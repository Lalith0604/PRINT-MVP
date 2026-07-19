"""
whatsapp_server.py -- WhatsApp Print Assistant via Twilio + Flask

How it works:
  WhatsApp message -> Twilio -> this Flask server -> Groq AI -> print_pdf.py -> printer
  Then replies back on WhatsApp with the result.

------------------------------------------------------------
ONE-TIME SETUP (do this once)
------------------------------------------------------------

STEP 1 -- Install dependencies:
    pip install twilio flask groq

STEP 2 -- Get a free Twilio account:
    https://www.twilio.com/try-twilio

STEP 3 -- Set up Twilio WhatsApp Sandbox:
    -> Twilio Console -> Messaging -> Try it out -> Send a WhatsApp message
    -> Follow the instructions to join the sandbox from your WhatsApp

STEP 4 -- Get your Twilio credentials:
    -> Twilio Console -> Account Info -> copy Account SID and Auth Token

STEP 5 -- Get a free Groq API key:
    https://console.groq.com

STEP 6 -- Set environment variables (run these in CMD before starting server):
    set TWILIO_ACCOUNT_SID=ACxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
    set TWILIO_AUTH_TOKEN=your_auth_token_here
    set TWILIO_WHATSAPP_FROM=whatsapp:+14155238886
    set GROQ_API_KEY=your_groq_key_here

STEP 7 -- Install and start ngrok to expose your local server:
    Download: https://ngrok.com/download
    Run:      ngrok http 5000
    Copy the https URL shown (e.g. https://abc123.ngrok.io)

STEP 8 -- Set the Twilio Webhook URL:
    -> Twilio Console -> Messaging -> Settings -> WhatsApp Sandbox Settings
    -> "When a message comes in" -> paste your ngrok URL + /webhook
    -> e.g. https://abc123.ngrok.io/webhook
    -> Save

STEP 9 -- Start this server:
    python whatsapp_server.py

------------------------------------------------------------
USAGE -- Two ways to use:
------------------------------------------------------------
  Option 1 - Send a file directly on WhatsApp:
    -> Attach any PDF/Excel/PPT/Word/Image file and send it
    -> Bot replies: "Got your file! How to print?"
    -> You reply: "bw 2 copies" or "color on HP LaserJet"
    -> Bot prints it!

  Option 2 - Type a command:
    Print C:\\docs\\report.pdf in black and white 2 copies
    Print C:\\photo.jpg in color on HP LaserJet
    List my printers
    help
"""

import os
import sys
import json
import subprocess
import threading

from flask import Flask, request
from twilio.twiml.messaging_response import MessagingResponse
from twilio.rest import Client as TwilioClient
import urllib.request
import mimetypes

app = Flask(__name__)

# -- Config from environment variables ----------------------------------------

TWILIO_ACCOUNT_SID   = os.environ.get("TWILIO_ACCOUNT_SID", "")
TWILIO_AUTH_TOKEN    = os.environ.get("TWILIO_AUTH_TOKEN", "")
TWILIO_WHATSAPP_FROM = os.environ.get("TWILIO_WHATSAPP_FROM", "whatsapp:+14155238886")
GROQ_API_KEY         = os.environ.get("GROQ_API_KEY", "")

# Path to print_pdf.py -- must be in the same folder
# Folder where files sent via WhatsApp are saved temporarily
DOWNLOAD_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "whatsapp_received")
os.makedirs(DOWNLOAD_DIR, exist_ok=True)

# In-memory session store: tracks pending file per sender
# Format: { "whatsapp:+91xxxxxxx": { "file": "path", "filename": "name.pdf" } }
pending_sessions = {}
# file_buffer: collects files sent within GROUP_WINDOW seconds into one batch
# structure: { sender: {"files": [...], "last_time": float, "timer": Timer} }
file_buffer    = {}
GROUP_WINDOW   = 5  # seconds to wait for more files before asking how to print
SCRIPT_DIR   = os.path.dirname(os.path.abspath(__file__))
PRINT_SCRIPT = os.path.join(SCRIPT_DIR, "print_pdf.py")

# -- Groq system prompt (same logic as ai_print.py) ---------------------------

SYSTEM_PROMPT = """
You are a smart print assistant. The user will describe what they want to do
with one or more files in plain English via WhatsApp.

Return ONLY a valid JSON object -- no explanation, no markdown, no code fences.

JSON schema:
{
  "jobs": [
    {
      "command":  string,        // printers | print | toxpdf | pptpdf | wordpdf | imgpdf | download | xlprint | pptprint | wordprint | imgprint
      "file":     string | null, // file path, null only for printers command
      "printer":  string | null, // printer name if mentioned, else null
      "color":    "bw" | "color",// default "bw"
      "copies":   integer,       // default 1
      "out":      string | null  // output folder if mentioned, else null
    }
  ],
  "clarify": string | null       // short question if request is unclear, else null
}

Command rules (by file extension):
- LIST/SHOW printers -> "printers" (no file)
- .pdf  -> "print"
- .xlsx/.xls/.xlsm -> "xlprint"
- .pptx/.ppt/.pptm/.odp -> "pptprint"
- .docx/.doc/.odt/.rtf -> "wordprint"
- .jpg/.jpeg/.png/.bmp/.tiff/.gif/.webp -> "imgprint"
- CONVERT only (no print): toxpdf / pptpdf / wordpdf / imgpdf
- DOWNLOAD/SAVE PDF -> "download"

Color: "bw"/"black and white"/"grayscale" -> "bw" | "color"/"colour" -> "color" | default "bw"
Copies: extract number if mentioned, else 1
Multi-file: create one job per file.

Return only JSON. Nothing else.
"""

# -- Help message --------------------------------------------------------------

HELP_MSG = """[PRINTERS] *AI Print Assistant*

*Two ways to print:*

1️⃣ *Send a file directly* (easiest!)
   Just attach and send any file — I will ask how you want it printed.
   Supported: PDF, Excel, PowerPoint, Word, Images (JPG/PNG/BMP/TIFF/WEBP)

2️⃣ *Type a command*
   Mention the file path and preferences, e.g.:
   • Print C:\\docs\\report.pdf bw 2 copies
   • Print C:\\photo.jpg in color on HP LaserJet
   • List my printers

*After sending a file, reply with your print preferences:*
   • bw 2 copies
   • color on HP LaserJet
   • 3 copies black and white
   • color (uses default printer)

*Options:*
   • Color: "in color" or "black and white" (default: bw)
   • Copies: "3 copies" or "print 5"
   • Printer: "on HP LaserJet" or "Canon printer"

Reply *cancel* to discard a pending file.
"""

# -- Groq AI parser ------------------------------------------------------------

def download_whatsapp_file(media_url: str, filename: str) -> str:
    """
    Download a file sent via WhatsApp (hosted on Twilio) to DOWNLOAD_DIR.
    Twilio requires Basic Auth using Account SID + Auth Token.
    Returns the local saved file path.
    """
    local_path = os.path.join(DOWNLOAD_DIR, filename)

    # Build authenticated request (Twilio media requires credentials)
    password_mgr = urllib.request.HTTPPasswordMgrWithDefaultRealm()
    password_mgr.add_password(None, media_url, TWILIO_ACCOUNT_SID, TWILIO_AUTH_TOKEN)
    auth_handler = urllib.request.HTTPBasicAuthHandler(password_mgr)
    opener = urllib.request.build_opener(auth_handler)

    with opener.open(media_url) as response:
        with open(local_path, "wb") as f:
            f.write(response.read())

    return local_path


def guess_extension(content_type: str, media_url: str) -> str:
    """Guess file extension from MIME type or URL."""
    ext = mimetypes.guess_extension(content_type.split(";")[0].strip()) or ""
    # Fix common wrong guesses by Python's mimetypes module
    fix = {".jpe": ".jpg", ".jpeg": ".jpg", ".tiff": ".tiff", ".htm": ".html"}
    ext = fix.get(ext, ext)
    if not ext:
        # fallback: try to get from URL
        url_path = media_url.split("?")[0]
        ext = os.path.splitext(url_path)[-1] or ".bin"
    return ext


def get_print_command(file_path: str) -> str:
    """Return the right print_pdf.py command based on file extension."""
    ext = os.path.splitext(file_path)[1].lower()
    mapping = {
        ".pdf":  "print",
        ".xlsx": "xlprint", ".xls": "xlprint", ".xlsm": "xlprint",
        ".pptx": "pptprint", ".ppt": "pptprint", ".pptm": "pptprint", ".odp": "pptprint",
        ".docx": "wordprint", ".doc": "wordprint", ".odt": "wordprint", ".rtf": "wordprint",
        ".jpg":  "imgprint", ".jpeg": "imgprint", ".png": "imgprint",
        ".bmp":  "imgprint", ".tiff": "imgprint", ".tif": "imgprint",
        ".gif":  "imgprint", ".webp": "imgprint",
    }
    return mapping.get(ext, "")


def file_type_label(file_path: str) -> str:
    """Human-friendly file type label."""
    ext = os.path.splitext(file_path)[1].lower()
    labels = {
        ".pdf": "PDF", ".xlsx": "Excel", ".xls": "Excel", ".xlsm": "Excel",
        ".pptx": "PowerPoint", ".ppt": "PowerPoint", ".pptm": "PowerPoint",
        ".docx": "Word", ".doc": "Word", ".odt": "Word", ".rtf": "Word",
        ".jpg": "Image", ".jpeg": "Image", ".png": "Image", ".bmp": "Image",
        ".tiff": "Image", ".tif": "Image", ".gif": "Image", ".webp": "Image",
    }
    return labels.get(ext, "File")


def get_groq_client():

    try:
        from groq import Groq
    except ImportError:
        return None
    if not GROQ_API_KEY:
        return None
    return Groq(api_key=GROQ_API_KEY)


def parse_intent(user_text: str) -> dict:
    """Send user message to Groq and return parsed JSON intent."""
    client = get_groq_client()
    if not client:
        return {"clarify": "Groq API key not configured. Please set GROQ_API_KEY."}

    response = client.chat.completions.create(
        model="llama-3.3-70b-versatile",
        messages=[
            {"role": "system", "content": SYSTEM_PROMPT},
            {"role": "user",   "content": user_text},
        ],
        temperature=0,
        max_tokens=800,
    )

    raw = response.choices[0].message.content.strip()
    if raw.startswith("```"):
        raw = raw.split("```")[1]
        if raw.startswith("json"):
            raw = raw[4:]
    raw = raw.strip()

    try:
        parsed = json.loads(raw)
        if "jobs" not in parsed and "command" in parsed:
            parsed = {"jobs": [parsed], "clarify": parsed.get("clarify")}
        return parsed
    except json.JSONDecodeError:
        return {"clarify": "Sorry, I couldn't understand that. Could you rephrase?"}


# -- Job runner ----------------------------------------------------------------

def run_job(job: dict) -> tuple[bool, str]:
    """
    Execute one print/convert job.
    Returns (success: bool, message: str) to send back via WhatsApp.
    """
    command = job.get("command")
    file    = job.get("file")
    printer = job.get("printer")
    color   = job.get("color", "bw")
    copies  = int(job.get("copies", 1))
    out     = job.get("out")

    if not command:
        return False, "[FAIL] Could not determine command."

    # printers -- list available printers
    if command == "printers":
        result = subprocess.run(
            [sys.executable, PRINT_SCRIPT, "printers"],
            capture_output=True, text=True, encoding="utf-8",
            env={**os.environ, "PYTHONIOENCODING": "utf-8"}
        )
        if result.returncode == 0:
            output = result.stdout.strip() or "No printers found."
            return True, f"[PRINTERS] *Available Printers:*\n{output}"
        return False, f"[FAIL] Could not list printers:\n{result.stderr.strip()}"

    if not file:
        return False, "[FAIL] No file path found in your message."

    if not os.path.isfile(file):
        return False, f"[FAIL] File not found on this computer:\n`{file}`"

    # Build CLI command
    cmd = [sys.executable, PRINT_SCRIPT, command, file]

    print_commands   = {"print", "xlprint", "pptprint", "wordprint", "imgprint"}
    convert_commands = {"toxpdf", "pptpdf", "wordpdf", "imgpdf"}

    if command in print_commands:
        if printer:
            cmd += ["--printer", printer]
        cmd += ["--color", color, "--copies", str(copies)]
        if out:
            cmd += ["--out", out]

    elif command in convert_commands:
        if out:
            cmd += ["--out", out]

    elif command == "download":
        if out:
            cmd += ["--out", out]
        else:
            return False, "[FAIL] Please mention where to save the file (e.g. 'save to C:\\Downloads')."

    result = subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8",
                            errors="replace",
                            env={**os.environ, "PYTHONIOENCODING": "utf-8"})


    fname   = os.path.basename(file)
    color_l = "Color" if color == "color" else "B&W"

    if result.returncode == 0:
        if command in print_commands:
            return True, (
                f"[OK] *Printed:* {fname}\n"
                f"   Printer : {printer or 'default'}\n"
                f"   Color   : {color_l}\n"
                f"   Copies  : {copies}"
            )
        else:
            return True, f"[OK] *Converted:* {fname} -> PDF"
    else:
        err = result.stderr.strip() or result.stdout.strip()
        return False, f"[FAIL] Failed for `{fname}`:\n{err[:300]}"


def build_job_from_pending(session: dict, instructions: str) -> dict:
    """
    Use Groq to parse the user's intent (print vs convert) and settings
    (color, copies, printer) from their reply text.
    """
    file_path = session["file"]
    ext       = os.path.splitext(file_path)[1].lower()

    # Map extension to both print and convert commands
    EXT_MAP = {
        ".pdf":  {"print": "print",     "convert": None},
        ".xlsx": {"print": "xlprint",   "convert": "toxpdf"},
        ".xls":  {"print": "xlprint",   "convert": "toxpdf"},
        ".xlsm": {"print": "xlprint",   "convert": "toxpdf"},
        ".pptx": {"print": "pptprint",  "convert": "pptpdf"},
        ".ppt":  {"print": "pptprint",  "convert": "pptpdf"},
        ".pptm": {"print": "pptprint",  "convert": "pptpdf"},
        ".odp":  {"print": "pptprint",  "convert": "pptpdf"},
        ".docx": {"print": "wordprint", "convert": "wordpdf"},
        ".doc":  {"print": "wordprint", "convert": "wordpdf"},
        ".odt":  {"print": "wordprint", "convert": "wordpdf"},
        ".rtf":  {"print": "wordprint", "convert": "wordpdf"},
        ".jpg":  {"print": "imgprint",  "convert": "imgpdf"},
        ".jpeg": {"print": "imgprint",  "convert": "imgpdf"},
        ".png":  {"print": "imgprint",  "convert": "imgpdf"},
        ".bmp":  {"print": "imgprint",  "convert": "imgpdf"},
        ".tiff": {"print": "imgprint",  "convert": "imgpdf"},
        ".tif":  {"print": "imgprint",  "convert": "imgpdf"},
        ".gif":  {"print": "imgprint",  "convert": "imgpdf"},
        ".webp": {"print": "imgprint",  "convert": "imgpdf"},
    }

    cmds = EXT_MAP.get(ext)
    if not cmds:
        return {}

    # Ask Groq: action (print or convert) + settings
    settings_prompt = f"""
The user sent a file and is giving instructions for what to do with it.
File extension: {ext}

Return ONLY valid JSON, no explanation, no markdown:
{{
  "action":   "print" | "convert",
  "printer":  string | null,
  "color":    "bw" | "color",
  "copies":   integer,
  "save_to":  string | null
}}

Rules:
- action: "convert" ONLY if user says convert/save as pdf/make pdf — otherwise "print"
- color: MUST be "bw" if user says any of: bw, b&w, black and white, black & white, grayscale, greyscale, no color — MUST be "color" if user says color or colour — DEFAULT is "bw" if not mentioned
- copies: extract any integer mentioned (e.g. "2 copies", "3x", "print 5") — default 1
- printer: printer name only if explicitly named (e.g. "HP LaserJet") — else null
- save_to: folder path only if user says "save to" or "export to" — else null

IMPORTANT: "all in bw", "all in 2 copies bw", "all bw" — color must be "bw"
IMPORTANT: "2 copies" means copies=2, not color

User instruction: "{instructions}"
"""
    client = get_groq_client()
    settings = {"action": "print", "printer": None, "color": "bw", "copies": 1, "save_to": None}

    if client:
        try:
            response = client.chat.completions.create(
                model="llama-3.3-70b-versatile",
                messages=[{"role": "user", "content": settings_prompt}],
                temperature=0, max_tokens=150,
            )
            raw = response.choices[0].message.content.strip()
            if raw.startswith("```"):
                raw = raw.split("```")[1]
                if raw.startswith("json"):
                    raw = raw[4:]
            parsed = json.loads(raw.strip())
            settings.update(parsed)
        except Exception as e:
            print(f"[Groq] settings parse error: {e}")

    # Pick command based on action
    action  = settings.get("action", "print")
    command = cmds["convert"] if action == "convert" and cmds["convert"] else cmds["print"]

    return {
        "command": command,
        "file":    file_path,
        "printer": settings.get("printer"),
        "color":   settings.get("color", "bw"),
        "copies":  int(settings.get("copies", 1)),
        "out":     settings.get("save_to"),
    }


# -- Flask webhook -------------------------------------------------------------

def send_whatsapp_reply(to: str, body: str):
    """Send a WhatsApp message via Twilio REST API (used for async replies)."""
    try:
        client = TwilioClient(TWILIO_ACCOUNT_SID, TWILIO_AUTH_TOKEN)
        client.messages.create(
            from_=TWILIO_WHATSAPP_FROM,
            to=to,
            body=body
        )
        print(f"[SERVER] Reply sent to {to}: {repr(body[:80])}")
    except Exception as e:
        print(f"[SERVER] Failed to send reply: {e}")


@app.route("/webhook", methods=["POST"])
def webhook():
    """
    Receive incoming WhatsApp message from Twilio and process it.

    Two flows:
      A) User sends a FILE  -> download it, save it, ask how to print.
      B) User sends TEXT    -> if pending file exists use it + instructions,
                              else treat as a normal text command.
    """
    incoming_msg = request.form.get("Body", "").strip()
    sender       = request.form.get("From", "")
    num_media    = int(request.form.get("NumMedia", 0))

    print(f"\n[WhatsApp] From: {sender}")
    print(f"[WhatsApp] Message: {incoming_msg!r}  |  NumMedia: {num_media}")

    resp = MessagingResponse()
    msg  = resp.message()

    # ── Help ──────────────────────────────────────────────────────────────────
    if incoming_msg.lower() in ("help", "hi", "hello", "hey", "?") and num_media == 0:
        msg.body(HELP_MSG)
        return str(resp)

    # ── Cancel pending session ─────────────────────────────────────────────────
    if incoming_msg.lower() in ("cancel", "stop it", "never mind", "nevermind"):
        if sender in pending_sessions:
            fname = pending_sessions[sender]["filename"]
            del pending_sessions[sender]
            msg.body(f"Cancelled. {fname} will not be printed.")
        else:
            msg.body("Nothing to cancel.")
        return str(resp)

    # -- File received: buffer it, wait GROUP_WINDOW seconds for more files -----
    if num_media > 0:
        import time as _time

        FILE_EXTENSIONS = (
            ".pdf", ".png", ".jpg", ".jpeg", ".bmp", ".tiff", ".tif", ".gif", ".webp",
            ".xlsx", ".xls", ".xlsm", ".pptx", ".ppt", ".pptm", ".odp",
            ".docx", ".doc", ".odt", ".rtf"
        )

        def is_filename_not_instruction(text):
            t = text.strip().lower()
            if any(t.endswith(e) for e in FILE_EXTENSIONS):
                return True
            if "." in t and " " not in t:
                return True
            return False

        # Download this single file
        num       = sender.replace("whatsapp:+", "").replace(":", "")
        timestamp = int(_time.time() * 1000)
        media_url  = request.form.get("MediaUrl0", "")
        media_type = request.form.get("MediaContentType0", "application/octet-stream")
        ext        = guess_extension(media_type, media_url)
        fname      = f"wa_{num}_{timestamp}{ext}"

        print(f"[WhatsApp] File received: {fname}  type={media_type}")

        try:
            local_path = download_whatsapp_file(media_url, fname)
        except Exception as e:
            msg.body(f"Sorry, could not download your file.\nError: {e}")
            return str(resp)

        command = get_print_command(local_path)
        if not command:
            os.remove(local_path)
            msg.body(
                f"Sorry, I don\'t support that file type ({ext}).\n"
                "Supported: PDF, Excel, PowerPoint, Word, Images"
            )
            return str(resp)

        file_entry = {
            "file":     local_path,
            "filename": fname,
            "ftype":    file_type_label(local_path),
        }

        # Caption = real print instruction? Print immediately.
        caption_is_instruction = (
            incoming_msg
            and len(incoming_msg.strip()) > 2
            and not is_filename_not_instruction(incoming_msg)
        )

        if caption_is_instruction:
            print(f"[WhatsApp] Caption detected: {incoming_msg!r}")
            def run_caption(entry=file_entry, instr=incoming_msg, sndr=sender):
                job = build_job_from_pending(entry, instr)
                if job:
                    success, reply = run_job(job)
                    send_whatsapp_reply(sndr, reply)
                else:
                    send_whatsapp_reply(sndr, "Could not process " + entry["filename"] + ".")
            threading.Thread(target=run_caption, daemon=True).start()
            msg.body("Got your *" + file_entry["ftype"] + "* file! Processing now...")
            return str(resp)

        # No caption — add to grouping buffer, ask after GROUP_WINDOW seconds
        def flush_buffer(sndr=sender):
            buf = file_buffer.pop(sndr, None)
            if not buf:
                return
            files = buf["files"]
            if len(files) == 1:
                receipt = f"Got your *" + files[0]["ftype"] + "* file! [OK]"
            else:
                lines_r = [f"Got *{len(files)} files*! [OK]"]
                for i, d in enumerate(files, 1):
                    lines_r.append(f"  {i}. " + d["ftype"] + ": " + d["filename"])
                receipt = "\n".join(lines_r)
            pending_sessions[sndr] = {"files": files}
            send_whatsapp_reply(sndr,
                f"{receipt}\n\n"
                "How would you like to print? Reply with:\n"
                "  *bw* - black & white (default)\n"
                "  *color* - full color\n"
                "  *2 copies* - number of copies\n"
                "  *HP LaserJet* - printer name\n\n"
                "Combine: *bw 2 copies* or *color on HP LaserJet*\n"
                "Or reply *cancel* to discard."
            )

        if sender in file_buffer:
            file_buffer[sender]["timer"].cancel()
            file_buffer[sender]["files"].append(file_entry)
            print(f"[WhatsApp] Added to buffer: {len(file_buffer[sender]['files'])} files so far")
        else:
            file_buffer[sender] = {"files": [file_entry]}
            print(f"[WhatsApp] New buffer started — waiting {GROUP_WINDOW}s for more files")

        t = threading.Timer(GROUP_WINDOW, flush_buffer, args=[sender])
        file_buffer[sender]["timer"] = t
        t.start()

        msg.body("")
        return str(resp)

    # ── FLOW B: User sent text ─────────────────────────────────────────────────

    # Check if there are pending files waiting for print instructions
    if sender in pending_sessions:
        session      = pending_sessions.pop(sender)
        instructions = incoming_msg if len(incoming_msg) > 3 else "bw 1 copy"

        # Support both new format (files list) and old format (single file)
        files = session.get("files", [])
        if not files and "file" in session:
            files = [{"file": session["file"], "filename": session.get("filename", "file"),
                      "ftype": file_type_label(session["file"])}]

        if not files:
            msg.body("No files found. Please resend your file.")
            return str(resp)

        total = len(files)

        def run_all_pending(files=files, instructions=instructions, sender=sender, total=total):
            results = []
            for i, d in enumerate(files, 1):
                job = build_job_from_pending(d, instructions)
                if job:
                    success, reply = run_job(job)
                    label = f"[{i}/{total}] " if total > 1 else ""
                    results.append(label + reply)
                else:
                    results.append(f"[{i}/{total}] Could not process " + d.get("filename", "file"))
            send_whatsapp_reply(sender, "\n\n".join(results))

        threading.Thread(target=run_all_pending, daemon=True).start()
        count_msg = f"all {total} files" if total > 1 else "your file"
        msg.body(f"Got it! Printing {count_msg} now... I will send you a confirmation shortly.")
        return str(resp)

    # No pending file — normal text command flow
    if not incoming_msg:
        msg.body("Please send a file or type a command. Type *help* for examples.")
        return str(resp)

    # Parse with Groq AI
    intent = parse_intent(incoming_msg)

    if intent.get("clarify"):
        msg.body(f"(?) {intent['clarify']}")
        return str(resp)

    jobs  = intent.get("jobs", [])
    total = len(jobs)

    if total == 0:
        msg.body("I couldn\'t understand that. Type *help* to see examples.")
        return str(resp)

    if total == 1:
        success, reply = run_job(jobs[0])
        msg.body(reply)
        return str(resp)

    # Multiple jobs
    results   = []
    successes = 0
    for i, job in enumerate(jobs, start=1):
        fname   = os.path.basename(job.get("file") or "") or f"Job {i}"
        success, reply = run_job(job)
        if success:
            successes += 1
            results.append(f"[OK] [{i}/{total}] {fname}")
        else:
            results.append(f"[X] [{i}/{total}] {fname}: {reply.replace('[X] ', '')}")

    summary = "\n".join(results)
    final   = f"Print Summary ({successes}/{total} succeeded):\n\n{summary}"
    msg.body(final)
    return str(resp)


@app.route("/", methods=["GET"])
def index():
    """Health check endpoint."""
    return "[OK] WhatsApp Print Server is running!", 200


# -- Startup checks ------------------------------------------------------------

def check_config():
    ok = True
    print("\n" + "-"*54)
    print("  WhatsApp Print Server -- startup check")
    print("-"*54)

    if not GROQ_API_KEY:
        print("  [X] GROQ_API_KEY not set")
        print("    -> set GROQ_API_KEY=your_key")
        ok = False
    else:
        print("  [OK] GROQ_API_KEY found")

    if not TWILIO_ACCOUNT_SID:
        print("  [X] TWILIO_ACCOUNT_SID not set")
        ok = False
    else:
        print("  [OK] TWILIO_ACCOUNT_SID found")

    if not TWILIO_AUTH_TOKEN:
        print("  [X] TWILIO_AUTH_TOKEN not set")
        ok = False
    else:
        print("  [OK] TWILIO_AUTH_TOKEN found")

    if not os.path.isfile(PRINT_SCRIPT):
        print(f"  [X] print_pdf.py not found at: {PRINT_SCRIPT}")
        ok = False
    else:
        print(f"  [OK] print_pdf.py found")

    print("-"*54)
    if ok:
        print("  [OK] All good! Server starting on http://localhost:5000")
        print("  [OK] Webhook URL to paste in Twilio:")
        print("      https://<your-ngrok-url>/webhook")
    else:
        print("  [!] Fix the issues above before using.")
    print("-"*54 + "\n")

    return ok


# -- Main ----------------------------------------------------------------------

if __name__ == "__main__":
    check_config()
    # debug=False in production; use debug=True only for development
    app.run(host="0.0.0.0", port=5000, debug=False)