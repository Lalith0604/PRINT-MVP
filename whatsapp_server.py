"""
whatsapp_server.py — WhatsApp Print Assistant via Twilio + Flask

How it works:
  WhatsApp message → Twilio → this Flask server → Groq AI → print_pdf.py → printer
  Then replies back on WhatsApp with the result.

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
ONE-TIME SETUP (do this once)
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

STEP 1 — Install dependencies:
    pip install twilio flask groq

STEP 2 — Get a free Twilio account:
    https://www.twilio.com/try-twilio

STEP 3 — Set up Twilio WhatsApp Sandbox:
    → Twilio Console → Messaging → Try it out → Send a WhatsApp message
    → Follow the instructions to join the sandbox from your WhatsApp

STEP 4 — Get your Twilio credentials:
    → Twilio Console → Account Info → copy Account SID and Auth Token

STEP 5 — Get a free Groq API key:
    https://console.groq.com

STEP 6 — Set environment variables (run these in CMD before starting server):
    set TWILIO_ACCOUNT_SID=ACxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
    set TWILIO_AUTH_TOKEN=your_auth_token_here
    set TWILIO_WHATSAPP_FROM=whatsapp:+14155238886
    set GROQ_API_KEY=your_groq_key_here

STEP 7 — Install and start ngrok to expose your local server:
    Download: https://ngrok.com/download
    Run:      ngrok http 5000
    Copy the https URL shown (e.g. https://abc123.ngrok.io)

STEP 8 — Set the Twilio Webhook URL:
    → Twilio Console → Messaging → Settings → WhatsApp Sandbox Settings
    → "When a message comes in" → paste your ngrok URL + /webhook
    → e.g. https://abc123.ngrok.io/webhook
    → Save

STEP 9 — Start this server:
    python whatsapp_server.py

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
USAGE — Send these WhatsApp messages to your Twilio number:
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    Print C:\\docs\report.pdf in black and white 2 copies
    Print C:\\pics\\photo.jpg in color on HP LaserJet
    Print C:\report.pdf and C:\\slides.pptx bw 3 copies
    Convert C:\\data\\budget.xlsx to PDF save to C:\\Downloads
    List my printers
    help
"""

import os
import sys
import json
import subprocess

from flask import Flask, request
from twilio.twiml.messaging_response import MessagingResponse
from twilio.rest import Client as TwilioClient

app = Flask(__name__)

# ── Config from environment variables ────────────────────────────────────────

TWILIO_ACCOUNT_SID   = os.environ.get("TWILIO_ACCOUNT_SID", "")
TWILIO_AUTH_TOKEN    = os.environ.get("TWILIO_AUTH_TOKEN", "")
TWILIO_WHATSAPP_FROM = os.environ.get("TWILIO_WHATSAPP_FROM", "whatsapp:+14155238886")
GROQ_API_KEY         = os.environ.get("GROQ_API_KEY", "")

# Path to print_pdf.py — must be in the same folder
SCRIPT_DIR   = os.path.dirname(os.path.abspath(__file__))
PRINT_SCRIPT = os.path.join(SCRIPT_DIR, "print_pdf.py")

# ── Groq system prompt (same logic as ai_print.py) ───────────────────────────

SYSTEM_PROMPT = """
You are a smart print assistant. The user will describe what they want to do
with one or more files in plain English via WhatsApp.

Return ONLY a valid JSON object — no explanation, no markdown, no code fences.

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
- LIST/SHOW printers → "printers" (no file)
- .pdf  → "print"
- .xlsx/.xls/.xlsm → "xlprint"
- .pptx/.ppt/.pptm/.odp → "pptprint"
- .docx/.doc/.odt/.rtf → "wordprint"
- .jpg/.jpeg/.png/.bmp/.tiff/.gif/.webp → "imgprint"
- CONVERT only (no print): toxpdf / pptpdf / wordpdf / imgpdf
- DOWNLOAD/SAVE PDF → "download"

Color: "bw"/"black and white"/"grayscale" → "bw" | "color"/"colour" → "color" | default "bw"
Copies: extract number if mentioned, else 1
Multi-file: create one job per file.

Return only JSON. Nothing else.
"""

# ── Help message ──────────────────────────────────────────────────────────────

HELP_MSG = """🖨 *AI Print Assistant*

Send me a message describing what to print!

*Examples:*
• Print C:\\docs\\report.pdf bw 2 copies
• Print C:\\photo.jpg in color on HP LaserJet
• Print C:\\report.pdf and C:\\slides.pptx bw
• Convert C:\\budget.xlsx to PDF save to C:\\Downloads
• List my printers

*Supported files:*
PDF, Excel, PowerPoint, Word, Images (JPG/PNG/BMP/TIFF/WEBP)

*Options you can mention:*
• Color: "in color" or "black and white" (default: bw)
• Copies: "3 copies" or "print 5"
• Printer: "on HP LaserJet" or "Canon printer"
• Save to: "save to C:\\Downloads"
"""

# ── Groq AI parser ────────────────────────────────────────────────────────────

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


# ── Job runner ────────────────────────────────────────────────────────────────

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
        return False, "❌ Could not determine command."

    # printers — list available printers
    if command == "printers":
        result = subprocess.run(
            [sys.executable, PRINT_SCRIPT, "printers"],
            capture_output=True, text=True
        )
        if result.returncode == 0:
            output = result.stdout.strip() or "No printers found."
            return True, f"🖨 *Available Printers:*\n{output}"
        return False, f"❌ Could not list printers:\n{result.stderr.strip()}"

    if not file:
        return False, "❌ No file path found in your message."

    if not os.path.isfile(file):
        return False, f"❌ File not found on this computer:\n`{file}`"

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
            return False, "❌ Please mention where to save the file (e.g. 'save to C:\\Downloads')."

    result = subprocess.run(cmd, capture_output=True, text=True)

    fname   = os.path.basename(file)
    color_l = "Color" if color == "color" else "B&W"

    if result.returncode == 0:
        if command in print_commands:
            return True, (
                f"✅ *Printed:* {fname}\n"
                f"   Printer : {printer or 'default'}\n"
                f"   Color   : {color_l}\n"
                f"   Copies  : {copies}"
            )
        else:
            return True, f"✅ *Converted:* {fname} → PDF"
    else:
        err = result.stderr.strip() or result.stdout.strip()
        return False, f"❌ Failed for `{fname}`:\n{err[:300]}"


# ── Flask webhook ─────────────────────────────────────────────────────────────

@app.route("/webhook", methods=["POST"])
def webhook():
    """Receive incoming WhatsApp message from Twilio and process it."""
    incoming_msg = request.form.get("Body", "").strip()
    sender       = request.form.get("From", "")

    print(f"\n[WhatsApp] From: {sender}")
    print(f"[WhatsApp] Message: {incoming_msg}")

    resp = MessagingResponse()
    msg  = resp.message()

    # Handle help
    if incoming_msg.lower() in ("help", "hi", "hello", "hey", "?"):
        msg.body(HELP_MSG)
        return str(resp)

    # Parse with Groq AI
    intent = parse_intent(incoming_msg)

    # Groq needs clarification
    if intent.get("clarify"):
        msg.body(f"🤔 {intent['clarify']}")
        return str(resp)

    jobs  = intent.get("jobs", [])
    total = len(jobs)

    if total == 0:
        msg.body("❌ I couldn't understand your request. Type *help* to see examples.")
        return str(resp)

    # Single job
    if total == 1:
        success, reply = run_job(jobs[0])
        msg.body(reply)
        return str(resp)

    # Multiple jobs — run one by one, collect results
    results   = []
    successes = 0
    for i, job in enumerate(jobs, start=1):
        fname   = os.path.basename(job.get("file") or "") or f"Job {i}"
        success, reply = run_job(job)
        if success:
            successes += 1
            results.append(f"✅ [{i}/{total}] {fname}")
        else:
            results.append(f"❌ [{i}/{total}] {fname}: {reply.replace('❌ ', '')}")

    summary = "\n".join(results)
    final   = f"🖨 *Print Summary ({successes}/{total} succeeded):*\n\n{summary}"
    msg.body(final)
    return str(resp)


@app.route("/", methods=["GET"])
def index():
    """Health check endpoint."""
    return "✅ WhatsApp Print Server is running!", 200


# ── Startup checks ────────────────────────────────────────────────────────────

def check_config():
    ok = True
    print("\n" + "━"*54)
    print("  WhatsApp Print Server — startup check")
    print("━"*54)

    if not GROQ_API_KEY:
        print("  ✗ GROQ_API_KEY not set")
        print("    → set GROQ_API_KEY=your_key")
        ok = False
    else:
        print("  ✓ GROQ_API_KEY found")

    if not TWILIO_ACCOUNT_SID:
        print("  ✗ TWILIO_ACCOUNT_SID not set")
        ok = False
    else:
        print("  ✓ TWILIO_ACCOUNT_SID found")

    if not TWILIO_AUTH_TOKEN:
        print("  ✗ TWILIO_AUTH_TOKEN not set")
        ok = False
    else:
        print("  ✓ TWILIO_AUTH_TOKEN found")

    if not os.path.isfile(PRINT_SCRIPT):
        print(f"  ✗ print_pdf.py not found at: {PRINT_SCRIPT}")
        ok = False
    else:
        print(f"  ✓ print_pdf.py found")

    print("━"*54)
    if ok:
        print("  ✓ All good! Server starting on http://localhost:5000")
        print("  ✓ Webhook URL to paste in Twilio:")
        print("      https://<your-ngrok-url>/webhook")
    else:
        print("  ⚠ Fix the issues above before using.")
    print("━"*54 + "\n")

    return ok


# ── Main ──────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    check_config()
    # debug=False in production; use debug=True only for development
    app.run(host="0.0.0.0", port=5000, debug=False)