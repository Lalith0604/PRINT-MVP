"""
ai_print.py — AI-powered print assistant using Groq.

Just describe what you want to print in plain English.
The AI understands your intent and runs the right print command.

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
SETUP
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
  1. Get a free Groq API key at: https://console.groq.com
  2. Set your API key (one-time):
       Windows CMD:   set GROQ_API_KEY=your_key_here
       PowerShell:    $env:GROQ_API_KEY="your_key_here"
     Or paste it when prompted on first run.

  3. Install dependency:
       pip install groq

in cmd
       # Install Groq
pip install groq

# Get a free API key at https://console.groq.com
# Then set it (CMD):
set GROQ_API_KEY=your_key_here

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
USAGE
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
  python ai_print.py

  Then just type naturally, for example:
    > Print C:/reports/sales.pdf in black and white, 3 copies
    > Convert my file C:/docs/report.docx to PDF and save to C:/Downloads
    > Print C:/pics/photo.png in color on my HP printer
    > Print C:/slides/deck.pptx 2 copies black and white to HP LaserJet

  Type 'help' to see example prompts.
  Type 'quit' or 'exit' to stop.
"""

import os
import sys
import json
import subprocess

# ── Groq client setup ─────────────────────────────────────────────────────────

def get_groq_client():
    try:
        from groq import Groq
    except ImportError:
        print("Groq not installed. Run: pip install groq")
        sys.exit(1)

    api_key = os.environ.get("GROQ_API_KEY", "").strip()
    if not api_key:
        print("\n  Groq API key not found in environment.")
        print("  Get a free key at: https://console.groq.com")
        api_key = input("  Paste your GROQ_API_KEY here: ").strip()
        if not api_key:
            print("No API key provided. Exiting.")
            sys.exit(1)
        # Save for this session
        os.environ["GROQ_API_KEY"] = api_key

    return Groq(api_key=api_key)


# ── System prompt for Groq ────────────────────────────────────────────────────

SYSTEM_PROMPT = """
You are a smart print assistant. The user will describe what they want to do
with a file (print it, convert it, save a copy, etc.) in plain English.

Your job is to extract the intent and return ONLY a valid JSON object — no
explanation, no markdown, no code fences — just raw JSON.

JSON schema:
{
  "command":      string,   // one of: print | toxpdf | pptpdf | wordpdf | imgpdf | download | xlprint | pptprint | wordprint | imgprint
  "file":         string,   // the file path the user mentioned (required)
  "printer":      string | null,   // printer name if mentioned, else null
  "color":        "bw" | "color",  // default "bw" unless user says color/colour
  "copies":       integer,          // default 1
  "out":          string | null,   // output folder if mentioned, else null
  "clarify":      string | null    // if you cannot determine the command or file, set this to a short question to ask the user; leave null if you have enough info
}

Command selection rules:
- User wants to PRINT a PDF → "print"
- User wants to PRINT an Excel/xlsx/xls → "xlprint"
- User wants to PRINT a PowerPoint/pptx/ppt → "pptprint"
- User wants to PRINT a Word/docx/doc → "wordprint"
- User wants to PRINT an image (jpg/png/bmp/tiff/gif/webp) → "imgprint"
- User wants to CONVERT Excel to PDF only (no print) → "toxpdf"
- User wants to CONVERT PowerPoint to PDF only → "pptpdf"
- User wants to CONVERT Word to PDF only → "wordpdf"
- User wants to CONVERT image to PDF only → "imgpdf"
- User wants to DOWNLOAD/SAVE/COPY a PDF → "download"

Color rules:
- "black and white", "bw", "grayscale", "greyscale", "no color" → "bw"
- "color", "colour", "in color", "coloured" → "color"
- If not mentioned → default "bw"

Copies rules:
- Extract any number mentioned: "3 copies", "print 5", "×2" → that number
- If not mentioned → 1

Always return valid JSON. Never return anything except the JSON object.
"""


# ── Call Groq to parse the user intent ────────────────────────────────────────

def parse_intent(client, user_input: str) -> dict:
    response = client.chat.completions.create(
        model="llama-3.3-70b-versatile",  # fast + smart Groq model
        messages=[
            {"role": "system", "content": SYSTEM_PROMPT},
            {"role": "user",   "content": user_input},
        ],
        temperature=0,
        max_tokens=300,
    )
    raw = response.choices[0].message.content.strip()

    # Strip any accidental markdown fences
    if raw.startswith("```"):
        raw = raw.split("```")[1]
        if raw.startswith("json"):
            raw = raw[4:]
    raw = raw.strip()

    try:
        return json.loads(raw)
    except json.JSONDecodeError:
        return {"clarify": f"I couldn't parse your request. Could you rephrase it?\n(Raw AI output: {raw})"}


# ── Build and run the print_pdf.py command ────────────────────────────────────

def run_command(intent: dict) -> None:
    """Translate the parsed JSON intent into a print_pdf.py CLI call."""

    # Path to print_pdf.py — assumes same folder as this script
    script_dir = os.path.dirname(os.path.abspath(__file__))
    print_script = os.path.join(script_dir, "print_pdf.py")

    if not os.path.isfile(print_script):
        print(f"\n  ✗ Could not find print_pdf.py at: {print_script}")
        print("    Make sure ai_print.py and print_pdf.py are in the same folder.")
        return

    command = intent.get("command")
    file    = intent.get("file")
    printer = intent.get("printer")
    color   = intent.get("color", "bw")
    copies  = int(intent.get("copies", 1))
    out     = intent.get("out")

    if not command:
        print("  ✗ AI could not determine the command. Please try again.")
        return

    if not file:
        print("  ✗ AI could not find a file path in your message. Please include the file path.")
        return

    # Build the command list
    cmd = [sys.executable, print_script, command, file]

    # Append options depending on command type
    print_commands = {"print", "xlprint", "pptprint", "wordprint", "imgprint"}
    convert_commands = {"toxpdf", "pptpdf", "wordpdf", "imgpdf"}

    if command in print_commands:
        if printer:
            cmd += ["--printer", printer]
        cmd += ["--color", color]
        cmd += ["--copies", str(copies)]
        if out:
            cmd += ["--out", out]

    elif command in convert_commands:
        if out:
            cmd += ["--out", out]

    elif command == "download":
        if out:
            cmd += ["--out", out]
        else:
            # download requires --out, ask for it
            out = input("  Where do you want to save it? Enter folder path: ").strip()
            if out:
                cmd += ["--out", out]
            else:
                print("  ✗ Output folder is required for download. Cancelled.")
                return

    # Show the command being run
    display_cmd = " ".join(f'"{c}"' if " " in c else c for c in cmd)
    print(f"\n  Running: {display_cmd}\n")

    # Execute
    result = subprocess.run(cmd)
    if result.returncode != 0:
        print("\n  ✗ Command failed. See error above.")


# ── Pretty summary of what AI understood ─────────────────────────────────────

def summarise_intent(intent: dict) -> str:
    cmd     = intent.get("command", "?")
    file    = intent.get("file", "?")
    printer = intent.get("printer") or "system default"
    color   = intent.get("color", "bw")
    copies  = intent.get("copies", 1)
    out     = intent.get("out") or "same folder as file"

    cmd_labels = {
        "print":      "Print PDF",
        "xlprint":    "Convert Excel → PDF → Print",
        "pptprint":   "Convert PowerPoint → PDF → Print",
        "wordprint":  "Convert Word → PDF → Print",
        "imgprint":   "Convert Image → PDF → Print",
        "toxpdf":     "Convert Excel → PDF",
        "pptpdf":     "Convert PowerPoint → PDF",
        "wordpdf":    "Convert Word → PDF",
        "imgpdf":     "Convert Image → PDF",
        "download":   "Download / Save PDF copy",
    }

    lines = [
        f"  Action   : {cmd_labels.get(cmd, cmd)}",
        f"  File     : {file}",
    ]
    if cmd in {"print", "xlprint", "pptprint", "wordprint", "imgprint"}:
        lines += [
            f"  Printer  : {printer}",
            f"  Color    : {'Full Color' if color == 'color' else 'Black & White'}",
            f"  Copies   : {copies}",
        ]
    if cmd not in {"print"} and cmd != "download" or out != "same folder as file":
        lines.append(f"  Save to  : {out}")

    return "\n".join(lines)


# ── Help text ─────────────────────────────────────────────────────────────────

HELP_TEXT = """
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
 AI PRINT ASSISTANT — Example prompts
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

  Print a PDF:
    Print C:/reports/sales.pdf in black and white
    Print C:/reports/sales.pdf, 3 copies, color, HP LaserJet

  Print Excel:
    Print C:/data/budget.xlsx 2 copies bw
    Convert and print C:/data/budget.xlsx in color

  Print PowerPoint:
    Print C:/slides/deck.pptx black and white 5 copies

  Print Word:
    Print C:/docs/letter.docx on Canon printer in color

  Print Image:
    Print C:/pics/photo.jpg in color, 2 copies

  Convert to PDF only (no printing):
    Convert C:/data/report.xlsx to PDF and save to C:/exports
    Convert C:/slides/presentation.pptx to PDF

  Save a PDF copy:
    Save a copy of C:/reports/final.pdf to C:/Users/lalit/Downloads

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
"""

# ── Main interactive loop ─────────────────────────────────────────────────────

def main():
    print("\n" + "━"*54)
    print("  🖨  AI Print Assistant  (powered by Groq)")
    print("━"*54)
    print("  Type your request in plain English.")
    print("  Type 'help' for examples, 'quit' to exit.")
    print("━"*54 + "\n")

    client = get_groq_client()

    while True:
        try:
            user_input = input("You: ").strip()
        except (KeyboardInterrupt, EOFError):
            print("\n  Goodbye!")
            break

        if not user_input:
            continue

        if user_input.lower() in ("quit", "exit", "q"):
            print("  Goodbye!")
            break

        if user_input.lower() in ("help", "h", "?"):
            print(HELP_TEXT)
            continue

        print("\n  Thinking…")
        intent = parse_intent(client, user_input)

        # If AI needs clarification
        if intent.get("clarify"):
            print(f"\n  AI: {intent['clarify']}\n")
            continue

        # Show summary and confirm
        print("\n  Here's what I understood:\n")
        print(summarise_intent(intent))
        print()

        confirm = input("  Proceed? [Y/n]: ").strip().lower()
        if confirm in ("", "y", "yes"):
            run_command(intent)
        else:
            print("  Cancelled.\n")

        print()


if __name__ == "__main__":
    main()