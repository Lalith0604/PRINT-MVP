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
# Get a free API key at https://console.groq.com
# Then set it (CMD):
set GROQ_API_KEY=your_key_here



You: Print C:\docs\report.pdf, C:\pics\photo.jpg and C:\slides\deck.pptx in black and white 2 copies

  Here's what I understood (3 files):

  ── Job 1 of 3 ──────────────────────
  Action   : Print PDF
  File     : C:\docs\report.pdf
  Printer  : system default
  Color    : Black & White
  Copies   : 2
  ── Job 2 of 3 ──────────────────────
  Action   : Convert Image → PDF → Print
  File     : C:\pics\photo.jpg
  ...
  ── Job 3 of 3 ──────────────────────
  Action   : Convert PowerPoint → PDF → Print
  File     : C:\slides\deck.pptx
  ...

  Proceed? [Y/n]: y

  [1/3] Running: ...   ✓
  [2/3] Running: ...   ✓
  [3/3] Running: ...   ✓

  ━━━ Done: 3/3 jobs completed successfully. ━━━
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
with one or more files in plain English. They may also ask for utility actions
like listing printers.

Your job is to extract ALL file intents and return ONLY a valid JSON object —
no explanation, no markdown, no code fences — just raw JSON.

JSON schema:
{
  "jobs": [
    {
      "command":  string,        // one of: printers | print | toxpdf | pptpdf | wordpdf | imgpdf | download | xlprint | pptprint | wordprint | imgprint
      "file":     string | null, // the file path for this job. null only for "printers"
      "printer":  string | null, // printer name if mentioned, else null
      "color":    "bw" | "color",// default "bw" unless user says color/colour
      "copies":   integer,       // default 1
      "out":      string | null  // output folder if mentioned, else null
    }
  ],
  "clarify": string | null       // if you truly cannot understand the request, ask a short question. null otherwise
}

MULTI-FILE RULES:
- If the user mentions multiple files, create one job object per file in the "jobs" array.
- Each file gets its own command based on its extension.
- Shared settings (color, copies, printer) apply to ALL jobs unless stated differently per file.
- Example: "print sales.pdf, photo.jpg and slides.pptx in bw 2 copies"
  → 3 jobs: print + imgprint + pptprint, all with color=bw, copies=2

Command selection rules (per file extension):
- User wants to LIST / SHOW / GET printers → "printers" (one job, no file)
- .pdf → "print"
- .xlsx / .xls / .xlsm → "xlprint"
- .pptx / .ppt / .pptm / .odp → "pptprint"
- .docx / .doc / .odt / .rtf → "wordprint"
- .jpg / .jpeg / .png / .bmp / .tiff / .gif / .webp → "imgprint"
- CONVERT only (no print): use toxpdf / pptpdf / wordpdf / imgpdf
- DOWNLOAD/SAVE/COPY a PDF → "download"

Color rules:
- "black and white", "bw", "grayscale", "greyscale", "no color" → "bw"
- "color", "colour", "in color", "coloured" → "color"
- If not mentioned → default "bw"

Copies rules:
- Extract any number mentioned: "3 copies", "print 5", "×2" → that number
- If not mentioned → 1

Always return valid JSON with a "jobs" array. Never return anything else.
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
        max_tokens=800,  # increased for multi-file responses
    )
    raw = response.choices[0].message.content.strip()

    # Strip any accidental markdown fences
    if raw.startswith("```"):
        raw = raw.split("```")[1]
        if raw.startswith("json"):
            raw = raw[4:]
    raw = raw.strip()

    try:
        parsed = json.loads(raw)
        # Normalise: if old single-job format returned, wrap it
        if "jobs" not in parsed and "command" in parsed:
            parsed = {"jobs": [parsed], "clarify": parsed.get("clarify")}
        return parsed
    except json.JSONDecodeError:
        return {"clarify": f"I couldn't parse your request. Could you rephrase it?\n(Raw AI output: {raw})"}


# ── Build and run the print_pdf.py command ────────────────────────────────────

def run_single_job(job: dict, print_script: str, job_num: int = None, total: int = None) -> bool:
    """
    Run one print/convert job. Returns True on success, False on failure.
    job_num / total are used for progress display when printing multiple files.
    """
    command = job.get("command")
    file    = job.get("file")
    printer = job.get("printer")
    color   = job.get("color", "bw")
    copies  = int(job.get("copies", 1))
    out     = job.get("out")

    prefix = ""
    if job_num is not None and total is not None and total > 1:
        prefix = f"[{job_num}/{total}] "

    if not command:
        print(f"  {prefix}✗ Could not determine command. Skipping.")
        return False

    # printers command needs no file
    if command == "printers":
        cmd = [sys.executable, print_script, "printers"]
        display_cmd = " ".join(cmd)
        print(f"\n  {prefix}Running: {display_cmd}\n")
        subprocess.run(cmd)
        return True

    if not file:
        print(f"  {prefix}✗ No file path found. Skipping.")
        return False

    # Build the CLI command
    cmd = [sys.executable, print_script, command, file]

    print_commands   = {"print", "xlprint", "pptprint", "wordprint", "imgprint"}
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
            out = input(f"  {prefix}Where to save? Enter folder path: ").strip()
            if out:
                cmd += ["--out", out]
            else:
                print(f"  {prefix}✗ Output folder required. Skipping.")
                return False

    display_cmd = " ".join(f'"{c}"' if " " in c else c for c in cmd)
    print(f"\n  {prefix}Running: {display_cmd}\n")

    result = subprocess.run(cmd)
    if result.returncode != 0:
        print(f"\n  {prefix}✗ Command failed. See error above.")
        return False
    return True


def run_command(intent: dict) -> None:
    """
    Run all jobs from the parsed intent, one by one.
    Shows progress when multiple files are involved.
    """
    script_dir   = os.path.dirname(os.path.abspath(__file__))
    print_script = os.path.join(script_dir, "print_pdf.py")

    if not os.path.isfile(print_script):
        print(f"\n  ✗ Could not find print_pdf.py at: {print_script}")
        print("    Make sure ai_print.py and print_pdf.py are in the same folder.")
        return

    jobs  = intent.get("jobs", [])
    total = len(jobs)

    if total == 0:
        print("  ✗ No jobs found. Please try again.")
        return

    success_count = 0
    for i, job in enumerate(jobs, start=1):
        ok = run_single_job(job, print_script, job_num=i, total=total)
        if ok:
            success_count += 1
        if i < total:
            print()  # blank line between jobs

    if total > 1:
        print(f"\n  ━━━ Done: {success_count}/{total} jobs completed successfully. ━━━")


# ── Pretty summary of what AI understood ─────────────────────────────────────

def summarise_intent(intent: dict) -> str:
    cmd_labels = {
        "printers":   "List available printers",
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
    print_commands = {"print", "xlprint", "pptprint", "wordprint", "imgprint"}

    jobs  = intent.get("jobs", [])
    total = len(jobs)
    lines = []

    for i, job in enumerate(jobs, start=1):
        cmd     = job.get("command", "?")
        file    = job.get("file", "?")
        printer = job.get("printer") or "system default"
        color   = job.get("color", "bw")
        copies  = job.get("copies", 1)
        out     = job.get("out") or "same folder as file"

        if total > 1:
            lines.append(f"  ── Job {i} of {total} ──────────────────────")

        if cmd == "printers":
            lines.append("  Action   : List all available printers")
            continue

        lines.append(f"  Action   : {cmd_labels.get(cmd, cmd)}")
        lines.append(f"  File     : {file}")
        if cmd in print_commands:
            lines.append(f"  Printer  : {printer}")
            lines.append(f"  Color    : {'Full Color' if color == 'color' else 'Black & White'}")
            lines.append(f"  Copies   : {copies}")
        if out != "same folder as file":
            lines.append(f"  Save to  : {out}")

    return "\n".join(lines)


# ── Help text ─────────────────────────────────────────────────────────────────

HELP_TEXT = """
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
 AI PRINT ASSISTANT — Example prompts
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

  List printers:
    List my printers
    Show available printers
    What printers do I have?

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

  Multiple files at once:
    Print C:/docs/report.pdf and C:/pics/photo.jpg in black and white
    Print C:/report.pdf, C:/slides.pptx and C:/photo.png 2 copies color
    Print C:/a.pdf C:/b.xlsx C:/c.pptx on HP LaserJet bw 3 copies

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