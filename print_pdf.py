"""
print_pdf.py — PDF, Excel, PowerPoint, Word & Image printing utility for Windows.

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
FEATURES
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
  1.  Print a PDF file to a printer
  2.  Convert Excel  (.xlsx / .xls / .xlsm)           → PDF
  3.  Convert PowerPoint (.pptx / .ppt / .pptm / .odp) → PDF
  4.  Convert Word   (.docx / .doc / .odt / .rtf)     → PDF
  5.  Convert Image  (JPG / PNG / BMP / TIFF / WEBP)  → PDF
  6.  Download / save a copy of any PDF to a folder
  7.  Excel     → PDF → Print  (one command)
  8.  PowerPoint → PDF → Print (one command)
  9.  Word      → PDF → Print  (one command)
  10. Image     → PDF → Print  (one command)

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
COLOR vs BLACK & WHITE  (--color flag)
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
  All print commands accept an optional --color flag:

    --color bw     → Black & white print (DEFAULT if flag is omitted).
                     The PDF is converted to grayscale by Ghostscript
                     before printing, so NO color ink is used regardless
                     of what the printer driver is set to.

    --color color  → Full color print. The PDF is sent as-is.

  If Ghostscript is not installed and --color bw is used, the script
  prints a warning and continues WITHOUT grayscale conversion.

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
NUMBER OF COPIES  (--copies / -n flag)
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
  All print commands accept an optional --copies (or -n) flag:

    --copies 1     → Print 1 copy (DEFAULT if flag is omitted).
    --copies 5     → Print 5 copies of the document.

  Implemented via SumatraPDF's -print-settings "<n>x" option.

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
REQUIREMENTS
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
  Software to install:
    • SumatraPDF  (for printing)
        https://www.sumatrapdfreader.org/download-free-pdf-viewer
        → Place SumatraPDF.exe next to this script OR install normally.

    • LibreOffice  (for Excel / PowerPoint / Word → PDF conversion)
        https://www.libreoffice.org/download/download/

    • Ghostscript  (for black & white / grayscale conversion)
        https://www.ghostscript.com/releases/gsdnld.html
        → Install the 64-bit version (gswin64c).
        → Only needed when printing in black & white (the default).

  Python packages:
    pip install pillow    # for Image → PDF
    pip install pywin32   # for listing printers

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
USAGE
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
  List printers:
    python print_pdf.py printers

  Print a PDF:
    python print_pdf.py print "report.pdf"
    python print_pdf.py print "report.pdf" --printer "HP LaserJet"
    python print_pdf.py print "report.pdf" --color color      # full color
    python print_pdf.py print "report.pdf" --color bw         # black & white (default)
    python print_pdf.py print "report.pdf" --copies 3         # print 3 copies

  Convert to PDF only (no printing):
    python print_pdf.py toxpdf    "data.xlsx"    --out "C:/exports"
    python print_pdf.py pptpdf    "slides.pptx"  --out "C:/exports"
    python print_pdf.py wordpdf   "report.docx"  --out "C:/exports"
    python print_pdf.py imgpdf    "photo.jpg"    --out "C:/exports"

  Download / save a PDF copy:
    python print_pdf.py download "report.pdf" --out "C:/Users/lalit/Downloads"

  Convert + Print in one step:
    python print_pdf.py xlprint   "data.xlsx"   --out "C:/Users/lalit/Downloads"
    python print_pdf.py pptprint  "slides.pptx" --out "C:/Users/lalit/Downloads"
    python print_pdf.py wordprint "report.docx" --out "C:/Users/lalit/Downloads"
    python print_pdf.py imgprint  "photo.jpg"   --out "C:/Users/lalit/Downloads"

    # With color, copies, and printer options:
    python print_pdf.py xlprint "data.xlsx" --printer "HP LaserJet" --color color --copies 2 --out "C:/exports"
    python print_pdf.py print   "report.pdf" --copies 3               # print 3 copies
    python print_pdf.py xlprint "data.xlsx"  --copies 5 --color color  # 5 color copies
"""

import sys
import os
import shutil
import subprocess
import platform
import argparse
import tempfile


IMAGE_EXTS = [".jpg", ".jpeg", ".png", ".bmp", ".tiff", ".tif", ".gif", ".webp"]
EXCEL_EXTS = [".xlsx", ".xls", ".xlsm"]
PPT_EXTS   = [".pptx", ".ppt", ".pptm", ".odp"]
WORD_EXTS  = [".docx", ".doc", ".docm", ".odt", ".rtf"]


# ─────────────────────────────────────────────────────────────────────────────
# Helpers
# ─────────────────────────────────────────────────────────────────────────────

def _require_windows():
    if platform.system() != "Windows":
        raise OSError("This script currently supports Windows only.")


def _check_file(path: str, exts=None):
    if not os.path.isfile(path):
        raise FileNotFoundError(f"File not found: {path}")
    if exts and not any(path.lower().endswith(e) for e in exts):
        raise ValueError(f"Expected file with extension {exts}, got: {path}")


def _find_sumatra(hint: str = None) -> str:
    candidates = []
    if hint:
        candidates.append(hint)
    script_dir = os.path.dirname(os.path.abspath(__file__))
    candidates.append(os.path.join(script_dir, "SumatraPDF.exe"))
    candidates += [
        r"C:\Program Files\SumatraPDF\SumatraPDF.exe",
        r"C:\Program Files (x86)\SumatraPDF\SumatraPDF.exe",
    ]
    for p in candidates:
        if os.path.isfile(p):
            return p
    result = subprocess.run(["where", "SumatraPDF"], capture_output=True, text=True)
    if result.returncode == 0:
        return result.stdout.strip().splitlines()[0]
    raise FileNotFoundError(
        "SumatraPDF.exe not found.\n"
        "  • Download: https://www.sumatrapdfreader.org/download-free-pdf-viewer\n"
        "  • Place SumatraPDF.exe next to this script, OR use --sumatra <path>"
    )


def _find_libreoffice() -> str:
    candidates = [
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
    ]
    for p in candidates:
        if os.path.isfile(p):
            return p
    result = subprocess.run(["where", "soffice"], capture_output=True, text=True)
    if result.returncode == 0:
        return result.stdout.strip().splitlines()[0]
    raise FileNotFoundError(
        "LibreOffice (soffice.exe) not found.\n"
        "  • Download: https://www.libreoffice.org/download/download/\n"
        "  • Install it, then re-run this script."
    )



def _apply_grayscale_to_pdf(pdf_path: str) -> str:
    """
    Convert a PDF to grayscale (black & white) in-place using Ghostscript.

    How it works:
      - Ghostscript re-renders every page of the PDF using a gray color model.
      - The result overwrites the original file (via a temp file).
      - This guarantees NO color ink is used when the file is sent to the printer,
        regardless of the printer driver's own color settings.

    Falls back gracefully if Ghostscript is not installed — prints a warning
    and leaves the PDF untouched.
    Returns the (same) pdf_path.
    """
    gs_candidates = [
        r"C:\Program Files\gs\gs10.04.0\bin\gswin64c.exe",
        r"C:\Program Files\gs\gs10.03.1\bin\gswin64c.exe",
        r"C:\Program Files\gs\gs10.02.1\bin\gswin64c.exe",
        r"C:\Program Files (x86)\gs\gs10.04.0\bin\gswin32c.exe",
        r"C:\Program Files (x86)\gs\gs10.03.1\bin\gswin32c.exe",
    ]
    gs = None
    for p in gs_candidates:
        if os.path.isfile(p):
            gs = p
            break
    if not gs:
        res = subprocess.run(["where", "gswin64c"], capture_output=True, text=True)
        if res.returncode == 0:
            gs = res.stdout.strip().splitlines()[0]
    if not gs:
        res = subprocess.run(["where", "gswin32c"], capture_output=True, text=True)
        if res.returncode == 0:
            gs = res.stdout.strip().splitlines()[0]

    if not gs:
        print("  ⚠ Ghostscript not found — skipping grayscale conversion.")
        print("    Download: https://www.ghostscript.com/releases/gsdnld.html")
        return pdf_path

    import tempfile as _tmp
    with _tmp.NamedTemporaryFile(suffix=".pdf", delete=False) as tf:
        tmp_out = tf.name

    try:
        cmd = [
            gs, "-q", "-dBATCH", "-dNOPAUSE", "-dSAFER",
            "-sDEVICE=pdfwrite",
            "-sColorConversionStrategy=Gray",
            "-dProcessColorModel=/DeviceGray",
            f"-sOutputFile={tmp_out}",
            pdf_path,
        ]
        result = subprocess.run(cmd, capture_output=True, text=True)
        if result.returncode != 0:
            print(f"  ⚠ Ghostscript grayscale failed: {result.stderr.strip()}")
            return pdf_path
        shutil.move(tmp_out, pdf_path)
        print("  ✓ Converted to grayscale (black & white).")
    except Exception as e:
        print(f"  ⚠ Grayscale conversion error: {e}")
    finally:
        if os.path.isfile(tmp_out):
            os.remove(tmp_out)

    return pdf_path


def get_available_printers():
    import win32print
    return [p[2] for p in win32print.EnumPrinters(
        win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS
    )]


# ─────────────────────────────────────────────────────────────────────────────
# Feature 1 — Print PDF
# ─────────────────────────────────────────────────────────────────────────────

def print_pdf(pdf_path: str, printer_name: str = None, sumatra_hint: str = None, color: str = "bw", copies: int = 1):
    """Send a PDF to a printer using SumatraPDF. color='bw' or 'color'. copies=number of copies to print."""
    _require_windows()
    _check_file(pdf_path, [".pdf"])

    sumatra = _find_sumatra(sumatra_hint)

    print(f"  PDF      : {os.path.abspath(pdf_path)}")
    print(f"  Printer  : {printer_name or '(system default)'}")
    print(f"  Color    : {'Color' if color == 'color' else 'Black & White (default)'}")
    print(f"  Copies   : {copies}")
    print(f"  Sumatra  : {sumatra}")

    # --color bw  (default): strip all color from the PDF before printing.
    # --color color         : send the PDF as-is for full-color printing.
    # If Ghostscript is missing, a warning is shown and we continue without grayscale.
    if color != "color":
        _apply_grayscale_to_pdf(pdf_path)

    # Send the job once per requested copy.
    # SumatraPDF does not have a native -copies flag, so we submit the job
    # multiple times — each submission is one complete copy of the document.
    for i in range(copies):
        if copies > 1:
            print(f"  Printing copy {i + 1} of {copies} …")
        if printer_name:
            cmd = [sumatra, "-print-to", printer_name, "-silent", pdf_path]
        else:
            cmd = [sumatra, "-print-to-default", "-silent", pdf_path]

        result = subprocess.run(cmd, capture_output=True, text=True)
        if result.returncode != 0:
            raise RuntimeError(
                f"SumatraPDF failed on copy {i + 1} (exit {result.returncode}):\n"
                f"  {result.stderr.strip() or result.stdout.strip()}"
            )

    print(f"  ✓ {copies} copy/copies sent to printer successfully.")


# ─────────────────────────────────────────────────────────────────────────────
# Feature 2 — Convert Excel → PDF
# ─────────────────────────────────────────────────────────────────────────────

def excel_to_pdf(excel_path: str, out_dir: str = None) -> str:
    """Convert an Excel file to PDF using LibreOffice."""
    _require_windows()
    _check_file(excel_path, EXCEL_EXTS)

    soffice = _find_libreoffice()
    abs_excel = os.path.abspath(excel_path)
    ext = os.path.splitext(abs_excel)[1]
    orig_base = os.path.splitext(os.path.basename(abs_excel))[0]

    if out_dir:
        os.makedirs(out_dir, exist_ok=True)
    else:
        out_dir = os.path.dirname(abs_excel)

    print(f"  Excel      : {abs_excel}")
    print(f"  Output dir : {out_dir}")
    print(f"  LibreOffice: {soffice}")

    with tempfile.TemporaryDirectory() as tmp:
        safe_copy = os.path.join(tmp, "input" + ext)
        shutil.copy2(abs_excel, safe_copy)

        cmd = [soffice, "--headless", "--convert-to", "pdf", "--outdir", tmp, safe_copy]
        result = subprocess.run(cmd, capture_output=True, text=True)

        if result.returncode != 0:
            raise RuntimeError(
                f"LibreOffice conversion failed (exit {result.returncode}):\n"
                f"  {result.stderr.strip() or result.stdout.strip()}"
            )

        tmp_pdf = os.path.join(tmp, "input.pdf")
        if not os.path.isfile(tmp_pdf):
            raise RuntimeError(
                f"Conversion succeeded but PDF not found.\n"
                f"LibreOffice output: {result.stdout.strip()}"
            )

        final_pdf = os.path.join(out_dir, orig_base + ".pdf")
        shutil.move(tmp_pdf, final_pdf)

    print(f"  ✓ PDF created : {final_pdf}")
    return final_pdf



# ─────────────────────────────────────────────────────────────────────────────
# Feature 3 — Convert PowerPoint → PDF
# ─────────────────────────────────────────────────────────────────────────────

def ppt_to_pdf(ppt_path: str, out_dir: str = None) -> str:
    """
    Convert a PowerPoint file (.pptx/.ppt/.pptm/.odp) to PDF using LibreOffice.
    Each slide becomes one page in the PDF.
    Returns the path of the generated PDF.
    """
    _require_windows()
    _check_file(ppt_path, PPT_EXTS)

    soffice = _find_libreoffice()
    abs_ppt = os.path.abspath(ppt_path)
    ext = os.path.splitext(abs_ppt)[1]
    orig_base = os.path.splitext(os.path.basename(abs_ppt))[0]

    if out_dir:
        os.makedirs(out_dir, exist_ok=True)
    else:
        out_dir = os.path.dirname(abs_ppt)

    print(f"  PowerPoint : {abs_ppt}")
    print(f"  Output dir : {out_dir}")
    print(f"  LibreOffice: {soffice}")

    with tempfile.TemporaryDirectory() as tmp:
        safe_copy = os.path.join(tmp, "input" + ext)
        shutil.copy2(abs_ppt, safe_copy)

        cmd = [soffice, "--headless", "--convert-to", "pdf", "--outdir", tmp, safe_copy]
        result = subprocess.run(cmd, capture_output=True, text=True)

        if result.returncode != 0:
            raise RuntimeError(
                f"LibreOffice conversion failed (exit {result.returncode}):\n"
                f"  {result.stderr.strip() or result.stdout.strip()}"
            )

        tmp_pdf = os.path.join(tmp, "input.pdf")
        if not os.path.isfile(tmp_pdf):
            raise RuntimeError(
                f"Conversion succeeded but PDF not found.\n"
                f"LibreOffice output: {result.stdout.strip()}"
            )

        final_pdf = os.path.join(out_dir, orig_base + ".pdf")
        shutil.move(tmp_pdf, final_pdf)

    print(f"  \u2713 PDF created : {final_pdf}")
    return final_pdf


# ─────────────────────────────────────────────────────────────────────────────
# Feature 4 — Convert Image → PDF
# ─────────────────────────────────────────────────────────────────────────────

# ─────────────────────────────────────────────────────────────────────────────
# Feature 4 — Convert Image → PDF
# ─────────────────────────────────────────────────────────────────────────────

def image_to_pdf(image_path: str, out_dir: str = None) -> str:
    """
    Convert an image file (JPG, PNG, BMP, TIFF, GIF, WEBP) to a PDF
    using Pillow. The image is fitted to A4 size.
    Returns the path of the generated PDF.
    """
    try:
        from PIL import Image
    except ImportError:
        raise RuntimeError(
            "Pillow is not installed.\n"
            "  Run: pip install pillow"
        )

    _check_file(image_path, IMAGE_EXTS)

    abs_img = os.path.abspath(image_path)
    orig_base = os.path.splitext(os.path.basename(abs_img))[0]

    if out_dir:
        os.makedirs(out_dir, exist_ok=True)
    else:
        out_dir = os.path.dirname(abs_img)

    print(f"  Image      : {abs_img}")
    print(f"  Output dir : {out_dir}")

    # A4 at 150 DPI → 1240 x 1754 px
    A4_W, A4_H = 1240, 1754

    img = Image.open(abs_img)

    # Convert to RGB (handles PNG transparency, RGBA, palette modes)
    if img.mode in ("RGBA", "P", "LA"):
        background = Image.new("RGB", img.size, (255, 255, 255))
        if img.mode == "P":
            img = img.convert("RGBA")
        background.paste(img, mask=img.split()[-1] if img.mode in ("RGBA", "LA") else None)
        img = background
    elif img.mode != "RGB":
        img = img.convert("RGB")

    # Fit image inside A4, preserving aspect ratio
    img.thumbnail((A4_W, A4_H), Image.LANCZOS)

    # Place on white A4 canvas, centred
    canvas = Image.new("RGB", (A4_W, A4_H), (255, 255, 255))
    x = (A4_W - img.width) // 2
    y = (A4_H - img.height) // 2
    canvas.paste(img, (x, y))

    final_pdf = os.path.join(out_dir, orig_base + ".pdf")
    canvas.save(final_pdf, "PDF", resolution=150)

    print(f"  ✓ PDF created : {final_pdf}")
    return final_pdf


# ─────────────────────────────────────────────────────────────────────────────
# Feature 5 — Download / copy a PDF
# ─────────────────────────────────────────────────────────────────────────────

def download_pdf(pdf_path: str, out_dir: str) -> str:
    """Copy a PDF to out_dir. Returns the destination path."""
    _check_file(pdf_path, [".pdf"])
    os.makedirs(out_dir, exist_ok=True)

    dest = os.path.join(out_dir, os.path.basename(pdf_path))
    shutil.copy2(pdf_path, dest)

    print(f"  Source   : {os.path.abspath(pdf_path)}")
    print(f"  Saved to : {dest}")
    print("  ✓ PDF downloaded (copied) successfully.")
    return dest


# ─────────────────────────────────────────────────────────────────────────────
# Feature 6 — Excel → PDF → Print
# ─────────────────────────────────────────────────────────────────────────────

def excel_print(excel_path: str, printer_name: str = None,
                out_dir: str = None, sumatra_hint: str = None, color: str = "bw", copies: int = 1):
    """Convert Excel to PDF, then print it."""
    print("[Step 1/2] Converting Excel to PDF …")
    pdf_path = excel_to_pdf(excel_path, out_dir)
    print("[Step 2/2] Sending PDF to printer …")
    print_pdf(pdf_path, printer_name, sumatra_hint, color, copies)


# ─────────────────────────────────────────────────────────────────────────────
# Feature 7 — Image → PDF → Print
# ─────────────────────────────────────────────────────────────────────────────

def image_print(image_path: str, printer_name: str = None,
                out_dir: str = None, sumatra_hint: str = None, color: str = "bw", copies: int = 1):
    """Convert image to PDF, then print it."""
    print("[Step 1/2] Converting image to PDF …")
    pdf_path = image_to_pdf(image_path, out_dir)
    print("[Step 2/2] Sending PDF to printer …")
    print_pdf(pdf_path, printer_name, sumatra_hint, color, copies)



# ─────────────────────────────────────────────────────────────────────────────
# Feature 8 — PowerPoint → PDF → Print
# ─────────────────────────────────────────────────────────────────────────────

def ppt_print(ppt_path: str, printer_name: str = None,
              out_dir: str = None, sumatra_hint: str = None):
    """Convert PowerPoint to PDF, then print it."""
    print("[Step 1/2] Converting PowerPoint to PDF \u2026")
    pdf_path = ppt_to_pdf(ppt_path, out_dir)
    print("[Step 2/2] Sending PDF to printer \u2026")
    print_pdf(pdf_path, printer_name, sumatra_hint)


# ─────────────────────────────────────────────────────────────────────────────
# Feature 9 — Convert Word → PDF
# ─────────────────────────────────────────────────────────────────────────────

def word_to_pdf(word_path: str, out_dir: str = None) -> str:
    """
    Convert a Word file (.docx/.doc/.odt/.rtf) to PDF using LibreOffice.
    Returns the path of the generated PDF.
    """
    _require_windows()
    _check_file(word_path, WORD_EXTS)

    soffice = _find_libreoffice()
    abs_word = os.path.abspath(word_path)
    ext = os.path.splitext(abs_word)[1]
    orig_base = os.path.splitext(os.path.basename(abs_word))[0]

    if out_dir:
        os.makedirs(out_dir, exist_ok=True)
    else:
        out_dir = os.path.dirname(abs_word)

    print(f"  Word       : {abs_word}")
    print(f"  Output dir : {out_dir}")
    print(f"  LibreOffice: {soffice}")

    with tempfile.TemporaryDirectory() as tmp:
        safe_copy = os.path.join(tmp, "input" + ext)
        shutil.copy2(abs_word, safe_copy)

        cmd = [soffice, "--headless", "--convert-to", "pdf", "--outdir", tmp, safe_copy]
        result = subprocess.run(cmd, capture_output=True, text=True)

        if result.returncode != 0:
            raise RuntimeError(
                f"LibreOffice conversion failed (exit {result.returncode}):\n"
                f"  {result.stderr.strip() or result.stdout.strip()}"
            )

        tmp_pdf = os.path.join(tmp, "input.pdf")
        if not os.path.isfile(tmp_pdf):
            raise RuntimeError(
                f"Conversion succeeded but PDF not found.\n"
                f"LibreOffice output: {result.stdout.strip()}"
            )

        final_pdf = os.path.join(out_dir, orig_base + ".pdf")
        shutil.move(tmp_pdf, final_pdf)

    print(f"  ✓ PDF created : {final_pdf}")
    return final_pdf


# ─────────────────────────────────────────────────────────────────────────────
# Feature 10 — Word → PDF → Print
# ─────────────────────────────────────────────────────────────────────────────

def word_print(word_path: str, printer_name: str = None,
               out_dir: str = None, sumatra_hint: str = None, color: str = "bw", copies: int = 1):
    """Convert Word document to PDF, then print it."""
    print("[Step 1/2] Converting Word to PDF …")
    pdf_path = word_to_pdf(word_path, out_dir)
    print("[Step 2/2] Sending PDF to printer …")
    print_pdf(pdf_path, printer_name, sumatra_hint, color, copies)

# ─────────────────────────────────────────────────────────────────────────────
# CLI
# ─────────────────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(
        prog="print_pdf.py",
        description="PDF, Excel & Image printing utility for Windows.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=__doc__,
    )
    sub = parser.add_subparsers(dest="command", metavar="COMMAND")

    # printers
    sub.add_parser("printers", help="List all available printers.")

    # print
    p_print = sub.add_parser("print", help="Print a PDF file.")
    p_print.add_argument("pdf", help="Path to the PDF file.")
    p_print.add_argument("--printer", "-p", help="Printer name (default: system default).")
    p_print.add_argument("--sumatra", help="Path to SumatraPDF.exe.")
    p_print.add_argument("--copies", "-n", type=int, default=1, metavar="N",
                        help="Number of copies to print (default: 1).")
    p_print.add_argument("--color", choices=["bw", "color"], default="bw",
                        help='Color mode: "bw" = black & white (default, uses Ghostscript to strip color), '
                             '"color" = full color. If --color is omitted, bw is used.')

    # toxpdf  (Excel → PDF)
    p_toxpdf = sub.add_parser("toxpdf", help="Convert Excel file to PDF.")
    p_toxpdf.add_argument("excel", help="Path to the .xlsx / .xls file.")
    p_toxpdf.add_argument("--out", "-o", help="Output folder (default: same folder as Excel).")

    # pptpdf  (PowerPoint → PDF)
    p_pptpdf = sub.add_parser("pptpdf", help="Convert PowerPoint file to PDF.")
    p_pptpdf.add_argument("ppt", help="Path to the .pptx / .ppt file.")
    p_pptpdf.add_argument("--out", "-o", help="Output folder (default: same folder as PPT).")

    # wordpdf (Word → PDF)
    p_wordpdf = sub.add_parser("wordpdf", help="Convert Word file to PDF.")
    p_wordpdf.add_argument("word", help="Path to the .docx / .doc file.")
    p_wordpdf.add_argument("--out", "-o", help="Output folder (default: same folder as Word file).")

    # imgpdf  (Image → PDF)
    p_imgpdf = sub.add_parser("imgpdf", help="Convert image file to PDF.")
    p_imgpdf.add_argument("image", help="Path to the image file (JPG/PNG/BMP/TIFF/GIF/WEBP).")
    p_imgpdf.add_argument("--out", "-o", help="Output folder (default: same folder as image).")

    # download
    p_dl = sub.add_parser("download", help="Save/download a copy of a PDF.")
    p_dl.add_argument("pdf", help="Path to the PDF file.")
    p_dl.add_argument("--out", "-o", required=True, help="Destination folder.")

    # xlprint (Excel → PDF → print)
    p_xlp = sub.add_parser("xlprint", help="Convert Excel to PDF and print it.")
    p_xlp.add_argument("excel", help="Path to the .xlsx / .xls file.")
    p_xlp.add_argument("--printer", "-p", help="Printer name (default: system default).")
    p_xlp.add_argument("--out", "-o", help="Folder to save the PDF.")
    p_xlp.add_argument("--sumatra", help="Path to SumatraPDF.exe.")
    p_xlp.add_argument("--copies", "-n", type=int, default=1, metavar="N",
                        help="Number of copies to print (default: 1).")
    p_xlp.add_argument("--color", choices=["bw", "color"], default="bw",
                        help='Color mode: "bw" = black & white (default, uses Ghostscript to strip color), '
                             '"color" = full color. If --color is omitted, bw is used.')

    # pptprint (PowerPoint → PDF → print)
    p_pptp = sub.add_parser("pptprint", help="Convert PowerPoint to PDF and print it.")
    p_pptp.add_argument("ppt", help="Path to the .pptx / .ppt file.")
    p_pptp.add_argument("--printer", "-p", help="Printer name (default: system default).")
    p_pptp.add_argument("--out", "-o", help="Folder to save the PDF.")
    p_pptp.add_argument("--sumatra", help="Path to SumatraPDF.exe.")
    p_pptp.add_argument("--copies", "-n", type=int, default=1, metavar="N",
                        help="Number of copies to print (default: 1).")
    p_pptp.add_argument("--color", choices=["bw", "color"], default="bw",
                        help='Color mode: "bw" = black & white (default, uses Ghostscript to strip color), '
                             '"color" = full color. If --color is omitted, bw is used.')

    # wordprint (Word → PDF → print)
    p_wordp = sub.add_parser("wordprint", help="Convert Word file to PDF and print it.")
    p_wordp.add_argument("word", help="Path to the .docx / .doc file.")
    p_wordp.add_argument("--printer", "-p", help="Printer name (default: system default).")
    p_wordp.add_argument("--out", "-o", help="Folder to save the PDF.")
    p_wordp.add_argument("--sumatra", help="Path to SumatraPDF.exe.")
    p_wordp.add_argument("--copies", "-n", type=int, default=1, metavar="N",
                        help="Number of copies to print (default: 1).")
    p_wordp.add_argument("--color", choices=["bw", "color"], default="bw",
                        help='Color mode: "bw" = black & white (default, uses Ghostscript to strip color), '
                             '"color" = full color. If --color is omitted, bw is used.')

    # imgprint (Image → PDF → print)
    p_imp = sub.add_parser("imgprint", help="Convert image to PDF and print it.")
    p_imp.add_argument("image", help="Path to the image file (JPG/PNG/BMP/TIFF/GIF/WEBP).")
    p_imp.add_argument("--printer", "-p", help="Printer name (default: system default).")
    p_imp.add_argument("--out", "-o", help="Folder to save the PDF.")
    p_imp.add_argument("--sumatra", help="Path to SumatraPDF.exe.")
    p_imp.add_argument("--copies", "-n", type=int, default=1, metavar="N",
                        help="Number of copies to print (default: 1).")
    p_imp.add_argument("--color", choices=["bw", "color"], default="bw",
                        help='Color mode: "bw" = black & white (default, uses Ghostscript to strip color), '
                             '"color" = full color. If --color is omitted, bw is used.')

    args = parser.parse_args()

    if not args.command:
        parser.print_help()
        sys.exit(0)

    try:
        if args.command == "printers":
            printers = get_available_printers()
            if printers:
                print("Available printers:")
                for p in printers:
                    print(f"  • {p}")
            else:
                print("No printers found.")

        elif args.command == "print":
            print_pdf(args.pdf, args.printer, args.sumatra, args.color, args.copies)

        elif args.command == "toxpdf":
            excel_to_pdf(args.excel, args.out)

        elif args.command == "pptpdf":
            ppt_to_pdf(args.ppt, args.out)

        elif args.command == "wordpdf":
            word_to_pdf(args.word, args.out)

        elif args.command == "imgpdf":
            image_to_pdf(args.image, args.out)

        elif args.command == "download":
            download_pdf(args.pdf, args.out)

        elif args.command == "xlprint":
            excel_print(args.excel, args.printer, args.out, args.sumatra, args.color, args.copies)

        elif args.command == "pptprint":
            ppt_print(args.ppt, args.printer, args.out, args.sumatra, args.color, args.copies)

        elif args.command == "wordprint":
            word_print(args.word, args.printer, args.out, args.sumatra, args.color, args.copies)

        elif args.command == "imgprint":
            image_print(args.image, args.printer, args.out, args.sumatra, args.color, args.copies)

    except (FileNotFoundError, ValueError, OSError, RuntimeError) as exc:
        print(f"\nError: {exc}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()