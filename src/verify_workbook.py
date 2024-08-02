"""Recalculate a workbook with LibreOffice (headless) and verify it.

- Writes build/BergWerk_Shopify_Cost_Control_calculated.xlsx (values cached)
- Prints the brief check figures and the model-integrity checks
- Exits with code 1 if any check fails or any formula returns an error

Run:  python src/verify_workbook.py
"""
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

from openpyxl import load_workbook

ROOT = Path(__file__).resolve().parents[1]
SRC = ROOT / "excel" / "BergWerk_Shopify_Cost_Control.xlsx"
BUILD = ROOT / "build"
CALC = BUILD / "BergWerk_Shopify_Cost_Control_calculated.xlsx"
ERRORS = ("#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#NUM!", "#NULL!",
          "Err:")


def find_soffice():
    for n in ("soffice", "libreoffice"):
        if shutil.which(n):
            return shutil.which(n)
    for p in (r"C:\Program Files\LibreOffice\program\soffice.exe",
              r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
              "/Applications/LibreOffice.app/Contents/MacOS/soffice"):
        if Path(p).exists():
            return p
    sys.exit("LibreOffice not found. Install it to recalculate the workbook.")


def recalc(src, dst):
    """Open src in headless LibreOffice, recalculate, save values to dst.
    A throwaway user profile avoids clashes with a running LibreOffice."""
    src, dst = Path(src), Path(dst)
    dst.parent.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory() as tmp:
        profile = Path(tmp, "profile").as_uri()
        outdir = Path(tmp, "out")
        subprocess.run([find_soffice(), f"-env:UserInstallation={profile}",
                        "--headless", "--norestore", "--convert-to",
                        "xlsx:Calc MS Excel 2007 XML", "--outdir",
                        str(outdir), str(src)],
                       check=True, capture_output=True, timeout=300)
        shutil.move(str(outdir / src.name), dst)
    return dst


def named(wb, name):
    """Value of a single-cell named range."""
    (sheet, ref), = wb.defined_names[name].destinations
    return wb[sheet][ref.replace("$", "")].value


def named_rows(wb, name):
    """Values of a multi-cell named range as a list of row lists."""
    (sheet, ref), = wb.defined_names[name].destinations
    return [[c.value for c in row] for row in wb[sheet][ref.replace("$", "")]]


def formula_errors(wb):
    return [f"{ws.title}!{c.coordinate}={c.value}" for ws in wb.worksheets
            for row in ws.iter_rows() for c in row
            if isinstance(c.value, str) and c.value.startswith(ERRORS)]


def main():
    wb = load_workbook(recalc(SRC, CALC), data_only=True)
    ok = True
    print(f"{'Brief check figure':32}{'Model':>14}{'Expected':>14}  Result")
    for m, _, model, exp, res in named_rows(wb, "CHECKS_TABLE"):
        ok &= res == "PASS"
        print(f"{m:32}{str(model):>14}{str(exp):>14}  {res}")
    print(f"\nBrief figures: {named(wb, 'BRIEF_CHECKS')} | "
          f"Model integrity: {named(wb, 'INTEGRITY')}")
    ok &= named(wb, "BRIEF_CHECKS") == "ALL PASS"
    ok &= named(wb, "INTEGRITY") == "ALL PASS"
    print(f"Cost: {named(wb, 'COST_RAG')} | Schedule: "
          f"{named(wb, 'SCHED_RAG')} | Forecast launch: "
          f"{named(wb, 'FORECAST_LAUNCH'):%d %b %Y}")
    errs = formula_errors(wb)
    if errs:
        ok = False
        print("Formula errors:", errs[:20])
    sys.exit(0 if ok else 1)


if __name__ == "__main__":
    main()
