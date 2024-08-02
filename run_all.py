"""Build the workbook, verify it in LibreOffice, then build the report."""
import subprocess
import sys

for step in ("build_workbook.py", "verify_workbook.py", "build_report.py"):
    print(f"\n=== {step} ===")
    subprocess.run([sys.executable, f"src/{step}"], check=True)
