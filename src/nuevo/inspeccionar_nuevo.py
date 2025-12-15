import pandas as pd
from pathlib import Path

FILES = [
    Path(__file__).parent / "455 Report_AseguradoraSaldos_Unificado noviembre 2025 5924.xlsx",
    Path(__file__).parent / "plantillaTC.xlsx",
]

ROWS = 8

for f in FILES:
    print(f"\n=== {f.name}")
    if not f.exists():
        print("[ERROR] No existe")
        continue
    xls = pd.ExcelFile(f)
    print("Hojas:", xls.sheet_names)
    for sheet in xls.sheet_names[:2]:
        df = xls.parse(sheet, nrows=ROWS, header=None)
        print(f"-- {sheet} (primeras {ROWS} filas)")
        print(df.to_string(index=False))
