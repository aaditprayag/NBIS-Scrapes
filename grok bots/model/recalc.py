import subprocess, sys, os, shutil, openpyxl
def recalc(src, outdir="/tmp/rc"):
    os.makedirs(outdir, exist_ok=True)
    subprocess.run(["soffice", "--headless", "--calc", "--convert-to", "xlsx:Calc MS Excel 2007 XML", "--outdir", outdir, src], check=True, capture_output=True, timeout=300)
    return os.path.join(outdir, os.path.basename(src))
if __name__ == "__main__":
    p = recalc(sys.argv[1])
    wb = openpyxl.load_workbook(p, data_only=True)
    errs = {}
    for ws in wb:
        for row in ws.iter_rows():
            for c in row:
                if isinstance(c.value, str) and c.value.startswith(("#", "Err:")) and c.value[:4] in ("#REF", "#VAL", "#NAM", "#DIV", "#N/A", "#NUM", "Err:"):
                    errs.setdefault(ws.title, []).append((c.coordinate, c.value))
    for k, v in errs.items(): print(k, len(v), v[:12])
    print("done", p)
