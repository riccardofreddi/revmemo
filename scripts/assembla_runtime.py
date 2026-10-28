# -*- coding: utf-8 -*-
"""Assembla runtime/ — Python Windows embeddable + openpyxl, zero installazione."""
import zipfile, glob, os

BUILD = "/home/z/my-project/scripts/revmemo_build"
RUNTIME = "/home/z/my-project/download/RevMemo/runtime"

# 1) decomprimi python embeddable
with zipfile.ZipFile(os.path.join(BUILD, "python_embed.zip")) as z:
    z.extractall(RUNTIME)
print("Python embeddable estratto:", len(os.listdir(RUNTIME)), "voci")

# 2) scrivi python312._pth: percorsi che python.exe caricherà
pth_content = "python312.zip\n.\nLib/site-packages\n"
with open(os.path.join(RUNTIME, "python312._pth"), "w") as f:
    f.write(pth_content)
print("python312._pth scritto")

# 3) estrai le wheel (sono zip) dentro Lib/site-packages
import zipfile as zf
for whl in glob.glob(os.path.join(BUILD, "wheels", "*.whl")):
    with zf.ZipFile(whl) as z:
        z.extractall(os.path.join(RUNTIME, "Lib", "site-packages"))
    print("estratta:", os.path.basename(whl))

# 4) verifica struttura openpyxl
op = os.path.join(RUNTIME, "Lib", "site-packages", "openpyxl")
print("openpyxl presente:", os.path.isdir(op))
print("openpyxl/__init__.py:", os.path.isfile(os.path.join(op, "__init__.py")))
print("ssl DLLs:", [f for f in os.listdir(RUNTIME) if "ssl" in f.lower() or "crypto" in f.lower()])
