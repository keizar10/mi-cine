#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Recalcula todas las fórmulas del Excel con LibreOffice y guarda los valores
en el archivo. Sin este paso, Google Drive / el móvil mostrarían 0 en las
columnas de fórmulas (Puntos premios, hoja Análisis…).

USO:  python3 recalcular.py "Movie Data Base by Keizar.xlsx"
REQUISITOS: LibreOffice instalado (soffice en el PATH).
"""
import os
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

MACRO = """<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE script:module PUBLIC "-//OpenOffice.org//DTD OfficeDocument 1.0//EN" "module.dtd">
<script:module xmlns:script="http://openoffice.org/2000/script" script:name="Module1" script:language="StarBasic">
    Sub RecalculateAndSave()
      ThisComponent.calculateAll()
      ThisComponent.store()
      ThisComponent.close(True)
    End Sub
</script:module>"""


def main(path):
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if not soffice:
        sys.exit("ERROR: no encuentro LibreOffice (soffice).")

    xlsx = Path(path).resolve()
    if not xlsx.exists():
        sys.exit(f"ERROR: no existe {xlsx}")

    with tempfile.TemporaryDirectory() as tmp:
        profile = Path(tmp) / "perfil"
        url = profile.as_uri()

        # 1) crear el perfil de LibreOffice
        subprocess.run([soffice, "--headless", "--terminate_after_init",
                        f"-env:UserInstallation={url}"],
                       capture_output=True, timeout=120)

        macro_dir = profile / "user" / "basic" / "Standard"
        macro_dir.mkdir(parents=True, exist_ok=True)
        (macro_dir / "Module1.xba").write_text(MACRO, encoding="utf-8")

        # registrar el módulo en la biblioteca Standard
        script_xlb = macro_dir / "script.xlb"
        script_xlb.write_text(
            '<?xml version="1.0" encoding="UTF-8"?>\n'
            '<!DOCTYPE library:library PUBLIC "-//OpenOffice.org//DTD OfficeDocument 1.0//EN" "library.dtd">\n'
            '<library:library xmlns:library="http://openoffice.org/2000/library" '
            'library:name="Standard" library:readonly="false" library:passwordprotected="false">\n'
            ' <library:element library:name="Module1"/>\n'
            '</library:library>', encoding="utf-8")

        antes = xlsx.stat().st_mtime
        # 2) abrir el Excel, recalcular, guardar y cerrar
        r = subprocess.run([soffice, "--headless", "--norestore",
                            f"-env:UserInstallation={url}",
                            "macro:///Standard.Module1.RecalculateAndSave",
                            str(xlsx)],
                           capture_output=True, text=True, timeout=600)

    if xlsx.stat().st_mtime <= antes:
        print(r.stdout, r.stderr)
        sys.exit("ERROR: LibreOffice no guardó el archivo; fórmulas sin recalcular.")
    print(f"OK: fórmulas recalculadas en {xlsx.name}")


if __name__ == "__main__":
    main(sys.argv[1] if len(sys.argv) > 1 else "Movie Data Base by Keizar.xlsx")
