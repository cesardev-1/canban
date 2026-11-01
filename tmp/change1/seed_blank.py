"""Turn a base ODS into the blank template (canban_blank.ods).

Usage: python3 seed_blank.py <path.ods>

It removes every task, activity and catalog entry, then writes only the
essential seed: In progress (base), To do (base), Trash (trash) and the
activity state "active" (none).
"""
import os
import subprocess
import sys
import time

import uno
from com.sun.star.beans import PropertyValue

ODS = os.path.abspath(sys.argv[1] if len(sys.argv) > 1 else "canban_blank.ods")
BASE = os.path.abspath("tmp/change1")
_NAME = os.path.splitext(os.path.basename(ODS))[0]
PORT = 2300 + (sum(ord(c) for c in _NAME) % 200)
PROFILE = os.path.join(BASE, "lo_profile_" + _NAME)

STATES = [
    (1, "In progress", "base", ""),
    (2, "To do", "base", ""),
    (3, "Trash", "trash", ""),
]


def main() -> None:
    os.makedirs(PROFILE, exist_ok=True)
    try:
        os.remove(os.path.join(os.path.dirname(ODS), ".~lock." + os.path.basename(ODS) + "#"))
    except OSError:
        pass

    proc = subprocess.Popen(
        [
            "soffice", "--headless", "--invisible", "--norestore", "--nologo",
            "--nofirststartwizard", "-env:UserInstallation=file://" + PROFILE,
            "--accept=socket,host=localhost,port=%d;urp;" % PORT,
        ],
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
    )
    try:
        local = uno.getComponentContext()
        resolver = local.ServiceManager.createInstanceWithContext(
            "com.sun.star.bridge.UnoUrlResolver", local)
        ctx = None
        for _ in range(120):
            try:
                ctx = resolver.resolve(
                    "uno:socket,host=localhost,port=%d;urp;StarOffice.ComponentContext" % PORT)
                break
            except Exception:
                time.sleep(0.5)
        if ctx is None:
            raise RuntimeError("could not connect to LibreOffice")

        desktop = ctx.ServiceManager.createInstanceWithContext(
            "com.sun.star.frame.Desktop", ctx)
        hidden = PropertyValue(); hidden.Name = "Hidden"; hidden.Value = True
        doc = desktop.loadComponentFromURL(
            uno.systemPathToFileUrl(ODS), "_blank", 0, (hidden,))
        if doc is None:
            raise RuntimeError("could not open " + ODS)

        data = doc.Sheets.getByName("data")
        view = doc.Sheets.getByName("view")
        provider = doc.getScriptProvider()

        # Clear A2 silently first so any content-change event is harmless.
        provider.getScript(
            "vnd.sun.star.script:Standard.Module1.SetActivityFilterSilently?language=Basic&location=document"
        ).invoke(("",), (), ())

        cursor = data.createCursor()
        cursor.gotoEndOfUsedArea(False)
        total_rows = cursor.RangeAddress.EndRow + 1
        if total_rows < 1:
            total_rows = 1

        # Clear tasks (A-F) and catalogs (H-P).
        clear_cols = [0, 1, 2, 3, 4, 5, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16]
        for row in range(1, total_rows):
            for col in clear_cols:
                data.getCellByPosition(col, row).String = ""

        # Essential task states.
        for row, name, role, scope in STATES:
            data.getCellByPosition(9, row).String = name
            data.getCellByPosition(10, row).String = role
            data.getCellByPosition(11, row).String = scope

        # Essential activity state.
        data.getCellByPosition(12, 1).String = "active"
        data.getCellByPosition(13, 1).String = "none"
        data.getCellByPosition(14, 1).String = ""
        data.getCellByPosition(15, 1).String = ""

        # Default board states.
        view.getCellByPosition(0, 2).String = "In progress"
        view.getCellByPosition(1, 2).String = "To do"

        doc.store()
        doc.close(True)
        print("seeded blank", ODS)
    finally:
        proc.terminate()
        try:
            proc.wait(timeout=15)
        except Exception:
            proc.kill()


if __name__ == "__main__":
    main()
