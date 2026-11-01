"""Refresh the canonical ODS files with the latest macros and dialogs.

Source of truth: canban_trial.ods (full sample data and current schema).

Pipeline:
  1. Build tmp/change1/base.ods from canban_trial.ods, injecting the final
     macros and the dialogs.
  2. Copy base.ods to canban_blank.ods and canban_trial.ods.
  3. Run seed_blank.py on canban_blank.ods and seed_trial.py on
     canban_trial.ods to (re)generate their data.
"""
import shutil
import zipfile

SRC = "canban_trial.ods"
BASE = "tmp/change1/base.ods"
VARIANTS = ["canban_blank.ods", "canban_trial.ods"]
MACROS = "src/macros.bas"
DIALOG_ITEM = "tmp/change1/DialogPanelItem.xml"
DIALOG_TASK = "tmp/change1/DialogPanelTask.xml"
DIALOG_ACTSTATE = "tmp/change1/DialogPanelActState.xml"

DIALOG_LB = """<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE library:library PUBLIC "-//OpenOffice.org//DTD OfficeDocument 1.0//EN" "library.dtd">
<library:library xmlns:library="http://openoffice.org/2000/library" library:name="Standard" library:readonly="false" library:passwordprotected="false">
 <library:element library:name="DialogPanelTask"/>
 <library:element library:name="DialogPanelItem"/>
 <library:element library:name="DialogPanelActState"/>
</library:library>
"""


def build_module() -> bytes:
    code = open(MACROS, encoding="utf-8-sig").read()
    esc = (
        code.replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace('"', "&quot;")
        .replace("'", "&apos;")
    )
    return (
        '<?xml version="1.0" encoding="UTF-8"?>\n'
        '<!DOCTYPE script:module PUBLIC "-//OpenOffice.org//DTD OfficeDocument 1.0//EN" "module.dtd">\n'
        '<script:module xmlns:script="http://openoffice.org/2000/script" '
        'script:name="Module1" script:language="StarBasic" script:moduleType="normal">'
        + esc
        + "</script:module>"
    ).encode("utf-8")


def add_manifest_entry(manifest: str) -> str:
    if "DialogPanelActState.xml" in manifest:
        return manifest
    entry = (
        ' <manifest:file-entry manifest:full-path="Dialogs/Standard/DialogPanelActState.xml" '
        'manifest:media-type="text/xml"/>\n'
    )
    return manifest.replace(
        ' <manifest:file-entry manifest:full-path="Dialogs/Standard/dialog-lb.xml"',
        entry + ' <manifest:file-entry manifest:full-path="Dialogs/Standard/dialog-lb.xml"',
    )


def main() -> None:
    module = build_module()
    dialog_item = open(DIALOG_ITEM, "rb").read()
    dialog_task = open(DIALOG_TASK, "rb").read()
    dialog_actstate = open(DIALOG_ACTSTATE, "rb").read()

    added = False
    with zipfile.ZipFile(SRC, "r") as zin, zipfile.ZipFile(BASE, "w") as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "Basic/Standard/Module1.xml":
                data = module
            elif item.filename == "Dialogs/Standard/DialogPanelItem.xml":
                data = dialog_item
            elif item.filename == "Dialogs/Standard/DialogPanelTask.xml":
                data = dialog_task
            elif item.filename == "Dialogs/Standard/DialogPanelActState.xml":
                data = dialog_actstate
                added = True
            elif item.filename == "Dialogs/Standard/dialog-lb.xml":
                data = DIALOG_LB.encode("utf-8")
            elif item.filename == "META-INF/manifest.xml":
                data = add_manifest_entry(data.decode("utf-8")).encode("utf-8")
            zout.writestr(item, data)

        if not added:
            zout.writestr("Dialogs/Standard/DialogPanelActState.xml", dialog_actstate)

    for name in VARIANTS:
        shutil.copyfile(BASE, name)
    print("wrote", BASE, "and", ", ".join(VARIANTS))


if __name__ == "__main__":
    main()
