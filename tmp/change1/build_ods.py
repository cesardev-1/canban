"""Build the base ODS and copy it into the three deliverables.

Pipeline:
  1. This script builds tmp/change1/base.ods from canban.ods, injecting the
     final macros, the dialogs, the data headers and the event bindings.
  2. base.ods is copied to canban_blank.ods, canban_trial.ods and
     canban_test.ods.
  3. Run seed_blank.py on canban_blank.ods and seed_trial.py on
     canban_trial.ods / canban_test.ods to populate the sample data.
"""
import html
import re
import shutil
import zipfile

SRC = "canban.ods"
BASE = "tmp/change1/base.ods"
VARIANTS = ["canban_blank.ods", "canban_trial.ods", "canban_test.ods"]
MACROS = "src/macros.bas"
DIALOG_ITEM = "tmp/change1/DialogPanelItem.xml"
DIALOG_TASK = "tmp/change1/DialogPanelTask.xml"
DIALOG_ACTSTATE = "tmp/change1/DialogPanelActState.xml"

# Logical data rows that receive a K value (row index -> role text).
ROLE_ROWS = {1: "base", 2: "base", 11: "trash"}

HEADERS = [
    "Activity", "Task", "State", "Description", "Order", "ID_Task", "",
    "Activities", "ActState", "States", "StateRole", "StateScope",
    "ActStateName", "ActStateType", "ActStateParam", "ActStateMessage",
]

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


def cell(value: str) -> str:
    if value == "":
        return "<table:table-cell/>"
    return (
        '<table:table-cell office:value-type="string" '
        'calcext:value-type="string"><text:p>%s</text:p></table:table-cell>'
        % html.escape(value, quote=False)
    )


def rebuild_header_row(row: str) -> str:
    open_tag = row[: row.index(">") + 1]
    parts = [open_tag]
    for value in HEADERS:
        parts.append(cell(value))
    parts.append('<table:table-cell table:number-columns-repeated="%d"/>' % (16384 - len(HEADERS)))
    parts.append("</table:table-row>")
    return "".join(parts)


def seed_k(content: str) -> str:
    match = re.search(r'(<table:table table:name="data".*?</table:table>)', content, re.S)
    if not match:
        raise RuntimeError("data table not found")
    table = match.group(1)
    rows = list(re.finditer(r"<table:table-row\b.*?</table:table-row>", table, re.S))

    marker = '<table:table-cell table:number-columns-repeated="16374"/>'
    out = []
    last = 0
    for index, row_match in enumerate(rows):
        row = row_match.group(0)
        if index == 0:
            new_row = rebuild_header_row(row)
        elif index in ROLE_ROWS:
            pos = row.rfind(marker)
            if pos == -1:
                raise RuntimeError("trailing cell not found in row element %d" % index)
            new_row = (
                row[:pos]
                + cell(ROLE_ROWS[index])
                + '<table:table-cell table:number-columns-repeated="16373"/>'
                + row[pos + len(marker):]
            )
        else:
            new_row = row
        out.append(table[last:row_match.start()])
        out.append(new_row)
        last = row_match.end()
    out.append(table[last:])
    new_table = "".join(out)
    return content[: match.start(1)] + new_table + content[match.end(1):]


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
            elif item.filename == "content.xml":
                content = data.decode("utf-8")
                content = content.replace(
                    "Standard.Module1.StartListenerSeleccion",
                    "Standard.Module1.StartListenerSelection",
                )
                content = content.replace(
                    "Standard.Module1.RemoveListenerSeleccion",
                    "Standard.Module1.RemoveListenerSelection",
                )
                data = seed_k(content).encode("utf-8")
            zout.writestr(item, data)

        if not added:
            zout.writestr("Dialogs/Standard/DialogPanelActState.xml", dialog_actstate)

    for name in VARIANTS:
        shutil.copyfile(BASE, name)
    print("wrote", BASE, "and", ", ".join(VARIANTS))


if __name__ == "__main__":
    main()
