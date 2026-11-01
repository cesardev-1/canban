#!/usr/bin/env python3
"""Migrate user data from an old Canban ODS into the current version.

The script reads the old file's data sheet (tasks, activities, task states,
roles and, when present, activity states) and the view filters, then writes
them into a copy of the current template (canban_blank.ods).

It is pure Python: no running LibreOffice instance is required.

Usage:
    python3 migrate/migrate_ods.py
    python3 migrate/migrate_ods.py --source old.ods --dest out_dir --name result.ods

If arguments are omitted, the script prompts for them.
"""
import argparse
import html
import os
import re
import sys
import zipfile

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
PROJECT_DIR = os.path.dirname(SCRIPT_DIR)
BLANK_TEMPLATE = os.path.join(PROJECT_DIR, "canban_blank.ods")
DEFAULT_DEST = os.path.join(SCRIPT_DIR, "migration_result")

HEADERS = [
    "Activity", "Task", "State", "Description", "Order", "ID_Task", "",
    "Activities", "ActState", "States", "StateRole", "StateScope",
    "ActStateName", "ActStateType", "ActStateParam", "ActStateMessage",
]

# Legacy state names -> role (used when the old file has no StateRole column).
LEGACY_ROLES = {
    "trash": "trash",
    "papelera": "trash",
    "in progress": "base",
    "en progreso": "base",
    "to do": "base",
    "por hacer": "base",
}


# ---------------------------------------------------------------------------
# XML helpers
# ---------------------------------------------------------------------------
def xml_text(fragment: str) -> str:
    parts = re.findall(r"<text:p[^>]*>(.*?)</text:p>", fragment, re.S)
    if not parts:
        return ""
    return " ".join(
        html.unescape(re.sub(r"<[^>]+>", "", p)).strip() for p in parts
    ).strip()


def cell_xml(value: str) -> str:
    if value == "":
        return "<table:table-cell/>"
    return (
        '<table:table-cell office:value-type="string" '
        'calcext:value-type="string"><text:p>%s</text:p></table:table-cell>'
        % html.escape(value, quote=False)
    )


def number_xml(value: str) -> str:
    value = str(value).strip()
    if value == "":
        return "<table:table-cell/>"
    try:
        float(value)
    except ValueError:
        return cell_xml(value)
    return (
        '<table:table-cell office:value-type="float" office:value="%s" '
        'calcext:value-type="float"><text:p>%s</text:p></table:table-cell>'
        % (html.escape(value, quote=False), html.escape(value, quote=False))
    )


# ---------------------------------------------------------------------------
# ODS reading
# ---------------------------------------------------------------------------
def get_table(content: str, name: str) -> str:
    match = re.search(
        r'<table:table\b[^>]*table:name="%s".*?</table:table>' % re.escape(name),
        content,
        re.S,
    )
    if not match:
        raise RuntimeError("table %r not found" % name)
    return match.group(0)


def parse_rows(table_xml: str) -> list:
    """Return logical rows (list of 16 cell strings) for content rows."""
    rows = []
    for row_match in re.finditer(r"<table:table-row\b.*?</table:table-row>", table_xml, re.S):
        row = row_match.group(0)
        repeat_match = re.search(r'table:number-rows-repeated="(\d+)"', row)
        repeat = int(repeat_match.group(1)) if repeat_match else 1

        cells = []
        for cell_match in re.finditer(
            r"<table:(?:covered-)?table-cell\b[^>]*?(?:/>|>.*?</table:(?:covered-)?table-cell>)",
            row,
            re.S,
        ):
            tag = cell_match.group(0)
            col_repeat_match = re.search(r'table:number-columns-repeated="(\d+)"', tag)
            col_repeat = int(col_repeat_match.group(1)) if col_repeat_match else 1
            text = xml_text(tag)
            cells.extend([text] * min(col_repeat, 16))
            if len(cells) >= 16:
                break

        if len(cells) < 16:
            cells.extend([""] * (16 - len(cells)))
        cells = cells[:16]

        if not any(cells):
            # Empty row (trailing repeats can be huge): skip.
            continue
        for _ in range(min(repeat, 50)):
            rows.append(list(cells))
    return rows


def extract_data(content: str) -> dict:
    rows = parse_rows(get_table(content, "data"))
    # Drop the header row (first non-empty row) when present.
    if rows and rows[0][0].strip().lower() == "activity":
        rows = rows[1:]

    activities = []
    states = []
    actstates = []
    tasks = []

    for r in rows:
        if r[0] or r[1]:
            tasks.append((r[0], r[1], r[2], r[3], r[4], r[5]))
        if r[7]:
            activities.append((r[7], r[8]))
        if r[9]:
            states.append((r[9], r[10], r[11]))
        if r[12]:
            actstates.append((r[12], r[13], r[14], r[15]))

    # Infer roles when the old file has no StateRole values.
    has_roles = any(s[1] for s in states)
    if not has_roles:
        states = [
            (name, LEGACY_ROLES.get(name.strip().lower(), ""), scope)
            for name, _role, scope in states
        ]

    # "active" activity state is required by the current model.
    if not any(a[0].strip().lower() == "active" for a in actstates):
        actstates.append(("active", "none", "", ""))

    return {
        "activities": activities,
        "states": states,
        "actstates": actstates,
        "tasks": tasks,
    }


def extract_view(content: str) -> tuple:
    view = get_table(content, "view")
    rows = parse_rows(view)
    a2 = rows[1][0] if len(rows) > 1 else ""
    a3 = rows[2][0] if len(rows) > 2 else ""
    b3 = rows[2][1] if len(rows) > 2 else ""
    return a2, a3, b3


# ---------------------------------------------------------------------------
# ODS writing
# ---------------------------------------------------------------------------
def build_data_table(prefix: str, data: dict) -> str:
    n = max(
        len(data["activities"]),
        len(data["states"]),
        len(data["actstates"]),
        len(data["tasks"]),
    )

    out = [prefix]

    # Header row.
    header = ["<table:table-row>"]
    for value in HEADERS:
        header.append(cell_xml(value))
    header.append(
        '<table:table-cell table:number-columns-repeated="%d"/>' % (16384 - len(HEADERS))
    )
    header.append("</table:table-row>")
    out.append("".join(header))

    for i in range(n):
        cells = []
        if i < len(data["tasks"]):
            a, b, c, d, e, f = data["tasks"][i]
            cells.extend([cell_xml(a), cell_xml(b), cell_xml(c), cell_xml(d), number_xml(e), number_xml(f)])
        else:
            cells.extend(["<table:table-cell/>"] * 6)
        cells.append("<table:table-cell/>")  # G
        if i < len(data["activities"]):
            name, state = data["activities"][i]
            cells.extend([cell_xml(name), cell_xml(state)])
        else:
            cells.extend(["<table:table-cell/>", "<table:table-cell/>"])
        if i < len(data["states"]):
            name, role, scope = data["states"][i]
            cells.extend([cell_xml(name), cell_xml(role), cell_xml(scope)])
        else:
            cells.extend(["<table:table-cell/>"] * 3)
        if i < len(data["actstates"]):
            name, kind, param, message = data["actstates"][i]
            cells.extend([cell_xml(name), cell_xml(kind), cell_xml(param), cell_xml(message)])
        else:
            cells.extend(["<table:table-cell/>"] * 4)
        cells.append(
            '<table:table-cell table:number-columns-repeated="%d"/>' % (16384 - len(cells))
        )
        out.append("<table:table-row>%s</table:table-row>" % "".join(cells))

    remaining = 1048576 - (1 + n)
    if remaining > 0:
        out.append(
            '<table:table-row table:number-rows-repeated="%d">'
            '<table:table-cell table:number-columns-repeated="16384"/>'
            "</table:table-row>" % remaining
        )

    return "".join(out)


def replace_data_table(content: str, data: dict) -> str:
    match = re.search(r'(<table:table\b[^>]*table:name="data".*?</table:table>)', content, re.S)
    if not match:
        raise RuntimeError("data table not found")
    table = match.group(1)
    first_row = re.search(r"<table:table-row\b", table)
    prefix = table[: first_row.start()] if first_row else table.replace("</table:table>", "")
    new_table = build_data_table(prefix, data) + "</table:table>"
    return content[: match.start(1)] + new_table + content[match.end(1):]


def set_validation_cell(content: str, valname: str, value: str) -> str:
    # Self-closing cell (for example an empty A2).
    self_pat = re.compile(
        r'<table:table-cell\b[^>]*table:content-validation-name="%s"[^>]*/>' % re.escape(valname)
    )
    match = self_pat.search(content)
    if match:
        if value == "":
            return content
        tag = match.group(0)[:-2].rstrip()
        if "office:value-type" not in tag:
            tag += ' office:value-type="string" calcext:value-type="string"'
        new = tag + "><text:p>%s</text:p></table:table-cell>" % html.escape(value, quote=False)
        return content[: match.start()] + new + content[match.end():]

    full_pat = re.compile(
        r'(<table:table-cell\b[^>]*table:content-validation-name="%s"[^>]*>)(.*?)(</table:table-cell>)'
        % re.escape(valname),
        re.S,
    )
    match = full_pat.search(content)
    if match:
        inner = "<text:p>%s</text:p>" % html.escape(value, quote=False) if value != "" else "<text:p/>"
        new = match.group(1) + inner + match.group(3)
        return content[: match.start()] + new + content[match.end():]
    return content


def write_result(data: dict, view: tuple, dest_path: str) -> None:
    if not os.path.isfile(BLANK_TEMPLATE):
        raise RuntimeError("template not found: " + BLANK_TEMPLATE)

    with zipfile.ZipFile(BLANK_TEMPLATE, "r") as zin, zipfile.ZipFile(dest_path, "w") as zout:
        for item in zin.infolist():
            payload = zin.read(item.filename)
            if item.filename == "content.xml":
                content = payload.decode("utf-8")
                content = replace_data_table(content, data)
                a2, a3, b3 = view
                content = set_validation_cell(content, "val1", a2)
                content = set_validation_cell(content, "val2", a3)
                content = set_validation_cell(content, "val3", b3)
                payload = content.encode("utf-8")
            zout.writestr(item, payload)


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------
def prompt(text: str, default: str = "") -> str:
    suffix = " [%s]" % default if default else ""
    answer = input(text + suffix + ": ").strip()
    return answer or default


def main() -> int:
    parser = argparse.ArgumentParser(description="Migrate an old Canban ODS to the current version.")
    parser.add_argument("--source", "-s", default=None, help="old ODS file")
    parser.add_argument("--dest", "-d", default=None, help="destination directory")
    parser.add_argument("--name", "-n", default=None, help="resulting file name")
    args = parser.parse_args()

    source = args.source
    if not source:
        source = prompt("Path to the old ODS file")
    source = os.path.abspath(os.path.expanduser(source))
    if not os.path.isfile(source):
        print("ERROR: source not found:", source)
        return 2

    dest_dir = args.dest
    if not dest_dir:
        dest_dir = prompt("Destination directory", DEFAULT_DEST)
    dest_dir = os.path.abspath(os.path.expanduser(dest_dir))

    name = args.name
    if not name:
        default_name = os.path.splitext(os.path.basename(source))[0] + "_migrated.ods"
        name = prompt("Result file name", default_name)
    if not name.lower().endswith(".ods"):
        name += ".ods"

    try:
        with zipfile.ZipFile(source, "r") as zin:
            content = zin.read("content.xml").decode("utf-8")
        data = extract_data(content)
        view = extract_view(content)
    except Exception as exc:  # noqa: BLE001
        print("ERROR: could not read the source file:", exc)
        return 3

    os.makedirs(dest_dir, exist_ok=True)
    dest_path = os.path.join(dest_dir, name)
    try:
        write_result(data, view, dest_path)
    except Exception as exc:  # noqa: BLE001
        print("ERROR: could not write the result:", exc)
        return 4

    print("Migration completed:", dest_path)
    print("  activities :", len(data["activities"]))
    print("  task states:", len(data["states"]))
    print("  activities states:", len(data["actstates"]))
    print("  tasks      :", len(data["tasks"]))
    print("  view       :", view[0] or "(empty)", "/", view[1] or "(empty)", "/", view[2] or "(empty)")
    return 0


if __name__ == "__main__":
    sys.exit(main())
