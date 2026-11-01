"""Populate canban_draft.ods with atomic sample data for Change 3.

Run after build_draft_ods.py. It opens the draft, writes the sample
activity states, activities, scoped states and task assignments, then
stores the document in place.
"""
import os
import subprocess
import time

import uno
from com.sun.star.beans import PropertyValue

PORT = 2221
BASE = os.path.abspath("tmp/change1")
PROFILE = os.path.join(BASE, "lo_profile_seed")
ODS = os.path.abspath("canban_draft.ods")

# Activity states catalog: M=name, N=type, O=param, P=message.
ACTIVITY_STATES = [
    ("suspended", "manual", "",
     "This activity is suspended. Complete the pending requirement and confirm to resume."),
    ("scheduled", "date-from", "2099-01-01",
     "This activity is scheduled. It will be available from 2099-01-01."),
    ("expired", "date-until", "2000-01-01",
     "This activity is closed and can no longer be accessed."),
    ("review", "manual", "",
     "This activity is under review. Confirm the review to resume."),
    ("active", "none", "", ""),
]

# Activities: H=name, I=current activity state.
ACTIVITIES = [
    (1, "Personal", "suspended"),
    (2, "School", "scheduled"),
    (3, "Archive", "expired"),
    (4, "Ideas", "active"),
]

# Task states: J=name, K=role, L=scope (owner activity, empty = general).
SCOPED_STATES = [
    (4, "Essay", "", "School"),
    (5, "Thesis", "base", "School"),
    (6, "Errands", "", "Personal"),
]

# Task assignments: row, state.
TASK_STATES = [
    (3, "Errands"),
    (5, "Thesis"),
    (6, "Essay"),
]


def main() -> None:
    os.makedirs(PROFILE, exist_ok=True)
    try:
        os.remove(os.path.join(BASE, ".~lock.canban_draft.ods#"))
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

        # 1. Activity states catalog.
        for index, (name, kind, param, message) in enumerate(ACTIVITY_STATES, start=1):
            data.getCellByPosition(12, index).String = name
            data.getCellByPosition(13, index).String = kind
            data.getCellByPosition(14, index).String = param
            data.getCellByPosition(15, index).String = message

        # 2. Activities and their current state.
        for row, name, state in ACTIVITIES:
            data.getCellByPosition(7, row).String = name
            data.getCellByPosition(8, row).String = state

        # 3. Scoped task states.
        for row, name, role, scope in SCOPED_STATES:
            data.getCellByPosition(9, row).String = name
            data.getCellByPosition(10, row).String = role
            data.getCellByPosition(11, row).String = scope

        # 4. Task assignments to scoped states.
        for row, state in TASK_STATES:
            data.getCellByPosition(2, row).String = state

        # 5. Show one general and one scoped state on the board.
        view.getCellByPosition(0, 2).String = "In progress"
        view.getCellByPosition(1, 2).String = "Errands"

        doc.store()
        doc.close(True)
        print("seeded", ODS)
    finally:
        proc.terminate()


if __name__ == "__main__":
    main()
