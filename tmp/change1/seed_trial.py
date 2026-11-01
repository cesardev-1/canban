"""Populate a variant ODS with the Change 3 sample data.

Usage: python3 seed_trial.py <path.ods>

The sample covers every activity-state condition type, general and scoped
task states, and per-activity base states. It finishes by leaving the
blocked "Personal" activity selected so the startup access service can be
observed on the next open.
"""
import os
import subprocess
import sys
import time

import uno
from com.sun.star.beans import PropertyValue

ODS = os.path.abspath(sys.argv[1] if len(sys.argv) > 1 else "canban_trial.ods")
BASE = os.path.abspath("tmp/change1")
_NAME = os.path.splitext(os.path.basename(ODS))[0]
PORT = 2300 + (sum(ord(c) for c in _NAME) % 200)
PROFILE = os.path.join(BASE, "lo_profile_" + _NAME)

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
    (4, "Work", ""),
    (5, "Ideas", ""),
]

# Task states: J=name, K=role, L=scope (owner activity, empty = general).
SCOPED_STATES = [
    (4, "Urgent", "", "Personal"),
    (5, "Errands", "", "Personal"),
    (6, "Homework", "base", "School"),
    (7, "Study", "", "School"),
    (8, "Meeting", "", "Work"),
    (9, "Report", "base", "Work"),
]

# Existing task row -> state.
TASK_STATES = [
    (2, "Urgent"),
    (3, "Errands"),
    (4, "Homework"),
    (5, "Study"),
    (6, "Homework"),
]

# New tasks: row, activity, title, state, description, order, id.
NEW_TASKS = [
    (7, "Work", "Prepare quarterly report", "Report",
     "Draft and review the quarterly numbers", 7, 7),
    (8, "Work", "Team meeting", "Meeting",
     "Weekly sync with the team", 8, 8),
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

        # 4. Existing task state assignments.
        for row, state in TASK_STATES:
            data.getCellByPosition(2, row).String = state

        # 5. New tasks for the Work activity.
        for row, activity, title, state, description, order, task_id in NEW_TASKS:
            data.getCellByPosition(0, row).String = activity
            data.getCellByPosition(1, row).String = title
            data.getCellByPosition(2, row).String = state
            data.getCellByPosition(3, row).String = description
            data.getCellByPosition(4, row).Value = order
            data.getCellByPosition(5, row).Value = task_id

        # 6. Show the accessible Work states on the board.
        view.getCellByPosition(0, 2).String = "Report"
        view.getCellByPosition(1, 2).String = "Meeting"

        # 7. Leave the blocked activity selected so the startup service can
        #    switch to an accessible one on the next open.
        provider = doc.getScriptProvider()
        script = provider.getScript(
            "vnd.sun.star.script:Standard.Module1.SetActivityFilterSilently?language=Basic&location=document")
        script.invoke(("Personal",), (), ())

        doc.store()
        doc.close(True)
        print("seeded", ODS)
    finally:
        proc.terminate()
        try:
            proc.wait(timeout=15)
        except Exception:
            proc.kill()


if __name__ == "__main__":
    main()
