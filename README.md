# Canban

"Kanban Board" for LibreOffice Calc.

<img width="1020" height="561" alt="imagen" src="https://github.com/user-attachments/assets/dd68521a-ecbb-4d2f-86db-53252e5606f0" />


This project is a visual organization and task management tool in Kanban format, developed entirely within **LibreOffice Calc** using **Basic (StarBasic)** programming from the UNO API.

The system is designed under a separation of concerns, which is divided into an interactive visual board (Presentation Layer) and a structured flat record (Data Layer).

## Features
- **Fluid and Agile Workspace:** The Kanban board updates and interacts immediately when changing projects or status columns, offering an agile, lag-free user experience, free from the usual slowdowns of conventional spreadsheets.
- **Dynamic Detail Visualization:** By simply clicking on any task on the board, the side panel instantly displays its full title and detailed description, allowing the Kanban cards to be kept short, clean, and readable.
- **Intuitive Capture and Interaction:** The single task form visually adapts to the context of use. It offers simplified options for rapid creation and transforms its colors and buttons to clearly alert the user during sensitive operations such as sending to the trash or permanent deletion.
- **Security and Traceability of Your Tasks:** Each task is linked to a unique and persistent record in the internal database. This ensures that priorities can be reorganized, moved up, or moved down with complete certainty that texts and descriptions will never be mixed up or lost.
- **Separation of Operation and Configuration:** The system separates the daily organization routine from the workspace structuring. Through a dedicated configuration panel, it is possible to centrally add or remove projects and columns without interfering with quick task capture.
- **Activity States and Scoped States:** Activities can carry their own access-gating states (manual or date-based rules with a custom message), and task states can be general or scoped to a specific activity.
- **Local Approach and Total Privacy:** The tool is 100% self-contained and works completely offline. It requires no registrations, cloud accounts, or external dependencies; all your information remains privately secured on your computer.

This project aims to offer a local, lightweight, customizable, and privacy-oriented alternative for the agile management of straightforward projects, without relying on external web services or third-party databases.

## Getting Started (Installation)
1. Download one of the spreadsheet files from this repository:
   - `canban_blank.ods`: empty template for new users.
   - `canban_trial.ods`: sample data to explore the system (also useful for testing new features).
2. Open the file with LibreOffice Calc.
3. If a message appears indicating that macros are disabled, refer to the **Macro Security** section below.

If you already have data in a previous version, use the migration tool described in
**Migrating from a Previous Version** instead of starting from the template.

## Macro Security
For macros to work correctly on your computer:
1. In LibreOffice, go to **Tools > Options > LibreOffice > Security**.
2. Click on **Macro Security > Trusted Sources**.
3. Add the folder where you downloaded this file to the **Trusted Sources** list (recommended), or set the security level to **Medium**.

## Fundamental Commands
1. **Disable the menus** in LibreOffice Calc using the `Ctrl + Alt + Q` shortcut.
2. To **add a new task**, use the `Ctrl + Alt + N` shortcut.
3. To **edit an existing task**, use the `Ctrl + Alt + U` shortcut.
4. To **add/remove a status**, use the `Ctrl + Alt + E` shortcut.
5. To **add/remove an activity**, use the `Ctrl + Alt + A` shortcut.
6. To **send a task to the trash**, use the `Ctrl + Alt + P` shortcut.
7. To **delete (permanently erase) a task**, use the `Ctrl + Alt + D` shortcut.
8. To **manage activity states**, use the `Ctrl + Alt + S` shortcut.


## Distributions
- `canban_blank.ods`: template with only the essential base states and no tasks or activities.
- `canban_trial.ods`: sample data covering activity states, per-activity states and tasks. It is also the file to use when testing new features.

## Migrating from a Previous Version
Use the Python migration tool to move your data from an older ODS into the
current version:

```
python3 migrate/migrate_ods.py --source old.ods
```

Without arguments, the tool prompts for the source file, the destination folder
(default `migrate/migration_result/`) and the result file name. The migrated file
keeps your tasks, activities and states; base/trash roles are inferred from the
legacy names and can be adjusted later in the state manager.

## How to Contribute or Modify Macros
If you wish to make changes to the code:
1. The macro source code is located in the `src/macros.bas` file.
2. You can make modifications directly in the LibreOffice Calc Basic editor.
3. When finished, remember to export the modified module and replace the `src/macros.bas` file to keep the repository updated.

## License
This project is licensed under the MIT License - see the `LICENSE` file for details.
