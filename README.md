# Canban

"Kanban Board" for LibreOffice Calc.

This project is a visual organization and task management tool in Kanban format, developed entirely within **LibreOffice Calc** using **Basic (StarBasic)** programming from the UNO API.

The system is designed under a separation of concerns, which is divided into an interactive visual board (Presentation Layer) and a structured flat record (Data Layer).

## Features
- **Fluid and Agile Workspace:** The Kanban board updates and interacts immediately when changing projects or status columns, offering an agile, lag-free user experience, free from the usual slowdowns of conventional spreadsheets.
- **Dynamic Detail Visualization:** By simply clicking on any task on the board, the side panel instantly displays its full title and detailed description, allowing the Kanban cards to be kept short, clean, and readable.
- **Intuitive Capture and Interaction:** The single task form visually adapts to the context of use. It offers simplified options for rapid creation and transforms its colors and buttons to clearly alert the user during sensitive operations such as sending to the trash or permanent deletion.
- **Security and Traceability of Your Tasks:** Each task is linked to a unique and persistent record in the internal database. This ensures that priorities can be reorganized, moved up, or moved down with complete certainty that texts and descriptions will never be mixed up or lost.
- **Separation of Operation and Configuration:** The system separates the daily organization routine from the workspace structuring. Through a dedicated configuration panel, it is possible to centrally add or remove projects and columns without interfering with quick task capture.
- **Local Approach and Total Privacy:** The tool is 100% self-contained and works completely offline. It requires no registrations, cloud accounts, or external dependencies; all your information remains privately secured on your computer.

This project aims to offer a local, lightweight, customizable, and privacy-oriented alternative for the agile management of straightforward projects, without relying on external web services or third-party databases.

## Getting Started (Installation)
1. Download the `canban.ods` file from this repository.
2. Open the file with LibreOffice Calc.
3. If a message appears indicating that macros are disabled, refer to the **Macro Security** section below.

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


## How to Contribute or Modify Macros
If you wish to make changes to the code:
1. The macro source code is located in the `src/macros.bas` file.
2. You can make modifications directly in the LibreOffice Calc Basic editor.
3. When finished, remember to export the modified module and replace the `src/macros.bas` file to keep the repository updated.

## License
This project is licensed under the MIT License - see the `LICENSE` file for details.
