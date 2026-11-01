# Original Plan — Canban (State/Activity Management Branch)

> Historical document. This is the original plan and architecture analysis produced
> at the start of the branch, before implementation. It is kept as a reference for
> comparing the original proposal with the final implementation.

## 1. Initial Request

The branch is dedicated to improving the management experience for states and
activities. All natural-language content in the project must be in English.
Spanish leftovers are considered legacy.

Requested improvements:

- [x] **State/activity management menus**
   - `Ctrl + Alt + E` (states) and `Ctrl + Alt + A` (activities) must allow:
     moving, adding, deleting, modifying, and (for states) marking items as base.
   - Establish an order, either from adjacent cells next to the catalog columns
     in the `data` sheet, or derived from the item position in the matrix.
     The matrix-position approach avoids adding unnecessary content, but its
     consequences on changes must be analyzed.
   - For the `states` column, implement a way to mark states as `base`. The code
     currently hardcodes state names (for example, the state where tasks are sent
     to the trash). These minimum required states must be marked `base` in the
     matrix instead of declared in code. This allows users to modify the value
     stored in the base-state cells and other states.

- [x] **Consequences of modifying states and activities**
   - When states/activities are modified, reorganize the columns and remove blank
     cells (currently gaps are left after deleting items).
   - When activity/state names change, reassign everything to the new name.

3. **New state/activity data structures**
   - Activities can also have states (states assignable only to activities, not
     to tasks), changing the behavior/content of the `A3` and `B3` dropdowns in
     the `view` sheet.
   - Activity states can trigger behaviors. Example: a `suspended` activity state
     opens a popup warning that a condition must be completed before returning to
     `active`, and the popup allows marking the condition as fulfilled or not.
     The feature must be general and user-defined from the matrix (text and
     condition), not hardcoded.
   - Define general states, activity states, and per-activity states.
     - Per-activity states are states that only the tasks of one activity can
       have (not states that activities themselves have).
     - General states are shared by all tasks (current behavior).
   - Proposal: create two whole custom matrices for state and activity
     information (similar to the task matrix), so they can be managed and
     extended with properties in an orderly way.

## 2. Extracted Architecture

### Layers
- **Presentation layer:** `view` sheet (interactive Kanban board).
- **Data layer:** `data` sheet (flat record).
- **Logic:** `Module1` Basic/StarBasic macros. Everything is macro-driven; no
  formulas are written into the spreadsheet matrix.

### `data` sheet
| Col | Index | Role |
|-----|-------|------|
| A | 0 | Activity |
| B | 1 | Task title |
| C | 2 | State |
| D | 3 | Description |
| E | 4 | Order (decimal weight) |
| F | 5 | ID_Task (immutable) |
| G | 6 | empty |
| H | 7 | **Activities** catalog (`H2:H100`) |
| I | 8 | empty |
| J | 9 | **States** catalog (`J2:J20`) |

- Row 0 = headers; data starts at row index 1 (sheet row 2).
- The `Trash` state is at the end of `J` and acts as the system trash state.
- Catalog cells share physical rows with tasks.

### `view` sheet
- `A2` (merged A2:D2) -> activity filter (dropdown `val1` -> `data.H2:H1048576`).
- `A3` -> left state column (dropdown `val2` -> `data.J2:J18`).
- `B3` -> right state column (dropdown `val3` -> `data.J2:J18`).
- Columns `A`/`B`, rows index 3..299 -> task cards.
- `D3` title and `D4` description (injected).
- `E`/`F` hidden IDs for columns A/B.
- `E1` debug cell with the focused ID.

### Engine components
- **Render:** `DrawingBoard()` clears rows 3..299, reads filters A2/A3/B3,
  traverses `data`, and paints tasks into columns A/B based on state, storing
  hidden IDs in E/F.
- **Sort:** `CustomDataSort()` heals IDs/orders, sorts by (Activity `H` ->
  State `J` -> Order `E`) and renumbers sequentially, filtering ghost rows.
- **Priority lookup:** `GetListIndex()` returns the physical list position in
  H/J as the priority; 999 when not found.
- **Cascade delete:** `RunCascadeDelete()` validates hardcoded base states,
  counts and deletes tasks plus the catalog entry.
- **IDs:** `GetNextAvailableId()` / `AssignMissingIdsBD()` reuse gaps in `F`.
- **Order/position:** `GetNextOrder()` / `AssignMissingOrdersBD()` use decimal
  interpolation in `E`.
- **Description:** `RunInjectionDescription()` injects title/description by ID.
- **Listeners:** `StartListenerSelection`, `VisorSeleccion_selectionChanged`,
  `Visor_onContentChanged`.
- **Read hub:** `InitializeDialogServices()` and services
  `ServPopulateCmbActivity/State`, `ServManageTxtTitle/Description`,
  `ServCalculateNumPosition`, `ServPopulateCmbItem`.
- **Write hub:** `SaveDataFromDialog()` and services
  `ServSaveCmbActivity/State/TxtTitle/TxtDescription/NumPosition(+Default)`.
- **Summoners:** `ManageActivityes()`, `ManageStates()` -> `OpenItemPanel()` ->
  `DialogPanelItem`.
- **Catalog:** `SaveCatalogFromDialog()` + `CmbItemTextModified()`.
- **Tasks:** `AddNewTask`, `EditActiveTask`, `MoveTaskToTrash`,
  `DeleteActiveTask` -> `DialogPanelTask`.

### Hardcoded values at the time
- `"Trash"`: `MoveTaskToTrash`, exclusion in `ServPopulateCmbState`, rule in
  `ServCalculateNumPosition`.
- `"In progress"` and `"To do"`: base-state validation in `RunCascadeDelete` and
  defaults in `AddNewTask` / `ServCalculateNumPosition`.
- `"General"`: default activity in `ServCalculateNumPosition`.

### Findings relevant to the changes
1. Legacy Spanish references in `DialogPanelTask.xml`
   (`EliminarTareaActiva`, `MoverTareaAPapelera`, `CapturarCtrlEnter`) did not
   match the English `Module1`.
2. Residual Spanish in `macros.bas` comments/identifiers, plus `README_ES.md`
   and `macros_es_legacy.bas`.
3. Range mismatch: code reads states from `J2:J20`, validations use `J2:J18`;
   activities in code `H2:H100`, validations `H2:H29` / `H2:H1048576`.
4. Deleting catalog items leaves blank cells in H/J.

## 3. Original Roadmap

The architecture is clear and stable: catalogs in H/J, range-based validations,
hardcoded base states, and one unified dialog per item type. The natural entry
point for Change 1 is `DialogPanelItem` + `SaveCatalogFromDialog` +
`RunCascadeDelete`. The key decision is ordering by adjacent cells versus matrix
position.

Branch work, in order:

1. **Change 1 — State/activity management**
   - Reorder, add, delete, modify, and (states only) mark as `base`.
   - Order by matrix position in `H`/`J`.
   - Remove hardcoded base states (`In progress`, `To do`, `Trash`, `General`).

2. **Change 2 — Consequences**
   - Compact gaps in `H`/`J` after deletion.
   - Reassign everything when renaming states/activities (by ID or full sweep,
     to be analyzed).

3. **Change 3 — New structures**
   - General states, activity states (only for activities), and per-activity
     states (only for tasks of that activity).
   - User-defined conditional behaviors (for example, a `suspended` popup with a
     condition) driven by the matrix, not hardcoded.
   - New dropdown values/behavior in `A3`/`B3`.

## 4. Initial Decisions

- Expose the plan in chat first; validate before writing files.
- Order by matrix position (avoid extra "order" columns).
- Fix legacy references in `.ods`, `.bas`, and `README`; do **not** modify files
  marked as legacy (for example, `macros_es_legacy.bas`).
- Keep `canban.ods` untouched until the draft is reviewed and approved.

## 5. Notes

- This document intentionally reflects the **original** plan and does not include
  later refinements (position dropdown, role dropdown, `Modify` button with
  icons, etc.), which were introduced during implementation review.
