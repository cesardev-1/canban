# Roadmap

This document records future features that are intentionally out of scope for
the current change set.

## Tags (labels)

Status: **future**. Change 3 is finalized in this release; the tags feature
described below is not part of it.

### Concept

Tags are a second level of organization, below states. The user-facing
hierarchy is:

```
Activity -> State -> Tag -> Task
```

- States are mandatory for every task.
- Tags are optional: a task may have zero or more tags.
- Tags are stored per activity (not per task), and can be:
  - **General tags**: available to tasks of any activity.
  - **Per-activity tags**: available only to tasks of a specific activity.

### Manual/administrative meaning

Tags group and filter tasks around concepts that are not workflow states
(for example `Essay`, `Thesis`, `Urgent`). They are independent from states,
but conceptually subordinate to them (states remain the primary organization).

### View changes

- `view.A4` and `view.B4` become the tag selectors for the two board columns.
- The task rendering range moves down one row (tasks start at row 5 instead of 4).
- An empty tag selector shows all tasks and the tasks without a tag.
- The hidden ID columns and the description panel shift accordingly.

### Management dialog

- A dedicated tag manager with the same corrected structure as the state
  manager:
  1. The container selector (General or an activity) is the first field.
  2. The item list shows only tags that belong to the selected container.
  3. Add / delete / modify flows follow the same atomic pattern used by
     `Manage States`.

### Data model (proposal)

- Add a tag catalog with its own columns (name, owning activity or empty for
  general tags).
- Store the assigned tags on the task record (for example, a delimited list
  or one column per tag slot).
- Keep tags optional and independent from `State` and `StateScope`.

### Open questions

- Maximum number of tags per task.
- Whether a task can carry tags from activities other than its own.
- How tags interact with sorting (state first, then tag, then order).
