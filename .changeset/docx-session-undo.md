---
"@betteroffice/docx": minor
"@betteroffice/docx-react": minor
"@betteroffice/rust-crates": patch
---

Keep one local undo history across document stories, group rapid keystrokes in WebAssembly, and preserve native undo in other inputs. Replace story-scoped history helpers with session-wide tracking and changed-story reporting.

Migrate each removed API as follows: `historyStory()` returns the changed stories via `historyStories()` (sorted, empty before the first local edit instead of `null`); `undoDepth()` and `redoDepth()` are gone, query `canUndo()` and `canRedo()` instead; `markUndoGroup(startDepth)` is gone, rapid keystrokes now coalesce in WebAssembly with no host bookkeeping; `applyLocalUpdate(update, story)` drops its story argument and becomes `applyLocalUpdate(update)`; `beginUndoCapture(story, includeTableStories?)` drops its arguments and becomes `beginUndoCapture()`; `computeLayout()` no longer returns `blocks` and `measures`, read them lazily from `getLayoutKernelInputs(computation.layout)` as `measured` and `options`.
