use super::{EditSessionCore, EngineError, EngineResult};
use crate::UndoCaptureMode;

impl EditSessionCore {
    /// Starts local-origin undo tracking across every story. Call this after
    /// import or seeding but before the first edit, so the initial document is
    /// not an undo step; later calls keep the history.
    pub fn track_undo(&self) {
        self.undo.track(self.engine.doc());
    }

    /// Closes the current undo capture without adding an empty step.
    pub fn add_undo_boundary(&self) {
        self.undo.add_undo_barrier();
    }

    /// Changes grouping policy while retaining undo and redo history.
    pub fn set_undo_capture_mode(&self, mode: &str) -> EngineResult<()> {
        let mode = match mode {
            "auto" => UndoCaptureMode::Auto,
            "manual" => UndoCaptureMode::Manual,
            _ => {
                return Err(EngineError::message(
                    "undo capture mode must be auto or manual",
                ));
            }
        };
        self.undo.set_capture_mode(mode);
        Ok(())
    }

    /// Pins the input span used to measure the next capture gap.
    pub fn set_undo_input_time(&self, first: f64, last: f64) {
        self.undo.set_input_time(Some((first as u64, last as u64)));
    }

    /// Clears the pinned input span.
    pub fn clear_undo_input_time(&self) {
        self.undo.set_input_time(None);
    }

    /// Current undo grouping policy.
    pub fn undo_capture_mode(&self) -> String {
        match self.undo.capture_mode() {
            UndoCaptureMode::Auto => "auto",
            UndoCaptureMode::Manual => "manual",
        }
        .to_owned()
    }

    /// Selects a story, closing capture unless manual grouping is selected.
    pub fn select_story(&self, story: &str) {
        self.undo.select_story(story);
    }

    /// Reverts the latest local-origin step and reports whether anything was
    /// reverted. Remote and system transactions are excluded by the manager's
    /// tracked-origin policy; `false` before tracking starts.
    pub fn undo(&self) -> bool {
        self.undo.undo()
    }

    /// Reapplies the latest locally undone transaction and reports whether
    /// anything was reapplied.
    pub fn redo(&self) -> bool {
        self.undo.redo()
    }

    /// Whether [`EditSessionCore::undo`] would revert something.
    pub fn can_undo(&self) -> bool {
        self.undo.can_undo()
    }

    /// Whether [`EditSessionCore::redo`] would reapply something.
    pub fn can_redo(&self) -> bool {
        self.undo.can_redo()
    }

    /// Stories changed by the latest undo or redo, sorted.
    pub fn history_stories(&self) -> Vec<String> {
        self.undo.changed_stories()
    }
}
