//! Local-origin undo management.
//!
//! WASM note: yrs gates `SystemClock`, `Options::default()`, and `UndoManager::new` off
//! `wasm32-unknown-unknown`, so this module builds its `Options` explicitly around an injectable
//! [`Clock`]. Hosts with a real clock (JS `Date.now`) should use
//! [`EditingDoc::undo_scope_with_clock`]; the plain [`EditingDoc::undo_scope`] picks a per-target
//! default.

use std::cell::RefCell;
use std::collections::HashSet;
use std::sync::Arc;

use yrs::sync::time::Clock;
use yrs::{Origin, ReadTxn, Transact};

use crate::op::OpResult;
use crate::{EditingDoc, STORIES, story_ref};

/// Undo capture window.
pub const UNDO_CAPTURE_TIMEOUT_MS: u64 = 500;

/// Target undo depth; yrs exposes no stack-trim API.
pub const UNDO_DEPTH: usize = 100;

/// The contract-shaped undo surface over yrs [`yrs::undo::UndoManager`].
pub struct DocUndoManager {
    inner: yrs::undo::UndoManager<()>,
}

/// The default clock for [`EditingDoc::undo_scope`].
///
/// - Native targets read the system clock, matching yrs's own default.
/// - `wasm32-unknown-unknown` has no ambient clock; the fallback advances a monotonic counter by
///   a full capture window per reading, so every transaction forms its own undo group (safe but
///   ungrouped). Hosts that want real 500 ms grouping inject `Date.now` through
///   [`EditingDoc::undo_scope_with_clock`].
pub(crate) fn default_clock() -> Arc<dyn Clock> {
    #[cfg(not(all(target_family = "wasm", target_os = "unknown")))]
    {
        Arc::new(|| {
            std::time::SystemTime::now()
                .duration_since(std::time::UNIX_EPOCH)
                .map(|elapsed| elapsed.as_millis() as u64)
                .unwrap_or_default()
        })
    }
    #[cfg(all(target_family = "wasm", target_os = "unknown"))]
    {
        use std::sync::atomic::{AtomicU64, Ordering};
        let ticks = AtomicU64::new(0);
        Arc::new(move || ticks.fetch_add(UNDO_CAPTURE_TIMEOUT_MS + 1, Ordering::Relaxed))
    }
}

impl EditingDoc {
    /// Builds an undo manager scoped to the given stories with the contract policy:
    /// tracked-origins = local only, capture timeout 500 ms. `Origin::System` (and agent/remote)
    /// transactions use string origins and are therefore untracked.
    pub fn undo_scope(&self, story_ids: &[&str]) -> OpResult<DocUndoManager> {
        self.undo_scope_with_clock(story_ids, default_clock())
    }

    /// [`EditingDoc::undo_scope`] with a host-injected clock (milliseconds since an arbitrary
    /// epoch) — the WASM-safe constructor, mirroring how `EditCtx.now_iso` keeps the clock a
    /// host concern.
    pub fn undo_scope_with_clock(
        &self,
        story_ids: &[&str],
        clock: Arc<dyn Clock>,
    ) -> OpResult<DocUndoManager> {
        let mut stories = Vec::with_capacity(story_ids.len());
        {
            let txn = self.yrs_doc().transact();
            for story_id in story_ids {
                stories.push(story_ref(&txn, story_id)?);
            }
        }
        Ok(DocUndoManager {
            inner: build_manager(self, &stories, clock),
        })
    }
}

/// The one WASM-safe constructor every undo path routes through: explicit `Options` (no
/// `Options::default()` / `UndoManager::new`, both unavailable on `wasm32-unknown-unknown`).
pub(crate) fn build_manager(
    doc: &EditingDoc,
    stories: &[yrs::TextRef],
    clock: Arc<dyn Clock>,
) -> yrs::undo::UndoManager<()> {
    let options = yrs::undo::Options {
        capture_timeout_millis: UNDO_CAPTURE_TIMEOUT_MS,
        tracked_origins: HashSet::from([Origin::from(doc.client_id())]),
        capture_transaction: None,
        timestamp: clock,
        init_undo_stack: Vec::new(),
        init_redo_stack: Vec::new(),
    };
    let mut manager = yrs::undo::UndoManager::with_options(options);
    for story in stories {
        manager.expand_scope(doc.yrs_doc(), story);
    }
    manager
}

impl DocUndoManager {
    pub fn undo(&mut self) -> bool {
        self.inner.undo_blocking()
    }

    pub fn redo(&mut self) -> bool {
        self.inner.redo_blocking()
    }

    pub fn can_undo(&self) -> bool {
        self.inner.can_undo()
    }

    pub fn can_redo(&self) -> bool {
        self.inner.can_redo()
    }

    /// Closes the current capture group.
    pub fn add_undo_barrier(&mut self) {
        self.inner.reset();
    }

    /// Current undo stack size (see [`UNDO_DEPTH`]).
    pub fn undo_depth(&self) -> usize {
        self.inner.undo_stack().len()
    }

    /// Current redo stack size — the redo twin of [`DocUndoManager::undo_depth`].
    pub fn redo_depth(&self) -> usize {
        self.inner.redo_stack().len()
    }

    /// Clears both stacks (file-load reset).
    pub fn clear(&mut self) {
        self.inner.clear_all();
    }

    /// Low-level escape hatch for the transport/awareness bridges.
    pub fn raw(&mut self) -> &mut yrs::undo::UndoManager<()> {
        &mut self.inner
    }
}

/// One tracked undo scope per session, shared by the WASM boundary and the native facade.
///
/// Every accessor is inert until a scope is tracked, so a host may call it before import.
#[derive(Default)]
pub struct UndoSession {
    manager: RefCell<Option<DocUndoManager>>,
    scope: RefCell<Option<String>>,
}

impl UndoSession {
    pub fn new() -> Self {
        Self::default()
    }

    /// Tracks one story, replacing any other scope and its history. Re-tracking is a no-op.
    pub fn track(&self, doc: &EditingDoc, story: &str) -> OpResult<()> {
        if self.scope.borrow().as_deref() == Some(story) {
            return Ok(());
        }
        let manager = doc.undo_scope(&[story])?;
        *self.manager.borrow_mut() = Some(manager);
        *self.scope.borrow_mut() = Some(story.to_owned());
        Ok(())
    }

    /// [`UndoSession::track`] widened to the stories root so structural table edits also
    /// restore the cell stories they created or destroyed. Tracked as its own scope.
    pub fn track_table(&self, doc: &EditingDoc, story: &str) -> OpResult<()> {
        let scope = format!("table:{story}");
        if self.scope.borrow().as_deref() == Some(scope.as_str()) {
            return Ok(());
        }
        let mut manager = doc.undo_scope(&[story])?;
        let txn = doc.yrs_doc().transact();
        let stories = txn
            .get_map(STORIES)
            .expect("stories root is declared by EditingDoc::new");
        drop(txn);
        manager.raw().expand_scope(doc.yrs_doc(), &stories);
        *self.manager.borrow_mut() = Some(manager);
        *self.scope.borrow_mut() = Some(scope);
        Ok(())
    }

    pub fn undo(&self) -> bool {
        self.manager
            .borrow_mut()
            .as_mut()
            .is_some_and(DocUndoManager::undo)
    }

    pub fn redo(&self) -> bool {
        self.manager
            .borrow_mut()
            .as_mut()
            .is_some_and(DocUndoManager::redo)
    }

    pub fn can_undo(&self) -> bool {
        self.manager
            .borrow()
            .as_ref()
            .is_some_and(DocUndoManager::can_undo)
    }

    pub fn can_redo(&self) -> bool {
        self.manager
            .borrow()
            .as_ref()
            .is_some_and(DocUndoManager::can_redo)
    }

    pub fn add_undo_barrier(&self) {
        if let Some(manager) = self.manager.borrow_mut().as_mut() {
            manager.add_undo_barrier();
        }
    }

    pub fn undo_depth(&self) -> usize {
        self.manager
            .borrow()
            .as_ref()
            .map_or(0, DocUndoManager::undo_depth)
    }

    pub fn redo_depth(&self) -> usize {
        self.manager
            .borrow()
            .as_ref()
            .map_or(0, DocUndoManager::redo_depth)
    }
}
