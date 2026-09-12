//! Local-origin undo management.
//!
//! One manager per session, scoped on the stories root, so every story — body, headers, footers,
//! notes and table cells — shares one ordered history. yrs gates `SystemClock`,
//! `Options::default()` and `UndoManager::new` off `wasm32-unknown-unknown`, so the manager is
//! built from explicit `Options` around an injectable [`Clock`]: native code reads the system
//! clock and the wasm host injects `Date.now`.

use std::cell::RefCell;
use std::collections::HashSet;
use std::sync::{Arc, Mutex};

use yrs::sync::time::Clock;
use yrs::{Map, Origin, Out, ReadTxn, Subscription, Transact};

use crate::{EditingDoc, STORIES};

/// Undo capture window.
pub const UNDO_CAPTURE_TIMEOUT_MS: u64 = 500;

/// Target undo depth; yrs exposes no stack-trim API.
pub const UNDO_DEPTH: usize = 100;

/// The contract-shaped undo surface over yrs [`yrs::undo::UndoManager`].
pub struct DocUndoManager {
    inner: yrs::undo::UndoManager<()>,
    changed_stories: Arc<Mutex<Vec<String>>>,
    _popped: Subscription,
}

/// System clock on native targets. `wasm32-unknown-unknown` has no ambient clock, so the fallback
/// advances a counter by a full capture window per reading: every transaction is its own step.
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
    /// Local-origin undo manager over every story, on the default clock.
    pub fn undo_manager(&self) -> DocUndoManager {
        DocUndoManager::new(self, default_clock())
    }
}

impl DocUndoManager {
    /// Tracks the local client id only, so agent, remote and system transactions (string
    /// origins) never enter the history; groups edits within [`UNDO_CAPTURE_TIMEOUT_MS`].
    fn new(doc: &EditingDoc, clock: Arc<dyn Clock>) -> Self {
        let options = yrs::undo::Options {
            capture_timeout_millis: UNDO_CAPTURE_TIMEOUT_MS,
            tracked_origins: HashSet::from([Origin::from(doc.client_id())]),
            capture_transaction: None,
            timestamp: clock,
            init_undo_stack: Vec::new(),
            init_redo_stack: Vec::new(),
        };
        let mut inner = yrs::undo::UndoManager::with_options(options);
        let root = stories_root(&doc.yrs_doc().transact());
        inner.expand_scope(doc.yrs_doc(), &root);
        let changed_stories = Arc::new(Mutex::new(Vec::new()));
        let popped = {
            let changed_stories = Arc::clone(&changed_stories);
            inner.observe_item_popped(move |txn, event| {
                let mut changed: Vec<String> = stories_root(txn)
                    .iter(txn)
                    .filter_map(|(story, value)| match value {
                        Out::YText(text) if event.has_changed(&text) => Some(story.to_owned()),
                        _ => None,
                    })
                    .collect();
                changed.sort();
                *lock(&changed_stories) = changed;
            })
        };
        Self {
            inner,
            changed_stories,
            _popped: popped,
        }
    }

    pub fn undo(&mut self) -> bool {
        lock(&self.changed_stories).clear();
        self.inner.undo_blocking()
    }

    pub fn redo(&mut self) -> bool {
        lock(&self.changed_stories).clear();
        self.inner.redo_blocking()
    }

    pub fn can_undo(&self) -> bool {
        self.inner.can_undo()
    }

    pub fn can_redo(&self) -> bool {
        self.inner.can_redo()
    }

    /// Stories changed by the latest undo or redo, sorted.
    pub fn changed_stories(&self) -> Vec<String> {
        lock(&self.changed_stories).clone()
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
}

fn stories_root(txn: &impl ReadTxn) -> yrs::MapRef {
    txn.get_map(STORIES)
        .expect("stories root is declared by EditingDoc::new")
}

fn lock(stories: &Mutex<Vec<String>>) -> std::sync::MutexGuard<'_, Vec<String>> {
    stories
        .lock()
        .unwrap_or_else(|poisoned| poisoned.into_inner())
}

/// The session's undo state, shared by the WASM boundary and the native facade.
///
/// Inert until [`UndoSession::track`] runs, so a host may import or seed first without the
/// initial document becoming an undo step.
pub struct UndoSession {
    clock: Arc<dyn Clock>,
    manager: RefCell<Option<DocUndoManager>>,
    story: RefCell<Option<String>>,
}

impl Default for UndoSession {
    fn default() -> Self {
        Self::with_clock(default_clock())
    }
}

impl UndoSession {
    pub fn new() -> Self {
        Self::default()
    }

    /// Groups edits within [`UNDO_CAPTURE_TIMEOUT_MS`] as read from `clock` (milliseconds).
    pub fn with_clock(clock: Arc<dyn Clock>) -> Self {
        Self {
            clock,
            manager: RefCell::new(None),
            story: RefCell::new(None),
        }
    }

    /// Starts tracking local edits in every story; later calls keep the history.
    pub fn track(&self, doc: &EditingDoc) {
        let mut manager = self.manager.borrow_mut();
        if manager.is_none() {
            *manager = Some(DocUndoManager::new(doc, Arc::clone(&self.clock)));
        }
    }

    /// Records the story holding the caret; moving to another story closes the capture group,
    /// so each undo step stays within one story.
    pub fn select_story(&self, story: &str) {
        if self.story.borrow().as_deref() == Some(story) {
            return;
        }
        *self.story.borrow_mut() = Some(story.to_owned());
        self.add_undo_barrier();
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

    /// Stories changed by the latest undo or redo, sorted.
    pub fn changed_stories(&self) -> Vec<String> {
        self.manager
            .borrow()
            .as_ref()
            .map(DocUndoManager::changed_stories)
            .unwrap_or_default()
    }

    pub fn add_undo_barrier(&self) {
        if let Some(manager) = self.manager.borrow_mut().as_mut() {
            manager.add_undo_barrier();
        }
    }
}

#[cfg(test)]
mod tests {
    use std::sync::atomic::{AtomicU64, Ordering};

    use super::*;
    use crate::{EditCtx, FormatPolicy, Position};

    const BODY: &str = "body";
    const HEADER: &str = "header:rId7";

    fn seed() -> EditingDoc {
        let doc = EditingDoc::new(100);
        doc.create_story(BODY, "body", "Normal", "left").unwrap();
        doc.create_story(HEADER, "header", "Header", "left")
            .unwrap();
        doc
    }

    fn append(doc: &EditingDoc, story: &str, text: &str) {
        let end = doc.story_len(story).unwrap() - 1;
        doc.insert_text(
            &EditCtx::local("", ""),
            Position::new(story, end),
            text,
            FormatPolicy::Plain,
        )
        .unwrap();
    }

    fn text(doc: &EditingDoc, story: &str) -> String {
        doc.paragraphs(story).unwrap()[0].text.clone()
    }

    /// A session whose clock only moves when the test advances `now`.
    fn stepped_session() -> (UndoSession, Arc<AtomicU64>) {
        let now = Arc::new(AtomicU64::new(1_000));
        let clock = Arc::clone(&now);
        let session = UndoSession::with_clock(Arc::new(move || clock.load(Ordering::Relaxed)));
        (session, now)
    }

    #[test]
    fn one_history_spans_every_story_in_order() {
        let doc = seed();
        let (undo, now) = stepped_session();
        undo.track(&doc);

        append(&doc, BODY, "!");
        now.fetch_add(UNDO_CAPTURE_TIMEOUT_MS + 1, Ordering::Relaxed);
        append(&doc, HEADER, "?");
        now.fetch_add(UNDO_CAPTURE_TIMEOUT_MS + 1, Ordering::Relaxed);
        undo.track(&doc);

        assert!(undo.undo());
        assert_eq!(text(&doc, HEADER), "header");
        assert_eq!(text(&doc, BODY), "body!");
        assert_eq!(undo.changed_stories(), [HEADER]);

        assert!(undo.undo());
        assert_eq!(text(&doc, BODY), "body");
        assert_eq!(undo.changed_stories(), [BODY]);
        assert!(!undo.undo());

        assert!(undo.redo());
        assert_eq!(text(&doc, BODY), "body!");
        assert_eq!(undo.changed_stories(), [BODY]);
    }

    #[test]
    fn keystrokes_inside_the_capture_window_form_one_step() {
        let doc = seed();
        let (undo, now) = stepped_session();
        undo.track(&doc);

        append(&doc, BODY, "a");
        now.fetch_add(100, Ordering::Relaxed);
        append(&doc, BODY, "b");
        assert!(undo.undo());
        assert_eq!(text(&doc, BODY), "body");

        now.fetch_add(UNDO_CAPTURE_TIMEOUT_MS + 1, Ordering::Relaxed);
        append(&doc, BODY, "c");
        now.fetch_add(600, Ordering::Relaxed);
        append(&doc, BODY, "d");
        assert!(undo.undo());
        assert_eq!(text(&doc, BODY), "bodyc");
        assert!(undo.undo());
        assert_eq!(text(&doc, BODY), "body");
    }

    #[test]
    fn moving_the_caret_to_another_story_closes_the_capture_group() {
        let doc = seed();
        let (undo, _now) = stepped_session();
        undo.track(&doc);

        undo.select_story(BODY);
        append(&doc, BODY, "!");
        undo.select_story(HEADER);
        append(&doc, HEADER, "?");

        assert!(undo.undo());
        assert_eq!(text(&doc, HEADER), "header");
        assert_eq!(text(&doc, BODY), "body!");
        assert_eq!(undo.changed_stories(), [HEADER]);
    }

    #[test]
    fn session_is_inert_until_tracked() {
        let doc = seed();
        let undo = UndoSession::new();
        append(&doc, BODY, "seed");
        assert!(!undo.can_undo());
        assert!(!undo.undo());

        undo.track(&doc);
        assert!(!undo.can_undo());
        append(&doc, BODY, "!");
        assert!(undo.can_undo());
    }
}
