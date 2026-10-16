use std::collections::HashSet;
#[cfg(any(feature = "wasm", test))]
use std::collections::{BTreeSet, HashMap};
use std::sync::{Arc, Mutex};

#[cfg(any(feature = "wasm", test))]
use serde::Serialize;
use yrs::sync::time::Clock;
#[cfg(any(feature = "wasm", test))]
use yrs::types::{Event, PathSegment};
#[cfg(any(feature = "wasm", test))]
use yrs::{DeepObservable, Map, MapRef, Subscription};
use yrs::{Doc, Origin, ReadTxn, Transact};

use crate::{COMMENTS, EditError, EditResult, META, SHAPES, SLIDE_ORDER, SLIDES, STORIES};

use crate::DeckSession;

const CAPTURE_TIMEOUT_MS: u64 = 500;

#[cfg(test)]
type HistorySets = Vec<(yrs::IdSet, yrs::IdSet)>;
#[cfg(test)]
type ClientClockCounts = Vec<(yrs::ClientID, u64)>;
#[cfg(test)]
type HistoryCounts = Vec<(ClientClockCounts, ClientClockCounts)>;

/// Policy for grouping tracked local transactions.
#[derive(Clone, Copy, Debug, Default, Eq, PartialEq)]
pub enum UndoCaptureMode {
    #[default]
    Auto,
    Manual,
}

struct CaptureClock {
    source: Arc<dyn Clock>,
    state: Mutex<CaptureClockState>,
}

struct CaptureClockState {
    mode: UndoCaptureMode,
    holds: usize,
    ticks: u64,
    last_source: u64,
}

impl CaptureClock {
    fn new(source: Arc<dyn Clock>) -> Self {
        Self {
            state: Mutex::new(CaptureClockState {
                mode: UndoCaptureMode::Auto,
                holds: 0,
                ticks: 1,
                last_source: source.now(),
            }),
            source,
        }
    }

    fn mode(&self) -> UndoCaptureMode {
        self.state.lock().unwrap().mode
    }

    fn set_mode(&self, mode: UndoCaptureMode) {
        self.state.lock().unwrap().mode = mode;
    }
}

impl Clock for CaptureClock {
    fn now(&self) -> u64 {
        let mut state = self.state.lock().unwrap();
        let now = self.source.now();
        let elapsed = match (state.mode, state.holds) {
            (UndoCaptureMode::Auto, 0) => now.saturating_sub(state.last_source),
            _ => 0,
        };
        state.last_source = now;
        state.ticks = state.ticks.saturating_add(elapsed);
        state.ticks
    }
}

struct CaptureClockHold(Arc<CaptureClock>);

impl Drop for CaptureClockHold {
    fn drop(&mut self) {
        self.0.state.lock().unwrap().holds -= 1;
    }
}

impl DeckSession {
    pub(crate) fn with_undo_capture<T>(
        &self,
        operation: impl FnOnce() -> EditResult<T>,
    ) -> EditResult<T> {
        self.automatic_undo_barrier();
        let clock = self.undo.borrow().clock.clone();
        clock.state.lock().unwrap().holds += 1;
        let hold = CaptureClockHold(clock);
        let result = operation();
        drop(hold);
        self.automatic_undo_barrier();
        result
    }
}

pub struct DeckUndoManager {
    clock: Arc<CaptureClock>,
    inner: yrs::undo::UndoManager<()>,
    #[cfg(any(feature = "wasm", test))]
    history: Option<HistoryObserver>,
}

#[cfg(any(feature = "wasm", test))]
#[derive(Default, Serialize)]
#[serde(rename_all = "camelCase")]
pub(crate) struct HistoryReceipt {
    pub applied: bool,
    pub changed_slide_ids: BTreeSet<String>,
    pub deck_size_changed: bool,
    pub slide_order_changed: bool,
}

#[cfg(any(feature = "wasm", test))]
#[derive(Default)]
struct HistoryState {
    owners: HashMap<String, String>,
    receipt: Option<HistoryReceipt>,
}

#[cfg(any(feature = "wasm", test))]
struct HistoryObserver {
    state: Arc<Mutex<HistoryState>>,
    _subscriptions: Vec<Subscription>,
}

#[cfg(any(feature = "wasm", test))]
fn index_shape<T: ReadTxn>(txn: &T, state: &mut HistoryState, shape_id: &str, slide_id: &str) {
    if state
        .owners
        .insert(shape_id.to_owned(), slide_id.to_owned())
        .as_deref()
        == Some(slide_id)
    {
        return;
    }
    refresh_shape(txn, state, shape_id, slide_id);
}

#[cfg(any(feature = "wasm", test))]
fn refresh_shape<T: ReadTxn>(txn: &T, state: &mut HistoryState, shape_id: &str, slide_id: &str) {
    let Some(shape) = txn
        .get_map(SHAPES)
        .and_then(|shapes| shapes.get(txn, shape_id))
        .and_then(|value| value.cast::<MapRef>().ok())
    else {
        return;
    };
    let Ok(stories) = crate::deck::map_string_array(&shape, txn, "textStories") else {
        return;
    };
    for story in stories {
        state.owners.insert(story, slide_id.to_owned());
    }
    let Ok(children) = crate::deck::map_string_array(&shape, txn, "children") else {
        return;
    };
    for child in children {
        index_shape(txn, state, &child, slide_id);
    }
}

#[cfg(any(feature = "wasm", test))]
fn index_slide<T: ReadTxn>(txn: &T, state: &mut HistoryState, slide_id: &str) {
    let Some(slide) = txn
        .get_map(SLIDES)
        .and_then(|slides| slides.get(txn, slide_id))
        .and_then(|value| value.cast::<MapRef>().ok())
    else {
        return;
    };
    if let Ok(order) = crate::deck::slide_shape_order(&slide, txn) {
        for id in crate::deck::string_array_ref(&order, txn) {
            index_shape(txn, state, &id, slide_id);
        }
    }
}

#[cfg(any(feature = "wasm", test))]
impl DeckSession {
    pub(crate) fn enable_history_receipts(&self) -> EditResult<()> {
        let mut undo = self.undo.borrow_mut();
        if undo.history.is_some() {
            return Ok(());
        }
        let txn = self.doc.transact();
        let mut initial = HistoryState::default();
        let slides = crate::deck::required_map(&txn, SLIDES)?;
        for (id, _) in slides.iter(&txn) {
            index_slide(&txn, &mut initial, id);
        }
        for (id, value) in crate::deck::required_map(&txn, COMMENTS)?.iter(&txn) {
            if let Ok(map) = value.cast::<MapRef>()
                && let Some(slide_id) = crate::deck::map_string(&map, &txn, "slideId")
            {
                initial.owners.insert(id.to_owned(), slide_id);
            }
        }
        let state = Arc::new(Mutex::new(initial));
        let mut subscriptions = Vec::new();
        for root in [SLIDES, SHAPES, STORIES, COMMENTS, META] {
            let map = crate::deck::required_map(&txn, root)?;
            let state = state.clone();
            subscriptions.push(map.observe_deep(move |txn, events| {
                let mut state = state.lock().unwrap();
                let mut keys = BTreeSet::new();
                for event in events.iter() {
                    if let Some(PathSegment::Key(key)) = event.path().front() {
                        keys.insert(key.to_string());
                    } else if let Event::Map(event) = event {
                        keys.extend(event.keys(txn).keys().map(|key| key.to_string()));
                    }
                }
                for key in keys {
                    let previous_owner = state.owners.get(&key).cloned();
                    let owner = match root {
                        SLIDES => Some(key.clone()),
                        COMMENTS => txn
                            .get_map(COMMENTS)
                            .and_then(|map| map.get(txn, &key))
                            .and_then(|value| value.cast::<MapRef>().ok())
                            .and_then(|map| crate::deck::map_string(&map, txn, "slideId"))
                            .or_else(|| previous_owner.clone()),
                        _ => previous_owner.clone(),
                    };
                    if let Some(receipt) = &mut state.receipt {
                        if root == META {
                            if key == "widthEmu" || key == "heightEmu" {
                                receipt.deck_size_changed = true;
                            } else if let Some(slides) = txn.get_map(SLIDES) {
                                receipt
                                    .changed_slide_ids
                                    .extend(slides.keys(txn).map(str::to_owned));
                            }
                        }
                        if let Some(id) = &owner {
                            receipt.changed_slide_ids.insert(id.clone());
                        }
                        if root == COMMENTS {
                            receipt.changed_slide_ids.extend(previous_owner);
                        }
                    }
                    if root == SLIDES {
                        index_slide(txn, &mut state, &key);
                    } else if let Some(id) = owner {
                        if root == SHAPES {
                            refresh_shape(txn, &mut state, &key, &id);
                        } else if root == COMMENTS {
                            state.owners.insert(key, id);
                        }
                    }
                }
            }));
        }
        let order = txn
            .get_array(SLIDE_ORDER)
            .ok_or_else(|| EditError::InvalidState("missing slide order".to_owned()))?;
        let order_state = state.clone();
        subscriptions.push(order.observe_deep(move |_, _| {
            if let Some(receipt) = &mut order_state.lock().unwrap().receipt {
                receipt.slide_order_changed = true;
            }
        }));
        undo.history = Some(HistoryObserver {
            state,
            _subscriptions: subscriptions,
        });
        Ok(())
    }

    pub(crate) fn history_step(&self, redo: bool) -> EditResult<HistoryReceipt> {
        self.enable_history_receipts()?;
        let mut undo = self.undo.borrow_mut();
        let state = undo.history.as_ref().unwrap().state.clone();
        state.lock().unwrap().receipt = Some(HistoryReceipt::default());
        let applied = if redo { undo.redo() } else { undo.undo() };
        let mut receipt = state.lock().unwrap().receipt.take().unwrap();
        receipt.applied = applied;
        Ok(receipt)
    }
}

impl DeckUndoManager {
    pub(crate) fn new(doc: &Doc, client_id: u64) -> EditResult<Self> {
        let txn = doc.transact();
        let order = txn
            .get_array(SLIDE_ORDER)
            .ok_or_else(|| EditError::InvalidState("missing slide order".to_owned()))?;
        let slides = txn
            .get_map(SLIDES)
            .ok_or_else(|| EditError::InvalidState("missing slides map".to_owned()))?;
        let shapes = txn
            .get_map(SHAPES)
            .ok_or_else(|| EditError::InvalidState("missing shapes map".to_owned()))?;
        let stories = txn
            .get_map(STORIES)
            .ok_or_else(|| EditError::InvalidState("missing stories map".to_owned()))?;
        let comments = txn
            .get_map(COMMENTS)
            .ok_or_else(|| EditError::InvalidState("missing comments map".to_owned()))?;
        let meta = txn
            .get_map(META)
            .ok_or_else(|| EditError::InvalidState("missing meta map".to_owned()))?;
        drop(txn);
        let clock = Arc::new(CaptureClock::new(default_clock()));
        let options = yrs::undo::Options {
            capture_timeout_millis: CAPTURE_TIMEOUT_MS,
            tracked_origins: HashSet::from([Origin::from(client_id)]),
            capture_transaction: None,
            timestamp: clock.clone(),
            init_undo_stack: Vec::new(),
            init_redo_stack: Vec::new(),
        };
        let mut inner = yrs::undo::UndoManager::with_options(options);
        inner.expand_scope(doc, &order);
        inner.expand_scope(doc, &slides);
        inner.expand_scope(doc, &shapes);
        inner.expand_scope(doc, &stories);
        inner.expand_scope(doc, &comments);
        inner.expand_scope(doc, &meta);
        Ok(Self {
            inner,
            clock,
            #[cfg(any(feature = "wasm", test))]
            history: None,
        })
    }

    pub(crate) fn assert_replay_capture(&self) -> Result<(), crate::peer::PeerError> {
        if self.capture_mode() != UndoCaptureMode::Manual {
            return Err(crate::peer::PeerError::new(
                "capturePolicy",
                "replay requires manual undo capture",
            ));
        }
        Ok(())
    }

    pub(crate) fn diagnostics(&self) -> (usize, usize) {
        (self.inner.undo_stack().len(), self.inner.redo_stack().len())
    }

    #[cfg(test)]
    pub(crate) fn stack_diagnostics(&self) -> (HistorySets, HistorySets) {
        let capture = |stack: &[yrs::undo::StackItem<()>]| {
            stack
                .iter()
                .map(|item| (item.deletions().clone(), item.insertions().clone()))
                .collect()
        };
        (
            capture(self.inner.undo_stack()),
            capture(self.inner.redo_stack()),
        )
    }

    #[cfg(test)]
    pub(crate) fn stack_clock_counts(&self) -> (HistoryCounts, HistoryCounts) {
        let counts = |set: &yrs::IdSet| -> ClientClockCounts {
            set.iter()
                .map(|(client, ranges)| {
                    let count = ranges
                        .iter()
                        .map(|range| u64::from(range.end - range.start))
                        .sum();
                    (*client, count)
                })
                .collect()
        };
        let capture = |stack: &[yrs::undo::StackItem<()>]| -> HistoryCounts {
            stack
                .iter()
                .map(|item| (counts(item.deletions()), counts(item.insertions())))
                .collect()
        };
        (
            capture(self.inner.undo_stack()),
            capture(self.inner.redo_stack()),
        )
    }

    pub fn capture_mode(&self) -> UndoCaptureMode {
        self.clock.mode()
    }

    pub fn set_capture_mode(&mut self, mode: UndoCaptureMode) {
        if self.capture_mode() != mode {
            self.add_undo_barrier();
            self.clock.set_mode(mode);
        }
    }

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

    pub fn add_undo_barrier(&mut self) {
        self.inner.reset();
    }
}

fn default_clock() -> Arc<dyn Clock> {
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
        Arc::new(move || ticks.fetch_add(CAPTURE_TIMEOUT_MS + 1, Ordering::Relaxed))
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::{EditCtx, ReadRequest, ShapeDraft, ShapeRect, TextStyle};

    fn session() -> DeckSession {
        let mut package = pptx_parse::PptxPackage::default();
        package.presentation.width_emu = 9144000;
        package.presentation.height_emu = 6858000;
        DeckSession::from_package_with_fingerprint(package, "history-receipts".to_owned(), 7)
            .unwrap()
    }

    #[test]
    fn history_receipts_scope_text_without_reading_untouched_slide() {
        let session = session();
        let context = EditCtx::local("test");
        session.set_undo_capture_mode(UndoCaptureMode::Manual);
        session.enable_history_receipts().unwrap();
        let mut slides = Vec::new();
        let mut shapes = Vec::new();
        for index in 0..2 {
            let slide = session.insert_slide(&context, index, None).unwrap();
            let shape = session
                .add_text_box(
                    &context,
                    &slide.slide_id,
                    &ShapeDraft {
                        name: "Text".to_owned(),
                        rect: ShapeRect {
                            x: 0,
                            y: 0,
                            width: 1000,
                            height: 1000,
                        },
                        text: "alpha".to_owned(),
                        style: TextStyle::default(),
                    },
                )
                .unwrap();
            slides.push(slide.slide_id);
            shapes.push(shape.shape_id);
        }
        session.add_undo_barrier();
        {
            let mut txn = session.doc.transact_mut_with(crate::REMOTE_ORIGIN);
            let shape = crate::deck::shape_ref(&txn, &shapes[1]).unwrap();
            shape.remove(&mut txn, "width");
        }
        assert!(session.snapshot().is_err());
        let story_id = format!("story:{}:0", shapes[0]);
        session
            .insert_text(&context, &story_id, 0, "Z", &TextStyle::default())
            .unwrap();
        session.add_undo_barrier();
        for (redo, text) in [(false, "alpha"), (true, "Zalpha")] {
            let receipt = session.history_step(redo).unwrap();
            assert!(receipt.applied);
            assert_eq!(
                receipt.changed_slide_ids,
                BTreeSet::from([slides[0].clone()])
            );
            assert!(!receipt.deck_size_changed);
            assert!(!receipt.slide_order_changed);
            let read = session
                .read_command_content(&ReadRequest {
                    slide_ids: Some(vec![slides[0].clone()]),
                })
                .unwrap()
                .unwrap();
            let value = serde_json::to_value(read).unwrap();
            assert_eq!(value["stories"][0]["text"], text);
        }
        assert!(session.snapshot().is_err());
    }

    #[test]
    fn history_receipts_report_slide_order_and_deck_size() {
        let session = session();
        let context = EditCtx::local("test");
        session.enable_history_receipts().unwrap();
        session.set_undo_capture_mode(UndoCaptureMode::Manual);
        let slide = session.insert_slide(&context, 0, None).unwrap();
        session.add_undo_barrier();
        for redo in [false, true] {
            let receipt = session.history_step(redo).unwrap();
            assert!(receipt.applied);
            assert!(receipt.slide_order_changed);
            assert!(!receipt.deck_size_changed);
            assert_eq!(
                receipt.changed_slide_ids,
                BTreeSet::from([slide.slide_id.clone()])
            );
        }
        session.add_undo_barrier();
        {
            let mut txn = session.transact_for(&context);
            let meta = crate::deck::required_map(&txn, META).unwrap();
            meta.insert(&mut txn, "widthEmu", 1000_f64);
        }
        session.add_undo_barrier();
        for redo in [false, true] {
            let receipt = session.history_step(redo).unwrap();
            assert!(receipt.applied);
            assert!(receipt.deck_size_changed);
            assert!(!receipt.slide_order_changed);
        }
    }
}
