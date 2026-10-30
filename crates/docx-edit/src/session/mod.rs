use std::cell::{Cell, RefCell};
use std::collections::VecDeque;
use std::rc::Rc;
use std::sync::Arc;

use ooxml_opc::PackageBytes;
use yrs::{StickyIndex, Subscription};

use crate::compare::{CompareApplied, CompareLimits};
use crate::{EngineSession, UndoSession};

pub use self::error::{EngineError, EngineErrorKind, EngineResult};
pub use self::host::Host;

pub(crate) mod args;
mod controls;
mod crdt;
mod display;
mod error;
mod exports;
mod fonts;
mod frames;
mod history;
mod host;
mod identity;
mod input;
mod json;
mod layout;
mod metadata;
mod open;
mod raw;
mod reads;
mod review;
mod selection;
mod tables;
mod text;

#[cfg(test)]
mod tests;

pub struct ParagraphIdentitiesCache {
    pub(crate) version: crate::DocumentVersion,
    pub(crate) source: Option<Arc<crate::identity::SourceIndex>>,
    pub(crate) json: String,
}

pub struct UpdateEventObserver {
    pub(crate) pending: Rc<RefCell<VecDeque<Vec<u8>>>>,
    pub(crate) _subscription: Subscription,
}

/// Per-peer selection state. These sticky positions are deliberately held
/// outside the yrs document: an awareness transport may publish them, but they
/// are never serialized as document content or included in save updates.
#[derive(Clone)]
pub struct LocalSelection {
    pub(crate) story: String,
    pub(crate) anchor: StickyIndex,
    pub(crate) head: StickyIndex,
}

/// One endpoint of a local table selection. The cell story is the stable
/// identity; row/column are a fallback if that cell is removed locally.
pub struct LocalCellPoint {
    pub(crate) cell_story: String,
    pub(crate) row: u32,
    pub(crate) column: u32,
}

/// Per-peer cell selection, held outside the yrs document. The sticky table
/// position survives unrelated edits in the parent story; cell-story identity
/// lets each endpoint follow inserted/deleted rows and columns.
pub struct LocalCellSelection {
    pub(crate) parent_story: String,
    pub(crate) table: StickyIndex,
    pub(crate) anchor: LocalCellPoint,
    pub(crate) head: LocalCellPoint,
}

pub struct EditSessionCore {
    host: Host,
    pub(crate) engine: EngineSession,
    pub(crate) paragraph_identities_cache: RefCell<Option<ParagraphIdentitiesCache>>,
    /// This session's measurement fonts; see the module docs.
    pub(crate) fonts: docx_layout::MeasureFonts,
    pub(crate) docx_source: RefCell<Option<PackageBytes>>,
    paragraph_styles: RefCell<Option<crate::seed::StyleResolver>>,
    /// The [`crate::seed::package_digest`] of `docx_source`, when known.
    pub(crate) docx_digest: RefCell<Option<String>>,
    awaiting_comment_baseline: Cell<bool>,
    update_event_observer: Option<UpdateEventObserver>,
    pub(crate) undo: UndoSession,
    pub(crate) selection: RefCell<Option<LocalSelection>>,
    cell_selection: RefCell<Option<LocalCellSelection>>,
    last_apply_profile_json: RefCell<String>,
    resident_deleted_units: Cell<u32>,
    /// The story and resulting doc epoch of the last resident text edit, until a relayout takes it.
    pub(crate) resident_text_edit: RefCell<Option<(String, u64)>>,
    /// The comparison applied here, awaiting its saved bytes.
    compared: RefCell<Option<(Box<CompareApplied>, CompareLimits)>>,
    /// Seed images as `media:{n}` tokens rather than `data:` URLs.
    media_tokens: Cell<bool>,
}

impl EditSessionCore {
    pub fn new(client_id: u64, host: Host) -> Self {
        let clock_host = host.clone();
        let session = Self {
            host: host.clone(),
            engine: EngineSession::new(client_id),
            paragraph_identities_cache: RefCell::new(None),
            fonts: docx_layout::MeasureFonts::default(),
            docx_source: RefCell::new(None),
            paragraph_styles: RefCell::new(None),
            docx_digest: RefCell::new(None),
            awaiting_comment_baseline: Cell::new(false),
            update_event_observer: None,
            undo: UndoSession::with_clock(Arc::new(move || clock_host.clock_ms())),
            selection: RefCell::new(None),
            cell_selection: RefCell::new(None),
            last_apply_profile_json: RefCell::new("{}".to_owned()),
            resident_deleted_units: Cell::new(0),
            resident_text_edit: RefCell::new(None),
            compared: RefCell::new(None),
            media_tokens: Cell::new(false),
        };
        session.engine.doc().rotate_version(host.entropy());
        session
    }

    pub fn client_id(&self) -> f64 {
        self.engine.doc().client_id() as f64
    }
}
