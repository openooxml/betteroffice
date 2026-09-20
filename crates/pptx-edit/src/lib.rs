//! Collaborative yrs-backed PPTX deck model.

use std::cell::RefCell;
use std::panic::{AssertUnwindSafe, catch_unwind};
use std::sync::Arc;
use std::sync::atomic::{AtomicU64, Ordering};

use yrs::updates::decoder::{Decode, Decoder, DecoderV1};

use pptx_parse::PptxPackage;
use sha2::{Digest, Sha256};
use yrs::updates::encoder::Encode;
use yrs::{
    ClientID, Doc, OffsetKind, Options, ReadTxn, StateVector, StickyIndex, Subscription, Transact,
    Update, WriteTxn,
};

mod comments;
mod deck;
mod effects;
mod model;
mod outline_gradients;
mod proposal_diff;
mod proposals;
mod save;
mod search;
mod source_run_properties;
mod story;
mod undo;

pub use model::*;
pub use proposal_diff::*;
pub use proposals::*;
pub use search::TextSearchMatch;
pub use undo::DeckUndoManager;

#[cfg(feature = "wasm")]
pub mod wasm;

pub(crate) const META: &str = "pptx:meta";
pub(crate) const SLIDE_ORDER: &str = "pptx:slide-order";
pub(crate) const SLIDES: &str = "pptx:slides";
pub(crate) const SHAPES: &str = "pptx:shapes";
pub(crate) const STORIES: &str = "pptx:stories";
pub(crate) const COMMENTS: &str = "pptx:comments";
pub(crate) const REMOTE_ORIGIN: &str = "pptx:remote";
pub(crate) const HYDRATE_ORIGIN: &str = "pptx:hydrate";
pub(crate) const MIGRATE_ORIGIN: &str = "pptx:migrate";
pub(crate) const PILCROW_KIND: &str = "pilcrow";
pub(crate) const KIND: &str = "_kind";
pub(crate) const PARA_ID: &str = "paraId";
pub(crate) const BOOTSTRAP_CLIENT_ID: u64 = (1_u64 << 53) - 1;
pub const MAX_SAFE_CLIENT_ID: u64 = BOOTSTRAP_CLIENT_ID - 1;
pub const MAX_UPDATE_BYTES: usize = 64 * 1024 * 1024;
const MAX_STATE_VECTOR_ENTRIES: u32 = 65_536;

pub type UpdateSubscription = Subscription;

#[derive(Clone, Debug, Eq, PartialEq)]
pub struct CaretAnchor {
    story_id: String,
    position: StickyIndex,
}

pub struct DeckSession {
    pub(crate) doc: Doc,
    /// Persistent clone of `doc` that remote updates stage against before
    /// adoption, so a rejected update leaves `doc` untouched without paying a
    /// full clone per apply. Built lazily and re-synced by diff.
    staged: RefCell<Option<Doc>>,
    /// The `packageJson` bytes `doc` currently stores — proven to parse — so
    /// updates that leave them untouched skip the deserialize.
    package_json: RefCell<Arc<[u8]>>,
    client_id: u64,
    id_counter: AtomicU64,
    package: Arc<PptxPackage>,
    undo: RefCell<DeckUndoManager>,
    proposals: RefCell<proposals::ProposalStore>,
}

impl DeckSession {
    pub fn open(bytes: &[u8], client_id: u64) -> EditResult<Self> {
        let package =
            pptx_parse::parse_pptx(bytes).map_err(|error| EditError::Parse(error.to_string()))?;
        Self::from_package_with_source(package, bytes, client_id)
    }

    /// Opens an edit session from an already parsed package. The fingerprint
    /// is taken from a re-zip of the package; prefer
    /// [`Self::from_package_with_source`] when the file bytes are at hand, so
    /// peers can match them against the update.
    pub fn from_package(package: PptxPackage, client_id: u64) -> EditResult<Self> {
        let bytes = pptx_parse::write_pptx(&package)
            .map_err(|error| EditError::Parse(error.to_string()))?;
        let fingerprint = format!("{:x}", Sha256::digest(bytes));
        Self::from_package_with_fingerprint(package, fingerprint, client_id)
    }

    /// Opens an edit session from a parsed package, fingerprinting the file
    /// bytes it was parsed from.
    pub fn from_package_with_source(
        package: PptxPackage,
        source: &[u8],
        client_id: u64,
    ) -> EditResult<Self> {
        let fingerprint = format!("{:x}", Sha256::digest(source));
        Self::from_package_with_fingerprint(package, fingerprint, client_id)
    }

    fn from_package_with_fingerprint(
        package: PptxPackage,
        fingerprint: String,
        client_id: u64,
    ) -> EditResult<Self> {
        validate_client_id(client_id)?;
        let bootstrap = doc_with_client_id(BOOTSTRAP_CLIENT_ID);
        deck::seed_doc(&bootstrap, &package, &fingerprint)?;
        let baseline = bootstrap
            .transact()
            .encode_state_as_update_v1(&StateVector::default());
        let doc = doc_with_client_id(client_id);
        hydrate_doc(&doc, &baseline)?;
        deck::validate_doc(&doc)?;
        Self::assemble(doc, client_id, package)
    }

    pub fn open_from_update(update: &[u8], client_id: u64) -> EditResult<Self> {
        validate_client_id(client_id)?;
        if update.len() > MAX_UPDATE_BYTES {
            return Err(EditError::InvalidUpdate(format!(
                "update exceeds {MAX_UPDATE_BYTES} bytes"
            )));
        }
        let doc = doc_with_client_id(client_id);
        hydrate_doc(&doc, update)?;
        deck::migrate_doc(&doc)?;
        deck::validate_doc(&doc)?;
        let package = deck::package_from_doc(&doc)?;
        Self::assemble(doc, client_id, package)
    }

    fn assemble(doc: Doc, client_id: u64, package: PptxPackage) -> EditResult<Self> {
        let undo = DeckUndoManager::new(&doc, client_id)?;
        let package_json = deck::package_json_bytes(&doc)?;
        Ok(Self {
            doc,
            staged: RefCell::new(None),
            package_json: RefCell::new(package_json),
            client_id,
            id_counter: AtomicU64::new(0),
            package: Arc::new(package),
            undo: RefCell::new(undo),
            proposals: Default::default(),
        })
    }

    /// Like [`Self::open_from_update`], but re-attaches the source file the
    /// update was seeded from, so the session can save. The bytes must hash to
    /// the fingerprint recorded in the update.
    pub fn open_from_update_with_source(
        update: &[u8],
        source: &[u8],
        client_id: u64,
    ) -> EditResult<Self> {
        let session = Self::open_from_update(update, client_id)?;
        let recorded = deck::fingerprint_from_doc(&session.doc)?;
        let actual = format!("{:x}", Sha256::digest(source));
        if recorded != actual {
            return Err(EditError::Parse(
                "source bytes do not match the fingerprint recorded in the update".to_owned(),
            ));
        }
        let package = if session.package.models_connectors() {
            pptx_parse::parse_pptx(source)
        } else {
            pptx_parse::parse_pptx_without_connectors(source)
        }
        .map_err(|error| EditError::Parse(error.to_string()))?;
        comments::import_source_comments(&session, &package)?;
        deck::import_source_render_data(&session.doc, &package)?;
        source_run_properties::import_source(
            &session,
            &package,
            source_run_properties::SourceProperty::Baseline,
        )?;
        deck::import_source_ole_pictures(&session.doc, &package)?;
        effects::import_source(&session.doc, &package)?;
        source_run_properties::import_source(
            &session,
            &package,
            source_run_properties::SourceProperty::Spacing,
        )?;
        story::import_source_numbering_restarts(&session.doc, &package)?;
        outline_gradients::import_source(&session, &package)?;
        let session = Self {
            package: Arc::new(package),
            ..session
        };
        // Re-record proven `packageJson` bytes after import passes.
        *session.package_json.borrow_mut() = deck::package_json_bytes(&session.doc)?;
        Ok(session)
    }

    pub fn client_id(&self) -> u64 {
        self.client_id
    }

    pub fn package(&self) -> &PptxPackage {
        &self.package
    }

    pub fn yrs_doc(&self) -> &Doc {
        &self.doc
    }

    pub fn encode_state_vector_v1(&self) -> Vec<u8> {
        self.doc.transact().state_vector().encode_v1()
    }

    pub fn encode_state_as_update_v1(&self) -> Vec<u8> {
        self.doc
            .transact()
            .encode_state_as_update_v1(&StateVector::default())
    }

    pub fn encode_diff_v1(&self, remote_state_vector: &[u8]) -> EditResult<Vec<u8>> {
        let state_vector =
            decode_state_vector_v1(remote_state_vector).map_err(EditError::InvalidStateVector)?;
        Ok(self.doc.transact().encode_diff_v1(&state_vector))
    }

    pub fn apply_update_v1(&self, bytes: &[u8]) -> EditResult<DeckSnapshot> {
        if bytes.len() > MAX_UPDATE_BYTES {
            return Err(EditError::InvalidUpdate(format!(
                "update exceeds {MAX_UPDATE_BYTES} bytes"
            )));
        }
        let incoming = decode_update_v1(bytes).map_err(EditError::InvalidUpdate)?;
        self.sync_staged()?;
        let validated = {
            let staged = self.staged.borrow();
            let staged = staged.as_ref().expect("staged doc synced");
            // Drop the mut txn (committing) before validation reads the store.
            let applied = staged
                .transact_mut_with(REMOTE_ORIGIN)
                .apply_update(incoming)
                .map_err(|error| EditError::InvalidUpdate(error.to_string()));
            applied.and_then(|()| {
                deck::validate_remote_doc(staged, &self.package, &self.package_json.borrow())
            })
        };
        let (snapshot, package_json) = match validated {
            Ok(ok) => ok,
            Err(error) => {
                self.rebuild_staged()?;
                return Err(error);
            }
        };

        let incoming = decode_update_v1(bytes).map_err(EditError::InvalidUpdate)?;
        self.doc
            .transact_mut_with(REMOTE_ORIGIN)
            .apply_update(incoming)
            .map_err(|error| EditError::InvalidUpdate(error.to_string()))?;
        *self.package_json.borrow_mut() = package_json;
        Ok(snapshot)
    }

    /// Diff-replays `doc`'s state into `staged`; re-hydrates on drift.
    fn sync_staged(&self) -> EditResult<()> {
        let doc_state = DocState::of(&self.doc);
        let staged_state = self.staged.borrow().as_ref().map(DocState::of);
        match staged_state {
            None => self.rebuild_staged(),
            Some(staged_state) if staged_state == doc_state => Ok(()),
            Some(staged_state) if state_vector_exceeds(&staged_state.sv, &doc_state.sv) => {
                self.rebuild_staged()
            }
            Some(staged_state) => {
                let diff = self.doc.transact().encode_diff_v1(&staged_state.sv);
                let caught_up = decode_update_v1(&diff).and_then(|update| {
                    self.staged
                        .borrow()
                        .as_ref()
                        .expect("staged doc synced")
                        .transact_mut()
                        .apply_update(update)
                        .map_err(|error| error.to_string())
                });
                let synced = caught_up.is_ok()
                    && self
                        .staged
                        .borrow()
                        .as_ref()
                        .map(|staged| DocState::of(staged) == doc_state)
                        .unwrap_or(false);
                if synced {
                    Ok(())
                } else {
                    self.rebuild_staged()
                }
            }
        }
    }

    fn rebuild_staged(&self) -> EditResult<()> {
        let staged = doc_with_client_id(self.client_id);
        hydrate_doc(&staged, &self.encode_state_as_update_v1())?;
        *self.staged.borrow_mut() = Some(staged);
        Ok(())
    }

    pub fn observe_update_v1<F>(&self, callback: F) -> EditResult<Subscription>
    where
        F: Fn(UpdateEvent) + 'static,
    {
        self.doc
            .observe_update_v1(move |txn, event| {
                let origin = if txn
                    .origin()
                    .is_some_and(|origin| origin.as_ref() == REMOTE_ORIGIN.as_bytes())
                {
                    UpdateOrigin::Remote
                } else {
                    UpdateOrigin::Local
                };
                let update = UpdateEvent {
                    update: event.update.clone(),
                    origin,
                };
                let _ = catch_unwind(AssertUnwindSafe(|| callback(update)));
            })
            .map_err(|error| EditError::Observer(error.to_string()))
    }

    pub fn undo(&self) -> bool {
        self.undo.borrow_mut().undo()
    }

    pub fn redo(&self) -> bool {
        self.undo.borrow_mut().redo()
    }

    pub fn can_undo(&self) -> bool {
        self.undo.borrow().can_undo()
    }

    pub fn can_redo(&self) -> bool {
        self.undo.borrow().can_redo()
    }

    pub fn add_undo_barrier(&self) {
        self.undo.borrow_mut().add_undo_barrier();
    }

    pub(crate) fn transact_for(&self, context: &EditCtx) -> yrs::TransactionMut<'_> {
        match context.origin {
            EditOrigin::Local => self.doc.transact_mut_with(self.client_id),
            EditOrigin::Agent => self.doc.transact_mut_with("pptx:agent"),
            EditOrigin::Remote => self.doc.transact_mut_with(REMOTE_ORIGIN),
            EditOrigin::System => self.doc.transact_mut_with("pptx:system"),
        }
    }

    pub(crate) fn next_id(&self, prefix: &str) -> String {
        let counter = self.id_counter.fetch_add(1, Ordering::Relaxed);
        format!("{prefix}:{}:{counter}", self.client_id)
    }
}

pub(crate) fn doc_with_client_id(client_id: u64) -> Doc {
    let mut options = Options::with_client_id(ClientID::new(client_id));
    options.offset_kind = OffsetKind::Utf16;
    Doc::with_options(options)
}

fn validate_client_id(client_id: u64) -> EditResult<()> {
    if client_id == 0 || client_id > MAX_SAFE_CLIENT_ID {
        return Err(EditError::InvalidClientId(client_id));
    }
    Ok(())
}

fn state_vector_exceeds(a: &StateVector, b: &StateVector) -> bool {
    a.iter().any(|(client, clock)| *clock > b.get(client))
}

/// Replication fingerprint: the state vector misses deletion-only edits, so
/// the delete set is compared alongside it.
#[derive(PartialEq)]
struct DocState {
    sv: StateVector,
    ds: Vec<u8>,
}

impl DocState {
    fn of(doc: &Doc) -> Self {
        let snapshot = doc.transact().snapshot();
        Self {
            sv: snapshot.state_map,
            ds: snapshot.delete_set.encode_v1(),
        }
    }
}

fn hydrate_doc(doc: &Doc, bytes: &[u8]) -> EditResult<()> {
    let update = decode_update_v1(bytes).map_err(EditError::InvalidUpdate)?;
    let mut txn = doc.transact_mut_with(HYDRATE_ORIGIN);
    txn.apply_update(update)
        .map_err(|error| EditError::InvalidUpdate(error.to_string()))?;
    // An empty container carries no items, so an update cannot carry it either.
    // Every peer registers the roots the same way, which stays convergent.
    txn.get_or_insert_array(SLIDE_ORDER);
    for root in [META, SLIDES, SHAPES, STORIES, COMMENTS] {
        txn.get_or_insert_map(root);
    }
    Ok(())
}

fn decode_update_v1(bytes: &[u8]) -> Result<Update, String> {
    let mut decoder = DecoderV1::from(bytes);
    let update = Update::decode(&mut decoder).map_err(|error| error.to_string())?;
    if !decoder
        .read_to_end()
        .map_err(|error| error.to_string())?
        .is_empty()
    {
        return Err("update contains trailing bytes".to_owned());
    }
    Ok(update)
}

fn decode_state_vector_v1(bytes: &[u8]) -> Result<StateVector, String> {
    validate_state_vector_entry_count(bytes)?;
    let mut decoder = DecoderV1::from(bytes);
    let state_vector = StateVector::decode(&mut decoder).map_err(|error| error.to_string())?;
    if !decoder
        .read_to_end()
        .map_err(|error| error.to_string())?
        .is_empty()
    {
        return Err("state vector contains trailing bytes".to_owned());
    }
    Ok(state_vector)
}

fn validate_state_vector_entry_count(bytes: &[u8]) -> Result<(), String> {
    let Some((&first, _)) = bytes.split_first() else {
        return Err("state vector is empty".to_owned());
    };
    let mut value = u32::from(first & 0x7f);
    let mut shift = 7;
    let mut used = 1;
    let mut byte = first;
    while byte & 0x80 != 0 {
        if used == 5 || used >= bytes.len() {
            return Err("invalid state vector entry count".to_owned());
        }
        byte = bytes[used];
        if used == 4 && byte > 0x0f {
            return Err("invalid state vector entry count".to_owned());
        }
        value |= u32::from(byte & 0x7f) << shift;
        shift += 7;
        used += 1;
    }
    if value > MAX_STATE_VECTOR_ENTRIES {
        return Err(format!(
            "state vector contains {value} entries, exceeds the {MAX_STATE_VECTOR_ENTRIES}-entry limit"
        ));
    }
    if value as usize > bytes.len().saturating_sub(used) / 2 {
        return Err("state vector entry count exceeds its payload".to_owned());
    }
    Ok(())
}
