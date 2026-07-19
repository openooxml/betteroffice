//! Collaborative yrs-backed PPTX deck model.

use std::cell::RefCell;
use std::panic::{AssertUnwindSafe, catch_unwind};
use std::sync::Arc;
use std::sync::atomic::{AtomicU64, Ordering};

use pptx_parse::PptxPackage;
use sha2::{Digest, Sha256};
use yrs::updates::decoder::{Decode, Decoder, DecoderV1};
use yrs::updates::encoder::Encode;
use yrs::{
    ClientID, Doc, OffsetKind, Options, ReadTxn, StateVector, Subscription, Transact, Update,
};

mod deck;
mod model;
mod story;
mod undo;

pub use model::*;
pub use undo::DeckUndoManager;

#[cfg(feature = "wasm")]
pub mod wasm;

pub(crate) const META: &str = "pptx:meta";
pub(crate) const SLIDE_ORDER: &str = "pptx:slide-order";
pub(crate) const SLIDES: &str = "pptx:slides";
pub(crate) const SHAPES: &str = "pptx:shapes";
pub(crate) const STORIES: &str = "pptx:stories";
pub(crate) const REMOTE_ORIGIN: &str = "pptx:remote";
pub(crate) const HYDRATE_ORIGIN: &str = "pptx:hydrate";
pub(crate) const PILCROW_KIND: &str = "pilcrow";
pub(crate) const KIND: &str = "_kind";
pub(crate) const PARA_ID: &str = "paraId";
const BOOTSTRAP_CLIENT_ID: u64 = (1_u64 << 53) - 1;
const MAX_SAFE_CLIENT_ID: u64 = BOOTSTRAP_CLIENT_ID - 1;
const MAX_UPDATE_BYTES: usize = 64 * 1024 * 1024;
const MAX_STATE_VECTOR_ENTRIES: u32 = 65_536;

pub struct DeckSession {
    pub(crate) doc: Doc,
    client_id: u64,
    id_counter: AtomicU64,
    package: Arc<PptxPackage>,
    undo: RefCell<DeckUndoManager>,
}

impl DeckSession {
    pub fn open(bytes: &[u8], client_id: u64) -> EditResult<Self> {
        validate_client_id(client_id)?;
        let package =
            pptx_parse::parse_pptx(bytes).map_err(|error| EditError::Parse(error.to_string()))?;
        let fingerprint = format!("{:x}", Sha256::digest(bytes));
        let bootstrap = doc_with_client_id(BOOTSTRAP_CLIENT_ID);
        deck::seed_doc(&bootstrap, &package, &fingerprint)?;
        let baseline = bootstrap
            .transact()
            .encode_state_as_update_v1(&StateVector::default());
        let doc = doc_with_client_id(client_id);
        hydrate_doc(&doc, &baseline)?;
        deck::validate_doc(&doc)?;
        let undo = DeckUndoManager::new(&doc, client_id)?;
        Ok(Self {
            doc,
            client_id,
            id_counter: AtomicU64::new(0),
            package: Arc::new(package),
            undo: RefCell::new(undo),
        })
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
        deck::validate_doc(&doc)?;
        let package = deck::package_from_doc(&doc)?;
        let undo = DeckUndoManager::new(&doc, client_id)?;
        Ok(Self {
            doc,
            client_id,
            id_counter: AtomicU64::new(0),
            package: Arc::new(package),
            undo: RefCell::new(undo),
        })
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
        let staged = doc_with_client_id(self.client_id);
        hydrate_doc(&staged, &self.encode_state_as_update_v1())?;
        staged
            .transact_mut_with(REMOTE_ORIGIN)
            .apply_update(incoming)
            .map_err(|error| EditError::InvalidUpdate(error.to_string()))?;
        deck::validate_doc(&staged)?;

        let incoming = decode_update_v1(bytes).map_err(EditError::InvalidUpdate)?;
        self.doc
            .transact_mut_with(REMOTE_ORIGIN)
            .apply_update(incoming)
            .map_err(|error| EditError::InvalidUpdate(error.to_string()))?;
        self.snapshot()
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

fn doc_with_client_id(client_id: u64) -> Doc {
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

fn hydrate_doc(doc: &Doc, bytes: &[u8]) -> EditResult<()> {
    let update = decode_update_v1(bytes).map_err(EditError::InvalidUpdate)?;
    doc.transact_mut_with(HYDRATE_ORIGIN)
        .apply_update(update)
        .map_err(|error| EditError::InvalidUpdate(error.to_string()))
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
