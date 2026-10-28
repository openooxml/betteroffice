use std::cell::RefCell;
use std::collections::{BTreeMap, VecDeque};
use std::rc::Rc;
use std::sync::Arc;

use ooxml_opc::PackageBytes;
use serde_json::{Value, json};

use super::{EditSessionCore, EngineError, EngineResult, UpdateEventObserver};
use crate::engine::RelayoutTrigger;
use crate::presence::apply_update_with_typing_inference;
use crate::{EditingDoc, SeedParagraph};

fn seed_paragraph(value: &Value) -> (String, String, String) {
    let text = value
        .get("text")
        .and_then(Value::as_str)
        .unwrap_or_default()
        .to_owned();
    let p_style = value
        .get("pStyle")
        .and_then(Value::as_str)
        .unwrap_or("Normal")
        .to_owned();
    let alignment = value
        .get("alignment")
        .and_then(Value::as_str)
        .unwrap_or("left")
        .to_owned();
    (text, p_style, alignment)
}

impl EditSessionCore {
    /// Hydrates from a yrs v1 update. The first load after an unseeded open retains prior
    /// comment writes and marks loaded fields differing from seed placeholders as authored.
    pub fn load(&self, update: &[u8]) -> EngineResult<()> {
        let session_id = self.engine.doc().session_id();
        let written = self.awaiting_comment_baseline.get().then(|| {
            self.engine
                .doc()
                .source_metadata()
                .map(|source| source.read().comment_writes.snapshot())
                .unwrap_or_default()
        });
        self.engine.doc().load_state_v1(update).map_err(EngineError::message)?;
        if self.engine.doc().session_id() != session_id {
            self.engine.set_relayout_trigger(RelayoutTrigger::Open);
        }
        if let Some(written) = written {
            self.engine.doc().rebase_comment_writes(written);
            self.awaiting_comment_baseline.set(false);
        }
        self.engine.doc().rotate_version(self.host.entropy());
        Ok(())
    }

    pub fn encode_peer_metadata(&self) -> EngineResult<Vec<u8>> {
        self.engine
            .doc()
            .encode_peer_metadata()
            .map_err(EngineError::from)
    }

    pub fn bootstrap_peer(
        &self,
        state: &[u8],
        metadata: &[u8],
        source: Option<Vec<u8>>,
    ) -> EngineResult<()> {
        let source = source.map(PackageBytes::from);
        let bootstrap = self
            .engine
            .doc()
            .prepare_peer_bootstrap(state, metadata, source)
            .map_err(EngineError::from)?;
        let retained = self
            .engine
            .doc()
            .install_peer_bootstrap(bootstrap, self.host.entropy())
            .map_err(EngineError::from)?;
        self.docx_source.replace(Some(retained.source));
        self.paragraph_styles.replace(None);
        self.docx_digest.replace(Some(retained.digest));
        self.awaiting_comment_baseline.set(true);
        self.engine.set_partial_document(false);
        self.engine.set_relayout_trigger(RelayoutTrigger::Open);
        self.engine.doc().rotate_version(self.host.entropy());
        self.load(state)
    }


    /// Starts a new opening of the document; see [`EditingDoc::begin_opening`].
    pub fn begin_opening(&self, generation: Option<String>) {
        self.engine.doc().begin_opening(generation.as_deref());
        self.engine.set_relayout_trigger(RelayoutTrigger::Open);
    }

    /// Declares header/footer alias groups before host-side seeding.
    pub fn set_header_footer_aliases(&self, json: &str) -> EngineResult<()> {
        self.engine
            .doc()
            .set_header_footer_aliases(json)
            .map_err(EngineError::message)
    }

    /// Returns active alias-to-canonical relationship ids as a JSON object.
    pub fn header_footer_aliases_json(&self) -> String {
        let aliases: BTreeMap<_, _> = self
            .engine
            .doc()
            .header_footer_aliases()
            .into_iter()
            .collect();
        serde_json::to_string(&aliases).expect("header/footer aliases serialize")
    }

    /// Unions seeded opaque sequence names into document state.
    pub fn seed_opaque_sequences(&self, names_json: &str) -> EngineResult<()> {
        let names: Vec<String> = serde_json::from_str(names_json).map_err(EngineError::message)?;
        crate::seed::seed_opaque_sequences(self.engine.doc(), &names, None);
        Ok(())
    }

    /// Whether [`EditSessionCore::open_docx`] seeds images as `media:{n}` tokens,
    /// which only a replica opened from the same package resolves, instead of
    /// `data:` URLs. Off by default.
    pub fn set_media_tokens(&self, enabled: bool) {
        self.media_tokens.set(enabled);
    }

    /// The digests mapping the `data:` URLs this replica seeded to
    /// `media:{n}` tokens, for [`EditSessionCore::load_media_sources`]; empty
    /// when it seeded none.
    pub fn media_sources_json(&self) -> String {
        let sources = self.engine.doc().media_sources();
        if sources.is_empty() {
            String::new()
        } else {
            sources.to_json()
        }
    }

    /// Lays this replica's `data:` image sources out as the `media:{n}` tokens
    /// another replica seeded them from.
    pub fn load_media_sources(&self, json: &str) -> EngineResult<()> {
        let sources = if json.is_empty() {
            crate::media::MediaSources::default()
        } else {
            crate::media::MediaSources::from_json(json)
                .map_err(|error| EngineError::message(&error))?
        };
        self.engine.doc().set_media_sources(sources);
        Ok(())
    }

    /// The separator notes for another replica; empty when the package has none.
    pub fn note_separators_state(&self) -> EngineResult<Vec<u8>> {
        self.engine
            .doc()
            .note_separator_state()
            .map(|state| state.map_or_else(Vec::new, |state| state.to_vec()))
            .map_err(|error| EngineError::message(&error))
    }

    /// Loads separator notes from a yrs v1 update; empty clears them.
    pub fn load_note_separators(&self, state: &[u8]) -> EngineResult<()> {
        if self.engine.doc().has_note_separator_state(state) {
            return Ok(());
        }
        let state = if state.is_empty() {
            None
        } else {
            EditingDoc::new(1)
                .load_state_v1(state)
                .map_err(|error| EngineError::message(error.to_string()))?;
            Some(Arc::from(state))
        };
        self.engine.doc().set_note_separator_state(state);
        Ok(())
    }

    /// Marks whether the session's document is part of a package, as a
    /// replica of a preview is: its layouts render NUMPAGES empty.
    pub fn set_partial_document(&self, partial: bool) {
        self.engine.set_partial_document(partial);
    }

    /// Re-parses the DOCX bytes retained by the last
    /// [`EditSessionCore::open_docx`] and returns the COMPLETE package envelope as
    /// JSON, or `None` when no DOCX has been opened. Unlike `open_docx` this
    /// keeps every part; it reflects the source file, not later edits.
    pub fn materialize_docx(&self) -> EngineResult<Option<String>> {
        let source = self.docx_source.borrow();
        let Some(source) = source.as_ref() else {
            return Ok(None);
        };
        let digest = self
            .docx_digest
            .borrow()
            .clone()
            .unwrap_or_else(|| crate::seed::package_digest(source));
        let (envelope, _) =
            crate::seed::parse_docx_package_with_digest(source, digest)
                .map_err(EngineError::message)?;
        let json = serde_json::to_string(&envelope).map_err(EngineError::message)?;
        Ok(Some(json))
    }

    /// The bytes a `media:{n}` image source displays, as the part's
    /// [`EditSessionCore::media_type`]; `undefined` for any other source or a
    /// part that cannot be read.
    pub fn media_bytes(&self, token: &str) -> Option<Vec<u8>> {
        let media = self.engine.doc().media_table()?;
        let index = docx_parse::media::media_token_index(token)?;
        media.bytes(index).ok().map(std::borrow::Cow::into_owned)
    }

    /// The media type of [`EditSessionCore::media_bytes`].
    pub fn media_type(&self, token: &str) -> Option<String> {
        let media = self.engine.doc().media_table()?;
        media
            .mime_type(docx_parse::media::media_token_index(token)?)
            .map(str::to_owned)
    }

    /// The `data:` URL a `media:{n}` image source stands for.
    pub fn media_data_url(&self, token: &str) -> Option<String> {
        self.engine.doc().media_table()?.resolve(token)
    }

    /// Seeds stories from JSON:
    /// `[{"storyId","paragraphs":[{"text","pStyle"?,"alignment"?}, …]}, …]`.
    /// `pStyle` defaults to `"Normal"` and `alignment` to `"left"`; paragraph
    /// text must not contain paragraph breaks. Receipt:
    /// `{storyId: [paraId, …]}` with each story's paragraphs in document
    /// order. Errors when a story entry has no `storyId`, no `paragraphs`
    /// array, or an empty one, and when a story id already exists. Seeding a
    /// document that has no opening yet starts one; see
    /// [`EditingDoc::begin_opening`].
    pub fn load_json(&self, stories_json: &str) -> EngineResult<String> {
        let value: Value = serde_json::from_str(stories_json).map_err(EngineError::message)?;
        let entries = value
            .as_array()
            .ok_or_else(|| EngineError::message("load_json expects an array of stories"))?;
        let mut receipt = serde_json::Map::new();
        for entry in entries {
            let story_id = entry
                .get("storyId")
                .and_then(Value::as_str)
                .ok_or_else(|| {
                    EngineError::message("a story entry requires a string \"storyId\"")
                })?;
            let paragraphs = entry
                .get("paragraphs")
                .and_then(Value::as_array)
                .ok_or_else(|| {
                    EngineError::message("a story entry requires a \"paragraphs\" array")
                })?;
            if paragraphs.is_empty() {
                return Err(EngineError::message(format!(
                    "story {story_id:?} requires at least one paragraph"
                )));
            }

            let seed_paragraphs: Vec<SeedParagraph> = paragraphs
                .iter()
                .map(|paragraph| {
                    let (text, p_style, alignment) = seed_paragraph(paragraph);
                    SeedParagraph {
                        text,
                        p_style,
                        alignment,
                    }
                })
                .collect();
            let para_ids: Vec<Value> = self
                .engine
                .doc()
                .seed_story(story_id, &seed_paragraphs)
                .map_err(EngineError::message)?
                .into_iter()
                .map(Value::String)
                .collect();
            receipt.insert(story_id.to_owned(), Value::Array(para_ids));
            self.awaiting_comment_baseline.set(false);
        }
        if !self.engine.doc().has_opening() {
            self.engine.doc().begin_opening(None);
        }
        self.engine.set_relayout_trigger(RelayoutTrigger::Open);
        serde_json::to_string(&Value::Object(receipt)).map_err(EngineError::message)
    }

    /// The full document state as one yrs v1 update. Hand it to
    /// [`EditSessionCore::load`] on a fresh replica to reproduce this document.
    pub fn encode_state(&self) -> Vec<u8> {
        self.engine.doc().encode_state_as_update_v1()
    }

    /// This replica's yrs v1 state vector — what a peer sends so
    /// [`EditSessionCore::encode_diff`] can compute the update it is missing.
    pub fn encode_state_vector(&self) -> Vec<u8> {
        self.engine.doc().encode_state_vector_v1()
    }

    /// Whether a remote update is waiting on one that has not arrived yet.
    pub fn has_pending_update(&self) -> bool {
        self.engine.doc().holds_pending_update()
    }

    /// The yrs v1 update carrying everything this replica has that the peer
    /// described by `remote_state_vector` does not. Errors on a malformed
    /// state vector.
    pub fn encode_diff(&self, remote_state_vector: &[u8]) -> EngineResult<Vec<u8>> {
        self.engine
            .doc()
            .encode_diff_v1(remote_state_vector)
            .map_err(EngineError::message)
    }

    /// [`EditSessionCore::encode_diff`] as ordered updates for a sync that must
    /// travel in bounded frames: the whole diff when it fits
    /// `max_chunk_bytes`, otherwise causally ordered struct updates followed
    /// by the delete set alone. Each fits the limit except one that carries a
    /// single larger struct. Errors on a malformed state vector or a zero
    /// limit.
    pub fn encode_diff_chunks(
        &self,
        remote_state_vector: &[u8],
        max_chunk_bytes: u32,
    ) -> EngineResult<Vec<Vec<u8>>> {
        self.engine
            .doc()
            .encode_diff_chunks_v1(remote_state_vector, max_chunk_bytes as usize)
            .map_err(EngineError::message)
    }

    /// Applies a remote or incremental yrs v1 update. It commits without a
    /// local origin, so this replica's undo manager never claims it. Errors on
    /// a malformed update.
    pub fn apply_update(&self, update: &[u8]) -> EngineResult<()> {
        self.engine.doc().apply_update_v1(update).map_err(EngineError::message)
    }

    pub fn apply_update_chunks(&self, chunks: &[Vec<u8>]) -> EngineResult<()> {
        self.engine
            .doc()
            .apply_update_chunks_v1(chunks)
            .map_err(EngineError::message)
    }

    /// [`EditSessionCore::apply_update`] plus an inference about where the peer
    /// that authored it is typing, for caret presence. Returns
    /// `{"clientId","story","paraId","endOffset"}` for the end of that peer's
    /// last text insertion, or `"null"` when the update carries work from more
    /// than one client, ends in a non-text insertion, or inserts nothing.
    /// Errors on a malformed update.
    pub fn apply_update_with_inference(&self, update: &[u8]) -> EngineResult<String> {
        let inference =
            apply_update_with_typing_inference(self.engine.doc(), update)
                .map_err(EngineError::message)?;
        serde_json::to_string(&inference.map(|inference| {
            json!({
                "clientId": inference.client_id,
                "story": inference.story,
                "paraId": inference.para_id,
                "endOffset": inference.end_offset,
            })
        }))
        .map_err(EngineError::message)
    }

    /// Applies an update produced by this document's dedicated local worker.
    /// The local origin lets the main replica's UndoManager retain ownership of
    /// the edit; remote/collaboration updates must use `apply_update` instead.
    pub fn apply_local_update(&self, update: &[u8]) -> EngineResult<()> {
        self.engine
            .doc()
            .apply_local_update_v1(update)
            .map_err(EngineError::message)
    }

    /// Applies an update another replica of this document committed for a host batch.
    /// It commits outside undo history and notifies as a local change.
    pub fn apply_host_update(&self, update: &[u8]) -> EngineResult<()> {
        self.undo.add_undo_barrier();
        let result = self
            .engine
            .doc()
            .apply_host_update_v1(update)
            .map_err(EngineError::message);
        self.undo.add_undo_barrier();
        result
    }

    /// Starts queueing committed transactions for
    /// [`EditSessionCore::drain_update_event`] instead of pushing them through a
    /// callback. Idempotent while observation is running.
    pub fn start_update_event_observation(&mut self) -> EngineResult<()> {
        if self.update_event_observer.is_some() {
            return Ok(());
        }
        let pending = Rc::new(RefCell::new(VecDeque::new()));
        let observed = Rc::clone(&pending);
        let subscription = self
            .engine
            .doc()
            .observe_update_v1(move |update, remote| {
                let mut encoded = Vec::with_capacity(update.len() + 1);
                encoded.push(u8::from(remote));
                encoded.extend_from_slice(update);
                observed.borrow_mut().push_back(encoded);
            })
            .map_err(EngineError::message)?;
        self.update_event_observer = Some(UpdateEventObserver {
            pending,
            _subscription: subscription,
        });
        Ok(())
    }

    /// Pops the oldest queued transaction, in arrival order. Byte 0 is `1`
    /// when the transaction had no local origin and `0` otherwise; the rest is
    /// the v1 update. Empty when the queue is drained or observation never
    /// started.
    pub fn drain_update_event(&self) -> Vec<u8> {
        self.update_event_observer
            .as_ref()
            .and_then(|observer| observer.pending.borrow_mut().pop_front())
            .unwrap_or_default()
    }

    /// Stops queueing and discards anything not yet drained.
    pub fn clear_update_event_observation(&mut self) {
        self.update_event_observer = None;
    }
}
