use wasm_bindgen::prelude::*;

use crate::DeckSession;
use crate::engine::{self, EngineError};

#[wasm_bindgen]
extern "C" {
    #[wasm_bindgen(js_namespace = performance, js_name = now)]
    fn performance_now() -> f64;
}

#[wasm_bindgen]
pub struct PptxDocument {
    inner: engine::Document,
}

const HOST: engine::Host = engine::Host {
    entropy: js_entropy,
    timer_ms: performance_now,
};

#[wasm_bindgen]
impl PptxDocument {
    #[doc(hidden)]
    #[wasm_bindgen(js_name = openReplayBaseline)]
    pub fn open_replay_baseline(
        source: &[u8],
        client_id: Option<f64>,
        initial_update: Option<Vec<u8>>,
    ) -> Result<PptxDocument, JsValue> {
        engine::Document::open_replay_baseline(source, client_id, initial_update, HOST)
            .map(|inner| Self { inner })
            .map_err(to_js)
    }

    #[doc(hidden)]
    #[wasm_bindgen(js_name = peerHydrationJson)]
    pub fn peer_hydration_json(&self, fonts_json: &str) -> Result<String, JsValue> {
        self.inner.peer_hydration_json(fonts_json).map_err(to_js)
    }

    #[doc(hidden)]
    #[wasm_bindgen(js_name = openPeerDeckJson)]
    pub fn open_peer_deck_json(
        source: &[u8],
        identity_json: &str,
        initial_update: Option<Vec<u8>>,
    ) -> Result<PptxDocument, JsValue> {
        engine::Document::open_peer_deck_json(source, identity_json, initial_update, HOST)
            .map(|inner| Self { inner })
            .map_err(to_js)
    }

    #[doc(hidden)]
    #[wasm_bindgen(js_name = registerPeerFontsJson)]
    pub fn register_peer_fonts_json(&self, fonts_json: &str) -> Result<(), JsValue> {
        self.inner
            .register_peer_fonts_json(fonts_json)
            .map_err(to_js)
    }

    #[doc(hidden)]
    #[wasm_bindgen(js_name = adoptPeerIdentity)]
    pub fn adopt_peer_identity(&self) -> Result<(), JsValue> {
        self.inner.adopt_peer_identity().map_err(to_js)
    }

    #[doc(hidden)]
    #[wasm_bindgen(js_name = replayJson)]
    pub fn replay_json(&self, envelope: &str) -> Result<String, JsValue> {
        self.inner.replay_json(envelope).map_err(to_js)
    }

    #[wasm_bindgen(js_name = proposeJson)]
    pub fn propose_json(&self, args: &str) -> Result<String, JsValue> {
        self.inner.propose_json(args).map_err(to_js)
    }

    #[wasm_bindgen(js_name = listProposalsJson)]
    pub fn list_proposals_json(&self) -> Result<String, JsValue> {
        self.inner.list_proposals_json().map_err(to_js)
    }

    #[wasm_bindgen(js_name = previewProposalJson)]
    pub fn preview_proposal_json(&self, args: &str) -> Result<String, JsValue> {
        self.inner.preview_proposal_json(args).map_err(to_js)
    }

    #[wasm_bindgen(js_name = acceptProposalJson)]
    pub fn accept_proposal_json(&self, args: &str) -> Result<String, JsValue> {
        self.inner.accept_proposal_json(args).map_err(to_js)
    }

    #[wasm_bindgen(js_name = rejectProposalJson)]
    pub fn reject_proposal_json(&self, args: &str) -> Result<String, JsValue> {
        self.inner.reject_proposal_json(args).map_err(to_js)
    }

    /// Opens a session that can exchange updates. It is not marked collaborative until
    /// `setCollaborative` is called.
    #[wasm_bindgen(js_name = openCollaborative)]
    pub fn open_collaborative(bytes: &[u8], client_id: f64) -> Result<PptxDocument, JsValue> {
        engine::Document::open_collaborative(bytes, client_id, HOST)
            .map(|inner| Self { inner })
            .map_err(to_js)
    }

    /// `source` is the file the update was seeded from; when it matches the
    /// recorded fingerprint the session keeps its part bytes and can save, and
    /// failing to reattach it throws.
    /// Any other bytes fall back to the bare update session, whose `saveBytes`
    /// fails — joining a room must not depend on carrying the right file.
    #[wasm_bindgen(js_name = openCollaborativeFromUpdate)]
    pub fn open_collaborative_from_update(
        update: &[u8],
        client_id: f64,
        source: Option<Vec<u8>>,
    ) -> Result<PptxDocument, JsValue> {
        engine::Document::open_collaborative_from_update(update, client_id, source, HOST)
            .map(|inner| Self { inner })
            .map_err(to_js)
    }

    #[wasm_bindgen(getter, js_name = clientId)]
    pub fn client_id(&self) -> f64 {
        self.inner.client_id()
    }

    #[wasm_bindgen(js_name = snapshotJson)]
    pub fn snapshot_json(&self) -> Result<String, JsValue> {
        self.inner.snapshot_json().map_err(to_js)
    }

    #[wasm_bindgen(js_name = sessionMetadataJson)]
    pub fn session_metadata_json(&self) -> Result<String, JsValue> {
        self.inner.session_metadata_json().map_err(to_js)
    }

    // Version-checked host edits. Requests and results are JSON; policy outcomes come back as
    // `{"ok":true,...}` or `{"ok":false,"version","failure":{"code","message","stepIndex"?,
    // "conflictingStepIndex"?,"target"?}}`, while malformed requests throw.

    /// The session-scoped version token of the committed deck state.
    #[wasm_bindgen(js_name = documentVersion)]
    pub fn document_version(&self) -> String {
        self.inner.document_version()
    }

    /// `{"slideIds"?}` -> `{"ok":true,"version","slides","stories"}`.
    #[wasm_bindgen(js_name = readContentJson)]
    pub fn read_content_json(&self, request: &str) -> Result<String, JsValue> {
        self.inner.read_content_json(request).map_err(to_js)
    }

    #[wasm_bindgen(js_name = readCommandContentJson)]
    pub fn read_command_content_json(&self, request: &str) -> Result<String, JsValue> {
        self.inner.read_command_content_json(request).map_err(to_js)
    }

    /// `{"text","within"?,"limit"?}` -> `{"ok":true,"version","matches","truncated"}`.
    #[wasm_bindgen(js_name = findTextJson)]
    pub fn find_text_json(&self, request: &str) -> Result<String, JsValue> {
        self.inner.find_text_json(request).map_err(to_js)
    }

    /// Runs every check of `applyEditsJson`, staging included, without changing anything:
    /// `{"ok":true,"baseVersion","wouldApply","previews"}`.
    #[wasm_bindgen(js_name = validateEditsJson)]
    pub fn validate_edits_json(&self, request: &str) -> Result<String, JsValue> {
        self.inner.validate_edits_json(request).map_err(to_js)
    }

    /// Applies a batch all-or-nothing: `{"ok":true,"baseVersion","version","applied","source",
    /// "changedSlides","changedStories","receipts"}`.
    #[wasm_bindgen(js_name = applyEditsJson)]
    pub fn apply_edits_json(&self, request: &str) -> Result<String, JsValue> {
        self.inner.apply_edits_json(request).map_err(to_js)
    }

    /// Structured export of the committed deck: `{"includeHiddenSlides"?,"includeHiddenShapes"?,
    /// "includeNotes"?,"includeComments"?,"includeFormatting"?,"maxBlocks"?,"maxBytes"?}` ->
    /// `{"ok":true,"version","content"}` or `{"ok":false,"version","failure"}`. Nothing changes.
    #[wasm_bindgen(js_name = exportStructuredJson)]
    pub fn export_structured_json(&self, options: &str) -> Result<String, JsValue> {
        self.inner.export_structured_json(options).map_err(to_js)
    }

    /// `exportStructuredJson` rendered as Markdown from the same read:
    /// `{"ok":true,"version","content":{"markdown","anchors","diagnostics","truncated"}}`.
    #[wasm_bindgen(js_name = exportMarkdownJson)]
    pub fn export_markdown_json(&self, options: &str) -> Result<String, JsValue> {
        self.inner.export_markdown_json(options).map_err(to_js)
    }

    #[wasm_bindgen(js_name = searchTextJson)]
    pub fn search_text_json(&self, args: &str) -> Result<String, JsValue> {
        self.inner.search_text_json(args).map_err(to_js)
    }

    #[wasm_bindgen(js_name = storyJson)]
    pub fn story_json(&self, args: &str) -> Result<String, JsValue> {
        self.inner.story_json(args).map_err(to_js)
    }

    #[wasm_bindgen(js_name = storyFormattingJson)]
    pub fn story_formatting_json(&self, args: &str) -> Result<String, JsValue> {
        self.inner.story_formatting_json(args).map_err(to_js)
    }

    /// Anchors a caret offset so that later edits, undo and remote updates move it.
    #[wasm_bindgen(js_name = anchorCaretJson)]
    pub fn anchor_caret_json(&self, args: &str) -> Result<String, JsValue> {
        self.inner.anchor_caret_json(args).map_err(to_js)
    }

    /// The current offset of an anchor from `anchorCaretJson`, or `null` once its story is gone.
    #[wasm_bindgen(js_name = resolveCaretAnchorJson)]
    pub fn resolve_caret_anchor_json(&self, args: &str) -> Result<String, JsValue> {
        self.inner.resolve_caret_anchor_json(args).map_err(to_js)
    }

    #[wasm_bindgen(js_name = mediaBytes)]
    pub fn media_bytes(&self, part_path: &str) -> Result<Vec<u8>, JsValue> {
        self.inner.media_bytes(part_path).map_err(to_js)
    }

    #[wasm_bindgen(js_name = metafileSvg)]
    pub fn metafile_svg(&self, part_path: &str) -> Option<Vec<u8>> {
        self.session()
            .metafile_svg(part_path)
            .map(|svg| svg.to_vec())
    }

    /// Serializes the deck back to `.pptx` bytes, edits included.
    #[wasm_bindgen(js_name = saveBytes)]
    pub fn save_bytes(&self) -> Result<Vec<u8>, JsValue> {
        self.inner.save_bytes().map_err(to_js)
    }

    #[wasm_bindgen(js_name = encodeStateVector)]
    pub fn encode_state_vector(&self) -> Vec<u8> {
        self.inner.encode_state_vector()
    }

    #[wasm_bindgen(js_name = encodeStateAsUpdate)]
    pub fn encode_state_as_update(&self) -> Vec<u8> {
        self.inner.encode_state_as_update()
    }

    #[wasm_bindgen(js_name = encodeDiff)]
    pub fn encode_diff(&self, remote_state_vector: &[u8]) -> Result<Vec<u8>, JsValue> {
        self.inner.encode_diff(remote_state_vector).map_err(to_js)
    }

    #[wasm_bindgen(js_name = setCollaborative)]
    pub fn set_collaborative(&self) -> Result<(), JsValue> {
        self.inner.set_collaborative().map_err(to_js)
    }

    #[wasm_bindgen(js_name = applyUpdateJson)]
    pub fn apply_update_json(&self, update: &[u8]) -> Result<String, JsValue> {
        self.inner.apply_update_json(update).map_err(to_js)
    }

    #[wasm_bindgen(js_name = startUpdateObservation)]
    pub fn start_update_observation(&mut self) -> Result<(), JsValue> {
        self.inner.start_update_observation().map_err(to_js)
    }

    #[wasm_bindgen(js_name = clearUpdateObservation)]
    pub fn clear_update_observation(&mut self) {
        self.inner.clear_update_observation()
    }

    #[wasm_bindgen(js_name = drainUpdateEvent)]
    pub fn drain_update_event(&self) -> Vec<u8> {
        self.inner.drain_update_event()
    }

    #[wasm_bindgen(js_name = insertTextJson)]
    pub fn insert_text_json(&self, args: &str) -> Result<String, JsValue> {
        self.inner.insert_text_json(args).map_err(to_js)
    }

    #[wasm_bindgen(js_name = insertTextProfiledJson)]
    pub fn insert_text_profiled_json(&self, args: &str) -> Result<String, JsValue> {
        self.inner.insert_text_profiled_json(args).map_err(to_js)
    }

    #[wasm_bindgen(js_name = deleteTextJson)]
    pub fn delete_text_json(&self, args: &str) -> Result<String, JsValue> {
        self.inner.delete_text_json(args).map_err(to_js)
    }

    #[wasm_bindgen(js_name = deleteTextProfiledJson)]
    pub fn delete_text_profiled_json(&self, args: &str) -> Result<String, JsValue> {
        self.inner.delete_text_profiled_json(args).map_err(to_js)
    }

    #[wasm_bindgen(js_name = formatTextJson)]
    pub fn format_text_json(&self, args: &str) -> Result<String, JsValue> {
        self.inner.format_text_json(args).map_err(to_js)
    }

    #[wasm_bindgen(js_name = setParagraphAlignmentJson)]
    pub fn set_paragraph_alignment_json(&self, args: &str) -> Result<String, JsValue> {
        self.inner.set_paragraph_alignment_json(args).map_err(to_js)
    }

    #[wasm_bindgen(js_name = addCommentJson)]
    pub fn add_comment_json(&self, args: &str) -> Result<String, JsValue> {
        self.inner.add_comment_json(args).map_err(to_js)
    }

    #[wasm_bindgen(js_name = replyToCommentJson)]
    pub fn reply_to_comment_json(&self, args: &str) -> Result<String, JsValue> {
        self.inner.reply_to_comment_json(args).map_err(to_js)
    }

    #[wasm_bindgen(js_name = setCommentPositionJson)]
    pub fn set_comment_position_json(&self, args: &str) -> Result<String, JsValue> {
        self.inner.set_comment_position_json(args).map_err(to_js)
    }

    #[wasm_bindgen(js_name = setCommentStatusJson)]
    pub fn set_comment_status_json(&self, args: &str) -> Result<String, JsValue> {
        self.inner.set_comment_status_json(args).map_err(to_js)
    }

    #[wasm_bindgen(js_name = removeCommentJson)]
    pub fn remove_comment_json(&self, args: &str) -> Result<String, JsValue> {
        self.inner.remove_comment_json(args).map_err(to_js)
    }

    #[wasm_bindgen(js_name = setCommentFlavorJson)]
    pub fn set_comment_flavor_json(&self, args: &str) -> Result<String, JsValue> {
        self.inner.set_comment_flavor_json(args).map_err(to_js)
    }

    #[wasm_bindgen(js_name = commentsJson)]
    pub fn comments_json(&self) -> Result<String, JsValue> {
        self.inner.comments_json().map_err(to_js)
    }

    #[wasm_bindgen(js_name = insertParagraphBreakJson)]
    pub fn insert_paragraph_break_json(&self, args: &str) -> Result<String, JsValue> {
        self.inner.insert_paragraph_break_json(args).map_err(to_js)
    }

    #[wasm_bindgen(js_name = setStoryParagraphsJson)]
    pub fn set_story_paragraphs_json(&self, args: &str) -> Result<String, JsValue> {
        self.inner.set_story_paragraphs_json(args).map_err(to_js)
    }

    #[wasm_bindgen(js_name = insertTableRowJson)]
    pub fn insert_table_row_json(&self, args: &str) -> Result<String, JsValue> {
        self.inner.insert_table_row_json(args).map_err(to_js)
    }

    #[wasm_bindgen(js_name = checkTableEditJson)]
    pub fn check_table_edit_json(&self, args: &str) -> Result<(), JsValue> {
        self.inner.check_table_edit_json(args).map_err(to_js)
    }

    #[wasm_bindgen(js_name = deleteTableRowJson)]
    pub fn delete_table_row_json(&self, args: &str) -> Result<String, JsValue> {
        self.inner.delete_table_row_json(args).map_err(to_js)
    }

    #[wasm_bindgen(js_name = insertTableColumnJson)]
    pub fn insert_table_column_json(&self, args: &str) -> Result<String, JsValue> {
        self.inner.insert_table_column_json(args).map_err(to_js)
    }

    #[wasm_bindgen(js_name = deleteTableColumnJson)]
    pub fn delete_table_column_json(&self, args: &str) -> Result<String, JsValue> {
        self.inner.delete_table_column_json(args).map_err(to_js)
    }

    #[wasm_bindgen(js_name = insertSlideJson)]
    pub fn insert_slide_json(&self, args: &str) -> Result<String, JsValue> {
        self.inner.insert_slide_json(args).map_err(to_js)
    }

    #[wasm_bindgen(js_name = duplicateSlideJson)]
    pub fn duplicate_slide_json(&self, args: &str) -> Result<String, JsValue> {
        self.inner.duplicate_slide_json(args).map_err(to_js)
    }

    #[wasm_bindgen(js_name = insertSlideProfiledJson)]
    pub fn insert_slide_profiled_json(&self, args: &str) -> Result<String, JsValue> {
        self.inner.insert_slide_profiled_json(args).map_err(to_js)
    }

    #[wasm_bindgen(js_name = deleteSlideJson)]
    pub fn delete_slide_json(&self, args: &str) -> Result<String, JsValue> {
        self.inner.delete_slide_json(args).map_err(to_js)
    }

    #[wasm_bindgen(js_name = setSlideLayoutJson)]
    pub fn set_slide_layout_json(&self, args: &str) -> Result<String, JsValue> {
        self.inner.set_slide_layout_json(args).map_err(to_js)
    }

    #[wasm_bindgen(js_name = setSlideNotesJson)]
    pub fn set_slide_notes_json(&self, args: &str) -> Result<String, JsValue> {
        self.inner.set_slide_notes_json(args).map_err(to_js)
    }

    #[wasm_bindgen(js_name = moveSlideJson)]
    pub fn move_slide_json(&self, args: &str) -> Result<String, JsValue> {
        self.inner.move_slide_json(args).map_err(to_js)
    }

    #[wasm_bindgen(js_name = addTextBoxJson)]
    pub fn add_text_box_json(&self, args: &str) -> Result<String, JsValue> {
        self.inner.add_text_box_json(args).map_err(to_js)
    }

    #[wasm_bindgen(js_name = addTextBoxProfiledJson)]
    pub fn add_text_box_profiled_json(&self, args: &str) -> Result<String, JsValue> {
        self.inner.add_text_box_profiled_json(args).map_err(to_js)
    }

    #[wasm_bindgen(js_name = addShapeJson)]
    pub fn add_shape_json(&self, args: &str) -> Result<String, JsValue> {
        self.inner.add_shape_json(args).map_err(to_js)
    }

    #[wasm_bindgen(js_name = addPictureJson)]
    pub fn add_picture_json(&self, args: &str) -> Result<String, JsValue> {
        self.inner.add_picture_json(args).map_err(to_js)
    }

    #[wasm_bindgen(js_name = removeShapeJson)]
    pub fn remove_shape_json(&self, args: &str) -> Result<String, JsValue> {
        self.inner.remove_shape_json(args).map_err(to_js)
    }

    #[wasm_bindgen(js_name = bringShapeToFrontJson)]
    pub fn bring_shape_to_front_json(&self, args: &str) -> Result<String, JsValue> {
        self.inner.bring_shape_to_front_json(args).map_err(to_js)
    }

    #[wasm_bindgen(js_name = sendShapeToBackJson)]
    pub fn send_shape_to_back_json(&self, args: &str) -> Result<String, JsValue> {
        self.inner.send_shape_to_back_json(args).map_err(to_js)
    }

    #[wasm_bindgen(js_name = bringShapeForwardJson)]
    pub fn bring_shape_forward_json(&self, args: &str) -> Result<String, JsValue> {
        self.inner.bring_shape_forward_json(args).map_err(to_js)
    }

    #[wasm_bindgen(js_name = sendShapeBackwardJson)]
    pub fn send_shape_backward_json(&self, args: &str) -> Result<String, JsValue> {
        self.inner.send_shape_backward_json(args).map_err(to_js)
    }

    #[wasm_bindgen(js_name = moveShapeJson)]
    pub fn move_shape_json(&self, args: &str) -> Result<String, JsValue> {
        self.inner.move_shape_json(args).map_err(to_js)
    }

    #[wasm_bindgen(js_name = moveShapeProfiledJson)]
    pub fn move_shape_profiled_json(&self, args: &str) -> Result<String, JsValue> {
        self.inner.move_shape_profiled_json(args).map_err(to_js)
    }

    #[wasm_bindgen(js_name = resizeShapeJson)]
    pub fn resize_shape_json(&self, args: &str) -> Result<String, JsValue> {
        self.inner.resize_shape_json(args).map_err(to_js)
    }

    #[wasm_bindgen(js_name = setShapeRectJson)]
    pub fn set_shape_rect_json(&self, args: &str) -> Result<String, JsValue> {
        self.inner.set_shape_rect_json(args).map_err(to_js)
    }

    #[wasm_bindgen(js_name = setShapeFillJson)]
    pub fn set_shape_fill_json(&self, args: &str) -> Result<String, JsValue> {
        self.inner.set_shape_fill_json(args).map_err(to_js)
    }

    #[wasm_bindgen(js_name = setShapeStrokeJson)]
    pub fn set_shape_stroke_json(&self, args: &str) -> Result<String, JsValue> {
        self.inner.set_shape_stroke_json(args).map_err(to_js)
    }

    #[wasm_bindgen(js_name = setShapeAdjustJson)]
    pub fn set_shape_adjust_json(&self, args: &str) -> Result<String, JsValue> {
        self.inner.set_shape_adjust_json(args).map_err(to_js)
    }

    #[wasm_bindgen(js_name = undoCaptureMode)]
    pub fn undo_capture_mode(&self) -> String {
        self.inner.undo_capture_mode()
    }

    #[wasm_bindgen(js_name = setUndoCaptureMode)]
    pub fn set_undo_capture_mode(&self, mode: &str) -> Result<(), JsValue> {
        self.inner.set_undo_capture_mode(mode).map_err(to_js)
    }

    #[wasm_bindgen(js_name = addUndoBoundary)]
    pub fn add_undo_boundary(&self) {
        self.inner.add_undo_boundary()
    }

    #[wasm_bindgen(js_name = undoJson)]
    pub fn undo_json(&self) -> Result<String, JsValue> {
        self.inner.undo_json().map_err(to_js)
    }

    #[wasm_bindgen(js_name = undoReceiptJson)]
    pub fn undo_receipt_json(&self) -> Result<String, JsValue> {
        self.inner.undo_receipt_json().map_err(to_js)
    }

    #[wasm_bindgen(js_name = redoReceiptJson)]
    pub fn redo_receipt_json(&self) -> Result<String, JsValue> {
        self.inner.redo_receipt_json().map_err(to_js)
    }

    /// `undoJson` timed at its undo, snapshot and serialize boundaries, as
    /// `{"receipt": ..., "profile": {"undoMs", "snapshotMs", "serializeMs"}}`.
    #[wasm_bindgen(js_name = undoProfiledJson)]
    pub fn undo_profiled_json(&self) -> Result<String, JsValue> {
        self.inner.undo_profiled_json().map_err(to_js)
    }

    #[wasm_bindgen(js_name = redoJson)]
    pub fn redo_json(&self) -> Result<String, JsValue> {
        self.inner.redo_json().map_err(to_js)
    }

    #[wasm_bindgen(js_name = canUndo)]
    pub fn can_undo(&self) -> bool {
        self.inner.can_undo()
    }

    #[wasm_bindgen(js_name = canRedo)]
    pub fn can_redo(&self) -> bool {
        self.inner.can_redo()
    }

    pub fn version() -> String {
        engine::version().to_owned()
    }
}

impl PptxDocument {
    pub fn session(&self) -> &DeckSession {
        self.inner.session()
    }
}

fn to_js(error: EngineError) -> JsValue {
    match error.kind {
        engine::ErrorKind::Message => JsValue::from_str(&error.message),
        engine::ErrorKind::Peer { code } => {
            let exception = js_sys::Error::new(&error.message);
            exception.set_name("PresentationPeerError");
            let _ = js_sys::Reflect::set(
                exception.as_ref(),
                &JsValue::from_str("code"),
                &JsValue::from_str(&code),
            );
            exception.into()
        }
        engine::ErrorKind::Decode => {
            let exception = js_sys::Error::new(&error.message);
            exception.set_name("PptxEditDecodeError");
            exception.into()
        }
    }
}

/// Host randomness for version nonces; the wasm target has no ambient entropy source.
fn js_entropy() -> u64 {
    #[cfg(target_arch = "wasm32")]
    {
        ((js_sys::Math::random() * 9_007_199_254_740_992.0) as u64)
            ^ (js_sys::Date::now() as u64).rotate_left(21)
    }
    #[cfg(not(target_arch = "wasm32"))]
    {
        0
    }
}
