use std::collections::{BTreeMap, VecDeque};
use std::sync::{Arc, Mutex};

use base64::Engine as _;
use serde::{Deserialize, Serialize};
use wasm_bindgen::prelude::*;
use yrs::Subscription;

use crate::structured::{PptxExportOptions, export_outcome_json};
use crate::{
    CommentFlavor, DeckSession, DeckSnapshot, EditCtx, EditRequest, FindRequest, MAX_REQUEST_BYTES,
    PictureDraft, PresetShapeDraft, ReadRequest, ShapeDraft, ShapeReceipt, ShapeRect, ShapeStroke,
    SlideReceipt, TextReceipt, TextStyle, TextStylePatch, TransformReceipt, UpdateEvent,
    UpdateOrigin, outcome_json, oversized_request,
};

#[wasm_bindgen]
extern "C" {
    #[wasm_bindgen(js_namespace = performance, js_name = now)]
    fn performance_now() -> f64;
}

#[wasm_bindgen]
pub struct PptxDocument {
    session: DeckSession,
    update_observer: Option<UpdateObserver>,
}

struct UpdateObserver {
    pending: Arc<Mutex<VecDeque<UpdateEvent>>>,
    _subscription: Subscription,
}

#[derive(Deserialize)]
#[serde(rename_all = "camelCase")]
struct SearchTextArgs {
    query: String,
    #[serde(default)]
    case_sensitive: bool,
    limit: Option<u32>,
}

#[derive(Deserialize)]
#[serde(rename_all = "camelCase")]
struct StoryArgs {
    story_id: String,
}

#[derive(Deserialize)]
#[serde(rename_all = "camelCase")]
struct InsertTextArgs {
    story_id: String,
    index: u32,
    text: String,
    #[serde(default)]
    style: TextStyle,
}

#[derive(Deserialize)]
#[serde(rename_all = "camelCase")]
struct DeleteTextArgs {
    story_id: String,
    start: u32,
    end: u32,
}

#[derive(Deserialize)]
#[serde(rename_all = "camelCase")]
struct FormatTextArgs {
    story_id: String,
    start: u32,
    end: u32,
    #[serde(default)]
    patch: TextStylePatch,
}

#[derive(Deserialize)]
#[serde(rename_all = "camelCase")]
struct SetParagraphAlignmentArgs {
    story_id: String,
    start: u32,
    end: u32,
    #[serde(default)]
    alignment: Option<String>,
}

#[derive(Deserialize)]
#[serde(rename_all = "camelCase")]
struct AddCommentArgs {
    slide_id: String,
    author: String,
    #[serde(default)]
    initials: String,
    text: String,
    created: String,
    #[serde(default)]
    x_emu: i64,
    #[serde(default)]
    y_emu: i64,
}

#[derive(Deserialize)]
#[serde(rename_all = "camelCase")]
struct ReplyCommentArgs {
    comment_id: String,
    author: String,
    #[serde(default)]
    initials: String,
    text: String,
    created: String,
}

#[derive(Deserialize)]
#[serde(rename_all = "camelCase")]
struct CommentPositionArgs {
    comment_id: String,
    x_emu: i64,
    y_emu: i64,
}

#[derive(Deserialize)]
#[serde(rename_all = "camelCase")]
struct CommentStatusArgs {
    comment_id: String,
    resolved: bool,
}

#[derive(Deserialize)]
#[serde(rename_all = "camelCase")]
struct CommentIdArgs {
    comment_id: String,
}

#[derive(Deserialize)]
#[serde(rename_all = "camelCase")]
struct CommentFlavorArgs {
    flavor: CommentFlavor,
}

#[derive(Deserialize)]
#[serde(rename_all = "camelCase")]
struct ParagraphBreakArgs {
    story_id: String,
    index: u32,
}

#[derive(Deserialize)]
#[serde(rename_all = "camelCase")]
struct InsertSlideArgs {
    index: u32,
    layout_part_path: Option<String>,
}

#[derive(Deserialize)]
#[serde(rename_all = "camelCase")]
struct SlideArgs {
    slide_id: String,
}

#[derive(Deserialize)]
#[serde(rename_all = "camelCase")]
struct SetSlideNotesArgs {
    slide_id: String,
    text: String,
}

#[derive(Deserialize)]
#[serde(rename_all = "camelCase")]
struct MoveSlideArgs {
    slide_id: String,
    to_index: u32,
}

#[derive(Deserialize)]
#[serde(rename_all = "camelCase")]
struct AddTextBoxArgs {
    slide_id: String,
    draft: ShapeDraft,
}

#[derive(Deserialize)]
#[serde(rename_all = "camelCase")]
struct AddShapeArgs {
    slide_id: String,
    draft: PresetShapeDraft,
}

#[derive(Deserialize)]
#[serde(rename_all = "camelCase")]
struct AddPictureArgs {
    slide_id: String,
    name: String,
    rect: ShapeRect,
    content_type: String,
    /// The image bytes, base64-encoded: JSON has no binary payload of its own.
    media_base64: String,
}

#[derive(Deserialize)]
#[serde(rename_all = "camelCase")]
struct ShapeArgs {
    slide_id: String,
    shape_id: String,
}

#[derive(Deserialize)]
#[serde(rename_all = "camelCase")]
struct MoveShapeArgs {
    slide_id: String,
    shape_id: String,
    x: i64,
    y: i64,
}

#[derive(Deserialize)]
#[serde(rename_all = "camelCase")]
struct ResizeShapeArgs {
    slide_id: String,
    shape_id: String,
    width: i64,
    height: i64,
}

#[derive(Deserialize)]
#[serde(rename_all = "camelCase")]
struct SetShapeRectArgs {
    slide_id: String,
    shape_id: String,
    rect: ShapeRect,
}

#[derive(Deserialize)]
#[serde(rename_all = "camelCase")]
struct SetShapeFillArgs {
    slide_id: String,
    shape_id: String,
    color: Option<String>,
}

#[derive(Deserialize)]
#[serde(rename_all = "camelCase")]
struct SetShapeStrokeArgs {
    slide_id: String,
    shape_id: String,
    #[serde(default)]
    stroke: ShapeStroke,
}

#[derive(Deserialize)]
#[serde(rename_all = "camelCase")]
struct SetShapeAdjustArgs {
    slide_id: String,
    shape_id: String,
    adjustments: BTreeMap<String, f64>,
}

#[derive(Serialize)]
#[serde(rename_all = "camelCase")]
struct HistoryResult {
    applied: bool,
    snapshot: DeckSnapshot,
}

/// boundary stage latencies of one profiled edit, in milliseconds.
#[derive(Serialize)]
#[serde(rename_all = "camelCase")]
struct EditProfile {
    parse_ms: f64,
    apply_ms: f64,
    serialize_ms: f64,
}

#[derive(Serialize)]
#[serde(rename_all = "camelCase")]
struct HistoryProfile {
    undo_ms: f64,
    snapshot_ms: f64,
    serialize_ms: f64,
}

#[derive(Deserialize)]
struct ProposalIdArgs {
    id: String,
    #[serde(default)]
    force: bool,
}

#[wasm_bindgen]
impl PptxDocument {
    #[wasm_bindgen(js_name = proposeJson)]
    pub fn propose_json(&self, args: &str) -> Result<String, JsValue> {
        json(
            self.session
                .propose(parse_args(args)?)
                .map_err(proposal_error)?,
        )
    }

    #[wasm_bindgen(js_name = listProposalsJson)]
    pub fn list_proposals_json(&self) -> Result<String, JsValue> {
        json(self.session.proposals().map_err(proposal_error)?)
    }

    #[wasm_bindgen(js_name = previewProposalJson)]
    pub fn preview_proposal_json(&self, args: &str) -> Result<String, JsValue> {
        let args: ProposalIdArgs = parse_args(args)?;
        json(
            self.session
                .preview_proposal(&args.id)
                .map_err(proposal_error)?,
        )
    }

    #[wasm_bindgen(js_name = acceptProposalJson)]
    pub fn accept_proposal_json(&self, args: &str) -> Result<String, JsValue> {
        let args: ProposalIdArgs = parse_args(args)?;
        json(
            self.session
                .accept_proposal(&args.id, args.force)
                .map_err(proposal_error)?,
        )
    }

    #[wasm_bindgen(js_name = rejectProposalJson)]
    pub fn reject_proposal_json(&self, args: &str) -> Result<String, JsValue> {
        let args: ProposalIdArgs = parse_args(args)?;
        json(self.session.reject_proposal(&args.id))
    }

    #[wasm_bindgen(js_name = openCollaborative)]
    pub fn open_collaborative(bytes: &[u8], client_id: f64) -> Result<PptxDocument, JsValue> {
        let client_id = parse_client_id(client_id)?;
        DeckSession::open(bytes, client_id)
            .map(Self::opened)
            .map_err(js_error)
    }

    /// `source` is the file the update was seeded from; when it matches the
    /// recorded fingerprint the session keeps its part bytes and can save.
    /// Any other bytes fall back to the bare update session, whose `saveBytes`
    /// fails — joining a room must not depend on carrying the right file.
    #[wasm_bindgen(js_name = openCollaborativeFromUpdate)]
    pub fn open_collaborative_from_update(
        update: &[u8],
        client_id: f64,
        source: Option<Vec<u8>>,
    ) -> Result<PptxDocument, JsValue> {
        let client_id = parse_client_id(client_id)?;
        let session = source
            .and_then(|source| {
                DeckSession::open_from_update_with_source(update, &source, client_id).ok()
            })
            .map_or_else(|| DeckSession::open_from_update(update, client_id), Ok)
            .map_err(js_error)?;
        Ok(Self::opened(session))
    }

    #[wasm_bindgen(getter, js_name = clientId)]
    pub fn client_id(&self) -> f64 {
        self.session.client_id() as f64
    }

    #[wasm_bindgen(js_name = snapshotJson)]
    pub fn snapshot_json(&self) -> Result<String, JsValue> {
        json(self.session.snapshot().map_err(js_error)?)
    }

    // Version-checked host edits. Requests and results are JSON; policy outcomes come back as
    // `{"ok":true,...}` or `{"ok":false,"version","failure":{"code","message","stepIndex"?,
    // "conflictingStepIndex"?,"target"?}}`, while malformed requests throw.

    /// The session-scoped version token of the committed deck state.
    #[wasm_bindgen(js_name = documentVersion)]
    pub fn document_version(&self) -> String {
        self.session.version().to_string()
    }

    /// `{"slideIds"?}` -> `{"ok":true,"version","slides","stories"}`.
    #[wasm_bindgen(js_name = readContentJson)]
    pub fn read_content_json(&self, request: &str) -> Result<String, JsValue> {
        if request.len() > MAX_REQUEST_BYTES {
            return self.oversized();
        }
        let request: ReadRequest = parse_args(request)?;
        outcome_json(&self.session.read_content(&request).map_err(js_error)?).map_err(js_error)
    }

    /// `{"text","within"?,"limit"?}` -> `{"ok":true,"version","matches","truncated"}`.
    #[wasm_bindgen(js_name = findTextJson)]
    pub fn find_text_json(&self, request: &str) -> Result<String, JsValue> {
        if request.len() > MAX_REQUEST_BYTES {
            return self.oversized();
        }
        let request: FindRequest = parse_args(request)?;
        outcome_json(&self.session.find_text(&request).map_err(js_error)?).map_err(js_error)
    }

    /// Runs every check of `applyEditsJson`, staging included, without changing anything:
    /// `{"ok":true,"baseVersion","wouldApply","previews"}`.
    #[wasm_bindgen(js_name = validateEditsJson)]
    pub fn validate_edits_json(&self, request: &str) -> Result<String, JsValue> {
        if request.len() > MAX_REQUEST_BYTES {
            return self.oversized();
        }
        let request: EditRequest = parse_args(request)?;
        outcome_json(&self.session.validate_edits(&request).map_err(js_error)?).map_err(js_error)
    }

    /// Applies a batch all-or-nothing: `{"ok":true,"baseVersion","version","applied","source",
    /// "changedSlides","changedStories","receipts"}`.
    #[wasm_bindgen(js_name = applyEditsJson)]
    pub fn apply_edits_json(&self, request: &str) -> Result<String, JsValue> {
        if request.len() > MAX_REQUEST_BYTES {
            return self.oversized();
        }
        let request: EditRequest = parse_args(request)?;
        outcome_json(&self.session.apply_edits(&request).map_err(js_error)?).map_err(js_error)
    }

    /// Structured export of the committed deck: `{"includeHiddenSlides"?,"includeHiddenShapes"?,
    /// "includeNotes"?,"includeComments"?,"includeFormatting"?,"maxBlocks"?,"maxBytes"?}` ->
    /// `{"ok":true,"version","content"}` or `{"ok":false,"version","failure"}`. Nothing changes.
    #[wasm_bindgen(js_name = exportStructuredJson)]
    pub fn export_structured_json(&self, options: &str) -> Result<String, JsValue> {
        let options: PptxExportOptions = parse_args(options)?;
        export_outcome_json(&self.session.export_structured(&options).map_err(js_error)?)
            .map_err(js_error)
    }

    /// `exportStructuredJson` rendered as Markdown from the same read:
    /// `{"ok":true,"version","content":{"markdown","anchors","diagnostics","truncated"}}`.
    #[wasm_bindgen(js_name = exportMarkdownJson)]
    pub fn export_markdown_json(&self, options: &str) -> Result<String, JsValue> {
        let options: PptxExportOptions = parse_args(options)?;
        export_outcome_json(&self.session.export_markdown(&options).map_err(js_error)?)
            .map_err(js_error)
    }

    #[wasm_bindgen(js_name = searchTextJson)]
    pub fn search_text_json(&self, args: &str) -> Result<String, JsValue> {
        let args: SearchTextArgs = parse_args(args)?;
        json(
            self.session
                .search_text(
                    &args.query,
                    args.case_sensitive,
                    args.limit.map(|value| value as usize),
                )
                .map_err(js_error)?,
        )
    }

    #[wasm_bindgen(js_name = storyJson)]
    pub fn story_json(&self, args: &str) -> Result<String, JsValue> {
        let args: StoryArgs = parse_args(args)?;
        json(self.session.story(&args.story_id).map_err(js_error)?)
    }

    #[wasm_bindgen(js_name = mediaBytes)]
    pub fn media_bytes(&self, part_path: &str) -> Result<Vec<u8>, JsValue> {
        self.session.media_bytes(part_path).map_err(js_error)
    }

    /// Serializes the deck back to `.pptx` bytes, edits included.
    #[wasm_bindgen(js_name = saveBytes)]
    pub fn save_bytes(&self) -> Result<Vec<u8>, JsValue> {
        self.session.save().map_err(js_error)
    }

    #[wasm_bindgen(js_name = encodeStateVector)]
    pub fn encode_state_vector(&self) -> Vec<u8> {
        self.session.encode_state_vector_v1()
    }

    #[wasm_bindgen(js_name = encodeStateAsUpdate)]
    pub fn encode_state_as_update(&self) -> Vec<u8> {
        self.session.encode_state_as_update_v1()
    }

    #[wasm_bindgen(js_name = encodeDiff)]
    pub fn encode_diff(&self, remote_state_vector: &[u8]) -> Result<Vec<u8>, JsValue> {
        self.session
            .encode_diff_v1(remote_state_vector)
            .map_err(js_error)
    }

    #[wasm_bindgen(js_name = applyUpdateJson)]
    pub fn apply_update_json(&self, update: &[u8]) -> Result<String, JsValue> {
        json(self.session.apply_update_v1(update).map_err(js_error)?)
    }

    #[wasm_bindgen(js_name = startUpdateObservation)]
    pub fn start_update_observation(&mut self) -> Result<(), JsValue> {
        if self.update_observer.is_some() {
            return Ok(());
        }
        let pending = Arc::new(Mutex::new(VecDeque::new()));
        let observed = Arc::clone(&pending);
        let subscription = self
            .session
            .observe_update_v1(move |event| {
                observed
                    .lock()
                    .unwrap_or_else(|poisoned| poisoned.into_inner())
                    .push_back(event);
            })
            .map_err(js_error)?;
        self.update_observer = Some(UpdateObserver {
            pending,
            _subscription: subscription,
        });
        Ok(())
    }

    #[wasm_bindgen(js_name = clearUpdateObservation)]
    pub fn clear_update_observation(&mut self) {
        self.update_observer = None;
    }

    #[wasm_bindgen(js_name = drainUpdateEvent)]
    pub fn drain_update_event(&self) -> Vec<u8> {
        let Some(observer) = &self.update_observer else {
            return Vec::new();
        };
        let event = observer
            .pending
            .lock()
            .unwrap_or_else(|poisoned| poisoned.into_inner())
            .pop_front();
        let Some(event) = event else {
            return Vec::new();
        };
        let mut encoded = Vec::with_capacity(event.update.len() + 1);
        encoded.push(match event.origin {
            UpdateOrigin::Local => 0,
            UpdateOrigin::Remote => 1,
        });
        encoded.extend_from_slice(&event.update);
        encoded
    }

    #[wasm_bindgen(js_name = insertTextJson)]
    pub fn insert_text_json(&self, args: &str) -> Result<String, JsValue> {
        json(self.insert_text(parse_args(args)?)?)
    }

    #[wasm_bindgen(js_name = insertTextProfiledJson)]
    pub fn insert_text_profiled_json(&self, args: &str) -> Result<String, JsValue> {
        profiled(args, &mut performance_now, |args| self.insert_text(args))
    }

    #[wasm_bindgen(js_name = deleteTextJson)]
    pub fn delete_text_json(&self, args: &str) -> Result<String, JsValue> {
        json(self.delete_text(parse_args(args)?)?)
    }

    #[wasm_bindgen(js_name = deleteTextProfiledJson)]
    pub fn delete_text_profiled_json(&self, args: &str) -> Result<String, JsValue> {
        profiled(args, &mut performance_now, |args| self.delete_text(args))
    }

    #[wasm_bindgen(js_name = formatTextJson)]
    pub fn format_text_json(&self, args: &str) -> Result<String, JsValue> {
        let args: FormatTextArgs = parse_args(args)?;
        json(
            self.session
                .format_text(
                    &local_context(),
                    &args.story_id,
                    args.start,
                    args.end,
                    &args.patch,
                )
                .map_err(js_error)?,
        )
    }

    #[wasm_bindgen(js_name = setParagraphAlignmentJson)]
    pub fn set_paragraph_alignment_json(&self, args: &str) -> Result<String, JsValue> {
        let args: SetParagraphAlignmentArgs = parse_args(args)?;
        json(
            self.session
                .set_paragraph_alignment(
                    &local_context(),
                    &args.story_id,
                    args.start,
                    args.end,
                    args.alignment.as_deref(),
                )
                .map_err(js_error)?,
        )
    }

    #[wasm_bindgen(js_name = addCommentJson)]
    pub fn add_comment_json(&self, args: &str) -> Result<String, JsValue> {
        let args: AddCommentArgs = parse_args(args)?;
        json(
            self.session
                .add_comment(
                    &local_context(),
                    &args.slide_id,
                    &args.author,
                    &args.initials,
                    &args.text,
                    &args.created,
                    args.x_emu,
                    args.y_emu,
                )
                .map_err(js_error)?,
        )
    }

    #[wasm_bindgen(js_name = replyToCommentJson)]
    pub fn reply_to_comment_json(&self, args: &str) -> Result<String, JsValue> {
        let args: ReplyCommentArgs = parse_args(args)?;
        json(
            self.session
                .reply_to_comment(
                    &local_context(),
                    &args.comment_id,
                    &args.author,
                    &args.initials,
                    &args.text,
                    &args.created,
                )
                .map_err(js_error)?,
        )
    }

    #[wasm_bindgen(js_name = setCommentPositionJson)]
    pub fn set_comment_position_json(&self, args: &str) -> Result<String, JsValue> {
        let args: CommentPositionArgs = parse_args(args)?;
        json(
            self.session
                .set_comment_position(&local_context(), &args.comment_id, args.x_emu, args.y_emu)
                .map_err(js_error)?,
        )
    }

    #[wasm_bindgen(js_name = setCommentStatusJson)]
    pub fn set_comment_status_json(&self, args: &str) -> Result<String, JsValue> {
        let args: CommentStatusArgs = parse_args(args)?;
        json(
            self.session
                .set_comment_status(&local_context(), &args.comment_id, args.resolved)
                .map_err(js_error)?,
        )
    }

    #[wasm_bindgen(js_name = removeCommentJson)]
    pub fn remove_comment_json(&self, args: &str) -> Result<String, JsValue> {
        let args: CommentIdArgs = parse_args(args)?;
        json(
            self.session
                .remove_comment(&local_context(), &args.comment_id)
                .map_err(js_error)?,
        )
    }

    #[wasm_bindgen(js_name = setCommentFlavorJson)]
    pub fn set_comment_flavor_json(&self, args: &str) -> Result<String, JsValue> {
        let args: CommentFlavorArgs = parse_args(args)?;
        json(
            self.session
                .set_comment_flavor(&local_context(), args.flavor)
                .map_err(js_error)?,
        )
    }

    #[wasm_bindgen(js_name = commentsJson)]
    pub fn comments_json(&self) -> Result<String, JsValue> {
        json(self.session.comments().map_err(js_error)?)
    }

    #[wasm_bindgen(js_name = insertParagraphBreakJson)]
    pub fn insert_paragraph_break_json(&self, args: &str) -> Result<String, JsValue> {
        let args: ParagraphBreakArgs = parse_args(args)?;
        json(
            self.session
                .insert_paragraph_break(&local_context(), &args.story_id, args.index)
                .map_err(js_error)?,
        )
    }

    #[wasm_bindgen(js_name = insertSlideJson)]
    pub fn insert_slide_json(&self, args: &str) -> Result<String, JsValue> {
        json(self.insert_slide(parse_args(args)?)?)
    }

    #[wasm_bindgen(js_name = insertSlideProfiledJson)]
    pub fn insert_slide_profiled_json(&self, args: &str) -> Result<String, JsValue> {
        profiled(args, &mut performance_now, |args| self.insert_slide(args))
    }

    #[wasm_bindgen(js_name = deleteSlideJson)]
    pub fn delete_slide_json(&self, args: &str) -> Result<String, JsValue> {
        let args: SlideArgs = parse_args(args)?;
        json(
            self.session
                .delete_slide(&local_context(), &args.slide_id)
                .map_err(js_error)?,
        )
    }

    #[wasm_bindgen(js_name = setSlideNotesJson)]
    pub fn set_slide_notes_json(&self, args: &str) -> Result<String, JsValue> {
        let args: SetSlideNotesArgs = parse_args(args)?;
        self.session
            .set_slide_notes(&local_context(), &args.slide_id, &args.text)
            .map_err(js_error)?;
        json(())
    }

    #[wasm_bindgen(js_name = moveSlideJson)]
    pub fn move_slide_json(&self, args: &str) -> Result<String, JsValue> {
        let args: MoveSlideArgs = parse_args(args)?;
        json(
            self.session
                .move_slide(&local_context(), &args.slide_id, args.to_index)
                .map_err(js_error)?,
        )
    }

    #[wasm_bindgen(js_name = addTextBoxJson)]
    pub fn add_text_box_json(&self, args: &str) -> Result<String, JsValue> {
        json(self.add_text_box(parse_args(args)?)?)
    }

    #[wasm_bindgen(js_name = addTextBoxProfiledJson)]
    pub fn add_text_box_profiled_json(&self, args: &str) -> Result<String, JsValue> {
        profiled(args, &mut performance_now, |args| self.add_text_box(args))
    }

    #[wasm_bindgen(js_name = addShapeJson)]
    pub fn add_shape_json(&self, args: &str) -> Result<String, JsValue> {
        let args: AddShapeArgs = parse_args(args)?;
        json(
            self.session
                .add_shape(&local_context(), &args.slide_id, &args.draft)
                .map_err(js_error)?,
        )
    }

    #[wasm_bindgen(js_name = addPictureJson)]
    pub fn add_picture_json(&self, args: &str) -> Result<String, JsValue> {
        let args: AddPictureArgs = parse_args(args)?;
        let media_bytes = base64::engine::general_purpose::STANDARD
            .decode(&args.media_base64)
            .map_err(|error| js_error(format!("invalid image data: {error}")))?;
        let draft = PictureDraft {
            name: args.name,
            rect: args.rect,
            content_type: args.content_type,
            media_bytes,
        };
        json(
            self.session
                .add_picture(&local_context(), &args.slide_id, &draft)
                .map_err(js_error)?,
        )
    }

    #[wasm_bindgen(js_name = removeShapeJson)]
    pub fn remove_shape_json(&self, args: &str) -> Result<String, JsValue> {
        let args: ShapeArgs = parse_args(args)?;
        json(
            self.session
                .remove_shape(&local_context(), &args.slide_id, &args.shape_id)
                .map_err(js_error)?,
        )
    }

    #[wasm_bindgen(js_name = bringShapeToFrontJson)]
    pub fn bring_shape_to_front_json(&self, args: &str) -> Result<String, JsValue> {
        let args: ShapeArgs = parse_args(args)?;
        json(
            self.session
                .bring_to_front(&local_context(), &args.slide_id, &args.shape_id)
                .map_err(js_error)?,
        )
    }

    #[wasm_bindgen(js_name = sendShapeToBackJson)]
    pub fn send_shape_to_back_json(&self, args: &str) -> Result<String, JsValue> {
        let args: ShapeArgs = parse_args(args)?;
        json(
            self.session
                .send_to_back(&local_context(), &args.slide_id, &args.shape_id)
                .map_err(js_error)?,
        )
    }

    #[wasm_bindgen(js_name = bringShapeForwardJson)]
    pub fn bring_shape_forward_json(&self, args: &str) -> Result<String, JsValue> {
        let args: ShapeArgs = parse_args(args)?;
        json(
            self.session
                .bring_forward(&local_context(), &args.slide_id, &args.shape_id)
                .map_err(js_error)?,
        )
    }

    #[wasm_bindgen(js_name = sendShapeBackwardJson)]
    pub fn send_shape_backward_json(&self, args: &str) -> Result<String, JsValue> {
        let args: ShapeArgs = parse_args(args)?;
        json(
            self.session
                .send_backward(&local_context(), &args.slide_id, &args.shape_id)
                .map_err(js_error)?,
        )
    }

    #[wasm_bindgen(js_name = moveShapeJson)]
    pub fn move_shape_json(&self, args: &str) -> Result<String, JsValue> {
        json(self.move_shape(parse_args(args)?)?)
    }

    #[wasm_bindgen(js_name = moveShapeProfiledJson)]
    pub fn move_shape_profiled_json(&self, args: &str) -> Result<String, JsValue> {
        profiled(args, &mut performance_now, |args| self.move_shape(args))
    }

    #[wasm_bindgen(js_name = resizeShapeJson)]
    pub fn resize_shape_json(&self, args: &str) -> Result<String, JsValue> {
        let args: ResizeShapeArgs = parse_args(args)?;
        json(
            self.session
                .resize_shape(
                    &local_context(),
                    &args.slide_id,
                    &args.shape_id,
                    args.width,
                    args.height,
                )
                .map_err(js_error)?,
        )
    }

    #[wasm_bindgen(js_name = setShapeRectJson)]
    pub fn set_shape_rect_json(&self, args: &str) -> Result<String, JsValue> {
        let args: SetShapeRectArgs = parse_args(args)?;
        json(
            self.session
                .set_shape_rect(&local_context(), &args.slide_id, &args.shape_id, args.rect)
                .map_err(js_error)?,
        )
    }

    #[wasm_bindgen(js_name = setShapeFillJson)]
    pub fn set_shape_fill_json(&self, args: &str) -> Result<String, JsValue> {
        let args: SetShapeFillArgs = parse_args(args)?;
        json(
            self.session
                .set_shape_fill(
                    &local_context(),
                    &args.slide_id,
                    &args.shape_id,
                    args.color.as_deref(),
                )
                .map_err(js_error)?,
        )
    }

    #[wasm_bindgen(js_name = setShapeStrokeJson)]
    pub fn set_shape_stroke_json(&self, args: &str) -> Result<String, JsValue> {
        let args: SetShapeStrokeArgs = parse_args(args)?;
        json(
            self.session
                .set_shape_stroke(
                    &local_context(),
                    &args.slide_id,
                    &args.shape_id,
                    &args.stroke,
                )
                .map_err(js_error)?,
        )
    }

    #[wasm_bindgen(js_name = setShapeAdjustJson)]
    pub fn set_shape_adjust_json(&self, args: &str) -> Result<String, JsValue> {
        let args: SetShapeAdjustArgs = parse_args(args)?;
        json(
            self.session
                .set_shape_adjust(
                    &local_context(),
                    &args.slide_id,
                    &args.shape_id,
                    &args.adjustments,
                )
                .map_err(js_error)?,
        )
    }

    #[wasm_bindgen(js_name = undoCaptureMode)]
    pub fn undo_capture_mode(&self) -> String {
        match self.session.undo_capture_mode() {
            crate::UndoCaptureMode::Auto => "auto",
            crate::UndoCaptureMode::Manual => "manual",
        }
        .to_owned()
    }

    #[wasm_bindgen(js_name = setUndoCaptureMode)]
    pub fn set_undo_capture_mode(&self, mode: &str) -> Result<(), JsValue> {
        let mode = match mode {
            "auto" => crate::UndoCaptureMode::Auto,
            "manual" => crate::UndoCaptureMode::Manual,
            _ => return Err(js_error("undo capture mode must be auto or manual")),
        };
        self.session.set_undo_capture_mode(mode);
        Ok(())
    }

    #[wasm_bindgen(js_name = addUndoBoundary)]
    pub fn add_undo_boundary(&self) {
        self.session.add_undo_barrier();
    }

    #[wasm_bindgen(js_name = undoJson)]
    pub fn undo_json(&self) -> Result<String, JsValue> {
        json(HistoryResult {
            applied: self.session.undo(),
            snapshot: self.session.snapshot().map_err(js_error)?,
        })
    }

    /// `undoJson` timed at its undo, snapshot and serialize boundaries, as
    /// `{"receipt": ..., "profile": {"undoMs", "snapshotMs", "serializeMs"}}`.
    #[wasm_bindgen(js_name = undoProfiledJson)]
    pub fn undo_profiled_json(&self) -> Result<String, JsValue> {
        let started = performance_now();
        let applied = self.session.undo();
        let undone = performance_now();
        let snapshot = self.session.snapshot().map_err(js_error)?;
        let snapshotted = performance_now();
        let receipt = json(HistoryResult { applied, snapshot })?;
        let serialized = performance_now();
        let profile = json(HistoryProfile {
            undo_ms: undone - started,
            snapshot_ms: snapshotted - undone,
            serialize_ms: serialized - snapshotted,
        })?;
        Ok(format!("{{\"receipt\":{receipt},\"profile\":{profile}}}"))
    }

    #[wasm_bindgen(js_name = redoJson)]
    pub fn redo_json(&self) -> Result<String, JsValue> {
        json(HistoryResult {
            applied: self.session.redo(),
            snapshot: self.session.snapshot().map_err(js_error)?,
        })
    }

    #[wasm_bindgen(js_name = canUndo)]
    pub fn can_undo(&self) -> bool {
        self.session.can_undo()
    }

    #[wasm_bindgen(js_name = canRedo)]
    pub fn can_redo(&self) -> bool {
        self.session.can_redo()
    }

    pub fn version() -> String {
        env!("CARGO_PKG_VERSION").to_owned()
    }
}

impl PptxDocument {
    fn oversized(&self) -> Result<String, JsValue> {
        outcome_json::<()>(&Err(oversized_request(self.session.version()))).map_err(js_error)
    }

    fn opened(session: DeckSession) -> Self {
        session.rotate_version(js_entropy());
        Self {
            session,
            update_observer: None,
        }
    }

    pub fn session(&self) -> &DeckSession {
        &self.session
    }

    fn insert_text(&self, args: InsertTextArgs) -> Result<TextReceipt, JsValue> {
        self.session
            .insert_text(
                &local_context(),
                &args.story_id,
                args.index,
                &args.text,
                &args.style,
            )
            .map_err(js_error)
    }

    fn delete_text(&self, args: DeleteTextArgs) -> Result<TextReceipt, JsValue> {
        self.session
            .delete_text(&local_context(), &args.story_id, args.start, args.end)
            .map_err(js_error)
    }

    fn insert_slide(&self, args: InsertSlideArgs) -> Result<SlideReceipt, JsValue> {
        self.session
            .insert_slide(
                &local_context(),
                args.index,
                args.layout_part_path.as_deref(),
            )
            .map_err(js_error)
    }

    fn add_text_box(&self, args: AddTextBoxArgs) -> Result<ShapeReceipt, JsValue> {
        self.session
            .add_text_box(&local_context(), &args.slide_id, &args.draft)
            .map_err(js_error)
    }

    fn move_shape(&self, args: MoveShapeArgs) -> Result<TransformReceipt, JsValue> {
        self.session
            .move_shape(
                &local_context(),
                &args.slide_id,
                &args.shape_id,
                args.x,
                args.y,
            )
            .map_err(js_error)
    }
}

/// one edit timed by `now` at its parse, apply and serialize boundaries, as
/// `{"receipt": <the usual json>, "profile": {"parseMs", "applyMs", "serializeMs"}}`.
fn profiled<A: serde::de::DeserializeOwned, R: Serialize>(
    args: &str,
    now: &mut impl FnMut() -> f64,
    apply: impl FnOnce(A) -> Result<R, JsValue>,
) -> Result<String, JsValue> {
    let started = now();
    let args = parse_args(args)?;
    let parsed = now();
    let receipt = apply(args)?;
    let applied = now();
    let receipt = json(receipt)?;
    let serialized = now();
    let profile = json(EditProfile {
        parse_ms: parsed - started,
        apply_ms: applied - parsed,
        serialize_ms: serialized - applied,
    })?;
    Ok(format!("{{\"receipt\":{receipt},\"profile\":{profile}}}"))
}

fn local_context() -> EditCtx {
    EditCtx::local("wasm")
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

fn parse_args<T: serde::de::DeserializeOwned>(args: &str) -> Result<T, JsValue> {
    serde_json::from_str(args).map_err(js_error)
}

fn json(value: impl Serialize) -> Result<String, JsValue> {
    serde_json::to_string(&value).map_err(js_error)
}

fn parse_client_id(client_id: f64) -> Result<u64, JsValue> {
    if !client_id.is_finite()
        || client_id.fract() != 0.0
        || client_id < 1.0
        || client_id > super::MAX_SAFE_CLIENT_ID as f64
    {
        return Err(JsValue::from_str(
            "client ID must be a positive safe integer below Number.MAX_SAFE_INTEGER",
        ));
    }
    Ok(client_id as u64)
}

fn js_error(error: impl std::fmt::Display) -> JsValue {
    JsValue::from_str(&error.to_string())
}

fn proposal_error(error: crate::ProposalError) -> JsValue {
    match error {
        crate::ProposalError::Stale(targets) => JsValue::from_str(
            &serde_json::json!({
                "code": "staleProposal", "targets": targets,
            })
            .to_string(),
        ),
        other => js_error(other),
    }
}

#[cfg(test)]
mod tests {
    use serde_json::{Value, json};

    use super::PptxDocument;

    const DECK: &[u8] = include_bytes!("../../../apps/demo/public/betteroffice-demo.pptx");

    fn envelope(json: &str) -> Value {
        serde_json::from_str(json).unwrap()
    }

    #[test]
    fn batch_envelopes_default_and_refuse_as_data() {
        let document = PptxDocument::open_collaborative(DECK, 91.0).unwrap();
        let version = document.document_version();
        let read = envelope(&document.read_content_json("{}").unwrap());
        assert_eq!(read["ok"], true);
        assert_eq!(read["version"], version.as_str());
        let story = &read["stories"][0];
        let target = json!({
            "kind": "range", "slideId": story["slideId"], "shapeId": story["shapeId"],
            "storyId": story["storyId"], "start": 0, "end": 0,
        });
        let request = json!({
            "expectVersion": version,
            "steps": [{"op": "insertText", "target": target, "at": "start", "text": "Draft: "}],
        })
        .to_string();

        let validated = envelope(&document.validate_edits_json(&request).unwrap());
        assert_eq!(validated["ok"], true);
        assert_eq!(validated["wouldApply"], true);
        assert_eq!(document.document_version(), version);

        let applied = envelope(&document.apply_edits_json(&request).unwrap());
        assert_eq!(applied["ok"], true);
        assert_eq!(applied["applied"], true);
        assert_eq!(applied["source"], "host");
        assert_eq!(applied["baseVersion"], version.as_str());
        assert_eq!(applied["version"], document.document_version().as_str());
        assert_eq!(applied["receipts"][0]["target"]["end"], 7);

        let stale = envelope(&document.apply_edits_json(&request).unwrap());
        assert_eq!(stale["ok"], false);
        assert_eq!(stale["failure"]["code"], "stale-version");
        assert_eq!(stale["version"], document.document_version().as_str());
        assert!(document.can_undo());

        let found = envelope(
            &document
                .find_text_json(r#"{"text":"Draft: ","limit":1}"#)
                .unwrap(),
        );
        assert_eq!(found["ok"], true);
        assert_eq!(found["matches"][0]["range"]["start"], 0);
        let empty = envelope(&document.find_text_json(r#"{"text":""}"#).unwrap());
        assert_eq!(empty["failure"]["code"], "invalid-step");
        let oversized = " ".repeat(crate::MAX_REQUEST_BYTES + 1);
        let refused = envelope(&document.apply_edits_json(&oversized).unwrap());
        assert_eq!(refused["ok"], false);
        assert_eq!(refused["failure"]["code"], "limit-exceeded");
        assert_eq!(refused["version"], document.document_version().as_str());
    }

    #[test]
    fn export_envelopes_carry_the_version_or_a_refusal() {
        let document = PptxDocument::open_collaborative(DECK, 94.0).unwrap();
        let version = document.document_version();
        let read = envelope(&document.export_structured_json("{}").unwrap());
        assert_eq!(read["ok"], true);
        assert_eq!(read["version"], version.as_str());
        assert_eq!(read["content"]["schemaVersion"], 1);
        assert_eq!(read["content"]["anchorScope"], "session");
        assert_eq!(read["content"]["readingOrder"], "shapeTree");
        let markdown = envelope(&document.export_markdown_json("{}").unwrap());
        assert_eq!(markdown["ok"], true);
        assert!(
            markdown["content"]["markdown"]
                .as_str()
                .unwrap()
                .starts_with("<!-- pptx-export:0 -->")
        );
        let refused = envelope(
            &document
                .export_structured_json(r#"{"maxBytes":8}"#)
                .unwrap(),
        );
        assert_eq!(refused["ok"], false);
        assert_eq!(refused["version"], version.as_str());
        assert_eq!(refused["failure"]["code"], "invalid-options");
        assert!(
            refused["failure"]
                .as_object()
                .unwrap()
                .get("target")
                .is_some_and(Value::is_null)
        );
        assert_eq!(document.document_version(), version);
    }

    #[test]
    fn documents_opened_from_the_same_bytes_never_share_versions() {
        let first = PptxDocument::open_collaborative(DECK, 92.0).unwrap();
        let second = PptxDocument::open_collaborative(DECK, 92.0).unwrap();
        assert_ne!(first.document_version(), second.document_version());
        let joined = PptxDocument::open_collaborative_from_update(
            &first.encode_state_as_update(),
            93.0,
            None,
        )
        .unwrap();
        assert_ne!(joined.document_version(), first.document_version());
    }
}
