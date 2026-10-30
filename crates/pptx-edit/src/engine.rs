use std::collections::{BTreeMap, VecDeque};
use std::fmt;
use std::sync::{Arc, Mutex};

use base64::Engine as _;
use serde::{Deserialize, Serialize};
use yrs::updates::decoder::Decode;
use yrs::updates::encoder::Encode;
use yrs::{StickyIndex, Subscription};

use crate::structured::{PptxExportOptions, export_outcome_json};
use crate::{
    CaretAnchor, CommentFlavor, DeckSession, DeckSnapshot, EditCtx, EditRequest, FindRequest,
    MAX_REQUEST_BYTES, PictureDraft, PresetShapeDraft, ReadRequest, ShapeDraft, ShapeReceipt,
    ShapeRect, ShapeStroke, SlideReceipt, TextReceipt, TextStyle, TextStylePatch, TransformReceipt,
    UpdateEvent, UpdateOrigin, outcome_json, oversized_request,
};

pub fn parse_pptx_json(data: &[u8]) -> Result<String, EngineError> {
    let package =
        pptx_parse::parse_pptx(data).map_err(|error| EngineError::message(error.to_string()))?;
    serde_json::to_string(&package).map_err(|error| EngineError::message(error.to_string()))
}

pub fn parse_pptx_json_without_media(data: &[u8]) -> Result<String, EngineError> {
    let package =
        pptx_parse::parse_pptx(data).map_err(|error| EngineError::message(error.to_string()))?;
    let metadata = PptxPackageWithoutMedia {
        presentation: &package.presentation,
        slides: &package.slides,
        layouts: &package.layouts,
        masters: &package.masters,
        themes: &package.themes,
        charts: &package.charts,
        diagram_drawings: &package.diagram_drawings,
        media: package
            .media
            .iter()
            .map(|part| MediaSummary {
                part_path: &part.part_path,
                content_type: &part.content_type,
                byte_length: part.bytes.len(),
            })
            .collect(),
        table_styles: &package.table_styles,
        comment_authors: &package.comment_authors,
        comments: &package.comments,
        comment_flavor: &package.comment_flavor,
        relationships: &package.relationships,
        shape_elements: package.models_connectors().then_some("withConnectors"),
    };
    serde_json::to_string(&metadata).map_err(|error| EngineError::message(error.to_string()))
}

#[derive(Serialize)]
#[serde(rename_all = "camelCase")]
struct PptxPackageWithoutMedia<'a> {
    presentation: &'a pptx_parse::Presentation,
    slides: &'a [pptx_parse::Slide],
    layouts: &'a [pptx_parse::SlideLayout],
    masters: &'a [pptx_parse::SlideMaster],
    themes: &'a [pptx_parse::ThemePart],
    charts: &'a [pptx_parse::ChartPart],
    #[serde(skip_serializing_if = "Vec::is_empty")]
    diagram_drawings: &'a Vec<pptx_parse::DiagramDrawing>,
    media: Vec<MediaSummary<'a>>,
    #[serde(skip_serializing_if = "ooxml_drawingml::TableStyleList::is_empty")]
    table_styles: &'a ooxml_drawingml::TableStyleList,
    #[serde(skip_serializing_if = "Vec::is_empty")]
    comment_authors: &'a Vec<pptx_parse::CommentAuthor>,
    #[serde(skip_serializing_if = "Vec::is_empty")]
    comments: &'a Vec<pptx_parse::Comment>,
    #[serde(skip_serializing_if = "Option::is_none")]
    comment_flavor: &'a Option<pptx_parse::CommentFlavor>,
    relationships: &'a BTreeMap<String, Vec<pptx_parse::Relationship>>,
    #[serde(skip_serializing_if = "Option::is_none")]
    shape_elements: Option<&'static str>,
}

#[derive(Serialize)]
#[serde(rename_all = "camelCase")]
struct MediaSummary<'a> {
    part_path: &'a str,
    content_type: &'a str,
    byte_length: usize,
}

pub fn version() -> &'static str {
    env!("CARGO_PKG_VERSION")
}

#[derive(Clone, Debug, Eq, PartialEq)]
pub struct EngineError {
    pub kind: ErrorKind,
    pub message: String,
}

#[derive(Clone, Debug, Eq, PartialEq)]
pub enum ErrorKind {
    /// thrown as a plain string (today: JsValue::from_str)
    Message,
    /// thrown as an Error named "PresentationPeerError" with a string `code` property (today: peer_error)
    Peer { code: String },
    /// thrown as an Error named "PptxEditDecodeError" (today: validateEditsJson decode failure)
    Decode,
}

impl EngineError {
    pub fn message(message: impl Into<String>) -> Self {
        Self {
            kind: ErrorKind::Message,
            message: message.into(),
        }
    }

    pub fn peer(code: String, message: String) -> Self {
        Self {
            kind: ErrorKind::Peer { code },
            message,
        }
    }

    pub fn decode(message: String) -> Self {
        Self {
            kind: ErrorKind::Decode,
            message,
        }
    }
}

impl fmt::Display for EngineError {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        f.write_str(&self.message)
    }
}

impl std::error::Error for EngineError {}

#[derive(Clone, Copy)]
pub struct Host {
    pub entropy: fn() -> u64,
    pub timer_ms: fn() -> f64,
}

pub struct Document {
    session: DeckSession,
    update_observer: Option<UpdateObserver>,
    host: Host,
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
struct StoryFormattingArgs {
    slide_id: String,
    shape_id: String,
    story_id: String,
}

#[derive(Deserialize)]
#[serde(rename_all = "camelCase")]
struct AnchorCaretArgs {
    story_id: String,
    index: u32,
}

/// A caret anchor across the JSON boundary: the story and its base64 v1-encoded sticky index.
#[derive(Deserialize, Serialize)]
#[serde(rename_all = "camelCase")]
struct CaretAnchorJson {
    story_id: String,
    position: String,
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
    #[serde(default)]
    allow_paragraph_join: bool,
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
struct SetSlideLayoutArgs {
    slide_id: String,
    layout_part_path: String,
    #[serde(default)]
    mapping: Vec<(String, u32)>,
}

#[derive(Deserialize)]
#[serde(rename_all = "camelCase")]
struct MoveSlideArgs {
    slide_id: String,
    to_index: u32,
}

#[derive(Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
struct SetStoryParagraphsArgs {
    story_id: String,
    paragraphs: Vec<crate::TextParagraphDraft>,
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
#[serde(rename_all = "camelCase", deny_unknown_fields)]
struct TableCheckArgs {
    edit: crate::TableEdit,
    slide_id: String,
    shape_id: String,
    index: u32,
    #[serde(default)]
    position: Option<crate::TableInsertPosition>,
    #[serde(default)]
    cells: Vec<String>,
    #[serde(default)]
    width: crate::TableColumnWidth,
}

#[derive(Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
struct TableArgs {
    slide_id: String,
    shape_id: String,
    index: u32,
    #[serde(default)]
    position: Option<crate::TableInsertPosition>,
    #[serde(default)]
    cells: Vec<String>,
    #[serde(default)]
    width: crate::TableColumnWidth,
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

impl Document {
    pub fn open_replay_baseline(
        source: &[u8],
        client_id: Option<f64>,
        initial_update: Option<Vec<u8>>,
        host: Host,
    ) -> Result<Document, EngineError> {
        let client_id = client_id
            .map(|id| {
                if !id.is_finite()
                    || id.fract() != 0.0
                    || id < 1.0
                    || id > crate::MAX_SAFE_CLIENT_ID as f64
                {
                    return Err(peer_error(crate::peer::PeerError::new(
                        "clientId",
                        "client id must be a positive safe integer below Number.MAX_SAFE_INTEGER",
                    )));
                }
                Ok(id as u64)
            })
            .transpose()?;
        DeckSession::open_replay_baseline(source, client_id, initial_update.as_deref())
            .map(|session| Self::opened(session, host))
            .map_err(peer_error)
    }

    pub fn peer_hydration_json(&self, fonts_json: &str) -> Result<String, EngineError> {
        let fonts = serde_json::from_str(fonts_json).map_err(|error| peer_error(error.into()))?;
        self.session.peer_identity(fonts).map_err(peer_error)
    }

    pub fn open_peer_deck_json(
        source: &[u8],
        identity_json: &str,
        initial_update: Option<Vec<u8>>,
        host: Host,
    ) -> Result<Document, EngineError> {
        DeckSession::open_peer_deck(source, identity_json, initial_update.as_deref())
            .map(|session| Self {
                session,
                update_observer: None,
                host,
            })
            .map_err(peer_error)
    }

    pub fn register_peer_fonts_json(&self, fonts_json: &str) -> Result<(), EngineError> {
        let fonts = serde_json::from_str(fonts_json).map_err(|error| peer_error(error.into()))?;
        self.session.register_peer_fonts(fonts).map_err(peer_error)
    }

    pub fn adopt_peer_identity(&self) -> Result<(), EngineError> {
        self.session.adopt_peer_identity().map_err(peer_error)
    }

    pub fn replay_json(&self, envelope: &str) -> Result<String, EngineError> {
        self.session.replay_json(envelope).map_err(peer_error)
    }

    pub fn propose_json(&self, args: &str) -> Result<String, EngineError> {
        json(
            self.session
                .propose(parse_args(args)?)
                .map_err(proposal_error)?,
        )
    }

    pub fn list_proposals_json(&self) -> Result<String, EngineError> {
        json(self.session.proposals().map_err(proposal_error)?)
    }

    pub fn preview_proposal_json(&self, args: &str) -> Result<String, EngineError> {
        let args: ProposalIdArgs = parse_args(args)?;
        json(
            self.session
                .preview_proposal(&args.id)
                .map_err(proposal_error)?,
        )
    }

    pub fn accept_proposal_json(&self, args: &str) -> Result<String, EngineError> {
        let args: ProposalIdArgs = parse_args(args)?;
        json(
            self.session
                .accept_proposal(&args.id, args.force)
                .map_err(proposal_error)?,
        )
    }

    pub fn reject_proposal_json(&self, args: &str) -> Result<String, EngineError> {
        let args: ProposalIdArgs = parse_args(args)?;
        json(self.session.reject_proposal(&args.id))
    }

    /// Opens a session that can exchange updates. It is not marked collaborative until
    /// `setCollaborative` is called.
    pub fn open_collaborative(
        bytes: &[u8],
        client_id: f64,
        host: Host,
    ) -> Result<Document, EngineError> {
        let client_id = parse_client_id(client_id)?;
        DeckSession::open(bytes, client_id)
            .map(|session| Self::opened(session, host))
            .map_err(message_error)
    }

    /// `source` is the file the update was seeded from; when it matches the
    /// recorded fingerprint the session keeps its part bytes and can save, and
    /// failing to reattach it throws.
    /// Any other bytes fall back to the bare update session, whose `saveBytes`
    /// fails — joining a room must not depend on carrying the right file.
    pub fn open_collaborative_from_update(
        update: &[u8],
        client_id: f64,
        source: Option<Vec<u8>>,
        host: Host,
    ) -> Result<Document, EngineError> {
        let client_id = parse_client_id(client_id)?;
        let session = DeckSession::open_from_update(update, client_id).map_err(message_error)?;
        let session = match source {
            Some(source) if session.seeded_from(&source).map_err(message_error)? => {
                session.attach_source_owned(source).map_err(message_error)?
            }
            _ => session,
        };
        Ok(Self::opened(session, host))
    }

    pub fn client_id(&self) -> f64 {
        self.session.client_id() as f64
    }

    pub fn snapshot_json(&self) -> Result<String, EngineError> {
        json(self.session.snapshot().map_err(message_error)?)
    }

    pub fn session_metadata_json(&self) -> Result<String, EngineError> {
        json(self.session.session_metadata().map_err(message_error)?)
    }

    // Version-checked host edits. Requests and results are JSON; policy outcomes come back as
    // `{"ok":true,...}` or `{"ok":false,"version","failure":{"code","message","stepIndex"?,
    // "conflictingStepIndex"?,"target"?}}`, while malformed requests throw.

    /// The session-scoped version token of the committed deck state.
    pub fn document_version(&self) -> String {
        self.session.version().to_string()
    }

    /// `{"slideIds"?}` -> `{"ok":true,"version","slides","stories"}`.
    pub fn read_content_json(&self, request: &str) -> Result<String, EngineError> {
        if request.len() > MAX_REQUEST_BYTES {
            return self.oversized();
        }
        let request: ReadRequest = parse_args(request)?;
        outcome_json(&self.session.read_content(&request).map_err(message_error)?)
            .map_err(message_error)
    }

    pub fn read_command_content_json(&self, request: &str) -> Result<String, EngineError> {
        if request.len() > MAX_REQUEST_BYTES {
            return self.oversized();
        }
        let request: ReadRequest = parse_args(request)?;
        outcome_json(
            &self
                .session
                .read_command_content(&request)
                .map_err(message_error)?,
        )
        .map_err(message_error)
    }

    /// `{"text","within"?,"limit"?}` -> `{"ok":true,"version","matches","truncated"}`.
    pub fn find_text_json(&self, request: &str) -> Result<String, EngineError> {
        if request.len() > MAX_REQUEST_BYTES {
            return self.oversized();
        }
        let request: FindRequest = parse_args(request)?;
        outcome_json(&self.session.find_text(&request).map_err(message_error)?)
            .map_err(message_error)
    }

    /// Runs every check of `applyEditsJson`, staging included, without changing anything:
    /// `{"ok":true,"baseVersion","wouldApply","previews"}`.
    pub fn validate_edits_json(&self, request: &str) -> Result<String, EngineError> {
        if request.len() > MAX_REQUEST_BYTES {
            return self.oversized();
        }
        let request: EditRequest = serde_json::from_str(request)
            .map_err(|error| EngineError::decode(error.to_string()))?;
        outcome_json(
            &self
                .session
                .validate_edits(&request)
                .map_err(message_error)?,
        )
        .map_err(message_error)
    }

    /// Applies a batch all-or-nothing: `{"ok":true,"baseVersion","version","applied","source",
    /// "changedSlides","changedStories","receipts"}`.
    pub fn apply_edits_json(&self, request: &str) -> Result<String, EngineError> {
        if request.len() > MAX_REQUEST_BYTES {
            return self.oversized();
        }
        let request: EditRequest = parse_args(request)?;
        outcome_json(&self.session.apply_edits(&request).map_err(message_error)?)
            .map_err(message_error)
    }

    /// Structured export of the committed deck: `{"includeHiddenSlides"?,"includeHiddenShapes"?,
    /// "includeNotes"?,"includeComments"?,"includeFormatting"?,"maxBlocks"?,"maxBytes"?}` ->
    /// `{"ok":true,"version","content"}` or `{"ok":false,"version","failure"}`. Nothing changes.
    pub fn export_structured_json(&self, options: &str) -> Result<String, EngineError> {
        let options: PptxExportOptions = parse_args(options)?;
        export_outcome_json(
            &self
                .session
                .export_structured(&options)
                .map_err(message_error)?,
        )
        .map_err(message_error)
    }

    /// `exportStructuredJson` rendered as Markdown from the same read:
    /// `{"ok":true,"version","content":{"markdown","anchors","diagnostics","truncated"}}`.
    pub fn export_markdown_json(&self, options: &str) -> Result<String, EngineError> {
        let options: PptxExportOptions = parse_args(options)?;
        export_outcome_json(
            &self
                .session
                .export_markdown(&options)
                .map_err(message_error)?,
        )
        .map_err(message_error)
    }

    pub fn search_text_json(&self, args: &str) -> Result<String, EngineError> {
        let args: SearchTextArgs = parse_args(args)?;
        json(
            self.session
                .search_text(
                    &args.query,
                    args.case_sensitive,
                    args.limit.map(|value| value as usize),
                )
                .map_err(message_error)?,
        )
    }

    pub fn story_json(&self, args: &str) -> Result<String, EngineError> {
        let args: StoryArgs = parse_args(args)?;
        json(self.session.story(&args.story_id).map_err(message_error)?)
    }

    pub fn story_formatting_json(&self, args: &str) -> Result<String, EngineError> {
        let args: StoryFormattingArgs = parse_args(args)?;
        json(
            self.session
                .story_formatting(&args.slide_id, &args.shape_id, &args.story_id)
                .map_err(message_error)?,
        )
    }

    /// Anchors a caret offset so that later edits, undo and remote updates move it.
    pub fn anchor_caret_json(&self, args: &str) -> Result<String, EngineError> {
        let args: AnchorCaretArgs = parse_args(args)?;
        let anchor = self
            .session
            .anchor_caret(&args.story_id, args.index)
            .map_err(message_error)?;
        json(CaretAnchorJson {
            position: base64::engine::general_purpose::STANDARD.encode(anchor.position.encode_v1()),
            story_id: anchor.story_id,
        })
    }

    /// The current offset of an anchor from `anchorCaretJson`, or `null` once its story is gone.
    pub fn resolve_caret_anchor_json(&self, args: &str) -> Result<String, EngineError> {
        let args: CaretAnchorJson = parse_args(args)?;
        let bytes = base64::engine::general_purpose::STANDARD
            .decode(args.position)
            .map_err(message_error)?;
        let position = StickyIndex::decode_v1(&bytes).map_err(message_error)?;
        json(self.session.resolve_caret_anchor(&CaretAnchor {
            story_id: args.story_id,
            position,
        }))
    }

    pub fn media_bytes(&self, part_path: &str) -> Result<Vec<u8>, EngineError> {
        self.session.media_bytes(part_path).map_err(message_error)
    }

    /// Serializes the deck back to `.pptx` bytes, edits included.
    pub fn save_bytes(&self) -> Result<Vec<u8>, EngineError> {
        self.session.save().map_err(message_error)
    }

    pub fn encode_state_vector(&self) -> Vec<u8> {
        self.session.encode_state_vector_v1()
    }

    pub fn encode_state_as_update(&self) -> Vec<u8> {
        self.session.encode_state_as_update_v1()
    }

    pub fn encode_diff(&self, remote_state_vector: &[u8]) -> Result<Vec<u8>, EngineError> {
        self.session
            .encode_diff_v1(remote_state_vector)
            .map_err(message_error)
    }

    pub fn set_collaborative(&self) -> Result<(), EngineError> {
        self.session.set_collaborative().map_err(message_error)
    }

    pub fn apply_update_json(&self, update: &[u8]) -> Result<String, EngineError> {
        json(
            self.session
                .apply_update_v1(update)
                .map_err(message_error)?,
        )
    }

    pub fn start_update_observation(&mut self) -> Result<(), EngineError> {
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
            .map_err(message_error)?;
        self.update_observer = Some(UpdateObserver {
            pending,
            _subscription: subscription,
        });
        Ok(())
    }

    pub fn clear_update_observation(&mut self) {
        self.update_observer = None;
    }

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

    pub fn insert_text_json(&self, args: &str) -> Result<String, EngineError> {
        json(self.insert_text(parse_args(args)?)?)
    }

    pub fn insert_text_profiled_json(&self, args: &str) -> Result<String, EngineError> {
        let mut timer = self.host.timer_ms;
        profiled(args, &mut timer, |args| self.insert_text(args))
    }

    pub fn delete_text_json(&self, args: &str) -> Result<String, EngineError> {
        json(self.delete_text(parse_args(args)?)?)
    }

    pub fn delete_text_profiled_json(&self, args: &str) -> Result<String, EngineError> {
        let mut timer = self.host.timer_ms;
        profiled(args, &mut timer, |args| self.delete_text(args))
    }

    pub fn format_text_json(&self, args: &str) -> Result<String, EngineError> {
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
                .map_err(message_error)?,
        )
    }

    pub fn set_paragraph_alignment_json(&self, args: &str) -> Result<String, EngineError> {
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
                .map_err(message_error)?,
        )
    }

    pub fn add_comment_json(&self, args: &str) -> Result<String, EngineError> {
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
                .map_err(message_error)?,
        )
    }

    pub fn reply_to_comment_json(&self, args: &str) -> Result<String, EngineError> {
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
                .map_err(message_error)?,
        )
    }

    pub fn set_comment_position_json(&self, args: &str) -> Result<String, EngineError> {
        let args: CommentPositionArgs = parse_args(args)?;
        json(
            self.session
                .set_comment_position(&local_context(), &args.comment_id, args.x_emu, args.y_emu)
                .map_err(message_error)?,
        )
    }

    pub fn set_comment_status_json(&self, args: &str) -> Result<String, EngineError> {
        let args: CommentStatusArgs = parse_args(args)?;
        json(
            self.session
                .set_comment_status(&local_context(), &args.comment_id, args.resolved)
                .map_err(message_error)?,
        )
    }

    pub fn remove_comment_json(&self, args: &str) -> Result<String, EngineError> {
        let args: CommentIdArgs = parse_args(args)?;
        json(
            self.session
                .remove_comment(&local_context(), &args.comment_id)
                .map_err(message_error)?,
        )
    }

    pub fn set_comment_flavor_json(&self, args: &str) -> Result<String, EngineError> {
        let args: CommentFlavorArgs = parse_args(args)?;
        json(
            self.session
                .set_comment_flavor(&local_context(), args.flavor)
                .map_err(message_error)?,
        )
    }

    pub fn comments_json(&self) -> Result<String, EngineError> {
        json(self.session.comments().map_err(message_error)?)
    }

    pub fn insert_paragraph_break_json(&self, args: &str) -> Result<String, EngineError> {
        let args: ParagraphBreakArgs = parse_args(args)?;
        json(
            self.session
                .insert_paragraph_break(&local_context(), &args.story_id, args.index)
                .map_err(message_error)?,
        )
    }

    pub fn set_story_paragraphs_json(&self, args: &str) -> Result<String, EngineError> {
        let args: SetStoryParagraphsArgs = parse_args(args)?;
        json(
            self.session
                .set_story_paragraphs(&local_context(), &args.story_id, &args.paragraphs)
                .map_err(message_error)?,
        )
    }

    pub fn insert_table_row_json(&self, args: &str) -> Result<String, EngineError> {
        let args: TableArgs = parse_args(args)?;
        json(
            self.session
                .insert_table_row(
                    &local_context(),
                    &args.slide_id,
                    &args.shape_id,
                    args.index,
                    args.position
                        .ok_or_else(|| message_error("position must be before or after"))?,
                    &args.cells,
                )
                .map_err(message_error)?,
        )
    }

    pub fn check_table_edit_json(&self, args: &str) -> Result<(), EngineError> {
        let args: TableCheckArgs = parse_args(args)?;
        let insert = matches!(
            args.edit,
            crate::TableEdit::InsertRow | crate::TableEdit::InsertColumn
        );
        let position = match args.position {
            Some(position) => position,
            None if insert => return Err(message_error("position must be before or after")),
            None => crate::TableInsertPosition::After,
        };
        self.session
            .check_table_edit_with_width(
                args.edit,
                &args.slide_id,
                &args.shape_id,
                args.index,
                position,
                &args.cells,
                args.width,
            )
            .map_err(message_error)
    }

    pub fn delete_table_row_json(&self, args: &str) -> Result<String, EngineError> {
        let args: TableArgs = parse_args(args)?;
        json(
            self.session
                .delete_table_row(&local_context(), &args.slide_id, &args.shape_id, args.index)
                .map_err(message_error)?,
        )
    }

    pub fn insert_table_column_json(&self, args: &str) -> Result<String, EngineError> {
        let args: TableArgs = parse_args(args)?;
        json(
            self.session
                .insert_table_column_with_width(
                    &local_context(),
                    &args.slide_id,
                    &args.shape_id,
                    args.index,
                    args.position
                        .ok_or_else(|| message_error("position must be before or after"))?,
                    &args.cells,
                    args.width,
                )
                .map_err(message_error)?,
        )
    }

    pub fn delete_table_column_json(&self, args: &str) -> Result<String, EngineError> {
        let args: TableArgs = parse_args(args)?;
        json(
            self.session
                .delete_table_column(&local_context(), &args.slide_id, &args.shape_id, args.index)
                .map_err(message_error)?,
        )
    }

    pub fn insert_slide_json(&self, args: &str) -> Result<String, EngineError> {
        json(self.insert_slide(parse_args(args)?)?)
    }

    pub fn duplicate_slide_json(&self, args: &str) -> Result<String, EngineError> {
        let args: MoveSlideArgs = parse_args(args)?;
        json(
            self.session
                .duplicate_slide(&local_context(), &args.slide_id, args.to_index)
                .map_err(message_error)?,
        )
    }

    pub fn insert_slide_profiled_json(&self, args: &str) -> Result<String, EngineError> {
        let mut timer = self.host.timer_ms;
        profiled(args, &mut timer, |args| self.insert_slide(args))
    }

    pub fn delete_slide_json(&self, args: &str) -> Result<String, EngineError> {
        let args: SlideArgs = parse_args(args)?;
        json(
            self.session
                .delete_slide(&local_context(), &args.slide_id)
                .map_err(message_error)?,
        )
    }

    pub fn set_slide_layout_json(&self, args: &str) -> Result<String, EngineError> {
        let args: SetSlideLayoutArgs = parse_args(args)?;
        self.session
            .set_slide_layout(
                &local_context(),
                &args.slide_id,
                &args.layout_part_path,
                &args.mapping,
            )
            .map_err(message_error)?;
        json(())
    }

    pub fn set_slide_notes_json(&self, args: &str) -> Result<String, EngineError> {
        let args: SetSlideNotesArgs = parse_args(args)?;
        self.session
            .set_slide_notes(&local_context(), &args.slide_id, &args.text)
            .map_err(message_error)?;
        json(())
    }

    pub fn move_slide_json(&self, args: &str) -> Result<String, EngineError> {
        let args: MoveSlideArgs = parse_args(args)?;
        json(
            self.session
                .move_slide(&local_context(), &args.slide_id, args.to_index)
                .map_err(message_error)?,
        )
    }

    pub fn add_text_box_json(&self, args: &str) -> Result<String, EngineError> {
        json(self.add_text_box(parse_args(args)?)?)
    }

    pub fn add_text_box_profiled_json(&self, args: &str) -> Result<String, EngineError> {
        let mut timer = self.host.timer_ms;
        profiled(args, &mut timer, |args| self.add_text_box(args))
    }

    pub fn add_shape_json(&self, args: &str) -> Result<String, EngineError> {
        let args: AddShapeArgs = parse_args(args)?;
        json(
            self.session
                .add_shape(&local_context(), &args.slide_id, &args.draft)
                .map_err(message_error)?,
        )
    }

    pub fn add_picture_json(&self, args: &str) -> Result<String, EngineError> {
        let args: AddPictureArgs = parse_args(args)?;
        let media_bytes = base64::engine::general_purpose::STANDARD
            .decode(&args.media_base64)
            .map_err(|error| message_error(format!("invalid image data: {error}")))?;
        let draft = PictureDraft {
            name: args.name,
            rect: args.rect,
            content_type: args.content_type,
            media_bytes,
        };
        json(
            self.session
                .add_picture(&local_context(), &args.slide_id, &draft)
                .map_err(message_error)?,
        )
    }

    pub fn remove_shape_json(&self, args: &str) -> Result<String, EngineError> {
        let args: ShapeArgs = parse_args(args)?;
        json(
            self.session
                .remove_shape(&local_context(), &args.slide_id, &args.shape_id)
                .map_err(message_error)?,
        )
    }

    pub fn bring_shape_to_front_json(&self, args: &str) -> Result<String, EngineError> {
        let args: ShapeArgs = parse_args(args)?;
        json(
            self.session
                .bring_to_front(&local_context(), &args.slide_id, &args.shape_id)
                .map_err(message_error)?,
        )
    }

    pub fn send_shape_to_back_json(&self, args: &str) -> Result<String, EngineError> {
        let args: ShapeArgs = parse_args(args)?;
        json(
            self.session
                .send_to_back(&local_context(), &args.slide_id, &args.shape_id)
                .map_err(message_error)?,
        )
    }

    pub fn bring_shape_forward_json(&self, args: &str) -> Result<String, EngineError> {
        let args: ShapeArgs = parse_args(args)?;
        json(
            self.session
                .bring_forward(&local_context(), &args.slide_id, &args.shape_id)
                .map_err(message_error)?,
        )
    }

    pub fn send_shape_backward_json(&self, args: &str) -> Result<String, EngineError> {
        let args: ShapeArgs = parse_args(args)?;
        json(
            self.session
                .send_backward(&local_context(), &args.slide_id, &args.shape_id)
                .map_err(message_error)?,
        )
    }

    pub fn move_shape_json(&self, args: &str) -> Result<String, EngineError> {
        json(self.move_shape(parse_args(args)?)?)
    }

    pub fn move_shape_profiled_json(&self, args: &str) -> Result<String, EngineError> {
        let mut timer = self.host.timer_ms;
        profiled(args, &mut timer, |args| self.move_shape(args))
    }

    pub fn resize_shape_json(&self, args: &str) -> Result<String, EngineError> {
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
                .map_err(message_error)?,
        )
    }

    pub fn set_shape_rect_json(&self, args: &str) -> Result<String, EngineError> {
        let args: SetShapeRectArgs = parse_args(args)?;
        json(
            self.session
                .set_shape_rect(&local_context(), &args.slide_id, &args.shape_id, args.rect)
                .map_err(message_error)?,
        )
    }

    pub fn set_shape_fill_json(&self, args: &str) -> Result<String, EngineError> {
        let args: SetShapeFillArgs = parse_args(args)?;
        json(
            self.session
                .set_shape_fill(
                    &local_context(),
                    &args.slide_id,
                    &args.shape_id,
                    args.color.as_deref(),
                )
                .map_err(message_error)?,
        )
    }

    pub fn set_shape_stroke_json(&self, args: &str) -> Result<String, EngineError> {
        let args: SetShapeStrokeArgs = parse_args(args)?;
        json(
            self.session
                .set_shape_stroke(
                    &local_context(),
                    &args.slide_id,
                    &args.shape_id,
                    &args.stroke,
                )
                .map_err(message_error)?,
        )
    }

    pub fn set_shape_adjust_json(&self, args: &str) -> Result<String, EngineError> {
        let args: SetShapeAdjustArgs = parse_args(args)?;
        json(
            self.session
                .set_shape_adjust(
                    &local_context(),
                    &args.slide_id,
                    &args.shape_id,
                    &args.adjustments,
                )
                .map_err(message_error)?,
        )
    }

    pub fn undo_capture_mode(&self) -> String {
        match self.session.undo_capture_mode() {
            crate::UndoCaptureMode::Auto => "auto",
            crate::UndoCaptureMode::Manual => "manual",
        }
        .to_owned()
    }

    pub fn set_undo_capture_mode(&self, mode: &str) -> Result<(), EngineError> {
        let mode = match mode {
            "auto" => crate::UndoCaptureMode::Auto,
            "manual" => crate::UndoCaptureMode::Manual,
            _ => return Err(message_error("undo capture mode must be auto or manual")),
        };
        self.session.set_undo_capture_mode(mode);
        Ok(())
    }

    pub fn add_undo_boundary(&self) {
        self.session.add_undo_barrier();
    }

    pub fn undo_json(&self) -> Result<String, EngineError> {
        json(HistoryResult {
            applied: self.session.undo(),
            snapshot: self.session.snapshot().map_err(message_error)?,
        })
    }

    pub fn undo_receipt_json(&self) -> Result<String, EngineError> {
        json(self.session.history_step(false).map_err(message_error)?)
    }

    pub fn redo_receipt_json(&self) -> Result<String, EngineError> {
        json(self.session.history_step(true).map_err(message_error)?)
    }

    /// `undoJson` timed at its undo, snapshot and serialize boundaries, as
    /// `{"receipt": ..., "profile": {"undoMs", "snapshotMs", "serializeMs"}}`.
    pub fn undo_profiled_json(&self) -> Result<String, EngineError> {
        let started = (self.host.timer_ms)();
        let applied = self.session.undo();
        let undone = (self.host.timer_ms)();
        let snapshot = self.session.snapshot().map_err(message_error)?;
        let snapshotted = (self.host.timer_ms)();
        let receipt = json(HistoryResult { applied, snapshot })?;
        let serialized = (self.host.timer_ms)();
        let profile = json(HistoryProfile {
            undo_ms: undone - started,
            snapshot_ms: snapshotted - undone,
            serialize_ms: serialized - snapshotted,
        })?;
        Ok(format!("{{\"receipt\":{receipt},\"profile\":{profile}}}"))
    }

    pub fn redo_json(&self) -> Result<String, EngineError> {
        json(HistoryResult {
            applied: self.session.redo(),
            snapshot: self.session.snapshot().map_err(message_error)?,
        })
    }

    pub fn can_undo(&self) -> bool {
        self.session.can_undo()
    }

    pub fn can_redo(&self) -> bool {
        self.session.can_redo()
    }
}

impl Document {
    fn oversized(&self) -> Result<String, EngineError> {
        outcome_json::<()>(&Err(oversized_request(self.session.version()))).map_err(message_error)
    }

    fn opened(session: DeckSession, host: Host) -> Self {
        session.rotate_version((host.entropy)());
        Self {
            session,
            update_observer: None,
            host,
        }
    }

    pub fn session(&self) -> &DeckSession {
        &self.session
    }

    fn insert_text(&self, args: InsertTextArgs) -> Result<TextReceipt, EngineError> {
        self.session
            .insert_text(
                &local_context(),
                &args.story_id,
                args.index,
                &args.text,
                &args.style,
            )
            .map_err(message_error)
    }

    fn delete_text(&self, args: DeleteTextArgs) -> Result<TextReceipt, EngineError> {
        if args.allow_paragraph_join {
            return self
                .session
                .delete_text_joining_paragraphs(
                    &local_context(),
                    &args.story_id,
                    args.start,
                    args.end,
                )
                .map_err(message_error);
        }
        self.session
            .delete_text(&local_context(), &args.story_id, args.start, args.end)
            .map_err(message_error)
    }

    fn insert_slide(&self, args: InsertSlideArgs) -> Result<SlideReceipt, EngineError> {
        self.session
            .insert_slide(
                &local_context(),
                args.index,
                args.layout_part_path.as_deref(),
            )
            .map_err(message_error)
    }

    fn add_text_box(&self, args: AddTextBoxArgs) -> Result<ShapeReceipt, EngineError> {
        self.session
            .add_text_box(&local_context(), &args.slide_id, &args.draft)
            .map_err(message_error)
    }

    fn move_shape(&self, args: MoveShapeArgs) -> Result<TransformReceipt, EngineError> {
        self.session
            .move_shape(
                &local_context(),
                &args.slide_id,
                &args.shape_id,
                args.x,
                args.y,
            )
            .map_err(message_error)
    }
}

/// one edit timed by `now` at its parse, apply and serialize boundaries, as
/// `{"receipt": <the usual json>, "profile": {"parseMs", "applyMs", "serializeMs"}}`.
fn profiled<A: serde::de::DeserializeOwned, R: Serialize>(
    args: &str,
    now: &mut impl FnMut() -> f64,
    apply: impl FnOnce(A) -> Result<R, EngineError>,
) -> Result<String, EngineError> {
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

fn parse_args<T: serde::de::DeserializeOwned>(args: &str) -> Result<T, EngineError> {
    serde_json::from_str(args).map_err(message_error)
}

fn json(value: impl Serialize) -> Result<String, EngineError> {
    serde_json::to_string(&value).map_err(message_error)
}

fn parse_client_id(client_id: f64) -> Result<u64, EngineError> {
    if !client_id.is_finite()
        || client_id.fract() != 0.0
        || client_id < 1.0
        || client_id > super::MAX_SAFE_CLIENT_ID as f64
    {
        return Err(message_error(
            "client ID must be a positive safe integer below Number.MAX_SAFE_INTEGER",
        ));
    }
    Ok(client_id as u64)
}

fn message_error(error: impl fmt::Display) -> EngineError {
    EngineError::message(error.to_string())
}

fn peer_error(error: crate::peer::PeerError) -> EngineError {
    EngineError::peer(error.code, error.message)
}

fn proposal_error(error: crate::ProposalError) -> EngineError {
    match error {
        crate::ProposalError::Stale(targets) => EngineError::message(
            serde_json::json!({
                "code": "staleProposal", "targets": targets,
            })
            .to_string(),
        ),
        other => message_error(other),
    }
}

#[cfg(test)]
mod tests {
    use serde_json::{Value, json};

    use super::{Document, Host};

    const DECK: &[u8] = include_bytes!("../../../apps/demo/public/betteroffice-demo.pptx");

    fn test_entropy() -> u64 {
        0
    }

    fn test_timer() -> f64 {
        0.0
    }

    const TEST_HOST: Host = Host {
        entropy: test_entropy,
        timer_ms: test_timer,
    };

    fn envelope(json: &str) -> Value {
        serde_json::from_str(json).unwrap()
    }

    #[test]
    fn batch_envelopes_default_and_refuse_as_data() {
        let document = Document::open_collaborative(DECK, 91.0, TEST_HOST).unwrap();
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
        let document = Document::open_collaborative(DECK, 94.0, TEST_HOST).unwrap();
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
        let first = Document::open_collaborative(DECK, 92.0, TEST_HOST).unwrap();
        let second = Document::open_collaborative(DECK, 92.0, TEST_HOST).unwrap();
        assert_ne!(first.document_version(), second.document_version());
        let joined = Document::open_collaborative_from_update(
            &first.encode_state_as_update(),
            93.0,
            None,
            TEST_HOST,
        )
        .unwrap();
        assert_ne!(joined.document_version(), first.document_version());
    }
}
