use std::collections::{BTreeMap, BTreeSet};
use std::sync::atomic::Ordering;

use base64::Engine as _;
use serde::{Deserialize, Serialize, de::DeserializeOwned};
use serde_json::{Value, json};

use crate::peer::{PeerError, PeerOpening};
use crate::{
    CommentFlavor, DeckSession, EditCtx, EditRequest, PictureDraft, PresetShapeDraft, ProposalEdit,
    ProposalRequest, ShapeDraft, ShapeRect, ShapeStroke, TextStyle, TextStylePatch,
};

const MAX_SAFE_INTEGER: u64 = (1_u64 << 53) - 1;

#[derive(Default)]
pub(crate) struct ReplayState {
    pub ready: bool,
    pub opening: Option<PeerOpening>,
    pub sequence: u64,
    revision: u64,
    failed: bool,
    applying: bool,
}

impl ReplayState {
    pub(crate) fn can_capture(&self) -> bool {
        self.opening.is_none() && self.sequence == 0 && !self.failed && !self.applying
    }
}

#[derive(Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
struct Envelope {
    sequence: u64,
    base_version: String,
    op: WireOp,
    #[serde(default)]
    expected_outcome: Option<Value>,
}

#[derive(Deserialize)]
#[serde(deny_unknown_fields)]
struct WireOp {
    method: String,
    args: Vec<Value>,
}

#[derive(Serialize)]
#[serde(rename_all = "camelCase")]
struct Reply {
    sequence: u64,
    revision: u64,
    version: String,
    engine_version: String,
    consumed: bool,
    outcome: Value,
}

enum Op {
    InsertText(String, u32, String, TextStyle),
    DeleteText(String, u32, u32),
    FormatText(String, u32, u32, TextStylePatch),
    InsertParagraphBreak(String, u32),
    SetParagraphAlignment(String, u32, u32, Option<String>),
    InsertSlide(u32, Option<String>),
    DeleteSlide(String),
    MoveSlide(String, u32),
    SetSlideNotes(String, String),
    AddTextBox(String, ShapeDraft),
    AddShape(String, PresetShapeDraft),
    AddPicture(String, PictureDraft),
    RemoveShape(String, String),
    MoveShape(String, String, i64, i64),
    ResizeShape(String, String, i64, i64),
    SetShapeRect(String, String, ShapeRect),
    SetShapeFill(String, String, Option<String>),
    SetShapeStroke(String, String, ShapeStroke),
    SetShapeAdjust(String, String, BTreeMap<String, f64>),
    BringShapeToFront(String, String),
    SendShapeToBack(String, String),
    BringShapeForward(String, String),
    SendShapeBackward(String, String),
    AddComment(String, CommentDraft),
    ReplyToComment(String, ReplyDraft),
    SetCommentStatus(String, bool),
    SetCommentPosition(String, Position),
    RemoveComment(String),
    SetCommentFlavor(CommentFlavor),
    Propose(ProposalRequest),
    AcceptProposal(String, bool),
    RejectProposal(String),
    ApplyEdits(EditRequest),
    AddUndoBoundary,
    Undo,
    Redo,
}

#[derive(Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
struct CommentDraft {
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
#[serde(rename_all = "camelCase", deny_unknown_fields)]
struct ReplyDraft {
    author: String,
    #[serde(default)]
    initials: String,
    text: String,
    created: String,
}

#[derive(Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
struct Position {
    x_emu: i64,
    y_emu: i64,
}

#[derive(Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
struct PictureJson {
    name: String,
    rect: ShapeRect,
    content_type: String,
    media_base64: String,
}

#[derive(Default, Deserialize)]
#[serde(deny_unknown_fields)]
struct AcceptOptions {
    #[serde(default)]
    force: bool,
}

fn tuple<T: DeserializeOwned>(args: Vec<Value>, min: usize, max: usize) -> Result<T, PeerError> {
    if args.len() < min || args.len() > max {
        return Err(PeerError::new("arguments", "wrong argument tuple length"));
    }
    let mut args = args;
    args.resize(max, Value::Null);
    Ok(serde_json::from_value(Value::Array(args))?)
}

fn validate_numbers(value: &Value, field: &str) -> Result<(), PeerError> {
    match value {
        Value::Number(number) => {
            let finite = number.as_f64().is_some_and(f64::is_finite);
            let integer_field = matches!(field, "x" | "y" | "width" | "height" | "xEmu" | "yEmu");
            if !finite
                || (integer_field
                    && !number
                        .as_i64()
                        .is_some_and(|n| n.unsigned_abs() <= MAX_SAFE_INTEGER))
            {
                return Err(PeerError::new(
                    "arguments",
                    format!("invalid numeric {field}"),
                ));
            }
        }
        Value::Array(array) => {
            for value in array {
                validate_numbers(value, field)?;
            }
        }
        Value::Object(object) => {
            for (field, value) in object {
                validate_numbers(value, field)?;
            }
        }
        _ => {}
    }
    Ok(())
}

fn canonical_numbers(value: &mut Value) {
    match value {
        Value::Number(number) if number.is_f64() => {
            let float = number.as_f64().unwrap_or(f64::NAN);
            if float.fract() == 0.0 && float >= -(2_f64.powi(63)) && float < 2_f64.powi(63) {
                *value = Value::from(float as i64);
            } else if float.fract() == 0.0 && float >= 0.0 && float < 2_f64.powi(64) {
                *value = Value::from(float as u64);
            }
        }
        Value::Array(items) => items.iter_mut().for_each(canonical_numbers),
        Value::Object(object) => object.values_mut().for_each(canonical_numbers),
        _ => {}
    }
}

fn fields(value: &Value, allowed: &[&str]) -> Result<(), PeerError> {
    let object = value
        .as_object()
        .ok_or_else(|| PeerError::new("arguments", "expected an object"))?;
    if object
        .keys()
        .any(|field| !allowed.contains(&field.as_str()))
    {
        return Err(PeerError::new("arguments", "unknown nested argument field"));
    }
    Ok(())
}

fn validate_nested(value: &Value) -> Result<(), PeerError> {
    match value {
        Value::Array(values) => {
            for value in values {
                validate_nested(value)?;
            }
        }
        Value::Object(object) => {
            for (key, value) in object {
                if !value.is_null() {
                    match key.as_str() {
                        "style" => fields(
                            value,
                            &[
                                "bold",
                                "italic",
                                "fontSizePt",
                                "color",
                                "fontFamily",
                                "underline",
                                "spacingPt",
                                "baselinePct",
                                "caps",
                            ],
                        )?,
                        "patch" => fields(
                            value,
                            &[
                                "bold",
                                "italic",
                                "fontSizePt",
                                "color",
                                "fontFamily",
                                "underline",
                                "spacingPt",
                                "baselinePct",
                            ],
                        )?,
                        "rect" => fields(value, &["x", "y", "width", "height"])?,
                        "stroke" => fields(value, &["color", "widthPt"])?,
                        _ => {}
                    }
                }
                validate_nested(value)?;
            }
        }
        _ => {}
    }
    Ok(())
}

impl WireOp {
    fn parse(self) -> Result<Op, PeerError> {
        validate_numbers(&Value::Array(self.args.clone()), "")?;
        validate_nested(&Value::Array(self.args.clone()))?;
        match self.method.as_str() {
            "insertText" if self.args.get(3).is_some_and(|value| !value.is_null()) => fields(
                &self.args[3],
                &[
                    "bold",
                    "italic",
                    "fontSizePt",
                    "color",
                    "fontFamily",
                    "underline",
                    "spacingPt",
                    "baselinePct",
                    "caps",
                ],
            )?,
            "formatText" if self.args.len() >= 4 => fields(
                &self.args[3],
                &[
                    "bold",
                    "italic",
                    "fontSizePt",
                    "color",
                    "fontFamily",
                    "underline",
                    "spacingPt",
                    "baselinePct",
                ],
            )?,
            "addTextBox" if self.args.len() >= 2 => {
                fields(&self.args[1], &["name", "rect", "text", "style"])?
            }
            "addShape" if self.args.len() >= 2 => {
                fields(&self.args[1], &["name", "geometry", "rect", "fill"])?
            }
            "setShapeRect" if self.args.len() >= 3 => {
                fields(&self.args[2], &["x", "y", "width", "height"])?
            }
            "setShapeStroke" if self.args.len() >= 3 => {
                fields(&self.args[2], &["color", "widthPt"])?
            }
            _ => {}
        }
        let a = self.args;
        Ok(match self.method.as_str() {
            "insertText" => {
                let (id, at, text, style): (String, u32, String, Option<TextStyle>) =
                    tuple(a, 3, 4)?;
                Op::InsertText(id, at, text, style.unwrap_or_default())
            }
            "deleteText" => {
                let (id, start, end) = tuple(a, 3, 3)?;
                Op::DeleteText(id, start, end)
            }
            "formatText" => {
                let (id, start, end, patch) = tuple(a, 4, 4)?;
                Op::FormatText(id, start, end, patch)
            }
            "insertParagraphBreak" => {
                let (id, at) = tuple(a, 2, 2)?;
                Op::InsertParagraphBreak(id, at)
            }
            "setParagraphAlignment" => {
                let (id, start, end, align) = tuple(a, 4, 4)?;
                Op::SetParagraphAlignment(id, start, end, align)
            }
            "insertSlide" => {
                let (index, layout) = tuple(a, 1, 2)?;
                Op::InsertSlide(index, layout)
            }
            "deleteSlide" => {
                let (id,) = tuple(a, 1, 1)?;
                Op::DeleteSlide(id)
            }
            "moveSlide" => {
                let (id, index) = tuple(a, 2, 2)?;
                Op::MoveSlide(id, index)
            }
            "setSlideNotes" => {
                let (id, text) = tuple(a, 2, 2)?;
                Op::SetSlideNotes(id, text)
            }
            "addTextBox" => {
                let (id, draft) = tuple(a, 2, 2)?;
                Op::AddTextBox(id, draft)
            }
            "addShape" => {
                let (id, draft) = tuple(a, 2, 2)?;
                Op::AddShape(id, draft)
            }
            "addPicture" => {
                let (id, draft): (String, PictureJson) = tuple(a, 2, 2)?;
                let media_bytes = base64::engine::general_purpose::STANDARD
                    .decode(&draft.media_base64)
                    .map_err(|error| PeerError::new("arguments", error))?;
                if media_bytes.len() > 8 * 1024 * 1024 {
                    return Err(PeerError::new("arguments", "image exceeds engine limit"));
                }
                Op::AddPicture(
                    id,
                    PictureDraft {
                        name: draft.name,
                        rect: draft.rect,
                        content_type: draft.content_type,
                        media_bytes,
                    },
                )
            }
            "removeShape" => {
                let (slide, shape) = tuple(a, 2, 2)?;
                Op::RemoveShape(slide, shape)
            }
            "moveShape" => {
                let (slide, shape, x, y) = tuple(a, 4, 4)?;
                Op::MoveShape(slide, shape, x, y)
            }
            "resizeShape" => {
                let (slide, shape, w, h) = tuple(a, 4, 4)?;
                Op::ResizeShape(slide, shape, w, h)
            }
            "setShapeRect" => {
                let (slide, shape, rect) = tuple(a, 3, 3)?;
                Op::SetShapeRect(slide, shape, rect)
            }
            "setShapeFill" => {
                let (slide, shape, color) = tuple(a, 3, 3)?;
                Op::SetShapeFill(slide, shape, color)
            }
            "setShapeStroke" => {
                let (slide, shape, stroke) = tuple(a, 3, 3)?;
                Op::SetShapeStroke(slide, shape, stroke)
            }
            "setShapeAdjust" => {
                let (slide, shape, adj) = tuple(a, 3, 3)?;
                Op::SetShapeAdjust(slide, shape, adj)
            }
            "bringShapeToFront" => {
                let (slide, shape) = tuple(a, 2, 2)?;
                Op::BringShapeToFront(slide, shape)
            }
            "sendShapeToBack" => {
                let (slide, shape) = tuple(a, 2, 2)?;
                Op::SendShapeToBack(slide, shape)
            }
            "bringShapeForward" => {
                let (slide, shape) = tuple(a, 2, 2)?;
                Op::BringShapeForward(slide, shape)
            }
            "sendShapeBackward" => {
                let (slide, shape) = tuple(a, 2, 2)?;
                Op::SendShapeBackward(slide, shape)
            }
            "addComment" => {
                let (slide, draft) = tuple(a, 2, 2)?;
                Op::AddComment(slide, draft)
            }
            "replyToComment" => {
                let (id, draft) = tuple(a, 2, 2)?;
                Op::ReplyToComment(id, draft)
            }
            "setCommentStatus" => {
                let (id, resolved) = tuple(a, 2, 2)?;
                Op::SetCommentStatus(id, resolved)
            }
            "setCommentPosition" => {
                let (id, pos) = tuple(a, 2, 2)?;
                Op::SetCommentPosition(id, pos)
            }
            "removeComment" => {
                let (id,) = tuple(a, 1, 1)?;
                Op::RemoveComment(id)
            }
            "setCommentFlavor" => {
                let (flavor,) = tuple(a, 1, 1)?;
                Op::SetCommentFlavor(flavor)
            }
            "propose" => {
                let (agent_id, note, edits): (String, Option<String>, Vec<ProposalEdit>) =
                    tuple(a, 3, 3)?;
                Op::Propose(ProposalRequest {
                    agent_id,
                    note,
                    edits,
                })
            }
            "acceptProposal" => {
                let (id, options): (String, Option<AcceptOptions>) = tuple(a, 1, 2)?;
                Op::AcceptProposal(id, options.unwrap_or_default().force)
            }
            "rejectProposal" => {
                let (id,) = tuple(a, 1, 1)?;
                Op::RejectProposal(id)
            }
            "applyEdits" => {
                let (request,) = tuple(a, 1, 1)?;
                Op::ApplyEdits(request)
            }
            "addUndoBoundary" | "undo" | "redo" => {
                if !a.is_empty() {
                    return Err(PeerError::new("arguments", "history takes no arguments"));
                }
                match self.method.as_str() {
                    "undo" => Op::Undo,
                    "redo" => Op::Redo,
                    _ => Op::AddUndoBoundary,
                }
            }
            _ => return Err(PeerError::new("method", "unsupported replay method")),
        })
    }
}

fn value<T: Serialize>(receipt: T) -> Result<Value, PeerError> {
    Ok(serde_json::to_value(receipt)?)
}

impl Op {
    fn dispatch(&self, session: &DeckSession) -> Result<Value, PeerError> {
        let ctx = EditCtx::local("host");
        match self {
            Self::InsertText(id, at, text, style) => {
                value(session.insert_text(&ctx, id, *at, text, style)?)
            }
            Self::DeleteText(id, start, end) => value(session.delete_text(&ctx, id, *start, *end)?),
            Self::FormatText(id, start, end, patch) => {
                value(session.format_text(&ctx, id, *start, *end, patch)?)
            }
            Self::InsertParagraphBreak(id, at) => {
                value(session.insert_paragraph_break(&ctx, id, *at)?)
            }
            Self::SetParagraphAlignment(id, start, end, align) => {
                value(session.set_paragraph_alignment(&ctx, id, *start, *end, align.as_deref())?)
            }
            Self::InsertSlide(at, layout) => {
                value(session.insert_slide(&ctx, *at, layout.as_deref())?)
            }
            Self::DeleteSlide(id) => value(session.delete_slide(&ctx, id)?),
            Self::MoveSlide(id, at) => value(session.move_slide(&ctx, id, *at)?),
            Self::SetSlideNotes(id, text) => {
                session.set_slide_notes(&ctx, id, text)?;
                Ok(Value::Null)
            }
            Self::AddTextBox(id, draft) => value(session.add_text_box(&ctx, id, draft)?),
            Self::AddShape(id, draft) => value(session.add_shape(&ctx, id, draft)?),
            Self::AddPicture(id, draft) => value(session.add_picture(&ctx, id, draft)?),
            Self::RemoveShape(slide, shape) => value(session.remove_shape(&ctx, slide, shape)?),
            Self::MoveShape(slide, shape, x, y) => {
                value(session.move_shape(&ctx, slide, shape, *x, *y)?)
            }
            Self::ResizeShape(slide, shape, w, h) => {
                value(session.resize_shape(&ctx, slide, shape, *w, *h)?)
            }
            Self::SetShapeRect(slide, shape, rect) => {
                value(session.set_shape_rect(&ctx, slide, shape, *rect)?)
            }
            Self::SetShapeFill(slide, shape, color) => {
                value(session.set_shape_fill(&ctx, slide, shape, color.as_deref())?)
            }
            Self::SetShapeStroke(slide, shape, stroke) => {
                value(session.set_shape_stroke(&ctx, slide, shape, stroke)?)
            }
            Self::SetShapeAdjust(slide, shape, adj) => {
                value(session.set_shape_adjust(&ctx, slide, shape, adj)?)
            }
            Self::BringShapeToFront(slide, shape) => {
                value(session.bring_to_front(&ctx, slide, shape)?)
            }
            Self::SendShapeToBack(slide, shape) => value(session.send_to_back(&ctx, slide, shape)?),
            Self::BringShapeForward(slide, shape) => {
                value(session.bring_forward(&ctx, slide, shape)?)
            }
            Self::SendShapeBackward(slide, shape) => {
                value(session.send_backward(&ctx, slide, shape)?)
            }
            Self::AddComment(slide, d) => value(session.add_comment(
                &ctx,
                slide,
                &d.author,
                &d.initials,
                &d.text,
                &d.created,
                d.x_emu,
                d.y_emu,
            )?),
            Self::ReplyToComment(id, d) => value(session.reply_to_comment(
                &ctx,
                id,
                &d.author,
                &d.initials,
                &d.text,
                &d.created,
            )?),
            Self::SetCommentStatus(id, status) => {
                value(session.set_comment_status(&ctx, id, *status)?)
            }
            Self::SetCommentPosition(id, p) => {
                value(session.set_comment_position(&ctx, id, p.x_emu, p.y_emu)?)
            }
            Self::RemoveComment(id) => value(session.remove_comment(&ctx, id)?),
            Self::SetCommentFlavor(flavor) => value(session.set_comment_flavor(&ctx, *flavor)?),
            Self::Propose(request) => {
                let proposal = session.propose(request.clone())?;
                Ok(
                    json!({ "proposalId": proposal.id, "changedTargets": proposal.changes.iter().map(|change| change.key()).collect::<BTreeSet<_>>() }),
                )
            }
            Self::AcceptProposal(id, force) => value(session.accept_proposal_compact(id, *force)?),
            Self::RejectProposal(id) => {
                Ok(json!({ "proposalId": id, "rejected": session.reject_proposal(id) }))
            }
            Self::ApplyEdits(request) => Ok(serde_json::from_str(&crate::outcome_json(
                &session.apply_edits(request)?,
            )?)?),
            Self::AddUndoBoundary => {
                session.add_undo_barrier();
                Ok(Value::Null)
            }
            Self::Undo => Ok(json!({ "applied": session.undo() })),
            Self::Redo => Ok(json!({ "applied": session.redo() })),
        }
    }

    fn targets(&self, result: &Value, applied: bool) -> BTreeSet<String> {
        let mut targets = BTreeSet::new();
        if !applied {
            return targets;
        }
        for key in ["slideId", "shapeId", "storyId", "commentId"] {
            if let Some(id) = result.get(key).and_then(Value::as_str) {
                targets.insert(id.to_owned());
            }
        }
        for key in ["changedTargets", "changedSlides", "changedStories"] {
            if let Some(ids) = result.get(key).and_then(Value::as_array) {
                targets.extend(ids.iter().filter_map(Value::as_str).map(str::to_owned));
            }
        }
        match self {
            Self::SetSlideNotes(id, _) => {
                targets.insert(id.clone());
            }
            Self::Undo | Self::Redo | Self::SetCommentFlavor(_) => {
                targets.insert("deck".into());
            }
            _ => {}
        }
        targets
    }
}

impl DeckSession {
    fn replay_diagnostics(&self) -> (u64, u64, (u64, usize), (usize, usize)) {
        (
            self.epoch(),
            self.id_counter.load(Ordering::Relaxed),
            self.proposals.borrow().counters(),
            self.undo.borrow().diagnostics(),
        )
    }

    #[doc(hidden)]
    pub fn replay_json(&self, input: &str) -> Result<String, PeerError> {
        if input.len() > crate::MAX_REQUEST_BYTES {
            return Err(PeerError::new("arguments", "envelope exceeds engine limit"));
        }
        let mut envelope: Envelope = serde_json::from_str(input)?;
        if let Some(expected) = envelope.expected_outcome.as_mut() {
            canonical_numbers(expected);
            fields(
                expected,
                &["result", "applied", "changedTargets", "canUndo", "canRedo"],
            )?;
            if expected.get("result").is_none()
                || ["applied", "canUndo", "canRedo"]
                    .iter()
                    .any(|key| !expected.get(*key).is_some_and(Value::is_boolean))
                || !expected
                    .get("changedTargets")
                    .and_then(Value::as_array)
                    .is_some_and(|targets| targets.iter().all(Value::is_string))
            {
                return Err(PeerError::new("arguments", "malformed expected outcome"));
            }
        }
        let op = envelope.op.parse()?;
        self.undo.borrow().assert_replay_capture()?;
        {
            let state = self.replay_state.borrow();
            if !state.ready || state.failed || state.applying || state.opening.is_some() {
                return Err(PeerError::new("stage", "replay lane is not ready"));
            }
            if envelope.sequence > MAX_SAFE_INTEGER
                || state.sequence.checked_add(1) != Some(envelope.sequence)
            {
                return Err(PeerError::new(
                    "sequence",
                    "replay sequence is not the next safe integer",
                ));
            }
            if envelope.base_version != self.version().as_str() {
                return Err(PeerError::new("version", "replay base version differs"));
            }
        }
        let before = self.replay_diagnostics();
        self.replay_state.borrow_mut().applying = true;
        let dispatched = op.dispatch(self);
        self.replay_state.borrow_mut().applying = false;
        let result = match dispatched {
            Ok(result) => result,
            Err(error) => {
                if self.replay_diagnostics() != before {
                    self.replay_state.borrow_mut().failed = true;
                }
                return Err(error);
            }
        };
        let refused = result.get("ok") == Some(&Value::Bool(false));
        let applied = self.epoch() != before.0;
        let mut outcome = json!({
            "result": result,
            "applied": applied,
            "changedTargets": op.targets(&result, applied),
            "canUndo": self.can_undo(),
            "canRedo": self.can_redo(),
        });
        canonical_numbers(&mut outcome);
        let mut state = self.replay_state.borrow_mut();
        if !refused {
            state.sequence = envelope.sequence;
            state.revision += 1;
        }
        if envelope
            .expected_outcome
            .as_ref()
            .is_some_and(|expected| expected != &outcome)
        {
            state.failed = true;
            return Err(PeerError::new(
                "outcomeMismatch",
                "canonical replay outcome differs; retain peer for recovery",
            ));
        }
        Ok(serde_json::to_string(&Reply {
            sequence: state.sequence,
            revision: state.revision,
            version: self.version().to_string(),
            engine_version: self.version().to_string(),
            consumed: !refused,
            outcome,
        })?)
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::peer::tests::{apply_pair, pair, replay, source};
    use crate::{ProposalEdit, ProposalRequest};

    fn assert_observable(worker: &DeckSession, peer: &DeckSession, prefix: &str) {
        let snapshot = worker.snapshot().unwrap();
        assert_eq!(
            worker.save().unwrap(),
            peer.save().unwrap(),
            "save bytes at {prefix}"
        );
        assert_eq!(
            snapshot,
            peer.snapshot().unwrap(),
            "stories and paragraph ids at {prefix}"
        );
        assert_eq!(
            worker.proposals().unwrap(),
            peer.proposals().unwrap(),
            "full proposals at {prefix}"
        );
        assert_eq!(
            worker.replay_diagnostics(),
            peer.replay_diagnostics(),
            "counters and history at {prefix}"
        );
        assert_eq!(
            worker.undo.borrow().stack_clock_counts(),
            peer.undo.borrow().stack_clock_counts(),
            "history entry clock counts at {prefix}"
        );
        assert_eq!(worker.version(), peer.version(), "versions at {prefix}");
        for slide in &snapshot.slides {
            for shape in &slide.shapes {
                for story in &shape.text_stories {
                    for index in 0..story.length {
                        let worker_anchor = worker.anchor_caret(&story.id, index).unwrap();
                        let peer_anchor = peer.anchor_caret(&story.id, index).unwrap();
                        assert_eq!(
                            worker.resolve_caret_anchor(&worker_anchor),
                            Some(index),
                            "worker caret at {prefix}/{index}"
                        );
                        assert_eq!(
                            peer.resolve_caret_anchor(&peer_anchor),
                            Some(index),
                            "peer caret at {prefix}/{index}"
                        );
                    }
                }
            }
        }
    }

    fn assert_ordered_history(worker: &DeckSession, peer: &DeckSession, prefix: &str) {
        assert_eq!(
            worker.undo.borrow().stack_diagnostics(),
            peer.undo.borrow().stack_diagnostics(),
            "history ids before restoration at {prefix}"
        );
        for slide in worker.snapshot().unwrap().slides {
            for shape in slide.shapes {
                for story in shape.text_stories {
                    for index in 0..story.length {
                        let anchor = worker.anchor_caret(&story.id, index).unwrap();
                        assert_eq!(
                            anchor,
                            peer.anchor_caret(&story.id, index).unwrap(),
                            "caret ids before restoration at {prefix}/{index}"
                        );
                        assert_eq!(
                            peer.resolve_caret_anchor(&anchor),
                            Some(index),
                            "shared caret before restoration at {prefix}/{index}"
                        );
                    }
                }
            }
        }
    }

    fn history_paths(worker: &DeckSession, peer: &DeckSession, prefix: &str) {
        let final_state = worker.snapshot().unwrap();
        let final_bytes = worker.save().unwrap();
        let mut undos = 0;
        while worker.can_undo() {
            assert!(undos < 256, "undo failed to terminate at {prefix}");
            let result = apply_pair(worker, peer, &json!({"method":"undo", "args":[]}));
            assert_eq!(
                result["outcome"]["result"]["applied"], true,
                "undo at {prefix}/{undos}"
            );
            assert_observable(worker, peer, &format!("{prefix}/undo/{undos}"));
            undos += 1;
        }
        let result = apply_pair(worker, peer, &json!({"method":"undo", "args":[]}));
        assert_eq!(
            result["outcome"]["result"]["applied"], false,
            "exhausted undo at {prefix}"
        );
        let mut redos = 0;
        while worker.can_redo() {
            assert!(redos < 256, "redo failed to terminate at {prefix}");
            let result = apply_pair(worker, peer, &json!({"method":"redo", "args":[]}));
            assert_eq!(
                result["outcome"]["result"]["applied"], true,
                "redo at {prefix}/{redos}"
            );
            assert_observable(worker, peer, &format!("{prefix}/redo/{redos}"));
            redos += 1;
        }
        assert_eq!(undos, redos, "history path lengths at {prefix}");
        let result = apply_pair(worker, peer, &json!({"method":"redo", "args":[]}));
        assert_eq!(
            result["outcome"]["result"]["applied"], false,
            "exhausted redo at {prefix}"
        );
        assert_eq!(
            worker.snapshot().unwrap(),
            final_state,
            "redo restores content at {prefix}"
        );
        assert_eq!(
            worker.save().unwrap(),
            final_bytes,
            "redo restores bytes at {prefix}"
        );
    }

    fn operations(session: &DeckSession) -> Vec<Value> {
        let snapshot = session.snapshot().unwrap();
        let slide = &snapshot.slides[0].id;
        let shape = &snapshot.slides[0].shapes[0].id;
        let story = &snapshot.slides[0].shapes[0].text_stories[0].id;
        let rect = json!({"x":100,"y":200,"width":300000,"height":400000});
        let client = session.client_id();
        let textbox = format!("shape:{client}:1");
        let new_slide = format!("slide:{client}:0");
        let comment = format!("comment:{client}:6");
        vec![
            json!({"method":"applyEdits","args":[{"expectVersion":session.version(),"history":"separate","steps":[{"op":"setSlideNotes","target":{"slideId":slide},"text":"batch notes"}]}]}),
            json!({"method":"insertSlide","args":[1]}),
            json!({"method":"moveSlide","args":[new_slide,0]}),
            json!({"method":"addUndoBoundary","args":[]}),
            json!({"method":"setSlideNotes","args":[slide,"speaker notes"]}),
            json!({"method":"addTextBox","args":[slide,{"name":"New text","rect":rect,"text":"Box","style":{}}]}),
            json!({"method":"addShape","args":[slide,{"name":"Rectangle","geometry":"rect","rect":rect,"fill":"#123456"}]}),
            json!({"method":"addPicture","args":[slide,{"name":"Image","rect":rect,"contentType":"image/png","mediaBase64":"iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+a1ioAAAAASUVORK5CYII="}]}),
            json!({"method":"insertText","args":[story,0,"🙂",{"bold":true,"italic":true,"fontSizePt":18,"color":"#ABCDEF","fontFamily":"A","underline":"sng","spacingPt":2,"baselinePct":5}]}),
            json!({"method":"deleteText","args":[story,0,2]}),
            json!({"method":"formatText","args":[story,0,3,{"bold":false,"italic":true,"fontSizePt":20,"color":"#FF0000","spacingPt":1,"baselinePct":0}]}),
            json!({"method":"insertParagraphBreak","args":[story,2]}),
            json!({"method":"setParagraphAlignment","args":[story,0,3,"ctr"]}),
            json!({"method":"moveShape","args":[slide,shape,500,600]}),
            json!({"method":"resizeShape","args":[slide,shape,500000,600000]}),
            json!({"method":"setShapeRect","args":[slide,shape,rect]}),
            json!({"method":"setShapeFill","args":[slide,shape,"#112233"]}),
            json!({"method":"setShapeStroke","args":[slide,shape,{"color":"#332211","widthPt":2}]}),
            json!({"method":"setShapeAdjust","args":[slide,shape,{"adj":12000,"adj1":100}]}),
            json!({"method":"bringShapeToFront","args":[slide,shape]}),
            json!({"method":"sendShapeToBack","args":[slide,shape]}),
            json!({"method":"bringShapeForward","args":[slide,shape]}),
            json!({"method":"sendShapeBackward","args":[slide,shape]}),
            json!({"method":"removeShape","args":[slide,textbox]}),
            json!({"method":"setCommentFlavor","args":["modern"]}),
            json!({"method":"addComment","args":[slide,{"author":"Human","initials":"H","text":"Comment","created":"2026-10-05T11:22:33.456Z","xEmu":123,"yEmu":456}]}),
            json!({"method":"replyToComment","args":[comment,{"author":"Peer","text":"Reply","created":"2026-10-05T11:22:33.456Z"}]}),
            json!({"method":"setCommentPosition","args":[comment,{"xEmu":789,"yEmu":123}]}),
            json!({"method":"setCommentStatus","args":[comment,true]}),
            json!({"method":"removeComment","args":[comment]}),
            json!({"method":"propose","args":["agent",null,[{"type":"setSlideNotes","slideId":slide,"text":"Proposed notes"}]]}),
            json!({"method":"acceptProposal","args":["p1",{"force":false}]}),
            json!({"method":"propose","args":["agent","reject this",[{"type":"setShapeFill","slideId":slide,"shapeId":shape,"color":"#FFFFFF"}]]}),
            json!({"method":"rejectProposal","args":["p2"]}),
            json!({"method":"rejectProposal","args":["absent"]}),
            json!({"method":"deleteSlide","args":[new_slide]}),
            json!({"method":"addUndoBoundary","args":[]}),
        ]
    }

    #[test]
    fn every_replay_prefix_has_equal_observable_state() {
        let (base, _) = pair();
        let ops = operations(&base);
        for end in 0..=ops.len() {
            let (worker, peer) = pair();
            let ops = operations(&worker);
            for (index, op) in ops[..end].iter().enumerate() {
                apply_pair(&worker, &peer, op);
                assert_observable(
                    &worker,
                    &peer,
                    &format!("prefix {end}, operation {index}: {op}"),
                );
                assert_ordered_history(&worker, &peer, &format!("prefix {end}, operation {index}"));
            }
            let label = format!(
                "prefix {end}: {}",
                end.checked_sub(1)
                    .map(|index| ops[index].to_string())
                    .unwrap_or_else(|| "baseline".into())
            );
            history_paths(&worker, &peer, &label);
        }
    }

    #[test]
    fn full_undo_and_redo_paths_match() {
        let (worker, peer) = pair();
        for op in operations(&worker) {
            apply_pair(&worker, &peer, &op);
            assert_observable(&worker, &peer, &format!("full replay: {op}"));
            assert_ordered_history(&worker, &peer, &format!("full replay: {op}"));
        }
        history_paths(&worker, &peer, "full replay");
    }

    #[test]
    fn history_is_independent_of_dispatch_delay() {
        let (worker, peer) = pair();
        let slide = worker.slide_ids().unwrap()[0].clone();
        for (index, text) in ["one", "two", "three"].iter().enumerate() {
            let op = json!({"method":"setSlideNotes","args":[slide,text]});
            let result = replay(&worker, &op, None);
            std::thread::sleep(std::time::Duration::from_millis(550));
            assert_eq!(result, replay(&peer, &op, Some(result["outcome"].clone())));
            assert_eq!(
                worker.undo.borrow().diagnostics(),
                (1, 0),
                "manual group {index}"
            );
        }
        apply_pair(
            &worker,
            &peer,
            &json!({"method":"addUndoBoundary","args":[]}),
        );
        apply_pair(
            &worker,
            &peer,
            &json!({"method":"setSlideNotes","args":[slide,"next group"]}),
        );
        assert_eq!(worker.undo.borrow().diagnostics(), (2, 0));
        history_paths(&worker, &peer, "delayed dispatch");
    }

    #[test]
    fn proposal_preview_accept_reject_preserve_identity() {
        let (worker, peer) = pair();
        let slide = worker.slide_ids().unwrap()[0].clone();
        let before = worker.replay_diagnostics();
        let propose = json!({"method":"propose","args":["agent",null,[{"type":"setSlideNotes","slideId":slide,"text":"proposal"}]]});
        apply_pair(&worker, &peer, &propose);
        let live_update = worker.encode_state_as_update_v1();
        assert_eq!(
            worker.preview_proposal("p1").unwrap(),
            peer.preview_proposal("p1").unwrap()
        );
        assert_eq!(worker.encode_state_as_update_v1(), live_update);
        assert_eq!(worker.replay_diagnostics().1, before.1);
        apply_pair(
            &worker,
            &peer,
            &json!({"method":"acceptProposal","args":["p1"]}),
        );
        assert!(worker.proposals().unwrap().is_empty());
        apply_pair(&worker, &peer, &propose);
        apply_pair(
            &worker,
            &peer,
            &json!({"method":"rejectProposal","args":["p2"]}),
        );
        assert_observable(&worker, &peer, "proposal preview/accept/reject");
        let baseline = DeckSession::open_replay_baseline(&source(), Some(71), None).unwrap();
        let proposal = baseline
            .propose(ProposalRequest {
                agent_id: "agent".into(),
                note: None,
                edits: vec![ProposalEdit::SetSlideNotes {
                    slide_id: slide,
                    text: "discard".into(),
                }],
            })
            .unwrap();
        assert!(baseline.reject_proposal(&proposal.id));
        let identity = baseline.peer_identity(Vec::new()).unwrap();
        let hydrated = DeckSession::open_peer_deck(&source(), &identity, None).unwrap();
        hydrated.register_peer_fonts(Vec::new()).unwrap();
        hydrated.adopt_peer_identity().unwrap();
        assert_eq!(
            baseline.proposals.borrow().counters(),
            hydrated.proposals.borrow().counters()
        );
        assert_eq!(
            replay(&baseline, &propose, None),
            replay(&hydrated, &propose, None)
        );
    }

    #[test]
    fn stale_proposal_acceptance_is_atomic() {
        let (worker, peer) = pair();
        let slide = worker.slide_ids().unwrap()[0].clone();
        apply_pair(
            &worker,
            &peer,
            &json!({"method":"propose","args":["agent",null,[{"type":"setSlideNotes","slideId":slide,"text":"proposal"}]]}),
        );
        apply_pair(
            &worker,
            &peer,
            &json!({"method":"setSlideNotes","args":[slide,"human change"]}),
        );
        let before = worker.replay_diagnostics();
        let proposals = worker.proposals().unwrap();
        let bytes = worker.save().unwrap();
        let sequence = worker.replay_state.borrow().sequence;
        let envelope = json!({"sequence":sequence+1,"baseVersion":worker.version(),"op":{"method":"acceptProposal","args":["p1"]}}).to_string();
        assert!(matches!(worker.replay_json(&envelope), Err(error) if error.code == "proposal"));
        assert!(peer.replay_json(&envelope).is_err());
        assert_eq!(worker.replay_diagnostics(), before);
        assert_eq!(worker.proposals().unwrap(), proposals);
        assert_eq!(worker.save().unwrap(), bytes);
        assert_eq!(worker.replay_state.borrow().sequence, sequence);
        assert_observable(&worker, &peer, "stale acceptance refusal");
        apply_pair(
            &worker,
            &peer,
            &json!({"method":"acceptProposal","args":["p1",{"force":true}]}),
        );
        assert!(worker.proposals().unwrap().is_empty());
        assert_observable(&worker, &peer, "forced acceptance");
    }

    #[test]
    fn media_and_comment_timestamps_match() {
        let (worker, peer) = pair();
        let slide = worker.slide_ids().unwrap()[0].clone();
        let picture = operations(&worker)
            .into_iter()
            .find(|op| op["method"] == "addPicture")
            .unwrap();
        apply_pair(&worker, &peer, &picture);
        apply_pair(
            &worker,
            &peer,
            &json!({"method":"setCommentFlavor","args":["modern"]}),
        );
        let created = "2026-10-05T11:22:33.456Z";
        let comment = apply_pair(
            &worker,
            &peer,
            &json!({"method":"addComment","args":[slide,{"author":"Human","text":"Comment","created":created,"xEmu":100,"yEmu":200}]}),
        );
        let id = comment["outcome"]["result"]["commentId"].as_str().unwrap();
        apply_pair(
            &worker,
            &peer,
            &json!({"method":"replyToComment","args":[id,{"author":"Peer","text":"Reply","created":created}]}),
        );
        apply_pair(
            &worker,
            &peer,
            &json!({"method":"setCommentPosition","args":[id,{"xEmu":300,"yEmu":400}]}),
        );
        apply_pair(
            &worker,
            &peer,
            &json!({"method":"setCommentStatus","args":[id,true]}),
        );
        assert_eq!(worker.comments().unwrap().len(), 2);
        assert!(
            worker
                .comments()
                .unwrap()
                .iter()
                .all(|comment| comment.created.as_deref() == Some(created))
        );
        let snapshot = worker.snapshot().unwrap();
        let media = snapshot.slides[0]
            .shapes
            .iter()
            .find_map(|shape| shape.pending_media.as_ref())
            .unwrap();
        assert_eq!(media.content_type, "image/png");
        assert_eq!(
            media.base64,
            picture["args"][1]["mediaBase64"].as_str().unwrap()
        );
        assert_observable(&worker, &peer, "media and timestamps");
        apply_pair(
            &worker,
            &peer,
            &json!({"method":"removeComment","args":[id]}),
        );
        assert!(worker.comments().unwrap().is_empty());
        history_paths(&worker, &peer, "comment removal with replies");
    }

    #[test]
    fn rejected_creators_preserve_allocator_and_history() {
        let (worker, peer) = pair();
        let before = worker.replay_diagnostics();
        for op in [
            json!({"method":"insertSlide","args":[99]}),
            json!({"method":"addTextBox","args":["missing",{"name":"Box","rect":{"x":0,"y":0,"width":1,"height":1},"text":"text","style":{}}]}),
            json!({"method":"addComment","args":["missing",{"author":"Human","text":"text","created":"2026-10-05T00:00:00Z"}]}),
            json!({"method":"replyToComment","args":["missing",{"author":"Human","text":"text","created":"2026-10-05T00:00:00Z"}]}),
        ] {
            for session in [&worker, &peer] {
                let input =
                    json!({"sequence":1,"baseVersion":session.version(),"op":op}).to_string();
                assert!(session.replay_json(&input).is_err(), "must refuse {op}");
                assert_eq!(
                    session.replay_diagnostics(),
                    before,
                    "refused creator changed counters/history: {op}"
                );
                assert_eq!(
                    session.replay_state.borrow().sequence,
                    0,
                    "refused sequence: {op}"
                );
            }
        }
        apply_pair(&worker, &peer, &json!({"method":"insertSlide","args":[1]}));
        assert_observable(&worker, &peer, "valid creation after refused creators");
    }

    #[test]
    fn batch_history_and_refusals_preserve_replay_state() {
        let (worker, peer) = pair();
        let slide = worker.slide_ids().unwrap()[0].clone();
        let request = |version: String, history: &str, text: &str| {
            json!({
                "method": "applyEdits", "args": [{
                    "expectVersion": version, "history": history,
                    "steps": [{"op":"setSlideNotes","target":{"slideId":slide},"text":text}],
                }],
            })
        };
        let op = request(worker.version().to_string(), "separate", "tracked");
        let result = apply_pair(&worker, &peer, &op);
        assert_eq!(result["outcome"]["result"]["ok"], true);
        assert_eq!(worker.undo.borrow().diagnostics(), (1, 0));
        let sequence = worker.replay_state.borrow().sequence;
        let refused = apply_pair(&worker, &peer, &op);
        assert_eq!(refused["outcome"]["result"]["ok"], false);
        assert_eq!(refused["consumed"], false);
        assert_eq!(worker.replay_state.borrow().sequence, sequence);
        apply_pair(
            &worker,
            &peer,
            &request(worker.version().to_string(), "none", "untracked"),
        );
        assert_eq!(worker.undo.borrow().diagnostics(), (1, 0));
        assert_observable(&worker, &peer, "untracked batch");
        apply_pair(&worker, &peer, &json!({"method":"undo","args":[]}));
        assert_observable(&worker, &peer, "undo after untracked batch");
        apply_pair(&worker, &peer, &json!({"method":"redo","args":[]}));
        assert_observable(&worker, &peer, "redo after untracked batch");
    }

    #[test]
    fn outcome_mismatch_keeps_recovery_content_and_fails_replay() {
        let (worker, _) = pair();
        let slide = worker.slide_ids().unwrap()[0].clone();
        let input = json!({"sequence":1,"baseVersion":worker.version(),
            "op":{"method":"setSlideNotes","args":[slide,"recover me"]},
            "expectedOutcome":{"result":null,"applied":false,"changedTargets":[],"canUndo":false,"canRedo":false}})
        .to_string();
        assert!(
            matches!(worker.replay_json(&input), Err(error) if error.code == "outcomeMismatch")
        );
        let bytes = worker.save().unwrap();
        assert_eq!(
            DeckSession::open(&bytes, 72)
                .unwrap()
                .snapshot()
                .unwrap()
                .slides[0]
                .notes,
            "recover me"
        );
        assert!(worker.replay_state.borrow().failed);
        assert!(worker.replay_json(&json!({"sequence":2,"baseVersion":worker.version(),"op":{"method":"undo","args":[]}}).to_string()).is_err());
    }

    #[test]
    fn replay_refuses_auto_capture_and_reentrant_dispatch() {
        let (worker, _) = pair();
        let input =
            json!({"sequence":1,"baseVersion":worker.version(),"op":{"method":"undo","args":[]}})
                .to_string();
        worker.set_undo_capture_mode(crate::UndoCaptureMode::Auto);
        assert!(matches!(worker.replay_json(&input), Err(error) if error.code == "capturePolicy"));
        assert_eq!(worker.replay_state.borrow().sequence, 0);
        worker.set_undo_capture_mode(crate::UndoCaptureMode::Manual);
        let worker = std::rc::Rc::new(worker);
        let observed = std::rc::Rc::clone(&worker);
        let refused = std::rc::Rc::new(std::cell::Cell::new(false));
        let refusal = std::rc::Rc::clone(&refused);
        let subscription = worker.observe_update_v1(move |_| {
            let input = json!({"sequence":1,"baseVersion":observed.version(),"op":{"method":"undo","args":[]}}).to_string();
            refusal.set(matches!(observed.replay_json(&input), Err(error) if error.code == "stage"));
        }).unwrap();
        let slide = worker.slide_ids().unwrap()[0].clone();
        replay(
            &worker,
            &json!({"method":"setSlideNotes","args":[slide,"outer"]}),
            None,
        );
        assert!(refused.get());
        assert_eq!(worker.replay_state.borrow().sequence, 1);
        assert_eq!(worker.snapshot().unwrap().slides[0].notes, "outer");
        drop(subscription);
    }

    #[test]
    fn malformed_envelopes_and_successful_noops_preserve_order() {
        let (worker, peer) = pair();
        let before = worker.replay_diagnostics();
        for op in [
            json!({"method":"undo","args":[1]}),
            json!({"method":"applyUpdate","args":[]}),
            json!({"method":"insertSlide","args":[1.5]}),
            json!({"method":"setCommentPosition","args":["missing",{"xEmu":9007199254740992_u64,"yEmu":0}]}),
            json!({"method":"propose","args":["agent",null,[{"type":"unknown"}]]}),
        ] {
            assert!(
                worker
                    .replay_json(
                        &json!({"sequence":1,"baseVersion":worker.version(),"op":op}).to_string()
                    )
                    .is_err(),
                "malformed {op}"
            );
            assert_eq!(
                worker.replay_diagnostics(),
                before,
                "malformed mutation {op}"
            );
        }
        for sequence in [0, 2, MAX_SAFE_INTEGER + 1] {
            assert!(worker.replay_json(&json!({"sequence":sequence,"baseVersion":worker.version(),"op":{"method":"undo","args":[]}}).to_string()).is_err());
        }
        for (index, method) in ["undo", "redo", "rejectProposal"].into_iter().enumerate() {
            let args = if method == "rejectProposal" {
                json!(["absent"])
            } else {
                json!([])
            };
            let result = apply_pair(&worker, &peer, &json!({"method":method,"args":args}));
            assert_eq!(result["sequence"], json!(index + 1));
            assert_eq!(result["consumed"], true);
            assert_eq!(result["outcome"]["applied"], false);
        }
        assert_eq!(worker.replay_diagnostics(), before);
    }
}
