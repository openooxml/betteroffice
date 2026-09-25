//! Version-checked, all-or-nothing host edit batches.
//!
//! A batch resolves every step against one captured state and returns policy failures as data.
//! The plan then runs on a private stage that shares this session's client id; its update is
//! rehearsed against an untouched copy of the base and adopted as one transaction.

use std::cmp::Reverse;
use std::collections::BTreeSet;
use std::fmt;
use std::sync::atomic::{AtomicU64, Ordering};

use ooxml_drawingml::{ShapeFill, ShapeOutline};
use serde::{Deserialize, Deserializer, Serialize};
use serde_json::Value;
use yrs::{ReadTxn, Transact, Update};

use crate::deck::{shape_fill, stroked_outline, validate_rect};
use crate::model::validate_xml_text;
use crate::proposals::inherited_style;
use crate::staging::Adoption;
use crate::story::{validate_alignment, validate_style_values};
use crate::target::{
    Deck, ParagraphView, Selection, ShapeTarget, StoryTarget, TextRange, TextTarget, keeps_fields,
};
use crate::{
    DeckSession, DeckSnapshot, EditCtx, EditError, EditResult, MAX_UPDATE_BYTES, ShapeKind,
    ShapeRect, ShapeSnapshot, ShapeStroke, StorySnapshot, TextStyle, TextStylePatch,
};

const MAX_STEPS: usize = 128;
const MAX_INSERTED_UNITS: usize = 1_048_576;
const MAX_STAGING_BYTES: usize = 256 * 1024 * 1024;
/// Largest request a read, search or batch accepts, measured as JSON.
pub const MAX_REQUEST_BYTES: usize = 16 * 1024 * 1024;

/// Charges serialized JSON sizes against a byte allowance.
pub(crate) struct ByteBudget {
    remaining: usize,
}

impl ByteBudget {
    pub fn new(limit: usize) -> Self {
        Self { remaining: limit }
    }

    /// Takes `value`'s JSON size from the allowance; false once it would run out.
    pub fn charge(&mut self, value: &impl Serialize) -> bool {
        struct Counter(usize);
        impl std::io::Write for Counter {
            fn write(&mut self, bytes: &[u8]) -> std::io::Result<usize> {
                self.0 = self.0.checked_sub(bytes.len()).ok_or_else(|| {
                    std::io::Error::new(std::io::ErrorKind::OutOfMemory, "over budget")
                })?;
                Ok(bytes.len())
            }

            fn flush(&mut self) -> std::io::Result<()> {
                Ok(())
            }
        }
        let mut counter = Counter(self.remaining);
        let fits = serde_json::to_writer(&mut counter, value).is_ok();
        if fits {
            self.remaining = counter.0;
        }
        fits
    }
}

/// An opaque, session-scoped optimistic-concurrency token. Compare for equality only.
#[derive(Clone, Debug, Deserialize, Eq, Hash, PartialEq, Serialize)]
#[serde(transparent)]
pub struct DocumentVersion(String);

impl DocumentVersion {
    pub fn as_str(&self) -> &str {
        &self.0
    }
}

impl From<String> for DocumentVersion {
    fn from(token: String) -> Self {
        Self(token)
    }
}

impl From<&str> for DocumentVersion {
    fn from(token: &str) -> Self {
        Self(token.to_owned())
    }
}

impl fmt::Display for DocumentVersion {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        f.write_str(&self.0)
    }
}

pub(crate) fn version_token(nonce: u64, epoch: u64) -> DocumentVersion {
    DocumentVersion(format!("{nonce:016x}-{epoch}"))
}

static NONCE_SEQUENCE: AtomicU64 = AtomicU64::new(0);

fn clock_entropy() -> u64 {
    #[cfg(not(all(target_family = "wasm", target_os = "unknown")))]
    {
        std::time::SystemTime::now()
            .duration_since(std::time::UNIX_EPOCH)
            .map(|elapsed| elapsed.as_nanos() as u64)
            .unwrap_or_default()
    }
    #[cfg(all(target_family = "wasm", target_os = "unknown"))]
    {
        0
    }
}

fn splitmix64(mut value: u64) -> u64 {
    value = value.wrapping_add(0x9e37_79b9_7f4a_7c15);
    value = (value ^ (value >> 30)).wrapping_mul(0xbf58_476d_1ce4_e5b9);
    value = (value ^ (value >> 27)).wrapping_mul(0x94d0_49bb_1331_11eb);
    value ^ (value >> 31)
}

pub(crate) fn mint_nonce(client_id: u64, entropy: u64) -> u64 {
    let sequence = NONCE_SEQUENCE.fetch_add(1, Ordering::Relaxed);
    splitmix64(
        splitmix64(sequence ^ client_id.rotate_left(32))
            ^ splitmix64(entropy)
            ^ splitmix64(clock_entropy()),
    )
}

/// Who asked for a batch. Provenance only: it selects neither permissions nor history.
#[derive(Clone, Copy, Debug, Default, Deserialize, Eq, PartialEq, Serialize)]
#[serde(rename_all = "camelCase")]
pub enum EditSource {
    #[default]
    Host,
    Agent,
}

/// How an applied batch enters local undo history.
#[derive(Clone, Copy, Debug, Default, Deserialize, Eq, PartialEq, Serialize)]
#[serde(rename_all = "camelCase")]
pub enum EditHistory {
    /// Exactly one undo step, whatever the capture mode.
    #[default]
    Separate,
    /// Outside undo history; existing undo and redo entries stay.
    None,
}

/// Which end of a target an insertion goes to.
#[derive(Clone, Copy, Debug, Deserialize, Eq, PartialEq, Serialize)]
#[serde(rename_all = "camelCase")]
pub enum TargetEdge {
    Start,
    End,
}

/// A required field that may be `null`.
fn nullable<'de, D: Deserializer<'de>, T: Deserialize<'de>>(
    deserializer: D,
) -> Result<Option<T>, D::Error> {
    Option::deserialize(deserializer)
}

/// Refuses a step unless its target currently reads exactly `text`.
#[derive(Clone, Debug, Deserialize, Eq, PartialEq, Serialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct TextGuard {
    pub text: String,
}

/// Refuses a step unless the shape's rectangle is exactly `rect`.
#[derive(Clone, Copy, Debug, Deserialize, Eq, PartialEq, Serialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct RectGuard {
    pub rect: ShapeRect,
}

/// Refuses a step unless the shape's authored fill is exactly `fill`; `None` is no fill of its
/// own.
#[derive(Clone, Debug, Deserialize, PartialEq, Serialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct FillGuard {
    #[serde(deserialize_with = "nullable")]
    pub fill: Option<Box<ShapeFill>>,
}

/// Refuses a step unless the shape's authored outline is exactly `outline`.
#[derive(Clone, Debug, Deserialize, PartialEq, Serialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct OutlineGuard {
    #[serde(deserialize_with = "nullable")]
    pub outline: Option<Box<ShapeOutline>>,
}

#[derive(Clone, Debug, Deserialize, Eq, PartialEq, Serialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct SlideTarget {
    pub slide_id: String,
}

/// One batch step. Text steps stay inside one paragraph, except `formatText` and
/// `setParagraphAlignment`, which may span paragraphs of one story.
#[derive(Clone, Debug, Deserialize, PartialEq, Serialize)]
#[serde(
    tag = "op",
    rename_all = "camelCase",
    rename_all_fields = "camelCase",
    deny_unknown_fields
)]
pub enum EditStep {
    InsertText {
        target: TextTarget,
        at: TargetEdge,
        text: String,
        #[serde(default, skip_serializing_if = "Option::is_none")]
        expect: Option<TextGuard>,
    },
    ReplaceText {
        target: TextTarget,
        text: String,
        #[serde(default, skip_serializing_if = "Option::is_none")]
        expect: Option<TextGuard>,
    },
    DeleteText {
        target: TextTarget,
        #[serde(default, skip_serializing_if = "Option::is_none")]
        expect: Option<TextGuard>,
    },
    FormatText {
        target: TextTarget,
        patch: TextStylePatch,
        #[serde(default, skip_serializing_if = "Option::is_none")]
        expect: Option<TextGuard>,
    },
    /// `None` restores the inherited alignment.
    SetParagraphAlignment {
        target: TextTarget,
        #[serde(deserialize_with = "nullable")]
        alignment: Option<String>,
        #[serde(default, skip_serializing_if = "Option::is_none")]
        expect: Option<TextGuard>,
    },
    /// Replaces the slide's speaker notes; empty text clears them.
    SetSlideNotes {
        target: SlideTarget,
        text: String,
        #[serde(default, skip_serializing_if = "Option::is_none")]
        expect: Option<TextGuard>,
    },
    SetShapeRect {
        target: ShapeTarget,
        rect: ShapeRect,
        #[serde(default, skip_serializing_if = "Option::is_none")]
        expect: Option<RectGuard>,
    },
    /// A solid `#RRGGBB` fill; `None` removes the fill.
    SetShapeFill {
        target: ShapeTarget,
        #[serde(deserialize_with = "nullable")]
        color: Option<String>,
        #[serde(default, skip_serializing_if = "Option::is_none")]
        expect: Option<FillGuard>,
    },
    SetShapeStroke {
        target: ShapeTarget,
        stroke: ShapeStroke,
        #[serde(default, skip_serializing_if = "Option::is_none")]
        expect: Option<OutlineGuard>,
    },
}

fn units(text: &str) -> usize {
    text.encode_utf16().count()
}

impl EditStep {
    /// UTF-16 units the step writes.
    fn inserted_units(&self) -> usize {
        match self {
            Self::InsertText { text, .. }
            | Self::ReplaceText { text, .. }
            | Self::SetSlideNotes { text, .. } => units(text),
            _ => 0,
        }
    }
}

#[derive(Clone, Debug, Deserialize, PartialEq, Serialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct EditRequest {
    /// The version the targets were read at.
    pub expect_version: DocumentVersion,
    #[serde(default)]
    pub source: EditSource,
    #[serde(default)]
    pub history: EditHistory,
    pub steps: Vec<EditStep>,
}

#[derive(Clone, Copy, Debug, Deserialize, Eq, PartialEq, Serialize)]
#[serde(rename_all = "kebab-case")]
pub enum EditFailureCode {
    StaleVersion,
    MissingTarget,
    AmbiguousTarget,
    ContentMismatch,
    OverlappingSteps,
    Unsupported,
    InvalidStep,
    LimitExceeded,
}

/// What a failure, preview or receipt refers to.
#[derive(Clone, Debug, Deserialize, Eq, PartialEq, Serialize)]
#[serde(
    tag = "kind",
    rename_all = "camelCase",
    rename_all_fields = "camelCase"
)]
pub enum EditTarget {
    Range(TextRange),
    Search { within: StoryTarget, text: String },
    Story(StoryTarget),
    Shape(ShapeTarget),
    Slide(SlideTarget),
}

impl From<TextTarget> for EditTarget {
    fn from(target: TextTarget) -> Self {
        match target {
            TextTarget::Range(range) => Self::Range(range),
            TextTarget::Search { within, text } => Self::Search { within, text },
        }
    }
}

#[derive(Clone, Debug, Deserialize, Eq, PartialEq, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct EditFailure {
    pub code: EditFailureCode,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub step_index: Option<u32>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub conflicting_step_index: Option<u32>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub target: Option<Box<EditTarget>>,
    pub message: String,
}

impl EditFailure {
    fn at(mut self, step_index: u32) -> Self {
        self.step_index = Some(step_index);
        self
    }
}

/// A policy refusal; the deck is untouched at `version`.
#[derive(Clone, Debug, Deserialize, Eq, PartialEq, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct EditRefusal {
    pub version: DocumentVersion,
    pub failure: EditFailure,
}

/// Characters a failure message keeps.
const MAX_MESSAGE_CHARS: usize = 512;
/// JSON bytes of a target a failure still echoes.
const MAX_ECHOED_TARGET_BYTES: usize = 64 * 1024;

/// A failure small enough to return whatever text caused it: long messages are cut, and a
/// target too large to echo is left out.
pub(crate) fn failure(
    code: EditFailureCode,
    message: impl Into<String>,
    target: Option<&EditTarget>,
) -> EditFailure {
    let mut message = message.into();
    if let Some((cut, _)) = message.char_indices().nth(MAX_MESSAGE_CHARS) {
        message.truncate(cut);
        message.push('…');
    }
    EditFailure {
        code,
        step_index: None,
        conflicting_step_index: None,
        target: target
            .filter(|target| ByteBudget::new(MAX_ECHOED_TARGET_BYTES).charge(target))
            .cloned()
            .map(Box::new),
        message,
    }
}

/// `text` quoted for a diagnostic, cut to its first 64 characters so a refusal stays small.
pub(crate) fn quoted(text: &str) -> String {
    match text.char_indices().nth(64) {
        Some((cut, _)) => format!("{:?}…", &text[..cut]),
        None => format!("{text:?}"),
    }
}

pub(crate) fn refusal(version: DocumentVersion, failure: EditFailure) -> EditRefusal {
    EditRefusal { version, failure }
}

/// The refusal for a JSON request longer than [`MAX_REQUEST_BYTES`], before decoding it.
pub fn oversized_request(version: DocumentVersion) -> EditRefusal {
    refusal(version, request_limit())
}

pub(crate) fn request_limit() -> EditFailure {
    failure(
        EditFailureCode::LimitExceeded,
        format!("a request holds at most {MAX_REQUEST_BYTES} bytes of JSON"),
        None,
    )
}

/// What one step did. Text targets describe the final state; a deletion reports its collapsed
/// boundary.
#[derive(Clone, Debug, Deserialize, Eq, PartialEq, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct EditReceipt {
    pub step_index: u32,
    pub changed: bool,
    pub target: EditTarget,
}

/// What one step would do at its resolved pre-batch target. It reserves nothing.
#[derive(Clone, Debug, Deserialize, Eq, PartialEq, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct EditPreview {
    pub step_index: u32,
    pub target: EditTarget,
    pub would_change: bool,
}

#[derive(Clone, Debug, Deserialize, Eq, PartialEq, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct EditValidation {
    pub base_version: DocumentVersion,
    pub would_apply: bool,
    pub previews: Vec<EditPreview>,
}

#[derive(Clone, Debug, Deserialize, Eq, PartialEq, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct EditApplication {
    pub base_version: DocumentVersion,
    pub version: DocumentVersion,
    /// False when every step was a no-op: no history, update or version change.
    pub applied: bool,
    pub source: EditSource,
    pub changed_slides: Vec<String>,
    pub changed_stories: Vec<String>,
    /// One per request step, in request order.
    pub receipts: Vec<EditReceipt>,
}

pub type ValidationOutcome = Result<EditValidation, EditRefusal>;
pub type EditOutcome = Result<EditApplication, EditRefusal>;

/// JSON of a policy outcome: the success body or the refusal, tagged with `ok`.
pub fn outcome_json<T: Serialize>(
    outcome: &Result<T, EditRefusal>,
) -> Result<String, serde_json::Error> {
    let mut value = match outcome {
        Ok(body) => serde_json::to_value(body)?,
        Err(refusal) => serde_json::to_value(refusal)?,
    };
    if let Value::Object(object) = &mut value {
        object.insert("ok".to_owned(), Value::Bool(outcome.is_ok()));
    }
    serde_json::to_string(&value)
}

/// A text range or boundary a step writes, in the captured state.
#[derive(Clone, Copy)]
enum Claim {
    Span(u32, u32),
    Point(u32),
}

impl Claim {
    fn conflicts(self, other: Self) -> bool {
        match (self, other) {
            (Self::Span(a, b), Self::Span(c, d)) => a < d && c < b,
            (Self::Point(p), Self::Point(q)) => p == q,
            (Self::Point(p), Self::Span(a, b)) | (Self::Span(a, b), Self::Point(p)) => {
                a <= p && p <= b
            }
        }
    }
}

/// A slide or shape property a step replaces whole.
#[derive(Clone, Copy, Eq, PartialEq)]
enum Property {
    Notes,
    Rect,
    Fill,
    Stroke,
}

enum Effect {
    Insert {
        at: u32,
        text: String,
        style: TextStyle,
    },
    Replace {
        start: u32,
        end: u32,
        text: String,
        style: TextStyle,
    },
    Delete {
        start: u32,
        end: u32,
    },
    Format {
        start: u32,
        end: u32,
        patch: TextStylePatch,
    },
    Align {
        start: u32,
        end: u32,
        alignment: Option<String>,
    },
    Notes(String),
    Rect(ShapeRect),
    Fill(Option<String>),
    Stroke(ShapeStroke),
}

impl Effect {
    /// Where the effect starts and how much it grows its story, for effects that move text.
    fn shift(&self) -> Option<(u32, i64)> {
        match self {
            Self::Insert { at, text, .. } => Some((*at, units(text) as i64)),
            Self::Replace {
                start, end, text, ..
            } => Some((*start, units(text) as i64 - i64::from(end - start))),
            Self::Delete { start, end } => Some((*start, -i64::from(end - start))),
            _ => None,
        }
    }
}

/// Where a receipt finds a step's result in the final state.
enum Location {
    /// `len` units from `start`, which moves with earlier edits.
    Grown { start: u32, len: u32 },
    /// The resolved range, both ends moving with earlier edits.
    Kept,
    /// The step's slide or shape.
    Target,
}

struct Planned {
    index: u32,
    /// The resolved pre-batch target.
    target: EditTarget,
    slide_id: String,
    shape_id: Option<String>,
    story_id: Option<String>,
    claims: Vec<Claim>,
    paragraphs: Vec<String>,
    /// Conflicts with every other write to its paragraphs.
    exclusive: bool,
    property: Option<Property>,
    effect: Option<Effect>,
    location: Location,
}

impl Planned {
    fn text(
        index: u32,
        selection: &Selection<'_>,
        claims: Vec<Claim>,
        paragraphs: Vec<String>,
        exclusive: bool,
        effect: Option<Effect>,
        location: Location,
    ) -> Self {
        Self {
            index,
            target: EditTarget::Range(selection.range()),
            slide_id: selection.view.slide.id.clone(),
            shape_id: Some(selection.view.shape.id.clone()),
            story_id: Some(selection.view.story.id.clone()),
            claims,
            paragraphs,
            exclusive,
            property: None,
            effect,
            location,
        }
    }

    fn property(
        index: u32,
        target: EditTarget,
        slide_id: &str,
        shape_id: Option<&str>,
        property: Property,
        effect: Option<Effect>,
    ) -> Self {
        Self {
            index,
            target,
            slide_id: slide_id.to_owned(),
            shape_id: shape_id.map(str::to_owned),
            story_id: None,
            claims: Vec::new(),
            paragraphs: Vec::new(),
            exclusive: false,
            property: Some(property),
            effect,
            location: Location::Target,
        }
    }

    fn owned_property(&self) -> Option<(&str, Property)> {
        let property = self.property?;
        let owner = match property {
            Property::Notes => self.slide_id.as_str(),
            _ => self.shape_id.as_deref()?,
        };
        Some((owner, property))
    }

    fn shift(&self) -> Option<(u32, i64)> {
        self.effect.as_ref()?.shift()
    }

    fn conflicts(&self, other: &Self) -> bool {
        if self.property.is_some() && self.owned_property() == other.owned_property() {
            return true;
        }
        if self.story_id.is_none() || self.story_id != other.story_id {
            return false;
        }
        let shared = |owner: &Self, other: &Self| {
            owner.exclusive
                && owner
                    .paragraphs
                    .iter()
                    .any(|id| other.paragraphs.contains(id))
        };
        shared(self, other)
            || shared(other, self)
            || self
                .claims
                .iter()
                .any(|mine| other.claims.iter().any(|theirs| mine.conflicts(*theirs)))
    }
}

struct Plan {
    base_version: DocumentVersion,
    nonce: u64,
    epoch: u64,
    id_counter: u64,
    source: EditSource,
    before: DeckSnapshot,
    steps: Vec<Planned>,
}

impl Plan {
    fn would_apply(&self) -> bool {
        self.steps.iter().any(|planned| planned.effect.is_some())
    }

    fn application(self, applied: bool, version: DocumentVersion) -> EditApplication {
        let changed = |id: fn(&Planned) -> Option<&String>| -> Vec<String> {
            self.steps
                .iter()
                .filter(|planned| applied && planned.effect.is_some())
                .filter_map(id)
                .cloned()
                .collect::<BTreeSet<_>>()
                .into_iter()
                .collect()
        };
        EditApplication {
            changed_slides: changed(|planned| Some(&planned.slide_id)),
            changed_stories: changed(|planned| planned.story_id.as_ref()),
            receipts: receipts(&self.steps, applied),
            base_version: self.base_version,
            version,
            applied,
            source: self.source,
        }
    }
}

/// A plan executed on a private stage, with the update that adopts it.
struct Staged {
    stage: DeckSession,
    update: Update,
}

fn unsupported(message: impl Into<String>, target: &EditTarget) -> EditFailure {
    failure(EditFailureCode::Unsupported, message, Some(target))
}

fn invalid(error: impl fmt::Display, target: &EditTarget) -> EditFailure {
    failure(
        EditFailureCode::InvalidStep,
        error.to_string(),
        Some(target),
    )
}

fn mismatch(message: impl Into<String>, target: &EditTarget) -> EditFailure {
    failure(EditFailureCode::ContentMismatch, message, Some(target))
}

fn pending_failure() -> EditFailure {
    failure(
        EditFailureCode::Unsupported,
        "the deck holds updates that are not integrated yet",
        None,
    )
}

fn check_budget(request: &EditRequest) -> Result<(), EditFailure> {
    let limit = |message: String| failure(EditFailureCode::LimitExceeded, message, None);
    if request.steps.len() > MAX_STEPS {
        return Err(limit(format!("a batch holds at most {MAX_STEPS} steps")));
    }
    if !ByteBudget::new(MAX_REQUEST_BYTES).charge(request) {
        return Err(request_limit());
    }
    let inserted: usize = request.steps.iter().map(EditStep::inserted_units).sum();
    if inserted > MAX_INSERTED_UNITS {
        return Err(limit(format!(
            "a batch inserts at most {MAX_INSERTED_UNITS} UTF-16 units"
        )));
    }
    Ok(())
}

fn check_text(
    expect: Option<&TextGuard>,
    actual: &str,
    target: &EditTarget,
) -> Result<(), EditFailure> {
    match expect {
        Some(expected) if expected.text != actual => Err(mismatch(
            format!(
                "the target reads {}, not the expected {}",
                quoted(actual),
                quoted(&expected.text)
            ),
            target,
        )),
        _ => Ok(()),
    }
}

fn check_inserted(text: &str, target: &EditTarget) -> Result<(), EditFailure> {
    if text.contains(['\n', '\r', '\u{b}', '\u{2028}', '\u{2029}']) {
        return Err(invalid(
            "inserted text may not contain paragraph or line breaks",
            target,
        ));
    }
    validate_xml_text(text).map_err(|error| invalid(error, target))
}

fn check_editable(paragraph: &ParagraphView, target: &EditTarget) -> Result<(), EditFailure> {
    if paragraph.editable {
        Ok(())
    } else {
        Err(unsupported(
            format!(
                "saving could not keep the fields of paragraph {} once it changes",
                quoted(&paragraph.id)
            ),
            target,
        ))
    }
}

fn check_point(paragraph: &ParagraphView, at: u32, target: &EditTarget) -> Result<(), EditFailure> {
    if paragraph
        .fields
        .iter()
        .any(|field| field.start < at && at < field.end)
    {
        return Err(unsupported(
            "text cannot be inserted inside a field",
            target,
        ));
    }
    Ok(())
}

fn check_span(
    paragraph: &ParagraphView,
    start: u32,
    end: u32,
    target: &EditTarget,
) -> Result<(), EditFailure> {
    if paragraph
        .line_breaks
        .iter()
        .any(|at| start <= *at && *at < end)
    {
        return Err(unsupported(
            "replacing or deleting line breaks is not supported",
            target,
        ));
    }
    if paragraph
        .fields
        .iter()
        .any(|field| start < field.end && field.start < end)
    {
        return Err(unsupported(
            "text steps cannot change a field's text",
            target,
        ));
    }
    Ok(())
}

/// The formatting new text at `at` takes, when the deck can store it.
fn writable_style(
    story: &StorySnapshot,
    at: u32,
    target: &EditTarget,
) -> Result<TextStyle, EditFailure> {
    let style = inherited_style(story, at);
    validate_style_values(
        style.font_family.as_deref(),
        style.underline.as_deref(),
        style.color.as_deref(),
        style.font_size_pt,
        style.spacing_pt,
        style.baseline_pct,
    )
    .map_err(|error| {
        unsupported(
            format!("the surrounding formatting cannot be written: {error}"),
            target,
        )
    })?;
    Ok(style)
}

/// Whether applying `patch` changes any text in `start..end`.
fn patch_changes(story: &StorySnapshot, start: u32, end: u32, patch: &TextStylePatch) -> bool {
    fn differs<T: PartialEq>(current: &Option<T>, patched: &Option<T>) -> bool {
        patched
            .as_ref()
            .is_some_and(|value| current.as_ref() != Some(value))
    }
    let mut offset = 0;
    for paragraph in &story.paragraphs {
        for run in &paragraph.runs {
            let (from, to) = (offset, offset + units(&run.text) as u32);
            offset = to;
            let style = &run.style;
            if from < end
                && start < to
                && (differs(&style.bold, &patch.bold)
                    || differs(&style.italic, &patch.italic)
                    || differs(&style.font_size_pt, &patch.font_size_pt)
                    || differs(&style.color, &patch.color)
                    || differs(&style.font_family, &patch.font_family)
                    || differs(&style.underline, &patch.underline)
                    || differs(&style.spacing_pt, &patch.spacing_pt)
                    || differs(&style.baseline_pct, &patch.baseline_pct))
            {
                return true;
            }
        }
        offset += 1;
    }
    false
}

enum TextWrite<'s> {
    Insert(TargetEdge, &'s str),
    Replace(&'s str),
    Delete,
}

/// Insertion, replacement and deletion: plain text inside one paragraph.
fn plan_text_write(
    deck: &Deck<'_>,
    index: u32,
    target: &TextTarget,
    expect: Option<&TextGuard>,
    write: TextWrite<'_>,
) -> Result<Planned, EditFailure> {
    let requested = EditTarget::from(target.clone());
    let selection = deck.text(target)?;
    let selected = selection.text();
    check_text(expect, &selected, &requested)?;
    let Some(paragraph) = selection.paragraph() else {
        return Err(unsupported(
            "a text range must stay within one paragraph",
            &requested,
        ));
    };
    check_editable(paragraph, &requested)?;
    let story = selection.view.story;
    let (start, end) = (selection.start, selection.end);
    let (claim, effect, location) = match write {
        TextWrite::Insert(edge, text) => {
            check_inserted(text, &requested)?;
            let at = if edge == TargetEdge::Start {
                start
            } else {
                end
            };
            check_point(paragraph, at, &requested)?;
            let effect = if text.is_empty() {
                None
            } else {
                Some(Effect::Insert {
                    at,
                    text: text.to_owned(),
                    style: writable_style(story, at, &requested)?,
                })
            };
            let len = if effect.is_some() { units(text) } else { 0 };
            (
                Claim::Point(at),
                effect,
                Location::Grown {
                    start: at,
                    len: len as u32,
                },
            )
        }
        TextWrite::Replace(text) => {
            check_inserted(text, &requested)?;
            let claim = if start == end {
                check_point(paragraph, start, &requested)?;
                Claim::Point(start)
            } else {
                check_span(paragraph, start, end, &requested)?;
                Claim::Span(start, end)
            };
            let effect = if selected == text {
                None
            } else {
                Some(Effect::Replace {
                    start,
                    end,
                    text: text.to_owned(),
                    style: writable_style(story, start, &requested)?,
                })
            };
            let len = if effect.is_some() {
                units(text) as u32
            } else {
                end - start
            };
            (claim, effect, Location::Grown { start, len })
        }
        TextWrite::Delete if start == end => {
            (Claim::Point(start), None, Location::Grown { start, len: 0 })
        }
        TextWrite::Delete => {
            check_span(paragraph, start, end, &requested)?;
            (
                Claim::Span(start, end),
                Some(Effect::Delete { start, end }),
                Location::Grown { start, len: 0 },
            )
        }
    };
    let paragraphs = vec![paragraph.id.clone()];
    Ok(Planned::text(
        index,
        &selection,
        vec![claim],
        paragraphs,
        false,
        effect,
        location,
    ))
}

fn plan_format(
    deck: &Deck<'_>,
    index: u32,
    target: &TextTarget,
    patch: &TextStylePatch,
    expect: Option<&TextGuard>,
) -> Result<Planned, EditFailure> {
    let requested = EditTarget::from(target.clone());
    let selection = deck.text(target)?;
    check_text(expect, &selection.text(), &requested)?;
    validate_style_values(
        patch.font_family.as_deref(),
        patch.underline.as_deref(),
        patch.color.as_deref(),
        patch.font_size_pt,
        patch.spacing_pt,
        patch.baseline_pct,
    )
    .map_err(|error| invalid(error, &requested))?;
    let (start, end) = (selection.start, selection.end);
    let selected = selection.selected();
    for paragraph in &selected {
        check_editable(paragraph, &requested)?;
        if paragraph.fields.iter().any(|field| {
            (field.start < start && start < field.end) || (field.start < end && end < field.end)
        }) {
            return Err(unsupported(
                "formatting part of a field would unbind it",
                &requested,
            ));
        }
    }
    let effect =
        (start < end && patch_changes(selection.view.story, start, end, patch)).then(|| {
            Effect::Format {
                start,
                end,
                patch: patch.clone(),
            }
        });
    let claim = if start == end {
        Claim::Point(start)
    } else {
        Claim::Span(start, end)
    };
    let paragraphs = selected
        .iter()
        .map(|paragraph| paragraph.id.clone())
        .collect();
    Ok(Planned::text(
        index,
        &selection,
        vec![claim],
        paragraphs,
        false,
        effect,
        Location::Kept,
    ))
}

fn plan_alignment(
    deck: &Deck<'_>,
    index: u32,
    target: &TextTarget,
    alignment: &Option<String>,
    expect: Option<&TextGuard>,
) -> Result<Planned, EditFailure> {
    let requested = EditTarget::from(target.clone());
    let selection = deck.text(target)?;
    check_text(expect, &selection.text(), &requested)?;
    validate_alignment(alignment.as_deref()).map_err(|error| invalid(error, &requested))?;
    let selected = selection.selected();
    let effect = selected
        .iter()
        .any(|paragraph| paragraph.alignment != *alignment)
        .then(|| Effect::Align {
            start: selection.start,
            end: selection.end,
            alignment: alignment.clone(),
        });
    let paragraphs = selected
        .iter()
        .map(|paragraph| paragraph.id.clone())
        .collect();
    Ok(Planned::text(
        index,
        &selection,
        Vec::new(),
        paragraphs,
        true,
        effect,
        Location::Kept,
    ))
}

fn plan_notes(
    deck: &Deck<'_>,
    index: u32,
    target: &SlideTarget,
    text: &str,
    expect: Option<&TextGuard>,
) -> Result<Planned, EditFailure> {
    let requested = EditTarget::Slide(target.clone());
    let slide = deck.slide(&target.slide_id, &requested)?;
    check_text(expect, &slide.notes, &requested)?;
    validate_xml_text(text).map_err(|error| invalid(error, &requested))?;
    let effect = (slide.notes != text).then(|| Effect::Notes(text.to_owned()));
    Ok(Planned::property(
        index,
        requested,
        &slide.id,
        None,
        Property::Notes,
        effect,
    ))
}

/// A shape that geometry, fill and outline steps may address.
fn top_level_shape<'a>(
    deck: &Deck<'a>,
    target: &ShapeTarget,
    requested: &EditTarget,
) -> Result<&'a ShapeSnapshot, EditFailure> {
    let (_, shape, top_level) = deck.shape(target, requested)?;
    if !top_level {
        return Err(unsupported(
            "geometry, fill and outline steps address shapes at the top of a slide",
            requested,
        ));
    }
    Ok(shape)
}

fn require_preset(shape: &ShapeSnapshot, target: &EditTarget) -> Result<(), EditFailure> {
    if shape.kind == ShapeKind::Shape {
        Ok(())
    } else {
        Err(unsupported(
            "only preset shapes support shape styling",
            target,
        ))
    }
}

fn plan_rect(
    deck: &Deck<'_>,
    index: u32,
    target: &ShapeTarget,
    rect: ShapeRect,
    expect: Option<&RectGuard>,
) -> Result<Planned, EditFailure> {
    let requested = EditTarget::Shape(target.clone());
    let shape = top_level_shape(deck, target, &requested)?;
    let current = ShapeRect {
        x: shape.x,
        y: shape.y,
        width: shape.width,
        height: shape.height,
    };
    if let Some(guard) = expect
        && guard.rect != current
    {
        return Err(mismatch(
            format!(
                "the shape's rectangle is {current:?}, not the expected {:?}",
                guard.rect
            ),
            &requested,
        ));
    }
    validate_rect(rect).map_err(|error| invalid(error, &requested))?;
    let effect = (current != rect).then_some(Effect::Rect(rect));
    Ok(Planned::property(
        index,
        requested,
        &target.slide_id,
        Some(&shape.id),
        Property::Rect,
        effect,
    ))
}

fn plan_fill(
    deck: &Deck<'_>,
    index: u32,
    target: &ShapeTarget,
    color: &Option<String>,
    expect: Option<&FillGuard>,
) -> Result<Planned, EditFailure> {
    let requested = EditTarget::Shape(target.clone());
    let shape = top_level_shape(deck, target, &requested)?;
    if let Some(guard) = expect
        && guard.fill.as_deref() != shape.fill.as_ref()
    {
        return Err(mismatch(
            "the shape's fill differs from the expected fill",
            &requested,
        ));
    }
    require_preset(shape, &requested)?;
    let fill = shape_fill(color.as_deref()).map_err(|error| invalid(error, &requested))?;
    let effect = (shape.fill.as_ref() != Some(&fill)).then(|| Effect::Fill(color.clone()));
    Ok(Planned::property(
        index,
        requested,
        &target.slide_id,
        Some(&shape.id),
        Property::Fill,
        effect,
    ))
}

fn plan_stroke(
    deck: &Deck<'_>,
    index: u32,
    target: &ShapeTarget,
    stroke: &ShapeStroke,
    expect: Option<&OutlineGuard>,
) -> Result<Planned, EditFailure> {
    let requested = EditTarget::Shape(target.clone());
    let shape = top_level_shape(deck, target, &requested)?;
    if let Some(guard) = expect
        && guard.outline.as_deref() != shape.outline.as_ref()
    {
        return Err(mismatch(
            "the shape's outline differs from the expected outline",
            &requested,
        ));
    }
    require_preset(shape, &requested)?;
    let outline = stroked_outline(shape.outline.clone(), stroke)
        .map_err(|error| invalid(error, &requested))?;
    let effect = (shape.outline.as_ref() != Some(&outline)).then(|| Effect::Stroke(stroke.clone()));
    Ok(Planned::property(
        index,
        requested,
        &target.slide_id,
        Some(&shape.id),
        Property::Stroke,
        effect,
    ))
}

fn plan_step(deck: &Deck<'_>, index: u32, step: &EditStep) -> Result<Planned, EditFailure> {
    match step {
        EditStep::InsertText {
            target,
            at,
            text,
            expect,
        } => plan_text_write(
            deck,
            index,
            target,
            expect.as_ref(),
            TextWrite::Insert(*at, text),
        ),
        EditStep::ReplaceText {
            target,
            text,
            expect,
        } => plan_text_write(
            deck,
            index,
            target,
            expect.as_ref(),
            TextWrite::Replace(text),
        ),
        EditStep::DeleteText { target, expect } => {
            plan_text_write(deck, index, target, expect.as_ref(), TextWrite::Delete)
        }
        EditStep::FormatText {
            target,
            patch,
            expect,
        } => plan_format(deck, index, target, patch, expect.as_ref()),
        EditStep::SetParagraphAlignment {
            target,
            alignment,
            expect,
        } => plan_alignment(deck, index, target, alignment, expect.as_ref()),
        EditStep::SetSlideNotes {
            target,
            text,
            expect,
        } => plan_notes(deck, index, target, text, expect.as_ref()),
        EditStep::SetShapeRect {
            target,
            rect,
            expect,
        } => plan_rect(deck, index, target, *rect, expect.as_ref()),
        EditStep::SetShapeFill {
            target,
            color,
            expect,
        } => plan_fill(deck, index, target, color, expect.as_ref()),
        EditStep::SetShapeStroke {
            target,
            stroke,
            expect,
        } => plan_stroke(deck, index, target, stroke, expect.as_ref()),
    }
}

/// Runs the plan on `stage`: effects that keep offsets first, then text changes from the end
/// of each story backwards, so every captured offset stays valid when its step runs.
fn execute(stage: &DeckSession, steps: &[Planned]) -> EditResult<()> {
    let context = EditCtx::local("batch");
    let run = |planned: &Planned| -> EditResult<()> {
        let slide = planned.slide_id.as_str();
        let shape = planned.shape_id.as_deref().unwrap_or_default();
        let story = planned.story_id.as_deref().unwrap_or_default();
        match &planned.effect {
            None => {}
            Some(Effect::Insert { at, text, style }) => {
                stage.insert_text(&context, story, *at, text, style)?;
            }
            Some(Effect::Replace {
                start,
                end,
                text,
                style,
            }) => {
                stage.delete_text(&context, story, *start, *end)?;
                if !text.is_empty() {
                    stage.insert_text(&context, story, *start, text, style)?;
                }
            }
            Some(Effect::Delete { start, end }) => {
                stage.delete_text(&context, story, *start, *end)?;
            }
            Some(Effect::Format { start, end, patch }) => {
                stage.format_text(&context, story, *start, *end, patch)?;
            }
            Some(Effect::Align {
                start,
                end,
                alignment,
            }) => {
                stage.set_paragraph_alignment(
                    &context,
                    story,
                    *start,
                    *end,
                    alignment.as_deref(),
                )?;
            }
            Some(Effect::Notes(text)) => stage.set_slide_notes(&context, slide, text)?,
            Some(Effect::Rect(rect)) => {
                stage.set_shape_rect(&context, slide, shape, *rect)?;
            }
            Some(Effect::Fill(color)) => {
                stage.set_shape_fill(&context, slide, shape, color.as_deref())?;
            }
            Some(Effect::Stroke(stroke)) => {
                stage.set_shape_stroke(&context, slide, shape, stroke)?;
            }
        }
        Ok(())
    };
    let (mut moving, fixed): (Vec<&Planned>, Vec<&Planned>) = steps
        .iter()
        .filter(|planned| planned.effect.is_some())
        .partition(|planned| planned.shift().is_some());
    moving.sort_by_key(|planned| Reverse(planned.shift().map(|(start, _)| start)));
    for planned in fixed.into_iter().chain(moving) {
        run(planned).map_err(|error| {
            EditError::InvalidState(format!("staged step {} failed: {error}", planned.index))
        })?;
    }
    Ok(())
}

/// Receipts in request order; unapplied receipts keep their pre-batch targets.
fn receipts(steps: &[Planned], applied: bool) -> Vec<EditReceipt> {
    let shifted = |story: Option<&str>, raw: u32| -> u32 {
        if !applied {
            return raw;
        }
        let delta: i64 = steps
            .iter()
            .filter(|planned| planned.story_id.as_deref() == story)
            .filter_map(Planned::shift)
            .filter(|(start, _)| *start < raw)
            .map(|(_, delta)| delta)
            .sum();
        (i64::from(raw) + delta).max(0) as u32
    };
    steps
        .iter()
        .map(|planned| {
            let story = planned.story_id.as_deref();
            let target = match (&planned.location, &planned.target) {
                (Location::Grown { start, len }, EditTarget::Range(range)) => {
                    let start = shifted(story, *start);
                    EditTarget::Range(TextRange {
                        start,
                        end: start + len,
                        ..range.clone()
                    })
                }
                (Location::Kept, EditTarget::Range(range)) => EditTarget::Range(TextRange {
                    start: shifted(story, range.start),
                    end: shifted(story, range.end),
                    ..range.clone()
                }),
                _ => planned.target.clone(),
            };
            EditReceipt {
                step_index: planned.index,
                changed: applied && planned.effect.is_some(),
                target,
            }
        })
        .collect()
}

impl DeckSession {
    fn plan(&self, request: &EditRequest) -> EditResult<Result<Plan, EditRefusal>> {
        let nonce = self.version_nonce.load(Ordering::Relaxed);
        let epoch = self.epoch();
        let version = version_token(nonce, epoch);
        let refuse = |failure: EditFailure| -> EditResult<Result<Plan, EditRefusal>> {
            Ok(Err(refusal(version.clone(), failure)))
        };
        if let Err(failure) = check_budget(request) {
            return refuse(failure);
        }
        if request.expect_version != version {
            return refuse(failure(
                EditFailureCode::StaleVersion,
                "the deck changed since the expected version was read",
                None,
            ));
        }
        if self.has_pending_updates() {
            return refuse(pending_failure());
        }
        let before = self.snapshot()?;
        let mut steps: Vec<Planned> = Vec::with_capacity(request.steps.len());
        {
            let deck = Deck::new(&before, &self.package);
            for (index, step) in request.steps.iter().enumerate() {
                let index = index as u32;
                let planned = match plan_step(&deck, index, step) {
                    Ok(planned) => planned,
                    Err(failure) => return refuse(failure.at(index)),
                };
                if let Some(earlier) = steps.iter().find(|earlier| earlier.conflicts(&planned)) {
                    let mut conflict = failure(
                        EditFailureCode::OverlappingSteps,
                        format!("step {index} conflicts with step {}", earlier.index),
                        Some(&planned.target),
                    )
                    .at(index);
                    conflict.conflicting_step_index = Some(earlier.index);
                    return refuse(conflict);
                }
                steps.push(planned);
            }
        }
        Ok(Ok(Plan {
            base_version: version,
            nonce,
            epoch,
            id_counter: self.id_counter.load(Ordering::Relaxed),
            source: request.source,
            before,
            steps,
        }))
    }

    /// Executes a plan on a private stage and rehearses its adoption. `None` when the staged
    /// deck equals the captured one.
    fn stage_plan(&self, plan: &Plan) -> EditResult<Result<Option<Staged>, EditRefusal>> {
        let refuse = |code: EditFailureCode, message: String| {
            Ok(Err(refusal(
                plan.base_version.clone(),
                failure(code, message, None),
            )))
        };
        let base = self.state_update_v1();
        if base.len() > MAX_STAGING_BYTES {
            return refuse(
                EditFailureCode::LimitExceeded,
                format!("batches stage decks of at most {MAX_STAGING_BYTES} bytes"),
            );
        }
        let base_state = self.doc.transact().state_vector();
        let stage = self.stage()?;
        execute(&stage, &plan.steps)?;
        let staged = match stage.validated_snapshot() {
            Ok(staged) => staged,
            Err(error) => {
                return refuse(
                    EditFailureCode::Unsupported,
                    format!("the batch would leave the deck invalid: {error}"),
                );
            }
        };
        if staged == plan.before {
            return Ok(Ok(None));
        }
        for planned in plan.steps.iter().filter(|planned| planned.effect.is_some()) {
            let Some(story) = &planned.story_id else {
                continue;
            };
            if let Some(paragraph) = planned
                .paragraphs
                .iter()
                .find(|paragraph| !keeps_fields(&self.package, &staged, story, paragraph))
            {
                let failure = unsupported(
                    format!(
                        "saving would turn a field of paragraph {} into plain text",
                        quoted(paragraph)
                    ),
                    &planned.target,
                )
                .at(planned.index);
                return Ok(Err(refusal(plan.base_version.clone(), failure)));
            }
        }
        let (bytes, update) = stage.staged_update(&base_state)?;
        if bytes.len() > MAX_UPDATE_BYTES {
            return refuse(
                EditFailureCode::LimitExceeded,
                format!("a batch's update may not exceed {MAX_UPDATE_BYTES} bytes"),
            );
        }
        self.rehearse(&base, &bytes, &stage, &staged)?;
        Ok(Ok(Some(Staged { stage, update })))
    }

    fn check_commit(&self, plan: &Plan) -> Result<(), EditRefusal> {
        if self.version_nonce.load(Ordering::Relaxed) != plan.nonce
            || self.epoch() != plan.epoch
            || self.id_counter.load(Ordering::Relaxed) != plan.id_counter
        {
            return Err(refusal(
                self.version(),
                failure(
                    EditFailureCode::StaleVersion,
                    "the deck changed while the batch was staged",
                    None,
                ),
            ));
        }
        if self.has_pending_updates() {
            return Err(refusal(self.version(), pending_failure()));
        }
        Ok(())
    }

    /// Resolves, stages and rehearses a batch like [`DeckSession::apply_edits`], then discards
    /// it: nothing changes and no ids are reserved. `Err` means an internal failure.
    pub fn validate_edits(&self, request: &EditRequest) -> EditResult<ValidationOutcome> {
        let plan = match self.plan(request)? {
            Ok(plan) => plan,
            Err(refusal) => return Ok(Err(refusal)),
        };
        let would_apply = plan.would_apply()
            && match self.stage_plan(&plan)? {
                Ok(staged) => staged.is_some(),
                Err(refusal) => return Ok(Err(refusal)),
            };
        Ok(Ok(EditValidation {
            previews: plan
                .steps
                .iter()
                .map(|planned| EditPreview {
                    step_index: planned.index,
                    target: planned.target.clone(),
                    would_change: would_apply && planned.effect.is_some(),
                })
                .collect(),
            base_version: plan.base_version,
            would_apply,
        }))
    }

    /// Applies every step or none, resolving all targets against `expect_version`. Policy
    /// failures come back as an [`EditRefusal`] with the deck, history and id allocation
    /// untouched; `Err` means an internal failure. An applied batch commits one transaction.
    pub fn apply_edits(&self, request: &EditRequest) -> EditResult<EditOutcome> {
        let plan = match self.plan(request)? {
            Ok(plan) => plan,
            Err(refusal) => return Ok(Err(refusal)),
        };
        let staged = if plan.would_apply() {
            match self.stage_plan(&plan)? {
                Ok(staged) => staged,
                Err(refusal) => return Ok(Err(refusal)),
            }
        } else {
            None
        };
        let Some(staged) = staged else {
            let version = plan.base_version.clone();
            return Ok(Ok(plan.application(false, version)));
        };
        if let Err(refusal) = self.check_commit(&plan) {
            return Ok(Err(refusal));
        }
        let adoption = match request.history {
            EditHistory::Separate => Adoption::Tracked,
            EditHistory::None => Adoption::Untracked,
        };
        self.adopt(&staged.stage, staged.update, adoption)?;
        Ok(Ok(plan.application(true, self.version())))
    }
}
