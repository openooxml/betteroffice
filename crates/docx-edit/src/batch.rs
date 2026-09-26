//! Version-checked, all-or-nothing host edit batches.
//!
//! A batch resolves every step against one captured state, returns policy failures as data,
//! executes the resulting plan on a private clone that shares this replica's client id, rehearses
//! the clone's update against an untouched copy of the base, and adopts it as one transaction.

use std::collections::{BTreeSet, HashMap};
use std::fmt;
use std::sync::atomic::{AtomicU64, Ordering};

use serde::de::Error as _;
use serde::{Deserialize, Deserializer, Serialize};
use serde_json::Value;
use yrs::updates::decoder::Decode;
use yrs::{ReadTxn, StateVector, Transact, Update};

use crate::content_controls::{
    ContentControlSelector, ControlRecord, ControlValue, Site, ValueUnavailable, content_items,
    item_stamped,
};
use crate::ops::content_control::{
    ControlFill, ParagraphFill, TextRefusal, fill_text, formatting, inline_content, property_patch,
};
use crate::ops::paragraph::{ParagraphRecord, StylePlan, plan_paragraph_style};
use crate::ops::text::validate_text;
use crate::ops::utf16_len;
use crate::read_types::Anchor;
use crate::seed::{Restoration, SourceMetadata};
use crate::target::{
    EditTextView, ParagraphTarget, SearchScope, StoryView, TextRange, TextTarget, Views,
    comment_spans, field_result_blocks,
};
use crate::{
    EditCtx, EditError, EditResult, EditingDoc, FormatPolicy, Position, StoryRange, UndoSession,
    deterministic, story_ref,
};

const MAX_STEPS: usize = 128;
const MAX_INSERTED_UNITS: usize = 1_048_576;
const MAX_INSERTED_PARAGRAPHS: usize = 1_024;
const MAX_STAGING_BYTES: usize = 256 * 1024 * 1024;
/// Transaction origin of batches that stay out of local undo history.
const HOST_ORIGIN: &str = "host";

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

/// Who asked for a batch. Provenance only: history and revision authorship are separate.
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

/// Refuses a step unless its target currently reads exactly `text` (accepted view for
/// paragraph steps, the target's own view for text steps).
#[derive(Clone, Debug, Deserialize, Eq, PartialEq, Serialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct EditGuard {
    pub text: String,
}

/// Records a text step as a tracked change by `author` at the host-supplied `date`.
#[derive(Clone, Debug, Deserialize, Eq, PartialEq, Serialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct EditSuggestion {
    pub author: String,
    pub date: String,
}

/// One new paragraph; without `style_id` it takes the anchor paragraph's style.
#[derive(Clone, Debug, Deserialize, Eq, PartialEq, Serialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct ParagraphInput {
    pub text: String,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub style_id: Option<String>,
}

#[derive(Clone, Debug, Deserialize, Eq, PartialEq, Serialize)]
#[serde(
    tag = "op",
    rename_all = "camelCase",
    rename_all_fields = "camelCase",
    deny_unknown_fields
)]
#[non_exhaustive]
pub enum EditOperation {
    InsertText {
        target: TextTarget,
        at: TargetEdge,
        text: String,
    },
    ReplaceText {
        target: TextTarget,
        text: String,
    },
    DeleteText {
        target: TextTarget,
    },
    InsertParagraphs {
        target: ParagraphTarget,
        at: TargetEdge,
        paragraphs: Vec<ParagraphInput>,
    },
    /// Deletes the inclusive, contiguous span of complete paragraphs.
    DeleteParagraphs {
        story: String,
        first_para_id: String,
        last_para_id: String,
    },
    /// Applies a paragraph style defined in the document, with its paragraph and run effects.
    SetParagraphStyle {
        target: ParagraphTarget,
        style_id: String,
    },
    /// Replaces the content of a plain- or rich-text content control with plain text and clears
    /// its placeholder state. LF breaks lines inline and paragraphs in a block control.
    SetContentControlText {
        target: ContentControlSelector,
        text: String,
    },
}

/// One batch step; on the wire the operation's fields sit beside `expect` and `suggest`.
#[derive(Clone, Debug, Eq, PartialEq)]
pub struct EditStep {
    pub operation: EditOperation,
    pub expect: Option<EditGuard>,
    pub suggest: Option<EditSuggestion>,
}

impl EditStep {
    pub fn new(operation: EditOperation) -> Self {
        Self {
            operation,
            expect: None,
            suggest: None,
        }
    }
}

impl<'de> Deserialize<'de> for EditStep {
    fn deserialize<D: Deserializer<'de>>(deserializer: D) -> Result<Self, D::Error> {
        let mut value = Value::deserialize(deserializer)?;
        let object = value
            .as_object_mut()
            .ok_or_else(|| D::Error::custom("an edit step must be an object"))?;
        let expect = object.remove("expect").unwrap_or(Value::Null);
        let suggest = object.remove("suggest").unwrap_or(Value::Null);
        Ok(Self {
            expect: serde_json::from_value(expect).map_err(D::Error::custom)?,
            suggest: serde_json::from_value(suggest).map_err(D::Error::custom)?,
            operation: serde_json::from_value(value).map_err(D::Error::custom)?,
        })
    }
}

#[derive(Clone, Debug, Deserialize, Eq, PartialEq)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct EditRequest {
    pub expect_version: DocumentVersion,
    #[serde(default)]
    pub source: EditSource,
    #[serde(default)]
    pub history: EditHistory,
    pub steps: Vec<EditStep>,
}

impl EditRequest {
    /// Parses a request. A `setContentControlText` text holding an unpaired UTF-16 surrogate,
    /// which JSON written from JavaScript can escape but no Rust string can hold, fails its step
    /// as `invalid-text` instead of failing the parse.
    pub fn from_json(json: &str) -> Result<Result<Self, EditFailure>, serde_json::Error> {
        let error = match serde_json::from_str(json) {
            Ok(request) => return Ok(Ok(request)),
            Err(error) => error,
        };
        let Some(Ok(request)) =
            mark_lone_surrogates(json).map(|marked| serde_json::from_str::<Self>(&marked))
        else {
            return Err(error);
        };
        request
            .steps
            .iter()
            .enumerate()
            .find_map(|(index, step)| match &step.operation {
                EditOperation::SetContentControlText { target, text }
                    if text.contains(LONE_SURROGATE) =>
                {
                    Some(
                        failure(
                            EditFailureCode::InvalidStep,
                            "the text holds an unpaired UTF-16 surrogate".to_owned(),
                            Some(EditTarget::ContentControl {
                                selector: target.clone(),
                            }),
                        )
                        .because(EditFailureReason::InvalidText)
                        .at(index as u32),
                    )
                }
                _ => None,
            })
            .map(Err)
            .ok_or(error)
    }
}

/// What [`mark_lone_surrogates`] writes for an unpaired surrogate; content controls refuse it.
const LONE_SURROGATE: char = '\u{FFFF}';

/// `json` with each escaped unpaired UTF-16 surrogate replaced by [`LONE_SURROGATE`], or `None`
/// when it holds none.
fn mark_lone_surrogates(json: &str) -> Option<String> {
    let bytes = json.as_bytes();
    let unit = |at: usize| {
        (bytes.get(at) == Some(&b'\\') && bytes.get(at + 1) == Some(&b'u'))
            .then(|| json.get(at + 2..at + 6))
            .flatten()
            .and_then(|hex| u16::from_str_radix(hex, 16).ok())
    };
    let mut marked = String::with_capacity(json.len());
    let (mut copied, mut at) = (0, 0);
    while at < bytes.len() {
        if bytes[at] != b'\\' {
            at += 1;
            continue;
        }
        match unit(at) {
            Some(0xD800..=0xDBFF) if matches!(unit(at + 6), Some(0xDC00..=0xDFFF)) => at += 12,
            Some(0xD800..=0xDFFF) => {
                marked.push_str(&json[copied..at]);
                marked.push_str("\\uffff");
                at += 6;
                copied = at;
            }
            Some(_) => at += 6,
            None => at += 2,
        }
    }
    (copied > 0).then(|| {
        marked.push_str(&json[copied..]);
        marked
    })
}

#[derive(Clone, Copy, Debug, Deserialize, Eq, PartialEq, Serialize)]
#[serde(rename_all = "kebab-case")]
#[non_exhaustive]
pub enum EditFailureCode {
    StaleVersion,
    MissingTarget,
    AmbiguousTarget,
    ContentMismatch,
    OverlappingSteps,
    LockedTarget,
    TrackedRevisionConflict,
    Unsupported,
    InvalidStep,
    LimitExceeded,
}

/// Why a content-control step was refused, beside its failure code.
#[derive(Clone, Copy, Debug, Deserialize, Eq, PartialEq, Serialize)]
#[serde(rename_all = "kebab-case")]
pub enum EditFailureReason {
    MissingControl,
    MissingTag,
    AmbiguousTag,
    AmbiguousControlId,
    ContentLocked,
    BoundControl,
    UnsupportedControlType,
    UnsupportedChildren,
    NestedControls,
    UnknownLock,
    UnsupportedSuggestion,
    ProvenanceUnavailable,
    UnsupportedStory,
    MultilineNotAllowed,
    InvalidText,
}

/// What a failure or preview refers to: a step's text target or a paragraph span.
#[derive(Clone, Debug, Deserialize, Eq, PartialEq, Serialize)]
#[serde(
    tag = "kind",
    rename_all = "camelCase",
    rename_all_fields = "camelCase"
)]
#[non_exhaustive]
pub enum EditTarget {
    Paragraph(ParagraphTarget),
    Range(TextRange),
    Search {
        text: String,
        within: SearchScope,
        view: EditTextView,
    },
    Paragraphs {
        story: String,
        first_para_id: String,
        last_para_id: String,
    },
    ContentControl {
        selector: ContentControlSelector,
    },
}

/// The content control a step resolved to, and where it is.
#[derive(Clone, Debug, Deserialize, Eq, PartialEq, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct ResolvedControl {
    pub control_id: String,
    pub anchor: Anchor,
}

#[derive(Clone, Debug, Deserialize, Eq, PartialEq, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct EditFailure {
    pub code: EditFailureCode,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub reason: Option<EditFailureReason>,
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

    fn because(mut self, reason: EditFailureReason) -> Self {
        self.reason = Some(reason);
        self
    }
}

/// A policy refusal; the document is untouched at `version`.
#[derive(Clone, Debug, Deserialize, Eq, PartialEq, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct EditRefusal {
    pub version: DocumentVersion,
    pub failure: EditFailure,
}

pub(crate) fn failure(
    code: EditFailureCode,
    message: String,
    target: Option<EditTarget>,
) -> EditFailure {
    EditFailure {
        code,
        reason: None,
        step_index: None,
        conflicting_step_index: None,
        target: target.map(Box::new),
        message,
    }
}

pub(crate) fn refusal(version: DocumentVersion, failure: EditFailure) -> EditRefusal {
    EditRefusal { version, failure }
}

/// What one step did, in request order. Ranges use the final accepted view; a deletion reports
/// its collapsed boundary, and removed paragraph ids are history, not anchors.
#[derive(Clone, Debug, Deserialize, Eq, PartialEq, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct EditReceipt {
    pub step_index: u32,
    pub changed: bool,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub range: Option<TextRange>,
    pub new_paragraphs: Vec<ParagraphTarget>,
    pub removed_paragraphs: Vec<ParagraphTarget>,
    pub revision_ids: Vec<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub control: Option<ResolvedControl>,
}

/// What one step would do, resolved against the validated state. It reserves nothing.
#[derive(Clone, Debug, Deserialize, Eq, PartialEq, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct EditPreview {
    pub step_index: u32,
    pub target: EditTarget,
    pub would_change: bool,
    pub new_paragraph_count: u32,
    pub would_create_revisions: bool,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub control: Option<ResolvedControl>,
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
    /// False when every step was a no-op: no history, notification or version change.
    pub applied: bool,
    pub source: EditSource,
    pub changed_stories: Vec<String>,
    pub receipts: Vec<EditReceipt>,
}

/// JSON of a policy outcome: the success body or the refusal, tagged with `ok`.
#[cfg(feature = "wasm")]
pub(crate) fn outcome_json<T: Serialize, R: Serialize>(
    outcome: &Result<T, R>,
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

enum Effect {
    Insert {
        at: u32,
        text: String,
        ctx: EditCtx,
    },
    Replace {
        start: u32,
        end: u32,
        text: String,
        ctx: EditCtx,
    },
    Delete {
        start: u32,
        end: u32,
        ctx: EditCtx,
    },
    Paragraphs {
        at: u32,
        records: Vec<ParagraphRecord>,
    },
    RemoveParagraphs {
        start: u32,
        end: u32,
    },
    Style(StylePlan),
    Control(ControlFill),
}

impl Effect {
    /// The story index the effect starts at, for effects that shift later indices.
    fn shifting_start(&self) -> Option<u32> {
        match self {
            Self::Insert { at, .. } | Self::Paragraphs { at, .. } => Some(*at),
            Self::Replace { start, .. }
            | Self::Delete { start, .. }
            | Self::RemoveParagraphs { start, .. } => Some(*start),
            Self::Style(_) | Self::Control(_) => None,
        }
    }
}

/// A raw story interval or boundary a step mutates, in the captured state.
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

/// Where a step's receipt range comes from.
enum Shape {
    /// `len` final units starting at captured index `start`.
    Raw {
        start: u32,
        len: u32,
    },
    Paragraph(String),
    NewParagraphs,
    /// Every paragraph of the step's story.
    Story,
}

struct Planned {
    index: u32,
    story: String,
    target: EditTarget,
    claims: Vec<Claim>,
    touched: Vec<String>,
    exclusive: bool,
    effect: Option<Effect>,
    shape: Shape,
    removed: Vec<String>,
    new_paragraph_count: u32,
    suggest: bool,
    control: Option<ResolvedControl>,
    /// Another story the step changes: a block control's owning story.
    also_changes: Option<String>,
    /// Fills the copy of the step's control that another story reading its part holds; it
    /// has no preview or receipt of its own.
    companion: bool,
}

impl Planned {
    fn conflicts(&self, other: &Self) -> bool {
        if self.story != other.story {
            return false;
        }
        let shared = |owner: &Self, other: &Self| {
            owner.exclusive && owner.touched.iter().any(|id| other.touched.contains(id))
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
    source: EditSource,
    steps: Vec<Planned>,
}

/// The captured base state a batch stages against.
struct Base {
    update: Vec<u8>,
    state_vector: StateVector,
}

/// A plan executed on a private clone, with the update that adopts it.
struct Staged {
    stage: EditingDoc,
    executed: Vec<Option<Executed>>,
    changed_stories: Vec<String>,
    update: Vec<u8>,
}

fn unsupported(message: impl Into<String>, target: &EditTarget) -> EditFailure {
    failure(
        EditFailureCode::Unsupported,
        message.into(),
        Some(target.clone()),
    )
}

fn revision_conflict(message: impl Into<String>, target: &EditTarget) -> EditFailure {
    failure(
        EditFailureCode::TrackedRevisionConflict,
        message.into(),
        Some(target.clone()),
    )
}

fn invalid(message: impl Into<String>, target: &EditTarget) -> EditFailure {
    failure(
        EditFailureCode::InvalidStep,
        message.into(),
        Some(target.clone()),
    )
}

fn guard(expect: &Option<EditGuard>, actual: &str, target: &EditTarget) -> Result<(), EditFailure> {
    match expect {
        Some(expected) if expected.text != actual => Err(failure(
            EditFailureCode::ContentMismatch,
            format!(
                "the target reads {actual:?}, not the expected {:?}",
                expected.text
            ),
            Some(target.clone()),
        )),
        _ => Ok(()),
    }
}

fn edit_ctx(suggest: &Option<EditSuggestion>, target: &EditTarget) -> Result<EditCtx, EditFailure> {
    match suggest {
        None => Ok(EditCtx::local(String::new(), String::new())),
        Some(suggestion) if suggestion.author.is_empty() || suggestion.date.is_empty() => Err(
            invalid("a suggestion requires an author and a date", target),
        ),
        Some(suggestion) => {
            Ok(EditCtx::local(suggestion.author.clone(), suggestion.date.clone()).suggesting())
        }
    }
}

fn require_source<'a>(
    source: Option<&'a SourceMetadata>,
    target: &EditTarget,
) -> Result<&'a SourceMetadata, EditFailure> {
    source.ok_or_else(|| {
        unsupported(
            "paragraph structure and styles need a document opened from DOCX bytes in this session",
            target,
        )
    })
}

fn refuse_suggestion(step: &EditStep, target: &EditTarget) -> Result<(), EditFailure> {
    if step.suggest.is_some() {
        return Err(unsupported(
            "this operation cannot be recorded as a tracked change",
            target,
        ));
    }
    Ok(())
}

fn inserted_units(step: &EditStep) -> (usize, usize) {
    match &step.operation {
        EditOperation::InsertText { text, .. }
        | EditOperation::ReplaceText { text, .. }
        | EditOperation::SetContentControlText { text, .. } => (utf16_len(text) as usize, 0),
        EditOperation::InsertParagraphs { paragraphs, .. } => (
            paragraphs
                .iter()
                .map(|paragraph| utf16_len(&paragraph.text) as usize)
                .sum(),
            paragraphs.len(),
        ),
        _ => (0, 0),
    }
}

fn check_budget(request: &EditRequest) -> Result<(), EditFailure> {
    let limit = |message: String| failure(EditFailureCode::LimitExceeded, message, None);
    if request.steps.len() > MAX_STEPS {
        return Err(limit(format!("a batch holds at most {MAX_STEPS} steps")));
    }
    let (units, paragraphs) = request.steps.iter().map(inserted_units).fold(
        (0, 0),
        |(units, paragraphs), (more_units, more_paragraphs)| {
            (units + more_units, paragraphs + more_paragraphs)
        },
    );
    if units > MAX_INSERTED_UNITS {
        return Err(limit(format!(
            "a batch inserts at most {MAX_INSERTED_UNITS} UTF-16 units"
        )));
    }
    if paragraphs > MAX_INSERTED_PARAGRAPHS {
        return Err(limit(format!(
            "a batch inserts at most {MAX_INSERTED_PARAGRAPHS} paragraphs"
        )));
    }
    Ok(())
}

/// Text steps: insertion at an unambiguous boundary, or a contiguous run of plain text.
fn plan_text<T: ReadTxn>(
    views: &mut Views<'_, T>,
    source: Option<&SourceMetadata>,
    index: u32,
    step: &EditStep,
    target: &TextTarget,
    replacement: Option<(&str, Option<TargetEdge>)>,
) -> Result<Planned, EditFailure> {
    let requested = EditTarget::from(target.clone());
    let selection = views.text(target)?;
    guard(&step.expect, &selection.text(), &requested)?;
    let story = selection.view.story.clone();
    views.check_story_writable(&story, &requested)?;
    check_opaque_order(views, source, &story, &requested)?;
    if selection.view.structurally_revised(selection.paragraph) {
        return Err(revision_conflict(
            "the paragraph has a pending paragraph-mark revision",
            &requested,
        ));
    }
    if selection.paragraph().run_revisions {
        return Err(revision_conflict(
            "the paragraph's runs carry tracked formatting changes",
            &requested,
        ));
    }
    let ctx = edit_ctx(&step.suggest, &requested)?;
    let suggesting = step.suggest.is_some();
    if let Some((text, _)) = replacement {
        validate_text(text).map_err(|_| {
            invalid(
                "inserted text may not contain paragraph or line breaks; use insertParagraphs",
                &requested,
            )
        })?;
    }
    let resolved = EditTarget::Range(selection.range());
    let paragraph = selection.paragraph();
    let previous_mark_change = selection.paragraph > 0
        && paragraph.node_start == paragraph.start
        && selection.view.paragraphs[selection.paragraph - 1]
            .mark
            .property_change();
    let inline_unit = |raw: u32| raw >= paragraph.node_start && raw < paragraph.pilcrow;
    let point = |offset: u32| -> Result<u32, EditFailure> {
        let (left, right) = (paragraph.raw_after(offset), paragraph.raw_at(offset));
        if left != right {
            return Err(revision_conflict(
                "the boundary is ambiguous next to hidden tracked content",
                &requested,
            ));
        }
        let source = if offset > 0 {
            Some(left - 1)
        } else {
            (offset < paragraph.len()).then_some(right)
        };
        if source.is_some_and(|raw| paragraph.stamped_at(raw)) {
            return Err(revision_conflict(
                "the insertion would inherit formatting from a tracked change",
                &requested,
            ));
        }
        if suggesting
            && ((left > 0 && inline_unit(left - 1) && paragraph.stamped_at(left - 1))
                || (inline_unit(right) && paragraph.stamped_at(right))
                || paragraph.mark.property_change()
                || (offset == 0 && previous_mark_change))
        {
            return Err(revision_conflict(
                "a suggestion here would merge into an existing tracked change",
                &requested,
            ));
        }
        Ok(left)
    };
    let span = |start: u32, end: u32| -> Result<(u32, u32), EditFailure> {
        let (raw_start, raw_end) = (paragraph.raw_at(start), paragraph.raw_after(end));
        if raw_end - raw_start != end - start {
            return Err(revision_conflict(
                "the range bridges hidden tracked content",
                &requested,
            ));
        }
        if paragraph.has_stamped_in(start, end) {
            return Err(revision_conflict(
                "the range overlaps a tracked change",
                &requested,
            ));
        }
        if paragraph.has_atom_in(start, end) {
            return Err(unsupported(
                "replacing or deleting inline atoms is not supported",
                &requested,
            ));
        }
        if suggesting
            && ((raw_start > 0
                && inline_unit(raw_start - 1)
                && paragraph.stamped_at(raw_start - 1))
                || (inline_unit(raw_end) && paragraph.stamped_at(raw_end)))
        {
            return Err(revision_conflict(
                "a suggestion here would merge into an existing tracked change",
                &requested,
            ));
        }
        Ok((raw_start, raw_end))
    };
    let selected = selection.text();
    let (start, end) = (selection.start, selection.end);
    let (claims, effect, shape) = match replacement {
        Some((text, Some(edge))) => {
            let at = point(if edge == TargetEdge::Start {
                start
            } else {
                end
            })?;
            let effect = (!text.is_empty()).then(|| Effect::Insert {
                at,
                text: text.to_owned(),
                ctx,
            });
            let len = if effect.is_some() { utf16_len(text) } else { 0 };
            (
                vec![Claim::Point(at)],
                effect,
                Shape::Raw { start: at, len },
            )
        }
        Some((text, None)) => {
            let (raw_start, raw_end) = if start == end {
                let at = point(start)?;
                (at, at)
            } else {
                span(start, end)?
            };
            let claim = if raw_start == raw_end {
                Claim::Point(raw_start)
            } else {
                Claim::Span(raw_start, raw_end)
            };
            let changed = selected != text;
            let effect = changed.then(|| Effect::Replace {
                start: raw_start,
                end: raw_end,
                text: text.to_owned(),
                ctx,
            });
            let len = if changed {
                utf16_len(text)
            } else {
                raw_end - raw_start
            };
            (
                vec![claim],
                effect,
                Shape::Raw {
                    start: raw_start,
                    len,
                },
            )
        }
        None => {
            if start == end {
                let at = paragraph.raw_at(start);
                (
                    vec![Claim::Point(at)],
                    None,
                    Shape::Raw { start: at, len: 0 },
                )
            } else {
                let (raw_start, raw_end) = span(start, end)?;
                (
                    vec![Claim::Span(raw_start, raw_end)],
                    Some(Effect::Delete {
                        start: raw_start,
                        end: raw_end,
                        ctx,
                    }),
                    Shape::Raw {
                        start: raw_start,
                        len: 0,
                    },
                )
            }
        }
    };
    Ok(Planned {
        index,
        story,
        target: resolved,
        claims,
        touched: vec![paragraph.para_id.clone()],
        exclusive: false,
        effect,
        shape,
        removed: Vec::new(),
        new_paragraph_count: 0,
        suggest: suggesting,
        control: None,
        also_changes: None,
        companion: false,
    })
}

fn plan_insert_paragraphs<T: ReadTxn>(
    views: &mut Views<'_, T>,
    source: Option<&SourceMetadata>,
    index: u32,
    step: &EditStep,
    target: &ParagraphTarget,
    at: TargetEdge,
    paragraphs: &[ParagraphInput],
) -> Result<Planned, EditFailure> {
    let requested = EditTarget::Paragraph(target.clone());
    let (story, anchor_index) = views.paragraph(target, EditTextView::Accepted)?;
    let anchor = &story.paragraphs[anchor_index];
    guard(&step.expect, &anchor.text, &requested)?;
    views.check_story_writable(&target.story, &requested)?;
    refuse_suggestion(step, &requested)?;
    if story.structurally_revised(anchor_index) {
        return Err(revision_conflict(
            "the anchor paragraph has a pending paragraph-mark revision",
            &requested,
        ));
    }
    let source = require_source(source, &requested)?;
    check_opaque_order(views, Some(source), &target.story, &requested)?;
    let anchors = opaque_anchors(source, &story, &[], &requested)?;
    let (boundary, follower) = match at {
        TargetEdge::End => {
            let next = story
                .paragraphs
                .get(anchor_index + 1)
                .filter(|next| next.node_start == next.start);
            (anchor.pilcrow + 1, next.map(|next| next.para_id.as_str()))
        }
        TargetEdge::Start => (anchor.node_start, Some(anchor.para_id.as_str())),
    };
    if follower.is_some_and(|id| anchors.iter().any(|held| held == id)) {
        return Err(unsupported(
            "an opaque block sits at the insertion point",
            &requested,
        ));
    }
    let anchor_style = anchor.mark.style_id();
    let mut records = Vec::with_capacity(paragraphs.len());
    for paragraph in paragraphs {
        validate_text(&paragraph.text).map_err(|_| {
            invalid(
                "paragraph text may not contain paragraph or line breaks",
                &requested,
            )
        })?;
        if let Some(style_id) = &paragraph.style_id
            && !source.has_paragraph_style(style_id)
        {
            return Err(invalid(
                format!("style {style_id:?} is not a paragraph style of this document"),
                &requested,
            ));
        }
        let style = paragraph.style_id.as_deref().or(anchor_style.as_deref());
        let styled = source
            .styled_paragraph(style)
            .map_err(|message| invalid(message, &requested))?;
        if styled.properties.iter().any(|(key, _)| key == "numPr") {
            let named = style.map_or_else(
                || "the default paragraph style".to_owned(),
                |id| format!("style {id:?}"),
            );
            return Err(unsupported(
                format!(
                    "{named} defines list numbering, which v1 batches do not apply; {NUMBERING_FOLLOW_UP}"
                ),
                &requested,
            ));
        }
        records.push(ParagraphRecord {
            text: paragraph.text.clone(),
            properties: styled.properties,
            run: styled.run,
        });
    }
    let count = records.len() as u32;
    Ok(Planned {
        index,
        story: target.story.clone(),
        target: requested,
        claims: vec![Claim::Point(boundary)],
        touched: Vec::new(),
        exclusive: false,
        effect: (!records.is_empty()).then_some(Effect::Paragraphs {
            at: boundary,
            records,
        }),
        shape: Shape::NewParagraphs,
        removed: Vec::new(),
        new_paragraph_count: count,
        suggest: false,
        control: None,
        also_changes: None,
        companion: false,
    })
}

const NUMBERING_FOLLOW_UP: &str = "retaining numbering definitions for batches is a follow-up";

const DISPLACED_OPAQUE: &str =
    "saving would move an opaque XML block past the blocks that follow it";

fn opaque_restorations(
    source: &SourceMetadata,
    story: &StoryView,
    removed: &[String],
) -> Vec<Restoration> {
    let alive: std::collections::HashSet<&str> = story
        .paragraphs
        .iter()
        .map(|paragraph| paragraph.para_id.as_str())
        .filter(|id| !removed.iter().any(|removed| removed == id))
        .collect();
    source.opaque_restorations(&story.story, |id| alive.contains(id))
}

/// The paragraph the save projection restores each raw XML block of `story` in front of once
/// `removed` is gone. Refuses when a block would be restored past other blocks, or only by
/// position, which shifts with unrelated paragraph counts.
fn opaque_anchors(
    source: &SourceMetadata,
    story: &StoryView,
    removed: &[String],
    target: &EditTarget,
) -> Result<Vec<String>, EditFailure> {
    opaque_restorations(source, story, removed)
        .into_iter()
        .map(|restoration| match restoration {
            Restoration::Before(id) => Ok(id),
            Restoration::Displaced => Err(unsupported(DISPLACED_OPAQUE, target)),
            Restoration::Unanchored => Err(unsupported(
                "an opaque XML block in this story would lose the paragraph it is restored before",
                target,
            )),
        })
        .collect()
}

/// Refuses changing `story` when saving it or a story owning it, which the save projection
/// re-projects with it, would move a raw XML block past the blocks after it.
fn check_opaque_order<T: ReadTxn>(
    views: &mut Views<'_, T>,
    source: Option<&SourceMetadata>,
    story: &str,
    target: &EditTarget,
) -> Result<(), EditFailure> {
    let Some(source) = source else {
        return Ok(());
    };
    let ownership = views.ownership();
    let owners = ownership.chain(story).map_err(|message| {
        failure(
            EditFailureCode::LimitExceeded,
            message,
            Some(target.clone()),
        )
    })?;
    for story in std::iter::once(story).chain(owners.iter().map(|owner| owner.parent.as_str())) {
        if views
            .story(story, EditTextView::Accepted)
            .is_some_and(|view| {
                opaque_restorations(source, &view, &[])
                    .iter()
                    .any(|restoration| matches!(restoration, Restoration::Displaced))
            })
        {
            return Err(unsupported(DISPLACED_OPAQUE, target));
        }
    }
    Ok(())
}

fn bookmarks_balanced<'a>(
    paragraphs: impl Iterator<Item = &'a crate::target::ParagraphView>,
) -> bool {
    let mut open: Vec<String> = Vec::new();
    let mut closed: Vec<String> = Vec::new();
    for paragraph in paragraphs {
        let Some(yrs::Any::Array(bookmarks)) = paragraph.mark.properties.get("bookmarks") else {
            continue;
        };
        for bookmark in bookmarks.iter() {
            let yrs::Any::Map(bookmark) = bookmark else {
                continue;
            };
            let id = bookmark
                .get("id")
                .map(|id| format!("{id:?}"))
                .unwrap_or_default();
            match bookmark.get("kind") {
                Some(yrs::Any::String(kind)) if kind.as_ref() == "start" => open.push(id),
                Some(yrs::Any::String(kind)) if kind.as_ref() == "end" => closed.push(id),
                _ => {}
            }
        }
    }
    open.sort();
    closed.sort();
    open == closed
}

fn plan_delete_paragraphs<T: ReadTxn>(
    views: &mut Views<'_, T>,
    source: Option<&SourceMetadata>,
    index: u32,
    step: &EditStep,
    story_id: &str,
    first_para_id: &str,
    last_para_id: &str,
) -> Result<Planned, EditFailure> {
    let requested = EditTarget::Paragraphs {
        story: story_id.to_owned(),
        first_para_id: first_para_id.to_owned(),
        last_para_id: last_para_id.to_owned(),
    };
    let lookup = |views: &mut Views<'_, T>, para_id: &str| {
        views
            .paragraph(
                &ParagraphTarget {
                    story: story_id.to_owned(),
                    para_id: para_id.to_owned(),
                },
                EditTextView::Accepted,
            )
            .map_err(|mut failure| {
                failure.target = Some(Box::new(requested.clone()));
                failure
            })
    };
    let (story, first) = lookup(views, first_para_id)?;
    let (_, last) = lookup(views, last_para_id)?;
    if last < first {
        return Err(invalid("the last paragraph precedes the first", &requested));
    }
    let span = &story.paragraphs[first..=last];
    let joined: Vec<&str> = span
        .iter()
        .map(|paragraph| paragraph.text.as_str())
        .collect();
    guard(&step.expect, &joined.join("\n"), &requested)?;
    views.check_story_writable(story_id, &requested)?;
    refuse_suggestion(step, &requested)?;
    if let Some(offset) = (first..=last).find(|index| {
        story.structurally_revised(*index) || story.paragraphs[*index].has_revisions()
    }) {
        return Err(revision_conflict(
            format!(
                "paragraph {:?} carries a tracked change",
                story.paragraphs[offset].para_id
            ),
            &requested,
        ));
    }
    let source = require_source(source, &requested)?;
    check_opaque_order(views, Some(source), story_id, &requested)?;
    if first == 0 && last + 1 == story.paragraphs.len() {
        return Err(unsupported(
            "a story must keep at least one paragraph",
            &requested,
        ));
    }
    let head = &story.paragraphs[first];
    if last + 1 == story.paragraphs.len() && head.node_start > head.start {
        return Err(unsupported(
            "deleting the final paragraphs would leave a block at the end of the story",
            &requested,
        ));
    }
    for (offset, paragraph) in span.iter().enumerate() {
        if offset > 0 && paragraph.node_start > paragraph.start {
            return Err(unsupported(
                "the span contains a table, content control or break block",
                &requested,
            ));
        }
        if let Some((kind, _)) = paragraph
            .embeds
            .iter()
            .find(|(kind, _)| !matches!(kind.as_str(), "break" | "image"))
        {
            return Err(unsupported(
                format!("deleting the span would destroy a {kind:?} embed"),
                &requested,
            ));
        }
        if paragraph.mark.section() {
            return Err(unsupported("the span contains a section break", &requested));
        }
    }
    let ids: Vec<String> = span
        .iter()
        .map(|paragraph| paragraph.para_id.clone())
        .collect();
    let anchors = opaque_anchors(source, &story, &[], &requested)?;
    if anchors
        .iter()
        .any(|held| held != &head.para_id && ids.contains(held))
    {
        return Err(unsupported("the span contains an opaque block", &requested));
    }
    let follower = story
        .paragraphs
        .get(last + 1)
        .filter(|next| next.node_start == next.start);
    if opaque_anchors(source, &story, &ids, &requested)?
        .iter()
        .zip(&anchors)
        .any(|(moved, held)| moved != held && follower.is_none_or(|next| moved != &next.para_id))
    {
        return Err(unsupported(
            "deleting the span would move an opaque block past other content",
            &requested,
        ));
    }
    if !bookmarks_balanced(span.iter()) {
        return Err(unsupported(
            "a bookmark crosses the span boundary",
            &requested,
        ));
    }
    let bound = field_result_blocks(&story, views.txn());
    if span
        .iter()
        .any(|paragraph| bound.contains(&paragraph.para_id))
    {
        return Err(unsupported(
            "the span holds a field's cached result",
            &requested,
        ));
    }
    let (start, end) = (head.node_start, story.paragraphs[last].pilcrow + 1);
    if comment_spans(views.txn(), story_id)
        .into_iter()
        .any(|(from, to)| from < end && to > start)
    {
        return Err(unsupported("a comment is anchored in the span", &requested));
    }
    Ok(Planned {
        index,
        story: story_id.to_owned(),
        target: requested,
        claims: vec![
            Claim::Span(start, end),
            Claim::Point(start),
            Claim::Point(end),
        ],
        touched: ids.clone(),
        exclusive: true,
        effect: Some(Effect::RemoveParagraphs { start, end }),
        shape: Shape::Raw { start, len: 0 },
        removed: ids,
        new_paragraph_count: 0,
        suggest: false,
        control: None,
        also_changes: None,
        companion: false,
    })
}

fn plan_style<T: ReadTxn>(
    views: &mut Views<'_, T>,
    source: Option<&SourceMetadata>,
    index: u32,
    step: &EditStep,
    target: &ParagraphTarget,
    style_id: &str,
) -> Result<Planned, EditFailure> {
    let requested = EditTarget::Paragraph(target.clone());
    let (story, paragraph_index) = views.paragraph(target, EditTextView::Accepted)?;
    let paragraph = &story.paragraphs[paragraph_index];
    guard(&step.expect, &paragraph.text, &requested)?;
    views.check_story_writable(&target.story, &requested)?;
    refuse_suggestion(step, &requested)?;
    if story.structurally_revised(paragraph_index) || paragraph.has_revisions() {
        return Err(revision_conflict(
            "the paragraph carries a tracked change",
            &requested,
        ));
    }
    let source = require_source(source, &requested)?;
    check_opaque_order(views, Some(source), &target.story, &requested)?;
    if !source.has_paragraph_style(style_id) {
        return Err(invalid(
            format!("style {style_id:?} is not a paragraph style of this document"),
            &requested,
        ));
    }
    let next = source
        .styled_paragraph(Some(style_id))
        .map_err(|message| invalid(message, &requested))?;
    if paragraph.mark.numbering() {
        return Err(unsupported(
            format!("v1 batches do not restyle numbered paragraphs; {NUMBERING_FOLLOW_UP}"),
            &requested,
        ));
    }
    if next.properties.iter().any(|(key, _)| key == "numPr") {
        return Err(unsupported(
            format!(
                "style {style_id:?} defines list numbering, which v1 batches do not apply; {NUMBERING_FOLLOW_UP}"
            ),
            &requested,
        ));
    }
    let previous = source
        .styled_paragraph(paragraph.mark.style_id().as_deref())
        .map_err(|message| invalid(message, &requested))?;
    let text = story_ref(views.txn(), &target.story).map_err(|error| {
        failure(
            EditFailureCode::MissingTarget,
            error.to_string(),
            Some(requested.clone()),
        )
    })?;
    let chunks = views
        .doc()
        .chunk_snapshot(&target.story, &text, views.txn());
    let plan = plan_paragraph_style(
        &chunks,
        &paragraph.para_id,
        paragraph.pilcrow,
        paragraph.node_start,
        &paragraph.mark.properties,
        &previous,
        &next,
        style_id,
    );
    Ok(Planned {
        index,
        story: target.story.clone(),
        target: requested,
        claims: Vec::new(),
        touched: vec![paragraph.para_id.clone()],
        exclusive: true,
        effect: (!plan.is_empty()).then_some(Effect::Style(plan)),
        shape: Shape::Paragraph(paragraph.para_id.clone()),
        removed: Vec::new(),
        new_paragraph_count: 0,
        suggest: false,
        control: None,
        also_changes: None,
        companion: false,
    })
}

fn control_failure(
    code: EditFailureCode,
    reason: EditFailureReason,
    message: impl Into<String>,
    target: &EditTarget,
) -> EditFailure {
    failure(code, message.into(), Some(target.clone())).because(reason)
}

/// The one control `selector` names in the complete inventory, independent of read filters.
fn resolve_control<'a>(
    inventory: &'a crate::content_controls::Inventory,
    selector: &ContentControlSelector,
    target: &EditTarget,
) -> Result<(usize, &'a ControlRecord), EditFailure> {
    let addressable = || {
        inventory
            .records
            .iter()
            .enumerate()
            .filter(|(_, record)| !record.alias)
    };
    let (matches, missing, ambiguous, named): (Vec<(usize, &ControlRecord)>, _, _, String) =
        match selector {
            ContentControlSelector::Id { control_id } => (
                addressable()
                    .filter(|(_, record)| record.id() == control_id)
                    .collect(),
                EditFailureReason::MissingControl,
                EditFailureReason::AmbiguousControlId,
                format!("control id {control_id:?}"),
            ),
            ContentControlSelector::Tag { tag } => {
                if inventory.unidentified {
                    return Err(control_failure(
                        EditFailureCode::Unsupported,
                        EditFailureReason::ProvenanceUnavailable,
                        format!(
                            "content controls may exist that this session cannot see, so tag {tag:?} cannot be shown to be unique"
                        ),
                        target,
                    ));
                }
                (
                    addressable()
                        .filter(|(_, record)| {
                            record.control.metadata.tag.as_deref() == Some(tag.as_str())
                        })
                        .collect(),
                    EditFailureReason::MissingTag,
                    EditFailureReason::AmbiguousTag,
                    format!("tag {tag:?}"),
                )
            }
        };
    match matches.as_slice() {
        [] => Err(control_failure(
            EditFailureCode::MissingTarget,
            missing,
            format!("no content control has {named}"),
            target,
        )),
        [found] => Ok(*found),
        many => Err(control_failure(
            EditFailureCode::AmbiguousTarget,
            ambiguous,
            format!("{} content controls have {named}", many.len()),
            target,
        )),
    }
}

/// The parsed `w:rPr` of a control's own properties.
fn control_run_properties(payload: &HashMap<String, yrs::Any>) -> Option<Value> {
    if let Some(yrs::Any::String(json)) = payload.get("propertiesJson") {
        return serde_json::from_str::<Value>(json)
            .ok()?
            .get("runProperties")
            .cloned();
    }
    let yrs::Any::String(raw) = payload.get("rawPropertiesXml")? else {
        return None;
    };
    serde_json::to_value(
        docx_parse::parse_sdt_properties_xml(raw)
            .ok()?
            .run_properties?,
    )
    .ok()
}

fn plan_control<T: ReadTxn>(
    views: &mut Views<'_, T>,
    source: Option<&SourceMetadata>,
    index: u32,
    step: &EditStep,
    selector: &ContentControlSelector,
    text: &str,
) -> Result<Vec<Planned>, EditFailure> {
    let requested = EditTarget::ContentControl {
        selector: selector.clone(),
    };
    let inventory = views.controls()?;
    let (position, record) = resolve_control(&inventory, selector, &requested)?;
    let control = &record.control;
    let resolved = EditTarget::ContentControl {
        selector: ContentControlSelector::Id {
            control_id: record.id().to_owned(),
        },
    };
    let refuse = |code: EditFailureCode, reason: EditFailureReason, message: &str| {
        control_failure(code, reason, message, &resolved)
    };
    let current = match &control.value {
        ControlValue::Text { text } => Some(text.as_str()),
        ControlValue::Unavailable { .. } => None,
    };
    if let Some(expected) = &step.expect
        && current != Some(expected.text.as_str())
    {
        return Err(failure(
            EditFailureCode::ContentMismatch,
            match current {
                Some(actual) => format!(
                    "the control reads {actual:?}, not the expected {:?}",
                    expected.text
                ),
                None => "the control has no text value to compare".to_owned(),
            },
            Some(resolved.clone()),
        ));
    }
    if step.suggest.is_some() {
        return Err(refuse(
            EditFailureCode::Unsupported,
            EditFailureReason::UnsupportedSuggestion,
            "filling a content control cannot be recorded as a tracked change",
        ));
    }
    if matches!(record.site, Site::Source) {
        return Err(refuse(
            EditFailureCode::Unsupported,
            EditFailureReason::UnsupportedStory,
            "the control lives in source content the session does not edit",
        ));
    }
    if !matches!(
        control.metadata.control_type.as_str(),
        "plainText" | "richText"
    ) {
        return Err(refuse(
            EditFailureCode::Unsupported,
            EditFailureReason::UnsupportedControlType,
            &format!(
                "{} controls are not filled with text",
                control.metadata.control_type
            ),
        ));
    }
    if control.metadata.data_bound {
        return Err(refuse(
            EditFailureCode::Unsupported,
            EditFailureReason::BoundControl,
            "the control is bound to custom XML data",
        ));
    }
    if !control.effective_lock.known {
        return Err(refuse(
            EditFailureCode::Unsupported,
            EditFailureReason::UnknownLock,
            "the control or a control containing it carries an unknown lock",
        ));
    }
    if inventory.divergent.contains(&position) {
        return Err(refuse(
            EditFailureCode::Unsupported,
            EditFailureReason::UnsupportedStory,
            "another story reads the control's part but now holds different controls, so a fill cannot keep both in step",
        ));
    }
    if record
        .safety
        .as_ref()
        .is_some_and(|safety| safety.uncertain)
    {
        return Err(refuse(
            EditFailureCode::Unsupported,
            EditFailureReason::UnknownLock,
            "the control's lock or data binding cannot be read reliably from its XML",
        ));
    }
    if control.effective_lock.content {
        return Err(refuse(
            EditFailureCode::LockedTarget,
            EditFailureReason::ContentLocked,
            "the control's content is locked",
        ));
    }
    if control.parent_control_id.is_some() || record.children {
        return Err(refuse(
            EditFailureCode::Unsupported,
            EditFailureReason::NestedControls,
            "controls nested in or containing other controls are not filled",
        ));
    }
    let source = source.ok_or_else(|| {
        refuse(
            EditFailureCode::Unsupported,
            EditFailureReason::ProvenanceUnavailable,
            "filling a control needs a document opened from DOCX bytes in this session",
        )
    })?;
    match &record.safety {
        None => {
            return Err(refuse(
                EditFailureCode::Unsupported,
                EditFailureReason::ProvenanceUnavailable,
                "the control could not be matched to its source content, so a fill cannot be shown to be safe",
            ));
        }
        Some(safety) if safety.revisions => {
            return Err(revision_conflict(
                "the control's source content carries tracked changes",
                &resolved,
            ));
        }
        Some(safety) if !safety.safe() => {
            return Err(refuse(
                EditFailureCode::Unsupported,
                EditFailureReason::UnsupportedChildren,
                &format!(
                    "the control holds {} that a text fill would destroy",
                    safety.unsupported.join(", ")
                ),
            ));
        }
        Some(_) => {}
    }
    if record.stamped
        || matches!(
            control.value,
            ControlValue::Unavailable {
                reason: ValueUnavailable::TrackedRevisions
            }
        )
    {
        return Err(revision_conflict(
            "the control carries a tracked change",
            &resolved,
        ));
    }
    if matches!(
        control.value,
        ControlValue::Unavailable {
            reason: ValueUnavailable::NonTextContent
        }
    ) {
        return Err(refuse(
            EditFailureCode::Unsupported,
            EditFailureReason::UnsupportedChildren,
            "the control holds content other than text, tabs and line breaks",
        ));
    }
    let allow_breaks =
        control.metadata.control_type == "richText" || control.multi_line == Some(true);
    let text = fill_text(text, allow_breaks).map_err(|refusal| match refusal {
        TextRefusal::Invalid(message) => refuse(
            EditFailureCode::InvalidStep,
            EditFailureReason::InvalidText,
            &message,
        ),
        TextRefusal::Multiline => refuse(
            EditFailureCode::InvalidStep,
            EditFailureReason::MultilineNotAllowed,
            "the plain-text control does not accept line breaks",
        ),
    })?;
    let copies = inventory.companions.get(&position).into_iter().flatten();
    let mut planned = Vec::new();
    for (copy, target) in std::iter::once(record)
        .chain(copies.map(|copy| &inventory.records[*copy]))
        .enumerate()
    {
        let patch = property_patch(&target.payload).map_err(|message| {
            failure(
                EditFailureCode::Unsupported,
                message,
                Some(resolved.clone()),
            )
        })?;
        let unchanged = match &target.control.value {
            ControlValue::Text { text: current } => *current == text,
            ControlValue::Unavailable { .. } => false,
        };
        let fill = Fill {
            index,
            record: target,
            resolved: resolved.clone(),
            source,
            unchanged,
            text: text.clone(),
            patch,
            drop_value: target
                .payload
                .get("value")
                .is_some_and(|value| !matches!(value, yrs::Any::Null | yrs::Any::Undefined)),
        };
        let mut step = match &target.site {
            Site::Inline { raw, para_id } => plan_inline_fill(views, fill, *raw, para_id),
            Site::Block { raw, child } => plan_block_fill(views, fill, *raw, child),
            Site::Nested | Site::Source => Err(fill.refuse(
                EditFailureCode::Unsupported,
                EditFailureReason::NestedControls,
                "the control cannot be filled",
            )),
        }?;
        step.companion = copy > 0;
        planned.push(step);
    }
    Ok(planned)
}

/// A fill whose target, guard and text passed the checks both placements share.
struct Fill<'a> {
    index: u32,
    record: &'a ControlRecord,
    resolved: EditTarget,
    source: &'a SourceMetadata,
    text: String,
    patch: Vec<(String, yrs::Any)>,
    unchanged: bool,
    drop_value: bool,
}

impl Fill<'_> {
    fn refuse(
        &self,
        code: EditFailureCode,
        reason: EditFailureReason,
        message: &str,
    ) -> EditFailure {
        control_failure(code, reason, message, &self.resolved)
    }

    /// The formatting typed text takes in a paragraph of `style` inside this control.
    fn control_run(&self, style: Option<&str>) -> Result<Vec<(String, yrs::Any)>, EditFailure> {
        self.source
            .control_run(style, control_run_properties(&self.record.payload).as_ref())
            .map_err(|message| invalid(message, &self.resolved))
    }

    fn placeholder(&self) -> bool {
        self.record.control.metadata.showing_placeholder
    }

    fn writable<T: ReadTxn>(
        &self,
        views: &mut Views<'_, T>,
        story: &str,
    ) -> Result<(), EditFailure> {
        views
            .check_story_writable(story, &self.resolved)
            .map_err(|failure| match failure.code {
                EditFailureCode::LockedTarget => failure.because(EditFailureReason::ContentLocked),
                _ => failure,
            })?;
        check_opaque_order(views, Some(self.source), story, &self.resolved)
    }

    fn resolved_control(&self) -> Option<ResolvedControl> {
        Some(ResolvedControl {
            control_id: self.record.id().to_owned(),
            anchor: self.record.control.anchor.clone(),
        })
    }
}

/// An inline fill replaces the control embed's frozen content in place.
fn plan_inline_fill<T: ReadTxn>(
    views: &mut Views<'_, T>,
    fill: Fill<'_>,
    raw: u32,
    para_id: &str,
) -> Result<Planned, EditFailure> {
    let story = fill.record.story.clone();
    fill.writable(views, &story)?;
    let (view, paragraph) = views.paragraph(
        &ParagraphTarget {
            story: story.clone(),
            para_id: para_id.to_owned(),
        },
        EditTextView::Accepted,
    )?;
    let paragraph = &view.paragraphs[paragraph];
    if paragraph.run_revisions {
        return Err(revision_conflict(
            "the paragraph's runs carry tracked formatting changes",
            &fill.resolved,
        ));
    }
    let content = if fill.unchanged {
        None
    } else {
        let first = content_items(&fill.record.payload)
            .iter()
            .filter_map(|item| match item {
                yrs::Any::Map(item) => Some(item),
                _ => None,
            })
            .find(|item| {
                matches!(item.get("kind"), Some(yrs::Any::String(kind)) if kind.as_ref() == "text")
            });
        let attrs = match first {
            Some(item) if !fill.placeholder() && !item_stamped(item) => match item.get("attrs") {
                Some(yrs::Any::Map(attrs)) => formatting(attrs),
                _ => Vec::new(),
            },
            _ => fill.control_run(paragraph.mark.style_id().as_deref())?,
        };
        Some(inline_content(&fill.text, &attrs))
    };
    let control = fill.resolved_control();
    let effect = (content.is_some() || !fill.patch.is_empty() || fill.drop_value).then(|| {
        Effect::Control(ControlFill::Inline {
            raw,
            content,
            drop_value: fill.drop_value,
            patch: fill.patch,
        })
    });
    Ok(Planned {
        index: fill.index,
        story,
        target: fill.resolved,
        claims: vec![Claim::Span(raw, raw + 1)],
        touched: vec![para_id.to_owned()],
        exclusive: false,
        effect,
        shape: Shape::Raw { start: raw, len: 1 },
        removed: Vec::new(),
        new_paragraph_count: 0,
        suggest: false,
        control,
        also_changes: None,
        companion: false,
    })
}

/// A block fill rewrites the control's child story: surviving paragraphs keep their ids and
/// properties, extra ones take the first paragraph's style defaults.
fn plan_block_fill<T: ReadTxn>(
    views: &mut Views<'_, T>,
    fill: Fill<'_>,
    raw: u32,
    child: &str,
) -> Result<Planned, EditFailure> {
    fill.writable(views, child)?;
    let unsupported = |message: &str| {
        fill.refuse(
            EditFailureCode::Unsupported,
            EditFailureReason::UnsupportedChildren,
            message,
        )
    };
    let view = views
        .story(child, EditTextView::Accepted)
        .ok_or_else(|| unsupported("the control's story was not found"))?;
    if !fill.source.opaque_restorations(child, |_| true).is_empty() {
        return Err(unsupported("the control holds raw XML blocks"));
    }
    for (position, paragraph) in view.paragraphs.iter().enumerate() {
        if view.structurally_revised(position) || paragraph.has_revisions() {
            return Err(revision_conflict(
                "a paragraph of the control carries a tracked change",
                &fill.resolved,
            ));
        }
        if paragraph.node_start > paragraph.start {
            return Err(unsupported(
                "the control holds a table, a nested control or a break block",
            ));
        }
        if let Some((kind, _)) = paragraph.embeds.iter().find(|(kind, _)| kind != "break") {
            return Err(unsupported(&format!("the control holds a {kind:?} embed")));
        }
        if paragraph.mark.section() {
            return Err(unsupported("the control holds a section break"));
        }
        if matches!(paragraph.mark.properties.get("bookmarks"), Some(yrs::Any::Array(bookmarks)) if !bookmarks.is_empty())
        {
            return Err(unsupported("the control holds a bookmark"));
        }
    }
    let bound = field_result_blocks(&view, views.txn());
    if view
        .paragraphs
        .iter()
        .any(|paragraph| bound.contains(&paragraph.para_id))
    {
        return Err(unsupported("the control holds a field's cached result"));
    }
    if !comment_spans(views.txn(), child).is_empty() {
        return Err(unsupported("a comment is anchored in the control"));
    }
    let (Some(first), Some(last)) = (view.paragraphs.first(), view.paragraphs.last()) else {
        return Err(unsupported("the control's story holds no paragraph"));
    };
    let lines: Vec<&str> = fill.text.split('\n').collect();
    if lines.len().saturating_sub(view.paragraphs.len()) > MAX_INSERTED_PARAGRAPHS {
        return Err(failure(
            EditFailureCode::LimitExceeded,
            format!("a batch inserts at most {MAX_INSERTED_PARAGRAPHS} paragraphs"),
            Some(fill.resolved.clone()),
        ));
    }
    let text = story_ref(views.txn(), child)
        .map_err(|_| unsupported("the control's story was not found"))?;
    let chunks = views.doc().chunk_snapshot(child, &text, views.txn());
    let first_text = |paragraph: &crate::target::ParagraphView| {
        chunks
            .iter()
            .filter(|chunk| chunk.start >= paragraph.node_start && chunk.end() <= paragraph.pilcrow)
            .find(|chunk| matches!(chunk.kind, crate::ops::ChunkKind::Text(_)))
            .map(|chunk| formatting(&chunk.attrs.clone().into_iter().collect()))
    };
    let baseline = match first_text(first) {
        Some(attrs) if !fill.placeholder() => attrs,
        _ => fill.control_run(first.mark.style_id().as_deref())?,
    };
    let mut paragraphs = Vec::new();
    let mut remove = None;
    let mut insert = None;
    let mut removed = Vec::new();
    let mut new_paragraph_count = 0;
    if !fill.unchanged {
        for (paragraph, line) in view.paragraphs.iter().zip(&lines) {
            if paragraph.embeds.is_empty() && paragraph.text == *line {
                continue;
            }
            let attrs = match first_text(paragraph) {
                Some(attrs) if !fill.placeholder() => attrs,
                _ => fill.control_run(paragraph.mark.style_id().as_deref())?,
            };
            paragraphs.push(ParagraphFill {
                node_start: paragraph.node_start,
                pilcrow: paragraph.pilcrow,
                text: (*line).to_owned(),
                attrs,
            });
        }
        if lines.len() < view.paragraphs.len() {
            remove = Some((view.paragraphs[lines.len()].start, last.pilcrow + 1));
            removed = view.paragraphs[lines.len()..]
                .iter()
                .map(|paragraph| paragraph.para_id.clone())
                .collect();
        }
        if lines.len() > view.paragraphs.len() {
            let styled = fill
                .source
                .styled_paragraph(first.mark.style_id().as_deref())
                .map_err(|message| invalid(message, &fill.resolved))?;
            if styled.properties.iter().any(|(key, _)| key == "numPr") {
                return Err(failure(
                    EditFailureCode::Unsupported,
                    format!(
                        "the control's paragraph style defines list numbering, which v1 batches do not apply to new paragraphs; {NUMBERING_FOLLOW_UP}"
                    ),
                    Some(fill.resolved.clone()),
                ));
            }
            let records: Vec<ParagraphRecord> = lines[view.paragraphs.len()..]
                .iter()
                .map(|line| ParagraphRecord {
                    text: (*line).to_owned(),
                    properties: styled.properties.clone(),
                    run: baseline.clone(),
                })
                .collect();
            new_paragraph_count = records.len() as u32;
            insert = Some((last.pilcrow + 1, records));
        }
    }
    let control = fill.resolved_control();
    let parent = fill.record.story.clone();
    let changes_content = !paragraphs.is_empty() || remove.is_some() || insert.is_some();
    let effect = (changes_content || !fill.patch.is_empty() || fill.drop_value).then(|| {
        Effect::Control(ControlFill::Block {
            parent: parent.clone(),
            raw,
            drop_value: fill.drop_value,
            patch: fill.patch,
            child: child.to_owned(),
            paragraphs,
            remove,
            insert,
        })
    });
    Ok(Planned {
        index: fill.index,
        story: child.to_owned(),
        target: fill.resolved,
        claims: vec![Claim::Span(0, last.pilcrow + 1)],
        touched: view
            .paragraphs
            .iter()
            .map(|paragraph| paragraph.para_id.clone())
            .collect(),
        exclusive: true,
        effect,
        shape: Shape::Story,
        removed,
        new_paragraph_count,
        suggest: false,
        control,
        also_changes: Some(parent),
        companion: false,
    })
}

fn plan_step<T: ReadTxn>(
    views: &mut Views<'_, T>,
    source: Option<&SourceMetadata>,
    index: u32,
    step: &EditStep,
) -> Result<Vec<Planned>, EditFailure> {
    let planned = match &step.operation {
        EditOperation::InsertText { target, at, text } => {
            plan_text(views, source, index, step, target, Some((text, Some(*at))))
        }
        EditOperation::ReplaceText { target, text } => {
            plan_text(views, source, index, step, target, Some((text, None)))
        }
        EditOperation::DeleteText { target } => plan_text(views, source, index, step, target, None),
        EditOperation::InsertParagraphs {
            target,
            at,
            paragraphs,
        } => plan_insert_paragraphs(views, source, index, step, target, *at, paragraphs),
        EditOperation::DeleteParagraphs {
            story,
            first_para_id,
            last_para_id,
        } => plan_delete_paragraphs(
            views,
            source,
            index,
            step,
            story,
            first_para_id,
            last_para_id,
        ),
        EditOperation::SetParagraphStyle { target, style_id } => {
            plan_style(views, source, index, step, target, style_id)
        }
        EditOperation::SetContentControlText { target, text } => {
            return plan_control(views, source, index, step, target, text);
        }
    }?;
    Ok(vec![planned])
}

/// One executed step: the story-length change and what the operation minted.
#[derive(Default)]
struct Executed {
    delta: i64,
    new_ids: Vec<String>,
    revision_ids: Vec<String>,
}

fn staging_error(index: u32, error: impl fmt::Display) -> EditError {
    EditError::InvalidUpdate(format!("staged step {index} failed: {error}"))
}

fn execute(stage: &EditingDoc, steps: &[Planned]) -> EditResult<Vec<Option<Executed>>> {
    let mut executed: Vec<Option<Executed>> = steps.iter().map(|_| None).collect();
    for (slot, planned) in steps.iter().enumerate() {
        match &planned.effect {
            Some(Effect::Style(plan)) => {
                stage
                    .apply_style_plan(&planned.story, plan)
                    .map_err(|error| staging_error(planned.index, error))?;
                executed[slot] = Some(Executed::default());
            }
            Some(Effect::Control(fill)) => {
                let new_ids = stage
                    .apply_control_fill(&planned.story, fill)
                    .map_err(|error| staging_error(planned.index, error))?;
                executed[slot] = Some(Executed {
                    new_ids,
                    ..Executed::default()
                });
            }
            _ => {}
        }
    }
    let mut order: Vec<usize> = (0..steps.len())
        .filter(|slot| {
            steps[*slot]
                .effect
                .as_ref()
                .is_some_and(|effect| effect.shifting_start().is_some())
        })
        .collect();
    order.sort_by_key(|slot| {
        std::cmp::Reverse(
            steps[*slot]
                .effect
                .as_ref()
                .and_then(Effect::shifting_start),
        )
    });
    for slot in order {
        let planned = &steps[slot];
        let story = planned.story.as_str();
        let before = stage.story_len(story)?;
        let mut outcome = Executed::default();
        let fail = |error: crate::OpError| staging_error(planned.index, error);
        match planned.effect.as_ref() {
            Some(Effect::Insert { at, text, ctx }) => {
                let receipt = stage
                    .insert_text(ctx, Position::new(story, *at), text, FormatPolicy::Inherit)
                    .map_err(fail)?;
                outcome.revision_ids = receipt.revision_ids;
            }
            Some(Effect::Replace {
                start,
                end,
                text,
                ctx,
            }) => {
                let receipt = stage
                    .replace_range(ctx, StoryRange::new(story, *start, *end), text)
                    .map_err(fail)?;
                outcome.revision_ids = receipt.revision_ids;
            }
            Some(Effect::Delete { start, end, ctx }) => {
                let receipt = stage
                    .delete_range(ctx, StoryRange::new(story, *start, *end))
                    .map_err(fail)?;
                outcome.revision_ids = receipt.revision_ids;
            }
            Some(Effect::Paragraphs { at, records }) => {
                outcome.new_ids = stage
                    .insert_paragraph_records(story, *at, records)
                    .map_err(fail)?;
            }
            Some(Effect::RemoveParagraphs { start, end }) => {
                stage
                    .remove_paragraph_records(story, *start, *end)
                    .map_err(fail)?;
            }
            Some(Effect::Style(_) | Effect::Control(_)) | None => continue,
        }
        outcome.delta = i64::from(stage.story_len(story)?) - i64::from(before);
        executed[slot] = Some(outcome);
    }
    Ok(executed)
}

fn shifted(steps: &[Planned], executed: &[Option<Executed>], story: &str, raw: u32) -> u32 {
    let delta: i64 = steps
        .iter()
        .zip(executed)
        .filter(|(planned, _)| planned.story == story)
        .filter_map(|(planned, outcome)| {
            let start = planned.effect.as_ref()?.shifting_start()?;
            (start < raw).then(|| outcome.as_ref().map_or(0, |outcome| outcome.delta))
        })
        .sum();
    (i64::from(raw) + delta).max(0) as u32
}

fn receipts<T: ReadTxn>(
    views: &mut Views<'_, T>,
    steps: &[Planned],
    executed: &[Option<Executed>],
) -> Vec<EditReceipt> {
    steps
        .iter()
        .zip(executed)
        .filter(|(planned, _)| !planned.companion)
        .map(|(planned, outcome)| {
            let story = views.story(&planned.story, EditTextView::Accepted);
            let paragraph_range = |story: &StoryView, para_id: &str| {
                let paragraph = story
                    .paragraphs
                    .iter()
                    .find(|paragraph| paragraph.para_id == para_id)?;
                Some((paragraph.para_id.clone(), paragraph.len()))
            };
            let range = story.as_deref().and_then(|story| match &planned.shape {
                Shape::Raw { start, len } => {
                    let start = shifted(steps, executed, &planned.story, *start);
                    story.range_of_raw(start, start + len)
                }
                Shape::Paragraph(para_id) => {
                    let (para_id, len) = paragraph_range(story, para_id)?;
                    Some(text_range(&planned.story, &para_id, 0, &para_id, len))
                }
                Shape::NewParagraphs => {
                    let ids = &outcome.as_ref()?.new_ids;
                    let (first, _) = paragraph_range(story, ids.first()?)?;
                    let (last, len) = paragraph_range(story, ids.last()?)?;
                    Some(text_range(&planned.story, &first, 0, &last, len))
                }
                Shape::Story => {
                    let first = story.paragraphs.first()?;
                    let last = story.paragraphs.last()?;
                    Some(text_range(
                        &planned.story,
                        &first.para_id,
                        0,
                        &last.para_id,
                        last.len(),
                    ))
                }
            });
            let paragraphs = |ids: &[String]| {
                ids.iter()
                    .map(|para_id| ParagraphTarget {
                        story: planned.story.clone(),
                        para_id: para_id.clone(),
                    })
                    .collect()
            };
            EditReceipt {
                step_index: planned.index,
                changed: outcome.is_some(),
                range,
                new_paragraphs: outcome
                    .as_ref()
                    .map_or_else(Vec::new, |outcome| paragraphs(&outcome.new_ids)),
                removed_paragraphs: if outcome.is_some() {
                    paragraphs(&planned.removed)
                } else {
                    Vec::new()
                },
                revision_ids: outcome
                    .as_ref()
                    .map_or_else(Vec::new, |outcome| outcome.revision_ids.clone()),
                control: planned.control.clone(),
            }
        })
        .collect()
}

fn text_range(story: &str, first: &str, start: u32, last: &str, end: u32) -> TextRange {
    TextRange {
        story: story.to_owned(),
        start: crate::TextPosition {
            para_id: first.to_owned(),
            offset: start,
        },
        end: crate::TextPosition {
            para_id: last.to_owned(),
            offset: end,
        },
        view: EditTextView::Accepted,
    }
}

impl EditingDoc {
    fn plan(
        &self,
        request: &EditRequest,
        capture: bool,
    ) -> Result<(Plan, Option<Base>), EditRefusal> {
        let nonce = self.version_nonce.load(Ordering::Relaxed);
        let epoch = self.epoch.load(Ordering::Relaxed);
        let version = version_token(nonce, epoch);
        check_budget(request).map_err(|failure| refusal(version.clone(), failure))?;
        if request.expect_version != version {
            return Err(refusal(
                version,
                failure(
                    EditFailureCode::StaleVersion,
                    "the document changed since the expected version was read".to_owned(),
                    None,
                ),
            ));
        }
        let txn = self.yrs_doc().transact();
        if txn.store().pending_update().is_some() || txn.store().pending_ds().is_some() {
            return Err(refusal(
                version,
                failure(
                    EditFailureCode::Unsupported,
                    "the document holds updates that are not integrated yet".to_owned(),
                    None,
                ),
            ));
        }
        let source = self.source_metadata();
        let mut views = Views::new(self, &txn);
        let mut steps = Vec::with_capacity(request.steps.len());
        for (index, step) in request.steps.iter().enumerate() {
            let index = index as u32;
            let planned = plan_step(&mut views, source.as_deref(), index, step)
                .map_err(|failure| refusal(version.clone(), failure.at(index)))?;
            for planned in planned {
                if let Some(earlier) = steps
                    .iter()
                    .find(|earlier: &&Planned| earlier.conflicts(&planned))
                {
                    let mut conflict = failure(
                        EditFailureCode::OverlappingSteps,
                        format!("step {index} conflicts with step {}", earlier.index),
                        Some(planned.target.clone()),
                    )
                    .at(index);
                    conflict.conflicting_step_index = Some(earlier.index);
                    return Err(refusal(version, conflict));
                }
                steps.push(planned);
            }
        }
        let inserted: usize = steps
            .iter()
            .map(|planned| planned.new_paragraph_count as usize)
            .sum();
        if inserted > MAX_INSERTED_PARAGRAPHS {
            return Err(refusal(
                version,
                failure(
                    EditFailureCode::LimitExceeded,
                    format!("a batch inserts at most {MAX_INSERTED_PARAGRAPHS} paragraphs"),
                    None,
                ),
            ));
        }
        let base = if capture && steps.iter().any(|planned| planned.effect.is_some()) {
            let update = deterministic::encode_state_as_update_v1(&txn, &StateVector::default());
            if update.len() > MAX_STAGING_BYTES {
                return Err(refusal(
                    version,
                    failure(
                        EditFailureCode::LimitExceeded,
                        format!("batches stage documents of at most {MAX_STAGING_BYTES} bytes"),
                        None,
                    ),
                ));
            }
            Some(Base {
                update,
                state_vector: txn.state_vector(),
            })
        } else {
            None
        };
        Ok((
            Plan {
                base_version: version,
                nonce,
                epoch,
                source: request.source,
                steps,
            },
            base,
        ))
    }

    /// Resolves, stages and rehearses a batch like [`EditingDoc::apply_edits`], then discards it:
    /// nothing changes and no ids are reserved. `Err` means an internal failure.
    pub fn validate_edits(
        &self,
        request: &EditRequest,
    ) -> EditResult<Result<EditValidation, EditRefusal>> {
        let (plan, base) = match self.plan(request, true) {
            Ok(planned) => planned,
            Err(refusal) => return Ok(Err(refusal)),
        };
        if let Some(base) = base
            && let Err(refusal) = self.stage(&plan, base)?
        {
            return Ok(Err(refusal));
        }
        Ok(Ok(EditValidation {
            base_version: plan.base_version,
            would_apply: plan.steps.iter().any(|planned| planned.effect.is_some()),
            previews: plan
                .steps
                .iter()
                .filter(|planned| !planned.companion)
                .map(|planned| EditPreview {
                    step_index: planned.index,
                    target: planned.target.clone(),
                    would_change: planned.effect.is_some(),
                    new_paragraph_count: planned.new_paragraph_count,
                    would_create_revisions: planned.suggest && planned.effect.is_some(),
                    control: planned.control.clone(),
                })
                .collect(),
        }))
    }

    /// Applies every step or none. Policy failures come back as an [`EditRefusal`] with the
    /// document, history and id allocation untouched; `Err` means an internal failure.
    ///
    /// `history` must be this document's undo session. An applied batch is one committed
    /// transaction: one undo step for [`EditHistory::Separate`], none for [`EditHistory::None`].
    pub fn apply_edits(
        &self,
        request: &EditRequest,
        history: &UndoSession,
    ) -> EditResult<Result<EditApplication, EditRefusal>> {
        if !history.belongs_to(self) {
            return Err(EditError::InvalidUpdate(
                "the undo history belongs to another document".to_owned(),
            ));
        }
        let (plan, base) = match self.plan(request, true) {
            Ok(planned) => planned,
            Err(refusal) => return Ok(Err(refusal)),
        };
        let Some(base) = base else {
            let txn = self.yrs_doc().transact();
            let mut views = Views::new(self, &txn);
            let executed: Vec<Option<Executed>> = plan.steps.iter().map(|_| None).collect();
            return Ok(Ok(EditApplication {
                version: plan.base_version.clone(),
                base_version: plan.base_version,
                applied: false,
                source: plan.source,
                changed_stories: Vec::new(),
                receipts: receipts(&mut views, &plan.steps, &executed),
            }));
        };
        let staged = match self.stage(&plan, base)? {
            Ok(staged) => staged,
            Err(refusal) => return Ok(Err(refusal)),
        };
        let adoption = Update::decode_v1(&staged.update)
            .map_err(|error| EditError::InvalidUpdate(error.to_string()))?;
        let receipts = {
            let txn = staged.stage.yrs_doc().transact();
            let mut views = Views::new(&staged.stage, &txn);
            receipts(&mut views, &plan.steps, &staged.executed)
        };
        if let Err(refusal) = self.check_commit(&plan) {
            return Ok(Err(refusal));
        }
        match request.history {
            EditHistory::Separate => history.track(self),
            EditHistory::None => {}
        }
        history.add_undo_barrier();
        let valued = {
            let mut txn = match request.history {
                EditHistory::Separate => self.yrs_doc().transact_mut_with(self.client_id),
                EditHistory::None => self.yrs_doc().transact_mut_with(HOST_ORIGIN),
            };
            if self.epoch.load(Ordering::Relaxed) != plan.epoch
                || self.version_nonce.load(Ordering::Relaxed) != plan.nonce
            {
                drop(txn);
                return Ok(Err(stale(self.version())));
            }
            let valued = valued_embeds(&txn, &plan.steps)?;
            txn.apply_update(adoption)
                .map_err(|error| EditError::InvalidUpdate(error.to_string()))?;
            valued
        };
        history.add_undo_barrier();
        self.drop_values(&valued);
        self.id_counter.store(
            staged.stage.id_counter.load(Ordering::Relaxed),
            Ordering::Relaxed,
        );
        Ok(Ok(EditApplication {
            base_version: plan.base_version,
            version: self.version(),
            applied: true,
            source: plan.source,
            changed_stories: staged.changed_stories,
            receipts,
        }))
    }

    /// Executes a plan on a private clone that shares this replica's client id and id counter,
    /// validates the staged stories, and rehearses the resulting update against the base.
    fn stage(&self, plan: &Plan, base: Base) -> EditResult<Result<Staged, EditRefusal>> {
        let stage = EditingDoc::new(self.client_id);
        stage.apply_update_v1(&base.update)?;
        stage
            .id_counter
            .store(self.id_counter.load(Ordering::Relaxed), Ordering::Relaxed);
        let executed = execute(&stage, &plan.steps)?;
        let changed_stories: Vec<String> = plan
            .steps
            .iter()
            .zip(&executed)
            .filter(|(_, outcome)| outcome.is_some())
            .flat_map(|(planned, _)| {
                std::iter::once(planned.story.clone()).chain(planned.also_changes.clone())
            })
            .collect::<BTreeSet<_>>()
            .into_iter()
            .collect();
        if let Err(failure) = validate_stage(&stage, &executed, &changed_stories) {
            return Ok(Err(refusal(plan.base_version.clone(), failure)));
        }
        let update = deterministic::encode_diff_v1(&stage.yrs_doc().transact(), &base.state_vector);
        rehearse(
            self.client_id,
            &base.update,
            &update,
            &stage,
            &changed_stories,
        )?;
        Ok(Ok(Staged {
            stage,
            executed,
            changed_stories,
            update,
        }))
    }

    /// Drops the authored values of the text controls a committed batch filled, in a transaction
    /// of its own so undoing the fill never brings them back.
    fn drop_values(&self, valued: &[yrs::MapRef]) {
        if valued.is_empty() {
            return;
        }
        let mut txn = self.yrs_doc().transact_mut_with(HOST_ORIGIN);
        for map in valued {
            if !AsRef::<yrs::branch::Branch>::as_ref(map).is_deleted() {
                yrs::Map::remove(map, &mut txn, "value");
            }
        }
    }

    fn check_commit(&self, plan: &Plan) -> Result<(), EditRefusal> {
        let txn = self.yrs_doc().transact();
        let current = self.version();
        if self.version_nonce.load(Ordering::Relaxed) != plan.nonce
            || self.epoch.load(Ordering::Relaxed) != plan.epoch
        {
            return Err(stale(current));
        }
        if txn.store().pending_update().is_some() || txn.store().pending_ds().is_some() {
            return Err(refusal(
                current,
                failure(
                    EditFailureCode::Unsupported,
                    "the document holds updates that are not integrated yet".to_owned(),
                    None,
                ),
            ));
        }
        Ok(())
    }
}

/// The embeds of the text controls `steps` fill that carry an authored value, read from the state
/// they were planned on.
fn valued_embeds<T: ReadTxn>(txn: &T, steps: &[Planned]) -> EditResult<Vec<yrs::MapRef>> {
    steps
        .iter()
        .filter_map(|planned| match &planned.effect {
            Some(Effect::Control(ControlFill::Inline {
                raw,
                drop_value: true,
                ..
            })) => Some((planned.story.as_str(), *raw)),
            Some(Effect::Control(ControlFill::Block {
                parent,
                raw,
                drop_value: true,
                ..
            })) => Some((parent.as_str(), *raw)),
            _ => None,
        })
        .map(|(story, raw)| {
            let text = story_ref(txn, story)?;
            crate::ops::embed::embed_map_at(&text, txn, raw)
                .map_err(|error| EditError::InvalidUpdate(error.to_string()))
        })
        .collect()
}

fn stale(version: DocumentVersion) -> EditRefusal {
    refusal(
        version,
        failure(
            EditFailureCode::StaleVersion,
            "the document changed while the batch was staged".to_owned(),
            None,
        ),
    )
}

/// Checks the staged stories are well formed and the minted paragraph ids are unique.
fn validate_stage(
    stage: &EditingDoc,
    executed: &[Option<Executed>],
    changed_stories: &[String],
) -> Result<(), EditFailure> {
    let txn = stage.yrs_doc().transact();
    let malformed = |message: String| failure(EditFailureCode::Unsupported, message, None);
    let mut ids: Vec<String> = Vec::new();
    for story_id in changed_stories {
        let Ok(story) = story_ref(&txn, story_id) else {
            return Err(malformed(format!(
                "story {story_id:?} vanished while staging"
            )));
        };
        let pilcrows = crate::pilcrows(&story, &txn);
        let len = yrs::Text::len(&story, &txn);
        if pilcrows.last().is_none_or(|(index, _)| index + 1 != len) {
            return Err(malformed(format!(
                "story {story_id:?} would no longer end in a paragraph mark"
            )));
        }
        ids.extend(
            pilcrows
                .iter()
                .filter_map(|(_, map)| crate::map_string(map, &txn, crate::PARA_ID)),
        );
    }
    let minted: Vec<&String> = executed
        .iter()
        .flatten()
        .flat_map(|outcome| &outcome.new_ids)
        .collect();
    if minted
        .iter()
        .any(|id| ids.iter().filter(|existing| existing == id).count() > 1)
    {
        return Err(malformed(
            "a minted paragraph id collides with an existing paragraph".to_owned(),
        ));
    }
    Ok(())
}

/// Integrates the staged update into an untouched copy of the base and checks it reproduces the
/// staged stories; adoption then cannot meet a failure the rehearsal did not.
fn rehearse(
    client_id: u64,
    base: &[u8],
    update: &[u8],
    stage: &EditingDoc,
    changed_stories: &[String],
) -> EditResult<()> {
    let rehearsal = EditingDoc::new(client_id);
    rehearsal.apply_update_v1(base)?;
    rehearsal.apply_update_v1(update)?;
    let pending = {
        let txn = rehearsal.yrs_doc().transact();
        txn.store().pending_update().is_some() || txn.store().pending_ds().is_some()
    };
    let diverged = pending
        || rehearsal.encode_state_vector_v1() != stage.encode_state_vector_v1()
        || changed_stories
            .iter()
            .any(|story| rehearsal.story_segments(story).ok() != stage.story_segments(story).ok());
    if diverged {
        return Err(EditError::InvalidUpdate(
            "the rehearsed adoption diverged from the staged batch".to_owned(),
        ));
    }
    Ok(())
}
