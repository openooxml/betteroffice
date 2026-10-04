use std::borrow::Cow;
use std::collections::{BTreeMap, BTreeSet, HashMap, HashSet};
use std::fmt;
use std::sync::{Arc, Mutex};

use ooxml_opc::PackageBytes;
use serde::Deserialize;
use serde::de::{MapAccess, SeqAccess, Visitor};
use serde_json::{Map, Value, json};
use yrs::types::Attrs;
use yrs::{Any, Array as _, Map as YrsMap, Out, ReadTxn, Text as _, Transact};

use crate::control_source::safety_key;
use crate::identity::{
    OOXML_PARA_ID, PARA_ORIGIN, SOURCE_PARA_ID, SYNTHETIC, SeededParagraph, SourceIndex,
    SourcePackage, SourcePartInput, SourceStoryKind,
};
use crate::script_fonts::ScriptFontUse;
use crate::structured::source::{
    CellLayout, CommentWrites, InlineRecord, InlineSource, Pin, Provenance, RawSource, ReadSource,
    Relocated, Represented, RowLayout, SourceMerge, SourceParts, Step, TableLayout, Witness,
    story_root,
};
use crate::structured::{BreakType, Revision, RevisionKind, StoryKind};
use crate::{EditCtx, EditingDoc, RawOp};

type JsonObject = BTreeMap<String, Value>;

/// Marks state seeded with sequence metadata. State written before sequence numbering
/// lacks this key and keeps cached SEQ results.
pub(crate) const OPAQUE_SEQUENCES: &str = "opaqueSequences";

/// Writes the first marker, so a seed keeps the operation ids an earlier seed of the same
/// document gave its stories and session: replicas that seeded it before still converge.
const SEQUENCE_METADATA_CLIENT: u64 = 0x1_0000_05e9;

pub(crate) fn seed_opaque_sequences(document: &EditingDoc, names: &[String], seeded: Option<bool>) {
    let mut txn = document.transact_for(&EditCtx::system(""));
    let session = txn
        .get_map(crate::identity::SESSION)
        .expect("session root is declared by EditingDoc::new");
    let previous = session.get(&txn, OPAQUE_SEQUENCES);
    if let Some(seeded) = seeded {
        debug_assert_eq!(seeded, holds_sequence_fields(&txn));
    }
    if previous.is_none()
        && names.is_empty()
        && !seeded.unwrap_or_else(|| holds_sequence_fields(&txn))
    {
        return;
    }
    let mut opaque_sequences: BTreeSet<String> = names.iter().cloned().collect();
    if let Some(Out::Any(Any::Array(previous))) = &previous {
        opaque_sequences.extend(previous.iter().filter_map(|value| match value {
            Any::String(name) => Some(name.to_string()),
            _ => None,
        }));
    }
    let names = Any::Array(
        opaque_sequences
            .into_iter()
            .map(Any::from)
            .collect::<Vec<_>>()
            .into(),
    );
    if previous.is_some() {
        if previous != Some(Out::Any(names.clone())) {
            session.insert(&mut txn, OPAQUE_SEQUENCES, names);
        }
        return;
    }
    drop(txn);
    let marker = yrs::Doc::with_client_id(SEQUENCE_METADATA_CLIENT);
    let root = marker.get_or_insert_map(crate::identity::SESSION);
    let mut marker_txn = marker.transact_mut();
    root.insert(&mut marker_txn, OPAQUE_SEQUENCES, names);
    let update = marker_txn.encode_update_v1();
    drop(marker_txn);
    document
        .apply_verbatim_v1(&update)
        .expect("a one-entry session update applies");
}

/// Whether a story embeds anything that names a SEQ field (a field instruction, nested
/// sequence names, shape text). A document without one seeds exactly as before.
fn holds_sequence_fields<T: ReadTxn>(txn: &T) -> bool {
    let Some(stories) = txn.get_map(crate::STORIES) else {
        return false;
    };
    let mut pending: Vec<Out> = stories.iter(txn).map(|(_, value)| value).collect();
    let mut anys: Vec<Any> = Vec::new();
    while let Some(value) = pending.pop() {
        match value {
            Out::YText(text) => pending.extend(
                text.diff(txn, yrs::types::text::YChange::identity)
                    .into_iter()
                    .map(|chunk| chunk.insert)
                    .filter(|insert| !matches!(insert, Out::Any(Any::String(_)))),
            ),
            Out::YMap(map) => pending.extend(map.iter(txn).map(|(_, value)| value)),
            Out::YArray(array) => pending.extend(array.iter(txn)),
            Out::Any(any) => anys.push(any),
            _ => {}
        }
    }
    for any in &anys {
        let mut values = vec![any];
        while let Some(value) = values.pop() {
            match value {
                Any::String(text) if text.contains("SEQ") => return true,
                Any::Array(items) => values.extend(items.iter()),
                Any::Map(entries) => values.extend(entries.values()),
                _ => {}
            }
        }
    }
    false
}

/// Whether the embed values a fresh seed writes hold a string with SEQ, as
/// [`holds_sequence_fields`] reads them back; `None` for ops a seed never writes.
fn seeded_sequence_fields<'a>(ops: impl IntoIterator<Item = &'a RawOp>) -> Option<bool> {
    let mut found = false;
    for op in ops {
        match op {
            RawOp::InsertEmbed { kind, payload, .. } if !found => {
                let pilcrow = kind == crate::PILCROW_KIND;
                let mut values: Vec<&Any> = payload
                    .iter()
                    .filter(|(key, _)| {
                        !pilcrow
                            || !matches!(key.as_str(), OOXML_PARA_ID | SOURCE_PARA_ID | PARA_ORIGIN)
                    })
                    .map(|(_, value)| value)
                    .collect();
                found =
                    !payload.iter().any(|(key, _)| key == crate::KIND_KEY) && kind.contains("SEQ");
                while !found && let Some(value) = values.pop() {
                    match value {
                        Any::String(text) => found = text.contains("SEQ"),
                        Any::Array(items) => values.extend(items.iter()),
                        Any::Map(entries) => values.extend(entries.values()),
                        _ => {}
                    }
                }
            }
            RawOp::Insert { .. } | RawOp::Format { .. } | RawOp::InsertEmbed { .. } => {}
            RawOp::SetComment { .. } | RawOp::RemoveComment { .. } => {}
            RawOp::Delete { .. } | RawOp::SetEmbedAttr { .. } => return None,
        }
    }
    Some(found)
}

/// A parsed node's `w:p` occurrence in its part: read for identity, never seeded.
const SOURCE_ORDINAL: &str = "sourceOrdinal";

#[derive(Clone)]
struct Mark {
    name: String,
    attrs: Vec<(String, Value)>,
}

#[derive(Clone)]
enum UnitContent {
    Text(String),
    Embed { kind: String, payload: JsonObject },
}

#[derive(Clone)]
struct InlineUnit {
    content: UnitContent,
    attrs: JsonObject,
    pm_size: u32,
    marks: Vec<Mark>,
}

#[derive(Debug, PartialEq, Eq)]
struct CommentMark {
    unit: usize,
    start: bool,
    id: String,
}

struct StoryPlan {
    story_id: String,
    units: Vec<InlineUnit>,
    comment_marks: Vec<CommentMark>,
    comment_coverage: Vec<(String, Vec<(u32, u32)>)>,
    /// How many units [`StoryPlan::width`] has measured, and their width.
    measured: (usize, u32),
}

impl StoryPlan {
    /// The story index the next unit gets.
    fn width(&mut self) -> u32 {
        let (count, width) = self.measured;
        let width = width + self.units[count..].iter().map(unit_width).sum::<u32>();
        self.measured = (self.units.len(), width);
        width
    }
}

struct ProjectedCell<'a> {
    paragraph_formatting: Option<Value>,
    attrs: JsonObject,
    content: Cow<'a, [Value]>,
}

struct ProjectedRow<'a> {
    attrs: JsonObject,
    cells: Vec<ProjectedCell<'a>>,
}

struct ProjectedTable<'a> {
    attrs: JsonObject,
    rows: Vec<ProjectedRow<'a>>,
}

#[derive(Clone, Copy)]
struct StoryOptions {
    include_page_breaks: bool,
    append_body_tail: bool,
    seed_comments: bool,
}

struct LoweringContext {
    styles: StyleResolver,
    theme: Option<Value>,
    source_json: Arc<BTreeMap<String, String>>,
    plans: Vec<StoryPlan>,
    compatibility_mode: u8,
    /// The root story being lowered, for [`Self::paragraphs`].
    root: String,
    /// Every lowered paragraph in document order, nested stories in place.
    paragraphs: Vec<SeededParagraph>,
    opaque_sequences: Vec<String>,
    source: SourceStructure,
    provenance: Provenance,
    /// Each source story's steps from its part's root element.
    locators: HashMap<String, Vec<Step>>,
}

/// A story's source blocks as the save projection sees them when it puts raw XML back: each raw
/// block goes in front of the first following source paragraph that still exists.
enum SourceBlock {
    Raw,
    /// A paragraph with a source id, which a raw block can be restored in front of.
    Anchor(String),
    /// A table, control or id-less paragraph, which that search skips over.
    Other,
}

/// Where the save projection puts a raw XML block back.
pub(crate) enum Restoration {
    /// In front of this live paragraph, where the source had it.
    Before(String),
    /// In front of a later paragraph, past blocks the source had after it.
    Displaced,
    /// By position, as no source paragraph follows it any more.
    Unanchored,
}

/// Source structure the story stream does not carry, retained for batch planning.
#[derive(Default)]
struct SourceStructure {
    /// Source block order of the stories that hold raw XML blocks.
    blocks: HashMap<String, Vec<SourceBlock>>,
    /// Paragraphs whose runs carry tracked formatting changes.
    run_revisions: HashSet<(String, String)>,
}

/// Package context retained after lowering, for planning host edits and exporting in Rust.
pub(crate) struct SourceMetadata {
    styles: StyleResolver,
    structure: SourceStructure,
    read: ReadSource,
}

/// A paragraph's style-derived pilcrow properties and run formatting.
pub(crate) struct StyledParagraph {
    pub properties: Vec<(String, Any)>,
    pub run: Vec<(String, Any)>,
}

/// Pilcrow keys seeded from direct formatting with a paragraph-style fallback.
const STYLE_FALLBACK_KEYS: [&str; 23] = [
    "alignment",
    "spaceBefore",
    "spaceAfter",
    "spaceBeforeLines",
    "spaceAfterLines",
    "beforeAutospacing",
    "afterAutospacing",
    "lineSpacing",
    "lineSpacingRule",
    "indentRight",
    "borders",
    "shading",
    "tabs",
    "pageBreakBefore",
    "keepNext",
    "keepLines",
    "widowControl",
    "contextualSpacing",
    "snapToGrid",
    "autoSpaceDE",
    "autoSpaceDN",
    "outlineLevel",
    "bidi",
];

/// Every pilcrow key a seeded paragraph resolves through its style.
pub(crate) fn style_resolved_keys() -> impl Iterator<Item = &'static str> {
    STYLE_FALLBACK_KEYS.into_iter().chain([
        "spacingExplicit",
        "indentLeft",
        "indentFirstLine",
        "hangingIndent",
        "defaultTextFormatting",
    ])
}

impl SourceMetadata {
    pub(crate) fn has_paragraph_style(&self, style_id: &str) -> bool {
        self.styles
            .style(style_id)
            .is_some_and(|style| string(field(Some(style), "type")) == Some("paragraph"))
    }

    /// What seeding produces for an unformatted paragraph carrying only `style_id`.
    pub(crate) fn styled_paragraph(
        &self,
        style_id: Option<&str>,
    ) -> Result<StyledParagraph, String> {
        let formatting = style_id.map_or_else(|| json!({}), |id| json!({ "styleId": id }));
        let paragraph = json!({ "type": "paragraph", "formatting": formatting, "content": [] });
        let attrs = paragraph_attrs(&paragraph, &self.styles, &[], &[], None);
        let run = marks_to_attrs(&formatting_to_marks(
            paragraph_style_formatting(&paragraph, &self.styles, None).as_deref(),
        ));
        Ok(StyledParagraph {
            properties: payload(para_attrs_to_ppr(attrs))?,
            run: payload(run)?,
        })
    }

    /// The run formatting typed text takes in a paragraph of `style_id` inside a content control
    /// whose own run properties are `control` (a parsed `w:rPr`).
    pub(crate) fn control_run(
        &self,
        style_id: Option<&str>,
        control: Option<&Value>,
    ) -> Result<Vec<(String, Any)>, String> {
        let formatting = style_id.map_or_else(|| json!({}), |id| json!({ "styleId": id }));
        let paragraph = json!({ "type": "paragraph", "formatting": formatting, "content": [] });
        let style = paragraph_style_formatting(&paragraph, &self.styles, None);
        let run_style = self.styles.run_style_own(string(field(control, "styleId")));
        let inherited = merge_text_formatting(style.as_deref(), run_style);
        let merged = merge_text_formatting(inherited.as_ref(), control);
        payload(marks_to_attrs(&formatting_to_marks(merged.as_ref())))
    }

    /// How the save projection restores each raw XML block of `story` when only the paragraphs
    /// `alive` accepts remain.
    pub(crate) fn opaque_restorations(
        &self,
        story: &str,
        alive: impl Fn(&str) -> bool,
    ) -> Vec<Restoration> {
        let Some(blocks) = self.structure.blocks.get(story) else {
            return Vec::new();
        };
        blocks
            .iter()
            .enumerate()
            .filter(|(_, block)| matches!(block, SourceBlock::Raw))
            .map(|(index, _)| {
                let mut skipped = false;
                for block in &blocks[index + 1..] {
                    match block {
                        SourceBlock::Anchor(id) if alive(id) => {
                            return if skipped {
                                Restoration::Displaced
                            } else {
                                Restoration::Before(id.clone())
                            };
                        }
                        SourceBlock::Other => skipped = true,
                        _ => {}
                    }
                }
                Restoration::Unanchored
            })
            .collect()
    }

    pub(crate) fn read(&self) -> &ReadSource {
        &self.read
    }

    /// The document's numbering definitions.
    pub(crate) fn numbering(&self) -> Arc<docx_parse::NumberingMap> {
        Arc::clone(&self.read.numbering)
    }

    /// Records the comment writes committed to `doc` from now on.
    #[cfg(feature = "wasm")]
    pub(crate) fn watch_comments(&mut self, doc: &EditingDoc) {
        self.read.comment_writes = CommentWrites::watch(doc);
    }

    pub(crate) fn has_style(&self, style_id: &str) -> bool {
        self.styles.style(style_id).is_some()
    }

    /// The paragraph style that applies when a paragraph names none or an undefined one.
    pub(crate) fn default_paragraph_style(&self) -> Option<&str> {
        self.styles.default_paragraph.as_deref()
    }

    /// The outline level `style_id` defines, its `basedOn` chain included.
    pub(crate) fn style_outline_level(&self, style_id: &str) -> Option<f64> {
        number(field(
            field(self.styles.style(style_id), "pPr"),
            "outlineLevel",
        ))
    }

    /// Whether paragraph styles `left` and `right` give runs the same bold, italic, underline,
    /// strike, vertical alignment and hidden state.
    pub(crate) fn same_run_marks(&self, left: Option<&str>, right: Option<&str>) -> bool {
        let marks = |style_id: Option<&str>| {
            let style = self.styles.resolve_paragraph_style(style_id);
            [
                "bold",
                "italic",
                "underline",
                "strike",
                "vertAlign",
                "hidden",
            ]
            .map(|key| {
                field(style.run.as_deref(), key)
                    .cloned()
                    .unwrap_or(Value::Null)
            })
        };
        marks(left) == marks(right)
    }

    /// The outline level of the document's default paragraph properties.
    pub(crate) fn default_outline_level(&self) -> Option<f64> {
        number(field(
            field(self.styles.doc_defaults.as_ref(), "pPr"),
            "outlineLevel",
        ))
    }

    pub(crate) fn run_revision(&self, story: &str, para_id: &str) -> bool {
        self.structure
            .run_revisions
            .contains(&(story.to_owned(), para_id.to_owned()))
    }

    /// The stories holding a paragraph whose source runs carry tracked formatting changes.
    pub(crate) fn run_revision_stories(&self) -> impl Iterator<Item = &str> {
        self.structure
            .run_revisions
            .iter()
            .map(|(story, _)| story.as_str())
    }
}

fn has_run_property_changes(value: &Value) -> bool {
    match value {
        Value::Object(object) => {
            (string(object.get("type")) == Some("run")
                && !array(object.get("propertyChanges")).is_empty())
                || object.values().any(has_run_property_changes)
        }
        Value::Array(values) => values.iter().any(has_run_property_changes),
        _ => false,
    }
}

fn compatibility_mode_from_package(package: Option<&Value>) -> u8 {
    field(field(package, "settings"), "compatibilityFlags")
        .and_then(|flags| field(Some(flags), "compatibilityMode"))
        .and_then(|value| value.as_f64())
        .filter(|value| value.is_finite() && (0.0..=255.0).contains(value))
        .map(|value| value as u8)
        .unwrap_or(12)
}

enum OrderedValue {
    Null,
    Bool(bool),
    Number(serde_json::Number),
    String(String),
    Array(Vec<OrderedValue>),
    Object(Vec<(String, OrderedValue)>),
}

struct OrderedValueVisitor;

impl<'de> Visitor<'de> for OrderedValueVisitor {
    type Value = OrderedValue;

    fn expecting(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str("a JSON value")
    }

    fn visit_unit<E>(self) -> Result<Self::Value, E> {
        Ok(OrderedValue::Null)
    }

    fn visit_none<E>(self) -> Result<Self::Value, E> {
        Ok(OrderedValue::Null)
    }

    fn visit_bool<E>(self, value: bool) -> Result<Self::Value, E> {
        Ok(OrderedValue::Bool(value))
    }

    fn visit_i64<E>(self, value: i64) -> Result<Self::Value, E> {
        Ok(OrderedValue::Number(value.into()))
    }

    fn visit_u64<E>(self, value: u64) -> Result<Self::Value, E> {
        Ok(OrderedValue::Number(value.into()))
    }

    fn visit_f64<E>(self, value: f64) -> Result<Self::Value, E>
    where
        E: serde::de::Error,
    {
        serde_json::Number::from_f64(value)
            .map(OrderedValue::Number)
            .ok_or_else(|| E::custom("non-finite JSON number"))
    }

    fn visit_str<E>(self, value: &str) -> Result<Self::Value, E> {
        Ok(OrderedValue::String(value.to_owned()))
    }

    fn visit_string<E>(self, value: String) -> Result<Self::Value, E> {
        Ok(OrderedValue::String(value))
    }

    fn visit_seq<A>(self, mut sequence: A) -> Result<Self::Value, A::Error>
    where
        A: SeqAccess<'de>,
    {
        let mut values = Vec::new();
        while let Some(value) = sequence.next_element()? {
            values.push(value);
        }
        Ok(OrderedValue::Array(values))
    }

    fn visit_map<A>(self, mut object: A) -> Result<Self::Value, A::Error>
    where
        A: MapAccess<'de>,
    {
        let mut entries = Vec::new();
        while let Some(entry) = object.next_entry::<String, OrderedValue>()? {
            if entry.0 != SOURCE_ORDINAL {
                entries.push(entry);
            }
        }
        Ok(OrderedValue::Object(entries))
    }
}

impl<'de> Deserialize<'de> for OrderedValue {
    fn deserialize<D>(deserializer: D) -> Result<Self, D::Error>
    where
        D: serde::Deserializer<'de>,
    {
        deserializer.deserialize_any(OrderedValueVisitor)
    }
}

#[derive(Clone, Default)]
struct StyleResolver {
    enabled: bool,
    styles: BTreeMap<String, Value>,
    doc_defaults: Option<Value>,
    default_paragraph: Option<String>,
    default_table: Option<String>,
    default_character: Option<String>,
    table_paragraph_formatting: Option<Value>,
    memo: StyleMemo,
}

struct ResolvedParagraphStyle {
    paragraph: Option<Value>,
    run: Option<Arc<Value>>,
    default_run: Option<Arc<Value>>,
}

/// Maximum approximate retained value size across both style memos.
const STYLE_MEMO_BYTES: usize = 4 << 20;

/// Approximate fixed cost charged for every memo entry.
const STYLE_MEMO_ENTRY_BYTES: usize = 64;

/// Paragraph styles grouped by exact lookup result.
#[derive(Default)]
struct ParagraphStyleMemo {
    absent: Option<Arc<ResolvedParagraphStyle>>,
    undefined: Option<Arc<ResolvedParagraphStyle>>,
    styles: HashMap<String, Arc<ResolvedParagraphStyle>>,
}

/// Run styles grouped by exact lookup result.
#[derive(Default)]
struct RunStyleMemo {
    unstyled: Option<Option<Arc<Value>>>,
    styles: HashMap<String, Option<Arc<Value>>>,
}

/// Style memos sharing one retained value budget.
#[derive(Default)]
struct MemoState {
    paragraphs: ParagraphStyleMemo,
    runs: RunStyleMemo,
    bytes: usize,
}

#[derive(Default)]
struct StyleMemo {
    state: Mutex<MemoState>,
}

impl Clone for StyleMemo {
    fn clone(&self) -> Self {
        Self::default()
    }
}

/// Estimates JSON storage from nodes, strings, and object keys.
fn approx_bytes(value: &Value) -> usize {
    32 + match value {
        Value::String(value) => value.len(),
        Value::Array(values) => values.iter().map(approx_bytes).sum(),
        Value::Object(values) => values
            .iter()
            .map(|(key, value)| key.len() + approx_bytes(value))
            .sum(),
        _ => 0,
    }
}

fn object(value: Option<&Value>) -> Option<&Map<String, Value>> {
    value?.as_object()
}

fn array(value: Option<&Value>) -> &[Value] {
    value
        .and_then(Value::as_array)
        .map(Vec::as_slice)
        .unwrap_or(&[])
}

fn identical_json(left: &Value, right: &Value) -> bool {
    match (left, right) {
        (Value::Number(left), Value::Number(right)) => {
            left == right && left.as_f64().map(f64::to_bits) == right.as_f64().map(f64::to_bits)
        }
        (Value::Array(left), Value::Array(right)) => {
            left.len() == right.len()
                && left
                    .iter()
                    .zip(right)
                    .all(|(left, right)| identical_json(left, right))
        }
        (Value::Object(left), Value::Object(right)) => {
            left.len() == right.len()
                && left
                    .iter()
                    .zip(right)
                    .all(|((left_key, left), (right_key, right))| {
                        left_key == right_key && identical_json(left, right)
                    })
        }
        _ => left == right,
    }
}

fn field<'a>(value: Option<&'a Value>, key: &str) -> Option<&'a Value> {
    object(value)?.get(key)
}

fn string(value: Option<&Value>) -> Option<&str> {
    value?.as_str()
}

fn number(value: Option<&Value>) -> Option<f64> {
    value?.as_f64()
}

fn boolean(value: Option<&Value>) -> Option<bool> {
    value?.as_bool()
}

fn truthy(value: Option<&Value>) -> bool {
    match value {
        None | Some(Value::Null) => false,
        Some(Value::Bool(value)) => *value,
        Some(Value::Number(value)) => value.as_f64().is_some_and(|value| value != 0.0),
        Some(Value::String(value)) => !value.is_empty(),
        Some(Value::Array(_) | Value::Object(_)) => true,
    }
}

fn nullish(value: Option<&Value>) -> Value {
    value.cloned().unwrap_or(Value::Null)
}

fn js_string(value: &Value) -> String {
    match value {
        Value::String(value) => value.clone(),
        Value::Number(value) => {
            if let Some(number) = value.as_f64()
                && number.fract() == 0.0
                && number.abs() <= 9_007_199_254_740_991.0
            {
                return format!("{number:.0}");
            }
            value.to_string()
        }
        Value::Bool(value) => value.to_string(),
        Value::Null => "null".to_owned(),
        _ => serde_json::to_string(value).unwrap_or_default(),
    }
}

fn utf16_len(value: &str) -> u32 {
    value.encode_utf16().count() as u32
}

fn ordered_object(
    entries: impl IntoIterator<Item = (impl Into<String>, Value)>,
) -> Vec<(String, Value)> {
    entries
        .into_iter()
        .map(|(key, value)| (key.into(), value))
        .collect()
}

fn ordered_json(entries: &[(String, Value)]) -> String {
    let mut output = String::from("{");
    for (index, (key, value)) in entries.iter().enumerate() {
        if index > 0 {
            output.push(',');
        }
        output.push_str(&serde_json::to_string(key).unwrap());
        output.push(':');
        js_json_into(value, &mut output);
    }
    output.push('}');
    output
}

fn js_json(value: &Value) -> String {
    let mut out = String::new();
    js_json_into(value, &mut out);
    out
}

/// Appends JSON with JavaScript number formatting.
fn js_json_into(value: &Value, out: &mut String) {
    match value {
        Value::Null => out.push_str("null"),
        Value::Bool(value) => out.push_str(&value.to_string()),
        Value::Number(value) => {
            let value = value.as_f64().unwrap_or_default();
            if value == 0.0 {
                out.push('0');
            } else {
                out.push_str(ryu_js::Buffer::new().format(value));
            }
        }
        Value::String(value) => out.push_str(&serde_json::to_string(value).unwrap()),
        Value::Array(values) => {
            out.push('[');
            for (index, value) in values.iter().enumerate() {
                if index > 0 {
                    out.push(',');
                }
                js_json_into(value, out);
            }
            out.push(']');
        }
        Value::Object(values) => {
            out.push('{');
            for (index, (key, value)) in values.iter().enumerate() {
                if index > 0 {
                    out.push(',');
                }
                out.push_str(&serde_json::to_string(key).unwrap());
                out.push(':');
                js_json_into(value, out);
            }
            out.push('}');
        }
    }
}

impl OrderedValue {
    fn value(&self) -> Value {
        match self {
            Self::Null => Value::Null,
            Self::Bool(value) => Value::Bool(*value),
            Self::Number(value) => Value::Number(value.clone()),
            Self::String(value) => Value::String(value.clone()),
            Self::Array(values) => Value::Array(values.iter().map(Self::value).collect()),
            Self::Object(entries) => Value::Object(
                entries
                    .iter()
                    .map(|(key, value)| (key.clone(), value.value()))
                    .collect(),
            ),
        }
    }

    fn js_json(&self) -> String {
        let mut out = String::new();
        self.js_json_into(&mut out);
        out
    }

    /// Appends JSON in source order with JavaScript number formatting.
    fn js_json_into(&self, out: &mut String) {
        match self {
            Self::Null => out.push_str("null"),
            Self::Bool(value) => out.push_str(&value.to_string()),
            Self::Number(value) => {
                let value = value.as_f64().unwrap_or_default();
                if value == 0.0 {
                    out.push('0');
                } else {
                    out.push_str(ryu_js::Buffer::new().format(value));
                }
            }
            Self::String(value) => out.push_str(&serde_json::to_string(value).unwrap()),
            Self::Array(values) => {
                out.push('[');
                for (index, value) in values.iter().enumerate() {
                    if index > 0 {
                        out.push(',');
                    }
                    value.js_json_into(out);
                }
                out.push(']');
            }
            Self::Object(entries) => {
                out.push('{');
                for (index, (key, value)) in entries.iter().enumerate() {
                    if index > 0 {
                        out.push(',');
                    }
                    out.push_str(&serde_json::to_string(key).unwrap());
                    out.push(':');
                    value.js_json_into(out);
                }
                out.push('}');
            }
        }
    }

    fn insert_source_json(&self, output: &mut BTreeMap<String, String>) {
        output
            .entry(serde_json::to_string(&self.value()).unwrap())
            .or_insert_with(|| self.js_json());
    }

    fn collect_source_json(&self, output: &mut BTreeMap<String, String>) {
        match self {
            Self::Array(values) => {
                for value in values {
                    value.collect_source_json(output);
                }
            }
            Self::Object(entries) => {
                let node_type = entries.iter().find_map(|(key, value)| {
                    (key == "type")
                        .then_some(value)
                        .and_then(|value| match value {
                            Self::String(value) => Some(value.as_str()),
                            _ => None,
                        })
                });
                if matches!(
                    node_type,
                    Some("simpleField" | "complexField" | "shape" | "chart")
                ) {
                    self.insert_source_json(output);
                }
                if matches!(node_type, Some("inlineSdt" | "blockSdt"))
                    && let Some((_, properties)) =
                        entries.iter().find(|(key, _)| key == "properties")
                {
                    properties.insert_source_json(output);
                    if let Self::Object(properties) = properties {
                        for (_, value) in properties
                            .iter()
                            .filter(|(key, _)| matches!(key.as_str(), "listItems" | "dataBinding"))
                        {
                            value.insert_source_json(output);
                        }
                    }
                }
                for (_, value) in entries {
                    value.collect_source_json(output);
                }
            }
            _ => {}
        }
    }
}

fn needs_source_json(value: &Value) -> bool {
    match value {
        Value::Array(values) => values.iter().any(needs_source_json),
        Value::Object(values) => {
            matches!(
                string(values.get("type")),
                Some("simpleField" | "complexField" | "shape" | "chart" | "inlineSdt" | "blockSdt")
            ) || values.values().any(needs_source_json)
        }
        _ => false,
    }
}

fn source_json(value: &Value, values: &BTreeMap<String, String>) -> String {
    let mut value = value.clone();
    strip_source_ordinals(&mut value);
    serde_json::to_string(&value)
        .ok()
        .and_then(|key| values.get(&key).cloned())
        .unwrap_or_else(|| js_json(&value))
}

fn strip_source_ordinals(value: &mut Value) {
    match value {
        Value::Array(values) => values.iter_mut().for_each(strip_source_ordinals),
        Value::Object(values) => {
            values.remove(SOURCE_ORDINAL);
            values.values_mut().for_each(strip_source_ordinals);
        }
        _ => {}
    }
}

/// Removes object nulls recursively while retaining array nulls.
fn drop_nulls_in_place(value: &mut Value) {
    match value {
        Value::Array(values) => values.iter_mut().for_each(drop_nulls_in_place),
        Value::Object(values) => {
            values.retain(|_, value| !value.is_null());
            values.values_mut().for_each(drop_nulls_in_place);
        }
        _ => {}
    }
}

fn drop_nulls(mut value: Value) -> Value {
    drop_nulls_in_place(&mut value);
    value
}

fn map_from_value(value: Value) -> JsonObject {
    match value {
        Value::Object(values) => values
            .into_iter()
            .filter(|(_, value)| !value.is_null())
            .map(|(key, mut value)| {
                drop_nulls_in_place(&mut value);
                (key, value)
            })
            .collect(),
        _ => JsonObject::new(),
    }
}

fn value_from_map(value: &JsonObject) -> Value {
    Value::Object(
        value
            .iter()
            .map(|(key, value)| (key.clone(), value.clone()))
            .collect(),
    )
}

fn any_from_value(value: Value) -> Result<Any, String> {
    match value {
        Value::Null => Ok(Any::Null),
        Value::Bool(value) => Ok(Any::Bool(value)),
        Value::Number(value) if value.is_i64() => Ok(Any::from(value.as_i64().unwrap())),
        Value::Number(value) if value.is_u64() => Any::try_from(value.as_u64().unwrap())
            .map_err(|value| format!("JSON number {value} exceeds the yrs integer range")),
        Value::Number(value) => Ok(Any::Number(
            value
                .as_f64()
                .ok_or_else(|| format!("invalid JSON number {value}"))?,
        )),
        Value::String(value) => Ok(Any::String(Arc::from(value))),
        Value::Array(values) => values
            .into_iter()
            .map(any_from_value)
            .collect::<Result<Vec<_>, _>>()
            .map(Arc::from)
            .map(Any::Array),
        Value::Object(values) => {
            let mut entries = values
                .into_iter()
                .map(|(key, value)| Ok((key, any_from_value(value)?)))
                .collect::<Result<Vec<_>, String>>()?;
            entries.sort_unstable_by(|left, right| left.0.cmp(&right.0));
            Ok(Any::Map(Arc::new(entries.into_iter().collect())))
        }
    }
}

fn yrs_attrs(values: JsonObject) -> Result<Attrs, String> {
    let mut entries = values
        .into_iter()
        .map(|(key, value)| Ok((Arc::<str>::from(key), any_from_value(value)?)))
        .collect::<Result<Vec<_>, String>>()?;
    entries.sort_unstable_by(|left, right| left.0.cmp(&right.0));
    Ok(entries.into_iter().collect())
}

fn payload(values: JsonObject) -> Result<Vec<(String, Any)>, String> {
    values
        .into_iter()
        .map(|(key, value)| Ok((key, any_from_value(value)?)))
        .collect()
}

fn merge_plain(target: Option<&Value>, source: Option<&Value>) -> Option<Value> {
    match (object(target), object(source)) {
        (None, None) => None,
        (Some(target), None) => Some(Value::Object(target.clone())),
        (None, Some(source)) => Some(Value::Object(source.clone())),
        (Some(target), Some(source)) => {
            let mut result = target.clone();
            for (key, value) in source {
                result.insert(key.clone(), value.clone());
            }
            Some(Value::Object(result))
        }
    }
}

fn merge_font_family(target: Option<&Value>, source: &Value) -> Value {
    let mut result = object(target).cloned().unwrap_or_default();
    let source = source.as_object().cloned().unwrap_or_default();
    for (explicit, theme) in [
        ("ascii", "asciiTheme"),
        ("hAnsi", "hAnsiTheme"),
        ("eastAsia", "eastAsiaTheme"),
        ("cs", "csTheme"),
    ] {
        if source.contains_key(explicit) || source.contains_key(theme) {
            result.remove(explicit);
            result.remove(theme);
            if let Some(value) = source.get(explicit) {
                result.insert(explicit.to_owned(), value.clone());
            }
            if let Some(value) = source.get(theme) {
                result.insert(theme.to_owned(), value.clone());
            }
        }
    }
    for (key, value) in source {
        if ![
            "ascii",
            "asciiTheme",
            "hAnsi",
            "hAnsiTheme",
            "eastAsia",
            "eastAsiaTheme",
            "cs",
            "csTheme",
        ]
        .contains(&key.as_str())
        {
            result.insert(key, value);
        }
    }
    Value::Object(result)
}

fn merge_text_formatting(target: Option<&Value>, source: Option<&Value>) -> Option<Value> {
    if object(source).is_none() {
        return target.cloned();
    }
    if object(target).is_none() {
        return source.cloned();
    }
    merge_text_formatting_owned(target.cloned(), source)
}

fn merge_text_formatting_owned(target: Option<Value>, source: Option<&Value>) -> Option<Value> {
    let Some(source_object) = object(source) else {
        return target;
    };
    let Some(Value::Object(mut result)) = target else {
        return source.cloned();
    };
    for (key, value) in source_object {
        if key == "fontFamily" && value.is_object() {
            result.insert(key.clone(), merge_font_family(result.get(key), value));
        } else if key == "color" && value.is_object() {
            let explicit = truthy(field(Some(value), "rgb"))
                || truthy(field(Some(value), "themeColor"))
                || truthy(field(Some(value), "themeTint"))
                || truthy(field(Some(value), "themeShade"));
            if !truthy(field(Some(value), "auto")) || explicit {
                result.insert(key.clone(), value.clone());
            }
        } else if value.is_object() {
            result.insert(
                key.clone(),
                merge_plain(result.get(key), Some(value)).unwrap(),
            );
        } else {
            result.insert(key.clone(), value.clone());
        }
    }
    Some(Value::Object(result))
}

fn merge_paragraph_formatting(target: Option<&Value>, source: Option<&Value>) -> Option<Value> {
    let Some(source) = object(source) else {
        return target.cloned();
    };
    let mut result = object(target).cloned().unwrap_or_default();
    if let Some(value) = source
        .get("indentFirstLine")
        .filter(|value| !value.is_null())
    {
        result.insert("indentFirstLine".to_owned(), value.clone());
        match source.get("hangingIndent").filter(|value| !value.is_null()) {
            Some(hanging) => {
                result.insert("hangingIndent".to_owned(), hanging.clone());
            }
            None => {
                result.remove("hangingIndent");
            }
        }
    }
    for (key, value) in source {
        if key == "runProperties" {
            if let Some(merged) = merge_text_formatting(result.get(key), Some(value)) {
                result.insert(key.clone(), merged);
            }
        } else if ["borders", "numPr", "frame"].contains(&key.as_str()) {
            result.insert(
                key.clone(),
                merge_plain(result.get(key), Some(value)).unwrap_or_else(|| value.clone()),
            );
        } else if matches!(key.as_str(), "indentFirstLine" | "hangingIndent") {
            continue;
        } else {
            result.insert(key.clone(), value.clone());
        }
    }
    Some(Value::Object(result))
}

impl StyleResolver {
    fn new(definitions: Option<&Value>) -> Self {
        let Some(definitions) = object(definitions) else {
            return Self::default();
        };
        let mut resolver = Self {
            enabled: true,
            doc_defaults: definitions.get("docDefaults").cloned(),
            ..Self::default()
        };
        for style in array(definitions.get("styles")) {
            let Some(style_id) = string(field(Some(style), "styleId")) else {
                continue;
            };
            resolver.styles.insert(style_id.to_owned(), style.clone());
        }
        resolver.default_paragraph = resolver.find_default("paragraph").or_else(|| {
            resolver
                .styles
                .contains_key("Normal")
                .then(|| "Normal".to_owned())
        });
        resolver.default_table = resolver.find_default("table");
        resolver.default_character = resolver.find_default("character");
        resolver
    }

    fn find_default(&self, style_type: &str) -> Option<String> {
        self.styles.iter().find_map(|(id, style)| {
            (string(field(Some(style), "type")) == Some(style_type)
                && truthy(field(Some(style), "default")))
            .then(|| id.clone())
        })
    }

    fn style(&self, style_id: &str) -> Option<&Value> {
        self.styles.get(style_id)
    }

    fn default_style(&self, style_type: &str) -> Option<&Value> {
        let id = match style_type {
            "paragraph" => self.default_paragraph.as_deref(),
            "table" => self.default_table.as_deref(),
            "character" => self.default_character.as_deref(),
            _ => None,
        };
        id.and_then(|id| self.style(id))
    }

    fn set_table_paragraph_formatting(&mut self, formatting: Option<Value>) -> Option<Value> {
        let previous = std::mem::replace(&mut self.table_paragraph_formatting, formatting);
        let unchanged = match (&previous, &self.table_paragraph_formatting) {
            (None, None) => true,
            (Some(previous), Some(current)) => identical_json(previous, current),
            _ => false,
        };
        if !unchanged {
            self.memo = StyleMemo::default();
        }
        previous
    }

    fn restore_table_paragraph_formatting(&mut self, state: Option<Value>) {
        let _ = self.set_table_paragraph_formatting(state);
    }

    fn resolve_paragraph_style(&self, style_id: Option<&str>) -> Arc<ResolvedParagraphStyle> {
        let defined_id = style_id.filter(|id| self.style(id).is_some());
        let mut memo = self
            .memo
            .state
            .lock()
            .unwrap_or_else(|error| error.into_inner());
        let cached = match (style_id, defined_id) {
            (None, _) => memo.paragraphs.absent.as_ref(),
            (_, Some(id)) => memo.paragraphs.styles.get(id),
            _ => memo.paragraphs.undefined.as_ref(),
        };
        if let Some(cached) = cached {
            return Arc::clone(cached);
        }
        let (paragraph, run) = self.resolve_paragraph_style_uncached(style_id);
        let run = run.map(Arc::new);
        let default_character = self
            .default_style("character")
            .and_then(|style| field(Some(style), "rPr"));
        let default_run = if default_character.is_some() {
            merge_text_formatting(run.as_deref(), default_character).map(Arc::new)
        } else {
            run.clone()
        };
        let shared_default = default_character.is_none();
        let size = STYLE_MEMO_ENTRY_BYTES
            + defined_id.map_or(0, str::len)
            + [
                paragraph.as_ref(),
                run.as_deref(),
                default_run.as_deref().filter(|_| !shared_default),
            ]
            .into_iter()
            .flatten()
            .map(approx_bytes)
            .sum::<usize>();
        let resolved = Arc::new(ResolvedParagraphStyle {
            paragraph,
            run,
            default_run,
        });
        if memo.bytes + size <= STYLE_MEMO_BYTES {
            memo.bytes += size;
            match (style_id, defined_id) {
                (None, _) => memo.paragraphs.absent = Some(Arc::clone(&resolved)),
                (_, Some(id)) => {
                    memo.paragraphs
                        .styles
                        .insert(id.to_owned(), Arc::clone(&resolved));
                }
                _ => memo.paragraphs.undefined = Some(Arc::clone(&resolved)),
            }
        }
        resolved
    }

    fn resolve_paragraph_style_uncached(
        &self,
        style_id: Option<&str>,
    ) -> (Option<Value>, Option<Value>) {
        let mut paragraph = merge_paragraph_formatting(
            field(self.doc_defaults.as_ref(), "pPr"),
            self.table_paragraph_formatting.as_ref(),
        );
        let mut run = field(self.doc_defaults.as_ref(), "rPr").cloned();
        let style = style_id
            .and_then(|id| self.style(id))
            .or_else(|| self.default_style("paragraph"));
        if let Some(style) = style {
            paragraph = merge_paragraph_formatting(paragraph.as_ref(), field(Some(style), "pPr"));
            run = merge_text_formatting(run.as_ref(), field(Some(style), "rPr"));
        }
        if style_id.is_some() && style.is_none() {
            if let Some(style) = self.default_style("paragraph") {
                paragraph =
                    merge_paragraph_formatting(paragraph.as_ref(), field(Some(style), "pPr"));
                run = merge_text_formatting(run.as_ref(), field(Some(style), "rPr"));
            }
        }
        if style_id.is_none() && style.is_none() && self.doc_defaults.is_none() {
            paragraph = merge_paragraph_formatting(
                Some(&json!({
                    "spaceAfter": 160,
                    "lineSpacing": 259,
                    "lineSpacingRule": "auto"
                })),
                self.table_paragraph_formatting.as_ref(),
            );
        }
        (paragraph, run)
    }

    fn resolve_run_style(&self, style_id: Option<&str>) -> Option<Arc<Value>> {
        let defined_id = style_id.filter(|id| self.style(id).is_some());
        let mut memo = self
            .memo
            .state
            .lock()
            .unwrap_or_else(|error| error.into_inner());
        let cached = match defined_id {
            Some(id) => memo.runs.styles.get(id),
            None => memo.runs.unstyled.as_ref(),
        };
        if let Some(cached) = cached {
            return cached.clone();
        }
        let resolved = self.resolve_run_style_uncached(style_id).map(Arc::new);
        let size = STYLE_MEMO_ENTRY_BYTES
            + defined_id.map_or(0, str::len)
            + resolved.as_deref().map_or(0, approx_bytes);
        if memo.bytes + size <= STYLE_MEMO_BYTES {
            memo.bytes += size;
            match defined_id {
                Some(id) => {
                    memo.runs.styles.insert(id.to_owned(), resolved.clone());
                }
                None => memo.runs.unstyled = Some(resolved.clone()),
            }
        }
        resolved
    }

    fn resolve_run_style_uncached(&self, style_id: Option<&str>) -> Option<Value> {
        let mut result = field(self.doc_defaults.as_ref(), "rPr").cloned();
        result = merge_text_formatting(
            result.as_ref(),
            self.default_style("character")
                .and_then(|style| field(Some(style), "rPr")),
        );
        if let Some(style) = style_id.and_then(|id| self.style(id)) {
            result = merge_text_formatting(result.as_ref(), field(Some(style), "rPr"));
        }
        result
    }

    fn run_style_own(&self, style_id: Option<&str>) -> Option<&Value> {
        style_id
            .and_then(|id| self.style(id))
            .and_then(|style| field(Some(style), "rPr"))
    }
}

fn mark(name: &str, attrs: Vec<(String, Value)>) -> Mark {
    Mark {
        name: name.to_owned(),
        attrs,
    }
}

fn formatting_to_marks(formatting: Option<&Value>) -> Vec<Mark> {
    let mut marks = Vec::new();
    let Some(formatting) = object(formatting) else {
        return marks;
    };
    if truthy(formatting.get("bold")) {
        marks.push(mark("bold", vec![]));
    }
    if truthy(formatting.get("italic")) {
        marks.push(mark("italic", vec![]));
    }
    if let Some(underline) = object(formatting.get("underline"))
        && string(underline.get("style")) != Some("none")
    {
        marks.push(mark(
            "underline",
            ordered_object([
                (
                    "style",
                    underline.get("style").cloned().unwrap_or(Value::Null),
                ),
                ("color", nullish(underline.get("color"))),
            ]),
        ));
    }
    if truthy(formatting.get("strike")) || truthy(formatting.get("doubleStrike")) {
        marks.push(mark(
            "strike",
            ordered_object([(
                "double",
                Value::Bool(truthy(formatting.get("doubleStrike"))),
            )]),
        ));
    }
    if let Some(color) = object(formatting.get("color"))
        && !truthy(color.get("auto"))
    {
        marks.push(mark(
            "textColor",
            ordered_object([
                ("rgb", nullish(color.get("rgb"))),
                ("themeColor", nullish(color.get("themeColor"))),
                ("themeTint", nullish(color.get("themeTint"))),
                ("themeShade", nullish(color.get("themeShade"))),
            ]),
        ));
    }
    let shading_fill =
        object(formatting.get("shading")).and_then(|shading| object(shading.get("fill")));
    let shading_highlight = shading_fill.and_then(|fill| {
        let pattern =
            string(object(formatting.get("shading")).and_then(|shading| shading.get("pattern")));
        (pattern.is_none() || pattern == Some("clear"))
            .then(|| string(fill.get("rgb")))
            .flatten()
            .filter(|_| !truthy(fill.get("auto")))
            .map(|rgb| {
                if rgb.starts_with('#') {
                    rgb.to_owned()
                } else {
                    format!("#{rgb}")
                }
            })
    });
    let highlight = string(formatting.get("highlight"))
        .filter(|value| *value != "none")
        .map(str::to_owned)
        .or(shading_highlight);
    if let Some(highlight) = highlight {
        marks.push(mark(
            "highlight",
            ordered_object([("color", Value::String(highlight))]),
        ));
    }
    if formatting.contains_key("fontSize") || formatting.contains_key("fontSizeCs") {
        marks.push(mark(
            "fontSize",
            ordered_object([
                ("size", nullish(formatting.get("fontSize"))),
                ("sizeCs", nullish(formatting.get("fontSizeCs"))),
            ]),
        ));
    }
    if let Some(font) = object(formatting.get("fontFamily")) {
        marks.push(mark(
            "fontFamily",
            ordered_object([
                ("ascii", nullish(font.get("ascii"))),
                ("hAnsi", nullish(font.get("hAnsi"))),
                ("eastAsia", nullish(font.get("eastAsia"))),
                ("cs", nullish(font.get("cs"))),
                ("asciiTheme", nullish(font.get("asciiTheme"))),
                ("hAnsiTheme", nullish(font.get("hAnsiTheme"))),
                ("eastAsiaTheme", nullish(font.get("eastAsiaTheme"))),
                ("csTheme", nullish(font.get("csTheme"))),
            ]),
        ));
    }
    match string(formatting.get("vertAlign")) {
        Some("superscript") => marks.push(mark("superscript", vec![])),
        Some("subscript") => marks.push(mark("subscript", vec![])),
        _ => {}
    }
    for (key, name) in [
        ("allCaps", "allCaps"),
        ("smallCaps", "smallCaps"),
        ("emboss", "emboss"),
        ("imprint", "imprint"),
        ("shadow", "textShadow"),
        ("outline", "textOutline"),
        ("hidden", "hidden"),
        ("rtl", "rtl"),
    ] {
        if truthy(formatting.get(key)) {
            marks.push(mark(name, vec![]));
        }
    }
    // Document-grid opt-out (w:snapToGrid, default on): only an authored off
    // becomes a mark, mirroring documentToYrs.
    if formatting.get("snapToGrid") == Some(&Value::Bool(false)) {
        marks.push(mark("snapToGrid", vec![]));
    }
    if ["spacing", "position", "scale", "kerning"]
        .iter()
        .any(|key| formatting.contains_key(*key))
    {
        marks.push(mark(
            "characterSpacing",
            ordered_object([
                ("spacing", nullish(formatting.get("spacing"))),
                ("position", nullish(formatting.get("position"))),
                ("scale", nullish(formatting.get("scale"))),
                ("kerning", nullish(formatting.get("kerning"))),
            ]),
        ));
    }
    if let Some(value) = string(formatting.get("emphasisMark")).filter(|value| *value != "none") {
        marks.push(mark(
            "emphasisMark",
            ordered_object([("type", Value::String(value.to_owned()))]),
        ));
    }
    if let Some(value) = string(formatting.get("effect")).filter(|value| *value != "none") {
        marks.push(mark(
            "textEffect",
            ordered_object([("effect", Value::String(value.to_owned()))]),
        ));
    }
    if let Some(value) = formatting.get("modernEffects") {
        marks.push(mark(
            "modernTextEffects",
            ordered_object([("effects", value.clone())]),
        ));
    }
    if let Some(value) = formatting.get("styleId") {
        marks.push(mark(
            "runStyle",
            ordered_object([("styleId", value.clone())]),
        ));
    }
    marks
}

/// The story attributes of a paragraph mark's run defaults (`defaultTextFormatting`), lowered as
/// seeding lowers a run's formatting.
pub(crate) fn mark_run_attrs(defaults: &Any) -> Vec<(String, Any)> {
    serde_json::to_value(defaults)
        .ok()
        .and_then(|value| payload(marks_to_attrs(&formatting_to_marks(Some(&value)))).ok())
        .unwrap_or_default()
}

fn mark_attrs(mark: &Mark) -> Value {
    Value::Object(mark.attrs.iter().cloned().collect())
}

fn marks_to_attrs(marks: &[Mark]) -> JsonObject {
    let boolean_marks = [
        "bold",
        "italic",
        "superscript",
        "subscript",
        "allCaps",
        "smallCaps",
        "emboss",
        "imprint",
        "textShadow",
        "textOutline",
        "hidden",
        "rtl",
    ];
    let mut attrs = JsonObject::new();
    for mark in marks {
        if mark.name == "comment" || mark.name == "footnoteRef" {
            continue;
        }
        if boolean_marks.contains(&mark.name.as_str()) {
            attrs.insert(mark.name.clone(), Value::Bool(true));
        } else if mark.name == "snapToGrid" {
            attrs.insert("snapToGrid".to_owned(), Value::Bool(false));
        } else if mark.name == "highlight" {
            attrs.insert(
                "highlight".to_owned(),
                mark.attrs
                    .iter()
                    .find(|(key, _)| key == "color")
                    .map(|(_, value)| value.clone())
                    .unwrap_or(Value::Null),
            );
        } else if mark.name == "insertion" || mark.name == "deletion" {
            let get = |name: &str| {
                mark.attrs
                    .iter()
                    .find(|(key, _)| key == name)
                    .map(|(_, value)| value.clone())
                    .unwrap_or(Value::Null)
            };
            attrs.insert(
                if mark.name == "insertion" {
                    "ins".to_owned()
                } else {
                    "del".to_owned()
                },
                drop_nulls(json!({
                    "id": get("revisionId"),
                    "author": get("author"),
                    "date": get("date")
                })),
            );
        } else {
            attrs.insert(mark.name.clone(), drop_nulls(mark_attrs(mark)));
        }
    }
    attrs
}

fn marks_key(marks: &[Mark]) -> String {
    let mut values: Vec<String> = marks
        .iter()
        .filter(|mark| mark.name != "hyperlink" && mark.name != "comment")
        .map(|mark| format!("{}:{}", mark.name, ordered_json(&mark.attrs)))
        .collect();
    values.sort();
    values.join("|")
}

fn with_mark(marks: &[Mark], next: Mark) -> Vec<Mark> {
    let name = next.name.clone();
    marks
        .iter()
        .filter(|mark| mark.name != name)
        .cloned()
        .chain(std::iter::once(next))
        .collect()
}

fn text_unit(text: String, marks: &[Mark]) -> InlineUnit {
    InlineUnit {
        pm_size: utf16_len(&text),
        content: UnitContent::Text(text),
        attrs: marks_to_attrs(marks),
        marks: marks.to_vec(),
    }
}

fn embed_unit(kind: &str, payload: JsonObject, marks: &[Mark], pm_size: u32) -> InlineUnit {
    InlineUnit {
        content: UnitContent::Embed {
            kind: kind.to_owned(),
            payload,
        },
        attrs: marks_to_attrs(marks),
        pm_size,
        marks: marks.to_vec(),
    }
}

fn run_marks(run: &Value, style_formatting: Option<&Value>, styles: &StyleResolver) -> Vec<Mark> {
    let formatting = field(Some(run), "formatting");
    let style_id = string(field(formatting, "styleId"));
    let run_style = styles.run_style_own(style_id);
    let inherited = merge_text_formatting(style_formatting, run_style);
    let merged = merge_text_formatting_owned(inherited, formatting);
    let mut marks = formatting_to_marks(merged.as_ref());
    let hyperlink_style = style_id.is_some_and(is_hyperlink_style_name)
        || style_id
            .and_then(|id| styles.style(id))
            .and_then(|style| string(field(Some(style), "name")))
            .is_some_and(is_hyperlink_style_name);
    if hyperlink_style {
        for (property, name) in [("color", "textColor"), ("underline", "underline")] {
            if field(formatting, property).is_none()
                && field(run_style, property).is_some()
                && let Some(mark) = marks.iter_mut().find(|mark| mark.name == name)
            {
                mark.attrs
                    .push(("inheritedHyperlink".to_owned(), Value::Bool(true)));
            }
        }
    }
    marks
}

fn is_hyperlink_style_name(name: &str) -> bool {
    name.eq_ignore_ascii_case("Hyperlink") || name.eq_ignore_ascii_case("FollowedHyperlink")
}

fn emu_to_pixels(value: f64) -> f64 {
    value / 914_400.0 * 96.0
}

fn image_payload(image: &Value) -> JsonObject {
    let size = field(Some(image), "size");
    let wrap = field(Some(image), "wrap");
    let position = field(Some(image), "position");
    let transform = field(Some(image), "transform");
    let outline = field(Some(image), "outline");
    let wrap_type = string(field(wrap, "type")).unwrap_or_default();
    let wrap_text = string(field(wrap, "wrapText"));
    let horizontal = field(position, "horizontal");
    let vertical = field(position, "vertical");
    let alignment = string(field(horizontal, "alignment"));
    let css_float = if wrap_type == "inline" || wrap_type == "topAndBottom" {
        "none"
    } else if ["square", "tight", "through"].contains(&wrap_type) {
        if wrap_text == Some("left") {
            "right"
        } else if wrap_text == Some("right") {
            "left"
        } else if matches!(alignment, Some("left" | "right")) {
            alignment.unwrap()
        } else {
            "none"
        }
    } else {
        "none"
    };
    let display_mode = if wrap_type == "inline" {
        "inline"
    } else if wrap_type == "topAndBottom" {
        "block"
    } else if matches!(wrap_type, "behind" | "inFront") || css_float != "none" {
        "float"
    } else {
        "block"
    };
    let mut transforms = Vec::new();
    if let Some(rotation) = number(field(transform, "rotation")).filter(|value| *value != 0.0) {
        transforms.push(format!("rotate({rotation}deg)"));
    }
    if truthy(field(transform, "flipH")) {
        transforms.push("scaleX(-1)".to_owned());
    }
    if truthy(field(transform, "flipV")) {
        transforms.push("scaleY(-1)".to_owned());
    }
    let outline_width = number(field(outline, "width")).filter(|value| *value != 0.0);
    let border_width =
        outline_width.map(|value| (value / 914_400.0 * 96.0 * 100.0).round() / 100.0);
    let border_color =
        string(field(field(outline, "color"), "rgb")).map(|value| format!("#{value}"));
    let border_style = outline_width.map(|_| match string(field(outline, "style")) {
        Some("dot" | "sysDot") => "dotted",
        Some(
            "dash" | "lgDash" | "dashDot" | "lgDashDot" | "lgDashDotDot" | "sysDash" | "sysDashDot"
            | "sysDashDotDot",
        ) => "dashed",
        _ => "solid",
    });
    let axis = |axis: Option<&Value>| {
        axis.map(|axis| {
            json!({
                "relativeTo": nullish(field(Some(axis), "relativeTo")),
                "posOffset": nullish(field(Some(axis), "posOffset")),
                "align": nullish(field(Some(axis), "alignment"))
            })
        })
    };
    map_from_value(json!({
        "src": string(field(Some(image), "src")).unwrap_or_default(),
        "alt": nullish(field(Some(image), "alt")),
        "title": nullish(field(Some(image), "title")),
        "width": number(field(size, "width")).filter(|value| *value != 0.0).map(emu_to_pixels),
        "height": number(field(size, "height")).filter(|value| *value != 0.0).map(emu_to_pixels),
        "rId": nullish(field(Some(image), "rId")),
        "wrapType": wrap_type,
        "displayMode": display_mode,
        "cssFloat": css_float,
        "transform": (!transforms.is_empty()).then(|| transforms.join(" ")),
        "distTop": number(field(wrap, "distT")).map(emu_to_pixels),
        "distBottom": number(field(wrap, "distB")).map(emu_to_pixels),
        "distLeft": number(field(wrap, "distL")).map(emu_to_pixels),
        "distRight": number(field(wrap, "distR")).map(emu_to_pixels),
        "position": position.map(|_| json!({
            "horizontal": axis(horizontal),
            "vertical": axis(vertical),
            "relativeHeight": nullish(field(position, "relativeHeight")),
            "behindDoc": nullish(field(position, "behindDoc"))
        })),
        "borderWidth": border_width,
        "borderColor": border_color,
        "borderColorValue": nullish(field(outline, "color")),
        "borderStyle": border_style,
        "wrapText": wrap_text,
        "hlinkHref": nullish(field(Some(image), "hlinkHref")),
        "cropTop": nullish(field(field(Some(image), "crop"), "top")),
        "cropRight": nullish(field(field(Some(image), "crop"), "right")),
        "cropBottom": nullish(field(field(Some(image), "crop"), "bottom")),
        "cropLeft": nullish(field(field(Some(image), "crop"), "left")),
        "shapeType": nullish(field(Some(image), "shapeType")),
        "opacity": nullish(field(Some(image), "opacity")),
        "effectExtentTop": number(field(field(Some(image), "padding"), "top"))
            .map(emu_to_pixels),
        "effectExtentBottom": number(field(field(Some(image), "padding"), "bottom"))
            .map(emu_to_pixels),
        "effectExtentLeft": number(field(field(Some(image), "padding"), "left"))
            .map(emu_to_pixels),
        "effectExtentRight": number(field(field(Some(image), "padding"), "right"))
            .map(emu_to_pixels),
        "layoutInCell": nullish(field(Some(image), "layoutInCell")),
        "allowOverlap": nullish(field(Some(image), "allowOverlap"))
    }))
}

fn shape_payload(shape: &Value, source: &BTreeMap<String, String>) -> JsonObject {
    map_from_value(json!({ "shapeJson": source_json(shape, source) }))
}

fn chart_payload(chart: &Value, source: &BTreeMap<String, String>) -> JsonObject {
    let size = field(Some(chart), "size");
    map_from_value(json!({
        "chartJson": source_json(chart, source),
        "chartType": nullish(field(Some(chart), "chartType")),
        "title": nullish(field(Some(chart), "title")),
        "width": number(field(size, "width")).filter(|value| *value != 0.0).map(emu_to_pixels).unwrap_or(320.0),
        "height": number(field(size, "height")).filter(|value| *value != 0.0).map(emu_to_pixels).unwrap_or(220.0),
        "rId": nullish(field(Some(chart), "rId")),
        "path": nullish(field(Some(chart), "path"))
    }))
}

pub(crate) fn numeric_field_instruction(instruction: &str) -> bool {
    let instruction = instruction.trim();
    !instruction.is_empty() && instruction.bytes().all(|byte| byte.is_ascii_digit())
}

fn field_payload(
    field_value: &Value,
    style_formatting: Option<&Value>,
    source: &BTreeMap<String, String>,
) -> (JsonObject, Vec<Mark>) {
    let kind = string(field(Some(field_value), "type")).unwrap_or_default();
    let runs = if kind == "simpleField" {
        array(field(Some(field_value), "content"))
    } else {
        array(field(Some(field_value), "fieldResult"))
    };
    let mut display_text = String::new();
    let mut field_formatting = None;
    for child in runs {
        if string(field(Some(child), "type")) != Some("run") {
            continue;
        }
        for content in array(field(Some(child), "content")) {
            if string(field(Some(content), "type")) == Some("text") {
                display_text.push_str(string(field(Some(content), "text")).unwrap_or_default());
            }
        }
        if field_formatting.is_none() {
            field_formatting = field(Some(child), "formatting");
        }
    }
    let formatting = field_formatting.or_else(|| {
        (kind == "complexField")
            .then(|| field(Some(field_value), "formatting"))
            .flatten()
    });
    let merged = merge_text_formatting(style_formatting, formatting);
    (
        map_from_value(json!({
            "fieldType": nullish(field(Some(field_value), "fieldType")),
            "instruction": nullish(field(Some(field_value), "instruction")),
            "displayText": display_text,
            "fieldKind": if kind == "simpleField" { "simple" } else { "complex" },
            "fldLock": boolean(field(Some(field_value), "fldLock")).unwrap_or(false),
            "dirty": boolean(field(Some(field_value), "dirty")).unwrap_or(false),
            "displayMode": string(field(field(Some(field_value), "fieldTree"), "displayMode")).unwrap_or("result"),
            "hasCachedResult": !display_text.is_empty(),
            "fieldData": source_json(field_value, source),
            "modelKind": "field"
        })),
        formatting_to_marks(merged.as_ref()),
    )
}

/// The sequences of SEQ fields nested anywhere in a field's code or result.
pub(crate) fn nested_sequence_names(field_value: &Value) -> Vec<String> {
    let mut names = Vec::new();
    let mut seen = HashSet::new();
    let mut pending = Vec::new();
    if let Value::Object(map) = field_value {
        pending.extend(map.values().rev().map(|value| (value, false)));
    }
    while let Some((value, scanned_hyperlink)) = pending.pop() {
        match value {
            Value::Object(map) => {
                if matches!(
                    string(map.get("type")),
                    Some("complexField" | "simpleField")
                ) && let Some(name) = string(map.get("instruction"))
                    .and_then(docx_layout::sequence_fields::sequence_name)
                    && seen.insert(name.clone())
                {
                    names.push(name);
                }
                if string(map.get("type")) == Some("hyperlink") && !scanned_hyperlink {
                    names.extend(
                        hyperlink_sequence_names(value)
                            .into_iter()
                            .filter(|name| seen.insert(name.clone())),
                    );
                }
                let scanned_hyperlink =
                    scanned_hyperlink || string(map.get("type")) == Some("hyperlink");
                pending.extend(map.values().rev().map(|value| (value, scanned_hyperlink)));
            }
            Value::Array(values) => {
                pending.extend(values.iter().rev().map(|value| (value, scanned_hyperlink)))
            }
            _ => {}
        }
    }
    names
}

pub(crate) fn hyperlink_sequence_names(hyperlink: &Value) -> Vec<String> {
    let mut pending = vec![hyperlink];
    let mut instructions: Vec<Option<String>> = Vec::new();
    let mut names = Vec::new();
    while let Some(node) = pending.pop() {
        let kind = string(field(Some(node), "type"));
        match kind {
            Some("hyperlink") => {
                let children = field(Some(node), "structuredChildren")
                    .or_else(|| field(Some(node), "children"));
                pending.extend(array(children).iter().rev());
            }
            Some("inlineSdt" | "insertion" | "deletion" | "moveFrom" | "moveTo") => {
                pending.extend(array(field(Some(node), "content")).iter().rev());
            }
            Some("simpleField" | "complexField") => {
                names.extend(
                    string(field(Some(node), "instruction"))
                        .and_then(docx_layout::sequence_fields::sequence_name),
                );
                let tree = field(Some(node), "fieldTree");
                let result = field(field(Some(node), "structuredResult"), "inline")
                    .or_else(|| field(field(tree, "result"), "inline"))
                    .or_else(|| {
                        field(
                            Some(node),
                            if kind == Some("simpleField") {
                                "content"
                            } else {
                                "fieldResult"
                            },
                        )
                    });
                pending.extend(array(result).iter().rev());
                if kind == Some("complexField") {
                    let code = field(field(Some(node), "structuredCode"), "inline")
                        .or_else(|| field(field(tree, "code"), "inline"))
                        .or_else(|| field(Some(node), "fieldCode"));
                    pending.extend(array(code).iter().rev());
                }
            }
            Some("run") => {}
            _ => continue,
        }
        if kind != Some("run") {
            continue;
        }
        for content in array(field(Some(node), "content")) {
            let instruction = match string(field(Some(content), "type")) {
                Some("fieldChar") => match string(field(Some(content), "charType")) {
                    Some("begin") => {
                        instructions.push(Some(String::new()));
                        None
                    }
                    Some("separate") => instructions.last_mut().and_then(Option::take),
                    Some("end") => instructions.pop().flatten(),
                    _ => None,
                },
                Some("instrText") => {
                    if let Some(Some(instruction)) = instructions.last_mut() {
                        instruction
                            .push_str(string(field(Some(content), "text")).unwrap_or_default());
                    }
                    None
                }
                _ => None,
            };
            names.extend(
                instruction
                    .as_deref()
                    .and_then(docx_layout::sequence_fields::sequence_name),
            );
        }
    }
    names
}

fn math_payload(math: &Value) -> JsonObject {
    map_from_value(json!({
        "display": nullish(field(Some(math), "display")),
        "ommlXml": nullish(field(Some(math), "ommlXml")),
        "plainText": string(field(Some(math), "plainText")).unwrap_or_default()
    }))
}

fn hyperlink_mark(hyperlink: &Value) -> Mark {
    let href = string(field(Some(hyperlink), "href"))
        .filter(|value| !value.is_empty())
        .map(str::to_owned)
        .or_else(|| {
            string(field(Some(hyperlink), "anchor"))
                .filter(|value| !value.is_empty())
                .map(|value| format!("#{value}"))
        })
        .unwrap_or_default();
    mark(
        "hyperlink",
        ordered_object([
            ("href", Value::String(href)),
            ("tooltip", nullish(field(Some(hyperlink), "tooltip"))),
            ("rId", nullish(field(Some(hyperlink), "rId"))),
        ]),
    )
}

fn note_ref_unit(id: &Value, note_type: &str, marks: &[Mark]) -> InlineUnit {
    let note_mark = mark(
        "footnoteRef",
        ordered_object([
            ("id", Value::String(js_string(id))),
            ("noteType", Value::String(note_type.to_owned())),
        ]),
    );
    let all_marks: Vec<Mark> = marks
        .iter()
        .cloned()
        .chain(std::iter::once(note_mark))
        .collect();
    embed_unit(
        "noteRef",
        map_from_value(if note_type == "endnote" {
            json!({ "endnoteRefId": id })
        } else {
            json!({ "footnoteRefId": id })
        }),
        &all_marks,
        1,
    )
}

/// The marks a drawing keeps: hidden text and its tracked insertion or deletion.
fn drawing_marks(marks: &[Mark]) -> Vec<Mark> {
    marks
        .iter()
        .filter(|mark| matches!(mark.name.as_str(), "hidden" | "insertion" | "deletion"))
        .cloned()
        .collect()
}

/// Returns nonempty text emitted by a run content node.
fn run_content_text(content: &Value) -> Option<&str> {
    string(field(Some(content), "text")).filter(|text| !text.is_empty())
}

/// Whether a run break emits an inline unit.
fn run_content_is_text_wrapping_break(content: &Value) -> bool {
    string(field(Some(content), "breakType")).is_none_or(|kind| kind == "textWrapping")
}

/// Parses a run symbol's Unicode scalar value.
fn run_content_symbol_char(content: &Value) -> Option<char> {
    string(field(Some(content), "char"))
        .and_then(|value| u32::from_str_radix(value, 16).ok())
        .and_then(char::from_u32)
}

/// Counts emitted units without lowering a run content node.
fn run_content_unit_count(content: &Value) -> usize {
    match string(field(Some(content), "type")).unwrap_or_default() {
        "text" => usize::from(run_content_text(content).is_some()),
        "break" => usize::from(run_content_is_text_wrapping_break(content)),
        "symbol" => usize::from(run_content_symbol_char(content).is_some()),
        "footnoteRef" | "endnoteRef" => usize::from(field(Some(content), "id").is_some()),
        "tab" | "softHyphen" | "noBreakHyphen" | "commentReference" | "drawing"
        | "horizontalRule" | "shape" | "chart" => 1,
        _ => 0,
    }
}

/// Measures emitted UTF-16 text and single-width embeds without lowering.
fn run_content_width(content: &Value) -> u32 {
    match string(field(Some(content), "type")).unwrap_or_default() {
        "text" => run_content_text(content).map(utf16_len).unwrap_or_default(),
        "symbol" => run_content_symbol_char(content)
            .map(|codepoint| codepoint.len_utf16() as u32)
            .unwrap_or_default(),
        _ => run_content_unit_count(content) as u32,
    }
}

fn run_content_to_units(
    content: &Value,
    marks: &[Mark],
    source: &BTreeMap<String, String>,
) -> Vec<InlineUnit> {
    match string(field(Some(content), "type")).unwrap_or_default() {
        "text" => run_content_text(content)
            .map(|text| vec![text_unit(text.to_owned(), marks)])
            .unwrap_or_default(),
        "tab" => vec![text_unit("\t".to_owned(), marks)],
        "break" if run_content_is_text_wrapping_break(content) => {
            vec![embed_unit("break", JsonObject::new(), marks, 1)]
        }
        "softHyphen" => vec![text_unit("\u{00ad}".to_owned(), marks)],
        "noBreakHyphen" => vec![text_unit("\u{2011}".to_owned(), marks)],
        "symbol" => {
            let Some(codepoint) = run_content_symbol_char(content) else {
                return vec![];
            };
            let font = string(field(Some(content), "font"))
                .filter(|value| !value.is_empty())
                .map(|value| Value::String(value.to_owned()))
                .unwrap_or(Value::Null);
            let symbol_mark = mark(
                "fontFamily",
                ordered_object([
                    ("ascii", font.clone()),
                    ("hAnsi", font.clone()),
                    ("eastAsia", font.clone()),
                    ("cs", font),
                    ("asciiTheme", Value::Null),
                    ("hAnsiTheme", Value::Null),
                    ("eastAsiaTheme", Value::Null),
                    ("csTheme", Value::Null),
                ]),
            );
            vec![text_unit(
                codepoint.to_string(),
                &with_mark(marks, symbol_mark),
            )]
        }
        "commentReference" => {
            let mut value = json!({
                "fieldType": "COMMENT",
                "instruction": "",
                "displayText": "",
                "fieldKind": "simple",
                "fldLock": false,
                "dirty": false,
                "displayMode": "result",
                "hasCachedResult": false,
                "modelKind": "commentReference"
            });
            if let Some(id) = field(Some(content), "id") {
                value
                    .as_object_mut()
                    .unwrap()
                    .insert("commentId".to_owned(), id.clone());
            }
            vec![embed_unit("field", map_from_value(value), &[], 1)]
        }
        "drawing" => vec![embed_unit(
            "image",
            image_payload(field(Some(content), "image").unwrap_or(&Value::Null)),
            &drawing_marks(marks),
            1,
        )],
        "horizontalRule" => vec![embed_unit(
            "horizontalRule",
            map_from_value(json!({"rule": field(Some(content), "rule")})),
            marks,
            1,
        )],
        "shape" => vec![embed_unit(
            "shape",
            shape_payload(
                field(Some(content), "shape").unwrap_or(&Value::Null),
                source,
            ),
            &drawing_marks(marks),
            1,
        )],
        "chart" => vec![embed_unit(
            "chart",
            chart_payload(
                field(Some(content), "chart").unwrap_or(&Value::Null),
                source,
            ),
            &drawing_marks(marks),
            1,
        )],
        "footnoteRef" => field(Some(content), "id")
            .map(|id| note_ref_unit(id, "footnote", marks))
            .into_iter()
            .collect(),
        "endnoteRef" => field(Some(content), "id")
            .map(|id| note_ref_unit(id, "endnote", marks))
            .into_iter()
            .collect(),
        _ => vec![],
    }
}

fn run_to_units(
    run: &Value,
    style_formatting: Option<&Value>,
    styles: &StyleResolver,
    extra_marks: &[Mark],
    source: &BTreeMap<String, String>,
) -> Vec<InlineUnit> {
    let marks: Vec<Mark> = run_marks(run, style_formatting, styles)
        .into_iter()
        .chain(extra_marks.iter().cloned())
        .collect();
    array(field(Some(run), "content"))
        .iter()
        .flat_map(|content| run_content_to_units(content, &marks, source))
        .collect()
}

fn has_tracked_control(value: &Value, control: bool, revision: bool) -> bool {
    let mut pending = vec![(value, control, revision)];
    while let Some((node, control, revision)) = pending.pop() {
        let kind = string(field(Some(node), "type"));
        let control = control || kind == Some("inlineSdt");
        let revision =
            revision || matches!(kind, Some("insertion" | "deletion" | "moveFrom" | "moveTo"));
        if control && revision {
            return true;
        }
        let children = match kind {
            Some("inlineSdt" | "insertion" | "deletion" | "moveFrom" | "moveTo") => {
                field(Some(node), "content")
            }
            Some("hyperlink") => {
                field(Some(node), "structuredChildren").or_else(|| field(Some(node), "children"))
            }
            _ => None,
        };
        pending.extend(
            array(children)
                .iter()
                .map(|child| (child, control, revision)),
        );
    }
    false
}

fn hyperlink_to_units(
    hyperlink: &Value,
    style_formatting: Option<&Value>,
    styles: &StyleResolver,
    extra_marks: &[Mark],
    source: &BTreeMap<String, String>,
    opaque_sequences: &mut Vec<String>,
) -> Vec<InlineUnit> {
    let mut units = Vec::new();
    let link = hyperlink_mark(hyperlink);
    let children =
        field(Some(hyperlink), "structuredChildren").or_else(|| field(Some(hyperlink), "children"));
    for child in array(children) {
        match string(field(Some(child), "type")).unwrap_or_default() {
            "run" => {
                let marks: Vec<Mark> = run_marks(child, style_formatting, styles)
                    .into_iter()
                    .chain(extra_marks.iter().cloned())
                    .chain(std::iter::once(link.clone()))
                    .collect();
                for content in array(field(Some(child), "content")) {
                    units.extend(run_content_to_units(content, &marks, source));
                }
            }
            "simpleField" | "complexField" => {
                opaque_sequences.extend(nested_sequence_names(child));
                let (payload, marks) = field_payload(child, style_formatting, source);
                let marks: Vec<Mark> = marks
                    .into_iter()
                    .chain(extra_marks.iter().cloned())
                    .chain(std::iter::once(link.clone()))
                    .collect();
                units.push(embed_unit("field", payload, &marks, 1));
            }
            "mathEquation" => {
                let marks: Vec<Mark> = extra_marks
                    .iter()
                    .cloned()
                    .chain(std::iter::once(link.clone()))
                    .collect();
                units.push(embed_unit("math", math_payload(child), &marks, 1));
            }
            "inlineSdt"
                if extra_marks
                    .iter()
                    .any(|mark| matches!(mark.name.as_str(), "insertion" | "deletion"))
                    || has_tracked_control(child, false, false) =>
            {
                let marks: Vec<Mark> = extra_marks
                    .iter()
                    .cloned()
                    .chain(std::iter::once(link.clone()))
                    .collect();
                units.push(embed_unit(
                    "sdt",
                    sdt_payload(child, style_formatting, styles, source, opaque_sequences),
                    &marks,
                    2,
                ));
            }
            _ => {}
        }
    }
    units
}

fn field_to_units(
    value: &Value,
    style_formatting: Option<&Value>,
    styles: &StyleResolver,
    source: &BTreeMap<String, String>,
    projection_id: usize,
    opaque_sequences: &mut Vec<String>,
) -> Vec<InlineUnit> {
    let result = array(field(field(Some(value), "structuredResult"), "inline"));
    let code = array(field(field(Some(value), "structuredCode"), "inline"));
    let projected_children: Vec<_> = code
        .iter()
        .enumerate()
        .map(|(index, child)| (-(index as isize) - 1, child))
        .chain(
            result
                .iter()
                .enumerate()
                .map(|(index, child)| (index as isize, child)),
        )
        .collect();
    if numeric_field_instruction(string(field(Some(value), "instruction")).unwrap_or_default())
        || string(field(Some(value), "type")) != Some("complexField")
        || !projected_children.iter().any(|(_, child)| {
            matches!(
                string(field(Some(child), "type")),
                Some("hyperlink" | "simpleField")
            )
        })
    {
        opaque_sequences.extend(nested_sequence_names(value));
        let (payload, marks) = field_payload(value, style_formatting, source);
        return vec![embed_unit("field", payload, &marks, 1)];
    }
    let mut units = Vec::new();
    let mut children = Vec::new();
    for (index, child) in projected_children {
        let mut projected = match string(field(Some(child), "type")) {
            Some("hyperlink") => {
                opaque_sequences.extend(hyperlink_sequence_names(child));
                hyperlink_to_units(
                    child,
                    style_formatting,
                    styles,
                    &[],
                    source,
                    opaque_sequences,
                )
            }
            Some("simpleField") => {
                opaque_sequences.extend(nested_sequence_names(child));
                let (payload, marks) = field_payload(child, style_formatting, source);
                vec![embed_unit("field", payload, &marks, 1)]
            }
            _ => continue,
        };
        let items: Vec<Value> = projected.iter().map(|unit| match &unit.content {
            UnitContent::Text(text) => json!({"kind":"text", "text":text, "attributes":unit.attrs}),
            UnitContent::Embed {kind, payload} => json!({"kind":"embed", "embedKind":kind, "payload":payload, "attributes":unit.attrs}),
        }).collect();
        children.push(json!({"index":index, "items":items}));
        for unit in &mut projected {
            unit.attrs.insert(
                "fieldResult".to_owned(),
                json!({"id":projection_id, "index":index}),
            );
        }
        units.extend(projected);
    }
    let mut visible = value.clone();
    visible["fieldResult"] = Value::Array(
        result
            .iter()
            .filter(|child| string(field(Some(child), "type")) == Some("run"))
            .cloned()
            .collect(),
    );
    let (mut payload, marks) = field_payload(&visible, style_formatting, source);
    let sequence_owner = string(field(Some(value), "instruction"))
        .and_then(docx_layout::sequence_fields::sequence_name)
        .is_some();
    opaque_sequences.extend(nested_sequence_names(if sequence_owner {
        value
    } else {
        &visible
    }));
    payload.insert(
        "fieldData".to_owned(),
        Value::String(source_json(value, source)),
    );
    payload.insert(
        "resultProjection".to_owned(),
        json!({"id":projection_id, "children":children}),
    );
    units.push(embed_unit("field", payload, &marks, 1));
    units
}

fn tracked_mark(info: &Value, kind: &str, is_move_pair: bool) -> Mark {
    mark(
        kind,
        ordered_object([
            ("revisionId", nullish(field(Some(info), "id"))),
            ("author", nullish(field(Some(info), "author"))),
            ("date", nullish(field(Some(info), "date"))),
            ("isMovePair", Value::Bool(is_move_pair)),
        ]),
    )
}

fn tracked_to_units(
    content: &Value,
    style_formatting: Option<&Value>,
    styles: &StyleResolver,
    source: &BTreeMap<String, String>,
    opaque_sequences: &mut Vec<String>,
) -> Vec<InlineUnit> {
    tracked_to_units_in_control(
        content,
        style_formatting,
        styles,
        source,
        opaque_sequences,
        false,
    )
}

fn tracked_to_units_in_control(
    content: &Value,
    style_formatting: Option<&Value>,
    styles: &StyleResolver,
    source: &BTreeMap<String, String>,
    opaque_sequences: &mut Vec<String>,
    in_control: bool,
) -> Vec<InlineUnit> {
    let content_type = string(field(Some(content), "type")).unwrap_or_default();
    let kind = if matches!(content_type, "insertion" | "moveTo") {
        "insertion"
    } else {
        "deletion"
    };
    let marker = tracked_mark(
        field(Some(content), "info").unwrap_or(&Value::Null),
        kind,
        matches!(content_type, "moveFrom" | "moveTo"),
    );
    let mut units = Vec::new();
    for child in array(field(Some(content), "content")) {
        if string(field(Some(child), "type")) == Some("run") {
            units.extend(run_to_units(
                child,
                style_formatting,
                styles,
                std::slice::from_ref(&marker),
                source,
            ));
        } else if string(field(Some(child), "type")) == Some("hyperlink") {
            opaque_sequences.extend(hyperlink_sequence_names(child));
            let linked = hyperlink_to_units(
                child,
                style_formatting,
                styles,
                std::slice::from_ref(&marker),
                source,
                opaque_sequences,
            );
            units.extend(linked);
        } else if in_control
            || string(field(Some(child), "type")) == Some("inlineSdt")
            || has_tracked_control(child, false, true)
        {
            let inherited = marks_to_attrs(std::slice::from_ref(&marker));
            for mut unit in inline_container_units(
                child,
                style_formatting,
                styles,
                source,
                opaque_sequences,
                in_control,
            ) {
                let mut attrs = inherited.clone();
                attrs.extend(unit.attrs);
                unit.attrs = attrs;
                if !unit.marks.iter().any(|mark| mark.name == marker.name) {
                    unit.marks.push(marker.clone());
                }
                units.push(unit);
            }
        }
    }
    units
}

fn inline_container_units(
    child: &Value,
    style_formatting: Option<&Value>,
    styles: &StyleResolver,
    source: &BTreeMap<String, String>,
    opaque_sequences: &mut Vec<String>,
    in_control: bool,
) -> Vec<InlineUnit> {
    match string(field(Some(child), "type")).unwrap_or_default() {
        "run" => run_to_units(child, style_formatting, styles, &[], source),
        "hyperlink" => {
            opaque_sequences.extend(hyperlink_sequence_names(child));
            hyperlink_to_units(
                child,
                style_formatting,
                styles,
                &[],
                source,
                opaque_sequences,
            )
        }
        "simpleField" | "complexField" => {
            opaque_sequences.extend(nested_sequence_names(child));
            let (payload, marks) = field_payload(child, style_formatting, source);
            vec![embed_unit("field", payload, &marks, 1)]
        }
        "inlineSdt" => vec![embed_unit(
            "sdt",
            sdt_payload(child, style_formatting, styles, source, opaque_sequences),
            &[],
            2,
        )],
        "mathEquation" => vec![embed_unit("math", math_payload(child), &[], 1)],
        "insertion" | "deletion" | "moveFrom" | "moveTo" => tracked_to_units_in_control(
            child,
            style_formatting,
            styles,
            source,
            opaque_sequences,
            in_control,
        ),
        _ => Vec::new(),
    }
}

fn sdt_properties_attrs(properties: &Value, source: &BTreeMap<String, String>) -> JsonObject {
    map_from_value(json!({
        "sdtType": nullish(field(Some(properties), "sdtType")),
        "id": nullish(field(Some(properties), "id")),
        "alias": nullish(field(Some(properties), "alias")),
        "tag": nullish(field(Some(properties), "tag")),
        "lock": nullish(field(Some(properties), "lock")),
        "placeholder": nullish(field(Some(properties), "placeholder")),
        "showingPlaceholder": boolean(field(Some(properties), "showingPlaceholder")).unwrap_or(false),
        "dateFormat": nullish(field(Some(properties), "dateFormat")),
        "listItems": field(Some(properties), "listItems").map(|value| source_json(value, source)),
        "checked": nullish(field(Some(properties), "checked")),
        "dataBinding": field(Some(properties), "dataBinding").map(|value| source_json(value, source)),
        "multiLine": nullish(field(Some(properties), "multiLine")),
        "rawPropertiesXml": nullish(field(Some(properties), "rawPropertiesXml")),
        "rawEndPropertiesXml": nullish(field(Some(properties), "rawEndPropertiesXml"))
    }))
}

fn sdt_payload(
    sdt: &Value,
    style_formatting: Option<&Value>,
    styles: &StyleResolver,
    source: &BTreeMap<String, String>,
    opaque_sequences: &mut Vec<String>,
) -> JsonObject {
    let mut content = Vec::new();
    let append = |content: &mut Vec<Value>, unit: InlineUnit| match unit.content {
        UnitContent::Text(text) if text == "\t" => {
            content.push(json!({ "kind": "tab", "attrs": value_from_map(&unit.attrs) }));
        }
        UnitContent::Text(text) => {
            if let Some(previous) = content.last_mut()
                && string(field(Some(&*previous), "kind")) == Some("text")
                && field(Some(&*previous), "attrs") == Some(&value_from_map(&unit.attrs))
            {
                let previous_text = previous
                    .as_object_mut()
                    .and_then(|value| value.get_mut("text"))
                    .and_then(|value| value.as_str())
                    .unwrap_or_default()
                    .to_owned();
                previous.as_object_mut().unwrap().insert(
                    "text".to_owned(),
                    Value::String(format!("{previous_text}{text}")),
                );
            } else {
                content.push(json!({
                    "kind": "text",
                    "text": text,
                    "attrs": value_from_map(&unit.attrs)
                }));
            }
        }
        UnitContent::Embed { kind, payload } => {
            content.push(json!({
                "kind": kind,
                "payload": value_from_map(&payload),
                "attrs": value_from_map(&unit.attrs)
            }));
        }
    };
    for child in array(field(Some(sdt), "content")) {
        match string(field(Some(child), "type")).unwrap_or_default() {
            "run" => {
                for unit in run_to_units(child, style_formatting, styles, &[], source) {
                    append(&mut content, unit);
                }
            }
            "hyperlink" => {
                opaque_sequences.extend(hyperlink_sequence_names(child));
                for unit in hyperlink_to_units(
                    child,
                    style_formatting,
                    styles,
                    &[],
                    source,
                    opaque_sequences,
                ) {
                    append(&mut content, unit);
                }
            }
            "simpleField" | "complexField" => {
                opaque_sequences.extend(nested_sequence_names(child));
                let (payload, marks) = field_payload(child, style_formatting, source);
                append(&mut content, embed_unit("field", payload, &marks, 1));
            }
            "inlineSdt" => append(
                &mut content,
                embed_unit(
                    "sdt",
                    sdt_payload(child, style_formatting, styles, source, opaque_sequences),
                    &[],
                    1,
                ),
            ),
            "mathEquation" => append(
                &mut content,
                embed_unit("math", math_payload(child), &[], 1),
            ),
            "insertion" | "deletion" | "moveFrom" | "moveTo" => {
                for unit in tracked_to_units_in_control(
                    child,
                    style_formatting,
                    styles,
                    source,
                    opaque_sequences,
                    true,
                ) {
                    append(&mut content, unit);
                }
            }
            _ => {}
        }
    }
    let properties = field(Some(sdt), "properties").unwrap_or(&Value::Null);
    let mut result = sdt_properties_attrs(properties, source);
    result.insert(
        "propertiesJson".to_owned(),
        Value::String(source_json(properties, source)),
    );
    result.insert("content".to_owned(), Value::Array(content));
    result
}

fn paragraph_style_formatting(
    paragraph: &Value,
    styles: &StyleResolver,
    extra: Option<&Value>,
) -> Option<Arc<Value>> {
    let style_id = string(field(field(Some(paragraph), "formatting"), "styleId"));
    let style = styles
        .enabled
        .then(|| styles.resolve_paragraph_style(style_id).run.clone())
        .flatten();
    if object(extra).is_none() {
        return style;
    }
    merge_text_formatting(style.as_deref(), extra).map(Arc::new)
}

/// Note number marks carry no story unit, so the run boundary cache is the only
/// place a saved paragraph can learn they were there.
fn note_ref_mark_types(run: &Value) -> Vec<Value> {
    array(field(Some(run), "content"))
        .iter()
        .filter_map(
            |content| match string(field(Some(content), "type")).unwrap_or_default() {
                "footnoteRefMark" => Some(Value::String("footnote".to_owned())),
                "endnoteRefMark" => Some(Value::String("endnote".to_owned())),
                _ => None,
            },
        )
        .collect()
}

fn units_text(units: &[InlineUnit]) -> String {
    units
        .iter()
        .map(|unit| match &unit.content {
            UnitContent::Text(text) => text.clone(),
            UnitContent::Embed { payload, .. } => payload
                .get("footnoteRefId")
                .or_else(|| payload.get("endnoteRefId"))
                .map(js_string)
                .unwrap_or_default(),
        })
        .collect()
}

/// `w:br w:type="page"|"column"`, which the story carries as a block embed
/// beside the paragraph instead of as an inline unit.
fn flow_break_type(content: &Value) -> Option<&'static str> {
    if string(field(Some(content), "type")) != Some("break") {
        return None;
    }
    match string(field(Some(content), "breakType")) {
        Some("page") => Some("page"),
        Some("column") => Some("column"),
        _ => None,
    }
}

/// Where a run's flow breaks sit in its text. They occupy no story unit, so
/// the save projection rebuilds them from these offsets.
fn flow_break_offsets(run: &Value, source: &BTreeMap<String, String>) -> Vec<Value> {
    let contents = array(field(Some(run), "content"));
    if !contents.iter().any(|item| flow_break_type(item).is_some()) {
        return Vec::new();
    }
    let mut breaks = Vec::new();
    let mut offset = 0usize;
    for content in contents {
        if let Some(kind) = flow_break_type(content) {
            breaks.push(json!({ "offset": offset, "type": kind }));
            continue;
        }
        offset += units_text(&run_content_to_units(content, &[], source))
            .encode_utf16()
            .count();
    }
    breaks
}

fn run_boundary(
    run: &Value,
    units: &[InlineUnit],
    source: &BTreeMap<String, String>,
) -> Option<Value> {
    if units
        .iter()
        .any(|unit| matches!(&unit.content, UnitContent::Embed { kind, .. } if kind != "noteRef"))
    {
        return None;
    }
    let keys: Vec<_> = units.iter().map(|unit| marks_key(&unit.marks)).collect();
    if keys
        .first()
        .is_some_and(|first| keys.iter().any(|key| key != first))
    {
        return None;
    }
    let note_marks = note_ref_mark_types(run);
    let breaks = flow_break_offsets(run, source);
    let mut boundary = Map::new();
    boundary.insert("text".to_owned(), Value::String(units_text(units)));
    if !note_marks.is_empty() {
        boundary.insert("noteMarks".to_owned(), Value::Array(note_marks));
    }
    if !breaks.is_empty() {
        boundary.insert("breaks".to_owned(), Value::Array(breaks));
    }
    if let Some(key) = keys.first() {
        boundary.insert("marksKey".to_owned(), Value::String(key.clone()));
    }
    if let Some(formatting) = field(Some(run), "formatting") {
        boundary.insert("formatting".to_owned(), formatting.clone());
    }
    if let Some(changes) = field(Some(run), "propertyChanges") {
        boundary.insert("propertyChanges".to_owned(), changes.clone());
    }
    Some(Value::Object(boundary))
}

fn resolved_text_formatting(formatting: Option<&Value>, styles: &StyleResolver) -> Option<Value> {
    let style = formatting
        .and_then(|formatting| string(field(Some(formatting), "styleId")))
        .and_then(|style_id| styles.resolve_run_style(Some(style_id)));
    merge_text_formatting(style.as_deref(), formatting)
}

fn paragraph_attrs(
    paragraph: &Value,
    styles: &StyleResolver,
    units: &[InlineUnit],
    unit_counts: &[usize],
    run_boundaries: Option<Vec<Value>>,
) -> JsonObject {
    let formatting = field(Some(paragraph), "formatting");
    let style_id = string(field(formatting, "styleId"));
    let list = field(Some(paragraph), "listRendering");
    let direct_value = field(formatting, "indentFirstLine");
    let direct_nonzero = direct_value.filter(|value| number(Some(value)) != Some(0.0));
    let list_value = field(list, "indentFirstLine");
    let (selected_first, selected_hanging) = if let Some(value) = direct_nonzero {
        (Some(value), field(formatting, "hangingIndent"))
    } else if let Some(value) = list_value {
        (Some(value), field(list, "hangingIndent"))
    } else if let Some(value) = direct_value {
        (Some(value), field(formatting, "hangingIndent"))
    } else {
        (None, None)
    };
    let mut attrs = JsonObject::new();
    for (key, value) in [
        ("paraId", field(Some(paragraph), "paraId")),
        ("textId", field(Some(paragraph), "textId")),
        ("numPr", field(formatting, "numPr")),
        ("numPrFromStyle", field(formatting, "numPrFromStyle")),
        ("listNumFmt", field(list, "numFmt")),
        ("listIsBullet", field(list, "isBullet")),
        ("listMarker", field(list, "marker")),
        (
            "listMarkerHidden",
            field(list, "markerHidden").filter(|value| truthy(Some(value))),
        ),
        ("listMarkerBold", field(list, "markerBold")),
        ("listMarkerItalic", field(list, "markerItalic")),
        ("listMarkerColor", field(list, "markerColor")),
        (
            "listLevelNumFmts",
            field(list, "levelNumFmts").filter(|value| truthy(Some(value))),
        ),
        ("listAbstractNumId", field(list, "abstractNumId")),
        ("listStartOverride", field(list, "startOverride")),
        ("_originalFormatting", formatting),
    ] {
        if let Some(value) = value.filter(|value| !value.is_null()) {
            attrs.insert(key.to_owned(), drop_nulls(value.clone()));
        }
    }
    for (key, value) in [
        ("styleId", style_id),
        (
            "listMarkerFontFamily",
            string(field(list, "markerFontFamily")).filter(|value| !value.is_empty()),
        ),
        (
            "listMarkerSuffix",
            string(field(list, "markerSuffix")).filter(|value| !value.is_empty()),
        ),
    ] {
        if let Some(value) = value {
            attrs.insert(key.to_owned(), Value::String(value.to_owned()));
        }
    }
    if let Some(size) = number(field(list, "markerFontSize")).filter(|value| *value != 0.0) {
        attrs.insert("listMarkerFontSize".to_owned(), json!(size));
    }
    if styles.enabled {
        let style = styles.resolve_paragraph_style(style_id);
        let style_ppr_ref = style.paragraph.as_ref();
        for key in STYLE_FALLBACK_KEYS {
            attrs.insert(
                key.to_owned(),
                field(formatting, key)
                    .or_else(|| field(style_ppr_ref, key))
                    .cloned()
                    .unwrap_or(Value::Null),
            );
        }
        attrs.insert(
            "spacingExplicit".to_owned(),
            truthy(field(formatting, "spacingExplicit"))
                .then(|| field(formatting, "spacingExplicit").cloned())
                .flatten()
                .unwrap_or(Value::Null),
        );
        let numbering_removed = number(field(field(formatting, "numPr"), "numId")) == Some(0.0)
            && field(style_ppr_ref, "numPr").is_some()
            && number(field(field(style_ppr_ref, "numPr"), "numId")) != Some(0.0);
        attrs.insert(
            "indentLeft".to_owned(),
            field(formatting, "indentLeft")
                .or_else(|| field(list, "indentLeft"))
                .or_else(|| field(style_ppr_ref, "indentLeft"))
                .cloned()
                .unwrap_or(Value::Null),
        );
        attrs.insert(
            "indentFirstLine".to_owned(),
            if selected_first.is_some() {
                selected_first
            } else if numbering_removed {
                None
            } else {
                field(style_ppr_ref, "indentFirstLine")
            }
            .cloned()
            .unwrap_or(Value::Null),
        );
        attrs.insert(
            "hangingIndent".to_owned(),
            if selected_first.is_some() {
                selected_hanging
            } else if numbering_removed {
                None
            } else {
                field(style_ppr_ref, "hangingIndent")
            }
            .cloned()
            .unwrap_or(Value::Bool(false)),
        );
        let direct = resolved_text_formatting(field(formatting, "runProperties"), styles);
        attrs.insert(
            "defaultTextFormatting".to_owned(),
            merge_text_formatting(style.default_run.as_deref(), direct.as_ref())
                .unwrap_or(Value::Null),
        );
        if field(formatting, "numPr").is_none()
            && field(style_ppr_ref, "numPr").is_some()
            && number(field(field(style_ppr_ref, "numPr"), "numId")) != Some(0.0)
        {
            let num_pr = field(style_ppr_ref, "numPr").unwrap().clone();
            attrs.insert("numPr".to_owned(), num_pr.clone());
            attrs.insert("numPrFromStyle".to_owned(), num_pr);
        }
    } else {
        for key in [
            "alignment",
            "spaceBefore",
            "spaceAfter",
            "spaceBeforeLines",
            "spaceAfterLines",
            "beforeAutospacing",
            "afterAutospacing",
            "lineSpacing",
            "lineSpacingRule",
            "indentRight",
            "borders",
            "shading",
            "tabs",
            "pageBreakBefore",
            "keepNext",
            "keepLines",
            "widowControl",
            "snapToGrid",
            "autoSpaceDE",
            "autoSpaceDN",
            "outlineLevel",
            "bidi",
        ] {
            attrs.insert(
                key.to_owned(),
                field(formatting, key).cloned().unwrap_or(Value::Null),
            );
        }
        attrs.insert(
            "spacingExplicit".to_owned(),
            truthy(field(formatting, "spacingExplicit"))
                .then(|| field(formatting, "spacingExplicit").cloned())
                .flatten()
                .unwrap_or(Value::Null),
        );
        attrs.insert(
            "indentLeft".to_owned(),
            field(formatting, "indentLeft")
                .or_else(|| field(list, "indentLeft"))
                .cloned()
                .unwrap_or(Value::Null),
        );
        attrs.insert(
            "indentFirstLine".to_owned(),
            selected_first.cloned().unwrap_or(Value::Null),
        );
        attrs.insert(
            "hangingIndent".to_owned(),
            selected_hanging.cloned().unwrap_or(Value::Bool(false)),
        );
        attrs.insert(
            "defaultTextFormatting".to_owned(),
            field(formatting, "runProperties")
                .cloned()
                .unwrap_or(Value::Null),
        );
    }
    if let Some(section) = field(Some(paragraph), "sectionProperties") {
        attrs.insert("_sectionProperties".to_owned(), section.clone());
        if let Some(start @ ("nextPage" | "continuous" | "oddPage" | "evenPage" | "nextColumn")) =
            string(field(Some(section), "sectionStart"))
        {
            attrs.insert(
                "sectionBreakType".to_owned(),
                Value::String(start.to_owned()),
            );
        }
    }
    if truthy(field(Some(paragraph), "renderedPageBreakBefore")) {
        attrs.insert("renderedPageBreakBefore".to_owned(), Value::Bool(true));
    }
    if paragraph_starts_with_page_break(paragraph) {
        attrs.insert("pageBreakBeforeRun".to_owned(), Value::Bool(true));
    }
    for (source, target) in [("pPrIns", "pPrIns"), ("pPrDel", "pPrDel")] {
        if let Some(info) = field(Some(paragraph), source) {
            attrs.insert(
                target.to_owned(),
                json!({
                    "revisionId": nullish(field(Some(info), "id")),
                    "author": nullish(field(Some(info), "author")),
                    "date": nullish(field(Some(info), "date"))
                }),
            );
        }
    }
    if !array(field(Some(paragraph), "propertyChanges")).is_empty() {
        attrs.insert(
            "pPrChange".to_owned(),
            field(Some(paragraph), "propertyChanges").unwrap().clone(),
        );
    }
    let mut bookmarks = Vec::new();
    let mut unit_index = 0usize;
    let mut pm_offset = 0u32;
    for (content_index, content) in array(field(Some(paragraph), "content")).iter().enumerate() {
        match string(field(Some(content), "type")).unwrap_or_default() {
            "bookmarkStart" => {
                let mut bookmark = json!({
                    "id": nullish(field(Some(content), "id")),
                    "name": nullish(field(Some(content), "name")),
                    "kind": "start",
                    "offset": pm_offset
                });
                for key in ["colFirst", "colLast"] {
                    if let Some(value) = field(Some(content), key) {
                        bookmark
                            .as_object_mut()
                            .unwrap()
                            .insert(key.to_owned(), value.clone());
                    }
                }
                bookmarks.push(bookmark);
            }
            "bookmarkEnd" => bookmarks.push(json!({
                "id": nullish(field(Some(content), "id")),
                "kind": "end",
                "offset": pm_offset
            })),
            _ => {
                for _ in 0..unit_counts.get(content_index).copied().unwrap_or_default() {
                    pm_offset += units.get(unit_index).map(|unit| unit.pm_size).unwrap_or(0);
                    unit_index += 1;
                }
            }
        }
    }
    if !bookmarks.is_empty() {
        attrs.insert("bookmarks".to_owned(), Value::Array(bookmarks));
    }
    if let Some(boundaries) = run_boundaries.filter(|boundaries| !boundaries.is_empty()) {
        attrs.insert(
            "_originalRunBoundaries".to_owned(),
            Value::Array(boundaries),
        );
    }
    attrs
}

fn para_attrs_to_ppr(attrs: JsonObject) -> JsonObject {
    attrs
        .into_iter()
        .filter(|(key, value)| {
            ![
                "paraId",
                "textId",
                "renderedPageBreakBefore",
                "numPrFromStyle",
            ]
            .contains(&key.as_str())
                && !value.is_null()
        })
        .map(|(key, value)| {
            (
                match key.as_str() {
                    "styleId" => "pStyle".to_owned(),
                    "_sectionProperties" => "sectPr".to_owned(),
                    _ => key,
                },
                drop_nulls(value),
            )
        })
        .collect()
}

/// A paragraph content node seeding leaves out, before the unit at `unit`.
struct Omitted {
    unit: usize,
    element: String,
    /// Inside the content control whose embed is the unit at `unit`.
    in_control: bool,
}

/// Where a paragraph's page or column break sits in the source.
#[derive(Clone, Copy, Eq, PartialEq)]
enum BreakPlace {
    Paragraph,
    /// Inside the content control whose embed is the unit.
    Control,
    /// Inside a field, whose cached result carries it.
    Field,
}

/// A page or column break of a paragraph, listed in the order [`inline_tokens`] meets it.
struct FlowBreak {
    kind: BreakType,
    /// The index into the paragraph's units of the unit the break precedes.
    unit: usize,
    place: BreakPlace,
    revision: Option<Revision>,
    /// For a break inside a content control: its UTF-16 offset into the control's content,
    /// then into each nested control's.
    control_offset: Option<Vec<u32>>,
}

fn break_kind(content: &Value) -> Option<BreakType> {
    match flow_break_type(content)? {
        "column" => Some(BreakType::Column),
        _ => Some(BreakType::Page),
    }
}

fn tracked_revision(content: &Value) -> Revision {
    let info = field(Some(content), "info");
    let text = |key: &str| {
        field(info, key).and_then(|value| match value {
            Value::String(value) if !value.is_empty() => Some(value.clone()),
            Value::Number(number) => Some(crate::structured::source::number_text(number)),
            _ => None,
        })
    };
    Revision {
        kind: match string(field(Some(content), "type")).unwrap_or_default() {
            "insertion" => RevisionKind::Insertion,
            "moveTo" => RevisionKind::MoveTo,
            "moveFrom" => RevisionKind::MoveFrom,
            _ => RevisionKind::Deletion,
        },
        id: text("id"),
        author: text("author"),
        date: text("date"),
    }
}

/// The breaks of one run's content, placed after the units before them.
fn run_breaks(
    run: &Value,
    start: usize,
    place: BreakPlace,
    revision: Option<&Revision>,
    output: &mut Vec<FlowBreak>,
) -> usize {
    let mut offset = 0;
    for content in array(field(Some(run), "content")) {
        match break_kind(content) {
            Some(kind) => output.push(FlowBreak {
                kind,
                unit: start + offset,
                place,
                revision: revision.cloned(),
                control_offset: None,
            }),
            None => offset += run_content_unit_count(content),
        }
    }
    offset
}

fn units_width(units: &[InlineUnit]) -> u32 {
    units.iter().map(unit_width).sum()
}

/// Where each break inside an inline content control sits in the control's frozen content, in
/// [`inline_tokens`] order: UTF-16 offsets that count tabs and embeds as one, as
/// [`sdt_payload`] lowers the content, followed by the offsets into each nested control. Also
/// how many breaks [`inline_tokens`] finds in the control, read in the same single pass.
fn control_break_offsets(
    sdt: &Value,
    styles: &StyleResolver,
    source: &BTreeMap<String, String>,
) -> (Vec<Vec<u32>>, usize) {
    let mut offsets = Vec::new();
    let mut count = 0usize;
    let mut offset = 0u32;
    let run = |run: &Value, mut at: u32, offsets: &mut Vec<Vec<u32>>| {
        for item in array(field(Some(run), "content")) {
            match break_kind(item) {
                Some(_) => offsets.push(vec![at]),
                None => at += run_content_width(item),
            }
        }
        at
    };
    let breaks_below = |node: &Value| {
        let mut tokens = Vec::new();
        inline_tokens(std::slice::from_ref(node), &mut tokens);
        tokens
            .iter()
            .filter(|token| matches!(**token, "pageBreak" | "columnBreak"))
            .count()
    };
    for child in array(field(Some(sdt), "content")) {
        let kind = string(field(Some(child), "type")).unwrap_or_default();
        let below = if kind == "inlineSdt" {
            0
        } else {
            breaks_below(child)
        };
        count += below;
        match kind {
            "run" => offset = run(child, offset, &mut offsets),
            "hyperlink" => {
                let mut at = offset;
                for inner in array(field(Some(child), "children")) {
                    if string(field(Some(inner), "type")) == Some("run") {
                        at = run(inner, at, &mut offsets);
                    }
                }
                offset += units_width(&hyperlink_to_units(
                    child,
                    None,
                    styles,
                    &[],
                    source,
                    &mut Vec::new(),
                ));
            }
            "inlineSdt" => {
                let (nested, below) = control_break_offsets(child, styles, source);
                count += below;
                if nested.len() == below {
                    offsets.extend(
                        nested
                            .into_iter()
                            .map(|path| std::iter::once(offset).chain(path).collect::<Vec<u32>>()),
                    );
                } else {
                    offsets.extend(std::iter::repeat_n(vec![offset], below));
                }
                offset += 1;
            }
            "simpleField" | "complexField" | "mathEquation" => {
                offsets.extend(std::iter::repeat_n(vec![offset], below));
                offset += 1;
            }
            "insertion" | "deletion" | "moveFrom" | "moveTo" => {
                let (nested, _) = control_break_offsets(child, styles, source);
                offsets.extend(nested.into_iter().map(|mut path| {
                    if let Some(first) = path.first_mut() {
                        *first += offset;
                    }
                    path
                }));
                offset += units_width(&tracked_to_units_in_control(
                    child,
                    None,
                    styles,
                    source,
                    &mut Vec::new(),
                    true,
                ));
            }
            _ => {}
        }
    }
    (offsets, count)
}

/// The breaks of one paragraph content node, in [`inline_tokens`] order; `start` is the index of
/// the node's first unit. Breaks inside a control get their offsets into it when `positions`;
/// a control nested in another leaves that to the outermost.
fn content_breaks(
    content: &Value,
    start: usize,
    styles: &StyleResolver,
    source: &BTreeMap<String, String>,
    output: &mut Vec<FlowBreak>,
    positions: bool,
    in_control: bool,
) {
    let runs = |key: &str, output: &mut Vec<FlowBreak>| {
        let mut found = Vec::new();
        for child in array(field(Some(content), key)) {
            if string(field(Some(child), "type")) == Some("run") {
                run_breaks(child, start, BreakPlace::Field, None, &mut found);
            }
        }
        output.extend(found.into_iter().map(|found| FlowBreak {
            unit: start,
            ..found
        }));
    };
    match string(field(Some(content), "type")).unwrap_or_default() {
        "run" => {
            run_breaks(content, start, BreakPlace::Paragraph, None, output);
        }
        "hyperlink" => {
            let mut offset = 0;
            for child in array(field(Some(content), "children")) {
                if string(field(Some(child), "type")) == Some("run") {
                    offset +=
                        run_breaks(child, start + offset, BreakPlace::Paragraph, None, output);
                }
            }
        }
        "simpleField" => runs("content", output),
        "complexField" => {
            runs("fieldCode", output);
            runs("fieldResult", output);
        }
        "inlineSdt" => {
            let mut nested = Vec::new();
            for child in array(field(Some(content), "content")) {
                content_breaks(child, start, styles, source, &mut nested, false, true);
            }
            let offsets = positions
                .then(|| control_break_offsets(content, styles, source).0)
                .filter(|offsets| offsets.len() == nested.len());
            output.extend(
                nested
                    .into_iter()
                    .enumerate()
                    .map(|(index, found)| FlowBreak {
                        unit: start,
                        place: if found.place == BreakPlace::Field {
                            BreakPlace::Field
                        } else {
                            BreakPlace::Control
                        },
                        control_offset: offsets.as_ref().map(|offsets| offsets[index].clone()),
                        ..found
                    }),
            );
        }
        "insertion" | "deletion" | "moveFrom" | "moveTo" => {
            let revision = tracked_revision(content);
            let mut offset = 0;
            for child in array(field(Some(content), "content")) {
                if string(field(Some(child), "type")) == Some("run") {
                    offset += run_breaks(
                        child,
                        start + offset,
                        BreakPlace::Paragraph,
                        Some(&revision),
                        output,
                    );
                } else if in_control
                    || string(field(Some(child), "type")) == Some("inlineSdt")
                    || has_tracked_control(child, false, true)
                {
                    let first = output.len();
                    content_breaks(
                        child,
                        start + offset,
                        styles,
                        source,
                        output,
                        positions,
                        in_control,
                    );
                    for found in &mut output[first..] {
                        if found.revision.is_none() {
                            found.revision = Some(revision.clone());
                        }
                    }
                    offset += inline_container_units(
                        child,
                        None,
                        styles,
                        source,
                        &mut Vec::new(),
                        in_control,
                    )
                    .len();
                } else {
                    offset +=
                        hyperlink_to_units(child, None, styles, &[], source, &mut Vec::new()).len();
                }
            }
        }
        _ => {}
    }
}

fn drawing_element(kind: &str) -> String {
    match kind {
        "alternateContent" => "mc:AlternateContent".to_owned(),
        kind => format!("w:{kind}"),
    }
}

/// The unmodelled source nodes inside one paragraph content node, in source order.
fn unmodelled_nodes(content: &Value, output: &mut Vec<String>) {
    let children = |key: &str| array(field(Some(content), key));
    match string(field(Some(content), "type")).unwrap_or_default() {
        "run" => output.extend(
            children("content")
                .iter()
                .filter(|item| string(field(Some(item), "type")) == Some("opaqueDrawing"))
                .map(|item| drawing_element(string(field(Some(item), "kind")).unwrap_or_default())),
        ),
        "hyperlink" => {
            let nodes = field(Some(content), "structuredChildren")
                .or_else(|| field(Some(content), "children"));
            for child in array(nodes) {
                unmodelled_nodes(child, output);
            }
        }
        "insertion" | "deletion" | "moveFrom" | "moveTo" => {
            for child in children("content") {
                unmodelled_nodes(child, output);
            }
        }
        "simpleField" => {
            for child in children("content") {
                unmodelled_nodes(child, output);
            }
        }
        "complexField" => {
            for child in children("fieldResult") {
                unmodelled_nodes(child, output);
            }
        }
        "inlineSdt" => {
            for child in children("content") {
                match string(field(Some(child), "type")).unwrap_or_default() {
                    "run" | "hyperlink" | "simpleField" | "complexField" | "inlineSdt"
                    | "insertion" | "deletion" | "moveFrom" | "moveTo" => {
                        unmodelled_nodes(child, output)
                    }
                    "mathEquation" | "bookmarkStart" | "bookmarkEnd" | "commentRangeStart"
                    | "commentRangeEnd" => {}
                    "rawXml" => output.push(crate::structured::source::element_name(
                        string(field(Some(child), "xml")).unwrap_or_default(),
                    )),
                    other => output.push(other.to_owned()),
                }
            }
        }
        "rawXml" => output.push(crate::structured::source::element_name(
            string(field(Some(content), "xml")).unwrap_or_default(),
        )),
        _ => {}
    }
}

/// A paragraph's units and pilcrow properties, with the content seeding leaves out of them.
struct ParagraphUnits {
    units: Vec<InlineUnit>,
    comment_marks: Vec<CommentMark>,
    ppr: JsonObject,
    omitted: Vec<Omitted>,
    breaks: Vec<FlowBreak>,
    opaque_sequences: Vec<String>,
}

fn paragraph_units(
    paragraph: &Value,
    styles: &StyleResolver,
    extra_run_formatting: Option<&Value>,
    source: &BTreeMap<String, String>,
) -> ParagraphUnits {
    let mut units = Vec::new();
    let mut omitted = Vec::new();
    let mut breaks = Vec::new();
    let mut opaque_sequences = Vec::new();
    let mut comment_marks = Vec::new();
    let mut boundaries = Some(Vec::new());
    let mut unit_counts = Vec::new();
    let style_formatting = paragraph_style_formatting(paragraph, styles, extra_run_formatting);
    for content in array(field(Some(paragraph), "content")) {
        let start = units.len();
        match string(field(Some(content), "type")).unwrap_or_default() {
            "commentRangeStart" | "commentRangeEnd" => {
                if let Some(id) = field(Some(content), "id") {
                    comment_marks.push(CommentMark {
                        unit: start,
                        start: string(field(Some(content), "type")) == Some("commentRangeStart"),
                        id: js_string(id),
                    });
                }
            }
            "run" => {
                let run_units =
                    run_to_units(content, style_formatting.as_deref(), styles, &[], source);
                if let Some(run_boundaries) = &mut boundaries {
                    if let Some(boundary) = run_boundary(content, &run_units, source) {
                        run_boundaries.push(boundary);
                    } else {
                        boundaries = None;
                    }
                }
                units.extend(run_units);
            }
            "hyperlink" => {
                boundaries = None;
                opaque_sequences.extend(hyperlink_sequence_names(content));
                let linked = hyperlink_to_units(
                    content,
                    style_formatting.as_deref(),
                    styles,
                    &[],
                    source,
                    &mut opaque_sequences,
                );
                units.extend(linked);
            }
            "simpleField" | "complexField" => {
                boundaries = None;
                units.extend(field_to_units(
                    content,
                    style_formatting.as_deref(),
                    styles,
                    source,
                    unit_counts.len(),
                    &mut opaque_sequences,
                ));
            }
            "inlineSdt" => {
                boundaries = None;
                units.push(embed_unit(
                    "sdt",
                    sdt_payload(
                        content,
                        style_formatting.as_deref(),
                        styles,
                        source,
                        &mut opaque_sequences,
                    ),
                    &[],
                    2,
                ));
            }
            "insertion" | "deletion" | "moveFrom" | "moveTo" => {
                boundaries = None;
                units.extend(tracked_to_units(
                    content,
                    style_formatting.as_deref(),
                    styles,
                    source,
                    &mut opaque_sequences,
                ));
            }
            "mathEquation" => {
                boundaries = None;
                units.push(embed_unit("math", math_payload(content), &[], 1));
            }
            "bookmarkStart" | "bookmarkEnd" | "rawXml" => {}
            _ => boundaries = None,
        }
        content_breaks(content, start, styles, source, &mut breaks, true, false);
        let mut elements = Vec::new();
        unmodelled_nodes(content, &mut elements);
        if !elements.is_empty() {
            let kind = string(field(Some(content), "type")).unwrap_or_default();
            let (unit, in_control) = match kind {
                "inlineSdt" => (units.len() - 1, true),
                "run" => (start + run_prefix_units(content), false),
                "rawXml" => (start, false),
                _ => (units.len(), false),
            };
            omitted.extend(elements.into_iter().map(|element| Omitted {
                unit,
                element,
                in_control,
            }));
        }
        unit_counts.push(units.len() - start);
    }
    let attrs = paragraph_attrs(paragraph, styles, &units, &unit_counts, boundaries);
    ParagraphUnits {
        ppr: para_attrs_to_ppr(attrs),
        units,
        comment_marks,
        omitted,
        breaks,
        opaque_sequences,
    }
}

/// Units a run seeds before its first unmodelled drawing.
fn run_prefix_units(run: &Value) -> usize {
    array(field(Some(run), "content"))
        .iter()
        .take_while(|item| string(field(Some(item), "type")) != Some("opaqueDrawing"))
        .map(run_content_unit_count)
        .sum()
}

fn run_tokens(run: &Value, tokens: &mut Vec<&'static str>) {
    for content in array(field(Some(run), "content")) {
        if string(field(Some(content), "type")) == Some("break")
            && matches!(
                string(field(Some(content), "breakType")),
                Some("page" | "column")
            )
        {
            tokens.push(
                if string(field(Some(content), "breakType")) == Some("column") {
                    "columnBreak"
                } else {
                    "pageBreak"
                },
            );
        } else if string(field(Some(content), "type")) != Some("text")
            || !string(field(Some(content), "text"))
                .unwrap_or_default()
                .is_empty()
        {
            tokens.push("visible");
        }
    }
}

fn inline_tokens(content: &[Value], tokens: &mut Vec<&'static str>) {
    for item in content {
        match string(field(Some(item), "type")).unwrap_or_default() {
            "run" => run_tokens(item, tokens),
            "hyperlink" => {
                for child in array(field(Some(item), "children")) {
                    if string(field(Some(child), "type")) == Some("run") {
                        run_tokens(child, tokens);
                    }
                }
            }
            "simpleField" => {
                for child in array(field(Some(item), "content")) {
                    if string(field(Some(child), "type")) == Some("run") {
                        run_tokens(child, tokens);
                    }
                }
            }
            "complexField" => {
                for key in ["fieldCode", "fieldResult"] {
                    for child in array(field(Some(item), key)) {
                        run_tokens(child, tokens);
                    }
                }
            }
            "inlineSdt" => inline_tokens(array(field(Some(item), "content")), tokens),
            "insertion" | "deletion" | "moveFrom" | "moveTo" => {
                for child in array(field(Some(item), "content")) {
                    if string(field(Some(child), "type")) == Some("run") {
                        run_tokens(child, tokens);
                    }
                }
            }
            "mathEquation" => tokens.push("visible"),
            _ => {}
        }
    }
}

fn paragraph_starts_with_page_break(paragraph: &Value) -> bool {
    let mut tokens = Vec::new();
    inline_tokens(array(field(Some(paragraph), "content")), &mut tokens);
    tokens.first() == Some(&"pageBreak") && tokens.contains(&"visible")
}

fn paragraph_flow_breaks(paragraph: &Value) -> (Vec<&'static str>, Vec<&'static str>) {
    let mut tokens = Vec::new();
    inline_tokens(array(field(Some(paragraph), "content")), &mut tokens);
    if !tokens.contains(&"visible") {
        let split = tokens
            .iter()
            .rposition(|token| *token == "columnBreak")
            .map_or(0, |index| index + 1);
        return (tokens[..split].to_vec(), tokens[split..].to_vec());
    }
    let mut leading = None;
    let mut trailing = Vec::new();
    let mut visible = false;
    for token in tokens {
        if matches!(token, "pageBreak" | "columnBreak") {
            if visible || leading.is_some() {
                trailing.push(token);
            } else {
                leading = Some(token);
            }
        } else {
            visible = true;
        }
    }
    (
        leading
            .filter(|kind| *kind == "columnBreak")
            .into_iter()
            .collect(),
        trailing,
    )
}

fn modifier(value: &str) -> f64 {
    let prefix: String = value
        .chars()
        .take_while(|character| character.is_ascii_hexdigit())
        .collect();
    u8::from_str_radix(&prefix, 16)
        .map(|value| f64::from(value) / 255.0)
        .unwrap_or(1.0)
}

fn rgb_channels(value: &str) -> [u8; 3] {
    let mut normalized = value.trim_start_matches('#').to_owned();
    while normalized.len() < 6 {
        normalized.insert(0, '0');
    }
    normalized.truncate(6);
    [
        u8::from_str_radix(&normalized[0..2], 16).unwrap_or(0),
        u8::from_str_radix(&normalized[2..4], 16).unwrap_or(0),
        u8::from_str_radix(&normalized[4..6], 16).unwrap_or(0),
    ]
}

fn resolve_color_to_hex(color: Option<&Value>, theme: Option<&Value>) -> Option<String> {
    let color = object(color)?;
    if truthy(color.get("auto")) {
        return None;
    }
    if let Some(theme_color) = string(color.get("themeColor"))
        && let Some(theme) = theme
    {
        let slot = match theme_color {
            "dark1" | "text1" | "tx1" => "dk1",
            "light1" | "background1" | "bg1" => "lt1",
            "dark2" | "text2" | "tx2" => "dk2",
            "light2" | "background2" | "bg2" => "lt2",
            "hyperlink" => "hlink",
            "followedHyperlink" => "folHlink",
            value => value,
        };
        let known = [
            "dk1", "lt1", "dk2", "lt2", "accent1", "accent2", "accent3", "accent4", "accent5",
            "accent6", "hlink", "folHlink",
        ];
        let mut hex = if known.contains(&slot) {
            string(field(field(Some(theme), "colorScheme"), slot))
                .or_else(|| string(color.get("rgb")))
                .unwrap_or("000000")
                .to_owned()
        } else {
            string(color.get("rgb")).unwrap_or("000000").to_owned()
        };
        let mut channels = rgb_channels(&hex);
        if let Some(tint) = string(color.get("themeTint")) {
            let tint = modifier(tint);
            channels = channels
                .map(|channel| (f64::from(channel) * tint + 255.0 * (1.0 - tint)).round() as u8);
        } else if let Some(shade) = string(color.get("themeShade")) {
            let shade = modifier(shade);
            channels = channels.map(|channel| (f64::from(channel) * shade).round() as u8);
        }
        hex = format!("{:02X}{:02X}{:02X}", channels[0], channels[1], channels[2]);
        return Some(hex);
    }
    string(color.get("rgb"))
        .filter(|value| *value != "auto")
        .map(|value| value.trim_start_matches('#').to_ascii_uppercase())
}

fn calculate_row_spans(table: &Value) -> BTreeMap<(usize, usize), (usize, bool)> {
    let mut result = BTreeMap::new();
    let mut active = BTreeMap::<usize, usize>::new();
    for (row_index, row) in array(field(Some(table), "rows")).iter().enumerate() {
        let mut column = 0usize;
        let cells: Vec<_> = array(field(Some(row), "cells"))
            .iter()
            .map(|cell| {
                let current = column;
                column += number(field(field(Some(cell), "formatting"), "gridSpan")).unwrap_or(1.0)
                    as usize;
                (
                    current,
                    string(field(field(Some(cell), "formatting"), "vMerge")),
                )
            })
            .collect();
        let empty = !cells.is_empty()
            && cells
                .iter()
                .all(|(column, merge)| *merge == Some("continue") && active.contains_key(column));
        if empty {
            for (column, _) in cells {
                active.remove(&column);
                result.insert((row_index, column), (1, false));
            }
            continue;
        }
        for (column, merge) in cells {
            match merge {
                Some("restart") => {
                    active.insert(column, row_index);
                    result.insert((row_index, column), (1, false));
                }
                Some("continue") => {
                    if let Some(start) = active.get(&column).copied() {
                        if let Some(owner) = result.get_mut(&(start, column)) {
                            owner.0 += 1;
                        }
                        result.insert((row_index, column), (1, true));
                    } else {
                        result.insert((row_index, column), (1, false));
                    }
                }
                _ => {
                    active.remove(&column);
                    result.insert((row_index, column), (1, false));
                }
            }
        }
    }
    result
}

fn revision_attrs(info: &Value) -> Value {
    json!({
        "revisionId": nullish(field(Some(info), "id")),
        "author": nullish(field(Some(info), "author")),
        "date": nullish(field(Some(info), "date"))
    })
}

/// Maps each physical cell edge to the border side that feeds it: the matching
/// outer side where the cell sits on the boundary, `insideH`/`insideV` within.
pub(crate) fn border_side_sources(
    first_row: bool,
    last_row: bool,
    first_column: bool,
    last_column: bool,
) -> [(&'static str, &'static str); 4] {
    [
        ("top", if first_row { "top" } else { "insideH" }),
        ("bottom", if last_row { "bottom" } else { "insideH" }),
        ("left", if first_column { "left" } else { "insideV" }),
        ("right", if last_column { "right" } else { "insideV" }),
    ]
}

fn cell_borders(
    formatting: Option<&Value>,
    table_borders: Option<&Value>,
    first_row: bool,
    last_row: bool,
    first_column: bool,
    last_column: bool,
) -> Option<Value> {
    let inherited = object(table_borders).map(|borders| {
        Value::Object(
            border_side_sources(first_row, last_row, first_column, last_column)
                .into_iter()
                .map(|(edge, source)| (edge.to_owned(), nullish(borders.get(source))))
                .collect(),
        )
    });
    let direct = field(formatting, "borders");
    if inherited.is_none() && direct.is_none() {
        None
    } else {
        merge_plain(inherited.as_ref(), direct)
    }
}

struct CellOptions<'a> {
    is_header: bool,
    rowspan: usize,
    grid_width: Option<f64>,
    first_row: bool,
    last_row: bool,
    first_column: bool,
    last_column: bool,
    table_borders: Option<&'a Value>,
    default_margins: Option<&'a Value>,
    theme: Option<&'a Value>,
    table_bidi: bool,
}

fn structural_attrs(mut attrs: JsonObject, skipped: &[&str]) -> JsonObject {
    attrs.retain(|key, value| !skipped.contains(&key.as_str()) && !value.is_null());
    attrs.values_mut().for_each(drop_nulls_in_place);
    attrs
}

fn project_cell<'a>(cell: &'a Value, options: CellOptions<'_>) -> ProjectedCell<'a> {
    let formatting = field(Some(cell), "formatting");
    let background =
        resolve_color_to_hex(field(field(formatting, "shading"), "fill"), options.theme);
    let width = field(field(formatting, "width"), "value")
        .cloned()
        .or_else(|| options.grid_width.map(|value| json!(value)));
    let width_type = field(field(formatting, "width"), "type")
        .cloned()
        .or_else(|| options.grid_width.map(|_| Value::String("pct".to_owned())));
    let margins = if let Some(margins) = field(formatting, "margins") {
        Some(json!({
            "top": nullish(field(field(Some(margins), "top"), "value")),
            "bottom": nullish(field(field(Some(margins), "bottom"), "value")),
            "left": nullish(
                field(field(Some(margins), "left"), "value").or_else(|| {
                    field(
                        field(Some(margins), if options.table_bidi { "end" } else { "start" }),
                        "value",
                    )
                })
            ),
            "right": nullish(
                field(field(Some(margins), "right"), "value").or_else(|| {
                    field(
                        field(Some(margins), if options.table_bidi { "start" } else { "end" }),
                        "value",
                    )
                })
            )
        }))
    } else {
        options.default_margins.cloned()
    };
    let mut attrs = map_from_value(json!({
        "colspan": number(field(formatting, "gridSpan")).unwrap_or(1.0),
        "rowspan": options.rowspan,
        "width": width,
        "widthType": width_type,
        "verticalAlign": nullish(field(formatting, "verticalAlign")),
        "backgroundColor": background,
        "borders": cell_borders(
            formatting,
            options.table_borders,
            options.first_row,
            options.last_row,
            options.first_column,
            options.last_column,
        ),
        "margins": margins,
        "textDirection": nullish(field(formatting, "textDirection")),
        "noWrap": boolean(field(formatting, "noWrap")).unwrap_or(false),
        "_originalFormatting": nullish(formatting),
        "_originalResolvedFill": background
    }));
    if let Some(change) = field(Some(cell), "structuralChange") {
        let info = revision_attrs(field(Some(change), "info").unwrap_or(&Value::Null));
        match string(field(Some(change), "type")).unwrap_or_default() {
            "tableCellInsertion" => {
                attrs.insert(
                    "cellMarker".to_owned(),
                    json!({ "kind": "ins", "info": info }),
                );
            }
            "tableCellDeletion" => {
                attrs.insert(
                    "cellMarker".to_owned(),
                    json!({ "kind": "del", "info": info }),
                );
            }
            "tableCellMerge" => {
                let mut marker = json!({
                    "kind": "merge",
                    "info": info,
                    "vMerge": string(field(Some(change), "vMerge")).unwrap_or("cont")
                });
                if let Some(value) =
                    string(field(Some(change), "vMergeOrig")).filter(|value| !value.is_empty())
                {
                    marker
                        .as_object_mut()
                        .unwrap()
                        .insert("vMergeOrig".to_owned(), Value::String(value.to_owned()));
                }
                attrs.insert("cellMarker".to_owned(), marker);
            }
            _ => {}
        }
    }
    if !array(field(Some(cell), "propertyChanges")).is_empty() {
        attrs.insert(
            "tcPrChange".to_owned(),
            field(Some(cell), "propertyChanges").unwrap().clone(),
        );
    }
    let mut attrs = structural_attrs(attrs, &[]);
    if options.is_header {
        attrs.insert("header".to_owned(), Value::Bool(true));
    }
    let content = array(field(Some(cell), "content"));
    ProjectedCell {
        paragraph_formatting: None,
        attrs,
        content: if content.is_empty() {
            Cow::Owned(vec![json!({ "type": "paragraph", "content": [] })])
        } else {
            Cow::Borrowed(content)
        },
    }
}

struct TableStyleContext<'a> {
    column_count: usize,
    style: Option<&'a Value>,
    borders: Option<&'a Value>,
    margins: Option<&'a Value>,
    theme: Option<&'a Value>,
}

fn table_column_count(table: &Value) -> usize {
    array(field(Some(table), "rows"))
        .iter()
        .map(|row| {
            let row_formatting = field(Some(row), "formatting");
            let omitted = number(field(row_formatting, "gridBefore")).unwrap_or(0.0) as usize
                + number(field(row_formatting, "gridAfter")).unwrap_or(0.0) as usize;
            omitted
                + array(field(Some(row), "cells"))
                    .iter()
                    .map(|cell| {
                        number(field(field(Some(cell), "formatting"), "gridSpan")).unwrap_or(1.0)
                            as usize
                    })
                    .sum::<usize>()
        })
        .max()
        .unwrap_or(0)
        .max(array(field(Some(table), "columnWidths")).len())
}

fn table_cell_paragraph_formatting(
    table: &Value,
    style: Option<&Value>,
    row_index: usize,
    start_column: usize,
    end_column: usize,
    columns: usize,
) -> Option<Value> {
    let mut result = field(style, "pPr").cloned();
    let parts = array(field(style, "tblStylePr"));
    if !parts.iter().any(|part| field(Some(part), "pPr").is_some()) {
        return result;
    }
    let formatting = field(Some(table), "formatting");
    let style_formatting = field(style, "tblPr");
    let look = field(formatting, "look").or_else(|| field(style_formatting, "look"));
    let mask = string(field(look, "value"))
        .and_then(|value| u32::from_str_radix(value, 16).ok())
        .unwrap_or(if look.is_none() { 0x04a0 } else { 0 });
    let flag = |key, bit| boolean(field(look, key)).unwrap_or(mask & bit != 0);
    let first_row = flag("firstRow", 0x20);
    let last_row = flag("lastRow", 0x40);
    let first_column = flag("firstColumn", 0x80);
    let last_column = flag("lastColumn", 0x100);
    let rows = array(field(Some(table), "rows"));
    let grid_before = number(field(
        field(rows.get(row_index), "formatting"),
        "gridBefore",
    ))
    .unwrap_or(0.0) as usize;
    let start_column = start_column + grid_before;
    let end_column = end_column + grid_before;
    let at_first_row = first_row && row_index == 0;
    let at_last_row = last_row && row_index + 1 == rows.len();
    let at_first_column = first_column && start_column == 0;
    let at_last_column = last_column && end_column == columns;
    let band_size = |key| {
        number(field(formatting, key).or_else(|| field(style_formatting, key))).unwrap_or(1.0)
    };
    let row_band_size = band_size("styleRowBandSize");
    let column_band_size = band_size("styleColBandSize");
    let mut regions = Vec::new();
    if !flag("noHBand", 0x200) && !at_first_row && !at_last_row && row_band_size > 0.0 {
        let band =
            (row_index.saturating_sub(usize::from(first_row)) as f64 / row_band_size).floor();
        regions.push(if band % 2.0 == 0.0 {
            "band1Horz"
        } else {
            "band2Horz"
        });
    }
    if !flag("noVBand", 0x400) && !at_first_column && !at_last_column && column_band_size > 0.0 {
        let band = (start_column.saturating_sub(usize::from(first_column)) as f64
            / column_band_size)
            .floor();
        regions.push(if band % 2.0 == 0.0 {
            "band1Vert"
        } else {
            "band2Vert"
        });
    }
    for (region, active) in [
        ("firstCol", at_first_column),
        ("lastCol", at_last_column),
        ("firstRow", at_first_row),
        ("lastRow", at_last_row),
        ("nwCell", at_first_row && at_first_column),
        ("neCell", at_first_row && at_last_column),
        ("swCell", at_last_row && at_first_column),
        ("seCell", at_last_row && at_last_column),
    ] {
        if active {
            regions.push(region);
        }
    }
    for region in regions {
        let conditional = parts
            .iter()
            .find(|part| string(field(Some(part), "type")) == Some(region));
        result = merge_paragraph_formatting(result.as_ref(), field(conditional, "pPr"));
    }
    result
}

fn project_row<'a>(
    row: &'a Value,
    table: &Value,
    row_index: usize,
    row_spans: &BTreeMap<(usize, usize), (usize, bool)>,
    style_context: &TableStyleContext<'_>,
) -> ProjectedRow<'a> {
    let formatting = field(Some(row), "formatting");
    let mut attrs = map_from_value(json!({
        "height": nullish(field(field(formatting, "height"), "value")),
        "heightRule": nullish(field(formatting, "heightRule")),
        "isHeader": truthy(field(formatting, "header")),
        "_originalFormatting": nullish(formatting)
    }));
    if let Some(change) = field(Some(row), "structuralChange") {
        let value = revision_attrs(field(Some(change), "info").unwrap_or(&Value::Null));
        match string(field(Some(change), "type")).unwrap_or_default() {
            "tableRowInsertion" => {
                attrs.insert("trIns".to_owned(), value);
            }
            "tableRowDeletion" => {
                attrs.insert("trDel".to_owned(), value);
            }
            _ => {}
        }
    }
    if !array(field(Some(row), "propertyChanges")).is_empty() {
        attrs.insert(
            "trPrChange".to_owned(),
            field(Some(row), "propertyChanges").unwrap().clone(),
        );
    }
    let widths = array(field(Some(table), "columnWidths"));
    let total_width: f64 = widths.iter().filter_map(Value::as_f64).sum();
    let rows = array(field(Some(table), "rows"));
    let total_columns = if !widths.is_empty() {
        widths.len()
    } else {
        rows.iter()
            .map(|row| {
                array(field(Some(row), "cells"))
                    .iter()
                    .map(|cell| {
                        number(field(field(Some(cell), "formatting"), "gridSpan")).unwrap_or(1.0)
                            as usize
                    })
                    .sum()
            })
            .max()
            .unwrap_or(0)
    };
    let cells_source = array(field(Some(row), "cells"));
    let mut column = 0usize;
    let mut cells = Vec::new();
    for cell in cells_source {
        let colspan =
            number(field(field(Some(cell), "formatting"), "gridSpan")).unwrap_or(1.0) as usize;
        let start_column = column;
        let span = row_spans.get(&(row_index, start_column));
        let grid_width = (!widths.is_empty() && total_width > 0.0).then(|| {
            let cell_width: f64 = widths
                .iter()
                .skip(start_column)
                .take(colspan)
                .filter_map(Value::as_f64)
                .sum();
            (cell_width / total_width * 100.0).round()
        });
        column += colspan;
        if span.is_some_and(|(_, skip)| *skip) {
            continue;
        }
        let mut projected = project_cell(
            cell,
            CellOptions {
                is_header: row_index == 0
                    && truthy(field(
                        field(field(Some(table), "formatting"), "look"),
                        "firstRow",
                    )),
                rowspan: span.map(|(rowspan, _)| *rowspan).unwrap_or(1),
                grid_width,
                first_row: row_index == 0,
                last_row: row_index + 1 == rows.len(),
                first_column: start_column == 0,
                last_column: column == total_columns,
                table_borders: style_context.borders,
                default_margins: style_context.margins,
                theme: style_context.theme,
                table_bidi: truthy(field(field(Some(table), "formatting"), "bidi")),
            },
        );
        projected.paragraph_formatting = table_cell_paragraph_formatting(
            table,
            style_context.style,
            row_index,
            start_column,
            column,
            style_context.column_count,
        );
        cells.push(projected);
    }
    if cells.is_empty() {
        let synthetic = if total_columns > 1 {
            json!({
                "type": "tableCell",
                "formatting": { "gridSpan": total_columns },
                "content": [{ "type": "paragraph", "content": [] }]
            })
        } else {
            json!({
                "type": "tableCell",
                "content": [{ "type": "paragraph", "content": [] }]
            })
        };
        let projected = project_cell(
            &synthetic,
            CellOptions {
                is_header: row_index == 0
                    && truthy(field(
                        field(field(Some(table), "formatting"), "look"),
                        "firstRow",
                    )),
                rowspan: 1,
                grid_width: (total_width > 0.0).then_some(100.0),
                first_row: row_index == 0,
                last_row: row_index + 1 == rows.len(),
                first_column: true,
                last_column: true,
                table_borders: style_context.borders,
                default_margins: style_context.margins,
                theme: style_context.theme,
                table_bidi: truthy(field(field(Some(table), "formatting"), "bidi")),
            },
        );
        cells.push(ProjectedCell {
            paragraph_formatting: table_cell_paragraph_formatting(
                table,
                style_context.style,
                row_index,
                0,
                total_columns,
                style_context.column_count,
            ),
            attrs: projected.attrs,
            content: Cow::Owned(projected.content.into_owned()),
        });
    }
    ProjectedRow {
        attrs: structural_attrs(attrs, &[]),
        cells,
    }
}

fn project_table<'a>(
    table: &'a Value,
    styles: &StyleResolver,
    theme: Option<&Value>,
    compatibility_mode: u8,
) -> ProjectedTable<'a> {
    let formatting = field(Some(table), "formatting");
    let default_style = styles.default_style("table");
    let style_id = string(field(formatting, "styleId"));
    let effective_style_id =
        style_id.or_else(|| default_style.and_then(|style| string(field(Some(style), "styleId"))));
    let table_style = effective_style_id.and_then(|id| styles.style(id));
    let borders = field(formatting, "borders")
        .or_else(|| field(field(table_style, "tblPr"), "borders"))
        .or_else(|| field(field(default_style, "tblPr"), "borders"));
    let margin_layers: Vec<&Value> = [
        field(formatting, "cellMargins"),
        field(field(table_style, "tblPr"), "cellMargins")
            .filter(|margins| !margins.is_null())
            .or_else(|| field(field(default_style, "tblPr"), "cellMargins")),
    ]
    .into_iter()
    .flatten()
    .filter(|margins| !margins.is_null())
    .collect();
    let margin_side = |keys: &[&str]| {
        margin_layers.iter().find_map(|margins| {
            keys.iter().find_map(|key| {
                field(field(Some(*margins), key), "value").filter(|value| !value.is_null())
            })
        })
    };
    let (logical_left, logical_right) = if truthy(field(formatting, "bidi")) {
        ("end", "start")
    } else {
        ("start", "end")
    };
    let default_margins = (!margin_layers.is_empty()).then(|| {
        drop_nulls(json!({
            "top": nullish(margin_side(&["top"])),
            "bottom": nullish(margin_side(&["bottom"])),
            "left": nullish(margin_side(&["left", logical_left])),
            "right": nullish(margin_side(&["right", logical_right]))
        }))
    });
    let mut based_on = Vec::new();
    let mut visited = BTreeSet::new();
    let mut inherited = table_style;
    while let Some(parent_id) = inherited.and_then(|style| string(field(Some(style), "basedOn"))) {
        if based_on.len() >= 32 || !visited.insert(parent_id.to_owned()) {
            break;
        }
        based_on.insert(0, Value::String(parent_id.to_owned()));
        inherited = styles.style(parent_id);
    }
    let mut original_formatting = object(formatting).cloned().unwrap_or_default();
    original_formatting.insert(
        "styleCascade".to_owned(),
        drop_nulls(json!({
            "selectedStyleId": style_id.filter(|value| !value.is_empty()),
            "defaultStyleId": default_style
                .and_then(|style| string(field(Some(style), "styleId")))
                .filter(|value| !value.is_empty()),
            "basedOnStyleIds": (!based_on.is_empty()).then_some(based_on)
        })),
    );
    let mut attrs = map_from_value(json!({
        "styleId": style_id,
        "width": nullish(field(field(formatting, "width"), "value")),
        "widthType": nullish(field(field(formatting, "width"), "type")),
        "justification": nullish(field(formatting, "justification")),
        "columnWidths": nullish(field(Some(table), "columnWidths")),
        "tableLayout": nullish(field(formatting, "layout")),
        "floating": nullish(field(formatting, "floating")),
        "cellMargins": default_margins,
        "look": nullish(field(formatting, "look")),
        "bidi": truthy(field(formatting, "bidi")).then_some(true),
        "compatibilityMode": if compatibility_mode == 12 {
            Value::Null
        } else {
            json!(compatibility_mode as f64)
        },
        "_originalFormatting": Value::Object(original_formatting)
    }));
    if !array(field(Some(table), "propertyChanges")).is_empty() {
        attrs.insert(
            "tblPrChange".to_owned(),
            field(Some(table), "propertyChanges").unwrap().clone(),
        );
    }
    let row_spans = calculate_row_spans(table);
    let column_count = table_column_count(table);
    let rows = array(field(Some(table), "rows"))
        .iter()
        .enumerate()
        .map(|(row_index, row)| {
            project_row(
                row,
                table,
                row_index,
                &row_spans,
                &TableStyleContext {
                    column_count,
                    style: table_style.or(default_style),
                    borders,
                    margins: default_margins.as_ref(),
                    theme,
                },
            )
        })
        .collect();
    ProjectedTable { attrs, rows }
}

fn table_cell_story_id(parent: &str, table: usize, row: usize, cell: usize) -> String {
    format!("{parent}:t{table}:r{row}c{cell}")
}

/// Hands each block of a story the identity `visit_story` seeds it with.
#[derive(Clone, Copy, Default)]
struct BlockCursor {
    paragraph: usize,
    table: usize,
    sdt: usize,
}

impl BlockCursor {
    fn take(&mut self, story_id: &str, block: &Value) -> Option<String> {
        match string(field(Some(block), "type")).unwrap_or_default() {
            "rawXml" => None,
            "paragraph" => {
                let id = string(field(Some(block), "paraId"))
                    .filter(|value| {
                        !value.is_empty() && !truthy(field(Some(block), "repeatedParaId"))
                    })
                    .map_or_else(|| format!("{story_id}:p{}", self.paragraph), str::to_owned);
                self.paragraph += 1;
                Some(id)
            }
            "table" => {
                let id = format!("{story_id}:t{}", self.table);
                self.table += 1;
                Some(id)
            }
            _ => {
                let id = format!("{story_id}:sdt{}", self.sdt);
                self.sdt += 1;
                Some(id)
            }
        }
    }
}

/// How many story blocks a suppressed field's cached result duplicates.
fn cached_result_block_count(data: &Value) -> Option<usize> {
    let blocks = array(field(field(Some(data), "structuredResult"), "blocks"));
    blocks
        .last()
        .is_some_and(|block| {
            string(field(Some(block), "type")) == Some("paragraph")
                && array(field(Some(block), "content")).is_empty()
        })
        .then(|| blocks.len())
}

/// Binds each suppressed field to the story blocks its cached result duplicates.
fn bind_field_result_blocks(
    units: &mut [InlineUnit],
    story_id: &str,
    blocks: &[Value],
    owner: usize,
    after_owner: BlockCursor,
    table_ids: &mut BTreeMap<usize, String>,
) {
    for unit in units {
        let UnitContent::Embed { kind, payload } = &mut unit.content else {
            continue;
        };
        // `instruction` is the key the render side suppresses on.
        if kind.as_str() != "field"
            || !numeric_field_instruction(
                payload
                    .get("instruction")
                    .and_then(Value::as_str)
                    .unwrap_or_default(),
            )
        {
            continue;
        }
        let Some(count) = payload
            .get("fieldData")
            .and_then(Value::as_str)
            .and_then(|data| serde_json::from_str::<Value>(data).ok())
            .as_ref()
            .and_then(cached_result_block_count)
        else {
            continue;
        };
        let Some(duplicated) = blocks.get(owner + 1..owner + 1 + count) else {
            continue;
        };
        let mut cursor = after_owner;
        let mut ids = Vec::with_capacity(duplicated.len());
        for (offset, block) in duplicated.iter().enumerate() {
            let is_table = string(field(Some(block), "type")) == Some("table");
            let Some(id) = cursor.take(story_id, block) else {
                continue;
            };
            if is_table {
                table_ids.insert(owner + 1 + offset, id.clone());
            }
            ids.push(Value::String(id));
        }
        if ids.is_empty() {
            continue;
        }
        payload.insert("fieldResultBlocks".to_owned(), Value::Array(ids));
    }
}

/// Binds each field whose code runs past its paragraph's mark to the paragraph that paragraph
/// joins (`fieldCodeTarget`) and the paragraphs in between whose marks the code hides
/// (`fieldCodeMarks`). Word shows the field's paragraph and the target, the paragraph holding the
/// field's separator or, when the field ends within its code, its code's last paragraph, as one
/// paragraph. Tables among those blocks leave the field unbound.
fn bind_field_code_blocks(
    units: &mut [InlineUnit],
    story_id: &str,
    blocks: &[Value],
    owner: usize,
    after_owner: BlockCursor,
) {
    for unit in units {
        let UnitContent::Embed { kind, payload } = &mut unit.content else {
            continue;
        };
        if kind.as_str() != "field" {
            continue;
        }
        let Some(data) = payload
            .get("fieldData")
            .and_then(Value::as_str)
            .filter(|data| data.contains("\"blocks\""))
            .and_then(|data| serde_json::from_str::<Value>(data).ok())
        else {
            continue;
        };
        let code = array(field(field(Some(&data), "structuredCode"), "blocks")).len();
        let has_result = !array(field(field(Some(&data), "structuredResult"), "blocks")).is_empty();
        let hidden = match (code, has_result) {
            (0, true)
                if blocks
                    .get(owner + 1)
                    .is_some_and(opens_with_field_separator) =>
            {
                0
            }
            (0, _) => continue,
            (code, true) => code,
            (code, false) => code - 1,
        };
        let Some(joined) = blocks.get(owner + 1..owner + 1 + hidden + 1) else {
            continue;
        };
        if joined
            .iter()
            .any(|block| string(field(Some(block), "type")) != Some("paragraph"))
        {
            continue;
        }
        let mut cursor = after_owner;
        let mut ids: Vec<Value> = joined
            .iter()
            .filter_map(|block| cursor.take(story_id, block))
            .map(Value::String)
            .collect();
        let Some(target) = ids.pop() else {
            continue;
        };
        payload.insert("fieldCodeMarks".to_owned(), Value::Array(ids));
        payload.insert("fieldCodeTarget".to_owned(), target);
    }
}

/// Whether a paragraph holds, outside any field of its own, the separator of a field that began
/// before it.
fn opens_with_field_separator(block: &Value) -> bool {
    array(field(Some(block), "content")).iter().any(|content| {
        string(field(Some(content), "type")) == Some("run")
            && array(field(Some(content), "content")).iter().any(|item| {
                string(field(Some(item), "type")) == Some("fieldChar")
                    && string(field(Some(item), "charType")) == Some("separate")
            })
    })
}

fn add_comment_coverage(plan: &mut StoryPlan) {
    let add_range =
        |coverage: &mut Vec<(String, Vec<(u32, u32)>)>, id: &str, start: u32, end: u32| {
            if end <= start {
                return;
            }
            let ranges = &mut coverage.iter_mut().find(|(key, _)| key == id).unwrap().1;
            if let Some(previous) = ranges.last_mut()
                && previous.1 == start
            {
                previous.1 = end;
            } else {
                ranges.push((start, end));
            }
        };
    let mut offset = 0u32;
    let mut open: Vec<(String, u32, Option<u32>)> = Vec::new();
    let mut marks = plan.comment_marks.iter().peekable();
    for unit_index in 0..=plan.units.len() {
        while marks.peek().is_some_and(|mark| mark.unit == unit_index) {
            let mark = marks.next().unwrap();
            if mark.start {
                if open.iter().any(|(id, _, _)| id == &mark.id) {
                    continue;
                }
                if !plan.comment_coverage.iter().any(|(id, _)| id == &mark.id) {
                    plan.comment_coverage.push((mark.id.clone(), Vec::new()));
                }
                open.push((mark.id.clone(), offset, None));
            } else if let Some(index) = open.iter().position(|(id, _, _)| id == &mark.id) {
                let (id, start, _) = open.remove(index);
                add_range(&mut plan.comment_coverage, &id, start, offset);
            }
        }
        if let Some(unit) = plan.units.get(unit_index) {
            if matches!(&unit.content, UnitContent::Embed { kind, .. } if kind == "pilcrow") {
                for (_, _, paragraph_end) in &mut open {
                    paragraph_end.get_or_insert(offset);
                }
            }
            offset += unit_width(unit);
        }
    }
    for (id, start, paragraph_end) in open {
        add_range(
            &mut plan.comment_coverage,
            &id,
            start,
            paragraph_end.unwrap_or(offset),
        );
    }
    plan.comment_coverage
        .retain(|(_, ranges)| !ranges.is_empty());
}

/// For each row of `table`, the source index of every cell seeding keeps as a cell story.
fn source_cells(table: &Value) -> Vec<Vec<usize>> {
    let spans = calculate_row_spans(table);
    array(field(Some(table), "rows"))
        .iter()
        .enumerate()
        .map(|(row_index, row)| {
            let mut column = 0usize;
            array(field(Some(row), "cells"))
                .iter()
                .enumerate()
                .filter_map(|(index, cell)| {
                    let start = column;
                    column += number(field(field(Some(cell), "formatting"), "gridSpan"))
                        .unwrap_or(1.0) as usize;
                    (!spans
                        .get(&(row_index, start))
                        .is_some_and(|(_, skipped)| *skipped))
                    .then_some(index)
                })
                .collect()
        })
        .collect()
}

fn unit_width(unit: &InlineUnit) -> u32 {
    match &unit.content {
        UnitContent::Text(text) => utf16_len(text),
        UnitContent::Embed { .. } => 1,
    }
}

/// Whether a source block subtree holds any text or drawing.
fn has_content(value: &Value) -> bool {
    match value {
        Value::Object(object) => match string(object.get("type")) {
            Some("text") => string(object.get("text")).is_some_and(|text| !text.is_empty()),
            Some("drawing" | "shape" | "chart" | "opaqueDrawing" | "mathEquation") => true,
            _ => object.values().any(has_content),
        },
        Value::Array(values) => values.iter().any(has_content),
        _ => false,
    }
}

/// Where each source cell of `table` sits on its grid, with the cell story seeding made for it.
fn table_layout(table: &Value, story_id: &str, table_index: usize) -> TableLayout {
    let sources = source_cells(table);
    let count = |value: Option<&Value>, default: f64| {
        number(value)
            .filter(|value| value.is_finite())
            .unwrap_or(default)
            .clamp(0.0, f64::from(u16::MAX)) as u32
    };
    let mut grid_columns = array(field(Some(table), "columnWidths")).len() as u32;
    let rows = array(field(Some(table), "rows"))
        .iter()
        .enumerate()
        .map(|(row_index, row)| {
            let formatting = field(Some(row), "formatting");
            let grid_before = count(field(formatting, "gridBefore"), 0.0);
            let grid_after = count(field(formatting, "gridAfter"), 0.0);
            let mut column = grid_before;
            let cells = array(field(Some(row), "cells"))
                .iter()
                .enumerate()
                .map(|(index, cell)| {
                    let formatting = field(Some(cell), "formatting");
                    let span = count(field(formatting, "gridSpan"), 1.0).max(1);
                    let story = sources
                        .get(row_index)
                        .and_then(|kept| kept.binary_search(&index).ok())
                        .map(|projected| {
                            table_cell_story_id(story_id, table_index, row_index, projected)
                        });
                    let layout = CellLayout {
                        column,
                        span,
                        merge: match string(field(formatting, "vMerge")) {
                            Some("restart") => SourceMerge::Restart,
                            Some("continue") => SourceMerge::Continue,
                            _ => SourceMerge::None,
                        },
                        story,
                        content: has_content(field(Some(cell), "content").unwrap_or(&Value::Null)),
                    };
                    column = column.saturating_add(span);
                    layout
                })
                .collect();
            grid_columns = grid_columns.max(column.saturating_add(grid_after));
            RowLayout {
                grid_before,
                grid_after,
                cells,
            }
        })
        .collect();
    TableLayout { grid_columns, rows }
}

/// Records the revision identity of every move below `value`.
fn record_moves(value: &Value, moves: &mut HashSet<String>) {
    match value {
        Value::Object(object) => {
            if let Some(kind @ ("moveFrom" | "moveTo")) = string(object.get("type")) {
                let revision = tracked_revision(value);
                moves.insert(crate::structured::source::move_key(
                    kind == "moveTo",
                    revision.id.as_deref().unwrap_or_default(),
                    revision.author.as_deref().unwrap_or_default(),
                    revision.date.as_deref().unwrap_or_default(),
                ));
            }
            for child in object.values() {
                record_moves(child, moves);
            }
        }
        Value::Array(values) => {
            for child in values {
                record_moves(child, moves);
            }
        }
        _ => {}
    }
}

/// Records where a paragraph's page and column breaks sat, and which break embeds seeding moved
/// out of the paragraph stand for them. `embeds` are the story indices of those embeds, in order.
fn record_breaks(
    context: &mut LoweringContext,
    paragraph: &Value,
    story_id: &str,
    para_id: &str,
    breaks: Vec<FlowBreak>,
    embeds: Vec<u32>,
) {
    let mut tokens = Vec::new();
    inline_tokens(array(field(Some(paragraph), "content")), &mut tokens);
    let expected = tokens
        .iter()
        .filter(|token| matches!(**token, "pageBreak" | "columnBreak"))
        .count();
    if breaks.len() != expected || breaks.is_empty() {
        return;
    }
    let leading = tokens.first() == Some(&"pageBreak") && tokens.contains(&"visible");
    let mut embeds = embeds.into_iter();
    for (index, found) in breaks.into_iter().enumerate() {
        let witness = if index == 0 && leading {
            Witness::Leading
        } else {
            match embeds.next() {
                Some(unit) => {
                    context.provenance.relocated.push(Relocated {
                        pin: Pin::new(story_id, unit),
                        para_id: para_id.to_owned(),
                    });
                    Witness::Embed(context.provenance.relocated.len() - 1)
                }
                None => Witness::Invisible,
            }
        };
        if found.place == BreakPlace::Field {
            continue;
        }
        context.provenance.inline.push(InlineRecord {
            pin: Pin::new(story_id, found.unit as u32),
            para_id: para_id.to_owned(),
            in_control: found.place == BreakPlace::Control,
            control_offset: found.control_offset,
            content: InlineSource::Break {
                kind: found.kind,
                revision: found.revision,
            },
            witness,
        });
    }
}

fn visit_story(
    context: &mut LoweringContext,
    story_id: String,
    source_blocks: &[Value],
    options: StoryOptions,
) {
    let plan_index = context.plans.len();
    context.plans.push(StoryPlan {
        story_id: story_id.clone(),
        units: Vec::new(),
        comment_marks: Vec::new(),
        comment_coverage: Vec::new(),
        measured: (0, 0),
    });
    let empty_story;
    let source = !source_blocks.is_empty();
    let blocks = if source_blocks.is_empty() {
        empty_story = [json!({ "type": "paragraph", "content": [] })];
        &empty_story[..]
    } else {
        source_blocks
    };
    let mut cursor = BlockCursor::default();
    let mut result_table_ids = BTreeMap::new();
    let mut last_kind = None;
    let mut block_order = Vec::new();
    let mut source_order = Vec::new();
    for (block_index, block) in blocks.iter().enumerate() {
        let position = cursor;
        let Some(block_id) = cursor.take(&story_id, block) else {
            block_order.push(SourceBlock::Raw);
            source_order.push(None);
            let xml = string(field(Some(block), "xml")).unwrap_or_default();
            let raws = context
                .provenance
                .raw_blocks
                .entry(story_id.clone())
                .or_default();
            raws.push(crate::structured::source::element_name(xml));
            let index = raws.len() - 1;
            if let Some(steps) = context.locators.get(&story_id) {
                let mut steps = steps.clone();
                steps.push(Step::Block(block_index));
                context.provenance.raw_sources.push(RawSource {
                    story: story_id.clone(),
                    index,
                    steps,
                    xml: xml.to_owned(),
                });
            }
            continue;
        };
        let kind = string(field(Some(block), "type")).unwrap_or_default();
        source_order.push(Some(block_id.clone()));
        block_order.push(
            if kind == "paragraph"
                && string(field(Some(block), "paraId")).is_some_and(|id| !id.is_empty())
            {
                SourceBlock::Anchor(block_id.clone())
            } else {
                SourceBlock::Other
            },
        );
        match kind {
            "paragraph" => {
                context
                    .provenance
                    .paragraph_sources
                    .entry(story_id.clone())
                    .or_default()
                    .push((!source_blocks.is_empty()).then_some(block_index));
                if has_run_property_changes(field(Some(block), "content").unwrap_or(&Value::Null)) {
                    context
                        .source
                        .run_revisions
                        .insert((story_id.clone(), block_id.clone()));
                }
                record_moves(block, &mut context.provenance.moves);
                let (leading_breaks, trailing_breaks) = paragraph_flow_breaks(block);
                let mut embeds = Vec::new();
                if options.include_page_breaks {
                    for kind in leading_breaks {
                        embeds.push(context.plans[plan_index].width());
                        context.plans[plan_index].units.push(embed_unit(
                            kind,
                            JsonObject::new(),
                            &[],
                            1,
                        ));
                    }
                }
                let ParagraphUnits {
                    mut units,
                    comment_marks,
                    mut ppr,
                    omitted,
                    breaks,
                    opaque_sequences,
                } = paragraph_units(block, &context.styles, None, &context.source_json);
                context.opaque_sequences.extend(opaque_sequences);
                let base = context.plans[plan_index].width();
                let offsets: Vec<u32> = std::iter::once(0)
                    .chain(units.iter().scan(0, |width, unit| {
                        *width += unit_width(unit);
                        Some(*width)
                    }))
                    .collect();
                let end = offsets.last().copied().unwrap_or_default();
                for omission in omitted {
                    let at = offsets.get(omission.unit).copied().unwrap_or(end);
                    context.provenance.inline.push(InlineRecord {
                        pin: Pin::new(&story_id, base + at),
                        para_id: block_id.clone(),
                        in_control: omission.in_control,
                        control_offset: None,
                        content: InlineSource::Omitted {
                            element: omission.element,
                        },
                        witness: Witness::Invisible,
                    });
                }
                let source_para_id = string(field(Some(block), "paraId"))
                    .filter(|value| source && !value.is_empty())
                    .map(str::to_owned);
                if let Some(id) = &source_para_id {
                    ppr.insert(SOURCE_PARA_ID.to_owned(), Value::String(id.clone()));
                }
                if !source {
                    ppr.insert(PARA_ORIGIN.to_owned(), Value::String(SYNTHETIC.to_owned()));
                }
                ppr.insert("paraId".to_owned(), Value::String(block_id.clone()));
                context.paragraphs.push(SeededParagraph {
                    root: context.root.clone(),
                    key: block_id.clone(),
                    source_para_id,
                    ordinal: number(field(Some(block), SOURCE_ORDINAL))
                        .map(|ordinal| ordinal as u32),
                    source,
                });
                bind_field_result_blocks(
                    &mut units,
                    &story_id,
                    blocks,
                    block_index,
                    cursor,
                    &mut result_table_ids,
                );
                bind_field_code_blocks(&mut units, &story_id, blocks, block_index, cursor);
                let unit_base = context.plans[plan_index].units.len();
                context.plans[plan_index]
                    .comment_marks
                    .extend(comment_marks.into_iter().map(|mark| CommentMark {
                        unit: unit_base + mark.unit,
                        ..mark
                    }));
                context.plans[plan_index].units.extend(units);
                context.plans[plan_index]
                    .units
                    .push(embed_unit("pilcrow", ppr, &[], 1));
                if options.include_page_breaks {
                    for kind in trailing_breaks {
                        embeds.push(context.plans[plan_index].width());
                        context.plans[plan_index].units.push(embed_unit(
                            kind,
                            JsonObject::new(),
                            &[],
                            1,
                        ));
                    }
                }
                let breaks: Option<Vec<FlowBreak>> = breaks
                    .into_iter()
                    .map(|found| {
                        Some(FlowBreak {
                            unit: (base + offsets.get(found.unit)?) as usize,
                            ..found
                        })
                    })
                    .collect();
                record_breaks(
                    context,
                    block,
                    &story_id,
                    &block_id,
                    breaks.unwrap_or_default(),
                    embeds,
                );
                last_kind = Some("paragraph");
            }
            "table" => {
                let current_table = position.table;
                context.provenance.tables.insert(
                    format!("{story_id}:t{current_table}"),
                    table_layout(block, &story_id, current_table),
                );
                let ProjectedTable {
                    mut attrs,
                    mut rows,
                } = project_table(
                    block,
                    &context.styles,
                    context.theme.as_ref(),
                    context.compatibility_mode,
                );
                let payload_rows = rows
                    .iter_mut()
                    .enumerate()
                    .map(|(row_index, row)| {
                        let cells = row
                            .cells
                            .iter_mut()
                            .enumerate()
                            .map(|(cell_index, cell)| {
                                Value::Object(Map::from_iter([
                                    (
                                        "tcPr".to_owned(),
                                        Value::Object(
                                            std::mem::take(&mut cell.attrs).into_iter().collect(),
                                        ),
                                    ),
                                    (
                                        "story".to_owned(),
                                        Value::String(table_cell_story_id(
                                            &story_id,
                                            current_table,
                                            row_index,
                                            cell_index,
                                        )),
                                    ),
                                ]))
                            })
                            .collect();
                        Value::Object(Map::from_iter([
                            (
                                "trPr".to_owned(),
                                Value::Object(std::mem::take(&mut row.attrs).into_iter().collect()),
                            ),
                            ("cells".to_owned(), Value::Array(cells)),
                        ]))
                    })
                    .collect();
                let mut grid = match attrs.remove("columnWidths") {
                    Some(Value::Array(grid)) => grid,
                    _ => Vec::new(),
                };
                let tbl_pr = structural_attrs(attrs, &[]);
                grid.iter_mut().for_each(drop_nulls_in_place);
                let mut payload = JsonObject::from([
                    (
                        "tblPr".to_owned(),
                        Value::Object(tbl_pr.into_iter().collect()),
                    ),
                    ("grid".to_owned(), Value::Array(grid)),
                    ("rows".to_owned(), Value::Array(payload_rows)),
                ]);
                payload.values_mut().for_each(drop_nulls_in_place);
                if let Some(id) = result_table_ids.remove(&block_index) {
                    payload.insert("blockId".to_owned(), Value::String(id));
                }
                context.plans[plan_index]
                    .units
                    .push(embed_unit("table", payload, &[], 1));
                let previous_table_state = context.styles.set_table_paragraph_formatting(None);
                let sources = source_cells(block);
                for (row_index, row) in rows.into_iter().enumerate() {
                    for (cell_index, cell) in row.cells.into_iter().enumerate() {
                        context
                            .styles
                            .set_table_paragraph_formatting(cell.paragraph_formatting);
                        let source_cell = sources
                            .get(row_index)
                            .and_then(|cells| cells.get(cell_index))
                            .copied();
                        if let (Some(steps), Some(source_cell)) =
                            (context.locators.get(&story_id), source_cell)
                        {
                            let mut steps = steps.clone();
                            steps.extend([
                                Step::Block(block_index),
                                Step::Row(row_index),
                                Step::Cell(source_cell),
                            ]);
                            context.locators.insert(
                                table_cell_story_id(
                                    &story_id,
                                    current_table,
                                    row_index,
                                    cell_index,
                                ),
                                steps,
                            );
                        }
                        visit_story(
                            context,
                            table_cell_story_id(&story_id, current_table, row_index, cell_index),
                            &cell.content,
                            StoryOptions {
                                include_page_breaks: false,
                                append_body_tail: false,
                                seed_comments: options.seed_comments,
                            },
                        );
                    }
                }
                context
                    .styles
                    .restore_table_paragraph_formatting(previous_table_state);
                last_kind = Some("table");
            }
            _ => {
                let child_story = block_id;
                let mut properties = sdt_properties_attrs(
                    field(Some(block), "properties").unwrap_or(&Value::Null),
                    &context.source_json,
                );
                properties.insert("story".to_owned(), Value::String(child_story.clone()));
                context.plans[plan_index]
                    .units
                    .push(embed_unit("blockSdt", properties, &[], 1));
                if let Some(steps) = context.locators.get(&story_id) {
                    let mut steps = steps.clone();
                    steps.extend([Step::Block(block_index), Step::Content]);
                    context.locators.insert(child_story.clone(), steps);
                }
                visit_story(
                    context,
                    child_story,
                    array(field(Some(block), "content")),
                    StoryOptions {
                        include_page_breaks: options.include_page_breaks,
                        append_body_tail: false,
                        seed_comments: options.seed_comments,
                    },
                );
                last_kind = Some("blockSdt");
            }
        }
    }
    if block_order
        .iter()
        .any(|block| matches!(block, SourceBlock::Raw))
    {
        context.source.blocks.insert(story_id.clone(), block_order);
        context
            .provenance
            .block_order
            .insert(story_id.clone(), source_order);
    }

    if options.append_body_tail && matches!(last_kind, Some("table" | "blockSdt")) {
        context
            .provenance
            .paragraph_sources
            .entry(story_id.clone())
            .or_default()
            .push(None);
        let key = format!("{story_id}:p{}", cursor.paragraph);
        context.plans[plan_index].units.push(embed_unit(
            "pilcrow",
            map_from_value(json!({
                "hangingIndent": false,
                "paraId": key,
                (PARA_ORIGIN): SYNTHETIC
            })),
            &[],
            1,
        ));
        context.paragraphs.push(SeededParagraph {
            root: context.root.clone(),
            key,
            source_para_id: None,
            ordinal: None,
            source: false,
        });
    }
    if options.seed_comments {
        add_comment_coverage(&mut context.plans[plan_index]);
    }
}

fn collect_font_entry(key: &str, value: &Value, fonts: &mut BTreeSet<String>) {
    if matches!(
        key,
        "fontFamily" | "listMarkerFontFamily" | "markerFontFamily"
    ) {
        match value {
            Value::String(name) if !name.trim().is_empty() => {
                fonts.insert(name.trim().to_owned());
            }
            Value::Object(slots) => {
                for key in ["ascii", "hAnsi", "eastAsia", "cs"] {
                    if let Some(name) = slots
                        .get(key)
                        .and_then(Value::as_str)
                        .filter(|name| !name.trim().is_empty())
                    {
                        fonts.insert(name.trim().to_owned());
                    }
                }
            }
            _ => {}
        }
    }
}

fn collect_fonts_from_value(value: &Value, fonts: &mut BTreeSet<String>) {
    match value {
        Value::Array(values) => {
            for value in values {
                collect_fonts_from_value(value, fonts);
            }
        }
        Value::Object(values) => {
            for (key, value) in values {
                collect_font_entry(key, value, fonts);
                collect_fonts_from_value(value, fonts);
            }
        }
        _ => {}
    }
}

fn collect_font_table_fonts(envelope: &docx_parse::S9WireEnvelope, fonts: &mut BTreeSet<String>) {
    for font in &envelope.document.package.font_table.fonts {
        if !font.name.trim().is_empty() {
            fonts.insert(font.name.trim().to_owned());
        }
        if let Some(name) = font
            .alt_name
            .as_deref()
            .filter(|name| !name.trim().is_empty())
        {
            fonts.insert(name.trim().to_owned());
        }
    }
}

fn units_to_raw_ops(
    units: Vec<InlineUnit>,
    referenced_fonts: &mut BTreeSet<String>,
    mut script_fonts: Option<&mut ScriptFontUse>,
) -> Result<Vec<RawOp>, String> {
    let mut ops = vec![RawOp::Delete { index: 0, len: 1 }];
    let mut index = 0u32;
    let mut text = String::new();
    let mut attrs = JsonObject::new();
    let flush = |ops: &mut Vec<RawOp>,
                 index: &mut u32,
                 text: &mut String,
                 attrs: &mut JsonObject|
     -> Result<(), String> {
        if text.is_empty() {
            return Ok(());
        }
        let inserted = std::mem::take(text);
        let len = utf16_len(&inserted);
        ops.push(RawOp::Insert {
            index: *index,
            text: inserted,
            attrs: yrs_attrs(std::mem::take(attrs))?,
        });
        *index += len;
        Ok(())
    };
    for unit in units {
        if let Some(script_fonts) = script_fonts.as_deref_mut() {
            match &unit.content {
                UnitContent::Text(value) => script_fonts.text(value, &unit.attrs),
                UnitContent::Embed { payload, .. } => script_fonts.embed(payload, &unit.attrs),
            }
        }
        for (key, value) in &unit.attrs {
            collect_font_entry(key, value, referenced_fonts);
            collect_fonts_from_value(value, referenced_fonts);
        }
        if let UnitContent::Embed { payload, .. } = &unit.content {
            for (key, value) in payload {
                collect_font_entry(key, value, referenced_fonts);
                collect_fonts_from_value(value, referenced_fonts);
            }
        }
        match unit.content {
            UnitContent::Text(value) => {
                if text.is_empty() {
                    attrs = unit.attrs;
                } else if unit.attrs != attrs {
                    flush(&mut ops, &mut index, &mut text, &mut attrs)?;
                    attrs = unit.attrs;
                }
                text.push_str(&value);
            }
            UnitContent::Embed {
                kind,
                payload: values,
            } => {
                flush(&mut ops, &mut index, &mut text, &mut attrs)?;
                ops.push(RawOp::InsertEmbed {
                    index,
                    kind,
                    payload: payload(values)?,
                    attrs: yrs_attrs(unit.attrs)?,
                });
                index += 1;
            }
        }
    }
    flush(&mut ops, &mut index, &mut text, &mut attrs)?;
    Ok(ops)
}

fn seed_plan(
    plan: StoryPlan,
    script_fonts: Option<&mut ScriptFontUse>,
) -> Result<(String, Vec<RawOp>, BTreeSet<String>), String> {
    let StoryPlan {
        story_id,
        units,
        comment_coverage,
        ..
    } = plan;
    let mut referenced_fonts = BTreeSet::new();
    let mut ops = units_to_raw_ops(units, &mut referenced_fonts, script_fonts)?;
    if !comment_coverage.is_empty() {
        ops.extend(
            comment_coverage
                .into_iter()
                .map(|(id, ranges)| RawOp::SetComment {
                    id,
                    ranges,
                    author: String::new(),
                    date: String::new(),
                    body: Any::Null,
                }),
        );
    }
    Ok((story_id, ops, referenced_fonts))
}

fn entry_parts(entry: &Value) -> Option<(&str, &Value)> {
    let entry = entry.as_array()?;
    Some((entry.first()?.as_str()?, entry.get(1)?))
}

/// Parses a DOCX for editing, each paragraph carrying its source occurrence.
#[cfg(test)]
pub(crate) fn parse_docx_for_edit(bytes: &[u8]) -> Result<docx_parse::S9WireEnvelope, String> {
    parse_docx_package_with_digest(bytes, package_digest(bytes)).map(|(envelope, _)| envelope)
}

/// The SHA-256 of a whole package, hex encoded: the seed of its generated
/// IDs and the identity of its source index.
pub(crate) fn package_digest(bytes: &[u8]) -> String {
    use sha2::{Digest, Sha256};
    format!("{:x}", Sha256::digest(bytes))
}

/// `digest` as a package digest: the lowercase hex form [`package_digest`] gives.
#[cfg(feature = "wasm")]
pub(crate) fn checked_package_digest(digest: &str) -> Result<String, String> {
    if digest.len() == 64
        && digest
            .bytes()
            .all(|byte| matches!(byte, b'0'..=b'9' | b'a'..=b'f'))
    {
        Ok(digest.to_owned())
    } else {
        Err("a package digest is 64 lowercase hex digits".to_owned())
    }
}

/// Parses a DOCX for editing with the inflated parts the identity index
/// reads. `digest` is its [`package_digest`], so the parser does not hash the
/// package again.
#[cfg(any(test, feature = "wasm"))]
pub(crate) fn parse_docx_package_with_digest(
    bytes: &[u8],
    digest: String,
) -> Result<(docx_parse::S9WireEnvelope, Vec<(String, Vec<u8>)>), String> {
    docx_parse::parse_docx_s9_wire_parts_with_limits(
        bytes,
        docx_parse::S9ParseOptions {
            source_ordinals: true,
            determinism_seed: Some(digest),
            ..docx_parse::S9ParseOptions::default()
        },
        &docx_parse::xml::ParseLimits::default(),
    )
    .map_err(|error| error.to_string())
}

/// [`parse_docx_package_with_digest`] leaving the media in the package: an
/// image names its part by a `media:{n}` token of the returned table.
pub(crate) fn parse_docx_package_with_media(
    bytes: PackageBytes,
    digest: String,
) -> Result<
    (
        docx_parse::S9WireEnvelope,
        Vec<(String, Vec<u8>)>,
        docx_parse::media::MediaTable,
    ),
    String,
> {
    docx_parse::parse_docx_s9_wire_with_media_table_bytes(
        bytes,
        docx_parse::S9ParseOptions {
            source_ordinals: true,
            determinism_seed: Some(digest),
            ..docx_parse::S9ParseOptions::default()
        },
        &docx_parse::xml::ParseLimits::default(),
    )
    .map_err(|error| error.to_string())
}

/// [`parse_docx_for_edit`] plus the parts provenance resolves against.
pub(crate) fn parse_docx_with_parts(
    bytes: &[u8],
) -> Result<(docx_parse::S9WireEnvelope, SourceParts), String> {
    let (envelope, parts) = docx_parse::parse_docx_s9_wire_parts_with_limits(
        bytes,
        docx_parse::S9ParseOptions::default(),
        &docx_parse::ParseLimits::default(),
    )
    .map_err(|error| error.to_string())?;
    Ok((envelope, SourceParts::new(parts)))
}

type SourceRoot = (String, SourceStoryKind, Option<String>);

/// One package lowered into story plans, with what its identity index and structured reads
/// need.
struct LoweredDocx {
    context: LoweringContext,
    referenced_fonts: BTreeSet<String>,
    /// `None` for a seed whose caller does not report unused script fonts.
    script_fonts: Option<ScriptFontUse>,
    roots: Vec<SourceRoot>,
    relationships: Vec<(String, docx_parse::Relationship)>,
    read: ReadSource,
}

/// Lowers `envelope`, resolving source provenance against `parts` when the package's parts are
/// at hand.
fn lower_docx(
    envelope: docx_parse::S9WireEnvelope,
    parts: Option<&SourceParts>,
) -> Result<LoweredDocx, String> {
    lower_docx_with(envelope, parts, true)
}

/// [`lower_docx`], without the retained source JSON that only seeded payloads carry when
/// `payloads` is `false`.
fn lower_docx_with(
    mut envelope: docx_parse::S9WireEnvelope,
    parts: Option<&SourceParts>,
    payloads: bool,
) -> Result<LoweredDocx, String> {
    envelope.document.package.media_entries.clear();
    let relationships = envelope.document.package.relationship_entries.clone();
    let mut referenced_fonts = BTreeSet::new();
    collect_font_table_fonts(&envelope, &mut referenced_fonts);
    let mut script_fonts = ScriptFontUse::default();
    script_fonts.font_table(&envelope.document.package.font_table.fonts);
    let parsed = serde_json::to_value(&envelope.document).map_err(|error| error.to_string())?;
    collect_fonts_from_value(&parsed, &mut referenced_fonts);
    let source_json = if payloads && needs_source_json(&parsed) {
        let serialized =
            serde_json::to_string(&envelope.document).map_err(|error| error.to_string())?;
        let ordered: OrderedValue =
            serde_json::from_str(&serialized).map_err(|error| error.to_string())?;
        let mut values = BTreeMap::new();
        ordered.collect_source_json(&mut values);
        values
    } else {
        BTreeMap::new()
    };
    drop(envelope);
    let package =
        field(Some(&parsed), "package").ok_or_else(|| "parsed DOCX has no package".to_owned())?;
    let mut read = read_source(package, &parsed, parts);
    let (mut context, roots) = lower_package(package, source_json);
    drop(parsed);
    read.provenance = std::mem::take(&mut context.provenance);
    read.seeded_comments = seeded_comments(&context.plans);
    if let Some(parts) = parts {
        let comment_raw = comment_raw_sources(&context.styles, &read);
        let represented = represented_controls(&context.plans, &read);
        read.resolve_sources(parts, comment_raw, &represented);
    }
    Ok(LoweredDocx {
        context,
        referenced_fonts,
        script_fonts: Some(script_fonts),
        roots,
        relationships,
        read,
    })
}

/// Seeds `envelope` and retains its package context, resolving source provenance against
/// `parts` when the package's parts are at hand.
pub(crate) fn seed_parsed_docx_with(
    document: &EditingDoc,
    envelope: docx_parse::S9WireEnvelope,
    parts: Option<&SourceParts>,
) -> Result<Vec<String>, String> {
    let mut lowered = lower_docx(envelope, parts)?;
    lowered.script_fonts = None;
    seed_lowered(document, lowered, None, SeedMedia::AsParsed).map(|fonts| fonts.referenced)
}

/// The fonts a seeded document references, and those of them it names only
/// for East Asian or complex-script text it does not contain.
pub(crate) struct SeededFonts {
    pub(crate) referenced: Vec<String>,
    #[cfg_attr(not(feature = "wasm"), allow(dead_code))]
    pub(crate) unused_script: Vec<String>,
}

/// How seeding writes the images of a package parsed against a media table.
pub(crate) enum SeedMedia<'a> {
    /// As parsed: `media:{n}` tokens, which only a reader holding the
    /// package resolves.
    AsParsed,
    /// As their parts' `data:` URLs, which every replica reads; with
    /// `layout_tokens`, layout still carries the tokens.
    DataUrls {
        table: &'a docx_parse::media::MediaTable,
        layout_tokens: bool,
    },
}

/// Seeds every lowered story into `document` and retains the package context, with the identity
/// index when there is one.
fn seed_lowered(
    document: &EditingDoc,
    lowered: LoweredDocx,
    index: Option<SourceIndex>,
    media: SeedMedia<'_>,
) -> Result<SeededFonts, String> {
    let LoweredDocx {
        context,
        mut referenced_fonts,
        mut script_fonts,
        mut read,
        ..
    } = lowered;
    let scan_sequences = context.opaque_sequences.is_empty() && {
        let txn = document.yrs_doc().transact();
        txn.get_map(crate::STORIES)
            .expect("stories root is declared by EditingDoc::new")
            .len(&txn)
            == 0
            && txn
                .get_map(crate::identity::SESSION)
                .expect("session root is declared by EditingDoc::new")
                .get(&txn, OPAQUE_SEQUENCES)
                .is_none()
    };
    document
        .create_empty_stories(
            &context
                .plans
                .iter()
                .map(|plan| plan.story_id.clone())
                .collect::<Vec<_>>(),
        )
        .map_err(|error| error.to_string())?;
    let mut batches = Vec::with_capacity(context.plans.len());
    let mut deletes = Vec::with_capacity(context.plans.len());
    for plan in context.plans {
        let (story_id, mut ops, fonts) = seed_plan(plan, script_fonts.as_mut())?;
        deletes.push((story_id.clone(), vec![ops.remove(0)]));
        batches.push((story_id, ops));
        referenced_fonts.extend(fonts);
    }
    let sources = match media {
        SeedMedia::AsParsed => crate::media::MediaSources::default(),
        SeedMedia::DataUrls {
            table,
            layout_tokens,
        } => crate::media::write_data_urls(
            batches.iter_mut().flat_map(|(_, ops)| ops.iter_mut()),
            table,
            layout_tokens,
        )?,
    };
    let seeded = scan_sequences
        .then(|| seeded_sequence_fields(batches.iter().flat_map(|(_, ops)| ops)))
        .flatten();
    let ctx = EditCtx::local(String::new(), String::new());
    document
        .apply_raw_story_batches(deletes, &ctx)
        .map_err(|error| error.to_string())?;
    let ranges = document
        .apply_raw_seed_batches(batches, &ctx)
        .map_err(|error| error.to_string())?;
    seed_opaque_sequences(document, &context.opaque_sequences, seeded);
    document.set_media_sources(sources);
    read.pin(document, &ranges);
    read.comment_writes = CommentWrites::watch(document);
    if let Some(index) = index {
        document.retain_source(SourcePackage::Ready(Arc::new(index)));
    }
    document.install_source(
        SourceMetadata {
            styles: context.styles,
            structure: context.source,
            read,
        },
        0,
    );
    if let Some(source) = document.source_metadata()
        && !source.read().ambiguous_safety.is_empty()
    {
        let txn = document.yrs_doc().transact();
        if let Ok(inventory) = crate::content_controls::Inventory::build(document, &txn) {
            let _ = source
                .read()
                .embed_safety
                .set(inventory.occurrence_safety(source.read()));
        }
    }
    Ok(SeededFonts {
        unused_script: script_fonts
            .map(|scan| scan.unused(&referenced_fonts))
            .unwrap_or_default(),
        referenced: referenced_fonts.into_iter().collect(),
    })
}

/// Where the raw XML blocks of every source comment body sit in the comments part, as lowering
/// the body into a story of its own records them.
fn comment_raw_sources(styles: &StyleResolver, read: &ReadSource) -> Vec<RawSource> {
    let mut context = scratch_context(styles.clone());
    for comment in &read.comments {
        let story = format!("comment:{}", comment.id);
        context
            .locators
            .insert(story.clone(), vec![Step::Comment(comment.id.clone())]);
        visit_story(&mut context, story, &comment.body, scratch_options());
    }
    context.provenance.raw_sources
}

fn scratch_context(styles: StyleResolver) -> LoweringContext {
    LoweringContext {
        styles,
        theme: None,
        source_json: Arc::new(BTreeMap::new()),
        plans: Vec::new(),
        compatibility_mode: 12,
        root: String::new(),
        paragraphs: Vec::new(),
        opaque_sequences: Vec::new(),
        source: SourceStructure::default(),
        provenance: Provenance::default(),
        locators: HashMap::new(),
    }
}

fn scratch_options() -> StoryOptions {
    StoryOptions {
        include_page_breaks: false,
        append_body_tail: false,
        seed_comments: false,
    }
}

/// How many controls with each captured `w:sdtPr` seeding represents in each source part:
/// control embeds, the controls nested in them and block controls. A header or footer part that
/// several relationships reference counts once; every other story, each note included, counts.
fn represented_controls(plans: &[StoryPlan], read: &ReadSource) -> Represented {
    fn count(raw: Option<&Value>, part: &str, represented: &mut Represented) {
        let key = safety_key(string(raw));
        *represented.entry((part.to_owned(), key)).or_default() += 1;
    }
    fn nested(content: Option<&Value>, part: &str, represented: &mut Represented) {
        for item in array(content) {
            if string(field(Some(item), "kind")) == Some("sdt") {
                let payload = field(Some(item), "payload");
                count(field(payload, "rawPropertiesXml"), part, represented);
                nested(field(payload, "content"), part, represented);
            }
        }
    }
    let mut readers: HashMap<String, &str> = HashMap::new();
    let mut represented = Represented::new();
    for plan in plans {
        let root = story_root(&plan.story_id);
        let Some(part) = read.story_part(root) else {
            continue;
        };
        let alias = read
            .story(root)
            .is_some_and(|story| matches!(story.kind, StoryKind::Header | StoryKind::Footer));
        if alias && *readers.entry(part.clone()).or_insert(root) != root {
            continue;
        }
        for unit in &plan.units {
            let UnitContent::Embed { kind, payload } = &unit.content else {
                continue;
            };
            if matches!(kind.as_str(), "sdt" | "blockSdt") {
                count(payload.get("rawPropertiesXml"), &part, &mut represented);
                nested(payload.get("content"), &part, &mut represented);
            }
        }
    }
    represented
}

fn read_source(package: &Value, parsed: &Value, parts: Option<&SourceParts>) -> ReadSource {
    let document = parts.map_or(crate::structured::source::DOCUMENT_PART, |parts| {
        parts.document.as_str()
    });
    ReadSource::from_package(package, document, warnings(parsed))
}

fn seeded_comments(plans: &[StoryPlan]) -> HashSet<String> {
    plans
        .iter()
        .flat_map(|plan| plan.comment_coverage.iter().map(|(id, _)| id.clone()))
        .collect()
}

fn warnings(parsed: &Value) -> Vec<String> {
    array(field(Some(parsed), "warnings"))
        .iter()
        .filter_map(|warning| warning.as_str().map(str::to_owned))
        .collect()
}

/// Lowers each `(story, blocks)` into `doc` as its own story with `source`'s styles, for content
/// the session does not hold, such as comment bodies and field results. Returns what lowering
/// left out, pinned to `doc`.
pub(crate) fn seed_blocks(
    doc: &EditingDoc,
    source: Option<&SourceMetadata>,
    stories: &[(String, &[Value])],
) -> Result<Provenance, String> {
    let mut context = scratch_context(
        source
            .map(|source| source.styles.clone())
            .unwrap_or_default(),
    );
    for (story, blocks) in stories {
        visit_story(&mut context, story.clone(), blocks, scratch_options());
    }
    doc.create_empty_stories(
        &context
            .plans
            .iter()
            .map(|plan| plan.story_id.clone())
            .collect::<Vec<_>>(),
    )
    .map_err(|error| error.to_string())?;
    let mut provenance = std::mem::take(&mut context.provenance);
    let mut batches = Vec::with_capacity(context.plans.len());
    for plan in context.plans {
        let (story_id, ops, _) = seed_plan(plan, None)?;
        batches.push((story_id, ops));
    }
    let ranges = doc
        .apply_raw_seed_batches(batches, &EditCtx::local(String::new(), String::new()))
        .map_err(|error| error.to_string())?;
    provenance.pin(doc, &ranges);
    Ok(provenance)
}

/// Source metadata, identity index and referenced fonts for a package whose stories arrive
/// another way, such as shared state, from one lowering.
#[cfg(feature = "wasm")]
pub(crate) fn replica_source(
    envelope: docx_parse::S9WireEnvelope,
    parts: Vec<(String, Vec<u8>)>,
    bytes: PackageBytes,
    digest: String,
) -> Result<(SourceMetadata, SourceIndex, Vec<String>), String> {
    let ids = PackageIds::scan(&parts);
    let parts = SourceParts::new(parts);
    let LoweredDocx {
        mut context,
        referenced_fonts,
        roots,
        relationships,
        read,
        ..
    } = lower_docx_with(envelope, Some(&parts), false)?;
    let index = build_source_index(
        bytes,
        digest,
        &parts,
        ids,
        roots,
        &relationships,
        std::mem::take(&mut context.paragraphs),
    );
    Ok((
        SourceMetadata {
            styles: context.styles,
            structure: context.source,
            read,
        },
        index,
        referenced_fonts.into_iter().collect(),
    ))
}

fn lower_package(
    package: &Value,
    source_json: BTreeMap<String, String>,
) -> (LoweringContext, Vec<SourceRoot>) {
    let compatibility_mode = compatibility_mode_from_package(Some(package));
    let mut context = LoweringContext {
        styles: StyleResolver::new(field(Some(package), "styles")),
        theme: field(Some(package), "theme").cloned(),
        source_json: Arc::new(source_json),
        plans: Vec::new(),
        compatibility_mode,
        root: "body".to_owned(),
        paragraphs: Vec::new(),
        opaque_sequences: Vec::new(),
        source: SourceStructure::default(),
        provenance: Provenance::default(),
        locators: HashMap::from([("body".to_owned(), vec![Step::Body])]),
    };
    let mut roots = vec![("body".to_owned(), SourceStoryKind::Body, None)];
    visit_story(
        &mut context,
        "body".to_owned(),
        array(field(field(Some(package), "document"), "content")),
        StoryOptions {
            include_page_breaks: true,
            append_body_tail: true,
            seed_comments: true,
        },
    );
    for (key, kind) in [
        ("headerEntries", SourceStoryKind::Header),
        ("footerEntries", SourceStoryKind::Footer),
    ] {
        for entry in array(field(Some(package), key)) {
            let Some((relationship_id, part)) = entry_parts(entry) else {
                continue;
            };
            let story_id = format!("hf:{relationship_id}");
            if context.plans.iter().any(|plan| plan.story_id == story_id) {
                continue;
            }
            context.root = story_id.clone();
            context.locators.insert(story_id.clone(), Vec::new());
            roots.push((story_id.clone(), kind, Some(relationship_id.to_owned())));
            visit_story(
                &mut context,
                story_id,
                array(field(Some(part), "content")),
                StoryOptions {
                    include_page_breaks: false,
                    append_body_tail: false,
                    seed_comments: true,
                },
            );
        }
    }
    for (key, prefix, element, kind) in [
        ("footnotes", "fn", "footnote", SourceStoryKind::Footnote),
        ("endnotes", "en", "endnote", SourceStoryKind::Endnote),
    ] {
        for note in array(field(Some(package), key)) {
            let Some(id) = field(Some(note), "id") else {
                continue;
            };
            let story_id = format!("{prefix}:{}", js_string(id));
            context.root = story_id.clone();
            context
                .locators
                .insert(story_id.clone(), vec![Step::Note(element, js_string(id))]);
            roots.push((story_id.clone(), kind, Some(js_string(id))));
            visit_story(
                &mut context,
                story_id,
                array(field(Some(note), "content")),
                StoryOptions {
                    include_page_breaks: false,
                    append_body_tail: false,
                    seed_comments: true,
                },
            );
        }
    }
    (context, roots)
}

const COMMENT_COMPANIONS: [&str; 2] = ["word/commentsExtended.xml", "word/commentsIds.xml"];

/// What the identity index reads from every part of a package, before its parts narrow to the
/// stories: every paragraph ID an XML part uses, and the IDs the comment companion parts
/// reference.
struct PackageIds {
    occupied: BTreeSet<u32>,
    comment_references: BTreeSet<u32>,
}

impl PackageIds {
    fn scan(parts: &[(String, Vec<u8>)]) -> Self {
        Self {
            occupied: docx_parse::paragraph_identity::package_paragraph_ids(parts),
            comment_references: COMMENT_COMPANIONS
                .iter()
                .filter_map(|path| {
                    parts
                        .iter()
                        .find(|(candidate, _)| candidate.eq_ignore_ascii_case(path))
                })
                .flat_map(|(_, xml)| docx_parse::paragraph_identity::paragraph_id_attributes(xml))
                .collect(),
        }
    }
}

/// Indexes the package's paragraph identities: every story part with the
/// root stories seeded from it, and the IDs `ids` read from the whole package.
fn build_source_index(
    bytes: PackageBytes,
    digest: String,
    parts: &SourceParts,
    ids: PackageIds,
    roots: Vec<SourceRoot>,
    relationships: &[(String, docx_parse::Relationship)],
    paragraphs: Vec<SeededParagraph>,
) -> SourceIndex {
    use crate::structured::source::{COMMENTS_PART, ENDNOTES_PART, FOOTNOTES_PART};
    let document_path = &parts.document;
    let find = |path: &str| {
        parts
            .parts
            .iter()
            .find(|(candidate, _)| candidate.eq_ignore_ascii_case(path))
    };
    let mut inputs: Vec<SourcePartInput> = Vec::new();
    let mut add = |path: &str, kind: SourceStoryKind, root: Option<(String, Option<String>)>| {
        let Some((path, xml)) = find(path) else {
            return;
        };
        if let Some(input) = inputs.iter_mut().find(|input| input.path == *path) {
            input.roots.extend(root);
            return;
        }
        let Ok(xml) = std::str::from_utf8(xml) else {
            return;
        };
        inputs.push(SourcePartInput {
            path: path.clone(),
            kind,
            xml: xml.to_owned(),
            roots: root.into_iter().collect(),
        });
    };
    for (story_id, kind, item) in roots {
        let path = match kind {
            SourceStoryKind::Body => Some(document_path.clone()),
            SourceStoryKind::Footnote => Some(FOOTNOTES_PART.to_owned()),
            SourceStoryKind::Endnote => Some(ENDNOTES_PART.to_owned()),
            SourceStoryKind::Comment => Some(COMMENTS_PART.to_owned()),
            SourceStoryKind::Header | SourceStoryKind::Footer => relationships
                .iter()
                .find(|(id, _)| Some(id) == item.as_ref())
                .and_then(|(_, relationship)| {
                    match docx_parse::resolve_relationship_target(document_path, relationship) {
                        Ok(docx_parse::RelationshipTarget::Internal(path)) => Some(path),
                        _ => None,
                    }
                }),
        };
        let item =
            item.filter(|_| matches!(kind, SourceStoryKind::Footnote | SourceStoryKind::Endnote));
        if let Some(path) = path {
            add(&path, kind, Some((story_id, item)));
        }
    }
    add(FOOTNOTES_PART, SourceStoryKind::Footnote, None);
    add(ENDNOTES_PART, SourceStoryKind::Endnote, None);
    add(COMMENTS_PART, SourceStoryKind::Comment, None);
    SourceIndex::new(
        digest,
        bytes,
        ids.occupied,
        inputs,
        ids.comment_references,
        paragraphs,
    )
}

/// The identity index of a package, lowered without seeding.
pub(crate) fn source_index(
    bytes: PackageBytes,
    digest: Option<String>,
) -> Result<SourceIndex, String> {
    let digest = digest.unwrap_or_else(|| package_digest(&bytes));
    let (envelope, parts, _) = parse_docx_package_with_media(bytes.clone(), digest.clone())?;
    let ids = PackageIds::scan(&parts);
    let parts = SourceParts::new(parts);
    let lowered = lower_docx(envelope, None)?;
    Ok(build_source_index(
        bytes,
        digest,
        &parts,
        ids,
        lowered.roots,
        &lowered.relationships,
        lowered.context.paragraphs,
    ))
}

/// Seeds every story of the parsed package, resolving source provenance against `parts`, and
/// retains its identity index.
/// `digest` is the package's [`package_digest`].
pub(crate) fn seed_parsed_docx(
    document: &EditingDoc,
    envelope: docx_parse::S9WireEnvelope,
    parts: Vec<(String, Vec<u8>)>,
    bytes: PackageBytes,
    digest: String,
    media: SeedMedia<'_>,
) -> Result<SeededFonts, String> {
    let ids = PackageIds::scan(&parts);
    let parts = SourceParts::new(parts);
    let mut lowered = lower_docx(envelope, Some(&parts))?;
    let index = build_source_index(
        bytes,
        digest,
        &parts,
        ids,
        std::mem::take(&mut lowered.roots),
        &lowered.relationships,
        std::mem::take(&mut lowered.context.paragraphs),
    );
    seed_lowered(document, lowered, Some(index), media)
}

/// Opens a DOCX: seeds every story and starts a new opening with a fresh
/// generation, so its session anchors are its own; see
/// [`EditingDoc::begin_opening`].
pub fn seed_from_docx(document: &EditingDoc, bytes: &[u8]) -> Result<(), String> {
    seed_stories(document, bytes)?;
    document.begin_opening(None);
    Ok(())
}

/// [`seed_from_docx`] with a fixed opening generation, for a deterministic
/// seed every replica loads as one session.
pub fn seed_from_docx_with_generation(
    document: &EditingDoc,
    bytes: &[u8],
    generation: &str,
) -> Result<(), String> {
    seed_stories(document, bytes)?;
    document.begin_opening(Some(generation));
    Ok(())
}

/// Seeds every story of a DOCX without starting an opening.
/// Seeds `bytes` with only the body's first `blocks` blocks, parsed and
/// lowered, and the other stories they or the pages need, to paint the first
/// pages before the document is seeded in full. The result is for display
/// only. Its cut is not a real end of the document, so it is laid out with a
/// prefix pass that stops short of it
/// (`layout_document_with_regions_prefix_retained_json`); a pass that reaches
/// it needs more blocks. Seeds nothing and returns `false` for a document no
/// cut of which lays out like the whole (see
/// [`docx_parse::parse_docx_s9_preview_from_parts`]), which opens in full.
pub fn seed_docx_preview(
    document: &EditingDoc,
    bytes: &[u8],
    blocks: usize,
) -> Result<bool, String> {
    let Some((envelope, media, _)) = parse_docx_preview(bytes.into(), blocks, None)? else {
        return Ok(false);
    };
    seed_preview_envelope(document, envelope, media).map(|_| true)
}

/// The parse [`seed_docx_preview`] seeds from and the media its images name,
/// or `None` when it refuses one.
pub(crate) fn parse_docx_preview(
    bytes: PackageBytes,
    blocks: usize,
    paragraph_budget: Option<usize>,
) -> Result<
    Option<(
        docx_parse::S9WireEnvelope,
        docx_parse::media::MediaTable,
        bool,
    )>,
    String,
> {
    let (parts, media) =
        docx_parse::media_table_parts_bytes(&bytes).map_err(|error| error.to_string())?;
    let envelope = docx_parse::parse_docx_s9_preview_with_media_table_with_budget(
        &parts,
        &media,
        blocks,
        docx_parse::S9ParseOptions {
            source_ordinals: true,
            determinism_seed: Some(PREVIEW_SEED.to_owned()),
            ..docx_parse::S9ParseOptions::default()
        },
        &docx_parse::xml::ParseLimits::default(),
        paragraph_budget,
    )
    .map_err(|error| error.to_string())?;
    Ok(envelope.map(|(envelope, budget_stopped)| (envelope, media, budget_stopped)))
}

/// Seeds a preview parse and keeps the media its images name; returns the
/// fonts it references, and those of them it names only for East Asian or
/// complex-script text its cut does not contain.
pub(crate) fn seed_preview_envelope(
    document: &EditingDoc,
    envelope: docx_parse::S9WireEnvelope,
    media: docx_parse::media::MediaTable,
) -> Result<SeededFonts, String> {
    let mut lowered = lower_docx(envelope, None)?;
    retain_referenced_body_stories(&mut lowered.context.plans);
    let fonts = seed_lowered(document, lowered, None, SeedMedia::AsParsed)?;
    document.install_media(media);
    Ok(fonts)
}

/// The seed for IDs a preview's parse generates. A preview is never saved,
/// so it does not hash the package for the IDs a full open would generate.
const PREVIEW_SEED: &str = "0000000000000000000000000000000000000000000000000000000000000001";

/// Drops the body's nested stories, such as table cells, that the (cut) body
/// no longer reaches. Headers, footers, notes and comments stay whole.
fn retain_referenced_body_stories(plans: &mut Vec<StoryPlan>) {
    fn reach<'a>(value: &'a Value, found: &mut Vec<&'a str>) {
        match value {
            Value::String(text) => found.push(text),
            Value::Array(values) => values.iter().for_each(|value| reach(value, found)),
            Value::Object(fields) => fields.values().for_each(|value| reach(value, found)),
            _ => {}
        }
    }
    let nested = |id: &str| id.starts_with("body:");
    let index: HashMap<String, usize> = plans
        .iter()
        .enumerate()
        .map(|(position, plan)| (plan.story_id.clone(), position))
        .collect();
    let mut keep: HashSet<usize> = plans
        .iter()
        .enumerate()
        .filter(|(_, plan)| !nested(&plan.story_id))
        .map(|(position, _)| position)
        .collect();
    let mut queue: Vec<usize> = keep.iter().copied().collect();
    while let Some(position) = queue.pop() {
        let mut found = Vec::new();
        for unit in &plans[position].units {
            if let UnitContent::Embed { payload, .. } = &unit.content {
                payload.values().for_each(|value| reach(value, &mut found));
            }
        }
        for id in found {
            if let Some(&child) = index.get(id)
                && keep.insert(child)
            {
                queue.push(child);
            }
        }
    }
    let mut position = 0;
    plans.retain(|_| {
        position += 1;
        keep.contains(&(position - 1))
    });
}

/// Seeds `bytes` as the browser editor opens them: `data:` URLs in the
/// stories, tokens in layout.
#[cfg(test)]
pub(crate) fn seed_with_layout_tokens(document: &EditingDoc, bytes: &[u8]) -> Result<(), String> {
    let source = PackageBytes::from(bytes);
    let digest = package_digest(&source);
    let (envelope, parts, media) = parse_docx_package_with_media(source.clone(), digest.clone())?;
    seed_parsed_docx(
        document,
        envelope,
        parts,
        source,
        digest,
        SeedMedia::DataUrls {
            table: &media,
            layout_tokens: true,
        },
    )?;
    document.install_media(media);
    document.begin_opening(None);
    Ok(())
}

pub(crate) fn seed_stories(document: &EditingDoc, bytes: &[u8]) -> Result<(), String> {
    let digest = package_digest(bytes);
    let bytes = PackageBytes::from(bytes);
    let (envelope, parts, media) = parse_docx_package_with_media(bytes.clone(), digest.clone())?;
    seed_parsed_docx(
        document,
        envelope,
        parts,
        bytes,
        digest,
        SeedMedia::DataUrls {
            table: &media,
            layout_tokens: false,
        },
    )?;
    document.install_media(media);
    Ok(())
}

#[cfg(test)]
#[allow(dead_code)]
#[path = "../tests/support/structured_fixture.rs"]
pub(crate) mod fixture;

#[cfg(test)]
mod tests {
    fn assert_pins_match_story(doc: &EditingDoc, provenance: &Provenance) {
        let txn = doc.yrs_doc().transact();
        let pins = provenance
            .inline
            .iter()
            .map(|record| &record.pin)
            .chain(provenance.relocated.iter().map(|record| &record.pin));
        for pin in pins {
            assert_eq!(
                pin.position,
                Pin::sticky(&txn, &pin.story, pin.unit),
                "{}:{}",
                pin.story,
                pin.unit
            );
        }
    }

    #[test]
    fn seed_ranges_pin_every_story_unit_like_sticky_index() {
        let mut parts = fixture::principal_parts();
        for (path, bytes) in &mut parts {
            let paragraph = match path.as_str() {
                "word/document.xml" => Some(("10000006", "Inner")),
                "word/footnotes.xml" => Some(("20000004", "Footnote text")),
                "word/comments.xml" => Some(("30000001", "Please review")),
                _ => None,
            };
            if let Some((id, text)) = paragraph {
                let mut original = fixture::para(id, &fixture::run(text));
                if path.as_str() == "word/document.xml" {
                    original = original.replace(" xml:space=\"preserve\"", "");
                }
                let content = format!(
                    r#"{}<m:oMath><m:r><m:t>x=1</m:t></m:r></m:oMath><w:r><w:br w:type="page"/><w:br w:type="column"/></w:r>{}"#,
                    fixture::run("a😀b"),
                    fixture::run(text)
                );
                let xml = std::str::from_utf8(bytes).unwrap();
                assert!(xml.contains(&original));
                *bytes = xml
                    .replace(&original, &fixture::para(id, &content))
                    .into_bytes();
            }
        }
        let bytes = ooxml_opc::rezip_parts(&parts).unwrap();
        let LoweredDocx {
            context, mut read, ..
        } = lower_docx(parse_docx_for_edit(&bytes).unwrap(), None).unwrap();
        let doc = EditingDoc::new(74102);
        let stories: Vec<String> = context
            .plans
            .iter()
            .map(|plan| plan.story_id.clone())
            .collect();
        doc.create_empty_stories(&stories).unwrap();
        let batches = context
            .plans
            .into_iter()
            .map(|plan| {
                let (story, ops, _) = seed_plan(plan, None).unwrap();
                (story, ops)
            })
            .collect();
        let ranges = doc
            .apply_raw_seed_batches(batches, &EditCtx::local("", ""))
            .unwrap();
        assert_eq!(ranges.len(), stories.len());
        assert!(ranges.contains_key("fn:1"));
        assert!(read.provenance.tables.len() >= 2);
        assert!(!read.provenance.relocated.is_empty());
        assert!(
            read.provenance
                .inline
                .iter()
                .any(|record| matches!(record.content, InlineSource::Omitted { .. }))
        );
        assert!(
            read.provenance
                .inline
                .iter()
                .any(|record| matches!(record.content, InlineSource::Break { .. }))
        );
        for (story, range) in &ranges {
            for unit in 0..=range.len + 1 {
                read.provenance.relocated.push(Relocated {
                    pin: Pin::new(story, unit),
                    para_id: String::new(),
                });
            }
        }
        read.provenance.relocated.push(Relocated {
            pin: Pin::new("missing", 0),
            para_id: String::new(),
        });
        read.pin(&doc, &ranges);
        assert_pins_match_story(&doc, &read.provenance);
        let comments: Vec<_> = read
            .comments
            .iter()
            .map(|comment| (format!("comment:{}", comment.id), comment.body.as_slice()))
            .collect();
        let provenance = seed_blocks(&doc, None, &comments).unwrap();
        assert!(!provenance.inline.is_empty());
        assert_pins_match_story(&doc, &provenance);

        let seeded = EditingDoc::new(74103);
        crate::seed_from_docx(&seeded, &bytes).unwrap();
        let source = seeded.source_metadata().unwrap();
        assert_pins_match_story(&seeded, &source.read().provenance);
    }

    #[test]
    fn stories_outside_the_seed_contract_pin_with_the_slow_path() {
        let doc = EditingDoc::new(74104);
        doc.create_story("nonempty", "a😀b", "Normal", "left")
            .unwrap();
        doc.create_empty_stories(&["nonmonotonic".into(), "formatted".into(), "deleted".into()])
            .unwrap();
        let insert = |index, text: &str| RawOp::Insert {
            index,
            text: text.to_owned(),
            attrs: Attrs::new(),
        };
        let batches = vec![
            ("nonempty".into(), vec![insert(0, "x")]),
            (
                "nonmonotonic".into(),
                vec![
                    RawOp::Delete { index: 0, len: 1 },
                    insert(0, "a😀b"),
                    insert(0, "x"),
                ],
            ),
            (
                "formatted".into(),
                vec![
                    RawOp::Delete { index: 0, len: 1 },
                    insert(0, "a😀b"),
                    RawOp::Format {
                        index: 0,
                        len: 1,
                        attrs: Attrs::from([(Arc::from("bold"), Any::Bool(true))]),
                    },
                    insert(4, "x"),
                ],
            ),
            (
                "deleted".into(),
                vec![
                    RawOp::Delete { index: 0, len: 1 },
                    insert(0, "a😀b"),
                    RawOp::Delete { index: 0, len: 1 },
                ],
            ),
        ];
        let ranges = doc
            .apply_raw_seed_batches(batches, &EditCtx::local("", ""))
            .unwrap();
        assert!(ranges.is_empty());
        let mut provenance = Provenance::default();
        {
            let txn = doc.yrs_doc().transact();
            for story in [
                "nonempty",
                "nonmonotonic",
                "formatted",
                "deleted",
                "missing",
            ] {
                let len = crate::story_ref(&txn, story).map_or(0, |text| text.len(&txn));
                for unit in 0..=len + 1 {
                    provenance.relocated.push(Relocated {
                        pin: Pin::new(story, unit),
                        para_id: String::new(),
                    });
                }
            }
        }
        provenance.pin(&doc, &ranges);
        assert_pins_match_story(&doc, &provenance);
    }

    #[test]
    fn placeholder_delete_commit_preserves_single_transaction_seed_bytes() {
        let bytes = include_bytes!("../../../apps/demo/public/betteroffice-demo.docx");
        let mut lowered = lower_docx(parse_docx_for_edit(bytes).unwrap(), None).unwrap();
        let original = EditingDoc::new(1);
        original
            .create_empty_stories(
                &lowered
                    .context
                    .plans
                    .iter()
                    .map(|plan| plan.story_id.clone())
                    .collect::<Vec<_>>(),
            )
            .unwrap();
        let batches = lowered
            .context
            .plans
            .into_iter()
            .map(|plan| {
                let (story_id, ops, _) = seed_plan(plan, lowered.script_fonts.as_mut()).unwrap();
                (story_id, ops)
            })
            .collect();
        original
            .apply_raw_story_batches(batches, &EditCtx::local(String::new(), String::new()))
            .unwrap();
        seed_opaque_sequences(&original, &lowered.context.opaque_sequences, None);

        let split = EditingDoc::new(1);
        let lowered = lower_docx(parse_docx_for_edit(bytes).unwrap(), None).unwrap();
        seed_lowered(&split, lowered, None, SeedMedia::AsParsed).unwrap();
        assert_eq!(
            split.encode_state_as_update_v1(),
            original.encode_state_as_update_v1()
        );
    }

    #[test]
    fn opaque_sequence_names_accumulate_in_document_state() {
        let doc = EditingDoc::new(1);
        let empty = doc.encode_state_vector_v1();
        seed_opaque_sequences(&doc, &[], None);
        assert_eq!(doc.encode_state_vector_v1(), empty);
        seed_opaque_sequences(
            &doc,
            &["table".into(), "figure".into(), "table".into()],
            None,
        );
        {
            let txn = doc.yrs_doc().transact();
            assert_eq!(txn.state_vector().get(&doc.yrs_doc().client_id()), 0);
            assert_eq!(
                txn.state_vector()
                    .get(&yrs::ClientID::new(SEQUENCE_METADATA_CLIENT)),
                1
            );
        }
        seed_opaque_sequences(&doc, &["other".into(), "figure".into()], None);
        let before = doc.encode_state_vector_v1();
        seed_opaque_sequences(&doc, &[], None);
        seed_opaque_sequences(
            &doc,
            &["table".into(), "figure".into(), "other".into()],
            None,
        );
        assert_eq!(doc.encode_state_vector_v1(), before);
        doc.begin_opening(Some("opening"));
        let txn = doc.yrs_doc().transact();
        assert_eq!(
            txn.get_map(crate::identity::SESSION)
                .unwrap()
                .get(&txn, OPAQUE_SEQUENCES),
            Some(Out::Any(Any::Array(
                ["figure", "other", "table"].map(Any::from).to_vec().into()
            )))
        );
    }

    #[test]
    fn fresh_seeds_mark_sequence_strings_in_embed_values() {
        for (label, content, present) in [
            ("no sequence", fixture::run("Ordinary text"), false),
            (
                "field instruction",
                r#"<w:fldSimple w:instr=" SEQ Figure "><w:r><w:t>1</w:t></w:r></w:fldSimple>"#
                    .to_owned(),
                true,
            ),
            (
                "image description",
                fixture::image("rIdImage", "SEQUENCE"),
                true,
            ),
            (
                "plain text",
                format!(
                    r#"{}<w:hyperlink w:anchor="top">{}</w:hyperlink>"#,
                    fixture::run("SEQ"),
                    fixture::run("Link")
                ),
                false,
            ),
            ("run boundary payload", fixture::run("SEQ"), true),
        ] {
            let bytes = fixture::Package::new(&fixture::para("00000001", &content))
                .rel("rIdImage", "image", "media/image1.png")
                .bytes();
            let document = EditingDoc::new(1);
            crate::seed_from_docx(&document, &bytes).unwrap();
            let txn = document.yrs_doc().transact();
            assert_eq!(
                txn.get_map(crate::identity::SESSION)
                    .unwrap()
                    .get(&txn, OPAQUE_SEQUENCES),
                present.then(|| Out::Any(Any::Array(Vec::new().into()))),
                "{label}"
            );
        }
    }

    #[test]
    fn sequence_marker_reads_final_image_data_urls() {
        let mut parts = fixture::Package::new(&fixture::para(
            "00000001",
            &fixture::image("rIdImage", "Picture"),
        ))
        .rel("rIdImage", "image", "media/image1.png")
        .parts();
        parts
            .iter_mut()
            .find(|(path, _)| path == "word/media/image1.png")
            .unwrap()
            .1
            .extend_from_slice(b"HD\0");
        let bytes = ooxml_opc::rezip_parts(&parts).unwrap();
        let document = EditingDoc::new(1);
        crate::seed_from_docx(&document, &bytes).unwrap();
        let txn = document.yrs_doc().transact();
        assert_eq!(
            txn.get_map(crate::identity::SESSION)
                .unwrap()
                .get(&txn, OPAQUE_SEQUENCES),
            Some(Out::Any(Any::Array(Vec::new().into())))
        );
    }

    #[test]
    fn sequence_marker_keeps_existing_stories_in_the_walk() {
        let bytes =
            fixture::Package::new(&fixture::para("00000001", &fixture::run("Ordinary text")))
                .bytes();
        for (style, present) in [("Normal", false), ("SEQUENCE", true)] {
            let document = EditingDoc::new(1);
            document
                .create_story("existing", "", style, "left")
                .unwrap();
            crate::seed_from_docx(&document, &bytes).unwrap();
            let txn = document.yrs_doc().transact();
            assert_eq!(
                txn.get_map(crate::identity::SESSION)
                    .unwrap()
                    .get(&txn, OPAQUE_SEQUENCES),
                present.then(|| Out::Any(Any::Array(Vec::new().into()))),
                "{style}"
            );
        }
    }

    #[test]
    fn nested_sequence_names_keep_first_seen_order() {
        let field = json!({"type": "complexField", "instruction": "SEQ Outer",
        "structuredCode": {"inline": [
            {"type": "simpleField", "instruction": "sEq \"Figure\""},
            {"type": "complexField", "instruction": "SEQ Table"},
            {"type": "simpleField", "instruction": "SEQ FIGURE"}
        ]},
        "structuredResult": {"inline": [
            {"type": "complexField", "instruction": "SEQ TABLE"},
            {"type": "simpleField", "instruction": "SEQ Other"}
        ]}});
        assert_eq!(nested_sequence_names(&field), ["figure", "table", "other"]);
    }

    #[test]
    fn hyperlink_sequence_names_follow_inline_containers_in_order() {
        let hyperlink = json!({"type": "hyperlink", "children": [], "structuredChildren": [
            {"type": "run", "content": [{"type": "fieldChar", "charType": "begin"}]},
            {"type": "inlineSdt", "content": [
                {"type": "hyperlink", "children": [
                    {"type": "run", "content": [{"type": "instrText", "text": " sEq \"Fig"}]}
                ]},
                {"type": "simpleField", "instruction": "QUOTE", "content": [],
                    "structuredResult": {"inline": [
                        {"type": "run", "content": [{"type": "instrText", "text": "ure\" "}]}
                    ]}
                }
            ]},
            {"type": "complexField", "instruction": "QUOTE", "fieldCode": [], "fieldResult": [],
                "structuredCode": {"inline": [
                    {"type": "run", "content": [
                        {"type": "fieldChar", "charType": "begin"},
                        {"type": "instrText", "text": "SEQ Table"}
                    ]}
                ]},
                "structuredResult": {"inline": [
                    {"type": "run", "content": [{"type": "fieldChar", "charType": "end"}]}
                ]}
            },
            {"type": "run", "content": [
                {"type": "fieldChar", "charType": "separate"},
                {"type": "instrText", "text": "SEQ Ignored"},
                {"type": "fieldChar", "charType": "end"}
            ]}
        ]});
        assert_eq!(hyperlink_sequence_names(&hyperlink), ["table", "figure"]);
    }

    #[test]
    fn nested_sequence_metadata_covers_all_field_views_in_inline_wrappers() {
        let field = json!({
            "type": "complexField", "fieldType": "QUOTE", "instruction": "QUOTE",
            "fieldCode": [], "fieldResult": [],
            "structuredCode": {"inline": []}, "structuredResult": {"inline": []},
            "fieldTree": {"children": [{"result": {"inline": [
                {"type": "simpleField", "fieldType": "SEQ", "instruction": "SEQ Figure", "content": []}
            ]}}]}
        });
        for content in [
            field.clone(),
            json!({"type": "hyperlink", "children": [], "structuredChildren": [field.clone()]}),
            json!({"type": "inlineSdt", "properties": {}, "content": [field]}),
        ] {
            let paragraph = paragraph_units(
                &json!({"type": "paragraph", "content": [content]}),
                &StyleResolver::new(None),
                None,
                &BTreeMap::new(),
            );
            assert_eq!(paragraph.opaque_sequences, ["figure"]);
            for unit in paragraph.units {
                if let UnitContent::Embed { payload, .. } = unit.content {
                    assert!(
                        !serde_json::to_string(&payload)
                            .unwrap()
                            .contains("nestedSequences")
                    );
                }
            }
        }
    }

    fn seed_body(blocks: &[Value]) -> EditingDoc {
        let mut context = LoweringContext {
            styles: StyleResolver::new(None),
            theme: None,
            source_json: Arc::new(BTreeMap::new()),
            plans: Vec::new(),
            compatibility_mode: 12,
            root: "body".to_owned(),
            paragraphs: Vec::new(),
            opaque_sequences: Vec::new(),
            source: SourceStructure::default(),
            provenance: Provenance::default(),
            locators: HashMap::new(),
        };
        visit_story(
            &mut context,
            "body".to_owned(),
            blocks,
            StoryOptions {
                include_page_breaks: true,
                append_body_tail: false,
                seed_comments: false,
            },
        );
        let document = EditingDoc::new(74101);
        document
            .create_empty_stories(
                &context
                    .plans
                    .iter()
                    .map(|plan| plan.story_id.clone())
                    .collect::<Vec<_>>(),
            )
            .unwrap();
        let batches = context
            .plans
            .into_iter()
            .map(|plan| {
                let (story_id, ops, _) = seed_plan(plan, None).unwrap();
                (story_id, ops)
            })
            .collect();
        document
            .apply_raw_story_batches(batches, &EditCtx::local(String::new(), String::new()))
            .unwrap();
        document
    }

    fn rendered(document: &EditingDoc) -> (usize, String) {
        let blocks = crate::bridge::yrs_doc_to_layout_blocks(
            document,
            "body",
            &crate::bridge::RenderEnv::default(),
        )
        .unwrap();
        let json = serde_json::to_string(&blocks).unwrap();
        (blocks.len(), json)
    }

    fn run(text: &str) -> Value {
        json!({"type":"run","content":[{"type":"text","text":text}]})
    }

    fn block_field(instruction: &str, blocks: &[Value]) -> Value {
        json!({
            "type":"complexField", "fieldType":"UNKNOWN", "instruction":instruction,
            "fieldCode":[], "fieldResult":[run("Cached first")],
            "structuredResult":{
                "inline":[{"type":"hyperlink","anchor":"bookmark","children":[run("Cached first")]}],
                "blocks":blocks
            }
        })
    }

    #[test]
    fn numeric_fields_hide_cached_paragraphs_without_changing_the_story() {
        // `w14:paraId` is optional and plenty of real documents carry none, so
        // the binding cannot key on it.
        for (instruction, para_ids) in [("0", true), ("TOC", true), ("0", false), ("TOC", false)] {
            let id = |name: &str| if para_ids { json!(name) } else { json!(null) };
            let cached =
                json!({"type":"paragraph","paraId":id("cached"),"content":[run("Cached second") ]});
            let end = json!({"type":"paragraph","paraId":id("end"),"content":[]});
            let field = block_field(instruction, &[cached.clone(), end.clone()]);
            let document = seed_body(&[
                json!({"type":"paragraph","paraId":id("owner"),"content":[field]}),
                cached,
                end,
                json!({"type":"paragraph","paraId":id("after"),"content":[run("After")]}),
            ]);
            let before = crate::story_checksum(&document, "body").unwrap();
            let (count, output) = rendered(&document);
            let label = format!("{instruction} paraIds={para_ids}");
            assert_eq!(count, if instruction == "0" { 2 } else { 4 }, "{label}");
            assert_eq!(output.contains("Cached"), instruction != "0", "{label}");
            assert!(output.contains("After"), "{label}");
            assert_eq!(
                before,
                crate::story_checksum(&document, "body").unwrap(),
                "{label}"
            );
        }
    }

    #[test]
    fn a_numeric_field_the_preview_leaves_out_releases_its_cached_blocks() {
        let cached = json!({"type":"paragraph","paraId":"cached","content":[run("Cached second")]});
        let end = json!({"type":"paragraph","paraId":"end","content":[]});
        let field = block_field("0", &[cached.clone(), end.clone()]);
        let document = seed_body(&[
            json!({"type":"paragraph","paraId":"owner","content":[field]}),
            cached,
            end,
        ]);
        let stamp = Any::Map(Arc::new(HashMap::from([
            ("id".to_owned(), Any::from("9")),
            ("author".to_owned(), Any::from("Ann")),
            ("date".to_owned(), Any::from("2026-09-29T12:00:00Z")),
        ])));
        document
            .apply_raw_ops(
                "body",
                vec![crate::RawOp::Format {
                    index: 0,
                    len: 1,
                    attrs: Attrs::from([(Arc::from("ins"), stamp)]),
                }],
                &crate::EditCtx::local("", ""),
            )
            .unwrap();
        let lower = |env: crate::bridge::RenderEnv| {
            serde_json::to_string(
                &crate::bridge::yrs_doc_to_layout_blocks(&document, "body", &env).unwrap(),
            )
            .unwrap()
        };
        let preview = |decision| {
            lower(crate::bridge::RenderEnv::default().with_revision_preview("9", decision))
        };
        assert!(!lower(crate::bridge::RenderEnv::default()).contains("Cached"));
        assert!(!preview(crate::bridge::RevisionPreview::Accepted).contains("Cached"));
        assert!(preview(crate::bridge::RevisionPreview::Rejected).contains("Cached second"));
    }

    #[test]
    fn numeric_fields_hide_non_paragraph_cached_blocks() {
        for instruction in ["0", "TOC"] {
            let cell = json!({"type":"tableCell","content":[
                json!({"type":"paragraph","content":[run("Cached cell")]})
            ]});
            let table = json!({"type":"table","rows":[{"type":"tableRow","cells":[cell]}]});
            let sdt = json!({"type":"blockSdt","properties":{},"content":[
                json!({"type":"paragraph","content":[run("Cached sdt")]})
            ]});
            let end = json!({"type":"paragraph","content":[]});
            let field = block_field(instruction, &[table.clone(), sdt.clone(), end.clone()]);
            let document = seed_body(&[
                json!({"type":"paragraph","content":[field]}),
                table,
                sdt,
                end,
                json!({"type":"paragraph","content":[run("After")]}),
            ]);
            let before = crate::story_checksum(&document, "body").unwrap();
            let (count, output) = rendered(&document);
            assert_eq!(
                count,
                if instruction == "0" { 2 } else { 5 },
                "{instruction}"
            );
            assert_eq!(
                output.contains("Cached"),
                instruction != "0",
                "{instruction}"
            );
            assert!(output.contains("After"), "{instruction}");
            assert_eq!(
                before,
                crate::story_checksum(&document, "body").unwrap(),
                "{instruction}"
            );
        }
    }

    #[test]
    fn splitting_a_cached_paragraph_shows_the_authored_half_only() {
        let cached = json!({"type":"paragraph","content":[run("Cached second")]});
        let end = json!({"type":"paragraph","content":[]});
        let document = seed_body(&[
            json!({"type":"paragraph","content":[block_field("0", &[cached.clone(), end.clone()])]}),
            cached,
            end,
            json!({"type":"paragraph","content":[run("After")]}),
        ]);
        assert_eq!(rendered(&document).0, 2);
        // The field embed, the owner's pilcrow, then the cached text: pressing
        // Enter at the end of the cached paragraph and typing into the new one.
        let cached_pilcrow = 2 + "Cached second".encode_utf16().count() as u32;
        let ctx = EditCtx::local("Ada".to_owned(), "2026-01-01T00:00:00Z".to_owned());
        document
            .split_paragraph(&ctx, crate::Position::new("body", cached_pilcrow), None)
            .unwrap();
        document
            .insert_text(
                &ctx,
                crate::Position::new("body", cached_pilcrow + 1),
                "Authored",
                crate::FormatPolicy::Plain,
            )
            .unwrap();
        let (count, output) = rendered(&document);
        assert_eq!(count, 3);
        assert!(output.contains("Authored"));
        assert!(output.contains("After"));
        assert!(!output.contains("Cached"));
    }

    #[test]
    fn numeric_field_detection_preserves_formulas_and_named_fields() {
        assert!(numeric_field_instruction(" 123 "));
        for instruction in ["", "= 0", "QUOTE 0", "PAGE", "CustomField", "123abc"] {
            assert!(!numeric_field_instruction(instruction));
        }
    }

    #[test]
    fn complex_field_results_keep_hyperlink_units_and_style() {
        let link = json!({"type":"hyperlink","anchor":"_Toc1","children":[{"type":"run","formatting":{"styleId":"Hyperlink"},"content":[{"type":"text","text":"Heading"}]}]});
        let value = json!({"type":"complexField","fieldType":"TOC","instruction":"TOC", "fieldCode":[], "fieldResult":link["children"], "structuredResult":{"inline":[link, {"type":"simpleField","fieldType":"PAGE","instruction":" PAGE ","content":[{"type":"run","content":[{"type":"text","text":"1"}]}]}]}});
        let styles = StyleResolver::new(Some(
            &json!({"styles":[{"type":"character","styleId":"Hyperlink","rPr":{"color":{"rgb":"0563C1"}}}]}),
        ));
        let units =
            paragraph_units(&json!({"content":[value]}), &styles, None, &BTreeMap::new()).units;
        assert!(matches!(&units[0].content, UnitContent::Text(text) if text == "Heading"));
        assert_eq!(units[0].attrs["hyperlink"]["href"], json!("#_Toc1"));
        assert_eq!(units[0].attrs["textColor"]["rgb"], json!("0563C1"));
        let UnitContent::Embed {
            payload: nested, ..
        } = &units[1].content
        else {
            panic!("missing nested field")
        };
        assert_eq!(nested["fieldType"], json!("PAGE"));
        let UnitContent::Embed { payload, .. } = &units[2].content else {
            panic!("missing field")
        };
        assert_eq!(payload["displayText"], json!(""));
        assert!(
            payload["fieldData"]
                .as_str()
                .unwrap()
                .contains("structuredResult")
        );
    }

    #[test]
    fn comprehensive_native_layout_accepts_authored_page_number_start() {
        let bytes = include_bytes!(
            "../../betteroffice-docx/tests/corpus/fixtures/wordprocessingml-comprehensive.docx"
        );
        let parsed =
            docx_parse::parse_docx_s9_wire(bytes, docx_parse::S9ParseOptions::default()).unwrap();
        let package = parsed.document.package;
        let mut sections: Vec<Value> = package
            .document
            .sections
            .unwrap()
            .into_iter()
            .map(|section| json!({"properties":section.properties}))
            .collect();
        assert!(
            sections.iter().any(
                |section| section["properties"]["pageNumbering"]["start"].as_f64() == Some(1.0)
            )
        );
        sections.push(json!({"properties":package.document.final_section_properties}));
        let request = json!({"bodyStory":"body", "options":{"pageGap":24}, "regions":{"sections":sections, "settings":package.settings}, "renderEnv":{}}).to_string();
        let engine = crate::EngineSession::new(74003);
        seed_from_docx(engine.doc(), bytes).unwrap();
        engine.layout_font_requirements_json(&request).unwrap();
        let layout: Value =
            serde_json::from_str(&engine.layout_document_with_regions_json(&request).unwrap())
                .unwrap();
        assert!(!layout["layout"]["pages"].as_array().unwrap().is_empty());
    }

    #[test]
    fn raw_blocks_seed_no_content_control_or_child_story() {
        let mut context = LoweringContext {
            styles: StyleResolver::new(None),
            theme: None,
            source_json: Arc::new(BTreeMap::new()),
            plans: Vec::new(),
            compatibility_mode: 12,
            root: "body".to_owned(),
            paragraphs: Vec::new(),
            opaque_sequences: Vec::new(),
            source: SourceStructure::default(),
            provenance: Provenance::default(),
            locators: HashMap::new(),
        };
        visit_story(
            &mut context,
            "body".to_owned(),
            &[
                json!({"type":"paragraph","content":[]}),
                json!({"type":"rawXml","xml":"<x:block/>"}),
                json!({"type":"paragraph","content":[]}),
            ],
            StoryOptions {
                include_page_breaks: true,
                append_body_tail: true,
                seed_comments: true,
            },
        );
        assert_eq!(context.plans.len(), 1);
        assert_eq!(context.plans[0].units.len(), 2);
        assert!(context.plans[0].units.iter().all(
            |unit| matches!(&unit.content, UnitContent::Embed { kind, .. } if kind == "pilcrow")
        ));
    }

    #[test]
    fn raw_inline_nodes_leave_run_boundaries_intact() {
        let properties = paragraph_units(
            &json!({"content":[
                {"type":"run","content":[{"type":"text","text":"A"}]},
                {"type":"rawXml","xml":"<x:mark/>"},
                {"type":"run","content":[{"type":"text","text":"B"}]}
            ]}),
            &StyleResolver::new(None),
            None,
            &BTreeMap::new(),
        )
        .ppr;
        assert_eq!(
            properties["_originalRunBoundaries"]
                .as_array()
                .unwrap()
                .len(),
            2
        );
    }

    #[test]
    fn numbering_indents_precede_styles_and_ignore_zero_first_line() {
        let style_data = json!({"styles":[{"styleId":"List","type":"paragraph","pPr":{"indentLeft":720,"indentFirstLine":180,"hangingIndent":false}}]});
        for styles in [
            StyleResolver::new(Some(&style_data)),
            StyleResolver::new(None),
        ] {
            for direct in [
                json!({}),
                json!({"indentFirstLine":0}),
                json!({"indentFirstLine":0,"hangingIndent":true}),
            ] {
                let mut formatting = direct;
                formatting["styleId"] = json!("List");
                let properties = paragraph_attrs(
                    &json!({"formatting":formatting,"listRendering":{"indentLeft":1440,"indentFirstLine":-360,"hangingIndent":true},"content":[]}),
                    &styles,
                    &[],
                    &[],
                    None,
                );
                assert_eq!(properties["indentLeft"], json!(1440));
                assert_eq!(properties["indentFirstLine"], json!(-360));
                assert_eq!(properties["hangingIndent"], json!(true));
            }
            let properties = paragraph_attrs(
                &json!({"formatting":{"styleId":"List","indentLeft":0,"indentFirstLine":240,"hangingIndent":false},"listRendering":{"indentLeft":1440,"indentFirstLine":-360,"hangingIndent":true},"content":[]}),
                &styles,
                &[],
                &[],
                None,
            );
            assert_eq!(properties["indentLeft"], json!(0));
            assert_eq!(properties["indentFirstLine"], json!(240));
            assert_eq!(properties["hangingIndent"], json!(false));
        }
    }

    #[test]
    fn first_line_value_and_kind_share_one_source() {
        let hanging_styles = StyleResolver::new(Some(
            &json!({"styles":[{"styleId":"Normal","type":"paragraph","default":true,"pPr":{"indentLeft":1450,"indentFirstLine":-730,"hangingIndent":true}}]}),
        ));
        let first_styles = StyleResolver::new(Some(
            &json!({"styles":[{"styleId":"Normal","type":"paragraph","default":true,"pPr":{"indentLeft":1450,"indentFirstLine":720}}]}),
        ));
        let properties = paragraph_attrs(
            &json!({"formatting":{"indentLeft":2160,"indentFirstLine":720},"content":[]}),
            &hanging_styles,
            &[],
            &[],
            None,
        );
        assert_eq!(properties["indentLeft"], json!(2160));
        assert_eq!(properties["indentFirstLine"], json!(720));
        assert_eq!(properties["hangingIndent"], json!(false));
        let properties = paragraph_attrs(
            &json!({"formatting":{"indentLeft":720,"indentFirstLine":0},"content":[]}),
            &hanging_styles,
            &[],
            &[],
            None,
        );
        assert_eq!(properties["indentFirstLine"], json!(0));
        assert_eq!(properties["hangingIndent"], json!(false));
        let properties = paragraph_attrs(
            &json!({"formatting":{"indentLeft":1425},"content":[]}),
            &hanging_styles,
            &[],
            &[],
            None,
        );
        assert_eq!(properties["indentLeft"], json!(1425));
        assert_eq!(properties["indentFirstLine"], json!(-730));
        assert_eq!(properties["hangingIndent"], json!(true));
        let properties = paragraph_attrs(
            &json!({"formatting":{"indentLeft":2160,"indentFirstLine":-720,"hangingIndent":true},"content":[]}),
            &first_styles,
            &[],
            &[],
            None,
        );
        assert_eq!(properties["indentFirstLine"], json!(-720));
        assert_eq!(properties["hangingIndent"], json!(true));
        let properties = paragraph_attrs(
            &json!({"formatting":{"indentFirstLine":720},"listRendering":{"indentLeft":1440,"indentFirstLine":-360,"hangingIndent":true},"content":[]}),
            &hanging_styles,
            &[],
            &[],
            None,
        );
        assert_eq!(properties["indentFirstLine"], json!(720));
        assert_eq!(properties["hangingIndent"], json!(false));
        let properties = paragraph_attrs(
            &json!({"formatting":{},"listRendering":{"indentLeft":1440,"indentFirstLine":300},"content":[]}),
            &hanging_styles,
            &[],
            &[],
            None,
        );
        assert_eq!(properties["indentFirstLine"], json!(300));
        assert_eq!(properties["hangingIndent"], json!(false));
        let properties = paragraph_attrs(
            &json!({"formatting":{"indentFirstLine":0},"listRendering":{"indentLeft":1440,"indentFirstLine":-360,"hangingIndent":true},"content":[]}),
            &hanging_styles,
            &[],
            &[],
            None,
        );
        assert_eq!(properties["indentFirstLine"], json!(-360));
        assert_eq!(properties["hangingIndent"], json!(true));
        let properties = paragraph_attrs(
            &json!({"formatting":{"indentFirstLine":-720,"hangingIndent":true},"listRendering":{"indentLeft":2145},"content":[]}),
            &hanging_styles,
            &[],
            &[],
            None,
        );
        assert_eq!(properties["indentLeft"], json!(2145));
        assert_eq!(properties["indentFirstLine"], json!(-720));
        assert_eq!(properties["hangingIndent"], json!(true));
    }

    #[test]
    fn derived_first_line_without_flag_clears_base_hanging() {
        let styles = StyleResolver::new(Some(
            &json!({"docDefaults":{"pPr":{"indentFirstLine":-730,"hangingIndent":true}},"styles":[{"styleId":"Derived","type":"paragraph","pPr":{"indentFirstLine":200}}]}),
        ));
        let properties = paragraph_attrs(
            &json!({"formatting":{"styleId":"Derived"},"content":[]}),
            &styles,
            &[],
            &[],
            None,
        );
        assert_eq!(properties["indentFirstLine"], json!(200));
        assert_eq!(properties["hangingIndent"], json!(false));
        let styles = StyleResolver::new(Some(
            &json!({"docDefaults":{"pPr":{"indentFirstLine":-730,"hangingIndent":true}},"styles":[{"styleId":"Derived","type":"paragraph","pPr":{"indentFirstLine":0}}]}),
        ));
        let properties = paragraph_attrs(
            &json!({"formatting":{"styleId":"Derived"},"content":[]}),
            &styles,
            &[],
            &[],
            None,
        );
        assert_eq!(properties["indentFirstLine"], json!(0));
        assert_eq!(properties["hangingIndent"], json!(false));
        let styles = StyleResolver::new(Some(
            &json!({"docDefaults":{"pPr":{"indentFirstLine":200}},"styles":[{"styleId":"Derived","type":"paragraph","pPr":{"indentFirstLine":-360,"hangingIndent":true}}]}),
        ));
        let properties = paragraph_attrs(
            &json!({"formatting":{"styleId":"Derived"},"content":[]}),
            &styles,
            &[],
            &[],
            None,
        );
        assert_eq!(properties["indentFirstLine"], json!(-360));
        assert_eq!(properties["hangingIndent"], json!(true));
    }

    #[test]
    fn indent_kind_pairs_merge_atomically_including_zero() {
        let base = json!({"indentFirstLine":-730,"hangingIndent":true});
        let derived = json!({"indentFirstLine":200});
        let merged = merge_paragraph_formatting(Some(&base), Some(&derived)).unwrap();
        assert_eq!(merged["indentFirstLine"], json!(200));
        assert!(merged.get("hangingIndent").is_none());
        let derived = json!({"indentFirstLine":0});
        let merged = merge_paragraph_formatting(Some(&base), Some(&derived)).unwrap();
        assert_eq!(merged["indentFirstLine"], json!(0));
        assert!(merged.get("hangingIndent").is_none());
        let base = json!({"indentFirstLine":200});
        let derived = json!({"hangingIndent":true});
        let merged = merge_paragraph_formatting(Some(&base), Some(&derived)).unwrap();
        assert_eq!(merged["indentFirstLine"], json!(200));
        assert!(merged.get("hangingIndent").is_none());
        let base = json!({"indentFirstLine":-730,"hangingIndent":true});
        let derived = json!({"indentLeft":720});
        let merged = merge_paragraph_formatting(Some(&base), Some(&derived)).unwrap();
        assert_eq!(merged["indentFirstLine"], json!(-730));
        assert_eq!(merged["hangingIndent"], json!(true));
    }

    #[test]
    fn parsed_indent_xml_seeds_matching_attrs() {
        fn ppr(xml: &str) -> Value {
            let limits = docx_parse::xml::ParseLimits::default();
            let mut budget = docx_parse::xml::ParseBudget::new(&limits);
            let root = docx_parse::xml::parse_xml(xml.as_bytes(), "formatting.xml", &mut budget)
                .unwrap()
                .root()
                .unwrap()
                .clone();
            serde_json::to_value(docx_parse::parse_paragraph_properties(Some(&root), None).unwrap())
                .unwrap()
        }
        let style_ppr = ppr(r#"<w:pPr><w:ind w:left="1450" w:hanging="730"/></w:pPr>"#);
        assert_eq!(style_ppr["indentFirstLine"], json!(-730.0));
        assert_eq!(style_ppr["hangingIndent"], json!(true));
        let direct_ppr = ppr(r#"<w:pPr><w:ind w:left="2160" w:firstLine="720"/></w:pPr>"#);
        assert_eq!(direct_ppr["indentFirstLine"], json!(720.0));
        assert!(direct_ppr.get("hangingIndent").is_none());
        let styles = StyleResolver::new(Some(
            &json!({"styles":[{"styleId":"Normal","type":"paragraph","default":true,"pPr":style_ppr}]}),
        ));
        let properties = paragraph_attrs(
            &json!({"formatting": direct_ppr, "content": []}),
            &styles,
            &[],
            &[],
            None,
        );
        assert_eq!(properties["indentLeft"], json!(2160.0));
        assert_eq!(properties["indentFirstLine"], json!(720.0));
        assert_eq!(properties["hangingIndent"], json!(false));
    }

    #[test]
    fn numbering_level_marker_format_preserves_explicit_off() {
        let styles = StyleResolver::new(None);
        let properties = paragraph_attrs(
            &json!({"formatting":{},"listRendering":{"marker":"1.","numFmt":"decimal","markerBold":false,"markerItalic":false,"markerColor":{"rgb":"000000"},"markerFontFamily":"Times New Roman","markerFontSize":12.0},"content":[]}),
            &styles,
            &[],
            &[],
            None,
        );
        assert_eq!(properties["listMarker"], json!("1."));
        assert_eq!(properties["listMarkerBold"], json!(false));
        assert_eq!(properties["listMarkerItalic"], json!(false));
        assert_eq!(properties["listMarkerColor"], json!({"rgb":"000000"}));
        assert_eq!(properties["listMarkerFontFamily"], json!("Times New Roman"));
        assert_eq!(properties["listMarkerFontSize"], json!(12.0));
    }

    #[test]
    fn paragraph_and_run_formatting_preserve_uncached_resolution() {
        let definitions = json!({
            "docDefaults": { "pPr": { "alignment": "left" }, "rPr": { "bold": false } },
            "styles": [
                { "styleId": "Normal", "type": "paragraph", "default": true,
                    "pPr": { "spaceAfter": 120 }, "rPr": { "bold": true } },
                { "styleId": "Body", "type": "paragraph", "rPr": { "italic": true } },
                { "styleId": "body", "type": "paragraph", "rPr": { "italic": false } },
                { "styleId": "Character", "type": "character", "default": true,
                    "rPr": { "fontFamily": { "asciiTheme": "minorHAnsi" } } },
                { "styleId": "Accent", "type": "character",
                    "rPr": { "fontFamily": { "ascii": "Example" }, "bold": false } }
            ]
        });
        let extras = [
            Value::Null,
            json!(false),
            json!(5),
            json!("extra"),
            json!([null, {}]),
            json!({}),
            json!({ "bold": false, "fontFamily": { "ascii": "Direct" } }),
        ];
        let assert_formatting = |styles: &StyleResolver, style_id: Option<&str>| {
            let paragraph = json!({ "formatting": { "styleId": style_id }, "content": [] });
            let (ppr, run) = styles.resolve_paragraph_style_uncached(style_id);
            let inherited_run = styles.enabled.then_some(run.as_ref()).flatten();
            for extra in std::iter::once(None).chain(extras.iter().map(Some)) {
                let actual = paragraph_style_formatting(&paragraph, styles, extra);
                let expected = merge_text_formatting(inherited_run, extra);
                assert_eq!(actual.as_deref(), expected.as_ref());
            }
            let attrs = paragraph_attrs(&paragraph, styles, &[], &[], None);
            if styles.enabled {
                for key in STYLE_FALLBACK_KEYS {
                    assert_eq!(
                        attrs.get(key),
                        Some(field(ppr.as_ref(), key).unwrap_or(&Value::Null))
                    );
                }
                let character = styles
                    .default_style("character")
                    .and_then(|style| field(Some(style), "rPr"));
                let expected = merge_text_formatting(run.as_ref(), character);
                assert_eq!(
                    attrs.get("defaultTextFormatting"),
                    Some(expected.as_ref().unwrap_or(&Value::Null))
                );
            }
            let formatting = json!({ "styleId": style_id, "color": { "rgb": "123456" } });
            let run = style_id.and_then(|id| styles.resolve_run_style_uncached(Some(id)));
            assert_eq!(
                resolved_text_formatting(Some(&formatting), styles),
                merge_text_formatting(run.as_ref(), Some(&formatting))
            );
            assert_eq!(
                styles.resolve_run_style(style_id).as_deref(),
                styles.resolve_run_style_uncached(style_id).as_ref()
            );
            let memo = styles.memo.state.lock().unwrap();
            assert!(memo.bytes <= STYLE_MEMO_BYTES);
            assert!(memo.paragraphs.styles.len() <= styles.styles.len());
            assert!(memo.runs.styles.len() <= styles.styles.len());
        };
        for styles in [
            StyleResolver::new(Some(&definitions)),
            StyleResolver::new(Some(&json!({}))),
            StyleResolver::new(None),
        ] {
            for _ in 0..2 {
                for style_id in [
                    None,
                    Some(""),
                    Some("Missing"),
                    Some("Normal"),
                    Some("Body"),
                    Some("body"),
                    Some("Character"),
                    Some("Accent"),
                ] {
                    assert_formatting(&styles, style_id);
                }
            }
        }
        let styles = StyleResolver::new(Some(&json!({})));
        let absent = styles.resolve_paragraph_style(None);
        let undefined = styles.resolve_paragraph_style(Some("Missing"));
        assert_eq!(
            absent.paragraph,
            Some(json!({ "spaceAfter": 160, "lineSpacing": 259, "lineSpacingRule": "auto" }))
        );
        assert_eq!(undefined.paragraph, None);
        assert!(!Arc::ptr_eq(&absent, &undefined));
        assert!(Arc::ptr_eq(
            &undefined,
            &styles.resolve_paragraph_style(Some("AnotherMissing"))
        ));

        let font_name = "x".repeat(64 << 10);
        let mut definitions = json!({
            "docDefaults": {
                "pPr": { "alignment": "left", "runProperties": { "fontFamily": { "ascii": font_name } } },
                "rPr": { "bold": false, "fontFamily": { "ascii": font_name } }
            },
            "styles": [
                { "styleId": "Character", "type": "character", "default": true,
                    "rPr": { "italic": true } }
            ]
        });
        definitions["styles"]
            .as_array_mut()
            .unwrap()
            .extend((0..200).map(|index| {
                json!({
                    "styleId": format!("Defined{index}"), "type": "paragraph",
                    "pPr": { "spaceAfter": index }, "rPr": { "bold": index % 2 == 0 }
                })
            }));
        let mut styles = StyleResolver::new(Some(&definitions));
        let mut style_ids = Vec::new();
        for index in 0..500 {
            style_ids.push(format!("Undefined{index}"));
            if index < 200 {
                style_ids.push(format!("Defined{index}"));
            }
        }
        assert_formatting(&styles, None);
        assert!(Arc::ptr_eq(
            &styles.resolve_run_style(None).unwrap(),
            &styles.resolve_run_style(Some("Undefined0")).unwrap()
        ));
        for _ in 0..2 {
            for style_id in &style_ids {
                assert_formatting(&styles, Some(style_id));
            }
        }
        {
            let memo = styles.memo.state.lock().unwrap();
            assert!(memo.bytes > STYLE_MEMO_BYTES / 2);
            assert!(memo.paragraphs.absent.is_some());
            assert!(memo.paragraphs.undefined.is_some());
            assert!(memo.runs.unstyled.is_some());
            assert!(memo.paragraphs.styles.len() < 200);
            assert!(memo.runs.styles.len() < 200);
        }
        let cloned = styles.clone();
        assert_eq!(cloned.memo.state.lock().unwrap().bytes, 0);
        let previous = styles.set_table_paragraph_formatting(Some(json!({ "alignment": "right" })));
        assert_eq!(previous, None);
        assert_eq!(styles.memo.state.lock().unwrap().bytes, 0);
        assert_formatting(&styles, Some("Defined0"));
        let outer = styles.set_table_paragraph_formatting(Some(json!({ "alignment": "center" })));
        assert_eq!(outer, Some(json!({ "alignment": "right" })));
        assert_eq!(styles.memo.state.lock().unwrap().bytes, 0);
        assert_formatting(&styles, Some("Defined0"));
        styles.restore_table_paragraph_formatting(outer);
        assert_eq!(styles.memo.state.lock().unwrap().bytes, 0);
        assert_formatting(&styles, Some("Defined0"));
        styles.restore_table_paragraph_formatting(previous);
        assert_eq!(styles.memo.state.lock().unwrap().bytes, 0);
        assert_formatting(&styles, Some("Defined0"));
    }

    #[test]
    fn style_memo_bounds_empty_styles_with_long_ids() {
        let ids = (0..5000)
            .map(|index| format!("{index:0>1024}"))
            .collect::<Vec<_>>();
        let definitions = json!({
            "styles": ids
                .iter()
                .map(|id| json!({ "styleId": id, "type": "paragraph" }))
                .collect::<Vec<_>>()
        });
        let styles = StyleResolver::new(Some(&definitions));
        for id in &ids {
            let resolved = styles.resolve_paragraph_style(Some(id));
            let (paragraph, run) = styles.resolve_paragraph_style_uncached(Some(id));
            assert_eq!(resolved.paragraph, paragraph);
            assert_eq!(resolved.run.as_deref(), run.as_ref());
            assert_eq!(
                styles.resolve_run_style(Some(id)).as_deref(),
                styles.resolve_run_style_uncached(Some(id)).as_ref()
            );
        }
        let memo = styles.memo.state.lock().unwrap();
        assert!(memo.bytes <= STYLE_MEMO_BYTES);
        assert!(memo.paragraphs.styles.len() + memo.runs.styles.len() < ids.len());
    }

    #[test]
    fn owned_text_formatting_matches_reference_merging() {
        fn reference(target: Option<&Value>, source: Option<&Value>) -> Option<Value> {
            let Some(source_object) = object(source) else {
                return target.cloned();
            };
            let Some(target_object) = object(target) else {
                return source.cloned();
            };
            let mut result = target_object.clone();
            for (key, value) in source_object {
                if key == "fontFamily" && value.is_object() {
                    result.insert(
                        key.clone(),
                        merge_font_family(target_object.get(key), value),
                    );
                } else if key == "color" && value.is_object() {
                    let explicit = ["rgb", "themeColor", "themeTint", "themeShade"]
                        .iter()
                        .any(|key| truthy(field(Some(value), key)));
                    if !truthy(field(Some(value), "auto")) || explicit {
                        result.insert(key.clone(), value.clone());
                    }
                } else if value.is_object() {
                    result.insert(
                        key.clone(),
                        merge_plain(target_object.get(key), Some(value)).unwrap(),
                    );
                } else {
                    result.insert(key.clone(), value.clone());
                }
            }
            Some(Value::Object(result))
        }
        let values = [
            Value::Null,
            json!(false),
            json!(5),
            json!("formatting"),
            json!([null, {}]),
            json!({}),
            json!({ "bold": true, "italic": false, "fontSize": 22.0 }),
            json!({ "fontSize": -0.0, "spacing": 0.0 }),
            json!({ "fontSize": 0.0, "spacing": -0.0 }),
            json!({ "fontFamily": { "ascii": "Explicit", "hAnsiTheme": "minorHAnsi", "cs": "Complex", "extra": null } }),
            json!({ "fontFamily": { "asciiTheme": "majorHAnsi", "hAnsi": "Direct", "csTheme": "minorBidi", "extra": false } }),
            json!({ "fontFamily": null, "color": { "rgb": "123456" }, "underline": { "style": "single", "color": null } }),
            json!({ "color": { "auto": true }, "underline": { "style": "none" } }),
            json!({ "color": { "auto": true, "themeColor": "accent1" }, "underline": { "color": { "rgb": "ABCDEF" } } }),
            json!({ "color": { "auto": false, "rgb": null }, "nested": { "items": [null, { "keep": 1 }], "omit": null } }),
        ];
        for target in std::iter::once(None).chain(values.iter().map(Some)) {
            for source in std::iter::once(None).chain(values.iter().map(Some)) {
                let expected = reference(target, source);
                assert_eq!(merge_text_formatting(target, source), expected);
                let actual = merge_text_formatting_owned(target.cloned(), source);
                assert_eq!(actual, expected);
                assert_eq!(
                    serde_json::to_vec(&actual).unwrap(),
                    serde_json::to_vec(&expected).unwrap()
                );
                let inherited = reference(target, source);
                for direct in std::iter::once(None).chain(values.iter().map(Some)) {
                    assert_eq!(
                        merge_text_formatting_owned(inherited.clone(), direct),
                        reference(inherited.as_ref(), direct)
                    );
                }
            }
        }
    }

    #[test]
    fn paragraph_style_memo_tracks_table_context_values() {
        let definitions = json!({
            "styles": [
                { "styleId": "Shared", "type": "paragraph", "rPr": { "bold": true } },
                { "styleId": "shared", "type": "paragraph", "rPr": { "bold": false } }
            ]
        });
        let mut styles = StyleResolver::new(Some(&definitions));
        for context in [
            None,
            Some(Value::Null),
            Some(json!({})),
            Some(json!({ "alignment": "center", "runProperties": { "bold": false } })),
            Some(json!({ "alignment": "right", "runProperties": { "bold": false } })),
            Some(
                json!({ "spaceBefore": -0.0, "runProperties": { "fontSize": -0.0 }, "tabs": [{ "position": -0.0 }] }),
            ),
            Some(
                json!({ "spaceBefore": 0.0, "runProperties": { "fontSize": 0.0 }, "tabs": [{ "position": 0.0 }] }),
            ),
            None,
        ] {
            let previous = styles.set_table_paragraph_formatting(context.clone());
            if previous.is_some() != context.is_some()
                || serde_json::to_vec(&previous).unwrap() != serde_json::to_vec(&context).unwrap()
            {
                assert_eq!(styles.memo.state.lock().unwrap().bytes, 0);
            }
            let cached: Vec<_> = [None, Some("Missing"), Some("Shared"), Some("shared")]
                .into_iter()
                .map(|id| {
                    let resolved = styles.resolve_paragraph_style(id);
                    let (paragraph, run) = styles.resolve_paragraph_style_uncached(id);
                    assert_eq!(resolved.paragraph, paragraph);
                    assert_eq!(
                        serde_json::to_vec(&resolved.paragraph).unwrap(),
                        serde_json::to_vec(&paragraph).unwrap()
                    );
                    assert_eq!(resolved.run.as_deref(), run.as_ref());
                    (id, resolved)
                })
                .collect();
            let bytes = styles.memo.state.lock().unwrap().bytes;
            let same = styles.set_table_paragraph_formatting(context.clone());
            assert_eq!(same, context);
            assert_eq!(styles.memo.state.lock().unwrap().bytes, bytes);
            styles.restore_table_paragraph_formatting(same);
            assert_eq!(styles.memo.state.lock().unwrap().bytes, bytes);
            for (id, resolved) in cached {
                assert!(Arc::ptr_eq(&resolved, &styles.resolve_paragraph_style(id)));
            }
        }
    }

    #[test]
    fn table_lowering_preserves_reference_payloads_and_synthetic_stories() {
        let paragraph = |id, text| {
            json!({
                "type": "paragraph", "paraId": id,
                "content": [{ "type": "run", "content": [{ "type": "text", "text": text }] }]
            })
        };
        let nested = json!({
            "type": "table", "columnWidths": [900.0],
            "rows": [
                { "cells": [{ "content": [paragraph("10000002", "inner")] }] },
                { "cells": [] }
            ]
        });
        let mut table = json!({
            "type": "table",
            "formatting": {
                "look": { "firstRow": true },
                "borders": { "top": { "style": "single", "color": { "rgb": "123456", "omit": null } } },
                "floating": { "items": [null, { "keep": 1, "omit": null }], "omit": null }
            },
            "propertyChanges": [null, { "keep": 1, "omit": null }],
            "rows": [
                {
                    "formatting": { "header": true, "height": { "value": 200.0, "omit": null } },
                    "propertyChanges": [null, { "keep": false, "omit": null }],
                    "cells": [
                        {
                            "formatting": { "gridSpan": 2, "vMerge": "restart" },
                            "propertyChanges": [null, { "nested": { "keep": 0, "omit": null } }],
                            "content": [paragraph("10000001", "before"), nested, paragraph("10000003", "after")]
                        },
                        { "content": [] }
                    ]
                },
                { "cells": [
                    { "formatting": { "gridSpan": 2, "vMerge": "continue" }, "content": [paragraph("10000004", "skip")] },
                    { "content": [paragraph("10000005", "right")] }
                ] },
                { "cells": [] }
            ]
        });
        let continuation = json!({
            "type": "table", "columnWidths": [100.0, 200.0, 300.0],
            "rows": [
                { "cells": [{ "formatting": { "gridSpan": 3, "vMerge": "restart" }, "content": [] }] },
                { "cells": [{ "formatting": { "gridSpan": 3, "vMerge": "continue" }, "content": [] }] }
            ]
        });
        let reference = |table: &ProjectedTable<'_>, parent: &str, table_index| {
            let rows: Vec<_> = table
                .rows
                .iter()
                .enumerate()
                .map(|(row_index, row)| {
                    json!({
                        "trPr": value_from_map(&row.attrs),
                        "cells": row.cells.iter().enumerate().map(|(cell_index, cell)| {
                            json!({
                                "tcPr": value_from_map(&cell.attrs),
                                "story": table_cell_story_id(parent, table_index, row_index, cell_index)
                            })
                        }).collect::<Vec<_>>()
                    })
                })
                .collect();
            let tbl_pr = structural_attrs(table.attrs.clone(), &["columnWidths"]);
            let mut grid = table
                .attrs
                .get("columnWidths")
                .and_then(Value::as_array)
                .cloned()
                .unwrap_or_default();
            grid.iter_mut().for_each(drop_nulls_in_place);
            map_from_value(json!({ "tblPr": value_from_map(&tbl_pr), "grid": grid, "rows": rows }))
        };
        for grid in [
            None,
            Some(Value::Null),
            Some(json!("invalid")),
            Some(json!([])),
            Some(json!([1200.0, 1800.0, 2000.0])),
            Some(json!([-0.0, 0.0, 1.0])),
            Some(
                json!([1200.0, null, { "items": [null, { "keep": 1, "omit": null }], "omit": null }]),
            ),
        ] {
            table.as_object_mut().unwrap().remove("columnWidths");
            if let Some(grid) = grid {
                table
                    .as_object_mut()
                    .unwrap()
                    .insert("columnWidths".to_owned(), grid);
            }
            let package = json!({ "document": { "content": [
                table, continuation, { "type": "table", "rows": [] }
            ] } });
            let (context, _) = lower_package(&package, BTreeMap::new());
            assert_eq!(
                context
                    .plans
                    .iter()
                    .map(|plan| plan.story_id.as_str())
                    .collect::<Vec<_>>(),
                [
                    "body",
                    "body:t0:r0c0",
                    "body:t0:r0c0:t0:r0c0",
                    "body:t0:r0c0:t0:r1c0",
                    "body:t0:r0c1",
                    "body:t0:r1c0",
                    "body:t0:r2c0",
                    "body:t1:r0c0",
                    "body:t1:r1c0",
                ]
            );
            for (parent, table_index, source) in [
                ("body", 0, &table),
                (
                    "body:t0:r0c0",
                    0,
                    &table["rows"][0]["cells"][0]["content"][1],
                ),
                ("body", 1, &continuation),
                ("body", 2, &package["document"]["content"][2]),
            ] {
                let projected = project_table(source, &context.styles, None, 12);
                let plan = context
                    .plans
                    .iter()
                    .find(|plan| plan.story_id == parent)
                    .unwrap();
                let actual = plan
                    .units
                    .iter()
                    .filter_map(|unit| match &unit.content {
                        UnitContent::Embed { kind, payload } if kind == "table" => Some(payload),
                        _ => None,
                    })
                    .nth(table_index)
                    .unwrap();
                let expected = reference(&projected, parent, table_index);
                assert_eq!(actual, &expected);
                assert_eq!(
                    serde_json::to_vec(actual).unwrap(),
                    serde_json::to_vec(&expected).unwrap()
                );
            }
            let projected = project_table(&table, &context.styles, None, 12);
            let source_content = table["rows"][0]["cells"][0]["content"].as_array().unwrap();
            let Cow::Borrowed(content) = &projected.rows[0].cells[0].content else {
                panic!("nonempty source content must be borrowed");
            };
            assert!(std::ptr::eq(*content, source_content.as_slice()));
            assert_eq!(projected.rows[0].cells[0].attrs["rowspan"], json!(2));
            assert_eq!(projected.rows[0].cells[0].attrs["colspan"], json!(2.0));
            assert_eq!(projected.rows[1].cells.len(), 1);
            for cell in [&projected.rows[0].cells[1], &projected.rows[2].cells[0]] {
                assert!(matches!(&cell.content, Cow::Owned(_)));
                assert_eq!(
                    cell.content.as_ref(),
                    &[json!({ "type": "paragraph", "content": [] })]
                );
            }
            for story in [
                "body:t0:r0c0:t0:r1c0",
                "body:t0:r0c1",
                "body:t0:r2c0",
                "body:t1:r0c0",
                "body:t1:r1c0",
            ] {
                let plan = context
                    .plans
                    .iter()
                    .find(|plan| plan.story_id == story)
                    .unwrap();
                assert_eq!(plan.units.len(), 1);
                let unit = &plan.units[0];
                assert_eq!(unit.pm_size, 1);
                assert!(unit.attrs.is_empty());
                assert!(unit.marks.is_empty());
                let UnitContent::Embed { kind, payload } = &unit.content else {
                    panic!("an empty cell must contain a pilcrow");
                };
                assert_eq!(kind, "pilcrow");
                assert_eq!(payload["paraId"], json!(format!("{story}:p0")));
                assert!(payload.get(SOURCE_PARA_ID).is_none());
                assert!(payload.get(PARA_ORIGIN).is_none());
            }
            let batches = |context: LoweringContext| {
                context
                    .plans
                    .into_iter()
                    .map(|plan| seed_plan(plan, None).unwrap())
                    .collect::<Vec<_>>()
            };
            let expected = batches(context);
            for _ in 0..2 {
                assert_eq!(
                    batches(lower_package(&package, BTreeMap::new()).0),
                    expected
                );
            }
        }
    }

    #[test]
    fn table_cell_paragraph_styles_preserve_body_and_nested_table_formatting() {
        let runs: String = (0..32)
            .map(|index| {
                format!(
                    r#"<w:r><w:rPr><w:b w:val="{}"/></w:rPr><w:t>Same{index}😀</w:t></w:r>"#,
                    u8::from(index % 2 == 0)
                )
            })
            .collect();
        let paragraph = |id| {
            fixture::para(
                id,
                &format!(r#"<w:pPr><w:pStyle w:val="Shared"/></w:pPr>{runs}"#),
            )
        };
        let nested = format!(
            r#"<w:tbl><w:tblPr><w:tblStyle w:val="Inner"/></w:tblPr><w:tblGrid><w:gridCol w:w="1000"/></w:tblGrid><w:tr><w:tc>{}</w:tc></w:tr></w:tbl>"#,
            paragraph("10000004")
        );
        let table = format!(
            r#"<w:tbl><w:tblPr><w:tblStyle w:val="Cells"/><w:tblLook w:val="0780"/></w:tblPr><w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid><w:tr><w:tc>{}{}{nested}{}</w:tc><w:tc>{}</w:tc><w:tc>{}</w:tc><w:tc>{}{}</w:tc></w:tr></w:tbl>"#,
            paragraph("10000002"),
            paragraph("10000003"),
            paragraph("10000005"),
            paragraph("10000009"),
            paragraph("1000000A"),
            paragraph("10000006"),
            paragraph("10000007")
        );
        let styles = format!(
            r#"<w:styles {}><w:docDefaults><w:pPrDefault><w:pPr><w:jc w:val="left"/><w:spacing w:before="10"/><w:ind w:right="80"/></w:pPr></w:pPrDefault><w:rPrDefault><w:rPr><w:sz w:val="22"/></w:rPr></w:rPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:pPr><w:spacing w:after="100"/></w:pPr></w:style><w:style w:type="paragraph" w:styleId="Shared"><w:basedOn w:val="Normal"/><w:rPr><w:b/></w:rPr></w:style><w:style w:type="character" w:default="1" w:styleId="DefaultCharacter"><w:rPr><w:i/></w:rPr></w:style><w:style w:type="table" w:styleId="Cells"><w:pPr><w:ind w:right="240"/></w:pPr><w:tblStylePr w:type="firstCol"><w:pPr><w:jc w:val="center"/><w:spacing w:before="200"/></w:pPr></w:tblStylePr><w:tblStylePr w:type="lastCol"><w:pPr><w:jc w:val="right"/><w:spacing w:before="400"/></w:pPr></w:tblStylePr></w:style><w:style w:type="table" w:styleId="Inner"><w:pPr><w:jc w:val="both"/><w:spacing w:before="600"/><w:ind w:right="600"/></w:pPr></w:style></w:styles>"#,
            fixture::namespaces()
        );
        let body = format!("{}{table}{}", paragraph("10000001"), paragraph("10000008"));
        let bytes = fixture::Package::new(&body).styles(&styles).bytes();
        let lowered = lower_docx(parse_docx_for_edit(&bytes).unwrap(), None).unwrap();
        assert_eq!(
            lowered
                .context
                .plans
                .iter()
                .map(|plan| plan.story_id.as_str())
                .collect::<Vec<_>>(),
            [
                "body",
                "body:t0:r0c0",
                "body:t0:r0c0:t0:r0c0",
                "body:t0:r0c1",
                "body:t0:r0c2",
                "body:t0:r0c3",
            ]
        );
        for plan in &lowered.context.plans {
            for (index, unit) in plan
                .units
                .iter()
                .filter(|unit| matches!(&unit.content, UnitContent::Text(_)))
                .enumerate()
            {
                let index = index % 32;
                let UnitContent::Text(text) = &unit.content else {
                    unreachable!();
                };
                assert_eq!(text, &format!("Same{index}😀"));
                assert_eq!(unit.pm_size, utf16_len(text));
                let mut expected =
                    JsonObject::from([("fontSize".to_owned(), json!({ "size": 22.0 }))]);
                if index % 2 == 0 {
                    expected.insert("bold".to_owned(), Value::Bool(true));
                }
                assert_eq!(unit.attrs, expected);
            }
        }
        let actual: BTreeMap<_, Vec<_>> = lowered
            .context
            .plans
            .iter()
            .map(|plan| {
                let attrs = plan
                    .units
                    .iter()
                    .filter_map(|unit| match &unit.content {
                        UnitContent::Embed { kind, payload } if kind == "pilcrow" => {
                            Some(value_from_map(payload))
                        }
                        _ => None,
                    })
                    .collect();
                (plan.story_id.clone(), attrs)
            })
            .collect();
        let expected = |alignment, before, right, id| {
            json!({
                "paraId": id,
                "sourceParaId": id,
                "alignment": alignment,
                "spaceBefore": before,
                "spaceAfter": 100.0,
                "indentRight": right,
                "hangingIndent": false,
                "defaultTextFormatting": { "fontSize": 22.0, "bold": true, "italic": true },
                "pStyle": "Shared",
                "_originalFormatting": { "styleId": "Shared" },
                "_originalRunBoundaries": (0..32).map(|index| json!({
                    "text": format!("Same{index}😀"),
                    "formatting": { "bold": index % 2 == 0 },
                    "marksKey": if index % 2 == 0 {
                        r#"bold:{}|fontSize:{"size":22,"sizeCs":null}"#
                    } else {
                        r#"fontSize:{"size":22,"sizeCs":null}"#
                    }
                })).collect::<Vec<_>>()
            })
        };
        assert_eq!(
            actual,
            BTreeMap::from([
                (
                    "body".to_owned(),
                    vec![
                        expected("left", 10.0, 80.0, "10000001"),
                        expected("left", 10.0, 80.0, "10000008"),
                    ],
                ),
                (
                    "body:t0:r0c0".to_owned(),
                    vec![
                        expected("center", 200.0, 240.0, "10000002"),
                        expected("center", 200.0, 240.0, "10000003"),
                        expected("center", 200.0, 240.0, "10000005"),
                    ],
                ),
                (
                    "body:t0:r0c0:t0:r0c0".to_owned(),
                    vec![expected("both", 600.0, 600.0, "10000004")],
                ),
                (
                    "body:t0:r0c1".to_owned(),
                    vec![expected("left", 10.0, 240.0, "10000009")],
                ),
                (
                    "body:t0:r0c2".to_owned(),
                    vec![expected("left", 10.0, 240.0, "1000000A")],
                ),
                (
                    "body:t0:r0c3".to_owned(),
                    vec![
                        expected("right", 400.0, 240.0, "10000006"),
                        expected("right", 400.0, 240.0, "10000007"),
                    ],
                ),
            ])
        );
        let expected = EditingDoc::new(74105);
        seed_parsed_docx_with(&expected, parse_docx_for_edit(&bytes).unwrap(), None).unwrap();
        for _ in 0..3 {
            let actual = EditingDoc::new(74105);
            seed_parsed_docx_with(&actual, parse_docx_for_edit(&bytes).unwrap(), None).unwrap();
            assert_eq!(
                actual.encode_state_as_update_v1(),
                expected.encode_state_as_update_v1()
            );
            assert_eq!(
                actual.encode_state_vector_v1(),
                expected.encode_state_vector_v1()
            );
        }
    }

    #[test]
    fn paragraph_attrs_preserve_list_values_and_recursive_null_removal() {
        let paragraph = json!({
            "paraId": "12345678",
            "textId": "ABCDEF01",
            "formatting": {
                "styleId": "List",
                "numPr": { "numId": 5, "ilvl": null, "nested": { "keep": 1, "omit": null } },
                "numPrFromStyle": { "numId": 5, "ilvl": null },
                "nested": { "omit": null, "items": [null, { "keep": false, "omit": null }] },
                "omit": null
            },
            "listRendering": {
                "numFmt": "decimal",
                "isBullet": false,
                "marker": "5.",
                "markerHidden": true,
                "markerFontFamily": "Example",
                "markerFontSize": 12,
                "markerBold": false,
                "markerItalic": true,
                "markerColor": { "rgb": "123456", "themeColor": null },
                "markerSuffix": "tab",
                "levelNumFmts": ["decimal", null, { "format": "bullet", "omit": null }],
                "abstractNumId": 0,
                "startOverride": 5
            },
            "content": []
        });
        let actual = paragraph_attrs(&paragraph, &StyleResolver::new(None), &[], &[], None);
        let expected: JsonObject = serde_json::from_str(
            r#"{
            "paraId": "12345678",
            "textId": "ABCDEF01",
            "styleId": "List",
            "numPr": { "numId": 5, "nested": { "keep": 1 } },
            "numPrFromStyle": { "numId": 5 },
            "listNumFmt": "decimal",
            "listIsBullet": false,
            "listMarker": "5.",
            "listMarkerHidden": true,
            "listMarkerFontFamily": "Example",
            "listMarkerFontSize": 12.0,
            "listMarkerBold": false,
            "listMarkerItalic": true,
            "listMarkerColor": { "rgb": "123456" },
            "listMarkerSuffix": "tab",
            "listLevelNumFmts": ["decimal", null, { "format": "bullet" }],
            "listAbstractNumId": 0,
            "listStartOverride": 5,
            "_originalFormatting": {
                "styleId": "List",
                "numPr": { "numId": 5, "nested": { "keep": 1 } },
                "numPrFromStyle": { "numId": 5 },
                "nested": { "items": [null, { "keep": false }] }
            },
            "alignment": null,
            "spaceBefore": null,
            "spaceAfter": null,
            "spaceBeforeLines": null,
            "spaceAfterLines": null,
            "beforeAutospacing": null,
            "afterAutospacing": null,
            "lineSpacing": null,
            "lineSpacingRule": null,
            "indentRight": null,
            "borders": null,
            "shading": null,
            "tabs": null,
            "pageBreakBefore": null,
            "keepNext": null,
            "keepLines": null,
            "widowControl": null,
            "snapToGrid": null,
            "autoSpaceDE": null,
            "autoSpaceDN": null,
            "outlineLevel": null,
            "bidi": null,
            "spacingExplicit": null,
            "indentLeft": null,
            "indentFirstLine": null,
            "hangingIndent": false,
            "defaultTextFormatting": null
        }"#,
        )
        .unwrap();
        assert_eq!(actual, expected);
        assert_eq!(
            serde_json::to_string(&actual).unwrap(),
            serde_json::to_string(&expected).unwrap()
        );
    }

    use super::*;

    #[test]
    fn a_passed_package_digest_parses_like_the_parser_hashing_it() {
        let bytes = include_bytes!("../../../apps/demo/public/betteroffice-demo.docx");
        let (hashed, _) = docx_parse::parse_docx_s9_wire_parts_with_limits(
            bytes,
            docx_parse::S9ParseOptions {
                source_ordinals: true,
                ..docx_parse::S9ParseOptions::default()
            },
            &docx_parse::xml::ParseLimits::default(),
        )
        .unwrap();
        let (passed, _) = parse_docx_package_with_digest(bytes, package_digest(bytes)).unwrap();
        assert_eq!(
            serde_json::to_value(&passed).unwrap(),
            serde_json::to_value(&hashed).unwrap()
        );
    }

    #[test]
    fn seeding_data_urls_fails_on_an_image_part_that_cannot_be_read() {
        let mut state = 0x2545_f491_u32;
        let mut png = b"\x89PNG\r\n\x1a\n".to_vec();
        png.extend((0..8192).map(|_| {
            state ^= state << 13;
            state ^= state >> 17;
            state ^= state << 5;
            state as u8
        }));
        let mut bytes = ooxml_opc::rezip_parts(&[
            ("[Content_Types].xml".to_owned(), br#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>"#.to_vec()),
            ("_rels/.rels".to_owned(), br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"#.to_vec()),
            ("word/_rels/document.xml.rels".to_owned(), br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdImage" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/picture.png"/></Relationships>"#.to_vec()),
            ("word/media/picture.png".to_owned(), png.clone()),
            ("word/document.xml".to_owned(), br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><w:body><w:p><w:r><w:drawing><wp:inline><wp:extent cx="914400" cy="457200"/><wp:docPr id="1" name="picture"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:blipFill><a:blip r:embed="rIdImage"/></pic:blipFill></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p></w:body></w:document>"#.to_vec()),
        ])
        .unwrap();
        seed_with_layout_tokens(&EditingDoc::new(9), &bytes).unwrap();
        let late = &png[6000..6032];
        let at = bytes
            .windows(late.len())
            .position(|window| window == late)
            .expect("incompressible image bytes are stored verbatim");
        bytes[at + 16] ^= 0xff;
        assert!(seed_with_layout_tokens(&EditingDoc::new(10), &bytes).is_err());
    }

    #[test]
    fn a_seeded_image_names_its_part_and_resolves_in_any_replica_of_the_package() {
        let png = [0x89, b'P', b'N', b'G', 1, 2, 3, 4];
        let bytes = ooxml_opc::rezip_parts(&[
            ("[Content_Types].xml".to_owned(), br#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>"#.to_vec()),
            ("_rels/.rels".to_owned(), br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"#.to_vec()),
            ("word/_rels/document.xml.rels".to_owned(), br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdImage" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/picture.png"/></Relationships>"#.to_vec()),
            ("word/media/unused.png".to_owned(), vec![9; 16]),
            ("word/media/picture.png".to_owned(), png.to_vec()),
            ("word/document.xml".to_owned(), br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><w:body><w:p><w:r><w:drawing><wp:inline><wp:extent cx="914400" cy="457200"/><wp:docPr id="1" name="picture"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:blipFill><a:blip r:embed="rIdImage"/></pic:blipFill></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p></w:body></w:document>"#.to_vec()),
        ])
        .unwrap();
        let native = EditingDoc::new(6);
        seed_from_docx(&native, &bytes).unwrap();
        assert!(native.media_sources().is_empty());
        let seeded = EditingDoc::new(7);
        seed_with_layout_tokens(&seeded, &bytes).unwrap();
        let image_src = |doc: &EditingDoc| {
            let blocks = crate::bridge::yrs_doc_to_layout_blocks(
                doc,
                "body",
                &crate::bridge::RenderEnv {
                    media_tokens: true,
                    ..Default::default()
                },
            )
            .unwrap();
            let docx_layout::types::LayoutBlock::Paragraph(paragraph) = &blocks[0] else {
                panic!("the image paragraph stays a paragraph");
            };
            paragraph
                .runs
                .iter()
                .find_map(|run| match run {
                    docx_layout::types::Run::Image(image) => Some(image.src.clone()),
                    _ => None,
                })
                .unwrap()
        };
        assert_eq!(image_src(&seeded), "media:1");
        let data_url = "data:image/png;base64,iVBORwECAwQ=".to_owned();
        assert_eq!(image_src(&native), data_url);
        assert_eq!(
            seeded.media_table().unwrap().resolve("media:1"),
            Some(data_url.clone())
        );

        let state = seeded.encode_state_as_update_v1();
        assert!(
            state
                .windows(data_url.len())
                .any(|window| window == data_url.as_bytes())
        );

        let replica = EditingDoc::new(8);
        replica.apply_update_v1(&state).unwrap();
        assert!(replica.media_table().is_none());
        assert_eq!(image_src(&replica), data_url);
        replica.set_media_sources(
            crate::media::MediaSources::from_json(&seeded.media_sources().to_json()).unwrap(),
        );
        assert_eq!(image_src(&replica), "media:1");
        replica.retain_source_docx(bytes);
        assert_eq!(
            replica.media_table().unwrap().resolve("media:1"),
            Some(data_url)
        );
    }

    #[test]
    fn resolved_images_and_fonts_survive_media_projection() {
        let src = "data:image/png;base64,AQID";
        for with_field in [false, true] {
            let mut envelope = parse_docx_for_edit(include_bytes!(
                "../../../apps/demo/public/betteroffice-demo.docx"
            ))
            .unwrap();
            let mut content = vec![json!({
                "type": "run",
                "formatting": {"fontFamily": {"ascii": "Image Caption"}},
                "content": [{"type": "drawing", "image": {
                    "type": "image", "rId": "rIdImage", "src": src,
                    "size": {"width": 914400, "height": 457200},
                    "wrap": {"type": "inline"}
                }}]
            })];
            if with_field {
                content.push(json!({
                    "type": "simpleField", "instruction": " PAGE ", "fieldType": "PAGE",
                    "content": [{"type": "run", "content": [{"type": "text", "text": "1"}]}]
                }));
            }
            envelope.document.package.document.content = serde_json::from_value(json!([{
                "type": "paragraph", "paraId": "image", "content": content
            }]))
            .unwrap();
            envelope.document.package.media_entries = vec![(
                "word/media/image.png".to_owned(),
                Arc::new(docx_parse::media::MediaFile {
                    path: "word/media/image.png".to_owned(),
                    filename: Some("image.png".to_owned()),
                    mime_type: "image/png".to_owned(),
                    base64: "AQID".to_owned(),
                    data_url: src.to_owned(),
                }),
            )];
            let mut without_media = envelope.clone();
            without_media.document.package.media_entries.clear();
            let with_media_doc = EditingDoc::new(7);
            let without_media_doc = EditingDoc::new(7);
            let fonts = seed_parsed_docx_with(&with_media_doc, envelope, None).unwrap();
            assert_eq!(
                fonts,
                seed_parsed_docx_with(&without_media_doc, without_media, None).unwrap()
            );
            assert!(fonts.iter().any(|font| font == "Image Caption"));
            assert_eq!(
                with_media_doc.encode_state_as_update_v1(),
                without_media_doc.encode_state_as_update_v1()
            );
            let blocks = crate::bridge::yrs_doc_to_layout_blocks(
                &with_media_doc,
                "body",
                &crate::bridge::RenderEnv::default(),
            )
            .unwrap();
            let docx_layout::types::LayoutBlock::Paragraph(paragraph) = &blocks[0] else {
                panic!("image paragraph must remain a paragraph");
            };
            assert!(paragraph.runs.iter().any(|run| {
                matches!(run, docx_layout::types::Run::Image(image)
                    if image.src == src && image.width == 96.0 && image.height == 48.0)
            }));
        }
    }

    /// Run measurements match lowering for every content branch.
    #[test]
    fn run_content_measurements_match_lowering() {
        let contents = json!([
            null,
            {},
            {"type": "unknown"},
            {"type": "text"},
            {"type": "text", "text": null},
            {"type": "text", "text": 5},
            {"type": "text", "text": ""},
            {"type": "text", "text": "plain text"},
            {"type": "text", "text": "a😀b"},
            {"type": "tab"},
            {"type": "break"},
            {"type": "break", "breakType": null},
            {"type": "break", "breakType": 5},
            {"type": "break", "breakType": "textWrapping"},
            {"type": "break", "breakType": "page"},
            {"type": "break", "breakType": "column"},
            {"type": "break", "breakType": "unknown"},
            {"type": "softHyphen"},
            {"type": "noBreakHyphen"},
            {"type": "symbol"},
            {"type": "symbol", "char": ""},
            {"type": "symbol", "char": "invalid"},
            {"type": "symbol", "char": "D800"},
            {"type": "symbol", "char": "110000"},
            {"type": "symbol", "char": "100000000"},
            {"type": "symbol", "char": "0000"},
            {"type": "symbol", "char": "0041", "font": "Symbol"},
            {"type": "symbol", "char": "1f600"},
            {"type": "commentReference"},
            {"type": "commentReference", "id": null},
            {"type": "commentReference", "id": "1"},
            {"type": "drawing"},
            {"type": "drawing", "image": {"src": "image.png"}},
            {"type": "horizontalRule"},
            {"type": "horizontalRule", "rule": {"width": 5}},
            {"type": "shape"},
            {"type": "shape", "shape": {"type": "shape", "z": 1}},
            {"type": "chart"},
            {"type": "chart", "chart": {"type": "chart", "z": 2}},
            {"type": "footnoteRef"},
            {"type": "footnoteRef", "id": null},
            {"type": "footnoteRef", "id": "12"},
            {"type": "footnoteRef", "id": 12},
            {"type": "endnoteRef"},
            {"type": "endnoteRef", "id": null},
            {"type": "endnoteRef", "id": "12"},
            {"type": "endnoteRef", "id": false}
        ]);
        let marks = [mark("bold", vec![]), mark("hidden", vec![])];
        for source in [
            BTreeMap::new(),
            BTreeMap::from([
                (
                    r#"{"type":"shape","z":1}"#.to_owned(),
                    r#"{"z":1,"type":"shape"}"#.to_owned(),
                ),
                (
                    r#"{"type":"chart","z":2}"#.to_owned(),
                    r#"{"z":2,"type":"chart"}"#.to_owned(),
                ),
            ]),
        ] {
            for marks in [&[][..], marks.as_slice()] {
                for content in array(Some(&contents)) {
                    let units = run_content_to_units(content, marks, &source);
                    assert_eq!(run_content_unit_count(content), units.len(), "{content}");
                    assert_eq!(
                        run_content_width(content),
                        units.iter().map(unit_width).sum::<u32>(),
                        "{content}"
                    );
                }
            }
        }
    }

    /// JSON writers preserve escaping, number formatting, and entry order.
    #[test]
    fn json_writers_preserve_exact_text() {
        let value = json!({
            "array": [null, true, false, -0.0, 0, 1e21, 0.1, 5, {
                "a": "\"\\\n\r\t\u{0008}\u{000c}\u{0000}😀",
                "b": [[], {}]
            }],
            "key\"\\\n": "escaped"
        });
        assert_eq!(
            js_json(&value),
            r#"{"array":[null,true,false,0,0,1e+21,0.1,5,{"a":"\"\\\n\r\t\b\f\u0000😀","b":[[],{}]}],"key\"\\\n":"escaped"}"#
        );
        let entries = ordered_object([
            ("z", json!([5, {"a": 0.1, "b": null}])),
            ("a\"\\\n", json!("\"\\\n")),
            ("null", Value::Null),
        ]);
        assert_eq!(
            ordered_json(&entries),
            r#"{"z":[5,{"a":0.1,"b":null}],"a\"\\\n":"\"\\\n","null":null}"#
        );
        let ordered: OrderedValue = serde_json::from_str(
            r#"{"z":[null,true,false,-0.0,0,1e21,0.1,5,{"b":[],"a":"\"\\\n"}],"a":{}}"#,
        )
        .unwrap();
        assert_eq!(
            ordered.js_json(),
            r#"{"z":[null,true,false,0,0,1e+21,0.1,5,{"b":[],"a":"\"\\\n"}],"a":{}}"#
        );
    }

    /// Object nulls disappear while array nulls and empty containers remain.
    #[test]
    fn map_from_value_preserves_array_nulls() {
        let value = json!({
            "a": null,
            "b": {
                "a": null,
                "b": [null, {"a": null, "b": 0}, [null, {"a": null, "b": false}]]
            },
            "c": [null, {}, []],
            "d": {},
            "e": []
        });
        assert_eq!(
            value_from_map(&map_from_value(value)),
            json!({
                "b": {"b": [null, {"b": 0}, [null, {"b": false}]]},
                "c": [null, {}, []],
                "d": {},
                "e": []
            })
        );
        for value in [
            Value::Null,
            json!(false),
            json!(5),
            json!("text"),
            json!([null, {"a": null}]),
        ] {
            assert_eq!(map_from_value(value), JsonObject::new());
        }
    }

    #[test]
    fn source_json_preserves_wire_order_with_js_number_formatting() {
        let ordered: OrderedValue =
            serde_json::from_str(r#"{"type":"shape","z":1.0,"nested":{"b":2,"a":3}}"#).unwrap();
        let value = ordered.value();
        let mut source = BTreeMap::new();
        ordered.collect_source_json(&mut source);

        assert_eq!(
            source_json(&value, &source),
            r#"{"type":"shape","z":1,"nested":{"b":2,"a":3}}"#
        );
    }

    #[test]
    fn source_json_never_carries_source_ordinals() {
        let wire = r#"{"type":"shape","textBody":{"content":[{"type":"paragraph","sourceOrdinal":4}]},"z":1.0}"#;
        let ordered: OrderedValue = serde_json::from_str(wire).unwrap();
        let parsed: Value = serde_json::from_str(wire).unwrap();
        let mut source = BTreeMap::new();
        ordered.collect_source_json(&mut source);

        assert_eq!(
            source_json(&parsed, &source),
            r#"{"type":"shape","textBody":{"content":[{"type":"paragraph"}]},"z":1}"#
        );
        assert_eq!(
            source_json(&parsed, &BTreeMap::new()),
            r#"{"textBody":{"content":[{"type":"paragraph"}]},"type":"shape","z":1}"#
        );
    }

    fn widow_control_styles() -> Value {
        // Style chains are already merged: Body carries Normal's authored off,
        // while docDefaults sits under a style that leaves the toggle absent.
        json!({
            "docDefaults": { "pPr": { "widowControl": false } },
            "styles": [
                { "styleId": "Normal", "type": "paragraph", "default": true, "pPr": {} },
                { "styleId": "Body", "type": "paragraph", "pPr": { "widowControl": false } },
                { "styleId": "Quote", "type": "paragraph", "pPr": { "widowControl": true } }
            ]
        })
    }

    fn seeded_widow_control(styles: &StyleResolver, formatting: Value) -> Option<Value> {
        paragraph_attrs(
            &json!({ "formatting": formatting, "content": [] }),
            styles,
            &[],
            &[],
            None,
        )
        .get("widowControl")
        .cloned()
    }

    fn snap_grid_styles() -> Value {
        json!({
            "docDefaults": { "pPr": {} },
            "styles": [
                { "styleId": "Normal", "type": "paragraph", "default": true, "pPr": {} },
                { "styleId": "Body", "type": "paragraph", "pPr": { "snapToGrid": false } },
                { "styleId": "Quote", "type": "paragraph", "pPr": { "snapToGrid": true } }
            ]
        })
    }

    fn seeded_snap_to_grid(styles: &StyleResolver, formatting: Value) -> Option<Value> {
        paragraph_attrs(
            &json!({ "formatting": formatting, "content": [] }),
            styles,
            &[],
            &[],
            None,
        )
        .get("snapToGrid")
        .cloned()
    }

    #[test]
    fn snap_to_grid_is_seeded_from_the_style_and_direct_formatting() {
        let styles = StyleResolver::new(Some(&snap_grid_styles()));

        assert_eq!(seeded_snap_to_grid(&styles, json!({})), Some(Value::Null));
        assert_eq!(
            seeded_snap_to_grid(&styles, json!({ "styleId": "Body" })),
            Some(Value::Bool(false))
        );
        assert_eq!(
            seeded_snap_to_grid(&styles, json!({ "styleId": "Quote" })),
            Some(Value::Bool(true))
        );
        assert_eq!(
            seeded_snap_to_grid(&styles, json!({ "styleId": "Body", "snapToGrid": true })),
            Some(Value::Bool(true)),
            "a direct on overrides a style that opts out"
        );
        assert_eq!(
            seeded_snap_to_grid(&styles, json!({ "styleId": "Quote", "snapToGrid": false })),
            Some(Value::Bool(false)),
            "a direct off overrides a style that turns the toggle back on"
        );
    }

    #[test]
    fn widow_control_is_seeded_from_doc_defaults_the_style_and_direct_formatting() {
        let styles = StyleResolver::new(Some(&widow_control_styles()));

        assert_eq!(
            seeded_widow_control(&styles, json!({})),
            Some(Value::Bool(false)),
            "docDefaults reaches a paragraph whose style is silent"
        );
        assert_eq!(
            seeded_widow_control(&styles, json!({ "styleId": "Body" })),
            Some(Value::Bool(false))
        );
        assert_eq!(
            seeded_widow_control(&styles, json!({ "styleId": "Quote" })),
            Some(Value::Bool(true))
        );
        assert_eq!(
            seeded_widow_control(&styles, json!({ "styleId": "Body", "widowControl": true })),
            Some(Value::Bool(true))
        );
        assert_eq!(
            seeded_widow_control(
                &styles,
                json!({ "styleId": "Quote", "widowControl": false })
            ),
            Some(Value::Bool(false)),
            "a direct off overrides a style that turns the toggle back on"
        );
    }

    #[test]
    fn note_ref_marks_seed_no_story_unit_and_land_in_the_run_boundary() {
        let styles = StyleResolver::new(None);
        for (content_type, note_type) in [
            ("footnoteRefMark", "footnote"),
            ("endnoteRefMark", "endnote"),
        ] {
            let run = json!({
                "type": "run",
                "formatting": { "styleId": "FootnoteReference" },
                "content": [{ "type": content_type }, { "type": content_type }],
            });
            let units = run_to_units(&run, None, &styles, &[], &BTreeMap::new());
            assert!(units.is_empty());
            let boundary = run_boundary(&run, &units, &BTreeMap::new()).unwrap();
            assert_eq!(
                boundary.get("noteMarks"),
                Some(&json!([note_type, note_type]))
            );
            assert_eq!(boundary.get("text"), Some(&Value::String(String::new())));
        }
    }

    #[test]
    fn flow_breaks_land_in_the_run_boundary_at_their_text_offset() {
        let styles = StyleResolver::new(None);
        let wrapping = json!({
            "type": "run",
            "content": [
                { "type": "text", "text": "AB" },
                { "type": "break", "breakType": "page" },
                { "type": "break", "breakType": "textWrapping" },
            ],
        });
        let units = run_to_units(&wrapping, None, &styles, &[], &BTreeMap::new());
        // A wrapping break is an inline unit, so the run keeps no boundary.
        assert!(run_boundary(&wrapping, &units, &BTreeMap::new()).is_none());

        let run = json!({
            "type": "run",
            "content": [
                { "type": "text", "text": "AB" },
                { "type": "break", "breakType": "page" },
                { "type": "text", "text": "C" },
                { "type": "break", "breakType": "column" },
            ],
        });
        let units = run_to_units(&run, None, &styles, &[], &BTreeMap::new());
        let boundary = run_boundary(&run, &units, &BTreeMap::new()).unwrap();
        assert_eq!(boundary.get("text"), Some(&json!("ABC")));
        assert_eq!(
            boundary.get("breaks"),
            Some(&json!([
                { "offset": 2, "type": "page" },
                { "offset": 3, "type": "column" },
            ]))
        );
    }

    #[test]
    fn reused_run_units_keep_comments_out_of_saved_boundaries() {
        let ParagraphUnits {
            units,
            comment_marks,
            ppr,
            ..
        } = paragraph_units(
            &json!({"content": [
                {"type": "commentRangeStart", "id": 7},
                {"type": "run", "formatting": {"bold": true}, "content": [
                    {"type": "text", "text": "A"},
                    {"type": "tab"},
                    {"type": "softHyphen"}
                ]},
                {"type": "commentRangeEnd", "id": 7},
                {"type": "run", "content": [{"type": "footnoteRef", "id": 12}]}
            ]}),
            &StyleResolver::new(None),
            None,
            &BTreeMap::new(),
        );
        assert_eq!(units.len(), 4);
        assert_eq!(
            comment_marks,
            [
                CommentMark {
                    unit: 0,
                    start: true,
                    id: "7".to_owned()
                },
                CommentMark {
                    unit: 3,
                    start: false,
                    id: "7".to_owned()
                },
            ]
        );
        assert_eq!(units[3].pm_size, 1);
        let boundaries = ppr["_originalRunBoundaries"].as_array().unwrap();
        assert_eq!(boundaries.len(), 2);
        assert_eq!(boundaries[0]["text"], "A\t\u{00ad}");
        assert_eq!(boundaries[0]["marksKey"], "bold:{}");
        assert_eq!(boundaries[1]["text"], "12");
    }

    #[test]
    fn hyperlink_units_belong_to_the_comment_around_them() {
        let link = |text: &str| {
            json!({"type": "hyperlink", "href": "https://example.com/", "children": [
                {"type": "run", "content": [{"type": "text", "text": text}]}
            ]})
        };
        let ParagraphUnits {
            units,
            comment_marks,
            ..
        } = paragraph_units(
            &json!({"content": [
                {"type": "run", "content": [{"type": "text", "text": "See "}]},
                {"type": "commentRangeStart", "id": 7},
                link("the"),
                {"type": "commentRangeEnd", "id": 7},
                link(" link")
            ]}),
            &StyleResolver::new(None),
            None,
            &BTreeMap::new(),
        );
        assert_eq!(units.len(), 3);
        assert_eq!(units_text(&units), "See the link");
        assert!(
            units[1..]
                .iter()
                .all(|unit| unit.marks.iter().any(|mark| mark.name == "hyperlink"))
        );
        assert_eq!(
            comment_marks,
            [
                CommentMark {
                    unit: 1,
                    start: true,
                    id: "7".to_owned()
                },
                CommentMark {
                    unit: 2,
                    start: false,
                    id: "7".to_owned()
                },
            ]
        );
    }

    #[test]
    fn comment_coverage_follows_overlapping_and_cross_paragraph_markers() {
        let mut context = LoweringContext {
            styles: StyleResolver::new(None),
            theme: None,
            source_json: Arc::new(BTreeMap::new()),
            plans: Vec::new(),
            compatibility_mode: 12,
            root: "body".to_owned(),
            paragraphs: Vec::new(),
            opaque_sequences: Vec::new(),
            source: SourceStructure::default(),
            provenance: Provenance::default(),
            locators: HashMap::new(),
        };
        visit_story(
            &mut context,
            "body".to_owned(),
            &[
                json!({"type": "paragraph", "content": [
                    {"type": "commentRangeStart", "id": 3},
                    run("A"),
                    {"type": "commentRangeStart", "id": 4},
                    run("B"),
                    {"type": "commentRangeEnd", "id": 3},
                    run("C"),
                    {"type": "commentRangeEnd", "id": 4}
                ]}),
                json!({"type": "paragraph", "content": [
                    run("Before "),
                    {"type": "commentRangeStart", "id": 5},
                    run("first")
                ]}),
                json!({"type": "paragraph", "content": [
                    run("second"),
                    {"type": "commentRangeEnd", "id": 5},
                    run(" after "),
                    {"type": "commentRangeStart", "id": 6},
                    run("tail")
                ]}),
                json!({"type": "paragraph", "content": [
                    {"type": "commentRangeEnd", "id": 7},
                    {"type": "commentRangeStart", "id": 0},
                    run("😀"),
                    {"type": "commentRangeEnd", "id": 0},
                    run("!")
                ]}),
            ],
            StoryOptions {
                include_page_breaks: false,
                append_body_tail: false,
                seed_comments: true,
            },
        );
        assert_eq!(context.plans.len(), 1);
        assert_eq!(
            context.plans[0].comment_coverage,
            [
                ("3".to_owned(), vec![(0, 2)]),
                ("4".to_owned(), vec![(1, 3)]),
                ("5".to_owned(), vec![(11, 23)]),
                ("6".to_owned(), vec![(30, 34)]),
                ("0".to_owned(), vec![(35, 37)]),
            ]
        );
    }

    #[test]
    fn multi_digit_note_references_seed_one_position() {
        let styles = StyleResolver::new(None);
        let units = run_to_units(
            &json!({
                "type": "run",
                "content": [{ "type": "footnoteRef", "id": 12 }],
            }),
            None,
            &styles,
            &[],
            &BTreeMap::new(),
        );
        assert_eq!(units.len(), 1);
        assert_eq!(units[0].pm_size, 1);
    }

    #[test]
    fn widow_control_left_unauthored_anywhere_seeds_null() {
        let styles = StyleResolver::new(Some(&json!({
            "styles": [{ "styleId": "Normal", "type": "paragraph", "default": true, "pPr": {} }]
        })));

        assert_eq!(seeded_widow_control(&styles, json!({})), Some(Value::Null));
    }
}
