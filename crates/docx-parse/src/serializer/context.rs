//! Deterministic recursive serializer state.

use std::collections::{BTreeMap, BTreeSet, HashMap, HashSet};

use crate::inline::DirectionScope;
use crate::paragraph::HexIdAllocator;
use crate::paragraph_identity::{allocate_paragraph_id, format_paragraph_id};
use crate::xml::ParseError;

use super::paragraph::is_safe_attribute_name;
use super::s10::SerializerDeterminism;
use super::xml_writer::escape_xml;

/// The most kept paragraphs one save describes; its count covers them all.
const KEPT_PARAGRAPH_LIMIT: usize = 200;

/// A paragraph a save wrote as its source bytes because writing it would have dropped structure.
#[derive(Clone, Debug, serde::Serialize)]
#[serde(rename_all = "camelCase")]
pub struct KeptParagraph {
    pub part: String,
    pub ordinal: u32,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub para_id: Option<String>,
    pub text: String,
    /// `"paragraph"`, or `"part"` when no safe splice applied and the whole part was kept.
    pub scope: String,
    pub dropped: Vec<String>,
}

/// A surviving paragraph of a part written whole that holds fewer structural elements than its
/// source.
#[derive(Clone, Debug, serde::Serialize)]
#[serde(rename_all = "camelCase")]
pub struct DroppedParagraph {
    pub ordinal: u32,
    pub para_id: String,
    pub text: String,
    pub dropped: Vec<String>,
}

/// A story part written whole with fewer structural elements than its source. The edits may
/// account for them, as a deleted table row does.
#[derive(Clone, Debug, serde::Serialize)]
#[serde(rename_all = "camelCase")]
pub struct DroppedStructure {
    pub part: String,
    pub dropped: Vec<String>,
    pub paragraphs: Vec<DroppedParagraph>,
}

/// A save's diagnostics: the text warnings and the typed values the save reports.
#[derive(Clone, Debug, Default)]
pub struct S13SaveReport {
    /// User-facing diagnostics.
    pub warnings: Vec<String>,
    /// Paragraphs kept as their source bytes, counted per part.
    pub kept_unchanged: BTreeMap<String, u32>,
    pub kept_paragraphs: Vec<KeptParagraph>,
    pub structure_dropped: Vec<DroppedStructure>,
}

/// The `w:dir` and `w:bdo` elements open where a paragraph is being written, and how many of
/// them enclose the wrapper element being written.
#[derive(Clone, Debug, Default)]
pub(crate) struct DirectionScopes {
    open: Vec<DirectionScope>,
    base: usize,
    /// How many open scopes the next node with width keeps open, as noted before a node without.
    kept: Option<usize>,
}

pub(crate) fn shared_direction_prefix(left: &[DirectionScope], right: &[DirectionScope]) -> usize {
    left.iter().zip(right).take_while(|(a, b)| a == b).count()
}

impl DirectionScopes {
    fn close_to(&mut self, len: usize) -> String {
        self.kept = None;
        let mut xml = String::new();
        let at = len.min(self.open.len());
        for scope in self.open.drain(at..).rev() {
            xml.push_str("</");
            xml.push_str(scope.kind.element());
            xml.push('>');
        }
        xml
    }

    /// The tags that leave exactly the scopes of `target` open for a run, never closing one
    /// opened outside the wrapper element being written.
    pub(crate) fn transition(&mut self, target: &[DirectionScope]) -> String {
        let shared = shared_direction_prefix(&self.open, target);
        let mut xml = self.close_to(shared.max(self.base));
        for scope in &target[shared..] {
            if self
                .open
                .iter()
                .any(|open| open.kind == scope.kind && open.id == scope.id)
            {
                continue;
            }
            xml.push('<');
            xml.push_str(scope.kind.element());
            if let Some(val) = &scope.val {
                xml.push_str(" w:val=\"");
                xml.push_str(&escape_xml(val));
                xml.push('"');
            }
            for attribute in &scope.extra {
                if is_safe_attribute_name(&attribute.name) {
                    xml.push(' ');
                    xml.push_str(&attribute.name);
                    xml.push_str("=\"");
                    xml.push_str(&escape_xml(&attribute.value));
                    xml.push('"');
                }
            }
            xml.push('>');
            self.open.push(scope.clone());
        }
        xml
    }

    /// Whether a scope opened inside the wrapper element being written is open, so the next node
    /// with width decides which of them a node without width stays inside.
    pub(crate) fn any_open(&self) -> bool {
        self.open.len() > self.base
    }

    /// Notes, before a node without width, the scopes the next node with width keeps open where
    /// it starts; `None` when no such node follows inside the wrapper element being written.
    pub(crate) fn look_ahead(&mut self, next: Option<&[DirectionScope]>) {
        self.kept = next.map(|next| shared_direction_prefix(&self.open, next));
    }

    /// The tags closing, before a node that carries no scope, each scope opened inside the
    /// wrapper element being written: before a node without width, only those the next node
    /// with width does not keep open, as [`Self::look_ahead`] noted.
    pub(crate) fn unscoped(&mut self, zero_width: bool) -> String {
        let kept = self.kept.take().filter(|_| zero_width).unwrap_or_default();
        self.close_to(kept.max(self.base))
    }

    /// Enters a wrapper element whose runs share `scopes`: closes, before it, each open scope
    /// they do not share, every one for a wrapper without runs, and keeps the others outside it.
    pub(crate) fn enter_wrapper(&mut self, scopes: Option<&[DirectionScope]>) -> (String, usize) {
        let shared = shared_direction_prefix(&self.open, scopes.unwrap_or_default());
        let xml = self.close_to(shared.max(self.base));
        let base = std::mem::replace(&mut self.base, self.open.len());
        (xml, base)
    }

    /// The tags closing the scopes opened inside the wrapper element being left.
    pub(crate) fn leave_wrapper(&mut self, base: usize) -> String {
        let xml = self.close_to(self.base);
        self.base = base;
        xml
    }

    pub(crate) fn close_all(&mut self) -> String {
        self.base = 0;
        self.close_to(0)
    }
}

/// Per-serialization state. Generated identities are taken only from the
/// injected seed; serializers never consult ambient randomness or a clock.
#[derive(Clone, Debug)]
pub struct SerializerContext {
    ids: HexIdAllocator,
    now: String,
    rendered_page_breaks: Vec<bool>,
    chart_drawings: HashMap<String, String>,
    group_drawings: HashMap<String, String>,
    warnings: Vec<String>,
    paragraph_replace_counts: BTreeMap<String, u32>,
    kept_unchanged: BTreeMap<String, u32>,
    kept_paragraphs: Vec<KeptParagraph>,
    structure_dropped: Vec<DroppedStructure>,
    paragraph_ids: BTreeSet<u32>,
    pub(crate) deletion: bool,
    pub(crate) in_control: bool,
    pub(crate) in_revision: bool,
    pub(crate) vmerge_continuation: u32,
    pub(crate) direction: DirectionScopes,
    splice: Option<SpliceRecorder>,
    document_root: Option<String>,
    form_data_token: Option<String>,
    form_data_written_token: bool,
    deferred_form_data: Vec<(String, String, String)>,
}

const FORM_DATA_TOKEN: &str = "<!--bo-ffdata ";

/// The tokens written in place of form field data, by field name and reason.
pub(crate) struct DeferredFormData {
    pub(crate) tokens: Vec<(String, String, String)>,
    /// Whether written form field data holds token text, so only position tells tokens apart.
    pub(crate) in_written_data: bool,
}

/// The XML written for each model paragraph a spliced part addresses by `sourceOrdinal`.
#[derive(Clone, Debug, Default)]
struct SpliceRecorder {
    expected: HashSet<u32>,
    recorded: RecordedParagraphs,
    conflict: bool,
}

/// The paragraphs written for a spliced part, outside any other paragraph.
#[derive(Clone, Debug, Default)]
pub(crate) struct RecordedParagraphs {
    /// The XML of each paragraph written from a source paragraph, by `sourceOrdinal`.
    pub(crate) written: HashMap<u32, String>,
    /// The XML of each paragraph without a `sourceOrdinal`, in writing order.
    pub(crate) inserted: Vec<String>,
    /// Every recorded paragraph in writing order: its `sourceOrdinal`, or `None` for the next
    /// of `inserted`.
    pub(crate) order: Vec<Option<u32>>,
    /// Whether one source paragraph was written twice differently, or one with a
    /// `sourceOrdinal` the part does not address was written; only the first is recorded.
    pub(crate) conflict: bool,
    pub(crate) token_prefix: Option<String>,
    /// The empty paragraph written for each table cell without content, as the number of
    /// `order` entries written before it.
    pub(crate) placeholders: Vec<usize>,
}

impl SerializerContext {
    pub fn new(determinism: &SerializerDeterminism) -> Result<Self, ParseError> {
        determinism.validate()?;
        Ok(Self {
            ids: HexIdAllocator::from_sha256(&determinism.seed)?,
            now: determinism.now.clone(),
            rendered_page_breaks: Vec::new(),
            chart_drawings: HashMap::new(),
            group_drawings: HashMap::new(),
            warnings: Vec::new(),
            paragraph_replace_counts: BTreeMap::new(),
            kept_unchanged: BTreeMap::new(),
            kept_paragraphs: Vec::new(),
            structure_dropped: Vec::new(),
            paragraph_ids: BTreeSet::new(),
            deletion: false,
            in_control: false,
            in_revision: false,
            vmerge_continuation: 0,
            direction: DirectionScopes::default(),
            splice: None,
            document_root: None,
            form_data_token: None,
            form_data_written_token: false,
            deferred_form_data: Vec::new(),
        })
    }

    /// Writes an unreplayable `w:ffData` as a token absent from `source` instead of failing.
    pub(crate) fn begin_form_data_tokens(&mut self, source: &str) {
        let mut prefix = FORM_DATA_TOKEN.to_owned();
        while source.contains(&prefix) {
            prefix.push('_');
        }
        self.form_data_token = Some(prefix);
        self.form_data_written_token = false;
        self.deferred_form_data.clear();
    }

    /// Records whether written form field data holds the token text itself.
    pub(crate) fn note_form_data(&mut self, raw: &str) {
        if let Some(prefix) = self.form_data_token.as_deref()
            && raw.contains(prefix)
        {
            self.form_data_written_token = true;
        }
    }

    pub(crate) fn defer_form_data(&mut self, field: &str, detail: &str) -> Option<String> {
        let prefix = self.form_data_token.as_deref()?;
        let token = format!("{prefix}{}-->", self.deferred_form_data.len());
        self.deferred_form_data
            .push((token.clone(), field.to_owned(), detail.to_owned()));
        Some(token)
    }

    pub(crate) fn end_form_data_tokens(&mut self) -> DeferredFormData {
        self.form_data_token = None;
        DeferredFormData {
            tokens: std::mem::take(&mut self.deferred_form_data),
            in_written_data: std::mem::take(&mut self.form_data_written_token),
        }
    }

    pub(crate) fn set_document_source(&mut self, source: Option<&str>) {
        self.document_root = source.and_then(|xml| {
            super::splice::root_tag(xml).map(|root| xml[..root.range.end].to_owned())
        });
    }

    pub(crate) fn document_root(&self) -> Option<&str> {
        self.document_root.as_deref()
    }

    /// Paragraph IDs the output already uses, which [`Self::allocate_paragraph_id`] avoids.
    pub fn reserve_paragraph_ids(&mut self, ids: impl IntoIterator<Item = u32>) {
        self.paragraph_ids.extend(ids);
    }

    /// A paragraph ID for `owner` through the shared deterministic allocator.
    pub fn allocate_paragraph_id(&mut self, owner: &str) -> String {
        let id = allocate_paragraph_id(owner, &self.paragraph_ids)
            .expect("parse budgets keep packages far below 2^31 paragraphs");
        self.paragraph_ids.insert(id);
        format_paragraph_id(id)
    }

    /// Authored `w:drawing` placements for one story part, keyed by chart
    /// relationship id. Sessions seeded before drawings were replayed carry
    /// chart runs without one; the source part still names their placement.
    pub fn set_chart_drawings(&mut self, drawings: HashMap<String, String>) {
        self.chart_drawings = drawings;
    }

    pub(crate) fn chart_drawing(&self, relationship_id: &str) -> Option<&str> {
        self.chart_drawings.get(relationship_id).map(String::as_str)
    }

    /// Source `w:drawing` elements holding a group or canvas in one story part, keyed by
    /// `wp:docPr` id; an id used by more than one drawing is left out.
    pub fn set_group_drawings(&mut self, drawings: HashMap<String, String>) {
        self.group_drawings = drawings;
    }

    pub(crate) fn group_drawing(&self, id: &str) -> Option<&str> {
        self.group_drawings.get(id).map(String::as_str)
    }

    /// Records a non-fatal save diagnostic for the caller to report.
    pub fn warn(&mut self, warning: impl Into<String>) {
        self.warnings.push(warning.into());
    }

    pub(crate) fn count_paragraph_replace(&mut self, reason: &str) {
        *self
            .paragraph_replace_counts
            .entry(reason.to_owned())
            .or_default() += 1;
    }

    /// Counts a paragraph of `part` written as its source bytes because its rewritten XML would
    /// drop structure, though its run formatting changed.
    pub(crate) fn count_kept_unchanged(&mut self, part: &str) {
        *self.kept_unchanged.entry(part.to_owned()).or_default() += 1;
    }

    /// Whether another kept paragraph's details fit in the save's report.
    pub(crate) fn wants_kept_paragraph(&self) -> bool {
        self.kept_paragraphs.len() < KEPT_PARAGRAPH_LIMIT
    }

    /// Reports a story part written whole with fewer structural elements than its source.
    pub(crate) fn add_structure_dropped(&mut self, dropped: DroppedStructure) {
        self.structure_dropped.push(dropped);
    }

    /// Adds a kept paragraph's details to the save's report.
    pub(crate) fn add_kept_paragraph(&mut self, paragraph: KeptParagraph) {
        self.kept_paragraphs.push(paragraph);
    }

    /// Drains the diagnostics recorded so far.
    pub fn take_warnings(&mut self) -> Vec<String> {
        std::mem::take(&mut self.warnings)
    }

    fn write_wire_diagnostics(&self, wire: &mut Vec<String>) {
        if !self.kept_unchanged.is_empty() {
            wire.push(format!(
                "DOCX_KEPT_UNCHANGED structure {}",
                serde_json::to_string(&self.kept_unchanged).expect("string keys and integer counts")
            ));
        }
        if !self.structure_dropped.is_empty() {
            wire.push(format!(
                "DOCX_STRUCTURE_DROPPED {}",
                serde_json::to_string(&self.structure_dropped).expect("strings and integers")
            ));
        }
        if !self.kept_paragraphs.is_empty() {
            wire.push(format!(
                "DOCX_KEPT_PARAGRAPHS {}",
                serde_json::to_string(&self.kept_paragraphs).expect("strings and integers")
            ));
        }
        if !self.paragraph_replace_counts.is_empty() {
            wire.push(format!(
                "DOCX_MINIMAL_PATCH replace counts {}",
                serde_json::to_string(&self.paragraph_replace_counts).expect("string keys and integer counts")
            ));
        }
    }

    /// Drains the diagnostics recorded so far, typed beside their text.
    pub(crate) fn take_report(&mut self, wire: Option<&mut Vec<String>>) -> S13SaveReport {
        if let Some(wire) = wire {
            self.write_wire_diagnostics(wire);
        }
        S13SaveReport {
            warnings: self.take_warnings(),
            kept_unchanged: std::mem::take(&mut self.kept_unchanged),
            kept_paragraphs: std::mem::take(&mut self.kept_paragraphs),
            structure_dropped: std::mem::take(&mut self.structure_dropped),
        }
    }

    pub fn allocate_hex_id(&mut self) -> String {
        self.ids.allocate()
    }

    /// Injected clock for timestamp-bearing part writers.
    pub fn now(&self) -> &str {
        &self.now
    }

    /// `wp:docPr/@id` is an unsigned decimal integer. Reuse the canonical
    /// xorshift stream, converting its valid long-hex value to decimal.
    pub fn allocate_drawing_id(&mut self) -> String {
        let hex = self.allocate_hex_id();
        u32::from_str_radix(&hex, 16)
            .expect("HexIdAllocator always returns eight hexadecimal digits")
            .to_string()
    }

    /// Records the XML written for the model paragraphs whose `sourceOrdinal` is in `expected`
    /// until [`Self::end_splice`], with tokens when `token_prefix` is supplied.
    pub(crate) fn begin_splice(
        &mut self,
        expected: impl IntoIterator<Item = u32>,
        token_prefix: Option<String>,
    ) {
        self.splice = Some(SpliceRecorder {
            expected: expected.into_iter().collect(),
            recorded: RecordedParagraphs {
                token_prefix,
                ..RecordedParagraphs::default()
            },
            ..SpliceRecorder::default()
        });
    }

    /// The recorded paragraphs.
    pub(crate) fn end_splice(&mut self) -> Option<RecordedParagraphs> {
        self.splice.take().map(|recorder| RecordedParagraphs {
            conflict: recorder.conflict,
            ..recorder.recorded
        })
    }

    /// Records a paragraph written outside any other paragraph: by its `sourceOrdinal`, which
    /// must be one the part addresses, or as a paragraph the source lacks. The token to write
    /// before it, if any.
    pub(crate) fn record_paragraph(&mut self, ordinal: Option<u32>, xml: &str) -> Option<String> {
        if !self.rendered_page_breaks.is_empty() {
            return None;
        }
        let recorder = self.splice.as_mut()?;
        let recorded = &mut recorder.recorded;
        match ordinal {
            None => recorded.inserted.push(xml.to_owned()),
            Some(ordinal) if recorder.expected.contains(&ordinal) => {
                if let Some(previous) = recorded.written.get(&ordinal) {
                    recorder.conflict |= previous != xml;
                    return None;
                }
                recorded.written.insert(ordinal, xml.to_owned());
            }
            Some(_) if self.vmerge_continuation > 0 => return None,
            Some(_) => {
                recorder.conflict = true;
                return None;
            }
        }
        recorded.order.push(ordinal);
        recorded
            .token_prefix
            .as_deref()
            .map(|prefix| super::permissions::token(prefix, recorded.order.len() - 1))
    }

    /// Records the empty paragraph written outside any other paragraph for a table cell
    /// without content.
    pub(crate) fn record_placeholder(&mut self) {
        if !self.rendered_page_breaks.is_empty() {
            return;
        }
        if let Some(recorder) = self.splice.as_mut() {
            let recorded = &mut recorder.recorded;
            recorded.placeholders.push(recorded.order.len());
        }
    }

    pub(crate) fn enter_paragraph(&mut self, rendered_page_break_before: bool) {
        self.rendered_page_breaks.push(rendered_page_break_before);
    }

    pub(crate) fn leave_paragraph(&mut self) {
        let _ = self.rendered_page_breaks.pop();
    }

    /// Consumes every active rendered-page-break marker.
    pub(crate) fn take_rendered_page_breaks(&mut self) -> usize {
        let mut count = 0;
        for pending in &mut self.rendered_page_breaks {
            if *pending {
                *pending = false;
                count += 1;
            }
        }
        count
    }
}
