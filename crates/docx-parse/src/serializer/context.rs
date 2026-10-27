//! Deterministic recursive serializer state.

use std::collections::{BTreeMap, BTreeSet, HashMap, HashSet};

use crate::paragraph::HexIdAllocator;
use crate::paragraph_identity::{allocate_paragraph_id, format_paragraph_id};
use crate::xml::ParseError;

use super::s10::SerializerDeterminism;

/// Per-serialization state. Generated identities are taken only from the
/// injected seed; serializers never consult ambient randomness or a clock.
#[derive(Clone, Debug)]
pub struct SerializerContext {
    ids: HexIdAllocator,
    now: String,
    rendered_page_breaks: Vec<bool>,
    chart_drawings: HashMap<String, String>,
    warnings: Vec<String>,
    paragraph_replace_counts: BTreeMap<String, u32>,
    paragraph_ids: BTreeSet<u32>,
    pub(crate) deletion: bool,
    pub(crate) in_control: bool,
    pub(crate) in_revision: bool,
    splice: Option<SpliceRecorder>,
    document_root: Option<String>,
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
            warnings: Vec::new(),
            paragraph_replace_counts: BTreeMap::new(),
            paragraph_ids: BTreeSet::new(),
            deletion: false,
            in_control: false,
            in_revision: false,
            splice: None,
            document_root: None,
        })
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

    /// Drains the diagnostics recorded so far.
    pub fn take_warnings(&mut self) -> Vec<String> {
        if !self.paragraph_replace_counts.is_empty() {
            let counts = std::mem::take(&mut self.paragraph_replace_counts);
            self.warn(format!(
                "DOCX_MINIMAL_PATCH replace counts {}",
                serde_json::to_string(&counts).expect("string keys and integer counts")
            ));
        }
        std::mem::take(&mut self.warnings)
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
