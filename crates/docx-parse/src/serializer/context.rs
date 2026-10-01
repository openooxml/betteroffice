//! Deterministic recursive serializer state.

use std::collections::{BTreeSet, HashMap, HashSet};

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
    paragraph_ids: BTreeSet<u32>,
    pub(crate) deletion: bool,
    pub(crate) in_control: bool,
    pub(crate) in_revision: bool,
    splice: Option<SpliceRecorder>,
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
}

impl SerializerContext {
    pub fn new(determinism: &SerializerDeterminism) -> Result<Self, ParseError> {
        determinism.validate()?;
        Ok(Self {
            ids: HexIdAllocator::from_sha256(&determinism.seed)?,
            now: determinism.now.clone(),
            rendered_page_breaks: Vec::new(),
            paragraph_ids: BTreeSet::new(),
            deletion: false,
            in_control: false,
            in_revision: false,
            splice: None,
        })
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
    /// until [`Self::end_splice`].
    pub(crate) fn begin_splice(&mut self, expected: impl IntoIterator<Item = u32>) {
        self.splice = Some(SpliceRecorder {
            expected: expected.into_iter().collect(),
            ..SpliceRecorder::default()
        });
    }

    /// The recorded paragraphs; `None` when one source paragraph was written twice differently,
    /// or one with a `sourceOrdinal` the part does not address was written.
    pub(crate) fn end_splice(&mut self) -> Option<RecordedParagraphs> {
        self.splice
            .take()
            .filter(|recorder| !recorder.conflict)
            .map(|recorder| recorder.recorded)
    }

    /// Records a paragraph written outside any other paragraph: by its `sourceOrdinal`, which
    /// must be one the part addresses, or as a paragraph the source lacks.
    pub(crate) fn record_paragraph(&mut self, ordinal: Option<u32>, xml: &str) {
        if !self.rendered_page_breaks.is_empty() {
            return;
        }
        let Some(recorder) = self.splice.as_mut() else {
            return;
        };
        let recorded = &mut recorder.recorded;
        match ordinal {
            None => recorded.inserted.push(xml.to_owned()),
            Some(ordinal) if recorder.expected.contains(&ordinal) => {
                if let Some(previous) = recorded.written.get(&ordinal) {
                    recorder.conflict |= previous != xml;
                    return;
                }
                recorded.written.insert(ordinal, xml.to_owned());
            }
            Some(_) => {
                recorder.conflict = true;
                return;
            }
        }
        recorded.order.push(ordinal);
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
