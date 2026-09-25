//! Deterministic recursive serializer state.

use std::collections::BTreeSet;

use crate::paragraph::HexIdAllocator;
use crate::paragraph_identity::{allocate_paragraph_id, format_paragraph_id};
use crate::xml::ParseError;

use super::s10::SerializerDeterminism;

/// Per-serialization state. Generated identities are taken only from the
/// injected seed; serializers never consult ambient randomness or a clock.
#[derive(Debug)]
pub struct SerializerContext {
    ids: HexIdAllocator,
    now: String,
    rendered_page_breaks: Vec<bool>,
    paragraph_ids: BTreeSet<u32>,
}

impl SerializerContext {
    pub fn new(determinism: &SerializerDeterminism) -> Result<Self, ParseError> {
        determinism.validate()?;
        Ok(Self {
            ids: HexIdAllocator::from_sha256(&determinism.seed)?,
            now: determinism.now.clone(),
            rendered_page_breaks: Vec::new(),
            paragraph_ids: BTreeSet::new(),
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
