//! Deterministic recursive serializer state.

use std::collections::HashMap;

use crate::paragraph::HexIdAllocator;
use crate::xml::ParseError;

use super::s10::SerializerDeterminism;

/// Per-serialization state. Generated identities are taken only from the
/// injected seed; serializers never consult ambient randomness or a clock.
#[derive(Debug)]
pub struct SerializerContext {
    ids: HexIdAllocator,
    now: String,
    rendered_page_breaks: Vec<bool>,
    chart_drawings: HashMap<String, String>,
    warnings: Vec<String>,
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
        })
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

    /// Drains the diagnostics recorded so far.
    pub fn take_warnings(&mut self) -> Vec<String> {
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
