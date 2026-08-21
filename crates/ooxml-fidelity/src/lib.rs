//! Fidelity oracles over package bytes.
//!
//! The oracles parse saved bytes with their own XML reader, never the typed
//! model, so they cannot share the parser's blind spots. Two oracles, neither
//! compensating for the other: the structural fingerprint answers "is this the
//! same tree?", the semantic digest answers "did a save→reopen keep the
//! meaning?". The element census catches drops nobody predicted. Governed by
//! `openspec/changes/docx-word-fidelity/`.

mod census;
mod error;
mod fingerprint;
mod registry;
mod report;
pub mod wml;
mod xml;

pub use census::{Census, Loss, element_census, losses};
pub use error::FidelityError;
pub use fingerprint::{
    short_fingerprint, structural_fingerprint, structural_fingerprint_excluding,
};
pub use registry::{ComparisonMode, DECLARED_NORMALIZATIONS, Normalization, comparison_mode};
pub use report::roundtrip_findings;
pub use xml::{XML_NAMESPACE, XmlAttribute, XmlElement, XmlLimits, XmlNode, parse_part};

/// True when a part name denotes an XML part the oracles read.
pub fn is_xml_part(name: &str) -> bool {
    name.ends_with(".xml") || name.ends_with(".rels")
}
