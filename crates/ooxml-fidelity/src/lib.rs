//! Fidelity oracles over package bytes, read by their own parser so they
//! cannot share the engine's blind spots. Neither covers for the other: the
//! fingerprint answers "the same tree?", the digest "the same meaning after a
//! save→reopen?", the census catches drops nobody predicted.

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

/// One package part: its name and its bytes.
pub type Part = (String, Vec<u8>);

/// True when a part name denotes an XML part the oracles read.
pub fn is_xml_part(name: &str) -> bool {
    name.ends_with(".xml") || name.ends_with(".rels")
}
