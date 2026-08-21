//! Shared round-trip helpers for the fidelity and corpus gates.

use betteroffice_docx::Document;

pub type Parts = Vec<(String, Vec<u8>)>;

pub fn parts_of(package: &[u8]) -> Parts {
    ooxml_opc::unzip_parts(package).unwrap()
}

pub fn save_unedited(package: &[u8]) -> Vec<u8> {
    Document::open(package).unwrap().save().unwrap()
}

pub fn roundtrip_report(before: &Parts, after: &Parts) -> Vec<String> {
    ooxml_fidelity::roundtrip_findings(before, after).unwrap()
}
