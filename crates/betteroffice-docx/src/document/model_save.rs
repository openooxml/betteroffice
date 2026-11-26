use std::collections::BTreeMap;

use docx_parse::serializer::{S13SaveRequest, SerializerDeterminism};
use docx_parse::{
    S9ParseOptions, parse_docx_s9_wire_with_limits, write_docx_s13_parts_with_report,
};
use ooxml_opc::{MAX_TOTAL_UNCOMPRESSED_BYTES, unzip_parts_where};

use super::{body_from_wire, save_report};
use crate::{Document, Result, SaveOptions, SaveReport};

const CORE_PROPERTIES: &str = "docProps/core.xml";
const CONTENT_TYPES: &str = "[Content_Types].xml";

fn is_package_metadata(path: &str) -> bool {
    matches!(path, CORE_PROPERTIES | CONTENT_TYPES) || path.ends_with(".rels")
}

impl Document {
    pub(super) fn save_model_report(&self, options: SaveOptions) -> Result<SaveReport> {
        let determinism = SerializerDeterminism {
            seed: self.seed.clone(),
            now: options.now,
        };
        let request = self.save_request(
            determinism,
            options.update_modified_date,
            options.modified_by,
            None,
        );
        let matches = self.model_matches_source(&request);
        #[cfg(test)]
        super::session_save_tests::count_package_write();
        let (bytes, report) =
            write_docx_s13_parts_with_report(request, &self.original_parts, None)?;
        if matches && self.package_metadata_matches(&bytes) {
            return Ok(SaveReport {
                bytes: self.source.clone(),
                kept_unchanged: None,
                structure_dropped: None,
                whole_model_write: false,
            });
        }

        Ok(save_report(bytes, report, true))
    }

    fn model_matches_source(&self, request: &S13SaveRequest) -> bool {
        let Ok(parsed) = parse_docx_s9_wire_with_limits(
            &self.source,
            S9ParseOptions {
                source_ordinals: true,
                ..S9ParseOptions::default()
            },
            &self.parse_limits,
        ) else {
            return false;
        };
        let package = parsed.document.package;
        let original = S13SaveRequest {
            determinism: request.determinism.clone(),
            document: body_from_wire(package.document),
            header_entries: package.header_entries.unwrap_or_default(),
            footer_entries: package.footer_entries.unwrap_or_default(),
            footnotes: package.footnotes.unwrap_or_default(),
            endnotes: package.endnotes.unwrap_or_default(),
            footnote_separators: package.footnote_separators.unwrap_or_default(),
            endnote_separators: package.endnote_separators.unwrap_or_default(),
            relationship_entries: package.relationship_entries,
            numbering: Some(package.numbering),
            options: request.options.clone(),
            selective: None,
            paragraph_ids: None,
        };
        if original.document != self.model.body || original != *request {
            return false;
        }
        // Signed zero can change emitted attributes even when typed equality holds.
        match (serde_json::to_vec(&original), serde_json::to_vec(request)) {
            (Ok(original), Ok(current)) => original == current,
            _ => false,
        }
    }

    fn package_metadata_matches(&self, written: &[u8]) -> bool {
        let Ok(written_parts) =
            unzip_parts_where(written, MAX_TOTAL_UNCOMPRESSED_BYTES, is_package_metadata)
        else {
            return false;
        };
        let source_parts: BTreeMap<_, _> = self
            .original_parts
            .iter()
            .filter(|(path, _)| is_package_metadata(path))
            .map(|(path, bytes)| (path, bytes))
            .collect();
        let written_parts: BTreeMap<_, _> = written_parts
            .iter()
            .map(|(path, bytes)| (path, bytes))
            .collect();
        source_parts == written_parts
    }
}
