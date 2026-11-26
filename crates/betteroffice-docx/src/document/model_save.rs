use docx_parse::serializer::{S13SaveRequest, SerializerDeterminism, update_core_properties};
use docx_parse::{S9ParseOptions, parse_docx_s9_wire_with_limits, write_docx_s13_parts_with_report};

use super::{body_from_wire, save_report};
use crate::{Document, Result, SaveOptions, SaveReport};

const CORE_PROPERTIES: &str = "docProps/core.xml";

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
        if self.model_matches_source(&request)
            && let Some(bytes) = self.save_unchanged_model(&request)
        {
            return Ok(SaveReport {
                bytes,
                kept_unchanged: None,
                structure_dropped: None,
                whole_model_write: false,
            });
        }

        #[cfg(test)]
        super::session_save_tests::count_package_write();
        let (bytes, report) =
            write_docx_s13_parts_with_report(request, &self.original_parts, None)?;
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
        if original != *request {
            return false;
        }
        // Signed zero can change emitted attributes even when typed equality holds.
        match (serde_json::to_vec(&original), serde_json::to_vec(request)) {
            (Ok(original), Ok(current)) => original == current,
            _ => false,
        }
    }

    fn save_unchanged_model(&self, request: &S13SaveRequest) -> Option<Vec<u8>> {
        request.determinism.validate().ok()?;
        let options = &request.options;
        if !options.update_modified_date && options.modified_by.is_none() {
            return Some(self.source.clone());
        }
        let Some((index, (_, core))) = self
            .original_parts
            .iter()
            .enumerate()
            .find(|(_, (path, _))| path == CORE_PROPERTIES)
        else {
            return Some(self.source.clone());
        };
        let updated = update_core_properties(
            &String::from_utf8_lossy(core),
            options.update_modified_date,
            options.modified_by.as_deref(),
            &request.determinism.now,
        );
        if updated.as_bytes() == core.as_slice() {
            return Some(self.source.clone());
        }
        let mut parts: Vec<_> = self
            .original_parts
            .iter()
            .map(|(path, bytes)| (path.clone(), bytes.as_slice()))
            .collect();
        parts[index].1 = updated.as_bytes();
        ooxml_opc::rezip_parts_preserving(&parts, &self.source).ok()
    }
}
