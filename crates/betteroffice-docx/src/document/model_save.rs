use std::collections::HashMap;

use docx_parse::serializer::{S13SaveRequest, SerializerDeterminism, update_core_properties};
use docx_parse::{S9ParseOptions, parse_docx_s9_wire_with_limits, write_docx_s13_parts_with_report};

use super::{Dirty, body_from_wire, save_report};
use crate::{Document, Result, SaveOptions, SaveReport};

const CORE_PROPERTIES: &str = "docProps/core.xml";
const MODEL_PARTS_REWRITTEN: &str = "DOCX_MODEL_PARTS_REWRITTEN";

impl Document {
    /// The save report and warnings from saving a mutably borrowed model.
    pub fn save_report_with_warnings(&self) -> Result<(SaveReport, Vec<String>)> {
        self.save_report_with_warnings_and_options(SaveOptions::default())
    }

    /// The save report and model-save warnings with the supplied options.
    pub fn save_report_with_warnings_and_options(
        &self,
        options: SaveOptions,
    ) -> Result<(SaveReport, Vec<String>)> {
        if self.dirty == Dirty::ModelChanged {
            return self.save_model_report(options, true);
        }
        Ok((self.save_report_with_options(options)?, Vec::new()))
    }

    pub(super) fn save_model_report(
        &self,
        options: SaveOptions,
        include_warnings: bool,
    ) -> Result<(SaveReport, Vec<String>)> {
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
            return Ok((save_report(bytes, Default::default(), false), Vec::new()));
        }

        #[cfg(test)]
        super::session_save_tests::count_package_write();
        let (bytes, mut report) =
            write_docx_s13_parts_with_report(request, &self.original_parts, None)?;
        let warnings = if include_warnings {
            report
                .warnings
                .push(model_parts_rewritten_warning(&self.original_parts, &bytes));
            std::mem::take(&mut report.warnings)
        } else {
            Vec::new()
        };
        Ok((save_report(bytes, report, true), warnings))
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

fn model_parts_rewritten_warning(original: &[(String, Vec<u8>)], output: &[u8]) -> String {
    let parts = ooxml_opc::unzip_parts(output)
        .ok()
        .and_then(|written| {
            let original: HashMap<_, _> =
                original.iter().map(|(path, bytes)| (path, bytes)).collect();
            let mut paths: Vec<_> = written
                .iter()
                .filter(|(path, bytes)| {
                    path != CORE_PROPERTIES && original.get(path).copied() != Some(bytes)
                })
                .map(|(path, _)| path)
                .collect();
            paths.sort_unstable();
            serde_json::to_string(&paths).ok()
        })
        .unwrap_or_else(|| "(part names unavailable)".to_owned());
    format!(
        "{MODEL_PARTS_REWRITTEN} {parts} Use createYrsSession/seedFromDocx/saveYrsDocx \
         to keep unchanged content when editing."
    )
}
