use std::sync::Arc;

use ooxml_opc::PackageBytes;
use serde::Serialize;

use super::{EditSessionCore, EngineError, EngineResult};
use crate::engine::RelayoutTrigger;
use crate::seed::checked_package_digest;

#[derive(Serialize)]
#[serde(rename_all = "camelCase")]
struct DocxHostWire {
    envelope: docx_parse::S9WireEnvelope,
    referenced_fonts: Vec<String>,
    /// Empty unless the stories were seeded from the package (a preview's
    /// from its cut).
    unused_script_fonts: Vec<String>,
    /// A preview whose cut holds the whole body.
    #[serde(skip_serializing_if = "std::ops::Not::not")]
    whole_body: bool,
}

fn thin_header_footer(
    entries: &Option<Vec<(String, docx_parse::HeaderFooter)>>,
) -> Option<Vec<(String, docx_parse::HeaderFooter)>> {
    entries.as_ref().map(|entries| {
        entries
            .iter()
            .map(|(relationship_id, part)| {
                (
                    relationship_id.clone(),
                    docx_parse::HeaderFooter {
                        story_type: part.story_type.clone(),
                        hdr_ftr_type: part.hdr_ftr_type.clone(),
                        content: Vec::new(),
                        custom_root_bindings: part.custom_root_bindings.clone(),
                        watermark: part.watermark.clone(),
                    },
                )
            })
            .collect()
    })
}

fn thin_notes(notes: &Option<Vec<docx_parse::Note>>) -> Option<Vec<docx_parse::Note>> {
    notes.as_ref().map(|notes| {
        notes
            .iter()
            .map(|note| docx_parse::Note {
                story_type: note.story_type.clone(),
                custom_root_bindings: note.custom_root_bindings.clone(),
                id: note.id,
                note_type: note.note_type.clone(),
                content: Vec::new(),
                verbatim_xml: None,
                source_ordinal: None,
            })
            .collect()
    })
}

fn thin_docx_envelope(envelope: &docx_parse::S9WireEnvelope) -> docx_parse::S9WireEnvelope {
    let package = &envelope.document.package;
    let sections = package.document.sections.as_ref().map(|sections| {
        sections
            .iter()
            .map(|section| docx_parse::S9SectionWire {
                id: section.id.clone(),
                properties: section.properties.clone(),
                content_start: 0,
                content_end: 0,
            })
            .collect()
    });
    docx_parse::S9WireEnvelope {
        wire_version: envelope.wire_version,
        document: docx_parse::S9DocumentWire {
            package: docx_parse::S9PackageWire {
                document: docx_parse::S9DocumentBodyWire {
                    content: Vec::new(),
                    sections,
                    final_section_properties: package.document.final_section_properties.clone(),
                    custom_root_bindings: package.document.custom_root_bindings.clone(),
                    comments: package.document.comments.clone(),
                },
                styles: package.styles.clone(),
                theme: package.theme.clone(),
                numbering: docx_parse::NumberingDefinitions::default(),
                settings: package.settings.clone(),
                font_table: package.font_table.clone(),
                header_entries: thin_header_footer(&package.header_entries),
                footer_entries: thin_header_footer(&package.footer_entries),
                footnotes: thin_notes(&package.footnotes),
                endnotes: thin_notes(&package.endnotes),
                footnote_separators: None,
                endnote_separators: None,
                relationship_entries: package.relationship_entries.clone(),
                media_entries: Vec::new(),
                chart_entries: Vec::new(),
            },
            template_variables: None,
            warnings: envelope.document.warnings.clone(),
        },
        embedded_font_parts: envelope.embedded_font_parts.clone(),
        font_table_relationships_xml: envelope.font_table_relationships_xml.clone(),
        canonical_base64: None,
        canonical_sha256: None,
    }
}

impl EditSessionCore {
    /// The source package's style cascade, built on first use and cached until the source changes.
    pub(crate) fn paragraph_style_resolver(
        &self,
    ) -> EngineResult<std::cell::RefMut<'_, Option<crate::seed::StyleResolver>>> {
        let mut styles = self.paragraph_styles.borrow_mut();
        if styles.is_none()
            && let Some(source) = self.docx_source.borrow().as_ref()
        {
            let parts = ooxml_opc::unzip_parts_where(
                source,
                ooxml_opc::MAX_TOTAL_UNCOMPRESSED_BYTES,
                |name| {
                    matches!(
                        name,
                        "word/styles.xml" | "word/theme/theme1.xml" | "word/settings.xml"
                    )
                },
            )
            .map_err(EngineError::message)?;
            *styles =
                Some(crate::seed::StyleResolver::from_parts(&parts).map_err(EngineError::message)?);
        }
        Ok(styles)
    }

    pub fn seed_from_docx(&self, bytes: &[u8], generation: Option<String>) -> EngineResult<String> {
        self.open_docx_inner(bytes.into(), true, generation.as_deref(), None)
    }

    pub fn open_docx(
        &self,
        bytes: &[u8],
        seed_stories: bool,
        generation: Option<String>,
        digest: Option<String>,
    ) -> EngineResult<String> {
        self.open_docx_inner(
            bytes.into(),
            seed_stories,
            generation.as_deref(),
            digest.as_deref(),
        )
    }

    pub fn open_docx_preview(&self, bytes: &[u8], blocks: u32) -> EngineResult<Option<String>> {
        self.open_docx_preview_with_budget(bytes, blocks, None)
    }

    pub fn open_docx_preview_with_budget(
        &self,
        bytes: &[u8],
        blocks: u32,
        paragraph_budget: Option<u32>,
    ) -> EngineResult<Option<String>> {
        self.open_preview_source(
            bytes,
            blocks as usize,
            paragraph_budget.map(|value| value as usize),
        )
        .map_err(|error| EngineError::message(&error))
    }

    fn open_docx_inner(
        &self,
        bytes: PackageBytes,
        seed_stories: bool,
        generation: Option<&str>,
        digest: Option<&str>,
    ) -> EngineResult<String> {
        self.open_docx_retaining_source(bytes, seed_stories, generation, digest)
            .map_err(|error| EngineError::message(&error))
    }

    #[cfg(test)]
    pub(crate) fn open_docx_retaining(
        &self,
        bytes: &[u8],
        seed_stories: bool,
        generation: Option<&str>,
        digest: Option<&str>,
    ) -> Result<String, String> {
        self.open_docx_retaining_source(bytes.into(), seed_stories, generation, digest)
    }

    /// Opens `source`, retaining it and its digest only once opening succeeds.
    pub(crate) fn open_docx_retaining_source(
        &self,
        source: PackageBytes,
        seed_stories: bool,
        generation: Option<&str>,
        digest: Option<&str>,
    ) -> Result<String, String> {
        let digest = match digest {
            Some(digest) => checked_package_digest(digest)?,
            None => crate::seed::package_digest(&source),
        };
        let (envelope, parts, media) =
            crate::seed::parse_docx_package_with_media(source.clone(), digest.clone())?;
        let host_envelope = thin_docx_envelope(&envelope);
        // Only stories seeded into an empty replica are what a fresh seed of the package holds.
        let fresh_seed = seed_stories && self.engine.doc().holds_nothing();
        let fonts = if seed_stories {
            let seed_media = if self.media_tokens.get() {
                crate::seed::SeedMedia::AsParsed
            } else {
                crate::seed::SeedMedia::DataUrls {
                    table: &media,
                    layout_tokens: true,
                }
            };
            let fonts = crate::seed::seed_parsed_docx(
                self.engine.doc(),
                envelope,
                parts,
                source.clone(),
                digest.clone(),
                seed_media,
            )
            .map_err(|error| error.to_string())?;
            self.awaiting_comment_baseline.set(false);
            self.engine.doc().begin_opening(generation);
            self.engine.doc().install_media(media);
            fonts
        } else {
            let (mut metadata, index, fonts) =
                crate::seed::replica_source(envelope, parts, source.clone(), digest.clone())?;
            metadata.watch_comments(self.engine.doc());
            self.engine
                .doc()
                .install_source(metadata, self.host.entropy());
            self.awaiting_comment_baseline.set(true);
            self.engine
                .doc()
                .retain_source(crate::identity::SourcePackage::Ready(Arc::new(index)));
            self.engine.doc().install_media(media);
            self.engine
                .doc()
                .set_media_sources(crate::media::MediaSources::default());
            crate::seed::SeededFonts {
                referenced: fonts,
                unused_script: Vec::new(),
            }
        };
        let host = DocxHostWire {
            envelope: host_envelope,
            referenced_fonts: fonts.referenced,
            unused_script_fonts: fonts.unused_script,
            whole_body: false,
        };
        let json = serde_json::to_string(&host).map_err(|error| error.to_string())?;
        self.docx_source.replace(Some(source));
        self.paragraph_styles.replace(None);
        self.docx_digest.replace(Some(digest));
        self.engine.set_partial_document(false);
        self.engine.set_relayout_trigger(RelayoutTrigger::Open);
        self.engine.doc().rotate_version(self.host.entropy());
        if fresh_seed {
            self.engine.doc().mark_seeded();
        }
        Ok(json)
    }

    #[cfg(test)]
    pub(crate) fn open_preview(
        &self,
        bytes: &[u8],
        blocks: usize,
    ) -> Result<Option<String>, String> {
        self.open_preview_source(bytes, blocks, None)
    }

    pub(crate) fn open_preview_source(
        &self,
        bytes: impl Into<PackageBytes>,
        blocks: usize,
        paragraph_budget: Option<usize>,
    ) -> Result<Option<String>, String> {
        // A preview holds a cut of the document and must never be saved, so
        // it only opens into a session with nothing to save.
        if self.docx_source.borrow().is_some() || !self.story_ids().is_empty() {
            return Err("a preview opens only into an empty session".to_owned());
        }
        let Some((envelope, media, budget_stopped)) =
            crate::seed::parse_docx_preview(bytes.into(), blocks, paragraph_budget)?
        else {
            return Ok(None);
        };
        let host_envelope = thin_docx_envelope(&envelope);
        let whole_body =
            !budget_stopped && envelope.document.package.document.content.len() < blocks;
        let fonts = crate::seed::seed_preview_envelope(self.engine.doc(), envelope, media)?;
        self.awaiting_comment_baseline.set(false);
        self.engine.set_partial_document(true);
        self.engine.set_relayout_trigger(RelayoutTrigger::Open);
        self.engine.doc().rotate_version(self.host.entropy());
        serde_json::to_string(&DocxHostWire {
            envelope: host_envelope,
            referenced_fonts: fonts.referenced,
            unused_script_fonts: fonts.unused_script,
            whole_body,
        })
        .map(Some)
        .map_err(|error| error.to_string())
    }

    pub fn seed_from_docx_owned(
        &self,
        bytes: Vec<u8>,
        generation: Option<String>,
    ) -> EngineResult<String> {
        self.open_docx_inner(bytes.into(), true, generation.as_deref(), None)
    }

    pub fn open_docx_owned(
        &self,
        bytes: Vec<u8>,
        seed_stories: bool,
        generation: Option<String>,
        digest: Option<String>,
    ) -> EngineResult<String> {
        self.open_docx_inner(
            bytes.into(),
            seed_stories,
            generation.as_deref(),
            digest.as_deref(),
        )
    }

    pub fn open_docx_preview_owned(
        &self,
        bytes: Vec<u8>,
        blocks: u32,
    ) -> EngineResult<Option<String>> {
        self.open_docx_preview_with_budget_owned(bytes, blocks, None)
    }

    pub fn open_docx_preview_with_budget_owned(
        &self,
        bytes: Vec<u8>,
        blocks: u32,
        paragraph_budget: Option<u32>,
    ) -> EngineResult<Option<String>> {
        self.open_preview_source(
            bytes,
            blocks as usize,
            paragraph_budget.map(|value| value as usize),
        )
        .map_err(|error| EngineError::message(&error))
    }
}
