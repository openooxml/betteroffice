use std::collections::{BTreeMap, BTreeSet, btree_map::Entry};
use std::sync::Arc;
#[cfg(any(feature = "fonts", feature = "raster"))]
use std::sync::Mutex;
#[cfg(feature = "fonts")]
use std::sync::PoisonError;

#[cfg(feature = "fonts")]
use betteroffice_fonts::FontManager;

use docx_edit::content_controls::{
    ContentControlQuery, ContentControlsOptions, ContentControlsSnapshot,
    find_package_content_controls, list_package_content_controls,
};
use docx_edit::structured::{
    DocxStructuredContent, ExportOptions, MarkdownContent, MarkdownOptions,
    export_package_structured,
};
use docx_edit::{
    EditCtx, EditingDoc, ParagraphIdentities, ParagraphRef, ParagraphSavePlan, Receipt,
    SegmentContent, SpliceAnchor, StoryRange, StorySegment, seed_from_docx_with_generation,
};
use docx_layout::types::Input as LayoutInput;
use docx_parse::block::BlockContent;
use docx_parse::document::{DocumentBody, Section, get_paragraph_text};
use docx_parse::inline::{InlineNode, Run, RunContent, RunType};
use docx_parse::notes::Note;
use docx_parse::numbering::NumberingDefinitions;
use docx_parse::paragraph::{Paragraph, ParagraphContent};
use docx_parse::paragraph_identity::{
    allocate_paragraph_id, format_paragraph_id, package_paragraph_ids,
};
use docx_parse::relationships::Relationship;
use docx_parse::s9::{
    S9DocumentBodyWire, S9DocumentWire, S9PackageWire, S9ParseOptions, S9SectionWire,
    S9WireEnvelope,
};
use docx_parse::serializer::paragraph_ids::{S13TouchedPart, S13WholePart};
use docx_parse::serializer::{
    S13ParagraphId, S13ParagraphIds, S13PatchedPart, S13SaveOptions, S13SaveRequest,
    S13SpliceAnchor, S13SplicedPart, SerializerDeterminism, write_docx_s13_parts,
};
use docx_parse::table::Table;
use docx_parse::xml::ParseLimits;
use sha2::{Digest, Sha256};

use crate::types::DEFAULT_SERIALIZATION_TIME;
use crate::{
    DocumentModel, DocumentStructure, Error, LayoutResult, OpenOptions, Result, SaveOptions,
};

/// The product worker's save seed caps (`residentEngineWorker.ts`).
const SAVE_SEED_MAX_PARAGRAPHS: u32 = 4096;
const SAVE_SEED_MAX_INFLATED_BYTES: u32 = 16 * 1024 * 1024;

pub struct Document {
    /// Inflated package parts retained from open for `save` to reuse.
    pub(crate) original_parts: Vec<(String, Vec<u8>)>,
    source: Vec<u8>,
    seed: String,
    model: DocumentModel,
    /// By body paragraph ordinal, the fresh ID handed to a paragraph that
    /// repeats an earlier paragraph's ID and the authored ID it saves with
    /// until it is edited or the model is changed.
    authored_ids: BTreeMap<usize, (String, String)>,
    opened_ids: BTreeMap<usize, (String, String)>,
    /// What a save writes when the model holds only the session edits; `None` once it cannot.
    session_baseline: Option<SavedFields>,
    session_edits: Vec<SessionEdit>,
    dirty: Dirty,
    #[cfg(any(feature = "fonts", feature = "raster"))]
    pub(crate) fonts: Mutex<crate::font::FontRegistry>,
    #[cfg(feature = "fonts")]
    pub(crate) font_manager: FontManager,
    #[cfg(feature = "raster")]
    pub(crate) images: crate::render::ImageRegistry,
}

/// What changed since open.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum Dirty {
    Clean,
    LoggedEdits,
    /// Set by any `model_mut` borrow.
    ModelChanged,
}

impl Document {
    pub fn open(bytes: &[u8]) -> Result<Self> {
        Self::open_with_limits(bytes, &ParseLimits::default())
    }

    #[cfg(feature = "fonts")]
    pub fn open_with_options(bytes: &[u8], options: &OpenOptions) -> Result<Self> {
        Self::open_with_limits_and_options(bytes, &ParseLimits::default(), options)
    }

    /// [`Document::open`] under a caller-supplied parse budget. A document that
    /// exceeds any limit is rejected rather than truncated.
    pub fn open_with_limits(bytes: &[u8], limits: &ParseLimits) -> Result<Self> {
        Self::open_with_limits_and_options(bytes, limits, &OpenOptions::default())
    }

    fn open_with_limits_and_options(
        bytes: &[u8],
        limits: &ParseLimits,
        options: &OpenOptions,
    ) -> Result<Self> {
        #[cfg(not(feature = "fonts"))]
        let _ = options;
        #[cfg(feature = "fonts")]
        let font_manager = FontManager::new(&options.font_source, options.font_loading)?;
        let (parsed, original_parts) = docx_parse::parse_docx_s9_wire_parts_with_limits(
            bytes,
            S9ParseOptions {
                source_ordinals: true,
                ..S9ParseOptions::default()
            },
            limits,
        )?;
        let mut document = parsed.document;
        let authored_ids =
            address_repeated_paragraphs(&mut document.package.document.content, &original_parts);
        let model = model_from_package(
            document.package,
            document.template_variables.unwrap_or_default(),
            document.warnings.unwrap_or_default(),
        );
        let document = Self {
            original_parts,
            source: bytes.to_vec(),
            seed: format!("{:x}", Sha256::digest(bytes)),
            session_baseline: Some(SavedFields::of(&model)),
            session_edits: Vec::new(),
            dirty: Dirty::Clean,
            model,
            opened_ids: authored_ids.clone(),
            authored_ids,
            #[cfg(any(feature = "fonts", feature = "raster"))]
            fonts: Mutex::new(crate::font::FontRegistry::default()),
            #[cfg(feature = "fonts")]
            font_manager,
            #[cfg(feature = "raster")]
            images: crate::render::ImageRegistry::default(),
        };
        #[cfg(feature = "fonts")]
        let mut document = document;
        #[cfg(feature = "fonts")]
        document.load_embedded_fonts()?;
        #[cfg(feature = "fonts")]
        for font in &options.fonts {
            document.register_font(&font.family, font.bold, font.italic, &font.data)?;
        }
        Ok(document)
    }

    pub fn model(&self) -> &DocumentModel {
        &self.model
    }

    /// The model, for any change. Paragraphs that repeat an earlier
    /// paragraph's ID then save with the fresh IDs they read with.
    pub fn model_mut(&mut self) -> &mut DocumentModel {
        self.dirty = Dirty::ModelChanged;
        self.authored_ids.clear();
        &mut self.model
    }

    pub fn into_model(self) -> DocumentModel {
        self.model
    }

    pub fn body(&self) -> &DocumentBody {
        &self.model.body
    }

    pub fn headers(&self) -> &[(String, docx_parse::HeaderFooter)] {
        &self.model.headers
    }

    pub fn footers(&self) -> &[(String, docx_parse::HeaderFooter)] {
        &self.model.footers
    }

    pub fn sections(&self) -> &[Section] {
        self.model.body.sections.as_deref().unwrap_or_default()
    }

    /// Every body paragraph in document order, table cells and content controls
    /// included. One that repeats an earlier paragraph's ID carries a fresh ID,
    /// which a save writes once the paragraph is edited.
    pub fn paragraphs(&self) -> Vec<&Paragraph> {
        let mut paragraphs = Vec::new();
        collect_paragraphs(&self.model.body.content, &mut paragraphs);
        paragraphs
    }

    pub fn tables(&self) -> Vec<&Table> {
        let mut tables = Vec::new();
        collect_tables(&self.model.body.content, &mut tables);
        tables
    }

    /// The body paragraph with `para_id`; `None` when no paragraph or several have it.
    pub fn paragraph(&self, para_id: &str) -> Option<&Paragraph> {
        self.unique_paragraph(para_id).ok()
    }

    fn unique_paragraph(&self, para_id: &str) -> Result<&Paragraph> {
        let mut matches = self
            .paragraphs()
            .into_iter()
            .filter(|paragraph| paragraph.para_id.as_deref() == Some(para_id));
        let found = matches
            .next()
            .ok_or_else(|| Error::ParagraphNotFound(para_id.to_owned()))?;
        match matches.next() {
            Some(_) => Err(Error::AmbiguousParagraph(para_id.to_owned())),
            None => Ok(found),
        }
    }

    pub fn structure(&self) -> DocumentStructure {
        DocumentStructure {
            body_paragraphs: self.paragraphs().len(),
            body_tables: self.tables().len(),
            sections: self.sections().len(),
            headers: self.model.headers.len(),
            footers: self.model.footers.len(),
            footnotes: self.model.footnotes.len(),
            endnotes: self.model.endnotes.len(),
        }
    }

    pub fn replace_paragraph_text(&mut self, para_id: &str, text: &str) -> Result<Receipt> {
        let context = EditCtx::local("betteroffice-docx", DEFAULT_SERIALIZATION_TIME);
        self.replace_paragraph_text_with(para_id, text, 1, &context)
    }

    pub fn replace_paragraph_text_with(
        &mut self,
        para_id: &str,
        text: &str,
        client_id: u64,
        context: &EditCtx,
    ) -> Result<Receipt> {
        self.unique_paragraph(para_id)?;
        let paragraph = find_paragraph_mut(&mut self.model.body.content, para_id)
            .ok_or_else(|| Error::ParagraphNotFound(para_id.to_owned()))?;
        let template = plain_run_template(paragraph)
            .ok_or_else(|| Error::UnsupportedParagraphEdit(para_id.to_owned()))?;
        let current = get_paragraph_text(paragraph);
        let style = paragraph
            .formatting
            .as_ref()
            .and_then(|formatting| formatting.style_id.as_deref())
            .unwrap_or("Normal");
        let alignment = paragraph
            .formatting
            .as_ref()
            .and_then(|formatting| formatting.alignment.as_deref())
            .unwrap_or("left");
        let editor = EditingDoc::new(client_id);
        editor.create_story_with_paragraph_id("body", para_id, &current, style, alignment)?;
        let receipt = editor.replace_range(
            context,
            StoryRange::new("body", 0, utf16_len(&current)),
            text,
        )?;
        let edited = editor
            .paragraphs("body")?
            .into_iter()
            .next()
            .map(|paragraph| paragraph.text)
            .unwrap_or_default();
        replace_plain_run(paragraph, edited.clone(), template);
        let replacement = paragraph.clone();
        replace_section_paragraphs(&mut self.model.body, para_id, &replacement);
        if self.dirty == Dirty::Clean {
            self.dirty = Dirty::LoggedEdits;
        }
        self.authored_ids.retain(|_, (fresh, _)| fresh != para_id);
        let mirrored = match (&mut self.session_baseline, replacement.source_ordinal) {
            (Some(baseline), Some(ordinal)) if context.suggesting.is_none() => baseline
                .replace_text(para_id, &edited)
                .then(|| SessionEdit {
                    para_id: para_id.to_owned(),
                    ordinal,
                    text: text.to_owned(),
                    expected: edited,
                    context: context.clone(),
                }),
            _ => None,
        };
        match mirrored {
            Some(edit) => self.session_edits.push(edit),
            None => self.session_baseline = None,
        }
        Ok(receipt)
    }

    /// Exports the current model as read-only structured content. Anchors address the returned
    /// snapshot; source-part provenance is kept only where the opened package still matches.
    pub fn export_structured(&self, options: &ExportOptions) -> Result<DocxStructuredContent> {
        Ok(export_package_structured(
            self.envelope(),
            &self.original_parts,
            options,
        )?)
    }

    /// [`Document::export_structured`] rendered as Markdown.
    pub fn export_markdown(&self, options: &ExportOptions) -> Result<MarkdownContent> {
        let content = self.export_structured(options)?;
        render_docx_markdown(
            &content,
            &MarkdownOptions {
                max_bytes: options.max_bytes,
            },
        )
    }

    /// Lists the content controls of the current model in document order. Control ids and
    /// anchors address the returned snapshot; controls whose source content no longer matches
    /// the opened package cannot be shown fillable. An editing session fills them.
    pub fn list_content_controls(
        &self,
        options: &ContentControlsOptions,
    ) -> Result<ContentControlsSnapshot> {
        Ok(list_package_content_controls(
            self.envelope(),
            &self.original_parts,
            options,
        )?)
    }

    /// The content controls of the current model that match `query` exactly.
    pub fn find_content_controls(
        &self,
        query: &ContentControlQuery,
        options: &ContentControlsOptions,
    ) -> Result<ContentControlsSnapshot> {
        Ok(find_package_content_controls(
            self.envelope(),
            &self.original_parts,
            query,
            options,
        )?)
    }

    /// The current model as the parsed package seeding reads.
    fn envelope(&self) -> S9WireEnvelope {
        let model = self.model.clone();
        S9WireEnvelope {
            wire_version: 1,
            document: S9DocumentWire {
                package: S9PackageWire {
                    document: S9DocumentBodyWire::from(self.with_authored_ids(model.body)),
                    styles: model.styles,
                    theme: model.theme,
                    numbering: model.numbering,
                    settings: model.settings,
                    font_table: model.font_table,
                    header_entries: Some(model.headers),
                    footer_entries: Some(model.footers),
                    footnotes: Some(model.footnotes),
                    endnotes: Some(model.endnotes),
                    footnote_separators: Some(model.footnote_separators),
                    endnote_separators: Some(model.endnote_separators),
                    relationship_entries: model.relationships,
                    media_entries: Vec::new(),
                    chart_entries: model.charts,
                },
                template_variables: None,
                warnings: (!model.warnings.is_empty()).then_some(model.warnings),
            },
            embedded_font_parts: Vec::new(),
            font_table_relationships_xml: None,
            canonical_base64: None,
            canonical_sha256: None,
        }
    }

    #[cfg(feature = "fonts")]
    pub fn layout(&self, mut input: LayoutInput) -> Result<LayoutResult> {
        let requirements =
            self.font_requirements(input.measured.iter().map(|measured| &measured.block));
        let mut fonts = self.fonts.lock().unwrap_or_else(PoisonError::into_inner);
        fonts.resolve(&self.font_manager, &requirements)?;
        if fonts.chains.is_empty() {
            let layout = docx_layout::compute_layout_input(&mut input)?;
            let display_list =
                docx_layout::build_display_list(&input, &layout).map_err(Error::DisplayList)?;
            return Ok(LayoutResult {
                layout,
                display_list,
            });
        }
        let chains: BTreeMap<String, Vec<u32>> = fonts
            .chains
            .iter()
            .map(|(key, ids)| (key.clone(), ids.iter().map(|id| id.to_u32()).collect()))
            .collect();
        let package = serde_json::json!({
            "settings": { "compatibilityFlags": &self.model.settings.compatibility_flags }
        });
        let mut config: docx_layout::measure_blocks::MeasurementConfig =
            serde_json::from_value(docx_edit::layout_request::measurement(&package))
                .map_err(|error| Error::Render(error.to_string()))?;
        config.font_chains = chains.clone();
        let widths = crate::font::measurement_widths(&input)?;
        docx_layout::with_measure_font_store(&mut fonts.store, || {
            let has_floats = input
                .measured
                .iter()
                .zip(&widths)
                .any(|(measured, (width, _))| has_floating_zones(&measured.block, *width, &config));
            if has_floats {
                return;
            }
            for (measured, (width, remeasure)) in input.measured.iter_mut().zip(widths) {
                if !remeasure {
                    continue;
                }
                let mut block = measured.block.clone();
                if let Ok(measure) =
                    docx_layout::measure_blocks::measure_block(&mut block, width, &config)
                    && !docx_layout::measure_blocks::measured_synthetically([&measure])
                {
                    measured.block = block;
                    measured.measure = measure;
                }
            }
        });
        let layout = docx_layout::compute_layout_input(&mut input)?;
        let extras = serde_json::json!({ "fontChains": chains }).to_string();
        let display_list =
            docx_layout::display_list::build_display_list_value_from_resident_with_fonts(
                &input,
                &layout,
                &extras,
                &fonts.store,
            )
            .map_err(Error::DisplayList)?;
        Ok(LayoutResult {
            layout,
            display_list,
        })
    }

    /// The fonts `blocks` need, with the document's own CJK script first as the
    /// editing engine orders them for the browser and `render_png`.
    #[cfg(feature = "fonts")]
    fn font_requirements<'a>(
        &self,
        blocks: impl IntoIterator<Item = &'a docx_layout::types::LayoutBlock>,
    ) -> Vec<docx_layout::measure_blocks::FontRequirement> {
        let mut collector = docx_layout::measure_blocks::FontRequirementCollector::default();
        collector.prefer_east_asian_language(self.default_east_asian_language());
        collector.collect(blocks, "Calibri");
        collector.finish().into_values().collect()
    }

    #[cfg(feature = "fonts")]
    fn default_east_asian_language(&self) -> Option<&str> {
        self.model
            .styles
            .as_ref()?
            .doc_defaults
            .as_ref()?
            .r_pr
            .as_ref()?
            .language
            .as_ref()?
            .east_asia
            .as_deref()
    }

    #[cfg(not(feature = "fonts"))]
    pub fn layout(&self, mut input: LayoutInput) -> Result<LayoutResult> {
        let layout = docx_layout::compute_layout_input(&mut input)?;
        let display_list =
            docx_layout::build_display_list(&input, &layout).map_err(Error::DisplayList)?;
        Ok(LayoutResult {
            layout,
            display_list,
        })
    }

    pub fn save(&self) -> Result<Vec<u8>> {
        self.save_with_options(SaveOptions::default())
    }

    pub fn save_with_options(&self, options: SaveOptions) -> Result<Vec<u8>> {
        let determinism = SerializerDeterminism {
            seed: self.seed.clone(),
            now: options.now,
        };
        if self.dirty == Dirty::Clean
            && !options.update_modified_date
            && options.modified_by.is_none()
        {
            determinism.validate()?;
            return Ok(self.source.clone());
        }
        let request = S13SaveRequest {
            determinism,
            document: self.with_authored_ids(self.model.body.clone()),
            header_entries: self.model.headers.clone(),
            footer_entries: self.model.footers.clone(),
            footnotes: self.model.footnotes.clone(),
            endnotes: self.model.endnotes.clone(),
            footnote_separators: self.model.footnote_separators.clone(),
            endnote_separators: self.model.endnote_separators.clone(),
            relationship_entries: self.model.relationships.clone(),
            numbering: Some(self.model.numbering.clone()),
            options: S13SaveOptions {
                update_modified_date: options.update_modified_date,
                modified_by: options.modified_by,
            },
            selective: None,
            paragraph_ids: self.session_save_plan()?,
        };
        #[cfg(test)]
        session_save_tests::count_package_write();
        write_docx_s13_parts(request, &self.original_parts, None).map_err(Error::from)
    }

    /// The editing session's paragraph save plan for the current model, as the
    /// product's session save writes it; `None` saves the whole model.
    fn session_save_plan(&self) -> Result<Option<S13ParagraphIds>> {
        if self.dirty == Dirty::ModelChanged {
            return Ok(None);
        }
        let Some(baseline) = &self.session_baseline else {
            return Ok(None);
        };
        if !baseline.saves_as(&self.model) {
            return Ok(None);
        }
        let Some(reassigned) = self.reassigned_ids() else {
            return Ok(None);
        };
        let session = EditingDoc::new(1);
        #[cfg(test)]
        session_save_tests::count_session_seed();
        let seeded = seed_from_docx_with_generation(&session, &self.source, &self.seed[..32]);
        #[cfg(test)]
        let seeded = seeded.and_then(|()| session_save_tests::injected_seed_error());
        if seeded.is_err() {
            return Ok(None);
        }
        session.prime_save_seed_states(SAVE_SEED_MAX_PARAGRAPHS, SAVE_SEED_MAX_INFLATED_BYTES);
        let identities = session.paragraph_identities();
        let body_part = body_part(&identities);
        let Some(seeded) = session_paragraphs(&session, &body_part) else {
            return Ok(None);
        };
        let replayed = self
            .session_edits
            .iter()
            .all(|edit| edit.replay(&session, &identities, &body_part).is_ok());
        if !replayed {
            return Ok(None);
        }
        match session_paragraphs(&session, &body_part) {
            Some(replayed) if self.session_agrees(baseline, &seeded, &replayed) => {}
            _ => return Ok(None),
        }
        let mut plan = session_paragraph_ids(session.paragraph_save_plan());
        if !reassigned.is_empty() {
            if let Some(patched) = plan
                .patched_parts
                .iter_mut()
                .find(|patched| patched.part == body_part)
            {
                patched
                    .para_ids
                    .retain(|(ordinal, _)| !reassigned.iter().any(|(source, _)| source == ordinal));
                patched.para_ids.extend(reassigned);
                patched.para_ids.sort_by_key(|(ordinal, _)| *ordinal);
            } else if plan
                .spliced_parts
                .iter()
                .any(|spliced| spliced.part == body_part)
            {
                let moved: BTreeSet<u32> = reassigned.iter().map(|(ordinal, _)| *ordinal).collect();
                plan.assignments
                    .retain(|id| id.part != body_part || !moved.contains(&id.ordinal));
                for (ordinal, para_id) in reassigned {
                    plan.assignments.push(S13ParagraphId {
                        part: body_part.clone(),
                        ordinal,
                        para_id,
                    });
                }
            } else {
                return Ok(None);
            }
        }
        Ok(Some(plan))
    }

    /// Whether each changed body paragraph saves from the session as the model holds it.
    fn session_agrees(
        &self,
        baseline: &SavedFields,
        seeded: &SessionParagraphs,
        replayed: &SessionParagraphs,
    ) -> bool {
        if seeded.count != replayed.count || seeded.body.len() != replayed.body.len() {
            return false;
        }
        let mut model = BTreeMap::new();
        for paragraph in self.paragraphs() {
            if let Some(ordinal) = paragraph.source_ordinal
                && model.insert(ordinal, paragraph).is_some()
            {
                return false;
            }
        }
        let edited: BTreeSet<u32> = self
            .session_edits
            .iter()
            .map(|session_edit| session_edit.ordinal)
            .collect();
        if !edited.iter().all(|at| seeded.body.contains_key(at)) {
            return false;
        }
        let session = baseline.paragraphs_by_ordinal();
        seeded.body.iter().all(|(ordinal, before)| {
            let Some(after) = replayed.body.get(ordinal) else {
                return false;
            };
            if !same_properties(model.get(ordinal).copied(), session.get(ordinal).copied()) {
                return false;
            }
            if after == before && !edited.contains(ordinal) {
                return true;
            }
            model
                .get(ordinal)
                .is_some_and(|paragraph| get_paragraph_text(paragraph) == after.text)
                && after.keeps_formatting_of(before)
        })
    }

    /// `(source ordinal, ID)` of each body paragraph that saves with the fresh
    /// ID it was given on open; `None` when one has no source ordinal.
    fn reassigned_ids(&self) -> Option<Vec<(u32, String)>> {
        let paragraphs = self.paragraphs();
        self.opened_ids
            .iter()
            .filter(|(ordinal, _)| !self.authored_ids.contains_key(*ordinal))
            .filter_map(|(ordinal, (fresh, _))| {
                let paragraph = paragraphs.get(*ordinal)?;
                (paragraph.para_id.as_ref() == Some(fresh)).then(|| {
                    paragraph
                        .source_ordinal
                        .map(|source| (source, fresh.clone()))
                })
            })
            .collect()
    }

    /// `body` as it saves: a paragraph still at its ordinal with its fresh ID
    /// keeps its authored ID.
    fn with_authored_ids(&self, mut body: DocumentBody) -> DocumentBody {
        if !self.authored_ids.is_empty() {
            let mut ordinal = 0usize;
            visit_paragraphs_mut(&mut body.content, &mut |paragraph| {
                ordinal += 1;
                let Some((fresh, authored)) = self.authored_ids.get(&(ordinal - 1)) else {
                    return;
                };
                if paragraph.para_id.as_ref() == Some(fresh) {
                    let paragraph = Arc::make_mut(paragraph);
                    paragraph.para_id = Some(authored.clone());
                    paragraph.repeated_para_id = Some(true);
                }
            });
        }
        body
    }
}

/// The model fields a save writes.
struct SavedFields {
    body: DocumentBody,
    headers: Vec<(String, docx_parse::HeaderFooter)>,
    footers: Vec<(String, docx_parse::HeaderFooter)>,
    footnotes: Vec<Note>,
    endnotes: Vec<Note>,
    footnote_separators: Vec<Note>,
    endnote_separators: Vec<Note>,
    relationships: Vec<(String, Relationship)>,
    numbering: NumberingDefinitions,
}

impl SavedFields {
    fn of(model: &DocumentModel) -> Self {
        Self {
            body: model.body.clone(),
            headers: model.headers.clone(),
            footers: model.footers.clone(),
            footnotes: model.footnotes.clone(),
            endnotes: model.endnotes.clone(),
            footnote_separators: model.footnote_separators.clone(),
            endnote_separators: model.endnote_separators.clone(),
            relationships: model.relationships.clone(),
            numbering: model.numbering.clone(),
        }
    }

    fn saves_as(&self, model: &DocumentModel) -> bool {
        self.body == model.body
            && self.headers == model.headers
            && self.footers == model.footers
            && self.footnotes == model.footnotes
            && self.endnotes == model.endnotes
            && self.footnote_separators == model.footnote_separators
            && self.endnote_separators == model.endnote_separators
            && self.relationships == model.relationships
            && self.numbering == model.numbering
    }

    /// Replaces only the text of the paragraph with `para_id`, as the logged edit did.
    fn replace_text(&mut self, para_id: &str, text: &str) -> bool {
        let Some(paragraph) = find_paragraph_mut(&mut self.body.content, para_id) else {
            return false;
        };
        let Some(template) = plain_run_template(paragraph) else {
            return false;
        };
        replace_plain_run(paragraph, text.to_owned(), template);
        let replacement = paragraph.clone();
        replace_section_paragraphs(&mut self.body, para_id, &replacement);
        true
    }

    /// The body paragraphs by source ordinal.
    fn paragraphs_by_ordinal(&self) -> BTreeMap<u32, &Paragraph> {
        let mut paragraphs = Vec::new();
        collect_paragraphs(&self.body.content, &mut paragraphs);
        paragraphs
            .into_iter()
            .filter_map(|paragraph| paragraph.source_ordinal.map(|ordinal| (ordinal, paragraph)))
            .collect()
    }
}

/// A text replacement the save replays into the editing session.
struct SessionEdit {
    para_id: String,
    /// The paragraph's `w:p` occurrence in the body part.
    ordinal: u32,
    text: String,
    expected: String,
    context: EditCtx,
}

impl SessionEdit {
    fn replay(
        &self,
        session: &EditingDoc,
        identities: &ParagraphIdentities,
        body_part: &str,
    ) -> Result<()> {
        let (story, key) = identities
            .paragraphs
            .iter()
            .find_map(|identity| match (&identity.paragraph, &identity.source) {
                (ParagraphRef::Session { story, para_id }, Some(source))
                    if source.paragraph_ordinal == self.ordinal
                        && part_path(&source.part_uri) == body_part =>
                {
                    Some((story, para_id))
                }
                _ => None,
            })
            .ok_or_else(|| Error::ParagraphNotFound(self.para_id.clone()))?;
        let range = story_paragraph_range(session, story, key)?
            .ok_or_else(|| Error::ParagraphNotFound(self.para_id.clone()))?;
        session.replace_range(&self.context, range, &self.text)?;
        let replayed = session
            .paragraphs(story)?
            .into_iter()
            .find(|paragraph| paragraph.para_id == *key)
            .map(|paragraph| paragraph.text);
        if replayed.as_deref() != Some(self.expected.as_str()) {
            return Err(Error::UnsupportedParagraphEdit(self.para_id.clone()));
        }
        Ok(())
    }
}

/// An editing session's paragraphs: how many, and the body part's by source ordinal.
struct SessionParagraphs {
    count: usize,
    body: BTreeMap<u32, SessionParagraph>,
}

#[derive(PartialEq)]
struct SessionParagraph {
    text: String,
    runs: Vec<StorySegment>,
}

impl SessionParagraph {
    fn new(runs: Vec<StorySegment>) -> Self {
        let text: String = runs
            .iter()
            .filter_map(|run| match &run.content {
                SegmentContent::Text(text) => Some(text.as_str()),
                _ => None,
            })
            .collect();
        Self { text, runs }
    }

    /// Every run carries the one formatting `before` had, none when it had no text.
    fn keeps_formatting_of(&self, before: &SessionParagraph) -> bool {
        match before.runs.first() {
            Some(first) => before
                .runs
                .iter()
                .chain(&self.runs)
                .all(|run| run.attributes == first.attributes),
            None => self.runs.iter().all(|run| run.attributes.is_empty()),
        }
    }
}

/// Whether both hold the same paragraph properties and the same runs, run properties included.
fn same_properties(model: Option<&Paragraph>, session: Option<&Paragraph>) -> bool {
    match (model, session) {
        (Some(model), Some(session)) => {
            model.formatting == session.formatting && model.content == session.content
        }
        (model, session) => model.is_none() && session.is_none(),
    }
}

fn session_paragraphs(session: &EditingDoc, body_part: &str) -> Option<SessionParagraphs> {
    let identities = session.paragraph_identities();
    let mut stories: BTreeMap<String, BTreeMap<String, SessionParagraph>> = BTreeMap::new();
    let mut paragraphs = SessionParagraphs {
        count: 0,
        body: BTreeMap::new(),
    };
    for identity in &identities.paragraphs {
        let ParagraphRef::Session { story, para_id } = &identity.paragraph else {
            continue;
        };
        paragraphs.count += 1;
        let Some(source) = identity
            .source
            .as_ref()
            .filter(|source| part_path(&source.part_uri) == body_part)
        else {
            continue;
        };
        let in_story = match stories.entry(story.clone()) {
            Entry::Occupied(entry) => entry.into_mut(),
            Entry::Vacant(entry) => entry.insert(story_paragraphs(session, story).ok()?),
        };
        let paragraph = in_story.remove(para_id)?;
        let ordinal = source.paragraph_ordinal;
        if paragraphs.body.insert(ordinal, paragraph).is_some() {
            return None;
        }
    }
    Some(paragraphs)
}

/// The text runs of each paragraph in `story`, by pilcrow key.
fn story_paragraphs(
    session: &EditingDoc,
    story: &str,
) -> Result<BTreeMap<String, SessionParagraph>> {
    let mut paragraphs = BTreeMap::new();
    let mut runs = Vec::new();
    for segment in session.story_segments(story)? {
        if matches!(segment.content, SegmentContent::Text(_)) {
            runs.push(segment);
        } else if let SegmentContent::Pilcrow(properties) = segment.content {
            paragraphs.insert(
                properties.para_id,
                SessionParagraph::new(std::mem::take(&mut runs)),
            );
        }
    }
    Ok(paragraphs)
}

/// The content of the paragraph closed by the pilcrow `key` in `story`.
fn story_paragraph_range(
    session: &EditingDoc,
    story: &str,
    key: &str,
) -> Result<Option<StoryRange>> {
    let (mut start, mut offset) = (0, 0);
    for segment in session.story_segments(story)? {
        match segment.content {
            SegmentContent::Text(text) => offset += utf16_len(&text),
            SegmentContent::Pilcrow(properties) if properties.para_id == key => {
                return Ok(Some(StoryRange::new(story, start, offset)));
            }
            SegmentContent::Pilcrow(_) => {
                offset += 1;
                start = offset;
            }
            SegmentContent::OtherEmbed { .. } => offset += 1,
        }
    }
    Ok(None)
}

fn part_path(uri: &str) -> &str {
    uri.strip_prefix('/').unwrap_or(uri)
}

fn body_part(identities: &ParagraphIdentities) -> String {
    identities
        .paragraphs
        .iter()
        .find_map(|identity| match (&identity.paragraph, &identity.source) {
            (ParagraphRef::Session { story, .. }, Some(source)) if story == "body" => {
                Some(part_path(&source.part_uri).to_owned())
            }
            _ => None,
        })
        .unwrap_or_else(|| "word/document.xml".to_owned())
}

/// `plan` as the package writer's request field, every planned part kept.
fn session_paragraph_ids(plan: ParagraphSavePlan) -> S13ParagraphIds {
    S13ParagraphIds {
        removed_comment_ids: Vec::new(),
        section_edits: None,
        assignments: plan
            .assignments
            .into_iter()
            .map(|(part, ordinal, para_id)| S13ParagraphId {
                part,
                ordinal,
                para_id,
            })
            .collect(),
        patched_parts: plan
            .patched_parts
            .into_iter()
            .map(|(part, para_ids)| S13PatchedPart { part, para_ids })
            .collect(),
        spliced_parts: plan
            .spliced_parts
            .into_iter()
            .map(|part| S13SplicedPart {
                part: part.part,
                sha256: part.sha256,
                paragraphs: part
                    .paragraphs
                    .into_iter()
                    .map(|(ordinal, _)| ordinal)
                    .collect(),
                changed: part.changed,
                inserted: part
                    .inserted
                    .into_iter()
                    .map(|anchor| match anchor {
                        SpliceAnchor::Before(ordinal) => S13SpliceAnchor::Before(ordinal),
                        SpliceAnchor::After(ordinal) => S13SpliceAnchor::After(ordinal),
                    })
                    .collect(),
                removed: part.removed,
                paragraph_patches: part.paragraph_patches,
                legacy_paragraphs: part.legacy_paragraphs,
                table_patches: part.table_patches,
                format_only: part.format_only,
            })
            .collect(),
        whole_parts: plan
            .whole_parts
            .into_iter()
            .map(|part| S13WholePart {
                part: part.part,
                sha256: part.sha256,
                changed: part.changed,
            })
            .collect(),
        touched_parts: plan
            .touched_parts
            .into_iter()
            .map(|(part, touched)| S13TouchedPart { part, touched })
            .collect(),
    }
}

fn replace_section_paragraphs(body: &mut DocumentBody, para_id: &str, replacement: &Paragraph) {
    for section in body.sections.iter_mut().flatten() {
        if let Some(paragraph) = find_paragraph_mut(&mut section.content, para_id) {
            *paragraph = replacement.clone();
        }
    }
}

/// Gives each paragraph that repeats an earlier paragraph's ID a fresh one no
/// package part uses, returning the fresh and authored IDs by paragraph ordinal.
fn address_repeated_paragraphs(
    blocks: &mut [BlockContent],
    parts: &[(String, Vec<u8>)],
) -> BTreeMap<usize, (String, String)> {
    let mut authored_ids = BTreeMap::new();
    let mut occupied = None;
    let mut next_ordinal = 0usize;
    visit_paragraphs_mut(blocks, &mut |paragraph| {
        let ordinal = next_ordinal;
        next_ordinal += 1;
        if paragraph.repeated_para_id != Some(true) {
            return;
        }
        let occupied = occupied.get_or_insert_with(|| package_paragraph_ids(parts));
        let Some(id) = allocate_paragraph_id(&format!("body#{ordinal}"), occupied) else {
            return;
        };
        occupied.insert(id);
        let paragraph = Arc::make_mut(paragraph);
        let fresh = format_paragraph_id(id);
        if let Some(authored) = paragraph.para_id.replace(fresh.clone()) {
            authored_ids.insert(ordinal, (fresh, authored));
        }
        paragraph.repeated_para_id = None;
    });
    authored_ids
}

fn visit_paragraphs_mut(blocks: &mut [BlockContent], visit: &mut impl FnMut(&mut Arc<Paragraph>)) {
    for block in blocks {
        match block {
            BlockContent::Paragraph(paragraph) => visit(paragraph),
            BlockContent::Table(table) => {
                for row in &mut Arc::make_mut(table).rows {
                    for cell in &mut row.cells {
                        visit_paragraphs_mut(&mut cell.content, visit);
                    }
                }
            }
            BlockContent::BlockSdt(sdt) => {
                visit_paragraphs_mut(&mut Arc::make_mut(sdt).content, visit)
            }
            BlockContent::RawXml(_) => {}
        }
    }
}

/// Renders structured content as Markdown within `options.max_bytes`.
pub fn render_docx_markdown(
    content: &DocxStructuredContent,
    options: &MarkdownOptions,
) -> Result<MarkdownContent> {
    Ok(docx_edit::structured::render_docx_markdown(
        content, options,
    )?)
}

type RunTemplate = Option<(
    Option<docx_parse::TextFormatting>,
    Option<Vec<docx_parse::inline::RunPropertyChange>>,
)>;

fn plain_run_template(paragraph: &Paragraph) -> Option<RunTemplate> {
    if paragraph.content.is_empty() {
        return Some(None);
    }
    let [ParagraphContent::Inline(InlineNode::Run(run))] = paragraph.content.as_slice() else {
        return None;
    };
    if !run
        .content
        .iter()
        .all(|content| matches!(content, RunContent::Text { .. }))
    {
        return None;
    }
    Some(Some((run.formatting.clone(), run.property_changes.clone())))
}

fn replace_plain_run(paragraph: &mut Paragraph, text: String, template: RunTemplate) {
    if text.is_empty() {
        paragraph.content.clear();
        return;
    }
    let (formatting, property_changes) = template.unwrap_or_default();
    paragraph.content = vec![ParagraphContent::Inline(InlineNode::Run(Run {
        node_type: RunType::Run,
        formatting,
        property_changes,
        content: vec![RunContent::Text {
            preserve_space: (text.trim() != text).then_some(true),
            text,
        }],
    }))];
}

fn model_from_package(
    package: S9PackageWire,
    template_variables: Vec<String>,
    warnings: Vec<String>,
) -> DocumentModel {
    let S9PackageWire {
        document,
        styles,
        theme,
        numbering,
        settings,
        font_table,
        header_entries,
        footer_entries,
        footnotes,
        endnotes,
        footnote_separators,
        endnote_separators,
        relationship_entries,
        media_entries,
        chart_entries,
    } = package;
    DocumentModel {
        body: body_from_wire(document),
        styles,
        theme,
        numbering,
        settings,
        font_table,
        headers: header_entries.unwrap_or_default(),
        footers: footer_entries.unwrap_or_default(),
        footnotes: footnotes.unwrap_or_default(),
        endnotes: endnotes.unwrap_or_default(),
        footnote_separators: footnote_separators.unwrap_or_default(),
        endnote_separators: endnote_separators.unwrap_or_default(),
        relationships: relationship_entries,
        media: media_entries
            .into_iter()
            .map(|(key, file)| {
                (
                    key,
                    Arc::try_unwrap(file).unwrap_or_else(|file| (*file).clone()),
                )
            })
            .collect(),
        charts: chart_entries,
        template_variables,
        warnings,
    }
}

fn body_from_wire(wire: S9DocumentBodyWire) -> DocumentBody {
    let sections = wire.sections.map(|sections| {
        sections
            .into_iter()
            .map(|section| section_from_wire(section, &wire.content))
            .collect()
    });
    DocumentBody {
        content: wire.content,
        sections,
        final_section_properties: wire.final_section_properties,
        custom_root_bindings: wire.custom_root_bindings,
        comments: wire.comments,
    }
}

fn section_from_wire(section: S9SectionWire, content: &[BlockContent]) -> Section {
    Section {
        id: section.id,
        properties: section.properties,
        content: content
            .get(section.content_start..section.content_end)
            .unwrap_or_default()
            .to_vec(),
    }
}

fn collect_paragraphs<'a>(blocks: &'a [BlockContent], output: &mut Vec<&'a Paragraph>) {
    for block in blocks {
        match block {
            BlockContent::Paragraph(paragraph) => output.push(paragraph),
            BlockContent::Table(table) => {
                for row in &table.rows {
                    for cell in &row.cells {
                        collect_paragraphs(&cell.content, output);
                    }
                }
            }
            BlockContent::BlockSdt(sdt) => collect_paragraphs(&sdt.content, output),
            BlockContent::RawXml(_) => {}
        }
    }
}

fn collect_tables<'a>(blocks: &'a [BlockContent], output: &mut Vec<&'a Table>) {
    for block in blocks {
        match block {
            BlockContent::Paragraph(_) => {}
            BlockContent::Table(table) => {
                output.push(table);
                for row in &table.rows {
                    for cell in &row.cells {
                        collect_tables(&cell.content, output);
                    }
                }
            }
            BlockContent::BlockSdt(sdt) => collect_tables(&sdt.content, output),
            BlockContent::RawXml(_) => {}
        }
    }
}

fn find_paragraph_mut<'a>(
    blocks: &'a mut [BlockContent],
    para_id: &str,
) -> Option<&'a mut Paragraph> {
    for block in blocks {
        match block {
            BlockContent::Paragraph(paragraph) if paragraph.para_id.as_deref() == Some(para_id) => {
                return Some(Arc::make_mut(paragraph));
            }
            BlockContent::Paragraph(_) => {}
            BlockContent::Table(table) => {
                for row in &mut Arc::make_mut(table).rows {
                    for cell in &mut row.cells {
                        if let Some(paragraph) = find_paragraph_mut(&mut cell.content, para_id) {
                            return Some(paragraph);
                        }
                    }
                }
            }
            BlockContent::BlockSdt(sdt) => {
                if let Some(paragraph) =
                    find_paragraph_mut(&mut Arc::make_mut(sdt).content, para_id)
                {
                    return Some(paragraph);
                }
            }
            BlockContent::RawXml(_) => {}
        }
    }
    None
}

fn utf16_len(text: &str) -> u32 {
    text.encode_utf16().count() as u32
}

#[cfg(feature = "fonts")]
fn has_floating_zones(
    block: &docx_layout::types::LayoutBlock,
    width: f64,
    config: &docx_layout::measure_blocks::MeasurementConfig,
) -> bool {
    use docx_layout::types::LayoutBlock;
    if docx_layout::measure_blocks::has_floating_zones(
        std::slice::from_ref(block),
        width,
        config,
        None,
    )
    .unwrap_or(true)
    {
        return true;
    }
    match block {
        LayoutBlock::Table(table) => table
            .rows
            .iter()
            .flat_map(|row| &row.cells)
            .flat_map(|cell| &cell.blocks)
            .any(|block| has_floating_zones(block, width, config)),
        LayoutBlock::TextBox(text_box) => text_box.content.iter().any(|paragraph| {
            has_floating_zones(&LayoutBlock::Paragraph(paragraph.clone()), width, config)
        }),
        _ => false,
    }
}

#[cfg(test)]
mod session_save_tests {
    use super::{
        Document, EditCtx, InlineNode, Paragraph, ParagraphContent, SessionEdit,
        find_paragraph_mut, get_paragraph_text, replace_plain_run, replace_section_paragraphs,
    };

    use std::cell::{Cell, RefCell};

    use docx_parse::{TabStop, TextFormatting};

    thread_local! {
        static SEED_ERROR: RefCell<Option<String>> = const { RefCell::new(None) };
        static SESSION_SEEDS: Cell<usize> = const { Cell::new(0) };
        static PACKAGE_WRITES: Cell<usize> = const { Cell::new(0) };
    }

    pub(super) fn injected_seed_error() -> Result<(), String> {
        SEED_ERROR.with(|error| error.borrow().clone().map_or(Ok(()), Err))
    }

    pub(super) fn count_session_seed() {
        SESSION_SEEDS.with(|seeds| seeds.set(seeds.get() + 1));
    }

    pub(super) fn count_package_write() {
        PACKAGE_WRITES.with(|writes| writes.set(writes.get() + 1));
    }

    /// `(session seeds, package writes)` during `run`.
    fn counted(run: impl FnOnce()) -> (usize, usize) {
        SESSION_SEEDS.with(|seeds| seeds.set(0));
        PACKAGE_WRITES.with(|writes| writes.set(0));
        run();
        (
            SESSION_SEEDS.with(|seeds| seeds.get()),
            PACKAGE_WRITES.with(|writes| writes.get()),
        )
    }

    const NAMESPACES: &str = r#"xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" mc:Ignorable="w14""#;

    fn paragraph(id: &str, content: &str) -> String {
        format!(
            r#"<w:p w14:paraId="{id}" w14:textId="77777777" w:rsidR="00A1B2C3" w:rsidRDefault="00A1B2C3"><w:pPr><w:spacing w:after="120"/></w:pPr>{content}</w:p>"#
        )
    }

    fn run(text: &str) -> String {
        format!(r#"<w:r w:rsidRPr="00D4E5F6"><w:t xml:space="preserve">{text}</w:t></w:r>"#)
    }

    fn source() -> String {
        document_xml(
            &[
                first_paragraph(),
                paragraph("0000000B", &run("Charlie")),
                paragraph("0000000C", &run("Delta")),
            ]
            .concat(),
        )
    }

    fn document_xml(body: &str) -> String {
        format!(
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n<w:document {NAMESPACES}><w:body>{body}<w:sectPr><w:pgSz w:w=\"12240\" w:h=\"15840\"/></w:sectPr></w:body></w:document>"
        )
    }

    fn open() -> Document {
        open_with(&source())
    }

    fn open_with(document_xml: &str) -> Document {
        Document::open(&package(document_xml)).unwrap()
    }

    fn package(document_xml: &str) -> Vec<u8> {
        let parts = vec![
            (
                "[Content_Types].xml".to_owned(),
                br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>"#.to_vec(),
            ),
            (
                "_rels/.rels".to_owned(),
                br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"#.to_vec(),
            ),
            ("word/document.xml".to_owned(), document_xml.as_bytes().to_vec()),
        ];
        ooxml_opc::rezip_parts(&parts).unwrap()
    }

    fn reopened(document: &Document) -> Document {
        Document::open(&document.save().unwrap()).unwrap()
    }

    fn change_model(document: &mut Document, para_id: &str, change: impl FnOnce(&mut Paragraph)) {
        let body = &mut document.model_mut().body;
        let paragraph = find_paragraph_mut(&mut body.content, para_id).unwrap();
        change(&mut *paragraph);
        let replacement = paragraph.clone();
        replace_section_paragraphs(body, para_id, &replacement);
    }

    fn center(paragraph: &mut Paragraph) {
        let formatting = paragraph.formatting.get_or_insert_with(Default::default);
        formatting.alignment = Some("center".to_owned());
    }

    fn tab_stop(paragraph: &mut Paragraph) {
        let formatting = paragraph.formatting.get_or_insert_with(Default::default);
        formatting.tabs = Some(vec![TabStop {
            position: 4320.0,
            alignment: "center".to_owned(),
            leader: None,
        }]);
    }

    fn alignment(paragraph: &Paragraph) -> Option<&str> {
        paragraph.formatting.as_ref()?.alignment.as_deref()
    }

    fn bold_run(paragraph: &mut Paragraph) {
        let Some(ParagraphContent::Inline(InlineNode::Run(run))) = paragraph.content.first_mut()
        else {
            panic!("paragraph without a first run");
        };
        run.formatting = Some(TextFormatting {
            bold: Some(true),
            ..Default::default()
        });
    }

    fn run_bold(paragraph: &Paragraph) -> Option<bool> {
        match paragraph.content.first() {
            Some(ParagraphContent::Inline(InlineNode::Run(run))) => run.formatting.as_ref()?.bold,
            _ => None,
        }
    }

    fn main_part(document: &Document) -> String {
        let parts = ooxml_opc::unzip_parts(&document.save().unwrap()).unwrap();
        let (_, bytes) = parts
            .into_iter()
            .find(|(name, _)| name == "word/document.xml")
            .unwrap();
        String::from_utf8(bytes).unwrap()
    }

    fn full_write(mut document: Document) -> String {
        document.session_baseline = None;
        main_part(&document)
    }

    fn assert_plan_keeps(document: &Document, unchanged: &[&str], edited: &str) {
        assert!(document.session_save_plan().unwrap().is_some());
        let saved = main_part(document);
        assert!(saved.contains(edited), "{saved}");
        let source = source();
        for paragraph in unchanged {
            assert!(
                source.contains(*paragraph) && saved.contains(*paragraph),
                "{saved}"
            );
        }
    }

    fn logged(document: &Document, index: usize, text: &str) -> SessionEdit {
        let paragraph = document.paragraphs()[index];
        SessionEdit {
            para_id: paragraph.para_id.clone().unwrap(),
            ordinal: paragraph.source_ordinal.unwrap(),
            text: text.to_owned(),
            expected: text.to_owned(),
            context: EditCtx::local("betteroffice-docx", "1970-01-01T00:00:00.000Z"),
        }
    }

    fn first_paragraph() -> String {
        paragraph(
            "0000000A",
            &format!(
                r#"{}<w:proofErr w:type="spellStart"/>{}<w:proofErr w:type="spellEnd"/>"#,
                run("Alpha "),
                run("bravvo")
            ),
        )
    }

    #[test]
    fn replace_paragraph_text_saves_through_the_plan() {
        let mut document = open();
        document
            .replace_paragraph_text("0000000B", "Charlie edited")
            .unwrap();
        let unchanged = [first_paragraph(), paragraph("0000000C", &run("Delta"))];
        let unchanged: Vec<&str> = unchanged.iter().map(String::as_str).collect();
        assert_plan_keeps(&document, &unchanged, "Charlie edited");
    }

    #[test]
    fn replace_paragraph_text_with_saves_through_the_plan() {
        let mut document = open();
        let context = EditCtx::local("Reviewer", "2026-01-02T03:04:05.000Z");
        document
            .replace_paragraph_text_with("0000000C", "Delta edited", 7, &context)
            .unwrap();
        document
            .replace_paragraph_text_with("0000000C", "Delta edited twice", 7, &context)
            .unwrap();
        let unchanged = [first_paragraph(), paragraph("0000000B", &run("Charlie"))];
        let unchanged: Vec<&str> = unchanged.iter().map(String::as_str).collect();
        assert_plan_keeps(&document, &unchanged, "Delta edited twice");
    }

    #[test]
    fn an_unedited_save_builds_no_session() {
        let document = open();
        let counts = counted(|| {
            document.save().unwrap();
        });
        assert_eq!(counts, (0, 0));
    }

    #[test]
    fn an_unedited_save_writes_every_source_entry_as_it_is() {
        let source = package(&source());
        let saved = Document::open(&source).unwrap().save().unwrap();
        assert_eq!(
            ooxml_opc::unzip_parts(&saved).unwrap(),
            ooxml_opc::unzip_parts(&source).unwrap()
        );
        assert_eq!(saved, source);
    }

    #[test]
    fn known_limit_model_mut_borrow_without_change_takes_the_full_write() {
        let mut document = open();
        document.model_mut();
        let counts = counted(|| {
            document.save().unwrap();
        });
        assert_eq!(counts, (0, 1));
    }

    #[test]
    fn a_logged_edit_saves_through_one_session() {
        let mut document = open();
        document
            .replace_paragraph_text("0000000B", "Charlie edited")
            .unwrap();
        let counts = counted(|| {
            document.save().unwrap();
        });
        assert_eq!(counts, (1, 1));
        let unchanged = [first_paragraph(), paragraph("0000000C", &run("Delta"))];
        let unchanged: Vec<&str> = unchanged.iter().map(String::as_str).collect();
        assert_plan_keeps(&document, &unchanged, "Charlie edited");
    }

    #[test]
    fn diverging_session_takes_the_full_write() {
        let mut document = open();
        document
            .replace_paragraph_text("0000000B", "Charlie edited")
            .unwrap();
        let planted = logged(&document, 2, "Planted");
        document.session_edits.push(planted);
        assert!(document.session_save_plan().unwrap().is_none());
        let saved = main_part(&document);
        assert!(!saved.contains("Planted"), "{saved}");
        let mut expected = open();
        expected
            .replace_paragraph_text("0000000B", "Charlie edited")
            .unwrap();
        assert_eq!(saved, full_write(expected));
    }

    #[test]
    fn failed_seed_takes_the_full_write() {
        let mut document = open();
        document
            .replace_paragraph_text("0000000B", "Charlie edited")
            .unwrap();
        SEED_ERROR.with(|error| {
            *error.borrow_mut() = Some("XML element limit exceeded".to_owned());
        });
        let plan = document.session_save_plan();
        let saved = main_part(&document);
        SEED_ERROR.with(|error| *error.borrow_mut() = None);
        assert!(plan.unwrap().is_none());
        let mut expected = open();
        expected
            .replace_paragraph_text("0000000B", "Charlie edited")
            .unwrap();
        assert_eq!(saved, full_write(expected));
    }

    #[test]
    fn failed_replay_takes_the_full_write() {
        let mut document = open();
        document
            .replace_paragraph_text("0000000B", "Charlie edited")
            .unwrap();
        let mut missing = logged(&document, 2, "Missing");
        missing.ordinal = 99;
        document.session_edits.push(missing);
        assert!(document.session_save_plan().unwrap().is_none());
        let mut expected = open();
        expected
            .replace_paragraph_text("0000000B", "Charlie edited")
            .unwrap();
        assert_eq!(main_part(&document), full_write(expected));
    }

    #[test]
    fn model_mut_edit_after_logged_edits_is_saved() {
        let mut document = open();
        document
            .replace_paragraph_text("0000000B", "Charlie edited")
            .unwrap();
        let body = &mut document.model_mut().body;
        let paragraph = find_paragraph_mut(&mut body.content, "0000000C").unwrap();
        replace_plain_run(paragraph, "Delta by model".to_owned(), None);
        let replacement = paragraph.clone();
        replace_section_paragraphs(body, "0000000C", &replacement);
        assert!(document.session_save_plan().unwrap().is_none());
        let saved = main_part(&document);
        assert!(saved.contains("Charlie edited"), "{saved}");
        assert!(saved.contains("Delta by model"), "{saved}");
    }

    #[test]
    fn model_mut_borrow_takes_the_full_write() {
        let mut document = open();
        document
            .replace_paragraph_text("0000000B", "Charlie edited")
            .unwrap();
        document.model_mut();
        assert!(document.session_save_plan().unwrap().is_none());
        let counts = counted(|| {
            document.save().unwrap();
        });
        assert_eq!(counts, (0, 1));
        let mut expected = open();
        expected
            .replace_paragraph_text("0000000B", "Charlie edited")
            .unwrap();
        assert_eq!(main_part(&document), full_write(expected));
    }

    #[test]
    fn model_alignment_then_logged_edit_on_one_paragraph_both_survive() {
        let mut document = open();
        change_model(&mut document, "0000000B", center);
        document
            .replace_paragraph_text("0000000B", "Charlie edited")
            .unwrap();
        assert!(document.session_save_plan().unwrap().is_none());
        let saved = reopened(&document);
        let edited = saved.paragraph("0000000B").unwrap();
        assert_eq!(get_paragraph_text(edited), "Charlie edited");
        assert_eq!(alignment(edited), Some("center"));
    }

    #[test]
    fn model_alignment_and_logged_edit_on_other_paragraphs_both_survive() {
        let mut document = open();
        change_model(&mut document, "0000000C", center);
        document
            .replace_paragraph_text("0000000B", "Charlie edited")
            .unwrap();
        assert!(document.session_save_plan().unwrap().is_none());
        let saved = reopened(&document);
        let changed = saved.paragraph("0000000C").unwrap();
        assert_eq!(get_paragraph_text(changed), "Delta");
        assert_eq!(alignment(changed), Some("center"));
        let edited = saved.paragraph("0000000B").unwrap();
        assert_eq!(get_paragraph_text(edited), "Charlie edited");
    }

    #[test]
    fn logged_edit_then_model_alignment_on_one_paragraph_both_survive() {
        let mut document = open();
        document
            .replace_paragraph_text("0000000B", "Charlie edited")
            .unwrap();
        change_model(&mut document, "0000000B", center);
        assert!(document.session_save_plan().unwrap().is_none());
        let saved = reopened(&document);
        let edited = saved.paragraph("0000000B").unwrap();
        assert_eq!(get_paragraph_text(edited), "Charlie edited");
        assert_eq!(alignment(edited), Some("center"));
    }

    fn bold_then_edit(document: &mut Document) {
        change_model(document, "0000000B", bold_run);
        document
            .replace_paragraph_text("0000000B", "Charlie edited")
            .unwrap();
    }

    #[test]
    fn model_run_property_then_logged_edit_saves_as_the_full_write() {
        let mut document = open();
        bold_then_edit(&mut document);
        assert!(document.session_save_plan().unwrap().is_none());
        let mut expected = open();
        bold_then_edit(&mut expected);
        assert_eq!(main_part(&document), full_write(expected));
        let saved = reopened(&document);
        let edited = saved.paragraph("0000000B").unwrap();
        assert_eq!(get_paragraph_text(edited), "Charlie edited");
        assert_eq!(run_bold(edited), Some(true));
    }

    #[test]
    fn model_tab_stop_then_logged_edit_on_one_paragraph_both_survive() {
        let mut document = open();
        change_model(&mut document, "0000000B", tab_stop);
        document
            .replace_paragraph_text("0000000B", "Charlie edited")
            .unwrap();
        assert!(document.session_save_plan().unwrap().is_none());
        let saved = reopened(&document);
        let edited = saved.paragraph("0000000B").unwrap();
        assert_eq!(get_paragraph_text(edited), "Charlie edited");
        let tabs = edited
            .formatting
            .as_ref()
            .and_then(|formatting| formatting.tabs.as_deref())
            .unwrap_or_default();
        assert!(tabs.iter().any(|tab| tab.alignment == "center"), "{tabs:?}");
    }

    #[test]
    fn known_limit_edited_empty_paragraph_with_mark_formatting_takes_the_full_write() {
        let empty = r#"<w:p w14:paraId="0000000E" w14:textId="77777777"><w:pPr><w:rPr><w:b/></w:rPr></w:pPr></w:p>"#;
        let source = document_xml(
            &[
                first_paragraph(),
                empty.to_owned(),
                paragraph("0000000C", &run("Delta")),
            ]
            .concat(),
        );
        let mut document = open_with(&source);
        document.replace_paragraph_text("0000000E", "X").unwrap();
        assert!(document.session_save_plan().unwrap().is_none());
        let saved = main_part(&document);
        assert!(!saved.contains("w:proofErr"), "{saved}");
        let mut expected = open_with(&source);
        expected.replace_paragraph_text("0000000E", "X").unwrap();
        assert_eq!(saved, full_write(expected));
        let saved = reopened(&document);
        assert_eq!(
            get_paragraph_text(saved.paragraph("0000000E").unwrap()),
            "X"
        );
    }
}

#[cfg(all(test, feature = "fonts"))]
mod font_requirement_tests {
    use super::Document;
    use docx_layout::types::LayoutBlock;
    use serde_json::json;

    fn document(language: &str) -> Document {
        let styles = format!(
            r#"<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:docDefaults><w:rPrDefault><w:rPr><w:lang w:val="en-US" w:eastAsia="{language}"/></w:rPr></w:rPrDefault></w:docDefaults></w:styles>"#
        );
        let parts = [
            ("[Content_Types].xml", r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>"#.to_owned()),
            ("_rels/.rels", r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="doc" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"#.to_owned()),
            ("word/_rels/document.xml.rels", r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>"#.to_owned()),
            ("word/styles.xml", styles),
            ("word/document.xml", r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>x</w:t></w:r></w:p></w:body></w:document>"#.to_owned()),
        ];
        let parts: Vec<_> = parts
            .into_iter()
            .map(|(name, value)| (name.to_owned(), value.into_bytes()))
            .collect();
        Document::open(&ooxml_opc::rezip_parts(&parts).unwrap()).unwrap()
    }

    #[test]
    fn layout_puts_the_documents_own_cjk_script_first() {
        let paragraph = |id: &str, language: &str| -> LayoutBlock {
            serde_json::from_value(json!({"kind": "paragraph", "id": id, "runs": [{
                "kind": "text", "text": "骨", "fontFamily": "Calibri",
                "language": {"eastAsia": language}
            }]}))
            .unwrap()
        };
        let blocks = [paragraph("zh", "zh-CN"), paragraph("ja", "ja-JP")];
        for (language, expected) in [
            ("ja-JP", ["cjk-jp", "cjk-sc"]),
            ("zh-CN", ["cjk-sc", "cjk-jp"]),
        ] {
            let requirements = document(language).font_requirements(&blocks);
            let calibri = requirements
                .iter()
                .find(|requirement| requirement.key == "calibri|0|0")
                .unwrap();
            assert_eq!(calibri.scripts, expected, "{language}");
        }
    }
}
