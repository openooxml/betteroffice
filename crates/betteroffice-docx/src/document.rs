use docx_edit::{EditCtx, EditingDoc, Receipt, StoryRange};
use docx_layout::types::Input as LayoutInput;
use docx_parse::block::BlockContent;
use docx_parse::document::{DocumentBody, Section, get_paragraph_text};
use docx_parse::inline::{InlineNode, Run, RunContent, RunType};
use docx_parse::paragraph::{Paragraph, ParagraphContent};
use docx_parse::s9::{S9DocumentBodyWire, S9PackageWire, S9ParseOptions, S9SectionWire};
use docx_parse::serializer::{
    S13SaveOptions, S13SaveRequest, SerializerDeterminism, write_docx_s13,
};
use docx_parse::table::Table;
use sha2::{Digest, Sha256};

use crate::types::DEFAULT_SERIALIZATION_TIME;
use crate::{DocumentModel, DocumentStructure, Error, LayoutResult, Result, SaveOptions};

pub struct Document {
    original: Vec<u8>,
    seed: String,
    model: DocumentModel,
}

impl Document {
    pub fn open(bytes: &[u8]) -> Result<Self> {
        let parsed = docx_parse::parse_docx_s9_wire(bytes, S9ParseOptions::default())?;
        let document = parsed.document;
        let model = model_from_package(
            document.package,
            document.template_variables.unwrap_or_default(),
            document.warnings.unwrap_or_default(),
        );
        Ok(Self {
            original: bytes.to_vec(),
            seed: format!("{:x}", Sha256::digest(bytes)),
            model,
        })
    }

    pub fn model(&self) -> &DocumentModel {
        &self.model
    }

    pub fn model_mut(&mut self) -> &mut DocumentModel {
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

    pub fn paragraph(&self, para_id: &str) -> Option<&Paragraph> {
        self.paragraphs()
            .into_iter()
            .find(|paragraph| paragraph.para_id.as_deref() == Some(para_id))
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
        replace_plain_run(paragraph, edited, template);
        let replacement = paragraph.clone();
        if let Some(sections) = &mut self.model.body.sections {
            for section in sections {
                if let Some(section_paragraph) = find_paragraph_mut(&mut section.content, para_id) {
                    *section_paragraph = replacement.clone();
                }
            }
        }
        Ok(receipt)
    }

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
        let request = S13SaveRequest {
            determinism: SerializerDeterminism {
                seed: self.seed.clone(),
                now: options.now,
            },
            document: self.model.body.clone(),
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
        };
        write_docx_s13(request, &self.original).map_err(Error::from)
    }
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
        media: media_entries,
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
                return Some(paragraph);
            }
            BlockContent::Paragraph(_) => {}
            BlockContent::Table(table) => {
                for row in &mut table.rows {
                    for cell in &mut row.cells {
                        if let Some(paragraph) = find_paragraph_mut(&mut cell.content, para_id) {
                            return Some(paragraph);
                        }
                    }
                }
            }
            BlockContent::BlockSdt(sdt) => {
                if let Some(paragraph) = find_paragraph_mut(&mut sdt.content, para_id) {
                    return Some(paragraph);
                }
            }
        }
    }
    None
}

fn utf16_len(text: &str) -> u32 {
    text.encode_utf16().count() as u32
}
