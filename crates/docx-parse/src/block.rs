//! Shared story dispatcher for body and recursively nested block content.

use serde::{Deserialize, Serialize};

use crate::chart::ChartPartsMap;
use crate::inline::{
    ComplexField, InlineNode, MathEquation, MathType, RunContent, SdtProperties,
    StructuredFieldContent, StructuredFieldTree, parse_sdt_properties,
};
use crate::media::MediaMap;
use crate::numbering::NumberingMap;
use crate::paragraph::{
    DrawingContext, HexIdAllocator, Paragraph, ParagraphContent, parse_paragraph,
};
use crate::relationships::RelationshipMap;
use crate::shape::{RelativeRect, Shape, ShapeTextBody};
use crate::smart_art::SmartArtContext;
use crate::styles::{DocDefaults, StyleMap};
use crate::table::{
    Table, TableCell, TableRow, infer_implicit_single_cell_row_spans,
    parse_document_table_cell_properties, parse_document_table_properties,
    parse_document_table_row_properties, parse_table_cell_property_changes,
    parse_table_cell_structural_change, parse_table_grid, parse_table_property_changes,
    parse_table_row_property_changes, parse_table_row_structural_change,
};
use crate::text_box::{get_text_box_content_element, parse_text_box};
use crate::theme::Theme;
use crate::xml::{ParseBudget, ParseError, XmlElement, XmlNode};

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
pub struct BlockSdt {
    #[serde(rename = "type")]
    pub node_type: String,
    pub properties: SdtProperties,
    pub content: Vec<BlockContent>,
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(untagged)]
pub enum BlockContent {
    Paragraph(Paragraph),
    Table(Table),
    BlockSdt(BlockSdt),
}

impl BlockContent {
    pub fn node_type(&self) -> &str {
        match self {
            Self::Paragraph(paragraph) => &paragraph.node_type,
            Self::Table(table) => &table.node_type,
            Self::BlockSdt(sdt) => &sdt.node_type,
        }
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum FieldMode {
    Code,
    Result,
}

#[derive(Clone, Debug, Default)]
struct FieldEvents {
    external_separates: usize,
    external_ends: usize,
    unmatched_modes: Vec<FieldMode>,
}

#[derive(Clone, Debug)]
struct FieldRecord {
    owner_block: usize,
    owner_content: usize,
    code_blocks: Vec<usize>,
    result_blocks: Vec<usize>,
}

#[derive(Clone, Copy, Debug)]
struct OpenField {
    record: usize,
    mode: FieldMode,
}

/// One dispatcher owns every recursive story boundary. S7 plugs typed table
/// parsing into the `Table` arm without introducing a paragraph/table callback
/// cycle.
pub struct StoryParser<'a, 'limits> {
    pub relationships: Option<&'a RelationshipMap>,
    pub theme: Option<&'a Theme>,
    pub styles: Option<&'a StyleMap>,
    pub doc_defaults: Option<&'a DocDefaults>,
    pub numbering: Option<&'a NumberingMap>,
    pub media: &'a MediaMap,
    pub charts: &'a ChartPartsMap,
    pub smart_art: &'a mut SmartArtContext,
    pub budget: &'a mut ParseBudget<'limits>,
    pub ids: &'a mut HexIdAllocator,
    pub part: &'a str,
}

impl StoryParser<'_, '_> {
    pub fn parse_blocks(
        &mut self,
        parent: &XmlElement,
        depth: usize,
        in_header_footer: bool,
    ) -> Result<Vec<BlockContent>, ParseError> {
        self.budget.check_nesting_depth(depth, self.part)?;
        let mut content = Vec::new();
        let mut records: Vec<FieldRecord> = Vec::new();
        let mut open_fields: Vec<OpenField> = Vec::new();

        for child in parent.child_elements() {
            let recognized = matches!(
                child.local_name(),
                "p" | "tbl" | "sdt" | "oMath" | "oMathPara"
            );
            if !recognized {
                continue;
            }
            let events = scan_field_block_events(child);
            if events.external_separates > 0
                && let Some(open) = open_fields.last_mut()
            {
                open.mode = FieldMode::Result;
            }
            self.budget.charge_block(self.part)?;
            let mut parsed = match child.local_name() {
                "p" => {
                    let mut drawing = DrawingContext {
                        media: self.media,
                        charts: self.charts,
                        smart_art: &mut *self.smart_art,
                    };
                    let mut paragraph = parse_paragraph(
                        child,
                        self.relationships,
                        self.theme,
                        self.styles,
                        self.doc_defaults,
                        self.numbering,
                        self.part,
                        self.budget,
                        self.ids,
                        Some(&mut drawing),
                        in_header_footer,
                        depth,
                    )?;
                    self.enrich_paragraph_text_boxes(&mut paragraph, child, depth)?;
                    BlockContent::Paragraph(paragraph)
                }
                "tbl" => BlockContent::Table(self.parse_table(child, depth, in_header_footer)?),
                "sdt" => {
                    BlockContent::BlockSdt(self.parse_block_sdt(child, depth, in_header_footer)?)
                }
                "oMath" | "oMathPara" => {
                    self.budget.charge_paragraph(self.part)?;
                    BlockContent::Paragraph(math_paragraph(child))
                }
                _ => unreachable!(),
            };

            remove_external_field_end_runs(&mut parsed, events.external_ends);
            let parsed_index = content.len();
            for open in &open_fields {
                let target = &mut records[open.record];
                match open.mode {
                    FieldMode::Code => target.code_blocks.push(parsed_index),
                    FieldMode::Result => target.result_blocks.push(parsed_index),
                }
            }
            for _ in 0..events.external_ends {
                if open_fields.pop().is_none() {
                    break;
                }
            }
            content.push(parsed);

            if !events.unmatched_modes.is_empty() {
                let candidates = top_level_complex_field_indices(&content[parsed_index]);
                let start = candidates
                    .len()
                    .saturating_sub(events.unmatched_modes.len());
                for (content_index, mode) in candidates[start..]
                    .iter()
                    .copied()
                    .zip(events.unmatched_modes.iter().copied())
                {
                    let record = records.len();
                    records.push(FieldRecord {
                        owner_block: parsed_index,
                        owner_content: content_index,
                        code_blocks: Vec::new(),
                        result_blocks: Vec::new(),
                    });
                    open_fields.push(OpenField { record, mode });
                }
            }
        }

        attach_recorded_field_blocks(&mut content, records);
        Ok(content)
    }

    fn parse_block_sdt(
        &mut self,
        element: &XmlElement,
        depth: usize,
        in_header_footer: bool,
    ) -> Result<BlockSdt, ParseError> {
        let next_depth = depth.saturating_add(1);
        self.budget.check_nesting_depth(next_depth, self.part)?;
        let nested = match element.child("w", "sdtContent") {
            Some(container) => self.parse_blocks(container, next_depth, in_header_footer)?,
            None => Vec::new(),
        };
        Ok(BlockSdt {
            node_type: "blockSdt".to_owned(),
            properties: parse_sdt_properties(
                element.child("w", "sdtPr"),
                element.child("w", "sdtEndPr"),
                self.theme,
            ),
            content: nested,
        })
    }

    fn parse_table(
        &mut self,
        element: &XmlElement,
        depth: usize,
        in_header_footer: bool,
    ) -> Result<Table, ParseError> {
        self.budget.charge_table(self.part)?;
        let properties = element.child("w", "tblPr");
        let formatting = parse_document_table_properties(properties);
        let mut table = Table {
            node_type: "table".to_owned(),
            property_changes: parse_table_property_changes(properties, formatting.as_ref()),
            formatting,
            column_widths: parse_table_grid(element.child("w", "tblGrid")),
            rows: Vec::new(),
        };
        for row in element.children_named("w", "tr") {
            self.budget.charge_table_row(self.part)?;
            table
                .rows
                .push(self.parse_table_row(row, depth, in_header_footer)?);
        }
        infer_implicit_single_cell_row_spans(&mut table);
        Ok(table)
    }

    fn parse_table_row(
        &mut self,
        element: &XmlElement,
        depth: usize,
        in_header_footer: bool,
    ) -> Result<TableRow, ParseError> {
        let properties = element.child("w", "trPr");
        let formatting = parse_document_table_row_properties(properties);
        let mut row = TableRow {
            node_type: "tableRow".to_owned(),
            property_changes: parse_table_row_property_changes(properties, formatting.as_ref()),
            structural_change: parse_table_row_structural_change(properties),
            formatting,
            cells: Vec::new(),
        };
        for cell in element.children_named("w", "tc") {
            self.budget.charge_table_cell(self.part)?;
            row.cells
                .push(self.parse_table_cell(cell, depth, in_header_footer)?);
        }
        Ok(row)
    }

    fn parse_table_cell(
        &mut self,
        element: &XmlElement,
        depth: usize,
        in_header_footer: bool,
    ) -> Result<TableCell, ParseError> {
        let properties = element.child("w", "tcPr");
        let formatting = parse_document_table_cell_properties(properties);
        let cell_depth = depth.saturating_add(1);
        self.budget.check_nesting_depth(cell_depth, self.part)?;
        let mut content = self.parse_blocks(element, cell_depth, in_header_footer)?;
        if content.is_empty() {
            self.budget.charge_block(self.part)?;
            self.budget.charge_paragraph(self.part)?;
            content.push(BlockContent::Paragraph(empty_paragraph()));
        }
        Ok(TableCell {
            node_type: "tableCell".to_owned(),
            property_changes: parse_table_cell_property_changes(properties, formatting.as_ref()),
            structural_change: parse_table_cell_structural_change(properties),
            formatting,
            content,
        })
    }

    fn enrich_paragraph_text_boxes(
        &mut self,
        paragraph: &mut Paragraph,
        source: &XmlElement,
        depth: usize,
    ) -> Result<(), ParseError> {
        if paragraph.content.is_empty() {
            return Ok(());
        }
        let mut run_index = 0usize;
        for run in source
            .child_elements()
            .filter(|child| child.local_name() == "r")
        {
            for child in run.child_elements() {
                match child.local_name() {
                    "drawing" => {
                        self.enrich_text_box_drawing(paragraph, child, run_index, depth)?
                    }
                    "AlternateContent" => {
                        let branch = child
                            .child_elements()
                            .find(|branch| branch.local_name() == "Choice")
                            .or_else(|| {
                                child
                                    .child_elements()
                                    .find(|branch| branch.local_name() == "Fallback")
                            });
                        if let Some(branch) = branch {
                            for drawing in branch
                                .child_elements()
                                .filter(|child| child.local_name() == "drawing")
                            {
                                self.enrich_text_box_drawing(paragraph, drawing, run_index, depth)?;
                            }
                        }
                    }
                    _ => {}
                }
            }
            run_index = run_index.saturating_add(1);
        }
        Ok(())
    }

    fn enrich_text_box_drawing(
        &mut self,
        paragraph: &mut Paragraph,
        drawing: &XmlElement,
        run_index: usize,
        depth: usize,
    ) -> Result<(), ParseError> {
        let Some(text_box) = parse_text_box(drawing) else {
            return Ok(());
        };
        let mut blocks = Vec::new();
        if let Some(shape) = find_descendant_by_name(drawing, "wps:wsp", 0)
            && let Some(container) = get_text_box_content_element(shape)
        {
            blocks = self.parse_blocks(container, depth.saturating_add(1), false)?;
        }
        let mut shape = Shape::empty("rect".to_owned(), text_box.size);
        shape.id = text_box.id;
        shape.position = text_box.position;
        shape.wrap = text_box.wrap;
        shape.fill = text_box.fill;
        shape.outline = text_box.outline;
        shape.text_body = Some(ShapeTextBody {
            vertical: None,
            rotation: None,
            anchor: None,
            anchor_center: None,
            auto_fit: None,
            margins: text_box.margins.map(|margins| RelativeRect {
                left: margins.left,
                top: margins.top,
                right: margins.right,
                bottom: margins.bottom,
            }),
            content: blocks
                .into_iter()
                .map(serde_json::to_value)
                .collect::<Result<_, _>>()
                .map_err(|error| ParseError::Canonical(error.to_string()))?,
        });
        let mut target = run_index;
        if target >= paragraph.content.len() {
            let Some(last_run) = paragraph.content.iter().rposition(|content| {
                matches!(content, ParagraphContent::Inline(InlineNode::Run(_)))
            }) else {
                return Ok(());
            };
            target = last_run;
        }
        if let Some(ParagraphContent::Inline(InlineNode::Run(run))) =
            paragraph.content.get_mut(target)
        {
            run.content.push(RunContent::Shape {
                shape: Box::new(shape),
            });
        }
        Ok(())
    }
}

fn scan_field_block_events(root: &XmlElement) -> FieldEvents {
    let mut events = FieldEvents::default();
    let mut modes = Vec::new();
    let mut stack = vec![root];
    while let Some(element) = stack.pop() {
        if element.local_name() == "fldChar" {
            match element.attribute(Some("w"), "fldCharType") {
                Some("begin") => modes.push(FieldMode::Code),
                Some("separate") => match modes.last_mut() {
                    Some(mode) => *mode = FieldMode::Result,
                    None => events.external_separates += 1,
                },
                Some("end") => {
                    if modes.pop().is_none() {
                        events.external_ends += 1;
                    }
                }
                _ => {}
            }
        }
        let children: Vec<_> = element.child_elements().collect();
        stack.extend(children.into_iter().rev());
    }
    events.unmatched_modes = modes;
    events
}

fn top_level_complex_field_indices(block: &BlockContent) -> Vec<usize> {
    let BlockContent::Paragraph(paragraph) = block else {
        return Vec::new();
    };
    paragraph
        .content
        .iter()
        .enumerate()
        .filter_map(|(index, content)| {
            matches!(
                content,
                ParagraphContent::Inline(InlineNode::ComplexField(_))
            )
            .then_some(index)
        })
        .collect()
}

fn attach_recorded_field_blocks(content: &mut [BlockContent], records: Vec<FieldRecord>) {
    for record in records.into_iter().rev() {
        let code: Vec<_> = record
            .code_blocks
            .iter()
            .filter_map(|index| content.get(*index).cloned())
            .collect();
        let result: Vec<_> = record
            .result_blocks
            .iter()
            .filter_map(|index| content.get(*index).cloned())
            .collect();
        let Some(field) = complex_field_mut(content, record.owner_block, record.owner_content)
        else {
            continue;
        };
        if !code.is_empty() {
            let structured = field
                .structured_code
                .get_or_insert_with(|| StructuredFieldContent {
                    inline: Some(
                        field
                            .field_code
                            .iter()
                            .cloned()
                            .map(InlineNode::Run)
                            .collect(),
                    ),
                    blocks: None,
                });
            structured.blocks.get_or_insert_with(Vec::new).extend(code);
            field.field_tree.get_or_insert_with(default_field_tree).code = Some(structured.clone());
        }
        if !result.is_empty() {
            let structured =
                field
                    .structured_result
                    .get_or_insert_with(|| StructuredFieldContent {
                        inline: Some(
                            field
                                .field_result
                                .iter()
                                .cloned()
                                .map(InlineNode::Run)
                                .collect(),
                        ),
                        blocks: None,
                    });
            structured
                .blocks
                .get_or_insert_with(Vec::new)
                .extend(result);
            field
                .field_tree
                .get_or_insert_with(default_field_tree)
                .result = Some(structured.clone());
        }
    }
}

fn complex_field_mut(
    content: &mut [BlockContent],
    block_index: usize,
    content_index: usize,
) -> Option<&mut ComplexField> {
    let BlockContent::Paragraph(paragraph) = content.get_mut(block_index)? else {
        return None;
    };
    let ParagraphContent::Inline(InlineNode::ComplexField(field)) =
        paragraph.content.get_mut(content_index)?
    else {
        return None;
    };
    Some(field)
}

fn default_field_tree() -> StructuredFieldTree {
    StructuredFieldTree {
        version: Some(1.0),
        code: None,
        result: None,
        children: None,
        display_mode: Some("result".to_owned()),
    }
}

fn remove_external_field_end_runs(block: &mut BlockContent, count: usize) {
    let BlockContent::Paragraph(paragraph) = block else {
        return;
    };
    let mut remaining = count;
    paragraph.content.retain(|content| {
        if remaining == 0 {
            return true;
        }
        let ParagraphContent::Inline(InlineNode::Run(run)) = content else {
            return true;
        };
        let non_instruction: Vec<_> = run
            .content
            .iter()
            .filter(|content| !matches!(content, RunContent::InstrText { .. }))
            .collect();
        let end_only = !non_instruction.is_empty()
            && non_instruction.iter().all(|content| {
                matches!(
                    content,
                    RunContent::FieldChar { char_type, .. } if char_type == "end"
                )
            });
        if end_only {
            remaining -= 1;
            false
        } else {
            true
        }
    });
}

fn math_paragraph(element: &XmlElement) -> Paragraph {
    Paragraph {
        node_type: "paragraph".to_owned(),
        para_id: None,
        text_id: None,
        formatting: None,
        property_changes: None,
        p_pr_ins: None,
        p_pr_del: None,
        content: vec![ParagraphContent::Inline(InlineNode::Math(MathEquation {
            node_type: MathType::MathEquation,
            display: if element.local_name() == "oMathPara" {
                "block"
            } else {
                "inline"
            }
            .to_owned(),
            omml_xml: element.to_incumbent_xml(),
            plain_text: {
                let mut text = String::new();
                append_text(element, &mut text);
                (!text.is_empty()).then_some(text)
            },
        }))],
        list_rendering: None,
        rendered_page_break_before: None,
        section_properties: None,
    }
}

fn empty_paragraph() -> Paragraph {
    Paragraph {
        node_type: "paragraph".to_owned(),
        para_id: None,
        text_id: None,
        formatting: None,
        property_changes: None,
        p_pr_ins: None,
        p_pr_del: None,
        content: Vec::new(),
        list_rendering: None,
        rendered_page_break_before: None,
        section_properties: None,
    }
}

fn append_text(element: &XmlElement, output: &mut String) {
    for child in &element.children {
        match child {
            XmlNode::Text(text) => output.push_str(text),
            XmlNode::Element(child) => append_text(child, output),
            XmlNode::CData(_) => {}
        }
    }
}

fn find_descendant_by_name<'a>(
    element: &'a XmlElement,
    full_name: &str,
    depth: usize,
) -> Option<&'a XmlElement> {
    if depth > 64 {
        return None;
    }
    if element.name == full_name {
        return Some(element);
    }
    element
        .child_elements()
        .find_map(|child| find_descendant_by_name(child, full_name, depth + 1))
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::xml::{ParseLimits, parse_xml};
    use indexmap::IndexMap;

    fn parse(xml: &str) -> Vec<BlockContent> {
        let limits = ParseLimits::default();
        let mut budget = ParseBudget::new(&limits);
        let document = parse_xml(xml.as_bytes(), "word/document.xml", &mut budget).unwrap();
        let media = IndexMap::new();
        let charts = IndexMap::new();
        let mut smart_art = SmartArtContext::default();
        let mut ids = HexIdAllocator::from_sha256(&"0".repeat(64)).unwrap();
        StoryParser {
            relationships: None,
            theme: None,
            styles: None,
            doc_defaults: None,
            numbering: None,
            media: &media,
            charts: &charts,
            smart_art: &mut smart_art,
            budget: &mut budget,
            ids: &mut ids,
            part: "word/document.xml",
        }
        .parse_blocks(document.root().unwrap(), 0, false)
        .unwrap()
    }

    #[test]
    fn preserves_block_sdts_empty_paragraphs_math_and_the_s7_table_boundary() {
        let blocks = parse(
            r#"<w:body xmlns:w="w" xmlns:m="m">
              <w:p/>
              <w:sdt><w:sdtPr><w:alias w:val="box"/></w:sdtPr><w:sdtContent><w:p><w:r><w:t>x</w:t></w:r></w:p><w:tbl/></w:sdtContent></w:sdt>
              <m:oMathPara><m:oMath><m:r><m:t>y</m:t></m:r></m:oMath></m:oMathPara>
            </w:body>"#,
        );
        let BlockContent::Paragraph(empty) = &blocks[0] else {
            panic!("empty paragraph")
        };
        assert!(empty.content.is_empty());
        let BlockContent::BlockSdt(sdt) = &blocks[1] else {
            panic!("block sdt")
        };
        assert_eq!(sdt.properties.alias.as_deref(), Some("box"));
        assert!(matches!(sdt.content[1], BlockContent::Table(_)));
        let BlockContent::Paragraph(math) = &blocks[2] else {
            panic!("math block")
        };
        assert!(matches!(
            math.content[0],
            ParagraphContent::Inline(InlineNode::Math(_))
        ));
    }

    #[test]
    fn pins_fields_and_bookmarks_spanning_blocks() {
        let blocks = parse(
            r#"<w:body xmlns:w="w">
              <w:p><w:bookmarkStart w:id="1" w:name="span"/><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText>TOC</w:instrText></w:r></w:p>
              <w:p><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>entry</w:t></w:r></w:p>
              <w:tbl/>
              <w:p><w:r><w:fldChar w:fldCharType="end"/></w:r><w:bookmarkEnd w:id="1"/></w:p>
            </w:body>"#,
        );
        let BlockContent::Paragraph(opening) = &blocks[0] else {
            panic!("opening")
        };
        assert!(matches!(
            opening.content[0],
            ParagraphContent::Inline(InlineNode::BookmarkStart(_))
        ));
        let field = opening
            .content
            .iter()
            .find_map(|content| match content {
                ParagraphContent::Inline(InlineNode::ComplexField(field)) => Some(field),
                _ => None,
            })
            .unwrap();
        let result = field.structured_result.as_ref().unwrap();
        assert_eq!(result.blocks.as_ref().unwrap().len(), 3);
        assert!(matches!(
            result.blocks.as_ref().unwrap()[1],
            BlockContent::Table(_)
        ));
        let BlockContent::Paragraph(closing) = &blocks[3] else {
            panic!("closing")
        };
        assert_eq!(closing.content.len(), 1);
        assert!(matches!(
            closing.content[0],
            ParagraphContent::Inline(InlineNode::BookmarkEnd(_))
        ));
    }

    #[test]
    fn parses_typed_rows_cells_and_nested_tables_through_the_same_dispatcher() {
        let blocks = parse(
            r#"<w:body xmlns:w="w"><w:tbl>
              <w:tblPr><w:tblW w:w="5000" w:type="dxa"/><w:tblpPr w:horzAnchor="page"/></w:tblPr>
              <w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="3000"/></w:tblGrid>
              <w:tr><w:trPr><w:tblHeader/></w:trPr>
                <w:tc><w:tcPr><w:gridSpan w:val="2"/></w:tcPr>
                  <w:p><w:r><w:t>outer</w:t></w:r></w:p>
                  <w:tbl><w:tr><w:tc/></w:tr></w:tbl>
                </w:tc>
              </w:tr>
            </w:tbl></w:body>"#,
        );
        let BlockContent::Table(table) = &blocks[0] else {
            panic!("table")
        };
        assert_eq!(table.column_widths.as_deref(), Some(&[2000.0, 3000.0][..]));
        assert!(crate::table::is_floating_table(table));
        assert!(crate::table::has_header_row(table));
        assert_eq!(crate::table::get_table_column_count(table), 2);
        assert_eq!(crate::table::get_table_text(table), "outer");
        let cell = &table.rows[0].cells[0];
        assert!(crate::table::is_cell_horizontally_merged(cell));
        assert!(matches!(cell.content[1], BlockContent::Table(_)));
        let BlockContent::Table(nested) = &cell.content[1] else {
            unreachable!()
        };
        assert!(matches!(
            nested.rows[0].cells[0].content[0],
            BlockContent::Paragraph(_)
        ));
    }

    #[test]
    fn text_box_story_uses_the_shared_dispatcher_for_ordered_blocks() {
        let blocks = parse(
            r#"<w:body xmlns:w="w" xmlns:wp="wp" xmlns:a="a" xmlns:wps="wps"><w:p><w:r><w:drawing>
              <wp:inline><wp:extent cx="914400" cy="457200"/><a:graphic><a:graphicData><wps:wsp>
                <wps:txbx><w:txbxContent>
                  <w:p><w:r><w:t>first</w:t></w:r></w:p>
                  <w:tbl><w:tr><w:tc><w:p><w:r><w:t>cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl>
                  <w:sdt><w:sdtPr><w:alias w:val="nested"/></w:sdtPr><w:sdtContent><w:p/></w:sdtContent></w:sdt>
                </w:txbxContent></wps:txbx>
              </wps:wsp></a:graphicData></a:graphic></wp:inline>
            </w:drawing></w:r></w:p></w:body>"#,
        );
        let BlockContent::Paragraph(paragraph) = &blocks[0] else {
            panic!("paragraph")
        };
        let shape = paragraph
            .content
            .iter()
            .find_map(|content| match content {
                ParagraphContent::Inline(InlineNode::Run(run)) => {
                    run.content.iter().find_map(|content| match content {
                        RunContent::Shape { shape } => Some(shape.as_ref()),
                        _ => None,
                    })
                }
                _ => None,
            })
            .expect("text box shape");
        let content = &shape.text_body.as_ref().expect("text body").content;
        assert_eq!(
            content
                .iter()
                .map(|block| block["type"].as_str().unwrap())
                .collect::<Vec<_>>(),
            ["paragraph", "table", "blockSdt"]
        );
        assert_eq!(
            content[1]["rows"][0]["cells"][0]["content"][0]["type"],
            "paragraph"
        );
    }

    #[test]
    fn cell_story_keeps_block_field_ownership_and_order() {
        let blocks = parse(
            r#"<w:body xmlns:w="w"><w:tbl><w:tr><w:tc>
              <w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText>TOC</w:instrText></w:r></w:p>
              <w:p><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>entry</w:t></w:r></w:p>
              <w:tbl/>
              <w:p><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>
            </w:tc></w:tr></w:tbl></w:body>"#,
        );
        let BlockContent::Table(table) = &blocks[0] else {
            panic!("table")
        };
        let content = &table.rows[0].cells[0].content;
        assert!(matches!(content[0], BlockContent::Paragraph(_)));
        assert!(matches!(content[1], BlockContent::Paragraph(_)));
        assert!(matches!(content[2], BlockContent::Table(_)));
        assert!(matches!(content[3], BlockContent::Paragraph(_)));
        let BlockContent::Paragraph(opening) = &content[0] else {
            unreachable!()
        };
        let field = opening
            .content
            .iter()
            .find_map(|content| match content {
                ParagraphContent::Inline(InlineNode::ComplexField(field)) => Some(field),
                _ => None,
            })
            .unwrap();
        assert_eq!(
            field
                .structured_result
                .as_ref()
                .and_then(|result| result.blocks.as_ref())
                .unwrap()
                .len(),
            3
        );
    }

    #[test]
    fn table_row_cell_and_nesting_budgets_fail_stably_before_unbounded_growth() {
        fn parse_with_limits(
            xml: &str,
            limits: &ParseLimits,
        ) -> Result<Vec<BlockContent>, ParseError> {
            let mut budget = ParseBudget::new(limits);
            let document = parse_xml(xml.as_bytes(), "word/document.xml", &mut budget).unwrap();
            let media = IndexMap::new();
            let charts = IndexMap::new();
            let mut smart_art = SmartArtContext::default();
            let mut ids = HexIdAllocator::from_sha256(&"0".repeat(64)).unwrap();
            StoryParser {
                relationships: None,
                theme: None,
                styles: None,
                doc_defaults: None,
                numbering: None,
                media: &media,
                charts: &charts,
                smart_art: &mut smart_art,
                budget: &mut budget,
                ids: &mut ids,
                part: "word/document.xml",
            }
            .parse_blocks(document.root().unwrap(), 0, false)
        }

        let rows = parse_with_limits(
            r#"<w:body xmlns:w="w"><w:tbl><w:tr/><w:tr/></w:tbl></w:body>"#,
            &ParseLimits {
                max_table_rows: 1,
                ..ParseLimits::default()
            },
        )
        .unwrap_err();
        assert!(matches!(
            rows,
            ParseError::ResourceLimit {
                kind: "tableRows",
                ..
            }
        ));

        let cells = parse_with_limits(
            r#"<w:body xmlns:w="w"><w:tbl><w:tr><w:tc/><w:tc/></w:tr></w:tbl></w:body>"#,
            &ParseLimits {
                max_table_cells: 1,
                ..ParseLimits::default()
            },
        )
        .unwrap_err();
        assert!(matches!(
            cells,
            ParseError::ResourceLimit {
                kind: "tableCells",
                ..
            }
        ));

        let depth = parse_with_limits(
            r#"<w:body xmlns:w="w"><w:tbl><w:tr><w:tc><w:tbl><w:tr><w:tc/></w:tr></w:tbl></w:tc></w:tr></w:tbl></w:body>"#,
            &ParseLimits {
                max_nesting_depth: 1,
                ..ParseLimits::default()
            },
        )
        .unwrap_err();
        assert!(matches!(
            depth,
            ParseError::ResourceLimit {
                kind: "nestingDepth",
                ..
            }
        ));
    }
}
