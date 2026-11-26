use std::collections::BTreeSet;

use docx_edit::projection::{ProjectionError, ProjectionFamily, ProjectionInput, ProjectionOutput};
use docx_edit::{EditCtx, EditingDoc, FormatPolicy, Position, seed_from_docx};
use docx_parse::BlockContent;
use docx_parse::inline::{InlineNode, Run, RunContent};
use docx_parse::paragraph::{Paragraph, ParagraphContent};

const WORD: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

fn seeded(body: &str) -> EditingDoc {
    let parts = [
        ("[Content_Types].xml", r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/></Types>"#.to_owned()),
        ("_rels/.rels", r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="main" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"#.to_owned()),
        ("word/document.xml", format!(
            r#"<w:document xmlns:w="{WORD}" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:body>{body}</w:body></w:document>"#
        )),
        ("word/_rels/document.xml.rels", r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="numbering" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/></Relationships>"#.to_owned()),
        ("word/numbering.xml", format!(
            r#"<w:numbering xmlns:w="{WORD}"><w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>"#
        )),
    ].map(|(path, xml)| (path.to_owned(), xml.into_bytes()));
    let bytes = ooxml_opc::rezip_parts(&parts).unwrap();
    let doc = EditingDoc::new(41);
    seed_from_docx(&doc, &bytes).unwrap();
    doc
}

fn body(projection: &ProjectionOutput) -> &[BlockContent] {
    &projection.document.package.document.content
}

fn paragraph(block: &BlockContent) -> &Paragraph {
    let BlockContent::Paragraph(paragraph) = block else { panic!("expected paragraph") };
    paragraph
}

fn run(content: &ParagraphContent) -> &Run {
    let ParagraphContent::Inline(InlineNode::Run(run)) = content else { panic!("expected run") };
    run
}

fn text(paragraph: &Paragraph) -> String {
    paragraph.content.iter().flat_map(|content| &run(content).content)
        .filter_map(|content| match content {
            RunContent::Text { text, .. } => Some(text.as_str()),
            _ => None,
        }).collect()
}

#[test]
fn projects_plain_paragraphs_in_session_order() {
    let doc = seeded(concat!(
        r#"<w:p w14:paraId="00000011" w14:textId="00000021"><w:r><w:t>A😀</w:t></w:r><w:r><w:t>B</w:t></w:r></w:p>"#,
        r#"<w:p w14:paraId="00000012"/>"#,
        r#"<w:p w14:paraId="00000013"><w:r><w:t>Last</w:t></w:r></w:p>"#,
    ));
    let output = doc.project_document(&ProjectionInput::default()).unwrap();
    let [first, empty, last] = body(&output) else { panic!("expected three paragraphs") };
    let first = paragraph(first);
    assert_eq!(first.para_id.as_deref(), Some("00000011"));
    assert_eq!(first.text_id.as_deref(), Some("00000021"));
    assert_eq!(first.source_ordinal, Some(0));
    assert_eq!(first.content.len(), 2);
    assert_eq!(text(first), "A😀B");
    assert!(paragraph(empty).content.is_empty());
    assert_eq!(text(paragraph(last)), "Last");
    assert_eq!(output.paragraph_session_keys, ["00000011", "00000012", "00000013"]);
    assert!(output.created_list_sites.is_empty());
    assert!(output.session_revision_placeholders.is_empty());
}

#[test]
fn restores_authored_run_properties_and_boundaries() {
    let doc = seeded(concat!(
        r#"<w:p><w:r><w:rPr><w:b w:val="0"/><w:i/><w:sz w:val="26"/><w:rFonts w:ascii="Arial" w:eastAsia="Yu Gothic"/><w:lang w:val="en-US" w:eastAsia="ja-JP"/></w:rPr><w:t>First</w:t></w:r>"#,
        r#"<w:r><w:rPr><w:b/></w:rPr><w:t>Second</w:t></w:r></w:p>"#,
    ));
    let output = doc.project_document(&ProjectionInput::default()).unwrap();
    let [first, second] = paragraph(&body(&output)[0]).content.as_slice() else {
        panic!("expected both runs");
    };
    let format = run(first).formatting.as_ref().unwrap();
    assert_eq!(format.bold, Some(false));
    assert_eq!(format.italic, Some(true));
    assert_eq!(format.font_size, Some(26.0));
    let fonts = format.font_family.as_ref().unwrap();
    assert_eq!(fonts.ascii.as_deref(), Some("Arial"));
    assert_eq!(fonts.east_asia.as_deref(), Some("Yu Gothic"));
    let language = format.language.as_ref().unwrap();
    assert_eq!(language.latin.as_deref(), Some("en-US"));
    assert_eq!(language.east_asia.as_deref(), Some("ja-JP"));
    assert_eq!(run(second).formatting.as_ref().unwrap().bold, Some(true));
}

#[test]
fn projects_authored_paragraph_properties_and_numbering() {
    let doc = seeded(concat!(
        r#"<w:p><w:pPr><w:jc w:val="right"/><w:spacing w:before="120" w:after="0" w:line="320" w:lineRule="exact"/>"#,
        r#"<w:ind w:left="480" w:hanging="240"/><w:keepNext/><w:widowControl w:val="0"/>"#,
        r#"<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>Numbered</w:t></w:r></w:p>"#,
    ));
    let result = doc.project_document(&ProjectionInput::default()).unwrap();
    let format = paragraph(&body(&result)[0]).formatting.as_ref().unwrap();
    assert_eq!(format.alignment.as_deref(), Some("right"));
    assert_eq!(format.space_before, Some(120.0));
    assert_eq!(format.space_after, Some(0.0));
    assert_eq!(format.line_spacing, Some(320.0));
    assert_eq!(format.line_spacing_rule.as_deref(), Some("exact"));
    assert_eq!(format.indent_left, Some(480.0));
    assert_eq!(format.hanging_indent, Some(true));
    assert_eq!(format.keep_next, Some(true));
    assert_eq!(format.widow_control, Some(false));
    let numbering = format.num_pr.as_ref().unwrap();
    assert_eq!(numbering.num_id, Some(1.0));
    assert_eq!(numbering.ilvl, Some(0.0));
}

#[test]
fn reprojects_after_an_edit_and_carries_the_last_assembled_body() {
    let doc = seeded(r#"<w:p w14:paraId="00000011"><w:r><w:t>A😀Z</w:t></w:r></w:p>"#);
    let defaults = ProjectionInput::default();
    let before = doc.project_document(&defaults).unwrap();
    doc.insert_text(
        &EditCtx::local("Ada", "2026-09-24T00:00:00Z"),
        Position::new("body", 3),
        "!",
        FormatPolicy::Inherit,
    ).unwrap();
    let story = doc.project_story("body", &defaults).unwrap();
    let unchanged = doc.project_document(&ProjectionInput {
        root_stories: Some(BTreeSet::new()),
        ..ProjectionInput::default()
    }).unwrap();
    assert_eq!(text(paragraph(&body(&unchanged)[0])), "A😀Z");
    let after = doc.project_document(&defaults).unwrap();
    assert_eq!(text(paragraph(&body(&after)[0])), "A😀!Z");
    assert_eq!(body(&story), body(&after));
    assert_eq!(after.paragraph_session_keys, before.paragraph_session_keys);
    assert_eq!(story.paragraph_session_keys, after.paragraph_session_keys);
}

#[test]
fn defers_tab_in_its_paragraph_and_carries_the_source_tab() {
    let doc = seeded(r#"<w:p><w:r><w:t>A</w:t><w:tab/><w:t>B</w:t></w:r></w:p>"#);
    let output = doc.project_document(&ProjectionInput::default()).unwrap();
    let tab = output.deferred.iter().find(|deferred| deferred.kind == "tab").unwrap();
    assert_eq!(tab.family, ProjectionFamily::Fields);
    assert_eq!(tab.story_id, "body");
    assert_eq!(tab.paragraph, Some(0));
    let content = &paragraph(&body(&output)[0]).content;
    assert!(content.iter().any(|content| run(content).content.contains(&RunContent::Tab)));
    assert_eq!(text(paragraph(&body(&output)[0])), "AB");
}

#[test]
fn carries_a_deferred_table_without_projecting_its_paragraph_keys() {
    let doc = seeded(concat!(
        r#"<w:p w14:paraId="00000011"><w:r><w:t>First</w:t></w:r></w:p>"#,
        r#"<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="1440"/></w:tblGrid><w:tr><w:tc><w:p><w:r><w:t>Cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl>"#,
        r#"<w:p w14:paraId="00000012"><w:r><w:t>Last</w:t></w:r></w:p>"#,
    ));
    let output = doc.project_document(&ProjectionInput::default()).unwrap();
    let [first, BlockContent::Table(table), last] = body(&output) else {
        panic!("expected a table between two paragraphs");
    };
    assert_eq!(text(paragraph(first)), "First");
    assert_eq!(text(paragraph(&table.rows[0].cells[0].content[0])), "Cell");
    assert_eq!(text(paragraph(last)), "Last");
    assert_eq!(output.paragraph_session_keys, ["00000011", "00000012"]);
    assert!(output.deferred.iter().any(|deferral| {
        deferral.family == ProjectionFamily::Tables && deferral.kind == "table"
            && deferral.story_id == "body" && deferral.paragraph == Some(1)
    }));
}

#[test]
fn requires_a_retained_source() {
    let doc = EditingDoc::new(41);
    assert!(matches!(
        doc.project_document(&ProjectionInput::default()),
        Err(ProjectionError::MissingSource)
    ));
}
