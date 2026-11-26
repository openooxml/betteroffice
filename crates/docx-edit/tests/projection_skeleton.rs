use std::collections::BTreeSet;

use docx_edit::projection::{ProjectionError, ProjectionFamily, ProjectionInput, ProjectionOutput};
use docx_edit::{EditCtx, EditingDoc, FormatPolicy, Position, seed_from_docx};
use docx_parse::BlockContent;
use docx_parse::inline::{InlineNode, Run, RunContent};
use docx_parse::paragraph::{Paragraph, ParagraphContent};

const WORD: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

fn seeded(body: &str) -> EditingDoc {
    seeded_with_parts(body, "", "", &[])
}

fn seeded_with_parts(
    body: &str,
    overrides: &str,
    relationships: &str,
    extra: &[(&str, String)],
) -> EditingDoc {
    let mut parts = [
        ("[Content_Types].xml", format!(r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>{overrides}</Types>"#)),
        ("_rels/.rels", r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="main" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"#.to_owned()),
        ("word/document.xml", format!(
            r#"<w:document xmlns:w="{WORD}" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body>{body}</w:body></w:document>"#
        )),
        ("word/_rels/document.xml.rels", format!(r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="numbering" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>{relationships}</Relationships>"#)),
        ("word/numbering.xml", format!(
            r#"<w:numbering xmlns:w="{WORD}"><w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>"#
        )),
    ].into_iter().map(|(path, xml)| (path.to_owned(), xml.into_bytes())).collect::<Vec<_>>();
    parts.extend(extra.iter().map(|(path, xml)| ((*path).to_owned(), xml.as_bytes().to_vec())));
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
    assert_eq!(story.blocks.as_slice(), body(&after));
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

fn seeded_with_auxiliary_stories() -> EditingDoc {
    let parts = [
        ("header", "header1", "header1", "hdr", "", "Header"),
        ("footer", "footer1", "footer1", "ftr", "", "Footer"),
        ("footnotes", "footnotes", "footnotes", "footnotes", "footnote", "Footnote"),
        ("endnotes", "endnotes", "endnotes", "endnotes", "endnote", "Endnote"),
    ];
    let mut overrides = String::new();
    let mut relationships = String::new();
    let mut extra = Vec::new();
    let paths: Vec<_> = parts.iter().map(|(_, file, _, _, _, _)| format!("word/{file}.xml")).collect();
    for ((kind, file, rel, tag, child, text), path) in parts.into_iter().zip(&paths) {
        use std::fmt::Write;
        write!(overrides, r#"<Override PartName="/{path}" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.{kind}+xml"/>"#).unwrap();
        write!(relationships, r#"<Relationship Id="{rel}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/{kind}" Target="{file}.xml"/>"#).unwrap();
        let paragraph = format!("<w:p><w:r><w:t>{text}</w:t></w:r></w:p>");
        let content = if child.is_empty() { paragraph } else {
            format!(r#"<w:{child} w:id="1">{paragraph}</w:{child}>"#)
        };
        extra.push((path.as_str(), format!(r#"<w:{tag} xmlns:w="{WORD}">{content}</w:{tag}>"#)));
    }
    seeded_with_parts(
        concat!(
            r#"<w:p w14:paraId="00000011"><w:r><w:t>Body</w:t></w:r></w:p>"#,
            r#"<w:sectPr><w:headerReference w:type="default" r:id="header1"/><w:footerReference w:type="default" r:id="footer1"/></w:sectPr>"#,
        ),
        &overrides,
        &relationships,
        &extra,
    )
}

#[test]
fn projects_each_root_as_the_matching_part_of_the_document() {
    let doc = seeded_with_auxiliary_stories();
    let input = ProjectionInput::default();
    let document = doc.project_document(&input).unwrap();
    let package = &document.document.package;
    let body = &package.document.content;
    let header = &package.header_entries.as_ref().unwrap()[0].1.content;
    let footer = &package.footer_entries.as_ref().unwrap()[0].1.content;
    let footnote = &package.footnotes.as_ref().unwrap()[0].content;
    let endnote = &package.endnotes.as_ref().unwrap()[0].content;
    let mut keys = Vec::new();
    let mut placeholders = BTreeSet::new();
    let mut sites = Vec::new();
    let mut deferred = Vec::new();
    for (id, blocks) in [
        ("body", body),
        ("en:1", endnote),
        ("fn:1", footnote),
        ("hf:footer1", footer),
        ("hf:header1", header),
    ] {
        let story = doc.project_story(id, &input).unwrap();
        assert_eq!(story.story_id, id);
        assert_eq!(&story.blocks, blocks);
        let offset = keys.len();
        keys.extend(story.paragraph_session_keys);
        placeholders.extend(story.session_revision_placeholders);
        sites.extend(story.created_list_sites.into_iter().map(|mut site| {
            site.paragraph += offset;
            site
        }));
        deferred.extend(story.deferred);
    }
    assert_eq!(document.paragraph_session_keys, keys);
    assert_eq!(document.session_revision_placeholders, placeholders);
    assert_eq!(document.created_list_sites, sites);
    assert_eq!(document.deferred, deferred);
}

#[cfg(debug_assertions)]
#[test]
fn projecting_a_header_does_not_visit_the_body() {
    let doc = seeded_with_auxiliary_stories();
    let header = doc.project_story("hf:header1", &ProjectionInput::default()).unwrap();
    assert_eq!(text(paragraph(&header.blocks[0])), "Header");
    assert!(doc.projection_has_visited_source_story("hf:header1"));
    assert!(!doc.projection_has_visited_source_story("body"));
}

#[test]
fn requires_a_retained_source() {
    let doc = EditingDoc::new(41);
    assert!(matches!(
        doc.project_document(&ProjectionInput::default()),
        Err(ProjectionError::MissingSource)
    ));
}
