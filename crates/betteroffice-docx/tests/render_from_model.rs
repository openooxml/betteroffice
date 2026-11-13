#![cfg(all(
    feature = "raster",
    feature = "fonts-latin",
    feature = "fonts-intl",
    feature = "fonts-cjk"
))]

use std::path::PathBuf;
use std::sync::Arc;

use betteroffice_docx::{
    BlockContent, Document, ParseLimits, PngTarget, RenderOptions, RenderedPage, get_paragraph_text,
};

const CARLITO: &[u8] = include_bytes!("../../docx-raster/tests/assets/Carlito-Regular.ttf");
const EDITED: &str = "Edited from the model";
const PAGE_CAP: usize = 64;

fn fixtures() -> Vec<(&'static str, Vec<u8>)> {
    let root = PathBuf::from(env!("CARGO_MANIFEST_DIR"));
    [
        (
            "comprehensive",
            "tests/corpus/fixtures/wordprocessingml-comprehensive.docx",
        ),
        ("demo", "tests/corpus/fixtures/betteroffice-demo.docx"),
        (
            "theme-accent",
            "../docx-edit/tests/fixtures/layout-request/theme-accent.docx",
        ),
        (
            "principal",
            "../docx-edit/tests/fixtures/structured-export/principal.docx",
        ),
    ]
    .into_iter()
    .map(|(name, path)| (name, std::fs::read(root.join(path)).unwrap()))
    .collect()
}

fn with_fonts(mut document: Document) -> Document {
    for bold in [false, true] {
        document
            .register_font("Calibri", bold, false, CARLITO)
            .unwrap();
    }
    document
}

fn open(bytes: &[u8]) -> Document {
    with_fonts(Document::open(bytes).unwrap())
}

fn raised_limits() -> ParseLimits {
    let limits = ParseLimits::default();
    ParseLimits {
        max_paragraphs: limits.max_paragraphs * 4,
        ..limits
    }
}

/// Every page as the saved document renders it.
fn saved_pages(document: &Document) -> Vec<RenderedPage> {
    let mut pages = Vec::new();
    while pages.len() < PAGE_CAP {
        let page = pages.len();
        match document.render_png_saved(PngTarget { page }, &RenderOptions::default()) {
            Ok(rendered) => pages.push(rendered),
            Err(error) if error.to_string() == format!("page ordinal {page} is out of range") => {
                break;
            }
            Err(error) => panic!("saved render of page {page}: {error}"),
        }
    }
    assert!(!pages.is_empty(), "the saved document renders no page");
    pages
}

fn sampled(count: usize) -> Vec<usize> {
    let mut pages = vec![0, count / 2, count - 1];
    pages.dedup();
    pages
}

/// The page count, once the first, middle and last pages match the saved render through the path `from_model` names.
fn assert_renders_as_saved(document: &Document, label: &str, from_model: bool) -> usize {
    let saved = saved_pages(document);
    for page in sampled(saved.len()) {
        let before = Document::model_renders();
        let rendered = document
            .render_png(PngTarget { page }, &RenderOptions::default())
            .unwrap_or_else(|error| panic!("{label} page {page}: {error}"));
        assert_eq!(
            Document::model_renders() - before,
            usize::from(from_model),
            "{label} page {page}: render path"
        );
        assert!(
            rendered == saved[page],
            "{label} page {page}: the PNG differs from the saved render"
        );
    }
    saved.len()
}

fn edit_a_body_paragraph(document: &mut Document) -> bool {
    let candidates: Vec<String> = document
        .paragraphs()
        .into_iter()
        .filter(|paragraph| !get_paragraph_text(paragraph).is_empty())
        .filter_map(|paragraph| paragraph.para_id.clone())
        .collect();
    candidates
        .iter()
        .any(|para_id| document.replace_paragraph_text(para_id, EDITED).is_ok())
}

fn center_the_first_paragraph(document: &mut Document) {
    let body = &mut document.model_mut().body;
    let Some(BlockContent::Paragraph(first)) = body
        .content
        .iter_mut()
        .find(|block| matches!(block, BlockContent::Paragraph(_)))
    else {
        panic!("no body paragraph");
    };
    let paragraph = Arc::make_mut(first);
    paragraph
        .formatting
        .get_or_insert_with(Default::default)
        .alignment = Some("center".to_owned());
    let centered = paragraph.clone();
    for section in body.sections.iter_mut().flatten() {
        for block in &mut section.content {
            if let BlockContent::Paragraph(paragraph) = block
                && paragraph.para_id == centered.para_id
            {
                *paragraph = Arc::new(centered.clone());
            }
        }
    }
}

fn mark_formatted_empty_paragraph_docx() -> Vec<u8> {
    let namespaces = r#"xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" mc:Ignorable="w14""#;
    let body = concat!(
        r#"<w:p w14:paraId="0000000A" w14:textId="77777777"><w:r><w:t xml:space="preserve">Alpha </w:t></w:r><w:proofErr w:type="spellStart"/><w:r><w:t>bravvo</w:t></w:r><w:proofErr w:type="spellEnd"/></w:p>"#,
        r#"<w:p w14:paraId="0000000E" w14:textId="77777777"><w:pPr><w:rPr><w:b/></w:rPr></w:pPr></w:p>"#,
        r#"<w:p w14:paraId="0000000C" w14:textId="77777777"><w:r><w:t>Delta</w:t></w:r></w:p>"#,
    );
    let document = format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {namespaces}><w:body>{body}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body></w:document>"#
    );
    let parts = vec![
        (
            "[Content_Types].xml".to_owned(),
            br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>"#.to_vec(),
        ),
        (
            "_rels/.rels".to_owned(),
            br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"#.to_vec(),
        ),
        ("word/document.xml".to_owned(), document.into_bytes()),
    ];
    ooxml_opc::rezip_parts(&parts).unwrap()
}

fn saved_main_part(document: &Document) -> String {
    let parts = ooxml_opc::unzip_parts(&document.save().unwrap()).unwrap();
    let (_, bytes) = parts
        .into_iter()
        .find(|(name, _)| name == "word/document.xml")
        .unwrap();
    String::from_utf8(bytes).unwrap()
}

#[test]
fn an_unedited_document_renders_from_the_model_as_saved() {
    let mut longest = 0;
    for (name, bytes) in fixtures() {
        longest = longest.max(assert_renders_as_saved(&open(&bytes), name, true));
    }
    assert!(longest >= 3, "no fixture has a middle page");
}

#[test]
fn a_logged_edit_renders_from_the_model_as_saved() {
    for (name, bytes) in fixtures() {
        let mut document = open(&bytes);
        assert!(
            edit_a_body_paragraph(&mut document),
            "{name}: no paragraph takes a logged edit"
        );
        assert_renders_as_saved(&document, &format!("{name} edited"), true);
    }
}

#[test]
fn a_logged_edit_the_save_writes_in_full_renders_from_the_model_as_saved() {
    let mut document = open(&mark_formatted_empty_paragraph_docx());
    document.replace_paragraph_text("0000000E", "X").unwrap();
    let saved = saved_main_part(&document);
    assert!(!saved.contains("w:proofErr"), "{saved}");
    assert_renders_as_saved(&document, "mark-formatted empty paragraph", true);
}

#[test]
fn a_model_change_renders_through_the_save() {
    for (name, bytes) in fixtures() {
        let mut document = open(&bytes);
        center_the_first_paragraph(&mut document);
        assert_renders_as_saved(&document, &format!("{name} changed"), false);
        assert!(edit_a_body_paragraph(&mut document));
        assert_renders_as_saved(&document, &format!("{name} changed then edited"), false);
    }
}

#[test]
fn raised_parse_limits_render_from_the_model_as_saved() {
    for (name, bytes) in fixtures() {
        let mut document =
            with_fonts(Document::open_with_limits(&bytes, &raised_limits()).unwrap());
        assert_renders_as_saved(&document, &format!("{name} raised limits"), true);
        assert!(edit_a_body_paragraph(&mut document));
        assert_renders_as_saved(&document, &format!("{name} raised limits edited"), true);
    }
}

#[test]
fn an_out_of_range_page_fails_as_the_saved_render_does() {
    let (_, bytes) = fixtures().remove(0);
    let mut document = open(&bytes);
    for edited in [false, true] {
        if edited {
            assert!(edit_a_body_paragraph(&mut document));
        }
        let count = saved_pages(&document).len();
        for page in [count, count + 1, usize::MAX] {
            let target = PngTarget { page };
            let options = RenderOptions::default();
            let before = Document::model_renders();
            let fast = document.render_png(target, &options).unwrap_err();
            assert_eq!(Document::model_renders() - before, 1, "page {page}");
            let saved = document.render_png_saved(target, &options).unwrap_err();
            assert_eq!(format!("{fast:?}"), format!("{saved:?}"), "page {page}");
            assert_eq!(
                fast.to_string(),
                format!("page ordinal {page} is out of range")
            );
        }
    }
}
