#![cfg(all(
    feature = "raster",
    feature = "fonts-latin",
    feature = "fonts-intl",
    feature = "fonts-cjk"
))]

use std::path::PathBuf;
use std::sync::Arc;

use betteroffice_docx::{
    BlockContent, Document, Error, ParseLimits, PngTarget, RenderOptions, RenderedPage,
    get_paragraph_text,
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

fn paragraph_limits(max_paragraphs: usize) -> ParseLimits {
    ParseLimits {
        max_paragraphs,
        ..ParseLimits::default()
    }
}

#[cfg(feature = "render-path-counter")]
fn model_renders() -> usize {
    Document::model_renders()
}

#[cfg(not(feature = "render-path-counter"))]
fn model_renders() -> usize {
    0
}

fn assert_render_path(before: usize, from_model: bool, label: &str) {
    if cfg!(feature = "render-path-counter") {
        assert_eq!(
            model_renders() - before,
            usize::from(from_model),
            "{label}: render path"
        );
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
        let before = model_renders();
        let rendered = document
            .render_png(PngTarget { page }, &RenderOptions::default())
            .unwrap_or_else(|error| panic!("{label} page {page}: {error}"));
        assert_render_path(before, from_model, &format!("{label} page {page}"));
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
    docx_with_body(concat!(
        r#"<w:p w14:paraId="0000000A" w14:textId="77777777"><w:r><w:t xml:space="preserve">Alpha </w:t></w:r><w:proofErr w:type="spellStart"/><w:r><w:t>bravvo</w:t></w:r><w:proofErr w:type="spellEnd"/></w:p>"#,
        r#"<w:p w14:paraId="0000000E" w14:textId="77777777"><w:pPr><w:rPr><w:b/></w:rPr></w:pPr></w:p>"#,
        r#"<w:p w14:paraId="0000000C" w14:textId="77777777"><w:r><w:t>Delta</w:t></w:r></w:p>"#,
    ))
}

/// One page per paragraph, each after the first behind a page break, the first with `first_attributes` too.
fn one_page_per_paragraph_docx(pages: usize, first_attributes: &str) -> Vec<u8> {
    let body: String = (0..pages)
        .map(|page| {
            let (page_break, attributes) = if page == 0 {
                ("", first_attributes)
            } else {
                (r#"<w:r><w:br w:type="page"/></w:r>"#, "")
            };
            format!(
                r#"<w:p w14:paraId="{:08X}" w14:textId="77777777"{attributes}>{page_break}<w:r><w:t>Page {page}</w:t></w:r></w:p>"#,
                page + 1
            )
        })
        .collect();
    docx_with_body(&body)
}

fn docx_with_body(body: &str) -> Vec<u8> {
    let namespaces = r#"xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" mc:Ignorable="w14""#;
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
fn a_logged_edit_renders_through_the_save() {
    for (name, bytes) in fixtures() {
        let mut document = open(&bytes);
        assert!(
            edit_a_body_paragraph(&mut document),
            "{name}: no paragraph takes a logged edit"
        );
        assert_renders_as_saved(&document, &format!("{name} edited"), false);
    }
}

#[test]
fn a_logged_edit_the_save_writes_in_full_renders_through_the_save() {
    let mut document = open(&mark_formatted_empty_paragraph_docx());
    document.replace_paragraph_text("0000000E", "X").unwrap();
    let saved = saved_main_part(&document);
    assert!(!saved.contains("w:proofErr"), "{saved}");
    assert_renders_as_saved(&document, "mark-formatted empty paragraph", false);
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

/// The page count, once every page renders through the path `from_model` names and the next is out of range.
fn assert_every_page_renders(document: &Document, label: &str, from_model: bool) -> usize {
    let mut page = 0;
    loop {
        let before = model_renders();
        match document.render_png(PngTarget { page }, &RenderOptions::default()) {
            Ok(_) => page += 1,
            Err(error) => {
                assert_eq!(
                    error.to_string(),
                    format!("page ordinal {page} is out of range"),
                    "{label} page {page}"
                );
                assert_render_path(before, from_model, &format!("{label} page {page}"));
                return page;
            }
        }
        assert_render_path(before, from_model, &format!("{label} page {}", page - 1));
        assert!(page <= PAGE_CAP, "{label}: more than {PAGE_CAP} pages");
    }
}

#[test]
fn a_document_opened_under_raised_limits_renders() {
    let pages = 3;
    let bytes = one_page_per_paragraph_docx(pages, "");
    let Err(error) = Document::open_with_limits(&bytes, &paragraph_limits(pages - 1)) else {
        panic!("a {pages}-paragraph document opened under a {}-paragraph budget", pages - 1);
    };
    assert!(
        error.to_string().contains("paragraph"),
        "unexpected error: {error}"
    );
    let mut document =
        with_fonts(Document::open_with_limits(&bytes, &paragraph_limits(pages * 2)).unwrap());
    assert_eq!(
        assert_every_page_renders(&document, "raised limits", true),
        pages
    );
    assert!(edit_a_body_paragraph(&mut document));
    assert_eq!(
        assert_every_page_renders(&document, "raised limits edited", false),
        pages
    );
}

#[test]
fn a_document_one_attribute_past_the_default_limit_renders_from_the_model() {
    let pages = 3;
    let limit = ParseLimits::default().max_attributes_per_element;
    let attributes: String = std::iter::once(r#" xmlns:x="urn:betteroffice:test""#.to_owned())
        .chain((3..=limit).map(|index| format!(r#" x:a{index}="""#)))
        .collect();
    assert_eq!(attributes.matches('=').count() + 2, limit + 1);
    let bytes = one_page_per_paragraph_docx(pages, &attributes);
    let over_the_limit = |error: &Error| {
        matches!(
            error,
            Error::Parse(docx_parse::ParseError::ResourceLimit {
                kind: "attributesPerElement",
                ..
            })
        )
    };
    let Err(error) = Document::open(&bytes) else {
        panic!("a paragraph with {} attributes opened under the default limit", limit + 1);
    };
    assert!(over_the_limit(&error), "unexpected error: {error:?}");
    let raised = ParseLimits {
        max_attributes_per_element: limit + 1,
        ..ParseLimits::default()
    };
    let document = with_fonts(Document::open_with_limits(&bytes, &raised).unwrap());
    let saved = document
        .render_png_saved(PngTarget { page: 0 }, &RenderOptions::default())
        .unwrap_err();
    assert!(over_the_limit(&saved), "unexpected error: {saved:?}");
    assert_eq!(
        assert_every_page_renders(&document, "one attribute past the limit", true),
        pages
    );
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
            let before = model_renders();
            let fast = document.render_png(target, &options).unwrap_err();
            assert_render_path(before, !edited, &format!("page {page}"));
            let saved = document.render_png_saved(target, &options).unwrap_err();
            assert_eq!(format!("{fast:?}"), format!("{saved:?}"), "page {page}");
            assert_eq!(
                fast.to_string(),
                format!("page ordinal {page} is out of range")
            );
        }
    }
}
