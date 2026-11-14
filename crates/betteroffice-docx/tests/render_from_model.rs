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

/// Documents with a body paragraph [`Document::replace_paragraph_text`] accepts.
fn editable_fixtures() -> Vec<(&'static str, Vec<u8>)> {
    let mut editable: Vec<_> = fixtures()
        .into_iter()
        .filter(|(name, _)| *name == "principal")
        .collect();
    editable.push(("three pages", one_page_per_paragraph_docx(3, "")));
    editable
}

fn edit_a_body_paragraph(document: &mut Document) {
    let candidates: Vec<String> = document
        .paragraphs()
        .into_iter()
        .filter(|paragraph| !get_paragraph_text(paragraph).is_empty())
        .filter_map(|paragraph| paragraph.para_id.clone())
        .collect();
    let mut last = None;
    for para_id in &candidates {
        match document.replace_paragraph_text(para_id, EDITED) {
            Ok(_) => return,
            Err(error) => last = Some(error),
        }
    }
    panic!(
        "none of {} paragraphs takes a logged edit, the last refusal: {last:?}",
        candidates.len()
    );
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
    docx_with_header_and_body(None, body)
}

/// [`docx_with_body`] with a default header holding `header` when one is given.
fn docx_with_header_and_body(header: Option<&str>, body: &str) -> Vec<u8> {
    let namespaces = r#"xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" mc:Ignorable="w14""#;
    let header_reference = if header.is_some() {
        r#"<w:headerReference w:type="default" r:id="rIdHeader1"/>"#
    } else {
        ""
    };
    let header_override = if header.is_some() {
        r#"<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>"#
    } else {
        ""
    };
    let document = format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {namespaces}><w:body>{body}<w:sectPr>{header_reference}<w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body></w:document>"#
    );
    let content_types = format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>{header_override}</Types>"#
    );
    let mut parts = vec![
        ("[Content_Types].xml".to_owned(), content_types.into_bytes()),
        (
            "_rels/.rels".to_owned(),
            br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"#.to_vec(),
        ),
        ("word/document.xml".to_owned(), document.into_bytes()),
    ];
    if let Some(header) = header {
        parts.push((
            "word/_rels/document.xml.rels".to_owned(),
            br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdHeader1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/></Relationships>"#.to_vec(),
        ));
        parts.push((
            "word/header1.xml".to_owned(),
            format!(
                r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:hdr {namespaces}>{header}</w:hdr>"#
            )
            .into_bytes(),
        ));
    }
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
    for (name, bytes) in editable_fixtures() {
        let mut document = open(&bytes);
        edit_a_body_paragraph(&mut document);
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
    for (name, bytes) in editable_fixtures() {
        let mut document = open(&bytes);
        center_the_first_paragraph(&mut document);
        assert_renders_as_saved(&document, &format!("{name} changed"), false);
        edit_a_body_paragraph(&mut document);
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
        panic!(
            "a {pages}-paragraph document opened under a {}-paragraph budget",
            pages - 1
        );
    };
    assert!(
        error.to_string().contains("paragraph"),
        "unexpected error: {error}"
    );
    let mut document =
        with_fonts(Document::open_with_limits(&bytes, &paragraph_limits(pages * 2)).unwrap());
    assert_eq!(
        assert_every_page_renders(&document, "raised limits", false),
        pages
    );
    edit_a_body_paragraph(&mut document);
    assert_eq!(
        assert_every_page_renders(&document, "raised limits edited", false),
        pages
    );
}

#[test]
fn a_document_one_attribute_past_the_default_limit_renders_as_the_saved_render_does() {
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
        panic!(
            "a paragraph with {} attributes opened under the default limit",
            limit + 1
        );
    };
    assert!(over_the_limit(&error), "unexpected error: {error:?}");
    let raised = ParseLimits {
        max_attributes_per_element: limit + 1,
        ..ParseLimits::default()
    };
    let document = with_fonts(Document::open_with_limits(&bytes, &raised).unwrap());
    let target = PngTarget { page: 0 };
    let options = RenderOptions::default();
    let saved = document.render_png_saved(target, &options).unwrap_err();
    assert!(over_the_limit(&saved), "unexpected error: {saved:?}");
    let before = model_renders();
    let fast = document.render_png(target, &options).unwrap_err();
    assert_render_path(before, false, "one attribute past the limit");
    assert!(over_the_limit(&fast), "unexpected error: {fast:?}");
    assert_eq!(format!("{fast:?}"), format!("{saved:?}"));
}

#[test]
fn an_unreadable_paragraph_renders_as_the_saved_render_does() {
    let document = open(&docx_with_body(concat!(
        r#"<w:p w14:paraId="00000001" w14:textId="77777777"><w:r><w:t>Before</w:t></w:r></w:p>"#,
        r#"<w:p w14:paraId="00000002" w14:textId="77777777"><w:ins w:id="100000000000000000000" w:author="Author" w:date="2024-01-01T00:00:00Z"><w:r><w:t>Inserted</w:t></w:r></w:ins></w:p>"#,
    )));
    let target = PngTarget { page: 0 };
    let options = RenderOptions::default();
    let saved = document.render_png_saved(target, &options);
    assert!(
        matches!(saved, Err(Error::Render(_))),
        "unexpected saved render: {:?}",
        saved.as_ref().err()
    );
    let before = model_renders();
    let fast = document.render_png(target, &options);
    assert_render_path(before, false, "an unreadable paragraph");
    assert_eq!(
        format!("{:?}", fast.as_ref().err()),
        format!("{:?}", saved.as_ref().err())
    );
}

const UNREADABLE_INSERTION: &str = r#"<w:ins w:id="100000000000000000000" w:author="Author" w:date="2024-01-01T00:00:00Z"><w:r><w:t>Inserted</w:t></w:r></w:ins>"#;

/// Asserts `page` renders through the save, with what the saved render returns.
fn assert_renders_through_the_save(document: &Document, page: usize, label: &str) {
    let target = PngTarget { page };
    let options = RenderOptions::default();
    let saved = document.render_png_saved(target, &options);
    let before = model_renders();
    let fast = document.render_png(target, &options);
    assert_render_path(before, false, label);
    match (&fast, &saved) {
        (Ok(fast), Ok(saved)) => {
            assert!(fast == saved, "{label}: the PNG differs from the saved render");
        }
        _ => assert_eq!(
            format!("{:?}", fast.as_ref().err()),
            format!("{:?}", saved.as_ref().err()),
            "{label}"
        ),
    }
}

/// Three pages, the last holding an unreadable paragraph when `unreadable` is set.
fn pages_before_an_unreadable_paragraph(unreadable: bool) -> Vec<u8> {
    let mut body = concat!(
        r#"<w:p w14:paraId="00000001" w14:textId="77777777"><w:r><w:t>Page 0</w:t></w:r></w:p>"#,
        r#"<w:p w14:paraId="00000002" w14:textId="77777777"><w:r><w:br w:type="page"/></w:r><w:r><w:t>Page 1</w:t></w:r></w:p>"#,
        r#"<w:p w14:paraId="00000003" w14:textId="77777777"><w:r><w:br w:type="page"/></w:r></w:p>"#,
    )
    .to_owned();
    if unreadable {
        body.push_str(&format!(
            r#"<w:p w14:paraId="00000004" w14:textId="77777777">{UNREADABLE_INSERTION}</w:p>"#
        ));
    }
    docx_with_body(&body)
}

#[test]
fn a_resident_conversion_error_renders_as_the_saved_render_does() {
    let header = format!(
        r#"<w:p w14:paraId="00000010" w14:textId="77777777">{UNREADABLE_INSERTION}</w:p>"#
    );
    let document = open(&docx_with_header_and_body(
        Some(&header),
        r#"<w:p w14:paraId="00000001" w14:textId="77777777"><w:r><w:t>Body</w:t></w:r></w:p>"#,
    ));
    assert_renders_through_the_save(&document, 0, "an unreadable header");
}

#[test]
fn a_page_without_the_unreadable_paragraph_renders_from_the_model() {
    let document = open(&pages_before_an_unreadable_paragraph(true));
    let readable = open(&pages_before_an_unreadable_paragraph(false));
    let options = RenderOptions::default();
    for page in [0, 1] {
        let target = PngTarget { page };
        let before = model_renders();
        let rendered = document
            .render_png(target, &options)
            .unwrap_or_else(|error| panic!("page {page}: {error}"));
        assert_render_path(before, true, &format!("page {page}"));
        let expected = readable.render_png(target, &options).unwrap();
        assert!(
            rendered == expected,
            "page {page}: the PNG differs from the document without the unreadable paragraph"
        );
    }
}

#[test]
fn the_page_of_an_unreadable_paragraph_fails_as_the_saved_render_does() {
    let document = open(&pages_before_an_unreadable_paragraph(true));
    let saved = document.render_png_saved(PngTarget { page: 2 }, &RenderOptions::default());
    assert!(
        matches!(saved, Err(Error::Render(_))),
        "unexpected saved render: {:?}",
        saved.as_ref().err()
    );
    assert_renders_through_the_save(&document, 2, "the unreadable paragraph's page");
}

#[test]
fn an_out_of_range_page_fails_as_the_saved_render_does() {
    let mut document = open(&one_page_per_paragraph_docx(3, ""));
    for edited in [false, true] {
        if edited {
            edit_a_body_paragraph(&mut document);
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
