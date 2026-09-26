use std::path::PathBuf;
use std::sync::Arc;

use betteroffice_docx::structured::{
    Anchor, AnchorScope, BlockKind, ExportFailureCode, InlineKind, export_docx_structured,
};
use betteroffice_docx::{
    BlockContent, Document, Error, ExportOptions, MarkdownOptions, RevisionView, StorySelection,
    render_docx_markdown,
};

fn fixture(name: &str) -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../docx-edit/tests/fixtures/structured-export")
        .join(name)
}

fn principal() -> Vec<u8> {
    std::fs::read(fixture("principal.docx")).unwrap()
}

fn all(view: RevisionView) -> ExportOptions {
    ExportOptions {
        stories: Some(vec![
            StorySelection::Body,
            StorySelection::Headers,
            StorySelection::Footers,
            StorySelection::Footnotes,
            StorySelection::Endnotes,
            StorySelection::Comments,
        ]),
        ..ExportOptions::new(view)
    }
}

fn body_text(content: &betteroffice_docx::DocxStructuredContent) -> Vec<String> {
    content.stories[0]
        .blocks
        .iter()
        .filter_map(|block| match &block.content {
            BlockKind::Paragraph { paragraph }
            | BlockKind::Heading { paragraph, .. }
            | BlockKind::ListItem { paragraph, .. } => Some(
                paragraph
                    .inlines
                    .iter()
                    .filter_map(|inline| match &inline.content {
                        InlineKind::Text { text } => Some(text.as_str()),
                        _ => None,
                    })
                    .collect(),
            ),
            _ => None,
        })
        .collect()
}

#[test]
fn native_exports_match_the_bytes_export_and_golden_files() {
    let bytes = principal();
    let document = Document::open(&bytes).unwrap();
    for (view, name) in [
        (RevisionView::Accepted, "accepted"),
        (RevisionView::Original, "original"),
        (RevisionView::Markup, "markup"),
    ] {
        let content = document.export_structured(&all(view)).unwrap();
        assert_eq!(content.anchor_scope, AnchorScope::Snapshot);
        assert_eq!(content, export_docx_structured(&bytes, &all(view)).unwrap());
        let expected = std::fs::read_to_string(fixture(&format!("principal.{name}.json"))).unwrap();
        assert_eq!(
            serde_json::to_string_pretty(&content).unwrap() + "\n",
            expected
        );
        let markdown = document.export_markdown(&all(view)).unwrap();
        assert_eq!(
            markdown,
            render_docx_markdown(&content, &MarkdownOptions::default()).unwrap()
        );
    }
}

#[test]
fn native_exports_read_the_current_model() {
    let mut document = Document::open(&principal()).unwrap();
    document
        .replace_paragraph_text("00000017", "Edited through the facade")
        .unwrap();
    let blocks = &mut document.model_mut().body.content;
    let removed = blocks
        .iter()
        .position(|block| {
            matches!(block, BlockContent::Paragraph(paragraph) if paragraph.para_id.as_deref() == Some("00000009"))
        })
        .unwrap();
    blocks.remove(removed);
    if let BlockContent::Paragraph(paragraph) = &mut blocks[0] {
        Arc::make_mut(paragraph).para_id = Some("0000ABCD".to_owned());
    }
    let content = document
        .export_structured(&ExportOptions::new(RevisionView::Accepted))
        .unwrap();
    let texts = body_text(&content);
    assert!(texts.contains(&"Edited through the facade".to_owned()));
    assert!(!texts.contains(&"Second".to_owned()));
    assert!(matches!(
        &content.stories[0].blocks[0].anchor,
        Anchor::Paragraph { para_id, .. } if para_id == "0000ABCD"
    ));
}

#[test]
fn native_refusals_are_typed_errors() {
    let document = Document::open(&principal()).unwrap();
    let error = document
        .export_structured(&ExportOptions {
            max_bytes: Some(1),
            ..ExportOptions::new(RevisionView::Markup)
        })
        .unwrap_err();
    assert!(matches!(
        error,
        Error::Export(ref failure) if failure.code == ExportFailureCode::InvalidOptions
    ));
}

#[test]
fn native_exports_leave_the_model_and_the_saved_package_alone() {
    let bytes = principal();
    let mut parts = ooxml_opc::unzip_parts(&bytes).unwrap();
    let document_xml = parts
        .iter_mut()
        .find(|(name, _)| name == "word/document.xml")
        .unwrap();
    let xml = String::from_utf8(document_xml.1.clone())
        .unwrap()
        .replace(r#"<w:p w14:paraId="00000017">"#, "<w:p>");
    document_xml.1 = xml.into_bytes();
    let bytes = ooxml_opc::rezip_parts(&parts).unwrap();
    let document = Document::open(&bytes).unwrap();
    let model = document.model().clone();
    let saved = document
        .save_with_options(betteroffice_docx::SaveOptions::default())
        .unwrap();
    for view in [
        RevisionView::Accepted,
        RevisionView::Original,
        RevisionView::Markup,
    ] {
        document.export_structured(&all(view)).unwrap();
        document.export_markdown(&all(view)).unwrap();
    }
    assert!(document.model() == &model);
    let resaved = document
        .save_with_options(betteroffice_docx::SaveOptions::default())
        .unwrap();
    assert_eq!(resaved, saved);
    let resaved_parts = ooxml_opc::unzip_parts(&resaved).unwrap();
    let resaved_document = &resaved_parts
        .iter()
        .find(|(name, _)| name == "word/document.xml")
        .unwrap()
        .1;
    let text = String::from_utf8_lossy(resaved_document);
    assert!(text.contains("Section two"));
    assert!(
        !text.contains("body:p"),
        "exporting mints no paragraph ids into the package"
    );
}

#[test]
fn native_exports_read_multi_run_fields_inside_comments() {
    let parts: Vec<(String, Vec<u8>)> = ooxml_opc::unzip_parts(&principal())
        .unwrap()
        .into_iter()
        .map(|(name, bytes)| {
            if name != "word/comments.xml" {
                return (name, bytes);
            }
            let xml = String::from_utf8(bytes).unwrap().replace(
                r#"<w:r><w:t xml:space="preserve">Please review</w:t></w:r>"#,
                r#"<w:fldSimple w:instr=" REF Summary "><w:r><w:rPr><w:b/></w:rPr><w:t>First</w:t></w:r><w:r><w:t>Second</w:t></w:r></w:fldSimple>"#,
            );
            (name, xml.into_bytes())
        })
        .collect();
    let bytes = ooxml_opc::rezip_parts(&parts).unwrap();
    let options = ExportOptions {
        stories: Some(vec![StorySelection::Comments]),
        ..ExportOptions::new(RevisionView::Accepted)
    };
    let (sender, receiver) = std::sync::mpsc::channel();
    std::thread::spawn(move || {
        let native = Document::open(&bytes)
            .unwrap()
            .export_structured(&options)
            .unwrap();
        sender
            .send((native, export_docx_structured(&bytes, &options).unwrap()))
            .unwrap();
    });
    let (native, exported) = receiver
        .recv_timeout(std::time::Duration::from_secs(60))
        .expect("both exports finish");
    assert_eq!(native, exported);
    let fields: Vec<&InlineKind> = native
        .stories
        .iter()
        .flat_map(|story| &story.blocks)
        .filter_map(|block| match &block.content {
            BlockKind::Paragraph { paragraph } => Some(&paragraph.inlines),
            _ => None,
        })
        .flatten()
        .map(|inline| &inline.content)
        .filter(|content| matches!(content, InlineKind::Field { .. }))
        .collect();
    assert_eq!(fields.len(), 1);
}
