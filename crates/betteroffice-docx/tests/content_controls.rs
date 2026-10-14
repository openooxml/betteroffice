use std::sync::Arc;

use betteroffice_docx::content_controls::{AnchorScope, ControlValue, list_docx_content_controls};
use betteroffice_docx::{
    BlockContent, ContentControlQuery, ContentControlsOptions, Document, Error, InlineNode,
    ParagraphContent,
};

fn template() -> Vec<u8> {
    std::fs::read(concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../../packages/docx/src/yrs/__fixtures__/content-controls/template.docx"
    ))
    .unwrap()
}

#[test]
fn lists_the_controls_the_bytes_read_lists() {
    let bytes = template();
    let document = Document::open(&bytes).unwrap();
    let saved = document.save().unwrap();
    let native = document
        .list_content_controls(&ContentControlsOptions::default())
        .unwrap();
    let from_bytes =
        list_docx_content_controls(&bytes, &ContentControlsOptions::default()).unwrap();
    assert_eq!(native, from_bytes);
    assert_eq!(native.anchor_scope, AnchorScope::Snapshot);
    assert_eq!(native.controls.len(), 7);
    let query: ContentControlQuery =
        serde_json::from_value(serde_json::json!({"kind": "tag", "tag": "account.reference"}))
            .unwrap();
    let found = document
        .find_content_controls(&query, &ContentControlsOptions::default())
        .unwrap();
    assert_eq!(found.controls.len(), 2);
    assert_eq!(document.save().unwrap(), saved);
}

#[test]
fn reads_the_current_model_including_its_edits() {
    let mut document = Document::open(&template()).unwrap();
    for block in &mut document.model_mut().body.content {
        let BlockContent::Paragraph(paragraph) = block else {
            continue;
        };
        if paragraph.para_id.as_deref() != Some("10000004") {
            continue;
        }
        for content in &mut Arc::make_mut(paragraph).content {
            if let ParagraphContent::Inline(InlineNode::InlineSdt(sdt)) = content {
                sdt.properties.tag = Some("account.previous".to_owned());
            }
        }
    }
    let snapshot = document
        .list_content_controls(&ContentControlsOptions::default())
        .unwrap();
    let tags: Vec<&str> = snapshot
        .controls
        .iter()
        .filter_map(|control| control.metadata.tag.as_deref())
        .collect();
    assert!(tags.contains(&"account.previous"));
    assert_eq!(
        tags.iter()
            .filter(|tag| **tag == "account.reference")
            .count(),
        1
    );
    let previous = snapshot
        .controls
        .iter()
        .find(|control| control.metadata.tag.as_deref() == Some("account.previous"))
        .unwrap();
    assert_eq!(
        previous.value,
        ControlValue::Text {
            text: "REF-OLD".to_owned()
        }
    );
}

#[test]
fn multi_line_follows_the_typed_property() {
    let mut document = Document::open(&template()).unwrap();
    let address = |document: &Document| {
        document
            .list_content_controls(&ContentControlsOptions::default())
            .unwrap()
            .controls
            .into_iter()
            .find(|control| control.metadata.tag.as_deref() == Some("customer.address"))
            .unwrap()
            .multi_line
    };
    assert_eq!(address(&document), Some(true));
    for block in &mut document.model_mut().body.content {
        if let BlockContent::BlockSdt(sdt) = block {
            Arc::make_mut(sdt).properties.multi_line = Some(false);
        }
    }
    assert_eq!(address(&document), Some(false));
}

#[test]
fn refusals_are_export_errors() {
    let document = Document::open(&template()).unwrap();
    let error = document
        .list_content_controls(&ContentControlsOptions {
            max_controls: Some(1),
            ..ContentControlsOptions::default()
        })
        .unwrap_err();
    assert!(matches!(error, Error::Export(_)), "{error}");
}

#[test]
fn block_controls_in_table_cells_survive_save() {
    let w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    let document = format!(
        r#"<w:document xmlns:w="{w}"><w:body><w:tbl><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid><w:tr><w:tc><w:sdt><w:sdtPr><w:alias w:val="Cell"/><w:tag w:val="cell"/><w:id w:val="101"/><w:text/></w:sdtPr><w:sdtContent><w:p><w:r><w:t>cell text</w:t></w:r></w:p></w:sdtContent></w:sdt></w:tc></w:tr></w:tbl><w:sectPr/></w:body></w:document>"#
    );
    let bytes = ooxml_opc::rezip_parts(&[
        (
            "[Content_Types].xml".to_owned(),
            br#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>"#.to_vec(),
        ),
        (
            "_rels/.rels".to_owned(),
            br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"#.to_vec(),
        ),
        ("word/document.xml".to_owned(), document.into_bytes()),
    ])
    .unwrap();
    let saved = Document::open(&bytes).unwrap().save().unwrap();
    let controls = list_docx_content_controls(&saved, &ContentControlsOptions::default())
        .unwrap()
        .controls;
    let [control] = controls.as_slice() else {
        panic!("one control: {controls:?}");
    };
    assert_eq!(control.metadata.tag.as_deref(), Some("cell"));
    assert_eq!(control.metadata.alias.as_deref(), Some("Cell"));
    assert_eq!(control.metadata.ooxml_id.as_deref(), Some("101"));
    assert_eq!(
        control.value,
        ControlValue::Text {
            text: "cell text".to_owned()
        }
    );
}
