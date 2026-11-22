use std::sync::Arc;

use betteroffice_docx::{BlockContent, Document};

const CONTENT_TYPES: &str = concat!(
    r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#,
    r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">"#,
    r#"<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>"#,
    r#"<Default Extension="xml" ContentType="application/xml"/>"#,
    r#"<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>"#,
    "</Types>"
);
const ROOT_RELS: &str = concat!(
    r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#,
    r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">"#,
    r#"<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>"#,
    "</Relationships>"
);
const DOCUMENT_START: &str = concat!(
    r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#,
    r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" "#,
    r#"xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:body>"#
);
const FIELDS: [(&str, &str); 3] = [
    ("PAGE", "1"),
    (r"REF target \h", "Cached text"),
    (r#"DATE \@ "d MMMM yyyy""#, "20 November 2026"),
];

fn instruction(code: &str) -> String {
    format!(r#"<w:instrText xml:space="preserve"> {code} </w:instrText>"#)
}

/// One run holding the whole field `code` with its cached `result`.
fn field_run(code: &str, result: &str) -> String {
    format!(
        concat!(
            r#"<w:r><w:fldChar w:fldCharType="begin"/>{}"#,
            r#"<w:fldChar w:fldCharType="separate"/><w:t>{}</w:t><w:fldChar w:fldCharType="end"/></w:r>"#
        ),
        instruction(code),
        result
    )
}

/// Three paragraphs, each "A" and then one of `FIELDS` in one run.
fn document_xml() -> String {
    let body: String = FIELDS
        .iter()
        .enumerate()
        .map(|(index, (code, result))| {
            format!(
                r#"<w:p w14:paraId="7777777{index}"><w:r><w:t>A</w:t></w:r>{}</w:p>"#,
                field_run(code, result)
            )
        })
        .collect();
    format!("{DOCUMENT_START}{body}<w:sectPr/></w:body></w:document>")
}

fn package(document_xml: &str) -> Vec<u8> {
    let parts = [
        (
            "[Content_Types].xml".to_owned(),
            CONTENT_TYPES.as_bytes().to_vec(),
        ),
        ("_rels/.rels".to_owned(), ROOT_RELS.as_bytes().to_vec()),
        (
            "word/document.xml".to_owned(),
            document_xml.as_bytes().to_vec(),
        ),
    ];
    ooxml_opc::rezip_parts(&parts).unwrap()
}

fn saved_document_xml(saved: &[u8]) -> String {
    let (_, xml) = ooxml_opc::unzip_parts(saved)
        .unwrap()
        .into_iter()
        .find(|(path, _)| path == "word/document.xml")
        .unwrap();
    String::from_utf8(xml).unwrap()
}

fn paragraphs(xml: &str) -> Vec<&str> {
    xml.split("<w:p ")
        .skip(1)
        .map(|rest| &rest[..rest.find("</w:p>").unwrap()])
        .collect()
}

#[test]
fn a_rewritten_paragraph_keeps_its_run_holding_one_field_as_read() {
    let mut document = Document::open(&package(&document_xml())).unwrap();
    for block in &mut document.model_mut().body.content {
        if let BlockContent::Paragraph(paragraph) = block {
            Arc::make_mut(paragraph)
                .formatting
                .get_or_insert_with(Default::default)
                .alignment = Some("center".to_owned());
        }
    }
    let xml = saved_document_xml(&document.save().unwrap());
    let paragraphs = paragraphs(&xml);
    assert_eq!(paragraphs.len(), FIELDS.len(), "{xml}");
    for ((code, result), paragraph) in FIELDS.iter().zip(paragraphs) {
        assert!(
            paragraph.contains(r#"<w:jc w:val="center"/>"#),
            "{paragraph}"
        );
        assert!(paragraph.contains(&field_run(code, result)), "{paragraph}");
        assert!(
            paragraph.contains(&format!("<w:t>{result}</w:t>")),
            "{paragraph}"
        );
        assert!(paragraph.contains(&instruction(code)), "{paragraph}");
    }
}

#[test]
fn a_save_without_an_edit_keeps_the_document_part_byte_for_byte() {
    let source = document_xml();
    let document = Document::open(&package(&source)).unwrap();
    assert_eq!(saved_document_xml(&document.save().unwrap()), source);
}
