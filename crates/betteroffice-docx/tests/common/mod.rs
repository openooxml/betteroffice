//! Shared round-trip helpers for the fidelity, corpus, and defect gates.
//! Each gate binary uses a subset, so unused helpers are expected here.
#![allow(dead_code)]

use betteroffice_docx::Document;

pub type Parts = Vec<ooxml_fidelity::Part>;

pub fn parts_of(package: &[u8]) -> Parts {
    ooxml_opc::unzip_parts(package).unwrap()
}

pub fn save_unedited(package: &[u8]) -> Vec<u8> {
    Document::open(package).unwrap().save().unwrap()
}

pub fn roundtrip_report(before: &Parts, after: &Parts) -> Vec<String> {
    ooxml_fidelity::roundtrip_findings(before, after).unwrap()
}

/// A modelled story, a header, media, and an unmodelled XML part. Paragraph
/// `11111111` carries bookmarks; `22222222` is a plain table-cell paragraph.
pub fn sample_docx() -> Vec<u8> {
    let parts = vec![
        (
            "[Content_Types].xml".to_owned(),
            br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/></Types>"#.to_vec(),
        ),
        (
            "_rels/.rels".to_owned(),
            br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"#.to_vec(),
        ),
        (
            "word/_rels/document.xml.rels".to_owned(),
            br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdHeader" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/></Relationships>"#.to_vec(),
        ),
        (
            "word/document.xml".to_owned(),
            br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p w14:paraId="11111111"><w:pPr><w:jc w:val="center"/><w:rPr><w:b/></w:rPr></w:pPr><w:r><w:rPr><w:b/><w:sz w:val="28"/></w:rPr><w:t xml:space="preserve">Hello </w:t></w:r><w:bookmarkStart w:id="1" w:name="mark"/><w:r><w:t>DOCX</w:t></w:r><w:bookmarkEnd w:id="1"/></w:p><w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="2400"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="2400" w:type="dxa"/></w:tcPr><w:p w14:paraId="22222222"><w:r><w:t>Cell text</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:p w14:paraId="33333333"><w:r><w:tab/><w:t>Second</w:t><w:br/><w:t>section</w:t></w:r></w:p><w:sectPr><w:headerReference w:type="default" r:id="rIdHeader"/><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body></w:document>"#.to_vec(),
        ),
        (
            "word/header1.xml".to_owned(),
            br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:p w14:paraId="44444444"><w:r><w:t>Native header</w:t></w:r></w:p></w:hdr>"#.to_vec(),
        ),
        (
            "word/media/image1.png".to_owned(),
            vec![0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A, 0x00, 0x01, 0x02, 0x03],
        ),
        (
            "customXml/item1.xml".to_owned(),
            br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><data xmlns="urn:custom"><value>kept</value></data>"#.to_vec(),
        ),
    ];
    ooxml_opc::rezip_parts(&parts).unwrap()
}

pub fn with_document_xml(package: &[u8], edit: impl Fn(String) -> String) -> Vec<u8> {
    with_part(package, "word/document.xml", edit)
}

pub fn with_part(package: &[u8], part: &str, edit: impl Fn(String) -> String) -> Vec<u8> {
    let mut parts = ooxml_opc::unzip_parts(package).unwrap();
    for (name, bytes) in &mut parts {
        if name == part {
            *bytes = edit(String::from_utf8(bytes.clone()).unwrap()).into_bytes();
        }
    }
    ooxml_opc::rezip_parts(&parts).unwrap()
}
