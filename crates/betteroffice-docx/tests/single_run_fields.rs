use betteroffice_docx::{BlockContent, Document};

fn field(code: &str, result: &str) -> String {
    format!(
        concat!(
            r#"<w:r><w:fldChar w:fldCharType="begin"/><w:instrText xml:space="preserve">{}</w:instrText>"#,
            r#"<w:fldChar w:fldCharType="separate"/><w:t>{}</w:t><w:fldChar w:fldCharType="end"/></w:r>"#
        ),
        code, result
    )
}

#[test]
fn a_field_held_in_one_run_loses_its_cached_result_and_instruction_spacing_on_a_rewrite_as_on_main()
{
    let cases = [
        ("page", field(" PAGE ", "1")),
        ("ref", field(" REF target \\h ", "Cached text")),
        (
            "date",
            field(" DATE \\@ \"d MMMM yyyy\" ", "20 November 2026"),
        ),
    ];
    let body: String = cases
        .iter()
        .enumerate()
        .map(|(index, (_, run))| {
            format!(r#"<w:p w14:paraId="7777777{index}"><w:r><w:t>A</w:t></w:r>{run}</w:p>"#)
        })
        .collect();
    let document_xml = format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:body>{body}<w:sectPr/></w:body></w:document>"#
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
        ("word/document.xml".to_owned(), document_xml.into_bytes()),
    ];
    let mut document = Document::open(&ooxml_opc::rezip_parts(&parts).unwrap()).unwrap();
    for block in &mut document.model_mut().body.content {
        if let BlockContent::Paragraph(paragraph) = block {
            std::sync::Arc::make_mut(paragraph)
                .formatting
                .get_or_insert_with(Default::default)
                .alignment = Some("center".to_owned());
        }
    }
    let saved = document.save().unwrap();
    let parts = ooxml_opc::unzip_parts(&saved).unwrap();
    let xml = parts
        .iter()
        .find(|(name, _)| name == "word/document.xml")
        .map(|(_, bytes)| String::from_utf8_lossy(bytes).into_owned())
        .unwrap();
    let paragraphs: Vec<String> = xml
        .split("<w:p ")
        .skip(1)
        .map(|rest| {
            format!(
                "<w:p {}",
                &rest[..rest.find("</w:p>").map_or(rest.len(), |at| at + 6)]
            )
        })
        .collect();
    let out: Vec<String> = cases
        .iter()
        .zip(paragraphs)
        .map(|((name, _), paragraph)| format!("{name}: {paragraph}"))
        .collect();
    let expected = [
        r#"page: <w:p w14:paraId="77777770"><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:t>A</w:t></w:r><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText>PAGE</w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>"#,
        r#"ref: <w:p w14:paraId="77777771"><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:t>A</w:t></w:r><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText>REF target \h</w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>"#,
        r#"date: <w:p w14:paraId="77777772"><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:t>A</w:t></w:r><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText>DATE \@ &quot;d MMMM yyyy&quot;</w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>"#,
    ];
    assert_eq!(out, expected);
}
