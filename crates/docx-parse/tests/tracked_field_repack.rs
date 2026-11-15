use docx_parse::serializer::{S13SaveRequest, write_docx_s13};
use docx_parse::{S9ParseOptions, parse_docx_s9_wire};
use serde_json::{Value, json};

const W: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

fn field(deleted: bool) -> String {
    let (code, text) = if deleted {
        ("delInstrText", "delText")
    } else {
        ("instrText", "t")
    };
    format!(
        r#"<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:{code} xml:space="preserve"> HYPERLINK "https://example.test" </w:{code}></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:{text}>Result</w:{text}></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>"#
    )
}

fn package(content: &str) -> Vec<u8> {
    ooxml_opc::rezip_parts(&[
        ("[Content_Types].xml".to_owned(), br#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>"#.to_vec()),
        ("word/document.xml".to_owned(), format!(r#"<w:document xmlns:w="{W}"><w:body><w:p>{content}</w:p></w:body></w:document>"#).into_bytes()),
    ]).unwrap()
}

fn repacked_xml(original: &[u8]) -> String {
    let model: Value = serde_json::to_value(
        parse_docx_s9_wire(original, S9ParseOptions::default())
            .unwrap()
            .document
            .package
            .document
            .content,
    )
    .unwrap();
    let request: S13SaveRequest = serde_json::from_value(json!({
        "determinism": {"seed": "0".repeat(64), "now": "2000-01-01T00:00:00.000Z"},
        "document": {"content": model},
        "options": {"updateModifiedDate": false}
    }))
    .unwrap();
    let saved = write_docx_s13(request, original).unwrap();
    let xml = ooxml_opc::unzip_parts(&saved)
        .unwrap()
        .into_iter()
        .find(|(path, _)| path == "word/document.xml")
        .unwrap()
        .1;
    String::from_utf8(xml).unwrap()
}

/// Word refuses `delText` and `delInstrText` inside `moveFrom`; only `del` uses them.
#[test]
fn fields_inside_revisions_repack_with_word_text_elements() {
    for (tag, deleted) in [
        ("ins", false),
        ("del", true),
        ("moveFrom", false),
        ("moveTo", false),
    ] {
        let xml = repacked_xml(&package(&format!(
            r#"<w:{tag} w:id="1" w:author="A" w:date="2026-01-01T00:00:00Z">{}</w:{tag}>"#,
            field(tag == "del")
        )));
        let body = &xml[xml.find(&format!("<w:{tag} ")).expect(tag)..];
        let (code, text) = if deleted {
            ("<w:delInstrText", "<w:delText>Result</w:delText>")
        } else {
            ("<w:instrText", "<w:t>Result</w:t>")
        };
        assert!(body.contains("w:fldCharType=\"begin\""), "{tag}: {xml}");
        assert!(body.contains(code), "{tag}: {xml}");
        assert!(body.contains(text), "{tag}: {xml}");
        assert!(body.contains("w:fldCharType=\"end\""), "{tag}: {xml}");
        if !deleted {
            assert!(!body.contains("<w:del"), "{tag}: {xml}");
        }
    }
}
