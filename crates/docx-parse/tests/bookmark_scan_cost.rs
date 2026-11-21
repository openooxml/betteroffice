use std::time::{Duration, Instant};

use docx_parse::serializer::{S13SaveRequest, write_docx_s13};
use docx_parse::{S9ParseOptions, parse_docx_s9_wire};
use serde_json::{Value, json};

const NS: &str = r#"xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006""#;

fn package_bytes(body: &str, comments: Option<&str>) -> Vec<u8> {
    let mut parts = vec![
        ("[Content_Types].xml".to_owned(), format!(r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>{}</Types>"#,
            if comments.is_some() { r#"<Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/>"# } else { "" }).into_bytes()),
        ("word/_rels/document.xml.rels".to_owned(), format!(r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">{}</Relationships>"#,
            if comments.is_some() { r#"<Relationship Id="c1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml"/>"# } else { "" }).into_bytes()),
        ("word/document.xml".to_owned(), format!("<w:document {NS}><w:body>{body}<w:sectPr/></w:body></w:document>").into_bytes()),
    ];
    if let Some(comments) = comments {
        parts.push(("word/comments.xml".to_owned(), format!("<w:comments {NS}>{comments}</w:comments>").into_bytes()));
    }
    ooxml_opc::rezip_parts(&parts).unwrap()
}

fn open_time(source: &[u8]) -> Duration {
    let start = Instant::now();
    parse_docx_s9_wire(source, S9ParseOptions::default()).unwrap();
    start.elapsed()
}

fn best_of_two(mut run: impl FnMut() -> Duration) -> Duration {
    run().min(run())
}

fn comment_save_time(pairs: usize) -> Duration {
    let mut blocks = String::new();
    for id in 0..pairs {
        blocks.push_str(&format!(r#"<w:bookmarkStart w:id="{id}" w:name="b{id}"/>"#));
    }
    for id in 0..pairs {
        blocks.push_str(&format!(r#"<w:bookmarkEnd w:id="{id}"/>"#));
    }
    let comment = format!(
        r#"<w:comment w:id="1" w:author="R">{blocks}<w:p><w:r><w:t>c</w:t></w:r></w:p></w:comment>"#
    );
    let body = r#"<w:p><w:commentRangeStart w:id="1"/><w:r><w:t>x</w:t></w:r><w:commentRangeEnd w:id="1"/><w:r><w:commentReference w:id="1"/></w:r></w:p>"#;
    let source = package_bytes(body, Some(&comment));
    let parsed = parse_docx_s9_wire(&source, S9ParseOptions::default()).unwrap();
    let mut package = serde_json::to_value(parsed.document.package).unwrap();
    package["document"].as_object_mut().unwrap().remove("sections");
    if let Some(text) = package.pointer_mut("/document/content/0/content/1/content/0/text") {
        *text = Value::String("edited".to_owned());
    }
    let mut request = json!({
        "determinism": {"seed": "0".repeat(64), "now": "2000-01-01T00:00:00.000Z"},
        "document": package["document"],
        "options": {"updateModifiedDate": false}
    });
    if let Some(entries) = package.get("relationshipEntries") {
        request["relationshipEntries"] = entries.clone();
    }
    best_of_two(|| {
        let request: S13SaveRequest = serde_json::from_value(request.clone()).unwrap();
        let start = Instant::now();
        write_docx_s13(request, &source).unwrap();
        start.elapsed()
    })
}

#[test]
fn comment_block_bookmarks_save_in_linear_time() {
    let small = comment_save_time(5_000);
    let large = comment_save_time(40_000);
    assert!(
        large < small.max(Duration::from_millis(5)) * 24,
        "8x the bookmark pairs took {large:?} against {small:?}"
    );
}

#[test]
fn nested_alternate_content_without_bookmarks_opens_in_linear_time() {
    let leaves = 100_000;
    let flat = format!(
        r#"<w:p><w:r><w:t>x</w:t></w:r></w:p><w:p><w:r><mc:AlternateContent><w:txbxContent>{}</w:txbxContent></mc:AlternateContent></w:r></w:p>"#,
        "<a:x/>".repeat(leaves)
    );
    let nested = format!(
        r#"<w:p><w:r><w:t>x</w:t></w:r></w:p><w:p><w:r>{}<w:txbxContent>{}{}{}</w:txbxContent>{}</w:r></w:p>"#,
        "<mc:AlternateContent>".repeat(60),
        "<w:customXml>".repeat(60),
        "<a:x/>".repeat(leaves),
        "</w:customXml>".repeat(60),
        "</mc:AlternateContent>".repeat(60),
    );
    let flat = package_bytes(&flat, None);
    let nested = package_bytes(&nested, None);
    let flat_time = best_of_two(|| open_time(&flat));
    let nested_time = best_of_two(|| open_time(&nested));
    assert!(
        nested_time < flat_time.max(Duration::from_millis(20)) * 10,
        "nesting the same leaves took {nested_time:?} against {flat_time:?}"
    );
}
