use docx_parse::serializer::{S13SaveRequest, write_docx_s13};
use docx_parse::{S9ParseOptions, parse_docx_s9_wire};
use serde_json::json;

const TABLE_START: &str = r#"<w:bookmarkStart w:name='Table&amp;Range' w:colLast="2" w:id = '7' w:colFirst="0" w:displacedByCustomXml="next" x:flag="a&#38;b" />"#;
const TABLE_END: &str = r#"<w:bookmarkEnd w:displacedByCustomXml='prev' w:id="7"></w:bookmarkEnd>"#;
const SECTION_START: &str = r#"<w:bookmarkStart w:id="8" w:name="SectionRange"/>"#;
const SECTION_END: &str = r#"<w:bookmarkEnd w:id="8"/>"#;
const TABLE: &str = r#"<w:tbl><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid><w:tr><w:tc><w:p/></w:tc></w:tr></w:tbl>"#;
const SECTION: &str = r#"<w:p><w:pPr><w:sectPr><w:type w:val="nextPage"/></w:sectPr></w:pPr><w:r><w:t>Section</w:t></w:r></w:p>"#;
const PARAGRAPH: &str = "<w:p><w:r><w:t>Unrelated</w:t></w:r></w:p>";

fn package(body: &str) -> Vec<u8> {
    ooxml_opc::rezip_parts(&[
        ("[Content_Types].xml".to_owned(), br#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>"#.to_vec()),
        ("word/document.xml".to_owned(), format!(r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:x="urn:bookmark"><w:body>{body}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>"#).into_bytes()),
    ]).unwrap()
}

fn rewrite(body: &str, edit: bool) -> (Vec<String>, String) {
    let source = package(body);
    let parsed = parse_docx_s9_wire(&source, S9ParseOptions::default()).unwrap();
    let mut content = serde_json::to_value(&parsed.document.package.document.content).unwrap();
    let types = content
        .as_array()
        .unwrap()
        .iter()
        .map(|block| block["type"].as_str().unwrap().to_owned())
        .collect();
    if edit {
        content[0]["content"][0]["content"][0]["text"] = json!("Edited");
    }
    (
        types,
        save_content(
            &source,
            content,
            serde_json::to_value(&parsed.document.package.document.final_section_properties)
                .unwrap(),
        ),
    )
}

fn save_content(
    source: &[u8],
    content: serde_json::Value,
    final_section_properties: serde_json::Value,
) -> String {
    let request: S13SaveRequest = serde_json::from_value(json!({
        "determinism": {"seed": "0".repeat(64), "now": "2000-01-01T00:00:00.000Z"},
        "document": {"content": content, "finalSectionProperties": final_section_properties},
        "options": {"updateModifiedDate": false}
    }))
    .unwrap();
    let saved = write_docx_s13(request, source).unwrap();
    let xml = ooxml_opc::unzip_parts(&saved)
        .unwrap()
        .into_iter()
        .find(|(name, _)| name == "word/document.xml")
        .unwrap()
        .1;
    String::from_utf8(xml).unwrap()
}

fn assert_order(xml: &str, fragments: &[&str]) {
    let mut offset = 0;
    for fragment in fragments {
        assert_eq!(xml.matches(fragment).count(), 1, "{fragment}: {xml}");
        offset += xml[offset..].find(fragment).unwrap() + fragment.len();
    }
}

#[test]
fn body_bookmarks_span_table_byte_stably() {
    let (types, xml) = rewrite(
        &format!("{TABLE_START}{TABLE}{TABLE_END}{PARAGRAPH}"),
        false,
    );
    assert_eq!(types, ["rawXml", "table", "rawXml", "paragraph"]);
    assert_order(
        &xml,
        &[TABLE_START, "<w:tbl>", "</w:tbl>", TABLE_END, "Unrelated"],
    );
}

#[test]
fn body_bookmarks_span_section_break_byte_stably() {
    let (types, xml) = rewrite(
        &format!("{SECTION_START}{SECTION}{SECTION_END}{PARAGRAPH}"),
        false,
    );
    assert_eq!(types, ["rawXml", "paragraph", "rawXml", "paragraph"]);
    let range = xml
        .split_once(SECTION_START)
        .unwrap()
        .1
        .split_once(SECTION_END)
        .unwrap()
        .0;
    assert!(range.contains("<w:sectPr"));
    assert_order(
        &xml,
        &[SECTION_START, ">Section<", SECTION_END, "Unrelated"],
    );
}

#[test]
fn body_bookmarks_survive_unrelated_edit_and_whole_part_rewrite() {
    let (_, xml) = rewrite(
        &format!("{PARAGRAPH}{TABLE_START}{TABLE}{TABLE_END}{SECTION_START}{SECTION}{SECTION_END}"),
        true,
    );
    assert_order(
        &xml,
        &[
            "Edited",
            TABLE_START,
            "<w:tbl>",
            "</w:tbl>",
            TABLE_END,
            SECTION_START,
            ">Section<",
            SECTION_END,
        ],
    );
}

#[test]
fn body_bookmarks_end_before_a_table_on_whole_part_rewrite() {
    let (_, xml) = rewrite(
        &format!("{PARAGRAPH}{TABLE_START}{SECTION}{TABLE_END}{TABLE}"),
        true,
    );
    assert_order(
        &xml,
        &[
            "Edited",
            TABLE_START,
            ">Section<",
            TABLE_END,
            "<w:tbl>",
            "</w:tbl>",
        ],
    );
    assert!(xml.rfind("<w:sectPr").unwrap() > xml.find(TABLE_END).unwrap());
}

#[test]
fn body_bookmarks_moved_into_a_paragraph_are_not_written_twice() {
    let source = package(&format!("{PARAGRAPH}{TABLE_START}{TABLE}{TABLE_END}"));
    let parsed = parse_docx_s9_wire(&source, S9ParseOptions::default()).unwrap();
    let mut content = serde_json::to_value(&parsed.document.package.document.content).unwrap();
    content
        .as_array_mut()
        .unwrap()
        .retain(|block| block["type"] != "rawXml");
    content[0]["content"] = json!([
        {"type": "bookmarkStart", "id": 7, "name": "Table&Range"},
        {"type": "run", "content": [{"type": "text", "text": "Edited"}]},
        {"type": "bookmarkEnd", "id": 7}
    ]);
    let xml = save_content(
        &source,
        content,
        serde_json::to_value(&parsed.document.package.document.final_section_properties).unwrap(),
    );
    assert_eq!(xml.matches("<w:bookmarkStart").count(), 1, "{xml}");
    assert_eq!(xml.matches("<w:bookmarkEnd").count(), 1, "{xml}");
    assert_order(
        &xml,
        &[
            r#"<w:bookmarkStart w:id="7" w:name="Table&amp;Range"/>"#,
            "Edited",
            r#"<w:bookmarkEnd w:id="7"/>"#,
            "<w:tbl>",
        ],
    );
}
