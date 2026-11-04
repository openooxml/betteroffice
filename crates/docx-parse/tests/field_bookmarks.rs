use docx_parse::serializer::{S13SaveRequest, write_docx_s13};
use docx_parse::{BlockContent, S9ParseOptions, get_paragraph_text, parse_docx_s9_wire};
use serde_json::{Value, json};

const NS: &str = r#"xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main""#;

fn fixture(body: &str) -> Vec<u8> {
    let parts = vec![
        ("[Content_Types].xml".to_owned(), r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>"#.to_owned().into_bytes()),
        ("_rels/.rels".to_owned(), r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="doc" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"#.to_owned().into_bytes()),
        ("word/document.xml".to_owned(), format!("<w:document {NS}><w:body>{body}<w:sectPr/></w:body></w:document>").into_bytes()),
    ];
    ooxml_opc::rezip_parts(&parts).unwrap()
}

fn comment_fixture(body: &str) -> Vec<u8> {
    let mut parts = ooxml_opc::unzip_parts(&fixture(body)).unwrap();
    let (_, content_types) = parts
        .iter_mut()
        .find(|(name, _)| name == "[Content_Types].xml")
        .unwrap();
    *content_types = std::str::from_utf8(content_types)
        .unwrap()
        .replace(
            "</Types>",
            r#"<Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/></Types>"#,
        )
        .into_bytes();
    parts.extend([
        ("word/_rels/document.xml.rels".to_owned(), br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="comments" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml"/></Relationships>"#.to_vec()),
        ("word/comments.xml".to_owned(), format!(r#"<w:comments {NS} xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:comment w:id="99" w:author="Reviewer"><w:p w14:paraId="00000099"><w:r><w:t>Comment</w:t></w:r></w:p></w:comment></w:comments>"#).into_bytes()),
    ]);
    ooxml_opc::rezip_parts(&parts).unwrap()
}

fn saved_docx(source: &[u8]) -> Vec<u8> {
    let parsed = parse_docx_s9_wire(source, S9ParseOptions::default()).unwrap();
    let mut package: Value = serde_json::to_value(parsed.document.package).unwrap();
    package["document"]
        .as_object_mut()
        .unwrap()
        .remove("sections");
    save_package(source, &package)
}

fn save_package(source: &[u8], package: &Value) -> Vec<u8> {
    let request: S13SaveRequest = serde_json::from_value(json!({
        "determinism": {"seed": "0".repeat(64), "now": "2000-01-01T00:00:00.000Z"},
        "document": package["document"],
        "options": {"updateModifiedDate": false}
    }))
    .unwrap();
    write_docx_s13(request, source).unwrap()
}

fn document_xml(source: &[u8]) -> String {
    String::from_utf8(
        ooxml_opc::unzip_parts(source)
            .unwrap()
            .into_iter()
            .find(|(name, _)| name == "word/document.xml")
            .unwrap()
            .1,
    )
    .unwrap()
}

fn saved_document(source: &[u8]) -> String {
    document_xml(&saved_docx(source))
}

fn paragraph_texts(blocks: &[BlockContent]) -> Vec<String> {
    let mut texts = Vec::new();
    for block in blocks {
        match block {
            BlockContent::Paragraph(paragraph) => texts.push(get_paragraph_text(paragraph)),
            BlockContent::Table(table) => {
                for row in &table.rows {
                    for cell in &row.cells {
                        texts.extend(paragraph_texts(&cell.content));
                    }
                }
            }
            BlockContent::BlockSdt(sdt) => texts.extend(paragraph_texts(&sdt.content)),
            BlockContent::RawXml(_) => {}
        }
    }
    texts
}

fn repacked_field_document(body: &str, expected_texts: &[&str]) -> String {
    let source = fixture(body);
    let parsed = parse_docx_s9_wire(&source, S9ParseOptions::default()).unwrap();
    let texts = paragraph_texts(&parsed.document.package.document.content);
    assert_eq!(texts, expected_texts);
    let saved = saved_docx(&source);
    let reopened = parse_docx_s9_wire(&saved, S9ParseOptions::default()).unwrap();
    assert_eq!(
        paragraph_texts(&reopened.document.package.document.content),
        texts
    );
    document_xml(&saved)
}

fn rewritten_field_result(result: &str, runs: Value, expected_text: &str) -> String {
    let source = fixture(&format!(
        "<w:p>{BEGIN}{}{SEPARATE}{result}</w:p><w:p>{END}</w:p>",
        instruction("QUOTE")
    ));
    let parsed = parse_docx_s9_wire(&source, S9ParseOptions::default()).unwrap();
    let mut package = serde_json::to_value(parsed.document.package).unwrap();
    package["document"]
        .as_object_mut()
        .unwrap()
        .remove("sections");
    let field = package["document"]["content"][0]["content"]
        .as_array_mut()
        .unwrap()
        .iter_mut()
        .find(|node| node["type"] == "complexField")
        .unwrap();
    assert!(field["structuredResult"]["blocks"].is_array());
    field["fieldResult"] = runs;
    let saved = save_package(&source, &package);
    let reopened = parse_docx_s9_wire(&saved, S9ParseOptions::default()).unwrap();
    assert_eq!(
        paragraph_texts(&reopened.document.package.document.content),
        [expected_text, ""]
    );
    let xml = document_xml(&saved);
    assert_field(&xml, "QUOTE");
    assert_eq!(xml.matches("<w:bookmarkStart").count(), 1, "{xml}");
    assert_eq!(xml.matches("<w:bookmarkEnd").count(), 1, "{xml}");
    xml
}

fn normalized_field_result(result: &str, expected_text: &str) -> String {
    let body = format!(
        "<w:p><w:commentRangeStart w:id=\"99\"/>{BEGIN}{}{SEPARATE}{result}<w:commentRangeEnd w:id=\"99\"/></w:p><w:p>{END}</w:p>",
        instruction("QUOTE")
    );
    let source = comment_fixture(&body);
    let parsed = parse_docx_s9_wire(&source, S9ParseOptions::default()).unwrap();
    assert_eq!(
        paragraph_texts(&parsed.document.package.document.content),
        [expected_text, ""]
    );
    let mut package = serde_json::to_value(parsed.document.package).unwrap();
    assert!(
        package["document"]["comments"]
            .as_array()
            .unwrap()
            .iter()
            .any(|comment| comment["id"].as_f64() == Some(99.0))
    );
    let content = package["document"]["content"][0]["content"]
        .as_array()
        .unwrap();
    assert!(
        content
            .iter()
            .any(|node| { node["type"] == "commentRangeEnd" && node["id"].as_f64() == Some(99.0) })
    );
    let field = content
        .iter()
        .find(|node| node["type"] == "complexField")
        .unwrap();
    assert!(field["structuredResult"]["blocks"].is_array());
    assert_eq!(
        field["fieldResult"]
            .as_array()
            .unwrap()
            .iter()
            .flat_map(|run| run["content"].as_array().unwrap())
            .filter(|item| item["type"] == "commentReference")
            .count(),
        2
    );
    package["document"]
        .as_object_mut()
        .unwrap()
        .remove("sections");
    let saved = save_package(&source, &package);
    let reopened = parse_docx_s9_wire(&saved, S9ParseOptions::default()).unwrap();
    assert_eq!(
        paragraph_texts(&reopened.document.package.document.content),
        [expected_text, ""]
    );
    let xml = document_xml(&saved);
    assert_field(&xml, "QUOTE");
    assert_eq!(xml.matches("<w:bookmarkStart").count(), 1, "{xml}");
    assert_eq!(xml.matches("<w:bookmarkEnd").count(), 1, "{xml}");
    assert_eq!(xml.matches("<w:commentReference").count(), 1, "{xml}");
    xml
}

fn assert_field(xml: &str, instruction: &str) {
    for kind in ["begin", "separate", "end"] {
        assert_eq!(
            xml.matches(&format!("w:fldCharType=\"{kind}\"")).count(),
            1,
            "{xml}"
        );
    }
    assert_eq!(xml.matches("<w:instrText").count(), 1, "{xml}");
    assert_order(
        xml,
        &[
            "w:fldCharType=\"begin\"",
            "<w:instrText",
            instruction,
            "</w:instrText>",
            "w:fldCharType=\"separate\"",
            "w:fldCharType=\"end\"",
        ],
    );
}

fn assert_order(xml: &str, fragments: &[&str]) {
    let mut offset = 0;
    for fragment in fragments {
        let at = xml[offset..]
            .find(fragment)
            .unwrap_or_else(|| panic!("{fragment} missing or out of order: {xml}"));
        offset += at + fragment.len();
    }
}

const BEGIN: &str = r#"<w:r><w:fldChar w:fldCharType="begin"/></w:r>"#;
const SEPARATE: &str = r#"<w:r><w:fldChar w:fldCharType="separate"/></w:r>"#;
const END: &str = r#"<w:r><w:fldChar w:fldCharType="end"/></w:r>"#;
const START: &str = r#"<w:bookmarkStart w:id="0" w:name="Text1"/>"#;
const STOP: &str = r#"<w:bookmarkEnd w:id="0"/>"#;

fn instruction(text: &str) -> String {
    format!(r#"<w:r><w:instrText xml:space="preserve"> {text} </w:instrText></w:r>"#)
}

#[test]
fn a_form_field_bookmark_in_the_field_code_survives_a_save() {
    let body = format!(
        r#"<w:p>{BEGIN}{START}{}{SEPARATE}<w:r><w:t>Value</w:t></w:r>{END}{STOP}</w:p>"#,
        instruction("FORMTEXT")
    );
    let xml = saved_document(&fixture(&body));
    assert_order(
        &xml,
        &[
            "w:fldCharType=\"begin\"",
            START,
            "FORMTEXT",
            "w:fldCharType=\"separate\"",
            ">Value<",
            "w:fldCharType=\"end\"",
            STOP,
        ],
    );
}

#[test]
fn bookmarks_on_both_sides_of_the_field_separator_survive_a_save() {
    let body = format!(
        r#"<w:p>{BEGIN}{}{START}{SEPARATE}<w:r><w:t>Value</w:t></w:r>{STOP}{END}</w:p>"#,
        instruction("REF Target")
    );
    let xml = saved_document(&fixture(&body));
    assert_order(
        &xml,
        &[
            "REF Target",
            START,
            "w:fldCharType=\"separate\"",
            ">Value<",
            STOP,
            "w:fldCharType=\"end\"",
        ],
    );
}

#[test]
fn a_nested_field_in_the_field_code_survives_a_save() {
    let nested = format!(
        r#"{BEGIN}{}{SEPARATE}<w:r><w:t>inner</w:t></w:r>{END}"#,
        instruction("MERGEFIELD Name")
    );
    let body = format!(
        r#"<w:p>{BEGIN}{}{nested}{}{SEPARATE}<w:r><w:t>outer</w:t></w:r>{END}</w:p>"#,
        instruction("IF"),
        instruction("= \"x\" \"y\" \"z\""),
    );
    let xml = saved_document(&fixture(&body));
    assert_order(
        &xml,
        &[
            "IF",
            "MERGEFIELD Name",
            ">inner<",
            "&quot;y&quot;",
            ">outer<",
        ],
    );
    assert_eq!(xml.matches("w:fldCharType=\"begin\"").count(), 2, "{xml}");
}

#[test]
fn a_simple_field_in_the_field_code_is_written_once() {
    let body = format!(
        r#"<w:p>{BEGIN}{}<w:fldSimple w:instr="AUTHOR"><w:r><w:t>Ada</w:t></w:r></w:fldSimple>{START}{SEPARATE}<w:r><w:t>Ada</w:t></w:r>{END}{STOP}</w:p>"#,
        instruction("QUOTE")
    );
    let xml = saved_document(&fixture(&body));
    assert_eq!(xml.matches(">Ada<").count(), 2, "{xml}");
    assert_order(
        &xml,
        &[
            "QUOTE",
            "w:instr=\"AUTHOR\"",
            ">Ada<",
            START,
            "\"separate\"",
            ">Ada<",
            STOP,
        ],
    );
}

#[test]
fn an_edited_field_code_drops_its_old_nested_fields_and_keeps_its_bookmarks() {
    let nested = format!(
        r#"{BEGIN}{}{SEPARATE}<w:r><w:t>2</w:t></w:r>{END}"#,
        instruction("PAGE")
    );
    let body = format!(
        r#"<w:p>{BEGIN}{START}{}{nested}{}{SEPARATE}<w:r><w:t>2</w:t></w:r>{END}{STOP}</w:p>"#,
        instruction("="),
        instruction("/2+1"),
    );
    let source = fixture(&body);
    let parsed = parse_docx_s9_wire(&source, S9ParseOptions::default()).unwrap();
    let mut package: Value = serde_json::to_value(parsed.document.package).unwrap();
    package["document"]
        .as_object_mut()
        .unwrap()
        .remove("sections");
    let content = package["document"]["content"][0]["content"]
        .as_array_mut()
        .unwrap();
    let field = content
        .iter_mut()
        .find(|node| node["type"] == "complexField")
        .unwrap();
    field["instruction"] = json!("NUMPAGES");
    field["fieldCode"] = json!([]);
    let request: S13SaveRequest = serde_json::from_value(json!({
        "determinism": {"seed": "0".repeat(64), "now": "2000-01-01T00:00:00.000Z"},
        "document": package["document"],
        "options": {"updateModifiedDate": false}
    }))
    .unwrap();
    let saved = write_docx_s13(request, &source).unwrap();
    let xml = String::from_utf8(
        ooxml_opc::unzip_parts(&saved)
            .unwrap()
            .into_iter()
            .find(|(name, _)| name == "word/document.xml")
            .unwrap()
            .1,
    )
    .unwrap();
    assert!(!xml.contains("PAGE<") && !xml.contains(" PAGE "), "{xml}");
    assert_order(
        &xml,
        &[
            "\"begin\"",
            START,
            "NUMPAGES",
            "\"separate\"",
            "\"end\"",
            STOP,
        ],
    );
}

#[test]
fn a_nested_field_after_an_instruction_in_the_begin_run_survives_a_save() {
    let nested = format!(
        r#"{BEGIN}{}{SEPARATE}<w:r><w:t>4</w:t></w:r>{END}"#,
        instruction("PAGE")
    );
    let body = format!(
        r#"<w:p><w:r><w:fldChar w:fldCharType="begin"/><w:instrText xml:space="preserve"> = </w:instrText></w:r>{nested}{SEPARATE}<w:r><w:t>4</w:t></w:r>{END}</w:p>"#
    );
    let xml = saved_document(&fixture(&body));
    assert_order(
        &xml,
        &[
            "\"begin\"",
            " = ",
            "\"begin\"",
            " PAGE ",
            "\"separate\"",
            "\"end\"",
            "\"separate\"",
            "\"end\"",
        ],
    );
    assert_eq!(xml.matches("w:fldCharType=\"begin\"").count(), 2, "{xml}");
}

#[test]
fn hyperlink_result_bookmarks_in_a_table_cell_survive_repack() {
    let start1 = r#"<w:bookmarkStart w:id="21" w:name="_Hlt1"/>"#;
    let start2 = r#"<w:bookmarkStart w:id="22" w:name="_Hlt2"/>"#;
    let end1 = r#"<w:bookmarkEnd w:id="21"/>"#;
    let end2 = r#"<w:bookmarkEnd w:id="22"/>"#;
    for later_end in [false, true] {
        let ending = if later_end {
            format!("</w:p><w:p>{END}</w:p>")
        } else {
            format!("{END}</w:p>")
        };
        let body = format!(
            r#"<w:tbl><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid><w:tr><w:tc><w:p>{BEGIN}{}<w:r/>{SEPARATE}<w:r><w:t>Managin</w:t></w:r>{start1}{start2}<w:r><w:t>g</w:t></w:r>{end1}{end2}<w:r><w:rPr><w:rStyle w:val="Hyperlink"/></w:rPr><w:t> speed</w:t></w:r>{ending}</w:tc></w:tr></w:tbl>"#,
            instruction(r#"HYPERLINK "https://example.com/a""#)
        );
        let expected = if later_end {
            vec!["Managing speed", ""]
        } else {
            vec!["Managing speed"]
        };
        let xml = repacked_field_document(&body, &expected);
        assert_field(&xml, "HYPERLINK &quot;https://example.com/a&quot;");
        assert_eq!(xml.matches("<w:bookmarkStart").count(), 2, "{xml}");
        assert_eq!(xml.matches("<w:bookmarkEnd").count(), 2, "{xml}");
        assert_order(
            &xml,
            &[
                "<w:tc>",
                "w:fldCharType=\"separate\"",
                ">Managin<",
                start1,
                start2,
                ">g<",
                end1,
                end2,
                "<w:rStyle w:val=\"Hyperlink\"/>",
                "> speed<",
                "w:fldCharType=\"end\"",
                "</w:tc>",
            ],
        );
    }
}

#[test]
fn formtext_result_bookmark_with_a_body_level_end_survives_repack() {
    let begin = r#"<w:r><w:fldChar w:fldCharType="begin"><w:ffData><w:name w:val="Title"/><w:enabled/><w:calcOnExit w:val="0"/><w:textInput/></w:ffData></w:fldChar></w:r>"#;
    let start = r#"<w:bookmarkStart w:id="1" w:name="_GoBack"/>"#;
    let stop = r#"<w:bookmarkEnd w:id="1"/>"#;
    for (end, continuation) in [
        (END, String::new()),
        ("", String::new()),
        (
            "",
            format!("<w:p><w:r><w:t>continued</w:t></w:r>{END}</w:p>"),
        ),
    ] {
        let body = format!(
            "<w:p>{begin}{}{SEPARATE}{start}<w:r><w:t>Some title</w:t></w:r>{end}</w:p>{stop}{continuation}",
            instruction("FORMTEXT")
        );
        let expected = if continuation.is_empty() {
            vec!["Some title"]
        } else {
            vec!["Some title", "continued"]
        };
        let xml = repacked_field_document(&body, &expected);
        assert_field(&xml, "FORMTEXT");
        assert_eq!(xml.matches("<w:bookmarkStart").count(), 1, "{xml}");
        assert_eq!(xml.matches("<w:bookmarkEnd").count(), 1, "{xml}");
        assert_order(
            &xml,
            &[
                "w:fldCharType=\"separate\"",
                start,
                ">Some title<",
                "w:fldCharType=\"end\"",
                "</w:p>",
                stop,
            ],
        );
        if !continuation.is_empty() {
            assert_order(&xml, &[stop, ">continued<"]);
        }
    }
}

#[test]
fn hyperlink_result_bookmarks_in_a_heading_survive_repack() {
    let start = r#"<w:bookmarkStart w:id="83" w:name="_Toc1"/>"#;
    let stop = r#"<w:bookmarkEnd w:id="83"/>"#;
    for later_end in [false, true] {
        let ending = if later_end {
            format!("</w:p><w:p>{END}</w:p>")
        } else {
            format!("{END}</w:p>")
        };
        let body = format!(
            r#"<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr>{BEGIN}{}{SEPARATE}{start}<w:r><w:t>Heading text</w:t></w:r>{stop}{ending}"#,
            instruction(r#"HYPERLINK "https://example.com/b""#)
        );
        let expected = if later_end {
            vec!["Heading text", ""]
        } else {
            vec!["Heading text"]
        };
        let xml = repacked_field_document(&body, &expected);
        assert_field(&xml, "HYPERLINK &quot;https://example.com/b&quot;");
        assert_eq!(xml.matches("<w:bookmarkStart").count(), 1, "{xml}");
        assert_eq!(xml.matches("<w:bookmarkEnd").count(), 1, "{xml}");
        assert_order(
            &xml,
            &[
                "<w:pStyle w:val=\"Heading1\"/>",
                "w:fldCharType=\"separate\"",
                start,
                ">Heading text<",
                stop,
                "w:fldCharType=\"end\"",
                "</w:p>",
            ],
        );
    }
}

#[test]
fn split_result_runs_keep_the_entire_bookmark() {
    let result = format!("{START}<w:r><w:t>Heading text</w:t></w:r>{STOP}");
    let xml = rewritten_field_result(
        &result,
        json!([
            {"type": "run", "content": [{"type": "text", "text": "Heading "}]},
            {"type": "run", "formatting": {"bold": true}, "content": [{"type": "text", "text": "text"}]}
        ]),
        "Heading text",
    );
    assert_order(&xml, &[START, ">Heading <", "<w:b/>", ">text<", STOP]);
}

#[test]
fn removed_result_runs_keep_bookmarks_around_remaining_text() {
    let result = format!(
        "<w:commentRangeEnd w:id=\"99\"/><w:r><w:commentReference w:id=\"99\"/></w:r><w:r><w:commentReference w:id=\"99\"/></w:r>{START}<w:r><w:t>Heading text</w:t></w:r>{STOP}"
    );
    for (runs, expected_text) in [
        (
            json!([{"type": "run", "content": [{"type": "text", "text": "Heading text"}]}]),
            "Heading text",
        ),
        (
            json!([{"type": "run", "content": [{"type": "text", "text": "Heading"}]}]),
            "Heading",
        ),
        (
            json!([
                {"type": "run", "content": [{"type": "commentReference", "id": 99}]},
                {"type": "run", "content": [{"type": "commentReference", "id": 99}]},
                {"type": "run", "content": [{"type": "text", "text": "Heading text"}]}
            ]),
            "Heading text",
        ),
        (json!([]), ""),
    ] {
        let xml = rewritten_field_result(&result, runs, expected_text);
        if expected_text.is_empty() {
            assert_order(&xml, &[START, STOP]);
        } else {
            assert_order(&xml, &[START, &format!(">{expected_text}<"), STOP]);
        }
    }
}

#[test]
fn mid_run_bookmark_offsets_widen_to_whole_result_runs() {
    let result = format!(
        "<w:r><w:t>Head</w:t></w:r>{START}<w:r><w:t>ing</w:t></w:r>{STOP}<w:r><w:t> text</w:t></w:r>"
    );
    let xml = rewritten_field_result(
        &result,
        json!([{"type": "run", "content": [{"type": "text", "text": "Heading text"}]}]),
        "Heading text",
    );
    assert_order(&xml, &[START, "<w:r><w:t>Heading text</w:t></w:r>", STOP]);
}

#[test]
fn restyled_tab_result_keeps_its_enclosing_bookmark() {
    let result = format!("{START}<w:r><w:tab/></w:r>{STOP}");
    let xml = rewritten_field_result(
        &result,
        json!([{"type": "run", "formatting": {"bold": true}, "content": [{"type": "tab"}]}]),
        "",
    );
    assert_order(
        &xml,
        &[START, "<w:r><w:rPr><w:b/></w:rPr><w:tab/></w:r>", STOP],
    );
}

#[test]
fn duplicate_result_references_keep_both_text_runs_bookmarked() {
    let reference = r#"<w:r><w:commentReference w:id="99"/></w:r>"#;
    let result = format!(
        "{reference}{reference}{START}<w:r><w:t>A</w:t></w:r><w:r><w:t>B</w:t></w:r>{STOP}"
    );
    let xml = normalized_field_result(&result, "AB");
    assert_order(&xml, &[START, ">A<", ">B<", STOP]);
}

#[test]
fn removed_reference_between_result_runs_keeps_the_bookmark_end_in_place() {
    let reference = r#"<w:r><w:commentReference w:id="99"/></w:r>"#;
    let result = format!(
        "{reference}{START}<w:r><w:t>A</w:t></w:r>{reference}<w:r><w:t>B</w:t></w:r>{STOP}<w:r><w:t>C</w:t></w:r>"
    );
    let xml = normalized_field_result(&result, "ABC");
    assert_order(&xml, &[START, ">A<", ">B<", STOP, ">C<"]);
}
