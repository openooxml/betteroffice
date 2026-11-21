use docx_parse::serializer::{S13SaveRequest, write_docx_s13};
use docx_parse::{S9ParseOptions, parse_docx_s5_projection, parse_docx_s9_wire};
use serde_json::{Value, json};

const W: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const MARK: &str = "<w:lastRenderedPageBreak/>";
const FORM_DATA: &str = r#"<w:ffData><w:name w:val="Text1"/></w:ffData>"#;
const CONTENT_TYPES: &str = concat!(
    r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">"#,
    r#"<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>"#,
    r#"<Default Extension="xml" ContentType="application/xml"/>"#,
    r#"<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>"#,
    "</Types>"
);
const ROOT_RELS: &str = concat!(
    r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">"#,
    r#"<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>"#,
    "</Relationships>"
);
const STYLES_TYPE: &str = concat!(
    r#"<Override PartName="/word/styles.xml" "#,
    r#"ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>"#
);
const DOCUMENT_RELS: &str = concat!(
    r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">"#,
    r#"<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>"#,
    "</Relationships>"
);
const STYLES: &str = concat!(
    r#"<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">"#,
    r#"<w:style w:type="character" w:styleId="Raised"><w:name w:val="Raised"/>"#,
    r#"<w:rPr><w:vertAlign w:val="superscript"/></w:rPr></w:style></w:styles>"#
);
const BASELINE: &str = r#"<w:vertAlign w:val="baseline"/>"#;

fn chr(kind: &str) -> String {
    format!(r#"<w:fldChar w:fldCharType="{kind}"/>"#)
}

fn instr(code: &str) -> String {
    format!(r#"<w:instrText xml:space="preserve"> {code} </w:instrText>"#)
}

fn text(value: &str) -> String {
    format!(r#"<w:t xml:space="preserve">{value}</w:t>"#)
}

fn run(content: &str) -> String {
    format!("<w:r>{content}</w:r>")
}

fn whole_field(code: &str) -> String {
    [chr("begin"), instr(code), chr("separate"), chr("end")].concat()
}

fn page_of_pages() -> String {
    [whole_field("PAGE"), text(" of "), whole_field("NUMPAGES")].concat()
}

fn field_run() -> String {
    run(&page_of_pages())
}

fn between(middle: &str) -> String {
    format!(
        "<w:p>{}{middle}{}</w:p>",
        run(&text("AB")),
        run(&text("CD"))
    )
}

fn package(body: &str) -> Vec<u8> {
    let document =
        format!(r#"<w:document xmlns:w="{W}"><w:body>{body}<w:sectPr/></w:body></w:document>"#);
    ooxml_opc::rezip_parts(&[
        (
            "[Content_Types].xml".to_owned(),
            CONTENT_TYPES.as_bytes().to_vec(),
        ),
        ("_rels/.rels".to_owned(), ROOT_RELS.as_bytes().to_vec()),
        ("word/document.xml".to_owned(), document.into_bytes()),
    ])
    .unwrap()
}

/// `package(body)` with a styles part holding the character style `Raised`, which raises its
/// text.
fn styled_package(body: &str) -> Vec<u8> {
    let mut parts = ooxml_opc::unzip_parts(&package(body)).unwrap();
    parts.retain(|(path, _)| path != "[Content_Types].xml");
    let types = CONTENT_TYPES.replace("</Types>", &format!("{STYLES_TYPE}</Types>"));
    parts.extend([
        ("[Content_Types].xml".to_owned(), types.into_bytes()),
        (
            "word/_rels/document.xml.rels".to_owned(),
            DOCUMENT_RELS.as_bytes().to_vec(),
        ),
        ("word/styles.xml".to_owned(), STYLES.as_bytes().to_vec()),
    ]);
    ooxml_opc::rezip_parts(&parts).unwrap()
}

/// The two fields in one run whose properties set the baseline over the style `Raised`.
fn raised_run() -> String {
    let properties = format!(r#"<w:rPr><w:rStyle w:val="Raised"/>{BASELINE}</w:rPr>"#);
    run(&format!("{properties}{}", page_of_pages()))
}

fn body_paragraph(bytes: &[u8], index: usize) -> Vec<Value> {
    let parsed = parse_docx_s9_wire(bytes, S9ParseOptions::default()).unwrap();
    let content = &parsed.document.package.document.content;
    let paragraph = serde_json::to_value(&content[index]).unwrap();
    paragraph["content"].as_array().unwrap().clone()
}

fn container_paragraph(bytes: &[u8], index: usize) -> Vec<Value> {
    let projection = parse_docx_s5_projection(bytes).unwrap();
    let part = projection
        .xml_parts
        .iter()
        .find(|part| part.path == "word/document.xml")
        .unwrap();
    let inlines = serde_json::to_value(&part.paragraph_inlines[index]).unwrap();
    inlines.as_array().unwrap().clone()
}

fn both_walkers(bytes: &[u8], index: usize) -> [Vec<Value>; 2] {
    [
        body_paragraph(bytes, index),
        container_paragraph(bytes, index),
    ]
}

fn kinds(nodes: &[Value]) -> Vec<&str> {
    nodes
        .iter()
        .map(|node| node["type"].as_str().unwrap_or_default())
        .collect()
}

fn save(source: &[u8], edit: impl FnOnce(&mut Value)) -> String {
    let parsed = parse_docx_s9_wire(source, S9ParseOptions::default()).unwrap();
    let mut package: Value = serde_json::to_value(parsed.document.package).unwrap();
    package["document"]
        .as_object_mut()
        .unwrap()
        .remove("sections");
    let mut request = json!({
        "determinism": {"seed": "0".repeat(64), "now": "2000-01-01T00:00:00.000Z"},
        "document": package["document"],
        "options": {"updateModifiedDate": false}
    });
    edit(&mut request["document"]);
    let request: S13SaveRequest = serde_json::from_value(request).unwrap();
    let saved = write_docx_s13(request, source).unwrap();
    let (_, xml) = ooxml_opc::unzip_parts(&saved)
        .unwrap()
        .into_iter()
        .find(|(path, _)| path == "word/document.xml")
        .unwrap();
    String::from_utf8(xml).unwrap()
}

fn visible(xml: &str) -> String {
    let mut output = String::new();
    for piece in xml.split("<w:t").skip(1) {
        let body = match piece.strip_prefix('>') {
            Some(body) => body,
            None => match piece
                .strip_prefix(' ')
                .and_then(|rest| rest.split_once('>'))
            {
                Some((_, body)) => body,
                None => continue,
            },
        };
        output.push_str(body.split("</w:t>").next().unwrap_or_default());
    }
    output
}

/// The text written before and after the one copy of `source` in its paragraph.
fn around(xml: &str, source: &str) -> Option<(String, String)> {
    if xml.matches(source).count() != 1 {
        return None;
    }
    let at = xml.find(source)?;
    let start = xml[..at].rfind("<w:p>").max(xml[..at].rfind("<w:p "))?;
    let after = at + source.len();
    let end = after + xml[after..].find("</w:p>")?;
    Some((visible(&xml[start..at]), visible(&xml[after..end])))
}

fn pair(before: &str, after: &str) -> Option<(String, String)> {
    Some((before.to_owned(), after.to_owned()))
}

fn without_source(value: &mut Value) {
    match value {
        Value::Object(object) => {
            object.remove("sourceRunXml");
            object.remove("sourceRunDigest");
            object.values_mut().for_each(without_source);
        }
        Value::Array(items) => items.iter_mut().for_each(without_source),
        _ => {}
    }
}

/// The document saved after `edit`, and saved after `edit` without the kept run, as on main.
fn save_and_main(source: &[u8], edit: impl Fn(&mut Value)) -> (String, String) {
    let saved = save(source, &edit);
    let main = save(source, |document| {
        edit(document);
        without_source(document);
    });
    (saved, main)
}

fn field_of(document: &mut Value) -> &mut Value {
    &mut document["content"][0]["content"][1]
}

/// The content of the first tracked change in `value`.
fn tracked_content(value: &mut Value) -> Option<&mut Value> {
    let tracked = matches!(
        value["type"].as_str(),
        Some("deletion" | "moveFrom" | "insertion")
    );
    if tracked {
        return value.get_mut("content");
    }
    match value {
        Value::Object(object) => object.values_mut().find_map(tracked_content),
        Value::Array(items) => items.iter_mut().find_map(tracked_content),
        _ => None,
    }
}

/// The text written from the first rendered page break to the end of its paragraph.
fn after_mark(xml: &str) -> Option<String> {
    let at = xml.find(MARK)?;
    let end = at + xml[at..].find("</w:p>")?;
    Some(visible(&xml[at..end]))
}

/// Saves `bytes` with and without the kept run, asserts both write the same rendered page
/// breaks at the same text, and returns the save with it.
fn marks_as_main(bytes: &[u8]) -> String {
    let (saved, main) = save_and_main(bytes, |_| {});
    assert_eq!(
        saved.matches(MARK).count(),
        main.matches(MARK).count(),
        "{saved}\n{main}"
    );
    assert_eq!(after_mark(&saved), after_mark(&main), "{saved}\n{main}");
    saved
}

#[test]
fn the_run_parses_to_the_field_main_parses() {
    let bytes = package(&between(&field_run()));
    for content in both_walkers(&bytes, 0) {
        assert_eq!(
            kinds(&content),
            ["run", "complexField", "run"],
            "{content:?}"
        );
        let field = &content[1];
        assert_eq!(field["instruction"], "PAGE  NUMPAGES", "{field}");
        assert_eq!(field["fieldType"], "PAGE", "{field}");
        assert_eq!(field["fieldCode"], json!([]), "{field}");
        assert_eq!(field["fieldResult"], json!([]), "{field}");
        assert!(field.get("structuredCode").is_none(), "{field}");
        assert!(field.get("structuredResult").is_none(), "{field}");
    }
}

#[test]
fn the_field_carries_its_run_as_read() {
    let source = field_run();
    let bytes = package(&between(&source));
    for content in both_walkers(&bytes, 0) {
        assert_eq!(content[1]["sourceRunXml"], source.as_str(), "{content:?}");
        let digest = content[1]["sourceRunDigest"].as_str().unwrap_or_default();
        assert_eq!(digest.len(), 64, "{content:?}");
    }
}

#[test]
fn a_full_rewrite_writes_the_run_as_read_between_its_neighbours() {
    let source = field_run();
    let xml = save(&package(&between(&source)), |_| {});
    assert_eq!(around(&xml, &source), pair("AB", "CD"), "{xml}");
}

#[test]
fn a_written_result_writes_the_field_from_the_model() {
    let source = field_run();
    let xml = save(&package(&between(&source)), |document| {
        document["content"][0]["content"][1]["fieldResult"] =
            json!([{"type": "run", "content": [{"type": "text", "text": "9"}]}]);
    });
    assert!(!xml.contains(&source), "{xml}");
    assert!(!xml.contains(" of "), "{xml}");
    assert!(xml.contains(">PAGE  NUMPAGES</w:instrText>"), "{xml}");
    assert!(xml.contains("<w:t>9</w:t>"), "{xml}");
}

#[test]
fn an_edited_field_is_written_as_main_writes_it() {
    let source = field_run();
    let bytes = package(&between(&source));
    let edits: [fn(&mut Value); 3] = [
        |document| field_of(document)["instruction"] = json!("DATE"),
        |document| field_of(document)["fldLock"] = json!(true),
        |document| field_of(document)["formatting"] = json!({"bold": true}),
    ];
    for edit in edits {
        let (saved, main) = save_and_main(&bytes, edit);
        assert!(!saved.contains(&source), "{saved}");
        assert_eq!(saved, main);
    }
}

#[test]
fn a_removed_baseline_over_a_raised_style_is_written_as_main_writes_it() {
    let source = raised_run();
    let bytes = styled_package(&between(&source));
    let (saved, main) = save_and_main(&bytes, |document| {
        let formatting = field_of(document)["formatting"].as_object_mut().unwrap();
        assert_eq!(formatting.remove("vertAlign"), Some(json!("baseline")));
    });
    assert!(!saved.contains(&source), "{saved}");
    assert!(!saved.contains(BASELINE), "{saved}");
    assert!(saved.contains(">PAGE  NUMPAGES</w:instrText>"), "{saved}");
    assert_eq!(saved, main);
}

#[test]
fn an_untouched_field_over_a_raised_style_is_written_as_read() {
    let source = raised_run();
    let xml = save(&styled_package(&between(&source)), |_| {});
    assert_eq!(around(&xml, &source), pair("AB", "CD"), "{xml}");
}

#[test]
fn a_field_moved_into_a_tracked_change_in_a_control_is_written_as_main_writes_it() {
    let source = field_run();
    for (wrapper, tag) in [
        ("del", "w:delText"),
        ("moveFrom", "w:delText"),
        ("ins", "w:t"),
    ] {
        let change = format!(
            r#"<w:{wrapper} w:id="1" w:author="A">{}</w:{wrapper}>"#,
            run(&format!("<{tag}>x</{tag}>"))
        );
        let control = format!("<w:p><w:sdt><w:sdtContent>{change}</w:sdtContent></w:sdt></w:p>");
        let bytes = package(&format!("{}{control}", between(&source)));
        let moved = |document: &mut Value| {
            let field = field_of(document).clone();
            *tracked_content(&mut document["content"][1]).unwrap() = json!([field]);
        };
        let saved = save(&bytes, moved);
        let main = save(&bytes, |document| {
            moved(document);
            without_source(&mut document["content"][1]);
        });
        assert_eq!(saved.matches(&source).count(), 1, "{saved}");
        assert_eq!(saved, main);
    }
}

#[test]
fn a_rendered_page_break_inside_the_run_is_written_where_main_writes_it() {
    let kept = field_run();
    let marked = run(&format!("{MARK}{}", page_of_pages()));
    let bytes = package(&format!("<w:p>{marked}{}</w:p>", run(&text("CD"))));
    for content in both_walkers(&bytes, 0) {
        assert_eq!(content[0]["sourceRunXml"], kept.as_str(), "{content:?}");
    }
    let saved = marks_as_main(&bytes);
    assert_eq!(around(&saved, &kept), pair("", "CD"), "{saved}");
}

#[test]
fn a_rendered_page_break_after_run_properties_is_written_where_main_writes_it() {
    let properties = "<w:rPr><w:b/></w:rPr>";
    let kept = run(&format!("{properties}{}", page_of_pages()));
    let marked = run(&format!("{properties}{MARK}{}", page_of_pages()));
    let bytes = package(&format!("<w:p>{marked}{}</w:p>", run(&text("CD"))));
    for content in both_walkers(&bytes, 0) {
        assert_eq!(content[0]["sourceRunXml"], kept.as_str(), "{content:?}");
    }
    let saved = marks_as_main(&bytes);
    assert_eq!(around(&saved, &kept), pair("", "CD"), "{saved}");
    assert!(!saved.contains(&format!("<w:r>{MARK}<w:rPr>")), "{saved}");
}

#[test]
fn a_rendered_page_break_before_the_run_is_written_once_before_it() {
    let source = field_run();
    let first = run(&format!("{MARK}{}", text("AB")));
    let bytes = package(&format!("<w:p>{first}{source}{}</w:p>", run(&text("CD"))));
    let xml = save(&bytes, |_| {});
    assert_eq!(xml.matches(MARK).count(), 1, "{xml}");
    assert!(xml.find(MARK) < xml.find("AB"), "{xml}");
    assert_eq!(around(&xml, &source), pair("AB", "CD"), "{xml}");
}

#[test]
fn other_field_runs_carry_no_source_run() {
    let nested = [
        chr("begin"),
        instr("PAGE"),
        whole_field("NUMPAGES"),
        chr("separate"),
        chr("end"),
        text(" of "),
        whole_field("NUMPAGES"),
    ]
    .concat();
    let open_around = [
        run(&chr("begin")),
        run(&instr("IF 1 = 1")),
        run(&chr("separate")),
        field_run(),
        run(&chr("end")),
    ]
    .concat();
    let unbalanced = [
        run(&format!("{}{}", page_of_pages(), chr("begin"))),
        run(&instr("DATE")),
        run(&chr("separate")),
        run(&chr("end")),
    ]
    .concat();
    let form = [
        format!(r#"<w:fldChar w:fldCharType="begin">{FORM_DATA}</w:fldChar>"#),
        instr("FORMTEXT"),
        chr("separate"),
        chr("end"),
        text(" of "),
        whole_field("NUMPAGES"),
    ]
    .concat();
    let controls = [
        run(&whole_field("PAGE")),
        run(&nested),
        open_around,
        unbalanced,
        run(&form),
        format!(
            r#"<w:hyperlink w:anchor="target">{}</w:hyperlink>"#,
            field_run()
        ),
    ];
    for control in controls {
        let bytes = package(&between(&control));
        for content in both_walkers(&bytes, 0) {
            let serialized = serde_json::to_string(&content).unwrap();
            assert!(
                !serialized.contains("sourceRunXml"),
                "{control}\n{serialized}"
            );
        }
    }
}
