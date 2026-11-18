use std::io::{self, Write};

use docx_parse::serializer::write_docx_s13;
use docx_parse::{BlockContent, S9ParseOptions, parse_docx_s9_wire};
use serde_json::{Value, json};

const W: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const R: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const PLAIN: usize = 4;
const CLOSER: &str = "CLOSERZ";
const FIELD_TAG: &str = r#""type":"complexField""#;
const NESTING_LIMIT: usize = 32;
/// Serialised body bytes allowed past `depth / 4` times the depth-4 body.
const BYTE_SLACK: usize = 4096;
/// Field entries allowed past `depth / 4` times the depth-4 count.
const FIELD_SLACK: usize = 4;
const BODY_BOOKMARK_START: &str = r#"<w:bookmarkStart w:id="0" w:name="Between"/>"#;
const BODY_BOOKMARK_END: &str = r#"<w:bookmarkEnd w:id="0"/>"#;
const FOREIGN_BLOCK: &str = r#"<ext:marker xmlns:ext="urn:example:foreign"/>"#;
/// Sources whose gates must ask the shared predicate whether a field spans paragraphs.
const GATE_SOURCES: [&str; 3] = [
    include_str!("../src/serializer/paragraph.rs"),
    include_str!("../src/serializer/comment_references.rs"),
    include_str!("../src/serializer/s13.rs"),
];
const BARE_SPAN_TESTS: [&str; 2] = ["blocks.is_none()", "blocks.is_some()"];

fn run_xml(inner: &str) -> String {
    format!("<w:r>{inner}</w:r>")
}

fn text_run(text: &str) -> String {
    run_xml(&format!(r#"<w:t xml:space="preserve">{text}</w:t>"#))
}

fn opener_text(level: usize) -> String {
    format!("OPN{level:02}Z")
}

fn result_text(level: usize) -> String {
    format!("RES{level:02}Z")
}

fn instruction(level: usize) -> String {
    format!(r#"QUOTE "L{level}""#)
}

fn plain_text(index: usize) -> String {
    format!("PLN{index:04}Z")
}

fn open_field(level: usize) -> String {
    [
        run_xml(r#"<w:fldChar w:fldCharType="begin"/>"#),
        run_xml(&format!(
            r#"<w:instrText xml:space="preserve"> {} </w:instrText>"#,
            instruction(level)
        )),
        run_xml(r#"<w:fldChar w:fldCharType="separate"/>"#),
        text_run(&result_text(level)),
    ]
    .concat()
}

/// Paragraph k opens field k and leaves it open; `PLAIN` plain paragraphs follow, then one
/// paragraph holding all `depth` ends.
fn stair_xml(depth: usize) -> String {
    let mut paragraphs = Vec::new();
    for level in 1..=depth {
        paragraphs.push(format!(
            "<w:p>{}{}</w:p>",
            text_run(&opener_text(level)),
            open_field(level)
        ));
    }
    for index in 0..PLAIN {
        paragraphs.push(format!("<w:p>{}</w:p>", text_run(&plain_text(index))));
    }
    paragraphs.push(format!(
        "<w:p>{}{}</w:p>",
        field_end().repeat(depth),
        text_run(CLOSER)
    ));
    paragraphs.concat()
}

fn field_end() -> String {
    run_xml(r#"<w:fldChar w:fldCharType="end"/>"#)
}

/// One field opened in the first paragraph and ended in the third, with a body-level bookmark
/// before the second paragraph and a foreign block before the third.
fn interrupted_xml() -> String {
    [
        format!("<w:p>{}{}</w:p>", text_run(&opener_text(1)), open_field(1)),
        BODY_BOOKMARK_START.to_owned(),
        format!("<w:p>{}</w:p>", text_run(&plain_text(0))),
        FOREIGN_BLOCK.to_owned(),
        format!("<w:p>{}{}</w:p>", field_end(), text_run(CLOSER)),
        BODY_BOOKMARK_END.to_owned(),
    ]
    .concat()
}

fn stair(depth: usize) -> Vec<u8> {
    package(&stair_xml(depth))
}

fn package(body: &str) -> Vec<u8> {
    let document = format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="{W}" xmlns:r="{R}"><w:body>{body}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>"#
    );
    ooxml_opc::rezip_parts(&[
        ("[Content_Types].xml".to_owned(), br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>"#.to_vec()),
        ("_rels/.rels".to_owned(), format!(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="{R}/officeDocument" Target="word/document.xml"/></Relationships>"#).into_bytes()),
        ("word/_rels/document.xml.rels".to_owned(), br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>"#.to_vec()),
        ("word/document.xml".to_owned(), document.into_bytes()),
    ])
    .unwrap()
}

fn parse_body(bytes: &[u8]) -> Vec<BlockContent> {
    parse_docx_s9_wire(bytes, S9ParseOptions::default())
        .unwrap()
        .document
        .package
        .document
        .content
}

fn body_values(bytes: &[u8]) -> Vec<Value> {
    parse_body(bytes)
        .iter()
        .map(|block| serde_json::to_value(block).unwrap())
        .collect()
}

struct Capped {
    bytes: Vec<u8>,
    cap: usize,
}

impl Write for Capped {
    fn write(&mut self, buf: &[u8]) -> io::Result<usize> {
        if self.bytes.len() + buf.len() > self.cap {
            return Err(io::Error::other("over the cap"));
        }
        self.bytes.extend_from_slice(buf);
        Ok(buf.len())
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}

/// The serialised body of the stair at `depth`, or `None` once it passes `cap` bytes.
fn capped_body(depth: usize, cap: usize) -> Option<Vec<u8>> {
    let body = parse_body(&stair(depth));
    let mut writer = Capped {
        bytes: Vec::new(),
        cap,
    };
    serde_json::to_writer(&mut writer, &body).ok()?;
    Some(writer.bytes)
}

fn field_entries(bytes: &[u8]) -> usize {
    bytes
        .windows(FIELD_TAG.len())
        .filter(|window| *window == FIELD_TAG.as_bytes())
        .count()
}

fn run_text(run: &Value) -> String {
    run["content"]
        .as_array()
        .into_iter()
        .flatten()
        .filter_map(|item| item["text"].as_str())
        .collect()
}

/// The text of the runs directly in a paragraph, outside its fields.
fn own_text(block: &Value) -> String {
    block["content"]
        .as_array()
        .into_iter()
        .flatten()
        .filter(|node| node["type"] == "run")
        .map(run_text)
        .collect()
}

fn top_level_fields(block: &Value) -> Vec<&Value> {
    block["content"]
        .as_array()
        .into_iter()
        .flatten()
        .filter(|node| node["type"] == "complexField")
        .collect()
}

/// The paragraphs `field` records as its result, resolved to their own text: those among the
/// `blockCount` body blocks after its owner paragraph `owner`, or the copies it carries.
fn recorded_result_texts(field: &Value, owner: usize, body: &[Value]) -> Vec<String> {
    let result = &field["structuredResult"];
    let recorded = match result["blockCount"].as_u64() {
        Some(count) => {
            let count = usize::try_from(count).unwrap();
            body[owner + 1..=owner + count].to_vec()
        }
        None => result["blocks"].as_array().cloned().unwrap_or_default(),
    };
    recorded
        .iter()
        .filter(|block| block["type"] != "rawXml")
        .map(own_text)
        .collect()
}

fn expected_paragraph_texts(depth: usize) -> Vec<String> {
    (1..=depth)
        .map(opener_text)
        .chain((0..PLAIN).map(plain_text))
        .chain([CLOSER.to_owned()])
        .collect()
}

fn expected_run_texts(depth: usize) -> Vec<String> {
    (1..=depth)
        .flat_map(|level| [opener_text(level), result_text(level)])
        .chain((0..PLAIN).map(plain_text))
        .chain([CLOSER.to_owned()])
        .collect()
}

fn save(original: &[u8], body: Vec<Value>) -> Vec<u8> {
    let request = json!({
        "determinism": {"seed": "0".repeat(64), "now": "2000-01-01T00:00:00.000Z"},
        "document": {"content": body},
        "options": {"updateModifiedDate": false}
    });
    let request = serde_json::from_str(&request.to_string()).unwrap();
    write_docx_s13(request, original).unwrap()
}

fn document_xml(bytes: &[u8]) -> String {
    let xml = ooxml_opc::unzip_parts(bytes)
        .unwrap()
        .into_iter()
        .find(|(path, _)| path == "word/document.xml")
        .unwrap()
        .1;
    String::from_utf8(xml).unwrap()
}

fn field_char_counts(xml: &str) -> [usize; 3] {
    ["begin", "separate", "end"].map(|kind| {
        xml.matches(format!(r#"w:fldCharType="{kind}""#).as_str())
            .count()
    })
}

/// The text inside each `tag` element of `xml`, in order.
fn tagged_texts(xml: &str, tag: &str) -> Vec<String> {
    let open = format!("<{tag}");
    let close = format!("</{tag}>");
    let mut texts = Vec::new();
    let mut rest = xml;
    while let Some(at) = rest.find(&open) {
        rest = &rest[at + open.len()..];
        if !rest.starts_with([' ', '>']) {
            continue;
        }
        let start = rest.find('>').unwrap() + 1;
        let end = rest.find(&close).unwrap();
        texts.push(rest[start..end].to_owned());
        rest = &rest[end..];
    }
    texts
}

fn on_a_large_stack(test: fn()) {
    std::thread::Builder::new()
        .stack_size(64 << 20)
        .spawn(test)
        .unwrap()
        .join()
        .unwrap();
}

#[test]
fn fields_open_across_paragraphs_serialise_linearly_in_depth() {
    let base = capped_body(4, usize::MAX).unwrap();
    let base_fields = field_entries(&base);
    for depth in [16, 8, 12] {
        let bound = depth / 4 * base.len() + BYTE_SLACK;
        let Some(body) = capped_body(depth, bound) else {
            panic!(
                "depth {depth}: body passed {bound} bytes; depth 4 was {}",
                base.len()
            );
        };
        let fields = field_entries(&body);
        assert!(
            (depth..=depth / 4 * base_fields + FIELD_SLACK).contains(&fields),
            "depth {depth}: {fields} field entries; depth 4 had {base_fields}"
        );
    }
}

#[test]
fn fields_open_across_paragraphs_at_the_nesting_limit_open_within_a_linear_budget() {
    on_a_large_stack(open_at_the_nesting_limit);
}

fn open_at_the_nesting_limit() {
    let base = capped_body(4, usize::MAX).unwrap();
    let bound = NESTING_LIMIT / 4 * base.len() + BYTE_SLACK;
    let Some(body) = capped_body(NESTING_LIMIT, bound) else {
        panic!(
            "depth {NESTING_LIMIT}: body passed {bound} bytes; depth 4 was {}",
            base.len()
        );
    };
    assert!(field_entries(&body) >= NESTING_LIMIT);
    let blocks: Vec<Value> = serde_json::from_slice(&body).unwrap();
    let texts: Vec<_> = blocks.iter().map(own_text).collect();
    assert_eq!(texts, expected_paragraph_texts(NESTING_LIMIT));
}

#[test]
fn fields_open_across_paragraphs_keep_each_level_and_its_later_paragraphs() {
    for depth in [4, 8] {
        let body = body_values(&stair(depth));
        let expected = expected_paragraph_texts(depth);
        let texts: Vec<_> = body.iter().map(own_text).collect();
        assert_eq!(texts, expected);
        for (index, paragraph) in body.iter().take(depth).enumerate() {
            let level = index + 1;
            let fields = top_level_fields(paragraph);
            assert_eq!(fields.len(), 1, "depth {depth} level {level}");
            let field = fields[0];
            assert_eq!(field["instruction"], instruction(level));
            let result: String = field["fieldResult"]
                .as_array()
                .unwrap()
                .iter()
                .map(run_text)
                .collect();
            assert_eq!(result, result_text(level));
            assert_eq!(
                recorded_result_texts(field, index, &body),
                expected[level..],
                "depth {depth} level {level}"
            );
            assert!(
                field["structuredCode"]["blocks"].is_null(),
                "depth {depth} level {level}"
            );
        }
    }
}

#[test]
fn a_field_open_across_a_body_bookmark_and_foreign_markup_counts_their_indexes() {
    let body = body_values(&package(&interrupted_xml()));
    let kinds: Vec<_> = body
        .iter()
        .map(|block| block["type"].as_str().unwrap())
        .collect();
    assert_eq!(
        kinds.join(" "),
        "paragraph rawXml paragraph rawXml paragraph rawXml"
    );
    let fields = top_level_fields(&body[0]);
    assert_eq!(fields.len(), 1);
    let field = fields[0];
    assert_eq!(field["instruction"], instruction(1));
    assert_eq!(
        recorded_result_texts(field, 0, &body),
        [plain_text(0), CLOSER.to_owned()]
    );
    let result = &field["structuredResult"];
    if let Some(count) = result["blockCount"].as_u64() {
        assert_eq!(count, 4);
        assert!(result["blocks"].is_null());
    }
    assert!(field["structuredCode"]["blockCount"].is_null());
    assert!(field["structuredCode"]["blocks"].is_null());
}

#[test]
fn spanning_field_gates_ask_the_shared_predicate() {
    for (index, source) in GATE_SOURCES.iter().enumerate() {
        for pattern in BARE_SPAN_TESTS {
            assert!(!source.contains(pattern), "source {index}: {pattern}");
        }
    }
}

#[test]
fn saving_fields_open_across_paragraphs_writes_each_field_and_paragraph_once() {
    on_a_large_stack(save_each_field_and_paragraph_once);
}

fn save_each_field_and_paragraph_once() {
    for depth in [4, 8] {
        let bytes = stair(depth);
        let saved = save(&bytes, body_values(&bytes));
        let xml = document_xml(&saved);
        assert_eq!(field_char_counts(&xml), [depth; 3]);
        let instructions: Vec<_> = tagged_texts(&xml, "w:instrText")
            .iter()
            .map(|text| text.trim().replace("&quot;", "\""))
            .collect();
        let expected: Vec<_> = (1..=depth).map(instruction).collect();
        assert_eq!(instructions, expected);
        assert_eq!(tagged_texts(&xml, "w:t"), expected_run_texts(depth));
        let reopened: Vec<_> = body_values(&saved).iter().map(own_text).collect();
        assert_eq!(reopened, expected_paragraph_texts(depth));
    }
}
