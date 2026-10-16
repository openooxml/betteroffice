use std::alloc::{GlobalAlloc, Layout, System};
use std::cell::Cell;

use docx_parse::serializer::{S13SaveRequest, write_docx_s13};
use docx_parse::{ParseError, S9ParseOptions, parse_docx_s9_wire};
use serde_json::{Value, json};

#[global_allocator]
static ALLOCATOR: CountingAllocator = CountingAllocator;

struct CountingAllocator;

thread_local! {
    static ALLOCATED: Cell<u64> = const { Cell::new(0) };
}

unsafe impl GlobalAlloc for CountingAllocator {
    unsafe fn alloc(&self, layout: Layout) -> *mut u8 {
        let _ = ALLOCATED.try_with(|counter| counter.set(counter.get() + layout.size() as u64));
        unsafe { System.alloc(layout) }
    }

    unsafe fn dealloc(&self, pointer: *mut u8, layout: Layout) {
        unsafe { System.dealloc(pointer, layout) }
    }
}

const W: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const R: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const MARKER: &str = "nested banner";
const INSTRUCTION: &str = r#"INCLUDEPICTURE "https://example.test/banner.png" \* MERGEFORMATINET"#;

fn run(content: &str) -> String {
    format!(
        r#"<w:r><w:rPr><w:rFonts w:ascii="Segoe UI" w:hAnsi="Segoe UI" w:cs="Segoe UI"/><w:color w:val="444444"/><w:szCs w:val="20"/><w:lang w:val="en-US"/></w:rPr>{content}</w:r>"#
    )
}

fn picture() -> String {
    run(&format!(
        r##"<w:pict><v:shape id="_x0000_i1025" type="#_x0000_t75" alt="{MARKER}" style="width:450pt;height:59.25pt"><v:imagedata r:id="rId1" r:href="rId2"/></v:shape></w:pict>"##
    ))
}

/// The runs opening `depth` INCLUDEPICTURE fields, each one's result holding the next.
fn complex_open(depth: usize) -> String {
    [
        run(r#"<w:fldChar w:fldCharType="begin"/>"#),
        run(&format!(
            r#"<w:instrText xml:space="preserve"> {INSTRUCTION} </w:instrText>"#
        )),
        run(r#"<w:fldChar w:fldCharType="separate"/>"#),
    ]
    .concat()
    .repeat(depth)
}

fn complex_close(depth: usize) -> String {
    run(r#"<w:fldChar w:fldCharType="end"/>"#).repeat(depth)
}

/// `depth` INCLUDEPICTURE fields, each one's result holding the next, around one picture.
fn complex_chain(depth: usize) -> String {
    format!(
        "{}{}{}",
        complex_open(depth),
        picture(),
        complex_close(depth)
    )
}

fn simple_chain(depth: usize) -> String {
    let open = format!(
        r#"<w:fldSimple w:instr="{}">"#,
        INSTRUCTION.replace('"', "&quot;")
    );
    format!(
        "{}{}{}",
        open.repeat(depth),
        picture(),
        "</w:fldSimple>".repeat(depth)
    )
}

fn png() -> Vec<u8> {
    let mut bytes = b"\x89PNG\r\n\x1a\n\0\0\0\x0dIHDR".to_vec();
    bytes.extend(600u32.to_be_bytes());
    bytes.extend(79u32.to_be_bytes());
    bytes.extend([8, 6, 0, 0, 0]);
    bytes
}

fn package(body: &str, header: &str) -> Vec<u8> {
    let namespaces = format!(
        r#"xmlns:w="{W}" xmlns:r="{R}" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office""#
    );
    let rels = |extra: &str| {
        format!(
            r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="{R}/image" Target="media/image1.png"/><Relationship Id="rId2" Type="{R}/image" Target="https://example.test/banner.png" TargetMode="External"/>{extra}</Relationships>"#
        )
        .into_bytes()
    };
    ooxml_opc::rezip_parts(&[
        ("[Content_Types].xml".to_owned(), br#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/></Types>"#.to_vec()),
        ("_rels/.rels".to_owned(), format!(r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="{R}/officeDocument" Target="word/document.xml"/></Relationships>"#).into_bytes()),
        ("word/_rels/document.xml.rels".to_owned(), rels(&format!(r#"<Relationship Id="rIdHeader" Type="{R}/header" Target="header1.xml"/>"#))),
        ("word/_rels/header1.xml.rels".to_owned(), rels("")),
        ("word/document.xml".to_owned(), format!(r#"<w:document {namespaces}><w:body><w:p>{body}</w:p><w:sectPr><w:headerReference w:type="default" r:id="rIdHeader"/></w:sectPr></w:body></w:document>"#).into_bytes()),
        ("word/header1.xml".to_owned(), format!(r#"<w:hdr {namespaces}><w:p>{header}</w:p></w:hdr>"#).into_bytes()),
        ("word/media/image1.png".to_owned(), png()),
    ])
    .unwrap()
}

/// The body and header paragraphs of a parsed package.
fn parse(bytes: &[u8]) -> (Value, Value) {
    let package = parse_docx_s9_wire(bytes, S9ParseOptions::default())
        .unwrap()
        .document
        .package;
    let header = &package.header_entries.as_ref().unwrap()[0].1.content;
    (
        serde_json::to_value(&package.document.content[0]).unwrap(),
        serde_json::to_value(&header[0]).unwrap(),
    )
}

/// Complex fields, simple fields and picture copies anywhere in `value`.
fn census(value: &Value) -> [usize; 3] {
    let mut counts = [0; 3];
    let mut pending = vec![value];
    while let Some(value) = pending.pop() {
        match value {
            Value::Object(map) => {
                match map.get("type").and_then(Value::as_str) {
                    Some("complexField") => counts[0] += 1,
                    Some("simpleField") => counts[1] += 1,
                    _ => {}
                }
                pending.extend(map.values());
            }
            Value::Array(values) => pending.extend(values),
            Value::String(text) if text.contains(MARKER) => counts[2] += 1,
            _ => {}
        }
    }
    counts
}

/// Follows the first node of each structured result down from the paragraph's
/// outermost field: how many fields it passes and the node it ends on.
fn innermost_result(paragraph: &Value) -> (usize, &Value) {
    let mut node = &paragraph["content"][0];
    let mut levels = 0;
    while matches!(node["type"].as_str(), Some("complexField" | "simpleField")) {
        assert_eq!(node["instruction"], INSTRUCTION);
        levels += 1;
        node = &node["structuredResult"]["inline"][0];
    }
    (levels, node)
}

/// A save request for `body`, its object keys sorted as `serde_json::Value` writes them.
fn save_request(body: Value) -> String {
    json!({
        "determinism": {"seed": "0".repeat(64), "now": "2000-01-01T00:00:00.000Z"},
        "document": {"content": [body]},
        "options": {"updateModifiedDate": false}
    })
    .to_string()
}

fn save(original: &[u8], body: Value) -> Vec<u8> {
    let request = serde_json::from_str(&save_request(body)).unwrap();
    write_docx_s13(request, original).unwrap()
}

/// Bytes allocated decoding the save request of a body paragraph with `depth` nested fields.
fn save_request_allocation(depth: usize) -> u64 {
    let (body, _) = parse(&package(&complex_chain(depth), ""));
    let request = save_request(body);
    let before = ALLOCATED.with(Cell::get);
    let decoded = serde_json::from_str::<S13SaveRequest>(&request);
    let allocated = ALLOCATED.with(Cell::get) - before;
    assert!(decoded.is_ok());
    allocated
}

fn field_char_counts(xml: &str) -> [usize; 3] {
    ["begin", "separate", "end"].map(|kind| {
        xml.matches(format!(r#"w:fldCharType="{kind}""#).as_str())
            .count()
    })
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

#[test]
fn nested_fields_hold_each_level_once() {
    for (chain, kind) in [(complex_chain as fn(usize) -> String, 0), (simple_chain, 1)] {
        let (single, _) = parse(&package(&chain(1), ""));
        let pictures = census(&single)[2];
        assert!(pictures > 0);
        let leaf = &single["content"][0];
        assert_eq!(leaf["fieldTree"]["result"], leaf["structuredResult"]);
        let (body, header) = parse(&package(&chain(10), &chain(10)));
        for paragraph in [&body, &header] {
            let counts = census(paragraph);
            assert_eq!(counts[kind], 10);
            assert_eq!(counts[2], pictures);
            assert!(paragraph["content"][0]["fieldTree"].get("result").is_none());
            let (levels, inner) = innermost_result(paragraph);
            assert_eq!(levels, 10);
            assert_eq!(inner["type"], "run");
            assert!(census(inner)[2] > 0);
        }
    }
}

#[test]
fn saving_nested_fields_is_not_exponential_in_depth() {
    let allocations = [8, 12, 16].map(save_request_allocation);
    for pair in allocations.windows(2) {
        assert!(
            pair[1] < 4 * pair[0],
            "{allocations:?} bytes at depths 8, 12, 16"
        );
    }
}

#[test]
fn fields_nested_past_the_limit_open_and_save_every_level() {
    on_a_large_stack(open_and_save_past_the_limit);
}

// Debug builds overflow a 2 MB thread from depth 33; release at 2 MB and wasm open depth 40.
fn on_a_large_stack(test: fn()) {
    std::thread::Builder::new()
        .stack_size(64 << 20)
        .spawn(test)
        .unwrap()
        .join()
        .unwrap();
}

fn open_and_save_past_the_limit() {
    let depth = 40;
    let bytes = package(&complex_chain(depth), &complex_chain(depth));
    let (body, header) = parse(&bytes);
    for paragraph in [&body, &header] {
        let (levels, _) = innermost_result(paragraph);
        assert!((1..depth).contains(&levels));
        assert_eq!(census(paragraph)[0], levels);
        assert!(census(paragraph)[2] > 0);
    }
    let saved = save(&bytes, body.clone());
    let xml = document_xml(&saved);
    for kind in ["begin", "separate", "end"] {
        let marker = format!(r#"w:fldCharType="{kind}""#);
        assert_eq!(xml.matches(marker.as_str()).count(), depth);
    }
    assert_eq!(xml.matches("INCLUDEPICTURE").count(), depth);
    let (reopened, _) = parse(&saved);
    assert_eq!(innermost_result(&reopened).0, innermost_result(&body).0);
    assert!(census(&reopened)[2] > 0);
}

#[test]
fn field_characters_past_the_limit_save_as_often_as_they_open() {
    on_a_large_stack(save_field_characters_past_the_limit);
}

fn save_field_characters_past_the_limit() {
    let two = |kind: &str| {
        run(&format!(
            r#"<w:fldChar w:fldCharType="{kind}"/><w:fldChar w:fldCharType="{kind}"/>"#
        ))
    };
    let body = format!(
        "{}{}{}{}{}",
        complex_open(33),
        two("begin"),
        picture(),
        two("end"),
        complex_close(33)
    );
    let bytes = package(&body, "");
    let (paragraph, _) = parse(&bytes);
    let saved = document_xml(&save(&bytes, paragraph));
    assert_eq!(field_char_counts(&saved), [35, 33, 35]);
    assert_eq!(field_char_counts(&saved), field_char_counts(&body));
}

#[test]
fn fields_past_the_limit_that_leave_their_scope_are_refused() {
    let open = format!("{}{}", complex_open(33), picture());
    for body in [
        format!(
            r#"{open}<w:hyperlink w:anchor="top">{}</w:hyperlink>{}"#,
            complex_close(1),
            complex_close(32)
        ),
        format!("{open}</w:p><w:p>{}", complex_close(33)),
        format!(
            "{open}{}{}",
            run(r#"<w:fldChar w:fldCharType="end"/><w:fldChar w:fldCharType="end"/>"#),
            complex_close(31)
        ),
    ] {
        let error = parse_docx_s9_wire(&package(&body, ""), S9ParseOptions::default()).err();
        assert!(
            matches!(
                error,
                Some(ParseError::ResourceLimit {
                    kind: "fieldDepth",
                    ..
                })
            ),
            "{error:?}"
        );
    }
}
