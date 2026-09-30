use docx_edit::bridge::{RenderEnv, yrs_doc_to_layout_blocks};
use docx_edit::{EditingDoc, seed_from_docx};
use serde_json::Value;

const W: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

fn run(text: &str) -> String {
    format!(r#"<w:r><w:t xml:space="preserve">{text}</w:t></w:r>"#)
}

/// A complex field whose cached result is "7".
fn field(instruction: &str) -> String {
    format!(
        r#"<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> {instruction} </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>7</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>"#
    )
}

fn paragraph(content: &str) -> String {
    format!("<w:p>{}{content}</w:p>", run("Caption "))
}

fn text_box(content: &str) -> String {
    format!(
        r#"<w:r><w:drawing><wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="1" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="margin"><wp:align>right</wp:align></wp:positionH><wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV><wp:extent cx="1600000" cy="228600"/><wp:wrapNone/><wp:docPr id="1" name="Box 1"/><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp><wps:cNvSpPr txBox="1"/><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1600000" cy="228600"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></wps:spPr><wps:txbx><w:txbxContent><w:p>{content}</w:p></w:txbxContent></wps:txbx><wps:bodyPr/></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>"#
    )
}

fn document(body: &str) -> Vec<u8> {
    let parts = [
        ("[Content_Types].xml", r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>"#.to_owned()),
        ("_rels/.rels", r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="doc" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"#.to_owned()),
        ("word/document.xml", format!(r#"<w:document xmlns:w="{W}" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><w:body>{body}<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr></w:body></w:document>"#)),
    ];
    ooxml_opc::rezip_parts(
        &parts
            .into_iter()
            .map(|(name, value)| (name.to_owned(), value.into_bytes()))
            .collect::<Vec<_>>(),
    )
    .unwrap()
}

/// The results of the body's SEQ fields in lowering order, each with whether
/// it sits in a text box or shape.
fn boxed_sequence_results(body: &str) -> Vec<(bool, String)> {
    fn collect(value: &Value, boxed: bool, out: &mut Vec<(bool, String)>) {
        match value {
            Value::Object(map) if map.get("rawType").and_then(Value::as_str) == Some("SEQ") => {
                out.push((boxed, map["fallback"].as_str().unwrap().to_owned()))
            }
            Value::Object(map) => {
                let kind = map.get("kind").and_then(Value::as_str);
                let boxed = boxed || matches!(kind, Some("shape" | "textBox"));
                map.values().for_each(|value| collect(value, boxed, out))
            }
            Value::Array(items) => items.iter().for_each(|value| collect(value, boxed, out)),
            _ => {}
        }
    }
    let doc = EditingDoc::new(1);
    seed_from_docx(&doc, &document(body)).unwrap();
    let blocks = yrs_doc_to_layout_blocks(&doc, "body", &RenderEnv::default()).unwrap();
    let mut results = Vec::new();
    collect(&serde_json::to_value(blocks).unwrap(), false, &mut results);
    results
}

fn sequence_results(body: &str) -> Vec<String> {
    boxed_sequence_results(body)
        .into_iter()
        .map(|(_, result)| result)
        .collect()
}

/// Each instruction with the result Word 16 shows after updating fields
/// (checked against its PDF export), or "7", the cached result, where
/// BetterOffice keeps it.
const CASES: &[(&str, &str)] = &[
    (r"SEQ Figure \* ARABIC", "1"),
    (r"SEQ Figure", "2"),
    (r"SEQ figure", "3"),
    (r"SEQ Figure \c", "3"),
    (r"SEQ Figure \r 10", "10"),
    (r"SEQ Figure \h", ""),
    (r"SEQ Figure", "12"),
    (r"SEQ Figure \h \* ARABIC", "13"),
    (r"SEQ Figure \h \* MERGEFORMAT", ""),
    (r"SEQ Figure \c \r 20", "14"),
    (r"SEQ Figure \r 20 \c", "20"),
    (r"SEQ Figure \* Arabic \* roman", "xxi"),
    (r#"SEQ "Figure""#, "22"),
    (r"SEQ Table \* ROMAN", "I"),
    (r"SEQ Table \* roman", "ii"),
    (r"SEQ Table \* ALPHABETIC", "C"),
    (r"SEQ Table \* alphabetic", "d"),
    (r"SEQ Table \* Roman", "V"),
    (r"SEQ Other \c", "0"),
    (r"SEQ Letters \r 27 \* ALPHABETIC", "AA"),
    (r"SEQ Letters \r 53 \* alphabetic", "aaa"),
    (r"SEQ Numerals \r 4000 \* ROMAN", "MMMM"),
    (r"SEQ Zero \r 0 \* ROMAN", ""),
    (r#"SEQ Figure \# "00""#, "7"),
    (r"SEQ Figure \* CardText", "7"),
];

#[test]
fn seq_fields_number_like_word() {
    let mut body: String = CASES
        .iter()
        .map(|(instruction, _)| paragraph(&field(instruction)))
        .collect();
    let mut expected: Vec<&str> = CASES.iter().map(|(_, result)| *result).collect();
    // A locked field counts and keeps its cached result.
    body += &paragraph(&field("SEQ Figure").replacen(
        r#"w:fldCharType="begin""#,
        r#"w:fldCharType="begin" w:fldLock="true""#,
        1,
    ));
    expected.push("7");
    body += &format!(
        r#"<w:tbl><w:tblPr><w:tblW w:w="9000" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="9000"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="9000" w:type="dxa"/></w:tcPr>{}</w:tc></w:tr></w:tbl>"#,
        paragraph(&field("SEQ Figure"))
    );
    expected.push("26");
    body +=
        &paragraph(r#"<w:fldSimple w:instr=" SEQ Figure "><w:r><w:t>7</w:t></w:r></w:fldSimple>"#);
    expected.push("27");
    // Word numbers a chapter-reset sequence, and shows an error for a SEQ
    // without a name; both keep their cached results here.
    body += &paragraph(&field(r"SEQ Chapter \s 1"));
    body += &paragraph(&field("SEQ Chapter"));
    body += &paragraph(&field("SEQ"));
    expected.extend(["7", "7", "7"]);

    assert_eq!(sequence_results(&body), expected);
}

#[test]
fn a_sequence_with_a_nested_field_keeps_its_cached_results() {
    let nested = format!(
        r#"<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> QUOTE "</w:instrText></w:r>{}<w:r><w:instrText xml:space="preserve">" </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>7</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>"#,
        field("SEQ figure")
    );
    let body = [
        paragraph(&field("SEQ Figure")),
        paragraph(&nested),
        paragraph(&field("SEQ Figure")),
        paragraph(&field("SEQ Table")),
    ]
    .concat();
    assert_eq!(sequence_results(&body), ["7", "7", "1"]);
}

#[test]
fn a_text_box_with_a_projected_nested_sequence_keeps_cached_results() {
    let nested = format!(
        r#"<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> QUOTE "</w:instrText></w:r>{}<w:r><w:instrText xml:space="preserve">" </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:hyperlink w:anchor="top">{}</w:hyperlink><w:r><w:fldChar w:fldCharType="end"/></w:r>"#,
        field("SEQ Figure").replacen("<w:t>7</w:t>", "<w:t>1</w:t>", 1),
        run("1")
    );
    let body = [
        format!("<w:p>{}</w:p>", text_box(&nested)),
        paragraph(&field("SEQ Figure").replacen("<w:t>7</w:t>", "<w:t>2</w:t>", 1)),
        paragraph(&field("SEQ Table")),
    ]
    .concat();
    assert_eq!(
        boxed_sequence_results(&body),
        [(false, "2"), (false, "1")].map(|(boxed, result)| (boxed, result.to_owned()))
    );
}

#[test]
fn a_text_box_with_a_projected_sequence_keeps_cached_results() {
    let boxed = format!(
        r#"<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> SEQ Figure </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:hyperlink w:anchor="top">{}</w:hyperlink><w:r><w:fldChar w:fldCharType="end"/></w:r>"#,
        run("1")
    );
    let body = [
        format!("<w:p>{}</w:p>", text_box(&boxed)),
        paragraph(&field("SEQ Figure").replacen("<w:t>7</w:t>", "<w:t>2</w:t>", 1)),
        paragraph(&field("SEQ Table")),
    ]
    .concat();
    assert_eq!(
        boxed_sequence_results(&body),
        [(false, "2"), (false, "1")].map(|(boxed, result)| (boxed, result.to_owned()))
    );
}

#[test]
fn a_text_box_anchored_at_a_paragraph_start_counts_before_the_paragraph() {
    let body = [
        paragraph(&field("SEQ Figure")),
        format!(
            "<w:p>{}{}{}</w:p>",
            text_box(&(run("Boxed ") + &field("SEQ Figure"))),
            run("Caption "),
            field("SEQ Figure")
        ),
        paragraph(&field("SEQ Figure")),
    ]
    .concat();
    assert_eq!(
        boxed_sequence_results(&body),
        [(false, "1"), (true, "2"), (false, "3"), (false, "4")]
            .map(|(boxed, result)| (boxed, result.to_owned()))
    );
}
