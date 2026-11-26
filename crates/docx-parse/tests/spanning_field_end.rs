use docx_parse::serializer::{S13SaveRequest, write_docx_s13};
use docx_parse::{S9ParseOptions, parse_docx_s5_projection, parse_docx_s9_wire};
use serde_json::{Value, json};
use sha2::{Digest, Sha256};
use std::error::Error;

const W: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
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
const DOCUMENT_PART: &str = "word/document.xml";
const BEGIN: &str = r#"w:fldCharType="begin""#;
const END: &str = r#"w:fldCharType="end""#;
const BODY_NEXT: &str = "BODYNEXT";
const INSERTED: &str = "INSERTED";
const ENTRY_TWO: &str = "ENTRYTWO";
const PAGE_MARKERS: [&str; 3] = ["PAGEONE", "PAGETWO", "PAGETHREE"];

#[derive(Clone, Copy)]
enum Edit {
    Text(&'static str),
    Paragraph(ParagraphEdit),
    None,
}

#[derive(Clone, Copy)]
enum ParagraphEdit {
    InsertAfter(usize),
    Delete(usize),
}

const S1_EDITS: &[(&str, Edit)] = &[
    ("last", Edit::Text("SPANEND")),
    ("first", Edit::Text("SPANSTART")),
    ("after", Edit::Text(BODY_NEXT)),
    ("none", Edit::None),
    ("insert_inside", Edit::Paragraph(ParagraphEdit::InsertAfter(0))),
];
const TOC_EDITS: &[(&str, Edit)] = &[
    ("last", Edit::Text("ENTRYTHREE")),
    ("middle", Edit::Text(ENTRY_TWO)),
    ("first", Edit::Text("ENTRYONE")),
    ("after", Edit::Text(BODY_NEXT)),
    ("none", Edit::None),
    ("insert_inside", Edit::Paragraph(ParagraphEdit::InsertAfter(1))),
    ("delete_middle", Edit::Paragraph(ParagraphEdit::Delete(1))),
];

fn field_char(kind: &str) -> String {
    format!(r#"<w:r><w:fldChar w:fldCharType="{kind}"/></w:r>"#)
}

fn field_start(code: &str) -> String {
    let begin = field_char("begin");
    let separate = field_char("separate");
    format!(
        r#"{begin}<w:r><w:instrText xml:space="preserve"> {code} </w:instrText></w:r>{separate}"#
    )
}

fn text_run(text: &str) -> String {
    format!("<w:r><w:t>{text}</w:t></w:r>")
}

fn entry(number: usize, marker: &str, page_marker: &str) -> String {
    let label = text_run(marker);
    let pageref = field_start(&format!(r"PAGEREF _Toc{number} \h"));
    let page = text_run(page_marker);
    let end = field_char("end");
    format!(
        r#"<w:hyperlink w:anchor="_Toc{number}" w:history="1">{label}<w:r><w:tab/></w:r>{pageref}{page}{end}</w:hyperlink>"#
    )
}

fn body(scenario: &str) -> String {
    let end = field_char("end");
    let mut paragraphs = if scenario == "S1" {
        vec![
            format!("{}{}", field_start("QUOTE"), text_run("SPANSTART")),
            format!("{}{end}{}", text_run("SPANEND"), text_run("AFTERSPAN")),
        ]
    } else {
        let toc = field_start(r#"TOC \o "1-3" \h \z \u"#);
        let first = format!("{toc}{}", entry(1, "ENTRYONE", PAGE_MARKERS[0]));
        let second = entry(2, ENTRY_TWO, PAGE_MARKERS[1]);
        let third = entry(3, "ENTRYTHREE", PAGE_MARKERS[2]);
        if scenario == "TB" {
            vec![first, second, third, end]
        } else {
            vec![first, second, format!("{third}{end}")]
        }
    };
    paragraphs.push(text_run(BODY_NEXT));
    paragraphs
        .into_iter()
        .map(|content| format!("<w:p>{content}</w:p>"))
        .collect()
}

fn package(body: &str) -> Result<Vec<u8>, Box<dyn Error>> {
    let document =
        format!(r#"<w:document xmlns:w="{W}"><w:body>{body}<w:sectPr/></w:body></w:document>"#);
    Ok(ooxml_opc::rezip_parts(&[
        (
            "[Content_Types].xml".to_owned(),
            CONTENT_TYPES.as_bytes().to_vec(),
        ),
        ("_rels/.rels".to_owned(), ROOT_RELS.as_bytes().to_vec()),
        (DOCUMENT_PART.to_owned(), document.into_bytes()),
    ])?)
}

fn document_xml(bytes: &[u8]) -> Result<String, Box<dyn Error>> {
    let (_, xml) = ooxml_opc::unzip_parts(bytes)?
        .into_iter()
        .find(|(path, _)| path == DOCUMENT_PART)
        .ok_or("missing word/document.xml")?;
    Ok(String::from_utf8(xml)?)
}

fn edit_runs(value: &mut Value, marker: &str, edited: &str) -> usize {
    match value {
        Value::Object(node) if node.get("type").and_then(Value::as_str) == Some("run") => {
            let Some(content) = node.get_mut("content").and_then(Value::as_array_mut) else {
                return 0;
            };
            let mut changed = 0;
            for child in content {
                if child["type"] == "text" && child["text"] == marker {
                    child["text"] = json!(edited);
                    changed += 1;
                }
            }
            changed
        }
        Value::Object(node) => node
            .values_mut()
            .map(|child| edit_runs(child, marker, edited))
            .sum(),
        Value::Array(items) => items
            .iter_mut()
            .map(|child| edit_runs(child, marker, edited))
            .sum(),
        _ => 0,
    }
}

fn splice_paragraphs(
    paragraphs: &mut Vec<Value>,
    edit: ParagraphEdit,
    digest: &str,
) -> Result<Value, Box<dyn Error>> {
    let (ParagraphEdit::InsertAfter(index) | ParagraphEdit::Delete(index)) = edit;
    if index >= paragraphs.len() {
        return Err("missing paragraph for structural edit".into());
    }
    for (ordinal, paragraph) in paragraphs.iter_mut().enumerate() {
        paragraph["sourceOrdinal"] = json!(ordinal);
    }
    let (inserted, removed) = match edit {
        ParagraphEdit::InsertAfter(index) => {
            let mut paragraph = paragraphs.last().ok_or("missing plain paragraph")?.clone();
            paragraph
                .as_object_mut()
                .ok_or("invalid plain paragraph")?
                .remove("sourceOrdinal");
            if edit_runs(&mut paragraph, BODY_NEXT, INSERTED) == 0 {
                return Err("missing plain paragraph text".into());
            }
            paragraphs.insert(index + 1, paragraph);
            (vec![json!({"after": index})], Vec::new())
        }
        ParagraphEdit::Delete(index) => {
            paragraphs.remove(index);
            (Vec::new(), vec![index])
        }
    };
    let ordinals: Vec<_> = paragraphs
        .iter()
        .filter_map(|paragraph| paragraph["sourceOrdinal"].as_u64())
        .collect();
    Ok(json!({"splicedParts": [{
        "part": DOCUMENT_PART,
        "sha256": digest,
        "paragraphs": ordinals,
        "inserted": inserted,
        "removed": removed,
    }]}))
}

fn save(source: &[u8], edit: Edit) -> Result<Vec<u8>, Box<dyn Error>> {
    let parsed = parse_docx_s9_wire(source, S9ParseOptions::default())?;
    let mut package = serde_json::to_value(parsed.document.package)?;
    package["document"]
        .as_object_mut()
        .ok_or("missing document model")?
        .remove("sections");
    let mut request = json!({
        "determinism": {"seed": "0".repeat(64), "now": "2000-01-01T00:00:00.000Z"},
        "document": package["document"],
        "options": {"updateModifiedDate": false}
    });
    let digest = format!("{:x}", Sha256::digest(document_xml(source)?.as_bytes()));
    let paragraphs = request["document"]["content"]
        .as_array_mut()
        .ok_or("missing body blocks")?;
    if let Edit::Paragraph(edit) = edit {
        request["paragraphIds"] = splice_paragraphs(paragraphs, edit, &digest)?;
    } else {
        let mut targets = Vec::new();
        if let Edit::Text(marker) = edit {
            let edited = format!("{marker} edited");
            let block = paragraphs
                .iter_mut()
                .enumerate()
                .find_map(|(index, value)| (edit_runs(value, marker, &edited) > 0).then_some(index))
                .ok_or("missing edit marker")?;
            targets.push(json!({"path": [0, block], "block": block}));
        }
        request["selective"] = json!({
            "sourceParagraphs": {"partSha256": digest, "paragraphs": targets}
        });
    }
    let request: S13SaveRequest = serde_json::from_value(request)?;
    Ok(write_docx_s13(request, source)?)
}

struct FieldSpan {
    begin: usize,
    end: Option<usize>,
    outer: bool,
}

fn field_diagnostics(xml: &str) -> (Option<usize>, String, String) {
    let mut fields = Vec::new();
    let mut open = Vec::new();
    let mut instrs = Vec::new();
    let mut end_para = None;
    for (index, paragraph) in xml.split("</w:p>").enumerate() {
        for tag in paragraph.split("<w:").skip(1) {
            let Some((opening, content)) = tag.split_once('>') else {
                continue;
            };
            let name = opening
                .split(|ch: char| ch.is_ascii_whitespace() || ch == '/')
                .next();
            match name {
                Some("fldChar") if opening.contains(BEGIN) => {
                    open.push(fields.len());
                    fields.push(FieldSpan {
                        begin: index,
                        end: None,
                        outer: false,
                    });
                }
                Some("fldChar") if opening.contains(END) => {
                    if let Some(field) = open.pop() {
                        fields[field].end = Some(index);
                        if fields[field].outer {
                            end_para.get_or_insert(index);
                        }
                    }
                }
                Some("instrText") => {
                    let raw = if opening.ends_with('/') {
                        ""
                    } else {
                        content.split('<').next().unwrap_or_default()
                    };
                    let code = quick_xml::escape::unescape(raw).unwrap_or_else(|_| raw.into());
                    if matches!(code.split_whitespace().next(), Some("QUOTE" | "TOC"))
                        && let Some(&field) = open.last()
                    {
                        fields[field].outer = true;
                    }
                    instrs.push(code.trim().to_owned());
                }
                _ => {}
            }
        }
    }
    let fields = fields
        .iter()
        .map(|field| {
            let end = field
                .end
                .map_or_else(|| "?".to_owned(), |index| index.to_string());
            format!("b{}-e{end}", field.begin)
        })
        .collect::<Vec<_>>()
        .join(",");
    (end_para, instrs.join(";"), fields)
}

fn error_message(error: &dyn Error) -> String {
    let message = error.to_string();
    let mut message = message.split('<').next().unwrap_or_default();
    if let Some((prefix, _)) = message.split_once("&lt;") {
        message = prefix;
    }
    let mut message = message.split_whitespace().collect::<Vec<_>>().join(" ");
    for marker in [
        BODY_NEXT,
        INSERTED,
        "SPANSTART",
        "SPANEND",
        "AFTERSPAN",
        "ENTRYONE",
        ENTRY_TWO,
        "ENTRYTHREE",
    ]
    .into_iter()
    .chain(PAGE_MARKERS)
    {
        message = message
            .replace(&format!("{marker} edited"), "[redacted]")
            .replace(marker, "[redacted]");
    }
    message.chars().take(120).collect()
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

fn has_marker(value: &Value, marker: &str) -> bool {
    match value {
        Value::Object(node) => {
            (value["type"] == "text"
                && value["text"]
                    .as_str()
                    .is_some_and(|text| text.contains(marker)))
                || node.values().any(|child| has_marker(child, marker))
        }
        Value::Array(items) => items.iter().any(|child| has_marker(child, marker)),
        _ => false,
    }
}

fn field_contains_body_next(value: &Value, following: &[Value]) -> bool {
    match value {
        Value::Object(node) => {
            if value["type"] == "complexField" {
                let result = &value["structuredResult"];
                let code_blocks = value["structuredCode"]["blockCount"].as_u64().unwrap_or(0);
                let result_blocks = result["blockCount"].as_u64().unwrap_or(0);
                if has_marker(result, BODY_NEXT)
                    || has_marker(&value["fieldResult"], BODY_NEXT)
                    || following
                        .iter()
                        .skip(code_blocks as usize)
                        .take(result_blocks as usize)
                        .any(|block| has_marker(block, BODY_NEXT))
                {
                    return true;
                }
            }
            node.values()
                .any(|child| field_contains_body_next(child, following))
        }
        Value::Array(items) => items
            .iter()
            .any(|child| field_contains_body_next(child, following)),
        _ => false,
    }
}

fn nested_flags(present: impl Fn(&str) -> bool) -> String {
    PAGE_MARKERS
        .iter()
        .map(|marker| if present(marker) { 'y' } else { 'n' })
        .collect()
}

fn reopen(bytes: &[u8]) -> Result<(&'static str, String), Box<dyn Error>> {
    let projection = parse_docx_s5_projection(bytes)?;
    projection
        .xml_parts
        .iter()
        .find(|part| part.path == DOCUMENT_PART)
        .ok_or("missing main part in inline projection")?;
    let parsed = parse_docx_s9_wire(bytes, S9ParseOptions::default())?;
    let blocks = parsed
        .document
        .package
        .document
        .content
        .iter()
        .map(serde_json::to_value)
        .collect::<Result<Vec<_>, _>>()?;
    let swallowed = blocks
        .iter()
        .enumerate()
        .any(|(index, block)| field_contains_body_next(block, &blocks[index + 1..]));
    let nested = nested_flags(|marker| blocks.iter().any(|block| has_marker(block, marker)));
    Ok((if swallowed { "swallowed" } else { "ok" }, nested))
}

#[derive(Debug)]
struct Measurement {
    begins: usize,
    ends: usize,
    end_para: Option<usize>,
    text_ok: bool,
    reopen: &'static str,
    nested: String,
    nested_model: String,
    instrs: String,
    fields: String,
    error: Option<String>,
}

fn edits(scenario: &str) -> &'static [(&'static str, Edit)] {
    if scenario == "S1" { S1_EDITS } else { TOC_EDITS }
}

fn measure(scenario: &str, edit: Edit) -> Result<Measurement, Box<dyn Error>> {
    let source = package(&body(scenario))?;
    let saved = save(&source, edit)?;
    let xml = document_xml(&saved)?;
    let text = visible(&xml);
    let nested = nested_flags(|target| text.contains(target));
    let deleted_middle = matches!(edit, Edit::Paragraph(ParagraphEdit::Delete(_)));
    let inserted_inside = matches!(edit, Edit::Paragraph(ParagraphEdit::InsertAfter(_)));
    let expected_nested = if deleted_middle { "yny" } else { "yyy" };
    let text_ok = edits(scenario)
        .iter()
        .filter_map(|&(_, edit)| match edit {
            Edit::Text(target) => Some(target),
            _ => None,
        })
        .filter(|&target| !deleted_middle || target != ENTRY_TWO)
        .chain((scenario == "S1").then_some("AFTERSPAN"))
        .chain(inserted_inside.then_some(INSERTED))
        .all(|target| {
            if matches!(edit, Edit::Text(marker) if marker == target) {
                text.contains(&format!("{target} edited"))
            } else {
                text.contains(target)
            }
        }) && (scenario == "S1" || nested == expected_nested)
        && (!deleted_middle || !text.contains(ENTRY_TWO));
    let (end_para, instrs, fields) = field_diagnostics(&xml);
    let (reopen, nested_model, error) = match reopen(&saved) {
        Ok((status, nested)) => (status, nested, None),
        Err(error) => ("err", "nnn".to_owned(), Some(error_message(error.as_ref()))),
    };
    Ok(Measurement {
        begins: xml.matches(BEGIN).count(),
        ends: xml.matches(END).count(),
        end_para,
        text_ok,
        reopen,
        nested,
        nested_model,
        instrs,
        fields,
        error,
    })
}

#[test]
#[ignore]
fn spanning_field_end_probe() {
    for scenario in ["S1", "TA", "TB"] {
        for &(name, edit) in edits(scenario) {
            let measured = measure(scenario, edit).unwrap_or_else(|error| Measurement {
                begins: 0,
                ends: 0,
                end_para: None,
                text_ok: false,
                reopen: "err",
                nested: "nnn".to_owned(),
                nested_model: "nnn".to_owned(),
                instrs: String::new(),
                fields: String::new(),
                error: Some(error_message(error.as_ref())),
            });
            let end_para = measured
                .end_para
                .map_or_else(|| "none".to_owned(), |index| index.to_string());
            let text_ok = if measured.text_ok { "yes" } else { "no" };
            let error = measured
                .error
                .as_deref()
                .map_or_else(String::new, |message| format!(" | error={message}"));
            println!(
                "SPAN facade {scenario} {name} | begins={} ends={} | end_para={end_para} | \
                 text_ok={text_ok} | reopen={} | nested={} | nested_model={} | route_note=n/a | \
                 instrs={} | fields={}{error}",
                measured.begins,
                measured.ends,
                measured.reopen,
                measured.nested,
                measured.nested_model,
                measured.instrs,
                measured.fields
            );
        }
    }
}

#[test]
fn a_spanning_field_keeps_its_end_when_its_closing_paragraph_is_edited() {
    let measured = [
        ("S1", measure("S1", Edit::Text("SPANEND"))),
        ("TA", measure("TA", Edit::Text("ENTRYTHREE"))),
    ];
    let passes = measured.iter().all(|(_, result)| {
        result
            .as_ref()
            .is_ok_and(|result| result.begins == result.ends && result.reopen == "ok")
    });
    assert!(passes, "{measured:#?}");
}
