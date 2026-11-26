use docx_parse::serializer::{S13SaveRequest, write_docx_s13};
use docx_parse::{BlockContent, InlineNode, ParagraphContent, S9ParseOptions, parse_docx_s9_wire};
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
const SEPARATE: &str = r#"w:fldCharType="separate""#;
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

#[derive(Debug, PartialEq, Eq)]
struct FieldSpan {
    begin: usize,
    end: Option<usize>,
    separates: Vec<usize>,
    instructions: Vec<String>,
    result: String,
    parent: Option<usize>,
}

#[derive(Debug, PartialEq, Eq, Default)]
struct Story {
    paragraphs: Vec<String>,
    fields: Vec<FieldSpan>,
    instructions: Vec<String>,
    ends: usize,
}

fn inspect(xml: &str) -> Story {
    let mut story = Story::default();
    let mut open = Vec::<usize>::new();
    for (index, paragraph) in xml
        .split("</w:p>")
        .filter(|paragraph| paragraph.contains("<w:p>") || paragraph.contains("<w:p "))
        .enumerate()
    {
        let mut visible = String::new();
        for tag in paragraph.split("<w:").skip(1) {
            let Some((opening, content)) = tag.split_once('>') else {
                continue;
            };
            let name = opening
                .split(|ch: char| ch.is_ascii_whitespace() || ch == '/')
                .next();
            match name {
                Some("fldChar") if opening.contains(BEGIN) => {
                    story.fields.push(FieldSpan {
                        begin: index,
                        end: None,
                        separates: Vec::new(),
                        instructions: Vec::new(),
                        result: String::new(),
                        parent: open.last().copied(),
                    });
                    open.push(story.fields.len() - 1);
                }
                Some("fldChar") if opening.contains(SEPARATE) => {
                    let field = *open.last().expect("field separator without a begin");
                    story.fields[field].separates.push(index);
                }
                Some("fldChar") if opening.contains(END) => {
                    story.ends += 1;
                    let field = open.pop().expect("field end without a begin");
                    story.fields[field].end = Some(index);
                }
                Some("instrText" | "t") => {
                    let raw = if opening.ends_with('/') {
                        ""
                    } else {
                        content.split('<').next().unwrap_or_default()
                    };
                    let text = quick_xml::escape::unescape(raw).expect("valid XML text");
                    if name == Some("instrText") {
                        let code = text.into_owned();
                        let field = *open.last().expect("field instruction without a begin");
                        assert!(
                            story.fields[field].separates.is_empty(),
                            "instruction after separator"
                        );
                        story.fields[field].instructions.push(code.clone());
                        story.instructions.push(code);
                    } else {
                        visible.push_str(&text);
                        for &field in &open {
                            if !story.fields[field].separates.is_empty() {
                                story.fields[field].result.push_str(&text);
                            }
                        }
                    }
                }
                _ => {}
            }
        }
        story.paragraphs.push(visible);
    }
    assert!(open.is_empty(), "fields without ends: {open:?}");
    assert_eq!(story.fields.len(), story.ends);
    story
}

fn assert_save(scenario: &str, edit: Edit) -> Result<(), Box<dyn Error>> {
    let source = package(&body(scenario))?;
    let mut expected = inspect(&document_xml(&source)?);
    let saved = save(&source, edit)?;
    let actual = inspect(&document_xml(&saved)?);

    match edit {
        Edit::Text(marker) => {
            let edited = format!("{marker} edited");
            for text in &mut expected.paragraphs {
                *text = text.replace(marker, &edited);
            }
            for field in &mut expected.fields {
                field.result = field.result.replace(marker, &edited);
            }
        }
        Edit::Paragraph(ParagraphEdit::InsertAfter(index)) => {
            expected.paragraphs.insert(index + 1, INSERTED.to_owned());
            let offset = expected.paragraphs[..=index]
                .iter()
                .map(String::len)
                .sum::<usize>();
            expected.fields[0].result.insert_str(offset, INSERTED);
            shift_fields(&mut expected.fields, index, 1);
        }
        Edit::Paragraph(ParagraphEdit::Delete(index)) => {
            let removed = expected.paragraphs.remove(index);
            let field = expected.fields.remove(index + 1);
            expected.fields[0].result = expected.fields[0].result.replace(&removed, "");
            expected.instructions.retain(|code| !field.instructions.contains(code));
            expected.ends -= 1;
            shift_fields(&mut expected.fields, index, -1);
        }
        Edit::None => {}
    }

    assert_eq!(
        actual.paragraphs, expected.paragraphs,
        "{scenario} paragraph text"
    );
    assert_eq!(
        actual.instructions, expected.instructions,
        "{scenario} instructions"
    );
    assert_eq!(
        actual.fields, expected.fields,
        "{scenario} field markers and results"
    );
    let reopened = parse_docx_s9_wire(&saved, S9ParseOptions::default())?;
    let BlockContent::Paragraph(first) = &reopened.document.package.document.content[0] else {
        panic!("missing opening paragraph");
    };
    let field = first
        .content
        .iter()
        .find_map(|item| match item {
            ParagraphContent::Inline(InlineNode::ComplexField(field)) => Some(field),
            _ => None,
        })
        .expect("missing spanning field after reopen");
    let count = field.structured_result.as_ref().and_then(|result| result.block_count);
    assert_eq!(count, expected.fields[0].end);
    Ok(())
}

fn shift_fields(fields: &mut [FieldSpan], after: usize, by: isize) {
    for field in fields {
        for index in std::iter::once(&mut field.begin)
            .chain(field.end.iter_mut())
            .chain(&mut field.separates)
        {
            if *index > after {
                *index = index.checked_add_signed(by).expect("valid paragraph index");
            }
        }
    }
}

#[test]
fn single_spanning_field_saves() {
    for &(name, edit) in S1_EDITS {
        assert_save("S1", edit).unwrap_or_else(|error| panic!("S1 {name}: {error}"));
    }
}

#[test]
fn toc_ending_in_the_last_entry_saves() {
    for &(name, edit) in TOC_EDITS {
        assert_save("TA", edit).unwrap_or_else(|error| panic!("TA {name}: {error}"));
    }
}

#[test]
fn toc_ending_in_its_own_paragraph_saves() {
    for &(name, edit) in TOC_EDITS {
        assert_save("TB", edit).unwrap_or_else(|error| panic!("TB {name}: {error}"));
    }
}

#[test]
fn a_spanning_field_keeps_its_end_when_its_closing_paragraph_is_edited() {
    assert_save("S1", Edit::Text("SPANEND")).expect("closing QUOTE paragraph saves");
    assert_save("TA", Edit::Text("ENTRYTHREE")).expect("closing TOC entry saves");
}
