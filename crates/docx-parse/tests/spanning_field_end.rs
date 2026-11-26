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
    TextOnBothSides,
    Paragraph(ParagraphEdit),
    None,
}

#[derive(Clone, Copy, Debug)]
enum WriteRoute {
    SourceParagraphs,
    ChangedParaIds,
    Splice,
    Whole,
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
        .enumerate()
        .map(|(index, content)| paragraph(index, &content))
        .collect()
}

fn paragraph(index: usize, content: &str) -> String {
    format!(r#"<w:p w14:paraId="{:08X}">{content}</w:p>"#, index + 1)
}

fn package(body: &str) -> Result<Vec<u8>, Box<dyn Error>> {
    let document = format!(
        r#"<w:document xmlns:w="{W}" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:v="urn:schemas-microsoft-com:vml"><w:body>{body}<w:sectPr/></w:body></w:document>"#
    );
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
    addressed: bool,
) -> Result<Value, Box<dyn Error>> {
    let (ParagraphEdit::InsertAfter(index) | ParagraphEdit::Delete(index)) = edit;
    if index >= paragraphs.len() {
        return Err("missing paragraph for structural edit".into());
    }
    if addressed {
        for (ordinal, paragraph) in paragraphs.iter_mut().enumerate() {
            paragraph["sourceOrdinal"] = json!(ordinal);
        }
    }
    let (inserted, removed) = match edit {
        ParagraphEdit::InsertAfter(index) => {
            let mut paragraph = paragraphs.last().ok_or("missing plain paragraph")?.clone();
            paragraph
                .as_object_mut()
                .ok_or("invalid plain paragraph")?
                .remove("sourceOrdinal");
            paragraph.as_object_mut().ok_or("invalid plain paragraph")?.remove("paraId");
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

fn save(source: &[u8], edit: Edit, route: WriteRoute) -> Result<Vec<u8>, Box<dyn Error>> {
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
        let plan = splice_paragraphs(paragraphs, edit, &digest, matches!(route, WriteRoute::Splice))?;
        if matches!(route, WriteRoute::Splice) {
            request["paragraphIds"] = plan;
        }
    } else {
        let mut targets = Vec::new();
        let mut changed_ids = Vec::new();
        let markers = match edit {
            Edit::Text(ref marker) => std::slice::from_ref(marker),
            Edit::TextOnBothSides => &["SPANEND", "AFTERSPAN"],
            _ => &[],
        };
        for &marker in markers {
            let edited = format!("{marker} edited");
            let block = paragraphs
                .iter_mut()
                .enumerate()
                .find_map(|(index, value)| (edit_runs(value, marker, &edited) > 0).then_some(index))
                .ok_or("missing edit marker")?;
            let target = json!({"path": [0, block], "block": block});
            if !targets.contains(&target) {
                targets.push(target);
                changed_ids.push(paragraphs[block]["paraId"].clone());
            }
        }
        match route {
            WriteRoute::SourceParagraphs => {
                request["selective"] = json!({
                    "sourceParagraphs": {"partSha256": digest, "paragraphs": targets}
                });
            }
            WriteRoute::ChangedParaIds => {
                request["selective"] = json!({"changedParaIds": changed_ids});
            }
            WriteRoute::Splice | WriteRoute::Whole => {}
        }
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

fn assert_save(scenario: &str, edit: Edit, route: WriteRoute) -> Result<(), Box<dyn Error>> {
    let source = package(&body(scenario))?;
    let mut expected = inspect(&document_xml(&source)?);
    let saved = save(&source, edit, route)?;
    let actual = inspect(&document_xml(&saved)?);

    match edit {
        Edit::Text(_) | Edit::TextOnBothSides => {
            let markers = match edit {
                Edit::Text(ref marker) => std::slice::from_ref(marker),
                _ => &["SPANEND", "AFTERSPAN"],
            };
            for &marker in markers {
                let edited = format!("{marker} edited");
                for text in &mut expected.paragraphs {
                    *text = text.replace(marker, &edited);
                }
                for field in &mut expected.fields {
                    field.result = field.result.replace(marker, &edited);
                }
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
    assert_matrix("S1", S1_EDITS);
}

#[test]
fn toc_ending_in_the_last_entry_saves() {
    assert_matrix("TA", TOC_EDITS);
}

#[test]
fn toc_ending_in_its_own_paragraph_saves() {
    assert_matrix("TB", TOC_EDITS);
}

#[test]
fn spanning_results_keep_links_without_page_numbers() -> Result<(), Box<dyn Error>> {
    let label = text_run("ENTRYONE");
    let control = format!(
        "<w:sdt><w:sdtPr/><w:sdtContent>{}</w:sdtContent></w:sdt>",
        text_run("Control")
    );
    let math = r#"<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"><m:r><m:t>x</m:t></m:r></m:oMath>"#;
    let simple = r#"<w:fldSimple w:instr="AUTHOR"><w:r><w:t>Author</w:t></w:r></w:fldSimple>"#;
    for (code, suffix) in [
        ("QUOTE", String::new()),
        (r"TOC \h", String::new()),
        (r"TOC \h", format!("{control}{math}{simple}")),
    ] {
        let entry = format!(r#"<w:hyperlink w:anchor="_Toc1">{label}{suffix}</w:hyperlink>"#);
        let source = package(&[
            paragraph(0, &(field_start(code) + &entry)),
            paragraph(1, &field_char("end")),
        ].concat())?;
        for route in [
            WriteRoute::Whole,
            WriteRoute::SourceParagraphs,
            WriteRoute::ChangedParaIds,
        ] {
            let saved = save(&source, Edit::Text("ENTRYONE"), route)?;
            let xml = document_xml(&saved)?;
            assert_eq!(inspect(&xml).fields[0].end, Some(1), "{xml}");
            let link = xml
                .split_once(r#"<w:hyperlink w:anchor="_Toc1">"#)
                .unwrap()
                .1;
            let link = link.split_once("</w:hyperlink>").unwrap().0;
            assert!(link.contains(">ENTRYONE edited<"), "{link}");
            if !suffix.is_empty() {
                for element in ["<w:sdt>", "<m:oMath", "<w:fldSimple"] {
                    assert!(link.contains(element), "{link}");
                }
            }
        }
    }
    Ok(())
}

fn assert_matrix(scenario: &str, edits: &[(&str, Edit)]) {
    for &(name, edit) in edits {
        let routes: &[_] = if matches!(edit, Edit::Paragraph(_)) {
            &[WriteRoute::Splice]
        } else {
            &[WriteRoute::SourceParagraphs, WriteRoute::ChangedParaIds]
        };
        for &route in routes {
            assert_save(scenario, edit, route)
                .unwrap_or_else(|error| panic!("{scenario} {name} {route:?}: {error}"));
        }
    }
}

#[test]
fn a_spanning_field_keeps_its_end_when_its_closing_paragraph_is_edited() {
    assert_matrix("S1", &[("both sides of the end", Edit::TextOnBothSides)]);
    assert_matrix("TA", &[("closing entry", Edit::Text("ENTRYTHREE"))]);
}

#[test]
fn a_removed_toc_closing_paragraph_carries_the_end_to_the_last_surviving_entry() -> Result<(), Box<dyn Error>> {
    let end = field_char("end");
    let second = entry(2, ENTRY_TWO, PAGE_MARKERS[1]);
    let third = entry(3, "ENTRYTHREE", PAGE_MARKERS[2]);
    let inline_end = body("TA");
    let separate_end = body("TB");
    for (xml, removed, expected) in [
        (
            &inline_end,
            2,
            inline_end.replace(&paragraph(2, &format!("{third}{end}")), "")
                .replace(&second, &format!("{second}{end}")),
        ),
        (
            &separate_end,
            3,
            separate_end.replace(&paragraph(3, &end), "")
                .replace(&third, &format!("{third}{end}")),
        ),
    ] {
        for route in [WriteRoute::Splice, WriteRoute::Whole] {
            let source_xml = if matches!(route, WriteRoute::Whole) {
                (1..=5).fold(xml.to_string(), |xml, id| {
                    xml.replace(&format!(r#" w14:paraId="{id:08X}""#), "")
                })
            } else {
                xml.to_string()
            };
            let source = package(&source_xml)?;
            let saved = save(&source, Edit::Paragraph(ParagraphEdit::Delete(removed)), route)?;
            assert_eq!(inspect(&document_xml(&saved)?), inspect(&expected), "{route:?}");
        }
    }
    Ok(())
}

#[test]
fn marker_only_revisions_survive_on_both_selective_routes() -> Result<(), Box<dyn Error>> {
    for wrapper in ["ins", "del"] {
        let source = body("S1").replace(
            &field_char("end"),
            &format!(r#"<w:{wrapper} w:id="17" w:author="Reviewer" w:date="2000-01-01T00:00:00Z">{}</w:{wrapper}>"#, field_char("end")),
        );
        for route in [WriteRoute::SourceParagraphs, WriteRoute::ChangedParaIds] {
            let saved = save(&package(&source)?, Edit::Text("SPANEND"), route)?;
            let xml = document_xml(&saved)?;
            assert_eq!(inspect(&xml), inspect(&source.replace("SPANEND", "SPANEND edited")));
            let revision = xml.split_once(&format!(r#"<w:{wrapper} w:id="17" w:author="Reviewer" w:date="2000-01-01T00:00:00Z">"#))
                .and_then(|(_, tail)| tail.split_once(&format!("</w:{wrapper}>")))
                .map(|(revision, _)| revision);
            assert!(revision.is_some_and(|revision| revision.contains(END)), "{xml}");
        }
    }
    Ok(())
}

#[test]
fn consecutive_spans_keep_separate_runs_in_source_order() -> Result<(), Box<dyn Error>> {
    let end = field_char("end");
    let first = field_start("QUOTE FIRST") + &text_run("FIRST");
    let second = field_start("QUOTE SECOND") + &text_run("SECOND");
    let source = [
        paragraph(0, &first),
        paragraph(1, &format!("{}{end}{}{second}", text_run("FIRSTEND"), text_run("BETWEEN"))),
        paragraph(2, &format!("{}{end}{}", text_run("SECONDEND"), text_run("OUTSIDE"))),
    ].concat();
    for route in [WriteRoute::SourceParagraphs, WriteRoute::ChangedParaIds, WriteRoute::Whole] {
        let saved = save(&package(&source)?, Edit::Text("BETWEEN"), route)?;
        assert_eq!(inspect(&document_xml(&saved)?), inspect(&source.replace("BETWEEN", "BETWEEN edited")));
    }
    Ok(())
}

#[test]
fn shared_ends_keep_their_run_or_revision() -> Result<(), Box<dyn Error>> {
    let original = body("S1");
    let plain_end = field_char("end");
    let following = text_run("AFTERSPAN");
    let original_boundary = format!("{plain_end}{following}");
    for boundary in [
        r#"<w:r><w:fldChar w:fldCharType="end"/><w:t>AFTERSPAN</w:t></w:r>"#.to_owned(),
        format!(r#"<w:ins w:id="17" w:author="Reviewer" w:date="2000-01-01T00:00:00Z">{plain_end}{following}</w:ins>"#),
    ] {
        let xml = original.replace(&original_boundary, &boundary);
        let source = package(&xml)?;
        for route in [WriteRoute::Whole, WriteRoute::SourceParagraphs, WriteRoute::ChangedParaIds] {
            for marker in ["SPANSTART", "SPANEND"] {
                let saved = save(&source, Edit::Text(marker), route)?;
                let saved_xml = document_xml(&saved)?;
                assert_eq!(inspect(&saved_xml), inspect(&xml.replace(marker, &format!("{marker} edited"))), "{route:?}");
                assert!(saved_xml.contains(&boundary), "{saved_xml}");
            }
        }
    }
    Ok(())
}

#[test]
fn tracked_opener_keeps_the_main_save_behavior() -> Result<(), Box<dyn Error>> {
    let empty = r#"<w:ins w:id="17" w:author="Reviewer" w:date="2000-01-01T00:00:00Z">"#;
    let revision = format!("{empty}{}{}</w:ins>", field_start("QUOTE"), text_run("TRACKED"));
    let source = [
        paragraph(0, &revision),
        paragraph(1, &format!("{}{}", text_run("CLOSING"), field_char("end"))),
        paragraph(2, &text_run(BODY_NEXT)),
    ].concat();
    for route in [WriteRoute::Whole, WriteRoute::SourceParagraphs] {
        let saved = save(&package(&source)?, Edit::Text(BODY_NEXT), route)?;
        let xml = document_xml(&saved)?;
        assert!(xml.contains(">BODYNEXT edited<"), "{xml}");
        if matches!(route, WriteRoute::Whole) {
            assert!(xml.contains(&format!("{empty}</w:ins>")), "{xml}");
            assert!(!xml.contains("w:fldChar"), "{xml}");
        } else {
            assert!(xml.contains(&revision), "{xml}");
            assert_eq!(inspect(&xml), inspect(&source.replace(BODY_NEXT, "BODYNEXT edited")));
        }
    }
    Ok(())
}

#[test]
fn nested_spanning_opens_keep_the_legacy_full_write() -> Result<(), Box<dyn Error>> {
    let end = field_char("end");
    let outer = field_start("QUOTE OUTER") + &text_run("OUTER");
    let inner = field_start("QUOTE INNER") + &text_run("INNER");
    let source = [
        paragraph(0, &format!("{outer}{inner}")),
        paragraph(1, &format!("{}{end}", text_run("INNEREND"))),
        paragraph(2, &format!("{}{end}{}", text_run("OUTEREND"), text_run("AFTER"))),
    ].concat();
    let expected = [
        paragraph(0, &format!("{outer}{end}")),
        paragraph(1, &text_run("INNEREND")),
        paragraph(2, &format!("{}{}", text_run("OUTEREND"), text_run("AFTER edited"))),
    ].concat();
    let saved = save(&package(&source)?, Edit::Text("AFTER"), WriteRoute::Whole)?;
    assert_eq!(inspect(&document_xml(&saved)?), inspect(&expected));
    Ok(())
}

#[test]
fn cross_container_closes_keep_the_legacy_opening_paragraph() -> Result<(), Box<dyn Error>> {
    let end = field_char("end");
    let opener = field_start("QUOTE OUTER") + &text_run("OUTER");
    let nested = field_start("QUOTE LOCAL") + &text_run("LOCAL") + &end;
    let source_opening = paragraph(0, &format!("{opener}{nested}"));
    let box_paragraph = paragraph(2, &end);
    for container in [
        format!("<w:tbl><w:tr><w:tc>{}</w:tc></w:tr></w:tbl>", paragraph(2, &end)),
        paragraph(1, &format!("<w:r><w:pict><v:shape><v:textbox><w:txbxContent>{box_paragraph}</w:txbxContent></v:textbox></v:shape></w:pict></w:r>")),
    ] {
        let source = package(&format!("{source_opening}{container}"))?;
        let saved = save(&source, Edit::Text("OUTER"), WriteRoute::SourceParagraphs)?;
        let xml = document_xml(&saved)?;
        let opening = &xml[..xml.find("</w:p>").unwrap() + "</w:p>".len()];
        assert_eq!(inspect(opening), inspect(&paragraph(0, &format!("{}{end}", opener.replace(">OUTER<", ">OUTER edited<")))));
        assert!(xml.contains(&container), "{xml}");
    }
    Ok(())
}

#[test]
fn selective_writes_without_a_span_keep_the_other_paragraph_verbatim() -> Result<(), Box<dyn Error>> {
    let target = paragraph(0, &text_run("TARGET"));
    let kept = r#"<w:p w14:paraId="00000002" w:rsidR="0000000A"><w:r><w:t><![CDATA[Kept]]></w:t></w:r></w:p>"#;
    let source = package(&format!("{target}{kept}"))?;
    for route in [WriteRoute::SourceParagraphs, WriteRoute::ChangedParaIds] {
        let saved = save(&source, Edit::Text("TARGET"), route)?;
        let xml = document_xml(&saved)?;
        assert!(xml.contains(">TARGET edited<"));
        assert!(xml.contains(kept));
    }
    Ok(())
}

#[test]
fn changed_ids_keep_untouched_spanning_paragraphs_verbatim() -> Result<(), Box<dyn Error>> {
    let source = body("S1").replace(">SPANEND<", "><![CDATA[SPANEND]]><");
    let source = package(&source)?;
    let saved = save(&source, Edit::Text(BODY_NEXT), WriteRoute::ChangedParaIds)?;
    assert_eq!(document_xml(&saved)?, document_xml(&source)?.replace(BODY_NEXT, "BODYNEXT edited"));
    Ok(())
}

#[test]
fn a_same_paragraph_raw_end_stays_on_its_original_side_of_the_next_field() -> Result<(), Box<dyn Error>> {
    let end = field_char("end");
    let preceding = format!(r#"<w:hyperlink w:anchor="top">{}</w:hyperlink>{end}"#, field_char("begin"));
    let next = field_start("QUOTE NEXT") + &text_run("NEXT") + &end;
    let source = paragraph(0, &format!("{preceding}{next}"));
    for route in [WriteRoute::SourceParagraphs, WriteRoute::ChangedParaIds] {
        let saved = save(&package(&source)?, Edit::Text("NEXT"), route)?;
        assert_eq!(inspect(&document_xml(&saved)?), inspect(&source.replace(">NEXT<", ">NEXT edited<")));
    }
    Ok(())
}
