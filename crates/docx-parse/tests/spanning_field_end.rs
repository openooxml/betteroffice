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
    Replace(&'static str, &'static str),
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
    (
        "insert_inside",
        Edit::Paragraph(ParagraphEdit::InsertAfter(0)),
    ),
];
const TOC_EDITS: &[(&str, Edit)] = &[
    ("last", Edit::Text("ENTRYTHREE")),
    ("middle", Edit::Text(ENTRY_TWO)),
    ("first", Edit::Text("ENTRYONE")),
    ("after", Edit::Text(BODY_NEXT)),
    ("none", Edit::None),
    (
        "insert_inside",
        Edit::Paragraph(ParagraphEdit::InsertAfter(1)),
    ),
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
        let toc = field_start(if scenario == "QUOTE" {
            "QUOTE"
        } else {
            r#"TOC \o "1-3" \h \z \u"#
        });
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

fn block_sdt(content: &str) -> String {
    format!(
        concat!(
            r#"<w:sdt><w:sdtPr><w:id w:val="42"/><w:docPartObj>"#,
            r#"<w:docPartGallery w:val="Table of Contents"/><w:docPartUnique/>"#,
            "</w:docPartObj></w:sdtPr><w:sdtContent>{}</w:sdtContent></w:sdt>"
        ),
        content
    )
}

fn without_paragraph_ids(xml: &str) -> String {
    (1..=5).fold(xml.to_owned(), |xml, id| {
        xml.replace(&format!(r#" w14:paraId="{id:08X}""#), "")
    })
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
    insertion_donor: Option<Value>,
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
            let mut paragraph = insertion_donor
                .or_else(|| paragraphs.last().cloned())
                .ok_or("missing plain paragraph")?;
            paragraph
                .as_object_mut()
                .ok_or("invalid plain paragraph")?
                .remove("sourceOrdinal");
            paragraph
                .as_object_mut()
                .ok_or("invalid plain paragraph")?
                .remove("paraId");
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
    save_story(source, edit, route, "/document/content")
}

fn save_story(
    source: &[u8],
    edit: Edit,
    route: WriteRoute,
    story_path: &str,
) -> Result<Vec<u8>, Box<dyn Error>> {
    let mut request = save_request(source, false)?;
    let digest = format!("{:x}", Sha256::digest(document_xml(source)?.as_bytes()));
    let insertion_donor = if story_path != "/document/content"
        && matches!(edit, Edit::Paragraph(ParagraphEdit::InsertAfter(_)))
    {
        request
            .pointer("/document/content")
            .and_then(Value::as_array)
            .and_then(|blocks| blocks.last())
            .filter(|block| block["type"] == "paragraph")
            .cloned()
    } else {
        None
    };
    let paragraphs = request
        .pointer_mut(story_path)
        .and_then(Value::as_array_mut)
        .ok_or("missing story blocks")?;
    if let Edit::Paragraph(edit) = edit {
        let plan = splice_paragraphs(
            paragraphs,
            edit,
            &digest,
            matches!(route, WriteRoute::Splice),
            insertion_donor,
        )?;
        if matches!(route, WriteRoute::Splice) {
            request["paragraphIds"] = plan;
        }
    } else {
        let mut targets = Vec::new();
        let mut changed_ids = Vec::new();
        let mut changed = Vec::new();
        let markers = match edit {
            Edit::Text(ref marker) | Edit::Replace(ref marker, _) => std::slice::from_ref(marker),
            Edit::TextOnBothSides => &["SPANEND", "AFTERSPAN"],
            _ => &[],
        };
        for &marker in markers {
            let edited = match edit {
                Edit::Replace(_, text) => text.to_owned(),
                _ => format!("{marker} edited"),
            };
            let block = paragraphs
                .iter_mut()
                .enumerate()
                .find_map(|(index, value)| (edit_runs(value, marker, &edited) > 0).then_some(index))
                .ok_or("missing edit marker")?;
            let target = json!({"path": [0, block], "block": block});
            if !targets.contains(&target) {
                targets.push(target);
                changed_ids.push(paragraphs[block]["paraId"].clone());
                changed.push(block);
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
            WriteRoute::Splice => {
                for (ordinal, paragraph) in paragraphs.iter_mut().enumerate() {
                    paragraph["sourceOrdinal"] = json!(ordinal);
                }
                let ordinals: Vec<_> = (0..paragraphs.len()).collect();
                request["paragraphIds"] = json!({"splicedParts": [{
                    "part": DOCUMENT_PART,
                    "sha256": digest,
                    "paragraphs": ordinals,
                    "changed": changed,
                }]});
            }
            WriteRoute::Whole => {}
        }
    }
    let request: S13SaveRequest = serde_json::from_value(request)?;
    Ok(write_docx_s13(request, source)?)
}

fn save_request(source: &[u8], source_ordinals: bool) -> Result<Value, Box<dyn Error>> {
    let options = S9ParseOptions {
        source_ordinals,
        ..Default::default()
    };
    let parsed = parse_docx_s9_wire(source, options)?;
    let mut package = serde_json::to_value(parsed.document.package)?;
    package["document"]
        .as_object_mut()
        .ok_or("missing document model")?
        .remove("sections");
    Ok(json!({
        "determinism": {"seed": "0".repeat(64), "now": "2000-01-01T00:00:00.000Z"},
        "document": package["document"],
        "options": {"updateModifiedDate": false}
    }))
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
    inspect_story(xml, "story")
}

fn story_paragraphs(xml: &str) -> impl Iterator<Item = &str> {
    xml.split("</w:p>")
        .filter(|paragraph| paragraph.contains("<w:p>") || paragraph.contains("<w:p "))
}

fn inspect_story(xml: &str, case: &str) -> Story {
    let saved = scan_story(xml);
    assert_eq!(saved.unpaired, 0, "{case}: unpaired field characters");
    saved.story
}

#[derive(Default)]
struct SavedStory {
    story: Story,
    text: String,
    tabs: Vec<usize>,
    ends_at: Vec<Option<usize>>,
    markers: Vec<String>,
    marker_fields: Vec<Vec<usize>>,
    unpaired: usize,
}

fn attribute<'a>(tag: &'a str, name: &str) -> Option<&'a str> {
    tag.split_once(&format!(r#" {name}=""#))?
        .1
        .split_once('"')
        .map(|(value, _)| value)
}

fn scan_story(xml: &str) -> SavedStory {
    let mut saved = SavedStory::default();
    let mut open: Vec<usize> = Vec::new();
    let mut paragraphs: Vec<usize> = Vec::new();
    let mut bookmarks = std::collections::HashMap::new();
    let mut revisions = Vec::new();
    let mut offset = 0;
    let mut run_depth = 0;
    for tag in xml.split('<').skip(1) {
        let Some((opening, content)) = tag.split_once('>') else {
            continue;
        };
        let name = opening
            .trim_end_matches('/')
            .split_ascii_whitespace()
            .next()
            .unwrap_or_default();
        if name == "w:p" {
            let index = saved.story.paragraphs.len();
            saved.story.paragraphs.push(String::new());
            saved.tabs.push(0);
            if !opening.ends_with('/') {
                paragraphs.push(index);
            }
        } else if name == "/w:p" {
            paragraphs.pop();
        } else if name == "w:r" && !opening.ends_with('/') {
            run_depth += 1;
        } else if name == "/w:r" {
            run_depth -= 1;
        }

        let id = attribute(opening, "w:id").unwrap_or("-");
        let marker = match name {
            "w:bookmarkStart" => {
                let name = attribute(opening, "w:name").unwrap_or(id);
                bookmarks.insert(id, name);
                Some(("bs", name))
            }
            "w:bookmarkEnd" => Some(("be", bookmarks.get(id).copied().unwrap_or(id))),
            "w:commentRangeStart" => Some(("cs", id)),
            "w:commentRangeEnd" => Some(("ce", id)),
            "w:commentReference" => Some(("cr", id)),
            "w:moveFromRangeStart" | "w:moveFromRangeEnd"
            | "w:moveToRangeStart" | "w:moveToRangeEnd" => Some((&name[2..], id)),
            "w:ins" | "w:del" | "w:moveFrom" | "w:moveTo" => {
                if !opening.ends_with('/') {
                    revisions.push(id);
                }
                Some((&name[2..], id))
            }
            "/w:ins" => Some(("/ins", revisions.pop().unwrap_or("-"))),
            "/w:del" => Some(("/del", revisions.pop().unwrap_or("-"))),
            "/w:moveFrom" => Some(("/moveFrom", revisions.pop().unwrap_or("-"))),
            "/w:moveTo" => Some(("/moveTo", revisions.pop().unwrap_or("-"))),
            _ => None,
        };
        if let Some((kind, id)) = marker {
            saved.markers.push(format!("{kind}:{id}@{offset}"));
            saved.marker_fields.push(open.clone());
        }
        let Some(&index) = paragraphs.last() else {
            continue;
        };
        let visible = open.iter().all(|&field| !saved.story.fields[field].separates.is_empty());
        match name {
            "w:fldChar" if opening.contains(BEGIN) => {
                saved.story.fields.push(FieldSpan {
                    begin: index,
                    end: None,
                    separates: Vec::new(),
                    instructions: Vec::new(),
                    result: String::new(),
                    parent: open.last().copied(),
                });
                saved.ends_at.push(None);
                open.push(saved.story.fields.len() - 1);
            }
            "w:fldChar" if opening.contains(SEPARATE) => {
                if let Some(&field) = open.last() {
                    saved.story.fields[field].separates.push(index);
                } else {
                    saved.unpaired += 1;
                }
            }
            "w:fldChar" if opening.contains(END) => {
                saved.story.ends += 1;
                if let Some(field) = open.pop() {
                    saved.story.fields[field].end = Some(index);
                    saved.ends_at[field] = Some(offset);
                } else {
                    saved.unpaired += 1;
                }
            }
            "w:instrText" if !opening.ends_with('/') => {
                let code = quick_xml::escape::unescape(content).unwrap().into_owned();
                if let Some(&field) = open.last() {
                    saved.unpaired += usize::from(!saved.story.fields[field].separates.is_empty());
                    saved.story.fields[field].instructions.push(code.clone());
                } else {
                    saved.unpaired += 1;
                }
                saved.story.instructions.push(code);
            }
            "w:t" if visible && !opening.ends_with('/') => {
                let text = quick_xml::escape::unescape(content).unwrap();
                saved.story.paragraphs[index].push_str(&text);
                for &field in &open {
                    if !saved.story.fields[field].separates.is_empty() {
                        saved.story.fields[field].result.push_str(&text);
                    }
                }
                offset += text.encode_utf16().count();
                saved.text.push_str(&text);
            }
            "w:tab" if visible && run_depth > 0 => {
                saved.tabs[index] += 1;
                offset += 1;
                saved.text.push('\t');
            }
            _ => {}
        }
    }
    saved.unpaired += open.len();
    saved
}

fn closed_pagerefs(story: &Story) -> Vec<String> {
    story
        .fields
        .iter()
        .filter(|field| field.end.is_some())
        .map(|field| field.instructions.concat())
        .filter(|code| code.split_whitespace().next() == Some("PAGEREF"))
        .collect()
}

fn assert_spanning_save(xml: &str, expected_xml: &str, case: &str) {
    let actual = scan_story(xml);
    let expected = scan_story(expected_xml);
    let mut actual_refs = closed_pagerefs(&actual.story);
    let expected_refs = closed_pagerefs(&expected.story);
    let kept = expected_refs
        .iter()
        .filter(|code| {
            actual_refs
                .iter()
                .position(|actual| actual == *code)
                .map(|index| actual_refs.swap_remove(index))
                .is_some()
        })
        .count();
    let instruction = if actual.story.instructions == expected.story.instructions {
        "same"
    } else {
        "changed"
    };
    let end_para = actual
        .story
        .fields
        .first()
        .and_then(|field| field.end)
        .map_or_else(|| "none".to_owned(), |index| index.to_string());
    let after = expected.ends_at.first().copied().flatten().and_then(|end| {
        let units: Vec<_> = expected.text.encode_utf16().skip(end).collect();
        let tail = String::from_utf16_lossy(&units);
        actual
            .text
            .rfind(&tail)
            .map(|start| actual.text[..start].encode_utf16().count())
    });
    let outside = actual
        .ends_at
        .first()
        .copied()
        .flatten()
        .zip(after)
        .is_some_and(|(end, start)| end <= start);
    let markers = actual.markers.join(",");
    println!(
        "SPANDIFF rust {case} begins={} ends={} instr={instruction} pageref={kept}/{} \
         endpara={end_para} after={} markers={}",
        actual.story.fields.len(),
        actual.story.ends,
        expected_refs.len(),
        if outside { "outside" } else { "inside" },
        if markers.is_empty() { "-" } else { &markers },
    );
    assert_eq!(
        actual.unpaired, expected.unpaired,
        "{case}: unpaired field characters"
    );
    assert_eq!(actual.story, expected.story, "{case}: field boundaries and text");
    assert_eq!(actual.tabs, expected.tabs, "{case}: tabs per paragraph");
    assert_eq!(actual.ends_at, expected.ends_at, "{case}: field end offsets");
    assert_eq!(actual.markers, expected.markers, "{case}: markers");
    assert_eq!(
        actual.marker_fields, expected.marker_fields,
        "{case}: markers within fields"
    );
}

fn assert_spanning_cases(cases: &[(String, String, String)]) {
    let failed: Vec<_> = cases
        .iter()
        .filter_map(|(xml, expected, case)| {
            std::panic::catch_unwind(|| assert_spanning_save(xml, expected, case))
                .is_err()
                .then_some(case)
        })
        .collect();
    assert!(failed.is_empty(), "failed cases: {failed:?}");
}

fn assert_save(
    scenario: &str,
    edit: Edit,
    route: WriteRoute,
    in_sdt: bool,
) -> Result<(), Box<dyn Error>> {
    let body = body(scenario);
    let source = package(&if in_sdt { block_sdt(&body) } else { body })?;
    let mut expected = inspect(&document_xml(&source)?);
    let story_path = if in_sdt {
        "/document/content/0/content"
    } else {
        "/document/content"
    };
    let saved = save_story(&source, edit, route, story_path)?;
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
            expected
                .instructions
                .retain(|code| !field.instructions.contains(code));
            expected.ends -= 1;
            shift_fields(&mut expected.fields, index, -1);
        }
        Edit::None => {}
        Edit::Replace(..) => unreachable!(),
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
    let mut content = &reopened.document.package.document.content;
    if in_sdt {
        let BlockContent::BlockSdt(sdt) = &content[0] else {
            panic!("missing block SDT");
        };
        content = &sdt.content;
    }
    let BlockContent::Paragraph(first) = &content[0] else {
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
    let count = field
        .structured_result
        .as_ref()
        .and_then(|result| result.block_count);
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
fn toc_inside_a_block_sdt_saves() -> Result<(), Box<dyn Error>> {
    for edit in [
        Edit::Text("ENTRYONE"),
        Edit::Text("ENTRYTHREE"),
        Edit::Paragraph(ParagraphEdit::InsertAfter(1)),
        Edit::Paragraph(ParagraphEdit::Delete(1)),
    ] {
        for route in [WriteRoute::Splice, WriteRoute::Whole] {
            assert_save("TA", edit, route, true)?;
        }
        if matches!(edit, Edit::Text(_)) {
            assert_save("TA", edit, WriteRoute::ChangedParaIds, true)?;
        }
    }
    Ok(())
}

#[test]
fn spanning_quote_keeps_its_complete_nested_result_field() -> Result<(), Box<dyn Error>> {
    let end = field_char("end");
    let source = [
        paragraph(
            0,
            &(field_start("QUOTE OUTER")
                + &text_run("OUTER")
                + &field_start("QUOTE INNER")
                + &text_run("INNER")
                + &end),
        ),
        paragraph(1, &(text_run("CLOSING") + &end)),
    ]
    .concat();
    for route in [WriteRoute::SourceParagraphs, WriteRoute::Whole] {
        assert_save("QUOTE", Edit::Text("ENTRYONE"), route, false)?;
        let saved = save(&package(&source)?, Edit::Text("OUTER"), route)?;
        assert_eq!(
            inspect(&document_xml(&saved)?),
            inspect(&source.replace(">OUTER<", ">OUTER edited<"))
        );
    }
    Ok(())
}

#[test]
fn spanning_quote_keeps_bookmarks_around_a_projected_pageref() -> Result<(), Box<dyn Error>> {
    let start = r#"<w:bookmarkStart w:id="41" w:name="Entry"/>"#;
    let stop = r#"<w:bookmarkEnd w:id="41"/>"#;
    let entry = entry(1, "ENTRYONE", PAGE_MARKERS[0]);
    let source = paragraph(0, &format!("{}{start}{entry}{stop}", field_start("QUOTE")))
        + &paragraph(1, &field_char("end"));
    let expected = inspect(&source.replace(">ENTRYONE<", ">ENTRYONE edited<"));
    for route in [WriteRoute::SourceParagraphs, WriteRoute::Whole] {
        let saved = save(&package(&source)?, Edit::Text("ENTRYONE"), route)?;
        let xml = document_xml(&saved)?;
        assert_eq!(inspect(&xml), expected);
        assert!(xml.contains(&format!("{start}<w:hyperlink")), "{xml}");
        assert!(xml.contains(&format!("</w:hyperlink>{stop}")), "{xml}");
    }
    Ok(())
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
        let source = package(
            &[
                paragraph(0, &(field_start(code) + &entry)),
                paragraph(1, &field_char("end")),
            ]
            .concat(),
        )?;
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
            assert_save(scenario, edit, route, false)
                .unwrap_or_else(|error| panic!("{scenario} {name} {route:?}: {error}"));
        }
    }
}

const WORD_TOC_CODE: &str = r#"TOC \o "1-3" \h \z \u"#;
const WORD_TOC_TITLES: [&str; 3] = ["ENTRYONE", ENTRY_TWO, "ENTRYTHREE"];
const WORD_FIGURE_TITLES: [&str; 3] = ["FIGUREONE", "FIGURETWO", "FIGURETHREE"];

#[derive(Default)]
struct WordTocOptions {
    inline_end: bool,
    numbered: bool,
    unlinked: bool,
    omit_level_one_pages: bool,
}

struct WordToc {
    paragraphs: Vec<String>,
    titles: Vec<&'static str>,
    in_control: bool,
}

impl WordToc {
    fn xml(&self, paragraphs: &[String]) -> String {
        let entries = paragraphs
            .iter()
            .enumerate()
            .map(|(index, content)| paragraph(index, content))
            .collect::<String>();
        let entries = if self.in_control {
            block_sdt(&entries)
        } else {
            entries
        };
        entries + &paragraph(paragraphs.len(), &text_run(BODY_NEXT))
    }
}

fn word_toc(
    code: &str,
    titles: &[&'static str],
    first_anchor: usize,
    options: WordTocOptions,
) -> WordToc {
    let tab = "<w:r><w:rPr><w:webHidden/></w:rPr><w:tab/></w:r>";
    let page = "<w:r><w:rPr><w:webHidden/></w:rPr><w:t>3</w:t></w:r>";
    let end = field_char("end");
    let mut paragraphs = Vec::new();
    for (index, title) in titles.iter().enumerate() {
        let level = if options.numbered || options.omit_level_one_pages {
            [1, 2, 1][index]
        } else {
            index + 1
        };
        let anchor = format!("_Toc{}", first_anchor + index);
        let mut content = format!(r#"<w:pPr><w:pStyle w:val="TOC{level}"/></w:pPr>"#);
        if index == 0 {
            content.push_str(&field_start(code));
        }
        if !options.unlinked {
            content.push_str(&format!(
                r#"<w:hyperlink w:anchor="{anchor}" w:history="1">"#
            ));
        }
        if options.numbered {
            content.push_str(&text_run(["1", "1.1", "2"][index]));
            content.push_str(tab);
        }
        content.push_str(&text_run(title));
        if !options.omit_level_one_pages || level != 1 {
            content.push_str(tab);
            content.push_str(&field_start(&format!(r"PAGEREF {anchor} \h")));
            content.push_str(page);
            content.push_str(&end);
        }
        if !options.unlinked {
            content.push_str("</w:hyperlink>");
        }
        if options.inline_end && index + 1 == titles.len() {
            content.push_str(&end);
        }
        paragraphs.push(content);
    }
    if !options.inline_end {
        paragraphs.push(end);
    }
    WordToc {
        paragraphs,
        titles: titles.to_vec(),
        in_control: false,
    }
}

fn word_toc_w1() -> WordToc {
    word_toc(
        WORD_TOC_CODE,
        &WORD_TOC_TITLES,
        1,
        WordTocOptions::default(),
    )
}

fn word_toc_w2() -> WordToc {
    word_toc(
        WORD_TOC_CODE,
        &WORD_TOC_TITLES,
        1,
        WordTocOptions {
            inline_end: true,
            ..WordTocOptions::default()
        },
    )
}

fn word_toc_w3() -> WordToc {
    word_toc(
        WORD_TOC_CODE,
        &WORD_TOC_TITLES,
        1,
        WordTocOptions {
            numbered: true,
            ..WordTocOptions::default()
        },
    )
}

fn word_toc_w4() -> WordToc {
    word_toc(
        r#"TOC \o "1-3""#,
        &WORD_TOC_TITLES,
        1,
        WordTocOptions {
            unlinked: true,
            ..WordTocOptions::default()
        },
    )
}

fn word_toc_w5() -> WordToc {
    word_toc(
        r#"TOC \o "1-3" \h \z \u \n "1-1""#,
        &WORD_TOC_TITLES,
        1,
        WordTocOptions {
            omit_level_one_pages: true,
            ..WordTocOptions::default()
        },
    )
}

fn word_toc_w6() -> WordToc {
    let mut toc = word_toc_w1();
    toc.in_control = true;
    toc.paragraphs.insert(
        0,
        r#"<w:pPr><w:pStyle w:val="TOCHeading"/></w:pPr>"#.to_owned() + &text_run("Contents"),
    );
    toc
}

fn word_toc_w7() -> WordToc {
    let mut toc = word_toc_w1();
    let figures = word_toc(
        r#"TOC \h \z \c "Figure""#,
        &WORD_FIGURE_TITLES,
        4,
        WordTocOptions::default(),
    );
    toc.paragraphs.extend(figures.paragraphs);
    toc.titles.extend(figures.titles);
    toc
}

fn word_toc_w8() -> WordToc {
    word_toc(
        WORD_TOC_CODE,
        &WORD_TOC_TITLES[..1],
        1,
        WordTocOptions::default(),
    )
}

fn edit_word_toc_first(toc: &WordToc) -> Edit {
    Edit::Text(toc.titles[0])
}

fn edit_word_toc_last(toc: &WordToc) -> Edit {
    Edit::Text(toc.titles[toc.titles.len() - 1])
}

fn edit_word_toc_middle(toc: &WordToc) -> Edit {
    Edit::Text(toc.titles.get(1).copied().unwrap_or(toc.titles[0]))
}

fn insert_after_word_toc_first(toc: &WordToc) -> Edit {
    Edit::Paragraph(ParagraphEdit::InsertAfter(usize::from(toc.in_control)))
}

fn delete_word_toc_middle(toc: &WordToc) -> Edit {
    if toc.titles.len() == 1 {
        Edit::None
    } else {
        Edit::Paragraph(ParagraphEdit::Delete(usize::from(toc.in_control) + 1))
    }
}

fn edit_after_word_toc(_: &WordToc) -> Edit {
    Edit::Text(BODY_NEXT)
}

type WordTocEdit = fn(&WordToc) -> Edit;
const WORD_TOC_EDITS: [(&str, WordTocEdit); 6] = [
    ("E1 edit first title", edit_word_toc_first),
    ("E2 edit last title", edit_word_toc_last),
    ("E3 edit middle title", edit_word_toc_middle),
    ("E4 insert after first entry", insert_after_word_toc_first),
    ("E5 delete middle entry", delete_word_toc_middle),
    ("E6 edit following body paragraph", edit_after_word_toc),
];

fn assert_word_toc_save(toc: &WordToc, edit: Edit, case: &str) -> Result<(), Box<dyn Error>> {
    let source = package(&toc.xml(&toc.paragraphs))?;
    let mut expected_paragraphs = toc.paragraphs.clone();
    match edit {
        Edit::Paragraph(ParagraphEdit::InsertAfter(index)) => {
            expected_paragraphs.insert(index + 1, text_run(INSERTED));
        }
        Edit::Paragraph(ParagraphEdit::Delete(index)) => {
            expected_paragraphs.remove(index);
        }
        Edit::Text(_) => {}
        _ => unreachable!("{case}: unexpected edit"),
    }
    let mut expected_xml = toc.xml(&expected_paragraphs);
    if let Edit::Text(marker) = edit {
        expected_xml = expected_xml.replace(&format!(">{marker}<"), &format!(">{marker} edited<"));
    }
    let inner_story = toc.in_control && !matches!(edit, Edit::Text(BODY_NEXT));
    let story_path = if inner_story {
        "/document/content/0/content"
    } else {
        "/document/content"
    };
    let route = match (toc.in_control, edit) {
        (true, Edit::Paragraph(_)) => WriteRoute::Whole,
        (true, _) => WriteRoute::ChangedParaIds,
        (false, Edit::Paragraph(_)) => WriteRoute::Splice,
        (false, _) => WriteRoute::SourceParagraphs,
    };
    let xml = document_xml(&save_story(&source, edit, route, story_path)?)?;
    assert_spanning_save(&xml, &expected_xml, case);
    if toc.in_control {
        let paragraphs_in_control = xml
            .split_once("<w:sdtContent>")
            .and_then(|(_, content)| content.split_once("</w:sdtContent>"))
            .map(|(content, _)| story_paragraphs(content).count());
        assert_eq!(
            paragraphs_in_control,
            Some(expected_paragraphs.len()),
            "{case}: paragraphs inside the TOC content control"
        );
    }
    Ok(())
}

macro_rules! word_toc_save_tests {
    ($($test:ident: $shape:literal => $build:ident),+ $(,)?) => {
        $(
            #[test]
            fn $test() {
                let toc = $build();
                let mut failed = Vec::new();
                for (name, edit) in WORD_TOC_EDITS {
                    let case = format!("{} {name}", $shape);
                    let edit = edit(&toc);
                    if !matches!(edit, Edit::None)
                        && std::panic::catch_unwind(|| {
                            assert_word_toc_save(&toc, edit, &case)
                                .unwrap_or_else(|error| panic!("{case}: {error}"));
                        })
                        .is_err()
                    {
                        failed.push(case);
                    }
                }
                assert!(failed.is_empty(), "failed cases: {failed:?}");
            }
        )+
    };
}

word_toc_save_tests! {
    word_toc_hyperlinked_plain_saves: "W1" => word_toc_w1,
    word_toc_inline_end_saves: "W2" => word_toc_w2,
    word_toc_numbered_saves: "W3" => word_toc_w3,
    word_toc_unlinked_saves: "W4" => word_toc_w4,
    word_toc_without_level_one_pages_saves: "W5" => word_toc_w5,
    word_toc_in_block_control_saves: "W6" => word_toc_w6,
    word_toc_followed_by_figures_saves: "W7" => word_toc_w7,
    word_toc_single_entry_saves: "W8" => word_toc_w8,
}

#[test]
fn a_spanning_field_keeps_its_end_when_its_closing_paragraph_is_edited() {
    assert_matrix("S1", &[("both sides of the end", Edit::TextOnBothSides)]);
    assert_matrix("TA", &[("closing entry", Edit::Text("ENTRYTHREE"))]);
}

#[test]
fn whole_write_matches_an_end_only_paragraph_after_insertion() -> Result<(), Box<dyn Error>> {
    let source = without_paragraph_ids(&body("TB"));
    let end = without_paragraph_ids(&paragraph(3, &field_char("end")));
    let expected = source.replace(&end, &(paragraph(9, &text_run(INSERTED)) + &end));
    let saved = save(
        &package(&source)?,
        Edit::Paragraph(ParagraphEdit::InsertAfter(2)),
        WriteRoute::Whole,
    )?;
    assert_eq!(inspect(&document_xml(&saved)?), inspect(&expected));
    Ok(())
}

#[test]
fn whole_write_matches_repeated_toc_entries_by_anchor() -> Result<(), Box<dyn Error>> {
    let repeated = |number| entry(number, "Same title", "7");
    let second = paragraph(1, &repeated(2));
    let source = paragraph(0, &(field_start(r"TOC \h") + &repeated(1)))
        + &second
        + &paragraph(2, &(repeated(3) + &field_char("end")))
        + &paragraph(3, &text_run(BODY_NEXT));
    let expected = source.replace(
        &second,
        &(second.clone() + &paragraph(9, &text_run(INSERTED))),
    );
    let saved = save(
        &package(&without_paragraph_ids(&source))?,
        Edit::Paragraph(ParagraphEdit::InsertAfter(1)),
        WriteRoute::Whole,
    )?;
    let xml = document_xml(&saved)?;
    assert_eq!(inspect(&xml), inspect(&expected));
    for number in 1..=3 {
        assert!(
            xml.contains(&format!(r#"w:anchor="_Toc{number}""#)),
            "{xml}"
        );
    }
    Ok(())
}

#[test]
fn block_bookmarks_between_toc_entries_keep_the_field_open() -> Result<(), Box<dyn Error>> {
    let start = r#"<w:bookmarkStart w:id="41" w:name="Entry"/>"#;
    let stop = r#"<w:bookmarkEnd w:id="41"/>"#;
    let middle = paragraph(1, &entry(2, ENTRY_TWO, PAGE_MARKERS[1]));
    let source = body("TA").replace(&middle, &format!("{start}{middle}{stop}"));
    let packaged = package(&source)?;
    for route in [
        WriteRoute::Whole,
        WriteRoute::SourceParagraphs,
        WriteRoute::ChangedParaIds,
    ] {
        for label in ["ENTRYONE", "ENTRYTHREE"] {
            let saved = save(&packaged, Edit::Text(label), route)?;
            let xml = document_xml(&saved)?;
            let expected = source.replace(label, &format!("{label} edited"));
            assert_eq!(inspect(&xml), inspect(&expected));
            assert!(xml.contains(&format!("</w:p>{start}<w:p")), "{xml}");
            assert!(xml.contains(&format!("</w:p>{stop}<w:p")), "{xml}");
        }
    }
    Ok(())
}

#[test]
fn a_toc_closes_after_the_bookmark_end_in_the_last_entry() -> Result<(), Box<dyn Error>> {
    let start = r#"<w:bookmarkStart w:id="41" w:name="Entry"/>"#;
    let stop = r#"<w:bookmarkEnd w:id="41"/>"#;
    let last = entry(3, "ENTRYTHREE", PAGE_MARKERS[2]);
    let source = body("TA").replace(&last, &format!("{start}{last}{stop}"));
    let expected = inspect(&source.replace("ENTRYTHREE", "ENTRYTHREE edited"));
    for route in [
        WriteRoute::Whole,
        WriteRoute::SourceParagraphs,
        WriteRoute::ChangedParaIds,
    ] {
        let saved = save(&package(&source)?, Edit::Text("ENTRYTHREE"), route)?;
        let xml = document_xml(&saved)?;
        assert_eq!(inspect(&xml), expected);
        let closing = format!("</w:hyperlink>{stop}{}", field_char("end"));
        assert!(xml.contains(&closing), "{xml}");
    }
    Ok(())
}

#[test]
fn a_carried_end_follows_an_opener_with_an_empty_result() -> Result<(), Box<dyn Error>> {
    let end = field_char("end");
    let opening = field_start("QUOTE");
    let following = paragraph(2, &text_run(BODY_NEXT));
    let source = [
        paragraph(0, &opening),
        paragraph(1, &end),
        following.clone(),
    ]
    .concat();
    let expected = inspect(&(paragraph(0, &(opening + &end)) + &following));
    for route in [WriteRoute::Splice, WriteRoute::Whole] {
        let saved = save(
            &package(&source)?,
            Edit::Paragraph(ParagraphEdit::Delete(1)),
            route,
        )?;
        assert_eq!(inspect(&document_xml(&saved)?), expected, "{route:?}");
    }
    Ok(())
}

#[test]
fn a_source_run_boundary_keeps_repeated_text_on_its_edited_side() -> Result<(), Box<dyn Error>> {
    let source = body("S1")
        .replace(">SPANEND<", ">A<")
        .replace(">AFTERSPAN<", ">B<");
    let source_bytes = package(&source)?;
    for route in [WriteRoute::SourceParagraphs, WriteRoute::Whole] {
        for (marker, text) in [("A", "AB"), ("B", "AB")] {
            let saved = save(&source_bytes, Edit::Replace(marker, text), route)?;
            assert_eq!(
                inspect(&document_xml(&saved)?),
                inspect(&source.replace(&format!(">{marker}<"), &format!(">{text}<"))),
                "{route:?} {marker}"
            );
        }
    }
    Ok(())
}

#[test]
fn a_removed_toc_closing_paragraph_carries_the_end_to_the_last_surviving_entry()
-> Result<(), Box<dyn Error>> {
    let end = field_char("end");
    let second = entry(2, ENTRY_TWO, PAGE_MARKERS[1]);
    let third = entry(3, "ENTRYTHREE", PAGE_MARKERS[2]);
    let inline_end = body("TA");
    let separate_end = body("TB");
    for (xml, removed, expected) in [
        (
            &inline_end,
            2,
            inline_end
                .replace(&paragraph(2, &format!("{third}{end}")), "")
                .replace(&second, &format!("{second}{end}")),
        ),
        (
            &separate_end,
            3,
            separate_end
                .replace(&paragraph(3, &end), "")
                .replace(&third, &format!("{third}{end}")),
        ),
    ] {
        for route in [WriteRoute::Splice, WriteRoute::Whole] {
            let source_xml = if matches!(route, WriteRoute::Whole) {
                without_paragraph_ids(xml)
            } else {
                xml.to_string()
            };
            let source = package(&source_xml)?;
            let saved = save(
                &source,
                Edit::Paragraph(ParagraphEdit::Delete(removed)),
                route,
            )?;
            assert_eq!(
                inspect(&document_xml(&saved)?),
                inspect(&expected),
                "{route:?}"
            );
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
            assert_eq!(
                inspect(&xml),
                inspect(&source.replace("SPANEND", "SPANEND edited"))
            );
            let revision = xml
                .split_once(&format!(
                    r#"<w:{wrapper} w:id="17" w:author="Reviewer" w:date="2000-01-01T00:00:00Z">"#
                ))
                .and_then(|(_, tail)| tail.split_once(&format!("</w:{wrapper}>")))
                .map(|(revision, _)| revision);
            assert!(
                revision.is_some_and(|revision| revision.contains(END)),
                "{xml}"
            );
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
        paragraph(
            1,
            &format!(
                "{}{end}{}{second}",
                text_run("FIRSTEND"),
                text_run("BETWEEN")
            ),
        ),
        paragraph(
            2,
            &format!("{}{end}{}", text_run("SECONDEND"), text_run("OUTSIDE")),
        ),
    ]
    .concat();
    for route in [
        WriteRoute::SourceParagraphs,
        WriteRoute::ChangedParaIds,
        WriteRoute::Whole,
    ] {
        let saved = save(&package(&source)?, Edit::Text("BETWEEN"), route)?;
        assert_eq!(
            inspect(&document_xml(&saved)?),
            inspect(&source.replace("BETWEEN", "BETWEEN edited"))
        );
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
        format!(
            r#"<w:ins w:id="17" w:author="Reviewer" w:date="2000-01-01T00:00:00Z">{plain_end}{following}</w:ins>"#
        ),
    ] {
        let xml = original.replace(&original_boundary, &boundary);
        let source = package(&xml)?;
        for route in [
            WriteRoute::Whole,
            WriteRoute::SourceParagraphs,
            WriteRoute::ChangedParaIds,
        ] {
            for marker in ["SPANSTART", "SPANEND"] {
                let saved = save(&source, Edit::Text(marker), route)?;
                let saved_xml = document_xml(&saved)?;
                assert_eq!(
                    inspect(&saved_xml),
                    inspect(&xml.replace(marker, &format!("{marker} edited"))),
                    "{route:?}"
                );
                assert!(saved_xml.contains(&boundary), "{saved_xml}");
            }
        }
    }
    Ok(())
}

#[test]
fn tracked_opener_keeps_the_main_save_behavior() -> Result<(), Box<dyn Error>> {
    let empty = r#"<w:ins w:id="17" w:author="Reviewer" w:date="2000-01-01T00:00:00Z">"#;
    let revision = format!(
        "{empty}{}{}</w:ins>",
        field_start("QUOTE"),
        text_run("TRACKED")
    );
    let source = [
        paragraph(0, &revision),
        paragraph(1, &format!("{}{}", text_run("CLOSING"), field_char("end"))),
        paragraph(2, &text_run(BODY_NEXT)),
    ]
    .concat();
    for route in [WriteRoute::Whole, WriteRoute::SourceParagraphs] {
        let saved = save(&package(&source)?, Edit::Text(BODY_NEXT), route)?;
        let xml = document_xml(&saved)?;
        assert!(xml.contains(">BODYNEXT edited<"), "{xml}");
        if matches!(route, WriteRoute::Whole) {
            assert!(xml.contains(&format!("{empty}</w:ins>")), "{xml}");
            assert!(!xml.contains("w:fldChar"), "{xml}");
        } else {
            assert!(xml.contains(&revision), "{xml}");
            assert_eq!(
                inspect(&xml),
                inspect(&source.replace(BODY_NEXT, "BODYNEXT edited"))
            );
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
        paragraph(
            2,
            &format!("{}{end}{}", text_run("OUTEREND"), text_run("AFTER")),
        ),
    ]
    .concat();
    let expected = [
        paragraph(0, &format!("{outer}{end}")),
        paragraph(1, &text_run("INNEREND")),
        paragraph(
            2,
            &format!("{}{}", text_run("OUTEREND"), text_run("AFTER edited")),
        ),
    ]
    .concat();
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
        block_sdt(&box_paragraph),
        format!(
            "<w:tbl><w:tr><w:tc>{}</w:tc></w:tr></w:tbl>",
            paragraph(2, &end)
        ),
        paragraph(
            1,
            &format!(
                "<w:r><w:pict><v:shape><v:textbox><w:txbxContent>{box_paragraph}</w:txbxContent></v:textbox></v:shape></w:pict></w:r>"
            ),
        ),
    ] {
        let source = package(&format!("{source_opening}{container}"))?;
        let saved = save(&source, Edit::Text("OUTER"), WriteRoute::SourceParagraphs)?;
        let xml = document_xml(&saved)?;
        let opening = &xml[..xml.find("</w:p>").unwrap() + "</w:p>".len()];
        assert_eq!(
            inspect(opening),
            inspect(&paragraph(
                0,
                &format!("{}{end}", opener.replace(">OUTER<", ">OUTER edited<"))
            ))
        );
        assert!(xml.contains(&container), "{xml}");
    }
    Ok(())
}

#[test]
fn a_field_spanning_paragraphs_in_one_cell_keeps_its_end() -> Result<(), Box<dyn Error>> {
    let source = format!("<w:tbl><w:tr><w:tc>{}</w:tc></w:tr></w:tbl>", body("S1"));
    let packaged = package(&source)?;
    for route in [WriteRoute::ChangedParaIds, WriteRoute::Whole] {
        for label in ["SPANSTART", "SPANEND"] {
            let saved = save_story(
                &packaged,
                Edit::Text(label),
                route,
                "/document/content/0/rows/0/cells/0/content",
            )?;
            let expected = source.replace(label, &format!("{label} edited"));
            assert_eq!(inspect(&document_xml(&saved)?), inspect(&expected));
        }
    }
    Ok(())
}

#[test]
fn cross_cell_closes_keep_the_legacy_opening_paragraph() -> Result<(), Box<dyn Error>> {
    let end = field_char("end");
    let opening = field_start("QUOTE") + &text_run("SPANSTART");
    let first_cell = paragraph(0, &opening) + &paragraph(1, &text_run("CONTINUATION"));
    let second_cell = paragraph(2, &end);
    let source = package(&format!(
        "<w:tbl><w:tr><w:tc>{first_cell}</w:tc><w:tc>{second_cell}</w:tc></w:tr></w:tbl>"
    ))?;
    let saved = save_story(
        &source,
        Edit::Text("SPANSTART"),
        WriteRoute::ChangedParaIds,
        "/document/content/0/rows/0/cells/0/content",
    )?;
    let xml = document_xml(&saved)?;
    let first_cell = xml.split_once("</w:tc>").unwrap().0;
    assert_eq!(inspect(first_cell).fields[0].end, Some(0), "{xml}");
    assert!(
        xml.contains(&format!("<w:tc>{second_cell}</w:tc>")),
        "{xml}"
    );
    Ok(())
}

#[test]
fn a_shared_end_and_begin_run_keeps_the_legacy_opening_paragraph() -> Result<(), Box<dyn Error>> {
    let first = field_start("QUOTE FIRST") + &text_run("FIRST");
    let shared = [
        format!("<w:r><w:fldChar {END}/><w:fldChar {BEGIN}/></w:r>"),
        "<w:r><w:instrText>QUOTE SECOND</w:instrText></w:r>".to_owned(),
        field_char("separate"),
        text_run("SECOND"),
    ]
    .concat();
    let source = [
        paragraph(0, &first),
        paragraph(1, &shared),
        paragraph(2, &(text_run("LAST") + &field_char("end"))),
    ]
    .concat();
    for route in [WriteRoute::SourceParagraphs, WriteRoute::Whole] {
        let saved = save(&package(&source)?, Edit::Text("FIRST"), route)?;
        let xml = document_xml(&saved)?;
        let opening = xml.split_once("</w:p>").unwrap().0;
        let first = first.replace(">FIRST<", ">FIRST edited<") + &field_char("end");
        assert_eq!(inspect(opening), inspect(&paragraph(0, &first)));
        assert!(xml.contains(">SECOND<") && xml.contains(">LAST<"), "{xml}");
        if matches!(route, WriteRoute::SourceParagraphs) {
            assert!(xml.contains(&shared), "{xml}");
        }
    }
    Ok(())
}

#[test]
fn selective_writes_without_a_span_keep_the_other_paragraph_verbatim() -> Result<(), Box<dyn Error>>
{
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
    assert_eq!(
        document_xml(&saved)?,
        document_xml(&source)?.replace(BODY_NEXT, "BODYNEXT edited")
    );
    Ok(())
}

#[test]
fn a_same_paragraph_raw_end_stays_on_its_original_side_of_the_next_field()
-> Result<(), Box<dyn Error>> {
    let end = field_char("end");
    let preceding = format!(
        r#"<w:hyperlink w:anchor="top">{}</w:hyperlink>{end}"#,
        field_char("begin")
    );
    let next = field_start("QUOTE NEXT") + &text_run("NEXT") + &end;
    let source = paragraph(0, &format!("{preceding}{next}"));
    for route in [WriteRoute::SourceParagraphs, WriteRoute::ChangedParaIds] {
        let saved = save(&package(&source)?, Edit::Text("NEXT"), route)?;
        assert_eq!(
            inspect(&document_xml(&saved)?),
            inspect(&source.replace(">NEXT<", ">NEXT edited<"))
        );
    }
    Ok(())
}

#[test]
fn finding_s2_mixed_bookmark_endpoints_keep_the_closing_field_end() -> Result<(), Box<dyn Error>> {
    let start = r#"<w:bookmarkStart w:id="41" w:name="Toc1"/>"#;
    let stop = r#"<w:bookmarkEnd w:id="41"/>"#;
    let middle = paragraph(1, &entry(2, ENTRY_TWO, PAGE_MARKERS[1]));
    let last = entry(3, "ENTRYTHREE", PAGE_MARKERS[2]);
    let source = body("TA")
        .replace(&middle, &format!("{start}{middle}"))
        .replace(&last, &format!("{last}{stop}"));
    let saved = save(
        &package(&source)?,
        Edit::Text("ENTRYTHREE"),
        WriteRoute::ChangedParaIds,
    )?;
    assert_spanning_save(
        &document_xml(&saved)?,
        &source.replace("ENTRYTHREE", "ENTRYTHREE edited"),
        "finding_s2_mixed_bookmarks",
    );
    Ok(())
}

#[test]
fn finding_s3_opening_bookmark_keeps_the_unlinked_pageref() -> Result<(), Box<dyn Error>> {
    let mut toc = word_toc_w4();
    let title = text_run("ENTRYONE");
    let bookmarked = format!(
        r#"<w:bookmarkStart w:id="41" w:name="Toc1"/>{title}<w:bookmarkEnd w:id="41"/>"#
    );
    toc.paragraphs[0] = toc.paragraphs[0].replace(&title, &bookmarked);
    let source = toc.xml(&toc.paragraphs);
    let saved = save(
        &package(&source)?,
        Edit::Text("ENTRYONE"),
        WriteRoute::ChangedParaIds,
    )?;
    assert_spanning_save(
        &document_xml(&saved)?,
        &source.replace("ENTRYONE", "ENTRYONE edited"),
        "finding_s3_bookmark_and_pageref",
    );
    Ok(())
}

#[test]
fn finding_s6_whole_save_matches_unlinked_entries_by_pageref() -> Result<(), Box<dyn Error>> {
    let toc = word_toc(
        "TOC",
        &["Same title"; 3],
        1,
        WordTocOptions {
            inline_end: true,
            unlinked: true,
            ..WordTocOptions::default()
        },
    );
    let source = without_paragraph_ids(&toc.xml(&toc.paragraphs));
    let mut expected = toc.paragraphs.clone();
    expected.insert(2, text_run(INSERTED));
    let saved = save(
        &package(&source)?,
        Edit::Paragraph(ParagraphEdit::InsertAfter(1)),
        WriteRoute::Whole,
    )?;
    assert_spanning_save(
        &document_xml(&saved)?,
        &toc.xml(&expected),
        "finding_s6_repeated_unlinked_entries",
    );
    Ok(())
}

#[test]
fn finding_n2_deleting_the_closer_keeps_the_end_before_trailing_blocks()
-> Result<(), Box<dyn Error>> {
    let mut toc = word_toc_w4();
    let end = field_char("end");
    toc.paragraphs.truncate(3);
    toc.paragraphs[2].push_str(&end);
    let bookmark_start = r#"<w:bookmarkStart w:id="41" w:name="Toc1"/>"#;
    let bookmark_end = r#"<w:bookmarkEnd w:id="41"/>"#;
    let tail = paragraph(3, &text_run(BODY_NEXT));
    let source = format!(
        "{bookmark_start}{}{bookmark_end}{tail}",
        toc.paragraphs
            .iter()
            .enumerate()
            .map(|(i, p)| paragraph(i, p))
            .collect::<String>()
    );
    let expected = format!(
        "{bookmark_start}{}{}{bookmark_end}",
        paragraph(0, &toc.paragraphs[0]),
        paragraph(1, &(toc.paragraphs[1].clone() + &end)),
    );
    let bytes = package(&source)?;
    let mut request = save_request(&bytes, true)?;
    let blocks = request["document"]["content"].as_array_mut().unwrap();
    blocks.retain(|block| !matches!(block["sourceOrdinal"].as_u64(), Some(2 | 3)));
    let digest = format!("{:x}", Sha256::digest(document_xml(&bytes)?.as_bytes()));
    request["paragraphIds"] = json!({"splicedParts": [{
        "part": DOCUMENT_PART, "sha256": digest, "paragraphs": [0, 1], "removed": [2, 3],
    }]});
    let saved = write_docx_s13(serde_json::from_value(request)?, &bytes)?;
    assert_spanning_save(
        &document_xml(&saved)?,
        &expected,
        "finding_n2_trailing_blocks",
    );
    Ok(())
}

#[test]
fn finding_n3_insertion_next_to_a_container_keeps_its_span() -> Result<(), Box<dyn Error>> {
    let before = paragraph(4, &text_run("BEFORE"));
    let tail = paragraph(5, &text_run(BODY_NEXT));
    let inner = body("S1");
    let containers = [
        (
            "table",
            format!("<w:tbl><w:tr><w:tc>{inner}</w:tc></w:tr></w:tbl>"),
        ),
        (
            "sdt",
            block_sdt(&inner).replace(r#"<w:id w:val="42"/>"#, ""),
        ),
    ];
    let mut cases = Vec::new();
    for (container, inner) in containers {
        let source = format!("{before}{inner}{tail}");
        let saved = save(
            &package(&source)?,
            Edit::Paragraph(ParagraphEdit::InsertAfter(0)),
            WriteRoute::Whole,
        )?;
        cases.push((
            document_xml(&saved)?,
            format!("{before}{}{inner}{tail}", paragraph(6, &text_run(INSERTED))),
            format!("finding_n3_insert_before_{container}"),
        ));
    }
    assert_spanning_cases(&cases);
    Ok(())
}

#[test]
fn finding_n5_the_end_follows_a_simple_field_or_inline_control() -> Result<(), Box<dyn Error>> {
    let text = text_run("LEADING");
    let wrappers = [
        (
            "simple_field",
            format!(r#"<w:fldSimple w:instr="QUOTE">{text}</w:fldSimple>"#),
        ),
        (
            "inline_control",
            format!(r#"<w:sdt><w:sdtPr><w:id w:val="42"/></w:sdtPr><w:sdtContent>{text}</w:sdtContent></w:sdt>"#),
        ),
    ];
    let mut cases = Vec::new();
    for (wrapper, leading) in wrappers {
        let source = body("S1").replace(&text_run("SPANEND"), &leading);
        let saved = save(&package(&source)?, Edit::Text("AFTERSPAN"), WriteRoute::Whole)?;
        cases.push((
            document_xml(&saved)?,
            source.replace("AFTERSPAN", "AFTERSPAN edited"),
            format!("finding_n5_end_after_{wrapper}"),
        ));
    }
    assert_spanning_cases(&cases);
    Ok(())
}

#[test]
fn finding_n6_the_end_splits_a_run_before_its_tab() -> Result<(), Box<dyn Error>> {
    let tabbed = "<w:r><w:tab/><w:t>AFTERSPAN</w:t></w:r>";
    let source = body("S1").replace(&text_run("AFTERSPAN"), tabbed);
    let bytes = package(&source)?;
    let mut request = save_request(&bytes, false)?;
    let content = request
        .pointer_mut("/document/content/1/content")
        .and_then(Value::as_array_mut)
        .ok_or("missing closing paragraph runs")?;
    content.clear();
    content.push(json!({
        "type": "run",
        "content": [
            {"type": "text", "text": "SPANEND edited"},
            {"type": "tab"},
            {"type": "text", "text": "AFTERSPAN"},
        ],
    }));
    let saved = write_docx_s13(serde_json::from_value(request)?, &bytes)?;
    assert_spanning_save(
        &document_xml(&saved)?,
        &source.replace("SPANEND", "SPANEND edited"),
        "finding_n6_end_before_tab_in_run",
    );
    Ok(())
}

#[test]
fn finding_n7_selective_save_accounts_for_shapes_in_spanning_results()
-> Result<(), Box<dyn Error>> {
    let shape = concat!(
        r#"<w:r><w:drawing><wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">"#,
        r#"<wp:extent cx="914400" cy="914400"/><a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">"#,
        r#"<a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">"#,
        r#"<wps:wsp xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:spPr>"#,
        r#"<a:prstGeom prst="rect"/></wps:spPr><wps:bodyPr/></wps:wsp>"#,
        "</a:graphicData></a:graphic></wp:inline></w:drawing></w:r>"
    );
    let control = format!("<w:sdt><w:sdtPr/><w:sdtContent>{shape}</w:sdtContent></w:sdt>");
    let linked = |name: &str| {
        format!(
            r#"<w:hyperlink w:anchor="{name}">{}{control}</w:hyperlink>"#,
            text_run(name),
        )
    };
    let end = field_char("end");
    let source = [
        paragraph(0, &(field_start("QUOTE FIRST") + &linked("FIRST"))),
        paragraph(1, &(text_run("FIRSTEND") + &end)),
        paragraph(2, &(field_start("QUOTE SECOND") + &linked("SECOND"))),
        paragraph(3, &(text_run("SECONDEND") + &end)),
        paragraph(4, &text_run(BODY_NEXT)),
    ]
    .concat();
    let bytes = package(&source)?;
    let selective = document_xml(&save(&bytes, Edit::Text("SECOND"), WriteRoute::ChangedParaIds)?)?;
    assert_spanning_save(
        &selective,
        &source.replace(">SECOND<", ">SECOND edited<"),
        "finding_n7_selective_shape_ids",
    );
    let whole = document_xml(&save(&bytes, Edit::Text("SECOND"), WriteRoute::Whole)?)?;
    let final_id = |xml: &str| {
        xml.rsplit_once("<wp:docPr")
            .and_then(|(_, tag)| attribute(tag.split_once('>')?.0, "id"))
            .map(str::to_owned)
    };
    let whole_id = final_id(&whole);
    assert!(whole_id.is_some(), "shape in the edited field has no DrawingML ID");
    assert_eq!(final_id(&selective), whole_id);
    Ok(())
}
