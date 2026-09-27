//! Content controls under paragraph identities: fills allocate paragraph identities as other
//! batch paragraphs do, and control ids and anchors name paragraphs by the session keys batches
//! target and exports report.

#[allow(dead_code)]
#[path = "support/structured_fixture.rs"]
mod fixture;

use docx_edit::content_controls::{Anchor, ContentControlsOptions};
use docx_edit::structured::{BlockKind, ExportOptions, InlineKind, RevisionView};
use docx_edit::{
    EditApplication, EditCtx, EditFailureCode, EditFailureReason, EditRequest, EditingDoc,
    FormatPolicy, ParagraphIdOrigin, ParagraphIdentity, ParagraphOrigin, ParagraphRef, Position,
    RawOp, UndoSession, seed_from_docx,
};
use fixture::{Package, para, run};
use serde_json::{Value, json};
use yrs::Any;

/// The child story of the first block control in the body.
const CONTROL: &str = "body:sdt0";

fn ctx() -> EditCtx {
    EditCtx::local("", "")
}

fn open(bytes: &[u8]) -> EditingDoc {
    let doc = EditingDoc::new(4242);
    seed_from_docx(&doc, bytes).unwrap();
    doc
}

fn sdt(properties: &str, content: &str) -> String {
    format!("<w:sdt><w:sdtPr>{properties}</w:sdtPr><w:sdtContent>{content}</w:sdtContent></w:sdt>")
}

fn block_sdt(tag: &str, content: &str) -> String {
    sdt(&format!(r#"<w:tag w:val="{tag}"/><w:richText/>"#), content)
}

fn inline_sdt(tag: &str, text: &str) -> String {
    sdt(&format!(r#"<w:tag w:val="{tag}"/><w:text/>"#), &run(text))
}

/// A block control holding one paragraph, then a body paragraph; `comment` is the body of a
/// comment the package retains outside every story.
fn address_package(comment: Option<&str>) -> Vec<u8> {
    let body = format!(
        "{}{}",
        block_sdt("address", &para("0E000001", &run("One"))),
        para("0E000002", &run("After"))
    );
    let package = Package::new(&body);
    match comment {
        Some(paragraph) => package.part(
            "comments.xml",
            "rIdComments",
            "comments",
            "comments",
            &format!(
                r#"<w:comments {}><w:comment w:id="9" w:author="Ann">{paragraph}</w:comment></w:comments>"#,
                fixture::namespaces()
            ),
        ),
        None => package,
    }
    .bytes()
}

fn fill(control_id: &str, text: &str) -> Value {
    json!({"op": "setContentControlText", "target": {"kind": "id", "controlId": control_id}, "text": text})
}

fn request(doc: &EditingDoc, steps: Vec<Value>) -> EditRequest {
    serde_json::from_value(json!({"expectVersion": doc.version().as_str(), "steps": steps}))
        .unwrap()
}

fn apply(doc: &EditingDoc, steps: Vec<Value>) -> EditApplication {
    doc.apply_edits(&request(doc, steps), &UndoSession::new())
        .unwrap()
        .unwrap_or_else(|refusal| panic!("refused: {refusal:?}"))
}

fn reason(doc: &EditingDoc, steps: Vec<Value>) -> (EditFailureCode, Option<EditFailureReason>) {
    let failure = doc
        .apply_edits(&request(doc, steps), &UndoSession::new())
        .unwrap()
        .expect_err("the batch is refused")
        .failure;
    (failure.code, failure.reason)
}

fn new_keys(applied: &EditApplication) -> Vec<String> {
    applied.receipts[0]
        .new_paragraphs
        .iter()
        .map(|paragraph| paragraph.para_id.clone())
        .collect()
}

fn identities(doc: &EditingDoc) -> Vec<ParagraphIdentity> {
    doc.paragraph_identities().paragraphs
}

fn identity(doc: &EditingDoc, story: &str, key: &str) -> ParagraphIdentity {
    identities(doc)
        .into_iter()
        .find(|identity| {
            identity.paragraph
                == ParagraphRef::Session {
                    story: story.to_owned(),
                    para_id: key.to_owned(),
                }
        })
        .unwrap_or_else(|| panic!("no paragraph {key} in {story}"))
}

fn texts(doc: &EditingDoc, story: &str) -> Vec<String> {
    doc.paragraphs(story)
        .unwrap()
        .into_iter()
        .map(|paragraph| paragraph.text)
        .collect()
}

/// Types the paragraphs a fill of `lines` after the control's first paragraph adds: a split at
/// the end of the previous paragraph, then its text.
fn type_after(doc: &EditingDoc, first: &str, lines: &[&str]) {
    let mut previous = first.to_owned();
    for line in lines {
        let mark = doc.paragraph_mark_position(&previous).unwrap();
        let split = doc.split_paragraph(&ctx(), mark, None).unwrap();
        let at = doc.paragraph_mark_position(&split.second_para_id).unwrap();
        doc.insert_text(&ctx(), at, line, FormatPolicy::Inherit)
            .unwrap();
        previous = split.second_para_id;
    }
}

#[test]
fn block_fills_allocate_and_adopt_identities_as_direct_edits_do() {
    let bytes = address_package(None);
    let (filled, typed) = (open(&bytes), open(&bytes));
    let mut reserved = String::new();
    for doc in [&filled, &typed] {
        let minted = doc.create_story("probe", "", "Normal", "left").unwrap();
        let (client, counter) = minted.split_once(':').unwrap();
        reserved = format!("{client}:{}", counter.parse::<u64>().unwrap() + 1);
        let ops = [
            RawOp::InsertEmbed {
                index: 0,
                kind: "pilcrow".to_owned(),
                payload: vec![("paraId".to_owned(), Any::from(reserved.as_str()))],
                attrs: Default::default(),
            },
            RawOp::Delete { index: 0, len: 1 },
        ];
        for op in ops {
            doc.apply_raw_ops(CONTROL, vec![op], &ctx()).unwrap();
        }
    }
    let applied = apply(&filled, vec![fill(CONTROL, "One\nNew\nMore")]);
    let minted = new_keys(&applied);
    assert_eq!(minted.len(), 2);
    assert!(
        !minted.contains(&reserved),
        "a deleted paragraph's key stays reserved"
    );
    type_after(&typed, "0E000001", &["New", "More"]);

    assert_eq!(texts(&filled, CONTROL), ["One", "New", "More"]);
    assert_eq!(texts(&filled, CONTROL), texts(&typed, CONTROL));
    assert_eq!(identities(&filled), identities(&typed));
    let mut ids = Vec::new();
    for key in &minted {
        let identity = identity(&filled, CONTROL, key);
        assert_eq!(identity.origin, ParagraphOrigin::Authored);
        assert_eq!(identity.id_origin, Some(ParagraphIdOrigin::Authored));
        ids.push(
            identity
                .ooxml_para_id
                .expect("a filled paragraph is claimed an ID"),
        );
    }
    assert_eq!(
        identity(&filled, CONTROL, "0E000001")
            .ooxml_para_id
            .as_deref(),
        Some("0E000001")
    );
    assert!(ids[0] != ids[1] && ids.iter().all(|id| id != "0E000001" && id != "0E000002"));
}

#[test]
fn block_fills_never_take_an_id_retained_content_holds() {
    let probe = open(&address_package(None));
    let key = new_keys(&apply(&probe, vec![fill(CONTROL, "One\nNew")]))[0].clone();
    let natural = identity(&probe, CONTROL, &key)
        .ooxml_para_id
        .expect("a filled paragraph saves with a Word paragraph ID");
    let doc = open(&address_package(Some(&para(&natural, &run("Remark")))));
    assert_eq!(
        new_keys(&apply(&doc, vec![fill(CONTROL, "One\nNew")])),
        std::slice::from_ref(&key)
    );
    let id = identity(&doc, CONTROL, &key).ooxml_para_id.unwrap();
    assert_ne!(id, natural, "the comment paragraph keeps its ID");
}

#[test]
fn a_fill_authors_into_an_editor_only_control_paragraph() {
    let bytes = Package::new(&format!(
        "{}{}",
        block_sdt("empty", ""),
        para("0E000002", &run("After"))
    ))
    .bytes();
    let (filled, typed) = (open(&bytes), open(&bytes));
    let key = filled.paragraphs(CONTROL).unwrap()[0].para_id.clone();
    assert_eq!(
        identity(&filled, CONTROL, &key).origin,
        ParagraphOrigin::Synthetic
    );
    apply(&filled, vec![fill(CONTROL, "Filled")]);
    let at = typed.paragraph_mark_position(&key).unwrap();
    typed
        .insert_text(&ctx(), at, "Filled", FormatPolicy::Inherit)
        .unwrap();

    assert_eq!(texts(&filled, CONTROL), ["Filled"]);
    assert_eq!(identities(&filled), identities(&typed));
    let identity = identity(&filled, CONTROL, &key);
    assert_eq!(identity.origin, ParagraphOrigin::Authored);
    assert_eq!(identity.id_origin, Some(ParagraphIdOrigin::Authored));
    assert!(identity.ooxml_para_id.is_some());
}

/// Each listed control's id, tag and the session key its id embeds.
fn listed(doc: &EditingDoc) -> Vec<(String, String, String)> {
    doc.list_content_controls(&ContentControlsOptions::default())
        .unwrap()
        .content
        .controls
        .into_iter()
        .map(|control| {
            let id = control.metadata.control_id;
            let key = id.split('|').nth(1).unwrap().to_owned();
            assert_eq!(
                control.anchor,
                Anchor::Control {
                    story: "body".to_owned(),
                    control_id: id.clone(),
                }
            );
            (id, control.metadata.tag.unwrap(), key)
        })
        .collect()
}

/// Each exported inline control's id, the key of the paragraph block holding it and the key its
/// own range anchor names.
fn exported(doc: &EditingDoc) -> Vec<(String, String, String)> {
    let read = doc
        .export_structured(&ExportOptions::new(RevisionView::Accepted))
        .unwrap();
    let mut found = Vec::new();
    for block in read.content.stories.iter().flat_map(|story| &story.blocks) {
        let BlockKind::Paragraph { paragraph } = &block.content else {
            continue;
        };
        let Anchor::Paragraph { para_id, .. } = &block.anchor else {
            panic!("paragraph without a location: {block:?}");
        };
        for inline in &paragraph.inlines {
            if let InlineKind::ContentControl { control, .. } = &inline.content {
                let Anchor::Range(range) = &inline.anchor else {
                    panic!("control without a range: {inline:?}");
                };
                assert_eq!(range.start.para_id, range.end.para_id);
                found.push((
                    control.control_id.clone(),
                    para_id.clone(),
                    range.start.para_id.clone(),
                ));
            }
        }
    }
    found
}

#[test]
fn control_ids_name_paragraphs_by_the_keys_batches_and_exports_use() {
    let bytes = Package::new(
        &[
            para("0F000001", &inline_sdt("first", "a")),
            para("0F000001", &inline_sdt("repeated", "b")),
            format!("<w:p>{}</w:p>", inline_sdt("missing", "c")),
            para(
                "0F000003",
                &format!("{}{}", run("Lead "), inline_sdt("moved", "d")),
            ),
        ]
        .concat(),
    )
    .bytes();
    let doc = open(&bytes);
    let mark = doc.paragraph_mark_position("0F000003").unwrap();
    let split = doc
        .split_paragraph(&ctx(), Position::new("body", mark.index - 1), None)
        .unwrap();
    let controls = listed(&doc);
    let keys: Vec<(&str, &str)> = controls
        .iter()
        .map(|(_, tag, key)| (tag.as_str(), key.as_str()))
        .collect();
    assert_eq!(
        keys,
        [
            ("first", "0F000001"),
            ("repeated", "body:p1"),
            ("missing", "body:p2"),
            ("moved", split.second_para_id.as_str()),
        ],
        "session keys, not Word paragraph IDs"
    );
    let expected: Vec<(String, String, String)> = controls
        .iter()
        .map(|(id, _, key)| (id.clone(), key.clone(), key.clone()))
        .collect();
    assert_eq!(exported(&doc), expected);

    let applied = apply(
        &doc,
        controls
            .iter()
            .map(|(id, tag, _)| fill(id, &tag.to_uppercase()))
            .collect(),
    );
    assert!(applied.applied);
    let steps: Vec<Value> = controls
        .iter()
        .map(|(_, tag, key)| {
            json!({
                "op": "insertText",
                "target": {"kind": "paragraph", "story": "body", "paraId": key},
                "at": "start",
                "text": format!("{tag}:"),
            })
        })
        .collect();
    apply(&doc, steps);
    assert_eq!(listed(&doc), controls, "edits around a control keep its id");
    assert_eq!(exported(&doc), expected);
    let paragraphs = doc.paragraphs("body").unwrap();
    for (_, tag, key) in &controls {
        let paragraph = paragraphs
            .iter()
            .find(|paragraph| &paragraph.para_id == key)
            .unwrap();
        assert!(paragraph.text.starts_with(&format!("{tag}:")), "{key}");
    }
    let values: Vec<String> = doc
        .list_content_controls(&ContentControlsOptions::default())
        .unwrap()
        .content
        .controls
        .into_iter()
        .map(|control| match control.value {
            docx_edit::content_controls::ControlValue::Text { text } => text,
            other => panic!("no text value: {other:?}"),
        })
        .collect();
    assert_eq!(values, ["FIRST", "REPEATED", "MISSING", "MOVED"]);
}

#[test]
fn controls_in_paragraphs_repeating_a_word_id_keep_their_own_safety() {
    let properties = r#"<w:tag w:val="twin"/><w:text/>"#;
    let marked = format!(
        r#"<w:bookmarkStart w:id="1" w:name="kept"/>{}<w:bookmarkEnd w:id="1"/>"#,
        run("marked")
    );
    let bytes = Package::new(&format!(
        "{}{}",
        para("0F000030", &sdt(properties, &run("plain"))),
        para("0F000030", &sdt(properties, &marked))
    ))
    .bytes();
    let doc = open(&bytes);
    let ids: Vec<String> = listed(&doc).into_iter().map(|(id, _, _)| id).collect();
    assert_eq!(ids, ["body|0F000030|0", "body|body:p1|0"]);
    apply(&doc, vec![fill(&ids[0], "filled")]);
    assert_eq!(
        reason(&doc, vec![fill(&ids[1], "x")]),
        (
            EditFailureCode::Unsupported,
            Some(EditFailureReason::UnsupportedChildren)
        )
    );
}
