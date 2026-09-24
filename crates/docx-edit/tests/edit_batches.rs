use std::cell::Cell;
use std::rc::Rc;
use std::sync::Arc;
use std::sync::atomic::{AtomicU64, Ordering};

use docx_edit::{
    AtomKind, ChangeTarget, DocumentVersion, EditApplication, EditCtx, EditFailureCode, EditGuard,
    EditHistory, EditOperation, EditRefusal, EditRequest, EditSource, EditStep, EditSuggestion,
    EditTarget, EditTextView, EditingDoc, FindTextRequest, FormatPolicy, ParaAttrDelta,
    ParaSelector, ParagraphInput, ParagraphTarget, Position, RawOp, ReadParagraphsRequest,
    SearchScope, SegmentContent, StoryRange, TargetEdge, TextPosition, TextRange, TextTarget,
    UndoCaptureMode, UndoSession, seed_from_docx,
};
use yrs::Any;

const NS: &str = concat!(
    r#"xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" "#,
    r#"xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" "#,
    r#"xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" "#,
    r#"xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" "#,
    r#"xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" "#,
    r#"xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" "#,
    r#"xmlns:bofx="urn:fidelity""#
);

const STYLES: &str = r#"<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:pPr><w:spacing w:after="120"/></w:pPr></w:style><w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:pPr><w:keepNext/><w:spacing w:before="240" w:after="60"/></w:pPr><w:rPr><w:b/><w:sz w:val="32"/></w:rPr></w:style><w:style w:type="paragraph" w:styleId="Quote"><w:name w:val="Quote"/><w:pPr><w:ind w:left="720"/><w:jc w:val="center"/></w:pPr><w:rPr><w:i/><w:caps/></w:rPr></w:style><w:style w:type="paragraph" w:styleId="ListNumber"><w:name w:val="List Number"/><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr></w:style><w:style w:type="character" w:styleId="Strong"><w:name w:val="Strong"/><w:rPr><w:b/></w:rPr></w:style>"#;

const DATE: &str = "2026-09-24T12:00:00Z";

fn docx(body: &str) -> Vec<u8> {
    let content_types = r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/><Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/><Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/></Types>"#;
    let root_rels = r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"#;
    let rel = |id: &str, kind: &str, target: &str| {
        format!(
            r#"<Relationship Id="{id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/{kind}" Target="{target}"/>"#
        )
    };
    let document_rels = format!(
        r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">{}{}{}{}{}</Relationships>"#,
        rel("rIdStyles", "styles", "styles.xml"),
        rel("rIdHeader", "header", "header1.xml"),
        rel("rIdNotes", "footnotes", "footnotes.xml"),
        rel("rIdComments", "comments", "comments.xml"),
        rel("rIdImage", "image", "media/one.png"),
    );
    let document = format!(
        r#"<w:document {NS}><w:body>{body}<w:sectPr><w:headerReference w:type="default" r:id="rIdHeader"/></w:sectPr></w:body></w:document>"#
    );
    let header = format!(
        r#"<w:hdr {NS}><w:p w14:paraId="0000E001"><w:r><w:t>Header text</w:t></w:r></w:p></w:hdr>"#
    );
    let notes = format!(
        r#"<w:footnotes {NS}><w:footnote w:id="1"><w:p><w:r><w:t>Note text</w:t></w:r></w:p></w:footnote></w:footnotes>"#
    );
    let comments = format!(
        r#"<w:comments {NS}><w:comment w:id="9" w:author="Ann" w:date="{DATE}"><w:p><w:r><w:t>Remark</w:t></w:r></w:p></w:comment></w:comments>"#
    );
    let styles = format!(r#"<w:styles {NS}>{STYLES}</w:styles>"#);
    ooxml_opc::rezip_parts(&[
        (
            "[Content_Types].xml".to_owned(),
            content_types.as_bytes().to_vec(),
        ),
        ("_rels/.rels".to_owned(), root_rels.as_bytes().to_vec()),
        (
            "word/_rels/document.xml.rels".to_owned(),
            document_rels.into_bytes(),
        ),
        ("word/document.xml".to_owned(), document.into_bytes()),
        ("word/styles.xml".to_owned(), styles.into_bytes()),
        ("word/header1.xml".to_owned(), header.into_bytes()),
        ("word/footnotes.xml".to_owned(), notes.into_bytes()),
        ("word/comments.xml".to_owned(), comments.into_bytes()),
        ("word/media/one.png".to_owned(), vec![137, 80, 78, 71]),
    ])
    .unwrap()
}

fn p(id: &str, content: &str) -> String {
    format!(r#"<w:p w14:paraId="{id}">{content}</w:p>"#)
}

fn styled(id: &str, style: &str, content: &str) -> String {
    format!(r#"<w:p w14:paraId="{id}"><w:pPr><w:pStyle w:val="{style}"/></w:pPr>{content}</w:p>"#)
}

fn r(text: &str) -> String {
    format!(r#"<w:r><w:t xml:space="preserve">{text}</w:t></w:r>"#)
}

const BREAK: &str = "<w:r><w:br/></w:r>";
const IMAGE: &str = r#"<w:r><w:drawing><wp:inline><wp:extent cx="914400" cy="457200"/><wp:docPr id="1" name="one"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:blipFill><a:blip r:embed="rIdImage"/></pic:blipFill></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>"#;
const INLINE_CONTROL: &str = r#"<w:sdt><w:sdtPr><w:tag w:val="inline"/></w:sdtPr><w:sdtContent><w:r><w:t>ctl</w:t></w:r></w:sdtContent></w:sdt>"#;
const NOTE_REFERENCE: &str = r#"<w:r><w:footnoteReference w:id="1"/></w:r>"#;
const RAW: &str = r#"<bofx:block bofx:value="opaque"/>"#;
const TAB: &str = "<w:r><w:tab/></w:r>";

fn open(body: &str) -> EditingDoc {
    let doc = EditingDoc::new(7001);
    seed_from_docx(&doc, &docx(body)).unwrap();
    doc
}

fn basic_body() -> String {
    format!(
        "{}{}{}{}",
        p("00000001", &r("Alpha beta gamma")),
        p(
            "00000002",
            &format!("{}{BREAK}{}", r("Line one"), r("target words"))
        ),
        styled("00000003", "Heading1", &r("Title")),
        p("00000004", &r("Delta")),
    )
}

fn basic() -> EditingDoc {
    open(&basic_body())
}

fn para(story: &str, id: &str) -> ParagraphTarget {
    ParagraphTarget {
        story: story.to_owned(),
        para_id: id.to_owned(),
    }
}

fn body(id: &str) -> ParagraphTarget {
    para("body", id)
}

fn search(text: &str, id: &str) -> TextTarget {
    TextTarget::Search {
        text: text.to_owned(),
        within: SearchScope::Paragraph(body(id)),
        view: EditTextView::Accepted,
    }
}

fn range(id: &str, start: u32, end: u32, view: EditTextView) -> TextTarget {
    TextTarget::Range(TextRange {
        story: "body".to_owned(),
        start: TextPosition {
            para_id: id.to_owned(),
            offset: start,
        },
        end: TextPosition {
            para_id: id.to_owned(),
            offset: end,
        },
        view,
    })
}

fn step(operation: EditOperation) -> EditStep {
    EditStep::new(operation)
}

fn replace(target: TextTarget, text: &str) -> EditStep {
    step(EditOperation::ReplaceText {
        target,
        text: text.to_owned(),
    })
}

fn insert(target: TextTarget, at: TargetEdge, text: &str) -> EditStep {
    step(EditOperation::InsertText {
        target,
        at,
        text: text.to_owned(),
    })
}

fn delete(target: TextTarget) -> EditStep {
    step(EditOperation::DeleteText { target })
}

fn new_paragraph(text: &str, style: Option<&str>) -> ParagraphInput {
    ParagraphInput {
        text: text.to_owned(),
        style_id: style.map(str::to_owned),
    }
}

fn insert_paragraphs(id: &str, at: TargetEdge, paragraphs: Vec<ParagraphInput>) -> EditStep {
    step(EditOperation::InsertParagraphs {
        target: body(id),
        at,
        paragraphs,
    })
}

fn delete_paragraphs(first: &str, last: &str) -> EditStep {
    step(EditOperation::DeleteParagraphs {
        story: "body".to_owned(),
        first_para_id: first.to_owned(),
        last_para_id: last.to_owned(),
    })
}

fn set_style(id: &str, style: &str) -> EditStep {
    step(EditOperation::SetParagraphStyle {
        target: body(id),
        style_id: style.to_owned(),
    })
}

fn suggested(mut step: EditStep, author: &str) -> EditStep {
    step.suggest = Some(EditSuggestion {
        author: author.to_owned(),
        date: DATE.to_owned(),
    });
    step
}

fn guarded(mut step: EditStep, text: &str) -> EditStep {
    step.expect = Some(EditGuard {
        text: text.to_owned(),
    });
    step
}

fn request(doc: &EditingDoc, steps: Vec<EditStep>) -> EditRequest {
    EditRequest {
        expect_version: doc.version(),
        source: EditSource::Host,
        history: EditHistory::Separate,
        steps,
    }
}

fn apply_request(doc: &EditingDoc, undo: &UndoSession, request: &EditRequest) -> EditApplication {
    doc.apply_edits(request, undo)
        .unwrap()
        .unwrap_or_else(|refusal| panic!("batch refused: {refusal:?}"))
}

fn apply(doc: &EditingDoc, undo: &UndoSession, steps: Vec<EditStep>) -> EditApplication {
    apply_request(doc, undo, &request(doc, steps))
}

fn refuse(doc: &EditingDoc, undo: &UndoSession, steps: Vec<EditStep>) -> EditRefusal {
    match doc.apply_edits(&request(doc, steps), undo).unwrap() {
        Ok(applied) => panic!("batch applied: {applied:?}"),
        Err(refusal) => refusal,
    }
}

fn code(doc: &EditingDoc, steps: Vec<EditStep>) -> EditFailureCode {
    refuse(doc, &UndoSession::new(), steps).failure.code
}

fn texts(doc: &EditingDoc, story: &str, view: EditTextView) -> Vec<String> {
    doc.read_paragraphs(&ReadParagraphsRequest {
        story: Some(story.to_owned()),
        para_ids: None,
        view,
    })
    .unwrap()
    .paragraphs
    .into_iter()
    .map(|paragraph| paragraph.text)
    .collect()
}

fn accepted(doc: &EditingDoc) -> Vec<String> {
    texts(doc, "body", EditTextView::Accepted)
}

fn para_ids(doc: &EditingDoc) -> Vec<String> {
    doc.paragraphs("body")
        .unwrap()
        .into_iter()
        .map(|paragraph| paragraph.para_id)
        .collect()
}

fn properties(doc: &EditingDoc, id: &str) -> std::collections::BTreeMap<String, Any> {
    doc.paragraphs("body")
        .unwrap()
        .into_iter()
        .find(|paragraph| paragraph.para_id == id)
        .unwrap()
        .properties
}

/// Formatting attributes of the text segment containing `marker`.
fn marks(doc: &EditingDoc, story: &str, marker: &str) -> std::collections::BTreeMap<String, Any> {
    doc.story_segments(story)
        .unwrap()
        .into_iter()
        .find_map(|segment| match segment.content {
            SegmentContent::Text(text) if text.contains(marker) => Some(segment.attributes),
            _ => None,
        })
        .unwrap_or_else(|| panic!("no text segment holds {marker:?}"))
}

fn number(value: Option<&Any>) -> Option<f64> {
    match value? {
        Any::Number(value) => Some(*value),
        Any::BigInt(value) => Some(*value as f64),
        _ => None,
    }
}

fn active(map: &std::collections::BTreeMap<String, Any>, key: &str) -> bool {
    map.get(key).is_some_and(|value| *value != Any::Null)
}

/// Counts committed update events on `doc`.
fn notifications(doc: &EditingDoc) -> (Rc<Cell<u32>>, yrs::Subscription) {
    let count = Rc::new(Cell::new(0));
    let observed = Rc::clone(&count);
    let subscription = doc
        .yrs_doc()
        .observe_update_v1(move |_, _| observed.set(observed.get() + 1))
        .unwrap();
    (count, subscription)
}

/// An undo session whose clock only moves when the test advances it.
fn stepped_undo() -> (UndoSession, Arc<AtomicU64>) {
    let now = Arc::new(AtomicU64::new(1_000));
    let clock = Arc::clone(&now);
    (
        UndoSession::with_clock(Arc::new(move || clock.load(Ordering::Relaxed))),
        now,
    )
}

fn type_text(doc: &EditingDoc, undo: &UndoSession, id: &str, text: &str) {
    undo.track(doc);
    let position = doc.paragraph_mark_position(id).unwrap();
    doc.insert_text(
        &EditCtx::local("", ""),
        position,
        text,
        FormatPolicy::Inherit,
    )
    .unwrap();
}

/// The next id a document mints, observed through a throwaway story.
fn next_minted_id(doc: &EditingDoc, story: &str) -> String {
    doc.create_story(story, "", "Normal", "left").unwrap()
}

#[test]
fn reads_project_atoms_views_and_versions() {
    let doc = open(&format!(
        "{}{}",
        p(
            "00000001",
            &format!(
                "{}{BREAK}{}{IMAGE}{INLINE_CONTROL}{NOTE_REFERENCE}{TAB}{}",
                r("Line one"),
                r("two"),
                r("end")
            )
        ),
        p(
            "00000002",
            &format!(
                r#"{}<w:ins w:id="1" w:author="Ann" w:date="{DATE}">{}</w:ins><w:del w:id="2" w:author="Ann" w:date="{DATE}"><w:r><w:delText>gone</w:delText></w:r></w:del>{}"#,
                r("keep "),
                r("added"),
                r(" end")
            )
        ),
    ));
    let read = doc
        .read_paragraphs(&ReadParagraphsRequest {
            story: None,
            para_ids: Some(vec!["00000001".to_owned()]),
            view: EditTextView::Accepted,
        })
        .unwrap();
    assert_eq!(read.version, doc.version());
    let paragraph = &read.paragraphs[0];
    assert_eq!(
        paragraph.text,
        "Line one\u{FFFC}two\u{FFFC}\u{FFFC}\u{FFFC}\tend"
    );
    assert_eq!(
        paragraph
            .atoms
            .iter()
            .map(|atom| (atom.offset, atom.kind))
            .collect::<Vec<_>>(),
        [
            (8, AtomKind::LineBreak),
            (12, AtomKind::Image),
            (13, AtomKind::ContentControl),
            (14, AtomKind::NoteReference),
        ]
    );
    assert_eq!(
        texts(&doc, "body", EditTextView::Accepted)[1],
        "keep added end"
    );
    assert_eq!(
        texts(&doc, "body", EditTextView::Original)[1],
        "keep gone end"
    );
    assert_eq!(
        texts(&doc, "hf:rIdHeader", EditTextView::Accepted),
        ["Header text"]
    );
    let missing = doc
        .read_paragraphs(&ReadParagraphsRequest {
            story: None,
            para_ids: Some(vec!["nope".to_owned()]),
            view: EditTextView::Accepted,
        })
        .unwrap_err();
    assert_eq!(missing.failure.code, EditFailureCode::MissingTarget);
    assert_eq!(missing.version, doc.version());
}

#[test]
fn find_text_counts_overlapping_matches_and_truncates() {
    let doc = open(&p("00000001", &r("aaaa")));
    let find = |limit| {
        doc.find_text(&FindTextRequest {
            text: "aa".to_owned(),
            within: SearchScope::Story {
                story: "body".to_owned(),
            },
            view: EditTextView::Accepted,
            limit,
        })
        .unwrap()
    };
    let all = find(None);
    assert_eq!(
        all.matches
            .iter()
            .map(|found| found.range.start.offset)
            .collect::<Vec<_>>(),
        [0, 1, 2]
    );
    assert!(!all.truncated);
    let limited = find(Some(2));
    assert_eq!(limited.matches.len(), 2);
    assert!(limited.truncated);
    assert_eq!(
        code(&doc, vec![replace(search("aa", "00000001"), "b")]),
        EditFailureCode::AmbiguousTarget
    );
}

#[test]
fn legacy_offset_bug_cases_resolve_after_every_atom_kind() {
    for atom in [BREAK, IMAGE, INLINE_CONTROL, NOTE_REFERENCE] {
        let doc = open(&p(
            "00000001",
            &format!("{}{atom}{}", r("lead "), r("target tail")),
        ));
        let undo = UndoSession::new();
        let applied = apply(
            &doc,
            &undo,
            vec![replace(search("target", "00000001"), "chosen")],
        );
        assert_eq!(accepted(&doc)[0], "lead \u{FFFC}chosen tail", "{atom}");
        let range = applied.receipts[0].range.clone().unwrap();
        assert_eq!((range.start.offset, range.end.offset), (6, 12), "{atom}");
        let span = doc
            .resolve_text_span(&search("tail", "00000001"))
            .unwrap()
            .unwrap();
        assert_eq!(span.end - span.start, 4, "{atom}");
        let loc = doc.loc_range_of(&span).unwrap();
        assert_eq!((loc.start.offset, loc.end.offset), (13, 17), "{atom}");
        let segments: String = doc
            .story_segments("body")
            .unwrap()
            .into_iter()
            .filter_map(|segment| match segment.content {
                SegmentContent::Text(text) => Some(text),
                _ => None,
            })
            .collect();
        assert_eq!(segments, "lead chosen tail", "{atom}");
    }
}

#[test]
fn two_step_batches_commit_both_or_nothing() {
    let doc = basic();
    let control = basic();
    let undo = UndoSession::new();
    let before = doc.encode_state_as_update_v1();
    let version = doc.version();
    let (events, _subscription) = notifications(&doc);
    let refusal = refuse(
        &doc,
        &undo,
        vec![
            replace(search("beta", "00000001"), "BETA"),
            replace(search("missing", "00000004"), "x"),
        ],
    );
    assert_eq!(refusal.failure.code, EditFailureCode::MissingTarget);
    assert_eq!(refusal.failure.step_index, Some(1));
    assert_eq!(refusal.version, version);
    assert_eq!(doc.version(), version);
    assert_eq!(doc.encode_state_as_update_v1(), before);
    assert_eq!(events.get(), 0);
    assert!(!undo.can_undo());
    assert_eq!(
        next_minted_id(&doc, "probe"),
        next_minted_id(&control, "probe")
    );

    let doc = basic();
    let base = doc.version();
    let (events, _subscription) = notifications(&doc);
    let applied = apply(
        &doc,
        &undo,
        vec![
            replace(search("beta", "00000001"), "BETA"),
            replace(search("Delta", "00000004"), "Epsilon"),
        ],
    );
    assert!(applied.applied);
    assert_eq!(applied.base_version, base);
    assert_eq!(events.get(), 1);
    assert_eq!(applied.version, doc.version());
    assert_ne!(applied.version, applied.base_version);
    assert_eq!(applied.changed_stories, ["body"]);
    assert_eq!(accepted(&doc)[0], "Alpha BETA gamma");
    assert_eq!(accepted(&doc)[3], "Epsilon");
    assert!(undo.undo());
    assert_eq!(accepted(&doc)[0], "Alpha beta gamma");
    assert_eq!(accepted(&doc)[3], "Delta");
}

#[test]
fn every_committed_change_and_reopen_invalidates_versions() {
    let doc = basic();
    let undo = UndoSession::new();
    let initial = doc.version();
    undo.track(&doc);
    doc.delete_range(&EditCtx::local("", ""), StoryRange::new("body", 0, 1))
        .unwrap();
    let after_delete = doc.version();
    assert_ne!(after_delete, initial);
    assert!(undo.undo());
    assert_ne!(doc.version(), after_delete);
    let after_undo = doc.version();
    assert!(undo.redo());
    assert_ne!(doc.version(), after_undo);

    let peer = EditingDoc::new(7002);
    peer.apply_update_v1(&doc.encode_state_as_update_v1())
        .unwrap();
    let before_remote = doc.version();
    peer.insert_text(
        &EditCtx::local("", ""),
        Position::new("body", 0),
        "remote ",
        FormatPolicy::Plain,
    )
    .unwrap();
    doc.apply_update_v1(&peer.encode_state_as_update_v1())
        .unwrap();
    assert_ne!(doc.version(), before_remote);
    let stale = request(&doc, vec![replace(search("remote", "00000001"), "x")]);
    peer.insert_text(
        &EditCtx::local("", ""),
        Position::new("body", 0),
        "again ",
        FormatPolicy::Plain,
    )
    .unwrap();
    doc.apply_update_v1(&peer.encode_state_as_update_v1())
        .unwrap();
    let refusal = doc.apply_edits(&stale, &undo).unwrap().unwrap_err();
    assert_eq!(refusal.failure.code, EditFailureCode::StaleVersion);

    let first = basic();
    let reopened = basic();
    assert_ne!(first.version(), reopened.version());
    let from_first = request(&first, vec![replace(search("beta", "00000001"), "x")]);
    assert_eq!(
        reopened
            .apply_edits(&from_first, &UndoSession::new())
            .unwrap()
            .unwrap_err()
            .failure
            .code,
        EditFailureCode::StaleVersion
    );
    let version = DocumentVersion::from(reopened.version().to_string());
    assert_eq!(version, reopened.version());
}

#[test]
fn no_op_batches_create_no_history_notification_or_version() {
    let doc = basic();
    let undo = UndoSession::new();
    type_text(&doc, &undo, "00000004", "!");
    assert!(undo.undo());
    assert!(undo.can_redo());
    let version = doc.version();
    let (events, _subscription) = notifications(&doc);
    for steps in [
        vec![],
        vec![insert(search("beta", "00000001"), TargetEdge::End, "")],
        vec![delete(range("00000001", 3, 3, EditTextView::Accepted))],
        vec![replace(search("beta", "00000001"), "beta")],
        vec![set_style("00000003", "Heading1")],
        vec![insert_paragraphs("00000001", TargetEdge::End, vec![])],
    ] {
        let applied = apply(&doc, &undo, steps);
        assert!(!applied.applied);
        assert_eq!(applied.version, version);
        assert!(applied.changed_stories.is_empty());
        assert!(applied.receipts.iter().all(|receipt| !receipt.changed));
    }
    assert_eq!(events.get(), 0);
    assert_eq!(doc.version(), version);
    assert!(undo.can_redo());
    assert!(!undo.can_undo());
}

#[test]
fn a_no_op_step_still_checks_version_targets_and_guards() {
    let doc = basic();
    let undo = UndoSession::new();
    let mut stale = request(&doc, vec![replace(search("beta", "00000001"), "beta")]);
    stale.expect_version = DocumentVersion::from("0-0");
    assert_eq!(
        doc.apply_edits(&stale, &undo)
            .unwrap()
            .unwrap_err()
            .failure
            .code,
        EditFailureCode::StaleVersion
    );
    assert_eq!(
        code(
            &doc,
            vec![replace(search("nothing", "00000001"), "nothing")]
        ),
        EditFailureCode::MissingTarget
    );
    assert_eq!(
        code(
            &doc,
            vec![guarded(replace(search("beta", "00000001"), "beta"), "BETA")]
        ),
        EditFailureCode::ContentMismatch
    );
}

#[test]
fn one_undo_step_between_typing_in_both_capture_modes() {
    for mode in [UndoCaptureMode::Auto, UndoCaptureMode::Manual] {
        let doc = basic();
        let (undo, now) = stepped_undo();
        undo.set_capture_mode(mode);
        type_text(&doc, &undo, "00000001", "a");
        now.fetch_add(10, Ordering::Relaxed);
        apply(
            &doc,
            &undo,
            vec![
                replace(search("Delta", "00000004"), "Omega"),
                step(EditOperation::ReplaceText {
                    target: TextTarget::Paragraph(para("hf:rIdHeader", "0000E001")),
                    text: "New header".to_owned(),
                }),
            ],
        );
        now.fetch_add(10, Ordering::Relaxed);
        type_text(&doc, &undo, "00000001", "b");
        assert!(undo.undo(), "{mode:?}");
        assert_eq!(accepted(&doc)[0], "Alpha beta gammaa", "{mode:?}");
        assert_eq!(accepted(&doc)[3], "Omega", "{mode:?}");
        assert!(undo.undo(), "{mode:?}");
        assert_eq!(accepted(&doc)[3], "Delta", "{mode:?}");
        assert_eq!(
            texts(&doc, "hf:rIdHeader", EditTextView::Accepted),
            ["Header text"],
            "{mode:?}"
        );
        assert_eq!(undo.changed_stories(), ["body", "hf:rIdHeader"], "{mode:?}");
        assert_eq!(accepted(&doc)[0], "Alpha beta gammaa", "{mode:?}");
        assert!(undo.undo(), "{mode:?}");
        assert_eq!(accepted(&doc)[0], "Alpha beta gamma", "{mode:?}");
        assert!(!undo.undo(), "{mode:?}");
    }
}

#[test]
fn history_and_source_are_independent() {
    let doc = basic();
    let undo = UndoSession::new();
    let mut agent = request(&doc, vec![replace(search("beta", "00000001"), "BETA")]);
    agent.source = EditSource::Agent;
    let applied = apply_request(&doc, &undo, &agent);
    assert_eq!(applied.source, EditSource::Agent);
    assert!(undo.can_undo());
    assert!(undo.undo());
    assert!(undo.can_redo());

    let mut untracked = request(&doc, vec![replace(search("Delta", "00000004"), "Omega")]);
    untracked.history = EditHistory::None;
    let applied = apply_request(&doc, &undo, &untracked);
    assert_eq!(applied.source, EditSource::Host);
    assert!(undo.can_redo());
    assert!(!undo.can_undo());
    assert!(undo.redo());
    assert_eq!(accepted(&doc)[0], "Alpha BETA gamma");
    assert_eq!(accepted(&doc)[3], "Omega");

    let fresh = basic();
    let history = UndoSession::new();
    let mut none = request(&fresh, vec![replace(search("beta", "00000001"), "x")]);
    none.history = EditHistory::None;
    apply_request(&fresh, &history, &none);
    history.track(&fresh);
    assert!(!history.can_undo());
    let foreign = basic();
    assert!(
        foreign
            .apply_edits(
                &request(&foreign, vec![replace(search("beta", "00000001"), "y")]),
                &history
            )
            .is_err()
    );
}

#[test]
fn targets_fail_as_typed_data() {
    let doc = open(&format!(
        "{}{}",
        p("000000DD", &r("first")),
        p("00000003", &r("plain text"))
    ));
    doc.apply_raw_ops(
        "body",
        vec![RawOp::InsertEmbed {
            index: 0,
            kind: "pilcrow".to_owned(),
            payload: vec![("paraId".to_owned(), Any::from("000000DD"))],
            attrs: Default::default(),
        }],
        &EditCtx::local("", ""),
    )
    .unwrap();
    assert_eq!(
        code(&doc, vec![replace(search("x", "missing"), "y")]),
        EditFailureCode::MissingTarget
    );
    assert_eq!(
        code(
            &doc,
            vec![replace(TextTarget::Paragraph(body("000000DD")), "y")]
        ),
        EditFailureCode::AmbiguousTarget
    );
    let refusal = refuse(
        &doc,
        &UndoSession::new(),
        vec![guarded(
            replace(search("plain", "00000003"), "PLAIN"),
            "plan",
        )],
    );
    assert_eq!(refusal.failure.code, EditFailureCode::ContentMismatch);
    assert_eq!(refusal.failure.step_index, Some(0));
    apply(
        &doc,
        &UndoSession::new(),
        vec![guarded(
            replace(search("plain", "00000003"), "PLAIN"),
            "plain",
        )],
    );
    for target in [
        range("00000003", 4, 2, EditTextView::Accepted),
        range("00000003", 0, 99, EditTextView::Accepted),
    ] {
        assert_eq!(
            code(&doc, vec![delete(target)]),
            EditFailureCode::InvalidStep
        );
    }
    assert_eq!(
        code(&doc, vec![replace(search("", "00000003"), "y")]),
        EditFailureCode::InvalidStep
    );
    assert_eq!(
        code(
            &doc,
            vec![insert(search("PLAIN", "00000003"), TargetEdge::End, "a\nb")]
        ),
        EditFailureCode::InvalidStep
    );
    let cross = TextTarget::Range(TextRange {
        story: "body".to_owned(),
        start: TextPosition {
            para_id: "000000DD".to_owned(),
            offset: 0,
        },
        end: TextPosition {
            para_id: "00000003".to_owned(),
            offset: 1,
        },
        view: EditTextView::Accepted,
    });
    assert_eq!(
        code(&doc, vec![delete(cross)]),
        EditFailureCode::Unsupported
    );
}

#[test]
fn every_step_targets_the_pre_batch_state() {
    let doc = basic();
    let undo = UndoSession::new();
    let applied = apply(
        &doc,
        &undo,
        vec![
            replace(range("00000001", 11, 16, EditTextView::Accepted), "GAMMA!"),
            replace(range("00000001", 0, 5, EditTextView::Accepted), "A"),
            replace(range("00000001", 5, 6, EditTextView::Accepted), "_"),
            insert(search("Delta", "00000004"), TargetEdge::Start, ">"),
        ],
    );
    assert_eq!(accepted(&doc)[0], "A_beta GAMMA!");
    assert_eq!(accepted(&doc)[3], ">Delta");
    let offsets: Vec<(u32, u32)> = applied
        .receipts
        .iter()
        .map(|receipt| {
            let range = receipt.range.as_ref().unwrap();
            (range.start.offset, range.end.offset)
        })
        .collect();
    assert_eq!(offsets, [(7, 13), (0, 1), (1, 2), (0, 1)]);
    assert!(
        applied
            .receipts
            .iter()
            .all(|receipt| receipt.range.as_ref().unwrap().view == EditTextView::Accepted)
    );
}

#[test]
fn conflicting_effects_are_refused() {
    let doc = basic();
    let accepted_range = |start, end| range("00000001", start, end, EditTextView::Accepted);
    let cases = [
        vec![
            replace(accepted_range(0, 6), "x"),
            replace(accepted_range(5, 10), "y"),
        ],
        vec![
            insert(accepted_range(5, 5), TargetEdge::Start, "x"),
            insert(accepted_range(5, 5), TargetEdge::End, "y"),
        ],
        vec![
            replace(accepted_range(0, 5), "x"),
            insert(accepted_range(5, 5), TargetEdge::Start, "y"),
        ],
        vec![
            set_style("00000001", "Quote"),
            replace(accepted_range(0, 5), "x"),
        ],
        vec![
            insert_paragraphs("00000001", TargetEdge::End, vec![new_paragraph("a", None)]),
            insert_paragraphs(
                "00000002",
                TargetEdge::Start,
                vec![new_paragraph("b", None)],
            ),
        ],
        vec![
            delete_paragraphs("00000003", "00000004"),
            replace(search("Delta", "00000004"), "x"),
        ],
        vec![
            delete_paragraphs("00000003", "00000003"),
            insert_paragraphs("00000003", TargetEdge::End, vec![new_paragraph("a", None)]),
        ],
        vec![
            set_style("00000003", "Quote"),
            set_style("00000003", "Normal"),
        ],
    ];
    for steps in cases {
        let refusal = refuse(&doc, &UndoSession::new(), steps);
        assert_eq!(refusal.failure.code, EditFailureCode::OverlappingSteps);
        assert_eq!(refusal.failure.step_index, Some(1));
        assert_eq!(refusal.failure.conflicting_step_index, Some(0));
    }
    apply(
        &basic(),
        &UndoSession::new(),
        vec![
            set_style("00000001", "Quote"),
            insert_paragraphs("00000001", TargetEdge::End, vec![new_paragraph("a", None)]),
            replace(search("target", "00000002"), "x"),
        ],
    );
}

#[test]
fn locks_follow_actual_control_ownership() {
    let doc = open(&format!(
        r#"{}<w:sdt><w:sdtPr><w:lock w:val="contentLocked"/><w:tag w:val="locked"/></w:sdtPr><w:sdtContent>{}<w:tbl><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid><w:tr><w:tc>{}</w:tc></w:tr></w:tbl>{}</w:sdtContent></w:sdt><w:sdt><w:sdtPr><w:lock w:val="sdtLocked"/></w:sdtPr><w:sdtContent>{}</w:sdtContent></w:sdt>{}"#,
        p("00000001", &r("outside")),
        p("0000A001", &r("locked text")),
        p("0000C001", &r("cell text")),
        p("0000A002", &r("after table")),
        p("0000B001", &r("removal lock only")),
        p("00000009", &r("tail")),
    ));
    let locked = |story: &str, id: &str| replace(TextTarget::Paragraph(para(story, id)), "changed");
    assert_eq!(
        code(&doc, vec![locked("body:sdt0", "0000A001")]),
        EditFailureCode::LockedTarget
    );
    assert_eq!(
        code(&doc, vec![locked("body:sdt0:t0:r0c0", "0000C001")]),
        EditFailureCode::LockedTarget
    );
    let undo = UndoSession::new();
    apply(&doc, &undo, vec![locked("body:sdt1", "0000B001")]);
    assert_eq!(
        texts(&doc, "body:sdt1", EditTextView::Accepted),
        ["changed"]
    );
    assert_eq!(
        texts(&doc, "body:sdt0", EditTextView::Accepted),
        ["locked text", "after table"]
    );
}

#[test]
fn existing_revisions_conflict_only_where_touched() {
    let doc = open(&format!(
        "{}{}",
        p(
            "00000001",
            &format!(
                r#"{}<w:ins w:id="1" w:author="Ann" w:date="{DATE}">{}</w:ins>{}<w:del w:id="2" w:author="Ann" w:date="{DATE}"><w:r><w:delText>gone</w:delText></w:r></w:del>{}"#,
                r("keep "),
                r("added"),
                r(" mid "),
                r(" end")
            )
        ),
        p("00000002", &r("clean")),
    ));
    assert_eq!(accepted(&doc)[0], "keep added mid  end");
    for steps in [
        vec![replace(search("added", "00000001"), "x")],
        vec![delete(range("00000001", 13, 16, EditTextView::Accepted))],
        vec![insert(
            range("00000001", 15, 15, EditTextView::Accepted),
            TargetEdge::Start,
            "x",
        )],
        vec![replace(
            search("gone", "00000001").in_view(EditTextView::Original),
            "x",
        )],
        vec![set_style("00000001", "Quote")],
        vec![delete_paragraphs("00000001", "00000001")],
    ] {
        assert_eq!(code(&doc, steps), EditFailureCode::TrackedRevisionConflict);
    }
    let undo = UndoSession::new();
    apply(
        &doc,
        &undo,
        vec![
            replace(search("mid", "00000001"), "MID"),
            replace(search("clean", "00000002"), "CLEAN"),
        ],
    );
    assert_eq!(accepted(&doc), ["keep added MID  end", "CLEAN"]);
}

trait InView {
    fn in_view(self, view: EditTextView) -> Self;
}

impl InView for TextTarget {
    fn in_view(self, view: EditTextView) -> Self {
        match self {
            TextTarget::Search { text, within, .. } => TextTarget::Search { text, within, view },
            other => other,
        }
    }
}

#[test]
fn paragraph_mark_revisions_refuse_reads_and_writes_of_affected_paragraphs() {
    let doc = open(&format!(
        r#"{}<w:p w14:paraId="00000002"><w:pPr><w:rPr><w:ins w:id="3" w:author="Ann" w:date="{DATE}"/></w:rPr></w:pPr>{}</w:p>{}"#,
        p("00000001", &r("before")),
        r("split here"),
        p("00000003", &r("after split")),
    ));
    let read = |ids: Option<Vec<&str>>| {
        doc.read_paragraphs(&ReadParagraphsRequest {
            story: None,
            para_ids: ids.map(|ids| ids.into_iter().map(str::to_owned).collect()),
            view: EditTextView::Accepted,
        })
        .map(|read| read.paragraphs.len())
        .map_err(|refusal| refusal.failure.code)
    };
    assert_eq!(read(None), Err(EditFailureCode::Unsupported));
    assert_eq!(read(Some(vec!["00000001"])), Ok(1));
    assert_eq!(
        code(&doc, vec![replace(search("after", "00000003"), "x")]),
        EditFailureCode::TrackedRevisionConflict
    );
    apply(
        &doc,
        &UndoSession::new(),
        vec![replace(search("before", "00000001"), "BEFORE")],
    );
}

#[test]
fn offsets_are_unicode_scalar_boundaries() {
    let doc = open(&p("00000001", &r("a😀be\u{0301}x")));
    assert_eq!(
        code(
            &doc,
            vec![delete(range("00000001", 2, 3, EditTextView::Accepted))]
        ),
        EditFailureCode::InvalidStep
    );
    let undo = UndoSession::new();
    apply(
        &doc,
        &undo,
        vec![insert(
            range("00000001", 5, 5, EditTextView::Accepted),
            TargetEdge::Start,
            "|",
        )],
    );
    assert_eq!(accepted(&doc)[0], "a😀be|\u{0301}x");
}

#[test]
fn inline_atoms_cannot_be_replaced_but_text_around_them_can() {
    let doc = open(&p(
        "00000001",
        &format!("{}{IMAGE}{}", r("left"), r("right")),
    ));
    assert_eq!(
        code(
            &doc,
            vec![delete(range("00000001", 3, 6, EditTextView::Accepted))]
        ),
        EditFailureCode::Unsupported
    );
    apply(
        &doc,
        &UndoSession::new(),
        vec![
            replace(range("00000001", 0, 4, EditTextView::Accepted), "L"),
            replace(range("00000001", 5, 10, EditTextView::Accepted), "R"),
        ],
    );
    assert_eq!(accepted(&doc)[0], "L\u{FFFC}R");
}

#[test]
fn inserted_paragraphs_keep_existing_identities_and_use_style_defaults() {
    let doc = basic();
    let before = para_ids(&doc);
    let anchor = properties(&doc, "00000003");
    let applied = apply(
        &doc,
        &UndoSession::new(),
        vec![
            insert_paragraphs(
                "00000003",
                TargetEdge::End,
                vec![
                    new_paragraph("Same style", None),
                    new_paragraph("Quoted", Some("Quote")),
                ],
            ),
            insert_paragraphs(
                "00000001",
                TargetEdge::Start,
                vec![new_paragraph("First", None)],
            ),
        ],
    );
    assert_eq!(
        accepted(&doc),
        [
            "First",
            "Alpha beta gamma",
            "Line one\u{FFFC}target words",
            "Title",
            "Same style",
            "Quoted",
            "Delta"
        ]
    );
    let after = para_ids(&doc);
    assert!(before.iter().all(|id| after.contains(id)));
    assert_eq!(properties(&doc, "00000003"), anchor);
    let created = &applied.receipts[0].new_paragraphs;
    assert_eq!(created.len(), 2);
    let range = applied.receipts[0].range.as_ref().unwrap();
    assert_eq!(range.start.para_id, created[0].para_id);
    assert_eq!(range.end.para_id, created[1].para_id);
    assert_eq!(range.end.offset, 6);
    let same = properties(&doc, &created[0].para_id);
    assert_eq!(same.get("pStyle"), Some(&Any::from("Heading1")));
    assert_eq!(same.get("keepNext"), Some(&Any::Bool(true)));
    assert!(!same.contains_key("bookmarks"));
    let quoted = properties(&doc, &created[1].para_id);
    assert_eq!(quoted.get("pStyle"), Some(&Any::from("Quote")));
    assert_eq!(quoted.get("alignment"), Some(&Any::from("center")));
    assert!(active(&marks(&doc, "body", "Same style"), "bold"));
    assert!(active(&marks(&doc, "body", "Quoted"), "italic"));
    assert!(!active(&marks(&doc, "body", "Quoted"), "bold"));
    let minted = doc
        .split_paragraph(&EditCtx::local("", ""), Position::new("body", 1), None)
        .unwrap();
    assert!(
        !created
            .iter()
            .any(|target| target.para_id == minted.second_para_id)
    );
}

#[test]
fn deleted_paragraphs_donate_nothing_to_their_neighbours() {
    let doc = basic();
    let survivor = properties(&doc, "00000004");
    let applied = apply(
        &doc,
        &UndoSession::new(),
        vec![delete_paragraphs("00000002", "00000003")],
    );
    assert_eq!(accepted(&doc), ["Alpha beta gamma", "Delta"]);
    assert_eq!(properties(&doc, "00000004"), survivor);
    assert_eq!(
        applied.receipts[0]
            .removed_paragraphs
            .iter()
            .map(|target| target.para_id.as_str())
            .collect::<Vec<_>>(),
        ["00000002", "00000003"]
    );
    let range = applied.receipts[0].range.as_ref().unwrap();
    assert_eq!(
        (range.start.para_id.as_str(), range.start.offset),
        ("00000004", 0)
    );
    let tail = basic();
    apply(
        &tail,
        &UndoSession::new(),
        vec![delete_paragraphs("00000003", "00000004")],
    );
    assert_eq!(
        accepted(&tail),
        ["Alpha beta gamma", "Line one\u{FFFC}target words"]
    );
}

#[test]
fn structural_deletions_refuse_unsafe_spans() {
    let doc = open(&format!(
        r#"{}{}{RAW}{}{}<w:tbl><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid><w:tr><w:tc>{}</w:tc></w:tr></w:tbl>{}{}{}{}{}"#,
        p("00000001", &r("one")),
        p("00000002", &r("two")),
        p("00000003", &r("three")),
        p("00000004", &r("four")),
        p("0000C001", &r("cell")),
        p("00000005", &r("five")),
        r#"<w:p w14:paraId="00000006"><w:pPr><w:sectPr/></w:pPr><w:r><w:t>section</w:t></w:r></w:p>"#,
        p(
            "00000007",
            &format!(r#"<w:bookmarkStart w:id="5" w:name="mark"/>{}"#, r("seven"))
        ),
        p(
            "00000008",
            &format!(r#"{}<w:bookmarkEnd w:id="5"/>"#, r("eight"))
        ),
        p(
            "00000009",
            &format!(
                r#"<w:commentRangeStart w:id="9"/>{}<w:commentRangeEnd w:id="9"/><w:r><w:commentReference w:id="9"/></w:r>"#,
                r("commented")
            )
        ),
    ));
    for (first, last) in [
        ("00000002", "00000003"),
        ("00000004", "00000005"),
        ("00000006", "00000006"),
        ("00000007", "00000007"),
        ("00000009", "00000009"),
        ("00000001", "00000009"),
    ] {
        assert_eq!(
            code(&doc, vec![delete_paragraphs(first, last)]),
            EditFailureCode::Unsupported,
            "{first}..{last}"
        );
    }
    assert_eq!(
        code(
            &doc,
            vec![insert_paragraphs(
                "00000002",
                TargetEdge::End,
                vec![new_paragraph("x", None)]
            )]
        ),
        EditFailureCode::Unsupported
    );
    assert_eq!(
        code(
            &doc,
            vec![insert_paragraphs(
                "00000003",
                TargetEdge::Start,
                vec![new_paragraph("x", None)]
            )]
        ),
        EditFailureCode::Unsupported
    );
    assert_eq!(
        code(&doc, vec![delete_paragraphs("00000003", "00000001")]),
        EditFailureCode::InvalidStep
    );
    let undo = UndoSession::new();
    apply(&doc, &undo, vec![delete_paragraphs("00000007", "00000008")]);
    apply(&doc, &undo, vec![delete_paragraphs("00000001", "00000001")]);
    assert_eq!(accepted(&doc)[0], "two");
    let only = open(&p("00000001", &r("only")));
    assert_eq!(
        code(&only, vec![delete_paragraphs("00000001", "00000001")]),
        EditFailureCode::Unsupported
    );
}

#[test]
fn paragraph_styles_apply_their_full_effect_in_rust() {
    let doc = open(&format!(
        r#"<w:p w14:paraId="00000001"><w:pPr><w:pStyle w:val="Heading1"/><w:ind w:left="1440"/></w:pPr>{}<w:r><w:rPr><w:u w:val="single"/></w:rPr><w:t>underlined</w:t></w:r></w:p>{}"#,
        r("Heading "),
        p("00000002", &r("plain")),
    ));
    let undo = UndoSession::new();
    apply(&doc, &undo, vec![set_style("00000001", "Quote")]);
    let quoted = properties(&doc, "00000001");
    assert_eq!(quoted.get("pStyle"), Some(&Any::from("Quote")));
    assert_eq!(number(quoted.get("indentLeft")), Some(720.0));
    assert_eq!(quoted.get("alignment"), Some(&Any::from("center")));
    assert!(!quoted.contains_key("keepNext"));
    let Some(Any::Map(original)) = quoted.get("_originalFormatting") else {
        panic!("seeded paragraphs keep their source formatting record");
    };
    assert_eq!(original.get("styleId"), Some(&Any::from("Quote")));
    assert!(!original.contains_key("indentLeft"));
    let heading = marks(&doc, "body", "Heading");
    assert!(!active(&heading, "bold"));
    assert!(!active(&heading, "fontSize"));
    assert!(active(&heading, "italic"));
    assert!(active(&heading, "allCaps"));
    assert!(!active(&marks(&doc, "body", "underlined"), "underline"));
    assert!(undo.undo());
    assert_eq!(
        properties(&doc, "00000001").get("pStyle"),
        Some(&Any::from("Heading1"))
    );

    assert_eq!(
        code(&doc, vec![set_style("00000002", "Missing")]),
        EditFailureCode::InvalidStep
    );
    assert_eq!(
        code(&doc, vec![set_style("00000002", "Strong")]),
        EditFailureCode::InvalidStep
    );
    assert_eq!(
        code(&doc, vec![suggested(set_style("00000002", "Quote"), "Ann")]),
        EditFailureCode::Unsupported
    );
    for steps in [
        vec![set_style("00000002", "ListNumber")],
        vec![insert_paragraphs(
            "00000002",
            TargetEdge::End,
            vec![new_paragraph("x", Some("ListNumber"))],
        )],
    ] {
        let refusal = refuse(&doc, &UndoSession::new(), steps).failure;
        assert_eq!(refusal.code, EditFailureCode::Unsupported);
        assert_eq!(
            refusal.message,
            "style \"ListNumber\" defines list numbering, which v1 batches do not apply; retaining numbering definitions for batches is a follow-up"
        );
    }
}

#[test]
fn structure_and_styles_need_retained_source_metadata() {
    let doc = EditingDoc::new(7003);
    let id = doc.create_story("body", "text", "Normal", "left").unwrap();
    for steps in [
        vec![insert_paragraphs(
            &id,
            TargetEdge::End,
            vec![new_paragraph("x", None)],
        )],
        vec![set_style(&id, "Normal")],
    ] {
        assert_eq!(code(&doc, steps), EditFailureCode::Unsupported);
    }
    apply(
        &doc,
        &UndoSession::new(),
        vec![replace(TextTarget::Paragraph(body(&id)), "changed")],
    );
    assert_eq!(accepted(&doc), ["changed"]);
}

#[test]
fn suggested_text_steps_author_revisions_that_resolve() {
    let doc = basic();
    let undo = UndoSession::new();
    let applied = apply(
        &doc,
        &undo,
        vec![
            suggested(replace(search("beta", "00000001"), "BETA"), "Ann"),
            suggested(delete(search("Delta", "00000004")), "Ann"),
            suggested(
                insert(search("Title", "00000003"), TargetEdge::End, "!"),
                "Ann",
            ),
        ],
    );
    assert!(
        applied
            .receipts
            .iter()
            .all(|receipt| receipt.revision_ids.len() == 1)
    );
    assert_eq!(accepted(&doc)[0], "Alpha BETA gamma");
    assert_eq!(
        texts(&doc, "body", EditTextView::Original)[0],
        "Alpha beta gamma"
    );
    assert_eq!(accepted(&doc)[3], "");
    let replace_revision = applied.receipts[0].revision_ids[0].clone();
    doc.accept_change(
        &EditCtx::local("", ""),
        &ChangeTarget::Revision(replace_revision),
    )
    .unwrap();
    assert_eq!(
        texts(&doc, "body", EditTextView::Original)[0],
        "Alpha BETA gamma"
    );
    let delete_revision = applied.receipts[1].revision_ids[0].clone();
    doc.reject_change(
        &EditCtx::local("", ""),
        &ChangeTarget::Revision(delete_revision),
    )
    .unwrap();
    assert_eq!(accepted(&doc)[3], "Delta");
    let mut missing_author = suggested(replace(search("gamma", "00000001"), "x"), "");
    missing_author.suggest.as_mut().unwrap().date.clear();
    assert_eq!(
        code(&doc, vec![missing_author]),
        EditFailureCode::InvalidStep
    );
    assert_eq!(
        code(
            &doc,
            vec![suggested(
                insert_paragraphs("00000001", TargetEdge::End, vec![new_paragraph("x", None)]),
                "Ann"
            )]
        ),
        EditFailureCode::Unsupported
    );
}

#[test]
fn validation_previews_the_resolved_plan_without_reserving() {
    let doc = basic();
    let steps = vec![
        suggested(replace(search("beta", "00000001"), "BETA"), "Ann"),
        insert_paragraphs(
            "00000001",
            TargetEdge::End,
            vec![new_paragraph("a", None), new_paragraph("b", None)],
        ),
        replace(search("Delta", "00000004"), "Delta"),
    ];
    let before = doc.encode_state_as_update_v1();
    let validation = doc
        .validate_edits(&request(&doc, steps.clone()))
        .unwrap()
        .unwrap();
    assert_eq!(doc.encode_state_as_update_v1(), before);
    assert_eq!(validation.base_version, doc.version());
    assert!(validation.would_apply);
    let previews = &validation.previews;
    assert!(matches!(
        &previews[0].target,
        EditTarget::Range(range) if (range.start.offset, range.end.offset) == (6, 10)
    ));
    assert!(previews[0].would_create_revisions);
    assert_eq!(previews[1].new_paragraph_count, 2);
    assert!(!previews[2].would_change);
    let expected = request(&doc, steps.clone());
    let peer = EditingDoc::new(7004);
    peer.apply_update_v1(&doc.encode_state_as_update_v1())
        .unwrap();
    peer.insert_text(
        &EditCtx::local("", ""),
        Position::new("body", 0),
        "x",
        FormatPolicy::Plain,
    )
    .unwrap();
    doc.apply_update_v1(&peer.encode_state_as_update_v1())
        .unwrap();
    assert_eq!(
        doc.apply_edits(&expected, &UndoSession::new())
            .unwrap()
            .unwrap_err()
            .failure
            .code,
        EditFailureCode::StaleVersion
    );
    apply(&doc, &UndoSession::new(), steps);
}

#[test]
fn budgets_return_limit_exceeded() {
    let doc = basic();
    let many = (0..129)
        .map(|_| insert(search("Delta", "00000004"), TargetEdge::End, "x"))
        .collect();
    assert_eq!(code(&doc, many), EditFailureCode::LimitExceeded);
    let long = "x".repeat(1_048_577);
    assert_eq!(
        code(&doc, vec![replace(search("Delta", "00000004"), &long)]),
        EditFailureCode::LimitExceeded
    );
    let paragraphs = (0..1_025).map(|_| new_paragraph("p", None)).collect();
    assert_eq!(
        code(
            &doc,
            vec![insert_paragraphs("00000004", TargetEdge::End, paragraphs)]
        ),
        EditFailureCode::LimitExceeded
    );
}

#[test]
fn pending_updates_make_batches_unsupported() {
    let doc = basic();
    let peer = EditingDoc::new(7005);
    peer.apply_update_v1(&doc.encode_state_as_update_v1())
        .unwrap();
    peer.insert_text(
        &EditCtx::local("", ""),
        Position::new("body", 0),
        "one ",
        FormatPolicy::Plain,
    )
    .unwrap();
    let known = peer.encode_state_vector_v1();
    peer.insert_text(
        &EditCtx::local("", ""),
        Position::new("body", 4),
        "two ",
        FormatPolicy::Plain,
    )
    .unwrap();
    doc.apply_update_v1(&peer.encode_diff_v1(&known).unwrap())
        .unwrap();
    assert_eq!(
        code(&doc, vec![replace(search("beta", "00000001"), "x")]),
        EditFailureCode::Unsupported
    );
}

#[test]
fn requests_decode_from_json_and_malformed_input_is_an_error() {
    let doc = basic();
    let json = format!(
        r#"{{"expectVersion":"{}","steps":[{{"op":"replaceText","target":{{"kind":"search","text":"beta","within":{{"kind":"paragraph","story":"body","paraId":"00000001"}},"view":"accepted"}},"text":"B","expect":{{"text":"beta"}},"suggest":null}},{{"op":"insertParagraphs","target":{{"story":"body","paraId":"00000004"}},"at":"end","paragraphs":[{{"text":"x","styleId":"Quote"}}]}},{{"op":"deleteParagraphs","story":"body","firstParaId":"00000003","lastParaId":"00000003"}}]}}"#,
        doc.version()
    );
    let parsed: EditRequest = serde_json::from_str(&json).unwrap();
    assert_eq!(parsed.source, EditSource::Host);
    assert_eq!(parsed.history, EditHistory::Separate);
    assert_eq!(parsed.steps.len(), 3);
    assert!(parsed.steps[0].suggest.is_none());
    apply_request(&doc, &UndoSession::new(), &parsed);
    for malformed in [
        r#"{"steps":[]}"#,
        r#"{"expectVersion":"v","steps":[{"op":"moveText"}]}"#,
        r#"{"expectVersion":"v","steps":[],"extra":1}"#,
        r#"{"expectVersion":"v","steps":[{"op":"deleteText","target":{"kind":"range","story":"body","start":{"paraId":"p","offset":-1},"end":{"paraId":"p","offset":1},"view":"accepted"}}]}"#,
        r#"{"expectVersion":"v","steps":[{"op":"deleteText","target":{"kind":"paragraph","story":"body","paraId":"p"},"expect":{"txt":"x"}}]}"#,
        r#"{"expectVersion":"v","steps":[{"op":"deleteText","target":{"kind":"paragraph","story":"body","paraId":"p","view":"accepted"}}]}"#,
        r#"{"expectVersion":"v","history":"sometimes","steps":[]}"#,
        r#"{"expectVersion":"v","steps":[{"op":"deleteText","target":{"kind":"search","text":"x","within":{"kind":"story","story":"body"},"view":"final"}}]}"#,
        r#"{"expectVersion":"v","steps":[{"op":"deleteText","target":{"kind":"range","story":"body","start":{"paraId":"p","offset":1.5},"end":{"paraId":"p","offset":2},"view":"accepted"}}]}"#,
    ] {
        assert!(
            serde_json::from_str::<EditRequest>(malformed).is_err(),
            "{malformed}"
        );
    }
    let refusal = serde_json::to_value(refuse(
        &doc,
        &UndoSession::new(),
        vec![replace(search("zzz", "00000001"), "x")],
    ))
    .unwrap();
    assert_eq!(refusal["failure"]["code"], "missing-target");
    assert_eq!(refusal["failure"]["stepIndex"], 0);
    assert_eq!(refusal["failure"]["target"]["kind"], "search");
}

#[test]
fn deleting_the_story_tail_keeps_a_final_paragraph_mark() {
    let table = format!(
        r#"<w:tbl><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid><w:tr><w:tc>{}</w:tc></w:tr></w:tbl>"#,
        p("0000C001", &r("cell"))
    );
    let doc = open(&format!(
        "{}{table}{}",
        p("00000001", &r("head")),
        p("00000002", &r("after table"))
    ));
    assert_eq!(
        code(&doc, vec![delete_paragraphs("00000002", "00000002")]),
        EditFailureCode::Unsupported
    );
    let trailing = open(&format!("{}{table}", p("00000001", &r("head"))));
    let tail = para_ids(&trailing).pop().unwrap();
    assert_eq!(
        code(&trailing, vec![delete_paragraphs(&tail, &tail)]),
        EditFailureCode::Unsupported
    );
    apply(
        &doc,
        &UndoSession::new(),
        vec![delete_paragraphs("00000001", "00000001")],
    );
    assert_eq!(accepted(&doc), ["after table"]);
}

#[test]
fn opaque_blocks_follow_the_live_paragraphs_they_are_restored_before() {
    let doc = open(&format!(
        "{}{RAW}{}{}",
        p("00000001", &r("A")),
        p("00000002", &r("B")),
        p("00000003", &r("C"))
    ));
    let undo = UndoSession::new();
    let add = |id: &str, at| insert_paragraphs(id, at, vec![new_paragraph("new", None)]);
    assert_eq!(
        code(&doc, vec![add("00000001", TargetEdge::End)]),
        EditFailureCode::Unsupported
    );
    apply(&doc, &undo, vec![add("00000003", TargetEdge::Start)]);
    assert_eq!(
        code(&doc, vec![delete_paragraphs("00000002", "00000002")]),
        EditFailureCode::Unsupported
    );
    assert!(undo.undo());
    apply(&doc, &undo, vec![delete_paragraphs("00000002", "00000002")]);
    for steps in [
        vec![add("00000003", TargetEdge::Start)],
        vec![add("00000001", TargetEdge::End)],
        vec![delete_paragraphs("00000003", "00000003")],
    ] {
        assert_eq!(code(&doc, steps), EditFailureCode::Unsupported);
    }
    assert!(undo.undo());
    assert_eq!(accepted(&doc), ["A", "B", "C"]);
    apply(&doc, &undo, vec![add("00000003", TargetEdge::Start)]);
    assert_eq!(accepted(&doc), ["A", "B", "new", "C"]);
}

#[test]
fn opaque_blocks_before_other_blocks_refuse_steps_whose_save_would_move_them() {
    let table = |id: &str| {
        format!(
            r#"<w:tbl><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid><w:tr><w:tc>{}</w:tc></w:tr></w:tbl>"#,
            p(id, &r("cell"))
        )
    };
    let blocks = format!(
        r#"{}<w:sdt><w:sdtPr><w:tag w:val="nested"/></w:sdtPr><w:sdtContent>{}{}</w:sdtContent></w:sdt>"#,
        table("0000C001"),
        table("0000D002"),
        p("0000D001", &r("control"))
    );
    let add = |id: &str, at| insert_paragraphs(id, at, vec![new_paragraph("X", None)]);
    let owned = |story: &str, id: &str| replace(TextTarget::Paragraph(para(story, id)), "edited");
    let nested = || {
        [
            owned("body:t0:r0c0", "0000C001"),
            owned("body:sdt0", "0000D001"),
            owned("body:sdt0:t0:r0c0", "0000D002"),
        ]
    };
    let displaced = open(&format!(
        "{}{RAW}{blocks}{}",
        p("00000001", &r("A")),
        p("00000002", &r("B"))
    ));
    for step in [
        add("00000001", TargetEdge::End),
        replace(search("A", "00000001"), "A2"),
        set_style("00000002", "Quote"),
    ]
    .into_iter()
    .chain(nested())
    {
        assert_eq!(code(&displaced, vec![step]), EditFailureCode::Unsupported);
    }
    let in_place = open(&format!(
        "{}{blocks}{RAW}{}",
        p("00000001", &r("A")),
        p("00000002", &r("B"))
    ));
    assert_eq!(
        code(&in_place, vec![add("00000002", TargetEdge::Start)]),
        EditFailureCode::Unsupported
    );
    let mut steps = vec![
        add("00000001", TargetEdge::End),
        replace(search("B", "00000002"), "B2"),
    ];
    steps.extend(nested());
    apply(&in_place, &UndoSession::new(), steps);
    assert_eq!(accepted(&in_place), ["A", "X", "B2"]);
    assert_eq!(
        texts(&in_place, "body:sdt0:t0:r0c0", EditTextView::Accepted),
        ["edited"]
    );
}

#[test]
fn tracked_run_formatting_changes_refuse_edits_that_would_drop_them() {
    let changed = format!(
        r#"<w:r><w:rPr><w:b/><w:rPrChange w:id="7" w:author="Ann" w:date="{DATE}"><w:rPr/></w:rPrChange></w:rPr><w:t>bold</w:t></w:r>"#
    );
    let doc = open(&format!(
        "{}{}{}",
        p("00000001", &format!("{}{changed}", r("plain "))),
        p(
            "00000002",
            &format!(
                r#"<w:hyperlink w:anchor="top"><w:r><w:t>link</w:t></w:r></w:hyperlink>{changed}"#
            )
        ),
        p("00000003", &r("free")),
    ));
    for id in ["00000001", "00000002"] {
        for steps in [
            vec![insert(
                TextTarget::Paragraph(body(id)),
                TargetEdge::End,
                "!",
            )],
            vec![set_style(id, "Quote")],
            vec![delete_paragraphs(id, id)],
        ] {
            assert_eq!(
                code(&doc, steps),
                EditFailureCode::TrackedRevisionConflict,
                "{id}"
            );
        }
    }
    apply(
        &doc,
        &UndoSession::new(),
        vec![
            replace(search("free", "00000003"), "FREE"),
            insert_paragraphs(
                "00000001",
                TargetEdge::End,
                vec![new_paragraph("next", None)],
            ),
        ],
    );
}

#[test]
fn searches_stream_matches_with_incremental_offsets() {
    let doc = open(&p(
        "00000001",
        &r(&format!("😀a😀a{}", "b".repeat(200_000))),
    ));
    let find = |text: &str, limit| {
        doc.find_text(&FindTextRequest {
            text: text.to_owned(),
            within: SearchScope::Paragraph(body("00000001")),
            view: EditTextView::Accepted,
            limit,
        })
        .unwrap()
    };
    let found = find("a", None);
    assert_eq!(
        found
            .matches
            .iter()
            .map(|found| (found.range.start.offset, found.range.end.offset))
            .collect::<Vec<_>>(),
        [(2, 3), (5, 6)]
    );
    let limited = find("bb", Some(1));
    assert_eq!(limited.matches.len(), 1);
    assert_eq!(limited.matches[0].range.start.offset, 6);
    assert_eq!(limited.matches[0].text, "bb");
    assert!(limited.truncated);
    assert_eq!(
        code(&doc, vec![delete(search("bb", "00000001"))]),
        EditFailureCode::AmbiguousTarget
    );
}

#[test]
fn original_view_reads_imported_and_authored_style_changes() {
    let doc = open(&format!(
        r#"<w:p w14:paraId="00000001"><w:pPr><w:pStyle w:val="Quote"/><w:pPrChange w:id="4" w:author="Ann" w:date="{DATE}"><w:pPr><w:pStyle w:val="Heading1"/></w:pPr></w:pPrChange></w:pPr>{}</w:p>{}"#,
        r("imported"),
        styled("00000002", "Heading1", &r("authored"))
    ));
    doc.set_paragraph_attrs(
        &EditCtx::local("Ann", DATE).suggesting(),
        &ParaSelector::One("00000002".to_owned()),
        &ParaAttrDelta {
            other: std::collections::BTreeMap::from([(
                "pStyle".to_owned(),
                Some(Any::from("Quote")),
            )]),
            ..ParaAttrDelta::default()
        },
    )
    .unwrap();
    let styles = |view| {
        doc.read_paragraphs(&ReadParagraphsRequest {
            story: None,
            para_ids: None,
            view,
        })
        .unwrap()
        .paragraphs
        .into_iter()
        .map(|paragraph| paragraph.style_id)
        .collect::<Vec<_>>()
    };
    assert_eq!(
        styles(EditTextView::Accepted),
        [Some("Quote".to_owned()), Some("Quote".to_owned())]
    );
    assert_eq!(
        styles(EditTextView::Original),
        [Some("Heading1".to_owned()), Some("Heading1".to_owned())]
    );
}

#[test]
fn story_searches_refuse_matches_in_duplicated_paragraph_ids() {
    let doc = open(&format!(
        "{}{}",
        p("000000DD", &r("unique words")),
        p("00000003", &r("other"))
    ));
    doc.apply_raw_ops(
        "body",
        vec![RawOp::InsertEmbed {
            index: 0,
            kind: "pilcrow".to_owned(),
            payload: vec![("paraId".to_owned(), Any::from("000000DD"))],
            attrs: Default::default(),
        }],
        &EditCtx::local("", ""),
    )
    .unwrap();
    let story = TextTarget::Search {
        text: "unique".to_owned(),
        within: SearchScope::Story {
            story: "body".to_owned(),
        },
        view: EditTextView::Accepted,
    };
    assert_eq!(
        code(&doc, vec![replace(story, "x")]),
        EditFailureCode::AmbiguousTarget
    );
}

#[test]
fn validation_runs_the_staged_checks_application_runs() {
    let doc = basic();
    let minted = next_minted_id(&doc, "probe");
    let (client, counter) = minted.split_once(':').unwrap();
    let collision = format!("{client}:{}", counter.parse::<u64>().unwrap() + 1);
    doc.apply_raw_ops(
        "body",
        vec![RawOp::InsertEmbed {
            index: 0,
            kind: "pilcrow".to_owned(),
            payload: vec![("paraId".to_owned(), Any::from(collision.as_str()))],
            attrs: Default::default(),
        }],
        &EditCtx::local("", ""),
    )
    .unwrap();
    let steps = vec![insert_paragraphs(
        "00000004",
        TargetEdge::End,
        vec![new_paragraph("collides", None)],
    )];
    let validated = doc
        .validate_edits(&request(&doc, steps.clone()))
        .unwrap()
        .unwrap_err();
    assert_eq!(validated.failure.code, EditFailureCode::Unsupported);
    assert_eq!(code(&doc, steps), EditFailureCode::Unsupported);
}
