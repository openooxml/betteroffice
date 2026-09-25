use std::collections::BTreeMap;

use yrs::{Any, Transact};

use super::verify::verify_saved_comparison;
use super::*;
use crate::ops::ChunkKind;
use crate::{ChangeTarget, DEL, EditCtx, INS};

const W: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const REL: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const OFFICE: &str = "application/vnd.openxmlformats-officedocument";
const AUTHOR: &str = "Reviewer <R&D>";
const DATE: &str = "2024-05-06T07:08:09+02:00";
const UTC: &str = "2024-05-06T05:08:09Z";

fn ns() -> String {
    format!(
        r#"xmlns:w="{W}" xmlns:r="{REL}" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml""#
    )
}

fn p(text: &str) -> String {
    format!(r#"<w:p><w:r><w:t xml:space="preserve">{text}</w:t></w:r></w:p>"#)
}

fn parts(body: &str) -> BTreeMap<String, Vec<u8>> {
    let ns = ns();
    let mut parts = BTreeMap::new();
    let mut set = |name: &str, content: String| {
        parts.insert(name.to_owned(), content.into_bytes());
    };
    set(
        "[Content_Types].xml",
        format!(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="{OFFICE}.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="{OFFICE}.wordprocessingml.styles+xml"/><Override PartName="/word/settings.xml" ContentType="{OFFICE}.wordprocessingml.settings+xml"/><Override PartName="/word/header1.xml" ContentType="{OFFICE}.wordprocessingml.header+xml"/><Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/></Types>"#
        ),
    );
    set(
        "_rels/.rels",
        format!(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdDoc" Type="{REL}/officeDocument" Target="word/document.xml"/><Relationship Id="rIdCore" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/></Relationships>"#
        ),
    );
    set(
        "word/_rels/document.xml.rels",
        format!(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdStyles" Type="{REL}/styles" Target="styles.xml"/><Relationship Id="rIdSettings" Type="{REL}/settings" Target="settings.xml"/><Relationship Id="rIdHeader" Type="{REL}/header" Target="header1.xml"/><Relationship Id="rIdLink" Type="{REL}/hyperlink" Target="https://example.com/" TargetMode="External"/></Relationships>"#
        ),
    );
    set(
        "word/styles.xml",
        format!(
            r#"<w:styles {ns}><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style><w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:pPr><w:keepNext/></w:pPr><w:rPr><w:b/></w:rPr></w:style></w:styles>"#
        ),
    );
    set(
        "word/settings.xml",
        format!(
            r#"<w:settings {ns}><w:zoom w:percent="100"/><w:rsids><w:rsidRoot w:val="00A1B2C3"/></w:rsids></w:settings>"#
        ),
    );
    set(
        "word/header1.xml",
        format!(r#"<w:hdr {ns}>{}</w:hdr>"#, p("Header text")),
    );
    set(
        "docProps/core.xml",
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:title>Fixture</dc:title><dcterms:modified xsi:type="dcterms:W3CDTF">2024-01-01T00:00:00Z</dcterms:modified></cp:coreProperties>"#.to_owned(),
    );
    set(
        "word/document.xml",
        format!(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {ns}><w:body>{body}<w:sectPr><w:headerReference w:type="default" r:id="rIdHeader"/></w:sectPr></w:body></w:document>"#
        ),
    );
    parts
}

fn zip(parts: &BTreeMap<String, Vec<u8>>) -> Vec<u8> {
    let entries: Vec<(String, Vec<u8>)> = parts
        .iter()
        .map(|(name, bytes)| (name.clone(), bytes.clone()))
        .collect();
    ooxml_opc::rezip_parts(&entries).unwrap()
}

fn docx(body: &str) -> Vec<u8> {
    zip(&parts(body))
}

fn options(extra: &str) -> CompareOptions {
    parse_options(&format!(
        r#"{{"author":{},"date":"{DATE}"{extra}}}"#,
        serde_json::to_string(AUTHOR).unwrap()
    ))
    .unwrap()
    .unwrap()
}

fn compare(
    original: &[u8],
    revised: &[u8],
    options: &CompareOptions,
) -> (EditingDoc, CompareOutcome) {
    let doc = EditingDoc::new(9);
    let outcome = compare_into(&doc, &UndoSession::new(), original, revised, options).unwrap();
    (doc, outcome)
}

fn texts(doc: &EditingDoc, view: EditTextView) -> Vec<String> {
    let txn = doc.yrs_doc().transact();
    StoryView::build(doc, &txn, BODY, view)
        .unwrap()
        .paragraphs
        .into_iter()
        .map(|paragraph| paragraph.text)
        .collect()
}

fn refused(outcome: &CompareOutcome) -> Vec<CompareDiagnosticCode> {
    match outcome {
        CompareOutcome::Refused(diagnostics) => diagnostics.items.iter().map(|d| d.code).collect(),
        _ => panic!("the comparison was not refused"),
    }
}

fn applied(outcome: &CompareOutcome) -> (&[ComparedChange], &SaveManifest) {
    let applied = applied_state(outcome);
    (&applied.changes, &applied.manifest)
}

fn applied_state(outcome: &CompareOutcome) -> &CompareApplied {
    match outcome {
        CompareOutcome::Applied(applied) => applied,
        CompareOutcome::Refused(diagnostics) => panic!("refused: {:?}", diagnostics.items),
        CompareOutcome::Unchanged(_) => panic!("unchanged"),
    }
}

/// Every stamped unit's revision key and value in the body, in story order.
fn stamps(doc: &EditingDoc) -> Vec<(String, String, Any)> {
    let txn = doc.yrs_doc().transact();
    let story = crate::story_ref(&txn, BODY).unwrap();
    crate::ops::snapshot(&story, &txn)
        .into_iter()
        .flat_map(|chunk| {
            let text = match &chunk.kind {
                ChunkKind::Text(text) => text.clone(),
                _ => "\u{FFFC}".to_owned(),
            };
            [INS, DEL]
                .into_iter()
                .filter(|key| chunk.attr_active(key))
                .map(|key| (key.to_owned(), text.clone(), chunk.attrs[key].clone()))
                .collect::<Vec<_>>()
        })
        .collect()
}

fn field(value: &Any, key: &str) -> String {
    match value {
        Any::Map(map) => match map.get(key) {
            Some(Any::String(value)) => value.to_string(),
            other => format!("{other:?}"),
        },
        _ => String::new(),
    }
}

fn resolve_all(doc: &EditingDoc, manifest: &SaveManifest, accept: bool) {
    let ctx = EditCtx::local(String::new(), String::new());
    for id in manifest.revision_ids.keys() {
        let target = ChangeTarget::Revision(id.clone());
        if accept {
            doc.accept_change(&ctx, &target).unwrap();
        } else {
            doc.reject_change(&ctx, &target).unwrap();
        }
    }
}

#[test]
fn text_differences_become_attributed_tracked_changes() {
    let original = docx(
        &[
            p("The quick brown fox."),
            p("Keep this."),
            p("Remove the word here."),
            p("Add at end"),
        ]
        .concat(),
    );
    let revised = docx(
        &[
            p("The slow brown fox."),
            p("Keep this."),
            p("Remove the here."),
            p("Add at end now"),
        ]
        .concat(),
    );
    let (doc, outcome) = compare(&original, &revised, &options(""));
    let (changes, manifest) = applied(&outcome);
    let summary: Vec<_> = changes
        .iter()
        .map(|change| {
            (
                change.id.as_str(),
                change.kind,
                change.original.text.as_str(),
                change.revised.text.as_str(),
                change.original.path.clone(),
                (change.original.start, change.original.end),
                (change.revised.start, change.revised.end),
            )
        })
        .collect();
    assert_eq!(
        summary,
        vec![
            (
                "change-0",
                ChangeKind::Replacement,
                "quick",
                "slow",
                vec![0, 0],
                (4, 9),
                (4, 8)
            ),
            (
                "change-1",
                ChangeKind::Deletion,
                "word ",
                "",
                vec![0, 2],
                (11, 16),
                (11, 11)
            ),
            (
                "change-2",
                ChangeKind::Insertion,
                "",
                " now",
                vec![0, 3],
                (10, 10),
                (10, 14)
            ),
        ]
    );
    let original_sha = sha256(
        &parts(
            &[
                p("The quick brown fox."),
                p("Keep this."),
                p("Remove the word here."),
                p("Add at end"),
            ]
            .concat(),
        )["word/document.xml"],
    );
    assert_eq!(changes[0].original.part, "word/document.xml");
    assert_eq!(changes[0].original.part_sha256, original_sha);
    assert_eq!(manifest.part_sha256, original_sha);
    assert_eq!(
        manifest.paragraphs,
        vec![
            SourceParagraphTarget {
                path: vec![0, 0],
                block: 0
            },
            SourceParagraphTarget {
                path: vec![0, 2],
                block: 2
            },
            SourceParagraphTarget {
                path: vec![0, 3],
                block: 3
            },
        ]
    );
    let mut ids: Vec<u32> = manifest
        .revision_ids
        .values()
        .flat_map(|numbers| numbers.deletion.into_iter().chain(numbers.insertion))
        .collect();
    ids.sort_unstable();
    assert_eq!(ids, vec![1, 2, 3, 4]);
    assert_eq!(manifest.now, "2024-05-06T05:08:09.000Z");
    assert_eq!(manifest.seed, sha256(&original));
    assert_eq!(
        texts(&doc, EditTextView::Accepted),
        [
            "The slow brown fox.",
            "Keep this.",
            "Remove the here.",
            "Add at end now"
        ]
    );
    assert_eq!(
        texts(&doc, EditTextView::Original),
        [
            "The quick brown fox.",
            "Keep this.",
            "Remove the word here.",
            "Add at end"
        ]
    );
    let stamped = stamps(&doc);
    assert_eq!(
        stamped
            .iter()
            .map(|(key, text, _)| (key.as_str(), text.as_str()))
            .collect::<Vec<_>>(),
        vec![(INS, "slow"), (DEL, "quick"), (DEL, "word "), (INS, " now")]
    );
    assert!(
        stamped
            .iter()
            .all(|(_, _, revision)| field(revision, "author") == AUTHOR
                && field(revision, "date") == UTC)
    );
    resolve_all(&doc, manifest, true);
    assert_eq!(
        texts(&doc, EditTextView::Original),
        [
            "The slow brown fox.",
            "Keep this.",
            "Remove the here.",
            "Add at end now"
        ]
    );
    let (rejected, outcome) = compare(&original, &revised, &options(""));
    resolve_all(&rejected, applied(&outcome).1, false);
    assert_eq!(
        texts(&rejected, EditTextView::Accepted),
        [
            "The quick brown fox.",
            "Keep this.",
            "Remove the word here.",
            "Add at end"
        ]
    );
    assert!(stamps(&rejected).is_empty());
}

#[test]
fn char_granularity_reports_grapheme_changes() {
    let original = docx(&p("colour"));
    let revised = docx(&p("color"));
    let (_, word) = compare(&original, &revised, &options(""));
    assert_eq!(applied(&word).0[0].original.text, "colour");
    let (_, chars) = compare(&original, &revised, &options(r#","granularity":"char""#));
    let change = &applied(&chars).0[0];
    assert_eq!(
        (change.kind, change.original.text.as_str()),
        (ChangeKind::Deletion, "u")
    );
}

#[test]
fn identical_documents_are_a_no_op_with_bookkeeping_noted() {
    let original = docx(&p("Same"));
    let mut revised =
        parts(&p("Same").replace("<w:p>", r#"<w:p w:rsidR="00FF00FF" w14:paraId="1234ABCD">"#));
    revised.insert(
        "docProps/core.xml".to_owned(),
        String::from_utf8(parts("")["docProps/core.xml"].clone())
            .unwrap()
            .replace("2024-01-01", "2025-02-02")
            .into_bytes(),
    );
    revised.insert(
        "word/settings.xml".to_owned(),
        String::from_utf8(parts("")["word/settings.xml"].clone())
            .unwrap()
            .replace("00A1B2C3", "00D4E5F6")
            .into_bytes(),
    );
    let (_, outcome) = compare(&original, &zip(&revised), &options(""));
    let CompareOutcome::Unchanged(diagnostics) = outcome else {
        panic!("expected a no-op");
    };
    assert_eq!(
        diagnostics
            .items
            .iter()
            .map(|d| (d.code, d.severity))
            .collect::<Vec<_>>(),
        vec![
            (CompareDiagnosticCode::MetadataDifference, Severity::Info),
            (CompareDiagnosticCode::MetadataDifference, Severity::Info),
        ]
    );
}

#[test]
fn clearing_a_paragraph_keeps_its_mark() {
    let (doc, outcome) = compare(
        &docx(&[p("Gone text"), p("Stay")].concat()),
        &docx(&["<w:p/>".to_owned(), p("Stay")].concat()),
        &options(""),
    );
    let (changes, _) = applied(&outcome);
    assert_eq!(changes[0].kind, ChangeKind::Deletion);
    assert_eq!(texts(&doc, EditTextView::Accepted), ["", "Stay"]);
    assert_eq!(texts(&doc, EditTextView::Original), ["Gone text", "Stay"]);
}

#[test]
fn whole_paragraph_changes_are_refused_in_both_modes() {
    let original = docx(&[p("One"), p("Two"), p("Three")].concat());
    for (revised, code) in [
        (
            docx(&[p("One"), p("Two"), p("New"), p("Three")].concat()),
            CompareDiagnosticCode::ParagraphInsertion,
        ),
        (
            docx(&[p("One"), p("Three")].concat()),
            CompareDiagnosticCode::ParagraphDeletion,
        ),
        (
            docx(&[p("One"), "<w:p/>".to_owned(), p("Two"), p("Three")].concat()),
            CompareDiagnosticCode::ParagraphInsertion,
        ),
    ] {
        for policy in ["fail", "report"] {
            let (_, outcome) = compare(
                &original,
                &revised,
                &options(&format!(r#","unsupported":"{policy}""#)),
            );
            assert_eq!(refused(&outcome), vec![code], "{policy}");
        }
    }
    let revised = docx(&[p("Uno"), p("Two"), p("New"), p("Three")].concat());
    let (_, failed) = compare(&original, &revised, &options(""));
    assert_eq!(refused(&failed).len(), 1);
    let (_, reported) = compare(
        &original,
        &docx(&[p("One changed"), p("Two"), p("New"), p("Three"), p("More")].concat()),
        &options(r#","unsupported":"report""#),
    );
    assert_eq!(
        refused(&reported),
        vec![
            CompareDiagnosticCode::ParagraphInsertion,
            CompareDiagnosticCode::ParagraphInsertion
        ]
    );
}

#[test]
fn existing_revisions_anywhere_are_refused() {
    let tracked = r#"<w:p><w:ins w:id="7" w:author="A" w:date="2024-01-01T00:00:00Z"><w:r><w:t>x</w:t></w:r></w:ins></w:p>"#;
    let mark = r#"<w:p><w:pPr><w:rPr><w:del w:id="8" w:author="A" w:date="2024-01-01T00:00:00Z"/></w:rPr></w:pPr><w:r><w:t>x</w:t></w:r></w:p>"#;
    let clean = docx(&p("x"));
    for body in [tracked, mark] {
        let (_, outcome) = compare(&docx(body), &clean, &options(""));
        assert_eq!(
            refused(&outcome),
            vec![CompareDiagnosticCode::ExistingRevisions]
        );
        let (_, outcome) = compare(&clean, &docx(body), &options(""));
        assert_eq!(
            refused(&outcome),
            vec![CompareDiagnosticCode::ExistingRevisions]
        );
    }
    let mut header = parts(&p("x"));
    header.insert(
        "word/header1.xml".to_owned(),
        format!(
            r#"<w:hdr {}><w:p><w:r><w:rPr><w:rPrChange w:id="3" w:author="A"><w:rPr/></w:rPrChange></w:rPr><w:t>H</w:t></w:r></w:p></w:hdr>"#,
            ns()
        )
        .into_bytes(),
    );
    let (_, outcome) = compare(
        &clean,
        &zip(&header),
        &options(r#","unsupported":"report""#),
    );
    let codes = refused(&outcome);
    assert_eq!(codes[0], CompareDiagnosticCode::ExistingRevisions);
}

#[test]
fn tables_are_kept_and_their_changes_diagnosed() {
    let table = |text: &str| {
        format!(
            r#"<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr>{}</w:tc></w:tr></w:tbl>"#,
            p(text)
        )
    };
    let original = docx(&[p("Before"), table("Cell"), p("After")].concat());
    let (doc, outcome) = compare(
        &original,
        &docx(&[p("Before edit"), table("Cell"), p("After")].concat()),
        &options(""),
    );
    let (changes, manifest) = applied(&outcome);
    assert_eq!(changes.len(), 1);
    assert_eq!(manifest.paragraphs[0].block, 0);
    assert_eq!(texts(&doc, EditTextView::Accepted)[0], "Before edit");
    let (_, changed) = compare(
        &original,
        &docx(&[p("Before"), table("Cell edited"), p("After")].concat()),
        &options(""),
    );
    assert_eq!(refused(&changed), vec![CompareDiagnosticCode::TableChange]);
    let (_, added) = compare(
        &original,
        &docx(&[p("Before"), table("Cell"), table("New"), p("After")].concat()),
        &options(""),
    );
    assert_eq!(refused(&added), vec![CompareDiagnosticCode::TableChange]);
    let (_, after_table) = compare(
        &original,
        &docx(&[p("Before"), table("Cell"), p("After, edited")].concat()),
        &options(""),
    );
    assert_eq!(applied(&after_table).1.paragraphs[0].block, 2);
}

#[test]
fn formatting_changes_are_refused_rather_than_redlined() {
    let run = |props: &str, text: &str| {
        format!(r#"<w:r><w:rPr>{props}</w:rPr><w:t xml:space="preserve">{text}</w:t></w:r>"#)
    };
    let original = docx(&format!(
        "<w:p>{}{}</w:p>",
        run("", "plain "),
        run("", "word")
    ));
    let bold_only = docx(&format!(
        "<w:p>{}{}</w:p>",
        run("", "plain "),
        run("<w:b/>", "word")
    ));
    let (_, outcome) = compare(&original, &bold_only, &options(""));
    assert_eq!(
        refused(&outcome),
        vec![CompareDiagnosticCode::FormattingChange]
    );
    let bold_and_longer = docx(&format!(
        "<w:p>{}{}</w:p>",
        run("", "plain "),
        run("<w:b/>", "words")
    ));
    let (_, outcome) = compare(&original, &bold_and_longer, &options(""));
    assert_eq!(
        refused(&outcome),
        vec![CompareDiagnosticCode::FormattingChange]
    );
    let restyled = docx(&format!(
        r#"<w:p><w:pPr><w:jc w:val="center"/></w:pPr>{}{}</w:p>"#,
        run("", "plain "),
        run("", "words")
    ));
    let (_, outcome) = compare(&original, &restyled, &options(""));
    assert_eq!(
        refused(&outcome),
        vec![CompareDiagnosticCode::FormattingChange]
    );
    let mut styles = parts(&format!(
        "<w:p>{}{}</w:p>",
        run("", "plain "),
        run("", "word")
    ));
    styles.insert(
        "word/styles.xml".to_owned(),
        String::from_utf8(styles["word/styles.xml"].clone())
            .unwrap()
            .replace("<w:keepNext/>", "<w:keepLines/>")
            .into_bytes(),
    );
    let (_, outcome) = compare(&original, &zip(&styles), &options(""));
    assert_eq!(
        refused(&outcome),
        vec![CompareDiagnosticCode::FormattingChange]
    );
    let unsupported = docx(&format!(
        "<w:p>{}{}</w:p>",
        run("", "plain "),
        run(r#"<w:lang w:val="de-DE"/>"#, "Wort")
    ));
    let (_, outcome) = compare(&original, &unsupported, &options(""));
    assert_eq!(
        refused(&outcome),
        vec![CompareDiagnosticCode::UnsupportedFormatting]
    );
}

#[test]
fn inserted_runs_carry_revised_formatting_and_deleted_runs_keep_theirs() {
    let run = |props: &str, text: &str| {
        format!(r#"<w:r><w:rPr>{props}</w:rPr><w:t xml:space="preserve">{text}</w:t></w:r>"#)
    };
    let original = docx(&format!(
        "<w:p>{}{}</w:p>",
        run("", "Keep "),
        run("<w:i/>", "old")
    ));
    let revised = docx(&format!(
        "<w:p>{}{}{}</w:p>",
        run("", "Keep "),
        run("<w:b/>", "new"),
        run(r#"<w:color w:val="FF0000"/>"#, " hue")
    ));
    let (doc, outcome) = compare(&original, &revised, &options(""));
    assert_eq!(applied(&outcome).0.len(), 1);
    let txn = doc.yrs_doc().transact();
    let story = crate::story_ref(&txn, BODY).unwrap();
    let styled: Vec<(String, bool, bool, bool, bool)> = crate::ops::snapshot(&story, &txn)
        .into_iter()
        .filter_map(|chunk| match &chunk.kind {
            ChunkKind::Text(text) => Some((
                text.clone(),
                chunk.attr_active(INS),
                chunk.attr_active(DEL),
                chunk.attr_active("bold"),
                chunk.attr_active("italic"),
            )),
            _ => None,
        })
        .collect();
    assert_eq!(
        styled,
        vec![
            ("Keep ".to_owned(), false, false, false, false),
            ("new".to_owned(), true, false, true, false),
            (" hue".to_owned(), true, false, false, false),
            ("old".to_owned(), false, true, false, true),
        ]
    );
}

#[test]
fn fields_hyperlinks_and_objects_block_edits_to_their_paragraphs() {
    let link = |text: &str| {
        format!(
            r#"<w:p><w:r><w:t xml:space="preserve">See </w:t></w:r><w:hyperlink r:id="rIdLink"><w:r><w:t>{text}</w:t></w:r></w:hyperlink></w:p>"#
        )
    };
    let original = docx(&[link("site"), p("Plain")].concat());
    let (_, kept) = compare(
        &original,
        &docx(&[link("site"), p("Plain edit")].concat()),
        &options(""),
    );
    assert_eq!(applied(&kept).0.len(), 1);
    let (_, edited) = compare(
        &original,
        &docx(&[link("page"), p("Plain")].concat()),
        &options(""),
    );
    assert_eq!(
        refused(&edited),
        vec![CompareDiagnosticCode::UnsupportedContent]
    );
    let field = |instruction: &str| {
        format!(
            r#"<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> {instruction} </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>1</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>"#
        )
    };
    let (_, outcome) = compare(
        &docx(&field("PAGE")),
        &docx(&field("NUMPAGES")),
        &options(""),
    );
    assert_eq!(refused(&outcome), vec![CompareDiagnosticCode::FieldChange]);
}

#[test]
fn moves_and_unresolved_correspondence_are_diagnosed() {
    let original = docx(&[p("Alpha clause"), p("Beta clause"), p("Gamma clause")].concat());
    let (_, moved) = compare(
        &original,
        &docx(&[p("Beta clause"), p("Alpha clause"), p("Gamma clause")].concat()),
        &options(""),
    );
    assert_eq!(refused(&moved), vec![CompareDiagnosticCode::ParagraphMove]);
    let (_, ambiguous) = compare(
        &docx(&[p("Start"), p("red green"), p("blue yellow"), p("End")].concat()),
        &docx(&[p("Start"), p("one two"), p("three four"), p("End")].concat()),
        &options(""),
    );
    assert_eq!(
        refused(&ambiguous),
        vec![CompareDiagnosticCode::AmbiguousAlignment]
    );
}

#[test]
fn parts_outside_the_body_are_diagnosed() {
    let original = parts(&p("Body"));
    let mut header = original.clone();
    header.insert(
        "word/header1.xml".to_owned(),
        format!(r#"<w:hdr {}>{}</w:hdr>"#, ns(), p("New header")).into_bytes(),
    );
    let mut custom = original.clone();
    custom.insert("customXml/item1.xml".to_owned(), b"<root/>".to_vec());
    let mut section = original.clone();
    section.insert(
        "word/document.xml".to_owned(),
        String::from_utf8(section["word/document.xml"].clone())
            .unwrap()
            .replace(
                "<w:sectPr>",
                r#"<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>"#,
            )
            .into_bytes(),
    );
    for (revised, code) in [
        (header, CompareDiagnosticCode::OutOfScopeChange),
        (custom, CompareDiagnosticCode::OpaquePartChange),
        (section, CompareDiagnosticCode::StructureChange),
    ] {
        let (_, outcome) = compare(&zip(&original), &zip(&revised), &options(""));
        assert_eq!(refused(&outcome), vec![code]);
    }
}

#[test]
fn unicode_text_keeps_exact_spelling_and_offsets() {
    let cases = [
        ("Hi \u{1F600} there", "Hi \u{1F601} there"),
        ("cafe\u{0301} au lait", "caf\u{00e9} au lait"),
        (
            "\u{4F60}\u{597D}\u{4E16}\u{754C}",
            "\u{4F60}\u{597D}\u{670B}\u{53CB}",
        ),
        (
            "\u{05E9}\u{05DC}\u{05D5}\u{05DD} world",
            "\u{05E9}\u{05DC}\u{05D5}\u{05DD} there",
        ),
        ("a\u{1F468}\u{200D}\u{1F469}\u{200D}\u{1F467}b", "ab"),
    ];
    for granularity in ["word", "char"] {
        for (before, after) in cases {
            let (doc, outcome) = compare(
                &docx(&p(before)),
                &docx(&p(after)),
                &options(&format!(r#","granularity":"{granularity}""#)),
            );
            let (changes, _) = applied(&outcome);
            for change in changes {
                assert_eq!(
                    text::slice(before, change.original.start..change.original.end),
                    change.original.text
                );
                assert_eq!(
                    text::slice(after, change.revised.start..change.revised.end),
                    change.revised.text
                );
            }
            assert_eq!(texts(&doc, EditTextView::Accepted), [after]);
            assert_eq!(texts(&doc, EditTextView::Original), [before]);
        }
    }
    let special = |tail: &str| {
        format!(
            r#"<w:p><w:r><w:t>co</w:t><w:softHyphen/><w:t>op</w:t><w:tab/><w:t>x</w:t><w:noBreakHyphen/><w:t>{tail}</w:t></w:r></w:p>"#
        )
    };
    let (doc, outcome) = compare(&docx(&special("y")), &docx(&special("z")), &options(""));
    assert_eq!(applied(&outcome).0.len(), 1);
    assert_eq!(
        texts(&doc, EditTextView::Accepted),
        ["co\u{00ad}op\tx\u{2011}z"]
    );
}

#[test]
fn line_breaks_are_kept_but_not_inserted() {
    let broken =
        |tail: &str| format!(r#"<w:p><w:r><w:t>one</w:t><w:br/><w:t>{tail}</w:t></w:r></w:p>"#);
    let (doc, outcome) = compare(&docx(&broken("two")), &docx(&broken("three")), &options(""));
    assert_eq!(applied(&outcome).0[0].original.text, "two");
    assert_eq!(texts(&doc, EditTextView::Accepted), ["one\u{FFFC}three"]);
    let (_, outcome) = compare(&docx(&p("one two")), &docx(&broken("two")), &options(""));
    assert_eq!(
        refused(&outcome),
        vec![CompareDiagnosticCode::UnsupportedContent]
    );
}

#[test]
fn limits_and_diagnostics_are_bounded() {
    let original = docx(&[p("a one"), p("b two")].concat());
    let revised = docx(&[p("a uno"), p("b dos")].concat());
    let (_, outcome) = compare(
        &original,
        &revised,
        &options(r#","limits":{"maxChanges":1}"#),
    );
    assert_eq!(
        refused(&outcome),
        vec![CompareDiagnosticCode::LimitExceeded]
    );
    let (_, outcome) = compare(
        &original,
        &revised,
        &options(r#","limits":{"maxParagraphs":1}"#),
    );
    assert_eq!(
        refused(&outcome),
        vec![CompareDiagnosticCode::LimitExceeded]
    );
    let (_, outcome) = compare(
        &original,
        &revised,
        &options(r#","limits":{"maxInputBytes":100}"#),
    );
    assert_eq!(
        refused(&outcome),
        vec![CompareDiagnosticCode::LimitExceeded]
    );
    let (_, outcome) = compare(
        &original,
        &revised,
        &options(r#","limits":{"maxDiffCells":2}"#),
    );
    assert_eq!(
        refused(&outcome),
        vec![CompareDiagnosticCode::LimitExceeded]
    );
    let many = docx(&[p("a"), p("b"), p("c"), p("d")].concat());
    let (_, outcome) = compare(
        &many,
        &docx(
            &[
                p("a"),
                p("new 1"),
                p("b"),
                p("new 2"),
                p("c"),
                p("new 3"),
                p("d"),
            ]
            .concat(),
        ),
        &options(r#","unsupported":"report","limits":{"maxDiagnostics":2}"#),
    );
    assert_eq!(
        refused(&outcome),
        vec![
            CompareDiagnosticCode::ParagraphInsertion,
            CompareDiagnosticCode::DiagnosticsTruncated
        ]
    );
    let json = result_json(
        true,
        &[],
        &[CompareDiagnostic {
            code: CompareDiagnosticCode::MetadataDifference,
            severity: Severity::Info,
            message: "x".repeat(64),
            locations: Vec::new(),
        }],
        32,
    )
    .unwrap();
    assert!(json.contains("limit-exceeded") && json.contains("\"ok\":false"));
}

#[test]
fn options_are_refused_as_diagnostics_and_malformed_ones_throw() {
    let refusal = |json: &str| parse_options(json).unwrap().unwrap_err().code;
    assert_eq!(
        refusal(r#"{"author":" ","date":"2024-01-01T00:00:00Z"}"#),
        CompareDiagnosticCode::InvalidOptions
    );
    assert_eq!(
        refusal(r#"{"author":"A","date":"2024-01-01T00:00:00"}"#),
        CompareDiagnosticCode::InvalidOptions
    );
    assert_eq!(
        refusal(r#"{"author":"A","date":"2024-01-01T00:00:00Z","limits":{"maxChanges":1000}}"#),
        CompareDiagnosticCode::InvalidOptions
    );
    assert!(parse_options(r#"{"author":"A"}"#).is_err());
    assert!(parse_options(r#"{"author":"A","date":"2024-01-01T00:00:00Z","extra":1}"#).is_err());
    let (_, outcome) = compare(b"not a zip", &docx(&p("x")), &options(""));
    assert_eq!(refused(&outcome), vec![CompareDiagnosticCode::InvalidDocx]);
}

#[test]
fn repeated_paragraph_ids_warn_and_changes_target_positions() {
    let with_id = |text: &str| p(text).replace("<w:p>", r#"<w:p w14:paraId="0000ABCD">"#);
    let (doc, outcome) = compare(
        &docx(&[with_id("First"), with_id("Second")].concat()),
        &docx(&[with_id("First"), with_id("Second edit")].concat()),
        &options(""),
    );
    let diagnostics = &applied_state(&outcome).diagnostics.items;
    assert_eq!(
        diagnostics[0].code,
        CompareDiagnosticCode::AmbiguousIdentity
    );
    assert_eq!(diagnostics[0].severity, Severity::Warning);
    assert_eq!(
        texts(&doc, EditTextView::Accepted),
        ["First", "Second edit"]
    );
}

#[test]
fn a_fresh_session_is_required() {
    let doc = EditingDoc::new(9);
    let bytes = docx(&p("x"));
    compare_into(&doc, &UndoSession::new(), &bytes, &bytes, &options("")).unwrap();
    assert!(compare_into(&doc, &UndoSession::new(), &bytes, &bytes, &options("")).is_err());
}

/// The original with each changed paragraph written as `redline` renders it.
fn saved(
    original: &BTreeMap<String, Vec<u8>>,
    outcome: &CompareOutcome,
    redline: impl Fn(&ComparedChange, RevisionNumbers) -> String,
    rest: impl Fn(&str) -> String,
) -> Vec<u8> {
    let CompareApplied {
        changes,
        postconditions,
        ..
    } = applied_state(outcome);
    let xml = String::from_utf8(original["word/document.xml"].clone()).unwrap();
    let mut patched = xml.clone();
    for paragraph in postconditions.paragraphs.iter().rev() {
        let span = docx_parse::element_span(&xml, &paragraph.path).unwrap();
        let full = texts_of(&xml, &paragraph.path);
        let mut content = String::new();
        let mut cursor = 0;
        for index in paragraph.changes.clone() {
            let change = &changes[index];
            content.push_str(&rest(text::slice(&full, cursor..change.original.start)));
            content.push_str(&redline(change, postconditions.change_ids[index]));
            cursor = change.original.end;
        }
        content.push_str(&rest(text::slice(&full, cursor..u32::MAX)));
        let fragment = &xml[span.clone()];
        let ppr = fragment
            .find("<w:pPr>")
            .zip(fragment.find("</w:pPr>"))
            .map_or("", |(start, end)| &fragment[start..end + "</w:pPr>".len()]);
        patched.replace_range(span, &format!("<w:p>{ppr}{content}</w:p>"));
    }
    let mut parts = original.clone();
    parts.insert("word/document.xml".to_owned(), patched.into_bytes());
    zip(&parts)
}

fn texts_of(xml: &str, path: &[u32]) -> String {
    let span = docx_parse::element_span(xml, path).unwrap();
    let fragment = &xml[span];
    fragment
        .split("<w:t")
        .skip(1)
        .map(|piece| {
            let start = piece.find('>').unwrap() + 1;
            let end = piece.find("</w:t>").unwrap();
            piece[start..end].to_owned()
        })
        .collect()
}

fn escape(text: &str) -> String {
    text.replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('>', "&gt;")
}

fn plain(text: &str) -> String {
    if text.is_empty() {
        String::new()
    } else {
        format!(
            r#"<w:r><w:t xml:space="preserve">{}</w:t></w:r>"#,
            escape(text)
        )
    }
}

fn faithful(change: &ComparedChange, ids: RevisionNumbers) -> String {
    let mut out = String::new();
    if let Some(id) = ids.deletion {
        out.push_str(&format!(
            r#"<w:del w:id="{id}" w:author="{}" w:date="{UTC}"><w:r><w:delText xml:space="preserve">{}</w:delText></w:r></w:del>"#,
            escape(AUTHOR),
            escape(&change.original.text)
        ));
    }
    if let Some(id) = ids.insertion {
        out.push_str(&format!(
            r#"<w:ins w:id="{id}" w:author="{}" w:date="{UTC}">{}</w:ins>"#,
            escape(AUTHOR),
            plain(&change.revised.text)
        ));
    }
    out
}

#[test]
fn saved_results_are_verified_against_both_inputs() {
    let body = [p("Alpha beta."), p("Unchanged."), p("Gamma delta.")].concat();
    let original = parts(&body);
    let revised = docx(&[p("Alpha gamma."), p("Unchanged."), p("Gamma delta!")].concat());
    let (_, outcome) = compare(&zip(&original), &revised, &options(""));
    let postconditions = &applied_state(&outcome).postconditions;
    let limits = options("").limits;
    let good = saved(&original, &outcome, faithful, plain);
    assert_eq!(
        verify_saved_comparison(&good, postconditions, &limits),
        Ok(())
    );

    let code = |bytes: &[u8]| {
        verify_saved_comparison(bytes, postconditions, &limits)
            .unwrap_err()
            .into_iter()
            .map(|diagnostic| diagnostic.code)
            .collect::<Vec<_>>()
    };
    let wrong_author = saved(
        &original,
        &outcome,
        |change, ids| faithful(change, ids).replace(&escape(AUTHOR), "Someone"),
        plain,
    );
    assert_eq!(
        code(&wrong_author),
        vec![CompareDiagnosticCode::RoundtripMismatch]
    );
    let wrong_id = saved(
        &original,
        &outcome,
        |change, ids| {
            faithful(
                change,
                RevisionNumbers {
                    deletion: ids.deletion.map(|id| id + 100),
                    ..ids
                },
            )
        },
        plain,
    );
    assert_eq!(
        code(&wrong_id),
        vec![CompareDiagnosticCode::RoundtripMismatch]
    );
    let bolded = saved(&original, &outcome, faithful, |text| {
        plain(text).replace("<w:r>", "<w:r><w:rPr><w:b/></w:rPr>")
    });
    assert_eq!(
        code(&bolded),
        vec![CompareDiagnosticCode::RoundtripMismatch]
    );
    let untracked = saved(
        &original,
        &outcome,
        |change, _| plain(&change.revised.text),
        plain,
    );
    assert_eq!(
        code(&untracked),
        vec![CompareDiagnosticCode::RoundtripMismatch]
    );
    let mut header = unzip(&good);
    header.insert("word/header1.xml".to_owned(), b"<w:hdr/>".to_vec());
    assert_eq!(
        code(&zip(&header)),
        vec![CompareDiagnosticCode::RoundtripMismatch]
    );
    let mut outside = unzip(&good);
    let document = String::from_utf8(outside["word/document.xml"].clone())
        .unwrap()
        .replace("Unchanged.", "Unchanged!");
    outside.insert("word/document.xml".to_owned(), document.into_bytes());
    assert_eq!(
        code(&zip(&outside)),
        vec![CompareDiagnosticCode::RoundtripMismatch]
    );
    assert_eq!(
        verify_saved_comparison(
            &good,
            postconditions,
            &CompareLimits {
                max_output_bytes: 10,
                ..limits
            }
        )
        .unwrap_err()[0]
            .code,
        CompareDiagnosticCode::LimitExceeded
    );
}

fn unzip(bytes: &[u8]) -> BTreeMap<String, Vec<u8>> {
    ooxml_opc::unzip_parts(bytes).unwrap().into_iter().collect()
}

#[test]
fn revision_ids_follow_the_original_annotation_ids() {
    let bookmarked = r#"<w:p><w:bookmarkStart w:id="41" w:name="kept"/><w:r><w:t>Anchor</w:t></w:r><w:bookmarkEnd w:id="41"/></w:p>"#;
    let foreign = r#"<w:p><w:r><w:t id="900">Other</w:t></w:r></w:p>"#;
    let (_, outcome) = compare(
        &docx(&[bookmarked.to_owned(), foreign.to_owned(), p("old words")].concat()),
        &docx(&[bookmarked.to_owned(), foreign.to_owned(), p("new words")].concat()),
        &options(""),
    );
    let ids: Vec<RevisionNumbers> = applied(&outcome).1.revision_ids.values().copied().collect();
    assert_eq!(
        ids,
        vec![RevisionNumbers {
            deletion: Some(42),
            insertion: Some(43)
        }]
    );
}

#[test]
fn a_late_batch_refusal_applies_nothing() {
    let opaque = r#"<bofx:block xmlns:bofx="urn:fidelity" bofx:value="kept"/>"#;
    let table = format!(
        r#"<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid><w:tr><w:tc>{}</w:tc></w:tr></w:tbl>"#,
        p("Cell")
    );
    let last = p("Last").replace("<w:p>", r#"<w:p w14:paraId="0000AAAA">"#);
    let body = |text: &str| [p(text), opaque.to_owned(), table.clone(), last.clone()].concat();
    let (doc, outcome) = compare(&docx(&body("Before")), &docx(&body("After")), &options(""));
    assert_eq!(refused(&outcome), vec![CompareDiagnosticCode::BatchRefused]);
    assert_eq!(texts(&doc, EditTextView::Accepted), ["Before", "Last"]);
    assert!(stamps(&doc).is_empty());
}

#[test]
fn wrapper_and_body_level_changes_are_structure_changes() {
    let wrapped = |element: &str, text: &str| {
        format!(
            r#"<w:customXml w:element="{element}">{}</w:customXml>"#,
            p(text)
        )
    };
    let original = docx(&[wrapped("clause", "Same"), p("Other")].concat());
    for revised in [
        [wrapped("term", "Same"), p("Other")].concat(),
        [p("Same"), p("Other")].concat(),
        [
            wrapped("clause", "Same"),
            r#"<w:bookmarkStart w:id="5" w:name="mark"/><w:bookmarkEnd w:id="5"/>"#.to_owned(),
            p("Other"),
        ]
        .concat(),
    ] {
        let (_, outcome) = compare(&original, &docx(&revised), &options(""));
        assert_eq!(
            refused(&outcome),
            vec![CompareDiagnosticCode::StructureChange]
        );
    }
    let (doc, outcome) = compare(
        &original,
        &docx(&[wrapped("clause", "Same edited"), p("Other")].concat()),
        &options(""),
    );
    assert_eq!(applied(&outcome).1.paragraphs[0].path, vec![0, 0, 0]);
    assert_eq!(
        texts(&doc, EditTextView::Accepted),
        ["Same edited", "Other"]
    );
}

#[test]
fn formatting_the_save_path_cannot_keep_is_refused() {
    let shaded = |fill: &str, text: &str| {
        format!(
            r#"<w:p><w:r><w:t xml:space="preserve">Keep </w:t></w:r><w:r><w:rPr><w:shd {fill}/></w:rPr><w:t>{text}</w:t></w:r></w:p>"#
        )
    };
    let pattern = r#"w:val="pct20" w:color="auto" w:fill="FFFF00""#;
    let (_, outcome) = compare(
        &docx(&shaded(pattern, "old")),
        &docx(&shaded(pattern, "new")),
        &options(""),
    );
    assert_eq!(
        refused(&outcome),
        vec![CompareDiagnosticCode::UnsupportedFormatting]
    );
    let fill = r#"w:val="clear" w:color="auto" w:fill="FFFF00""#;
    let (_, outcome) = compare(
        &docx(&shaded(fill, "old")),
        &docx(&shaded(fill, "new")),
        &options(""),
    );
    applied(&outcome);
    let hinted = |text: &str| {
        format!(
            r#"<w:p><w:r><w:rPr><w:rFonts w:hint="eastAsia" w:ascii="Arial"/></w:rPr><w:t>{text}</w:t></w:r></w:p>"#
        )
    };
    let (_, outcome) = compare(&docx(&hinted("a b")), &docx(&hinted("a c")), &options(""));
    assert_eq!(
        refused(&outcome),
        vec![CompareDiagnosticCode::UnsupportedFormatting]
    );
}

#[test]
fn retained_characters_are_checked_in_every_granularity() {
    let original = docx(&p("the cat sat"));
    let revised = docx(&format!(
        r#"<w:p>{}<w:r><w:rPr><w:b/></w:rPr><w:t>car</w:t></w:r>{}</w:p>"#,
        r#"<w:r><w:t xml:space="preserve">the </w:t></w:r>"#,
        r#"<w:r><w:t xml:space="preserve"> sat</w:t></w:r>"#
    ));
    for granularity in ["word", "char"] {
        let (_, outcome) = compare(
            &original,
            &revised,
            &options(&format!(r#","granularity":"{granularity}""#)),
        );
        assert_eq!(
            refused(&outcome),
            vec![CompareDiagnosticCode::FormattingChange],
            "{granularity}"
        );
    }
}

#[test]
fn parts_are_found_by_content_type_and_must_be_readable() {
    let tracked = format!(
        r#"<w:hdr {}><w:p><w:ins w:id="1" w:author="A"><w:r><w:t>x</w:t></w:r></w:ins></w:p></w:hdr>"#,
        ns()
    );
    let with = |name: &str, content_type: &str, content: &str| {
        let mut parts = parts(&p("Body"));
        parts.insert(name.to_owned(), content.as_bytes().to_vec());
        let types = String::from_utf8(parts["[Content_Types].xml"].clone())
            .unwrap()
            .replace(
                "</Types>",
                &format!(r#"<Override PartName="/{name}" ContentType="{content_type}"/></Types>"#),
            );
        parts.insert("[Content_Types].xml".to_owned(), types.into_bytes());
        zip(&parts)
    };
    let header = format!("{OFFICE}.wordprocessingml.header+xml");
    for bytes in [
        with("word/HEADER2.XML", &header, &tracked),
        with("word/extra.bin", "application/xml", &tracked),
    ] {
        let (_, outcome) = compare(&bytes, &bytes, &options(""));
        assert_eq!(
            refused(&outcome),
            vec![CompareDiagnosticCode::ExistingRevisions]
        );
    }
    let broken = with("word/extra.bin", "application/xml", "<unclosed>");
    let (_, outcome) = compare(&broken, &broken, &options(""));
    assert_eq!(refused(&outcome), vec![CompareDiagnosticCode::InvalidDocx]);
    let mut untyped = parts(&p("Body"));
    untyped.insert("word/orphan.dat".to_owned(), b"data".to_vec());
    let untyped = zip(&untyped);
    let (_, outcome) = compare(&untyped, &untyped, &options(""));
    assert_eq!(refused(&outcome), vec![CompareDiagnosticCode::InvalidDocx]);
}

#[test]
fn inspection_limits_cover_every_story_before_seeding() {
    let original = docx(&[p("one"), p("two"), p("three")].concat());
    let revised = docx(&[p("one"), p("two"), p("three!")].concat());
    let (doc, outcome) = compare(
        &original,
        &revised,
        &options(r#","limits":{"maxParagraphs":3}"#),
    );
    assert_eq!(
        refused(&outcome),
        vec![CompareDiagnosticCode::LimitExceeded]
    );
    assert!(doc.story_len(BODY).is_err());
    let header_units = "Header text".len() * 2;
    let body_units = "onetwothree".len() + "onetwothree!".len();
    let exact = format!(
        r#","limits":{{"maxTextUnits":{}}}"#,
        header_units + body_units
    );
    applied(&compare(&original, &revised, &options(&exact)).1);
    let short = format!(
        r#","limits":{{"maxTextUnits":{}}}"#,
        header_units + body_units - 1
    );
    let (doc, outcome) = compare(&original, &revised, &options(&short));
    assert_eq!(
        refused(&outcome),
        vec![CompareDiagnosticCode::LimitExceeded]
    );
    assert!(doc.story_len(BODY).is_err());
}

#[test]
fn every_outcome_respects_the_result_limits() {
    let same = docx(&p("Same"));
    let (_, outcome) = compare(&same, &same, &options(r#","limits":{"maxOutputBytes":1}"#));
    assert_eq!(
        refused(&outcome),
        vec![CompareDiagnosticCode::LimitExceeded]
    );
    let with_id = |text: &str| p(text).replace("<w:p>", r#"<w:p w14:paraId="0000ABCD">"#);
    let original = docx(&[with_id("First"), with_id("Second")].concat());
    let revised = docx(&[with_id("First"), with_id("Second edit")].concat());
    let limited = options(r#","limits":{"maxDiagnostics":1}"#);
    let (_, outcome) = compare(&original, &revised, &limited);
    let CompareOutcome::Applied(applied) = outcome else {
        panic!("expected changes");
    };
    let result: serde_json::Value =
        serde_json::from_str(&applied.finish(&original, &limited.limits).unwrap()).unwrap();
    assert_eq!(result["ok"], false);
    assert_eq!(result["diagnostics"][0]["code"], "diagnostics-truncated");
    let (_, outcome) = compare(&original, &revised, &options(""));
    let CompareOutcome::Applied(applied) = outcome else {
        panic!("expected changes");
    };
    let tiny = CompareLimits {
        max_result_bytes: 16,
        ..options("").limits
    };
    let result: serde_json::Value =
        serde_json::from_str(&applied.fail("no space", &tiny).unwrap()).unwrap();
    assert_eq!(result["diagnostics"][0]["code"], "limit-exceeded");
}

fn styled(body: &str, styles: &str) -> BTreeMap<String, Vec<u8>> {
    let mut parts = parts(body);
    parts.insert(
        "word/styles.xml".to_owned(),
        format!(
            r#"<w:styles {}>{styles}<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style></w:styles>"#,
            ns()
        )
        .into_bytes(),
    );
    parts
}

/// Saves `outcome` with every run of the changed paragraphs carrying `props`.
fn saved_with(
    original: &BTreeMap<String, Vec<u8>>,
    outcome: &CompareOutcome,
    props: &str,
) -> Vec<u8> {
    let run = |text: &str, element: &str| {
        format!(
            r#"<w:r><w:rPr>{props}</w:rPr><w:{element} xml:space="preserve">{}</w:{element}></w:r>"#,
            escape(text)
        )
    };
    saved(
        original,
        outcome,
        |change, ids| {
            format!(
                r#"<w:del w:id="{}" w:author="{}" w:date="{UTC}">{}</w:del><w:ins w:id="{}" w:author="{}" w:date="{UTC}">{}</w:ins>"#,
                ids.deletion.unwrap(),
                escape(AUTHOR),
                run(&change.original.text, "delText"),
                ids.insertion.unwrap(),
                escape(AUTHOR),
                run(&change.revised.text, "t")
            )
        },
        |text| {
            if text.is_empty() {
                String::new()
            } else {
                run(text, "t")
            }
        },
    )
}

#[test]
fn verification_compares_effective_formatting_with_the_source() {
    let defaults = r#"<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Carlito" w:hAnsi="Carlito"/><w:sz w:val="22"/><w:szCs w:val="24"/></w:rPr></w:rPrDefault></w:docDefaults>"#;
    let body = |text: &str| {
        format!(r#"<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>{text} words</w:t></w:r></w:p>"#)
    };
    let original = styled(&body("old"), defaults);
    let revised = zip(&styled(&body("new"), defaults));
    let (_, outcome) = compare(&zip(&original), &revised, &options(""));
    let postconditions = &applied_state(&outcome).postconditions;
    let limits = options("").limits;
    let verify = |props: &str| {
        verify_saved_comparison(
            &saved_with(&original, &outcome, props),
            postconditions,
            &limits,
        )
        .map_err(|diagnostics| diagnostics[0].code)
    };
    assert_eq!(verify("<w:b/>"), Ok(()));
    assert_eq!(
        verify(
            r#"<w:rFonts w:ascii="Carlito" w:hAnsi="Carlito"/><w:b/><w:sz w:val="22"/><w:szCs w:val="24"/>"#
        ),
        Ok(())
    );
    for invented in [
        "<w:b/><w:bCs/>",
        r#"<w:b/><w:szCs w:val="22"/>"#,
        r#"<w:rFonts w:ascii="Carlito" w:hAnsi="Carlito" w:cs="Carlito"/><w:b/>"#,
    ] {
        assert_eq!(
            verify(invented),
            Err(CompareDiagnosticCode::RoundtripMismatch),
            "{invented}"
        );
    }
}

#[test]
fn style_toggles_are_read_at_every_level_of_a_chain() {
    let styles = r#"<w:style w:type="paragraph" w:styleId="Strong"><w:rPr><w:b/></w:rPr></w:style><w:style w:type="paragraph" w:styleId="Toggled"><w:basedOn w:val="Strong"/><w:rPr><w:b/></w:rPr></w:style><w:style w:type="paragraph" w:styleId="Off"><w:basedOn w:val="Strong"/><w:rPr><w:b w:val="0"/></w:rPr></w:style>"#;
    let limits = options("").limits;
    for style in ["Toggled", "Off"] {
        let body = |text: &str| {
            format!(
                r#"<w:p><w:pPr><w:pStyle w:val="{style}"/></w:pPr><w:r><w:t>{text} words</w:t></w:r></w:p>"#
            )
        };
        let original = styled(&body("old"), styles);
        let revised = zip(&styled(&body("new"), styles));
        let (_, outcome) = compare(&zip(&original), &revised, &options(""));
        let postconditions = &applied_state(&outcome).postconditions;
        let verify = |props: &str| {
            verify_saved_comparison(
                &saved_with(&original, &outcome, props),
                postconditions,
                &limits,
            )
            .map_err(|diagnostics| diagnostics[0].code)
        };
        assert_eq!(verify(""), Ok(()), "{style}");
        assert_eq!(
            verify("<w:b/>"),
            Err(CompareDiagnosticCode::RoundtripMismatch),
            "{style}"
        );
    }
}
