#[allow(dead_code)]
#[path = "support/page_fixture.rs"]
mod fixture;

#[path = "support/revision_boundary.rs"]
mod boundary;

use docx_edit::bridge::{RenderEnv, RevisionPreview};
use docx_edit::structured::{ExportFailureCode, PageExportOptions, RevisionView};
use docx_edit::{
    ChangeKind, EditCtx, EditHistory, EditOperation, EditRequest, EditSource, EditStep,
    EditSuggestion, EditTextView, EngineSession, FormatPolicy, ParagraphTarget, Position,
    SearchScope, StoryRange, TargetEdge, TextTarget, UndoSession, seed_from_docx,
};
use docx_layout::display_list::{Primitive, RevisionKind};
use serde_json::{Value, json};

use RevisionPreview::{Accepted, Rejected};

const FONT: &[u8] = include_bytes!("../../ooxml-text/tests/fonts/LiberationSans-Regular.ttf");
const NS: &str = r#"xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml""#;
const PROPOSALS: &str = r#"<w:p w14:paraId="00000001"><w:r><w:t xml:space="preserve">Alpha </w:t></w:r><w:r><w:rPr><w:b/></w:rPr><w:t>beta</w:t></w:r><w:r><w:t xml:space="preserve"> gamma</w:t></w:r></w:p><w:p w14:paraId="00000002"><w:r><w:t>Delta</w:t></w:r></w:p><w:p w14:paraId="00000003"><w:r><w:t>Title</w:t></w:r></w:p>"#;

fn document(body: &str) -> Vec<u8> {
    let parts = [
        ("[Content_Types].xml", r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/></Types>"#.to_owned()),
        ("_rels/.rels", r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"#.to_owned()),
        ("word/_rels/document.xml.rels", r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdNumbering" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/></Relationships>"#.to_owned()),
        ("word/numbering.xml", format!(r#"<w:numbering {NS}><w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>"#)),
        ("word/document.xml", format!(r#"<w:document {NS}><w:body>{body}</w:body></w:document>"#)),
    ];
    ooxml_opc::rezip_parts(
        &parts
            .into_iter()
            .map(|(name, value)| (name.to_owned(), value.into_bytes()))
            .collect::<Vec<_>>(),
    )
    .unwrap()
}

fn search(text: &str, para_id: &str) -> TextTarget {
    TextTarget::Search {
        text: text.to_owned(),
        within: SearchScope::Paragraph(ParagraphTarget {
            story: "body".to_owned(),
            para_id: para_id.to_owned(),
        }),
        view: EditTextView::Accepted,
    }
}

fn suggested(operation: EditOperation) -> EditStep {
    let mut step = EditStep::new(operation);
    step.suggest = Some(EditSuggestion {
        author: "Ann".to_owned(),
        date: "2026-09-29T12:00:00Z".to_owned(),
    });
    step
}

/// A suggested replacement, deletion and insertion, one per paragraph, and their revision ids.
fn proposals() -> (EngineSession, [String; 3]) {
    proposals_in(&document(PROPOSALS))
}

fn proposals_in(bytes: &[u8]) -> (EngineSession, [String; 3]) {
    let engine = EngineSession::new(75101);
    seed_from_docx(engine.doc(), bytes).unwrap();
    let request = EditRequest {
        expect_version: engine.doc().version(),
        source: EditSource::Host,
        history: EditHistory::None,
        steps: vec![
            suggested(EditOperation::ReplaceText {
                target: search("beta", "00000001"),
                text: "BETA".to_owned(),
            }),
            suggested(EditOperation::DeleteText {
                target: search("Delta", "00000002"),
            }),
            suggested(EditOperation::InsertText {
                target: search("Title", "00000003"),
                at: TargetEdge::End,
                text: "!".to_owned(),
            }),
        ],
    };
    let applied = engine
        .doc()
        .apply_edits(&request, &UndoSession::new())
        .unwrap()
        .unwrap();
    let ids: Vec<String> = applied
        .receipts
        .into_iter()
        .map(|receipt| {
            assert_eq!(receipt.revision_ids.len(), 1);
            receipt.revision_ids[0].clone()
        })
        .collect();
    (engine, ids.try_into().unwrap())
}

fn preview(entries: &[(&str, RevisionPreview)]) -> RenderEnv {
    entries
        .iter()
        .fold(RenderEnv::default(), |env, (id, decision)| {
            env.with_revision_preview(*id, *decision)
        })
}

fn lower(engine: &EngineSession, env: &RenderEnv) -> Value {
    serde_json::from_str(&engine.lower_story_json("body", env).unwrap()).unwrap()
}

/// `(text, pmStart, pmEnd, tracked kind)` per run of paragraph `index`.
fn runs(blocks: &Value, index: usize) -> Vec<(String, f64, f64, &'static str)> {
    blocks[index]["runs"]
        .as_array()
        .unwrap()
        .iter()
        .map(|run| {
            let kind = match (run["isInsertion"] == true, run["isDeletion"] == true) {
                (true, true) => "ins+del",
                (true, false) => "ins",
                (false, true) => "del",
                (false, false) => "",
            };
            (
                run["text"].as_str().unwrap_or_default().to_owned(),
                run["pmStart"].as_f64().unwrap(),
                run["pmEnd"].as_f64().unwrap(),
                kind,
            )
        })
        .collect()
}

fn run(text: &str, start: f64, end: f64, kind: &'static str) -> (String, f64, f64, &'static str) {
    (text.to_owned(), start, end, kind)
}

fn paragraph_bounds(blocks: &Value) -> Vec<(Value, Value)> {
    blocks
        .as_array()
        .unwrap()
        .iter()
        .map(|block| (block["pmStart"].clone(), block["pmEnd"].clone()))
        .collect()
}

#[test]
fn previews_of_changes_before_a_table_keep_valid_block_boundaries() {
    for (inserted_mark, replacement) in [(false, ""), (false, "X"), (true, "")] {
        let engine = EngineSession::new(75110);
        let plain = EditCtx::local("", "");
        let suggest = EditCtx::local("Ann", "2026-09-29T12:00:00Z").suggesting();
        let split_ctx = if inserted_mark { &suggest } else { &plain };
        let split = boundary::seed(engine.doc(), "table", split_ctx);
        let id = if inserted_mark {
            split.revision_ids[0].clone()
        } else {
            let receipt = if replacement.is_empty() {
                engine
                    .doc()
                    .delete_range(&suggest, StoryRange::new("body", 0, 4))
            } else {
                engine
                    .doc()
                    .replace_range(&suggest, StoryRange::new("body", 0, 4), replacement)
            }
            .unwrap();
            receipt.revision_ids[0].clone()
        };
        let native = lower(&engine, &RenderEnv::default());
        for decision in [Accepted, Rejected] {
            let blocks = lower(&engine, &preview(&[(&id, decision)]));
            assert_eq!(
                blocks
                    .as_array()
                    .unwrap()
                    .iter()
                    .map(|block| block["kind"].as_str().unwrap())
                    .collect::<Vec<_>>(),
                ["paragraph", "table", "paragraph"]
            );
            assert_eq!(paragraph_bounds(&blocks), paragraph_bounds(&native));
            let text: String = runs(&blocks, 0).into_iter().map(|run| run.0).collect();
            let expected = match decision {
                Accepted if !inserted_mark => replacement,
                _ => "old",
            };
            assert_eq!(text, expected);
        }
    }
}

#[test]
fn each_decision_shows_its_outcome_at_the_source_positions() {
    let (engine, [replace, delete, insert]) = proposals();
    let native = lower(&engine, &RenderEnv::default());
    assert_eq!(
        runs(&native, 0),
        [
            run("Alpha ", 1.0, 7.0, ""),
            run("beta", 7.0, 11.0, "del"),
            run("BETA", 11.0, 15.0, "ins"),
            run(" gamma", 15.0, 21.0, ""),
        ]
    );
    let expected = |decision| match decision {
        None => (runs(&native, 0), runs(&native, 1), runs(&native, 2)),
        Some(Accepted) => (
            vec![
                run("Alpha ", 1.0, 7.0, ""),
                run("BETA", 11.0, 15.0, ""),
                run(" gamma", 15.0, 21.0, ""),
            ],
            vec![],
            vec![run("Title!", 30.0, 36.0, "")],
        ),
        Some(Rejected) => (
            vec![
                run("Alpha ", 1.0, 7.0, ""),
                run("beta", 7.0, 11.0, ""),
                run(" gamma", 15.0, 21.0, ""),
            ],
            vec![run("Delta", 23.0, 28.0, "")],
            vec![run("Title", 30.0, 35.0, "")],
        ),
    };
    let transitions = [
        None,
        Some(Accepted),
        Some(Rejected),
        None,
        Some(Rejected),
        Some(Accepted),
        None,
    ];
    for decision in transitions {
        let env = match decision {
            Some(decision) => preview(&[
                (&replace, decision),
                (&delete, decision),
                (&insert, decision),
            ]),
            None => RenderEnv::default(),
        };
        let blocks = lower(&engine, &env);
        assert_eq!(
            (runs(&blocks, 0), runs(&blocks, 1), runs(&blocks, 2)),
            expected(decision),
            "{decision:?}"
        );
        assert_eq!(paragraph_bounds(&blocks), paragraph_bounds(&native));
        assert_eq!(blocks, lower(&proposals().0, &env), "{decision:?}");
    }
}

#[test]
fn proposals_decide_independently() {
    let (engine, [replace, delete, insert]) = proposals();
    let native = lower(&engine, &RenderEnv::default());
    let mixed = lower(
        &engine,
        &preview(&[(&replace, Accepted), (&delete, Rejected)]),
    );
    assert_eq!(
        runs(&mixed, 0),
        [
            run("Alpha ", 1.0, 7.0, ""),
            run("BETA", 11.0, 15.0, ""),
            run(" gamma", 15.0, 21.0, ""),
        ]
    );
    assert_eq!(runs(&mixed, 1), [run("Delta", 23.0, 28.0, "")]);
    assert_eq!(mixed[2], native[2]);
    let other = lower(&engine, &preview(&[(&insert, Rejected)]));
    assert_eq!(other[0], native[0]);
    assert_eq!(other[1], native[1]);
    assert_eq!(runs(&other, 2), [run("Title", 30.0, 35.0, "")]);
}

#[test]
fn decided_runs_keep_their_formatting_and_lose_only_the_markup() {
    let (engine, [replace, ..]) = proposals();
    let native = lower(&engine, &RenderEnv::default());
    let without_markup = |run: &Value| {
        let mut run = run.clone();
        for key in [
            "isInsertion",
            "isDeletion",
            "changeAuthor",
            "changeDate",
            "changeRevisionId",
            "logicalOrder",
        ] {
            run.as_object_mut().unwrap().remove(key);
        }
        run
    };
    let accepted = lower(&engine, &preview(&[(&replace, Accepted)]));
    assert_eq!(accepted[0]["runs"][1]["bold"], true);
    assert_eq!(
        without_markup(&accepted[0]["runs"][1]),
        without_markup(&native[0]["runs"][2])
    );
    let rejected = lower(&engine, &preview(&[(&replace, Rejected)]));
    assert_eq!(rejected[0]["runs"][1]["bold"], true);
    assert_eq!(
        without_markup(&rejected[0]["runs"][1]),
        without_markup(&native[0]["runs"][1])
    );
    assert_eq!(
        without_markup(&rejected[0]["runs"][2]),
        without_markup(&native[0]["runs"][3])
    );
    assert_eq!(lower(&engine, &RenderEnv::default()), native);
}

#[test]
fn imported_revisions_are_keyed_by_their_listed_ids() {
    let body = r#"<w:p><w:r><w:t xml:space="preserve">A </w:t></w:r><w:ins w:id="1" w:author="Bo" w:date="2026-09-29T12:00:00Z"><w:r><w:t>new</w:t></w:r></w:ins><w:del w:id="2" w:author="Bo" w:date="2026-09-29T12:00:00Z"><w:r><w:delText>old</w:delText></w:r></w:del><w:r><w:t xml:space="preserve"> Z</w:t></w:r></w:p>"#;
    let engine = EngineSession::new(75102);
    seed_from_docx(engine.doc(), &document(body)).unwrap();
    let ids: Vec<String> = engine
        .doc()
        .list_revisions()
        .unwrap()
        .into_iter()
        .map(|revision| revision.change.revision_id)
        .collect();
    assert_eq!(ids, ["1", "2"]);
    let view = |entries: &[(&str, RevisionPreview)]| runs(&lower(&engine, &preview(entries)), 0);
    let native = view(&[]);
    assert_eq!(
        native,
        [
            run("A ", 1.0, 3.0, ""),
            run("new", 3.0, 6.0, "ins"),
            run("old", 6.0, 9.0, "del"),
            run(" Z", 9.0, 11.0, ""),
        ]
    );
    assert_eq!(
        view(&[("1", Accepted), ("2", Rejected)]),
        [run("A newold Z", 1.0, 11.0, "")]
    );
    assert_eq!(
        view(&[("1", Rejected), ("2", Accepted)]),
        [run("A ", 1.0, 3.0, ""), run(" Z", 9.0, 11.0, "")]
    );
    assert_eq!(view(&[("5", Accepted), ("9:9", Rejected)]), native);
}

#[test]
fn a_deleted_insertion_shows_whichever_side_is_still_pending() {
    let engine = EngineSession::new(75103);
    seed_from_docx(
        engine.doc(),
        &document(r#"<w:p><w:r><w:t>AZ</w:t></w:r></w:p>"#),
    )
    .unwrap();
    let suggest = |author| EditCtx::local(author, "2026-09-29T12:00:00Z").suggesting();
    engine
        .doc()
        .insert_text(
            &suggest("Ann"),
            Position::new("body", 1),
            "both",
            FormatPolicy::Plain,
        )
        .unwrap();
    engine
        .doc()
        .delete_range(&suggest("Bob"), StoryRange::new("body", 1, 5))
        .unwrap();
    let revisions = engine.doc().list_revisions().unwrap();
    let id = |kind| {
        revisions
            .iter()
            .find(|revision| revision.change.kind == kind)
            .unwrap()
            .change
            .revision_id
            .clone()
    };
    let (insertion, deletion) = (id(ChangeKind::Insertion), id(ChangeKind::Deletion));
    let both = |entries: &[(&str, RevisionPreview)]| {
        runs(&lower(&engine, &preview(entries)), 0)
            .into_iter()
            .find(|run| run.0.contains("both"))
            .map(|run| run.3)
    };
    assert_eq!(both(&[]), Some("ins+del"));
    assert_eq!(both(&[(&insertion, Accepted)]), Some("del"));
    assert_eq!(both(&[(&deletion, Rejected)]), Some("ins"));
    assert_eq!(
        both(&[(&insertion, Accepted), (&deletion, Rejected)]),
        Some("")
    );
    assert_eq!(both(&[(&insertion, Rejected)]), None);
    assert_eq!(both(&[(&deletion, Accepted)]), None);
    assert_eq!(both(&[(&insertion, Rejected), (&deletion, Rejected)]), None);
}

#[test]
fn a_previewed_control_revision_covers_the_controls_content() {
    let body = r#"<w:p><w:r><w:t xml:space="preserve">A </w:t></w:r><w:sdt><w:sdtPr><w:tag w:val="outer"/></w:sdtPr><w:sdtContent><w:r><w:t>SECRET</w:t></w:r><w:sdt><w:sdtPr><w:tag w:val="inner"/></w:sdtPr><w:sdtContent><w:r><w:t>nested</w:t></w:r></w:sdtContent></w:sdt></w:sdtContent></w:sdt><w:r><w:t xml:space="preserve"> Z</w:t></w:r></w:p>"#;
    let engine = EngineSession::new(75104);
    seed_from_docx(engine.doc(), &document(body)).unwrap();
    engine
        .doc()
        .delete_range(
            &EditCtx::local("Ann", "2026-09-29T12:00:00Z").suggesting(),
            StoryRange::new("body", 2, 3),
        )
        .unwrap();
    let revisions = engine.doc().list_revisions().unwrap();
    assert_eq!(revisions.len(), 1);
    let id = revisions[0].change.revision_id.clone();
    let texts = |entries: &[(&str, RevisionPreview)]| -> Vec<String> {
        runs(&lower(&engine, &preview(entries)), 0)
            .into_iter()
            .map(|run| run.0)
            .collect()
    };
    assert_eq!(texts(&[]), texts(&[(&id, Rejected)]));
    assert!(texts(&[]).concat().contains("SECRETnested"));
    let accepted = texts(&[(&id, Accepted)]);
    assert_eq!(accepted.concat(), "A  Z");
    let native = lower(&engine, &RenderEnv::default());
    let hidden = lower(&engine, &preview(&[(&id, Accepted)]));
    assert_eq!(hidden[0]["pmEnd"], native[0]["pmEnd"]);
    assert_eq!(
        hidden[0]["runs"].as_array().unwrap().last().unwrap()["pmStart"],
        native[0]["runs"].as_array().unwrap().last().unwrap()["pmStart"]
    );
}

#[test]
fn a_table_cell_previews_its_own_revisions() {
    let body = r#"<w:tbl><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid><w:tr><w:tc><w:p w14:paraId="00000010"><w:r><w:t>cell beta</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:p w14:paraId="00000011"><w:r><w:t>After</w:t></w:r></w:p>"#;
    let engine = EngineSession::new(75105);
    seed_from_docx(engine.doc(), &document(body)).unwrap();
    let request = EditRequest {
        expect_version: engine.doc().version(),
        source: EditSource::Host,
        history: EditHistory::None,
        steps: vec![suggested(EditOperation::ReplaceText {
            target: TextTarget::Search {
                text: "beta".to_owned(),
                within: SearchScope::Paragraph(ParagraphTarget {
                    story: "body:t0:r0c0".to_owned(),
                    para_id: "00000010".to_owned(),
                }),
                view: EditTextView::Accepted,
            },
            text: "BETA".to_owned(),
        })],
    };
    let applied = engine
        .doc()
        .apply_edits(&request, &UndoSession::new())
        .unwrap()
        .unwrap();
    let id = applied.receipts[0].revision_ids[0].clone();
    let cell_runs = |env: &RenderEnv| {
        let blocks = lower(&engine, env);
        runs(&blocks[0]["rows"][0]["cells"][0]["blocks"], 0)
            .into_iter()
            .map(|run| (run.0, run.3))
            .collect::<Vec<_>>()
    };
    assert_eq!(
        cell_runs(&RenderEnv::default()),
        [
            ("cell ".to_owned(), ""),
            ("beta".to_owned(), "del"),
            ("BETA".to_owned(), "ins")
        ]
    );
    assert_eq!(
        cell_runs(&preview(&[(&id, Accepted)])),
        [("cell ".to_owned(), ""), ("BETA".to_owned(), "")]
    );
    assert_eq!(
        cell_runs(&preview(&[(&id, Rejected)])),
        [("cell beta".to_owned(), "")]
    );
}

#[test]
fn the_preview_parses_leniently_and_keys_the_lowering_cache() {
    let parsed: RenderEnv = serde_json::from_value(json!({
        "revisionPreview": {"a": "accepted", "b": "rejected", "c": "proposed", "d": 1, "e": null}
    }))
    .unwrap();
    assert_eq!(parsed, preview(&[("a", Accepted), ("b", Rejected)]));
    for empty in [json!({}), json!(null), json!("accepted"), json!(["a"])] {
        let parsed: RenderEnv =
            serde_json::from_value(json!({ "revisionPreview": empty })).unwrap();
        assert_eq!(parsed, RenderEnv::default());
    }
    assert_eq!(
        serde_json::to_value(RenderEnv::default())
            .unwrap()
            .get("revisionPreview"),
        None
    );
    assert_eq!(
        serde_json::to_value(preview(&[("a", Accepted)])).unwrap()["revisionPreview"],
        json!({"a": "accepted"})
    );

    let (engine, [replace, ..]) = proposals();
    let misses = || engine.stats().lower_cache_misses;
    lower(&engine, &RenderEnv::default());
    let start = misses();
    let empty: RenderEnv = serde_json::from_value(json!({"revisionPreview": {}})).unwrap();
    lower(&engine, &empty);
    assert_eq!(misses(), start);
    lower(&engine, &preview(&[(&replace, Accepted)]));
    assert_eq!(misses(), start + 1);
    lower(&engine, &preview(&[(&replace, Accepted)]));
    assert_eq!(misses(), start + 1);
    lower(&engine, &preview(&[(&replace, Rejected)]));
    assert_eq!(misses(), start + 2);
}

#[test]
fn previewing_leaves_the_document_and_its_revisions_alone() {
    let (engine, [replace, delete, insert]) = proposals();
    let state = engine.doc().encode_state_as_update_v1();
    let version = engine.doc().version();
    let revisions = engine.doc().list_revisions().unwrap();
    for decision in [Accepted, Rejected] {
        lower(
            &engine,
            &preview(&[
                (&replace, decision),
                (&delete, decision),
                (&insert, decision),
            ]),
        );
    }
    assert_eq!(engine.doc().encode_state_as_update_v1(), state);
    assert_eq!(engine.doc().version(), version);
    assert_eq!(engine.doc().list_revisions().unwrap(), revisions);
}

fn layout_request(env: &RenderEnv, font: u32) -> String {
    json!({
        "bodyStory": "body", "renderEnv": env, "options": {},
        "measurement": {"fontChains": {"calibri|0|0": [font]}, "defaults": {"fontFamily": "Calibri", "fontSize": 12}}
    })
    .to_string()
}

/// Painted text, split where the revision kind changes.
fn painted(engine: &EngineSession) -> Vec<(String, Option<RevisionKind>)> {
    let mut segments: Vec<(String, Option<RevisionKind>)> = Vec::new();
    engine
        .with_display_list(|list| {
            for primitive in list.pages.iter().flat_map(|page| &page.primitives) {
                let Primitive::Text(text) = primitive else {
                    continue;
                };
                let kind = text.attrs.revision.as_ref().map(|revision| revision.kind);
                match segments.last_mut() {
                    Some((painted, last)) if *last == kind => painted.push_str(&text.text),
                    _ => segments.push((text.text.clone(), kind)),
                }
            }
        })
        .unwrap();
    segments
}

fn segment(text: &str, kind: Option<RevisionKind>) -> (String, Option<RevisionKind>) {
    (text.to_owned(), kind)
}

#[test]
fn a_changed_preview_rebuilds_the_retained_frame() {
    let font = docx_layout::register_measure_font(FONT).unwrap();
    let (engine, [replace, delete, insert]) = proposals();
    let native = RenderEnv::default();
    engine
        .layout_document_with_regions_json(&layout_request(&native, font))
        .unwrap();
    engine.build_display_list_frame("{}", 0).unwrap();
    let tracked = painted(&engine);
    assert_eq!(
        tracked,
        [
            segment("Alpha ", None),
            segment("beta", Some(RevisionKind::Del)),
            segment("BETA", Some(RevisionKind::Ins)),
            segment(" gamma", None),
            segment("Delta", Some(RevisionKind::Del)),
            segment("Title", None),
            segment("!", Some(RevisionKind::Ins)),
        ]
    );

    let mut epoch = engine.stats().frame_epoch;
    let mut frame = |env: &RenderEnv| {
        engine
            .layout_document_with_regions_json(&layout_request(env, font))
            .unwrap();
        engine.build_display_list_frame("{}", epoch).unwrap();
        assert!(engine.stats().frame_epoch > epoch);
        epoch = engine.stats().frame_epoch;
        painted(&engine)
    };
    assert_eq!(
        frame(&preview(&[(&insert, Accepted)]))[4..],
        [
            segment("Delta", Some(RevisionKind::Del)),
            segment("Title!", None)
        ]
    );
    let all = preview(&[
        (&replace, Accepted),
        (&delete, Accepted),
        (&insert, Accepted),
    ]);
    assert_eq!(frame(&all), [segment("Alpha BETA gammaTitle!", None)]);
    assert_eq!(frame(&native), tracked);
}

#[test]
fn a_previewed_page_break_revision_moves_the_following_page() {
    let font = docx_layout::register_measure_font(FONT).unwrap();
    let two_paragraphs =
        r#"<w:p><w:r><w:t>Before</w:t></w:r></w:p><w:p><w:r><w:t>After</w:t></w:r></w:p>"#;
    let with_break = |ctx: &EditCtx| {
        let engine = EngineSession::new(75106);
        seed_from_docx(engine.doc(), &document(two_paragraphs)).unwrap();
        engine
            .doc()
            .insert_embed(ctx, Position::new("body", 7), "pageBreak", vec![])
            .unwrap();
        engine
    };
    let pages = |engine: &EngineSession, env: &RenderEnv| {
        let output: Value = serde_json::from_str(
            &engine
                .layout_document_with_regions_json(&layout_request(env, font))
                .unwrap(),
        )
        .unwrap();
        output["layout"]["pages"].as_array().unwrap().len()
    };
    let suggest = EditCtx::local("Ann", "2026-09-29T12:00:00Z").suggesting();
    let inserted = with_break(&suggest);
    let id = inserted.doc().list_revisions().unwrap()[0]
        .change
        .revision_id
        .clone();
    assert_eq!(pages(&inserted, &RenderEnv::default()), 2);
    assert_eq!(pages(&inserted, &preview(&[(&id, Accepted)])), 2);
    assert_eq!(pages(&inserted, &preview(&[(&id, Rejected)])), 1);

    let deleted = with_break(&EditCtx::local("", ""));
    deleted
        .doc()
        .delete_range(&suggest, StoryRange::new("body", 7, 8))
        .unwrap();
    let id = deleted.doc().list_revisions().unwrap()[0]
        .change
        .revision_id
        .clone();
    assert_eq!(pages(&deleted, &RenderEnv::default()), 2);
    assert_eq!(pages(&deleted, &preview(&[(&id, Accepted)])), 1);
    assert_eq!(pages(&deleted, &preview(&[(&id, Rejected)])), 2);
    assert_eq!(pages(&deleted, &RenderEnv::default()), 2);
}

#[test]
fn a_previewed_block_control_revision_covers_its_blocks() {
    let body = r#"<w:sdt><w:sdtPr><w:tag w:val="block"/></w:sdtPr><w:sdtContent><w:p><w:r><w:t>Inside</w:t></w:r></w:p></w:sdtContent></w:sdt><w:p><w:r><w:t>After</w:t></w:r></w:p>"#;
    let engine = EngineSession::new(75107);
    seed_from_docx(engine.doc(), &document(body)).unwrap();
    engine
        .doc()
        .delete_range(
            &EditCtx::local("Ann", "2026-09-29T12:00:00Z").suggesting(),
            StoryRange::new("body", 0, 1),
        )
        .unwrap();
    let id = engine.doc().list_revisions().unwrap()[0]
        .change
        .revision_id
        .clone();
    let shown = |env: &RenderEnv| lower(&engine, env).to_string();
    assert!(shown(&RenderEnv::default()).contains("Inside"));
    assert!(shown(&preview(&[(&id, Rejected)])).contains("Inside"));
    let accepted = shown(&preview(&[(&id, Accepted)]));
    assert!(!accepted.contains("Inside") && accepted.contains("After"));
}

#[test]
fn content_the_preview_leaves_out_does_not_count_in_lists() {
    let item = |text: &str| {
        format!(
            r#"<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>{text}</w:t></w:r></w:p>"#
        )
    };
    let control = format!(
        r#"<w:sdt><w:sdtPr><w:tag w:val="block"/></w:sdtPr><w:sdtContent>{}</w:sdtContent></w:sdt>"#,
        item("Inside")
    );
    let table = format!(
        r#"<w:tbl><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid><w:tr><w:tc>{}</w:tc></w:tr></w:tbl>"#,
        item("Inside")
    );
    for (client, container) in [(75108, control), (75109, table)] {
        let engine = EngineSession::new(client);
        seed_from_docx(
            engine.doc(),
            &document(&format!("{container}{}", item("After"))),
        )
        .unwrap();
        engine
            .doc()
            .delete_range(
                &EditCtx::local("Ann", "2026-09-29T12:00:00Z").suggesting(),
                StoryRange::new("body", 0, 1),
            )
            .unwrap();
        let revisions = engine.doc().list_revisions().unwrap();
        let id = revisions
            .iter()
            .find(|revision| revision.story == "body")
            .unwrap()
            .change
            .revision_id
            .clone();
        let after = |env: &RenderEnv| {
            let blocks = lower(&engine, env);
            let paragraph = blocks
                .as_array()
                .unwrap()
                .iter()
                .find(|block| block["runs"][0]["text"] == "After")
                .unwrap()
                .clone();
            paragraph["attrs"]["listMarker"].clone()
        };
        assert_eq!(after(&RenderEnv::default()), "2.", "{client}");
        assert_eq!(after(&preview(&[(&id, Rejected)])), "2.", "{client}");
        assert_eq!(after(&preview(&[(&id, Accepted)])), "1.", "{client}");
    }
}

#[test]
fn hidden_ranges_collapse_to_the_neighbouring_edges() {
    let font = docx_layout::register_measure_font(FONT).unwrap();
    let (engine, [replace, delete, _]) = proposals();
    let env = preview(&[(&replace, Accepted), (&delete, Accepted)]);
    engine
        .layout_document_with_regions_json(&layout_request(&env, font))
        .unwrap();
    engine.build_display_list_frame("{}", 0).unwrap();
    engine
        .with_display_list(|list| {
            assert!(docx_layout::hit::range_rects(list, 7, 11).is_empty());
            let before = docx_layout::hit::caret_rect(list, 7).unwrap();
            let after = docx_layout::hit::caret_rect(list, 11).unwrap();
            assert_eq!((before.page_index, before.y), (after.page_index, after.y));
            assert!((before.x - after.x).abs() < 0.01, "{before:?} {after:?}");
            let inserted = docx_layout::hit::range_rects(list, 11, 15);
            assert_eq!(inserted.len(), 1);
            assert!((inserted[0].x - after.x).abs() < 0.01);

            let empty_line = docx_layout::hit::caret_rect(list, 23).unwrap();
            let marks = docx_layout::hit::range_rects(list, 23, 28);
            assert_eq!(marks.len(), 1);
            assert_eq!((marks[0].x, marks[0].y), (empty_line.x, empty_line.y));
            assert!(docx_layout::hit::caret_rect(list, 25).is_none());
        })
        .unwrap();
}

#[test]
fn a_paged_export_refuses_a_previewed_layout() {
    let font = docx_layout::register_measure_font(FONT).unwrap();
    let bytes = document(&format!(
        r#"{PROPOSALS}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>"#
    ));
    let (engine, [replace, ..]) = proposals_in(&bytes);
    let mut request = fixture::region_request(&engine, &bytes, font);
    let options = PageExportOptions::new(RevisionView::Markup);
    let mut export = |env: RenderEnv| {
        request["renderEnv"] = serde_json::to_value(env).unwrap();
        engine
            .layout_document_with_regions_json(&request.to_string())
            .unwrap();
        engine
            .export_structured_with_pages(&options)
            .map(|_| ())
            .map_err(|refusal| refusal.failure.code)
    };
    assert_eq!(
        export(preview(&[(&replace, Accepted)])),
        Err(ExportFailureCode::UnsupportedRevisionLayout)
    );
    assert_eq!(export(RenderEnv::default()), Ok(()));
}

#[test]
fn a_suggested_replacement_that_ends_before_a_block_embed_stays_in_its_paragraph() {
    let table = r#"<w:tbl><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid><w:tr><w:tc><w:p w14:paraId="00000002"><w:r><w:t>cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl>"#;
    let suggest = |author: &str| EditCtx::local(author, "2026-09-29T12:00:00Z").suggesting();
    let first_runs = |engine: &EngineSession| -> Vec<(String, &'static str)> {
        runs(&lower(engine, &RenderEnv::default()), 0)
            .into_iter()
            .map(|run| (run.0, run.3))
            .collect()
    };

    // "old" and its paragraph mark, up to a table; Ann's own pending "Z" inside
    // the range is retracted without moving the insertion ahead of "ab".
    let engine = EngineSession::new(75112);
    let body = format!(
        r#"<w:p w14:paraId="00000001"><w:r><w:t>abold</w:t></w:r></w:p>{table}<w:p w14:paraId="00000003"><w:r><w:t>tail</w:t></w:r></w:p>"#
    );
    seed_from_docx(engine.doc(), &document(&body)).unwrap();
    engine
        .doc()
        .insert_text(
            &suggest("Ann"),
            Position::new("body", 2),
            "Z",
            FormatPolicy::Inherit,
        )
        .unwrap();
    engine
        .doc()
        .replace_range(&suggest("Ann"), StoryRange::new("body", 2, 7), "X")
        .unwrap();
    assert_eq!(
        first_runs(&engine),
        [
            ("ab".to_owned(), ""),
            ("X".to_owned(), "ins"),
            ("old".to_owned(), "del")
        ]
    );
    assert_eq!(lower(&engine, &RenderEnv::default())[1]["kind"], "table");

    // A story that ends with a page break after its last paragraph mark.
    let engine = EngineSession::new(75114);
    seed_from_docx(
        engine.doc(),
        &document(r#"<w:p w14:paraId="00000001"><w:r><w:t>old</w:t></w:r><w:r><w:br w:type="page"/></w:r></w:p>"#),
    )
    .unwrap();
    let len = engine.doc().story_len("body").unwrap();
    engine
        .doc()
        .replace_range(&suggest("Ann"), StoryRange::new("body", 0, len), "X")
        .unwrap();
    assert_eq!(
        first_runs(&engine),
        [("X".to_owned(), "ins"), ("old".to_owned(), "del")]
    );
}
