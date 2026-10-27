use docx_edit::*;
use serde_json::json;

const NS: &str = r#"xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml""#;
const DATE: &str = "2026-01-01T00:00:00Z";
const NUMBER: &str = r#"<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>"#;

fn p(id: usize, properties: &str, runs: &str) -> String {
    format!(r#"<w:p w14:paraId="{id:08X}"><w:pPr>{properties}</w:pPr>{runs}</w:p>"#)
}

fn run(text: &str) -> String {
    format!("<w:r><w:t>{text}</w:t></w:r>")
}

fn mark(kind: &str) -> String {
    format!(r#"<w:rPr><w:{kind} w:id="1" w:author="Author" w:date="{DATE}"/></w:rPr>"#)
}

fn tracked(kind: &str, text: &str) -> String {
    let tag = if kind == "del" { "delText" } else { "t" };
    format!(
        r#"<w:{kind} w:id="2" w:author="Author" w:date="{DATE}"><w:r><w:{tag}>{text}</w:{tag}></w:r></w:{kind}>"#
    )
}

fn table() -> String {
    format!(
        r#"<w:tbl><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid><w:tr><w:tc>{}</w:tc></w:tr></w:tbl>"#,
        p(90, "", &run("cell"))
    )
}

fn package(body: &str) -> Vec<u8> {
    let parts = vec![
        ("[Content_Types].xml".to_owned(), br#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/></Types>"#.to_vec()),
        ("_rels/.rels".to_owned(), br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="main" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"#.to_vec()),
        ("word/document.xml".to_owned(), format!("<w:document {NS}><w:body>{body}</w:body></w:document>").into_bytes()),
        ("word/_rels/document.xml.rels".to_owned(), br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="styles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="numbering" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/></Relationships>"#.to_vec()),
        ("word/styles.xml".to_owned(), format!(r#"<w:styles {NS}><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style><w:style w:type="paragraph" w:styleId="Heading1"><w:pPr><w:outlineLvl w:val="0"/></w:pPr></w:style></w:styles>"#).into_bytes()),
        ("word/numbering.xml".to_owned(), format!(r#"<w:numbering {NS}><w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>"#).into_bytes()),
    ];
    ooxml_opc::rezip_parts(&parts).unwrap()
}

fn open(bytes: &[u8]) -> EditingDoc {
    let doc = EditingDoc::new(7181);
    seed_from_docx(&doc, bytes).unwrap();
    doc
}

fn read(doc: &EditingDoc, view: EditTextView) -> Vec<ParagraphText> {
    doc.read_paragraphs(&ReadParagraphsRequest {
        story: None,
        para_ids: None,
        view,
    })
    .unwrap()
    .paragraphs
}

fn stories(doc: &EditingDoc, view: EditTextView) -> Vec<StoryText> {
    doc.read_stories(&ReadStoriesRequest {
        stories: None,
        by_root: false,
        view,
        expect_version: None,
    })
    .unwrap()
    .stories
}

fn oracle(bytes: &[u8], edit: impl Fn(&EditingDoc)) -> EditingDoc {
    let doc = open(bytes);
    edit(&doc);
    for view in [EditTextView::Accepted, EditTextView::Original] {
        let resolved = open(bytes);
        edit(&resolved);
        let ctx = EditCtx::local("Author", DATE);
        for story in resolved.list_stories().stories {
            let end = resolved.story_len(&story.story).unwrap();
            if end == 0 {
                continue;
            }
            let target = ChangeTarget::Range(StoryRange::new(&story.story, 0, end));
            match view {
                EditTextView::Accepted => resolved.accept_change(&ctx, &target),
                EditTextView::Original => resolved.reject_change(&ctx, &target),
            }
            .unwrap();
        }
        assert_eq!(read(&doc, view), read(&resolved, view));
        assert_eq!(
            serde_json::to_value(stories(&doc, view)).unwrap(),
            serde_json::to_value(stories(&resolved, view)).unwrap()
        );
    }
    doc
}

fn literal(doc: &EditingDoc, view: EditTextView, expected: &[(&str, &str)]) {
    let paragraphs = read(doc, view);
    let actual: Vec<_> = paragraphs
        .iter()
        .map(|p| (p.para_id.as_str(), p.text.as_str()))
        .collect();
    assert_eq!(actual, expected);
}

#[test]
fn package_split_merge_and_chain_follow_resolution_in_both_views() {
    for (kind, joined_view, split_view) in [
        ("ins", EditTextView::Original, EditTextView::Accepted),
        ("del", EditTextView::Accepted, EditTextView::Original),
    ] {
        let bytes = package(&format!(
            "{}{}{}",
            p(1, &mark(kind), &run("ab")),
            p(2, "", &run("cd")),
            p(3, "", &run("tail"))
        ));
        let doc = oracle(&bytes, |_| {});
        literal(
            &doc,
            joined_view,
            &[("00000002", "abcd"), ("00000003", "tail")],
        );
        literal(
            &doc,
            split_view,
            &[("00000001", "ab"), ("00000002", "cd"), ("00000003", "tail")],
        );
        let bytes = package(&format!(
            "{}{}{}",
            p(1, &mark(kind), &run("ab")),
            p(2, &mark(kind), &run("cd")),
            p(3, "", &run("ef"))
        ));
        let doc = oracle(&bytes, |_| {});
        literal(&doc, joined_view, &[("00000003", "abcdef")]);
        literal(
            &doc,
            split_view,
            &[("00000001", "ab"), ("00000002", "cd"), ("00000003", "ef")],
        );
    }
}

#[test]
fn package_insert_delete_both_stamps_and_final_marks_follow_resolution() {
    for kind in ["ins", "del"] {
        let bytes = package(&format!(
            "{}{}{}",
            p(1, &mark(kind), &tracked(kind, "new")),
            p(2, "", &run("kept")),
            p(3, &mark(kind), &run("last"))
        ));
        oracle(&bytes, |_| {});
    }
    let both =
        r#"<w:rPr><w:ins w:id="1" w:author="Author"/><w:del w:id="2" w:author="Author"/></w:rPr>"#;
    let doc = oracle(
        &package(&format!("{}{}", p(1, both, &run("a")), p(2, "", &run("b")))),
        |_| {},
    );
    for view in [EditTextView::Accepted, EditTextView::Original] {
        literal(&doc, view, &[("00000002", "ab")]);
    }
    let doc = oracle(&package(&p(1, &mark("del"), &run("last"))), |_| {});
    for view in [EditTextView::Accepted, EditTextView::Original] {
        literal(&doc, view, &[("00000001", "last")]);
    }
}

#[test]
fn block_boundaries_preserve_marks_only_when_content_precedes_them() {
    for kind in ["ins", "del"] {
        for content in [run("before"), tracked(kind, "hidden"), String::new()] {
            let bytes = package(&format!(
                "{}{}{}",
                p(1, &mark(kind), &content),
                table(),
                p(2, "", &run("after"))
            ));
            oracle(&bytes, |_| {});
        }
    }
}

#[test]
fn survivor_properties_and_numbering_counters_follow_resolution() {
    for kind in ["ins", "del"] {
        for first_numbered in [false, true] {
            let first = format!("{}{}", mark(kind), if first_numbered { NUMBER } else { "" });
            let second = format!(
                r#"<w:pStyle w:val="Heading1"/>{}"#,
                if first_numbered { "" } else { NUMBER }
            );
            let bytes = package(&format!(
                "{}{}{}",
                p(1, &first, &run("one")),
                p(2, &second, &run("two")),
                p(3, NUMBER, &run("three"))
            ));
            let doc = oracle(&bytes, |_| {});
            let view = if kind == "del" {
                EditTextView::Accepted
            } else {
                EditTextView::Original
            };
            let paragraphs = read(&doc, view);
            assert_eq!(paragraphs[0].para_id, "00000002");
            assert_eq!(paragraphs[0].style_id.as_deref(), Some("Heading1"));
            assert_eq!(paragraphs[0].outline_level, Some(0));
            assert_eq!(paragraphs[0].list.is_some(), !first_numbered);
            assert_eq!(
                paragraphs[1].list.as_ref().unwrap().marker.as_deref(),
                Some(if first_numbered { "1." } else { "2." })
            );
        }
    }
}

fn request(doc: &EditingDoc, operation: EditOperation, suggest: bool) -> EditRequest {
    let mut step = EditStep::new(operation);
    if suggest {
        step.suggest = Some(EditSuggestion {
            author: "Author".into(),
            date: DATE.into(),
        });
    }
    EditRequest {
        expect_version: doc.version(),
        source: EditSource::Host,
        history: EditHistory::Separate,
        steps: vec![step],
    }
}

fn engine_edit(doc: &EditingDoc, kind: &str) {
    let ctx = EditCtx::local("Author", DATE).suggesting();
    match kind {
        "split" => {
            doc.split_paragraph(&ctx, Position::new("body", 2), None)
                .unwrap();
        }
        "merge" => {
            doc.merge_paragraphs(&ctx, "00000001", MergeDirection::Forward)
                .unwrap();
        }
        "delete" => {
            doc.delete_range(&ctx, StoryRange::new("body", 0, 5))
                .unwrap();
        }
        "insert" => {
            doc.split_paragraph(&ctx, Position::new("body", 4), None)
                .unwrap();
            doc.insert_text(&ctx, Position::new("body", 5), "new", FormatPolicy::Plain)
                .unwrap();
        }
        _ => unreachable!(),
    }
}

#[test]
fn engine_made_split_merge_insert_and_delete_follow_resolution() {
    let bytes = package(&format!(
        "{}{}",
        p(1, "", &run("abcd")),
        p(2, "", &run("efgh"))
    ));
    for kind in ["split", "merge", "insert", "delete"] {
        let doc = oracle(&bytes, |doc| engine_edit(doc, kind));
        if kind == "merge" {
            literal(&doc, EditTextView::Accepted, &[("00000002", "abcdefgh")]);
            literal(
                &doc,
                EditTextView::Original,
                &[("00000001", "abcd"), ("00000002", "efgh")],
            );
        }
        if kind == "split" {
            let raw = doc.paragraphs("body").unwrap();
            assert_eq!(raw[0].para_id, "00000001");
            assert_eq!(
                read(&doc, EditTextView::Accepted)
                    .iter()
                    .map(|p| p.text.as_str())
                    .collect::<Vec<_>>(),
                ["ab", "cd", "efgh"]
            );
            let original = read(&doc, EditTextView::Original);
            assert_eq!(original[0].para_id, raw[1].para_id);
            assert_eq!(original[0].text, "abcd");
            assert_eq!(original[1].para_id, "00000002");
        }
    }
}

#[test]
fn subset_reads_and_search_use_survivor_ids_and_projected_offsets() {
    for (kind, view) in [
        ("del", EditTextView::Accepted),
        ("ins", EditTextView::Original),
    ] {
        let doc = open(&package(&format!(
            "{}{}{}",
            p(1, &mark(kind), &run("ab")),
            p(2, "", &run("cd")),
            p(3, "", &run("tail"))
        )));
        let all = read(&doc, view);
        for (id, expected) in [("00000002", &all[0]), ("00000003", &all[1])] {
            let subset = doc
                .read_paragraphs(&ReadParagraphsRequest {
                    story: None,
                    para_ids: Some(vec![id.into()]),
                    view,
                })
                .unwrap();
            assert_eq!(subset.paragraphs, vec![expected.clone()]);
        }
        let refusal = doc
            .read_paragraphs(&ReadParagraphsRequest {
                story: None,
                para_ids: Some(vec!["00000001".into()]),
                view,
            })
            .unwrap_err();
        assert_eq!(refusal.failure.code, EditFailureCode::MissingTarget);
        assert!(
            refusal
                .failure
                .message
                .contains("joins the following paragraph")
        );
        assert!(
            refusal
                .failure
                .message
                .to_lowercase()
                .contains(if kind == "del" {
                    "accepted"
                } else {
                    "original"
                })
        );
        for within in [
            SearchScope::Story {
                story: "body".into(),
            },
            SearchScope::Paragraph(ParagraphTarget {
                story: "body".into(),
                para_id: "00000002".into(),
            }),
        ] {
            let found = doc
                .find_text(&FindTextRequest {
                    text: "bc".into(),
                    within,
                    view,
                    limit: None,
                })
                .unwrap();
            assert_eq!(found.matches.len(), 1);
            let range = &found.matches[0].range;
            assert_eq!(range.start.para_id, "00000002");
            assert_eq!(range.end.para_id, "00000002");
            assert_eq!((range.start.offset, range.end.offset), (1, 3));
        }
    }
}

#[test]
fn projected_read_ranges_and_search_edits_are_refused_without_mutation() {
    for (kind, view) in [
        ("del", EditTextView::Accepted),
        ("ins", EditTextView::Original),
    ] {
        let doc = open(&package(&format!(
            "{}{}",
            p(1, &mark(kind), &run("abc")),
            p(2, "", &run("defgh"))
        )));
        let paragraph = &read(&doc, view)[0];
        assert_eq!(paragraph.text, "abcdefgh");
        let mut targets = Vec::new();
        for (start, end) in [(1, 2), (4, 5), (6, 7)] {
            targets.push(TextTarget::Range(TextRange {
                story: paragraph.story.clone(),
                start: TextPosition {
                    para_id: paragraph.para_id.clone(),
                    offset: start,
                },
                end: TextPosition {
                    para_id: paragraph.para_id.clone(),
                    offset: end,
                },
                view,
            }));
        }
        targets.push(TextTarget::Search {
            text: "cd".into(),
            within: SearchScope::Story {
                story: "body".into(),
            },
            view,
        });
        for target in targets {
            let expected = match &target {
                TextTarget::Range(range) if range.end.offset > 5 => EditFailureCode::InvalidStep,
                _ => EditFailureCode::TrackedRevisionConflict,
            };
            let before = doc.encode_state_as_update_v1();
            let request = request(
                &doc,
                EditOperation::ReplaceText {
                    target,
                    text: "X".into(),
                },
                false,
            );
            assert_eq!(
                doc.validate_edits(&request)
                    .unwrap()
                    .unwrap_err()
                    .failure
                    .code,
                expected
            );
            assert_eq!(doc.encode_state_as_update_v1(), before);
            assert_eq!(
                doc.apply_edits(&request, &UndoSession::new())
                    .unwrap()
                    .unwrap_err()
                    .failure
                    .code,
                expected
            );
            assert_eq!(doc.encode_state_as_update_v1(), before);
        }
    }
}

#[test]
fn anchors_still_resolve_raw_survivor_and_joined_away_marks() {
    for kind in ["split", "merge"] {
        let doc = open(&package(&format!(
            "{}{}",
            p(1, "", &run("abcd")),
            p(2, "", &run("efgh"))
        )));
        engine_edit(&doc, kind);
        let raw = doc.paragraphs("body").unwrap();
        for paragraph in &raw[..2] {
            let anchor = ParagraphAnchor::Session {
                session_id: doc.session_id(),
                story: "body".into(),
                para_id: paragraph.para_id.clone(),
            };
            assert_eq!(
                doc.resolve_paragraph_anchor(&anchor),
                AnchorResolution::Found(ParagraphRef::Session {
                    story: "body".into(),
                    para_id: paragraph.para_id.clone()
                })
            );
        }
    }
}

#[test]
fn joined_atom_offsets_are_shifted_in_utf16_units() {
    let bytes = package(&format!(
        "{}{}",
        p(1, &mark("del"), &format!("{}<w:r><w:br/></w:r>", run("😀"))),
        p(2, "", &format!("{}<w:r><w:br/></w:r>", run("b")))
    ));
    let doc = oracle(&bytes, |_| {});
    let paragraphs = read(&doc, EditTextView::Accepted);
    assert_eq!(paragraphs[0].text, "😀\u{fffc}b\u{fffc}");
    assert_eq!(
        paragraphs[0]
            .atoms
            .iter()
            .map(|a| a.offset)
            .collect::<Vec<_>>(),
        [2, 4]
    );
}

#[test]
fn run_revisions_without_mark_revisions_keep_literal_read_records() {
    let bytes = package(&format!(
        "{}{}",
        p(
            1,
            r#"<w:pStyle w:val="Normal"/>"#,
            &format!("{}{}{}", run("a"), tracked("ins", "b"), tracked("del", "c"))
        ),
        p(2, r#"<w:pStyle w:val="Normal"/>"#, &run("tail"))
    ));
    let doc = oracle(&bytes, |_| {});
    for (view, text) in [
        (EditTextView::Accepted, "ab"),
        (EditTextView::Original, "ac"),
    ] {
        let value = serde_json::to_value(read(&doc, view)).unwrap();
        assert_eq!(
            value,
            json!([
                {"story":"body", "paraId":"00000001", "text":text, "styleId":"Normal", "outlineLevel":null, "list":null, "atoms":[]},
                {"story":"body", "paraId":"00000002", "text":"tail", "styleId":"Normal", "outlineLevel":null, "list":null, "atoms":[]}
            ])
        );
    }
}
