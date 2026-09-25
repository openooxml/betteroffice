//! One export feature per package: numbering counters, source grids, revisions, field results,
//! comments, shared identities, breaks, relationships, placeholders, Markdown tables and limits.

#[allow(dead_code)]
#[path = "support/structured_fixture.rs"]
mod fixture;

use docx_edit::structured::{
    Anchor, Block, BlockKind, CachedResult, DiagnosticCode, DocxStructuredContent, ExportOptions,
    FormattingMark, Inline, InlineKind, MarkdownOptions, RevisionKind, RevisionView, StoryKind,
    StorySelection, VerticalMerge, export_docx_markdown, export_docx_structured,
    render_docx_markdown,
};
use docx_edit::{EditCtx, EditingDoc, RawOp, seed_from_docx};
use fixture::{Package, image, para, run};
use yrs::{Map, MapRef, ReadTxn, Transact};

const COMMENTS: &str = "comments";

fn options(view: RevisionView) -> ExportOptions {
    ExportOptions::new(view)
}

fn with_stories(view: RevisionView, stories: &[StorySelection]) -> ExportOptions {
    ExportOptions {
        stories: Some(stories.to_vec()),
        ..options(view)
    }
}

fn export(bytes: &[u8], options: &ExportOptions) -> DocxStructuredContent {
    export_docx_structured(bytes, options).unwrap()
}

fn open(bytes: &[u8]) -> EditingDoc {
    let doc = EditingDoc::new(8080);
    seed_from_docx(&doc, bytes).unwrap();
    doc
}

fn body(content: &DocxStructuredContent) -> &[Block] {
    &content.stories[0].blocks
}

fn inlines(block: &Block) -> &[Inline] {
    match &block.content {
        BlockKind::Paragraph { paragraph }
        | BlockKind::Heading { paragraph, .. }
        | BlockKind::ListItem { paragraph, .. } => &paragraph.inlines,
        other => panic!("not a paragraph: {other:?}"),
    }
}

fn label(inline: &Inline) -> String {
    match &inline.content {
        InlineKind::Text { text } => text.clone(),
        InlineKind::Tab => "<tab>".to_owned(),
        InlineKind::Break { break_type } => format!("<{break_type:?}>"),
        InlineKind::Image { .. } => "<image>".to_owned(),
        InlineKind::Field { .. } => "<field>".to_owned(),
        InlineKind::Unsupported { element, .. } => format!("<{element}>"),
        other => format!("{other:?}"),
    }
}

fn labels(block: &Block) -> Vec<String> {
    inlines(block).iter().map(label).collect()
}

fn paragraph<'a>(content: &'a DocxStructuredContent, para_id: &str) -> &'a Block {
    body(content)
        .iter()
        .find(|block| {
            matches!(&block.anchor, Anchor::Paragraph { para_id: id, .. } if id == para_id)
                && paragraph_like(block)
        })
        .unwrap_or_else(|| panic!("no paragraph {para_id}"))
}

fn paragraph_like(block: &Block) -> bool {
    matches!(
        block.content,
        BlockKind::Paragraph { .. } | BlockKind::Heading { .. } | BlockKind::ListItem { .. }
    )
}

fn count(content: &DocxStructuredContent, code: DiagnosticCode) -> usize {
    content
        .diagnostics
        .iter()
        .filter(|diagnostic| diagnostic.code == code)
        .count()
}

fn numbered(id: &str, num: u32, level: u32, text: &str) -> String {
    format!(
        r#"<w:p w14:paraId="{id}"><w:pPr><w:numPr><w:ilvl w:val="{level}"/><w:numId w:val="{num}"/></w:numPr></w:pPr>{}</w:p>"#,
        run(text)
    )
}

fn numbering_part(abstracts: &str, instances: &str) -> String {
    format!(
        r#"<w:numbering {}>{abstracts}{instances}</w:numbering>"#,
        fixture::namespaces()
    )
}

fn markers(content: &DocxStructuredContent) -> Vec<Option<String>> {
    body(content)
        .iter()
        .filter_map(|block| match &block.content {
            BlockKind::ListItem { list, .. } => Some(list.marker.clone()),
            _ => None,
        })
        .collect()
}

#[test]
fn list_counters_follow_starts_overrides_and_restarts() {
    let abstracts = concat!(
        r#"<w:abstractNum w:abstractNumId="0">"#,
        r#"<w:lvl w:ilvl="0"><w:start w:val="5"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl>"#,
        r#"<w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="lowerLetter"/><w:lvlText w:val="%2."/><w:lvlRestart w:val="0"/></w:lvl>"#,
        r#"<w:lvl w:ilvl="2"><w:start w:val="1"/><w:numFmt w:val="lowerRoman"/><w:lvlText w:val="%3."/></w:lvl>"#,
        r#"</w:abstractNum>"#
    );
    let instances = concat!(
        r#"<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>"#,
        r#"<w:num w:numId="2"><w:abstractNumId w:val="0"/></w:num>"#,
        r#"<w:num w:numId="3"><w:abstractNumId w:val="0"/><w:lvlOverride w:ilvl="0"><w:startOverride w:val="1"/></w:lvlOverride></w:num>"#
    );
    let paragraphs = [
        (1, 0),
        (1, 1),
        (1, 2),
        (1, 0),
        (1, 1),
        (1, 2),
        (2, 0),
        (3, 0),
        (3, 0),
        (1, 0),
    ];
    let xml: String = paragraphs
        .iter()
        .enumerate()
        .map(|(index, (num, level))| {
            numbered(
                &format!("1000{index:04}"),
                *num,
                *level,
                &format!("Item {index}"),
            )
        })
        .collect();
    let bytes = Package::new(&xml)
        .numbering(&numbering_part(abstracts, instances))
        .bytes();
    let content = export(&bytes, &options(RevisionView::Accepted));
    let expected = ["5.", "a.", "i.", "6.", "b.", "i.", "7.", "1.", "2.", "8."];
    assert_eq!(
        markers(&content),
        expected.map(|marker| Some(marker.to_owned()))
    );
    let doc = open(&bytes);
    let rendered: Vec<Option<String>> =
        docx_edit::bridge::yrs_doc_to_layout_blocks(&doc, "body", &Default::default())
            .unwrap()
            .into_iter()
            .filter_map(|block| match block {
                docx_layout::types::LayoutBlock::Paragraph(paragraph) => {
                    Some(paragraph.attrs?.list_marker)
                }
                _ => None,
            })
            .collect();
    assert_eq!(
        rendered,
        markers(&content),
        "the renderer reads the same markers"
    );
}

#[test]
fn invalid_numbering_levels_are_diagnosed_without_counters() {
    let abstracts = r#"<w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl></w:abstractNum>"#;
    let xml = [
        numbered("20000001", 1, 12, "Too deep"),
        numbered("20000002", 1, 0, "First"),
    ]
    .concat();
    let bytes = Package::new(&xml)
        .numbering(&numbering_part(
            abstracts,
            r#"<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>"#,
        ))
        .bytes();
    let content = export(&bytes, &options(RevisionView::Accepted));
    let BlockKind::ListItem { list, .. } = &paragraph(&content, "20000001").content else {
        panic!("a numbered paragraph stays a list item");
    };
    assert_eq!((list.level, list.marker.as_deref()), (8, None));
    assert!(content.diagnostics.iter().any(|diagnostic| {
        diagnostic.code == DiagnosticCode::UnsupportedNumbering
            && diagnostic.message.contains("outside 0..8")
    }));
    assert_eq!(markers(&content)[1].as_deref(), Some("1."));
}

fn tables(content: &DocxStructuredContent) -> Vec<&docx_edit::structured::TableData> {
    body(content)
        .iter()
        .filter_map(|block| match &block.content {
            BlockKind::Table { table } => Some(table),
            _ => None,
        })
        .collect()
}

#[test]
fn vertical_merges_follow_the_source_grid() {
    let cell = |properties: &str, content: &str| {
        format!(r#"<w:tc><w:tcPr>{properties}</w:tcPr>{content}</w:tc>"#)
    };
    let column = concat!(
        r#"<w:tbl><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid>"#,
        r#"<w:tr><w:tc><w:tcPr><w:vMerge w:val="restart"/></w:tcPr><w:p w14:paraId="30000001"><w:r><w:t>Top</w:t></w:r></w:p></w:tc></w:tr>"#,
        r#"<w:tr><w:tc><w:tcPr><w:vMerge/></w:tcPr><w:p w14:paraId="30000002"><w:r><w:t>Hidden</w:t></w:r></w:p></w:tc></w:tr>"#,
        r#"</w:tbl>"#
    );
    let shifted = format!(
        r#"<w:tbl><w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid><w:tr>{}{}</w:tr><w:tr><w:trPr><w:gridBefore w:val="1"/></w:trPr>{}</w:tr></w:tbl>"#,
        cell("", &para("30000003", &run("Left"))),
        cell(
            r#"<w:vMerge w:val="restart"/>"#,
            &para("30000004", &run("Right"))
        ),
        cell(r#"<w:vMerge/>"#, &para("30000005", "")),
    );
    let xml = format!(
        "{column}{}{shifted}{}",
        para("30000006", ""),
        para("30000007", "")
    );
    let content = export(
        &Package::new(&xml).bytes(),
        &options(RevisionView::Accepted),
    );
    let found = tables(&content);
    type Grid = Vec<Vec<(u32, u32, VerticalMerge, Option<(u32, u32)>)>>;
    let summary = |table: &docx_edit::structured::TableData| -> Grid {
        table
            .rows
            .iter()
            .map(|row| {
                row.cells
                    .iter()
                    .map(|cell| {
                        (
                            cell.column,
                            cell.row_span,
                            cell.vertical_merge,
                            cell.merge_origin.map(|origin| (origin.row, origin.column)),
                        )
                    })
                    .collect()
            })
            .collect()
    };
    assert_eq!(
        summary(found[0]),
        [
            vec![(0, 2, VerticalMerge::Restart, None)],
            vec![(0, 0, VerticalMerge::Continue, Some((0, 0)))],
        ]
    );
    assert_eq!(
        summary(found[1]),
        [
            vec![
                (0, 1, VerticalMerge::None, None),
                (1, 2, VerticalMerge::Restart, None)
            ],
            vec![(1, 0, VerticalMerge::Continue, Some((0, 1)))],
        ]
    );
    assert_eq!(found[1].rows[1].grid_before, 1);
    assert_eq!(
        count(&content, DiagnosticCode::MergeContinuationContentOmitted),
        1,
        "the fully continuing row's text is omitted and diagnosed"
    );
}

#[test]
fn moves_and_tracked_drawings_follow_every_view() {
    let xml = format!(
        r#"<w:p w14:paraId="40000001">{}<w:moveFrom w:id="31" w:author="Ann" w:date="2026-02-01T00:00:00Z"><w:r><w:delText>moved</w:delText></w:r></w:moveFrom><w:moveTo w:id="32" w:author="Ann" w:date="2026-02-01T00:00:00Z"><w:r><w:t>placed</w:t></w:r></w:moveTo><w:del w:id="33" w:author="Bob" w:date="2026-02-02T00:00:00Z">{}</w:del></w:p>"#,
        run("Keep "),
        image("rIdImage", "Removed picture")
    );
    let bytes = Package::new(&xml)
        .rel("rIdImage", "image", "media/image1.png")
        .bytes();
    let view = |view| labels(&body(&export(&bytes, &options(view)))[0]);
    assert_eq!(view(RevisionView::Accepted), ["Keep placed"]);
    assert_eq!(view(RevisionView::Original), ["Keep moved", "<image>"]);
    let markup = export(&bytes, &options(RevisionView::Markup));
    let attributed: Vec<(String, Vec<RevisionKind>)> = inlines(&body(&markup)[0])
        .iter()
        .map(|inline| {
            (
                label(inline),
                inline
                    .revisions
                    .iter()
                    .map(|revision| revision.kind)
                    .collect(),
            )
        })
        .collect();
    assert_eq!(
        attributed,
        [
            ("Keep ".to_owned(), vec![]),
            ("moved".to_owned(), vec![RevisionKind::MoveFrom]),
            ("placed".to_owned(), vec![RevisionKind::MoveTo]),
            ("<image>".to_owned(), vec![RevisionKind::Deletion]),
        ]
    );
    let session = open(&bytes);
    let live = session
        .export_structured(&options(RevisionView::Accepted))
        .unwrap()
        .content;
    assert_eq!(labels(&body(&live)[0]), ["Keep placed"]);
}

fn field_results(block: &Block) -> Vec<(String, CachedResult)> {
    inlines(block)
        .iter()
        .filter_map(|inline| match &inline.content {
            InlineKind::Field {
                field_type,
                cached_result,
                ..
            } => Some((field_type.clone(), cached_result.clone())),
            _ => None,
        })
        .collect()
}

fn field_anchor(block: &Block) -> Anchor {
    inlines(block)
        .iter()
        .find(|inline| matches!(inline.content, InlineKind::Field { .. }))
        .unwrap()
        .anchor
        .clone()
}

fn complex_field(instruction: &str, result: &str) -> String {
    format!(
        r#"<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve">{instruction}</w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>{result}<w:r><w:fldChar w:fldCharType="end"/></w:r>"#
    )
}

#[test]
fn structured_field_results_keep_runs_tabs_and_images() {
    let xml = [
        para(
            "50000001",
            &complex_field(
                " REF _Ref1 \\h ",
                r#"<w:r><w:rPr><w:b/></w:rPr><w:t>Bold</w:t></w:r><w:r><w:tab/><w:t>plain</w:t></w:r>"#,
            ),
        ),
        para(
            "50000002",
            &complex_field(" INCLUDEPICTURE \"x\" ", &image("rIdImage", "Pictured")),
        ),
    ]
    .concat();
    let bytes = Package::new(&xml)
        .rel("rIdImage", "image", "media/image1.png")
        .bytes();
    for content in [
        export(&bytes, &options(RevisionView::Accepted)),
        open(&bytes)
            .export_structured(&options(RevisionView::Accepted))
            .unwrap()
            .content,
    ] {
        let first = paragraph(&content, "50000001");
        let anchor = field_anchor(first);
        let results = field_results(first);
        let [(_, CachedResult::Inline { inlines })] = results.as_slice() else {
            panic!("an inline cached result");
        };
        let result: Vec<(String, Option<Vec<FormattingMark>>)> = inlines
            .iter()
            .map(|inline| (label(inline), inline.marks.clone()))
            .collect();
        assert_eq!(
            result,
            [
                ("Bold".to_owned(), Some(vec![FormattingMark::Bold])),
                ("<tab>".to_owned(), Some(vec![])),
                ("plain".to_owned(), Some(vec![])),
            ]
        );
        assert!(inlines.iter().all(|inline| inline.anchor == anchor));
        let second = paragraph(&content, "50000002");
        let results = field_results(second);
        let [(_, CachedResult::Inline { inlines })] = results.as_slice() else {
            panic!("an inline cached result");
        };
        assert!(matches!(
            &inlines[0].content,
            InlineKind::Image { alt_text: Some(alt), part: Some(part), .. }
                if alt == "Pictured" && part == "word/media/image1.png"
        ));
        let anchor = field_anchor(second);
        assert!(content.diagnostics.iter().any(|diagnostic| {
            diagnostic.code == DiagnosticCode::ImageDataOmitted
                && diagnostic.anchor.as_ref() == Some(&anchor)
        }));
    }
}

#[test]
fn numeric_field_result_blocks_stay_inside_their_field() {
    let xml = [
        para(
            "51000001",
            &format!(
                r#"{}<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> 7 </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>first</w:t></w:r>"#,
                run("Lead ")
            ),
        ),
        para("51000002", &run("second")),
        para("51000003", r#"<w:r><w:fldChar w:fldCharType="end"/></w:r>"#),
        para("51000004", &run("after")),
    ]
    .concat();
    let bytes = Package::new(&xml).bytes();
    let content = export(&bytes, &options(RevisionView::Accepted));
    let owner = paragraph(&content, "51000001");
    let anchor = field_anchor(owner);
    let results = field_results(owner);
    let [(_, CachedResult::Blocks { blocks })] = results.as_slice() else {
        panic!("a block cached result: {:?}", field_results(owner));
    };
    let texts: Vec<String> = blocks.iter().map(|block| labels(block).concat()).collect();
    assert_eq!(texts, ["first", "second", ""]);
    assert!(blocks.iter().all(|block| block.anchor == anchor));
    assert!(
        body(&content)
            .iter()
            .all(|block| !matches!(&block.anchor, Anchor::Paragraph { para_id, .. } if para_id == "51000002")),
        "result blocks are not exported twice"
    );
    let markdown = render_docx_markdown(&content, &MarkdownOptions::default()).unwrap();
    let lead = markdown.markdown.find("Lead").unwrap();
    let second = markdown.markdown.find("second").unwrap();
    let after = markdown.markdown.find("after").unwrap();
    assert!(lead < second && second < after, "{}", markdown.markdown);
    assert_eq!(
        markdown.anchors.len(),
        markdown.markdown.matches("<!-- docx-export:").count()
    );
    assert!(markdown.anchors.iter().any(|entry| entry.anchor == anchor));
}

fn comment_package(done: bool) -> Package {
    let xml = para(
        "60000001",
        &format!(
            r#"<w:commentRangeStart w:id="1"/>{}<w:commentRangeEnd w:id="1"/><w:r><w:commentReference w:id="1"/></w:r>"#,
            run("Annotated")
        ),
    );
    let comments = format!(
        r#"<w:comments {}><w:comment w:id="1" w:author="Ann" w:date="2026-03-01T00:00:00Z"{}>{}</w:comment></w:comments>"#,
        fixture::namespaces(),
        if done { r#" w:done="1""# } else { "" },
        para("60000002", &run("Resolved remark"))
    );
    Package::new(&xml).part("comments.xml", "rIdComments", COMMENTS, COMMENTS, &comments)
}

fn comment_metadata(content: &DocxStructuredContent) -> docx_edit::structured::CommentMetadata {
    content
        .stories
        .iter()
        .find_map(|story| story.comment.clone())
        .unwrap()
}

#[test]
fn current_comment_values_win_over_the_source() {
    let bytes = comment_package(true).bytes();
    let options = with_stories(RevisionView::Accepted, &[StorySelection::Comments]);
    assert!(comment_metadata(&export(&bytes, &options)).resolved);
    let doc = open(&bytes);
    let read =
        |doc: &EditingDoc| comment_metadata(&doc.export_structured(&options).unwrap().content);
    let seeded = read(&doc);
    assert!(seeded.resolved, "an unwritten store field reads the source");
    assert_eq!(seeded.author.as_deref(), Some("Ann"));
    {
        let mut txn = doc.yrs_doc().transact_mut();
        let comments = txn.get_map("comments").unwrap();
        let comment: MapRef = comments.get(&txn, "1").unwrap().cast().unwrap();
        comment.insert(&mut txn, "done", false);
        comment.insert(&mut txn, "author", "");
        comment.insert(&mut txn, "date", "2027-01-01T00:00:00Z");
    }
    let current = read(&doc);
    assert!(
        !current.resolved,
        "a reopened comment is no longer resolved"
    );
    assert_eq!(current.author, None, "a cleared author stays cleared");
    assert_eq!(current.date.as_deref(), Some("2027-01-01T00:00:00Z"));
    let unresolved = open(&comment_package(false).bytes());
    {
        let mut txn = unresolved.yrs_doc().transact_mut();
        let comments = txn.get_map("comments").unwrap();
        let comment: MapRef = comments.get(&txn, "1").unwrap().cast().unwrap();
        comment.insert(&mut txn, "done", true);
    }
    assert!(read(&unresolved).resolved);
}

#[test]
fn shared_paragraph_ids_never_look_like_edit_targets() {
    let xml = [
        para("70000001", &run("First twin")),
        para("70000001", &run("Second twin")),
        para("70000002", &run("Unique")),
    ]
    .concat();
    let bytes = Package::new(&xml).bytes();
    let content = export(&bytes, &options(RevisionView::Accepted));
    let ids: Vec<String> = body(&content)
        .iter()
        .map(|block| match &block.anchor {
            Anchor::Paragraph { para_id, .. } => para_id.clone(),
            other => panic!("{other:?}"),
        })
        .collect();
    assert_eq!(ids[0], "70000001");
    assert_ne!(
        ids[1], "70000001",
        "parsing gives a repeated id a fresh one"
    );
    assert_eq!(count(&content, DiagnosticCode::AmbiguousIdentity), 0);
    let doc = open(&bytes);
    let pilcrow = doc.paragraph_mark_position(&ids[1]).unwrap();
    doc.apply_raw_ops(
        "body",
        vec![RawOp::SetEmbedAttr {
            index: pilcrow.index,
            key: "paraId".to_owned(),
            value: yrs::Any::from("70000001"),
        }],
        &EditCtx::local("", ""),
    )
    .unwrap();
    let content = doc
        .export_structured(&options(RevisionView::Accepted))
        .unwrap()
        .content;
    let unlocated = Anchor::Paragraph {
        story: "body".to_owned(),
        para_id: String::new(),
    };
    for block in &body(&content)[..2] {
        assert_eq!(block.anchor, unlocated);
        assert!(
            inlines(block)
                .iter()
                .all(|inline| inline.anchor == unlocated)
        );
    }
    assert!(
        matches!(&body(&content)[2].anchor, Anchor::Paragraph { para_id, .. } if para_id == "70000002")
    );
    assert_eq!(count(&content, DiagnosticCode::AmbiguousIdentity), 2);
}

/// A package whose comment body holds a cross-reference with two differently formatted
/// cached runs.
fn field_comment_package() -> Package {
    let xml = para(
        "61000001",
        &format!(
            r#"<w:commentRangeStart w:id="1"/>{}<w:commentRangeEnd w:id="1"/><w:r><w:commentReference w:id="1"/></w:r>"#,
            run("Annotated")
        ),
    );
    let comments = format!(
        r#"<w:comments {}><w:comment w:id="1" w:author="Ann"><w:p w14:paraId="61000002"><w:fldSimple w:instr=" REF Summary "><w:r><w:rPr><w:b/></w:rPr><w:t>First</w:t></w:r><w:r><w:t>Second</w:t></w:r></w:fldSimple></w:p></w:comment></w:comments>"#,
        fixture::namespaces()
    );
    Package::new(&xml).part("comments.xml", "rIdComments", COMMENTS, COMMENTS, &comments)
}

#[test]
fn multi_run_fields_inside_comments_export() {
    let (sender, receiver) = std::sync::mpsc::channel();
    std::thread::spawn(move || {
        let bytes = field_comment_package().bytes();
        let options = with_stories(RevisionView::Accepted, &[StorySelection::Comments]);
        sender
            .send([
                export(&bytes, &options),
                open(&bytes).export_structured(&options).unwrap().content,
            ])
            .unwrap();
    });
    let exports = receiver
        .recv_timeout(std::time::Duration::from_secs(60))
        .expect("both exports finish");
    for content in exports {
        let story = content
            .stories
            .iter()
            .find(|story| story.kind == StoryKind::Comment)
            .unwrap();
        let results = field_results(&story.blocks[0]);
        let [(_, CachedResult::Inline { inlines })] = results.as_slice() else {
            panic!("an inline cached result: {results:?}");
        };
        let result: Vec<(String, Option<Vec<FormattingMark>>)> = inlines
            .iter()
            .map(|inline| (label(inline), inline.marks.clone()))
            .collect();
        assert_eq!(
            result,
            [
                ("First".to_owned(), Some(vec![FormattingMark::Bold])),
                ("Second".to_owned(), Some(vec![])),
            ]
        );
    }
}

#[test]
fn comment_bodies_keep_their_omissions() {
    let marker = r#"<bofx:mark bofx:value="kept"/>"#;
    let xml = para(
        "80000001",
        &format!(
            r#"<w:commentRangeStart w:id="4"/>{}<w:commentRangeEnd w:id="4"/><w:r><w:commentReference w:id="4"/></w:r>"#,
            run("Noted")
        ),
    );
    let comments = format!(
        r#"<w:comments {}><w:comment w:id="4" w:author="Ann"><bofx:block bofx:value="opaque"/>{}</w:comment></w:comments>"#,
        fixture::namespaces(),
        para("80000002", &format!("{}{marker}", run("See ")))
    );
    let bytes = Package::new(&xml)
        .part("comments.xml", "rIdComments", COMMENTS, COMMENTS, &comments)
        .bytes();
    let options = with_stories(RevisionView::Accepted, &[StorySelection::Comments]);
    for content in [
        export(&bytes, &options),
        open(&bytes).export_structured(&options).unwrap().content,
    ] {
        let story = content
            .stories
            .iter()
            .find(|story| story.kind == StoryKind::Comment)
            .unwrap();
        assert!(matches!(
            &story.blocks[0],
            Block { content: BlockKind::Unsupported { element }, anchor: Anchor::SourcePart { part, path, .. }, .. }
                if element == "bofx:block" && part == "word/comments.xml" && path == &[0, 0]
        ));
        assert_eq!(labels(&story.blocks[1]), ["See ", "<bofx:mark>"]);
        assert_eq!(count(&content, DiagnosticCode::UnsupportedContent), 2);
    }
}

#[test]
fn images_resolve_against_the_part_that_owns_them() {
    let header = format!(
        r#"<w:hdr {}>{}</w:hdr>"#,
        fixture::namespaces(),
        para("90000001", &image("rIdLogo", "Header logo"))
    );
    let xml = format!(
        r#"{}<w:sectPr><w:headerReference w:type="default" r:id="rIdHeader"/></w:sectPr>"#,
        para("90000002", &image("rIdMissing", "Lost"))
    );
    let bytes = Package::new(&xml)
        .part("header1.xml", "rIdHeader", "header", "header", &header)
        .part_rel("header1.xml", "rIdLogo", "image", "media/image1.png")
        .bytes();
    let content = export(
        &bytes,
        &with_stories(
            RevisionView::Accepted,
            &[StorySelection::Body, StorySelection::Headers],
        ),
    );
    let image_part = |block: &Block| {
        inlines(block)
            .iter()
            .find_map(|inline| match &inline.content {
                InlineKind::Image { part, .. } => Some(part.clone()),
                _ => None,
            })
    };
    let header = content
        .stories
        .iter()
        .find(|story| story.kind == StoryKind::Header)
        .unwrap();
    assert_eq!(
        image_part(&header.blocks[0]),
        Some(Some("word/media/image1.png".to_owned()))
    );
    assert_eq!(image_part(&body(&content)[0]), Some(None));
    assert!(content.diagnostics.iter().any(|diagnostic| {
        diagnostic.code == DiagnosticCode::UnresolvedReference
            && diagnostic.message.contains("rIdMissing")
    }));
}

fn page_break() -> &'static str {
    r#"<w:r><w:br w:type="page"/></w:r>"#
}

#[test]
fn page_and_column_breaks_stay_where_the_source_has_them() {
    let cell = format!(
        r#"<w:tbl><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid><w:tr><w:tc>{}</w:tc></w:tr></w:tbl>"#,
        para(
            "0A000004",
            &format!("{}{}{}", run("Cell"), page_break(), run("text"))
        )
    );
    let xml = [
        para(
            "0A000001",
            &format!("{}{}{}", run("Before"), page_break(), run("After")),
        ),
        para("0A000002", &format!("{}{}", page_break(), run("Leading"))),
        para(
            "0A000003",
            &format!(
                "{}{}",
                run("Column"),
                r#"<w:r><w:br w:type="column"/></w:r>"#
            ),
        ),
        cell,
        para(
            "0A000005",
            &format!(
                r#"{}<w:ins w:id="51" w:author="Ann"><w:r><w:br w:type="page"/></w:r></w:ins>{}"#,
                run("Tracked"),
                run("break")
            ),
        ),
        para("0A000006", &run("End")),
    ]
    .concat();
    let bytes = Package::new(&xml).bytes();
    let doc = open(&bytes);
    for content in [
        export(&bytes, &options(RevisionView::Markup)),
        doc.export_structured(&options(RevisionView::Markup))
            .unwrap()
            .content,
    ] {
        assert_eq!(
            labels(paragraph(&content, "0A000001")),
            ["Before", "<Page>", "After"]
        );
        assert_eq!(
            labels(paragraph(&content, "0A000002")),
            ["<Page>", "Leading"]
        );
        assert_eq!(
            labels(paragraph(&content, "0A000003")),
            ["Column", "<Column>"]
        );
        assert!(
            body(&content)
                .iter()
                .all(|block| !matches!(block.content, BlockKind::Break { .. })),
            "relocated breaks are not exported again between blocks"
        );
        let table = tables(&content)[0];
        assert_eq!(
            labels(&table.rows[0].cells[0].blocks[0]),
            ["Cell", "<Page>", "text"]
        );
        let tracked = inlines(paragraph(&content, "0A000005"));
        assert_eq!(
            tracked[1]
                .revisions
                .iter()
                .map(|revision| revision.kind)
                .collect::<Vec<_>>(),
            [RevisionKind::Insertion]
        );
    }
    assert_eq!(
        labels(paragraph(
            &export(&bytes, &options(RevisionView::Original)),
            "0A000005"
        )),
        ["Trackedbreak"]
    );
    let relocated = {
        let read = doc
            .read_paragraphs(&docx_edit::ReadParagraphsRequest {
                story: Some("body".to_owned()),
                para_ids: Some(vec!["0A000001".to_owned()]),
                view: docx_edit::EditTextView::Accepted,
            })
            .unwrap();
        assert_eq!(read.paragraphs[0].text, "BeforeAfter");
        doc.paragraph_mark_position("0A000001").unwrap().index + 1
    };
    doc.apply_raw_ops(
        "body",
        vec![RawOp::Delete {
            index: relocated,
            len: 1,
        }],
        &EditCtx::local("", ""),
    )
    .unwrap();
    let edited = doc
        .export_structured(&options(RevisionView::Accepted))
        .unwrap()
        .content;
    assert_eq!(
        labels(paragraph(&edited, "0A000001")),
        ["BeforeAfter"],
        "a break deleted in the session is gone"
    );
}

#[test]
fn unreconstructable_history_becomes_anchored_placeholders() {
    let bold_change = r#"<w:r><w:rPr><w:b/><w:rPrChange w:id="61" w:author="Ann"><w:rPr/></w:rPrChange></w:rPr><w:t>Now bold</w:t></w:r>"#;
    let colour_change = r#"<w:r><w:rPr><w:color w:val="FF0000"/><w:rPrChange w:id="62" w:author="Ann"><w:rPr/></w:rPrChange></w:rPr><w:t>Now red</w:t></w:r>"#;
    let deleted_cell = format!(
        r#"<w:tbl><w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid><w:tr><w:tc>{}</w:tc><w:tc><w:tcPr><w:cellDel w:id="63" w:author="Ann"/></w:tcPr>{}</w:tc></w:tr></w:tbl>"#,
        para("0B000003", &run("Kept")),
        para("0B000004", &run("Cut"))
    );
    let inserted_row = format!(
        r#"<w:tbl><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid><w:tr><w:trPr><w:ins w:id="64" w:author="Ann"/></w:trPr><w:tc>{}</w:tc></w:tr></w:tbl>"#,
        para("0B000005", &run("New row"))
    );
    let xml = [
        para("0B000001", bold_change),
        para("0B000002", colour_change),
        deleted_cell,
        para("0B000006", ""),
        inserted_row,
        para("0B000007", &run("End")),
    ]
    .concat();
    let bytes = Package::new(&xml).bytes();
    let placeholder = |content: &DocxStructuredContent, para_id: &str| {
        body(content).iter().any(|block| {
            matches!(&block.anchor, Anchor::Paragraph { para_id: id, .. } if id == para_id)
                && matches!(&block.content, BlockKind::Unsupported { element } if element == "w:p")
        })
    };
    let accepted = export(&bytes, &options(RevisionView::Accepted));
    assert!(matches!(
        inlines(paragraph(&accepted, "0B000001"))[0]
            .marks
            .as_deref(),
        Some([FormattingMark::Bold])
    ));
    let original = export(&bytes, &options(RevisionView::Original));
    assert!(placeholder(&original, "0B000001"));
    assert_eq!(labels(paragraph(&original, "0B000002")), ["Now red"]);
    let without_marks = export(
        &bytes,
        &ExportOptions {
            include_formatting: Some(false),
            ..options(RevisionView::Original)
        },
    );
    assert_eq!(labels(paragraph(&without_marks, "0B000001")), ["Now bold"]);
    let table_placeholders = |content: &DocxStructuredContent| {
        body(content)
            .iter()
            .filter(|block| matches!(&block.content, BlockKind::Unsupported { element } if element == "w:tbl"))
            .count()
    };
    assert_eq!(
        table_placeholders(&accepted),
        1,
        "a deleted cell cannot be removed"
    );
    assert_eq!(
        table_placeholders(&export(&bytes, &options(RevisionView::Markup))),
        2,
        "the markup view cannot attribute cell or row changes"
    );
    assert!(count(&original, DiagnosticCode::UnsupportedRevision) >= 2);
}

#[test]
fn html_tables_escape_text_and_mark_every_nested_block() {
    let nested = format!(
        r#"<w:tbl><w:tblGrid><w:gridCol w:w="900"/></w:tblGrid><w:tr><w:tc>{}</w:tc></w:tr></w:tbl>{}"#,
        para("0C000004", &run("Inner")),
        para("0C000005", "")
    );
    let link = r#"<w:hyperlink r:id="rIdScript"><w:r><w:t>click</w:t></w:r></w:hyperlink>"#;
    let holder = para("0C000003", &run("Holder"));
    let xml = format!(
        r#"<w:tbl><w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:gridSpan w:val="2"/></w:tcPr>{}</w:tc></w:tr><w:tr><w:tc>{}</w:tc><w:tc>{holder}{nested}</w:tc></w:tr></w:tbl>{}"#,
        para(
            "0C000001",
            &format!(
                "{}<w:r><w:rPr><w:b/></w:rPr><w:t>Strong</w:t></w:r>",
                run("&lt;script&gt;alert(1)&lt;/script&gt; &amp; **not bold** ")
            )
        ),
        para("0C000002", link),
        para("0C000006", &run("After"))
    );
    let bytes = Package::new(&xml)
        .rel(
            "rIdScript",
            "hyperlink\" TargetMode=\"External",
            "javascript:alert(1)",
        )
        .bytes();
    let markdown = export_docx_markdown(&bytes, &options(RevisionView::Accepted)).unwrap();
    let text = &markdown.markdown;
    assert!(
        text.contains(
            "&lt;script&gt;alert(1)&lt;/script&gt; &amp; **not bold** <strong>Strong</strong>"
        ),
        "{text}"
    );
    assert!(
        !text.contains("<script>") && !text.contains("\\*"),
        "{text}"
    );
    assert!(!text.contains("javascript:"), "{text}");
    assert_eq!(text.matches("<table>").count(), 2, "{text}");
    assert_eq!(
        text.matches("<!-- docx-export:").count(),
        markdown.anchors.len()
    );
    let content = export(&bytes, &options(RevisionView::Accepted));
    let mut nested_blocks = 0;
    fn visit(block: &Block, count: &mut usize) {
        *count += 1;
        match &block.content {
            BlockKind::Table { table } => {
                for cell in table.rows.iter().flat_map(|row| &row.cells) {
                    for block in &cell.blocks {
                        visit(block, count);
                    }
                }
            }
            BlockKind::ContentControl { blocks, .. } => {
                for block in blocks {
                    visit(block, count);
                }
            }
            _ => {}
        }
    }
    for block in body(&content) {
        visit(block, &mut nested_blocks);
    }
    assert_eq!(
        markdown.anchors.len(),
        nested_blocks,
        "every block has a marker"
    );
    let mut scripted = content.clone();
    let BlockKind::Table { table } = &mut scripted.stories[0].blocks[0].content else {
        panic!("a table first");
    };
    let BlockKind::Paragraph { paragraph } = &mut table.rows[1].cells[0].blocks[0].content else {
        panic!("a paragraph in the cell");
    };
    paragraph.inlines[0].link = Some(docx_edit::structured::Link {
        href: "java\tscript:alert(1)".to_owned(),
        title: Some("\"><img onerror=x>".to_owned()),
    });
    let copy = table.rows[1].cells[0].blocks[0].clone();
    scripted.stories[0].blocks.push(copy);
    let rendered = render_docx_markdown(&scripted, &MarkdownOptions::default()).unwrap();
    assert!(
        !rendered.markdown.contains("script:"),
        "{}",
        rendered.markdown
    );
    assert!(rendered.diagnostics.iter().any(|diagnostic| {
        diagnostic.code == DiagnosticCode::MarkdownLossy
            && diagnostic.message.contains("are not linked")
    }));
    paragraph_link_is_attribute_escaped();
}

fn paragraph_link_is_attribute_escaped() {
    let xml = format!(
        r#"<w:tbl><w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:gridSpan w:val="2"/></w:tcPr>{}</w:tc></w:tr></w:tbl>{}"#,
        para(
            "0C000007",
            r#"<w:hyperlink r:id="rIdQuote" w:tooltip="say &quot;hi&quot; &lt;b&gt;"><w:r><w:t>quoted</w:t></w:r></w:hyperlink>"#
        ),
        para("0C000008", "")
    );
    let bytes = Package::new(&xml)
        .rel(
            "rIdQuote",
            "hyperlink\" TargetMode=\"External",
            "https://example.com/?a=1&amp;b=%22x%22",
        )
        .bytes();
    let markdown = export_docx_markdown(&bytes, &options(RevisionView::Accepted)).unwrap();
    assert!(
        markdown
            .markdown
            .contains(r#"<a href="https://example.com/?a=1&amp;b=%22x%22" title="say &quot;hi&quot; &lt;b&gt;">quoted</a>"#),
        "{}",
        markdown.markdown
    );
}

#[test]
fn tooltips_and_encoded_script_targets_stay_inert() {
    let xml = [
        para(
            "0C100001",
            r#"<w:hyperlink r:id="rIdSafe" w:tooltip="tip&#10;&#10;&lt;img src=x onerror=alert(1)&gt;&#10;&#10;"><w:r><w:t>tip</w:t></w:r></w:hyperlink>"#,
        ),
        para(
            "0C100002",
            r#"<w:hyperlink r:id="rIdScript"><w:r><w:t>click</w:t></w:r></w:hyperlink>"#,
        ),
    ]
    .concat();
    let bytes = Package::new(&xml)
        .rel(
            "rIdSafe",
            "hyperlink\" TargetMode=\"External",
            "https://example.com",
        )
        .rel(
            "rIdScript",
            "hyperlink\" TargetMode=\"External",
            "javascript&amp;#58;alert(1)",
        )
        .bytes();
    let markdown = export_docx_markdown(&bytes, &options(RevisionView::Accepted))
        .unwrap()
        .markdown;
    assert!(
        markdown.contains(r#"[tip](https://example.com "tip  \<img src=x onerror=alert(1)\>  ")"#),
        "{markdown}"
    );
    assert!(markdown.contains("\nclick\n"), "{markdown}");
    assert!(!markdown.contains("javascript"), "{markdown}");
}

#[test]
fn shared_header_parts_are_exported_once() {
    let header = format!(
        r#"<w:hdr {}>{}</w:hdr>"#,
        fixture::namespaces(),
        para("0D000001", &run("Shared header"))
    );
    let xml = format!(
        r#"<w:p w14:paraId="0D000002"><w:pPr><w:sectPr><w:headerReference w:type="default" r:id="rIdA"/></w:sectPr></w:pPr>{}</w:p>{}<w:sectPr><w:headerReference w:type="default" r:id="rIdB"/></w:sectPr>"#,
        run("One"),
        para("0D000003", &run("Two"))
    );
    let bytes = Package::new(&xml)
        .part("header1.xml", "rIdA", "header", "header", &header)
        .rel("rIdB", "header", "header1.xml")
        .bytes();
    let content = export(
        &bytes,
        &with_stories(RevisionView::Accepted, &[StorySelection::Headers]),
    );
    let headers: Vec<(String, Vec<u32>)> = content
        .stories
        .iter()
        .filter(|story| story.kind == StoryKind::Header)
        .map(|story| {
            (
                story.story.clone(),
                story.uses.iter().map(|used| used.section_index).collect(),
            )
        })
        .collect();
    assert_eq!(headers, [("hf:rIdA".to_owned(), vec![0, 1])]);
    let doc = open(&bytes);
    let story = |doc: &EditingDoc| {
        doc.export_structured(&with_stories(
            RevisionView::Accepted,
            &[StorySelection::Headers],
        ))
        .unwrap()
        .content
    };
    assert_eq!(story(&doc).stories.len(), 1);
    doc.toggle_format(
        &EditCtx::local("", ""),
        docx_edit::StoryRange::new("hf:rIdB", 0, 6),
        docx_edit::SimpleFormat::Bold,
    )
    .unwrap();
    let diverged = story(&doc);
    assert_eq!(
        diverged
            .stories
            .iter()
            .map(|story| story.story.as_str())
            .collect::<Vec<_>>(),
        ["hf:rIdA", "hf:rIdB"],
        "a formatting edit to one alias keeps both"
    );
    assert!(diverged.diagnostics.iter().any(|diagnostic| {
        diagnostic.code == DiagnosticCode::AmbiguousIdentity
            && diagnostic.severity == docx_edit::structured::Severity::Warning
    }));
    let named = open(&bytes);
    for alias in ["hf:rIdA", "hf:rIdB"] {
        named
            .insert_text(
                &EditCtx::local("", ""),
                docx_edit::Position::new(alias, 0),
                alias,
                docx_edit::FormatPolicy::Inherit,
            )
            .unwrap();
    }
    assert_eq!(
        story(&named)
            .stories
            .iter()
            .map(|story| story.story.as_str())
            .collect::<Vec<_>>(),
        ["hf:rIdA", "hf:rIdB"],
        "aliases whose text names their own story still differ"
    );
}

#[test]
fn huge_paragraphs_stop_without_losing_the_prefix() {
    let huge = "x".repeat(2_000_000);
    let xml = [
        para("0E000001", &run("Small")),
        para("0E000002", &run(&huge)),
        para("0E000003", &run("After")),
    ]
    .concat();
    let bytes = Package::new(&xml).bytes();
    let within = |max_bytes: u32| {
        export(
            &bytes,
            &ExportOptions {
                max_bytes: Some(max_bytes),
                ..options(RevisionView::Accepted)
            },
        )
    };
    let content = within(1_048_576);
    assert!(content.truncated);
    assert_eq!(body(&content).len(), 1);
    assert_eq!(labels(&body(&content)[0]), ["Small"]);
    assert!(serde_json::to_vec(&content).unwrap().len() <= 1_048_576);
    let refused = within(4_096);
    assert!(refused.truncated);
    assert_eq!(
        body(&refused)
            .iter()
            .map(|block| labels(block).concat())
            .collect::<Vec<_>>(),
        ["Small"],
        "the leading block that fits is kept"
    );
    assert!(
        refused
            .diagnostics
            .last()
            .is_some_and(|diagnostic| diagnostic.message.contains("too large to project")),
        "the paragraph past the projection budget is never read"
    );
    let mut kept = 0;
    for max_bytes in [4_096, 65_536, 262_144, 524_288, 1_048_576, 4_194_304] {
        let content = within(max_bytes);
        let blocks: Vec<String> = body(&content)
            .iter()
            .map(|block| labels(block).concat())
            .collect();
        assert!(blocks.len() >= kept, "{max_bytes} bytes kept fewer blocks");
        assert!(blocks.first().is_none_or(|first| first == "Small"));
        kept = blocks.len();
    }
    assert_eq!(kept, 3);
}

#[test]
fn a_long_manual_keeps_its_title_within_a_small_budget() {
    let manual: String = (0..1_200)
        .map(|index| {
            para(
                &format!("{:08X}", 0x0E70_0000 + index),
                &run(&"m".repeat(1_000)),
            )
        })
        .collect();
    let xml = format!(
        r#"{}{manual}"#,
        para(
            "0E6F0001",
            &format!(
                r#"<w:commentRangeStart w:id="1"/>{}<w:commentRangeEnd w:id="1"/><w:r><w:commentReference w:id="1"/></w:r>"#,
                run("Manual")
            ),
        )
    );
    let comments = format!(
        r#"<w:comments {}><w:comment w:id="1" w:author="Ann">{}</w:comment></w:comments>"#,
        fixture::namespaces(),
        para("0E6F0002", &run("Check"))
    );
    let bytes = Package::new(&xml)
        .part("comments.xml", "rIdComments", COMMENTS, COMMENTS, &comments)
        .bytes();
    let options = ExportOptions {
        max_bytes: Some(4_096),
        ..options(RevisionView::Accepted)
    };
    for content in [
        export(&bytes, &options),
        open(&bytes).export_structured(&options).unwrap().content,
    ] {
        assert!(content.truncated);
        assert_eq!(labels(&body(&content)[0])[0], "Manual");
        assert!(serde_json::to_vec(&content).unwrap().len() <= 4_096);
    }
}

#[test]
fn a_leading_table_is_kept_before_a_paragraph_past_the_projection_budget() {
    let xml = format!(
        r#"<w:tbl><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid><w:tr><w:tc>{}</w:tc></w:tr></w:tbl>{}"#,
        para("0E800001", &run("Manual")),
        para("0E800002", &run(&"d".repeat(2_000_000)))
    );
    let bytes = Package::new(&xml).bytes();
    let options = ExportOptions {
        max_bytes: Some(4_096),
        ..options(RevisionView::Accepted)
    };
    for content in [
        export(&bytes, &options),
        open(&bytes).export_structured(&options).unwrap().content,
    ] {
        assert!(content.truncated);
        assert_eq!(body(&content).len(), 1);
        let table = tables(&content)[0];
        assert_eq!(labels(&table.rows[0].cells[0].blocks[0]), ["Manual"]);
        assert!(
            content
                .diagnostics
                .last()
                .is_some_and(|diagnostic| diagnostic.message.contains("too large to project"))
        );
    }
}

#[test]
fn simple_field_results_with_tabs_and_breaks_open_and_export() {
    let xml = [
        r#"<w:p w14:paraId="0E900001"><w:fldSimple w:instr=" REF Summary "><w:r><w:t>First</w:t><w:tab/><w:br w:type="page"/><w:t>Second</w:t></w:r></w:fldSimple></w:p>"#.to_owned(),
        r#"<w:p w14:paraId="0E900002"><w:r><w:t>Lead</w:t></w:r><w:fldSimple w:instr=" REF Summary "><w:r><w:t>A</w:t><w:tab/><w:br w:type="page"/><w:t>B</w:t></w:r></w:fldSimple><w:r><w:br w:type="page"/><w:t>Tail</w:t></w:r></w:p>"#.to_owned(),
        para("0E900003", &run("After")),
    ]
    .concat();
    let bytes = Package::new(&xml).bytes();
    for content in [
        export(&bytes, &options(RevisionView::Accepted)),
        open(&bytes)
            .export_structured(&options(RevisionView::Accepted))
            .unwrap()
            .content,
    ] {
        assert_eq!(field_results(paragraph(&content, "0E900001")).len(), 1);
        let second = paragraph(&content, "0E900002");
        assert_eq!(field_results(second).len(), 1);
        let texts = labels(second);
        assert_eq!(texts.first().map(String::as_str), Some("Lead"));
        assert_eq!(texts.last().map(String::as_str), Some("Tail"));
        assert_eq!(labels(paragraph(&content, "0E900003")), ["After"]);
    }
}

#[test]
fn an_oversized_field_payload_stops_after_the_prefix() {
    let huge = "z".repeat(2_000_000);
    let xml = [
        para("0E200001", &run("Small")),
        para(
            "0E200002",
            &complex_field(" REF _Ref1 \\h ", &format!("{}{}", run(&huge), run("tail"))),
        ),
        para("0E200003", &run("After")),
    ]
    .concat();
    let content = export(
        &Package::new(&xml).bytes(),
        &ExportOptions {
            max_bytes: Some(65_536),
            ..options(RevisionView::Accepted)
        },
    );
    assert!(content.truncated);
    assert_eq!(
        body(&content)
            .iter()
            .map(|block| labels(block).concat())
            .collect::<Vec<_>>(),
        ["Small"]
    );
    assert!(
        content
            .diagnostics
            .last()
            .is_some_and(|diagnostic| diagnostic.message.contains("too large to project"))
    );
}

#[test]
fn an_oversized_comment_body_is_refused_before_it_is_read() {
    let huge = "c".repeat(2_000_000);
    let xml = para(
        "0E400001",
        &format!(
            r#"<w:commentRangeStart w:id="1"/>{}<w:commentRangeEnd w:id="1"/><w:r><w:commentReference w:id="1"/></w:r>"#,
            run("Annotated")
        ),
    );
    let comments = format!(
        r#"<w:comments {}><w:comment w:id="1" w:author="Ann">{}</w:comment></w:comments>"#,
        fixture::namespaces(),
        para("0E400002", &run(&huge))
    );
    let bytes = Package::new(&xml)
        .part("comments.xml", "rIdComments", COMMENTS, COMMENTS, &comments)
        .bytes();
    let options = ExportOptions {
        max_bytes: Some(65_536),
        ..with_stories(
            RevisionView::Accepted,
            &[StorySelection::Body, StorySelection::Comments],
        )
    };
    for content in [
        export(&bytes, &options),
        open(&bytes).export_structured(&options).unwrap().content,
    ] {
        assert!(content.truncated);
        assert_eq!(labels(&body(&content)[0])[0], "Annotated");
        assert!(
            content
                .stories
                .iter()
                .filter(|story| story.kind == StoryKind::Comment)
                .all(|story| story.blocks.is_empty())
        );
        assert!(
            content
                .diagnostics
                .last()
                .is_some_and(|diagnostic| diagnostic.message.contains("too large to project"))
        );
    }
}

#[test]
fn field_result_wrappers_do_not_count_as_blocks() {
    let xml = para(
        "0E500001",
        &complex_field(
            " REF _Ref1 \\h ",
            r#"<w:r><w:rPr><w:b/></w:rPr><w:t>Bold</w:t></w:r><w:r><w:t xml:space="preserve"> plain</w:t></w:r>"#,
        ),
    );
    let bytes = Package::new(&xml).bytes();
    let options = ExportOptions {
        max_blocks: Some(1),
        ..options(RevisionView::Accepted)
    };
    for content in [
        export(&bytes, &options),
        open(&bytes).export_structured(&options).unwrap().content,
    ] {
        assert!(!content.truncated);
        assert_eq!(body(&content).len(), 1);
        assert!(matches!(
            field_results(&body(&content)[0]).as_slice(),
            [(_, CachedResult::Inline { inlines })] if inlines.len() == 2
        ));
    }
}

#[test]
fn field_results_inherit_the_paragraph_style() {
    let styles = format!(
        r#"<w:styles {}><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style><w:style w:type="paragraph" w:styleId="Loud"><w:name w:val="Loud"/><w:rPr><w:b/></w:rPr></w:style></w:styles>"#,
        fixture::namespaces()
    );
    let xml = format!(
        r#"<w:p w14:paraId="0E600001"><w:pPr><w:pStyle w:val="Loud"/></w:pPr>{}</w:p>"#,
        complex_field(
            " REF _Ref1 \\h ",
            r#"<w:r><w:t>A</w:t></w:r><w:r><w:rPr><w:i/></w:rPr><w:t>B</w:t></w:r>"#,
        )
    );
    let bytes = Package::new(&xml).styles(&styles).bytes();
    for content in [
        export(&bytes, &options(RevisionView::Accepted)),
        open(&bytes)
            .export_structured(&options(RevisionView::Accepted))
            .unwrap()
            .content,
    ] {
        let results = field_results(&body(&content)[0]);
        let [(_, CachedResult::Inline { inlines })] = results.as_slice() else {
            panic!("an inline cached result");
        };
        let marks: Vec<(String, Option<Vec<FormattingMark>>)> = inlines
            .iter()
            .map(|inline| (label(inline), inline.marks.clone()))
            .collect();
        assert_eq!(
            marks,
            [
                ("A".to_owned(), Some(vec![FormattingMark::Bold])),
                (
                    "B".to_owned(),
                    Some(vec![FormattingMark::Bold, FormattingMark::Italic])
                ),
            ]
        );
    }
}

#[test]
fn a_completed_field_owner_is_kept_before_a_block_that_does_not_fit() {
    let xml = [
        para(
            "0E100001",
            r#"<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> 7 </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>first</w:t></w:r>"#,
        ),
        para("0E100002", &run("second")),
        para("0E100003", r#"<w:r><w:fldChar w:fldCharType="end"/></w:r>"#),
        para("0E100004", &run(&"y".repeat(100_000))),
    ]
    .concat();
    let content = export(
        &Package::new(&xml).bytes(),
        &ExportOptions {
            max_bytes: Some(16_384),
            ..options(RevisionView::Accepted)
        },
    );
    assert!(content.truncated);
    assert_eq!(body(&content).len(), 1);
    assert!(matches!(
        field_results(&body(&content)[0]).as_slice(),
        [(_, CachedResult::Blocks { blocks })] if blocks.len() == 3
    ));
}

#[test]
fn diagnostics_stay_capped() {
    let xml: String = (0..1_200)
        .map(|index| {
            para(
                &format!("1{index:07X}"),
                r#"<bofx:mark bofx:value="kept"/>"#,
            )
        })
        .collect();
    let content = export(
        &Package::new(&xml).bytes(),
        &options(RevisionView::Accepted),
    );
    assert!(content.diagnostics.len() <= 1_000);
    assert_eq!(count(&content, DiagnosticCode::UnsupportedContent), 100);
    assert!(content.diagnostics.iter().any(|diagnostic| {
        diagnostic.code == DiagnosticCode::UnsupportedContent && diagnostic.anchor.is_none()
    }));
    assert_eq!(body(&content).len(), 1_200);
}

#[test]
fn deep_nesting_stops_at_the_depth_limit() {
    let mut nested = para("01000000", &run("Deepest"));
    for level in 0..40 {
        nested = format!(
            r#"<w:tbl><w:tblGrid><w:gridCol w:w="900"/></w:tblGrid><w:tr><w:tc>{nested}{}</w:tc></w:tr></w:tbl>"#,
            para(&format!("02{level:06}"), "")
        );
    }
    let xml = format!("{nested}{}", para("03000000", &run("After")));
    let bytes = Package::new(&xml).bytes();
    let content = std::thread::Builder::new()
        .stack_size(256 << 20)
        .spawn(move || export(&bytes, &options(RevisionView::Accepted)))
        .unwrap()
        .join()
        .unwrap();
    assert!(content.diagnostics.iter().any(|diagnostic| {
        diagnostic.code == DiagnosticCode::UnsupportedContent
            && diagnostic.message.contains("nested more than")
    }));
    assert_eq!(labels(paragraph(&content, "03000000")), ["After"]);
}

#[test]
fn session_headings_match_the_export() {
    let bytes = fixture::principal_docx();
    let doc = open(&bytes);
    let content = doc
        .export_structured(&options(RevisionView::Accepted))
        .unwrap()
        .content;
    let exported: Vec<(String, docx_edit::structured::HeadingInfo)> = body(&content)
        .iter()
        .filter_map(|block| {
            let heading = match &block.content {
                BlockKind::Heading { heading, .. } => heading.clone(),
                BlockKind::ListItem {
                    heading: Some(heading),
                    ..
                } => heading.clone(),
                _ => return None,
            };
            let Anchor::Paragraph { para_id, .. } = &block.anchor else {
                return None;
            };
            Some((para_id.clone(), heading))
        })
        .collect();
    assert!(exported.len() >= 4);
    assert_eq!(doc.paragraph_headings("body").unwrap(), exported);
}

#[test]
fn projected_field_results_stay_inside_their_field_in_result_order() {
    let result = concat!(
        r#"<w:r><w:t xml:space="preserve">A </w:t></w:r>"#,
        r#"<w:hyperlink w:anchor="_Toc1"><w:r><w:t>B</w:t></w:r></w:hyperlink>"#,
        r#"<w:fldSimple w:instr=" PAGE "><w:r><w:t>7</w:t></w:r></w:fldSimple>"#,
        r#"<w:r><w:t xml:space="preserve"> C</w:t></w:r>"#
    );
    let xml = [
        para("0F100001", &complex_field(" TOC \\o ", result)),
        para("0F100002", &run("After")),
    ]
    .concat();
    let bytes = Package::new(&xml).bytes();
    for content in [
        export(&bytes, &options(RevisionView::Accepted)),
        open(&bytes)
            .export_structured(&options(RevisionView::Accepted))
            .unwrap()
            .content,
    ] {
        let owner = paragraph(&content, "0F100001");
        assert_eq!(
            labels(owner),
            ["<field>"],
            "no result content escapes the field"
        );
        let anchor = field_anchor(owner);
        let results = field_results(owner);
        let [(_, CachedResult::Inline { inlines })] = results.as_slice() else {
            panic!("an inline cached result: {results:?}");
        };
        let order: Vec<(String, bool)> = inlines
            .iter()
            .map(|inline| (label(inline), inline.link.is_some()))
            .collect();
        assert_eq!(
            order,
            [
                ("A ".to_owned(), false),
                ("B".to_owned(), true),
                ("<field>".to_owned(), false),
                (" C".to_owned(), false),
            ]
        );
        assert!(inlines.iter().all(|inline| inline.anchor == anchor));
    }
}

#[test]
fn breaks_inside_inline_controls_stay_at_their_offsets() {
    let control = format!(
        r#"<w:sdt><w:sdtPr><w:tag w:val="clause"/><w:id w:val="7"/></w:sdtPr><w:sdtContent>{}<w:r><w:br w:type="page"/></w:r>{}</w:sdtContent></w:sdt>"#,
        run("Alpha"),
        run("Beta")
    );
    let xml = [
        para("0F200001", &format!("{}{control}", run("Before "))),
        para("0F200002", ""),
    ]
    .concat();
    let bytes = Package::new(&xml).bytes();
    for content in [
        export(&bytes, &options(RevisionView::Accepted)),
        open(&bytes)
            .export_structured(&options(RevisionView::Accepted))
            .unwrap()
            .content,
    ] {
        let children: Vec<String> = inlines(paragraph(&content, "0F200001"))
            .iter()
            .find_map(|inline| match &inline.content {
                InlineKind::ContentControl { inlines, .. } => {
                    Some(inlines.iter().map(label).collect())
                }
                _ => None,
            })
            .unwrap();
        assert_eq!(children, ["Alpha", "<Page>", "Beta"]);
    }
}

fn control_tree(inlines: &[Inline]) -> Vec<String> {
    inlines
        .iter()
        .map(|inline| match &inline.content {
            InlineKind::ContentControl { inlines, .. } => {
                format!("[{}]", control_tree(inlines).join(" "))
            }
            _ => label(inline),
        })
        .collect()
}

#[test]
fn breaks_inside_nested_inline_controls_stay_in_the_inner_control() {
    let control = |id: &str, content: &str| {
        format!(
            r#"<w:sdt><w:sdtPr><w:id w:val="{id}"/></w:sdtPr><w:sdtContent>{content}</w:sdtContent></w:sdt>"#
        )
    };
    let inner = control("9", &format!("{}{}{}", run("A"), page_break(), run("B")));
    let outer = control("8", &format!("{}{inner}{}", run("X"), run("Y")));
    let xml = [para("0F210001", &outer), para("0F210002", "")].concat();
    let bytes = Package::new(&xml).bytes();
    for content in [
        export(&bytes, &options(RevisionView::Accepted)),
        open(&bytes)
            .export_structured(&options(RevisionView::Accepted))
            .unwrap()
            .content,
    ] {
        assert_eq!(
            control_tree(inlines(paragraph(&content, "0F210001"))),
            ["[X [A <Page> B] Y]"]
        );
    }
}

#[test]
fn original_view_style_changes_that_alter_marks_become_placeholders() {
    let styles = format!(
        r#"<w:styles {}><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style><w:style w:type="paragraph" w:styleId="Loud"><w:name w:val="Loud"/><w:rPr><w:b/></w:rPr></w:style><w:style w:type="paragraph" w:styleId="Spaced"><w:name w:val="Spaced"/><w:pPr><w:spacing w:after="400"/></w:pPr></w:style></w:styles>"#,
        fixture::namespaces()
    );
    let changed = |id: &str, style: &str, previous: &str, text: &str| {
        format!(
            r#"<w:p w14:paraId="{id}"><w:pPr><w:pStyle w:val="{style}"/><w:pPrChange w:id="81" w:author="Ann"><w:pPr>{previous}</w:pPr></w:pPrChange></w:pPr>{}</w:p>"#,
            run(text)
        )
    };
    let normal = r#"<w:pStyle w:val="Normal"/>"#;
    let xml = [
        changed("0F300001", "Loud", normal, "Now loud"),
        changed("0F300002", "Spaced", normal, "Now spaced"),
        changed("0F300003", "Spaced", "", "No earlier properties"),
    ]
    .concat();
    let bytes = Package::new(&xml).styles(&styles).bytes();
    let original = export(&bytes, &options(RevisionView::Original));
    assert!(matches!(
        &body(&original)[0].content,
        BlockKind::Unsupported { element } if element == "w:p"
    ));
    let spaced = &body(&original)[1];
    assert_eq!(labels(spaced), ["Now spaced"]);
    let BlockKind::Paragraph { paragraph } = &spaced.content else {
        panic!("a paragraph");
    };
    assert_eq!(
        paragraph.style_id.as_deref(),
        Some("Normal"),
        "the original view reads the earlier style"
    );
    assert!(matches!(
        &body(&original)[2].content,
        BlockKind::Unsupported { element } if element == "w:p"
    ));
    let accepted = export(&bytes, &options(RevisionView::Accepted));
    assert!(matches!(
        inlines(&body(&accepted)[0])[0].marks.as_deref(),
        Some([FormattingMark::Bold])
    ));
}

#[test]
fn numbers_a_format_cannot_write_are_diagnosed() {
    let abstracts = r#"<w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="2000000000"/><w:numFmt w:val="upperRoman"/><w:lvlText w:val="%1."/></w:lvl></w:abstractNum>"#;
    let xml = numbered("0F400001", 1, 0, "Huge");
    let bytes = Package::new(&xml)
        .numbering(&numbering_part(
            abstracts,
            r#"<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>"#,
        ))
        .bytes();
    let content = export(&bytes, &options(RevisionView::Accepted));
    assert_eq!(markers(&content), [None]);
    assert!(content.diagnostics.iter().any(|diagnostic| {
        diagnostic.code == DiagnosticCode::UnsupportedNumbering
            && diagnostic.message.contains("2000000000")
    }));
}

#[test]
fn shared_ids_leave_cell_and_comment_anchors_without_a_location() {
    let xml = format!(
        r#"<w:tbl><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid><w:tr><w:tc>{}{}</w:tc></w:tr></w:tbl>{}"#,
        para("0F500001", &run("One")),
        para("0F500002", &run("Two")),
        para(
            "0F500003",
            &format!(
                r#"<w:commentRangeStart w:id="5"/>{}<w:commentRangeEnd w:id="5"/><w:r><w:commentReference w:id="5"/></w:r>"#,
                run("Noted")
            )
        )
    );
    let comments = format!(
        r#"<w:comments {}><w:comment w:id="5" w:author="Ann">{}</w:comment></w:comments>"#,
        fixture::namespaces(),
        para("0F500004", &run("Remark"))
    );
    let bytes = Package::new(&format!("{xml}{}", para("0F500005", &run("Tail"))))
        .part("comments.xml", "rIdComments", COMMENTS, COMMENTS, &comments)
        .bytes();
    let doc = open(&bytes);
    let rename = |story: &str, para_id: &str, to: &str| {
        let position = doc.paragraph_mark_position(para_id).unwrap();
        assert_eq!(position.story, story);
        doc.apply_raw_ops(
            story,
            vec![RawOp::SetEmbedAttr {
                index: position.index,
                key: "paraId".to_owned(),
                value: yrs::Any::from(to),
            }],
            &EditCtx::local("", ""),
        )
        .unwrap();
    };
    rename("body:t0:r0c0", "0F500002", "0F500001");
    rename("body", "0F500005", "0F500003");
    let content = doc
        .export_structured(&with_stories(
            RevisionView::Accepted,
            &[StorySelection::Body, StorySelection::Comments],
        ))
        .unwrap()
        .content;
    let table = tables(&content)[0];
    assert!(matches!(
        table.rows[0].cells[0].anchor,
        Anchor::Table { .. }
    ));
    let comment = comment_metadata(&content);
    assert!(matches!(
        comment.anchors.as_slice(),
        [Anchor::Paragraph { para_id, .. }] if para_id.is_empty()
    ));
    assert!(count(&content, DiagnosticCode::AmbiguousIdentity) >= 3);
}
