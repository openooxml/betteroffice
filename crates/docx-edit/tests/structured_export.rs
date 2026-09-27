#[allow(dead_code)]
#[path = "support/structured_fixture.rs"]
mod fixture;

use std::cell::Cell;
use std::rc::Rc;

use docx_edit::structured::{
    Anchor, AnchorScope, Block, BlockKind, CachedResult, DiagnosticCode, DocxStructuredContent,
    ExportFailureCode, ExportOptions, FormattingMark, Inline, InlineKind, MarkdownOptions,
    OutlineSource, RevisionKind, RevisionView, StoryKind, StorySelection, UnlocatedReason,
    VerticalMerge, export_docx_markdown, export_docx_structured, render_docx_markdown,
};
use docx_edit::{
    EditHistory, EditOperation, EditRequest, EditSource, EditStep, EditTextView, EditingDoc,
    ParagraphTarget, RawOp, ReadParagraphsRequest, TargetEdge, TextTarget, UndoSession,
    seed_from_docx,
};
use yrs::Any;

const ALL: [StorySelection; 6] = [
    StorySelection::Body,
    StorySelection::Headers,
    StorySelection::Footers,
    StorySelection::Footnotes,
    StorySelection::Endnotes,
    StorySelection::Comments,
];

fn options(view: RevisionView) -> ExportOptions {
    ExportOptions::new(view)
}

fn all(view: RevisionView) -> ExportOptions {
    ExportOptions {
        stories: Some(ALL.to_vec()),
        ..options(view)
    }
}

fn export(bytes: &[u8], options: &ExportOptions) -> DocxStructuredContent {
    export_docx_structured(bytes, options).unwrap()
}

fn open(bytes: &[u8]) -> EditingDoc {
    let doc = EditingDoc::new(4242);
    seed_from_docx(&doc, bytes).unwrap();
    doc
}

fn body(content: &DocxStructuredContent) -> &[Block] {
    &content
        .stories
        .iter()
        .find(|story| story.story == "body")
        .unwrap()
        .blocks
}

fn inlines(block: &Block) -> &[Inline] {
    match &block.content {
        BlockKind::Paragraph { paragraph }
        | BlockKind::Heading { paragraph, .. }
        | BlockKind::ListItem { paragraph, .. } => &paragraph.inlines,
        other => panic!("not a paragraph: {other:?}"),
    }
}

fn text(block: &Block) -> String {
    inlines(block)
        .iter()
        .map(|inline| match &inline.content {
            InlineKind::Text { text } => text.clone(),
            InlineKind::Tab => "\t".to_owned(),
            _ => "\u{FFFC}".to_owned(),
        })
        .collect()
}

fn block<'a>(content: &'a DocxStructuredContent, para_id: &str) -> &'a Block {
    body(content)
        .iter()
        .find(|block| {
            matches!(&block.anchor, Anchor::Paragraph { para_id: id, .. } if id == para_id)
                && !matches!(block.content, BlockKind::SectionBreak { .. })
        })
        .unwrap_or_else(|| panic!("no block for {para_id}"))
}

fn codes(content: &DocxStructuredContent) -> Vec<DiagnosticCode> {
    content
        .diagnostics
        .iter()
        .map(|diagnostic| diagnostic.code)
        .collect()
}

#[test]
fn principal_fixture_matches_its_builder() {
    let committed = ooxml_opc::unzip_parts(&fixture::principal_docx()).unwrap();
    assert_eq!(committed, fixture::principal_parts());
}

#[test]
fn golden_exports_cover_every_view() {
    let bytes = fixture::principal_docx();
    for (view, name) in [
        (RevisionView::Accepted, "accepted"),
        (RevisionView::Original, "original"),
        (RevisionView::Markup, "markup"),
    ] {
        let content = export(&bytes, &all(view));
        fixture::golden(
            &format!("principal.{name}.json"),
            &(serde_json::to_string_pretty(&content).unwrap() + "\n"),
        );
        let markdown = render_docx_markdown(&content, &MarkdownOptions::default()).unwrap();
        fixture::golden(&format!("principal.{name}.md"), &markdown.markdown);
    }
}

#[test]
fn headings_resolve_through_direct_style_and_builtin_sources() {
    let content = export(&fixture::principal_docx(), &options(RevisionView::Accepted));
    let heading = |para_id: &str| match &block(&content, para_id).content {
        BlockKind::Heading { heading, .. } => Some(heading.clone()),
        BlockKind::ListItem { heading, .. } => heading.clone(),
        _ => None,
    };
    let styled = |style: &str| OutlineSource::Style {
        style_id: style.to_owned(),
    };
    assert_eq!(heading("00000001").unwrap().source, styled("Heading1"));
    let direct = heading("00000002").unwrap();
    assert_eq!(
        (direct.outline_level, direct.source),
        (1, OutlineSource::Direct)
    );
    assert_eq!(heading("00000003").unwrap().source, styled("Title"));
    assert_eq!(
        heading("00000004").unwrap().source,
        OutlineSource::BuiltinStyleId {
            style_id: "Heading2".to_owned()
        }
    );
    assert_eq!(heading("00000005"), None, "outline level 9 is body text");
    let numbered = &block(&content, "0000000B").content;
    assert!(
        matches!(numbered, BlockKind::ListItem { heading: Some(_), list, .. } if list.marker.as_deref() == Some("I."))
    );
}

#[test]
fn document_default_outline_levels_are_reported_as_such() {
    let mut parts = fixture::principal_parts();
    let styles = parts
        .iter_mut()
        .find(|(name, _)| name == "word/styles.xml")
        .unwrap();
    let xml = String::from_utf8(styles.1.clone()).unwrap().replace(
        "<w:style w:type=\"paragraph\" w:default=\"1\"",
        "<w:docDefaults><w:pPrDefault><w:pPr><w:outlineLvl w:val=\"4\"/></w:pPr></w:pPrDefault></w:docDefaults><w:style w:type=\"paragraph\" w:default=\"1\"",
    );
    styles.1 = xml.into_bytes();
    let content = export(
        &ooxml_opc::rezip_parts(&parts).unwrap(),
        &options(RevisionView::Accepted),
    );
    let BlockKind::Heading { heading, .. } = &block(&content, "00000017").content else {
        panic!("the document default makes every plain paragraph a heading");
    };
    assert_eq!(heading.outline_level, 4);
    assert_eq!(heading.source, OutlineSource::DocumentDefault);
}

#[test]
fn list_markers_match_the_rendered_markers() {
    let bytes = fixture::principal_docx();
    let content = export(&bytes, &options(RevisionView::Accepted));
    let exported: Vec<Option<String>> = body(&content)
        .iter()
        .filter_map(|block| match &block.content {
            BlockKind::ListItem { list, .. } => Some(list.marker.clone()),
            _ => None,
        })
        .collect();
    assert_eq!(
        exported,
        [
            Some("1."),
            Some("a)"),
            Some("\u{2022}"),
            Some("2."),
            Some("7."),
            Some("I."),
            None
        ]
        .map(|marker| marker.map(str::to_owned))
    );
    let doc = open(&bytes);
    let rendered: Vec<String> =
        docx_edit::bridge::yrs_doc_to_layout_blocks(&doc, "body", &Default::default())
            .unwrap()
            .into_iter()
            .filter_map(|block| match block {
                docx_layout::types::LayoutBlock::Paragraph(paragraph) => {
                    paragraph.attrs?.list_marker
                }
                _ => None,
            })
            .collect();
    let resolved: Vec<String> = exported.into_iter().flatten().collect();
    assert_eq!(&rendered[..resolved.len()], resolved.as_slice());
    let unrendered = content
        .diagnostics
        .iter()
        .find(|diagnostic| diagnostic.code == DiagnosticCode::UnsupportedNumbering)
        .unwrap();
    assert!(unrendered.message.contains("cardinalText"));
}

#[test]
fn tables_report_grid_positions_spans_and_merge_continuations() {
    let content = export(&fixture::principal_docx(), &options(RevisionView::Accepted));
    let table = body(&content)
        .iter()
        .find_map(|block| match &block.content {
            BlockKind::Table { table } => Some(table),
            _ => None,
        })
        .unwrap();
    assert_eq!(table.grid_columns, 3);
    assert!(table.rows[0].header);
    let positions: Vec<Vec<(u32, u32, u32)>> = table
        .rows
        .iter()
        .map(|row| {
            row.cells
                .iter()
                .map(|cell| (cell.column, cell.grid_span, cell.row_span))
                .collect()
        })
        .collect();
    assert_eq!(
        positions,
        [
            vec![(0, 2, 1), (2, 1, 1)],
            vec![(0, 1, 2), (1, 1, 1), (2, 1, 1)],
            vec![(0, 1, 0), (1, 1, 1), (2, 1, 1)],
            vec![(1, 2, 1)],
        ]
    );
    let owner = &table.rows[1].cells[0];
    assert_eq!(owner.vertical_merge, VerticalMerge::Restart);
    let continuation = &table.rows[2].cells[0];
    assert_eq!(continuation.vertical_merge, VerticalMerge::Continue);
    assert_eq!(
        continuation
            .merge_origin
            .map(|origin| (origin.row, origin.column)),
        Some((1, 0))
    );
    assert!(continuation.blocks.is_empty() && continuation.story.is_none());
    assert_eq!(table.rows[3].grid_before, 1);
    let nested = &table.rows[1].cells[2].blocks;
    assert!(
        nested
            .iter()
            .any(|block| matches!(block.content, BlockKind::Table { .. }))
    );
    assert!(codes(&content).contains(&DiagnosticCode::MergeContinuationContentOmitted));
}

#[test]
fn controls_keep_distinct_identities_for_duplicate_tags() {
    let content = export(&fixture::principal_docx(), &options(RevisionView::Accepted));
    let BlockKind::ContentControl {
        control,
        story,
        blocks,
    } = &body(&content)
        .iter()
        .find(|block| matches!(block.content, BlockKind::ContentControl { .. }))
        .unwrap()
        .content
    else {
        unreachable!()
    };
    assert_eq!(control.tag.as_deref(), Some("clause"));
    assert_eq!(control.ooxml_id.as_deref(), Some("-501"));
    assert!(control.showing_placeholder);
    assert_eq!(story.as_deref(), Some(control.control_id.as_str()));
    assert_eq!(text(&blocks[0]), "Inside a block control");
    let controls: Vec<_> = inlines(block(&content, "0000000E"))
        .iter()
        .filter_map(|inline| match &inline.content {
            InlineKind::ContentControl { control, inlines } => {
                assert!(inlines.iter().all(|child| matches!(
                    &child.anchor,
                    Anchor::Control { control_id, .. } if control_id == &control.control_id
                )));
                Some(control.clone())
            }
            _ => None,
        })
        .collect();
    assert_eq!(controls.len(), 2);
    assert_eq!(controls[0].tag, controls[1].tag);
    assert_ne!(controls[0].control_id, controls[1].control_id);
    assert_eq!(
        (
            controls[0].ooxml_id.as_deref(),
            controls[1].ooxml_id.as_deref()
        ),
        (Some("11"), Some("12"))
    );
}

#[test]
fn revision_views_project_insertions_and_deletions() {
    let bytes = fixture::principal_docx();
    let accepted = export(&bytes, &options(RevisionView::Accepted));
    let original = export(&bytes, &options(RevisionView::Original));
    let markup = export(&bytes, &options(RevisionView::Markup));
    assert_eq!(text(block(&accepted, "0000000F")), "Keep added end");
    assert_eq!(text(block(&original, "0000000F")), "Keep removed end");
    let attributed: Vec<(String, Vec<RevisionKind>, EditTextView)> =
        inlines(block(&markup, "0000000F"))
            .iter()
            .map(|inline| {
                let InlineKind::Text { text } = &inline.content else {
                    panic!("text only")
                };
                let Anchor::Range(range) = &inline.anchor else {
                    panic!("range anchors")
                };
                (
                    text.clone(),
                    inline
                        .revisions
                        .iter()
                        .map(|revision| revision.kind)
                        .collect(),
                    range.view,
                )
            })
            .collect();
    assert_eq!(
        attributed,
        [
            ("Keep ".to_owned(), vec![], EditTextView::Accepted),
            (
                "added ".to_owned(),
                vec![RevisionKind::Insertion],
                EditTextView::Accepted
            ),
            (
                "removed ".to_owned(),
                vec![RevisionKind::Deletion],
                EditTextView::Original
            ),
            ("end".to_owned(), vec![], EditTextView::Accepted),
        ]
    );
    let insertion = &inlines(block(&markup, "0000000F"))[1].revisions[0];
    assert_eq!(
        (insertion.id.as_deref(), insertion.author.as_deref()),
        (Some("21"), Some("Ann"))
    );
    assert!(codes(&accepted).contains(&DiagnosticCode::RevisionContentExcluded));
    assert!(codes(&original).contains(&DiagnosticCode::RevisionContentExcluded));
    assert!(!codes(&markup).contains(&DiagnosticCode::RevisionContentExcluded));
}

/// Every range anchor, sliced out of the batch read of its view, is the exported text.
fn assert_ranges_agree(doc: &EditingDoc, content: &DocxStructuredContent) {
    let mut checked = 0;
    for story in &content.stories {
        if story.kind == StoryKind::Comment {
            continue;
        }
        for block in &story.blocks {
            visit(doc, block, &mut checked);
        }
    }
    assert!(checked > 20, "only {checked} ranges were checked");

    fn visit(doc: &EditingDoc, block: &Block, checked: &mut usize) {
        match &block.content {
            BlockKind::Paragraph { .. }
            | BlockKind::Heading { .. }
            | BlockKind::ListItem { .. } => {
                for inline in inlines(block) {
                    let Anchor::Range(range) = &inline.anchor else {
                        continue;
                    };
                    let read = doc
                        .read_paragraphs(&ReadParagraphsRequest {
                            story: Some(range.story.clone()),
                            para_ids: Some(vec![range.start.para_id.clone()]),
                            view: range.view,
                        })
                        .unwrap();
                    let paragraph = &read.paragraphs[0];
                    let units: Vec<u16> = paragraph.text.encode_utf16().collect();
                    let slice = String::from_utf16(
                        &units[range.start.offset as usize..range.end.offset as usize],
                    )
                    .unwrap();
                    match &inline.content {
                        InlineKind::Text { text } => assert_eq!(&slice, text),
                        InlineKind::Tab => assert_eq!(slice, "\t"),
                        _ => {
                            assert_eq!(slice, "\u{FFFC}");
                            assert!(
                                paragraph
                                    .atoms
                                    .iter()
                                    .any(|atom| atom.offset == range.start.offset)
                            );
                        }
                    }
                    *checked += 1;
                }
            }
            BlockKind::Table { table } => {
                for cell in table.rows.iter().flat_map(|row| &row.cells) {
                    for block in &cell.blocks {
                        visit(doc, block, checked);
                    }
                }
            }
            BlockKind::ContentControl { blocks, .. } => {
                for block in blocks {
                    visit(doc, block, checked);
                }
            }
            _ => {}
        }
    }
}

#[test]
fn range_anchors_agree_with_batch_reads_in_every_view() {
    let doc = open(&fixture::principal_docx());
    for view in [
        RevisionView::Accepted,
        RevisionView::Original,
        RevisionView::Markup,
    ] {
        let content = doc.export_structured(&all(view)).unwrap().content;
        assert_ranges_agree(&doc, &content);
    }
}

#[test]
fn atoms_tabs_and_surrogates_follow_the_offset_contract() {
    let content = export(&fixture::principal_docx(), &options(RevisionView::Accepted));
    let paragraph = inlines(block(&content, "00000014"));
    let summary: Vec<(String, u32, u32)> = paragraph
        .iter()
        .map(|inline| {
            let Anchor::Range(range) = &inline.anchor else {
                panic!("range anchors")
            };
            let label = match &inline.content {
                InlineKind::Text { text } => text.clone(),
                InlineKind::Tab => "<tab>".to_owned(),
                InlineKind::Break { .. } => "<break>".to_owned(),
                other => panic!("unexpected {other:?}"),
            };
            (label, range.start.offset, range.end.offset)
        })
        .collect();
    assert_eq!(
        summary,
        [
            ("Atoms:".to_owned(), 0, 6),
            ("<tab>".to_owned(), 6, 7),
            ("tabbed".to_owned(), 7, 13),
            ("<break>".to_owned(), 13, 14),
            ("e\u{301} \u{1F600} \u{FFFC} ".to_owned(), 14, 22),
            ("done".to_owned(), 22, 26),
        ]
    );
}

#[test]
fn images_and_fields_report_metadata_and_cached_results() {
    let content = export(&fixture::principal_docx(), &options(RevisionView::Accepted));
    let image = inlines(block(&content, "00000010"))
        .iter()
        .find_map(|inline| match &inline.content {
            InlineKind::Image {
                alt_text,
                relationship_id,
                part,
                ..
            } => Some((alt_text.clone(), relationship_id.clone(), part.clone())),
            _ => None,
        })
        .unwrap();
    assert_eq!(
        image,
        (
            Some("Company logo".to_owned()),
            Some("rIdImage".to_owned()),
            Some("word/media/logo.png".to_owned())
        )
    );
    assert!(codes(&content).contains(&DiagnosticCode::ImageDataOmitted));
    let fields: Vec<(String, String)> = inlines(block(&content, "00000011"))
        .iter()
        .filter_map(|inline| match &inline.content {
            InlineKind::Field {
                field_type,
                cached_result,
                ..
            } => Some((
                field_type.clone(),
                match cached_result {
                    CachedResult::Missing => "missing".to_owned(),
                    CachedResult::Inline { inlines } => inlines
                        .iter()
                        .map(|inline| match &inline.content {
                            InlineKind::Text { text } => text.clone(),
                            _ => String::new(),
                        })
                        .collect(),
                    CachedResult::Blocks { .. } => "blocks".to_owned(),
                },
            )),
            _ => None,
        })
        .collect();
    assert_eq!(
        fields,
        [
            ("PAGE".to_owned(), "4".to_owned()),
            ("DATE".to_owned(), "2026".to_owned()),
            ("AUTHOR".to_owned(), "missing".to_owned())
        ]
    );
    let count = |code| {
        codes(&content)
            .into_iter()
            .filter(|found| *found == code)
            .count()
    };
    assert_eq!(count(DiagnosticCode::FieldCachedResult), 4);
    assert_eq!(count(DiagnosticCode::MissingFieldResult), 1);
}

#[test]
fn empty_cached_results_differ_from_missing_ones() {
    let mut parts = fixture::principal_parts();
    let document = parts
        .iter_mut()
        .find(|(name, _)| name == "word/document.xml")
        .unwrap();
    let xml = String::from_utf8(document.1.clone()).unwrap().replace(
        "<w:fldSimple w:instr=\" AUTHOR \"/>",
        "<w:r><w:fldChar w:fldCharType=\"begin\"/></w:r><w:r><w:instrText>AUTHOR</w:instrText></w:r><w:r><w:fldChar w:fldCharType=\"separate\"/></w:r><w:r><w:t></w:t></w:r><w:r><w:fldChar w:fldCharType=\"end\"/></w:r>",
    );
    document.1 = xml.into_bytes();
    let content = export(
        &ooxml_opc::rezip_parts(&parts).unwrap(),
        &options(RevisionView::Accepted),
    );
    let author = inlines(block(&content, "00000011"))
        .iter()
        .find_map(|inline| match &inline.content {
            InlineKind::Field {
                field_type,
                cached_result,
                ..
            } if field_type == "AUTHOR" => Some(cached_result.clone()),
            _ => None,
        })
        .unwrap();
    assert_eq!(author, CachedResult::Inline { inlines: vec![] });
}

#[test]
fn unsupported_content_keeps_placeholders_and_source_provenance() {
    let bytes = fixture::principal_docx();
    let content = export(&bytes, &options(RevisionView::Accepted));
    let raw = body(&content)
        .iter()
        .find(|block| matches!(&block.content, BlockKind::Unsupported { element } if element == "bofx:block"))
        .unwrap();
    let Anchor::SourcePart {
        part,
        part_sha256,
        path,
    } = &raw.anchor
    else {
        panic!("raw blocks carry source provenance")
    };
    assert_eq!(part, "word/document.xml");
    assert_eq!(part_sha256.len(), 64);
    let parts = ooxml_opc::unzip_parts(&bytes).unwrap();
    let xml = &parts
        .iter()
        .find(|(name, _)| name == "word/document.xml")
        .unwrap()
        .1;
    let limits = docx_parse::ParseLimits::default();
    let parsed =
        docx_parse::parse_xml(xml, part, &mut docx_parse::ParseBudget::new(&limits)).unwrap();
    let mut element = parsed.root().unwrap();
    for index in path {
        element = element.child_elements().nth(*index as usize).unwrap();
    }
    assert_eq!(element.name, "bofx:block");
    let elements: Vec<(String, Option<String>)> = inlines(block(&content, "00000015"))
        .iter()
        .filter_map(|inline| match &inline.content {
            InlineKind::Unsupported { element, alt_text } => {
                Some((element.clone(), alt_text.clone()))
            }
            _ => None,
        })
        .collect();
    assert_eq!(
        elements,
        [
            ("wps:wsp".to_owned(), Some("Process arrow".to_owned())),
            ("m:oMath".to_owned(), Some("x=1".to_owned())),
            ("w:drawing".to_owned(), None),
        ]
    );
    let unsupported = content
        .diagnostics
        .iter()
        .filter(|diagnostic| diagnostic.code == DiagnosticCode::UnsupportedContent)
        .count();
    assert_eq!(unsupported, 4);
}

#[test]
fn stories_are_explicit_and_uses_record_every_referencing_section() {
    let bytes = fixture::principal_docx();
    let default = export(&bytes, &options(RevisionView::Accepted));
    assert_eq!(default.stories.len(), 1);
    let omitted: Vec<&str> = default
        .diagnostics
        .iter()
        .filter(|diagnostic| diagnostic.code == DiagnosticCode::StoriesOmitted)
        .map(|diagnostic| diagnostic.message.as_str())
        .collect();
    assert_eq!(omitted.len(), 5, "{omitted:?}");
    let everything = export(&bytes, &all(RevisionView::Accepted));
    let summary: Vec<(String, StoryKind)> = everything
        .stories
        .iter()
        .map(|story| (story.story.clone(), story.kind))
        .collect();
    assert_eq!(
        summary,
        [
            ("body", StoryKind::Body),
            ("hf:rIdHeader1", StoryKind::Header),
            ("hf:rIdHeader2", StoryKind::Header),
            ("hf:rIdFooter1", StoryKind::Footer),
            ("fn:1", StoryKind::Footnote),
            ("en:1", StoryKind::Endnote),
            ("comment:1", StoryKind::Comment),
            ("comment:2", StoryKind::Comment),
        ]
        .map(|(story, kind)| (story.to_owned(), kind))
    );
    let header = &everything.stories[1];
    assert_eq!(header.part.as_deref(), Some("word/header1.xml"));
    assert_eq!(
        header
            .uses
            .iter()
            .map(|used| used.section_index)
            .collect::<Vec<_>>(),
        [0, 1]
    );
    let comment = everything.stories[6].comment.as_ref().unwrap();
    assert_eq!(
        (comment.author.as_deref(), comment.anchors.len()),
        (Some("Ann"), 1)
    );
    assert!(matches!(
        &everything.stories[7].blocks[0].anchor,
        Anchor::SourcePart { part, path, .. } if part == "word/comments.xml" && path == &[1]
    ));
    let reference = inlines(block(&everything, "00000012"))
        .iter()
        .find_map(|inline| match &inline.content {
            InlineKind::CommentReference { story, .. } => story.clone(),
            _ => None,
        });
    assert_eq!(reference.as_deref(), Some("comment:1"));
}

#[test]
fn limits_truncate_at_whole_root_blocks_within_budget() {
    let bytes = fixture::principal_docx();
    let full = export(&bytes, &all(RevisionView::Markup));
    let length = serde_json::to_string(&full).unwrap().len() as u32;
    let fits = export(
        &bytes,
        &ExportOptions {
            max_bytes: Some(length),
            ..all(RevisionView::Markup)
        },
    );
    assert!(!fits.truncated);
    assert_eq!(fits, full);
    for max_bytes in [1_024, 4_096, length / 2, length - 1] {
        let truncated = export(
            &bytes,
            &ExportOptions {
                max_bytes: Some(max_bytes),
                ..all(RevisionView::Markup)
            },
        );
        assert!(truncated.truncated);
        assert!(serde_json::to_string(&truncated).unwrap().len() <= max_bytes as usize);
        assert_eq!(
            truncated
                .diagnostics
                .last()
                .map(|diagnostic| diagnostic.code),
            Some(DiagnosticCode::Truncated)
        );
    }
    let one = export(
        &bytes,
        &ExportOptions {
            max_blocks: Some(1),
            ..options(RevisionView::Accepted)
        },
    );
    assert!(one.truncated);
    assert_eq!(body(&one).len(), 1);
    let table_position = body(&full)
        .iter()
        .position(|block| matches!(block.content, BlockKind::Table { .. }))
        .unwrap();
    let before_table = body(&full)[..table_position].len() as u32;
    let stopped = export(
        &bytes,
        &ExportOptions {
            max_blocks: Some(before_table + 5),
            ..options(RevisionView::Markup)
        },
    );
    assert_eq!(
        body(&stopped).len() as u32,
        before_table,
        "a table counts every nested block and is never split"
    );
}

#[test]
fn unusable_limits_are_refused_as_data() {
    let bytes = fixture::principal_docx();
    let refused = |options: ExportOptions| match export_docx_structured(&bytes, &options) {
        Err(docx_edit::structured::ExportError::Refused(failure)) => failure.code,
        other => panic!("expected a refusal, got {other:?}"),
    };
    let base = options(RevisionView::Accepted);
    assert_eq!(
        refused(ExportOptions {
            max_bytes: Some(100),
            ..base.clone()
        }),
        ExportFailureCode::InvalidOptions
    );
    assert_eq!(
        refused(ExportOptions {
            max_bytes: Some(u32::MAX),
            ..base.clone()
        }),
        ExportFailureCode::LimitExceeded
    );
    assert_eq!(
        refused(ExportOptions {
            max_blocks: Some(0),
            ..base.clone()
        }),
        ExportFailureCode::InvalidOptions
    );
    assert_eq!(
        refused(ExportOptions {
            stories: Some(Vec::new()),
            ..base.clone()
        }),
        ExportFailureCode::InvalidOptions
    );
    let doc = open(&bytes);
    let refusal = doc
        .export_structured(&ExportOptions {
            max_blocks: Some(0),
            ..base.clone()
        })
        .unwrap_err();
    assert_eq!(refusal.version, doc.version());
    let wire = serde_json::to_value(&refusal.failure).unwrap();
    assert_eq!(wire["target"], serde_json::Value::Null);
    assert!(wire.as_object().unwrap().contains_key("target"));
    let empty = EditingDoc::new(6);
    let refusal = empty.export_structured(&base).unwrap_err();
    assert_eq!(refusal.failure.code, ExportFailureCode::Unsupported);
    assert_eq!(
        serde_json::to_value(&refusal.failure).unwrap()["code"],
        "unsupported"
    );
    assert!(
        serde_json::from_str::<ExportOptions>(r#"{"revisionView":"accepted","maxBytes":-1}"#)
            .is_err()
    );
    assert!(
        serde_json::from_str::<ExportOptions>(r#"{"revisionView":"accepted","pages":true}"#)
            .is_err()
    );
}

#[test]
fn exports_are_deterministic_across_replicas_and_chunking() {
    let bytes = fixture::principal_docx();
    let options = all(RevisionView::Markup);
    let first = export(&bytes, &options);
    assert_eq!(export(&bytes, &options), first);
    let left = EditingDoc::new(1);
    let right = EditingDoc::new(987_654);
    seed_from_docx(&left, &bytes).unwrap();
    seed_from_docx(&right, &bytes).unwrap();
    let left_read = left.export_structured(&options).unwrap();
    let right_read = right.export_structured(&options).unwrap();
    assert_ne!(left_read.version, right_read.version);
    assert_eq!(left_read.content, right_read.content);
    assert_eq!(left_read.content.anchor_scope, AnchorScope::Session);
    let mut snapshot = left_read.content.clone();
    snapshot.anchor_scope = AnchorScope::Snapshot;
    assert_eq!(snapshot, first);
    let paragraph = left
        .read_paragraphs(&ReadParagraphsRequest {
            story: None,
            para_ids: Some(vec!["00000017".to_owned()]),
            view: EditTextView::Accepted,
        })
        .unwrap();
    assert_eq!(paragraph.paragraphs[0].text, "Section two");
    let start = left.paragraph_mark_position("00000017").unwrap().index - 11;
    let span = docx_edit::StoryRange::new("body", start + 2, start + 5);
    let ctx = docx_edit::EditCtx::local("", "");
    left.toggle_format(&ctx, span.clone(), docx_edit::SimpleFormat::Bold)
        .unwrap();
    left.toggle_format(&ctx, span, docx_edit::SimpleFormat::Bold)
        .unwrap();
    let rechunked = left.export_structured(&options).unwrap();
    assert_ne!(rechunked.version, left_read.version);
    assert_eq!(rechunked.content, left_read.content);
}

fn replace(doc: &EditingDoc, history: &UndoSession, para_id: &str, text: &str) {
    let request = EditRequest {
        expect_version: doc.version(),
        source: EditSource::Host,
        history: EditHistory::Separate,
        steps: vec![EditStep::new(EditOperation::ReplaceText {
            target: TextTarget::Paragraph(ParagraphTarget {
                story: "body".to_owned(),
                para_id: para_id.to_owned(),
            }),
            text: text.to_owned(),
        })],
    };
    doc.apply_edits(&request, history).unwrap().unwrap();
}

#[test]
fn exporting_is_read_only() {
    let bytes = fixture::principal_docx();
    let doc = open(&bytes);
    let control = open(&bytes);
    let history = UndoSession::default();
    let control_history = UndoSession::default();
    history.track(&doc);
    control_history.track(&control);
    for (doc, history) in [(&doc, &history), (&control, &control_history)] {
        replace(doc, history, "00000017", "First edit");
        replace(doc, history, "00000013", "Second edit");
        assert!(history.undo());
    }
    assert!(history.can_undo() && history.can_redo());
    let updates = Rc::new(Cell::new(0));
    let counted = Rc::clone(&updates);
    let _subscription = doc
        .yrs_doc()
        .observe_update_v1(move |_, _| counted.set(counted.get() + 1))
        .unwrap();
    let version = doc.version();
    let state = doc.encode_state_as_update_v1();
    let vector = doc.encode_state_vector_v1();
    for view in [
        RevisionView::Accepted,
        RevisionView::Original,
        RevisionView::Markup,
    ] {
        let read = doc.export_structured(&all(view)).unwrap();
        assert_eq!(read.version, version);
        doc.export_markdown(&all(view)).unwrap();
        doc.export_structured(&ExportOptions {
            max_bytes: Some(1_024),
            ..all(view)
        })
        .unwrap();
    }
    assert_eq!(doc.version(), version);
    assert_eq!(doc.encode_state_as_update_v1(), state);
    assert_eq!(doc.encode_state_vector_v1(), vector);
    assert_eq!(updates.get(), 0);
    assert!(history.can_undo() && history.can_redo());
    let range = |doc: &EditingDoc| {
        let start = doc.paragraph_mark_position("00000017").unwrap().index - 3;
        docx_edit::StoryRange::new("body", start, start + 2)
    };
    assert_eq!(
        doc.add_comment(&[range(&doc)], "A", "", Any::Null).unwrap(),
        control
            .add_comment(&[range(&control)], "A", "", Any::Null)
            .unwrap(),
        "exporting allocates no ids"
    );
    assert!(history.redo() && control_history.redo());
    assert!(history.undo() && control_history.undo());
    assert_eq!(
        doc.story_segments("body").unwrap(),
        control.story_segments("body").unwrap(),
        "the history replays exactly as a replica's that never exported"
    );
}

#[test]
fn live_edits_appear_and_removed_content_does_not_return() {
    let doc = open(&fixture::principal_docx());
    let history = UndoSession::default();
    let target = |para_id: &str| ParagraphTarget {
        story: "body".to_owned(),
        para_id: para_id.to_owned(),
    };
    let request = EditRequest {
        expect_version: doc.version(),
        source: EditSource::Host,
        history: EditHistory::Separate,
        steps: vec![
            EditStep::new(EditOperation::ReplaceText {
                target: TextTarget::Paragraph(target("00000017")),
                text: "Rewritten".to_owned(),
            }),
            EditStep::new(EditOperation::InsertText {
                target: TextTarget::Paragraph(target("00000015")),
                at: TargetEdge::Start,
                text: "Prefix ".to_owned(),
            }),
        ],
    };
    doc.apply_edits(&request, &history).unwrap().unwrap();
    let content = doc
        .export_structured(&options(RevisionView::Accepted))
        .unwrap()
        .content;
    assert_eq!(text(block(&content, "00000017")), "Rewritten");
    let shifted: Vec<(String, Option<u32>)> = inlines(block(&content, "00000015"))
        .iter()
        .map(|inline| {
            let label = match &inline.content {
                InlineKind::Text { text } => text.clone(),
                InlineKind::Unsupported { element, .. } => element.clone(),
                _ => String::new(),
            };
            let offset = match &inline.anchor {
                Anchor::Range(range) => Some(range.start.offset),
                _ => None,
            };
            (label, offset)
        })
        .collect();
    assert_eq!(
        shifted.last(),
        Some(&("end".to_owned(), Some(41))),
        "{shifted:?}"
    );
    assert_eq!(
        shifted[shifted.len() - 2],
        ("w:drawing".to_owned(), None),
        "the pinned omission moves with the text around it"
    );
    let delete = EditRequest {
        expect_version: doc.version(),
        source: EditSource::Host,
        history: EditHistory::Separate,
        steps: vec![EditStep::new(EditOperation::DeleteParagraphs {
            story: "body".to_owned(),
            first_para_id: "00000017".to_owned(),
            last_para_id: "00000017".to_owned(),
        })],
    };
    doc.apply_edits(&delete, &history)
        .unwrap()
        .expect("the paragraph can be deleted");
    let content = doc
        .export_structured(&options(RevisionView::Accepted))
        .unwrap()
        .content;
    assert!(body(&content).iter().all(|block| !matches!(
        &block.anchor,
        Anchor::Paragraph { para_id, .. } if para_id == "00000017"
    )));
    assert!(!text_of(&content).contains("Rewritten"));
}

fn text_of(content: &DocxStructuredContent) -> String {
    serde_json::to_string(content).unwrap()
}

#[test]
fn formatting_marks_can_be_left_out_without_losing_links() {
    let bytes = fixture::principal_docx();
    let content = export(
        &bytes,
        &ExportOptions {
            include_formatting: Some(false),
            ..options(RevisionView::Accepted)
        },
    );
    assert!(
        inlines(block(&content, "00000001"))
            .iter()
            .all(|inline| inline.marks.is_none())
    );
    let link = inlines(block(&content, "00000013"))[1]
        .link
        .clone()
        .unwrap();
    assert_eq!(
        (link.href.as_str(), link.title.as_deref()),
        ("https://example.com/docs", Some("Visit"))
    );
    assert!(codes(&content).contains(&DiagnosticCode::FormattingOmitted));
    let formatted = export(&bytes, &options(RevisionView::Accepted));
    assert_eq!(
        inlines(block(&formatted, "00000001"))[0].marks.as_deref(),
        Some([FormattingMark::Bold].as_slice())
    );
}

#[test]
fn sessions_without_source_context_still_export_what_they_hold() {
    let doc = EditingDoc::new(9);
    doc.create_story("body", "Plain *text*", "Heading1", "left")
        .unwrap();
    let content = doc
        .export_structured(&options(RevisionView::Accepted))
        .unwrap()
        .content;
    assert!(codes(&content).contains(&DiagnosticCode::ProvenanceUnavailable));
    let block = &body(&content)[0];
    assert!(matches!(
        &block.content,
        BlockKind::Heading { heading, .. } if heading.source == OutlineSource::BuiltinStyleId { style_id: "Heading1".to_owned() }
    ));
    let markdown = render_docx_markdown(&content, &MarkdownOptions::default()).unwrap();
    assert!(markdown.markdown.contains("# Plain \\*text\\*"));
}

#[test]
fn duplicate_paragraph_ids_are_reported_as_ambiguous() {
    let doc = EditingDoc::new(5);
    let para = doc.create_story("body", "One", "Normal", "left").unwrap();
    doc.apply_raw_ops(
        "body",
        vec![
            RawOp::Insert {
                index: 0,
                text: "Two".to_owned(),
                attrs: Default::default(),
            },
            RawOp::InsertEmbed {
                index: 3,
                kind: "pilcrow".to_owned(),
                payload: vec![("paraId".to_owned(), Any::from(para.as_str()))],
                attrs: Default::default(),
            },
        ],
        &docx_edit::EditCtx::local("", ""),
    )
    .unwrap();
    let content = doc
        .export_structured(&options(RevisionView::Accepted))
        .unwrap()
        .content;
    let ambiguous = content
        .diagnostics
        .iter()
        .filter(|diagnostic| diagnostic.code == DiagnosticCode::AmbiguousIdentity)
        .count();
    assert_eq!(ambiguous, 2);
    for block in body(&content) {
        assert!(matches!(
            &block.anchor,
            Anchor::Unlocated {
                reason: UnlocatedReason::DuplicateParagraphId,
                ..
            }
        ));
        assert!(
            inlines(block)
                .iter()
                .all(|inline| inline.anchor == block.anchor)
        );
    }
}

#[test]
fn markdown_carries_markers_anchors_and_its_own_budget() {
    let bytes = fixture::principal_docx();
    let markdown = export_docx_markdown(&bytes, &all(RevisionView::Markup)).unwrap();
    let markers = markdown.markdown.matches("<!-- docx-export:").count();
    assert_eq!(markers, markdown.anchors.len());
    assert!(markdown.markdown.contains("<ins data-author=\"Ann\""));
    assert!(markdown.markdown.contains("<table>"));
    assert!(markdown.markdown.contains("![Company logo]()"));
    let small = render_docx_markdown(
        &export(&bytes, &all(RevisionView::Markup)),
        &MarkdownOptions {
            max_bytes: Some(1_024),
        },
    )
    .unwrap();
    assert!(small.truncated && small.markdown.len() <= 1_024);
    assert_eq!(
        small.anchors.len(),
        small.markdown.matches("<!-- docx-export:").count()
    );
    let mut content = export(&bytes, &options(RevisionView::Accepted));
    content.schema_version = 1;
    let mut value = serde_json::to_value(&content).unwrap();
    value["schemaVersion"] = serde_json::json!(2);
    assert!(serde_json::from_value::<DocxStructuredContent>(value).is_err());
}
