#[path = "support/page_fixture.rs"]
mod fixture;

use std::collections::{BTreeMap, BTreeSet, HashMap, HashSet};

use docx_edit::structured::{
    Anchor, AtomCoverage, Block, BlockKind, DocxLayoutMap, DocxPagedStructuredContent,
    ExportFailureCode, FragmentSlice, InlineKind, MarkdownOptions, NotePlacement, NumberingStatus,
    OccurrenceRegion, PageAnnotations, PageDiagnosticCode, PageExportOptions, PageFragment,
    PageMarkdownOptions, RevisionView, StorySelection, render_docx_markdown,
    render_docx_markdown_with_pages,
};
use docx_edit::{
    EditHistory, EditOperation, EditRequest, EditSource, EditStep, EditTextView, EngineSession,
    ParagraphTarget, ReadParagraphsRequest, TextTarget, UndoSession,
};

const LAID_OUT: [StorySelection; 5] = [
    StorySelection::Body,
    StorySelection::Headers,
    StorySelection::Footers,
    StorySelection::Footnotes,
    StorySelection::Endnotes,
];

fn options(view: RevisionView) -> PageExportOptions {
    PageExportOptions {
        stories: Some(LAID_OUT.to_vec()),
        ..PageExportOptions::new(view)
    }
}

fn export(
    engine: &EngineSession,
    options: &PageExportOptions,
) -> DocxPagedStructuredContent<DocxLayoutMap> {
    engine
        .export_structured_with_pages(options)
        .unwrap()
        .content
}

fn refusal(engine: &EngineSession, options: &PageExportOptions) -> ExportFailureCode {
    engine
        .export_structured_with_pages(options)
        .unwrap_err()
        .failure
        .code
}

/// Every exported block by id, nested ones included.
fn blocks(content: &DocxPagedStructuredContent<DocxLayoutMap>) -> HashMap<String, Block> {
    fn visit(block: &Block, out: &mut HashMap<String, Block>) {
        out.insert(block.id.clone(), block.clone());
        match &block.content {
            BlockKind::Table { table } => {
                for cell in table.rows.iter().flat_map(|row| &row.cells) {
                    for block in &cell.blocks {
                        visit(block, out);
                    }
                }
            }
            BlockKind::ContentControl { blocks, .. } => {
                for block in blocks {
                    visit(block, out);
                }
            }
            _ => {}
        }
    }
    let mut out = HashMap::new();
    for story in &content.structured.stories {
        for block in &story.blocks {
            visit(block, &mut out);
        }
    }
    out
}

fn fragments_of<'a>(map: &'a DocxLayoutMap, node: &str) -> Vec<&'a PageFragment> {
    map.fragments
        .iter()
        .filter(|fragment| fragment.node_id == node)
        .collect()
}

fn text_range(fragment: &PageFragment) -> Option<(u32, u32, EditTextView)> {
    match &fragment.slice {
        FragmentSlice::Text { range } => Some((range.start.offset, range.end.offset, range.view)),
        _ => None,
    }
}

fn body_paragraph<'a>(
    content: &'a DocxPagedStructuredContent<DocxLayoutMap>,
    para_id: &str,
) -> &'a Block {
    content.structured.stories[0]
        .blocks
        .iter()
        .find(
            |block| matches!(&block.anchor, Anchor::Paragraph { para_id: id, .. } if id == para_id),
        )
        .unwrap()
}

#[test]
fn fixture_matches_its_builder() {
    let committed = ooxml_opc::unzip_parts(&fixture::pages_docx()).unwrap();
    assert_eq!(committed, fixture::page_parts(true));
}

#[test]
fn pages_carry_physical_and_displayed_numbers() {
    let (engine, _) = fixture::laid_out(&fixture::pages_docx(), 7);
    let map = export(&engine, &options(RevisionView::Markup)).layout;
    let labels: Vec<_> = map
        .pages
        .iter()
        .map(|page| page.displayed_label.as_str())
        .collect();
    assert_eq!(labels, ["i", "ii", "iii", "iv", "1", "2", "C", "4"]);
    let indexes: Vec<_> = map.pages.iter().map(|page| page.page_index).collect();
    assert_eq!(indexes, (0..8).collect::<Vec<_>>());
    assert_eq!(map.pages[0].displayed_number, map.pages[4].displayed_number);
    let fillers: Vec<_> = map
        .pages
        .iter()
        .filter(|page| page.parity_filler)
        .map(|page| page.page_index)
        .collect();
    assert_eq!(fillers, [5]);
    assert_eq!(map.pages[5].section_index, 1);
    assert_eq!(map.pages[6].section_index, 2);
    assert_eq!(map.pages[6].numbering_format, "upperLetter");
    assert_eq!(map.pages[7].numbering_status, NumberingStatus::Fallback);
    assert!(
        map.pages[..7]
            .iter()
            .all(|page| page.numbering_status == NumberingStatus::Resolved)
    );
    assert!(map.diagnostics.iter().any(|diagnostic| {
        diagnostic.code == PageDiagnosticCode::UnsupportedNumbering
            && diagnostic.page_index == Some(7)
    }));
    assert_eq!(
        (map.pages[0].size.width, map.pages[0].size.height),
        (480.0, 384.0)
    );
    assert!(map.occurrences.iter().any(
        |occurrence| occurrence.page_index == 5 && occurrence.region == OccurrenceRegion::Body
    ));
    assert!(
        !map.occurrences
            .iter()
            .any(|occurrence| occurrence.page_index == 5
                && occurrence.region != OccurrenceRegion::Body),
        "a parity filler shows no header, footer or note"
    );
}

#[test]
fn paragraphs_split_across_pages_have_exact_ranges() {
    let bytes = fixture::unrevised_docx();
    let (engine, _) = fixture::laid_out(&bytes, 7);
    let content = export(&engine, &options(RevisionView::Accepted));
    let map = &content.layout;
    let read = engine
        .doc()
        .read_paragraphs(&ReadParagraphsRequest {
            story: None,
            para_ids: None,
            view: EditTextView::Accepted,
        })
        .unwrap();
    let texts: HashMap<_, _> = read
        .paragraphs
        .iter()
        .map(|paragraph| (paragraph.para_id.clone(), paragraph.clone()))
        .collect();
    let exported = blocks(&content);
    let mut ordered: Vec<_> = exported.iter().collect();
    ordered.sort_by_key(|(_, block)| match &block.anchor {
        Anchor::Paragraph { para_id, .. } => para_id.clone(),
        _ => String::new(),
    });
    let mut checked = 0;
    let mut text_breaks: Vec<(String, u32, u32, u32)> = Vec::new();
    for (id, block) in ordered {
        let (BlockKind::Paragraph { paragraph } | BlockKind::Heading { paragraph, .. }) =
            &block.content
        else {
            continue;
        };
        let Anchor::Paragraph { story, para_id } = &block.anchor else {
            continue;
        };
        if story != "body" {
            continue;
        }
        let projected = &texts[para_id];
        let own = fragments_of(map, id);
        assert!(!own.is_empty(), "paragraph {para_id} is placed");
        let pages: Vec<_> = own.iter().map(|fragment| fragment.page_index).collect();
        assert!(
            pages.windows(2).all(|pair| pair[0] < pair[1]),
            "{para_id}: {pages:?}"
        );
        assert!(!own.first().unwrap().continued_from_previous);
        assert!(!own.last().unwrap().continued_on_next);
        for pair in own.windows(2) {
            assert!(pair[0].continued_on_next && pair[1].continued_from_previous);
        }
        let units: Vec<u16> = projected.text.encode_utf16().collect();
        let mut covered = vec![0u32; units.len()];
        let mut rebuilt = vec![None; units.len()];
        for inline in &paragraph.inlines {
            let parts = fragments_of(map, &inline.id);
            let Anchor::Range(anchor) = &inline.anchor else {
                continue;
            };
            assert!(!parts.is_empty(), "{para_id}: {} is placed", inline.id);
            match &inline.content {
                InlineKind::Text { .. } | InlineKind::Tab => {
                    let mut previous_end = anchor.start.offset;
                    for part in &parts {
                        let (start, end, view) = text_range(part).unwrap();
                        assert_eq!(view, EditTextView::Accepted);
                        assert_eq!(start, previous_end, "{para_id}: text slices continue");
                        for offset in [start, end] {
                            let offset = offset as usize;
                            assert!(
                                offset == units.len() || !(0xDC00..0xE000).contains(&units[offset]),
                                "{para_id}: no surrogate pair is split at {offset}"
                            );
                        }
                        for unit in start..end {
                            covered[unit as usize] += 1;
                            rebuilt[unit as usize] = Some(units[unit as usize]);
                        }
                        text_breaks.push((para_id.clone(), part.page_index, start, end));
                        previous_end = end;
                    }
                    assert_eq!(
                        previous_end, anchor.end.offset,
                        "{para_id}: text slices cover"
                    );
                }
                _ => {
                    assert_eq!(parts.len(), 1, "{para_id}: one atom fragment");
                    assert!(matches!(
                        parts[0].slice,
                        FragmentSlice::Atom {
                            coverage: AtomCoverage::Whole,
                            ..
                        }
                    ));
                    covered[anchor.start.offset as usize] += 1;
                    rebuilt[anchor.start.offset as usize] =
                        Some(units[anchor.start.offset as usize]);
                }
            }
        }
        assert!(
            covered.iter().all(|count| *count == 1),
            "{para_id}: {covered:?}"
        );
        let rebuilt: Vec<u16> = rebuilt.into_iter().map(Option::unwrap).collect();
        assert_eq!(
            String::from_utf16(&rebuilt).unwrap(),
            projected.text,
            "{para_id}: the slices rebuild the paragraph"
        );
        checked += 1;
    }
    assert!(checked > 8);
    let units: Vec<u16> = texts["00000002"].text.encode_utf16().collect();
    let boundaries: Vec<u32> = text_breaks
        .windows(2)
        .filter(|pair| pair[0].0 == "00000002" && pair[1].0 == "00000002" && pair[0].1 < pair[1].1)
        .map(|pair| pair[1].2)
        .collect();
    assert_eq!(boundaries.len(), 1, "the long paragraph breaks once");
    let at = boundaries[0] as usize;
    assert!(
        units[at - 2..at] == [0xD83D, 0xDE00] && units[at..at + 2] == [0xD83D, 0xDE00],
        "the page break falls between two emoji"
    );
    let split = body_paragraph(&content, "00000002");
    let pages: HashSet<_> = fragments_of(map, &split.id)
        .iter()
        .map(|fragment| fragment.page_index)
        .collect();
    assert_eq!(pages.len(), 2, "the long paragraph spans two pages");
    let text = &texts["00000002"].text;
    assert!(text.contains('\u{1F600}') && text.contains('\u{FFFC}') && text.contains('\t'));
    let literal = text.encode_utf16().position(|unit| unit == 0xFFFC).unwrap() as u32;
    assert!(
        map.fragments.iter().any(|fragment| matches!(
            text_range(fragment),
            Some((start, end, _)) if start <= literal && literal < end
        )),
        "a literal U+FFFC is text, not an atom"
    );
}

#[test]
fn atoms_keep_their_owning_ranges() {
    let (engine, _) = fixture::laid_out(&fixture::unrevised_docx(), 7);
    let content = export(&engine, &options(RevisionView::Accepted));
    let paragraph = body_paragraph(&content, "00000003");
    let BlockKind::Paragraph { paragraph: data } = &paragraph.content else {
        panic!("paragraph expected");
    };
    let mut kinds = Vec::new();
    for inline in &data.inlines {
        let parts = fragments_of(&content.layout, &inline.id);
        if let FragmentSlice::Atom { range, coverage } = &parts[0].slice {
            assert_eq!(*coverage, AtomCoverage::Whole);
            assert_eq!(
                range.clone().map(Anchor::Range),
                Some(inline.anchor.clone())
            );
            kinds.push(match &inline.content {
                InlineKind::Field { .. } => "field",
                InlineKind::Break { .. } => "break",
                InlineKind::NoteReference { .. } => "note",
                InlineKind::ContentControl { .. } => "control",
                _ => "other",
            });
        }
    }
    assert_eq!(kinds, ["field", "break", "note", "control"]);
}

#[test]
fn tables_report_row_windows_repeated_headers_and_split_rows() {
    let (engine, _) = fixture::laid_out(&fixture::unrevised_docx(), 7);
    let content = export(&engine, &options(RevisionView::Accepted));
    let map = &content.layout;
    let table = content.structured.stories[0]
        .blocks
        .iter()
        .find(|block| matches!(block.content, BlockKind::Table { .. }))
        .unwrap();
    let windows: Vec<_> = fragments_of(map, &table.id)
        .into_iter()
        .map(|fragment| {
            let FragmentSlice::Table { rows } = &fragment.slice else {
                panic!("table slice expected");
            };
            (fragment.page_index, rows.clone())
        })
        .collect();
    assert_eq!(windows.len(), 3);
    let (_, first) = &windows[0];
    assert!(first.iter().all(|row| !row.repeated_header));
    let BlockKind::Table { table: data } = &table.content else {
        unreachable!()
    };
    let split = data.rows.len() as u32 - 1;
    assert!(first.iter().all(|row| row.row_index < split));
    for (index, (_, rows)) in windows.iter().enumerate().skip(1) {
        assert_eq!(rows[0].row_index, 0);
        assert!(
            rows[0].repeated_header,
            "carried fragments repeat the header row"
        );
        let tall = rows.iter().find(|row| row.row_index == split).unwrap();
        assert_eq!(tall.continued_from_previous, index == 2);
        assert_eq!(tall.continued_on_next, index == 1);
    }
    let tall_cell = |row: usize| {
        let BlockKind::Table { table } = &table.content else {
            unreachable!()
        };
        table.rows[row].cells[0].blocks.clone()
    };
    let cell_pages: Vec<u32> = tall_cell(split as usize)
        .iter()
        .map(|block| fragments_of(map, &block.id)[0].page_index)
        .collect();
    assert!(cell_pages.windows(2).all(|pair| pair[0] <= pair[1]));
    assert!(
        cell_pages.first() < cell_pages.last(),
        "the row's lines split"
    );
    for block in tall_cell(split as usize) {
        assert_eq!(
            fragments_of(map, &block.id).len(),
            1,
            "each line is placed once"
        );
    }
    let header = &tall_cell(0)[0];
    let header_pages: Vec<_> = fragments_of(map, &header.id)
        .iter()
        .map(|fragment| (fragment.page_index, fragment.repeated_table_header))
        .collect();
    assert_eq!(header_pages.len(), 3);
    assert!(!header_pages[0].1 && header_pages[1].1 && header_pages[2].1);
    for fragment in &map.fragments {
        if fragment.repeated_table_header {
            assert!(!fragment.continued_from_previous && !fragment.continued_on_next);
        }
    }
    let merged = &tall_cell(1)[0];
    assert_eq!(fragments_of(map, &merged.id).len(), 1);
    assert!(
        !map.diagnostics
            .iter()
            .any(|diagnostic| diagnostic.code == PageDiagnosticCode::UnsupportedNoteLayout)
    );
}

/// A body whose table row splits across pages with footnote 1 referenced at the end of the
/// row's tall cell.
fn late_reference_body() -> String {
    let tall: String = (0..24)
        .map(|index| {
            fixture::p(
                &format!("300000{index:02X}"),
                &fixture::r(&format!("Tall line {index}")),
            )
        })
        .collect();
    let reference = fixture::p(
        "30000100",
        &format!(
            r#"{}<w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteReference w:id="1"/></w:r>"#,
            fixture::r("Late reference")
        ),
    );
    format!(
        r#"{}<w:tbl><w:tblPr><w:tblW w:w="4800" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="4800"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="4800" w:type="dxa"/></w:tcPr>{tall}{reference}</w:tc></w:tr></w:tbl>{}<w:sectPr><w:pgSz w:w="7200" w:h="5760"/><w:pgMar w:top="720" w:right="720" w:bottom="720" w:left="720" w:header="300" w:footer="300" w:gutter="0"/></w:sectPr>"#,
        fixture::p("00000001", &fixture::r(&"lead words ".repeat(30))),
        fixture::p("00000002", &fixture::r("After the table")),
    )
}

#[test]
fn a_footnote_away_from_its_reference_in_a_split_row_is_refused() {
    let (engine, _) = fixture::laid_out(&fixture::with_body(&late_reference_body()), 7);
    for stories in [
        None,
        Some(vec![StorySelection::Body]),
        Some(vec![StorySelection::Headers]),
    ] {
        let mut selected = PageExportOptions::new(RevisionView::Markup);
        selected.stories = stories;
        let failure = engine
            .export_structured_with_pages(&selected)
            .unwrap_err()
            .failure;
        assert_eq!(failure.code, ExportFailureCode::Unsupported);
        assert!(
            failure.message.contains("Footnote 1"),
            "{}",
            failure.message
        );
    }
}

#[test]
fn headers_and_footers_are_one_story_with_occurrences_per_page() {
    let (engine, _) = fixture::laid_out(&fixture::unrevised_docx(), 7);
    let content = export(&engine, &options(RevisionView::Accepted));
    let map = &content.layout;
    let stories: Vec<_> = content
        .structured
        .stories
        .iter()
        .filter(|story| story.story.starts_with("hf:"))
        .map(|story| story.story.as_str())
        .collect();
    assert_eq!(stories, ["hf:rIdHeader1", "hf:rIdHeader2", "hf:rIdFooter1"]);
    let shown = |story: &str| {
        map.occurrences
            .iter()
            .filter(|occurrence| occurrence.story == story)
            .map(|occurrence| occurrence.page_index)
            .collect::<Vec<_>>()
    };
    assert_eq!(shown("hf:rIdHeader1"), [0, 4]);
    assert_eq!(shown("hf:rIdHeader2"), [1, 2, 3, 6, 7]);
    assert_eq!(shown("hf:rIdFooter1"), [1, 2, 3, 6, 7]);
    let first = map
        .occurrences
        .iter()
        .find(|occurrence| occurrence.page_index == 0 && occurrence.story == "hf:rIdHeader1")
        .unwrap();
    assert!(matches!(
        first.region,
        OccurrenceRegion::Header {
            variant: docx_edit::structured::HeaderFooterVariant::First
        }
    ));
    assert_eq!(first.part.as_deref(), Some("word/header1.xml"));
    let running = content
        .structured
        .stories
        .iter()
        .find(|story| story.story == "hf:rIdHeader2")
        .unwrap();
    let block = &running.blocks[0];
    let occurrences: Vec<_> = fragments_of(map, &block.id)
        .iter()
        .map(|fragment| fragment.occurrence_id.clone())
        .collect();
    assert_eq!(occurrences.len(), 5, "one fragment per occurrence");
    assert!(
        fragments_of(map, &block.id)
            .iter()
            .all(|fragment| !fragment.continued_from_previous && !fragment.continued_on_next)
    );
}

#[test]
fn notes_are_placed_where_the_layout_puts_them() {
    let (engine, _) = fixture::laid_out(&fixture::unrevised_docx(), 7);
    let map = export(&engine, &options(RevisionView::Accepted)).layout;
    let notes: BTreeMap<_, _> = map
        .occurrences
        .iter()
        .filter_map(|occurrence| match &occurrence.region {
            OccurrenceRegion::Footnote { note_id, placement }
            | OccurrenceRegion::Endnote { note_id, placement } => Some((
                occurrence.story.clone(),
                (note_id.clone(), *placement, occurrence.page_index),
            )),
            _ => None,
        })
        .collect();
    assert_eq!(
        notes["fn:1"],
        ("1".to_owned(), NotePlacement::PageBottom, 1)
    );
    assert_eq!(
        notes["fn:2"].2, 2,
        "the note is on its reference's page, the split row's first"
    );
    assert_eq!(notes["fn:3"].2, 4);
    assert_eq!(
        notes["en:1"],
        ("1".to_owned(), NotePlacement::DocumentEnd, 7)
    );
    for story in ["fn:1", "fn:2", "fn:3", "en:1"] {
        assert!(
            map.fragments
                .iter()
                .any(|fragment| fragment.occurrence_id.ends_with(&format!(
                    ".{}",
                    story.replace("fn:", "footnote.").replace("en:", "endnote.")
                )))
        );
    }
}

#[test]
fn markup_anchors_and_refused_views() {
    let (engine, _) = fixture::laid_out(&fixture::pages_docx(), 7);
    let content = export(&engine, &options(RevisionView::Markup));
    let revised = body_paragraph(&content, "00000009");
    let BlockKind::Paragraph { paragraph } = &revised.content else {
        panic!("paragraph expected");
    };
    for inline in &paragraph.inlines {
        let Anchor::Range(anchor) = &inline.anchor else {
            continue;
        };
        let parts = fragments_of(&content.layout, &inline.id);
        assert!(!parts.is_empty(), "{} is placed", inline.id);
        let mut previous_end = anchor.start.offset;
        for fragment in parts {
            let (start, end, view) = text_range(fragment).unwrap();
            assert_eq!(view, anchor.view);
            assert_eq!(start, previous_end, "{} slices continue", inline.id);
            previous_end = end;
        }
        assert_eq!(
            previous_end, anchor.end.offset,
            "{} slices cover",
            inline.id
        );
    }
    let views: HashSet<_> = paragraph
        .inlines
        .iter()
        .filter_map(|inline| match &inline.anchor {
            Anchor::Range(range) => Some(range.view),
            _ => None,
        })
        .collect();
    assert_eq!(views.len(), 2, "deleted text keeps original-view ranges");
    for view in [RevisionView::Accepted, RevisionView::Original] {
        assert_eq!(
            refusal(&engine, &options(view)),
            ExportFailureCode::UnsupportedRevisionLayout
        );
        let mut body_only = PageExportOptions::new(view);
        body_only.stories = Some(vec![StorySelection::Footnotes]);
        assert_eq!(
            refusal(&engine, &body_only),
            ExportFailureCode::UnsupportedRevisionLayout,
            "revisions in stories left out of the export still refuse"
        );
    }
    let (clean, _) = fixture::laid_out(&fixture::unrevised_docx(), 7);
    let accepted = export(&clean, &options(RevisionView::Accepted)).layout;
    let original = export(&clean, &options(RevisionView::Original)).layout;
    assert_eq!(accepted.fragments.len(), original.fragments.len());
    assert_eq!(accepted.pages, original.pages);
}

fn replace(engine: &EngineSession, history: &UndoSession, para_id: &str, text: &str) {
    replace_in(engine, history, "body", para_id, text);
}

fn replace_in(
    engine: &EngineSession,
    history: &UndoSession,
    story: &str,
    para_id: &str,
    text: &str,
) {
    let doc = engine.doc();
    let request = EditRequest {
        expect_version: doc.version(),
        source: EditSource::Host,
        history: EditHistory::Separate,
        steps: vec![EditStep::new(EditOperation::ReplaceText {
            target: TextTarget::Paragraph(ParagraphTarget {
                story: story.to_owned(),
                para_id: para_id.to_owned(),
            }),
            text: text.to_owned(),
        })],
    };
    doc.apply_edits(&request, history).unwrap().unwrap();
}

#[test]
fn stale_layouts_are_refused() {
    let bytes = fixture::unrevised_docx();
    let unlaid = EngineSession::new(9);
    docx_edit::seed_from_docx(unlaid.doc(), &bytes).unwrap();
    assert_eq!(
        refusal(&unlaid, &options(RevisionView::Markup)),
        ExportFailureCode::LayoutUnavailable
    );
    let (engine, request) = fixture::laid_out(&bytes, 7);
    let layout_version = export(&engine, &options(RevisionView::Markup))
        .layout
        .layout_version;
    let mut expected = options(RevisionView::Markup);
    expected.expect_layout_version = Some(layout_version.clone());
    export(&engine, &expected);
    expected.expect_layout_version = Some(format!("{layout_version}x"));
    assert_eq!(refusal(&engine, &expected), ExportFailureCode::StaleLayout);

    let history = UndoSession::default();
    history.track(engine.doc());
    replace(&engine, &history, "00000001", "Page map, edited");
    assert_eq!(
        refusal(&engine, &options(RevisionView::Markup)),
        ExportFailureCode::StaleDocument
    );
    engine
        .layout_document_with_regions_retained_json(&request)
        .unwrap();
    let edited = export(&engine, &options(RevisionView::Markup));
    assert_ne!(edited.layout.layout_version, layout_version);
    assert!(history.undo());
    assert_eq!(
        refusal(&engine, &options(RevisionView::Markup)),
        ExportFailureCode::StaleDocument,
        "undo changes the document"
    );
    engine
        .layout_document_with_regions_retained_json(&request)
        .unwrap();
    export(&engine, &options(RevisionView::Markup));
    assert!(history.redo());
    assert_eq!(
        refusal(&engine, &options(RevisionView::Markup)),
        ExportFailureCode::StaleDocument,
        "redo changes the document"
    );
    engine
        .layout_document_with_regions_retained_json(&request)
        .unwrap();
    let redone = export(&engine, &options(RevisionView::Markup));
    assert_eq!(redone.layout.fragments, edited.layout.fragments);

    let remote = EngineSession::new(11);
    remote
        .doc()
        .apply_update_v1(&engine.doc().encode_state_as_update_v1())
        .unwrap();
    replace(&remote, &UndoSession::default(), "00000001", "Remote");
    engine
        .doc()
        .apply_update_v1(&remote.doc().encode_state_as_update_v1())
        .unwrap();
    assert_eq!(
        refusal(&engine, &options(RevisionView::Markup)),
        ExportFailureCode::StaleDocument,
        "a remote update changes the document"
    );
    engine
        .layout_document_with_regions_retained_json(&request)
        .unwrap();
    export(&engine, &options(RevisionView::Markup));

    docx_layout::clear_measure_fonts();
    docx_layout::register_measure_font(fixture::FONT).unwrap();
    assert_eq!(
        refusal(&engine, &options(RevisionView::Markup)),
        ExportFailureCode::StaleLayout,
        "another session reset the shared font store"
    );
}

#[test]
fn layout_options_and_fonts_are_fingerprinted_by_content() {
    let bytes = fixture::unrevised_docx();
    let (engine, request) = fixture::laid_out(&bytes, 7);
    let base = export(&engine, &options(RevisionView::Markup))
        .layout
        .provenance;
    let mut gap: serde_json::Value = serde_json::from_str(&request).unwrap();
    gap["options"]["pageGap"] = serde_json::json!(80);
    engine
        .layout_document_with_regions_retained_json(&gap.to_string())
        .unwrap();
    let spaced = export(&engine, &options(RevisionView::Markup))
        .layout
        .provenance;
    assert_eq!(
        spaced.options_fingerprint, base.options_fingerprint,
        "the page gap is not layout"
    );
    let mut env = gap.clone();
    env["renderEnv"] = serde_json::json!({"showHiddenText": true});
    engine
        .layout_document_with_regions_retained_json(&env.to_string())
        .unwrap();
    let hidden = export(&engine, &options(RevisionView::Markup))
        .layout
        .provenance;
    assert_ne!(hidden.options_fingerprint, base.options_fingerprint);
    assert_eq!(hidden.font_set_fingerprint, base.font_set_fingerprint);

    let substitute = docx_layout::register_substitute_measure_font(0, "Calibri").unwrap();
    let mut swapped = env.clone();
    for chain in swapped["measurement"]["fontChains"]
        .as_object_mut()
        .unwrap()
        .values_mut()
    {
        *chain = serde_json::json!([substitute, 0]);
    }
    engine
        .layout_document_with_regions_retained_json(&swapped.to_string())
        .unwrap();
    let chained = export(&engine, &options(RevisionView::Markup))
        .layout
        .provenance;
    assert_ne!(chained.font_set_fingerprint, base.font_set_fingerprint);
}

#[test]
fn exporting_is_read_only() {
    let (engine, _) = fixture::laid_out(&fixture::unrevised_docx(), 7);
    let history = UndoSession::default();
    history.track(engine.doc());
    replace(&engine, &history, "00000007", "Short");
    let request = serde_json::to_string(&fixture::region_request(
        &engine,
        &fixture::unrevised_docx(),
        0,
    ))
    .unwrap();
    engine
        .layout_document_with_regions_retained_json(&request)
        .unwrap();
    let version = engine.doc().version();
    let state = engine.doc().encode_state_as_update_v1();
    let stats = engine.stats();
    let mut geometry = options(RevisionView::Markup);
    geometry.include_geometry = Some(true);
    for options in [options(RevisionView::Markup), geometry] {
        let read = engine.export_structured_with_pages(&options).unwrap();
        assert_eq!(read.version, version);
        assert_eq!(read.content.layout.document_version, version);
    }
    assert_eq!(engine.doc().version(), version);
    assert_eq!(engine.doc().encode_state_as_update_v1(), state);
    assert!(history.can_undo() && !history.can_redo());
    let after = engine.stats();
    assert_eq!(after.layout_epoch, stats.layout_epoch);
    assert_eq!(after.frame_epoch, stats.frame_epoch);
    assert_eq!(after.display_builds, stats.display_builds);
}

#[test]
fn snapshot_maps_are_deterministic() {
    let bytes = fixture::unrevised_docx();
    let snapshot = |client| {
        let (engine, _) = fixture::laid_out(&bytes, client);
        let mut options = options(RevisionView::Markup);
        options.include_geometry = Some(true);
        let paged = engine.export_snapshot_with_pages(&options).unwrap();
        serde_json::to_string(&paged).unwrap()
    };
    let first = snapshot(3);
    assert_eq!(first, snapshot(99));
    assert!(!first.contains("documentVersion") && !first.contains("layoutVersion"));
    let (engine, _) = fixture::laid_out(&bytes, 3);
    let session = export(&engine, &options(RevisionView::Markup));
    let snap: DocxPagedStructuredContent<docx_edit::structured::DocxSnapshotLayoutMap> =
        serde_json::from_str(&first).unwrap();
    assert_eq!(snap.layout.pages, session.layout.pages);
    assert_eq!(snap.layout.fragments.len(), session.layout.fragments.len());
}

#[test]
fn limits_truncate_the_map_without_dangling_references() {
    let (engine, _) = fixture::laid_out(&fixture::unrevised_docx(), 7);
    let full = export(&engine, &options(RevisionView::Markup)).layout;
    assert!(!full.truncated);
    for (fragments, bytes) in [(Some(7), None), (None, Some(4_096)), (None, Some(1_024))] {
        let mut limited = options(RevisionView::Markup);
        limited.max_fragments = fragments;
        limited.max_layout_bytes = bytes;
        let map = export(&engine, &limited).layout;
        assert!(map.truncated);
        assert!(serde_json::to_vec(&map).unwrap().len() <= bytes.unwrap_or(u32::MAX) as usize);
        assert!(map.fragments.len() <= fragments.unwrap_or(u32::MAX) as usize);
        assert_eq!(
            map.diagnostics.last().unwrap().code,
            PageDiagnosticCode::Truncated
        );
        let occurrences: HashSet<_> = map
            .occurrences
            .iter()
            .map(|occurrence| &occurrence.id)
            .collect();
        assert!(
            map.fragments
                .iter()
                .all(|fragment| occurrences.contains(&fragment.occurrence_id))
        );
        assert_eq!(&full.fragments[..map.fragments.len()], &map.fragments[..]);
    }
    assert!(
        serde_json::from_str::<PageExportOptions>(r#"{"revisionView":"markup","pages":true}"#)
            .is_err(),
        "unknown options are malformed, not refused"
    );
    for (fragments, bytes, code) in [
        (Some(0), None, ExportFailureCode::InvalidOptions),
        (Some(2_000_000), None, ExportFailureCode::LimitExceeded),
        (None, Some(10), ExportFailureCode::InvalidOptions),
    ] {
        let mut limited = options(RevisionView::Markup);
        limited.max_fragments = fragments;
        limited.max_layout_bytes = bytes;
        assert_eq!(refusal(&engine, &limited), code);
    }
}

#[test]
fn markdown_page_markers_are_optional_and_checked() {
    let (engine, _) = fixture::laid_out(&fixture::unrevised_docx(), 7);
    let paged = export(&engine, &options(RevisionView::Markup));
    let plain = render_docx_markdown(&paged.structured, &MarkdownOptions::default()).unwrap();
    let unmarked = render_docx_markdown_with_pages(
        &paged.structured,
        PageAnnotations::from(&paged.layout),
        &PageMarkdownOptions::default(),
    )
    .unwrap();
    assert_eq!(unmarked, plain, "markers are off by default");
    let marked = render_docx_markdown_with_pages(
        &paged.structured,
        PageAnnotations::from(&paged.layout),
        &PageMarkdownOptions {
            page_markers: Some(true),
            ..PageMarkdownOptions::default()
        },
    )
    .unwrap();
    assert!(
        marked
            .markdown
            .contains("<!-- docx-export:0 --><!-- docx-pages: 0=i -->")
    );
    assert!(marked.markdown.contains("<!-- docx-pages: 0=i 1=ii -->"));
    assert!(
        marked
            .markdown
            .contains("<!-- docx-pages: 1=ii 2=iii 3=iv 6=C 7=4 -->")
    );
    assert_eq!(marked.anchors, plain.anchors);
    let placed: HashSet<_> = paged
        .layout
        .fragments
        .iter()
        .filter(|fragment| {
            matches!(
                fragment.slice,
                FragmentSlice::Block | FragmentSlice::Table { .. }
            )
        })
        .map(|fragment| &fragment.node_id)
        .collect();
    let section_breaks = blocks(&paged)
        .values()
        .filter(|block| matches!(block.content, BlockKind::SectionBreak { .. }))
        .count();
    let shown: HashSet<_> = paged
        .layout
        .occurrences
        .iter()
        .map(|occurrence| occurrence.story.as_str())
        .collect();
    let unshown: usize = paged
        .structured
        .stories
        .iter()
        .filter(|story| !shown.contains(story.story.as_str()))
        .map(|story| story.blocks.len())
        .sum();
    assert_eq!(
        marked.markdown.matches("<!-- docx-pages:").count(),
        placed.len(),
        "every placed block is annotated, and nothing else"
    );
    assert_eq!(
        marked.markdown.matches("<!-- docx-export:").count(),
        placed.len() + section_breaks + unshown,
        "only section breaks, which take no room, and stories no page shows go unannotated"
    );
    let mut other = paged.structured.clone();
    other.truncated = !other.truncated;
    assert_eq!(
        render_docx_markdown_with_pages(
            &other,
            PageAnnotations::from(&paged.layout),
            &PageMarkdownOptions::default(),
        )
        .unwrap_err()
        .code,
        ExportFailureCode::InvalidOptions,
        "a map from other content is refused"
    );
}

#[test]
fn geometry_is_optional_page_local_and_leaves_ranges_alone() {
    let (engine, _) = fixture::laid_out(&fixture::unrevised_docx(), 7);
    let plain = export(&engine, &options(RevisionView::Markup)).layout;
    assert!(
        plain
            .fragments
            .iter()
            .all(|fragment| fragment.geometry.is_none())
    );
    let mut with_geometry = options(RevisionView::Markup);
    with_geometry.include_geometry = Some(true);
    let map = export(&engine, &with_geometry).layout;
    assert_eq!(map.fragments.len(), plain.fragments.len());
    for (left, right) in map.fragments.iter().zip(&plain.fragments) {
        assert_eq!(left.slice, right.slice);
        assert_eq!(left.page_index, right.page_index);
    }
    let mut measured = 0;
    for fragment in &map.fragments {
        let page = &map.pages[fragment.page_index as usize];
        let Some(geometry) = &fragment.geometry else {
            continue;
        };
        for rect in &geometry.rects {
            assert!(rect.x >= 0.0 && rect.y >= 0.0, "{fragment:?}");
            assert!(rect.x + rect.width <= page.size.width + 0.5, "{fragment:?}");
            assert!(
                rect.y + rect.height <= page.size.height + 0.5,
                "{fragment:?}"
            );
            measured += 1;
        }
    }
    assert!(measured > map.fragments.len() / 2);
}

/// Every text a primitive list paints, a line break wherever the baseline moves.
fn painted(primitives: &[docx_layout::display_list::Primitive]) -> String {
    let mut text = String::from("\n");
    let mut baseline = None;
    for primitive in primitives {
        let (run, y) = match primitive {
            docx_layout::display_list::Primitive::Text(run) => {
                (&run.text, run.baseline_y.as_f64().unwrap_or_default())
            }
            docx_layout::display_list::Primitive::GlyphRun(run) => {
                (&run.text, run.glyphs.first().map_or(0.0, |glyph| glyph.y))
            }
            _ => continue,
        };
        if baseline.is_some_and(|previous: f64| (previous - y).abs() > 0.5) {
            text.push('\n');
        }
        baseline = Some(y);
        text.push_str(run);
    }
    text.push('\n');
    text
}

/// The fixture's numbered words in `text`, however the paint glued them together.
fn numbered_words(text: &str) -> HashSet<String> {
    let mut words = HashSet::new();
    for prefix in ["alpha", "beta", "gamma", "delta", "epsilon", "zeta"] {
        for (at, _) in text.match_indices(prefix) {
            let digits: String = text[at + prefix.len()..]
                .chars()
                .take_while(char::is_ascii_digit)
                .collect();
            if !digits.is_empty() {
                words.insert(format!("{prefix}{digits}"));
            }
        }
    }
    words
}

#[test]
fn page_references_match_the_painted_pages() {
    let bytes = fixture::unrevised_docx();
    let (engine, request) = fixture::laid_out(&bytes, 7);
    let output: serde_json::Value = serde_json::from_str(
        &engine
            .layout_document_with_regions_retained_json(&request)
            .unwrap(),
    )
    .unwrap();
    let request: serde_json::Value = serde_json::from_str(&request).unwrap();
    let extras = serde_json::json!({
        "headersFooters": output["headersFooters"],
        "fontChains": request["measurement"]["fontChains"],
    });
    engine
        .build_display_list_frame(&extras.to_string(), 0)
        .unwrap();
    let pages: Vec<[String; 4]> = engine
        .with_display_list(|list| {
            list.pages
                .iter()
                .map(|page| {
                    [
                        painted(&page.primitives),
                        page.header
                            .as_ref()
                            .map_or_else(String::new, |band| painted(&band.primitives)),
                        page.footer
                            .as_ref()
                            .map_or_else(String::new, |band| painted(&band.primitives)),
                        page.note_areas
                            .iter()
                            .map(|area| painted(&area.primitives))
                            .collect(),
                    ]
                })
                .collect()
        })
        .unwrap();
    let content = export(&engine, &options(RevisionView::Accepted));
    let map = &content.layout;
    assert_eq!(pages.len(), map.pages.len());
    let region = |occurrence: &str| {
        let occurrence = map
            .occurrences
            .iter()
            .find(|candidate| candidate.id == occurrence)
            .unwrap();
        match occurrence.region {
            OccurrenceRegion::Body => 0,
            OccurrenceRegion::Header { .. } => 1,
            OccurrenceRegion::Footer { .. } => 2,
            _ => 3,
        }
    };
    let inline_text: HashMap<String, String> = blocks(&content)
        .values()
        .filter_map(|block| match &block.content {
            BlockKind::Paragraph { paragraph } | BlockKind::Heading { paragraph, .. } => {
                Some(paragraph)
            }
            _ => None,
        })
        .flat_map(|paragraph| &paragraph.inlines)
        .filter_map(|inline| match &inline.content {
            InlineKind::Text { text } => Some((inline.id.clone(), text.clone())),
            _ => None,
        })
        .collect();
    let unique = |token: &str| {
        ["alpha", "beta", "gamma", "delta", "epsilon", "zeta"]
            .iter()
            .any(|prefix| {
                token
                    .strip_prefix(prefix)
                    .is_some_and(|rest| rest.parse::<u32>().is_ok())
            })
    };
    let mut checked = 0;
    for fragment in &map.fragments {
        let (Some((start, end, _)), Some(text), Some(Anchor::Range(anchor))) = (
            text_range(fragment),
            inline_text.get(&fragment.node_id),
            Some(&fragment.anchor),
        ) else {
            continue;
        };
        let units: Vec<u16> = text.encode_utf16().collect();
        let slice = String::from_utf16(
            &units[(start - anchor.start.offset) as usize..(end - anchor.start.offset) as usize],
        )
        .unwrap();
        let here = &pages[fragment.page_index as usize][region(&fragment.occurrence_id)];
        for token in slice.split_whitespace() {
            assert!(
                here.contains(token),
                "{token:?} of {} is painted on page {}",
                fragment.node_id,
                fragment.page_index
            );
            if unique(token) {
                let painted_on: Vec<_> = pages
                    .iter()
                    .enumerate()
                    .filter(|(_, regions)| numbered_words(&regions[0]).contains(token))
                    .map(|(index, _)| index as u32)
                    .collect();
                assert_eq!(painted_on, [fragment.page_index], "{token}");
                checked += 1;
            }
        }
        for line in slice
            .lines()
            .filter(|line| line.starts_with("Tall cell line"))
        {
            assert!(
                here.contains(&format!("\n{line}\n")),
                "{line} on page {}",
                fragment.page_index
            );
            checked += 1;
        }
    }
    assert!(checked > 250, "{checked}");
}

#[test]
fn atoms_split_across_pages_are_partial_and_anchor_only() {
    let control = (0..140)
        .map(|index| format!("word{index}"))
        .collect::<Vec<_>>()
        .join(" ");
    let body = format!(
        r#"{}<w:sdt><w:sdtPr><w:tag w:val="long"/><w:id w:val="9"/><w:text/></w:sdtPr><w:sdtContent>{}</w:sdtContent></w:sdt>{}<w:sectPr><w:pgSz w:w="7200" w:h="5760"/><w:pgMar w:top="720" w:right="720" w:bottom="720" w:left="720" w:header="300" w:footer="300" w:gutter="0"/></w:sectPr>"#,
        fixture::p("00000001", &fixture::r(&"lead ".repeat(160))),
        fixture::r(&control),
        ""
    );
    let body = body.replacen("<w:sdt>", r#"<w:p w14:paraId="00000002"><w:sdt>"#, 1);
    let body = body.replacen("</w:sdt>", "</w:sdt></w:p>", 1);
    let bytes = fixture::with_body(&body);
    let (engine, _) = fixture::laid_out(&bytes, 7);
    let content = export(&engine, &PageExportOptions::new(RevisionView::Markup));
    let paragraph = body_paragraph(&content, "00000002");
    let BlockKind::Paragraph { paragraph: data } = &paragraph.content else {
        panic!("paragraph expected");
    };
    let control = data
        .inlines
        .iter()
        .find(|inline| matches!(inline.content, InlineKind::ContentControl { .. }))
        .unwrap();
    let parts = fragments_of(&content.layout, &control.id);
    assert_eq!(parts.len(), 2, "{parts:?}");
    for part in &parts {
        let FragmentSlice::Atom { range, coverage } = &part.slice else {
            panic!("atom slice expected");
        };
        assert_eq!(*coverage, AtomCoverage::Partial);
        assert_eq!(
            range.clone().map(Anchor::Range),
            Some(control.anchor.clone())
        );
    }
    assert!(parts[0].continued_on_next && parts[1].continued_from_previous);
    let anchor_only: Vec<_> = content
        .layout
        .diagnostics
        .iter()
        .filter(|diagnostic| diagnostic.code == PageDiagnosticCode::AnchorOnly)
        .map(|diagnostic| (diagnostic.node_id.clone(), diagnostic.page_index))
        .collect();
    assert_eq!(
        anchor_only,
        parts
            .iter()
            .map(|part| (Some(control.id.clone()), Some(part.page_index)))
            .collect::<Vec<_>>()
    );
}

#[test]
fn a_resident_relayout_republishes_the_layout() {
    let body = format!(
        "{}{}<w:sectPr><w:pgSz w:w=\"7200\" w:h=\"5760\"/><w:pgMar w:top=\"720\" w:right=\"720\" w:bottom=\"720\" w:left=\"720\" w:header=\"300\" w:footer=\"300\" w:gutter=\"0\"/></w:sectPr>",
        fixture::p("00000001", &fixture::r("A single section")),
        fixture::p("00000002", &fixture::r(&"filler ".repeat(200)))
    );
    let bytes = fixture::with_body(&body);
    let (engine, request) = fixture::laid_out(&bytes, 7);
    let request: serde_json::Value = serde_json::from_str(&request).unwrap();
    let extras = serde_json::json!({"fontChains": request["measurement"]["fontChains"]});
    let frame = engine.build_display_list_frame(&extras.to_string(), 0);
    assert!(frame.is_ok());
    let before = export(&engine, &PageExportOptions::new(RevisionView::Markup)).layout;
    replace(
        &engine,
        &UndoSession::default(),
        "00000001",
        "A single, edited section",
    );
    assert_eq!(
        refusal(&engine, &PageExportOptions::new(RevisionView::Markup)),
        ExportFailureCode::StaleDocument
    );
    let incremental = engine.stats().incremental_pagination_calls;
    engine.apply_and_layout("body", 1).unwrap();
    assert!(engine.stats().incremental_pagination_calls > incremental);
    let after = export(&engine, &PageExportOptions::new(RevisionView::Markup));
    assert_ne!(after.layout.layout_version, before.layout_version);
    let heading = &after.structured.stories[0].blocks[0];
    assert!(
        after
            .layout
            .fragments
            .iter()
            .any(|fragment| fragment.block_id == heading.id
                && matches!(text_range(fragment), Some((0, 24, _))))
    );
}

#[test]
fn section_metadata_from_an_older_document_is_refused() {
    let bytes = fixture::unrevised_docx();
    let (engine, request) = fixture::laid_out(&bytes, 7);
    let mut request: serde_json::Value = serde_json::from_str(&request).unwrap();
    let sections = request["regions"]["sections"].as_array_mut().unwrap();
    sections.truncate(3);
    engine
        .layout_document_with_regions_retained_json(&request.to_string())
        .unwrap();
    assert_eq!(
        refusal(&engine, &options(RevisionView::Markup)),
        ExportFailureCode::StaleLayout
    );
}

#[test]
fn even_pages_show_the_even_header_and_orphan_notes_no_page() {
    let (engine, _) = fixture::laid_out(&fixture::even_headers_docx(), 7);
    let map = export(&engine, &options(RevisionView::Accepted)).layout;
    let shown = |story: &str| {
        map.occurrences
            .iter()
            .filter(|occurrence| occurrence.story == story)
            .map(|occurrence| (occurrence.page_index, occurrence.region.clone()))
            .collect::<Vec<_>>()
    };
    let even = shown("hf:rIdHeader1");
    assert!(!even.is_empty());
    for (page, region) in &even {
        assert_eq!(page % 2, 1, "page index {page} is an even page number");
        assert_eq!(
            *region,
            OccurrenceRegion::Header {
                variant: docx_edit::structured::HeaderFooterVariant::Even
            }
        );
    }
    let default = shown("hf:rIdHeader2");
    assert!(default.iter().all(|(page, _)| page % 2 == 0));
    assert!(map.diagnostics.iter().any(|diagnostic| {
        diagnostic.code == PageDiagnosticCode::NotLaidOut && diagnostic.message.contains("fn:4")
    }));
    assert!(
        !map.occurrences
            .iter()
            .any(|occurrence| occurrence.story == "fn:4")
    );
}

#[test]
fn notes_that_never_settle_are_refused() {
    let note: String = (0..40)
        .map(|index| {
            fixture::p(
                &format!("5000{index:04X}"),
                &fixture::r(&format!("note line {index}")),
            )
        })
        .collect();
    let body = format!(
        "{}{}<w:sectPr><w:pgSz w:w=\"7200\" w:h=\"5760\"/><w:pgMar w:top=\"720\" w:right=\"720\" w:bottom=\"720\" w:left=\"720\" w:header=\"300\" w:footer=\"300\" w:gutter=\"0\"/></w:sectPr>",
        (0..3)
            .map(|index| fixture::p(
                &format!("6000{index:04X}"),
                &fixture::r(&"filler words ".repeat(20))
            ))
            .collect::<String>(),
        fixture::p(
            "00000001",
            &format!(
                r#"{}<w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteReference w:id="1"/></w:r>{}"#,
                fixture::r("Anchor"),
                fixture::r(&"tail words ".repeat(40))
            )
        )
    );
    let (engine, _) = fixture::laid_out(&fixture::with_body_and_note(&body, &note), 7);
    assert_eq!(
        refusal(&engine, &options(RevisionView::Markup)),
        ExportFailureCode::LayoutNotConverged
    );
}

#[test]
fn a_page_holding_two_sections_has_an_occurrence_for_each() {
    let geometry = r#"<w:pgSz w:w="7200" w:h="5760"/><w:pgMar w:top="720" w:right="720" w:bottom="720" w:left="720" w:header="300" w:footer="300" w:gutter="0"/>"#;
    let body = format!(
        r#"<w:p w14:paraId="00000001"><w:pPr><w:sectPr>{geometry}</w:sectPr></w:pPr>{}</w:p>{}<w:sectPr><w:type w:val="continuous"/>{geometry}</w:sectPr>"#,
        fixture::r("First section"),
        fixture::p("00000002", &fixture::r("Second section on the same page"))
    );
    let (engine, _) = fixture::laid_out(&fixture::with_body(&body), 7);
    let content = export(&engine, &PageExportOptions::new(RevisionView::Markup));
    let map = &content.layout;
    assert_eq!(map.pages.len(), 1);
    let bodies: Vec<_> = map
        .occurrences
        .iter()
        .map(|occurrence| {
            (
                occurrence.page_index,
                occurrence.section_index,
                &occurrence.region,
            )
        })
        .collect();
    assert_eq!(
        bodies,
        [
            (0, Some(0), &OccurrenceRegion::Body),
            (0, Some(1), &OccurrenceRegion::Body)
        ]
    );
    let second = body_paragraph(&content, "00000002");
    assert!(
        fragments_of(map, &second.id)
            .iter()
            .all(|fragment| fragment.occurrence_id == map.occurrences[1].id)
    );
}

#[test]
fn a_truncated_export_is_mapped_without_dangling_nodes() {
    let (engine, _) = fixture::laid_out(&fixture::unrevised_docx(), 7);
    let mut limited = options(RevisionView::Markup);
    limited.max_blocks = Some(3);
    let content = export(&engine, &limited);
    assert!(content.structured.truncated);
    let mut ids: HashSet<String> = blocks(&content).into_keys().collect();
    for block in blocks(&content).values() {
        if let BlockKind::Paragraph { paragraph } | BlockKind::Heading { paragraph, .. } =
            &block.content
        {
            ids.extend(paragraph.inlines.iter().map(|inline| inline.id.clone()));
        }
    }
    assert!(!content.layout.fragments.is_empty());
    for fragment in &content.layout.fragments {
        assert!(ids.contains(&fragment.node_id), "{}", fragment.node_id);
        assert!(ids.contains(&fragment.block_id), "{}", fragment.block_id);
    }
    for diagnostic in &content.layout.diagnostics {
        if let Some(node) = &diagnostic.node_id {
            assert!(ids.contains(node), "{node}");
        }
        assert_ne!(diagnostic.code, PageDiagnosticCode::UnmappedContent);
    }
}

/// `request` with the final section's top margin changed, as a host's page setup changes it:
/// in the final section record and in the repeat of it the request builder appends.
fn final_margin(request: &mut serde_json::Value) {
    let sections = request["regions"]["sections"].as_array_mut().unwrap();
    let last = sections.len() - 1;
    for section in &mut sections[last - 1..] {
        section["properties"]["marginTop"] = serde_json::json!(1440);
    }
}

#[test]
fn only_a_layout_of_the_sessions_own_stories_and_metadata_is_exported() {
    let bytes = fixture::unrevised_docx();
    let (engine, request) = fixture::laid_out(&bytes, 7);
    let request: serde_json::Value = serde_json::from_str(&request).unwrap();
    let relaid = |request: &serde_json::Value| {
        engine
            .layout_document_with_regions_retained_json(&request.to_string())
            .unwrap();
        engine.export_structured_with_pages(&options(RevisionView::Markup))
    };
    let mut supplied = request.clone();
    supplied.as_object_mut().unwrap().remove("bodyStory");
    assert_eq!(
        relaid(&supplied).unwrap_err().failure.code,
        ExportFailureCode::LayoutUnavailable,
        "caller-measured blocks are not the session's lowering"
    );
    relaid(&request).unwrap();

    let mut margins = request.clone();
    margins["regions"]["sections"][1]["properties"]["marginLeft"] = serde_json::json!(1440);
    let mut settings = request.clone();
    settings["regions"]["settings"]["evenAndOddHeaders"] = serde_json::json!(true);
    let mut last = request.clone();
    final_margin(&mut last);
    let mut surplus = request.clone();
    let sections = surplus["regions"]["sections"].as_array_mut().unwrap();
    sections.push(sections.last().unwrap().clone());
    let mut notes = request.clone();
    notes["notes"]["contents"]
        .as_array_mut()
        .unwrap()
        .retain(|content| content["id"] != 3);
    for (stale, what) in [
        (margins, "an inner section"),
        (settings, "the settings"),
        (last, "the final section"),
        (surplus, "a section record past the final one"),
        (notes, "a referenced note"),
    ] {
        assert_eq!(
            relaid(&stale).unwrap_err().failure.code,
            ExportFailureCode::StaleLayout,
            "{what} differs from the document"
        );
    }
}

#[test]
fn an_editor_export_needs_the_layout_of_its_current_inputs() {
    let bytes = fixture::unrevised_docx();
    let (engine, request) = fixture::laid_out(&bytes, 7);
    let current: serde_json::Value = serde_json::from_str(&request).unwrap();
    let paged = |current: &serde_json::Value| {
        engine
            .export_structured_with_pages_for(&options(RevisionView::Markup), &current.to_string())
    };
    let base = paged(&current).unwrap().content.layout;
    let mut gap = current.clone();
    gap["options"]["pageGap"] = serde_json::json!(80);
    assert_eq!(paged(&gap).unwrap().content.layout, base);
    let mut hidden = current.clone();
    hidden["renderEnv"] = serde_json::json!({"showHiddenText": true});
    let mut defaults = current.clone();
    defaults["measurement"]["defaults"]["fontSize"] = serde_json::json!(12);
    let substitute = docx_layout::register_substitute_measure_font(0, "Calibri").unwrap();
    let mut chains = current.clone();
    for chain in chains["measurement"]["fontChains"]
        .as_object_mut()
        .unwrap()
        .values_mut()
    {
        *chain = serde_json::json!([substitute, 0]);
    }
    for (changed, what) in [
        (hidden, "render environment"),
        (defaults, "defaults"),
        (chains, "chains"),
    ] {
        assert_eq!(
            paged(&changed).map(|_| ()).unwrap_err().failure.code,
            ExportFailureCode::StaleLayout,
            "{what}"
        );
    }

    let mut page_setup = current.clone();
    final_margin(&mut page_setup);
    engine
        .layout_document_with_regions_retained_json(&page_setup.to_string())
        .unwrap();
    assert_eq!(
        refusal(&engine, &options(RevisionView::Markup)),
        ExportFailureCode::StaleLayout,
        "without an editor the source package owns the final section"
    );
    paged(&page_setup).unwrap();
    assert_eq!(
        paged(&current).unwrap_err().failure.code,
        ExportFailureCode::StaleLayout
    );
}

#[test]
fn layouts_measured_without_their_fonts_are_unavailable() {
    let bytes = fixture::unrevised_docx();
    let (engine, request) = fixture::laid_out(&bytes, 7);
    let request: serde_json::Value = serde_json::from_str(&request).unwrap();
    let mut empty = request.clone();
    empty["measurement"]["fontChains"] = serde_json::json!({});
    let mut unregistered = request.clone();
    for chain in unregistered["measurement"]["fontChains"]
        .as_object_mut()
        .unwrap()
        .values_mut()
    {
        *chain = serde_json::json!([99]);
    }
    let mut blank = request.clone();
    for chain in blank["measurement"]["fontChains"]
        .as_object_mut()
        .unwrap()
        .values_mut()
    {
        *chain = serde_json::json!([]);
    }
    for missing in [&empty, &unregistered, &blank] {
        engine
            .layout_document_with_regions_retained_json(&missing.to_string())
            .unwrap();
        assert_eq!(
            refusal(&engine, &options(RevisionView::Markup)),
            ExportFailureCode::LayoutUnavailable
        );
    }

    let private = EngineSession::new(12);
    docx_edit::seed_from_docx(private.doc(), &bytes).unwrap();
    let failure = private
        .export_snapshot_with_private_fonts(&[], &blank.to_string(), &options(RevisionView::Markup))
        .unwrap()
        .unwrap_err();
    assert_eq!(failure.code, ExportFailureCode::LayoutUnavailable);
}

/// Two pages of plain text with no notes.
fn plain_body() -> String {
    format!(
        "{}{}<w:sectPr><w:pgSz w:w=\"7200\" w:h=\"5760\"/><w:pgMar w:top=\"720\" w:right=\"720\" w:bottom=\"720\" w:left=\"720\" w:header=\"300\" w:footer=\"300\" w:gutter=\"0\"/></w:sectPr>",
        fixture::p("00000001", &fixture::r(&"measured words ".repeat(120))),
        fixture::p("00000002", &fixture::r(&"more words ".repeat(120)))
    )
}

#[test]
fn a_private_export_leaves_the_shared_fonts_alone() {
    let bytes = fixture::with_body(&plain_body());
    let (engine, request) = fixture::laid_out(&bytes, 7);
    let store = docx_layout::measure_store_id();
    let live = export(&engine, &options(RevisionView::Markup)).layout;

    let mut snapshots = Vec::new();
    for font in [fixture::OTHER_FONT, fixture::FONT] {
        let private = EngineSession::new(12);
        docx_edit::seed_from_docx(private.doc(), &bytes).unwrap();
        let private_request = fixture::region_request(&private, &bytes, 0).to_string();
        let paged = private
            .export_snapshot_with_private_fonts(
                &[font],
                &private_request,
                &options(RevisionView::Markup),
            )
            .unwrap()
            .unwrap();
        snapshots.push(paged.layout);
        assert_eq!(docx_layout::measure_store_id(), store);
        let after = export(&engine, &options(RevisionView::Markup)).layout;
        assert_eq!(
            after, live,
            "the live layout is still current and unchanged"
        );
    }
    assert_ne!(
        snapshots[0].provenance.font_set_fingerprint, snapshots[1].provenance.font_set_fingerprint,
        "the two private exports measured with their own fonts"
    );
    assert_eq!(
        snapshots[1].provenance.font_set_fingerprint,
        live.provenance.font_set_fingerprint
    );
    assert_ne!(snapshots[0].fragments, snapshots[1].fragments);
    engine
        .layout_document_with_regions_retained_json(&request)
        .unwrap();
    let relaid = export(&engine, &options(RevisionView::Markup)).layout;
    assert_eq!(relaid.pages, live.pages);
    assert_eq!(relaid.fragments, live.fragments);
    assert_eq!(
        relaid.provenance.font_set_fingerprint, live.provenance.font_set_fingerprint,
        "a relayout measures with the live fonts again"
    );
}

#[test]
fn merged_header_aliases_keep_their_cells_and_controls() {
    let (engine, _) = fixture::laid_out(&fixture::aliased_headers_docx(), 7);
    let content = export(&engine, &options(RevisionView::Markup));
    let map = &content.layout;
    let headers: Vec<_> = content
        .structured
        .stories
        .iter()
        .filter(|story| story.kind == docx_edit::structured::StoryKind::Header)
        .collect();
    assert_eq!(headers.len(), 1, "the two headers export as one story");
    let story = &headers[0].story;
    let exported = blocks(&content);
    let node = |text: &str| {
        exported
            .values()
            .find(|block| {
                matches!(&block.anchor, Anchor::Paragraph { story: owner, .. } if owner != "body")
                    && matches!(&block.content, BlockKind::Paragraph { paragraph } if paragraph.inlines.iter().any(|inline| matches!(&inline.content, InlineKind::Text { text: shown } if shown == text)))
            })
            .unwrap()
            .id
            .clone()
    };
    let control = headers[0]
        .blocks
        .iter()
        .find(|block| matches!(block.content, BlockKind::ContentControl { .. }))
        .unwrap()
        .id
        .clone();
    let occurrences: Vec<_> = map
        .occurrences
        .iter()
        .filter(|occurrence| &occurrence.story == story)
        .collect();
    assert!(occurrences.len() > 2);
    for occurrence in occurrences {
        for id in [
            node("Shared header"),
            node("Header cell"),
            node("Header control"),
            control.clone(),
        ] {
            assert!(
                map.fragments
                    .iter()
                    .any(|fragment| fragment.occurrence_id == occurrence.id
                        && fragment.node_id == id),
                "{id} on {}",
                occurrence.id
            );
        }
    }
    assert!(
        !map.diagnostics
            .iter()
            .any(|diagnostic| diagnostic.code == PageDiagnosticCode::UnmappedContent),
        "{:?}",
        map.diagnostics
    );
}

#[test]
fn a_part_shared_by_two_roles_is_shown_in_the_role_each_page_selects() {
    let (engine, _) = fixture::laid_out(&fixture::shared_header_docx(), 7);
    let map = export(&engine, &options(RevisionView::Markup)).layout;
    let roles: Vec<_> = map
        .occurrences
        .iter()
        .filter(|occurrence| occurrence.story == "hf:rIdHeader2" && occurrence.page_index < 2)
        .map(|occurrence| (occurrence.page_index, occurrence.region.clone()))
        .collect();
    assert_eq!(
        roles,
        [
            (
                0,
                OccurrenceRegion::Header {
                    variant: docx_edit::structured::HeaderFooterVariant::First
                }
            ),
            (
                1,
                OccurrenceRegion::Header {
                    variant: docx_edit::structured::HeaderFooterVariant::Default
                }
            ),
        ]
    );
}

#[test]
fn endnotes_follow_the_placement_the_settings_choose() {
    let (engine, _) = fixture::laid_out(&fixture::section_endnotes_docx(), 7);
    let map = export(&engine, &options(RevisionView::Markup)).layout;
    let endnote = map
        .occurrences
        .iter()
        .find(|occurrence| occurrence.story == "en:1")
        .unwrap();
    assert_eq!(
        endnote.region,
        OccurrenceRegion::Endnote {
            note_id: "1".to_owned(),
            placement: NotePlacement::SectionEnd
        }
    );
    assert!(
        map.fragments
            .iter()
            .any(|fragment| fragment.occurrence_id == endnote.id)
    );
}

#[test]
fn a_note_too_tall_for_its_page_is_diagnosed_not_placed() {
    let note: String = (0..40)
        .map(|index| {
            fixture::p(
                &format!("5000{index:04X}"),
                &fixture::r(&format!("note line {index}")),
            )
        })
        .collect();
    let body = format!(
        "{}{}<w:sectPr><w:pgSz w:w=\"7200\" w:h=\"5760\"/><w:pgMar w:top=\"720\" w:right=\"720\" w:bottom=\"720\" w:left=\"720\" w:header=\"300\" w:footer=\"300\" w:gutter=\"0\"/></w:sectPr>",
        fixture::p(
            "00000001",
            &format!(
                r#"<w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteReference w:id="1"/></w:r>{}"#,
                fixture::r("Anchor")
            )
        ),
        fixture::p("00000002", &fixture::r(&"tail words ".repeat(40)))
    );
    let (engine, _) = fixture::laid_out(&fixture::with_body_and_note(&body, &note), 7);
    let map = export(&engine, &options(RevisionView::Markup)).layout;
    let occurrence = map
        .occurrences
        .iter()
        .find(|occurrence| occurrence.story == "fn:1")
        .unwrap();
    assert!(
        !map.fragments
            .iter()
            .any(|fragment| fragment.occurrence_id == occurrence.id),
        "no content is claimed for a note its page cannot hold"
    );
    assert!(map.diagnostics.iter().any(|diagnostic| {
        diagnostic.code == PageDiagnosticCode::UnsupportedNoteLayout
            && diagnostic.page_index == Some(occurrence.page_index)
    }));
}

#[test]
fn a_relayout_after_the_fonts_change_measures_again() {
    let bytes = fixture::with_body(&plain_body());
    let (engine, request) = fixture::laid_out(&bytes, 7);
    let before = export(&engine, &options(RevisionView::Markup)).layout;
    docx_layout::clear_measure_fonts();
    assert_eq!(
        docx_layout::register_measure_font(fixture::OTHER_FONT).unwrap(),
        0,
        "the other font takes the old font's id"
    );
    assert_eq!(
        refusal(&engine, &options(RevisionView::Markup)),
        ExportFailureCode::StaleLayout
    );
    engine
        .layout_document_with_regions_retained_json(&request)
        .unwrap();
    let relaid = export(&engine, &options(RevisionView::Markup)).layout;
    let (fresh, _) = fixture::laid_out_with(&bytes, 8, fixture::OTHER_FONT);
    let baseline = export(&fresh, &options(RevisionView::Markup)).layout;
    assert_ne!(baseline.fragments, before.fragments);
    assert_eq!(relaid.pages, baseline.pages);
    assert_eq!(relaid.fragments, baseline.fragments);
    assert_eq!(
        relaid.provenance.font_set_fingerprint,
        baseline.provenance.font_set_fingerprint
    );
    docx_layout::register_measure_font(fixture::FONT).unwrap();
    assert_eq!(
        refusal(&fresh, &options(RevisionView::Markup)),
        ExportFailureCode::StaleLayout,
        "a font registered after the layout could have measured it"
    );
}

#[test]
fn text_in_the_default_family_needs_that_family_measured() {
    let bytes = fixture::without_default_fonts(&plain_body());
    docx_layout::clear_measure_fonts();
    let font = docx_layout::register_measure_font(fixture::FONT).unwrap();
    let engine = EngineSession::new(7);
    docx_edit::seed_from_docx(engine.doc(), &bytes).unwrap();
    let mut request = fixture::region_request(&engine, &bytes, font);
    request["measurement"]["defaults"]["fontFamily"] = serde_json::json!("Liberation Sans");
    let requirements: Vec<serde_json::Value> = serde_json::from_str(
        &engine
            .layout_font_requirements_json(&request.to_string())
            .unwrap(),
    )
    .unwrap();
    let keys: Vec<_> = requirements
        .iter()
        .map(|requirement| requirement["key"].as_str().unwrap().to_owned())
        .collect();
    assert!(keys.contains(&"liberation sans|0|0".to_owned()), "{keys:?}");
    assert!(
        !keys.iter().any(|key| key.starts_with("calibri|")),
        "{keys:?}"
    );
    let mut chained = request.clone();
    chained["measurement"]["fontChains"] = keys
        .iter()
        .map(|key| (key.clone(), serde_json::json!([font])))
        .collect::<serde_json::Map<_, _>>()
        .into();
    let laid = |request: &serde_json::Value| {
        engine
            .layout_document_with_regions_retained_json(&request.to_string())
            .unwrap();
        engine.export_structured_with_pages(&options(RevisionView::Markup))
    };
    laid(&chained).unwrap();
    assert_eq!(
        laid(&request).unwrap_err().failure.code,
        ExportFailureCode::LayoutUnavailable,
        "chains resolved for Calibri leave the default family unmeasured"
    );
    let mut undefined = chained.clone();
    undefined["measurement"]["defaults"] = serde_json::Value::Null;
    undefined["measurement"]["fontChains"]["calibri|0|0"] = serde_json::json!([font]);
    assert_eq!(
        laid(&undefined).unwrap_err().failure.code,
        ExportFailureCode::LayoutUnavailable,
        "text measured with stand-in metrics is not exported"
    );
}

#[test]
fn text_a_header_table_clips_has_no_page_fragments() {
    let (engine, _) = fixture::laid_out(&fixture::clipped_header_docx(), 7);
    let content = export(&engine, &options(RevisionView::Markup));
    let map = &content.layout;
    let exported = blocks(&content);
    let line = |text: &str| {
        exported
            .values()
            .find(|block| {
                matches!(&block.content, BlockKind::Paragraph { paragraph } if paragraph.inlines.iter().any(|inline| matches!(&inline.content, InlineKind::Text { text: shown } if shown == text)))
            })
            .unwrap()
            .id
            .clone()
    };
    let occurrences: Vec<_> = map
        .occurrences
        .iter()
        .filter(|occurrence| occurrence.story == "hf:rIdHeader2")
        .collect();
    assert_eq!(occurrences.len(), 2);
    let (shown, cut, clipped) = (
        line("Clipped line 0"),
        line("Clipped line 1"),
        line("Clipped line 3"),
    );
    for occurrence in occurrences {
        let on = |node: &str| {
            map.fragments
                .iter()
                .any(|fragment| fragment.occurrence_id == occurrence.id && fragment.node_id == node)
        };
        assert!(on(&shown), "the row shows its first line");
        assert!(on(&cut), "the row shows the top of its second line");
        assert!(!on(&clipped), "the row clips its last line");
        assert!(map.diagnostics.iter().any(|diagnostic| {
            diagnostic.code == PageDiagnosticCode::ClippedContent
                && diagnostic.node_id.as_deref() == Some(cut.as_str())
                && diagnostic.page_index == Some(occurrence.page_index)
        }));
    }
    assert!(map.diagnostics.iter().any(|diagnostic| {
        diagnostic.code == PageDiagnosticCode::NotLaidOut
            && diagnostic.node_id.as_deref() == Some(clipped.as_str())
    }));
}

#[test]
fn a_body_relayout_after_a_header_edit_lays_the_header_out_again() {
    let bytes = fixture::headed_docx(&fixture::p("22000001", &fixture::r("Running header")));
    let edit = |engine: &EngineSession| {
        let history = UndoSession::default();
        replace_in(
            engine,
            &history,
            "hf:rIdHeader2",
            "22000001",
            &"A running header grown to two lines ".repeat(3),
        );
        let words: Vec<_> = (0..120).map(|index| format!("filler{index}")).collect();
        replace(
            engine,
            &history,
            "00000002",
            &format!("edited {}", words.join(" ")),
        );
    };
    let laid = |client| {
        let (engine, request) = fixture::laid_out(&bytes, client);
        let mut request: serde_json::Value = serde_json::from_str(&request).unwrap();
        request["notes"]["contents"] = serde_json::json!([]);
        request["regions"]["sections"]
            .as_array_mut()
            .unwrap()
            .truncate(1);
        let request = request.to_string();
        engine
            .layout_document_with_regions_retained_json(&request)
            .unwrap();
        (engine, request)
    };
    let (engine, _) = laid(7);
    let extras = serde_json::json!({"fontChains": {}});
    engine
        .build_display_list_frame(&extras.to_string(), 0)
        .unwrap();
    let before = export(&engine, &options(RevisionView::Markup)).layout;
    edit(&engine);
    engine.apply_and_layout("body", 1).unwrap();
    let after = export(&engine, &options(RevisionView::Markup)).layout;

    let (fresh, request) = laid(8);
    edit(&fresh);
    fresh
        .layout_document_with_regions_retained_json(&request)
        .unwrap();
    let baseline = export(&fresh, &options(RevisionView::Markup)).layout;
    let body_pages = |map: &DocxLayoutMap| {
        map.fragments
            .iter()
            .filter(|fragment| fragment.occurrence_id.contains(".body."))
            .map(|fragment| {
                (
                    fragment.node_id.clone(),
                    fragment.page_index,
                    fragment.slice.clone(),
                )
            })
            .collect::<Vec<_>>()
    };
    assert_eq!(baseline.pages.len(), before.pages.len());
    assert_ne!(
        body_pages(&baseline),
        body_pages(&before),
        "the grown header moves body text across pages"
    );
    assert_eq!(after.pages, baseline.pages);
    assert_eq!(after.fragments, baseline.fragments);
}

#[test]
fn a_layout_with_a_section_the_document_no_longer_has_is_refused() {
    let (engine, request) = fixture::laid_out(&fixture::without_third_section_docx(), 7);
    let current = export(&engine, &options(RevisionView::Markup)).layout;
    let stale = fixture::region_request(&engine, &fixture::unrevised_docx(), 0);
    engine
        .layout_document_with_regions_retained_json(&stale.to_string())
        .unwrap();
    assert_eq!(
        refusal(&engine, &options(RevisionView::Markup)),
        ExportFailureCode::StaleLayout
    );
    engine
        .layout_document_with_regions_retained_json(&request)
        .unwrap();
    let relaid = export(&engine, &options(RevisionView::Markup)).layout;
    assert_eq!(relaid.pages, current.pages);
    let sections: BTreeSet<_> = relaid.pages.iter().map(|page| page.section_index).collect();
    assert_eq!(sections.into_iter().collect::<Vec<_>>(), [0, 1, 2]);
    assert!(
        relaid
            .pages
            .iter()
            .all(|page| page.numbering_format != "upperLetter"),
        "the lettered section is gone"
    );
    assert!(relaid.occurrences.iter().any(|occurrence| {
        occurrence.story == "hf:rIdHeader2" && occurrence.section_index == Some(2)
    }));
}

#[test]
fn a_limited_map_stops_mapping_past_its_limits() {
    let control = (0..140)
        .map(|index| format!("word{index}"))
        .collect::<Vec<_>>()
        .join(" ");
    let body = format!(
        r#"{}{}<w:p w14:paraId="00000002"><w:sdt><w:sdtPr><w:tag w:val="long"/><w:id w:val="9"/><w:text/></w:sdtPr><w:sdtContent>{}</w:sdtContent></w:sdt></w:p><w:sectPr><w:pgSz w:w="7200" w:h="5760"/><w:pgMar w:top="720" w:right="720" w:bottom="720" w:left="720" w:header="300" w:footer="300" w:gutter="0"/></w:sectPr>"#,
        (0..12)
            .map(|index| fixture::p(
                &format!("6000{index:04X}"),
                &fixture::r(&"filler words ".repeat(60))
            ))
            .collect::<String>(),
        fixture::p("00000001", &fixture::r(&"lead ".repeat(160))),
        fixture::r(&control),
    );
    let (engine, _) = fixture::laid_out(&fixture::with_body(&body), 7);
    let full = export(&engine, &PageExportOptions::new(RevisionView::Markup)).layout;
    let partial = |map: &DocxLayoutMap| {
        map.diagnostics
            .iter()
            .filter(|diagnostic| diagnostic.code == PageDiagnosticCode::AnchorOnly)
            .count()
    };
    assert_eq!(
        partial(&full),
        2,
        "the control splits across two late pages"
    );
    assert!(full.pages.len() > 4);
    let mut limited = PageExportOptions::new(RevisionView::Markup);
    limited.max_fragments = Some(1);
    let map = export(&engine, &limited).layout;
    assert!(map.truncated);
    assert_eq!(map.pages, full.pages);
    assert_eq!(map.fragments, full.fragments[..1]);
    assert_eq!(
        partial(&map),
        0,
        "the pages past the limit are never mapped"
    );
}
