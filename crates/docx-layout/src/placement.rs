//! Where laid-out content lands. Every physical page is read as its regions — the body flow
//! (split per source section), the header and footer bands the page shows, and the notes its
//! note areas hold — and each region lists the paragraph line windows, table row windows and
//! atomic blocks placed there. What a table fragment shows inside its cells comes from the
//! display list's own paint plan, and band selection from its band resolver, so a placement
//! never disagrees with the page it describes.

use std::collections::HashMap;
use std::ops::Range;

use crate::display_list::{
    ShownPart, TableBlockIn, TableExtentIn, TableFragmentIn, display_mirror, visible_table_content,
};
use crate::footnotes::{NoteContent, NoteKind, note_reference_map_id};
use crate::header_footer::{HeaderFooterKind, HeaderFooterPayload, HeaderFooterType};
use crate::hf_bands::{BandPage, BandSettings, select_band_variant};
use crate::types::{
    BlockExtent, Fragment, HeaderFooterRefs, Layout, LayoutBlock, MeasuredBlock, Page,
    ParagraphBlock, ParagraphExtent, Run, TableBlock, TableExtent, TableFragment,
};

/// How far a note area may reach into the body above it before it counts as overflowing, in
/// CSS pixels.
const NOTE_AREA_TOLERANCE: f64 = 0.5;

/// The region of a page content was placed in.
#[derive(Clone, Debug, PartialEq)]
pub enum PlacedRegion<'a> {
    /// The body flow of one source section.
    Body { section_index: usize },
    /// The header or footer part a page shows, selected for `section_index`.
    Band {
        kind: HeaderFooterKind,
        hf_type: HeaderFooterType,
        r_id: &'a str,
        section_index: usize,
    },
    /// One note in a note area.
    Note {
        kind: NoteKind,
        id: i64,
        placement: &'a str,
    },
}

/// One page's regions in reading order: body sections, header, footer, then notes.
#[derive(Debug)]
pub struct PlacedPage<'a> {
    pub page_index: usize,
    pub page: &'a Page,
    pub regions: Vec<PlacedRegionContent<'a>>,
}

#[derive(Debug)]
pub struct PlacedRegionContent<'a> {
    pub region: PlacedRegion<'a>,
    pub items: Vec<PlacedItem<'a>>,
}

/// A piece of one block placed in a region. Tables are listed before the content of their
/// cells, which follows as items of its own.
#[derive(Clone, Debug)]
pub enum PlacedItem<'a> {
    Paragraph(PlacedParagraph<'a>),
    Table(PlacedTable<'a>),
    /// An image, shape, chart or text box placed whole.
    Block(&'a LayoutBlock),
}

/// The line window `[lines.start, lines.end)` of a paragraph shown in one region.
#[derive(Clone, Debug)]
pub struct PlacedParagraph<'a> {
    pub block: &'a ParagraphBlock,
    pub measure: &'a ParagraphExtent,
    pub lines: Range<usize>,
    pub continued_from_previous: bool,
    pub continued_on_next: bool,
    /// Painted again as part of a repeated table header row.
    pub repeated_header: bool,
    /// Inside a table row whose earlier part is on another page.
    pub in_row_continuation: bool,
    /// Some line of the window is only partly inside its table cell's clip box.
    pub clipped: bool,
    /// The page fragment's box `[x, y, width, height]`, for body flow.
    pub frame: Option<[f64; 4]>,
}

/// The rows of a table shown in one region.
#[derive(Clone, Debug)]
pub struct PlacedTable<'a> {
    pub block: &'a TableBlock,
    pub rows: Vec<PlacedRow>,
    pub continued_from_previous: bool,
    pub continued_on_next: bool,
    pub repeated_header: bool,
    /// The page fragment's box `[x, y, width, height]`, for body flow.
    pub frame: Option<[f64; 4]>,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct PlacedRow {
    pub row_index: usize,
    pub continued_from_previous: bool,
    pub continued_on_next: bool,
    pub repeated_header: bool,
}

/// A page fragment or note area whose content could not be located.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct PlacementIssue {
    pub page_index: usize,
    pub kind: PlacementIssueKind,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub enum PlacementIssueKind {
    /// A body fragment names no block of the measured arena.
    UnresolvedFragment,
    /// A note area names a note with no laid-out content.
    MissingNote { kind: NoteKind, id: i64 },
    /// The page shows a header or footer part whose content was not measured.
    UnmeasuredBand,
    /// A note area does not fit between the body and the bottom margin; its notes would have
    /// to continue on the next page, which the layout does not do, so they are not placed.
    NoteOverflow { kind: NoteKind, id: i64 },
    /// A footnote's reference is shown on another page than the footnote: the layout pairs a
    /// footnote with the first page of the table row holding its reference.
    NoteReferenceElsewhere { id: i64, reference_page: usize },
}

/// The pages of a layout with everything placed on them.
#[derive(Debug, Default)]
pub struct Placements<'a> {
    pub pages: Vec<PlacedPage<'a>>,
    pub issues: Vec<PlacementIssue>,
}

/// Inputs of one completed region layout.
pub struct PlacementInput<'a> {
    pub layout: &'a Layout,
    pub measured: &'a [MeasuredBlock],
    pub headers_footers: Option<&'a HeaderFooterPayload>,
    /// Whether the layout composed header and footer bands at all.
    pub bands_composed: bool,
    pub notes: &'a [NoteContent],
}

/// Reads every page of a region layout. Linear in the placed content.
pub fn place_layout<'a>(input: &PlacementInput<'a>) -> Placements<'a> {
    let arena = Arena::new(input.measured);
    let notes: HashMap<i64, &NoteContent> = input
        .notes
        .iter()
        .map(|content| (content.map_id(), content))
        .collect();
    let mut tables: HashMap<usize, TableGeometry<'a>> = HashMap::new();
    let mut bands: HashMap<usize, (Vec<PlacedItem<'a>>, bool)> = HashMap::new();
    let mut output = Placements::default();
    // The first page showing each footnote reference inside a table, and each footnote's page.
    let mut cell_references: HashMap<i64, usize> = HashMap::new();
    let mut footnote_pages: HashMap<i64, usize> = HashMap::new();
    for (page_index, page) in input.layout.pages.iter().enumerate() {
        let mut body: Vec<(usize, Vec<PlacedItem<'a>>)> = Vec::new();
        for fragment in &page.fragments {
            let Some(index) = arena.find(fragment) else {
                output.issues.push(PlacementIssue {
                    page_index,
                    kind: PlacementIssueKind::UnresolvedFragment,
                });
                continue;
            };
            let section_index = arena.sections[index];
            let items = match body.last_mut() {
                Some((section, items)) if *section == section_index => items,
                _ => {
                    body.push((section_index, Vec::new()));
                    &mut body.last_mut().expect("section pushed").1
                }
            };
            let measured = &input.measured[index];
            match (fragment, &measured.block, &measured.measure) {
                (
                    Fragment::Paragraph(fragment),
                    LayoutBlock::Paragraph(block),
                    BlockExtent::Paragraph(measure),
                ) => items.push(PlacedItem::Paragraph(PlacedParagraph {
                    block,
                    measure,
                    lines: fragment.from_line..fragment.to_line.min(measure.lines.len()),
                    continued_from_previous: fragment.carried_from_prev == Some(true),
                    continued_on_next: fragment.carried_to_next == Some(true),
                    repeated_header: false,
                    in_row_continuation: false,
                    clipped: false,
                    frame: Some([fragment.x, fragment.y, fragment.width, fragment.height]),
                })),
                (
                    Fragment::Table(fragment),
                    LayoutBlock::Table(block),
                    BlockExtent::Table(measure),
                ) => {
                    let geometry = tables
                        .entry(index)
                        .or_insert_with(|| TableGeometry::new(block, measure));
                    let start = items.len();
                    if !place_table_fragment(geometry, fragment, items) {
                        output.issues.push(PlacementIssue {
                            page_index,
                            kind: PlacementIssueKind::UnresolvedFragment,
                        });
                    }
                    for item in &items[start..] {
                        if let PlacedItem::Paragraph(paragraph) = item
                            && !paragraph.repeated_header
                        {
                            for id in footnote_references(paragraph) {
                                cell_references.entry(id).or_insert(page_index);
                            }
                        }
                    }
                }
                (_, block, _) => items.push(PlacedItem::Block(block)),
            }
        }
        let body_bottom = page
            .fragments
            .iter()
            .filter_map(|fragment| match fragment {
                Fragment::Paragraph(fragment) => Some(fragment.y + fragment.height),
                Fragment::Table(fragment) => Some(fragment.y + fragment.height),
                _ => None,
            })
            .fold(page.margins.top, f64::max);
        let mut regions: Vec<PlacedRegionContent<'a>> = if body.is_empty() {
            vec![PlacedRegionContent {
                region: PlacedRegion::Body {
                    section_index: page.region_section_index,
                },
                items: Vec::new(),
            }]
        } else {
            body.into_iter()
                .map(|(section_index, items)| PlacedRegionContent {
                    region: PlacedRegion::Body { section_index },
                    items,
                })
                .collect()
        };
        if page.parity_filler != Some(true) {
            for kind in [HeaderFooterKind::Header, HeaderFooterKind::Footer] {
                match input.headers_footers {
                    Some(payload) => {
                        if let Some((band, hf_type)) =
                            band_for_page(payload, page, page_index, kind)
                        {
                            let (items, readable) = bands
                                .entry(std::ptr::from_ref(band) as usize)
                                .or_insert_with(|| {
                                    let mut items = Vec::new();
                                    let readable = whole_items(
                                        band.measured
                                            .iter()
                                            .map(|measured| (&measured.block, &measured.measure)),
                                        &mut items,
                                    );
                                    (items, readable)
                                })
                                .clone();
                            if !readable {
                                output.issues.push(PlacementIssue {
                                    page_index,
                                    kind: PlacementIssueKind::UnresolvedFragment,
                                });
                            }
                            regions.push(PlacedRegionContent {
                                region: PlacedRegion::Band {
                                    kind,
                                    hf_type,
                                    r_id: &band.r_id,
                                    section_index: band.section_index,
                                },
                                items,
                            });
                        }
                    }
                    None if input.bands_composed && page_refs(page, kind) => {
                        output.issues.push(PlacementIssue {
                            page_index,
                            kind: PlacementIssueKind::UnmeasuredBand,
                        });
                    }
                    None => {}
                }
            }
        }
        for area in page.note_areas.iter().flatten() {
            let kind = match area.kind.as_deref() {
                Some("endnote") => NoteKind::Endnote,
                _ => NoteKind::Footnote,
            };
            let placement = area.placement.as_deref().unwrap_or(match kind {
                NoteKind::Footnote => "pageBottom",
                NoteKind::Endnote => "docEnd",
            });
            let overflows = area
                .y
                .is_some_and(|y| y + NOTE_AREA_TOLERANCE < body_bottom);
            for note in area.notes.iter().flatten() {
                let Some(id) = note.id else {
                    continue;
                };
                if kind == NoteKind::Footnote {
                    footnote_pages.entry(id).or_insert(page_index);
                }
                let Some(content) = notes.get(&note_reference_map_id(id, kind)) else {
                    output.issues.push(PlacementIssue {
                        page_index,
                        kind: PlacementIssueKind::MissingNote { kind, id },
                    });
                    continue;
                };
                let mut items = Vec::new();
                if overflows {
                    output.issues.push(PlacementIssue {
                        page_index,
                        kind: PlacementIssueKind::NoteOverflow { kind, id },
                    });
                } else if !whole_items(content.blocks.iter().zip(&content.measures), &mut items) {
                    output.issues.push(PlacementIssue {
                        page_index,
                        kind: PlacementIssueKind::UnresolvedFragment,
                    });
                }
                regions.push(PlacedRegionContent {
                    region: PlacedRegion::Note {
                        kind,
                        id,
                        placement,
                    },
                    items,
                });
            }
        }
        output.pages.push(PlacedPage {
            page_index,
            page,
            regions,
        });
    }
    let mut elsewhere: Vec<(i64, usize, usize)> = cell_references
        .into_iter()
        .filter_map(|(id, reference_page)| {
            let note_page = *footnote_pages.get(&id)?;
            (note_page != reference_page).then_some((id, reference_page, note_page))
        })
        .collect();
    elsewhere.sort_unstable();
    for (id, reference_page, note_page) in elsewhere {
        output.issues.push(PlacementIssue {
            page_index: note_page,
            kind: PlacementIssueKind::NoteReferenceElsewhere { id, reference_page },
        });
    }
    output
}

/// The ids of the footnotes a placed paragraph shows references to.
fn footnote_references(paragraph: &PlacedParagraph<'_>) -> Vec<i64> {
    line_window_slices(paragraph.block, paragraph.measure, paragraph.lines.clone())
        .into_iter()
        .filter_map(|slice| match paragraph.block.runs.get(slice.run_index) {
            Some(Run::Text(text)) => text.fmt.footnote_ref_id.map(|id| id as i64),
            _ => None,
        })
        .collect()
}

fn page_refs(page: &Page, kind: HeaderFooterKind) -> bool {
    page.header_footer_refs
        .as_ref()
        .is_some_and(|refs| match kind {
            HeaderFooterKind::Header => {
                refs.header_default.is_some()
                    || refs.header_first.is_some()
                    || refs.header_even.is_some()
            }
            HeaderFooterKind::Footer => {
                refs.footer_default.is_some()
                    || refs.footer_first.is_some()
                    || refs.footer_even.is_some()
            }
        })
}

fn refs_for(
    refs: &HeaderFooterRefs,
    kind: HeaderFooterKind,
    hf_type: HeaderFooterType,
) -> Option<&str> {
    match (kind, hf_type) {
        (HeaderFooterKind::Header, HeaderFooterType::Default) => refs.header_default.as_deref(),
        (HeaderFooterKind::Header, HeaderFooterType::First) => refs.header_first.as_deref(),
        (HeaderFooterKind::Header, HeaderFooterType::Even) => refs.header_even.as_deref(),
        (HeaderFooterKind::Footer, HeaderFooterType::Default) => refs.footer_default.as_deref(),
        (HeaderFooterKind::Footer, HeaderFooterType::First) => refs.footer_first.as_deref(),
        (HeaderFooterKind::Footer, HeaderFooterType::Even) => refs.footer_even.as_deref(),
    }
}

/// The header or footer variant the display list paints on `page`, with the role it shows it in.
fn band_for_page<'a>(
    payload: &'a HeaderFooterPayload,
    page: &'a Page,
    page_index: usize,
    kind: HeaderFooterKind,
) -> Option<(
    &'a crate::header_footer::HeaderFooterVariant,
    HeaderFooterType,
)> {
    select_band_variant(
        &payload.variants,
        |variant| {
            (
                variant.kind,
                variant.hf_type,
                variant.r_id.as_str(),
                Some(variant.section_index),
            )
        },
        &BandSettings {
            title_pg: payload.title_pg,
            even_and_odd_headers: payload.even_and_odd_headers,
            title_page_sections: &payload.title_page_sections,
            even_and_odd_sections: &payload.even_and_odd_sections,
        },
        &BandPage {
            section_index: page.section_index.map(|value| value as usize),
            section_page_index: page.section_page_index,
            page_number: if page.number > 0 {
                u64::from(page.number)
            } else {
                page_index as u64 + 1
            },
            has_refs: page.header_footer_refs.is_some(),
        },
        |kind, hf_type| {
            page.header_footer_refs
                .as_ref()
                .and_then(|refs| refs_for(refs, kind, hf_type))
        },
        kind,
    )
}

/// Every block of a header, footer or note story, which is placed whole: paragraphs with all
/// their lines, and tables with all rows and what their cells show by the painting rules. False
/// when a table cannot be read.
fn whole_items<'a>(
    blocks: impl Iterator<Item = (&'a LayoutBlock, &'a BlockExtent)>,
    items: &mut Vec<PlacedItem<'a>>,
) -> bool {
    let mut readable = true;
    for (block, measure) in blocks {
        match (block, measure) {
            (LayoutBlock::Paragraph(block), BlockExtent::Paragraph(measure)) => {
                items.push(PlacedItem::Paragraph(PlacedParagraph {
                    block,
                    measure,
                    lines: 0..measure.lines.len(),
                    continued_from_previous: false,
                    continued_on_next: false,
                    repeated_header: false,
                    in_row_continuation: false,
                    clipped: false,
                    frame: None,
                }));
            }
            (LayoutBlock::Table(table), BlockExtent::Table(extent)) => {
                let whole = TableFragment {
                    block_id: table.id.clone(),
                    x: 0.0,
                    y: 0.0,
                    width: extent.total_width,
                    pm_start: None,
                    pm_end: None,
                    row_start: 0,
                    row_end: table.rows.len(),
                    height: extent.total_height,
                    is_floating: None,
                    carried_from_prev: None,
                    carried_to_next: None,
                    header_row_count: None,
                    clip_top: None,
                    clip_bottom: None,
                };
                let at = items.len();
                readable &= place_table_fragment(&TableGeometry::new(table, extent), &whole, items);
                if let Some(PlacedItem::Table(placed)) = items.get_mut(at) {
                    placed.frame = None;
                }
            }
            (
                LayoutBlock::SectionBreak(_)
                | LayoutBlock::PageBreak(_)
                | LayoutBlock::ColumnBreak(_),
                _,
            )
            | (LayoutBlock::Unsupported, _) => {}
            (block, _) => items.push(PlacedItem::Block(block)),
        }
    }
    readable
}

/// A table and its display mirror, converted once however many pages it spans.
struct TableGeometry<'a> {
    block: &'a TableBlock,
    measure: &'a TableExtent,
    mirror: Option<(TableBlockIn, TableExtentIn)>,
}

impl<'a> TableGeometry<'a> {
    fn new(block: &'a TableBlock, measure: &'a TableExtent) -> Self {
        Self {
            block,
            measure,
            mirror: display_mirror(block).zip(display_mirror(measure)),
        }
    }
}

/// The rows a table fragment shows and the content its cells show, as the display list paints
/// them. Repeated header rows come first. False when the fragment cannot be read.
fn place_table_fragment<'a>(
    geometry: &TableGeometry<'a>,
    fragment: &TableFragment,
    items: &mut Vec<PlacedItem<'a>>,
) -> bool {
    let rows = geometry.measure.rows.len().min(geometry.block.rows.len());
    let row_start = fragment.row_start.min(rows);
    let row_end = fragment.row_end.min(rows).max(row_start);
    let carried = fragment.carried_from_prev == Some(true);
    let header_rows = if carried {
        (fragment.header_row_count.unwrap_or(0.0).max(0.0) as usize).min(row_start)
    } else {
        0
    };
    let clip_top = fragment.clip_top.unwrap_or(0.0);
    let mut placed_rows: Vec<PlacedRow> = (0..header_rows)
        .map(|row_index| PlacedRow {
            row_index,
            continued_from_previous: false,
            continued_on_next: false,
            repeated_header: true,
        })
        .collect();
    for row_index in row_start..row_end {
        placed_rows.push(PlacedRow {
            row_index,
            continued_from_previous: row_index == row_start && clip_top > 0.0,
            continued_on_next: row_index + 1 == row_end && fragment.clip_bottom.is_some(),
            repeated_header: false,
        });
    }
    items.push(PlacedItem::Table(PlacedTable {
        block: geometry.block,
        rows: placed_rows,
        continued_from_previous: carried,
        continued_on_next: fragment.carried_to_next == Some(true),
        repeated_header: false,
        frame: Some([fragment.x, fragment.y, fragment.width, fragment.height]),
    }));
    let (Some((block, measure)), Some(fragment)) = (
        geometry.mirror.as_ref(),
        display_mirror::<TableFragmentIn>(fragment),
    ) else {
        return false;
    };
    for visible in visible_table_content(&fragment, block, measure) {
        let Some((block, measure)) = cell_block(geometry.block, geometry.measure, &visible.path)
        else {
            continue;
        };
        match (visible.shown, block, measure) {
            (
                ShownPart::Lines(lines, clipped),
                LayoutBlock::Paragraph(block),
                BlockExtent::Paragraph(measure),
            ) => {
                items.push(PlacedItem::Paragraph(PlacedParagraph {
                    block,
                    measure,
                    continued_from_previous: lines.start > 0,
                    continued_on_next: lines.end < measure.lines.len(),
                    lines,
                    repeated_header: visible.repeated_header,
                    in_row_continuation: visible.continuation,
                    clipped,
                    frame: None,
                }));
            }
            (ShownPart::Rows(rows), LayoutBlock::Table(table), BlockExtent::Table(extent)) => {
                let count = table.rows.len().min(extent.rows.len());
                items.push(PlacedItem::Table(PlacedTable {
                    block: table,
                    continued_from_previous: visible.continuation
                        || rows.first().is_some_and(|row| *row > 0),
                    continued_on_next: rows.last().is_some_and(|row| row + 1 < count),
                    rows: rows
                        .into_iter()
                        .filter(|row| *row < count)
                        .map(|row_index| PlacedRow {
                            row_index,
                            continued_from_previous: false,
                            continued_on_next: false,
                            repeated_header: visible.repeated_header,
                        })
                        .collect(),
                    repeated_header: visible.repeated_header,
                    frame: None,
                }));
            }
            (ShownPart::Whole, block, _) => items.push(PlacedItem::Block(block)),
            _ => {}
        }
    }
    true
}

/// The block a cell path names, with its measure.
fn cell_block<'a>(
    mut table: &'a TableBlock,
    mut extent: &'a TableExtent,
    path: &[(usize, usize, usize)],
) -> Option<(&'a LayoutBlock, &'a BlockExtent)> {
    let (last, steps) = path.split_last()?;
    for &(row, cell, index) in steps {
        match (
            table.rows.get(row)?.cells.get(cell)?.blocks.get(index)?,
            extent.rows.get(row)?.cells.get(cell)?.blocks.get(index)?,
        ) {
            (LayoutBlock::Table(nested), BlockExtent::Table(nested_extent)) => {
                table = nested;
                extent = nested_extent;
            }
            _ => return None,
        }
    }
    let &(row, cell, index) = last;
    Some((
        table.rows.get(row)?.cells.get(cell)?.blocks.get(index)?,
        extent.rows.get(row)?.cells.get(cell)?.blocks.get(index)?,
    ))
}

/// The measured body blocks a page fragment can name, found by kind and document position.
struct Arena<'a> {
    /// Paragraph blocks by start position; their ranges are disjoint.
    paragraphs: Vec<(f64, f64, usize)>,
    /// Other placeable blocks by kind and start position.
    positioned: HashMap<(u8, u64), usize>,
    /// The source section of every arena block.
    sections: Vec<usize>,
    measured: &'a [MeasuredBlock],
}

fn kind_tag(block: &LayoutBlock) -> Option<u8> {
    match block {
        LayoutBlock::Table(_) => Some(1),
        LayoutBlock::Image(_) => Some(2),
        LayoutBlock::Shape(_) => Some(3),
        LayoutBlock::Chart(_) => Some(4),
        LayoutBlock::TextBox(_) => Some(5),
        _ => None,
    }
}

fn fragment_tag(fragment: &Fragment) -> Option<u8> {
    match fragment {
        Fragment::Paragraph(_) => None,
        Fragment::Table(_) => Some(1),
        Fragment::Image(_) => Some(2),
        Fragment::Shape(_) => Some(3),
        Fragment::Chart(_) => Some(4),
        Fragment::TextBox(_) => Some(5),
    }
}

fn fragment_pm_start(fragment: &Fragment) -> Option<f64> {
    match fragment {
        Fragment::Paragraph(value) => value.pm_start,
        Fragment::Table(value) => value.pm_start,
        Fragment::Image(value) => value.pm_start,
        Fragment::Shape(value) => value.pm_start,
        Fragment::Chart(value) => value.pm_start,
        Fragment::TextBox(value) => value.pm_start,
    }
}

impl<'a> Arena<'a> {
    fn new(measured: &'a [MeasuredBlock]) -> Self {
        let mut paragraphs = Vec::new();
        let mut positioned = HashMap::new();
        let mut sections = Vec::with_capacity(measured.len());
        let mut section = 0;
        for (index, entry) in measured.iter().enumerate() {
            sections.push(section);
            match &entry.block {
                LayoutBlock::Paragraph(block) => {
                    if let (Some(start), Some(end)) = (block.pm_start, block.pm_end) {
                        paragraphs.push((start, end, index));
                    }
                }
                LayoutBlock::SectionBreak(_) => section += 1,
                block => {
                    if let (Some(tag), Some(start)) = (kind_tag(block), block.pm_start()) {
                        positioned.insert((tag, start.to_bits()), index);
                    }
                }
            }
        }
        paragraphs.sort_by(|left, right| left.0.total_cmp(&right.0));
        Self {
            paragraphs,
            positioned,
            sections,
            measured,
        }
    }

    fn find(&self, fragment: &Fragment) -> Option<usize> {
        let start = fragment_pm_start(fragment)?;
        let index = match fragment_tag(fragment) {
            None => {
                let after = self.paragraphs.partition_point(|entry| entry.0 <= start);
                let &(block_start, block_end, index) =
                    self.paragraphs.get(after.checked_sub(1)?)?;
                (start >= block_start && start <= block_end).then_some(index)?
            }
            Some(tag) => *self.positioned.get(&(tag, start.to_bits()))?,
        };
        let expected = match fragment {
            Fragment::Paragraph(value) => &value.block_id,
            Fragment::Table(value) => &value.block_id,
            Fragment::Image(value) => &value.block_id,
            Fragment::Shape(value) => &value.block_id,
            Fragment::Chart(value) => &value.block_id,
            Fragment::TextBox(value) => &value.block_id,
        };
        (self.measured[index].block.block_id() == Some(expected)).then_some(index)
    }
}

/// One run's placed part: its index and the document positions it covers.
#[derive(Clone, Copy, Debug, PartialEq)]
pub struct RunSlice {
    pub run_index: usize,
    pub pm_start: f64,
    pub pm_end: f64,
}

/// The run parts lines `[lines.start, lines.end)` of a paragraph show, in logical order. Text is
/// sliced by UTF-16 offsets, which follow its positions; a one-position text run (a note mark or
/// an equation stand-in, whatever its label) and every other run count whole. A soft break
/// belongs to the line it ends. Linear in the runs the lines show, never in their text.
pub fn line_window_slices(
    block: &ParagraphBlock,
    measure: &ParagraphExtent,
    lines: Range<usize>,
) -> Vec<RunSlice> {
    let mut slices: Vec<RunSlice> = Vec::new();
    for line in measure.lines.get(lines).into_iter().flatten() {
        for run_index in line.head_run..=line.tail_run {
            let Some(run) = block.runs.get(run_index) else {
                continue;
            };
            let (Some(pm_start), Some(pm_end)) = (run.pm_start(), run.pm_end()) else {
                continue;
            };
            let tail = run_index == line.tail_run;
            let head = run_index == line.head_run;
            let slice = match run {
                Run::Text(text) => {
                    let positions = (pm_end - pm_start).max(0.0) as usize;
                    let units = if positions <= 1 {
                        crate::resolve_lines::utf16_len(&text.text)
                    } else {
                        positions
                    };
                    let start = if head { line.head_char.min(units) } else { 0 };
                    let end = if tail {
                        line.tail_char.min(units)
                    } else {
                        units
                    };
                    if end <= start {
                        continue;
                    }
                    if positions <= 1 {
                        (pm_start, pm_end)
                    } else {
                        (pm_start + start as f64, pm_start + end as f64)
                    }
                }
                Run::LineBreak(_) => (pm_start, pm_end),
                _ if tail && line.tail_char == 0 => continue,
                _ => (pm_start, pm_end),
            };
            match slices.last_mut() {
                Some(previous) if previous.run_index == run_index && previous.pm_end == slice.0 => {
                    previous.pm_end = slice.1;
                }
                _ => slices.push(RunSlice {
                    run_index,
                    pm_start: slice.0,
                    pm_end: slice.1,
                }),
            }
        }
    }
    slices
}

#[cfg(test)]
mod tests {
    use serde_json::json;

    use super::*;

    fn line(head: (usize, usize), tail: (usize, usize)) -> serde_json::Value {
        json!({
            "headRun": head.0, "headChar": head.1, "tailRun": tail.0, "tailChar": tail.1,
            "width": 0.0, "ascent": 0.0, "descent": 0.0, "lineHeight": 10.0
        })
    }

    /// A paragraph starting at `start` with `lines` lines of 10px, one character each.
    fn paragraph(start: u64, lines: usize) -> (serde_json::Value, serde_json::Value) {
        let end = start + 1 + lines as u64;
        (
            json!({
                "kind": "paragraph", "id": format!("p{start}"),
                "runs": [{"kind": "text", "text": "a".repeat(lines), "pmStart": start + 1, "pmEnd": end}],
                "pmStart": start, "pmEnd": end + 1
            }),
            json!({
                "kind": "paragraph",
                "lines": (0..lines).map(|index| line((0, index), (0, index + 1))).collect::<Vec<_>>(),
                "totalHeight": lines as f64 * 10.0
            }),
        )
    }

    /// A table of 100px columns: per row its height and cells, each with a row span and blocks.
    fn table(
        id: &str,
        rows: Vec<(f64, Vec<(u32, Vec<(serde_json::Value, serde_json::Value)>)>)>,
    ) -> (serde_json::Value, serde_json::Value) {
        let columns = rows.iter().map(|(_, cells)| cells.len()).max().unwrap_or(0);
        let block = json!({
            "kind": "table", "id": id,
            "rows": rows.iter().enumerate().map(|(row, (_, cells))| json!({
                "id": format!("{id}r{row}"),
                "cells": cells.iter().enumerate().map(|(cell, (span, blocks))| json!({
                    "id": format!("{id}r{row}c{cell}"),
                    "rowSpan": span,
                    "blocks": blocks.iter().map(|(block, _)| block.clone()).collect::<Vec<_>>()
                })).collect::<Vec<_>>()
            })).collect::<Vec<_>>()
        });
        let extent = json!({
            "kind": "table",
            "rows": rows.iter().map(|(height, cells)| json!({
                "height": height,
                "cells": cells.iter().map(|(_, blocks)| json!({
                    "width": 100.0,
                    "height": height,
                    "blocks": blocks.iter().map(|(_, extent)| extent.clone()).collect::<Vec<_>>()
                })).collect::<Vec<_>>()
            })).collect::<Vec<_>>(),
            "columnWidths": vec![100.0; columns],
            "totalWidth": 100.0 * columns as f64,
            "totalHeight": rows.iter().map(|(height, _)| height).sum::<f64>()
        });
        (block, extent)
    }

    /// The paragraphs, by start position, and nested table rows one table fragment shows.
    fn shown(
        table: &(serde_json::Value, serde_json::Value),
        fragment: TableFragment,
    ) -> Vec<(i64, Range<usize>, bool)> {
        let block: TableBlock = serde_json::from_value(table.0.clone()).unwrap();
        let extent: TableExtent = serde_json::from_value(table.1.clone()).unwrap();
        let mut items = Vec::new();
        assert!(place_table_fragment(
            &TableGeometry::new(&block, &extent),
            &fragment,
            &mut items
        ));
        items
            .iter()
            .skip(1)
            .map(|item| match item {
                PlacedItem::Paragraph(paragraph) => (
                    paragraph.block.pm_start.unwrap() as i64,
                    paragraph.lines.clone(),
                    paragraph.in_row_continuation,
                ),
                PlacedItem::Table(table) => (
                    -1,
                    table.rows.first().map_or(0, |row| row.row_index)
                        ..table.rows.last().map_or(0, |row| row.row_index + 1),
                    table.continued_from_previous,
                ),
                PlacedItem::Block(_) => (-2, 0..0, false),
            })
            .collect()
    }

    /// A page slice of table `t`: its rows, top and height, and whether it is continued from
    /// the previous page or on the next, with the row cuts `(clip_top, clip_bottom)`.
    fn fragment(
        rows: Range<usize>,
        y: f64,
        height: f64,
        (from, next): (bool, bool),
        (clip_top, clip_bottom): (Option<f64>, Option<f64>),
    ) -> TableFragment {
        TableFragment {
            block_id: crate::types::BlockId::Str("t".to_owned()),
            x: 0.0,
            y,
            width: 200.0,
            pm_start: None,
            pm_end: None,
            row_start: rows.start,
            row_end: rows.end,
            height,
            is_floating: None,
            carried_from_prev: from.then_some(true),
            carried_to_next: next.then_some(true),
            header_row_count: None,
            clip_top,
            clip_bottom,
        }
    }

    #[test]
    fn a_vertically_merged_cell_shows_the_lines_each_page_paints() {
        let merged = table(
            "t",
            vec![
                (
                    20.0,
                    vec![(2, vec![paragraph(0, 4)]), (1, vec![paragraph(100, 1)])],
                ),
                (20.0, vec![(1, vec![paragraph(200, 1)])]),
            ],
        );

        assert_eq!(
            shown(
                &merged,
                fragment(0..1, 100.0, 20.0, (false, true), (None, None))
            ),
            [(0, 0..2, false), (100, 0..1, false)]
        );
        assert_eq!(
            shown(
                &merged,
                fragment(1..2, 50.0, 20.0, (true, false), (None, None))
            ),
            [(0, 2..4, true), (200, 0..1, false)],
            "the merged cell's later lines are painted, and exported, on the next page"
        );
    }

    #[test]
    fn a_line_its_cell_cuts_through_is_shown_and_marked_clipped() {
        let (block, extent) = table("t", vec![(25.0, vec![(1, vec![paragraph(0, 4)])])]);
        let block: TableBlock = serde_json::from_value(block).unwrap();
        let extent: TableExtent = serde_json::from_value(extent).unwrap();
        let mut items = Vec::new();
        assert!(place_table_fragment(
            &TableGeometry::new(&block, &extent),
            &fragment(0..1, 100.0, 25.0, (false, false), (None, None)),
            &mut items
        ));
        let Some(PlacedItem::Paragraph(paragraph)) = items.get(1) else {
            panic!("a paragraph is placed");
        };
        assert_eq!(
            paragraph.lines,
            0..3,
            "a line cut in half still shows its top"
        );
        assert!(paragraph.clipped);
    }

    #[test]
    fn content_clipped_by_its_cell_is_not_shown() {
        let clipped = table("t", vec![(20.0, vec![(1, vec![paragraph(0, 4)])])]);

        assert_eq!(
            shown(
                &clipped,
                fragment(0..1, 100.0, 20.0, (false, false), (None, None))
            ),
            [(0, 0..2, false)]
        );
    }

    #[test]
    fn a_nested_table_in_a_split_row_shows_its_rows_on_each_page() {
        let nested = table(
            "n",
            vec![
                (10.0, vec![(1, vec![paragraph(10, 1)])]),
                (10.0, vec![(1, vec![paragraph(20, 1)])]),
            ],
        );
        let outer = table("t", vec![(20.0, vec![(1, vec![nested])])]);

        assert_eq!(
            shown(
                &outer,
                fragment(0..1, 100.0, 10.0, (false, true), (None, Some(10.0)))
            ),
            [(-1, 0..1, false), (10, 0..1, false)]
        );
        assert_eq!(
            shown(
                &outer,
                fragment(0..1, 50.0, 10.0, (true, false), (Some(10.0), None))
            ),
            [(-1, 1..2, true), (20, 0..1, true)]
        );
    }

    #[test]
    fn line_windows_slice_text_by_utf16_and_keep_atoms_whole() {
        let block: ParagraphBlock = serde_json::from_value(json!({
            "id": "p",
            "runs": [
                {"kind": "text", "text": "hello \u{1F600} world", "pmStart": 1, "pmEnd": 15},
                {"kind": "lineBreak", "pmStart": 15, "pmEnd": 16},
                {"kind": "text", "text": "12", "pmStart": 16, "pmEnd": 17},
                {"kind": "tab", "pmStart": 17, "pmEnd": 18},
                {"kind": "text", "text": "end", "pmStart": 18, "pmEnd": 21}
            ],
            "pmStart": 0,
            "pmEnd": 22
        }))
        .unwrap();
        let measure: ParagraphExtent = serde_json::from_value(json!({
            "lines": [line((0, 0), (0, 8)), line((0, 8), (1, 0)), line((2, 0), (4, 3))],
            "totalHeight": 30.0
        }))
        .unwrap();
        let slices = |lines: Range<usize>| {
            line_window_slices(&block, &measure, lines)
                .into_iter()
                .map(|slice| (slice.run_index, slice.pm_start, slice.pm_end))
                .collect::<Vec<_>>()
        };
        assert_eq!(slices(0..1), [(0, 1.0, 9.0)]);
        assert_eq!(slices(1..2), [(0, 9.0, 15.0), (1, 15.0, 16.0)]);
        assert_eq!(
            slices(2..3),
            [(2, 16.0, 17.0), (3, 17.0, 18.0), (4, 18.0, 21.0)],
            "a note mark's label spans its one position"
        );
        assert_eq!(
            slices(0..2),
            [(0, 1.0, 15.0), (1, 15.0, 16.0)],
            "a run continuing onto the next line of the window stays one slice"
        );
    }
}
