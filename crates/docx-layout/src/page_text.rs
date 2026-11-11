//! The text a laid-out page shows, read from its fragments without building the page.

use std::borrow::Cow;

use serde::Serialize;

use crate::display_list::{
    resident_block_id_key, resident_fragment_block_key, source_blocks_by_key,
};
use crate::types::{
    BlockExtent, Fragment, Input, Layout, LayoutBlock, Page, ParagraphBlock, Run, TableBlock,
    TableFragment, TypesetRow,
};

/// One paragraph of a page's text: what the page shows of it, in reading order.
#[derive(Debug, Clone, PartialEq, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct PageTextParagraph {
    pub text: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub para_id: Option<String>,
    /// The paragraph style as written.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub style_id: Option<String>,
    /// Set for a paragraph inside a table cell, whose text reads apart from the body's.
    #[serde(skip_serializing_if = "std::ops::Not::not")]
    pub in_table: bool,
}

/// The text of one page of a layout.
#[derive(Debug, Clone, PartialEq, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct PageText {
    pub page_index: usize,
    pub paragraphs: Vec<PageTextParagraph>,
}

/// What `PAGE` and `NUMPAGES` fields show on the page being read.
#[derive(Clone, Copy)]
struct PageFields<'a> {
    label: Option<&'a str>,
    number: u64,
    /// The document's page count; zero while the layout is partial.
    total: u64,
    cached_totals: bool,
}

fn run_text<'a>(run: &'a Run, page: PageFields<'_>) -> Option<Cow<'a, str>> {
    match run {
        Run::Text(run) => Some(Cow::Borrowed(run.text.as_str())),
        Run::Field(run) => match run.field_type.as_str() {
            "PAGE" => Some(Cow::Owned(
                crate::regions::page_field_text(page.label, page.number).into_owned(),
            )),
            "NUMPAGES" if page.total == 0 => run
                .fallback
                .as_deref()
                .filter(|_| page.cached_totals)
                .map(Cow::Borrowed),
            "NUMPAGES" => Some(Cow::Owned(page.total.to_string())),
            _ => run.fallback.as_deref().map(Cow::Borrowed),
        },
        Run::Tab(_) => Some(Cow::Borrowed("\t")),
        Run::LineBreak(_) => Some(Cow::Borrowed("\n")),
        _ => None,
    }
}

/// The part of `text` starting `from` and ending `to` UTF-16 units into it, clamped.
fn utf16_slice(text: &str, from: usize, to: usize) -> String {
    let units: Vec<u16> = text.encode_utf16().collect();
    let start = from.min(units.len());
    let end = to.max(start).min(units.len());
    String::from_utf16_lossy(&units[start..end])
}

/// The text `lines` of `paragraph` show, each run cut as the display cuts it.
fn lines_text(paragraph: &ParagraphBlock, lines: &[TypesetRow], page: PageFields<'_>) -> String {
    let mut text = String::new();
    for line in lines {
        for run_index in line.head_run..=line.tail_run {
            let Some(run_text) = paragraph
                .runs
                .get(run_index)
                .and_then(|run| run_text(run, page))
            else {
                continue;
            };
            if !matches!(paragraph.runs[run_index], Run::Text(_)) {
                text.push_str(&run_text);
                continue;
            }
            let total = run_text.encode_utf16().count();
            let start = if run_index == line.head_run {
                line.head_char.min(total)
            } else {
                0
            };
            let end = if run_index == line.tail_run {
                line.tail_char.min(total)
            } else {
                total
            };
            text.push_str(&utf16_slice(&run_text, start, end.max(start)));
        }
    }
    text
}

/// The whole text of `paragraph`.
fn paragraph_text(paragraph: &ParagraphBlock, page: PageFields<'_>) -> String {
    paragraph
        .runs
        .iter()
        .filter_map(|run| run_text(run, page))
        .collect()
}

fn push(
    paragraphs: &mut Vec<PageTextParagraph>,
    text: String,
    paragraph: &ParagraphBlock,
    in_table: bool,
) {
    if text.chars().any(|c| !c.is_whitespace()) {
        paragraphs.push(PageTextParagraph {
            text,
            para_id: paragraph.para_id.clone(),
            style_id: paragraph
                .attrs
                .as_ref()
                .and_then(|attrs| attrs.style_id.clone()),
            in_table,
        });
    }
}

fn table_text(
    table: &TableBlock,
    rows: std::ops::Range<usize>,
    page: PageFields<'_>,
    paragraphs: &mut Vec<PageTextParagraph>,
) {
    let end = rows.end.min(table.rows.len());
    for row in &table.rows[rows.start.min(end)..end] {
        for cell in &row.cells {
            for nested in &cell.blocks {
                match nested {
                    LayoutBlock::Paragraph(paragraph) => {
                        push(paragraphs, paragraph_text(paragraph, page), paragraph, true)
                    }
                    LayoutBlock::Table(table) => {
                        table_text(table, 0..table.rows.len(), page, paragraphs);
                    }
                    _ => {}
                }
            }
        }
    }
}

/// The first row `fragment` reads: a row split across pages or columns reads
/// whole where it starts, so a continuation skips the rows that an earlier
/// fragment of the table, on this page or the one before, already placed.
fn first_unread_row(layout: &Layout, page_index: usize, fragment: &TableFragment) -> usize {
    let key = resident_block_id_key(&fragment.block_id);
    let earlier_on_page = layout.pages[page_index].fragments.iter().take_while(
        |candidate| !matches!(candidate, Fragment::Table(table) if std::ptr::eq(table, fragment)),
    );
    let previous_page = page_index
        .checked_sub(1)
        .and_then(|index| layout.pages.get(index))
        .into_iter()
        .flat_map(|page| &page.fragments);
    let read_to = previous_page
        .chain(earlier_on_page)
        .filter_map(|candidate| match candidate {
            Fragment::Table(table) if resident_block_id_key(&table.block_id) == key => {
                Some(table.row_end)
            }
            _ => None,
        })
        .max();
    match read_to {
        Some(row_end) if row_end > fragment.row_start => row_end.min(fragment.row_end),
        _ => fragment.row_start,
    }
}

fn page_fields<'a>(page: &'a Page, layout: &Layout) -> PageFields<'a> {
    PageFields {
        label: page.page_label.as_deref(),
        number: u64::from(page.number),
        total: if layout.partial {
            0
        } else {
            layout.pages.len() as u64
        },
        cached_totals: layout.cached_page_totals,
    }
}

/// The text each of `pages` shows: body paragraphs, with a paragraph split
/// across pages cut where its fragment is, and the rows of tables the page
/// places, cell by cell. A row split across pages or columns reads where it
/// starts, and header rows a table repeats on later pages read only once.
/// `PAGE` and `NUMPAGES` fields read as the page shows them. Text boxes,
/// shapes, images, charts, headers, footers and notes are not read. A page
/// index past the layout is skipped.
pub fn layout_page_texts(pagination: &Input, layout: &Layout, pages: &[usize]) -> Vec<PageText> {
    let selected: Vec<usize> = pages
        .iter()
        .copied()
        .filter(|&index| index < layout.pages.len())
        .collect();
    let blocks = source_blocks_by_key(
        pagination,
        selected.iter().map(|&index| &layout.pages[index]),
    );
    selected
        .into_iter()
        .map(|page_index| {
            let page = &layout.pages[page_index];
            let fields = page_fields(page, layout);
            let mut paragraphs = Vec::new();
            for fragment in &page.fragments {
                let Some(measured) = blocks.get(resident_fragment_block_key(fragment).as_ref())
                else {
                    continue;
                };
                match (fragment, &measured.block) {
                    (Fragment::Paragraph(fragment), LayoutBlock::Paragraph(paragraph)) => {
                        let text = match &measured.measure {
                            BlockExtent::Paragraph(extent) if !extent.lines.is_empty() => {
                                let end = fragment.to_line.min(extent.lines.len());
                                let start = fragment.from_line.min(end);
                                lines_text(paragraph, &extent.lines[start..end], fields)
                            }
                            _ if fragment.from_line > 0 => String::new(),
                            _ => paragraph_text(paragraph, fields),
                        };
                        push(&mut paragraphs, text, paragraph, false);
                    }
                    (Fragment::Table(fragment), LayoutBlock::Table(table)) => {
                        let start = first_unread_row(layout, page_index, fragment);
                        table_text(table, start..fragment.row_end, fields, &mut paragraphs);
                    }
                    _ => {}
                }
            }
            PageText {
                page_index,
                paragraphs,
            }
        })
        .collect()
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::types::FieldRun;

    fn field(field_type: &str, fallback: Option<&str>) -> Run {
        Run::Field(FieldRun {
            fmt: Default::default(),
            field_type: field_type.to_owned(),
            raw_type: None,
            instruction: None,
            fallback: fallback.map(str::to_owned),
            locked: false,
            nested_sequences: Vec::new(),
            pm_start: None,
            pm_end: None,
        })
    }

    fn shown(run: &Run, total: u64, cached_totals: bool) -> Option<String> {
        let page = PageFields {
            label: None,
            number: 2,
            total,
            cached_totals,
        };
        run_text(run, page).map(Cow::into_owned)
    }

    #[test]
    fn page_count_fields_read_as_the_display_paints_them() {
        let pages = field("NUMPAGES", Some("9"));
        assert_eq!(shown(&pages, 3, false).as_deref(), Some("3"));
        assert_eq!(shown(&pages, 3, true).as_deref(), Some("3"));
        assert_eq!(shown(&pages, 0, true).as_deref(), Some("9"));
        assert_eq!(shown(&pages, 0, false), None);
    }

    #[test]
    fn page_number_fields_read_the_page_and_others_their_cached_result() {
        assert_eq!(
            shown(&field("PAGE", Some("1")), 3, false).as_deref(),
            Some("2")
        );
        assert_eq!(
            shown(&field("DATE", Some("Today")), 3, false).as_deref(),
            Some("Today")
        );
    }

    fn row(head_run: usize, head_char: usize, tail_run: usize, tail_char: usize) -> TypesetRow {
        serde_json::from_value(serde_json::json!({
            "headRun": head_run, "headChar": head_char, "tailRun": tail_run, "tailChar": tail_char,
            "width": 1.0, "ascent": 1.0, "descent": 1.0, "lineHeight": 1.0,
        }))
        .unwrap()
    }

    #[test]
    fn each_fragment_reads_only_its_lines_without_body_positions() {
        let paragraph: ParagraphBlock = serde_json::from_value(serde_json::json!({
            "id": "p",
            "runs": [{ "kind": "text", "text": "Ab" }, { "kind": "text", "text": "cd" }],
        }))
        .unwrap();
        let page = PageFields {
            label: None,
            number: 1,
            total: 2,
            cached_totals: false,
        };
        assert_eq!(lines_text(&paragraph, &[row(0, 0, 0, 1)], page), "A");
        assert_eq!(lines_text(&paragraph, &[row(0, 1, 1, 2)], page), "bcd");
        assert_eq!(paragraph_text(&paragraph, page), "Abcd");
    }

    #[test]
    fn a_cut_inside_a_surrogate_pair_keeps_the_units_either_side() {
        assert_eq!(utf16_slice("a😀b", 3, 4), "b");
        assert_eq!(utf16_slice("a😀b", 0, 9), "a😀b");
    }
}
