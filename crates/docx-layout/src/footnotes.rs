use std::collections::HashMap;

use serde::{Deserialize, Serialize};

use crate::LayoutError;
use crate::regions::{DocumentRegions, format_number};
use crate::types::{
    BlockExtent, BlockId, Fragment, Layout, LayoutBlock, NoteAreaContract, NoteLayoutItemContract,
    Page, ParagraphBlock, Run, RunFormatting, TextRun,
};

pub const FOOTNOTE_SEPARATOR_HEIGHT: f64 = 12.0;
pub const FOOTNOTE_COLUMN_GAP_PX: f64 = 24.0;
pub const MAX_FOOTNOTE_LAYOUT_PASSES: usize = 6;
const FOOTNOTE_FONT_SIZE_PT: f64 = 8.0;

#[derive(Clone, Debug, PartialEq)]
pub struct OrderedMap<K, V> {
    entries: Vec<(K, V)>,
}

impl<K: PartialEq, V> OrderedMap<K, V> {
    pub fn new() -> Self {
        Self {
            entries: Vec::new(),
        }
    }

    pub fn is_empty(&self) -> bool {
        self.entries.is_empty()
    }

    pub fn get(&self, key: &K) -> Option<&V> {
        self.entries
            .iter()
            .find(|(candidate, _)| candidate == key)
            .map(|(_, value)| value)
    }

    pub fn get_mut(&mut self, key: &K) -> Option<&mut V> {
        self.entries
            .iter_mut()
            .find(|(candidate, _)| candidate == key)
            .map(|(_, value)| value)
    }

    pub fn set(&mut self, key: K, value: V) {
        if let Some(current) = self.get_mut(&key) {
            *current = value;
        } else {
            self.entries.push((key, value));
        }
    }

    pub fn iter(&self) -> std::slice::Iter<'_, (K, V)> {
        self.entries.iter()
    }
}

impl<K: PartialEq, V> Default for OrderedMap<K, V> {
    fn default() -> Self {
        Self::new()
    }
}

impl<K: PartialEq, V> FromIterator<(K, V)> for OrderedMap<K, V> {
    fn from_iter<I: IntoIterator<Item = (K, V)>>(iter: I) -> Self {
        let mut map = Self::new();
        for (key, value) in iter {
            map.set(key, value);
        }
        map
    }
}

#[derive(Clone, Copy, Debug, Default, Eq, Hash, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum NoteKind {
    #[default]
    Footnote,
    Endnote,
}

impl NoteKind {
    fn as_str(self) -> &'static str {
        match self {
            Self::Footnote => "footnote",
            Self::Endnote => "endnote",
        }
    }
}

#[derive(Clone, Debug, Default, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct NoteAnchor {
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub doc_start: Option<f64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub doc_end: Option<f64>,
}

#[derive(Clone, Debug, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct NoteContent {
    pub id: i64,
    #[serde(default)]
    pub note_kind: NoteKind,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub display_number: Option<i64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub display_label: Option<String>,
    #[serde(default)]
    pub blocks: Vec<LayoutBlock>,
    #[serde(default)]
    pub measures: Vec<BlockExtent>,
    pub height: f64,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub anchor: Option<NoteAnchor>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub custom_mark_follows: Option<bool>,
}

impl NoteContent {
    pub fn map_id(&self) -> i64 {
        note_reference_map_id(self.id, self.note_kind)
    }
}

#[derive(Clone, Debug, Default, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct NoteLayoutInput {
    #[serde(default)]
    pub contents: Vec<NoteContent>,
}

#[derive(Clone, Debug, PartialEq)]
pub struct NoteRefLocation {
    pub note_id: i64,
    pub note_kind: NoteKind,
    pub pm_pos: f64,
    pub table_block_id: Option<BlockId>,
    pub row_index: Option<usize>,
}

impl NoteRefLocation {
    pub fn map_id(&self) -> i64 {
        note_reference_map_id(self.note_id, self.note_kind)
    }
}

#[derive(Clone, Debug, PartialEq)]
pub struct NotePresentation {
    pub display_number: i64,
    pub display_label: String,
    pub anchor: NoteAnchor,
}

pub struct StabilizedNoteLayout {
    pub layout: Layout,
    pub page_note_map: OrderedMap<u32, Vec<i64>>,
    pub reserved_heights: OrderedMap<u32, f64>,
    pub converged: bool,
}

pub fn note_reference_map_id(note_id: i64, kind: NoteKind) -> i64 {
    match kind {
        NoteKind::Footnote => note_id,
        NoteKind::Endnote => -note_id.abs() - 1,
    }
}

pub fn collect_note_refs(blocks: &[LayoutBlock]) -> Vec<NoteRefLocation> {
    struct TableContext<'a> {
        id: &'a BlockId,
        row_index: usize,
    }

    fn collect_paragraph(
        paragraph: &ParagraphBlock,
        table: Option<&TableContext<'_>>,
        refs: &mut Vec<NoteRefLocation>,
    ) {
        for run in &paragraph.runs {
            let Run::Text(text) = run else {
                continue;
            };
            let (note_id, note_kind) = if let Some(id) = text.fmt.footnote_ref_id {
                (id as i64, NoteKind::Footnote)
            } else if let Some(id) = text.fmt.endnote_ref_id {
                (id as i64, NoteKind::Endnote)
            } else {
                continue;
            };
            refs.push(NoteRefLocation {
                note_id,
                note_kind,
                pm_pos: text.pm_start.unwrap_or(0.0),
                table_block_id: table.map(|context| context.id.clone()),
                row_index: table.map(|context| context.row_index),
            });
        }
    }

    fn walk(
        blocks: &[LayoutBlock],
        table: Option<&TableContext<'_>>,
        refs: &mut Vec<NoteRefLocation>,
    ) {
        for block in blocks {
            match block {
                LayoutBlock::Paragraph(paragraph) => collect_paragraph(paragraph, table, refs),
                LayoutBlock::Table(table_block) => {
                    for (row_index, row) in table_block.rows.iter().enumerate() {
                        let own_context;
                        let context = if let Some(table) = table {
                            table
                        } else {
                            own_context = TableContext {
                                id: &table_block.id,
                                row_index,
                            };
                            &own_context
                        };
                        for cell in &row.cells {
                            walk(&cell.blocks, Some(context), refs);
                        }
                    }
                }
                LayoutBlock::TextBox(text_box) => {
                    for paragraph in &text_box.content {
                        collect_paragraph(paragraph, table, refs);
                    }
                }
                _ => {}
            }
        }
    }

    let mut refs = Vec::new();
    walk(blocks, None, &mut refs);
    refs
}

fn fragment_pm_range(fragment: &Fragment) -> (Option<f64>, Option<f64>) {
    match fragment {
        Fragment::Paragraph(value) => (value.pm_start, value.pm_end),
        Fragment::Table(value) => (value.pm_start, value.pm_end),
        Fragment::Image(value) => (value.pm_start, value.pm_end),
        Fragment::Shape(value) => (value.pm_start, value.pm_end),
        Fragment::Chart(value) => (value.pm_start, value.pm_end),
        Fragment::TextBox(value) => (value.pm_start, value.pm_end),
    }
}

fn fragment_matches_ref(fragment: &Fragment, reference: &NoteRefLocation) -> bool {
    if let (Some(table_id), Some(row_index)) = (&reference.table_block_id, reference.row_index) {
        return matches!(
            fragment,
            Fragment::Table(table)
                if &table.block_id == table_id
                    && row_index >= table.row_start
                    && row_index < table.row_end
        );
    }
    let (start, end) = fragment_pm_range(fragment);
    start.is_some_and(|start| start >= 0.0 && reference.pm_pos >= start)
        && end.is_some_and(|end| end >= 0.0 && reference.pm_pos < end)
}

pub fn map_note_anchors_to_pages(
    pages: &[Page],
    refs: &[NoteRefLocation],
) -> OrderedMap<u32, Vec<i64>> {
    let mut result = OrderedMap::new();
    for reference in refs {
        if let Some(page) = pages.iter().find(|page| {
            page.fragments
                .iter()
                .any(|fragment| fragment_matches_ref(fragment, reference))
        }) {
            append_unique(&mut result, page.number, reference.map_id());
        }
    }
    result
}

fn append_unique(map: &mut OrderedMap<u32, Vec<i64>>, page_number: u32, map_id: i64) {
    if let Some(ids) = map.get_mut(&page_number) {
        if !ids.contains(&map_id) {
            ids.push(map_id);
        }
    } else {
        map.set(page_number, vec![map_id]);
    }
}

fn page_for_map_id<'a>(
    pages: &'a [Page],
    anchors: &OrderedMap<u32, Vec<i64>>,
    map_id: i64,
) -> Option<&'a Page> {
    let page_number = anchors
        .iter()
        .find(|(_, ids)| ids.contains(&map_id))
        .map(|(page_number, _)| *page_number)?;
    pages.iter().find(|page| page.number == page_number)
}

pub fn map_notes_to_pages(
    pages: &[Page],
    refs: &[NoteRefLocation],
    regions: &DocumentRegions,
) -> OrderedMap<u32, Vec<i64>> {
    let anchors = map_note_anchors_to_pages(pages, refs);
    let mut result = OrderedMap::new();
    for reference in refs {
        let map_id = reference.map_id();
        let Some(anchor_page) = page_for_map_id(pages, &anchors, map_id) else {
            continue;
        };
        let properties = regions.note_properties(
            anchor_page.region_section_index,
            reference.note_kind.as_str(),
        );
        let default_position = match reference.note_kind {
            NoteKind::Footnote => "pageBottom",
            NoteKind::Endnote => "docEnd",
        };
        let target = match properties.position.as_deref().unwrap_or(default_position) {
            "docEnd" => pages.last().unwrap_or(anchor_page),
            "sectEnd" => pages
                .iter()
                .rev()
                .find(|page| page.region_section_index == anchor_page.region_section_index)
                .unwrap_or(anchor_page),
            _ => anchor_page,
        };
        append_unique(&mut result, target.number, map_id);
    }
    result
}

pub fn build_note_presentations(
    refs: &[NoteRefLocation],
    pages: &[Page],
    regions: &DocumentRegions,
) -> OrderedMap<i64, NotePresentation> {
    let anchors = map_note_anchors_to_pages(pages, refs);
    let mut result = OrderedMap::new();
    let mut counters = HashMap::<String, i64>::new();
    for reference in refs {
        let map_id = reference.map_id();
        if result.get(&map_id).is_some() {
            continue;
        }
        let Some(page) = page_for_map_id(pages, &anchors, map_id) else {
            continue;
        };
        let properties =
            regions.note_properties(page.region_section_index, reference.note_kind.as_str());
        let key = match properties.num_restart.as_deref() {
            Some("eachPage") => format!("{}:page:{}", reference.note_kind.as_str(), page.number),
            Some("eachSect") => format!(
                "{}:section:{}",
                reference.note_kind.as_str(),
                page.section_id
                    .clone()
                    .unwrap_or_else(|| page.region_section_index.to_string())
            ),
            _ => format!("{}:continuous", reference.note_kind.as_str()),
        };
        let display_number = counters
            .get(&key)
            .copied()
            .unwrap_or(properties.num_start.unwrap_or(1));
        counters.insert(key, display_number + 1);
        result.set(
            map_id,
            NotePresentation {
                display_number,
                display_label: format_number(
                    display_number,
                    properties.num_fmt.as_deref().unwrap_or("decimal"),
                ),
                anchor: NoteAnchor {
                    doc_start: Some(reference.pm_pos),
                    doc_end: Some(reference.pm_pos + 1.0),
                },
            },
        );
    }
    result
}

pub fn assign_note_presentations(
    contents: &mut [NoteContent],
    presentations: &OrderedMap<i64, NotePresentation>,
) {
    for content in contents {
        let Some(presentation) = presentations.get(&content.map_id()) else {
            continue;
        };
        content.display_number = Some(presentation.display_number);
        content
            .display_label
            .clone_from(&Some(presentation.display_label.clone()));
        content.anchor = Some(presentation.anchor.clone());
    }
}

pub fn apply_note_presentation(
    blocks: &mut Vec<LayoutBlock>,
    display_number: i64,
    display_label: &str,
) {
    if blocks.is_empty() {
        blocks.push(LayoutBlock::Paragraph(ParagraphBlock {
            sdt_groups: None,
            id: BlockId::Str(format!("fn-empty-{display_number}")),
            para_id: None,
            runs: Vec::new(),
            attrs: None,
            pm_start: None,
            pm_end: None,
        }));
    }
    for block in blocks.iter_mut() {
        let LayoutBlock::Paragraph(paragraph) = block else {
            continue;
        };
        for run in &mut paragraph.runs {
            match run {
                Run::Text(text) if text.fmt.font_size.is_none() => {
                    text.fmt.font_size = Some(FOOTNOTE_FONT_SIZE_PT);
                }
                Run::Tab(tab) if tab.fmt.font_size.is_none() => {
                    tab.fmt.font_size = Some(FOOTNOTE_FONT_SIZE_PT);
                }
                _ => {}
            }
        }
    }
    let Some(LayoutBlock::Paragraph(first)) = blocks.first_mut() else {
        return;
    };
    let font_family = first.runs.iter().find_map(|run| match run {
        Run::Text(text) => text
            .fmt
            .font_family
            .as_ref()
            .filter(|family| !family.is_empty())
            .cloned(),
        _ => None,
    });
    first.runs.insert(
        0,
        Run::Text(TextRun {
            fmt: RunFormatting {
                font_family,
                font_size: Some(FOOTNOTE_FONT_SIZE_PT),
                superscript: Some(true),
                ..Default::default()
            },
            text: format!("{display_label}  "),
            pm_start: None,
            pm_end: None,
            inline_sdt_widget: None,
        }),
    );
}

pub trait HasHeight {
    fn height(&self) -> f64;
}

impl<T: HasHeight + ?Sized> HasHeight for &T {
    fn height(&self) -> f64 {
        (*self).height()
    }
}

impl HasHeight for NoteContent {
    fn height(&self) -> f64 {
        self.height
    }
}

pub fn distribute_notes_into_columns<T: HasHeight>(items: Vec<T>, columns: u64) -> Vec<Vec<T>> {
    let columns = columns.max(1) as usize;
    if columns == 1 || items.len() <= 1 {
        return vec![items];
    }
    let target = items.iter().map(HasHeight::height).sum::<f64>() / columns as f64;
    let mut result = vec![Vec::new()];
    let mut column_height = 0.0;
    for item in items {
        if result.len() < columns
            && column_height > 0.0
            && column_height + item.height() / 2.0 > target
        {
            result.push(Vec::new());
            column_height = 0.0;
        }
        column_height += item.height();
        result.last_mut().expect("column exists").push(item);
    }
    result
}

fn content_map(contents: &[NoteContent]) -> OrderedMap<i64, &NoteContent> {
    contents
        .iter()
        .map(|content| (content.map_id(), content))
        .collect()
}

pub fn footnote_columns_by_page(pages: &[Page], regions: &DocumentRegions) -> OrderedMap<u32, u64> {
    pages
        .iter()
        .map(|page| {
            (
                page.number,
                regions.footnote_columns(page.region_section_index),
            )
        })
        .collect()
}

pub fn calculate_note_reserved_heights(
    page_note_map: &OrderedMap<u32, Vec<i64>>,
    contents: &[NoteContent],
    columns_by_page: &OrderedMap<u32, u64>,
) -> OrderedMap<u32, f64> {
    let contents = content_map(contents);
    let mut reserved = OrderedMap::new();
    for (page_number, note_ids) in page_note_map.iter() {
        let mut total_height = 0.0;
        for kind in [NoteKind::Footnote, NoteKind::Endnote] {
            let heights: Vec<Height> = note_ids
                .iter()
                .filter_map(|id| contents.get(id).copied())
                .filter(|content| content.note_kind == kind && content.height > 0.0)
                .map(|content| Height(content.height))
                .collect();
            if heights.is_empty() {
                continue;
            }
            let columns = if kind == NoteKind::Footnote {
                columns_by_page.get(page_number).copied().unwrap_or(1)
            } else {
                1
            };
            let tallest = distribute_notes_into_columns(heights, columns)
                .iter()
                .map(|column| column.iter().map(HasHeight::height).sum::<f64>())
                .fold(0.0_f64, f64::max);
            total_height += tallest + FOOTNOTE_SEPARATOR_HEIGHT;
        }
        if total_height > 0.0 {
            reserved.set(*page_number, total_height);
        }
    }
    reserved
}

#[derive(Clone, Copy)]
struct Height(f64);

impl HasHeight for Height {
    fn height(&self) -> f64 {
        self.0
    }
}

pub fn reserved_heights_equal(left: &OrderedMap<u32, f64>, right: &OrderedMap<u32, f64>) -> bool {
    left.entries.len() == right.entries.len()
        && left
            .iter()
            .all(|(page, height)| right.get(page).copied() == Some(*height))
}

fn reserved_heights_cover(
    reserved: &OrderedMap<u32, f64>,
    required: &OrderedMap<u32, f64>,
) -> bool {
    required
        .iter()
        .all(|(page, height)| reserved.get(page).copied().unwrap_or(0.0) >= *height)
}

fn merge_reserved_heights(
    left: &OrderedMap<u32, f64>,
    right: &OrderedMap<u32, f64>,
) -> OrderedMap<u32, f64> {
    let mut merged = left.clone();
    for (page, height) in right.iter() {
        merged.set(*page, merged.get(page).copied().unwrap_or(0.0).max(*height));
    }
    merged
}

pub fn stabilize_note_layout<F>(
    mut layout_with_reserved: F,
    refs: &[NoteRefLocation],
    contents: &[NoteContent],
    initial_layout: Layout,
    regions: &DocumentRegions,
) -> Result<StabilizedNoteLayout, LayoutError>
where
    F: FnMut(&OrderedMap<u32, f64>) -> Result<Layout, LayoutError>,
{
    let mut page_note_map = map_notes_to_pages(&initial_layout.pages, refs, regions);
    let mut columns = footnote_columns_by_page(&initial_layout.pages, regions);
    let mut reserved = calculate_note_reserved_heights(&page_note_map, contents, &columns);
    if reserved.is_empty() {
        return Ok(StabilizedNoteLayout {
            layout: initial_layout,
            page_note_map,
            reserved_heights: reserved,
            converged: true,
        });
    }

    let mut layout = initial_layout;
    let mut converged = false;
    for _ in 0..MAX_FOOTNOTE_LAYOUT_PASSES {
        layout = layout_with_reserved(&reserved)?;
        page_note_map = map_notes_to_pages(&layout.pages, refs, regions);
        columns = footnote_columns_by_page(&layout.pages, regions);
        let next = calculate_note_reserved_heights(&page_note_map, contents, &columns);
        if reserved_heights_equal(&reserved, &next) {
            reserved = next;
            converged = true;
            break;
        }
        reserved = next;
    }

    if !converged {
        let mut fallback = reserved;
        let mut covered = false;
        for _ in 0..MAX_FOOTNOTE_LAYOUT_PASSES {
            layout = layout_with_reserved(&fallback)?;
            page_note_map = map_notes_to_pages(&layout.pages, refs, regions);
            columns = footnote_columns_by_page(&layout.pages, regions);
            let required = calculate_note_reserved_heights(&page_note_map, contents, &columns);
            if reserved_heights_cover(&fallback, &required) {
                covered = true;
                break;
            }
            fallback = merge_reserved_heights(&fallback, &required);
        }
        if !covered {
            layout = layout_with_reserved(&fallback)?;
            page_note_map = map_notes_to_pages(&layout.pages, refs, regions);
        }
        reserved = fallback;
    }

    stamp_note_pages(&mut layout, &page_note_map, regions);
    Ok(StabilizedNoteLayout {
        layout,
        page_note_map,
        reserved_heights: reserved,
        converged,
    })
}

pub fn stamp_note_pages(
    layout: &mut Layout,
    page_note_map: &OrderedMap<u32, Vec<i64>>,
    regions: &DocumentRegions,
) {
    for page in &mut layout.pages {
        let Some(ids) = page_note_map.get(&page.number) else {
            continue;
        };
        page.footnote_ids = Some(ids.iter().map(|id| *id as f64).collect());
        let columns = regions.footnote_columns(page.region_section_index);
        if columns > 1 {
            page.footnote_columns = Some(columns as f64);
        }
    }
}

fn fragment_bottom(fragment: &Fragment) -> f64 {
    match fragment {
        Fragment::Paragraph(value) => value.y + value.height,
        Fragment::Table(value) => value.y + value.height,
        Fragment::Image(value) => value.y + value.height,
        Fragment::Shape(value) => value.y + value.height,
        Fragment::Chart(value) => value.y + value.height,
        Fragment::TextBox(value) => value.y + value.height,
    }
}

fn note_item(content: &NoteContent) -> NoteLayoutItemContract {
    NoteLayoutItemContract {
        kind: Some(content.note_kind.as_str().to_owned()),
        id: Some(content.id),
        display_label: content.display_label.clone().or_else(|| {
            content
                .display_number
                .map(|display_number| display_number.to_string())
        }),
        blocks: Some(
            content
                .blocks
                .iter()
                .filter_map(|block| serde_json::to_value(block).ok())
                .collect(),
        ),
        measures: Some(
            content
                .measures
                .iter()
                .filter_map(|measure| serde_json::to_value(measure).ok())
                .collect(),
        ),
        height: Some(content.height),
        anchor_doc_start: content
            .anchor
            .as_ref()
            .and_then(|anchor| anchor.doc_start)
            .map(|value| value as i64),
        anchor_doc_end: content
            .anchor
            .as_ref()
            .and_then(|anchor| anchor.doc_end)
            .map(|value| value as i64),
        custom_mark_follows: content.custom_mark_follows,
    }
}

pub fn attach_note_areas(
    layout: &mut Layout,
    page_note_map: &OrderedMap<u32, Vec<i64>>,
    contents: &[NoteContent],
    regions: &DocumentRegions,
) {
    let contents = content_map(contents);
    for page in &mut layout.pages {
        let Some(ids) = page_note_map.get(&page.number) else {
            continue;
        };
        let mut groups = OrderedMap::<NoteKind, Vec<&NoteContent>>::new();
        for content in ids.iter().filter_map(|id| contents.get(id).copied()) {
            if let Some(group) = groups.get_mut(&content.note_kind) {
                group.push(content);
            } else {
                groups.set(content.note_kind, vec![content]);
            }
        }
        let content_bottom = page.size.h - page.margins.bottom;
        let last_body_bottom = page
            .fragments
            .iter()
            .map(fragment_bottom)
            .fold(page.margins.top, f64::max);
        let mut bottom_cursor = content_bottom;
        let mut beneath_text_cursor = last_body_bottom;
        let mut areas = Vec::new();
        for (kind, group) in groups.iter() {
            let properties = regions.note_properties(page.region_section_index, kind.as_str());
            let default_position = match kind {
                NoteKind::Footnote => "pageBottom",
                NoteKind::Endnote => "docEnd",
            };
            let placement = properties
                .position
                .unwrap_or_else(|| default_position.to_owned());
            let columns = if *kind == NoteKind::Footnote {
                regions.footnote_columns(page.region_section_index)
            } else {
                1
            };
            let height = distribute_notes_into_columns(group.clone(), columns)
                .iter()
                .map(|column| column.iter().map(|content| content.height).sum::<f64>())
                .fold(0.0_f64, f64::max)
                + FOOTNOTE_SEPARATOR_HEIGHT;
            let y = if placement == "beneathText" {
                let y = beneath_text_cursor.min(content_bottom - height);
                beneath_text_cursor = y + height;
                y
            } else {
                bottom_cursor -= height;
                bottom_cursor
            };
            areas.push(NoteAreaContract {
                page_index: None,
                section_id: page.section_id.clone(),
                kind: Some(kind.as_str().to_owned()),
                placement: Some(placement),
                y: Some(y),
                height: Some(height),
                columns: Some(columns),
                separator: None,
                notes: Some(group.iter().map(|content| note_item(content)).collect()),
            });
        }
        if !areas.is_empty() {
            page.note_areas = Some(areas);
        }
    }
}

#[cfg(test)]
mod tests {
    use serde_json::json;

    use super::*;
    use crate::regions::{NoteProperties, NoteSettings, RegionSection};
    use crate::types::{PageMargins, ParagraphFragment, Size, TableFragment};

    fn block(value: serde_json::Value) -> LayoutBlock {
        serde_json::from_value(value).unwrap()
    }

    fn page(number: u32, section: usize, fragments: Vec<Fragment>) -> Page {
        Page {
            number,
            fragments,
            margins: PageMargins {
                top: 96.0,
                right: 96.0,
                bottom: 96.0,
                left: 96.0,
                header: None,
                footer: None,
            },
            size: Size {
                w: 816.0,
                h: 1056.0,
            },
            orientation: None,
            section_index: Some(section as u64),
            region_section_index: section,
            header_footer_refs: None,
            footnote_ids: None,
            footnote_reserved_height: None,
            footnote_columns: None,
            columns: None,
            section_id: Some(section.to_string()),
            section_page_index: None,
            section_page_number: None,
            page_label: None,
            page_numbering: None,
            header_distance: None,
            footer_distance: None,
            page_borders: None,
            watermark: None,
            vertical_align: None,
            note_areas: None,
        }
    }

    fn paragraph_fragment(start: f64, end: f64) -> Fragment {
        Fragment::Paragraph(ParagraphFragment {
            block_id: BlockId::Str("p".to_owned()),
            x: 96.0,
            y: 96.0,
            width: 624.0,
            pm_start: Some(start),
            pm_end: Some(end),
            from_line: 0,
            to_line: 1,
            height: 20.0,
            carried_from_prev: None,
            carried_to_next: None,
            resolved_lines: None,
        })
    }

    fn layout(pages: Vec<Page>) -> Layout {
        Layout {
            page_size: Size {
                w: 816.0,
                h: 1056.0,
            },
            pages,
            columns: None,
            headers: None,
            footers: None,
            page_gap: None,
        }
    }

    fn content(id: i64, kind: NoteKind, height: f64) -> NoteContent {
        NoteContent {
            id,
            note_kind: kind,
            display_number: None,
            display_label: None,
            blocks: Vec::new(),
            measures: Vec::new(),
            height,
            anchor: None,
            custom_mark_follows: None,
        }
    }

    #[test]
    fn collects_footnotes_endnotes_and_outer_table_rows() {
        let blocks = vec![
            block(json!({
                "kind": "paragraph",
                "id": "p1",
                "runs": [
                    {"kind": "text", "text": "a", "pmStart": 10, "footnoteRefId": 7},
                    {"kind": "text", "text": "b", "pmStart": 11, "endnoteRefId": 7}
                ]
            })),
            block(json!({
                "kind": "table",
                "id": "t1",
                "rows": [{
                    "id": "r1",
                    "cells": [{
                        "id": "c1",
                        "blocks": [{
                            "kind": "paragraph",
                            "id": "cp",
                            "runs": [{"kind": "text", "text": "c", "pmStart": 20, "footnoteRefId": 8}]
                        }]
                    }]
                }]
            })),
        ];

        let refs = collect_note_refs(&blocks);
        assert_eq!(refs.len(), 3);
        assert_eq!(refs[0].map_id(), 7);
        assert_eq!(refs[1].map_id(), -8);
        assert_eq!(refs[2].table_block_id, Some(BlockId::Str("t1".to_owned())));
        assert_eq!(refs[2].row_index, Some(0));
    }

    #[test]
    fn maps_split_table_rows_to_their_actual_pages() {
        let fragment = |start, end| {
            Fragment::Table(TableFragment {
                block_id: BlockId::Str("t1".to_owned()),
                x: 96.0,
                y: 96.0,
                width: 624.0,
                pm_start: Some(5.0),
                pm_end: Some(80.0),
                row_start: start,
                row_end: end,
                height: 40.0,
                is_floating: None,
                carried_from_prev: None,
                carried_to_next: None,
                header_row_count: None,
                clip_top: None,
                clip_bottom: None,
            })
        };
        let pages = vec![
            page(1, 0, vec![fragment(0, 2)]),
            page(2, 0, vec![fragment(2, 4)]),
        ];
        let refs = vec![
            NoteRefLocation {
                note_id: 1,
                note_kind: NoteKind::Footnote,
                pm_pos: 10.0,
                table_block_id: Some(BlockId::Str("t1".to_owned())),
                row_index: Some(1),
            },
            NoteRefLocation {
                note_id: 2,
                note_kind: NoteKind::Footnote,
                pm_pos: 20.0,
                table_block_id: Some(BlockId::Str("t1".to_owned())),
                row_index: Some(2),
            },
        ];

        let mapped = map_note_anchors_to_pages(&pages, &refs);
        assert_eq!(mapped.get(&1), Some(&vec![1]));
        assert_eq!(mapped.get(&2), Some(&vec![2]));
    }

    #[test]
    fn places_endnotes_at_document_end_without_id_collisions() {
        let pages = vec![
            page(1, 0, vec![paragraph_fragment(0.0, 20.0)]),
            page(2, 0, vec![paragraph_fragment(20.0, 40.0)]),
        ];
        let refs = vec![
            NoteRefLocation {
                note_id: 1,
                note_kind: NoteKind::Footnote,
                pm_pos: 5.0,
                table_block_id: None,
                row_index: None,
            },
            NoteRefLocation {
                note_id: 1,
                note_kind: NoteKind::Endnote,
                pm_pos: 6.0,
                table_block_id: None,
                row_index: None,
            },
        ];

        let mapped = map_notes_to_pages(&pages, &refs, &DocumentRegions::default());
        assert_eq!(mapped.get(&1), Some(&vec![1]));
        assert_eq!(mapped.get(&2), Some(&vec![-2]));
    }

    #[test]
    fn balances_footnotes_but_stacks_endnotes_separately() {
        let page_map = [(1, vec![1, 2, 3, 4, -2, -3])].into_iter().collect();
        let columns = [(1, 2)].into_iter().collect();
        let contents = vec![
            content(1, NoteKind::Footnote, 10.0),
            content(2, NoteKind::Footnote, 10.0),
            content(3, NoteKind::Footnote, 10.0),
            content(4, NoteKind::Footnote, 10.0),
            content(1, NoteKind::Endnote, 8.0),
            content(2, NoteKind::Endnote, 9.0),
        ];

        let reserved = calculate_note_reserved_heights(&page_map, &contents, &columns);
        assert_eq!(reserved.get(&1), Some(&(20.0 + 12.0 + 17.0 + 12.0)));
    }

    #[test]
    fn restarts_labels_per_page_and_uses_authored_format() {
        let pages = vec![
            page(1, 0, vec![paragraph_fragment(0.0, 10.0)]),
            page(2, 0, vec![paragraph_fragment(10.0, 20.0)]),
        ];
        let refs = vec![
            NoteRefLocation {
                note_id: 1,
                note_kind: NoteKind::Footnote,
                pm_pos: 2.0,
                table_block_id: None,
                row_index: None,
            },
            NoteRefLocation {
                note_id: 2,
                note_kind: NoteKind::Footnote,
                pm_pos: 12.0,
                table_block_id: None,
                row_index: None,
            },
        ];
        let regions = DocumentRegions {
            note_settings: NoteSettings {
                footnote: NoteProperties {
                    num_fmt: Some("upperRoman".to_owned()),
                    num_start: Some(3),
                    num_restart: Some("eachPage".to_owned()),
                    ..Default::default()
                },
                ..Default::default()
            },
            ..Default::default()
        };

        let presentations = build_note_presentations(&refs, &pages, &regions);
        assert_eq!(presentations.get(&1).unwrap().display_label, "III");
        assert_eq!(presentations.get(&2).unwrap().display_label, "III");
    }

    #[test]
    fn stabilization_reenters_until_reservations_match_the_latest_layout() {
        let initial = layout(vec![page(1, 0, vec![paragraph_fragment(0.0, 10.0)])]);
        let shifted = layout(vec![
            page(1, 0, Vec::new()),
            page(2, 0, vec![paragraph_fragment(0.0, 10.0)]),
        ]);
        let refs = vec![NoteRefLocation {
            note_id: 1,
            note_kind: NoteKind::Footnote,
            pm_pos: 5.0,
            table_block_id: None,
            row_index: None,
        }];
        let contents = vec![content(1, NoteKind::Footnote, 20.0)];
        let mut passes = 0;
        let result = stabilize_note_layout(
            |_| {
                passes += 1;
                Ok(shifted.clone())
            },
            &refs,
            &contents,
            initial,
            &DocumentRegions::default(),
        )
        .unwrap();

        assert!(result.converged);
        assert_eq!(passes, 2);
        assert_eq!(result.reserved_heights.get(&2), Some(&32.0));
        assert_eq!(result.layout.pages[1].footnote_ids, Some(vec![1.0]));
    }

    #[test]
    fn attaches_typed_beneath_text_note_area() {
        let mut output = layout(vec![page(1, 0, vec![paragraph_fragment(0.0, 10.0)])]);
        let map = [(1, vec![1])].into_iter().collect();
        let contents = vec![content(1, NoteKind::Footnote, 20.0)];
        let regions = DocumentRegions {
            sections: vec![RegionSection {
                note_settings: NoteSettings {
                    footnote: NoteProperties {
                        position: Some("beneathText".to_owned()),
                        ..Default::default()
                    },
                    ..Default::default()
                },
                ..Default::default()
            }],
            ..Default::default()
        };

        attach_note_areas(&mut output, &map, &contents, &regions);
        let area = &output.pages[0].note_areas.as_ref().unwrap()[0];
        assert_eq!(area.placement.as_deref(), Some("beneathText"));
        assert_eq!(area.y, Some(116.0));
        assert_eq!(area.height, Some(32.0));
    }

    #[test]
    fn presentation_uses_note_font_and_label() {
        let mut blocks = vec![block(json!({
            "kind": "paragraph",
            "id": "fn",
            "runs": [{"kind": "text", "text": "note", "fontFamily": "Cambria"}]
        }))];

        apply_note_presentation(&mut blocks, 3, "III");
        let LayoutBlock::Paragraph(paragraph) = &blocks[0] else {
            panic!("paragraph expected");
        };
        let Run::Text(marker) = &paragraph.runs[0] else {
            panic!("text expected");
        };
        assert_eq!(marker.text, "III  ");
        assert_eq!(marker.fmt.font_family.as_deref(), Some("Cambria"));
        assert_eq!(marker.fmt.superscript, Some(true));
    }
}
