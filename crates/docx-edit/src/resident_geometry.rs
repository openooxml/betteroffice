use std::collections::{BTreeMap, HashMap, HashSet};

use serde::{Deserialize, Serialize};
use yrs::Any;

use crate::geometry::{GeometryPositionOutline, GeometryRead};
use crate::{EditingDoc, SegmentContent};

#[derive(Clone, Deserialize, Serialize)]
#[serde(rename_all = "camelCase")]
pub(crate) struct ResidentLoc {
    pub story: String,
    pub para_id: String,
    pub offset: u64,
}

#[derive(Default, Deserialize)]
pub(crate) struct ResidentRequest {
    pub ranges: Vec<[u64; 2]>,
    pub selection: Option<[ResidentLoc; 2]>,
}

#[derive(Clone, Serialize)]
#[serde(rename_all = "camelCase")]
pub(crate) struct ResidentParagraph {
    pub story: String,
    pub para_id: String,
    pub index: u32,
    pub from: u64,
    pub to: u64,
    pub leading: u64,
    /// Story units, leading blocks included.
    pub length: u64,
    pub story_order: usize,
    #[serde(skip_serializing_if = "Vec::is_empty")]
    pub controls: Vec<ResidentInlineControl>,
}

/// An inline control node: its story offset in the paragraph and its display range.
#[derive(Clone, Serialize)]
#[serde(rename_all = "camelCase")]
pub(crate) struct ResidentInlineControl {
    pub offset: u64,
    pub from: u64,
    pub to: u64,
}

/// A block control node and the story holding its content.
#[derive(Clone, Serialize)]
#[serde(rename_all = "camelCase")]
pub(crate) struct ResidentControl {
    pub story: String,
    pub content_story: String,
    pub from: u64,
    pub to: u64,
}

#[derive(Clone, Serialize)]
#[serde(rename_all = "camelCase")]
pub(crate) struct ResidentCell {
    pub story: String,
    pub para_id: String,
    pub row: usize,
    pub column: usize,
    pub rowspan: usize,
    pub colspan: usize,
    pub from: u64,
    pub to: u64,
}

#[derive(Clone, Serialize)]
#[serde(rename_all = "camelCase")]
pub(crate) struct ResidentTable {
    pub story: String,
    pub table_index: usize,
    pub from: u64,
    pub to: u64,
    pub at: ResidentLoc,
    pub cells: Vec<ResidentCell>,
}

#[derive(Serialize)]
pub(crate) struct ResidentGeometry {
    pub paragraphs: Vec<ResidentParagraph>,
    pub tables: Vec<ResidentTable>,
    pub controls: Vec<ResidentControl>,
    pub selection: Option<[u64; 2]>,
}

#[derive(Clone)]
enum Item {
    Paragraph(ResidentParagraph),
    Table(usize),
    Cell(usize, ResidentCell),
    Control(usize),
}

struct Entry {
    from: usize,
    to: usize,
    item: Item,
    /// The paragraph's inline controls: the point of each node's start and its display units.
    controls: Vec<(usize, u64)>,
}

/// A record before indexing: its range, its item and its inline control nodes.
type Record = (u64, u64, Item, Vec<(u64, u64)>);

pub(crate) struct ResidentGeometryIndex {
    epoch: u64,
    revision: u64,
    points: Vec<(u64, bool)>,
    shifts: Vec<i64>,
    entries: Vec<Entry>,
    maxima: Vec<usize>,
    carets: HashMap<(String, String), usize>,
    stories: HashSet<String>,
    tables: Vec<ResidentTable>,
    table_entries: Vec<usize>,
    controls: Vec<ResidentControl>,
}

fn span(value: Option<&Any>) -> usize {
    match value {
        Some(Any::Number(value)) if value.is_finite() && *value >= 1.0 && value.fract() == 0.0 => {
            *value as usize
        }
        _ => 1,
    }
}

fn map(value: Option<&Any>) -> Option<&HashMap<String, Any>> {
    match value {
        Some(Any::Map(value)) => Some(value),
        _ => None,
    }
}

fn array(value: Option<&Any>) -> &[Any] {
    match value {
        Some(Any::Array(value)) => value,
        _ => &[],
    }
}

fn child(value: Option<&Any>, fallback: String) -> String {
    match value {
        Some(Any::String(value)) => value.to_string(),
        _ => fallback,
    }
}

impl ResidentGeometryIndex {
    fn build(doc: &EditingDoc, root: &str, epoch: u64) -> Result<Self, String> {
        let GeometryRead::Value(outline) = doc.geometry_position_outline(root) else {
            return Err("unsupported geometry projection".to_owned());
        };
        Self::from_outline(doc, root, epoch, outline)
    }

    fn from_outline(
        doc: &EditingDoc,
        root: &str,
        epoch: u64,
        outline: GeometryPositionOutline,
    ) -> Result<Self, String> {
        let mut records: Vec<Record> = Vec::new();
        let mut tables = Vec::new();
        let mut controls = Vec::new();
        for (story, projected) in &outline {
            let paragraphs: HashMap<_, _> = projected
                .paragraphs
                .iter()
                .map(|paragraph| (paragraph.para_id.as_str(), paragraph))
                .collect();
            let units = doc
                .story_segment_units(story)
                .map_err(|error| error.to_string())?;
            let mut cursor = projected.content_start;
            let mut table_index = 0;
            for (index, unit) in units.iter().enumerate() {
                let Some(properties) = unit.iter().find_map(|segment| match &segment.content {
                    SegmentContent::Pilcrow(properties) => Some(properties),
                    _ => None,
                }) else {
                    continue;
                };
                let Some(paragraph) = paragraphs.get(properties.para_id.as_str()) else {
                    continue;
                };
                let from = projected.content_start + paragraph.display_start + 1;
                let to = from + paragraph.length;
                let mut expanded = 0;
                let inline = paragraph
                    .controls
                    .iter()
                    .map(|&[offset, size]| {
                        let start = from + offset + expanded;
                        expanded += size - 1;
                        (start, size)
                    })
                    .collect();
                records.push((
                    from,
                    to,
                    Item::Paragraph(ResidentParagraph {
                        story: story.clone(),
                        para_id: properties.para_id.clone(),
                        index: index as u32,
                        from,
                        to,
                        leading: paragraph.leading,
                        length: paragraph.length - expanded + paragraph.leading,
                        story_order: 0,
                        controls: Vec::new(),
                    }),
                    inline,
                ));
                let mut leading = 0;
                for segment in unit {
                    let SegmentContent::OtherEmbed { kind, payload } = &segment.content else {
                        continue;
                    };
                    if kind == "table" {
                        let id = tables.len();
                        let start = cursor;
                        let mut size = 0;
                        let mut occupied: HashMap<usize, HashSet<usize>> = HashMap::new();
                        let mut cells = Vec::new();
                        for (row, data) in array(payload.get("rows")).iter().enumerate() {
                            let mut row_size = 0;
                            let mut column = 0;
                            for (cell_index, data) in
                                array(map(Some(data)).and_then(|row| row.get("cells")))
                                    .iter()
                                    .enumerate()
                            {
                                while occupied
                                    .get(&row)
                                    .is_some_and(|used| used.contains(&column))
                                {
                                    column += 1;
                                }
                                let data = map(Some(data));
                                let properties = map(data.and_then(|cell| cell.get("tcPr")));
                                let rowspan =
                                    span(properties.and_then(|value| value.get("rowspan")));
                                let colspan =
                                    span(properties.and_then(|value| value.get("colspan")));
                                let cell_story = child(
                                    data.and_then(|cell| cell.get("story")),
                                    format!("{story}:t{table_index}:r{row}c{cell_index}"),
                                );
                                let content_size = outline
                                    .get(&cell_story)
                                    .ok_or_else(|| "missing cell projection".to_owned())?
                                    .size;
                                let cell_from = start + 2 + size + row_size;
                                let cell_to = cell_from + content_size + 2;
                                let para_id = outline[&cell_story]
                                    .paragraphs
                                    .first()
                                    .ok_or_else(|| "missing cell paragraph".to_owned())?
                                    .para_id
                                    .clone();
                                let cell = ResidentCell {
                                    story: cell_story,
                                    para_id,
                                    row,
                                    column,
                                    rowspan,
                                    colspan,
                                    from: cell_from,
                                    to: cell_to,
                                };
                                records.push((
                                    cell_from,
                                    cell_to,
                                    Item::Cell(id, cell.clone()),
                                    Vec::new(),
                                ));
                                cells.push(cell);
                                for row in row..row + rowspan {
                                    occupied
                                        .entry(row)
                                        .or_default()
                                        .extend(column..column + colspan);
                                }
                                column += colspan;
                                row_size += content_size + 2;
                            }
                            size += row_size + 2;
                        }
                        let end = start + size + 2;
                        tables.push(ResidentTable {
                            story: story.clone(),
                            table_index,
                            from: start,
                            to: end,
                            at: ResidentLoc {
                                story: story.clone(),
                                para_id: paragraph.para_id.clone(),
                                offset: leading,
                            },
                            cells,
                        });
                        records.push((start, end, Item::Table(id), Vec::new()));
                        cursor = end;
                        leading += 1;
                        table_index += 1;
                    } else if kind == "blockSdt" {
                        let content_story = child(payload.get("story"), format!("{story}:sdt0"));
                        let size = outline
                            .get(&content_story)
                            .ok_or_else(|| "missing control projection".to_owned())?
                            .size;
                        let end = cursor + size + 2;
                        records.push((cursor, end, Item::Control(controls.len()), Vec::new()));
                        controls.push(ResidentControl {
                            story: story.clone(),
                            content_story,
                            from: cursor,
                            to: end,
                        });
                        cursor = end;
                        leading += 1;
                    } else if kind == "pageBreak" || kind == "columnBreak" {
                        cursor += 1;
                        leading += 1;
                    }
                }
                cursor = projected.content_start + paragraph.display_start + paragraph.length + 2;
            }
        }
        let mut containers: Vec<(u64, u64, Vec<&str>)> = tables
            .iter()
            .map(|table| {
                let stories = table.cells.iter().map(|cell| cell.story.as_str());
                (table.from, table.to, stories.collect())
            })
            .chain(controls.iter().map(|control| {
                (
                    control.from,
                    control.to,
                    vec![control.content_story.as_str()],
                )
            }))
            .collect();
        containers.sort_by(|a, b| nested_first((a.0, a.1), (b.0, b.1)));
        let mut story_order = HashMap::from([(root.to_owned(), 0)]);
        for (_, _, stories) in containers {
            for story in stories {
                let next = story_order.len();
                story_order.entry(story.to_owned()).or_insert(next);
            }
        }
        let mut keys = BTreeMap::new();
        for (from, to, _, inline) in &records {
            keys.insert((*from, false), 0);
            keys.insert((*to, true), 0);
            // Text typed at a control's start moves the control.
            for (start, _) in inline {
                keys.insert((*start, true), 0);
            }
        }
        let points: Vec<_> = keys.keys().copied().collect();
        for (index, key) in points.iter().enumerate() {
            keys.insert(*key, index);
        }
        let mut entries: Vec<_> = records
            .into_iter()
            .map(|(from, to, mut item, inline)| {
                if let Item::Paragraph(paragraph) = &mut item {
                    paragraph.story_order =
                        *story_order.get(&paragraph.story).unwrap_or(&usize::MAX);
                }
                Entry {
                    from: keys[&(from, false)],
                    to: keys[&(to, true)],
                    item,
                    controls: inline
                        .into_iter()
                        .map(|(start, size)| (keys[&(start, true)], size))
                        .collect(),
                }
            })
            .collect();
        entries.sort_by_key(|entry| entry.from);
        let mut table_entries = vec![0; tables.len()];
        for (index, entry) in entries.iter().enumerate() {
            if let Item::Table(id) = entry.item {
                table_entries[id] = index;
            }
        }
        let carets = entries
            .iter()
            .enumerate()
            .filter_map(|(index, entry)| {
                if let Item::Paragraph(paragraph) = &entry.item {
                    Some(((paragraph.story.clone(), paragraph.para_id.clone()), index))
                } else {
                    None
                }
            })
            .collect();
        let mut result = Self {
            epoch,
            revision: doc.story_revisions.lock().unwrap().geometry,
            shifts: vec![0; points.len() + 1],
            points,
            maxima: vec![0; entries.len().max(1) * 4],
            entries,
            carets,
            stories: outline.keys().cloned().collect(),
            tables,
            table_entries,
            controls,
        };
        if !result.entries.is_empty() {
            result.build_maxima(1, 0, result.entries.len());
        }
        Ok(result)
    }

    fn build_maxima(&mut self, node: usize, from: usize, to: usize) -> usize {
        let maximum = if to - from == 1 {
            self.entries[from].to
        } else {
            let middle = (from + to) / 2;
            let left = self.build_maxima(node * 2, from, middle);
            let right = self.build_maxima(node * 2 + 1, middle, to);
            left.max(right)
        };
        self.maxima[node] = maximum;
        maximum
    }

    fn position(&self, point: usize) -> u64 {
        let mut index = point + 1;
        let mut shift = 0;
        while index > 0 {
            shift += self.shifts[index];
            index &= index - 1;
        }
        (self.points[point].0 as i64 + shift) as u64
    }

    /// Display units from `entry`'s content start `from` to its story offset `offset`.
    fn display_offset(&self, entry: &Entry, from: u64, offset: u64) -> u64 {
        let mut expanded = 0;
        for &(point, size) in &entry.controls {
            if self.position(point).saturating_sub(from + expanded) >= offset {
                break;
            }
            expanded += size - 1;
        }
        offset.saturating_add(expanded)
    }

    fn locate(&self, at: &ResidentLoc) -> Option<u64> {
        let entry = &self.entries[*self.carets.get(&(at.story.clone(), at.para_id.clone()))?];
        let Item::Paragraph(paragraph) = &entry.item else {
            return None;
        };
        let from = self.position(entry.from);
        let to = self.position(entry.to);
        let offset = at.offset.saturating_sub(paragraph.leading);
        Some(from + self.display_offset(entry, from, offset).min(to - from))
    }

    fn move_text(
        &mut self,
        story: &str,
        para_id: &str,
        offset: u32,
        delta: i64,
        insertion: bool,
    ) -> bool {
        let Some(&entry) = self.carets.get(&(story.to_owned(), para_id.to_owned())) else {
            return false;
        };
        let start = self.position(self.entries[entry].from);
        let end = self.position(self.entries[entry].to);
        let offset = self.display_offset(&self.entries[entry], start, u64::from(offset));
        if offset > end - start {
            return false;
        }
        let at = start + offset;
        let (mut left, mut right) = (0, self.points.len());
        while left < right {
            let middle = (left + right) / 2;
            let position = self.position(middle);
            let before = if insertion {
                (position, self.points[middle].1) <= (at, false)
            } else {
                position < at
            };
            if before {
                left = middle + 1;
            } else {
                right = middle;
            }
        }
        let mut index = left + 1;
        while index < self.shifts.len() {
            self.shifts[index] += delta;
            index += index.isolate_lowest_one();
        }
        true
    }

    fn intersect(
        &self,
        node: usize,
        from: usize,
        to: usize,
        range: [u64; 2],
        found: &mut HashSet<usize>,
    ) {
        if self.position(self.maxima[node]) < range[0]
            || self.position(self.entries[from].from) > range[1]
        {
            return;
        }
        if to - from == 1 {
            found.insert(from);
        } else {
            let middle = (from + to) / 2;
            self.intersect(node * 2, from, middle, range, found);
            self.intersect(node * 2 + 1, middle, to, range, found);
        }
    }

    fn query(&self, request: &ResidentRequest) -> ResidentGeometry {
        let mut found = HashSet::new();
        if !self.entries.is_empty() {
            for &range in &request.ranges {
                self.intersect(1, 0, self.entries.len(), range, &mut found);
            }
        }
        let mut paragraphs = Vec::new();
        let mut selected: HashMap<usize, Vec<ResidentCell>> = HashMap::new();
        let mut controls = Vec::new();
        for index in found {
            let entry = &self.entries[index];
            match &entry.item {
                Item::Paragraph(value) => {
                    let mut value = value.clone();
                    value.from = self.position(entry.from);
                    value.to = self.position(entry.to);
                    let mut expanded = 0;
                    for &(point, size) in &entry.controls {
                        let from = self.position(point);
                        value.controls.push(ResidentInlineControl {
                            offset: from - value.from - expanded + value.leading,
                            from,
                            to: from + size,
                        });
                        expanded += size - 1;
                    }
                    value.length = value.to - value.from - expanded + value.leading;
                    if value.story_order != usize::MAX {
                        paragraphs.push(value);
                    }
                }
                Item::Control(id) => {
                    let mut value = self.controls[*id].clone();
                    value.from = self.position(entry.from);
                    value.to = self.position(entry.to);
                    controls.push(value);
                }
                Item::Table(id) => {
                    selected.entry(*id).or_default();
                }
                Item::Cell(id, value) => {
                    let mut value = value.clone();
                    value.from = self.position(entry.from);
                    value.to = self.position(entry.to);
                    selected.entry(*id).or_default().push(value);
                }
            }
        }
        paragraphs.sort_by_key(|paragraph| (paragraph.story_order, paragraph.index));
        let mut tables: Vec<_> = selected
            .into_iter()
            .map(|(id, mut cells)| {
                let table = &self.tables[id];
                let source = &self.entries[self.table_entries[id]];
                cells.sort_by_key(|cell| (cell.row, cell.column));
                ResidentTable {
                    story: table.story.clone(),
                    table_index: table.table_index,
                    from: self.position(source.from),
                    to: self.position(source.to),
                    at: table.at.clone(),
                    cells,
                }
            })
            .collect();
        tables.sort_by(|a, b| nested_first((a.from, a.to), (b.from, b.to)));
        controls.sort_by(|a, b| nested_first((a.from, a.to), (b.from, b.to)));
        let selection = request
            .selection
            .as_ref()
            .and_then(|range| Some([self.locate(&range[0])?, self.locate(&range[1])?]));
        ResidentGeometry {
            paragraphs,
            tables,
            controls,
            selection,
        }
    }
}

/// Orders ranges nested ones first, then by start.
fn nested_first(a: (u64, u64), b: (u64, u64)) -> std::cmp::Ordering {
    if a == b {
        std::cmp::Ordering::Equal
    } else if a.0 <= b.0 && a.1 >= b.1 {
        std::cmp::Ordering::Greater
    } else if b.0 <= a.0 && b.1 >= a.1 {
        std::cmp::Ordering::Less
    } else {
        a.0.cmp(&b.0)
    }
}

impl EditingDoc {
    pub(crate) fn resident_geometry(
        &self,
        root: &str,
        request: &ResidentRequest,
    ) -> Result<ResidentGeometry, String> {
        let epoch = self.committed_epoch();
        let revision = self.story_revisions.lock().unwrap().geometry;
        {
            let mut cache = self.resident_geometry.lock().unwrap();
            if let Some(index) = cache
                .get_mut(root)
                .filter(|index| index.revision == revision)
            {
                index.epoch = epoch;
                return Ok(index.query(request));
            }
        }
        let index = ResidentGeometryIndex::build(self, root, epoch)?;
        let result = index.query(request);
        self.resident_geometry
            .lock()
            .unwrap()
            .insert(root.to_owned(), index);
        Ok(result)
    }

    pub(crate) fn advance_resident_geometry(
        &self,
        story: &str,
        before: u64,
        after: u64,
        at: Option<(&str, u32)>,
        delta: i64,
        insertion: bool,
    ) {
        let revision = self.story_revisions.lock().unwrap().geometry;
        self.resident_geometry.lock().unwrap().retain(|_, index| {
            if index.epoch != before {
                return false;
            }
            if index.stories.contains(story) {
                let Some((para_id, offset)) = at else {
                    return false;
                };
                if !index.move_text(story, para_id, offset, delta, insertion) {
                    return false;
                }
            }
            index.epoch = after;
            index.revision = revision;
            true
        });
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::bridge::{RenderEnv, yrs_doc_to_mapped_layout_blocks};
    use crate::{EditCtx, FormatPolicy, KIND_KEY, Position, StoryRange, story_ref};
    use docx_layout::types::{LayoutBlock, Run};
    use serde_json::{Value, json};
    use std::collections::BTreeSet;
    use yrs::{MapPrelim, Text, Transact};

    fn story(doc: &EditingDoc, id: &str, segments: Value) {
        doc.create_story(id, "", "Normal", "left").unwrap();
        let mut txn = doc.doc.transact_mut();
        let story = story_ref(&txn, id).unwrap();
        story.remove_range(&mut txn, 0, 1);
        let mut offset = 0;
        for value in segments.as_array().unwrap() {
            if let Some(text) = value.get("text").and_then(Value::as_str) {
                story.insert(&mut txn, offset, text);
                offset += text.encode_utf16().count() as u32;
            } else {
                let payload = value.as_object().unwrap().iter().map(|(key, value)| {
                    (
                        if key == "kind" {
                            KIND_KEY.to_owned()
                        } else {
                            key.clone()
                        },
                        Any::from_json(&value.to_string()).unwrap(),
                    )
                });
                story.insert_embed(&mut txn, offset, MapPrelim::from_iter(payload));
                offset += 1;
            }
        }
    }

    #[test]
    fn resident_queries_match_rebuilds_after_insertions_and_deletions() {
        let doc = EditingDoc::new(88001);
        story(
            &doc,
            "body:t0:r0c0",
            json!([
                {"text":"Cell"}, {"kind":"pilcrow", "paraId":"cell"}
            ]),
        );
        story(
            &doc,
            "body",
            json!([
                {"text":"First"}, {"kind":"pilcrow", "paraId":"first"},
                {"kind":"table", "rows":[{"cells":[{"story":"body:t0:r0c0"}]}]},
                {"kind":"pilcrow", "paraId":"table"},
                {"text":"Last"}, {"kind":"pilcrow", "paraId":"last"}
            ]),
        );
        let mut index = ResidentGeometryIndex::build(&doc, "body", doc.committed_epoch()).unwrap();
        for (story, para, offset, units, insertion, story_offset) in [
            ("body", "first", 5, 2, true, 5),
            ("body:t0:r0c0", "cell", 0, 3, true, 0),
            ("body", "first", 4, 3, false, 1),
            ("body:t0:r0c0", "cell", 7, 2, false, 5),
            ("body", "first", 4, 4, false, 0),
            ("body", "first", 0, 1, true, 0),
        ] {
            assert!(index.move_text(
                story,
                para,
                offset,
                if insertion { units } else { -units },
                insertion
            ));
            let mut txn = doc.doc.transact_mut();
            let text = story_ref(&txn, story).unwrap();
            if insertion {
                text.insert(&mut txn, story_offset, &"+".repeat(units as usize));
            } else {
                text.remove_range(&mut txn, story_offset, units as u32);
            }
            drop(txn);
            let rebuilt =
                ResidentGeometryIndex::build(&doc, "body", doc.committed_epoch()).unwrap();
            for ranges in [
                vec![[0, 8]],
                vec![[8, 20]],
                vec![[0, 3], [22, 40]],
                vec![[0, 100]],
            ] {
                let request = ResidentRequest {
                    ranges,
                    selection: Some([
                        ResidentLoc {
                            story: story.into(),
                            para_id: para.into(),
                            offset: 0,
                        },
                        ResidentLoc {
                            story: story.into(),
                            para_id: para.into(),
                            offset: 3,
                        },
                    ]),
                };
                assert_eq!(
                    serde_json::to_value(index.query(&request)).unwrap(),
                    serde_json::to_value(rebuilt.query(&request)).unwrap()
                );
            }
        }
    }

    fn package(body: &str) -> Vec<u8> {
        let parts = [
            (
                "[Content_Types].xml",
                r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>"#.to_owned(),
            ),
            (
                "_rels/.rels",
                r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"#.to_owned(),
            ),
            (
                "word/document.xml",
                format!(
                    r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:body>{body}<w:sectPr/></w:body></w:document>"#
                ),
            ),
        ];
        ooxml_opc::rezip_parts(
            &parts
                .into_iter()
                .map(|(name, xml)| (name.to_owned(), xml.into_bytes()))
                .collect::<Vec<_>>(),
        )
        .unwrap()
    }

    fn controls_doc() -> EditingDoc {
        let p = |id: &str, content: String| format!(r#"<w:p w14:paraId="{id}">{content}</w:p>"#);
        let r = |text: &str| format!("<w:r><w:t>{text}</w:t></w:r>");
        let sdt = |properties: &str, content: String| {
            format!(
                "<w:sdt><w:sdtPr>{properties}</w:sdtPr><w:sdtContent>{content}</w:sdtContent></w:sdt>"
            )
        };
        let checkbox = |id: u32| {
            sdt(
                &format!(
                    r#"<w:id w:val="{id}"/><w14:checkbox><w14:checked w14:val="0"/></w14:checkbox>"#
                ),
                r("☐"),
            )
        };
        let body = [
            p("00000001", checkbox(1) + &r("A")),
            p("00000002", checkbox(2) + &r("B")),
            sdt(
                r#"<w:id w:val="3"/><w:tag w:val="block"/>"#,
                p("00000003", r("Block")),
            ),
            p(
                "00000004",
                r("Text")
                    + &sdt(
                        r#"<w:id w:val="4"/><w:tag w:val="name"/><w:text/>"#,
                        r("Name"),
                    )
                    + &r("End"),
            ),
            format!(
                r#"<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid><w:tr><w:tc>{}</w:tc></w:tr></w:tbl>"#,
                p("00000005", checkbox(5) + &r("C"))
            ),
            p("00000006", r("Tail")),
        ]
        .concat();
        let doc = EditingDoc::new(88002);
        crate::seed_from_docx(&doc, &package(&body)).unwrap();
        doc
    }

    type Paragraphs = BTreeSet<(String, String, u64, u64)>;
    type Inline = BTreeSet<(String, u64, u64)>;
    type Blocks = BTreeSet<(String, u64, u64)>;

    /// Paragraph content ranges, inline control nodes by paragraph and block control nodes by
    /// content story, as the layout bridge places them.
    fn layout_positions(doc: &EditingDoc) -> (Paragraphs, Inline, Blocks) {
        fn visit(blocks: &[LayoutBlock], ends: &mut HashMap<u64, (u64, Option<String>)>) {
            for block in blocks {
                match block {
                    LayoutBlock::Paragraph(paragraph) => {
                        let group = paragraph
                            .sdt_groups
                            .as_ref()
                            .and_then(|groups| groups.first());
                        ends.insert(
                            paragraph.pm_start.unwrap() as u64,
                            (
                                paragraph.pm_end.unwrap() as u64,
                                group.map(|group| group.id.clone()),
                            ),
                        );
                    }
                    LayoutBlock::Table(table) => {
                        for cell in table.rows.iter().flat_map(|row| &row.cells) {
                            visit(&cell.blocks, ends);
                        }
                    }
                    _ => {}
                }
            }
        }
        let (blocks, map) =
            yrs_doc_to_mapped_layout_blocks(doc, "body", &RenderEnv::default()).unwrap();
        let mut ends = HashMap::new();
        visit(&blocks, &mut ends);
        let mut paragraphs = BTreeSet::new();
        let mut groups: BTreeMap<String, (String, u64)> = BTreeMap::new();
        for &(start, source) in &map.paragraph_blocks {
            let (slot, para_id) = &map.paragraphs[source as usize];
            let story = map.stories[*slot as usize].clone();
            let (end, group) = ends[&start].clone();
            paragraphs.insert((story.clone(), para_id.clone(), start + 1, end - 1));
            if let Some(group) = group {
                let entry = groups.entry(group).or_insert((story, end));
                entry.1 = entry.1.max(end);
            }
        }
        let blocks = groups
            .into_iter()
            .map(|(group, (story, end))| {
                let from = group.strip_prefix("sdt@").unwrap().parse().unwrap();
                (story, from, end + 1)
            })
            .collect();
        let mut atoms: BTreeMap<(u32, u32), (u64, u64)> = BTreeMap::new();
        for span in map.spans.iter().filter(|span| span.atom) {
            let range = atoms
                .entry((span.paragraph, span.raw_start))
                .or_insert((span.pm_start, span.pm_end));
            range.0 = range.0.min(span.pm_start);
            range.1 = range.1.max(span.pm_end);
        }
        let inline = atoms
            .into_iter()
            .map(|((paragraph, _), (start, end))| {
                (
                    map.paragraphs[paragraph as usize].1.clone(),
                    start - 1,
                    end + 1,
                )
            })
            .collect();
        (paragraphs, inline, blocks)
    }

    fn resident_positions(value: &Value) -> (Paragraphs, Inline, Blocks) {
        let number = |value: &Value| value.as_u64().unwrap();
        let text = |value: &Value| value.as_str().unwrap().to_owned();
        let mut paragraphs = BTreeSet::new();
        let mut inline = BTreeSet::new();
        for paragraph in value["paragraphs"].as_array().unwrap() {
            paragraphs.insert((
                text(&paragraph["story"]),
                text(&paragraph["paraId"]),
                number(&paragraph["from"]),
                number(&paragraph["to"]),
            ));
            for control in paragraph["controls"].as_array().into_iter().flatten() {
                inline.insert((
                    text(&paragraph["paraId"]),
                    number(&control["from"]),
                    number(&control["to"]),
                ));
            }
        }
        let blocks = value["controls"]
            .as_array()
            .into_iter()
            .flatten()
            .map(|control| {
                (
                    text(&control["contentStory"]),
                    number(&control["from"]),
                    number(&control["to"]),
                )
            })
            .collect();
        (paragraphs, inline, blocks)
    }

    fn outline_paragraphs(doc: &EditingDoc) -> Paragraphs {
        let GeometryRead::Value(outline) = doc.geometry_position_outline("body") else {
            panic!("legacy outline");
        };
        outline
            .iter()
            .flat_map(|(story, projected)| {
                projected.paragraphs.iter().map(move |paragraph| {
                    let from = projected.content_start + paragraph.display_start + 1;
                    (
                        story.clone(),
                        paragraph.para_id.clone(),
                        from,
                        from + paragraph.length,
                    )
                })
            })
            .collect()
    }

    fn whole(selection: Option<[ResidentLoc; 2]>) -> ResidentRequest {
        ResidentRequest {
            ranges: vec![[0, 10_000]],
            selection,
        }
    }

    fn at(story: &str, para_id: &str, offset: u64) -> ResidentLoc {
        ResidentLoc {
            story: story.into(),
            para_id: para_id.into(),
            offset,
        }
    }

    #[test]
    fn resident_and_outline_positions_match_the_layout_after_inline_and_block_controls() {
        let doc = controls_doc();
        let layout = layout_positions(&doc);
        assert_eq!(layout.0.len(), 6);
        assert_eq!(layout.1.len(), 4);
        assert_eq!(layout.2.len(), 1);
        assert_eq!(outline_paragraphs(&doc), layout.0);
        let resident = doc
            .resident_geometry(
                "body",
                &whole(Some([at("body", "00000004", 6), at("body", "00000004", 9)])),
            )
            .unwrap();
        let value = serde_json::to_value(&resident).unwrap();
        assert_eq!(resident_positions(&value), layout);
        let (blocks, _) =
            yrs_doc_to_mapped_layout_blocks(&doc, "body", &RenderEnv::default()).unwrap();
        let end = blocks
            .iter()
            .find_map(|block| match block {
                LayoutBlock::Paragraph(paragraph) => paragraph.runs.iter().find_map(|run| {
                    matches!(run, Run::Text(text) if text.text == "End")
                        .then(|| [run.pm_start().unwrap() as u64, run.pm_end().unwrap() as u64])
                }),
                _ => None,
            })
            .unwrap();
        assert_eq!(value["selection"], json!(end));
        assert_eq!(value["controls"][0]["story"], "body");
    }

    #[test]
    fn resident_geometry_after_inline_controls_follows_edits_like_a_rebuild() {
        let doc = controls_doc();
        let request = whole(Some([at("body", "00000004", 0), at("body", "00000004", 9)]));
        doc.resident_geometry("body", &request).unwrap();
        let ctx = EditCtx::local("", "");
        let raw = |story: &str, para_id: &str, offset: u32| {
            doc.paragraph_index(story)
                .unwrap()
                .para_span(para_id)
                .unwrap()
                .0
                + offset
        };
        let check = |label: &str, incremental: bool| {
            let cold = serde_json::to_value(
                ResidentGeometryIndex::build(&doc, "body", doc.committed_epoch())
                    .unwrap()
                    .query(&request),
            )
            .unwrap();
            let current = if incremental {
                let revision = doc.story_revisions.lock().unwrap().geometry;
                let cache = doc.resident_geometry.lock().unwrap();
                let index = cache.get("body").expect(label);
                assert_eq!(index.revision, revision, "{label}");
                serde_json::to_value(index.query(&request)).unwrap()
            } else {
                serde_json::to_value(doc.resident_geometry("body", &request).unwrap()).unwrap()
            };
            assert_eq!(current, cold, "{label}");
            assert_eq!(resident_positions(&cold), layout_positions(&doc), "{label}");
        };
        for (label, para_id, offset, text) in [
            ("after the inline control", "00000004", 6, "+"),
            ("right before the inline control", "00000004", 5, "x"),
            ("before a leading checkbox", "00000001", 0, "y"),
            ("after a cell checkbox", "00000005", 1, "z"),
            ("inside the block control", "00000003", 5, "w"),
        ] {
            let story = outline_paragraphs(&doc)
                .into_iter()
                .find(|paragraph| paragraph.1 == para_id)
                .unwrap()
                .0;
            let index = raw(&story, para_id, offset);
            let before = doc.committed_epoch();
            doc.insert_text(
                &ctx,
                Position::new(story.as_str(), index),
                text,
                FormatPolicy::Inherit,
            )
            .unwrap();
            doc.advance_indexes_after_text_insert(
                &story,
                before,
                doc.committed_epoch(),
                index,
                text,
            );
            check(label, true);
        }
        let start = raw("body", "00000004", 7);
        let before = doc.committed_epoch();
        doc.delete_range(&ctx, StoryRange::new("body", start, start + 1))
            .unwrap();
        doc.advance_indexes_after_text_delete(
            "body",
            before,
            doc.committed_epoch(),
            start,
            start + 1,
        );
        check("deleting after the inline control", true);
        let fill: crate::EditRequest = serde_json::from_value(json!({
            "expectVersion": doc.version().as_str(),
            "steps": [{"op": "setContentControlText", "target": {"kind": "tag", "tag": "name"}, "text": "Longer name"}]
        }))
        .unwrap();
        doc.apply_edits(&fill, &crate::UndoSession::new())
            .unwrap()
            .unwrap();
        check("typing inside the inline control", false);
    }
}
