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
    pub length: u64,
    pub story_order: usize,
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
    pub selection: Option<[u64; 2]>,
}

#[derive(Clone)]
enum Item {
    Paragraph(ResidentParagraph),
    Table(usize),
    Cell(usize, ResidentCell),
}

struct Entry {
    from: usize,
    to: usize,
    item: Item,
}

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
        let mut records = Vec::new();
        let mut tables = Vec::new();
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
                        length: paragraph.length + paragraph.leading,
                        story_order: 0,
                    }),
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
                                records.push((cell_from, cell_to, Item::Cell(id, cell.clone())));
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
                        records.push((start, end, Item::Table(id)));
                        cursor = end;
                        leading += 1;
                        table_index += 1;
                    } else if kind == "blockSdt" {
                        let story = child(payload.get("story"), format!("{story}:sdt0"));
                        cursor += outline
                            .get(&story)
                            .ok_or_else(|| "missing control projection".to_owned())?
                            .size
                            + 2;
                        leading += 1;
                    } else if kind == "pageBreak" || kind == "columnBreak" {
                        cursor += 1;
                        leading += 1;
                    }
                }
                cursor = projected.content_start + paragraph.display_start + paragraph.length + 2;
            }
        }
        let mut order: Vec<_> = (0..tables.len()).collect();
        order.sort_by(|a, b| {
            if a == b {
                return std::cmp::Ordering::Equal;
            }
            let a = &tables[*a];
            let b = &tables[*b];
            if a.from <= b.from && a.to >= b.to {
                std::cmp::Ordering::Greater
            } else if b.from <= a.from && b.to >= a.to {
                std::cmp::Ordering::Less
            } else {
                a.from.cmp(&b.from)
            }
        });
        let mut story_order = HashMap::from([(root.to_owned(), 0)]);
        for id in order {
            for cell in &tables[id].cells {
                let next = story_order.len();
                story_order.entry(cell.story.clone()).or_insert(next);
            }
        }
        let mut keys = BTreeMap::new();
        for (from, to, _) in &records {
            keys.insert((*from, false), 0);
            keys.insert((*to, true), 0);
        }
        let points: Vec<_> = keys.keys().copied().collect();
        for (index, key) in points.iter().enumerate() {
            keys.insert(*key, index);
        }
        let mut entries: Vec<_> = records
            .into_iter()
            .map(|(from, to, mut item)| {
                if let Item::Paragraph(paragraph) = &mut item {
                    paragraph.story_order =
                        *story_order.get(&paragraph.story).unwrap_or(&usize::MAX);
                }
                Entry {
                    from: keys[&(from, false)],
                    to: keys[&(to, true)],
                    item,
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

    fn locate(&self, at: &ResidentLoc) -> Option<u64> {
        let entry = &self.entries[*self.carets.get(&(at.story.clone(), at.para_id.clone()))?];
        let Item::Paragraph(paragraph) = &entry.item else {
            return None;
        };
        let from = self.position(entry.from);
        let to = self.position(entry.to);
        Some(from + at.offset.saturating_sub(paragraph.leading).min(to - from))
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
        if u64::from(offset) > end - start {
            return false;
        }
        let at = start + u64::from(offset);
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
        for index in found {
            let entry = &self.entries[index];
            match &entry.item {
                Item::Paragraph(value) => {
                    let mut value = value.clone();
                    value.from = self.position(entry.from);
                    value.to = self.position(entry.to);
                    value.length = value.to - value.from + value.leading;
                    if value.story_order != usize::MAX {
                        paragraphs.push(value);
                    }
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
        tables.sort_by(|a, b| {
            if a.from == b.from && a.to == b.to {
                std::cmp::Ordering::Equal
            } else if a.from <= b.from && a.to >= b.to {
                std::cmp::Ordering::Greater
            } else if b.from <= a.from && b.to >= a.to {
                std::cmp::Ordering::Less
            } else {
                a.from.cmp(&b.from)
            }
        });
        let selection = request
            .selection
            .as_ref()
            .and_then(|range| Some([self.locate(&range[0])?, self.locate(&range[1])?]));
        ResidentGeometry {
            paragraphs,
            tables,
            selection,
        }
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
    use crate::{KIND_KEY, story_ref};
    use serde_json::{Value, json};
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
}
