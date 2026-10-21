use std::collections::{BTreeMap, BTreeSet};
use std::ops::Range;

use serde_json::Value;

use crate::paragraph_identity::{attribute, paragraph_occurrences, tags};

use super::context::RecordedParagraphs;
use super::paragraph_ids::S13TablePatch;
use super::s13::element_span;
use super::splice::recorded_entries;

#[derive(Default)]
pub(super) struct TableEdits {
    pub edits: Vec<(Range<usize>, u8, String)>,
    pub paragraphs: BTreeSet<u32>,
    pub inserted: BTreeSet<usize>,
    pub reasons: Vec<String>,
    pub tables: BTreeSet<usize>,
}

struct Child<'a> {
    name: &'a str,
    span: Range<usize>,
}

fn children(xml: &str) -> Option<Vec<Child<'_>>> {
    let mut result = Vec::new();
    let mut depth = 0usize;
    for tag in tags(xml)? {
        if tag.end {
            depth = depth.checked_sub(1)?;
        } else {
            if depth == 1 {
                let span = element_span(&xml[tag.range.start..], &[])?;
                result.push(Child {
                    name: tag.name,
                    span: tag.range.start + span.start..tag.range.start + span.end,
                });
            }
            depth += usize::from(!tag.empty);
        }
    }
    Some(result)
}

fn named(xml: &str, name: &str) -> Option<Vec<Range<usize>>> {
    Some(
        children(xml)?
            .into_iter()
            .filter(|child| child.name == name)
            .map(|child| child.span)
            .collect(),
    )
}

fn tables(xml: &str) -> Option<Vec<Range<usize>>> {
    let mut result = Vec::new();
    for tag in tags(xml)? {
        if tag.name == "w:tbl" && !tag.end {
            let span = element_span(&xml[tag.range.start..], &[])?;
            result.push(tag.range.start..tag.range.start + span.end);
        }
    }
    Some(result)
}

fn rank(name: &str) -> usize {
    [
        "w:cnfStyle",
        "w:tcW",
        "w:gridSpan",
        "w:hMerge",
        "w:vMerge",
        "w:tcBorders",
        "w:shd",
        "w:noWrap",
        "w:tcMar",
        "w:textDirection",
        "w:tcFitText",
        "w:vAlign",
        "w:hideMark",
        "w:headers",
        "w:cellIns",
        "w:cellDel",
        "w:cellMerge",
        "w:tcPrChange",
    ]
    .iter()
    .position(|candidate| *candidate == name)
    .unwrap_or(usize::MAX)
}

fn property(xml: &str, name: &str) -> Option<Option<String>> {
    let properties = named(xml, "w:tcPr")?;
    let Some(span) = properties.first() else {
        return Some(None);
    };
    let properties = &xml[span.clone()];
    let children = named(properties, name)?;
    if children.len() > 1 {
        return None;
    }
    Some(
        children
            .first()
            .map(|span| properties[span.clone()].to_owned()),
    )
}

fn property_value(xml: Option<&str>, name: &str) -> Option<String> {
    let Some(xml) = xml else {
        return Some(if name == "w:gridSpan" { "1" } else { "none" }.to_owned());
    };
    let tag = tags(xml)?.into_iter().next()?;
    let value = attribute(&tag, "w:val").map_or("continue", |range| &xml[range]);
    if name == "w:gridSpan" {
        Some(value.parse::<u32>().ok()?.to_string())
    } else {
        Some(value.to_owned())
    }
}

fn patch_child(
    xml: &str,
    name: &str,
    replacement: Option<&str>,
    rank: fn(&str) -> usize,
) -> Option<String> {
    let parts = children(xml)?;
    let mut output = xml.to_owned();
    if let Some(child) = parts.iter().find(|child| child.name == name) {
        output.replace_range(child.span.clone(), replacement.unwrap_or_default());
    } else if let Some(replacement) = replacement {
        let scanned = tags(xml)?;
        let root = scanned.first()?;
        if root.empty {
            let opening = xml.strip_suffix("/>")?;
            output = format!("{opening}>{replacement}</{}>", root.name);
        } else {
            let at = parts
                .iter()
                .find(|child| rank(child.name) > rank(name))
                .map_or(scanned.last()?.range.start, |child| child.span.start);
            output.insert_str(at, replacement);
        }
    }
    Some(output)
}

fn patch_property(cell: &str, name: &str, replacement: Option<&str>) -> Option<String> {
    let properties = named(cell, "w:tcPr")?;
    let mut output = cell.to_owned();
    if let Some(span) = properties.first() {
        let patched = patch_child(&cell[span.clone()], name, replacement, rank)?;
        output.replace_range(span.clone(), &patched);
    } else if let Some(replacement) = replacement {
        let at = tags(cell)?.first()?.range.end;
        output.insert_str(at, &format!("<w:tcPr>{replacement}</w:tcPr>"));
    }
    Some(output)
}

fn merge_properties(source: &str, target: &str) -> Option<String> {
    let mut result = source.to_owned();
    for name in ["w:gridSpan", "w:vMerge"] {
        let old = property(&result, name)?;
        let new = property(target, name)?;
        if property_value(old.as_deref(), name)? != property_value(new.as_deref(), name)? {
            result = patch_property(&result, name, new.as_deref())?;
        }
    }
    Some(result)
}

fn stories(row: &Value) -> Option<Vec<&str>> {
    row["cells"]
        .as_array()?
        .iter()
        .map(|cell| cell["story"].as_str())
        .collect()
}

fn row_sources(before: &[Value], after: &[Value]) -> Option<Vec<Option<usize>>> {
    let old = before.iter().map(stories).collect::<Option<Vec<_>>>()?;
    let mut used = BTreeSet::new();
    let mut result = Vec::new();
    for (at, row) in after.iter().enumerate() {
        let keys = stories(row)?;
        let source = old
            .iter()
            .position(|previous| keys.iter().any(|key| previous.contains(key)))
            .or_else(|| (before.len() == after.len()).then_some(at));
        if let Some(source) = source
            && !used.insert(source)
        {
            return None;
        }
        result.push(source);
    }
    if result
        .iter()
        .flatten()
        .copied()
        .collect::<Vec<_>>()
        .windows(2)
        .any(|pair| pair[0] >= pair[1])
    {
        return None;
    }
    Some(result)
}

fn changed_key(before: &Value, after: &Value, ignored: &[&str]) -> Option<String> {
    let old = before.as_object()?;
    let new = after.as_object()?;
    old.keys()
        .chain(new.keys())
        .find(|key| !ignored.contains(&key.as_str()) && old.get(*key) != new.get(*key))
        .cloned()
}

fn row_reason(before: &Value, after: &Value) -> Option<String> {
    if let Some(key) = changed_key(&before["trPr"], &after["trPr"], &[]) {
        return Some(format!("table:{key}"));
    }
    for cell in after["cells"].as_array()? {
        let Some(old) = before["cells"]
            .as_array()?
            .iter()
            .find(|old| old["story"] == cell["story"])
        else {
            return Some("table:cells".to_owned());
        };
        if let Some(key) = changed_key(
            &old["tcPr"],
            &cell["tcPr"],
            &["colspan", "rowspan", "gridSpan", "vMerge"],
        ) {
            return Some(format!("table:{key}"));
        }
    }
    None
}

struct Paragraphs<'a> {
    source: &'a str,
    target: &'a str,
    recorded: &'a RecordedParagraphs,
    rendered: &'a BTreeMap<u32, String>,
    target_entries: BTreeMap<usize, Option<usize>>,
    insertion_indices: BTreeMap<usize, usize>,
    source_ordinals: BTreeMap<usize, u32>,
}

impl Paragraphs<'_> {
    fn cell(
        &self,
        source: &str,
        source_at: usize,
        target: &str,
        target_at: usize,
    ) -> Option<String> {
        let old = named(source, "w:p")?;
        let new = named(target, "w:p")?;
        if children(source)?
            .iter()
            .any(|child| !matches!(child.name, "w:tcPr" | "w:p"))
            || children(target)?
                .iter()
                .any(|child| !matches!(child.name, "w:tcPr" | "w:p"))
        {
            return None;
        }
        let mut paragraphs = Vec::new();
        for span in &new {
            let entry = self.target_entries.get(&(target_at + span.start))?;
            let value = match entry {
                Some(entry) => match self.recorded.order[*entry] {
                    Some(ordinal) => self.rendered.get(&ordinal)?.clone(),
                    None => self.recorded.inserted[*self.insertion_indices.get(entry)?].clone(),
                },
                None => {
                    let old = old.get(paragraphs.len())?;
                    let moved = self
                        .source_ordinals
                        .get(&(source_at + old.start))
                        .is_some_and(|ordinal| self.recorded.written.contains_key(ordinal));
                    if moved {
                        target[span.clone()].to_owned()
                    } else {
                        source[old.clone()].to_owned()
                    }
                }
            };
            paragraphs.push(value);
        }
        let mut result = source.to_owned();
        if old.len() > paragraphs.len() {
            return None;
        }
        for (span, paragraph) in old.iter().zip(&paragraphs).rev() {
            result.replace_range(span.clone(), paragraph);
        }
        if paragraphs.len() > old.len() {
            let at = old.last()?.end;
            let shift: isize = old
                .iter()
                .zip(&paragraphs)
                .map(|(span, paragraph)| paragraph.len() as isize - span.len() as isize)
                .sum();
            result.insert_str(at.checked_add_signed(shift)?, &paragraphs[old.len()..].concat());
        }
        merge_properties(&result, target)
    }

    fn row(&self, source: Range<usize>, target: Range<usize>) -> Option<String> {
        let old = &self.source[source.clone()];
        let new = &self.target[target.clone()];
        let old_cells = named(old, "w:tc")?;
        let new_cells = named(new, "w:tc")?;
        let mut used = BTreeSet::new();
        let mut replacements = BTreeMap::new();
        for (at, target_cell) in new_cells.iter().enumerate() {
            let target_xml = &new[target_cell.clone()];
            let first = paragraph_occurrences(target_xml)?.first().cloned();
            let ordinal = first.and_then(|paragraph| {
                let entry = self
                    .target_entries
                    .get(&(target.start + target_cell.start + paragraph.tag.start))?
                    .as_ref()?;
                self.recorded.order[*entry]
            });
            let donor = ordinal
                .and_then(|ordinal| {
                    old_cells.iter().position(|cell| {
                        self.source_ordinals.iter().any(|(start, value)| {
                            *value == ordinal
                                && (source.start + cell.start..source.start + cell.end)
                                    .contains(start)
                        })
                    })
                })
                .unwrap_or(at);
            let cell = old_cells.get(donor)?;
            if !used.insert(donor) {
                return None;
            }
            replacements.insert(
                donor,
                self.cell(
                    &old[cell.clone()],
                    source.start + cell.start,
                    target_xml,
                    target.start + target_cell.start,
                )?,
            );
        }
        let mut result = old.to_owned();
        for (at, cell) in old_cells.iter().enumerate().rev() {
            result.replace_range(cell.clone(), replacements.get(&at).map_or("", String::as_str));
        }
        Some(result)
    }
}

fn same_merge_shape(source: &str, target: &str) -> Option<bool> {
    let old = named(source, "w:tc")?;
    let new = named(target, "w:tc")?;
    if old.len() != new.len() {
        return Some(false);
    }
    for (old, new) in old.iter().zip(&new) {
        for name in ["w:gridSpan", "w:vMerge"] {
            let old = property(&source[old.clone()], name)?;
            let new = property(&target[new.clone()], name)?;
            if property_value(old.as_deref(), name)? != property_value(new.as_deref(), name)? {
                return Some(false);
            }
        }
    }
    Some(true)
}

fn cover_insertions(result: &mut TableEdits, paragraphs: &Paragraphs<'_>, span: Range<usize>) {
    for (_, entry) in paragraphs.target_entries.range(span) {
        if let Some(entry) = entry
            && let Some(index) = paragraphs.insertion_indices.get(entry)
        {
            result.inserted.insert(*index);
        }
    }
}

fn table_rank(name: &str) -> usize {
    [
        "w:tblStyle",
        "w:tblpPr",
        "w:tblOverlap",
        "w:bidiVisual",
        "w:tblStyleRowBandSize",
        "w:tblStyleColBandSize",
        "w:tblW",
        "w:jc",
        "w:tblCellSpacing",
        "w:tblInd",
        "w:tblBorders",
        "w:shd",
        "w:tblLayout",
        "w:tblCellMar",
        "w:tblLook",
        "w:tblCaption",
        "w:tblDescription",
        "w:tblPrChange",
    ]
    .iter()
    .position(|candidate| *candidate == name)
    .unwrap_or(usize::MAX)
}

fn table_property(key: &str) -> Option<&'static str> {
    Some(match key {
        "styleId" => "w:tblStyle",
        "floating" => "w:tblpPr",
        "overlap" => "w:tblOverlap",
        "bidi" => "w:bidiVisual",
        "width" | "widthType" => "w:tblW",
        "justification" => "w:jc",
        "cellSpacing" => "w:tblCellSpacing",
        "indent" => "w:tblInd",
        "borders" => "w:tblBorders",
        "shading" => "w:shd",
        "tableLayout" | "layout" => "w:tblLayout",
        "cellMargins" => "w:tblCellMar",
        "look" => "w:tblLook",
        "tblPrChange" => "w:tblPrChange",
        _ => return None,
    })
}

fn metadata_edits(
    source: &str,
    target: &str,
    offset: usize,
    patch: &S13TablePatch,
) -> Option<Vec<(Range<usize>, u8, String)>> {
    let mut edits = Vec::new();
    if patch.before["grid"] != patch.after["grid"] {
        let old = named(source, "w:tblGrid")?;
        let new = named(target, "w:tblGrid")?;
        let replacement = new.first().map_or("", |span| &target[span.clone()]);
        let span = old.first().cloned().unwrap_or_else(|| {
            let at = named(source, "w:tr")
                .and_then(|rows| rows.first().cloned())
                .map_or(source.len(), |row| row.start);
            at..at
        });
        edits.push((
            offset + span.start..offset + span.end,
            0,
            replacement.to_owned(),
        ));
    }
    let before = patch.before["tblPr"].as_object()?;
    let after = patch.after["tblPr"].as_object()?;
    let mut properties = BTreeSet::new();
    for key in before.keys().chain(after.keys()) {
        if before.get(key) == after.get(key) {
            continue;
        }
        if key == "_originalFormatting" {
            let empty = serde_json::Map::new();
            let old = before.get(key).and_then(Value::as_object).unwrap_or(&empty);
            let new = after.get(key).and_then(Value::as_object).unwrap_or(&empty);
            for key in old.keys().chain(new.keys()) {
                if old.get(key) != new.get(key)
                    && let Some(name) = table_property(key)
                {
                    properties.insert(name);
                }
            }
        } else if let Some(name) = table_property(key) {
            properties.insert(name);
        }
    }
    if !properties.is_empty() {
        let old = named(source, "w:tblPr")?;
        let new = named(target, "w:tblPr")?;
        let original = old.first().map_or("<w:tblPr/>", |span| &source[span.clone()]);
        let written = new.first().map_or("<w:tblPr/>", |span| &target[span.clone()]);
        let mut patched = original.to_owned();
        let mut properties: Vec<_> = properties.into_iter().collect();
        properties.sort_by_key(|name| table_rank(name));
        for name in properties {
            let child = named(written, name)?;
            let replacement = child.first().map(|span| &written[span.clone()]);
            patched = patch_child(&patched, name, replacement, table_rank)?;
        }
        let span = old.first().cloned().unwrap_or_else(|| {
            let at = tags(source)
                .and_then(|tags| tags.first().map(|tag| tag.range.end))
                .unwrap_or_default();
            at..at
        });
        if patched != original {
            edits.push((offset + span.start..offset + span.end, 0, patched));
        }
    }
    Some(edits)
}

pub(super) fn patch_tables(
    source: &str,
    target: &str,
    recorded: &RecordedParagraphs,
    patches: &[S13TablePatch],
    rendered: &BTreeMap<u32, String>,
) -> Option<TableEdits> {
    let mut result = TableEdits::default();
    if patches.is_empty() {
        return Some(result);
    }
    let source_tables = tables(source)?;
    let target_tables = tables(target)?;
    if source_tables.len() != target_tables.len() {
        return None;
    }
    let source_paragraphs = paragraph_occurrences(source)?;
    let target_paragraphs = paragraph_occurrences(target)?;
    let entries = recorded_entries(target, recorded)?;
    let mut next = 0;
    let insertion_indices = recorded
        .order
        .iter()
        .enumerate()
        .filter(|(_, ordinal)| ordinal.is_none())
        .map(|(at, _)| {
            let index = next;
            next += 1;
            (at, index)
        })
        .collect();
    let paragraphs = Paragraphs {
        source,
        target,
        recorded,
        rendered,
        insertion_indices,
        target_entries: target_paragraphs
            .iter()
            .zip(entries)
            .map(|(paragraph, entry)| (paragraph.tag.start, entry))
            .collect(),
        source_ordinals: source_paragraphs
            .iter()
            .map(|paragraph| (paragraph.tag.start, paragraph.ordinal))
            .collect(),
    };
    for patch in patches {
        let occurrence = source_paragraphs.get(patch.ordinal as usize)?;
        let table = source_tables
            .iter()
            .enumerate()
            .filter(|(_, span)| span.contains(&occurrence.tag.start))
            .min_by_key(|(_, span)| span.len())?
            .0;
        if !result.tables.insert(table) {
            return None;
        }
        let source_span = &source_tables[table];
        let target_span = &target_tables[table];
        let old_xml = &source[source_span.clone()];
        let new_xml = &target[target_span.clone()];
        let old_rows = named(old_xml, "w:tr")?;
        let new_rows = named(new_xml, "w:tr")?;
        let before = patch.before["rows"].as_array()?;
        let after = patch.after["rows"].as_array()?;
        if old_rows.len() != before.len() || new_rows.len() != after.len() {
            return None;
        }
        let origins = row_sources(before, after)?;
        let metadata = metadata_edits(old_xml, new_xml, source_span.start, patch)?;
        result.edits.extend(metadata);
        let table_reason = changed_key(&patch.before, &patch.after, &["rows"])
            .map(|key| format!("table:{key}"));
        for (at, row) in new_rows.iter().enumerate() {
            let target_row = target_span.start + row.start..target_span.start + row.end;
            if let Some(origin) = origins[at] {
                let source_row = source_span.start + old_rows[origin].start
                    ..source_span.start + old_rows[origin].end;
                let reason = row_reason(&before[origin], &after[at])
                    .or_else(|| table_reason.clone().filter(|_| at == 0));
                if before[origin] == after[at]
                    && reason.is_none()
                    && same_merge_shape(&source[source_row.clone()], &target[target_row.clone()])?
                {
                    continue;
                }
                for paragraph in &source_paragraphs {
                    if source_row.contains(&paragraph.tag.start) {
                        result.paragraphs.insert(paragraph.ordinal);
                    }
                }
                cover_insertions(&mut result, &paragraphs, target_row.clone());
                let patched = if reason.is_none() {
                    paragraphs.row(source_row.clone(), target_row.clone())
                } else {
                    None
                };
                let xml = match patched {
                    Some(xml) => xml,
                    None => {
                        result
                            .reasons
                            .push(reason.unwrap_or_else(|| "table:structure".to_owned()));
                        target[target_row].to_owned()
                    }
                };
                if xml != source[source_row.clone()] {
                    result.edits.push((source_row, 2, xml));
                }
            } else {
                let donor = origins[..at]
                    .iter()
                    .rev()
                    .flatten()
                    .next()
                    .copied()
                    .or_else(|| origins[at + 1..].iter().flatten().next().copied())?;
                let source_row = source_span.start + old_rows[donor].start
                    ..source_span.start + old_rows[donor].end;
                let position = if origins[..at].iter().any(Option::is_some) {
                    source_row.end
                } else {
                    source_row.start
                };
                cover_insertions(&mut result, &paragraphs, target_row.clone());
                let xml = match paragraphs.row(source_row, target_row.clone()) {
                    Some(xml) => xml,
                    None => {
                        result.reasons.push("table:structure".to_owned());
                        target[target_row].to_owned()
                    }
                };
                result.edits.push((position..position, 1, xml));
            }
        }
        for (at, row) in old_rows.iter().enumerate() {
            if !origins.contains(&Some(at)) {
                let span = source_span.start + row.start..source_span.start + row.end;
                for paragraph in &source_paragraphs {
                    if span.contains(&paragraph.tag.start) {
                        result.paragraphs.insert(paragraph.ordinal);
                    }
                }
                result.edits.push((span, 2, String::new()));
            }
        }
    }
    Some(result)
}
