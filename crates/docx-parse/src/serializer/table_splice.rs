use std::collections::{BTreeMap, BTreeSet};
use std::ops::Range;

use serde_json::Value;

use crate::paragraph_identity::{paragraph_occurrences, simple_table_spans, tags};

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

fn stories(row: &Value) -> Option<Vec<&str>> {
    row["cells"]
        .as_array()?
        .iter()
        .map(|cell| cell["story"].as_str())
        .collect()
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
    for (old, cell) in before["cells"]
        .as_array()?
        .iter()
        .zip(after["cells"].as_array()?)
    {
        if let Some(key) = changed_key(&old["tcPr"], &cell["tcPr"], &[]) {
            return Some(format!("table:{key}"));
        }
    }
    None
}

struct Paragraphs<'a> {
    target: &'a str,
    recorded: &'a RecordedParagraphs,
    rendered: &'a BTreeMap<u32, String>,
    target_entries: BTreeMap<usize, Option<usize>>,
    insertion_indices: BTreeMap<usize, usize>,
}

impl<'a> Paragraphs<'a> {
    fn new(
        target: &'a str,
        recorded: &'a RecordedParagraphs,
        rendered: &'a BTreeMap<u32, String>,
    ) -> Option<Self> {
        let target_paragraphs = paragraph_occurrences(target)?;
        let entries = recorded_entries(target, recorded)?;
        let insertion_indices = recorded
            .order
            .iter()
            .enumerate()
            .filter(|(_, ordinal)| ordinal.is_none())
            .enumerate()
            .map(|(index, (at, _))| (at, index))
            .collect();
        Some(Self {
            target,
            recorded,
            rendered,
            insertion_indices,
            target_entries: target_paragraphs
                .iter()
                .zip(entries)
                .map(|(paragraph, entry)| (paragraph.tag.start, entry))
                .collect(),
        })
    }

    fn replacement_row(&self, target: Range<usize>) -> Option<String> {
        let mut result = self.target[target.clone()].to_owned();
        for (start, entry) in self.target_entries.range(target.clone()).rev() {
            let Some(ordinal) = entry.and_then(|entry| self.recorded.order[entry]) else {
                continue;
            };
            let xml = self.rendered.get(&ordinal)?;
            let span = element_span(&self.target[*start..], &[])?;
            result.replace_range(*start - target.start..*start + span.end - target.start, xml);
        }
        Some(result)
    }
}

fn cover_insertions(
    result: &mut TableEdits,
    paragraphs: &Paragraphs<'_>,
    span: Range<usize>,
) -> Option<()> {
    for (_, entry) in paragraphs.target_entries.range(span) {
        if let Some(entry) = entry
            && let Some(index) = paragraphs.insertion_indices.get(entry)
            && !result.inserted.insert(*index)
        {
            return None;
        }
    }
    Some(())
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
        "styleRowBandSize" => "w:tblStyleRowBandSize",
        "styleColBandSize" => "w:tblStyleColBandSize",
        "caption" => "w:tblCaption",
        "description" => "w:tblDescription",
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
    if changed_key(&patch.before, &patch.after, &["rows", "grid", "tblPr"]).is_some() {
        return None;
    }
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
                if old.get(key) != new.get(key) && key != "styleCascade" {
                    properties.insert(table_property(key)?);
                }
            }
        } else {
            properties.insert(table_property(key)?);
        }
    }
    if !properties.is_empty() {
        let old = named(source, "w:tblPr")?;
        let new = named(target, "w:tblPr")?;
        let original = old
            .first()
            .map_or("<w:tblPr/>", |span| &source[span.clone()]);
        let written = new
            .first()
            .map_or("<w:tblPr/>", |span| &target[span.clone()]);
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

pub(super) fn patch_spans(
    source: &str,
    target: &str,
    recorded: &RecordedParagraphs,
    patches: &[S13TablePatch],
) -> Option<Vec<(Range<usize>, Range<usize>)>> {
    let mut result = Vec::new();
    if patches.is_empty() {
        return Some(result);
    }
    let source_tables = simple_table_spans(source)?;
    let target_tables = simple_table_spans(target)?;
    let source_paragraphs = paragraph_occurrences(source)?;
    let target_paragraphs = paragraph_occurrences(target)?;
    let entries = recorded_entries(target, recorded)?;
    let mut used = BTreeSet::new();
    for patch in patches {
        let occurrence = source_paragraphs.get(patch.ordinal as usize)?;
        let table = source_tables
            .iter()
            .position(|span| span.contains(&occurrence.tag.start))?;
        if !used.insert(table) {
            return None;
        }
        let source_span = &source_tables[table];
        let mut targets = BTreeSet::new();
        for (paragraph, entry) in target_paragraphs.iter().zip(&entries) {
            let Some(ordinal) = entry.and_then(|entry| recorded.order[entry]) else {
                continue;
            };
            if source_span.contains(&source_paragraphs.get(ordinal as usize)?.tag.start) {
                targets.insert(
                    target_tables
                        .iter()
                        .position(|span| span.contains(&paragraph.tag.start))?,
                );
            }
        }
        if targets.len() != 1 {
            return None;
        }
        let target_span = &target_tables[*targets.first()?];
        result.push((source_span.clone(), target_span.clone()));
    }
    Some(result)
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
    let spans = patch_spans(source, target, recorded, patches)?;
    let source_paragraphs = paragraph_occurrences(source)?;
    let paragraphs = Paragraphs::new(target, recorded, rendered)?;
    for (patch, (source_span, target_span)) in patches.iter().zip(&spans) {
        let old_xml = &source[source_span.clone()];
        let new_xml = &target[target_span.clone()];
        let old_rows = named(old_xml, "w:tr")?;
        let new_rows = named(new_xml, "w:tr")?;
        let before = patch.before["rows"].as_array()?;
        let after = patch.after["rows"].as_array()?;
        if old_rows.len() != before.len()
            || new_rows.len() != after.len()
            || before.len() != after.len()
        {
            return None;
        }
        for (at, (before, after)) in before.iter().zip(after).enumerate() {
            let old_cells = stories(before)?;
            if old_cells != stories(after)?
                || named(&old_xml[old_rows[at].clone()], "w:tc")?.len() != old_cells.len()
                || named(&new_xml[new_rows[at].clone()], "w:tc")?.len() != old_cells.len()
            {
                return None;
            }
        }
        let metadata = metadata_edits(old_xml, new_xml, source_span.start, patch)?;
        result.edits.extend(metadata);
        for (at, row) in new_rows.iter().enumerate() {
            if before[at] == after[at] {
                continue;
            }
            let target_row = target_span.start + row.start..target_span.start + row.end;
            let source_row =
                source_span.start + old_rows[at].start..source_span.start + old_rows[at].end;
            let reason = row_reason(&before[at], &after[at]);
            for paragraph in &source_paragraphs {
                if source_row.contains(&paragraph.tag.start) {
                    result.paragraphs.insert(paragraph.ordinal);
                }
            }
            cover_insertions(&mut result, &paragraphs, target_row.clone())?;
            result
                .reasons
                .push(reason.unwrap_or_else(|| "table:properties".to_owned()));
            let xml = paragraphs.replacement_row(target_row)?;
            if xml != source[source_row.clone()] {
                result.edits.push((source_row, 2, xml));
            }
        }
    }
    Some(result)
}
