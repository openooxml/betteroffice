use std::collections::{BTreeMap, BTreeSet};
use std::ops::Range;

use serde_json::Value;

use crate::paragraph_identity::{
    ParagraphOccurrence, attribute, paragraph_occurrences, patch_paragraph_ids, tags, unescaped,
};

use super::super::context::RecordedParagraphs;
use super::super::paragraph_ids::{S13SpliceAnchor, S13SplicedPart, S13TablePatch};
use super::super::s13::element_span;
use super::{
    paragraph_span, recorded_entries, required_prefixes, root_binds, root_with_prefixes,
    splice_story_part_inner, unique_ranges,
};

struct Table {
    index: usize,
    source: Range<usize>,
    target: Range<usize>,
    affected: bool,
    legacy: bool,
}

fn outer_tables(xml: &str) -> Option<Vec<Range<usize>>> {
    let mut tables = Vec::new();
    let mut end = 0;
    for tag in tags(xml)? {
        if tag.name == "w:tbl" && !tag.end && tag.range.start >= end {
            let span = element_span(&xml[tag.range.start..], &[])?;
            end = tag.range.start + span.end;
            tables.push(tag.range.start..end);
        }
    }
    Some(tables)
}

fn complex_xml(xml: &str) -> Option<bool> {
    let mut tables = 0;
    for tag in tags(xml)? {
        if tag.end {
            continue;
        }
        match tag.name {
            "w:tbl" => {
                tables += 1;
                if tables > 1 {
                    return Some(true);
                }
            }
            "w:gridBefore" | "w:gridAfter" => return Some(true),
            "w:gridSpan" => {
                let value = unescaped(xml, attribute(&tag, "w:val")?)?;
                if value.parse::<u32>().ok()? > 1 {
                    return Some(true);
                }
            }
            "w:vMerge" => {
                let value = if let Some(span) = attribute(&tag, "w:val") {
                    unescaped(xml, span)?
                } else {
                    "continue".to_owned()
                };
                if value != "restart" {
                    return Some(true);
                }
            }
            _ => {}
        }
    }
    Some(false)
}

pub(super) fn has_complex_tables(xml: &str) -> bool {
    outer_tables(xml).is_some_and(|tables| {
        tables
            .iter()
            .any(|span| complex_xml(&xml[span.clone()]) == Some(true))
    })
}

fn topology(xml: &str) -> Option<Vec<(String, bool)>> {
    Some(
        tags(xml)?
            .into_iter()
            .filter(|tag| matches!(tag.name, "w:tbl" | "w:tr" | "w:tc"))
            .map(|tag| (tag.name.to_owned(), tag.end))
            .collect(),
    )
}

fn complex_payload(value: &Value) -> bool {
    match value {
        Value::Object(object) => object.iter().any(|(key, value)| {
            matches!(key.as_str(), "gridBefore" | "gridAfter")
                || (matches!(key.as_str(), "gridSpan" | "colspan" | "rowspan")
                    && value.as_f64().is_some_and(|value| value > 1.0))
                || (key == "vMerge" && value.as_str() == Some("continue"))
                || complex_payload(value)
        }),
        Value::Array(values) => values.iter().any(complex_payload),
        _ => false,
    }
}

fn legacy_patch(patch: &S13TablePatch) -> Option<bool> {
    let mut before = patch.before.as_object()?.clone();
    let mut after = patch.after.as_object()?.clone();
    let old = before.remove("rows")?;
    let new = after.remove("rows")?;
    if before != after {
        return Some(false);
    }
    let old = old.as_array()?;
    for row in new.as_array()? {
        let story = row["cells"].as_array()?.first()?["story"].as_str()?;
        if let Some(previous) = old.iter().find(|row| row["cells"][0]["story"] == story) {
            if previous != row {
                return Some(false);
            }
        } else if complex_payload(&row["trPr"]) {
            return Some(false);
        }
    }
    Some(true)
}

struct Input<'a> {
    source: &'a str,
    target: &'a str,
    recorded: &'a RecordedParagraphs,
    part: &'a S13SplicedPart,
    source_paragraphs: Vec<ParagraphOccurrence>,
    target_paragraphs: Vec<ParagraphOccurrence>,
    entries: Vec<Option<usize>>,
    insertions: BTreeMap<usize, usize>,
}

impl<'a> Input<'a> {
    fn new(
        source: &'a str,
        target: &'a str,
        recorded: &'a RecordedParagraphs,
        part: &'a S13SplicedPart,
    ) -> Option<Self> {
        Some(Self {
            source,
            target,
            recorded,
            part,
            source_paragraphs: paragraph_occurrences(source)?,
            target_paragraphs: paragraph_occurrences(target)?,
            entries: recorded_entries(target, recorded)?,
            insertions: recorded
                .order
                .iter()
                .enumerate()
                .filter(|(_, ordinal)| ordinal.is_none())
                .enumerate()
                .map(|(index, (entry, _))| (entry, index))
                .collect(),
        })
    }

    fn tables(&self) -> Option<Vec<Table>> {
        let source = outer_tables(self.source)?;
        let target = outer_tables(self.target)?;
        if source.len() != target.len() {
            return None;
        }
        let mut tables = Vec::new();
        for (index, (source, target)) in source.into_iter().zip(target).enumerate() {
            let contains = |ordinal: u32| {
                self.source_paragraphs
                    .get(ordinal as usize)
                    .is_some_and(|paragraph| source.contains(&paragraph.tag.start))
            };
            let patches: Vec<_> = self
                .part
                .table_patches
                .iter()
                .filter(|patch| contains(patch.ordinal))
                .collect();
            if !complex_xml(&self.source[source.clone()])?
                && !complex_xml(&self.target[target.clone()])?
                && !patches.iter().any(|patch| {
                    complex_payload(&patch.before) || complex_payload(&patch.after)
                })
            {
                continue;
            }
            let explicit = !patches.is_empty()
                || self
                    .part
                    .changed
                    .iter()
                    .chain(&self.part.removed)
                    .copied()
                    .any(contains)
                || self
                    .target_paragraphs
                    .iter()
                    .enumerate()
                    .any(|(at, paragraph)| {
                        target.contains(&paragraph.tag.start)
                            && self.entries[at]
                                .is_some_and(|entry| self.recorded.order[entry].is_none())
                    });
            let structural = topology(&self.source[source.clone()])?
                != topology(&self.target[target.clone()])?;
            let affected = explicit || structural;
            let nested = tags(&self.source[source.clone()])?
                .iter()
                .filter(|tag| tag.name == "w:tbl" && !tag.end)
                .count()
                > 1;
            let legacy = (explicit || !structural)
                && (!nested || patches.is_empty())
                && patches.iter().all(|patch| legacy_patch(patch) == Some(true));
            tables.push(Table {
                index,
                source,
                target,
                affected,
                legacy,
            });
        }
        Some(tables)
    }
}

#[derive(Default)]
struct Projection {
    xml: String,
    recorded: RecordedParagraphs,
    inserted: Vec<S13SpliceAnchor>,
    target_ordinals: BTreeSet<u32>,
}

impl Projection {
    fn copy(&mut self, input: &Input<'_>, span: Range<usize>, source: bool) -> Option<()> {
        let (xml, paragraphs) = if source {
            (input.source, &input.source_paragraphs)
        } else {
            (input.target, &input.target_paragraphs)
        };
        self.xml.push_str(&xml[span.clone()]);
        let mut end = span.start;
        for (at, paragraph) in paragraphs.iter().enumerate() {
            if !span.contains(&paragraph.tag.start) || paragraph.tag.start < end {
                continue;
            }
            let range = paragraph_span(xml, paragraph)?;
            end = range.end;
            let ordinal = if source {
                Some(paragraph.ordinal)
            } else if let Some(entry) = input.entries[at] {
                input.recorded.order[entry]
            } else {
                self.recorded.placeholders.push(self.recorded.order.len());
                continue;
            };
            let value = if source {
                xml[range].to_owned()
            } else if let Some(ordinal) = ordinal {
                self.target_ordinals.insert(ordinal);
                input.recorded.written.get(&ordinal)?.clone()
            } else {
                let index = input.insertions.get(&input.entries[at]?)?;
                self.inserted.push(*input.part.inserted.get(*index)?);
                input.recorded.inserted.get(*index)?.clone()
            };
            if let Some(ordinal) = ordinal {
                if self.recorded.written.insert(ordinal, value).is_some() {
                    return None;
                }
            } else {
                self.recorded.inserted.push(value);
            }
            self.recorded.order.push(ordinal);
        }
        Some(())
    }

    fn part(&self, original: &S13SplicedPart) -> S13SplicedPart {
        let mut part = original.clone();
        part.paragraphs = self.recorded.written.keys().copied().collect();
        part.paragraphs.sort_unstable();
        part.changed
            .retain(|ordinal| self.target_ordinals.contains(ordinal));
        part.removed
            .retain(|ordinal| !self.recorded.written.contains_key(ordinal));
        part.inserted.clone_from(&self.inserted);
        part.paragraph_patches
            .retain(|patch| part.changed.contains(&patch.ordinal));
        part.legacy_paragraphs
            .retain(|patch| part.changed.contains(&patch.ordinal));
        part
    }
}

fn replacement(
    input: &Input<'_>,
    table: &Table,
    assignments: &BTreeMap<u32, String>,
    patched: &BTreeMap<u32, String>,
) -> Option<String> {
    let mut result = input.target[table.target.clone()].to_owned();
    for (at, paragraph) in input.target_paragraphs.iter().enumerate().rev() {
        if !table.target.contains(&paragraph.tag.start) {
            continue;
        }
        let Some(ordinal) = input.entries[at].and_then(|entry| input.recorded.order[entry]) else {
            continue;
        };
        let span = paragraph_span(input.target, paragraph)?;
        let xml = match patched.get(&ordinal) {
            Some(xml) => xml.as_str(),
            None if assignments.contains_key(&ordinal) => &input.target[span.clone()],
            None => continue,
        };
        let id = assignments.get(&ordinal).or(paragraph.para_id.as_ref());
        let xml = if let Some(id) = id {
            patch_paragraph_ids(xml, &BTreeMap::from([(0, id.clone())]), false)?
        } else {
            xml.to_owned()
        };
        result.replace_range(
            span.start - table.target.start..span.end - table.target.start,
            &xml,
        );
    }
    Some(result)
}

pub(super) fn splice(
    source: &str,
    serialized: &str,
    recorded: &RecordedParagraphs,
    part: &S13SplicedPart,
    assignments: &BTreeMap<u32, String>,
    patched: &BTreeMap<u32, String>,
    replace_reasons: &mut Vec<String>,
) -> Option<String> {
    let removed: BTreeSet<_> = part.removed.iter().copied().collect();
    if recorded.conflict
        || recorded.written.len() != part.paragraphs.len()
        || part
            .paragraphs
            .iter()
            .any(|ordinal| !recorded.written.contains_key(ordinal))
        || removed.len() != part.removed.len()
        || removed
            .iter()
            .any(|ordinal| recorded.written.contains_key(ordinal))
        || recorded.inserted.len() != part.inserted.len()
        || patched.keys().any(|ordinal| !part.changed.contains(ordinal))
    {
        return None;
    }
    let input = Input::new(source, serialized, recorded, part)?;
    let tables = input.tables();
    if tables.as_ref().is_some_and(Vec::is_empty)
        || (tables.is_none() && part.table_patches.is_empty())
    {
        return splice_story_part_inner(
            source,
            serialized,
            recorded,
            part,
            assignments,
            patched,
            replace_reasons,
        );
    }
    let tables = tables?;
    let mut replacements = Vec::new();
    for table in &tables {
        if !table.affected {
            continue;
        }
        let mut xml = None;
        if table.legacy {
            let mut projection = Projection::default();
            projection.copy(&input, 0..table.source.start, true)?;
            projection.copy(&input, table.target.clone(), false)?;
            projection.copy(&input, table.source.end..source.len(), true)?;
            let mut local = projection.part(part);
            local.table_patches.clear();
            let patches = patched
                .iter()
                .filter(|(ordinal, _)| local.changed.contains(ordinal))
                .map(|(&ordinal, xml)| (ordinal, xml.clone()))
                .collect();
            if let Some(output) = splice_story_part_inner(
                source,
                &projection.xml,
                &projection.recorded,
                &local,
                assignments,
                &patches,
                &mut Vec::new(),
            ) {
                xml = Some(output[outer_tables(&output)?.get(table.index)?.clone()].to_owned());
            }
        }
        if let Some(xml) = xml {
            replacements.push((table.index, xml, false));
        } else {
            replacements.push((
                table.index,
                replacement(&input, table, assignments, patched)?,
                true,
            ));
        }
    }
    // Isolate replacement tables so projection loss cannot reject unrelated edits.
    let mut projection = Projection::default();
    let mut cursor = 0;
    for table in &tables {
        projection.copy(&input, cursor..table.target.start, false)?;
        projection.copy(&input, table.source.clone(), true)?;
        cursor = table.target.end;
    }
    projection.copy(&input, cursor..serialized.len(), false)?;
    let mut outside = projection.part(part);
    outside.table_patches.retain(|patch| {
        !tables.iter().any(|table| {
            input
                .source_paragraphs
                .get(patch.ordinal as usize)
                .is_some_and(|paragraph| table.source.contains(&paragraph.tag.start))
        })
    });
    let patches = patched
        .iter()
        .filter(|(ordinal, _)| outside.changed.contains(ordinal))
        .map(|(&ordinal, xml)| (ordinal, xml.clone()))
        .collect();
    let mut reasons = Vec::new();
    let output = splice_story_part_inner(
        source,
        &projection.xml,
        &projection.recorded,
        &outside,
        assignments,
        &patches,
        &mut reasons,
    )?;
    let mut prefixes = BTreeSet::new();
    for (_, xml, replaced) in &replacements {
        if *replaced {
            prefixes.extend(required_prefixes(xml)?);
        }
    }
    let mut output = root_with_prefixes(&output, serialized, &prefixes)?.into_owned();
    if !root_binds(&output, serialized, &prefixes, false)? {
        return None;
    }
    let spans = outer_tables(&output)?;
    for (index, xml, replaced) in replacements.into_iter().rev() {
        output.replace_range(spans.get(index)?.clone(), &xml);
        if replaced {
            reasons.push("complex-table:structure".to_owned());
        }
    }
    if !unique_ranges([output.as_str()])? {
        return None;
    }
    replace_reasons.extend(reasons);
    Some(output)
}
