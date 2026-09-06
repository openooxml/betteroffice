use std::{collections::HashMap, sync::Arc};

use pptx_parse::PptxPackage;
use serde_json::Value;
use yrs::types::Attrs;
use yrs::{Any, Map, Out, ReadTxn, Text, TextRef, Transact};

use crate::{
    DeckSession, EditError, EditResult, META, MIGRATE_ORIGIN, ShapeSnapshot, StorySnapshot,
};

#[derive(Clone, Copy)]
pub(crate) enum SourceProperty {
    Baseline,
    Spacing,
}

impl SourceProperty {
    fn keys(self) -> (&'static str, &'static str, &'static str) {
        match self {
            Self::Baseline => ("baselinesPendingSource", "baselinePct", "baseline"),
            Self::Spacing => ("spacingPendingSource", "spacingPt", "spacing"),
        }
    }

    fn value(self, style: &crate::TextStyle) -> Option<f64> {
        match self {
            Self::Baseline => style.baseline_pct,
            Self::Spacing => style.spacing_pt,
        }
    }
}

pub(crate) fn import_source(
    session: &DeckSession,
    source: &PptxPackage,
    property: SourceProperty,
) -> EditResult<()> {
    let (pending_key, json_key, attribute) = property.keys();
    let pending = {
        let txn = session.doc.transact();
        txn.get_map(META)
            .is_some_and(|meta| meta.get(&txn, pending_key) == Some(Out::Any(Any::Bool(true))))
    };
    if !pending {
        return Ok(());
    }
    let source_json =
        serde_json::to_value(source).map_err(|error| EditError::Json(error.to_string()))?;
    if !has_property(&source_json, json_key) {
        return Ok(());
    }
    let fresh = DeckSession::from_package(source.clone(), 33100)?;
    let original = fresh.snapshot()?;
    let current = session.snapshot()?;
    let mut sources = HashMap::new();
    for slide in &original.slides {
        collect_stories(&slide.shapes, &mut sources);
    }
    let mut targets = HashMap::new();
    for slide in &current.slides {
        collect_stories(&slide.shapes, &mut targets);
    }
    let mut patches = Vec::new();
    for (id, target) in targets {
        let Some(source) = sources.get(id) else {
            continue;
        };
        let source_tokens = tokens(source, property);
        if !source_tokens.iter().any(|(_, baseline)| baseline.is_some()) {
            continue;
        }
        let target_tokens = tokens(target, property);
        let pairs = unchanged_pairs(&source_tokens, &target_tokens)?;
        let mut offset = 0u32;
        let positions: Vec<_> = target_tokens
            .iter()
            .map(|(ch, _)| {
                let start = offset;
                offset += ch.len_utf16() as u32;
                (start, offset)
            })
            .collect();
        for (source_index, target_index) in pairs {
            if target_tokens[target_index].1.is_none()
                && let Some(baseline) = source_tokens[source_index].1
            {
                let (start, end) = positions[target_index];
                patches.push((id, start, end, baseline));
            }
        }
    }
    let mut package = serde_json::to_value(crate::deck::package_from_doc(&session.doc)?)
        .map_err(|error| EditError::Json(error.to_string()))?;
    merge_property(&mut package, &source_json, json_key);
    let package: PptxPackage =
        serde_json::from_value(package).map_err(|error| EditError::Json(error.to_string()))?;
    let bytes = serde_json::to_vec(&package).map_err(|error| EditError::Json(error.to_string()))?;
    let mut txn = session.doc.transact_mut_with(MIGRATE_ORIGIN);
    let stories = txn
        .get_map(crate::STORIES)
        .ok_or_else(|| EditError::InvalidState("missing stories".into()))?;
    for (id, start, end, baseline) in patches {
        let story = stories
            .get(&txn, id)
            .and_then(|value| value.cast::<TextRef>().ok())
            .ok_or_else(|| EditError::StoryNotFound(id.into()))?;
        story.format(
            &mut txn,
            start,
            end - start,
            Attrs::from([(attribute.into(), Any::Number(baseline))]),
        );
    }
    let meta = txn
        .get_map(META)
        .ok_or_else(|| EditError::InvalidState("missing metadata".into()))?;
    meta.insert(&mut txn, "packageJson", Any::Buffer(Arc::from(bytes)));
    meta.remove(&mut txn, pending_key);
    Ok(())
}

fn collect_stories<'a>(
    shapes: &'a [ShapeSnapshot],
    stories: &mut HashMap<&'a str, &'a StorySnapshot>,
) {
    for shape in shapes {
        for story in &shape.text_stories {
            stories.insert(&story.id, story);
        }
        collect_stories(&shape.children, stories);
    }
}

fn tokens(story: &StorySnapshot, property: SourceProperty) -> Vec<(char, Option<f64>)> {
    let mut tokens = Vec::new();
    for paragraph in &story.paragraphs {
        for run in &paragraph.runs {
            tokens.extend(
                run.text
                    .chars()
                    .map(|unit| (unit, property.value(&run.style))),
            );
        }
        tokens.push(('\0', None));
    }
    tokens
}

fn unchanged_pairs(
    source: &[(char, Option<f64>)],
    target: &[(char, Option<f64>)],
) -> EditResult<Vec<(usize, usize)>> {
    let mut prefix = 0;
    while prefix < source.len().min(target.len()) && source[prefix].0 == target[prefix].0 {
        prefix += 1;
    }
    let mut suffix = 0;
    while suffix < source.len().min(target.len()) - prefix
        && source[source.len() - suffix - 1].0 == target[target.len() - suffix - 1].0
    {
        suffix += 1;
    }
    let rows = source.len() - prefix - suffix + 1;
    let cols = target.len() - prefix - suffix + 1;
    let cells = rows
        .checked_mul(cols)
        .filter(|cells| *cells <= 4_000_000)
        .ok_or_else(|| {
            EditError::InvalidState("source run property recovery exceeds text diff limit".into())
        })?;
    let mut lengths = vec![0u32; cells];
    for i in (0..rows - 1).rev() {
        for j in (0..cols - 1).rev() {
            lengths[i * cols + j] = if source[prefix + i].0 == target[prefix + j].0 {
                lengths[(i + 1) * cols + j + 1] + 1
            } else {
                lengths[(i + 1) * cols + j].max(lengths[i * cols + j + 1])
            };
        }
    }
    let mut pairs: Vec<_> = (0..prefix).map(|i| (i, i)).collect();
    let (mut i, mut j) = (0, 0);
    while i + 1 < rows && j + 1 < cols {
        if source[prefix + i].0 == target[prefix + j].0 {
            pairs.push((prefix + i, prefix + j));
            i += 1;
            j += 1;
        } else if lengths[(i + 1) * cols + j] >= lengths[i * cols + j + 1] {
            i += 1;
        } else {
            j += 1;
        }
    }
    pairs.extend((0..suffix).map(|i| (source.len() - suffix + i, target.len() - suffix + i)));
    Ok(pairs)
}

fn merge_property(target: &mut Value, source: &Value, key: &str) {
    match (target, source) {
        (Value::Object(target), Value::Object(source)) => {
            if let Some(baseline) = source.get(key) {
                target.insert(key.into(), baseline.clone());
            }
            for (child_key, target) in target {
                if let Some(source) = source.get(child_key) {
                    merge_property(target, source, key);
                }
            }
        }
        (Value::Array(target), Value::Array(source)) => {
            for (target, source) in target.iter_mut().zip(source) {
                merge_property(target, source, key);
            }
        }
        _ => {}
    }
}

fn has_property(value: &Value, key: &str) -> bool {
    match value {
        Value::Object(object) => {
            object.contains_key(key) || object.values().any(|value| has_property(value, key))
        }
        Value::Array(array) => array.iter().any(|value| has_property(value, key)),
        _ => false,
    }
}
