use std::collections::HashMap;

use serde_json::Value;
use yrs::types::Attrs;
use yrs::{Any, Map, Out, ReadTxn, Text, TextRef, Transact};

use crate::deck::SourceImport;
use crate::{
    DeckSession, EditError, EditResult, META, MIGRATE_ORIGIN, ShapeSnapshot, StorySnapshot,
};

#[derive(Clone, Copy)]
pub(crate) enum SourceProperty {
    Baseline,
    Spacing,
    Caps,
    Color,
}

impl SourceProperty {
    /// The pending flag, the package keys the source restores, and the run attribute.
    fn keys(self) -> (&'static str, &'static [&'static str], &'static str) {
        match self {
            Self::Baseline => ("baselinesPendingSource", &["baselinePct"], "baseline"),
            Self::Spacing => ("spacingPendingSource", &["spacingPt"], "spacing"),
            Self::Caps => ("capsPendingSource", &["caps"], "caps"),
            Self::Color => (
                "colorsPendingSource",
                &["colorMap", "colorMapOverride"],
                "color",
            ),
        }
    }

    fn value(self, style: &crate::TextStyle) -> Option<Any> {
        match self {
            Self::Baseline => style.baseline_pct.map(Any::Number),
            Self::Spacing => style.spacing_pt.map(Any::Number),
            Self::Caps => style.caps.map(|caps| Any::from(caps.as_attribute())),
            Self::Color => style.color.as_deref().map(Any::from),
        }
    }
}

pub(crate) fn import_source(
    session: &DeckSession,
    import: &mut SourceImport<'_>,
    property: SourceProperty,
) -> EditResult<()> {
    let (pending_key, json_keys, attribute) = property.keys();
    let pending = {
        let txn = session.doc.transact();
        txn.get_map(META)
            .is_some_and(|meta| meta.get(&txn, pending_key) == Some(Out::Any(Any::Bool(true))))
    };
    if !pending {
        return Ok(());
    }
    let source_json =
        serde_json::to_value(import.source).map_err(|error| EditError::Json(error.to_string()))?;
    if !json_keys.iter().any(|key| has_property(&source_json, key)) {
        return Ok(());
    }
    let legacy_snapshot = match property {
        SourceProperty::Color => Some(crate::deck::legacy_baseline_snapshot(import.source)?),
        _ => None,
    };
    let mut legacy = HashMap::new();
    for slide in legacy_snapshot.iter().flat_map(|snapshot| &snapshot.slides) {
        collect_stories(&slide.shapes, &mut legacy);
    }
    let mut sources = HashMap::new();
    {
        let original = import.source_snapshot()?;
        for slide in &original.slides {
            collect_stories(&slide.shapes, &mut sources);
        }
    }
    let current = session.snapshot()?;
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
        let legacy_values: Vec<_> = match legacy.get(id) {
            Some(story) => tokens(story, property)
                .into_iter()
                .map(|(_, value)| value)
                .collect(),
            None => vec![None; source_tokens.len()],
        };
        if source_tokens
            .iter()
            .zip(&legacy_values)
            .all(|((_, value), legacy)| value == legacy)
        {
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
            let value = &source_tokens[source_index].1;
            let legacy = &legacy_values[source_index];
            if value != legacy && target_tokens[target_index].1 == *legacy {
                let (start, end) = positions[target_index];
                patches.push((id, start, end, value.clone().unwrap_or(Any::Null)));
            }
        }
    }
    let mut package = serde_json::to_value(&import.package)
        .map_err(|error| EditError::Json(error.to_string()))?;
    for key in json_keys {
        merge_property(&mut package, &source_json, key);
    }
    import.package =
        serde_json::from_value(package).map_err(|error| EditError::Json(error.to_string()))?;
    let mut txn = session.doc.transact_mut_with(MIGRATE_ORIGIN);
    let stories = txn
        .get_map(crate::STORIES)
        .ok_or_else(|| EditError::InvalidState("missing stories".into()))?;
    for (id, start, end, value) in patches {
        let story = stories
            .get(&txn, id)
            .and_then(|value| value.cast::<TextRef>().ok())
            .ok_or_else(|| EditError::StoryNotFound(id.into()))?;
        story.format(
            &mut txn,
            start,
            end - start,
            Attrs::from([(attribute.into(), value)]),
        );
    }
    let meta = txn
        .get_map(META)
        .ok_or_else(|| EditError::InvalidState("missing metadata".into()))?;
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

fn tokens(story: &StorySnapshot, property: SourceProperty) -> Vec<(char, Option<Any>)> {
    let mut tokens = Vec::new();
    for paragraph in &story.paragraphs {
        for run in &paragraph.runs {
            let value = property.value(&run.style);
            tokens.extend(run.text.chars().map(|unit| (unit, value.clone())));
        }
        tokens.push(('\0', None));
    }
    tokens
}

fn unchanged_pairs(
    source: &[(char, Option<Any>)],
    target: &[(char, Option<Any>)],
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
