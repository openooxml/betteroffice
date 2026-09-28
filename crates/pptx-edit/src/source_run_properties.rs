use std::collections::HashMap;
use std::ops::Range;

use serde_json::Value;
use yrs::types::Attrs;
use yrs::types::text::YChange;
use yrs::{Any, Assoc, ClientID, IndexedSequence, Map, Out, ReadTxn, Text, TextRef, Transact};

use crate::deck::{SourceImport, map_string};
use crate::story::style_from_attrs;
use crate::{
    BOOTSTRAP_CLIENT_ID, DeckSession, EditError, EditResult, META, MIGRATE_ORIGIN, PARA_ID,
    ShapeSnapshot, StorySnapshot,
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
    let mut sources = Vec::new();
    for slide in &import.source_snapshot()?.slides {
        collect_source_tokens(&slide.shapes, &legacy, property, &mut sources);
    }
    let mut package = serde_json::to_value(&import.package)
        .map_err(|error| EditError::Json(error.to_string()))?;
    for key in json_keys {
        merge_property(&mut package, &source_json, key);
    }
    import.package =
        serde_json::from_value(package).map_err(|error| EditError::Json(error.to_string()))?;
    let mut patches = Vec::new();
    {
        let txn = session.doc.transact();
        let stories = txn
            .get_map(crate::STORIES)
            .ok_or_else(|| EditError::InvalidState("missing stories".into()))?;
        for (id, source) in &sources {
            let Some(story) = stories
                .get(&txn, id)
                .and_then(|value| value.cast::<TextRef>().ok())
            else {
                continue;
            };
            let seeded = seeded_tokens(&story, &txn, id, property);
            for (start, end, value) in recovered(source, &seeded) {
                patches.push((story.clone(), start, end, value));
            }
        }
    }
    let mut txn = session.doc.transact_mut_with(MIGRATE_ORIGIN);
    for (story, start, end, value) in patches {
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

#[derive(Clone, Copy, PartialEq)]
enum Token {
    Char(char),
    Paragraph(usize),
}

struct SourceToken {
    token: Token,
    value: Option<Any>,
    legacy: Option<Any>,
}

struct SeededToken {
    token: Token,
    start: u32,
    end: u32,
    value: Option<Any>,
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

/// Tokenizes each source story whose values differ anywhere from what a legacy seed stored.
fn collect_source_tokens(
    shapes: &[ShapeSnapshot],
    legacy: &HashMap<&str, &StorySnapshot>,
    property: SourceProperty,
    sources: &mut Vec<(String, Vec<SourceToken>)>,
) {
    for shape in shapes {
        for story in &shape.text_stories {
            let values = token_values(story, property);
            let legacy = match legacy.get(story.id.as_str()) {
                Some(legacy) => token_values(legacy, property),
                None => vec![None; values.len()],
            };
            if values == legacy {
                continue;
            }
            let tokens = story
                .paragraphs
                .iter()
                .enumerate()
                .flat_map(|(index, paragraph)| {
                    paragraph
                        .runs
                        .iter()
                        .flat_map(|run| run.text.chars().map(Token::Char))
                        .chain([Token::Paragraph(index)])
                })
                .zip(values.into_iter().zip(legacy))
                .map(|(token, (value, legacy))| SourceToken {
                    token,
                    value,
                    legacy,
                })
                .collect();
            sources.push((story.id.clone(), tokens));
        }
        collect_source_tokens(&shape.children, legacy, property, sources);
    }
}

fn token_values(story: &StorySnapshot, property: SourceProperty) -> Vec<Option<Any>> {
    let mut values = Vec::new();
    for paragraph in &story.paragraphs {
        for run in &paragraph.runs {
            let value = property.value(&run.style);
            values.extend(run.text.chars().map(|_| value.clone()));
        }
        values.push(None);
    }
    values
}

/// The story's characters the seed wrote, told apart from later insertions by
/// the CRDT client that inserted them, and the seeded paragraph ends.
fn seeded_tokens<T: ReadTxn>(
    story: &TextRef,
    txn: &T,
    story_id: &str,
    property: SourceProperty,
) -> Vec<SeededToken> {
    let prefix = format!("para:{story_id}:");
    let bootstrap = ClientID::new(BOOTSTRAP_CLIENT_ID);
    let mut tokens = Vec::new();
    let mut offset = 0;
    for diff in story.diff(txn, YChange::identity) {
        match &diff.insert {
            Out::Any(Any::String(text)) => {
                let value = property.value(&style_from_attrs(diff.attributes.as_deref()));
                for unit in text.chars() {
                    let end = offset + unit.len_utf16() as u32;
                    if story
                        .sticky_index(txn, offset, Assoc::After)
                        .and_then(|index| index.id().map(|id| id.client == bootstrap))
                        .unwrap_or(false)
                    {
                        tokens.push(SeededToken {
                            token: Token::Char(unit),
                            start: offset,
                            end,
                            value: value.clone(),
                        });
                    }
                    offset = end;
                }
            }
            Out::YMap(pilcrow) => {
                if let Some(index) = map_string(pilcrow, txn, PARA_ID)
                    .and_then(|id| id.strip_prefix(&prefix)?.parse().ok())
                {
                    tokens.push(SeededToken {
                        token: Token::Paragraph(index),
                        start: offset,
                        end: offset + 1,
                        value: None,
                    });
                }
                offset += 1;
            }
            _ => offset += 1,
        }
    }
    tokens
}

/// The ranges of seeded characters still holding their legacy value, with the
/// source value to write. Seeded paragraph ends pin each stretch of characters
/// to its source paragraphs.
fn recovered(source: &[SourceToken], seeded: &[SeededToken]) -> Vec<(u32, u32, Any)> {
    let mut paragraphs = HashMap::new();
    let mut changes = Vec::with_capacity(source.len());
    let mut previous = None;
    let mut count = 0;
    for (position, token) in source.iter().enumerate() {
        match token.token {
            Token::Paragraph(index) => {
                paragraphs.insert(index, position);
            }
            Token::Char(_) => {
                let values = (&token.value, &token.legacy);
                count += usize::from(previous.is_some_and(|previous| previous != values));
                previous = Some(values);
            }
        }
        changes.push(count);
    }
    let mut ranges = Vec::new();
    let mut from = 0;
    let mut stretch = Vec::new();
    for token in seeded {
        match token.token {
            Token::Char(_) => stretch.push(token),
            Token::Paragraph(index) => {
                if let Some(&until) = paragraphs.get(&index).filter(|until| **until >= from) {
                    recover_stretch(source, &changes, from..until, &stretch, &mut ranges);
                    from = until + 1;
                    stretch.clear();
                }
            }
        }
    }
    recover_stretch(source, &changes, from..source.len(), &stretch, &mut ranges);
    ranges
}

/// A character's candidate source positions run from its earliest to its latest
/// in-order match within `window`; it is written only when they all carry the
/// same values.
fn recover_stretch(
    source: &[SourceToken],
    changes: &[usize],
    window: Range<usize>,
    stretch: &[&SeededToken],
    ranges: &mut Vec<(u32, u32, Any)>,
) {
    let from = window.start;
    let window = &source[window];
    let (Some(earliest), Some(latest)) = (earliest(window, stretch), latest(window, stretch))
    else {
        return;
    };
    for ((token, earliest), latest) in stretch.iter().zip(earliest).zip(latest) {
        let (earliest, latest) = (from + earliest, from + latest);
        let candidate = &source[earliest];
        if changes[earliest] != changes[latest]
            || candidate.value == candidate.legacy
            || token.value != candidate.legacy
        {
            continue;
        }
        let value = candidate.value.clone().unwrap_or(Any::Null);
        match ranges.last_mut() {
            Some((_, end, last)) if *end == token.start && *last == value => *end = token.end,
            _ => ranges.push((token.start, token.end, value)),
        }
    }
}

fn earliest(window: &[SourceToken], stretch: &[&SeededToken]) -> Option<Vec<usize>> {
    let mut positions = Vec::with_capacity(stretch.len());
    let mut next = 0;
    for token in stretch {
        next += window[next..]
            .iter()
            .position(|source| source.token == token.token)?;
        positions.push(next);
        next += 1;
    }
    Some(positions)
}

fn latest(window: &[SourceToken], stretch: &[&SeededToken]) -> Option<Vec<usize>> {
    let mut positions = vec![0; stretch.len()];
    let mut end = window.len();
    for (position, token) in positions.iter_mut().zip(stretch).rev() {
        end = window[..end]
            .iter()
            .rposition(|source| source.token == token.token)?;
        *position = end;
    }
    Some(positions)
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
