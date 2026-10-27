use std::borrow::Cow;
use std::collections::{BTreeMap, HashMap, HashSet};
use std::sync::Arc;

use yrs::types::text::YChange;
use yrs::{Any, Map, MapRef, Out, ReadTxn, Text};

use crate::batch::DocumentVersion;
use crate::heading::{OutlineInputs, current_inputs, resolve_heading};
use crate::list_marker::{
    ListState, Unrenderable, list_level_formats, list_marker, numbering_level, renders_format,
};
use crate::read_types::HeadingInfo;
use crate::seed::SourceMetadata;
use crate::structured::{ListInfo, MarkerSuffix};
use crate::target::{EditTextView, Mark, ParagraphList};
use crate::{EditingDoc, KIND_KEY, map_string, story_ref};

fn present(value: Option<&Any>) -> Option<&Any> {
    value.filter(|value| !matches!(value, Any::Null | Any::Undefined))
}

fn any_text(value: Option<&Any>) -> Option<String> {
    match present(value)? {
        Any::String(value) => Some(value.to_string()),
        Any::Number(value) if value.is_finite() && value.fract() == 0.0 => {
            Some(format!("{value:.0}"))
        }
        Any::Number(value) => Some(value.to_string()),
        Any::BigInt(value) => Some(value.to_string()),
        _ => None,
    }
}

fn any_number(value: Option<&Any>) -> Option<f64> {
    match present(value)? {
        Any::Number(value) => Some(*value),
        Any::BigInt(value) => Some(*value as f64),
        Any::String(value) => value.parse().ok(),
        _ => None,
    }
}

fn any_map(value: Option<&Any>) -> Option<&HashMap<String, Any>> {
    match present(value)? {
        Any::Map(map) => Some(map),
        _ => None,
    }
}

pub(crate) fn heading(
    properties: &BTreeMap<String, Any>,
    style_id: Option<&str>,
    original: bool,
    source: Option<&SourceMetadata>,
) -> Option<HeadingInfo> {
    let inputs = match (original, previous(properties)) {
        (true, Some(previous)) => OutlineInputs {
            effective: None,
            direct: any_number(previous.get("outlineLevel")),
            style_id,
        },
        _ => OutlineInputs {
            style_id,
            ..current_inputs(properties)
        },
    };
    resolve_heading(&inputs, source)
}

fn previous(properties: &BTreeMap<String, Any>) -> Option<&HashMap<String, Any>> {
    let Any::Array(changes) = properties.get("pPrChange")? else {
        return None;
    };
    any_map(any_map(changes.first())?.get("previousFormatting"))
}

pub(crate) fn numbering_properties<'a>(
    properties: &'a BTreeMap<String, Any>,
    style_id: Option<&str>,
    original: bool,
    source: Option<&SourceMetadata>,
) -> Cow<'a, BTreeMap<String, Any>> {
    let Some(previous) = original.then(|| previous(properties)).flatten() else {
        return Cow::Borrowed(properties);
    };
    let mut values = BTreeMap::new();
    let inherited = source.and_then(|source| source.style_numbering(style_id));
    let num_pr = previous.get("numPr").cloned().or(inherited);
    if let Some(num_pr) = num_pr {
        if let Some(num) = any_map(Some(&num_pr))
            && let Some(source) = source
            && let Some(rendering) = docx_parse::compute_list_rendering(
                any_number(num.get("numId")),
                any_number(num.get("ilvl")),
                &source.numbering(),
            )
        {
            values.insert("listNumFmt".into(), rendering.num_fmt.into());
            values.insert("listIsBullet".into(), Any::Bool(rendering.is_bullet));
            let marker = if rendering.is_bullet {
                let mut characters = rendering.marker.trim_ascii().chars();
                rendering
                    .marker_font_family
                    .as_deref()
                    .and_then(ooxml_text::SymbolFont::named)
                    .and_then(|face| match (characters.next(), characters.next()) {
                        (Some(character), None) => face.shape(character),
                        _ => None,
                    })
                    .map(|shape| shape.to_string())
                    .unwrap_or_else(|| docx_parse::convert_bullet_to_unicode(&rendering.marker))
            } else {
                rendering.marker
            };
            values.insert("listMarker".into(), marker.into());
            values.insert(
                "listMarkerHidden".into(),
                Any::Bool(rendering.marker_hidden == Some(true)),
            );
            values.insert(
                "listLevelNumFmts".into(),
                Any::Array(
                    rendering
                        .level_num_fmts
                        .into_iter()
                        .map(Any::from)
                        .collect::<Vec<_>>()
                        .into(),
                ),
            );
            if let Some(id) = rendering.abstract_num_id {
                values.insert("listAbstractNumId".into(), Any::Number(id));
            }
            if let Some(start) = rendering.start_override {
                values.insert("listStartOverride".into(), Any::Number(start));
            }
            if let Some(suffix) = rendering.marker_suffix {
                values.insert("listMarkerSuffix".into(), suffix.into());
            }
        }
        values.insert("numPr".into(), num_pr);
    }
    Cow::Owned(values)
}

pub(crate) fn list_info(
    values: &BTreeMap<String, Any>,
    list: &mut ListState,
    mut note: impl FnMut(String),
) -> Option<ListInfo> {
    let num_pr = any_map(values.get("numPr"))?;
    let num_id = any_number(num_pr.get("numId")).filter(|id| *id != 0.0)?;
    let num_id = any_text(num_pr.get("numId")).unwrap_or_else(|| num_id.to_string());
    let abstract_num_id = any_text(values.get("listAbstractNumId"));
    let suffix = match any_text(values.get("listMarkerSuffix")).as_deref() {
        Some("space") => MarkerSuffix::Space,
        Some("nothing") => MarkerSuffix::Nothing,
        _ => MarkerSuffix::Tab,
    };
    let marker_hidden = matches!(values.get("listMarkerHidden"), Some(Any::Bool(true)));
    let level_value = any_number(num_pr.get("ilvl")).unwrap_or(0.0);
    let Some(level) = numbering_level(level_value) else {
        note(format!(
            "Numbering level {level_value} is outside 0..8, so the marker is unresolved."
        ));
        return Some(ListInfo {
            num_id,
            abstract_num_id,
            level: if level_value > 8.0 { 8 } else { 0 },
            format: String::new(),
            marker: None,
            suffix,
            marker_hidden,
        });
    };
    let bullet = matches!(values.get("listIsBullet"), Some(Any::Bool(true)));
    let mut formats = list_level_formats(values);
    formats.truncate(9);
    let format = if bullet {
        Some("bullet".to_owned())
    } else {
        formats
            .get(level)
            .cloned()
            .or_else(|| any_text(values.get("listNumFmt")))
    };
    let rendered = list_marker(values, list);
    let template = any_text(values.get("listMarker"));
    let referenced: Vec<String> = match &template {
        Some(template) if !bullet => (1..=9usize)
            .filter(|level| template.contains(&format!("%{level}")))
            .filter_map(|level| formats.get(level - 1).cloned().or_else(|| format.clone()))
            .collect(),
        _ => format.iter().cloned().collect(),
    };
    let unsupported = referenced
        .iter()
        .find(|format| !renders_format(format))
        .cloned();
    let marker = match (&format, &unsupported) {
        (None, _) => {
            note(format!(
                "Numbering instance {num_id} has no resolved format, so its marker is unresolved."
            ));
            None
        }
        (_, Some(unsupported)) => {
            note(format!(
                "Numbering format {unsupported} is not rendered, so the marker is unresolved."
            ));
            None
        }
        _ => match rendered {
            Some(Err(Unrenderable { value, format })) => {
                note(format!(
                    "Number {value} cannot be written in format {format}, so the marker is unresolved."
                ));
                None
            }
            Some(Ok(marker)) => Some(marker),
            None => Some(String::new()),
        },
    };
    Some(ListInfo {
        num_id,
        abstract_num_id,
        level: level as u8,
        format: format.unwrap_or_default(),
        marker,
        suffix,
        marker_hidden,
    })
}

#[derive(Clone, Default)]
pub(crate) struct ParagraphFacts {
    pub outline_level: Option<u8>,
    pub list: Option<ParagraphList>,
}

pub(crate) type RootFacts = HashMap<(String, String), ParagraphFacts>;

#[derive(Default)]
pub(crate) struct FactsCache {
    version: Option<DocumentVersion>,
    roots: HashMap<(String, EditTextView), Arc<RootFacts>>,
}

impl EditingDoc {
    pub(crate) fn root_paragraph_facts<T: ReadTxn>(
        &self,
        txn: &T,
        root: &str,
        view: EditTextView,
    ) -> Arc<RootFacts> {
        let version = self.version();
        let mut cache = self.paragraph_facts.lock().unwrap();
        if cache.version.as_ref() != Some(&version) {
            cache.roots.clear();
            cache.version = Some(version);
        }
        let key = (root.to_owned(), view);
        Arc::clone(cache.roots.entry(key).or_insert_with(|| {
            let source = self.source_metadata();
            let mut list = ListState::new(source.as_ref().map(|source| source.numbering()));
            let mut facts = HashMap::new();
            facts_story(
                txn,
                root,
                view,
                source.as_deref(),
                &mut list,
                &mut facts,
                &mut HashSet::new(),
            );
            Arc::new(facts)
        }))
    }
}

fn fact_properties<T: ReadTxn>(map: &MapRef, txn: &T) -> BTreeMap<String, Any> {
    let mut properties = BTreeMap::new();
    for key in [
        "pStyle",
        "outlineLevel",
        "numPr",
        "listAbstractNumId",
        "listMarkerSuffix",
        "listMarkerHidden",
        "listIsBullet",
        "listLevelNumFmts",
        "listNumFmt",
        "listMarker",
        "listStartOverride",
        "pPrIns",
        "pPrDel",
    ] {
        if let Some(Out::Any(value)) = map.get(txn, key) {
            properties.insert(key.to_owned(), value);
        }
    }
    if let Some(Out::Any(Any::Map(direct))) = map.get(txn, "_originalFormatting") {
        let outline = direct
            .get("outlineLevel")
            .map(|value| ("outlineLevel".to_owned(), value.clone()));
        properties.insert(
            "_originalFormatting".into(),
            Any::Map(Arc::new(outline.into_iter().collect())),
        );
    }
    if let Some(Out::Any(Any::Array(changes))) = map.get(txn, "pPrChange")
        && let Some(change) = changes.first().and_then(|change| any_map(Some(change)))
    {
        let previous = any_map(change.get("previousFormatting"));
        let earlier = ["pStyle", "styleId", "numPr", "outlineLevel"]
            .into_iter()
            .filter_map(|key| {
                previous
                    .and_then(|previous| previous.get(key))
                    .map(|value| (key.to_owned(), value.clone()))
            })
            .collect();
        properties.insert(
            "pPrChange".into(),
            Any::Array(
                vec![Any::Map(Arc::new(HashMap::from([(
                    "previousFormatting".to_owned(),
                    Any::Map(Arc::new(earlier)),
                )])))]
                .into(),
            ),
        );
    }
    properties
}

fn facts_story<T: ReadTxn>(
    txn: &T,
    story: &str,
    view: EditTextView,
    source: Option<&SourceMetadata>,
    list: &mut ListState,
    facts: &mut RootFacts,
    active: &mut HashSet<String>,
) {
    if active.len() >= 32 || !active.insert(story.to_owned()) {
        return;
    }
    let Ok(text) = story_ref(txn, story) else {
        active.remove(story);
        return;
    };
    let mut previous_hidden = false;
    for diff in text.diff(txn, YChange::identity) {
        let Out::YMap(map) = diff.insert else {
            continue;
        };
        if crate::is_pilcrow(&map, txn) {
            let id = map_string(&map, txn, crate::PARA_ID).unwrap_or_default();
            let mark = Mark {
                properties: fact_properties(&map, txn),
                revision: false,
            };
            let original = view == EditTextView::Original;
            let (attribute, property) = if original {
                (crate::INS, "pPrIns")
            } else {
                (crate::DEL, "pPrDel")
            };
            let hidden = diff
                .attributes
                .as_ref()
                .and_then(|attrs| present(attrs.get(attribute)))
                .is_some()
                || present(mark.properties.get(property)).is_some();
            let revised = hidden || previous_hidden;
            previous_hidden = hidden;
            if revised {
                continue;
            }
            let style = if original {
                mark.original_style_id()
            } else {
                mark.style_id()
            };
            let outline_level = heading(&mark.properties, style.as_deref(), original, source)
                .map(|heading| heading.outline_level);
            let values = numbering_properties(&mark.properties, style.as_deref(), original, source);
            let info = list_info(&values, list, |_| {});
            facts.insert(
                (story.to_owned(), id),
                ParagraphFacts {
                    outline_level,
                    list: info.map(|info| ParagraphList {
                        marker: info.marker,
                        level: info.level,
                        num_id: info.num_id,
                        format: info.format,
                    }),
                },
            );
            continue;
        }
        let hidden = match view {
            EditTextView::Accepted => crate::DEL,
            EditTextView::Original => crate::INS,
        };
        if diff
            .attributes
            .as_ref()
            .and_then(|attrs| attrs.get(hidden))
            .is_some_and(|value| !matches!(value, Any::Null | Any::Undefined))
        {
            continue;
        }
        let mut children = Vec::new();
        match map_string(&map, txn, KIND_KEY).as_deref() {
            Some("blockSdt") => {
                children.extend(map_string(&map, txn, "story"));
            }
            Some("table") => {
                if let Some(Out::Any(Any::Array(rows))) = map.get(txn, "rows") {
                    for row in rows.iter().filter_map(|row| any_map(Some(row))) {
                        let hidden_row = match view {
                            EditTextView::Accepted => "trDel",
                            EditTextView::Original => "trIns",
                        };
                        if any_map(row.get("trPr"))
                            .and_then(|pr| present(pr.get(hidden_row)))
                            .is_some()
                        {
                            continue;
                        }
                        if let Some(Any::Array(cells)) = row.get("cells") {
                            for cell in cells.iter().filter_map(|cell| any_map(Some(cell))) {
                                children.extend(any_text(cell.get("story")));
                            }
                        }
                    }
                }
            }
            _ => {}
        }
        for child in children {
            facts_story(txn, &child, view, source, list, facts, active);
        }
    }
    active.remove(story);
}
