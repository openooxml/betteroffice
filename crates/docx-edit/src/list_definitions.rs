//! List definitions the editor creates, beside the numbering definitions of the source package.

use std::collections::{BTreeMap, BTreeSet, HashMap};

use yrs::{Any, Map, MapRef, Out, ReadTxn, TextRef, Transact};

use crate::bridge::{any_bool, map_number};
use crate::list_marker::{LEVELS, ListState, numbering_level};
use crate::paragraph_facts::any_text;
use crate::seed::SourceMetadata;
use crate::{EditCtx, EditError, EditResult, EditingDoc};

/// The shared root holding the list definitions the editor creates, by session id.
pub const LIST_DEFINITIONS: &str = "listDefinitions";

/// The paragraph attribute naming the created definition a paragraph is on, as
/// `{"key": <string>, "numId": <number>}`; honoured while `numPr.numId` still equals `numId`.
pub const LIST_DEF: &str = "listDef";

/// The bullets of Word's default bullet list, in Symbol, Courier New and Wingdings, repeating.
const BULLET_TEXTS: [&str; 3] = ["\u{f0b7}", "o", "\u{f0a7}"];

/// The number formats of Word's default numbered list, repeating.
const NUMBER_FORMATS: [&str; 3] = ["decimal", "lowerLetter", "lowerRoman"];

/// The list attributes a paragraph joining a source list takes from its neighbour.
const SOURCE_LIST_KEYS: [&str; 14] = [
    "numPr",
    "listIsBullet",
    "listNumFmt",
    "listMarker",
    "listLevelNumFmts",
    "listAbstractNumId",
    "listStartOverride",
    "listMarkerHidden",
    "listMarkerBold",
    "listMarkerItalic",
    "listMarkerColor",
    "listMarkerFontFamily",
    "listMarkerFontSize",
    "listMarkerSuffix",
];

/// The kind of list the editor creates.
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum ListKind {
    Bullet,
    Numbered,
}

impl ListKind {
    /// `bullet` or `numbered`.
    pub fn name(self) -> &'static str {
        match self {
            Self::Bullet => "bullet",
            Self::Numbered => "numbered",
        }
    }

    fn from_name(name: &str) -> Option<Self> {
        match name {
            "bullet" => Some(Self::Bullet),
            "numbered" => Some(Self::Numbered),
            _ => None,
        }
    }

    /// The `w:numFmt` of `level` in a list of this kind the editor creates.
    pub fn level_format(self, level: usize) -> &'static str {
        match self {
            Self::Bullet => "bullet",
            Self::Numbered => NUMBER_FORMATS[level % NUMBER_FORMATS.len()],
        }
    }

    /// The `w:lvlText` of `level` in a list of this kind the editor creates.
    fn level_text(self, level: usize) -> String {
        match self {
            Self::Bullet => BULLET_TEXTS[level % BULLET_TEXTS.len()].to_owned(),
            Self::Numbered => format!("%{}.", level + 1),
        }
    }

    /// The `numPr.numId` a paragraph on a created list of this kind carries, which marks the
    /// kind for peers that do not read [`LIST_DEF`].
    fn marker_num_id(self) -> f64 {
        match self {
            Self::Bullet => 1.0,
            Self::Numbered => 2.0,
        }
    }
}

/// A list definition the editor created.
#[derive(Clone, Debug, PartialEq)]
pub struct CreatedListDefinition {
    pub key: String,
    pub kind: ListKind,
    /// The `w:lvlText` of levels 0 to 8.
    pub level_texts: Vec<String>,
}

impl CreatedListDefinition {
    /// The `w:numFmt` of each level.
    pub fn level_formats(&self) -> Vec<String> {
        (0..self.level_texts.len())
            .map(|level| self.kind.level_format(level).to_owned())
            .collect()
    }
}

/// The list a paragraph is on.
#[derive(Clone, Debug, PartialEq)]
pub enum ListIdentity {
    /// A list of the source package, by its `w:numId`.
    Source { num_id: f64, level: f64 },
    /// A list the editor created, by its key in [`LIST_DEFINITIONS`].
    Created { key: String, level: f64 },
}

impl ListIdentity {
    /// The level the paragraph is at, as written.
    pub fn level(&self) -> f64 {
        match self {
            Self::Source { level, .. } | Self::Created { level, .. } => *level,
        }
    }
}

/// The key of the created definition `values` names in [`LIST_DEF`], while its `numPr` still
/// carries the numId recorded beside the key.
fn created_key(values: &BTreeMap<String, Any>) -> Option<&str> {
    let Some(Any::Map(num_pr)) = values.get("numPr") else {
        return None;
    };
    let Some(Any::Map(reference)) = values.get(LIST_DEF) else {
        return None;
    };
    let Some(Any::String(key)) = reference.get("key") else {
        return None;
    };
    let recorded = map_number(reference, "numId")?;
    (map_number(num_pr, "numId") == Some(recorded)).then_some(key.as_ref())
}

/// The numbering instance and level the `numPr` of `values` names, `(numId, ilvl)`, whatever list
/// the paragraph is on; `None` when it names none.
pub(crate) fn numbered_level(values: &BTreeMap<String, Any>) -> Option<(f64, f64)> {
    let Some(Any::Map(num_pr)) = values.get("numPr") else {
        return None;
    };
    let num_id = map_number(num_pr, "numId").filter(|id| *id != 0.0)?;
    Some((num_id, map_number(num_pr, "ilvl").unwrap_or(0.0)))
}

/// The list a paragraph with the attributes `values` is on, or `None` when it is on no list.
pub fn resolve_list_identity(values: &BTreeMap<String, Any>) -> Option<ListIdentity> {
    let (num_id, level) = numbered_level(values)?;
    Some(match created_key(values) {
        Some(key) => ListIdentity::Created {
            key: key.to_owned(),
            level,
        },
        None => ListIdentity::Source { num_id, level },
    })
}

/// The `numPr` of `values` as written, `(numId, ilvl)`, which layout reads.
pub(crate) fn written_numbering(
    values: &BTreeMap<String, Any>,
) -> Option<(Option<f64>, Option<f64>)> {
    let Some(Any::Map(num_pr)) = values.get("numPr") else {
        return None;
    };
    Some((map_number(num_pr, "numId"), map_number(num_pr, "ilvl")))
}

/// The id the read surfaces give the numbering instance `num_pr` names, as written; `None` on no
/// list.
pub(crate) fn written_list_id(num_pr: &HashMap<String, Any>) -> Option<String> {
    let num_id = map_number(num_pr, "numId").filter(|id| *id != 0.0)?;
    Some(any_text(num_pr.get("numId")).unwrap_or_else(|| num_id.to_string()))
}

/// How the source package renders the list an earlier `numPr` of a paragraph names.
pub(crate) fn source_list_rendering(
    num_pr: &HashMap<String, Any>,
    source: &SourceMetadata,
) -> Option<docx_parse::ListRendering> {
    docx_parse::compute_list_rendering(
        map_number(num_pr, "numId"),
        map_number(num_pr, "ilvl"),
        &source.numbering(),
    )
}

/// The list state over the effective definitions: the numbering of the source package and the
/// definitions the editor created, as `txn` reads them.
pub(crate) fn effective_list_state<T: ReadTxn>(
    source: Option<&SourceMetadata>,
    txn: &T,
) -> ListState {
    let numbering = source.map(|source| source.numbering());
    ListState::new(numbering).with_created(created_definitions(txn))
}

fn created_definitions<T: ReadTxn>(txn: &T) -> BTreeMap<String, CreatedListDefinition> {
    let Some(root) = txn.get_map(LIST_DEFINITIONS) else {
        return BTreeMap::new();
    };
    root.iter(txn)
        .filter_map(|(key, value)| {
            let Out::Any(Any::Map(entry)) = value else {
                return None;
            };
            let Some(Any::String(kind)) = entry.get("kind") else {
                return None;
            };
            let Some(Any::Array(texts)) = entry.get("levelTexts") else {
                return None;
            };
            let kind = ListKind::from_name(kind)?;
            let mut level_texts = texts
                .iter()
                .take(LEVELS)
                .map(|text| match text {
                    Any::String(text) => Some(xml_text(text)),
                    _ => None,
                })
                .collect::<Option<Vec<_>>>()?;
            for level in level_texts.len()..LEVELS {
                level_texts.push(kind.level_text(level));
            }
            let definition = CreatedListDefinition {
                key: key.to_owned(),
                kind,
                level_texts,
            };
            Some((key.to_owned(), definition))
        })
        .collect()
}

/// `text` without the characters XML 1.0 does not allow.
fn xml_text(text: &str) -> String {
    text.chars()
        .filter(|&character| {
            matches!(character, '\t' | '\n' | '\r' | ' '..='\u{d7ff}' | '\u{e000}'..='\u{fffd}')
                || character >= '\u{10000}'
        })
        .collect()
}

/// The kind of list a paragraph on a list shows, from its bullet flag and its format.
fn shown_kind(values: &BTreeMap<String, Any>) -> ListKind {
    let bullet_format =
        matches!(values.get("listNumFmt"), Some(Any::String(format)) if &**format == "bullet");
    if values.get("listIsBullet").and_then(any_bool) == Some(true) || bullet_format {
        ListKind::Bullet
    } else {
        ListKind::Numbered
    }
}

/// Whether the list of a paragraph with `values` comes from its style: it is on no created list and
/// its `numPr` names the numbering the style gives, which seeding keeps only in the
/// `numPrFromStyle` of the paragraph's `_originalFormatting`.
fn numbered_by_style(values: &BTreeMap<String, Any>) -> bool {
    let original = match values.get("_originalFormatting") {
        Some(Any::Map(original)) => original.get("numPrFromStyle"),
        _ => None,
    };
    let style = values.get("numPrFromStyle").or(original);
    let (Some(Any::Map(style)), Some(Any::Map(num_pr))) = (style, values.get("numPr")) else {
        return false;
    };
    created_key(values).is_none()
        && map_number(style, "numId").is_some_and(|id| map_number(num_pr, "numId") == Some(id))
}

/// The attributes that put a paragraph on level `level` of the created list `key` of `kind`.
fn created_attrs(kind: ListKind, key: &str, level: f64) -> BTreeMap<String, Any> {
    let num_id = kind.marker_num_id();
    let shown = numbering_level(level).unwrap_or(0);
    let marker = match kind {
        ListKind::Bullet => docx_parse::convert_bullet_to_unicode(&kind.level_text(shown)),
        ListKind::Numbered => kind.level_text(shown),
    };
    let formats = match kind {
        ListKind::Bullet => Any::Null,
        ListKind::Numbered => Any::Array(
            (0..LEVELS)
                .map(|level| Any::from(kind.level_format(level)))
                .collect::<Vec<_>>()
                .into(),
        ),
    };
    let num_pr = HashMap::from([
        ("numId".to_owned(), Any::Number(num_id)),
        ("ilvl".to_owned(), Any::Number(level)),
    ]);
    let reference = HashMap::from([
        ("key".to_owned(), Any::from(key)),
        ("numId".to_owned(), Any::Number(num_id)),
    ]);
    BTreeMap::from([
        ("numPr".to_owned(), Any::from(num_pr)),
        (LIST_DEF.to_owned(), Any::from(reference)),
        (
            "listIsBullet".to_owned(),
            Any::Bool(kind == ListKind::Bullet),
        ),
        ("listNumFmt".to_owned(), Any::from(kind.level_format(shown))),
        ("listMarker".to_owned(), Any::String(marker.into())),
        ("listLevelNumFmts".to_owned(), formats),
        ("listAbstractNumId".to_owned(), Any::Null),
        ("listStartOverride".to_owned(), Any::Null),
        ("listMarkerHidden".to_owned(), Any::Null),
        ("listMarkerBold".to_owned(), Any::Null),
        ("listMarkerItalic".to_owned(), Any::Null),
        ("listMarkerColor".to_owned(), Any::Null),
        ("listMarkerFontFamily".to_owned(), Any::Null),
        ("listMarkerFontSize".to_owned(), Any::Null),
        ("listMarkerSuffix".to_owned(), Any::Null),
    ])
}

/// The attributes that put a paragraph on the source list of its neighbour with `values`, at
/// `level` when given, else at the neighbour's, when the list shows as `kind` there. At the
/// neighbour's level the list attributes are the neighbour's; at another level the list qualifies
/// only when `source` renders that level, and the attributes are those seeding stores for it.
fn source_attrs(
    values: &BTreeMap<String, Any>,
    kind: ListKind,
    level: Option<f64>,
    source: Option<&SourceMetadata>,
) -> Option<BTreeMap<String, Any>> {
    let num_pr = match (level, values.get("numPr")) {
        (Some(level), Some(Any::Map(num_pr))) => {
            let mut num_pr = (**num_pr).clone();
            let moved = map_number(&num_pr, "ilvl").unwrap_or(0.0) != level;
            num_pr.insert("ilvl".to_owned(), Any::Number(level));
            Some((num_pr, moved))
        }
        _ => None,
    };
    let rendered = match (&num_pr, source) {
        (Some((num_pr, true)), source) => Some(rendered_level(num_pr, source?)?),
        _ => None,
    };
    let from = rendered.as_ref().unwrap_or(values);
    if shown_kind(from) != kind {
        return None;
    }
    let mut attrs: BTreeMap<String, Any> = SOURCE_LIST_KEYS
        .into_iter()
        .map(|key| (key.to_owned(), from.get(key).cloned().unwrap_or(Any::Null)))
        .collect();
    if let Some((num_pr, _)) = num_pr {
        attrs.insert("numPr".to_owned(), Any::from(num_pr));
    }
    attrs.insert(LIST_DEF.to_owned(), Any::Null);
    Some(attrs)
}

/// The list attributes seeding stores for a paragraph on the source list and level `num_pr`
/// names, as `source` renders them; `None` when it renders no such level.
fn rendered_level(
    num_pr: &HashMap<String, Any>,
    source: &SourceMetadata,
) -> Option<BTreeMap<String, Any>> {
    crate::seed::rendered_list_attrs(source_list_rendering(num_pr, source)?)
}

/// The attributes that put a paragraph on the list of the paragraph with `values`, at `level` when
/// given, else at that paragraph's, when that list has that level and shows as `kind` there; a
/// paragraph naming a created list whose definition is absent is on the list its numId names.
fn joined_list(
    values: &BTreeMap<String, Any>,
    definitions: &BTreeMap<String, CreatedListDefinition>,
    kind: ListKind,
    level: Option<f64>,
    source: Option<&SourceMetadata>,
) -> Option<BTreeMap<String, Any>> {
    match resolve_list_identity(values)? {
        ListIdentity::Created { key, level: theirs } if definitions.contains_key(&key) => {
            let level = level.unwrap_or(theirs);
            (definitions[&key].kind == kind).then(|| created_attrs(kind, &key, level))
        }
        _ => source_attrs(values, kind, level, source),
    }
}

/// The attributes a pilcrow holds.
fn pilcrow_values<T: ReadTxn>(pilcrow: &MapRef, txn: &T) -> BTreeMap<String, Any> {
    pilcrow
        .iter(txn)
        .filter_map(|(key, value)| match value {
            Out::Any(value) => Some((key.to_owned(), value)),
            _ => None,
        })
        .collect()
}

impl EditingDoc {
    /// The list definitions the editor created, by key.
    pub fn created_list_definitions(&self) -> Vec<CreatedListDefinition> {
        created_definitions(&self.yrs_doc().transact())
            .into_values()
            .collect()
    }

    /// The keys of the created definitions some paragraph of a live story is on.
    pub fn list_definition_keys_in_use(&self) -> BTreeSet<String> {
        let txn = self.yrs_doc().transact();
        let definitions = created_definitions(&txn);
        let Some(stories) = txn.get_map(crate::STORIES) else {
            return BTreeSet::new();
        };
        crate::live_stories(&stories, &txn)
            .filter_map(|(_, story)| story.cast::<TextRef>().ok())
            .flat_map(|story| crate::pilcrows(&story, &txn))
            .filter_map(|(_, pilcrow)| {
                let values = pilcrow_values(&pilcrow, &txn);
                let ListIdentity::Created { key, .. } = resolve_list_identity(&values)? else {
                    return None;
                };
                definitions.contains_key(&key).then_some(key)
            })
            .collect()
    }

    /// The attributes that put paragraph `para_id` of `story` on a list of `kind`: the list it is
    /// on when that is of that kind, else the list of the paragraph directly above, else below,
    /// when it is of that kind and does not come from that paragraph's style, otherwise a
    /// definition created under a fresh key in [`LIST_DEFINITIONS`]. A paragraph already on a list
    /// keeps its own level, and joins a neighbour's list only when that list has that level in
    /// that kind; one on no list takes the level of the paragraph it joins, or level 0 on a new
    /// list.
    pub fn ensure_list_definition(
        &self,
        ctx: &EditCtx,
        story: &str,
        para_id: &str,
        kind: ListKind,
    ) -> EditResult<BTreeMap<String, Any>> {
        self.ensure_list_definition_after(ctx, story, para_id, kind, None)
    }

    /// As [`Self::ensure_list_definition`] for a paragraph of a selection whose first paragraph
    /// is `lead`: when `lead` is on a list of `kind`, the paragraph goes on that list before any
    /// other is looked at, at its own level where that list has it in that kind, else at
    /// `lead`'s.
    pub fn ensure_list_definition_after(
        &self,
        ctx: &EditCtx,
        story: &str,
        para_id: &str,
        kind: ListKind,
        lead: Option<&str>,
    ) -> EditResult<BTreeMap<String, Any>> {
        let (joined, level) = {
            let txn = self.yrs_doc().transact();
            let text = crate::story_ref(&txn, story)?;
            let pilcrows = crate::pilcrows(&text, &txn);
            let position = |id: &str| {
                pilcrows.iter().position(|(_, pilcrow)| {
                    crate::map_string(pilcrow, &txn, crate::PARA_ID).as_deref() == Some(id)
                })
            };
            let index = position(para_id)
                .ok_or_else(|| EditError::ParagraphNotFound(para_id.to_owned()))?;
            let definitions = created_definitions(&txn);
            let metadata = self.source_metadata();
            let source = metadata.as_deref();
            let own = pilcrow_values(&pilcrows[index].1, &txn);
            let level = numbered_level(&own).map(|(_, level)| level);
            let join = |values: &BTreeMap<String, Any>| {
                joined_list(values, &definitions, kind, level, source)
            };
            let lead_at = match lead {
                Some(lead) if lead != para_id => position(lead),
                _ => None,
            };
            let led = lead_at
                .map(|at| pilcrow_values(&pilcrows[at].1, &txn))
                .and_then(|values| {
                    join(&values)
                        .or_else(|| joined_list(&values, &definitions, kind, None, source))
                });
            let joined = led.or_else(|| join(&own)).or_else(|| {
                [index.checked_sub(1), Some(index + 1)]
                    .into_iter()
                    .flatten()
                    .filter_map(|at| pilcrows.get(at))
                    .map(|(_, pilcrow)| pilcrow_values(pilcrow, &txn))
                    .filter(|values| !numbered_by_style(values))
                    .find_map(|values| join(&values))
            });
            (joined, level)
        };
        if let Some(attrs) = joined {
            return Ok(attrs);
        }
        let mut txn = self.transact_for(ctx);
        let root = txn
            .get_map(LIST_DEFINITIONS)
            .expect("list definitions root is declared");
        let key = self.next_id(&txn);
        let level_texts: Vec<Any> = (0..LEVELS)
            .map(|level| Any::String(kind.level_text(level).into()))
            .collect();
        let definition = HashMap::from([
            ("kind".to_owned(), Any::from(kind.name())),
            ("levelTexts".to_owned(), Any::Array(level_texts.into())),
        ]);
        root.insert(&mut txn, key.clone(), Any::from(definition));
        Ok(created_attrs(kind, &key, level.unwrap_or(0.0)))
    }
}
