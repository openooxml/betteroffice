//! List definitions the editor creates, beside the numbering definitions of the source package.

use std::collections::{BTreeMap, HashMap};

use yrs::{Any, Map, Out, ReadTxn, Transact};

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
const SOURCE_LIST_KEYS: [&str; 9] = [
    "numPr",
    "listIsBullet",
    "listNumFmt",
    "listMarker",
    "listLevelNumFmts",
    "listAbstractNumId",
    "listStartOverride",
    "listMarkerFontFamily",
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

/// The list a paragraph with the attributes `values` is on, or `None` when it is on no list.
pub fn resolve_list_identity(values: &BTreeMap<String, Any>) -> Option<ListIdentity> {
    let Some(Any::Map(num_pr)) = values.get("numPr") else {
        return None;
    };
    let num_id = map_number(num_pr, "numId").filter(|id| *id != 0.0)?;
    let level = map_number(num_pr, "ilvl").unwrap_or(0.0);
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

/// The id the read surfaces give the list the `num_pr` of `values` puts a paragraph on: the key
/// of a created definition, otherwise the `numId` as written; `None` on no list.
pub(crate) fn list_instance_id(
    values: &BTreeMap<String, Any>,
    num_pr: &HashMap<String, Any>,
) -> Option<String> {
    if let Some(key) = created_key(values) {
        return Some(key.to_owned());
    }
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
            let level_texts = texts
                .iter()
                .map(|text| match text {
                    Any::String(text) => Some(text.to_string()),
                    _ => None,
                })
                .collect::<Option<Vec<_>>>()?;
            let definition = CreatedListDefinition {
                key: key.to_owned(),
                kind: ListKind::from_name(kind)?,
                level_texts,
            };
            Some((key.to_owned(), definition))
        })
        .collect()
}

/// The kind of list a paragraph on a list shows.
fn shown_kind(values: &BTreeMap<String, Any>) -> ListKind {
    if values.get("listIsBullet").and_then(any_bool) == Some(true) {
        ListKind::Bullet
    } else {
        ListKind::Numbered
    }
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
        ("listIsBullet".to_owned(), Any::Bool(kind == ListKind::Bullet)),
        ("listNumFmt".to_owned(), Any::from(kind.level_format(shown))),
        ("listMarker".to_owned(), Any::String(marker.into())),
        ("listLevelNumFmts".to_owned(), formats),
        ("listAbstractNumId".to_owned(), Any::Null),
        ("listStartOverride".to_owned(), Any::Null),
    ])
}

/// The attributes that put a paragraph on the source list of its neighbour with `values`.
fn source_attrs(values: &BTreeMap<String, Any>) -> BTreeMap<String, Any> {
    let mut attrs: BTreeMap<String, Any> = SOURCE_LIST_KEYS
        .into_iter()
        .map(|key| (key.to_owned(), values.get(key).cloned().unwrap_or(Any::Null)))
        .collect();
    attrs.insert(LIST_DEF.to_owned(), Any::Null);
    attrs
}

impl EditingDoc {
    /// The list definitions the editor created, by key.
    pub fn created_list_definitions(&self) -> Vec<CreatedListDefinition> {
        created_definitions(&self.yrs_doc().transact())
            .into_values()
            .collect()
    }

    /// The attributes that put paragraph `para_id` of `story` on a list of `kind`: the list of
    /// the paragraph directly above or below when it is of that kind, otherwise a definition
    /// created under a fresh key in [`LIST_DEFINITIONS`].
    pub fn ensure_list_definition(
        &self,
        ctx: &EditCtx,
        story: &str,
        para_id: &str,
        kind: ListKind,
    ) -> EditResult<BTreeMap<String, Any>> {
        let paragraphs = self.paragraphs(story)?;
        let index = paragraphs
            .iter()
            .position(|paragraph| paragraph.para_id == para_id)
            .ok_or_else(|| EditError::ParagraphNotFound(para_id.to_owned()))?;
        let definitions = created_definitions(&self.yrs_doc().transact());
        let above = index.checked_sub(1).and_then(|at| paragraphs.get(at));
        for neighbour in [above, paragraphs.get(index + 1)].into_iter().flatten() {
            let values = &neighbour.properties;
            match resolve_list_identity(values) {
                Some(ListIdentity::Created { key, level }) => {
                    let shown = definitions
                        .get(&key)
                        .map_or_else(|| shown_kind(values), |definition| definition.kind);
                    if shown == kind {
                        return Ok(created_attrs(kind, &key, level));
                    }
                }
                Some(ListIdentity::Source { .. }) if shown_kind(values) == kind => {
                    return Ok(source_attrs(values));
                }
                Some(ListIdentity::Source { .. }) | None => {}
            }
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
        Ok(created_attrs(kind, &key, 0.0))
    }
}
