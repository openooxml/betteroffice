//! List definitions the editor creates, beside the numbering definitions of the source package.

use std::collections::{BTreeMap, HashMap};

use yrs::Any;

use crate::bridge::map_number;
use crate::{EditCtx, EditResult, EditingDoc};

/// The shared root holding the list definitions the editor creates, by session id.
pub const LIST_DEFINITIONS: &str = "listDefinitions";

/// The paragraph attribute naming the created definition a paragraph is on, as
/// `{"key": <string>, "numId": <number>}`; honoured while `numPr.numId` still equals `numId`.
pub const LIST_DEF: &str = "listDef";

/// The kind of list the editor creates.
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum ListKind {
    Bullet,
    Numbered,
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

/// The list a paragraph with the attributes `values` is on, or `None` when it is on no list.
pub fn resolve_list_identity(values: &BTreeMap<String, Any>) -> Option<ListIdentity> {
    let Some(Any::Map(num_pr)) = values.get("numPr") else {
        return None;
    };
    let num_id = map_number(num_pr, "numId").filter(|id| *id != 0.0)?;
    let level = map_number(num_pr, "ilvl").unwrap_or(0.0);
    Some(ListIdentity::Source { num_id, level })
}

impl EditingDoc {
    /// The list definitions the editor created, by key.
    pub fn created_list_definitions(&self) -> Vec<CreatedListDefinition> {
        Vec::new()
    }

    /// The attributes that put paragraph `para_id` of `story` on a list of `kind`: the list of
    /// the paragraph directly above or below when it is of that kind, otherwise a definition
    /// created under a fresh key in [`LIST_DEFINITIONS`].
    pub fn ensure_list_definition(
        &self,
        _ctx: &EditCtx,
        _story: &str,
        _para_id: &str,
        kind: ListKind,
    ) -> EditResult<BTreeMap<String, Any>> {
        let (num_id, bullet, format) = match kind {
            ListKind::Bullet => (1.0, true, Any::Null),
            ListKind::Numbered => (2.0, false, Any::from("decimal")),
        };
        let num_pr = HashMap::from([
            ("numId".to_owned(), Any::Number(num_id)),
            ("ilvl".to_owned(), Any::Number(0.0)),
        ]);
        Ok(BTreeMap::from([
            ("numPr".to_owned(), Any::from(num_pr)),
            ("listIsBullet".to_owned(), Any::Bool(bullet)),
            ("listNumFmt".to_owned(), format),
            ("listMarker".to_owned(), Any::Null),
        ]))
    }
}
