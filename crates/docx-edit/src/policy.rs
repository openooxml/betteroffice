//! Edit policy read from actual owner references: which block control or table cell owns a
//! story, whether a control locks its content, and whether tracked structure owns it.

use std::collections::{HashMap, HashSet};

use yrs::{Any, Map, Out, ReadTxn};

use crate::ops::ChunkKind;
use crate::{DEL, EditingDoc, INS, KIND_KEY, map_string};

/// Owner links followed before a story's ancestry counts as unbounded.
pub(crate) const MAX_OWNERSHIP_DEPTH: usize = 32;

/// The embed owning one child story.
pub(crate) struct Owner {
    pub parent: String,
    /// The `w:lock` value of an owning block content control.
    pub lock: Option<String>,
    /// The owning embed, row or table carries a pending revision.
    pub revision: bool,
}

impl Owner {
    pub fn content_locked(&self) -> bool {
        matches!(
            self.lock.as_deref(),
            Some("contentLocked" | "sdtContentLocked")
        )
    }
}

pub(crate) struct Ownership {
    owners: HashMap<String, Owner>,
}

fn active(value: Option<&Any>) -> bool {
    value.is_some_and(|value| !matches!(value, Any::Null | Any::Undefined))
}

impl Ownership {
    pub fn build<T: ReadTxn>(doc: &EditingDoc, txn: &T) -> Self {
        let mut owners = HashMap::new();
        let Some(stories) = txn.get_map(crate::STORIES) else {
            return Self { owners };
        };
        for (story_id, value) in stories.iter(txn) {
            let Out::YText(story) = value else {
                continue;
            };
            for chunk in doc.chunk_snapshot(story_id, &story, txn).iter() {
                let ChunkKind::Embed(Some(map)) = &chunk.kind else {
                    continue;
                };
                let revision = chunk.attr_active(INS) || chunk.attr_active(DEL);
                match map_string(map, txn, KIND_KEY).as_deref() {
                    Some("table") => {
                        let Some(Out::Any(Any::Array(rows))) = map.get(txn, "rows") else {
                            continue;
                        };
                        for row in rows.iter() {
                            let Any::Map(row) = row else { continue };
                            let row_revision = match row.get("trPr") {
                                Some(Any::Map(properties)) => {
                                    active(properties.get("trIns"))
                                        || active(properties.get("trDel"))
                                }
                                _ => false,
                            };
                            let Some(Any::Array(cells)) = row.get("cells") else {
                                continue;
                            };
                            for cell in cells.iter() {
                                let Any::Map(cell) = cell else { continue };
                                if let Some(Any::String(child)) = cell.get("story") {
                                    owners.insert(
                                        child.to_string(),
                                        Owner {
                                            parent: story_id.to_owned(),
                                            lock: None,
                                            revision: revision || row_revision,
                                        },
                                    );
                                }
                            }
                        }
                    }
                    Some("blockSdt") => {
                        if let Some(child) = map_string(map, txn, "story") {
                            owners.insert(
                                child,
                                Owner {
                                    parent: story_id.to_owned(),
                                    lock: map_string(map, txn, "lock"),
                                    revision,
                                },
                            );
                        }
                    }
                    _ => {}
                }
            }
        }
        Self { owners }
    }

    /// Whether an embed of another story owns `story`.
    pub fn owns(&self, story: &str) -> bool {
        self.owners.contains_key(story)
    }

    /// The owners of `story`, nearest first; errors on a cycle or excessive nesting.
    pub fn chain(&self, story: &str) -> Result<Vec<&Owner>, String> {
        let mut chain = Vec::new();
        let mut seen = HashSet::new();
        let mut current = story;
        while let Some(owner) = self.owners.get(current) {
            if chain.len() == MAX_OWNERSHIP_DEPTH || !seen.insert(current) {
                return Err(format!(
                    "story {story:?} is nested more than {MAX_OWNERSHIP_DEPTH} levels deep"
                ));
            }
            chain.push(owner);
            current = &owner.parent;
        }
        Ok(chain)
    }
}
