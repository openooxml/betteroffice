use std::collections::{HashMap, HashSet};

use serde::{Deserialize, Serialize};
use yrs::{Map, ReadTxn, Transact};

use crate::EditingDoc;
use crate::batch::DocumentVersion;
use crate::structured::source::{cell_table, story_root};
use crate::structured::{StoryKind, StoryUse, header_footer_index};
use crate::target::Views;

/// What a session story holds. Comment bodies are not session stories, so there is no comment
/// kind.
#[derive(Clone, Copy, Debug, Deserialize, Eq, Hash, PartialEq, Serialize)]
#[serde(rename_all = "kebab-case")]
pub enum StoryInfoKind {
    Body,
    TableCell,
    ContentControl,
    Header,
    Footer,
    Footnote,
    Endnote,
    Other,
}

impl StoryInfoKind {
    fn parse(name: &str) -> Option<Self> {
        Some(match name {
            "body" => Self::Body,
            "table-cell" => Self::TableCell,
            "content-control" => Self::ContentControl,
            "header" => Self::Header,
            "footer" => Self::Footer,
            "footnote" => Self::Footnote,
            "endnote" => Self::Endnote,
            "other" => Self::Other,
            _ => return None,
        })
    }
}

#[derive(Clone, Debug, Deserialize, Eq, PartialEq, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct StoryInfo {
    pub story: String,
    pub kind: StoryInfoKind,
    /// The story a table cell or block content control sits in.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub parent: Option<String>,
    /// The top-level story this one belongs to; itself for a top-level story.
    pub root: String,
    /// A header's or footer's package part, when the session was opened from DOCX bytes.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub part: Option<String>,
    /// A header's or footer's sections whose properties reference the part, inheritance
    /// applied, whether or not a page shows it; a preview's may be incomplete.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub uses: Option<Vec<StoryUse>>,
}

#[derive(Clone, Debug, Deserialize, Eq, PartialEq, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct ListStoriesResponse {
    pub version: DocumentVersion,
    pub stories: Vec<StoryInfo>,
}

/// `"all"`, or story kinds and story ids in any mix.
#[derive(Clone, Debug, Deserialize, Eq, PartialEq, Serialize)]
#[serde(untagged)]
pub enum StorySelection {
    All(AllStories),
    Listed(Vec<String>),
}

#[derive(Clone, Copy, Debug, Deserialize, Eq, PartialEq, Serialize)]
#[serde(rename_all = "camelCase")]
pub enum AllStories {
    All,
}

fn nesting(story: &str) -> Option<(&str, StoryInfoKind)> {
    let digits = |text: &str| !text.is_empty() && text.bytes().all(|byte| byte.is_ascii_digit());
    if let Some(table) = cell_table(story)
        && let Some(cut) = table.rfind(":t")
        && digits(&table[cut + 2..])
    {
        return Some((&table[..cut], StoryInfoKind::TableCell));
    }
    let cut = story.rfind(":sdt")?;
    digits(&story[cut + 4..]).then(|| (&story[..cut], StoryInfoKind::ContentControl))
}

/// Every story of `views`, sorted by id.
pub(crate) fn story_infos<T: ReadTxn>(views: &mut Views<'_, T>) -> Vec<StoryInfo> {
    let txn = views.txn();
    let mut ids: Vec<String> = txn
        .get_map(crate::STORIES)
        .map(|stories| stories.keys(txn).map(str::to_owned).collect())
        .unwrap_or_default();
    ids.sort();
    let index = header_footer_index(views);
    ids.into_iter()
        .map(|story| {
            let root = story_root(&story).to_owned();
            if let Some((parent, kind)) = nesting(&story) {
                let parent = Some(parent.to_owned());
                return StoryInfo {
                    story,
                    kind,
                    parent,
                    root,
                    part: None,
                    uses: None,
                };
            }
            let (kind, part, uses) = match index.roles.get(&story) {
                Some((role, part)) => (
                    if *role == StoryKind::Header {
                        StoryInfoKind::Header
                    } else {
                        StoryInfoKind::Footer
                    },
                    part.clone(),
                    Some(index.uses.get(&story).cloned().unwrap_or_default()),
                ),
                None if story == "body" => (StoryInfoKind::Body, None, None),
                None if story.starts_with("fn:") => (StoryInfoKind::Footnote, None, None),
                None if story.starts_with("en:") => (StoryInfoKind::Endnote, None, None),
                None => (StoryInfoKind::Other, None, None),
            };
            StoryInfo {
                story,
                kind,
                parent: None,
                root,
                part,
                uses,
            }
        })
        .collect()
}

/// The selected story ids: every story sorted, or the listed entries in order, a kind standing
/// for its stories in id order, without repeats. `by_root` matches a kind against each story's
/// root.
pub(crate) fn select_stories<T: ReadTxn>(
    views: &mut Views<'_, T>,
    selection: &StorySelection,
    by_root: bool,
) -> Vec<String> {
    let entries = match selection {
        StorySelection::All(_) => {
            return story_infos(views)
                .into_iter()
                .map(|info| info.story)
                .collect();
        }
        StorySelection::Listed(entries) => entries,
    };
    let infos = entries
        .iter()
        .any(|entry| StoryInfoKind::parse(entry).is_some())
        .then(|| {
            let infos = story_infos(views);
            let kinds: HashMap<String, StoryInfoKind> = infos
                .iter()
                .map(|info| (info.story.clone(), info.kind))
                .collect();
            infos
                .into_iter()
                .map(|info| {
                    let kind = if by_root {
                        kinds.get(&info.root).copied().unwrap_or(info.kind)
                    } else {
                        info.kind
                    };
                    (info.story, kind)
                })
                .collect::<Vec<_>>()
        })
        .unwrap_or_default();
    let mut seen = HashSet::new();
    let mut selected = Vec::new();
    for entry in entries {
        match StoryInfoKind::parse(entry) {
            Some(kind) => {
                for (story, _) in infos.iter().filter(|(_, matched)| *matched == kind) {
                    if seen.insert(story.as_str()) {
                        selected.push(story.clone());
                    }
                }
            }
            None => {
                if seen.insert(entry.as_str()) {
                    selected.push(entry.clone());
                }
            }
        }
    }
    selected
}

impl EditingDoc {
    /// Every story with its kind, container and, for a header or footer, its part and uses.
    pub fn list_stories(&self) -> ListStoriesResponse {
        let version = self.version();
        let txn = self.yrs_doc().transact();
        let mut views = Views::committed(self, &txn);
        ListStoriesResponse {
            version,
            stories: story_infos(&mut views),
        }
    }
}

#[cfg(test)]
mod tests {
    use yrs::Any;

    use super::*;
    use crate::{EditCtx, RawOp};

    #[test]
    fn without_a_source_section_references_name_headers_and_footers() {
        let doc = EditingDoc::new(100);
        doc.create_story("body", "x", "Normal", "left").unwrap();
        for story in ["hf:rIdH", "hf:rIdF", "hf:rIdX"] {
            doc.create_story(story, "y", "Normal", "left").unwrap();
        }
        let section = Any::from_json(
            r#"{"headerReferences":[{"type":"default","rId":"rIdH"}],"footerReferences":[{"type":"first","rId":"rIdF"}]}"#,
        )
        .unwrap();
        doc.apply_raw_ops(
            "body",
            vec![RawOp::SetEmbedAttr {
                index: 1,
                key: "sectPr".to_owned(),
                value: section,
            }],
            &EditCtx::local("", ""),
        )
        .unwrap();
        let listed = serde_json::to_value(doc.list_stories().stories).unwrap();
        assert_eq!(
            listed,
            serde_json::json!([
                { "story": "body", "kind": "body", "root": "body" },
                {
                    "story": "hf:rIdF", "kind": "footer", "root": "hf:rIdF",
                    "uses": [
                        { "sectionIndex": 0, "variant": "first" },
                        { "sectionIndex": 1, "variant": "first" },
                    ],
                },
                {
                    "story": "hf:rIdH", "kind": "header", "root": "hf:rIdH",
                    "uses": [
                        { "sectionIndex": 0, "variant": "default" },
                        { "sectionIndex": 1, "variant": "default" },
                    ],
                },
                { "story": "hf:rIdX", "kind": "other", "root": "hf:rIdX" },
            ])
        );
    }
}
