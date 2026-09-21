use std::borrow::Cow;
use std::collections::{BTreeMap, HashSet};
use std::fmt;

use serde::Serialize;
use yrs::{Any, Map, ReadTxn, Transact};

use crate::{EditingDoc, SegmentContent, read_state::table_cell_stories};

/// Paragraph-local UTF-16 offsets.
#[derive(Clone, Debug, Eq, PartialEq, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct TextSearchMatch {
    pub story: String,
    pub para_id: String,
    pub start: u32,
    pub end: u32,
    pub text: String,
}

#[derive(Debug)]
pub struct TextSearchError(String);

impl fmt::Display for TextSearchError {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        self.0.fmt(f)
    }
}

impl std::error::Error for TextSearchError {}

/// Literal query; case-insensitive matching equates characters that share an uppercase form.
struct Needle {
    folded: String,
    case_sensitive: bool,
}

impl Needle {
    fn new(query: &str, case_sensitive: bool) -> Self {
        let folded = if case_sensitive {
            query.to_owned()
        } else {
            query.chars().map(fold_char).collect()
        };
        Self {
            folded,
            case_sensitive,
        }
    }

    /// Lazily yields the non-overlapping byte ranges of `text` equal to the query, in order.
    fn find_all<'a>(&'a self, text: &'a str) -> Matches<'a> {
        let (haystack, origins) = if self.case_sensitive {
            (Cow::Borrowed(text), Vec::new())
        } else {
            let mut folded = String::with_capacity(text.len());
            let mut origins = Vec::with_capacity(text.len() + 1);
            for (offset, ch) in text.char_indices() {
                origins.push((folded.len(), offset));
                folded.push(fold_char(ch));
            }
            origins.push((folded.len(), text.len()));
            (Cow::Owned(folded), origins)
        };
        Matches {
            needle: &self.folded,
            haystack,
            origins,
            cursor: 0,
        }
    }
}

fn fold_char(ch: char) -> char {
    let mut upper = ch.to_uppercase();
    if let (Some(single), None) = (upper.next(), upper.next()) {
        let mut lower = single.to_lowercase();
        if let (Some(single), None) = (lower.next(), lower.next()) {
            return single;
        }
    }
    let mut lower = ch.to_lowercase();
    match (lower.next(), lower.next()) {
        (Some(single), None) => single,
        _ => ch,
    }
}

/// Byte ranges in the original text; empty `origins` means the haystack is that text.
struct Matches<'a> {
    needle: &'a str,
    haystack: Cow<'a, str>,
    origins: Vec<(usize, usize)>,
    cursor: usize,
}

impl Iterator for Matches<'_> {
    type Item = (usize, usize);

    fn next(&mut self) -> Option<Self::Item> {
        if self.needle.is_empty() {
            return None;
        }
        let start = self.cursor + self.haystack[self.cursor..].find(self.needle)?;
        let end = start + self.needle.len();
        self.cursor = end;
        Some((self.original(start), self.original(end)))
    }
}

impl Matches<'_> {
    fn original(&self, folded: usize) -> usize {
        if self.origins.is_empty() {
            return folded;
        }
        match self.origins.binary_search_by_key(&folded, |&(at, _)| at) {
            Ok(index) => self.origins[index].1,
            Err(index) => self.origins[index.min(self.origins.len() - 1)].1,
        }
    }
}

enum SearchPart {
    Text { offset: u32, text: String },
    Story(String),
}

enum SearchTask {
    Story(String),
    Text {
        story: String,
        para_id: String,
        offset: u32,
        text: String,
    },
}

fn cell_stories(payload: &BTreeMap<String, Any>) -> Vec<String> {
    let mut stories = Vec::new();
    if let Some(Any::Array(rows)) = payload.get("rows") {
        for row in rows.iter() {
            let Any::Map(row) = row else { continue };
            let Some(Any::Array(cells)) = row.get("cells") else {
                continue;
            };
            for cell in cells.iter() {
                let Any::Map(cell) = cell else { continue };
                if let Some(Any::String(story)) = cell.get("story") {
                    stories.push(story.to_string());
                }
            }
        }
    }
    stories
}

impl EditingDoc {
    /// Searches body and table content in reading order, then other stories by ID.
    pub fn search_text(
        &self,
        query: &str,
        case_sensitive: bool,
        limit: Option<usize>,
    ) -> Result<Vec<TextSearchMatch>, TextSearchError> {
        let limit = limit.unwrap_or(usize::MAX);
        if query.is_empty() || limit == 0 {
            return Ok(Vec::new());
        }
        let needle = Needle::new(query, case_sensitive);
        let txn = self.yrs_doc().transact();
        let Some(stories) = txn.get_map(crate::STORIES) else {
            return Ok(Vec::new());
        };
        let cells = table_cell_stories(self, &txn);
        let mut ids: Vec<String> = stories.keys(&txn).map(str::to_owned).collect();
        ids.sort_by(|a, b| {
            (a != "body", cells.contains(a), a).cmp(&(b != "body", cells.contains(b), b))
        });
        let mut tasks: Vec<_> = ids.into_iter().rev().map(SearchTask::Story).collect();
        let mut visited = HashSet::new();
        let mut matches = Vec::new();
        while let Some(task) = tasks.pop() {
            match task {
                SearchTask::Story(story) => {
                    if !visited.insert(story.clone()) {
                        continue;
                    }
                    let mut ordered = Vec::new();
                    let mut parts = Vec::new();
                    let mut offset = 0;
                    let mut text = String::new();
                    for segment in self
                        .story_segments(&story)
                        .map_err(|error| TextSearchError(error.to_string()))?
                    {
                        if let SegmentContent::Text(value) = segment.content {
                            text.push_str(&value);
                            continue;
                        }
                        if !text.is_empty() {
                            let length = text.encode_utf16().count() as u32;
                            parts.push(SearchPart::Text {
                                offset,
                                text: std::mem::take(&mut text),
                            });
                            offset += length;
                        }
                        match segment.content {
                            SegmentContent::Pilcrow(paragraph) => {
                                for part in parts.drain(..) {
                                    ordered.push(match part {
                                        SearchPart::Text { offset, text } => SearchTask::Text {
                                            story: story.clone(),
                                            para_id: paragraph.para_id.clone(),
                                            offset,
                                            text,
                                        },
                                        SearchPart::Story(story) => SearchTask::Story(story),
                                    });
                                }
                                offset = 0;
                            }
                            SegmentContent::OtherEmbed { kind, payload } => {
                                if kind == "table" {
                                    parts.extend(
                                        cell_stories(&payload).into_iter().map(SearchPart::Story),
                                    );
                                }
                                offset += 1;
                            }
                            SegmentContent::Text(_) => unreachable!(),
                        }
                    }
                    tasks.extend(ordered.into_iter().rev());
                }
                SearchTask::Text {
                    story,
                    para_id,
                    offset,
                    text,
                } => {
                    let mut byte_offset = 0;
                    let mut position = offset;
                    for (from, to) in needle.find_all(&text) {
                        position += text[byte_offset..from].encode_utf16().count() as u32;
                        let found = &text[from..to];
                        let end = position + found.encode_utf16().count() as u32;
                        matches.push(TextSearchMatch {
                            story: story.clone(),
                            para_id: para_id.clone(),
                            start: position,
                            end,
                            text: found.to_owned(),
                        });
                        if matches.len() == limit {
                            return Ok(matches);
                        }
                        byte_offset = to;
                        position = end;
                    }
                }
            }
        }
        Ok(matches)
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::{EditCtx, FormatPolicy, Position};

    #[test]
    fn literal_unicode_offsets_limits_and_read_only() {
        let doc = EditingDoc::new(83001);
        let para_id = doc
            .create_story("body", "😀Σςσ [a] [A]", "Normal", "left")
            .unwrap();
        let before = doc.encode_state_as_update_v1();
        let matches = doc.search_text("σ", false, None).unwrap();
        assert_eq!(
            matches.iter().map(|m| (m.start, m.end)).collect::<Vec<_>>(),
            [(2, 3), (3, 4), (4, 5)]
        );
        assert!(matches.iter().all(|m| m.para_id == para_id));
        assert_eq!(
            doc.search_text("[a]", false, Some(1)).unwrap()[0].text,
            "[a]"
        );
        assert_eq!(doc.search_text("[a]", true, None).unwrap().len(), 1);
        assert!(doc.search_text("", false, None).unwrap().is_empty());
        assert!(doc.search_text("σ", false, Some(0)).unwrap().is_empty());
        assert_eq!(doc.encode_state_as_update_v1(), before);
    }

    #[test]
    fn tables_and_nested_cells_follow_document_order() {
        let doc = EditingDoc::new(83002);
        let ctx = EditCtx::local("test", "2026-09-18T00:00:00Z");
        doc.create_story("body", "needle before needle after", "Normal", "left")
            .unwrap();
        doc.split_paragraph(&ctx, Position::new("body", 14), None)
            .unwrap();
        let table = doc
            .insert_table(&ctx, Position::new("body", 14), 12, 1)
            .unwrap();
        for (index, story) in table.created_story_ids.iter().enumerate() {
            doc.insert_text(
                &ctx,
                Position::new(story, 0),
                &format!("needle row{index}"),
                FormatPolicy::Inherit,
            )
            .unwrap();
        }
        let nested = doc
            .insert_table(&ctx, Position::new(&table.created_story_ids[0], 11), 1, 1)
            .unwrap();
        doc.insert_text(
            &ctx,
            Position::new(&nested.created_story_ids[0], 0),
            "needle nested",
            FormatPolicy::Inherit,
        )
        .unwrap();
        doc.create_story("header:rId1", "needle header", "Normal", "left")
            .unwrap();
        let matches = doc.search_text("needle", false, None).unwrap();
        let mut expected = vec![
            "body".to_owned(),
            table.created_story_ids[0].clone(),
            nested.created_story_ids[0].clone(),
        ];
        expected.extend(table.created_story_ids[1..].iter().cloned());
        expected.extend(["body".to_owned(), "header:rId1".to_owned()]);
        assert_eq!(
            matches.iter().map(|m| m.story.clone()).collect::<Vec<_>>(),
            expected
        );
        assert_eq!(
            doc.search_text("needle", false, Some(3)).unwrap(),
            matches[..3]
        );
        assert_eq!(matches[matches.len() - 2].start, 0);
    }

    #[test]
    fn needle_equates_shared_uppercase_forms_and_keeps_byte_ranges() {
        let text = "\u{212A}elvin k STRASSE ẞ ß AAA";
        let ranges = |query, case_sensitive| {
            Needle::new(query, case_sensitive)
                .find_all(text)
                .collect::<Vec<_>>()
        };
        assert_eq!(ranges("k", false), [(0, 3), (9, 10)]);
        assert_eq!(ranges("k", true), [(9, 10)]);
        assert_eq!(ranges("ß", false), [(19, 22), (23, 25)]);
        assert_eq!(ranges("ss", false), [(15, 17)]);
        assert_eq!(ranges("aa", false), [(26, 28)]);
        assert_eq!(Needle::new("", false).find_all(text).next(), None);
    }
}
