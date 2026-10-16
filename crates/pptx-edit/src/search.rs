use std::borrow::Cow;
use std::collections::HashSet;

use serde::Serialize;
use yrs::{Map, TextRef, Transact};

use crate::deck::{
    live_shape_order, map_string_array, required_map, required_order, shape_ref, slide_ref,
    slide_shape_order, string_array_ref,
};
use crate::{DeckSession, EditError, EditResult, STORIES, story::snapshot_story};

/// Story-local UTF-16 offsets.
#[derive(Clone, Debug, Eq, PartialEq, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct TextSearchMatch {
    pub slide_index: usize,
    pub slide_id: String,
    pub shape_id: String,
    pub story_id: String,
    pub start: u32,
    pub end: u32,
    pub text: String,
}

/// Literal query with Unicode simple case-insensitive matching.
pub(crate) struct Needle {
    folded: String,
    case_sensitive: bool,
}

impl Needle {
    pub(crate) fn new(query: &str, case_sensitive: bool) -> Self {
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
        self.matches(text, false)
    }

    /// Like [`Needle::find_all`], but every occurrence counts, overlapping ones included.
    pub(crate) fn find_overlapping<'a>(&'a self, text: &'a str) -> Matches<'a> {
        self.matches(text, true)
    }

    fn matches<'a>(&'a self, text: &'a str, overlapping: bool) -> Matches<'a> {
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
            overlapping,
        }
    }
}

fn fold_char(ch: char) -> char {
    match ch {
        '\u{0131}' => return ch,
        '\u{1FD3}' => return '\u{0390}',
        '\u{1FE3}' => return '\u{03B0}',
        '\u{FB05}' => return '\u{FB06}',
        _ => {}
    }
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
pub(crate) struct Matches<'a> {
    needle: &'a str,
    haystack: Cow<'a, str>,
    origins: Vec<(usize, usize)>,
    cursor: usize,
    /// Resume one character after a match's start rather than at its end.
    overlapping: bool,
}

impl Iterator for Matches<'_> {
    type Item = (usize, usize);

    fn next(&mut self) -> Option<Self::Item> {
        if self.needle.is_empty() {
            return None;
        }
        let start = self.cursor + self.haystack[self.cursor..].find(self.needle)?;
        let end = start + self.needle.len();
        self.cursor = if self.overlapping {
            start
                + self.haystack[start..]
                    .chars()
                    .next()
                    .map_or(1, char::len_utf8)
        } else {
            end
        };
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

impl DeckSession {
    /// Searches slides, shapes, and table cells in document order.
    pub fn search_text(
        &self,
        query: &str,
        case_sensitive: bool,
        limit: Option<usize>,
    ) -> EditResult<Vec<TextSearchMatch>> {
        let limit = limit.unwrap_or(usize::MAX);
        if query.is_empty() || limit == 0 {
            return Ok(Vec::new());
        }
        let needle = Needle::new(query, case_sensitive);
        let txn = self.doc.transact();
        let stories = required_map(&txn, STORIES)?;
        let mut seen_slides = HashSet::new();
        let mut matches = Vec::new();
        for (slide_index, slide_id) in string_array_ref(&required_order(&txn)?, &txn)
            .into_iter()
            .filter(|id| seen_slides.insert(id.clone()))
            .enumerate()
        {
            let slide = slide_ref(&txn, &slide_id)?;
            let mut shapes = live_shape_order(&slide_shape_order(&slide, &txn)?, &txn)?;
            shapes.reverse();
            let mut seen_shapes = HashSet::new();
            while let Some(shape_id) = shapes.pop() {
                if !seen_shapes.insert(shape_id.clone()) {
                    continue;
                }
                let shape = shape_ref(&txn, &shape_id)?;
                for story_id in map_string_array(&shape, &txn, "textStories")? {
                    let story = stories
                        .get(&txn, &story_id)
                        .and_then(|value| value.cast::<TextRef>().ok())
                        .ok_or_else(|| EditError::StoryNotFound(story_id.clone()))?;
                    let text = snapshot_story(&story, &txn, &story_id)?.plain_text();
                    let mut byte_offset = 0;
                    let mut position = 0;
                    for (from, to) in needle.find_all(&text) {
                        position += text[byte_offset..from].encode_utf16().count() as u32;
                        let found = &text[from..to];
                        let end = position + found.encode_utf16().count() as u32;
                        matches.push(TextSearchMatch {
                            slide_index,
                            slide_id: slide_id.clone(),
                            shape_id: shape_id.clone(),
                            story_id: story_id.clone(),
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
                shapes.extend(
                    map_string_array(&shape, &txn, "children")?
                        .into_iter()
                        .rev(),
                );
            }
        }
        Ok(matches)
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::{EditCtx, ShapeDraft, ShapeRect, TextStyle};

    #[test]
    fn literal_unicode_offsets_paragraphs_limits_and_read_only() {
        let session = DeckSession::open(
            include_bytes!("../../../apps/demo/public/betteroffice-demo.pptx"),
            84001,
        )
        .unwrap();
        let slide = session.snapshot().unwrap().slides[0].id.clone();
        let added = session
            .add_text_box(
                &EditCtx::local("test"),
                &slide,
                &ShapeDraft {
                    name: "Search".to_owned(),
                    rect: ShapeRect {
                        x: 0,
                        y: 0,
                        width: 1_000_000,
                        height: 1_000_000,
                    },
                    text: "😀Σςσ [a]\n[A]".to_owned(),
                    style: TextStyle::default(),
                },
            )
            .unwrap();
        let before = session.encode_state_as_update_v1();
        let matches = session.search_text("σ", false, None).unwrap();
        assert_eq!(
            matches.iter().map(|m| (m.start, m.end)).collect::<Vec<_>>(),
            [(2, 3), (3, 4), (4, 5)]
        );
        assert!(matches.iter().all(|m| m.shape_id == added.shape_id));
        assert_eq!(
            session
                .search_text("[a]", false, None)
                .unwrap()
                .iter()
                .map(|m| m.start)
                .collect::<Vec<_>>(),
            [6, 10]
        );
        assert_eq!(session.search_text("[a]", true, None).unwrap().len(), 1);
        assert_eq!(
            session.search_text("σ", false, Some(1)).unwrap(),
            matches[..1]
        );
        assert!(session.search_text("", false, None).unwrap().is_empty());
        assert!(session.search_text("σ", false, Some(0)).unwrap().is_empty());
        assert_eq!(session.encode_state_as_update_v1(), before);
    }

    #[test]
    fn searches_nested_group_shapes() {
        let session = DeckSession::open(
            include_bytes!("../tests/fixtures/deck-schema-v2-nested-connectors.pptx"),
            84002,
        )
        .unwrap();
        let snapshot = session.snapshot().unwrap();
        let slide = &snapshot.slides[0];
        let mut shapes: Vec<_> = slide.shapes.iter().collect();
        let mut nested_story = None;
        while let Some(shape) = shapes.pop() {
            for child in &shape.children {
                if let Some(story) = child
                    .text_stories
                    .iter()
                    .find(|story| story.plain_text().contains("Before"))
                {
                    nested_story = Some((child.id.clone(), story.id.clone()));
                }
                shapes.push(child);
            }
        }
        let (shape_id, story_id) = nested_story.expect("fixture has nested text");
        let matches = session.search_text("before", false, None).unwrap();
        assert_eq!(matches.len(), 1);
        assert_eq!(matches[0].shape_id, shape_id);
        assert_eq!(matches[0].story_id, story_id);
        assert_eq!(matches[0].slide_id, slide.id);
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

    #[test]
    fn overlapping_matches_count_every_occurrence() {
        let needle = Needle::new("aa", true);
        assert_eq!(needle.find_all("aaa").collect::<Vec<_>>(), [(0, 2)]);
        assert_eq!(
            needle.find_overlapping("aaa").collect::<Vec<_>>(),
            [(0, 2), (1, 3)]
        );
        let emoji = Needle::new("😀😀", true);
        assert_eq!(
            emoji.find_overlapping("😀😀😀").collect::<Vec<_>>(),
            [(0, 8), (4, 12)]
        );
    }
}
