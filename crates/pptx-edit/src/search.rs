use std::collections::HashSet;

use regex::RegexBuilder;
use serde::Serialize;
use yrs::{Map, TextRef, Transact};

use crate::deck::{
    map_string_array, required_map, required_order, shape_ref, slide_ref, slide_shape_order,
    string_array_ref,
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
        let pattern = RegexBuilder::new(&regex::escape(query))
            .case_insensitive(!case_sensitive)
            .build()
            .map_err(|error| EditError::InvalidText(error.to_string()))?;
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
            let mut shapes = string_array_ref(&slide_shape_order(&slide, &txn)?, &txn);
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
                    for found in pattern.find_iter(&text) {
                        position += text[byte_offset..found.start()].encode_utf16().count() as u32;
                        let end = position + found.as_str().encode_utf16().count() as u32;
                        matches.push(TextSearchMatch {
                            slide_index,
                            slide_id: slide_id.clone(),
                            shape_id: shape_id.clone(),
                            story_id: story_id.clone(),
                            start: position,
                            end,
                            text: found.as_str().to_owned(),
                        });
                        if matches.len() == limit {
                            return Ok(matches);
                        }
                        byte_offset = found.end();
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
}
