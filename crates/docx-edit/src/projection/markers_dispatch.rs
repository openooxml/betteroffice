use docx_parse::paragraph::Paragraph;
use yrs::Any;

use super::{
    Lowered, ProjectionContext, ProjectionError, ProjectionFamily, ProjectionHooks, StoryPosition,
};
use crate::{ParagraphProperties, StorySegment};

pub(super) struct Hooks;
pub(super) const MARKS: &[&str] = &["ins", "del", "rPrChange"];

#[derive(Default)]
pub(super) struct StoryData {
    raw_mark_offsets: Vec<usize>,
}

impl ProjectionHooks for Hooks {
    fn family(&self) -> ProjectionFamily {
        ProjectionFamily::Markers
    }

    fn setup_story(
        &self,
        context: &mut ProjectionContext<'_, '_>,
        story: &mut StoryPosition<'_>,
        _segments: &[StorySegment],
    ) -> Result<(), ProjectionError> {
        if context.source.has_comment_ranges || context.markers.has_comments(story.id) {
            story
                .result
                .defer(self.family(), "commentRanges", story.id, None);
        }
        story.state.markers.raw_mark_offsets = context.markers.raw_offsets(context.txn, story.id);
        if context.source.has_revision_ranges {
            story
                .result
                .defer(self.family(), "revisions", story.id, None);
        }
        Ok(())
    }

    fn wrap(
        &self,
        _context: &mut ProjectionContext<'_, '_>,
        story: &mut StoryPosition<'_>,
        segment: &StorySegment,
        _content: &mut Lowered,
    ) -> Result<(), ProjectionError> {
        for kind in MARKS
            .iter()
            .copied()
            .filter(|kind| super::story::has_mark(segment, kind))
        {
            story.result.defer(
                self.family(),
                kind,
                story.id,
                Some(story.state.paragraph_index),
            );
        }
        Ok(())
    }

    fn finish_paragraph(
        &self,
        _context: &mut ProjectionContext<'_, '_>,
        story: &mut StoryPosition<'_>,
        props: Option<&ParagraphProperties>,
        source: Option<&Paragraph>,
        paragraph: &mut Paragraph,
    ) -> Result<(), ProjectionError> {
        let index = Some(story.state.paragraph_index);
        let live = |kind| {
            props
                .and_then(|props| props.values.get(kind))
                .is_some_and(|value| match value {
                    Any::Null | Any::Undefined => false,
                    Any::Array(entries) => !entries.is_empty(),
                    _ => true,
                })
        };
        for (kind, marker) in [
            ("pPrIns", &mut paragraph.p_pr_ins),
            ("pPrDel", &mut paragraph.p_pr_del),
        ] {
            if live(kind) {
                story.result.defer(self.family(), kind, story.id, index);
            } else {
                *marker = None;
            }
        }
        if live("pPrChange") {
            story
                .result
                .defer(self.family(), "pPrChange", story.id, index);
        } else {
            paragraph.property_changes = None;
        }
        if live("bookmarks") {
            story
                .result
                .defer(self.family(), "bookmarks", story.id, index);
        } else {
            paragraph
                .content
                .retain(|entry| !matches!(entry.node_type(), "bookmarkStart" | "bookmarkEnd"));
        }
        if story.state.markers.raw_mark_offsets[story
            .state
            .markers
            .raw_mark_offsets
            .partition_point(|offset| *offset < story.state.paragraph_start)..]
            .first()
            .is_some_and(|offset| *offset <= story.state.offset)
            || source.is_some_and(|source| {
                source
                    .content
                    .iter()
                    .any(|inline| inline.node_type() == "rawXml")
            })
        {
            story
                .result
                .defer(self.family(), "rawInline", story.id, index);
        }
        Ok(())
    }
}
