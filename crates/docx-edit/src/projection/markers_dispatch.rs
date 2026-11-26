use docx_parse::paragraph::Paragraph;
use yrs::{Any, Map, ReadTxn};

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
        let source = context.source.stories.get(story.id);
        if source.is_some_and(|source| source.blocks.iter().any(super::story::has_comment_ranges))
            || context.txn.get_map(crate::COMMENTS).is_some_and(|comments| {
            comments.keys(context.txn).any(|id| {
                context.doc.comment_anchors_in(context.txn, id)
                    .is_ok_and(|anchors| anchors.iter().any(|anchor| anchor.story == story.id))
            })
        }) {
            story.result.defer(self.family(), "commentRanges", story.id, None);
        }
        story.state.markers.raw_mark_offsets = context.doc.resolve_raw_marks_in(context.txn)
            .into_iter().filter(|mark| mark.story == story.id)
            .map(|mark| mark.offset as usize).collect();
        if source.is_some_and(|source| source.blocks.iter().any(super::story::has_revision_ranges)) {
            story.result.defer(self.family(), "revisions", story.id, None);
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
        for kind in MARKS.iter().copied().filter(|kind| super::story::has_mark(segment, kind)) {
            story.result.defer(self.family(), kind, story.id, Some(story.state.paragraph_index));
        }
        Ok(())
    }

    fn finish_paragraph(
        &self,
        _context: &mut ProjectionContext<'_, '_>,
        story: &mut StoryPosition<'_>,
        props: Option<&ParagraphProperties>,
        source: Option<&Paragraph>,
        _paragraph: &mut Paragraph,
    ) -> Result<(), ProjectionError> {
        let index = Some(story.state.paragraph_index);
        for kind in ["pPrIns", "pPrDel", "pPrChange", "bookmarks"] {
            if props.and_then(|props| props.values.get(kind)).is_some_and(|value| !matches!(value, Any::Null | Any::Undefined)) {
                story.result.defer(self.family(), kind, story.id, index);
            }
        }
        if story.state.markers.raw_mark_offsets.iter()
            .any(|offset| (story.state.paragraph_start..=story.state.offset).contains(offset))
            || source.is_some_and(|source| source.content.iter().any(|inline| inline.node_type() == "rawXml"))
        {
            story.result.defer(self.family(), "rawInline", story.id, index);
        }
        Ok(())
    }
}
