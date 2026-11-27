use docx_parse::paragraph::Paragraph;
use yrs::Any;

use super::{
    Attrs, Lowered, ProjectionContext, ProjectionError, ProjectionFamily, ProjectionHooks,
    StoryPosition, runs::AuthoredRuns,
};
use crate::{ParagraphProperties, SegmentContent, StorySegment};

pub(super) struct Hooks;
pub(super) type StoryData = ();

pub(super) const MARKS: &[&str] = &["fieldResult", "hyperlink"];

pub(super) fn defer_boundary_content(
    story: &mut StoryPosition<'_>,
    attrs: &Attrs,
    authored: &AuthoredRuns<'_>,
) {
    let paragraph = Some(story.state.paragraph_index);
    if matches!(attrs.get("pageBreakBeforeRun"), Some(Any::Bool(true))) {
        story
            .result
            .defer(ProjectionFamily::Fields, "pageBreak", story.id, paragraph);
    }
    if authored.has_note_marks() {
        story
            .result
            .defer(ProjectionFamily::Fields, "noteMarks", story.id, paragraph);
    }
}

impl ProjectionHooks for Hooks {
    fn family(&self) -> ProjectionFamily {
        ProjectionFamily::Fields
    }

    fn wrap(
        &self,
        _context: &mut ProjectionContext<'_, '_>,
        story: &mut StoryPosition<'_>,
        segment: &StorySegment,
        _lowered: &mut Lowered,
    ) -> Result<(), ProjectionError> {
        if matches!(segment.content, SegmentContent::Text(_)) {
            return Ok(());
        }
        for mark in MARKS
            .iter()
            .copied()
            .filter(|mark| super::story::has_mark(segment, mark))
        {
            story.defer_here(self.family(), mark);
        }
        Ok(())
    }

    fn finish_paragraph(
        &self,
        _context: &mut ProjectionContext<'_, '_>,
        story: &mut StoryPosition<'_>,
        _properties: Option<&ParagraphProperties>,
        source: Option<&Paragraph>,
        paragraph: &mut Paragraph,
    ) -> Result<(), ProjectionError> {
        if let Some(rendered_break) = source.and_then(|source| source.rendered_page_break_before) {
            story.defer_here(self.family(), "pageBreak");
            paragraph.rendered_page_break_before = Some(rendered_break);
        }
        Ok(())
    }
}
