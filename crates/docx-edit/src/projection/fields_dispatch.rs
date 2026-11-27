use yrs::Any;

use super::{
    Attrs, Lowered, ProjectionContext, ProjectionError, ProjectionFamily, ProjectionHooks,
    StoryPosition, runs::AuthoredRuns,
};
use crate::StorySegment;

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
        for mark in MARKS
            .iter()
            .copied()
            .filter(|mark| super::story::has_mark(segment, mark))
        {
            story.result.defer(
                self.family(),
                mark,
                story.id,
                Some(story.state.paragraph_index),
            );
        }
        Ok(())
    }
}
