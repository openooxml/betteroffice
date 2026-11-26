use super::{
    Lowered, ProjectionContext, ProjectionError, ProjectionFamily, ProjectionHooks, StoryPosition,
};
use crate::StorySegment;

pub(super) struct Hooks;
pub(super) type StoryData = ();

pub(super) const MARKS: &[&str] = &["fieldResult", "hyperlink"];

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
