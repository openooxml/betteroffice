use docx_parse::paragraph::Paragraph;

use super::{
    ProjectionContext, ProjectionError, ProjectionFamily, ProjectionHooks, StoryPosition,
    StoryProjection,
};
use crate::ParagraphProperties;

pub(super) struct Hooks;

#[derive(Default)]
pub(super) struct StoryData {
    section_ranges_deferred: bool,
}

impl ProjectionHooks for Hooks {
    fn family(&self) -> ProjectionFamily {
        ProjectionFamily::StoriesAux
    }

    fn project_root(
        &self,
        _context: &mut ProjectionContext<'_, '_>,
        story_id: &str,
    ) -> Result<StoryProjection, ProjectionError> {
        let kind = if story_id.starts_with("hf:") { "headerFooterStory" } else { "noteStory" };
        let mut result = StoryProjection::default();
        result.defer(self.family(), kind, story_id, None);
        Ok(result)
    }

    fn finish_paragraph(
        &self,
        _context: &mut ProjectionContext<'_, '_>,
        story: &mut StoryPosition<'_>,
        properties: Option<&ParagraphProperties>,
        _source: Option<&Paragraph>,
        paragraph: &mut Paragraph,
    ) -> Result<(), ProjectionError> {
        if paragraph.section_properties.is_some() || properties.is_some_and(|properties| {
            properties.values.contains_key("sectPr") || properties.values.contains_key("sectionBreakType")
        }) {
            story.state.aux.section_ranges_deferred = true;
        }
        Ok(())
    }

    fn finish_story(
        &self,
        context: &mut ProjectionContext<'_, '_>,
        story: &mut StoryPosition<'_>,
    ) -> Result<(), ProjectionError> {
        if story.id == "body"
            && (story.state.aux.section_ranges_deferred
                || context.source.envelope.document.package.document.sections.is_some()
                || context.source.envelope.document.package.document.final_section_properties.is_some())
        {
            story.result.defer(self.family(), "sectionRanges", story.id, None);
        }
        Ok(())
    }
}
