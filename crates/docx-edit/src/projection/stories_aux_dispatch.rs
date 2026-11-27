use docx_parse::BlockContent;
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

pub(super) fn project_root(story_id: &str, blocks: &[BlockContent]) -> StoryProjection {
    let kind = if story_id.starts_with("hf:") {
        "headerFooterStory"
    } else {
        "noteStory"
    };
    let mut result = StoryProjection {
        story_id: story_id.to_owned(),
        blocks: blocks.to_vec(),
        ..StoryProjection::default()
    };
    result.defer(ProjectionFamily::StoriesAux, kind, story_id, None);
    result
}

impl ProjectionHooks for Hooks {
    fn family(&self) -> ProjectionFamily {
        ProjectionFamily::StoriesAux
    }

    fn finish_paragraph(
        &self,
        _context: &mut ProjectionContext<'_, '_>,
        story: &mut StoryPosition<'_>,
        properties: Option<&ParagraphProperties>,
        source: Option<&Paragraph>,
        paragraph: &mut Paragraph,
    ) -> Result<(), ProjectionError> {
        if matches!(
            properties.and_then(|properties| crate::resolve_list_identity(&properties.values)),
            Some(crate::ListIdentity::Created { .. })
        ) {
            story.result.defer(
                self.family(),
                crate::LIST_DEF,
                story.id,
                Some(story.state.paragraph_index),
            );
        }
        if let Some(section) = source.and_then(|source| source.section_properties.as_ref()) {
            story.defer_here(self.family(), "sectionRanges");
            paragraph.section_properties = Some(section.clone());
        }
        if paragraph.section_properties.is_some()
            || properties.is_some_and(|properties| {
                properties.values.contains_key("sectPr")
                    || properties.values.contains_key("sectionBreakType")
            })
        {
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
            && (story.state.aux.section_ranges_deferred || context.package_body_has_sections)
        {
            story
                .result
                .defer(self.family(), "sectionRanges", story.id, None);
        }
        Ok(())
    }
}
