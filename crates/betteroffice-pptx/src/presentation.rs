use pptx_edit::{
    DeckSession, DeckSnapshot, EditCtx, ShapeDraft, ShapeReceipt, SlideReceipt, StorySnapshot,
    TextReceipt, TextStyle, TextStylePatch, TransformReceipt,
};
use pptx_parse::{
    MediaPart, ParseLimits, PptxPackage, Presentation as PresentationModel, Slide, SlideLayout,
    SlideMaster, ThemePart,
};
use pptx_render::{RenderedSlide, SlideRenderer};

use crate::Result;

const STANDALONE_CLIENT_ID: u64 = 1;

pub struct Presentation {
    session: DeckSession,
    renderer: SlideRenderer,
}

impl Presentation {
    pub fn open(bytes: &[u8]) -> Result<Self> {
        Self::open_with_limits_internal(bytes, &ParseLimits::default(), STANDALONE_CLIENT_ID)
    }

    pub fn open_with_limits(bytes: &[u8], limits: &ParseLimits) -> Result<Self> {
        Self::open_with_limits_internal(bytes, limits, STANDALONE_CLIENT_ID)
    }

    pub fn open_collaborative(bytes: &[u8], client_id: u64) -> Result<Self> {
        Self::open_with_limits_internal(bytes, &ParseLimits::default(), client_id)
    }

    pub fn open_collaborative_with_limits(
        bytes: &[u8],
        client_id: u64,
        limits: &ParseLimits,
    ) -> Result<Self> {
        Self::open_with_limits_internal(bytes, limits, client_id)
    }

    fn open_with_limits_internal(
        bytes: &[u8],
        limits: &ParseLimits,
        client_id: u64,
    ) -> Result<Self> {
        let package = pptx_parse::parse_pptx_with_limits(bytes, limits)?;
        let session = DeckSession::from_package(package, client_id)?;
        Ok(Self {
            session,
            renderer: SlideRenderer::new(),
        })
    }

    pub fn client_id(&self) -> u64 {
        self.session.client_id()
    }

    pub fn package(&self) -> &PptxPackage {
        self.session.package()
    }

    pub fn model(&self) -> &PresentationModel {
        &self.package().presentation
    }

    pub fn slides(&self) -> &[Slide] {
        &self.package().slides
    }

    pub fn layouts(&self) -> &[SlideLayout] {
        &self.package().layouts
    }

    pub fn masters(&self) -> &[SlideMaster] {
        &self.package().masters
    }

    pub fn themes(&self) -> &[ThemePart] {
        &self.package().themes
    }

    pub fn media(&self) -> &[MediaPart] {
        &self.package().media
    }

    pub fn snapshot(&self) -> Result<DeckSnapshot> {
        Ok(self.session.snapshot()?)
    }

    pub fn story(&self, story_id: &str) -> Result<StorySnapshot> {
        Ok(self.session.story(story_id)?)
    }

    pub fn insert_slide(
        &self,
        context: &EditCtx,
        index: u32,
        layout_part_path: Option<&str>,
    ) -> Result<SlideReceipt> {
        Ok(self
            .session
            .insert_slide(context, index, layout_part_path)?)
    }

    pub fn delete_slide(&self, context: &EditCtx, slide_id: &str) -> Result<SlideReceipt> {
        Ok(self.session.delete_slide(context, slide_id)?)
    }

    pub fn move_slide(
        &self,
        context: &EditCtx,
        slide_id: &str,
        to_index: u32,
    ) -> Result<SlideReceipt> {
        Ok(self.session.move_slide(context, slide_id, to_index)?)
    }

    pub fn add_text_box(
        &self,
        context: &EditCtx,
        slide_id: &str,
        draft: &ShapeDraft,
    ) -> Result<ShapeReceipt> {
        Ok(self.session.add_text_box(context, slide_id, draft)?)
    }

    pub fn remove_shape(
        &self,
        context: &EditCtx,
        slide_id: &str,
        shape_id: &str,
    ) -> Result<ShapeReceipt> {
        Ok(self.session.remove_shape(context, slide_id, shape_id)?)
    }

    pub fn move_shape(
        &self,
        context: &EditCtx,
        slide_id: &str,
        shape_id: &str,
        x: i64,
        y: i64,
    ) -> Result<TransformReceipt> {
        Ok(self.session.move_shape(context, slide_id, shape_id, x, y)?)
    }

    pub fn resize_shape(
        &self,
        context: &EditCtx,
        slide_id: &str,
        shape_id: &str,
        width: i64,
        height: i64,
    ) -> Result<TransformReceipt> {
        Ok(self
            .session
            .resize_shape(context, slide_id, shape_id, width, height)?)
    }

    pub fn insert_text(
        &self,
        context: &EditCtx,
        story_id: &str,
        index: u32,
        text: &str,
        style: &TextStyle,
    ) -> Result<TextReceipt> {
        Ok(self
            .session
            .insert_text(context, story_id, index, text, style)?)
    }

    pub fn delete_text(
        &self,
        context: &EditCtx,
        story_id: &str,
        start: u32,
        end: u32,
    ) -> Result<TextReceipt> {
        Ok(self.session.delete_text(context, story_id, start, end)?)
    }

    pub fn format_text(
        &self,
        context: &EditCtx,
        story_id: &str,
        start: u32,
        end: u32,
        patch: &TextStylePatch,
    ) -> Result<TextReceipt> {
        Ok(self
            .session
            .format_text(context, story_id, start, end, patch)?)
    }

    pub fn insert_paragraph_break(
        &self,
        context: &EditCtx,
        story_id: &str,
        index: u32,
    ) -> Result<TextReceipt> {
        Ok(self
            .session
            .insert_paragraph_break(context, story_id, index)?)
    }

    pub fn register_font(
        &mut self,
        family: &str,
        bold: bool,
        italic: bool,
        bytes: &[u8],
    ) -> Result<u32> {
        Ok(self.renderer.register_font(family, bold, italic, bytes)?)
    }

    pub fn render_slide(&self, slide_index: usize) -> Result<RenderedSlide> {
        let snapshot = self.session.snapshot()?;
        Ok(self
            .renderer
            .layout_slide(self.session.package(), &snapshot, slide_index)?)
    }

    /// Serializes the byte-preserved source package.
    pub fn save(&self) -> Result<Vec<u8>> {
        Ok(pptx_parse::write_pptx(self.session.package())?)
    }

    pub fn encode_state_vector_v1(&self) -> Vec<u8> {
        self.session.encode_state_vector_v1()
    }

    pub fn encode_state_as_update_v1(&self) -> Vec<u8> {
        self.session.encode_state_as_update_v1()
    }

    pub fn encode_diff_v1(&self, remote_state_vector: &[u8]) -> Result<Vec<u8>> {
        Ok(self.session.encode_diff_v1(remote_state_vector)?)
    }

    pub fn apply_update_v1(&self, update: &[u8]) -> Result<DeckSnapshot> {
        Ok(self.session.apply_update_v1(update)?)
    }

    pub fn undo(&self) -> bool {
        self.session.undo()
    }

    pub fn redo(&self) -> bool {
        self.session.redo()
    }

    pub fn can_undo(&self) -> bool {
        self.session.can_undo()
    }

    pub fn can_redo(&self) -> bool {
        self.session.can_redo()
    }

    pub fn add_undo_barrier(&self) {
        self.session.add_undo_barrier();
    }
}
