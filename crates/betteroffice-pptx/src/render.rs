//! Server-side rasterization of a laid-out slide.

use std::sync::{Mutex, PoisonError};

use pptx_raster::{AssetMap, GlyphCache, RenderResources};

use crate::{Error, Presentation, Result};

pub use pptx_raster::{
    Background, MAX_IMAGE_BYTES, MAX_IMAGE_PIXELS, MAX_SLIDE_DIM, MAX_SLIDE_PIXELS, RenderOptions,
    RenderedSlide as RenderedPng,
};

/// Glyph outlines shared across every slide one deck exports. Behind a `Mutex`
/// so a [`Presentation`] holding one stays `Send + Sync`.
#[derive(Default)]
pub(crate) struct GlyphRegistry {
    cache: Mutex<GlyphCache>,
}

impl Presentation {
    /// Rasterizes one slide to deterministic PNG bytes. Media resolves from the
    /// package, so nothing needs registering beyond the fonts
    /// [`Presentation::register_font`] took; a picture the backend cannot draw
    /// is skipped and counted rather than failing the render. Glyph outlines are
    /// cached on the deck, so later slides reuse what earlier ones extracted.
    pub fn render_png(&self, slide_index: usize, options: &RenderOptions) -> Result<RenderedPng> {
        let rendered = self.render_slide(slide_index)?;
        let images: AssetMap<'_> = self
            .media()
            .iter()
            .map(|part| (part.part_path.as_str(), part.bytes.as_slice()))
            .collect();
        let resources = RenderResources::new(self.renderer().fonts(), &images)
            .with_label_font(self.renderer().fallback_font());
        let mut cache = self
            .glyphs()
            .cache
            .lock()
            .unwrap_or_else(PoisonError::into_inner);
        pptx_raster::render_slide_cached(&rendered.display_list, &resources, options, &mut cache)
            .map_err(Error::Raster)
    }
}
