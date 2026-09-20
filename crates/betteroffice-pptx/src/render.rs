//! Server-side rasterization of a laid-out slide.

use std::sync::{Mutex, PoisonError};

use pptx_raster::{AssetMap, GlyphCache, ImageCache, RenderResources};

use crate::{Error, Presentation, Result};

pub use pptx_raster::{
    Background, MAX_IMAGE_BYTES, MAX_IMAGE_PIXELS, MAX_SLIDE_DIM, MAX_SLIDE_PIXELS, RenderOptions,
    RenderedSlide as RenderedPng,
};

/// Glyph outlines and decoded assets shared across every slide one deck
/// exports. Behind `Mutex`es so a [`Presentation`] holding one stays
/// `Send + Sync`.
#[derive(Default)]
pub(crate) struct RenderCaches {
    glyphs: Mutex<GlyphCache>,
    images: Mutex<ImageCache>,
}

impl Presentation {
    /// Rasterizes one slide to deterministic PNG bytes. Media resolves from the
    /// package, so nothing needs registering beyond the fonts
    /// [`Presentation::register_font`] took; a picture the backend cannot draw
    /// is skipped and counted rather than failing the render. Glyph outlines and
    /// decoded images are cached on the deck.
    pub fn render_png(&self, slide_index: usize, options: &RenderOptions) -> Result<RenderedPng> {
        let rendered = self.render_slide(slide_index)?;
        let mut images: AssetMap<'_> = self
            .media()
            .iter()
            .map(|part| (part.part_path.as_str(), part.bytes.as_slice()))
            .collect();
        let pending: Vec<(&str, Vec<u8>)> = rendered
            .display_list
            .primitives
            .iter()
            .filter_map(|primitive| match primitive {
                pptx_render::Primitive::Image {
                    asset_id: Some(id), ..
                } if id.starts_with("pending-media:") => Some(id.as_str()),
                _ => None,
            })
            .map(|id| self.media_bytes(id).map(|bytes| (id, bytes)))
            .collect::<Result<_>>()?;
        images.extend(pending.iter().map(|(id, bytes)| (*id, bytes.as_slice())));
        let resources = RenderResources::new(self.renderer().fonts(), &images)
            .with_label_font(self.renderer().fallback_font());
        let mut glyphs = self
            .caches()
            .glyphs
            .lock()
            .unwrap_or_else(PoisonError::into_inner);
        let mut cached_images = self
            .caches()
            .images
            .lock()
            .unwrap_or_else(PoisonError::into_inner);
        pptx_raster::render_slide_shared(
            &rendered.display_list,
            &resources,
            options,
            &mut glyphs,
            &mut cached_images,
        )
        .map_err(Error::Raster)
    }
}
