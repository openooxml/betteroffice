//! Server-side rasterization of a display list.

use std::sync::{Mutex, PoisonError};

use docx_layout::display_list::DisplayList;
use docx_raster::{
    ImageCache, ImageMap, ImageScope, RenderResources, RenderedPage, scoped_image_key,
};
use serde_json::Number;

use crate::{Document, Error, Result};

const MAX_IMAGE_BYTES: usize = 32 * 1024 * 1024;
const MAX_IMAGES: usize = 256;
pub const MAX_PIXMAP_DIM: u32 = docx_raster::MAX_PAGE_DIM;
pub const MAX_PIXMAP_PIXELS: u64 = docx_raster::MAX_PAGE_PIXELS;

/// Caller-supplied bytes for relationship ids the display list did not already
/// carry as `data:` URLs, keyed by owning part, plus the decoded images a page
/// leaves for the next — behind a `Mutex` so the document stays `Send + Sync`.
#[derive(Default)]
pub(crate) struct ImageRegistry {
    entries: ImageMap,
    cache: Mutex<ImageCache>,
}

impl ImageRegistry {
    fn register(&mut self, scope: ImageScope<'_>, rel_id: &str, bytes: &[u8]) -> Result<()> {
        if bytes.len() > MAX_IMAGE_BYTES {
            return Err(Error::ResourceLimit(format!(
                "image exceeds {MAX_IMAGE_BYTES} bytes"
            )));
        }
        if rel_id.is_empty() {
            return Err(Error::Image("image relationship id is empty".to_owned()));
        }
        let key = scoped_image_key(scope, rel_id);
        if self.entries.len() >= MAX_IMAGES && !self.entries.contains_key(&key) {
            return Err(Error::ResourceLimit(format!(
                "more than {MAX_IMAGES} images"
            )));
        }
        self.entries.insert(key, bytes.to_vec());
        Ok(())
    }
}

fn validate_render_size(width: u32, height: u32) -> Result<()> {
    if width > MAX_PIXMAP_DIM || height > MAX_PIXMAP_DIM {
        return Err(Error::RenderTooLarge {
            width,
            height,
            max: MAX_PIXMAP_DIM,
        });
    }
    if u64::from(width) * u64::from(height) > MAX_PIXMAP_PIXELS {
        return Err(Error::RenderAreaTooLarge {
            width,
            height,
            max_pixels: MAX_PIXMAP_PIXELS,
        });
    }
    Ok(())
}

/// The pixel extent the raster backend would allocate, or `None` when the
/// backend rejects the number before allocating anything.
fn page_pixels(number: &Number) -> Option<u32> {
    let value = number.as_f64()?;
    if !value.is_finite() || value <= 0.0 {
        return None;
    }
    Some(value.ceil().min(f64::from(u32::MAX)) as u32)
}

impl Document {
    /// Supplies bytes for one relationship id the display list left
    /// unresolved, scoped to the part that owns the id. Images are capped at
    /// 256 and each at 32 MiB.
    pub fn register_image(
        &mut self,
        scope: ImageScope<'_>,
        rel_id: &str,
        bytes: &[u8],
    ) -> Result<()> {
        self.images.register(scope, rel_id, bytes)
    }

    /// Rasterizes one display-list page to deterministic PNG bytes. A page past
    /// [`MAX_PIXMAP_DIM`] or [`MAX_PIXMAP_PIXELS`] is refused before any surface
    /// is allocated; an image the backend cannot draw is skipped and counted.
    /// Glyph outlines and decoded images are cached on the document, so a page
    /// reuses what earlier pages extracted and decoded; an export past either
    /// cache's cap re-does what it evicted.
    pub fn render_png(
        &self,
        display_list: &DisplayList,
        page_ordinal: usize,
    ) -> Result<RenderedPage> {
        if let Some(page) = display_list.pages.get(page_ordinal)
            && let (Some(width), Some(height)) =
                (page_pixels(&page.width), page_pixels(&page.height))
        {
            validate_render_size(width, height)?;
        }
        #[cfg(feature = "fonts")]
        self.prepare_display_fonts(display_list, page_ordinal)?;
        let mut fonts = self.fonts.lock().unwrap_or_else(PoisonError::into_inner);
        let mut images = self
            .images
            .cache
            .lock()
            .unwrap_or_else(PoisonError::into_inner);
        let crate::font::FontRegistry {
            store,
            chains,
            glyphs,
            ..
        } = &mut *fonts;
        let resources = RenderResources::new(store, chains, &self.images.entries);
        docx_raster::render_page_cached(display_list, page_ordinal, &resources, glyphs, &mut images)
            .map_err(Error::Render)
    }
}
