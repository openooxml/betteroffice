//! Server-side rasterization of a display list.

use std::sync::{Mutex, PoisonError};

use docx_layout::display_list::DisplayList;
use docx_raster::{FontChains, ImageMap, RenderResources};
use ooxml_text::FontStore;
use serde_json::Number;

use crate::{Document, Error, Result};

const MAX_FONT_BYTES: usize = 32 * 1024 * 1024;
const MAX_FONTS: usize = 256;
pub const MAX_PIXMAP_DIM: u32 = 16_384;
pub const MAX_PIXMAP_PIXELS: u64 = 16_777_216;

/// Registered faces plus the `family|bold|italic` chains the raster backend
/// resolves runs against. The store sits behind a `Mutex` so a [`Document`]
/// holding one stays `Send + Sync`.
#[derive(Default)]
pub(crate) struct FontRegistry {
    store: Mutex<FontStore>,
    chains: FontChains,
    faces: usize,
}

impl FontRegistry {
    fn register(&mut self, family: &str, bold: bool, italic: bool, bytes: &[u8]) -> Result<u32> {
        if bytes.len() > MAX_FONT_BYTES {
            return Err(Error::ResourceLimit(format!(
                "font exceeds {MAX_FONT_BYTES} bytes"
            )));
        }
        if self.faces >= MAX_FONTS {
            return Err(Error::ResourceLimit(format!(
                "more than {MAX_FONTS} font faces"
            )));
        }
        let family = family.trim();
        if family.is_empty() {
            return Err(Error::Font("font family is empty".to_owned()));
        }
        let id = self
            .store
            .get_mut()
            .unwrap_or_else(PoisonError::into_inner)
            .register(bytes.to_vec())
            .map_err(|error| Error::Font(error.to_string()))?;
        self.chains
            .entry(chain_key(family, bold, italic))
            .or_default()
            .push(id);
        self.faces += 1;
        Ok(id.to_u32())
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

fn chain_key(family: &str, bold: bool, italic: bool) -> String {
    format!(
        "{}|{}|{}",
        family.to_lowercase(),
        u8::from(bold),
        u8::from(italic)
    )
}

impl Document {
    /// Registers one font face for rendering, returning its font id. Faces are
    /// capped at 256 and each font at 32 MiB; either budget is a
    /// [`Error::ResourceLimit`] rather than a truncated accept.
    pub fn register_font(
        &mut self,
        family: &str,
        bold: bool,
        italic: bool,
        bytes: &[u8],
    ) -> Result<u32> {
        self.fonts.register(family, bold, italic, bytes)
    }

    /// Rasterizes one display-list page to deterministic PNG bytes.
    ///
    /// A page wider or taller than [`MAX_PIXMAP_DIM`], or larger in area than
    /// [`MAX_PIXMAP_PIXELS`], is refused before any surface is allocated.
    ///
    /// Embedded images arrive as `data:` URLs on the primitive, so no image
    /// side table is consulted.
    pub fn render_png(&self, display_list: &DisplayList, page_ordinal: usize) -> Result<Vec<u8>> {
        if let Some(page) = display_list.pages.get(page_ordinal)
            && let (Some(width), Some(height)) =
                (page_pixels(&page.width), page_pixels(&page.height))
        {
            validate_render_size(width, height)?;
        }
        let store = self
            .fonts
            .store
            .lock()
            .unwrap_or_else(PoisonError::into_inner);
        let images = ImageMap::new();
        let resources = RenderResources::new(&store, &self.fonts.chains, &images);
        docx_raster::render_png(display_list, page_ordinal, &resources).map_err(Error::Render)
    }
}
