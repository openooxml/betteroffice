//! Whole-document PDF export.

use std::sync::PoisonError;

use crate::{Document, Error, Result};

#[cfg(feature = "fonts")]
use betteroffice_fonts::LastResortOffice;

impl Document {
    /// Exports the complete, paginated document as a PDF.
    ///
    /// Layout runs through the same path as [`Document::render_png`], including
    /// font resolution. Every page must have been built; a display list with an
    /// `unbuilt` page is rejected rather than exported partially.
    pub fn export_pdf(&self) -> Result<Vec<u8>> {
        #[cfg(feature = "fonts")]
        self.register_cjk_fallbacks();
        let (display_list, fonts) = self.layout_display_list()?;
        if let Some(page) = display_list.pages.iter().find(|page| page.unbuilt) {
            return Err(Error::Render(format!(
                "page {} was not built; PDF export refuses partial documents",
                page.page_index
            )));
        }
        let images = self.scoped_image_bytes();
        docx_pdf::export_pdf_with_fonts_and_images(&display_list, &fonts.store, &images)
            .map_err(|error| Error::Render(error.to_string()))
    }

    /// Registers script fallback faces for Han/Kana/Hangul before layout.
    ///
    /// The layout chain can pick a Japanese face for a mixed CJK run; adding a
    /// Simplified Chinese face to the store lets the painter resolve the
    /// missing scalars without changing the laid-out chains. Loaded once per
    /// document, and silently skipped when the provider has no such face.
    #[cfg(feature = "fonts")]
    fn register_cjk_fallbacks(&self) {
        let mut fonts = self.fonts.lock().unwrap_or_else(PoisonError::into_inner);
        if fonts.cjk_fallbacks {
            return;
        }
        fonts.cjk_fallbacks = true;
        for script in ["Hani", "Hira", "Kana", "Hang"] {
            let scripts = [script.to_owned()];
            for face in self.font_manager.resolve(
                "Noto Sans SC",
                false,
                false,
                &scripts,
                LastResortOffice::Word,
            ) {
                if let Ok(bytes) = face.load() {
                    let _ = fonts.store.register_slice(&bytes);
                }
            }
        }
    }
}
