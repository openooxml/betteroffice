//! Whole-document PDF export.

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
        let (display_list, fonts) = self.layout_display_list()?;
        #[cfg(feature = "fonts")]
        let mut fonts = fonts;
        #[cfg(feature = "fonts")]
        self.register_cjk_fallbacks(&mut fonts)?;
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

    /// Registers script fallback faces for Han/Kana/Hangul after layout.
    ///
    /// The layout chain can pick a Japanese face for a mixed CJK run; adding a
    /// Simplified Chinese face to the store lets the painter resolve the
    /// missing scalars without changing the laid-out chains. Registering after
    /// layout keeps the document's own faces ahead of the fallbacks in painter
    /// resolution order. Loaded once per document, and silently skipped when
    /// the provider has no such face.
    #[cfg(feature = "fonts")]
    fn register_cjk_fallbacks(&self, fonts: &mut crate::font::FontRegistry) -> Result<()> {
        if fonts.cjk_fallbacks {
            return Ok(());
        }
        let mut seen = std::collections::HashSet::new();
        for script in ["cjk-sc", "cjk-tc", "cjk-jp", "cjk-kr"] {
            let scripts = [script.to_owned()];
            for face in self.font_manager.resolve(
                "Noto Sans SC",
                false,
                false,
                &scripts,
                LastResortOffice::Word,
            ) {
                if !seen.insert((face.family.clone(), face.bold, face.italic)) {
                    continue;
                }
                if let Ok(bytes) = face.load() {
                    fonts.register_export_fallback(&bytes)?;
                }
            }
        }
        fonts.cjk_fallbacks = true;
        Ok(())
    }
}
