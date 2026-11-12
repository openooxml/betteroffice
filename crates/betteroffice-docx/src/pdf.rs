//! Whole-document PDF export.

use crate::{Document, Error, PdfExportOptions, Result};

#[cfg(feature = "fonts")]
use crate::PdfConformance;
#[cfg(feature = "fonts")]
use betteroffice_fonts::LastResortOffice;

impl Document {
    /// Exports the complete, paginated document as a plain PDF 1.7.
    ///
    /// Equivalent to [`Document::export_pdf_with_options`] with
    /// [`PdfExportOptions::default`], so the historical bytes are unchanged.
    /// Layout runs through the same path as [`Document::render_png`], including
    /// font resolution. Every page must have been built; a display list with an
    /// `unbuilt` page is rejected rather than exported partially.
    pub fn export_pdf(&self) -> Result<Vec<u8>> {
        self.export_pdf_with_options(&PdfExportOptions::default())
    }

    /// Exports the complete, paginated document under explicit options.
    ///
    /// `options.conformance` selects the output profile.
    /// [`PdfConformance::Standard`](docx_pdf::PdfConformance::Standard) keeps
    /// the historical plain PDF 1.7 output, including its Base-14 text
    /// fallback. [`PdfConformance::PdfA2u`](docx_pdf::PdfConformance::PdfA2u)
    /// targets ISO 19005-2 (PDF/A-2u): it requires embedded fonts and returns
    /// an error for content the writer cannot represent conformantly (for
    /// example text that would need the Base-14 fallback) instead of silently
    /// downgrading. `options.metadata` populates `/Info` and XMP; absent fields
    /// are omitted rather than defaulted.
    ///
    /// Layout, CJK fallback registration, image scoping and the unbuilt-page
    /// refusal are the same as [`Document::export_pdf`].
    pub fn export_pdf_with_options(&self, options: &PdfExportOptions) -> Result<Vec<u8>> {
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
        // A PDF/A painter may not emit `.notdef`, so characters outside every
        // document, CJK and bundled face need the monochrome emoji fallback.
        // Standard keeps its historical face priority and never registers it.
        // The face is export-local: remove it from the cached registry after
        // both successful and rejected exports so a later Standard export sees
        // exactly the same store as before.
        #[cfg(feature = "fonts")]
        let pdfa_store_len = if options.conformance == PdfConformance::PdfA2u {
            let len = fonts.store.font_count();
            fonts.register_pdfa_fallbacks()?;
            Some(len)
        } else {
            None
        };
        let images = self.scoped_image_bytes();
        let result = docx_pdf::export_pdf_with_options(
            &display_list,
            Some(&fonts.store),
            &images,
            Some(fonts.chains()),
            options,
        )
        .map_err(|error| Error::Render(error.to_string()));
        #[cfg(feature = "fonts")]
        if let Some(len) = pdfa_store_len {
            fonts.release_pdfa_fallbacks(len);
        }
        result
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
