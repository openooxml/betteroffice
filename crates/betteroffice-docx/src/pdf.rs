//! Whole-document PDF export.

use crate::{Document, Error, Result};

impl Document {
    /// Exports the complete, paginated document as a PDF.
    ///
    /// Layout runs through the same path as [`Document::render_png`], including
    /// font resolution. Every page must have been built; a display list with an
    /// `unbuilt` page is rejected rather than exported partially.
    pub fn export_pdf(&self) -> Result<Vec<u8>> {
        let (display_list, fonts) = self.layout_display_list()?;
        if let Some(page) = display_list.pages.iter().find(|page| page.unbuilt) {
            return Err(Error::Render(format!(
                "page {} was not built; PDF export refuses partial documents",
                page.page_index
            )));
        }
        let images = self.image_bytes_by_rel_id();
        docx_pdf::export_pdf_with_fonts_and_images(&display_list, &fonts.store, &images)
            .map_err(|error| Error::Render(error.to_string()))
    }
}
