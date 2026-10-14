//! Headless VSDX export to PowerPoint and Word.
//!
//! Each page renders through `vsdx-render` into its display list, then maps
//! onto native DrawingML shapes so output stays editable in its host.

mod docx;
mod metadata;
mod pptx;
mod shared;

use vsdx_parse::VsdxPackage;
use vsdx_render::{Renderer, VsdxDisplayList};

pub use metadata::{ShapeDatum, shape_data};

#[derive(Debug)]
pub enum ExportError {
    NoPages,
    Render(String),
    Resolve(String),
    Package(String),
}

impl std::fmt::Display for ExportError {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::NoPages => write!(f, "diagram has no pages"),
            Self::Render(reason) => write!(f, "render failed: {reason}"),
            Self::Resolve(reason) => write!(f, "resolve failed: {reason}"),
            Self::Package(reason) => write!(f, "package failed: {reason}"),
        }
    }
}

impl std::error::Error for ExportError {}

/// Shape ids that degraded during export, by category.
#[derive(Debug, Default)]
pub struct ExportReport {
    pub bbox_fallbacks: Vec<String>,
    pub empty_geometry: Vec<String>,
    pub unsupported_images: Vec<String>,
    pub skipped: Vec<String>,
}

impl ExportReport {
    /// Counts every shape that degraded or was skipped.
    pub fn degraded_shapes(&self) -> usize {
        self.bbox_fallbacks
            .iter()
            .chain(&self.empty_geometry)
            .chain(&self.unsupported_images)
            .chain(&self.skipped)
            .collect::<std::collections::BTreeSet<_>>()
            .len()
    }

    /// One-line caller-facing summary, empty when nothing degraded.
    pub fn summary(&self) -> String {
        let mut parts = Vec::new();
        if !self.bbox_fallbacks.is_empty() {
            parts.push(format!(
                "{} shape{} exported as bounding boxes",
                self.bbox_fallbacks.len(),
                if self.bbox_fallbacks.len() == 1 {
                    ""
                } else {
                    "s"
                }
            ));
        }
        if !self.empty_geometry.is_empty() {
            parts.push(format!(
                "{} shape{} exported as rectangles",
                self.empty_geometry.len(),
                if self.empty_geometry.len() == 1 {
                    ""
                } else {
                    "s"
                }
            ));
        }
        if !self.unsupported_images.is_empty() {
            parts.push(format!(
                "{} image{} unsupported",
                self.unsupported_images.len(),
                if self.unsupported_images.len() == 1 {
                    ""
                } else {
                    "s"
                }
            ));
        }
        if !self.skipped.is_empty() {
            parts.push(format!(
                "{} shape{} skipped",
                self.skipped.len(),
                if self.skipped.len() == 1 { "" } else { "s" }
            ));
        }
        parts.join(", ")
    }
}

/// Exported package bytes with the degradation report.
pub struct ExportOutcome {
    pub bytes: Vec<u8>,
    pub report: ExportReport,
}

/// Exports every page as one slide carrying native shapes.
pub fn export_pptx(package: &VsdxPackage) -> Result<Vec<u8>, ExportError> {
    Ok(export_pptx_with_report(package)?.bytes)
}

/// Exports every page as one slide, reporting what degraded.
pub fn export_pptx_with_report(package: &VsdxPackage) -> Result<ExportOutcome, ExportError> {
    let pages = display_lists(package)?;
    let mut report = ExportReport::default();
    let bytes = pptx::build(&pages, package, &mut report)?;
    Ok(ExportOutcome { bytes, report })
}

/// Exports every page as shapes plus a shape-data table.
pub fn export_docx(package: &VsdxPackage) -> Result<Vec<u8>, ExportError> {
    Ok(export_docx_with_report(package)?.bytes)
}

/// Exports every page as shapes plus a table, reporting what degraded.
pub fn export_docx_with_report(package: &VsdxPackage) -> Result<ExportOutcome, ExportError> {
    let pages = display_lists(package)?;
    let mut report = ExportReport::default();
    let bytes = docx::build(&pages, package, &mut report)?;
    Ok(ExportOutcome { bytes, report })
}

pub(crate) struct Page {
    pub part: String,
    pub name: String,
    pub width_in: f64,
    pub height_in: f64,
    pub list: VsdxDisplayList,
}

fn display_lists(package: &VsdxPackage) -> Result<Vec<Page>, ExportError> {
    if package.page_part_paths.is_empty() {
        return Err(ExportError::NoPages);
    }
    let renderer = Renderer::default();
    let mut pages = Vec::with_capacity(package.page_part_paths.len());
    for part in &package.page_part_paths {
        let list = renderer
            .layout_page(package, part)
            .map_err(|error| ExportError::Render(error.to_string()))?;
        let width_in = f64::from(list.width) / 96.0;
        let height_in = f64::from(list.height) / 96.0;
        pages.push(Page {
            part: part.clone(),
            name: metadata::page_name_for(package, part),
            width_in,
            height_in,
            list,
        });
    }
    Ok(pages)
}
