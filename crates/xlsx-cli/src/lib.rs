//! Library path behind the `xlsx` binary.

use betteroffice_xlsx::{SheetId, Workbook};

pub use betteroffice_xlsx::{MAX_PIXMAP_DIM, RenderOptions, RenderedPng};

/// one-line description of a sheet for the `info` command.
#[derive(Debug, Clone)]
pub struct SheetSummary {
    pub index: u32,
    pub name: String,
    /// a1 used range, or `None` for an empty sheet.
    pub used_range: Option<String>,
    pub cell_count: usize,
}

pub fn load_workbook(bytes: &[u8]) -> Result<Workbook, String> {
    Workbook::open_for_read(bytes).map_err(|error| error.to_string())
}

/// resolve a `--sheet` selector: name match wins, then 0-based index;
/// `None` selects the first sheet.
pub fn resolve_sheet(wb: &Workbook, selector: Option<&str>) -> Result<SheetId, String> {
    let Some(sel) = selector else {
        return Ok(SheetId(0));
    };
    if let Some(id) = wb.sheet_id(sel) {
        return Ok(id);
    }
    if let Ok(idx) = sel.parse::<usize>() {
        if idx < wb.sheet_count() {
            return Ok(SheetId(idx as u32));
        }
        return Err(format!(
            "sheet index {idx} out of range (workbook has {} sheet(s))",
            wb.sheet_count()
        ));
    }
    Err(format!("no sheet named {sel:?}"))
}

pub fn render(wb: &Workbook, sheet: SheetId, opts: &RenderOptions) -> Result<RenderedPng, String> {
    wb.render_sheet(sheet, opts)
        .map_err(|error| error.to_string())
}

/// one summary per sheet, in workbook order.
pub fn sheet_summaries(wb: &Workbook) -> Vec<SheetSummary> {
    wb.model()
        .sheets
        .iter()
        .enumerate()
        .map(|(i, s)| SheetSummary {
            index: i as u32,
            name: s.name.clone(),
            used_range: s.used_range().map(|r| r.to_a1()),
            cell_count: s.iter_cells().count(),
        })
        .collect()
}
