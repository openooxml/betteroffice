//! library path behind the `xlsx` binary: load a workbook, resolve a sheet,
//! rasterize a region to png. every entry point returns `Result<_, String>`.

use xlsx_model::{CellRange, SheetId, Workbook};
use xlsx_raster::render_png;
use xlsx_render::{build_display_list, scaled, viewport_for_range, viewport_for_used_range};

/// hard cap on either output dimension in pixels; oversized ranges are
/// rejected before the pixmap is allocated.
pub const MAX_PIXMAP_DIM: u32 = 16_384;

/// knobs for one render: region (else the sheet's used range), output scale,
/// and pixel caps that crop the content region.
#[derive(Debug, Clone)]
pub struct RenderOptions {
    pub range: Option<CellRange>,
    pub scale: f32,
    pub max_width: Option<u32>,
    pub max_height: Option<u32>,
}

impl Default for RenderOptions {
    fn default() -> Self {
        Self {
            range: None,
            scale: 1.0,
            max_width: None,
            max_height: None,
        }
    }
}

/// a rendered frame: the encoded png plus its pixel dimensions.
#[derive(Debug, Clone)]
pub struct RenderedPng {
    pub bytes: Vec<u8>,
    pub width: u32,
    pub height: u32,
}

/// one-line description of a sheet for the `info` command.
#[derive(Debug, Clone)]
pub struct SheetSummary {
    pub index: u32,
    pub name: String,
    /// a1 used range, or `None` for an empty sheet.
    pub used_range: Option<String>,
    pub cell_count: usize,
}

/// unzip and parse raw `.xlsx` bytes into a workbook, refusing an empty one.
pub fn load_workbook(bytes: &[u8]) -> Result<Workbook, String> {
    let parts = ooxml_opc::unzip_parts(bytes)?;
    let workbook = xlsx_parse::parse_workbook(&parts).map_err(|e| e.to_string())?;
    if workbook.sheets.is_empty() {
        return Err("workbook has no sheets".to_string());
    }
    Ok(workbook)
}

/// resolve a `--sheet` selector: name match wins, then 0-based index;
/// `None` selects the first sheet.
pub fn resolve_sheet(wb: &Workbook, selector: Option<&str>) -> Result<SheetId, String> {
    let Some(sel) = selector else {
        return Ok(SheetId(0));
    };
    if let Some((id, _)) = wb.sheet_by_name(sel) {
        return Ok(id);
    }
    if let Ok(idx) = sel.parse::<usize>() {
        if idx < wb.sheets.len() {
            return Ok(SheetId(idx as u32));
        }
        return Err(format!(
            "sheet index {idx} out of range (workbook has {} sheet(s))",
            wb.sheets.len()
        ));
    }
    Err(format!("no sheet named {sel:?}"))
}

/// build the display list for the chosen region, scale it, and rasterize to
/// png; the pixmap-dimension guard runs before anything is allocated.
pub fn render(wb: &Workbook, sheet: SheetId, opts: &RenderOptions) -> Result<RenderedPng, String> {
    if !(opts.scale.is_finite() && opts.scale > 0.0) {
        return Err(format!(
            "scale must be a positive number, got {}",
            opts.scale
        ));
    }

    let sheet_ref = wb
        .sheet(sheet)
        .ok_or_else(|| "sheet index out of range".to_string())?;

    let mut vp = match opts.range {
        Some(range) => viewport_for_range(sheet_ref, range),
        None => viewport_for_used_range(sheet_ref),
    };

    // caps are output pixels, measured after scaling
    if let Some(w) = opts.max_width {
        vp.width = vp.width.min(w as f32 / opts.scale);
    }
    if let Some(h) = opts.max_height {
        vp.height = vp.height.min(h as f32 / opts.scale);
    }

    let out_w = ((vp.width * opts.scale).ceil() as u32).max(1);
    let out_h = ((vp.height * opts.scale).ceil() as u32).max(1);
    if out_w > MAX_PIXMAP_DIM || out_h > MAX_PIXMAP_DIM {
        return Err(format!(
            "requested render is {out_w}x{out_h}px, exceeds the {MAX_PIXMAP_DIM}px per-side cap; \
             narrow the range or lower --scale"
        ));
    }

    let dl = build_display_list(wb, sheet, &vp);
    let dl = if opts.scale == 1.0 {
        dl
    } else {
        scaled(dl, opts.scale)
    };
    let bytes = render_png(&dl)?;
    Ok(RenderedPng {
        bytes,
        width: out_w,
        height: out_h,
    })
}

/// one summary per sheet, in workbook order.
pub fn sheet_summaries(wb: &Workbook) -> Vec<SheetSummary> {
    wb.sheets
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
