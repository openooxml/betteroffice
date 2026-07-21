//! `xlsx_model::Workbook` -> minimal valid xlsx parts. structural round-trip:
//! whatever `read` captures comes back out.

use std::collections::HashMap;
use std::io;

use quick_xml::Writer;
use quick_xml::events::{BytesDecl, BytesEnd, BytesStart, BytesText, Event};
use xlsx_model::addr::RowId;
use xlsx_model::styles::{Alignment, Border, BorderEdge, Color, Fill, Font, Stylesheet, Xf};
use xlsx_model::{Cell, CellRef, CellValue, DateSystem, Sheet, Workbook};

use crate::ParseError;
use crate::xml::xml_err;

const NS_MAIN: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const NS_R: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const NS_CT: &str = "http://schemas.openxmlformats.org/package/2006/content-types";
const NS_PKG_REL: &str = "http://schemas.openxmlformats.org/package/2006/relationships";
const CT_WORKSHEET: &str =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml";
const CT_SST: &str =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml";
const REL_WORKSHEET: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet";
const REL_SST: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings";
const CT_STYLES: &str = "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml";
const CT_THEME: &str = "application/vnd.openxmlformats-officedocument.theme+xml";
const REL_STYLES: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles";
const REL_THEME: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme";
const NS_DML: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";

/// serialize a workbook to opc parts in a fixed, deterministic order.
pub fn serialize_workbook(wb: &Workbook) -> Result<Vec<(String, Vec<u8>)>, ParseError> {
    let have_sst = !wb.shared_strings.is_empty();
    let have_styles = !wb.styles.is_empty();
    let mut parts = vec![
        (
            "[Content_Types].xml".to_string(),
            content_types(wb, have_sst, have_styles)?,
        ),
        ("_rels/.rels".to_string(), root_rels()?),
        ("xl/workbook.xml".to_string(), workbook_xml(wb)?),
        (
            "xl/_rels/workbook.xml.rels".to_string(),
            workbook_rels(wb, have_sst, have_styles)?,
        ),
    ];
    if have_sst {
        parts.push(("xl/sharedStrings.xml".to_string(), shared_strings_xml(wb)?));
    }
    if have_styles {
        parts.push(("xl/styles.xml".to_string(), styles_xml(&wb.styles)?));
        parts.push(("xl/theme/theme1.xml".to_string(), theme_xml(&wb.styles)?));
    }
    for (i, sheet) in wb.sheets.iter().enumerate() {
        parts.push((
            format!("xl/worksheets/sheet{}.xml", i + 1),
            worksheet_xml(sheet, wb)?,
        ));
    }
    Ok(parts)
}

/// run a builder against a fresh writer that already emitted the xml decl.
fn doc<F>(f: F) -> Result<Vec<u8>, ParseError>
where
    F: FnOnce(&mut Writer<Vec<u8>>) -> io::Result<()>,
{
    let mut w = Writer::new(Vec::new());
    w.write_event(Event::Decl(BytesDecl::new(
        "1.0",
        Some("UTF-8"),
        Some("yes"),
    )))
    .map_err(xml_err)?;
    f(&mut w).map_err(xml_err)?;
    Ok(w.into_inner())
}

fn content_types(wb: &Workbook, have_sst: bool, have_styles: bool) -> Result<Vec<u8>, ParseError> {
    doc(|w| {
        w.create_element("Types")
            .with_attribute(("xmlns", NS_CT))
            .write_inner_content(|w| {
                w.create_element("Default")
                    .with_attribute(("Extension", "rels"))
                    .with_attribute((
                        "ContentType",
                        "application/vnd.openxmlformats-package.relationships+xml",
                    ))
                    .write_empty()?;
                w.create_element("Default")
                    .with_attribute(("Extension", "xml"))
                    .with_attribute(("ContentType", "application/xml"))
                    .write_empty()?;
                w.create_element("Override")
                    .with_attribute(("PartName", "/xl/workbook.xml"))
                    .with_attribute((
                        "ContentType",
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
                    ))
                    .write_empty()?;
                if have_sst {
                    w.create_element("Override")
                        .with_attribute(("PartName", "/xl/sharedStrings.xml"))
                        .with_attribute(("ContentType", CT_SST))
                        .write_empty()?;
                }
                if have_styles {
                    w.create_element("Override")
                        .with_attribute(("PartName", "/xl/styles.xml"))
                        .with_attribute(("ContentType", CT_STYLES))
                        .write_empty()?;
                    w.create_element("Override")
                        .with_attribute(("PartName", "/xl/theme/theme1.xml"))
                        .with_attribute(("ContentType", CT_THEME))
                        .write_empty()?;
                }
                for i in 0..wb.sheets.len() {
                    let part = format!("/xl/worksheets/sheet{}.xml", i + 1);
                    w.create_element("Override")
                        .with_attribute(("PartName", part.as_str()))
                        .with_attribute(("ContentType", CT_WORKSHEET))
                        .write_empty()?;
                }
                Ok(())
            })?;
        Ok(())
    })
}

fn root_rels() -> Result<Vec<u8>, ParseError> {
    doc(|w| {
        w.create_element("Relationships")
            .with_attribute(("xmlns", NS_PKG_REL))
            .write_inner_content(|w| {
                w.create_element("Relationship")
                    .with_attribute(("Id", "rId1"))
                    .with_attribute((
                        "Type",
                        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument",
                    ))
                    .with_attribute(("Target", "xl/workbook.xml"))
                    .write_empty()?;
                Ok(())
            })?;
        Ok(())
    })
}

fn workbook_xml(wb: &Workbook) -> Result<Vec<u8>, ParseError> {
    doc(|w| {
        w.create_element("workbook")
            .with_attribute(("xmlns", NS_MAIN))
            .with_attribute(("xmlns:r", NS_R))
            .write_inner_content(|w| {
                if wb.date_system == DateSystem::V1904 {
                    w.create_element("workbookPr")
                        .with_attribute(("date1904", "1"))
                        .write_empty()?;
                }
                w.create_element("sheets").write_inner_content(|w| {
                    for (i, sheet) in wb.sheets.iter().enumerate() {
                        let rid = format!("rId{}", i + 1);
                        let sid = (i + 1).to_string();
                        w.create_element("sheet")
                            .with_attribute(("name", sheet.name.as_str()))
                            .with_attribute(("sheetId", sid.as_str()))
                            .with_attribute(("r:id", rid.as_str()))
                            .write_empty()?;
                    }
                    Ok(())
                })?;
                Ok(())
            })?;
        Ok(())
    })
}

fn workbook_rels(wb: &Workbook, have_sst: bool, have_styles: bool) -> Result<Vec<u8>, ParseError> {
    doc(|w| {
        w.create_element("Relationships")
            .with_attribute(("xmlns", NS_PKG_REL))
            .write_inner_content(|w| {
                let mut next = wb.sheets.len() + 1;
                for i in 0..wb.sheets.len() {
                    let rid = format!("rId{}", i + 1);
                    let target = format!("worksheets/sheet{}.xml", i + 1);
                    w.create_element("Relationship")
                        .with_attribute(("Id", rid.as_str()))
                        .with_attribute(("Type", REL_WORKSHEET))
                        .with_attribute(("Target", target.as_str()))
                        .write_empty()?;
                }
                if have_sst {
                    let rid = format!("rId{next}");
                    next += 1;
                    w.create_element("Relationship")
                        .with_attribute(("Id", rid.as_str()))
                        .with_attribute(("Type", REL_SST))
                        .with_attribute(("Target", "sharedStrings.xml"))
                        .write_empty()?;
                }
                if have_styles {
                    let rid = format!("rId{next}");
                    next += 1;
                    w.create_element("Relationship")
                        .with_attribute(("Id", rid.as_str()))
                        .with_attribute(("Type", REL_STYLES))
                        .with_attribute(("Target", "styles.xml"))
                        .write_empty()?;
                    let rid = format!("rId{next}");
                    w.create_element("Relationship")
                        .with_attribute(("Id", rid.as_str()))
                        .with_attribute(("Type", REL_THEME))
                        .with_attribute(("Target", "theme/theme1.xml"))
                        .write_empty()?;
                }
                Ok(())
            })?;
        Ok(())
    })
}

fn shared_strings_xml(wb: &Workbook) -> Result<Vec<u8>, ParseError> {
    let count = wb.shared_strings.len().to_string();
    doc(|w| {
        w.create_element("sst")
            .with_attribute(("xmlns", NS_MAIN))
            .with_attribute(("count", count.as_str()))
            .with_attribute(("uniqueCount", count.as_str()))
            .write_inner_content(|w| {
                for s in &wb.shared_strings {
                    w.create_element("si").write_inner_content(|w| {
                        write_text_el(w, s)?;
                        Ok(())
                    })?;
                }
                Ok(())
            })?;
        Ok(())
    })
}

fn worksheet_xml(sheet: &Sheet, wb: &Workbook) -> Result<Vec<u8>, ParseError> {
    let sst_index: HashMap<&str, usize> = wb
        .shared_strings
        .iter()
        .enumerate()
        .map(|(i, s)| (s.as_str(), i))
        .collect();

    let mut rows: Vec<RowId> = sheet.iter_cells().map(|(r, _)| r.row).collect();
    rows.extend(sheet.row_heights.keys().copied());
    rows.sort_unstable();
    rows.dedup();

    doc(|w| {
        w.create_element("worksheet")
            .with_attribute(("xmlns", NS_MAIN))
            .write_inner_content(|w| {
                write_cols(w, sheet)?;
                w.create_element("sheetData").write_inner_content(|w| {
                    for &row in &rows {
                        write_row(w, sheet, row, &sst_index)?;
                    }
                    Ok(())
                })?;
                write_merges(w, sheet)?;
                Ok(())
            })?;
        Ok(())
    })
}

fn write_cols(w: &mut Writer<Vec<u8>>, sheet: &Sheet) -> io::Result<()> {
    if sheet.col_widths.is_empty() {
        return Ok(());
    }
    w.create_element("cols").write_inner_content(|w| {
        for (&col, &width) in &sheet.col_widths {
            let n = (col as u64 + 1).to_string();
            w.create_element("col")
                .with_attribute(("min", n.as_str()))
                .with_attribute(("max", n.as_str()))
                .with_attribute(("width", fmt_num(width).as_str()))
                .with_attribute(("customWidth", "1"))
                .write_empty()?;
        }
        Ok(())
    })?;
    Ok(())
}

fn write_row(
    w: &mut Writer<Vec<u8>>,
    sheet: &Sheet,
    row: RowId,
    sst_index: &HashMap<&str, usize>,
) -> io::Result<()> {
    let r = (row as u64 + 1).to_string();
    let mut start = BytesStart::new("row");
    start.push_attribute(("r", r.as_str()));
    let ht = sheet.row_heights.get(&row).map(|h| fmt_num(*h));
    if let Some(h) = &ht {
        start.push_attribute(("ht", h.as_str()));
        start.push_attribute(("customHeight", "1"));
    }
    w.write_event(Event::Start(start))?;
    for (addr, cell) in sheet.iter_cells().filter(|(a, _)| a.row == row) {
        write_cell(w, addr, cell, sst_index)?;
    }
    w.write_event(Event::End(BytesEnd::new("row")))?;
    Ok(())
}

/// serialize a single cell, choosing the `t` type and body from its value and
/// whether it carries a formula.
fn write_cell(
    w: &mut Writer<Vec<u8>>,
    addr: CellRef,
    cell: &Cell,
    sst_index: &HashMap<&str, usize>,
) -> io::Result<()> {
    let a1 = addr.to_a1();
    let has_formula = cell.formula.is_some();

    let mut ty: Option<&str> = None;
    let mut value: Option<String> = None;
    let mut inline: Option<String> = None;
    match &cell.value {
        CellValue::Empty => {}
        CellValue::Number { value: n } => value = Some(fmt_num(*n)),
        CellValue::Bool { value: b } => {
            ty = Some("b");
            value = Some(if *b { "1" } else { "0" }.to_string());
        }
        CellValue::Error { value: e } => {
            ty = Some("e");
            value = Some(e.as_str().to_string());
        }
        CellValue::Text { value: s } => {
            if has_formula {
                ty = Some("str");
                value = Some(s.clone());
            } else if let Some(idx) = sst_index.get(s.as_str()) {
                ty = Some("s");
                value = Some(idx.to_string());
            } else {
                ty = Some("inlineStr");
                inline = Some(s.clone());
            }
        }
    }

    let mut start = BytesStart::new("c");
    start.push_attribute(("r", a1.as_str()));
    let style = cell.style.map(|s| s.to_string());
    if let Some(s) = &style {
        start.push_attribute(("s", s.as_str()));
    }
    if let Some(t) = ty {
        start.push_attribute(("t", t));
    }
    w.write_event(Event::Start(start))?;
    if let Some(f) = &cell.formula {
        w.create_element("f")
            .write_text_content(BytesText::new(f))?;
    }
    if let Some(v) = &value {
        w.create_element("v")
            .write_text_content(BytesText::new(v))?;
    } else if let Some(s) = &inline {
        w.create_element("is").write_inner_content(|w| {
            write_text_el(w, s)?;
            Ok(())
        })?;
    }
    w.write_event(Event::End(BytesEnd::new("c")))?;
    Ok(())
}

fn write_merges(w: &mut Writer<Vec<u8>>, sheet: &Sheet) -> io::Result<()> {
    if sheet.merges.is_empty() {
        return Ok(());
    }
    let count = sheet.merges.len().to_string();
    w.create_element("mergeCells")
        .with_attribute(("count", count.as_str()))
        .write_inner_content(|w| {
            for m in &sheet.merges {
                w.create_element("mergeCell")
                    .with_attribute(("ref", m.to_a1().as_str()))
                    .write_empty()?;
            }
            Ok(())
        })?;
    Ok(())
}

/// write a `<t xml:space="preserve">` element so leading/trailing whitespace
/// survives the round-trip.
fn write_text_el(w: &mut Writer<Vec<u8>>, text: &str) -> io::Result<()> {
    w.create_element("t")
        .with_attribute(("xml:space", "preserve"))
        .write_text_content(BytesText::new(text))?;
    Ok(())
}

/// serialize the style tables verbatim; callers building a stylesheet from
/// scratch must include the sml convention entries for excel to accept it.
fn styles_xml(ss: &Stylesheet) -> Result<Vec<u8>, ParseError> {
    doc(|w| {
        w.create_element("styleSheet")
            .with_attribute(("xmlns", NS_MAIN))
            .write_inner_content(|w| {
                write_num_fmts(w, ss)?;
                write_fonts(w, ss)?;
                write_fills(w, ss)?;
                write_borders(w, ss)?;
                write_cell_xfs(w, ss)?;
                Ok(())
            })?;
        Ok(())
    })
}

fn write_num_fmts(w: &mut Writer<Vec<u8>>, ss: &Stylesheet) -> io::Result<()> {
    if ss.num_fmts.is_empty() {
        return Ok(());
    }
    let count = ss.num_fmts.len().to_string();
    w.create_element("numFmts")
        .with_attribute(("count", count.as_str()))
        .write_inner_content(|w| {
            for (id, code) in &ss.num_fmts {
                let id = id.to_string();
                w.create_element("numFmt")
                    .with_attribute(("numFmtId", id.as_str()))
                    .with_attribute(("formatCode", code.as_str()))
                    .write_empty()?;
            }
            Ok(())
        })?;
    Ok(())
}

fn write_fonts(w: &mut Writer<Vec<u8>>, ss: &Stylesheet) -> io::Result<()> {
    if ss.fonts.is_empty() {
        return Ok(());
    }
    let count = ss.fonts.len().to_string();
    w.create_element("fonts")
        .with_attribute(("count", count.as_str()))
        .write_inner_content(|w| {
            for font in &ss.fonts {
                write_font(w, font)?;
            }
            Ok(())
        })?;
    Ok(())
}

fn write_font(w: &mut Writer<Vec<u8>>, font: &Font) -> io::Result<()> {
    w.create_element("font").write_inner_content(|w| {
        if font.bold {
            w.create_element("b").write_empty()?;
        }
        if font.italic {
            w.create_element("i").write_empty()?;
        }
        if font.underline {
            w.create_element("u").write_empty()?;
        }
        if font.strike {
            w.create_element("strike").write_empty()?;
        }
        if let Some(sz) = font.size_pt {
            w.create_element("sz")
                .with_attribute(("val", fmt_num(sz).as_str()))
                .write_empty()?;
        }
        if let Some(c) = &font.color {
            write_color(w, "color", c)?;
        }
        if let Some(name) = &font.name {
            w.create_element("name")
                .with_attribute(("val", name.as_str()))
                .write_empty()?;
        }
        Ok(())
    })?;
    Ok(())
}

fn write_fills(w: &mut Writer<Vec<u8>>, ss: &Stylesheet) -> io::Result<()> {
    if ss.fills.is_empty() {
        return Ok(());
    }
    let count = ss.fills.len().to_string();
    w.create_element("fills")
        .with_attribute(("count", count.as_str()))
        .write_inner_content(|w| {
            for fill in &ss.fills {
                write_fill(w, fill)?;
            }
            Ok(())
        })?;
    Ok(())
}

fn write_fill(w: &mut Writer<Vec<u8>>, fill: &Fill) -> io::Result<()> {
    w.create_element("fill")
        .write_inner_content(|w| match fill {
            Fill::None => {
                w.create_element("patternFill")
                    .with_attribute(("patternType", "none"))
                    .write_empty()?;
                Ok(())
            }
            Fill::Solid(color) => {
                w.create_element("patternFill")
                    .with_attribute(("patternType", "solid"))
                    .write_inner_content(|w| {
                        write_color(w, "fgColor", color)?;
                        Ok(())
                    })?;
                Ok(())
            }
        })?;
    Ok(())
}

fn write_borders(w: &mut Writer<Vec<u8>>, ss: &Stylesheet) -> io::Result<()> {
    if ss.borders.is_empty() {
        return Ok(());
    }
    let count = ss.borders.len().to_string();
    w.create_element("borders")
        .with_attribute(("count", count.as_str()))
        .write_inner_content(|w| {
            for border in &ss.borders {
                write_border(w, border)?;
            }
            Ok(())
        })?;
    Ok(())
}

fn write_border(w: &mut Writer<Vec<u8>>, border: &Border) -> io::Result<()> {
    w.create_element("border").write_inner_content(|w| {
        write_edge(w, "left", &border.left)?;
        write_edge(w, "right", &border.right)?;
        write_edge(w, "top", &border.top)?;
        write_edge(w, "bottom", &border.bottom)?;
        w.create_element("diagonal").write_empty()?;
        Ok(())
    })?;
    Ok(())
}

fn write_edge(w: &mut Writer<Vec<u8>>, name: &str, edge: &Option<BorderEdge>) -> io::Result<()> {
    match edge {
        None => {
            w.create_element(name).write_empty()?;
        }
        Some(ed) => {
            w.create_element(name)
                .with_attribute(("style", ed.style.as_sml()))
                .write_inner_content(|w| {
                    if let Some(c) = &ed.color {
                        write_color(w, "color", c)?;
                    }
                    Ok(())
                })?;
        }
    }
    Ok(())
}

fn write_cell_xfs(w: &mut Writer<Vec<u8>>, ss: &Stylesheet) -> io::Result<()> {
    if ss.cell_xfs.is_empty() {
        return Ok(());
    }
    let count = ss.cell_xfs.len().to_string();
    w.create_element("cellXfs")
        .with_attribute(("count", count.as_str()))
        .write_inner_content(|w| {
            for xf in &ss.cell_xfs {
                write_xf(w, xf)?;
            }
            Ok(())
        })?;
    Ok(())
}

/// write one cellXfs `<xf>`. unset facets serialize as index 0 with no
/// `applyX` flag, so the reader restores them to `None`.
fn write_xf(w: &mut Writer<Vec<u8>>, xf: &Xf) -> io::Result<()> {
    let num_fmt_id = xf.num_fmt_id.unwrap_or(0).to_string();
    let font_id = xf.font.unwrap_or(0).to_string();
    let fill_id = xf.fill.unwrap_or(0).to_string();
    let border_id = xf.border.unwrap_or(0).to_string();

    let mut el = BytesStart::new("xf");
    el.push_attribute(("numFmtId", num_fmt_id.as_str()));
    el.push_attribute(("fontId", font_id.as_str()));
    el.push_attribute(("fillId", fill_id.as_str()));
    el.push_attribute(("borderId", border_id.as_str()));
    if xf.num_fmt_id.is_some() {
        el.push_attribute(("applyNumberFormat", "1"));
    }
    if xf.font.is_some() {
        el.push_attribute(("applyFont", "1"));
    }
    if xf.fill.is_some() {
        el.push_attribute(("applyFill", "1"));
    }
    if xf.border.is_some() {
        el.push_attribute(("applyBorder", "1"));
    }
    if xf.alignment.is_some() {
        el.push_attribute(("applyAlignment", "1"));
    }

    match &xf.alignment {
        None => w.write_event(Event::Empty(el))?,
        Some(a) => {
            w.write_event(Event::Start(el))?;
            write_alignment(w, a)?;
            w.write_event(Event::End(BytesEnd::new("xf")))?;
        }
    }
    Ok(())
}

fn write_alignment(w: &mut Writer<Vec<u8>>, a: &Alignment) -> io::Result<()> {
    let mut el = BytesStart::new("alignment");
    if let Some(h) = a.h {
        el.push_attribute(("horizontal", h.as_sml()));
    }
    if let Some(v) = a.v {
        el.push_attribute(("vertical", v.as_sml()));
    }
    if a.wrap_text {
        el.push_attribute(("wrapText", "1"));
    }
    if a.shrink_to_fit {
        el.push_attribute(("shrinkToFit", "1"));
    }
    w.write_event(Event::Empty(el))?;
    Ok(())
}

/// write a `CT_Color` element carrying whichever representation the `Color`
/// holds. rgb is emitted as `FFrrggbb` (opaque).
fn write_color(w: &mut Writer<Vec<u8>>, name: &str, color: &Color) -> io::Result<()> {
    let mut el = BytesStart::new(name.to_string());
    let rgb;
    let idx;
    let theme;
    let tint;
    match color {
        Color::Rgb(hex) => {
            rgb = format!("FF{}", hex.trim_start_matches('#').to_ascii_uppercase());
            el.push_attribute(("rgb", rgb.as_str()));
        }
        Color::Indexed(i) => {
            idx = i.to_string();
            el.push_attribute(("indexed", idx.as_str()));
        }
        Color::Theme { idx: i, tint: t } => {
            theme = i.to_string();
            el.push_attribute(("theme", theme.as_str()));
            if *t != 0.0 {
                tint = format!("{t}");
                el.push_attribute(("tint", tint.as_str()));
            }
        }
        Color::Auto => {
            el.push_attribute(("auto", "1"));
        }
    }
    w.write_event(Event::Empty(el))?;
    Ok(())
}

/// emit a minimal but schema-shaped `theme1.xml`: the 12-color clrScheme plus
/// stub font/format schemes so excel accepts the part.
fn theme_xml(ss: &Stylesheet) -> Result<Vec<u8>, ParseError> {
    let c = &ss.theme.colors;
    let slots = [
        "dk1", "lt1", "dk2", "lt2", "accent1", "accent2", "accent3", "accent4", "accent5",
        "accent6", "hlink", "folHlink",
    ];
    doc(|w| {
        w.create_element("a:theme")
            .with_attribute(("xmlns:a", NS_DML))
            .with_attribute(("name", "Office Theme"))
            .write_inner_content(|w| {
                w.create_element("a:themeElements")
                    .write_inner_content(|w| {
                        w.create_element("a:clrScheme")
                            .with_attribute(("name", "Office"))
                            .write_inner_content(|w| {
                                for (slot, hex) in slots.iter().zip(c.iter()) {
                                    let val = hex.trim_start_matches('#').to_ascii_uppercase();
                                    w.create_element(format!("a:{slot}")).write_inner_content(
                                        |w| {
                                            w.create_element("a:srgbClr")
                                                .with_attribute(("val", val.as_str()))
                                                .write_empty()?;
                                            Ok(())
                                        },
                                    )?;
                                }
                                Ok(())
                            })?;
                        write_stub_font_scheme(w)?;
                        write_stub_fmt_scheme(w)?;
                        Ok(())
                    })?;
                Ok(())
            })?;
        Ok(())
    })
}

/// a minimal `a:fontScheme` (major/minor latin only) so the theme validates.
fn write_stub_font_scheme(w: &mut Writer<Vec<u8>>) -> io::Result<()> {
    w.create_element("a:fontScheme")
        .with_attribute(("name", "Office"))
        .write_inner_content(|w| {
            for major_minor in ["a:majorFont", "a:minorFont"] {
                w.create_element(major_minor).write_inner_content(|w| {
                    w.create_element("a:latin")
                        .with_attribute(("typeface", "Calibri"))
                        .write_empty()?;
                    w.create_element("a:ea")
                        .with_attribute(("typeface", ""))
                        .write_empty()?;
                    w.create_element("a:cs")
                        .with_attribute(("typeface", ""))
                        .write_empty()?;
                    Ok(())
                })?;
            }
            Ok(())
        })?;
    Ok(())
}

/// a minimal `a:fmtScheme`; excel requires the element even though we do not
/// model fill/line/effect styles.
fn write_stub_fmt_scheme(w: &mut Writer<Vec<u8>>) -> io::Result<()> {
    w.create_element("a:fmtScheme")
        .with_attribute(("name", "Office"))
        .write_inner_content(|w| {
            w.create_element("a:fillStyleLst")
                .write_inner_content(|w| {
                    for _ in 0..3 {
                        w.create_element("a:solidFill").write_inner_content(|w| {
                            w.create_element("a:schemeClr")
                                .with_attribute(("val", "phClr"))
                                .write_empty()?;
                            Ok(())
                        })?;
                    }
                    Ok(())
                })?;
            w.create_element("a:lnStyleLst").write_inner_content(|w| {
                for _ in 0..3 {
                    w.create_element("a:ln").write_empty()?;
                }
                Ok(())
            })?;
            w.create_element("a:effectStyleLst")
                .write_inner_content(|w| {
                    for _ in 0..3 {
                        w.create_element("a:effectStyle").write_inner_content(|w| {
                            w.create_element("a:effectLst").write_empty()?;
                            Ok(())
                        })?;
                    }
                    Ok(())
                })?;
            w.create_element("a:bgFillStyleLst")
                .write_inner_content(|w| {
                    for _ in 0..3 {
                        w.create_element("a:solidFill").write_inner_content(|w| {
                            w.create_element("a:schemeClr")
                                .with_attribute(("val", "phClr"))
                                .write_empty()?;
                            Ok(())
                        })?;
                    }
                    Ok(())
                })?;
            Ok(())
        })?;
    Ok(())
}

/// format a number the way excel writes cell values: integers without a
/// trailing `.0`.
fn fmt_num(n: f64) -> String {
    if n.is_finite() && n.fract() == 0.0 && n.abs() < 1e15 {
        return format!("{}", n as i64);
    }
    format!("{n}")
}
