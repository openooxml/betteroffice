//! `xl/styles.xml` + `xl/theme/theme1.xml` -> `xlsx_model::Stylesheet`.
//! streaming and count-capped; missing parts resolve to defaults.

use quick_xml::events::{BytesStart, Event};
use xlsx_model::styles::{
    Alignment, Border, BorderEdge, BorderStyle, Color, Fill, Font, HAlign, Stylesheet, Theme,
    VAlign, Xf,
};

use crate::xml::{attr, local_name, next_event, reader};
use crate::{MAX_STYLE_ENTRIES, ParseError};

/// build the stylesheet from the optional style and theme parts; either falls
/// back to the model defaults.
pub(crate) fn parse_stylesheet(
    styles: Option<&[u8]>,
    theme: Option<&[u8]>,
) -> Result<Stylesheet, ParseError> {
    let theme = match theme {
        Some(bytes) => parse_theme(bytes)?,
        None => Theme::default(),
    };
    let mut sheet = match styles {
        Some(bytes) => parse_styles(bytes)?,
        None => Stylesheet::default(),
    };
    sheet.theme = theme;
    Ok(sheet)
}

/// which top-level pool the cursor is inside; disambiguates elements that recur
/// across sections (`<xf>` in cellStyleXfs vs cellXfs, `<color>` everywhere).
#[derive(PartialEq)]
enum Section {
    None,
    Fonts,
    Fills,
    Borders,
    CellStyleXfs,
    CellXfs,
}

/// parse the style pools (numFmts, fonts, fills, borders, cellXfs). the theme is
/// filled in by the caller.
fn parse_styles(data: &[u8]) -> Result<Stylesheet, ParseError> {
    let mut reader = reader(data);
    let mut buf = Vec::new();
    let mut depth = 0;

    let mut ss = Stylesheet::default();
    let mut section = Section::None;
    let mut font: Option<Font> = None;
    let mut fill: Option<Fill> = None;
    let mut border: Option<Border> = None;
    let mut edge: Option<(u8, BorderEdge)> = None;
    let mut xf: Option<Xf> = None;

    loop {
        match next_event(&mut reader, &mut buf, &mut depth)? {
            Event::Start(e) => match local_name(&e).as_slice() {
                b"numFmts" => section = Section::None,
                b"fonts" => section = Section::Fonts,
                b"fills" => section = Section::Fills,
                b"borders" => section = Section::Borders,
                b"cellStyleXfs" => section = Section::CellStyleXfs,
                b"cellXfs" => section = Section::CellXfs,
                b"numFmt" => {
                    if let (Some(id), Some(code)) = (
                        attr(&e, b"numFmtId")?.and_then(|v| v.parse::<u16>().ok()),
                        attr(&e, b"formatCode")?,
                    ) {
                        cap(ss.num_fmts.len())?;
                        ss.num_fmts.push((id, code));
                    }
                }
                b"font" if section == Section::Fonts => font = Some(Font::default()),
                b"fill" if section == Section::Fills => fill = Some(Fill::None),
                b"border" if section == Section::Borders => border = Some(Border::default()),
                b"b" if font.is_some() => set_bool(&e, &mut font, |f| &mut f.bold)?,
                b"i" if font.is_some() => set_bool(&e, &mut font, |f| &mut f.italic)?,
                b"u" if font.is_some() => set_bool(&e, &mut font, |f| &mut f.underline)?,
                b"strike" if font.is_some() => set_bool(&e, &mut font, |f| &mut f.strike)?,
                b"sz" if font.is_some() => {
                    if let Some(v) = attr(&e, b"val")?.and_then(|v| v.parse::<f64>().ok()) {
                        font.as_mut().unwrap().size_pt = Some(v);
                    }
                }
                b"name" if font.is_some() => {
                    if let Some(v) = attr(&e, b"val")? {
                        font.as_mut().unwrap().name = Some(v);
                    }
                }
                b"color" if font.is_some() => {
                    font.as_mut().unwrap().color = parse_color(&e)?;
                }
                b"patternFill" if fill.is_some() => {
                    if attr(&e, b"patternType")?.as_deref() != Some("none") {
                        // any non-none pattern collapses to solid auto until fgColor is seen
                        *fill.as_mut().unwrap() = Fill::Solid(Color::Auto);
                    }
                }
                b"fgColor" if fill.is_some() => {
                    if let Some(c) = parse_color(&e)? {
                        *fill.as_mut().unwrap() = Fill::Solid(c);
                    }
                }
                b"left" | b"start" if border.is_some() => edge = begin_edge(&e, 0)?,
                b"right" | b"end" if border.is_some() => edge = begin_edge(&e, 1)?,
                b"top" if border.is_some() => edge = begin_edge(&e, 2)?,
                b"bottom" if border.is_some() => edge = begin_edge(&e, 3)?,
                b"color" if edge.is_some() => {
                    if let Some((_, ed)) = edge.as_mut() {
                        ed.color = parse_color(&e)?;
                    }
                }
                b"xf" if section == Section::CellXfs => xf = Some(parse_xf(&e)?),
                b"alignment" if xf.is_some() && section == Section::CellXfs => {
                    let a = parse_alignment(&e)?;
                    if !a.is_empty() {
                        xf.as_mut().unwrap().alignment = Some(a);
                    }
                }
                _ => {}
            },
            Event::End(e) => match e.name().local_name().as_ref() {
                b"font" => {
                    if let Some(f) = font.take() {
                        cap(ss.fonts.len())?;
                        ss.fonts.push(f);
                    }
                }
                b"fill" => {
                    if let Some(f) = fill.take() {
                        cap(ss.fills.len())?;
                        ss.fills.push(f);
                    }
                }
                b"border" => {
                    if let Some(b) = border.take() {
                        cap(ss.borders.len())?;
                        ss.borders.push(b);
                    }
                }
                b"left" | b"start" | b"right" | b"end" | b"top" | b"bottom" => {
                    if let (Some((kind, ed)), Some(bd)) = (edge.take(), border.as_mut()) {
                        match kind {
                            0 => bd.left = Some(ed),
                            1 => bd.right = Some(ed),
                            2 => bd.top = Some(ed),
                            _ => bd.bottom = Some(ed),
                        }
                    }
                }
                b"xf" if section == Section::CellXfs => {
                    if let Some(x) = xf.take() {
                        cap(ss.cell_xfs.len())?;
                        ss.cell_xfs.push(x);
                    }
                }
                _ => {}
            },
            Event::Eof => break,
            _ => {}
        }
    }
    Ok(ss)
}

/// reject a pool that would grow past the cap before pushing another entry.
fn cap(len: usize) -> Result<(), ParseError> {
    if len >= MAX_STYLE_ENTRIES {
        return Err(ParseError::TooManyStyles);
    }
    Ok(())
}

/// set a boolean font facet from a `CT_BooleanProperty` element (`val`
/// defaults true when the element is present).
fn set_bool(
    e: &BytesStart,
    font: &mut Option<Font>,
    field: impl Fn(&mut Font) -> &mut bool,
) -> Result<(), ParseError> {
    let on = match attr(e, b"val")? {
        Some(v) => is_truthy(&v),
        None => true,
    };
    if let Some(f) = font.as_mut() {
        *field(f) = on;
    }
    Ok(())
}

/// begin a border edge if its `style` is a real weight (`none`/absent -> no
/// edge).
fn begin_edge(e: &BytesStart, kind: u8) -> Result<Option<(u8, BorderEdge)>, ParseError> {
    match attr(e, b"style")? {
        Some(s) if s != "none" => Ok(Some((
            kind,
            BorderEdge {
                style: BorderStyle::from_sml(&s),
                color: None,
            },
        ))),
        _ => Ok(None),
    }
}

/// build an `Xf`, folding the `applyX` flags in: an index is stored only when
/// its facet is applied.
fn parse_xf(e: &BytesStart) -> Result<Xf, ParseError> {
    let num_fmt = index_u16(e, b"numFmtId")?;
    let font = index_u32(e, b"fontId")?;
    let fill = index_u32(e, b"fillId")?;
    let border = index_u32(e, b"borderId")?;
    Ok(Xf {
        font: applied(e, b"applyFont")?.then_some(font).flatten(),
        fill: applied(e, b"applyFill")?.then_some(fill).flatten(),
        border: applied(e, b"applyBorder")?.then_some(border).flatten(),
        num_fmt_id: applied(e, b"applyNumberFormat")?
            .then_some(num_fmt)
            .flatten(),
        alignment: None,
    })
}

/// whether an `applyX` flag is set (truthy).
fn applied(e: &BytesStart, name: &[u8]) -> Result<bool, ParseError> {
    Ok(attr(e, name)?.map(|v| is_truthy(&v)).unwrap_or(false))
}

fn index_u32(e: &BytesStart, name: &[u8]) -> Result<Option<u32>, ParseError> {
    Ok(attr(e, name)?.and_then(|v| v.parse::<u32>().ok()))
}

fn index_u16(e: &BytesStart, name: &[u8]) -> Result<Option<u16>, ParseError> {
    Ok(attr(e, name)?.and_then(|v| v.parse::<u16>().ok()))
}

/// parse an `<alignment>` element's horizontal/vertical/wrapText attributes.
fn parse_alignment(e: &BytesStart) -> Result<Alignment, ParseError> {
    Ok(Alignment {
        h: attr(e, b"horizontal")?.and_then(|v| HAlign::from_sml(&v)),
        v: attr(e, b"vertical")?.and_then(|v| VAlign::from_sml(&v)),
        wrap_text: attr(e, b"wrapText")?
            .map(|v| is_truthy(&v))
            .unwrap_or(false),
    })
}

/// parse a `CT_Color` element (`rgb`/`theme`+`tint`/`indexed`/`auto`) into a
/// `Color`; `None` for an empty color element.
fn parse_color(e: &BytesStart) -> Result<Option<Color>, ParseError> {
    if let Some(rgb) = attr(e, b"rgb")? {
        return Ok(normalize_rgb(&rgb).map(Color::Rgb));
    }
    if let Some(theme) = attr(e, b"theme")?.and_then(|v| v.parse::<u8>().ok()) {
        let tint = attr(e, b"tint")?
            .and_then(|v| v.parse::<f64>().ok())
            .unwrap_or(0.0);
        return Ok(Some(Color::Theme { idx: theme, tint }));
    }
    if let Some(indexed) = attr(e, b"indexed")?.and_then(|v| v.parse::<u8>().ok()) {
        return Ok(Some(Color::Indexed(indexed)));
    }
    if attr(e, b"auto")?.as_deref() == Some("1") {
        return Ok(Some(Color::Auto));
    }
    Ok(None)
}

/// normalize an `aarrggbb` or `rrggbb` hex to `#rrggbb`, dropping any alpha.
fn normalize_rgb(v: &str) -> Option<String> {
    let hex = v.trim();
    let rgb = match hex.len() {
        8 => &hex[2..],
        6 => hex,
        _ => return None,
    };
    if rgb.chars().all(|c| c.is_ascii_hexdigit()) {
        Some(format!("#{}", rgb.to_ascii_lowercase()))
    } else {
        None
    }
}

/// parse `theme1.xml`'s `a:clrScheme` into the 12 slot colors, in declaration
/// order (dk1, lt1, dk2, lt2, accent1..6, hlink, folHlink).
fn parse_theme(data: &[u8]) -> Result<Theme, ParseError> {
    let mut reader = reader(data);
    let mut buf = Vec::new();
    let mut depth = 0;

    let mut theme = Theme::default();
    let mut slot: Option<usize> = None;
    let mut in_scheme = false;

    loop {
        match next_event(&mut reader, &mut buf, &mut depth)? {
            Event::Start(e) => match local_name(&e).as_slice() {
                b"clrScheme" => in_scheme = true,
                name if in_scheme && slot.is_none() => {
                    slot = slot_index(name);
                }
                b"srgbClr" => {
                    if let (Some(i), Some(val)) = (slot, attr(&e, b"val")?)
                        && let Some(c) = normalize_rgb(&val)
                    {
                        theme.colors[i] = c;
                    }
                }
                b"sysClr" => {
                    if let Some(i) = slot {
                        theme.colors[i] = sys_color(&e)?;
                    }
                }
                _ => {}
            },
            Event::End(e) => match e.name().local_name().as_ref() {
                b"clrScheme" => in_scheme = false,
                name if slot_index(name).is_some() => slot = None,
                _ => {}
            },
            Event::Eof => break,
            _ => {}
        }
    }
    Ok(theme)
}

/// resolve a `<a:sysClr>` to a hex: prefer the cached `lastClr`, else map the
/// well-known values, else black.
fn sys_color(e: &BytesStart) -> Result<String, ParseError> {
    if let Some(last) = attr(e, b"lastClr")?.and_then(|v| normalize_rgb(&v)) {
        return Ok(last);
    }
    Ok(match attr(e, b"val")?.as_deref() {
        Some("window") => "#ffffff".into(),
        Some("windowText") => "#000000".into(),
        _ => "#000000".into(),
    })
}

/// map a clrScheme slot element name to its declaration-order index.
fn slot_index(name: &[u8]) -> Option<usize> {
    Some(match name {
        b"dk1" => 0,
        b"lt1" => 1,
        b"dk2" => 2,
        b"lt2" => 3,
        b"accent1" => 4,
        b"accent2" => 5,
        b"accent3" => 6,
        b"accent4" => 7,
        b"accent5" => 8,
        b"accent6" => 9,
        b"hlink" => 10,
        b"folHlink" => 11,
        _ => return None,
    })
}

fn is_truthy(v: &str) -> bool {
    matches!(v, "1" | "true" | "on")
}
