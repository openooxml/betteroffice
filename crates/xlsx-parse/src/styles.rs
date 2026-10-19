//! `xl/styles.xml` + `xl/theme/theme1.xml` -> `xlsx_model::Stylesheet`.
//! streaming and count-capped; missing parts resolve to defaults.

use quick_xml::events::{BytesStart, Event};
use xlsx_model::styles::{
    Alignment, Border, BorderEdge, BorderStyle, Color, Fill, Font, HAlign, PatternType, Protection,
    SavedPlain, Stylesheet, Theme, VAlign, Xf,
};

use crate::xml::{attr, local_name, next_event, reader, xml_err};
use crate::{MAX_STYLE_ENTRIES, ParseError};

#[cfg(feature = "test-counters")]
thread_local! {
    static LEGACY_STYLE_READS: std::cell::Cell<u64> = const { std::cell::Cell::new(0) };
}

#[cfg(feature = "test-counters")]
#[doc(hidden)]
pub fn legacy_style_reads() -> u64 {
    LEGACY_STYLE_READS.get()
}

/// build the stylesheet from the optional style and theme parts; either falls
/// back to the model defaults.
pub(crate) fn parse_stylesheet(
    styles: Option<&[u8]>,
    theme: Option<&[u8]>,
) -> Result<(Stylesheet, Option<Stylesheet>), ParseError> {
    let theme = match theme {
        Some(bytes) => parse_theme(bytes)?,
        None => Theme::default(),
    };
    let (mut sheet, legacy_xfs) = match styles {
        Some(bytes) => parse_styles(bytes)?,
        None => (Stylesheet::default(), Vec::new()),
    };
    sheet.theme = theme;
    let legacy = (legacy_xfs != sheet.cell_xfs).then(|| {
        let mut legacy = sheet.clone();
        legacy.cell_xfs = legacy_xfs;
        legacy
    });
    Ok((sheet, legacy))
}

pub(crate) fn legacy_style_readings(
    styles: Option<&[u8]>,
    sheet: &Stylesheet,
    legacy: Option<&Stylesheet>,
    references_zero_style: bool,
) -> Result<Vec<Stylesheet>, ParseError> {
    let mut readings = Vec::new();
    if let Some(bytes) = styles {
        let (mut unfiltered, xfs) = parse_style_pools(bytes, true)?;
        unfiltered.theme = sheet.theme.clone();
        let mut held = unfiltered.clone();
        held.cell_xfs = xfs;
        for mut reading in [unfiltered, held] {
            if reading.cell_xfs.is_empty() && references_zero_style {
                reading.cell_xfs.push(Xf::default());
            }
            if &reading != sheet && legacy != Some(&reading) && !readings.contains(&reading) {
                readings.push(reading);
            }
        }
    }
    Ok(readings)
}

/// A stylesheet pool selected by logical parentage.
#[derive(PartialEq)]
enum Section {
    None,
    NumFmts,
    Fonts,
    Fills,
    Borders,
    CellStyleXfs,
    CellXfs,
}

/// parse the style pools (numFmts, fonts, fills, borders, cellXfs). the theme is
/// filled in by the caller.
fn parse_styles(data: &[u8]) -> Result<(Stylesheet, Vec<Xf>), ParseError> {
    parse_style_pools(data, false)
}

struct Scope {
    depth: usize,
    read: bool,
    transparent: bool,
    alternate: bool,
    taken: bool,
}

fn parse_style_pools(data: &[u8], unfiltered: bool) -> Result<(Stylesheet, Vec<Xf>), ParseError> {
    #[cfg(feature = "test-counters")]
    if unfiltered {
        LEGACY_STYLE_READS.set(LEGACY_STYLE_READS.get() + 1);
    }
    let mut reader = reader(data);
    let mut buf = Vec::new();
    let mut physical_depth = 0;
    let mut scopes: Vec<Scope> = Vec::new();

    let mut ss = Stylesheet::default();
    let mut section = Section::None;
    let mut font: Option<Font> = None;
    let mut fill: Option<Fill> = None;
    let mut border: Option<Border> = None;
    let mut edge: Option<(u8, BorderEdge)> = None;
    let mut xf: Option<Xf> = None;
    let mut explicit = [false; 4];
    let mut protection_applied = true;
    let mut legacy_cell_xfs: Vec<Xf> = Vec::new();
    let mut colors_depth = None;
    let mut indexed_colors_depth = None;
    let mut markup: Option<Markup> = None;
    let mut default_fonts = Vec::new();
    let mut default_fills = Vec::new();
    let mut default_borders = Vec::new();
    let mut plain_xfs = Vec::new();

    loop {
        let event = next_event(&mut reader, &mut buf, &mut physical_depth)?;
        let depth = if unfiltered {
            physical_depth
        } else {
            match &event {
                Event::Start(e) => {
                    let name = local_name(e);
                    let parent = scopes.last_mut();
                    let mut read = parent.as_ref().is_none_or(|parent| parent.read);
                    let parent_depth = parent.as_ref().map_or(0, |parent| parent.depth);
                    let alternate = name == b"AlternateContent";
                    let branch = matches!(name.as_slice(), b"Choice" | b"Fallback")
                        && parent.as_ref().is_some_and(|parent| parent.alternate);
                    if branch && let Some(parent) = parent {
                        read &= !parent.taken;
                        parent.taken = true;
                    }
                    let transparent = alternate || branch;
                    let depth = parent_depth + usize::from(!transparent);
                    scopes.push(Scope {
                        depth,
                        read,
                        transparent,
                        alternate,
                        taken: false,
                    });
                    if !read || transparent {
                        continue;
                    }
                    depth
                }
                Event::End(_) => {
                    let scope = scopes.pop();
                    if scope.is_some_and(|scope| !scope.read || scope.transparent) {
                        continue;
                    }
                    scopes.last().map_or(0, |scope| scope.depth)
                }
                _ => scopes.last().map_or(0, |scope| scope.depth),
            }
        };
        if let (Event::Start(e), Some(open)) = (&event, markup.as_mut()) {
            open.child(e, depth)?;
        }
        match event {
            Event::Start(e) => match local_name(&e).as_slice() {
                b"colors" if depth == 2 => colors_depth = Some(depth),
                b"indexedColors" if colors_depth == Some(depth - 1) => {
                    indexed_colors_depth = Some(depth);
                }
                b"rgbColor" if indexed_colors_depth == Some(depth - 1) => {
                    cap(ss.indexed_colors.len())?;
                    let rgb = match attr(&e, b"rgb")? {
                        Some(value) => normalize_rgb(&value).ok_or_else(|| {
                            ParseError::Xml("invalid indexed palette color".into())
                        })?,
                        None => String::new(),
                    };
                    ss.indexed_colors.push(rgb);
                }
                b"numFmts" if unfiltered || depth == 2 => {
                    section = if unfiltered {
                        Section::None
                    } else {
                        Section::NumFmts
                    }
                }
                b"fonts" if unfiltered || depth == 2 => section = Section::Fonts,
                b"fills" if unfiltered || depth == 2 => section = Section::Fills,
                b"borders" if unfiltered || depth == 2 => section = Section::Borders,
                b"cellStyleXfs" if unfiltered || depth == 2 => section = Section::CellStyleXfs,
                b"cellXfs" if unfiltered || depth == 2 => section = Section::CellXfs,
                b"numFmt" if unfiltered || (section == Section::NumFmts && depth == 3) => {
                    if let (Some(id), Some(code)) = (
                        attr(&e, b"numFmtId")?.and_then(|v| v.parse::<u16>().ok()),
                        attr(&e, b"formatCode")?,
                    ) {
                        cap(ss.num_fmts.len())?;
                        ss.num_fmts.push((id, code));
                    }
                }
                b"font" if section == Section::Fonts && (unfiltered || depth == 3) => {
                    font = Some(Font::default());
                    markup = Some(Markup::open(Entry::Font, &e, depth)?);
                }
                b"fill" if section == Section::Fills && (unfiltered || depth == 3) => {
                    fill = Some(Fill::None);
                    markup = Some(Markup::open(Entry::Fill, &e, depth)?);
                }
                b"border" if section == Section::Borders && (unfiltered || depth == 3) => {
                    border = Some(Border::default());
                    markup = Some(Markup::open(Entry::Border, &e, depth)?);
                }
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
                    let token = attr(&e, b"patternType")?;
                    *fill.as_mut().unwrap() = match token.as_deref() {
                        Some("none") => Fill::None,
                        token => match token.and_then(PatternType::from_sml) {
                            Some(pattern) => Fill::Pattern {
                                pattern,
                                fg: None,
                                bg: None,
                            },
                            None => Fill::Solid(Color::Auto),
                        },
                    };
                }
                b"fgColor" if fill.is_some() => {
                    let color = parse_color(&e)?;
                    match fill.as_mut().unwrap() {
                        Fill::Pattern { fg, .. } => *fg = color,
                        other => {
                            if let Some(color) = color {
                                *other = Fill::Solid(color);
                            }
                        }
                    }
                }
                b"bgColor" if fill.is_some() => {
                    if let Fill::Pattern { bg, .. } = fill.as_mut().unwrap() {
                        *bg = parse_color(&e)?;
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
                b"xf" if section == Section::CellXfs && (unfiltered || depth == 3) => {
                    let (parsed, flags) = parse_xf(&e)?;
                    xf = Some(parsed);
                    explicit = flags;
                    markup = Some(Markup::open(Entry::Xf, &e, depth)?);
                    protection_applied = applied(&e, b"applyProtection")?.0;
                }
                b"alignment" if xf.is_some() && section == Section::CellXfs => {
                    let a = parse_alignment(&e)?;
                    if !a.is_empty() {
                        xf.as_mut().unwrap().alignment = Some(a);
                    }
                }
                b"protection"
                    if xf.is_some() && section == Section::CellXfs && protection_applied =>
                {
                    xf.as_mut().unwrap().protection = Some(Protection {
                        locked: attr(&e, b"locked")?.map(|v| is_truthy(&v)).unwrap_or(true),
                        hidden: attr(&e, b"hidden")?.is_some_and(|v| is_truthy(&v)),
                    });
                }
                _ => {}
            },
            Event::End(_) if !unfiltered && depth == 1 => {
                section = Section::None;
                colors_depth = None;
            }
            Event::End(e) => match e.name().local_name().as_ref() {
                b"colors" if colors_depth == Some(depth + 1) => colors_depth = None,
                b"indexedColors" if indexed_colors_depth == Some(depth + 1) => {
                    indexed_colors_depth = None;
                }
                b"font" => {
                    if let Some(f) = font.take() {
                        cap(ss.fonts.len())?;
                        ss.fonts.push(f);
                        default_fonts.push(markup.take().is_some_and(|open| open.plain));
                    }
                }
                b"fill" => {
                    if let Some(f) = fill.take() {
                        cap(ss.fills.len())?;
                        ss.fills.push(f);
                        default_fills.push(markup.take().is_some_and(|open| open.plain));
                    }
                }
                b"border" => {
                    if let Some(b) = border.take() {
                        cap(ss.borders.len())?;
                        ss.borders.push(b);
                        default_borders.push(markup.take().is_some_and(|open| open.plain));
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
                        legacy_cell_xfs.push(held_back(&x, explicit));
                        ss.cell_xfs.push(x);
                        if let Some(open) = markup.take().filter(|open| open.plain) {
                            plain_xfs.push((ss.cell_xfs.len() as u32 - 1, open.ids));
                        }
                    }
                }
                _ => {}
            },
            Event::Eof => break,
            _ => {}
        }
    }
    let xf = plain_xfs
        .into_iter()
        .find(|&(index, [font, fill, border])| {
            names_default(&default_fonts, font)
                && names_default(&default_fills, fill)
                && names_default(&default_borders, border)
                && ss.resolved_format(Some(index)) == ss.resolved_format(None)
        })
        .map(|(index, _)| index);
    ss.set_saved_plain(SavedPlain {
        xf,
        font: first_default(&default_fonts),
        fill: first_default(&default_fills),
        border: first_default(&default_borders),
    });
    Ok((ss, legacy_cell_xfs))
}

/// The pool entries whose markup a save keeps while the model leaves them alone.
#[derive(Clone, Copy)]
enum Entry {
    Font,
    Fill,
    Border,
    Xf,
}

/// An entry being read: whether its markup so far is what a save writes for a default entry
/// or the plain xf, and for an xf the font, fill and border it names.
struct Markup {
    entry: Entry,
    depth: usize,
    plain: bool,
    ids: [u32; 3],
}

impl Markup {
    fn open(entry: Entry, e: &BytesStart, depth: usize) -> Result<Self, ParseError> {
        let ids = match entry {
            Entry::Xf => plain_xf_ids(e)?,
            _ => attributes_within(e, &[])?.then_some([0; 3]),
        };
        Ok(Self {
            entry,
            depth,
            plain: ids.is_some(),
            ids: ids.unwrap_or_default(),
        })
    }

    /// A descendant keeps the entry plain only as a child a save writes for a default fill or
    /// border, or a modeled alignment.
    fn child(&mut self, e: &BytesStart, depth: usize) -> Result<(), ParseError> {
        if !self.plain || depth <= self.depth {
            return Ok(());
        }
        if depth > self.depth + 1 {
            self.plain = false;
            return Ok(());
        }
        let name = local_name(e);
        self.plain = match self.entry {
            Entry::Font => false,
            Entry::Fill => {
                name == b"patternFill"
                    && attributes_within(e, &["patternType"])?
                    && attr(e, b"patternType")?.as_deref() == Some("none")
            }
            Entry::Border => {
                DEFAULT_BORDER_EDGES.contains(&name.as_slice()) && attributes_within(e, &[])?
            }
            Entry::Xf => name == b"alignment" && attributes_within(e, MODELED_ALIGNMENT)?,
        };
        Ok(())
    }
}

const DEFAULT_BORDER_EDGES: &[&[u8]] = &[b"left", b"right", b"top", b"bottom", b"diagonal"];

const MODELED_ALIGNMENT: &[&str] = &["horizontal", "vertical", "wrapText", "shrinkToFit"];

const PLAIN_XF_ATTRIBUTES: &[&str] = &[
    "numFmtId",
    "fontId",
    "fillId",
    "borderId",
    "xfId",
    "applyNumberFormat",
    "applyFont",
    "applyFill",
    "applyBorder",
    "applyAlignment",
];

/// The font, fill and border an xf names when its attributes are ones a save writes for the
/// plain xf: general number format, no parent style and no protection, prefix or pivot flags.
fn plain_xf_ids(e: &BytesStart) -> Result<Option<[u32; 3]>, ParseError> {
    if !attributes_within(e, PLAIN_XF_ATTRIBUTES)? {
        return Ok(None);
    }
    for name in [b"numFmtId".as_slice(), b"xfId".as_slice()] {
        if attr(e, name)?.is_some_and(|value| value != "0") {
            return Ok(None);
        }
    }
    let mut ids = [0; 3];
    let names = [
        b"fontId".as_slice(),
        b"fillId".as_slice(),
        b"borderId".as_slice(),
    ];
    for (id, name) in ids.iter_mut().zip(names) {
        if let Some(value) = attr(e, name)? {
            let Ok(value) = value.parse() else {
                return Ok(None);
            };
            *id = value;
        }
    }
    Ok(Some(ids))
}

fn attributes_within(e: &BytesStart, names: &[&str]) -> Result<bool, ParseError> {
    for attribute in e.attributes() {
        let attribute = attribute.map_err(xml_err)?;
        if !names
            .iter()
            .any(|name| attribute.key.as_ref() == name.as_bytes())
        {
            return Ok(false);
        }
    }
    Ok(true)
}

/// Whether the pool entry `id` names saves as the default entry: its markup does, or the
/// source has no such entry.
fn names_default(defaults: &[bool], id: u32) -> bool {
    defaults.get(id as usize).copied().unwrap_or(true)
}

fn first_default(defaults: &[bool]) -> Option<u32> {
    defaults
        .iter()
        .position(|default| *default)
        .map(|index| index as u32)
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
/// its facet is applied. an absent `applyX` applies, per ECMA-376 §18.8.45;
/// only an explicit falsy flag holds the facet back.
fn parse_xf(e: &BytesStart) -> Result<(Xf, [bool; 4]), ParseError> {
    let num_fmt = index_u16(e, b"numFmtId")?;
    let font = index_u32(e, b"fontId")?;
    let fill = index_u32(e, b"fillId")?;
    let border = index_u32(e, b"borderId")?;
    let (font_on, font_set) = applied(e, b"applyFont")?;
    let (fill_on, fill_set) = applied(e, b"applyFill")?;
    let (border_on, border_set) = applied(e, b"applyBorder")?;
    let (num_fmt_on, num_fmt_set) = applied(e, b"applyNumberFormat")?;
    Ok((
        Xf {
            font: font_on.then_some(font).flatten(),
            fill: fill_on.then_some(fill).flatten(),
            border: border_on.then_some(border).flatten(),
            num_fmt_id: num_fmt_on.then_some(num_fmt).flatten(),
            alignment: None,
            protection: None,
            xf_id: index_u32(e, b"xfId")?,
            quote_prefix: attr(e, b"quotePrefix")?.is_some_and(|v| is_truthy(&v)),
            pivot_button: attr(e, b"pivotButton")?.is_some_and(|v| is_truthy(&v)),
        },
        [font_set, fill_set, border_set, num_fmt_set],
    ))
}

/// whether an `applyX` flag applies, and whether it was written at all: absent
/// applies, present follows its value.
fn applied(e: &BytesStart, name: &[u8]) -> Result<(bool, bool), ParseError> {
    Ok(match attr(e, name)? {
        Some(value) => (is_truthy(&value), true),
        None => (true, false),
    })
}

/// the `Xf` releases that needed an explicit `applyX` read; a legacy
/// collaboration fingerprint is the only thing that asks for it.
fn held_back(xf: &Xf, explicit: [bool; 4]) -> Xf {
    Xf {
        font: explicit[0].then_some(xf.font).flatten(),
        fill: explicit[1].then_some(xf.fill).flatten(),
        border: explicit[2].then_some(xf.border).flatten(),
        num_fmt_id: explicit[3].then_some(xf.num_fmt_id).flatten(),
        ..xf.clone()
    }
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
        shrink_to_fit: attr(e, b"shrinkToFit")?
            .map(|v| is_truthy(&v))
            .unwrap_or(false),
        text_rotation: index_u16(e, b"textRotation")?,
        indent: index_u32(e, b"indent")?,
        relative_indent: attr(e, b"relativeIndent")?.and_then(|v| v.parse::<i32>().ok()),
        justify_last_line: attr(e, b"justifyLastLine")?.map(|v| is_truthy(&v)),
        reading_order: index_u32(e, b"readingOrder")?,
    })
}

/// parse a `CT_Color` element (`rgb`/`theme`+`tint`/`indexed`/`auto`) into a
/// `Color`; `None` for an empty color element.
pub(crate) fn parse_color(e: &BytesStart) -> Result<Option<Color>, ParseError> {
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
    if !hex.bytes().all(|byte| byte.is_ascii_hexdigit()) {
        return None;
    }
    let rgb = match hex.len() {
        8 => hex.get(2..)?,
        6 => hex,
        _ => return None,
    };
    Some(format!("#{}", rgb.to_ascii_lowercase()))
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

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn format_facets_child_apply_flags_and_defaults() {
        for (flags, protected) in [
            ("", true),
            (r#"applyAlignment="1" applyProtection="true""#, true),
            (r#"applyAlignment="0" applyProtection="false""#, false),
        ] {
            let xml = format!(
                r#"<styleSheet><cellXfs><xf xfId="1" quotePrefix="true" pivotButton="1" {flags}>
<alignment horizontal="center" textRotation="0" indent="0" relativeIndent="0" justifyLastLine="0" readingOrder="0"/>
<protection hidden="1"/></xf></cellXfs></styleSheet>"#
            );
            let (styles, _) = parse_styles(xml.as_bytes()).unwrap();
            let xf = &styles.cell_xfs[0];
            assert_eq!(xf.xf_id, Some(1));
            assert!(xf.quote_prefix);
            assert!(xf.pivot_button);
            assert_eq!(
                xf.alignment,
                Some(Alignment {
                    h: Some(HAlign::Center),
                    text_rotation: Some(0),
                    indent: Some(0),
                    relative_indent: Some(0),
                    justify_last_line: Some(false),
                    reading_order: Some(0),
                    ..Default::default()
                })
            );
            let protection = Protection {
                locked: true,
                hidden: true,
            };
            assert_eq!(xf.protection, protected.then_some(protection));
        }
    }
}
