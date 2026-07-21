//! style types (fonts, fills, borders, xf chains, theme colors). pure data;
//! the cellXfs indirection chain is walked through the `Stylesheet` accessors.

use serde::{Deserialize, Serialize};

/// a color reference as it appears in a `<color>` element (§18.8.3).
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub enum Color {
    /// resolved `#rrggbb` (alpha dropped at parse; excel stores `aarrggbb`).
    Rgb(String),
    /// theme slot index (see [`Theme::slot`]) plus a signed tint in [-1.0, 1.0].
    Theme { idx: u8, tint: f64 },
    /// index into the legacy 64-entry palette.
    Indexed(u8),
    /// system/automatic color; the host decides (usually black text).
    Auto,
}

impl Color {
    /// resolve to a final `#rrggbb`, or `None` for automatic/out-of-range.
    pub fn resolve(&self, theme: &Theme) -> Option<String> {
        match self {
            Color::Rgb(s) => Some(s.clone()),
            Color::Auto => None,
            Color::Indexed(i) => indexed_color(*i).map(str::to_string),
            Color::Theme { idx, tint } => {
                let base = theme.slot(*idx)?;
                Some(apply_tint(base, *tint))
            }
        }
    }
}

/// a cell font. only the facets we render are modelled; unset fields inherit.
#[derive(Debug, Clone, PartialEq, Default, Serialize, Deserialize)]
pub struct Font {
    pub name: Option<String>,
    pub size_pt: Option<f64>,
    pub bold: bool,
    pub italic: bool,
    pub underline: bool,
    pub strike: bool,
    pub color: Option<Color>,
}

/// a cell fill. non-solid patterns collapse to their foreground color.
#[derive(Debug, Clone, PartialEq, Default, Serialize, Deserialize)]
pub enum Fill {
    #[default]
    None,
    Solid(Color),
}

/// border line weight/style, collapsed from the full sml set.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum BorderStyle {
    Thin,
    Medium,
    Thick,
    Dashed,
    Dotted,
    Double,
    Hair,
}

impl BorderStyle {
    /// map an sml `ST_BorderStyle` token; unknown weights fall back to `thin`.
    pub fn from_sml(s: &str) -> Self {
        match s {
            "thin" => BorderStyle::Thin,
            "medium" | "mediumDashed" | "mediumDashDot" | "mediumDashDotDot" => BorderStyle::Medium,
            "thick" => BorderStyle::Thick,
            "dashed" | "dashDot" | "dashDotDot" | "slantDashDot" => BorderStyle::Dashed,
            "dotted" => BorderStyle::Dotted,
            "double" => BorderStyle::Double,
            "hair" => BorderStyle::Hair,
            _ => BorderStyle::Thin,
        }
    }

    /// the sml token for this weight.
    pub fn as_sml(&self) -> &'static str {
        match self {
            BorderStyle::Thin => "thin",
            BorderStyle::Medium => "medium",
            BorderStyle::Thick => "thick",
            BorderStyle::Dashed => "dashed",
            BorderStyle::Dotted => "dotted",
            BorderStyle::Double => "double",
            BorderStyle::Hair => "hair",
        }
    }
}

/// one edge of a cell border.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct BorderEdge {
    pub style: BorderStyle,
    pub color: Option<Color>,
}

/// the four cell edges; `None` on an edge means no border there.
#[derive(Debug, Clone, PartialEq, Default, Serialize, Deserialize)]
pub struct Border {
    pub left: Option<BorderEdge>,
    pub right: Option<BorderEdge>,
    pub top: Option<BorderEdge>,
    pub bottom: Option<BorderEdge>,
}

/// horizontal alignment (`ST_HorizontalAlignment`).
#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum HAlign {
    General,
    Left,
    Center,
    Right,
    Fill,
    Justify,
    CenterContinuous,
    Distributed,
}

impl HAlign {
    pub fn from_sml(s: &str) -> Option<Self> {
        Some(match s {
            "general" => HAlign::General,
            "left" => HAlign::Left,
            "center" => HAlign::Center,
            "right" => HAlign::Right,
            "fill" => HAlign::Fill,
            "justify" => HAlign::Justify,
            "centerContinuous" => HAlign::CenterContinuous,
            "distributed" => HAlign::Distributed,
            _ => return None,
        })
    }

    pub fn as_sml(&self) -> &'static str {
        match self {
            HAlign::General => "general",
            HAlign::Left => "left",
            HAlign::Center => "center",
            HAlign::Right => "right",
            HAlign::Fill => "fill",
            HAlign::Justify => "justify",
            HAlign::CenterContinuous => "centerContinuous",
            HAlign::Distributed => "distributed",
        }
    }
}

/// vertical alignment (`ST_VerticalAlignment`).
#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum VAlign {
    Top,
    Center,
    Bottom,
    Justify,
    Distributed,
}

impl VAlign {
    pub fn from_sml(s: &str) -> Option<Self> {
        Some(match s {
            "top" => VAlign::Top,
            "center" => VAlign::Center,
            "bottom" => VAlign::Bottom,
            "justify" => VAlign::Justify,
            "distributed" => VAlign::Distributed,
            _ => return None,
        })
    }

    pub fn as_sml(&self) -> &'static str {
        match self {
            VAlign::Top => "top",
            VAlign::Center => "center",
            VAlign::Bottom => "bottom",
            VAlign::Justify => "justify",
            VAlign::Distributed => "distributed",
        }
    }
}

/// cell text alignment; unset horizontal defaults to `general`, vertical to `bottom`.
#[derive(Debug, Clone, PartialEq, Default, Serialize, Deserialize)]
pub struct Alignment {
    pub h: Option<HAlign>,
    pub v: Option<VAlign>,
    pub wrap_text: bool,
    pub shrink_to_fit: bool,
}

impl Alignment {
    /// true when the element carries no meaningful alignment.
    pub fn is_empty(&self) -> bool {
        self.h.is_none() && self.v.is_none() && !self.wrap_text && !self.shrink_to_fit
    }
}

/// a cell format record (`CT_Xf` in cellXfs). a `None` pool index means
/// "inherit / default", not "index 0".
#[derive(Debug, Clone, PartialEq, Default, Serialize, Deserialize)]
pub struct Xf {
    pub font: Option<u32>,
    pub fill: Option<u32>,
    pub border: Option<u32>,
    pub num_fmt_id: Option<u16>,
    pub alignment: Option<Alignment>,
}

/// the workbook theme colors in clrScheme declaration order
/// `[dk1, lt1, dk2, lt2, accent1..6, hlink, folHlink]`; index via [`Theme::slot`].
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct Theme {
    pub colors: [String; 12],
}

impl Default for Theme {
    fn default() -> Self {
        Theme {
            colors: [
                "#000000".into(),
                "#ffffff".into(),
                "#44546a".into(),
                "#e7e6e6".into(),
                "#4472c4".into(),
                "#ed7d31".into(),
                "#a5a5a5".into(),
                "#ffc000".into(),
                "#5b9bd5".into(),
                "#70ad47".into(),
                "#0563c1".into(),
                "#954f72".into(),
            ],
        }
    }
}

impl Theme {
    /// resolve a `theme="n"` index to a slot color. excel's index order swaps
    /// the first two light/dark pairs relative to declaration order.
    pub fn slot(&self, idx: u8) -> Option<&str> {
        let pos = match idx {
            0 => 1,
            1 => 0,
            2 => 3,
            3 => 2,
            n => n as usize,
        };
        self.colors.get(pos).map(String::as_str)
    }
}

/// the parsed style tables plus the resolved theme. a cell's `s` indexes
/// `cell_xfs`; `num_fmts` holds only custom codes (id >= 164).
#[derive(Debug, Clone, PartialEq, Default, Serialize, Deserialize)]
pub struct Stylesheet {
    pub fonts: Vec<Font>,
    pub fills: Vec<Fill>,
    pub borders: Vec<Border>,
    pub cell_xfs: Vec<Xf>,
    pub num_fmts: Vec<(u16, String)>,
    pub theme: Theme,
}

/// resolved number format for a cell: a custom code string or a builtin id.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum FormatCode<'a> {
    Custom(&'a str),
    Builtin(u16),
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(tag = "kind", rename_all = "camelCase")]
pub enum NumberFormat {
    Builtin { id: u16 },
    Custom { pattern: String },
}

impl Default for NumberFormat {
    fn default() -> Self {
        Self::Builtin { id: 0 }
    }
}

#[derive(Debug, Clone, PartialEq, Default, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct CellFormat {
    pub font: Font,
    pub fill: Fill,
    pub border: Border,
    pub number_format: NumberFormat,
    pub alignment: Alignment,
}

impl Stylesheet {
    /// true when no style data is present, so the serializer skips the part.
    pub fn is_empty(&self) -> bool {
        self.fonts.is_empty()
            && self.fills.is_empty()
            && self.borders.is_empty()
            && self.cell_xfs.is_empty()
            && self.num_fmts.is_empty()
            && self.theme == Theme::default()
    }

    /// the `Xf` a cell's `s` index selects.
    pub fn xf(&self, style_index: u32) -> Option<&Xf> {
        self.cell_xfs.get(style_index as usize)
    }

    /// resolved font for a cell style, or `None` when inherited/unset.
    pub fn font_for(&self, style_index: u32) -> Option<&Font> {
        let idx = self.xf(style_index)?.font?;
        self.fonts.get(idx as usize)
    }

    /// resolved fill for a cell style.
    pub fn fill_for(&self, style_index: u32) -> Option<&Fill> {
        let idx = self.xf(style_index)?.fill?;
        self.fills.get(idx as usize)
    }

    /// resolved border for a cell style.
    pub fn border_for(&self, style_index: u32) -> Option<&Border> {
        let idx = self.xf(style_index)?.border?;
        self.borders.get(idx as usize)
    }

    /// resolved alignment for a cell style.
    pub fn alignment_for(&self, style_index: u32) -> Option<&Alignment> {
        self.xf(style_index)?.alignment.as_ref()
    }

    /// the number-format for a cell style: custom code when id >= 164, else the
    /// builtin id. a missing/unset xf resolves to builtin `0` (General).
    pub fn format_code_for(&self, style_index: u32) -> FormatCode<'_> {
        let id = self
            .xf(style_index)
            .and_then(|xf| xf.num_fmt_id)
            .unwrap_or(0);
        if id >= 164
            && let Some((_, code)) = self.num_fmts.iter().find(|(k, _)| *k == id)
        {
            return FormatCode::Custom(code);
        }
        FormatCode::Builtin(id)
    }

    pub fn cell_format(&self, style_index: Option<u32>) -> CellFormat {
        let Some(style_index) = style_index else {
            return CellFormat::default();
        };
        let number_format = match self.format_code_for(style_index) {
            FormatCode::Builtin(id) => NumberFormat::Builtin { id },
            FormatCode::Custom(pattern) => NumberFormat::Custom {
                pattern: pattern.to_string(),
            },
        };
        CellFormat {
            font: self.font_for(style_index).cloned().unwrap_or_default(),
            fill: self.fill_for(style_index).cloned().unwrap_or_default(),
            border: self.border_for(style_index).cloned().unwrap_or_default(),
            number_format,
            alignment: self.alignment_for(style_index).cloned().unwrap_or_default(),
        }
    }

    pub fn intern_cell_format(&mut self, format: &CellFormat) -> Option<u32> {
        if format == &CellFormat::default() {
            return None;
        }
        let font = intern(&mut self.fonts, &format.font);
        let fill = intern(&mut self.fills, &format.fill);
        let border = intern(&mut self.borders, &format.border);
        let num_fmt_id = match &format.number_format {
            NumberFormat::Builtin { id } => Some(*id).filter(|id| *id != 0),
            NumberFormat::Custom { pattern } => Some(self.intern_number_format(pattern)),
        };
        let alignment = (!format.alignment.is_empty()).then(|| format.alignment.clone());
        let xf = Xf {
            font: (format.font != Font::default()).then_some(font),
            fill: (format.fill != Fill::default()).then_some(fill),
            border: (format.border != Border::default()).then_some(border),
            num_fmt_id,
            alignment,
        };
        Some(intern(&mut self.cell_xfs, &xf))
    }

    fn intern_number_format(&mut self, pattern: &str) -> u16 {
        if let Some((id, _)) = self.num_fmts.iter().find(|(_, code)| code == pattern) {
            return *id;
        }
        let id = (164..=u16::MAX)
            .find(|id| self.num_fmts.iter().all(|(used, _)| used != id))
            .unwrap_or(u16::MAX);
        self.num_fmts.push((id, pattern.to_string()));
        id
    }
}

fn intern<T: PartialEq + Clone>(values: &mut Vec<T>, value: &T) -> u32 {
    if let Some(index) = values.iter().position(|candidate| candidate == value) {
        return index as u32;
    }
    values.push(value.clone());
    (values.len() - 1) as u32
}

/// apply a spreadsheetml tint to a `#rrggbb` in hsl luminance per §18.8.3:
/// negative darkens (`L*(1+tint)`), positive lightens (`L*(1-tint)+tint`).
fn apply_tint(hex: &str, tint: f64) -> String {
    let (r, g, b) = match parse_hex(hex) {
        Some(rgb) => rgb,
        None => return hex.to_string(),
    };
    if tint == 0.0 {
        return format!("#{r:02x}{g:02x}{b:02x}");
    }
    let (h, s, l) = rgb_to_hsl(r, g, b);
    let l2 = if tint < 0.0 {
        l * (1.0 + tint)
    } else {
        l * (1.0 - tint) + tint
    }
    .clamp(0.0, 1.0);
    let (r2, g2, b2) = hsl_to_rgb(h, s, l2);
    format!("#{r2:02x}{g2:02x}{b2:02x}")
}

/// parse `#rrggbb` (leading `#` optional) into `(r, g, b)` bytes.
fn parse_hex(hex: &str) -> Option<(u8, u8, u8)> {
    let h = hex.strip_prefix('#').unwrap_or(hex);
    if h.len() != 6 {
        return None;
    }
    let r = u8::from_str_radix(&h[0..2], 16).ok()?;
    let g = u8::from_str_radix(&h[2..4], 16).ok()?;
    let b = u8::from_str_radix(&h[4..6], 16).ok()?;
    Some((r, g, b))
}

/// rgb bytes -> hsl with hue in degrees [0,360) and s/l in [0,1].
fn rgb_to_hsl(r: u8, g: u8, b: u8) -> (f64, f64, f64) {
    let (r, g, b) = (r as f64 / 255.0, g as f64 / 255.0, b as f64 / 255.0);
    let max = r.max(g).max(b);
    let min = r.min(g).min(b);
    let l = (max + min) / 2.0;
    let d = max - min;
    if d == 0.0 {
        return (0.0, 0.0, l);
    }
    let s = d / (1.0 - (2.0 * l - 1.0).abs());
    let h = if max == r {
        60.0 * (((g - b) / d).rem_euclid(6.0))
    } else if max == g {
        60.0 * ((b - r) / d + 2.0)
    } else {
        60.0 * ((r - g) / d + 4.0)
    };
    (h, s, l)
}

/// hsl (hue degrees, s/l in [0,1]) -> rgb bytes, rounding each channel.
fn hsl_to_rgb(h: f64, s: f64, l: f64) -> (u8, u8, u8) {
    let c = (1.0 - (2.0 * l - 1.0).abs()) * s;
    let hp = h / 60.0;
    let x = c * (1.0 - (hp.rem_euclid(2.0) - 1.0).abs());
    let (r1, g1, b1) = match hp as i32 {
        0 => (c, x, 0.0),
        1 => (x, c, 0.0),
        2 => (0.0, c, x),
        3 => (0.0, x, c),
        4 => (x, 0.0, c),
        _ => (c, 0.0, x),
    };
    let m = l - c / 2.0;
    let to_byte = |v: f64| ((v + m) * 255.0).round().clamp(0.0, 255.0) as u8;
    (to_byte(r1), to_byte(g1), to_byte(b1))
}

/// the legacy 64-entry indexed palette (biff8 default); system indices 64/65
/// resolve to `None`.
fn indexed_color(i: u8) -> Option<&'static str> {
    const PALETTE: [&str; 64] = [
        "#000000", "#ffffff", "#ff0000", "#00ff00", "#0000ff", "#ffff00", "#ff00ff", "#00ffff",
        "#000000", "#ffffff", "#ff0000", "#00ff00", "#0000ff", "#ffff00", "#ff00ff", "#00ffff",
        "#800000", "#008000", "#000080", "#808000", "#800080", "#008080", "#c0c0c0", "#808080",
        "#9999ff", "#993366", "#ffffcc", "#ccffff", "#660066", "#ff8080", "#0066cc", "#ccccff",
        "#000080", "#ff00ff", "#ffff00", "#00ffff", "#800080", "#800000", "#008080", "#0000ff",
        "#00ccff", "#ccffff", "#ccffcc", "#ffff99", "#99ccff", "#ff99cc", "#cc99ff", "#ffcc99",
        "#3366ff", "#33cccc", "#99cc00", "#ffcc00", "#ff9900", "#ff6600", "#666699", "#969696",
        "#003366", "#339966", "#003300", "#333300", "#993300", "#993366", "#333399", "#333333",
    ];
    PALETTE.get(i as usize).copied()
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn theme_slot_swaps_first_two_pairs() {
        let t = Theme::default();
        assert_eq!(t.slot(0), Some("#ffffff"));
        assert_eq!(t.slot(1), Some("#000000"));
        assert_eq!(t.slot(2), Some("#e7e6e6"));
        assert_eq!(t.slot(3), Some("#44546a"));
        assert_eq!(t.slot(4), Some("#4472c4"));
        assert_eq!(t.slot(11), Some("#954f72"));
        assert_eq!(t.slot(12), None);
    }

    #[test]
    fn tint_zero_is_identity() {
        let t = Theme::default();
        let c = Color::Theme { idx: 4, tint: 0.0 };
        assert_eq!(c.resolve(&t).as_deref(), Some("#4472c4"));
    }

    #[test]
    fn tint_negative_darkens_accent1_matches_excel() {
        let t = Theme::default();
        let c = Color::Theme {
            idx: 4,
            tint: -0.25,
        };
        assert_eq!(c.resolve(&t).as_deref(), Some("#2f5597"));
    }

    #[test]
    fn tint_positive_lightens_accent1_matches_excel() {
        let t = Theme::default();
        let c = Color::Theme { idx: 4, tint: 0.4 };
        assert_eq!(c.resolve(&t).as_deref(), Some("#8faadc"));
    }

    #[test]
    fn indexed_and_rgb_and_auto_resolve() {
        let t = Theme::default();
        assert_eq!(Color::Indexed(2).resolve(&t).as_deref(), Some("#ff0000"));
        assert_eq!(Color::Indexed(64).resolve(&t), None);
        assert_eq!(
            Color::Rgb("#123456".into()).resolve(&t).as_deref(),
            Some("#123456")
        );
        assert_eq!(Color::Auto.resolve(&t), None);
    }

    #[test]
    fn accessors_walk_the_indirection_chain() {
        let mut ss = Stylesheet {
            fonts: vec![
                Font::default(),
                Font {
                    bold: true,
                    ..Font::default()
                },
            ],
            fills: vec![Fill::None, Fill::Solid(Color::Rgb("#ffff00".into()))],
            borders: vec![Border::default()],
            cell_xfs: vec![
                Xf::default(),
                Xf {
                    font: Some(1),
                    fill: Some(1),
                    border: Some(0),
                    num_fmt_id: Some(164),
                    alignment: Some(Alignment {
                        h: Some(HAlign::Center),
                        v: Some(VAlign::Center),
                        wrap_text: true,
                        shrink_to_fit: false,
                    }),
                },
            ],
            num_fmts: vec![(164, "0.0\"%\"".into())],
            theme: Theme::default(),
        };

        assert!(ss.font_for(1).unwrap().bold);
        assert_eq!(
            ss.fill_for(1),
            Some(&Fill::Solid(Color::Rgb("#ffff00".into())))
        );
        assert_eq!(ss.border_for(1), Some(&Border::default()));
        assert_eq!(ss.format_code_for(1), FormatCode::Custom("0.0\"%\""));
        assert_eq!(ss.format_code_for(0), FormatCode::Builtin(0));
        assert!(ss.font_for(0).is_none());
        assert!(ss.xf(99).is_none());
        assert_eq!(ss.format_code_for(99), FormatCode::Builtin(0));

        ss.cell_xfs[1].num_fmt_id = Some(14);
        assert_eq!(ss.format_code_for(1), FormatCode::Builtin(14));
    }

    #[test]
    fn cell_formats_intern_and_resolve() {
        let mut styles = Stylesheet::default();
        let format = CellFormat {
            font: Font {
                name: Some("Arial".into()),
                bold: true,
                color: Some(Color::Rgb("#123456".into())),
                ..Font::default()
            },
            number_format: NumberFormat::Custom {
                pattern: "0.000".into(),
            },
            alignment: Alignment {
                h: Some(HAlign::Center),
                wrap_text: true,
                ..Alignment::default()
            },
            ..CellFormat::default()
        };
        let first = styles.intern_cell_format(&format).unwrap();
        let second = styles.intern_cell_format(&format).unwrap();
        assert_eq!(first, second);
        assert_eq!(styles.cell_format(Some(first)), format);
        assert_eq!(styles.intern_cell_format(&CellFormat::default()), None);
    }
}
