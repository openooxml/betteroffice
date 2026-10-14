use crate::display_list::{Diagnostic, Shadow};
use crate::paint;
use vsdx_eval::PageShapeReferences;
use vsdx_parse::{ThemeEffectColor, ThemeEffects, VsdxPackage};
use vsdx_resolve::{Lookup, ResolvedShape, Resolver};

const EMU_PER_INCH: f64 = 914400.0;
const POINTS_PER_INCH: f64 = 72.0;

/// Page default shadow: simple offsets plus the oblique fallback flag.
pub struct PageShadow {
    pub oblique: bool,
    pub offset_x_in: f64,
    pub offset_y_in: f64,
    pub scale: f64,
}

impl Default for PageShadow {
    fn default() -> Self {
        Self {
            oblique: false,
            offset_x_in: 0.0,
            offset_y_in: 0.0,
            scale: 1.0,
        }
    }
}

/// Caller context a shadow needs beyond the resolved shape.
pub struct ShadowContext<'a> {
    pub page: &'a PageShadow,
    pub parent_show: Option<i64>,
}

pub struct ShadowOutcome {
    pub shadow: Option<Shadow>,
    pub diagnostics: Vec<Diagnostic>,
}

/// Page `Shdw*` defaults; missing cells behave as a zero simple shadow.
pub fn page_shadow(resolver: &Resolver<'_>, package: &VsdxPackage, page: &str) -> PageShadow {
    let mut result = PageShadow::default();
    let Some(id) = package.page_part_ids.get(page) else {
        return result;
    };
    let Some(sheet) = package.page_sheets.get(id) else {
        return result;
    };
    let Ok(resolved) = resolver.resolve_sheet(sheet) else {
        return result;
    };
    if let Some(value) = cached(&resolved, "ShdwType") {
        result.oblique = value == 2.0;
    }
    if let Some(value) = cached(&resolved, "ShdwOffsetX") {
        result.offset_x_in = value;
    }
    if let Some(value) = cached(&resolved, "ShdwOffsetY") {
        result.offset_y_in = value;
    }
    if let Some(value) = cached(&resolved, "ShdwScaleFactor") {
        result.scale = value;
    }
    result
}

/// Resolved `ShapeShdwShow`, defaulting to Visio's zero.
pub fn show_value(
    package: &VsdxPackage,
    references: Option<&PageShapeReferences>,
    shape: &ResolvedShape,
    shape_id: u32,
) -> i64 {
    paint::resolved_number(package, references, shape, shape_id, "ShapeShdwShow")
        .filter(|value| value.is_finite())
        .map(|value| value.round() as i64)
        .unwrap_or(0)
}

/// Shape shadow in scene inches, or `None` when the cells specify none.
pub fn resolve(
    package: &VsdxPackage,
    references: Option<&PageShapeReferences>,
    shape: &ResolvedShape,
    shape_id: u32,
    context: &ShadowContext<'_>,
) -> ShadowOutcome {
    let mut diagnostics = Vec::new();
    let show = show_value(package, references, shape, shape_id);
    if !eligible(show, context.parent_show) {
        return ShadowOutcome {
            shadow: None,
            diagnostics,
        };
    }
    let shape_type = int_cell(package, references, shape, shape_id, "ShapeShdwType");
    let offsets = (
        paint::resolved_number(package, references, shape, shape_id, "ShapeShdwOffsetX"),
        paint::resolved_number(package, references, shape, shape_id, "ShapeShdwOffsetY"),
    );
    let theme = theme_shadow(package, references, shape, shape_id);
    let oblique = shape_type == Some(2) || shape_type.is_none() && context.page.oblique;
    let source = select_source(
        show,
        shape_type,
        offsets,
        theme.is_some(),
        context.page,
        legacy_pattern(package, references, shape, shape_id),
    );
    let Some(source) = source else {
        return ShadowOutcome {
            shadow: None,
            diagnostics,
        };
    };
    if oblique {
        diagnostics.push(Diagnostic::for_code(
            "unsupported-shadow-oblique",
            "oblique shadows render as simple offsets",
        ));
    }
    let (dx, dy, blur) = match source {
        ShadowSource::Theme => {
            let theme = theme.as_ref().cloned().unwrap_or_default();
            (theme.dx_in, theme.dy_in, theme.blur_in)
        }
        ShadowSource::Explicit(x, y) => (x, y, blur_points(package, references, shape, shape_id)),
        ShadowSource::Page => (
            context.page.offset_x_in * context.page.scale,
            context.page.offset_y_in * context.page.scale,
            0.0,
        ),
        ShadowSource::Legacy => (
            context.page.offset_x_in * context.page.scale,
            context.page.offset_y_in * context.page.scale,
            0.0,
        ),
    };
    let scale = if matches!(source, ShadowSource::Theme) {
        1.0
    } else {
        paint::resolved_number(package, references, shape, shape_id, "ShapeShdwScaleFactor")
            .filter(|value| value.is_finite() && *value >= 0.0)
            .unwrap_or(1.0)
    };
    let (dx, dy) = (dx * scale, dy * scale);
    if !dx.is_finite() || !dy.is_finite() || !blur.is_finite() {
        return ShadowOutcome {
            shadow: None,
            diagnostics,
        };
    }
    if dx == 0.0 && dy == 0.0 && blur <= 0.0 {
        return ShadowOutcome {
            shadow: None,
            diagnostics,
        };
    }
    let (color, opacity) = shadow_color(
        package,
        references,
        shape,
        shape_id,
        &theme,
        &mut diagnostics,
    );
    let color = with_opacity(&color, opacity);
    ShadowOutcome {
        shadow: Some(Shadow {
            color,
            blur_in: blur.max(0.0) as f32,
            offset_x_in: dx as f32,
            offset_y_in: dy as f32,
        }),
        diagnostics,
    }
}

#[derive(Clone, Copy, PartialEq)]
enum ShadowSource {
    Theme,
    Explicit(f64, f64),
    Page,
    Legacy,
}

fn eligible(show: i64, parent_show: Option<i64>) -> bool {
    if !matches!(show, 0..=2) {
        return false;
    }
    match parent_show {
        None => true,
        Some(parent) => (show == 2 && parent != 0) || (show == 0 && parent == 1),
    }
}

fn select_source(
    show: i64,
    shape_type: Option<i64>,
    offsets: (Option<f64>, Option<f64>),
    themed: bool,
    page: &PageShadow,
    legacy: bool,
) -> Option<ShadowSource> {
    let explicit = match offsets {
        (Some(x), Some(y)) if x.is_finite() && y.is_finite() => Some((x, y)),
        _ => None,
    };
    if themed && shape_type != Some(1) && explicit.is_none() {
        return Some(ShadowSource::Theme);
    }
    if let Some((x, y)) = explicit
        && (x != 0.0 || y != 0.0)
    {
        return Some(ShadowSource::Explicit(x, y));
    }
    if shape_type == Some(1) || shape_type.is_none() && page.oblique {
        return Some(ShadowSource::Page);
    }
    if matches!(show, 1 | 2) && (page.offset_x_in != 0.0 || page.offset_y_in != 0.0) {
        return Some(ShadowSource::Page);
    }
    legacy.then_some(ShadowSource::Legacy)
}

fn legacy_pattern(
    package: &VsdxPackage,
    references: Option<&PageShapeReferences>,
    shape: &ResolvedShape,
    shape_id: u32,
) -> bool {
    paint::resolved_number(package, references, shape, shape_id, "ShdwPattern")
        .is_some_and(|pattern| (1.0..=40.0).contains(&pattern))
}

fn blur_points(
    package: &VsdxPackage,
    references: Option<&PageShapeReferences>,
    shape: &ResolvedShape,
    shape_id: u32,
) -> f64 {
    paint::resolved_number(package, references, shape, shape_id, "ShapeShdwBlur")
        .filter(|value| value.is_finite() && *value >= 0.0)
        .map(|points| points / POINTS_PER_INCH)
        .unwrap_or(0.0)
}

fn int_cell(
    package: &VsdxPackage,
    references: Option<&PageShapeReferences>,
    shape: &ResolvedShape,
    shape_id: u32,
    name: &str,
) -> Option<i64> {
    paint::resolved_number(package, references, shape, shape_id, name)
        .filter(|value| value.is_finite() && value.fract() == 0.0)
        .map(|value| value as i64)
}

fn cached(shape: &ResolvedShape, name: &str) -> Option<f64> {
    match shape.cell(name)? {
        Lookup::Found(cell) => cell
            .cell
            .value
            .as_deref()?
            .parse()
            .ok()
            .filter(|value: &f64| value.is_finite()),
        Lookup::Deleted | Lookup::Absent => None,
    }
}

fn shadow_color(
    package: &VsdxPackage,
    references: Option<&PageShapeReferences>,
    shape: &ResolvedShape,
    shape_id: u32,
    theme: &Option<ThemeShadow>,
    diagnostics: &mut Vec<Diagnostic>,
) -> (String, f64) {
    if let Ok(color) = paint::colour(package, references, shape, shape_id, "ShdwForegnd") {
        return (color, opacity(package, references, shape, shape_id));
    }
    if let Some(color) = cached_color(package, shape, "ShdwForegnd") {
        return (color, opacity(package, references, shape, shape_id));
    }
    if let Some(theme) = theme
        && let Some(color) = theme.color.clone()
    {
        return (color, theme.opacity);
    }
    diagnostics.push(Diagnostic::for_code(
        "unresolvable-shadow-colour",
        "shadow colour falls back to the palette default",
    ));
    (
        crate::palette_colour(package, 0.0).unwrap_or_else(|| "#000000".into()),
        1.0,
    )
}

/// `ShdwForegndTrans` is a transparency, so opacity is its complement.
fn opacity(
    package: &VsdxPackage,
    references: Option<&PageShapeReferences>,
    shape: &ResolvedShape,
    shape_id: u32,
) -> f64 {
    paint::resolved_number(package, references, shape, shape_id, "ShdwForegndTrans")
        .filter(|value| value.is_finite())
        .map(|value| 1.0 - value.clamp(0.0, 1.0))
        .unwrap_or(1.0)
}

fn cached_color(package: &VsdxPackage, shape: &ResolvedShape, name: &str) -> Option<String> {
    let value = match shape.cell(name)? {
        Lookup::Found(cell) => cell.cell.value.as_deref()?,
        Lookup::Deleted | Lookup::Absent => return None,
    };
    let hex = value.strip_prefix('#').unwrap_or(value);
    if hex.len() == 6 && hex.bytes().all(|byte| byte.is_ascii_hexdigit()) {
        return Some(format!("#{}", hex.to_ascii_uppercase()));
    }
    crate::palette_colour(package, value.parse().ok()?)
}

fn with_opacity(hex: &str, opacity: f64) -> String {
    if opacity >= 1.0 {
        return hex.to_owned();
    }
    let digits = hex.strip_prefix('#').unwrap_or(hex);
    if digits.len() != 6 {
        return hex.to_owned();
    }
    format!(
        "{hex}{:02X}",
        (opacity.clamp(0.0, 1.0) * 255.0).round() as u8
    )
}

#[derive(Clone, Debug)]
struct ThemeShadow {
    dx_in: f64,
    dy_in: f64,
    blur_in: f64,
    color: Option<String>,
    opacity: f64,
}

impl Default for ThemeShadow {
    fn default() -> Self {
        Self {
            dx_in: 0.0,
            dy_in: 0.0,
            blur_in: 0.0,
            color: None,
            opacity: 1.0,
        }
    }
}

fn theme_shadow(
    package: &VsdxPackage,
    references: Option<&PageShapeReferences>,
    shape: &ResolvedShape,
    shape_id: u32,
) -> Option<ThemeShadow> {
    let (part, effects) = select_theme(package, shape, references, shape_id)?;
    let matrix = int_cell(
        package,
        references,
        shape,
        shape_id,
        "QuickStyleEffectsMatrix",
    );
    let index = match matrix {
        Some(position) if (1..=6).contains(&position) => Some((position - 1) as usize),
        Some(variant) if (100..=199).contains(&variant) => {
            variant_effect_index(effects, shape, package, references, shape_id, variant)
        }
        _ => None,
    }?;
    let outer = effects.effect_styles.get(index)?.outer_shadow.as_ref()?;
    let radians = outer.direction_60k as f64 / 60000.0 * std::f64::consts::PI / 180.0;
    let dist_in = outer.dist_emu as f64 / EMU_PER_INCH;
    let theme = package.themes.get(&part);
    let color = match &outer.color {
        ThemeEffectColor::Srgb(hex) => Some(format!("#{hex}")),
        ThemeEffectColor::Scheme(slot) if !slot.eq_ignore_ascii_case("phClr") => theme
            .and_then(|theme| theme.color_scheme.get(&slot.to_ascii_lowercase()))
            .map(|hex| format!("#{hex}").to_ascii_uppercase()),
        ThemeEffectColor::Scheme(_) | ThemeEffectColor::Placeholder => {
            quick_shadow_color(package, references, shape, shape_id, theme, effects, matrix)
        }
    };
    Some(ThemeShadow {
        dx_in: dist_in * radians.cos(),
        dy_in: -dist_in * radians.sin(),
        blur_in: (outer.blur_emu as f64 / EMU_PER_INCH).max(0.0),
        color,
        opacity: outer
            .alpha_1000pct
            .map(|alpha| (alpha as f64 / 100000.0).clamp(0.0, 1.0))
            .unwrap_or(1.0),
    })
}

fn select_theme<'a>(
    package: &'a VsdxPackage,
    shape: &ResolvedShape,
    references: Option<&PageShapeReferences>,
    shape_id: u32,
) -> Option<(u32, &'a ThemeEffects)> {
    if package.theme_effects.is_empty() {
        return None;
    }
    let want = shape
        .theme_index()
        .or_else(|| shape.color_scheme_index())
        .or_else(|| {
            int_cell(package, references, shape, shape_id, "EffectSchemeIndex")
                .and_then(|index| u32::try_from(index).ok())
        });
    match want {
        Some(index) => package
            .theme_effects
            .iter()
            .find(|(_, effects)| effects.scheme_id == Some(index))
            .map(|(part, effects)| (*part, effects))
            .or_else(|| {
                (package.theme_effects.len() == 1)
                    .then(|| package.theme_effects.iter().next())
                    .flatten()
                    .map(|(part, effects)| (*part, effects))
            }),
        None => package
            .theme_effects
            .iter()
            .next()
            .map(|(part, effects)| (*part, effects)),
    }
}

fn variant_effect_index(
    effects: &ThemeEffects,
    shape: &ResolvedShape,
    package: &VsdxPackage,
    references: Option<&PageShapeReferences>,
    shape_id: u32,
    variant: i64,
) -> Option<usize> {
    let quick_type = int_cell(package, references, shape, shape_id, "QuickStyleType").unwrap_or(0);
    let scheme = effects
        .variation_schemes
        .iter()
        .find(|scheme| scheme.embellishment == quick_type as u32)
        .or_else(|| effects.variation_schemes.first())?;
    let last = scheme.effect_indexes.len().checked_sub(1)? as i64;
    let position = (variant - 100).clamp(0, last);
    let index = *scheme.effect_indexes.get(position as usize)? as usize;
    (index >= 1).then(|| index - 1)
}

fn quick_shadow_color(
    package: &VsdxPackage,
    references: Option<&PageShapeReferences>,
    shape: &ResolvedShape,
    shape_id: u32,
    theme: Option<&ooxml_drawingml::Theme>,
    effects: &ThemeEffects,
    matrix: Option<i64>,
) -> Option<String> {
    const SLOTS: [&str; 8] = [
        "dk1", "lt1", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6",
    ];
    let slot = int_cell(
        package,
        references,
        shape,
        shape_id,
        "QuickStyleShadowColor",
    )?;
    if (0..=7).contains(&slot) {
        return theme
            .and_then(|theme| theme.color_scheme.get(SLOTS[slot as usize]))
            .map(|hex| format!("#{hex}").to_ascii_uppercase());
    }
    if (100..=106).contains(&slot) {
        let scheme = match matrix {
            Some(variant) if (100..=199).contains(&variant) => {
                let last = effects.variation_colors.len().checked_sub(1)? as i64;
                (variant - 100).clamp(0, last) as usize
            }
            _ => 0,
        };
        return effects
            .variation_colors
            .get(scheme)?
            .get((slot - 100) as usize)?
            .as_deref()
            .map(|hex| format!("#{hex}"));
    }
    None
}
