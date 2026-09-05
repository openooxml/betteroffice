//! `a:fmtScheme` style lists and the `a:fillRef`/`a:lnRef` references into them.

use serde::{Deserialize, Serialize};

use crate::{ColorValue, ShapeFill, ShapeOutline, Theme, resolve_color_value_to_hex_with_theme};

/// Line width applied to a style reference the theme does not define.
const DEFAULT_STYLE_LINE_WIDTH_EMU: f64 = 9_525.0;

/// First `a:fillRef` index naming `a:bgFillStyleLst` instead of `a:fillStyleLst`.
const BACKGROUND_FILL_INDEX_BASE: u32 = 1_001;

/// Theme style lists with stable, one-based slots.
#[derive(Clone, Debug, Default, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ThemeFormatScheme {
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub fills: Vec<Option<ShapeFill>>,
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub lines: Vec<Option<ShapeOutline>>,
    /// `a:bgFillStyleLst`, which `a:fillRef` indices from 1001 name.
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub background_fills: Vec<Option<ShapeFill>>,
}

impl ThemeFormatScheme {
    pub fn is_empty(&self) -> bool {
        self.fills.is_empty() && self.lines.is_empty() && self.background_fills.is_empty()
    }

    fn fill(&self, index: u32) -> Option<&ShapeFill> {
        let (list, first) = if index >= BACKGROUND_FILL_INDEX_BASE {
            (&self.background_fills, BACKGROUND_FILL_INDEX_BASE)
        } else {
            (&self.fills, 1)
        };
        list.get(index.checked_sub(first)? as usize)?.as_ref()
    }

    fn line(&self, index: u32) -> Option<&ShapeOutline> {
        self.lines.get(index.checked_sub(1)? as usize)?.as_ref()
    }
}

/// A style index and its placeholder colour.
#[derive(Clone, Debug, Default, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct StyleReference {
    pub index: u32,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub color: Option<ColorValue>,
}

/// Theme references and unsupported explicit fills.
#[derive(Clone, Debug, Default, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ShapeStyle {
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub font_color: Option<ColorValue>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub fill: Option<StyleReference>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub line: Option<StyleReference>,
    #[serde(default, skip_serializing_if = "is_false")]
    pub fill_disabled: bool,
    #[serde(default, skip_serializing_if = "is_false")]
    pub line_disabled: bool,
}

fn is_false(value: &bool) -> bool {
    !value
}

impl ShapeStyle {
    pub fn is_empty(&self) -> bool {
        self.font_color.is_none()
            && self.fill.is_none()
            && self.line.is_none()
            && !self.fill_disabled
            && !self.line_disabled
    }
}

/// Resolves a fill reference; 0 and 1000 mean no fill.
pub fn style_fill(
    scheme: &ThemeFormatScheme,
    reference: &StyleReference,
    theme: &Theme,
) -> ShapeFill {
    if matches!(reference.index, 0 | 1000) {
        return ShapeFill::named("none");
    }
    let placeholder = placeholder_color(reference, theme);
    let Some(fill) = scheme.fill(reference.index) else {
        return ShapeFill {
            fill_type: "solid".to_owned(),
            color: placeholder,
            gradient: None,
        };
    };
    let mut resolved = fill.clone();
    resolved.color = substitute_color(fill.color.as_ref(), placeholder.as_ref());
    if let Some(gradient) = &mut resolved.gradient {
        for stop in &mut gradient.stops {
            if let Some(color) = substitute_color(Some(&stop.color), placeholder.as_ref()) {
                stop.color = color;
            }
        }
    }
    resolved
}

/// Resolves `a:lnRef` against the theme. Index 0 means no outline.
pub fn style_outline(
    scheme: &ThemeFormatScheme,
    reference: &StyleReference,
    theme: &Theme,
) -> Option<ShapeOutline> {
    if reference.index == 0 {
        return None;
    }
    let placeholder = placeholder_color(reference, theme);
    let Some(outline) = scheme.line(reference.index) else {
        return Some(ShapeOutline {
            width: Some(DEFAULT_STYLE_LINE_WIDTH_EMU),
            color: placeholder,
            ..ShapeOutline::default()
        });
    };
    Some(ShapeOutline {
        width: outline.width.or(Some(DEFAULT_STYLE_LINE_WIDTH_EMU)),
        color: substitute_color(outline.color.as_ref(), placeholder.as_ref()),
        ..outline.clone()
    })
}

/// Applies reference modifiers before style modifiers.
fn placeholder_color(reference: &StyleReference, theme: &Theme) -> Option<ColorValue> {
    let color = reference.color.as_ref()?;
    let rgb = resolve_color_value_to_hex_with_theme(Some(color), Some(theme))?;
    Some(ColorValue {
        rgb: Some(rgb.trim_start_matches('#').to_owned()),
        alpha: color.alpha,
        ..ColorValue::default()
    })
}

/// Replaces the base colour while retaining style modifiers.
fn substitute_color(
    value: Option<&ColorValue>,
    placeholder: Option<&ColorValue>,
) -> Option<ColorValue> {
    let value = value?;
    if value.theme_color.as_deref() != Some("phClr") {
        return Some(value.clone());
    }
    let Some(placeholder) = placeholder else {
        return Some(value.clone());
    };
    Some(ColorValue {
        rgb: placeholder.rgb.clone(),
        theme_color: None,
        auto: None,
        alpha: value.alpha.or(placeholder.alpha),
        ..value.clone()
    })
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::{GradientFill, GradientStop};

    fn rgb(value: &str) -> ColorValue {
        ColorValue {
            rgb: Some(value.to_owned()),
            ..Default::default()
        }
    }

    fn solid(color: ColorValue) -> ShapeFill {
        ShapeFill {
            fill_type: "solid".to_owned(),
            color: Some(color),
            gradient: None,
        }
    }

    fn reference(index: u32) -> StyleReference {
        StyleReference {
            index,
            color: Some(rgb("808080")),
        }
    }

    #[test]
    fn fill_indices_keep_holes_and_select_background_styles() {
        let scheme = ThemeFormatScheme {
            fills: vec![Some(solid(rgb("112233"))), None, Some(solid(rgb("445566")))],
            background_fills: vec![Some(solid(rgb("778899"))), Some(solid(rgb("AABBCC")))],
            ..Default::default()
        };
        let theme = Theme::default();
        for (index, expected) in [
            (1, "112233"),
            (2, "808080"),
            (3, "445566"),
            (999, "808080"),
            (1001, "778899"),
            (1002, "AABBCC"),
            (1003, "808080"),
        ] {
            assert_eq!(
                style_fill(&scheme, &reference(index), &theme)
                    .color
                    .unwrap()
                    .rgb
                    .as_deref(),
                Some(expected),
                "index {index}"
            );
        }
        for index in [0, 1000] {
            assert_eq!(
                style_fill(&scheme, &reference(index), &theme),
                ShapeFill::named("none")
            );
        }
    }

    #[test]
    fn placeholder_modifiers_compose_in_solids_lines_and_every_gradient_stop() {
        let first = ColorValue {
            theme_color: Some("phClr".to_owned()),
            luminance_modulation: Some(0.5),
            luminance_offset: Some(0.1),
            ..Default::default()
        };
        let last = ColorValue {
            luminance_modulation: Some(0.25),
            luminance_offset: Some(0.2),
            ..first.clone()
        };
        let scheme = ThemeFormatScheme {
            fills: vec![
                Some(solid(first.clone())),
                Some(ShapeFill {
                    fill_type: "gradient".to_owned(),
                    color: None,
                    gradient: Some(GradientFill {
                        gradient_type: "linear".to_owned(),
                        angle: Some(90.0),
                        stops: vec![
                            GradientStop {
                                position: 0.0,
                                color: first.clone(),
                            },
                            GradientStop {
                                position: 50000.0,
                                color: rgb("123456"),
                            },
                            GradientStop {
                                position: 100000.0,
                                color: last,
                            },
                        ],
                    }),
                }),
            ],
            lines: vec![Some(ShapeOutline {
                color: Some(first),
                ..Default::default()
            })],
            ..Default::default()
        };
        let mut theme = Theme::default();
        theme.color_scheme.accent1 = "808080".to_owned();
        let reference = |index| StyleReference {
            index,
            color: Some(ColorValue {
                theme_color: Some("accent1".to_owned()),
                luminance_modulation: Some(0.5),
                luminance_offset: Some(0.1),
                ..Default::default()
            }),
        };
        let hex = |color: &ColorValue| {
            resolve_color_value_to_hex_with_theme(Some(color), Some(&theme)).unwrap()
        };
        assert_eq!(
            hex(style_fill(&scheme, &reference(1), &theme)
                .color
                .as_ref()
                .unwrap()),
            "#474747"
        );
        assert_eq!(
            hex(style_outline(&scheme, &reference(1), &theme)
                .unwrap()
                .color
                .as_ref()
                .unwrap()),
            "#474747"
        );
        let gradient = style_fill(&scheme, &reference(2), &theme).gradient.unwrap();
        assert_eq!(gradient.angle, Some(90.0));
        assert_eq!(
            gradient
                .stops
                .iter()
                .map(|s| (s.position, hex(&s.color)))
                .collect::<Vec<_>>(),
            vec![
                (0.0, "#474747".to_owned()),
                (50000.0, "#123456".to_owned()),
                (100000.0, "#4A4A4A".to_owned())
            ]
        );
    }

    #[test]
    fn line_indices_preserve_no_fill_slots_and_explicit_style_properties() {
        let line = ShapeOutline {
            width: Some(28575.0),
            color: Some(rgb("123456")),
            style: Some("dash".to_owned()),
            cap: Some("rnd".to_owned()),
            ..Default::default()
        };
        let scheme = ThemeFormatScheme {
            lines: vec![Some(ShapeOutline::default()), None, Some(line.clone())],
            ..Default::default()
        };
        let theme = Theme::default();
        assert_eq!(style_outline(&scheme, &reference(0), &theme), None);
        assert!(
            style_outline(&scheme, &reference(1), &theme)
                .unwrap()
                .color
                .is_none()
        );
        assert_eq!(style_outline(&scheme, &reference(3), &theme), Some(line));
        for index in [2, 9] {
            let outline = style_outline(&scheme, &reference(index), &theme).unwrap();
            assert_eq!(outline.width, Some(9525.0));
            assert_eq!(outline.color.unwrap().rgb.as_deref(), Some("808080"));
        }
    }
}
