//! Complete S3 styles/numbering package projection used by the differential corpus gate.

use base64::Engine as _;
use serde::{Deserialize, Serialize};

use crate::canonical::{canonical_sha256, from_serializable, to_canonical_bytes};
use crate::numbering::{
    ListLevel, ListRendering, NumberingDefinitions, NumberingMap, compute_list_rendering,
    get_bullet_character, is_bullet_level, parse_numbering, render_list_marker,
};
use crate::settings::parse_settings;
use crate::styles::{
    Style, StyleDefinitions, StyleMap, get_default_character_style, get_default_paragraph_style,
    get_default_table_style, parse_style_definitions,
};
use crate::theme::{apply_theme_font_lang, parse_theme};
use crate::xml::{ParseBudget, ParseError, ParseLimits};

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct S3Projection {
    pub styles: Option<StyleDefinitions>,
    pub style_map_entries: Vec<(String, Style)>,
    pub style_defaults: StyleDefaults,
    pub numbering: NumberingDefinitions,
    pub numbering_resolutions: Vec<NumberingResolution>,
}

#[derive(Clone, Debug, Default, PartialEq, Eq, Serialize, Deserialize)]
pub struct StyleDefaults {
    #[serde(skip_serializing_if = "Option::is_none")]
    pub paragraph: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub character: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub table: Option<String>,
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct NumberingResolution {
    pub num_id: f64,
    pub levels: Vec<ResolvedNumberingLevel>,
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ResolvedNumberingLevel {
    pub ilvl: f64,
    pub level: ListLevel,
    pub rendering: Option<ListRendering>,
    pub rendered_marker: Option<String>,
    pub bullet_character: String,
    pub is_bullet: bool,
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct S3WireEnvelope {
    pub wire_version: u8,
    pub projection: S3Projection,
    pub canonical_base64: String,
    pub canonical_sha256: String,
}

pub fn parse_docx_s3_projection(data: &[u8]) -> Result<S3Projection, ParseError> {
    let parts = ooxml_opc::unzip_parts(data).map_err(ParseError::Container)?;
    let limits = ParseLimits::default();
    let mut budget = ParseBudget::new(&limits);
    let settings = parse_settings(
        find_part(&parts, "word/settings.xml"),
        "word/settings.xml",
        &mut budget,
    )?;
    let mut theme = parse_theme(
        find_part(&parts, "word/theme/theme1.xml"),
        "word/theme/theme1.xml",
        &mut budget,
    )?;
    apply_theme_font_lang(&mut theme, settings.theme_font_lang.as_ref());

    let styles = find_part(&parts, "word/styles.xml")
        .map(|xml| parse_style_definitions(xml, Some(&theme), "word/styles.xml", &mut budget))
        .transpose()?;
    let style_map = styles
        .as_ref()
        .map(style_map_from_definitions)
        .unwrap_or_default();
    let style_defaults = StyleDefaults {
        paragraph: get_default_paragraph_style(&style_map).map(|style| style.style_id.clone()),
        character: get_default_character_style(&style_map).map(|style| style.style_id.clone()),
        table: get_default_table_style(&style_map).map(|style| style.style_id.clone()),
    };
    let style_map_entries = style_map.into_iter().collect();

    let numbering = parse_numbering(
        find_part(&parts, "word/numbering.xml"),
        "word/numbering.xml",
        &mut budget,
    )?;
    let numbering_resolutions = project_numbering_resolutions(&numbering);
    Ok(S3Projection {
        styles,
        style_map_entries,
        style_defaults,
        numbering: numbering.definitions,
        numbering_resolutions,
    })
}

pub fn s3_wire_envelope(projection: S3Projection) -> Result<S3WireEnvelope, ParseError> {
    let canonical =
        from_serializable(&projection).map_err(|error| ParseError::Canonical(error.to_string()))?;
    let bytes =
        to_canonical_bytes(&canonical).map_err(|error| ParseError::Canonical(error.to_string()))?;
    let sha =
        canonical_sha256(&canonical).map_err(|error| ParseError::Canonical(error.to_string()))?;
    Ok(S3WireEnvelope {
        wire_version: 1,
        projection,
        canonical_base64: base64::engine::general_purpose::STANDARD.encode(bytes),
        canonical_sha256: sha,
    })
}

pub fn parse_docx_s3_wire(data: &[u8]) -> Result<S3WireEnvelope, ParseError> {
    s3_wire_envelope(parse_docx_s3_projection(data)?)
}

fn style_map_from_definitions(definitions: &StyleDefinitions) -> StyleMap {
    definitions
        .styles
        .iter()
        .map(|style| (style.style_id.clone(), style.clone()))
        .collect()
}

fn project_numbering_resolutions(numbering: &NumberingMap) -> Vec<NumberingResolution> {
    let mut num_ids = Vec::new();
    for instance in &numbering.definitions.nums {
        if !num_ids.iter().any(|value| *value == instance.num_id) {
            num_ids.push(instance.num_id);
        }
    }
    num_ids
        .into_iter()
        .map(|num_id| {
            let mut levels = Vec::new();
            for ilvl in 0..=8 {
                let ilvl = ilvl as f64;
                let Some(level) = numbering.get_level(num_id, ilvl) else {
                    continue;
                };
                let rendering = compute_list_rendering(Some(num_id), Some(ilvl), numbering);
                let rendered_marker = rendering.as_ref().map(|rendering| {
                    render_list_marker(
                        &rendering.marker,
                        &[1, 2, 3, 4, 5, 6, 7, 8, 9],
                        &rendering.level_num_fmts,
                    )
                });
                levels.push(ResolvedNumberingLevel {
                    ilvl,
                    bullet_character: get_bullet_character(&level),
                    is_bullet: is_bullet_level(&level),
                    level,
                    rendering,
                    rendered_marker,
                });
            }
            NumberingResolution { num_id, levels }
        })
        .collect()
}

fn find_part<'a>(parts: &'a [(String, Vec<u8>)], path: &str) -> Option<&'a [u8]> {
    parts
        .iter()
        .find(|(candidate, _)| candidate == path)
        .map(|(_, bytes)| bytes.as_slice())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn projection_helpers_preserve_order_defaults_and_rendered_levels() {
        let styles = StyleDefinitions {
            styles: vec![
                Style {
                    style_id: "Normal".into(),
                    style_type: "paragraph".into(),
                    default: Some(true),
                    ..Style::default()
                },
                Style {
                    style_id: "DefaultTable".into(),
                    style_type: "table".into(),
                    default: Some(true),
                    ..Style::default()
                },
            ],
            ..StyleDefinitions::default()
        };
        let map = style_map_from_definitions(&styles);
        assert_eq!(
            get_default_paragraph_style(&map).unwrap().style_id,
            "Normal"
        );
        assert_eq!(
            get_default_table_style(&map).unwrap().style_id,
            "DefaultTable"
        );

        let numbering = NumberingMap {
            definitions: NumberingDefinitions {
                abstract_nums: vec![crate::numbering::AbstractNumbering {
                    abstract_num_id: 1.0,
                    multi_level_type: None,
                    num_style_link: None,
                    style_link: None,
                    levels: vec![ListLevel {
                        ilvl: 0.0,
                        start: None,
                        num_fmt: "decimal".into(),
                        lvl_text: "%1.".into(),
                        lvl_jc: None,
                        suffix: None,
                        p_pr: None,
                        r_pr: None,
                        lvl_restart: None,
                        is_lgl: None,
                        legacy: None,
                    }],
                    name: None,
                }],
                nums: vec![crate::numbering::NumberingInstance {
                    num_id: 2.0,
                    abstract_num_id: 1.0,
                    level_overrides: None,
                }],
            },
        };
        let resolutions = project_numbering_resolutions(&numbering);
        assert_eq!(
            resolutions[0].levels[0].rendered_marker.as_deref(),
            Some("1.")
        );
    }
}
