use std::collections::BTreeMap;

use serde::{Deserialize, Serialize};
use serde_json::Value;

use crate::footnotes::NoteLayoutInput;
use crate::measure_blocks::MeasurementConfig;
use crate::types::{
    ColumnLayout, HeaderFooterRefs, Input, Layout, LayoutBlock, NoteAreaContract, PageMargins, Size,
};

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct DocumentRegions {
    #[serde(default)]
    pub sections: Vec<RegionSection>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub headers_footers: Option<Value>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub note_areas: Option<Vec<NoteAreaContract>>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub watermark: Option<Value>,
    #[serde(default)]
    pub note_settings: NoteSettings,
    #[serde(default)]
    pub even_and_odd_headers: bool,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub settings: Option<AuthoredRegionSettings>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct RegionSection {
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub section_id: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub page_size: Option<Size>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub margins: Option<PageMargins>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub columns: Option<ColumnLayout>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub header_footer_refs: Option<HeaderFooterRefs>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub page_numbering: Option<PageNumbering>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub header_distance: Option<f64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub footer_distance: Option<f64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub page_borders: Option<Value>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub watermark: Option<Value>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub vertical_align: Option<String>,
    #[serde(default)]
    pub note_settings: NoteSettings,
    #[serde(default)]
    pub title_pg: bool,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub even_and_odd_headers: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub properties: Option<AuthoredSectionProperties>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub section_start: Option<crate::types::SectionBreakType>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct AuthoredRegionSettings {
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub footnote_pr: Option<NoteProperties>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub endnote_pr: Option<NoteProperties>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub even_and_odd_headers: Option<bool>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct AuthoredSectionProperties {
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub page_width: Option<f64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub page_height: Option<f64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub margin_top: Option<f64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub margin_right: Option<f64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub margin_bottom: Option<f64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub margin_left: Option<f64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub header_distance: Option<f64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub footer_distance: Option<f64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub gutter: Option<f64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub column_count: Option<f64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub column_space: Option<f64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub equal_width: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub separator: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub columns: Option<Vec<AuthoredColumn>>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub footnote_columns: Option<u64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub section_start: Option<crate::types::SectionBreakType>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub vertical_align: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub header_references: Option<Vec<AuthoredHeaderFooterReference>>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub footer_references: Option<Vec<AuthoredHeaderFooterReference>>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub title_pg: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub even_and_odd_headers: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub page_numbering: Option<PageNumbering>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub page_borders: Option<Value>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub watermark: Option<Value>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub footnote_pr: Option<NoteProperties>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub endnote_pr: Option<NoteProperties>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct AuthoredColumn {
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub width: Option<f64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub space: Option<f64>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct AuthoredHeaderFooterReference {
    pub r#type: String,
    #[serde(rename = "rId")]
    pub r_id: String,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct NoteSettings {
    #[serde(default)]
    pub footnote: NoteProperties,
    #[serde(default)]
    pub endnote: NoteProperties,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub footnote_columns: Option<u64>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct NoteProperties {
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub position: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub num_fmt: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub num_start: Option<i64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub num_restart: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub custom_number_format: Option<String>,
}

impl NoteProperties {
    pub fn overlay(&mut self, authored: &Self) {
        if authored.position.is_some() {
            self.position.clone_from(&authored.position);
        }
        if authored.num_fmt.is_some() {
            self.num_fmt.clone_from(&authored.num_fmt);
        }
        if authored.num_start.is_some() {
            self.num_start = authored.num_start;
        }
        if authored.num_restart.is_some() {
            self.num_restart.clone_from(&authored.num_restart);
        }
        if authored.custom_number_format.is_some() {
            self.custom_number_format
                .clone_from(&authored.custom_number_format);
        }
    }
}

impl DocumentRegions {
    pub fn normalize_authored(&mut self) {
        if let Some(settings) = &self.settings {
            if let Some(footnote) = &settings.footnote_pr {
                self.note_settings.footnote.overlay(footnote);
            }
            if let Some(endnote) = &settings.endnote_pr {
                self.note_settings.endnote.overlay(endnote);
            }
            if let Some(even_and_odd_headers) = settings.even_and_odd_headers {
                self.even_and_odd_headers = even_and_odd_headers;
            }
        }
        for section in &mut self.sections {
            section.normalize_authored();
        }
    }

    pub fn note_properties(&self, section_index: usize, kind: &str) -> NoteProperties {
        let mut properties = if kind == "endnote" {
            self.note_settings.endnote.clone()
        } else {
            self.note_settings.footnote.clone()
        };
        if let Some(section) = self
            .sections
            .get(section_index)
            .or_else(|| self.sections.last())
        {
            let authored = if kind == "endnote" {
                &section.note_settings.endnote
            } else {
                &section.note_settings.footnote
            };
            properties.overlay(authored);
        }
        properties
    }

    pub fn footnote_columns(&self, section_index: usize) -> u64 {
        self.sections
            .get(section_index)
            .or_else(|| self.sections.last())
            .and_then(|section| section.note_settings.footnote_columns)
            .or(self.note_settings.footnote_columns)
            .unwrap_or(1)
            .max(1)
    }
}

impl RegionSection {
    fn normalize_authored(&mut self) {
        let Some(properties) = self.properties.as_ref() else {
            return;
        };
        let page_width = properties.page_width.filter(|value| *value != 0.0);
        let page_height = properties.page_height.filter(|value| *value != 0.0);
        self.page_size = Some(Size {
            w: page_width.map_or(816.0, twips_to_pixels),
            h: page_height.map_or(1056.0, twips_to_pixels),
        });
        let gutter = properties.gutter.map_or(0.0, twips_to_pixels);
        self.margins = Some(PageMargins {
            top: properties.margin_top.map_or(96.0, twips_to_pixels),
            right: properties.margin_right.map_or(96.0, twips_to_pixels),
            bottom: properties.margin_bottom.map_or(96.0, twips_to_pixels),
            left: properties.margin_left.map_or(96.0, twips_to_pixels) + gutter,
            header: Some(properties.header_distance.map_or(48.0, twips_to_pixels)),
            footer: Some(properties.footer_distance.map_or(48.0, twips_to_pixels)),
        });
        let count = properties.column_count.unwrap_or_else(|| {
            properties
                .columns
                .as_ref()
                .map_or(1.0, |columns| columns.len() as f64)
        });
        self.columns = (count > 1.0).then(|| ColumnLayout {
            count,
            gap: twips_to_pixels(properties.column_space.unwrap_or(720.0)),
            equal_width: Some(properties.equal_width.unwrap_or(true)),
            separator: properties.separator,
            columns: properties.columns.as_ref().map(|columns| {
                columns
                    .iter()
                    .take(count as usize)
                    .map(|column| crate::types::ColumnDefinition {
                        width: column.width.map(twips_to_pixels),
                        space: column.space.map(twips_to_pixels),
                    })
                    .collect()
            }),
        });
        self.header_distance = properties.header_distance.map(twips_to_pixels);
        self.footer_distance = properties.footer_distance.map(twips_to_pixels);
        self.page_borders.clone_from(&properties.page_borders);
        self.watermark.clone_from(&properties.watermark);
        self.vertical_align.clone_from(&properties.vertical_align);
        self.page_numbering.clone_from(&properties.page_numbering);
        self.title_pg = properties.title_pg.unwrap_or(false);
        self.even_and_odd_headers = properties.even_and_odd_headers;
        self.section_start = properties.section_start;
        self.note_settings.footnote_columns = properties.footnote_columns;
        if let Some(footnote) = &properties.footnote_pr {
            self.note_settings.footnote.overlay(footnote);
        }
        if let Some(endnote) = &properties.endnote_pr {
            self.note_settings.endnote.overlay(endnote);
        }
        self.header_footer_refs = authored_header_footer_refs(properties);
    }
}

fn authored_header_footer_refs(properties: &AuthoredSectionProperties) -> Option<HeaderFooterRefs> {
    let mut refs = HeaderFooterRefs {
        header_default: None,
        header_first: None,
        header_even: None,
        footer_default: None,
        footer_first: None,
        footer_even: None,
    };
    for reference in properties.header_references.iter().flatten() {
        match reference.r#type.as_str() {
            "default" => refs.header_default = Some(reference.r_id.clone()),
            "first" => refs.header_first = Some(reference.r_id.clone()),
            "even" => refs.header_even = Some(reference.r_id.clone()),
            _ => {}
        }
    }
    for reference in properties.footer_references.iter().flatten() {
        match reference.r#type.as_str() {
            "default" => refs.footer_default = Some(reference.r_id.clone()),
            "first" => refs.footer_first = Some(reference.r_id.clone()),
            "even" => refs.footer_even = Some(reference.r_id.clone()),
            _ => {}
        }
    }
    (refs.header_default.is_some()
        || refs.header_first.is_some()
        || refs.header_even.is_some()
        || refs.footer_default.is_some()
        || refs.footer_first.is_some()
        || refs.footer_even.is_some())
    .then_some(refs)
}

fn twips_to_pixels(twips: f64) -> f64 {
    (twips / 1440.0 * 96.0).round()
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct PageNumbering {
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub start: Option<u64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub format: Option<String>,
}

#[derive(Debug, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct RegionLayoutInput {
    #[serde(default)]
    pub measured: Vec<crate::types::MeasuredBlock>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub body_story: Option<String>,
    #[serde(default)]
    pub options: crate::types::LayoutOptions,
    #[serde(default)]
    pub regions: DocumentRegions,
    #[serde(default)]
    pub notes: NoteLayoutInput,
    #[serde(default)]
    pub measurement: MeasurementConfig,
    #[serde(default)]
    pub render_env: Value,
}

impl RegionLayoutInput {
    pub fn split(
        mut self,
    ) -> (
        Input,
        DocumentRegions,
        NoteLayoutInput,
        MeasurementConfig,
        Value,
        Option<String>,
    ) {
        self.regions.normalize_authored();
        (
            Input {
                measured: self.measured,
                options: self.options,
            },
            self.regions,
            self.notes,
            self.measurement,
            self.render_env,
            self.body_story,
        )
    }
}

pub fn apply_section_geometry(input: &mut Input, regions: &DocumentRegions) {
    apply_section_geometry_to_blocks(&mut input.measured, &mut input.options, regions);
}

pub fn apply_section_geometry_to_blocks<T>(
    blocks: &mut [T],
    options: &mut crate::types::LayoutOptions,
    regions: &DocumentRegions,
) where
    T: RegionBlock,
{
    let Some(first) = regions.sections.first() else {
        return;
    };
    if first.page_size.is_some() {
        options.page_size.clone_from(&first.page_size);
    }
    if first.margins.is_some() {
        options.margins.clone_from(&first.margins);
    }
    if first.columns.is_some() {
        options.columns.clone_from(&first.columns);
    }
    if let Some(last) = regions.sections.last() {
        if last.page_size.is_some() {
            options.final_page_size.clone_from(&last.page_size);
        }
        if last.margins.is_some() {
            options.final_margins.clone_from(&last.margins);
        }
        options.columns.clone_from(&last.columns);
        options.body_break_type = last.section_start;
    }
    let mut section_index = 0;
    for item in blocks {
        let LayoutBlock::SectionBreak(section_break) = item.block_mut() else {
            continue;
        };
        let Some(section) = regions.sections.get(section_index) else {
            break;
        };
        if section.page_size.is_some() {
            section_break.page_size.clone_from(&section.page_size);
        }
        if section.margins.is_some() {
            section_break.margins.clone_from(&section.margins);
        }
        if section.columns.is_some() {
            section_break.columns.clone_from(&section.columns);
        }
        section_index += 1;
    }
}

pub trait RegionBlock {
    fn block_mut(&mut self) -> &mut LayoutBlock;
}

impl RegionBlock for crate::types::MeasuredBlock {
    fn block_mut(&mut self) -> &mut LayoutBlock {
        &mut self.block
    }
}

impl RegionBlock for LayoutBlock {
    fn block_mut(&mut self) -> &mut LayoutBlock {
        self
    }
}

pub fn apply_document_regions(layout: &mut Layout, regions: &DocumentRegions) {
    let mut page_counts = BTreeMap::<usize, u64>::new();
    for (page_index, page) in layout.pages.iter_mut().enumerate() {
        let section_index = page.region_section_index;
        page.section_index = Some(section_index as u64);
        let section = regions
            .sections
            .get(section_index)
            .or_else(|| regions.sections.last());
        let section_page_index = page_counts.entry(section_index).or_default();
        page.section_page_index = Some(*section_page_index);
        *section_page_index += 1;

        if let Some(section) = section {
            page.section_id = section
                .section_id
                .clone()
                .or_else(|| Some(section_index.to_string()));
            page.header_footer_refs = effective_header_footer_refs(regions, section_index);
            page.header_distance = section.header_distance;
            page.footer_distance = section.footer_distance;
            page.page_borders = section.page_borders.clone();
            page.watermark = section
                .watermark
                .clone()
                .or_else(|| regions.watermark.clone());
            page.vertical_align = section.vertical_align.clone();
            if let Some(numbering) = &section.page_numbering {
                let number = numbering.start.unwrap_or(1) + page.section_page_index.unwrap_or(0);
                page.section_page_number = Some(number);
                page.page_label = Some(format_number(
                    number as i64,
                    numbering.format.as_deref().unwrap_or("decimal"),
                ));
                page.page_numbering = serde_json::to_value(numbering).ok();
            }
        } else {
            page.section_id = Some(section_index.to_string());
            page.watermark = regions.watermark.clone();
        }

        if let Some(areas) = &regions.note_areas {
            let selected: Vec<_> = areas
                .iter()
                .filter(|area| area.page_index == Some(page_index as u64))
                .cloned()
                .collect();
            if !selected.is_empty() {
                page.note_areas = Some(selected);
            }
        }
    }
}

pub fn effective_header_footer_refs(
    regions: &DocumentRegions,
    section_index: usize,
) -> Option<HeaderFooterRefs> {
    let mut effective = HeaderFooterRefs {
        header_default: None,
        header_first: None,
        header_even: None,
        footer_default: None,
        footer_first: None,
        footer_even: None,
    };
    for section in regions.sections.iter().take(section_index + 1) {
        let Some(authored) = &section.header_footer_refs else {
            continue;
        };
        if authored.header_default.is_some() {
            effective
                .header_default
                .clone_from(&authored.header_default);
        }
        if authored.header_first.is_some() {
            effective.header_first.clone_from(&authored.header_first);
        }
        if authored.header_even.is_some() {
            effective.header_even.clone_from(&authored.header_even);
        }
        if authored.footer_default.is_some() {
            effective
                .footer_default
                .clone_from(&authored.footer_default);
        }
        if authored.footer_first.is_some() {
            effective.footer_first.clone_from(&authored.footer_first);
        }
        if authored.footer_even.is_some() {
            effective.footer_even.clone_from(&authored.footer_even);
        }
    }
    (effective.header_default.is_some()
        || effective.header_first.is_some()
        || effective.header_even.is_some()
        || effective.footer_default.is_some()
        || effective.footer_first.is_some()
        || effective.footer_even.is_some())
    .then_some(effective)
}

pub(crate) fn format_number(number: i64, format: &str) -> String {
    match format {
        "decimalZero" => format!("{number:02}"),
        "decimalZero3" => format!("{number:03}"),
        "decimalZero4" => format!("{number:04}"),
        "decimalZero5" => format!("{number:05}"),
        "upperRoman" => roman(number).to_uppercase(),
        "lowerRoman" => roman(number),
        "upperLetter" => letters(number).to_uppercase(),
        "lowerLetter" => letters(number),
        "ordinal" => ordinal(number),
        "bullet" => "•".to_owned(),
        "none" => String::new(),
        "decimalEnclosedParen" => format!("({number})"),
        "numberInDash" => format!("-{number}-"),
        _ => number.to_string(),
    }
}

fn roman(number: i64) -> String {
    if !(1..=3999).contains(&number) {
        return number.to_string();
    }
    let mut remaining = number;
    let mut output = String::new();
    for (value, numeral) in [
        (1000, "m"),
        (900, "cm"),
        (500, "d"),
        (400, "cd"),
        (100, "c"),
        (90, "xc"),
        (50, "l"),
        (40, "xl"),
        (10, "x"),
        (9, "ix"),
        (5, "v"),
        (4, "iv"),
        (1, "i"),
    ] {
        while remaining >= value {
            output.push_str(numeral);
            remaining -= value;
        }
    }
    output
}

fn letters(mut number: i64) -> String {
    if number == 0 {
        return String::new();
    }
    let mut output = Vec::new();
    while number > 0 {
        output.push((b'a' + ((number - 1) % 26) as u8) as char);
        number = (number - 1) / 26;
    }
    output.into_iter().rev().collect()
}

fn ordinal(number: i64) -> String {
    let suffix = if (11..=13).contains(&(number % 100)) {
        "th"
    } else {
        match number % 10 {
            1 => "st",
            2 => "nd",
            3 => "rd",
            _ => "th",
        }
    };
    format!("{number}{suffix}")
}

#[cfg(test)]
mod tests {
    use serde_json::json;

    use super::*;
    use crate::types::{Layout, Page, PageMargins, Size};

    fn page(number: u32, section_index: usize) -> Page {
        Page {
            number,
            fragments: Vec::new(),
            margins: PageMargins {
                top: 96.0,
                right: 96.0,
                bottom: 96.0,
                left: 96.0,
                header: Some(48.0),
                footer: Some(48.0),
            },
            size: Size {
                w: 816.0,
                h: 1056.0,
            },
            orientation: None,
            section_index: Some(section_index as u64),
            region_section_index: section_index,
            header_footer_refs: None,
            footnote_ids: None,
            footnote_reserved_height: None,
            footnote_columns: None,
            columns: None,
            section_id: None,
            section_page_index: None,
            section_page_number: None,
            page_label: None,
            page_numbering: None,
            header_distance: None,
            footer_distance: None,
            page_borders: None,
            watermark: None,
            vertical_align: None,
            note_areas: None,
        }
    }

    #[test]
    fn stamps_section_furniture_and_restart_labels() {
        let mut layout = Layout {
            page_size: Size {
                w: 816.0,
                h: 1056.0,
            },
            pages: vec![page(1, 0), page(2, 0), page(3, 1)],
            columns: None,
            headers: None,
            footers: None,
            page_gap: Some(20.0),
        };
        let regions = DocumentRegions {
            sections: vec![
                RegionSection {
                    section_id: Some("a".to_owned()),
                    page_numbering: Some(PageNumbering {
                        start: Some(3),
                        format: Some("upperRoman".to_owned()),
                    }),
                    header_distance: Some(24.0),
                    ..RegionSection::default()
                },
                RegionSection {
                    section_id: Some("b".to_owned()),
                    page_numbering: Some(PageNumbering {
                        start: Some(1),
                        format: Some("lowerLetter".to_owned()),
                    }),
                    footer_distance: Some(18.0),
                    ..RegionSection::default()
                },
            ],
            ..DocumentRegions::default()
        };

        apply_document_regions(&mut layout, &regions);

        assert_eq!(layout.pages[0].section_page_index, Some(0));
        assert_eq!(layout.pages[0].section_page_number, Some(3));
        assert_eq!(layout.pages[0].page_label.as_deref(), Some("III"));
        assert_eq!(layout.pages[1].page_label.as_deref(), Some("IV"));
        assert_eq!(layout.pages[2].section_page_index, Some(0));
        assert_eq!(layout.pages[2].page_label.as_deref(), Some("a"));
        assert_eq!(layout.pages[2].footer_distance, Some(18.0));
    }

    #[test]
    fn section_geometry_populates_options_and_breaks() {
        let request: RegionLayoutInput = serde_json::from_value(json!({
            "measured": [{
                "block": {"kind": "sectionBreak", "id": "break", "type": "nextPage"},
                "measure": {"kind": "sectionBreak"}
            }],
            "options": {},
            "regions": {
                "sections": [
                    {
                        "pageSize": {"w": 300, "h": 400},
                        "margins": {"top": 10, "right": 20, "bottom": 30, "left": 40},
                        "columns": {"count": 2, "gap": 12}
                    },
                    {
                        "pageSize": {"w": 500, "h": 600},
                        "margins": {"top": 50, "right": 60, "bottom": 70, "left": 80}
                    }
                ]
            }
        }))
        .unwrap();
        let (mut input, regions, _, _, _, _) = request.split();

        apply_section_geometry(&mut input, &regions);

        assert_eq!(input.options.page_size.as_ref().unwrap().w, 300.0);
        assert_eq!(input.options.final_page_size.as_ref().unwrap().w, 500.0);
        let LayoutBlock::SectionBreak(section_break) = &input.measured[0].block else {
            panic!("section break expected");
        };
        assert_eq!(section_break.columns.as_ref().unwrap().count, 2.0);
        assert_eq!(section_break.margins.as_ref().unwrap().top, 10.0);
    }

    #[test]
    fn header_footer_relationships_inherit_per_type() {
        let regions = DocumentRegions {
            sections: vec![
                RegionSection {
                    header_footer_refs: Some(HeaderFooterRefs {
                        header_default: Some("header-a".to_owned()),
                        header_first: None,
                        header_even: None,
                        footer_default: Some("footer-a".to_owned()),
                        footer_first: None,
                        footer_even: None,
                    }),
                    ..RegionSection::default()
                },
                RegionSection {
                    header_footer_refs: Some(HeaderFooterRefs {
                        header_default: None,
                        header_first: None,
                        header_even: Some("header-even-b".to_owned()),
                        footer_default: None,
                        footer_first: None,
                        footer_even: None,
                    }),
                    ..RegionSection::default()
                },
            ],
            ..DocumentRegions::default()
        };

        let effective = effective_header_footer_refs(&regions, 1).unwrap();
        assert_eq!(effective.header_default.as_deref(), Some("header-a"));
        assert_eq!(effective.header_even.as_deref(), Some("header-even-b"));
        assert_eq!(effective.footer_default.as_deref(), Some("footer-a"));
    }

    #[test]
    fn authored_section_properties_normalize_to_engine_geometry() {
        let request: RegionLayoutInput = serde_json::from_value(json!({
            "bodyStory": "body",
            "regions": {
                "settings": {
                    "evenAndOddHeaders": true,
                    "footnotePr": {"numFmt": "upperRoman", "numStart": 3}
                },
                "sections": [{
                    "sectionId": "main",
                    "properties": {
                        "pageWidth": 12240,
                        "pageHeight": 15840,
                        "marginTop": 0,
                        "marginRight": 1440,
                        "marginBottom": 1440,
                        "marginLeft": 720,
                        "gutter": 360,
                        "headerDistance": 0,
                        "columnCount": 2,
                        "columnSpace": 360,
                        "equalWidth": false,
                        "columns": [{"width": 3600, "space": 360}, {"width": 7200}],
                        "headerReferences": [{"type": "default", "rId": "rId1"}],
                        "footnoteColumns": 2,
                        "sectionStart": "oddPage"
                    }
                }]
            },
            "renderEnv": {}
        }))
        .unwrap();

        let (_, regions, _, _, _, _) = request.split();
        let section = &regions.sections[0];

        assert_eq!(section.page_size.as_ref().unwrap().w, 816.0);
        assert_eq!(section.margins.as_ref().unwrap().top, 0.0);
        assert_eq!(section.margins.as_ref().unwrap().left, 72.0);
        assert_eq!(section.margins.as_ref().unwrap().header, Some(0.0));
        assert_eq!(section.columns.as_ref().unwrap().gap, 24.0);
        assert_eq!(
            section.columns.as_ref().unwrap().columns.as_ref().unwrap()[0].width,
            Some(240.0)
        );
        assert_eq!(
            section
                .header_footer_refs
                .as_ref()
                .unwrap()
                .header_default
                .as_deref(),
            Some("rId1")
        );
        assert_eq!(section.note_settings.footnote_columns, Some(2));
        assert_eq!(
            section.section_start,
            Some(crate::types::SectionBreakType::OddPage)
        );
        assert!(regions.even_and_odd_headers);
        assert_eq!(regions.note_settings.footnote.num_start, Some(3));
    }
}
