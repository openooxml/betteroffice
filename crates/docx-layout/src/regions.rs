use std::collections::BTreeMap;

use serde::{Deserialize, Serialize};
use serde_json::Value;

use crate::footnotes::NoteLayoutInput;
use crate::types::{HeaderFooterRefs, Input, Layout, NoteAreaContract};

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
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct RegionSection {
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub section_id: Option<String>,
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

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct PageNumbering {
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub start: Option<u64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub format: Option<String>,
}

#[derive(Debug, Deserialize)]
pub struct RegionLayoutInput {
    pub measured: Vec<crate::types::MeasuredBlock>,
    #[serde(default)]
    pub options: crate::types::LayoutOptions,
    #[serde(default)]
    pub regions: DocumentRegions,
    #[serde(default)]
    pub notes: NoteLayoutInput,
}

impl RegionLayoutInput {
    pub fn split(self) -> (Input, DocumentRegions, NoteLayoutInput) {
        (
            Input {
                measured: self.measured,
                options: self.options,
            },
            self.regions,
            self.notes,
        )
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
            page.header_footer_refs = section.header_footer_refs.clone();
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
}
