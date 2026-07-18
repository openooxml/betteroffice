//! Bounded `word/fontTable.xml` leaf parser.

use serde::{Deserialize, Serialize};

use crate::settings::incumbent_utf8_text_boundary;
use crate::xml::{ParseBudget, ParseError, XmlElement, parse_xml};

#[derive(Clone, Debug, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct FontEmbed {
    pub rel_id: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub font_key: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub subsetted: Option<bool>,
}

#[derive(Clone, Debug, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct FontInfo {
    pub name: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub alt_name: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub panose1: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub charset: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub family: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub pitch: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub embed_regular: Option<FontEmbed>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub embed_bold: Option<FontEmbed>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub embed_italic: Option<FontEmbed>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub embed_bold_italic: Option<FontEmbed>,
}

#[derive(Clone, Debug, Default, PartialEq, Eq, Serialize, Deserialize)]
pub struct FontTable {
    pub fonts: Vec<FontInfo>,
}

pub fn parse_font_table(
    xml: Option<&[u8]>,
    part: &str,
    budget: &mut ParseBudget<'_>,
) -> Result<FontTable, ParseError> {
    let Some(xml) = xml.filter(|value| !value.iter().all(u8::is_ascii_whitespace)) else {
        return Ok(FontTable::default());
    };
    if !incumbent_utf8_text_boundary(xml) {
        return Ok(FontTable::default());
    }
    let document = parse_xml(xml, part, budget)?;
    Ok(parse_font_table_element(document.root()))
}

pub fn parse_font_table_element(root: Option<&XmlElement>) -> FontTable {
    let fonts = root
        .into_iter()
        .flat_map(|root| root.children_named("w", "font"))
        .filter_map(parse_font_info)
        .collect();
    FontTable { fonts }
}

fn parse_font_info(font: &XmlElement) -> Option<FontInfo> {
    let name = font.attribute(Some("w"), "name")?;
    if name.is_empty() {
        return None;
    }
    Some(FontInfo {
        name: name.to_owned(),
        alt_name: child_value(font, "altName"),
        panose1: child_value(font, "panose1"),
        charset: child_value(font, "charset"),
        family: child_value(font, "family").filter(|value| {
            matches!(
                value.as_str(),
                "decorative" | "modern" | "roman" | "script" | "swiss" | "auto"
            )
        }),
        pitch: child_value(font, "pitch")
            .filter(|value| matches!(value.as_str(), "default" | "fixed" | "variable")),
        embed_regular: parse_embed(font, "embedRegular"),
        embed_bold: parse_embed(font, "embedBold"),
        embed_italic: parse_embed(font, "embedItalic"),
        embed_bold_italic: parse_embed(font, "embedBoldItalic"),
    })
}

fn parse_embed(font: &XmlElement, name: &str) -> Option<FontEmbed> {
    let element = font.child("w", name)?;
    let rel_id = element.attribute(Some("r"), "id")?;
    if rel_id.is_empty() {
        return None;
    }
    Some(FontEmbed {
        rel_id: rel_id.to_owned(),
        font_key: element
            .attribute(Some("w"), "fontKey")
            .filter(|value| !value.is_empty())
            .map(str::to_owned),
        subsetted: matches!(
            element.attribute(Some("w"), "subsetted"),
            Some("true" | "1")
        )
        .then_some(true),
    })
}

fn child_value(font: &XmlElement, name: &str) -> Option<String> {
    font.child("w", name)
        .and_then(|element| element.attribute(Some("w"), "val"))
        .filter(|value| !value.is_empty())
        .map(str::to_owned)
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::xml::ParseLimits;

    fn parse(xml: Option<&str>) -> Result<FontTable, ParseError> {
        let limits = ParseLimits::default();
        parse_font_table(
            xml.map(str::as_bytes),
            "word/fontTable.xml",
            &mut ParseBudget::new(&limits),
        )
    }

    #[test]
    fn parses_metadata_faces_and_incumbent_enum_filtering() {
        let table = parse(Some(
            r#"<w:fonts><w:font w:name="Calibri"><w:altName w:val="Arial"/><w:family w:val="swiss"/><w:pitch w:val="variable"/><w:charset w:val="00"/><w:panose1 w:val="020F"/><w:embedRegular r:id="rId1" w:fontKey="{GUID}" w:subsetted="1"/><w:embedBold r:id=""/></w:font><w:font w:name="Future"><w:family w:val="future"/><w:pitch w:val="future"/></w:font><w:font/></w:fonts>"#,
        ))
        .unwrap();
        assert_eq!(table.fonts.len(), 2);
        let first = &table.fonts[0];
        assert_eq!(first.alt_name.as_deref(), Some("Arial"));
        assert_eq!(first.embed_regular.as_ref().unwrap().subsetted, Some(true));
        assert!(first.embed_bold.is_none());
        assert!(table.fonts[1].family.is_none());
    }

    #[test]
    fn empty_input_is_empty_and_dtd_is_rejected() {
        assert_eq!(parse(None).unwrap(), FontTable::default());
        assert_eq!(parse(Some("   \n")).unwrap(), FontTable::default());
        assert!(matches!(
            parse(Some("<!DOCTYPE x><w:fonts/>")),
            Err(ParseError::UnsafeXml { .. })
        ));
    }
}
