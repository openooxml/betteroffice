mod fields;
mod fonts;
mod identifiers;
mod mask;
mod media;
mod media_parts;
mod rels;
mod schema;
mod scrub;
mod styles;
mod visio;
mod xml;

use std::collections::HashSet;
use std::fmt;

use thiserror::Error;

use crate::mask::TextMasker;
use crate::media::{MediaBudget, replace_media_budgeted};
use crate::scrub::{collect_custom_xml_parts, normalize_part_name, prune_scrubbed_parts};
use crate::styles::StyleMap;
use crate::xml::redact_xml_with_styles;

#[derive(Clone, Copy, Debug, Default, Eq, PartialEq)]
pub enum Format {
    #[default]
    Auto,
    Docx,
    Xlsx,
    Pptx,
    Vsdx,
    Vstx,
}

impl Format {
    pub fn extension(self) -> Option<&'static str> {
        match self {
            Self::Auto => None,
            Self::Docx => Some("docx"),
            Self::Xlsx => Some("xlsx"),
            Self::Pptx => Some("pptx"),
            Self::Vsdx => Some("vsdx"),
            Self::Vstx => Some("vstx"),
        }
    }
}

impl fmt::Display for Format {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Auto => formatter.write_str("OOXML"),
            Self::Docx => formatter.write_str("DOCX"),
            Self::Xlsx => formatter.write_str("XLSX"),
            Self::Pptx => formatter.write_str("PPTX"),
            Self::Vsdx => formatter.write_str("VSDX"),
            Self::Vstx => formatter.write_str("VSTX"),
        }
    }
}

#[derive(Clone, Debug, Default, Eq, PartialEq)]
pub struct RedactionReport {
    pub format: Format,
    pub text_nodes: usize,
    pub characters: usize,
    pub attributes: usize,
    pub media_parts: usize,
    pub binary_parts: usize,
    pub custom_xml_parts: usize,
    pub xml_comments: usize,
}

#[derive(Clone, Copy, Debug, Default, Eq, PartialEq)]
pub struct RedactionOptions {
    pub random_characters: bool,
    pub preserve_media_size: bool,
}

#[derive(Debug, Error)]
pub enum RedactError {
    #[error("invalid OOXML package: {0}")]
    Container(String),
    #[error("could not detect DOCX, XLSX, PPTX, VSDX, or VSTX content")]
    UnknownFormat,
    #[error(
        "unsupported or ambiguous Visio format; only VSDX drawings and VSTX templates can be redacted"
    )]
    UnsupportedVisio,
    #[error("cannot safely redact Visio part {part}: {message}")]
    AmbiguousVisio { part: String, message: String },
    #[error("requested {requested}, but package is {detected}")]
    FormatMismatch { requested: Format, detected: Format },
    #[error("invalid XML in {part}: {message}")]
    Xml { part: String, message: String },
    #[error("invalid style reference in {part}: {message}")]
    StyleReference { part: String, message: String },
    #[error("could not replace image {part}: {message}")]
    Image { part: String, message: String },
    #[error("could not obtain secure randomness: {0}")]
    Randomness(String),
}

pub fn detect_format(bytes: &[u8]) -> Result<Format, RedactError> {
    let parts = ooxml_opc::unzip_parts(bytes).map_err(RedactError::Container)?;
    detect_parts(&parts)
}

pub fn redact(bytes: &[u8], format: Format) -> Result<Vec<u8>, RedactError> {
    redact_with_report(bytes, format).map(|(bytes, _)| bytes)
}

pub fn redact_with_report(
    bytes: &[u8],
    requested: Format,
) -> Result<(Vec<u8>, RedactionReport), RedactError> {
    redact_with_report_and_options(bytes, requested, &RedactionOptions::default())
}

pub fn redact_with_options(
    bytes: &[u8],
    requested: Format,
    options: &RedactionOptions,
) -> Result<Vec<u8>, RedactError> {
    redact_with_report_and_options(bytes, requested, options).map(|(bytes, _)| bytes)
}

pub fn redact_with_report_and_options(
    bytes: &[u8],
    requested: Format,
    options: &RedactionOptions,
) -> Result<(Vec<u8>, RedactionReport), RedactError> {
    let mut parts = ooxml_opc::unzip_parts(bytes).map_err(RedactError::Container)?;
    let detected = detect_parts(&parts)?;
    if requested != Format::Auto && requested != detected {
        return Err(RedactError::FormatMismatch {
            requested,
            detected,
        });
    }

    if visio::is_visio(detected) {
        vsdx_parse::parse_vsdx(bytes)
            .map_err(|_| visio::ambiguous("package", "Visio parser rejected input"))?;
    }
    let mut report = RedactionReport {
        format: detected,
        ..RedactionReport::default()
    };
    let custom_xml = if matches!(detected, Format::Docx | Format::Xlsx | Format::Pptx) {
        collect_custom_xml_parts(&parts)?
    } else {
        HashSet::new()
    };
    report.custom_xml_parts = parts
        .iter()
        .filter(|(path, _)| custom_xml.contains(&normalize_part_name(path)))
        .count();
    let scrubbed: HashSet<String> = parts
        .iter()
        .map(|(path, _)| normalize_part_name(path))
        .filter(|name| {
            !custom_xml.contains(name) && !media::is_replaceable_part(name) && !is_xml_part(name)
        })
        .collect();
    report.binary_parts = scrubbed.len();
    if matches!(detected, Format::Docx | Format::Pptx) {
        fonts::detach_scrubbed_fonts(&mut parts, &scrubbed)?;
    }
    let blanked = prune_scrubbed_parts(&mut parts, &scrubbed, &custom_xml)?;
    media_parts::convert_wdp_parts(&mut parts)?;
    if visio::is_visio(detected) {
        report.attributes += visio::normalize_relationships(&mut parts)?;
    }
    let (main_styles, glossary_styles) = match detected {
        Format::Docx => (
            StyleMap::collect_docx(&parts, &["word/styles.xml", "word/styleswitheffects.xml"])?,
            StyleMap::collect_docx(
                &parts,
                &[
                    "word/glossary/styles.xml",
                    "word/glossary/styleswitheffects.xml",
                ],
            )?,
        ),
        Format::Xlsx => (StyleMap::collect_xlsx(&parts)?, StyleMap::default()),
        _ => Default::default(),
    };
    let mut masker = TextMasker::new(options);
    let mut media_budget = MediaBudget::default();
    for (path, data) in &mut parts {
        let canonical = normalize_part_name(path);
        if blanked.contains(&canonical) {
            data.clear();
        } else if !media::is_replaceable_part(&canonical) {
            let styles = if canonical.starts_with("word/glossary/") {
                &glossary_styles
            } else {
                &main_styles
            };
            *data = redact_xml_with_styles(
                detected,
                &canonical,
                data,
                &mut report,
                styles,
                &mut masker,
            )?;
            media_budget.reserve(&canonical, data.len() as u64)?;
        }
    }
    for (path, data) in &mut parts {
        let canonical = normalize_part_name(path);
        if !blanked.contains(&canonical) && media::is_replaceable_part(&canonical) {
            *data =
                replace_media_budgeted(&canonical, data, options, &mut media_budget, &mut report)?;
        }
    }

    let output = ooxml_opc::rezip_parts(&parts).map_err(RedactError::Container)?;
    if visio::is_visio(detected) {
        vsdx_parse::parse_vsdx(&output)
            .map_err(|_| visio::ambiguous("package", "Visio parser rejected redacted output"))?;
    }
    Ok((output, report))
}

fn detect_parts(parts: &[(String, Vec<u8>)]) -> Result<Format, RedactError> {
    if let Some((_, content_types)) = parts
        .iter()
        .find(|(path, _)| normalize_part_name(path) == "[content_types].xml")
    {
        let text = String::from_utf8_lossy(content_types).to_ascii_lowercase();
        let visio = visio::classify_content_types(content_types)?;
        if visio.refused {
            return Err(RedactError::UnsupportedVisio);
        }
        let office_main = text.contains("wordprocessingml.document.main+xml")
            || text.contains("ms-word.document.macroenabled.main+xml")
            || text.contains("spreadsheetml.sheet.main+xml")
            || text.contains("ms-excel.sheet.macroenabled.main+xml")
            || text.contains("presentationml.presentation.main+xml")
            || text.contains("ms-powerpoint.presentation.macroenabled.main+xml");
        if visio.accepted_drawing || visio.accepted_template {
            if office_main || visio.accepted_drawing && visio.accepted_template {
                return Err(RedactError::UnsupportedVisio);
            }
            return match ooxml_opc::detect_package_kind(parts) {
                Ok(ooxml_opc::DocumentKind::Vsdx) => Ok(Format::Vsdx),
                Ok(ooxml_opc::DocumentKind::Vstx) => Ok(Format::Vstx),
                _ => Err(RedactError::UnsupportedVisio),
            };
        }
        if parts
            .iter()
            .any(|(path, _)| normalize_part_name(path).starts_with("visio/"))
        {
            return Err(RedactError::UnsupportedVisio);
        }
        if text.contains("wordprocessingml.document.main+xml")
            || text.contains("ms-word.document.macroenabled.main+xml")
        {
            return Ok(Format::Docx);
        }
        if text.contains("spreadsheetml.sheet.main+xml")
            || text.contains("ms-excel.sheet.macroenabled.main+xml")
        {
            return Ok(Format::Xlsx);
        }
        if text.contains("presentationml.presentation.main+xml")
            || text.contains("ms-powerpoint.presentation.macroenabled.main+xml")
        {
            return Ok(Format::Pptx);
        }
    }

    if parts
        .iter()
        .any(|(path, _)| normalize_part_name(path).starts_with("visio/"))
    {
        return Err(RedactError::UnsupportedVisio);
    }
    let has = |expected: &str| {
        parts
            .iter()
            .any(|(path, _)| path.eq_ignore_ascii_case(expected))
    };
    if has("word/document.xml") {
        Ok(Format::Docx)
    } else if has("xl/workbook.xml") {
        Ok(Format::Xlsx)
    } else if has("ppt/presentation.xml") {
        Ok(Format::Pptx)
    } else {
        Err(RedactError::UnknownFormat)
    }
}

fn is_xml_part(path: &str) -> bool {
    path.ends_with(".xml") || path.ends_with(".rels") || path.ends_with(".vml")
}

#[cfg(test)]
mod tests;
#[cfg(test)]
mod tests_identifiers;
#[cfg(test)]
mod tests_media_size;
#[cfg(test)]
mod tests_shared_fields;
#[cfg(test)]
mod tests_style_xsd;
