mod fonts;
mod mask;
mod media;
mod rels;
mod schema;
mod scrub;
mod styles;
mod xml;

use std::collections::HashSet;
use std::fmt;

use thiserror::Error;

use crate::mask::TextMasker;
use crate::media::replace_media;
use crate::scrub::{normalize_part_name, prune_scrubbed_parts};
use crate::styles::StyleMap;
use crate::xml::redact_xml_with_styles;

#[derive(Clone, Copy, Debug, Default, Eq, PartialEq)]
pub enum Format {
    #[default]
    Auto,
    Docx,
    Xlsx,
    Pptx,
}

impl Format {
    pub fn extension(self) -> Option<&'static str> {
        match self {
            Self::Auto => None,
            Self::Docx => Some("docx"),
            Self::Xlsx => Some("xlsx"),
            Self::Pptx => Some("pptx"),
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
    pub xml_comments: usize,
}

#[derive(Clone, Copy, Debug, Default, Eq, PartialEq)]
pub struct RedactionOptions {
    pub random_characters: bool,
}

#[derive(Debug, Error)]
pub enum RedactError {
    #[error("invalid OOXML package: {0}")]
    Container(String),
    #[error("could not detect DOCX, XLSX, or PPTX content")]
    UnknownFormat,
    #[error("Visio redaction is not supported; no safe redacted package can be produced")]
    UnsupportedVisio,
    #[error("requested {requested}, but package is {detected}")]
    FormatMismatch { requested: Format, detected: Format },
    #[error("invalid XML in {part}: {message}")]
    Xml { part: String, message: String },
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

    let mut report = RedactionReport {
        format: detected,
        ..RedactionReport::default()
    };
    let scrubbed: HashSet<String> = parts
        .iter()
        .map(|(path, _)| normalize_part_name(path))
        .filter(|name| !media::is_replaceable_part(name) && !is_xml_part(name))
        .collect();
    report.binary_parts = scrubbed.len();
    if matches!(detected, Format::Docx | Format::Pptx) {
        fonts::detach_scrubbed_fonts(&mut parts, &scrubbed)?;
    }
    let blanked = if scrubbed.is_empty() {
        HashSet::new()
    } else {
        prune_scrubbed_parts(&mut parts, &scrubbed)?
    };
    let collect_styles = |name: &str| {
        parts
            .iter()
            .find(|(path, _)| detected == Format::Docx && normalize_part_name(path) == name)
            .map(|(_, bytes)| StyleMap::collect(name, bytes))
            .transpose()
            .map(Option::unwrap_or_default)
    };
    let main_styles = collect_styles("word/styles.xml")?;
    let glossary_styles = collect_styles("word/glossary/styles.xml")?;
    let mut masker = TextMasker::new(options);
    for (path, data) in &mut parts {
        let canonical = normalize_part_name(path);
        if blanked.contains(&canonical) {
            data.clear();
        } else if media::is_replaceable_part(&canonical) {
            *data = replace_media(&canonical, data, &mut report)?;
        } else {
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
        }
    }

    let output = ooxml_opc::rezip_parts(&parts).map_err(RedactError::Container)?;
    Ok((output, report))
}

fn detect_parts(parts: &[(String, Vec<u8>)]) -> Result<Format, RedactError> {
    if parts
        .iter()
        .any(|(path, _)| normalize_part_name(path).starts_with("visio/"))
    {
        return Err(RedactError::UnsupportedVisio);
    }
    if let Some((_, content_types)) = parts
        .iter()
        .find(|(path, _)| normalize_part_name(path) == "[content_types].xml")
    {
        let text = String::from_utf8_lossy(content_types).to_ascii_lowercase();
        if declares_visio_main_part(content_types)? {
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

fn declares_visio_main_part(bytes: &[u8]) -> Result<bool, RedactError> {
    use quick_xml::events::Event;
    use quick_xml::name::ResolveResult;
    use quick_xml::{NsReader, XmlVersion};

    let mut reader = NsReader::from_reader(bytes);
    let error = |message: String| RedactError::Xml {
        part: "[Content_Types].xml".to_owned(),
        message,
    };
    loop {
        match reader
            .read_event()
            .map_err(|value| error(value.to_string()))?
        {
            Event::Start(start) | Event::Empty(start)
                if matches!(start.local_name().as_ref(), b"Override" | b"Default") =>
            {
                let namespace = reader.resolver().resolve_element(start.name()).0;
                if !matches!(namespace, ResolveResult::Unbound)
                    && !matches!(namespace, ResolveResult::Bound(ns) if ns.as_ref() == b"http://schemas.openxmlformats.org/package/2006/content-types")
                {
                    continue;
                }
                for attribute in start.attributes() {
                    let attribute = attribute.map_err(|value| error(value.to_string()))?;
                    if attribute.key.as_ref() == b"ContentType" {
                        let value = attribute
                            .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
                            .map_err(|value| error(value.to_string()))?
                            .to_ascii_lowercase();
                        if value.starts_with("application/vnd.ms-visio.")
                            && value.ends_with(".main+xml")
                        {
                            return Ok(true);
                        }
                    }
                }
            }
            Event::Eof => return Ok(false),
            _ => {}
        }
    }
}

fn is_xml_part(path: &str) -> bool {
    path.ends_with(".xml") || path.ends_with(".rels") || path.ends_with(".vml")
}

#[cfg(test)]
mod tests;
