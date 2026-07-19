mod media;
mod xml;

use std::fmt;

use thiserror::Error;

use crate::media::replace_media;
use crate::xml::redact_xml;

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

#[derive(Debug, Error)]
pub enum RedactError {
    #[error("invalid OOXML package: {0}")]
    Container(String),
    #[error("could not detect DOCX, XLSX, or PPTX content")]
    UnknownFormat,
    #[error("requested {requested}, but package is {detected}")]
    FormatMismatch { requested: Format, detected: Format },
    #[error("invalid XML in {part}: {message}")]
    Xml { part: String, message: String },
    #[error("could not replace image {part}: {message}")]
    Image { part: String, message: String },
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
    for (path, data) in &mut parts {
        let lower = path.to_ascii_lowercase();
        if media::is_replaceable_part(&lower) {
            *data = replace_media(path, data, &mut report)?;
        } else if is_xml_part(&lower) {
            *data = redact_xml(detected, path, data, &mut report)?;
        } else if is_sensitive_binary(&lower) {
            data.clear();
            report.binary_parts += 1;
        }
    }

    let output = ooxml_opc::rezip_parts(&parts).map_err(RedactError::Container)?;
    Ok((output, report))
}

fn detect_parts(parts: &[(String, Vec<u8>)]) -> Result<Format, RedactError> {
    if let Some((_, content_types)) = parts
        .iter()
        .find(|(path, _)| path.eq_ignore_ascii_case("[Content_Types].xml"))
    {
        let text = String::from_utf8_lossy(content_types).to_ascii_lowercase();
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

fn is_xml_part(path: &str) -> bool {
    path.ends_with(".xml") || path.ends_with(".rels") || path.ends_with(".vml")
}

fn is_sensitive_binary(path: &str) -> bool {
    path.ends_with("vbaproject.bin")
        || path.contains("/embeddings/")
        || path.contains("/activex/") && path.ends_with(".bin")
        || path.contains("/printersettings/")
}

#[cfg(test)]
mod tests;
