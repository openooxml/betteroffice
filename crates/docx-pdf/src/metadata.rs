//! PDF/A document metadata: the trailer `/Info` dictionary and the catalog
//! XMP packet.
//!
//! The caller reserves both object ids from the document's single id counter.
//! [`write_document_info`] fills the trailer `/Info` object and
//! [`write_xmp_metadata`] fills the catalog `/Metadata` stream. Only fields the
//! caller supplied are emitted; absent values are omitted rather than filled
//! with invented defaults, and the output contains no wall-clock timestamps, so
//! repeated exports of the same input are byte-identical.

use pdf_writer::{Finish, Name, Pdf, Ref, TextStr};

use crate::pdfa_limits::{PDFA_STRING_TOO_LONG, text_too_long};
use crate::{PdfConformance, PdfError, PdfMetadata};

/// Fixed `/Producer` value shared by `/Info` and `pdf:Producer`.
const PRODUCER: &str = concat!("BetterOffice ", env!("CARGO_PKG_VERSION"));

/// Error text for a string that does not follow the PDF date grammar.
const INVALID_DATE: &str = "document metadata date";

/// Error text for a character XML 1.0 cannot represent.
const INVALID_XML_CHAR: &str = "metadata contains a character XML cannot represent";

/// Fill the reserved trailer `/Info` dictionary.
///
/// The standard entries are written only when the corresponding
/// [`PdfMetadata`] field is present. `/Producer` is always written from the
/// compiled product version. `CreationDate`/`ModDate` are stored as PDF date
/// strings, validated here, and written byte-for-byte as supplied.
pub(crate) fn write_document_info(
    pdf: &mut Pdf,
    id: Ref,
    metadata: &PdfMetadata,
) -> Result<(), PdfError> {
    // Validate before touching the writer so an invalid date cannot leave a
    // half-written `/Info` object behind.
    let creation = metadata
        .creation_date
        .as_deref()
        .map(parse_pdf_date)
        .transpose()?;
    let modification = metadata
        .modification_date
        .as_deref()
        .map(parse_pdf_date)
        .transpose()?;

    // PDF/A-2 6.1.13 limits every PDF string. Validate all fields before
    // touching the writer so a rejected export cannot leave a partial object.
    for value in [
        metadata.title.as_deref(),
        metadata.author.as_deref(),
        metadata.subject.as_deref(),
        metadata.keywords.as_deref(),
        metadata.creator.as_deref(),
        creation.as_ref().map(|date| date.raw),
        modification.as_ref().map(|date| date.raw),
    ]
    .into_iter()
    .flatten()
    {
        if text_too_long(value) {
            return Err(PdfError::Unsupported(PDFA_STRING_TOO_LONG));
        }
    }

    let mut info = pdf.document_info(id);
    if let Some(title) = &metadata.title {
        info.title(TextStr(title));
    }
    if let Some(author) = &metadata.author {
        info.author(TextStr(author));
    }
    if let Some(subject) = &metadata.subject {
        info.subject(TextStr(subject));
    }
    if let Some(keywords) = &metadata.keywords {
        info.keywords(TextStr(keywords));
    }
    if let Some(creator) = &metadata.creator {
        info.creator(TextStr(creator));
    }
    info.producer(TextStr(PRODUCER));
    // `DocumentInfo` exposes pdf-writer's typed `Date`, which normalizes the
    // timezone (`+00'00'` becomes `Z`) and clamps out-of-range fields. The
    // dictionary derefs to `Dict`, so validate the PDF grammar above and write
    // the caller's exact (ASCII) date characters as a text string.
    if let Some(date) = creation {
        info.pair(Name(b"CreationDate"), TextStr(date.raw));
    }
    if let Some(date) = modification {
        info.pair(Name(b"ModDate"), TextStr(date.raw));
    }
    info.finish();
    Ok(())
}

/// Fill the reserved catalog `/Metadata` XMP stream.
///
/// Only [`PdfConformance::PdfA2u`] is supported: `Standard` output makes no
/// archival claim and must not carry a `pdfaid` packet. The packet identifies
/// PDF/A-2, conformance level U, and mirrors every present `/Info` entry.
pub(crate) fn write_xmp_metadata(
    pdf: &mut Pdf,
    id: Ref,
    conformance: PdfConformance,
    metadata: &PdfMetadata,
) -> Result<(), PdfError> {
    let packet = match conformance {
        PdfConformance::PdfA2u => xmp_packet(metadata)?,
        PdfConformance::Standard => {
            return Err(PdfError::Unsupported(
                "XMP metadata is only written for PDF/A-2u",
            ));
        }
    };
    pdf.metadata(id, packet.as_bytes()).finish();
    Ok(())
}

/// A validated PDF date, split into the fields needed for XMP.
#[derive(Debug, Clone, Copy)]
struct PdfDate<'a> {
    /// The original string, written verbatim into `/Info`.
    raw: &'a str,
    year: u16,
    month: Option<u8>,
    day: Option<u8>,
    hour: Option<u8>,
    minute: Option<u8>,
    second: Option<u8>,
    /// Present only after a complete `YYYYMMDDHHmmSS` time.
    offset: Option<UtcOffset>,
}

/// The `OHH'mm'` field of a PDF date.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum UtcOffset {
    /// `Z`.
    Zulu,
    /// `+HH'mm'`.
    Plus { hours: u8, minutes: u8 },
    /// `-HH'mm'`.
    Minus { hours: u8, minutes: u8 },
}

/// Validate a PDF date string and split it into fields.
///
/// The accepted grammar is ISO 32000-1 §7.9.4: `D:YYYYMMDDHHmmSSOHH'mm'`,
/// where year is required, later fields are optional but cannot be skipped,
/// and a timezone (which may only appear after a complete `…SS` time) is `Z`
/// or `+HH'mm'`/`-HH'mm'`. Out-of-range values (month 13, 30 February, hour
/// 24, offset minute 60) are rejected rather than normalized.
fn parse_pdf_date(raw: &str) -> Result<PdfDate<'_>, PdfError> {
    let value = raw
        .strip_prefix("D:")
        .ok_or(PdfError::Unsupported(INVALID_DATE))?;

    let digit_len = value
        .bytes()
        .take_while(|byte| byte.is_ascii_digit())
        .count();
    let (digits, suffix) = value.split_at(digit_len);
    if !(4..=14).contains(&digit_len) || digit_len % 2 != 0 {
        return Err(PdfError::Unsupported(INVALID_DATE));
    }
    let bytes = digits.as_bytes();
    let pair = |start: usize| -> Result<u8, PdfError> {
        two_digits(&bytes[start..start + 2]).ok_or(PdfError::Unsupported(INVALID_DATE))
    };

    let year = u16::from(pair(0)?) * 100 + u16::from(pair(2)?);
    let month = (digit_len >= 6).then(|| pair(4)).transpose()?;
    let day = (digit_len >= 8).then(|| pair(6)).transpose()?;
    let hour = (digit_len >= 10).then(|| pair(8)).transpose()?;
    let minute = (digit_len >= 12).then(|| pair(10)).transpose()?;
    let second = (digit_len >= 14).then(|| pair(12)).transpose()?;

    if let Some(month) = month {
        if !(1..=12).contains(&month) {
            return Err(PdfError::Unsupported(INVALID_DATE));
        }
    }
    if let Some(day) = day {
        let month = month.ok_or(PdfError::Unsupported(INVALID_DATE))?;
        if day == 0 || day > days_in_month(year, month) {
            return Err(PdfError::Unsupported(INVALID_DATE));
        }
    }
    if hour.is_some_and(|hour| hour > 23)
        || minute.is_some_and(|minute| minute > 59)
        || second.is_some_and(|second| second > 59)
    {
        return Err(PdfError::Unsupported(INVALID_DATE));
    }

    let offset = match suffix {
        "" => None,
        "Z" if digit_len == 14 => Some(UtcOffset::Zulu),
        _ => {
            let suffix = suffix.as_bytes();
            if digit_len != 14
                || suffix.len() != 7
                || !matches!(suffix[0], b'+' | b'-')
                || suffix[3] != b'\''
                || suffix[6] != b'\''
            {
                return Err(PdfError::Unsupported(INVALID_DATE));
            }
            let hours = two_digits(&suffix[1..3]).ok_or(PdfError::Unsupported(INVALID_DATE))?;
            let minutes = two_digits(&suffix[4..6]).ok_or(PdfError::Unsupported(INVALID_DATE))?;
            if hours > 23 || minutes > 59 {
                return Err(PdfError::Unsupported(INVALID_DATE));
            }
            Some(if suffix[0] == b'+' {
                UtcOffset::Plus { hours, minutes }
            } else {
                UtcOffset::Minus { hours, minutes }
            })
        }
    };

    Ok(PdfDate {
        raw,
        year,
        month,
        day,
        hour,
        minute,
        second,
        offset,
    })
}

/// Parse exactly two ASCII digits.
fn two_digits(bytes: &[u8]) -> Option<u8> {
    match bytes {
        [tens, ones] if tens.is_ascii_digit() && ones.is_ascii_digit() => {
            Some((*tens - b'0') * 10 + (*ones - b'0'))
        }
        _ => None,
    }
}

/// Days in `month`, honoring the Gregorian leap-year rule.
fn days_in_month(year: u16, month: u8) -> u8 {
    match month {
        1 | 3 | 5 | 7 | 8 | 10 | 12 => 31,
        4 | 6 | 9 | 11 => 30,
        2 if year % 4 == 0 && (year % 100 != 0 || year % 400 == 0) => 29,
        2 => 28,
        _ => 0,
    }
}

/// Convert a validated PDF date to the ISO 8601 form XMP date properties use.
///
/// The same precision as the source is kept: `D:2024` becomes `2024`,
/// `D:20240102030405+02'00'` becomes `2024-01-02T03:04:05+02:00`. No time is
/// invented for a partial date.
fn xmp_date(date: &PdfDate<'_>) -> String {
    let mut value = format!("{:04}", date.year);
    let Some(month) = date.month else {
        return value;
    };
    value.push_str(&format!("-{month:02}"));
    let Some(day) = date.day else {
        return value;
    };
    value.push_str(&format!("-{day:02}"));
    let Some(hour) = date.hour else {
        return value;
    };
    value.push_str(&format!("T{hour:02}"));
    let minute = date.minute.unwrap_or(0);
    value.push_str(&format!(":{minute:02}"));
    let Some(second) = date.second else {
        return value;
    };
    value.push_str(&format!(":{second:02}"));
    match date.offset {
        Some(UtcOffset::Zulu) => value.push('Z'),
        Some(UtcOffset::Plus { hours, minutes }) => {
            value.push_str(&format!("+{hours:02}:{minutes:02}"));
        }
        Some(UtcOffset::Minus { hours, minutes }) => {
            value.push_str(&format!("-{hours:02}:{minutes:02}"));
        }
        None => {}
    }
    value
}

/// Build a deterministic XMP packet for PDF/A-2u.
///
/// Element order is fixed. Every value is XML-escaped and UTF-8 encoded, and
/// absent metadata fields produce no element at all.
fn xmp_packet(metadata: &PdfMetadata) -> Result<String, PdfError> {
    // Validate dates before building anything so a bad value fails cleanly.
    let creation = metadata
        .creation_date
        .as_deref()
        .map(parse_pdf_date)
        .transpose()?;
    let modification = metadata
        .modification_date
        .as_deref()
        .map(parse_pdf_date)
        .transpose()?;

    let mut xml = String::with_capacity(1024);
    // PDF/A-2 6.6.2.1 forbids an `encoding` attribute in the packet header;
    // XMP defaults to UTF-8, and the stream bytes below are UTF-8.
    xml.push_str("<?xml version=\"1.0\"?>\n");
    // `x` names the standard XMP packet wrapper; the property namespaces are
    // `rdf`, `dc`, `xmp`, `pdf` and `pdfaid`.
    xml.push_str("<x:xmpmeta xmlns:x=\"adobe:ns:meta/\" x:xmptk=\"");
    escape_xml(PRODUCER, &mut xml)?;
    xml.push_str("\">\n");
    xml.push_str("<rdf:RDF xmlns:rdf=\"http://www.w3.org/1999/02/22-rdf-syntax-ns#\">\n");
    xml.push_str(concat!(
        "<rdf:Description rdf:about=\"\"",
        " xmlns:dc=\"http://purl.org/dc/elements/1.1/\"",
        " xmlns:xmp=\"http://ns.adobe.com/xap/1.0/\"",
        " xmlns:pdf=\"http://ns.adobe.com/pdf/1.3/\"",
        " xmlns:pdfaid=\"http://www.aiim.org/pdfa/ns/id/\">\n",
    ));

    if let Some(title) = &metadata.title {
        xml.push_str("<dc:title><rdf:Alt><rdf:li xml:lang=\"x-default\">");
        escape_xml(title, &mut xml)?;
        xml.push_str("</rdf:li></rdf:Alt></dc:title>\n");
    }
    if let Some(author) = &metadata.author {
        xml.push_str("<dc:creator><rdf:Seq><rdf:li>");
        escape_xml(author, &mut xml)?;
        xml.push_str("</rdf:li></rdf:Seq></dc:creator>\n");
    }
    if let Some(subject) = &metadata.subject {
        xml.push_str("<dc:description><rdf:Alt><rdf:li xml:lang=\"x-default\">");
        escape_xml(subject, &mut xml)?;
        xml.push_str("</rdf:li></rdf:Alt></dc:description>\n");
    }
    if let Some(keywords) = &metadata.keywords {
        xml.push_str("<pdf:Keywords>");
        escape_xml(keywords, &mut xml)?;
        xml.push_str("</pdf:Keywords>\n");
    }
    if let Some(creator) = &metadata.creator {
        xml.push_str("<xmp:CreatorTool>");
        escape_xml(creator, &mut xml)?;
        xml.push_str("</xmp:CreatorTool>\n");
    }
    if let Some(date) = creation {
        xml.push_str("<xmp:CreateDate>");
        xml.push_str(&xmp_date(&date));
        xml.push_str("</xmp:CreateDate>\n");
    }
    if let Some(date) = modification {
        xml.push_str("<xmp:ModifyDate>");
        xml.push_str(&xmp_date(&date));
        xml.push_str("</xmp:ModifyDate>\n");
    }
    xml.push_str("<pdf:Producer>");
    escape_xml(PRODUCER, &mut xml)?;
    xml.push_str("</pdf:Producer>\n");
    xml.push_str("<pdfaid:part>2</pdfaid:part>\n");
    xml.push_str("<pdfaid:conformance>U</pdfaid:conformance>\n");
    xml.push_str("</rdf:Description>\n</rdf:RDF>\n</x:xmpmeta>\n");
    Ok(xml)
}

/// Escape a value for XML text and attribute content and append it to `out`.
///
/// The five XML metacharacters are escaped. Characters that XML 1.0 cannot
/// represent are rejected instead of producing a malformed packet. Tab and
/// line feed remain literal. Carriage return is emitted as a character
/// reference because XML line-ending normalization would otherwise turn it
/// into LF and make the XMP value differ from the `/Info` string.
fn escape_xml(value: &str, out: &mut String) -> Result<(), PdfError> {
    for character in value.chars() {
        match character {
            '&' => out.push_str("&amp;"),
            '<' => out.push_str("&lt;"),
            '>' => out.push_str("&gt;"),
            '"' => out.push_str("&quot;"),
            '\'' => out.push_str("&apos;"),
            '\t' | '\n' => out.push(character),
            '\r' => out.push_str("&#13;"),
            '\u{0}'..='\u{8}'
            | '\u{b}'
            | '\u{c}'
            | '\u{e}'..='\u{1f}'
            | '\u{fffe}'
            | '\u{ffff}' => return Err(PdfError::Unsupported(INVALID_XML_CHAR)),
            _ => out.push(character),
        }
    }
    Ok(())
}
