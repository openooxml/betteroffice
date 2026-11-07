//! shared quick-xml plumbing: depth-capped pull loop, text collection,
//! attribute lookup, and opc path helpers. malformed input errors, never panics.

use std::borrow::Cow;

use quick_xml::Reader;
use quick_xml::events::attributes::Attribute;
use quick_xml::events::{BytesStart, BytesText, Event};
use quick_xml::name::QName;

use crate::{MAX_DEPTH, ParseError};

/// length of a leading utf-8 byte order mark. quick-xml skips one without
/// counting it in `buffer_position`.
pub(crate) fn utf8_bom_len(data: &[u8]) -> usize {
    if data.starts_with(b"\xEF\xBB\xBF") {
        3
    } else {
        0
    }
}

/// build a reader over an in-memory part. empty elements expand to start+end;
/// dtd handling stays at quick-xml's safe default (no external entities).
pub(crate) fn reader(data: &[u8]) -> Reader<&[u8]> {
    let mut r = Reader::from_reader(data);
    let cfg = r.config_mut();
    cfg.expand_empty_elements = true;
    cfg.check_end_names = true;
    r
}

/// pull one owned event, maintaining nesting depth and rejecting input past
/// `MAX_DEPTH`.
pub(crate) fn next_event(
    reader: &mut Reader<&[u8]>,
    buf: &mut Vec<u8>,
    depth: &mut usize,
) -> Result<Event<'static>, ParseError> {
    buf.clear();
    let ev = reader.read_event_into(buf).map_err(xml_err)?;
    match &ev {
        Event::Start(_) => {
            *depth += 1;
            if *depth > MAX_DEPTH {
                return Err(ParseError::DepthExceeded);
            }
        }
        Event::End(_) => *depth = depth.saturating_sub(1),
        _ => {}
    }
    Ok(ev.into_owned())
}

/// collect all descendant text of the element whose start was just read,
/// flattening nested markup. `depth` must already reflect that start element.
pub(crate) fn collect_text(
    reader: &mut Reader<&[u8]>,
    buf: &mut Vec<u8>,
    depth: &mut usize,
) -> Result<String, ParseError> {
    let target = *depth;
    let mut out = String::new();
    loop {
        match next_event(reader, buf, depth)? {
            Event::Text(t) => out.push_str(&t.xml10_content().map_err(xml_err)?),
            Event::CData(t) => out.push_str(&t.xml10_content().map_err(xml_err)?),
            Event::GeneralRef(r) => {
                let name = r.decode().map_err(xml_err)?;
                out.push_str(&resolve_entity(&name)?);
            }
            Event::End(_) if *depth < target => break,
            Event::Eof => return Err(ParseError::Malformed("unexpected eof in text".into())),
            _ => {}
        }
    }
    Ok(out)
}

/// resolve a bare entity name (`amp`, `#65`, `#x41`) to its text; anything
/// needing a dtd is rejected.
pub(crate) fn resolve_entity(name: &str) -> Result<String, ParseError> {
    let raw = format!("&{name};");
    quick_xml::escape::unescape(&raw)
        .map(|c| c.into_owned())
        .map_err(|_| ParseError::Malformed(format!("unresolvable entity &{name};")))
}

/// Formula text, from a cell, a defined name or a conditional format, escaped
/// for element content. Excel reads `_xHHHH_` in a formula literally, so a CR
/// and any other control character go out as character references, the only
/// form in which a file can bring one into a formula; a NUL, which no reference
/// can carry, is written as it came.
pub(crate) fn formula_text(text: &str) -> String {
    char_references(&quick_xml::escape::escape(text), &['\r'])
}

/// [`formula_text`] as a writer event.
pub(crate) fn text_event(text: &str) -> BytesText<'static> {
    BytesText::from_escaped(formula_text(text))
}

/// `escaped` with `also` and every character XML 1.0 cannot carry literally
/// written as decimal character references.
fn char_references(escaped: &str, also: &[char]) -> String {
    let mut out = String::with_capacity(escaped.len());
    for ch in escaped.chars() {
        if ch != '\0' && (also.contains(&ch) || (!is_xml_char(ch) && !matches!(ch, '\t' | '\n'))) {
            out.push_str(&format!("&#{};", ch as u32));
        } else {
            out.push(ch);
        }
    }
    out
}

fn is_xml_char(ch: char) -> bool {
    ch >= ' ' && ch != '\u{fffe}' && ch != '\u{ffff}'
}

/// An attribute whose value keeps its CR, LF, tab and other control characters,
/// written as character references since literal ones read back as spaces.
pub(crate) fn text_attribute<'a>(name: &'a str, value: &str) -> Attribute<'a> {
    let escaped = char_references(&quick_xml::escape::escape(value), &['\r', '\n', '\t']);
    Attribute {
        key: QName(name.as_bytes()),
        value: Cow::Owned(escaped.into_bytes()),
    }
}

/// Encode string item text the way Excel writes ST_Xstring: a CR and the other
/// characters XML 1.0 cannot carry become `_xHHHH_`, and an underscore that
/// would start such an escape is written as `_x005F_`.
pub(crate) fn encode_xstring(text: &str) -> String {
    xstring(text, false)
}

/// A number format code (`ST_Xstring`) as element text, the way Excel writes
/// it: [`encode_xstring`], then xml escaped. Writing one never fails.
pub(crate) fn format_code_text(code: &str) -> String {
    quick_xml::escape::escape(xstring(code, false)).into_owned()
}

/// A number format code (`ST_Xstring`) as an attribute value, the way Excel
/// writes it: [`encode_xstring`] with tab and line feed also as `_xHHHH_`, so
/// attribute normalization cannot fold them into spaces, then xml escaped.
/// Writing one never fails.
pub(crate) fn format_code_attribute_value(code: &str) -> String {
    quick_xml::escape::escape(xstring(code, true)).into_owned()
}

/// [`format_code_attribute_value`] as a writer attribute.
pub(crate) fn format_code_attribute<'a>(name: &'a str, code: &str) -> Attribute<'a> {
    Attribute {
        key: QName(name.as_bytes()),
        value: Cow::Owned(format_code_attribute_value(code).into_bytes()),
    }
}

fn xstring(text: &str, attribute: bool) -> String {
    let bytes = text.as_bytes();
    let mut out = String::with_capacity(text.len());
    for (index, ch) in text.char_indices() {
        match ch {
            '_' if bytes
                .get(index + 1)
                .is_some_and(|b| matches!(b, b'x' | b'X'))
                && bytes.get(index + 6) == Some(&b'_')
                && bytes
                    .get(index + 2..index + 6)
                    .is_some_and(|hex| hex.iter().all(u8::is_ascii_hexdigit)) =>
            {
                out.push_str("_x005F_");
            }
            '\t' | '\n' if !attribute => out.push(ch),
            _ if !is_xml_char(ch) => {
                out.push_str(&format!("_x{:04X}_", ch as u32));
            }
            _ => out.push(ch),
        }
    }
    out
}

/// look up an attribute by its local name (namespace-prefix agnostic),
/// returning the unescaped value.
pub(crate) fn attr(e: &BytesStart, local: &[u8]) -> Result<Option<String>, ParseError> {
    for a in e.attributes() {
        let a = a.map_err(xml_err)?;
        if a.key.local_name().as_ref() == local {
            return Ok(Some(
                a.normalized_value(quick_xml::XmlVersion::Implicit1_0)
                    .map_err(xml_err)?
                    .into_owned(),
            ));
        }
    }
    Ok(None)
}

/// the element's local name (tag without any namespace prefix).
pub(crate) fn local_name(e: &BytesStart) -> Vec<u8> {
    e.name().local_name().as_ref().to_vec()
}

/// resolve a relationship `Target` against its declaring part, collapsing
/// `.`/`..` so a smuggled traversal cannot escape the package.
pub(crate) fn resolve_part_path(base_dir: &str, target: &str) -> String {
    if let Some(abs) = target.strip_prefix('/') {
        return abs.to_string();
    }
    let mut segs: Vec<&str> = base_dir.split('/').filter(|s| !s.is_empty()).collect();
    for part in target.split('/') {
        match part {
            "" | "." => {}
            ".." => {
                segs.pop();
            }
            p => segs.push(p),
        }
    }
    segs.join("/")
}

/// find a part by name, tolerating a leading slash on either side.
pub(crate) fn find_part<'a>(parts: &'a [(String, Vec<u8>)], name: &str) -> Option<&'a [u8]> {
    let want = name.trim_start_matches('/');
    parts
        .iter()
        .find(|(n, _)| n.trim_start_matches('/') == want)
        .map(|(_, b)| b.as_slice())
}

pub(crate) fn xml_err<E: core::fmt::Display>(e: E) -> ParseError {
    ParseError::Xml(e.to_string())
}
