//! shared quick-xml plumbing: depth-capped pull loop, text collection,
//! attribute lookup, and opc path helpers. malformed input errors, never panics.

use quick_xml::Reader;
use quick_xml::events::{BytesStart, Event};

use crate::{MAX_DEPTH, ParseError};

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
            Event::Text(t) => out.push_str(&t.decode().map_err(xml_err)?),
            Event::CData(t) => out.push_str(&t.decode().map_err(xml_err)?),
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
fn resolve_entity(name: &str) -> Result<String, ParseError> {
    let raw = format!("&{name};");
    quick_xml::escape::unescape(&raw)
        .map(|c| c.into_owned())
        .map_err(|_| ParseError::Malformed(format!("unresolvable entity &{name};")))
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
