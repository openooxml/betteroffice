use std::collections::BTreeMap;

use quick_xml::events::Event;
use quick_xml::name::ResolveResult;
use quick_xml::{NsReader, Writer, XmlVersion};

use crate::RedactError;
use crate::scrub::normalize_part_name;

type IdMap = BTreeMap<String, String>;

pub(crate) fn normalize_relationships(
    parts: &mut [(String, Vec<u8>)],
) -> Result<usize, RedactError> {
    let mut relationships = BTreeMap::new();
    for (path, bytes) in parts.iter() {
        let path = normalize_part_name(path);
        if path.ends_with(".rels") && relationship_owner(&path).is_none() {
            return Err(super::ambiguous(
                &path,
                "unrecognized relationship part path",
            ));
        }
        if let Some(owner) = relationship_owner(&path) {
            let mut ids = IdMap::new();
            rewrite(bytes, &path, |element, key, value, namespace| {
                if element == "Relationship" && key == "Id" && namespace.is_none() {
                    let next = format!("rId{}", ids.len() + 1);
                    if value.is_empty() || ids.insert(value.to_owned(), next).is_some() {
                        return Err(super::ambiguous(
                            &path,
                            "duplicate or empty relationship id",
                        ));
                    }
                }
                Ok(None)
            })?;
            relationships.insert(owner, ids);
        }
    }
    let mut rows = IdMap::new();
    let mut changed = 0;
    for (path, bytes) in parts {
        let path = normalize_part_name(path);
        if !crate::is_xml_part(&path) {
            continue;
        }
        let owner = relationship_owner(&path);
        let ids = relationships.get(owner.as_deref().unwrap_or(&path));
        *bytes = rewrite(bytes, &path, |element, key, value, namespace| {
            let declaration = owner.is_some() && element == "Relationship" && key == "Id";
            let reference = namespace.is_some_and(super::relationship_namespace);
            if declaration || reference {
                let replacement = ids
                    .and_then(|ids| ids.get(value))
                    .cloned()
                    .ok_or_else(|| super::ambiguous(&path, "unresolved relationship reference"))?;
                changed += usize::from(replacement != value);
                return Ok(Some(replacement));
            }
            if element == "Row" && key == "N" && namespace.is_none() {
                let next = format!("Row{}", rows.len() + 1);
                let replacement = rows.entry(value.to_owned()).or_insert(next).clone();
                changed += usize::from(replacement != value);
                return Ok(Some(replacement));
            }
            Ok(None)
        })?;
    }
    Ok(changed)
}

fn relationship_owner(path: &str) -> Option<String> {
    if path == "_rels/.rels" {
        return Some(String::new());
    }
    let path = path.strip_suffix(".rels")?;
    if let Some(file) = path.strip_prefix("_rels/") {
        return Some(file.to_owned());
    }
    let (directory, file) = path.rsplit_once("/_rels/")?;
    Some(format!("{directory}/{file}"))
}

fn rewrite(
    bytes: &[u8],
    path: &str,
    mut attribute: impl FnMut(&str, &str, &str, Option<&[u8]>) -> Result<Option<String>, RedactError>,
) -> Result<Vec<u8>, RedactError> {
    let mut reader = NsReader::from_reader(bytes);
    let mut writer = Writer::new(Vec::with_capacity(bytes.len()));
    let mut depth = 0_usize;
    let mut root_seen = false;
    let error = |message: String| RedactError::Xml {
        part: path.into(),
        message,
    };
    loop {
        let event = reader
            .read_event()
            .map_err(|value| error(value.to_string()))?;
        let empty = matches!(event, Event::Empty(_));
        match event {
            Event::Start(start) | Event::Empty(start) => {
                let element = String::from_utf8_lossy(start.local_name().as_ref()).into_owned();
                if depth == 0 {
                    if root_seen {
                        return Err(super::ambiguous(path, "multiple XML roots"));
                    }
                    root_seen = true;
                    if path.ends_with(".rels") && element != "Relationships" {
                        return Err(super::ambiguous(path, "invalid relationship part root"));
                    }
                }
                if !empty {
                    depth += 1;
                    if depth > 256 {
                        return Err(super::ambiguous(path, "XML depth limit exceeded"));
                    }
                }
                let mut output = start.to_owned();
                output.clear_attributes();
                for attr in start.attributes() {
                    let attr = attr.map_err(|value| error(value.to_string()))?;
                    let key = String::from_utf8_lossy(attr.key.as_ref());
                    let value = attr
                        .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
                        .map_err(|value| error(value.to_string()))?;
                    let (namespace, _) = reader.resolver().resolve_attribute(attr.key);
                    let namespace = match &namespace {
                        ResolveResult::Bound(ns) => Some(ns.as_ref()),
                        ResolveResult::Unbound => None,
                        ResolveResult::Unknown(_) => {
                            return Err(super::ambiguous(path, "unbound attribute namespace"));
                        }
                    };
                    let replacement = attribute(&element, &key, &value, namespace)?;
                    output.push_attribute((key.as_ref(), replacement.as_deref().unwrap_or(&value)));
                }
                writer
                    .write_event(if empty {
                        Event::Empty(output)
                    } else {
                        Event::Start(output)
                    })
                    .map_err(|value| error(value.to_string()))?;
            }
            Event::DocType(_) => return Err(super::ambiguous(path, "DTD is forbidden")),
            Event::End(end) => {
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| super::ambiguous(path, "unmatched XML end"))?;
                writer
                    .write_event(Event::End(end))
                    .map_err(|value| error(value.to_string()))?;
            }
            Event::Eof => {
                if depth != 0 || !root_seen {
                    return Err(super::ambiguous(path, "incomplete XML document"));
                }
                break;
            }
            Event::Text(ref text)
                if depth == 0 && text.iter().any(|byte| !byte.is_ascii_whitespace()) =>
            {
                return Err(super::ambiguous(path, "text outside the XML root"));
            }
            Event::CData(_) | Event::GeneralRef(_) if depth == 0 => {
                return Err(super::ambiguous(path, "content outside the XML root"));
            }
            event => writer
                .write_event(event)
                .map_err(|value| error(value.to_string()))?,
        }
    }
    Ok(writer.into_inner())
}
