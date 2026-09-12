use std::collections::{HashMap, HashSet};

use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{QName, ResolveResult};
use quick_xml::{NsReader, Writer, XmlVersion};

use crate::RedactError;
use crate::rels::{external_relationship, unqualified_value};
use crate::scrub::{normalize_part_name, owner_of_rels_path, resolve_relationship_target};

pub(crate) fn detach_scrubbed_fonts(
    parts: &mut [(String, Vec<u8>)],
    scrubbed: &HashSet<String>,
) -> Result<(), RedactError> {
    let indices: HashMap<String, usize> = parts
        .iter()
        .enumerate()
        .map(|(index, (path, _))| (normalize_part_name(path), index))
        .collect();
    for rel_index in 0..parts.len() {
        let path = normalize_part_name(&parts[rel_index].0);
        let Some(owner) = owner_of_rels_path(&path) else {
            continue;
        };
        let Some(&owner_index) = indices.get(&owner) else {
            continue;
        };
        if !owner.ends_with(".xml") {
            continue;
        }
        let mut ids = HashSet::new();
        filter_xml(&parts[rel_index].1, &path, |reader, start| {
            if relationship(reader, start) {
                let attributes = attributes(reader, start, &path)?;
                if !external_relationship(&attributes, true)
                    && matches!(
                        unqualified_value(&attributes, "Type"),
                        Some(
                            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/font"
                                | "http://purl.oclc.org/ooxml/officeDocument/relationships/font"
                        )
                    )
                    && unqualified_value(&attributes, "Target")
                        .and_then(|target| resolve_relationship_target(&path, target))
                        .is_some_and(|target| scrubbed.contains(&target))
                    && let Some(id) = unqualified_value(&attributes, "Id")
                {
                    ids.insert(id.to_owned());
                }
            }
            Ok(false)
        })?;
        if ids.is_empty() {
            continue;
        }
        parts[owner_index].1 = filter_xml(&parts[owner_index].1, &owner, |reader, start| {
            if matches!(
                start.local_name().as_ref(),
                b"embedRegular" | b"embedBold" | b"embedItalic" | b"embedBoldItalic"
            ) && matches!(reader.resolver().resolve_element(start.name()).0,
                    ResolveResult::Bound(ns) if matches!(ns.as_ref(),
                        b"http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                            | b"http://purl.oclc.org/ooxml/wordprocessingml/main"))
            {
                return Ok(relationship_ids(reader, start, &owner)?
                    .iter()
                    .any(|id| ids.contains(id)));
            }
            Ok(false)
        })?;
        filter_xml(&parts[owner_index].1, &owner, |reader, start| {
            for (_, id) in attributes(reader, start, &owner)? {
                ids.remove(&id);
            }
            Ok(false)
        })?;
        parts[rel_index].1 = filter_xml(&parts[rel_index].1, &path, |reader, start| {
            Ok(relationship(reader, start)
                && unqualified_value(&attributes(reader, start, &path)?, "Id")
                    .is_some_and(|id| ids.contains(id)))
        })?;
    }
    Ok(())
}

fn relationship(reader: &NsReader<&[u8]>, start: &BytesStart<'_>) -> bool {
    start.local_name().as_ref() == b"Relationship"
        && matches!(reader.resolver().resolve_element(start.name()).0,
            ResolveResult::Bound(ns) if ns.as_ref()
                == b"http://schemas.openxmlformats.org/package/2006/relationships")
}

fn relationship_ids(
    reader: &NsReader<&[u8]>,
    start: &BytesStart<'_>,
    path: &str,
) -> Result<Vec<String>, RedactError> {
    Ok(attributes(reader, start, path)?
        .into_iter()
        .filter(|(key, _)| {
            QName(key.as_bytes()).local_name().as_ref() == b"id"
                && matches!(reader.resolver().resolve_attribute(QName(key.as_bytes())).0,
                    ResolveResult::Bound(ns) if matches!(ns.as_ref(),
                        b"http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                            | b"http://purl.oclc.org/ooxml/officeDocument/relationships"))
        })
        .map(|(_, value)| value)
        .collect())
}

fn attributes(
    reader: &NsReader<&[u8]>,
    start: &BytesStart<'_>,
    path: &str,
) -> Result<Vec<(String, String)>, RedactError> {
    start
        .attributes()
        .map(|attribute| {
            let attribute = attribute.map_err(|error| xml_error(path, error))?;
            let value = attribute
                .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
                .map_err(|error| xml_error(path, error))?;
            Ok((
                String::from_utf8_lossy(attribute.key.as_ref()).into_owned(),
                value.into_owned(),
            ))
        })
        .collect()
}

fn filter_xml(
    bytes: &[u8],
    path: &str,
    mut drop: impl FnMut(&NsReader<&[u8]>, &BytesStart<'_>) -> Result<bool, RedactError>,
) -> Result<Vec<u8>, RedactError> {
    let mut reader = NsReader::from_reader(bytes);
    let mut writer = Writer::new(Vec::with_capacity(bytes.len()));
    let mut skipped = 0;
    loop {
        let event = reader
            .read_event()
            .map_err(|error| xml_error(path, error))?;
        if matches!(event, Event::DocType(_)) {
            return Err(xml_error(path, "DTD/entity declarations are forbidden"));
        }
        if skipped > 0 {
            match event {
                Event::Start(_) => skipped += 1,
                Event::End(_) => skipped -= 1,
                Event::Eof => return Err(xml_error(path, "unterminated font reference")),
                _ => {}
            }
            continue;
        }
        match &event {
            Event::Start(start) if drop(&reader, start)? => {
                skipped = 1;
                continue;
            }
            Event::Empty(start) if drop(&reader, start)? => continue,
            Event::Eof => return Ok(writer.into_inner()),
            _ => {}
        }
        writer
            .write_event(event)
            .map_err(|error| xml_error(path, error))?;
    }
}

fn xml_error(path: &str, error: impl std::fmt::Display) -> RedactError {
    RedactError::Xml {
        part: path.to_owned(),
        message: error.to_string(),
    }
}
