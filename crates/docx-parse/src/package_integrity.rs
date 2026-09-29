use std::collections::HashSet;

use quick_xml::events::Event;
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::{NsReader, XmlVersion};

use crate::xml::{ParseLimits, namespaces};

pub(crate) fn package_warnings(parts: &[(String, Vec<u8>)], limits: &ParseLimits) -> Vec<String> {
    let mut warnings = Vec::new();
    let content_types = ooxml_opc::parse_content_types(parts);
    if content_types.is_ok()
        || matches!(
            content_types,
            Err(ooxml_opc::DocumentKindError::MissingContentTypes)
        )
    {
        let uncovered = parts
            .iter()
            .filter(|(path, _)| {
                !path.eq_ignore_ascii_case("[Content_Types].xml")
                    && !path.to_ascii_lowercase().ends_with(".rels")
                    && content_types
                        .as_ref()
                        .ok()
                        .and_then(|types| types.content_type_for(path))
                        .is_none()
            })
            .count();
        if uncovered > 0 {
            warnings.push(format!(
                "DOCX contains {uncovered} orphan OPC parts with no declared content type."
            ));
        }
    }

    let mut drawing_ids = HashSet::new();
    let mut bookmark_ids = HashSet::new();
    let mut repeated_drawings = 0;
    let mut repeated_bookmarks = 0;
    let mut xml_bytes = 0;
    let mut xml_events = 0;
    for (path, bytes) in parts {
        let content_type = content_types
            .as_ref()
            .ok()
            .and_then(|types| types.content_type_for(path));
        if !path.to_ascii_lowercase().ends_with(".xml")
            && !content_type
                .is_some_and(|value| value.ends_with("+xml") || value == "application/xml")
        {
            continue;
        }
        if bytes.len() > limits.max_xml_bytes {
            continue;
        }
        let mut reader = NsReader::from_reader(bytes.as_slice());
        let mut depth = 0;
        let mut alternates = Vec::new();
        let mut skipped_depth = None;
        while let Ok((namespace, event)) = reader.read_resolved_event() {
            xml_events += 1;
            if xml_events > limits.max_xml_events {
                break;
            }
            let empty = matches!(event, Event::Empty(_));
            match event {
                Event::Start(element) | Event::Empty(element) => {
                    let local = element.local_name();
                    let word = is_namespace(&namespace, namespaces::W);
                    if depth == 0 {
                        if !word
                            || !matches!(
                                local.as_ref(),
                                b"document"
                                    | b"hdr"
                                    | b"ftr"
                                    | b"footnotes"
                                    | b"endnotes"
                                    | b"comments"
                            )
                        {
                            break;
                        }
                        xml_bytes += bytes.len();
                        if xml_bytes > limits.max_xml_bytes {
                            break;
                        }
                    }
                    if depth >= limits.max_xml_depth {
                        break;
                    }
                    if skipped_depth.is_none() && is_namespace(&namespace, namespaces::MC) {
                        match local.as_ref() {
                            b"AlternateContent" if !empty => alternates.push((depth, false)),
                            b"Choice" | b"Fallback" => {
                                if let Some((parent_depth, chosen)) = alternates.last_mut()
                                    && depth == *parent_depth + 1
                                {
                                    if *chosen {
                                        skipped_depth = Some(depth);
                                    } else {
                                        *chosen = true;
                                    }
                                }
                            }
                            _ => {}
                        }
                    }
                    if skipped_depth.is_none() {
                        let drawing =
                            local.as_ref() == b"docPr" && is_namespace(&namespace, namespaces::WP);
                        let bookmark = word && local.as_ref() == b"bookmarkStart";
                        if drawing || bookmark {
                            for attribute in element
                                .attributes()
                                .take(limits.max_attributes_per_element)
                                .flatten()
                            {
                                let (namespace, name) =
                                    reader.resolver().resolve_attribute(attribute.key);
                                if name.as_ref() != b"id"
                                    || (drawing && namespace != ResolveResult::Unbound)
                                    || (bookmark && !is_namespace(&namespace, namespaces::W))
                                    || attribute.value.len() > limits.max_attribute_bytes
                                {
                                    continue;
                                }
                                let Ok(value) = attribute.decoded_and_normalized_value(
                                    XmlVersion::Implicit1_0,
                                    reader.decoder(),
                                ) else {
                                    continue;
                                };
                                if drawing {
                                    let value = value.trim();
                                    let id = value
                                        .parse::<u32>()
                                        .map(|id| id.to_string())
                                        .unwrap_or_else(|_| value.to_owned());
                                    if !id.is_empty() && !drawing_ids.insert(id) {
                                        repeated_drawings += 1;
                                    }
                                } else {
                                    let value = value.trim();
                                    let id = value
                                        .parse::<i64>()
                                        .map(|id| id.to_string())
                                        .unwrap_or_else(|_| value.to_owned());
                                    if !id.is_empty() && !bookmark_ids.insert(id) {
                                        repeated_bookmarks += 1;
                                    }
                                }
                            }
                        }
                    }
                    if empty {
                        if skipped_depth == Some(depth) {
                            skipped_depth = None;
                        }
                    } else {
                        depth += 1;
                    }
                }
                Event::End(_) => {
                    depth = depth.saturating_sub(1);
                    if skipped_depth == Some(depth) {
                        skipped_depth = None;
                    }
                    if alternates.last().is_some_and(|(start, _)| *start == depth) {
                        alternates.pop();
                    }
                }
                Event::Eof => break,
                _ => {}
            }
        }
        if xml_events > limits.max_xml_events {
            break;
        }
    }
    if repeated_drawings > 0 {
        warnings.push(format!(
            "DOCX contains {repeated_drawings} drawing doc properties with duplicate `wp:docPr id` values."
        ));
    }
    if repeated_bookmarks > 0 {
        warnings.push(format!(
            "DOCX contains {repeated_bookmarks} bookmark starts with duplicate `w:bookmarkStart w:id` values."
        ));
    }
    warnings
}

fn is_namespace(namespace: &ResolveResult<'_>, expected: &str) -> bool {
    if *namespace == ResolveResult::Bound(Namespace(expected.as_bytes())) {
        return true;
    }
    let strict = match expected {
        namespaces::W => "http://purl.oclc.org/ooxml/wordprocessingml/main",
        namespaces::WP => "http://purl.oclc.org/ooxml/drawingml/wordprocessingDrawing",
        _ => return false,
    };
    *namespace == ResolveResult::Bound(Namespace(strict.as_bytes()))
}
