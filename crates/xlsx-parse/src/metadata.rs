//! Adding a dynamic-array record to a preserved cell metadata part.

use std::ops::Range;

use quick_xml::Reader;
use quick_xml::events::{BytesStart, Event};

use crate::package::{XmlAttribute, attributes, set_attribute};
use crate::xml::xml_err;
use crate::{MAX_DEPTH, ParseError};

const DYNAMIC_TYPE: &str = r#"metadataType name="XLDAPR" minSupportedVersion="120000" copy="1" pasteAll="1" pasteValues="1" merge="1" splitFirst="1" rowColShift="1" clearFormats="1" clearComments="1" assign="1" coerce="1" cellMeta="1"/>"#;
const DYNAMIC_PROPERTIES: &str = r#"<xda:dynamicArrayProperties xmlns:xda="http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray" fDynamic="1" fCollapsed="0"/>"#;

/// One child of the part's root, with what an added record needs from it.
struct Child {
    local: String,
    name: String,
    attributes: Vec<XmlAttribute>,
    tag: Range<usize>,
    /// where its end tag begins, or its tag ends when it closes itself.
    end: usize,
    /// where its markup ends.
    close: usize,
    empty: bool,
    entries: usize,
    dynamic_type: Option<usize>,
}

/// `source` with a cell metadata record marking a dynamic array, reusing a
/// dynamic-array type or future-metadata block it already declares. Returns
/// the part and the `cm` index of the added record.
pub(crate) fn with_dynamic_array_record(source: &[u8]) -> Result<(Vec<u8>, u32), ParseError> {
    let (root_end, prefix, children) = scan(source)?;
    let named = |local: &str| children.iter().find(|child| child.local == local);
    let mut edits: Vec<(Range<usize>, String)> = Vec::new();
    let mut insert_after_types = String::new();

    let type_index = match named("metadataTypes") {
        Some(types) => match types.dynamic_type {
            Some(index) => index,
            None => {
                let element = format!("<{prefix}{DYNAMIC_TYPE}");
                edits.extend(append(types, &element)?);
                types.entries + 1
            }
        },
        None => {
            insert_after_types.push_str(&format!(
                r#"<{prefix}metadataTypes count="1"><{prefix}{DYNAMIC_TYPE}</{prefix}metadataTypes>"#
            ));
            1
        }
    };
    let block = format!(
        r#"<{prefix}bk><{prefix}extLst><{prefix}ext uri="{{bdbb8cdc-fa1e-496e-a857-3c3f30c029c3}}">{DYNAMIC_PROPERTIES}</{prefix}ext></{prefix}extLst></{prefix}bk>"#
    );
    let future = children.iter().find(|child| {
        child.local == "futureMetadata"
            && child
                .attributes
                .iter()
                .any(|attribute| attribute.local_name() == "name" && attribute.value == "XLDAPR")
    });
    let value_index = match future {
        Some(future) => {
            edits.extend(append(future, &block)?);
            future.entries
        }
        None => {
            insert_after_types.push_str(&format!(
                r#"<{prefix}futureMetadata name="XLDAPR" count="1">{block}</{prefix}futureMetadata>"#
            ));
            0
        }
    };
    let record =
        format!(r#"<{prefix}bk><{prefix}rc t="{type_index}" v="{value_index}"/></{prefix}bk>"#);
    let index = match named("cellMetadata") {
        Some(cells) => {
            edits.extend(append(cells, &record)?);
            cells.entries + 1
        }
        None => {
            insert_after_types.push_str(&format!(
                r#"<{prefix}cellMetadata count="1">{record}</{prefix}cellMetadata>"#
            ));
            1
        }
    };
    if !insert_after_types.is_empty() {
        let at = children
            .iter()
            .filter(|child| {
                matches!(
                    child.local.as_str(),
                    "metadataTypes" | "metadataStrings" | "mdxMetadata" | "futureMetadata"
                )
            })
            .map(|child| child.close)
            .max()
            .unwrap_or(root_end);
        edits.push((at..at, insert_after_types));
    }
    edits.sort_by_key(|(range, _)| std::cmp::Reverse(range.start));
    let mut out = source.to_vec();
    for (range, text) in edits {
        out.splice(range, text.into_bytes());
    }
    let index = u32::try_from(index)
        .map_err(|_| ParseError::Malformed("cell metadata holds too many records".to_owned()))?;
    Ok((out, index))
}

/// The edits that add `element` as the last entry of `child`, counting it.
fn append(child: &Child, element: &str) -> Result<Vec<(Range<usize>, String)>, ParseError> {
    let mut attributes = child.attributes.clone();
    if attributes
        .iter()
        .any(|attribute| attribute.local_name() == "count")
    {
        set_attribute(
            &mut attributes,
            "count",
            "count",
            (child.entries + 1).to_string(),
        );
    }
    let mut start = BytesStart::new(child.name.as_str());
    for attribute in &attributes {
        start.push_attribute((attribute.name.as_str(), attribute.value.as_str()));
    }
    let open = format!("<{}>", String::from_utf8_lossy(&start));
    Ok(if child.empty {
        vec![(
            child.tag.clone(),
            format!("{open}{element}</{}>", child.name),
        )]
    } else {
        vec![
            (child.tag.clone(), open),
            (child.end..child.end, element.to_owned()),
        ]
    })
}

/// The root's start-tag end, its namespace prefix, and its children.
fn scan(source: &[u8]) -> Result<(usize, String, Vec<Child>), ParseError> {
    let mut reader = Reader::from_reader(source);
    reader.config_mut().expand_empty_elements = false;
    reader.config_mut().check_end_names = true;
    let mut depth = 0_usize;
    let mut root_end = None;
    let mut prefix = String::new();
    let mut children: Vec<Child> = Vec::new();
    let mut open: Option<Child> = None;
    loop {
        let before = reader.buffer_position() as usize;
        let event = reader.read_event().map_err(xml_err)?;
        let after = reader.buffer_position() as usize;
        let (element, empty) = match event {
            Event::Start(element) => (element, false),
            Event::Empty(element) => (element, true),
            Event::End(_) => {
                depth = depth.saturating_sub(1);
                if depth == 1
                    && let Some(mut child) = open.take()
                {
                    child.end = before;
                    child.close = after;
                    children.push(child);
                }
                continue;
            }
            Event::Eof => break,
            _ => continue,
        };
        match depth {
            0 => {
                root_end = Some(after);
                let name = String::from_utf8_lossy(element.name().as_ref()).into_owned();
                if let Some((head, _)) = name.split_once(':') {
                    prefix = format!("{head}:");
                }
            }
            1 => {
                let child = Child {
                    local: String::from_utf8_lossy(element.local_name().as_ref()).into_owned(),
                    name: String::from_utf8_lossy(element.name().as_ref()).into_owned(),
                    attributes: attributes(&element)?,
                    tag: before..after,
                    end: after,
                    close: after,
                    empty,
                    entries: 0,
                    dynamic_type: None,
                };
                if empty {
                    children.push(child);
                } else {
                    open = Some(child);
                }
            }
            2 => {
                if let Some(child) = open.as_mut() {
                    let local = element.local_name();
                    if matches!(local.as_ref(), b"metadataType" | b"bk") {
                        child.entries += 1;
                    }
                    let dynamic = local.as_ref() == b"metadataType"
                        && attributes(&element)?.iter().any(|attribute| {
                            attribute.local_name() == "name" && attribute.value == "XLDAPR"
                        });
                    if dynamic && child.dynamic_type.is_none() {
                        child.dynamic_type = Some(child.entries);
                    }
                }
            }
            _ => {}
        }
        if !empty {
            depth += 1;
            if depth > MAX_DEPTH {
                return Err(ParseError::DepthExceeded);
            }
        }
    }
    let root_end = root_end
        .ok_or_else(|| ParseError::Malformed("cell metadata part has no root".to_owned()))?;
    Ok((root_end, prefix, children))
}

#[cfg(test)]
mod tests {
    use quick_xml::NsReader;
    use quick_xml::events::Event;
    use quick_xml::name::ResolveResult;

    use super::with_dynamic_array_record;

    const MAIN: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    const DYNAMIC: &str = "http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray";

    /// every element of `part` with the namespace it resolves to.
    fn resolved(part: &[u8]) -> Vec<(String, String)> {
        let mut reader = NsReader::from_reader(part);
        let mut out = Vec::new();
        loop {
            match reader.read_resolved_event().unwrap() {
                (namespace, Event::Start(element) | Event::Empty(element)) => {
                    let namespace = match namespace {
                        ResolveResult::Bound(namespace) => {
                            String::from_utf8(namespace.as_ref().to_vec()).unwrap()
                        }
                        _ => String::new(),
                    };
                    let local = String::from_utf8(element.local_name().as_ref().to_vec()).unwrap();
                    out.push((local, namespace));
                }
                (_, Event::Eof) => break,
                _ => {}
            }
        }
        out
    }

    fn assert_namespaces(part: &[u8]) {
        for (local, namespace) in resolved(part) {
            let expected = if local == "dynamicArrayProperties" {
                DYNAMIC
            } else {
                MAIN
            };
            assert_eq!(namespace, expected, "{local}");
        }
    }

    #[test]
    fn a_rich_value_part_gains_a_dynamic_array_record() {
        let source = br#"<metadata xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><metadataTypes count="1"><metadataType name="XLRICHVALUE" minSupportedVersion="120000"/></metadataTypes><futureMetadata name="XLRICHVALUE" count="1"><bk><extLst/></bk></futureMetadata><valueMetadata count="1"><bk><rc t="1" v="0"/></bk></valueMetadata></metadata>"#;
        let (part, cm) = with_dynamic_array_record(source).unwrap();
        assert_eq!(cm, 1);
        assert_namespaces(&part);
        let part = String::from_utf8(part).unwrap();
        assert!(
            part.contains(r#"<metadataTypes count="2"><metadataType name="XLRICHVALUE""#),
            "{part}"
        );
        assert!(part.contains(r#"<metadataType name="XLDAPR""#), "{part}");
        assert!(
            part.contains(r#"<futureMetadata name="XLDAPR" count="1"><bk>"#),
            "{part}"
        );
        assert!(
            part.contains(
                r#"<cellMetadata count="1"><bk><rc t="2" v="0"/></bk></cellMetadata><valueMetadata"#
            ),
            "{part}"
        );
        assert!(
            part.contains(r#"<valueMetadata count="1"><bk><rc t="1" v="0"/></bk></valueMetadata>"#),
            "{part}"
        );
    }

    #[test]
    fn existing_records_keep_their_indices() {
        let source = br#"<x:metadata xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><x:metadataTypes count="1"><x:metadataType name="XLDAPR"/></x:metadataTypes><x:futureMetadata name="XLDAPR" count="1"><x:bk/></x:futureMetadata><x:cellMetadata count="2"><x:bk/><x:bk/></x:cellMetadata></x:metadata>"#;
        let (part, cm) = with_dynamic_array_record(source).unwrap();
        assert_eq!(cm, 3);
        assert_namespaces(&part);
        let part = String::from_utf8(part).unwrap();
        assert!(
            part.contains(
                r#"<x:futureMetadata name="XLDAPR" count="2"><x:bk/><x:bk><x:extLst><x:ext "#
            ),
            "{part}"
        );
        assert!(
            part.contains(
                r#"<x:cellMetadata count="3"><x:bk/><x:bk/><x:bk><x:rc t="1" v="1"/></x:bk></x:cellMetadata>"#
            ),
            "{part}"
        );
    }

    /// a prefixed part without a dynamic-array type or block gains both, every
    /// element in the namespace its prefix names.
    #[test]
    fn a_prefixed_part_gains_a_qualified_record() {
        let source = br#"<x:metadata xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><x:metadataTypes count="1"><x:metadataType name="XLRICHVALUE"/></x:metadataTypes></x:metadata>"#;
        let (part, cm) = with_dynamic_array_record(source).unwrap();
        assert_eq!(cm, 1);
        assert_namespaces(&part);
    }
}
