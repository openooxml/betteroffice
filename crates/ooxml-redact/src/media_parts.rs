use std::collections::{BTreeMap, HashSet};

use quick_xml::events::{BytesStart, Event};
use quick_xml::name::ResolveResult;
use quick_xml::{NsReader, Writer, XmlVersion};

use crate::RedactError;
use crate::rels::{external_relationship, unqualified_value};
use crate::scrub::{normalize_part_name, resolve_relationship_target};

const CONTENT_TYPES: &str = "[content_types].xml";
const TYPES_NS: &[u8] = b"http://schemas.openxmlformats.org/package/2006/content-types";
const RELS_NS: &[u8] = b"http://schemas.openxmlformats.org/package/2006/relationships";
const HDPHOTO_REL: &str = "http://schemas.microsoft.com/office/2007/relationships/hdphoto";
const IMAGE_REL: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";

pub(crate) fn convert_wdp_parts(parts: &mut [(String, Vec<u8>)]) -> Result<(), RedactError> {
    let mut occupied: HashSet<_> = parts
        .iter()
        .map(|(path, _)| normalize_part_name(path))
        .collect();
    let mut renamed = BTreeMap::new();
    let mut index = 1;
    for (path, _) in parts.iter() {
        let canonical = normalize_part_name(path);
        if !canonical.ends_with(".wdp") {
            continue;
        }
        let directory = path.rsplit_once('/').map_or("", |(directory, _)| directory);
        let new_path = loop {
            let name = format!("redacted-image-{index}.png");
            let candidate = if directory.is_empty() {
                name
            } else {
                format!("{directory}/{name}")
            };
            index += 1;
            if occupied.insert(normalize_part_name(&candidate)) {
                break candidate;
            }
        };
        if renamed.insert(canonical, new_path).is_some() {
            return Err(RedactError::Container(
                "ambiguous WDP part names".to_owned(),
            ));
        }
    }
    if renamed.is_empty() {
        return Ok(());
    }
    if !parts
        .iter()
        .any(|(path, _)| normalize_part_name(path) == CONTENT_TYPES)
    {
        return Err(RedactError::Container(
            "WDP conversion requires content types".to_owned(),
        ));
    }
    for (path, bytes) in parts.iter_mut() {
        let canonical = normalize_part_name(path);
        if let Some(new_path) = renamed.get(&canonical) {
            *path = new_path.clone();
            bytes.clear();
        } else if canonical == CONTENT_TYPES || canonical.ends_with(".rels") {
            *bytes = rewrite_part(path, bytes, &renamed)?;
        }
    }
    Ok(())
}

fn rewrite_part(
    path: &str,
    bytes: &[u8],
    renamed: &BTreeMap<String, String>,
) -> Result<Vec<u8>, RedactError> {
    let content_types = normalize_part_name(path) == CONTENT_TYPES;
    let mut reader = NsReader::from_reader(bytes);
    let mut writer = Writer::new(Vec::with_capacity(bytes.len()));
    let mut depth = 0usize;
    let mut skipped = 0usize;
    let mut root_prefix = None;
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
                Event::Eof => return Err(xml_error(path, "unterminated content type")),
                _ => {}
            }
            continue;
        }
        let empty = matches!(event, Event::Empty(_));
        match event {
            Event::Start(start) | Event::Empty(start) => {
                let namespace = reader.resolver().resolve_element(start.name()).0;
                let owned = matches!(namespace, ResolveResult::Bound(ns) if ns.as_ref() == if content_types { TYPES_NS } else { RELS_NS });
                let local = start.local_name();
                let attributes = attributes(&reader, &start, path)?;
                if content_types && depth == 0 && local.as_ref() == b"Types" && owned {
                    let name = String::from_utf8_lossy(start.name().as_ref()).into_owned();
                    let prefix = name.rsplit_once(':').map_or("", |(prefix, _)| prefix);
                    root_prefix = Some(if prefix.is_empty() {
                        String::new()
                    } else {
                        format!("{prefix}:")
                    });
                    writer
                        .write_event(Event::Start(start.to_owned()))
                        .map_err(|error| xml_error(path, error))?;
                    if empty {
                        append_overrides(
                            &mut writer,
                            root_prefix.as_deref().unwrap(),
                            renamed,
                            path,
                        )?;
                        writer
                            .write_event(Event::End(start.to_end()))
                            .map_err(|error| xml_error(path, error))?;
                    } else {
                        depth += 1;
                    }
                    continue;
                }
                if content_types
                    && depth == 1
                    && owned
                    && drops_content_type(local.as_ref(), &attributes, renamed)
                {
                    if !empty {
                        skipped = 1;
                    }
                    continue;
                }
                let mut replacement = None;
                if !content_types
                    && owned
                    && local.as_ref() == b"Relationship"
                    && !external_relationship(&attributes, true)
                {
                    replacement = unqualified_value(&attributes, "Target")
                        .and_then(|target| resolve_relationship_target(path, target))
                        .and_then(|target| renamed.get(&target));
                }
                let mut output = start.into_owned();
                if let Some(target) = replacement {
                    output.clear_attributes();
                    for (key, value) in &attributes {
                        let value = if key == "Target" {
                            format!("/{target}")
                        } else if key == "Type" && value == HDPHOTO_REL {
                            IMAGE_REL.to_owned()
                        } else {
                            value.clone()
                        };
                        output.push_attribute((key.as_str(), value.as_str()));
                    }
                }
                writer
                    .write_event(if empty {
                        Event::Empty(output)
                    } else {
                        Event::Start(output)
                    })
                    .map_err(|error| xml_error(path, error))?;
                if !empty {
                    depth += 1;
                }
            }
            Event::End(end) => {
                if content_types && depth == 1 {
                    let prefix = root_prefix
                        .as_deref()
                        .ok_or_else(|| xml_error(path, "invalid content-types root"))?;
                    append_overrides(&mut writer, prefix, renamed, path)?;
                }
                depth = depth.saturating_sub(1);
                writer
                    .write_event(Event::End(end))
                    .map_err(|error| xml_error(path, error))?;
            }
            Event::Eof => {
                if content_types && root_prefix.is_none() {
                    return Err(xml_error(path, "invalid content-types root"));
                }
                return Ok(writer.into_inner());
            }
            other => writer
                .write_event(other)
                .map_err(|error| xml_error(path, error))?,
        }
    }
}

fn drops_content_type(
    local: &[u8],
    attributes: &[(String, String)],
    renamed: &BTreeMap<String, String>,
) -> bool {
    match local {
        b"Default" => unqualified_value(attributes, "Extension")
            .is_some_and(|extension| extension.eq_ignore_ascii_case("wdp")),
        b"Override" => unqualified_value(attributes, "PartName").is_some_and(|name| {
            let name = normalize_part_name(name);
            renamed.contains_key(&name)
                || renamed
                    .values()
                    .any(|value| normalize_part_name(value) == name)
        }),
        _ => false,
    }
}

fn append_overrides(
    writer: &mut Writer<Vec<u8>>,
    prefix: &str,
    renamed: &BTreeMap<String, String>,
    path: &str,
) -> Result<(), RedactError> {
    for name in renamed.values() {
        let mut entry = BytesStart::new(format!("{prefix}Override"));
        let part_name = format!("/{name}");
        entry.push_attribute(("PartName", part_name.as_str()));
        entry.push_attribute(("ContentType", "image/png"));
        writer
            .write_event(Event::Empty(entry))
            .map_err(|error| xml_error(path, error))?;
    }
    Ok(())
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
            Ok((
                String::from_utf8_lossy(attribute.key.as_ref()).into_owned(),
                attribute
                    .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
                    .map_err(|error| xml_error(path, error))?
                    .into_owned(),
            ))
        })
        .collect()
}

fn xml_error(path: &str, error: impl std::fmt::Display) -> RedactError {
    RedactError::Xml {
        part: path.to_owned(),
        message: error.to_string(),
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::{Format, redact_with_report};

    fn package(parts: Vec<(&str, &str)>) -> Vec<u8> {
        ooxml_opc::rezip_parts(
            &parts
                .into_iter()
                .map(|(name, value)| (name.to_owned(), value.as_bytes().to_vec()))
                .collect::<Vec<_>>(),
        )
        .unwrap()
    }

    #[test]
    fn wdp_parts_become_blank_png_with_consistent_targets_and_content_types() {
        for (main, content_type, folder, rel_path) in [
            (
                "word/document.xml",
                "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
                "word/media",
                "word/_rels/document.xml.rels",
            ),
            (
                "ppt/presentation.xml",
                "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml",
                "ppt/media",
                "ppt/_rels/presentation.xml.rels",
            ),
            (
                "xl/workbook.xml",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
                "xl/media",
                "xl/_rels/workbook.xml.rels",
            ),
        ] {
            let image = format!("{folder}/original.wdp");
            let collision = format!("{folder}/REDACTED-IMAGE-1.PNG");
            let types = format!(
                r#"<ct:Types xmlns:ct="http://schemas.openxmlformats.org/package/2006/content-types" xmlns:other="urn:foreign"><ct:Default Extension="wdp" ContentType="image/vnd.ms-photo"/><ct:Default Extension="png" ContentType="image/png"/><ct:Override PartName="/{main}" ContentType="{content_type}"/><ct:Override PartName="/{image}" ContentType="image/vnd.ms-photo"/><other:Default Extension="wdp" ContentType="foreign"/></ct:Types>"#
            );
            let relationships = format!(
                r#"<p:Relationships xmlns:p="http://schemas.openxmlformats.org/package/2006/relationships" xmlns:other="urn:foreign"><p:Relationship Id="relative" Type="{HDPHOTO_REL}" Target="media/original.wdp"/><p:Relationship Id="absolute" Type="{IMAGE_REL}" Target="/{image}"/><p:Relationship Id="normalized" Type="{IMAGE_REL}" Target="media/../media/original.wdp"/><p:Relationship Id="external" Type="{IMAGE_REL}" Target="https://private.invalid/original.wdp" TargetMode="External"/><other:Relationship Id="foreign" Target="media/original.wdp"/><p:Relationship Id="qualified" Type="{IMAGE_REL}" other:Target="media/original.wdp" Target="media/REDACTED-IMAGE-1.PNG"/></p:Relationships>"#
            );
            let document = r#"<root xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><a:xfrm><a:off x="1" y="2"/><a:ext cx="123456" cy="654321"/></a:xfrm><a:blip r:embed="relative"/></root>"#;
            let input = package(vec![
                ("[Content_Types].xml", &types),
                (main, document),
                (rel_path, &relationships),
                (&image, "PRIVATE_WDP_BYTES"),
                (&collision, "PRIVATE_PNG_BYTES"),
            ]);
            let (output, report) = redact_with_report(&input, Format::Auto).unwrap();
            assert_eq!(report.media_parts, 2);
            let parts: BTreeMap<_, _> = ooxml_opc::unzip_parts(&output)
                .unwrap()
                .into_iter()
                .collect();
            let replacement = format!("{folder}/redacted-image-2.png");
            assert!(!parts.contains_key(&image));
            assert!(parts.contains_key(&collision));
            assert_eq!(
                image::guess_format(&parts[&replacement]).unwrap(),
                image::ImageFormat::Png
            );
            let decoded = image::load_from_memory(&parts[&replacement]).unwrap();
            assert_eq!((decoded.width(), decoded.height()), (64, 64));
            assert_eq!(parts[main], document.as_bytes());
            let rels = String::from_utf8_lossy(&parts[rel_path]);
            assert_eq!(
                rels.matches(&format!("Target=\"/{replacement}\"")).count(),
                3
            );
            assert!(!rels.contains(HDPHOTO_REL));
            assert!(rels.contains("Id=\"foreign\" Target=\"media/original.wdp\""));
            assert!(rels.contains(
                "other:Target=\"media/original.wdp\" Target=\"media/REDACTED-IMAGE-1.PNG\""
            ));
            assert!(rels.contains("https://example.com"));
            let types = String::from_utf8_lossy(&parts["[Content_Types].xml"]);
            assert!(types.contains(&format!(
                "PartName=\"/{replacement}\" ContentType=\"image/png\""
            )));
            assert!(!types.contains("image/vnd.ms-photo"));
            assert!(types.contains("<other:Default Extension=\"wdp\" ContentType=\"foreign\""));
            assert!(
                parts
                    .values()
                    .all(|bytes| !String::from_utf8_lossy(bytes).contains("PRIVATE_"))
            );
        }
    }

    #[test]
    fn conversion_is_inert_without_wdp_and_supports_empty_content_type_roots() {
        let mut plain = vec![("word/document.xml".to_owned(), b"not XML".to_vec())];
        let original = plain.clone();
        convert_wdp_parts(&mut plain).unwrap();
        assert_eq!(plain, original);
        let mut parts = vec![
            (
                "[Content_Types].xml".to_owned(),
                br#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>"#
                    .to_vec(),
            ),
            ("image.WDP".to_owned(), b"private bytes".to_vec()),
        ];
        convert_wdp_parts(&mut parts).unwrap();
        assert_eq!(parts[1].0, "redacted-image-1.png");
        assert!(parts[1].1.is_empty());
        let types = String::from_utf8_lossy(&parts[0].1);
        assert!(
            types
                .contains("<Override PartName=\"/redacted-image-1.png\" ContentType=\"image/png\"")
        );
    }
}
