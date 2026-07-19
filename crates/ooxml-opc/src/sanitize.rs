use std::collections::HashSet;

use quick_xml::events::{BytesCData, BytesStart, BytesText, Event};
use quick_xml::{Reader, Writer, XmlVersion};

use crate::{rezip_parts, unzip_parts};

pub fn sanitize_package(data: &[u8]) -> Result<Vec<u8>, String> {
    sanitize_package_inner(data, None)
}

pub fn sanitize_package_for_format(data: &[u8], expected_format: &str) -> Result<Vec<u8>, String> {
    if !matches!(expected_format, "docx" | "xlsx" | "pptx") {
        return Err(format!("unsupported OOXML format: {expected_format}"));
    }
    sanitize_package_inner(data, Some(expected_format))
}

fn sanitize_package_inner(data: &[u8], expected_format: Option<&str>) -> Result<Vec<u8>, String> {
    let mut parts = unzip_parts(data)?;
    let detected = detect_format(&parts)?;
    if let Some(expected) = expected_format
        && detected != expected
    {
        return Err(format!(
            "claimed {expected} content does not match detected {detected} package"
        ));
    }

    let mut removed: HashSet<String> = parts
        .iter()
        .filter(|(path, _)| dangerous_path(path))
        .map(|(path, _)| normalize_part_name(path))
        .collect();
    if let Some((_, content_types)) = parts
        .iter()
        .find(|(path, _)| path.eq_ignore_ascii_case("[Content_Types].xml"))
    {
        removed.extend(dangerous_content_type_parts(content_types)?);
    }

    parts.retain(|(path, _)| !removed.contains(&normalize_part_name(path)));
    for (path, bytes) in &mut parts {
        let lower = path.to_ascii_lowercase();
        if lower == "[content_types].xml" {
            *bytes = sanitize_content_types(bytes, &removed, path)?;
        } else if lower.ends_with(".rels") {
            *bytes = sanitize_relationships(bytes, &removed, path)?;
        } else if is_xml_part(&lower) {
            *bytes = neutralize_fields(bytes, path)?;
        }
    }
    rezip_parts(&parts)
}

fn detect_format(parts: &[(String, Vec<u8>)]) -> Result<&'static str, String> {
    if let Some((_, bytes)) = parts
        .iter()
        .find(|(path, _)| path.eq_ignore_ascii_case("[Content_Types].xml"))
    {
        let content_types = String::from_utf8_lossy(bytes).to_ascii_lowercase();
        if content_types.contains("wordprocessingml.document.main+xml")
            || content_types.contains("ms-word.document.macroenabled.main+xml")
        {
            return Ok("docx");
        }
        if content_types.contains("spreadsheetml.sheet.main+xml")
            || content_types.contains("ms-excel.sheet.macroenabled.main+xml")
        {
            return Ok("xlsx");
        }
        if content_types.contains("presentationml.presentation.main+xml")
            || content_types.contains("ms-powerpoint.presentation.macroenabled.main+xml")
        {
            return Ok("pptx");
        }
    }
    if has_part(parts, "word/document.xml") {
        Ok("docx")
    } else if has_part(parts, "xl/workbook.xml") {
        Ok("xlsx")
    } else if has_part(parts, "ppt/presentation.xml") {
        Ok("pptx")
    } else {
        Err("could not detect DOCX, XLSX, or PPTX package".to_owned())
    }
}

fn has_part(parts: &[(String, Vec<u8>)], expected: &str) -> bool {
    parts
        .iter()
        .any(|(path, _)| path.eq_ignore_ascii_case(expected))
}

fn dangerous_content_type_parts(xml: &[u8]) -> Result<HashSet<String>, String> {
    let mut reader = Reader::from_reader(xml);
    let mut paths = HashSet::new();
    loop {
        match reader
            .read_event()
            .map_err(|error| format!("invalid [Content_Types].xml: {error}"))?
        {
            Event::Start(start) | Event::Empty(start)
                if start.name().local_name().as_ref() == b"Override" =>
            {
                let attributes = attributes(&reader, &start, "[Content_Types].xml")?;
                let part_name = attribute_value(&attributes, "PartName");
                let content_type = attribute_value(&attributes, "ContentType");
                if let (Some(part_name), Some(content_type)) = (part_name, content_type)
                    && dangerous_content_type(content_type)
                {
                    paths.insert(normalize_part_name(part_name));
                }
            }
            Event::DocType(_) => return Err("DTD is forbidden in [Content_Types].xml".to_owned()),
            Event::Eof => return Ok(paths),
            _ => {}
        }
    }
}

fn sanitize_content_types(
    xml: &[u8],
    removed: &HashSet<String>,
    path: &str,
) -> Result<Vec<u8>, String> {
    let mut reader = Reader::from_reader(xml);
    let mut writer = Writer::new(Vec::with_capacity(xml.len()));
    let mut skip_depth = 0_usize;
    loop {
        let event = reader
            .read_event()
            .map_err(|error| format!("invalid XML in {path}: {error}"))?;
        if skip_depth > 0 {
            match event {
                Event::Start(_) => skip_depth += 1,
                Event::End(_) => skip_depth -= 1,
                Event::Eof => return Err(format!("unexpected EOF in {path}")),
                _ => {}
            }
            continue;
        }
        match event {
            Event::Start(start) => {
                let local = start.name().local_name();
                let name = local.as_ref();
                let values = attributes(&reader, &start, path)?;
                if remove_content_type_entry(name, &values, removed) {
                    skip_depth = 1;
                } else {
                    write_start(
                        &mut writer,
                        rewrite_content_type(start, values),
                        false,
                        path,
                    )?;
                }
            }
            Event::Empty(start) => {
                let local = start.name().local_name();
                let name = local.as_ref();
                let values = attributes(&reader, &start, path)?;
                if !remove_content_type_entry(name, &values, removed) {
                    write_start(&mut writer, rewrite_content_type(start, values), true, path)?;
                }
            }
            Event::DocType(_) => return Err(format!("DTD is forbidden in {path}")),
            Event::Comment(_) | Event::PI(_) => {}
            Event::Eof => return Ok(writer.into_inner()),
            other => write(&mut writer, other, path)?,
        }
    }
}

fn remove_content_type_entry(
    element: &[u8],
    attributes: &[(String, String)],
    removed: &HashSet<String>,
) -> bool {
    let content_type = attribute_value(attributes, "ContentType");
    if content_type.is_some_and(dangerous_content_type) {
        return true;
    }
    element == b"Override"
        && attribute_value(attributes, "PartName")
            .is_some_and(|path| removed.contains(&normalize_part_name(path)))
}

fn rewrite_content_type(
    start: BytesStart<'_>,
    attributes: Vec<(String, String)>,
) -> BytesStart<'static> {
    let mut output = start.into_owned();
    output.clear_attributes();
    for (key, value) in attributes {
        let value = if attribute_local(&key) == "ContentType" {
            macro_free_content_type(&value).to_owned()
        } else {
            value
        };
        output.push_attribute((key.as_str(), value.as_str()));
    }
    output
}

fn macro_free_content_type(content_type: &str) -> &str {
    match content_type.to_ascii_lowercase().as_str() {
        "application/vnd.ms-word.document.macroenabled.main+xml" => {
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"
        }
        "application/vnd.ms-excel.sheet.macroenabled.main+xml" => {
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"
        }
        "application/vnd.ms-powerpoint.presentation.macroenabled.main+xml" => {
            "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"
        }
        _ => content_type,
    }
}

fn sanitize_relationships(
    xml: &[u8],
    removed: &HashSet<String>,
    path: &str,
) -> Result<Vec<u8>, String> {
    let mut reader = Reader::from_reader(xml);
    let mut writer = Writer::new(Vec::with_capacity(xml.len()));
    let mut skip_depth = 0_usize;
    loop {
        let event = reader
            .read_event()
            .map_err(|error| format!("invalid XML in {path}: {error}"))?;
        if skip_depth > 0 {
            match event {
                Event::Start(_) => skip_depth += 1,
                Event::End(_) => skip_depth -= 1,
                Event::Eof => return Err(format!("unexpected EOF in {path}")),
                _ => {}
            }
            continue;
        }
        match event {
            Event::Start(start) => {
                let remove = start.name().local_name().as_ref() == b"Relationship"
                    && remove_relationship(&attributes(&reader, &start, path)?, removed, path);
                if remove {
                    skip_depth = 1;
                } else {
                    write_start(&mut writer, start.into_owned(), false, path)?;
                }
            }
            Event::Empty(start) => {
                let remove = start.name().local_name().as_ref() == b"Relationship"
                    && remove_relationship(&attributes(&reader, &start, path)?, removed, path);
                if !remove {
                    write_start(&mut writer, start.into_owned(), true, path)?;
                }
            }
            Event::DocType(_) => return Err(format!("DTD is forbidden in {path}")),
            Event::Comment(_) | Event::PI(_) => {}
            Event::Eof => return Ok(writer.into_inner()),
            other => write(&mut writer, other, path)?,
        }
    }
}

fn remove_relationship(
    attributes: &[(String, String)],
    removed: &HashSet<String>,
    relationship_path: &str,
) -> bool {
    let target = attribute_value(attributes, "Target").unwrap_or_default();
    let target_mode = attribute_value(attributes, "TargetMode").unwrap_or_default();
    let relationship_type = attribute_value(attributes, "Type").unwrap_or_default();
    target_mode.eq_ignore_ascii_case("External")
        || external_target(target)
        || dangerous_relationship_type(relationship_type)
        || dangerous_path(target)
        || resolve_relationship_target(relationship_path, target)
            .is_some_and(|target| removed.contains(&target))
}

fn neutralize_fields(xml: &[u8], path: &str) -> Result<Vec<u8>, String> {
    let mut reader = Reader::from_reader(xml);
    reader.config_mut().trim_text(false);
    let mut writer = Writer::new(Vec::with_capacity(xml.len()));
    let mut stack = Vec::new();
    let mut skip_depth = 0_usize;
    loop {
        let event = reader
            .read_event()
            .map_err(|error| format!("invalid XML in {path}: {error}"))?;
        if skip_depth > 0 {
            match event {
                Event::Start(_) => skip_depth += 1,
                Event::End(_) => skip_depth -= 1,
                Event::Eof => return Err(format!("unexpected EOF in {path}")),
                _ => {}
            }
            continue;
        }
        match event {
            Event::Start(start) => {
                let local =
                    String::from_utf8_lossy(start.name().local_name().as_ref()).into_owned();
                if dangerous_element(&local) {
                    skip_depth = 1;
                    continue;
                }
                let output = if local == "fldSimple" {
                    neutralize_instruction_attribute(&reader, start, path)?
                } else {
                    start.into_owned()
                };
                stack.push(local);
                write_start(&mut writer, output, false, path)?;
            }
            Event::Empty(start) => {
                let local = start.name().local_name();
                if dangerous_element(&String::from_utf8_lossy(local.as_ref())) {
                    continue;
                }
                let output = if local.as_ref() == b"fldSimple" {
                    neutralize_instruction_attribute(&reader, start, path)?
                } else {
                    start.into_owned()
                };
                write_start(&mut writer, output, true, path)?;
            }
            Event::End(end) => {
                stack.pop();
                write(&mut writer, Event::End(end), path)?;
            }
            Event::Text(_) if stack.last().is_some_and(|name| field_element(name)) => {
                write(&mut writer, Event::Text(BytesText::new("0")), path)?;
            }
            Event::CData(_) if stack.last().is_some_and(|name| field_element(name)) => {
                write(&mut writer, Event::CData(BytesCData::new("0")), path)?;
            }
            Event::GeneralRef(_) if stack.last().is_some_and(|name| field_element(name)) => {
                write(&mut writer, Event::Text(BytesText::new("0")), path)?;
            }
            Event::DocType(_) => return Err(format!("DTD is forbidden in {path}")),
            Event::Comment(_) | Event::PI(_) => {}
            Event::Eof => return Ok(writer.into_inner()),
            other => write(&mut writer, other, path)?,
        }
    }
}

fn neutralize_instruction_attribute(
    reader: &Reader<&[u8]>,
    start: BytesStart<'_>,
    path: &str,
) -> Result<BytesStart<'static>, String> {
    let mut output = start.into_owned();
    let values = attributes(reader, &output, path)?;
    output.clear_attributes();
    for (key, value) in values {
        let value = if attribute_local(&key) == "instr" {
            "0"
        } else {
            &value
        };
        output.push_attribute((key.as_str(), value));
    }
    Ok(output)
}

fn field_element(name: &str) -> bool {
    matches!(
        name,
        "instrText" | "delInstrText" | "f" | "formula" | "formula1" | "formula2"
    )
}

fn dangerous_element(name: &str) -> bool {
    matches!(
        name,
        "ddeLink" | "object" | "oleLink" | "oleObject" | "OLEObject" | "oleObj" | "control"
    )
}

fn attributes(
    reader: &Reader<&[u8]>,
    start: &BytesStart<'_>,
    path: &str,
) -> Result<Vec<(String, String)>, String> {
    start
        .attributes()
        .map(|attribute| {
            let attribute = attribute.map_err(|error| format!("invalid XML in {path}: {error}"))?;
            let key = String::from_utf8_lossy(attribute.key.as_ref()).into_owned();
            let value = attribute
                .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
                .map_err(|error| format!("invalid XML in {path}: {error}"))?
                .into_owned();
            Ok((key, value))
        })
        .collect()
}

fn attribute_value<'a>(attributes: &'a [(String, String)], expected: &str) -> Option<&'a str> {
    attributes
        .iter()
        .find(|(name, _)| attribute_local(name).eq_ignore_ascii_case(expected))
        .map(|(_, value)| value.as_str())
}

fn attribute_local(name: &str) -> &str {
    name.rsplit_once(':').map_or(name, |(_, local)| local)
}

fn dangerous_path(path: &str) -> bool {
    let path = path.replace('\\', "/").to_ascii_lowercase();
    path.ends_with("vbaproject.bin")
        || path.ends_with("vbadata.xml")
        || path.contains("/macrosheets/")
        || path.contains("/embeddings/")
        || path.contains("/activex/")
        || path.contains("/oleobject")
}

fn dangerous_content_type(content_type: &str) -> bool {
    let content_type = content_type.to_ascii_lowercase();
    content_type.contains("vbaproject")
        || content_type.contains("vbadata")
        || content_type.contains("macrosheet")
        || content_type.contains("oleobject")
        || content_type.contains("activex")
        || content_type.contains("ms-package")
}

fn dangerous_relationship_type(relationship_type: &str) -> bool {
    let relationship_type = relationship_type.to_ascii_lowercase();
    relationship_type.contains("vbaproject")
        || relationship_type.contains("macrosheet")
        || relationship_type.contains("oleobject")
        || relationship_type.ends_with("/package")
        || relationship_type.contains("activex")
        || relationship_type.ends_with("/control")
}

fn external_target(target: &str) -> bool {
    let lower = target.trim().to_ascii_lowercase();
    lower.starts_with("//")
        || lower.contains("://")
        || matches!(
            lower.split_once(':').map(|(scheme, _)| scheme),
            Some("file" | "mailto" | "ftp" | "javascript" | "data")
        )
}

fn resolve_relationship_target(relationship_path: &str, target: &str) -> Option<String> {
    if target.is_empty() || external_target(target) {
        return None;
    }
    let clean_target = target.split(['?', '#']).next().unwrap_or_default();
    let relationship_path = relationship_path.replace('\\', "/");
    let mut segments: Vec<String> =
        if clean_target.starts_with('/') || relationship_path.eq_ignore_ascii_case("_rels/.rels") {
            Vec::new()
        } else {
            relationship_path
                .split("/_rels/")
                .next()
                .unwrap_or_default()
                .split('/')
                .filter(|segment| !segment.is_empty())
                .map(str::to_owned)
                .collect()
        };
    let clean_target = clean_target.trim_start_matches('/').replace('\\', "/");
    for segment in clean_target
        .split('/')
        .filter(|segment| !segment.is_empty() && *segment != ".")
    {
        if segment == ".." {
            segments.pop()?;
        } else {
            segments.push(segment.to_owned());
        }
    }
    Some(segments.join("/").to_ascii_lowercase())
}

fn normalize_part_name(path: &str) -> String {
    path.replace('\\', "/")
        .trim_start_matches('/')
        .to_ascii_lowercase()
}

fn is_xml_part(path: &str) -> bool {
    path.ends_with(".xml") || path.ends_with(".vml")
}

fn write_start(
    writer: &mut Writer<Vec<u8>>,
    start: BytesStart<'_>,
    empty: bool,
    path: &str,
) -> Result<(), String> {
    if empty {
        write(writer, Event::Empty(start), path)
    } else {
        write(writer, Event::Start(start), path)
    }
}

fn write(writer: &mut Writer<Vec<u8>>, event: Event<'_>, path: &str) -> Result<(), String> {
    writer
        .write_event(event)
        .map_err(|error| format!("writing sanitized XML for {path}: {error}"))
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn strips_external_macro_and_embedded_attack_vectors() {
        let package = rezip_parts(&[
            (
                "[Content_Types].xml".to_owned(),
                br#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="bin" ContentType="application/vnd.ms-office.vbaProject"/><Override PartName="/word/document.xml" ContentType="application/vnd.ms-word.document.macroEnabled.main+xml"/><Override PartName="/word/vbaProject.bin" ContentType="application/vnd.ms-office.vbaProject"/><Override PartName="/word/embeddings/object1.bin" ContentType="application/vnd.openxmlformats-officedocument.oleObject"/></Types>"#.to_vec(),
            ),
            (
                "word/document.xml".to_owned(),
                br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:o="urn:schemas-microsoft-com:office:office"><w:body><w:p><w:fldSimple w:instr="DDEAUTO secret"><w:r><w:instrText>HYPERLINK secret.example</w:instrText></w:r></w:fldSimple><w:object><o:OLEObject ProgID="Package"/></w:object></w:p></w:body></w:document>"#.to_vec(),
            ),
            (
                "word/_rels/document.xml.rels".to_owned(),
                br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://secret.example" TargetMode="External"/><Relationship Id="rId2" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProject" Target="vbaProject.bin"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject" Target="embeddings/object1.bin"/></Relationships>"#.to_vec(),
            ),
            ("word/vbaProject.bin".to_owned(), b"macro secret".to_vec()),
            (
                "word/embeddings/object1.bin".to_owned(),
                b"embedded secret".to_vec(),
            ),
        ])
        .unwrap();

        let sanitized = sanitize_package_for_format(&package, "docx").unwrap();
        let parts = unzip_parts(&sanitized).unwrap();
        assert_eq!(
            parts
                .iter()
                .map(|(path, _)| path.as_str())
                .collect::<Vec<_>>(),
            vec![
                "[Content_Types].xml",
                "word/document.xml",
                "word/_rels/document.xml.rels"
            ]
        );
        let all = parts
            .iter()
            .map(|(_, bytes)| String::from_utf8_lossy(bytes))
            .collect::<String>();
        assert!(!all.contains("secret"));
        assert!(!all.contains("TargetMode"));
        assert!(!all.contains("vbaProject"));
        assert!(!all.contains("oleObject"));
        assert!(!all.contains("ProgID"));
        assert!(all.contains("wordprocessingml.document.main+xml"));
        assert!(all.contains("w:instr=\"0\""));
        assert!(all.contains("<w:instrText>0</w:instrText>"));
    }

    #[test]
    fn validates_claimed_format() {
        let package = rezip_parts(&[
            (
                "[Content_Types].xml".to_owned(),
                br#"<Types><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/></Types>"#.to_vec(),
            ),
            ("xl/workbook.xml".to_owned(), b"<workbook/>".to_vec()),
        ])
        .unwrap();
        assert!(sanitize_package_for_format(&package, "docx").is_err());
    }

    #[test]
    fn accepts_and_rezips_all_three_formats() {
        let cases = [
            (
                "docx",
                "word/document.xml",
                "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
                "<w:document/>",
            ),
            (
                "xlsx",
                "xl/workbook.xml",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
                "<workbook><f>DDE secret</f><ddeLink ddeService=\"DDE secret\"/></workbook>",
            ),
            (
                "pptx",
                "ppt/presentation.xml",
                "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml",
                "<p:presentation/>",
            ),
        ];
        for (format, part_name, content_type, main_xml) in cases {
            let content_types = format!(
                r#"<Types><Override PartName="/{part_name}" ContentType="{content_type}"/></Types>"#
            );
            let package = rezip_parts(&[
                ("[Content_Types].xml".to_owned(), content_types.into_bytes()),
                (part_name.to_owned(), main_xml.as_bytes().to_vec()),
            ])
            .unwrap();
            let sanitized = sanitize_package_for_format(&package, format).unwrap();
            let parts = unzip_parts(&sanitized).unwrap();
            assert_eq!(parts.len(), 2);
            assert!(
                parts
                    .iter()
                    .all(|(_, bytes)| !String::from_utf8_lossy(bytes).contains("DDE secret"))
            );
        }
    }

    #[test]
    fn rejects_dtd_in_xml_parts() {
        let package = rezip_parts(&[
            (
                "[Content_Types].xml".to_owned(),
                br#"<Types><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>"#.to_vec(),
            ),
            (
                "word/document.xml".to_owned(),
                b"<!DOCTYPE x><w:document/>".to_vec(),
            ),
        ])
        .unwrap();
        assert!(sanitize_package(&package).unwrap_err().contains("DTD"));
    }
}
