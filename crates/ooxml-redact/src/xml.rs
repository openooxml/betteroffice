use quick_xml::events::{BytesCData, BytesStart, BytesText, Event};
use quick_xml::name::{QName, ResolveResult};
use quick_xml::{NsReader, Writer, XmlVersion};

use crate::mask::{TextMasker, placeholder};
use crate::rels::{self, attribute_local, is_unqualified};
use crate::schema;
use crate::styles::StyleMap;
use crate::{Format, RedactError, RedactionReport};

#[cfg(test)]
pub(crate) fn redact_xml(
    format: Format,
    path: &str,
    bytes: &[u8],
    report: &mut RedactionReport,
) -> Result<Vec<u8>, RedactError> {
    redact_xml_with_styles(
        format,
        path,
        bytes,
        report,
        &StyleMap::default(),
        &mut TextMasker::new(&Default::default()),
    )
}

pub(crate) fn redact_xml_with_styles(
    format: Format,
    path: &str,
    bytes: &[u8],
    report: &mut RedactionReport,
    styles: &StyleMap,
    masker: &mut TextMasker,
) -> Result<Vec<u8>, RedactError> {
    let mut reader = NsReader::from_reader(bytes);
    reader.config_mut().trim_text(false);
    let mut writer = Writer::new(Vec::with_capacity(bytes.len()));
    let mut stack = Vec::new();
    let mut state = RewriteState {
        format,
        path,
        report,
        styles,
        current_style: None,
        cell_type: None,
        custom_property: 0,
        masker,
    };
    let mut skipped_depth = 0;
    let mut preserve_skipped_end = false;

    loop {
        let event = reader
            .read_event()
            .map_err(|error| xml_error(path, error))?;
        if matches!(event, Event::DocType(_)) {
            return Err(xml_error(path, "DTD/entity declarations are forbidden"));
        }
        if skipped_depth > 0 {
            match event {
                Event::Start(_) => skipped_depth += 1,
                Event::End(end) => {
                    skipped_depth -= 1;
                    if skipped_depth == 0 && preserve_skipped_end {
                        writer
                            .write_event(Event::End(end))
                            .map_err(|error| xml_error(path, error))?;
                        preserve_skipped_end = false;
                    }
                }
                Event::Text(text) => {
                    let decoded = text.decode().map_err(|error| xml_error(path, error))?;
                    charge_text(state.report, &decoded);
                }
                Event::CData(text) => {
                    let decoded = text.decode().map_err(|error| xml_error(path, error))?;
                    charge_text(state.report, &decoded);
                }
                Event::Comment(_) | Event::PI(_) => state.report.xml_comments += 1,
                Event::GeneralRef(_) if preserve_skipped_end => {
                    state.report.text_nodes += 1;
                    state.report.characters += 1;
                }
                Event::Eof => return Err(xml_error(path, "unterminated redacted element")),
                _ => {}
            }
            continue;
        }
        if let Event::Start(start) | Event::Empty(start) = &event
            && let Some(value) = fixed_metadata_value(path, &reader, start.name())
        {
            let local = local_name(start.local_name().as_ref());
            let rewritten = rewrite_start(&reader, start.to_owned(), &local, &mut state)?;
            writer
                .write_event(Event::Start(rewritten.borrow()))
                .map_err(|error| xml_error(path, error))?;
            writer
                .write_event(Event::Text(BytesText::new(value)))
                .map_err(|error| xml_error(path, error))?;
            if matches!(event, Event::Empty(_)) {
                writer
                    .write_event(Event::End(rewritten.to_end()))
                    .map_err(|error| xml_error(path, error))?;
            } else {
                skipped_depth = 1;
                preserve_skipped_end = true;
            }
            continue;
        }
        match event {
            Event::Start(start) => {
                let local = local_name(start.name().local_name().as_ref());
                if schema_node(path, &reader, start.name()) && schema::drop_element(&local) {
                    skipped_depth = 1;
                    continue;
                }
                let rewritten = rewrite_start(&reader, start, &local, &mut state)?;
                stack.push(local);
                writer
                    .write_event(Event::Start(rewritten))
                    .map_err(|error| xml_error(path, error))?;
            }
            Event::Empty(start) => {
                let local = local_name(start.name().local_name().as_ref());
                if schema_node(path, &reader, start.name()) && schema::drop_element(&local) {
                    continue;
                }
                let empty_style = local == "style" && word_node(&reader, start.name());
                let rewritten = rewrite_start(&reader, start, &local, &mut state)?;
                if empty_style {
                    state.current_style = None;
                }
                writer
                    .write_event(Event::Empty(rewritten))
                    .map_err(|error| xml_error(path, error))?;
            }
            Event::End(end) => {
                if stack.last().is_some_and(|name| name == "c") {
                    state.cell_type = None;
                }
                if stack.last().is_some_and(|name| name == "style")
                    && word_node(&reader, end.name())
                {
                    state.current_style = None;
                }
                stack.pop();
                writer
                    .write_event(Event::End(end))
                    .map_err(|error| xml_error(path, error))?;
            }
            Event::Text(text) => {
                if let Some(kind) =
                    replacement_kind(format, path, &stack, state.cell_type.as_deref())
                {
                    let decoded = text.decode().map_err(|error| xml_error(path, error))?;
                    let unescaped = quick_xml::escape::unescape(&decoded)
                        .map_err(|error| xml_error(path, error))?;
                    let replacement = replace_text(&unescaped, kind, state.masker)?;
                    charge_text(state.report, &unescaped);
                    writer
                        .write_event(Event::Text(BytesText::new(&replacement)))
                        .map_err(|error| xml_error(path, error))?;
                } else {
                    writer
                        .write_event(Event::Text(text))
                        .map_err(|error| xml_error(path, error))?;
                }
            }
            Event::CData(text) => {
                if let Some(kind) =
                    replacement_kind(format, path, &stack, state.cell_type.as_deref())
                {
                    let decoded = text.decode().map_err(|error| xml_error(path, error))?;
                    let replacement = replace_text(&decoded, kind, state.masker)?;
                    charge_text(state.report, &decoded);
                    writer
                        .write_event(Event::CData(BytesCData::new(&replacement)))
                        .map_err(|error| xml_error(path, error))?;
                } else {
                    writer
                        .write_event(Event::CData(text))
                        .map_err(|error| xml_error(path, error))?;
                }
            }
            Event::GeneralRef(reference) => {
                if let Some(kind) =
                    replacement_kind(format, path, &stack, state.cell_type.as_deref())
                {
                    let replacement = if state.masker.is_random()
                        && matches!(kind, Replacement::Text)
                    {
                        let character = reference
                            .resolve_char_ref()
                            .map_err(|error| xml_error(path, error))?;
                        let decoded = reference.decode().map_err(|error| xml_error(path, error))?;
                        let text = character
                            .map(|character| character.to_string())
                            .unwrap_or_else(|| {
                                match decoded.as_ref() {
                                    "amp" => "&",
                                    "lt" => "<",
                                    "gt" => ">",
                                    "quot" => "\"",
                                    "apos" => "'",
                                    _ => "x",
                                }
                                .to_owned()
                            });
                        charge_text(state.report, &text);
                        state.masker.replace(&text)?
                    } else {
                        state.report.text_nodes += 1;
                        state.report.characters += 1;
                        "x".to_owned()
                    };
                    writer
                        .write_event(Event::Text(BytesText::new(&replacement)))
                        .map_err(|error| xml_error(path, error))?;
                } else {
                    writer
                        .write_event(Event::GeneralRef(reference))
                        .map_err(|error| xml_error(path, error))?;
                }
            }
            Event::Comment(_) | Event::PI(_) => {
                state.report.xml_comments += 1;
            }
            Event::Eof => break,
            other => writer
                .write_event(other)
                .map_err(|error| xml_error(path, error))?,
        }
    }

    Ok(writer.into_inner())
}

struct RewriteState<'a> {
    format: Format,
    path: &'a str,
    report: &'a mut RedactionReport,
    styles: &'a StyleMap,
    current_style: Option<String>,
    cell_type: Option<String>,
    custom_property: usize,
    masker: &'a mut TextMasker,
}

fn rewrite_start(
    reader: &NsReader<&[u8]>,
    start: BytesStart<'_>,
    element: &str,
    state: &mut RewriteState<'_>,
) -> Result<BytesStart<'static>, RedactError> {
    let format = state.format;
    let path = state.path;
    let schema = schema_node(path, reader, start.name());
    let word = format == Format::Docx && word_node(reader, start.name());
    let custom_property = path == "docprops/custom.xml"
        && element == "property"
        && matches!(reader.resolver().resolve_element(start.name()).0,
            ResolveResult::Bound(ns) if matches!(ns.as_ref(),
                b"http://schemas.openxmlformats.org/officeDocument/2006/custom-properties"
                | b"http://purl.oclc.org/ooxml/officeDocument/customProperties"));
    if custom_property {
        state.custom_property += 1;
    }
    let modern_comment = format == Format::Docx
        && element == "commentExtensible"
        && matches!(reader.resolver().resolve_element(start.name()).0,
            ResolveResult::Bound(ns) if ns.as_ref()
                == b"http://schemas.microsoft.com/office/word/2018/wordml/cex");
    let mut attributes = Vec::new();
    for attribute in start.attributes() {
        let attribute = attribute.map_err(|error| xml_error(path, error))?;
        let key = String::from_utf8_lossy(attribute.key.as_ref()).into_owned();
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
            .map_err(|error| xml_error(path, error))?
            .into_owned();
        attributes.push((key, value));
    }
    if word && element == "style" {
        state.current_style = attributes
            .iter()
            .find(|(key, _)| attribute_local(key) == "styleId" && word_attribute(reader, key))
            .map(|(_, value)| value.clone());
    }

    let (relationship, external) = relationship_mode(path, element, &attributes);
    let mut wrote_target_mode = false;
    if format == Format::Xlsx && element == "c" {
        state.cell_type = None;
    }
    let mut output = start.into_owned();
    output.clear_attributes();
    for (key, value) in attributes {
        let local = attribute_local(&key);
        let instance = path.starts_with("customxml/")
            && schema::is_instance_namespace(
                &reader.resolver().resolve_attribute(QName(key.as_bytes())).0,
            );
        if local.eq_ignore_ascii_case("gfxdata")
            && matches!(reader.resolver().resolve_attribute(QName(key.as_bytes())).0,
                ResolveResult::Bound(namespace) if namespace.as_ref() == b"urn:schemas-microsoft-com:office:office")
            || schema && is_unqualified(&key) && schema::drop_attribute(element, local)
            || instance
                && matches!(
                    local,
                    "type" | "schemaLocation" | "noNamespaceSchemaLocation"
                )
        {
            state.report.attributes += 1;
            continue;
        }
        if relationship && is_unqualified(&key) && local.eq_ignore_ascii_case("TargetMode") {
            if !external {
                output.push_attribute((key.as_str(), value.as_str()));
            } else if !wrote_target_mode {
                output.push_attribute(("TargetMode", "External"));
            }
            wrote_target_mode = true;
            continue;
        }
        let style_replacement = if word && word_attribute(reader, &key) {
            state
                .styles
                .replacement(element, local, &value, state.current_style.as_deref())
        } else {
            None
        };
        let replacement =
            if external && is_unqualified(&key) && local.eq_ignore_ascii_case("Target") {
                Some("https://example.com".to_owned())
            } else if custom_property && key == "name" {
                Some(format!("RedactedProperty{}", state.custom_property))
            } else if word && local == "date" && word_attribute(reader, &key)
                || modern_comment
                    && local == "dateUtc"
                    && matches!(reader.resolver().resolve_attribute(QName(key.as_bytes())).0,
                        ResolveResult::Bound(ns) if ns.as_ref()
                            == b"http://schemas.microsoft.com/office/word/2018/wordml/cex")
            {
                Some("1970-01-01T00:00:00Z".to_owned())
            } else if style_replacement.is_some() {
                style_replacement
            } else if instance && local == "nil" {
                let value = value.trim_matches([' ', '\t', '\r', '\n']);
                Some(
                    if matches!(value, "true" | "false" | "1" | "0") {
                        value
                    } else {
                        "false"
                    }
                    .to_owned(),
                )
            } else if !key.starts_with("xmlns")
                && !(schema && is_unqualified(&key) && schema::preserve_attribute(element, local))
                && sensitive_attribute(format, path, element, local, &value)
            {
                Some(state.masker.replace(&value)?)
            } else {
                None
            };
        if let Some(replacement) = replacement {
            if replacement != value {
                state.report.attributes += 1;
            }
            output.push_attribute((key.as_str(), replacement.as_str()));
        } else {
            output.push_attribute((key.as_str(), value.as_str()));
        }
        if format == Format::Xlsx && element == "c" && local == "t" {
            state.cell_type = Some(value);
        }
    }
    if relationship && external && !wrote_target_mode {
        output.push_attribute(("TargetMode", "External"));
    }
    Ok(output)
}

fn schema_node(path: &str, reader: &NsReader<&[u8]>, name: QName<'_>) -> bool {
    path.starts_with("customxml/")
        && schema::is_schema_namespace(&reader.resolver().resolve_element(name).0)
}

fn word_namespace(namespace: ResolveResult<'_>) -> bool {
    matches!(namespace, ResolveResult::Bound(ns) if matches!(ns.as_ref(),
        b"http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            | b"http://purl.oclc.org/ooxml/wordprocessingml/main"))
}

fn word_node(reader: &NsReader<&[u8]>, name: QName<'_>) -> bool {
    word_namespace(reader.resolver().resolve_element(name).0)
}

fn word_attribute(reader: &NsReader<&[u8]>, name: &str) -> bool {
    word_namespace(
        reader
            .resolver()
            .resolve_attribute(QName(name.as_bytes()))
            .0,
    )
}

fn fixed_metadata_value(
    path: &str,
    reader: &NsReader<&[u8]>,
    name: QName<'_>,
) -> Option<&'static str> {
    let application = path.eq_ignore_ascii_case("docprops/app.xml");
    let custom = path.eq_ignore_ascii_case("docprops/custom.xml");
    if !application && !custom {
        return None;
    }
    let namespace = reader.resolver().resolve_element(name).0;
    if application
        && name.local_name().as_ref() == b"AppVersion"
        && matches!(&namespace,
            ResolveResult::Bound(ns) if matches!(ns.as_ref(),
                b"http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
                | b"http://purl.oclc.org/ooxml/officeDocument/extendedProperties"))
    {
        Some("0.0000")
    } else if custom
        && matches!(
            name.local_name().as_ref(),
            b"i1"
                | b"i2"
                | b"i4"
                | b"i8"
                | b"int"
                | b"ui1"
                | b"ui2"
                | b"ui4"
                | b"ui8"
                | b"uint"
                | b"r4"
                | b"r8"
                | b"decimal"
        )
        && matches!(namespace,
            ResolveResult::Bound(ns) if matches!(ns.as_ref(),
                b"http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"
                | b"http://purl.oclc.org/ooxml/officeDocument/docPropsVTypes"))
    {
        Some("0")
    } else {
        None
    }
}

#[derive(Clone, Copy)]
enum Replacement {
    Text,
    Number,
    Formula,
    Date,
    Boolean,
    Error,
}

fn replacement_kind(
    format: Format,
    path: &str,
    stack: &[String],
    cell_type: Option<&str>,
) -> Option<Replacement> {
    let element = stack.last().map(String::as_str)?;
    let lower = path.to_ascii_lowercase();
    if lower.ends_with(".rels") {
        return None;
    }
    if lower == "docprops/core.xml" {
        return match element {
            "created" | "modified" | "lastPrinted" => Some(Replacement::Date),
            "revision" => Some(Replacement::Number),
            "title" | "subject" | "creator" | "keywords" | "description" | "lastModifiedBy"
            | "category" | "contentStatus" => Some(Replacement::Text),
            _ => None,
        };
    }
    if lower == "docprops/app.xml" {
        return matches!(
            element,
            "Application"
                | "AppVersion"
                | "Company"
                | "Manager"
                | "Template"
                | "HyperlinkBase"
                | "lpstr"
                | "lpwstr"
                | "bstr"
        )
        .then_some(Replacement::Text);
    }
    if lower == "docprops/custom.xml" {
        return match element {
            "i1" | "i2" | "i4" | "i8" | "int" | "uint" | "ui1" | "ui2" | "ui4" | "ui8" | "r4"
            | "r8" | "decimal" => Some(Replacement::Number),
            "bool" => Some(Replacement::Boolean),
            "date" | "filetime" => Some(Replacement::Date),
            _ => Some(Replacement::Text),
        };
    }
    if lower.starts_with("customxml/") && !lower.contains("itemprops") {
        return Some(Replacement::Text);
    }
    if lower.contains("/charts/") {
        return match element {
            "f" => Some(Replacement::Formula),
            "v" if stack.iter().any(|name| name == "strCache" || name == "tx") => {
                Some(Replacement::Text)
            }
            "v" => Some(Replacement::Number),
            "t" => Some(Replacement::Text),
            _ => None,
        };
    }

    match format {
        Format::Docx => matches!(element, "t" | "delText" | "instrText" | "delInstrText")
            .then_some(if matches!(element, "instrText" | "delInstrText") {
                Replacement::Formula
            } else {
                Replacement::Text
            }),
        Format::Pptx => matches!(element, "t" | "text").then_some(Replacement::Text),
        Format::Xlsx => match element {
            "t" | "author" | "oddHeader" | "oddFooter" | "evenHeader" | "evenFooter"
            | "firstHeader" | "firstFooter" => Some(Replacement::Text),
            "f" | "formula1" | "formula2" | "definedName" => Some(Replacement::Formula),
            "v" if cell_type == Some("s") => None,
            "v" if matches!(cell_type, Some("str" | "inlineStr")) => Some(Replacement::Text),
            "v" if cell_type == Some("e") => Some(Replacement::Error),
            "v" => Some(Replacement::Number),
            _ => None,
        },
        Format::Auto => None,
    }
}

fn replace_text(
    text: &str,
    kind: Replacement,
    masker: &mut TextMasker,
) -> Result<String, RedactError> {
    if text.trim().is_empty() {
        return Ok(text.to_owned());
    }
    Ok(match kind {
        Replacement::Text => return masker.replace(text),
        Replacement::Number => numeric_placeholder(text),
        Replacement::Formula => "0".to_owned(),
        Replacement::Date => "1970-01-01T00:00:00Z".to_owned(),
        Replacement::Boolean => "false".to_owned(),
        Replacement::Error => "#N/A".to_owned(),
    })
}

fn sensitive_attribute(
    format: Format,
    path: &str,
    element: &str,
    attribute: &str,
    value: &str,
) -> bool {
    let lower = path.to_ascii_lowercase();
    if lower.ends_with(".rels") {
        return false;
    }
    if lower == "docprops/custom.xml" && attribute == "name" {
        return true;
    }
    if lower.starts_with("customxml/") && !lower.contains("itemprops") {
        return !matches!(attribute, "id" | "Id");
    }
    if matches!(element, "docPr" | "cNvPr") && matches!(attribute, "name" | "descr" | "title") {
        return true;
    }
    if element == "textpath" && attribute == "string" {
        return true;
    }
    match format {
        Format::Docx => {
            matches!(attribute, "author" | "initials")
                || element == "fldSimple" && attribute == "instr"
                || element == "hyperlink" && matches!(attribute, "tooltip" | "tgtFrame")
                || matches!(element, "alias" | "tag" | "docVar")
                    && matches!(attribute, "name" | "val")
        }
        Format::Xlsx => {
            element == "sheet" && attribute == "name"
                || element == "definedName" && attribute == "name" && !value.starts_with("_xlnm.")
                || matches!(element, "table" | "tableColumn")
                    && matches!(attribute, "name" | "displayName")
                || element == "dataValidation"
                    && matches!(attribute, "prompt" | "promptTitle" | "error" | "errorTitle")
                || element == "hyperlink" && matches!(attribute, "display" | "tooltip" | "location")
                || element == "filter" && attribute == "val"
        }
        Format::Pptx => {
            element == "cSld" && attribute == "name"
                || element == "sldLayout" && attribute == "matchingName"
                || matches!(element, "theme" | "clrScheme" | "fontScheme" | "fmtScheme")
                    && attribute == "name"
                || element == "tblStyle" && attribute == "styleName"
                || element == "cmAuthor" && matches!(attribute, "name" | "initials")
                || element == "author" && matches!(attribute, "name" | "initials" | "userId")
                || element == "tag" && matches!(attribute, "name" | "val")
                || element == "custShow" && attribute == "name"
        }
        Format::Auto => false,
    }
}

fn numeric_placeholder(text: &str) -> String {
    let mut valid = true;
    let replacement: String = text
        .chars()
        .map(|character| match character {
            '0'..='9' => '8',
            '-' | '+' | '.' | 'e' | 'E' | ' ' | '\t' | '\r' | '\n' => character,
            _ => {
                valid = false;
                'x'
            }
        })
        .collect();
    if valid {
        replacement
    } else {
        placeholder(text)
    }
}

fn charge_text(report: &mut RedactionReport, text: &str) {
    if !text.trim().is_empty() {
        report.text_nodes += 1;
        report.characters += text
            .chars()
            .filter(|character| !character.is_whitespace())
            .count();
    }
}

fn local_name(name: &[u8]) -> String {
    String::from_utf8_lossy(name).into_owned()
}

/// Whether the element is a relationship in a package relationship part, and
/// whether it points outside the package.
fn relationship_mode(path: &str, element: &str, attributes: &[(String, String)]) -> (bool, bool) {
    if !element.eq_ignore_ascii_case("Relationship") {
        return (false, false);
    }
    let package_part = path.to_ascii_lowercase().ends_with(".rels");
    (
        package_part,
        rels::external_relationship(attributes, package_part),
    )
}

fn xml_error(path: &str, error: impl fmt::Display) -> RedactError {
    RedactError::Xml {
        part: path.to_owned(),
        message: error.to_string(),
    }
}

use std::fmt;
