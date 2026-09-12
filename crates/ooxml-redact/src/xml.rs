use quick_xml::events::{BytesCData, BytesStart, BytesText, Event};
use quick_xml::{Reader, Writer, XmlVersion};

use crate::rels::{self, attribute_local, is_unqualified};
use crate::{Format, RedactError, RedactionReport};

pub(crate) fn redact_xml(
    format: Format,
    path: &str,
    bytes: &[u8],
    report: &mut RedactionReport,
) -> Result<Vec<u8>, RedactError> {
    let mut reader = Reader::from_reader(bytes);
    reader.config_mut().trim_text(false);
    let mut writer = Writer::new(Vec::with_capacity(bytes.len()));
    let mut context = XmlContext {
        stack: Vec::new(),
        cell_type: None,
    };

    loop {
        let event = reader
            .read_event()
            .map_err(|error| xml_error(path, error))?;
        match event {
            Event::Start(start) => {
                let local = local_name(start.name().local_name().as_ref());
                let (rewritten, section_name) =
                    rewrite_start(format, path, &reader, start, &local, report, &mut context)?;
                context.stack.push(ElementContext {
                    name: local,
                    section_name,
                });
                writer
                    .write_event(Event::Start(rewritten))
                    .map_err(|error| xml_error(path, error))?;
            }
            Event::Empty(start) => {
                let local = local_name(start.name().local_name().as_ref());
                let (rewritten, _) =
                    rewrite_start(format, path, &reader, start, &local, report, &mut context)?;
                writer
                    .write_event(Event::Empty(rewritten))
                    .map_err(|error| xml_error(path, error))?;
            }
            Event::End(end) => {
                if context.stack.last().is_some_and(|entry| entry.name == "c") {
                    context.cell_type = None;
                }
                context.stack.pop();
                writer
                    .write_event(Event::End(end))
                    .map_err(|error| xml_error(path, error))?;
            }
            Event::Text(text) => {
                if let Some(kind) =
                    replacement_kind(format, path, &context.stack, context.cell_type.as_deref())
                {
                    let decoded = text.decode().map_err(|error| xml_error(path, error))?;
                    let unescaped = quick_xml::escape::unescape(&decoded)
                        .map_err(|error| xml_error(path, error))?;
                    let replacement = replace_text(&unescaped, kind);
                    charge_text(report, &unescaped);
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
                    replacement_kind(format, path, &context.stack, context.cell_type.as_deref())
                {
                    let decoded = text.decode().map_err(|error| xml_error(path, error))?;
                    let replacement = replace_text(&decoded, kind);
                    charge_text(report, &decoded);
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
                if replacement_kind(format, path, &context.stack, context.cell_type.as_deref())
                    .is_some()
                {
                    report.text_nodes += 1;
                    report.characters += 1;
                    writer
                        .write_event(Event::Text(BytesText::new("x")))
                        .map_err(|error| xml_error(path, error))?;
                } else {
                    writer
                        .write_event(Event::GeneralRef(reference))
                        .map_err(|error| xml_error(path, error))?;
                }
            }
            Event::Comment(_) | Event::PI(_) => {
                report.xml_comments += 1;
            }
            Event::DocType(_) => {
                return Err(RedactError::Xml {
                    part: path.to_owned(),
                    message: "DTD/entity declarations are forbidden".to_owned(),
                });
            }
            Event::Eof => break,
            other => writer
                .write_event(other)
                .map_err(|error| xml_error(path, error))?,
        }
    }

    Ok(writer.into_inner())
}

struct ElementContext {
    name: String,
    section_name: Option<String>,
}

struct XmlContext {
    stack: Vec<ElementContext>,
    cell_type: Option<String>,
}

fn rewrite_start(
    format: Format,
    path: &str,
    reader: &Reader<&[u8]>,
    start: BytesStart<'_>,
    element: &str,
    report: &mut RedactionReport,
    context: &mut XmlContext,
) -> Result<(BytesStart<'static>, Option<String>), RedactError> {
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

    let section_name = (element == "Section")
        .then(|| {
            attributes
                .iter()
                .find(|(key, _)| attribute_local(key) == "N")
                .map(|(_, value)| value.clone())
        })
        .flatten();
    let vsdx_user_data_cell = format == Format::Vsdx
        && element == "Cell"
        && context
            .stack
            .iter()
            .any(|context| matches!(context.section_name.as_deref(), Some("Property" | "User")))
        && attributes.iter().any(|(key, value)| {
            attribute_local(key) == "N" && matches!(value.as_str(), "Value" | "Prompt" | "Label")
        });

    let (relationship, external) = relationship_mode(path, element, &attributes);
    let mut wrote_target_mode = false;
    if format == Format::Xlsx && element == "c" {
        context.cell_type = None;
    }
    let mut output = start.into_owned();
    output.clear_attributes();
    for (key, value) in attributes {
        let local = attribute_local(&key);
        if relationship && is_unqualified(&key) && local.eq_ignore_ascii_case("TargetMode") {
            if !external {
                output.push_attribute((key.as_str(), value.as_str()));
            } else if !wrote_target_mode {
                output.push_attribute(("TargetMode", "External"));
            }
            wrote_target_mode = true;
            continue;
        }
        let replacement =
            if external && is_unqualified(&key) && local.eq_ignore_ascii_case("Target") {
                Some("https://example.com".to_owned())
            } else if !key.starts_with("xmlns")
                && sensitive_attribute(format, path, element, local, &value, vsdx_user_data_cell)
            {
                Some(placeholder(&value))
            } else {
                None
            };
        if let Some(replacement) = replacement {
            if replacement != value {
                report.attributes += 1;
            }
            output.push_attribute((key.as_str(), replacement.as_str()));
        } else {
            output.push_attribute((key.as_str(), value.as_str()));
        }
        if format == Format::Xlsx && element == "c" && local == "t" {
            context.cell_type = Some(value);
        }
    }
    if relationship && external && !wrote_target_mode {
        output.push_attribute(("TargetMode", "External"));
    }
    Ok((output, section_name))
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
    stack: &[ElementContext],
    cell_type: Option<&str>,
) -> Option<Replacement> {
    let element = stack.last().map(|context| context.name.as_str())?;
    let lower = path.to_ascii_lowercase();
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
            "v" if stack
                .iter()
                .any(|context| context.name == "strCache" || context.name == "tx") =>
            {
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
        Format::Vsdx => stack
            .iter()
            .any(|context| context.name == "Text")
            .then_some(Replacement::Text),
        Format::Auto => None,
    }
}

fn replace_text(text: &str, kind: Replacement) -> String {
    if text.trim().is_empty() {
        return text.to_owned();
    }
    match kind {
        Replacement::Text => placeholder(text),
        Replacement::Number => numeric_placeholder(text),
        Replacement::Formula => "0".to_owned(),
        Replacement::Date => "1970-01-01T00:00:00Z".to_owned(),
        Replacement::Boolean => "false".to_owned(),
        Replacement::Error => "#N/A".to_owned(),
    }
}

fn sensitive_attribute(
    format: Format,
    path: &str,
    element: &str,
    attribute: &str,
    value: &str,
    vsdx_user_data_cell: bool,
) -> bool {
    let lower = path.to_ascii_lowercase();
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
                || lower == "word/styles.xml" && element == "name" && attribute == "val"
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
                || element == "cmAuthor" && matches!(attribute, "name" | "initials")
                || element == "author" && matches!(attribute, "name" | "initials" | "userId")
                || element == "tag" && matches!(attribute, "name" | "val")
                || element == "custShow" && attribute == "name"
        }
        Format::Vsdx => vsdx_user_data_cell && matches!(attribute, "V" | "F"),
        Format::Auto => false,
    }
}

fn placeholder(text: &str) -> String {
    text.chars()
        .map(|character| {
            if character.is_whitespace() {
                character
            } else {
                'x'
            }
        })
        .collect()
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
