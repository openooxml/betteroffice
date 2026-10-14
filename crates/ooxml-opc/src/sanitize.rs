use std::borrow::Cow;
use std::collections::{BTreeMap, HashSet};

use quick_xml::events::{BytesCData, BytesStart, BytesText, Event};
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::{NsReader, Reader, Writer, XmlVersion};

use crate::{rezip_parts_preserving, unzip_parts};

const CONTENT_TYPES_NAMESPACE: &[u8] =
    b"http://schemas.openxmlformats.org/package/2006/content-types";
const RELATIONSHIPS_NAMESPACE: &[u8] =
    b"http://schemas.openxmlformats.org/package/2006/relationships";
const MARKUP_COMPATIBILITY_NAMESPACE: &[u8] =
    b"http://schemas.openxmlformats.org/markup-compatibility/2006";
const OFFICE_DOCUMENT_RELATIONSHIP: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
const STRICT_OFFICE_DOCUMENT_RELATIONSHIP: &str =
    "http://purl.oclc.org/ooxml/officeDocument/relationships/officeDocument";
const VISIO_DOCUMENT_RELATIONSHIP: &str =
    "http://schemas.microsoft.com/visio/2010/relationships/document";
const MAX_XML_DEPTH: usize = 256;
const MAX_PACKAGE_XML_EVENTS: usize = 4_000_000;
const MAX_PACKAGE_XML_ATTRIBUTE_BYTES: usize = 64 * 1024 * 1024;
const MAX_XML_ATTRIBUTES_PER_ELEMENT: usize = 1_024;
const MAX_XML_ATTRIBUTE_BYTES_PER_ELEMENT: usize = 4 * 1024 * 1024;

pub fn sanitize_package(data: &[u8]) -> Result<Vec<u8>, String> {
    sanitize_package_inner(data, None)
}

pub fn sanitize_package_for_format(data: &[u8], expected_format: &str) -> Result<Vec<u8>, String> {
    if !matches!(
        expected_format,
        "docx" | "dotx" | "xlsx" | "xltx" | "pptx" | "potx" | "ppsx" | "vsdx" | "vstx"
    ) {
        return Err(format!("unsupported OOXML format: {expected_format}"));
    }
    sanitize_package_inner(data, Some(expected_format))
}

fn sanitize_package_inner(data: &[u8], expected_format: Option<&str>) -> Result<Vec<u8>, String> {
    let mut parts = unzip_parts(data)?;
    let mut xml_budget = XmlBudget::default();
    let (detected, content_types) = if expected_format.is_some() {
        let (kind, content_types) = detect_package_kind_with_budget(&parts, &mut xml_budget)
            .map_err(|error| error.to_string())?;
        (kind, Some(content_types))
    } else {
        detect_format_with_budget(&parts, &mut xml_budget)?
    };
    if let Some(expected) = expected_format
        && detected.sanitized_format() != expected
    {
        return Err(format!(
            "claimed {expected} content does not match detected {} package",
            detected.format()
        ));
    }

    let mut removed: HashSet<String> = parts
        .iter()
        .filter(|(path, _)| {
            !path.eq_ignore_ascii_case("[Content_Types].xml")
                && dangerous_part(path, content_types.as_ref())
        })
        .map(|(path, _)| normalize_part_name(path).into_owned())
        .collect();
    if let Some(content_types) = &content_types {
        removed.extend(
            content_types
                .dangerous
                .iter()
                .filter(|path| path.as_str() != "[content_types].xml")
                .cloned(),
        );
    }

    if matches!(
        detected,
        DocumentKind::Xlsx
            | DocumentKind::Xlsm
            | DocumentKind::Xltx
            | DocumentKind::Xltm
            | DocumentKind::Xlam
    ) {
        sanitize_workbook_sheets(
            &mut parts,
            &mut removed,
            content_types.as_ref(),
            &mut xml_budget,
        )?;
    }
    if !removed.is_empty() {
        parts.retain(|(path, _)| !removed.contains(normalize_part_name(path).as_ref()));
    }
    for (path, bytes) in &mut parts {
        let lower = path.to_ascii_lowercase();
        let content_type = content_types
            .as_ref()
            .and_then(|content_types| content_types.content_type_for(path));
        if lower == "[content_types].xml" {
            *bytes = sanitize_content_types(bytes, &removed, path, &mut xml_budget)?;
        } else if lower.ends_with(".rels")
            || content_type.is_some_and(is_relationships_content_type)
        {
            *bytes = sanitize_relationships(
                bytes,
                &removed,
                content_types.as_ref(),
                path,
                &mut xml_budget,
            )?;
        } else if is_xml_part(&lower) || content_type.is_some_and(is_xml_content_type) {
            *bytes = neutralize_fields(bytes, path, &mut xml_budget)?;
        }
    }
    rezip_parts_preserving(&parts, data)
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, PartialOrd, Ord)]
#[non_exhaustive]
pub enum DocumentKind {
    Docx,
    Docm,
    Xlsx,
    Xlsm,
    Pptx,
    Pptm,
    Vsdx,
    Vsdm,
    Vssx,
    Vssm,
    Vstx,
    Vstm,
    Dotx,
    Dotm,
    Xltx,
    Xltm,
    Xlam,
    Potx,
    Potm,
    Ppsx,
    Ppsm,
    Ppam,
}

impl DocumentKind {
    pub fn format(self) -> &'static str {
        match self {
            Self::Docx => "docx",
            Self::Docm => "docm",
            Self::Dotx => "dotx",
            Self::Dotm => "dotm",
            Self::Xlsx => "xlsx",
            Self::Xlsm => "xlsm",
            Self::Xltx => "xltx",
            Self::Xltm => "xltm",
            Self::Xlam => "xlam",
            Self::Pptx => "pptx",
            Self::Pptm => "pptm",
            Self::Potx => "potx",
            Self::Potm => "potm",
            Self::Ppsx => "ppsx",
            Self::Ppsm => "ppsm",
            Self::Ppam => "ppam",
            Self::Vsdx => "vsdx",
            Self::Vsdm => "vsdm",
            Self::Vssx => "vssx",
            Self::Vssm => "vssm",
            Self::Vstx => "vstx",
            Self::Vstm => "vstm",
        }
    }

    fn sanitized_format(self) -> &'static str {
        match self {
            Self::Docm => "docx",
            Self::Dotm => "dotx",
            Self::Xlsm | Self::Xlam => "xlsx",
            Self::Xltm => "xltx",
            Self::Pptm | Self::Ppam => "pptx",
            Self::Potm => "potx",
            Self::Ppsm => "ppsx",
            _ => self.format(),
        }
    }
}

#[derive(Clone, Debug, PartialEq, Eq)]
#[non_exhaustive]
pub enum DocumentKindError {
    MissingContentTypes,
    InvalidContentTypes(String),
    MissingPackageRelationships,
    InvalidPackageRelationships(String),
    MissingMainDocumentRelationship,
    ConflictingMainDocumentRelationships(Vec<String>),
    MissingMainDocumentPart(String),
    MissingMainDocumentKind,
}

impl std::fmt::Display for DocumentKindError {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::MissingContentTypes => formatter.write_str("missing [Content_Types].xml"),
            Self::InvalidContentTypes(error) => {
                write!(formatter, "invalid [Content_Types].xml: {error}")
            }
            Self::MissingPackageRelationships => formatter.write_str("missing _rels/.rels"),
            Self::InvalidPackageRelationships(error) => {
                write!(formatter, "invalid _rels/.rels: {error}")
            }
            Self::MissingMainDocumentRelationship => {
                formatter.write_str("missing main document relationship")
            }
            Self::ConflictingMainDocumentRelationships(parts) => {
                write!(
                    formatter,
                    "conflicting main document relationships: {parts:?}"
                )
            }
            Self::MissingMainDocumentPart(part) => {
                write!(formatter, "missing main document part: {part}")
            }
            Self::MissingMainDocumentKind => {
                formatter.write_str("missing recognized main document content type")
            }
        }
    }
}

impl std::error::Error for DocumentKindError {}

pub fn detect_package_kind(parts: &[(String, Vec<u8>)]) -> Result<DocumentKind, DocumentKindError> {
    let mut xml_budget = XmlBudget::default();
    detect_package_kind_with_budget(parts, &mut xml_budget).map(|(kind, _)| kind)
}

fn detect_package_kind_with_budget(
    parts: &[(String, Vec<u8>)],
    xml_budget: &mut XmlBudget,
) -> Result<(DocumentKind, ContentTypes), DocumentKindError> {
    let part_names: HashSet<String> = parts
        .iter()
        .map(|(path, _)| normalize_part_name(path).into_owned())
        .collect();
    let main_document = main_document_part(parts, &part_names, xml_budget)?;
    let content_types = parse_content_types_with_budget(parts, xml_budget)?;
    let content_type = content_types
        .content_type_for(&main_document.part_name)
        .ok_or(DocumentKindError::MissingMainDocumentKind)?;
    let kind =
        declared_main_kind(content_type).ok_or(DocumentKindError::MissingMainDocumentKind)?;
    if main_document.relationship_kind.accepts(kind) {
        Ok((kind, content_types))
    } else {
        Err(invalid_relationships(
            "main relationship type does not match document content type",
        ))
    }
}

pub struct ContentTypes {
    overrides: BTreeMap<String, String>,
    defaults: BTreeMap<String, String>,
    /// Parts named by any `Override` (any depth, any namespace prefix) whose
    /// declared content type is macro/embedding-dangerous.
    dangerous: HashSet<String>,
}

impl ContentTypes {
    pub fn content_type_for(&self, part_name: &str) -> Option<&str> {
        let part_name = normalize_part_name(part_name);
        if let Some(content_type) = self.overrides.get(part_name.as_ref()) {
            return Some(content_type);
        }
        let (_, extension) = part_name.rsplit_once('.')?;
        self.defaults
            .get(&extension.to_ascii_lowercase())
            .map(String::as_str)
    }
}

pub fn parse_content_types(parts: &[(String, Vec<u8>)]) -> Result<ContentTypes, DocumentKindError> {
    parse_content_types_with_budget(parts, &mut XmlBudget::default())
}

fn parse_content_types_with_budget(
    parts: &[(String, Vec<u8>)],
    xml_budget: &mut XmlBudget,
) -> Result<ContentTypes, DocumentKindError> {
    let (_, bytes) = parts
        .iter()
        .find(|(path, _)| path.eq_ignore_ascii_case("[Content_Types].xml"))
        .ok_or(DocumentKindError::MissingContentTypes)?;
    validate_xml_limits(bytes, "[Content_Types].xml", xml_budget)
        .map_err(DocumentKindError::InvalidContentTypes)?;
    let mut reader = NsReader::from_reader(bytes.as_slice());
    let mut overrides = BTreeMap::new();
    let mut defaults = BTreeMap::new();
    let mut dangerous = HashSet::new();
    let mut declaration_count = 0_usize;
    let mut depth = 0_usize;
    let mut root_seen = false;
    loop {
        match read_decoded_namespace_event(&mut reader)
            .map_err(|error| DocumentKindError::InvalidContentTypes(error.to_string()))?
        {
            (namespace, Event::Start(start)) => {
                let local = start.name().local_name();
                if depth == 0 {
                    if root_seen {
                        return Err(invalid_content_types("multiple root elements"));
                    }
                    root_seen = true;
                    if local.as_ref() != b"Types" || !is_content_types_namespace(&namespace) {
                        return Err(invalid_content_types("unexpected root element"));
                    }
                } else if depth == 1
                    && is_content_types_namespace(&namespace)
                    && matches!(local.as_ref(), b"Override" | b"Default")
                {
                    record_content_type(
                        &reader,
                        &start,
                        &mut overrides,
                        &mut defaults,
                        &mut declaration_count,
                    )?;
                }
                if local.as_ref() == b"Override" {
                    record_dangerous_override(&reader, &start, &mut dangerous)?;
                }
                depth += 1;
            }
            (namespace, Event::Empty(start)) => {
                let local = start.name().local_name();
                if depth == 0 {
                    if root_seen {
                        return Err(invalid_content_types("multiple root elements"));
                    }
                    root_seen = true;
                    if local.as_ref() != b"Types" || !is_content_types_namespace(&namespace) {
                        return Err(invalid_content_types("unexpected root element"));
                    }
                } else if depth == 1
                    && is_content_types_namespace(&namespace)
                    && matches!(local.as_ref(), b"Override" | b"Default")
                {
                    record_content_type(
                        &reader,
                        &start,
                        &mut overrides,
                        &mut defaults,
                        &mut declaration_count,
                    )?;
                }
                if local.as_ref() == b"Override" {
                    record_dangerous_override(&reader, &start, &mut dangerous)?;
                }
            }
            (_, Event::End(_)) => {
                if depth == 0 {
                    return Err(invalid_content_types("unexpected closing element"));
                }
                depth -= 1;
            }
            (_, Event::DocType(_)) => {
                return Err(invalid_content_types("DTD is forbidden"));
            }
            (_, Event::Text(text)) if depth == 0 && !xml_whitespace(text.as_ref()) => {
                return Err(invalid_content_types("text outside root element"));
            }
            (_, Event::CData(_)) if depth == 0 => {
                return Err(invalid_content_types("CDATA outside root element"));
            }
            (_, Event::Eof) => {
                if !root_seen {
                    return Err(invalid_content_types("missing root element"));
                }
                if depth != 0 {
                    return Err(invalid_content_types("unexpected EOF"));
                }
                break;
            }
            _ => {}
        }
    }
    Ok(ContentTypes {
        overrides,
        defaults,
        dangerous,
    })
}

fn record_dangerous_override(
    reader: &Reader<&[u8]>,
    start: &BytesStart<'_>,
    dangerous: &mut HashSet<String>,
) -> Result<(), DocumentKindError> {
    let values = content_type_attributes(reader, start)?;
    if let (Some(part_name), Some(content_type)) = (
        attribute_value(&values, "PartName"),
        attribute_value(&values, "ContentType"),
    ) && dangerous_content_type(content_type)
    {
        dangerous.insert(normalize_part_name(part_name).into_owned());
    }
    Ok(())
}

// Namespace names are attribute values, so references in them spell the same URI. Each
// declaration holding one is decoded once, into a binding that shadows the raw one; a value
// that does not decode or bind keeps the raw binding.
fn read_decoded_namespace_event<'r, 'i>(
    reader: &'r mut NsReader<&'i [u8]>,
) -> quick_xml::Result<(ResolveResult<'r>, Event<'i>)> {
    let event = reader.read_event()?;
    if let Event::Start(start) | Event::Empty(start) = &event {
        let decoder = reader.decoder();
        for attribute in start.attributes().with_checks(false) {
            let Ok(attribute) = attribute else { break };
            if let Some(prefix) = attribute.key.as_namespace_binding()
                && attribute.value.contains(&b'&')
                && let Ok(value) =
                    attribute.decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)
            {
                let _ = reader
                    .resolver_mut()
                    .add(prefix, Namespace(value.as_bytes()));
            }
        }
    }
    Ok(reader.resolver().resolve_event(event))
}

fn is_content_types_namespace(namespace: &ResolveResult<'_>) -> bool {
    matches!(namespace, ResolveResult::Bound(namespace)
        if namespace.0 == CONTENT_TYPES_NAMESPACE)
}

fn invalid_content_types(error: &str) -> DocumentKindError {
    DocumentKindError::InvalidContentTypes(error.to_owned())
}

fn record_content_type(
    reader: &Reader<&[u8]>,
    start: &BytesStart<'_>,
    overrides: &mut BTreeMap<String, String>,
    defaults: &mut BTreeMap<String, String>,
    declaration_count: &mut usize,
) -> Result<(), DocumentKindError> {
    *declaration_count += 1;
    if *declaration_count > crate::MAX_ENTRY_COUNT * 2 {
        return Err(invalid_content_types("too many content type declarations"));
    }
    let values = content_type_attributes(reader, start)?;
    match start.name().local_name().as_ref() {
        b"Override" => {
            let part_name = attribute_value(&values, "PartName")
                .ok_or_else(|| invalid_content_types("Override missing PartName"))?;
            let content_type = attribute_value(&values, "ContentType")
                .ok_or_else(|| invalid_content_types("Override missing ContentType"))?;
            let part_name = normalize_part_name(part_name).into_owned();
            if overrides
                .insert(part_name.clone(), content_type.to_owned())
                .is_some()
            {
                return Err(invalid_content_types(&format!(
                    "duplicate Override for {part_name}"
                )));
            }
        }
        b"Default" => {
            let extension = attribute_value(&values, "Extension")
                .ok_or_else(|| invalid_content_types("Default missing Extension"))?
                .to_ascii_lowercase();
            let content_type = attribute_value(&values, "ContentType")
                .ok_or_else(|| invalid_content_types("Default missing ContentType"))?;
            if defaults
                .insert(extension.clone(), content_type.to_owned())
                .is_some()
            {
                return Err(invalid_content_types(&format!(
                    "duplicate Default for {extension}"
                )));
            }
        }
        _ => {}
    }
    Ok(())
}

fn content_type_attributes(
    reader: &Reader<&[u8]>,
    start: &BytesStart<'_>,
) -> Result<Vec<(String, String)>, DocumentKindError> {
    start
        .attributes()
        .map(|attribute| {
            let attribute = attribute
                .map_err(|error| DocumentKindError::InvalidContentTypes(error.to_string()))?;
            let key = String::from_utf8_lossy(attribute.key.as_ref()).into_owned();
            let value = attribute
                .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
                .map_err(|error| DocumentKindError::InvalidContentTypes(error.to_string()))?
                .into_owned();
            Ok((key, value))
        })
        .collect()
}

struct MainDocument {
    part_name: String,
    relationship_kind: MainRelationshipKind,
}

#[derive(Clone, Copy)]
enum MainRelationshipKind {
    OfficeDocument,
    VisioDocument,
}

impl MainRelationshipKind {
    fn accepts(self, kind: DocumentKind) -> bool {
        match self {
            Self::OfficeDocument => matches!(
                kind,
                DocumentKind::Docx
                    | DocumentKind::Docm
                    | DocumentKind::Dotx
                    | DocumentKind::Dotm
                    | DocumentKind::Xlsx
                    | DocumentKind::Xlsm
                    | DocumentKind::Xltx
                    | DocumentKind::Xltm
                    | DocumentKind::Xlam
                    | DocumentKind::Pptx
                    | DocumentKind::Pptm
                    | DocumentKind::Potx
                    | DocumentKind::Potm
                    | DocumentKind::Ppsx
                    | DocumentKind::Ppsm
                    | DocumentKind::Ppam
            ),
            Self::VisioDocument => matches!(
                kind,
                DocumentKind::Vsdx
                    | DocumentKind::Vsdm
                    | DocumentKind::Vssx
                    | DocumentKind::Vssm
                    | DocumentKind::Vstx
                    | DocumentKind::Vstm
            ),
        }
    }
}

fn main_document_part(
    parts: &[(String, Vec<u8>)],
    part_names: &HashSet<String>,
    xml_budget: &mut XmlBudget,
) -> Result<MainDocument, DocumentKindError> {
    let (_, bytes) = parts
        .iter()
        .find(|(path, _)| path.eq_ignore_ascii_case("_rels/.rels"))
        .ok_or(DocumentKindError::MissingPackageRelationships)?;
    validate_xml_limits(bytes, "_rels/.rels", xml_budget)
        .map_err(DocumentKindError::InvalidPackageRelationships)?;
    let mut reader = NsReader::from_reader(bytes.as_slice());
    let mut targets = Vec::new();
    let mut depth = 0_usize;
    let mut root_seen = false;
    loop {
        match read_decoded_namespace_event(&mut reader)
            .map_err(|error| DocumentKindError::InvalidPackageRelationships(error.to_string()))?
        {
            (namespace, Event::Start(start)) => {
                let local = start.name().local_name();
                if depth == 0 {
                    if root_seen {
                        return Err(invalid_relationships("multiple root elements"));
                    }
                    root_seen = true;
                    if local.as_ref() != b"Relationships" || !is_relationships_namespace(&namespace)
                    {
                        return Err(invalid_relationships("unexpected root element"));
                    }
                } else if depth == 1
                    && local.as_ref() == b"Relationship"
                    && is_relationships_namespace(&namespace)
                {
                    record_main_relationship(&reader, &start, &mut targets)?;
                }
                depth += 1;
            }
            (namespace, Event::Empty(start)) => {
                let local = start.name().local_name();
                if depth == 0 {
                    if root_seen {
                        return Err(invalid_relationships("multiple root elements"));
                    }
                    root_seen = true;
                    if local.as_ref() != b"Relationships" || !is_relationships_namespace(&namespace)
                    {
                        return Err(invalid_relationships("unexpected root element"));
                    }
                } else if depth == 1
                    && local.as_ref() == b"Relationship"
                    && is_relationships_namespace(&namespace)
                {
                    record_main_relationship(&reader, &start, &mut targets)?;
                }
            }
            (_, Event::End(_)) => {
                if depth == 0 {
                    return Err(invalid_relationships("unexpected closing element"));
                }
                depth -= 1;
            }
            (_, Event::DocType(_)) => return Err(invalid_relationships("DTD is forbidden")),
            (_, Event::Text(text)) if depth == 0 && !xml_whitespace(text.as_ref()) => {
                return Err(invalid_relationships("text outside root element"));
            }
            (_, Event::CData(_)) if depth == 0 => {
                return Err(invalid_relationships("CDATA outside root element"));
            }
            (_, Event::Eof) => {
                if !root_seen {
                    return Err(invalid_relationships("missing root element"));
                }
                if depth != 0 {
                    return Err(invalid_relationships("unexpected EOF"));
                }
                break;
            }
            _ => {}
        }
    }
    let (main_part, relationship_kind) = match targets.as_slice() {
        [] => return Err(DocumentKindError::MissingMainDocumentRelationship),
        [(target, relationship_kind)] => (target.clone(), *relationship_kind),
        _ => {
            return Err(DocumentKindError::ConflictingMainDocumentRelationships(
                targets.into_iter().map(|(target, _)| target).collect(),
            ));
        }
    };
    if !part_names.contains(&main_part) {
        return Err(DocumentKindError::MissingMainDocumentPart(main_part));
    }
    Ok(MainDocument {
        part_name: main_part,
        relationship_kind,
    })
}

fn is_relationships_namespace(namespace: &ResolveResult<'_>) -> bool {
    matches!(namespace, ResolveResult::Bound(namespace)
        if namespace.0 == RELATIONSHIPS_NAMESPACE)
}

fn invalid_relationships(error: &str) -> DocumentKindError {
    DocumentKindError::InvalidPackageRelationships(error.to_owned())
}

fn record_main_relationship(
    reader: &Reader<&[u8]>,
    start: &BytesStart<'_>,
    targets: &mut Vec<(String, MainRelationshipKind)>,
) -> Result<(), DocumentKindError> {
    let values = attributes(reader, start, "_rels/.rels")
        .map_err(DocumentKindError::InvalidPackageRelationships)?;
    let relationship_type = attribute_value(&values, "Type")
        .ok_or_else(|| invalid_relationships("Relationship missing Type"))?;
    let target = attribute_value(&values, "Target")
        .ok_or_else(|| invalid_relationships("Relationship missing Target"))?;
    let relationship_kind = match relationship_type {
        OFFICE_DOCUMENT_RELATIONSHIP | STRICT_OFFICE_DOCUMENT_RELATIONSHIP => {
            MainRelationshipKind::OfficeDocument
        }
        VISIO_DOCUMENT_RELATIONSHIP => MainRelationshipKind::VisioDocument,
        _ => return Ok(()),
    };
    let target_mode = attribute_value(&values, "TargetMode")
        .unwrap_or_default()
        .trim();
    if (!target_mode.is_empty() && !target_mode.eq_ignore_ascii_case("Internal"))
        || external_target(target)
    {
        return Err(invalid_relationships(
            "main document relationship must be internal",
        ));
    }
    // The main part decides the package kind, so a target above the root stays invalid.
    let target = resolve_target("_rels/.rels", target)
        .and_then(|(target, climbed)| (!climbed).then_some(target))
        .ok_or_else(|| invalid_relationships("invalid main document target"))?;
    if targets.len() < 2 {
        targets.push((target, relationship_kind));
    }
    Ok(())
}

fn xml_whitespace(bytes: &[u8]) -> bool {
    bytes.iter().all(u8::is_ascii_whitespace)
}

fn declared_main_kind(content_type: &str) -> Option<DocumentKind> {
    let content_type = content_type.to_ascii_lowercase();
    match content_type.as_str() {
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml" => {
            Some(DocumentKind::Docx)
        }
        "application/vnd.ms-word.document.macroenabled.main+xml" => Some(DocumentKind::Docm),
        "application/vnd.openxmlformats-officedocument.wordprocessingml.template.main+xml" => {
            Some(DocumentKind::Dotx)
        }
        "application/vnd.ms-word.template.macroenabledtemplate.main+xml" => {
            Some(DocumentKind::Dotm)
        }
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml" => {
            Some(DocumentKind::Xlsx)
        }
        "application/vnd.ms-excel.sheet.macroenabled.main+xml" => Some(DocumentKind::Xlsm),
        "application/vnd.openxmlformats-officedocument.spreadsheetml.template.main+xml" => {
            Some(DocumentKind::Xltx)
        }
        "application/vnd.ms-excel.template.macroenabled.main+xml" => Some(DocumentKind::Xltm),
        "application/vnd.ms-excel.addin.macroenabled.main+xml" => Some(DocumentKind::Xlam),
        "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml" => {
            Some(DocumentKind::Pptx)
        }
        "application/vnd.ms-powerpoint.presentation.macroenabled.main+xml" => {
            Some(DocumentKind::Pptm)
        }
        "application/vnd.openxmlformats-officedocument.presentationml.template.main+xml" => {
            Some(DocumentKind::Potx)
        }
        "application/vnd.ms-powerpoint.template.macroenabled.main+xml" => Some(DocumentKind::Potm),
        "application/vnd.openxmlformats-officedocument.presentationml.slideshow.main+xml" => {
            Some(DocumentKind::Ppsx)
        }
        "application/vnd.ms-powerpoint.slideshow.macroenabled.main+xml" => Some(DocumentKind::Ppsm),
        "application/vnd.ms-powerpoint.addin.macroenabled.main+xml" => Some(DocumentKind::Ppam),
        "application/vnd.ms-visio.drawing.main+xml" => Some(DocumentKind::Vsdx),
        "application/vnd.ms-visio.drawing.macroenabled.main+xml" => Some(DocumentKind::Vsdm),
        "application/vnd.ms-visio.stencil.main+xml" => Some(DocumentKind::Vssx),
        "application/vnd.ms-visio.stencil.macroenabled.main+xml" => Some(DocumentKind::Vssm),
        "application/vnd.ms-visio.template.main+xml" => Some(DocumentKind::Vstx),
        "application/vnd.ms-visio.template.macroenabled.main+xml" => Some(DocumentKind::Vstm),
        _ => None,
    }
}

#[cfg(test)]
fn detect_format(parts: &[(String, Vec<u8>)]) -> Result<DocumentKind, String> {
    let mut xml_budget = XmlBudget::default();
    detect_format_with_budget(parts, &mut xml_budget).map(|(kind, _)| kind)
}

fn detect_format_with_budget(
    parts: &[(String, Vec<u8>)],
    xml_budget: &mut XmlBudget,
) -> Result<(DocumentKind, Option<ContentTypes>), String> {
    match detect_package_kind_with_budget(parts, xml_budget) {
        Ok((kind, content_types)) => return Ok((kind, Some(content_types))),
        Err(error @ DocumentKindError::InvalidContentTypes(_))
        | Err(error @ DocumentKindError::InvalidPackageRelationships(_))
        | Err(error @ DocumentKindError::ConflictingMainDocumentRelationships(_)) => {
            return Err(error.to_string());
        }
        Err(_) => {}
    }
    let kind = if has_part(parts, "word/document.xml") {
        DocumentKind::Docx
    } else if has_part(parts, "xl/workbook.xml") {
        DocumentKind::Xlsx
    } else if has_part(parts, "ppt/presentation.xml") {
        DocumentKind::Pptx
    } else {
        return Err("could not detect DOCX, XLSX, PPTX, or VSDX package".to_owned());
    };
    let content_types = match parse_content_types_with_budget(parts, xml_budget) {
        Ok(content_types) => Some(content_types),
        Err(DocumentKindError::MissingContentTypes) => None,
        Err(error) => return Err(error.to_string()),
    };
    Ok((kind, content_types))
}

fn has_part(parts: &[(String, Vec<u8>)], expected: &str) -> bool {
    parts
        .iter()
        .any(|(path, _)| path.eq_ignore_ascii_case(expected))
}

fn sanitize_content_types(
    xml: &[u8],
    removed: &HashSet<String>,
    path: &str,
    xml_budget: &mut XmlBudget,
) -> Result<Vec<u8>, String> {
    let needs_validation = xml_budget.needs_validation(path);
    let mut reader = Reader::from_reader(xml);
    let mut writer = Writer::new(Vec::with_capacity(xml.len()));
    let mut skip_depth = 0_usize;
    let mut depth = 0_usize;
    loop {
        let event = reader
            .read_event()
            .map_err(|error| format!("invalid XML in {path}: {error}"))?;
        if needs_validation {
            validate_xml_event(&event, path, xml_budget, &mut depth)?;
        }
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
            .is_some_and(|path| removed.contains(normalize_part_name(path).as_ref()))
}

fn rewrite_content_type(
    start: BytesStart<'_>,
    attributes: Vec<(String, String)>,
) -> BytesStart<'static> {
    let mut output = start.into_owned();
    output.clear_attributes();
    for (key, value) in attributes {
        let value = if key == "ContentType" {
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
        "application/vnd.ms-word.template.macroenabledtemplate.main+xml" => {
            "application/vnd.openxmlformats-officedocument.wordprocessingml.template.main+xml"
        }
        "application/vnd.ms-excel.sheet.macroenabled.main+xml"
        | "application/vnd.ms-excel.addin.macroenabled.main+xml" => {
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"
        }
        "application/vnd.ms-excel.template.macroenabled.main+xml" => {
            "application/vnd.openxmlformats-officedocument.spreadsheetml.template.main+xml"
        }
        "application/vnd.ms-powerpoint.presentation.macroenabled.main+xml"
        | "application/vnd.ms-powerpoint.addin.macroenabled.main+xml" => {
            "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"
        }
        "application/vnd.ms-powerpoint.template.macroenabled.main+xml" => {
            "application/vnd.openxmlformats-officedocument.presentationml.template.main+xml"
        }
        "application/vnd.ms-powerpoint.slideshow.macroenabled.main+xml" => {
            "application/vnd.openxmlformats-officedocument.presentationml.slideshow.main+xml"
        }
        _ => content_type,
    }
}

fn sanitize_relationships(
    xml: &[u8],
    removed: &HashSet<String>,
    content_types: Option<&ContentTypes>,
    path: &str,
    xml_budget: &mut XmlBudget,
) -> Result<Vec<u8>, String> {
    let needs_validation = xml_budget.needs_validation(path);
    let mut reader = Reader::from_reader(xml);
    let mut writer = Writer::new(Vec::with_capacity(xml.len()));
    let mut skip_depth = 0_usize;
    let mut depth = 0_usize;
    loop {
        let event = reader
            .read_event()
            .map_err(|error| format!("invalid XML in {path}: {error}"))?;
        if needs_validation {
            validate_xml_event(&event, path, xml_budget, &mut depth)?;
        }
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
                    && remove_relationship(
                        &attributes(&reader, &start, path)?,
                        removed,
                        content_types,
                        path,
                    );
                if remove {
                    skip_depth = 1;
                } else {
                    write_start(&mut writer, start.into_owned(), false, path)?;
                }
            }
            Event::Empty(start) => {
                let remove = start.name().local_name().as_ref() == b"Relationship"
                    && remove_relationship(
                        &attributes(&reader, &start, path)?,
                        removed,
                        content_types,
                        path,
                    );
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
    content_types: Option<&ContentTypes>,
    relationship_path: &str,
) -> bool {
    let target = attribute_value(attributes, "Target").unwrap_or_default();
    let target_mode = attribute_value(attributes, "TargetMode").unwrap_or_default();
    let relationship_type = attribute_value(attributes, "Type").unwrap_or_default();
    target_mode.trim().eq_ignore_ascii_case("External")
        || external_target(target)
        || encoded_path_delimiter(target)
        || dangerous_relationship_type(relationship_type)
        || resolve_relationship_target(relationship_path, target).is_some_and(|target| {
            removed.contains(&target) || dangerous_part(&target, content_types)
        })
}

fn dangerous_part(path: &str, content_types: Option<&ContentTypes>) -> bool {
    content_types
        .and_then(|types| types.content_type_for(path))
        .map_or_else(|| dangerous_path(path), dangerous_content_type)
}

fn sanitize_workbook_sheets(
    parts: &mut [(String, Vec<u8>)],
    removed: &mut HashSet<String>,
    content_types: Option<&ContentTypes>,
    xml_budget: &mut XmlBudget,
) -> Result<(), String> {
    let part_names = parts
        .iter()
        .map(|(path, _)| normalize_part_name(path).into_owned())
        .collect();
    let workbook_path = main_document_part(parts, &part_names, xml_budget)
        .map(|main| main.part_name)
        .unwrap_or_else(|_| "xl/workbook.xml".to_owned());
    let (directory, filename) = workbook_path
        .rsplit_once('/')
        .unwrap_or(("", &workbook_path));
    let relationships_path = if directory.is_empty() {
        format!("_rels/{filename}.rels")
    } else {
        format!("{directory}/_rels/{filename}.rels")
    };
    let Some((path, bytes)) = parts
        .iter()
        .find(|(path, _)| normalize_part_name(path).as_ref() == relationships_path)
    else {
        return Ok(());
    };
    validate_xml_limits(bytes, path, xml_budget)?;
    let mut reader = NsReader::from_reader(bytes.as_slice());
    let mut kept_by_id = BTreeMap::new();
    let mut removed_macro_sheets = Vec::new();
    let mut calc_chains = Vec::new();
    let mut root_seen = false;
    let mut root_valid = false;
    let mut depth = 0;
    loop {
        let (namespace, event) = read_decoded_namespace_event(&mut reader)
            .map_err(|error| format!("invalid XML in {path}: {error}"))?;
        match &event {
            Event::Start(start) | Event::Empty(start) => {
                if depth == 0 {
                    root_valid = !root_seen
                        && start.local_name().as_ref() == b"Relationships"
                        && is_relationships_namespace(&namespace);
                    root_seen = true;
                }
                if depth == 1
                    && start.local_name().as_ref() == b"Relationship"
                    && is_relationships_namespace(&namespace)
                {
                    let values = attributes(&reader, start, path)?;
                    let is_removed = remove_relationship(&values, removed, content_types, path);
                    let relationship_type = attribute_value(&values, "Type")
                        .unwrap_or_default()
                        .to_ascii_lowercase();
                    let target = attribute_value(&values, "Target")
                        .and_then(|target| resolve_relationship_target(path, target));
                    let macro_sheet = relationship_type.contains("macrosheet")
                        || target.as_deref().is_some_and(|target| {
                            target.contains("/macrosheets/")
                                || content_types
                                    .and_then(|types| types.content_type_for(target))
                                    .is_some_and(|value| {
                                        value.to_ascii_lowercase().contains("macrosheet")
                                    })
                        });
                    if !is_removed && relationship_type.ends_with("/calcchain") {
                        calc_chains.extend(target);
                    }
                    // Id-less relationships cannot be referenced by a sheet.
                    if let Some(id) = attribute_value(&values, "Id") {
                        let kept = kept_by_id.entry(id.to_owned()).or_insert(0_usize);
                        if !is_removed {
                            *kept += 1;
                        } else if macro_sheet {
                            removed_macro_sheets.push(id.to_owned());
                        }
                    }
                }
                if matches!(event, Event::Start(_)) {
                    depth += 1;
                }
            }
            Event::End(_) => depth -= 1,
            Event::Eof => break,
            _ => {}
        }
    }
    let relationships: BTreeMap<_, _> = kept_by_id
        .into_iter()
        .map(|(id, kept)| {
            let relationship = match kept {
                0 => SheetRelationship::Removed,
                1 => SheetRelationship::Kept,
                _ => SheetRelationship::Ambiguous,
            };
            (id, relationship)
        })
        .collect();
    if !removed_macro_sheets
        .iter()
        .any(|id| matches!(relationships.get(id), Some(SheetRelationship::Removed)))
    {
        return Ok(());
    }
    let workbook = parts
        .iter()
        .position(|(path, _)| normalize_part_name(path).as_ref() == workbook_path)
        .ok_or_else(|| format!("missing workbook part: {workbook_path}"))?;
    let (workbook_name, xml) = &parts[workbook];
    if !rewrites_sheet_entries(xml, workbook_name, &relationships, xml_budget)? {
        return Ok(());
    }
    let Some(sheets) = workbook_sheet_map(xml, workbook_name, &relationships, xml_budget)? else {
        return Ok(());
    };
    if !root_valid {
        return Err(format!("invalid relationships in {path}"));
    }
    parts[workbook].1 = rewrite_workbook_sheets(xml, workbook_name, &sheets)?;
    // Calc chain entries name sheet ids; spreadsheet apps rebuild a missing chain.
    removed.extend(calc_chains);
    Ok(())
}

enum SheetRelationship {
    Kept,
    Removed,
    Ambiguous,
}

#[derive(Clone, Copy)]
enum WorkbookElement {
    Sheet,
    DefinedName,
    View,
    // Outside its list; only a rewrite rejects it.
    Misplaced,
    Other,
}

fn is_spreadsheet_namespace(namespace: &ResolveResult<'_>) -> bool {
    matches!(namespace, ResolveResult::Bound(namespace)
        if matches!(namespace.0, b"http://schemas.openxmlformats.org/spreadsheetml/2006/main" | b"http://purl.oclc.org/ooxml/spreadsheetml/main"))
}

fn is_markup_compatibility_namespace(namespace: &ResolveResult<'_>) -> bool {
    matches!(namespace, ResolveResult::Bound(namespace)
        if namespace.0 == MARKUP_COMPATIBILITY_NAMESPACE)
}

struct WorkbookReader<'a> {
    reader: NsReader<&'a [u8]>,
    stack: Vec<Vec<u8>>,
    root_seen: bool,
    sheets_seen: bool,
}

impl<'a> WorkbookReader<'a> {
    fn new(xml: &'a [u8]) -> Self {
        Self {
            reader: NsReader::from_reader(xml),
            stack: Vec::new(),
            root_seen: false,
            sheets_seen: false,
        }
    }

    fn read_event(&mut self, path: &str) -> Result<(Event<'a>, WorkbookElement), String> {
        let (namespace, event) = read_decoded_namespace_event(&mut self.reader)
            .map_err(|error| format!("invalid XML in {path}: {error}"))?;
        let mut element = WorkbookElement::Other;
        match &event {
            Event::Start(start) | Event::Empty(start) => {
                let spreadsheet = is_spreadsheet_namespace(&namespace);
                if matches!(&namespace, ResolveResult::Unknown(_)) {
                    return Err(format!("unbound workbook namespace in {path}"));
                }
                let local = start.local_name();
                if self.stack.is_empty() {
                    if self.root_seen || !spreadsheet || local.as_ref() != b"workbook" {
                        return Err(format!("invalid workbook root in {path}"));
                    }
                    self.root_seen = true;
                }
                if spreadsheet {
                    let parent = self.stack.last().map(Vec::as_slice);
                    if self.stack.len() == 1 && local.as_ref() == b"sheets" {
                        if self.sheets_seen {
                            return Err(format!("duplicate workbook sheets in {path}"));
                        }
                        self.sheets_seen = true;
                    }
                    element = match (self.stack.len(), parent, local.as_ref()) {
                        (2, Some(b"sheets"), b"sheet") => WorkbookElement::Sheet,
                        (2, Some(b"definedNames"), b"definedName") => WorkbookElement::DefinedName,
                        (2, Some(b"bookViews"), b"workbookView") => WorkbookElement::View,
                        (_, _, b"sheet") => {
                            return Err(format!("unexpected workbook element in {path}"));
                        }
                        (_, _, b"definedName" | b"workbookView") => WorkbookElement::Misplaced,
                        _ => WorkbookElement::Other,
                    };
                }
                if matches!(event, Event::Start(_)) {
                    self.stack.push(if spreadsheet {
                        local.as_ref().to_vec()
                    } else {
                        Vec::new()
                    });
                }
            }
            Event::End(_) => {
                self.stack.pop();
            }
            Event::Text(text) if self.stack.is_empty() && !xml_whitespace(text.as_ref()) => {
                return Err(format!("text outside workbook root in {path}"));
            }
            Event::CData(_) if self.stack.is_empty() => {
                return Err(format!("CDATA outside workbook root in {path}"));
            }
            Event::Eof if !self.root_seen || !self.sheets_seen || !self.stack.is_empty() => {
                return Err(format!("incomplete workbook in {path}"));
            }
            _ => {}
        }
        Ok((event, element))
    }
}

// Only relationship ids are decoded; other attribute values stay unread.
fn sheet_relationship_ids(
    reader: &NsReader<&[u8]>,
    start: &BytesStart<'_>,
    path: &str,
) -> Result<Vec<String>, String> {
    let mut ids = Vec::new();
    for attribute in start.attributes() {
        let attribute = attribute.map_err(|error| format!("invalid XML in {path}: {error}"))?;
        let (namespace, local) = reader.resolver().resolve_attribute(attribute.key);
        if local.as_ref() == b"id"
            && matches!(namespace, ResolveResult::Bound(namespace) if matches!(namespace.0, b"http://schemas.openxmlformats.org/officeDocument/2006/relationships" | b"http://purl.oclc.org/ooxml/officeDocument/relationships"))
        {
            let id = attribute
                .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
                .map_err(|error| format!("invalid XML in {path}: {error}"))?;
            ids.push(id.into_owned());
        }
    }
    Ok(ids)
}

fn workbook_sheet_id(
    reader: &NsReader<&[u8]>,
    start: &BytesStart<'_>,
    path: &str,
) -> Result<String, String> {
    let mut ids = sheet_relationship_ids(reader, start, path)?.into_iter();
    let id = ids
        .next()
        .ok_or_else(|| format!("sheet missing relationship id in {path}"))?;
    if ids.next().is_some() {
        return Err(format!("duplicate sheet relationship in {path}"));
    }
    Ok(id)
}

// Lenient: any `<sheet>` at any depth or namespace may name a removed id. Sheet-indexed
// entries under markup compatibility keep the workbook as it is.
fn rewrites_sheet_entries(
    xml: &[u8],
    path: &str,
    relationships: &BTreeMap<String, SheetRelationship>,
    xml_budget: &mut XmlBudget,
) -> Result<bool, String> {
    validate_xml_limits(xml, path, xml_budget)?;
    let mut reader = NsReader::from_reader(xml);
    let mut depth = 0_usize;
    let mut compatibility_depth = None;
    let mut names_removed = false;
    loop {
        let (namespace, event) = read_decoded_namespace_event(&mut reader)
            .map_err(|error| format!("invalid XML in {path}: {error}"))?;
        match &event {
            Event::Start(start) | Event::Empty(start) => {
                let spreadsheet = is_spreadsheet_namespace(&namespace);
                let compatibility = is_markup_compatibility_namespace(&namespace);
                let local = start.local_name();
                if compatibility_depth.is_some()
                    && spreadsheet
                    && matches!(local.as_ref(), b"sheet" | b"definedName" | b"workbookView")
                {
                    return Ok(false);
                }
                if !names_removed && local.as_ref() == b"sheet" {
                    names_removed =
                        sheet_relationship_ids(&reader, start, path)?
                            .iter()
                            .any(|id| {
                                matches!(relationships.get(id), Some(SheetRelationship::Removed))
                            });
                }
                if matches!(event, Event::Start(_)) {
                    if compatibility && compatibility_depth.is_none() {
                        compatibility_depth = Some(depth);
                    }
                    depth += 1;
                }
            }
            Event::End(_) => {
                depth -= 1;
                if compatibility_depth == Some(depth) {
                    compatibility_depth = None;
                }
            }
            Event::Eof => return Ok(names_removed),
            _ => {}
        }
    }
}

// Each `<sheet>` mapped to its index after removal; None when no sheet is removed.
fn workbook_sheet_map(
    xml: &[u8],
    path: &str,
    relationships: &BTreeMap<String, SheetRelationship>,
    xml_budget: &mut XmlBudget,
) -> Result<Option<Vec<Option<usize>>>, String> {
    validate_xml_limits(xml, path, xml_budget)?;
    let mut reader = WorkbookReader::new(xml);
    let mut sheets = Vec::new();
    let mut seen_ids = HashSet::new();
    let mut remaining = 0;
    loop {
        let (event, element) = reader.read_event(path)?;
        match event {
            Event::Start(start) | Event::Empty(start)
                if matches!(element, WorkbookElement::Sheet) =>
            {
                let id = workbook_sheet_id(&reader.reader, &start, path)?;
                if !seen_ids.insert(id.clone()) {
                    return Err(format!("duplicate sheet relationship in {path}"));
                }
                let removed = match relationships.get(&id) {
                    Some(SheetRelationship::Kept) => false,
                    Some(SheetRelationship::Removed) => true,
                    Some(SheetRelationship::Ambiguous) => {
                        return Err(format!("ambiguous sheet relationship in {path}"));
                    }
                    None => return Err(format!("unknown sheet relationship in {path}")),
                };
                if removed {
                    sheets.push(None);
                } else {
                    sheets.push(Some(remaining));
                    remaining += 1;
                }
            }
            Event::Eof => break,
            _ => {}
        }
    }
    if remaining == sheets.len() {
        return Ok(None);
    }
    if remaining == 0 {
        return Err(format!("no remaining workbook sheets in {path}"));
    }
    Ok(Some(sheets))
}

fn rewrite_workbook_sheets(
    xml: &[u8],
    path: &str,
    sheets: &[Option<usize>],
) -> Result<Vec<u8>, String> {
    // A view on a removed sheet moves to the next kept one: the count of kept sheets before it.
    let mut kept_before = Vec::with_capacity(sheets.len());
    let mut remaining = 0;
    for sheet in sheets {
        kept_before.push(remaining);
        remaining += usize::from(sheet.is_some());
    }
    let mut reader = WorkbookReader::new(xml);
    let mut writer = Writer::new(Vec::with_capacity(xml.len()));
    let mut sheet_index = 0;
    let mut skip_depth = 0;
    loop {
        let (event, element) = reader.read_event(path)?;
        if skip_depth > 0 {
            match event {
                Event::Start(_) => skip_depth += 1,
                Event::End(_) => skip_depth -= 1,
                _ => {}
            }
            continue;
        }
        let empty = matches!(event, Event::Empty(_));
        match event {
            Event::Start(start) | Event::Empty(start) => {
                let mut values = attributes(&reader.reader, &start, path)?;
                let remove = match element {
                    WorkbookElement::Sheet => {
                        let remove = sheets[sheet_index].is_none();
                        sheet_index += 1;
                        remove
                    }
                    WorkbookElement::DefinedName => {
                        if let Some(index) =
                            values.iter_mut().find(|(key, _)| key == "localSheetId")
                        {
                            let mapped = workbook_sheet_index(&index.1, sheets, path)?;
                            if let Some(mapped) = mapped {
                                index.1 = mapped.to_string();
                                false
                            } else {
                                true
                            }
                        } else {
                            false
                        }
                    }
                    WorkbookElement::View => {
                        for (key, value) in &mut values {
                            if matches!(key.as_str(), "activeTab" | "firstSheet") {
                                let mapped = value
                                    .parse::<usize>()
                                    .ok()
                                    .and_then(|index| kept_before.get(index))
                                    .ok_or_else(|| {
                                        format!("invalid workbook sheet index in {path}")
                                    })?;
                                *value = (*mapped).min(remaining - 1).to_string();
                            }
                        }
                        false
                    }
                    WorkbookElement::Misplaced => {
                        return Err(format!("unexpected workbook element in {path}"));
                    }
                    WorkbookElement::Other => false,
                };
                if remove {
                    if !empty {
                        skip_depth = 1;
                    }
                } else {
                    let mut output = start.into_owned();
                    output.clear_attributes();
                    for (key, value) in values {
                        output.push_attribute((key.as_str(), value.as_str()));
                    }
                    write_start(&mut writer, output, empty, path)?;
                }
            }
            Event::Eof => return Ok(writer.into_inner()),
            other => write(&mut writer, other, path)?,
        }
    }
}

fn workbook_sheet_index(
    value: &str,
    sheets: &[Option<usize>],
    path: &str,
) -> Result<Option<usize>, String> {
    value
        .parse::<usize>()
        .ok()
        .and_then(|index| sheets.get(index).copied())
        .ok_or_else(|| format!("invalid workbook sheet index in {path}"))
}

fn neutralize_fields(
    xml: &[u8],
    path: &str,
    xml_budget: &mut XmlBudget,
) -> Result<Vec<u8>, String> {
    let needs_validation = xml_budget.needs_validation(path);
    let mut reader = Reader::from_reader(xml);
    reader.config_mut().trim_text(false);
    let mut writer = Writer::new(Vec::with_capacity(xml.len()));
    let mut stack = Vec::new();
    let mut skip_depth = 0_usize;
    let mut depth = 0_usize;
    loop {
        let event = reader
            .read_event()
            .map_err(|error| format!("invalid XML in {path}: {error}"))?;
        if needs_validation {
            validate_xml_event(&event, path, xml_budget, &mut depth)?;
        }
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
        .find(|(name, _)| name == expected)
        .map(|(_, value)| value.as_str())
}

fn attribute_local(name: &str) -> &str {
    name.rsplit_once(':').map_or(name, |(_, local)| local)
}

fn dangerous_path(path: &str) -> bool {
    let path: Cow<'_, str> =
        if path.contains('\\') || path.bytes().any(|byte| byte.is_ascii_uppercase()) {
            Cow::Owned(path.replace('\\', "/").to_ascii_lowercase())
        } else {
            Cow::Borrowed(path)
        };
    path.ends_with("vbaproject.bin")
        || path.ends_with("vbaprojectsignature.bin")
        || path.ends_with("vbaprojectsignatureagile.bin")
        || path.ends_with("vbaprojectsignaturev3.bin")
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
    let target = target.trim();
    let mut prefix = target.bytes();
    let network_target = prefix
        .next()
        .is_some_and(|byte| matches!(byte, b'/' | b'\\'))
        && prefix
            .next()
            .is_some_and(|byte| matches!(byte, b'/' | b'\\'));
    network_target
        || target.split_once(':').is_some_and(|(scheme, _)| {
            let mut bytes = scheme.bytes();
            bytes.next().is_some_and(|byte| byte.is_ascii_alphabetic())
                && bytes
                    .all(|byte| byte.is_ascii_alphanumeric() || matches!(byte, b'+' | b'-' | b'.'))
        })
}

// OPC part names never percent-encode a separator or a dot, and readers disagree on where such
// a target points.
fn encoded_path_delimiter(target: &str) -> bool {
    target.as_bytes().windows(3).any(|window| {
        window[0] == b'%'
            && matches!(
                (window[1], window[2].to_ascii_uppercase()),
                (b'2', b'E' | b'F') | (b'5', b'C')
            )
    })
}

#[derive(Default)]
struct XmlBudget {
    events: usize,
    attribute_bytes: usize,
    validated_parts: HashSet<String>,
}

impl XmlBudget {
    fn needs_validation(&self, path: &str) -> bool {
        !self
            .validated_parts
            .contains(normalize_part_name(path).as_ref())
    }
}

fn validate_xml_limits(xml: &[u8], path: &str, budget: &mut XmlBudget) -> Result<(), String> {
    if !budget.needs_validation(path) {
        return Ok(());
    }
    let mut reader = Reader::from_reader(xml);
    let mut depth = 0_usize;
    loop {
        let event = reader
            .read_event()
            .map_err(|error| format!("invalid XML in {path}: {error}"))?;
        let done = matches!(event, Event::Eof);
        validate_xml_event(&event, path, budget, &mut depth)?;
        if done {
            return Ok(());
        }
    }
}

fn validate_xml_event(
    event: &Event<'_>,
    path: &str,
    budget: &mut XmlBudget,
    depth: &mut usize,
) -> Result<(), String> {
    budget.events = budget
        .events
        .checked_add(1)
        .ok_or_else(|| format!("package XML event limit exceeded in {path}"))?;
    if budget.events > MAX_PACKAGE_XML_EVENTS {
        return Err(format!("package XML event limit exceeded in {path}"));
    }
    match event {
        Event::Start(start) => {
            validate_xml_attributes(start, path, budget)?;
            *depth += 1;
            if *depth > MAX_XML_DEPTH {
                return Err(format!("XML depth limit exceeded in {path}"));
            }
        }
        Event::Empty(start) => {
            validate_xml_attributes(start, path, budget)?;
            if *depth >= MAX_XML_DEPTH {
                return Err(format!("XML depth limit exceeded in {path}"));
            }
        }
        Event::End(_) => {
            if *depth == 0 {
                return Err(format!("unexpected closing element in {path}"));
            }
            *depth -= 1;
        }
        Event::DocType(_) => return Err(format!("DTD is forbidden in {path}")),
        Event::Eof => {
            if *depth != 0 {
                return Err(format!("unexpected EOF in {path}"));
            }
            budget
                .validated_parts
                .insert(normalize_part_name(path).into_owned());
        }
        _ => {}
    }
    Ok(())
}

fn validate_xml_attributes(
    start: &BytesStart<'_>,
    path: &str,
    budget: &mut XmlBudget,
) -> Result<(), String> {
    let mut bytes = 0_usize;
    for (index, attribute) in start.attributes().enumerate() {
        if index >= MAX_XML_ATTRIBUTES_PER_ELEMENT {
            return Err(format!("XML attribute count limit exceeded in {path}"));
        }
        let attribute = attribute.map_err(|error| format!("invalid XML in {path}: {error}"))?;
        bytes = bytes
            .checked_add(attribute.key.as_ref().len() + attribute.value.as_ref().len())
            .ok_or_else(|| format!("XML attribute byte limit exceeded in {path}"))?;
        if bytes > MAX_XML_ATTRIBUTE_BYTES_PER_ELEMENT {
            return Err(format!("XML attribute byte limit exceeded in {path}"));
        }
        budget.attribute_bytes = budget
            .attribute_bytes
            .checked_add(attribute.key.as_ref().len() + attribute.value.as_ref().len())
            .ok_or_else(|| format!("package XML attribute byte limit exceeded in {path}"))?;
        if budget.attribute_bytes > MAX_PACKAGE_XML_ATTRIBUTE_BYTES {
            return Err(format!(
                "package XML attribute byte limit exceeded in {path}"
            ));
        }
    }
    Ok(())
}

fn resolve_relationship_target(relationship_path: &str, target: &str) -> Option<String> {
    resolve_target(relationship_path, target).map(|(target, _)| target)
}

// RFC 3986 dot segments: `..` at the package root stays there; the flag reports such a climb.
fn resolve_target(relationship_path: &str, target: &str) -> Option<(String, bool)> {
    if target.is_empty() || external_target(target) {
        return None;
    }
    let clean_target = target
        .split(['?', '#'])
        .next()
        .unwrap_or_default()
        .replace('\\', "/");
    let mut segments = if clean_target.starts_with('/') {
        Vec::new()
    } else {
        relationship_base(relationship_path)
    };
    let mut climbed = false;
    for segment in clean_target
        .split('/')
        .filter(|segment| !segment.is_empty() && *segment != ".")
    {
        if segment == ".." {
            climbed |= segments.pop().is_none();
        } else {
            segments.push(segment.to_owned());
        }
    }
    Some((segments.join("/").to_ascii_lowercase(), climbed))
}

// `<dir>/_rels/<name>.rels` describes `<dir>/<name>`; relative targets resolve against `<dir>`.
fn relationship_base(relationship_path: &str) -> Vec<String> {
    let mut segments: Vec<String> = relationship_path
        .split(['/', '\\'])
        .filter(|segment| !segment.is_empty())
        .map(str::to_owned)
        .collect();
    segments.pop();
    if segments
        .last()
        .is_some_and(|segment| segment.eq_ignore_ascii_case("_rels"))
    {
        segments.pop();
    }
    segments
}

fn normalize_part_name(path: &str) -> Cow<'_, str> {
    if !path.starts_with('/')
        && !path.contains('\\')
        && !path.bytes().any(|byte| byte.is_ascii_uppercase())
    {
        return Cow::Borrowed(path);
    }
    Cow::Owned(
        path.replace('\\', "/")
            .trim_start_matches('/')
            .to_ascii_lowercase(),
    )
}

fn is_xml_part(path: &str) -> bool {
    path.ends_with(".xml") || path.ends_with(".vml")
}

fn is_xml_content_type(content_type: &str) -> bool {
    let content_type = content_type
        .split_once(';')
        .map_or(content_type, |(content_type, _)| content_type)
        .trim();
    content_type.eq_ignore_ascii_case("application/xml")
        || content_type.eq_ignore_ascii_case("text/xml")
        || content_type.to_ascii_lowercase().ends_with("+xml")
}

fn is_relationships_content_type(content_type: &str) -> bool {
    content_type
        .trim()
        .eq_ignore_ascii_case("application/vnd.openxmlformats-package.relationships+xml")
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
    use crate::rezip_parts;

    fn content_types(part_name: &str, content_type: &str) -> (String, Vec<u8>) {
        (
            "[Content_Types].xml".to_owned(),
            format!(
                "<Types xmlns='{namespace}'><Override PartName='/{part_name}' ContentType='{content_type}'/></Types>",
                namespace = String::from_utf8_lossy(CONTENT_TYPES_NAMESPACE)
            )
            .into_bytes(),
        )
    }

    fn main_relationship(part_name: &str, relationship_type: &str) -> (String, Vec<u8>) {
        (
            "_rels/.rels".to_owned(),
            format!(
                "<Relationships xmlns='{namespace}'><Relationship Id='rId1' Type='{relationship_type}' Target='{}'/></Relationships>",
                part_name.trim_start_matches('/'),
                namespace = String::from_utf8_lossy(RELATIONSHIPS_NAMESPACE)
            )
            .into_bytes(),
        )
    }

    fn package_parts(
        part_name: &str,
        content_type: &str,
        main_xml: &[u8],
    ) -> Vec<(String, Vec<u8>)> {
        let relationship_type = if content_type.contains("visio") {
            VISIO_DOCUMENT_RELATIONSHIP
        } else {
            OFFICE_DOCUMENT_RELATIONSHIP
        };
        vec![
            content_types(part_name, content_type),
            main_relationship(part_name, relationship_type),
            (part_name.to_owned(), main_xml.to_vec()),
        ]
    }

    #[test]
    fn strips_external_macro_and_embedded_attack_vectors() {
        let package = rezip_parts(&[
            (
                "[Content_Types].xml".to_owned(),
                br#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="bin" ContentType="application/vnd.ms-office.vbaProject"/><Override PartName="/word/document.xml" ContentType="application/vnd.ms-word.document.macroEnabled.main+xml"/><Override PartName="/word/vbaProject.bin" ContentType="application/vnd.ms-office.vbaProject"/><Override PartName="/word/embeddings/object1.bin" ContentType="application/vnd.openxmlformats-officedocument.oleObject"/></Types>"#.to_vec(),
            ),
            main_relationship("word/document.xml", OFFICE_DOCUMENT_RELATIONSHIP),
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
                "_rels/.rels",
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
        let package = rezip_parts(&package_parts(
            "xl/workbook.xml",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
            b"<workbook/>",
        ))
        .unwrap();
        assert!(sanitize_package_for_format(&package, "docx").is_err());
    }

    const TEMPLATE_AND_ADDIN_KINDS: [[&str; 4]; 10] = [
        [
            "dotx",
            "application/vnd.openxmlformats-officedocument.wordprocessingml.template.main+xml",
            "dotx",
            "application/vnd.openxmlformats-officedocument.wordprocessingml.template.main+xml",
        ],
        [
            "dotm",
            "application/vnd.ms-word.template.macroEnabledTemplate.main+xml",
            "dotx",
            "application/vnd.openxmlformats-officedocument.wordprocessingml.template.main+xml",
        ],
        [
            "xltx",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.template.main+xml",
            "xltx",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.template.main+xml",
        ],
        [
            "xltm",
            "application/vnd.ms-excel.template.macroEnabled.main+xml",
            "xltx",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.template.main+xml",
        ],
        [
            "xlam",
            "application/vnd.ms-excel.addin.macroEnabled.main+xml",
            "xlsx",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
        ],
        [
            "potx",
            "application/vnd.openxmlformats-officedocument.presentationml.template.main+xml",
            "potx",
            "application/vnd.openxmlformats-officedocument.presentationml.template.main+xml",
        ],
        [
            "potm",
            "application/vnd.ms-powerpoint.template.macroEnabled.main+xml",
            "potx",
            "application/vnd.openxmlformats-officedocument.presentationml.template.main+xml",
        ],
        [
            "ppsx",
            "application/vnd.openxmlformats-officedocument.presentationml.slideshow.main+xml",
            "ppsx",
            "application/vnd.openxmlformats-officedocument.presentationml.slideshow.main+xml",
        ],
        [
            "ppsm",
            "application/vnd.ms-powerpoint.slideshow.macroEnabled.main+xml",
            "ppsx",
            "application/vnd.openxmlformats-officedocument.presentationml.slideshow.main+xml",
        ],
        [
            "ppam",
            "application/vnd.ms-powerpoint.addin.macroEnabled.main+xml",
            "pptx",
            "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml",
        ],
    ];

    fn office_main_part(format: &str) -> (&str, &str) {
        if format.starts_with('d') {
            ("word/document.xml", "docx")
        } else if format.starts_with('x') {
            ("xl/workbook.xml", "xlsx")
        } else {
            ("ppt/presentation.xml", "pptx")
        }
    }

    fn assert_all_parts_declared(parts: &[(String, Vec<u8>)]) {
        let types = parse_content_types(parts).unwrap();
        for (path, _) in parts {
            if path != "[Content_Types].xml" {
                assert!(types.content_type_for(path).is_some(), "undeclared {path}");
            }
        }
    }

    #[test]
    fn detects_template_slideshow_and_addin_kinds_case_insensitively() {
        for relationship_type in [
            OFFICE_DOCUMENT_RELATIONSHIP,
            STRICT_OFFICE_DOCUMENT_RELATIONSHIP,
        ] {
            for uppercase in [false, true] {
                let detected = TEMPLATE_AND_ADDIN_KINDS.map(|[format, content_type, _, _]| {
                    let (part_name, _) = office_main_part(format);
                    let content_type = if uppercase {
                        content_type.to_ascii_uppercase()
                    } else {
                        content_type.to_owned()
                    };
                    let mut parts = package_parts(part_name, &content_type, b"<root/>");
                    parts[1] = main_relationship(part_name, relationship_type);
                    detect_package_kind(&parts).map(DocumentKind::format)
                });
                assert_eq!(
                    detected,
                    TEMPLATE_AND_ADDIN_KINDS.map(|[format, _, _, _]| Ok(format))
                );
            }
        }
    }

    #[test]
    fn sanitizes_macro_templates_slideshows_and_addins() {
        for [format, content_type, sanitized_format, sanitized_type] in TEMPLATE_AND_ADDIN_KINDS {
            if format == sanitized_format {
                continue;
            }
            let (part_name, _) = office_main_part(format);
            let (directory, main_file) = part_name.rsplit_once('/').unwrap();
            for content_type in [content_type.to_owned(), content_type.to_ascii_uppercase()] {
                let mut parts = package_parts(part_name, &content_type, b"<root/>");
                parts[0].1 = String::from_utf8(parts[0].1.clone())
                    .unwrap()
                    .replace(
                        "</Types>",
                        &format!(
                            "<Default Extension='rels' ContentType='application/vnd.openxmlformats-package.relationships+xml'/><Override PartName='/{directory}/vbaProject.bin' ContentType='application/vnd.ms-office.vbaProject'/></Types>"
                        ),
                    )
                    .into_bytes();
                parts.push((
                    format!("{directory}/vbaProject.bin"),
                    b"macro payload".to_vec(),
                ));
                parts.push((
                    format!("{directory}/_rels/{main_file}.rels"),
                    br#"<Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'><Relationship Id='macro' Type='http://schemas.microsoft.com/office/2006/relationships/vbaProject' Target='vbaProject.bin'/></Relationships>"#.to_vec(),
                ));
                let package = rezip_parts(&parts).unwrap();
                let sanitized = sanitize_package(&package).unwrap();
                let sanitized_for_format =
                    sanitize_package_for_format(&package, sanitized_format).unwrap();
                assert_eq!(
                    unzip_parts(&sanitized_for_format).unwrap(),
                    unzip_parts(&sanitized).unwrap()
                );
                let output = unzip_parts(&sanitized).unwrap();
                assert!(
                    !has_part(&output, &format!("{directory}/vbaProject.bin")),
                    "{format}"
                );
                let all = output
                    .iter()
                    .map(|(_, bytes)| String::from_utf8_lossy(bytes).to_ascii_lowercase())
                    .collect::<String>();
                assert!(!all.contains("vbaproject"), "{format}");
                assert_eq!(
                    parse_content_types(&output)
                        .unwrap()
                        .content_type_for(part_name),
                    Some(sanitized_type),
                    "{format}"
                );
                assert_eq!(
                    detect_package_kind(&output).unwrap().format(),
                    sanitized_format
                );
                assert_all_parts_declared(&output);
            }
        }
    }

    #[test]
    fn validates_sanitized_template_slideshow_and_addin_formats() {
        for [format, content_type, sanitized_format, _] in TEMPLATE_AND_ADDIN_KINDS {
            let (part_name, base_format) = office_main_part(format);
            let package = rezip_parts(&package_parts(part_name, content_type, b"<root/>")).unwrap();
            assert!(
                sanitize_package_for_format(&package, sanitized_format).is_ok(),
                "{format}"
            );
            if sanitized_format == base_format {
                let template_format = if base_format == "xlsx" {
                    "xltx"
                } else {
                    "potx"
                };
                assert!(
                    sanitize_package_for_format(&package, template_format).is_err(),
                    "{format}"
                );
            } else {
                assert!(
                    sanitize_package_for_format(&package, base_format).is_err(),
                    "{format}"
                );
            }
            if format != sanitized_format {
                assert!(
                    sanitize_package_for_format(&package, format).is_err(),
                    "{format}"
                );
            }
        }
    }

    #[test]
    fn preserves_macro_free_templates_and_slideshows() {
        for [format, content_type, sanitized_format, _] in TEMPLATE_AND_ADDIN_KINDS {
            if format != sanitized_format {
                continue;
            }
            let (part_name, _) = office_main_part(format);
            let parts = package_parts(part_name, content_type, b"<root/>");
            let package = rezip_parts(&parts).unwrap();
            for sanitized in [
                sanitize_package(&package).unwrap(),
                sanitize_package_for_format(&package, format).unwrap(),
            ] {
                let output = unzip_parts(&sanitized).unwrap();
                assert_eq!(output.len(), parts.len());
                assert_eq!(
                    parse_content_types(&output)
                        .unwrap()
                        .content_type_for(part_name),
                    Some(content_type)
                );
                assert_eq!(detect_package_kind(&output).unwrap().format(), format);
                assert_eq!(
                    output.iter().find(|(path, _)| path == part_name).unwrap().1,
                    b"<root/>"
                );
            }
        }
    }

    fn default_payload_parts(content_type: &str, part_name: &str) -> Vec<(String, Vec<u8>)> {
        let mut parts = package_parts(
            "xl/workbook.xml",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
            b"<workbook/>",
        );
        parts[0].1 = String::from_utf8(parts[0].1.clone())
            .unwrap()
            .replace(
                "</Types>",
                &format!(
                    "<Default Extension='BIN' ContentType='{content_type}'/><Default Extension='rels' ContentType='application/vnd.openxmlformats-package.relationships+xml'/></Types>"
                ),
            )
            .into_bytes();
        parts.push((part_name.to_owned(), b"dangerous payload".to_vec()));
        parts.push((
            "xl/_rels/workbook.xml.rels".to_owned(),
            format!(
                "<Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'><Relationship Id='payload' Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/printerSettings' Target='/{}'/></Relationships>",
                part_name
            )
            .into_bytes(),
        ));
        parts
    }

    #[test]
    fn strips_vba_projects_declared_only_by_default() {
        for part_name in ["xl/vbaProject.bin", "xl/custom/payload.BIN"] {
            let parts = default_payload_parts("application/vnd.ms-office.vbaProject", part_name);
            let package = rezip_parts(&parts).unwrap();
            let output =
                unzip_parts(&sanitize_package_for_format(&package, "xlsx").unwrap()).unwrap();
            assert!(!has_part(&output, part_name), "{part_name}");
            let all = output
                .iter()
                .map(|(_, bytes)| String::from_utf8_lossy(bytes))
                .collect::<String>();
            assert!(!all.contains("vbaProject"));
            assert!(!all.contains(part_name));
            assert!(!all.contains("Id=\"payload\""));
            assert_all_parts_declared(&output);
        }
    }

    #[test]
    fn strips_signatures_declared_only_by_default() {
        for signature in [
            "vbaProjectSignature",
            "vbaProjectSignatureAgile",
            "vbaProjectSignatureV3",
        ] {
            for part_name in [
                format!("xl/{signature}.bin"),
                "xl/custom/signature.BIN".to_owned(),
            ] {
                let content_type = format!("application/vnd.ms-office.{signature}");
                let parts = default_payload_parts(&content_type, &part_name);
                let package = rezip_parts(&parts).unwrap();
                let output =
                    unzip_parts(&sanitize_package_for_format(&package, "xlsx").unwrap()).unwrap();
                assert!(!has_part(&output, &part_name), "{signature}: {part_name}");
                let all = output
                    .iter()
                    .map(|(_, bytes)| String::from_utf8_lossy(bytes))
                    .collect::<String>();
                assert!(!all.contains(signature));
                assert!(!all.contains(&part_name));
                assert!(!all.contains("Id=\"payload\""));
                assert_all_parts_declared(&output);
            }
        }
    }

    #[test]
    fn strips_signature_paths_without_content_types() {
        for signature in [
            "vbaProjectSignature",
            "vbaProjectSignatureAgile",
            "vbaProjectSignatureV3",
        ] {
            let part_name = format!("xl/{signature}.bin");
            let mut parts = default_payload_parts("application/octet-stream", &part_name);
            parts.remove(0);
            let package = rezip_parts(&parts).unwrap();
            let output = unzip_parts(&sanitize_package(&package).unwrap()).unwrap();
            assert!(!has_part(&output, &part_name), "{signature}");
            assert!(
                output
                    .iter()
                    .all(|(_, bytes)| !String::from_utf8_lossy(bytes).contains(signature))
            );
        }
    }

    #[test]
    fn preserves_safe_binary_override_next_to_dangerous_default() {
        let mut parts = default_payload_parts(
            "application/vnd.ms-office.vbaProject",
            "xl/custom/payload.bin",
        );
        let safe_part = "xl/printerSettings/printerSettings1.bin";
        let safe_type =
            "application/vnd.openxmlformats-officedocument.spreadsheetml.printerSettings";
        parts[0].1 = String::from_utf8(parts[0].1.clone())
            .unwrap()
            .replace(
                "</Types>",
                &format!("<Override PartName='/{safe_part}' ContentType='{safe_type}'/></Types>"),
            )
            .into_bytes();
        parts.push((safe_part.to_owned(), b"printer settings".to_vec()));
        parts[4].1 = String::from_utf8(parts[4].1.clone())
            .unwrap()
            .replace("</Relationships>", "<Relationship Id='safe' Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/printerSettings' Target='printerSettings/printerSettings1.bin'/></Relationships>")
            .into_bytes();
        let package = rezip_parts(&parts).unwrap();
        let output = unzip_parts(&sanitize_package_for_format(&package, "xlsx").unwrap()).unwrap();
        assert_eq!(
            output.iter().find(|(path, _)| path == safe_part).unwrap().1,
            b"printer settings"
        );
        assert_eq!(
            parse_content_types(&output)
                .unwrap()
                .content_type_for(safe_part),
            Some(safe_type)
        );
        let relationships = output
            .iter()
            .find(|(path, _)| path == "xl/_rels/workbook.xml.rels")
            .unwrap();
        assert!(String::from_utf8_lossy(&relationships.1).contains("printerSettings1.bin"));
        assert!(!has_part(&output, "xl/custom/payload.bin"));
        assert_all_parts_declared(&output);
    }

    #[test]
    fn leaves_no_parts_undeclared_after_removing_dangerous_defaults() {
        let parts = default_payload_parts(
            "application/vnd.ms-office.vbaProjectSignature",
            "xl/custom/payload.bin",
        );
        let package = rezip_parts(&parts).unwrap();
        let output = unzip_parts(&sanitize_package(&package).unwrap()).unwrap();
        assert_all_parts_declared(&output);
    }

    #[test]
    fn regression_preserves_manifest_with_vba_xml_default() {
        let parts = vec![
            (
                "[Content_Types].xml".to_owned(),
                br#"<Types xmlns='http://schemas.openxmlformats.org/package/2006/content-types'><Default Extension='xml' ContentType='application/vnd.ms-word.vbaData+xml'/><Default Extension='rels' ContentType='application/vnd.openxmlformats-package.relationships+xml'/><Override PartName='/word/document.xml' ContentType='application/vnd.ms-word.document.macroEnabled.main+xml'/><Override PartName='/word/settings.xml' ContentType='application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml'/><Override PartName='/word/custom/data.xml' ContentType='application/vnd.ms-word.vbaData+xml'/></Types>"#.to_vec(),
            ),
            main_relationship("word/document.xml", OFFICE_DOCUMENT_RELATIONSHIP),
            ("word/document.xml".to_owned(), b"<document/>".to_vec()),
            ("word/settings.xml".to_owned(), b"<settings/>".to_vec()),
            ("word/custom/data.xml".to_owned(), b"<vbaData/>".to_vec()),
            ("word/custom/default.xml".to_owned(), b"<vbaData/>".to_vec()),
            (
                "word/_rels/document.xml.rels".to_owned(),
                br#"<Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'><Relationship Id='safe' Type='settings' Target='settings.xml'/><Relationship Id='vba' Type='vbaData' Target='custom/data.xml'/><Relationship Id='default' Type='vbaData' Target='custom/default.xml'/></Relationships>"#.to_vec(),
            ),
        ];
        let package = rezip_parts(&parts).unwrap();
        for sanitized in [
            sanitize_package(&package).unwrap(),
            sanitize_package_for_format(&package, "docx").unwrap(),
        ] {
            let output = unzip_parts(&sanitized).unwrap();
            assert!(has_part(&output, "[Content_Types].xml"));
            assert_eq!(output.len(), parts.len() - 2);
            assert_all_parts_declared(&output);
            assert_eq!(detect_package_kind(&output), Ok(DocumentKind::Docx));
            assert_eq!(part_bytes(&output, "word/settings.xml"), b"<settings/>");
            let all = output
                .iter()
                .map(|(_, bytes)| String::from_utf8_lossy(bytes))
                .collect::<String>();
            assert!(!all.contains("vbaData"));
            assert!(!all.contains("custom/data.xml"));
            assert!(!all.contains("custom/default.xml"));
            assert!(
                String::from_utf8_lossy(part_bytes(&output, "word/_rels/document.xml.rels"))
                    .contains("Id='safe'")
            );
        }
    }

    fn part_bytes<'a>(parts: &'a [(String, Vec<u8>)], name: &str) -> &'a [u8] {
        &parts.iter().find(|(path, _)| path == name).unwrap().1
    }

    #[test]
    fn regression_preserves_safely_typed_signature_names() {
        let safe_type =
            "application/vnd.openxmlformats-officedocument.spreadsheetml.printerSettings";
        for signature in [
            "vbaProjectSignature",
            "vbaProjectSignatureAgile",
            "vbaProjectSignatureV3",
        ] {
            let part = format!("xl/printerSettings/{signature}.bin");
            for use_override in [false, true] {
                let mut parts = default_payload_parts(
                    if use_override {
                        "application/vnd.ms-office.vbaProject"
                    } else {
                        safe_type
                    },
                    &part,
                );
                parts[3].1 = b"printer settings".to_vec();
                if use_override {
                    parts[0].1 = String::from_utf8(parts[0].1.clone())
                        .unwrap()
                        .replace(
                            "</Types>",
                            &format!(
                                "<Override PartName='/{part}' ContentType='{safe_type}'/></Types>"
                            ),
                        )
                        .into_bytes();
                }
                let package = rezip_parts(&parts).unwrap();
                let output =
                    unzip_parts(&sanitize_package_for_format(&package, "xlsx").unwrap()).unwrap();
                assert!(
                    has_part(&output, &part),
                    "{signature}, override={use_override}"
                );
                assert_eq!(part_bytes(&output, &part), b"printer settings");
                assert_eq!(
                    parse_content_types(&output)
                        .unwrap()
                        .content_type_for(&part),
                    Some(safe_type)
                );
                let relationships =
                    String::from_utf8_lossy(part_bytes(&output, "xl/_rels/workbook.xml.rels"));
                assert!(relationships.contains("Id='payload'"));
                assert!(relationships.contains(&part));
                assert_all_parts_declared(&output);
            }
        }
    }

    const MACRO_WORKBOOK: &[u8] = br#"<workbook xmlns='http://schemas.openxmlformats.org/spreadsheetml/2006/main' xmlns:r='http://schemas.openxmlformats.org/officeDocument/2006/relationships'><bookViews><workbookView activeTab='2' firstSheet='0'/><workbookView activeTab='1' firstSheet='1'></workbookView><workbookView activeTab='0' firstSheet='2'/></bookViews><sheets><sheet name='First' sheetId='1' r:id='first'/><sheet name='Macro' sheetId='2' r:id='macro'></sheet><sheet name='Last' sheetId='3' r:id='last'/></sheets><definedNames><definedName name='first' localSheetId='0'>First!$A$1</definedName><definedName name='macro' localSheetId='1'>Macro!$A$1</definedName><definedName name='last' localSheetId='2'>Last!$A$1</definedName><definedName name='global'>First!$A$1</definedName><definedName name='empty' localSheetId='1'/></definedNames></workbook>"#;

    fn macro_sheet_parts(content_type: &str, workbook: &[u8]) -> Vec<(String, Vec<u8>)> {
        let mut parts = package_parts("xl/workbook.xml", content_type, workbook);
        parts[0].1 = String::from_utf8(parts[0].1.clone()).unwrap().replace("</Types>", "<Default Extension='xml' ContentType='application/xml'/><Default Extension='rels' ContentType='application/vnd.openxmlformats-package.relationships+xml'/><Override PartName='/xl/custom/sheet.data' ContentType='application/vnd.ms-excel.macrosheet+xml'/></Types>").into_bytes();
        parts.extend([
            ("xl/worksheets/first.xml".to_owned(), b"<worksheet/>".to_vec()),
            ("xl/custom/sheet.data".to_owned(), b"<macrosheet/>".to_vec()),
            ("xl/worksheets/last.xml".to_owned(), b"<worksheet/>".to_vec()),
            ("xl/_rels/workbook.xml.rels".to_owned(), br#"<Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'><Relationship Id='first' Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet' Target='worksheets/first.xml'/><Relationship Id='macro' Type='http://schemas.microsoft.com/office/2006/relationships/xlMacrosheet' Target='/xl/custom/sheet.data'/><Relationship Id='last' Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet' Target='worksheets/last.xml'/></Relationships>"#.to_vec()),
        ]);
        parts
    }

    fn element_attributes(xml: &[u8], element: &[u8]) -> Vec<Vec<(String, String)>> {
        let mut reader = Reader::from_reader(xml);
        let mut output = Vec::new();
        loop {
            match reader.read_event().unwrap() {
                Event::Start(start) | Event::Empty(start)
                    if start.local_name().as_ref() == element =>
                {
                    output.push(attributes(&reader, &start, "test.xml").unwrap())
                }
                Event::Eof => return output,
                _ => {}
            }
        }
    }

    fn assert_macro_sheet_metadata(content_type: &str, format: &str) {
        for strict in [false, true] {
            let workbook = if strict {
                String::from_utf8_lossy(MACRO_WORKBOOK)
                    .replace(
                        "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
                        "http://purl.oclc.org/ooxml/spreadsheetml/main",
                    )
                    .replace(
                        "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
                        "http://purl.oclc.org/ooxml/officeDocument/relationships",
                    )
                    .into_bytes()
            } else {
                MACRO_WORKBOOK.to_vec()
            };
            let parts = macro_sheet_parts(content_type, &workbook);
            let package = rezip_parts(&parts).unwrap();
            for sanitized in [
                sanitize_package(&package).unwrap(),
                sanitize_package_for_format(&package, format).unwrap(),
            ] {
                let output = unzip_parts(&sanitized).unwrap();
                let workbook = part_bytes(&output, "xl/workbook.xml");
                let sheets = element_attributes(workbook, b"sheet");
                assert_eq!(sheets.len(), 2);
                assert_eq!(attribute_value(&sheets[0], "r:id"), Some("first"));
                assert_eq!(attribute_value(&sheets[1], "r:id"), Some("last"));
                assert_eq!(attribute_value(&sheets[1], "sheetId"), Some("3"));
                let names = element_attributes(workbook, b"definedName");
                assert_eq!(names.len(), 3);
                assert_eq!(attribute_value(&names[0], "localSheetId"), Some("0"));
                assert_eq!(attribute_value(&names[1], "localSheetId"), Some("1"));
                assert_eq!(attribute_value(&names[1], "name"), Some("last"));
                assert_eq!(attribute_value(&names[2], "localSheetId"), None);
                let views = element_attributes(workbook, b"workbookView");
                let indices = views
                    .iter()
                    .map(|values| {
                        (
                            attribute_value(values, "activeTab"),
                            attribute_value(values, "firstSheet"),
                        )
                    })
                    .collect::<Vec<_>>();
                assert_eq!(
                    indices,
                    [
                        (Some("1"), Some("0")),
                        (Some("1"), Some("1")),
                        (Some("0"), Some("1"))
                    ]
                );
                assert!(!String::from_utf8_lossy(workbook).contains("macro"));
                assert!(
                    !String::from_utf8_lossy(part_bytes(&output, "xl/_rels/workbook.xml.rels"))
                        .contains("macro")
                );
                assert!(!has_part(&output, "xl/custom/sheet.data"));
                assert_all_parts_declared(&output);
            }
        }
    }

    #[test]
    fn regression_removes_xltm_macro_sheet_metadata() {
        assert_macro_sheet_metadata(
            "application/vnd.ms-excel.template.macroEnabled.main+xml",
            "xltx",
        );
    }

    #[test]
    fn regression_removes_xlam_macro_sheet_metadata() {
        assert_macro_sheet_metadata(
            "application/vnd.ms-excel.addin.macroEnabled.main+xml",
            "xlsx",
        );
    }

    #[test]
    fn regression_removes_xlsm_macro_sheet_metadata() {
        assert_macro_sheet_metadata(
            "application/vnd.ms-excel.sheet.macroEnabled.main+xml",
            "xlsx",
        );
    }

    #[test]
    fn regression_rejects_unsafe_macro_sheet_workbook_rewrites() {
        for workbook in [
            "<root xmlns:r='http://schemas.openxmlformats.org/officeDocument/2006/relationships'><sheet r:id='macro'/></root>".to_owned(),
            String::from_utf8_lossy(MACRO_WORKBOOK).replace(
                "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
                "urn:foreign",
            ),
            String::from_utf8_lossy(MACRO_WORKBOOK).replace("r:id='last'", "r:id='first'"),
            String::from_utf8_lossy(MACRO_WORKBOOK).replace("r:id='last'", "r:id='unknown'"),
            String::from_utf8_lossy(MACRO_WORKBOOK)
                .replace("localSheetId='2'", "localSheetId='invalid'"),
        ] {
            let parts = macro_sheet_parts(
                "application/vnd.ms-excel.sheet.macroEnabled.main+xml",
                workbook.as_bytes(),
            );
            let package = rezip_parts(&parts).unwrap();
            for result in [
                sanitize_package(&package),
                sanitize_package_for_format(&package, "xlsx"),
            ] {
                let error = result.unwrap_err();
                assert!(error.contains("xl/workbook.xml"), "{error}");
            }
        }
    }

    #[test]
    fn regression_clamps_views_when_last_macro_sheet_is_removed() {
        let workbook = String::from_utf8_lossy(MACRO_WORKBOOK)
            .replace("sheetId='2' r:id='macro'", "sheetId='2' r:id='last'")
            .replace("sheetId='3' r:id='last'", "sheetId='3' r:id='macro'");
        let parts = macro_sheet_parts(
            "application/vnd.ms-excel.sheet.macroEnabled.main+xml",
            workbook.as_bytes(),
        );
        let package = rezip_parts(&parts).unwrap();
        let output = unzip_parts(&sanitize_package_for_format(&package, "xlsx").unwrap()).unwrap();
        let workbook = part_bytes(&output, "xl/workbook.xml");
        assert_eq!(element_attributes(workbook, b"sheet").len(), 2);
        for values in element_attributes(workbook, b"workbookView") {
            for key in ["activeTab", "firstSheet"] {
                let index = attribute_value(&values, key)
                    .unwrap()
                    .parse::<usize>()
                    .unwrap();
                assert!(index < 2);
            }
        }
    }

    #[test]
    fn regression_default_main_types_preserve_typed_payloads() {
        for [format, content_type, sanitized_format, sanitized_type] in TEMPLATE_AND_ADDIN_KINDS {
            let (part_name, _) = office_main_part(format);
            let mut parts = package_parts(part_name, content_type, b"<root/>");
            parts[0].1 = format!("<Types xmlns='http://schemas.openxmlformats.org/package/2006/content-types'><Default Extension='XML' ContentType='{content_type}'/><Default Extension='bin' ContentType='application/vnd.openxmlformats-officedocument.spreadsheetml.printerSettings'/><Default Extension='rels' ContentType='application/vnd.openxmlformats-package.relationships+xml'/></Types>").into_bytes();
            let safe_part = "custom/vbaProjectSignature.bin";
            parts.push((safe_part.to_owned(), b"printer settings".to_vec()));
            assert_eq!(detect_package_kind(&parts).unwrap().format(), format);
            let package = rezip_parts(&parts).unwrap();
            for sanitized in [
                sanitize_package(&package).unwrap(),
                sanitize_package_for_format(&package, sanitized_format).unwrap(),
            ] {
                let output = unzip_parts(&sanitized).unwrap();
                assert_eq!(
                    detect_package_kind(&output).unwrap().format(),
                    sanitized_format
                );
                assert_eq!(
                    parse_content_types(&output)
                        .unwrap()
                        .content_type_for(part_name),
                    Some(sanitized_type)
                );
                assert!(has_part(&output, safe_part), "{format}");
                assert_all_parts_declared(&output);
            }
        }
    }

    #[test]
    fn regression_conflicting_main_override_wins_and_preserves_typed_payloads() {
        let mut parts = default_payload_parts(
            "application/vnd.openxmlformats-officedocument.spreadsheetml.printerSettings",
            "xl/printerSettings/vbaProjectSignature.bin",
        );
        parts[0].1 = String::from_utf8(parts[0].1.clone()).unwrap().replace("</Types>", "<Default Extension='xml' ContentType='application/vnd.ms-excel.template.macroEnabled.main+xml'/></Types>").replace("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml", "application/vnd.ms-excel.addin.macroEnabled.main+xml").into_bytes();
        assert_eq!(detect_package_kind(&parts), Ok(DocumentKind::Xlam));
        let package = rezip_parts(&parts).unwrap();
        assert!(sanitize_package_for_format(&package, "xltx").is_err());
        let output = unzip_parts(&sanitize_package_for_format(&package, "xlsx").unwrap()).unwrap();
        assert_eq!(detect_package_kind(&output), Ok(DocumentKind::Xlsx));
        let types = parse_content_types(&output).unwrap();
        assert_eq!(
            types.content_type_for("xl/workbook.xml"),
            Some("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml")
        );
        assert_eq!(
            types.content_type_for("custom/main.xml"),
            Some("application/vnd.openxmlformats-officedocument.spreadsheetml.template.main+xml")
        );
        assert!(has_part(
            &output,
            "xl/printerSettings/vbaProjectSignature.bin"
        ));
        assert!(
            String::from_utf8_lossy(part_bytes(&output, "xl/_rels/workbook.xml.rels"))
                .contains("Id='payload'")
        );
        assert_all_parts_declared(&output);
    }

    const XLSX_MAIN_TYPE: &str =
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml";
    const XLSM_MAIN_TYPE: &str = "application/vnd.ms-excel.sheet.macroEnabled.main+xml";
    const PRINTER_SETTINGS_TYPE: &str =
        "application/vnd.openxmlformats-officedocument.spreadsheetml.printerSettings";
    const CALC_CHAIN_TYPE: &str =
        "application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml";
    const ORDINARY_WORKBOOK: &[u8] = br#"<workbook xmlns='http://schemas.openxmlformats.org/spreadsheetml/2006/main' xmlns:r='http://schemas.openxmlformats.org/officeDocument/2006/relationships'><sheets><sheet name='First' sheetId='1' r:id='first'/><sheet name='Last' sheetId='3' r:id='last'/></sheets></workbook>"#;
    const WORKSHEET_RELATIONSHIPS: &str = "<Relationship Id='first' Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet' Target='worksheets/first.xml'/><Relationship Id='last' Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet' Target='worksheets/last.xml'/>";
    const EXTERNAL_LINK: &str = "Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink' Target='https://example.invalid/' TargetMode='External'";
    const CALC_CHAIN: &[u8] = br#"<calcChain xmlns='http://schemas.openxmlformats.org/spreadsheetml/2006/main'><c r='A1' i='2'/><c r='B1'/><c r='A1' i='1'/></calcChain>"#;

    fn relationships_xml(body: &str) -> Vec<u8> {
        format!(
            "<Relationships xmlns='{}'>{body}</Relationships>",
            String::from_utf8_lossy(RELATIONSHIPS_NAMESPACE)
        )
        .into_bytes()
    }

    fn declare(parts: &mut [(String, Vec<u8>)], declarations: &str) {
        parts[0].1 = String::from_utf8(parts[0].1.clone())
            .unwrap()
            .replace("</Types>", &format!("{declarations}</Types>"))
            .into_bytes();
    }

    fn append_relationships(parts: &mut [(String, Vec<u8>)], path: &str, body: &str) {
        let (_, bytes) = parts.iter_mut().find(|(name, _)| name == path).unwrap();
        *bytes = String::from_utf8(bytes.clone())
            .unwrap()
            .replace("</Relationships>", &format!("{body}</Relationships>"))
            .into_bytes();
    }

    fn ordinary_workbook_parts(relationships: &str) -> Vec<(String, Vec<u8>)> {
        let mut parts = package_parts("xl/workbook.xml", XLSX_MAIN_TYPE, ORDINARY_WORKBOOK);
        declare(
            &mut parts,
            "<Default Extension='xml' ContentType='application/xml'/><Default Extension='rels' ContentType='application/vnd.openxmlformats-package.relationships+xml'/>",
        );
        parts.extend([
            (
                "xl/worksheets/first.xml".to_owned(),
                b"<worksheet/>".to_vec(),
            ),
            (
                "xl/worksheets/last.xml".to_owned(),
                b"<worksheet/>".to_vec(),
            ),
            (
                "xl/_rels/workbook.xml.rels".to_owned(),
                relationships_xml(relationships),
            ),
        ]);
        parts
    }

    fn assert_first_and_last_sheets_kept(parts: &[(String, Vec<u8>)]) {
        let package = rezip_parts(parts).unwrap();
        for sanitized in [
            sanitize_package(&package).unwrap(),
            sanitize_package_for_format(&package, "xlsx").unwrap(),
        ] {
            let output = unzip_parts(&sanitized).unwrap();
            let sheets = element_attributes(part_bytes(&output, "xl/workbook.xml"), b"sheet");
            assert_eq!(sheets.len(), 2);
            assert_eq!(attribute_value(&sheets[0], "r:id"), Some("first"));
            assert_eq!(attribute_value(&sheets[1], "r:id"), Some("last"));
            let relationships =
                String::from_utf8_lossy(part_bytes(&output, "xl/_rels/workbook.xml.rels"));
            assert!(
                relationships.contains("worksheets/first.xml"),
                "{relationships}"
            );
            assert!(
                relationships.contains("worksheets/last.xml"),
                "{relationships}"
            );
            assert!(
                !relationships.contains("example.invalid"),
                "{relationships}"
            );
            assert!(!relationships.contains("macro"), "{relationships}");
            assert!(!has_part(&output, "xl/custom/sheet.data"));
            assert_all_parts_declared(&output);
        }
    }

    #[test]
    fn regression_ordinary_workbook_drops_id_less_external_relationships() {
        let body = format!("{WORKSHEET_RELATIONSHIPS}<Relationship {EXTERNAL_LINK}/>");
        assert_first_and_last_sheets_kept(&ordinary_workbook_parts(&body));
    }

    #[test]
    fn regression_ordinary_workbook_keeps_sheets_shadowed_by_discarded_ids() {
        let body = format!(
            "<Relationship Id='first' {EXTERNAL_LINK}/>{WORKSHEET_RELATIONSHIPS}<Relationship Id='last' {EXTERNAL_LINK}/>"
        );
        assert_first_and_last_sheets_kept(&ordinary_workbook_parts(&body));
    }

    #[test]
    fn regression_ordinary_workbook_tolerates_invalid_relationship_roots() {
        for (root, namespace) in [
            ("Relationships", "urn:foreign"),
            (
                "Rels",
                "http://schemas.openxmlformats.org/package/2006/relationships",
            ),
        ] {
            let mut parts = ordinary_workbook_parts("");
            parts.last_mut().unwrap().1 = format!(
                "<{root} xmlns='{namespace}'>{WORKSHEET_RELATIONSHIPS}<Relationship {EXTERNAL_LINK}/></{root}>"
            )
            .into_bytes();
            assert_first_and_last_sheets_kept(&parts);
        }
    }

    #[test]
    fn regression_macro_sheet_rewrite_ignores_id_less_and_discarded_duplicates() {
        let mut parts = macro_sheet_parts(XLSM_MAIN_TYPE, MACRO_WORKBOOK);
        append_relationships(
            &mut parts,
            "xl/_rels/workbook.xml.rels",
            &format!("<Relationship Id='first' {EXTERNAL_LINK}/><Relationship {EXTERNAL_LINK}/>"),
        );
        assert_first_and_last_sheets_kept(&parts);
    }

    fn root_sheet_parts(
        bin_type: &str,
        payload: &str,
        payload_type: &str,
        worksheet: &[u8],
    ) -> Vec<(String, Vec<u8>)> {
        let mut parts = ordinary_workbook_parts(
            &WORKSHEET_RELATIONSHIPS.replace("worksheets/first.xml", "../sheet1.xml"),
        );
        parts.retain(|(path, _)| path != "xl/worksheets/first.xml");
        declare(
            &mut parts,
            &format!(
                "<Default Extension='bin' ContentType='{bin_type}'/><Override PartName='/{payload}' ContentType='{payload_type}'/>"
            ),
        );
        parts.extend([
            ("sheet1.xml".to_owned(), worksheet.to_vec()),
            (
                "_rels/sheet1.xml.rels".to_owned(),
                relationships_xml(&format!(
                    "<Relationship Id='printer' Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/printerSettings' Target='{payload}'/>"
                )),
            ),
            (payload.to_owned(), b"printer settings".to_vec()),
        ]);
        parts
    }

    #[test]
    fn regression_root_level_relationships_keep_safely_typed_targets() {
        let worksheet = br#"<worksheet xmlns='http://schemas.openxmlformats.org/spreadsheetml/2006/main' xmlns:r='http://schemas.openxmlformats.org/officeDocument/2006/relationships'><pageSetup r:id='printer'/></worksheet>"#;
        let parts = root_sheet_parts(
            "application/vnd.ms-office.vbaProject",
            "printer.bin",
            PRINTER_SETTINGS_TYPE,
            worksheet,
        );
        let package = rezip_parts(&parts).unwrap();
        for sanitized in [
            sanitize_package(&package).unwrap(),
            sanitize_package_for_format(&package, "xlsx").unwrap(),
        ] {
            let output = unzip_parts(&sanitized).unwrap();
            assert_eq!(part_bytes(&output, "printer.bin"), b"printer settings");
            assert_eq!(
                parse_content_types(&output)
                    .unwrap()
                    .content_type_for("printer.bin"),
                Some(PRINTER_SETTINGS_TYPE)
            );
            let relationships =
                String::from_utf8_lossy(part_bytes(&output, "_rels/sheet1.xml.rels"));
            assert!(
                relationships.contains("Target='printer.bin'"),
                "{relationships}"
            );
            assert_eq!(part_bytes(&output, "sheet1.xml"), worksheet);
            assert_all_parts_declared(&output);
        }
    }

    #[test]
    fn regression_root_level_relationships_drop_dangerously_typed_targets() {
        let parts = root_sheet_parts(
            PRINTER_SETTINGS_TYPE,
            "payload.bin",
            "application/vnd.ms-office.vbaProject",
            b"<worksheet/>",
        );
        let package = rezip_parts(&parts).unwrap();
        for sanitized in [
            sanitize_package(&package).unwrap(),
            sanitize_package_for_format(&package, "xlsx").unwrap(),
        ] {
            let output = unzip_parts(&sanitized).unwrap();
            assert!(!has_part(&output, "payload.bin"));
            let relationships =
                String::from_utf8_lossy(part_bytes(&output, "_rels/sheet1.xml.rels"));
            assert!(!relationships.contains("payload.bin"), "{relationships}");
            let all = output
                .iter()
                .map(|(_, bytes)| String::from_utf8_lossy(bytes))
                .collect::<String>();
            assert!(!all.contains("payload.bin"));
            assert!(!all.contains("vbaProject"));
            assert_all_parts_declared(&output);
        }
    }

    fn add_calc_chain(parts: &mut Vec<(String, Vec<u8>)>, relationship_type: &str) {
        declare(
            parts,
            &format!("<Override PartName='/xl/calcChain.xml' ContentType='{CALC_CHAIN_TYPE}'/>"),
        );
        append_relationships(
            parts,
            "xl/_rels/workbook.xml.rels",
            &format!("<Relationship Id='calc' Type='{relationship_type}' Target='calcChain.xml'/>"),
        );
        parts.push(("xl/calcChain.xml".to_owned(), CALC_CHAIN.to_vec()));
    }

    #[test]
    fn regression_removes_calc_chain_with_macro_sheets() {
        let transitional =
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain";
        let mut parts = ordinary_workbook_parts(WORKSHEET_RELATIONSHIPS);
        add_calc_chain(&mut parts, transitional);
        let package = rezip_parts(&parts).unwrap();
        for sanitized in [
            sanitize_package(&package).unwrap(),
            sanitize_package_for_format(&package, "xlsx").unwrap(),
        ] {
            let output = unzip_parts(&sanitized).unwrap();
            assert_eq!(part_bytes(&output, "xl/calcChain.xml"), CALC_CHAIN);
            assert_eq!(
                parse_content_types(&output)
                    .unwrap()
                    .content_type_for("xl/calcChain.xml"),
                Some(CALC_CHAIN_TYPE)
            );
            assert!(
                String::from_utf8_lossy(part_bytes(&output, "xl/_rels/workbook.xml.rels"))
                    .contains("Id='calc'")
            );
        }
        for relationship_type in [
            transitional.to_owned(),
            transitional.to_ascii_uppercase(),
            "http://purl.oclc.org/ooxml/officeDocument/relationships/calcChain".to_owned(),
        ] {
            let mut parts = macro_sheet_parts(XLSM_MAIN_TYPE, MACRO_WORKBOOK);
            add_calc_chain(&mut parts, &relationship_type);
            let package = rezip_parts(&parts).unwrap();
            for sanitized in [
                sanitize_package(&package).unwrap(),
                sanitize_package_for_format(&package, "xlsx").unwrap(),
            ] {
                let output = unzip_parts(&sanitized).unwrap();
                assert!(
                    !has_part(&output, "xl/calcChain.xml"),
                    "{relationship_type}"
                );
                let workbook = part_bytes(&output, "xl/workbook.xml");
                assert_eq!(element_attributes(workbook, b"sheet").len(), 2);
                let all = output
                    .iter()
                    .map(|(_, bytes)| String::from_utf8_lossy(bytes).to_ascii_lowercase())
                    .collect::<String>();
                assert!(!all.contains("calcchain"), "{relationship_type}");
                assert_all_parts_declared(&output);
            }
        }
    }

    #[test]
    fn regression_removes_default_declared_calc_chain_with_macro_sheets() {
        let mut parts = macro_sheet_parts(XLSM_MAIN_TYPE, MACRO_WORKBOOK);
        declare(
            &mut parts,
            &format!("<Default Extension='chain' ContentType='{CALC_CHAIN_TYPE}'/>"),
        );
        append_relationships(
            &mut parts,
            "xl/_rels/workbook.xml.rels",
            "<Relationship Id='calc' Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain' Target='calc.chain'/>",
        );
        parts.push(("xl/calc.chain".to_owned(), CALC_CHAIN.to_vec()));
        for output in sanitize_both(&parts) {
            assert!(!has_part(&output, "xl/calc.chain"));
            let relationships =
                String::from_utf8_lossy(part_bytes(&output, "xl/_rels/workbook.xml.rels"));
            assert!(!relationships.contains("calc"), "{relationships}");
            assert_all_parts_declared(&output);
        }
    }

    const VBA_PROJECT_TYPE: &str = "application/vnd.ms-office.vbaProject";
    const NESTED_NAMES_WORKBOOK: &[u8] = br#"<workbook xmlns='http://schemas.openxmlformats.org/spreadsheetml/2006/main' xmlns:r='http://schemas.openxmlformats.org/officeDocument/2006/relationships' xmlns:mc='http://schemas.openxmlformats.org/markup-compatibility/2006'><sheets><sheet name='First' sheetId='1' r:id='first'/><sheet name='Last' sheetId='3' r:id='last'/></sheets><definedNames><mc:AlternateContent><mc:Choice Requires='r'><definedName name='choice'>First!$A$1</definedName></mc:Choice><mc:Fallback><definedName name='fallback'>Last!$A$1</definedName></mc:Fallback></mc:AlternateContent></definedNames></workbook>"#;

    fn replace_part_text(parts: &mut [(String, Vec<u8>)], path: &str, from: &str, to: &str) {
        let (_, bytes) = parts.iter_mut().find(|(name, _)| name == path).unwrap();
        let text = String::from_utf8(bytes.clone()).unwrap();
        assert!(text.contains(from), "{path}: {from}");
        *bytes = text.replace(from, to).into_bytes();
    }

    fn sanitize_both(parts: &[(String, Vec<u8>)]) -> [Vec<(String, Vec<u8>)>; 2] {
        let package = rezip_parts(parts).unwrap();
        [
            sanitize_package(&package).unwrap(),
            sanitize_package_for_format(&package, "xlsx").unwrap(),
        ]
        .map(|sanitized| unzip_parts(&sanitized).unwrap())
    }

    #[test]
    fn regression_unreferenced_macro_sheets_leave_the_workbook_untouched() {
        for root in ["Relationships", "Rels"] {
            let mut parts = macro_sheet_parts(XLSM_MAIN_TYPE, NESTED_NAMES_WORKBOOK);
            replace_part_text(
                &mut parts,
                "xl/_rels/workbook.xml.rels",
                "Relationships",
                root,
            );
            for output in sanitize_both(&parts) {
                assert_eq!(
                    part_bytes(&output, "xl/workbook.xml"),
                    NESTED_NAMES_WORKBOOK
                );
                let relationships =
                    String::from_utf8_lossy(part_bytes(&output, "xl/_rels/workbook.xml.rels"));
                assert!(
                    relationships.contains("worksheets/last.xml"),
                    "{relationships}"
                );
                assert!(!relationships.contains("macro"), "{relationships}");
                assert!(!has_part(&output, "xl/custom/sheet.data"));
                assert_all_parts_declared(&output);
            }
        }
    }

    fn wrap_alternate_content(workbook: &[u8], sheet: &str) -> String {
        let wrapped = format!(
            "<mc:AlternateContent xmlns:mc='http://schemas.openxmlformats.org/markup-compatibility/2006'><mc:Choice Requires='r'>{sheet}</mc:Choice><mc:Fallback>{sheet}</mc:Fallback></mc:AlternateContent>"
        );
        let workbook = String::from_utf8_lossy(workbook);
        assert!(workbook.contains(sheet), "{sheet}");
        workbook.replace(sheet, &wrapped)
    }

    #[test]
    fn regression_unreferenced_macro_sheets_ignore_workbook_structure() {
        for workbook in [
            wrap_alternate_content(
                NESTED_NAMES_WORKBOOK,
                "<sheet name='First' sheetId='1' r:id='first'/>",
            ),
            "<root/>".to_owned(),
            String::from_utf8_lossy(NESTED_NAMES_WORKBOOK).replace("name='Last'", "name='&x;'"),
            String::from_utf8_lossy(MACRO_WORKBOOK).replace("r:id='macro'", "r:id='first'"),
            String::from_utf8_lossy(MACRO_WORKBOOK).replace("r:id='macro'", "r:id='unknown'"),
        ] {
            for root in ["Relationships", "Rels"] {
                let mut parts = macro_sheet_parts(XLSM_MAIN_TYPE, workbook.as_bytes());
                replace_part_text(
                    &mut parts,
                    "xl/_rels/workbook.xml.rels",
                    "Relationships",
                    root,
                );
                for output in sanitize_both(&parts) {
                    assert_eq!(part_bytes(&output, "xl/workbook.xml"), workbook.as_bytes());
                    let relationships =
                        String::from_utf8_lossy(part_bytes(&output, "xl/_rels/workbook.xml.rels"));
                    for kept in ["worksheets/first.xml", "worksheets/last.xml"] {
                        assert!(relationships.contains(kept), "{relationships}");
                    }
                    assert!(!relationships.contains("macro"), "{relationships}");
                    assert!(!has_part(&output, "xl/custom/sheet.data"));
                    assert_all_parts_declared(&output);
                }
            }
        }
    }

    const ALTERNATE_CONTENT_ELEMENTS: [&str; 4] = [
        "<sheet name='Macro' sheetId='2' r:id='macro'></sheet>",
        "<sheet name='First' sheetId='1' r:id='first'/>",
        "<definedName name='last' localSheetId='2'>Last!$A$1</definedName>",
        "<workbookView activeTab='1' firstSheet='1'></workbookView>",
    ];

    fn alternate_content_parts(element: &str) -> (String, Vec<(String, Vec<u8>)>) {
        let workbook = wrap_alternate_content(MACRO_WORKBOOK, element);
        let mut parts = macro_sheet_parts(XLSM_MAIN_TYPE, workbook.as_bytes());
        add_calc_chain(
            &mut parts,
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain",
        );
        (workbook, parts)
    }

    #[test]
    fn regression_alternate_content_sheet_metadata_keeps_the_workbook() {
        for element in ALTERNATE_CONTENT_ELEMENTS {
            let (workbook, parts) = alternate_content_parts(element);
            for output in sanitize_both(&parts) {
                assert_eq!(part_bytes(&output, "xl/workbook.xml"), workbook.as_bytes());
                assert_eq!(part_bytes(&output, "xl/calcChain.xml"), CALC_CHAIN);
                let relationships =
                    String::from_utf8_lossy(part_bytes(&output, "xl/_rels/workbook.xml.rels"));
                for kept in ["worksheets/first.xml", "worksheets/last.xml", "Id='calc'"] {
                    assert!(relationships.contains(kept), "{relationships}");
                }
                assert!(!relationships.contains("macro"), "{relationships}");
                assert!(!has_part(&output, "xl/custom/sheet.data"));
                assert_all_parts_declared(&output);
            }
        }
    }

    // Escaping whole single-quoted declarations must not change what either entry point does;
    // re-serialized declarations come out double-quoted, so only those kept as written differ.
    fn assert_escaped_namespaces_match_literal(
        parts: &[(String, Vec<u8>)],
        escapes: &[(&str, &str)],
    ) -> [Vec<(String, Vec<u8>)>; 2] {
        let escape = |parts: &[(String, Vec<u8>)]| {
            parts
                .iter()
                .map(|(name, bytes)| {
                    let text = escapes.iter().fold(
                        String::from_utf8_lossy(bytes).into_owned(),
                        |text, (literal, escaped)| text.replace(literal, escaped),
                    );
                    (name.clone(), text)
                })
                .collect::<Vec<_>>()
        };
        for (literal, _) in escapes {
            assert!(
                parts
                    .iter()
                    .any(|(_, bytes)| String::from_utf8_lossy(bytes).contains(literal)),
                "{literal}"
            );
        }
        let escaped_parts = escape(parts)
            .into_iter()
            .map(|(name, text)| (name, text.into_bytes()))
            .collect::<Vec<_>>();
        let literal = sanitize_both(parts);
        let escaped = sanitize_both(&escaped_parts);
        for (literal, escaped) in literal.iter().zip(&escaped) {
            let escaped = escaped
                .iter()
                .map(|(name, bytes)| (name.clone(), String::from_utf8_lossy(bytes).into_owned()))
                .collect::<Vec<_>>();
            assert_eq!(escaped, escape(literal));
        }
        escaped
    }

    #[test]
    fn regression_escaped_compatibility_namespace_keeps_the_workbook() {
        for element in ALTERNATE_CONTENT_ELEMENTS {
            assert_escaped_namespaces_match_literal(
                &alternate_content_parts(element).1,
                &[(
                    "xmlns:mc='http://schemas.openxmlformats.org/markup-compatibility/2006'",
                    "xmlns:mc='http://schemas.openxmlformats.org/markup-compatibility/200&#54;'",
                )],
            );
        }
    }

    #[test]
    fn regression_escaped_workbook_namespaces_match_the_literal_spelling() {
        let prefixed = "<s:definedName xmlns:s='http://schemas.openxmlformats.org/spreadsheetml/2006/main' name='last' localSheetId='2'>Last!$A$1</s:definedName>";
        let direct = String::from_utf8_lossy(MACRO_WORKBOOK).replace(
            "<definedName name='last' localSheetId='2'>Last!$A$1</definedName>",
            prefixed,
        );
        let plain = String::from_utf8_lossy(MACRO_WORKBOOK).into_owned();
        let prefix = (
            "xmlns:s='http://schemas.openxmlformats.org/spreadsheetml/2006/main'",
            "xmlns:s='http://schemas.openxmlformats.org/spreadsheetml/2006/mai&#110;'",
        );
        for (workbook, escape) in [
            (wrap_alternate_content(direct.as_bytes(), prefixed), prefix),
            (direct, prefix),
            (
                plain.clone(),
                (
                    "xmlns='http://schemas.openxmlformats.org/spreadsheetml/2006/main'",
                    "xmlns='http://schemas.openxmlformats.org/spreadsheetml/2006/mai&#110;'",
                ),
            ),
            (
                plain,
                (
                    "xmlns:r='http://schemas.openxmlformats.org/officeDocument/2006/relationships'",
                    "xmlns:r='http://schemas.openxmlformats.org/officeDocument/2006/relationship&#115;'",
                ),
            ),
        ] {
            let parts = macro_sheet_parts(XLSM_MAIN_TYPE, workbook.as_bytes());
            for output in assert_escaped_namespaces_match_literal(&parts, &[escape]) {
                let workbook = part_bytes(&output, "xl/workbook.xml");
                let sheets = element_attributes(workbook, b"sheet").len();
                for values in element_attributes(workbook, b"definedName") {
                    if let Some(index) = attribute_value(&values, "localSheetId") {
                        assert!(index.parse::<usize>().unwrap() < sheets, "{index}");
                    }
                }
            }
        }
    }

    #[test]
    fn regression_escaped_package_namespaces_match_the_literal_spelling() {
        assert_escaped_namespaces_match_literal(
            &macro_sheet_parts(XLSM_MAIN_TYPE, MACRO_WORKBOOK),
            &[
                (
                    "xmlns='http://schemas.openxmlformats.org/package/2006/content-types'",
                    "xmlns='http://schemas.openxmlformats.org/package/2006/content-type&#115;'",
                ),
                (
                    "xmlns='http://schemas.openxmlformats.org/package/2006/relationships'",
                    "xmlns='http://schemas.openxmlformats.org/package/2006/relationship&#115;'",
                ),
            ],
        );
    }

    #[test]
    fn regression_remaps_views_across_many_removed_sheets() {
        const SHEETS: usize = 3_000;
        let removed = |index: usize| index.is_multiple_of(3) || index >= SHEETS - 3;
        let mut sheets = String::new();
        let mut views = String::new();
        let mut relationships = String::new();
        for index in 0..SHEETS {
            let (id, relationship) = if removed(index) {
                (
                    format!("m{index}"),
                    "Type='http://schemas.microsoft.com/office/2006/relationships/xlMacrosheet' Target='/xl/custom/sheet.data'",
                )
            } else {
                (
                    format!("k{index}"),
                    "Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet' Target='worksheets/first.xml'",
                )
            };
            sheets.push_str(&format!(
                "<sheet name='S{index}' sheetId='{}' r:id='{id}'/>",
                index + 1
            ));
            views.push_str(&format!(
                "<workbookView activeTab='{index}' firstSheet='{}'/>",
                SHEETS - 1 - index
            ));
            relationships.push_str(&format!("<Relationship Id='{id}' {relationship}/>"));
        }
        let workbook = format!(
            "<workbook xmlns='http://schemas.openxmlformats.org/spreadsheetml/2006/main' xmlns:r='http://schemas.openxmlformats.org/officeDocument/2006/relationships'><bookViews>{views}</bookViews><sheets>{sheets}</sheets></workbook>"
        );
        let mut parts = macro_sheet_parts(XLSM_MAIN_TYPE, workbook.as_bytes());
        let (_, rels) = parts
            .iter_mut()
            .find(|(name, _)| name == "xl/_rels/workbook.xml.rels")
            .unwrap();
        *rels = relationships_xml(&relationships);
        let kept = (0..SHEETS)
            .filter(|index| !removed(*index))
            .map(|index| format!("S{index}"))
            .collect::<Vec<_>>();
        // A view on a removed sheet moves to the next kept sheet, or the last one.
        let mut next_kept = kept.last().unwrap().clone();
        let mut selected = (0..SHEETS)
            .rev()
            .map(|index| {
                if !removed(index) {
                    next_kept = format!("S{index}");
                }
                next_kept.clone()
            })
            .collect::<Vec<_>>();
        selected.reverse();
        for output in sanitize_both(&parts) {
            let workbook = part_bytes(&output, "xl/workbook.xml");
            let names = element_attributes(workbook, b"sheet")
                .iter()
                .map(|values| attribute_value(values, "name").unwrap().to_owned())
                .collect::<Vec<_>>();
            assert_eq!(names, kept);
            let position = names
                .iter()
                .enumerate()
                .map(|(position, name)| (name.as_str(), position.to_string()))
                .collect::<std::collections::HashMap<_, _>>();
            let views = element_attributes(workbook, b"workbookView");
            assert_eq!(views.len(), SHEETS);
            for (index, values) in views.iter().enumerate() {
                assert_eq!(
                    attribute_value(values, "activeTab"),
                    Some(position[selected[index].as_str()].as_str()),
                    "{index}"
                );
                assert_eq!(
                    attribute_value(values, "firstSheet"),
                    Some(position[selected[SHEETS - 1 - index].as_str()].as_str()),
                    "{index}"
                );
            }
        }
    }

    #[test]
    fn regression_targets_climbing_above_the_root_resolve_at_the_root() {
        let mut parts = root_sheet_parts(
            VBA_PROJECT_TYPE,
            "vbaProject.bin",
            VBA_PROJECT_TYPE,
            b"<worksheet/>",
        );
        replace_part_text(
            &mut parts,
            "[Content_Types].xml",
            &format!("<Override PartName='/vbaProject.bin' ContentType='{VBA_PROJECT_TYPE}'/>"),
            "",
        );
        replace_part_text(
            &mut parts,
            "_rels/sheet1.xml.rels",
            "Target='vbaProject.bin'",
            "Target='../vbaProject.bin'",
        );
        for output in sanitize_both(&parts) {
            assert!(!has_part(&output, "vbaProject.bin"));
            let relationships =
                String::from_utf8_lossy(part_bytes(&output, "_rels/sheet1.xml.rels"));
            assert!(!relationships.contains("vbaProject"), "{relationships}");
            assert_all_parts_declared(&output);
        }

        let worksheet = br#"<worksheet xmlns='http://schemas.openxmlformats.org/spreadsheetml/2006/main' xmlns:r='http://schemas.openxmlformats.org/officeDocument/2006/relationships'><pageSetup r:id='printer'/></worksheet>"#;
        let mut parts = root_sheet_parts(
            VBA_PROJECT_TYPE,
            "printer.bin",
            PRINTER_SETTINGS_TYPE,
            worksheet,
        );
        replace_part_text(
            &mut parts,
            "_rels/sheet1.xml.rels",
            "Target='printer.bin'",
            "Target='../printer.bin'",
        );
        for output in sanitize_both(&parts) {
            assert_eq!(part_bytes(&output, "printer.bin"), b"printer settings");
            assert_eq!(
                parse_content_types(&output)
                    .unwrap()
                    .content_type_for("printer.bin"),
                Some(PRINTER_SETTINGS_TYPE)
            );
            let relationships =
                String::from_utf8_lossy(part_bytes(&output, "_rels/sheet1.xml.rels"));
            assert!(
                relationships.contains("Target='../printer.bin'"),
                "{relationships}"
            );
            assert_eq!(part_bytes(&output, "sheet1.xml"), worksheet);
            assert_all_parts_declared(&output);
        }
    }

    #[test]
    fn main_relationship_climbing_above_the_root_stays_invalid() {
        let mut parts = ordinary_workbook_parts(WORKSHEET_RELATIONSHIPS);
        replace_part_text(
            &mut parts,
            "_rels/.rels",
            "Target='xl/workbook.xml'",
            "Target='../xl/workbook.xml'",
        );
        assert_eq!(
            detect_package_kind(&parts),
            Err(invalid_relationships("invalid main document target"))
        );
        let package = rezip_parts(&parts).unwrap();
        for result in [
            sanitize_package(&package),
            sanitize_package_for_format(&package, "xlsx"),
        ] {
            let error = result.unwrap_err();
            assert!(error.contains("invalid main document target"), "{error}");
        }
    }

    #[test]
    fn regression_backslash_rooted_targets_resolve_from_the_root() {
        let mut parts = ordinary_workbook_parts(
            &WORKSHEET_RELATIONSHIPS.replace("worksheets/first.xml", "sheet1.xml"),
        );
        parts.retain(|(path, _)| path != "xl/worksheets/first.xml");
        declare(
            &mut parts,
            &format!(
                "<Default Extension='bin' ContentType='{PRINTER_SETTINGS_TYPE}'/><Override PartName='/vbaProject.bin' ContentType='{VBA_PROJECT_TYPE}'/>"
            ),
        );
        let printer = r"<Relationship Id='printer' Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/printerSettings' Target='\vbaProject.bin'/>";
        parts.extend([
            ("xl/sheet1.xml".to_owned(), b"<worksheet/>".to_vec()),
            (
                "xl/_rels/sheet1.xml.rels".to_owned(),
                relationships_xml(printer),
            ),
            ("vbaProject.bin".to_owned(), b"vba payload".to_vec()),
        ]);
        for output in sanitize_both(&parts) {
            assert!(!has_part(&output, "vbaProject.bin"));
            let relationships =
                String::from_utf8_lossy(part_bytes(&output, "xl/_rels/sheet1.xml.rels"));
            assert!(!relationships.contains("vbaProject"), "{relationships}");
            assert_all_parts_declared(&output);
        }
    }

    #[test]
    fn regression_drops_targets_with_percent_encoded_separators_or_dots() {
        for target in [
            "%5CvbaProject.bin",
            "%2FvbaProject.bin",
            "%5cvbaProject.bin",
            "%2e%2E/vbaProject.bin",
        ] {
            let mut parts = ordinary_workbook_parts(
                &WORKSHEET_RELATIONSHIPS.replace("worksheets/first.xml", "sheet1.xml"),
            );
            parts.retain(|(path, _)| path != "xl/worksheets/first.xml");
            declare(
                &mut parts,
                &format!(
                    "<Default Extension='bin' ContentType='{PRINTER_SETTINGS_TYPE}'/><Override PartName='/vbaProject.bin' ContentType='{VBA_PROJECT_TYPE}'/>"
                ),
            );
            let printer = format!(
                "<Relationship Id='printer' Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/printerSettings' Target='{target}'/>"
            );
            parts.extend([
                ("xl/sheet1.xml".to_owned(), b"<worksheet/>".to_vec()),
                (
                    "xl/_rels/sheet1.xml.rels".to_owned(),
                    relationships_xml(&printer),
                ),
                ("vbaProject.bin".to_owned(), b"vba payload".to_vec()),
            ]);
            for output in sanitize_both(&parts) {
                assert!(!has_part(&output, "vbaProject.bin"));
                let relationships =
                    String::from_utf8_lossy(part_bytes(&output, "xl/_rels/sheet1.xml.rels"));
                assert!(
                    !relationships.contains("vbaProject"),
                    "{target}: {relationships}"
                );
                assert_all_parts_declared(&output);
            }
        }
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
            let package =
                rezip_parts(&package_parts(part_name, content_type, main_xml.as_bytes())).unwrap();
            let sanitized = sanitize_package_for_format(&package, format).unwrap();
            let parts = unzip_parts(&sanitized).unwrap();
            assert_eq!(parts.len(), 3);
            assert!(
                parts
                    .iter()
                    .all(|(_, bytes)| !String::from_utf8_lossy(bytes).contains("DDE secret"))
            );
        }
    }

    #[test]
    fn accepts_vsdx_and_preserves_visio_fields_and_formulas() {
        let mut parts = package_parts(
            "visio/document.xml",
            "application/vnd.ms-visio.drawing.main+xml",
            b"<VisioDocument/>",
        );
        parts.push((
            "visio/pages/page1.xml".to_owned(),
            br#"<PageContents><Shapes><Shape><Cell N='PinX' F='Width*0.5' V='1'/><Text>Page <fld IX='0'/></Text></Shape></Shapes></PageContents>"#.to_vec(),
        ));
        let package = rezip_parts(&parts).unwrap();
        let sanitized = sanitize_package_for_format(&package, "vsdx").unwrap();
        let parts = unzip_parts(&sanitized).unwrap();
        let page = parts
            .iter()
            .find(|(path, _)| path == "visio/pages/page1.xml")
            .unwrap();
        let xml = String::from_utf8_lossy(&page.1);
        assert!(xml.contains("F='Width*0.5'"));
        assert!(xml.contains("<fld IX='0'/>"));
    }

    #[test]
    fn accepts_vstx_without_accepting_macro_templates() {
        for (content_type, accepted) in [
            ("application/vnd.ms-visio.template.main+xml", true),
            (
                "application/vnd.ms-visio.template.macroEnabled.main+xml",
                false,
            ),
            ("application/vnd.ms-visio.drawing.main+xml", false),
        ] {
            let package = rezip_parts(&package_parts(
                "visio/document.xml",
                content_type,
                b"<VisioDocument/>",
            ))
            .unwrap();
            assert_eq!(
                sanitize_package_for_format(&package, "vstx").is_ok(),
                accepted
            );
        }
    }

    #[test]
    fn rejects_recognized_non_drawing_visio_kinds() {
        for content_type in [
            "application/vnd.ms-visio.drawing.macroEnabled.main+xml",
            "application/vnd.ms-visio.stencil.main+xml",
            "application/vnd.ms-visio.stencil.macroEnabled.main+xml",
            "application/vnd.ms-visio.template.main+xml",
            "application/vnd.ms-visio.template.macroEnabled.main+xml",
        ] {
            let package = rezip_parts(&package_parts(
                "visio/document.xml",
                content_type,
                b"<VisioDocument/>",
            ))
            .unwrap();
            assert!(sanitize_package_for_format(&package, "vsdx").is_err());
        }
    }

    #[test]
    fn content_type_is_authoritative_for_visio_kind_detection() {
        let macro_enabled = rezip_parts(&package_parts(
            "visio/document.xml",
            "application/vnd.ms-visio.drawing.macroEnabled.main+xml",
            b"<VisioDocument/>",
        ))
        .unwrap();
        assert!(sanitize_package_for_format(&macro_enabled, "vsdx").is_err());

        let no_content_type = rezip_parts(&[(
            "visio/document.xml".to_owned(),
            b"<VisioDocument/>".to_vec(),
        )])
        .unwrap();
        assert!(sanitize_package_for_format(&no_content_type, "vsdx").is_err());
    }

    #[test]
    fn keeps_existing_format_detection_results() {
        let cases = [
            (
                "word/document.xml",
                "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
                DocumentKind::Docx,
            ),
            (
                "xl/workbook.xml",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
                DocumentKind::Xlsx,
            ),
            (
                "ppt/presentation.xml",
                "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml",
                DocumentKind::Pptx,
            ),
        ];
        for (part_name, content_type, expected) in cases {
            let parts = package_parts(part_name, content_type, b"<root/>");
            assert_eq!(detect_package_kind(&parts), Ok(expected));
            assert_eq!(detect_format(&parts).unwrap(), expected);
        }
    }

    #[test]
    fn preserves_macro_state_in_document_kinds() {
        for (content_type, expected, format) in [
            (
                "application/vnd.ms-word.document.macroEnabled.main+xml",
                DocumentKind::Docm,
                "docm",
            ),
            (
                "application/vnd.ms-excel.sheet.macroEnabled.main+xml",
                DocumentKind::Xlsm,
                "xlsm",
            ),
            (
                "application/vnd.ms-powerpoint.presentation.macroEnabled.main+xml",
                DocumentKind::Pptm,
                "pptm",
            ),
            (
                "application/vnd.ms-visio.drawing.macroEnabled.main+xml",
                DocumentKind::Vsdm,
                "vsdm",
            ),
            (
                "application/vnd.ms-visio.stencil.main+xml",
                DocumentKind::Vssx,
                "vssx",
            ),
            (
                "application/vnd.ms-visio.stencil.macroEnabled.main+xml",
                DocumentKind::Vssm,
                "vssm",
            ),
            (
                "application/vnd.ms-visio.template.main+xml",
                DocumentKind::Vstx,
                "vstx",
            ),
            (
                "application/vnd.ms-visio.template.macroEnabled.main+xml",
                DocumentKind::Vstm,
                "vstm",
            ),
        ] {
            let part_name = if content_type.contains("visio") {
                "visio/document.xml"
            } else if content_type.contains("word") {
                "word/document.xml"
            } else if content_type.contains("excel") {
                "xl/workbook.xml"
            } else {
                "ppt/presentation.xml"
            };
            let parts = package_parts(part_name, content_type, b"<root/>");
            let detected = detect_package_kind(&parts).unwrap();
            assert_eq!(detected, expected);
            assert_eq!(detected.format(), format);
        }
    }

    #[test]
    fn preserves_shipping_format_sanitizer_goldens() {
        // Compared as parts, not container bytes: zip records the writing host in
        // every central directory header, so the same output differs byte for byte
        // between a Windows and a Unix build.
        for (format, source, expected) in [
            (
                "docx",
                include_bytes!("../tests/fixtures/betteroffice-demo.docx").as_slice(),
                include_bytes!("../tests/fixtures/betteroffice-demo.docx.sanitized").as_slice(),
            ),
            (
                "xlsx",
                include_bytes!("../tests/fixtures/sample.xlsx").as_slice(),
                include_bytes!("../tests/fixtures/sample.xlsx.sanitized").as_slice(),
            ),
            (
                "pptx",
                include_bytes!("../tests/fixtures/betteroffice-demo.pptx").as_slice(),
                include_bytes!("../tests/fixtures/betteroffice-demo.pptx.sanitized").as_slice(),
            ),
        ] {
            let sanitized = sanitize_package_for_format(source, format).unwrap();
            assert_eq!(
                crate::unzip_parts(&sanitized).unwrap(),
                crate::unzip_parts(expected).unwrap(),
                "{format} sanitizer output drifted from its golden"
            );
        }
    }

    #[test]
    fn validates_content_types_and_resolves_main_relationship() {
        let macro_enabled = vec![
            (
                "[Content_Types].xml".to_owned(),
                br#"<Types xmlns='http://schemas.openxmlformats.org/package/2006/content-types'><Override PartName='/visio/document.xml' ContentType='application/vnd.ms-visio.drawing.macro&#69;nabled.main+xml'/></Types>"#.to_vec(),
            ),
            main_relationship("visio/document.xml", VISIO_DOCUMENT_RELATIONSHIP),
            ("visio/document.xml".to_owned(), Vec::new()),
        ];
        assert_eq!(detect_package_kind(&macro_enabled), Ok(DocumentKind::Vsdm));

        let comment = vec![
            (
                "[Content_Types].xml".to_owned(),
                br#"<Types xmlns='http://schemas.openxmlformats.org/package/2006/content-types'><!-- application/vnd.ms-visio.drawing.main+xml --><Override PartName='/word/document.xml' ContentType='application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml'/></Types>"#.to_vec(),
            ),
            main_relationship("word/document.xml", OFFICE_DOCUMENT_RELATIONSHIP),
            ("word/document.xml".to_owned(), Vec::new()),
        ];
        assert_eq!(detect_package_kind(&comment), Ok(DocumentKind::Docx));

        let alternate_path = package_parts(
            "drawing/main.data",
            "application/vnd.ms-visio.drawing.main+xml",
            b"<VisioDocument/>",
        );
        assert_eq!(detect_package_kind(&alternate_path), Ok(DocumentKind::Vsdx));

        let alternate_docx = package_parts(
            "custom/main.data",
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
            b"<w:document/>",
        );
        assert_eq!(detect_package_kind(&alternate_docx), Ok(DocumentKind::Docx));
    }

    #[test]
    fn root_relationship_prevents_decoy_kind_spoofing() {
        let parts = vec![
            (
                "[Content_Types].xml".to_owned(),
                br#"<Types xmlns='http://schemas.openxmlformats.org/package/2006/content-types'><Override PartName='/word/document.xml' ContentType='application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml'/><Override PartName='/visio/document.xml' ContentType='application/vnd.ms-visio.drawing.main+xml'/></Types>"#.to_vec(),
            ),
            main_relationship("word/document.xml", OFFICE_DOCUMENT_RELATIONSHIP),
            ("word/document.xml".to_owned(), Vec::new()),
            ("visio/document.xml".to_owned(), Vec::new()),
        ];
        assert_eq!(detect_package_kind(&parts), Ok(DocumentKind::Docx));
        let package = rezip_parts(&parts).unwrap();
        assert!(sanitize_package_for_format(&package, "vsdx").is_err());
    }

    #[test]
    fn rejects_mismatched_main_relationship_family() {
        let parts = vec![
            content_types(
                "visio/document.xml",
                "application/vnd.ms-visio.drawing.main+xml",
            ),
            main_relationship("visio/document.xml", OFFICE_DOCUMENT_RELATIONSHIP),
            ("visio/document.xml".to_owned(), Vec::new()),
        ];
        assert!(matches!(
            detect_package_kind(&parts),
            Err(DocumentKindError::InvalidPackageRelationships(_))
        ));
    }

    #[test]
    fn rejects_conflicting_main_relationships() {
        let parts = vec![
            (
                "[Content_Types].xml".to_owned(),
                br#"<Types xmlns='http://schemas.openxmlformats.org/package/2006/content-types'><Override PartName='/word/document.xml' ContentType='application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml'/><Override PartName='/visio/document.xml' ContentType='application/vnd.ms-visio.drawing.main+xml'/></Types>"#.to_vec(),
            ),
            (
                "_rels/.rels".to_owned(),
                format!(
                    "<Relationships xmlns='{namespace}'><Relationship Id='rId1' Type='{OFFICE_DOCUMENT_RELATIONSHIP}' Target='word/document.xml'/><Relationship Id='rId2' Type='{VISIO_DOCUMENT_RELATIONSHIP}' Target='visio/document.xml'/></Relationships>",
                    namespace = String::from_utf8_lossy(RELATIONSHIPS_NAMESPACE)
                )
                .into_bytes(),
            ),
            ("word/document.xml".to_owned(), Vec::new()),
            ("visio/document.xml".to_owned(), Vec::new()),
        ];
        assert!(matches!(
            detect_package_kind(&parts),
            Err(DocumentKindError::ConflictingMainDocumentRelationships(_))
        ));
    }

    #[test]
    fn ignores_nested_and_foreign_content_type_entries() {
        let nested = vec![
            (
                "[Content_Types].xml".to_owned(),
                br#"<Types xmlns='http://schemas.openxmlformats.org/package/2006/content-types'><Group><Override PartName='/visio/document.xml' ContentType='application/vnd.ms-visio.drawing.main+xml'/></Group></Types>"#.to_vec(),
            ),
            main_relationship("visio/document.xml", VISIO_DOCUMENT_RELATIONSHIP),
            ("visio/document.xml".to_owned(), Vec::new()),
        ];
        assert_eq!(
            detect_package_kind(&nested),
            Err(DocumentKindError::MissingMainDocumentKind)
        );

        let foreign = vec![
            (
                "[Content_Types].xml".to_owned(),
                br#"<Types xmlns='urn:foreign'><Override PartName='/visio/document.xml' ContentType='application/vnd.ms-visio.drawing.main+xml'/></Types>"#.to_vec(),
            ),
            main_relationship("visio/document.xml", VISIO_DOCUMENT_RELATIONSHIP),
            ("visio/document.xml".to_owned(), Vec::new()),
        ];
        assert!(matches!(
            detect_package_kind(&foreign),
            Err(DocumentKindError::InvalidContentTypes(_))
        ));
    }

    #[test]
    fn ignores_foreign_namespaced_direct_content_type_entries() {
        let parts = vec![
            (
                "[Content_Types].xml".to_owned(),
                br#"<Types xmlns='http://schemas.openxmlformats.org/package/2006/content-types' xmlns:evil='urn:evil'><evil:Override PartName='/visio/document.xml' ContentType='application/vnd.ms-visio.drawing.main+xml'/></Types>"#.to_vec(),
            ),
            main_relationship("visio/document.xml", VISIO_DOCUMENT_RELATIONSHIP),
            ("visio/document.xml".to_owned(), Vec::new()),
        ];
        assert_eq!(
            detect_package_kind(&parts),
            Err(DocumentKindError::MissingMainDocumentKind)
        );
    }

    #[test]
    fn accepts_prefixed_content_type_declarations() {
        let parts = vec![
            (
                "[Content_Types].xml".to_owned(),
                br#"<ct:Types xmlns:ct='http://schemas.openxmlformats.org/package/2006/content-types'><ct:Override PartName='/visio/document.xml' ContentType='application/vnd.ms-visio.drawing.main+xml'/></ct:Types>"#.to_vec(),
            ),
            main_relationship("visio/document.xml", VISIO_DOCUMENT_RELATIONSHIP),
            ("visio/document.xml".to_owned(), Vec::new()),
        ];
        assert_eq!(detect_package_kind(&parts), Ok(DocumentKind::Vsdx));
    }

    #[test]
    fn accepts_default_namespaced_content_type_declarations() {
        let parts = vec![
            (
                "[Content_Types].xml".to_owned(),
                br#"<Types xmlns='http://schemas.openxmlformats.org/package/2006/content-types'><Default Extension='data' ContentType='application/vnd.ms-visio.drawing.main+xml'/></Types>"#.to_vec(),
            ),
            main_relationship("custom/main.data", VISIO_DOCUMENT_RELATIONSHIP),
            ("custom/main.data".to_owned(), Vec::new()),
        ];
        assert_eq!(detect_package_kind(&parts), Ok(DocumentKind::Vsdx));
    }

    #[test]
    fn rejects_declared_main_part_that_is_absent() {
        let parts = vec![
            content_types(
                "visio/document.xml",
                "application/vnd.ms-visio.drawing.main+xml",
            ),
            main_relationship("visio/document.xml", VISIO_DOCUMENT_RELATIONSHIP),
        ];
        assert!(matches!(
            detect_package_kind(&parts),
            Err(DocumentKindError::MissingMainDocumentPart(_))
        ));

        let package = rezip_parts(&parts).unwrap();
        assert!(sanitize_package_for_format(&package, "vsdx").is_err());
    }

    #[test]
    fn rejects_foreign_attribute_shadowing() {
        let parts = vec![
            (
                "[Content_Types].xml".to_owned(),
                br#"<Types xmlns='http://schemas.openxmlformats.org/package/2006/content-types' xmlns:evil='urn:evil'><Override evil:PartName='/word/document.xml' evil:ContentType='application/vnd.ms-visio.drawing.main+xml' PartName='/visio/document.xml' ContentType='application/vnd.ms-visio.drawing.macroEnabled.main+xml'/></Types>"#.to_vec(),
            ),
            main_relationship("visio/document.xml", VISIO_DOCUMENT_RELATIONSHIP),
            ("visio/document.xml".to_owned(), Vec::new()),
        ];
        assert_eq!(detect_package_kind(&parts), Ok(DocumentKind::Vsdm));

        let package = rezip_parts(&parts).unwrap();
        assert!(sanitize_package_for_format(&package, "vsdx").is_err());
    }

    #[test]
    fn rejects_duplicate_content_type_mappings() {
        let duplicate_override = vec![
            (
                "[Content_Types].xml".to_owned(),
                br#"<Types xmlns='http://schemas.openxmlformats.org/package/2006/content-types'><Override PartName='/VISIO/document.xml' ContentType='application/vnd.ms-visio.drawing.main+xml'/><Override PartName='/visio/document.xml' ContentType='application/vnd.ms-visio.drawing.macroEnabled.main+xml'/></Types>"#.to_vec(),
            ),
            main_relationship("visio/document.xml", VISIO_DOCUMENT_RELATIONSHIP),
            ("visio/document.xml".to_owned(), Vec::new()),
        ];
        assert!(matches!(
            detect_package_kind(&duplicate_override),
            Err(DocumentKindError::InvalidContentTypes(_))
        ));

        let duplicate_default = vec![
            (
                "[Content_Types].xml".to_owned(),
                br#"<Types xmlns='http://schemas.openxmlformats.org/package/2006/content-types'><Default Extension='XML' ContentType='application/vnd.ms-visio.drawing.main+xml'/><Default Extension='xml' ContentType='application/vnd.ms-visio.drawing.macroEnabled.main+xml'/></Types>"#.to_vec(),
            ),
            main_relationship("visio/document.xml", VISIO_DOCUMENT_RELATIONSHIP),
            ("visio/document.xml".to_owned(), Vec::new()),
        ];
        assert!(matches!(
            detect_package_kind(&duplicate_default),
            Err(DocumentKindError::InvalidContentTypes(_))
        ));
    }

    #[test]
    fn format_specific_sanitization_requires_package_metadata() {
        let package =
            rezip_parts(&[("word/document.xml".to_owned(), b"<w:document/>".to_vec())]).unwrap();
        assert!(sanitize_package_for_format(&package, "docx").is_err());
    }

    #[test]
    fn rejects_truncated_package_metadata() {
        let parts = vec![
            (
                "[Content_Types].xml".to_owned(),
                br#"<Types xmlns='http://schemas.openxmlformats.org/package/2006/content-types'><Override PartName='/visio/document.xml' ContentType='application/vnd.ms-visio.drawing.main+xml'/>"#.to_vec(),
            ),
            main_relationship("visio/document.xml", VISIO_DOCUMENT_RELATIONSHIP),
            ("visio/document.xml".to_owned(), Vec::new()),
        ];
        assert!(matches!(
            detect_package_kind(&parts),
            Err(DocumentKindError::InvalidContentTypes(_))
        ));
    }

    #[test]
    fn rejects_xml_resource_limit_overruns() {
        let nested = format!(
            "<Types xmlns='{}'>{}{}</Types>",
            String::from_utf8_lossy(CONTENT_TYPES_NAMESPACE),
            "<x>".repeat(MAX_XML_DEPTH),
            "</x>".repeat(MAX_XML_DEPTH)
        );
        let mut parts = package_parts(
            "visio/document.xml",
            "application/vnd.ms-visio.drawing.main+xml",
            b"<VisioDocument/>",
        );
        parts[0].1 = nested.into_bytes();
        assert!(matches!(
            detect_package_kind(&parts),
            Err(DocumentKindError::InvalidContentTypes(error)) if error.contains("depth limit")
        ));

        let attributes = (0..MAX_XML_ATTRIBUTES_PER_ELEMENT)
            .map(|index| format!(" a{index}='x'"))
            .collect::<String>();
        parts[0].1 = format!(
            "<Types xmlns='{}'{attributes}/>",
            String::from_utf8_lossy(CONTENT_TYPES_NAMESPACE)
        )
        .into_bytes();
        assert!(matches!(
            detect_package_kind(&parts),
            Err(DocumentKindError::InvalidContentTypes(error)) if error.contains("attribute count limit")
        ));

        let document = format!(
            "{}{}",
            "<x>".repeat(MAX_XML_DEPTH + 1),
            "</x>".repeat(MAX_XML_DEPTH + 1)
        );
        let package = rezip_parts(&package_parts(
            "word/document.xml",
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
            document.as_bytes(),
        ))
        .unwrap();
        assert!(
            sanitize_package_for_format(&package, "docx")
                .unwrap_err()
                .contains("depth limit")
        );
    }

    #[test]
    fn enforces_package_scoped_xml_budgets() {
        let mut event_budget = XmlBudget {
            events: MAX_PACKAGE_XML_EVENTS - 4,
            ..XmlBudget::default()
        };
        validate_xml_limits(b"<x/>", "a.xml", &mut event_budget).unwrap();
        validate_xml_limits(b"<x/>", "b.xml", &mut event_budget).unwrap();
        assert!(
            validate_xml_limits(b"<x/>", "c.xml", &mut event_budget)
                .unwrap_err()
                .contains("package XML event limit")
        );

        let mut attribute_budget = XmlBudget {
            attribute_bytes: MAX_PACKAGE_XML_ATTRIBUTE_BYTES - 2,
            ..XmlBudget::default()
        };
        validate_xml_limits(b"<x a='b'/>", "a.xml", &mut attribute_budget).unwrap();
        assert!(
            validate_xml_limits(b"<x a='b'/>", "b.xml", &mut attribute_budget)
                .unwrap_err()
                .contains("package XML attribute byte limit")
        );

        let mut cache_budget = XmlBudget::default();
        validate_xml_limits(b"<x/>", "cached.xml", &mut cache_budget).unwrap();
        let cached_events = cache_budget.events;
        validate_xml_limits(b"<x/>", "cached.xml", &mut cache_budget).unwrap();
        assert_eq!(cache_budget.events, cached_events);
    }

    #[test]
    fn strips_external_relationship_targets() {
        let mut parts = package_parts(
            "word/document.xml",
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
            b"<w:document/>",
        );
        parts.push((
            "word/_rels/document.xml.rels".to_owned(),
            br#"<Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'><Relationship Id='rId1' Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink' Target='tel:+49123456789' TargetMode=' Internal '/><Relationship Id='rId2' Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink' Target='\\server\share\secret.docx' TargetMode='Internal'/><Relationship Id='rId3' Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink' Target='\/server/share/mixed-a.secret' TargetMode='Internal'/><Relationship Id='rId4' Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink' Target='/\server/share/mixed-b.secret' TargetMode='Internal'/></Relationships>"#.to_vec(),
        ));
        let package = rezip_parts(&parts).unwrap();
        let sanitized = sanitize_package_for_format(&package, "docx").unwrap();
        let all = unzip_parts(&sanitized)
            .unwrap()
            .into_iter()
            .flat_map(|(_, bytes)| bytes)
            .collect::<Vec<_>>();
        assert!(!String::from_utf8_lossy(&all).contains("tel:+49123456789"));
        assert!(!String::from_utf8_lossy(&all).contains(r"\\server\share\secret.docx"));
        assert!(!String::from_utf8_lossy(&all).contains("mixed-a.secret"));
        assert!(!String::from_utf8_lossy(&all).contains("mixed-b.secret"));
    }

    #[test]
    fn sanitizes_xml_parts_by_declared_content_type() {
        let parts = vec![
            (
                "[Content_Types].xml".to_owned(),
                br#"<Types xmlns='http://schemas.openxmlformats.org/package/2006/content-types'><Override PartName='/word/document.xml' ContentType='application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml'/><Override PartName='/word/header1.dat' ContentType='application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml'/></Types>"#.to_vec(),
            ),
            main_relationship("word/document.xml", OFFICE_DOCUMENT_RELATIONSHIP),
            ("word/document.xml".to_owned(), b"<w:document/>".to_vec()),
            (
                "word/header1.dat".to_owned(),
                br#"<w:hdr><w:fldSimple w:instr='DDEAUTO SECRET'/></w:hdr>"#.to_vec(),
            ),
        ];
        let package = rezip_parts(&parts).unwrap();
        let sanitized = sanitize_package_for_format(&package, "docx").unwrap();
        let header = unzip_parts(&sanitized)
            .unwrap()
            .into_iter()
            .find(|(path, _)| path == "word/header1.dat")
            .unwrap();
        assert!(!String::from_utf8_lossy(&header.1).contains("DDEAUTO SECRET"));
    }

    #[test]
    fn rejects_dtd_in_xml_parts() {
        let package = rezip_parts(&package_parts(
            "word/document.xml",
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
            b"<!DOCTYPE x><w:document/>",
        ))
        .unwrap();
        assert!(sanitize_package(&package).unwrap_err().contains("DTD"));
    }
}
