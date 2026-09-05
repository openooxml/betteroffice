use std::collections::{BTreeMap, HashSet};

use quick_xml::events::{BytesCData, BytesStart, BytesText, Event};
use quick_xml::name::ResolveResult;
use quick_xml::{NsReader, Reader, Writer, XmlVersion};

use crate::{rezip_parts, unzip_parts};

const CONTENT_TYPES_NAMESPACE: &[u8] =
    b"http://schemas.openxmlformats.org/package/2006/content-types";
const RELATIONSHIPS_NAMESPACE: &[u8] =
    b"http://schemas.openxmlformats.org/package/2006/relationships";
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
    if !matches!(expected_format, "docx" | "xlsx" | "pptx" | "vsdx") {
        return Err(format!("unsupported OOXML format: {expected_format}"));
    }
    sanitize_package_inner(data, Some(expected_format))
}

fn sanitize_package_inner(data: &[u8], expected_format: Option<&str>) -> Result<Vec<u8>, String> {
    let mut parts = unzip_parts(data)?;
    let mut xml_budget = XmlBudget::default();
    let detected = if expected_format.is_some() {
        detect_package_kind_with_budget(&parts, &mut xml_budget)
            .map_err(|error| error.to_string())?
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

    let content_types = match parse_content_types(&parts, &mut xml_budget) {
        Ok(content_types) => Some(content_types),
        Err(DocumentKindError::MissingContentTypes) => None,
        Err(error) => return Err(error.to_string()),
    };

    let mut removed: HashSet<String> = parts
        .iter()
        .filter(|(path, _)| dangerous_path(path))
        .map(|(path, _)| normalize_part_name(path))
        .collect();
    if let Some((_, content_types)) = parts
        .iter()
        .find(|(path, _)| path.eq_ignore_ascii_case("[Content_Types].xml"))
    {
        removed.extend(dangerous_content_type_parts(
            content_types,
            &mut xml_budget,
        )?);
    }

    parts.retain(|(path, _)| !removed.contains(&normalize_part_name(path)));
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
            *bytes = sanitize_relationships(bytes, &removed, path, &mut xml_budget)?;
        } else if is_xml_part(&lower) || content_type.is_some_and(is_xml_content_type) {
            *bytes = neutralize_fields(bytes, path, &mut xml_budget)?;
        }
    }
    rezip_parts(&parts)
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
}

impl DocumentKind {
    pub fn format(self) -> &'static str {
        match self {
            Self::Docx => "docx",
            Self::Docm => "docm",
            Self::Xlsx => "xlsx",
            Self::Xlsm => "xlsm",
            Self::Pptx => "pptx",
            Self::Pptm => "pptm",
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
            Self::Xlsm => "xlsx",
            Self::Pptm => "pptx",
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
    detect_package_kind_with_budget(parts, &mut xml_budget)
}

fn detect_package_kind_with_budget(
    parts: &[(String, Vec<u8>)],
    xml_budget: &mut XmlBudget,
) -> Result<DocumentKind, DocumentKindError> {
    let part_names: HashSet<String> = parts
        .iter()
        .map(|(path, _)| normalize_part_name(path))
        .collect();
    let main_document = main_document_part(parts, &part_names, xml_budget)?;
    let content_types = parse_content_types(parts, xml_budget)?;
    let content_type = content_types
        .content_type_for(&main_document.part_name)
        .ok_or(DocumentKindError::MissingMainDocumentKind)?;
    let kind =
        declared_main_kind(content_type).ok_or(DocumentKindError::MissingMainDocumentKind)?;
    if main_document.relationship_kind.accepts(kind) {
        Ok(kind)
    } else {
        Err(invalid_relationships(
            "main relationship type does not match document content type",
        ))
    }
}

struct ContentTypes {
    overrides: BTreeMap<String, String>,
    defaults: BTreeMap<String, String>,
}

impl ContentTypes {
    fn content_type_for(&self, part_name: &str) -> Option<&str> {
        let part_name = normalize_part_name(part_name);
        if let Some(content_type) = self.overrides.get(&part_name) {
            return Some(content_type);
        }
        let (_, extension) = part_name.rsplit_once('.')?;
        self.defaults
            .get(&extension.to_ascii_lowercase())
            .map(String::as_str)
    }
}

fn parse_content_types(
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
    let mut declaration_count = 0_usize;
    let mut depth = 0_usize;
    let mut root_seen = false;
    loop {
        match reader
            .read_resolved_event()
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
    })
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
            let part_name = normalize_part_name(part_name);
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
                    | DocumentKind::Xlsx
                    | DocumentKind::Xlsm
                    | DocumentKind::Pptx
                    | DocumentKind::Pptm
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
        match reader
            .read_resolved_event()
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
    let target = resolve_relationship_target("_rels/.rels", target)
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
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml" => {
            Some(DocumentKind::Xlsx)
        }
        "application/vnd.ms-excel.sheet.macroenabled.main+xml" => Some(DocumentKind::Xlsm),
        "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml" => {
            Some(DocumentKind::Pptx)
        }
        "application/vnd.ms-powerpoint.presentation.macroenabled.main+xml" => {
            Some(DocumentKind::Pptm)
        }
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
    detect_format_with_budget(parts, &mut xml_budget)
}

fn detect_format_with_budget(
    parts: &[(String, Vec<u8>)],
    xml_budget: &mut XmlBudget,
) -> Result<DocumentKind, String> {
    match detect_package_kind_with_budget(parts, xml_budget) {
        Ok(kind) => return Ok(kind),
        Err(error @ DocumentKindError::InvalidContentTypes(_))
        | Err(error @ DocumentKindError::InvalidPackageRelationships(_))
        | Err(error @ DocumentKindError::ConflictingMainDocumentRelationships(_)) => {
            return Err(error.to_string());
        }
        Err(_) => {}
    }
    if has_part(parts, "word/document.xml") {
        Ok(DocumentKind::Docx)
    } else if has_part(parts, "xl/workbook.xml") {
        Ok(DocumentKind::Xlsx)
    } else if has_part(parts, "ppt/presentation.xml") {
        Ok(DocumentKind::Pptx)
    } else {
        Err("could not detect DOCX, XLSX, PPTX, or VSDX package".to_owned())
    }
}

fn has_part(parts: &[(String, Vec<u8>)], expected: &str) -> bool {
    parts
        .iter()
        .any(|(path, _)| path.eq_ignore_ascii_case(expected))
}

fn dangerous_content_type_parts(
    xml: &[u8],
    xml_budget: &mut XmlBudget,
) -> Result<HashSet<String>, String> {
    validate_xml_limits(xml, "[Content_Types].xml", xml_budget)?;
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
    xml_budget: &mut XmlBudget,
) -> Result<Vec<u8>, String> {
    validate_xml_limits(xml, path, xml_budget)?;
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
    xml_budget: &mut XmlBudget,
) -> Result<Vec<u8>, String> {
    validate_xml_limits(xml, path, xml_budget)?;
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
    target_mode.trim().eq_ignore_ascii_case("External")
        || external_target(target)
        || dangerous_relationship_type(relationship_type)
        || dangerous_path(target)
        || resolve_relationship_target(relationship_path, target)
            .is_some_and(|target| removed.contains(&target))
}

fn neutralize_fields(
    xml: &[u8],
    path: &str,
    xml_budget: &mut XmlBudget,
) -> Result<Vec<u8>, String> {
    validate_xml_limits(xml, path, xml_budget)?;
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
        .find(|(name, _)| name == expected)
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

#[derive(Default)]
struct XmlBudget {
    events: usize,
    attribute_bytes: usize,
    validated_parts: HashSet<String>,
}

fn validate_xml_limits(xml: &[u8], path: &str, budget: &mut XmlBudget) -> Result<(), String> {
    let part_name = normalize_part_name(path);
    if budget.validated_parts.contains(&part_name) {
        return Ok(());
    }
    let mut reader = Reader::from_reader(xml);
    let mut depth = 0_usize;
    loop {
        let event = reader
            .read_event()
            .map_err(|error| format!("invalid XML in {path}: {error}"))?;
        budget.events = budget
            .events
            .checked_add(1)
            .ok_or_else(|| format!("package XML event limit exceeded in {path}"))?;
        if budget.events > MAX_PACKAGE_XML_EVENTS {
            return Err(format!("package XML event limit exceeded in {path}"));
        }
        match event {
            Event::Start(start) => {
                validate_xml_attributes(&start, path, budget)?;
                depth += 1;
                if depth > MAX_XML_DEPTH {
                    return Err(format!("XML depth limit exceeded in {path}"));
                }
            }
            Event::Empty(start) => {
                validate_xml_attributes(&start, path, budget)?;
                if depth >= MAX_XML_DEPTH {
                    return Err(format!("XML depth limit exceeded in {path}"));
                }
            }
            Event::End(_) => {
                if depth == 0 {
                    return Err(format!("unexpected closing element in {path}"));
                }
                depth -= 1;
            }
            Event::DocType(_) => return Err(format!("DTD is forbidden in {path}")),
            Event::Eof => {
                if depth != 0 {
                    return Err(format!("unexpected EOF in {path}"));
                }
                budget.validated_parts.insert(part_name);
                return Ok(());
            }
            _ => {}
        }
    }
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
