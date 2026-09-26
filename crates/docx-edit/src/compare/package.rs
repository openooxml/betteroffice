//! Package inspection: content types, tracked changes, inspection limits and parts outside the
//! main document.

use std::collections::{BTreeSet, HashMap};

use docx_parse::{XmlElement, XmlNode};

use super::xml::{Namespaces, canonical, is_revision, w_text_units};
use super::{CompareDiagnosticCode, CompareInput, Diagnostics, Package, Stop, location};
use crate::structured::Severity;
use crate::structured::source::relationship_part;

const CORE: &str = "application/vnd.openxmlformats-package.core-properties+xml";
const EXTENDED: &str = "application/vnd.openxmlformats-officedocument.extended-properties+xml";
const WORD: &str = "application/vnd.openxmlformats-officedocument.wordprocessingml.";

pub(crate) fn parse(bytes: &[u8], part: &str) -> Option<docx_parse::XmlDocument> {
    let limits = docx_parse::ParseLimits::default();
    docx_parse::parse_xml(bytes, part, &mut docx_parse::ParseBudget::new(&limits)).ok()
}

/// Each part's content type by `[Content_Types].xml`, or `None` when that part is unusable.
pub(crate) fn content_types(parts: &[(String, Vec<u8>)]) -> Option<HashMap<String, String>> {
    let (_, bytes) = parts
        .iter()
        .find(|(path, _)| path == "[Content_Types].xml")?;
    let document = parse(bytes, "[Content_Types].xml")?;
    let root = document.root()?;
    let mut defaults = HashMap::new();
    let mut overrides = HashMap::new();
    for child in root.child_elements() {
        let content_type = child.attribute(None, "ContentType");
        match (child.local_name(), content_type) {
            ("Default", Some(content_type)) => {
                let extension = child.attribute(None, "Extension")?.to_ascii_lowercase();
                defaults.insert(extension, content_type.to_owned());
            }
            ("Override", Some(content_type)) => {
                let name = child.attribute(None, "PartName")?;
                let name = name.strip_prefix('/').unwrap_or(name).to_ascii_lowercase();
                overrides.insert(name, content_type.to_owned());
            }
            _ => {}
        }
    }
    let mut types = HashMap::with_capacity(parts.len());
    for (path, _) in parts {
        if path == "[Content_Types].xml" {
            continue;
        }
        let lower = path.to_ascii_lowercase();
        let extension = lower.rsplit_once('.').map(|(_, extension)| extension);
        let content_type = overrides
            .get(&lower)
            .or_else(|| extension.and_then(|extension| defaults.get(extension)))?;
        types.insert(path.clone(), content_type.clone());
    }
    Some(types)
}

pub(crate) fn is_xml(content_type: &str) -> bool {
    let content_type = content_type.to_ascii_lowercase();
    content_type.ends_with("+xml")
        || content_type == "application/xml"
        || content_type == "text/xml"
}

/// Paragraphs and text an inspection may still read: per input and combined.
pub(crate) struct InspectionBudget {
    pub paragraphs: usize,
    pub max_paragraphs: usize,
    pub text_units: usize,
}

/// Whether `bytes` open like an XML document.
fn looks_like_xml(bytes: &[u8]) -> bool {
    let bytes = bytes.strip_prefix(b"\xEF\xBB\xBF").unwrap_or(bytes);
    bytes.iter().find(|byte| !byte.is_ascii_whitespace()) == Some(&b'<')
}

/// Parses every part of `package` that its content type or its content says is XML, refusing
/// tracked changes and unreadable XML-typed parts and charging paragraphs (per input) and text
/// (combined) to `budget` as it reads.
pub(crate) fn scan(
    package: &Package,
    budget: &mut InspectionBudget,
    diagnostics: &mut Diagnostics,
) -> Result<(), Stop> {
    budget.paragraphs = budget.max_paragraphs;
    let name = package.name();
    for (path, bytes) in &package.parts {
        let typed = package
            .content_types
            .get(path)
            .is_some_and(|content_type| is_xml(content_type));
        if !typed && !looks_like_xml(bytes) {
            continue;
        }
        let Some(document) = parse(bytes, path) else {
            if !typed {
                continue;
            }
            diagnostics.block(
                CompareDiagnosticCode::InvalidDocx,
                format!("{path} in the {name} document is not readable XML"),
                vec![location(package.input, Some(path), None)],
            )?;
            return Err(Stop);
        };
        let Some(root) = document.root() else {
            continue;
        };
        let mut walk = Walk {
            ns: Namespaces::default(),
            path: Vec::new(),
            revision: None,
            budget,
        };
        if walk.visit(root).is_err() {
            return Err(diagnostics.exhausted(format!(
                "the {name} document holds more paragraphs or text than maxParagraphs and maxTextUnits allow"
            )));
        }
        if let Some((element_path, element)) = walk.revision {
            diagnostics.block(
                CompareDiagnosticCode::ExistingRevisions,
                format!(
                    "{path} in the {name} document already holds tracked changes (w:{element}); accept or reject them first"
                ),
                vec![location(package.input, Some(path), Some(element_path))],
            )?;
        }
    }
    Ok(())
}

struct Walk<'a> {
    ns: Namespaces,
    path: Vec<u32>,
    revision: Option<(Vec<u32>, String)>,
    budget: &'a mut InspectionBudget,
}

struct Exhausted;

impl Walk<'_> {
    fn visit(&mut self, element: &XmlElement) -> Result<(), Exhausted> {
        let mark = self.ns.enter(element);
        if let Some(local) = self.ns.w_local(element) {
            if self.revision.is_none() && is_revision(local) {
                self.revision = Some((self.path.clone(), local.to_owned()));
            }
            if local == "p" {
                self.budget.paragraphs = self.budget.paragraphs.checked_sub(1).ok_or(Exhausted)?;
            }
            if let Some(units) = w_text_units(local, element) {
                self.budget.text_units =
                    self.budget.text_units.checked_sub(units).ok_or(Exhausted)?;
            }
        }
        let mut index = 0u32;
        for child in &element.children {
            if let XmlNode::Element(child) = child {
                self.path.push(index);
                let result = self.visit(child);
                self.path.pop();
                result?;
                index += 1;
            }
        }
        self.ns.leave(mark);
        Ok(())
    }
}

/// Bookkeeping a comparison ignores in a part of `content_type`, as `(namespace, local)` names
/// of direct children of its root.
fn bookkeeping(content_type: &str) -> Option<&'static [(&'static str, &'static str)]> {
    const CORE_FIELDS: &[(&str, &str)] = &[
        ("http://purl.org/dc/terms/", "modified"),
        (
            "http://schemas.openxmlformats.org/package/2006/metadata/core-properties",
            "lastModifiedBy",
        ),
        (
            "http://schemas.openxmlformats.org/package/2006/metadata/core-properties",
            "revision",
        ),
        (
            "http://schemas.openxmlformats.org/package/2006/metadata/core-properties",
            "lastPrinted",
        ),
    ];
    const EXTENDED_NS: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties";
    const EXTENDED_FIELDS: &[(&str, &str)] = &[
        (EXTENDED_NS, "TotalTime"),
        (EXTENDED_NS, "Pages"),
        (EXTENDED_NS, "Words"),
        (EXTENDED_NS, "Characters"),
        (EXTENDED_NS, "CharactersWithSpaces"),
        (EXTENDED_NS, "Lines"),
        (EXTENDED_NS, "Paragraphs"),
    ];
    const SETTINGS_FIELDS: &[(&str, &str)] = &[
        (
            "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
            "rsids",
        ),
        ("http://purl.oclc.org/ooxml/wordprocessingml/main", "rsids"),
    ];
    match content_type {
        CORE => Some(CORE_FIELDS),
        EXTENDED => Some(EXTENDED_FIELDS),
        _ if content_type == format!("{WORD}settings+xml") => Some(SETTINGS_FIELDS),
        _ => None,
    }
}

/// The canonical part with the direct children `ignored` names left out.
fn canonical_without(root: &XmlElement, ignored: &[(&str, &str)]) -> String {
    let mut ns = Namespaces::default();
    let mark = ns.enter(root);
    let mut filtered = root.clone();
    filtered.children.retain(|child| match child {
        XmlNode::Element(element) => {
            let mark = ns.enter(element);
            let (uri, local) = ns.resolve(&element.name, false);
            let keep = !ignored
                .iter()
                .any(|(namespace, name)| uri == Some(*namespace) && local == *name);
            ns.leave(mark);
            keep
        }
        _ => true,
    });
    ns.leave(mark);
    let mut out = String::new();
    canonical(&filtered, &mut Namespaces::default(), &mut out);
    out
}

fn classify(
    path: &str,
    content_type: Option<&str>,
    document_part: &str,
) -> (CompareDiagnosticCode, &'static str) {
    let kind = content_type
        .and_then(|content_type| content_type.strip_prefix(WORD))
        .unwrap_or_default();
    if matches!(
        kind,
        "header+xml"
            | "footer+xml"
            | "footnotes+xml"
            | "endnotes+xml"
            | "comments+xml"
            | "commentsExtended+xml"
            | "commentsIds+xml"
            | "commentsExtensible+xml"
            | "people+xml"
            | "document.glossary+xml"
    ) {
        (
            CompareDiagnosticCode::OutOfScopeChange,
            "content outside the body is not compared",
        )
    } else if matches!(
        kind,
        "styles+xml" | "stylesWithEffects+xml" | "numbering+xml" | "fontTable+xml"
    ) || content_type == Some("application/vnd.openxmlformats-officedocument.theme+xml")
    {
        (
            CompareDiagnosticCode::FormattingChange,
            "shared formatting definitions differ",
        )
    } else if path == relationship_part(document_part) {
        (
            CompareDiagnosticCode::StructureChange,
            "the main document's relationships differ",
        )
    } else {
        (
            CompareDiagnosticCode::OpaquePartChange,
            "a package part differs and is not compared",
        )
    }
}

/// Diagnoses every difference between the packages outside the main document part.
pub(crate) fn compare_parts(
    original: &Package,
    revised: &Package,
    diagnostics: &mut Diagnostics,
) -> Result<(), Stop> {
    fn index(package: &Package) -> HashMap<&str, &[u8]> {
        package
            .parts
            .iter()
            .map(|(path, bytes)| (path.as_str(), bytes.as_slice()))
            .collect()
    }
    let (before, after) = (index(original), index(revised));
    let paths: BTreeSet<&str> = before.keys().chain(after.keys()).copied().collect();
    let mut metadata = Vec::new();
    for path in paths {
        let (left, right) = (before.get(path).copied(), after.get(path).copied());
        if path == original.document_part && path == revised.document_part {
            continue;
        }
        if left == right {
            continue;
        }
        let content_type = original
            .content_types
            .get(path)
            .or_else(|| revised.content_types.get(path))
            .map(String::as_str);
        let same_type = original.content_types.get(path) == revised.content_types.get(path);
        if let (Some(left), Some(right), Some(ignored), true) =
            (left, right, content_type.and_then(bookkeeping), same_type)
        {
            let parsed = parse(left, path).zip(parse(right, path));
            if let Some((left, right)) = parsed
                && let (Some(left), Some(right)) = (left.root(), right.root())
                && canonical_without(left, ignored) == canonical_without(right, ignored)
            {
                metadata.push(path.to_owned());
                continue;
            }
        }
        let (code, reason) = classify(path, content_type, &original.document_part);
        let message = match (left, right) {
            (None, _) => format!("{path} exists only in the revised document: {reason}"),
            (_, None) => format!("{path} exists only in the original document: {reason}"),
            _ => format!("{path} differs: {reason}"),
        };
        let mut locations = Vec::new();
        if left.is_some() {
            locations.push(location(CompareInput::Original, Some(path), None));
        }
        if right.is_some() {
            locations.push(location(CompareInput::Revised, Some(path), None));
        }
        diagnostics.block(code, message, locations)?;
    }
    if original.document_part != revised.document_part {
        diagnostics.block(
            CompareDiagnosticCode::StructureChange,
            "the main document parts have different names",
            vec![
                location(CompareInput::Original, Some(&original.document_part), None),
                location(CompareInput::Revised, Some(&revised.document_part), None),
            ],
        )?;
    }
    for path in metadata {
        diagnostics.push(
            CompareDiagnosticCode::MetadataDifference,
            Severity::Info,
            format!("{path} differs only in bookkeeping; the original's is kept"),
            vec![location(CompareInput::Revised, Some(&path), None)],
        )?;
    }
    Ok(())
}
