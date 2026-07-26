use std::collections::{BTreeMap, HashMap, HashSet};

use quick_xml::events::{BytesEnd, BytesStart, Event};
use quick_xml::name::ResolveResult;
use quick_xml::{NsReader, Reader, Writer};
use xlsx_model::Workbook;

use crate::read::SharedStringCells;
use crate::xml::{attr, find_part, local_name, next_event, reader, resolve_part_path, xml_err};
use crate::{MAX_DEPTH, ParseError};

/// Source package data retained for ownership-aware writes.
#[derive(Clone, Debug)]
pub struct PreservedPackage {
    pub(crate) parts: Vec<(String, Vec<u8>)>,
    pub(crate) content_types: Vec<ContentTypeEntry>,
    pub(crate) root_relationships: Vec<Relationship>,
    pub(crate) workbook_relationships: Vec<Relationship>,
    pub(crate) workbook_template: XmlTemplate,
    pub(crate) workbook_pr_attributes: Option<Vec<XmlAttribute>>,
    pub(crate) workbook_sheets_attributes: Vec<XmlAttribute>,
    pub(crate) calc_pr_attributes: Option<Vec<XmlAttribute>>,
    pub(crate) sheets: Vec<PreservedSheet>,
    pub(crate) shared_strings: Option<PartReference>,
    pub(crate) shared_strings_template: Option<XmlTemplate>,
    pub(crate) styles: Option<PartReference>,
    pub(crate) theme: Option<PartReference>,
    pub(crate) calc_chains: Vec<PartReference>,
    pub(crate) stylesheet_template: Option<XmlTemplate>,
    pub(crate) original_workbook: Workbook,
}

impl PreservedPackage {
    pub(crate) fn capture(
        parts: &[(String, Vec<u8>)],
        workbook: &Workbook,
        shared_string_cells: &[SharedStringCells],
    ) -> Result<Self, ParseError> {
        let workbook_xml = find_part(parts, "xl/workbook.xml")
            .ok_or_else(|| ParseError::MissingPart("xl/workbook.xml".into()))?;
        let workbook_template = XmlTemplate::capture(workbook_xml)?;
        let workbook_pr_attributes = workbook_template
            .child("workbookPr")
            .map(|child| attributes_from_fragment(&child.bytes))
            .transpose()?;
        let workbook_sheets_attributes = workbook_template
            .child("sheets")
            .map(|child| attributes_from_fragment(&child.bytes))
            .transpose()?
            .unwrap_or_default();
        let calc_pr_attributes = workbook_template
            .child("calcPr")
            .map(|child| attributes_from_fragment(&child.bytes))
            .transpose()?;

        let workbook_relationships = find_part(parts, "xl/_rels/workbook.xml.rels")
            .map(parse_relationships)
            .transpose()?
            .unwrap_or_default();
        let relationship_by_id = workbook_relationships
            .iter()
            .filter_map(|relationship| relationship.id().map(|id| (id.to_owned(), relationship)))
            .collect::<HashMap<_, _>>();

        let sheet_entries = parse_sheet_entries(workbook_xml)?;
        let mut sheets = Vec::with_capacity(sheet_entries.len());
        for (index, entry) in sheet_entries.into_iter().enumerate() {
            let relationship = entry
                .relationship_id
                .as_deref()
                .and_then(|id| relationship_by_id.get(id).copied());
            let path = relationship
                .and_then(Relationship::target)
                .map(|target| resolve_part_path("xl", target))
                .unwrap_or_else(|| format!("xl/worksheets/sheet{}.xml", index + 1));
            let bytes =
                find_part(parts, &path).ok_or_else(|| ParseError::MissingPart(path.clone()))?;
            let relationship_type = relationship
                .and_then(|relationship| relationship.attribute("Type"))
                .map(str::to_owned);
            sheets.push(PreservedSheet {
                path,
                relationship_id: entry.relationship_id,
                relationship_type,
                sheet_id: entry.sheet_id.unwrap_or((index + 1) as u32),
                attributes: entry.attributes,
                template: XmlTemplate::capture(bytes)?,
                shared_string_cells: shared_string_cells.get(index).cloned().unwrap_or_default(),
            });
        }

        let shared_strings = part_reference(
            parts,
            &workbook_relationships,
            "sharedStrings",
            "xl/sharedStrings.xml",
        );
        let styles = part_reference(parts, &workbook_relationships, "styles", "xl/styles.xml");
        let theme = part_reference(
            parts,
            &workbook_relationships,
            "theme",
            "xl/theme/theme1.xml",
        );
        let shared_strings_template = shared_strings
            .as_ref()
            .and_then(|part| find_part(parts, &part.path))
            .map(XmlTemplate::capture)
            .transpose()?;
        let mut calc_chains = workbook_relationships
            .iter()
            .filter(|relationship| relationship.has_type("calcChain"))
            .filter_map(|relationship| {
                relationship.target().map(|target| PartReference {
                    path: resolve_part_path("xl", target),
                    relationship_id: relationship.id().map(str::to_owned),
                })
            })
            .collect::<Vec<_>>();
        if calc_chains.is_empty() && find_part(parts, "xl/calcChain.xml").is_some() {
            calc_chains.push(PartReference {
                path: "xl/calcChain.xml".to_owned(),
                relationship_id: None,
            });
        }
        let stylesheet_template = styles
            .as_ref()
            .and_then(|part| find_part(parts, &part.path))
            .map(XmlTemplate::capture)
            .transpose()?;

        Ok(Self {
            parts: parts.to_vec(),
            content_types: find_part(parts, "[Content_Types].xml")
                .map(parse_content_types)
                .transpose()?
                .unwrap_or_default(),
            root_relationships: find_part(parts, "_rels/.rels")
                .map(parse_relationships)
                .transpose()?
                .unwrap_or_default(),
            workbook_relationships,
            workbook_template,
            workbook_pr_attributes,
            workbook_sheets_attributes,
            calc_pr_attributes,
            sheets,
            shared_strings,
            shared_strings_template,
            styles,
            theme,
            calc_chains,
            stylesheet_template,
            original_workbook: workbook.clone(),
        })
    }

    /// Returns retained bytes for a package part.
    pub fn part_bytes(&self, path: &str) -> Option<&[u8]> {
        find_part(&self.parts, path)
    }

    #[doc(hidden)]
    pub fn source_sheet_count(&self) -> usize {
        self.sheets.len()
    }

    /// A preserved part that names sheets or addresses this crate never
    /// patches, so a rename, removal or axis edit would strand it.
    #[doc(hidden)]
    pub fn unpatchable_reference_part(&self) -> Option<&str> {
        const REFERENCE_BEARING: [&str; 3] = ["xl/charts/", "xl/pivottables/", "xl/pivotcache/"];
        self.parts.iter().find_map(|(path, _)| {
            let normalized = path.trim_start_matches('/').to_ascii_lowercase();
            REFERENCE_BEARING
                .iter()
                .any(|prefix| normalized.starts_with(prefix))
                .then_some(path.as_str())
        })
    }

    /// The shared-string entries a source sheet's cells were authored against.
    #[doc(hidden)]
    pub fn source_shared_string_cells(&self, index: usize) -> SharedStringCells {
        self.sheets
            .get(index)
            .map(|sheet| sheet.shared_string_cells.clone())
            .unwrap_or_default()
    }

    /// False for chartsheets, dialogsheets and any other non-worksheet sheet.
    #[doc(hidden)]
    pub fn source_sheet_is_worksheet(&self, index: usize) -> bool {
        self.sheets
            .get(index)
            .is_none_or(PreservedSheet::is_worksheet)
    }
}

#[derive(Clone, Debug)]
pub(crate) struct PreservedSheet {
    pub(crate) path: String,
    pub(crate) relationship_id: Option<String>,
    pub(crate) relationship_type: Option<String>,
    pub(crate) sheet_id: u32,
    pub(crate) attributes: Vec<XmlAttribute>,
    pub(crate) template: XmlTemplate,
    pub(crate) shared_string_cells: SharedStringCells,
}

impl PreservedSheet {
    pub(crate) fn is_worksheet(&self) -> bool {
        self.relationship_type
            .as_deref()
            .and_then(|relationship_type| relationship_type.rsplit('/').next())
            .is_none_or(|kind| kind == "worksheet")
    }
}

#[derive(Clone, Debug)]
pub(crate) struct PartReference {
    pub(crate) path: String,
    pub(crate) relationship_id: Option<String>,
}

#[derive(Clone, Debug)]
pub(crate) struct Relationship {
    pub(crate) attributes: Vec<XmlAttribute>,
}

impl Relationship {
    pub(crate) fn id(&self) -> Option<&str> {
        self.attribute("Id")
    }

    pub(crate) fn target(&self) -> Option<&str> {
        self.attribute("Target")
    }

    pub(crate) fn is_hyperlink(&self) -> bool {
        self.has_type("hyperlink")
    }

    pub(crate) fn has_type(&self, suffix: &str) -> bool {
        self.attribute("Type")
            .and_then(|value| value.rsplit('/').next())
            == Some(suffix)
    }

    pub(crate) fn attribute(&self, local_name: &str) -> Option<&str> {
        self.attributes
            .iter()
            .find(|attribute| attribute.local_name() == local_name)
            .map(|attribute| attribute.value.as_str())
    }
}

#[derive(Clone, Debug)]
pub(crate) struct ContentTypeEntry {
    pub(crate) element: String,
    pub(crate) attributes: Vec<XmlAttribute>,
}

impl ContentTypeEntry {
    pub(crate) fn attribute(&self, local_name: &str) -> Option<&str> {
        self.attributes
            .iter()
            .find(|attribute| attribute.local_name() == local_name)
            .map(|attribute| attribute.value.as_str())
    }
}

#[derive(Clone, Debug)]
pub(crate) struct XmlAttribute {
    pub(crate) name: String,
    pub(crate) value: String,
}

impl XmlAttribute {
    pub(crate) fn local_name(&self) -> &str {
        self.name.rsplit(':').next().unwrap_or(&self.name)
    }
}

#[derive(Clone, Debug)]
pub(crate) struct XmlTemplate {
    prefix: Vec<u8>,
    pub(crate) children: Vec<XmlChild>,
    trailing: Vec<u8>,
    suffix: Vec<u8>,
    root_name: String,
    root_prefix: Option<String>,
    root_namespace: Option<String>,
    namespaces: Vec<(Option<String>, String)>,
}

#[derive(Clone, Debug)]
pub(crate) struct XmlChild {
    before: Vec<u8>,
    pub(crate) local_name: String,
    pub(crate) bytes: Vec<u8>,
    namespace: Option<String>,
}

impl XmlTemplate {
    pub(crate) fn capture(data: &[u8]) -> Result<Self, ParseError> {
        let mut reader = NsReader::from_reader(data);
        let config = reader.config_mut();
        config.expand_empty_elements = false;
        config.check_end_names = true;

        let mut depth = 0_usize;
        let mut root_start = None;
        let mut root_end = None;
        let mut root_name = None;
        let mut root_prefix = None;
        let mut root_namespace = None;
        let mut namespaces = Vec::new();
        let mut active_child: Option<(usize, String, Option<String>)> = None;
        let mut spans = Vec::new();

        loop {
            let before = reader.buffer_position() as usize;
            let (namespace, event) = reader.read_resolved_event().map_err(xml_err)?;
            let namespace = resolved_namespace(namespace);
            let after = reader.buffer_position() as usize;
            match event {
                Event::Start(element) => {
                    if depth == 0 {
                        root_start = Some(after);
                        let name = String::from_utf8_lossy(element.name().as_ref()).into_owned();
                        root_prefix = name.rsplit_once(':').map(|(prefix, _)| prefix.to_owned());
                        root_name = Some(name);
                        root_namespace = namespace.clone();
                        namespaces = namespace_bindings(&element)?;
                    } else if depth == 1 {
                        active_child = Some((
                            before,
                            String::from_utf8_lossy(element.name().local_name().as_ref())
                                .into_owned(),
                            namespace,
                        ));
                    }
                    depth += 1;
                    if depth > MAX_DEPTH {
                        return Err(ParseError::DepthExceeded);
                    }
                }
                Event::Empty(element) if depth == 0 => {
                    return Self::empty_root(data, &element, namespace, before, after);
                }
                Event::Empty(element) if depth == 1 => {
                    spans.push((
                        before,
                        after,
                        String::from_utf8_lossy(element.name().local_name().as_ref()).into_owned(),
                        namespace,
                    ));
                }
                Event::End(_) if depth == 1 => {
                    root_end = Some(before);
                    break;
                }
                Event::End(_) => {
                    depth = depth.saturating_sub(1);
                    if depth == 1
                        && let Some((start, name, namespace)) = active_child.take()
                    {
                        spans.push((start, after, name, namespace));
                    }
                }
                Event::Eof => break,
                _ => {}
            }
        }

        let root_start =
            root_start.ok_or_else(|| ParseError::Malformed("xml part has no root".into()))?;
        let root_end =
            root_end.ok_or_else(|| ParseError::Malformed("xml root is not closed".into()))?;
        let mut cursor = root_start;
        let mut children = Vec::with_capacity(spans.len());
        for (start, end, local_name, namespace) in spans {
            if start < cursor || end > root_end {
                return Err(ParseError::Malformed("overlapping xml child spans".into()));
            }
            children.push(XmlChild {
                before: data[cursor..start].to_vec(),
                local_name,
                bytes: data[start..end].to_vec(),
                namespace,
            });
            cursor = end;
        }

        Ok(Self {
            prefix: data[..root_start].to_vec(),
            children,
            trailing: data[cursor..root_end].to_vec(),
            suffix: data[root_end..].to_vec(),
            root_name: root_name
                .ok_or_else(|| ParseError::Malformed("xml part has no root name".into()))?,
            root_prefix,
            root_namespace,
            namespaces,
        })
    }

    /// A schema-valid self-closing root (`<sst/>`) becomes an empty template
    /// whose expanded start and end tags bracket the insertion point.
    fn empty_root(
        data: &[u8],
        element: &BytesStart<'_>,
        namespace: Option<String>,
        before: usize,
        after: usize,
    ) -> Result<Self, ParseError> {
        let root_name = String::from_utf8_lossy(element.name().as_ref()).into_owned();
        let mut writer = Writer::new(data[..before].to_vec());
        writer
            .write_event(Event::Start(element.clone()))
            .map_err(xml_err)?;
        let mut suffix = format!("</{root_name}>").into_bytes();
        suffix.extend_from_slice(&data[after..]);
        Ok(Self {
            prefix: writer.into_inner(),
            children: Vec::new(),
            trailing: Vec::new(),
            suffix,
            root_prefix: root_name
                .rsplit_once(':')
                .map(|(prefix, _)| prefix.to_owned()),
            root_name,
            root_namespace: namespace,
            namespaces: namespace_bindings(element)?,
        })
    }

    /// Replaces modeled children in place; children outside the root namespace
    /// (`mc:AlternateContent` and friends) keep their position and are never
    /// used as ordering anchors.
    pub(crate) fn render(
        &self,
        replacements: Vec<(&'static str, Option<Vec<u8>>)>,
        rank: fn(&str) -> usize,
    ) -> Result<Vec<u8>, ParseError> {
        let mut replacements = replacements
            .into_iter()
            .map(|(name, bytes)| (name.to_owned(), bytes))
            .collect::<BTreeMap<_, _>>();
        for bytes in replacements.values_mut().flatten() {
            *bytes = self.qualify_fragment(bytes)?;
        }
        let source_names = self
            .children
            .iter()
            .filter(|child| self.is_root_child(child))
            .map(|child| child.local_name.as_str())
            .collect::<HashSet<_>>();
        let mut child_ranks = Vec::with_capacity(self.children.len());
        for child in &self.children {
            if self.is_root_child(child) {
                child_ranks.push(rank(&child.local_name));
                continue;
            }
            let covered = self.alternate_content_names(child)?;
            if let Some(owned) = covered
                .iter()
                .find(|name| replacements.contains_key(name.as_str()))
            {
                return Err(ParseError::UnsupportedEdit(format!(
                    "{owned} lives inside an mc:AlternateContent branch"
                )));
            }
            child_ranks.push(
                covered
                    .iter()
                    .map(|name| rank(name))
                    .min()
                    .unwrap_or(usize::MAX),
            );
        }
        let last_ranked_child = child_ranks.iter().rposition(|rank| *rank != usize::MAX);
        let mut emitted = HashSet::new();
        let mut output = Vec::new();
        output.extend_from_slice(&self.prefix);

        for (index, child) in self.children.iter().enumerate() {
            output.extend_from_slice(&child.before);
            let modeled = self.is_root_child(child);
            let child_rank = child_ranks[index];
            if child_rank != usize::MAX {
                let mut pending = replacements
                    .iter()
                    .filter(|(name, _)| {
                        !source_names.contains(name.as_str())
                            && !emitted.contains(name.as_str())
                            && rank(name) < child_rank
                    })
                    .collect::<Vec<_>>();
                pending.sort_by_key(|(name, _)| rank(name));
                for (name, bytes) in pending {
                    emitted.insert(name.clone());
                    if let Some(bytes) = bytes {
                        output.extend_from_slice(bytes);
                    }
                }
            }

            if modeled && let Some(bytes) = replacements.get(&child.local_name) {
                if emitted.insert(child.local_name.clone())
                    && let Some(bytes) = bytes
                {
                    output.extend_from_slice(bytes);
                }
            } else {
                output.extend_from_slice(&child.bytes);
            }

            if Some(index) == last_ranked_child {
                emit_pending(&replacements, &mut emitted, rank, &mut output);
            }
        }

        output.extend_from_slice(&self.trailing);
        emit_pending(&replacements, &mut emitted, rank, &mut output);
        output.extend_from_slice(&self.suffix);
        Ok(output)
    }

    /// The root-namespace elements each `mc:Choice`/`mc:Fallback` branch stands
    /// in for. They decide where the branch sits in child order and whether it
    /// hides an element the caller wants to replace.
    fn alternate_content_names(&self, child: &XmlChild) -> Result<Vec<String>, ParseError> {
        if child.local_name != "AlternateContent" {
            return Ok(Vec::new());
        }
        let mut writer = Writer::new(Vec::new());
        let mut root = BytesStart::new("mcRoot");
        for (prefix, namespace) in &self.namespaces {
            let name = match prefix {
                Some(prefix) => format!("xmlns:{prefix}"),
                None => "xmlns".to_owned(),
            };
            root.push_attribute((name.as_str(), namespace.as_str()));
        }
        writer.write_event(Event::Start(root)).map_err(xml_err)?;
        let mut document = writer.into_inner();
        document.extend_from_slice(&child.bytes);
        document.extend_from_slice(b"</mcRoot>");

        let mut reader = NsReader::from_reader(document.as_slice());
        reader.config_mut().expand_empty_elements = false;
        let mut depth = 0_usize;
        let mut names = Vec::new();
        loop {
            let (namespace, event) = reader.read_resolved_event().map_err(xml_err)?;
            let namespace = resolved_namespace(namespace);
            let branch_content = depth == 3 && namespace.as_deref() == self.root_namespace();
            match event {
                Event::Start(element) => {
                    if branch_content {
                        names.push(
                            String::from_utf8_lossy(element.name().local_name().as_ref())
                                .into_owned(),
                        );
                    }
                    depth += 1;
                    if depth > MAX_DEPTH {
                        return Err(ParseError::DepthExceeded);
                    }
                }
                Event::Empty(element) => {
                    if branch_content {
                        names.push(
                            String::from_utf8_lossy(element.name().local_name().as_ref())
                                .into_owned(),
                        );
                    }
                }
                Event::End(_) => depth = depth.saturating_sub(1),
                Event::Eof => break,
                _ => {}
            }
        }
        Ok(names)
    }

    /// Replaces every child with `local_name` by the supplied fragments, in
    /// order, keeping surrounding children untouched.
    pub(crate) fn render_repeated(
        &self,
        local_name: &str,
        replacements: &[Vec<u8>],
        root_attributes: &[(&str, String)],
    ) -> Result<Vec<u8>, ParseError> {
        let mut output = self.prefix_with_attributes(root_attributes)?;
        let last_repeated = self
            .children
            .iter()
            .rposition(|child| self.is_root_child_named(child, local_name));
        let mut replacement_index = 0;
        let mut emitted_without_source = false;

        for (index, child) in self.children.iter().enumerate() {
            output.extend_from_slice(&child.before);
            if last_repeated.is_none() && !emitted_without_source {
                for replacement in replacements {
                    output.extend_from_slice(replacement);
                }
                replacement_index = replacements.len();
                emitted_without_source = true;
            }
            if self.is_root_child_named(child, local_name) {
                if let Some(replacement) = replacements.get(replacement_index) {
                    output.extend_from_slice(replacement);
                }
                replacement_index += 1;
            } else {
                output.extend_from_slice(&child.bytes);
            }
            if Some(index) == last_repeated {
                for replacement in &replacements[replacement_index.min(replacements.len())..] {
                    output.extend_from_slice(replacement);
                }
                replacement_index = replacements.len();
            }
        }

        output.extend_from_slice(&self.trailing);
        for replacement in &replacements[replacement_index.min(replacements.len())..] {
            output.extend_from_slice(replacement);
        }
        output.extend_from_slice(&self.suffix);
        Ok(output)
    }

    /// Puts a generated fragment in the root's namespace. One default-namespace
    /// declaration beats repeating a caller-controlled prefix on every element,
    /// which turns a long prefix into quadratic output.
    pub(crate) fn qualify_fragment(&self, fragment: &[u8]) -> Result<Vec<u8>, ParseError> {
        let Some(prefix) = &self.root_prefix else {
            return Ok(fragment.to_vec());
        };
        if let Some(namespace) = self.root_namespace.as_deref().filter(is_bound_namespace) {
            return bind_default_namespace(fragment, namespace);
        }
        let mut reader = Reader::from_reader(fragment);
        let mut writer = Writer::new(Vec::new());
        loop {
            let event = reader.read_event().map_err(xml_err)?;
            let event = match event {
                Event::Start(mut element) if !element.name().as_ref().contains(&b':') => {
                    let name = format!(
                        "{prefix}:{}",
                        String::from_utf8_lossy(element.name().as_ref())
                    );
                    element.set_name(name.as_bytes());
                    Event::Start(element.into_owned())
                }
                Event::Empty(mut element) if !element.name().as_ref().contains(&b':') => {
                    let name = format!(
                        "{prefix}:{}",
                        String::from_utf8_lossy(element.name().as_ref())
                    );
                    element.set_name(name.as_bytes());
                    Event::Empty(element.into_owned())
                }
                Event::End(element) if !element.name().as_ref().contains(&b':') => {
                    let name = format!(
                        "{prefix}:{}",
                        String::from_utf8_lossy(element.name().as_ref())
                    );
                    Event::End(BytesEnd::new(name))
                }
                Event::Eof => break,
                event => event.into_owned(),
            };
            writer.write_event(event).map_err(xml_err)?;
        }
        Ok(writer.into_inner())
    }

    pub(crate) fn root_namespace(&self) -> Option<&str> {
        self.root_namespace.as_deref()
    }

    /// Finds the prefix the source bound to a namespace ending in `suffix`.
    pub(crate) fn namespace_binding(&self, suffix: &str) -> Option<(&str, &str)> {
        self.namespaces.iter().find_map(|(prefix, namespace)| {
            (namespace.ends_with(suffix))
                .then(|| prefix.as_deref().map(|prefix| (prefix, namespace.as_str())))
                .flatten()
        })
    }

    pub(crate) fn declares_namespace(&self, prefix: &str, namespace: &str) -> bool {
        self.namespaces
            .iter()
            .any(|(candidate, value)| candidate.as_deref() == Some(prefix) && value == namespace)
    }

    pub(crate) fn child(&self, local_name: &str) -> Option<&XmlChild> {
        self.children
            .iter()
            .find(|child| self.is_root_child_named(child, local_name))
    }

    pub(crate) fn children_named<'a>(
        &'a self,
        local_name: &'a str,
    ) -> impl Iterator<Item = &'a XmlChild> {
        self.children
            .iter()
            .filter(move |child| self.is_root_child_named(child, local_name))
    }

    fn prefix_with_attributes(&self, changes: &[(&str, String)]) -> Result<Vec<u8>, ParseError> {
        if changes.is_empty() {
            return Ok(self.prefix.clone());
        }
        let mut reader = Reader::from_reader(self.prefix.as_slice());
        loop {
            let before = reader.buffer_position() as usize;
            match reader.read_event().map_err(xml_err)? {
                Event::Start(element) => {
                    let after = reader.buffer_position() as usize;
                    let mut attributes = attributes(&element)?;
                    for (name, value) in changes {
                        set_attribute(&mut attributes, name, name, value.clone());
                    }
                    let mut root = BytesStart::new(self.root_name.as_str());
                    for attribute in &attributes {
                        root.push_attribute((attribute.name.as_str(), attribute.value.as_str()));
                    }
                    let mut writer = Writer::new(Vec::new());
                    writer.write_event(Event::Start(root)).map_err(xml_err)?;
                    let mut output = self.prefix[..before].to_vec();
                    output.extend_from_slice(&writer.into_inner());
                    output.extend_from_slice(&self.prefix[after..]);
                    return Ok(output);
                }
                Event::Eof => {
                    return Err(ParseError::Malformed(
                        "xml template root start is missing".to_owned(),
                    ));
                }
                _ => {}
            }
        }
    }

    pub(crate) fn is_root_child_named(&self, child: &XmlChild, local_name: &str) -> bool {
        child.local_name == local_name && self.is_root_child(child)
    }

    fn is_root_child(&self, child: &XmlChild) -> bool {
        child.namespace.as_deref() == self.root_namespace()
    }
}

fn emit_pending(
    replacements: &BTreeMap<String, Option<Vec<u8>>>,
    emitted: &mut HashSet<String>,
    rank: fn(&str) -> usize,
    output: &mut Vec<u8>,
) {
    let mut pending = replacements
        .iter()
        .filter(|(name, _)| !emitted.contains(name.as_str()))
        .collect::<Vec<_>>();
    pending.sort_by_key(|(name, _)| rank(name));
    for (name, bytes) in pending {
        emitted.insert(name.clone());
        if let Some(bytes) = bytes {
            output.extend_from_slice(bytes);
        }
    }
}

fn namespace_bindings(
    element: &BytesStart<'_>,
) -> Result<Vec<(Option<String>, String)>, ParseError> {
    element
        .attributes()
        .filter_map(|attribute| match attribute {
            Ok(attribute) => {
                let name = String::from_utf8_lossy(attribute.key.as_ref()).into_owned();
                let prefix = if name == "xmlns" {
                    Some(None)
                } else {
                    name.strip_prefix("xmlns:")
                        .map(|prefix| Some(prefix.to_owned()))
                }?;
                Some(
                    attribute
                        .normalized_value(quick_xml::XmlVersion::Implicit1_0)
                        .map(|value| (prefix, value.into_owned()))
                        .map_err(xml_err),
                )
            }
            Err(error) => Some(Err(xml_err(error))),
        })
        .collect()
}

/// `resolved_namespace` marks prefixes with no in-scope declaration, which
/// happens whenever a captured fragment is re-read on its own.
fn is_bound_namespace(namespace: &&str) -> bool {
    !namespace.starts_with('\0')
}

fn bind_default_namespace(fragment: &[u8], namespace: &str) -> Result<Vec<u8>, ParseError> {
    let mut reader = Reader::from_reader(fragment);
    let mut writer = Writer::new(Vec::new());
    let mut depth = 0_usize;
    loop {
        let event = match reader.read_event().map_err(xml_err)? {
            Event::Start(mut element) => {
                if depth == 0 {
                    element.push_attribute(("xmlns", namespace));
                }
                depth += 1;
                Event::Start(element.into_owned())
            }
            Event::Empty(mut element) => {
                if depth == 0 {
                    element.push_attribute(("xmlns", namespace));
                }
                Event::Empty(element.into_owned())
            }
            Event::End(element) => {
                depth = depth.saturating_sub(1);
                Event::End(element.into_owned())
            }
            Event::Eof => break,
            event => event.into_owned(),
        };
        writer.write_event(event).map_err(xml_err)?;
    }
    Ok(writer.into_inner())
}

fn resolved_namespace(namespace: ResolveResult<'_>) -> Option<String> {
    match namespace {
        ResolveResult::Unbound => None,
        ResolveResult::Bound(namespace) => {
            Some(String::from_utf8_lossy(namespace.as_ref()).into_owned())
        }
        ResolveResult::Unknown(prefix) => Some(format!("\0{}", String::from_utf8_lossy(&prefix))),
    }
}

pub(crate) fn set_attribute(
    attributes: &mut Vec<XmlAttribute>,
    local_name: &str,
    name: &str,
    value: String,
) {
    if let Some(attribute) = attributes
        .iter_mut()
        .find(|attribute| attribute.local_name() == local_name)
    {
        attribute.value = value;
    } else {
        attributes.push(XmlAttribute {
            name: name.to_owned(),
            value,
        });
    }
}

pub(crate) fn remove_attribute(attributes: &mut Vec<XmlAttribute>, local_name: &str) {
    attributes.retain(|attribute| attribute.local_name() != local_name);
}

pub(crate) fn relationship_part_path(part_path: &str) -> String {
    match part_path.rsplit_once('/') {
        Some((directory, filename)) => format!("{directory}/_rels/{filename}.rels"),
        None => format!("_rels/{part_path}.rels"),
    }
}

pub(crate) fn parse_relationships(data: &[u8]) -> Result<Vec<Relationship>, ParseError> {
    let mut reader = reader(data);
    let mut buffer = Vec::new();
    let mut depth = 0;
    let mut relationships = Vec::new();
    loop {
        match next_event(&mut reader, &mut buffer, &mut depth)? {
            Event::Start(element) if local_name(&element) == b"Relationship" => {
                relationships.push(Relationship {
                    attributes: attributes(&element)?,
                });
            }
            Event::Eof => break,
            _ => {}
        }
    }
    Ok(relationships)
}

fn parse_content_types(data: &[u8]) -> Result<Vec<ContentTypeEntry>, ParseError> {
    let mut reader = reader(data);
    let mut buffer = Vec::new();
    let mut depth = 0;
    let mut entries = Vec::new();
    loop {
        match next_event(&mut reader, &mut buffer, &mut depth)? {
            Event::Start(element) => {
                let name = local_name(&element);
                if matches!(name.as_slice(), b"Default" | b"Override") {
                    entries.push(ContentTypeEntry {
                        element: String::from_utf8_lossy(&name).into_owned(),
                        attributes: attributes(&element)?,
                    });
                }
            }
            Event::Eof => break,
            _ => {}
        }
    }
    Ok(entries)
}

struct SheetEntry {
    relationship_id: Option<String>,
    sheet_id: Option<u32>,
    attributes: Vec<XmlAttribute>,
}

fn parse_sheet_entries(data: &[u8]) -> Result<Vec<SheetEntry>, ParseError> {
    let mut reader = reader(data);
    let mut buffer = Vec::new();
    let mut depth = 0;
    let mut sheets_depth = None;
    let mut entries = Vec::new();
    loop {
        match next_event(&mut reader, &mut buffer, &mut depth)? {
            Event::Start(element) if local_name(&element) == b"sheets" => {
                sheets_depth = Some(depth);
            }
            Event::Start(element)
                if local_name(&element) == b"sheet"
                    && sheets_depth.is_some_and(|sheets| depth == sheets + 1) =>
            {
                entries.push(SheetEntry {
                    relationship_id: attr(&element, b"id")?,
                    sheet_id: attr(&element, b"sheetId")?.and_then(|id| id.parse().ok()),
                    attributes: attributes(&element)?,
                });
            }
            Event::End(element)
                if local_name_from_end(&element) == b"sheets"
                    && sheets_depth.is_some_and(|sheets| depth < sheets) =>
            {
                sheets_depth = None;
            }
            Event::Eof => break,
            _ => {}
        }
    }
    Ok(entries)
}

fn part_reference(
    parts: &[(String, Vec<u8>)],
    relationships: &[Relationship],
    relationship_type: &str,
    fallback: &str,
) -> Option<PartReference> {
    if let Some(relationship) = relationships
        .iter()
        .find(|relationship| relationship.has_type(relationship_type))
        && let Some(target) = relationship.target()
    {
        let path = resolve_part_path("xl", target);
        if find_part(parts, &path).is_some() {
            return Some(PartReference {
                path,
                relationship_id: relationship.id().map(str::to_owned),
            });
        }
    }
    find_part(parts, fallback).map(|_| PartReference {
        path: fallback.to_owned(),
        relationship_id: None,
    })
}

fn attributes(element: &BytesStart<'_>) -> Result<Vec<XmlAttribute>, ParseError> {
    element
        .attributes()
        .map(|attribute| {
            let attribute = attribute.map_err(xml_err)?;
            Ok(XmlAttribute {
                name: String::from_utf8_lossy(attribute.key.as_ref()).into_owned(),
                value: attribute
                    .normalized_value(quick_xml::XmlVersion::Implicit1_0)
                    .map_err(xml_err)?
                    .into_owned(),
            })
        })
        .collect()
}

pub(crate) fn attributes_from_fragment(fragment: &[u8]) -> Result<Vec<XmlAttribute>, ParseError> {
    let mut reader = Reader::from_reader(fragment);
    loop {
        match reader.read_event().map_err(xml_err)? {
            Event::Start(element) | Event::Empty(element) => return attributes(&element),
            Event::Eof => {
                return Err(ParseError::Malformed(
                    "xml fragment has no element".to_owned(),
                ));
            }
            _ => {}
        }
    }
}

fn local_name_from_end(element: &quick_xml::events::BytesEnd<'_>) -> Vec<u8> {
    element.name().local_name().as_ref().to_vec()
}
