//! `xlsx_model::Workbook` -> minimal valid xlsx parts. structural round-trip:
//! whatever `read` captures comes back out.

use std::collections::{BTreeMap, HashMap, HashSet, VecDeque};
use std::io;

use quick_xml::Writer;
use quick_xml::events::{BytesDecl, BytesEnd, BytesStart, BytesText, Event};
use xlsx_model::addr::RowId;
use xlsx_model::styles::{Alignment, Border, BorderEdge, Color, Fill, Font, Stylesheet, Xf};
use xlsx_model::{Cell, CellRef, CellValue, DateSystem, Sheet, Workbook};

use crate::ParseError;
use crate::package::{
    ContentTypeEntry, PartReference, PreservedPackage, Relationship, XmlAttribute, XmlTemplate,
    attributes_from_fragment, relationship_part_path, remove_attribute, set_attribute,
};
use crate::xml::xml_err;

const NS_MAIN: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const NS_STRICT_MAIN: &str = "http://purl.oclc.org/ooxml/spreadsheetml/main";
const NS_R: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const NS_CT: &str = "http://schemas.openxmlformats.org/package/2006/content-types";
const NS_PKG_REL: &str = "http://schemas.openxmlformats.org/package/2006/relationships";
const CT_WORKSHEET: &str =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml";
const CT_STRICT_WORKSHEET: &str = "application/vnd.ms-excel.worksheet+xml";
const CT_WORKBOOK: &str =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml";
const CT_SST: &str =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml";
const CT_STRICT_SST: &str = "application/vnd.ms-excel.sharedStrings+xml";
const REL_WORKSHEET: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet";
const REL_SST: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings";
const CT_STYLES: &str = "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml";
const CT_STRICT_STYLES: &str = "application/vnd.ms-excel.styles+xml";
const CT_THEME: &str = "application/vnd.openxmlformats-officedocument.theme+xml";
const REL_STYLES: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles";
const REL_THEME: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme";
const REL_OFFICE_DOCUMENT: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
const REL_HYPERLINK: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";
const NS_DML: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";

/// serialize a workbook to opc parts in a fixed, deterministic order.
pub fn serialize_workbook(wb: &Workbook) -> Result<Vec<(String, Vec<u8>)>, ParseError> {
    let have_sst = !wb.shared_strings.is_empty();
    let have_styles = !wb.styles.is_empty();
    let mut parts = vec![
        (
            "[Content_Types].xml".to_string(),
            content_types(wb, have_sst, have_styles)?,
        ),
        ("_rels/.rels".to_string(), root_rels()?),
        ("xl/workbook.xml".to_string(), workbook_xml(wb)?),
        (
            "xl/_rels/workbook.xml.rels".to_string(),
            workbook_rels(wb, have_sst, have_styles)?,
        ),
    ];
    if have_sst {
        parts.push(("xl/sharedStrings.xml".to_string(), shared_strings_xml(wb)?));
    }
    if have_styles {
        parts.push(("xl/styles.xml".to_string(), styles_xml(&wb.styles)?));
        parts.push(("xl/theme/theme1.xml".to_string(), theme_xml(&wb.styles)?));
    }
    for (i, sheet) in wb.sheets.iter().enumerate() {
        parts.push((
            format!("xl/worksheets/sheet{}.xml", i + 1),
            worksheet_xml(sheet, wb)?,
        ));
        if sheet
            .hyperlinks
            .iter()
            .any(|link| link.external_target.is_some())
        {
            parts.push((
                format!("xl/worksheets/_rels/sheet{}.xml.rels", i + 1),
                worksheet_rels_xml(sheet)?,
            ));
        }
    }
    Ok(parts)
}

/// Serializes owned parts over a preserved source package.
pub fn serialize_workbook_with_package(
    wb: &Workbook,
    package: &PreservedPackage,
) -> Result<Vec<(String, Vec<u8>)>, ParseError> {
    let origins = (0..wb.sheets.len())
        .map(|index| (index < package.sheets.len()).then_some(index))
        .collect::<Vec<_>>();
    serialize_workbook_with_package_and_origins(wb, package, &origins)
}

#[doc(hidden)]
pub fn serialize_workbook_with_package_and_origins(
    wb: &Workbook,
    package: &PreservedPackage,
    origins: &[Option<usize>],
) -> Result<Vec<(String, Vec<u8>)>, ParseError> {
    let edited = wb != &package.original_workbook
        || origins.len() != package.sheets.len()
        || origins
            .iter()
            .enumerate()
            .any(|(index, origin)| *origin != Some(index));
    serialize_workbook_with_package_and_origins_after_edits(wb, package, origins, edited)
}

/// `edited` false means the caller guarantees an untouched model, which is
/// returned as the source bytes; true drops the now-stale calculation chain.
#[doc(hidden)]
pub fn serialize_workbook_with_package_and_origins_after_edits(
    wb: &Workbook,
    package: &PreservedPackage,
    origins: &[Option<usize>],
    edited: bool,
) -> Result<Vec<(String, Vec<u8>)>, ParseError> {
    if origins.len() != wb.sheets.len() {
        return Err(ParseError::Malformed(
            "sheet origin count does not match workbook".to_owned(),
        ));
    }
    if !edited
        && origins.len() == package.sheets.len()
        && origins
            .iter()
            .enumerate()
            .all(|(index, origin)| *origin == Some(index))
    {
        return Ok(package.parts.clone());
    }

    let have_sst = !wb.shared_strings.is_empty();
    let have_styles = !wb.styles.is_empty();
    let main_namespace = package
        .workbook_template
        .root_namespace()
        .unwrap_or(NS_MAIN);
    let worksheet_relationship_type = relationship_type(package, "worksheet", REL_WORKSHEET);
    let shared_strings_relationship_type = relationship_type(package, "sharedStrings", REL_SST);
    let styles_relationship_type = relationship_type(package, "styles", REL_STYLES);
    let theme_relationship_type = relationship_type(package, "theme", REL_THEME);
    let mut used_relationship_ids = package
        .workbook_relationships
        .iter()
        .filter_map(Relationship::id)
        .map(str::to_owned)
        .collect::<HashSet<_>>();
    let mut used_paths = package
        .parts
        .iter()
        .map(|(path, _)| normalized_part_name(path))
        .collect::<HashSet<_>>();
    let sheets = plan_sheets(
        package,
        origins,
        &mut used_relationship_ids,
        &mut used_paths,
        &worksheet_relationship_type,
    )?;
    let shared_strings = have_sst.then(|| {
        plan_part(
            package.shared_strings.as_ref(),
            &package.workbook_relationships,
            "sharedStrings",
            &shared_strings_relationship_type,
            "xl/sharedStrings.xml",
            &mut used_relationship_ids,
        )
    });
    let styles = have_styles.then(|| {
        plan_part(
            package.styles.as_ref(),
            &package.workbook_relationships,
            "styles",
            &styles_relationship_type,
            "xl/styles.xml",
            &mut used_relationship_ids,
        )
    });
    let theme = have_styles.then(|| {
        plan_part(
            package.theme.as_ref(),
            &package.workbook_relationships,
            "theme",
            &theme_relationship_type,
            "xl/theme/theme1.xml",
            &mut used_relationship_ids,
        )
    });

    let mut parts = PartStore::new(package.parts.clone());
    let retained_origins = sheets
        .iter()
        .filter_map(|sheet| sheet.origin)
        .collect::<HashSet<_>>();
    for (index, source) in package.sheets.iter().enumerate() {
        if !retained_origins.contains(&index) {
            parts.remove(&source.path);
            parts.remove(&relationship_part_path(&source.path));
        }
    }
    if edited {
        for calc_chain in &package.calc_chains {
            parts.remove(&calc_chain.path);
            parts.remove(&relationship_part_path(&calc_chain.path));
        }
    }

    for (sheet, plan) in wb.sheets.iter().zip(&sheets) {
        let bytes = match plan.origin.and_then(|origin| package.sheets.get(origin)) {
            Some(source) if source.is_worksheet() => {
                worksheet_xml_with_template(sheet, wb, &source.template)?
            }
            Some(_) => continue,
            None => worksheet_xml_with_namespace(sheet, wb, main_namespace)?,
        };
        parts.set(plan.path.clone(), bytes);
    }

    let workbook = workbook_xml_with_template(wb, package, &sheets, edited)?;
    parts.set("xl/workbook.xml".to_owned(), workbook);
    parts.set(
        "xl/_rels/workbook.xml.rels".to_owned(),
        merged_workbook_relationships(
            package,
            &sheets,
            shared_strings.as_ref(),
            styles.as_ref(),
            theme.as_ref(),
            edited,
        )?,
    );
    parts.set(
        "_rels/.rels".to_owned(),
        merged_root_relationships(package)?,
    );
    parts.set(
        "[Content_Types].xml".to_owned(),
        merged_content_types(
            package,
            &sheets,
            shared_strings.as_ref(),
            styles.as_ref(),
            theme.as_ref(),
            edited,
        )?,
    );

    replace_shared_strings(
        &mut parts,
        wb,
        package,
        shared_strings.as_ref(),
        main_namespace,
    )?;
    replace_optional_part(
        &mut parts,
        package.styles.as_ref(),
        styles.as_ref(),
        || match &package.stylesheet_template {
            Some(template) => styles_xml_with_template(&wb.styles, template),
            None => styles_xml_with_namespace(&wb.styles, main_namespace),
        },
    )?;

    match (package.theme.as_ref(), theme.as_ref()) {
        (Some(source), Some(planned))
            if source.path == planned.path
                && wb.styles.theme == package.original_workbook.styles.theme => {}
        (source, Some(planned)) => {
            if let Some(source) = source
                && source.path != planned.path
            {
                parts.remove(&source.path);
            }
            parts.set(planned.path.clone(), theme_xml(&wb.styles)?);
        }
        (Some(source), None) => parts.remove(&source.path),
        (None, None) => {}
    }

    Ok(parts.finish())
}

#[derive(Clone)]
struct PlannedSheet {
    origin: Option<usize>,
    path: String,
    relationship: Relationship,
    sheet_id: u32,
    attributes: Vec<XmlAttribute>,
}

#[derive(Clone)]
struct PlannedPart {
    path: String,
    relationship: Relationship,
}

/// Hands out `sheetId` values that collide with no source sheet, continuing
/// past the highest source id before reusing gaps.
struct SheetIds {
    used: HashSet<u32>,
    next: u64,
}

impl SheetIds {
    fn new(package: &PreservedPackage) -> Self {
        let used = package
            .sheets
            .iter()
            .map(|sheet| sheet.sheet_id)
            .collect::<HashSet<_>>();
        let next = u64::from(used.iter().copied().max().unwrap_or(0)) + 1;
        Self { used, next }
    }

    fn allocate(&mut self) -> Result<u32, ParseError> {
        while self.next <= u64::from(u32::MAX) {
            let candidate = self.next as u32;
            self.next += 1;
            if self.used.insert(candidate) {
                return Ok(candidate);
            }
        }
        (1..=u32::MAX)
            .find(|candidate| self.used.insert(*candidate))
            .ok_or_else(|| ParseError::Malformed("worksheet ids exhausted".to_owned()))
    }
}

fn plan_sheets(
    package: &PreservedPackage,
    origins: &[Option<usize>],
    used_relationship_ids: &mut HashSet<String>,
    used_paths: &mut HashSet<String>,
    worksheet_relationship_type: &str,
) -> Result<Vec<PlannedSheet>, ParseError> {
    let mut claimed_origins = HashSet::new();
    let mut sheet_ids = SheetIds::new(package);
    origins
        .iter()
        .map(|origin| {
            let origin = origin
                .filter(|origin| *origin < package.sheets.len() && claimed_origins.insert(*origin));
            if let Some(origin) = origin {
                let source = &package.sheets[origin];
                let source_relationship = source.relationship_id.as_deref().and_then(|id| {
                    package
                        .workbook_relationships
                        .iter()
                        .find(|relationship| relationship.id() == Some(id))
                });
                let relationship = match source_relationship {
                    Some(relationship) => relationship.clone(),
                    None => new_relationship(
                        next_relationship_id(used_relationship_ids),
                        source
                            .relationship_type
                            .as_deref()
                            .unwrap_or(worksheet_relationship_type),
                        relative_to_xl(&source.path),
                    ),
                };
                return Ok(PlannedSheet {
                    origin: Some(origin),
                    path: source.path.clone(),
                    relationship,
                    sheet_id: source.sheet_id,
                    attributes: source.attributes.clone(),
                });
            }

            let path = next_sheet_path(used_paths);
            let relationship = new_relationship(
                next_relationship_id(used_relationship_ids),
                worksheet_relationship_type,
                relative_to_xl(&path),
            );
            Ok(PlannedSheet {
                origin: None,
                path,
                relationship,
                sheet_id: sheet_ids.allocate()?,
                attributes: Vec::new(),
            })
        })
        .collect()
}

fn plan_part(
    source: Option<&PartReference>,
    relationships: &[Relationship],
    type_suffix: &str,
    relationship_type: &str,
    fallback_path: &str,
    used_relationship_ids: &mut HashSet<String>,
) -> PlannedPart {
    if let Some(source) = source {
        let source_relationship = source
            .relationship_id
            .as_deref()
            .and_then(|id| {
                relationships
                    .iter()
                    .find(|relationship| relationship.id() == Some(id))
            })
            .or_else(|| {
                relationships
                    .iter()
                    .find(|relationship| relationship.has_type(type_suffix))
            });
        let relationship = match source_relationship {
            Some(relationship) => relationship.clone(),
            None => new_relationship(
                next_relationship_id(used_relationship_ids),
                relationship_type,
                relative_to_xl(&source.path),
            ),
        };
        return PlannedPart {
            path: source.path.clone(),
            relationship,
        };
    }

    PlannedPart {
        path: fallback_path.to_owned(),
        relationship: new_relationship(
            next_relationship_id(used_relationship_ids),
            relationship_type,
            relative_to_xl(fallback_path),
        ),
    }
}

/// Picks the relationship type the source uses, keeping Strict URIs intact.
fn relationship_type(package: &PreservedPackage, suffix: &str, fallback: &str) -> String {
    relationship_type_from(&package.workbook_relationships, suffix, fallback)
}

fn relationship_type_from(relationships: &[Relationship], suffix: &str, fallback: &str) -> String {
    if let Some(relationship_type) = relationships
        .iter()
        .find(|relationship| relationship.has_type(suffix))
        .and_then(|relationship| relationship.attribute("Type"))
    {
        return relationship_type.to_owned();
    }
    relationships
        .iter()
        .filter_map(|relationship| relationship.attribute("Type"))
        .find_map(|relationship_type| {
            let (base, _) = relationship_type.rsplit_once('/')?;
            base.ends_with("/officeDocument/relationships")
                .then(|| format!("{base}/{suffix}"))
        })
        .unwrap_or_else(|| fallback.to_owned())
}

fn workbook_relationship_namespace(package: &PreservedPackage) -> (String, String) {
    if let Some((prefix, namespace)) = package
        .workbook_template
        .namespace_binding("/officeDocument/relationships")
    {
        return (prefix.to_owned(), namespace.to_owned());
    }
    let namespace = package
        .workbook_relationships
        .iter()
        .filter_map(|relationship| relationship.attribute("Type"))
        .find_map(|relationship_type| {
            let (namespace, _) = relationship_type.rsplit_once('/')?;
            namespace
                .ends_with("/officeDocument/relationships")
                .then(|| namespace.to_owned())
        })
        .unwrap_or_else(|| NS_R.to_owned());
    ("r".to_owned(), namespace)
}

fn new_relationship(id: String, relationship_type: &str, target: String) -> Relationship {
    Relationship {
        attributes: vec![
            XmlAttribute {
                name: "Id".to_owned(),
                value: id,
            },
            XmlAttribute {
                name: "Type".to_owned(),
                value: relationship_type.to_owned(),
            },
            XmlAttribute {
                name: "Target".to_owned(),
                value: target,
            },
        ],
    }
}

fn next_relationship_id(used: &mut HashSet<String>) -> String {
    let mut index = 1_u64;
    loop {
        let candidate = format!("rId{index}");
        if used.insert(candidate.clone()) {
            return candidate;
        }
        index += 1;
    }
}

fn next_sheet_path(used: &mut HashSet<String>) -> String {
    let mut index = 1_u64;
    loop {
        let path = format!("xl/worksheets/sheet{index}.xml");
        if used.insert(normalized_part_name(&path)) {
            return path;
        }
        index += 1;
    }
}

fn relative_to_xl(path: &str) -> String {
    path.trim_start_matches('/')
        .strip_prefix("xl/")
        .unwrap_or(path.trim_start_matches('/'))
        .to_owned()
}

fn normalized_part_name(path: &str) -> String {
    path.trim_start_matches('/').to_ascii_lowercase()
}

fn replace_optional_part(
    parts: &mut PartStore,
    source: Option<&PartReference>,
    planned: Option<&PlannedPart>,
    serialize: impl FnOnce() -> Result<Vec<u8>, ParseError>,
) -> Result<(), ParseError> {
    match (source, planned) {
        (source, Some(planned)) => {
            if let Some(source) = source
                && source.path != planned.path
            {
                parts.remove(&source.path);
            }
            parts.set(planned.path.clone(), serialize()?);
        }
        (Some(source), None) => parts.remove(&source.path),
        (None, None) => {}
    }
    Ok(())
}

/// Regenerates only the shared string indices the model changed, so rich runs,
/// phonetic properties and sst extensions survive untouched.
fn replace_shared_strings(
    parts: &mut PartStore,
    wb: &Workbook,
    package: &PreservedPackage,
    planned: Option<&PlannedPart>,
    main_namespace: &str,
) -> Result<(), ParseError> {
    match (package.shared_strings.as_ref(), planned) {
        (Some(source), Some(planned))
            if source.path == planned.path
                && wb.shared_strings == package.original_workbook.shared_strings => {}
        (source, Some(planned)) => {
            if let Some(source) = source
                && source.path != planned.path
            {
                parts.remove(&source.path);
            }
            let bytes = match &package.shared_strings_template {
                Some(template) => shared_strings_xml_with_template(wb, package, template)?,
                None => shared_strings_xml_with_namespace(wb, main_namespace)?,
            };
            parts.set(planned.path.clone(), bytes);
        }
        (Some(source), None) => parts.remove(&source.path),
        (None, None) => {}
    }
    Ok(())
}

struct PartStore {
    parts: Vec<Option<(String, Vec<u8>)>>,
    positions: HashMap<String, usize>,
}

impl PartStore {
    fn new(parts: Vec<(String, Vec<u8>)>) -> Self {
        let positions = parts
            .iter()
            .enumerate()
            .map(|(index, (path, _))| (normalized_part_name(path), index))
            .collect();
        Self {
            parts: parts.into_iter().map(Some).collect(),
            positions,
        }
    }

    fn remove(&mut self, path: &str) {
        if let Some(index) = self.positions.remove(&normalized_part_name(path)) {
            self.parts[index] = None;
        }
    }

    fn set(&mut self, path: String, bytes: Vec<u8>) {
        let normalized = normalized_part_name(&path);
        if let Some(index) = self.positions.get(&normalized).copied() {
            let authored_path = self.parts[index]
                .as_ref()
                .map(|(path, _)| path.clone())
                .unwrap_or(path);
            self.parts[index] = Some((authored_path, bytes));
        } else {
            self.positions.insert(normalized, self.parts.len());
            self.parts.push(Some((path, bytes)));
        }
    }

    fn finish(self) -> Vec<(String, Vec<u8>)> {
        self.parts.into_iter().flatten().collect()
    }
}

/// run a builder against a fresh writer that already emitted the xml decl.
fn doc<F>(f: F) -> Result<Vec<u8>, ParseError>
where
    F: FnOnce(&mut Writer<Vec<u8>>) -> io::Result<()>,
{
    let mut w = Writer::new(Vec::new());
    w.write_event(Event::Decl(BytesDecl::new(
        "1.0",
        Some("UTF-8"),
        Some("yes"),
    )))
    .map_err(xml_err)?;
    f(&mut w).map_err(xml_err)?;
    Ok(w.into_inner())
}

fn fragment<F>(f: F) -> Result<Vec<u8>, ParseError>
where
    F: FnOnce(&mut Writer<Vec<u8>>) -> io::Result<()>,
{
    let mut writer = Writer::new(Vec::new());
    f(&mut writer).map_err(xml_err)?;
    Ok(writer.into_inner())
}

fn merged_root_relationships(package: &PreservedPackage) -> Result<Vec<u8>, ParseError> {
    let mut relationships = package.root_relationships.clone();
    if !relationships
        .iter()
        .any(|relationship| relationship.has_type("officeDocument"))
    {
        let mut used = relationships
            .iter()
            .filter_map(Relationship::id)
            .map(str::to_owned)
            .collect::<HashSet<_>>();
        relationships.push(new_relationship(
            next_relationship_id(&mut used),
            &relationship_type_from(
                &package.root_relationships,
                "officeDocument",
                REL_OFFICE_DOCUMENT,
            ),
            "xl/workbook.xml".to_owned(),
        ));
    }
    relationships_xml(&relationships)
}

fn merged_workbook_relationships(
    package: &PreservedPackage,
    sheets: &[PlannedSheet],
    shared_strings: Option<&PlannedPart>,
    styles: Option<&PlannedPart>,
    theme: Option<&PlannedPart>,
    edited: bool,
) -> Result<Vec<u8>, ParseError> {
    let planned = sheets
        .iter()
        .map(|sheet| &sheet.relationship)
        .chain(shared_strings.map(|part| &part.relationship))
        .chain(styles.map(|part| &part.relationship))
        .chain(theme.map(|part| &part.relationship))
        .collect::<Vec<_>>();
    let planned_by_id = planned
        .iter()
        .filter_map(|relationship| relationship.id().map(|id| (id.to_owned(), *relationship)))
        .collect::<HashMap<_, _>>();
    let mut source_owned_ids = package
        .sheets
        .iter()
        .filter_map(|sheet| sheet.relationship_id.clone())
        .collect::<HashSet<_>>();
    source_owned_ids.extend(
        [
            package.shared_strings.as_ref(),
            package.styles.as_ref(),
            package.theme.as_ref(),
        ]
        .into_iter()
        .flatten()
        .filter_map(|part| part.relationship_id.clone()),
    );

    let mut emitted = HashSet::new();
    let mut merged = Vec::new();
    for source in &package.workbook_relationships {
        if edited && source.has_type("calcChain") {
            continue;
        }
        let id = source.id();
        if let Some(planned) = id.and_then(|id| planned_by_id.get(id))
            && emitted.insert(id.unwrap().to_owned())
        {
            merged.push((*planned).clone());
        } else if id.is_some_and(|id| source_owned_ids.contains(id))
            || source.has_type("sharedStrings")
            || source.has_type("styles")
            || source.has_type("theme")
        {
        } else {
            merged.push(source.clone());
            if let Some(id) = id {
                emitted.insert(id.to_owned());
            }
        }
    }
    for relationship in planned {
        if let Some(id) = relationship.id()
            && emitted.insert(id.to_owned())
        {
            merged.push(relationship.clone());
        }
    }
    relationships_xml(&merged)
}

fn relationships_xml(relationships: &[Relationship]) -> Result<Vec<u8>, ParseError> {
    doc(|writer| {
        writer
            .create_element("Relationships")
            .with_attribute(("xmlns", NS_PKG_REL))
            .write_inner_content(|writer| {
                for relationship in relationships {
                    write_empty_element(writer, "Relationship", &relationship.attributes)?;
                }
                Ok(())
            })?;
        Ok(())
    })
}

fn merged_content_types(
    package: &PreservedPackage,
    sheets: &[PlannedSheet],
    shared_strings: Option<&PlannedPart>,
    styles: Option<&PlannedPart>,
    theme: Option<&PlannedPart>,
    edited: bool,
) -> Result<Vec<u8>, ParseError> {
    let mut desired = BTreeMap::new();
    let strict = package.workbook_template.root_namespace() == Some(NS_STRICT_MAIN);
    desired.insert(
        normalized_part_name("xl/workbook.xml"),
        source_content_type(package, "xl/workbook.xml").unwrap_or(CT_WORKBOOK),
    );
    let fallback_worksheet_content_type = package
        .sheets
        .iter()
        .find(|sheet| sheet.is_worksheet())
        .and_then(|sheet| source_content_type(package, &sheet.path))
        .unwrap_or(if strict {
            CT_STRICT_WORKSHEET
        } else {
            CT_WORKSHEET
        });
    for sheet in sheets {
        let content_type = sheet
            .origin
            .and_then(|origin| package.sheets.get(origin))
            .and_then(|source| source_content_type(package, &source.path))
            .unwrap_or(fallback_worksheet_content_type);
        desired.insert(normalized_part_name(&sheet.path), content_type);
    }
    if let Some(part) = shared_strings {
        desired.insert(
            normalized_part_name(&part.path),
            package
                .shared_strings
                .as_ref()
                .and_then(|source| source_content_type(package, &source.path))
                .unwrap_or(if strict { CT_STRICT_SST } else { CT_SST }),
        );
    }
    if let Some(part) = styles {
        desired.insert(
            normalized_part_name(&part.path),
            package
                .styles
                .as_ref()
                .and_then(|source| source_content_type(package, &source.path))
                .unwrap_or(if strict { CT_STRICT_STYLES } else { CT_STYLES }),
        );
    }
    if let Some(part) = theme {
        desired.insert(
            normalized_part_name(&part.path),
            package
                .theme
                .as_ref()
                .and_then(|source| source_content_type(package, &source.path))
                .unwrap_or(CT_THEME),
        );
    }

    let mut source_owned = HashSet::from([normalized_part_name("xl/workbook.xml")]);
    source_owned.extend(
        package
            .sheets
            .iter()
            .map(|sheet| normalized_part_name(&sheet.path)),
    );
    source_owned.extend(
        [
            package.shared_strings.as_ref(),
            package.styles.as_ref(),
            package.theme.as_ref(),
        ]
        .into_iter()
        .flatten()
        .map(|part| normalized_part_name(&part.path)),
    );

    let mut entries = Vec::new();
    let mut emitted_parts = HashSet::new();
    let mut default_extensions = HashMap::new();
    let calc_chain_paths = package
        .calc_chains
        .iter()
        .map(|part| normalized_part_name(&part.path))
        .collect::<HashSet<_>>();
    for entry in &package.content_types {
        if entry.element == "Default" {
            if let Some(extension) = entry.attribute("Extension") {
                default_extensions.insert(
                    extension.to_ascii_lowercase(),
                    entry
                        .attribute("ContentType")
                        .unwrap_or_default()
                        .to_owned(),
                );
            }
            entries.push(entry.clone());
            continue;
        }
        let Some(part_name) = entry.attribute("PartName") else {
            entries.push(entry.clone());
            continue;
        };
        let normalized = normalized_part_name(part_name);
        if edited && calc_chain_paths.contains(&normalized) {
            continue;
        }
        if source_owned.contains(&normalized) {
            if desired.contains_key(&normalized) && emitted_parts.insert(normalized) {
                entries.push(entry.clone());
            }
        } else {
            entries.push(entry.clone());
        }
    }

    for (path, content_type) in desired {
        let covered_by_default = path
            .rsplit_once('.')
            .and_then(|(_, extension)| default_extensions.get(extension))
            .is_some_and(|default| default == content_type);
        if covered_by_default {
            emitted_parts.insert(path);
            continue;
        }
        if emitted_parts.insert(path.clone()) {
            entries.push(ContentTypeEntry {
                element: "Override".to_owned(),
                attributes: vec![
                    XmlAttribute {
                        name: "PartName".to_owned(),
                        value: format!("/{path}"),
                    },
                    XmlAttribute {
                        name: "ContentType".to_owned(),
                        value: content_type.to_owned(),
                    },
                ],
            });
        }
    }
    if default_extensions
        .insert(
            "rels".to_owned(),
            "application/vnd.openxmlformats-package.relationships+xml".to_owned(),
        )
        .is_none()
    {
        entries.insert(
            0,
            ContentTypeEntry {
                element: "Default".to_owned(),
                attributes: vec![
                    XmlAttribute {
                        name: "Extension".to_owned(),
                        value: "rels".to_owned(),
                    },
                    XmlAttribute {
                        name: "ContentType".to_owned(),
                        value: "application/vnd.openxmlformats-package.relationships+xml"
                            .to_owned(),
                    },
                ],
            },
        );
    }
    if default_extensions
        .insert("xml".to_owned(), "application/xml".to_owned())
        .is_none()
    {
        entries.insert(
            default_extensions.contains_key("rels") as usize,
            ContentTypeEntry {
                element: "Default".to_owned(),
                attributes: vec![
                    XmlAttribute {
                        name: "Extension".to_owned(),
                        value: "xml".to_owned(),
                    },
                    XmlAttribute {
                        name: "ContentType".to_owned(),
                        value: "application/xml".to_owned(),
                    },
                ],
            },
        );
    }

    doc(|writer| {
        writer
            .create_element("Types")
            .with_attribute(("xmlns", NS_CT))
            .write_inner_content(|writer| {
                for entry in &entries {
                    write_empty_element(writer, &entry.element, &entry.attributes)?;
                }
                Ok(())
            })?;
        Ok(())
    })
}

/// The type OPC actually resolves for a part: its exact `Override`, else the
/// `Default` for its extension. Chartsheets and macro-enabled workbooks are
/// commonly typed by extension alone.
fn source_content_type<'a>(package: &'a PreservedPackage, path: &str) -> Option<&'a str> {
    let normalized = normalized_part_name(path);
    package
        .content_types
        .iter()
        .find_map(|entry| {
            (entry.element == "Override"
                && entry
                    .attribute("PartName")
                    .is_some_and(|part| normalized_part_name(part) == normalized))
            .then(|| entry.attribute("ContentType"))
            .flatten()
        })
        .or_else(|| default_content_type(package, path))
}

fn default_content_type<'a>(package: &'a PreservedPackage, path: &str) -> Option<&'a str> {
    let extension = path.rsplit_once('.')?.1;
    package.content_types.iter().find_map(|entry| {
        (entry.element == "Default"
            && entry
                .attribute("Extension")
                .is_some_and(|value| value.eq_ignore_ascii_case(extension)))
        .then(|| entry.attribute("ContentType"))
        .flatten()
    })
}

fn write_empty_element(
    writer: &mut Writer<Vec<u8>>,
    name: &str,
    attributes: &[XmlAttribute],
) -> io::Result<()> {
    let mut element = BytesStart::new(name);
    for attribute in attributes {
        element.push_attribute((attribute.name.as_str(), attribute.value.as_str()));
    }
    writer.write_event(Event::Empty(element))?;
    Ok(())
}

fn content_types(wb: &Workbook, have_sst: bool, have_styles: bool) -> Result<Vec<u8>, ParseError> {
    doc(|w| {
        w.create_element("Types")
            .with_attribute(("xmlns", NS_CT))
            .write_inner_content(|w| {
                w.create_element("Default")
                    .with_attribute(("Extension", "rels"))
                    .with_attribute((
                        "ContentType",
                        "application/vnd.openxmlformats-package.relationships+xml",
                    ))
                    .write_empty()?;
                w.create_element("Default")
                    .with_attribute(("Extension", "xml"))
                    .with_attribute(("ContentType", "application/xml"))
                    .write_empty()?;
                w.create_element("Override")
                    .with_attribute(("PartName", "/xl/workbook.xml"))
                    .with_attribute((
                        "ContentType",
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
                    ))
                    .write_empty()?;
                if have_sst {
                    w.create_element("Override")
                        .with_attribute(("PartName", "/xl/sharedStrings.xml"))
                        .with_attribute(("ContentType", CT_SST))
                        .write_empty()?;
                }
                if have_styles {
                    w.create_element("Override")
                        .with_attribute(("PartName", "/xl/styles.xml"))
                        .with_attribute(("ContentType", CT_STYLES))
                        .write_empty()?;
                    w.create_element("Override")
                        .with_attribute(("PartName", "/xl/theme/theme1.xml"))
                        .with_attribute(("ContentType", CT_THEME))
                        .write_empty()?;
                }
                for i in 0..wb.sheets.len() {
                    let part = format!("/xl/worksheets/sheet{}.xml", i + 1);
                    w.create_element("Override")
                        .with_attribute(("PartName", part.as_str()))
                        .with_attribute(("ContentType", CT_WORKSHEET))
                        .write_empty()?;
                }
                Ok(())
            })?;
        Ok(())
    })
}

fn root_rels() -> Result<Vec<u8>, ParseError> {
    doc(|w| {
        w.create_element("Relationships")
            .with_attribute(("xmlns", NS_PKG_REL))
            .write_inner_content(|w| {
                w.create_element("Relationship")
                    .with_attribute(("Id", "rId1"))
                    .with_attribute((
                        "Type",
                        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument",
                    ))
                    .with_attribute(("Target", "xl/workbook.xml"))
                    .write_empty()?;
                Ok(())
            })?;
        Ok(())
    })
}

fn workbook_xml(wb: &Workbook) -> Result<Vec<u8>, ParseError> {
    doc(|w| {
        w.create_element("workbook")
            .with_attribute(("xmlns", NS_MAIN))
            .with_attribute(("xmlns:r", NS_R))
            .write_inner_content(|w| {
                if wb.date_system == DateSystem::V1904 {
                    w.create_element("workbookPr")
                        .with_attribute(("date1904", "1"))
                        .write_empty()?;
                }
                w.create_element("sheets").write_inner_content(|w| {
                    for (i, sheet) in wb.sheets.iter().enumerate() {
                        let rid = format!("rId{}", i + 1);
                        let sid = (i + 1).to_string();
                        w.create_element("sheet")
                            .with_attribute(("name", sheet.name.as_str()))
                            .with_attribute(("sheetId", sid.as_str()))
                            .with_attribute(("r:id", rid.as_str()))
                            .write_empty()?;
                    }
                    Ok(())
                })?;
                write_defined_names(w, wb)?;
                Ok(())
            })?;
        Ok(())
    })
}

fn write_defined_names(w: &mut Writer<Vec<u8>>, wb: &Workbook) -> io::Result<()> {
    if wb.defined_names.is_empty() {
        return Ok(());
    }
    w.create_element("definedNames").write_inner_content(|w| {
        for defined in &wb.defined_names {
            let mut element = BytesStart::new("definedName");
            element.push_attribute(("name", defined.name.as_str()));
            let local_sheet = defined.local_sheet.map(|sheet| sheet.0.to_string());
            if let Some(local_sheet) = &local_sheet {
                element.push_attribute(("localSheetId", local_sheet.as_str()));
            }
            if defined.hidden {
                element.push_attribute(("hidden", "1"));
            }
            w.write_event(Event::Start(element))?;
            w.write_event(Event::Text(BytesText::new(&defined.formula)))?;
            w.write_event(Event::End(BytesEnd::new("definedName")))?;
        }
        Ok(())
    })?;
    Ok(())
}

fn workbook_xml_with_template(
    wb: &Workbook,
    package: &PreservedPackage,
    sheets: &[PlannedSheet],
    edited: bool,
) -> Result<Vec<u8>, ParseError> {
    let workbook_pr =
        if package.workbook_pr_attributes.is_some() || wb.date_system == DateSystem::V1904 {
            let mut attributes = package.workbook_pr_attributes.clone().unwrap_or_default();
            if wb.date_system == DateSystem::V1904 {
                set_attribute(&mut attributes, "date1904", "date1904", "1".to_owned());
            } else {
                remove_attribute(&mut attributes, "date1904");
            }
            Some(fragment(|writer| {
                write_empty_element(writer, "workbookPr", &attributes)
            })?)
        } else {
            None
        };
    let (relationship_prefix, relationship_namespace) = workbook_relationship_namespace(package);
    let mut sheets_attributes = package.workbook_sheets_attributes.clone();
    if !package
        .workbook_template
        .declares_namespace(&relationship_prefix, &relationship_namespace)
        && !sheets_attributes.iter().any(|attribute| {
            attribute.name == format!("xmlns:{relationship_prefix}")
                && attribute.value == relationship_namespace
        })
    {
        sheets_attributes.push(XmlAttribute {
            name: format!("xmlns:{relationship_prefix}"),
            value: relationship_namespace,
        });
    }
    let relationship_id_name = format!("{relationship_prefix}:id");
    let sheets_fragment = Some(fragment(|writer| {
        let mut element = BytesStart::new("sheets");
        for attribute in &sheets_attributes {
            element.push_attribute((attribute.name.as_str(), attribute.value.as_str()));
        }
        writer.write_event(Event::Start(element))?;
        for (sheet, plan) in wb.sheets.iter().zip(sheets) {
            let mut attributes = plan.attributes.clone();
            set_attribute(&mut attributes, "name", "name", sheet.name.clone());
            set_attribute(
                &mut attributes,
                "sheetId",
                "sheetId",
                plan.sheet_id.to_string(),
            );
            set_attribute(
                &mut attributes,
                "id",
                &relationship_id_name,
                plan.relationship.id().unwrap_or_default().to_owned(),
            );
            write_empty_element(writer, "sheet", &attributes)?;
        }
        writer.write_event(Event::End(BytesEnd::new("sheets")))?;
        Ok(())
    })?);
    let calc_pr = if edited {
        let mut attributes = package.calc_pr_attributes.clone().unwrap_or_default();
        set_attribute(
            &mut attributes,
            "fullCalcOnLoad",
            "fullCalcOnLoad",
            "1".to_owned(),
        );
        Some(fragment(|writer| {
            write_empty_element(writer, "calcPr", &attributes)
        })?)
    } else {
        package
            .workbook_template
            .child("calcPr")
            .map(|child| child.bytes.clone())
    };
    let mut replacements = vec![
        ("workbookPr", workbook_pr),
        ("sheets", sheets_fragment),
        ("calcPr", calc_pr),
    ];
    if let Some(child) = package.workbook_template.child("definedNames") {
        replacements.push((
            "definedNames",
            render_defined_names(&child.bytes, wb, package)?,
        ));
    }
    package
        .workbook_template
        .render(replacements, workbook_child_rank)
}

/// Patches retained `definedName` elements from the model, which the sheet ops
/// have already rescoped, rewritten or dropped. Source entries the model no
/// longer carries are dropped; every other byte of the element survives.
fn render_defined_names(
    fragment: &[u8],
    wb: &Workbook,
    package: &PreservedPackage,
) -> Result<Option<Vec<u8>>, ParseError> {
    if wb.defined_names == package.original_workbook.defined_names {
        return Ok(Some(fragment.to_vec()));
    }
    let template = XmlTemplate::capture(fragment)?;
    let sources = template.children_named("definedName").collect::<Vec<_>>();
    if sources.len() != package.original_workbook.defined_names.len() {
        return Ok(Some(fragment.to_vec()));
    }
    let mut current = wb.defined_names.iter().peekable();
    let mut replacements = Vec::new();
    for (source, original) in sources.iter().zip(&package.original_workbook.defined_names) {
        let Some(defined) = current
            .peek()
            .filter(|defined| defined.name == original.name)
        else {
            continue;
        };
        let defined = *defined;
        current.next();
        if defined == original {
            replacements.push(source.bytes.clone());
            continue;
        }
        replacements.push(template.qualify_fragment(&patch_defined_name(&source.bytes, defined)?)?);
    }
    if replacements.is_empty() {
        return Ok(None);
    }
    Ok(Some(template.render_repeated(
        "definedName",
        &replacements,
        &[],
    )?))
}

fn patch_defined_name(
    source: &[u8],
    defined: &xlsx_model::DefinedName,
) -> Result<Vec<u8>, ParseError> {
    let mut attributes = attributes_from_fragment(source)?;
    match defined.local_sheet {
        Some(sheet) => set_attribute(
            &mut attributes,
            "localSheetId",
            "localSheetId",
            sheet.0.to_string(),
        ),
        None => remove_attribute(&mut attributes, "localSheetId"),
    }
    fragment(|writer| {
        let mut element = BytesStart::new("definedName");
        for attribute in &attributes {
            element.push_attribute((attribute.name.as_str(), attribute.value.as_str()));
        }
        writer.write_event(Event::Start(element))?;
        writer.write_event(Event::Text(BytesText::new(&defined.formula)))?;
        writer.write_event(Event::End(BytesEnd::new("definedName")))?;
        Ok(())
    })
}

fn workbook_child_rank(name: &str) -> usize {
    match name {
        "fileVersion" => 0,
        "fileSharing" => 1,
        "workbookPr" => 2,
        "workbookProtection" => 3,
        "bookViews" => 4,
        "sheets" => 5,
        "functionGroups" => 6,
        "externalReferences" => 7,
        "definedNames" => 8,
        "calcPr" => 9,
        "oleSize" => 10,
        "customWorkbookViews" => 11,
        "pivotCaches" => 12,
        "smartTagPr" => 13,
        "smartTagTypes" => 14,
        "webPublishing" => 15,
        "fileRecoveryPr" => 16,
        "webPublishObjects" => 17,
        "extLst" => 18,
        _ => usize::MAX,
    }
}

fn workbook_rels(wb: &Workbook, have_sst: bool, have_styles: bool) -> Result<Vec<u8>, ParseError> {
    doc(|w| {
        w.create_element("Relationships")
            .with_attribute(("xmlns", NS_PKG_REL))
            .write_inner_content(|w| {
                let mut next = wb.sheets.len() + 1;
                for i in 0..wb.sheets.len() {
                    let rid = format!("rId{}", i + 1);
                    let target = format!("worksheets/sheet{}.xml", i + 1);
                    w.create_element("Relationship")
                        .with_attribute(("Id", rid.as_str()))
                        .with_attribute(("Type", REL_WORKSHEET))
                        .with_attribute(("Target", target.as_str()))
                        .write_empty()?;
                }
                if have_sst {
                    let rid = format!("rId{next}");
                    next += 1;
                    w.create_element("Relationship")
                        .with_attribute(("Id", rid.as_str()))
                        .with_attribute(("Type", REL_SST))
                        .with_attribute(("Target", "sharedStrings.xml"))
                        .write_empty()?;
                }
                if have_styles {
                    let rid = format!("rId{next}");
                    next += 1;
                    w.create_element("Relationship")
                        .with_attribute(("Id", rid.as_str()))
                        .with_attribute(("Type", REL_STYLES))
                        .with_attribute(("Target", "styles.xml"))
                        .write_empty()?;
                    let rid = format!("rId{next}");
                    w.create_element("Relationship")
                        .with_attribute(("Id", rid.as_str()))
                        .with_attribute(("Type", REL_THEME))
                        .with_attribute(("Target", "theme/theme1.xml"))
                        .write_empty()?;
                }
                Ok(())
            })?;
        Ok(())
    })
}

fn shared_strings_xml(wb: &Workbook) -> Result<Vec<u8>, ParseError> {
    shared_strings_xml_with_namespace(wb, NS_MAIN)
}

fn shared_strings_xml_with_namespace(
    wb: &Workbook,
    main_namespace: &str,
) -> Result<Vec<u8>, ParseError> {
    let count = wb.shared_strings.len().to_string();
    doc(|w| {
        w.create_element("sst")
            .with_attribute(("xmlns", main_namespace))
            .with_attribute(("count", count.as_str()))
            .with_attribute(("uniqueCount", count.as_str()))
            .write_inner_content(|w| {
                for s in &wb.shared_strings {
                    write_shared_string_item(w, s)?;
                }
                Ok(())
            })?;
        Ok(())
    })
}

fn shared_strings_xml_with_template(
    wb: &Workbook,
    package: &PreservedPackage,
    template: &XmlTemplate,
) -> Result<Vec<u8>, ParseError> {
    let source_items = template.children_named("si").collect::<Vec<_>>();
    let retained = retained_shared_string_items(
        &package.original_workbook.shared_strings,
        source_items.len(),
        &wb.shared_strings,
    );
    let mut items = Vec::with_capacity(wb.shared_strings.len());
    for (value, source) in wb.shared_strings.iter().zip(retained) {
        match source.and_then(|source| source_items.get(source)) {
            Some(source) => items.push(source.bytes.clone()),
            None => {
                let item = fragment(|writer| write_shared_string_item(writer, value))?;
                items.push(template.qualify_fragment(&item)?);
            }
        }
    }
    let count = wb.shared_strings.len().to_string();
    template.render_repeated(
        "si",
        &items,
        &[("count", count.clone()), ("uniqueCount", count)],
    )
}

/// Pairs each output shared string with the source `<si>` whose markup it may
/// reuse. Source items are claimed by plain value, in order and at most once,
/// so inserts, deletions and reorders carry rich runs, `phoneticPr` and
/// `extLst` along instead of stranding them on a stale index. A value with no
/// unclaimed source item is regenerated as plain text.
fn retained_shared_string_items(
    original: &[String],
    source_count: usize,
    values: &[String],
) -> Vec<Option<usize>> {
    let mut unclaimed: HashMap<&str, VecDeque<usize>> = HashMap::new();
    for (index, value) in original.iter().enumerate().take(source_count) {
        unclaimed
            .entry(value.as_str())
            .or_default()
            .push_back(index);
    }
    values
        .iter()
        .map(|value| {
            unclaimed
                .get_mut(value.as_str())
                .and_then(VecDeque::pop_front)
        })
        .collect()
}

fn write_shared_string_item(writer: &mut Writer<Vec<u8>>, value: &str) -> io::Result<()> {
    writer.create_element("si").write_inner_content(|writer| {
        write_text_el(writer, value)?;
        Ok(())
    })?;
    Ok(())
}

fn worksheet_xml(sheet: &Sheet, wb: &Workbook) -> Result<Vec<u8>, ParseError> {
    worksheet_xml_with_namespace(sheet, wb, NS_MAIN)
}

fn worksheet_xml_with_namespace(
    sheet: &Sheet,
    wb: &Workbook,
    main_namespace: &str,
) -> Result<Vec<u8>, ParseError> {
    doc(|writer| {
        let mut root = BytesStart::new("worksheet");
        root.push_attribute(("xmlns", main_namespace));
        if sheet
            .hyperlinks
            .iter()
            .any(|link| link.external_target.is_some())
        {
            root.push_attribute(("xmlns:r", NS_R));
        }
        writer.write_event(Event::Start(root))?;
        write_sheet_views(writer, sheet)?;
        write_cols(writer, sheet)?;
        write_sheet_data(writer, sheet, wb)?;
        write_merges(writer, sheet)?;
        write_hyperlinks(writer, sheet)?;
        writer.write_event(Event::End(BytesEnd::new("worksheet")))?;
        Ok(())
    })
}

/// Preserved fragments (filters, validations, anchors, hyperlink ranges) keep
/// their source geometry: valid XML, but stale after a row or column edit.
fn worksheet_xml_with_template(
    sheet: &Sheet,
    wb: &Workbook,
    template: &XmlTemplate,
) -> Result<Vec<u8>, ParseError> {
    let columns = (!sheet.col_widths.is_empty())
        .then(|| fragment(|writer| write_cols(writer, sheet)))
        .transpose()?;
    let sheet_data = Some(fragment(|writer| write_sheet_data(writer, sheet, wb))?);
    let merges = (!sheet.merges.is_empty())
        .then(|| fragment(|writer| write_merges(writer, sheet)))
        .transpose()?;
    template.render(
        vec![
            ("cols", columns),
            ("sheetData", sheet_data),
            ("mergeCells", merges),
        ],
        worksheet_child_rank,
    )
}

fn worksheet_child_rank(name: &str) -> usize {
    match name {
        "sheetPr" => 0,
        "dimension" => 1,
        "sheetViews" => 2,
        "sheetFormatPr" => 3,
        "cols" => 4,
        "sheetData" => 5,
        "sheetCalcPr" => 6,
        "sheetProtection" => 7,
        "protectedRanges" => 8,
        "scenarios" => 9,
        "autoFilter" => 10,
        "sortState" => 11,
        "dataConsolidate" => 12,
        "customSheetViews" => 13,
        "mergeCells" => 14,
        "phoneticPr" => 15,
        "conditionalFormatting" => 16,
        "dataValidations" => 17,
        "hyperlinks" => 18,
        "printOptions" => 19,
        "pageMargins" => 20,
        "pageSetup" => 21,
        "headerFooter" => 22,
        "rowBreaks" => 23,
        "colBreaks" => 24,
        "customProperties" => 25,
        "cellWatches" => 26,
        "ignoredErrors" => 27,
        "smartTags" => 28,
        "drawing" => 29,
        "legacyDrawing" => 30,
        "legacyDrawingHF" => 31,
        "picture" => 32,
        "oleObjects" => 33,
        "controls" => 34,
        "webPublishItems" => 35,
        "tableParts" => 36,
        "extLst" => 37,
        _ => usize::MAX,
    }
}

fn write_sheet_data(writer: &mut Writer<Vec<u8>>, sheet: &Sheet, wb: &Workbook) -> io::Result<()> {
    let sst_index: HashMap<&str, usize> = wb
        .shared_strings
        .iter()
        .enumerate()
        .map(|(i, s)| (s.as_str(), i))
        .collect();

    let mut rows: Vec<RowId> = sheet.iter_cells().map(|(r, _)| r.row).collect();
    rows.extend(sheet.row_heights.keys().copied());
    rows.sort_unstable();
    rows.dedup();

    writer
        .create_element("sheetData")
        .write_inner_content(|writer| {
            for &row in &rows {
                write_row(writer, sheet, row, &sst_index)?;
            }
            Ok(())
        })?;
    Ok(())
}

fn write_hyperlinks(w: &mut Writer<Vec<u8>>, sheet: &Sheet) -> io::Result<()> {
    if sheet.hyperlinks.is_empty() {
        return Ok(());
    }
    w.create_element("hyperlinks").write_inner_content(|w| {
        let mut external_index = 0;
        for link in &sheet.hyperlinks {
            let mut element = BytesStart::new("hyperlink");
            let reference = link.range.to_a1();
            element.push_attribute(("ref", reference.as_str()));
            let relationship_id = link.external_target.as_ref().map(|_| {
                external_index += 1;
                format!("rIdHyperlink{external_index}")
            });
            if let Some(relationship_id) = &relationship_id {
                element.push_attribute(("r:id", relationship_id.as_str()));
            }
            if let Some(location) = &link.location {
                element.push_attribute(("location", location.as_str()));
            }
            if let Some(tooltip) = &link.tooltip {
                element.push_attribute(("tooltip", tooltip.as_str()));
            }
            if let Some(display) = &link.display {
                element.push_attribute(("display", display.as_str()));
            }
            w.write_event(Event::Empty(element))?;
        }
        Ok(())
    })?;
    Ok(())
}

fn worksheet_rels_xml(sheet: &Sheet) -> Result<Vec<u8>, ParseError> {
    doc(|w| {
        w.create_element("Relationships")
            .with_attribute(("xmlns", NS_PKG_REL))
            .write_inner_content(|w| {
                let mut external_index = 0;
                for link in &sheet.hyperlinks {
                    let Some(target) = &link.external_target else {
                        continue;
                    };
                    external_index += 1;
                    let relationship_id = format!("rIdHyperlink{external_index}");
                    w.create_element("Relationship")
                        .with_attribute(("Id", relationship_id.as_str()))
                        .with_attribute(("Type", REL_HYPERLINK))
                        .with_attribute(("Target", target.as_str()))
                        .with_attribute(("TargetMode", "External"))
                        .write_empty()?;
                }
                Ok(())
            })?;
        Ok(())
    })
}

fn write_sheet_views(w: &mut Writer<Vec<u8>>, sheet: &Sheet) -> io::Result<()> {
    let Some(pane) = sheet.freeze_pane else {
        return Ok(());
    };
    w.create_element("sheetViews").write_inner_content(|w| {
        w.create_element("sheetView")
            .with_attribute(("workbookViewId", "0"))
            .write_inner_content(|w| {
                let mut element = BytesStart::new("pane");
                let x_split = pane.cols.to_string();
                let y_split = pane.rows.to_string();
                if pane.cols > 0 {
                    element.push_attribute(("xSplit", x_split.as_str()));
                }
                if pane.rows > 0 {
                    element.push_attribute(("ySplit", y_split.as_str()));
                }
                let top_left = CellRef::new(pane.top_left.row, pane.top_left.col).to_a1();
                element.push_attribute(("topLeftCell", top_left.as_str()));
                element.push_attribute(("activePane", active_pane(pane.rows, pane.cols)));
                element.push_attribute(("state", "frozen"));
                w.write_event(Event::Empty(element))?;
                Ok(())
            })?;
        Ok(())
    })?;
    Ok(())
}

fn active_pane(rows: u32, cols: u32) -> &'static str {
    match (rows > 0, cols > 0) {
        (true, true) => "bottomRight",
        (true, false) => "bottomLeft",
        (false, true) => "topRight",
        (false, false) => "topLeft",
    }
}

fn write_cols(w: &mut Writer<Vec<u8>>, sheet: &Sheet) -> io::Result<()> {
    if sheet.col_widths.is_empty() {
        return Ok(());
    }
    w.create_element("cols").write_inner_content(|w| {
        for (&col, &width) in &sheet.col_widths {
            let n = (col as u64 + 1).to_string();
            w.create_element("col")
                .with_attribute(("min", n.as_str()))
                .with_attribute(("max", n.as_str()))
                .with_attribute(("width", fmt_num(width).as_str()))
                .with_attribute(("customWidth", "1"))
                .write_empty()?;
        }
        Ok(())
    })?;
    Ok(())
}

fn write_row(
    w: &mut Writer<Vec<u8>>,
    sheet: &Sheet,
    row: RowId,
    sst_index: &HashMap<&str, usize>,
) -> io::Result<()> {
    let r = (row as u64 + 1).to_string();
    let mut start = BytesStart::new("row");
    start.push_attribute(("r", r.as_str()));
    let ht = sheet.row_heights.get(&row).map(|h| fmt_num(*h));
    if let Some(h) = &ht {
        start.push_attribute(("ht", h.as_str()));
        start.push_attribute(("customHeight", "1"));
    }
    w.write_event(Event::Start(start))?;
    for (addr, cell) in sheet.iter_cells().filter(|(a, _)| a.row == row) {
        write_cell(w, addr, cell, sst_index)?;
    }
    w.write_event(Event::End(BytesEnd::new("row")))?;
    Ok(())
}

/// serialize a single cell, choosing the `t` type and body from its value and
/// whether it carries a formula.
fn write_cell(
    w: &mut Writer<Vec<u8>>,
    addr: CellRef,
    cell: &Cell,
    sst_index: &HashMap<&str, usize>,
) -> io::Result<()> {
    let a1 = addr.to_a1();
    let has_formula = cell.formula.is_some();

    let mut ty: Option<&str> = None;
    let mut value: Option<String> = None;
    let mut inline: Option<String> = None;
    match &cell.value {
        CellValue::Empty => {}
        CellValue::Number { value: n } => value = Some(fmt_num(*n)),
        CellValue::Bool { value: b } => {
            ty = Some("b");
            value = Some(if *b { "1" } else { "0" }.to_string());
        }
        CellValue::Error { value: e } => {
            ty = Some("e");
            value = Some(e.as_str().to_string());
        }
        CellValue::Text { value: s } => {
            if has_formula {
                ty = Some("str");
                value = Some(s.clone());
            } else if let Some(idx) = sst_index.get(s.as_str()) {
                ty = Some("s");
                value = Some(idx.to_string());
            } else {
                ty = Some("inlineStr");
                inline = Some(s.clone());
            }
        }
    }

    let mut start = BytesStart::new("c");
    start.push_attribute(("r", a1.as_str()));
    let style = cell.style.map(|s| s.to_string());
    if let Some(s) = &style {
        start.push_attribute(("s", s.as_str()));
    }
    if let Some(t) = ty {
        start.push_attribute(("t", t));
    }
    w.write_event(Event::Start(start))?;
    if let Some(f) = &cell.formula {
        w.create_element("f")
            .write_text_content(BytesText::new(f))?;
    }
    if let Some(v) = &value {
        w.create_element("v")
            .write_text_content(BytesText::new(v))?;
    } else if let Some(s) = &inline {
        w.create_element("is").write_inner_content(|w| {
            write_text_el(w, s)?;
            Ok(())
        })?;
    }
    w.write_event(Event::End(BytesEnd::new("c")))?;
    Ok(())
}

fn write_merges(w: &mut Writer<Vec<u8>>, sheet: &Sheet) -> io::Result<()> {
    if sheet.merges.is_empty() {
        return Ok(());
    }
    let count = sheet.merges.len().to_string();
    w.create_element("mergeCells")
        .with_attribute(("count", count.as_str()))
        .write_inner_content(|w| {
            for m in &sheet.merges {
                w.create_element("mergeCell")
                    .with_attribute(("ref", m.to_a1().as_str()))
                    .write_empty()?;
            }
            Ok(())
        })?;
    Ok(())
}

/// write a `<t xml:space="preserve">` element so leading/trailing whitespace
/// survives the round-trip.
fn write_text_el(w: &mut Writer<Vec<u8>>, text: &str) -> io::Result<()> {
    w.create_element("t")
        .with_attribute(("xml:space", "preserve"))
        .write_text_content(BytesText::new(text))?;
    Ok(())
}

/// serialize the style tables verbatim; callers building a stylesheet from
/// scratch must include the sml convention entries for excel to accept it.
fn styles_xml(ss: &Stylesheet) -> Result<Vec<u8>, ParseError> {
    styles_xml_with_namespace(ss, NS_MAIN)
}

fn styles_xml_with_namespace(ss: &Stylesheet, main_namespace: &str) -> Result<Vec<u8>, ParseError> {
    doc(|w| {
        w.create_element("styleSheet")
            .with_attribute(("xmlns", main_namespace))
            .write_inner_content(|w| {
                write_num_fmts(w, ss)?;
                write_fonts(w, ss)?;
                write_fills(w, ss)?;
                write_borders(w, ss)?;
                write_cell_xfs(w, ss)?;
                Ok(())
            })?;
        Ok(())
    })
}

fn styles_xml_with_template(
    stylesheet: &Stylesheet,
    template: &XmlTemplate,
) -> Result<Vec<u8>, ParseError> {
    let num_fmts = (!stylesheet.num_fmts.is_empty())
        .then(|| fragment(|writer| write_num_fmts(writer, stylesheet)))
        .transpose()?;
    let fonts = (!stylesheet.fonts.is_empty())
        .then(|| fragment(|writer| write_fonts(writer, stylesheet)))
        .transpose()?;
    let fills = (!stylesheet.fills.is_empty())
        .then(|| fragment(|writer| write_fills(writer, stylesheet)))
        .transpose()?;
    let borders = (!stylesheet.borders.is_empty())
        .then(|| fragment(|writer| write_borders(writer, stylesheet)))
        .transpose()?;
    let cell_xfs = (!stylesheet.cell_xfs.is_empty())
        .then(|| fragment(|writer| write_cell_xfs(writer, stylesheet)))
        .transpose()?;
    template.render(
        vec![
            ("numFmts", num_fmts),
            ("fonts", fonts),
            ("fills", fills),
            ("borders", borders),
            ("cellXfs", cell_xfs),
        ],
        stylesheet_child_rank,
    )
}

fn stylesheet_child_rank(name: &str) -> usize {
    match name {
        "numFmts" => 0,
        "fonts" => 1,
        "fills" => 2,
        "borders" => 3,
        "cellStyleXfs" => 4,
        "cellXfs" => 5,
        "cellStyles" => 6,
        "dxfs" => 7,
        "tableStyles" => 8,
        "colors" => 9,
        "extLst" => 10,
        _ => usize::MAX,
    }
}

fn write_num_fmts(w: &mut Writer<Vec<u8>>, ss: &Stylesheet) -> io::Result<()> {
    if ss.num_fmts.is_empty() {
        return Ok(());
    }
    let count = ss.num_fmts.len().to_string();
    w.create_element("numFmts")
        .with_attribute(("count", count.as_str()))
        .write_inner_content(|w| {
            for (id, code) in &ss.num_fmts {
                let id = id.to_string();
                w.create_element("numFmt")
                    .with_attribute(("numFmtId", id.as_str()))
                    .with_attribute(("formatCode", code.as_str()))
                    .write_empty()?;
            }
            Ok(())
        })?;
    Ok(())
}

fn write_fonts(w: &mut Writer<Vec<u8>>, ss: &Stylesheet) -> io::Result<()> {
    if ss.fonts.is_empty() {
        return Ok(());
    }
    let count = ss.fonts.len().to_string();
    w.create_element("fonts")
        .with_attribute(("count", count.as_str()))
        .write_inner_content(|w| {
            for font in &ss.fonts {
                write_font(w, font)?;
            }
            Ok(())
        })?;
    Ok(())
}

fn write_font(w: &mut Writer<Vec<u8>>, font: &Font) -> io::Result<()> {
    w.create_element("font").write_inner_content(|w| {
        if font.bold {
            w.create_element("b").write_empty()?;
        }
        if font.italic {
            w.create_element("i").write_empty()?;
        }
        if font.underline {
            w.create_element("u").write_empty()?;
        }
        if font.strike {
            w.create_element("strike").write_empty()?;
        }
        if let Some(sz) = font.size_pt {
            w.create_element("sz")
                .with_attribute(("val", fmt_num(sz).as_str()))
                .write_empty()?;
        }
        if let Some(c) = &font.color {
            write_color(w, "color", c)?;
        }
        if let Some(name) = &font.name {
            w.create_element("name")
                .with_attribute(("val", name.as_str()))
                .write_empty()?;
        }
        Ok(())
    })?;
    Ok(())
}

fn write_fills(w: &mut Writer<Vec<u8>>, ss: &Stylesheet) -> io::Result<()> {
    if ss.fills.is_empty() {
        return Ok(());
    }
    let count = ss.fills.len().to_string();
    w.create_element("fills")
        .with_attribute(("count", count.as_str()))
        .write_inner_content(|w| {
            for fill in &ss.fills {
                write_fill(w, fill)?;
            }
            Ok(())
        })?;
    Ok(())
}

fn write_fill(w: &mut Writer<Vec<u8>>, fill: &Fill) -> io::Result<()> {
    w.create_element("fill")
        .write_inner_content(|w| match fill {
            Fill::None => {
                w.create_element("patternFill")
                    .with_attribute(("patternType", "none"))
                    .write_empty()?;
                Ok(())
            }
            Fill::Solid(color) => {
                w.create_element("patternFill")
                    .with_attribute(("patternType", "solid"))
                    .write_inner_content(|w| {
                        write_color(w, "fgColor", color)?;
                        Ok(())
                    })?;
                Ok(())
            }
        })?;
    Ok(())
}

fn write_borders(w: &mut Writer<Vec<u8>>, ss: &Stylesheet) -> io::Result<()> {
    if ss.borders.is_empty() {
        return Ok(());
    }
    let count = ss.borders.len().to_string();
    w.create_element("borders")
        .with_attribute(("count", count.as_str()))
        .write_inner_content(|w| {
            for border in &ss.borders {
                write_border(w, border)?;
            }
            Ok(())
        })?;
    Ok(())
}

fn write_border(w: &mut Writer<Vec<u8>>, border: &Border) -> io::Result<()> {
    w.create_element("border").write_inner_content(|w| {
        write_edge(w, "left", &border.left)?;
        write_edge(w, "right", &border.right)?;
        write_edge(w, "top", &border.top)?;
        write_edge(w, "bottom", &border.bottom)?;
        w.create_element("diagonal").write_empty()?;
        Ok(())
    })?;
    Ok(())
}

fn write_edge(w: &mut Writer<Vec<u8>>, name: &str, edge: &Option<BorderEdge>) -> io::Result<()> {
    match edge {
        None => {
            w.create_element(name).write_empty()?;
        }
        Some(ed) => {
            w.create_element(name)
                .with_attribute(("style", ed.style.as_sml()))
                .write_inner_content(|w| {
                    if let Some(c) = &ed.color {
                        write_color(w, "color", c)?;
                    }
                    Ok(())
                })?;
        }
    }
    Ok(())
}

fn write_cell_xfs(w: &mut Writer<Vec<u8>>, ss: &Stylesheet) -> io::Result<()> {
    if ss.cell_xfs.is_empty() {
        return Ok(());
    }
    let count = ss.cell_xfs.len().to_string();
    w.create_element("cellXfs")
        .with_attribute(("count", count.as_str()))
        .write_inner_content(|w| {
            for xf in &ss.cell_xfs {
                write_xf(w, xf)?;
            }
            Ok(())
        })?;
    Ok(())
}

/// write one cellXfs `<xf>`. unset facets serialize as index 0 with no
/// `applyX` flag, so the reader restores them to `None`.
fn write_xf(w: &mut Writer<Vec<u8>>, xf: &Xf) -> io::Result<()> {
    let num_fmt_id = xf.num_fmt_id.unwrap_or(0).to_string();
    let font_id = xf.font.unwrap_or(0).to_string();
    let fill_id = xf.fill.unwrap_or(0).to_string();
    let border_id = xf.border.unwrap_or(0).to_string();

    let mut el = BytesStart::new("xf");
    el.push_attribute(("numFmtId", num_fmt_id.as_str()));
    el.push_attribute(("fontId", font_id.as_str()));
    el.push_attribute(("fillId", fill_id.as_str()));
    el.push_attribute(("borderId", border_id.as_str()));
    if xf.num_fmt_id.is_some() {
        el.push_attribute(("applyNumberFormat", "1"));
    }
    if xf.font.is_some() {
        el.push_attribute(("applyFont", "1"));
    }
    if xf.fill.is_some() {
        el.push_attribute(("applyFill", "1"));
    }
    if xf.border.is_some() {
        el.push_attribute(("applyBorder", "1"));
    }
    if xf.alignment.is_some() {
        el.push_attribute(("applyAlignment", "1"));
    }

    match &xf.alignment {
        None => w.write_event(Event::Empty(el))?,
        Some(a) => {
            w.write_event(Event::Start(el))?;
            write_alignment(w, a)?;
            w.write_event(Event::End(BytesEnd::new("xf")))?;
        }
    }
    Ok(())
}

fn write_alignment(w: &mut Writer<Vec<u8>>, a: &Alignment) -> io::Result<()> {
    let mut el = BytesStart::new("alignment");
    if let Some(h) = a.h {
        el.push_attribute(("horizontal", h.as_sml()));
    }
    if let Some(v) = a.v {
        el.push_attribute(("vertical", v.as_sml()));
    }
    if a.wrap_text {
        el.push_attribute(("wrapText", "1"));
    }
    if a.shrink_to_fit {
        el.push_attribute(("shrinkToFit", "1"));
    }
    w.write_event(Event::Empty(el))?;
    Ok(())
}

/// write a `CT_Color` element carrying whichever representation the `Color`
/// holds. rgb is emitted as `FFrrggbb` (opaque).
fn write_color(w: &mut Writer<Vec<u8>>, name: &str, color: &Color) -> io::Result<()> {
    let mut el = BytesStart::new(name.to_string());
    let rgb;
    let idx;
    let theme;
    let tint;
    match color {
        Color::Rgb(hex) => {
            rgb = format!("FF{}", hex.trim_start_matches('#').to_ascii_uppercase());
            el.push_attribute(("rgb", rgb.as_str()));
        }
        Color::Indexed(i) => {
            idx = i.to_string();
            el.push_attribute(("indexed", idx.as_str()));
        }
        Color::Theme { idx: i, tint: t } => {
            theme = i.to_string();
            el.push_attribute(("theme", theme.as_str()));
            if *t != 0.0 {
                tint = format!("{t}");
                el.push_attribute(("tint", tint.as_str()));
            }
        }
        Color::Auto => {
            el.push_attribute(("auto", "1"));
        }
    }
    w.write_event(Event::Empty(el))?;
    Ok(())
}

/// emit a minimal but schema-shaped `theme1.xml`: the 12-color clrScheme plus
/// stub font/format schemes so excel accepts the part.
fn theme_xml(ss: &Stylesheet) -> Result<Vec<u8>, ParseError> {
    let c = &ss.theme.colors;
    let slots = [
        "dk1", "lt1", "dk2", "lt2", "accent1", "accent2", "accent3", "accent4", "accent5",
        "accent6", "hlink", "folHlink",
    ];
    doc(|w| {
        w.create_element("a:theme")
            .with_attribute(("xmlns:a", NS_DML))
            .with_attribute(("name", "Office Theme"))
            .write_inner_content(|w| {
                w.create_element("a:themeElements")
                    .write_inner_content(|w| {
                        w.create_element("a:clrScheme")
                            .with_attribute(("name", "Office"))
                            .write_inner_content(|w| {
                                for (slot, hex) in slots.iter().zip(c.iter()) {
                                    let val = hex.trim_start_matches('#').to_ascii_uppercase();
                                    w.create_element(format!("a:{slot}")).write_inner_content(
                                        |w| {
                                            w.create_element("a:srgbClr")
                                                .with_attribute(("val", val.as_str()))
                                                .write_empty()?;
                                            Ok(())
                                        },
                                    )?;
                                }
                                Ok(())
                            })?;
                        write_stub_font_scheme(w)?;
                        write_stub_fmt_scheme(w)?;
                        Ok(())
                    })?;
                Ok(())
            })?;
        Ok(())
    })
}

/// a minimal `a:fontScheme` (major/minor latin only) so the theme validates.
fn write_stub_font_scheme(w: &mut Writer<Vec<u8>>) -> io::Result<()> {
    w.create_element("a:fontScheme")
        .with_attribute(("name", "Office"))
        .write_inner_content(|w| {
            for major_minor in ["a:majorFont", "a:minorFont"] {
                w.create_element(major_minor).write_inner_content(|w| {
                    w.create_element("a:latin")
                        .with_attribute(("typeface", "Calibri"))
                        .write_empty()?;
                    w.create_element("a:ea")
                        .with_attribute(("typeface", ""))
                        .write_empty()?;
                    w.create_element("a:cs")
                        .with_attribute(("typeface", ""))
                        .write_empty()?;
                    Ok(())
                })?;
            }
            Ok(())
        })?;
    Ok(())
}

/// a minimal `a:fmtScheme`; excel requires the element even though we do not
/// model fill/line/effect styles.
fn write_stub_fmt_scheme(w: &mut Writer<Vec<u8>>) -> io::Result<()> {
    w.create_element("a:fmtScheme")
        .with_attribute(("name", "Office"))
        .write_inner_content(|w| {
            w.create_element("a:fillStyleLst")
                .write_inner_content(|w| {
                    for _ in 0..3 {
                        w.create_element("a:solidFill").write_inner_content(|w| {
                            w.create_element("a:schemeClr")
                                .with_attribute(("val", "phClr"))
                                .write_empty()?;
                            Ok(())
                        })?;
                    }
                    Ok(())
                })?;
            w.create_element("a:lnStyleLst").write_inner_content(|w| {
                for _ in 0..3 {
                    w.create_element("a:ln").write_empty()?;
                }
                Ok(())
            })?;
            w.create_element("a:effectStyleLst")
                .write_inner_content(|w| {
                    for _ in 0..3 {
                        w.create_element("a:effectStyle").write_inner_content(|w| {
                            w.create_element("a:effectLst").write_empty()?;
                            Ok(())
                        })?;
                    }
                    Ok(())
                })?;
            w.create_element("a:bgFillStyleLst")
                .write_inner_content(|w| {
                    for _ in 0..3 {
                        w.create_element("a:solidFill").write_inner_content(|w| {
                            w.create_element("a:schemeClr")
                                .with_attribute(("val", "phClr"))
                                .write_empty()?;
                            Ok(())
                        })?;
                    }
                    Ok(())
                })?;
            Ok(())
        })?;
    Ok(())
}

/// format a number the way excel writes cell values: integers without a
/// trailing `.0`.
fn fmt_num(n: f64) -> String {
    if n.is_finite() && n.fract() == 0.0 && n.abs() < 1e15 {
        return format!("{}", n as i64);
    }
    format!("{n}")
}
