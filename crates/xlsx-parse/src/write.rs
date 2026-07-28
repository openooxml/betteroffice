//! `xlsx_model::Workbook` -> xlsx parts. Without a source package this writes a
//! minimal valid workbook. With one, parts and sheets the model did not change
//! are copied through byte for byte; only what changed is reserialized, and
//! that reserialization carries just the subset `read` models.

use std::collections::{BTreeMap, HashMap, HashSet, VecDeque};
use std::io;

use quick_xml::Writer;
use quick_xml::events::{BytesDecl, BytesEnd, BytesStart, BytesText, Event};
use xlsx_model::addr::RowId;
use xlsx_model::styles::{Alignment, Border, BorderEdge, Color, Fill, Font, Stylesheet, Xf};
use xlsx_model::{Cell, CellRef, CellValue, ChartAnchor, ChartRef, DateSystem, Sheet, Workbook};

use crate::ParseError;
use crate::package::{
    ContentTypeEntry, PartReference, PreservedPackage, PreservedSheet, Relationship, XmlAttribute,
    XmlTemplate, attributes_from_fragment, effective_content_type, normalized_part_name,
    parse_relationships, relationship_part_path, remove_attribute, set_attribute,
};
use crate::read::SharedStringCells;
use crate::xml::{resolve_part_path, xml_err};

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
const NS_STRICT_DML: &str = "http://purl.oclc.org/ooxml/drawingml/main";

/// Generated relationship attributes always use this prefix. A source prefix
/// is caller-controlled and would be repeated on every generated sheet and
/// hyperlink, turning a long one into amplified output.
const GENERATED_REL_PREFIX: &str = "r";
const REL_PREFIX_DECLARATION: &str = "xmlns:r";
const REL_ID_ATTRIBUTE: &str = "r:id";

/// Bound on the relationship URI a save copies out of the source. Real ones are
/// under a hundred bytes; a larger one is rejected before any fragment that
/// would repeat it is built.
const MAX_RELATIONSHIP_NAMESPACE_BYTES: usize = 1024;

/// A Strict package must not gain a Transitional DrawingML theme.
fn drawingml_namespace(main_namespace: &str) -> &'static str {
    if main_namespace == NS_STRICT_MAIN {
        NS_STRICT_DML
    } else {
        NS_DML
    }
}

/// serialize a workbook to opc parts in a fixed, deterministic order. a chart
/// is refused: this writer emits no drawing, chart relationship or chart part,
/// so it would drop one rather than write it.
pub fn serialize_workbook(wb: &Workbook) -> Result<Vec<(String, Vec<u8>)>, ParseError> {
    if wb.sheets.iter().any(|sheet| !sheet.charts.is_empty()) {
        return Err(ParseError::UnsupportedEdit(
            "a chart can only be written back into the package it was read from".to_owned(),
        ));
    }
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
        parts.push((
            "xl/theme/theme1.xml".to_string(),
            theme_xml(&wb.styles, NS_DML)?,
        ));
    }
    for (i, sheet) in wb.sheets.iter().enumerate() {
        let links = HyperlinkPlan::new(sheet, &[], REL_HYPERLINK);
        parts.push((
            format!("xl/worksheets/sheet{}.xml", i + 1),
            worksheet_xml_with_namespace(sheet, wb, NS_MAIN, NS_R, &links, None)?,
        ));
        if !links.relationships.is_empty() {
            parts.push((
                format!("xl/worksheets/_rels/sheet{}.xml.rels", i + 1),
                relationships_xml(&links.relationships)?,
            ));
        }
    }
    Ok(parts)
}

/// Serializes owned parts over a preserved source package, guessing each
/// sheet's origin from its current index and its provenance from the source.
/// Only correct when nothing structural moved, so it is a test convenience
/// rather than an entry point.
#[cfg(test)]
pub(crate) fn serialize_workbook_with_package(
    wb: &Workbook,
    package: &PreservedPackage,
) -> Result<Vec<(String, Vec<u8>)>, ParseError> {
    let origins = (0..wb.sheets.len())
        .map(|index| (index < package.sheets.len()).then_some(index))
        .collect::<Vec<_>>();
    serialize_workbook_with_package_and_origins(wb, package, &origins)
}

#[cfg(test)]
pub(crate) fn serialize_workbook_with_package_and_origins(
    wb: &Workbook,
    package: &PreservedPackage,
    origins: &[Option<usize>],
) -> Result<Vec<(String, Vec<u8>)>, ParseError> {
    let changed = wb != &package.original_workbook
        || origins.len() != package.sheets.len()
        || origins
            .iter()
            .enumerate()
            .any(|(index, origin)| *origin != Some(index));
    let shared_string_cells = origins
        .iter()
        .map(|origin| {
            origin
                .map(|origin| package.source_shared_string_cells(origin))
                .unwrap_or_default()
        })
        .collect::<Vec<_>>();
    serialize_workbook_with_package_and_origins_after_edits(
        wb,
        package,
        origins,
        &shared_string_cells,
        SaveEdits {
            changed,
            moved_references: false,
        },
    )
}

/// What a save must be told about the edits behind it, because neither the
/// model nor the source package carries it.
#[doc(hidden)]
#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
pub struct SaveEdits {
    /// the model differs from the one the package was read with. False means
    /// the caller guarantees an untouched model, which is returned as the
    /// source bytes; true drops the now-stale calculation chain.
    pub changed: bool,
    /// an edit moved cell addresses, so every preserved part naming them must
    /// be one this crate can rewrite.
    pub moved_references: bool,
}

/// `origins` and `shared_string_cells` are indexed by current sheet and carry
/// what the model does not: the source sheet each one came from and the shared
/// string entry each of its cells was authored against.
#[doc(hidden)]
pub fn serialize_workbook_with_package_and_origins_after_edits(
    wb: &Workbook,
    package: &PreservedPackage,
    origins: &[Option<usize>],
    shared_string_cells: &[SharedStringCells],
    edits: SaveEdits,
) -> Result<Vec<(String, Vec<u8>)>, ParseError> {
    if origins.len() != wb.sheets.len() || shared_string_cells.len() != wb.sheets.len() {
        return Err(ParseError::Malformed(
            "sheet origin count does not match workbook".to_owned(),
        ));
    }
    let edited = edits.changed;
    if !edited
        && origins.len() == package.sheets.len()
        && origins
            .iter()
            .enumerate()
            .all(|(index, origin)| *origin == Some(index))
    {
        return Ok(package.parts.clone());
    }
    preflight_preserved_save(wb, package, origins, edits)?;

    let have_sst = !wb.shared_strings.is_empty();
    let have_styles = !wb.styles.is_empty() || package.styles.is_some();
    let main_namespace = package
        .workbook_template
        .root_namespace()
        .unwrap_or(NS_MAIN);
    let relationship_namespace = workbook_relationship_namespace(package)?;
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
    let shared_string_plan = package
        .shared_strings_template
        .as_ref()
        .map(|template| {
            SharedStringPlan::new(
                &package.original_workbook.shared_strings,
                template.children_named("si").count(),
                wb,
                shared_string_cells,
            )
        })
        .transpose()?;
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
    let synthesized_styles = package.styles.is_none() && !wb.styles.is_empty();
    let theme = (package.theme.is_some() || synthesized_styles).then(|| {
        plan_part(
            package.theme.as_ref(),
            &package.workbook_relationships,
            "theme",
            &theme_relationship_type,
            "xl/theme/theme1.xml",
            &mut used_relationship_ids,
        )
    });

    let mut parts = PartStore::new(&package.parts);
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
    patch_chart_parts(wb, package, origins, &mut parts)?;

    let shared_strings_stable = wb.shared_strings == package.original_workbook.shared_strings;
    let empty_provenance = SharedStringCells::new();
    for (index, (sheet, plan)) in wb.sheets.iter().zip(&sheets).enumerate() {
        let source = plan.origin.and_then(|origin| package.sheets.get(origin));
        let original = plan
            .origin
            .and_then(|origin| package.original_workbook.sheets.get(origin));
        let provenance = shared_string_cells.get(index).unwrap_or(&empty_provenance);
        let output = match source {
            Some(source) if source.is_worksheet() => {
                if shared_strings_stable
                    && original.is_some_and(|original| sheet_body_matches(sheet, original))
                {
                    continue;
                }
                worksheet_xml_with_template(
                    sheet,
                    wb,
                    original,
                    source,
                    package,
                    provenance,
                    shared_string_plan.as_ref(),
                )?
            }
            Some(_) => continue,
            None => {
                let links = HyperlinkPlan::new(
                    sheet,
                    &[],
                    &relationship_type_from(
                        &package.workbook_relationships,
                        "hyperlink",
                        REL_HYPERLINK,
                    ),
                );
                WorksheetOutput {
                    bytes: worksheet_xml_with_namespace(
                        sheet,
                        wb,
                        main_namespace,
                        &relationship_namespace,
                        &links,
                        shared_string_plan.as_ref(),
                    )?,
                    relationships: Some(links.relationships),
                }
            }
        };
        parts.set(plan.path.clone(), output.bytes)?;
        if let Some(relationships) = output.relationships {
            let path = relationship_part_path(&plan.path);
            if relationships.is_empty() {
                parts.remove(&path);
            } else {
                parts.set(path, relationships_xml(&relationships)?)?;
            }
        }
    }

    let workbook = workbook_xml_with_template(wb, package, &sheets, edited)?;
    parts.set("xl/workbook.xml".to_owned(), workbook)?;
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
    )?;
    parts.set(
        "_rels/.rels".to_owned(),
        merged_root_relationships(package)?,
    )?;
    let pruned = prune_unreachable_parts(&mut parts, package)?;
    parts.set(
        "[Content_Types].xml".to_owned(),
        merged_content_types(
            package,
            &sheets,
            shared_strings.as_ref(),
            styles.as_ref(),
            theme.as_ref(),
            edited,
            &pruned,
        )?,
    )?;

    replace_shared_strings(
        &mut parts,
        wb,
        package,
        shared_strings.as_ref(),
        main_namespace,
        shared_string_plan.as_ref(),
    )?;
    let styles_retained = matches!(
        (package.styles.as_ref(), styles.as_ref()),
        (Some(source), Some(planned)) if source.path == planned.path
    ) && wb.styles == package.original_workbook.styles;
    if !styles_retained {
        replace_optional_part(&mut parts, package.styles.as_ref(), styles.as_ref(), || {
            match &package.stylesheet_template {
                Some(template) => styles_xml_with_template(
                    &wb.styles,
                    &package.original_workbook.styles,
                    template,
                ),
                None => styles_xml_with_namespace(&wb.styles, main_namespace),
            }
        })?;
    }

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
            parts.set(
                planned.path.clone(),
                theme_xml(&wb.styles, drawingml_namespace(main_namespace))?,
            )?;
        }
        (Some(source), None) => parts.remove(&source.path),
        (None, None) => {}
    }

    Ok(parts.finish())
}

/// Writes moved chart references and anchors back into the parts that hold
/// them. Each part is patched once, from its source bytes, so unmodelled chart
/// markup survives byte for byte.
fn patch_chart_parts(
    wb: &Workbook,
    package: &PreservedPackage,
    origins: &[Option<usize>],
    parts: &mut PartStore<'_>,
) -> Result<(), ParseError> {
    let mut chart_refs: BTreeMap<&str, (&[ChartRef], &str)> = BTreeMap::new();
    let mut anchors: BTreeMap<&str, Vec<(usize, ChartAnchor)>> = BTreeMap::new();
    for (sheet, origin) in wb.sheets.iter().zip(origins) {
        let source = origin
            .and_then(|origin| package.original_workbook.sheets.get(origin))
            .map(|sheet| sheet.charts.as_slice())
            .unwrap_or_default();
        for chart in &sheet.charts {
            let original = source
                .iter()
                .find(|other| other.part == chart.part)
                .ok_or_else(|| unwritable_chart(&chart.part, &sheet.name))?;
            if original.drawing != chart.drawing || original.anchor_index != chart.anchor_index {
                return Err(ParseError::UnsupportedEdit(format!(
                    "chart {} no longer names the drawing and anchor it was read from",
                    chart.part
                )));
            }
            if original.refs.len() != chart.refs.len()
                || original
                    .refs
                    .iter()
                    .zip(&chart.refs)
                    .any(|(original, edited)| original.kind != edited.kind)
            {
                return Err(ParseError::UnsupportedEdit(format!(
                    "chart {} no longer holds the references it was read from",
                    chart.part
                )));
            }
            if original.refs != chart.refs {
                chart_refs.insert(
                    chart.part.as_str(),
                    (chart.refs.as_slice(), sheet.name.as_str()),
                );
            }
            if original.anchor != chart.anchor {
                anchors
                    .entry(chart.drawing.as_str())
                    .or_default()
                    .push((chart.anchor_index, chart.anchor));
            }
        }
    }
    for (path, (refs, owner)) in chart_refs {
        let source = package.part_bytes(path).ok_or_else(|| missing_part(path))?;
        parts.set(
            path.to_owned(),
            crate::chart::patch_chart_refs(source, refs, wb, owner)?,
        )?;
    }
    for (path, moved) in anchors {
        let source = package.part_bytes(path).ok_or_else(|| missing_part(path))?;
        parts.set(
            path.to_owned(),
            crate::chart::patch_drawing_anchors(source, &moved)?,
        )?;
    }
    Ok(())
}

fn missing_part(path: &str) -> ParseError {
    ParseError::UnsupportedEdit(format!(
        "{path} carries chart state to write back but is not in the source package"
    ))
}

fn unwritable_chart(part: &str, sheet: &str) -> ParseError {
    ParseError::UnsupportedEdit(format!(
        "chart {part} was not read from sheet {sheet}, and this crate cannot create one"
    ))
}

/// Everything a save over a preserved package must satisfy before a byte is
/// written: no part this crate cannot rewrite is stranded, and every chart the
/// model carries lines up one for one with the source it must be patched from.
fn preflight_preserved_save(
    wb: &Workbook,
    package: &PreservedPackage,
    origins: &[Option<usize>],
    edits: SaveEdits,
) -> Result<(), ParseError> {
    ensure_preserved_references_stay_valid(wb, package, origins, edits)?;
    ensure_chart_provenance(wb, package, origins)
}

/// Charts, pivot caches and their friends name sheets and cells this crate
/// never patches. Moving cells, or dropping, reordering or renaming a source
/// sheet while one of them is preserved strands it, so the serializer refuses
/// that here instead of trusting every caller to have checked.
fn ensure_preserved_references_stay_valid(
    wb: &Workbook,
    package: &PreservedPackage,
    origins: &[Option<usize>],
    edits: SaveEdits,
) -> Result<(), ParseError> {
    let Some(part) = package.unpatchable_reference_part() else {
        return Ok(());
    };
    let retained = origins.iter().flatten().copied().collect::<Vec<_>>();
    let kept_in_order = retained.len() == package.sheets.len()
        && retained
            .iter()
            .enumerate()
            .all(|(position, origin)| *origin == position);
    let renamed = origins.iter().zip(&wb.sheets).any(|(origin, sheet)| {
        origin.is_some_and(|origin| {
            package
                .original_workbook
                .sheets
                .get(origin)
                .is_some_and(|original| original.name != sheet.name)
        })
    });
    if kept_in_order && !renamed && !edits.moved_references {
        return Ok(());
    }
    Err(ParseError::UnsupportedEdit(format!(
        "{part} references sheets or cells this save would move, and it cannot be rewritten"
    )))
}

/// This crate patches chart parts in place; it can neither create one nor
/// delete one. Every chart a retained sheet carries must therefore name
/// exactly one chart the same source sheet was read with, and every chart that
/// source held must still be there.
fn ensure_chart_provenance(
    wb: &Workbook,
    package: &PreservedPackage,
    origins: &[Option<usize>],
) -> Result<(), ParseError> {
    for (sheet, origin) in wb.sheets.iter().zip(origins) {
        let source = origin
            .and_then(|origin| package.original_workbook.sheets.get(origin))
            .map(|sheet| sheet.charts.as_slice())
            .unwrap_or_default();
        for chart in &sheet.charts {
            if source
                .iter()
                .filter(|other| other.part == chart.part)
                .count()
                != 1
            {
                return Err(unwritable_chart(&chart.part, &sheet.name));
            }
        }
        for original in source {
            if !sheet.charts.iter().any(|chart| chart.part == original.part) {
                return Err(ParseError::UnsupportedEdit(format!(
                    "chart {} was dropped from sheet {}, and this crate cannot delete one",
                    original.part, sheet.name
                )));
            }
        }
    }
    Ok(())
}

/// The relationship ids a worksheet's `<hyperlink>` elements point at, beside
/// the worksheet relationship part that must back them. Non-hyperlink source
/// relationships (drawings, comments, tables) are carried through untouched.
struct HyperlinkPlan {
    ids: Vec<Option<String>>,
    relationships: Vec<Relationship>,
}

impl HyperlinkPlan {
    fn new(sheet: &Sheet, source: &[Relationship], fallback_type: &str) -> Self {
        let relationship_type = relationship_type_from(source, "hyperlink", fallback_type);
        let mut relationships = source
            .iter()
            .filter(|relationship| !relationship.has_type("hyperlink"))
            .cloned()
            .collect::<Vec<_>>();
        let mut used = relationships
            .iter()
            .filter_map(Relationship::id)
            .map(str::to_owned)
            .collect::<HashSet<_>>();
        let ids = sheet
            .hyperlinks
            .iter()
            .map(|link| {
                let target = link.external_target.as_ref()?;
                let id = next_relationship_id(&mut used);
                let mut relationship =
                    new_relationship(id.clone(), &relationship_type, target.clone());
                relationship.attributes.push(XmlAttribute {
                    name: "TargetMode".to_owned(),
                    value: "External".to_owned(),
                });
                relationships.push(relationship);
                Some(id)
            })
            .collect();
        Self { ids, relationships }
    }
}

/// Everything a worksheet part carries. The sheet name lives in the workbook
/// part, so a rename leaves the worksheet bytes reusable.
fn sheet_body_matches(sheet: &Sheet, original: &Sheet) -> bool {
    sheet.freeze_pane == original.freeze_pane
        && sheet.hyperlinks == original.hyperlinks
        && sheet.merges == original.merges
        && sheet.col_widths == original.col_widths
        && sheet.row_heights == original.row_heights
        && sheet.iter_cells().eq(original.iter_cells())
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

/// The relationship URI the source uses, Strict or Transitional. Only the URI
/// is taken from the source; the prefix bound to it is always
/// [`GENERATED_REL_PREFIX`].
fn workbook_relationship_namespace(package: &PreservedPackage) -> Result<String, ParseError> {
    if let Some((_, namespace)) = package
        .workbook_template
        .namespace_binding("/officeDocument/relationships")
    {
        return checked_relationship_namespace(namespace);
    }
    let namespace = package
        .workbook_relationships
        .iter()
        .filter_map(|relationship| relationship.attribute("Type"))
        .find_map(|relationship_type| {
            let (namespace, _) = relationship_type.rsplit_once('/')?;
            namespace
                .ends_with("/officeDocument/relationships")
                .then_some(namespace)
        })
        .unwrap_or(NS_R);
    checked_relationship_namespace(namespace)
}

/// The URI a generated fragment binds `r:` to: the part's own declaration when
/// it has one, so a Strict worksheet keeps its own, else the workbook's.
fn fragment_relationship_namespace(
    template: &XmlTemplate,
    package: &PreservedPackage,
) -> Result<String, ParseError> {
    match template.namespace_binding("/officeDocument/relationships") {
        Some((_, namespace)) => checked_relationship_namespace(namespace),
        None => workbook_relationship_namespace(package),
    }
}

fn checked_relationship_namespace(namespace: &str) -> Result<String, ParseError> {
    if namespace.len() > MAX_RELATIONSHIP_NAMESPACE_BYTES {
        return Err(ParseError::UnsupportedEdit(
            "relationship namespace exceeds the generated-attribute cap".to_owned(),
        ));
    }
    Ok(namespace.to_owned())
}

/// Declares [`GENERATED_REL_PREFIX`] on a generated element unless the part's
/// root already binds it to the same URI, replacing any shadowing declaration
/// the source put on that element.
fn bind_relationship_prefix(
    attributes: &mut Vec<XmlAttribute>,
    template: &XmlTemplate,
    namespace: &str,
) {
    if attributes
        .iter()
        .any(|attribute| attribute.name == REL_PREFIX_DECLARATION && attribute.value == namespace)
    {
        return;
    }
    attributes.retain(|attribute| attribute.name != REL_PREFIX_DECLARATION);
    if !template.declares_namespace(GENERATED_REL_PREFIX, namespace) {
        attributes.push(XmlAttribute {
            name: REL_PREFIX_DECLARATION.to_owned(),
            value: namespace.to_owned(),
        });
    }
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
            parts.set(planned.path.clone(), serialize()?)?;
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
    shared_string_plan: Option<&SharedStringPlan>,
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
                Some(template) => shared_strings_xml_with_template(
                    wb,
                    template,
                    shared_string_plan.ok_or_else(|| {
                        ParseError::Malformed("shared string identity plan is missing".to_owned())
                    })?,
                )?,
                None => shared_strings_xml_with_namespace(wb, main_namespace)?,
            };
            parts.set(planned.path.clone(), bytes)?;
        }
        (Some(source), None) => parts.remove(&source.path),
        (None, None) => {}
    }
    Ok(())
}

/// Headroom over the source package for everything a save regenerates. A
/// bounded input must not turn into an unbounded output.
const GENERATED_BYTES_HEADROOM: usize = 64 * 1024 * 1024;

enum Slot<'a> {
    Source(&'a str, &'a [u8]),
    Owned(String, Vec<u8>),
    Removed,
}

/// Borrows the source package so a save holds one copy of the retained bytes,
/// and refuses to generate unboundedly more than it started with.
struct PartStore<'a> {
    slots: Vec<Slot<'a>>,
    positions: HashMap<String, usize>,
    budget: usize,
}

impl<'a> PartStore<'a> {
    fn new(parts: &'a [(String, Vec<u8>)]) -> Self {
        let source_bytes = parts.iter().map(|(_, bytes)| bytes.len()).sum::<usize>();
        Self {
            slots: parts
                .iter()
                .map(|(path, bytes)| Slot::Source(path.as_str(), bytes.as_slice()))
                .collect(),
            positions: parts
                .iter()
                .enumerate()
                .map(|(index, (path, _))| (normalized_part_name(path), index))
                .collect(),
            budget: source_bytes
                .saturating_mul(2)
                .saturating_add(GENERATED_BYTES_HEADROOM),
        }
    }

    fn remove(&mut self, path: &str) {
        if let Some(index) = self.positions.remove(&normalized_part_name(path)) {
            self.slots[index] = Slot::Removed;
        }
    }

    fn set(&mut self, path: String, bytes: Vec<u8>) -> Result<(), ParseError> {
        self.budget = self.budget.checked_sub(bytes.len()).ok_or_else(|| {
            ParseError::UnsupportedEdit("generated parts exceeded the save budget".to_owned())
        })?;
        let normalized = normalized_part_name(&path);
        match self.positions.get(&normalized).copied() {
            Some(index) => {
                let authored = match &self.slots[index] {
                    Slot::Source(path, _) => (*path).to_owned(),
                    Slot::Owned(path, _) => path.clone(),
                    Slot::Removed => path,
                };
                self.slots[index] = Slot::Owned(authored, bytes);
            }
            None => {
                self.positions.insert(normalized, self.slots.len());
                self.slots.push(Slot::Owned(path, bytes));
            }
        }
        Ok(())
    }

    fn entries(&self) -> Vec<(&str, &[u8])> {
        self.slots
            .iter()
            .filter_map(|slot| match slot {
                Slot::Source(path, bytes) => Some((*path, *bytes)),
                Slot::Owned(path, bytes) => Some((path.as_str(), bytes.as_slice())),
                Slot::Removed => None,
            })
            .collect()
    }

    fn finish(self) -> Vec<(String, Vec<u8>)> {
        self.slots
            .into_iter()
            .filter_map(|slot| match slot {
                Slot::Source(path, bytes) => Some((path.to_owned(), bytes.to_vec())),
                Slot::Owned(path, bytes) => Some((path, bytes)),
                Slot::Removed => None,
            })
            .collect()
    }
}

/// Drops the parts this save left unreachable from the package roots: a
/// removed sheet's drawing, its chart, their caches and the images only they
/// referenced. A part the source package already held unreachable is left
/// alone, and one still reachable through another relationship stays.
fn prune_unreachable_parts(
    parts: &mut PartStore<'_>,
    package: &PreservedPackage,
) -> Result<HashSet<String>, ParseError> {
    let before = reachable_parts(
        &package
            .parts
            .iter()
            .map(|(path, bytes)| (path.as_str(), bytes.as_slice()))
            .collect::<Vec<_>>(),
    )?;
    let current = parts.entries();
    let after = reachable_parts(&current)?;
    let pruned = current
        .iter()
        .map(|(path, _)| normalized_part_name(path))
        .filter(|path| before.contains(path) && !after.contains(path))
        .collect::<HashSet<_>>();
    for path in &pruned {
        parts.remove(path);
    }
    Ok(pruned)
}

/// Every part reachable from `_rels/.rels` by following internal
/// relationships, plus the `.rels` parts that carry them.
fn reachable_parts(parts: &[(&str, &[u8])]) -> Result<HashSet<String>, ParseError> {
    let lookup = parts
        .iter()
        .map(|(path, bytes)| (normalized_part_name(path), *bytes))
        .collect::<HashMap<_, _>>();
    let mut reachable = HashSet::new();
    reachable.insert(normalized_part_name("[Content_Types].xml"));
    let root = normalized_part_name("_rels/.rels");
    let Some(bytes) = lookup.get(&root) else {
        return Ok(reachable);
    };
    reachable.insert(root);
    let mut queue = crate::chart::parse_relationships(bytes)?
        .into_iter()
        .map(|(_, _, target)| resolve_part_path("", &target))
        .collect::<VecDeque<String>>();
    while let Some(path) = queue.pop_front() {
        let normalized = normalized_part_name(&path);
        if !reachable.insert(normalized.clone()) {
            continue;
        }
        let rels = normalized_part_name(&relationship_part_path(&normalized));
        let Some(bytes) = lookup.get(&rels) else {
            continue;
        };
        reachable.insert(rels);
        let directory = crate::chart::directory_of(&normalized).to_owned();
        for (_, _, target) in crate::chart::parse_relationships(bytes)? {
            queue.push_back(resolve_part_path(&directory, &target));
        }
    }
    Ok(reachable)
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
    pruned: &HashSet<String>,
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
        if (edited && calc_chain_paths.contains(&normalized)) || pruned.contains(&normalized) {
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

fn source_content_type<'a>(package: &'a PreservedPackage, path: &str) -> Option<&'a str> {
    effective_content_type(&package.content_types, path).map(|resolved| resolved.value)
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
            write_defined_name(w, defined)?;
        }
        Ok(())
    })?;
    Ok(())
}

fn write_defined_name(
    w: &mut Writer<Vec<u8>>,
    defined: &xlsx_model::DefinedName,
) -> io::Result<()> {
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
    let relationship_namespace = workbook_relationship_namespace(package)?;
    let mut sheets_attributes = package.workbook_sheets_attributes.clone();
    bind_relationship_prefix(
        &mut sheets_attributes,
        &package.workbook_template,
        &relationship_namespace,
    );
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
                REL_ID_ATTRIBUTE,
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
    let ambiguous =
        ambiguous_defined_names(&package.original_workbook.defined_names, &wb.defined_names);
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
        if defined == original && !ambiguous.contains(defined.name.as_str()) {
            replacements.push(source.bytes.clone());
            continue;
        }
        let patched = if ambiguous.contains(defined.name.as_str()) {
            written_defined_name(defined)?
        } else {
            patch_defined_name(&source.bytes, defined)?
        };
        replacements.push(template.qualify_fragment(&patched)?);
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

/// Several `_xlnm.Print_Area` entries with different scopes are normal, and the
/// model carries no source identity to tell them apart. Once their count
/// changes, no pairing is trustworthy, so those entries are rewritten from the
/// model rather than inheriting another entry's markup.
fn ambiguous_defined_names<'a>(
    original: &'a [xlsx_model::DefinedName],
    current: &[xlsx_model::DefinedName],
) -> HashSet<&'a str> {
    let mut counts: HashMap<&str, (usize, usize)> = HashMap::new();
    for defined in original {
        counts.entry(defined.name.as_str()).or_default().0 += 1;
    }
    for defined in current {
        if let Some(entry) = counts.get_mut(defined.name.as_str()) {
            entry.1 += 1;
        }
    }
    counts
        .into_iter()
        .filter(|(_, (before, after))| *before > 1 && before != after)
        .map(|(name, _)| name)
        .collect()
}

fn written_defined_name(defined: &xlsx_model::DefinedName) -> Result<Vec<u8>, ParseError> {
    fragment(|writer| write_defined_name(writer, defined))
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
    template: &XmlTemplate,
    plan: &SharedStringPlan,
) -> Result<Vec<u8>, ParseError> {
    let source_items = template.children_named("si").collect::<Vec<_>>();
    let mut items = Vec::with_capacity(wb.shared_strings.len());
    for (value, source) in wb.shared_strings.iter().zip(&plan.retained) {
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

struct SharedStringPlan {
    retained: Vec<Option<usize>>,
    source_to_output: HashMap<usize, usize>,
    generated: HashMap<String, usize>,
}

impl SharedStringPlan {
    fn new(
        original: &[String],
        source_count: usize,
        wb: &Workbook,
        shared_string_cells: &[SharedStringCells],
    ) -> Result<Self, ParseError> {
        let values = &wb.shared_strings;
        let mut sources: HashMap<&str, Vec<usize>> = HashMap::new();
        let mut output_counts: HashMap<&str, usize> = HashMap::new();
        for (index, value) in original.iter().enumerate().take(source_count) {
            sources.entry(value.as_str()).or_default().push(index);
        }
        for value in values {
            *output_counts.entry(value.as_str()).or_default() += 1;
        }
        let mut required = HashSet::new();
        for (sheet, cells) in wb.sheets.iter().zip(shared_string_cells) {
            for (&(row, col), &source) in cells {
                let Some(value) = original.get(source).filter(|_| source < source_count) else {
                    continue;
                };
                let matches = sheet
                    .cell(CellRef::new(row, col))
                    .filter(|cell| cell.formula.is_none())
                    .is_some_and(|cell| {
                        matches!(&cell.value, CellValue::Text { value: current } if current == value)
                    });
                if matches {
                    required.insert(source);
                }
            }
        }
        let mut selected: HashMap<&str, VecDeque<usize>> = HashMap::new();
        let mut generated_remaining = HashMap::new();
        for (&value, &count) in &output_counts {
            let candidates = sources.get(value).map(Vec::as_slice).unwrap_or_default();
            let retained_count = count.min(candidates.len());
            let required_for_value = candidates
                .iter()
                .copied()
                .filter(|source| required.contains(source))
                .collect::<Vec<_>>();
            if required_for_value.len() > retained_count {
                return Err(ParseError::UnsupportedEdit(
                    "shared string duplicates cannot retain every authored entry".to_owned(),
                ));
            }
            let mut retained_sources = required_for_value.into_iter().collect::<HashSet<_>>();
            for source in candidates {
                if retained_sources.len() == retained_count {
                    break;
                }
                retained_sources.insert(*source);
            }
            let retained = candidates
                .iter()
                .copied()
                .filter(|source| retained_sources.contains(source))
                .collect::<Vec<_>>();
            generated_remaining.insert(value, count - retained.len());
            selected.insert(value, retained.into());
        }
        let retained = values
            .iter()
            .map(|value| {
                let generated = generated_remaining
                    .get_mut(value.as_str())
                    .filter(|remaining| **remaining > 0);
                if let Some(remaining) = generated {
                    *remaining -= 1;
                    None
                } else {
                    selected
                        .get_mut(value.as_str())
                        .and_then(VecDeque::pop_front)
                }
            })
            .collect::<Vec<_>>();
        let mut source_to_output = HashMap::new();
        let mut generated = HashMap::new();
        for (output, (value, source)) in values.iter().zip(&retained).enumerate() {
            match source {
                Some(source) => {
                    source_to_output.insert(*source, output);
                }
                None => {
                    generated.entry(value.clone()).or_insert(output);
                }
            }
        }
        Ok(Self {
            retained,
            source_to_output,
            generated,
        })
    }

    fn index_for(&self, value: &str, source: Option<usize>, values: &[String]) -> Option<usize> {
        source
            .and_then(|source| self.source_to_output.get(&source).copied())
            .filter(|output| values.get(*output).is_some_and(|stored| stored == value))
            .or_else(|| self.generated.get(value).copied())
    }
}

fn write_shared_string_item(writer: &mut Writer<Vec<u8>>, value: &str) -> io::Result<()> {
    writer.create_element("si").write_inner_content(|writer| {
        write_text_el(writer, value)?;
        Ok(())
    })?;
    Ok(())
}

fn worksheet_xml_with_namespace(
    sheet: &Sheet,
    wb: &Workbook,
    main_namespace: &str,
    relationship_namespace: &str,
    links: &HyperlinkPlan,
    shared_string_plan: Option<&SharedStringPlan>,
) -> Result<Vec<u8>, ParseError> {
    doc(|writer| {
        let mut root = BytesStart::new("worksheet");
        root.push_attribute(("xmlns", main_namespace));
        if links.relationships.iter().any(Relationship::is_hyperlink) {
            root.push_attribute((REL_PREFIX_DECLARATION, relationship_namespace));
        }
        writer.write_event(Event::Start(root))?;
        write_sheet_views(writer, sheet)?;
        write_cols(writer, sheet)?;
        write_sheet_data(
            writer,
            sheet,
            wb,
            &SharedStringCells::new(),
            shared_string_plan,
        )?;
        write_merges(writer, sheet)?;
        write_hyperlinks(writer, sheet, &links.ids, REL_ID_ATTRIBUTE, &[])?;
        writer.write_event(Event::End(BytesEnd::new("worksheet")))?;
        Ok(())
    })
}

/// A retained worksheet and the relationship part backing it, when the model's
/// hyperlinks moved and the part has to be rewritten.
struct WorksheetOutput {
    bytes: Vec<u8>,
    relationships: Option<Vec<Relationship>>,
}

/// Preserved fragments (filters, validations, anchors) keep their source
/// geometry: valid XML, but stale after a row or column edit.
fn worksheet_xml_with_template(
    sheet: &Sheet,
    wb: &Workbook,
    original: Option<&Sheet>,
    source: &PreservedSheet,
    package: &PreservedPackage,
    shared_string_cells: &SharedStringCells,
    shared_string_plan: Option<&SharedStringPlan>,
) -> Result<WorksheetOutput, ParseError> {
    let template = &source.template;
    let columns = (!sheet.col_widths.is_empty())
        .then(|| fragment(|writer| write_cols(writer, sheet)))
        .transpose()?;
    let sheet_data = Some(fragment(|writer| {
        write_sheet_data(writer, sheet, wb, shared_string_cells, shared_string_plan)
    })?);
    let merges = (!sheet.merges.is_empty())
        .then(|| fragment(|writer| write_merges(writer, sheet)))
        .transpose()?;
    let mut replacements = vec![
        ("cols", columns),
        ("sheetData", sheet_data),
        ("mergeCells", merges),
    ];

    if original.is_none_or(|original| original.freeze_pane != sheet.freeze_pane) {
        replacements.push(("sheetViews", patched_sheet_views(template, sheet)?));
    }

    let mut relationships = None;
    if original.is_none_or(|original| original.hyperlinks != sheet.hyperlinks) {
        let source_relationships = package
            .part_bytes(&relationship_part_path(&source.path))
            .map(parse_relationships)
            .transpose()?
            .unwrap_or_default();
        let links = HyperlinkPlan::new(
            sheet,
            &source_relationships,
            &relationship_type_from(&package.workbook_relationships, "hyperlink", REL_HYPERLINK),
        );
        replacements.push((
            "hyperlinks",
            patched_hyperlinks(template, package, sheet, &links)?,
        ));
        relationships = Some(links.relationships);
    }

    Ok(WorksheetOutput {
        bytes: template.render(replacements, worksheet_child_rank)?,
        relationships,
    })
}

/// Rewrites each `<sheetView>`'s pane while its zoom, selection and view flags
/// stay on the element they were authored on.
fn patched_sheet_views(
    template: &XmlTemplate,
    sheet: &Sheet,
) -> Result<Option<Vec<u8>>, ParseError> {
    let Some(child) = template.child("sheetViews") else {
        return sheet
            .freeze_pane
            .map(|_| fragment(|writer| write_sheet_views(writer, sheet)))
            .transpose();
    };
    let views = XmlTemplate::capture(&child.bytes)?;
    let sources = views.children_named("sheetView").collect::<Vec<_>>();
    if sources.is_empty() {
        return sheet
            .freeze_pane
            .map(|_| fragment(|writer| write_sheet_views(writer, sheet)))
            .transpose();
    }
    let pane = sheet
        .freeze_pane
        .map(|pane| fragment(|writer| write_pane(writer, pane)))
        .transpose()?;
    let mut items = Vec::with_capacity(sources.len());
    for view in sources {
        let view = XmlTemplate::capture(&view.bytes)?;
        items.push(view.render(vec![("pane", pane.clone())], sheet_view_child_rank)?);
    }
    views.render_repeated("sheetView", &items, &[]).map(Some)
}

fn sheet_view_child_rank(name: &str) -> usize {
    match name {
        "pane" => 0,
        "selection" => 1,
        "pivotSelection" => 2,
        "extLst" => 3,
        _ => usize::MAX,
    }
}

fn patched_hyperlinks(
    template: &XmlTemplate,
    package: &PreservedPackage,
    sheet: &Sheet,
    links: &HyperlinkPlan,
) -> Result<Option<Vec<u8>>, ParseError> {
    if links.ids.is_empty() {
        return Ok(None);
    }
    let needs_binding = links.relationships.iter().any(Relationship::is_hyperlink);
    let namespace = fragment_relationship_namespace(template, package)?;
    let declared = template.declares_namespace(GENERATED_REL_PREFIX, &namespace);
    let root_attributes = if needs_binding && !declared {
        vec![(REL_PREFIX_DECLARATION.to_owned(), namespace)]
    } else {
        Vec::new()
    };
    fragment(|writer| {
        write_hyperlinks(
            writer,
            sheet,
            &links.ids,
            REL_ID_ATTRIBUTE,
            &root_attributes,
        )
    })
    .map(Some)
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

fn write_sheet_data(
    writer: &mut Writer<Vec<u8>>,
    sheet: &Sheet,
    wb: &Workbook,
    retained: &SharedStringCells,
    shared_string_plan: Option<&SharedStringPlan>,
) -> io::Result<()> {
    let mut sst_index: HashMap<&str, usize> = HashMap::with_capacity(wb.shared_strings.len());
    if shared_string_plan.is_none() {
        for (index, value) in wb.shared_strings.iter().enumerate() {
            sst_index.entry(value.as_str()).or_insert(index);
        }
    }

    let mut rows: Vec<RowId> = sheet.iter_cells().map(|(r, _)| r.row).collect();
    rows.extend(sheet.row_heights.keys().copied());
    rows.sort_unstable();
    rows.dedup();

    writer
        .create_element("sheetData")
        .write_inner_content(|writer| {
            for &row in &rows {
                write_row(
                    writer,
                    sheet,
                    row,
                    wb,
                    &sst_index,
                    retained,
                    shared_string_plan,
                )?;
            }
            Ok(())
        })?;
    Ok(())
}

fn write_hyperlinks(
    w: &mut Writer<Vec<u8>>,
    sheet: &Sheet,
    ids: &[Option<String>],
    id_name: &str,
    root_attributes: &[(String, String)],
) -> io::Result<()> {
    if sheet.hyperlinks.is_empty() {
        return Ok(());
    }
    let mut root = BytesStart::new("hyperlinks");
    for (name, value) in root_attributes {
        root.push_attribute((name.as_str(), value.as_str()));
    }
    w.write_event(Event::Start(root))?;
    for (link, relationship_id) in sheet.hyperlinks.iter().zip(ids) {
        let mut element = BytesStart::new("hyperlink");
        let reference = link.range.to_a1();
        element.push_attribute(("ref", reference.as_str()));
        if let Some(relationship_id) = relationship_id {
            element.push_attribute((id_name, relationship_id.as_str()));
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
    w.write_event(Event::End(BytesEnd::new("hyperlinks")))?;
    Ok(())
}

fn write_sheet_views(w: &mut Writer<Vec<u8>>, sheet: &Sheet) -> io::Result<()> {
    let Some(pane) = sheet.freeze_pane else {
        return Ok(());
    };
    w.create_element("sheetViews").write_inner_content(|w| {
        w.create_element("sheetView")
            .with_attribute(("workbookViewId", "0"))
            .write_inner_content(|w| write_pane(w, pane))?;
        Ok(())
    })?;
    Ok(())
}

fn write_pane(w: &mut Writer<Vec<u8>>, pane: xlsx_model::FreezePane) -> io::Result<()> {
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
            let width = fmt_num(width);
            let mut element = BytesStart::new("col");
            element.push_attribute(("min", n.as_str()));
            element.push_attribute(("max", n.as_str()));
            element.push_attribute(("width", width.as_str()));
            element.push_attribute(("customWidth", "1"));
            if width == "0" {
                element.push_attribute(("hidden", "1"));
            }
            w.write_event(Event::Empty(element))?;
        }
        Ok(())
    })?;
    Ok(())
}

fn write_row(
    w: &mut Writer<Vec<u8>>,
    sheet: &Sheet,
    row: RowId,
    wb: &Workbook,
    sst_index: &HashMap<&str, usize>,
    retained: &SharedStringCells,
    shared_string_plan: Option<&SharedStringPlan>,
) -> io::Result<()> {
    let r = (row as u64 + 1).to_string();
    let mut start = BytesStart::new("row");
    start.push_attribute(("r", r.as_str()));
    let ht = sheet.row_heights.get(&row).map(|h| fmt_num(*h));
    if let Some(h) = &ht {
        start.push_attribute(("ht", h.as_str()));
        start.push_attribute(("customHeight", "1"));
        if h == "0" {
            start.push_attribute(("hidden", "1"));
        }
    }
    w.write_event(Event::Start(start))?;
    for (addr, cell) in sheet.iter_cells().filter(|(a, _)| a.row == row) {
        let source = retained.get(&(addr.row, addr.col)).copied();
        let retained = match (&cell.value, shared_string_plan) {
            (CellValue::Text { value }, Some(plan)) => {
                plan.index_for(value, source, &wb.shared_strings)
            }
            (CellValue::Text { value }, None) => source
                .filter(|index| wb.shared_strings.get(*index) == Some(value))
                .or_else(|| sst_index.get(value.as_str()).copied()),
            _ => None,
        };
        write_cell(w, addr, cell, sst_index, retained)?;
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
    retained: Option<usize>,
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
            } else if let Some(idx) = retained.or_else(|| sst_index.get(s.as_str()).copied()) {
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

/// Reuses each source pool entry the model left alone, so font schemes,
/// `gray125` fills and `xf` attributes the model does not carry survive an
/// edit that only appends new entries.
fn patched_pool<T: PartialEq>(
    template: &XmlTemplate,
    pool_name: &str,
    item_name: &str,
    values: &[T],
    original: &[T],
    write_item: impl Fn(&mut Writer<Vec<u8>>, &T) -> io::Result<()>,
    write_pool: impl FnOnce() -> Result<Vec<u8>, ParseError>,
) -> Result<Option<Vec<u8>>, ParseError> {
    if values.is_empty() {
        return Ok(None);
    }
    let Some(child) = template.child(pool_name) else {
        return write_pool().map(Some);
    };
    let pool = XmlTemplate::capture(&child.bytes)?;
    let sources = pool.children_named(item_name).collect::<Vec<_>>();
    let mut items = Vec::with_capacity(values.len());
    for (index, value) in values.iter().enumerate() {
        match sources
            .get(index)
            .filter(|_| original.get(index) == Some(value))
        {
            Some(source) => items.push(source.bytes.clone()),
            None => {
                let item = fragment(|writer| write_item(writer, value))?;
                items.push(pool.qualify_fragment(&item)?);
            }
        }
    }
    pool.render_repeated(item_name, &items, &[("count", values.len().to_string())])
        .map(Some)
}

fn styles_xml_with_template(
    stylesheet: &Stylesheet,
    original: &Stylesheet,
    template: &XmlTemplate,
) -> Result<Vec<u8>, ParseError> {
    let num_fmts = patched_pool(
        template,
        "numFmts",
        "numFmt",
        &stylesheet.num_fmts,
        &original.num_fmts,
        |writer, (id, code)| {
            writer
                .create_element("numFmt")
                .with_attribute(("numFmtId", id.to_string().as_str()))
                .with_attribute(("formatCode", code.as_str()))
                .write_empty()?;
            Ok(())
        },
        || fragment(|writer| write_num_fmts(writer, stylesheet)),
    )?;
    let fonts = patched_pool(
        template,
        "fonts",
        "font",
        &stylesheet.fonts,
        &original.fonts,
        write_font,
        || fragment(|writer| write_fonts(writer, stylesheet)),
    )?;
    let fills = patched_pool(
        template,
        "fills",
        "fill",
        &stylesheet.fills,
        &original.fills,
        write_fill,
        || fragment(|writer| write_fills(writer, stylesheet)),
    )?;
    let borders = patched_pool(
        template,
        "borders",
        "border",
        &stylesheet.borders,
        &original.borders,
        write_border,
        || fragment(|writer| write_borders(writer, stylesheet)),
    )?;
    let cell_xfs = patched_pool(
        template,
        "cellXfs",
        "xf",
        &stylesheet.cell_xfs,
        &original.cell_xfs,
        write_xf,
        || fragment(|writer| write_cell_xfs(writer, stylesheet)),
    )?;
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
fn theme_xml(ss: &Stylesheet, namespace: &str) -> Result<Vec<u8>, ParseError> {
    let c = &ss.theme.colors;
    let slots = [
        "dk1", "lt1", "dk2", "lt2", "accent1", "accent2", "accent3", "accent4", "accent5",
        "accent6", "hlink", "folHlink",
    ];
    doc(|w| {
        w.create_element("a:theme")
            .with_attribute(("xmlns:a", namespace))
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
