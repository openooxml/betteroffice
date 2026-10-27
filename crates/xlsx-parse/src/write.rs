//! `xlsx_model::Workbook` -> xlsx parts. Without a source package this writes a
//! minimal valid workbook. With one, parts and sheets the model did not change
//! are copied through byte for byte; only what changed is reserialized, and
//! that reserialization carries just the subset `read` models.

use std::borrow::Cow;
use std::collections::{BTreeMap, HashMap, HashSet, VecDeque, btree_map};
use std::io;
use std::iter::Peekable;
use std::ops::Range;

use quick_xml::events::{BytesDecl, BytesEnd, BytesStart, BytesText, Event};
use quick_xml::{Reader, Writer};
use xlsx_model::addr::{ColId, RowId};
use xlsx_model::styles::{Alignment, Border, BorderEdge, Color, Fill, Font, Stylesheet, Xf};
use xlsx_model::{
    Cell, CellRange, CellRef, CellValue, ChartAnchor, ChartRef, DateSystem, DifferentialFormat,
    Sheet, SheetChart, SheetId, Workbook,
};

use crate::ParseError;
use crate::axis::SheetAxes;
use crate::package::{
    ContentTypeEntry, PartReference, PreservedPackage, PreservedSheet, Relationship,
    SourceSheetKind, XmlAttribute, XmlTemplate, attributes_from_fragment, effective_content_type,
    normalized_part_name, parse_relationships, relationship_part_path, relationships_by_id,
    remove_attribute, set_attribute,
};
use crate::patch::{
    SheetPatch, SourceStyles, StyleMatch, keeps_source_conditional_formats, keeps_source_xf,
    retains_source_styles, sheet_data_regeneration_loss,
};
use crate::read::SharedStringCells;
use crate::xml::{resolve_part_path, xml_err};

mod app_titles;
mod source_coordinates;
mod style_derivation;
mod views;

/// A saved workbook's parts: generated entries are owned, entries the source
/// package already held are borrowed from it.
#[doc(hidden)]
pub type SerializedParts<'p> = Vec<(String, Cow<'p, [u8]>)>;

pub(crate) const NS_MAIN: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
pub(crate) const NS_STRICT_MAIN: &str = "http://purl.oclc.org/ooxml/spreadsheetml/main";
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

struct CfPlan {
    authored: Vec<DifferentialFormat>,
    assigned: HashMap<(usize, usize, usize), u32>,
}

impl CfPlan {
    fn new(wb: &Workbook) -> Self {
        let mut plan = Self {
            authored: Vec::new(),
            assigned: HashMap::new(),
        };
        let start = wb.styles.dxfs.len() as u32;
        for (sheet_index, sheet) in wb.sheets.iter().enumerate() {
            for (group_index, group) in sheet.conditional_formats.iter().enumerate() {
                for (rule_index, rule) in group.rules.iter().enumerate() {
                    if rule.raw_xml.is_none() {
                        let id = start + plan.authored.len() as u32;
                        plan.assigned
                            .insert((sheet_index, group_index, rule_index), id);
                        plan.authored.push(rule.format.clone());
                    }
                }
            }
        }
        plan
    }

    fn dxf_id(
        &self,
        sheet: usize,
        group: usize,
        rule: usize,
        fallback: Option<u32>,
    ) -> Option<u32> {
        self.assigned
            .get(&(sheet, group, rule))
            .copied()
            .or(fallback)
    }
}

fn escape_xml_text(text: &str) -> String {
    text.replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('>', "&gt;")
}

fn escape_xml_attr(text: &str) -> String {
    escape_xml_text(text).replace('"', "&quot;")
}

fn argb(color: &str) -> String {
    format!("FF{}", color.trim_start_matches('#').to_ascii_uppercase())
}

fn dxf_bytes(format: &DifferentialFormat) -> Vec<u8> {
    let mut out = String::from("<dxf>");
    if format.font_color.is_some() || format.bold.is_some() || format.italic.is_some() {
        out.push_str("<font>");
        match format.bold {
            Some(true) => out.push_str("<b/>"),
            Some(false) => out.push_str("<b val=\"0\"/>"),
            None => {}
        }
        match format.italic {
            Some(true) => out.push_str("<i/>"),
            Some(false) => out.push_str("<i val=\"0\"/>"),
            None => {}
        }
        if let Some(color) = &format.font_color {
            out.push_str(&format!("<color rgb=\"{}\"/>", argb(color)));
        }
        out.push_str("</font>");
    }
    if let Some(fill) = &format.fill {
        out.push_str(&format!(
            "<fill><patternFill patternType=\"solid\"><fgColor rgb=\"{}\"/></patternFill></fill>",
            argb(fill)
        ));
    }
    out.push_str("</dxf>");
    if out == "<dxf></dxf>" {
        return b"<dxf/>".to_vec();
    }
    out.into_bytes()
}

fn conditional_formatting_bytes(
    sheet: &Sheet,
    sheet_index: usize,
    plan: &CfPlan,
    source: Option<(&XmlTemplate, &Sheet)>,
) -> Result<Option<Vec<u8>>, ParseError> {
    let mut source_indices = sheet
        .conditional_formats
        .iter()
        .map(|group| {
            source.and_then(|(_, original)| {
                original.conditional_formats.iter().position(|original| {
                    group.rules.iter().any(|rule| {
                        rule.raw_xml.is_some() && original.rules.iter().any(|old| old.id == rule.id)
                    })
                })
            })
        })
        .collect::<Vec<_>>();
    let mut used = source_indices
        .iter()
        .flatten()
        .copied()
        .collect::<HashSet<_>>();
    if let Some((_, original)) = source {
        for (group, index) in sheet.conditional_formats.iter().zip(&mut source_indices) {
            if index.is_none() {
                *index = original
                    .conditional_formats
                    .iter()
                    .enumerate()
                    .find(|(index, original)| {
                        !used.contains(index) && original.sqref == group.sqref
                    })
                    .map(|(index, _)| index);
                if let Some(index) = index {
                    used.insert(*index);
                }
            }
        }
    }
    let mut out = Vec::new();
    for (group_index, group) in sheet.conditional_formats.iter().enumerate() {
        if group.rules.is_empty() {
            continue;
        }
        let wrapper = source
            .and_then(|(template, _)| {
                source_indices[group_index]
                    .and_then(|index| template.children_named("conditionalFormatting").nth(index))
            })
            .map(|child| XmlTemplate::capture(&child.bytes))
            .transpose()?;
        let mut rules = Vec::new();
        for (rule_index, rule) in group.rules.iter().enumerate() {
            if let Some(raw) = &rule.raw_xml {
                rules.push(raw.as_bytes().to_vec());
                continue;
            }
            let mut rule_xml = format!("<cfRule type=\"{}\"", escape_xml_attr(&rule.rule_type));
            if let Some(dxf) = plan.dxf_id(sheet_index, group_index, rule_index, rule.dxf_id) {
                rule_xml.push_str(&format!(" dxfId=\"{dxf}\""));
            }
            rule_xml.push_str(&format!(" priority=\"{}\"", rule.priority));
            if let Some(operator) = &rule.operator {
                rule_xml.push_str(&format!(" operator=\"{}\"", escape_xml_attr(operator)));
            }
            if rule.formulas.is_empty() {
                rule_xml.push_str("/>");
            } else {
                rule_xml.push('>');
                for formula in &rule.formulas {
                    rule_xml.push_str(&format!("<formula>{}</formula>", escape_xml_text(formula)));
                }
                rule_xml.push_str("</cfRule>");
            }
            rules.push(match &wrapper {
                Some(wrapper) => wrapper.qualify_fragment(rule_xml.as_bytes())?,
                None => rule_xml.into_bytes(),
            });
        }
        if let Some(wrapper) = wrapper {
            wrapper.require_unwrapped_child("cfRule")?;
            out.extend(wrapper.render_repeated("cfRule", &rules, &[])?);
        } else {
            out.extend_from_slice(
                format!(
                    "<conditionalFormatting sqref=\"{}\">",
                    escape_xml_attr(&group.sqref)
                )
                .as_bytes(),
            );
            for rule in rules {
                out.extend(rule);
            }
            out.extend_from_slice(b"</conditionalFormatting>");
        }
    }
    Ok((!out.is_empty()).then_some(out))
}

fn cf_only_delta(wb: &Workbook, original: &Workbook) -> bool {
    if wb.sheets.len() != original.sheets.len() {
        return false;
    }
    let strip = |book: &Workbook| {
        let mut book = book.clone();
        for sheet in &mut book.sheets {
            sheet.conditional_formats = Vec::new();
        }
        book
    };
    strip(wb) == strip(original)
}

/// serialize a workbook to opc parts in a fixed, deterministic order. a chart
/// is refused: this writer emits no drawing, chart relationship or chart part,
/// so it would drop one rather than write it.
pub fn serialize_workbook(wb: &Workbook) -> Result<Vec<(String, Vec<u8>)>, ParseError> {
    serialize_workbook_with_active_sheet(wb, SheetId(0))
}

#[doc(hidden)]
pub fn serialize_workbook_with_active_sheet(
    wb: &Workbook,
    active_sheet: SheetId,
) -> Result<Vec<(String, Vec<u8>)>, ParseError> {
    validate_indexed_colors(&wb.styles)?;
    if wb.sheets.iter().any(|sheet| !sheet.charts.is_empty()) {
        return Err(ParseError::UnsupportedEdit(
            "a chart can only be written back into the package it was read from".to_owned(),
        ));
    }
    let cf_plan = CfPlan::new(wb);
    let have_sst = !wb.shared_strings.is_empty();
    let have_styles = !wb.styles.is_empty() || !cf_plan.authored.is_empty();
    let mut parts = vec![
        (
            "[Content_Types].xml".to_string(),
            content_types(wb, have_sst, have_styles)?,
        ),
        ("_rels/.rels".to_string(), root_rels()?),
        (
            "xl/workbook.xml".to_string(),
            workbook_xml(wb, active_sheet)?,
        ),
        (
            "xl/_rels/workbook.xml.rels".to_string(),
            workbook_rels(wb, have_sst, have_styles)?,
        ),
    ];
    if have_sst {
        parts.push(("xl/sharedStrings.xml".to_string(), shared_strings_xml(wb)?));
    }
    if have_styles {
        parts.push((
            "xl/styles.xml".to_string(),
            styles_xml_with_authored(&wb.styles, &cf_plan.authored)?,
        ));
        parts.push((
            "xl/theme/theme1.xml".to_string(),
            theme_xml(&wb.styles, NS_DML)?,
        ));
    }
    for (i, sheet) in wb.sheets.iter().enumerate() {
        let links = HyperlinkPlan::new(sheet, &[], REL_HYPERLINK);
        parts.push((
            format!("xl/worksheets/sheet{}.xml", i + 1),
            worksheet_xml_with_namespace(sheet, wb, NS_MAIN, NS_R, &links, None, i, &cf_plan)?,
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
    let parts = serialize_workbook_with_package_and_origins_after_edits(
        wb,
        package,
        origins,
        &shared_string_cells,
        SaveEdits {
            changed,
            moved_references: false,
        },
    )?;
    Ok(parts
        .into_iter()
        .map(|(path, bytes)| (path, bytes.into_owned()))
        .collect())
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
    /// an edit moved cells a preserved part names but this crate cannot
    /// rewrite. The caller resolves that against
    /// [`PreservedPackage::unpatchable_references`], because only it knows
    /// which edits were applied.
    pub moved_references: bool,
}

/// `origins` and `shared_string_cells` are indexed by current sheet and carry
/// what the model does not: the source sheet each one came from and the shared
/// string entry each of its cells was authored against.
#[doc(hidden)]
pub fn serialize_workbook_with_package_and_origins_after_edits<'p>(
    wb: &Workbook,
    package: &'p PreservedPackage,
    origins: &[Option<usize>],
    shared_string_cells: &[SharedStringCells],
    edits: SaveEdits,
) -> Result<SerializedParts<'p>, ParseError> {
    let axes = vec![Some(SheetAxes::default()); wb.sheets.len()];
    serialize_workbook_with_package_and_origins_after_edits_and_active_sheet_with_axes(
        wb,
        package,
        origins,
        shared_string_cells,
        &axes,
        edits,
        package.active_sheet,
    )
}

#[doc(hidden)]
pub fn serialize_workbook_with_package_and_origins_after_edits_and_active_sheet<'p>(
    wb: &Workbook,
    package: &'p PreservedPackage,
    origins: &[Option<usize>],
    shared_string_cells: &[SharedStringCells],
    edits: SaveEdits,
    active_sheet: SheetId,
) -> Result<SerializedParts<'p>, ParseError> {
    let axes = vec![Some(SheetAxes::default()); wb.sheets.len()];
    serialize_workbook_with_package_and_origins_after_edits_and_active_sheet_with_axes(
        wb,
        package,
        origins,
        shared_string_cells,
        &axes,
        edits,
        active_sheet,
    )
}

/// `sheet_axes` carries, per current sheet, where its source rows and columns
/// sit after the row and column edits made since the package was read. Sheets
/// without one (`None`) are reserialized from the model, as are sheets whose
/// source cannot be patched cell by cell.
///
/// Parts the package already holds come back borrowed, so a save writes them
/// straight out of the retained copy instead of cloning each one.
#[doc(hidden)]
pub fn serialize_workbook_with_package_and_origins_after_edits_and_active_sheet_with_axes<'p>(
    wb: &Workbook,
    package: &'p PreservedPackage,
    origins: &[Option<usize>],
    shared_string_cells: &[SharedStringCells],
    sheet_axes: &[Option<SheetAxes>],
    edits: SaveEdits,
    active_sheet: SheetId,
) -> Result<SerializedParts<'p>, ParseError> {
    validate_indexed_colors(&wb.styles)?;
    if origins.len() != wb.sheets.len() || shared_string_cells.len() != wb.sheets.len() {
        return Err(ParseError::Malformed(
            "sheet origin count does not match workbook".to_owned(),
        ));
    }
    let edited = edits.changed;
    let claimed = claimed_origins(package, origins);
    let view_plan = views::ViewPlan::new(package, &claimed, active_sheet)?;
    if !edited
        && origins.len() == package.sheets.len()
        && origins
            .iter()
            .enumerate()
            .all(|(index, origin)| *origin == Some(index))
        && active_sheet == package.active_sheet
    {
        return Ok(package
            .parts
            .iter()
            .map(|(path, bytes)| (path.clone(), Cow::Borrowed(bytes.as_slice())))
            .collect());
    }
    preflight_preserved_save(wb, package, origins, edits)?;

    let mut cf_plan = CfPlan::new(wb);
    if !cf_plan.authored.is_empty()
        && let Some(template) = &package.stylesheet_template
    {
        let count = source_dxf_count(template)? as u32;
        for id in cf_plan.assigned.values_mut() {
            *id = count + (*id - wb.styles.dxfs.len() as u32);
        }
    }
    let cf_only = edited && cf_only_delta(wb, &package.original_workbook);
    let edited = edited && !cf_only;
    let have_sst = !wb.shared_strings.is_empty();
    let have_styles =
        !wb.styles.is_empty() || package.styles.is_some() || !cf_plan.authored.is_empty();
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
        &claimed,
        &mut used_relationship_ids,
        &mut used_paths,
        &worksheet_relationship_type,
    )?;
    let mut planned_ids = sheets
        .iter()
        .filter_map(|sheet| sheet.relationship.id().map(str::to_owned))
        .collect::<HashSet<_>>();
    let shared_strings = have_sst.then(|| {
        plan_part(
            package.shared_strings.as_ref(),
            &package.workbook_relationships,
            "sharedStrings",
            &shared_strings_relationship_type,
            "xl/sharedStrings.xml",
            &mut used_relationship_ids,
            &mut planned_ids,
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
            &mut planned_ids,
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
            &mut planned_ids,
        )
    });

    let mut parts = PartStore::new(&package.parts);
    if let Some(path) = app_titles::part_path(&package.root_relationships)
        && let Some(bytes) = parts.get(&path)
        && let Some(patched) = app_titles::patch(bytes, wb, &package.original_workbook, origins)
    {
        parts.set(path, patched)?;
    }
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
    let chart_refreshes = patch_chart_parts(wb, package, origins, &mut parts)?;

    let shared_strings_stable = wb.shared_strings == package.original_workbook.shared_strings;
    let style_plan = style_derivation::StylePlan::new(wb, package, origins, sheet_axes)?;
    let style_match = StyleMatch::new(&package.original_workbook.styles, &wb.styles)
        .with_derived(style_plan.pairs.clone());
    let empty_provenance = SharedStringCells::new();
    let axes_changed = sheet_axes.iter().flatten().any(|axes| !axes.is_identity());
    let defined_names = package
        .original_workbook
        .defined_names
        .iter()
        .chain(&wb.defined_names)
        .map(|defined| defined.name.as_str())
        .collect::<Vec<_>>();
    let sheet_names_stable = wb.sheets.len() == package.original_workbook.sheets.len()
        && origins.iter().flatten().collect::<HashSet<_>>().len() == wb.sheets.len()
        && wb
            .sheets
            .iter()
            .map(|sheet| sheet.name.to_ascii_lowercase())
            .collect::<HashSet<_>>()
            .len()
            == wb.sheets.len()
        && wb.sheets.iter().zip(origins).all(|(sheet, origin)| {
            origin
                .and_then(|origin| package.original_workbook.sheets.get(origin))
                .is_some_and(|original| original.name == sheet.name)
        });
    let reference_axes = wb
        .sheets
        .iter()
        .enumerate()
        .map(|(index, sheet)| {
            let axes = sheet_names_stable
                .then_some(index)
                .and_then(|index| sheet_axes.get(index))
                .and_then(Option::as_ref);
            (sheet.name.as_str(), axes)
        })
        .collect::<Vec<_>>();
    for (index, (sheet, plan)) in wb.sheets.iter().zip(&sheets).enumerate() {
        let source = plan.origin.and_then(|origin| package.sheets.get(origin));
        let original = plan
            .origin
            .and_then(|origin| package.original_workbook.sheets.get(origin));
        let provenance = shared_string_cells.get(index).unwrap_or(&empty_provenance);
        let axes = sheet_axes.get(index).and_then(|axes| axes.as_ref());
        let output = match source {
            Some(source) if source.is_worksheet() => {
                if shared_strings_stable
                    && original.is_some_and(|original| {
                        package.part_bytes(&source.path).is_some_and(|bytes| {
                            sheet_body_matches(
                                sheet,
                                original,
                                bytes,
                                axes,
                                &reference_axes,
                                axes_changed,
                                Some(&defined_names),
                                &style_match,
                            )
                        })
                    })
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
                    axes,
                    &style_match,
                    index,
                    &cf_plan,
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
                        index,
                        &cf_plan,
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

    if let Some(view_plan) = &view_plan {
        for (index, sheet) in sheets.iter().enumerate() {
            if let Some(bytes) = parts.get(&sheet.path)
                && let Some(patched) = view_plan.patch_sheet(bytes, index)?
            {
                parts.set(sheet.path.clone(), patched)?;
            }
        }
    }
    let workbook = workbook_xml_with_template(wb, package, &sheets, edited, view_plan.as_ref())?;
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
            &mut used_relationship_ids,
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
    ) && cf_plan.authored.is_empty()
        && retains_source_styles(&style_plan.styles, &package.original_workbook.styles);
    if !styles_retained {
        replace_optional_part(&mut parts, package.styles.as_ref(), styles.as_ref(), || {
            match &package.stylesheet_template {
                Some(template) => styles_xml_with_template(
                    &style_plan.styles,
                    &package.original_workbook.styles,
                    template,
                    &style_plan.overrides,
                    &cf_plan.authored,
                ),
                None => styles_xml_with_namespace(&wb.styles, main_namespace, &cf_plan.authored),
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

    for (path, bytes) in chart_refreshes {
        let _ = parts.set(path, bytes);
    }
    Ok(parts.finish())
}

/// What one sheet wants written into a chart part it anchors.
struct ChartClaim<'a> {
    chart: &'a SheetChart,
    refs: &'a [ChartRef],
    sheet: &'a str,
    original_refs: std::borrow::Cow<'a, [ChartRef]>,
    original_sheet: &'a str,
}

/// What one sheet wants written into a drawing anchor it holds.
struct AnchorClaim<'a> {
    anchor: ChartAnchor,
    sheet: &'a str,
    moved: bool,
}

/// Writes moved chart references and anchors back into the parts that hold
/// them. Each part is patched once, from its source bytes, so unmodelled chart
/// markup survives byte for byte. A part two sheets anchor is written only when
/// both want the same bytes out of it, because one part cannot hold two.
fn patch_chart_parts(
    wb: &Workbook,
    package: &PreservedPackage,
    origins: &[Option<usize>],
    parts: &mut PartStore<'_>,
) -> Result<Vec<(String, Vec<u8>)>, ParseError> {
    let mut chart_refs: BTreeMap<&str, Vec<ChartClaim<'_>>> = BTreeMap::new();
    let mut anchors: BTreeMap<&str, BTreeMap<usize, AnchorClaim<'_>>> = BTreeMap::new();
    for (sheet, origin) in wb.sheets.iter().zip(origins) {
        let original_sheet = origin.and_then(|origin| package.original_workbook.sheets.get(origin));
        let source = original_sheet
            .map(|sheet| sheet.charts.as_slice())
            .unwrap_or_default();
        for chart in &sheet.charts {
            let original = source_chart(source, chart, &sheet.name)?;
            if chart.raw_xml.is_none()
                && (original.refs.len() != chart.refs.len()
                    || original
                        .refs
                        .iter()
                        .zip(&chart.refs)
                        .any(|(original, edited)| original.kind != edited.kind))
            {
                return Err(ParseError::UnsupportedEdit(format!(
                    "chart {} no longer holds the references it was read from",
                    chart.part
                )));
            }
            chart_refs
                .entry(chart.part.as_str())
                .or_default()
                .push(ChartClaim {
                    chart,
                    refs: chart.refs.as_slice(),
                    sheet: sheet.name.as_str(),
                    original_refs: match &chart.raw_xml {
                        Some(xml) => {
                            std::borrow::Cow::Owned(crate::chart::chart_references(xml.as_bytes())?)
                        }
                        None => std::borrow::Cow::Borrowed(original.refs.as_slice()),
                    },
                    original_sheet: original_sheet.map_or(&sheet.name, |sheet| &sheet.name),
                });
            claim_anchor(
                anchors.entry(chart.drawing.as_str()).or_default(),
                chart,
                &sheet.name,
                original.anchor != chart.anchor,
            )?;
        }
    }
    let mut refreshes = Vec::new();
    for (path, claims) in chart_refs {
        let moved = claims.iter().any(|claim| {
            claim.refs != claim.original_refs.as_ref() || claim.chart.raw_xml.is_some()
        });
        let Some(source) = package.part_bytes(path) else {
            if moved {
                return Err(missing_part(path));
            }
            continue;
        };
        let rewritten = if moved {
            agreed_chart_bytes(source, path, &claims, wb, &package.original_workbook)?
                .unwrap_or_else(|| source.to_vec())
        } else {
            source.to_vec()
        };
        if let Some(bytes) =
            refreshed_chart_bytes(&rewritten, &claims, wb, &package.original_workbook)
            && bytes != rewritten
        {
            refreshes.push((path.to_owned(), bytes));
        }
        if rewritten != source {
            parts.set(path.to_owned(), rewritten)?;
        }
    }
    for (path, claims) in anchors {
        let moved = claims
            .into_iter()
            .filter(|(_, claim)| claim.moved)
            .map(|(index, claim)| (index, claim.anchor))
            .collect::<Vec<_>>();
        if moved.is_empty() {
            continue;
        }
        let source = package.part_bytes(path).ok_or_else(|| missing_part(path))?;
        parts.set(
            path.to_owned(),
            crate::chart::patch_drawing_anchors(source, &moved)?,
        )?;
    }
    Ok(refreshes)
}

/// Records what one sheet wants at a drawing anchor, refusing when a sheet
/// that shares the drawing already asked for a different one.
fn claim_anchor<'a>(
    claims: &mut BTreeMap<usize, AnchorClaim<'a>>,
    chart: &SheetChart,
    sheet: &'a str,
    moved: bool,
) -> Result<(), ParseError> {
    match claims.entry(chart.anchor_index) {
        btree_map::Entry::Vacant(slot) => {
            slot.insert(AnchorClaim {
                anchor: chart.anchor,
                sheet,
                moved,
            });
        }
        btree_map::Entry::Occupied(mut slot) => {
            if slot.get().anchor != chart.anchor {
                return Err(conflicting_shared_part(
                    &chart.drawing,
                    slot.get().sheet,
                    sheet,
                ));
            }
            slot.get_mut().moved |= moved;
        }
    }
    Ok(())
}

/// The bytes a chart part must take, refused when the sheets anchoring it do
/// not all patch it into the same thing. A shared part's references may be
/// unqualified, so the sheet a cache is rebuilt against is part of what the
/// sheets must agree on.
fn agreed_chart_bytes(
    source: &[u8],
    path: &str,
    claims: &[ChartClaim<'_>],
    wb: &Workbook,
    original: &Workbook,
) -> Result<Option<Vec<u8>>, ParseError> {
    let mut agreed: Option<(Vec<u8>, &str)> = None;
    for claim in claims {
        let base = match &claim.chart.raw_xml {
            Some(xml) => crate::tree::Part::decode(source)?.encode_text(xml),
            None => source.to_vec(),
        };
        let bytes = crate::chart::patch_chart_refs(
            &base,
            path,
            claim.refs,
            wb,
            claim.sheet,
            claim
                .chart
                .raw_xml
                .is_none()
                .then_some((original, claim.original_sheet)),
        )?;
        match &agreed {
            Some((first, owner)) if *first != bytes => {
                return Err(conflicting_shared_part(path, owner, claim.sheet));
            }
            Some(_) => {}
            None => agreed = Some((bytes, claim.sheet)),
        }
    }
    Ok(agreed.map(|(bytes, _)| bytes))
}

fn refreshed_chart_bytes(
    source: &[u8],
    claims: &[ChartClaim<'_>],
    wb: &Workbook,
    original: &Workbook,
) -> Option<Vec<u8>> {
    let plan = crate::chart::ChartRefreshPlan::new(source);
    let mut refresh = plan
        .save_refresh(wb, original)
        .edited(claims.iter().any(|claim| claim.chart.raw_xml.is_some()));
    let mut refreshed = None;
    let mut seen: BTreeMap<_, Vec<&[ChartRef]>> = BTreeMap::new();
    for claim in claims {
        let owner = if plan.owner_dependent() {
            claim.sheet
        } else {
            ""
        };
        let original_owner = if claim
            .original_refs
            .iter()
            .any(|reference| !reference.formula.contains('!'))
        {
            claim.original_sheet
        } else {
            ""
        };
        let references = seen.entry((owner, original_owner)).or_default();
        if references.contains(&claim.original_refs.as_ref()) {
            continue;
        }
        references.push(claim.original_refs.as_ref());
        let bytes = refresh
            .refresh_changed(
                claim.sheet,
                claim.original_sheet,
                claim.original_refs.as_ref(),
            )
            .unwrap_or_else(|| source.to_vec());
        match &refreshed {
            Some(first) if *first != bytes => return None,
            Some(_) => {}
            None => refreshed = Some(bytes),
        }
    }
    refreshed
}

/// Two sheets anchor one part and no longer agree on what it holds. The part
/// can carry only one of them, so the save is refused rather than rewriting the
/// chart the other sheet shows.
fn conflicting_shared_part(path: &str, first: &str, second: &str) -> ParseError {
    ParseError::UnsupportedEdit(format!(
        "{path} is anchored by both sheet {first} and sheet {second}, which no longer agree on what it holds, and one part cannot carry both"
    ))
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

fn unwritable_grid(sheet: &str, reason: impl std::fmt::Display) -> ParseError {
    ParseError::UnsupportedEdit(format!("sheet {sheet}: {reason}"))
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
/// never patches. Dropping, reordering or renaming a sheet one of them names —
/// or moving cells it names, which the caller reports — strands it, so the
/// serializer refuses that here instead of trusting every caller to have
/// checked. A sheet no preserved reference names moves freely.
fn ensure_preserved_references_stay_valid(
    wb: &Workbook,
    package: &PreservedPackage,
    origins: &[Option<usize>],
    edits: SaveEdits,
) -> Result<(), ParseError> {
    let refused = |part: &str| {
        ParseError::UnsupportedEdit(format!(
            "{part} references sheets or cells this save would move, and it cannot be rewritten"
        ))
    };
    if edits.moved_references
        && let Some(part) = package.unpatchable_reference_part()
    {
        return Err(refused(part));
    }
    for name in displaced_source_sheets(wb, package, origins) {
        if let Some(part) = package.reference_naming_sheet(&name) {
            return Err(refused(part));
        }
    }
    Ok(())
}

/// The names of the source sheets this save does not leave where and as they
/// were: one that is gone, one whose rank among the survivors changed, one
/// renamed.
fn displaced_source_sheets(
    wb: &Workbook,
    package: &PreservedPackage,
    origins: &[Option<usize>],
) -> Vec<String> {
    let retained = origins.iter().flatten().copied().collect::<Vec<_>>();
    let mut ranked = retained.clone();
    ranked.sort_unstable();
    let mut displaced = Vec::new();
    for (index, original) in package.original_workbook.sheets.iter().enumerate() {
        let rank = retained.iter().position(|origin| *origin == index);
        let kept_in_place = rank.is_some_and(|rank| ranked.get(rank) == Some(&index))
            && origins
                .iter()
                .zip(&wb.sheets)
                .any(|(origin, sheet)| *origin == Some(index) && sheet.name == original.name);
        if !kept_in_place {
            displaced.push(original.name.clone());
        }
    }
    displaced
}

/// This crate patches chart parts in place; it can neither create one nor
/// delete one. Every chart frame a retained sheet carries must therefore name
/// exactly one frame the same source sheet was read with, and every frame that
/// source held must still be there. Frames, not parts: two anchors may share
/// one chart part, and each of them is written back on its own.
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
            source_chart(source, chart, &sheet.name)?;
        }
        for original in source {
            match sheet
                .charts
                .iter()
                .filter(|chart| chart.is_same_frame(original))
                .count()
            {
                1 => {}
                0 => {
                    return Err(ParseError::UnsupportedEdit(format!(
                        "chart {} was dropped from sheet {}, and this crate cannot delete one",
                        original.part, sheet.name
                    )));
                }
                _ => {
                    return Err(ParseError::UnsupportedEdit(format!(
                        "sheet {} carries the anchor holding chart {} twice, and one anchor cannot hold two",
                        sheet.name, original.part
                    )));
                }
            }
        }
    }
    Ok(())
}

/// The source frame a modelled one must be patched from: same drawing anchor,
/// same chart part. A frame that repointed either has no source to patch, and
/// patching the one it used to hold would rewrite somebody else's chart.
fn source_chart<'a>(
    source: &'a [SheetChart],
    chart: &SheetChart,
    sheet: &str,
) -> Result<&'a SheetChart, ParseError> {
    let mut frames = source
        .iter()
        .filter(|other| other.is_same_frame(chart) && other.part == chart.part);
    match (frames.next(), frames.next()) {
        (Some(original), None) => Ok(original),
        (None, _) if source.iter().any(|other| other.part == chart.part) => {
            Err(ParseError::UnsupportedEdit(format!(
                "chart {} no longer names the drawing and anchor it was read from",
                chart.part
            )))
        }
        _ => Err(unwritable_chart(&chart.part, sheet)),
    }
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
#[allow(clippy::too_many_arguments)]
fn sheet_body_matches(
    sheet: &Sheet,
    original: &Sheet,
    source: &[u8],
    axes: Option<&SheetAxes>,
    reference_axes: &[(&str, Option<&SheetAxes>)],
    axes_changed: bool,
    defined_names: Option<&[&str]>,
    styles: &StyleMatch<'_>,
) -> bool {
    if !source_coordinates::unchanged_with_defined_names(
        source,
        axes,
        reference_axes,
        axes_changed,
        defined_names,
    ) {
        return false;
    }
    sheet.freeze_pane == original.freeze_pane
        && sheet.hyperlinks == original.hyperlinks
        && keeps_source_conditional_formats(
            &sheet.conditional_formats,
            &original.conditional_formats,
        )
        && sheet.auto_filter == original.auto_filter
        && sheet.merges == original.merges
        && sheet.col_widths == original.col_widths
        && sheet.row_heights == original.row_heights
        && sheet.col_styles == original.col_styles
        && sheet.row_styles == original.row_styles
        && sheet.array_formulas().eq(original.array_formulas())
        && if let Some(axes) = axes {
            let mut sources = original.iter_cells();
            sheet.iter_cells().all(|(at, cell)| {
                sources.next().is_some_and(|(source, original)| {
                    at == source
                        && axes.rows.current(source.row) == Some(source.row)
                        && axes.cols.current(source.col) == Some(source.col)
                        && styles.same_cell(original, cell)
                })
            }) && sources.next().is_none()
        } else {
            sheet.iter_cells().eq(original.iter_cells())
        }
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

/// The source sheet each sheet keeps: the first in-range claim of a source
/// index keeps it, later claims are new sheets.
fn claimed_origins(package: &PreservedPackage, origins: &[Option<usize>]) -> Vec<Option<usize>> {
    let mut claimed = HashSet::new();
    origins
        .iter()
        .map(|origin| {
            origin.filter(|origin| *origin < package.sheets.len() && claimed.insert(*origin))
        })
        .collect()
}

fn plan_sheets(
    package: &PreservedPackage,
    claimed_origins: &[Option<usize>],
    used_relationship_ids: &mut HashSet<String>,
    used_paths: &mut HashSet<String>,
    worksheet_relationship_type: &str,
) -> Result<Vec<PlannedSheet>, ParseError> {
    let mut sheet_ids = SheetIds::new(package);
    let relationship_by_id = relationships_by_id(&package.workbook_relationships);
    claimed_origins
        .iter()
        .map(|origin| {
            if let Some(origin) = *origin {
                let source = &package.sheets[origin];
                let source_relationship = source
                    .relationship_id
                    .as_deref()
                    .and_then(|id| relationship_by_id.get(id).copied());
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

/// Plans a part the reader finds by relationship type, so its relationship
/// takes a fresh id when a sheet or another part already holds the source one.
fn plan_part(
    source: Option<&PartReference>,
    relationships: &[Relationship],
    type_suffix: &str,
    relationship_type: &str,
    fallback_path: &str,
    used_relationship_ids: &mut HashSet<String>,
    planned_ids: &mut HashSet<String>,
) -> PlannedPart {
    let typed = |relationship: &&Relationship| {
        relationship.has_type(type_suffix) && relationship.target().is_some()
    };
    let source_relationship = source.and_then(|source| {
        source
            .relationship_id
            .as_deref()
            .and_then(|id| {
                relationships
                    .iter()
                    .filter(typed)
                    .find(|relationship| relationship.id() == Some(id))
            })
            .or_else(|| relationships.iter().find(typed))
    });
    let path = source.map_or_else(|| fallback_path.to_owned(), |source| source.path.clone());
    let mut relationship = match source_relationship {
        Some(relationship) => relationship.clone(),
        None => new_relationship(
            next_relationship_id(used_relationship_ids),
            relationship_type,
            relative_to_xl(&path),
        ),
    };
    if relationship
        .id()
        .is_none_or(|id| !planned_ids.insert(id.to_owned()))
    {
        let id = next_relationship_id(used_relationship_ids);
        planned_ids.insert(id.clone());
        set_attribute(&mut relationship.attributes, "Id", "Id", id);
    }
    PlannedPart { path, relationship }
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

    fn get(&self, path: &str) -> Option<&[u8]> {
        match &self.slots[*self.positions.get(&normalized_part_name(path))?] {
            Slot::Source(_, bytes) => Some(bytes),
            Slot::Owned(_, bytes) => Some(bytes),
            Slot::Removed => None,
        }
    }

    fn finish(self) -> Vec<(String, Cow<'a, [u8]>)> {
        self.slots
            .into_iter()
            .filter_map(|slot| match slot {
                Slot::Source(path, bytes) => Some((path.to_owned(), Cow::Borrowed(bytes))),
                Slot::Owned(path, bytes) => Some((path, Cow::Owned(bytes))),
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

pub(crate) fn fragment<F>(f: F) -> Result<Vec<u8>, ParseError>
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
    used_relationship_ids: &mut HashSet<String>,
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

    let resolved = relationships_by_id(&package.workbook_relationships);
    let mut emitted = HashSet::new();
    let mut merged = Vec::new();
    let mut renumbered = Vec::new();
    for source in &package.workbook_relationships {
        if edited && source.has_type("calcChain") {
            continue;
        }
        let id = source.id();
        if let Some(planned) = id.and_then(|id| planned_by_id.get(id)) {
            if emitted.insert(id.unwrap().to_owned()) {
                merged.push((*planned).clone());
            }
        } else if id.is_some_and(|id| source_owned_ids.contains(id))
            || source.has_type("sharedStrings")
            || source.has_type("styles")
            || source.has_type("theme")
            || source.target().is_none()
        {
        } else if let Some(id) = id
            && !resolved
                .get(id)
                .is_some_and(|resolved| std::ptr::eq(*resolved, source))
        {
            // a duplicate id the reader does not resolve keeps its part under a fresh id
            renumbered.push(source);
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
    for source in renumbered {
        let mut relationship = source.clone();
        set_attribute(
            &mut relationship.attributes,
            "Id",
            "Id",
            next_relationship_id(used_relationship_ids),
        );
        merged.push(relationship);
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

fn workbook_xml(wb: &Workbook, active_sheet: SheetId) -> Result<Vec<u8>, ParseError> {
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
                if active_sheet.0 != 0 {
                    let active_tab = active_sheet.0.to_string();
                    w.create_element("bookViews").write_inner_content(|w| {
                        w.create_element("workbookView")
                            .with_attribute(("activeTab", active_tab.as_str()))
                            .write_empty()?;
                        Ok(())
                    })?;
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
    view_plan: Option<&views::ViewPlan>,
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
    let calc_pr = if edited && holds_formulas(wb, package, sheets) {
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
    let mut template = package.workbook_template.clone();
    if wb.defined_names != package.original_workbook.defined_names {
        let origins = sheets.iter().map(|sheet| sheet.origin).collect::<Vec<_>>();
        let defined_names = edited_defined_names(wb, package, &origins)?;
        if let Some(index) = template
            .children
            .iter()
            .position(|child| template.is_root_child_named(child, "definedNames"))
            && let Some(bytes) = &defined_names
        {
            template.children[index].bytes = bytes.clone();
        } else {
            replacements.push(("definedNames", defined_names));
        }
    } else if let Some(child) = package.workbook_template.child("definedNames") {
        replacements.push(("definedNames", Some(child.bytes.clone())));
    }
    if let Some(view_plan) = view_plan {
        if let Some(index) = template
            .children
            .iter()
            .position(|child| template.is_root_child_named(child, "bookViews"))
        {
            template.children[index].bytes = view_plan.book_views.clone();
        } else {
            replacements.push(("bookViews", Some(view_plan.book_views.clone())));
        }
    }
    template.render(replacements, workbook_child_rank)
}

/// Whether the saved workbook holds a formula. Only worksheet cells are
/// modeled, and an `mc:AlternateContent` branch can hide one from the model,
/// so a macrosheet, an unknown sheet kind or such a source sheet counts too.
fn holds_formulas(wb: &Workbook, package: &PreservedPackage, sheets: &[PlannedSheet]) -> bool {
    sheets.iter().filter_map(|plan| plan.origin).any(|origin| {
        matches!(
            package.source_sheet_kind(origin),
            Some(SourceSheetKind::Macrosheet | SourceSheetKind::Other)
        ) || package
            .source_cell_facts(origin)
            .is_some_and(|facts| facts.alternate_formulas)
    }) || wb
        .sheets
        .iter()
        .any(|sheet| sheet.iter_cells().any(|(_, cell)| cell.formula.is_some()))
}

#[doc(hidden)]
pub fn validate_defined_name_edits(
    wb: &Workbook,
    package: &PreservedPackage,
    origins: &[Option<usize>],
) -> Result<(), ParseError> {
    if wb.defined_names != package.original_workbook.defined_names {
        edited_defined_names(wb, package, origins)?;
    }
    Ok(())
}

fn edited_defined_names(
    wb: &Workbook,
    package: &PreservedPackage,
    origins: &[Option<usize>],
) -> Result<Option<Vec<u8>>, ParseError> {
    let unsupported = || ParseError::UnsupportedEdit("ambiguous defined name source layout".into());
    let root = &package.workbook_template;
    root.require_unwrapped_child("definedNames")?;
    let containers = root.children_named("definedNames").collect::<Vec<_>>();
    if containers.len() > 1 {
        return Err(unsupported());
    }
    let Some(container) = containers.first() else {
        if !package.original_workbook.defined_names.is_empty() {
            return Err(unsupported());
        }
        return if wb.defined_names.is_empty() {
            Ok(None)
        } else {
            Ok(Some(fragment(|writer| write_defined_names(writer, wb))?))
        };
    };
    let template = root.capture_child(container)?;
    template.require_unwrapped_child("definedName")?;
    let sources = template.children_named("definedName").collect::<Vec<_>>();
    let originals = &package.original_workbook.defined_names;
    if sources.len() != originals.len() {
        return Err(unsupported());
    }
    let mut identities = HashSet::new();
    let mut scoped_counts = HashMap::<String, (usize, usize)>::new();
    for original in originals {
        if !identities.insert((original.name.to_ascii_lowercase(), original.local_sheet)) {
            return Err(unsupported());
        }
        if original.local_sheet.is_some() {
            scoped_counts
                .entry(original.name.to_ascii_lowercase())
                .or_default()
                .0 += 1;
        }
    }
    let mut current_identities = HashSet::new();
    let mut current_by_origin = HashMap::new();
    for (index, current) in wb.defined_names.iter().enumerate() {
        if !current_identities.insert((current.name.to_ascii_lowercase(), current.local_sheet)) {
            return Err(unsupported());
        }
        if current.local_sheet.is_some()
            && let Some(counts) = scoped_counts.get_mut(&current.name.to_ascii_lowercase())
        {
            counts.1 += 1;
        }
        let scope = match current.local_sheet {
            None => Some(None),
            Some(sheet) => origins.get(sheet.0 as usize).copied().flatten().map(Some),
        };
        if let Some(scope) = scope {
            current_by_origin.insert((current.name.to_ascii_lowercase(), scope), index);
        }
    }
    let mut used = HashSet::new();
    let mut replacements = Vec::new();
    for (source, original) in sources.iter().zip(originals) {
        if !XmlTemplate::capture(&source.bytes)?.children.is_empty() {
            return Err(unsupported());
        }
        let attributes = attributes_from_fragment(&source.bytes)?;
        let source_name = attributes
            .iter()
            .find(|attr| attr.name == "name")
            .map(|attr| attr.value.as_str());
        let source_scope = attributes
            .iter()
            .find(|attr| attr.name == "localSheetId")
            .map(|attr| attr.value.parse::<u32>())
            .transpose()
            .map_err(|_| unsupported())?;
        if source_name != Some(original.name.as_str())
            || source_scope != original.local_sheet.map(|sheet| sheet.0)
        {
            return Err(unsupported());
        }
        let identity = (
            original.name.to_ascii_lowercase(),
            original.local_sheet.map(|sheet| sheet.0 as usize),
        );
        if let Some(&index) = current_by_origin.get(&identity) {
            let current = &wb.defined_names[index];
            used.insert(index);
            replacements.push(if current == original {
                source.bytes.clone()
            } else if scoped_counts
                .get(&identity.0)
                .is_some_and(|&(before, after)| before > 1 && before != after)
                && original.local_sheet.is_some()
                && (current.name != original.name
                    || current.formula != original.formula
                    || current.hidden != original.hidden)
            {
                template.qualify_fragment(&written_defined_name(current)?)?
            } else {
                patch_defined_name(&source.bytes, current)?
            });
        }
    }
    for (index, current) in wb.defined_names.iter().enumerate() {
        if !used.contains(&index) {
            replacements.push(template.qualify_fragment(&written_defined_name(current)?)?);
        }
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

fn written_defined_name(defined: &xlsx_model::DefinedName) -> Result<Vec<u8>, ParseError> {
    fragment(|writer| write_defined_name(writer, defined))
}

fn patch_defined_name(
    source: &[u8],
    defined: &xlsx_model::DefinedName,
) -> Result<Vec<u8>, ParseError> {
    let mut reader = Reader::from_reader(source);
    let element_name = loop {
        match reader.read_event().map_err(xml_err)? {
            Event::Start(element) | Event::Empty(element) => {
                break String::from_utf8_lossy(element.name().as_ref()).into_owned();
            }
            Event::Eof => {
                return Err(ParseError::UnsupportedEdit(
                    "missing defined name element".into(),
                ));
            }
            _ => {}
        }
    };
    let mut attributes = attributes_from_fragment(source)?;
    for (name, value) in [
        ("name", Some(defined.name.clone())),
        ("hidden", defined.hidden.then(|| "1".to_owned())),
        (
            "localSheetId",
            defined.local_sheet.map(|sheet| sheet.0.to_string()),
        ),
    ] {
        if let Some(value) = value {
            if let Some(attribute) = attributes
                .iter_mut()
                .find(|attribute| attribute.name == name)
            {
                attribute.value = value;
            } else {
                attributes.push(XmlAttribute {
                    name: name.into(),
                    value,
                });
            }
        } else {
            attributes.retain(|attribute| attribute.name != name);
        }
    }
    fragment(|writer| {
        let mut element = BytesStart::new(element_name.as_str());
        for attribute in &attributes {
            element.push_attribute((attribute.name.as_str(), attribute.value.as_str()));
        }
        writer.write_event(Event::Start(element))?;
        writer.write_event(Event::Text(BytesText::new(&defined.formula)))?;
        writer.write_event(Event::End(BytesEnd::new(element_name.as_str())))?;
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

pub(crate) struct SharedStringPlan {
    retained: Vec<Option<usize>>,
    source_to_output: HashMap<usize, usize>,
    generated: HashMap<String, usize>,
}

impl SharedStringPlan {
    pub(crate) fn new(
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

    pub(crate) fn index_for(
        &self,
        value: &str,
        source: Option<usize>,
        values: &[String],
    ) -> Option<usize> {
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

#[allow(clippy::too_many_arguments)]
fn worksheet_xml_with_namespace(
    sheet: &Sheet,
    wb: &Workbook,
    main_namespace: &str,
    relationship_namespace: &str,
    links: &HyperlinkPlan,
    shared_string_plan: Option<&SharedStringPlan>,
    sheet_index: usize,
    cf_plan: &CfPlan,
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
            None,
        )?;
        if let Some(bytes) =
            auto_filter_bytes(sheet).map_err(|error| io::Error::other(error.to_string()))?
        {
            writer.get_mut().extend_from_slice(&bytes);
        }
        write_merges(writer, sheet)?;
        if let Some(bytes) = conditional_formatting_bytes(sheet, sheet_index, cf_plan, None)
            .map_err(|error| io::Error::other(error.to_string()))?
        {
            writer.get_mut().extend_from_slice(&bytes);
        }
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

/// The `<cols>` and `<sheetData>` elements with only the changed columns, rows
/// and cells rewritten and every other byte kept verbatim.
#[allow(clippy::too_many_arguments)]
fn patched_grid(
    sheet: &Sheet,
    wb: &Workbook,
    original: &Sheet,
    source: &PreservedSheet,
    axes: &SheetAxes,
    shared_string_cells: &SharedStringCells,
    shared_string_plan: Option<&SharedStringPlan>,
    styles: &StyleMatch<'_>,
) -> Result<(Option<Vec<u8>>, Vec<u8>), ParseError> {
    let mut sst_index: HashMap<&str, usize> = HashMap::with_capacity(wb.shared_strings.len());
    if shared_string_plan.is_none() {
        for (index, value) in wb.shared_strings.iter().enumerate() {
            sst_index.entry(value.as_str()).or_insert(index);
        }
    }
    let patch = SheetPatch {
        sheet,
        original,
        axes,
        workbook: wb,
        sst_index: &sst_index,
        retained: shared_string_cells,
        plan: shared_string_plan,
        styles,
    };
    let columns = patch
        .cols(
            source
                .template
                .child("cols")
                .map(|child| child.bytes.as_slice()),
        )
        .map_err(|error| unwritable_grid(&sheet.name, error))?;
    let sheet_data = match source.template.child("sheetData") {
        Some(child) => {
            let patched = patch
                .sheet_data(&child.bytes)
                .map_err(|error| unwritable_grid(&sheet.name, error))?;
            match patched {
                Some(patched) => Some(patched),
                None => match source_coordinates::explicit_rows(&child.bytes, axes)
                    .map_err(|error| unwritable_grid(&sheet.name, error))?
                {
                    Some(explicit) => patch
                        .sheet_data(&explicit)
                        .map_err(|error| unwritable_grid(&sheet.name, error))?,
                    None => None,
                },
            }
        }
        None => None,
    };
    let sheet_data = match sheet_data {
        Some(sheet_data) => sheet_data,
        None => {
            if let Some(child) = source.template.child("sheetData")
                && let Some(loss) = sheet_data_regeneration_loss(&child.bytes)
                    .map_err(|error| unwritable_grid(&sheet.name, error))?
            {
                return Err(unwritable_grid(&sheet.name, loss));
            }
            fragment(|writer| {
                write_sheet_data(
                    writer,
                    sheet,
                    wb,
                    shared_string_cells,
                    shared_string_plan,
                    Some(SourceStyles {
                        original,
                        axes,
                        styles,
                    }),
                )
            })?
        }
    };
    Ok((columns, sheet_data))
}

/// Preserved fragments (filters, validations, anchors) keep their source
/// geometry: valid XML, but stale after a row or column edit.
#[allow(clippy::too_many_arguments)]
fn worksheet_xml_with_template(
    sheet: &Sheet,
    wb: &Workbook,
    original: Option<&Sheet>,
    source: &PreservedSheet,
    package: &PreservedPackage,
    shared_string_cells: &SharedStringCells,
    shared_string_plan: Option<&SharedStringPlan>,
    sheet_axes: Option<&SheetAxes>,
    styles: &StyleMatch<'_>,
    sheet_index: usize,
    cf_plan: &CfPlan,
) -> Result<WorksheetOutput, ParseError> {
    let template = &source.template;
    if let Some(original) = original
        && original.auto_filter != sheet.auto_filter
        && sheet_axes.is_none_or(SheetAxes::is_identity)
        && wb.styles == package.original_workbook.styles
        && wb.shared_strings == package.original_workbook.shared_strings
    {
        let mut retained = original.clone();
        retained.auto_filter = sheet.auto_filter.clone();
        if retained == *sheet {
            let replacements = vec![("autoFilter", auto_filter_bytes(sheet)?)];
            return Ok(WorksheetOutput {
                bytes: render_auto_filter(template, replacements, sheet)?,
                relationships: None,
            });
        }
    }
    let patched = match (original, sheet_axes) {
        (Some(original), Some(axes)) => Some(patched_grid(
            sheet,
            wb,
            original,
            source,
            axes,
            shared_string_cells,
            shared_string_plan,
            styles,
        )?),
        _ => None,
    };
    let (columns, sheet_data) = match patched {
        Some((columns, sheet_data)) => (columns, Some(sheet_data)),
        None => (
            (!sheet.col_widths.is_empty() || !sheet.col_styles.is_empty())
                .then(|| fragment(|writer| write_cols(writer, sheet)))
                .transpose()?,
            Some(fragment(|writer| {
                write_sheet_data(
                    writer,
                    sheet,
                    wb,
                    shared_string_cells,
                    shared_string_plan,
                    original
                        .zip(sheet_axes)
                        .map(|(original, axes)| SourceStyles {
                            original,
                            axes,
                            styles,
                        }),
                )
            })?),
        ),
    };
    let merges = (!sheet.merges.is_empty())
        .then(|| fragment(|writer| write_merges(writer, sheet)))
        .transpose()?;
    let mut replacements = vec![("cols", columns), ("sheetData", sheet_data)];
    if original.is_none_or(|original| original.merges != sheet.merges) {
        replacements.push(("mergeCells", merges));
    }

    let extent = cell_extent(sheet);
    let authored = source_extent(template, sheet_axes);
    if template.child("dimension").is_some()
        && template.require_unwrapped_child("dimension").is_ok()
        && (original.is_none_or(|original| cell_extent(original) != extent)
            || authored != source_extent(template, None))
    {
        let reference =
            union_extent(extent, authored).map_or_else(|| "A1".to_owned(), |range| range.to_a1());
        replacements.push((
            "dimension",
            Some(fragment(|writer| {
                writer
                    .create_element("dimension")
                    .with_attribute(("ref", reference.as_str()))
                    .write_empty()
                    .map(|_| ())
            })?),
        ));
    }

    if original.is_none_or(|original| original.freeze_pane != sheet.freeze_pane) {
        replacements.push(("sheetViews", patched_sheet_views(template, sheet)?));
    }

    if original.is_none_or(|original| original.auto_filter != sheet.auto_filter) {
        replacements.push(("autoFilter", auto_filter_bytes(sheet)?));
    }

    if original.is_none_or(|original| {
        !keeps_source_conditional_formats(&sheet.conditional_formats, &original.conditional_formats)
    }) {
        replacements.push((
            "conditionalFormatting",
            conditional_formatting_bytes(
                sheet,
                sheet_index,
                cf_plan,
                original.map(|sheet| (template, sheet)),
            )?,
        ));
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
        bytes: render_auto_filter(template, replacements, sheet)?,
        relationships,
    })
}

fn render_auto_filter(
    template: &XmlTemplate,
    mut replacements: Vec<(&'static str, Option<Vec<u8>>)>,
    sheet: &Sheet,
) -> Result<Vec<u8>, ParseError> {
    if sheet
        .auto_filter
        .as_ref()
        .is_some_and(|filter| filter.raw_xml.is_some())
        && let Some(index) = template
            .children
            .iter()
            .position(|child| template.is_root_child_named(child, "autoFilter"))
        && let Some(replacement) = replacements
            .iter()
            .position(|(name, bytes)| *name == "autoFilter" && bytes.is_some())
    {
        let mut template = template.clone();
        template.children[index].bytes = replacements.remove(replacement).1.unwrap();
        return template.render(replacements, worksheet_child_rank);
    }
    template.render(replacements, worksheet_child_rank)
}

fn auto_filter_bytes(sheet: &Sheet) -> Result<Option<Vec<u8>>, ParseError> {
    let Some(filter) = &sheet.auto_filter else {
        return Ok(None);
    };
    let reference = filter.range.to_a1();
    let Some(raw) = &filter.raw_xml else {
        return Ok(Some(
            format!("<autoFilter ref=\"{reference}\"/>").into_bytes(),
        ));
    };
    let mut reader = Reader::from_reader(raw.as_bytes());
    let event = reader.read_event().map_err(xml_err)?;
    let (Event::Start(element) | Event::Empty(element)) = event else {
        return Err(ParseError::Malformed("invalid autoFilter source".into()));
    };
    for attribute in element.attributes() {
        let attribute = attribute.map_err(xml_err)?;
        if attribute.key.as_ref() == b"ref" {
            let value = attribute.value.as_ref();
            let start = 1 + value.as_ptr() as usize - element.as_ref().as_ptr() as usize;
            let mut bytes = raw.as_bytes().to_vec();
            bytes.splice(start..start + value.len(), reference.bytes());
            return Ok(Some(bytes));
        }
    }
    Err(ParseError::Malformed("autoFilter source has no ref".into()))
}

fn cell_extent(sheet: &Sheet) -> Option<CellRange> {
    let mut cells = sheet.iter_cells().map(|(at, _)| at);
    let first = cells.next()?;
    let (mut last_row, mut min_col, mut max_col) = (first.row, first.col, first.col);
    for at in cells {
        last_row = at.row;
        min_col = min_col.min(at.col);
        max_col = max_col.max(at.col);
    }
    Some(CellRange::new(
        CellRef::new(first.row, min_col),
        CellRef::new(last_row, max_col),
    ))
}

/// The source `<dimension>` moved through the sheet's row and column edits;
/// it still covers source cells the model does not hold.
fn source_extent(template: &XmlTemplate, axes: Option<&SheetAxes>) -> Option<CellRange> {
    let attributes = attributes_from_fragment(&template.child("dimension")?.bytes).ok()?;
    let reference = attributes
        .iter()
        .find(|attribute| attribute.local_name() == "ref")?;
    let range = CellRange::parse_a1(&reference.value).ok()?;
    let Some(axes) = axes else {
        return Some(range);
    };
    let span = |ranges: Vec<Range<u32>>| Some(ranges.first()?.start..ranges.last()?.end);
    let rows = span(
        axes.rows
            .current_ranges(range.start.row..range.end.row.saturating_add(1)),
    )?;
    let cols = span(
        axes.cols
            .current_ranges(range.start.col..range.end.col.saturating_add(1)),
    )?;
    Some(CellRange::new(
        CellRef::new(rows.start, cols.start),
        CellRef::new(rows.end - 1, cols.end - 1),
    ))
}

fn union_extent(a: Option<CellRange>, b: Option<CellRange>) -> Option<CellRange> {
    match (a, b) {
        (Some(a), Some(b)) => Some(CellRange::new(
            CellRef::new(a.start.row.min(b.start.row), a.start.col.min(b.start.col)),
            CellRef::new(a.end.row.max(b.end.row), a.end.col.max(b.end.col)),
        )),
        (a, b) => a.or(b),
    }
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
    source_styles: Option<SourceStyles<'_>>,
) -> io::Result<()> {
    let mut sst_index: HashMap<&str, usize> = HashMap::with_capacity(wb.shared_strings.len());
    if shared_string_plan.is_none() {
        for (index, value) in wb.shared_strings.iter().enumerate() {
            sst_index.entry(value.as_str()).or_insert(index);
        }
    }

    let mut cells = sheet.iter_cells().peekable();
    let mut heights = sheet.row_heights.keys().copied().peekable();
    let mut row_styles = sheet.row_styles.keys().copied().peekable();

    writer
        .create_element("sheetData")
        .write_inner_content(|writer| {
            loop {
                let cell_row = cells.peek().map(|(addr, _)| addr.row);
                let height_row = heights.peek().copied();
                let style_row = row_styles.peek().copied();
                let Some(row) = [cell_row, height_row, style_row]
                    .into_iter()
                    .flatten()
                    .min()
                else {
                    break;
                };
                write_row(
                    writer,
                    sheet,
                    row,
                    &mut cells,
                    wb,
                    &sst_index,
                    retained,
                    shared_string_plan,
                    source_styles,
                )?;
                if height_row == Some(row) {
                    heights.next();
                }
                if style_row == Some(row) {
                    row_styles.next();
                }
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

pub(crate) fn write_cols(w: &mut Writer<Vec<u8>>, sheet: &Sheet) -> io::Result<()> {
    let runs = col_runs(sheet, &[]);
    if runs.is_empty() {
        return Ok(());
    }
    w.create_element("cols").write_inner_content(|w| {
        for run in &runs {
            write_col_run(w, run)?;
        }
        Ok(())
    })?;
    Ok(())
}

/// One `<col>` element: columns sharing a width and a style.
pub(crate) struct ColRun {
    pub(crate) first: ColId,
    last: ColId,
    width: Option<f64>,
    hidden_width: Option<f64>,
    style: Option<u32>,
}

/// The model's widths and column styles as ascending, disjoint runs outside
/// `covered`. Widths stay one column per run.
pub(crate) fn col_runs(sheet: &Sheet, covered: &[Range<ColId>]) -> Vec<ColRun> {
    let mut sorted = covered.to_vec();
    sorted.sort_unstable_by_key(|range| range.start);
    let mut covered: Vec<Range<ColId>> = Vec::with_capacity(sorted.len());
    for range in sorted {
        match covered.last_mut() {
            Some(last) if range.start <= last.end => last.end = last.end.max(range.end),
            _ => covered.push(range),
        }
    }
    let styles = sheet.col_style_runs();
    let mut bounds =
        Vec::with_capacity(2 * (styles.len() + sheet.col_widths.len() + covered.len()));
    for run in &styles {
        bounds.extend([run.first, run.last.saturating_add(1)]);
    }
    for &col in sheet.col_widths.keys() {
        bounds.extend([col, col.saturating_add(1)]);
    }
    for range in &covered {
        bounds.extend([range.start, range.end]);
    }
    bounds.sort_unstable();
    bounds.dedup();
    let mut runs: Vec<ColRun> = Vec::new();
    for pair in bounds.windows(2) {
        let (first, last) = (pair[0], pair[1] - 1);
        let at = covered.partition_point(|range| range.end <= first);
        if covered.get(at).is_some_and(|range| range.start <= first) {
            continue;
        }
        let at = styles.partition_point(|run| run.last < first);
        let style = styles
            .get(at)
            .filter(|run| run.first <= first)
            .map(|run| run.xf);
        let width = sheet.col_widths.get(&first).copied();
        if width.is_none() && style.is_none() {
            continue;
        }
        match runs.last_mut() {
            Some(run)
                if width.is_none()
                    && run.width.is_none()
                    && run.style == style
                    && run.last.saturating_add(1) == first =>
            {
                run.last = last;
            }
            _ => runs.push(ColRun {
                first,
                last,
                width,
                hidden_width: sheet.hidden_col_widths.get(&first).copied(),
                style,
            }),
        }
    }
    runs
}

pub(crate) fn write_col_run<W: io::Write>(w: &mut Writer<W>, run: &ColRun) -> io::Result<()> {
    let first = (u64::from(run.first) + 1).to_string();
    let last = (u64::from(run.last) + 1).to_string();
    let hidden = run.width == Some(0.0);
    let width = run.width.map(|width| {
        fmt_num(if hidden {
            run.hidden_width.unwrap_or(width)
        } else {
            width
        })
    });
    let style = run.style.map(|style| style.to_string());
    let mut element = BytesStart::new("col");
    element.push_attribute(("min", first.as_str()));
    element.push_attribute(("max", last.as_str()));
    if let Some(width) = &width {
        element.push_attribute(("width", width.as_str()));
        element.push_attribute(("customWidth", "1"));
    }
    if let Some(style) = &style {
        element.push_attribute(("style", style.as_str()));
    }
    if hidden {
        element.push_attribute(("hidden", "1"));
    }
    w.write_event(Event::Empty(element))?;
    Ok(())
}

#[allow(clippy::too_many_arguments)]
pub(crate) fn write_row<'a, I>(
    w: &mut Writer<Vec<u8>>,
    sheet: &Sheet,
    row: RowId,
    cells: &mut Peekable<I>,
    wb: &Workbook,
    sst_index: &HashMap<&str, usize>,
    retained: &SharedStringCells,
    shared_string_plan: Option<&SharedStringPlan>,
    source_styles: Option<SourceStyles<'_>>,
) -> io::Result<()>
where
    I: Iterator<Item = (CellRef, &'a Cell)>,
{
    let r = (row as u64 + 1).to_string();
    let mut start = BytesStart::new("row");
    start.push_attribute(("r", r.as_str()));
    let style = sheet.row_styles.get(&row).map(u32::to_string);
    if let Some(style) = &style {
        start.push_attribute(("s", style.as_str()));
        start.push_attribute(("customFormat", "1"));
    }
    let hidden = sheet.row_heights.get(&row) == Some(&0.0);
    let ht = sheet.row_heights.get(&row).map(|h| {
        fmt_num(if hidden {
            sheet.hidden_row_heights.get(&row).copied().unwrap_or(*h)
        } else {
            *h
        })
    });
    if let Some(h) = &ht {
        start.push_attribute(("ht", h.as_str()));
        start.push_attribute(("customHeight", "1"));
        if hidden {
            start.push_attribute(("hidden", "1"));
        }
    }
    w.write_event(Event::Start(start))?;
    while let Some((addr, cell)) = cells.next_if(|(addr, _)| addr.row == row) {
        let retained = shared_string_index(cell, addr, wb, sst_index, retained, shared_string_plan);
        let style = source_styles.map_or(cell.style, |source| source.style(addr, cell));
        write_cell(
            w,
            addr,
            cell,
            style,
            sst_index,
            retained,
            sheet.array_formula(addr),
        )?;
    }
    w.write_event(Event::End(BytesEnd::new("row")))?;
    Ok(())
}

/// The shared-string index a cell serializes against: its authored entry when
/// it still holds it, else the output table.
pub(crate) fn shared_string_index(
    cell: &Cell,
    at: CellRef,
    wb: &Workbook,
    sst_index: &HashMap<&str, usize>,
    retained: &SharedStringCells,
    shared_string_plan: Option<&SharedStringPlan>,
) -> Option<usize> {
    let source = retained.get(&(at.row, at.col)).copied();
    match (&cell.value, shared_string_plan) {
        (CellValue::Text { value }, Some(plan)) => {
            plan.index_for(value, source, &wb.shared_strings)
        }
        (CellValue::Text { value }, None) => source
            .filter(|index| wb.shared_strings.get(*index) == Some(value))
            .or_else(|| sst_index.get(value.as_str()).copied()),
        _ => None,
    }
}

/// serialize a single cell with `style` as its `s`, choosing the `t` type and
/// body from its value and whether it carries a formula.
pub(crate) fn write_cell(
    w: &mut Writer<Vec<u8>>,
    addr: CellRef,
    cell: &Cell,
    style: Option<u32>,
    sst_index: &HashMap<&str, usize>,
    retained: Option<usize>,
    array_ref: Option<xlsx_model::CellRange>,
) -> io::Result<()> {
    let a1 = addr.to_a1();
    let CellValueMarkup { ty, value, inline } = cell_value_markup(cell, sst_index, retained);

    let mut start = BytesStart::new("c");
    start.push_attribute(("r", a1.as_str()));
    let style = style.map(|s| s.to_string());
    if let Some(s) = &style {
        start.push_attribute(("s", s.as_str()));
    }
    if let Some(t) = ty {
        start.push_attribute(("t", t));
    }
    w.write_event(Event::Start(start))?;
    if let Some(f) = &cell.formula {
        let reference = array_ref.map(|range| range.to_a1());
        let mut element = w.create_element("f");
        if let Some(reference) = &reference {
            element = element
                .with_attribute(("t", "array"))
                .with_attribute(("ref", reference.as_str()));
        }
        element.write_text_content(BytesText::new(f))?;
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

pub(crate) struct CellValueMarkup {
    pub(crate) ty: Option<&'static str>,
    pub(crate) value: Option<String>,
    inline: Option<String>,
}

pub(crate) fn cell_value_markup(
    cell: &Cell,
    sst_index: &HashMap<&str, usize>,
    retained: Option<usize>,
) -> CellValueMarkup {
    let mut markup = CellValueMarkup {
        ty: None,
        value: None,
        inline: None,
    };
    match &cell.value {
        CellValue::Empty => {}
        CellValue::Number { value } => markup.value = Some(fmt_num(*value)),
        CellValue::Bool { value } => {
            markup.ty = Some("b");
            markup.value = Some(if *value { "1" } else { "0" }.to_owned());
        }
        CellValue::Error { value } => {
            markup.ty = Some("e");
            markup.value = Some(value.as_str().to_owned());
        }
        CellValue::Text { value } => {
            if cell.formula.is_some() {
                markup.ty = Some("str");
                markup.value = Some(value.clone());
            } else if let Some(index) = retained.or_else(|| sst_index.get(value.as_str()).copied())
            {
                markup.ty = Some("s");
                markup.value = Some(index.to_string());
            } else {
                markup.ty = Some("inlineStr");
                markup.inline = Some(value.clone());
            }
        }
    }
    markup
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

fn styles_xml_with_authored(
    ss: &Stylesheet,
    authored: &[DifferentialFormat],
) -> Result<Vec<u8>, ParseError> {
    styles_xml_with_namespace(ss, NS_MAIN, authored)
}

fn write_dxfs(w: &mut Writer<Vec<u8>>, dxfs: &[DifferentialFormat]) -> io::Result<()> {
    if dxfs.is_empty() {
        return Ok(());
    }
    let count = dxfs.len().to_string();
    w.create_element("dxfs")
        .with_attribute(("count", count.as_str()))
        .write_inner_content(|w| {
            for dxf in dxfs {
                w.get_mut().extend_from_slice(&dxf_bytes(dxf));
            }
            Ok(())
        })?;
    Ok(())
}

fn styles_xml_with_namespace(
    ss: &Stylesheet,
    main_namespace: &str,
    authored: &[DifferentialFormat],
) -> Result<Vec<u8>, ParseError> {
    doc(|w| {
        w.create_element("styleSheet")
            .with_attribute(("xmlns", main_namespace))
            .write_inner_content(|w| {
                write_num_fmts(w, ss)?;
                write_fonts(w, ss)?;
                write_fills(w, ss)?;
                write_borders(w, ss)?;
                write_cell_xfs(w, ss)?;
                let mut dxfs = ss.dxfs.clone();
                dxfs.extend(authored.iter().cloned());
                write_dxfs(w, &dxfs)?;
                write_colors(w, ss)?;
                Ok(())
            })?;
        Ok(())
    })
}

fn source_dxf_count(template: &XmlTemplate) -> Result<usize, ParseError> {
    let Some(pool) = template.child("dxfs") else {
        return Ok(0);
    };
    let mut reader = crate::xml::reader(&pool.bytes);
    let mut buf = Vec::new();
    let mut depth = 0;
    let mut count = 0;
    loop {
        match crate::xml::next_event(&mut reader, &mut buf, &mut depth)? {
            Event::Start(e) if depth == 2 && e.local_name().as_ref() == b"dxf" => count += 1,
            Event::Eof => return Ok(count),
            _ => {}
        }
    }
}

fn patched_dxfs(
    template: &XmlTemplate,
    authored: &[DifferentialFormat],
) -> Result<Option<Vec<u8>>, ParseError> {
    if authored.is_empty() {
        return Ok(None);
    }
    template.require_unwrapped_child("dxfs")?;
    let Some(pool) = template
        .child("dxfs")
        .map(|child| XmlTemplate::capture(&child.bytes))
        .transpose()?
    else {
        return Ok(Some(fragment(|writer| {
            let dxfs = authored.to_vec();
            write_dxfs(writer, &dxfs)
        })?));
    };
    pool.require_unwrapped_child("dxf")?;
    let mut items = pool
        .children_named("dxf")
        .map(|child| child.bytes.clone())
        .collect::<Vec<_>>();
    for format in authored {
        items.push(pool.qualify_fragment(&dxf_bytes(format))?);
    }
    pool.render_repeated("dxf", &items, &[("count", items.len().to_string())])
        .map(Some)
}

/// Reuses each source pool entry `kept` says the model left alone, so font
/// schemes, `gray125` fills and `xf` attributes the model does not carry
/// survive an edit that only appends new entries.
#[allow(clippy::too_many_arguments)]
fn patched_pool<T>(
    template: &XmlTemplate,
    pool_name: &str,
    item_name: &str,
    values: &[T],
    kept: impl Fn(usize, &T) -> bool,
    write_item: impl Fn(&mut Writer<Vec<u8>>, &T) -> io::Result<()>,
    write_pool: impl FnOnce() -> Result<Vec<u8>, ParseError>,
    overrides: Option<&HashMap<usize, Vec<u8>>>,
) -> Result<Option<Vec<u8>>, ParseError> {
    template.require_unwrapped_child(pool_name)?;
    let pool = template
        .child(pool_name)
        .map(|child| XmlTemplate::capture(&child.bytes))
        .transpose()?;
    if let Some(pool) = &pool {
        pool.require_unwrapped_child(item_name)?;
    }
    if values.is_empty() {
        return Ok(None);
    }
    let Some(pool) = pool else {
        return write_pool().map(Some);
    };
    let sources = pool.children_named(item_name).collect::<Vec<_>>();
    let mut items = Vec::with_capacity(values.len());
    for (index, value) in values.iter().enumerate() {
        if let Some(bytes) = overrides.and_then(|values| values.get(&index)) {
            items.push(pool.qualify_fragment(bytes)?);
            continue;
        }
        match sources.get(index).filter(|_| kept(index, value)) {
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
    overrides: &style_derivation::StyleOverrides,
    authored_dxfs: &[DifferentialFormat],
) -> Result<Vec<u8>, ParseError> {
    let num_fmts = patched_pool(
        template,
        "numFmts",
        "numFmt",
        &stylesheet.num_fmts,
        |index, value| original.num_fmts.get(index) == Some(value),
        |writer, (id, code)| {
            writer
                .create_element("numFmt")
                .with_attribute(("numFmtId", id.to_string().as_str()))
                .with_attribute(("formatCode", code.as_str()))
                .write_empty()?;
            Ok(())
        },
        || fragment(|writer| write_num_fmts(writer, stylesheet)),
        None,
    )?;
    let fonts = patched_pool(
        template,
        "fonts",
        "font",
        &stylesheet.fonts,
        |index, value| original.fonts.get(index) == Some(value),
        write_font,
        || fragment(|writer| write_fonts(writer, stylesheet)),
        Some(&overrides.fonts),
    )?;
    let fills = patched_pool(
        template,
        "fills",
        "fill",
        &stylesheet.fills,
        |index, value| original.fills.get(index) == Some(value),
        write_fill,
        || fragment(|writer| write_fills(writer, stylesheet)),
        Some(&overrides.fills),
    )?;
    let borders = patched_pool(
        template,
        "borders",
        "border",
        &stylesheet.borders,
        |index, value| original.borders.get(index) == Some(value),
        write_border,
        || fragment(|writer| write_borders(writer, stylesheet)),
        Some(&overrides.borders),
    )?;
    let cell_xfs = patched_pool(
        template,
        "cellXfs",
        "xf",
        &stylesheet.cell_xfs,
        |index, _| keeps_source_xf(stylesheet, original, index),
        write_xf,
        || fragment(|writer| write_cell_xfs(writer, stylesheet)),
        Some(&overrides.xfs),
    )?;
    let mut replacements = vec![
        ("numFmts", num_fmts),
        ("fonts", fonts),
        ("fills", fills),
        ("borders", borders),
        ("cellXfs", cell_xfs),
    ];
    if !authored_dxfs.is_empty() {
        replacements.push(("dxfs", patched_dxfs(template, authored_dxfs)?));
    }
    if stylesheet.indexed_colors != original.indexed_colors {
        let colors = match template.child("colors") {
            Some(source) => {
                let colors = XmlTemplate::capture(&source.bytes)?;
                let indexed = if stylesheet.indexed_colors.is_empty() {
                    None
                } else {
                    Some(fragment(|writer| write_indexed_colors(writer, stylesheet))?)
                };
                Some(colors.render(vec![("indexedColors", indexed)], |name| {
                    if name == "indexedColors" { 0 } else { 1 }
                })?)
            }
            None if stylesheet.indexed_colors.is_empty() => None,
            None => Some(fragment(|writer| write_colors(writer, stylesheet))?),
        };
        replacements.push(("colors", colors));
    }
    template.render(replacements, stylesheet_child_rank)
}

fn write_colors(w: &mut Writer<Vec<u8>>, ss: &Stylesheet) -> io::Result<()> {
    if !ss.indexed_colors.is_empty() {
        w.create_element("colors")
            .write_inner_content(|w| write_indexed_colors(w, ss))?;
    }
    Ok(())
}

fn validate_indexed_colors(ss: &Stylesheet) -> Result<(), ParseError> {
    for (index, color) in ss.indexed_colors.iter().enumerate() {
        let rgb = color.strip_prefix('#').unwrap_or(color);
        if !color.is_empty()
            && (rgb.len() != 6 || !rgb.bytes().all(|byte| byte.is_ascii_hexdigit()))
        {
            return Err(ParseError::Malformed(format!(
                "indexed palette color at index {index} must be empty or six hexadecimal RGB digits with an optional '#'"
            )));
        }
    }
    Ok(())
}

fn write_indexed_colors(w: &mut Writer<Vec<u8>>, ss: &Stylesheet) -> io::Result<()> {
    w.create_element("indexedColors").write_inner_content(|w| {
        for rgb in &ss.indexed_colors {
            if rgb.is_empty() {
                w.create_element("rgbColor").write_empty()?;
            } else {
                let rgb = format!(
                    "FF{}",
                    rgb.strip_prefix('#').unwrap_or(rgb).to_ascii_uppercase()
                );
                w.create_element("rgbColor")
                    .with_attribute(("rgb", rgb.as_str()))
                    .write_empty()?;
            }
        }
        Ok(())
    })?;
    Ok(())
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
            Fill::Pattern { pattern, fg, bg } => {
                w.create_element("patternFill")
                    .with_attribute(("patternType", pattern.as_sml()))
                    .write_inner_content(|w| {
                        if let Some(color) = fg {
                            write_color(w, "fgColor", color)?;
                        }
                        if let Some(color) = bg {
                            write_color(w, "bgColor", color)?;
                        }
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
    let xf_id;
    if let Some(id) = xf.xf_id {
        xf_id = id.to_string();
        el.push_attribute(("xfId", xf_id.as_str()));
    }
    if xf.quote_prefix {
        el.push_attribute(("quotePrefix", "1"));
    }
    if xf.pivot_button {
        el.push_attribute(("pivotButton", "1"));
    }
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
    if xf.protection.is_some() {
        el.push_attribute(("applyProtection", "1"));
    }

    if xf.alignment.is_none() && xf.protection.is_none() {
        w.write_event(Event::Empty(el))?;
    } else {
        w.write_event(Event::Start(el))?;
        if let Some(a) = &xf.alignment {
            write_alignment(w, a)?;
        }
        if let Some(protection) = &xf.protection {
            w.create_element("protection")
                .with_attribute(("locked", if protection.locked { "1" } else { "0" }))
                .with_attribute(("hidden", if protection.hidden { "1" } else { "0" }))
                .write_empty()?;
        }
        w.write_event(Event::End(BytesEnd::new("xf")))?;
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
    let rotation;
    if let Some(value) = a.text_rotation {
        rotation = value.to_string();
        el.push_attribute(("textRotation", rotation.as_str()));
    }
    if a.wrap_text {
        el.push_attribute(("wrapText", "1"));
    }
    if a.shrink_to_fit {
        el.push_attribute(("shrinkToFit", "1"));
    }
    let indent;
    if let Some(value) = a.indent {
        indent = value.to_string();
        el.push_attribute(("indent", indent.as_str()));
    }
    let relative_indent;
    if let Some(value) = a.relative_indent {
        relative_indent = value.to_string();
        el.push_attribute(("relativeIndent", relative_indent.as_str()));
    }
    if let Some(value) = a.justify_last_line {
        el.push_attribute(("justifyLastLine", if value { "1" } else { "0" }));
    }
    let reading_order;
    if let Some(value) = a.reading_order {
        reading_order = value.to_string();
        el.push_attribute(("readingOrder", reading_order.as_str()));
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
pub(crate) fn fmt_num(n: f64) -> String {
    if n.is_finite() && n.fract() == 0.0 && n.abs() < 1e15 {
        return format!("{}", n as i64);
    }
    format!("{n}")
}
