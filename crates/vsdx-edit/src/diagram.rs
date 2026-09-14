use std::collections::HashSet;
use std::sync::Arc;

use vsdx_eval::{MutationContext, MutationOutcome, decide_mutation, evaluate};
use vsdx_parse::{
    Cell, CellLocator, CellRow, CellSheet, Connect, ConnectsChild, MutationGesture, ParseLimits,
    RowChild, SectionChild, Shape, ShapeChild, ShapesChild, SheetChild, StructuralEdit,
};
use vsdx_resolve::{Lookup, Resolver};
use yrs::{
    Any, Array, ArrayPrelim, ArrayRef, Doc, Map, MapPrelim, MapRef, Out, ReadTxn, Transact,
    TransactionMut, WriteTxn,
};

use crate::{
    CONNECTS, CellFormulaReceipt, CellSnapshot, ConnectorGlue, DiagramSession, DiagramSnapshot,
    EditCtx, EditError, EditResult, META, PAGE_ORDER, PAGES, PageSnapshot, SHEETS, STORIES,
    ShapeDraft, ShapeReceipt, ShapeSnapshot,
};

const SCHEMA_VERSION: f64 = 1.0;
pub(crate) const MAX_SHAPE_NESTING: usize = 256;
type SectionRows<'a> = Vec<(
    (String, Option<u32>),
    Vec<(Option<CellRow>, Vec<&'a CellSnapshot>)>,
)>;

pub(crate) fn seed_doc(
    doc: &Doc,
    package: &vsdx_parse::VsdxPackage,
    fingerprint: &str,
) -> EditResult<()> {
    let package_json =
        serde_json::to_vec(package).map_err(|error| EditError::Json(error.to_string()))?;
    let package_bytes =
        vsdx_parse::write_vsdx(package).map_err(|error| EditError::Parse(error.to_string()))?;
    let mut txn = doc.transact_mut_with("vsdx:bootstrap");
    let meta = txn.get_or_insert_map(META);
    meta.insert(&mut txn, "schemaVersion", SCHEMA_VERSION);
    meta.insert(&mut txn, "fingerprint", fingerprint);
    meta.insert(
        &mut txn,
        "packageJson",
        Any::Buffer(Arc::from(package_json)),
    );
    meta.insert(
        &mut txn,
        "packageBytes",
        Any::Buffer(Arc::from(package_bytes)),
    );
    meta.insert(&mut txn, "pageWidth", 0.0);
    meta.insert(&mut txn, "pageHeight", 0.0);
    let order = txn.get_or_insert_array(PAGE_ORDER);
    let pages = txn.get_or_insert_map(PAGES);
    let sheets = txn.get_or_insert_map(SHEETS);
    txn.get_or_insert_map(CONNECTS);
    let stories = txn.get_or_insert_map(STORIES);
    for path in &package.page_part_paths {
        let Some(page_id) = package.page_part_ids.get(path) else {
            continue;
        };
        let id = format!("page:{page_id}");
        order.push_back(&mut txn, id.as_str());
        let page = pages.insert(&mut txn, id.as_str(), MapPrelim::default());
        page.insert(&mut txn, "id", id.as_str());
        if let Some(name) = package.page_names.get(page_id) {
            page.insert(&mut txn, "name", name.as_str());
        }
        page.insert(&mut txn, "sourcePartPath", path.as_str());
        page.insert(
            &mut txn,
            "maxSourceId",
            package
                .page_contents
                .get(path)
                .map(largest_shape_id)
                .unwrap_or_default() as f64,
        );
        let shape_order = page.insert(&mut txn, "shapes", ArrayPrelim::default());
        if let Some(sheet) = package.page_contents.get(path) {
            let resolver = Resolver::new(package);
            for shape in sheet.shapes() {
                let shape_id = format!("{id}:shape:{}", shape.id);
                shape_order.push_back(&mut txn, shape_id.as_str());
                let resolved = resolver
                    .resolve_shape(path, shape.id)
                    .map_err(|error| EditError::InvalidState(error.to_string()))?;
                seed_shape(
                    &sheets, &stories, &mut txn, &shape_id, &id, None, path, sheet, shape,
                    &resolver, &resolved, 1,
                )?;
            }
        }
    }
    Ok(())
}

pub(crate) fn package_from_doc(doc: &Doc) -> EditResult<vsdx_parse::VsdxPackage> {
    let mut package = original_package_from_doc(doc)?;
    let snapshot = snapshot_doc(doc)?;
    let glue = glue_records(&doc.transact())?;
    materialize_snapshot(&mut package, &snapshot, &original_shape_ids(doc)?, &glue)?;
    package.page_part_paths = page_part_paths_for_snapshot(&package, &snapshot)?;
    Ok(package)
}

fn original_package_from_doc(doc: &Doc) -> EditResult<vsdx_parse::VsdxPackage> {
    let txn = doc.transact();
    let meta = required_map(&txn, META)?;
    if map_number(&meta, &txn, "schemaVersion") != Some(SCHEMA_VERSION) {
        return Err(EditError::InvalidState(
            "unsupported diagram schema version".to_owned(),
        ));
    }
    match meta.get(&txn, "packageBytes") {
        Some(Out::Any(Any::Buffer(bytes))) => vsdx_parse::parse_vsdx(&bytes)
            .map_err(|error| EditError::InvalidState(error.to_string())),
        _ => {
            let Some(Out::Any(Any::Buffer(bytes))) = meta.get(&txn, "packageJson") else {
                return Err(EditError::InvalidState("missing package data".to_owned()));
            };
            serde_json::from_slice::<vsdx_parse::VsdxPackage>(&bytes)
                .map_err(|error| EditError::InvalidState(error.to_string()))
        }
    }
}

fn original_shape_ids(doc: &Doc) -> EditResult<HashSet<String>> {
    let txn = doc.transact();
    let sheets = required_map(&txn, SHEETS)?;
    let mut ids = HashSet::new();
    for (id, value) in sheets.iter(&txn) {
        let Out::YMap(shape) = value else { continue };
        if shape_origin(&shape, &txn)? == ShapeOrigin::Original {
            ids.insert(id.to_owned());
        }
    }
    Ok(ids)
}

fn materialize_snapshot(
    package: &mut vsdx_parse::VsdxPackage,
    snapshot: &DiagramSnapshot,
    original_shape_ids: &HashSet<String>,
    glue: &[GlueRecord],
) -> EditResult<()> {
    for page in &snapshot.pages {
        let Some(sheet) = package.page_contents.get_mut(&page.source_part_path) else {
            continue;
        };
        let original_source_shape_ids = sheet_shape_ids(sheet);
        materialize_page_shapes(sheet, page, original_shape_ids)?;
        let projected_shape_ids = sheet_shape_ids(sheet);
        let deleted = original_source_shape_ids
            .difference(&projected_shape_ids)
            .copied()
            .collect();
        vsdx_parse::remove_connects_referencing_shapes(sheet, &deleted);
        materialize_page_glue(sheet, page, glue)?;
    }
    Ok(())
}

struct GlueRecord {
    id: String,
    page_id: String,
    connector_id: String,
    endpoint: GlueEndpoint,
    target_id: String,
    to_cell: String,
}

#[derive(Clone, Copy, PartialEq, Eq, Hash, PartialOrd, Ord)]
enum GlueEndpoint {
    Begin,
    End,
}

impl GlueEndpoint {
    fn name(self) -> &'static str {
        match self {
            Self::Begin => "begin",
            Self::End => "end",
        }
    }
    fn endpoint_cell(self) -> &'static str {
        match self {
            Self::Begin => "BeginX",
            Self::End => "EndX",
        }
    }
    fn parse(value: &str) -> Option<Self> {
        match value {
            "begin" => Some(Self::Begin),
            "end" => Some(Self::End),
            _ => None,
        }
    }
}

fn glue_key(connector_id: &str, endpoint: GlueEndpoint) -> String {
    format!("{connector_id}:{}", endpoint.name())
}

fn normalized_glue_target(cell: Option<&str>) -> Option<String> {
    match cell {
        None => Some("PinX".to_owned()),
        Some(value) if valid_glue_target(value) => Some(value.to_owned()),
        _ => None,
    }
}

fn valid_glue_target(cell: &str) -> bool {
    if matches!(cell, "PinX" | "PinY") {
        return true;
    }
    cell.strip_prefix("Connections.X")
        .and_then(|ordinal| ordinal.parse::<u32>().ok())
        .is_some_and(|ordinal| ordinal >= 1)
}

/// Reports whether `to_cell` names a connection row present on the target.
fn connection_point_exists<T: ReadTxn>(
    sheets: &yrs::MapRef,
    txn: &T,
    target_id: &str,
    to_cell: &str,
) -> bool {
    let Some(row) = to_cell
        .strip_prefix("Connections.X")
        .and_then(|ordinal| ordinal.parse::<u32>().ok())
        .and_then(|ordinal| ordinal.checked_sub(1))
    else {
        return true;
    };
    let Some(yrs::Out::YMap(shape)) = sheets.get(txn, target_id) else {
        return false;
    };
    let Ok(cells) = map_map(&shape, txn, "cells") else {
        return false;
    };
    for (_, value) in cells.iter(txn) {
        let yrs::Out::YMap(cell) = value else {
            continue;
        };
        if map_string(&cell, txn, "section").as_deref() != Some("Connection") {
            continue;
        }
        if map_u32(&cell, txn, "rowIndex").ok().flatten() == Some(row) {
            return true;
        }
    }
    false
}

fn glue_text_valid(value: &str) -> bool {
    let limits = ParseLimits::default();
    value.len() <= limits.max_attribute_bytes
        && !value.is_empty()
        && value
            .chars()
            .all(|c| matches!(c, '\t' | '\r' | '\n' | '\u{20}'..='\u{d7ff}' | '\u{e000}'..='\u{fffd}' | '\u{10000}'..='\u{10ffff}'))
}

/// Explicit numeric `OneD` decides alone, matching the resolver.
fn draft_is_one_d(draft: &ShapeDraft) -> bool {
    let mut endpoints = HashSet::new();
    for cell in &draft.cells {
        if cell.locator.section.is_some() {
            continue;
        }
        match cell.name.as_str() {
            "OneD" => {
                let formula = cell.formula.as_deref().unwrap_or_default();
                match formula.trim_start_matches('=').trim().parse::<f64>() {
                    Ok(value) => return value != 0.0,
                    Err(_) => return true,
                }
            }
            "BeginX" | "BeginY" | "EndX" | "EndY" => {
                endpoints.insert(cell.name.as_str());
            }
            _ => {}
        }
    }
    ["BeginX", "BeginY", "EndX", "EndY"]
        .into_iter()
        .all(|name| endpoints.contains(name))
}

fn glue_records<T: ReadTxn>(txn: &T) -> EditResult<Vec<GlueRecord>> {
    let Some(connects) = txn.get_map(CONNECTS) else {
        return Ok(Vec::new());
    };
    let mut records = Vec::with_capacity(connects.len(txn) as usize);
    for (key, value) in connects.iter(txn) {
        let Out::YMap(entry) = value else {
            return Err(EditError::InvalidState(
                "connector glue is not a map".to_owned(),
            ));
        };
        let endpoint = map_string(&entry, txn, "endpoint")
            .as_deref()
            .and_then(GlueEndpoint::parse)
            .ok_or_else(|| EditError::InvalidState("connector glue has no endpoint".to_owned()))?;
        let to_cell = map_string(&entry, txn, "toCell")
            .filter(|cell| valid_glue_target(cell))
            .ok_or_else(|| {
                EditError::InvalidState("connector glue has no target cell".to_owned())
            })?;
        records.push(GlueRecord {
            id: map_string(&entry, txn, "id")
                .filter(|id| id == key)
                .ok_or_else(|| {
                    EditError::InvalidState("connector glue ID does not match map key".to_owned())
                })?,
            page_id: map_string(&entry, txn, "pageId").ok_or_else(|| {
                EditError::InvalidState("connector glue is missing its page".to_owned())
            })?,
            connector_id: map_string(&entry, txn, "connectorId").ok_or_else(|| {
                EditError::InvalidState("connector glue is missing its connector".to_owned())
            })?,
            endpoint,
            target_id: map_string(&entry, txn, "targetId").ok_or_else(|| {
                EditError::InvalidState("connector glue is missing its target".to_owned())
            })?,
            to_cell,
        });
    }
    records.sort_by(|left, right| left.id.cmp(&right.id));
    Ok(records)
}

fn snapshot_shape_sources(page: &PageSnapshot) -> std::collections::BTreeMap<&str, u32> {
    let mut sources = std::collections::BTreeMap::new();
    let mut pending = page.shapes.iter().collect::<Vec<_>>();
    while let Some(shape) = pending.pop() {
        sources.insert(shape.id.as_str(), shape.source_id);
        pending.extend(shape.children.iter());
    }
    sources
}

fn materialize_page_glue(
    sheet: &mut vsdx_parse::Sheet,
    page: &PageSnapshot,
    glue: &[GlueRecord],
) -> EditResult<()> {
    let sources = snapshot_shape_sources(page);
    let mut pending = Vec::new();
    for record in glue.iter().filter(|record| record.page_id == page.id) {
        let (Some(connector), Some(target)) = (
            sources.get(record.connector_id.as_str()),
            sources.get(record.target_id.as_str()),
        ) else {
            return Err(EditError::InvalidState(
                "connector glue references a missing shape".to_owned(),
            ));
        };
        pending.push(Connect {
            from_sheet: *connector,
            from_cell: Some(record.endpoint.endpoint_cell().to_owned()),
            from_part: None,
            to_sheet: *target,
            to_cell: Some(record.to_cell.clone()),
            to_part: None,
            other_attrs: Vec::new(),
        });
    }
    if pending.is_empty() {
        return Ok(());
    }
    match sheet.children.iter_mut().find_map(|child| match child {
        SheetChild::Connects(connects) => Some(connects),
        _ => None,
    }) {
        Some(connects) => {
            connects.extend(pending.into_iter().map(ConnectsChild::Connect));
        }
        None => {
            sheet.children.push(SheetChild::Connects(
                pending.into_iter().map(ConnectsChild::Connect).collect(),
            ));
        }
    }
    Ok(())
}

fn sheet_shape_ids(sheet: &vsdx_parse::Sheet) -> HashSet<u32> {
    let mut ids = HashSet::new();
    let mut pending = sheet.shapes().collect::<Vec<_>>();
    while let Some(shape) = pending.pop() {
        ids.insert(shape.id);
        pending.extend(shape.shapes());
    }
    ids
}

fn materialize_page_shapes(
    sheet: &mut vsdx_parse::Sheet,
    page: &PageSnapshot,
    original_shape_ids: &HashSet<String>,
) -> EditResult<()> {
    let originals = sheet
        .shapes()
        .cloned()
        .map(|shape| (shape.id, shape))
        .collect::<std::collections::BTreeMap<_, _>>();
    let mut shapes = Vec::with_capacity(page.shapes.len());
    for snapshot in &page.shapes {
        let mut shape = if original_shape_ids.contains(&snapshot.id) {
            originals
                .get(&snapshot.source_id)
                .cloned()
                .unwrap_or_else(|| shape_from_snapshot(snapshot))
        } else {
            shape_from_snapshot(snapshot)
        };
        materialize_shape(&mut shape, snapshot, original_shape_ids, 1)?;
        shapes.push(ShapesChild::Shape(shape));
    }
    if let Some(SheetChild::Shapes(existing)) = sheet
        .children
        .iter_mut()
        .find(|child| matches!(child, SheetChild::Shapes(_)))
    {
        *existing = shapes;
    } else {
        sheet.children.push(SheetChild::Shapes(shapes));
    }
    Ok(())
}

fn shape_from_snapshot(snapshot: &ShapeSnapshot) -> Shape {
    Shape {
        id: snapshot.source_id,
        name: snapshot.name.clone(),
        name_u: None,
        shape_type: Some("Shape".to_owned()),
        master: None,
        master_shape: None,
        line_style: None,
        fill_style: None,
        text_style: None,
        children: Vec::new(),
        del: false,
        other_attrs: Vec::new(),
    }
}

fn largest_shape_id(sheet: &vsdx_parse::Sheet) -> u32 {
    let mut largest = 0;
    let mut pending = sheet.shapes().collect::<Vec<_>>();
    while let Some(shape) = pending.pop() {
        largest = largest.max(shape.id);
        pending.extend(shape.shapes());
    }
    largest
}

fn materialize_shape(
    shape: &mut Shape,
    snapshot: &ShapeSnapshot,
    original_shape_ids: &HashSet<String>,
    depth: usize,
) -> EditResult<()> {
    if depth > MAX_SHAPE_NESTING {
        return Err(EditError::InvalidState(
            "shape nesting exceeds maximum depth".to_owned(),
        ));
    }
    for cell in &snapshot.cells {
        materialize_cell(shape, cell);
    }
    let originals = shape
        .shapes()
        .cloned()
        .map(|child| (child.id, child))
        .collect::<std::collections::BTreeMap<_, _>>();
    let mut children = Vec::with_capacity(snapshot.children.len());
    for child_snapshot in &snapshot.children {
        let mut child = if original_shape_ids.contains(&child_snapshot.id) {
            originals
                .get(&child_snapshot.source_id)
                .cloned()
                .unwrap_or_else(|| shape_from_snapshot(child_snapshot))
        } else {
            shape_from_snapshot(child_snapshot)
        };
        materialize_shape(&mut child, child_snapshot, original_shape_ids, depth + 1)?;
        children.push(ShapesChild::Shape(child));
    }
    if let Some(ShapeChild::Shapes(existing)) = shape
        .children
        .iter_mut()
        .find(|child| matches!(child, ShapeChild::Shapes(_)))
    {
        *existing = children;
    } else {
        shape.children.push(ShapeChild::Shapes(children));
    }
    Ok(())
}

fn materialize_cell(shape: &mut Shape, snapshot: &CellSnapshot) {
    let locator = &snapshot.locator;
    let target = match &locator.section {
        None => shape.children.iter_mut().find_map(|child| match child {
            ShapeChild::Cell(cell) if cell.name == locator.cell_name => Some(cell),
            _ => None,
        }),
        Some(section_name) => shape.children.iter_mut().find_map(|child| {
            let ShapeChild::Section(section) = child else {
                return None;
            };
            if section.name != *section_name
                || section.index.unwrap_or(0) != locator.section_index.unwrap_or(0)
            {
                return None;
            }
            section.children.iter_mut().find_map(|child| {
                let SectionChild::Row(row) = child else {
                    return None;
                };
                let row_matches = match &locator.row {
                    Some(CellRow::Index(index)) => row.index == Some(*index),
                    Some(CellRow::Name(name)) => row.name.as_deref() == Some(name),
                    None => false,
                };
                row_matches.then(|| {
                    row.children.iter_mut().find_map(|child| match child {
                        RowChild::Cell(cell) if cell.name == locator.cell_name => Some(cell),
                        _ => None,
                    })
                })?
            })
        }),
    };
    if let Some(cell) = target {
        cell.formula = snapshot.formula.clone();
        cell.value = snapshot.value.clone();
    } else if locator.section.is_none() {
        shape.children.push(ShapeChild::Cell(Cell {
            name: locator.cell_name.clone(),
            formula: snapshot.formula.clone(),
            value: snapshot.value.clone(),
            unit: None,
            del: false,
            other_attrs: Vec::new(),
        }));
    } else if let (Some(section_name), Some(row)) = (&locator.section, &locator.row) {
        let cell = Cell {
            name: locator.cell_name.clone(),
            formula: snapshot.formula.clone(),
            value: snapshot.value.clone(),
            unit: None,
            del: false,
            other_attrs: Vec::new(),
        };
        let row_matches = |candidate: &vsdx_parse::Row| match row {
            CellRow::Index(index) => candidate.index == Some(*index),
            CellRow::Name(name) => candidate.name.as_deref() == Some(name),
        };
        if let Some(section) = shape.children.iter_mut().find_map(|child| match child {
            ShapeChild::Section(section)
                if section.name == *section_name
                    && section.index.unwrap_or(0) == locator.section_index.unwrap_or(0) =>
            {
                Some(section)
            }
            _ => None,
        }) {
            if let Some(existing_row) = section.children.iter_mut().find_map(|child| match child {
                SectionChild::Row(candidate) if row_matches(candidate) => Some(candidate),
                _ => None,
            }) {
                existing_row.children.push(RowChild::Cell(cell));
            } else {
                section.children.push(SectionChild::Row(vsdx_parse::Row {
                    index: match row {
                        CellRow::Index(index) => Some(*index),
                        CellRow::Name(_) => None,
                    },
                    name: match row {
                        CellRow::Index(_) => None,
                        CellRow::Name(name) => Some(name.clone()),
                    },
                    local_name: None,
                    row_type: snapshot.row_type.clone(),
                    del: false,
                    children: vec![RowChild::Cell(cell)],
                    other_attrs: Vec::new(),
                }));
            }
        } else {
            shape
                .children
                .push(ShapeChild::Section(vsdx_parse::Section {
                    name: section_name.clone(),
                    index: locator.section_index,
                    del: false,
                    children: vec![SectionChild::Row(vsdx_parse::Row {
                        index: match row {
                            CellRow::Index(index) => Some(*index),
                            CellRow::Name(_) => None,
                        },
                        name: match row {
                            CellRow::Index(_) => None,
                            CellRow::Name(name) => Some(name.clone()),
                        },
                        local_name: None,
                        row_type: snapshot.row_type.clone(),
                        del: false,
                        children: vec![RowChild::Cell(cell)],
                        other_attrs: Vec::new(),
                    })],
                    other_attrs: Vec::new(),
                }));
        }
    }
}

#[allow(clippy::too_many_arguments)]
fn seed_shape(
    sheets: &MapRef,
    stories: &MapRef,
    txn: &mut TransactionMut<'_>,
    id: &str,
    page_id: &str,
    parent_id: Option<&str>,
    page_path: &str,
    lookup: &vsdx_parse::Sheet,
    shape: &vsdx_parse::Shape,
    resolver: &Resolver<'_>,
    resolved: &vsdx_resolve::ResolvedShape,
    depth: usize,
) -> EditResult<()> {
    if depth > MAX_SHAPE_NESTING {
        return Err(EditError::InvalidState(
            "shape nesting exceeds maximum depth".to_owned(),
        ));
    }
    let map = sheets.insert(txn, id, MapPrelim::default());
    map.insert(txn, "id", id);
    map.insert(txn, "pageId", page_id);
    map.insert(txn, "sourceId", shape.id as f64);
    map.insert(txn, "origin", "original");
    if let Some(parent_id) = parent_id {
        map.insert(txn, "parentId", parent_id);
    }
    if let Some(name) = &shape.name {
        map.insert(txn, "name", name.as_str());
    }
    let cells = map.insert(txn, "cells", MapPrelim::default());
    let child_order = map.insert(txn, "shapes", ArrayPrelim::default());
    for (name, value) in &resolved.cells {
        if let Lookup::Found(cell) = value {
            seed_cell(
                &cells,
                txn,
                &CellLocator {
                    sheet: CellSheet::Page(0),
                    shape_id: Some(shape.id),
                    section: None,
                    section_index: None,
                    row: None,
                    cell_name: name.clone(),
                },
                cell.cell.formula.as_deref(),
                cell.cell.value.as_deref(),
                None,
            );
        }
    }
    for section in resolved.sections.values() {
        for (row_key, resolved_row) in &section.rows {
            let row = if let Some(name) = row_key.strip_prefix("N:") {
                CellRow::Name(name.to_owned())
            } else if let Some(index) = row_key
                .strip_prefix("IX:")
                .and_then(|value| value.parse().ok())
            {
                CellRow::Index(index)
            } else {
                continue;
            };
            for (name, value) in &resolved_row.cells {
                if let Lookup::Found(cell) = value {
                    seed_cell(
                        &cells,
                        txn,
                        &CellLocator {
                            sheet: CellSheet::Page(0),
                            shape_id: Some(shape.id),
                            section: Some(section.name.clone()),
                            section_index: section.index,
                            row: Some(row.clone()),
                            cell_name: name.clone(),
                        },
                        cell.cell.formula.as_deref(),
                        cell.cell.value.as_deref(),
                        resolved_row.row_type.as_deref(),
                    );
                }
            }
        }
    }
    let text = resolver
        .resolve_text_in_context(shape, lookup, resolved)
        .map_err(|error| EditError::InvalidState(error.to_string()))?;
    stories.insert(
        txn,
        id,
        serde_json::to_string(&text).map_err(|error| EditError::Json(error.to_string()))?,
    );
    for child in shape.shapes() {
        let child_id = format!("{id}:shape:{}", child.id);
        child_order.push_back(txn, child_id.as_str());
        let child_resolved = resolver
            .resolve_shape(page_path, child.id)
            .map_err(|error| EditError::InvalidState(error.to_string()))?;
        seed_shape(
            sheets,
            stories,
            txn,
            &child_id,
            page_id,
            Some(id),
            page_path,
            lookup,
            child,
            resolver,
            &child_resolved,
            depth + 1,
        )?;
    }
    Ok(())
}

fn seed_cell(
    cells: &MapRef,
    txn: &mut TransactionMut<'_>,
    locator: &CellLocator,
    formula: Option<&str>,
    value: Option<&str>,
    row_type: Option<&str>,
) {
    let key = locator_key(locator);
    let cell = cells.insert(txn, key.as_str(), MapPrelim::default());
    cell.insert(txn, "name", locator.cell_name.as_str());
    if let Some(section) = &locator.section {
        cell.insert(txn, "section", section.as_str());
    }
    if let Some(index) = locator.section_index {
        cell.insert(txn, "sectionIndex", index as f64);
    }
    if let Some(row) = &locator.row {
        match row {
            CellRow::Index(index) => {
                cell.insert(txn, "rowIndex", *index as f64);
            }
            CellRow::Name(name) => {
                cell.insert(txn, "rowName", name.as_str());
            }
        }
    }
    if let Some(row_type) = row_type {
        cell.insert(txn, "rowType", row_type);
    }
    if let Some(formula) = formula {
        cell.insert(txn, "formula", formula);
        cell.insert(txn, "baselineFormula", formula);
    }
    if let Some(value) = value {
        cell.insert(txn, "value", value);
    }
}

impl DiagramSession {
    pub fn snapshot(&self) -> EditResult<DiagramSnapshot> {
        snapshot_doc(&self.doc)
    }

    pub fn semantic_cell_edits(&self) -> EditResult<Vec<vsdx_parse::SemanticCellEdit>> {
        semantic_cell_edits(&self.doc)
    }

    pub fn save(&self) -> EditResult<Vec<u8>> {
        serialize_doc(&self.doc)
    }

    pub fn set_cell_formula(
        &self,
        context: &EditCtx,
        page_id: &str,
        shape_id: &str,
        cell_name: &str,
        formula: impl Into<String>,
    ) -> EditResult<CellFormulaReceipt> {
        self.set_cell_formula_at(
            context,
            page_id,
            shape_id,
            CellLocator {
                sheet: CellSheet::Page(0),
                shape_id: None,
                section: None,
                section_index: None,
                row: None,
                cell_name: cell_name.to_owned(),
            },
            formula,
        )
    }

    pub fn set_cell_formula_at(
        &self,
        context: &EditCtx,
        page_id: &str,
        shape_id: &str,
        locator: CellLocator,
        formula: impl Into<String>,
    ) -> EditResult<CellFormulaReceipt> {
        let formula = formula.into();
        let mut txn = self.transact_for(context);
        let context_for_policy = CrdtMutationContext::new(&txn, page_id, shape_id)?;
        let target = match decide_mutation(
            &context_for_policy,
            context_for_policy.locator(locator.clone()),
            gesture_for_cell(&locator.cell_name),
            formula.clone(),
            &ParseLimits::default(),
        ) {
            MutationOutcome::Allowed { target, .. } => target,
            MutationOutcome::Refused { reason } | MutationOutcome::Unsupported { reason } => {
                return Err(EditError::InvalidState(reason));
            }
        };
        let cell = cell_map(&mut txn, page_id, shape_id, &target)?;
        let before = map_string(&cell, &txn, "formula");
        cell.insert(&mut txn, "formula", formula.as_str());
        Ok(CellFormulaReceipt {
            page_id: page_id.to_owned(),
            shape_id: shape_id.to_owned(),
            cell_name: target.cell_name,
            before,
            after: formula,
        })
    }

    pub fn move_shape(
        &self,
        context: &EditCtx,
        page_id: &str,
        shape_id: &str,
        x_formula: impl Into<String>,
        y_formula: impl Into<String>,
    ) -> EditResult<[CellFormulaReceipt; 2]> {
        self.set_cell_formula_pair(
            context,
            page_id,
            shape_id,
            ("PinX", x_formula.into(), MutationGesture::MoveX),
            ("PinY", y_formula.into(), MutationGesture::MoveY),
        )
    }

    pub fn resize_shape(
        &self,
        context: &EditCtx,
        page_id: &str,
        shape_id: &str,
        width_formula: impl Into<String>,
        height_formula: impl Into<String>,
    ) -> EditResult<[CellFormulaReceipt; 2]> {
        self.set_cell_formula_pair(
            context,
            page_id,
            shape_id,
            ("Width", width_formula.into(), MutationGesture::ResizeWidth),
            (
                "Height",
                height_formula.into(),
                MutationGesture::ResizeHeight,
            ),
        )
    }

    pub fn reorder_shape(
        &self,
        context: &EditCtx,
        page_id: &str,
        shape_id: &str,
        to_index: u32,
    ) -> EditResult<ShapeReceipt> {
        reorder(self, context, page_id, shape_id, to_index)
    }
    pub fn reorder_page(
        &self,
        context: &EditCtx,
        page_id: &str,
        to_index: u32,
    ) -> EditResult<ShapeReceipt> {
        reorder_page(self, context, page_id, to_index)
    }

    pub fn add_shape(
        &self,
        context: &EditCtx,
        page_id: &str,
        draft: &ShapeDraft,
    ) -> EditResult<ShapeReceipt> {
        validate_shape_draft(draft)?;
        let mut txn = self.transact_for(context);
        let pages = txn
            .get_map(PAGES)
            .ok_or_else(|| EditError::InvalidState("missing pages map".to_owned()))?;
        let page = map_ref(&pages, &txn, page_id)?;
        let order = map_array(&page, &txn, "shapes")?;
        let index = order.len(&txn);
        let sheets = txn
            .get_map(SHEETS)
            .ok_or_else(|| EditError::InvalidState("missing sheets map".to_owned()))?;
        let id_prefix = format!("{page_id}:shape:added:{}:", self.client_id);
        if sheets.len(&txn) as usize >= ParseLimits::default().max_shapes {
            return Err(EditError::InvalidState(
                "shape count exceeds maximum".to_owned(),
            ));
        }
        let source_bound = map_u32(&page, &txn, "maxSourceId")?
            .ok_or_else(|| EditError::InvalidState("missing page source ID bound".to_owned()))?;
        let allocated = materialized_source_ids(&sheets, &txn, &order, source_bound)?;
        let largest = allocated.values().copied().max().unwrap_or(source_bound);
        largest.checked_add(1).ok_or_else(|| {
            EditError::InvalidState("cannot allocate a materialized source ID".to_owned())
        })?;
        let sequence = txn.state_vector().get(&yrs::ClientID::new(self.client_id));
        let id = format!("{id_prefix}{sequence}");
        let shape = sheets.insert(&mut txn, id.as_str(), MapPrelim::default());
        shape.insert(&mut txn, "id", id.as_str());
        shape.insert(&mut txn, "pageId", page_id);
        shape.insert(&mut txn, "sourceId", 0.0);
        shape.insert(&mut txn, "origin", "added");
        if let Some(name) = &draft.name {
            shape.insert(&mut txn, "name", name.as_str());
        }
        let cells = shape.insert(&mut txn, "cells", MapPrelim::default());
        shape.insert(&mut txn, "shapes", ArrayPrelim::default());
        for cell in &draft.cells {
            seed_cell(
                &cells,
                &mut txn,
                &cell.locator,
                cell.formula.as_deref(),
                cell.value.as_deref(),
                cell.row_type.as_deref(),
            );
        }
        order.push_back(&mut txn, id.as_str());
        Ok(ShapeReceipt {
            page_id: page_id.to_owned(),
            shape_id: id,
            from_index: None,
            to_index: Some(index),
        })
    }

    /// Adds a 1D connector shape glued at both ends in one atomic transaction.
    pub fn add_connector(
        &self,
        context: &EditCtx,
        page_id: &str,
        draft: &ShapeDraft,
        from: &ConnectorGlue,
        to: &ConnectorGlue,
    ) -> EditResult<ShapeReceipt> {
        validate_shape_draft(draft)?;
        if !draft_is_one_d(draft) {
            return Err(EditError::InvalidState(
                "connector draft must describe a 1D shape".to_owned(),
            ));
        }
        let from_cell = normalized_glue_target(from.to_cell.as_deref()).ok_or_else(|| {
            EditError::InvalidState("connector glue needs a valid target cell".to_owned())
        })?;
        let to_cell = normalized_glue_target(to.to_cell.as_deref()).ok_or_else(|| {
            EditError::InvalidState("connector glue needs a valid target cell".to_owned())
        })?;
        for cell in [&from_cell, &to_cell] {
            if !glue_text_valid(cell) {
                return Err(EditError::InvalidState(
                    "connector glue contains invalid XML attribute text".to_owned(),
                ));
            }
        }
        let mut txn = self.transact_for(context);
        let pages = txn
            .get_map(PAGES)
            .ok_or_else(|| EditError::InvalidState("missing pages map".to_owned()))?;
        let page = map_ref(&pages, &txn, page_id)?;
        let order = map_array(&page, &txn, "shapes")?;
        let index = order.len(&txn);
        let sheets = txn
            .get_map(SHEETS)
            .ok_or_else(|| EditError::InvalidState("missing sheets map".to_owned()))?;
        for target in [&from.shape_id, &to.shape_id] {
            let shape = match sheets.get(&txn, target.as_str()) {
                Some(Out::YMap(shape)) => shape,
                _ => return Err(EditError::ShapeNotFound(target.clone())),
            };
            if map_string(&shape, &txn, "pageId").as_deref() != Some(page_id) {
                return Err(EditError::ShapeNotFound(target.clone()));
            }
        }
        for (target, cell) in [
            (&from.shape_id, from_cell.as_str()),
            (&to.shape_id, to_cell.as_str()),
        ] {
            if !connection_point_exists(&sheets, &txn, target.as_str(), cell) {
                return Err(EditError::InvalidState(
                    "connector glue references a missing connection point".to_owned(),
                ));
            }
        }
        let id_prefix = format!("{page_id}:shape:added:{}:", self.client_id);
        if sheets.len(&txn) as usize >= ParseLimits::default().max_shapes {
            return Err(EditError::InvalidState(
                "shape count exceeds maximum".to_owned(),
            ));
        }
        let source_bound = map_u32(&page, &txn, "maxSourceId")?
            .ok_or_else(|| EditError::InvalidState("missing page source ID bound".to_owned()))?;
        let allocated = materialized_source_ids(&sheets, &txn, &order, source_bound)?;
        let largest = allocated.values().copied().max().unwrap_or(source_bound);
        largest.checked_add(1).ok_or_else(|| {
            EditError::InvalidState("cannot allocate a materialized source ID".to_owned())
        })?;
        let sequence = txn.state_vector().get(&yrs::ClientID::new(self.client_id));
        let id = format!("{id_prefix}{sequence}");
        let shape = sheets.insert(&mut txn, id.as_str(), MapPrelim::default());
        shape.insert(&mut txn, "id", id.as_str());
        shape.insert(&mut txn, "pageId", page_id);
        shape.insert(&mut txn, "sourceId", 0.0);
        shape.insert(&mut txn, "origin", "added");
        if let Some(name) = &draft.name {
            shape.insert(&mut txn, "name", name.as_str());
        }
        let cells = shape.insert(&mut txn, "cells", MapPrelim::default());
        shape.insert(&mut txn, "shapes", ArrayPrelim::default());
        for cell in &draft.cells {
            seed_cell(
                &cells,
                &mut txn,
                &cell.locator,
                cell.formula.as_deref(),
                cell.value.as_deref(),
                cell.row_type.as_deref(),
            );
        }
        order.push_back(&mut txn, id.as_str());
        let connects = txn.get_or_insert_map(CONNECTS);
        for (endpoint, target, cell) in [
            (GlueEndpoint::Begin, &from.shape_id, from_cell),
            (GlueEndpoint::End, &to.shape_id, to_cell),
        ] {
            let key = glue_key(&id, endpoint);
            let entry = connects.insert(&mut txn, key.as_str(), MapPrelim::default());
            entry.insert(&mut txn, "id", key.as_str());
            entry.insert(&mut txn, "pageId", page_id);
            entry.insert(&mut txn, "connectorId", id.as_str());
            entry.insert(&mut txn, "endpoint", endpoint.name());
            entry.insert(&mut txn, "targetId", target.as_str());
            entry.insert(&mut txn, "toCell", cell.as_str());
        }
        Ok(ShapeReceipt {
            page_id: page_id.to_owned(),
            shape_id: id,
            from_index: None,
            to_index: Some(index),
        })
    }

    pub fn delete_shape(
        &self,
        context: &EditCtx,
        page_id: &str,
        shape_id: &str,
    ) -> EditResult<ShapeReceipt> {
        let mut txn = self.transact_for(context);
        let context_for_policy = CrdtMutationContext::new(&txn, page_id, shape_id)?;
        match decide_mutation(
            &context_for_policy,
            context_for_policy.locator(CellLocator {
                sheet: CellSheet::Page(0),
                shape_id: None,
                section: None,
                section_index: None,
                row: None,
                cell_name: "LockDelete".to_owned(),
            }),
            MutationGesture::Delete,
            String::new(),
            &ParseLimits::default(),
        ) {
            MutationOutcome::Allowed { .. } => {}
            MutationOutcome::Refused { reason } | MutationOutcome::Unsupported { reason } => {
                return Err(EditError::InvalidState(reason));
            }
        }
        let pages = txn
            .get_map(PAGES)
            .ok_or_else(|| EditError::InvalidState("missing pages map".to_owned()))?;
        let page = map_ref(&pages, &txn, page_id)?;
        let root_order = map_array(&page, &txn, "shapes")?;
        let sheets = txn
            .get_map(SHEETS)
            .ok_or_else(|| EditError::InvalidState("missing sheets map".to_owned()))?;
        let entries = shape_tree_entries(&sheets, &txn, &root_order)?;
        let target = entries
            .iter()
            .position(|entry| entry.id == shape_id)
            .ok_or_else(|| EditError::ShapeNotFound(shape_id.to_owned()))?;
        let order = entries[target].order.clone();
        let from = entries[target].index;
        let depth = entries[target].depth;
        let removed = entries
            .iter()
            .skip(target)
            .enumerate()
            .take_while(|(offset, entry)| *offset == 0 || entry.depth > depth)
            .map(|(_, entry)| entry)
            .map(|entry| entry.id.clone())
            .collect::<Vec<_>>();
        order.remove_range(&mut txn, from, 1);
        for id in &removed {
            sheets.remove(&mut txn, id.as_str());
        }
        if let Some(connects) = txn.get_map(CONNECTS) {
            let gone = removed.iter().map(String::as_str).collect::<HashSet<_>>();
            let mut doomed = Vec::new();
            for (key, value) in connects.iter(&txn) {
                let Out::YMap(entry) = value else {
                    continue;
                };
                let touches = map_string(&entry, &txn, "connectorId")
                    .as_deref()
                    .is_some_and(|id| gone.contains(id))
                    || map_string(&entry, &txn, "targetId")
                        .as_deref()
                        .is_some_and(|id| gone.contains(id));
                if touches {
                    doomed.push(key.to_owned());
                }
            }
            for key in doomed {
                connects.remove(&mut txn, key.as_str());
            }
        }
        Ok(ShapeReceipt {
            page_id: page_id.to_owned(),
            shape_id: shape_id.to_owned(),
            from_index: Some(from),
            to_index: None,
        })
    }

    fn set_cell_formula_pair(
        &self,
        context: &EditCtx,
        page_id: &str,
        shape_id: &str,
        first: (&str, String, MutationGesture),
        second: (&str, String, MutationGesture),
    ) -> EditResult<[CellFormulaReceipt; 2]> {
        let mut txn = self.transact_for(context);
        let context_for_policy = CrdtMutationContext::new(&txn, page_id, shape_id)?;
        let decide =
            |(name, formula, gesture): (&str, String, MutationGesture)| match decide_mutation(
                &context_for_policy,
                context_for_policy.locator(CellLocator {
                    sheet: CellSheet::Page(0),
                    shape_id: None,
                    section: None,
                    section_index: None,
                    row: None,
                    cell_name: name.to_owned(),
                }),
                gesture,
                formula.clone(),
                &ParseLimits::default(),
            ) {
                MutationOutcome::Allowed { target, .. } => Ok((target, formula)),
                MutationOutcome::Refused { reason } | MutationOutcome::Unsupported { reason } => {
                    Err(EditError::InvalidState(reason))
                }
            };
        let (first_target, first_formula) = decide(first)?;
        let (second_target, second_formula) = decide(second)?;
        let first_cell = cell_map(&mut txn, page_id, shape_id, &first_target)?;
        let second_cell = cell_map(&mut txn, page_id, shape_id, &second_target)?;
        let first_before = map_string(&first_cell, &txn, "formula");
        let second_before = map_string(&second_cell, &txn, "formula");
        first_cell.insert(&mut txn, "formula", first_formula.as_str());
        second_cell.insert(&mut txn, "formula", second_formula.as_str());
        Ok([
            CellFormulaReceipt {
                page_id: page_id.to_owned(),
                shape_id: shape_id.to_owned(),
                cell_name: first_target.cell_name,
                before: first_before,
                after: first_formula,
            },
            CellFormulaReceipt {
                page_id: page_id.to_owned(),
                shape_id: shape_id.to_owned(),
                cell_name: second_target.cell_name,
                before: second_before,
                after: second_formula,
            },
        ])
    }
}

fn validate_shape_draft(draft: &ShapeDraft) -> EditResult<()> {
    let limits = ParseLimits::default();
    let text = |value: &str| -> EditResult<()> {
        if value.len() > limits.max_attribute_bytes || value.chars().any(|c| !matches!(c, '\t' | '\r' | '\n' | '\u{20}'..='\u{d7ff}' | '\u{e000}'..='\u{fffd}' | '\u{10000}'..='\u{10ffff}')) {
            return Err(EditError::InvalidState("draft contains invalid XML attribute text".to_owned()));
        }
        Ok(())
    };
    if let Some(name) = &draft.name {
        text(name)?;
    }
    if draft.cells.len() > limits.max_cells {
        return Err(EditError::InvalidState(
            "draft cell count exceeds maximum".to_owned(),
        ));
    }
    let mut locators = HashSet::new();
    let mut row_types = std::collections::BTreeMap::new();
    for cell in &draft.cells {
        let locator = &cell.locator;
        if cell.value.is_some() {
            return Err(EditError::InvalidState(
                "shape draft cells must not contain value".to_owned(),
            ));
        }
        if locator.cell_name.is_empty()
            || cell.name != locator.cell_name
            || locator.section.is_some() != locator.row.is_some()
            || locator.section.is_none()
                && (locator.section_index.is_some() || cell.row_type.is_some())
        {
            return Err(EditError::InvalidState(
                "draft contains an invalid cell locator".to_owned(),
            ));
        }
        text(&locator.cell_name)?;
        if let Some(section) = &locator.section {
            text(section)?;
        }
        if let Some(CellRow::Name(row)) = &locator.row {
            text(row)?;
        }
        let key = locator_key(locator);
        if !locators.insert(key.clone()) {
            return Err(EditError::InvalidState(
                "draft contains duplicate cell locators".to_owned(),
            ));
        }
        if let Some(row_type) = &cell.row_type {
            text(row_type)?;
            let row_key = key
                .rsplit_once('\u{1f}')
                .map(|(row, _)| row)
                .unwrap_or(&key)
                .to_owned();
            if row_types
                .insert(row_key, row_type)
                .is_some_and(|previous| previous != row_type)
            {
                return Err(EditError::InvalidState(
                    "draft contains conflicting row types".to_owned(),
                ));
            }
        }
        if let Some(formula) = &cell.formula {
            text(formula)?;
            vsdx_eval::parse(formula.trim_start_matches('='), &limits)
                .map_err(|error| EditError::InvalidState(error.to_string()))?;
        }
    }
    Ok(())
}

pub(crate) fn validate_doc(doc: &Doc) -> EditResult<()> {
    validate_schema(doc)?;
    serializable_doc(doc)
}

fn validate_schema(doc: &Doc) -> EditResult<()> {
    let txn = doc.transact();
    let meta = required_map(&txn, META)?;
    if map_number(&meta, &txn, "schemaVersion") != Some(SCHEMA_VERSION) {
        return Err(EditError::InvalidState(
            "unsupported diagram schema version".to_owned(),
        ));
    }
    for key in ["fingerprint", "packageJson"] {
        if meta.get(&txn, key).is_none() {
            return Err(EditError::InvalidState(format!(
                "missing diagram metadata {key}"
            )));
        }
    }
    let order = required_array(&txn, PAGE_ORDER)?;
    for root in [PAGES, SHEETS, STORIES] {
        required_map(&txn, root)?;
    }
    let pages = required_map(&txn, PAGES)?;
    let sheets = required_map(&txn, SHEETS)?;
    validate_acyclic_parents(&sheets, &txn)?;
    for index in 0..order.len(&txn) {
        let page_id = array_string(&order, &txn, index)
            .ok_or_else(|| EditError::InvalidState("page order contains non-string".to_owned()))?;
        let page = map_ref(&pages, &txn, &page_id)?;
        if required_string(&page, &txn, "id")? != page_id {
            return Err(EditError::InvalidState(
                "page ID does not match map key".to_owned(),
            ));
        }
        required_string(&page, &txn, "sourcePartPath")?;
        if map_u32(&page, &txn, "maxSourceId")?.is_none() {
            return Err(EditError::InvalidState(
                "missing page source ID bound".to_owned(),
            ));
        }
        for shape_id in reachable_shape_ids(&sheets, &txn, &page)? {
            let shape = map_ref(&sheets, &txn, &shape_id)?;
            if required_string(&shape, &txn, "id")? != shape_id {
                return Err(EditError::InvalidState(
                    "shape ID does not match map key".to_owned(),
                ));
            }
            if map_u32(&shape, &txn, "sourceId")?.is_none() {
                return Err(EditError::InvalidState("missing source ID".to_owned()));
            }
            shape_origin(&shape, &txn)?;
            let cells = map_map(&shape, &txn, "cells")?;
            for (key, cell) in cells.iter(&txn) {
                let Out::YMap(cell) = cell else {
                    return Err(EditError::InvalidState("cell is not a map".to_owned()));
                };
                let locator = cell_locator(&cell, &txn, 0, 0)?;
                if locator_key(&locator) != key {
                    return Err(EditError::InvalidState(
                        "cell locator does not match map key".to_owned(),
                    ));
                }
                for field in ["formula", "value"] {
                    if cell.get(&txn, field).is_some() && map_string(&cell, &txn, field).is_none() {
                        return Err(EditError::InvalidState(format!(
                            "cell {field} is not a string"
                        )));
                    }
                }
            }
        }
    }
    let attached = required_map(&txn, PAGES)?.iter(&txn).try_fold(
        HashSet::new(),
        |mut attached, (_, page)| -> EditResult<_> {
            let Out::YMap(page) = page else {
                return Err(EditError::InvalidState("page is not a map".to_owned()));
            };
            for shape_id in reachable_shape_ids(&sheets, &txn, &page)? {
                if !attached.insert(shape_id) {
                    return Err(EditError::InvalidState(
                        "shape is attached to multiple pages".to_owned(),
                    ));
                }
            }
            Ok(attached)
        },
    )?;
    if attached.len() != sheets.len(&txn) as usize {
        return Err(EditError::InvalidState(
            "shape is not reachable from a page shape order".to_owned(),
        ));
    }
    validate_glue_records(&txn)?;
    Ok(())
}

fn validate_glue_records<T: ReadTxn>(txn: &T) -> EditResult<()> {
    let Some(connects) = txn.get_map(CONNECTS) else {
        return Ok(());
    };
    let pages = required_map(txn, PAGES)?;
    let sheets = required_map(txn, SHEETS)?;
    let mut endpoints = HashSet::new();
    for (key, value) in connects.iter(txn) {
        let Out::YMap(entry) = value else {
            return Err(EditError::InvalidState(
                "connector glue is not a map".to_owned(),
            ));
        };
        if map_string(&entry, txn, "id").as_deref() != Some(key) {
            return Err(EditError::InvalidState(
                "connector glue ID does not match map key".to_owned(),
            ));
        }
        let page_id = map_string(&entry, txn, "pageId").ok_or_else(|| {
            EditError::InvalidState("connector glue is missing its page".to_owned())
        })?;
        map_ref(&pages, txn, &page_id).map_err(|_| {
            EditError::InvalidState("connector glue references a missing page".to_owned())
        })?;
        let endpoint = map_string(&entry, txn, "endpoint")
            .as_deref()
            .and_then(GlueEndpoint::parse)
            .ok_or_else(|| EditError::InvalidState("connector glue has no endpoint".to_owned()))?;
        let to_cell = map_string(&entry, txn, "toCell").ok_or_else(|| {
            EditError::InvalidState("connector glue has no target cell".to_owned())
        })?;
        if !valid_glue_target(&to_cell) || !glue_text_valid(&to_cell) {
            return Err(EditError::InvalidState(
                "connector glue has an invalid target cell".to_owned(),
            ));
        }
        for field in ["connectorId", "targetId"] {
            let target = map_string(&entry, txn, field).ok_or_else(|| {
                EditError::InvalidState(format!("connector glue is missing {field}"))
            })?;
            let shape = match sheets.get(txn, target.as_str()) {
                Some(Out::YMap(shape)) => shape,
                _ => {
                    return Err(EditError::InvalidState(
                        "connector glue references a missing shape".to_owned(),
                    ));
                }
            };
            if map_string(&shape, txn, "pageId").as_deref() != Some(page_id.as_str()) {
                return Err(EditError::InvalidState(
                    "connector glue crosses pages".to_owned(),
                ));
            }
        }
        let pair = (
            map_string(&entry, txn, "connectorId").unwrap_or_default(),
            endpoint,
        );
        if !endpoints.insert(pair) {
            return Err(EditError::InvalidState(
                "connector glue duplicates an endpoint".to_owned(),
            ));
        }
    }
    Ok(())
}

/// Rejects a cyclic or self-referential shape `parentId` chain.
fn validate_acyclic_parents<T: ReadTxn>(sheets: &MapRef, txn: &T) -> EditResult<()> {
    for (shape_id, _) in sheets.iter(txn) {
        let mut current = shape_id.to_owned();
        let mut seen = std::collections::BTreeSet::new();
        loop {
            if !seen.insert(current.clone()) {
                return Err(EditError::InvalidState(format!(
                    "shape {shape_id} has a cyclic parent chain"
                )));
            }
            if seen.len() > MAX_SHAPE_NESTING {
                return Err(EditError::InvalidState(
                    "shape nesting exceeds maximum depth".to_owned(),
                ));
            }
            let Some(Out::YMap(parent_shape)) = sheets.get(txn, current.as_str()) else {
                break;
            };
            match map_string(&parent_shape, txn, "parentId") {
                Some(parent_id) => current = parent_id,
                None => break,
            }
        }
    }
    Ok(())
}

pub(crate) fn normalize_concurrent_orders(staged: &Doc) -> EditResult<()> {
    let mut txn = staged.transact_mut_with(crate::REMOTE_ORIGIN);
    let sheets = required_map(&txn, SHEETS)?;
    let mut orders = vec![(required_array(&txn, PAGE_ORDER)?, false)];
    for root in [PAGES, SHEETS] {
        for (_, value) in required_map(&txn, root)?.iter(&txn) {
            if let Out::YMap(owner) = value
                && let Some(Out::YArray(order)) = owner.get(&txn, "shapes")
            {
                orders.push((order, true));
            }
        }
    }
    for (order, shape_order) in orders {
        let mut seen = HashSet::new();
        for index in (0..order.len(&txn)).rev() {
            let Some(id) = array_string(&order, &txn, index) else {
                continue;
            };
            if (shape_order && sheets.get(&txn, &id).is_none()) || !seen.insert(id) {
                order.remove_range(&mut txn, index, 1);
            }
        }
    }
    Ok(())
}

/// Drops staged glue referencing a missing shape, keeping the rest.
pub(crate) fn prune_concurrent_glue(_before: &Doc, staged: &Doc) -> EditResult<()> {
    let mut txn = staged.transact_mut_with(crate::REMOTE_ORIGIN);
    let sheets = required_map(&txn, SHEETS)?;
    let Some(connects) = txn.get_map(CONNECTS) else {
        return Ok(());
    };
    let mut doomed = Vec::new();
    for (key, value) in connects.iter(&txn) {
        let Out::YMap(entry) = value else {
            continue;
        };
        let connector = map_string(&entry, &txn, "connectorId");
        let target = map_string(&entry, &txn, "targetId");
        let (Some(connector), Some(target)) = (connector, target) else {
            continue;
        };
        if sheets.get(&txn, connector.as_str()).is_none()
            || sheets.get(&txn, target.as_str()).is_none()
        {
            doomed.push(key.to_owned());
        }
    }
    for key in doomed {
        connects.remove(&mut txn, key.as_str());
    }
    Ok(())
}

pub(crate) fn validate_remote_update(before: &Doc, staged: &Doc) -> EditResult<()> {
    validate_schema(staged)?;
    validate_immutable_metadata(before, staged)?;
    validate_session_topology(before, staged)?;
    validate_formula_mutations(before, staged)?;
    serializable_doc(staged)?;
    let before_identities = shape_identities(before)?;
    let after_identities = shape_identities(staged)?;
    let removed_shapes = before_identities
        .keys()
        .filter(|key| !after_identities.contains_key(key.as_str()))
        .cloned()
        .collect::<std::collections::BTreeSet<_>>();
    authorize_removed_shapes(before, &before_identities, &removed_shapes)?;
    for (key, identity) in &before_identities {
        if let Some(after) = after_identities.get(key)
            && after != identity
        {
            return Err(EditError::InvalidState(format!(
                "remote update changes the identity of shape {key}"
            )));
        }
    }
    for (key, (_, _, origin, _)) in &after_identities {
        if !before_identities.contains_key(key) && origin.as_deref() != Some("added") {
            return Err(EditError::InvalidState(format!(
                "remote update adds shape {key} without session provenance"
            )));
        }
    }
    let before_baselines = baseline_formulas(before)?;
    let after_baselines = baseline_formulas(staged)?;
    for (key, before_formula) in &before_baselines {
        match after_baselines.get(key) {
            Some(after_formula) => {
                if after_formula != before_formula {
                    return Err(EditError::InvalidState(format!(
                        "remote update changes the save baseline of {key}"
                    )));
                }
            }
            None if removed_shapes.contains(shape_id_prefix(key)) => {}
            None => {
                return Err(EditError::InvalidState(format!(
                    "remote update removes the save baseline of {key} while its shape survives"
                )));
            }
        }
    }
    let before_protected = protected_formulas(before)?;
    let after_protected = protected_formulas(staged)?;
    for (key, formula) in before_protected {
        match after_protected.get(&key) {
            Some(after_formula) => {
                if after_formula != &formula {
                    return Err(EditError::InvalidState(format!(
                        "remote update changes protected cell {key}"
                    )));
                }
            }
            None if removed_shapes.contains(shape_id_prefix(&key)) => {}
            None => {
                return Err(EditError::InvalidState(format!(
                    "remote update removes protected cell {key} while its shape survives"
                )));
            }
        }
    }
    validate_new_cells(before, staged, &before_identities)?;
    validate_remote_glue(before, staged, &after_identities, &removed_shapes)?;
    Ok(())
}

type GlueIdentity = (String, String, String, String, String);

fn glue_identities(doc: &Doc) -> EditResult<std::collections::BTreeMap<String, GlueIdentity>> {
    let txn = doc.transact();
    let Some(connects) = txn.get_map(CONNECTS) else {
        return Ok(std::collections::BTreeMap::new());
    };
    let mut identities = std::collections::BTreeMap::new();
    for (key, value) in connects.iter(&txn) {
        let Out::YMap(entry) = value else {
            continue;
        };
        identities.insert(
            key.to_owned(),
            (
                map_string(&entry, &txn, "pageId").unwrap_or_default(),
                map_string(&entry, &txn, "connectorId").unwrap_or_default(),
                map_string(&entry, &txn, "endpoint").unwrap_or_default(),
                map_string(&entry, &txn, "targetId").unwrap_or_default(),
                map_string(&entry, &txn, "toCell").unwrap_or_default(),
            ),
        );
    }
    Ok(identities)
}

fn validate_remote_glue(
    before: &Doc,
    staged: &Doc,
    after_identities: &std::collections::BTreeMap<String, ShapeIdentity>,
    removed_shapes: &std::collections::BTreeSet<String>,
) -> EditResult<()> {
    let before_glue = glue_identities(before)?;
    let after_glue = glue_identities(staged)?;
    for (key, identity) in &before_glue {
        match after_glue.get(key) {
            Some(after) if after == identity => {}
            Some(_) => {
                return Err(EditError::InvalidState(format!(
                    "remote update changes connector glue {key}"
                )));
            }
            None => {
                let (_, connector, _, target, _) = identity;
                if !removed_shapes.contains(connector) && !removed_shapes.contains(target) {
                    return Err(EditError::InvalidState(format!(
                        "remote update removes connector glue {key} while its shapes survive"
                    )));
                }
            }
        }
    }
    for (key, (_, connector, _, _, _)) in after_glue
        .iter()
        .filter(|(key, _)| !before_glue.contains_key(*key))
    {
        let origin = after_identities
            .get(connector)
            .and_then(|identity| identity.2.as_deref());
        if origin != Some("added") {
            return Err(EditError::InvalidState(format!(
                "remote update adds connector glue {key} without a session connector"
            )));
        }
    }
    Ok(())
}

/// Extracts the shape ID preceding the first slash in a formula key.
fn shape_id_prefix(key: &str) -> &str {
    key.split_once('/').map_or(key, |(shape_id, _)| shape_id)
}

/// Checks LockDelete for remote deletions whose parent survives.
fn authorize_removed_shapes(
    before: &Doc,
    before_identities: &std::collections::BTreeMap<String, ShapeIdentity>,
    removed_shapes: &std::collections::BTreeSet<String>,
) -> EditResult<()> {
    for shape_id in removed_shapes {
        let parent_also_removed = before_identities
            .get(shape_id)
            .and_then(|identity| identity.1.as_ref())
            .is_some_and(|parent_id| removed_shapes.contains(parent_id));
        if parent_also_removed {
            continue;
        }
        let page_id = before_identities
            .get(shape_id)
            .and_then(|identity| identity.0.as_deref())
            .ok_or_else(|| {
                EditError::InvalidState(format!(
                    "remote update deletes shape {shape_id} with no page to authorize its removal"
                ))
            })?;
        authorize_shape_deletion(before, page_id, shape_id)?;
    }
    Ok(())
}

/// Rejects remote deletion when LockDelete is enabled or guarded.
fn authorize_shape_deletion(before: &Doc, page_id: &str, shape_id: &str) -> EditResult<()> {
    let txn = before.transact();
    let context = CrdtMutationContext::new(&txn, page_id, shape_id)?;
    match decide_mutation(
        &context,
        context.locator(CellLocator {
            sheet: CellSheet::Page(0),
            shape_id: None,
            section: None,
            section_index: None,
            row: None,
            cell_name: "LockDelete".to_owned(),
        }),
        MutationGesture::Delete,
        String::new(),
        &ParseLimits::default(),
    ) {
        MutationOutcome::Allowed { .. } => Ok(()),
        MutationOutcome::Refused { reason } | MutationOutcome::Unsupported { reason } => {
            Err(EditError::InvalidState(format!(
                "remote update deletes shape {shape_id} without authorization: {reason}"
            )))
        }
    }
}

/// Enforces local locks and rejects new GUARDs on cells added to existing shapes.
fn validate_new_cells(
    before: &Doc,
    staged: &Doc,
    before_identities: &std::collections::BTreeMap<String, ShapeIdentity>,
) -> EditResult<()> {
    let before_txn = before.transact();
    let staged_txn = staged.transact();
    let before_sheets = required_map(&before_txn, SHEETS)?;
    let staged_sheets = required_map(&staged_txn, SHEETS)?;
    for shape_id in before_identities.keys() {
        let Some(Out::YMap(before_shape)) = before_sheets.get(&before_txn, shape_id) else {
            continue;
        };
        let Some(Out::YMap(staged_shape)) = staged_sheets.get(&staged_txn, shape_id) else {
            continue;
        };
        let before_cells = map_map(&before_shape, &before_txn, "cells")?;
        let staged_cells = map_map(&staged_shape, &staged_txn, "cells")?;
        let before_values = before_cells
            .iter(&before_txn)
            .filter_map(|(name, cell)| match cell {
                Out::YMap(cell) => Some((
                    name.to_owned(),
                    map_string(&cell, &before_txn, "formula")
                        .or_else(|| map_string(&cell, &before_txn, "value")),
                )),
                _ => None,
            })
            .collect::<std::collections::BTreeMap<_, _>>();
        for (key, cell) in staged_cells.iter(&staged_txn) {
            if before_cells.get(&before_txn, key).is_some() {
                continue;
            }
            let Out::YMap(cell) = cell else { continue };
            let Some(formula) = map_string(&cell, &staged_txn, "formula") else {
                continue;
            };
            let name = map_string(&cell, &staged_txn, "name").unwrap_or_default();
            let locked = [
                "LockMoveX",
                "LockMoveY",
                "LockWidth",
                "LockHeight",
                "LockAspect",
                "LockTextEdit",
                "LockFormat",
                "LockDelete",
            ]
            .iter()
            .any(|lock| {
                lock_target(lock) == Some(name.as_str())
                    && before_values
                        .get(*lock)
                        .and_then(|value| value.as_deref())
                        .is_some_and(|value| lock_is_enabled(value, &before_values))
            });
            if locked {
                return Err(EditError::InvalidState(format!(
                    "remote update adds a lock-protected cell {shape_id}/{key}"
                )));
            }
            if is_guarded(&formula) {
                return Err(EditError::InvalidState(format!(
                    "remote update adds a guarded cell {shape_id}/{key}"
                )));
            }
        }
    }
    Ok(())
}

fn validate_session_topology(before: &Doc, staged: &Doc) -> EditResult<()> {
    let before_txn = before.transact();
    let staged_txn = staged.transact();
    let before_order = required_array(&before_txn, PAGE_ORDER)?;
    let staged_order = required_array(&staged_txn, PAGE_ORDER)?;
    if before_order != staged_order {
        return Err(EditError::InvalidState(
            "remote update changes immutable page topology".to_owned(),
        ));
    }
    let before_pages = required_map(&before_txn, PAGES)?;
    let staged_pages = required_map(&staged_txn, PAGES)?;
    if before_pages.len(&before_txn) != staged_pages.len(&staged_txn) {
        return Err(EditError::InvalidState(
            "remote update changes immutable page topology".to_owned(),
        ));
    }
    for (page_id, before_page) in before_pages.iter(&before_txn) {
        let Out::YMap(before_page) = before_page else {
            return Err(EditError::InvalidState("page is not a map".to_owned()));
        };
        let staged_page = map_ref(&staged_pages, &staged_txn, page_id)?;
        for key in ["id", "sourcePartPath", "maxSourceId"] {
            if before_page.get(&before_txn, key) != staged_page.get(&staged_txn, key) {
                return Err(EditError::InvalidState(format!(
                    "remote update changes immutable page {key}"
                )));
            }
        }
    }
    let before_sheets = required_map(&before_txn, SHEETS)?;
    let staged_sheets = required_map(&staged_txn, SHEETS)?;
    for (shape_id, staged_shape) in staged_sheets.iter(&staged_txn) {
        if before_sheets.get(&before_txn, shape_id).is_some() {
            continue;
        }
        let Out::YMap(staged_shape) = staged_shape else {
            return Err(EditError::InvalidState("shape is not a map".to_owned()));
        };
        if shape_origin(&staged_shape, &staged_txn)? != ShapeOrigin::Added {
            return Err(EditError::InvalidState(
                "remote update adds a shape with original identity".to_owned(),
            ));
        }
    }
    for (shape_id, before_shape) in before_sheets.iter(&before_txn) {
        let Out::YMap(before_shape) = before_shape else {
            continue;
        };
        let Some(Out::YMap(staged_shape)) = staged_sheets.get(&staged_txn, shape_id) else {
            continue;
        };
        for key in ["id", "pageId", "sourceId", "origin", "parentId"] {
            if before_shape.get(&before_txn, key) != staged_shape.get(&staged_txn, key) {
                return Err(EditError::InvalidState(format!(
                    "remote update changes immutable shape {key}"
                )));
            }
        }
        if before_shape.get(&before_txn, "shapes").is_some()
            && !matches!(
                staged_shape.get(&staged_txn, "shapes"),
                Some(Out::YArray(_))
            )
        {
            return Err(EditError::InvalidState(
                "remote update removes required shape order".to_owned(),
            ));
        }
        let before_cells = map_map(&before_shape, &before_txn, "cells")?;
        let staged_cells = map_map(&staged_shape, &staged_txn, "cells")?;
        for (cell_id, before_cell) in before_cells.iter(&before_txn) {
            let Out::YMap(before_cell) = before_cell else {
                continue;
            };
            let staged_cell = map_ref(&staged_cells, &staged_txn, cell_id)?;
            if before_cell.get(&before_txn, "rowType") != staged_cell.get(&staged_txn, "rowType") {
                return Err(EditError::InvalidState(
                    "remote update changes immutable geometry row type".to_owned(),
                ));
            }
            if before_cell.get(&before_txn, "value") != staged_cell.get(&staged_txn, "value") {
                return Err(EditError::InvalidState(
                    "remote update changes untrusted cached cell value".to_owned(),
                ));
            }
            if before_cell.get(&before_txn, "baselineFormula")
                != staged_cell.get(&staged_txn, "baselineFormula")
            {
                return Err(EditError::InvalidState(
                    "remote update changes immutable cell baseline formula".to_owned(),
                ));
            }
        }
    }
    Ok(())
}

fn validate_immutable_metadata(before: &Doc, staged: &Doc) -> EditResult<()> {
    let before_txn = before.transact();
    let staged_txn = staged.transact();
    let before_meta = required_map(&before_txn, META)?;
    let staged_meta = required_map(&staged_txn, META)?;
    for key in ["fingerprint", "packageJson", "packageBytes"] {
        if before_meta.get(&before_txn, key) != staged_meta.get(&staged_txn, key) {
            return Err(EditError::InvalidState(format!(
                "remote update changes immutable diagram metadata {key}"
            )));
        }
    }
    Ok(())
}

fn validate_formula_mutations(before: &Doc, staged: &Doc) -> EditResult<()> {
    let before_txn = before.transact();
    let staged_txn = staged.transact();
    let pages = required_map(&before_txn, PAGES)?;
    let sheets = required_map(&before_txn, SHEETS)?;
    let staged_sheets = required_map(&staged_txn, SHEETS)?;
    for (page_id, page) in pages.iter(&before_txn) {
        let Out::YMap(page) = page else { continue };
        for shape_id in reachable_shape_ids(&sheets, &before_txn, &page)? {
            let shape = map_ref(&sheets, &before_txn, &shape_id)?;
            let cells = map_map(&shape, &before_txn, "cells")?;
            let context = CrdtMutationContext::new(&before_txn, page_id, &shape_id)?;
            let Some(Out::YMap(staged_shape)) = staged_sheets.get(&staged_txn, &shape_id) else {
                continue;
            };
            let staged_cells = map_map(&staged_shape, &staged_txn, "cells")?;
            for (key, cell) in cells.iter(&before_txn) {
                let Out::YMap(cell) = cell else { continue };
                let before_formula = map_string(&cell, &before_txn, "formula");
                let after_formula =
                    staged_cells
                        .get(&staged_txn, key)
                        .and_then(|cell| match cell {
                            Out::YMap(cell) => map_string(&cell, &staged_txn, "formula"),
                            _ => None,
                        });
                if before_formula == after_formula {
                    continue;
                }
                let formula = after_formula.ok_or_else(|| {
                    EditError::InvalidState(format!(
                        "remote update removes formula from {page_id}/{shape_id}/{key}"
                    ))
                })?;
                let locator = context.locator(cell_locator(&cell, &before_txn, 0, 0)?);
                match decide_mutation(
                    &context,
                    locator.clone(),
                    gesture_for_cell(&locator.cell_name),
                    formula,
                    &ParseLimits::default(),
                ) {
                    MutationOutcome::Allowed { target, .. } if target == locator => {}
                    MutationOutcome::Allowed { .. } => {
                        return Err(EditError::InvalidState(format!(
                            "remote update bypasses formula redirect at {page_id}/{shape_id}/{key}"
                        )));
                    }
                    MutationOutcome::Refused { reason }
                    | MutationOutcome::Unsupported { reason } => {
                        return Err(EditError::InvalidState(reason));
                    }
                }
            }
        }
    }
    for (shape_id, staged_shape) in staged_sheets.iter(&staged_txn) {
        let Out::YMap(staged_shape) = staged_shape else {
            return Err(EditError::InvalidState("shape is not a map".to_owned()));
        };
        let staged_cells = map_map(&staged_shape, &staged_txn, "cells")?;
        let before_cells = sheets
            .get(&before_txn, shape_id)
            .and_then(|shape| match shape {
                Out::YMap(shape) => map_map(&shape, &before_txn, "cells").ok(),
                _ => None,
            });
        for (key, cell) in staged_cells.iter(&staged_txn) {
            if before_cells
                .as_ref()
                .is_some_and(|cells| cells.get(&before_txn, key).is_some())
            {
                continue;
            }
            let Out::YMap(cell) = cell else {
                return Err(EditError::InvalidState("cell is not a map".to_owned()));
            };
            if cell.get(&staged_txn, "value").is_some() {
                return Err(EditError::InvalidState(
                    "remote update adds untrusted cached cell value".to_owned(),
                ));
            }
        }
    }
    Ok(())
}

type ShapeIdentity = (Option<String>, Option<String>, Option<String>, Option<f64>);

/// The save path routes a sheet by this metadata, so only the seed and a local add may write it.
fn shape_identities(doc: &Doc) -> EditResult<std::collections::BTreeMap<String, ShapeIdentity>> {
    let txn = doc.transact();
    let sheets = required_map(&txn, SHEETS)?;
    let mut identities = std::collections::BTreeMap::new();
    for (shape_id, shape) in sheets.iter(&txn) {
        let Out::YMap(shape) = shape else { continue };
        identities.insert(
            shape_id.to_owned(),
            (
                map_string(&shape, &txn, "pageId"),
                map_string(&shape, &txn, "parentId"),
                map_string(&shape, &txn, "origin"),
                map_number(&shape, &txn, "sourceId"),
            ),
        );
    }
    Ok(identities)
}

/// Baselines come from the package seed alone; a peer that could write one could hide an edit.
fn baseline_formulas(doc: &Doc) -> EditResult<std::collections::BTreeMap<String, String>> {
    let txn = doc.transact();
    let sheets = required_map(&txn, SHEETS)?;
    let mut baselines = std::collections::BTreeMap::new();
    for (shape_id, shape) in sheets.iter(&txn) {
        let Out::YMap(shape) = shape else { continue };
        let cells = map_map(&shape, &txn, "cells")?;
        for (key, cell) in cells.iter(&txn) {
            let Out::YMap(cell) = cell else { continue };
            if let Some(baseline) = map_string(&cell, &txn, "baselineFormula") {
                baselines.insert(format!("{shape_id}/{key}"), baseline);
            }
        }
    }
    Ok(baselines)
}

fn protected_formulas(doc: &Doc) -> EditResult<std::collections::BTreeMap<String, String>> {
    let txn = doc.transact();
    let sheets = required_map(&txn, SHEETS)?;
    let mut protected = std::collections::BTreeMap::new();
    for (shape_id, shape) in sheets.iter(&txn) {
        let Out::YMap(shape) = shape else { continue };
        let cells = map_map(&shape, &txn, "cells")?;
        let values = cells
            .iter(&txn)
            .filter_map(|(name, cell)| match cell {
                Out::YMap(cell) => Some((
                    name.to_string(),
                    map_string(&cell, &txn, "formula").or_else(|| map_string(&cell, &txn, "value")),
                )),
                _ => None,
            })
            .collect::<std::collections::BTreeMap<_, _>>();
        for (name, cell) in cells.iter(&txn) {
            let Out::YMap(cell) = cell else { continue };
            let formula = map_string(&cell, &txn, "formula");
            let locked = lock_target(name).is_some_and(|_| {
                values
                    .get(name)
                    .and_then(|value| value.as_deref())
                    .is_some_and(|value| lock_is_enabled(value, &values))
            });
            let protected_target = lock_target(name).is_none()
                && [
                    "LockMoveX",
                    "LockMoveY",
                    "LockWidth",
                    "LockHeight",
                    "LockAspect",
                    "LockTextEdit",
                    "LockFormat",
                    "LockDelete",
                ]
                .iter()
                .any(|lock| {
                    lock_target(lock) == Some(name)
                        && values
                            .get(*lock)
                            .and_then(|value| value.as_deref())
                            .is_some_and(|value| lock_is_enabled(value, &values))
                });
            if locked || protected_target || formula.as_deref().is_some_and(is_guarded) {
                protected.insert(format!("{shape_id}/{name}"), formula.unwrap_or_default());
            }
        }
    }
    Ok(protected)
}

fn lock_is_enabled(
    value: &str,
    formulas: &std::collections::BTreeMap<String, Option<String>>,
) -> bool {
    let formulas = formulas
        .iter()
        .filter_map(|(name, formula)| {
            formula
                .as_ref()
                .map(|formula| (name.clone(), formula.clone()))
        })
        .collect::<std::collections::BTreeMap<_, _>>();
    matches!(
        evaluate(value.trim_start_matches('='), &formulas, &ParseLimits::default()),
        vsdx_eval::Evaluation::Evaluated(result)
            if matches!(result.value, vsdx_eval::Value::Number(number) if number.number == 1.0)
    )
}

fn lock_target(lock: &str) -> Option<&str> {
    match lock {
        "LockMoveX" => Some("PinX"),
        "LockMoveY" => Some("PinY"),
        "LockWidth" => Some("Width"),
        "LockHeight" => Some("Height"),
        "LockAspect" => Some("Width"),
        "LockTextEdit" => Some("Text"),
        "LockFormat" | "LockDelete" => None,
        _ => None,
    }
}

fn is_guarded(formula: &str) -> bool {
    vsdx_eval::parse(formula.trim_start_matches('='), &ParseLimits::default())
        .map(|expression| {
            format!("{expression:?}")
                .to_ascii_uppercase()
                .contains("GUARD")
        })
        .unwrap_or(false)
}

fn reachable_shape_ids<T: ReadTxn>(
    sheets: &MapRef,
    txn: &T,
    page: &MapRef,
) -> EditResult<Vec<String>> {
    let roots = map_array(page, txn, "shapes")?;
    let page_id = required_string(page, txn, "id")?;
    let entries = shape_tree_entries(sheets, txn, &roots)?;
    for entry in &entries {
        let shape_id = &entry.id;
        let parent_id = entry.parent_id.as_deref();
        let shape = map_ref(sheets, txn, shape_id)?;
        let stored_parent_id = shape.get(txn, "parentId");
        if stored_parent_id.is_some() && map_string(&shape, txn, "parentId").is_none() {
            return Err(EditError::InvalidState(
                "shape parentId is not a string".to_owned(),
            ));
        }
        if map_string(&shape, txn, "parentId").as_deref() != parent_id {
            return Err(EditError::InvalidState(
                "shape parentId does not match shape order".to_owned(),
            ));
        }
        if shape.get(txn, "pageId").is_some()
            && map_string(&shape, txn, "pageId").as_deref() != Some(page_id.as_str())
        {
            return Err(EditError::InvalidState(
                "shape pageId does not match page shape order".to_owned(),
            ));
        }
    }
    Ok(entries.into_iter().map(|entry| entry.id).collect())
}

struct ShapeTreeEntry {
    id: String,
    parent_id: Option<String>,
    order: ArrayRef,
    index: u32,
    depth: usize,
}

fn shape_tree_entries<T: ReadTxn>(
    sheets: &MapRef,
    txn: &T,
    roots: &ArrayRef,
) -> EditResult<Vec<ShapeTreeEntry>> {
    let mut result = Vec::new();
    let mut seen = HashSet::new();
    let mut pending = (0..roots.len(txn))
        .rev()
        .map(|index| (roots.clone(), index, None, 1))
        .collect::<Vec<_>>();
    while let Some((order, index, parent_id, depth)) = pending.pop() {
        if depth > MAX_SHAPE_NESTING {
            return Err(EditError::InvalidState(
                "shape nesting exceeds maximum depth".to_owned(),
            ));
        }
        let id = array_string(&order, txn, index)
            .ok_or_else(|| EditError::InvalidState("shape order contains non-string".to_owned()))?;
        if !seen.insert(id.clone()) {
            return Err(EditError::InvalidState(
                "shape order contains a duplicate or cycle".to_owned(),
            ));
        }
        let shape = map_ref(sheets, txn, &id)?;
        if let Some(Out::YArray(children)) = shape.get(txn, "shapes") {
            pending.extend(
                (0..children.len(txn)).rev().map(|child_index| {
                    (children.clone(), child_index, Some(id.clone()), depth + 1)
                }),
            );
        }
        result.push(ShapeTreeEntry {
            id,
            parent_id,
            order,
            index,
            depth,
        });
    }
    Ok(result)
}

fn snapshot_doc(doc: &Doc) -> EditResult<DiagramSnapshot> {
    let txn = doc.transact();
    let order = required_array(&txn, PAGE_ORDER)?;
    let pages = required_map(&txn, PAGES)?;
    let sheets = required_map(&txn, SHEETS)?;
    let mut result = Vec::new();
    validate_acyclic_parents(&sheets, &txn)?;
    for index in 0..order.len(&txn) {
        let id = array_string(&order, &txn, index)
            .ok_or_else(|| EditError::InvalidState("page order contains non-string".to_owned()))?;
        let page = map_ref(&pages, &txn, &id)?;
        let shape_order = map_array(&page, &txn, "shapes")?;
        let source_ids = materialized_source_ids(
            &sheets,
            &txn,
            &shape_order,
            map_number(&page, &txn, "maxSourceId").unwrap_or_default() as u32,
        )?;
        let mut shapes = Vec::new();
        for shape_index in 0..shape_order.len(&txn) {
            let shape_id = array_string(&shape_order, &txn, shape_index).ok_or_else(|| {
                EditError::InvalidState("shape order contains non-string".to_owned())
            })?;
            shapes.push(snapshot_shape(
                &sheets,
                &txn,
                &shape_id,
                id.strip_prefix("page:")
                    .and_then(|value| value.parse().ok())
                    .unwrap_or_default(),
                &source_ids,
                1,
            )?);
        }
        result.push(PageSnapshot {
            id,
            source_part_path: required_string(&page, &txn, "sourcePartPath")?,
            name: map_string(&page, &txn, "name"),
            shapes,
        });
    }
    Ok(DiagramSnapshot { pages: result })
}

fn snapshot_shape<T: ReadTxn>(
    sheets: &MapRef,
    txn: &T,
    shape_id: &str,
    page_id: u32,
    source_ids: &std::collections::BTreeMap<String, u32>,
    depth: usize,
) -> EditResult<ShapeSnapshot> {
    if depth > MAX_SHAPE_NESTING {
        return Err(EditError::InvalidState(
            "shape nesting exceeds maximum depth".to_owned(),
        ));
    }
    let shape = map_ref(sheets, txn, shape_id)?;
    let stored_source_id = map_number(&shape, txn, "sourceId")
        .ok_or_else(|| EditError::InvalidState("missing source ID".to_owned()))?
        as u32;
    let source_id = source_ids
        .get(shape_id)
        .copied()
        .unwrap_or(stored_source_id);
    let cells = map_map(&shape, txn, "cells")?;
    let references = local_references(&cells, txn)?;
    let evaluate_locally = |formula: &str| evaluate_cached_formula(formula, &references);
    let mut snapshots = Vec::new();
    for (_key, value) in cells.iter(txn) {
        let Out::YMap(cell) = value else {
            return Err(EditError::InvalidState("cell is not a map".to_owned()));
        };
        let locator = cell_locator(&cell, txn, page_id, source_id)?;
        let formula = map_string(&cell, txn, "formula");
        let baseline = map_string(&cell, txn, "baselineFormula");
        let value = if formula == baseline {
            formula
                .as_deref()
                .and_then(evaluate_locally)
                .or_else(|| map_string(&cell, txn, "value"))
        } else {
            formula.as_deref().and_then(evaluate_locally)
        };
        snapshots.push(CellSnapshot {
            row_type: map_string(&cell, txn, "rowType"),
            name: locator.cell_name.clone(),
            locator,
            formula,
            value,
        });
    }
    snapshots.sort_by_key(|cell| {
        let row = match &cell.locator.row {
            Some(CellRow::Index(index)) => (Some(*index), None),
            Some(CellRow::Name(name)) => (None, Some(name.clone())),
            None => (None, None),
        };
        (
            cell.locator.section.clone(),
            cell.locator.section_index.unwrap_or(0),
            row,
            cell.name.clone(),
        )
    });
    let Some(Out::YArray(child_order)) = shape.get(txn, "shapes") else {
        return Ok(ShapeSnapshot {
            id: shape_id.to_owned(),
            source_id,
            name: map_string(&shape, txn, "name"),
            cells: snapshots,
            children: Vec::new(),
        });
    };
    let mut children = Vec::with_capacity(child_order.len(txn) as usize);
    for index in 0..child_order.len(txn) {
        let child_id = array_string(&child_order, txn, index)
            .ok_or_else(|| EditError::InvalidState("shape order contains non-string".to_owned()))?;
        children.push(snapshot_shape(
            sheets,
            txn,
            &child_id,
            page_id,
            source_ids,
            depth + 1,
        )?);
    }
    Ok(ShapeSnapshot {
        id: shape_id.to_owned(),
        source_id,
        name: map_string(&shape, txn, "name"),
        cells: snapshots,
        children,
    })
}

fn local_references<T: ReadTxn>(
    cells: &MapRef,
    txn: &T,
) -> EditResult<vsdx_resolve::ResolvedShape> {
    let mut references = vsdx_resolve::ResolvedShape::default();
    for (_, value) in cells.iter(txn) {
        let Out::YMap(cell) = value else {
            return Err(EditError::InvalidState("cell is not a map".to_owned()));
        };
        let locator = cell_locator(&cell, txn, 0, 0)?;
        let value = Lookup::Found(vsdx_resolve::ResolvedCell {
            cell: Cell {
                name: locator.cell_name.clone(),
                formula: map_string(&cell, txn, "formula"),
                value: map_string(&cell, txn, "value"),
                unit: None,
                del: false,
                other_attrs: Vec::new(),
            },
            provenance: vsdx_resolve::Provenance::Local,
        });
        match (&locator.section, &locator.row) {
            (Some(name), Some(row)) => {
                let section = references
                    .sections
                    .entry(vsdx_resolve::section_key(name, locator.section_index))
                    .or_insert_with(|| vsdx_resolve::ResolvedSection {
                        name: name.clone(),
                        index: locator.section_index,
                        ..Default::default()
                    });
                let key = match row {
                    CellRow::Index(index) => format!("IX:{index}"),
                    CellRow::Name(name) => format!("N:{name}"),
                };
                section
                    .rows
                    .entry(key.clone())
                    .or_insert_with(|| vsdx_resolve::ResolvedRow {
                        key,
                        ..Default::default()
                    })
                    .cells
                    .insert(locator.cell_name, value);
            }
            (None, None) => {
                references.cells.insert(locator.cell_name, value);
            }
            _ => {}
        }
    }
    Ok(references)
}

fn evaluate_cached_formula(
    formula: &str,
    references: &vsdx_resolve::ResolvedShape,
) -> Option<String> {
    match evaluate(
        formula.trim_start_matches('='),
        references,
        &ParseLimits::default(),
    ) {
        vsdx_eval::Evaluation::Evaluated(result) => match result.value {
            vsdx_eval::Value::Number(number) => Some(number.number.to_string()),
            vsdx_eval::Value::Color(_) => None,
        },
        _ => None,
    }
}

fn materialized_source_ids<T: ReadTxn>(
    sheets: &MapRef,
    txn: &T,
    roots: &ArrayRef,
    mut largest: u32,
) -> EditResult<std::collections::BTreeMap<String, u32>> {
    let entries = shape_tree_entries(sheets, txn, roots)?;
    let mut added = Vec::new();
    for entry in entries {
        let shape = map_ref(sheets, txn, &entry.id)?;
        let source_id = map_number(&shape, txn, "sourceId")
            .ok_or_else(|| EditError::InvalidState("missing source ID".to_owned()))?
            as u32;
        if shape_origin(&shape, txn)? == ShapeOrigin::Original {
            largest = largest.max(source_id);
        } else {
            added.push(entry.id);
        }
    }
    added.sort();
    let mut result = std::collections::BTreeMap::new();
    for id in added {
        largest = largest.checked_add(1).ok_or_else(|| {
            EditError::InvalidState("cannot allocate a materialized source ID".to_owned())
        })?;
        result.insert(id, largest);
    }
    Ok(result)
}

fn semantic_cell_edits(doc: &Doc) -> EditResult<Vec<vsdx_parse::SemanticCellEdit>> {
    let txn = doc.transact();
    let pages = required_map(&txn, PAGES)?;
    let sheets = required_map(&txn, SHEETS)?;
    let mut edits = Vec::new();
    for (page_id, page) in pages.iter(&txn) {
        let Out::YMap(_page) = page else { continue };
        let source_page_id = page_id
            .strip_prefix("page:")
            .and_then(|value| value.parse().ok())
            .ok_or_else(|| EditError::InvalidState("invalid page ID".to_owned()))?;
        for (_shape_id, shape) in sheets.iter(&txn) {
            let Out::YMap(shape) = shape else { continue };
            if map_string(&shape, &txn, "pageId").as_deref() == Some(page_id) {
                let source_id = map_number(&shape, &txn, "sourceId")
                    .ok_or_else(|| EditError::InvalidState("missing source ID".to_owned()))?
                    as u32;
                if shape_origin(&shape, &txn)? != ShapeOrigin::Original {
                    continue;
                }
                let cells = map_map(&shape, &txn, "cells")?;
                let references = local_references(&cells, &txn)?;
                for (_key, value) in cells.iter(&txn) {
                    let Out::YMap(cell) = value else { continue };
                    let formula = map_string(&cell, &txn, "formula");
                    let baseline = map_string(&cell, &txn, "baselineFormula");
                    if formula == baseline {
                        continue;
                    }
                    let Some(formula) = formula else { continue };
                    let locator = cell_locator(&cell, &txn, source_page_id, source_id)?;
                    let value = evaluate_cached_formula(&formula, &references);
                    edits.push(vsdx_parse::SemanticCellEdit {
                        locator: locator.clone(),
                        gesture: gesture_for_cell(&locator.cell_name),
                        formula: Some(formula),
                        value,
                    });
                }
            }
        }
    }
    Ok(edits)
}

fn serialize_doc(doc: &Doc) -> EditResult<Vec<u8>> {
    let package = original_package_from_doc(doc)?;
    let snapshot = snapshot_doc(doc)?;
    let structural = structural_edits(&package, doc, &snapshot)?;
    let bytes = if structural.is_empty() {
        return vsdx_parse::save_semantic_cell_edits(&package, &semantic_cell_edits(doc)?)
            .map_err(|error| EditError::Parse(error.to_string()));
    } else {
        vsdx_parse::save_structural_edits(&package, &structural)
            .map_err(|error| EditError::Parse(error.to_string()))?
    };
    let package =
        vsdx_parse::parse_vsdx(&bytes).map_err(|error| EditError::Parse(error.to_string()))?;
    vsdx_parse::save_semantic_cell_edits(&package, &semantic_cell_edits(doc)?)
        .map_err(|error| EditError::Parse(error.to_string()))
}

fn serializable_doc(doc: &Doc) -> EditResult<()> {
    #[cfg(test)]
    {
        let txn = doc.transact();
        let meta = required_map(&txn, META)?;
        if meta.get(&txn, "packageBytes").is_none()
            && matches!(meta.get(&txn, "packageJson"), Some(Out::Any(Any::Buffer(bytes))) if bytes.is_empty())
        {
            return Ok(());
        }
    }
    serialize_doc(doc).map(|_| ())
}

fn structural_edits(
    package: &vsdx_parse::VsdxPackage,
    doc: &Doc,
    snapshot: &DiagramSnapshot,
) -> EditResult<Vec<StructuralEdit>> {
    let txn = doc.transact();
    let pages = required_map(&txn, PAGES)?;
    let sheets = required_map(&txn, SHEETS)?;
    let glue = glue_records(&txn)?;
    let mut edits = Vec::new();
    let desired_page_ids =
        page_part_paths_for_snapshot(package, snapshot)?
            .iter()
            .map(|path| {
                package.page_part_ids.get(path).copied().ok_or_else(|| {
                    EditError::InvalidState("page is missing a source ID".to_owned())
                })
            })
            .collect::<EditResult<Vec<_>>>()?;
    let original_page_ids =
        package
            .page_part_paths
            .iter()
            .map(|path| {
                package.page_part_ids.get(path).copied().ok_or_else(|| {
                    EditError::InvalidState("page is missing a source ID".to_owned())
                })
            })
            .collect::<EditResult<Vec<_>>>()?;
    if desired_page_ids != original_page_ids {
        edits.push(StructuralEdit::ReorderPages {
            page_ids: desired_page_ids,
        });
    }
    for path in &package.page_part_paths {
        let Some(page_id) = package.page_part_ids.get(path) else {
            continue;
        };
        let Some(sheet) = package.page_contents.get(path) else {
            continue;
        };
        let Some(page) = snapshot
            .pages
            .iter()
            .find(|page| page.source_part_path == *path)
        else {
            continue;
        };
        let session_page = map_ref(&pages, &txn, &page.id)?;
        let roots = map_array(&session_page, &txn, "shapes")?;
        let snapshots = snapshot_shapes(page);
        let desired = structural_shape_entries(&sheets, &txn, &roots, &snapshots)?;
        let originals = original_shape_containers(sheet);
        let mut original_identities = HashSet::new();
        for shape in desired.iter().filter(|shape| shape.is_original) {
            let parent = shape.parent_id.as_ref().and_then(|parent_id| {
                desired
                    .iter()
                    .find(|candidate| candidate.id == *parent_id)
                    .filter(|parent| parent.is_original)
                    .map(|parent| parent.source_id)
            });
            if !originals
                .get(&parent)
                .is_some_and(|ids| ids.contains(&shape.source_id))
                || !original_identities.insert((parent, shape.source_id))
            {
                return Err(EditError::InvalidState(
                    "original shape identity does not match the source package".to_owned(),
                ));
            }
        }
        let mut added = std::collections::BTreeMap::new();
        for shape in desired.iter().filter(|shape| !shape.is_original) {
            if shape.parent_id.is_some() {
                return Err(EditError::InvalidState(
                    "added shapes cannot have a parent".to_owned(),
                ));
            }
            let snapshot = snapshots.get(shape.id.as_str()).ok_or_else(|| {
                EditError::InvalidState(format!(
                    "added shape {:?} is missing from snapshot",
                    shape.id
                ))
            })?;
            added.insert(shape.id.as_str(), shape.source_id);
            edits.push(StructuralEdit::AddShape {
                page_id: *page_id,
                shape_xml: shape_xml(snapshot).into_bytes(),
            });
        }
        let by_id = desired
            .iter()
            .map(|shape| (shape.id.as_str(), shape))
            .collect::<std::collections::BTreeMap<_, _>>();
        let originals_by_source = desired
            .iter()
            .filter(|shape| shape.is_original)
            .map(|shape| (shape.source_id, shape))
            .collect::<std::collections::BTreeMap<_, _>>();
        for (parent, original_order) in &originals {
            if let Some(parent) = parent {
                let parent = originals_by_source.get(parent);
                if !parent.is_some_and(|shape| shape.is_original) {
                    continue;
                }
            }
            let desired = desired
                .iter()
                .filter(|shape| match (&shape.parent_id, parent) {
                    (None, None) => true,
                    (Some(session_parent), Some(original_parent)) => {
                        by_id.get(session_parent.as_str()).is_some_and(|parent| {
                            parent.is_original && parent.source_id == *original_parent
                        })
                    }
                    _ => false,
                })
                .collect::<Vec<_>>();
            structural_container_edits(*page_id, original_order, &desired, &added, &mut edits)?;
        }
        let sources = desired
            .iter()
            .map(|shape| (shape.id.as_str(), shape.source_id))
            .collect::<std::collections::BTreeMap<_, _>>();
        for record in glue.iter().filter(|record| record.page_id == page.id) {
            let (Some(connector), Some(target)) = (
                sources.get(record.connector_id.as_str()),
                sources.get(record.target_id.as_str()),
            ) else {
                return Err(EditError::InvalidState(
                    "connector glue references a missing shape".to_owned(),
                ));
            };
            edits.push(StructuralEdit::AddConnect {
                page_id: *page_id,
                from_sheet: *connector,
                from_cell: record.endpoint.endpoint_cell().to_owned(),
                to_sheet: *target,
                to_cell: record.to_cell.clone(),
            });
        }
    }
    Ok(edits)
}

struct StructuralShapeEntry {
    id: String,
    parent_id: Option<String>,
    source_id: u32,
    is_original: bool,
}

#[derive(Clone, Copy, Eq, PartialEq)]
enum ShapeOrigin {
    Original,
    Added,
}

fn shape_origin<T: ReadTxn>(shape: &MapRef, txn: &T) -> EditResult<ShapeOrigin> {
    match required_string(shape, txn, "origin")?.as_str() {
        "original" => Ok(ShapeOrigin::Original),
        "added" => Ok(ShapeOrigin::Added),
        _ => Err(EditError::InvalidState("invalid shape origin".to_owned())),
    }
}

fn page_part_paths_for_snapshot(
    package: &vsdx_parse::VsdxPackage,
    snapshot: &DiagramSnapshot,
) -> EditResult<Vec<String>> {
    let paths = snapshot
        .pages
        .iter()
        .map(|page| page.source_part_path.clone())
        .collect::<Vec<_>>();
    let expected = package.page_part_paths.iter().collect::<HashSet<_>>();
    if paths.len() != expected.len()
        || paths.iter().collect::<HashSet<_>>().len() != paths.len()
        || !paths.iter().all(|path| expected.contains(path))
    {
        return Err(EditError::InvalidState(
            "page order does not match the original package".to_owned(),
        ));
    }
    Ok(paths)
}

fn structural_shape_entries<T: ReadTxn>(
    sheets: &MapRef,
    txn: &T,
    roots: &ArrayRef,
    snapshots: &std::collections::BTreeMap<&str, &ShapeSnapshot>,
) -> EditResult<Vec<StructuralShapeEntry>> {
    shape_tree_entries(sheets, txn, roots)?
        .into_iter()
        .map(|entry| {
            let shape = map_ref(sheets, txn, &entry.id)?;
            let source_id = snapshots
                .get(entry.id.as_str())
                .ok_or_else(|| {
                    EditError::InvalidState("shape is missing from snapshot".to_owned())
                })?
                .source_id;
            let is_original = shape_origin(&shape, txn)? == ShapeOrigin::Original;
            Ok(StructuralShapeEntry {
                id: entry.id,
                parent_id: entry.parent_id,
                source_id,
                is_original,
            })
        })
        .collect()
}

fn original_shape_containers(
    sheet: &vsdx_parse::Sheet,
) -> std::collections::BTreeMap<Option<u32>, Vec<u32>> {
    let mut containers = std::collections::BTreeMap::new();
    let mut pending = vec![(None, sheet.shapes().collect::<Vec<_>>())];
    while let Some((parent, shapes)) = pending.pop() {
        let ids = shapes.iter().map(|shape| shape.id).collect::<Vec<_>>();
        containers.insert(parent, ids);
        for shape in shapes.into_iter().rev() {
            pending.push((Some(shape.id), shape.shapes().collect()));
        }
    }
    containers
}

fn snapshot_shapes(page: &PageSnapshot) -> std::collections::BTreeMap<&str, &ShapeSnapshot> {
    let mut result = std::collections::BTreeMap::new();
    let mut pending = page.shapes.iter().collect::<Vec<_>>();
    while let Some(shape) = pending.pop() {
        result.insert(shape.id.as_str(), shape);
        pending.extend(shape.children.iter());
    }
    result
}

fn structural_container_edits(
    page_id: u32,
    originals: &[u32],
    desired: &[&StructuralShapeEntry],
    added: &std::collections::BTreeMap<&str, u32>,
    edits: &mut Vec<StructuralEdit>,
) -> EditResult<()> {
    let present = desired
        .iter()
        .filter(|shape| shape.is_original && originals.contains(&shape.source_id))
        .map(|shape| shape.source_id)
        .collect::<HashSet<_>>();
    for shape_id in originals {
        if !present.contains(shape_id) {
            edits.push(StructuralEdit::DeleteShape {
                page_id,
                shape_id: *shape_id,
            });
        }
    }
    let mut order = originals
        .iter()
        .filter(|shape| present.contains(shape))
        .copied()
        .collect::<Vec<_>>();
    for shape in desired.iter().filter(|shape| !shape.is_original) {
        order.push(*added.get(shape.id.as_str()).ok_or_else(|| {
            EditError::InvalidState(format!(
                "added shape {:?} is missing its allocated source ID",
                shape.id
            ))
        })?);
    }
    let desired_ids = desired
        .iter()
        .map(|shape| {
            if shape.is_original {
                Ok(shape.source_id)
            } else {
                added.get(shape.id.as_str()).copied().ok_or_else(|| {
                    EditError::InvalidState(format!(
                        "added shape {:?} is missing its allocated source ID",
                        shape.id
                    ))
                })
            }
        })
        .collect::<EditResult<Vec<_>>>()?;
    for (index, shape_id) in desired_ids.iter().enumerate() {
        if order.get(index) == Some(shape_id) {
            continue;
        }
        let before_shape_id = order.get(index).copied();
        edits.push(StructuralEdit::ReorderShape {
            page_id,
            shape_id: *shape_id,
            before_shape_id,
        });
        let from = order.iter().position(|id| id == shape_id).ok_or_else(|| {
            EditError::InvalidState(format!(
                "shape {shape_id} is requested in order but is absent from its container"
            ))
        })?;
        order.remove(from);
        order.insert(index, *shape_id);
    }
    Ok(())
}

fn shape_xml(shape: &ShapeSnapshot) -> String {
    let mut output = format!("<Shape Type=\"Shape\" ID=\"{}\"", shape.source_id);
    if let Some(name) = &shape.name {
        output.push_str(" Name=\"");
        xml_escape(&mut output, name);
        output.push('\"');
    }
    output.push('>');
    for cell in &shape.cells {
        if cell.locator.section.is_none() {
            cell_xml(&mut output, cell);
        }
    }
    let mut sections: SectionRows<'_> = Vec::new();
    for cell in &shape.cells {
        let Some(section) = &cell.locator.section else {
            continue;
        };
        let section_identity = (section.clone(), cell.locator.section_index);
        let section_index = sections
            .iter()
            .position(|(identity, _)| identity == &section_identity)
            .unwrap_or_else(|| {
                sections.push((section_identity, Vec::new()));
                sections.len() - 1
            });
        let rows = &mut sections[section_index].1;
        let row_index = rows
            .iter()
            .position(|(row, _)| row == &cell.locator.row)
            .unwrap_or_else(|| {
                rows.push((cell.locator.row.clone(), Vec::new()));
                rows.len() - 1
            });
        rows[row_index].1.push(cell);
    }
    for ((section, index), rows) in sections {
        output.push_str("<Section N=\"");
        xml_escape(&mut output, &section);
        if let Some(index) = index {
            output.push_str(&format!("\" IX=\"{index}"));
        }
        output.push_str("\">");
        for (row, cells) in rows {
            output.push_str("<Row");
            match row {
                Some(CellRow::Index(index)) => output.push_str(&format!(" IX=\"{index}\"")),
                Some(CellRow::Name(name)) => {
                    output.push_str(" N=\"");
                    xml_escape(&mut output, &name);
                    output.push('\"');
                }
                None => {}
            }
            if let Some(row_type) = cells.iter().find_map(|cell| cell.row_type.as_deref()) {
                output.push_str(" T=\"");
                xml_escape(&mut output, row_type);
                output.push('"');
            }
            output.push('>');
            for cell in cells {
                cell_xml(&mut output, cell);
            }
            output.push_str("</Row>");
        }
        output.push_str("</Section>");
    }
    output.push_str("</Shape>");
    output
}

fn cell_xml(output: &mut String, cell: &CellSnapshot) {
    output.push_str("<Cell N=\"");
    xml_escape(output, &cell.name);
    output.push('\"');
    if let Some(formula) = &cell.formula {
        output.push_str(" F=\"");
        xml_escape(output, formula);
        output.push('\"');
    }
    if let Some(value) = &cell.value {
        output.push_str(" V=\"");
        xml_escape(output, value);
        output.push('\"');
    }
    output.push_str("/>");
}

fn xml_escape(output: &mut String, value: &str) {
    for character in value.chars() {
        match character {
            '&' => output.push_str("&amp;"),
            '<' => output.push_str("&lt;"),
            '>' => output.push_str("&gt;"),
            '\"' => output.push_str("&quot;"),
            '\'' => output.push_str("&apos;"),
            _ => output.push(character),
        }
    }
}

fn cell_map(
    txn: &mut TransactionMut<'_>,
    page_id: &str,
    shape_id: &str,
    locator: &CellLocator,
) -> EditResult<MapRef> {
    let pages = txn
        .get_map(PAGES)
        .ok_or_else(|| EditError::InvalidState("missing pages map".to_owned()))?;
    map_ref(&pages, txn, page_id)?;
    let sheets = txn
        .get_map(SHEETS)
        .ok_or_else(|| EditError::InvalidState("missing sheets map".to_owned()))?;
    let shape = map_ref(&sheets, txn, shape_id)?;
    let cells = map_map(&shape, txn, "cells")?;
    let key = locator_key(locator);
    cells
        .get(txn, key.as_str())
        .and_then(|value| {
            if let Out::YMap(map) = value {
                Some(map)
            } else {
                None
            }
        })
        .ok_or(EditError::CellNotFound(key))
}
fn reorder(
    session: &DiagramSession,
    context: &EditCtx,
    page_id: &str,
    shape_id: &str,
    to: u32,
) -> EditResult<ShapeReceipt> {
    let mut txn = session.transact_for(context);
    let pages = txn
        .get_map(PAGES)
        .ok_or_else(|| EditError::InvalidState("missing pages map".to_owned()))?;
    let page = map_ref(&pages, &txn, page_id)?;
    let root_order = map_array(&page, &txn, "shapes")?;
    let sheets = txn
        .get_map(SHEETS)
        .ok_or_else(|| EditError::InvalidState("missing sheets map".to_owned()))?;
    let entries = shape_tree_entries(&sheets, &txn, &root_order)?;
    let target = entries
        .iter()
        .find(|entry| entry.id == shape_id)
        .ok_or_else(|| EditError::ShapeNotFound(shape_id.to_owned()))?;
    let order = target.order.clone();
    let length = order.len(&txn);
    if to >= length {
        return Err(EditError::OutOfBounds { index: to, length });
    }
    let from = target.index;
    order.remove_range(&mut txn, from, 1);
    order.insert(&mut txn, to, shape_id);
    Ok(ShapeReceipt {
        page_id: page_id.to_owned(),
        shape_id: shape_id.to_owned(),
        from_index: Some(from),
        to_index: Some(to),
    })
}
fn reorder_page(
    session: &DiagramSession,
    context: &EditCtx,
    page_id: &str,
    to: u32,
) -> EditResult<ShapeReceipt> {
    let mut txn = session.transact_for(context);
    let order = txn
        .get_array(PAGE_ORDER)
        .ok_or_else(|| EditError::InvalidState("missing page order".to_owned()))?;
    let length = order.len(&txn);
    if to >= length {
        return Err(EditError::OutOfBounds { index: to, length });
    }
    let from = (0..length)
        .find(|index| array_string(&order, &txn, *index).as_deref() == Some(page_id))
        .ok_or_else(|| EditError::PageNotFound(page_id.to_owned()))?;
    order.remove_range(&mut txn, from, 1);
    order.insert(&mut txn, to, page_id);
    Ok(ShapeReceipt {
        page_id: page_id.to_owned(),
        shape_id: page_id.to_owned(),
        from_index: Some(from),
        to_index: Some(to),
    })
}

struct CrdtMutationContext {
    page_id: u32,
    shape_id: u32,
    formulas: std::collections::BTreeMap<String, String>,
    values: std::collections::BTreeMap<String, String>,
    references: vsdx_resolve::ResolvedShape,
}

impl CrdtMutationContext {
    fn new<T: ReadTxn>(txn: &T, page_id: &str, shape_id: &str) -> EditResult<Self> {
        let pages = required_map(txn, PAGES)?;
        map_ref(&pages, txn, page_id)?;
        let sheets = required_map(txn, SHEETS)?;
        let shape = map_ref(&sheets, txn, shape_id)?;
        if map_string(&shape, txn, "pageId").as_deref() != Some(page_id) {
            return Err(EditError::ShapeNotFound(shape_id.to_owned()));
        }
        let cells = map_map(&shape, txn, "cells")?;
        let mut formulas = std::collections::BTreeMap::new();
        let mut values = std::collections::BTreeMap::new();
        for (_name, cell) in cells.iter(txn) {
            let Out::YMap(cell) = cell else { continue };
            let locator = cell_locator(&cell, txn, 0, 0)?;
            let key = locator_key(&locator);
            if let Some(formula) = map_string(&cell, txn, "formula") {
                formulas.insert(key.clone(), formula);
            }
            if let Some(value) = map_string(&cell, txn, "value") {
                values.insert(key, value);
            }
        }
        Ok(Self {
            page_id: page_id
                .trim_start_matches("page:")
                .parse()
                .unwrap_or_default(),
            shape_id: map_number(&shape, txn, "sourceId").unwrap_or_default() as u32,
            formulas,
            values,
            references: local_references(&cells, txn)?,
        })
    }

    fn locator(&self, mut locator: CellLocator) -> CellLocator {
        locator.sheet = CellSheet::Page(self.page_id);
        locator.shape_id = Some(self.shape_id);
        locator
    }
}

impl MutationContext for CrdtMutationContext {
    fn current_formula(&self, locator: &CellLocator) -> Result<Option<String>, String> {
        if locator.sheet != CellSheet::Page(self.page_id) || locator.shape_id != Some(self.shape_id)
        {
            return Err("cross-sheet mutation targets are not supported".to_owned());
        }
        Ok(self.formulas.get(&locator_key(locator)).cloned())
    }

    fn resolve_reference(
        &self,
        from: &CellLocator,
        reference: &str,
    ) -> Result<CellLocator, String> {
        if reference.contains('!') {
            return Err("cross-sheet SETATREF targets are not supported".to_owned());
        }
        if !self.formulas.contains_key(reference) && !self.values.contains_key(reference) {
            return Err(format!("SETATREF target does not exist: {reference}"));
        }
        Ok(CellLocator {
            cell_name: reference.to_owned(),
            ..from.clone()
        })
    }

    fn lock_enabled(&self, _locator: &CellLocator, lock: &str) -> Result<bool, String> {
        if lock.is_empty() {
            return Ok(false);
        }
        let formula = self.formulas.get(lock).or_else(|| self.values.get(lock));
        let Some(formula) = formula else {
            return Ok(false);
        };
        match vsdx_eval::evaluate_cell(
            lock,
            formula.trim_start_matches('='),
            &self.references,
            &ParseLimits::default(),
        ) {
            vsdx_eval::Evaluation::Evaluated(value) => match value.value {
                vsdx_eval::Value::Number(number) => Ok(number.number == 1.0),
                vsdx_eval::Value::Color(_) => Err(format!("cannot evaluate {lock}")),
            },
            _ => Err(format!("cannot evaluate {lock}")),
        }
    }
}

fn gesture_for_cell(cell_name: &str) -> MutationGesture {
    match cell_name {
        "PinX" => MutationGesture::MoveX,
        "PinY" => MutationGesture::MoveY,
        "Width" => MutationGesture::ResizeWidth,
        "Height" => MutationGesture::ResizeHeight,
        _ => MutationGesture::CellEdit,
    }
}

fn locator_key(locator: &CellLocator) -> String {
    let section = locator
        .section
        .as_ref()
        .map(|name| vsdx_resolve::section_key(name, locator.section_index));
    match (&section, &locator.row) {
        (Some(section), Some(CellRow::Index(row))) => {
            format!("{section}\u{1f}IX:{row}\u{1f}{}", locator.cell_name)
        }
        (Some(section), Some(CellRow::Name(row))) => {
            format!("{section}\u{1f}N:{row}\u{1f}{}", locator.cell_name)
        }
        _ => locator.cell_name.clone(),
    }
}

fn cell_locator<T: ReadTxn>(
    cell: &MapRef,
    txn: &T,
    page_id: u32,
    shape_id: u32,
) -> EditResult<CellLocator> {
    for key in ["section", "rowName", "rowType"] {
        if cell.get(txn, key).is_some() && map_string(cell, txn, key).is_none() {
            return Err(EditError::InvalidState(format!(
                "cell {key} is not a string"
            )));
        }
    }
    let row = match (
        map_u32(cell, txn, "rowIndex")?,
        map_string(cell, txn, "rowName"),
    ) {
        (Some(_), Some(_)) => {
            return Err(EditError::InvalidState(
                "cell has both row index and name".to_owned(),
            ));
        }
        (Some(index), None) => Some(CellRow::Index(index)),
        (None, Some(name)) => Some(CellRow::Name(name)),
        (None, None) => None,
    };
    let section = map_string(cell, txn, "section");
    let section_index = map_u32(cell, txn, "sectionIndex")?;
    if section.is_none() && (row.is_some() || section_index.is_some()) {
        return Err(EditError::InvalidState(
            "cell row/index requires a section".to_owned(),
        ));
    }
    Ok(CellLocator {
        sheet: CellSheet::Page(page_id),
        shape_id: Some(shape_id),
        section,
        section_index,
        row,
        cell_name: required_string(cell, txn, "name")?,
    })
}
fn required_map<T: ReadTxn>(txn: &T, name: &str) -> EditResult<MapRef> {
    txn.get_map(name)
        .ok_or_else(|| EditError::InvalidState(format!("missing {name} map")))
}
fn required_array<T: ReadTxn>(txn: &T, name: &str) -> EditResult<ArrayRef> {
    txn.get_array(name)
        .ok_or_else(|| EditError::InvalidState(format!("missing {name} array")))
}
fn map_ref<T: ReadTxn>(map: &MapRef, txn: &T, key: &str) -> EditResult<MapRef> {
    map.get(txn, key)
        .and_then(|value| {
            if let Out::YMap(map) = value {
                Some(map)
            } else {
                None
            }
        })
        .ok_or_else(|| EditError::PageNotFound(key.to_owned()))
}
fn map_map<T: ReadTxn>(map: &MapRef, txn: &T, key: &str) -> EditResult<MapRef> {
    map.get(txn, key)
        .and_then(|value| {
            if let Out::YMap(map) = value {
                Some(map)
            } else {
                None
            }
        })
        .ok_or_else(|| EditError::InvalidState(format!("missing {key} map")))
}
fn map_array<T: ReadTxn>(map: &MapRef, txn: &T, key: &str) -> EditResult<ArrayRef> {
    map.get(txn, key)
        .and_then(|value| {
            if let Out::YArray(array) = value {
                Some(array)
            } else {
                None
            }
        })
        .ok_or_else(|| EditError::InvalidState(format!("missing {key} array")))
}
fn map_string<T: ReadTxn>(map: &MapRef, txn: &T, key: &str) -> Option<String> {
    map.get(txn, key).and_then(|value| match value {
        Out::Any(Any::String(value)) => Some(value.to_string()),
        _ => None,
    })
}
fn required_string<T: ReadTxn>(map: &MapRef, txn: &T, key: &str) -> EditResult<String> {
    map_string(map, txn, key).ok_or_else(|| EditError::InvalidState(format!("missing {key}")))
}
fn map_u32<T: ReadTxn>(map: &MapRef, txn: &T, key: &str) -> EditResult<Option<u32>> {
    match map.get(txn, key) {
        None => Ok(None),
        Some(Out::Any(Any::Number(number)))
            if number.is_finite()
                && number >= 0.0
                && number <= f64::from(u32::MAX)
                && number.fract() == 0.0 =>
        {
            Ok(Some(number as u32))
        }
        _ => Err(EditError::InvalidState(format!(
            "{key} is not an unsigned 32-bit integer"
        ))),
    }
}

fn map_number<T: ReadTxn>(map: &MapRef, txn: &T, key: &str) -> Option<f64> {
    match map.get(txn, key) {
        Some(Out::Any(Any::Number(number))) => Some(number),
        _ => None,
    }
}
fn array_string<T: ReadTxn>(array: &ArrayRef, txn: &T, index: u32) -> Option<String> {
    array.get(txn, index).and_then(|value| match value {
        Out::Any(Any::String(value)) => Some(value.to_string()),
        _ => None,
    })
}
