use std::collections::HashSet;
use std::sync::Arc;

use vsdx_eval::{
    Evaluation, Expr, MutationContext, MutationOutcome, References, Value, decide_mutation,
    evaluate,
};
use vsdx_parse::{
    Cell, CellLocator, CellRow, CellSheet, MutationGesture, ParseLimits, RowChild, SectionChild,
    Shape, ShapeChild, ShapesChild, SheetChild, StructuralEdit, TextToken,
};
use vsdx_resolve::{Lookup, Resolver};
use yrs::{
    Any, Array, ArrayPrelim, ArrayRef, Doc, Map, MapPrelim, MapRef, Out, ReadTxn, Transact,
    TransactionMut, WriteTxn,
};

use crate::{
    CONNECTS, CellFormulaReceipt, CellFormulaWrite, CellSnapshot, CellWriteProbe, CellWriteQuery,
    ConnectorRouteReceipt, DiagramSession, DiagramSnapshot, EditCtx, EditError, EditResult, META,
    PAGE_ORDER, PAGES, PageSnapshot, SHEETS, STORIES, ShapeDataReceipt, ShapeDataWrite,
    ShapeDelete, ShapeDraft, ShapeMove, ShapeReceipt, ShapeSnapshot, ShapeTreeDraft, ShapeTreeGlue,
    TextReceipt,
};

mod connect;
mod containers;

const SCHEMA_VERSION: f64 = 2.0;
/// Stories held resolved text tokens as JSON before this version.
const TOKEN_STORY_SCHEMA_VERSION: f64 = 1.0;
pub(crate) const MAX_SHAPE_NESTING: usize = 256;
const MAX_CONNECTOR_ROUTE_POINTS: usize = 256;
const ROUTE_POINT_EPSILON: f64 = 1e-9;
type SectionRows<'a> = Vec<(
    (String, Option<u32>),
    Vec<(Option<CellRow>, Vec<&'a CellSnapshot>)>,
)>;
type StoryTexts = (
    std::collections::BTreeMap<String, String>,
    std::collections::BTreeMap<String, Vec<TextToken>>,
);
type DeleteCandidate = (usize, String, String, ArrayRef, u32, usize, Vec<String>);

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
    let glue = connect::glue_records(&doc.transact())?;
    let (texts, verbatim) = edited_story_texts(doc, &snapshot, &package)?;
    materialize_snapshot(
        doc,
        &mut package,
        &snapshot,
        &original_shape_ids(doc)?,
        &glue,
        &texts,
        &verbatim,
    )?;
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

/// Master IDs the opened package defines; sessions never add or remove masters.
fn original_master_ids(doc: &Doc) -> EditResult<std::collections::BTreeSet<u32>> {
    Ok(original_package_from_doc(doc)?
        .master_sheets
        .into_keys()
        .collect())
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
    doc: &Doc,
    package: &mut vsdx_parse::VsdxPackage,
    snapshot: &DiagramSnapshot,
    original_shape_ids: &HashSet<String>,
    glue: &[connect::GlueRecord],
    texts: &std::collections::BTreeMap<String, String>,
    verbatim: &std::collections::BTreeMap<String, Vec<TextToken>>,
) -> EditResult<()> {
    for page in &snapshot.pages {
        let Some(sheet) = package.page_contents.get_mut(&page.source_part_path) else {
            continue;
        };
        let original_source_shape_ids = sheet_shape_ids(sheet);
        materialize_page_shapes(
            sheet,
            page,
            original_shape_ids,
            &added_page_remaps(doc, page)?,
        )?;
        materialize_page_text(sheet, page, texts, verbatim);
        let projected_shape_ids = sheet_shape_ids(sheet);
        let deleted = original_source_shape_ids
            .difference(&projected_shape_ids)
            .copied()
            .collect();
        vsdx_parse::remove_connects_referencing_shapes(sheet, &deleted);
        connect::materialize_page_glue(sheet, page, glue);
    }
    Ok(())
}

/// Stored-to-final source ID maps for pasted subtrees on one snapshot page.
fn added_page_remaps(
    doc: &Doc,
    page: &PageSnapshot,
) -> EditResult<std::collections::BTreeMap<String, std::collections::BTreeMap<u32, u32>>> {
    let txn = doc.transact();
    let pages = required_map(&txn, PAGES)?;
    let sheets = required_map(&txn, SHEETS)?;
    let session_page = map_ref(&pages, &txn, &page.id)?;
    let roots = map_array(&session_page, &txn, "shapes")?;
    let mut finals = std::collections::BTreeMap::new();
    let mut pending = page.shapes.iter().collect::<Vec<_>>();
    while let Some(shape) = pending.pop() {
        finals.insert(shape.id.clone(), shape.source_id);
        pending.extend(shape.children.iter());
    }
    added_subtree_remaps(&sheets, &txn, &roots, &finals)
}

/// Projects edited plain text onto the parsed model; edited shapes collapse to one literal.
fn materialize_page_text(
    sheet: &mut vsdx_parse::Sheet,
    page: &PageSnapshot,
    texts: &std::collections::BTreeMap<String, String>,
    verbatim: &std::collections::BTreeMap<String, Vec<TextToken>>,
) {
    let mut pending = page.shapes.iter().collect::<Vec<_>>();
    while let Some(shape) = pending.pop() {
        if let Some(tokens) = verbatim
            .get(shape.id.as_str())
            .filter(|tokens| !tokens.is_empty())
            && let Some(target) = shape_by_source_mut(sheet, shape.source_id)
        {
            target
                .children
                .retain(|child| !matches!(child, ShapeChild::Text(_)));
            target.children.push(ShapeChild::Text(tokens.clone()));
        } else if let Some(text) = texts.get(shape.id.as_str())
            && let Some(target) = shape_by_source_mut(sheet, shape.source_id)
        {
            target
                .children
                .retain(|child| !matches!(child, ShapeChild::Text(_)));
            let tokens = match text.is_empty() {
                true => Vec::new(),
                false => vec![vsdx_parse::TextToken::Literal(text.clone())],
            };
            target.children.push(ShapeChild::Text(tokens));
        }
        pending.extend(shape.children.iter());
    }
}

/// Finds a page-local shape by its materialized source ID.
fn shape_by_source_mut(sheet: &mut vsdx_parse::Sheet, source_id: u32) -> Option<&mut Shape> {
    fn shape_mut(shape: &mut Shape, source_id: u32) -> Option<&mut Shape> {
        if shape.id == source_id {
            return Some(shape);
        }
        for child in &mut shape.children {
            let ShapeChild::Shapes(children) = child else {
                continue;
            };
            for candidate in children.iter_mut() {
                let ShapesChild::Shape(candidate) = candidate else {
                    continue;
                };
                if let Some(found) = shape_mut(candidate, source_id) {
                    return Some(found);
                }
            }
        }
        None
    }
    for child in &mut sheet.children {
        let SheetChild::Shapes(children) = child else {
            continue;
        };
        for candidate in children.iter_mut() {
            let ShapesChild::Shape(candidate) = candidate else {
                continue;
            };
            if let Some(found) = shape_mut(candidate, source_id) {
                return Some(found);
            }
        }
    }
    None
}

/// Maps diverging stories to session shape IDs for materialization.
fn edited_story_texts(
    doc: &Doc,
    snapshot: &DiagramSnapshot,
    package: &vsdx_parse::VsdxPackage,
) -> EditResult<StoryTexts> {
    let mut source_to_session = std::collections::BTreeMap::new();
    for page in &snapshot.pages {
        let Some(source_page_id) = page
            .id
            .strip_prefix("page:")
            .and_then(|value| value.parse::<u32>().ok())
        else {
            continue;
        };
        let mut pending = page.shapes.iter().collect::<Vec<_>>();
        while let Some(shape) = pending.pop() {
            source_to_session.insert((source_page_id, shape.source_id), shape.id.clone());
            pending.extend(shape.children.iter());
        }
    }
    let mut texts = std::collections::BTreeMap::new();
    for edit in semantic_text_edits_in(doc, package)? {
        let (source_page_id, source_shape_id) = match (&edit.locator.sheet, edit.locator.shape_id) {
            (CellSheet::Page(page_id), Some(shape_id)) => (*page_id, shape_id),
            _ => continue,
        };
        if let Some(session_id) = source_to_session.get(&(source_page_id, source_shape_id)) {
            texts.insert(session_id.clone(), edit.text);
        }
    }
    let mut verbatim = std::collections::BTreeMap::new();
    let txn = doc.transact();
    let sheets = required_map(&txn, SHEETS)?;
    let stories = txn.get_map(STORIES);
    for (shape_id, shape) in sheets.iter(&txn) {
        let Out::YMap(shape) = shape else { continue };
        if shape_origin(&shape, &txn)? != ShapeOrigin::Added {
            continue;
        }
        let current = match stories
            .as_ref()
            .and_then(|stories| stories.get(&txn, shape_id))
        {
            Some(Out::Any(Any::String(value))) => value.to_string(),
            _ => continue,
        };
        if !current.is_empty() {
            validate_story_text(&current)?;
            texts.insert(shape_id.to_owned(), current.clone());
        }
        let page_id = map_string(&shape, &txn, "pageId").and_then(|page_id| {
            page_id
                .strip_prefix("page:")
                .and_then(|value| value.parse::<u32>().ok())
        });
        let copy_source_id = map_u32(&shape, &txn, "copySourceId")?;
        let copy_source_page = map_u32(&shape, &txn, "copySourcePageId")?;
        if let (Some(page_id), Some(copy_source_id)) =
            (copy_source_page.or(page_id), copy_source_id)
            && let Some(tokens) = verbatim_source_text(package, page_id, copy_source_id, &current)
        {
            verbatim.insert(shape_id.to_owned(), tokens);
        }
    }
    Ok((texts, verbatim))
}

/// Verbatim `Text` tokens of the copied shape, when its story still matches.
fn verbatim_source_text(
    package: &vsdx_parse::VsdxPackage,
    source_page_id: u32,
    copy_source_id: u32,
    story: &str,
) -> Option<Vec<TextToken>> {
    let path = package
        .page_part_ids
        .iter()
        .find_map(|(path, id)| (*id == source_page_id).then(|| path.clone()))?;
    let contents = package.page_contents.get(&path)?;
    let original = find_shape_in(contents, copy_source_id)?;
    let local = original.text().filter(|tokens| !tokens.is_empty())?;
    let page_sheet = package
        .page_part_ids
        .get(&path)
        .and_then(|id| package.page_sheets.get(id))
        .unwrap_or(contents);
    let resolver = Resolver::new(package);
    let resolved = resolver.resolve_shape(&path, copy_source_id).ok()?;
    let tokens = resolver
        .resolve_text_in_context(original, page_sheet, &resolved)
        .ok()?;
    (plain_text(&tokens) == story).then(|| local.to_vec())
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
    remaps: &std::collections::BTreeMap<String, std::collections::BTreeMap<u32, u32>>,
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
        materialize_shape(&mut shape, snapshot, original_shape_ids, 1, remaps)?;
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
        shape_type: Some(if snapshot.children.is_empty() {
            "Shape".to_owned()
        } else {
            "Group".to_owned()
        }),
        master: snapshot.master,
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
    remaps: &std::collections::BTreeMap<String, std::collections::BTreeMap<u32, u32>>,
) -> EditResult<()> {
    if depth > MAX_SHAPE_NESTING {
        return Err(EditError::InvalidState(
            "shape nesting exceeds maximum depth".to_owned(),
        ));
    }
    let remap = remaps.get(snapshot.id.as_str());
    for cell in &snapshot.cells {
        materialize_cell(shape, cell, remap);
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
        materialize_shape(
            &mut child,
            child_snapshot,
            original_shape_ids,
            depth + 1,
            remaps,
        )?;
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

fn materialize_cell(
    shape: &mut Shape,
    snapshot: &CellSnapshot,
    remap: Option<&std::collections::BTreeMap<u32, u32>>,
) {
    let locator = &snapshot.locator;
    let formula = snapshot.formula.as_deref().map(|formula| match remap {
        Some(remap) => remap_sheet_refs(formula, remap).into_owned(),
        None => formula.to_owned(),
    });
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
        cell.formula = formula.clone();
        cell.value = snapshot.value.clone();
    } else if locator.section.is_none() {
        shape.children.push(ShapeChild::Cell(Cell {
            name: locator.cell_name.clone(),
            formula: formula.clone(),
            value: snapshot.value.clone(),
            unit: None,
            del: false,
            other_attrs: Vec::new(),
        }));
    } else if let (Some(section_name), Some(row)) = (&locator.section, &locator.row) {
        let cell = Cell {
            name: locator.cell_name.clone(),
            formula: formula.clone(),
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
    if let Some(master) = shape.master {
        map.insert(txn, "master", f64::from(master));
    }
    if let Some(parent_id) = parent_id {
        map.insert(txn, "parentId", parent_id);
    }
    if let Some(name) = &shape.name {
        map.insert(txn, "name", name.as_str());
    }
    if let Some(reason) = copy_refusal_reason(shape) {
        map.insert(txn, "copyRefusal", reason);
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
    stories.insert(txn, id, plain_text(&text));
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

/// Reason a shape cannot be copied losslessly, from its own package content.
fn copy_refusal_reason(shape: &vsdx_parse::Shape) -> Option<&'static str> {
    for child in &shape.children {
        match child {
            ShapeChild::ForeignData(_) => return Some("embedded media"),
            ShapeChild::Unknown(_) => return Some("unsupported shape content"),
            ShapeChild::Section(section) => {
                for child in &section.children {
                    let SectionChild::Row(row) = child else {
                        return Some("unsupported shape content");
                    };
                    if row.index.is_none() && row.name.is_none() {
                        return Some("unsupported geometry rows");
                    }
                    if row
                        .children
                        .iter()
                        .any(|child| !matches!(child, RowChild::Cell(_)))
                    {
                        return Some("unsupported shape content");
                    }
                }
            }
            ShapeChild::Text(_) | ShapeChild::Cell(_) | ShapeChild::Shapes(_) => {}
        }
    }
    None
}

/// Plain text behind resolved `Text` markers; paragraph and tab runs become controls.
fn plain_text(tokens: &[vsdx_resolve::ResolvedTextToken]) -> String {
    let mut output = String::new();
    for token in tokens {
        match token {
            vsdx_resolve::ResolvedTextToken::Literal(value) => output.push_str(value),
            vsdx_resolve::ResolvedTextToken::CharacterRun { .. } => {}
            vsdx_resolve::ResolvedTextToken::ParagraphRun { .. } => {
                if !output.is_empty() {
                    output.push('\n');
                }
            }
            vsdx_resolve::ResolvedTextToken::Tab { .. } => output.push('\t'),
            vsdx_resolve::ResolvedTextToken::Field { properties, .. } => {
                let value = properties.get("Value").and_then(|lookup| match lookup {
                    Lookup::Found(cell) => cell.cell.value.clone(),
                    Lookup::Deleted | Lookup::Absent => None,
                });
                if let Some(value) = value {
                    output.push_str(&value);
                }
            }
        }
    }
    output
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

/// Seeds a hand-routed Geometry cell with no baseline or cached value.
fn seed_route_cell(
    cells: &MapRef,
    txn: &mut TransactionMut<'_>,
    locator: &CellLocator,
    formula: &str,
    row_type: &str,
) {
    let key = locator_key(locator);
    let cell = cells.insert(txn, key.as_str(), MapPrelim::default());
    cell.insert(txn, "name", locator.cell_name.as_str());
    if let Some(section) = &locator.section {
        cell.insert(txn, "section", section.as_str());
    }
    if let Some(CellRow::Index(index)) = &locator.row {
        cell.insert(txn, "rowIndex", *index as f64);
    }
    cell.insert(txn, "rowType", row_type);
    cell.insert(txn, "formula", formula);
}

struct GeometryRow {
    row_type: Option<String>,
    has_x: bool,
    has_y: bool,
}

impl GeometryRow {
    fn has(&self, name: &str) -> bool {
        match name {
            "X" => self.has_x,
            _ => self.has_y,
        }
    }
}

/// Reuses the filed row indices in order, then appends past the highest one.
fn route_row_index(filed: &[u32], slot: usize) -> u32 {
    match filed.get(slot) {
        Some(index) => *index,
        None => filed.last().map_or(slot as u32, |last| {
            last.saturating_add((slot - filed.len() + 1) as u32)
        }),
    }
}

fn allow_route_cell(
    context: &CrdtMutationContext,
    locator: CellLocator,
    formula: &str,
) -> EditResult<()> {
    match decide_mutation(
        context,
        locator.clone(),
        MutationGesture::CellEdit,
        formula.to_owned(),
        &ParseLimits::default(),
    ) {
        MutationOutcome::Allowed { target, .. } if target == locator => Ok(()),
        MutationOutcome::Allowed { .. } => Err(EditError::InvalidState(
            "connector route edit bypasses formula redirect".to_owned(),
        )),
        MutationOutcome::Refused { reason } | MutationOutcome::Unsupported { reason } => {
            Err(EditError::InvalidState(reason))
        }
    }
}

/// A 1D shape resolves the way the glue resolver decides it: explicit OneD wins,
/// otherwise all four endpoint cells must resolve.
fn connector_route_shape(shape: &vsdx_resolve::ResolvedShape) -> bool {
    match route_shape_number(shape, "OneD") {
        Some(value) => value != 0.0,
        None => ["BeginX", "BeginY", "EndX", "EndY"]
            .into_iter()
            .all(|name| route_shape_number(shape, name).is_some()),
    }
}

fn route_shape_number(shape: &vsdx_resolve::ResolvedShape, name: &str) -> Option<f64> {
    let Lookup::Found(cell) = shape.cell(name)? else {
        return None;
    };
    route_cell_number(Some(&cell.cell), shape)
}

fn route_cell_number(cell: Option<&Cell>, shape: &vsdx_resolve::ResolvedShape) -> Option<f64> {
    let cell = cell?;
    if let Some(formula) = cell.formula.as_deref()
        && let vsdx_eval::Evaluation::Evaluated(result) = evaluate(
            formula.trim_start_matches('='),
            shape,
            &ParseLimits::default(),
        )
        && let vsdx_eval::Value::Number(number) = result.value
        && number.number.is_finite()
    {
        return Some(number.number);
    }
    cell.value
        .as_deref()?
        .trim()
        .parse::<f64>()
        .ok()
        .filter(|value| value.is_finite())
}

struct RouteFrame {
    a: f64,
    b: f64,
    c: f64,
    d: f64,
    e: f64,
    f: f64,
}

impl RouteFrame {
    fn unproject(&self, x: f64, y: f64) -> Option<(f64, f64)> {
        let determinant = self.a * self.d - self.b * self.c;
        if !determinant.is_finite() || determinant.abs() < 1e-12 {
            return None;
        }
        let x = x - self.e;
        let y = y - self.f;
        let local = (
            (self.d * x - self.c * y) / determinant,
            (self.a * y - self.b * x) / determinant,
        );
        (local.0.is_finite() && local.1.is_finite()).then_some(local)
    }
}

/// The connector's local frame, mirroring the resolver's unparented bounds affine.
fn connector_route_frame(shape: &vsdx_resolve::ResolvedShape) -> Option<RouteFrame> {
    let pin_x = route_shape_number(shape, "PinX")?;
    let pin_y = route_shape_number(shape, "PinY")?;
    let width = route_shape_number(shape, "Width")?;
    let height = route_shape_number(shape, "Height")?;
    if ![pin_x, pin_y, width, height]
        .into_iter()
        .all(|value| (value as f32).is_finite())
        || width == 0.0
        || height == 0.0
    {
        return None;
    }
    let loc_pin_x = route_shape_number(shape, "LocPinX").unwrap_or(width / 2.0);
    let loc_pin_y = route_shape_number(shape, "LocPinY").unwrap_or(height / 2.0);
    let angle = route_shape_number(shape, "Angle").unwrap_or(0.0);
    let flip_x = route_shape_number(shape, "FlipX").is_some_and(|value| value != 0.0);
    let flip_y = route_shape_number(shape, "FlipY").is_some_and(|value| value != 0.0);
    let (sin, cos) = angle.sin_cos();
    let (flip_x, flip_y) = (
        if flip_x { -1.0 } else { 1.0 },
        if flip_y { -1.0 } else { 1.0 },
    );
    Some(RouteFrame {
        a: cos * flip_x,
        b: sin * flip_x,
        c: -sin * flip_y,
        d: cos * flip_y,
        e: pin_x - cos * (flip_x * loc_pin_x) + sin * (flip_y * loc_pin_y),
        f: pin_y - sin * (flip_x * loc_pin_x) - cos * (flip_y * loc_pin_y),
    })
}

fn decimal(value: f64) -> String {
    if value == 0.0 {
        "0".to_owned()
    } else {
        value.to_string()
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
        if before.as_deref() == Some(formula.as_str()) {
            return Ok(CellFormulaReceipt {
                page_id: page_id.to_owned(),
                shape_id: shape_id.to_owned(),
                cell_name: target.cell_name,
                before,
                after: formula,
            });
        }
        cell.insert(&mut txn, "formula", formula.as_str());
        Ok(CellFormulaReceipt {
            page_id: page_id.to_owned(),
            shape_id: shape_id.to_owned(),
            cell_name: target.cell_name,
            before,
            after: formula,
        })
    }

    /// Writes a `Control` row's `X` and `Y` in one transaction, so a handle drag is one undo entry.
    pub fn set_control_handle(
        &self,
        context: &EditCtx,
        page_id: &str,
        shape_id: &str,
        row: &str,
        x_formula: Option<String>,
        y_formula: Option<String>,
    ) -> EditResult<Vec<CellFormulaReceipt>> {
        let requested = [("X", x_formula), ("Y", y_formula)]
            .into_iter()
            .filter_map(|(name, formula)| formula.map(|formula| (name, formula)))
            .collect::<Vec<_>>();
        if requested.is_empty() {
            return Err(EditError::InvalidState(
                "control handle edit writes no axis".to_owned(),
            ));
        }
        let mut txn = self.transact_for(context);
        let context_for_policy = CrdtMutationContext::new(&txn, page_id, shape_id)?;
        let mut targets = Vec::with_capacity(requested.len());
        for (name, formula) in requested {
            let locator = CellLocator {
                sheet: CellSheet::Page(0),
                shape_id: None,
                section: Some("Control".to_owned()),
                section_index: None,
                row: Some(CellRow::Name(row.to_owned())),
                cell_name: name.to_owned(),
            };
            match decide_mutation(
                &context_for_policy,
                context_for_policy.locator(locator),
                MutationGesture::CellEdit,
                formula.clone(),
                &ParseLimits::default(),
            ) {
                MutationOutcome::Allowed { target, .. } => targets.push((target, formula)),
                MutationOutcome::Refused { reason } | MutationOutcome::Unsupported { reason } => {
                    return Err(EditError::InvalidState(reason));
                }
            }
        }
        for (index, (target, _)) in targets.iter().enumerate() {
            if targets[..index].iter().any(|(seen, _)| seen == target) {
                return Err(EditError::InvalidState(format!(
                    "redirects converge on {} more than once",
                    target.cell_name
                )));
            }
        }
        let mut prepared = Vec::with_capacity(targets.len());
        for (target, formula) in targets {
            let cell = cell_map(&mut txn, page_id, shape_id, &target)?;
            prepared.push((target, formula, cell));
        }
        let mut receipts = Vec::with_capacity(prepared.len());
        for (target, formula, cell) in prepared {
            let before = map_string(&cell, &txn, "formula");
            cell.insert(&mut txn, "formula", formula.as_str());
            receipts.push(CellFormulaReceipt {
                page_id: page_id.to_owned(),
                shape_id: shape_id.to_owned(),
                cell_name: target.cell_name,
                before,
                after: formula,
            });
        }
        Ok(receipts)
    }

    pub fn shape_text(&self, page_id: &str, shape_id: &str) -> EditResult<String> {
        let txn = self.doc.transact();
        let pages = required_map(&txn, PAGES)?;
        map_ref(&pages, &txn, page_id)?;
        let sheets = required_map(&txn, SHEETS)?;
        let shape = map_ref(&sheets, &txn, shape_id)
            .map_err(|_| EditError::ShapeNotFound(shape_id.to_owned()))?;
        if map_string(&shape, &txn, "pageId").as_deref() != Some(page_id) {
            return Err(EditError::ShapeNotFound(shape_id.to_owned()));
        }
        Ok(story_text(&txn, shape_id))
    }

    pub fn set_shape_text(
        &self,
        context: &EditCtx,
        page_id: &str,
        shape_id: &str,
        text: impl Into<String>,
    ) -> EditResult<TextReceipt> {
        let text = text.into();
        validate_story_text(&text)?;
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
                cell_name: "Text".to_owned(),
            }),
            MutationGesture::TextEdit,
            text.clone(),
            &ParseLimits::default(),
        ) {
            MutationOutcome::Allowed { .. } => {}
            MutationOutcome::Refused { reason } | MutationOutcome::Unsupported { reason } => {
                return Err(EditError::InvalidState(reason));
            }
        }
        let before = story_text(&txn, shape_id);
        txn.get_or_insert_map(STORIES)
            .insert(&mut txn, shape_id, text.as_str());
        Ok(TextReceipt {
            page_id: page_id.to_owned(),
            shape_id: shape_id.to_owned(),
            before,
            after: text,
        })
    }

    pub fn semantic_text_edits(&self) -> EditResult<Vec<vsdx_parse::SemanticTextEdit>> {
        semantic_text_edits(&self.doc)
    }

    pub fn move_shape(
        &self,
        context: &EditCtx,
        page_id: &str,
        shape_id: &str,
        x_formula: impl Into<String>,
        y_formula: impl Into<String>,
    ) -> EditResult<[CellFormulaReceipt; 2]> {
        self.set_cell_formula_group(
            context,
            page_id,
            shape_id,
            [
                ("PinX", x_formula.into(), MutationGesture::MoveX),
                ("PinY", y_formula.into(), MutationGesture::MoveY),
            ],
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
        self.set_cell_formula_group(
            context,
            page_id,
            shape_id,
            [
                ("Width", width_formula.into(), MutationGesture::ResizeWidth),
                (
                    "Height",
                    height_formula.into(),
                    MutationGesture::ResizeHeight,
                ),
            ],
        )
    }

    pub fn set_shape_bounds(
        &self,
        context: &EditCtx,
        page_id: &str,
        shape_id: &str,
        formulas: [String; 4],
    ) -> EditResult<[CellFormulaReceipt; 4]> {
        let [x, y, width, height] = formulas;
        self.set_cell_formula_group(
            context,
            page_id,
            shape_id,
            [
                ("PinX", x, MutationGesture::MoveX),
                ("PinY", y, MutationGesture::MoveY),
                ("Width", width, MutationGesture::ResizeWidth),
                ("Height", height, MutationGesture::ResizeHeight),
            ],
        )
    }

    pub fn move_shapes(
        &self,
        context: &EditCtx,
        moves: &[ShapeMove],
    ) -> EditResult<Vec<[CellFormulaReceipt; 2]>> {
        if moves.is_empty() {
            return Ok(Vec::new());
        }
        let mut txn = self.transact_for(context);
        let mut pending = Vec::with_capacity(moves.len() * 2);
        for shape_move in moves {
            let context_for_policy =
                CrdtMutationContext::new(&txn, &shape_move.page_id, &shape_move.shape_id)?;
            for (name, formula, gesture) in [
                ("PinX", shape_move.x.clone(), MutationGesture::MoveX),
                ("PinY", shape_move.y.clone(), MutationGesture::MoveY),
            ] {
                match decide_mutation(
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
                    MutationOutcome::Allowed { target, .. } => pending.push((
                        shape_move.page_id.clone(),
                        shape_move.shape_id.clone(),
                        target,
                        formula,
                    )),
                    MutationOutcome::Refused { reason }
                    | MutationOutcome::Unsupported { reason } => {
                        return Err(EditError::InvalidState(reason));
                    }
                }
            }
        }
        let mut receipts = Vec::with_capacity(moves.len());
        for pair in pending.chunks(2) {
            let [first, second] = pair else {
                return Err(EditError::InvalidState(
                    "cell group arity changed".to_owned(),
                ));
            };
            let mut pair_receipts = Vec::with_capacity(2);
            for (page_id, shape_id, target, formula) in [first, second] {
                let cell = cell_map(&mut txn, page_id, shape_id, target)?;
                let before = map_string(&cell, &txn, "formula");
                cell.insert(&mut txn, "formula", formula.as_str());
                pair_receipts.push(CellFormulaReceipt {
                    page_id: page_id.clone(),
                    shape_id: shape_id.clone(),
                    cell_name: target.cell_name.clone(),
                    before,
                    after: formula.clone(),
                });
            }
            let [first, second] = pair_receipts
                .try_into()
                .map_err(|_| EditError::InvalidState("cell group arity changed".to_owned()))?;
            receipts.push([first, second]);
        }
        Ok(receipts)
    }

    pub fn delete_shapes(
        &self,
        context: &EditCtx,
        deletes: &[ShapeDelete],
    ) -> EditResult<Vec<ShapeReceipt>> {
        if deletes.is_empty() {
            return Ok(Vec::new());
        }
        let mut txn = self.transact_for(context);
        for entry in deletes {
            let pages = txn
                .get_map(PAGES)
                .ok_or_else(|| EditError::InvalidState("missing pages map".to_owned()))?;
            let page = map_ref(&pages, &txn, &entry.page_id)?;
            let root_order = map_array(&page, &txn, "shapes")?;
            let sheets = txn
                .get_map(SHEETS)
                .ok_or_else(|| EditError::InvalidState("missing sheets map".to_owned()))?;
            if !shape_tree_entries(&sheets, &txn, &root_order)?
                .iter()
                .any(|shape| shape.id == entry.shape_id)
            {
                return Err(EditError::ShapeNotFound(entry.shape_id.clone()));
            }
        }
        for entry in deletes {
            let context_for_policy =
                CrdtMutationContext::new(&txn, &entry.page_id, &entry.shape_id)?;
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
        }
        let mut receipts = Vec::with_capacity(deletes.len());
        let mut remaining: Vec<(String, String)> = deletes
            .iter()
            .map(|entry| (entry.page_id.clone(), entry.shape_id.clone()))
            .collect();
        while !remaining.is_empty() {
            let mut best: Option<DeleteCandidate> = None;
            for (position, (page_id, shape_id)) in remaining.iter().enumerate() {
                let pages = txn
                    .get_map(PAGES)
                    .ok_or_else(|| EditError::InvalidState("missing pages map".to_owned()))?;
                let page = match map_ref(&pages, &txn, page_id) {
                    Ok(page) => page,
                    Err(_) => continue,
                };
                let root_order = match map_array(&page, &txn, "shapes") {
                    Ok(order) => order,
                    Err(_) => continue,
                };
                let sheets = txn
                    .get_map(SHEETS)
                    .ok_or_else(|| EditError::InvalidState("missing sheets map".to_owned()))?;
                let entries = shape_tree_entries(&sheets, &txn, &root_order)?;
                let Some(target) = entries.iter().position(|entry| entry.id == *shape_id) else {
                    continue;
                };
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
                let replace = match &best {
                    None => true,
                    Some((_, _, _, _, best_from, _, _)) => from > *best_from,
                };
                if replace {
                    best = Some((
                        position,
                        page_id.clone(),
                        shape_id.clone(),
                        order,
                        from,
                        depth,
                        removed,
                    ));
                }
            }
            let Some((position, page_id, shape_id, order, from, _, removed)) = best else {
                let (page_id, shape_id) = &remaining[0];
                if deletes
                    .iter()
                    .any(|entry| &entry.page_id == page_id && &entry.shape_id == shape_id)
                {
                    return Err(EditError::ShapeNotFound(shape_id.clone()));
                }
                remaining.remove(0);
                continue;
            };
            {
                let sheets = txn
                    .get_map(SHEETS)
                    .ok_or_else(|| EditError::InvalidState("missing sheets map".to_owned()))?;
                order.remove_range(&mut txn, from, 1);
                for id in &removed {
                    sheets.remove(&mut txn, id.as_str());
                }
                if let Some(stories) = txn.get_map(STORIES) {
                    for id in &removed {
                        stories.remove(&mut txn, id.as_str());
                    }
                }
            }
            receipts.push(ShapeReceipt {
                page_id: page_id.clone(),
                shape_id: shape_id.clone(),
                from_index: Some(from),
                to_index: None,
            });
            remaining.remove(position);
            remaining.retain(|(_, shape_id)| !removed.contains(shape_id));
        }
        receipts.sort_by(|left, right| {
            deletes
                .iter()
                .position(|entry| entry.shape_id == left.shape_id)
                .cmp(
                    &deletes
                        .iter()
                        .position(|entry| entry.shape_id == right.shape_id),
                )
        });
        Ok(receipts)
    }

    pub fn set_cell_formulas(
        &self,
        context: &EditCtx,
        writes: &[CellFormulaWrite],
    ) -> EditResult<Vec<CellFormulaReceipt>> {
        if writes.is_empty() {
            return Ok(Vec::new());
        }
        let mut txn = self.transact_for(context);
        let mut pending = Vec::with_capacity(writes.len());
        for write in writes {
            let context_for_policy =
                CrdtMutationContext::new(&txn, &write.page_id, &write.shape_id)?;
            match decide_mutation(
                &context_for_policy,
                context_for_policy.locator(CellLocator {
                    sheet: CellSheet::Page(0),
                    shape_id: None,
                    section: None,
                    section_index: None,
                    row: None,
                    cell_name: write.cell_name.clone(),
                }),
                gesture_for_cell(&write.cell_name),
                write.formula.clone(),
                &ParseLimits::default(),
            ) {
                MutationOutcome::Allowed { target, .. } => pending.push((
                    write.page_id.clone(),
                    write.shape_id.clone(),
                    target,
                    write.formula.clone(),
                )),
                MutationOutcome::Refused { reason } | MutationOutcome::Unsupported { reason } => {
                    return Err(EditError::InvalidState(reason));
                }
            }
        }
        let mut receipts = Vec::with_capacity(pending.len());
        for (page_id, shape_id, target, formula) in pending {
            let cell = cell_map(&mut txn, &page_id, &shape_id, &target)?;
            let before = map_string(&cell, &txn, "formula");
            cell.insert(&mut txn, "formula", formula.as_str());
            receipts.push(CellFormulaReceipt {
                page_id,
                shape_id,
                cell_name: target.cell_name,
                before,
                after: formula,
            });
        }
        Ok(receipts)
    }

    /// Atomically writes shape-data values as one undo entry.
    pub fn set_shape_data(
        &self,
        context: &EditCtx,
        page_id: &str,
        shape_id: &str,
        writes: &[ShapeDataWrite],
    ) -> EditResult<Vec<ShapeDataReceipt>> {
        if writes.is_empty() {
            return Ok(Vec::new());
        }
        let mut txn = self.transact_for(context);
        let policy = CrdtMutationContext::new(&txn, page_id, shape_id)?;
        let mut receipts: Vec<ShapeDataReceipt> = Vec::with_capacity(writes.len());
        let mut pending = Vec::with_capacity(writes.len());
        let mut targets: std::collections::HashMap<String, usize> =
            std::collections::HashMap::new();
        for write in writes {
            let locator = policy.locator(shape_data_locator(&write.row, write.section_index));
            let before = policy.current_formula(&locator).ok().flatten();
            let mut receipt = ShapeDataReceipt {
                page_id: page_id.to_owned(),
                shape_id: shape_id.to_owned(),
                row_name: match &write.row {
                    CellRow::Name(name) => Some(name.clone()),
                    CellRow::Index(_) => None,
                },
                row_index: match &write.row {
                    CellRow::Index(index) => Some(*index),
                    CellRow::Name(_) => None,
                },
                section_index: write.section_index,
                before: before.clone(),
                after: None,
                refusal: None,
            };
            let refusal = cell_map(&mut txn, page_id, shape_id, &locator)
                .err()
                .map(|error| error.to_string())
                .or_else(|| shape_data_refusal(&policy, write, before.as_deref()));
            if let Some(reason) = refusal {
                receipt.refusal = Some(reason);
                receipts.push(receipt);
                continue;
            }
            match decide_mutation(
                &policy,
                locator,
                gesture_for_cell(SHAPE_DATA_CELL),
                write.formula.clone(),
                &ParseLimits::default(),
            ) {
                MutationOutcome::Allowed { target, formula } => {
                    match cell_map(&mut txn, page_id, shape_id, &target) {
                        Ok(cell) => {
                            let key = locator_key(&target);
                            if let Some(&previous) = targets.get(&key) {
                                let reason =
                                    "shape-data writes converge on the same cell".to_owned();
                                receipts[previous].refusal = Some(reason.clone());
                                receipt.refusal = Some(reason);
                            } else {
                                targets.insert(key, receipts.len());
                                receipt.before = map_string(&cell, &txn, "formula");
                                receipt.after = Some(formula.clone());
                                pending.push((cell, formula, receipts.len()));
                            }
                        }
                        Err(error) => receipt.refusal = Some(error.to_string()),
                    }
                }
                MutationOutcome::Refused { reason } | MutationOutcome::Unsupported { reason } => {
                    receipt.refusal = Some(reason);
                }
            }
            receipts.push(receipt);
        }
        if receipts.iter().any(ShapeDataReceipt::refused) {
            for receipt in &mut receipts {
                receipt.after = None;
            }
            return Ok(receipts);
        }
        for (cell, formula, index) in pending {
            if receipts[index].before.as_deref() != Some(formula.as_str()) {
                cell.insert(&mut txn, "formula", formula.as_str());
            }
        }
        Ok(receipts)
    }

    /// What the mutation policy would do with a write to each locator, writing nothing.
    pub fn probe_cell_writes(
        &self,
        page_id: &str,
        shape_id: &str,
        probes: &[CellWriteQuery],
    ) -> EditResult<Vec<CellWriteProbe>> {
        if probes.is_empty() {
            return Ok(Vec::new());
        }
        let txn = self.doc.transact();
        let policy = CrdtMutationContext::new(&txn, page_id, shape_id)?;
        Ok(probes
            .iter()
            .map(|query| {
                let cell_name = query.locator.cell_name.clone();
                match decide_mutation(
                    &policy,
                    policy.locator(query.locator.clone()),
                    query
                        .gesture
                        .unwrap_or_else(|| gesture_for_cell(&cell_name)),
                    PROBE_FORMULA.to_owned(),
                    &ParseLimits::default(),
                ) {
                    MutationOutcome::Allowed { target, .. } => CellWriteProbe {
                        cell_name,
                        allowed: true,
                        target_cell_name: Some(target.cell_name),
                        refusal: None,
                        reason: None,
                    },
                    MutationOutcome::Refused { reason } => CellWriteProbe {
                        cell_name,
                        allowed: false,
                        target_cell_name: None,
                        refusal: Some(refusal_kind(&reason).to_owned()),
                        reason: Some(reason),
                    },
                    MutationOutcome::Unsupported { reason } => CellWriteProbe {
                        cell_name,
                        allowed: false,
                        target_cell_name: None,
                        refusal: Some("unsupported".to_owned()),
                        reason: Some(reason),
                    },
                }
            })
            .collect())
    }

    pub fn resize_loc_pin(
        &self,
        page_id: &str,
        shape_id: &str,
        width: f64,
        height: f64,
    ) -> EditResult<[f64; 2]> {
        if !width.is_finite() || !height.is_finite() || width <= 0.0 || height <= 0.0 {
            return Err(EditError::InvalidState(
                "invalid resize dimensions".to_owned(),
            ));
        }
        let txn = self.doc.transact();
        let mut references = CrdtMutationContext::new(&txn, page_id, shape_id)?.references;
        for (name, value) in [("Width", width), ("Height", height)] {
            references.cells.insert(
                name.to_owned(),
                Lookup::Found(vsdx_resolve::ResolvedCell {
                    cell: Cell {
                        name: name.to_owned(),
                        formula: Some(value.to_string()),
                        value: None,
                        unit: None,
                        del: false,
                        other_attrs: Vec::new(),
                    },
                    provenance: vsdx_resolve::Provenance::Local,
                }),
            );
        }
        let loc_pin = |name: &str, fallback: f64| -> EditResult<f64> {
            if !references.cells.contains_key(name) {
                return Ok(fallback);
            }
            evaluate_cached_formula(name, &references)
                .and_then(|value| value.parse::<f64>().ok())
                .filter(|value| value.is_finite())
                .ok_or_else(|| {
                    EditError::InvalidState(format!("cannot evaluate {name} for resize"))
                })
        };
        Ok([
            loc_pin("LocPinX", width / 2.0)?,
            loc_pin("LocPinY", height / 2.0)?,
        ])
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
        validate_shape_draft(draft, false)?;
        if let Some(master) = draft.master
            && !original_master_ids(&self.doc)?.contains(&master)
        {
            return Err(EditError::InvalidState(
                "shape draft references an unknown master".to_owned(),
            ));
        }
        let mut txn = self.transact_for(context);
        insert_shape(&mut txn, self.client_id, page_id, draft)
    }

    /// Rewrites a 1D connector's filed route from scene-space points in one transaction.
    pub fn set_connector_route(
        &self,
        context: &EditCtx,
        page_id: &str,
        shape_id: &str,
        points: &[(f64, f64)],
    ) -> EditResult<ConnectorRouteReceipt> {
        if points.len() < 2 || points.len() > MAX_CONNECTOR_ROUTE_POINTS {
            return Err(EditError::InvalidState(
                "connector route needs between 2 and 256 points".to_owned(),
            ));
        }
        if points.iter().any(|(x, y)| !x.is_finite() || !y.is_finite()) {
            return Err(EditError::InvalidState(
                "connector route points must be finite".to_owned(),
            ));
        }
        let mut txn = self.transact_for(context);
        let context_for_policy = CrdtMutationContext::new(&txn, page_id, shape_id)?;
        let sheets = txn
            .get_map(SHEETS)
            .ok_or_else(|| EditError::InvalidState("missing sheets map".to_owned()))?;
        let shape = map_ref(&sheets, &txn, shape_id)?;
        if map_string(&shape, &txn, "parentId").is_some() {
            return Err(EditError::InvalidState(
                "nested connector routes are not hand-routable".to_owned(),
            ));
        }
        let page_number = page_id
            .trim_start_matches("page:")
            .parse()
            .unwrap_or_default();
        let source_id = map_number(&shape, &txn, "sourceId").unwrap_or_default() as u32;
        let cells = map_map(&shape, &txn, "cells")?;
        let references = local_references(&cells, &txn)?;
        if !connector_route_shape(&references) {
            return Err(EditError::InvalidState(
                "shape is not a 1D connector".to_owned(),
            ));
        }
        let frame = connector_route_frame(&references).ok_or_else(|| {
            EditError::InvalidState("connector frame cells do not resolve".to_owned())
        })?;
        let mut local = points
            .iter()
            .map(|point| frame.unproject(point.0, point.1))
            .collect::<Option<Vec<_>>>()
            .ok_or_else(|| {
                EditError::InvalidState("connector frame is not invertible".to_owned())
            })?;
        local.dedup_by(|right, left| {
            (right.0 - left.0).abs() < ROUTE_POINT_EPSILON
                && (right.1 - left.1).abs() < ROUTE_POINT_EPSILON
        });
        if local.len() < 2 {
            return Err(EditError::InvalidState(
                "connector route needs two distinct points".to_owned(),
            ));
        }
        let mut rows = std::collections::BTreeMap::new();
        for value in cells.iter(&txn).filter_map(|(_, value)| match value {
            Out::YMap(cell) => Some(cell),
            _ => None,
        }) {
            let locator = cell_locator(&value, &txn, 0, 0)?;
            if locator.section.as_deref() != Some("Geometry") {
                continue;
            }
            if locator.section_index.is_some() {
                return Err(EditError::InvalidState(
                    "indexed connector Geometry sections are not hand-routable".to_owned(),
                ));
            }
            let Some(CellRow::Index(index)) = locator.row else {
                continue;
            };
            let entry = rows.entry(index).or_insert(GeometryRow {
                row_type: map_string(&value, &txn, "rowType"),
                has_x: false,
                has_y: false,
            });
            match locator.cell_name.as_str() {
                "X" => entry.has_x = true,
                "Y" => entry.has_y = true,
                _ => {}
            }
        }
        let filed = rows.keys().copied().collect::<Vec<_>>();
        let last = *local.last().expect("route has at least two points");
        let mut planned = Vec::new();
        for (slot, point) in local
            .iter()
            .copied()
            .chain(std::iter::repeat_n(
                last,
                filed.len().saturating_sub(local.len()),
            ))
            .enumerate()
        {
            let index = route_row_index(&filed, slot);
            let expected = if slot == 0 { "MoveTo" } else { "LineTo" };
            let filed_type = rows.get(&index).and_then(|row| row.row_type.clone());
            if filed_type
                .as_deref()
                .is_some_and(|actual| actual != expected)
            {
                return Err(EditError::InvalidState(
                    "connector Geometry is not a single straight run".to_owned(),
                ));
            }
            let row_type = filed_type.unwrap_or_else(|| expected.to_owned());
            for (name, value) in [("X", point.0), ("Y", point.1)] {
                if slot >= local.len() && !rows.get(&index).is_some_and(|row| row.has(name)) {
                    continue;
                }
                let locator = CellLocator {
                    sheet: CellSheet::Page(page_number),
                    shape_id: Some(source_id),
                    section: Some("Geometry".into()),
                    section_index: None,
                    row: Some(CellRow::Index(index)),
                    cell_name: name.into(),
                };
                let formula = decimal(value);
                allow_route_cell(&context_for_policy, locator.clone(), &formula)?;
                planned.push((locator, formula, row_type.clone()));
            }
        }
        let prepared = planned
            .iter()
            .map(|(locator, formula, row_type)| {
                match cell_map(&mut txn, page_id, shape_id, locator) {
                    Ok(cell) => Ok((Some(cell), locator, formula, row_type)),
                    Err(EditError::CellNotFound(_)) => Ok((None, locator, formula, row_type)),
                    Err(error) => Err(error),
                }
            })
            .collect::<EditResult<Vec<_>>>()?;
        for (cell, locator, formula, row_type) in prepared {
            match cell {
                Some(cell) => {
                    cell.insert(&mut txn, "formula", formula.as_str());
                }
                None => seed_route_cell(&cells, &mut txn, locator, formula, row_type),
            }
        }
        Ok(ConnectorRouteReceipt {
            page_id: page_id.to_owned(),
            shape_id: shape_id.to_owned(),
            points: local.len() as u32,
        })
    }

    /** Adds a shape with initial text in one transaction, so paste stays one undo step. */
    pub fn add_shape_with_text(
        &self,
        context: &EditCtx,
        page_id: &str,
        draft: &ShapeDraft,
        text: String,
    ) -> EditResult<ShapeReceipt> {
        validate_shape_draft(draft, true)?;
        validate_story_text(&text)?;
        let mut txn = self.transact_for(context);
        let receipt = insert_shape(&mut txn, self.client_id, page_id, draft)?;
        txn.get_or_insert_map(STORIES)
            .insert(&mut txn, receipt.shape_id.as_str(), text.as_str());
        Ok(receipt)
    }

    /// Internal glue of the subtree rooted at a shape, addressed by copy sources.
    pub fn subtree_glue(&self, page_id: &str, shape_id: &str) -> EditResult<Vec<ShapeTreeGlue>> {
        let txn = self.doc.transact();
        let pages = required_map(&txn, PAGES)?;
        map_ref(&pages, &txn, page_id)?;
        let sheets = required_map(&txn, SHEETS)?;
        let shape = map_ref(&sheets, &txn, shape_id)
            .map_err(|_| EditError::ShapeNotFound(shape_id.to_owned()))?;
        if map_string(&shape, &txn, "pageId").as_deref() != Some(page_id) {
            return Err(EditError::ShapeNotFound(shape_id.to_owned()));
        }
        let mut members = HashSet::new();
        let mut originals = std::collections::BTreeMap::new();
        let mut pending = vec![(shape_id.to_owned(), 1)];
        while let Some((id, depth)) = pending.pop() {
            if depth > MAX_SHAPE_NESTING {
                return Err(EditError::InvalidState(
                    "shape nesting exceeds maximum depth".to_owned(),
                ));
            }
            if !members.insert(id.clone()) {
                continue;
            }
            let node = map_ref(&sheets, &txn, &id)?;
            if shape_origin(&node, &txn)? == ShapeOrigin::Original
                && let Some(source_id) = map_number(&node, &txn, "sourceId")
            {
                originals.insert(source_id as u32, id.clone());
            }
            if let Some(Out::YArray(children)) = node.get(&txn, "shapes") {
                for index in 0..children.len(&txn) {
                    let child = array_string(&children, &txn, index).ok_or_else(|| {
                        EditError::InvalidState("shape order contains non-string".to_owned())
                    })?;
                    pending.push((child, depth + 1));
                }
            }
        }
        let part = map_string(&map_ref(&pages, &txn, page_id)?, &txn, "sourcePartPath");
        let package = match originals.is_empty() {
            true => None,
            false => Some(original_package_from_doc(&self.doc)?),
        };
        let sheet = package
            .as_ref()
            .zip(part.as_ref())
            .and_then(|(package, part)| package.page_contents.get(part));
        connect::subtree_glue(&txn, page_id, &members, sheet, &originals)
    }

    /// Pastes a copied group subtree with fresh identities in one transaction.
    pub fn add_shape_tree(
        &self,
        context: &EditCtx,
        page_id: &str,
        draft: &ShapeTreeDraft,
    ) -> EditResult<ShapeReceipt> {
        let nodes = flatten_tree_draft(draft)?;
        let keys = tree_copy_keys(&nodes)?;
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
        let mut source_page_id: Option<String> = None;
        let mut source_page_numeric: Option<u32> = None;
        for node in &nodes {
            let source_id = node
                .draft
                .source_shape_id
                .as_ref()
                .ok_or_else(|| EditError::InvalidState("missing copy source".to_owned()))?;
            match sheets.get(&txn, source_id.as_str()) {
                Some(Out::YMap(source)) => {
                    if let Some(reason) = map_string(&source, &txn, "copyRefusal") {
                        return Err(EditError::InvalidState(format!(
                            "cannot copy shape {source_id} with {reason}"
                        )));
                    }
                    let owner = map_string(&source, &txn, "pageId").ok_or_else(|| {
                        EditError::InvalidState("copy source is missing its page".to_owned())
                    })?;
                    match &source_page_id {
                        Some(expected) if *expected == owner => {}
                        Some(_) => {
                            return Err(EditError::InvalidState(
                                "paste sources span multiple pages".to_owned(),
                            ));
                        }
                        None => source_page_id = Some(owner),
                    }
                    // Prefer the stored original page so copies of copies resolve their text.
                    let numeric = match map_u32(&source, &txn, "copySourcePageId")? {
                        Some(page) => page,
                        None => session_page_number(source_page_id.as_deref().unwrap_or_default())?,
                    };
                    match source_page_numeric {
                        Some(expected) if expected == numeric => {}
                        Some(_) => {
                            return Err(EditError::InvalidState(
                                "paste sources span multiple pages".to_owned(),
                            ));
                        }
                        None => source_page_numeric = Some(numeric),
                    }
                }
                _ => {
                    if let Some(reason) = &node.draft.copy_refusal {
                        return Err(EditError::InvalidState(format!(
                            "cannot copy shape {source_id} with {reason}"
                        )));
                    }
                }
            }
        }
        if source_page_numeric.is_none() {
            source_page_numeric = draft_source_page(&nodes)?;
        }
        if let Some(numeric) = source_page_numeric {
            map_ref(&pages, &txn, &format!("page:{numeric}"))
                .map_err(|_| EditError::InvalidState("paste source page is missing".to_owned()))?;
        }
        if sheets.len(&txn) as usize + nodes.len() > ParseLimits::default().max_shapes {
            return Err(EditError::InvalidState(
                "shape count exceeds maximum".to_owned(),
            ));
        }
        let source_bound = map_u32(&page, &txn, "maxSourceId")?
            .ok_or_else(|| EditError::InvalidState("missing page source ID bound".to_owned()))?;
        let allocated = materialized_source_ids(&sheets, &txn, &order, source_bound)?;
        let mut largest = allocated.values().copied().max().unwrap_or(source_bound);
        let mut eager = Vec::with_capacity(nodes.len());
        for _ in &nodes {
            largest = largest.checked_add(1).ok_or_else(|| {
                EditError::InvalidState("cannot allocate a materialized source ID".to_owned())
            })?;
            eager.push(largest);
        }
        let remap = keys
            .iter()
            .zip(&eager)
            .filter_map(|(key, eager)| key.map(|key| (key, *eager)))
            .collect::<std::collections::BTreeMap<_, _>>();
        if source_page_numeric != Some(session_page_number(page_id)?)
            && nodes.iter().any(|node| {
                node.draft.cells.iter().any(|cell| {
                    cell.formula.as_deref().is_some_and(|formula| {
                        sheet_ref_spans(formula)
                            .iter()
                            .any(|span| !remap.contains_key(&span.2))
                    })
                })
            })
        {
            return Err(EditError::InvalidState(
                "paste references a shape outside the copy on another page".to_owned(),
            ));
        }
        let sequence = txn.state_vector().get(&yrs::ClientID::new(self.client_id));
        let prefix = format!("{page_id}:shape:added:{}:{sequence}:", self.client_id);
        let fresh = (0..nodes.len())
            .map(|position| format!("{prefix}{position}"))
            .collect::<Vec<_>>();
        let locate = |source: &str| {
            nodes
                .iter()
                .position(|node| {
                    node.draft
                        .source_shape_id
                        .as_deref()
                        .is_some_and(|id| id == source)
                })
                .map(|position| fresh[position].as_str())
        };
        let mut seen_endpoints = HashSet::new();
        let mut carried = Vec::new();
        for glue in &draft.glue {
            let endpoint =
                connect::GlueEndpoint::parse(glue.endpoint.as_str()).ok_or_else(|| {
                    EditError::InvalidState("connector glue has no endpoint".to_owned())
                })?;
            if !connect::valid_glue_target(&glue.to_cell)
                || !connect::glue_text_valid(&glue.to_cell)
            {
                return Err(EditError::InvalidState(
                    "connector glue has an invalid target cell".to_owned(),
                ));
            }
            if !seen_endpoints.insert((glue.connector_source.clone(), endpoint)) {
                return Err(EditError::InvalidState(
                    "connector glue duplicates an endpoint".to_owned(),
                ));
            }
            let (Some(connector), Some(target)) =
                (locate(&glue.connector_source), locate(&glue.target_source))
            else {
                continue;
            };
            carried.push((endpoint, connector, target, glue.to_cell.clone()));
        }
        for (position, node) in nodes.iter().enumerate() {
            let shape = sheets.insert(&mut txn, fresh[position].as_str(), MapPrelim::default());
            shape.insert(&mut txn, "id", fresh[position].as_str());
            shape.insert(&mut txn, "pageId", page_id);
            shape.insert(&mut txn, "sourceId", eager[position] as f64);
            shape.insert(&mut txn, "origin", "added");
            if let Some(key) = keys[position] {
                shape.insert(&mut txn, "copySourceId", key as f64);
            }
            if let Some(page) = source_page_numeric {
                shape.insert(&mut txn, "copySourcePageId", page as f64);
            }
            if let Some(parent) = node.parent {
                shape.insert(&mut txn, "parentId", fresh[parent].as_str());
            }
            if let Some(name) = &node.draft.name {
                shape.insert(&mut txn, "name", name.as_str());
            }
            let cells = shape.insert(&mut txn, "cells", MapPrelim::default());
            for cell in &node.draft.cells {
                let rewritten = cell
                    .formula
                    .as_deref()
                    .map(|formula| remap_sheet_refs(formula, &remap));
                seed_cell(
                    &cells,
                    &mut txn,
                    &cell.locator,
                    rewritten.as_deref().or(cell.formula.as_deref()),
                    cell.value.as_deref(),
                    cell.row_type.as_deref(),
                );
            }
            shape.insert(&mut txn, "shapes", ArrayPrelim::default());
            match node.parent {
                Some(parent) => {
                    let owner = map_ref(&sheets, &txn, &fresh[parent])?;
                    map_array(&owner, &txn, "shapes")?
                        .push_back(&mut txn, fresh[position].as_str());
                }
                None => {
                    order.push_back(&mut txn, fresh[position].as_str());
                }
            }
            txn.get_or_insert_map(STORIES).insert(
                &mut txn,
                fresh[position].as_str(),
                node.draft.text.as_str(),
            );
        }
        let connects = txn.get_or_insert_map(CONNECTS);
        for (endpoint, connector, target, cell) in carried {
            let key = connect::glue_key(connector, endpoint);
            let entry = connects.insert(&mut txn, key.as_str(), MapPrelim::default());
            entry.insert(&mut txn, "id", key.as_str());
            entry.insert(&mut txn, "pageId", page_id);
            entry.insert(&mut txn, "connectorId", connector);
            entry.insert(&mut txn, "endpoint", endpoint.name());
            entry.insert(&mut txn, "targetId", target);
            entry.insert(&mut txn, "toCell", cell.as_str());
        }
        Ok(ShapeReceipt {
            page_id: page_id.to_owned(),
            shape_id: fresh.into_iter().next().ok_or_else(|| {
                EditError::InvalidState("paste draft is missing its copy source".to_owned())
            })?,
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
        if let Some(stories) = txn.get_map(STORIES) {
            for id in &removed {
                stories.remove(&mut txn, id.as_str());
            }
        }
        Ok(ShapeReceipt {
            page_id: page_id.to_owned(),
            shape_id: shape_id.to_owned(),
            from_index: Some(from),
            to_index: None,
        })
    }

    fn set_cell_formula_group<const N: usize>(
        &self,
        context: &EditCtx,
        page_id: &str,
        shape_id: &str,
        cells: [(&str, String, MutationGesture); N],
    ) -> EditResult<[CellFormulaReceipt; N]> {
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
        let targets = cells
            .into_iter()
            .map(decide)
            .collect::<EditResult<Vec<_>>>()?;
        for (index, (target, _)) in targets.iter().enumerate() {
            if targets[..index].iter().any(|(seen, _)| seen == target) {
                return Err(EditError::InvalidState(format!(
                    "redirects converge on {} more than once",
                    target.cell_name
                )));
            }
        }
        let prepared = targets
            .into_iter()
            .map(|(target, formula)| {
                let cell = cell_map(&mut txn, page_id, shape_id, &target)?;
                Ok((target, formula, cell))
            })
            .collect::<EditResult<Vec<_>>>()?;
        Ok(std::array::from_fn(|index| {
            let (target, formula, cell) = &prepared[index];
            let before = map_string(cell, &txn, "formula");
            cell.insert(&mut txn, "formula", formula.as_str());
            CellFormulaReceipt {
                page_id: page_id.to_owned(),
                shape_id: shape_id.to_owned(),
                cell_name: target.cell_name.clone(),
                before,
                after: formula.clone(),
            }
        }))
    }
}

fn insert_shape(
    txn: &mut TransactionMut<'_>,
    client_id: u64,
    page_id: &str,
    draft: &ShapeDraft,
) -> EditResult<ShapeReceipt> {
    let pages = txn
        .get_map(PAGES)
        .ok_or_else(|| EditError::InvalidState("missing pages map".to_owned()))?;
    let page = map_ref(&pages, txn, page_id)?;
    let order = map_array(&page, txn, "shapes")?;
    let index = order.len(txn);
    let sheets = txn
        .get_map(SHEETS)
        .ok_or_else(|| EditError::InvalidState("missing sheets map".to_owned()))?;
    let id_prefix = format!("{page_id}:shape:added:{}:", client_id);
    if sheets.len(txn) as usize >= ParseLimits::default().max_shapes {
        return Err(EditError::InvalidState(
            "shape count exceeds maximum".to_owned(),
        ));
    }
    let source_bound = map_u32(&page, txn, "maxSourceId")?
        .ok_or_else(|| EditError::InvalidState("missing page source ID bound".to_owned()))?;
    let allocated = materialized_source_ids(&sheets, txn, &order, source_bound)?;
    let largest = allocated.values().copied().max().unwrap_or(source_bound);
    largest.checked_add(1).ok_or_else(|| {
        EditError::InvalidState("cannot allocate a materialized source ID".to_owned())
    })?;
    let mut sequence = txn.state_vector().get(&yrs::ClientID::new(client_id));
    let id = loop {
        let candidate = format!("{id_prefix}{sequence}");
        if sheets.get(&*txn, candidate.as_str()).is_none() {
            break candidate;
        }
        sequence += 1;
    };
    let shape = sheets.insert(txn, id.as_str(), MapPrelim::default());
    shape.insert(txn, "id", id.as_str());
    shape.insert(txn, "pageId", page_id);
    shape.insert(txn, "sourceId", 0.0);
    shape.insert(txn, "origin", "added");
    if let Some(name) = &draft.name {
        shape.insert(txn, "name", name.as_str());
    }
    if let Some(master) = draft.master {
        shape.insert(txn, "master", f64::from(master));
    }
    let cells = shape.insert(txn, "cells", MapPrelim::default());
    shape.insert(txn, "shapes", ArrayPrelim::default());
    for cell in &draft.cells {
        seed_cell(
            &cells,
            txn,
            &cell.locator,
            cell.formula.as_deref(),
            cell.value.as_deref(),
            cell.row_type.as_deref(),
        );
    }
    order.push_back(txn, id.as_str());
    Ok(ShapeReceipt {
        page_id: page_id.to_owned(),
        shape_id: id,
        from_index: None,
        to_index: Some(index),
    })
}

fn validate_draft_master(package: &vsdx_parse::VsdxPackage, draft: &ShapeDraft) -> EditResult<()> {
    if let Some(master) = draft.master
        && !package.master_sheets.contains_key(&master)
    {
        return Err(EditError::InvalidState(
            "shape draft references an unknown master".to_owned(),
        ));
    }
    Ok(())
}

/** Paste reuses trusted cached values for formula-less cells; other drafts stay formula-only. */
fn validate_shape_draft(draft: &ShapeDraft, allow_values: bool) -> EditResult<()> {
    if draft.master == Some(0) {
        return Err(EditError::InvalidState(
            "shape draft references an invalid master".to_owned(),
        ));
    }
    validate_draft_cells(&draft.name, &draft.cells, allow_values)
}

fn validate_draft_cells(
    name: &Option<String>,
    cells: &[CellSnapshot],
    allow_values: bool,
) -> EditResult<()> {
    let limits = ParseLimits::default();
    let text = |value: &str| -> EditResult<()> {
        if value.len() > limits.max_attribute_bytes || value.chars().any(|c| !matches!(c, '\t' | '\r' | '\n' | '\u{20}'..='\u{d7ff}' | '\u{e000}'..='\u{fffd}' | '\u{10000}'..='\u{10ffff}')) {
            return Err(EditError::InvalidState("draft contains invalid XML attribute text".to_owned()));
        }
        Ok(())
    };
    if let Some(name) = name {
        text(name)?;
    }
    if cells.len() > limits.max_cells {
        return Err(EditError::InvalidState(
            "draft cell count exceeds maximum".to_owned(),
        ));
    }
    let mut locators = HashSet::new();
    let mut row_types = std::collections::BTreeMap::new();
    for cell in cells {
        let locator = &cell.locator;
        if !allow_values && cell.value.is_some() {
            return Err(EditError::InvalidState(
                "shape draft cells must not contain value".to_owned(),
            ));
        }
        if let Some(value) = &cell.value {
            text(value)?;
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

/// Carries a stored document forward; an unknown version is left for [`validate_schema`].
pub(crate) fn migrate_doc(doc: &Doc) -> EditResult<()> {
    let rewrites = {
        let txn = doc.transact();
        let meta = required_map(&txn, META)?;
        if map_number(&meta, &txn, "schemaVersion") != Some(TOKEN_STORY_SCHEMA_VERSION) {
            return Ok(());
        }
        let stories = required_map(&txn, STORIES)?;
        stories
            .iter(&txn)
            .filter_map(|(key, value)| {
                let Out::Any(Any::String(value)) = value else {
                    return None;
                };
                serde_json::from_str::<Vec<vsdx_resolve::ResolvedTextToken>>(&value)
                    .ok()
                    .map(|tokens| (key.to_owned(), plain_text(&tokens)))
            })
            .collect::<Vec<_>>()
    };
    let mut txn = doc.transact_mut_with(crate::MIGRATE_ORIGIN);
    let stories = required_map(&txn, STORIES)?;
    for (key, text) in &rewrites {
        stories.insert(&mut txn, key.as_str(), text.as_str());
    }
    required_map(&txn, META)?.insert(&mut txn, "schemaVersion", SCHEMA_VERSION);
    Ok(())
}

struct FlatTreeNode<'a> {
    draft: &'a ShapeTreeDraft,
    parent: Option<usize>,
}

/// Pre-order flattening with depth, cell, text and provenance checks.
fn flatten_tree_draft(draft: &ShapeTreeDraft) -> EditResult<Vec<FlatTreeNode<'_>>> {
    let mut nodes = Vec::new();
    let mut pending = vec![(draft, None, 1, true)];
    while let Some((draft, parent, depth, is_root)) = pending.pop() {
        if depth > MAX_SHAPE_NESTING {
            return Err(EditError::InvalidState(
                "shape nesting exceeds maximum depth".to_owned(),
            ));
        }
        if !is_root && !draft.glue.is_empty() {
            return Err(EditError::InvalidState(
                "paste glue belongs on the copied root".to_owned(),
            ));
        }
        validate_draft_cells(&draft.name, &draft.cells, true)?;
        validate_story_text(&draft.text)?;
        let index = nodes.len();
        nodes.push(FlatTreeNode { draft, parent });
        for child in draft.children.iter().rev() {
            pending.push((child, Some(index), depth + 1, false));
        }
    }
    let mut sources = HashSet::new();
    for node in &nodes {
        let source = node
            .draft
            .source_shape_id
            .as_ref()
            .filter(|source| !source.is_empty())
            .ok_or_else(|| {
                EditError::InvalidState("paste draft is missing its copy source".to_owned())
            })?;
        if !sources.insert(source) {
            return Err(EditError::InvalidState(
                "paste draft references its copy source twice".to_owned(),
            ));
        }
    }
    Ok(nodes)
}

/// Numeric identity each pasted node answers to, for remapping `Sheet.N!` references.
fn tree_copy_keys(nodes: &[FlatTreeNode<'_>]) -> EditResult<Vec<Option<u32>>> {
    let mut referenced = HashSet::new();
    for node in nodes {
        for cell in &node.draft.cells {
            if let Some(formula) = &cell.formula {
                referenced.extend(sheet_ref_spans(formula).iter().map(|span| span.2));
            }
        }
    }
    let mut keys = Vec::with_capacity(nodes.len());
    let mut seen = HashSet::new();
    for node in nodes {
        let key = match (node.draft.source_id, node.draft.copy_source_id) {
            (Some(live), _) if referenced.contains(&live) => Some(live),
            (_, Some(stored)) if referenced.contains(&stored) => Some(stored),
            (_, Some(stored)) => Some(stored),
            (Some(live), None) => Some(live),
            (None, None) => None,
        };
        if let Some(key) = key
            && !seen.insert(key)
        {
            return Err(EditError::InvalidState(
                "paste draft contains duplicate copy sources".to_owned(),
            ));
        }
        keys.push(key);
    }
    Ok(keys)
}

/// Numeric source page behind a `page:N` session ID.
fn session_page_number(page_id: &str) -> EditResult<u32> {
    page_id
        .strip_prefix("page:")
        .and_then(|value| value.parse::<u32>().ok())
        .ok_or_else(|| EditError::InvalidState("copy source is missing its page".to_owned()))
}

/// Single source page carried by a detached paste draft, if any.
fn draft_source_page(nodes: &[FlatTreeNode<'_>]) -> EditResult<Option<u32>> {
    let mut page = None;
    for node in nodes {
        let Some(candidate) = node.draft.copy_source_page_id else {
            continue;
        };
        match page {
            Some(expected) if expected == candidate => {}
            Some(_) => {
                return Err(EditError::InvalidState(
                    "paste sources span multiple pages".to_owned(),
                ));
            }
            None => page = Some(candidate),
        }
    }
    Ok(page)
}

/// `Sheet.N!` spans outside string literals, with the digit range and the target ID.
fn sheet_ref_spans(formula: &str) -> Vec<(usize, usize, u32)> {
    let bytes = formula.as_bytes();
    let mut spans = Vec::new();
    let mut index = 0;
    let mut in_string = false;
    while index < bytes.len() {
        let byte = bytes[index];
        if byte == b'"' {
            if in_string && bytes.get(index + 1) == Some(&b'"') {
                index += 2;
                continue;
            }
            in_string = !in_string;
            index += 1;
            continue;
        }
        if !in_string
            && (byte == b'S' || byte == b's')
            && let Some((start, end, id)) = sheet_ref_at(bytes, index)
        {
            spans.push((start, end, id));
            index = end + 1;
            continue;
        }
        index += 1;
    }
    spans
}

fn sheet_ref_at(bytes: &[u8], index: usize) -> Option<(usize, usize, u32)> {
    if index > 0
        && matches!(
            bytes[index - 1],
            b'0'..=b'9' | b'A'..=b'Z' | b'a'..=b'z' | b'.' | b'!' | b'_'
        )
    {
        return None;
    }
    let prefix = bytes.get(index..index + 6)?;
    if !prefix.eq_ignore_ascii_case(b"sheet.") {
        return None;
    }
    let mut end = index + 6;
    while bytes.get(end).is_some_and(|byte| byte.is_ascii_digit()) {
        end += 1;
    }
    if end == index + 6 || bytes.get(end) != Some(&b'!') {
        return None;
    }
    let id: u32 = std::str::from_utf8(&bytes[index + 6..end])
        .ok()?
        .parse()
        .ok()?;
    Some((index + 6, end, id))
}

/// Rewrites `Sheet.N!` targets through an old-to-new source ID map.
fn remap_sheet_refs<'a>(
    formula: &'a str,
    remap: &std::collections::BTreeMap<u32, u32>,
) -> std::borrow::Cow<'a, str> {
    if remap.is_empty() {
        return std::borrow::Cow::Borrowed(formula);
    }
    let spans = sheet_ref_spans(formula);
    if spans.iter().all(|span| !remap.contains_key(&span.2)) {
        return std::borrow::Cow::Borrowed(formula);
    }
    let mut output = String::with_capacity(formula.len());
    let mut cursor = 0;
    for (start, end, id) in spans {
        if let Some(target) = remap.get(&id) {
            output.push_str(&formula[cursor..start]);
            output.push_str(&target.to_string());
            cursor = end;
        }
    }
    output.push_str(&formula[cursor..]);
    std::borrow::Cow::Owned(output)
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
            if shape.get(&txn, "copyRefusal").is_some()
                && map_string(&shape, &txn, "copyRefusal").is_none_or(|refusal| refusal.is_empty())
            {
                return Err(EditError::InvalidState(
                    "shape copy refusal is not a non-empty string".to_owned(),
                ));
            }
            map_u32(&shape, &txn, "copySourceId")?;
            map_u32(&shape, &txn, "copySourcePageId")?;
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
    connect::glue_records(&txn)?;
    validate_story_records(&txn)?;
    Ok(())
}

fn validate_story_records<T: ReadTxn>(txn: &T) -> EditResult<()> {
    let Some(stories) = txn.get_map(STORIES) else {
        return Ok(());
    };
    let sheets = required_map(txn, SHEETS)?;
    for (key, value) in stories.iter(txn) {
        if !matches!(value, Out::Any(Any::String(_))) {
            return Err(EditError::InvalidState(
                "shape text is not a string".to_owned(),
            ));
        }
        if sheets.get(txn, key).is_none() {
            return Err(EditError::InvalidState(
                "shape text references a missing shape".to_owned(),
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

pub(crate) fn validate_remote_update(before: &Doc, staged: &Doc) -> EditResult<()> {
    validate_schema(staged)?;
    validate_immutable_metadata(before, staged)?;
    validate_session_topology(before, staged)?;
    validate_formula_mutations(before, staged)?;
    validate_remote_text(before, staged)?;
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
    validate_remote_masters(before, staged)?;
    connect::validate_remote_glue(before, staged)?;
    Ok(())
}

/// Rejects remote shapes whose master the opened package does not define.
fn validate_remote_masters(before: &Doc, staged: &Doc) -> EditResult<()> {
    let before_txn = before.transact();
    let staged_txn = staged.transact();
    let before_sheets = required_map(&before_txn, SHEETS)?;
    let staged_sheets = required_map(&staged_txn, SHEETS)?;
    let mut added = Vec::new();
    for (shape_id, staged_shape) in staged_sheets.iter(&staged_txn) {
        if before_sheets.get(&before_txn, shape_id).is_some() {
            continue;
        }
        let Out::YMap(staged_shape) = staged_shape else {
            continue;
        };
        if let Some(master) = map_u32(&staged_shape, &staged_txn, "master")? {
            added.push((shape_id.to_owned(), master));
        }
    }
    if added.is_empty() {
        return Ok(());
    }
    let known = original_master_ids(before)?;
    for (shape_id, master) in added {
        if !known.contains(&master) {
            return Err(EditError::InvalidState(format!(
                "remote update adds shape {shape_id} with unknown master {master}"
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
                "LockRotate",
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
        for key in [
            "id",
            "pageId",
            "sourceId",
            "origin",
            "parentId",
            "master",
            "copySourceId",
            "copySourcePageId",
            "copyRefusal",
        ] {
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
        let session_owned = shape_origin(&staged_shape, &staged_txn)? == ShapeOrigin::Added;
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
            if !session_owned && cell.get(&staged_txn, "value").is_some() {
                return Err(EditError::InvalidState(
                    "remote update adds untrusted cached cell value".to_owned(),
                ));
            }
        }
    }
    Ok(())
}

type ShapeIdentity = (Option<String>, Option<String>, Option<String>, Option<f64>);

/// Authorizes remote plain-text writes through the text-edit mutation policy.
fn validate_remote_text(before: &Doc, staged: &Doc) -> EditResult<()> {
    let before_txn = before.transact();
    let staged_txn = staged.transact();
    let before_sheets = required_map(&before_txn, SHEETS)?;
    for (shape_id, _) in required_map(&staged_txn, SHEETS)?.iter(&staged_txn) {
        let after_text = story_text(&staged_txn, shape_id);
        let Some(Out::YMap(before_shape)) = before_sheets.get(&before_txn, shape_id) else {
            validate_story_text(&after_text)?;
            continue;
        };
        if story_text(&before_txn, shape_id) == after_text {
            continue;
        }
        validate_story_text(&after_text)?;
        let page_id = map_string(&before_shape, &before_txn, "pageId")
            .ok_or_else(|| EditError::InvalidState("shape is missing its page".to_owned()))?;
        let context = CrdtMutationContext::new(&before_txn, &page_id, shape_id)?;
        match decide_mutation(
            &context,
            context.locator(CellLocator {
                sheet: CellSheet::Page(0),
                shape_id: None,
                section: None,
                section_index: None,
                row: None,
                cell_name: "Text".to_owned(),
            }),
            MutationGesture::TextEdit,
            after_text,
            &ParseLimits::default(),
        ) {
            MutationOutcome::Allowed { .. } => {}
            MutationOutcome::Refused { reason } | MutationOutcome::Unsupported { reason } => {
                return Err(EditError::InvalidState(reason));
            }
        }
    }
    Ok(())
}

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
                    "LockRotate",
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
        "LockRotate" => Some("Angle"),
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
    let master = map_u32(&shape, txn, "master")?;
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
            master,
            cells: snapshots,
            children: Vec::new(),
            copy_source_id: map_u32(&shape, txn, "copySourceId")?,
            copy_source_page_id: map_u32(&shape, txn, "copySourcePageId")?,
            copy_refusal: map_string(&shape, txn, "copyRefusal"),
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
        master,
        cells: snapshots,
        children,
        copy_source_id: map_u32(&shape, txn, "copySourceId")?,
        copy_source_page_id: map_u32(&shape, txn, "copySourcePageId")?,
        copy_refusal: map_string(&shape, txn, "copyRefusal"),
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

pub(crate) fn loc_pin_at_size<T: ReadTxn>(
    txn: &T,
    shape_id: &str,
    width: f64,
    height: f64,
) -> EditResult<(f64, f64)> {
    let sheets = required_map(txn, SHEETS)?;
    let shape = map_ref(&sheets, txn, shape_id)?;
    let cells = map_map(&shape, txn, "cells")?;
    let base = local_references(&cells, txn)?;
    let overrides = SizeOverrideRefs {
        base: &base,
        width: width.to_string(),
        height: height.to_string(),
    };
    Ok((
        loc_pin_component(&base, &overrides, "LocPinX", width),
        loc_pin_component(&base, &overrides, "LocPinY", height),
    ))
}

struct SizeOverrideRefs<'a> {
    base: &'a vsdx_resolve::ResolvedShape,
    width: String,
    height: String,
}

impl References for SizeOverrideRefs<'_> {
    fn formula(&self, name: &str) -> Option<&str> {
        if is_size_cell(name) {
            None
        } else {
            self.base.formula(name)
        }
    }

    fn value(&self, name: &str) -> Option<(&str, Option<&str>)> {
        self.size_value(name).or_else(|| self.base.value(name))
    }

    fn formula_in(&self, sheet: Option<u32>, name: &str) -> Option<&str> {
        if sheet.is_none() && is_size_cell(name) {
            None
        } else {
            self.base.formula_in(sheet, name)
        }
    }

    fn value_in(&self, sheet: Option<u32>, name: &str) -> Option<(&str, Option<&str>)> {
        if sheet.is_none() {
            self.size_value(name)
                .or_else(|| self.base.value_in(sheet, name))
        } else {
            self.base.value_in(sheet, name)
        }
    }

    fn formula_in_scoped(
        &self,
        sheet: Option<u32>,
        scope: Option<&str>,
        name: &str,
    ) -> Option<&str> {
        if sheet.is_none() && scope.is_none() && is_size_cell(name) {
            None
        } else {
            self.base.formula_in_scoped(sheet, scope, name)
        }
    }

    fn value_in_scoped(
        &self,
        sheet: Option<u32>,
        scope: Option<&str>,
        name: &str,
    ) -> Option<(&str, Option<&str>)> {
        if sheet.is_none() && scope.is_none() {
            self.size_value(name)
                .or_else(|| self.base.value_in_scoped(sheet, scope, name))
        } else {
            self.base.value_in_scoped(sheet, scope, name)
        }
    }

    fn reference_key(&self, sheet: Option<u32>, name: &str) -> String {
        self.base.reference_key(sheet, name)
    }
}

impl SizeOverrideRefs<'_> {
    fn size_value(&self, name: &str) -> Option<(&str, Option<&str>)> {
        if name.eq_ignore_ascii_case("Width") {
            Some((&self.width, None))
        } else if name.eq_ignore_ascii_case("Height") {
            Some((&self.height, None))
        } else {
            None
        }
    }
}

fn is_size_cell(name: &str) -> bool {
    name.eq_ignore_ascii_case("Width") || name.eq_ignore_ascii_case("Height")
}

fn numeric_reference(references: &impl References, name: &str) -> Option<f64> {
    let formula = references.formula(name)?;
    match evaluate(
        formula.trim_start_matches('='),
        references,
        &ParseLimits::default(),
    ) {
        Evaluation::Evaluated(result) => match result.value {
            Value::Number(number) => Some(number.number),
            Value::Color(_) => None,
        },
        _ => None,
    }
}

fn expression_references_pin(
    expression: &Expr,
    base: &vsdx_resolve::ResolvedShape,
    visited: &mut HashSet<String>,
    depth: u8,
) -> bool {
    match expression {
        Expr::Reference(name) => {
            if name.contains('!') {
                return false;
            }
            let cell = name.rsplit('.').next().unwrap_or(name);
            if cell.eq_ignore_ascii_case("PinX") || cell.eq_ignore_ascii_case("PinY") {
                return true;
            }
            if depth >= 32 || !visited.insert(name.to_ascii_uppercase()) {
                return false;
            }
            let Some(formula) = base.formula(name) else {
                return false;
            };
            let trimmed = formula.trim_start_matches('=').trim();
            if trimmed.is_empty() {
                return false;
            }
            match vsdx_eval::parse(trimmed, &ParseLimits::default()) {
                Ok(nested) => expression_references_pin(&nested, base, visited, depth + 1),
                Err(_) => false,
            }
        }
        Expr::Unary(inner) => expression_references_pin(inner, base, visited, depth),
        Expr::Binary(left, _, right) => {
            expression_references_pin(left, base, visited, depth)
                || expression_references_pin(right, base, visited, depth)
        }
        Expr::Call(_, arguments) => arguments
            .iter()
            .any(|argument| expression_references_pin(argument, base, visited, depth)),
        Expr::Number(_, _) | Expr::String(_) => false,
    }
}

fn loc_pin_component(
    base: &vsdx_resolve::ResolvedShape,
    overrides: &SizeOverrideRefs<'_>,
    name: &str,
    proposed_size: f64,
) -> f64 {
    let current = numeric_reference(base, name)
        .or_else(|| {
            base.value(name)
                .and_then(|(value, _)| value.parse::<f64>().ok())
        })
        .unwrap_or(proposed_size / 2.0);
    let Some(formula) = base.formula(name) else {
        return current;
    };
    let trimmed = formula.trim_start_matches('=').trim();
    if trimmed.is_empty() {
        return current;
    }
    match vsdx_eval::parse(trimmed, &ParseLimits::default()) {
        Ok(expression) if expression_references_pin(&expression, base, &mut HashSet::new(), 0) => {
            current
        }
        Ok(_) => match evaluate(trimmed, overrides, &ParseLimits::default()) {
            Evaluation::Evaluated(result) => match result.value {
                Value::Number(number) => number.number,
                Value::Color(_) => current,
            },
            _ => current,
        },
        Err(_) => current,
    }
}

fn materialized_source_ids<T: ReadTxn>(
    sheets: &MapRef,
    txn: &T,
    roots: &ArrayRef,
    mut largest: u32,
) -> EditResult<std::collections::BTreeMap<String, u32>> {
    let entries = shape_tree_entries(sheets, txn, roots)?;
    let mut taken = HashSet::new();
    let mut added = Vec::new();
    for entry in entries {
        let shape = map_ref(sheets, txn, &entry.id)?;
        let source_id = map_number(&shape, txn, "sourceId")
            .ok_or_else(|| EditError::InvalidState("missing source ID".to_owned()))?
            as u32;
        if shape_origin(&shape, txn)? == ShapeOrigin::Original {
            largest = largest.max(source_id);
            taken.insert(source_id);
        } else {
            added.push((entry.id, source_id));
        }
    }
    added.sort();
    let mut result = std::collections::BTreeMap::new();
    for (id, stored) in added {
        if stored != 0 && taken.insert(stored) {
            result.insert(id, stored);
            largest = largest.max(stored);
        } else {
            largest = largest.checked_add(1).ok_or_else(|| {
                EditError::InvalidState("cannot allocate a materialized source ID".to_owned())
            })?;
            taken.insert(largest);
            result.insert(id, largest);
        }
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
                        row_type: map_string(&cell, &txn, "rowType"),
                    });
                }
            }
        }
    }
    Ok(edits)
}

/// Plain-text overrides for original shapes whose stories diverge from the package.
fn semantic_text_edits(doc: &Doc) -> EditResult<Vec<vsdx_parse::SemanticTextEdit>> {
    let package = original_package_from_doc(doc)?;
    semantic_text_edits_in(doc, &package)
}

fn semantic_text_edits_in(
    doc: &Doc,
    package: &vsdx_parse::VsdxPackage,
) -> EditResult<Vec<vsdx_parse::SemanticTextEdit>> {
    let txn = doc.transact();
    let sheets = required_map(&txn, SHEETS)?;
    let stories = txn.get_map(STORIES);
    let resolver = Resolver::new(package);
    let mut edits = Vec::new();
    for (shape_id, shape) in sheets.iter(&txn) {
        let Out::YMap(shape) = shape else { continue };
        if shape_origin(&shape, &txn)? != ShapeOrigin::Original {
            continue;
        }
        let current = match stories
            .as_ref()
            .and_then(|stories| stories.get(&txn, shape_id))
        {
            None => continue,
            Some(Out::Any(Any::String(value))) => value.to_string(),
            Some(_) => {
                return Err(EditError::InvalidState(
                    "shape text is not a string".to_owned(),
                ));
            }
        };
        validate_story_text(&current)?;
        let page_id = map_string(&shape, &txn, "pageId")
            .ok_or_else(|| EditError::InvalidState("shape is missing its page".to_owned()))?;
        let source_id = map_number(&shape, &txn, "sourceId")
            .ok_or_else(|| EditError::InvalidState("missing source ID".to_owned()))?
            as u32;
        let source_page_id = page_id
            .strip_prefix("page:")
            .and_then(|value| value.parse::<u32>().ok())
            .ok_or_else(|| EditError::InvalidState("invalid page ID".to_owned()))?;
        let path = package
            .page_part_ids
            .iter()
            .find_map(|(path, id)| (*id == source_page_id).then(|| path.clone()))
            .ok_or_else(|| EditError::InvalidState("page is missing".to_owned()))?;
        let contents = package
            .page_contents
            .get(&path)
            .ok_or_else(|| EditError::InvalidState("page content is missing".to_owned()))?;
        let Some(original) = find_shape_in(contents, source_id) else {
            continue;
        };
        let resolved = resolver
            .resolve_shape(&path, source_id)
            .map_err(|error| EditError::InvalidState(error.to_string()))?;
        let tokens = resolver
            .resolve_text_in_context(original, contents, &resolved)
            .map_err(|error| EditError::InvalidState(error.to_string()))?;
        if plain_text(&tokens) == current {
            continue;
        }
        edits.push(vsdx_parse::SemanticTextEdit {
            locator: CellLocator {
                sheet: CellSheet::Page(source_page_id),
                shape_id: Some(source_id),
                section: None,
                section_index: None,
                row: None,
                cell_name: "Text".to_owned(),
            },
            text: current,
        });
    }
    Ok(edits)
}

/// Original package shape by page-local ID, descending into groups.
fn find_shape_in(sheet: &vsdx_parse::Sheet, source_id: u32) -> Option<&Shape> {
    fn visit<'a>(shapes: impl Iterator<Item = &'a Shape>, source_id: u32) -> Option<&'a Shape> {
        for shape in shapes {
            if shape.id == source_id {
                return Some(shape);
            }
            if let Some(found) = visit(shape.shapes(), source_id) {
                return Some(found);
            }
        }
        None
    }
    visit(sheet.shapes(), source_id)
}

fn serialize_doc(doc: &Doc) -> EditResult<Vec<u8>> {
    let package = original_package_from_doc(doc)?;
    let snapshot = snapshot_doc(doc)?;
    let structural = structural_edits(&package, doc, &snapshot)?;
    let cell_edits = semantic_cell_edits(doc)?;
    let text_edits = semantic_text_edits_in(doc, &package)?;
    let mut bytes = if structural.is_empty() {
        vsdx_parse::save_semantic_cell_edits(&package, &cell_edits)
            .map_err(|error| EditError::Parse(error.to_string()))?
    } else {
        let restructured = vsdx_parse::save_structural_edits(&package, &structural)
            .map_err(|error| EditError::Parse(error.to_string()))?;
        let staged = vsdx_parse::parse_vsdx(&restructured)
            .map_err(|error| EditError::Parse(error.to_string()))?;
        vsdx_parse::save_semantic_cell_edits(&staged, &cell_edits)
            .map_err(|error| EditError::Parse(error.to_string()))?
    };
    if !text_edits.is_empty() {
        let staged =
            vsdx_parse::parse_vsdx(&bytes).map_err(|error| EditError::Parse(error.to_string()))?;
        bytes = vsdx_parse::save_semantic_text_edits(&staged, &text_edits)
            .map_err(|error| EditError::Parse(error.to_string()))?;
    }
    Ok(bytes)
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
    let glue = connect::glue_records(&txn)?;
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
        let stories = txn.get_map(STORIES);
        let finals = snapshots
            .iter()
            .map(|(id, snapshot)| ((*id).to_owned(), snapshot.source_id))
            .collect::<std::collections::BTreeMap<_, _>>();
        let remaps = added_subtree_remaps(&sheets, &txn, &roots, &finals)?;
        for shape in desired.iter().filter(|shape| !shape.is_original) {
            let snapshot = snapshots.get(shape.id.as_str()).ok_or_else(|| {
                EditError::InvalidState(format!(
                    "added shape {:?} is missing from snapshot",
                    shape.id
                ))
            })?;
            added.insert(shape.id.as_str(), shape.source_id);
            if shape.parent_id.is_some() {
                continue;
            }
            let mut texts = std::collections::BTreeMap::new();
            let mut verbatim = std::collections::BTreeMap::new();
            let mut pending = vec![*snapshot];
            while let Some(node) = pending.pop() {
                let node_text =
                    stories
                        .as_ref()
                        .map(|stories| match stories.get(&txn, node.id.as_str()) {
                            Some(Out::Any(Any::String(value))) => value.to_string(),
                            _ => String::new(),
                        });
                if let Some(node_text) = node_text.as_ref() {
                    validate_story_text(node_text)?;
                }
                let node_text = node_text.unwrap_or_default();
                texts.insert(node.id.as_str(), node_text.clone());
                if let Some(copy_source_id) = node.copy_source_id
                    && let Some(tokens) = verbatim_source_text(
                        package,
                        node.copy_source_page_id.unwrap_or(*page_id),
                        copy_source_id,
                        &node_text,
                    )
                {
                    verbatim.insert(node.id.as_str(), tokens);
                }
                pending.extend(node.children.iter());
            }
            edits.push(StructuralEdit::AddShape {
                page_id: *page_id,
                shape_xml: shape_xml(snapshot, &texts, &verbatim, remaps.get(shape.id.as_str()))
                    .into_bytes(),
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
        for pending in connect::page_glue(page, &glue) {
            edits.push(StructuralEdit::AddConnect {
                page_id: *page_id,
                from_sheet: pending.from_sheet,
                from_cell: pending.from_cell.to_owned(),
                to_sheet: pending.to_sheet,
                to_cell: pending.to_cell,
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

/// Stored-to-final source ID maps for pasted subtrees, keyed by member session ID.
fn added_subtree_remaps<T: ReadTxn>(
    sheets: &MapRef,
    txn: &T,
    roots: &ArrayRef,
    finals: &std::collections::BTreeMap<String, u32>,
) -> EditResult<std::collections::BTreeMap<String, std::collections::BTreeMap<u32, u32>>> {
    let entries = shape_tree_entries(sheets, txn, roots)?;
    let mut is_added = std::collections::BTreeMap::new();
    for entry in &entries {
        let shape = map_ref(sheets, txn, &entry.id)?;
        is_added.insert(
            entry.id.as_str(),
            shape_origin(&shape, txn)? == ShapeOrigin::Added,
        );
    }
    let mut root_of = std::collections::BTreeMap::new();
    for entry in &entries {
        if !is_added.get(entry.id.as_str()).copied().unwrap_or(false) {
            continue;
        }
        match &entry.parent_id {
            None => {
                root_of.insert(entry.id.as_str(), entry.id.as_str());
            }
            Some(parent) => {
                if !is_added.get(parent.as_str()).copied().unwrap_or(false) {
                    return Err(EditError::InvalidState(
                        "added shapes cannot have an original parent".to_owned(),
                    ));
                }
                let root = root_of.get(parent.as_str()).copied().ok_or_else(|| {
                    EditError::InvalidState("shape is missing from its parent order".to_owned())
                })?;
                root_of.insert(entry.id.as_str(), root);
            }
        }
    }
    let mut maps = std::collections::BTreeMap::new();
    for entry in &entries {
        if !is_added.get(entry.id.as_str()).copied().unwrap_or(false) {
            continue;
        }
        let shape = map_ref(sheets, txn, &entry.id)?;
        let stored = map_number(&shape, txn, "sourceId")
            .ok_or_else(|| EditError::InvalidState("missing source ID".to_owned()))?
            as u32;
        let final_id = finals.get(&entry.id).copied().ok_or_else(|| {
            EditError::InvalidState("added shape is missing its allocated source ID".to_owned())
        })?;
        if stored != 0 {
            maps.entry(
                root_of
                    .get(entry.id.as_str())
                    .copied()
                    .unwrap_or_default()
                    .to_owned(),
            )
            .or_insert_with(std::collections::BTreeMap::new)
            .insert(stored, final_id);
        }
    }
    let mut remaps = std::collections::BTreeMap::new();
    for entry in &entries {
        if !is_added.get(entry.id.as_str()).copied().unwrap_or(false) {
            continue;
        }
        let root = root_of.get(entry.id.as_str()).copied().unwrap_or_default();
        if let Some(map) = maps.get(root) {
            remaps.insert(entry.id.clone(), map.clone());
        }
    }
    Ok(remaps)
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

fn shape_xml(
    shape: &ShapeSnapshot,
    texts: &std::collections::BTreeMap<&str, String>,
    verbatim: &std::collections::BTreeMap<&str, Vec<TextToken>>,
    remap: Option<&std::collections::BTreeMap<u32, u32>>,
) -> String {
    let mut output = format!(
        "<Shape Type=\"{}\" ID=\"{}\"",
        if shape.children.is_empty() {
            "Shape"
        } else {
            "Group"
        },
        shape.source_id
    );
    if let Some(master) = shape.master {
        output.push_str(&format!(" Master=\"{master}\""));
    }
    if let Some(name) = &shape.name {
        output.push_str(" Name=\"");
        xml_escape(&mut output, name);
        output.push('\"');
    }
    output.push('>');
    for cell in &shape.cells {
        if cell.locator.section.is_none() {
            cell_xml(&mut output, cell, remap);
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
                cell_xml(&mut output, cell, remap);
            }
            output.push_str("</Row>");
        }
        output.push_str("</Section>");
    }
    if let Some(tokens) = verbatim
        .get(shape.id.as_str())
        .filter(|tokens| !tokens.is_empty())
    {
        output.push_str("<Text>");
        for token in tokens {
            text_token_xml(&mut output, token);
        }
        output.push_str("</Text>");
    } else if let Some(text) = texts.get(shape.id.as_str()).filter(|text| !text.is_empty()) {
        output.push_str("<Text>");
        xml_escape(&mut output, text);
        output.push_str("</Text>");
    }
    if !shape.children.is_empty() {
        output.push_str("<Shapes>");
        for child in &shape.children {
            output.push_str(&shape_xml(child, texts, verbatim, remap));
        }
        output.push_str("</Shapes>");
    }
    output.push_str("</Shape>");
    output
}

fn text_token_xml(output: &mut String, token: &TextToken) {
    match token {
        TextToken::Literal(value) => {
            for character in value.chars() {
                match character {
                    '&' => output.push_str("&amp;"),
                    '<' => output.push_str("&lt;"),
                    '>' => output.push_str("&gt;"),
                    _ => output.push(character),
                }
            }
        }
        TextToken::CharacterRun(ix) => {
            output.push_str(&format!("<cp IX=\"{ix}\"></cp>"));
        }
        TextToken::ParagraphRun(ix) => {
            output.push_str(&format!("<pp IX=\"{ix}\"></pp>"));
        }
        TextToken::Tab(ix) => {
            output.push_str(&format!("<tp IX=\"{ix}\"></tp>"));
        }
        TextToken::Field(ix) => {
            output.push_str(&format!("<fld IX=\"{ix}\"></fld>"));
        }
    }
}

fn cell_xml(
    output: &mut String,
    cell: &CellSnapshot,
    remap: Option<&std::collections::BTreeMap<u32, u32>>,
) {
    output.push_str("<Cell N=\"");
    xml_escape(output, &cell.name);
    output.push('\"');
    let formula = cell.formula.as_deref().map(|formula| match remap {
        Some(remap) => remap_sheet_refs(formula, remap).into_owned(),
        None => formula.to_owned(),
    });
    if let Some(formula) = formula.as_deref() {
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

    fn current_text(&self, locator: &CellLocator) -> Option<String> {
        let key = locator_key(locator);
        self.values
            .get(&key)
            .or_else(|| self.formulas.get(&key))
            .cloned()
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
            section: None,
            section_index: None,
            row: None,
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

/// Splits a policy refusal into the two the UI words differently. The policy returns prose,
/// so this lives beside it rather than in a client that would have to match the same strings.
fn refusal_kind(reason: &str) -> &'static str {
    if reason.starts_with("GUARD") {
        "guard"
    } else {
        "lock"
    }
}

/// Stand-in value for a probe; the policy's refusal never reads the incoming formula.
const PROBE_FORMULA: &str = "0";

fn gesture_for_cell(cell_name: &str) -> MutationGesture {
    match cell_name {
        "PinX" => MutationGesture::MoveX,
        "PinY" => MutationGesture::MoveY,
        "Width" => MutationGesture::ResizeWidth,
        "Height" => MutationGesture::ResizeHeight,
        "Angle" => MutationGesture::Rotate,
        _ => MutationGesture::CellEdit,
    }
}

/// Current plain text behind a shape; absent stories read as empty.
fn story_text<T: ReadTxn>(txn: &T, shape_id: &str) -> String {
    txn.get_map(STORIES)
        .and_then(|stories| match stories.get(txn, shape_id) {
            Some(Out::Any(Any::String(value))) => Some(value.to_string()),
            _ => None,
        })
        .unwrap_or_default()
}

/// Rejects text the span-patching save could not write back.
fn validate_story_text(text: &str) -> EditResult<()> {
    let limits = ParseLimits::default();
    if text.len() > limits.max_xml_text_bytes {
        return Err(EditError::InvalidState(
            "shape text exceeds maximum length".to_owned(),
        ));
    }
    if text.chars().any(|c| !matches!(c, '\u{9}' | '\u{A}' | '\u{D}' | '\u{20}'..='\u{D7FF}' | '\u{E000}'..='\u{FFFD}' | '\u{10000}'..='\u{10FFFF}')) {
        return Err(EditError::InvalidState(
            "shape text contains a character forbidden by XML 1.0".to_owned(),
        ));
    }
    Ok(())
}

const SHAPE_DATA_CELL: &str = "Value";

fn shape_data_locator(row: &CellRow, section_index: Option<u32>) -> CellLocator {
    CellLocator {
        sheet: CellSheet::Page(0),
        shape_id: None,
        section: Some("Property".to_owned()),
        section_index,
        row: Some(row.clone()),
        cell_name: SHAPE_DATA_CELL.to_owned(),
    }
}

fn shape_data_type_refusal(type_code: Option<&str>) -> Option<&'static str> {
    match type_code
        .map(str::trim)
        .map(|code| code.trim_start_matches('=').trim())
    {
        Some("5") => Some("date"),
        Some("6") => Some("duration"),
        Some("7") => Some("currency"),
        _ => None,
    }
}

fn is_numeric_literal(formula: &str) -> bool {
    let trimmed = formula.trim();
    let formula = trimmed.strip_prefix('=').unwrap_or(trimmed);
    let Ok(mut expression) = vsdx_eval::parse(formula, &ParseLimits::default()) else {
        return false;
    };
    while let Expr::Unary(inner) = expression {
        expression = *inner;
    }
    matches!(expression, Expr::Number(number, vsdx_eval::Unit::Number) if number.is_finite())
}

fn shape_data_refusal(
    policy: &CrdtMutationContext,
    write: &ShapeDataWrite,
    before: Option<&str>,
) -> Option<String> {
    if write.formula.trim().is_empty() {
        return Some("a shape-data row cannot be set to an empty formula".to_owned());
    }
    let type_locator = policy.locator(CellLocator {
        cell_name: "Type".to_owned(),
        ..shape_data_locator(&write.row, write.section_index)
    });
    let type_code = policy.current_text(&type_locator);
    if let Some(kind) = shape_data_type_refusal(type_code.as_deref()) {
        return Some(format!(
            "a {kind} shape-data row keeps its typed value; editing it here would store text"
        ));
    }
    let numeric = matches!(
        type_code
            .as_deref()
            .map(str::trim)
            .map(|code| code.trim_start_matches('=').trim()),
        Some("2")
    );
    if numeric {
        if !is_numeric_literal(&write.formula) {
            return Some("a number shape-data row takes a numeric value".to_owned());
        }
        if before.is_some_and(|formula| !is_numeric_literal(formula)) {
            return Some(
                "a number shape-data row holding a formula is not editable as a value".to_owned(),
            );
        }
    }
    None
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

#[cfg(test)]
mod tests {
    use super::*;
    use crate::EditCtx;

    const GROUP_MASTER_TEXT: &[u8] =
        include_bytes!("../../vsdx-parse/tests/fixtures/group-master-text.vsdx");
    const SUB_SHAPE: &str = "page:1:shape:1:shape:2";
    const MASTER_PART: &str = "visio/masters/master7.xml";
    const PAGE_PART: &str = "visio/pages/page1.xml";

    fn part(bytes: &[u8], path: &str) -> Vec<u8> {
        vsdx_parse::parse_vsdx(bytes)
            .unwrap()
            .part_bytes(path)
            .unwrap()
            .to_vec()
    }

    fn story_values(doc: &Doc) -> std::collections::BTreeMap<String, String> {
        let txn = doc.transact();
        txn.get_map(STORIES)
            .unwrap()
            .iter(&txn)
            .map(|(id, value)| {
                let Out::Any(Any::String(text)) = value else {
                    panic!("story {id} is not a string");
                };
                (id.to_owned(), text.to_string())
            })
            .collect()
    }

    fn stamp_token_stories(
        session: &DiagramSession,
        stories: &std::collections::BTreeMap<String, String>,
    ) {
        let mut txn = session.doc.transact_mut_with(crate::HYDRATE_ORIGIN);
        let map = txn.get_or_insert_map(STORIES);
        for (id, text) in stories {
            let tokens = if text.is_empty() {
                Vec::new()
            } else {
                vec![vsdx_resolve::ResolvedTextToken::Literal(text.clone())]
            };
            let json = serde_json::to_string(&tokens).unwrap();
            map.insert(&mut txn, id.as_str(), json.as_str());
        }
        txn.get_or_insert_map(META)
            .insert(&mut txn, "schemaVersion", TOKEN_STORY_SCHEMA_VERSION);
    }

    #[test]
    fn legacy_token_stories_migrate_to_plain_text_on_open() {
        let session = DiagramSession::open(GROUP_MASTER_TEXT, 3).unwrap();
        let expected = story_values(&session.doc);
        assert!(expected.values().any(|text| text == "group label"));
        assert!(expected.values().any(String::is_empty));
        stamp_token_stories(&session, &expected);
        assert!(
            story_values(&session.doc)
                .values()
                .all(|text| text.starts_with('['))
        );

        let reopened =
            DiagramSession::open_from_update(&session.encode_state_as_update_v1(), 4).unwrap();
        assert_eq!(story_values(&reopened.doc), expected);
        assert_eq!(
            reopened.shape_text("page:1", SUB_SHAPE).unwrap(),
            "group label"
        );
        let txn = reopened.doc.transact();
        assert_eq!(
            map_number(&required_map(&txn, META).unwrap(), &txn, "schemaVersion"),
            Some(SCHEMA_VERSION)
        );
        drop(txn);
        assert!(semantic_text_edits(&reopened.doc).unwrap().is_empty());
        assert!(!reopened.can_undo());
        assert_eq!(
            reopened.save().unwrap(),
            vsdx_parse::write_vsdx(&vsdx_parse::parse_vsdx(GROUP_MASTER_TEXT).unwrap()).unwrap()
        );
    }

    #[test]
    fn remote_update_cannot_downgrade_the_schema_version() {
        let live = DiagramSession::open(GROUP_MASTER_TEXT, 5).unwrap();
        let peer = DiagramSession::open(GROUP_MASTER_TEXT, 6).unwrap();
        let expected = story_values(&live.doc);
        stamp_token_stories(&peer, &expected);

        let error = live
            .apply_update_v1(&peer.encode_state_as_update_v1())
            .unwrap_err();
        assert!(
            matches!(&error, EditError::InvalidState(reason) if reason.contains("schema version")),
            "unexpected error {error:?}"
        );
        assert_eq!(story_values(&live.doc), expected);
    }

    #[test]
    fn group_sub_shape_text_inherits_through_its_enclosing_master() {
        let session = DiagramSession::open(GROUP_MASTER_TEXT, 3).unwrap();
        assert_eq!(
            session.shape_text("page:1", SUB_SHAPE).unwrap(),
            "group label"
        );
        assert!(semantic_text_edits(&session.doc).unwrap().is_empty());
        assert_eq!(
            session.save().unwrap(),
            vsdx_parse::write_vsdx(&vsdx_parse::parse_vsdx(GROUP_MASTER_TEXT).unwrap()).unwrap()
        );
    }

    #[test]
    fn unedited_shapes_keep_their_original_text_markers() {
        let source = include_bytes!("../../vsdx-parse/tests/fixtures/text-accounting.vsdx");
        let session = DiagramSession::open(source, 3).unwrap();
        session
            .set_shape_text(&EditCtx::local("t"), "page:1", "page:1:shape:1", "edited")
            .unwrap();
        let package = session.package().unwrap();
        let sheet = &package.page_contents["visio/pages/page1.xml"];
        let text = |id: u32| {
            sheet
                .shapes()
                .find(|shape| shape.id == id)
                .unwrap()
                .text()
                .map(<[_]>::to_vec)
        };
        assert_eq!(
            text(1),
            Some(vec![vsdx_parse::TextToken::Literal("edited".to_owned())])
        );
        assert_eq!(text(2), Some(vec![vsdx_parse::TextToken::Field(0)]));
        assert_eq!(
            text(3),
            Some(vec![
                vsdx_parse::TextToken::CharacterRun(0),
                vsdx_parse::TextToken::Literal("style text".to_owned()),
            ])
        );
        assert_eq!(text(4), None);
    }

    #[test]
    fn text_given_to_an_added_shape_survives_saving() {
        let source = include_bytes!("../../vsdx-parse/tests/fixtures/text-accounting.vsdx");
        let session = DiagramSession::open(source, 3).unwrap();
        let context = EditCtx::local("t");
        let draft = ShapeDraft {
            name: None,
            master: None,
            cells: ["Width", "Height", "PinX", "PinY"]
                .into_iter()
                .map(|name| CellSnapshot {
                    row_type: None,
                    locator: CellLocator {
                        sheet: CellSheet::Page(1),
                        shape_id: None,
                        section: None,
                        section_index: None,
                        row: None,
                        cell_name: name.to_owned(),
                    },
                    name: name.to_owned(),
                    formula: Some("1".to_owned()),
                    value: None,
                })
                .collect(),
        };
        let receipt = session.add_shape(&context, "page:1", &draft).unwrap();
        session
            .set_shape_text(&context, "page:1", &receipt.shape_id, "added & <new>")
            .unwrap();
        let reopened = DiagramSession::open(&session.save().unwrap(), 4).unwrap();
        let added = reopened.snapshot().unwrap().pages[0]
            .shapes
            .last()
            .unwrap()
            .id
            .clone();
        assert_eq!(
            reopened.shape_text("page:1", &added).unwrap(),
            "added & <new>"
        );
    }

    #[test]
    fn text_edits_are_undoable() {
        let source = include_bytes!("../../vsdx-parse/tests/fixtures/text-accounting.vsdx");
        let session = DiagramSession::open(source, 3).unwrap();
        let before = session.shape_text("page:1", "page:1:shape:3").unwrap();
        session
            .set_shape_text(&EditCtx::local("t"), "page:1", "page:1:shape:3", "edited")
            .unwrap();
        session.add_undo_barrier();
        assert!(session.undo());
        assert_eq!(
            session.shape_text("page:1", "page:1:shape:3").unwrap(),
            before
        );
        assert!(session.redo());
        assert_eq!(
            session.shape_text("page:1", "page:1:shape:3").unwrap(),
            "edited"
        );
    }

    #[test]
    fn clearing_inherited_text_shadows_the_master_after_reopening() {
        let session = DiagramSession::open(GROUP_MASTER_TEXT, 3).unwrap();
        session
            .set_shape_text(&EditCtx::local("t"), "page:1", SUB_SHAPE, "")
            .unwrap();
        let saved = session.save().unwrap();
        assert_eq!(
            DiagramSession::open(&saved, 4)
                .unwrap()
                .shape_text("page:1", SUB_SHAPE)
                .unwrap(),
            ""
        );
        let package = session.package().unwrap();
        let projected = package.page_contents["visio/pages/page1.xml"]
            .shapes()
            .flat_map(vsdx_parse::Shape::shapes)
            .find(|shape| shape.id == 2)
            .unwrap()
            .text()
            .map(<[_]>::to_vec);
        assert_eq!(projected, Some(Vec::new()));
    }

    #[test]
    fn editing_an_inherited_sub_shape_rewrites_only_its_own_text() {
        let session = DiagramSession::open(GROUP_MASTER_TEXT, 3).unwrap();
        session
            .set_shape_text(&EditCtx::local("t"), "page:1", SUB_SHAPE, "renamed")
            .unwrap();
        let saved = session.save().unwrap();
        assert_eq!(
            DiagramSession::open(&saved, 4)
                .unwrap()
                .shape_text("page:1", SUB_SHAPE)
                .unwrap(),
            "renamed"
        );
        assert_eq!(
            part(&saved, MASTER_PART),
            part(GROUP_MASTER_TEXT, MASTER_PART)
        );
        let page = String::from_utf8(part(&saved, PAGE_PART)).unwrap();
        assert!(page.contains("<Text>renamed</Text>"), "{page}");
        assert_eq!(
            page.replace("<Text>renamed</Text>", ""),
            String::from_utf8(part(GROUP_MASTER_TEXT, PAGE_PART)).unwrap()
        );
    }
}
