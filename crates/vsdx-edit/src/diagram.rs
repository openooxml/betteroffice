use std::sync::Arc;

use vsdx_eval::{MutationContext, MutationOutcome, decide_mutation, evaluate};
use vsdx_parse::{
    Cell, CellLocator, CellRow, CellSheet, MutationGesture, ParseLimits, RowChild, SectionChild,
    Shape, ShapeChild, ShapesChild, SheetChild,
};
use vsdx_resolve::{Lookup, Resolver};
use yrs::{
    Any, Array, ArrayPrelim, ArrayRef, Doc, Map, MapPrelim, MapRef, Out, ReadTxn, Transact,
    TransactionMut, WriteTxn,
};

use crate::{
    CellFormulaReceipt, CellSnapshot, DiagramSession, DiagramSnapshot, EditCtx, EditError,
    EditResult, META, PAGE_ORDER, PAGES, PageSnapshot, SHEETS, STORIES, ShapeDraft, ShapeReceipt,
    ShapeSnapshot,
};

const SCHEMA_VERSION: f64 = 1.0;

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
    let stories = txn.get_or_insert_map(STORIES);
    for path in &package.page_part_paths {
        let Some(page_id) = package.page_part_ids.get(path) else {
            continue;
        };
        let id = format!("page:{page_id}");
        order.push_back(&mut txn, id.as_str());
        let page = pages.insert(&mut txn, id.as_str(), MapPrelim::default());
        page.insert(&mut txn, "id", id.as_str());
        page.insert(&mut txn, "sourcePartPath", path.as_str());
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
                    &resolver, &resolved,
                )?;
            }
        }
    }
    Ok(())
}

pub(crate) fn package_from_doc(doc: &Doc) -> EditResult<vsdx_parse::VsdxPackage> {
    let txn = doc.transact();
    let meta = required_map(&txn, META)?;
    if map_number(&meta, &txn, "schemaVersion") != Some(SCHEMA_VERSION) {
        return Err(EditError::InvalidState(
            "unsupported diagram schema version".to_owned(),
        ));
    }
    let mut package = match meta.get(&txn, "packageBytes") {
        Some(Out::Any(Any::Buffer(bytes))) => vsdx_parse::parse_vsdx(&bytes)
            .map_err(|error| EditError::InvalidState(error.to_string()))?,
        _ => {
            let Some(Out::Any(Any::Buffer(bytes))) = meta.get(&txn, "packageJson") else {
                return Err(EditError::InvalidState("missing package data".to_owned()));
            };
            serde_json::from_slice(&bytes)
                .map_err(|error| EditError::InvalidState(error.to_string()))?
        }
    };
    materialize_snapshot(&mut package, &snapshot_doc(doc)?);
    Ok(package)
}

fn materialize_snapshot(package: &mut vsdx_parse::VsdxPackage, snapshot: &DiagramSnapshot) {
    for page in &snapshot.pages {
        let Some(sheet) = package.page_contents.get_mut(&page.source_part_path) else {
            continue;
        };
        materialize_page_shapes(sheet, page);
    }
}

fn materialize_page_shapes(sheet: &mut vsdx_parse::Sheet, page: &PageSnapshot) {
    let originals = sheet
        .shapes()
        .cloned()
        .map(|shape| (shape.id, shape))
        .collect::<std::collections::BTreeMap<_, _>>();
    let mut next_id = largest_shape_id(sheet);
    let mut shapes = Vec::with_capacity(page.shapes.len());
    for snapshot in &page.shapes {
        let canonical_id = format!("{}:shape:{}", page.id, snapshot.source_id);
        let mut shape = if snapshot.id == canonical_id {
            originals
                .get(&snapshot.source_id)
                .cloned()
                .unwrap_or_else(|| shape_from_snapshot(snapshot, &mut next_id))
        } else {
            shape_from_snapshot(snapshot, &mut next_id)
        };
        materialize_shape(&mut shape, snapshot);
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
}

fn largest_shape_id(sheet: &vsdx_parse::Sheet) -> u32 {
    sheet
        .shapes()
        .map(largest_shape_id_including_children)
        .max()
        .unwrap_or_default()
}

fn largest_shape_id_including_children(shape: &Shape) -> u32 {
    shape
        .shapes()
        .map(largest_shape_id_including_children)
        .max()
        .unwrap_or_default()
        .max(shape.id)
}

fn shape_from_snapshot(snapshot: &ShapeSnapshot, next_id: &mut u32) -> Shape {
    *next_id = next_id.saturating_add(1);
    Shape {
        id: *next_id,
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

fn materialize_shape(shape: &mut Shape, snapshot: &ShapeSnapshot) {
    if shape.id == snapshot.source_id {
        for cell in &snapshot.cells {
            materialize_cell(shape, cell);
        }
    }
    if snapshot.children.is_empty() {
        return;
    }
    let originals = shape
        .shapes()
        .cloned()
        .map(|child| (child.id, child))
        .collect::<std::collections::BTreeMap<_, _>>();
    let mut next_id = shape
        .shapes()
        .map(largest_shape_id_including_children)
        .max()
        .unwrap_or(shape.id);
    let mut children = Vec::with_capacity(snapshot.children.len());
    for child_snapshot in &snapshot.children {
        let mut child = originals
            .get(&child_snapshot.source_id)
            .cloned()
            .unwrap_or_else(|| shape_from_snapshot(child_snapshot, &mut next_id));
        materialize_shape(&mut child, child_snapshot);
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
            if section.name != *section_name {
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
            ShapeChild::Section(section) if section.name == *section_name => Some(section),
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
                    row_type: None,
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
                    index: None,
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
                        row_type: None,
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
    page: &vsdx_parse::Sheet,
    shape: &vsdx_parse::Shape,
    resolver: &Resolver<'_>,
    resolved: &vsdx_resolve::ResolvedShape,
) -> EditResult<()> {
    let map = sheets.insert(txn, id, MapPrelim::default());
    map.insert(txn, "id", id);
    map.insert(txn, "pageId", page_id);
    map.insert(txn, "sourceId", shape.id as f64);
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
                    row: None,
                    cell_name: name.clone(),
                },
                cell.cell.formula.as_deref(),
                cell.cell.value.as_deref(),
            );
        }
    }
    for (section_name, section) in &resolved.sections {
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
                            section: Some(section_name.clone()),
                            row: Some(row.clone()),
                            cell_name: name.clone(),
                        },
                        cell.cell.formula.as_deref(),
                        cell.cell.value.as_deref(),
                    );
                }
            }
        }
    }
    let text = resolver
        .resolve_text(shape, page)
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
            page,
            child,
            resolver,
            &child_resolved,
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
) {
    let key = locator_key(locator);
    let cell = cells.insert(txn, key.as_str(), MapPrelim::default());
    cell.insert(txn, "name", locator.cell_name.as_str());
    if let Some(section) = &locator.section {
        cell.insert(txn, "section", section.as_str());
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
        let mut txn = self.transact_for(context);
        let pages = txn
            .get_map(PAGES)
            .ok_or_else(|| EditError::InvalidState("missing pages map".to_owned()))?;
        let page = map_ref(&pages, &txn, page_id)?;
        let order = map_array(&page, &txn, "shapes")?;
        let index = order.len(&txn);
        let id = self.next_id(&format!("{page_id}:shape"));
        let sheets = txn
            .get_map(SHEETS)
            .ok_or_else(|| EditError::InvalidState("missing sheets map".to_owned()))?;
        let shape = sheets.insert(&mut txn, id.as_str(), MapPrelim::default());
        shape.insert(&mut txn, "id", id.as_str());
        shape.insert(&mut txn, "pageId", page_id);
        shape.insert(&mut txn, "sourceId", draft.source_id as f64);
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

pub(crate) fn validate_doc(doc: &Doc) -> EditResult<()> {
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
    for index in 0..order.len(&txn) {
        let page_id = array_string(&order, &txn, index)
            .ok_or_else(|| EditError::InvalidState("page order contains non-string".to_owned()))?;
        let page = map_ref(&pages, &txn, &page_id)?;
        required_string(&page, &txn, "id")?;
        required_string(&page, &txn, "sourcePartPath")?;
        let shapes = map_array(&page, &txn, "shapes")?;
        for shape_index in 0..shapes.len(&txn) {
            let shape_id = array_string(&shapes, &txn, shape_index).ok_or_else(|| {
                EditError::InvalidState("shape order contains non-string".to_owned())
            })?;
            let shape = map_ref(&sheets, &txn, &shape_id)?;
            required_string(&shape, &txn, "id")?;
            if map_number(&shape, &txn, "sourceId").is_none() {
                return Err(EditError::InvalidState("missing source ID".to_owned()));
            }
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
    Ok(())
}

pub(crate) fn validate_remote_update(before: &Doc, staged: &Doc) -> EditResult<()> {
    validate_doc(staged)?;
    validate_immutable_metadata(before, staged)?;
    validate_session_topology(before, staged)?;
    validate_formula_mutations(before, staged)?;
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
        for key in ["id", "sourcePartPath"] {
            if before_page.get(&before_txn, key) != staged_page.get(&staged_txn, key) {
                return Err(EditError::InvalidState(format!(
                    "remote update changes immutable page {key}"
                )));
            }
        }
    }
    let before_sheets = required_map(&before_txn, SHEETS)?;
    let staged_sheets = required_map(&staged_txn, SHEETS)?;
    for (shape_id, before_shape) in before_sheets.iter(&before_txn) {
        let Out::YMap(before_shape) = before_shape else {
            continue;
        };
        let staged_shape = map_ref(&staged_sheets, &staged_txn, shape_id)?;
        let before_cells = map_map(&before_shape, &before_txn, "cells")?;
        let staged_cells = map_map(&staged_shape, &staged_txn, "cells")?;
        for (cell_id, before_cell) in before_cells.iter(&before_txn) {
            let Out::YMap(before_cell) = before_cell else {
                continue;
            };
            let staged_cell = map_ref(&staged_cells, &staged_txn, cell_id)?;
            if before_cell.get(&before_txn, "value") != staged_cell.get(&staged_txn, "value") {
                return Err(EditError::InvalidState(
                    "remote update changes untrusted cached cell value".to_owned(),
                ));
            }
        }
    }
    Ok(())
}

pub(crate) fn next_id_counter(doc: &Doc, client_id: u64) -> u64 {
    let txn = doc.transact();
    let Some(sheets) = txn.get_map(SHEETS) else {
        return 0;
    };
    sheets
        .iter(&txn)
        .filter_map(|(_, value)| match value {
            Out::YMap(shape) => map_string(&shape, &txn, "id"),
            _ => None,
        })
        .filter_map(|id| {
            id.rsplit_once(':').and_then(|(prefix, counter)| {
                prefix
                    .ends_with(&format!(":{client_id}"))
                    .then(|| counter.parse::<u64>().ok())
                    .flatten()
            })
        })
        .max()
        .and_then(|counter| counter.checked_add(1))
        .unwrap_or(0)
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
    for (page_id, page) in pages.iter(&before_txn) {
        let Out::YMap(page) = page else { continue };
        for shape_id in reachable_shape_ids(&sheets, &before_txn, &page)? {
            let shape = map_ref(&sheets, &before_txn, &shape_id)?;
            let cells = map_map(&shape, &before_txn, "cells")?;
            let context = CrdtMutationContext::new(&before_txn, page_id, &shape_id)?;
            let staged_shape = required_map(&staged_txn, SHEETS)
                .and_then(|sheets| map_ref(&sheets, &staged_txn, &shape_id))?;
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
    Ok(())
}

fn reachable_shape_ids<T: ReadTxn>(
    sheets: &MapRef,
    txn: &T,
    page: &MapRef,
) -> EditResult<Vec<String>> {
    let roots = map_array(page, txn, "shapes")?;
    let mut pending = Vec::new();
    for index in 0..roots.len(txn) {
        pending.push(array_string(&roots, txn, index).ok_or_else(|| {
            EditError::InvalidState("shape order contains non-string".to_owned())
        })?);
    }
    let mut result = Vec::new();
    while let Some(shape_id) = pending.pop() {
        if result.contains(&shape_id) {
            continue;
        }
        map_ref(sheets, txn, &shape_id)?;
        result.push(shape_id.clone());
        for (child_id, child) in sheets.iter(txn) {
            if let Out::YMap(child) = child
                && map_string(&child, txn, "parentId").as_deref() == Some(&shape_id)
            {
                pending.push(child_id.to_owned());
            }
        }
    }
    Ok(result)
}

fn snapshot_doc(doc: &Doc) -> EditResult<DiagramSnapshot> {
    let txn = doc.transact();
    let order = required_array(&txn, PAGE_ORDER)?;
    let pages = required_map(&txn, PAGES)?;
    let sheets = required_map(&txn, SHEETS)?;
    let mut result = Vec::new();
    for index in 0..order.len(&txn) {
        let id = array_string(&order, &txn, index)
            .ok_or_else(|| EditError::InvalidState("page order contains non-string".to_owned()))?;
        let page = map_ref(&pages, &txn, &id)?;
        let shape_order = map_array(&page, &txn, "shapes")?;
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
) -> EditResult<ShapeSnapshot> {
    let shape = map_ref(sheets, txn, shape_id)?;
    let source_id = map_number(&shape, txn, "sourceId")
        .ok_or_else(|| EditError::InvalidState("missing source ID".to_owned()))?
        as u32;
    let cells = map_map(&shape, txn, "cells")?;
    let mut snapshots = Vec::new();
    for (_key, value) in cells.iter(txn) {
        let Out::YMap(cell) = value else {
            return Err(EditError::InvalidState("cell is not a map".to_owned()));
        };
        let locator = cell_locator(&cell, txn, page_id, source_id)?;
        snapshots.push(CellSnapshot {
            name: locator.cell_name.clone(),
            locator,
            formula: map_string(&cell, txn, "formula"),
            value: map_string(&cell, txn, "value"),
        });
    }
    snapshots.sort_by_key(|cell| locator_key(&cell.locator));
    let children = if let Some(Out::YArray(child_order)) = shape.get(txn, "shapes") {
        let mut children = Vec::with_capacity(child_order.len(txn) as usize);
        for index in 0..child_order.len(txn) {
            let child_id = array_string(&child_order, txn, index).ok_or_else(|| {
                EditError::InvalidState("shape order contains non-string".to_owned())
            })?;
            children.push(snapshot_shape(sheets, txn, &child_id, page_id)?);
        }
        children
    } else {
        let mut ids = sheets
            .iter(txn)
            .filter_map(|(id, value)| match value {
                Out::YMap(child)
                    if map_string(&child, txn, "parentId").as_deref() == Some(shape_id) =>
                {
                    Some(id.to_owned())
                }
                _ => None,
            })
            .collect::<Vec<_>>();
        ids.sort();
        ids.into_iter()
            .map(|id| snapshot_shape(sheets, txn, &id, page_id))
            .collect::<EditResult<Vec<_>>>()?
    };
    Ok(ShapeSnapshot {
        id: shape_id.to_owned(),
        source_id,
        name: map_string(&shape, txn, "name"),
        cells: snapshots,
        children,
    })
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
                let cells = map_map(&shape, &txn, "cells")?;
                for (_key, value) in cells.iter(&txn) {
                    let Out::YMap(cell) = value else { continue };
                    let formula = map_string(&cell, &txn, "formula");
                    if formula == map_string(&cell, &txn, "baselineFormula") {
                        continue;
                    }
                    let Some(formula) = formula else { continue };
                    let locator = cell_locator(&cell, &txn, source_page_id, source_id)?;
                    edits.push(vsdx_parse::SemanticCellEdit {
                        locator: locator.clone(),
                        gesture: gesture_for_cell(&locator.cell_name),
                        formula: Some(formula),
                        value: None,
                    });
                }
            }
        }
    }
    Ok(edits)
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
    let order = map_array(&page, &txn, "shapes")?;
    let length = order.len(&txn);
    if to >= length {
        return Err(EditError::OutOfBounds { index: to, length });
    }
    let from = (0..length)
        .find(|index| array_string(&order, &txn, *index).as_deref() == Some(shape_id))
        .ok_or_else(|| EditError::ShapeNotFound(shape_id.to_owned()))?;
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
}

impl CrdtMutationContext {
    fn new<T: ReadTxn>(txn: &T, page_id: &str, shape_id: &str) -> EditResult<Self> {
        let pages = required_map(txn, PAGES)?;
        map_ref(&pages, txn, page_id)?;
        let sheets = required_map(txn, SHEETS)?;
        let shape = map_ref(&sheets, txn, shape_id)?;
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
        match evaluate(
            formula.trim_start_matches('='),
            &self.formulas,
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
    match (&locator.section, &locator.row) {
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
    let row = match (
        map_number(cell, txn, "rowIndex"),
        map_string(cell, txn, "rowName"),
    ) {
        (Some(index), _) => Some(CellRow::Index(index as u32)),
        (None, Some(name)) => Some(CellRow::Name(name)),
        (None, None) => None,
    };
    Ok(CellLocator {
        sheet: CellSheet::Page(page_id),
        shape_id: Some(shape_id),
        section: map_string(cell, txn, "section"),
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
