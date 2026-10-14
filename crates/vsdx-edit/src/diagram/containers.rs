use std::collections::{BTreeMap, BTreeSet};

use vsdx_parse::{CellLocator, CellSheet, MutationGesture};
use vsdx_resolve::dependson_refs;
use yrs::{Map, Out, ReadTxn, TransactionMut};

use super::{
    CrdtMutationContext, EditCtx, SHEETS, cell_map, evaluate_cached_formula, loc_pin_at_size,
    local_references, map_map, map_number, map_ref, map_string, required_map,
};
use crate::{CellFormulaReceipt, DiagramSession, EditError, EditResult};

const STRUCTURE_KEY: &str = "User\u{1f}N:msvStructureType\u{1f}Value";
const MARGIN_KEY: &str = "User\u{1f}N:msvSDContainerMargin\u{1f}Value";
const LOCKED_KEY: &str = "User\u{1f}N:msvSDContainerLocked\u{1f}Value";

struct Membership {
    by_source: BTreeMap<u32, String>,
    containers: BTreeMap<u32, BTreeSet<u32>>,
    structures: BTreeMap<u32, String>,
}

#[derive(Clone, Copy, Debug)]
struct Bounds {
    x0: f64,
    y0: f64,
    x1: f64,
    y1: f64,
}

impl Bounds {
    fn contains(self, other: Bounds) -> bool {
        self.x0 <= other.x0 && self.y0 <= other.y0 && self.x1 >= other.x1 && self.y1 >= other.y1
    }

    fn union(self, other: Bounds) -> Bounds {
        Bounds {
            x0: self.x0.min(other.x0),
            y0: self.y0.min(other.y0),
            x1: self.x1.max(other.x1),
            y1: self.y1.max(other.y1),
        }
    }

    fn expanded(self, margin: f64) -> Bounds {
        Bounds {
            x0: self.x0 - margin,
            y0: self.y0 - margin,
            x1: self.x1 + margin,
            y1: self.y1 + margin,
        }
    }
}

pub(crate) fn move_container(
    session: &DiagramSession,
    context: &EditCtx,
    page_id: &str,
    shape_id: &str,
    dx: f64,
    dy: f64,
) -> EditResult<Vec<CellFormulaReceipt>> {
    if !dx.is_finite() || !dy.is_finite() {
        return Err(EditError::InvalidState(
            "invalid container delta".to_owned(),
        ));
    }
    let mut txn = session.transact_for(context);
    let membership = membership(&txn, page_id)?;
    let container = source_of(&membership, page_id, shape_id)?;
    require_container(&membership, container)?;
    let mut targets = vec![container];
    targets.extend(transitive_members(&membership, container));
    let mut planned = Vec::new();
    for target in targets {
        let shape = membership.by_source[&target].clone();
        let (x, y) = pin(&txn, &shape)?;
        planned.push((
            shape.clone(),
            "PinX",
            MutationGesture::MoveX,
            format_number(x + dx),
        ));
        planned.push((shape, "PinY", MutationGesture::MoveY, format_number(y + dy)));
    }
    apply(&mut txn, page_id, &planned)
}

pub(crate) fn move_container_member(
    session: &DiagramSession,
    context: &EditCtx,
    page_id: &str,
    shape_id: &str,
    x_formula: String,
    y_formula: String,
) -> EditResult<Vec<CellFormulaReceipt>> {
    let mut txn = session.transact_for(context);
    let membership = membership(&txn, page_id)?;
    let member = source_of(&membership, page_id, shape_id)?;
    let mut planned = vec![
        (
            membership.by_source[&member].clone(),
            "PinX",
            MutationGesture::MoveX,
            x_formula,
        ),
        (
            membership.by_source[&member].clone(),
            "PinY",
            MutationGesture::MoveY,
            y_formula,
        ),
    ];
    let bounds = moved_bounds(&txn, &membership, member, &planned)?;
    let mut grown: BTreeMap<u32, Bounds> = BTreeMap::new();
    grown.insert(member, bounds);
    for container in ancestors(&membership, member) {
        if locked(&txn, &membership, container) {
            continue;
        }
        let margin = margin(&txn, &membership, container);
        let current = shape_bounds(&txn, &membership, container, &grown)?;
        let mut required: Option<Bounds> = None;
        if let Some(direct) = membership.containers.get(&container) {
            for peer in direct {
                let peer_bounds = shape_bounds(&txn, &membership, *peer, &grown)?;
                required = Some(required.map_or(peer_bounds, |union| union.union(peer_bounds)));
            }
        }
        let Some(required) = required else {
            continue;
        };
        if current.contains(required.expanded(margin)) {
            continue;
        }
        let outer = current.union(required.expanded(margin));
        planned.extend(resize_plan(&txn, &membership, container, outer)?);
        grown.insert(container, outer);
    }
    apply(&mut txn, page_id, &planned)
}

pub(crate) fn autofit_container(
    session: &DiagramSession,
    context: &EditCtx,
    page_id: &str,
    shape_id: &str,
) -> EditResult<Vec<CellFormulaReceipt>> {
    let mut txn = session.transact_for(context);
    let membership = membership(&txn, page_id)?;
    let container = source_of(&membership, page_id, shape_id)?;
    require_container(&membership, container)?;
    if locked(&txn, &membership, container) {
        return Ok(Vec::new());
    }
    let grown = BTreeMap::new();
    let margin = margin(&txn, &membership, container);
    let mut required: Option<Bounds> = None;
    if let Some(direct) = membership.containers.get(&container) {
        for peer in direct {
            let peer_bounds = shape_bounds(&txn, &membership, *peer, &grown)?;
            required = Some(required.map_or(peer_bounds, |union| union.union(peer_bounds)));
        }
    }
    let Some(required) = required else {
        return Ok(Vec::new());
    };
    let outer = required.expanded(margin);
    let planned = resize_plan(&txn, &membership, container, outer)?;
    apply(&mut txn, page_id, &planned)
}

fn membership<T: ReadTxn>(txn: &T, page_id: &str) -> EditResult<Membership> {
    let sheets = required_map(txn, SHEETS)?;
    let mut by_source = BTreeMap::new();
    let mut structures = BTreeMap::new();
    let mut relationships = BTreeMap::new();
    for (id, value) in sheets.iter(txn) {
        let Out::YMap(shape) = value else { continue };
        if map_string(&shape, txn, "pageId").as_deref() != Some(page_id) {
            continue;
        }
        let Some(source) = map_number(&shape, txn, "sourceId").map(|value| value as u32) else {
            continue;
        };
        by_source.insert(source, id.to_owned());
        let cells = map_map(&shape, txn, "cells")?;
        for (key, value) in cells.iter(txn) {
            let Out::YMap(cell) = value else { continue };
            if key == "Relationships" {
                if let Some(formula) = map_string(&cell, txn, "formula") {
                    relationships.insert(source, formula);
                }
                continue;
            }
            if key != STRUCTURE_KEY {
                continue;
            }
            let structure =
                map_string(&cell, txn, "formula").or_else(|| map_string(&cell, txn, "value"));
            if let Some(structure) = structure {
                structures.insert(source, structure);
            }
        }
    }
    let mut containers: BTreeMap<u32, BTreeSet<u32>> = BTreeMap::new();
    for (source, structure) in &structures {
        if structure == "Container" {
            containers.entry(*source).or_default();
        }
    }
    for (source, formula) in &relationships {
        for member in dependson_refs(formula, 1) {
            if containers.contains_key(source) && by_source.contains_key(&member) {
                containers.entry(*source).or_default().insert(member);
            }
        }
        for container in dependson_refs(formula, 4) {
            if containers.contains_key(&container) && by_source.contains_key(source) {
                containers.entry(container).or_default().insert(*source);
            }
        }
    }
    Ok(Membership {
        by_source,
        containers,
        structures,
    })
}

fn source_of(membership: &Membership, page_id: &str, shape_id: &str) -> EditResult<u32> {
    membership
        .by_source
        .iter()
        .find_map(|(source, id)| (id == shape_id).then_some(*source))
        .ok_or_else(|| EditError::ShapeNotFound(shape_id.to_owned() + " on " + page_id))
}

fn require_container(membership: &Membership, source: u32) -> EditResult<()> {
    if membership
        .structures
        .get(&source)
        .is_some_and(|structure| structure == "Container")
    {
        Ok(())
    } else {
        Err(EditError::InvalidState(format!(
            "shape {source} is not a container"
        )))
    }
}

fn transitive_members(membership: &Membership, container: u32) -> Vec<u32> {
    let mut seen = BTreeSet::new();
    let mut pending = vec![container];
    while let Some(current) = pending.pop() {
        if let Some(direct) = membership.containers.get(&current) {
            for member in direct {
                if seen.insert(*member) {
                    pending.push(*member);
                }
            }
        }
    }
    seen.into_iter().collect()
}

fn ancestors(membership: &Membership, member: u32) -> Vec<u32> {
    let mut owners: BTreeMap<u32, usize> = BTreeMap::new();
    let mut pending = vec![(member, 0)];
    while let Some((current, depth)) = pending.pop() {
        for (container, members) in &membership.containers {
            if members.contains(&current) && !owners.contains_key(container) {
                owners.insert(*container, depth + 1);
                pending.push((*container, depth + 1));
            }
        }
    }
    let mut ordered: Vec<(u32, usize)> = owners.into_iter().collect();
    ordered.sort_by_key(|(_, depth)| *depth);
    ordered.into_iter().map(|(id, _)| id).collect()
}

fn entry_number<T: ReadTxn>(txn: &T, shape: &str, key: &str) -> Option<f64> {
    let sheets = txn.get_map(SHEETS)?;
    let shape = match sheets.get(txn, shape)? {
        Out::YMap(shape) => shape,
        _ => return None,
    };
    let cells = match shape.get(txn, "cells")? {
        Out::YMap(cells) => cells,
        _ => return None,
    };
    let cell = match cells.get(txn, key)? {
        Out::YMap(cell) => cell,
        _ => return None,
    };
    let formula = map_string(&cell, txn, "formula");
    if let Some(formula) = formula {
        let references = local_references(&cells, txn).ok()?;
        if let Some(value) = evaluate_cached_formula(&formula, &references) {
            return value.parse::<f64>().ok().filter(|value| value.is_finite());
        }
    }
    map_string(&cell, txn, "value")?
        .parse::<f64>()
        .ok()
        .filter(|value| value.is_finite())
}

fn pin<T: ReadTxn>(txn: &T, shape: &str) -> EditResult<(f64, f64)> {
    let x = entry_number(txn, shape, "PinX");
    let y = entry_number(txn, shape, "PinY");
    match (x, y) {
        (Some(x), Some(y)) => Ok((x, y)),
        _ => Err(EditError::InvalidState(
            "cannot evaluate PinX/PinY for move".to_owned(),
        )),
    }
}

fn shape_bounds<T: ReadTxn>(
    txn: &T,
    membership: &Membership,
    source: u32,
    grown: &BTreeMap<u32, Bounds>,
) -> EditResult<Bounds> {
    if let Some(bounds) = grown.get(&source) {
        return Ok(*bounds);
    }
    let shape = &membership.by_source[&source];
    let pin = pin(txn, shape)?;
    let width = entry_number(txn, shape, "Width");
    let height = entry_number(txn, shape, "Height");
    let (Some(width), Some(height)) = (width, height) else {
        return Err(EditError::InvalidState(
            "cannot evaluate Width/Height for autofit".to_owned(),
        ));
    };
    let (loc_pin_x, loc_pin_y) = loc_pin_at_size(txn, shape, width, height)?;
    bounds_at(
        pin.0,
        pin.1,
        width,
        height,
        loc_pin_x,
        loc_pin_y,
        entry_number(txn, shape, "Angle").unwrap_or(0.0),
    )
}

fn moved_bounds<T: ReadTxn>(
    txn: &T,
    membership: &Membership,
    member: u32,
    planned: &[(String, &str, MutationGesture, String)],
) -> EditResult<Bounds> {
    let shape = &membership.by_source[&member];
    let cells = {
        let sheets = required_map(txn, SHEETS)?;
        let shape = map_ref(&sheets, txn, shape)?;
        map_map(&shape, txn, "cells")?
    };
    let mut references = local_references(&cells, txn)?;
    for (target, name, _, formula) in planned {
        if target == shape {
            references.cells.insert(
                (*name).to_owned(),
                vsdx_resolve::Lookup::Found(vsdx_resolve::ResolvedCell {
                    cell: vsdx_parse::Cell {
                        name: (*name).to_owned(),
                        formula: Some(formula.clone()),
                        value: None,
                        unit: None,
                        del: false,
                        other_attrs: Vec::new(),
                    },
                    provenance: vsdx_resolve::Provenance::Local,
                }),
            );
        }
    }
    let number = |name: &str| {
        references
            .cells
            .get(name)
            .and_then(|lookup| match lookup {
                vsdx_resolve::Lookup::Found(cell) => Some(cell),
                _ => None,
            })
            .and_then(|cell| {
                cell.cell
                    .formula
                    .as_deref()
                    .and_then(|formula| evaluate_cached_formula(formula, &references))
            })
            .and_then(|value| value.parse::<f64>().ok())
            .filter(|value| value.is_finite())
    };
    let (Some(x), Some(y), Some(width), Some(height)) = (
        number("PinX"),
        number("PinY"),
        number("Width"),
        number("Height"),
    ) else {
        return Err(EditError::InvalidState(
            "cannot evaluate member bounds for autofit".to_owned(),
        ));
    };
    let loc_pin_x = number("LocPinX").unwrap_or(width / 2.0);
    let loc_pin_y = number("LocPinY").unwrap_or(height / 2.0);
    bounds_at(
        x,
        y,
        width,
        height,
        loc_pin_x,
        loc_pin_y,
        number("Angle").unwrap_or(0.0),
    )
}

fn margin<T: ReadTxn>(txn: &T, membership: &Membership, container: u32) -> f64 {
    entry_number(txn, &membership.by_source[&container], MARGIN_KEY).unwrap_or(0.0)
}

fn locked<T: ReadTxn>(txn: &T, membership: &Membership, container: u32) -> bool {
    entry_number(txn, &membership.by_source[&container], LOCKED_KEY)
        .is_some_and(|value| value == 1.0)
}

fn resize_plan<T: ReadTxn>(
    txn: &T,
    membership: &Membership,
    container: u32,
    outer: Bounds,
) -> EditResult<Vec<(String, &'static str, MutationGesture, String)>> {
    let shape = membership.by_source[&container].clone();
    let width = outer.x1 - outer.x0;
    let height = outer.y1 - outer.y0;
    let (loc_pin_x, loc_pin_y) = loc_pin_at_size(txn, &shape, width, height)?;
    Ok(vec![
        (
            shape.clone(),
            "PinX",
            MutationGesture::MoveX,
            format_number(outer.x0 + loc_pin_x),
        ),
        (
            shape.clone(),
            "PinY",
            MutationGesture::MoveY,
            format_number(outer.y0 + loc_pin_y),
        ),
        (
            shape.clone(),
            "Width",
            MutationGesture::ResizeWidth,
            format_number(outer.x1 - outer.x0),
        ),
        (
            shape,
            "Height",
            MutationGesture::ResizeHeight,
            format_number(outer.y1 - outer.y0),
        ),
    ])
}

fn bounds_at(
    pin_x: f64,
    pin_y: f64,
    width: f64,
    height: f64,
    loc_pin_x: f64,
    loc_pin_y: f64,
    angle: f64,
) -> EditResult<Bounds> {
    let (sin, cos) = angle.sin_cos();
    let mut bounds = Bounds {
        x0: f64::INFINITY,
        y0: f64::INFINITY,
        x1: f64::NEG_INFINITY,
        y1: f64::NEG_INFINITY,
    };
    for (x, y) in [
        (-loc_pin_x, -loc_pin_y),
        (width - loc_pin_x, -loc_pin_y),
        (-loc_pin_x, height - loc_pin_y),
        (width - loc_pin_x, height - loc_pin_y),
    ] {
        let rotated_x = pin_x + x * cos - y * sin;
        let rotated_y = pin_y + x * sin + y * cos;
        bounds.x0 = bounds.x0.min(rotated_x);
        bounds.y0 = bounds.y0.min(rotated_y);
        bounds.x1 = bounds.x1.max(rotated_x);
        bounds.y1 = bounds.y1.max(rotated_y);
    }
    if bounds.x0.is_finite()
        && bounds.y0.is_finite()
        && bounds.x1.is_finite()
        && bounds.y1.is_finite()
    {
        Ok(bounds)
    } else {
        Err(EditError::InvalidState(
            "cannot evaluate member bounds for autofit".to_owned(),
        ))
    }
}

fn format_number(value: f64) -> String {
    format!("{value}")
}

fn apply(
    txn: &mut TransactionMut<'_>,
    page_id: &str,
    planned: &[(String, &str, MutationGesture, String)],
) -> EditResult<Vec<CellFormulaReceipt>> {
    let mut decided = Vec::with_capacity(planned.len());
    for (shape, name, gesture, formula) in planned {
        let policy = CrdtMutationContext::new(txn, page_id, shape)?;
        let target = match vsdx_eval::decide_mutation(
            &policy,
            policy.locator(CellLocator {
                sheet: CellSheet::Page(0),
                shape_id: None,
                section: None,
                section_index: None,
                row: None,
                cell_name: (*name).to_owned(),
            }),
            *gesture,
            formula.clone(),
            &vsdx_parse::ParseLimits::default(),
        ) {
            vsdx_eval::MutationOutcome::Allowed { target, .. } => target,
            vsdx_eval::MutationOutcome::Refused { reason }
            | vsdx_eval::MutationOutcome::Unsupported { reason } => {
                return Err(EditError::InvalidState(reason));
            }
        };
        decided.push((shape.clone(), target, formula.clone()));
    }
    let mut prepared = Vec::with_capacity(decided.len());
    for (shape, target, formula) in &decided {
        let cell = cell_map(txn, page_id, shape, target)?;
        prepared.push((shape.clone(), target.clone(), formula.clone(), cell));
    }
    Ok(prepared
        .iter()
        .map(|(shape, target, formula, cell)| {
            let before = map_string(cell, txn, "formula");
            cell.insert(txn, "formula", formula.as_str());
            CellFormulaReceipt {
                page_id: page_id.to_owned(),
                shape_id: shape.clone(),
                cell_name: target.cell_name.clone(),
                before,
                after: formula.clone(),
            }
        })
        .collect())
}

impl DiagramSession {
    pub fn move_container(
        &self,
        context: &EditCtx,
        page_id: &str,
        shape_id: &str,
        dx: f64,
        dy: f64,
    ) -> EditResult<Vec<CellFormulaReceipt>> {
        move_container(self, context, page_id, shape_id, dx, dy)
    }

    pub fn move_container_member(
        &self,
        context: &EditCtx,
        page_id: &str,
        shape_id: &str,
        x_formula: String,
        y_formula: String,
    ) -> EditResult<Vec<CellFormulaReceipt>> {
        move_container_member(self, context, page_id, shape_id, x_formula, y_formula)
    }

    pub fn autofit_container(
        &self,
        context: &EditCtx,
        page_id: &str,
        shape_id: &str,
    ) -> EditResult<Vec<CellFormulaReceipt>> {
        autofit_container(self, context, page_id, shape_id)
    }
}
