//! Read-only validation rules over resolved VSDX pages.

use std::collections::{BTreeMap, HashMap, HashSet};

use serde::{Deserialize, Serialize};
use vsdx_parse::{Shape, Sheet, VsdxPackage};
use vsdx_resolve::{Lookup, PageConnectivity, ResolvedShape, Resolver, SceneAffine, shape_data};

pub const RULE_DANGLING_CONNECTOR: &str = "dangling-connector";
pub const RULE_ISOLATED_SHAPE: &str = "isolated-shape";
pub const RULE_OVERLAPPING_SHAPES: &str = "overlapping-shapes";
pub const RULE_CONNECTOR_CROSSING: &str = "connector-crossing";
pub const RULE_EMPTY_SHAPE_DATA: &str = "empty-shape-data";

/// Pages with more shapes than this skip the two pairwise rules, which are quadratic.
pub const MAX_PAIRWISE_SHAPES: usize = 2_000;

/// Severity of a validation issue.
#[derive(Clone, Copy, Debug, PartialEq, Eq, PartialOrd, Ord, Serialize, Deserialize)]
#[serde(rename_all = "lowercase")]
pub enum Severity {
    Warning,
    Error,
}

/// Static descriptor for one validation rule.
#[derive(Clone, Debug, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct RuleDescriptor {
    pub id: String,
    pub title: String,
    pub severity: Severity,
}

/// One stable, sortable validation finding. Never mutates the document.
#[derive(Clone, Debug, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ValidationIssue {
    pub id: String,
    pub rule: String,
    pub severity: Severity,
    pub page_part: String,
    pub page_id: Option<u32>,
    pub shape_id: u32,
    pub other_shape_id: Option<u32>,
    pub endpoint: Option<String>,
    pub row: Option<String>,
}

/// Sorted validation findings for a document.
#[derive(Clone, Debug, Default, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ValidationReport {
    pub issues: Vec<ValidationIssue>,
}

/// Descriptors for every rule the engine implements, in stable order.
pub fn rules() -> Vec<RuleDescriptor> {
    [
        (
            RULE_DANGLING_CONNECTOR,
            "Connector glued at one end or neither",
            Severity::Error,
        ),
        (
            RULE_ISOLATED_SHAPE,
            "Shape with no connections in a connected diagram",
            Severity::Warning,
        ),
        (
            RULE_OVERLAPPING_SHAPES,
            "Overlapping shapes",
            Severity::Warning,
        ),
        (
            RULE_CONNECTOR_CROSSING,
            "Connector crossing a shape it is not glued to",
            Severity::Warning,
        ),
        (
            RULE_EMPTY_SHAPE_DATA,
            "Empty required Shape Data field",
            Severity::Warning,
        ),
    ]
    .into_iter()
    .map(|(id, title, severity)| RuleDescriptor {
        id: id.to_owned(),
        title: title.to_owned(),
        severity,
    })
    .collect()
}

/// Validates every page in catalog order. Read-only.
pub fn validate_package(package: &VsdxPackage) -> ValidationReport {
    let mut issues = Vec::new();
    for page_part in &package.page_part_paths {
        issues.extend(validate_page(package, page_part));
    }
    ValidationReport { issues }
}

/// Validates one page part. Returns no issues when the page cannot resolve, and
/// skips the pairwise rules beyond `MAX_PAIRWISE_SHAPES`.
pub fn validate_page(package: &VsdxPackage, page_part: &str) -> Vec<ValidationIssue> {
    let resolver = Resolver::new(package);
    let Ok(shapes) = resolver.resolve_page_shapes(page_part) else {
        return Vec::new();
    };
    let connectivity = resolver
        .resolve_page_connectivity_with(page_part, &shapes)
        .unwrap_or_default();
    let page_id = package.page_part_ids.get(page_part).copied();
    let tree = index_tree(package.page_contents.get(page_part));
    let bounds = scene_bounds(package.page_contents.get(page_part), &shapes, &tree);
    let mut issues = Vec::new();
    issues.extend(dangling_connectors(
        page_part,
        page_id,
        &shapes,
        &connectivity,
        &tree,
    ));
    issues.extend(isolated_shapes(
        page_part,
        page_id,
        &shapes,
        &connectivity,
        &tree,
    ));
    if bounds.len() <= MAX_PAIRWISE_SHAPES {
        issues.extend(overlapping_shapes(
            page_part, page_id, &shapes, &bounds, &tree,
        ));
        issues.extend(connector_crossings(
            page_part,
            page_id,
            &shapes,
            &connectivity,
            &bounds,
            &tree,
        ));
    }
    issues.extend(empty_shape_data(page_part, page_id, &shapes));
    issues.sort();
    issues
}

impl Ord for ValidationIssue {
    fn cmp(&self, other: &Self) -> std::cmp::Ordering {
        (
            &self.rule,
            &self.page_part,
            self.shape_id,
            self.other_shape_id,
            &self.endpoint,
            &self.row,
            &self.id,
        )
            .cmp(&(
                &other.rule,
                &other.page_part,
                other.shape_id,
                other.other_shape_id,
                &other.endpoint,
                &other.row,
                &other.id,
            ))
    }
}

impl PartialOrd for ValidationIssue {
    fn partial_cmp(&self, other: &Self) -> Option<std::cmp::Ordering> {
        Some(self.cmp(other))
    }
}

#[derive(Clone, Copy, Debug)]
struct Bounds {
    min_x: f64,
    min_y: f64,
    max_x: f64,
    max_y: f64,
}

impl Bounds {
    fn area(&self) -> f64 {
        (self.max_x - self.min_x).max(0.0) * (self.max_y - self.min_y).max(0.0)
    }
}

#[derive(Clone, Debug, Default)]
struct ShapeNode {
    parent: Option<u32>,
}

#[derive(Clone, Debug, Default)]
struct ShapeTree {
    nodes: HashMap<u32, ShapeNode>,
    children: HashMap<u32, Vec<u32>>,
}

fn index_tree(page: Option<&Sheet>) -> ShapeTree {
    let mut tree = ShapeTree::default();
    let Some(page) = page else {
        return tree;
    };
    let mut stack: Vec<(&Shape, Option<u32>)> = page.shapes().map(|shape| (shape, None)).collect();
    while let Some((shape, parent)) = stack.pop() {
        let children: Vec<&Shape> = shape.shapes().collect();
        tree.nodes.insert(shape.id, ShapeNode { parent });
        tree.children
            .insert(shape.id, children.iter().map(|child| child.id).collect());
        for child in children {
            stack.push((child, Some(shape.id)));
        }
    }
    tree
}

fn ancestors(tree: &ShapeTree, id: u32) -> HashSet<u32> {
    let mut out = HashSet::new();
    let mut current = tree.nodes.get(&id).and_then(|node| node.parent);
    while let Some(parent) = current {
        if !out.insert(parent) {
            break;
        }
        current = tree.nodes.get(&parent).and_then(|node| node.parent);
    }
    out
}

fn is_group(tree: &ShapeTree, id: u32) -> bool {
    tree.children
        .get(&id)
        .is_some_and(|children| !children.is_empty())
}

fn cell_number(shape: &ResolvedShape, name: &str) -> Option<f64> {
    let Lookup::Found(value) = shape.cell(name)? else {
        return None;
    };
    if let Some(number) = value
        .cell
        .formula
        .as_deref()
        .and_then(|formula| formula_number(shape, formula))
    {
        return Some(number);
    }
    value
        .cell
        .value
        .as_deref()?
        .parse::<f64>()
        .ok()
        .filter(|value| value.is_finite())
}

fn formula_number(shape: &ResolvedShape, formula: &str) -> Option<f64> {
    vsdx_formula::evaluate_number(
        formula,
        vsdx_formula::Limits {
            max_depth: 64,
            max_nodes: 1_024,
            max_tokens: 1_024,
        },
        &mut |name| {
            let Lookup::Found(value) = shape.cell(name.trim())? else {
                return None;
            };
            value
                .cell
                .formula
                .clone()
                .or_else(|| value.cell.value.clone())
        },
    )
}

fn is_1d(shape: &ResolvedShape) -> bool {
    match cell_number(shape, "OneD") {
        Some(value) => value != 0.0,
        None => ["BeginX", "BeginY", "EndX", "EndY"]
            .into_iter()
            .all(|name| cell_number(shape, name).is_some()),
    }
}

fn live_shape(shapes: &BTreeMap<u32, ResolvedShape>, id: u32) -> Option<&ResolvedShape> {
    shapes.get(&id).filter(|shape| !shape.deleted)
}

fn scene_bounds(
    page: Option<&Sheet>,
    shapes: &BTreeMap<u32, ResolvedShape>,
    tree: &ShapeTree,
) -> HashMap<u32, Bounds> {
    let Some(page) = page else {
        return HashMap::new();
    };
    let transforms =
        vsdx_resolve::scene_transforms(page, shapes, |_, shape, name| cell_number(shape, name));
    let mut bounds = HashMap::new();
    for (id, shape) in shapes {
        if shape.deleted || is_1d(shape) || is_group(tree, *id) {
            continue;
        }
        let (Some(width), Some(height)) =
            (cell_number(shape, "Width"), cell_number(shape, "Height"))
        else {
            continue;
        };
        if !(width > 0.0 && height > 0.0) {
            continue;
        }
        let Some(transform) = transforms.get(id).map(|entry| entry.scene) else {
            continue;
        };
        bounds.insert(*id, affine_bounds(transform, width, height));
    }
    let mut memo = HashMap::new();
    for id in shapes.keys() {
        if bounds.contains_key(id) {
            continue;
        }
        if let Some(union) = group_union(*id, shapes, tree, &bounds, &mut memo) {
            bounds.insert(*id, union);
        }
    }
    bounds
}

fn affine_bounds(transform: SceneAffine, width: f64, height: f64) -> Bounds {
    let corners = [
        transform.apply_point(0.0, 0.0),
        transform.apply_point(width, 0.0),
        transform.apply_point(0.0, height),
        transform.apply_point(width, height),
    ];
    Bounds {
        min_x: corners
            .iter()
            .map(|point| point.x)
            .reduce(f64::min)
            .unwrap_or(0.0),
        min_y: corners
            .iter()
            .map(|point| point.y)
            .reduce(f64::min)
            .unwrap_or(0.0),
        max_x: corners
            .iter()
            .map(|point| point.x)
            .reduce(f64::max)
            .unwrap_or(0.0),
        max_y: corners
            .iter()
            .map(|point| point.y)
            .reduce(f64::max)
            .unwrap_or(0.0),
    }
}

fn group_union(
    id: u32,
    shapes: &BTreeMap<u32, ResolvedShape>,
    tree: &ShapeTree,
    bounds: &HashMap<u32, Bounds>,
    memo: &mut HashMap<u32, Option<Bounds>>,
) -> Option<Bounds> {
    if let Some(cached) = memo.get(&id) {
        return *cached;
    }
    memo.insert(id, None);
    let mut union: Option<Bounds> = None;
    for child in tree.children.get(&id).cloned().unwrap_or_default() {
        let candidate = bounds.get(&child).copied().or_else(|| {
            if is_group(tree, child) {
                group_union(child, shapes, tree, bounds, memo)
            } else {
                None
            }
        });
        union = match (union, candidate) {
            (Some(left), Some(right)) => Some(Bounds {
                min_x: left.min_x.min(right.min_x),
                min_y: left.min_y.min(right.min_y),
                max_x: left.max_x.max(right.max_x),
                max_y: left.max_y.max(right.max_y),
            }),
            (Some(left), None) => Some(left),
            (None, right) => right,
        };
    }
    if shapes.get(&id).is_some_and(|shape| shape.deleted) {
        union = None;
    }
    memo.insert(id, union);
    union
}

fn dangling_connectors(
    page_part: &str,
    page_id: Option<u32>,
    shapes: &BTreeMap<u32, ResolvedShape>,
    connectivity: &PageConnectivity,
    tree: &ShapeTree,
) -> Vec<ValidationIssue> {
    let mut issues = Vec::new();
    for connector in connectivity.connectors.values() {
        if !connector.is_1d {
            continue;
        }
        if live_shape(shapes, connector.shape_id).is_none()
            || !tree.nodes.contains_key(&connector.shape_id)
        {
            continue;
        }
        let glued = |endpoint: vsdx_resolve::ConnectorEndpoint| {
            connector
                .glue
                .iter()
                .any(|glue| glue.endpoint == endpoint && glue.to.is_some())
        };
        let begin = glued(vsdx_resolve::ConnectorEndpoint::Begin);
        let end = glued(vsdx_resolve::ConnectorEndpoint::End);
        if begin && end {
            continue;
        }
        let endpoint = match (begin, end) {
            (false, true) => Some("begin".to_owned()),
            (true, false) => Some("end".to_owned()),
            _ => None,
        };
        issues.push(ValidationIssue {
            id: format!(
                "{RULE_DANGLING_CONNECTOR}:{page_part}:{}",
                connector.shape_id
            ),
            rule: RULE_DANGLING_CONNECTOR.to_owned(),
            severity: Severity::Error,
            page_part: page_part.to_owned(),
            page_id,
            shape_id: connector.shape_id,
            other_shape_id: None,
            endpoint,
            row: None,
        });
    }
    issues
}

fn isolated_shapes(
    page_part: &str,
    page_id: Option<u32>,
    shapes: &BTreeMap<u32, ResolvedShape>,
    connectivity: &PageConnectivity,
    tree: &ShapeTree,
) -> Vec<ValidationIssue> {
    let mut degree: HashMap<u32, u32> = HashMap::new();
    for connector in connectivity.connectors.values() {
        for glue in &connector.glue {
            let Some(target) = glue.to.as_ref() else {
                continue;
            };
            *degree.entry(connector.shape_id).or_default() += 1;
            *degree.entry(target.shape_id).or_default() += 1;
        }
    }
    if degree.is_empty() {
        return Vec::new();
    }
    shapes
        .iter()
        .filter(|(id, shape)| {
            !shape.deleted
                && !is_1d(shape)
                && !is_group(tree, **id)
                && tree.nodes.contains_key(*id)
                && degree.get(*id).copied().unwrap_or(0) == 0
        })
        .map(|(id, _)| ValidationIssue {
            id: format!("{RULE_ISOLATED_SHAPE}:{page_part}:{id}"),
            rule: RULE_ISOLATED_SHAPE.to_owned(),
            severity: Severity::Warning,
            page_part: page_part.to_owned(),
            page_id,
            shape_id: *id,
            other_shape_id: None,
            endpoint: None,
            row: None,
        })
        .collect()
}

fn overlapping_shapes(
    page_part: &str,
    page_id: Option<u32>,
    shapes: &BTreeMap<u32, ResolvedShape>,
    bounds: &HashMap<u32, Bounds>,
    tree: &ShapeTree,
) -> Vec<ValidationIssue> {
    let mut ids: Vec<u32> = bounds
        .keys()
        .copied()
        .filter(|id| live_shape(shapes, *id).is_some())
        .collect();
    ids.sort_by(|left, right| {
        bounds[left]
            .min_x
            .total_cmp(&bounds[right].min_x)
            .then_with(|| left.cmp(right))
    });
    let mut issues = Vec::new();
    for (index, left) in ids.iter().enumerate() {
        let left_bounds = &bounds[left];
        if left_bounds.area() <= 0.0 {
            continue;
        }
        let left_ancestors = ancestors(tree, *left);
        for right in ids.iter().skip(index + 1) {
            let right_bounds = &bounds[right];
            if right_bounds.min_x >= left_bounds.max_x {
                break;
            }
            if right_bounds.area() <= 0.0 {
                continue;
            }
            if left_ancestors.contains(right) || ancestors(tree, *right).contains(left) {
                continue;
            }
            if overlap_area(left_bounds, right_bounds) > 1e-9 {
                let (first, second) = if left < right {
                    (left, right)
                } else {
                    (right, left)
                };
                issues.push(ValidationIssue {
                    id: format!("{RULE_OVERLAPPING_SHAPES}:{page_part}:{first}:{second}"),
                    rule: RULE_OVERLAPPING_SHAPES.to_owned(),
                    severity: Severity::Warning,
                    page_part: page_part.to_owned(),
                    page_id,
                    shape_id: *first,
                    other_shape_id: Some(*second),
                    endpoint: None,
                    row: None,
                });
            }
        }
    }
    issues
}

fn overlap_area(left: &Bounds, right: &Bounds) -> f64 {
    let width = left.max_x.min(right.max_x) - left.min_x.max(right.min_x);
    let height = left.max_y.min(right.max_y) - left.min_y.max(right.min_y);
    if width <= 0.0 || height <= 0.0 {
        0.0
    } else {
        width * height
    }
}

fn connector_crossings(
    page_part: &str,
    page_id: Option<u32>,
    shapes: &BTreeMap<u32, ResolvedShape>,
    connectivity: &PageConnectivity,
    bounds: &HashMap<u32, Bounds>,
    tree: &ShapeTree,
) -> Vec<ValidationIssue> {
    let mut issues = Vec::new();
    for connector in connectivity.connectors.values() {
        if !connector.is_1d {
            continue;
        }
        let Some(shape) = live_shape(shapes, connector.shape_id) else {
            continue;
        };
        let (Some(begin), Some(end)) = (
            cell_number(shape, "BeginX").zip(cell_number(shape, "BeginY")),
            cell_number(shape, "EndX").zip(cell_number(shape, "EndY")),
        ) else {
            continue;
        };
        if begin == end {
            continue;
        }
        let glued: HashSet<u32> = connector
            .glue
            .iter()
            .filter_map(|glue| glue.to.as_ref().map(|target| target.shape_id))
            .collect();
        let kin: HashSet<u32> =
            ancestors(tree, connector.shape_id)
                .into_iter()
                .chain(tree.nodes.iter().filter_map(|(id, node)| {
                    (node.parent == Some(connector.shape_id)).then_some(*id)
                }))
                .collect();
        for (id, candidate) in bounds {
            if *id == connector.shape_id || glued.contains(id) || kin.contains(id) {
                continue;
            }
            if live_shape(shapes, *id).is_none() || is_group(tree, *id) {
                continue;
            }
            if segment_crosses_bounds(begin, end, candidate) {
                issues.push(ValidationIssue {
                    id: format!(
                        "{RULE_CONNECTOR_CROSSING}:{page_part}:{}:{id}",
                        connector.shape_id
                    ),
                    rule: RULE_CONNECTOR_CROSSING.to_owned(),
                    severity: Severity::Warning,
                    page_part: page_part.to_owned(),
                    page_id,
                    shape_id: connector.shape_id,
                    other_shape_id: Some(*id),
                    endpoint: None,
                    row: None,
                });
            }
        }
    }
    issues
}

fn strictly_inside(point: (f64, f64), bounds: &Bounds) -> bool {
    const EPSILON: f64 = 1e-9;
    point.0 > bounds.min_x + EPSILON
        && point.0 < bounds.max_x - EPSILON
        && point.1 > bounds.min_y + EPSILON
        && point.1 < bounds.max_y - EPSILON
}

fn segment_crosses_bounds(start: (f64, f64), end: (f64, f64), bounds: &Bounds) -> bool {
    if strictly_inside(start, bounds) || strictly_inside(end, bounds) {
        return true;
    }
    let Some((enter, exit)) = clip_segment(start, end, bounds) else {
        return false;
    };
    if exit - enter <= 1e-9 {
        return false;
    }
    let mid = enter + (exit - enter) / 2.0;
    strictly_inside(
        (
            start.0 + (end.0 - start.0) * mid,
            start.1 + (end.1 - start.1) * mid,
        ),
        bounds,
    )
}

fn clip_segment(start: (f64, f64), end: (f64, f64), bounds: &Bounds) -> Option<(f64, f64)> {
    let direction = (end.0 - start.0, end.1 - start.1);
    let (mut enter, mut exit) = (0.0_f64, 1.0_f64);
    for (origin, step, min, max) in [
        (start.0, direction.0, bounds.min_x, bounds.max_x),
        (start.1, direction.1, bounds.min_y, bounds.max_y),
    ] {
        if step.abs() <= f64::EPSILON {
            if origin < min || origin > max {
                return None;
            }
            continue;
        }
        let mut near = (min - origin) / step;
        let mut far = (max - origin) / step;
        if near > far {
            std::mem::swap(&mut near, &mut far);
        }
        enter = enter.max(near);
        exit = exit.min(far);
        if enter > exit {
            return None;
        }
    }
    Some((enter, exit))
}

fn empty_shape_data(
    page_part: &str,
    page_id: Option<u32>,
    shapes: &BTreeMap<u32, ResolvedShape>,
) -> Vec<ValidationIssue> {
    let mut issues = Vec::new();
    for (id, shape) in shapes {
        if shape.deleted {
            continue;
        }
        for property in shape_data(shape) {
            if !property.ask || !shape_data_empty(&property.value) {
                continue;
            }
            issues.push(ValidationIssue {
                id: format!("{RULE_EMPTY_SHAPE_DATA}:{page_part}:{id}:{}", property.row),
                rule: RULE_EMPTY_SHAPE_DATA.to_owned(),
                severity: Severity::Warning,
                page_part: page_part.to_owned(),
                page_id,
                shape_id: *id,
                other_shape_id: None,
                endpoint: None,
                row: Some(property.row.clone()),
            });
        }
    }
    issues
}

fn shape_data_empty(value: &vsdx_resolve::ShapeDataValue) -> bool {
    if value
        .value
        .as_deref()
        .is_some_and(|text| !text.trim().is_empty())
    {
        return false;
    }
    let Some(formula) = value.formula.as_deref() else {
        return true;
    };
    let expression = formula.trim().trim_start_matches('=').trim();
    if expression.is_empty() {
        return true;
    }
    match expression
        .strip_prefix('"')
        .and_then(|inner| inner.strip_suffix('"'))
    {
        Some(inner) => inner.replace("\"\"", "\"").trim().is_empty(),
        None => false,
    }
}

#[cfg(test)]
mod tests {
    use vsdx_parse::{
        Cell, Connect, ConnectsChild, Row, RowChild, Section, SectionChild, Shape, ShapeChild,
        ShapesChild, Sheet, SheetChild,
    };

    use super::*;

    fn cell(name: &str, value: &str) -> Cell {
        Cell {
            name: name.into(),
            formula: None,
            value: Some(value.into()),
            unit: None,
            del: false,
            other_attrs: Vec::new(),
        }
    }

    fn placed(id: u32, pin_x: f64, pin_y: f64) -> Shape {
        Shape {
            id,
            name: None,
            name_u: None,
            shape_type: Some("Shape".into()),
            master: None,
            master_shape: None,
            line_style: None,
            fill_style: None,
            text_style: None,
            children: ["Width", "Height"]
                .into_iter()
                .map(|name| ShapeChild::Cell(cell(name, "1")))
                .chain(
                    [
                        ("PinX", pin_x),
                        ("PinY", pin_y),
                        ("LocPinX", 0.5),
                        ("LocPinY", 0.5),
                    ]
                    .into_iter()
                    .map(|(name, value)| ShapeChild::Cell(cell(name, &value.to_string()))),
                )
                .collect(),
            del: false,
            other_attrs: Vec::new(),
        }
    }

    fn connector(id: u32, begin: (f64, f64), end: (f64, f64)) -> Shape {
        let mut shape = placed(id, 0.0, 0.0);
        shape.children.push(ShapeChild::Cell(cell("OneD", "1")));
        for (name, value) in [
            ("BeginX", begin.0),
            ("BeginY", begin.1),
            ("EndX", end.0),
            ("EndY", end.1),
        ] {
            shape
                .children
                .push(ShapeChild::Cell(cell(name, &value.to_string())));
        }
        shape
    }

    fn glue(from: u32, from_cell: &str, to: u32, to_cell: &str) -> Connect {
        Connect {
            from_sheet: from,
            from_cell: Some(from_cell.into()),
            from_part: Some(9),
            to_sheet: to,
            to_cell: Some(to_cell.into()),
            to_part: Some(3),
            other_attrs: Vec::new(),
        }
    }

    fn property_row(name: &str, cells: Vec<Cell>) -> Row {
        Row {
            index: None,
            name: Some(name.into()),
            local_name: None,
            row_type: None,
            del: false,
            children: cells.into_iter().map(RowChild::Cell).collect(),
            other_attrs: Vec::new(),
        }
    }

    fn package(shapes: Vec<Shape>, connects: Vec<Connect>) -> VsdxPackage {
        let sheet = Sheet {
            id: None,
            children: vec![
                SheetChild::Shapes(shapes.into_iter().map(ShapesChild::Shape).collect()),
                SheetChild::Connects(connects.into_iter().map(ConnectsChild::Connect).collect()),
            ],
            other_attrs: Vec::new(),
        };
        serde_json::from_value(serde_json::json!({
            "documentPartPath": "",
            "pagesPartPath": null,
            "mastersPartPath": null,
            "pagePartPaths": ["page"],
            "masterPartPaths": [],
            "themePartPaths": [],
            "windowsPartPath": null,
            "relationships": {},
            "documentSheet": null,
            "styleSheets": [],
            "colors": [],
            "faceNames": [],
            "pageSheets": {},
            "masterSheets": {},
            "pagePartIds": {"page": 1},
            "pageNames": {"1": "Page-1"},
            "masterPartIds": {},
            "pageContents": {"page": serde_json::to_value(&sheet).unwrap()},
            "masterContents": {}
        }))
        .unwrap()
    }

    fn rules_for(package: &VsdxPackage, rule: &str) -> Vec<ValidationIssue> {
        validate_package(package)
            .issues
            .into_iter()
            .filter(|issue| issue.rule == rule)
            .collect()
    }

    #[test]
    fn connector_glued_at_one_end_reports_that_endpoint() {
        let package = package(
            vec![placed(1, 1.0, 1.0), connector(2, (1.0, 1.0), (5.0, 5.0))],
            vec![glue(2, "BeginX", 1, "PinX")],
        );
        let issues = rules_for(&package, RULE_DANGLING_CONNECTOR);
        assert_eq!(issues.len(), 1);
        assert_eq!(issues[0].shape_id, 2);
        assert_eq!(issues[0].endpoint.as_deref(), Some("end"));
        assert_eq!(issues[0].severity, Severity::Error);
    }

    #[test]
    fn connector_glued_at_neither_end_reports_no_endpoint() {
        let package = package(
            vec![placed(1, 1.0, 1.0), connector(2, (1.0, 1.0), (5.0, 5.0))],
            vec![],
        );
        let issues = rules_for(&package, RULE_DANGLING_CONNECTOR);
        assert_eq!(issues.len(), 1);
        assert_eq!(issues[0].endpoint, None);
    }

    #[test]
    fn connector_glued_at_both_ends_is_clean() {
        let package = package(
            vec![
                placed(1, 1.0, 1.0),
                placed(2, 5.0, 5.0),
                connector(3, (1.0, 1.0), (5.0, 5.0)),
            ],
            vec![glue(3, "BeginX", 1, "PinX"), glue(3, "EndX", 2, "PinX")],
        );
        assert!(rules_for(&package, RULE_DANGLING_CONNECTOR).is_empty());
    }

    #[test]
    fn unconnected_shape_in_connected_diagram_is_isolated() {
        let package = package(
            vec![
                placed(1, 1.0, 1.0),
                placed(2, 5.0, 5.0),
                placed(3, 9.0, 9.0),
                connector(4, (1.0, 1.0), (5.0, 5.0)),
            ],
            vec![glue(4, "BeginX", 1, "PinX"), glue(4, "EndX", 2, "PinX")],
        );
        let issues = rules_for(&package, RULE_ISOLATED_SHAPE);
        assert_eq!(
            issues
                .iter()
                .map(|issue| issue.shape_id)
                .collect::<Vec<_>>(),
            [3]
        );
    }

    #[test]
    fn diagram_without_glue_has_no_isolated_shapes() {
        let package = package(vec![placed(1, 1.0, 1.0), placed(2, 5.0, 5.0)], vec![]);
        assert!(rules_for(&package, RULE_ISOLATED_SHAPE).is_empty());
    }

    #[test]
    fn overlapping_shapes_share_one_stable_issue() {
        let package = package(vec![placed(1, 1.0, 1.0), placed(2, 1.5, 1.0)], vec![]);
        let issues = rules_for(&package, RULE_OVERLAPPING_SHAPES);
        assert_eq!(issues.len(), 1);
        assert_eq!(issues[0].shape_id, 1);
        assert_eq!(issues[0].other_shape_id, Some(2));
        assert_eq!(issues[0].id, format!("{RULE_OVERLAPPING_SHAPES}:page:1:2"));
    }

    #[test]
    fn edge_touching_shapes_do_not_overlap() {
        let package = package(vec![placed(1, 1.0, 1.0), placed(2, 2.0, 1.0)], vec![]);
        assert!(rules_for(&package, RULE_OVERLAPPING_SHAPES).is_empty());
    }

    #[test]
    fn connector_crossing_an_unglued_shape_reports_it() {
        let package = package(
            vec![
                placed(1, 1.0, 1.0),
                placed(2, 9.0, 1.0),
                placed(3, 5.0, 1.0),
                connector(4, (1.0, 1.0), (9.0, 1.0)),
            ],
            vec![glue(4, "BeginX", 1, "PinX"), glue(4, "EndX", 2, "PinX")],
        );
        let issues = rules_for(&package, RULE_CONNECTOR_CROSSING);
        assert_eq!(issues.len(), 1);
        assert_eq!(issues[0].shape_id, 4);
        assert_eq!(issues[0].other_shape_id, Some(3));
    }

    #[test]
    fn connector_crossing_only_glued_shapes_is_clean() {
        let package = package(
            vec![
                placed(1, 1.0, 1.0),
                placed(2, 9.0, 1.0),
                connector(3, (1.0, 1.0), (9.0, 1.0)),
            ],
            vec![glue(3, "BeginX", 1, "PinX"), glue(3, "EndX", 2, "PinX")],
        );
        assert!(rules_for(&package, RULE_CONNECTOR_CROSSING).is_empty());
    }

    #[test]
    fn empty_required_shape_data_reports_its_row() {
        let mut shape = placed(1, 1.0, 1.0);
        shape.children.push(ShapeChild::Section(Section {
            name: "Property".into(),
            index: None,
            del: false,
            other_attrs: Vec::new(),
            children: vec![
                SectionChild::Row(property_row(
                    "Owner",
                    vec![
                        cell("Label", "Owner"),
                        cell("Value", ""),
                        cell("Type", "0"),
                        cell("Ask", "1"),
                    ],
                )),
                SectionChild::Row(property_row(
                    "Notes",
                    vec![
                        cell("Label", "Notes"),
                        cell("Value", ""),
                        cell("Type", "0"),
                        cell("Ask", "0"),
                    ],
                )),
                SectionChild::Row(property_row(
                    "Filled",
                    vec![
                        cell("Label", "Filled"),
                        cell("Value", "ready"),
                        cell("Type", "0"),
                        cell("Ask", "1"),
                    ],
                )),
            ],
        }));
        let package = package(vec![shape], vec![]);
        let issues = rules_for(&package, RULE_EMPTY_SHAPE_DATA);
        assert_eq!(issues.len(), 1);
        assert_eq!(issues[0].shape_id, 1);
        assert_eq!(issues[0].row.as_deref(), Some("Owner"));
    }

    #[test]
    fn report_is_deterministic_and_stable_across_runs() {
        let package = package(
            vec![
                placed(2, 1.5, 1.0),
                placed(1, 1.0, 1.0),
                connector(3, (1.0, 1.0), (5.0, 5.0)),
            ],
            vec![],
        );
        let first = validate_package(&package);
        let second = validate_package(&package);
        assert_eq!(first, second);
        let mut ids: Vec<_> = first.issues.iter().map(|issue| &issue.id).collect();
        ids.sort();
        ids.dedup();
        assert_eq!(ids.len(), first.issues.len());
        let sorted = {
            let mut issues = first.issues.clone();
            issues.sort();
            issues
        };
        assert_eq!(first.issues, sorted);
    }

    #[test]
    fn rules_describe_five_default_checks() {
        let ids: Vec<_> = rules().iter().map(|rule| rule.id.clone()).collect();
        assert_eq!(
            ids,
            [
                RULE_DANGLING_CONNECTOR,
                RULE_ISOLATED_SHAPE,
                RULE_OVERLAPPING_SHAPES,
                RULE_CONNECTOR_CROSSING,
                RULE_EMPTY_SHAPE_DATA,
            ]
        );
    }
}
