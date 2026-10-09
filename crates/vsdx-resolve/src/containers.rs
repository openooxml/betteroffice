//! Container membership resolved from ShapeSheet data.
//!
//! A Visio container is an ordinary shape whose inherited
//! `User.msvStructureType` is `Container`. Membership travels in the
//! page-local `Relationships` cell: `DEPENDSON(1, ...)` on the container
//! lists its members, `DEPENDSON(4, ...)` on a member lists its containers.

use std::collections::{BTreeMap, BTreeSet};

use serde::{Deserialize, Serialize};
use vsdx_formula::{Limits as FormulaLimits, evaluate_number};

use crate::{Lookup, ResolveError, ResolvedShape, Resolver};

/// A container shape and the members it owns.
#[derive(Clone, Debug, Default, PartialEq, Serialize, Deserialize)]
pub struct ContainerInfo {
    pub shape_id: u32,
    pub member_ids: Vec<u32>,
    pub margin: Option<f64>,
    pub resize: Option<i64>,
    pub locked: bool,
}

/// Container membership for one page.
#[derive(Clone, Debug, Default, PartialEq, Serialize, Deserialize)]
pub struct PageContainers {
    pub containers: BTreeMap<u32, ContainerInfo>,
    pub member_to_containers: BTreeMap<u32, Vec<u32>>,
}

impl PageContainers {
    /// Direct members of a container shape.
    pub fn members_of(&self, shape_id: u32) -> &[u32] {
        self.containers
            .get(&shape_id)
            .map(|info| info.member_ids.as_slice())
            .unwrap_or_default()
    }

    /// Containers directly owning a member shape.
    pub fn containers_of(&self, shape_id: u32) -> &[u32] {
        self.member_to_containers
            .get(&shape_id)
            .map(Vec::as_slice)
            .unwrap_or_default()
    }

    /// Direct and nested members of a container shape.
    pub fn transitive_members_of(&self, shape_id: u32) -> Vec<u32> {
        let mut seen = BTreeSet::new();
        let mut pending = vec![shape_id];
        while let Some(current) = pending.pop() {
            for member in self.members_of(current) {
                if seen.insert(*member) {
                    pending.push(*member);
                }
            }
        }
        seen.into_iter().collect()
    }
}

/// Shape IDs referenced by `DEPENDSON(kind, Sheet.*!...)` groups.
pub fn dependson_refs(formula: &str, kind: u32) -> Vec<u32> {
    let mut out = Vec::new();
    let bytes = formula.as_bytes();
    let mut index = 0;
    while index < bytes.len() {
        let rest = &formula[index..];
        let Some(offset) = find_call(rest, "DEPENDSON") else {
            break;
        };
        index += offset;
        let Some((group, length)) = balanced_group(&formula[index..]) else {
            break;
        };
        index += length;
        let mut parts = split_top_level(group);
        if parts.is_empty() {
            continue;
        }
        let head = parts.remove(0).trim();
        if head.parse::<u32>().ok() != Some(kind) {
            continue;
        }
        for part in parts {
            out.extend(sheet_refs(part));
        }
    }
    out.sort_unstable();
    out.dedup();
    out
}

/// Inherited structure type of a resolved shape, if any.
pub fn structure_type(shape: &ResolvedShape) -> Option<&str> {
    match shape.cell("User.msvStructureType")? {
        Lookup::Found(cell) => cell.cell.value.as_deref(),
        Lookup::Deleted | Lookup::Absent => None,
    }
}

/// Whether a resolved shape is a container.
pub fn is_container(shape: &ResolvedShape) -> bool {
    structure_type(shape).is_some_and(|value| value == "Container")
}

/// Resolved container membership for already-resolved page shapes.
pub fn page_containers(shapes: &BTreeMap<u32, ResolvedShape>) -> PageContainers {
    let mut containers = BTreeMap::new();
    let mut member_to_containers: BTreeMap<u32, BTreeSet<u32>> = BTreeMap::new();
    for (id, shape) in shapes {
        if !is_container(shape) {
            continue;
        }
        containers.insert(
            *id,
            ContainerInfo {
                shape_id: *id,
                member_ids: Vec::new(),
                margin: user_number(shape, "msvSDContainerMargin"),
                resize: user_number(shape, "msvSDContainerResize").and_then(|value| {
                    if value.is_finite() && value.fract() == 0.0 {
                        Some(value as i64)
                    } else {
                        None
                    }
                }),
                locked: user_flag(shape, "msvSDContainerLocked"),
            },
        );
    }
    if containers.is_empty() {
        return PageContainers::default();
    }
    let mut claimed: BTreeMap<u32, BTreeSet<u32>> = BTreeMap::new();
    for (id, shape) in shapes {
        let relationships = match shape.cell("Relationships") {
            Some(Lookup::Found(cell)) => cell,
            _ => continue,
        };
        let Some(formula) = relationships.cell.formula.as_deref() else {
            continue;
        };
        for member in dependson_refs(formula, 1) {
            if shapes.contains_key(&member) {
                claimed.entry(*id).or_default().insert(member);
            }
        }
        for container in dependson_refs(formula, 4) {
            if containers.contains_key(&container) && shapes.contains_key(id) {
                claimed.entry(container).or_default().insert(*id);
            }
        }
    }
    for (id, info) in &mut containers {
        if let Some(members) = claimed.get(id) {
            info.member_ids = members.iter().copied().collect();
        }
        for member in &info.member_ids {
            member_to_containers.entry(*member).or_default().insert(*id);
        }
    }
    PageContainers {
        containers,
        member_to_containers: member_to_containers
            .into_iter()
            .map(|(id, owners)| (id, owners.into_iter().collect()))
            .collect(),
    }
}

impl<'a> Resolver<'a> {
    /// Resolves container membership for one page part.
    pub fn resolve_page_containers(&self, page_part: &str) -> Result<PageContainers, ResolveError> {
        Ok(page_containers(&self.resolve_page_shapes(page_part)?))
    }
}

fn user_number(shape: &ResolvedShape, row: &str) -> Option<f64> {
    let found = match shape.cell(&format!("User.{row}")) {
        Some(Lookup::Found(cell)) => cell,
        _ => return None,
    };
    if let Some(value) = found
        .cell
        .value
        .as_deref()
        .and_then(|value| value.parse::<f64>().ok())
        .filter(|value| value.is_finite())
    {
        return Some(value);
    }
    let formula = found.cell.formula.as_deref()?;
    evaluate_number(
        formula,
        FormulaLimits {
            max_depth: 64,
            max_nodes: 1_024,
            max_tokens: 1_024,
        },
        &mut |name| {
            let found = match shape.cell(name.trim()) {
                Some(Lookup::Found(cell)) => cell,
                _ => return None,
            };
            found
                .cell
                .formula
                .clone()
                .or_else(|| found.cell.value.clone())
        },
    )
}

fn user_flag(shape: &ResolvedShape, row: &str) -> bool {
    let found = match shape.cell(&format!("User.{row}")) {
        Some(Lookup::Found(cell)) => cell,
        _ => return false,
    };
    found
        .cell
        .value
        .as_deref()
        .is_some_and(|value| value == "1")
        || found
            .cell
            .formula
            .as_deref()
            .is_some_and(|formula| formula.trim().trim_start_matches('=') == "1")
}

fn find_call(text: &str, name: &str) -> Option<usize> {
    let upper = text.to_ascii_uppercase();
    let mut search = 0;
    while let Some(offset) = upper[search..].find(name) {
        let start = search + offset;
        let inside_ident = text[..start].chars().next_back().is_some_and(|previous| {
            previous.is_ascii_alphanumeric() || matches!(previous, '.' | '!' | '_')
        });
        if !inside_ident {
            let mut cursor = start + name.len();
            while text
                .as_bytes()
                .get(cursor)
                .is_some_and(|byte| byte.is_ascii_whitespace())
            {
                cursor += 1;
            }
            if text.as_bytes().get(cursor) == Some(&b'(') {
                return Some(start);
            }
        }
        search = start + name.len();
    }
    None
}

fn balanced_group(text: &str) -> Option<(&str, usize)> {
    let bytes = text.as_bytes();
    let mut cursor = bytes.iter().position(|byte| *byte == b'(')?;
    let open = cursor;
    let mut depth: i32 = 0;
    while cursor < bytes.len() {
        match bytes[cursor] {
            b'(' => depth += 1,
            b')' => {
                depth -= 1;
                if depth == 0 {
                    return Some((&text[open + 1..cursor], cursor + 1));
                }
            }
            b'"' => {
                cursor += 1;
                while cursor < bytes.len() && bytes[cursor] != b'"' {
                    if bytes[cursor] == b'"' && bytes.get(cursor + 1) == Some(&b'"') {
                        cursor += 1;
                    }
                    cursor += 1;
                }
            }
            _ => {}
        }
        cursor += 1;
    }
    None
}

fn split_top_level(text: &str) -> Vec<&str> {
    let mut parts = Vec::new();
    let bytes = text.as_bytes();
    let mut depth: i32 = 0;
    let mut start = 0;
    let mut cursor = 0;
    while cursor < bytes.len() {
        match bytes[cursor] {
            b'(' => depth += 1,
            b')' => depth = depth.saturating_sub(1),
            b',' if depth == 0 => {
                parts.push(&text[start..cursor]);
                start = cursor + 1;
            }
            b'"' => {
                cursor += 1;
                while cursor < bytes.len() && bytes[cursor] != b'"' {
                    cursor += 1;
                }
            }
            _ => {}
        }
        cursor += 1;
    }
    parts.push(&text[start..]);
    parts
}

fn sheet_refs(text: &str) -> Vec<u32> {
    let mut out = Vec::new();
    let upper = text.to_ascii_uppercase();
    let mut search = 0;
    while let Some(offset) = upper[search..].find("SHEET.") {
        let mut cursor = search + offset + "SHEET.".len();
        let start = cursor;
        while text
            .as_bytes()
            .get(cursor)
            .is_some_and(|byte| byte.is_ascii_digit())
        {
            cursor += 1;
        }
        if cursor > start
            && text.as_bytes().get(cursor) == Some(&b'!')
            && let Ok(id) = text[start..cursor].parse::<u32>()
        {
            out.push(id);
        }
        search = cursor.max(start + 1);
    }
    out
}

#[cfg(test)]
mod tests {
    use vsdx_parse::{
        Cell, Row, RowChild, Section, SectionChild, Shape, ShapeChild, ShapesChild, Sheet,
        SheetChild, VsdxPackage,
    };

    use super::*;

    fn value_cell(name: &str, value: &str) -> Cell {
        Cell {
            name: name.into(),
            formula: None,
            value: Some(value.into()),
            unit: None,
            del: false,
            other_attrs: Vec::new(),
        }
    }

    fn formula_cell(name: &str, formula: &str, value: &str) -> Cell {
        Cell {
            name: name.into(),
            formula: Some(formula.into()),
            value: Some(value.into()),
            unit: None,
            del: false,
            other_attrs: Vec::new(),
        }
    }

    fn user_row(name: &str, value: &str) -> Row {
        Row {
            index: None,
            name: Some(name.into()),
            local_name: None,
            row_type: None,
            del: false,
            children: vec![RowChild::Cell(value_cell("Value", value))],
            other_attrs: Vec::new(),
        }
    }

    fn master_shape(children: Vec<ShapeChild>) -> Shape {
        Shape {
            id: 1,
            name: None,
            name_u: None,
            shape_type: None,
            master: None,
            master_shape: None,
            line_style: None,
            fill_style: None,
            text_style: None,
            children,
            del: false,
            other_attrs: Vec::new(),
        }
    }

    fn page_shape(id: u32, master: Option<u32>, children: Vec<ShapeChild>) -> Shape {
        Shape {
            id,
            name: None,
            name_u: None,
            shape_type: None,
            master,
            master_shape: None,
            line_style: None,
            fill_style: None,
            text_style: None,
            children,
            del: false,
            other_attrs: Vec::new(),
        }
    }

    fn package(page_shapes: Vec<Shape>, master_root: Shape) -> VsdxPackage {
        let mut package: VsdxPackage = serde_json::from_value(serde_json::json!({
            "documentPartPath": "", "pagesPartPath": null, "mastersPartPath": null,
            "pagePartPaths": ["page"], "masterPartPaths": ["master"], "themePartPaths": [],
            "windowsPartPath": null, "relationships": {}, "documentSheet": null,
            "styleSheets": [], "colors": [], "faceNames": [], "pageSheets": {},
            "masterSheets": {}, "pagePartIds": {"page": 1}, "masterPartIds": {"master": 7},
            "pageContents": {}, "masterContents": {}
        }))
        .unwrap();
        package.page_contents.insert(
            "page".into(),
            Sheet {
                id: None,
                children: vec![SheetChild::Shapes(
                    page_shapes.into_iter().map(ShapesChild::Shape).collect(),
                )],
                other_attrs: Vec::new(),
            },
        );
        package.master_contents.insert(
            "master".into(),
            Sheet {
                id: None,
                children: vec![SheetChild::Shapes(vec![ShapesChild::Shape(master_root)])],
                other_attrs: Vec::new(),
            },
        );
        package
    }

    fn container_master() -> Shape {
        master_shape(vec![ShapeChild::Section(Section {
            name: "User".into(),
            index: None,
            del: false,
            children: vec![
                SectionChild::Row(user_row("msvStructureType", "Container")),
                SectionChild::Row(user_row("msvSDContainerMargin", "0.25")),
                SectionChild::Row(user_row("msvSDContainerResize", "2")),
            ],
            other_attrs: Vec::new(),
        })])
    }

    #[test]
    fn dependson_lists_sheet_members_per_kind() {
        let formula = "SUM(DEPENDSON(1,Sheet.5!SheetRef(),Sheet.4!SheetRef()),DEPENDSON(4,Sheet.1!SheetRef()))";
        assert_eq!(dependson_refs(formula, 1), vec![4, 5]);
        assert_eq!(dependson_refs(formula, 4), vec![1]);
        assert!(dependson_refs(formula, 2).is_empty());
    }

    #[test]
    fn dependson_ignores_nested_parens_and_stray_text() {
        assert_eq!(dependson_refs("SUM(DEPENDSON(4))", 4), Vec::<u32>::new());
        assert_eq!(
            dependson_refs("DEPENDSON(1, Sheet.10!SheetRef() , Sheet.3!Foo(1,2))", 1),
            vec![3, 10]
        );
        assert!(dependson_refs("DEPENDSONX(1,Sheet.2!SheetRef())", 1).is_empty());
    }

    #[test]
    fn dependson_requires_identifier_boundaries() {
        assert!(dependson_refs("XDEPENDSON(1,Sheet.2!SheetRef())", 1).is_empty());
        assert!(dependson_refs("MY_DEPENDSON(1,Sheet.2!SheetRef())", 1).is_empty());
        assert!(dependson_refs("DEPENDSONX(1,Sheet.2!SheetRef())", 1).is_empty());
        assert_eq!(
            dependson_refs("DEPENDSON(1,Sheet.2!SheetRef())", 1),
            vec![2]
        );
        assert_eq!(
            dependson_refs("SUM( DEPENDSON(1,Sheet.2!SheetRef()))", 1),
            vec![2]
        );
        assert_eq!(
            dependson_refs(
                "DEPENDSON(4,Sheet.1!SheetRef()),DEPENDSON(1,Sheet.2!SheetRef())",
                1
            ),
            vec![2]
        );
        assert_eq!(
            dependson_refs("0+DEPENDSON(1,Sheet.2!SheetRef())", 1),
            vec![2]
        );
    }

    #[test]
    fn spurious_dependson_identifiers_do_not_register_membership() {
        let package = package(
            vec![
                page_shape(
                    1,
                    Some(7),
                    vec![ShapeChild::Cell(formula_cell(
                        "Relationships",
                        "XDEPENDSON(1,Sheet.2!SheetRef())",
                        "0",
                    ))],
                ),
                page_shape(2, None, vec![]),
            ],
            container_master(),
        );
        let containers = Resolver::new(&package)
            .resolve_page_containers("page")
            .unwrap();
        assert_eq!(containers.containers.len(), 1);
        assert!(containers.members_of(1).is_empty());
        assert!(containers.containers_of(2).is_empty());
    }

    #[test]
    fn containers_resolve_members_through_master_inheritance() {
        let package = package(
            vec![
                page_shape(
                    1,
                    Some(7),
                    vec![ShapeChild::Cell(formula_cell(
                        "Relationships",
                        "SUM(DEPENDSON(1,Sheet.2!SheetRef(),Sheet.9!SheetRef()))",
                        "0",
                    ))],
                ),
                page_shape(
                    2,
                    None,
                    vec![ShapeChild::Cell(formula_cell(
                        "Relationships",
                        "SUM(DEPENDSON(4,Sheet.1!SheetRef()))",
                        "0",
                    ))],
                ),
                page_shape(3, None, vec![]),
            ],
            container_master(),
        );
        let containers = Resolver::new(&package)
            .resolve_page_containers("page")
            .unwrap();
        assert_eq!(containers.containers.len(), 1);
        let info = &containers.containers[&1];
        assert_eq!(info.member_ids, vec![2]);
        assert_eq!(info.margin, Some(0.25));
        assert_eq!(info.resize, Some(2));
        assert!(!info.locked);
        assert_eq!(containers.containers_of(2), &[1]);
        assert!(containers.containers_of(3).is_empty());
    }

    #[test]
    fn member_side_claims_join_container_side_lists() {
        let package = package(
            vec![
                page_shape(1, Some(7), vec![]),
                page_shape(
                    2,
                    None,
                    vec![ShapeChild::Cell(formula_cell(
                        "Relationships",
                        "SUM(DEPENDSON(4,Sheet.1!SheetRef()))",
                        "0",
                    ))],
                ),
            ],
            container_master(),
        );
        let containers = Resolver::new(&package)
            .resolve_page_containers("page")
            .unwrap();
        assert_eq!(containers.members_of(1), &[2]);
    }

    #[test]
    fn non_container_structures_stay_outside_membership() {
        let mut master = container_master();
        if let Some(ShapeChild::Section(section)) = master.children.first_mut()
            && let Some(SectionChild::Row(row)) = section.children.first_mut()
        {
            row.children = vec![RowChild::Cell(value_cell("Value", "List"))];
        }
        let package = package(vec![page_shape(1, Some(7), vec![])], master);
        let containers = Resolver::new(&package)
            .resolve_page_containers("page")
            .unwrap();
        assert!(containers.containers.is_empty());
    }

    #[test]
    fn nested_containers_report_transitive_members() {
        let package = package(
            vec![
                page_shape(
                    1,
                    Some(7),
                    vec![ShapeChild::Cell(formula_cell(
                        "Relationships",
                        "SUM(DEPENDSON(1,Sheet.2!SheetRef()))",
                        "0",
                    ))],
                ),
                page_shape(
                    2,
                    Some(7),
                    vec![ShapeChild::Cell(formula_cell(
                        "Relationships",
                        "SUM(DEPENDSON(4,Sheet.1!SheetRef()),DEPENDSON(1,Sheet.3!SheetRef()))",
                        "0",
                    ))],
                ),
                page_shape(3, None, vec![]),
            ],
            container_master(),
        );
        let containers = Resolver::new(&package)
            .resolve_page_containers("page")
            .unwrap();
        assert_eq!(containers.members_of(1), &[2]);
        assert_eq!(containers.members_of(2), &[3]);
        assert_eq!(containers.transitive_members_of(1), vec![2, 3]);
    }
}
