//! dependency graph: which formula cells read which cells, answering "when this
//! cell changes, which formulas must re-evaluate?".

use std::collections::{HashMap, HashSet};

use xlsx_model::{CellRange, CellRef, ColId, DefinedName, RowId, SheetId, Workbook};

use crate::deps::references;
use crate::parser::{Expr, parse_formula};

/// a formula cell, normalized so `$`-anchoring never splits a node.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, PartialOrd, Ord)]
struct NodeKey {
    sheet: SheetId,
    row: RowId,
    col: ColId,
}

impl NodeKey {
    fn new(sheet: SheetId, cell: CellRef) -> Self {
        Self {
            sheet,
            row: cell.row,
            col: cell.col,
        }
    }

    fn cell(&self) -> CellRef {
        CellRef::new(self.row, self.col)
    }
}

/// case-insensitive names of functions whose value can change with no input
/// edit; cells calling them re-evaluate on every recalc.
const VOLATILE_FNS: [&str; 4] = ["TODAY", "NOW", "RAND", "RANDBETWEEN"];

pub struct DepGraph {
    /// sheet name -> id, snapshot at build time.
    names: HashMap<String, SheetId>,
    defined_names: Vec<DefinedName>,
    /// forward edges: formula node -> the cells/ranges it reads (sheets resolved).
    deps: HashMap<NodeKey, Vec<(SheetId, CellRange)>>,
    /// reverse index by sheet: `(range, dependent)` pairs read into that sheet.
    by_sheet: HashMap<SheetId, Vec<(CellRange, NodeKey)>>,
    /// formula cells that must re-evaluate every recalc regardless of edits.
    volatile: HashSet<NodeKey>,
}

impl DepGraph {
    /// build the whole graph from a workbook's stored formulas; unparseable
    /// formulas are skipped.
    pub fn build(wb: &Workbook) -> Self {
        let names = wb
            .sheets
            .iter()
            .enumerate()
            .map(|(i, s)| (s.name.to_lowercase(), SheetId(i as u32)))
            .collect();
        let mut g = DepGraph {
            names,
            defined_names: wb.defined_names.clone(),
            deps: HashMap::new(),
            by_sheet: HashMap::new(),
            volatile: HashSet::new(),
        };
        for (i, sheet) in wb.sheets.iter().enumerate() {
            let sid = SheetId(i as u32);
            for (cell, c) in sheet.iter_cells() {
                if let Some(src) = &c.formula {
                    g.install(NodeKey::new(sid, cell), src);
                }
            }
        }
        g
    }

    /// re-derive one node's edges after its formula changed, without touching
    /// the rest of the graph. `None` clears the node (cell no longer a formula).
    pub fn set_formula(&mut self, sheet: SheetId, cell: CellRef, formula: Option<&str>) {
        let key = NodeKey::new(sheet, cell);
        self.uninstall(key);
        if let Some(src) = formula {
            self.install(key, src);
        }
    }

    /// coarse invalidation for a sheet insert: ids shift, so rebuild wholesale.
    pub fn add_sheet(&mut self, wb: &Workbook) {
        *self = Self::build(wb);
    }

    /// coarse invalidation for a sheet removal: ids shift, so rebuild wholesale.
    pub fn remove_sheet(&mut self, wb: &Workbook) {
        *self = Self::build(wb);
    }

    /// formula cells that directly read `cell` on `sheet`; may contain
    /// duplicates, callers dedup.
    pub fn dependents_of(
        &self,
        sheet: SheetId,
        cell: CellRef,
    ) -> impl Iterator<Item = (SheetId, CellRef)> + '_ {
        let target = CellRef::new(cell.row, cell.col);
        self.by_sheet
            .get(&sheet)
            .into_iter()
            .flatten()
            .filter(move |(range, _)| range.contains(target))
            .map(|(_, node)| (node.sheet, node.cell()))
    }

    /// whether a cell is a (parseable) formula node.
    pub fn is_formula(&self, sheet: SheetId, cell: CellRef) -> bool {
        self.deps.contains_key(&NodeKey::new(sheet, cell))
    }

    /// every formula cell, in unspecified order.
    pub fn formula_cells(&self) -> impl Iterator<Item = (SheetId, CellRef)> + '_ {
        self.deps.keys().map(|k| (k.sheet, k.cell()))
    }

    /// every volatile formula cell, in unspecified order.
    pub fn volatile_cells(&self) -> impl Iterator<Item = (SheetId, CellRef)> + '_ {
        self.volatile.iter().map(|k| (k.sheet, k.cell()))
    }

    /// parse a formula and register its edges + volatility. no-op on parse error.
    fn install(&mut self, key: NodeKey, src: &str) {
        let Ok(expr) = parse_formula(src) else {
            return;
        };
        let edges = self.resolve_edges(key.sheet, &expr);
        for (sid, range) in &edges {
            self.by_sheet.entry(*sid).or_default().push((*range, key));
        }
        if self.is_volatile(key.sheet, &expr, &mut HashSet::new()) {
            self.volatile.insert(key);
        }
        self.deps.insert(key, edges);
    }

    /// drop a node's edges from every index it appears in.
    fn uninstall(&mut self, key: NodeKey) {
        if let Some(edges) = self.deps.remove(&key) {
            for (sid, _) in &edges {
                if let Some(list) = self.by_sheet.get_mut(sid) {
                    list.retain(|(_, node)| *node != key);
                }
            }
        }
        self.volatile.remove(&key);
    }

    /// resolve refs to concrete sheet ids; unqualified refs bind to the owning
    /// sheet, unknown sheet names drop the edge.
    fn resolve_edges(&self, owner: SheetId, expr: &Expr) -> Vec<(SheetId, CellRange)> {
        let mut out = Vec::new();
        let mut seen = HashSet::new();
        self.add_references(owner, expr, &mut out, &mut seen);
        self.add_defined_name_references(owner, expr, &mut out, &mut seen, &mut HashSet::new());
        out
    }

    fn add_references(
        &self,
        owner: SheetId,
        expr: &Expr,
        out: &mut Vec<(SheetId, CellRange)>,
        seen: &mut HashSet<(SheetId, u32, u32, u32, u32)>,
    ) {
        for (sheet, range) in references(expr) {
            let sid = match sheet {
                None => owner,
                Some(name) => match self.names.get(&name.to_lowercase()) {
                    Some(id) => *id,
                    None => continue,
                },
            };
            let key = (
                sid,
                range.start.row,
                range.start.col,
                range.end.row,
                range.end.col,
            );
            if seen.insert(key) {
                out.push((sid, range));
            }
        }
    }

    fn add_defined_name_references(
        &self,
        owner: SheetId,
        expr: &Expr,
        out: &mut Vec<(SheetId, CellRange)>,
        seen: &mut HashSet<(SheetId, u32, u32, u32, u32)>,
        name_stack: &mut HashSet<(SheetId, String)>,
    ) {
        match expr {
            Expr::Name { scope, name } => {
                let Some((lookup_sheet, defined)) = self.resolve_defined_name(owner, scope, name)
                else {
                    return;
                };
                let key = (lookup_sheet, name.to_lowercase());
                if !name_stack.insert(key.clone()) {
                    return;
                }
                if let Ok(expression) = parse_formula(
                    defined
                        .formula
                        .strip_prefix('=')
                        .unwrap_or(&defined.formula),
                ) {
                    let definition_sheet = defined.local_sheet.unwrap_or(lookup_sheet);
                    self.add_references(definition_sheet, &expression, out, seen);
                    self.add_defined_name_references(
                        definition_sheet,
                        &expression,
                        out,
                        seen,
                        name_stack,
                    );
                }
                name_stack.remove(&key);
            }
            Expr::Unary { expr, .. } | Expr::Percent(expr) => {
                self.add_defined_name_references(owner, expr, out, seen, name_stack);
            }
            Expr::Binary { lhs, rhs, .. } => {
                self.add_defined_name_references(owner, lhs, out, seen, name_stack);
                self.add_defined_name_references(owner, rhs, out, seen, name_stack);
            }
            Expr::FuncCall { args, .. } => {
                for argument in args {
                    self.add_defined_name_references(owner, argument, out, seen, name_stack);
                }
            }
            Expr::Number(_)
            | Expr::Text(_)
            | Expr::Bool(_)
            | Expr::Error(_)
            | Expr::Ref { .. }
            | Expr::Range { .. } => {}
        }
    }

    fn resolve_defined_name<'a>(
        &'a self,
        owner: SheetId,
        scope: &Option<String>,
        name: &str,
    ) -> Option<(SheetId, &'a DefinedName)> {
        let lookup_sheet = match scope {
            Some(scope) => *self.names.get(&scope.to_lowercase())?,
            None => owner,
        };
        let defined = self
            .defined_names
            .iter()
            .find(|defined| {
                defined.local_sheet == Some(lookup_sheet) && defined.name.eq_ignore_ascii_case(name)
            })
            .or_else(|| {
                self.defined_names.iter().find(|defined| {
                    defined.local_sheet.is_none() && defined.name.eq_ignore_ascii_case(name)
                })
            })?;
        Some((lookup_sheet, defined))
    }

    fn is_volatile(
        &self,
        owner: SheetId,
        expr: &Expr,
        name_stack: &mut HashSet<(SheetId, String)>,
    ) -> bool {
        match expr {
            Expr::FuncCall { name, args } => {
                let upper = name.to_ascii_uppercase();
                VOLATILE_FNS.contains(&upper.as_str())
                    || args
                        .iter()
                        .any(|argument| self.is_volatile(owner, argument, name_stack))
            }
            Expr::Name { scope, name } => {
                let Some((lookup_sheet, defined)) = self.resolve_defined_name(owner, scope, name)
                else {
                    return false;
                };
                let key = (lookup_sheet, name.to_lowercase());
                if !name_stack.insert(key.clone()) {
                    return false;
                }
                let volatile = parse_formula(
                    defined
                        .formula
                        .strip_prefix('=')
                        .unwrap_or(&defined.formula),
                )
                .is_ok_and(|expression| {
                    self.is_volatile(
                        defined.local_sheet.unwrap_or(lookup_sheet),
                        &expression,
                        name_stack,
                    )
                });
                name_stack.remove(&key);
                volatile
            }
            Expr::Unary { expr, .. } | Expr::Percent(expr) => {
                self.is_volatile(owner, expr, name_stack)
            }
            Expr::Binary { lhs, rhs, .. } => {
                self.is_volatile(owner, lhs, name_stack) || self.is_volatile(owner, rhs, name_stack)
            }
            _ => false,
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use xlsx_model::{Cell, CellValue, Sheet};

    fn a1(s: &str) -> CellRef {
        CellRef::parse_a1(s).unwrap()
    }

    fn formula_cell(f: &str) -> Cell {
        Cell {
            value: CellValue::Empty,
            formula: Some(f.to_string()),
            style: None,
        }
    }

    /// workbook with two sheets; caller populates cells.
    fn wb2() -> Workbook {
        let mut wb = Workbook::default();
        wb.sheets.push(Sheet::new("Sheet1"));
        wb.sheets.push(Sheet::new("Data"));
        wb
    }

    fn deps_a1(g: &DepGraph, sheet: &str, cell: &str, wb: &Workbook) -> Vec<String> {
        let sid = wb.sheet_by_name(sheet).unwrap().0;
        let mut out: Vec<String> = g
            .dependents_of(sid, a1(cell))
            .map(|(s, c)| format!("{}!{}", wb.sheet(s).unwrap().name, c.to_a1()))
            .collect();
        out.sort();
        out
    }

    #[test]
    fn single_cell_dependents() {
        let mut wb = wb2();
        wb.sheet_mut(SheetId(0))
            .unwrap()
            .set_cell(a1("B1"), formula_cell("A1+1"));
        let g = DepGraph::build(&wb);
        assert_eq!(deps_a1(&g, "Sheet1", "A1", &wb), vec!["Sheet1!B1"]);
        assert!(deps_a1(&g, "Sheet1", "Z9", &wb).is_empty());
    }

    #[test]
    fn range_dependents_without_materializing_edges() {
        let mut wb = wb2();
        wb.sheet_mut(SheetId(0))
            .unwrap()
            .set_cell(a1("C1"), formula_cell("SUM(A1:A1000)"));
        let g = DepGraph::build(&wb);
        assert_eq!(deps_a1(&g, "Sheet1", "A5", &wb), vec!["Sheet1!C1"]);
        assert_eq!(deps_a1(&g, "Sheet1", "A1000", &wb), vec!["Sheet1!C1"]);
        assert!(deps_a1(&g, "Sheet1", "A1001", &wb).is_empty());
        assert_eq!(
            g.deps
                .get(&NodeKey::new(SheetId(0), a1("C1")))
                .unwrap()
                .len(),
            1
        );
    }

    #[test]
    fn cross_sheet_edges_resolve_and_unknown_sheets_drop() {
        let mut wb = wb2();
        wb.sheet_mut(SheetId(1))
            .unwrap()
            .set_cell(a1("A1"), formula_cell("Sheet1!A1 + Ghost!B2"));
        let g = DepGraph::build(&wb);
        assert_eq!(deps_a1(&g, "Sheet1", "A1", &wb), vec!["Data!A1"]);
        assert_eq!(
            g.deps
                .get(&NodeKey::new(SheetId(1), a1("A1")))
                .unwrap()
                .len(),
            1
        );
    }

    #[test]
    fn set_formula_swaps_edges() {
        let mut wb = wb2();
        wb.sheet_mut(SheetId(0))
            .unwrap()
            .set_cell(a1("C1"), formula_cell("A1"));
        let mut g = DepGraph::build(&wb);
        assert_eq!(deps_a1(&g, "Sheet1", "A1", &wb), vec!["Sheet1!C1"]);

        g.set_formula(SheetId(0), a1("C1"), Some("B1"));
        assert!(deps_a1(&g, "Sheet1", "A1", &wb).is_empty());
        assert_eq!(deps_a1(&g, "Sheet1", "B1", &wb), vec!["Sheet1!C1"]);

        g.set_formula(SheetId(0), a1("C1"), None);
        assert!(deps_a1(&g, "Sheet1", "B1", &wb).is_empty());
        assert!(!g.is_formula(SheetId(0), a1("C1")));
    }

    #[test]
    fn volatile_detection() {
        let mut wb = wb2();
        let s = wb.sheet_mut(SheetId(0)).unwrap();
        s.set_cell(a1("A1"), formula_cell("NOW()"));
        s.set_cell(a1("A2"), formula_cell("A1 + TODAY()"));
        s.set_cell(a1("A3"), formula_cell("A1 + 1"));
        let g = DepGraph::build(&wb);
        let mut vol: Vec<String> = g.volatile_cells().map(|(_, c)| c.to_a1()).collect();
        vol.sort();
        assert_eq!(vol, vec!["A1", "A2"]);
    }

    #[test]
    fn anchored_refs_normalize_to_same_node() {
        let mut wb = wb2();
        wb.sheet_mut(SheetId(0))
            .unwrap()
            .set_cell(a1("B1"), formula_cell("$A$1 + 1"));
        let g = DepGraph::build(&wb);
        assert_eq!(deps_a1(&g, "Sheet1", "A1", &wb), vec!["Sheet1!B1"]);
    }

    #[test]
    fn resolved_sheet_aliases_deduplicate() {
        let mut wb = wb2();
        wb.sheet_mut(SheetId(0))
            .unwrap()
            .set_cell(a1("C1"), formula_cell("A1+$A$1+sheet1!A1"));
        let graph = DepGraph::build(&wb);
        assert_eq!(
            graph
                .deps
                .get(&NodeKey::new(SheetId(0), a1("C1")))
                .unwrap()
                .len(),
            1
        );
    }

    #[test]
    fn defined_name_edges_and_volatility_are_expanded() {
        let mut wb = wb2();
        wb.defined_names.extend([
            DefinedName {
                name: "Inputs".into(),
                formula: "Data!A1:A3".into(),
                local_sheet: None,
                hidden: false,
            },
            DefinedName {
                name: "Clock".into(),
                formula: "NOW()".into(),
                local_sheet: None,
                hidden: false,
            },
        ]);
        let sheet = wb.sheet_mut(SheetId(0)).unwrap();
        sheet.set_cell(a1("B1"), formula_cell("SUM(Inputs)"));
        sheet.set_cell(a1("B2"), formula_cell("Clock"));

        let graph = DepGraph::build(&wb);
        assert_eq!(deps_a1(&graph, "Data", "A2", &wb), vec!["Sheet1!B1"]);
        assert_eq!(
            graph.volatile_cells().collect::<Vec<_>>(),
            vec![(SheetId(0), a1("B2"))]
        );
    }
}
