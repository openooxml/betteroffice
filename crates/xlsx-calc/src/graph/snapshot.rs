use std::collections::{BTreeMap, BTreeSet, LinkedList};
use std::mem::size_of;

use super::snapshot_growth::{MapGrowth, SetGrowth};
use super::*;
use crate::reference::table_rect_with_columns;

#[doc(hidden)]
pub struct SnapshotGraphBuilder {
    graph: DepGraph,
    names_growth: MapGrowth<String, SheetId>,
    indices_growth: MapGrowth<(Option<SheetId>, String), usize>,
    tables_growth: MapGrowth<String, Table>,
    deps_growth: MapGrowth<NodeKey, NodeEntry>,
    reverse_growth: MapGrowth<SheetId, Vec<(CellRange, NodeKey)>>,
    volatile_growth: SetGrowth<NodeKey>,
    spills_growth: MapGrowth<NodeKey, CellRange>,
    spill_sheets_growth: MapGrowth<SheetId, Vec<(NodeKey, CellRange)>>,
    asts_growth: MapGrowth<String, Arc<Expr>>,
    migrated_entries: usize,
    phase: u8,
    index: usize,
    column: usize,
    table: Option<Table>,
    last_cell: Option<(RowId, ColId)>,
    formula: Option<Formula>,
    columns: BTreeMap<(usize, String), u32>,
    table_indices: BTreeMap<String, usize>,
    retired: LinkedList<Table>,
    reverse: BTreeMap<SheetId, Buffered<(CellRange, NodeKey)>>,
    spills: BTreeMap<SheetId, Buffered<(NodeKey, CellRange)>>,
}

impl Default for SnapshotGraphBuilder {
    fn default() -> Self {
        Self::new()
    }
}

struct Formula {
    key: NodeKey,
    ast: Arc<Expr>,
    edges: Buffered<(SheetId, CellRange)>,
    references: LinkedList<(SheetId, CellRange)>,
    seen: BTreeSet<(SheetId, u32, u32, u32, u32)>,
    names: LinkedList<DefinedNameUse>,
    volatile_names: LinkedList<DefinedNameUse>,
    expanded: BTreeSet<(SheetId, String)>,
    volatile_expanded: BTreeSet<(SheetId, String)>,
    tables: LinkedList<(String, TableSpec)>,
    root_tables: LinkedList<(String, TableSpec)>,
    volatile: bool,
}

struct Buffered<T> {
    queued: LinkedList<T>,
    values: Option<Vec<T>>,
}

impl<T> Buffered<T> {
    fn new() -> Self {
        Self {
            queued: LinkedList::new(),
            values: None,
        }
    }

    fn push(&mut self, value: T) {
        self.queued.push_back(value);
    }

    fn advance(&mut self, max_bytes: usize) -> Result<(bool, usize), String> {
        if self.values.is_none() {
            let mut values = Vec::new();
            values
                .try_reserve_exact(self.queued.len())
                .map_err(|_| "cannot allocate graph edges".to_owned())?;
            self.values = Some(values);
            return Ok((self.queued.is_empty(), 0));
        }
        if !self.queued.is_empty() {
            let bytes = admit(size_of::<T>(), max_bytes)?;
            let value = self
                .queued
                .pop_front()
                .ok_or_else(|| "snapshot graph edge is missing".to_owned())?;
            self.values
                .as_mut()
                .ok_or_else(|| "snapshot graph edge storage is missing".to_owned())?
                .push(value);
            return Ok((self.queued.is_empty(), bytes));
        }
        Ok((true, 0))
    }

    fn finish(self) -> Result<Vec<T>, String> {
        if !self.queued.is_empty() {
            return Err("snapshot graph edges are incomplete".to_owned());
        }
        self.values
            .ok_or_else(|| "snapshot graph edge storage is missing".to_owned())
    }
}

impl Formula {
    fn expression(&mut self, graph: &DepGraph, owner: SheetId, expr: &Expr) {
        let mut references = Vec::new();
        graph.add_references(owner, expr, &mut references, &mut HashSet::new());
        self.references.extend(references);
        let mut names = Vec::new();
        push_defined_name_uses(owner, expr, &mut names);
        for name in names {
            self.names.push_front(name);
        }
        self.tables.extend(table_uses(expr));
    }
}

fn table_uses(expr: &Expr) -> LinkedList<(String, TableSpec)> {
    let mut uses = Vec::new();
    push_table_uses(expr, &mut uses);
    uses.into_iter()
        .map(|(name, spec)| (name.to_owned(), spec.clone()))
        .collect()
}

fn name_cost(name: &DefinedNameUse) -> usize {
    name.1
        .as_ref()
        .map_or(0, String::len)
        .saturating_add(name.2.len())
        .saturating_add(16)
}

pub(super) fn admit(bytes: usize, max_bytes: usize) -> Result<usize, String> {
    if bytes > max_bytes {
        Err("snapshot graph record exceeds advance byte budget".to_owned())
    } else {
        Ok(bytes)
    }
}

impl SnapshotGraphBuilder {
    #[doc(hidden)]
    pub fn new() -> Self {
        Self {
            graph: DepGraph::empty(),
            names_growth: MapGrowth::default(),
            indices_growth: MapGrowth::default(),
            tables_growth: MapGrowth::default(),
            deps_growth: MapGrowth::default(),
            reverse_growth: MapGrowth::default(),
            volatile_growth: SetGrowth::default(),
            spills_growth: MapGrowth::default(),
            spill_sheets_growth: MapGrowth::default(),
            asts_growth: MapGrowth::default(),
            migrated_entries: 0,
            phase: 0,
            index: 0,
            column: 0,
            table: None,
            last_cell: None,
            formula: None,
            columns: BTreeMap::new(),
            table_indices: BTreeMap::new(),
            retired: LinkedList::new(),
            reverse: BTreeMap::new(),
            spills: BTreeMap::new(),
        }
    }

    #[doc(hidden)]
    pub fn advance(
        &mut self,
        model: &Workbook,
        max_bytes: usize,
    ) -> Result<(bool, usize, usize), String> {
        if let Some(bytes) = self.advance_capacity(max_bytes)? {
            self.migrated_entries += 1;
            return Ok((false, 1, bytes));
        }
        if let Some(formula) = &mut self.formula {
            if let Some(reference) = formula.references.front() {
                let bytes = admit(size_of::<(SheetId, CellRange)>(), max_bytes)?;
                let (sheet, range) = *reference;
                formula.references.pop_front();
                let key = (
                    sheet,
                    range.start.row,
                    range.start.col,
                    range.end.row,
                    range.end.col,
                );
                if formula.seen.insert(key) {
                    formula.edges.push((sheet, range));
                    self.reverse
                        .entry(sheet)
                        .or_insert_with(Buffered::new)
                        .push((range, formula.key));
                }
                return Ok((false, 1, bytes));
            }
            if let Some((name, spec)) = formula.tables.front() {
                let bytes = admit(
                    name.len()
                        .saturating_add(spec.first_column.as_ref().map_or(0, String::len))
                        .saturating_add(spec.last_column.as_ref().map_or(0, String::len))
                        .saturating_add(64),
                    max_bytes,
                )?;
                let normalized = name.to_lowercase();
                if let Some(table) = self.graph.tables.get(&normalized)
                    && let Some(index) = self.table_indices.get(&normalized)
                    && let Ok(range) =
                        table_rect_with_columns(table, spec, Some(formula.key.cell()), |column| {
                            self.columns
                                .get(&(*index, column.to_ascii_lowercase()))
                                .copied()
                                .filter(|index| {
                                    table
                                        .range
                                        .start
                                        .col
                                        .checked_add(*index)
                                        .is_some_and(|col| col <= table.range.end.col)
                                })
                        })
                {
                    formula.references.push_back((table.sheet, range));
                }
                formula.tables.pop_front();
                return Ok((false, 1, bytes));
            }
            if let Some(name) = formula.names.front() {
                let resolved = self.graph.resolve_defined_name(name.0, &name.1, &name.2);
                let bytes = admit(
                    name_cost(name)
                        .saturating_add(resolved.map_or(0, |(_, value)| value.formula.len())),
                    max_bytes,
                )?;
                if let Some((sheet, defined)) = resolved {
                    let key = (sheet, name.2.to_ascii_lowercase());
                    if formula.expanded.insert(key)
                        && let Some(expression) = parse_cached(
                            &self.graph.asts,
                            defined
                                .formula
                                .strip_prefix('=')
                                .unwrap_or(&defined.formula),
                        )
                    {
                        let owner = defined.local_sheet.unwrap_or(sheet);
                        formula.names.pop_front();
                        formula.expression(&self.graph, owner, &expression);
                        return Ok((false, 1, bytes));
                    }
                }
                formula.names.pop_front();
                return Ok((false, 1, bytes));
            }
            if !formula.root_tables.is_empty() {
                std::mem::swap(&mut formula.tables, &mut formula.root_tables);
                return Ok((false, 0, 0));
            }
            if let Some(name) = formula.volatile_names.front() {
                let resolved = self.graph.resolve_defined_name(name.0, &name.1, &name.2);
                let bytes = admit(
                    name_cost(name)
                        .saturating_add(resolved.map_or(0, |(_, value)| value.formula.len())),
                    max_bytes,
                )?;
                let mut uses = Vec::new();
                if !formula.volatile
                    && let Some((sheet, defined)) = resolved
                    && formula
                        .volatile_expanded
                        .insert((sheet, name.2.to_ascii_lowercase()))
                    && let Some(expression) = parse_cached(
                        &self.graph.asts,
                        defined
                            .formula
                            .strip_prefix('=')
                            .unwrap_or(&defined.formula),
                    )
                {
                    formula.volatile = push_volatile_name_uses(
                        defined.local_sheet.unwrap_or(sheet),
                        &expression,
                        &mut uses,
                    );
                }
                formula.volatile_names.pop_front();
                for name in uses {
                    formula.volatile_names.push_front(name);
                }
                return Ok((false, 1, bytes));
            }
            if let Some((_, name)) = formula
                .expanded
                .last()
                .or_else(|| formula.volatile_expanded.last())
            {
                let bytes = admit(name.len().saturating_add(16), max_bytes)?;
                if formula.expanded.pop_last().is_none() {
                    formula.volatile_expanded.pop_last();
                }
                return Ok((false, 1, bytes));
            }
            if !formula.seen.is_empty() {
                let bytes = admit(24, max_bytes)?;
                formula.seen.pop_last();
                return Ok((false, 1, bytes));
            }
            let (ready, bytes) = formula.edges.advance(max_bytes)?;
            if !ready || bytes != 0 {
                return Ok((false, 1, bytes));
            }
            let Some(formula) = self.formula.take() else {
                return Err("snapshot graph formula is missing".to_owned());
            };
            if formula.volatile {
                self.graph.volatile.insert(formula.key);
            }
            let range = model
                .sheets
                .get(formula.key.sheet.0 as usize)
                .and_then(|sheet| sheet.array_formula(formula.key.cell()));
            self.graph.deps.insert(
                formula.key,
                NodeEntry {
                    ast: formula.ast,
                    edges: formula.edges.finish()?,
                },
            );
            if let Some(range) = range
                && range.start != range.end
            {
                self.graph.spills.insert(formula.key, range);
                self.spills
                    .entry(formula.key.sheet)
                    .or_insert_with(Buffered::new)
                    .push((formula.key, range));
            }
            return Ok((false, 1, 0));
        }
        loop {
            match self.phase {
                0 => {
                    let Some(sheet) = model.sheets.get(self.index) else {
                        self.graph
                            .defined_names
                            .try_reserve_exact(model.defined_names.len())
                            .map_err(|_| "cannot allocate graph names".to_owned())?;
                        self.phase = 1;
                        self.index = 0;
                        continue;
                    };
                    let bytes = admit(sheet.name.len().saturating_add(32), max_bytes)?;
                    self.graph
                        .names
                        .insert(sheet.name.to_lowercase(), SheetId(self.index as u32));
                    self.index += 1;
                    return Ok((false, 1, bytes));
                }
                1 => {
                    let Some(name) = model.defined_names.get(self.index) else {
                        self.phase = 2;
                        self.index = 0;
                        continue;
                    };
                    let bytes = admit(
                        name.name
                            .len()
                            .saturating_add(name.formula.len())
                            .saturating_add(32),
                        max_bytes,
                    )?;
                    self.graph
                        .defined_name_indices
                        .entry((name.local_sheet, name.name.to_ascii_lowercase()))
                        .or_insert(self.index);
                    self.graph.defined_names.push(name.clone());
                    self.index += 1;
                    return Ok((false, 1, bytes));
                }
                2 => {
                    let Some(table) = model.tables.get(self.index) else {
                        self.phase = 3;
                        self.index = 0;
                        continue;
                    };
                    if self.table.is_none() {
                        let bytes = admit(
                            table.name.len().saturating_add(size_of::<Table>()),
                            max_bytes,
                        )?;
                        let mut columns = Vec::new();
                        columns
                            .try_reserve_exact(table.columns.len())
                            .map_err(|_| "cannot allocate graph columns".to_owned())?;
                        self.table = Some(Table {
                            name: table.name.clone(),
                            sheet: table.sheet,
                            range: table.range,
                            header_rows: table.header_rows,
                            totals_rows: table.totals_rows,
                            columns,
                        });
                        self.column = 0;
                        return Ok((false, 1, bytes));
                    }
                    if let Some(column) = table.columns.get(self.column) {
                        let bytes = admit(column.len().saturating_add(16), max_bytes)?;
                        self.columns
                            .entry((self.index, column.to_ascii_lowercase()))
                            .or_insert(self.column as u32);
                        self.table
                            .as_mut()
                            .ok_or_else(|| "snapshot graph table is missing".to_owned())?
                            .columns
                            .push(column.clone());
                        self.column += 1;
                        return Ok((false, 1, bytes));
                    }
                    let bytes = admit(table.name.len().saturating_add(16), max_bytes)?;
                    let table = self
                        .table
                        .take()
                        .ok_or_else(|| "snapshot graph table is missing".to_owned())?;
                    let name = table.name.to_lowercase();
                    self.table_indices.insert(name.clone(), self.index);
                    if let Some(old) = self.graph.tables.insert(name, table) {
                        self.retired.push_back(old);
                    }
                    self.index += 1;
                    return Ok((false, 1, bytes));
                }
                3 => {
                    let Some(sheet) = model.sheets.get(self.index) else {
                        self.phase = 4;
                        continue;
                    };
                    let Some((at, cell)) = super::builder::next_cell(sheet, self.last_cell) else {
                        let bytes = admit(1, max_bytes)?;
                        self.index += 1;
                        self.last_cell = None;
                        return Ok((false, 1, bytes));
                    };
                    let bytes = admit(cell.formula.as_ref().map_or(1, String::len), max_bytes)?;
                    self.last_cell = Some((at.row, at.col));
                    if let Some(source) = &cell.formula
                        && let Some(ast) = parse_cached(&self.graph.asts, source)
                    {
                        let key = NodeKey::new(SheetId(self.index as u32), at);
                        let mut formula = Formula {
                            key,
                            ast: Arc::clone(&ast),
                            edges: Buffered::new(),
                            references: LinkedList::new(),
                            seen: BTreeSet::new(),
                            names: LinkedList::new(),
                            volatile_names: LinkedList::new(),
                            expanded: BTreeSet::new(),
                            volatile_expanded: BTreeSet::new(),
                            tables: LinkedList::new(),
                            root_tables: LinkedList::new(),
                            volatile: false,
                        };
                        formula.expression(&self.graph, key.sheet, &ast);
                        std::mem::swap(&mut formula.tables, &mut formula.root_tables);
                        let mut uses = Vec::new();
                        formula.volatile = push_volatile_name_uses(key.sheet, &ast, &mut uses);
                        for name in uses {
                            formula.volatile_names.push_front(name);
                        }
                        self.formula = Some(formula);
                    }
                    return Ok((false, 1, bytes));
                }
                4 => {
                    if let Some((&sheet, _)) = self.reverse.first_key_value() {
                        let edges = self
                            .reverse
                            .get_mut(&sheet)
                            .ok_or_else(|| "snapshot reverse edges are missing".to_owned())?;
                        let (ready, bytes) = edges.advance(max_bytes)?;
                        if ready {
                            let (_, edges) = self
                                .reverse
                                .pop_first()
                                .ok_or_else(|| "snapshot reverse edges are missing".to_owned())?;
                            self.graph.by_sheet.insert(sheet, edges.finish()?);
                        }
                        return Ok((false, 1, bytes));
                    }
                    self.phase = 5;
                }
                5 => {
                    if let Some((&sheet, _)) = self.spills.first_key_value() {
                        let edges = self
                            .spills
                            .get_mut(&sheet)
                            .ok_or_else(|| "snapshot spill edges are missing".to_owned())?;
                        let (ready, bytes) = edges.advance(max_bytes)?;
                        if ready {
                            let (_, edges) = self
                                .spills
                                .pop_first()
                                .ok_or_else(|| "snapshot spill edges are missing".to_owned())?;
                            self.graph.spills_by_sheet.insert(sheet, edges.finish()?);
                        }
                        return Ok((false, 1, bytes));
                    }
                    self.phase = 6;
                }
                6 => {
                    if let Some(((_, name), _)) = self.columns.last_key_value() {
                        let bytes = admit(name.len().saturating_add(16), max_bytes)?;
                        self.columns.pop_last();
                        return Ok((false, 1, bytes));
                    }
                    if let Some((name, _)) = self.table_indices.last_key_value() {
                        let bytes = admit(name.len().saturating_add(16), max_bytes)?;
                        self.table_indices.pop_last();
                        return Ok((false, 1, bytes));
                    }
                    self.phase = 7;
                }
                7 => {
                    if let Some(table) = self.retired.front_mut() {
                        if let Some(column) = table.columns.last() {
                            let bytes = admit(column.len().saturating_add(16), max_bytes)?;
                            table.columns.pop();
                            return Ok((false, 1, bytes));
                        }
                        let bytes = admit(table.name.len().saturating_add(16), max_bytes)?;
                        self.retired.pop_front();
                        return Ok((false, 1, bytes));
                    }
                    self.phase = 8;
                }
                _ => return Ok((true, 0, 0)),
            }
        }
    }

    fn advance_capacity(&mut self, max_bytes: usize) -> Result<Option<usize>, String> {
        if let Some(bytes) =
            self.names_growth
                .ensure(&mut self.graph.names, String::len, max_bytes)?
        {
            return Ok(Some(bytes));
        }
        if let Some(bytes) = self.indices_growth.ensure(
            &mut self.graph.defined_name_indices,
            |key| key.1.len(),
            max_bytes,
        )? {
            return Ok(Some(bytes));
        }
        if let Some(bytes) =
            self.tables_growth
                .ensure(&mut self.graph.tables, String::len, max_bytes)?
        {
            return Ok(Some(bytes));
        }
        if let Some(bytes) = self
            .deps_growth
            .ensure(&mut self.graph.deps, |_| 0, max_bytes)?
        {
            return Ok(Some(bytes));
        }
        if let Some(bytes) =
            self.reverse_growth
                .ensure(&mut self.graph.by_sheet, |_| 0, max_bytes)?
        {
            return Ok(Some(bytes));
        }
        if let Some(bytes) = self
            .volatile_growth
            .ensure(&mut self.graph.volatile, max_bytes)?
        {
            return Ok(Some(bytes));
        }
        if let Some(bytes) = self
            .spills_growth
            .ensure(&mut self.graph.spills, |_| 0, max_bytes)?
        {
            return Ok(Some(bytes));
        }
        if let Some(bytes) =
            self.spill_sheets_growth
                .ensure(&mut self.graph.spills_by_sheet, |_| 0, max_bytes)?
        {
            return Ok(Some(bytes));
        }
        self.asts_growth.ensure(
            &mut self.graph.asts.lock().expect("parse cache poisoned"),
            String::len,
            max_bytes,
        )
    }

    #[doc(hidden)]
    pub fn migrated_entries(&self) -> usize {
        self.migrated_entries
    }

    #[doc(hidden)]
    pub fn finish(self) -> Option<DepGraph> {
        (self.phase == 8).then_some(self.graph)
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use xlsx_model::{Cell, CellValue, Sheet};

    #[test]
    fn fifty_thousand_formulas_migrate_graph_storage_within_each_step() {
        let mut sheet = Sheet::new("Data");
        for row in 0..50_000 {
            sheet.set_cell(
                CellRef::new(row, 0),
                Cell {
                    value: CellValue::Number { value: 1.0 },
                    formula: Some(format!("B{}+RAND()", row + 1)),
                    style: None,
                },
            );
        }
        let model = Workbook {
            sheets: vec![sheet],
            ..Workbook::default()
        };
        let mut builder = SnapshotGraphBuilder::new();
        loop {
            let before = builder.migrated_entries;
            let (ready, records, bytes) = builder.advance(&model, 16_384).unwrap();
            let moved = builder.migrated_entries - before;
            assert!(records <= 1);
            assert!(bytes <= 16_384);
            assert!(moved <= records);
            if ready {
                break;
            }
        }
        assert!(builder.migrated_entries >= 50_000);
        let graph = builder.finish().unwrap();
        assert_eq!(graph.deps.len(), 50_000);
        assert_eq!(graph.volatile.len(), 50_000);
        assert_eq!(graph.by_sheet[&SheetId(0)].len(), 50_000);
    }
}
