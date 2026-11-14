use super::*;

use crate::parser::{Expr, parse_formula};

use std::collections::BTreeMap;
use std::panic::{AssertUnwindSafe, catch_unwind};
use std::sync::Arc;

use xlsx_model::{Cell, Sheet};

const NOW_SERIAL: f64 = 45_000.25;
const RAND_SEED: u32 = 17;

#[derive(Debug, Clone, Copy, PartialEq)]
enum Edit {
    Unchanged,
    Value,
    Formula,
    RenameSheet,
    AddSheet,
    DeleteSheet,
    AddName,
    ChangeName,
    DeleteName,
    InsertRow,
    DeleteRow,
    InsertColumn,
    DeleteColumn,
}

const EDITS: [Edit; 13] = [
    Edit::Unchanged,
    Edit::Value,
    Edit::Formula,
    Edit::RenameSheet,
    Edit::AddSheet,
    Edit::DeleteSheet,
    Edit::AddName,
    Edit::ChangeName,
    Edit::DeleteName,
    Edit::InsertRow,
    Edit::DeleteRow,
    Edit::InsertColumn,
    Edit::DeleteColumn,
];

#[derive(Debug, Clone, Copy, PartialEq)]
enum Shift {
    Insert,
    Delete,
}

/// a slice of formula text; references name a sheet slot, so renames,
/// deletions and row or column shifts rewrite them as the workbook would.
#[derive(Debug, Clone)]
enum Piece {
    Text(String),
    Ref {
        sheet: Option<usize>,
        at: (u32, u32),
        abs: bool,
    },
    Range {
        sheet: Option<usize>,
        start: (u32, u32),
        end: (u32, u32),
        abs: bool,
    },
    Broken,
}

#[derive(Debug, Clone)]
enum Content {
    Number(f64),
    Formula(Vec<Piece>),
}

#[derive(Debug, Clone)]
struct Model {
    /// every sheet ever created: its name and whether it still exists.
    slots: Vec<(String, bool)>,
    /// the live slots in workbook order.
    order: Vec<usize>,
    /// per slot: address -> (cell identity, content).
    cells: Vec<BTreeMap<(u32, u32), (usize, Content)>>,
    names: Vec<(String, Vec<Piece>)>,
    next_id: usize,
}

fn text(source: &str) -> Piece {
    Piece::Text(source.into())
}

fn relative(sheet: Option<usize>, row: u32, col: u32) -> Piece {
    Piece::Ref {
        sheet,
        at: (row, col),
        abs: false,
    }
}

fn absolute(sheet: Option<usize>, row: u32, col: u32) -> Piece {
    Piece::Ref {
        sheet,
        at: (row, col),
        abs: true,
    }
}

fn formula(pieces: Vec<Piece>) -> Content {
    Content::Formula(pieces)
}

fn address((row, col): (u32, u32), abs: bool) -> String {
    CellRef {
        row,
        col,
        abs_row: abs,
        abs_col: abs,
    }
    .to_a1()
}

fn shift_point(point: (u32, u32), axis: usize, index: u32, shift: Shift) -> Option<(u32, u32)> {
    let mut point = [point.0, point.1];
    match shift {
        Shift::Insert => {
            if point[axis] >= index {
                point[axis] += 1;
            }
        }
        Shift::Delete => {
            if point[axis] == index {
                return None;
            }
            if point[axis] > index {
                point[axis] -= 1;
            }
        }
    }
    Some((point[0], point[1]))
}

fn shift_range(
    start: (u32, u32),
    end: (u32, u32),
    axis: usize,
    index: u32,
    shift: Shift,
) -> Option<((u32, u32), (u32, u32))> {
    let mut start = [start.0, start.1];
    let mut end = [end.0, end.1];
    match shift {
        Shift::Insert => {
            if start[axis] >= index {
                start[axis] += 1;
            }
            if end[axis] >= index {
                end[axis] += 1;
            }
        }
        Shift::Delete => {
            if start[axis] == index && end[axis] == index {
                return None;
            }
            if start[axis] > index {
                start[axis] -= 1;
            }
            if end[axis] >= index {
                end[axis] -= 1;
            }
        }
    }
    Some(((start[0], start[1]), (end[0], end[1])))
}

fn shift_pieces(
    pieces: &mut [Piece],
    owner: Option<usize>,
    slot: usize,
    axis: usize,
    index: u32,
    shift: Shift,
) {
    for piece in pieces.iter_mut() {
        let moved = match &*piece {
            Piece::Ref { sheet, at, abs } if sheet.or(owner) == Some(slot) => {
                let (sheet, abs) = (*sheet, *abs);
                let at = shift_point(*at, axis, index, shift);
                at.map(|at| Piece::Ref { sheet, at, abs })
            }
            Piece::Range {
                sheet,
                start,
                end,
                abs,
            } if sheet.or(owner) == Some(slot) => {
                let (sheet, abs) = (*sheet, *abs);
                let moved = shift_range(*start, *end, axis, index, shift);
                moved.map(|(start, end)| Piece::Range {
                    sheet,
                    start,
                    end,
                    abs,
                })
            }
            _ => continue,
        };
        *piece = moved.unwrap_or(Piece::Broken);
    }
}

impl Model {
    fn put(&mut self, slot: usize, at: (u32, u32), content: Content) {
        let id = match self.cells[slot].get(&at) {
            Some((id, _)) => *id,
            None => {
                self.next_id += 1;
                self.next_id
            }
        };
        self.cells[slot].insert(at, (id, content));
    }

    fn prefix(&self, sheet: Option<usize>) -> Option<String> {
        let Some(slot) = sheet else {
            return Some(String::new());
        };
        let (name, alive) = &self.slots[slot];
        if !*alive {
            None
        } else if name.contains(' ') {
            Some(format!("'{name}'!"))
        } else {
            Some(format!("{name}!"))
        }
    }

    fn text(&self, pieces: &[Piece]) -> String {
        let mut out = String::new();
        for piece in pieces {
            match piece {
                Piece::Text(source) => out.push_str(source),
                Piece::Broken => out.push_str("#REF!"),
                Piece::Ref { sheet, at, abs } => match self.prefix(*sheet) {
                    Some(prefix) => out.push_str(&format!("{prefix}{}", address(*at, *abs))),
                    None => out.push_str("#REF!"),
                },
                Piece::Range {
                    sheet,
                    start,
                    end,
                    abs,
                } => match self.prefix(*sheet) {
                    Some(prefix) => {
                        let (start, end) = (address(*start, *abs), address(*end, *abs));
                        out.push_str(&format!("{prefix}{start}:{end}"));
                    }
                    None => out.push_str("#REF!"),
                },
            }
        }
        out
    }

    /// the workbook this model describes; formula cells carry the value
    /// `values` holds for their identity, as an edited workbook keeps them.
    fn render(&self, values: &HashMap<usize, CellValue>) -> Workbook {
        let mut wb = Workbook::default();
        for &slot in &self.order {
            let mut sheet = Sheet::new(self.slots[slot].0.clone());
            for (&(row, col), (id, content)) in &self.cells[slot] {
                let cell = match content {
                    Content::Number(value) => Cell {
                        value: CellValue::Number { value: *value },
                        ..Cell::default()
                    },
                    Content::Formula(pieces) => Cell {
                        value: values.get(id).cloned().unwrap_or(CellValue::Empty),
                        formula: Some(self.text(pieces)),
                        style: None,
                    },
                };
                sheet.set_cell(CellRef::new(row, col), cell);
            }
            wb.sheets.push(sheet);
        }
        for (name, body) in &self.names {
            wb.defined_names.push(DefinedName {
                name: name.clone(),
                formula: self.text(body),
                local_sheet: None,
                hidden: false,
            });
        }
        wb
    }

    fn values(&self, wb: &Workbook) -> HashMap<usize, CellValue> {
        let mut values = HashMap::new();
        for (index, &slot) in self.order.iter().enumerate() {
            let Some(sheet) = wb.sheet(SheetId(index as u32)) else {
                continue;
            };
            for (&(row, col), (id, _)) in &self.cells[slot] {
                if let Some(cell) = sheet.cell(CellRef::new(row, col)) {
                    values.insert(*id, cell.value.clone());
                }
            }
        }
        values
    }

    /// inserts or deletes row or column `index` of `slot`: cells there move,
    /// and every reference into it moves, grows, shrinks or breaks.
    fn shift(&mut self, slot: usize, axis: usize, index: u32, shift: Shift) {
        let cells = std::mem::take(&mut self.cells[slot]);
        self.cells[slot] = cells
            .into_iter()
            .filter_map(|(at, cell)| Some((shift_point(at, axis, index, shift)?, cell)))
            .collect();
        for (owner, cells) in self.cells.iter_mut().enumerate() {
            for (_, content) in cells.values_mut() {
                if let Content::Formula(pieces) = content {
                    shift_pieces(pieces, Some(owner), slot, axis, index, shift);
                }
            }
        }
        for (_, body) in &mut self.names {
            shift_pieces(body, None, slot, axis, index, shift);
        }
    }

    fn apply(&mut self, edit: Edit) {
        match edit {
            Edit::Unchanged => {}
            Edit::Value => self.put(0, (1, 0), Content::Number(1000.0)),
            Edit::Formula => {
                let edited = vec![relative(None, 2, 0), text("*10")];
                self.put(0, (2, 1), formula(edited));
            }
            Edit::RenameSheet => self.slots[1].0 = "Renamed Two".into(),
            Edit::AddSheet => {
                self.slots.push(("Extra".into(), true));
                self.cells.push(BTreeMap::new());
                let slot = self.slots.len() - 1;
                self.put(slot, (0, 0), Content::Number(5.0));
                self.order.insert(1, slot);
            }
            Edit::DeleteSheet => {
                self.slots[1].1 = false;
                self.order.retain(|&slot| slot != 1);
                self.cells[1].clear();
            }
            Edit::AddName => {
                let later = vec![absolute(Some(0), 0, 0)];
                self.names.push(("Later".into(), later));
            }
            Edit::ChangeName => self.names[1].1 = vec![absolute(Some(0), 2, 0)],
            Edit::DeleteName => {
                self.names.remove(0);
            }
            Edit::InsertRow => self.shift(0, 0, 2, Shift::Insert),
            Edit::DeleteRow => self.shift(0, 0, 1, Shift::Delete),
            Edit::InsertColumn => self.shift(0, 1, 1, Shift::Insert),
            Edit::DeleteColumn => self.shift(0, 1, 1, Shift::Delete),
        }
    }
}

/// a workbook over four sheets whose shape and constants vary with `seed`:
/// columns of shifted relative formulas, columns repeating one absolute
/// text, chains, cross-sheet references and ranges, defined names, volatile
/// calls, a cycle and a formula over the per-formula cap.
fn generated(seed: u64) -> Model {
    let mut state = seed.wrapping_mul(0x9E37_79B9_7F4A_7C15) | 1;
    let mut draw = |bound: u64| -> u64 {
        state ^= state << 13;
        state ^= state >> 7;
        state ^= state << 17;
        state % bound
    };
    let mut model = Model {
        slots: vec![
            ("Sheet1".into(), true),
            ("Data Two".into(), true),
            ("Calc".into(), true),
            ("Big".into(), true),
        ],
        order: vec![0, 1, 2, 3],
        cells: vec![BTreeMap::new(); 4],
        names: Vec::new(),
        next_id: 0,
    };
    let factor = 1 + draw(3);
    for slot in [0, 1] {
        let rows = 6 + draw(6) as u32;
        for row in 0..rows {
            let value = (1 + draw(50)) as f64;
            model.put(slot, (row, 0), Content::Number(value));
            let shifted = vec![relative(None, row, 0), text(&format!("*{factor}+1"))];
            model.put(slot, (row, 1), formula(shifted));
            let repeated = vec![
                text("SUM("),
                Piece::Range {
                    sheet: None,
                    start: (0, 0),
                    end: (3, 0),
                    abs: true,
                },
                text(")+"),
                absolute(None, 0, 0),
            ];
            model.put(slot, (row, 2), formula(repeated));
            let mut chain = vec![relative(None, row, 1)];
            if row > 0 {
                chain.push(text("+"));
                chain.push(relative(None, row - 1, 3));
            }
            model.put(slot, (row, 3), formula(chain));
        }
    }
    model.put(0, (0, 4), formula(vec![text("RAND()")]));
    let now = vec![text("NOW()-NOW()+"), relative(None, 0, 0)];
    model.put(0, (1, 4), formula(now));
    model.put(0, (2, 4), formula(vec![text("INDIRECT(\"A2\")+1")]));
    model.put(0, (0, 5), formula(vec![text("Later*2")]));
    model.put(0, (1, 5), formula(vec![text("Extra!A1+1")]));
    let calc_rows = 4 + draw(4) as u32;
    for row in 0..calc_rows {
        let mut cross = vec![relative(Some(1), row, 1), text("+")];
        cross.push(absolute(Some(0), 0, 2));
        model.put(2, (row, 0), formula(cross));
    }
    let range = Piece::Range {
        sheet: Some(0),
        start: (0, 0),
        end: (5, 1),
        abs: false,
    };
    model.put(2, (0, 1), formula(vec![text("SUM("), range, text(")")]));
    model.put(2, (1, 1), formula(vec![text("SUM(Total)*Rate")]));
    model.put(2, (2, 1), formula(vec![text("SUM(Total)*Rate")]));
    let first = vec![relative(None, 0, 6), text("+1")];
    model.put(2, (0, 5), formula(first));
    let second = vec![relative(None, 0, 5), text("+1")];
    model.put(2, (0, 6), formula(second));
    model.put(2, (0, 7), formula(vec![text("SUM(Big!A1:XFD1048576)")]));
    model.put(3, (1_048_575, 16_383), Content::Number(1.0));
    let total = Piece::Range {
        sheet: Some(0),
        start: (0, 0),
        end: (3, 0),
        abs: true,
    };
    model.names.push(("Total".into(), vec![total]));
    let rate = vec![absolute(Some(1), 1, 0)];
    model.names.push(("Rate".into(), rate));
    model
}

struct Rebuilt {
    wb: Workbook,
    graph: DepGraph,
    result: RecalcResult,
}

fn rebuild(mut wb: Workbook, previous: Option<&DepGraph>) -> Rebuilt {
    let (graph, result) =
        rebuild_and_recalc_all_after(&mut wb, previous, Some(NOW_SERIAL), Some(RAND_SEED));
    Rebuilt { wb, graph, result }
}

fn evaluation_order(graph: &DepGraph) -> Vec<Step> {
    let every: HashSet<Key> = graph.formula_cells().map(|(s, c)| key(s, c)).collect();
    topo_order(graph, &every)
}

/// panics unless both rebuilds hold the same graph, parse cache, evaluation
/// order, recalc lists and workbook values.
fn assert_same_rebuild(reused: &Rebuilt, fresh: &Rebuilt, label: &str) {
    let graphs = catch_unwind(AssertUnwindSafe(|| {
        reused.graph.assert_matches(&fresh.graph);
    }));
    assert!(graphs.is_ok(), "{label}: graphs differ");
    assert_eq!(
        evaluation_order(&reused.graph),
        evaluation_order(&fresh.graph),
        "{label}"
    );
    assert_eq!(reused.result.changed, fresh.result.changed, "{label}");
    assert_eq!(
        reused.result.cycle_cells, fresh.result.cycle_cells,
        "{label}"
    );
    assert_eq!(
        reused.result.limited_cells, fresh.result.limited_cells,
        "{label}"
    );
    assert_eq!(
        reused.result.budget_exhausted, fresh.result.budget_exhausted,
        "{label}"
    );
    assert_eq!(
        reused.result.unevaluated_cells, fresh.result.unevaluated_cells,
        "{label}"
    );
    assert_eq!(reused.wb, fresh.wb, "{label}");
}

#[test]
fn a_full_rebuild_over_the_previous_graph_matches_a_fresh_build_after_each_edit() {
    for seed in 1..=6 {
        for edit in EDITS {
            let label = format!("seed {seed}, {edit:?}");
            let mut model = generated(seed);
            let mut before = model.render(&HashMap::new());
            let (mut previous, _) =
                rebuild_and_recalc_all_after(&mut before, None, Some(NOW_SERIAL), Some(RAND_SEED));
            let values = model.values(&before);
            model.apply(edit);
            let edited = model.render(&values);
            if edit == Edit::Formula {
                let (sheet, at) = (SheetId(0), CellRef::new(2, 1));
                let source = edited
                    .sheet(sheet)
                    .and_then(|s| s.cell(at))
                    .and_then(|c| c.formula.clone());
                previous.set_formula(sheet, at, source.as_deref());
            }
            let reused = rebuild(edited.clone(), Some(&previous));
            let fresh = rebuild(edited, None);
            assert!(!fresh.result.cycle_cells.is_empty(), "{label}");
            assert!(!fresh.result.limited_cells.is_empty(), "{label}");
            assert_same_rebuild(&reused, &fresh, &label);
        }
    }
}

#[test]
fn a_full_rebuild_keeps_the_parsed_tree_of_an_unchanged_formula() {
    let (sheet, at) = (SheetId(2), CellRef::new(0, 1));
    let mut model = generated(1);
    let mut before = model.render(&HashMap::new());
    let (previous, _) =
        rebuild_and_recalc_all_after(&mut before, None, Some(NOW_SERIAL), Some(RAND_SEED));
    let source = before
        .sheet(sheet)
        .and_then(|s| s.cell(at))
        .and_then(|c| c.formula.clone())
        .expect("a formula at Calc!B1");
    let kept = previous
        .asts()
        .lock()
        .expect("parse cache poisoned")
        .get(&source)
        .cloned()
        .expect("the previous graph parsed it");
    let values = model.values(&before);
    model.apply(Edit::Value);
    let edited = model.render(&values);
    let fresh = rebuild(edited.clone(), None);
    let reused = rebuild(edited, Some(&previous));
    let still = reused
        .wb
        .sheet(sheet)
        .and_then(|s| s.cell(at))
        .and_then(|c| c.formula.clone());
    assert_eq!(still.as_deref(), Some(source.as_str()));
    let cached = reused
        .graph
        .asts()
        .lock()
        .expect("parse cache poisoned")
        .get(&source)
        .cloned()
        .expect("the rebuilt graph holds it");
    assert!(Arc::ptr_eq(&kept, &cached));
    let node = reused.graph.ast(sheet, at).expect("a formula node");
    assert!(Arc::ptr_eq(&kept, &node));
    let plain = DepGraph::build(&fresh.wb);
    let parsed = plain
        .asts()
        .lock()
        .expect("parse cache poisoned")
        .get(&source)
        .cloned()
        .expect("a plain build parses it");
    assert!(!Arc::ptr_eq(&kept, &parsed));
    let node = plain.ast(sheet, at).expect("a formula node");
    assert!(!Arc::ptr_eq(&kept, &node));
    assert_same_rebuild(&reused, &fresh, "unchanged formula");
}

#[test]
fn the_rebuild_comparison_reports_a_wrong_parse_in_the_cache() {
    let model = generated(1);
    let wb = model.render(&HashMap::new());
    let reused = rebuild(wb.clone(), None);
    let fresh = rebuild(wb, None);
    assert_same_rebuild(&reused, &fresh, "unchanged");
    let source = reused
        .wb
        .sheet(SheetId(0))
        .and_then(|s| s.cell(CellRef::new(0, 1)))
        .and_then(|c| c.formula.clone())
        .expect("a formula at B1");
    let wrong = parse_formula("12345").expect("a number parses");
    reused
        .graph
        .asts()
        .lock()
        .expect("parse cache poisoned")
        .insert(source, Arc::new(wrong));
    let compared = catch_unwind(AssertUnwindSafe(|| {
        assert_same_rebuild(&reused, &fresh, "wrong parse");
    }));
    assert!(compared.is_err());
}

#[test]
fn a_graph_whose_parse_cache_was_taken_still_matches_a_fresh_build() {
    let (sheet, at) = (SheetId(0), CellRef::new(2, 1));
    let dirty = [(sheet, CellRef::new(1, 0)), (sheet, at)];
    for seed in 1..=6 {
        let label = format!("seed {seed}");
        let mut model = generated(seed);
        let before = model.render(&HashMap::new());
        let mut fresh = rebuild(before.clone(), None);
        let mut taken = rebuild(before, None);
        drop(DepGraph::build_reusing(&taken.wb, &taken.graph));
        let asts = taken.graph.asts().lock().expect("parse cache poisoned");
        assert!(asts.is_empty(), "{label}");
        drop(asts);
        let values = model.values(&taken.wb);
        model.apply(Edit::Value);
        model.apply(Edit::Formula);
        let edited = model.render(&values);
        let source = edited
            .sheet(sheet)
            .and_then(|s| s.cell(at))
            .and_then(|c| c.formula.clone());
        for rebuilt in [&mut taken, &mut fresh] {
            rebuilt.wb = edited.clone();
            rebuilt.graph.set_formula(sheet, at, source.as_deref());
            rebuilt.result = recalc_after_with_seed(
                &mut rebuilt.wb,
                &mut rebuilt.graph,
                &dirty,
                Some(NOW_SERIAL),
                Some(RAND_SEED),
            );
        }
        let kept: HashSet<String> = {
            let asts = taken.graph.asts().lock().expect("parse cache poisoned");
            asts.keys().cloned().collect()
        };
        fresh
            .graph
            .asts()
            .lock()
            .expect("parse cache poisoned")
            .retain(|parsed, _| kept.contains(parsed));
        assert_same_rebuild(&taken, &fresh, &label);
        let label = format!("{label}, full rebuild");
        let reused = rebuild(taken.wb.clone(), Some(&taken.graph));
        let rebuilt = rebuild(taken.wb, None);
        assert!(!rebuilt.result.cycle_cells.is_empty(), "{label}");
        assert_same_rebuild(&reused, &rebuilt, &label);
    }
}

fn at(source: &str) -> CellRef {
    CellRef::parse_a1(source).expect("an A1 address")
}

fn rect(source: &str) -> CellRange {
    CellRange::parse_a1(source).expect("an A1 range")
}

fn number(value: f64) -> CellValue {
    CellValue::Number { value }
}

fn sheets(names: &[&str]) -> Workbook {
    let mut wb = Workbook::default();
    wb.sheets.extend(names.iter().map(|name| Sheet::new(*name)));
    wb
}

fn put_number(wb: &mut Workbook, sheet: SheetId, address: &str, value: f64) {
    let cell = Cell {
        value: number(value),
        ..Cell::default()
    };
    wb.sheet_mut(sheet)
        .expect("a sheet")
        .set_cell(at(address), cell);
}

fn put_formula(wb: &mut Workbook, sheet: SheetId, address: &str, source: &str) {
    let cell = Cell {
        formula: Some(source.into()),
        ..Cell::default()
    };
    wb.sheet_mut(sheet)
        .expect("a sheet")
        .set_cell(at(address), cell);
}

fn shown(wb: &Workbook, sheet: SheetId, address: &str) -> CellValue {
    let cell = wb.sheet(sheet).and_then(|s| s.cell(at(address)));
    cell.map_or(CellValue::Empty, |cell| cell.value.clone())
}

struct Edited {
    reused: Rebuilt,
    fresh: Rebuilt,
    /// the parse cache of the previous graph before the reusing rebuild took it.
    cached: HashMap<String, Arc<Expr>>,
}

impl Edited {
    /// whether the reusing rebuild holds the very tree the previous graph
    /// parsed for `source`.
    fn kept(&self, source: &str) -> bool {
        let asts = self
            .reused
            .graph
            .asts()
            .lock()
            .expect("parse cache poisoned");
        match (self.cached.get(source), asts.get(source)) {
            (Some(before), Some(after)) => Arc::ptr_eq(before, after),
            _ => false,
        }
    }
}

/// recalculates `before`, applies `edit`, and rebuilds the result over the
/// previous graph and from scratch, which must agree.
fn rebuild_after(mut before: Workbook, edit: impl FnOnce(&mut Workbook), label: &str) -> Edited {
    let (previous, _) =
        rebuild_and_recalc_all_after(&mut before, None, Some(NOW_SERIAL), Some(RAND_SEED));
    let cached: HashMap<String, Arc<Expr>> = previous
        .asts()
        .lock()
        .expect("parse cache poisoned")
        .clone();
    edit(&mut before);
    let reused = rebuild(before.clone(), Some(&previous));
    let fresh = rebuild(before, None);
    assert_same_rebuild(&reused, &fresh, label);
    Edited {
        reused,
        fresh,
        cached,
    }
}

#[test]
fn a_rebuild_over_a_resized_table_matches_a_fresh_build() {
    let s = SheetId(0);
    let mut wb = sheets(&["Sheet1"]);
    for (address, qty) in [("B2", 10.0), ("B3", 20.0), ("B4", 30.0)] {
        put_number(&mut wb, s, address, qty);
    }
    wb.tables.push(Table {
        name: "Sales".into(),
        sheet: s,
        range: rect("A1:B4"),
        header_rows: 1,
        totals_rows: 0,
        columns: vec!["Item".into(), "Qty".into()],
    });
    put_formula(&mut wb, s, "D1", "SUM(Sales[Qty])");
    put_formula(&mut wb, s, "D2", "D1*2");
    let resize = |wb: &mut Workbook| {
        put_number(wb, s, "B5", 40.0);
        wb.tables[0].range = rect("A1:B5");
    };
    let edited = rebuild_after(wb, resize, "resized table");
    assert!(edited.kept("SUM(Sales[Qty])"));
    assert_eq!(shown(&edited.fresh.wb, s, "D1"), number(100.0));
    assert_eq!(shown(&edited.fresh.wb, s, "D2"), number(200.0));
}

#[test]
fn a_rebuild_over_array_formulas_matches_a_fresh_build() {
    let s = SheetId(0);
    for (rows, spill, total) in [(5.0, "C1:C5", 15.0), (2.0, "C1:C2", 3.0)] {
        let label = format!("a sequence of {rows}");
        let mut wb = sheets(&["Sheet1"]);
        for (row, a, b) in [(1, 3.0, 1.0), (2, 4.0, 2.0), (3, 5.0, 3.0)] {
            put_number(&mut wb, s, &format!("A{row}"), a);
            put_number(&mut wb, s, &format!("B{row}"), b);
        }
        put_formula(&mut wb, s, "C1", "_xlfn.SEQUENCE(A1)");
        put_formula(&mut wb, s, "D1", "SUM(C1:C10)");
        put_formula(&mut wb, s, "E1", "A1:A3*2");
        put_formula(&mut wb, s, "F1", "SUM(A1:A3*B1:B3)");
        let sheet = wb.sheet_mut(s).expect("a sheet");
        sheet.set_array_formula(at("C1"), rect("C1"));
        sheet.set_array_formula(at("E1"), rect("E1:E3"));
        sheet.set_array_formula(at("F1"), rect("F1"));
        let resize = |wb: &mut Workbook| put_number(wb, s, "A1", rows);
        let edited = rebuild_after(wb, resize, &label);
        for source in [
            "_xlfn.SEQUENCE(A1)",
            "SUM(C1:C10)",
            "A1:A3*2",
            "SUM(A1:A3*B1:B3)",
        ] {
            assert!(edited.kept(source), "{label}: {source}");
        }
        let fresh = &edited.fresh.wb;
        let spilled = fresh
            .sheet(s)
            .and_then(|sheet| sheet.array_formula(at("C1")));
        assert_eq!(spilled, Some(rect(spill)), "{label}");
        assert_eq!(shown(fresh, s, "D1"), number(total), "{label}");
        assert_eq!(shown(fresh, s, "E1"), number(rows * 2.0), "{label}");
        assert_eq!(shown(fresh, s, "E3"), number(10.0), "{label}");
        assert_eq!(shown(fresh, s, "F1"), number(rows + 8.0 + 15.0), "{label}");
    }
}

#[test]
fn a_rebuild_over_a_sheet_local_name_matches_a_fresh_build() {
    let (one, two) = (SheetId(0), SheetId(1));
    let mut wb = sheets(&["One", "Two"]);
    put_number(&mut wb, one, "A1", 2.0);
    put_number(&mut wb, two, "A1", 3.0);
    for (formula, local_sheet) in [("One!$A$1", None), ("Two!$A$1", Some(two))] {
        wb.defined_names.push(DefinedName {
            name: "Rate".into(),
            formula: formula.into(),
            local_sheet,
            hidden: false,
        });
    }
    put_formula(&mut wb, one, "B1", "Rate*10");
    put_formula(&mut wb, two, "B1", "Rate*10");
    put_formula(&mut wb, two, "B2", "One!Rate*10");
    let edit = |wb: &mut Workbook| put_number(wb, two, "A1", 4.0);
    let edited = rebuild_after(wb, edit, "sheet-local name");
    assert!(edited.kept("Rate*10"));
    assert!(edited.kept("One!Rate*10"));
    assert_eq!(shown(&edited.fresh.wb, one, "B1"), number(20.0));
    assert_eq!(shown(&edited.fresh.wb, two, "B1"), number(40.0));
    assert_eq!(shown(&edited.fresh.wb, two, "B2"), number(20.0));
}

#[test]
fn a_rebuild_over_stored_and_malformed_formula_text_matches_a_fresh_build() {
    let s = SheetId(0);
    let mut wb = sheets(&["Sheet1"]);
    put_number(&mut wb, s, "A1", 2.0);
    put_number(&mut wb, s, "A2", 3.0);
    wb.defined_names.push(DefinedName {
        name: "Scale".into(),
        formula: "=Sheet1!$A$2".into(),
        local_sheet: None,
        hidden: false,
    });
    put_formula(&mut wb, s, "B1", "=A1+1");
    put_formula(&mut wb, s, "B2", "A1+1");
    put_formula(&mut wb, s, "B3", "=SUM(A1:A2)*Scale");
    let malformed = Cell {
        value: number(7.0),
        formula: Some("SUM(A1".into()),
        style: None,
    };
    wb.sheet_mut(s)
        .expect("a sheet")
        .set_cell(at("C1"), malformed);
    put_formula(&mut wb, s, "C2", "A2*2");
    let edit = |wb: &mut Workbook| {
        put_number(wb, s, "A1", 5.0);
        put_formula(wb, s, "C2", "A2*");
    };
    let edited = rebuild_after(wb, edit, "stored and malformed text");
    assert!(edited.kept("A1+1"));
    assert!(edited.kept("SUM(A1:A2)*Scale"));
    assert!(!edited.cached.contains_key("SUM(A1"));
    for rebuilt in [&edited.reused, &edited.fresh] {
        let asts = rebuilt.graph.asts().lock().expect("parse cache poisoned");
        assert!(!asts.contains_key("A2*"));
        drop(asts);
        assert_eq!(shown(&rebuilt.wb, s, "C2"), CellValue::Empty);
    }
    let fresh = &edited.fresh.wb;
    assert_eq!(shown(fresh, s, "B1"), number(6.0));
    assert_eq!(shown(fresh, s, "B2"), number(6.0));
    assert_eq!(shown(fresh, s, "B3"), number(24.0));
    assert_eq!(shown(fresh, s, "C1"), number(7.0));
}

#[test]
fn a_rebuild_after_a_formula_gets_new_text_matches_a_fresh_build() {
    let s = SheetId(0);
    let mut wb = sheets(&["Sheet1"]);
    for (address, value) in [("A1", 1.0), ("A2", 2.0), ("A3", 3.0)] {
        put_number(&mut wb, s, address, value);
    }
    put_formula(&mut wb, s, "B1", "A1*2");
    put_formula(&mut wb, s, "B2", "A2*2");
    put_formula(&mut wb, s, "B3", "SUM(B1:B2)");
    put_formula(&mut wb, s, "B4", "A3*2");
    let edit = |wb: &mut Workbook| put_formula(wb, s, "B2", "A2*3");
    let edited = rebuild_after(wb, edit, "new formula text");
    for source in ["A1*2", "SUM(B1:B2)", "A3*2"] {
        assert!(edited.kept(source), "{source}");
    }
    assert!(!edited.cached.contains_key("A2*3"));
    let asts = edited
        .reused
        .graph
        .asts()
        .lock()
        .expect("parse cache poisoned");
    assert!(asts.contains_key("A2*3"));
    assert!(!asts.contains_key("A2*2"));
    drop(asts);
    assert_eq!(shown(&edited.fresh.wb, s, "B2"), number(6.0));
    assert_eq!(shown(&edited.fresh.wb, s, "B3"), number(8.0));
}
