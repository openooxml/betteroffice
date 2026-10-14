use std::num::NonZeroUsize;

use xlsx_model::{Cell, CellRange, CellRef, ColId, RowId, Sheet, SheetId, Workbook};

use super::{DepGraph, NodeKey};

/// Builds a dependency graph in bounded visits without evaluating formulas.
#[doc(hidden)]
pub struct DepGraphBuilder {
    graph: DepGraph,
    phase: Phase,
    index: usize,
    last_cell: Option<(RowId, ColId)>,
}

enum Phase {
    SheetNames,
    DefinedNames,
    Tables,
    Cells,
    Done,
}

impl DepGraphBuilder {
    pub fn new() -> Self {
        Self {
            graph: DepGraph::empty(),
            phase: Phase::SheetNames,
            index: 0,
            last_cell: None,
        }
    }

    /// Visits at most `max_items` entries of the same unchanged workbook.
    /// Each visit handles metadata, a stored cell, or a sheet boundary.
    pub fn advance(&mut self, model: &Workbook, max_items: NonZeroUsize) -> bool {
        let mut remaining = max_items.get();
        loop {
            match self.phase {
                Phase::SheetNames => {
                    let Some(sheet) = model.sheets.get(self.index) else {
                        self.phase = Phase::DefinedNames;
                        self.index = 0;
                        continue;
                    };
                    if remaining == 0 {
                        return false;
                    }
                    self.graph
                        .names
                        .insert(sheet.name.to_lowercase(), SheetId(self.index as u32));
                    self.index += 1;
                    remaining -= 1;
                }
                Phase::DefinedNames => {
                    let Some(defined) = model.defined_names.get(self.index) else {
                        self.phase = Phase::Tables;
                        self.index = 0;
                        continue;
                    };
                    if remaining == 0 {
                        return false;
                    }
                    self.graph
                        .defined_name_indices
                        .entry((defined.local_sheet, defined.name.to_ascii_lowercase()))
                        .or_insert(self.index);
                    self.graph.defined_names.push(defined.clone());
                    self.index += 1;
                    remaining -= 1;
                }
                Phase::Tables => {
                    let Some(table) = model.tables.get(self.index) else {
                        self.phase = Phase::Cells;
                        self.index = 0;
                        continue;
                    };
                    if remaining == 0 {
                        return false;
                    }
                    self.graph
                        .tables
                        .insert(table.name.to_lowercase(), table.clone());
                    self.index += 1;
                    remaining -= 1;
                }
                Phase::Cells => {
                    let Some(sheet) = model.sheets.get(self.index) else {
                        self.phase = Phase::Done;
                        continue;
                    };
                    if remaining == 0 {
                        return false;
                    }
                    remaining -= 1;
                    let Some((at, cell)) = next_cell(sheet, self.last_cell) else {
                        self.index += 1;
                        self.last_cell = None;
                        continue;
                    };
                    self.last_cell = Some((at.row, at.col));
                    if let Some(src) = &cell.formula {
                        let key = NodeKey::new(SheetId(self.index as u32), at);
                        self.graph.install(key, src);
                        if let Some(range) = sheet.array_formula(at) {
                            self.graph.install_spill(key, range);
                        }
                    }
                }
                Phase::Done => return true,
            }
        }
    }

    /// Returns the completed graph, or `None` while work remains.
    pub fn finish(self) -> Option<DepGraph> {
        matches!(self.phase, Phase::Done).then_some(self.graph)
    }
}

impl Default for DepGraphBuilder {
    fn default() -> Self {
        Self::new()
    }
}

pub(super) fn next_cell(sheet: &Sheet, last: Option<(RowId, ColId)>) -> Option<(CellRef, &Cell)> {
    let row = match last {
        Some((row, col)) => {
            if let Some(col) = col.checked_add(1)
                && let Some(cell) = sheet
                    .cells_in_range(CellRange::new(
                        CellRef::new(row, col),
                        CellRef::new(row, ColId::MAX),
                    ))
                    .next()
            {
                return Some(cell);
            }
            row.checked_add(1)?
        }
        None => 0,
    };
    sheet
        .cells_in_range(CellRange::new(
            CellRef::new(row, 0),
            CellRef::new(RowId::MAX, ColId::MAX),
        ))
        .next()
}

#[cfg(test)]
mod tests {
    use xlsx_model::{CellValue, DefinedName, Table};

    use super::*;

    fn a1(source: &str) -> CellRef {
        CellRef::parse_a1(source).unwrap()
    }

    fn range(source: &str) -> CellRange {
        CellRange::parse_a1(source).unwrap()
    }

    fn formula(source: &str) -> Cell {
        Cell {
            value: CellValue::Number { value: -0.0 },
            formula: Some(source.into()),
            style: Some(0),
        }
    }

    fn defined(name: &str, formula: &str, local_sheet: Option<SheetId>) -> DefinedName {
        DefinedName {
            name: name.into(),
            formula: formula.into(),
            local_sheet,
            hidden: false,
        }
    }

    fn table(name: &str, sheet: SheetId, source: &str) -> Table {
        Table {
            name: name.into(),
            sheet,
            range: range(source),
            header_rows: 1,
            totals_rows: 1,
            columns: vec!["Qty".into(), "Cost".into()],
        }
    }

    fn workbook() -> Workbook {
        let mut model = Workbook::default();
        model.sheets.extend([
            Sheet::new("Main"),
            Sheet::new("Data"),
            Sheet::new("DATA"),
            Sheet::new("Empty"),
        ]);
        model.defined_names.extend([
            defined("Inputs", "=Data!$A$1:$B$3", None),
            defined("inputs", "Main!Z1", None),
            defined("Inputs", "$C$1", Some(SheetId(0))),
            defined("INPUTS", "Z2", Some(SheetId(0))),
            defined("Clock", "=NOW()", None),
            defined("Nested", "SUM(Inputs)+Clock+SUM(Sales[Qty])", None),
            defined("RowInput", "Sales[[#This Row],[Qty]]", Some(SheetId(0))),
            defined("CycleLeft", "CycleRight", None),
            defined("CycleRight", "CycleLeft", None),
            defined("Broken", "(", None),
        ]);
        model.defined_names[1].hidden = true;
        model.defined_names[3].hidden = true;
        model.tables.extend([
            table("Sales", SheetId(1), "B1:C4"),
            table("SALES", SheetId(2), "D1:E5"),
        ]);
        let main = &mut model.sheets[0];
        for (at, source) in [
            ("Z1000000", "Data!A1+Main!A1+Ghost!A1"),
            ("F10", "SEQUENCE(3,2)"),
            ("F2", "SUM(Sales[Qty])+SUM(SALES[[#Totals],[Cost]])"),
            ("E2", "SUM(Nested)+Data!Inputs+RowInput"),
            ("D2", "SUM(Sales[[#This Row],[Qty]])"),
            ("C2", "Clock+SUM(Inputs)"),
            ("B2", "TODAY()+RAND()+RANDBETWEEN(1,2)+INDIRECT(\"A1\")"),
            ("A2", "OFFSET(A1,B1,0)"),
            ("K1", "D3+F11"),
            ("H1", ""),
            ("G1", "("),
            ("F1", "{1,-0;2,3}"),
            ("E1", "SUM(A1:OFFSET(A1,0,B1))"),
            ("D1", "SUM(A1:INDEX(A1:A9,3))"),
            ("C1", "OFFSET(A1,0,1)"),
            ("A1", "SUM(Inputs)+Data!A1+$C$1+Main!C1"),
            ("A3", "CycleLeft+Broken"),
            ("B5", "NOW()"),
        ] {
            main.set_cell(a1(at), formula(source));
        }
        main.set_cell(
            a1("B1000000"),
            Cell {
                value: CellValue::Number {
                    value: f64::from_bits(0x7ff8_0000_0000_0042),
                },
                formula: None,
                style: Some(0),
            },
        );
        main.set_array_formula(a1("F10"), range("F10:G12"));
        main.set_array_formula(a1("D2"), range("$D$2:$E$3"));
        main.set_array_formula(a1("A1"), range("A1"));
        main.set_array_formula(a1("G1"), range("G1:G3"));
        main.set_array_formula(a1("H1"), range("H1:H3"));
        main.set_array_formula(a1("M12"), range("M12:N13"));
        main.set_array_formula(a1("B1000000"), range("B1000000:C1000001"));
        model.sheets[1].set_cell(a1("A1"), formula("SUM(Inputs)"));
        model.sheets[2].set_cell(a1("A1"), formula("Main!A1+SUM(Sales[Qty])"));
        model.sheets[2].set_cell(a1("C3"), formula("NOW()"));
        model.sheets[2].set_array_formula(a1("C3"), range("C3:D4"));
        model
    }

    #[test]
    fn sliced_graph_matches_build_names_edges_volatility_and_spills() {
        let empty = Workbook::default();
        let mut metadata_only = Workbook::default();
        metadata_only
            .defined_names
            .push(defined("Clock", "=NOW()", None));
        metadata_only
            .tables
            .push(table("Detached", SheetId(0), "A1:B4"));
        let model = workbook();
        for model in [&empty, &metadata_only, &model] {
            let expected = DepGraph::build(model);
            let mut snapshot = crate::graph::SnapshotGraphBuilder::new();
            loop {
                let (ready, records, bytes) = snapshot.advance(model, 512).unwrap();
                assert!(records <= 1);
                assert!(bytes <= 512);
                if ready {
                    break;
                }
            }
            snapshot.finish().unwrap().assert_matches(&expected);
            let visits = model.sheets.len() * 2
                + model.defined_names.len()
                + model.tables.len()
                + model
                    .sheets
                    .iter()
                    .map(|sheet| sheet.iter_cells().count())
                    .sum::<usize>();
            for max_items in [1, 2, 7, usize::MAX] {
                let budget = NonZeroUsize::new(max_items).unwrap();
                let mut builder = DepGraphBuilder::new();
                let mut calls = 1;
                while !builder.advance(model, budget) {
                    calls += 1;
                    assert!(calls <= visits.max(1));
                }
                assert_eq!(calls, visits.div_ceil(max_items).max(1));
                assert!(builder.advance(model, budget));
                let actual = builder.finish().unwrap();
                actual.assert_matches(&expected);
                if !model.sheets.is_empty() {
                    assert_eq!(actual.names["data"], SheetId(2));
                    assert_eq!(actual.tables["sales"], model.tables[1]);
                    assert_eq!(actual.defined_name_indices[&(None, "inputs".into())], 0);
                    assert_eq!(
                        actual.defined_name_indices[&(Some(SheetId(0)), "inputs".into())],
                        2
                    );
                    assert!(!actual.is_formula(SheetId(0), a1("G1")));
                    assert!(!actual.is_formula(SheetId(0), a1("H1")));
                    assert!(
                        actual
                            .volatile
                            .contains(&NodeKey::new(SheetId(0), a1("E2")))
                    );
                    assert!(
                        !actual
                            .volatile
                            .contains(&NodeKey::new(SheetId(0), a1("C1")))
                    );
                    assert_eq!(actual.spills.len(), 3);
                    assert_eq!(
                        actual.spills_by_sheet[&SheetId(0)],
                        vec![
                            (NodeKey::new(SheetId(0), a1("D2")), range("$D$2:$E$3")),
                            (NodeKey::new(SheetId(0), a1("F10")), range("F10:G12")),
                        ]
                    );
                    assert_eq!(
                        actual
                            .dependents_of(SheetId(0), a1("D2"))
                            .collect::<Vec<_>>(),
                        expected
                            .dependents_of(SheetId(0), a1("D2"))
                            .collect::<Vec<_>>()
                    );
                    assert_eq!(
                        actual
                            .spill_sources(SheetId(0), range("D3:G11"))
                            .collect::<Vec<_>>(),
                        vec![(SheetId(0), a1("D2")), (SheetId(0), a1("F10"))]
                    );
                    let cached = &model.sheets[0].cell(a1("B1000000")).unwrap().value;
                    let CellValue::Number { value } = cached else {
                        panic!("expected cached number");
                    };
                    assert_eq!(value.to_bits(), 0x7ff8_0000_0000_0042);
                    let cached = model.sheets[0].cell(a1("A1")).unwrap();
                    let CellValue::Number { value } = &cached.value else {
                        panic!("expected cached number");
                    };
                    assert_eq!(value.to_bits(), (-0.0_f64).to_bits());
                    assert_eq!(cached.style, Some(0));
                }
            }
        }
    }

    #[test]
    fn graph_budget_resumes_metadata_and_sparse_cells() {
        let mut model = Workbook::default();
        model.sheets.extend([
            Sheet::new("Lead"),
            Sheet::new("Main"),
            Sheet::new("Middle"),
            Sheet::new("Tail"),
        ]);
        model.defined_names.extend([
            defined("Named", "Main!A1", None),
            defined("NAMED", "Main!B1", None),
            defined("Local", "=TODAY()", Some(SheetId(1))),
        ]);
        model.tables.extend([
            table("Items", SheetId(1), "B1:C4"),
            table("ITEMS", SheetId(1), "D1:E5"),
        ]);
        let cells = [
            (a1("Z1"), Some("Named+SUM(Items[Qty])")),
            (CellRef::new(0, ColId::MAX), Some("D1")),
            (a1("A500000"), Some("Local")),
            (a1("Z500000"), None),
            (a1("A1048576"), Some("A1")),
            (a1("XFD1048576"), Some("(")),
            (CellRef::new(RowId::MAX, 0), Some("B1")),
            (CellRef::new(RowId::MAX, ColId::MAX), Some("C1")),
        ];
        for (at, source) in cells.iter().rev() {
            model.sheets[1].set_cell(
                *at,
                Cell {
                    value: CellValue::Empty,
                    formula: source.map(str::to_string),
                    style: Some(0),
                },
            );
        }
        model.sheets[1].set_array_formula(a1("A500000"), range("A500000:B500001"));
        assert!(DepGraphBuilder::new().finish().is_none());
        let mut incomplete = DepGraphBuilder::new();
        let one = NonZeroUsize::new(1).unwrap();
        assert!(!incomplete.advance(&model, one));
        assert!(incomplete.finish().is_none());

        let mut builder = DepGraphBuilder::default();
        for count in 1..=model.sheets.len() {
            assert!(!builder.advance(&model, one));
            assert_eq!(builder.graph.names.len(), count);
            assert!(builder.graph.defined_names.is_empty());
            assert!(builder.graph.tables.is_empty());
            assert!(builder.graph.deps.is_empty());
        }
        for count in 1..=model.defined_names.len() {
            assert!(!builder.advance(&model, one));
            assert_eq!(builder.graph.defined_names.len(), count);
            assert!(builder.graph.tables.is_empty());
            assert!(builder.graph.deps.is_empty());
        }
        assert_eq!(
            builder.graph.defined_name_indices[&(None, "named".into())],
            0
        );
        for table in &model.tables {
            assert!(!builder.advance(&model, one));
            assert_eq!(builder.graph.tables.len(), 1);
            assert_eq!(&builder.graph.tables["items"], table);
            assert!(builder.graph.deps.is_empty());
        }
        assert!(!builder.advance(&model, one));
        assert_eq!(builder.index, 1);
        assert!(builder.last_cell.is_none());
        assert!(builder.graph.deps.is_empty());

        let mut installed = 0;
        for (at, source) in cells {
            assert!(!builder.advance(&model, one));
            if source.is_some_and(|source| source != "(") {
                installed += 1;
            }
            assert_eq!(builder.index, 1);
            assert_eq!(builder.last_cell, Some((at.row, at.col)));
            assert_eq!(builder.graph.deps.len(), installed);
        }
        assert_eq!(builder.graph.spills.len(), 1);
        for index in 2..model.sheets.len() {
            assert!(!builder.advance(&model, one));
            assert_eq!(builder.index, index);
            assert!(builder.last_cell.is_none());
        }
        assert!(builder.advance(&model, one));
        assert!(builder.advance(&model, one));
        builder
            .finish()
            .unwrap()
            .assert_matches(&DepGraph::build(&model));
    }

    #[test]
    fn snapshot_graph_bounds_name_expansion_and_wide_table_clones() {
        let mut model = Workbook::default();
        model.sheets.push(Sheet::new("Data"));
        for index in 0..300 {
            model.defined_names.push(defined(
                &format!("Name{index}"),
                &if index == 299 {
                    "A1+NOW()".to_owned()
                } else {
                    format!("Name{}+A{}", index + 1, index + 1)
                },
                None,
            ));
        }
        let mut wide = table("Wide", SheetId(0), "A1:XFD4");
        wide.columns = (0..2_000).map(|index| format!("Column{index}")).collect();
        model.tables.push(wide);
        model.sheets[0].set_cell(a1("A5"), formula("Name0+SUM(Wide[Column1999])"));
        let expected = DepGraph::build(&model);
        let mut builder = crate::graph::SnapshotGraphBuilder::new();
        let mut steps = 0;
        loop {
            let (ready, records, bytes) = builder.advance(&model, 256).unwrap();
            assert!(records <= 1);
            assert!(bytes <= 256);
            steps += 1;
            assert!(steps < 20_000);
            if ready {
                break;
            }
        }
        assert!(steps > 2_000);
        builder.finish().unwrap().assert_matches(&expected);
    }
}
