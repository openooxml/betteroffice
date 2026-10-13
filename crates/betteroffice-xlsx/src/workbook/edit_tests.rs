use super::*;
use crate::authority::{fast_set_cell_count, with_full_materialization};

pub(super) fn r(a1: &str) -> CellRef {
    CellRef::parse_a1(a1).unwrap()
}

pub(super) fn options() -> CalculationOptions {
    CalculationOptions {
        now_serial: Some(45_000.25),
    }
}

pub(super) fn workbook_bytes() -> Vec<u8> {
    let parts = [
        (
            "[Content_Types].xml",
            r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
<Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
<Override PartName="/xl/worksheets/sheet3.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>"#,
        ),
        (
            "_rels/.rels",
            r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>"#,
        ),
        (
            "xl/workbook.xml",
            r#"<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheets>
<sheet name="Sheet1" sheetId="1" r:id="rId1"/>
<sheet name="Sheet2" sheetId="2" r:id="rId2"/>
<sheet name="Sheet3" sheetId="3" r:id="rId3"/>
</sheets>
</workbook>"#,
        ),
        (
            "xl/_rels/workbook.xml.rels",
            r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>
<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet3.xml"/>
<Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
<Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>
</Relationships>"#,
        ),
        (
            "xl/styles.xml",
            r#"<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<fonts count="2">
<font>
<sz val="11"/>
<name val="Calibri"/>
</font>
<font>
<b/>
<sz val="11"/>
<name val="Calibri"/>
</font>
</fonts>
<fills count="2">
<fill>
<patternFill patternType="none"/>
</fill>
<fill>
<patternFill patternType="gray125"/>
</fill>
</fills>
<borders count="1">
<border>
<left/>
<right/>
<top/>
<bottom/>
<diagonal/>
</border>
</borders>
<cellStyleXfs count="1">
<xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
</cellStyleXfs>
<cellXfs count="2">
<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/>
</cellXfs>
<cellStyles count="1">
<cellStyle name="Normal" xfId="0" builtinId="0"/>
</cellStyles>
</styleSheet>"#,
        ),
        (
            "xl/sharedStrings.xml",
            r#"<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="5" uniqueCount="3">
<si>
<r>
<rPr>
<b/>
</rPr>
<t>Alpha</t>
</r>
<r>
<t xml:space="preserve"> rich</t>
</r>
</si>
<si>
<t>Alpha rich</t>
</si>
<si>
<t>Beta</t>
</si>
</sst>"#,
        ),
        (
            "xl/worksheets/sheet1.xml",
            r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<sheetData>
<row r="1">
<c r="A1">
<v>2</v>
</c>
<c r="B1">
<v>3</v>
</c>
<c r="C1">
<f>A1+B1</f>
<v>5</v>
</c>
<c r="D1">
<f>C1*2</f>
<v>10</v>
</c>
</row>
<row r="2">
<c r="A2" t="s">
<v>0</v>
</c>
<c r="B2" t="s">
<v>1</v>
</c>
<c r="C2" t="s">
<v>2</v>
</c>
<c r="D2" s="1">
<v>9</v>
</c>
</row>
<row r="3">
<c r="A3">
<f t="shared" si="0" ref="A3:B3">A1*2</f>
<v>4</v>
</c>
<c r="B3">
<f t="shared" si="0"></f>
<v>6</v>
</c>
</row>
<row r="4">
<c r="A4">
<f t="array" ref="A4:B4">A1:B1*2</f>
<v>4</v>
</c>
<c r="B4">
<v>6</v>
</c>
</row>
<row r="5">
<c r="A5">
<f>NOW()</f>
<v>45000.25</v>
</c>
<c r="B5">
<f>TODAY()</f>
<v>45000</v>
</c>
<c r="C5">
<f>A5+B5</f>
<v>90000.25</v>
</c>
</row>
</sheetData>
</worksheet>"#,
        ),
        (
            "xl/worksheets/sheet2.xml",
            r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<sheetData>
<row r="1">
<c r="A1">
<v>7</v>
</c>
<c r="B1" t="s">
<v>0</v>
</c>
<c r="C1">
<f>Sheet1!C1+A1</f>
<v>12</v>
</c>
<c r="D1" s="1">
<v>8</v>
</c>
</row>
</sheetData>
</worksheet>"#,
        ),
        (
            "xl/worksheets/sheet3.xml",
            r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<sheetData>
<row r="1">
<c r="A1">
<v>1</v>
</c>
<c r="B1">
<f>Sheet2!C1+A1</f>
<v>13</v>
</c>
<c r="C1" t="s">
<v>2</v>
</c>
</row>
</sheetData>
</worksheet>"#,
        ),
    ];
    let parts = parts.map(|(name, xml)| (name.to_string(), xml.as_bytes()));
    ooxml_opc::rezip_parts_borrowed(&parts).unwrap()
}

pub(super) fn open_identical(
    bytes: &[u8],
    collaborative: bool,
    rand_seed: Option<u32>,
) -> Workbook {
    let mut workbook = Workbook::open_internal(bytes, true, Some(73)).unwrap();
    if !collaborative {
        workbook.mode = WorkbookMode::Standalone;
    }
    workbook.version_nonce = "singleton-oracle".into();
    workbook.set_rand_seed(rand_seed);
    workbook
}

struct Lockstep {
    fast: Workbook,
    oracle: Workbook,
    fast_events: Arc<Mutex<Vec<UpdateEvent>>>,
    oracle_events: Arc<Mutex<Vec<UpdateEvent>>>,
    _subscriptions: [UpdateSubscription; 2],
    step: usize,
}

impl Lockstep {
    fn new(collaborative: bool) -> Self {
        Self::with_rand_seed(collaborative, None)
    }

    fn with_rand_seed(collaborative: bool, rand_seed: Option<u32>) -> Self {
        let bytes = workbook_bytes();
        let fast = open_identical(&bytes, collaborative, rand_seed);
        let oracle = open_identical(&bytes, collaborative, rand_seed);
        let fast_events = Arc::new(Mutex::new(Vec::new()));
        let oracle_events = Arc::new(Mutex::new(Vec::new()));
        let events = fast_events.clone();
        let fast_subscription = fast
            .observe_update_v1(move |event| {
                events.lock().unwrap().push(event);
            })
            .unwrap();
        let events = oracle_events.clone();
        let oracle_subscription = oracle
            .observe_update_v1(move |event| {
                events.lock().unwrap().push(event);
            })
            .unwrap();
        let pair = Self {
            fast,
            oracle,
            fast_events,
            oracle_events,
            _subscriptions: [fast_subscription, oracle_subscription],
            step: 0,
        };
        pair.assert_equal(true);
        pair
    }

    fn assert_equal(&self, save: bool) {
        assert_eq!(
            self.fast.model(),
            self.oracle.model(),
            "model at step {}",
            self.step
        );
        assert_eq!(
            self.fast.authority.encode_state_vector_v1(),
            self.oracle.authority.encode_state_vector_v1(),
            "vector at step {}",
            self.step
        );
        assert_eq!(
            self.fast.authority.encode_state_as_update_v1(),
            self.oracle.authority.encode_state_as_update_v1(),
            "state at step {}",
            self.step
        );
        assert_eq!(
            self.fast.history_state(),
            self.oracle.history_state(),
            "history at step {}",
            self.step
        );
        assert_eq!(
            self.fast.version(),
            self.oracle.version(),
            "version at step {}",
            self.step
        );
        assert_eq!(self.fast.preserved.origins, self.oracle.preserved.origins);
        assert_eq!(
            self.fast.preserved.shared_string_cells,
            self.oracle.preserved.shared_string_cells
        );
        assert_eq!(self.fast.preserved.axes, self.oracle.preserved.axes);
        assert_eq!(self.fast.preserved.created, self.oracle.preserved.created);
        let fast_events = std::mem::take(&mut *self.fast_events.lock().unwrap());
        let oracle_events = std::mem::take(&mut *self.oracle_events.lock().unwrap());
        assert_eq!(fast_events, oracle_events, "events at step {}", self.step);
        if save {
            assert_eq!(
                self.fast.save().unwrap(),
                self.oracle.save().unwrap(),
                "save at step {}",
                self.step
            );
        }
    }

    fn step(
        &mut self,
        save: bool,
        apply: impl Fn(&mut Workbook) -> Result<MutationResult>,
    ) -> std::result::Result<MutationResult, String> {
        self.step += 1;
        let result = apply(&mut self.fast).map_err(|error| format!("{error:?}"));
        let expected = with_full_materialization(|| apply(&mut self.oracle))
            .map_err(|error| format!("{error:?}"));
        assert_eq!(result, expected, "result at step {}", self.step);
        self.assert_equal(save);
        result
    }

    fn edit(
        &mut self,
        sheet: u32,
        at: CellRef,
        input: &str,
    ) -> std::result::Result<MutationResult, String> {
        self.step(true, |workbook| {
            workbook.edit_cell(SheetId(sheet), at, input, options())
        })
    }
}

#[test]
fn edit_cell_lockstep_matches_full_materialization() {
    let mut pair = Lockstep::new(false);
    let count = fast_set_cell_count();
    assert_eq!(
        pair.fast.model.sheets[0]
            .cell(r("B3"))
            .unwrap()
            .formula
            .as_deref(),
        Some("B1*2")
    );
    assert_eq!(
        pair.fast.model.sheets[0].array_formula(r("A4")),
        Some(CellRange::parse_a1("A4:B4").unwrap())
    );
    assert_eq!(
        pair.fast.preserved.shared_string_cells[0].get(&(1, 0)),
        Some(&0)
    );
    assert_eq!(
        pair.fast.preserved.shared_string_cells[0].get(&(1, 1)),
        Some(&1)
    );
    assert!(!pair.edit(0, r("A1"), "2").unwrap().applied);
    assert!(pair.edit(99, r("A1"), "3").is_err());
    assert!(pair.edit(0, CellRef::new(MAX_ROWS, 0), "3").is_err());
    assert!(pair.edit(0, CellRef::new(0, MAX_COLS), "3").is_err());
    let long_text = "x".repeat(xlsx_calc::eval::MAX_CELL_TEXT_CHARS + 1);
    let long_formula = format!("={}", "x".repeat(xlsx_calc::lexer::MAX_FORMULA_BYTES + 1));
    assert!(pair.edit(0, r("E1"), &long_text).is_err());
    assert!(pair.edit(0, r("E1"), &long_formula).is_err());
    assert!(pair.edit(0, r("E1"), "=SUM(").unwrap().applied);
    assert_eq!(
        pair.fast.model.sheets[0].cell(r("E1")).unwrap().value,
        CellValue::Text {
            value: "=SUM(".into(),
        }
    );
    pair.edit(0, r("E2"), "123").unwrap();
    let history = pair.fast.preserved_undo.last().unwrap();
    assert!(Arc::ptr_eq(
        &pair.fast.preserved.shared_string_cells,
        &history.before.shared_string_cells
    ));
    assert!(Arc::ptr_eq(
        &pair.fast.preserved.shared_string_cells,
        &history.after.shared_string_cells
    ));
    pair.edit(0, r("E2"), "").unwrap();
    let style = pair.fast.model.sheets[0].cell(r("D2")).unwrap().style;
    pair.edit(0, r("D2"), "123").unwrap();
    assert_eq!(
        pair.fast.model.sheets[0].cell(r("D2")).unwrap().style,
        style
    );
    pair.edit(0, r("D2"), "").unwrap();
    pair.edit(0, r("C1"), "=A1*B1").unwrap();
    pair.edit(0, r("A1"), "10").unwrap();
    pair.edit(0, r("C1"), "17").unwrap();
    pair.edit(0, r("B3"), "=B1+3").unwrap();
    pair.edit(0, r("A3"), "8").unwrap();
    pair.edit(0, r("B4"), "19").unwrap();
    pair.edit(0, r("A4"), "=A1:B1*3").unwrap();
    pair.edit(0, r("A2"), "changed rich").unwrap();
    pair.step(true, |workbook| workbook.undo(options()))
        .unwrap();
    pair.step(true, |workbook| workbook.redo(options()))
        .unwrap();
    pair.edit(0, r("B2"), "duplicate replacement").unwrap();
    pair.edit(0, r("A5"), "=NOW()+1").unwrap();
    pair.edit(0, r("B5"), "=TODAY()+1").unwrap();
    assert!(!pair.edit(0, r("A5"), "=NOW()+1").unwrap().applied);
    pair.edit(1, r("A1"), "11").unwrap();
    pair.edit(1, r("D1"), "=Sheet1!A1+1").unwrap();
    pair.step(true, |workbook| workbook.undo(options()))
        .unwrap();
    pair.step(true, |workbook| workbook.redo(options()))
        .unwrap();
    for op in [
        Op::InsertRows {
            sheet: SheetId(0),
            at: 1,
            count: 1,
        },
        Op::DeleteCols {
            sheet: SheetId(1),
            at: 1,
            count: 1,
        },
        Op::AddSheet {
            index: 1,
            name: "Added".into(),
        },
        Op::RemoveSheet { index: 1 },
        Op::MoveSheet { from: 0, to: 2 },
        Op::MoveSheet { from: 2, to: 0 },
    ] {
        pair.step(true, |workbook| {
            workbook.apply_ops(vec![op.clone()], options())
        })
        .unwrap();
        pair.step(true, |workbook| workbook.undo(options()))
            .unwrap();
        pair.step(true, |workbook| workbook.redo(options()))
            .unwrap();
    }
    pair.edit(0, r("E8"), "123").unwrap();
    pair.edit(0, r("E8"), "124").unwrap();
    assert!(fast_set_cell_count() > count + 15);
    pair.assert_equal(true);
}

#[test]
fn collaborative_singleton_lockstep_matches_full_materialization() {
    let mut pair = Lockstep::new(true);
    for (sheet, at, input) in [
        (0, "A1", "12"),
        (0, "C1", "=A1*B1"),
        (0, "A2", "changed rich"),
        (1, "D1", "=Sheet1!A1"),
        (0, "A5", "=NOW()+1"),
        (0, "B5", "=TODAY()+1"),
    ] {
        pair.edit(sheet, r(at), input).unwrap();
        pair.step(true, |workbook| workbook.undo(options()))
            .unwrap();
        pair.step(true, |workbook| workbook.redo(options()))
            .unwrap();
    }
    pair.assert_equal(true);
}

#[test]
fn rand_seeded_lockstep_matches_full_materialization() {
    let mut pair = Lockstep::with_rand_seed(false, Some(0x5eed));
    let count = fast_set_cell_count();
    for (sheet, at, input) in [
        (0, "A5", "=RANDBETWEEN(1,1000000)"),
        (0, "B5", "=RANDBETWEEN(1,1000)+A1"),
        (0, "A1", "7"),
        (1, "D1", "=RANDBETWEEN(1,1000)*Sheet1!A1"),
        (0, "A1", "8"),
    ] {
        pair.edit(sheet, r(at), input).unwrap();
        pair.step(true, |workbook| workbook.undo(options()))
            .unwrap();
        pair.step(true, |workbook| workbook.redo(options()))
            .unwrap();
    }
    assert!(fast_set_cell_count() > count);
    pair.assert_equal(true);
}

pub(super) struct Random(pub(super) u32);

impl Random {
    pub(super) fn next(&mut self, bound: u32) -> u32 {
        self.0 ^= self.0 << 13;
        self.0 ^= self.0 >> 17;
        self.0 ^= self.0 << 5;
        self.0 % bound
    }
}

#[test]
fn seeded_edit_cell_lockstep_matches_full_materialization() {
    let mut pair = Lockstep::new(false);
    let mut random = Random(0x5eed_1234);
    let count = fast_set_cell_count();
    for _ in 0..224 {
        let sheet = random.next(3);
        let at = CellRef::new(random.next(6), random.next(5));
        match random.next(10) {
            0 => {
                pair.step(true, |workbook| workbook.undo(options()))
                    .unwrap();
            }
            1 => {
                pair.step(true, |workbook| workbook.redo(options()))
                    .unwrap();
            }
            2 => {
                let patch = StylePatch {
                    bold: Some(random.next(2) == 0),
                    italic: Some(random.next(2) == 0),
                    ..Default::default()
                };
                pair.step(true, |workbook| {
                    workbook.patch_range_style(
                        SheetId(sheet),
                        CellRange::new(at, at),
                        patch.clone(),
                        options(),
                    )
                })
                .unwrap();
            }
            3 => assert!(pair.edit(99, at, "123").is_err()),
            4 => assert!(
                pair.edit(sheet, CellRef::new(MAX_ROWS, at.col), "123")
                    .is_err()
            ),
            5 => {
                let current = current_cell_state(&pair.fast.model, SheetId(sheet), at);
                if current.formula.is_none() && matches!(current.value, CellValue::Error { .. }) {
                    let result = pair
                        .step(true, |workbook| {
                            workbook.apply_ops(
                                vec![Op::SetCell {
                                    sheet: SheetId(sheet),
                                    at,
                                    cell: current.clone(),
                                }],
                                options(),
                            )
                        })
                        .unwrap();
                    assert!(!result.applied);
                    continue;
                }
                let input = match (current.formula, current.value) {
                    (Some(formula), _) => format!("={formula}"),
                    (_, CellValue::Empty) => String::new(),
                    (_, CellValue::Number { value }) => value.to_string(),
                    (_, CellValue::Text { value }) => format!("'{value}"),
                    (_, CellValue::Bool { value }) => value.to_string(),
                    (_, CellValue::Error { .. }) => unreachable!(),
                };
                assert!(!pair.edit(sheet, at, &input).unwrap().applied);
            }
            _ => {
                let input = match random.next(8) {
                    0 => String::new(),
                    1 => "Alpha rich".into(),
                    2 => "true".into(),
                    3 => "=A1+B1".into(),
                    4 => "=Sheet1!A1+Sheet2!A1".into(),
                    5 => "=NOW()+TODAY()".into(),
                    6 => format!("'text {}", random.next(8)),
                    _ => random.next(100).to_string(),
                };
                pair.edit(sheet, at, &input).unwrap();
            }
        }
    }
    assert!(fast_set_cell_count() > count + 20);
    pair.assert_equal(true);
}

#[test]
fn preserved_state_snapshots_copy_on_write() {
    let mut original = PreservedSheetState {
        origins: vec![Some(0), Some(1)],
        shared_string_cells: Arc::new(vec![
            BTreeMap::from([((1, 0), 0), ((1, 1), 1)]),
            BTreeMap::from([((0, 1), 2)]),
        ]),
        axes: vec![Some(xlsx_parse::SheetAxes::default()); 2],
        created: vec![false; 2],
    };
    let expected = original.shared_string_cells.as_ref().clone();
    let snapshot = original.clone();
    assert!(Arc::ptr_eq(
        &original.shared_string_cells,
        &snapshot.shared_string_cells
    ));
    original.shift(
        SheetId(0),
        &Op::InsertRows {
            sheet: SheetId(0),
            at: 1,
            count: 2,
        },
    );
    assert_eq!(
        original.shared_string_cells[0],
        BTreeMap::from([((3, 0), 0), ((3, 1), 1)])
    );
    original.insert(1);
    assert!(original.shared_string_cells[1].is_empty());
    original.remove(0);
    assert_eq!(original.shared_string_cells[1], expected[1]);
    original.resize(3);
    original.forget_shared_strings();
    assert!(original.shared_string_cells.iter().all(BTreeMap::is_empty));
    assert_eq!(*snapshot.shared_string_cells, expected);
}

#[test]
fn move_sheet_preserves_package_identity_and_undo() {
    let mut workbook = Workbook::open(&workbook_bytes()).unwrap();
    let names = workbook.sheet_names();
    let model = workbook.model.clone();
    let origins = workbook.preserved.origins.clone();
    let version = workbook.version();
    assert!(
        workbook
            .apply_ops(vec![Op::MoveSheet { from: 0, to: 2 }], options())
            .unwrap()
            .applied
    );
    assert_eq!(
        workbook.sheet_names(),
        vec![names[1].clone(), names[2].clone(), names[0].clone()]
    );
    assert_eq!(
        workbook.preserved.origins,
        vec![origins[1], origins[2], origins[0]]
    );
    assert_eq!(workbook.model.sheets[2], model.sheets[0]);
    let reopened = Workbook::open(&workbook.save().unwrap()).unwrap();
    assert_eq!(reopened.sheet_names(), workbook.sheet_names());
    assert_eq!(reopened.model.sheets[2], model.sheets[0]);
    assert_ne!(workbook.version(), version);
    workbook.undo(options()).unwrap();
    assert_eq!(workbook.model.sheets, model.sheets);
    assert_eq!(workbook.preserved.origins, origins);
    workbook.redo(options()).unwrap();
    assert_eq!(workbook.model.sheets[2], model.sheets[0]);
    let before = workbook.model.clone();
    assert!(
        workbook
            .apply_ops(
                vec![
                    Op::MoveSheet { from: 2, to: 0 },
                    Op::MoveSheet { from: 7, to: 0 }
                ],
                options()
            )
            .is_err()
    );
    assert_eq!(workbook.model, before);
}

#[test]
fn moved_tables_recover_after_move_undo_and_redo() {
    use xlsx_model::{DefinedName, Table};

    let mut model = WorkbookModel {
        sheets: vec![Sheet::new("A"), Sheet::new("B")],
        ..WorkbookModel::default()
    };
    for (sheet, name) in [(0, "Items"), (1, "Others")] {
        model.tables.push(Table {
            name: name.into(),
            sheet: SheetId(sheet),
            range: CellRange::new(r("A1"), r("A2")),
            header_rows: 1,
            totals_rows: 0,
            columns: vec!["Value".into()],
        });
        model.defined_names.push(DefinedName {
            name: "LocalValue".into(),
            formula: format!("{}!$A$2", model.sheets[sheet as usize].name),
            local_sheet: Some(SheetId(sheet)),
            hidden: false,
        });
        model.sheets[sheet as usize].hyperlinks.push(Hyperlink {
            range: CellRange::new(r("A1"), r("A1")),
            external_target: Some(format!("https://example.org/{sheet}")),
            location: None,
            tooltip: None,
            display: None,
        });
        model.sheets[sheet as usize].row_heights.insert(3, 0.0);
        model.sheets[sheet as usize].col_widths.insert(3, 0.0);
    }
    let mut workbook = Workbook::from_model(model).unwrap();
    for (sheet, at, input) in [
        (0, "A1", "Value"),
        (0, "A2", "10"),
        (1, "A1", "Value"),
        (1, "A2", "99"),
        (1, "B1", "=SUM(Items[Value])+SUM(Others[Value])"),
        (1, "C1", "=SUM(Items[Value])"),
        (1, "D1", "=SUM(Others[Value])"),
    ] {
        workbook
            .edit_cell(SheetId(sheet), r(at), input, options())
            .unwrap();
    }
    workbook
        .apply_ops(vec![Op::MoveSheet { from: 0, to: 1 }], options())
        .unwrap();
    for phase in 0..3 {
        let moved = phase != 1;
        assert_eq!(workbook.model.tables[0].sheet, SheetId(u32::from(moved)));
        assert_eq!(workbook.model.tables[1].sheet, SheetId(u32::from(!moved)));
        let checkpoint = workbook
            .recovery_checkpoint_with_history(1, None, true)
            .unwrap();
        let (mut restored, _, _, _) = Workbook::restore_recovery_checkpoint(&checkpoint).unwrap();
        restored.recalculate_all(options());
        assert_eq!(restored.model.tables, workbook.model.tables);
        assert_eq!(restored.model.defined_names, workbook.model.defined_names);
        assert_eq!(restored.model.sheets, workbook.model.sheets);
        let formula_sheet = SheetId(u32::from(!moved));
        assert_eq!(
            restored
                .model
                .sheet(formula_sheet)
                .unwrap()
                .cell(r("B1"))
                .unwrap()
                .value,
            CellValue::Number { value: 109.0 }
        );
        assert_eq!(
            restored
                .model
                .sheet(formula_sheet)
                .unwrap()
                .cell(r("C1"))
                .unwrap()
                .value,
            CellValue::Number { value: 10.0 }
        );
        assert_eq!(
            restored
                .model
                .sheet(formula_sheet)
                .unwrap()
                .cell(r("D1"))
                .unwrap()
                .value,
            CellValue::Number { value: 99.0 }
        );
        if phase == 0 {
            workbook.undo(options()).unwrap();
            restored.undo(options()).unwrap();
        } else if phase == 1 {
            workbook.redo(options()).unwrap();
            restored.redo(options()).unwrap();
        }
        assert_eq!(restored.model, workbook.model);
    }
}

#[test]
fn mixed_sheet_batch_admits_cells_at_their_current_index() {
    let mut workbook = Workbook::from_model(WorkbookModel {
        sheets: vec![Sheet::new("A"), Sheet::new("B")],
        ..WorkbookModel::default()
    })
    .unwrap();
    workbook
        .apply_ops(
            vec![
                Op::AddSheet {
                    index: 2,
                    name: "C".into(),
                },
                Op::MoveSheet { from: 0, to: 2 },
                Op::RenameSheet {
                    sheet: SheetId(0),
                    name: "Renamed".into(),
                },
                Op::SetCell {
                    sheet: SheetId(2),
                    at: r("A1"),
                    cell: cell_state_for_input_no_eval("42"),
                },
            ],
            options(),
        )
        .unwrap();
    assert_eq!(workbook.sheet_names(), vec!["Renamed", "C", "A"]);
    assert_eq!(
        workbook.model.sheets[2].cell(r("A1")).unwrap().value,
        CellValue::Number { value: 42.0 }
    );
    assert!(workbook.model.sheets[0].cell(r("A1")).is_none());
    assert!(workbook.model.sheets[1].cell(r("A1")).is_none());
}

fn workbook_with_chartsheet() -> Workbook {
    let mut parts = ooxml_opc::unzip_parts(&workbook_bytes()).unwrap();
    for (path, bytes) in &mut parts {
        if path == "xl/_rels/workbook.xml.rels" {
            let xml = String::from_utf8(bytes.clone()).unwrap().replace(
                "relationships/worksheet\" Target=\"worksheets/sheet2.xml",
                "relationships/chartsheet\" Target=\"worksheets/sheet2.xml",
            );
            *bytes = xml.into_bytes();
        } else if path == "xl/worksheets/sheet2.xml" {
            *bytes = br#"<chartsheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>"#.to_vec();
        }
    }
    let workbook = Workbook::open(&ooxml_opc::rezip_parts(&parts).unwrap()).unwrap();
    assert!(
        !workbook
            .source_package
            .as_ref()
            .unwrap()
            .facts()
            .source_sheet_is_worksheet(1)
    );
    workbook
}

#[test]
fn mixed_sheet_batch_refuses_moved_chartsheet_cell_edit() {
    let mut workbook = workbook_with_chartsheet();
    let before = workbook.model.clone();
    let error = workbook
        .apply_ops(
            vec![
                Op::MoveSheet { from: 1, to: 0 },
                Op::SetCell {
                    sheet: SheetId(0),
                    at: r("A1"),
                    cell: cell_state_for_input_no_eval("42"),
                },
            ],
            options(),
        )
        .unwrap_err();
    assert!(
        error.to_string().contains("not an editable worksheet"),
        "{error}"
    );
    assert_eq!(workbook.model, before);
}

#[test]
fn mixed_sheet_batch_admits_worksheet_at_old_chartsheet_index() {
    let mut workbook = workbook_with_chartsheet();
    let name = workbook.model.sheets[0].name.clone();
    workbook
        .apply_ops(
            vec![
                Op::MoveSheet { from: 1, to: 0 },
                Op::SetCell {
                    sheet: SheetId(1),
                    at: r("A1"),
                    cell: cell_state_for_input_no_eval("42"),
                },
            ],
            options(),
        )
        .unwrap();
    assert_eq!(workbook.model.sheets[1].name, name);
    assert_eq!(
        workbook.model.sheets[1].cell(r("A1")).unwrap().value,
        CellValue::Number { value: 42.0 }
    );
}

fn assert_active_sheet_batch(ops: Vec<Op>, before: SheetId, after: SheetId) {
    let mut workbook = Workbook::from_model(WorkbookModel {
        sheets: vec![Sheet::new("A"), Sheet::new("B")],
        ..WorkbookModel::default()
    })
    .unwrap();
    workbook.set_active_sheet(before).unwrap();
    let original_name = workbook.model.sheets[before.0 as usize].name.clone();
    workbook.apply_ops(ops, options()).unwrap();
    assert_eq!(workbook.active_sheet(), after);
    assert_eq!(workbook.model.sheets[after.0 as usize].name, "Renamed");
    workbook.undo(options()).unwrap();
    assert_eq!(workbook.active_sheet(), before);
    assert_eq!(workbook.model.sheets[before.0 as usize].name, original_name);
    workbook.redo(options()).unwrap();
    assert_eq!(workbook.active_sheet(), after);
    assert_eq!(workbook.model.sheets[after.0 as usize].name, "Renamed");
}

#[test]
fn active_sheet_follows_move_and_rename_batch() {
    assert_active_sheet_batch(
        vec![
            Op::MoveSheet { from: 0, to: 1 },
            Op::RenameSheet {
                sheet: SheetId(1),
                name: "Renamed".into(),
            },
        ],
        SheetId(0),
        SheetId(1),
    );
}

#[test]
fn active_sheet_follows_insert_and_rename_batch() {
    assert_active_sheet_batch(
        vec![
            Op::AddSheet {
                index: 0,
                name: "C".into(),
            },
            Op::RenameSheet {
                sheet: SheetId(1),
                name: "Renamed".into(),
            },
        ],
        SheetId(0),
        SheetId(1),
    );
}

#[test]
fn active_sheet_follows_remove_before_and_rename_batch() {
    assert_active_sheet_batch(
        vec![
            Op::RemoveSheet { index: 0 },
            Op::RenameSheet {
                sheet: SheetId(0),
                name: "Renamed".into(),
            },
        ],
        SheetId(1),
        SheetId(0),
    );
}
