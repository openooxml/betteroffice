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
