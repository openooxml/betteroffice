use super::*;
use crate::authority::{
    fast_range_format_count, fast_set_cell_count, full_materializations, local_hydrations,
    with_full_materialization,
};

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
        Self::from_bytes(&bytes, collaborative, rand_seed)
    }

    fn from_bytes(bytes: &[u8], collaborative: bool, rand_seed: Option<u32>) -> Self {
        let fast = open_identical(bytes, collaborative, rand_seed);
        let oracle = open_identical(bytes, collaborative, rand_seed);
        Self::from_workbooks(fast, oracle)
    }

    fn from_model(model: WorkbookModel) -> Self {
        let open = || {
            let mut workbook = Workbook::from_model_collaborative(model.clone(), 73).unwrap();
            workbook.mode = WorkbookMode::Standalone;
            workbook.version_nonce = "singleton-oracle".into();
            workbook
        };
        Self::from_workbooks(open(), open())
    }

    fn from_workbooks(fast: Workbook, oracle: Workbook) -> Self {
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
        let fast_events = std::mem::take(&mut *self.fast_events.lock().unwrap());
        let oracle_events = std::mem::take(&mut *self.oracle_events.lock().unwrap());
        assert_eq!(
            fast_events.len(),
            oracle_events.len(),
            "published update count at step {}",
            self.step
        );
        for (index, (fast, oracle)) in fast_events.iter().zip(&oracle_events).enumerate() {
            assert_eq!(
                fast.update, oracle.update,
                "published update {index} bytes at step {}",
                self.step
            );
            assert_eq!(
                fast.origin, oracle.origin,
                "published update {index} origin at step {}",
                self.step
            );
        }
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
        assert_eq!(
            self.fast.sheet_info().unwrap(),
            self.oracle.sheet_info().unwrap()
        );
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

    fn format_step(
        &mut self,
        apply: impl Fn(&mut Workbook) -> Result<MutationResult>,
    ) -> std::result::Result<MutationResult, String> {
        let result = self.step(true, apply)?;
        let input = self.step.to_string();
        assert!(self.edit(0, r("C1"), &input)?.applied);
        Ok(result)
    }
}

fn assert_range_format_avoids_staging(op: Op) {
    for observed in [false, true] {
        for edited in [false, true] {
            let mut workbook = open_identical(&workbook_bytes(), false, None);
            let _subscription = observed.then(|| workbook.observe_update_v1(|_| {}).unwrap());
            if edited {
                workbook
                    .edit_cell(SheetId(0), r("H9"), "5", options())
                    .unwrap();
            }
            let materializations = full_materializations();
            let hydrations = local_hydrations();
            let fast = fast_range_format_count();
            assert!(
                workbook
                    .apply_ops(vec![op.clone()], options())
                    .unwrap()
                    .applied
            );
            assert_eq!(full_materializations(), materializations);
            assert_eq!(local_hydrations(), hydrations);
            assert_eq!(fast_range_format_count(), fast + 1);
        }
    }
}

#[test]
fn range_format_style_patch_avoids_materialization_and_hydration() {
    assert_range_format_avoids_staging(Op::PatchRangeStyle {
        sheet: SheetId(0),
        range: CellRange::parse_a1("A1:C3").unwrap(),
        patch: StylePatch {
            bold: Some(true),
            fill_color: Some("#123456".into()),
            ..StylePatch::default()
        },
    });
}

#[test]
fn range_format_number_mutation_avoids_materialization_and_hydration() {
    assert_range_format_avoids_staging(Op::SetRangeNumberFormat {
        sheet: SheetId(0),
        range: CellRange::parse_a1("B2:C3").unwrap(),
        format: NumberFormatMutation::Currency,
    });
}

#[test]
fn consecutive_range_formats_and_cell_edit_avoid_materialization_and_hydration() {
    for observed in [false, true] {
        for synthesized in [false, true] {
            let mut workbook = if synthesized {
                Workbook::from_model(plain_style_synthesis_model()).unwrap()
            } else {
                open_identical(&workbook_bytes(), false, None)
            };
            let _subscription = observed.then(|| workbook.observe_update_v1(|_| {}).unwrap());
            let materializations = full_materializations();
            let hydrations = local_hydrations();
            let formats = fast_range_format_count();
            let cells = fast_set_cell_count();
            let (range, patch) = if synthesized {
                (
                    "A1:B1",
                    StylePatch {
                        bold: Some(false),
                        ..StylePatch::default()
                    },
                )
            } else {
                (
                    "A1:C3",
                    StylePatch {
                        bold: Some(true),
                        fill_color: Some("#123456".into()),
                        ..StylePatch::default()
                    },
                )
            };
            workbook
                .patch_range_style(
                    SheetId(0),
                    CellRange::parse_a1(range).unwrap(),
                    patch,
                    options(),
                )
                .unwrap();
            workbook
                .set_range_number_format(
                    SheetId(0),
                    CellRange::parse_a1("B2").unwrap(),
                    NumberFormatMutation::Currency,
                    options(),
                )
                .unwrap();
            workbook
                .edit_cell(SheetId(0), r("H9"), "5", options())
                .unwrap();
            assert_eq!(full_materializations(), materializations);
            assert_eq!(local_hydrations(), hydrations);
            assert_eq!(fast_range_format_count(), formats + 2);
            assert_eq!(fast_set_cell_count(), cells + 1);
        }
    }
}

#[test]
fn consecutive_range_formats_then_cell_edit_match_full_materialization() {
    let mut pair = Lockstep::new(false);
    pair.step(true, |workbook| {
        workbook.patch_range_style(
            SheetId(0),
            CellRange::parse_a1("A1:C3").unwrap(),
            StylePatch {
                bold: Some(true),
                fill_color: Some("#123456".into()),
                ..StylePatch::default()
            },
            options(),
        )
    })
    .unwrap();
    pair.step(true, |workbook| {
        workbook.set_range_number_format(
            SheetId(0),
            CellRange::parse_a1("B2").unwrap(),
            NumberFormatMutation::Currency,
            options(),
        )
    })
    .unwrap();
    pair.edit(0, r("H9"), "5").unwrap();
}

#[test]
fn collaborative_and_compound_range_formats_keep_staging() {
    for collaborative in [false, true] {
        let mut pair = Lockstep::new(collaborative);
        let count = fast_range_format_count();
        let hydrations = local_hydrations();
        pair.step(true, |workbook| {
            let op = Op::SetRangeNumberFormat {
                sheet: SheetId(0),
                range: CellRange::parse_a1("A1:C3").unwrap(),
                format: NumberFormatMutation::Currency,
            };
            let ops = if collaborative {
                vec![op]
            } else {
                vec![op.clone(), op]
            };
            workbook.apply_ops(ops, options())
        })
        .unwrap();
        pair.edit(0, r("C1"), "1").unwrap();
        assert_eq!(fast_range_format_count(), count);
        assert!(local_hydrations() > hydrations);
    }
}

#[test]
fn range_formats_match_full_materialization_and_history() {
    let mut model = Workbook::open(&workbook_bytes()).unwrap().model().clone();
    model.sheets[0]
        .merges
        .push(CellRange::parse_a1("E6:F7").unwrap());
    model.sheets[0].row_styles.insert(6, 1);
    model.sheets[0].col_styles.push(xlsx_model::ColStyle {
        first: 4,
        last: 6,
        xf: 1,
    });
    let bytes = Workbook::from_model(model).unwrap().save().unwrap();
    let mut pair = Lockstep::from_bytes(&bytes, false, None);
    let patches = [
        StylePatch {
            bold: Some(true),
            fill_color: Some("#123456".into()),
            ..StylePatch::default()
        },
        StylePatch {
            bold: Some(true),
            fill_color: Some("#123456".into()),
            ..StylePatch::default()
        },
        StylePatch {
            clear: vec![
                xlsx_ops::StyleProperty::Bold,
                xlsx_ops::StyleProperty::FillColor,
            ],
            ..StylePatch::default()
        },
        StylePatch {
            border: Some(xlsx_ops::BorderPatch {
                preset: Some(BorderPreset::Outer),
                color: Some("#abcdef".into()),
                ..xlsx_ops::BorderPatch::default()
            }),
            ..StylePatch::default()
        },
    ];
    for (range, patch) in ["A1:C3", "E6:G8", "A1:G8", "E6:F7"]
        .into_iter()
        .zip(patches)
    {
        pair.format_step(|workbook| {
            workbook.patch_range_style(
                SheetId(0),
                CellRange::parse_a1(range).unwrap(),
                patch.clone(),
                options(),
            )
        })
        .unwrap();
    }
    for (range, format) in [
        ("B2:C3", NumberFormatMutation::Currency),
        ("E6:G8", NumberFormatMutation::Currency),
        ("A1:G8", NumberFormatMutation::IncreaseDecimal),
        ("A1:G8", NumberFormatMutation::DecreaseDecimal),
        ("A1:G8", NumberFormatMutation::Automatic),
        ("D2:F7", NumberFormatMutation::PlainText),
        ("E6:F7", NumberFormatMutation::Date),
        (
            "E6:F7",
            NumberFormatMutation::Custom {
                pattern: "0.0000".into(),
            },
        ),
    ] {
        pair.format_step(|workbook| {
            workbook.set_range_number_format(
                SheetId(0),
                CellRange::parse_a1(range).unwrap(),
                format.clone(),
                options(),
            )
        })
        .unwrap();
    }
    let depth = pair.fast.history_state().undo_depth;
    for _ in 0..depth {
        pair.step(true, |workbook| workbook.undo(options()))
            .unwrap();
    }
    for _ in 0..depth {
        pair.step(true, |workbook| workbook.redo(options()))
            .unwrap();
    }
}

fn plain_style_synthesis_model() -> WorkbookModel {
    use xlsx_model::{Cell, Font, Xf};

    let mut model = WorkbookModel::default();
    model.styles.fonts = vec![
        Font {
            bold: true,
            ..Font::default()
        },
        Font {
            bold: true,
            italic: true,
            ..Font::default()
        },
    ];
    model.styles.cell_xfs = vec![
        Xf {
            font: Some(0),
            ..Xf::default()
        },
        Xf {
            font: Some(1),
            ..Xf::default()
        },
    ];
    let mut sheet = Sheet::new("Data");
    for (at, style) in [(r("A1"), 0), (r("B1"), 1)] {
        sheet.set_cell(
            at,
            Cell {
                value: CellValue::Number { value: 7.0 },
                style: Some(style),
                ..Cell::default()
            },
        );
    }
    model.sheets.push(sheet);
    model
}

#[test]
fn range_format_then_cell_edit_preserves_synthesized_plain_catalog() {
    for consecutive in [false, true] {
        let mut pair = Lockstep::from_model(plain_style_synthesis_model());
        assert_eq!(pair.fast.model.styles.cell_xfs.len(), 2);
        assert!(pair.fast.model.sheets[0].cell(r("C1")).is_none());
        pair.step(true, |workbook| {
            let result = workbook.patch_range_style(
                SheetId(0),
                CellRange::parse_a1("A1:B1").unwrap(),
                StylePatch {
                    bold: Some(false),
                    ..StylePatch::default()
                },
                options(),
            )?;
            assert_eq!(workbook.model.styles.cell_xfs.len(), 4);
            assert_eq!(
                workbook.model.sheets[0].cell(r("A1")).unwrap().style,
                Some(2)
            );
            assert_eq!(
                workbook.model.sheets[0].cell(r("B1")).unwrap().style,
                Some(3)
            );
            assert_eq!(
                workbook.model.styles.cell_format(Some(2)),
                CellFormat::default()
            );
            assert!(workbook.model.styles.cell_format(Some(3)).font.italic);
            Ok(result)
        })
        .unwrap();
        if consecutive {
            pair.step(true, |workbook| {
                workbook.set_range_number_format(
                    SheetId(0),
                    CellRange::parse_a1("B2").unwrap(),
                    NumberFormatMutation::Currency,
                    options(),
                )
            })
            .unwrap();
        }
        pair.edit(0, r("C1"), "1").unwrap();
        assert_eq!(pair.fast.cell(SheetId(0), r("C1")).unwrap().input, "1");
    }
}

#[test]
fn range_format_then_cell_edit_preserves_projected_style_alias() {
    for mode in 0..3 {
        let mut pair = Lockstep::from_model(plain_style_synthesis_model());
        pair.step(true, |workbook| {
            workbook.patch_range_style(
                SheetId(0),
                CellRange::parse_a1("A1:B1").unwrap(),
                StylePatch {
                    bold: Some(false),
                    ..StylePatch::default()
                },
                options(),
            )
        })
        .unwrap();
        pair.step(true, |workbook| match mode {
            0 => workbook.edit_cell(SheetId(0), r("A1"), "1", options()),
            1 => workbook.edit_cells(
                SheetId(0),
                &[CellInput {
                    cell: r("A1"),
                    input: "1".into(),
                }],
                options(),
            ),
            _ => workbook.apply_ops(
                vec![Op::SetCell {
                    sheet: SheetId(0),
                    at: r("A1"),
                    cell: CellState {
                        value: CellValue::Number { value: 1.0 },
                        style: current_cell_state(&workbook.model, SheetId(0), r("A1")).style,
                        ..CellState::default()
                    },
                }],
                options(),
            ),
        })
        .unwrap();
        pair.edit(0, r("C1"), "2").unwrap();
        assert_eq!(pair.fast.cell(SheetId(0), r("A1")).unwrap().input, "1");
    }
}

#[test]
fn range_format_then_staged_cell_edit_preserves_synthesized_plain_catalog() {
    for mode in 0..3 {
        let mut pair = Lockstep::from_model(plain_style_synthesis_model());
        pair.step(true, |workbook| {
            workbook.patch_range_style(
                SheetId(0),
                CellRange::parse_a1("A1:B1").unwrap(),
                StylePatch {
                    bold: Some(false),
                    ..StylePatch::default()
                },
                options(),
            )
        })
        .unwrap();
        pair.step(true, |workbook| match mode {
            0 => workbook.edit_cells(
                SheetId(0),
                &[CellInput {
                    cell: r("C1"),
                    input: "1".into(),
                }],
                options(),
            ),
            1 => workbook.edit_cells(
                SheetId(0),
                &[
                    CellInput {
                        cell: r("C1"),
                        input: "1".into(),
                    },
                    CellInput {
                        cell: r("D1"),
                        input: "2".into(),
                    },
                ],
                options(),
            ),
            _ => workbook.apply_ops(
                vec![Op::SetCell {
                    sheet: SheetId(0),
                    at: r("C1"),
                    cell: CellState {
                        value: CellValue::Number { value: 1.0 },
                        ..CellState::default()
                    },
                }],
                options(),
            ),
        })
        .unwrap();
        pair.edit(0, r("H9"), "3").unwrap();
        assert_eq!(pair.fast.cell(SheetId(0), r("C1")).unwrap().input, "1");
    }
}

#[derive(Debug)]
enum RangeFormatSequenceEdit {
    Operation(Box<Op>),
    Cell(SheetId, CellRef, String),
    Cells(SheetId, Vec<CellInput>),
    Undo,
    Redo,
}

impl RangeFormatSequenceEdit {
    fn is_fast_format(&self) -> bool {
        matches!(
            self,
            Self::Operation(op)
                if matches!(op.as_ref(), Op::PatchRangeStyle { .. } | Op::SetRangeNumberFormat { .. })
        )
    }

    fn requires_projection_refresh(&self) -> bool {
        !matches!(self, Self::Cell(..)) && !self.is_fast_format()
    }

    fn apply(&self, workbook: &mut Workbook) -> Result<MutationResult> {
        match self {
            Self::Operation(op) => match op.as_ref() {
                Op::PatchRangeStyle {
                    sheet,
                    range,
                    patch,
                } => workbook.patch_range_style(*sheet, *range, patch.clone(), options()),
                Op::SetRangeNumberFormat {
                    sheet,
                    range,
                    format,
                } => workbook.set_range_number_format(*sheet, *range, format.clone(), options()),
                op => workbook.apply_ops(vec![op.clone()], options()),
            },
            Self::Cell(sheet, at, input) => workbook.edit_cell(*sheet, *at, input, options()),
            Self::Cells(sheet, inputs) => workbook.edit_cells(*sheet, inputs, options()),
            Self::Undo => workbook.undo(options()),
            Self::Redo => workbook.redo(options()),
        }
    }
}

fn range_format_sequence_model(seed: u32) -> WorkbookModel {
    use xlsx_model::{Cell, ColStyle};

    let mut random = Random(seed);
    let mut model = plain_style_synthesis_model();
    model.sheets.push(Sheet::new("Other"));
    for (index, sheet) in model.sheets.iter_mut().enumerate() {
        for row in 0..8 {
            for col in 0..8 {
                if (row + col) % 3 == 0 && !(index == 0 && row == 0 && col < 2) {
                    sheet.set_cell(
                        CellRef::new(row, col),
                        Cell {
                            value: CellValue::Number {
                                value: f64::from(random.next(100)),
                            },
                            style: Some(random.next(2)),
                            ..Cell::default()
                        },
                    );
                }
            }
        }
        sheet.set_cell(
            r("C2"),
            Cell {
                value: CellValue::Number { value: 0.0 },
                formula: Some("A1+B1".into()),
                style: Some(1),
            },
        );
        sheet.col_styles.push(ColStyle {
            first: 2,
            last: 3,
            xf: 1,
        });
        sheet.row_styles.insert(4, 0);
    }
    model.sheets[0]
        .merges
        .push(CellRange::parse_a1("E7:F8").unwrap());
    model
}

fn random_range_style(random: &mut Random) -> StylePatch {
    let color = ["#123456", "#abcdef", "#ff8800"][random.next(3) as usize].to_string();
    match random.next(7) {
        0 => StylePatch {
            bold: Some(random.next(2) == 0),
            ..StylePatch::default()
        },
        1 => StylePatch {
            italic: Some(random.next(2) == 0),
            strikethrough: Some(random.next(2) == 0),
            ..StylePatch::default()
        },
        2 => StylePatch {
            text_color: Some(color),
            fill_color: Some("#aabbcc".into()),
            ..StylePatch::default()
        },
        3 => StylePatch {
            font_family: Some(["Arial", "Calibri", "Courier New"][random.next(3) as usize].into()),
            font_size: Some([9.0, 11.0, 18.0][random.next(3) as usize]),
            ..StylePatch::default()
        },
        4 => StylePatch {
            horizontal_alignment: Some(
                [
                    HorizontalAlignment::Left,
                    HorizontalAlignment::Center,
                    HorizontalAlignment::Right,
                ][random.next(3) as usize],
            ),
            vertical_alignment: Some(
                [
                    VerticalAlignment::Top,
                    VerticalAlignment::Middle,
                    VerticalAlignment::Bottom,
                ][random.next(3) as usize],
            ),
            text_wrapping: Some(
                [
                    TextWrapping::Overflow,
                    TextWrapping::Wrap,
                    TextWrapping::Clip,
                ][random.next(3) as usize],
            ),
            ..StylePatch::default()
        },
        5 => StylePatch {
            border: Some(xlsx_ops::BorderPatch {
                preset: Some(
                    [BorderPreset::All, BorderPreset::Outer, BorderPreset::Inner]
                        [random.next(3) as usize],
                ),
                style: Some(
                    [
                        BorderLineStyle::Solid,
                        BorderLineStyle::Dashed,
                        BorderLineStyle::Double,
                    ][random.next(3) as usize],
                ),
                color: Some(color),
            }),
            ..StylePatch::default()
        },
        _ => StylePatch {
            clear: vec![
                xlsx_ops::StyleProperty::Bold,
                xlsx_ops::StyleProperty::FillColor,
            ],
            ..StylePatch::default()
        },
    }
}

fn sequence_number_format(kind: u32, random: &mut Random) -> NumberFormatMutation {
    match kind {
        0 => NumberFormatMutation::Automatic,
        1 => NumberFormatMutation::PlainText,
        2 => NumberFormatMutation::Number,
        3 => NumberFormatMutation::Percent,
        4 => NumberFormatMutation::Scientific,
        5 => NumberFormatMutation::Currency,
        6 => NumberFormatMutation::Date,
        7 => NumberFormatMutation::Time,
        8 => NumberFormatMutation::Custom {
            pattern: format!("0.{}", "0".repeat((random.next(4) + 1) as usize)),
        },
        9 => NumberFormatMutation::IncreaseDecimal,
        _ => NumberFormatMutation::DecreaseDecimal,
    }
}

fn random_range_format_edit(
    random: &mut Random,
    workbook: &Workbook,
    kind: u32,
    number_kind: u32,
) -> RangeFormatSequenceEdit {
    let sheet = SheetId(random.next(2));
    let at = CellRef::new(random.next(6), random.next(6));
    let range = CellRange::new(
        at,
        CellRef::new(at.row + random.next(3), at.col + random.next(3)),
    );
    let op = match kind {
        0 => Op::PatchRangeStyle {
            sheet,
            range,
            patch: random_range_style(random),
        },
        1 => Op::SetRangeNumberFormat {
            sheet,
            range,
            format: sequence_number_format(number_kind, random),
        },
        2 => {
            let input = match random.next(4) {
                0 => random.next(100).to_string(),
                1 => format!("-{}.25", random.next(10)),
                2 => "true".into(),
                _ => format!("'text {}", random.next(10)),
            };
            return RangeFormatSequenceEdit::Cell(sheet, at, input);
        }
        3 => {
            let input = [
                "=A1+B1",
                "=SUM(A1:C3)",
                "=IF(A1>0,B1,0)",
                "=Other!A1+Data!B1",
                "=1/0",
            ][random.next(5) as usize];
            return RangeFormatSequenceEdit::Cell(sheet, at, input.into());
        }
        4 => Op::ApplyRangeFormat {
            sheet,
            range,
            format: CapturedFormat {
                rows: 1,
                columns: 1,
                formats: vec![CellFormat::default()],
                kept: Vec::new(),
                source_styles: Vec::new(),
                source_formats: Vec::new(),
            },
        },
        5 => return RangeFormatSequenceEdit::Cell(sheet, at, String::new()),
        6 => Op::InsertRows {
            sheet,
            at: at.row,
            count: 1,
        },
        7 => Op::DeleteRows {
            sheet,
            at: at.row,
            count: 1,
        },
        8 => Op::InsertCols {
            sheet,
            at: at.col,
            count: 1,
        },
        9 => Op::DeleteCols {
            sheet,
            at: at.col,
            count: 1,
        },
        10 => Op::MergeCells {
            sheet,
            range: if range.start == range.end {
                CellRange::new(at, CellRef::new(at.row, at.col + 1))
            } else {
                range
            },
        },
        11 => {
            let merges = &workbook.model.sheets[sheet.0 as usize].merges;
            let range = if merges.is_empty() {
                range
            } else {
                merges[random.next(merges.len() as u32) as usize]
            };
            Op::UnmergeCells { sheet, range }
        }
        12 => return RangeFormatSequenceEdit::Undo,
        13 => return RangeFormatSequenceEdit::Redo,
        14 | 15 => {
            let mut inputs = vec![CellInput {
                cell: at,
                input: (10_000 + random.next(100)).to_string(),
            }];
            if kind == 15 {
                inputs.push(CellInput {
                    cell: CellRef::new(at.row, at.col + 1),
                    input: "=A1+B1".into(),
                });
            }
            return RangeFormatSequenceEdit::Cells(sheet, inputs);
        }
        _ => Op::SetCell {
            sheet,
            at,
            cell: CellState {
                value: CellValue::Number {
                    value: f64::from(20_000 + random.next(100)),
                },
                style: current_cell_state(&workbook.model, sheet, at).style,
                ..CellState::default()
            },
        },
    };
    RangeFormatSequenceEdit::Operation(Box::new(op))
}

fn run_seeded_range_format_sequence(sequence: u32, check_fast_path: bool) -> [bool; 11] {
    let mut number_formats_used = [false; 11];
    let seed = 0x5246_0001 + sequence;
    let mut random = Random(seed);
    let mut pair = Lockstep::from_model(range_format_sequence_model(seed));
    pair.step = sequence as usize * 30;
    let mut order = [0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16];
    for index in (1..order.len()).rev() {
        order.swap(index, random.next((index + 1) as u32) as usize);
    }
    let mut first_number_format = true;
    let mut operations = Vec::new();
    let result = std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| {
        for step in 0..30 {
            let kind = if step < order.len() {
                order[step]
            } else {
                random.next(17)
            };
            let number_kind = if kind == 1 {
                let kind = if first_number_format {
                    sequence % 11
                } else {
                    random.next(11)
                };
                first_number_format = false;
                number_formats_used[kind as usize] = true;
                kind
            } else {
                0
            };
            let edit = random_range_format_edit(&mut random, &pair.fast, kind, number_kind);
            operations.push(format!("{step}: {edit:?}"));
            pair.step(false, |workbook| {
                let materializations = full_materializations();
                let hydrations = local_hydrations();
                let result = edit.apply(workbook);
                if check_fast_path
                    && edit.is_fast_format()
                    && !crate::authority::force_full_materialization()
                {
                    assert_eq!(
                        full_materializations(),
                        materializations,
                        "seed {seed:#x}, step {step}, {edit:?}"
                    );
                    assert_eq!(
                        local_hydrations(),
                        hydrations,
                        "seed {seed:#x}, step {step}, {edit:?}"
                    );
                }
                result
            })
            .unwrap_or_else(|error| panic!("seed {seed:#x}, step {step}, {edit:?}: {error}"));
            if check_fast_path && edit.requires_projection_refresh() {
                assert_eq!(
                    pair.fast.authority.materialize().unwrap(),
                    pair.oracle.authority.materialize().unwrap()
                );
            }
        }
        pair.assert_equal(true);
    }));
    if let Err(failure) = result {
        let message = failure
            .downcast_ref::<String>()
            .map(String::as_str)
            .or_else(|| failure.downcast_ref::<&str>().copied())
            .unwrap_or("non-string panic");
        panic!(
            "seed {seed:#x}, step {}, global step {}: {message}\noperations:\n{}",
            operations.len().saturating_sub(1),
            pair.step,
            operations.join("\n")
        );
    }
    number_formats_used
}

fn run_seeded_range_format_sequences(check_fast_path: bool) {
    let mut number_formats_used = [false; 11];
    for sequence in 0..50 {
        for (used, observed) in number_formats_used
            .iter_mut()
            .zip(run_seeded_range_format_sequence(sequence, check_fast_path))
        {
            *used |= observed;
        }
    }
    assert!(number_formats_used.into_iter().all(|used| used));
}

#[test]
fn range_format_seed_52460004_step_117_regression() {
    run_seeded_range_format_sequence(3, true);
}

#[test]
fn range_format_seed_52460029_step_1225_regression() {
    run_seeded_range_format_sequence(40, false);
}

#[test]
fn seeded_range_format_sequences_match_full_materialization() {
    run_seeded_range_format_sequences(false);
}

#[test]
fn seeded_range_format_sequences_avoid_materialization_and_hydration() {
    run_seeded_range_format_sequences(true);
}

#[test]
fn invalid_range_format_leaves_state_and_history_unchanged() {
    let mut pair = Lockstep::new(false);
    pair.edit(0, r("A1"), "17").unwrap();
    pair.step(true, |workbook| workbook.undo(options()))
        .unwrap();
    let model = pair.fast.model().clone();
    let state = pair.fast.encode_state_as_update_v1();
    let history = pair.fast.history_state();
    let version = pair.fast.version();
    let info = pair.fast.sheet_info().unwrap();
    let bytes = pair.fast.save().unwrap();
    let count = fast_range_format_count();
    for patch in [
        StylePatch {
            font_size: Some(0.0),
            ..StylePatch::default()
        },
        StylePatch {
            font_family: Some("f".repeat(70_000)),
            ..StylePatch::default()
        },
    ] {
        pair.fast.authority.materialize().unwrap();
        pair.oracle.authority.materialize().unwrap();
        assert!(
            pair.step(true, |workbook| {
                workbook.patch_range_style(
                    SheetId(0),
                    CellRange::parse_a1("A1:C3").unwrap(),
                    patch.clone(),
                    options(),
                )
            })
            .is_err()
        );
        assert_eq!(pair.fast.model(), &model);
        assert_eq!(pair.fast.encode_state_as_update_v1(), state);
        assert_eq!(pair.fast.history_state(), history);
        assert_eq!(pair.fast.version(), version);
        assert_eq!(pair.fast.sheet_info().unwrap(), info);
        assert_eq!(pair.fast.save().unwrap(), bytes);
        assert_eq!(fast_range_format_count(), count);
    }
    pair.step(true, |workbook| workbook.redo(options()))
        .unwrap();
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
