use betteroffice_xlsx::{
    CalculationOptions, CellRange, CellRef, NumberFormatMutation, Op, SheetId, SnapshotBudget,
    Workbook, WorkbookSnapshotBuilder, WorkbookSnapshotEncoder,
};
use ooxml_drawingml::chart::ChartSpace;

const FIXTURE: &[u8] = include_bytes!("../../../packages/xlsx/test-fixtures/charts.xlsx");
const BAR: &str = "xl/charts/chart1.xml";
const LINE: &str = "xl/charts/chart2.xml";

fn part<'a>(parts: &'a [(String, Vec<u8>)], path: &str) -> &'a [u8] {
    &parts.iter().find(|(name, _)| name == path).unwrap().1
}

fn rewrite(parts: &mut [(String, Vec<u8>)], path: &str, edit: impl FnOnce(String) -> String) {
    let (_, bytes) = parts.iter_mut().find(|(name, _)| name == path).unwrap();
    *bytes = edit(String::from_utf8(bytes.clone()).unwrap()).into_bytes();
}

fn edit(workbook: &mut Workbook, address: &str, input: &str) {
    workbook
        .edit_cell(
            SheetId(0),
            CellRef::parse_a1(address).unwrap(),
            input,
            CalculationOptions::default(),
        )
        .unwrap();
}

fn number_format(workbook: &mut Workbook, address: &str, format: NumberFormatMutation) {
    let at = CellRef::parse_a1(address).unwrap();
    workbook
        .apply_ops(
            vec![Op::SetRangeNumberFormat {
                sheet: SheetId(0),
                range: CellRange::new(at, at),
                format,
            }],
            CalculationOptions::default(),
        )
        .unwrap();
}

fn saved_parts(workbook: &Workbook) -> Vec<(String, Vec<u8>)> {
    ooxml_opc::unzip_parts(&workbook.save().unwrap()).unwrap()
}

fn cached_chart(parts: &[(String, Vec<u8>)], path: &str) -> ChartSpace {
    xlsx_parse::chart_space(part(parts, path), &Default::default()).unwrap()
}

fn series_bytes(bytes: &[u8], index: usize) -> &str {
    let xml = std::str::from_utf8(bytes).unwrap();
    let start = xml.match_indices("<c:ser>").nth(index).unwrap().0;
    let end = start + xml[start..].find("</c:ser>").unwrap() + "</c:ser>".len();
    &xml[start..end]
}

fn authored_parts() -> Vec<(String, Vec<u8>)> {
    let mut parts = ooxml_opc::unzip_parts(FIXTURE).unwrap();
    for (path, bytes) in &mut parts {
        if path.starts_with("xl/charts/") && path.ends_with(".xml") {
            *bytes = String::from_utf8(bytes.clone())
                .unwrap()
                .replace("<c:v>12</c:v>", "<c:v>12.0000</c:v>")
                .replace("<c:v>8</c:v>", "<c:v>8.0000</c:v>")
                .replace("<c:ptCount val=\"4\"/>", "<c:ptCount val='04' />")
                .replace(
                    "<c:numCache>",
                    "<c:numCache><c:formatCode>0.00</c:formatCode>",
                )
                .into_bytes();
        }
    }
    parts
}

#[test]
fn no_edit_save_preserves_the_chart_package() {
    let source = ooxml_opc::rezip_parts(&authored_parts()).unwrap();
    assert_eq!(Workbook::open(&source).unwrap().save().unwrap(), source);
}

#[test]
fn value_edit_updates_saved_numeric_cache() {
    let mut workbook = Workbook::open(FIXTURE).unwrap();
    edit(&mut workbook, "B2", "96");
    let parts = saved_parts(&workbook);
    let chart = cached_chart(&parts, BAR);
    assert_eq!(
        chart.plot_groups[0].series[0].values,
        [96.0, 18.0, 15.0, 24.0]
    );
}

#[test]
fn category_edit_updates_saved_string_cache() {
    let mut workbook = Workbook::open(FIXTURE).unwrap();
    edit(&mut workbook, "A2", "First & best");
    let parts = saved_parts(&workbook);
    let chart = cached_chart(&parts, BAR);
    for series in &chart.plot_groups[0].series {
        assert_eq!(series.categories, ["First & best", "Q2", "Q3", "Q4"]);
    }
}

#[test]
fn unrelated_edit_preserves_authored_chart_parts_until_a_source_changes() {
    let before = authored_parts();
    let source = ooxml_opc::rezip_parts(&before).unwrap();
    let mut workbook = Workbook::open(&source).unwrap();
    edit(&mut workbook, "Z20", "77");
    let after = saved_parts(&workbook);
    for (path, bytes) in &before {
        if path.starts_with("xl/charts/") {
            assert_eq!(part(&after, path), bytes, "{path}");
        }
    }
    edit(&mut workbook, "B2", "96");
    let after = saved_parts(&workbook);
    assert_eq!(
        cached_chart(&after, BAR).plot_groups[0].series[0].values[0],
        96.0
    );
    edit(&mut workbook, "B2", "12");
    let restored = saved_parts(&workbook);
    for (path, bytes) in &before {
        if path.starts_with("xl/charts/") {
            assert_eq!(part(&restored, path), bytes, "{path}");
        }
    }
}

#[test]
fn source_edit_preserves_the_other_series_authored_bytes() {
    let before = authored_parts();
    let mut workbook = Workbook::open(&ooxml_opc::rezip_parts(&before).unwrap()).unwrap();
    edit(&mut workbook, "B2", "96");
    let after = saved_parts(&workbook);
    assert_eq!(
        cached_chart(&after, BAR).plot_groups[0].series[0].values[0],
        96.0
    );
    assert_eq!(
        series_bytes(part(&after, BAR), 1),
        series_bytes(part(&before, BAR), 1)
    );
    assert!(series_bytes(part(&after, BAR), 0).contains("<c:formatCode>0.00</c:formatCode>"));
    assert_eq!(
        part(&after, "xl/charts/chart4.xml"),
        part(&before, "xl/charts/chart4.xml")
    );
}

#[test]
fn row_insert_moves_chart_references_and_refreshes_their_caches_on_later_edits() {
    let mut workbook = Workbook::open(FIXTURE).unwrap();
    edit(&mut workbook, "B2", "96");
    workbook
        .apply_ops(
            vec![Op::InsertRows {
                sheet: SheetId(0),
                at: 0,
                count: 1,
            }],
            CalculationOptions::default(),
        )
        .unwrap();
    let saved = workbook.save().unwrap();
    let parts = ooxml_opc::unzip_parts(&saved).unwrap();
    let xml = std::str::from_utf8(part(&parts, BAR)).unwrap();
    assert!(xml.contains("<c:f>Charts!$B$3:$B$6</c:f>"));
    assert!(xml.contains("<c:f>Charts!$A$3:$A$6</c:f>"));
    assert_eq!(
        cached_chart(&parts, BAR).plot_groups[0].series[0].values,
        [96.0, 18.0, 15.0, 24.0]
    );
    let mut reopened = Workbook::open(&saved).unwrap();
    edit(&mut reopened, "B3", "144");
    let parts = saved_parts(&reopened);
    assert_eq!(
        cached_chart(&parts, BAR).plot_groups[0].series[0].values,
        [144.0, 18.0, 15.0, 24.0]
    );
}

#[test]
fn external_chart_stays_authored_while_local_charts_refresh() {
    let mut before = authored_parts();
    rewrite(&mut before, LINE, |xml| {
        xml.replace("Charts!$B$2:$B$5", "[1]Charts!$B$2:$B$5")
    });
    let mut workbook = Workbook::open(&ooxml_opc::rezip_parts(&before).unwrap()).unwrap();
    edit(&mut workbook, "B2", "96");
    edit(&mut workbook, "A2", "First");
    let after = saved_parts(&workbook);
    assert_eq!(
        cached_chart(&after, BAR).plot_groups[0].series[0].values[0],
        96.0
    );
    assert_eq!(part(&after, LINE), part(&before, LINE));
}

#[test]
fn declined_chart_parts_stay_authored_while_local_charts_refresh() {
    for (from, to) in [
        ("Charts!$B$2:$B$5", "Charts!$B$2:$B$10000"),
        ("</c:chartSpace>", "<c:externalData/></c:chartSpace>"),
        ("</c:chartSpace>", "<c:pivotSource/></c:chartSpace>"),
        ("</c:chartSpace>", "<c:sqref/></c:chartSpace>"),
        (
            "<c:chartSpace ",
            "<c:chartSpace xmlns:cx=\"http://schemas.microsoft.com/office/drawing/2014/chartex\" ",
        ),
    ] {
        let mut before = authored_parts();
        rewrite(&mut before, LINE, |xml| xml.replace(from, to));
        let mut workbook = Workbook::open(&ooxml_opc::rezip_parts(&before).unwrap()).unwrap();
        edit(&mut workbook, "B2", "96");
        edit(&mut workbook, "A2", "First");
        let after = saved_parts(&workbook);
        assert_eq!(
            cached_chart(&after, BAR).plot_groups[0].series[0].values[0],
            96.0
        );
        assert_eq!(part(&after, LINE), part(&before, LINE), "{to}");
    }
}

#[test]
fn formula_result_edit_updates_saved_numeric_cache() {
    let mut before = ooxml_opc::unzip_parts(FIXTURE).unwrap();
    rewrite(&mut before, "xl/worksheets/sheet1.xml", |xml| {
        xml.replace(
            "<c r=\"B2\"><v>12</v></c>",
            "<c r=\"B2\"><f>D2*2</f><v>12</v></c>",
        )
        .replace(
            "<c r=\"C2\"><v>8</v></c>",
            "<c r=\"C2\"><v>8</v></c><c r=\"D2\"><v>6</v></c>",
        )
    });
    let mut workbook = Workbook::open(&ooxml_opc::rezip_parts(&before).unwrap()).unwrap();
    edit(&mut workbook, "D2", "48");
    let after = saved_parts(&workbook);
    assert_eq!(
        cached_chart(&after, BAR).plot_groups[0].series[0].values[0],
        96.0
    );
}

#[test]
fn source_number_format_edit_refreshes_only_its_series() {
    for format in [
        NumberFormatMutation::Custom {
            pattern: "0.00%".to_owned(),
        },
        NumberFormatMutation::Percent,
    ] {
        let before = authored_parts();
        let mut workbook = Workbook::open(&ooxml_opc::rezip_parts(&before).unwrap()).unwrap();
        number_format(&mut workbook, "B3", format);
        let after = saved_parts(&workbook);
        let series = series_bytes(part(&after, BAR), 0);
        assert!(
            series.contains("<c:formatCode>0.00%</c:formatCode>"),
            "{series}"
        );
        assert!(!series.contains("<c:formatCode>0.00</c:formatCode>"));
        assert_eq!(
            cached_chart(&after, BAR).plot_groups[0].series[0].values,
            [12.0, 18.0, 15.0, 24.0]
        );
        assert_eq!(
            series_bytes(part(&after, BAR), 1),
            series_bytes(part(&before, BAR), 1)
        );
    }
}

#[test]
fn equivalent_builtin_and_custom_formats_preserve_authored_chart_bytes() {
    let mut workbook = Workbook::open(FIXTURE).unwrap();
    number_format(
        &mut workbook,
        "B2",
        NumberFormatMutation::Custom {
            pattern: "0.00%".into(),
        },
    );
    let mut before = saved_parts(&workbook);
    rewrite(&mut before, BAR, |xml| {
        xml.replace("<c:v>12</c:v>", "<c:v>12.0000</c:v>")
    });
    let mut workbook = Workbook::open(&ooxml_opc::rezip_parts(&before).unwrap()).unwrap();
    let at = CellRef::new(1, 1);
    assert_eq!(
        workbook
            .model()
            .styles
            .resolved_format(workbook.model().sheets[0].effective_style(at))
            .number_format,
        xlsx_model::styles::FormatCode::Custom("0.00%")
    );
    number_format(&mut workbook, "B2", NumberFormatMutation::Percent);
    assert_eq!(
        workbook
            .model()
            .styles
            .resolved_format(workbook.model().sheets[0].effective_style(at))
            .number_format,
        xlsx_model::styles::FormatCode::Builtin(10)
    );
    let after = saved_parts(&workbook);
    for (path, bytes) in &before {
        if path.starts_with("xl/charts/") {
            assert_eq!(part(&after, path), bytes, "{path}");
        }
    }
}

#[test]
fn numeric_string_categories_use_cell_formats_and_the_workbook_date_system() {
    for (date1904, serial) in [(false, "45658"), (true, "44196")] {
        for moved in [false, true] {
            let mut before = authored_parts();
            if date1904 {
                rewrite(&mut before, "xl/workbook.xml", |xml| {
                    xml.replace("<sheets>", "<workbookPr date1904=\"1\"/><sheets>")
                });
            }
            let mut workbook = Workbook::open(&ooxml_opc::rezip_parts(&before).unwrap()).unwrap();
            for (address, value, pattern) in [
                ("A2", serial, "yyyy-mm-dd"),
                ("A3", "0.5", "0.00%"),
                ("A4", "1234.5", "#,##0.00"),
            ] {
                edit(&mut workbook, address, value);
                number_format(
                    &mut workbook,
                    address,
                    NumberFormatMutation::Custom {
                        pattern: pattern.into(),
                    },
                );
            }
            if moved {
                workbook
                    .apply_ops(
                        vec![Op::InsertRows {
                            sheet: SheetId(0),
                            at: 0,
                            count: 1,
                        }],
                        CalculationOptions::default(),
                    )
                    .unwrap();
            }
            let after = saved_parts(&workbook);
            for series in &cached_chart(&after, BAR).plot_groups[0].series {
                assert_eq!(
                    series.categories,
                    ["2025-01-01", "50.00%", "1,234.50", "Q4"],
                    "date1904={date1904}, moved={moved}"
                );
            }
            assert_eq!(
                cached_chart(&after, BAR).plot_groups[0].series[0].values,
                [12.0, 18.0, 15.0, 24.0]
            );
            let xml = std::str::from_utf8(part(&after, BAR)).unwrap();
            assert!(xml.contains(if moved {
                "Charts!$A$3:$A$6"
            } else {
                "Charts!$A$2:$A$5"
            }));
        }
    }
}

#[test]
fn snapshot_hydrated_save_refreshes_edits_after_adoption() {
    let before = authored_parts();
    let source = ooxml_opc::rezip_parts(&before).unwrap();
    let workbook = Workbook::open(&source).unwrap();
    let budget = SnapshotBudget::new(8, 65536).unwrap();
    let mut encoder = WorkbookSnapshotEncoder::new(&workbook, None, budget).unwrap();
    let mut builder = WorkbookSnapshotBuilder::new();
    while let Some(chunk) = encoder.next(&workbook).unwrap() {
        builder.push(&chunk).unwrap();
    }
    while !builder.advance(budget).unwrap().is_ready() {}
    let (mut restored, _) = builder.finish().unwrap().into_parts();
    edit(&mut restored, "B2", "96");
    for subsequent_edit in [false, true] {
        if subsequent_edit {
            edit(&mut restored, "B3", "144");
        }
        let after = saved_parts(&restored);
        assert_eq!(
            cached_chart(&after, BAR).plot_groups[0].series[0].values,
            [96.0, if subsequent_edit { 144.0 } else { 18.0 }, 15.0, 24.0]
        );
        assert_eq!(
            series_bytes(part(&after, BAR), 1),
            series_bytes(part(&before, BAR), 1)
        );
        assert_eq!(
            part(&after, "xl/charts/chart4.xml"),
            part(&before, "xl/charts/chart4.xml")
        );
    }
}

#[test]
fn a_collaborative_peer_saves_refreshed_caches_after_adopting_an_update() {
    let before = authored_parts();
    let source = ooxml_opc::rezip_parts(&before).unwrap();
    let mut writer = Workbook::open_collaborative(&source, 101).unwrap();
    let mut peer = Workbook::open_collaborative(&source, 102).unwrap();
    edit(&mut writer, "B2", "96");
    assert!(
        peer.apply_update_v1(
            &writer.encode_state_as_update_v1(),
            CalculationOptions::default()
        )
        .unwrap()
        .applied
    );
    edit(&mut peer, "B3", "144");
    let after = saved_parts(&peer);
    assert_eq!(
        cached_chart(&after, BAR).plot_groups[0].series[0].values,
        [96.0, 144.0, 15.0, 24.0]
    );
    assert_eq!(
        series_bytes(part(&after, BAR), 1),
        series_bytes(part(&before, BAR), 1)
    );
}

#[test]
fn undoing_a_source_edit_preserves_authored_chart_bytes() {
    let before = authored_parts();
    let mut workbook = Workbook::open(&ooxml_opc::rezip_parts(&before).unwrap()).unwrap();
    edit(&mut workbook, "Z20", "unrelated");
    edit(&mut workbook, "B2", "96");
    workbook.undo(CalculationOptions::default()).unwrap();
    let after = saved_parts(&workbook);
    for (path, bytes) in &before {
        if path.starts_with("xl/charts/") {
            assert_eq!(part(&after, path), bytes, "{path}");
        }
    }
}

#[test]
fn a_source_sheet_rename_keeps_formatted_string_caches() {
    let mut workbook = Workbook::open(FIXTURE).unwrap();
    edit(&mut workbook, "A2", "0.5");
    number_format(&mut workbook, "A2", NumberFormatMutation::Percent);
    workbook
        .apply_ops(
            vec![Op::RenameSheet {
                sheet: SheetId(0),
                name: "Renamed".into(),
            }],
            CalculationOptions::default(),
        )
        .unwrap();
    let after = saved_parts(&workbook);
    assert_eq!(
        cached_chart(&after, BAR).plot_groups[0].series[0].categories[0],
        "50.00%"
    );
    assert!(
        std::str::from_utf8(part(&after, BAR))
            .unwrap()
            .contains("Renamed!$A$2:$A$5")
    );
}
