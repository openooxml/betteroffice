use serde_json::{Value, json};

use super::edit_tests::{open_identical, options, r, workbook_bytes};
use super::*;

fn source(attributes: &str, children: &str, protected: bool) -> (Vec<u8>, String) {
    let xf = format!(
        r#"<xf numFmtId="2" fontId="0" fillId="0" borderId="0" {attributes}>{children}</xf>"#
    );
    let styles = format!(
        r#"<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<fonts count="1"><font><name val="Calibri"/><sz val="11"/></font></fonts>
<fills count="1"><fill><patternFill patternType="none"/></fill></fills>
<borders count="1"><border/></borders>
<cellStyleXfs count="2"><xf/><xf numFmtId="2"/></cellStyleXfs>
<cellXfs count="2"><xf/>{xf}</cellXfs>
<cellStyles count="2"><cellStyle name="Normal" xfId="0" builtinId="0"/>
<cellStyle name="Currency" xfId="1" builtinId="4"/></cellStyles></styleSheet>"#
    );
    let sheet = format!(
        r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<sheetData><row r="1"><c r="A1" s="1"><v>7</v></c><c r="B1" s="1"><v>8</v></c></row></sheetData>
{}</worksheet>"#,
        if protected {
            r#"<sheetProtection sheet="1"/>"#
        } else {
            ""
        }
    );
    let mut parts = ooxml_opc::unzip_parts(&workbook_bytes()).unwrap();
    for (name, bytes) in &mut parts {
        match name.as_str() {
            "xl/styles.xml" => *bytes = styles.as_bytes().to_vec(),
            "xl/worksheets/sheet1.xml" => *bytes = sheet.as_bytes().to_vec(),
            _ => {}
        }
    }
    (ooxml_opc::rezip_parts(&parts).unwrap(), xf)
}

fn part(bytes: &[u8], name: &str) -> String {
    let parts = ooxml_opc::unzip_parts(bytes).unwrap();
    String::from_utf8(parts.into_iter().find(|(path, _)| path == name).unwrap().1).unwrap()
}

fn format_value(workbook: &Workbook, address: &str) -> Value {
    let style = workbook.model().sheets[0].cell(r(address)).unwrap().style;
    serde_json::to_value(workbook.model().styles.cell_format(style)).unwrap()
}

fn preserved(attributes: &str, children: &str, key: &str, expected: Value) {
    let (bytes, raw_xf) = source(attributes, children, false);
    for collaborative in [false, true] {
        for edit in 0..6 {
            let mut workbook = open_identical(&bytes, collaborative, None);
            let original_index = workbook.model().sheets[0].cell(r("A1")).unwrap().style;
            let before = format_value(&workbook, "A1");
            let range = CellRange::new(r("A1"), r("A1"));
            let mut wanted = before.clone();
            match edit {
                0 => wanted["font"]["bold"] = json!(true),
                1 => wanted["fill"] = json!({"Solid": {"Rgb": "#123456"}}),
                2 => wanted["numberFormat"] = json!({"kind": "builtin", "id": 10}),
                3 => wanted["alignment"]["v"] = json!("Top"),
                4 => wanted["border"]["left"] = json!({"style": "Thin", "color": null}),
                _ => wanted["font"]["italic"] = json!(true),
            }
            if edit == 2 {
                workbook
                    .set_range_number_format(
                        SheetId(0),
                        range,
                        NumberFormatMutation::Percent,
                        options(),
                    )
                    .unwrap();
            } else if edit == 5 {
                let mut captured = workbook.capture_format(SheetId(0), range).unwrap();
                captured.formats[0].font.italic = true;
                workbook
                    .apply_format(SheetId(0), range, captured, options())
                    .unwrap();
            } else {
                let patch = match edit {
                    0 => StylePatch {
                        bold: Some(true),
                        ..Default::default()
                    },
                    1 => StylePatch {
                        fill_color: Some("#123456".into()),
                        ..Default::default()
                    },
                    3 => StylePatch {
                        vertical_alignment: Some(VerticalAlignment::Top),
                        ..Default::default()
                    },
                    _ => StylePatch {
                        border: Some(xlsx_ops::BorderPatch {
                            preset: Some(BorderPreset::Left),
                            ..Default::default()
                        }),
                        ..Default::default()
                    },
                };
                workbook
                    .apply_ops(
                        vec![Op::PatchRangeStyle {
                            sheet: SheetId(0),
                            range,
                            patch,
                        }],
                        options(),
                    )
                    .unwrap();
            }
            let edited_index = workbook.model().sheets[0].cell(r("A1")).unwrap().style;
            let saved = workbook.save().unwrap();
            let saved_parts = ooxml_opc::unzip_parts(&saved).unwrap();
            let parsed = xlsx_parse::parse_workbook(&saved_parts).unwrap();
            assert_eq!(parsed.sheets[0].cell(r("B1")).unwrap().style, Some(1));
            let reopened = Workbook::open(&saved).unwrap();
            let after = format_value(&reopened, "A1");
            assert_eq!(
                after.pointer(key),
                Some(&expected),
                "edit {edit}, collaborative {collaborative}"
            );
            assert_eq!(after, wanted, "edit {edit}");
            assert_eq!(
                reopened.model().sheets[0].cell(r("B1")).unwrap().style,
                original_index
            );
            let styles = part(&saved, "xl/styles.xml");
            assert!(styles.contains(&raw_xf), "{styles}");
            assert!(styles.contains(r#"<cellStyle name="Currency" xfId="1" builtinId="4"/>"#));
            if collaborative {
                let mut peer = Workbook::open_collaborative(&bytes, 74).unwrap();
                peer.apply_update_v1(&workbook.encode_state_as_update_v1(), options())
                    .unwrap();
                assert_eq!(format_value(&peer, "A1"), wanted);
            }
            workbook.undo(options()).unwrap();
            assert_eq!(
                workbook.model().sheets[0].cell(r("A1")).unwrap().style,
                original_index
            );
            assert_eq!(format_value(&workbook, "A1"), before);
            workbook.redo(options()).unwrap();
            assert_eq!(
                workbook.model().sheets[0].cell(r("A1")).unwrap().style,
                edited_index
            );
        }
    }
}

macro_rules! facet_test {
    ($name:ident, $attributes:expr, $children:expr, $key:expr, $($expected:tt)+) => {
        #[test]
        fn $name() {
            preserved($attributes, $children, $key, json!($($expected)+));
        }
    };
}

facet_test!(format_facets_protection, "applyProtection=\"1\"", "<protection locked=\"0\" hidden=\"1\"/>", "/protection", {"locked": false, "hidden": true});
facet_test!(format_facets_parent_style, "xfId=\"1\"", "", "/xfId", 1);
facet_test!(
    format_facets_quote_prefix,
    "quotePrefix=\"1\"",
    "",
    "/quotePrefix",
    true
);
facet_test!(
    format_facets_pivot_button,
    "pivotButton=\"1\"",
    "",
    "/pivotButton",
    true
);
facet_test!(
    format_facets_rotation,
    "",
    "<alignment textRotation=\"135\"/>",
    "/alignment/text_rotation",
    135
);
facet_test!(
    format_facets_zero_rotation,
    "",
    "<alignment textRotation=\"0\"/>",
    "/alignment/text_rotation",
    0
);
facet_test!(
    format_facets_max_rotation,
    "",
    "<alignment textRotation=\"180\"/>",
    "/alignment/text_rotation",
    180
);
facet_test!(
    format_facets_stacked,
    "",
    "<alignment textRotation=\"255\"/>",
    "/alignment/text_rotation",
    255
);
facet_test!(
    format_facets_indent,
    "",
    "<alignment indent=\"4\"/>",
    "/alignment/indent",
    4
);
facet_test!(
    format_facets_relative_indent,
    "",
    "<alignment relativeIndent=\"-2\"/>",
    "/alignment/relative_indent",
    -2
);
facet_test!(
    format_facets_justify_last_line,
    "",
    "<alignment justifyLastLine=\"1\"/>",
    "/alignment/justify_last_line",
    true
);
facet_test!(
    format_facets_reading_order,
    "",
    "<alignment readingOrder=\"2\"/>",
    "/alignment/reading_order",
    2
);
facet_test!(
    format_facets_fill_alignment,
    "",
    "<alignment horizontal=\"fill\"/>",
    "/alignment/h",
    "Fill"
);
facet_test!(
    format_facets_center_continuous,
    "",
    "<alignment horizontal=\"centerContinuous\"/>",
    "/alignment/h",
    "CenterContinuous"
);
facet_test!(
    format_facets_distributed,
    "",
    "<alignment horizontal=\"distributed\"/>",
    "/alignment/h",
    "Distributed"
);

#[test]
fn format_facets_protected_sheet_low_level_edit() {
    let (bytes, _) = source(
        "applyProtection=\"1\"",
        "<protection locked=\"0\" hidden=\"1\"/>",
        true,
    );
    let parts = ooxml_opc::unzip_parts(&bytes).unwrap();
    let parsed = xlsx_parse::parse_workbook_with_package(&parts).unwrap();
    let mut model = parsed.workbook.clone();
    let inverse = xlsx_ops::apply(
        &mut model,
        &Op::PatchRangeStyle {
            sheet: SheetId(0),
            range: CellRange::new(r("A1"), r("A1")),
            patch: StylePatch {
                bold: Some(true),
                ..Default::default()
            },
        },
    )
    .unwrap();
    let saved = xlsx_parse::serialize_workbook_with_package_and_origins_after_edits(
        &model,
        &parsed.package,
        &[Some(0), Some(1), Some(2)],
        &[Default::default(), Default::default(), Default::default()],
        xlsx_parse::SaveEdits {
            changed: true,
            moved_references: false,
        },
    )
    .unwrap();
    let saved = saved
        .into_iter()
        .map(|(name, xml)| (name, xml.into_owned()))
        .collect::<Vec<_>>();
    let reopened = xlsx_parse::parse_workbook(&saved).unwrap();
    let style = reopened.sheets[0].cell(r("A1")).unwrap().style;
    let format = serde_json::to_value(reopened.styles.cell_format(style)).unwrap();
    assert_eq!(
        format["protection"],
        json!({"locked": false, "hidden": true})
    );
    assert!(saved.iter().any(|(name, xml)| {
        name == "xl/worksheets/sheet1.xml"
            && std::str::from_utf8(xml)
                .unwrap()
                .contains("sheetProtection")
    }));
    xlsx_ops::apply_ops(&mut model, &inverse.0).unwrap();
    assert_eq!(model.sheets[0].cell(r("A1")).unwrap().style, Some(1));
}

#[test]
fn format_facets_clear_format_keeps_cell_contents() {
    let (bytes, _) = source(
        "xfId=\"1\" applyProtection=\"1\" quotePrefix=\"1\" pivotButton=\"1\"",
        "<alignment textRotation=\"255\" indent=\"4\"/><protection locked=\"0\" hidden=\"1\"/>",
        false,
    );
    let mut workbook = Workbook::open(&bytes).unwrap();
    let before = workbook.model().sheets[0].cell(r("A1")).unwrap().clone();
    let original_format = workbook.model().styles.cell_format(before.style);
    assert_eq!(
        format_value(&workbook, "A1")["protection"],
        json!({"locked": false, "hidden": true})
    );
    workbook
        .apply_format(
            SheetId(0),
            CellRange::new(r("A1"), r("A1")),
            CapturedFormat {
                rows: 1,
                columns: 1,
                formats: vec![CellFormat::default()],
            },
            options(),
        )
        .unwrap();
    let reopened = Workbook::open(&workbook.save().unwrap()).unwrap();
    assert_eq!(
        reopened
            .model()
            .styles
            .cell_format(reopened.model().sheets[0].cell(r("A1")).unwrap().style),
        CellFormat::default()
    );
    assert_eq!(
        reopened.model().sheets[0].cell(r("A1")).unwrap().value,
        CellValue::Number { value: 7.0 }
    );
    assert_eq!(
        format_value(&reopened, "B1"),
        serde_json::to_value(&original_format).unwrap()
    );
    workbook.undo(options()).unwrap();
    assert_eq!(workbook.model().sheets[0].cell(r("A1")), Some(&before));
    assert_eq!(
        workbook.model().styles.cell_format(before.style),
        original_format
    );
}
