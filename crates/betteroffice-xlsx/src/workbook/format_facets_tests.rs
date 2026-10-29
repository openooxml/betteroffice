use serde_json::{Value, json};

use super::edit_tests::{open_identical, options, r, workbook_bytes};
use super::*;

fn source(attributes: &str, children: &str, protected: bool) -> (Vec<u8>, String) {
    let xf = format!(
        r#"<xf numFmtId="2" fontId="0" fillId="0" borderId="0" {attributes}>{children}</xf>"#
    );
    let bytes = package(
        &format!("<xf/>{xf}"),
        r#"<c r="A1" s="1"><v>7</v></c><c r="B1" s="1"><v>8</v></c>"#,
        protected,
    );
    (bytes, xf)
}

pub(super) fn package(cell_xfs: &str, cells: &str, protected: bool) -> Vec<u8> {
    let count = cell_xfs.matches("<xf").count();
    let styles = format!(
        r#"<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<fonts count="1"><font><name val="Calibri"/><sz val="11"/></font></fonts>
<fills count="1"><fill><patternFill patternType="none"/></fill></fills>
<borders count="1"><border/></borders>
<cellStyleXfs count="2"><xf/><xf numFmtId="2"/></cellStyleXfs>
<cellXfs count="{count}">{cell_xfs}</cellXfs>
<cellStyles count="2"><cellStyle name="Normal" xfId="0" builtinId="0"/>
<cellStyle name="Currency" xfId="1" builtinId="4"/></cellStyles></styleSheet>"#
    );
    let sheet = format!(
        r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<sheetData><row r="1">{cells}</row></sheetData>
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
    ooxml_opc::rezip_parts(&parts).unwrap()
}

pub(super) fn part(bytes: &[u8], name: &str) -> String {
    let parts = ooxml_opc::unzip_parts(bytes).unwrap();
    String::from_utf8(parts.into_iter().find(|(path, _)| path == name).unwrap().1).unwrap()
}

fn format_value(workbook: &Workbook, address: &str) -> Value {
    let style = workbook.model().sheets[0].cell(r(address)).unwrap().style;
    serde_json::to_value(workbook.model().styles.cell_format(style)).unwrap()
}

#[test]
fn source_style_xml_markers_restore_idempotently() {
    let source = package(
        r#"<xf/><xf numFmtId="2"><extLst><ext uri="keep"/></extLst></xf>"#,
        r#"<c r="A1" s="1"><v>7</v></c>"#,
        false,
    );
    let classified = source_style_xml_markers(&source).unwrap();
    assert!(classified.xf_has_extra_xml(1));
    let mut restored: Stylesheet =
        serde_json::from_value(serde_json::to_value(&classified).unwrap()).unwrap();
    assert!(!restored.has_xf_xml_markers());
    restored.inherit_xf_xml_markers(&classified);
    assert!(restored.xf_has_extra_xml(1));
    let clean = source_style_xml_markers(&package("<xf/>", "", false)).unwrap();
    assert!(clean.has_xf_xml_markers());
    assert!(!clean.xf_has_extra_xml(0));
    restored.inherit_xf_xml_markers(&clean);
    assert!(restored.xf_has_extra_xml(1));
}

#[test]
fn source_styles_part_name_matches_open_resolution() {
    assert_eq!(xlsx_parse::styles_part_name(None).unwrap(), "xl/styles.xml");
    assert_eq!(
        xlsx_parse::styles_part_name(Some(b"<Relationships/>")).unwrap(),
        "xl/styles.xml"
    );
    for (target, expected) in [
        ("styles.xml", "xl/styles.xml"),
        ("customStyles.xml", "xl/customStyles.xml"),
        ("./styles/../customStyles.xml", "xl/customStyles.xml"),
        ("../customStyles.xml", "customStyles.xml"),
        ("/custom/styles.xml", "custom/styles.xml"),
    ] {
        let rels = format!(
            r#"<Relationships><Relationship Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="{target}"/></Relationships>"#
        );
        assert_eq!(
            xlsx_parse::styles_part_name(Some(rels.as_bytes())).unwrap(),
            expected
        );
    }
}

#[test]
fn source_style_xml_markers_ignore_unreadable_sources() {
    assert!(source_style_xml_markers(b"invalid zip").is_none());
    let source = package("<xf/>", "", false);
    for (name, xml) in [
        ("xl/styles.xml", None),
        (
            "xl/styles.xml",
            Some("<styleSheet><cellXfs><xf/></styleSheet>"),
        ),
        (
            "xl/styles.xml",
            Some(r#"<styleSheet><cellXfs><xf numFmtId="&unknown;"/></cellXfs></styleSheet>"#),
        ),
        (
            "xl/_rels/workbook.xml.rels",
            Some("<Relationships><Relationship></Relationships>"),
        ),
    ] {
        let mut parts = ooxml_opc::unzip_parts(&source).unwrap();
        parts.retain(|(path, _)| path != name);
        if let Some(xml) = xml {
            parts.push((name.to_string(), xml.as_bytes().to_vec()));
        }
        assert!(source_style_xml_markers(&ooxml_opc::rezip_parts(&parts).unwrap()).is_none());
    }
    let mut parts = ooxml_opc::unzip_parts(&source).unwrap();
    parts.retain(|(name, _)| name != "xl/_rels/workbook.xml.rels");
    let styles = source_style_xml_markers(&ooxml_opc::rezip_parts(&parts).unwrap()).unwrap();
    assert!(styles.has_xf_xml_markers());
    assert!(!styles.xf_has_extra_xml(0));
}

#[test]
fn captured_sort_keeps_default_and_extended_source_xfs() {
    let default_xf = r#"<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>"#;
    let extended_xf = r#"<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"><extLst><ext uri="{6E1A2B3C-4D5E-4F60-8A9B-0C1D2E3F4A5B}"><x:keep xmlns:x="urn:protected-sort" value="kept"/></ext></extLst></xf>"#;
    let bytes = package(
        &format!("{default_xf}<xf numFmtId=\"2\"/><xf numFmtId=\"14\"/>{extended_xf}"),
        r#"<c r="A1" s="3"><v>7</v></c><c r="B1"><v>8</v></c><c r="C1" s="0"><v>9</v></c>"#,
        false,
    );
    for collaborative in [false, true] {
        for plain_source in ["B1", "C1"] {
            for move_values in [false, true] {
                let mut workbook = open_identical(&bytes, collaborative, None);
                let capture = |address| {
                    serde_json::to_value(
                        workbook
                            .capture_format(SheetId(0), CellRange::new(r(address), r(address)))
                            .unwrap(),
                    )
                    .unwrap()
                };
                let plain = capture(plain_source);
                let extended = capture("A1");
                let mut payload = Vec::new();
                if move_values {
                    payload.extend([
                        json!({"type": "setCell", "sheet": 0, "at": {"row": 0, "col": 0},
                            "cell": {"value": {"kind": "number", "value": if plain_source == "B1" { 8 } else { 9 }}, "formula": null}}),
                        json!({"type": "setCell", "sheet": 0, "at": {"row": 0, "col": 1},
                            "cell": {"value": {"kind": "number", "value": 7}, "formula": null}}),
                    ]);
                }
                payload.push(json!({"type": "applyRangeFormat", "sheet": 0,
                    "range": {"start": {"row": 0, "col": 0}, "end": {"row": 0, "col": 1}},
                    "format": {"rows": 1, "columns": 2,
                        "formats": [plain["formats"][0], extended["formats"][0]],
                        "sourceStyles": [plain["sourceStyles"][0], extended["sourceStyles"][0]],
                        "sourceFormats": [plain["sourceFormats"][0], extended["sourceFormats"][0]]}}));
                let ops: Vec<Op> = serde_json::from_value(json!(payload)).unwrap();
                assert!(workbook.apply_ops(ops, options()).unwrap().applied);
                let expected_plain = (plain_source == "C1").then_some(0);
                assert_eq!(
                    workbook.model().sheets[0].cell(r("A1")).unwrap().style,
                    expected_plain,
                    "{collaborative}: {plain_source}: {move_values}"
                );
                assert_eq!(
                    workbook.model().sheets[0].cell(r("B1")).unwrap().style,
                    Some(3)
                );
                let saved = workbook.save().unwrap();
                assert_eq!(part(&saved, "xl/styles.xml"), part(&bytes, "xl/styles.xml"));
                let reopened = Workbook::open(&saved).unwrap();
                assert_eq!(
                    reopened.model().sheets[0].cell(r("A1")).unwrap().style,
                    expected_plain
                );
                assert_eq!(
                    reopened.model().sheets[0].cell(r("B1")).unwrap().style,
                    Some(3)
                );
            }
        }
    }
}

#[test]
fn painting_an_unstyled_cell_keeps_the_destination_cell_xml() {
    let bytes = package(
        r#"<xf/><xf numFmtId="4"/>"#,
        r#"<c r="A1"><v>7</v></c><c r="B1"><v>8</v></c>"#,
        false,
    );
    for collaborative in [false, true] {
        let mut workbook = open_identical(&bytes, collaborative, None);
        let captured = workbook
            .capture_format(SheetId(0), CellRange::new(r("A1"), r("A1")))
            .unwrap();
        assert_eq!(captured.source_styles, vec![None]);
        assert!(
            !workbook
                .apply_format(
                    SheetId(0),
                    CellRange::new(r("B1"), r("B1")),
                    captured,
                    options(),
                )
                .unwrap()
                .applied
        );
        assert_eq!(
            workbook.model().sheets[0].cell(r("B1")).unwrap().style,
            None
        );
        assert_eq!(
            part(&workbook.save().unwrap(), "xl/worksheets/sheet1.xml"),
            part(&bytes, "xl/worksheets/sheet1.xml")
        );
    }
}

#[test]
fn painting_a_missing_cell_in_a_styled_row_keeps_the_row_xf() {
    let row_xf = r#"<xf numFmtId="2" fontId="0"><extLst><ext uri="row-source"/></extLst></xf>"#;
    let source = package(
        &format!("<xf/>{row_xf}"),
        r#"<c r="B1"><v>8</v></c>"#,
        false,
    );
    let mut parts = ooxml_opc::unzip_parts(&source).unwrap();
    for (name, bytes) in &mut parts {
        if name == "xl/worksheets/sheet1.xml" {
            *bytes = String::from_utf8(bytes.clone())
                .unwrap()
                .replace(
                    "</sheetData>",
                    r#"<row r="2" s="1" customFormat="1"/></sheetData>"#,
                )
                .into_bytes();
        }
    }
    let bytes = ooxml_opc::rezip_parts(&parts).unwrap();
    for collaborative in [false, true] {
        let mut workbook = open_identical(&bytes, collaborative, None);
        let captured = workbook
            .capture_format(SheetId(0), CellRange::new(r("A2"), r("A2")))
            .unwrap();
        assert_eq!(captured.source_styles, vec![Some(1)]);
        workbook
            .apply_format(
                SheetId(0),
                CellRange::new(r("B1"), r("B1")),
                captured,
                options(),
            )
            .unwrap();
        assert_eq!(
            workbook.model().sheets[0].cell(r("B1")).unwrap().style,
            Some(1)
        );
        let saved = workbook.save().unwrap();
        assert_eq!(part(&saved, "xl/styles.xml"), part(&bytes, "xl/styles.xml"));
        assert_eq!(
            Workbook::open(&saved).unwrap().model().sheets[0]
                .cell(r("B1"))
                .unwrap()
                .style,
            Some(1)
        );
    }
}

#[test]
fn facet_interning_avoids_xfs_with_extra_xml() {
    for (attributes, children) in [
        ("", "<extLst><ext uri=\"extra-style\"/></extLst>"),
        ("", "<unknown value=\"extra-style\"/>"),
        ("", "<!--extra-style-->"),
        (" custom=\"extra-style\"", ""),
        ("", "<alignment custom=\"extra-style\"/>"),
        ("", "<protection custom=\"extra-style\"/>"),
    ] {
        for size in [2, 128] {
            for has_clean in [false, true] {
                let extra = format!(r#"<xf numFmtId="2"{attributes}>{children}</xf>"#);
                let fillers = (2..size)
                    .map(|index| format!(r#"<xf numFmtId="{}"/>"#, index + 10))
                    .collect::<String>();
                let clean = if !has_clean {
                    ""
                } else if children.starts_with("<protection") {
                    "<xf numFmtId=\"2\"><protection locked=\"1\" hidden=\"0\"/></xf>"
                } else {
                    "<xf numFmtId=\"2\"/>"
                };
                let bytes = package(
                    &format!("<xf/>{extra}{fillers}{clean}"),
                    r#"<c r="A1" s="1"><v>7</v></c><c r="B1"><v>8</v></c>"#,
                    false,
                );
                let mut workbook = Workbook::open(&bytes).unwrap();
                let mut styles = workbook.model().styles.clone();
                let format = styles.cell_format(Some(1));
                let index = styles.intern_cell_format(&format).unwrap().unwrap();
                assert_ne!(index, 1, "{size}: {has_clean}: {attributes}: {children}");
                assert_eq!(index, size);
                assert_eq!(styles.intern_cell_format(&format).unwrap(), Some(index));
                workbook
                    .apply_ops(
                        vec![Op::ApplyRangeFormat {
                            sheet: SheetId(0),
                            range: CellRange::new(r("B1"), r("B1")),
                            format: CapturedFormat {
                                rows: 1,
                                columns: 1,
                                formats: vec![format],
                                kept: Vec::new(),
                                source_styles: Vec::new(),
                                source_formats: Vec::new(),
                            },
                        }],
                        options(),
                    )
                    .unwrap();
                let saved = workbook.save().unwrap();
                let saved_styles = part(&saved, "xl/styles.xml");
                assert_eq!(saved_styles.matches("extra-style").count(), 1);
                assert_eq!(
                    Workbook::open(&saved).unwrap().model().sheets[0]
                        .cell(r("B1"))
                        .unwrap()
                        .style,
                    Some(index)
                );
            }
        }
    }
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
                kept: Vec::new(),
                source_styles: Vec::new(),
                source_formats: Vec::new(),
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

fn released_format() -> Value {
    json!({
        "font": {"name": null, "size_pt": null, "bold": true, "italic": false,
            "underline": false, "strike": false, "color": null},
        "fill": "None",
        "border": {"left": null, "right": null, "top": null, "bottom": null},
        "numberFormat": {"kind": "builtin", "id": 0},
        "alignment": {"h": null, "v": null, "wrap_text": false, "shrink_to_fit": false}
    })
}

#[test]
fn format_facets_released_captures_keep_the_facets_they_do_not_carry() {
    let (bytes, _) = source(
        "xfId=\"1\" applyProtection=\"1\" quotePrefix=\"1\"",
        "<alignment textRotation=\"90\"/><protection locked=\"0\"/>",
        false,
    );
    let range = CellRange::new(r("A1"), r("A1"));
    for locked in [None, Some(true)] {
        let mut format = released_format();
        if let Some(locked) = locked {
            format["protection"] = json!({ "locked": locked });
        }
        let captured = json!({"rows": 1, "columns": 1, "formats": [format]});
        let captured: CapturedFormat = serde_json::from_value(captured).unwrap();
        let mut workbook = Workbook::open(&bytes).unwrap();
        workbook
            .apply_format(SheetId(0), range, captured, options())
            .unwrap();
        let after = format_value(&workbook, "A1");
        assert_eq!(after["font"]["bold"], true);
        assert_eq!(
            after["protection"],
            json!({"locked": locked.is_some(), "hidden": false})
        );
        assert_eq!(after["alignment"]["text_rotation"], 90);
        assert_eq!(after["xfId"], 1);
        assert_eq!(after["quotePrefix"], true);
    }

    let mut workbook = Workbook::open(&bytes).unwrap();
    let plain = workbook
        .capture_format(SheetId(0), CellRange::new(r("C1"), r("C1")))
        .unwrap();
    let written = serde_json::to_value(&plain).unwrap();
    let format = written["formats"][0].as_object().unwrap();
    for key in ["protection", "xfId", "quotePrefix", "pivotButton"] {
        assert!(format.contains_key(key), "{key}: {written}");
    }
    let alignment = format["alignment"].as_object().unwrap();
    for key in [
        "text_rotation",
        "indent",
        "relative_indent",
        "justify_last_line",
        "reading_order",
    ] {
        assert!(alignment.contains_key(key), "{key}: {written}");
    }
    let replayed: CapturedFormat = serde_json::from_value(written).unwrap();
    assert_eq!(replayed, plain);
    workbook
        .apply_format(SheetId(0), range, replayed, options())
        .unwrap();
    assert_eq!(
        format_value(&workbook, "A1"),
        serde_json::to_value(CellFormat::default()).unwrap()
    );
}

#[test]
fn format_facets_survive_a_collaboration_state_saved_before_them() {
    let faceted = r#"<xf numFmtId="2" fontId="0" fillId="0" borderId="0" xfId="1" applyProtection="1"><alignment textRotation="90"/><protection locked="0" hidden="0"/></xf>"#;
    let bytes = package(
        &format!(r#"<xf/><xf numFmtId="2" fontId="0" fillId="0" borderId="0"/>{faceted}"#),
        r#"<c r="A1" s="2"><v>7</v></c><c r="B1" s="1"><v>8</v></c>"#,
        false,
    );
    let parts = ooxml_opc::unzip_parts(&bytes).unwrap();
    let mut released = xlsx_parse::parse_workbook(&parts).unwrap();
    for xf in &mut released.styles.cell_xfs {
        *xf = xlsx_model::Xf {
            font: xf.font,
            fill: xf.fill,
            border: xf.border,
            num_fmt_id: xf.num_fmt_id,
            alignment: xf
                .alignment
                .as_ref()
                .map(|alignment| xlsx_model::Alignment {
                    h: alignment.h,
                    v: alignment.v,
                    wrap_text: alignment.wrap_text,
                    shrink_to_fit: alignment.shrink_to_fit,
                    ..Default::default()
                })
                .filter(|alignment| !alignment.is_empty()),
            ..Default::default()
        };
    }
    let state = crate::authority::pre_facet_state_for_test(&released);
    let mut workbook = Workbook::open_collaborative(&bytes, 77).unwrap();
    workbook.apply_update_v1(&state, options()).unwrap();
    assert_eq!(
        workbook.model().sheets[0].cell(r("A1")).unwrap().style,
        Some(2)
    );
    assert_eq!(
        workbook.model().sheets[0].cell(r("B1")).unwrap().style,
        Some(1)
    );
    let format = format_value(&workbook, "A1");
    assert_eq!(
        format["protection"],
        json!({"locked": false, "hidden": false})
    );
    assert_eq!(format["alignment"]["text_rotation"], 90);
    assert_eq!(format["xfId"], 1);
    let saved = workbook.save().unwrap();
    assert!(part(&saved, "xl/styles.xml").contains(faceted));
    let saved_parts = ooxml_opc::unzip_parts(&saved).unwrap();
    let reopened = xlsx_parse::parse_workbook(&saved_parts).unwrap();
    assert_eq!(reopened.sheets[0].cell(r("A1")).unwrap().style, Some(2));
    assert_eq!(reopened.styles.cell_xfs.len(), 3);
}

#[test]
fn format_facets_open_hashes_cells_no_more_often_than_before_them() {
    let bytes = package(
        r#"<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="2" fontId="0" fillId="0" borderId="0" xfId="0"/>"#,
        r#"<c r="A1" s="1"><v>7</v></c><c r="B1" s="0"><v>8</v></c>"#,
        false,
    );
    let before = crate::authority::fingerprint_passes();
    Workbook::open(&bytes).unwrap();
    // releases before facets hashed schemas 3, 4, 5, 6 and 3 with defined names,
    // for the stylesheet and for its reading before applyX flags: ten passes
    assert!(crate::authority::fingerprint_passes() - before <= 10);
}

/// The collaboration state a release before format facets persisted for `bytes`.
fn released_state(bytes: &[u8]) -> Vec<u8> {
    let parts = ooxml_opc::unzip_parts(bytes).unwrap();
    let mut released = xlsx_parse::parse_workbook(&parts).unwrap();
    for xf in &mut released.styles.cell_xfs {
        *xf = xf.released_reading();
    }
    crate::authority::pre_facet_state_for_test(&released)
}

#[test]
fn format_facets_peers_reconnect_after_sharing_an_adopted_released_state() {
    let (bytes, _) = source("xfId=\"1\"", "<alignment textRotation=\"90\"/>", false);
    let state = released_state(&bytes);
    let mut a = Workbook::open_collaborative(&bytes, 77).unwrap();
    a.apply_update_v1(&state, options()).unwrap();
    let mut b = Workbook::open_collaborative(&bytes, 78).unwrap();
    b.apply_update_v1(&a.encode_state_as_update_v1(), options())
        .unwrap();
    b.edit_cell(SheetId(0), r("D1"), "6", options()).unwrap();
    a.edit_cell(SheetId(0), r("C1"), "5", options()).unwrap();
    a.apply_update_v1(&b.encode_state_as_update_v1(), options())
        .unwrap();
    b.apply_update_v1(&a.encode_state_as_update_v1(), options())
        .unwrap();
    for workbook in [&a, &b] {
        let sheet = &workbook.model().sheets[0];
        let value = |address| sheet.cell(r(address)).unwrap().value.clone();
        assert_eq!(value("C1"), xlsx_model::CellValue::Number { value: 5.0 });
        assert_eq!(value("D1"), xlsx_model::CellValue::Number { value: 6.0 });
        let format = format_value(workbook, "A1");
        assert_eq!(format["alignment"]["text_rotation"], 90);
        assert_eq!(format["xfId"], 1);
    }

    let mut edited = Workbook::open_collaborative(&bytes, 79).unwrap();
    edited
        .edit_cell(SheetId(0), r("C1"), "1", options())
        .unwrap();
    assert!(matches!(
        edited.apply_update_v1(&state, options()),
        Err(Error::CollaborativeStructureChanged)
    ));
}

fn with_palette(bytes: &[u8]) -> Vec<u8> {
    let palette = r#"<colors><indexedColors><rgbColor rgb="FF123456"/></indexedColors></colors>"#;
    let mut parts = ooxml_opc::unzip_parts(bytes).unwrap();
    for (name, part) in &mut parts {
        if name == "xl/styles.xml" {
            let styles = String::from_utf8(part.clone()).unwrap();
            *part = styles
                .replace("</styleSheet>", &format!("{palette}</styleSheet>"))
                .into_bytes();
        }
    }
    ooxml_opc::rezip_parts(&parts).unwrap()
}

#[test]
fn format_facets_cleared_in_a_current_state_stay_cleared_when_adopted() {
    let plain = package(
        r#"<xf/><xf numFmtId="2" fontId="0" fillId="0" borderId="0"><alignment textRotation="90"/></xf>"#,
        r#"<c r="A1" s="1"><v>7</v></c><c r="B1" s="1"><v>8</v></c>"#,
        false,
    );
    let range = CellRange::new(r("A1"), r("A1"));
    let mut peer = Workbook::open_collaborative(&plain, 81).unwrap();
    let mut captured = peer.capture_format(SheetId(0), range).unwrap();
    captured.formats[0].alignment.text_rotation = None;
    peer.apply_format(SheetId(0), range, captured, options())
        .unwrap();
    assert_eq!(
        format_value(&peer, "A1")["alignment"]["text_rotation"],
        Value::Null
    );

    let palette = with_palette(&plain);
    let mut workbook = Workbook::open_collaborative(&palette, 82).unwrap();
    assert!(!workbook.model().styles.indexed_colors.is_empty());
    workbook
        .apply_update_v1(&peer.encode_state_as_update_v1(), options())
        .unwrap();
    let saved = workbook.save().unwrap();
    let reopened = Workbook::open(&saved).unwrap();
    for workbook in [&workbook, &reopened] {
        let cleared = format_value(workbook, "A1");
        assert_eq!(cleared["alignment"]["text_rotation"], Value::Null);
        assert_eq!(cleared["numberFormat"]["id"], 2);
        assert_eq!(
            format_value(workbook, "B1")["alignment"]["text_rotation"],
            90
        );
    }
}
