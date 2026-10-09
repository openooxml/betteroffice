use xlsx_model::{Cell, CellRange, CellRef, CellValue, ErrorValue};

use super::style_match_tests::{cells, edit, save};
use crate::tests::{package, parse_workbook_with_package};
use crate::{ParsedWorkbook, SheetAxes, with_legacy_save_path};

const SHARED: &str = concat!(
    r#"<sheetData><row r="1"><c r="A1"><v>1</v></c>"#,
    r#"<c r="B1"><f t="shared" ref="B1:B3" si="0">A1*2</f><v>2</v></c>"#,
    r#"<c r="D1"><v>4</v></c></row>"#,
    r#"<row r="2"><c r="A2"><v>2</v></c><c r="B2"><f t="shared" si="0"/><v>4</v></c></row>"#,
    r#"<row r="3"><c r="A3"><v>3</v></c><c r="B3"><f t="shared" si="0"/><v>6</v></c></row></sheetData>"#,
);

fn parsed(source: &str) -> ParsedWorkbook {
    parse_workbook_with_package(&package(source, &["text &amp; &lt;value&gt;"], false)).unwrap()
}

fn reopen(xml: &str) -> ParsedWorkbook {
    let mut parts = package("", &[], false);
    parts
        .iter_mut()
        .find(|(path, _)| path == "xl/worksheets/sheet1.xml")
        .unwrap()
        .1 = xml.as_bytes().to_vec();
    parse_workbook_with_package(&parts).unwrap()
}

fn number(value: f64) -> CellValue {
    CellValue::Number { value }
}

fn formula(span: &str) -> &str {
    let start = span.find("<f").unwrap();
    let tag_end = start + span[start..].find('>').unwrap() + 1;
    let end = if span[..tag_end].ends_with("/>") {
        tag_end
    } else {
        tag_end + span[tag_end..].find("</f>").unwrap() + 4
    };
    &span[start..end]
}

#[test]
fn unrelated_constant_edit_keeps_shared_group_bytes() {
    let parsed = parsed(SHARED);
    let mut workbook = parsed.workbook.clone();
    edit(&mut workbook, "D1", |cell| cell.value = number(9.0));

    let saved = cells(&save(&parsed, &workbook, Some(SheetAxes::default())));
    let source = cells(SHARED);
    for address in ["B1", "B2", "B3"] {
        assert_eq!(saved[address], source[address], "{address}");
    }
}

#[test]
fn dependent_cache_edit_keeps_shared_formula_bytes() {
    let parsed = parsed(SHARED);
    let mut workbook = parsed.workbook.clone();
    edit(&mut workbook, "A2", |cell| cell.value = number(5.0));
    edit(&mut workbook, "B2", |cell| cell.value = number(10.0));

    let saved = cells(&save(&parsed, &workbook, Some(SheetAxes::default())));
    let source = cells(SHARED);
    for address in ["B1", "B2", "B3"] {
        assert_eq!(formula(&saved[address]), formula(&source[address]));
    }
    assert_eq!(saved["B2"], source["B2"].replace("<v>4</v>", "<v>10</v>"));
    for address in ["A1", "B1", "D1", "A3", "B3"] {
        assert_eq!(saved[address], source[address], "{address}");
    }
}

#[test]
fn cache_edit_with_trailing_children_uses_legacy_rewrite() {
    let source = concat!(
        r#"<sheetData><row r="1"><c r="A1" ph="1"><f ca="1" aca="1" bx="1" del1="1" del2="1">NOW()</f>"#,
        r#"<v>1.00</v><extLst><ext uri="value"><value>keep</value></ext></extLst></c></row></sheetData>"#,
    );
    let parsed = parsed(source);
    let mut workbook = parsed.workbook.clone();
    edit(&mut workbook, "A1", |cell| cell.value = number(2.5));

    let saved = save(&parsed, &workbook, Some(SheetAxes::default()));
    let (oracle, _) =
        with_legacy_save_path(|| save(&parsed, &workbook, Some(SheetAxes::default())));
    assert_eq!(saved, oracle);
    for xml in [&saved, &oracle] {
        assert_eq!(cells(xml)["A1"], r#"<c r="A1"><f>NOW()</f><v>2.5</v></c>"#);
        assert_eq!(
            reopen(xml).workbook.sheets[0].cell(CellRef::new(0, 0)),
            workbook.sheets[0].cell(CellRef::new(0, 0))
        );
    }
}

#[test]
fn cache_type_changes_use_writer_types_and_escaping() {
    let source =
        r#"<sheetData><row r="1"><c r="A1" t="n"><f ca="1">A2</f><v>1</v></c></row></sheetData>"#;
    let parsed = parsed(source);
    for (value, ty, cached) in [
        (
            CellValue::Error {
                value: ErrorValue::NA,
            },
            Some("e"),
            Some("#N/A"),
        ),
        (
            CellValue::Text {
                value: "text & <value>".to_owned(),
            },
            Some("str"),
            Some("text &amp; &lt;value&gt;"),
        ),
        (CellValue::Bool { value: false }, Some("b"), Some("0")),
        (number(3.25), None, Some("3.25")),
        (CellValue::Empty, None, None),
    ] {
        let mut workbook = parsed.workbook.clone();
        edit(&mut workbook, "A1", |cell| cell.value = value);
        let saved = cells(&save(&parsed, &workbook, Some(SheetAxes::default())));
        let ty = ty.map_or(String::new(), |ty| format!(r#" t="{ty}""#));
        let cached = cached.map_or(String::new(), |value| format!("<v>{value}</v>"));
        assert_eq!(
            saved["A1"],
            format!(r#"<c r="A1"{ty}><f ca="1">A2</f>{cached}</c>"#)
        );
    }
}

#[test]
fn missing_and_empty_cached_values_with_extensions_use_legacy_rewrite() {
    for value in ["", "<v/>"] {
        let source = format!(
            r#"<sheetData><row r="1"><c r="A1"><f ca="1">A2</f>{value}<extLst><ext uri="value"/></extLst></c></row></sheetData>"#
        );
        let parsed = parsed(&source);
        let mut workbook = parsed.workbook.clone();
        edit(&mut workbook, "A1", |cell| cell.value = number(2.0));
        let saved = save(&parsed, &workbook, Some(SheetAxes::default()));
        let (oracle, _) =
            with_legacy_save_path(|| save(&parsed, &workbook, Some(SheetAxes::default())));
        assert_eq!(saved, oracle);
        for xml in [&saved, &oracle] {
            assert_eq!(cells(xml)["A1"], r#"<c r="A1"><f>A2</f><v>2</v></c>"#);
            assert_eq!(
                reopen(xml).workbook.sheets[0].cell(CellRef::new(0, 0)),
                workbook.sheets[0].cell(CellRef::new(0, 0))
            );
        }
    }
}

#[test]
fn empty_formula_markup_and_positional_attributes_survive_unmoved_cache_edits() {
    for markup in [
        r#"<f ca="1"/>"#,
        r#"<f dt2D="1" dtr="1" r1="A2" r2="A3">A2+A3</f>"#,
        r#"<f  ca='1' >A2 &lt; 3</f>"#,
    ] {
        let source =
            format!(r#"<sheetData><row r="1"><c r="A1">{markup}<v>1</v></c></row></sheetData>"#);
        let parsed = parsed(&source);
        let mut workbook = parsed.workbook.clone();
        edit(&mut workbook, "A1", |cell| cell.value = number(2.0));
        let saved = cells(&save(&parsed, &workbook, Some(SheetAxes::default())));
        assert_eq!(saved["A1"], format!(r#"<c r="A1">{markup}<v>2</v></c>"#));
    }
}

#[test]
fn formula_change_dissolves_shared_group_like_legacy_save() {
    let parsed = parsed(SHARED);
    let mut workbook = parsed.workbook.clone();
    edit(&mut workbook, "B2", |cell| {
        cell.formula = Some("A2*3".to_owned())
    });

    let saved = save(&parsed, &workbook, Some(SheetAxes::default()));
    let (legacy, _) =
        with_legacy_save_path(|| save(&parsed, &workbook, Some(SheetAxes::default())));
    assert_eq!(saved, legacy);
    let saved = cells(&saved);
    for (address, formula) in [("B1", "A1*2"), ("B2", "A2*3"), ("B3", "A3*2")] {
        assert!(saved[address].contains(&format!("<f>{formula}</f>")));
        assert!(!saved[address].contains(r#"t="shared""#));
    }
}

#[test]
fn deleted_member_dissolves_shared_group_like_legacy_save() {
    let parsed = parsed(SHARED);
    let mut workbook = parsed.workbook.clone();
    edit(&mut workbook, "B2", |cell| *cell = Cell::default());

    let saved = save(&parsed, &workbook, Some(SheetAxes::default()));
    let (legacy, _) =
        with_legacy_save_path(|| save(&parsed, &workbook, Some(SheetAxes::default())));
    assert_eq!(saved, legacy);
    let saved = cells(&saved);
    assert!(!saved.contains_key("B2"));
    assert_eq!(saved["B1"], r#"<c r="B1"><f>A1*2</f><v>2</v></c>"#);
    assert_eq!(saved["B3"], r#"<c r="B3"><f>A3*2</f><v>6</v></c>"#);
}

#[test]
fn removed_formula_dissolves_shared_group_like_legacy_save() {
    let parsed = parsed(SHARED);
    let mut workbook = parsed.workbook.clone();
    edit(&mut workbook, "B2", |cell| cell.formula = None);
    let saved = save(&parsed, &workbook, Some(SheetAxes::default()));
    let (legacy, _) =
        with_legacy_save_path(|| save(&parsed, &workbook, Some(SheetAxes::default())));
    assert_eq!(saved, legacy);
    assert_eq!(cells(&saved)["B2"], r#"<c r="B2"><v>4</v></c>"#);
    assert!(!saved.contains(r#"t="shared""#));
}

#[test]
fn deleted_source_row_dissolves_shared_group_like_legacy_save() {
    let parsed = parsed(SHARED);
    let mut workbook = parsed.workbook.clone();
    workbook.sheets[0].remap_cells(|at| match at.row {
        0 => Some(at),
        1 => None,
        _ => Some(CellRef::new(at.row - 1, at.col)),
    });
    let mut axes = SheetAxes::default();
    axes.rows.delete(1, 1);
    let saved = save(&parsed, &workbook, Some(axes.clone()));
    let (legacy, _) = with_legacy_save_path(|| save(&parsed, &workbook, Some(axes)));
    assert_eq!(saved, legacy);
    assert!(!saved.contains(r#"t="shared""#));
}

#[test]
fn inserted_row_remaps_shared_ref_with_and_without_cache_edit() {
    for cache_edit in [false, true] {
        let parsed = parsed(SHARED);
        let mut workbook = parsed.workbook.clone();
        workbook.sheets[0].remap_cells(|at| Some(CellRef::new(at.row + 1, at.col)));
        let mut axes = SheetAxes::default();
        axes.rows.insert(0, 1);
        if cache_edit {
            edit(&mut workbook, "B2", |cell| cell.value = number(8.0));
            edit(&mut workbook, "B3", |cell| cell.value = number(10.0));
        }

        let saved = cells(&save(&parsed, &workbook, Some(axes)));
        let source = cells(SHARED);
        for (before, after) in [("B1", "B2"), ("B2", "B3"), ("B3", "B4")] {
            let mut expected = source[before]
                .replace(&format!(r#"r="{before}""#), &format!(r#"r="{after}""#))
                .replace(r#"ref="B1:B3""#, r#"ref="B2:B4""#);
            if cache_edit && after == "B2" {
                expected = expected.replace("<v>2</v>", "<v>8</v>");
            }
            if cache_edit && after == "B3" {
                expected = expected.replace("<v>4</v>", "<v>10</v>");
            }
            assert_eq!(saved[after], expected, "{before}: cache edit {cache_edit}");
        }
    }
}

#[test]
fn equivalent_styles_keep_shared_group_and_source_index() {
    let source = SHARED.replace(r#"r="B2""#, r#"r="B2" s="2""#);
    let mut parts = package(&source, &[], false);
    parts.push(("xl/styles.xml".to_owned(), concat!(
        r#"<styleSheet><cellXfs count="3"><xf numFmtId="0"/><xf numFmtId="2"/>"#,
        r#"<xf numFmtId="2" applyProtection="1"><protection locked="0"/></xf></cellXfs></styleSheet>"#,
    ).as_bytes().to_vec()));
    let parsed = parse_workbook_with_package(&parts).unwrap();
    for cache_edit in [false, true] {
        let mut workbook = parsed.workbook.clone();
        edit(&mut workbook, "B2", |cell| {
            cell.style = Some(1);
            if cache_edit {
                cell.value = number(10.0);
            }
        });
        edit(&mut workbook, "D1", |cell| cell.value = number(9.0));
        let saved = cells(&save(&parsed, &workbook, Some(SheetAxes::default())));
        let original = cells(&source);
        let expected = if cache_edit {
            original["B2"].replace("<v>4</v>", "<v>10</v>")
        } else {
            original["B2"].clone()
        };
        assert_eq!(saved["B2"], expected);
        assert_eq!(saved["B1"], original["B1"]);
        assert_eq!(saved["B3"], original["B3"]);
    }
}

#[test]
fn genuine_style_change_keeps_formula_and_writes_new_index() {
    let parsed = parsed(SHARED);
    let mut workbook = parsed.workbook.clone();
    edit(&mut workbook, "B2", |cell| cell.style = Some(1));
    let saved = cells(&save(&parsed, &workbook, Some(SheetAxes::default())));
    let source = cells(SHARED);
    assert_eq!(
        saved["B2"],
        r#"<c r="B2" s="1"><f t="shared" si="0"/><v>4</v></c>"#
    );
    assert_eq!(saved["B1"], source["B1"]);
    assert_eq!(saved["B3"], source["B3"]);
}

#[test]
fn uncertain_metadata_and_dirty_arrays_use_legacy_rewrite() {
    for (attributes, formula) in [
        (r#"cm="1" vm="2""#, r#"<f ca="1">A2</f>"#),
        (r#"cm="1""#, r#"<f t="array" ref="A1:A2" ca="1">A3</f>"#),
        ("", r#"<f t="array" ref="A1:A2" ca="1">A3</f>"#),
    ] {
        let source = format!(
            r#"<sheetData><row r="1"><c r="A1" {attributes}>{formula}<v>1</v></c></row></sheetData>"#
        );
        let parsed = parsed(&source);
        let mut workbook = parsed.workbook.clone();
        edit(&mut workbook, "A1", |cell| cell.value = number(2.0));
        let saved = save(&parsed, &workbook, Some(SheetAxes::default()));
        let (legacy, _) =
            with_legacy_save_path(|| save(&parsed, &workbook, Some(SheetAxes::default())));
        assert_eq!(saved, legacy);
        let expected = if formula.contains(r#"t="array""#) {
            r#"<c r="A1"><f t="array" ref="A1:A2">A3</f><v>2</v></c>"#
        } else {
            r#"<c r="A1"><f>A2</f><v>2</v></c>"#
        };
        assert_eq!(cells(&saved)["A1"], expected);
    }
}

#[test]
fn changed_spill_range_uses_legacy_rewrite() {
    let source = r#"<sheetData><row r="1"><c r="A1" cm="1"><f t="array" ref="A1:A2" ca="1">A3</f><v>1</v></c></row></sheetData>"#;
    let parsed = parsed(source);
    let mut workbook = parsed.workbook.clone();
    let at = CellRef::new(0, 0);
    workbook.sheets[0].set_array_formula(at, CellRange::new(at, at));
    edit(&mut workbook, "A1", |cell| cell.style = Some(1));
    let saved = save(&parsed, &workbook, Some(SheetAxes::default()));
    let (legacy, _) =
        with_legacy_save_path(|| save(&parsed, &workbook, Some(SheetAxes::default())));
    assert_eq!(saved, legacy);
    assert!(cells(&saved)["A1"].contains(r#"<f t="array" ref="A1">A3</f>"#));
}

#[test]
fn cache_types_and_styles_with_extensions_match_legacy_rewrite() {
    for cached in ["", "<v/>", "<v>1.00</v>"] {
        let source = format!(
            r#"<sheetData><row r="1"><c r="A1" ph="1"><f ca="1">A2</f>{cached}<extLst><ext uri="cache"/></extLst></c></row></sheetData>"#
        );
        let parsed = parsed(&source);
        for (value, ty, cached) in [
            (number(3.25), "", "<v>3.25</v>"),
            (CellValue::Bool { value: false }, r#" t="b""#, "<v>0</v>"),
            (
                CellValue::Error {
                    value: ErrorValue::NA,
                },
                r#" t="e""#,
                "<v>#N/A</v>",
            ),
            (
                CellValue::Text {
                    value: "text & <value>".to_owned(),
                },
                r#" t="str""#,
                "<v>text &amp; &lt;value&gt;</v>",
            ),
            (CellValue::Empty, "", ""),
        ] {
            let mut workbook = parsed.workbook.clone();
            edit(&mut workbook, "A1", |cell| {
                cell.value = value;
                cell.style = Some(1);
            });
            let saved = save(&parsed, &workbook, Some(SheetAxes::default()));
            let (oracle, _) =
                with_legacy_save_path(|| save(&parsed, &workbook, Some(SheetAxes::default())));
            assert_eq!(saved, oracle);
            for xml in [&saved, &oracle] {
                assert_eq!(
                    cells(xml)["A1"],
                    format!(r#"<c r="A1" s="1"{ty}><f>A2</f>{cached}</c>"#)
                );
                assert_eq!(
                    reopen(xml).workbook.sheets[0].cell(CellRef::new(0, 0)),
                    workbook.sheets[0].cell(CellRef::new(0, 0))
                );
            }
        }
    }
}

#[test]
fn shared_cache_edits_and_uniform_moves_match_oracle() {
    for moved in [false, true] {
        let parsed = parsed(SHARED);
        let mut workbook = parsed.workbook.clone();
        let mut axes = SheetAxes::default();
        if moved {
            workbook.sheets[0].remap_cells(|at| Some(CellRef::new(at.row + 1, at.col)));
            axes.rows.insert(0, 1);
        }
        let master = if moved { "B2" } else { "B1" };
        let follower = if moved { "B3" } else { "B2" };
        edit(&mut workbook, master, |cell| cell.value = number(8.0));
        edit(&mut workbook, follower, |cell| {
            cell.value = number(10.0);
            cell.style = Some(1);
        });
        let saved = save(&parsed, &workbook, Some(axes.clone()));
        let (oracle, _) = with_legacy_save_path(|| save(&parsed, &workbook, Some(axes)));
        assert_eq!(saved, oracle);
        let saved = cells(&saved);
        let reference = if moved { "B2:B4" } else { "B1:B3" };
        assert_eq!(
            saved[master],
            format!(
                r#"<c r="{master}"><f t="shared" ref="{reference}" si="0">A1*2</f><v>8</v></c>"#
            )
        );
        assert_eq!(
            saved[follower],
            format!(r#"<c r="{follower}" s="1"><f t="shared" si="0"/><v>10</v></c>"#)
        );
    }
}

#[test]
fn shared_master_metadata_fallback_rewrites_followers_and_reopens() {
    for metadata in [r#"cm="1""#, r#"vm="1""#] {
        let source = SHARED.replace(r#"r="B1""#, &format!(r#"r="B1" {metadata}"#));
        let parsed = parsed(&source);
        let mut workbook = parsed.workbook.clone();
        edit(&mut workbook, "B1", |cell| cell.value = number(8.0));
        let saved = save(&parsed, &workbook, Some(SheetAxes::default()));
        let (oracle, _) =
            with_legacy_save_path(|| save(&parsed, &workbook, Some(SheetAxes::default())));
        assert_eq!(saved, oracle);
        let saved_cells = cells(&saved);
        let reopened = reopen(&saved);
        for (address, formula, value) in [
            ("B1", "A1*2", 8.0),
            ("B2", "A2*2", 4.0),
            ("B3", "A3*2", 6.0),
        ] {
            assert_eq!(
                saved_cells[address],
                format!(r#"<c r="{address}"><f>{formula}</f><v>{value}</v></c>"#)
            );
            let at = CellRef::parse_a1(address).unwrap();
            assert_eq!(
                reopened.workbook.sheets[0].cell(at),
                workbook.sheets[0].cell(at)
            );
        }
    }
}

#[test]
fn style_change_with_cell_local_style_namespace_rewrites_and_reopens() {
    const MAIN: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    let source = format!(
        r#"<sheetData><row r="1"><c r="A1" s="1" xmlns:s="{MAIN}"><s:f>A2</s:f><s:v>1</s:v></c></row></sheetData>"#
    );
    let parsed = super::style_match_tests::parsed_sheet(&source);
    let mut workbook = parsed.workbook.clone();
    edit(&mut workbook, "A1", |cell| cell.style = Some(3));
    let saved = save(&parsed, &workbook, Some(SheetAxes::default()));
    let (oracle, _) =
        with_legacy_save_path(|| save(&parsed, &workbook, Some(SheetAxes::default())));
    assert_eq!(saved, oracle);
    for xml in [&saved, &oracle] {
        assert_eq!(cells(xml)["A1"], r#"<c r="A1" s="3"><f>A2</f><v>1</v></c>"#);
        assert_eq!(
            reopen(xml).workbook.sheets[0].cell(CellRef::new(0, 0)),
            workbook.sheets[0].cell(CellRef::new(0, 0))
        );
    }
}

#[test]
fn cache_type_changes_with_cell_local_type_namespace_rewrite_and_reopen() {
    const MAIN: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    let source = format!(
        r#"<sheetData><row r="1"><c r="A1" xmlns:t="{MAIN}"><t:f>A2</t:f><t:v>1</t:v></c></row></sheetData>"#
    );
    let parsed = parsed(&source);
    for (value, ty, cached) in [
        (CellValue::Bool { value: true }, "b", "1"),
        (
            CellValue::Error {
                value: ErrorValue::NA,
            },
            "e",
            "#N/A",
        ),
        (
            CellValue::Text {
                value: "text & <value>".to_owned(),
            },
            "str",
            "text &amp; &lt;value&gt;",
        ),
    ] {
        let mut workbook = parsed.workbook.clone();
        edit(&mut workbook, "A1", |cell| cell.value = value);
        let saved = save(&parsed, &workbook, Some(SheetAxes::default()));
        let (oracle, _) =
            with_legacy_save_path(|| save(&parsed, &workbook, Some(SheetAxes::default())));
        assert_eq!(saved, oracle);
        for xml in [&saved, &oracle] {
            assert_eq!(
                cells(xml)["A1"],
                format!(r#"<c r="A1" t="{ty}"><f>A2</f><v>{cached}</v></c>"#)
            );
            assert_eq!(
                reopen(xml).workbook.sheets[0].cell(CellRef::new(0, 0)),
                workbook.sheets[0].cell(CellRef::new(0, 0))
            );
        }
    }
}

#[test]
fn cell_namespace_collisions_use_legacy_rewrite_for_all_reader_attributes() {
    const MAIN: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    for prefix in ["r", "s", "t", "cm", "vm"] {
        let source = format!(
            r#"<sheetData><row r="1"><c r="A1" s="1" t="n" xmlns:{prefix}="{MAIN}"><{prefix}:f t="array" ref="A1:A2" ca="1">A2</{prefix}:f><{prefix}:v>1</{prefix}:v></c></row></sheetData>"#
        );
        let parsed = super::style_match_tests::parsed_sheet(&source);
        let mut workbook = parsed.workbook.clone();
        edit(&mut workbook, "A1", |cell| cell.value = number(2.0));
        let saved = save(&parsed, &workbook, Some(SheetAxes::default()));
        let (oracle, _) =
            with_legacy_save_path(|| save(&parsed, &workbook, Some(SheetAxes::default())));
        assert_eq!(saved, oracle);
        assert_eq!(
            cells(&saved)["A1"],
            r#"<c r="A1" s="1"><f t="array" ref="A1:A2">A2</f><v>2</v></c>"#,
            "{prefix}"
        );
        let reopened = reopen(&saved);
        let at = CellRef::new(0, 0);
        assert_eq!(
            reopened.workbook.sheets[0].cell(at),
            workbook.sheets[0].cell(at)
        );
        assert_eq!(
            reopened.workbook.sheets[0].array_formula(at),
            workbook.sheets[0].array_formula(at)
        );
    }
}

#[test]
fn cache_style_and_type_edits_with_prefixed_attributes_use_legacy_rewrite() {
    const MAIN: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    let source = format!(
        r#"<sheetData><row r="1"><c r="A1" s="1" t="n" xmlns:x="{MAIN}" x:s="3" x:t="n"><x:f ca="1">A2</x:f><x:v>1</x:v></c></row></sheetData>"#
    );
    let parsed = super::style_match_tests::parsed_sheet(&source);
    for (value, ty, cached) in [
        (number(2.0), "", "2"),
        (CellValue::Bool { value: true }, r#" t="b""#, "1"),
    ] {
        let mut workbook = parsed.workbook.clone();
        edit(&mut workbook, "A1", |cell| {
            cell.style = Some(3);
            cell.value = value;
        });
        let saved = save(&parsed, &workbook, Some(SheetAxes::default()));
        let (oracle, _) =
            with_legacy_save_path(|| save(&parsed, &workbook, Some(SheetAxes::default())));
        assert_eq!(saved, oracle);
        for xml in [&saved, &oracle] {
            assert_eq!(
                cells(xml)["A1"],
                format!(r#"<c r="A1" s="3"{ty}><f>A2</f><v>{cached}</v></c>"#)
            );
            assert_eq!(
                reopen(xml).workbook.sheets[0].cell(CellRef::new(0, 0)),
                workbook.sheets[0].cell(CellRef::new(0, 0))
            );
        }
    }
}

#[test]
fn cache_type_edits_preserve_namespace_and_prefixed_attributes() {
    const MAIN: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    for source_type in ["", r#" t="n""#] {
        let source = format!(
            r#"<sheetData><row r="1"><c r="A1" xmlns:t="{MAIN}" t:t="keep"{source_type}><t:f>A2</t:f><t:v>1</t:v></c></row></sheetData>"#
        );
        let parsed = parsed(&source);
        for (value, ty, cached) in [
            (number(2.0), "", "2"),
            (CellValue::Bool { value: true }, r#" t="b""#, "1"),
        ] {
            let mut workbook = parsed.workbook.clone();
            edit(&mut workbook, "A1", |cell| cell.value = value);
            let saved = save(&parsed, &workbook, Some(SheetAxes::default()));
            let (oracle, _) =
                with_legacy_save_path(|| save(&parsed, &workbook, Some(SheetAxes::default())));
            assert_eq!(saved, oracle);
            assert_eq!(
                cells(&saved)["A1"],
                format!(r#"<c r="A1"{ty}><f>A2</f><v>{cached}</v></c>"#)
            );
            assert_eq!(
                reopen(&saved).workbook.sheets[0].cell(CellRef::new(0, 0)),
                workbook.sheets[0].cell(CellRef::new(0, 0))
            );
        }
    }
}

#[test]
fn prefixed_cache_values_after_default_reset_use_legacy_rewrite() {
    const MAIN: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    for cached in ["<x:v>1</x:v>", "<x:v/>", ""] {
        let source = format!(
            r#"<sheetData><row r="1"><x:c r="A1" xmlns="" xmlns:x="{MAIN}"><x:f>A2</x:f>{cached}</x:c></row></sheetData>"#
        );
        let parsed = parsed(&source);
        let mut workbook = parsed.workbook.clone();
        edit(&mut workbook, "A1", |cell| cell.value = number(2.0));
        let saved = save(&parsed, &workbook, Some(SheetAxes::default()));
        let (oracle, _) =
            with_legacy_save_path(|| save(&parsed, &workbook, Some(SheetAxes::default())));
        assert_eq!(saved, oracle);
        for xml in [&saved, &oracle] {
            assert_eq!(
                xml.as_str(),
                r#"<worksheet><sheetData><row r="1"><c r="A1"><f>A2</f><v>2</v></c></row></sheetData></worksheet>"#
            );
            assert_eq!(
                reopen(xml).workbook.sheets[0].cell(CellRef::new(0, 0)),
                workbook.sheets[0].cell(CellRef::new(0, 0))
            );
        }
    }
}

#[test]
fn prefixed_type_attributes_use_legacy_rewrite_and_reopen() {
    const MAIN: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    for local_namespace in [true, false] {
        let namespace = format!(r#" xmlns:x="{MAIN}""#);
        let cell_namespace: &str = if local_namespace { &namespace } else { "" };
        let row_namespace: &str = if local_namespace { "" } else { &namespace };
        for (source_type, value, ty, cached) in [
            (
                r#"x:t="n""#,
                CellValue::Bool { value: true },
                r#" t="b""#,
                "1",
            ),
            (
                r#"x:t="n""#,
                CellValue::Error {
                    value: ErrorValue::NA,
                },
                r#" t="e""#,
                "#N/A",
            ),
            (r#"t="n" x:t="b""#, number(2.0), "", "2"),
        ] {
            let source = format!(
                r#"<sheetData><row r="1"{row_namespace}><c r="A1"{cell_namespace} {source_type}><f>A2</f><v>1</v></c></row></sheetData>"#
            );
            let parsed = parsed(&source);
            let mut workbook = parsed.workbook.clone();
            edit(&mut workbook, "A1", |cell| cell.value = value);
            let saved = save(&parsed, &workbook, Some(SheetAxes::default()));
            let (oracle, _) =
                with_legacy_save_path(|| save(&parsed, &workbook, Some(SheetAxes::default())));
            assert_eq!(saved, oracle);
            for xml in [&saved, &oracle] {
                assert_eq!(
                    cells(xml)["A1"],
                    format!(r#"<c r="A1"{ty}><f>A2</f><v>{cached}</v></c>"#)
                );
                assert_eq!(
                    reopen(xml).workbook.sheets[0].cell(CellRef::new(0, 0)),
                    workbook.sheets[0].cell(CellRef::new(0, 0))
                );
            }
        }
    }
}

#[test]
fn prefixed_style_attributes_use_legacy_rewrite_and_reopen() {
    const MAIN: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    for local_namespace in [true, false] {
        let namespace = format!(r#" xmlns:x="{MAIN}""#);
        let cell_namespace: &str = if local_namespace { &namespace } else { "" };
        let row_namespace: &str = if local_namespace { "" } else { &namespace };
        let source = format!(
            r#"<sheetData><row r="1"{row_namespace}><c r="A1" s="1"{cell_namespace} x:s="2"><f>A2</f><v>1</v></c></row></sheetData>"#
        );
        let parsed = super::style_match_tests::parsed_sheet(&source);
        for (style, attribute) in [(Some(3), r#" s="3""#), (None, "")] {
            let mut workbook = parsed.workbook.clone();
            edit(&mut workbook, "A1", |cell| cell.style = style);
            let saved = save(&parsed, &workbook, Some(SheetAxes::default()));
            let (oracle, _) =
                with_legacy_save_path(|| save(&parsed, &workbook, Some(SheetAxes::default())));
            assert_eq!(saved, oracle);
            for xml in [&saved, &oracle] {
                assert_eq!(
                    cells(xml)["A1"],
                    format!(r#"<c r="A1"{attribute}><f>A2</f><v>1</v></c>"#)
                );
                assert_eq!(
                    reopen(xml).workbook.sheets[0].cell(CellRef::new(0, 0)),
                    workbook.sheets[0].cell(CellRef::new(0, 0))
                );
            }
        }
    }
}

#[test]
fn nested_extension_values_use_legacy_rewrite_and_reopen() {
    let source = r#"<sheetData><row r="1"><c r="A1" xmlns:x="urn:ext"><f>A2</f><v>1</v><extLst><ext uri="u"><x:v>1</x:v></ext></extLst></c></row></sheetData>"#;
    let parsed = parsed(source);
    let mut workbook = parsed.workbook.clone();
    edit(&mut workbook, "A1", |cell| cell.value = number(2.0));
    let saved = save(&parsed, &workbook, Some(SheetAxes::default()));
    let (oracle, _) =
        with_legacy_save_path(|| save(&parsed, &workbook, Some(SheetAxes::default())));
    assert_eq!(saved, oracle);
    for xml in [&saved, &oracle] {
        assert_eq!(cells(xml)["A1"], r#"<c r="A1"><f>A2</f><v>2</v></c>"#);
        assert_eq!(
            reopen(xml).workbook.sheets[0].cell(CellRef::new(0, 0)),
            workbook.sheets[0].cell(CellRef::new(0, 0))
        );
    }
}

#[test]
fn plain_cache_edits_keep_surrounding_bytes_and_equivalent_style() {
    for cached in ["", "<v/>", "<v>1.00</v>"] {
        let source = format!(
            r#"<sheetData><row r="1" customFormat="1"><c r="A1" s="2" ph="1"> <!--formula--> <f  ca='1' aca='1' >A2 &lt; 3</f>
{cached} <!--tail--> </c><c r="B1"><v>7.00</v></c></row></sheetData>"#
        );
        let parsed = super::style_match_tests::parsed_sheet(&source);
        let mut workbook = parsed.workbook.clone();
        edit(&mut workbook, "A1", |cell| {
            cell.style = Some(1);
            cell.value = number(2.0);
        });
        let expected = if cached.is_empty() {
            source.replace("</f>", "</f><v>2</v>")
        } else {
            source.replace(cached, "<v>2</v>")
        };
        let saved = save(&parsed, &workbook, Some(SheetAxes::default()));
        let (oracle, _) =
            with_legacy_save_path(|| save(&parsed, &workbook, Some(SheetAxes::default())));
        assert_eq!(saved, oracle);
        for xml in [&saved, &oracle] {
            assert_eq!(xml, &format!("<worksheet>{expected}</worksheet>"));
            let reopened = reopen(xml);
            let at = CellRef::new(0, 0);
            let mut expected_cell = workbook.sheets[0].cell(at).unwrap().clone();
            expected_cell.style = Some(2);
            assert_eq!(reopened.workbook.sheets[0].cell(at), Some(&expected_cell));
            assert_eq!(
                reopened.workbook.sheets[0].cell(CellRef::new(0, 1)),
                workbook.sheets[0].cell(CellRef::new(0, 1))
            );
        }
    }
}

#[test]
fn non_plain_cache_shapes_use_legacy_rewrite_and_reopen() {
    for cell in [
        r#"<c r="A1" xmlns=""><f ca="1">A2</f><v>1</v></c>"#,
        r#"<c r="A1" xmlns:x="urn:ext"><f ca="1">A2</f><v>1</v></c>"#,
        r#"<c r="A1" x:flag="keep"><f ca="1">A2</f><v>1</v></c>"#,
        r#"<x:c r="A1"><f ca="1">A2</f><v>1</v></x:c>"#,
        r#"<c r="A1"><x:f ca="1">A2</x:f><v>1</v></c>"#,
        r#"<c r="A1"><f ca="1">A2</f><x:v>1</x:v></c>"#,
        r#"<c r="A1"><f ca="1" xmlns="">A2</f><v>1</v></c>"#,
        r#"<c r="A1"><f ca="1" xmlns:y="urn:ext">A2</f><v>1</v></c>"#,
        r#"<c r="A1"><f ca="1">A2</f><v xmlns="">1</v></c>"#,
        r#"<c r="A1"><f ca="1">A2</f><v xmlns:y="urn:ext">1</v></c>"#,
        r#"<c r="A1"><f ca="1">A2<inner/></f><v>1</v></c>"#,
        r#"<c r="A1"><f ca="1">A2<inner></inner></f><v>1</v></c>"#,
        r#"<c r="A1"><f ca="1">A2</f><v>1<inner/></v></c>"#,
        r#"<c r="A1"><f ca="1">A2</f><v>1<inner></inner></v></c>"#,
        r#"<c r="A1"><other/><f ca="1">A2</f><v>1</v></c>"#,
        r#"<c r="A1"><f ca="1">A2</f><other></other><v>1</v></c>"#,
    ] {
        let source = format!(r#"<sheetData><row r="1" xmlns:x="urn:ext">{cell}</row></sheetData>"#);
        let parsed = parsed(&source);
        let mut workbook = parsed.workbook.clone();
        edit(&mut workbook, "A1", |cell| cell.value = number(3.0));
        let saved = save(&parsed, &workbook, Some(SheetAxes::default()));
        let (oracle, _) =
            with_legacy_save_path(|| save(&parsed, &workbook, Some(SheetAxes::default())));
        assert_eq!(saved, oracle, "{cell}");
        for xml in [&saved, &oracle] {
            assert_eq!(
                cells(xml)["A1"],
                r#"<c r="A1"><f>A2</f><v>3</v></c>"#,
                "{cell}"
            );
            assert_eq!(
                reopen(xml).workbook.sheets[0].cell(CellRef::new(0, 0)),
                workbook.sheets[0].cell(CellRef::new(0, 0)),
                "{cell}"
            );
        }
    }
}

#[test]
fn added_array_state_with_cache_edit_rewrites_and_reopens() {
    let source =
        r#"<sheetData><row r="1"><c r="A1"><f ca="1">A3:A4</f><v>1</v></c></row></sheetData>"#;
    let parsed = parsed(source);
    let mut workbook = parsed.workbook.clone();
    let at = CellRef::new(0, 0);
    let range = CellRange::new(at, CellRef::new(1, 0));
    workbook.sheets[0].set_array_formula(at, range);
    edit(&mut workbook, "A1", |cell| cell.value = number(2.0));
    let saved = save(&parsed, &workbook, Some(SheetAxes::default()));
    let (oracle, _) =
        with_legacy_save_path(|| save(&parsed, &workbook, Some(SheetAxes::default())));
    assert_eq!(saved, oracle);
    assert_eq!(
        cells(&saved)["A1"],
        r#"<c r="A1"><f t="array" ref="A1:A2">A3:A4</f><v>2</v></c>"#
    );
    let reopened = reopen(&saved);
    assert_eq!(reopened.workbook.sheets[0].array_formula(at), Some(range));
    assert_eq!(
        reopened.workbook.sheets[0].cell(at),
        workbook.sheets[0].cell(at)
    );
}

#[test]
fn source_cache_type_transitions_match_oracle() {
    for (source_type, source_value, value, ty, cached) in [
        (
            "n",
            "1",
            CellValue::Text {
                value: "text".to_owned(),
            },
            r#" t="str""#,
            "text",
        ),
        (
            "str",
            "text",
            CellValue::Bool { value: true },
            r#" t="b""#,
            "1",
        ),
        ("e", "#N/A", number(2.0), "", "2"),
    ] {
        let source = format!(
            r#"<sheetData><row r="1"><c r="A1" t="{source_type}"><f ca="1">A2</f><v>{source_value}</v></c></row></sheetData>"#
        );
        let parsed = parsed(&source);
        let mut workbook = parsed.workbook.clone();
        edit(&mut workbook, "A1", |cell| cell.value = value);
        let saved = save(&parsed, &workbook, Some(SheetAxes::default()));
        let (oracle, _) =
            with_legacy_save_path(|| save(&parsed, &workbook, Some(SheetAxes::default())));
        assert_eq!(saved, oracle);
        assert_eq!(
            cells(&saved)["A1"],
            format!(r#"<c r="A1"{ty}><f ca="1">A2</f><v>{cached}</v></c>"#)
        );
    }
}

#[test]
fn ambiguous_cache_children_use_full_rewrite() {
    for children in [
        "<v>1</v><v>2</v>",
        "<is><t>text</t></is>",
        r#"<v xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">1</v>"#,
    ] {
        let source = format!(
            r#"<sheetData><row r="1"><c r="A1"><f ca="1">A2</f>{children}</c></row></sheetData>"#
        );
        let parsed = parsed(&source);
        let mut workbook = parsed.workbook.clone();
        edit(&mut workbook, "A1", |cell| cell.value = number(3.0));
        let saved = save(&parsed, &workbook, Some(SheetAxes::default()));
        let (oracle, _) =
            with_legacy_save_path(|| save(&parsed, &workbook, Some(SheetAxes::default())));
        assert_eq!(saved, oracle);
        assert_eq!(cells(&saved)["A1"], r#"<c r="A1"><f>A2</f><v>3</v></c>"#);
    }
}

#[test]
fn equivalent_style_cache_edit_keeps_source_index_like_oracle() {
    let source =
        r#"<sheetData><row r="1"><c r="A1" s="2"><f ca="1">A2</f><v>1</v></c></row></sheetData>"#;
    let parsed = super::style_match_tests::parsed_sheet(source);
    let mut workbook = parsed.workbook.clone();
    edit(&mut workbook, "A1", |cell| {
        cell.style = Some(1);
        cell.value = number(2.0);
    });
    assert_ne!(
        parsed.workbook.sheets[0]
            .cell(CellRef::new(0, 0))
            .unwrap()
            .style,
        workbook.sheets[0].cell(CellRef::new(0, 0)).unwrap().style
    );
    let saved = save(&parsed, &workbook, Some(SheetAxes::default()));
    let (oracle, _) =
        with_legacy_save_path(|| save(&parsed, &workbook, Some(SheetAxes::default())));
    assert_eq!(saved, oracle);
    assert_eq!(
        cells(&saved)["A1"],
        r#"<c r="A1" s="2"><f ca="1">A2</f><v>2</v></c>"#
    );
}
