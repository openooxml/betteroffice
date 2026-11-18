use xlsx_model::{CellRef, CellValue, Color, RunBaseline, RunUnderline};

use crate::tests::package;
use crate::{MAX_RICH_TEXT_RUNS_PER_ITEM, parse_workbook_with_package};

fn rich_package(items: &str, inline: &str) -> crate::ParsedWorkbook {
    let mut parts = package(
        &format!(
            r#"<sheetData><row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1" t="inlineStr"><is>{inline}</is></c></row></sheetData>"#
        ),
        &[],
        false,
    );
    parts.push((
        "xl/sharedStrings.xml".into(),
        format!("<sst>{items}</sst>").into_bytes(),
    ));
    parse_workbook_with_package(&parts).unwrap()
}

#[test]
fn shared_runs_keep_overrides_and_utf8_byte_offsets() {
    let parsed = rich_package(
        concat!(
            r#"<si><r><rPr><rFont val="Aptos"/><sz val="14"/><b/><i/><strike val="0"/><u val="doubleAccounting"/><color rgb="FF12AB34"/><vertAlign val="superscript"/></rPr><t xml:space="preserve"> é&amp; </t></r>"#,
            r#"<r><rPr><b val="0"/><i val="false"/><strike/><u val="none"/><color theme="4" tint="-0.25"/><vertAlign val="subscript"/></rPr><t>🙂</t></r><r><t>!</t></r></si>"#,
        ),
        "<t>plain</t>",
    );
    let text = &parsed.workbook.shared_strings[0];
    assert_eq!(text, " é& 🙂!");
    let runs = parsed.package.shared_string_text_runs(0).unwrap();
    assert_eq!(runs.len(), 3);
    assert_eq!((runs[0].start, runs[0].end), (0, 5));
    assert_eq!((runs[1].start, runs[1].end), (5, 9));
    assert_eq!((runs[2].start, runs[2].end), (9, 10));
    assert_eq!(&text[runs[1].start as usize..runs[1].end as usize], "🙂");
    let style = &runs[0].style;
    assert_eq!(style.name.as_deref(), Some("Aptos"));
    assert_eq!(style.size_pt, Some(14.0));
    assert_eq!(style.bold, Some(true));
    assert_eq!(style.italic, Some(true));
    assert_eq!(style.strike, Some(false));
    assert_eq!(style.underline, Some(RunUnderline::DoubleAccounting));
    assert_eq!(style.color, Some(Color::Rgb("#12ab34".into())));
    assert_eq!(style.baseline, Some(RunBaseline::Superscript));
    let style = &runs[1].style;
    assert_eq!(style.bold, Some(false));
    assert_eq!(style.italic, Some(false));
    assert_eq!(style.strike, Some(true));
    assert_eq!(style.underline, Some(RunUnderline::None));
    assert_eq!(
        style.color,
        Some(Color::Theme {
            idx: 4,
            tint: -0.25
        })
    );
    assert_eq!(style.baseline, Some(RunBaseline::Subscript));
    assert_eq!(runs[2].style.bold, None);
    assert_eq!(runs[2].style.baseline, None);
    assert!(
        parsed
            .package
            .inline_string_text_runs(0, CellRef::new(0, 1))
            .is_none()
    );
}

#[test]
fn inline_runs_preserve_entities_cdata_and_whitespace() {
    let parsed = rich_package(
        "<si><t>plain</t></si>",
        r#"<r><rPr><u/><color indexed="7"/><vertAlign val="baseline"/></rPr><t xml:space="preserve"> A&#x1F642; </t></r><r><t><![CDATA[<&>]]></t></r>"#,
    );
    let cell = parsed.workbook.sheets[0].cell(CellRef::new(0, 1)).unwrap();
    let CellValue::Text { value } = &cell.value else {
        panic!("expected text")
    };
    assert_eq!(value, " A🙂 <&>");
    let runs = parsed
        .package
        .inline_string_text_runs(0, CellRef::new(0, 1))
        .unwrap();
    assert_eq!((runs[0].start, runs[0].end), (0, 7));
    assert_eq!((runs[1].start, runs[1].end), (7, 10));
    assert_eq!(runs[0].style.underline, Some(RunUnderline::Single));
    assert_eq!(runs[0].style.color, Some(Color::Indexed(7)));
    assert_eq!(runs[0].style.baseline, Some(RunBaseline::Baseline));
    assert!(parsed.package.shared_string_text_runs(0).is_none());
}

#[test]
fn inline_runs_follow_implicit_cell_addresses() {
    let parsed = parse_workbook_with_package(&package(
        concat!(
            r#"<sheetData><row><c t="inlineStr"><is><r><rPr><b/></rPr><t>x</t></r></is></c>"#,
            r#"<c t="inlineStr"><is><r><rPr><i/></rPr><t>y</t></r></is></c></row>"#,
            r#"<row r="3"><c/><c t="inlineStr"><is><r><rPr><u/></rPr><t>z</t></r></is></c></row></sheetData>"#,
        ),
        &[],
        false,
    ))
    .unwrap();
    let runs = |row, col| {
        parsed
            .package
            .inline_string_text_runs(0, CellRef::new(row, col))
            .unwrap()[0]
            .style
            .clone()
    };
    assert_eq!(runs(0, 0).bold, Some(true));
    assert_eq!(runs(0, 1).italic, Some(true));
    assert_eq!(runs(2, 1).underline, Some(RunUnderline::Single));
    assert!(
        parsed
            .package
            .inline_string_text_runs(0, CellRef::new(2, 0))
            .is_none()
    );
}

#[test]
fn phonetic_text_stays_in_the_value_outside_the_base_runs() {
    let parsed = rich_package(
        r#"<si><r><rPr><b/></rPr><t>漢</t></r><rPh sb="0" eb="1"><t>kan</t></rPh><phoneticPr fontId="1" type="noConversion"/></si>"#,
        r#"<r><rPr><i/></rPr><t>字</t></r><rPh sb="0" eb="1"><t>ji</t></rPh><phoneticPr fontId="1"/>"#,
    );
    assert_eq!(parsed.workbook.shared_strings[0], "漢kan");
    let runs = parsed.package.shared_string_text_runs(0).unwrap();
    assert_eq!(runs.len(), 1);
    assert_eq!((runs[0].start, runs[0].end), (0, 3));
    assert_eq!(runs[0].style.bold, Some(true));
    assert_eq!(
        parsed.workbook.sheets[0]
            .cell(CellRef::new(0, 1))
            .unwrap()
            .value,
        CellValue::Text {
            value: "字ji".into()
        }
    );
    let runs = parsed
        .package
        .inline_string_text_runs(0, CellRef::new(0, 1))
        .unwrap();
    assert_eq!(runs.len(), 1);
    assert_eq!((runs[0].start, runs[0].end), (0, 3));
    assert_eq!(runs[0].style.italic, Some(true));
}

#[test]
fn inter_run_whitespace_keeps_legacy_plain_text() {
    let parsed = rich_package("<si>\n<r><t>x</t></r>\n</si>", "<t/>");
    assert_eq!(parsed.workbook.shared_strings, ["\nx\n"]);
    assert!(parsed.package.shared_string_text_runs(0).is_none());
}

#[test]
fn unmodelled_properties_keep_the_run_and_its_known_properties() {
    for property in [
        r#"<family val="2"/>"#,
        r#"<charset val="129"/>"#,
        r#"<scheme val="minor"/>"#,
        r#"<outline/>"#,
        r#"<shadow val="1"/>"#,
        r#"<condense val="0"/>"#,
        r#"<extend val="0"/>"#,
        r#"<unknown kind="future"/>"#,
    ] {
        let parsed = rich_package(
            &format!(
                r#"<si><r><t>first</t></r><r><rPr><b/>{property}<sz val="9"/></rPr><t>last</t></r></si>"#
            ),
            &format!(r#"<r><rPr>{property}<i/></rPr><t>inline</t></r>"#),
        );
        let runs = parsed
            .package
            .shared_string_text_runs(0)
            .unwrap_or_else(|| panic!("{property}"));
        assert_eq!(runs.len(), 2, "{property}");
        assert_eq!((runs[1].start, runs[1].end), (5, 9), "{property}");
        assert_eq!(runs[1].style.bold, Some(true), "{property}");
        assert_eq!(runs[1].style.size_pt, Some(9.0), "{property}");
        let runs = parsed
            .package
            .inline_string_text_runs(0, CellRef::new(0, 1))
            .unwrap_or_else(|| panic!("{property}"));
        assert_eq!(runs[0].style.italic, Some(true), "{property}");
    }
}

#[test]
fn malformed_or_repeated_known_properties_discard_the_whole_item() {
    for property in [
        r#"<sz val="NaN"/>"#,
        r#"<sz val="0"/>"#,
        r#"<b val="maybe"/>"#,
        r#"<u val="wavy"/>"#,
        r#"<vertAlign val="raised"/>"#,
        r#"<color theme="2" tint="inf"/>"#,
        r#"<color rgb="invalid"/>"#,
        r#"<b/><b/>"#,
        r#"<rFont val="Aptos"/><rFont val="Calibri"/>"#,
    ] {
        let parsed = rich_package(
            &format!(
                "<si><r><t>first</t></r><r><rPr>{property}</rPr><t>last</t></r></si><si><r><t>ok</t></r></si>"
            ),
            "<t/>",
        );
        assert_eq!(parsed.workbook.shared_strings[0], "firstlast");
        assert!(
            parsed.package.shared_string_text_runs(0).is_none(),
            "{property}"
        );
        assert!(parsed.package.shared_string_text_runs(1).is_some());
    }
}

#[test]
fn duplicate_plain_values_keep_distinct_run_styles_and_cached_slices() {
    let parsed = rich_package(
        "<si><r><rPr><b/></rPr><t>x</t></r></si><si><r><rPr><i/></rPr><t>x</t></r></si>",
        "<t/>",
    );
    let first = parsed.package.shared_string_text_runs(0).unwrap();
    assert_eq!(first[0].style.bold, Some(true));
    assert_eq!(
        parsed.package.shared_string_text_runs(1).unwrap()[0]
            .style
            .italic,
        Some(true)
    );
    assert!(std::ptr::eq(
        first,
        parsed.package.shared_string_text_runs(0).unwrap()
    ));
}

#[test]
fn per_item_budget_exhaustion_discards_only_the_item() {
    let item = "<r><t>x</t></r>".repeat(MAX_RICH_TEXT_RUNS_PER_ITEM + 1);
    let parsed = rich_package(
        &format!("<si>{item}</si><si><r><t>later</t></r></si>"),
        "<r><t>inline</t></r>",
    );
    assert_eq!(
        parsed.workbook.shared_strings[0].len(),
        MAX_RICH_TEXT_RUNS_PER_ITEM + 1
    );
    assert!(parsed.package.shared_string_text_runs(0).is_none());
    assert!(parsed.package.shared_string_text_runs(1).is_some());
    assert!(
        parsed
            .package
            .inline_string_text_runs(0, CellRef::new(0, 1))
            .is_some()
    );
}

#[test]
fn aggregate_budget_exhaustion_keeps_plain_text_and_prior_complete_items() {
    let item = format!(
        "<si>{}</si>",
        "<r><t>x</t></r>".repeat(MAX_RICH_TEXT_RUNS_PER_ITEM)
    );
    let complete_items = crate::MAX_RICH_TEXT_RUNS / MAX_RICH_TEXT_RUNS_PER_ITEM;
    let items = format!(
        "{}<si><r><t>later</t></r></si>",
        item.repeat(complete_items + 1)
    );
    let parsed = rich_package(&items, "<r><t>inline</t></r>");
    assert_eq!(parsed.workbook.shared_strings.len(), complete_items + 2);
    assert!(
        parsed
            .package
            .shared_string_text_runs(complete_items - 1)
            .is_some()
    );
    assert!(
        parsed
            .package
            .shared_string_text_runs(complete_items)
            .is_none()
    );
    assert!(
        parsed
            .package
            .shared_string_text_runs(complete_items + 1)
            .is_none()
    );
    assert!(
        parsed
            .package
            .inline_string_text_runs(0, CellRef::new(0, 1))
            .is_none()
    );
}

#[test]
fn indexing_runs_leaves_all_source_parts_byte_identical_on_save() {
    let parsed = rich_package(
        "<si><r><rPr><b/></rPr><t>shared</t></r></si>",
        "<r><rPr><i/></rPr><t>inline</t></r>",
    );
    assert!(parsed.package.shared_string_text_runs(0).is_some());
    assert!(
        parsed
            .package
            .inline_string_text_runs(0, CellRef::new(0, 1))
            .is_some()
    );
    let saved =
        crate::write::serialize_workbook_with_package(&parsed.workbook, &parsed.package).unwrap();
    for (path, bytes) in saved {
        assert_eq!(
            bytes.as_slice(),
            parsed.package.part_bytes(&path).unwrap(),
            "{path}"
        );
    }
}

#[test]
fn underline_kinds_and_automatic_color_are_retained() {
    for (value, expected) in [
        ("single", RunUnderline::Single),
        ("double", RunUnderline::Double),
        ("singleAccounting", RunUnderline::SingleAccounting),
        ("doubleAccounting", RunUnderline::DoubleAccounting),
        ("none", RunUnderline::None),
    ] {
        let parsed = rich_package(
            &format!(r#"<si><r><rPr><u val="{value}"/><color auto="1"/></rPr><t>x</t></r></si>"#),
            "<t/>",
        );
        let runs = parsed.package.shared_string_text_runs(0).unwrap();
        assert_eq!(runs[0].style.underline, Some(expected));
        assert_eq!(runs[0].style.color, Some(Color::Auto));
    }
}

#[test]
fn source_clones_share_the_index_even_when_cloned_before_first_query() {
    let parsed = rich_package("<si><r><t>shared</t></r></si>", "<t/>");
    let clone = parsed.package.clone();
    assert!(std::ptr::eq(
        parsed.package.shared_string_text_runs(0).unwrap(),
        clone.shared_string_text_runs(0).unwrap()
    ));
}

#[test]
fn package_run_budget_keeps_runs_first_come_and_counts_unmodelled_properties() {
    let per_item = 4_000;
    let full_items = crate::MAX_RICH_TEXT_RUNS / per_item;
    assert_eq!(full_items * per_item, crate::MAX_RICH_TEXT_RUNS);
    let family = r#"<r><rPr><family val="2"/><b/></rPr><t>x</t></r>"#;
    let plain = "<r><t>x</t></r>";
    let last = full_items - 1;
    let tail = full_items;
    for extra in [0, 1] {
        let mut items = String::new();
        for item in 0..full_items {
            let run = if item % 2 == 0 { family } else { plain };
            let count = if item == last {
                per_item + extra
            } else {
                per_item
            };
            items.push_str(&format!("<si>{}</si>", run.repeat(count)));
        }
        items.push_str("<si><r><rPr><i/></rPr><t>ta</t></r><r><t>il</t></r></si>");
        let parsed = rich_package(&items, "<r><t>inline</t></r>");
        for item in 0..last {
            let runs = parsed.package.shared_string_text_runs(item).unwrap();
            assert_eq!(runs.len(), per_item, "{item}");
        }
        let first = parsed.package.shared_string_text_runs(0).unwrap();
        assert_eq!(first[0].style.bold, Some(true));
        assert_eq!(
            parsed.workbook.shared_strings[last],
            "x".repeat(per_item + extra)
        );
        assert_eq!(
            parsed.package.shared_string_text_runs(last).is_some(),
            extra == 0
        );
        assert_eq!(parsed.workbook.shared_strings[tail], "tail");
        assert!(parsed.package.shared_string_text_runs(tail).is_none());
        assert_eq!(
            parsed.workbook.sheets[0]
                .cell(CellRef::new(0, 1))
                .unwrap()
                .value,
            CellValue::Text {
                value: "inline".into()
            }
        );
        assert!(
            parsed
                .package
                .inline_string_text_runs(0, CellRef::new(0, 1))
                .is_none()
        );
    }
}
