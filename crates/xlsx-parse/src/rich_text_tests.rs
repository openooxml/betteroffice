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
fn phonetics_and_inter_run_whitespace_keep_legacy_plain_text() {
    let parsed = rich_package(
        "<si><r><t>漢</t></r><rPh sb=\"0\" eb=\"1\"><t>kan</t></rPh></si><si>\n<r><t>x</t></r>\n</si>",
        "<r><t>字</t></r><rPh><t>ji</t></rPh>",
    );
    assert_eq!(parsed.workbook.shared_strings, ["漢kan", "\nx\n"]);
    assert!(parsed.package.shared_string_text_runs(0).is_none());
    assert!(parsed.package.shared_string_text_runs(1).is_none());
    assert_eq!(
        parsed.workbook.sheets[0]
            .cell(CellRef::new(0, 1))
            .unwrap()
            .value,
        CellValue::Text {
            value: "字ji".into()
        }
    );
    assert!(
        parsed
            .package
            .inline_string_text_runs(0, CellRef::new(0, 1))
            .is_none()
    );
}

#[test]
fn unsupported_or_invalid_styles_discard_the_whole_item() {
    for property in [
        r#"<sz val="NaN"/>"#,
        r#"<sz val="0"/>"#,
        r#"<b val="maybe"/>"#,
        r#"<u val="wavy"/>"#,
        r#"<vertAlign val="raised"/>"#,
        r#"<color theme="2" tint="inf"/>"#,
        r#"<color rgb="invalid"/>"#,
        r#"<outline/>"#,
        r#"<unknown/>"#,
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
