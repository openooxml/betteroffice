use std::collections::BTreeMap;

use xlsx_model::styles::Xf;
use xlsx_model::{CellRange, CellRef, SheetId, Workbook};

use super::StyleMatch;
use crate::tests::{package, parse_workbook_with_package};
use crate::{
    ParsedWorkbook, SaveEdits, SheetAxes,
    serialize_workbook_with_package_and_origins_after_edits_and_active_sheet_with_axes,
};

/// cellXfs[2] differs from cellXfs[1] only in protection; cellXfs[3] adds bold.
const STYLES: &str = concat!(
    r#"<styleSheet><fonts count="2"><font><sz val="11"/></font><font><b/><sz val="11"/></font></fonts>"#,
    r#"<fills count="1"><fill><patternFill patternType="none"/></fill></fills>"#,
    r#"<borders count="1"><border/></borders>"#,
    r#"<cellXfs count="4"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>"#,
    r#"<xf numFmtId="2" fontId="0" fillId="0" borderId="0" applyNumberFormat="1"/>"#,
    r#"<xf numFmtId="2" fontId="0" fillId="0" borderId="0" applyNumberFormat="1" applyProtection="1"><protection locked="0"/></xf>"#,
    r#"<xf numFmtId="2" fontId="1" fillId="0" borderId="0" applyNumberFormat="1" applyFont="1"/>"#,
    r#"</cellXfs></styleSheet>"#,
);

const SHEET: &str = concat!(
    r#"<sheetData><row r="1"><c r="A1" s="2"><v>1</v></c><c r="B1" s="2"><v>2</v></c>"#,
    r#"<c r="C1" s="2"><f t="shared" ref="C1:C2" si="0">B1*2</f><v>4</v></c></row>"#,
    r#"<row r="2"><c r="A2" s="1"><v>3</v></c><c r="B2" s="2"><v>5</v></c>"#,
    r#"<c r="C2" s="2"><f t="shared" si="0"/><v>10</v></c></row></sheetData>"#,
);

const ARRAY_SHEET: &str = concat!(
    r#"<sheetData><row r="1"><c r="A1" s="2"><f t="array" ref="A1:A2">1</f><v>1</v></c>"#,
    r#"<c r="B1" s="2"><v>2.0</v></c><c r="C1" s="1"><v>3</v></c></row>"#,
    r#"<row r="2"><c r="A2" s="2"><v>1.0</v></c></row></sheetData>"#,
);

fn parsed() -> ParsedWorkbook {
    parsed_sheet(SHEET)
}

pub(super) fn parsed_sheet(sheet: &str) -> ParsedWorkbook {
    let mut parts = package(sheet, &[], false);
    parts.push(("xl/styles.xml".to_owned(), STYLES.as_bytes().to_vec()));
    parse_workbook_with_package(&parts).unwrap()
}

/// The model as an open projects it: cells on cellXfs[2] move to the first
/// equivalent entry.
pub(super) fn projected(parsed: &ParsedWorkbook) -> Workbook {
    let mut workbook = parsed.workbook.clone();
    let sheet = &mut workbook.sheets[0];
    let moved = sheet
        .iter_cells()
        .filter(|(_, cell)| cell.style == Some(2))
        .map(|(at, cell)| (at, cell.clone()))
        .collect::<Vec<_>>();
    for (at, mut cell) in moved {
        cell.style = Some(1);
        sheet.set_cell(at, cell);
    }
    workbook
}

pub(super) fn edit(
    workbook: &mut Workbook,
    address: &str,
    edit: impl FnOnce(&mut xlsx_model::Cell),
) {
    let at = xlsx_model::CellRef::parse_a1(address).unwrap();
    let mut cell = workbook.sheets[0].cell(at).cloned().unwrap();
    edit(&mut cell);
    workbook.sheets[0].set_cell(at, cell);
}

pub(super) fn save(
    parsed: &ParsedWorkbook,
    workbook: &Workbook,
    axes: Option<SheetAxes>,
) -> String {
    let provenance = vec![parsed.package.source_shared_string_cells(0)];
    let saved = serialize_workbook_with_package_and_origins_after_edits_and_active_sheet_with_axes(
        workbook,
        &parsed.package,
        &[Some(0)],
        &provenance,
        &[axes],
        SaveEdits {
            changed: true,
            moved_references: false,
        },
        SheetId(0),
    )
    .unwrap();
    let (_, bytes) = saved
        .iter()
        .find(|(path, _)| path == "xl/worksheets/sheet1.xml")
        .unwrap();
    String::from_utf8(bytes.to_vec()).unwrap()
}

fn reopen(xml: &str) -> Workbook {
    let mut parts = package("", &[], false);
    parts
        .iter_mut()
        .find(|(path, _)| path == "xl/worksheets/sheet1.xml")
        .unwrap()
        .1 = xml.as_bytes().to_vec();
    parts.push(("xl/styles.xml".to_owned(), STYLES.as_bytes().to_vec()));
    parse_workbook_with_package(&parts).unwrap().workbook
}

/// Each `<c>` element's markup keyed by its `r` attribute.
pub(super) fn cells(xml: &str) -> BTreeMap<String, String> {
    let mut out = BTreeMap::new();
    let mut rest = xml;
    while let Some(start) = rest.find("<c ") {
        let tail = &rest[start..];
        let tag_end = tail.find('>').unwrap();
        let end = if tail[..tag_end].ends_with('/') {
            tag_end + 1
        } else {
            tail.find("</c>").unwrap() + "</c>".len()
        };
        let span = &tail[..end];
        let address = span
            .split("r=\"")
            .nth(1)
            .unwrap()
            .split('"')
            .next()
            .unwrap();
        out.insert(address.to_owned(), span.to_owned());
        rest = &tail[end..];
    }
    out
}

fn style_attribute(span: &str) -> Option<&str> {
    let tag = &span[..span.find('>').unwrap()];
    tag.split(" s=\"")
        .nth(1)
        .map(|rest| rest.split('"').next().unwrap())
}

#[test]
fn equivalence_needs_the_same_resolved_format_in_both_stylesheets() {
    let parsed = parsed();
    let original = &parsed.workbook.styles;
    let styles = StyleMatch::new(original, original);
    assert!(styles.equivalent(Some(2), Some(1)));
    assert!(styles.equivalent(Some(1), Some(2)));
    assert!(styles.equivalent(None, None));
    assert!(!styles.equivalent(Some(2), Some(3)));
    assert!(!styles.equivalent(Some(0), None));
    assert!(!styles.equivalent(None, Some(0)));
    assert!(!styles.equivalent(Some(9), Some(0)));
    assert!(!styles.equivalent(Some(0), Some(9)));
    assert_eq!(styles.written(Some(2), Some(1)), Some(2));
    assert_eq!(styles.written(Some(2), Some(3)), Some(3));

    let mut changed = original.clone();
    changed.cell_xfs[2] = Xf {
        font: Some(1),
        ..changed.cell_xfs[2].clone()
    };
    let styles = StyleMatch::new(original, &changed);
    assert!(!styles.equivalent(Some(2), Some(1)));
    assert_eq!(styles.written(Some(2), Some(1)), Some(1));

    let mut shortened = original.clone();
    shortened.cell_xfs.truncate(2);
    let styles = StyleMatch::new(original, &shortened);
    assert!(!styles.equivalent(Some(2), Some(1)));
}

#[test]
fn projected_indices_alone_leave_the_sheet_untouched() {
    let parsed = parsed();
    let workbook = projected(&parsed);
    assert_ne!(workbook, parsed.workbook);

    let saved = save(&parsed, &workbook, Some(SheetAxes::default()));

    assert_eq!(saved, format!("<worksheet>{SHEET}</worksheet>"));
}

#[test]
fn changed_source_xf_is_not_equivalent_even_when_its_resolved_format_matches() {
    let parsed = parsed();
    let mut workbook = projected(&parsed);
    workbook.styles.cell_xfs[2].alignment = Some(Default::default());
    assert_ne!(
        parsed.workbook.styles.cell_xfs[2],
        workbook.styles.cell_xfs[2]
    );
    assert_eq!(
        parsed.workbook.styles.resolved_format(Some(2)),
        workbook.styles.resolved_format(Some(2))
    );
    assert_eq!(
        workbook.styles.resolved_format(Some(2)),
        workbook.styles.resolved_format(Some(1))
    );
    let styles = StyleMatch::new(&parsed.workbook.styles, &workbook.styles);
    assert!(!styles.equivalent(Some(2), Some(1)));

    let saved = cells(&save(&parsed, &workbook, Some(SheetAxes::default())));

    assert_eq!(saved["A1"], r#"<c r="A1" s="1"><v>1</v></c>"#);
}

#[test]
fn moved_cells_do_not_borrow_equivalent_indices_from_the_same_position() {
    let source = concat!(
        r#"<sheetData><row r="1"><c r="A1" s="2"><v>1</v></c></row>"#,
        r#"<row r="2"><c r="A2" s="1"><v>1</v></c></row></sheetData>"#,
    );
    let parsed = parsed_sheet(source);
    let mut workbook = projected(&parsed);
    let a1 = CellRef::new(0, 0);
    let a2 = CellRef::new(1, 0);
    let inserted = workbook.sheets[0].cell(a1).cloned().unwrap();
    let surviving = parsed.workbook.sheets[0].cell(a2).cloned().unwrap();
    workbook.sheets[0].set_cell(a1, surviving);
    workbook.sheets[0].set_cell(a2, inserted);
    let mut axes = SheetAxes::default();
    axes.rows.delete(0, 1);
    axes.rows.insert(1, 1);
    let styles = StyleMatch::new(&parsed.workbook.styles, &workbook.styles);
    assert!(
        workbook.sheets[0]
            .iter_cells()
            .zip(parsed.workbook.sheets[0].iter_cells())
            .all(|((at, cell), (source, original))| {
                at == source && styles.same_cell(original, cell)
            })
    );

    let saved = save(&parsed, &workbook, Some(axes));

    assert_ne!(saved, format!("<worksheet>{source}</worksheet>"));
    assert_eq!(cells(&saved)["A1"], r#"<c r="A1" s="1"><v>1</v></c>"#);
}

#[test]
fn ordinary_formula_promotion_keeps_the_array_rectangle() {
    let source = concat!(
        r#"<sheetData><row r="1"><c r="A1" s="2"><f>1</f><v>1</v></c></row>"#,
        r#"<row r="2"><c r="A2" s="2"><v>1</v></c></row></sheetData>"#,
    );
    let parsed = parsed_sheet(source);
    let mut workbook = projected(&parsed);
    let at = CellRef::new(0, 0);
    let range = CellRange::new(at, CellRef::new(1, 0));
    workbook.sheets[0].set_array_formula(at, range);

    let saved = save(&parsed, &workbook, Some(SheetAxes::default()));

    assert_eq!(
        cells(&saved)["A1"],
        r#"<c r="A1" s="2"><f t="array" ref="A1:A2">1</f><v>1</v></c>"#
    );
    let reopened = reopen(&saved);
    assert_eq!(reopened.sheets[0].array_formula(at), Some(range));
    assert_eq!(
        reopened.sheets[0].cell(at).unwrap(),
        parsed.workbook.sheets[0].cell(at).unwrap()
    );
}

#[test]
fn shared_formula_promotion_rewrites_the_whole_group() {
    let source = concat!(
        r#"<sheetData><row r="1"><c r="B1"><v>1</v></c>"#,
        r#"<c r="C1" s="2"><f t="shared" ref="C1:C3" si="0">B1*2</f><v>2</v></c></row>"#,
        r#"<row r="2"><c r="B2"><v>2</v></c><c r="C2" s="2"><f t="shared" si="0"/><v>4</v></c></row>"#,
        r#"<row r="3"><c r="B3"><v>3</v></c><c r="C3" s="2"><f t="shared" si="0"/><v>6</v></c></row></sheetData>"#,
    );
    for address in ["C1", "C2"] {
        let parsed = parsed_sheet(source);
        let mut workbook = projected(&parsed);
        let at = CellRef::parse_a1(address).unwrap();
        let range = CellRange::new(at, CellRef::new(at.row + 1, at.col));
        workbook.sheets[0].set_array_formula(at, range);

        let saved = save(&parsed, &workbook, Some(SheetAxes::default()));

        assert!(
            cells(&saved)[address].contains(&format!(r#"<f t="array" ref="{}">"#, range.to_a1()))
        );
        assert!(!saved.contains(r#"t="shared""#));
        assert!(!saved.contains(r#"si=""#));
        let reopened = reopen(&saved);
        assert_eq!(reopened.sheets[0].array_formula(at), Some(range));
        for address in ["C1", "C2", "C3"] {
            let at = CellRef::parse_a1(address).unwrap();
            assert_eq!(
                reopened.sheets[0].cell(at).unwrap(),
                parsed.workbook.sheets[0].cell(at).unwrap(),
                "{address}"
            );
        }
    }
}

#[test]
fn equal_cells_with_nonidentity_axes_keep_surviving_row_attributes() {
    let source = concat!(
        r#"<sheetData><row r="1" outlineLevel="1"><c r="A1" s="2"><v>1</v></c></row>"#,
        r#"<row r="2" outlineLevel="2"><c r="A2" s="2"><v>1</v></c></row></sheetData>"#,
    );
    let parsed = parsed_sheet(source);
    let mut workbook = parsed.workbook.clone();
    let inserted = workbook.sheets[0]
        .cell(CellRef::new(0, 0))
        .cloned()
        .unwrap();
    workbook.sheets[0].remap_cells(|at| (at.row > 0).then(|| CellRef::new(at.row - 1, at.col)));
    workbook.sheets[0].set_cell(CellRef::new(1, 0), inserted);
    let mut axes = SheetAxes::default();
    axes.rows.delete(0, 1);
    axes.rows.insert(1, 1);
    assert!(!axes.is_identity());
    assert_eq!(workbook, parsed.workbook);

    let saved = save(&parsed, &workbook, Some(axes));

    assert_eq!(
        saved,
        concat!(
            r#"<worksheet><sheetData><row r="1" outlineLevel="2"><c r="A1" s="2"><v>1</v></c></row>"#,
            r#"<row r="2"><c r="A2" s="2"><v>1</v></c></row></sheetData></worksheet>"#,
        )
    );
    assert_eq!(
        save(&parsed, &workbook, None),
        format!("<worksheet>{source}</worksheet>")
    );
}

#[test]
fn array_rectangle_changes_alone_do_not_borrow_the_source_sheet() {
    let parsed = parsed_sheet(ARRAY_SHEET);
    let mut workbook = projected(&parsed);
    let at = CellRef::new(0, 0);
    workbook.sheets[0].set_array_formula(at, CellRange::new(at, at));

    let saved = save(&parsed, &workbook, Some(SheetAxes::default()));

    assert!(saved.contains(r#"<f t="array" ref="A1">1</f>"#));
    assert!(!saved.contains(r#"ref="A1:A2""#));
}

#[test]
fn patched_array_master_uses_the_current_rectangle_and_preserves_untouched_cells() {
    let parsed = parsed_sheet(ARRAY_SHEET);
    let mut workbook = projected(&parsed);
    let at = CellRef::new(0, 0);
    workbook.sheets[0].set_array_formula(at, CellRange::new(at, at));
    edit(&mut workbook, "C1", |cell| {
        cell.value = xlsx_model::CellValue::Number { value: 9.0 }
    });

    let saved = save(&parsed, &workbook, Some(SheetAxes::default()));

    assert!(saved.contains(r#"<f t="array" ref="A1">1</f>"#));
    assert!(!saved.contains(r#"ref="A1:A2""#));
    let saved = cells(&saved);
    let source = cells(ARRAY_SHEET);
    assert_eq!(saved["C1"], r#"<c r="C1" s="1"><v>9</v></c>"#);
    for address in ["B1", "A2"] {
        assert_eq!(saved[address], source[address], "{address}");
    }
}

#[test]
fn rewritten_cells_keep_their_equivalent_source_index() {
    let parsed = parsed();
    let mut workbook = projected(&parsed);
    edit(&mut workbook, "A1", |cell| {
        cell.value = xlsx_model::CellValue::Number { value: 9.0 }
    });
    edit(&mut workbook, "B1", |cell| cell.style = Some(3));

    let saved = cells(&save(&parsed, &workbook, Some(SheetAxes::default())));
    let source = cells(SHEET);

    assert_eq!(saved["A1"], r#"<c r="A1" s="2"><v>9</v></c>"#);
    assert_eq!(saved["B1"], r#"<c r="B1" s="3"><v>2</v></c>"#);
    for address in ["C1", "A2", "B2", "C2"] {
        assert_eq!(saved[address], source[address], "{address}");
    }
}

#[test]
fn regenerated_sheet_data_keeps_equivalent_source_indices() {
    let source = SHEET.replace(r#"<row r="1">"#, "<row>");
    assert!(super::scan_sheet_data(source.as_bytes()).unwrap().is_none());
    let parsed = parsed_sheet(&source);
    let mut workbook = projected(&parsed);
    edit(&mut workbook, "A1", |cell| {
        cell.value = xlsx_model::CellValue::Number { value: 9.0 }
    });
    edit(&mut workbook, "B1", |cell| cell.style = Some(3));

    let saved = cells(&save(&parsed, &workbook, Some(SheetAxes::default())));

    let styles = saved
        .iter()
        .map(|(address, span)| (address.as_str(), style_attribute(span)))
        .collect::<Vec<_>>();
    assert_eq!(
        styles,
        [
            ("A1", Some("2")),
            ("A2", Some("1")),
            ("B1", Some("3")),
            ("B2", Some("2")),
            ("C1", Some("2")),
            ("C2", Some("2")),
        ]
    );
}

#[test]
fn regenerated_sheet_data_without_axes_uses_model_indices() {
    let parsed = parsed();
    let mut workbook = projected(&parsed);
    edit(&mut workbook, "A1", |cell| {
        cell.value = xlsx_model::CellValue::Number { value: 9.0 }
    });
    edit(&mut workbook, "B1", |cell| cell.style = Some(3));

    let saved = cells(&save(&parsed, &workbook, None));

    let styles = saved
        .iter()
        .map(|(address, span)| (address.as_str(), style_attribute(span)))
        .collect::<Vec<_>>();
    assert_eq!(
        styles,
        [
            ("A1", Some("1")),
            ("A2", Some("1")),
            ("B1", Some("3")),
            ("B2", Some("1")),
            ("C1", Some("1")),
            ("C2", Some("1")),
        ]
    );
}
