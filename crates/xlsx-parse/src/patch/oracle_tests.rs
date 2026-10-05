use std::path::{Path, PathBuf};

use xlsx_model::{CellValue, ErrorValue};

use super::*;
use crate::axis::AxisMap;
use crate::package::XmlTemplate;
use crate::tests::{package, parse_workbook_with_package};

const CELLS: &str = concat!(
    r#"<sheetData><row r="1" spans="1:1" hidden="1">"#,
    r#"<c r="A1" cm="1"><v>1.00</v></c>"#,
    r#"<c r="B1" t="inlineStr"><is><r><rPr><b/></rPr><t>Rich</t></r><r><t xml:space="preserve"> inline</t></r></is></c>"#,
    r#"<c r="C1" t="s"><v>0</v></c><c r="D1" t="s"><v>1</v></c>"#,
    r#"<c r="E1" t="b"><v>1</v></c><c r="F1" t="e"><v>#DIV/0!</v></c>"#,
    r#"<c r="G1" s="3"/><c r="H1" cm="2"/><c r="I1"><f>A1+1</f><v>2</v></c>"#,
    r#"<c r="J1" t="s"/><c r="K1" t="s"><v>2</v></c></row>"#,
    r#"<row r="2" ht="22" customHeight="1">"#,
    r#"<c r="A2"><f t="shared" si="9" ref="A2:B4">A1+1</f><v>2</v></c>"#,
    r#"<c r="B2"><f t="shared" si="9"/><v>2</v></c>"#,
    r#"<c r="C2"><f t="array" ref="C2:D4">SEQUENCE(3,2)</f><v>1</v></c>"#,
    r#"<c r="D2"><v>2</v></c><c r="E2"><f r1="1" r2="2">A1</f><v>1</v></c>"#,
    r#"<c r="F2"><f ref="invalid">A1</f><v>1</v></c></row>"#,
    r#"<row r="4"><c r="A4"><f t="shared" si="9"/><v>2</v></c>"#,
    r#"<c r="B4"><f t="shared" si="9"/><v>2</v></c>"#,
    r#"<c r="C4"><v>5</v></c><c r="D4"><v>6</v></c></row>"#,
    r#"<row r="7" ht="0" customHeight="1" hidden="1"/>"#,
    r#"<row r="12"><c r="Z12"><v>12</v></c></row></sheetData>"#,
);

struct Case {
    source: Vec<u8>,
    original: Workbook,
    workbook: Workbook,
    axes: SheetAxes,
    retained: SharedStringCells,
}

impl Case {
    fn new(source: &str) -> Self {
        Self::from_parsed(parse_workbook_with_package(&package(source, &[], false)).unwrap())
    }

    fn mixed() -> Self {
        let mut parts = package(CELLS, &["Dup", "Dup", "other"], false);
        let strings = concat!(
            r#"<sst><si><r><rPr><b/></rPr><t>Dup</t></r></si>"#,
            r#"<si><t>Dup</t></si><si><t>other</t></si></sst>"#,
        );
        parts
            .iter_mut()
            .find(|(path, _)| path == "xl/sharedStrings.xml")
            .unwrap()
            .1 = strings.as_bytes().to_vec();
        Self::from_parsed(parse_workbook_with_package(&parts).unwrap())
    }

    fn from_parsed(parsed: crate::ParsedWorkbook) -> Self {
        let source = parsed.package.sheets[0]
            .template
            .child("sheetData")
            .unwrap()
            .bytes
            .clone();
        let retained = parsed.package.source_shared_string_cells(0);
        Self {
            source,
            workbook: parsed.workbook.clone(),
            original: parsed.workbook,
            axes: SheetAxes::default(),
            retained,
        }
    }

    fn projected(source: &str) -> Self {
        let parsed = super::style_match_tests::parsed_sheet(source);
        let workbook = super::style_match_tests::projected(&parsed);
        let mut case = Self::from_parsed(parsed);
        case.workbook = workbook;
        case
    }

    fn with_patch<R>(&self, planned: bool, run: impl FnOnce(&SheetPatch<'_>) -> R) -> R {
        let sst_index = self
            .workbook
            .shared_strings
            .iter()
            .enumerate()
            .map(|(index, value)| (value.as_str(), index))
            .collect();
        let plan = planned.then(|| {
            SharedStringPlan::new(
                &self.original.shared_strings,
                self.original.shared_strings.len(),
                &self.workbook,
                std::slice::from_ref(&self.retained),
            )
            .unwrap()
        });
        let styles = StyleMatch::new(&self.original.styles, &self.workbook.styles);
        run(&SheetPatch {
            sheet: &self.workbook.sheets[0],
            original: &self.original.sheets[0],
            axes: &self.axes,
            workbook: &self.workbook,
            sst_index: &sst_index,
            retained: &self.retained,
            plan: plan.as_ref(),
            styles: &styles,
        })
    }

    fn compare_source(&self, source: &[u8]) -> Result<Option<Vec<u8>>, ParseError> {
        let mut result = Ok(None);
        for planned in [false, true] {
            result = self.with_patch(planned, |patch| {
                assert_eq!(
                    patch.changed_source_cells(),
                    patch.changed_source_cells_oracle(),
                    "changed sources with plan {planned}"
                );
                let actual = patch.sheet_data(source);
                let oracle = patch.sheet_data_oracle(source);
                assert_eq!(actual, oracle, "sheetData with plan {planned}");
                actual
            });
        }
        result
    }

    fn compare(&self) -> Vec<u8> {
        self.compare_source(&self.source).unwrap().unwrap()
    }

    fn edit(&mut self, address: &str, edit: impl FnOnce(&mut Cell)) {
        let at = CellRef::parse_a1(address).unwrap();
        let mut cell = self.workbook.sheets[0]
            .cell(at)
            .cloned()
            .unwrap_or_default();
        edit(&mut cell);
        self.workbook.sheets[0].set_cell(at, cell);
    }

    fn shift(&mut self, rows: bool, insert: bool, at: u32, count: u32) {
        let limit = if rows { MAX_ROWS } else { MAX_COLS };
        let index = |index: u32| {
            if index < at {
                Some(index)
            } else if insert {
                index.checked_add(count).filter(|&index| index < limit)
            } else if index < at + count {
                None
            } else {
                Some(index - count)
            }
        };
        let remap = |cell: CellRef| {
            Some(if rows {
                CellRef::new(index(cell.row)?, cell.col)
            } else {
                CellRef::new(cell.row, index(cell.col)?)
            })
        };
        self.workbook.sheets[0].remap_cells(remap);
        if rows {
            let heights = std::mem::take(&mut self.workbook.sheets[0].row_heights);
            self.workbook.sheets[0].row_heights = heights
                .into_iter()
                .filter_map(|(row, height)| index(row).map(|row| (row, height)))
                .collect();
        }
        self.retained = std::mem::take(&mut self.retained)
            .into_iter()
            .filter_map(|((row, col), value)| {
                remap(CellRef::new(row, col)).map(|at| ((at.row, at.col), value))
            })
            .collect();
        let axis = if rows {
            &mut self.axes.rows
        } else {
            &mut self.axes.cols
        };
        if insert {
            axis.insert(at, count);
        } else {
            axis.delete(at, count);
        }
    }
}

fn number(value: f64) -> Cell {
    Cell {
        value: CellValue::Number { value },
        ..Cell::default()
    }
}

#[test]
fn projected_styles_and_value_edits_match_oracle_without_expanding_shared_groups() {
    let source = concat!(
        r#"<sheetData><row r="1"><c r="A1" s="2"><v>1</v></c>"#,
        r#"<c r="B1" s="2"><f t="shared" ref="B1:B2" si="0">1</f><v>1</v></c></row>"#,
        r#"<row r="2"><c r="B2" s="2"><f t="shared" si="0"/><v>1</v></c></row></sheetData>"#,
    );
    let mut case = Case::projected(source);
    assert_eq!(case.compare(), source.as_bytes());
    case.edit("A1", |cell| cell.value = CellValue::Number { value: 2.0 });

    assert_eq!(
        case.compare(),
        source
            .replace("<v>1</v></c><c", "<v>2</v></c><c")
            .as_bytes()
    );
}

#[test]
fn projected_array_rectangle_edits_match_oracle() {
    let source = concat!(
        r#"<sheetData><row r="1"><c r="A1" s="2"><f t="array" ref="A1:A2">1</f><v>1</v></c></row>"#,
        r#"<row r="2"><c r="A2" s="2"><v>1</v></c></row></sheetData>"#,
    );
    let mut case = Case::projected(source);
    let at = CellRef::new(0, 0);
    case.workbook.sheets[0].set_array_formula(at, CellRange::new(at, at));

    assert_eq!(
        case.compare(),
        source.replace(r#"ref="A1:A2""#, r#"ref="A1""#).as_bytes()
    );
    case.workbook.sheets[0].clear_array_formula(at);
    assert_eq!(
        case.compare(),
        source.replace(r#" t="array" ref="A1:A2""#, "").as_bytes()
    );
}

#[test]
fn projected_formula_promotions_match_oracle() {
    let source = concat!(
        r#"<sheetData><row r="1"><c r="A1" s="2"><f>1</f><v>1</v></c>"#,
        r#"<c r="B1" s="2"><f t="shared" ref="B1:B2" si="0">1</f><v>1</v></c></row>"#,
        r#"<row r="2"><c r="B2" s="2"><f t="shared" si="0"/><v>1</v></c></row></sheetData>"#,
    );
    for address in ["A1", "B1", "B2"] {
        let mut case = Case::projected(source);
        let at = CellRef::parse_a1(address).unwrap();
        let range = CellRange::new(at, CellRef::new(at.row + 1, at.col));
        case.workbook.sheets[0].set_array_formula(at, range);

        let saved = String::from_utf8(case.compare()).unwrap();

        assert!(saved.contains(&format!(
            r#"<c r="{address}" s="2"><f t="array" ref="{}">1</f><v>1</v></c>"#,
            range.to_a1()
        )));
        if address == "A1" {
            assert!(saved.contains(r#"<f t="shared" ref="B1:B2" si="0">1</f>"#));
            assert!(saved.contains(r#"<f t="shared" si="0"/>"#));
        } else {
            assert!(!saved.contains(r#"t="shared""#));
            assert!(!saved.contains(r#"si=""#));
        }
    }
}

#[test]
fn constructed_cells_and_edits_match_oracle() {
    let mut case = Case::mixed();
    case.compare();
    for address in ["A1", "C1", "D1", "G1", "H1", "J1", "A2", "B4", "C2", "Z12"] {
        case.edit(address, |cell| *cell = Cell::default());
        case.compare();
    }
    for address in ["B1", "C1", "K1"] {
        case.edit(address, |cell| {
            cell.value = CellValue::Text {
                value: "changed text".to_owned(),
            };
        });
        case.compare();
    }
    case.edit("E1", |cell| cell.value = CellValue::Bool { value: false });
    case.compare();
    case.edit("F1", |cell| {
        cell.value = CellValue::Error {
            value: ErrorValue::NA,
        };
    });
    case.compare();
    for address in ["H1", "D2", "B2"] {
        case.edit(address, |cell| cell.style = Some(5));
        case.compare();
    }
    for address in ["I1", "A2", "B2", "C2", "E2", "F2"] {
        case.edit(address, |cell| cell.formula = Some("SUM(A1:D4)".to_owned()));
        case.compare();
        case.edit(address, |cell| cell.formula = None);
        case.compare();
    }
    case.workbook.sheets[0].row_heights.insert(6, 30.0);
    case.compare();
    case.workbook.sheets[0].row_heights.remove(&1);
    case.compare();
    case.workbook.sheets[0].row_heights.insert(8, 15.0);
    case.compare();
}

#[test]
fn shared_strings_keep_rich_and_duplicate_provenance() {
    let mut case = Case::mixed();
    case.compare();
    case.workbook.shared_strings.insert(0, "added".to_owned());
    case.compare();
    case.workbook.shared_strings.insert(0, "Dup".to_owned());
    case.compare();
    case.workbook.shared_strings.rotate_left(2);
    case.compare();
    case.retained.remove(&(0, 2));
    case.compare();
    case.edit("D1", |cell| {
        cell.value = CellValue::Text {
            value: "other".to_owned(),
        };
    });
    case.compare();
}

#[test]
fn source_only_empty_cells_remain_verbatim() {
    let source = r#"<sheetData><row r="1"><c r="A1" cm="9"/><c r="B1"><v/></c></row><row r="3" ht="20" customHeight="1"/></sheetData>"#;
    let mut case = Case::new(source);
    assert_eq!(case.compare(), source.as_bytes());
    case.edit("C1", |cell| *cell = number(3.0));
    let written = String::from_utf8(case.compare()).unwrap();
    assert!(written.contains(r#"<c r="A1" cm="9"/><c r="B1"><v/></c>"#));
    case.edit("A1", |cell| cell.style = Some(2));
    case.compare();
    case.edit("A1", |cell| *cell = Cell::default());
    case.compare();
}

#[test]
fn empty_sheet_data_and_absolute_addresses_match_oracle() {
    for source in [
        "<sheetData/>",
        "<sheetData></sheetData>",
        r#"<sheetData><row r="1"><c r="$A$1"><v>1</v></c><c r="$C1" s="2"/></row></sheetData>"#,
    ] {
        let mut case = Case::new(source);
        case.compare();
        case.edit("B2", |cell| *cell = number(2.0));
        case.workbook.sheets[0].row_heights.insert(3, 18.0);
        case.compare();
        case.shift(true, true, 0, 1);
        case.shift(false, true, 1, 1);
        case.compare();
    }
}

#[test]
fn structural_edits_match_oracle() {
    for rows in [true, false] {
        for insert in [true, false] {
            for (at, count) in [(0, 1), (1, 2), (3, 1), (6, 2), (11, 1)] {
                let mut case = Case::mixed();
                case.shift(rows, insert, at, count);
                case.compare();
                case.edit("C3", |cell| *cell = number(99.0));
                case.compare();
                case.shift(!rows, !insert, 1, 1);
                case.compare();
                case.shift(rows, !insert, at, count);
                case.compare();
            }
        }
    }
}

#[test]
fn detector_records_source_changes_without_coordinate_dirtiness() {
    let source = r#"<sheetData><row r="1"><c r="A1"><v>1</v></c></row><row r="3"><c r="C3"><v>3</v></c></row></sheetData>"#;
    let mut case = Case::new(source);
    case.shift(true, true, 0, 1);
    case.shift(false, true, 1, 1);
    case.with_patch(true, |patch| {
        assert!(patch.changed_source_cells().is_empty())
    });
    case.compare();
    for address in ["A1", "B2"] {
        case.edit(address, |cell| *cell = number(7.0));
        case.with_patch(true, |patch| {
            assert!(patch.changed_source_cells().is_empty())
        });
        case.compare();
    }
    case.edit("C3", |cell| *cell = number(8.0));
    case.with_patch(true, |patch| {
        assert_eq!(patch.changed_source_cells(), BTreeSet::from([(1, 1)]));
    });
    case.compare();
    case.shift(true, false, 1, 1);
    case.with_patch(true, |patch| {
        assert_eq!(
            patch.changed_source_cells(),
            BTreeSet::from([(0, 0), (1, 1)])
        );
    });
    case.compare();
    case.edit("D3", |cell| *cell = Cell::default());
    case.with_patch(true, |patch| {
        assert_eq!(
            patch.changed_source_cells(),
            BTreeSet::from([(0, 0), (1, 1), (2, 2)])
        );
    });
    case.compare();
}

#[test]
fn unmapped_addresses_and_clipped_axes_match_oracle() {
    let mut case = Case::mixed();
    case.axes.rows = AxisMap::identity(4);
    case.axes.cols = AxisMap::identity(8);
    case.compare();
    case.shift(true, true, 2, 2);
    case.shift(false, false, 1, 2);
    case.compare();
    case.edit("Z20", |cell| *cell = number(20.0));
    case.compare();
}

#[test]
fn non_reflexive_values_match_oracle() {
    let mut case = Case::new(r#"<sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>"#);
    for value in [f64::NAN, f64::INFINITY, f64::NEG_INFINITY, -0.0] {
        case.original.sheets[0].set_cell(CellRef::new(0, 0), number(value));
        case.workbook.sheets[0].set_cell(CellRef::new(0, 0), number(value));
        case.compare();
    }
}

#[test]
fn unordered_duplicate_and_unreadable_addresses_keep_the_fallback() {
    let case = Case::mixed();
    for source in [
        r#"<sheetData><row r="1"><c r="B1"/><c r="A1"/></row></sheetData>"#,
        r#"<sheetData><row r="1"><c r="A1"/><c r="A1"/></row></sheetData>"#,
        r#"<sheetData><row r="2"/><row r="1"/></sheetData>"#,
        r#"<sheetData><row r="1"/><row r="1"/></sheetData>"#,
        r#"<sheetData><row r="1"><c r="A2"/></row></sheetData>"#,
        r#"<sheetData><row r="1"><c r="invalid"/></row></sheetData>"#,
        r#"<sheetData><row><c r="A1"/></row></sheetData>"#,
    ] {
        assert_eq!(case.compare_source(source.as_bytes()).unwrap(), None);
    }
}

#[test]
fn malformed_xml_errors_match_oracle() {
    let case = Case::mixed();
    for source in [
        "",
        "<sheetData>",
        r#"<sheetData><row r="1"></sheetData>"#,
        r#"<sheetData><row r="&missing;"/></sheetData>"#,
        r#"<sheetData><row r="1" r="2"/></sheetData>"#,
    ] {
        assert!(case.compare_source(source.as_bytes()).is_err(), "{source}");
    }
    let nested = format!(
        "<sheetData>{}{}</sheetData>",
        "<x>".repeat(MAX_DEPTH),
        "</x>".repeat(MAX_DEPTH)
    );
    assert_eq!(
        case.compare_source(nested.as_bytes()),
        Err(ParseError::DepthExceeded)
    );
}

#[test]
fn original_cursor_falls_back_without_losing_its_position() {
    let case = Case::mixed();
    let sheet = &case.original.sheets[0];
    let mut cursor = SourceCellCursor::new(sheet, sheet.iter_cells());
    EMISSION_LOOKUPS.with(|count| count.set(0));
    for (source, current, ordered) in [
        ("A1", "A1", true),
        ("C1", "C1", true),
        ("C1", "C1", false),
        ("B1", "B1", false),
        ("H1", "H1", true),
        ("G1", "G1", false),
        ("I1", "A1", false),
        ("K1", "K1", true),
        ("A2", "A2", true),
    ] {
        let source = CellRef::parse_a1(source).unwrap();
        let current = CellRef::parse_a1(current).unwrap();
        let (actual, actual_ordered) = cursor.at(source, current);
        assert_eq!(actual, sheet.cell(source));
        assert_eq!(actual_ordered, ordered);
    }
    assert_eq!(EMISSION_LOOKUPS.with(|count| count.get()), 3);
}

fn source_row(cells: &[(&str, Option<u32>)]) -> (Vec<u8>, SourceRow) {
    let mut data = b"<row r=\"1\">".to_vec();
    let tag = 0..data.len();
    let mut source_cells = Vec::new();
    for &(address, value) in cells {
        let start = data.len();
        let empty = value.is_none();
        data.extend_from_slice(
            format!(r#"<c r="{address}"{}>"#, if empty { "/" } else { "" }).as_bytes(),
        );
        let tag = start..data.len();
        if let Some(value) = value {
            data.extend_from_slice(format!("<v>{value}</v></c>").as_bytes());
        }
        source_cells.push(SourceCell {
            before: start..start,
            span: start..data.len(),
            tag,
            empty,
            at: CellRef::parse_a1(address).unwrap(),
            shared_string: false,
            formula: None,
        });
    }
    data.extend_from_slice(b"</row>");
    let row = SourceRow {
        before: 0..0,
        span: 0..data.len(),
        tag,
        empty: false,
        index: 0,
        cells: source_cells,
    };
    (data, row)
}

#[test]
fn row_emission_uses_point_lookups_for_backward_and_duplicate_cells() {
    let mut case = Case::new(concat!(
        r#"<sheetData><row r="1"><c r="A1"><v>1</v></c>"#,
        r#"<c r="B1"><v>2</v></c><c r="C1"><v>3</v></c></row></sheetData>"#,
    ));
    for edited in [false, true] {
        if edited {
            case.edit("A1", |cell| *cell = Cell::default());
            case.edit("B1", |cell| cell.style = Some(2));
        }
        for addresses in [
            vec![
                ("C1", Some(3)),
                ("A1", Some(1)),
                ("B1", Some(2)),
                ("E1", None),
            ],
            vec![
                ("A1", Some(1)),
                ("A1", Some(1)),
                ("C1", Some(3)),
                ("E1", None),
            ],
        ] {
            let (data, row) = source_row(&addresses);
            case.with_patch(true, |patch| {
                let dirty = DirtyFormulas {
                    groups: HashSet::new(),
                    masters: HashSet::new(),
                };
                let mut current = patch.sheet.iter_cells().peekable();
                let mut original =
                    SourceCellCursor::new(patch.original, patch.original.iter_cells());
                let mut actual = Vec::new();
                EMISSION_LOOKUPS.with(|count| count.set(0));
                let result = patch.emit_source_row(
                    &mut actual,
                    &data,
                    &row,
                    0,
                    &mut current,
                    &mut original,
                    &dirty,
                );
                assert!(EMISSION_LOOKUPS.with(|count| count.get()) >= 2);
                let mut current = patch.sheet.iter_cells().peekable();
                let mut oracle = Vec::new();
                let expected =
                    patch.emit_source_row_oracle(&mut oracle, &data, &row, 0, &mut current, &dirty);
                assert_eq!(result, expected);
                assert_eq!(actual, oracle);
            });
        }
    }
}

struct Xorshift(u64);

impl Xorshift {
    fn next(&mut self) -> u32 {
        self.0 ^= self.0 << 13;
        self.0 ^= self.0 >> 7;
        self.0 ^= self.0 << 17;
        self.0 as u32
    }
}

#[test]
fn seeded_random_edits_match_oracle_after_every_step() {
    let mut case = Case::mixed();
    let mut random = Xorshift(0x4d59_5df4_d0f3_3173);
    case.compare();
    for _ in 0..400 {
        let at = CellRef::new(random.next() % 14, random.next() % 10);
        let address = at.to_a1();
        match random.next() % 15 {
            0 => case.edit(&address, |cell| *cell = Cell::default()),
            1 => case.edit(&address, |cell| {
                *cell = number(f64::from(random.next() % 101))
            }),
            2 => {
                let values = ["Dup", "other", "Rich inline", "new", ""];
                let value = values[(random.next() % 5) as usize];
                case.edit(&address, |cell| {
                    cell.value = CellValue::Text {
                        value: value.to_owned(),
                    };
                });
            }
            3 => case.edit(&address, |cell| {
                cell.value = CellValue::Bool {
                    value: random.next().is_multiple_of(2),
                };
            }),
            4 => case.edit(&address, |cell| {
                cell.value = CellValue::Error {
                    value: ErrorValue::Ref,
                };
            }),
            5 => case.edit(&address, |cell| cell.style = Some(random.next() % 6)),
            6 => case.edit(&address, |cell| {
                cell.formula = Some("SUM(A1:C3)".to_owned())
            }),
            7 => case.edit(&address, |cell| cell.formula = None),
            8 => case.shift(true, true, at.row, 1 + random.next() % 2),
            9 => case.shift(true, false, at.row, 1 + random.next() % 2),
            10 => case.shift(false, true, at.col, 1 + random.next() % 2),
            11 => case.shift(false, false, at.col, 1 + random.next() % 2),
            12 => {
                case.workbook.sheets[0]
                    .row_heights
                    .insert(at.row, f64::from(random.next() % 40));
            }
            13 => {
                case.workbook.sheets[0].row_heights.remove(&at.row);
            }
            _ => case.edit(&address, |cell| cell.style = None),
        }
        case.compare();
    }
}

fn lookup_counts() -> (usize, usize) {
    (
        EMISSION_LOOKUPS.with(|count| count.replace(0)),
        DETECTOR_LOOKUPS.with(|count| count.replace(0)),
    )
}

#[test]
fn dense_sheet_uses_at_most_one_percent_of_oracle_point_lookups() {
    let mut source = String::from("<sheetData>");
    for row in 0..300 {
        source.push_str(&format!(r#"<row r="{}">"#, row + 1));
        for col in 0..20 {
            let at = CellRef::new(row, col).to_a1();
            source.push_str(&format!(r#"<c r="{at}"><v>{}</v></c>"#, row * 20 + col));
        }
        source.push_str("</row>");
    }
    source.push_str("</sheetData>");
    let mut case = Case::new(&source);
    for (row, col) in [(0, 0), (57, 7), (149, 19), (211, 3), (299, 18)] {
        case.edit(&CellRef::new(row, col).to_a1(), |cell| *cell = number(-1.0));
    }
    case.with_patch(true, |patch| {
        assert_eq!(
            patch.changed_source_cells(),
            patch.changed_source_cells_oracle()
        );
        lookup_counts();
        let actual = patch.sheet_data(&case.source);
        let current = lookup_counts();
        let oracle = patch.sheet_data_oracle(&case.source);
        let previous = lookup_counts();
        assert_eq!(actual, oracle);
        assert!(actual.unwrap().is_some());
        assert_eq!(previous, (12_000, 12_000));
        assert!(
            current.0 <= previous.0 / 100,
            "emission: {current:?} vs {previous:?}"
        );
        assert!(
            current.1 <= previous.1 / 100,
            "detector: {current:?} vs {previous:?}"
        );
    });
}

fn xml_files(directory: &Path, files: &mut Vec<PathBuf>) {
    for entry in std::fs::read_dir(directory).unwrap() {
        let entry = entry.unwrap();
        let path = entry.path();
        if entry.file_type().unwrap().is_dir() {
            xml_files(&path, files);
        } else if path.extension().is_some_and(|extension| extension == "xml") {
            files.push(path);
        }
    }
}

#[test]
fn worksheet_xml_files_match_oracle() {
    let directory = Path::new(env!("CARGO_MANIFEST_DIR")).join("tests/fixtures");
    let mut files = Vec::new();
    xml_files(&directory, &mut files);
    files.sort();
    let mut compared = 0;
    for path in files {
        let source = std::fs::read(&path).unwrap();
        let Ok(template) = XmlTemplate::capture(&source) else {
            continue;
        };
        if template.child("sheetData").is_none() {
            continue;
        }
        let mut parts = package("", &[], false);
        parts[2].1 = source;
        let Ok(parsed) = parse_workbook_with_package(&parts) else {
            continue;
        };
        let mut case = Case::from_parsed(parsed);
        let _ = case.compare_source(&case.source);
        let addresses = case.original.sheets[0]
            .iter_cells()
            .take(5)
            .map(|(at, _)| at)
            .collect::<Vec<_>>();
        for (index, at) in addresses.iter().enumerate() {
            case.edit(&at.to_a1(), |cell| match index % 3 {
                0 => *cell = Cell::default(),
                1 => cell.style = Some(1),
                _ => *cell = number(123.0),
            });
            let _ = case.compare_source(&case.source);
        }
        case.shift(true, true, 0, 1);
        let _ = case.compare_source(&case.source);
        case.shift(false, true, 1, 1);
        let _ = case.compare_source(&case.source);
        case.shift(true, false, 1, 1);
        let _ = case.compare_source(&case.source);
        case.shift(false, false, 2, 1);
        let _ = case.compare_source(&case.source);
        compared += 1;
    }
    assert!(compared > 0);
}
