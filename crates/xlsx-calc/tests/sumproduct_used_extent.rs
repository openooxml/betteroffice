//! SUMPRODUCT over sheets whose used extent sits above, below or around the
//! ranges it reads: whole columns and rows, bounded ranges past the data,
//! ranges outside it, mismatched sizes, an empty sheet and spills that grow the
//! extent during the recalculation. every result is compared bit for bit with
//! the table in `fixtures/sumproduct_used_extent.tsv`.
//!
//! with `SUMPRODUCT_USED_EXTENT_WRITE=<path>` set, the test writes the table it
//! computes to that path instead of comparing.

use std::fmt::Write as _;

use xlsx_calc::{EvalContext, evaluate, parse_formula, rebuild_and_recalc_all};
use xlsx_model::{Cell, CellRange, CellRef, CellValue, ErrorValue, Sheet, SheetId, Workbook};

const CHUNK: usize = 120;

const RANGES: &[&str] = &[
    "Data!A1:A20",
    "Data!B1:B20",
    "Data!C1:C20",
    "Data!D1:D20",
    "Data!E1:E20",
    "Data!A1:B20",
    "Data!A1:D20",
    "Data!B1:C20",
    "Data!A1:A19",
    "Data!A1:A5",
    "Data!A2:A21",
    "Data!A100:A119",
    "Data!A1:A1000",
    "Data!A1",
    "Data!A:A",
    "Data!B:B",
    "Data!A:B",
    "Data!F:F",
    "Data!F1:F40",
    "Data!G:G",
    "Data!G1:G100",
    "Data!1:1",
    "Data!2:2",
    "Data!1:2",
    "Data!A1:T1",
    "Data!A2:T2",
];

const EXPRESSIONS: &[&str] = &[
    "--(Data!A1:A20>2)",
    "(Data!B1:B20=\"x\")*1",
    "Data!A1:A20*Data!B1:B20",
    "Data!A1:A20+Data!C1:C20",
    "Data!A1:A20-Data!A1:A19",
    "Data!A:A*1",
    "(Data!A:A>0)*Data!B:B",
    "Data!A1:A20/Data!B1:B20",
    "-Data!C1:C20",
    "Data!A1:A20%",
    "Data!A1:A20^2",
    "Data!A1:A20&\"\"",
    "Data!A1:A1000*2",
    "Data!F:F+0",
    "Data!G:G*Data!F:F",
    "Data!1:1*1",
    "Data!A1:B20*Data!C1:D20",
];

const SCALARS: &[&str] = &[
    "1",
    "2.5",
    "TRUE",
    "FALSE",
    "\"3\"",
    "\"x\"",
    "#N/A",
    "#DIV/0!",
    "1/0",
    "{1;2;3}",
    "{1,2;3,4}",
    "{1;2;3;4;5;6;7;8;9;10;11;12;13;14;15;16;17;18;19;20}",
];

const NESTED: &[&str] = &[
    "OFFSET(Data!A1,0,0,20,1)",
    "OFFSET(Data!A1,0,0,20,2)",
    "OFFSET(Data!A1,40,0,20,1)",
    "TRANSPOSE(Data!A1:T1)",
    "TRANSPOSE(Data!A1:A20)",
    "IF(Data!A1:A20>0,1,0)",
    "IF(TRUE,Data!A1:A20)",
    "IF(Data!A:A=\"x\",1,0)",
    "INDEX(Data!A:B,0,1)",
    "N(Data!A1)",
    "ABS(Data!A1:A20)",
];

/// arguments too large to pair with everything: read once, and against the
/// partners only.
const LARGE: &[&str] = &[
    "Data!A1:A1048575",
    "Data!A1048557:A1048576",
    "Data!A1:A1048576",
    "Data!A2:A1048576",
    "Data!A1:A1048575*1",
];

const PARTNERS: &[&str] = &["Data!A1:A20", "Data!A:A", "Data!B1:B19", "#N/A"];

/// formulas the spill layout holds on its own data sheet, above and below
/// the anchors, so they run before and after the spills in either order.
const LOCAL: &[&str] = &[
    "SUMPRODUCT(F:F)",
    "SUMPRODUCT(F1:F40*1)",
    "SUMPRODUCT(G:G,G:G)",
    "SUMPRODUCT((F:F>3)*1)",
    "SUMPRODUCT(A1:A100)",
    "SUMPRODUCT(A:A,F:F)",
    "SUMPRODUCT(F1:F30,G25:G54)",
    "SUMPRODUCT(G:G)+SUMPRODUCT(A60:A70)",
];

const LAYOUTS: &[&str] = &[
    "dense", "short", "empty", "sparse", "below", "merged", "spill", "wide",
];

fn atoms() -> Vec<&'static str> {
    RANGES
        .iter()
        .chain(EXPRESSIONS)
        .chain(SCALARS)
        .chain(NESTED)
        .copied()
        .collect()
}

/// a fixed xorshift stream, so the sampled argument lists never change.
struct Stream(u64);

impl Stream {
    fn below(&mut self, bound: usize) -> usize {
        self.0 ^= self.0 << 13;
        self.0 ^= self.0 >> 7;
        self.0 ^= self.0 << 17;
        (self.0 % bound as u64) as usize
    }
}

fn formulas() -> Vec<String> {
    let atoms = atoms();
    let mut out: Vec<String> = Vec::new();
    let sumproduct = |args: &[&str]| format!("SUMPRODUCT({})", args.join(","));
    for atom in atoms.iter().chain(LARGE) {
        out.push(sumproduct(&[*atom]));
    }
    for partner in PARTNERS {
        for atom in atoms.iter().chain(LARGE) {
            out.push(sumproduct(&[*partner, *atom]));
            out.push(sumproduct(&[*atom, *partner]));
        }
    }
    let mut stream = Stream(0x9E37_79B9_7F4A_7C15);
    for count in [2, 3, 4] {
        for _ in 0..80 {
            let args: Vec<&str> = (0..count)
                .map(|_| atoms[stream.below(atoms.len())])
                .collect();
            out.push(sumproduct(&args));
        }
    }
    for wrap in [
        "SUMPRODUCT(SUMPRODUCT(Data!A1:A20),Data!B1:B20)",
        "SUMPRODUCT(Data!A1:A20,SUMPRODUCT(Data!A:A))",
        "SUMPRODUCT(Data!A1:A20)+SUMPRODUCT(Data!A1:A1000)",
        "IFERROR(SUMPRODUCT(Data!A1:A20,Data!A1:A19),-1)",
        "SUMPRODUCT(Data!A1:A20,Data!B1:B20,Data!C1:C20,Data!D1:D20)",
        "SUMPRODUCT(Data!A:A,Data!B:B,Data!F:F,Data!G:G)",
    ] {
        out.push(wrap.to_string());
    }
    let mut seen = std::collections::HashSet::new();
    out.retain(|formula| seen.insert(formula.clone()));
    out
}

fn value_at(row: u32, col: u32) -> Option<CellValue> {
    if (col == 3 && row % 5 == 2) || (row == 1 && col >= 8 && col.is_multiple_of(4)) {
        let value = if row.is_multiple_of(2) {
            ErrorValue::Div0
        } else {
            ErrorValue::NA
        };
        return Some(CellValue::Error { value });
    }
    let number = |value: f64| Some(CellValue::Number { value });
    match (row * 7 + col * 13) % 17 {
        0..=5 => number(f64::from(row) * 1.5 - 7.0 + f64::from(col)),
        6 => number(0.1 * f64::from(row)),
        7 => Some(CellValue::Text { value: "x".into() }),
        8 => Some(CellValue::Text { value: "3".into() }),
        9 => Some(CellValue::Bool { value: true }),
        10 => Some(CellValue::Bool { value: false }),
        11 | 16 => None,
        12 => number(1e10 / 3.0),
        13 => Some(CellValue::Text {
            value: String::new(),
        }),
        14 => number(-0.25 * f64::from(col + 1)),
        _ => number(f64::from(row + col)),
    }
}

fn fill(sheet: &mut Sheet, rows: std::ops::Range<u32>, cols: std::ops::Range<u32>) {
    for row in rows {
        for col in cols.clone() {
            if let Some(value) = value_at(row, col) {
                put(sheet, CellRef::new(row, col), value);
            }
        }
    }
}

fn put(sheet: &mut Sheet, at: CellRef, value: CellValue) {
    sheet.set_cell(
        at,
        Cell {
            value,
            ..Cell::default()
        },
    );
}

fn put_formula(sheet: &mut Sheet, at: CellRef, formula: &str, array: bool) {
    sheet.set_cell(
        at,
        Cell {
            value: CellValue::Empty,
            formula: Some(formula.to_string()),
            ..Cell::default()
        },
    );
    if array {
        sheet.set_array_formula(at, CellRange::new(at, at));
    }
}

fn a1(address: &str) -> CellRef {
    CellRef::parse_a1(address).unwrap()
}

fn data_sheet(layout: &str) -> Sheet {
    let mut sheet = Sheet::new("Data");
    match layout {
        "dense" => fill(&mut sheet, 0..40, 0..21),
        "short" => fill(&mut sheet, 0..5, 0..4),
        "empty" => {}
        "sparse" => {
            for (address, row, col) in [
                ("A1", 0, 0),
                ("A3", 2, 0),
                ("B7", 6, 1),
                ("D1500", 1499, 3),
                ("C2999", 2998, 2),
                ("A3000", 2999, 0),
            ] {
                let value = value_at(row, col).unwrap_or(CellValue::Number { value: 4.0 });
                put(&mut sheet, a1(address), value);
            }
        }
        "below" => fill(&mut sheet, 29..45, 0..5),
        "merged" => {
            fill(&mut sheet, 0..3, 0..2);
            sheet.merges.push(CellRange::parse_a1("C1:C200").unwrap());
        }
        "spill" => {
            fill(&mut sheet, 0..5, 0..2);
            for (address, formula) in [
                ("F1", "_xlfn.SEQUENCE(30)"),
                ("G25", "_xlfn.SEQUENCE(40,1,0.5,0.25)"),
            ] {
                put_formula(&mut sheet, a1(address), formula, true);
            }
            for (index, formula) in LOCAL.iter().enumerate() {
                let index = index as u32;
                put_formula(&mut sheet, CellRef::new(index, 10), formula, false);
                put_formula(&mut sheet, CellRef::new(299 + index, 10), formula, false);
            }
        }
        "wide" => fill(&mut sheet, 0..3, 0..31),
        _ => unreachable!("{layout}"),
    }
    sheet
}

fn encode(value: &CellValue) -> String {
    match value {
        CellValue::Empty => "z".to_string(),
        CellValue::Number { value } => format!("n{:016x}", value.to_bits()),
        CellValue::Text { value } => format!("t{value:?}"),
        CellValue::Bool { value } => format!("b{}", u8::from(*value)),
        CellValue::Error { value } => format!("e{}", value.as_str()),
    }
}

fn stored(workbook: &Workbook, sheet: usize, at: CellRef) -> String {
    workbook.sheets[sheet]
        .cell(at)
        .map_or_else(|| "none".to_string(), |cell| encode(&cell.value))
}

/// one recalculation of `formulas` beside the data of `layout`, each formula
/// in a plain cell and an array cell of a sheet before the data and in a plain
/// cell of a sheet after it, appending the table lines to `out`.
fn run_chunk(layout: &str, chunk: usize, first: usize, formulas: &[String], out: &mut String) {
    let mut workbook = Workbook::default();
    let mut before = Sheet::new("Before");
    let mut after = Sheet::new("After");
    for (offset, formula) in formulas.iter().enumerate() {
        let row = offset as u32;
        put_formula(&mut before, CellRef::new(row, 0), formula, false);
        put_formula(&mut before, CellRef::new(row, 1), formula, true);
        put_formula(&mut after, CellRef::new(row, 0), formula, false);
    }
    workbook.sheets.push(before);
    workbook.sheets.push(data_sheet(layout));
    workbook.sheets.push(after);
    let (_graph, result) = rebuild_and_recalc_all(&mut workbook, None);
    writeln!(
        out,
        "{layout}\tchunk{chunk}\tbudget_exhausted={}\tlimited={}\tcycles={}\tunevaluated={}",
        result.budget_exhausted,
        result.limited_cells.len(),
        result.cycle_cells.len(),
        result.unevaluated_cells.len()
    )
    .unwrap();
    for (offset, formula) in formulas.iter().enumerate() {
        let row = offset as u32;
        let plain = stored(&workbook, 0, CellRef::new(row, 0));
        let array = stored(&workbook, 0, CellRef::new(row, 1));
        let later = stored(&workbook, 2, CellRef::new(row, 0));
        let direct = encode(&evaluate(
            &parse_formula(formula).unwrap(),
            &EvalContext::new(&workbook, SheetId(0)),
        ));
        let same = |value: String| {
            if value == plain {
                "=".to_string()
            } else {
                value
            }
        };
        writeln!(
            out,
            "{layout}\t{}\t{plain}\t{}\t{}\t{}",
            first + offset,
            same(array),
            same(later),
            same(direct)
        )
        .unwrap();
    }
    if layout == "spill" {
        for index in 0..LOCAL.len() as u32 {
            writeln!(
                out,
                "{layout}\tlocal{index}\t{}\t{}",
                stored(&workbook, 1, CellRef::new(index, 10)),
                stored(&workbook, 1, CellRef::new(299 + index, 10))
            )
            .unwrap();
        }
        for address in ["F30", "F31", "G64", "G65"] {
            let value = stored(&workbook, 1, a1(address));
            writeln!(out, "{layout}\t{address}\t{value}").unwrap();
        }
    }
}

fn table() -> String {
    let formulas = formulas();
    let mut out = String::new();
    writeln!(out, "formulas\t{}", formulas.len()).unwrap();
    for layout in LAYOUTS {
        for (chunk, part) in formulas.chunks(CHUNK).enumerate() {
            run_chunk(layout, chunk, chunk * CHUNK, part, &mut out);
        }
    }
    out
}

#[test]
fn sumproduct_results_match_the_recorded_table() {
    let actual = table();
    if let Ok(path) = std::env::var("SUMPRODUCT_USED_EXTENT_WRITE") {
        std::fs::write(&path, &actual).unwrap();
        return;
    }
    let path = concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/tests/fixtures/sumproduct_used_extent.tsv"
    );
    let expected = std::fs::read_to_string(path).unwrap();
    let formulas = formulas();
    let describe = |line: &str| {
        let index = line
            .split('\t')
            .nth(1)
            .and_then(|field| field.parse::<usize>().ok());
        index
            .and_then(|index| formulas.get(index))
            .map_or_else(String::new, |formula| format!(" [{formula}]"))
    };
    let mut differences = Vec::new();
    for (want, got) in expected.lines().zip(actual.lines()) {
        if want != got && differences.len() < 20 {
            differences.push(format!("want {want}\n got {got}{}", describe(want)));
        }
    }
    assert!(
        differences.is_empty(),
        "{} differences, first ones:\n{}",
        expected
            .lines()
            .zip(actual.lines())
            .filter(|(want, got)| want != got)
            .count(),
        differences.join("\n")
    );
    assert_eq!(
        expected.lines().count(),
        actual.lines().count(),
        "table length"
    );
}
