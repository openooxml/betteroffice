//! end-to-end tests over the library path: build a real `.xlsx` in memory,
//! then load and render it via `xlsx_cli`'s public functions. no process
//! is spawned; png framing and dimensions are asserted, not text pixels.

use xlsx_cli::{
    MAX_PIXMAP_DIM, RenderOptions, load_workbook, render, resolve_sheet, sheet_summaries,
};
use xlsx_model::workbook::{Cell, Sheet};
use xlsx_model::{CellRange, CellRef, CellValue, Workbook};

/// build an in-memory `.xlsx`: fill a sheet, serialize to opc parts, zip them.
fn sample_xlsx() -> Vec<u8> {
    let mut sheet = Sheet::new("Data");
    for row in 0..8u32 {
        for col in 0..4u32 {
            sheet.set_cell(
                CellRef::new(row, col),
                Cell {
                    value: CellValue::Number {
                        value: (row * 4 + col) as f64,
                    },
                    ..Cell::default()
                },
            );
        }
    }
    let mut wb = Workbook::default();
    wb.sheets.push(sheet);
    wb.sheets.push(Sheet::new("Empty"));

    let parts = xlsx_parse::serialize_workbook(&wb).expect("serialize");
    ooxml_opc::rezip_parts(&parts).expect("rezip")
}

const PNG_MAGIC: [u8; 8] = [0x89, b'P', b'N', b'G', 0x0d, 0x0a, 0x1a, 0x0a];

/// read the width/height out of a png's IHDR chunk (bytes 16..24, big-endian).
fn png_dimensions(png: &[u8]) -> (u32, u32) {
    assert_eq!(&png[0..8], &PNG_MAGIC, "not a png");
    let w = u32::from_be_bytes([png[16], png[17], png[18], png[19]]);
    let h = u32::from_be_bytes([png[20], png[21], png[22], png[23]]);
    (w, h)
}

#[test]
fn renders_a_range_to_a_plausible_png() {
    let bytes = sample_xlsx();
    let wb = load_workbook(&bytes).unwrap();
    let sheet = resolve_sheet(&wb, Some("Data")).unwrap();

    let opts = RenderOptions {
        range: Some(CellRange::parse_a1("A1:D8").unwrap()),
        ..RenderOptions::default()
    };
    let out = render(&wb, sheet, &opts).unwrap();

    let (w, h) = png_dimensions(&out.bytes);
    assert_eq!((w, h), (out.width, out.height));
    // bounds fit four default columns (~64px) and eight default rows (20px)
    assert!((200..400).contains(&w), "width was {w}");
    assert!((150..200).contains(&h), "height was {h}");
}

#[test]
fn scale_doubles_output_dimensions() {
    let bytes = sample_xlsx();
    let wb = load_workbook(&bytes).unwrap();
    let sheet = resolve_sheet(&wb, None).unwrap();
    let range = CellRange::parse_a1("A1:D8").unwrap();

    let one = render(
        &wb,
        sheet,
        &RenderOptions {
            range: Some(range),
            ..RenderOptions::default()
        },
    )
    .unwrap();
    let two = render(
        &wb,
        sheet,
        &RenderOptions {
            range: Some(range),
            scale: 2.0,
            ..RenderOptions::default()
        },
    )
    .unwrap();

    let (w1, h1) = png_dimensions(&one.bytes);
    let (w2, h2) = png_dimensions(&two.bytes);
    assert!((w2 as i64 - 2 * w1 as i64).abs() <= 1, "{w1} vs {w2}");
    assert!((h2 as i64 - 2 * h1 as i64).abs() <= 1, "{h1} vs {h2}");
}

#[test]
fn used_range_render_defaults_when_no_range_given() {
    let bytes = sample_xlsx();
    let wb = load_workbook(&bytes).unwrap();
    let sheet = resolve_sheet(&wb, Some("Data")).unwrap();
    let out = render(&wb, sheet, &RenderOptions::default()).unwrap();
    let (w, h) = png_dimensions(&out.bytes);
    assert!(w > 0 && h > 0);
    assert!(w < 400 && h < 200);
}

#[test]
fn pixmap_guard_rejects_a_huge_range() {
    let bytes = sample_xlsx();
    let wb = load_workbook(&bytes).unwrap();
    let sheet = resolve_sheet(&wb, None).unwrap();

    let opts = RenderOptions {
        range: Some(CellRange::parse_a1("A1:XFD1000").unwrap()),
        ..RenderOptions::default()
    };
    let err = render(&wb, sheet, &opts).unwrap_err();
    assert!(err.contains("cap"), "unexpected error: {err}");
    assert!(err.contains(&MAX_PIXMAP_DIM.to_string()));
}

#[test]
fn scale_pushes_a_modest_range_past_the_guard() {
    let bytes = sample_xlsx();
    let wb = load_workbook(&bytes).unwrap();
    let sheet = resolve_sheet(&wb, None).unwrap();

    let range = CellRange::parse_a1("A1:Z400").unwrap();
    assert!(
        render(
            &wb,
            sheet,
            &RenderOptions {
                range: Some(range),
                ..RenderOptions::default()
            },
        )
        .is_ok()
    );
    let err = render(
        &wb,
        sheet,
        &RenderOptions {
            range: Some(range),
            scale: 40.0,
            ..RenderOptions::default()
        },
    )
    .unwrap_err();
    assert!(err.contains("cap"), "unexpected error: {err}");
}

#[test]
fn width_cap_crops_the_output() {
    let bytes = sample_xlsx();
    let wb = load_workbook(&bytes).unwrap();
    let sheet = resolve_sheet(&wb, None).unwrap();
    let out = render(
        &wb,
        sheet,
        &RenderOptions {
            range: Some(CellRange::parse_a1("A1:D8").unwrap()),
            max_width: Some(100),
            ..RenderOptions::default()
        },
    )
    .unwrap();
    let (w, _) = png_dimensions(&out.bytes);
    assert!(w <= 100, "cap not honored, width was {w}");
}

#[test]
fn resolve_sheet_by_name_index_and_errors() {
    let bytes = sample_xlsx();
    let wb = load_workbook(&bytes).unwrap();
    assert_eq!(resolve_sheet(&wb, None).unwrap().0, 0);
    assert_eq!(resolve_sheet(&wb, Some("Empty")).unwrap().0, 1);
    assert_eq!(resolve_sheet(&wb, Some("1")).unwrap().0, 1);
    assert!(resolve_sheet(&wb, Some("Nope")).is_err());
    assert!(resolve_sheet(&wb, Some("9")).is_err());
}

#[test]
fn info_summarizes_sheets() {
    let bytes = sample_xlsx();
    let wb = load_workbook(&bytes).unwrap();
    let summaries = sheet_summaries(&wb);
    assert_eq!(summaries.len(), 2);
    assert_eq!(summaries[0].name, "Data");
    assert_eq!(summaries[0].used_range.as_deref(), Some("A1:D8"));
    assert_eq!(summaries[0].cell_count, 32);
    assert_eq!(summaries[1].name, "Empty");
    assert_eq!(summaries[1].used_range, None);
    assert_eq!(summaries[1].cell_count, 0);
}

#[test]
fn rejects_non_xlsx_bytes() {
    assert!(load_workbook(b"not a zip at all").is_err());
}
