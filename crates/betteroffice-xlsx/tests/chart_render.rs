use betteroffice_xlsx::{
    CellRef, CellValue, ChartAnchor, DrawCmd, RenderOptions, SheetId, Viewport, Workbook,
};

const FIXTURE: &[u8] = include_bytes!("../../../packages/xlsx/test-fixtures/charts.xlsx");
const UNSUPPORTED: &[u8] =
    include_bytes!("../../../packages/xlsx/test-fixtures/unsupported-charts.xlsx");

#[test]
fn fixture_covers_anchor_families_display_list_and_raster() {
    let workbook = Workbook::open_for_read(FIXTURE).unwrap();
    let sheet = workbook.model().sheet(SheetId(0)).unwrap();
    assert_eq!(sheet.charts.len(), 4);
    assert!(matches!(
        sheet.charts.first().map(|chart| chart.anchor),
        Some(ChartAnchor::TwoCell { .. })
    ));
    assert!(
        sheet
            .charts
            .iter()
            .any(|chart| matches!(chart.anchor, ChartAnchor::OneCell { .. }))
    );
    assert!(
        sheet
            .charts
            .iter()
            .any(|chart| matches!(chart.anchor, ChartAnchor::Absolute { .. }))
    );

    let display_list = workbook
        .display_list_for(
            SheetId(0),
            &Viewport {
                x: 0.0,
                y: 0.0,
                width: 800.0,
                height: 800.0,
            },
        )
        .unwrap();
    assert_eq!(display_list.charts.len(), 4);
    assert!(
        display_list
            .commands
            .iter()
            .any(|command| matches!(command, DrawCmd::Path { .. }))
    );

    let png = workbook
        .render_sheet(
            SheetId(0),
            &RenderOptions {
                max_width: Some(900),
                max_height: Some(900),
                ..RenderOptions::default()
            },
        )
        .unwrap();
    assert!(png.bytes.starts_with(&[0x89, b'P', b'N', b'G']));
}

/// A 3-D plot and a stacked plot are both refused by the renderer. Neither may
/// take the worksheet down with it, in either backend.
#[test]
fn refused_charts_degrade_to_placeholders_and_keep_the_sheet() {
    let workbook = Workbook::open_for_read(UNSUPPORTED).unwrap();
    assert_eq!(workbook.model().sheet(SheetId(0)).unwrap().charts.len(), 2);

    let display_list = workbook
        .display_list_for(
            SheetId(0),
            &Viewport {
                x: 0.0,
                y: 0.0,
                width: 900.0,
                height: 500.0,
            },
        )
        .unwrap();

    let labels = display_list
        .charts
        .iter()
        .map(|chart| chart.label.as_str())
        .collect::<Vec<_>>();
    assert_eq!(
        labels,
        [
            "Chart, not shown",
            "Revenue stacked, bar chart, 1 series, 4 categories, not shown"
        ]
    );
    assert!(display_list.charts.iter().all(|chart| chart.placeholder));

    assert!(display_list.commands.iter().any(|command| matches!(
        command,
        DrawCmd::Text { text, chart: false, .. } if text == "Quarter"
    )));
    assert!(display_list.grid.row_offsets.len() > 1);
    assert_eq!(
        display_list
            .commands
            .iter()
            .filter(|command| matches!(
                command,
                DrawCmd::FillRect { color, clip: Some(_), .. } if color == "#f2f2f2"
            ))
            .count(),
        2
    );
    // two placeholders and nothing else on the chart layer: neither refused
    // plot leaked any of the ops it managed to emit before being rejected.
    assert_eq!(
        display_list
            .commands
            .iter()
            .filter(|command| match command {
                DrawCmd::FillRect { clip, .. }
                | DrawCmd::Line { clip, .. }
                | DrawCmd::Path { clip, .. } => clip.is_some(),
                DrawCmd::Text { chart, .. } => *chart,
            })
            .count(),
        10
    );

    let png = workbook
        .render_sheet(
            SheetId(0),
            &RenderOptions {
                max_width: Some(900),
                max_height: Some(600),
                ..RenderOptions::default()
            },
        )
        .unwrap();
    assert!(png.bytes.starts_with(&[0x89, b'P', b'N', b'G']));

    let pixels = decode_rgba(&png.bytes);
    assert!(pixels.contains(&[0xf2, 0xf2, 0xf2]));
    assert!(
        pixels
            .iter()
            .any(|pixel| *pixel != [0xff, 0xff, 0xff] && *pixel != [0xf2, 0xf2, 0xf2])
    );
}

fn decode_rgba(bytes: &[u8]) -> Vec<[u8; 3]> {
    let mut reader = png::Decoder::new(std::io::Cursor::new(bytes))
        .read_info()
        .unwrap();
    let mut buffer = vec![0; reader.output_buffer_size().unwrap()];
    let info = reader.next_frame(&mut buffer).unwrap();
    buffer[..info.buffer_size()]
        .chunks_exact(4)
        .map(|pixel| [pixel[0], pixel[1], pixel[2]])
        .collect()
}

const CHART_PART: &str = "xl/charts/chart2.xml";
const DRAWING_PART: &str = "xl/drawings/drawing1.xml";

/// `charts.xlsx` with one part's bytes replaced.
fn fixture_with(part: &str, bytes: &[u8]) -> Vec<u8> {
    let mut parts = ooxml_opc::unzip_parts(FIXTURE).unwrap();
    let slot = parts
        .iter_mut()
        .find(|(path, _)| path == part)
        .unwrap_or_else(|| panic!("{part} is not in the fixture"));
    slot.1 = bytes.to_vec();
    ooxml_opc::rezip_parts(&parts).unwrap()
}

/// Opens a fixture whose chart part carries `bytes` and asserts the workbook is
/// whole apart from that one chart.
fn assert_damaged_chart_is_dropped(bytes: &[u8]) {
    let package = fixture_with(CHART_PART, bytes);
    let workbook = Workbook::open(&package).unwrap();
    let model = workbook.model();
    assert_eq!(model.sheets.len(), 1);
    let sheet = model.sheet(SheetId(0)).unwrap();
    assert_eq!(
        sheet.cell(CellRef::new(0, 0)).map(|cell| &cell.value),
        Some(&CellValue::Text {
            value: "Quarter".to_owned()
        })
    );
    assert_eq!(
        sheet.cell(CellRef::new(4, 1)).map(|cell| &cell.value),
        Some(&CellValue::Number { value: 24.0 })
    );
    assert_eq!(sheet.charts.len(), 3);
    assert!(sheet.charts.iter().all(|chart| chart.part != CHART_PART));

    let display_list = workbook
        .display_list_for(
            SheetId(0),
            &Viewport {
                x: 0.0,
                y: 0.0,
                width: 800.0,
                height: 800.0,
            },
        )
        .unwrap();
    assert_eq!(display_list.charts.len(), 3);
    assert!(display_list.commands.iter().any(|command| matches!(
        command,
        DrawCmd::Text { text, chart: false, .. } if text == "Quarter"
    )));
}

#[test]
fn a_malformed_chart_part_drops_only_that_chart() {
    assert_damaged_chart_is_dropped(b"<c:chartSpace><c:chart></c:chartSpace>");
}

#[test]
fn an_empty_chart_part_drops_only_that_chart() {
    assert_damaged_chart_is_dropped(b"");
}

#[test]
fn a_doctype_bearing_chart_part_drops_only_that_chart() {
    assert_damaged_chart_is_dropped(
        br#"<?xml version="1.0"?><!DOCTYPE c:chartSpace [<!ENTITY x "y">]><c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"/>"#,
    );
}

#[test]
fn a_truncated_chart_part_drops_only_that_chart() {
    let mut parts = ooxml_opc::unzip_parts(FIXTURE).unwrap();
    let original = parts
        .iter_mut()
        .find(|(path, _)| path == CHART_PART)
        .map(|(_, bytes)| bytes.clone())
        .unwrap();
    assert_damaged_chart_is_dropped(&original[..original.len() / 2]);
}

#[test]
fn a_binary_chart_part_drops_only_that_chart() {
    assert_damaged_chart_is_dropped(&[0x00, 0xff, 0xfe, 0x01, 0x80, 0x7f]);
}

#[test]
fn an_undeclared_entity_in_a_chart_part_drops_only_that_chart() {
    assert_damaged_chart_is_dropped(
        br#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart><c:title>&nbsp;</c:title></c:chart></c:chartSpace>"#,
    );
}

/// A damaged drawing costs every chart it anchors, because anchor positions
/// name themselves by index, but never the sheet.
#[test]
fn a_malformed_drawing_part_drops_every_chart_and_keeps_the_sheet() {
    let package = fixture_with(DRAWING_PART, b"<xdr:wsDr><xdr:twoCellAnchor>");
    let workbook = Workbook::open(&package).unwrap();
    let sheet = workbook.model().sheet(SheetId(0)).unwrap();
    assert!(sheet.charts.is_empty());
    assert_eq!(
        sheet.cell(CellRef::new(0, 0)).map(|cell| &cell.value),
        Some(&CellValue::Text {
            value: "Quarter".to_owned()
        })
    );
    assert!(
        workbook
            .render_sheet(
                SheetId(0),
                &RenderOptions {
                    max_width: Some(900),
                    max_height: Some(900),
                    ..RenderOptions::default()
                },
            )
            .unwrap()
            .bytes
            .starts_with(&[0x89, b'P', b'N', b'G'])
    );
}

/// Refusing to parse a chart part must not stop the package from carrying it.
#[test]
fn an_unparseable_chart_part_survives_a_save_byte_for_byte() {
    const DAMAGED: &[u8] = b"<c:chartSpace><c:chart></c:chartSpace>";
    let package = fixture_with(CHART_PART, DAMAGED);
    let saved = Workbook::open(&package).unwrap().save().unwrap();

    let before = ooxml_opc::unzip_parts(&package).unwrap();
    let after = ooxml_opc::unzip_parts(&saved).unwrap();
    let part_bytes = |parts: &[(String, Vec<u8>)]| {
        parts
            .iter()
            .find(|(path, _)| path == CHART_PART)
            .map(|(_, bytes)| bytes.clone())
            .unwrap()
    };
    assert_eq!(part_bytes(&after), DAMAGED);
    assert_eq!(part_bytes(&after), part_bytes(&before));
    assert_eq!(
        after.iter().map(|(path, _)| path).collect::<Vec<_>>(),
        before.iter().map(|(path, _)| path).collect::<Vec<_>>()
    );
}
