use betteroffice_xlsx::{
    ChartAnchor, DrawCmd, MAX_PIXMAP_DIM, MAX_PIXMAP_PIXELS, RenderOptions, SheetId, Viewport,
    Workbook,
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

const SHOWCASE: &[u8] = include_bytes!("../../../apps/demo/public/showcase.xlsx");

/// showcase.xlsx with its one `twoCellAnchor` repinned `from_row` rows down,
/// keeping the chart's 15-row height. Data stays in A1:F11.
fn showcase_with_chart_at_row(from_row: u32) -> Vec<u8> {
    let mut parts = ooxml_opc::unzip_parts(SHOWCASE).unwrap();
    let drawing = parts
        .iter_mut()
        .find(|(name, _)| name == "xl/drawings/drawing1.xml")
        .expect("drawing part");
    drawing.1 = String::from_utf8(drawing.1.clone())
        .unwrap()
        .replace(
            "<xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>",
            &format!("<xdr:row>{from_row}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>"),
        )
        .replace(
            "<xdr:row>16</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>",
            &format!(
                "<xdr:row>{}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>",
                from_row + 15
            ),
        )
        .into_bytes();
    ooxml_opc::rezip_parts(&parts).unwrap()
}

/// A chart anchored below the data still grows the default render to take it in.
#[test]
fn a_reachable_chart_grows_the_default_render() {
    let workbook = Workbook::open_for_read(SHOWCASE).unwrap();
    let png = workbook
        .render_sheet(SheetId(0), &RenderOptions::default())
        .unwrap();
    assert_eq!((png.width, png.height), (1075, 340));

    let display_list = workbook
        .display_list_for(
            SheetId(0),
            &Viewport {
                x: 0.0,
                y: 0.0,
                width: png.width as f32,
                height: png.height as f32,
            },
        )
        .unwrap();
    assert_eq!(display_list.charts.len(), 1);
}

/// Where a chart sits may not decide whether a legal sheet renders at all: one
/// parked far below the data is out of frame, and the data still comes out.
#[test]
fn a_chart_parked_far_below_the_data_leaves_the_sheet_renderable() {
    for row in [800, 1_000, 60_000] {
        let bytes = showcase_with_chart_at_row(row);
        let workbook = Workbook::open_for_read(&bytes).unwrap();
        let png = workbook
            .render_sheet(SheetId(0), &RenderOptions::default())
            .unwrap();
        assert!(png.bytes.starts_with(&[0x89, b'P', b'N', b'G']));
        assert_eq!((png.width, png.height), (563, 240), "chart at row {row}");
        assert!(u64::from(png.width) * u64::from(png.height) <= MAX_PIXMAP_PIXELS);
        assert!(png.width.max(png.height) <= MAX_PIXMAP_DIM);
    }
}

/// The frame the default render declines is still the caller's to ask for.
#[test]
fn an_explicit_viewport_still_reaches_a_far_chart() {
    let bytes = showcase_with_chart_at_row(1_000);
    let workbook = Workbook::open_for_read(&bytes).unwrap();
    let display_list = workbook
        .display_list_for(
            SheetId(0),
            &Viewport {
                x: 0.0,
                y: 20_000.0,
                width: 1075.0,
                height: 400.0,
            },
        )
        .unwrap();
    assert_eq!(display_list.charts.len(), 1);
    assert!(!display_list.charts[0].placeholder);
}
