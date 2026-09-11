use betteroffice_xlsx::{
    CalculationOptions, CellRef, CellValue, ChartAnchor, DrawCmd, MAX_PIXMAP_DIM,
    MAX_PIXMAP_PIXELS, Op, RenderOptions, SheetId, Viewport, Workbook,
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
    let (pixels, _) = buffer[..info.buffer_size()].as_chunks::<4>();
    pixels
        .iter()
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
    let mut workbook = Workbook::open(&package).unwrap();
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

    assert_insert_is_refused(&mut workbook, SheetId(0));
    workbook
        .edit_cell(
            SheetId(0),
            CellRef::new(0, 0),
            "Period",
            CalculationOptions::default(),
        )
        .unwrap();
    let saved = workbook.save().unwrap();
    assert_eq!(
        part_of(&ooxml_opc::unzip_parts(&saved).unwrap(), DRAWING_PART),
        b"<xdr:wsDr><xdr:twoCellAnchor>"
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

const SECOND_DRAWING: &str = "xl/drawings/drawing2.xml";
const DRAWING_RELS: &str = "xl/drawings/_rels/drawing1.xml.rels";
const CONTENT_TYPES: &str = "[Content_Types].xml";

fn part_of(parts: &[(String, Vec<u8>)], path: &str) -> Vec<u8> {
    parts
        .iter()
        .find(|(existing, _)| existing == path)
        .map(|(_, bytes)| bytes.clone())
        .unwrap_or_else(|| panic!("{path} is not in the fixture"))
}

fn set_part(parts: &mut Vec<(String, Vec<u8>)>, path: &str, bytes: Vec<u8>) {
    match parts.iter_mut().find(|(existing, _)| existing == path) {
        Some(slot) => slot.1 = bytes,
        None => parts.push((path.to_owned(), bytes)),
    }
}

/// Rewrites a part's markup through `edit`.
fn rewrite_part(
    parts: &mut Vec<(String, Vec<u8>)>,
    path: &str,
    edit: impl FnOnce(String) -> String,
) {
    let source = String::from_utf8(part_of(parts, path)).unwrap();
    let edited = edit(source.clone());
    assert_ne!(edited, source, "{path} was left as it was");
    set_part(parts, path, edited.into_bytes());
}

/// `charts.xlsx` grown to a second sheet whose own drawing anchors the same four
/// chart parts, that drawing's markup passed through `damage`.
fn shared_chart_parts_fixture(damage: impl FnOnce(String) -> String) -> Vec<(String, Vec<u8>)> {
    let mut parts = ooxml_opc::unzip_parts(FIXTURE).unwrap();
    let sheet = part_of(&parts, "xl/worksheets/sheet1.xml");
    let drawing = String::from_utf8(part_of(&parts, DRAWING_PART)).unwrap();
    let drawing_rels = part_of(&parts, DRAWING_RELS);
    set_part(&mut parts, "xl/worksheets/sheet2.xml", sheet);
    set_part(
        &mut parts,
        "xl/worksheets/_rels/sheet2.xml.rels",
        br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdDrawing" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing2.xml"/></Relationships>"#.to_vec(),
    );
    set_part(&mut parts, SECOND_DRAWING, damage(drawing).into_bytes());
    set_part(
        &mut parts,
        "xl/drawings/_rels/drawing2.xml.rels",
        drawing_rels,
    );
    rewrite_part(&mut parts, "xl/workbook.xml", |text| {
        text.replace(
            r#"<sheet name="Charts" sheetId="1" r:id="rId1"/>"#,
            r#"<sheet name="Charts" sheetId="1" r:id="rId1"/><sheet name="Second" sheetId="2" r:id="rId2"/>"#,
        )
    });
    rewrite_part(&mut parts, "xl/_rels/workbook.xml.rels", |text| {
        text.replace(
            "</Relationships>",
            r#"<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/></Relationships>"#,
        )
    });
    rewrite_part(&mut parts, CONTENT_TYPES, |text| {
        text.replace(
            "</Types>",
            r#"<Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/drawings/drawing2.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/></Types>"#,
        )
    });
    parts
}

/// One anchor row this crate cannot read, which costs the whole drawing.
fn unreadable_anchor(drawing: String) -> String {
    drawing.replacen("<xdr:row>5</xdr:row>", "<xdr:row>not-a-row</xdr:row>", 1)
}

fn assert_insert_is_refused(workbook: &mut Workbook, sheet: SheetId) {
    let error = workbook
        .apply_ops(
            vec![Op::InsertRows {
                sheet,
                at: 0,
                count: 1,
            }],
            CalculationOptions::default(),
        )
        .unwrap_err();
    assert!(
        error
            .to_string()
            .contains("references cells this edit would move"),
        "{error}"
    );
}

fn assert_structural_edits_are_refused(package: &[u8]) {
    let mut workbook = Workbook::open(package).unwrap();
    assert_insert_is_refused(&mut workbook, SheetId(0));
}

/// A drawing this crate declined is one no save rewrites, so every structural
/// edit that would move what its anchors name is refused — on both sheets,
/// because the chart parts it holds are anchored from the other one too and so
/// look modelled.
#[test]
fn a_skipped_drawing_freezes_structural_edits_on_every_sheet() {
    let parts = shared_chart_parts_fixture(unreadable_anchor);
    let package = ooxml_opc::rezip_parts(&parts).unwrap();
    let mut workbook = Workbook::open(&package).unwrap();
    assert_eq!(workbook.model().sheet(SheetId(0)).unwrap().charts.len(), 4);
    assert!(
        workbook
            .model()
            .sheet(SheetId(1))
            .unwrap()
            .charts
            .is_empty()
    );

    assert_insert_is_refused(&mut workbook, SheetId(0));
    assert_insert_is_refused(&mut workbook, SheetId(1));
    let error = workbook
        .apply_ops(
            vec![Op::RemoveSheet { index: 1 }],
            CalculationOptions::default(),
        )
        .unwrap_err();
    assert!(
        error
            .to_string()
            .contains("references cells this edit would move"),
        "{error}"
    );

    workbook
        .edit_cell(
            SheetId(0),
            CellRef::new(0, 0),
            "Period",
            CalculationOptions::default(),
        )
        .unwrap();
    let saved = workbook.save().unwrap();
    let written = ooxml_opc::unzip_parts(&saved).unwrap();
    assert_eq!(
        part_of(&written, SECOND_DRAWING),
        part_of(&parts, SECOND_DRAWING)
    );
    let reopened = Workbook::open(&saved).unwrap();
    assert_eq!(
        reopened
            .model()
            .sheet(SheetId(0))
            .unwrap()
            .cell(CellRef::new(0, 0))
            .map(|cell| &cell.value),
        Some(&CellValue::Text {
            value: "Period".to_owned()
        })
    );
}

/// A chart part two sheets anchor is rewritten only when both of them want the
/// same bytes out of it. A drawing this crate declined leaves one of the two
/// sheets holding nothing, so that agreement stops being asked for — and an
/// unqualified reference in the shared part would be moved for one sheet while
/// the other one's data stayed where it was.
#[test]
fn a_skipped_drawing_does_not_let_a_shared_chart_be_rewritten_alone() {
    let mut parts = shared_chart_parts_fixture(unreadable_anchor);
    rewrite_part(&mut parts, "xl/charts/chart3.xml", |text| {
        text.replace("Charts!", "")
    });
    let package = ooxml_opc::rezip_parts(&parts).unwrap();
    let mut workbook = Workbook::open(&package).unwrap();
    assert_insert_is_refused(&mut workbook, SheetId(0));
}

/// A chart target outside the conventional layout that no content type claims
/// is in no inventory of the package, so only the decline itself can stop an
/// edit from stranding it.
#[test]
fn a_nonstandard_chart_target_that_cannot_be_read_freezes_structural_edits() {
    let mut parts = ooxml_opc::unzip_parts(FIXTURE).unwrap();
    parts.retain(|(path, _)| path != CHART_PART);
    set_part(
        &mut parts,
        "xl/custom/badplot.xml",
        b"<c:chartSpace><c:chart></c:chartSpace>".to_vec(),
    );
    rewrite_part(&mut parts, DRAWING_RELS, |text| {
        text.replace("../charts/chart2.xml", "../custom/badplot.xml")
    });
    rewrite_part(&mut parts, CONTENT_TYPES, |text| {
        text.replace(
            r#"<Override PartName="/xl/charts/chart2.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>"#,
            "",
        )
    });
    assert_structural_edits_are_refused(&ooxml_opc::rezip_parts(&parts).unwrap());
}

/// The same hole reached the other way: a chart target the package does not
/// hold at all, which nothing walking the parts can ever inventory.
#[test]
fn a_missing_chart_target_freezes_structural_edits() {
    let mut parts = ooxml_opc::unzip_parts(FIXTURE).unwrap();
    parts.retain(|(path, _)| path != CHART_PART);
    rewrite_part(&mut parts, CONTENT_TYPES, |text| {
        text.replace(
            r#"<Override PartName="/xl/charts/chart2.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>"#,
            "",
        )
    });
    assert_structural_edits_are_refused(&ooxml_opc::rezip_parts(&parts).unwrap());
}

/// An anchor that names a chart relationship the drawing rels never resolve,
/// with no chart part or content type left to veto on its behalf, is a frame
/// this crate cannot model, and its drawing is what a save would strand.
#[test]
fn an_anchor_whose_chart_relationship_is_unresolved_freezes_structural_edits() {
    let mut parts = ooxml_opc::unzip_parts(FIXTURE).unwrap();
    parts.retain(|(path, _)| path != CHART_PART);
    rewrite_part(&mut parts, DRAWING_RELS, |text| {
        text.replace(
            r#"<Relationship Id="rIdChart2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart2.xml"/>"#,
            "",
        )
    });
    rewrite_part(&mut parts, CONTENT_TYPES, |text| {
        text.replace(
            r#"<Override PartName="/xl/charts/chart2.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>"#,
            "",
        )
    });
    let package = ooxml_opc::rezip_parts(&parts).unwrap();
    assert_eq!(
        Workbook::open(&package)
            .unwrap()
            .model()
            .sheet(SheetId(0))
            .unwrap()
            .charts
            .len(),
        3
    );
    assert_structural_edits_are_refused(&package);
}

/// A `c:chart` that names no relationship is a frame this crate cannot model
/// either, and with nothing else left to veto on its behalf its drawing must.
#[test]
fn a_chart_element_without_a_relationship_freezes_structural_edits() {
    let mut parts = ooxml_opc::unzip_parts(FIXTURE).unwrap();
    parts.retain(|(path, _)| path != CHART_PART);
    rewrite_part(&mut parts, DRAWING_PART, |text| {
        text.replace(r#"<c:chart r:id="rIdChart2"/>"#, "<c:chart/>")
    });
    rewrite_part(&mut parts, DRAWING_RELS, |text| {
        text.replace(
            r#"<Relationship Id="rIdChart2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart2.xml"/>"#,
            "",
        )
    });
    rewrite_part(&mut parts, CONTENT_TYPES, |text| {
        text.replace(
            r#"<Override PartName="/xl/charts/chart2.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>"#,
            "",
        )
    });
    let package = ooxml_opc::rezip_parts(&parts).unwrap();
    assert_eq!(
        Workbook::open(&package)
            .unwrap()
            .model()
            .sheet(SheetId(0))
            .unwrap()
            .charts
            .len(),
        3
    );
    assert_structural_edits_are_refused(&package);
}

/// A declined chart target that another discovery already inventoried with
/// resolved areas has to be widened to name everything, not left as it was:
/// the pivot reader saw one cell, the chart reader saw nothing it could hold.
#[test]
fn a_declined_target_widens_an_entry_another_reader_resolved() {
    let mut parts = ooxml_opc::unzip_parts(FIXTURE).unwrap();
    parts.retain(|(path, _)| path != CHART_PART);
    rewrite_part(&mut parts, CONTENT_TYPES, |text| {
        text.replace(
            r#"<Override PartName="/xl/charts/chart2.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>"#,
            "",
        )
    });
    const PIVOT: &str = "xl/pivotCache/pivotCacheDefinition1.xml";
    set_part(
        &mut parts,
        PIVOT,
        br#"<pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><cacheSource type="worksheet"><worksheetSource sheet="Charts" ref="A1"/></cacheSource></pivotCacheDefinition>"#.to_vec(),
    );
    rewrite_part(&mut parts, DRAWING_RELS, |text| {
        text.replace(
            r#"Target="../charts/chart2.xml"/>"#,
            r#"Target="../pivotCache/pivotCacheDefinition1.xml"/>"#,
        )
    });
    rewrite_part(&mut parts, CONTENT_TYPES, |text| {
        text.replace(
            "</Types>",
            r#"<Override PartName="/xl/pivotCache/pivotCacheDefinition1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml"/></Types>"#,
        )
    });
    let package = ooxml_opc::rezip_parts(&parts).unwrap();
    let mut workbook = Workbook::open(&package).unwrap();
    assert_eq!(workbook.model().sheet(SheetId(0)).unwrap().charts.len(), 3);
    // an insert past A1 moves nothing the pivot named, so only the widened
    // entry can refuse it
    let refused = workbook.apply_ops(
        vec![Op::InsertRows {
            sheet: SheetId(0),
            at: 1,
            count: 1,
        }],
        CalculationOptions::default(),
    );
    assert!(
        refused.is_err(),
        "insert past the pivot's one cell was applied"
    );
}

/// Drawing relationships are reparsed by every real save, so opening over one
/// this crate cannot read would hand back a workbook that accepts an edit and
/// then refuses to write it. The open is where that can still be acted on.
#[test]
fn a_malformed_drawing_relationship_part_fails_the_open() {
    let package = fixture_with(
        DRAWING_RELS,
        b"<Relationships><Relationship></Relationships>",
    );
    assert!(Workbook::open(&package).is_err());
}
