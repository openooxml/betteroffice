//! Pictures and shapes follow a row or column insert or delete on save, as charts already do.
//! What a first version cannot place exactly refuses the edit instead of leaving it stale.

use std::collections::BTreeMap;

use betteroffice_xlsx::{CalculationOptions, Op, SheetId, Workbook};

const MAIN: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const RELS: &str = "http://schemas.openxmlformats.org/package/2006/relationships";
const DOC_RELS: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const DRAWING: &str = "xl/drawings/drawing1.xml";
const REFUSAL: &str =
    "xl/drawings/drawing1.xml references cells this edit would move, and it cannot be rewritten";

fn options() -> CalculationOptions {
    CalculationOptions::default()
}

fn relationships(entries: &[(&str, &str, &str)]) -> String {
    let mut xml =
        format!(r#"<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="{RELS}">"#);
    for (id, kind, target) in entries {
        xml.push_str(&format!(
            r#"<Relationship Id="{id}" Type="{DOC_RELS}/{kind}" Target="{target}"/>"#
        ));
    }
    xml + "</Relationships>"
}

fn marker(tag: &str, col: u32, row: u32) -> String {
    format!(
        "<xdr:{tag}><xdr:col>{col}</xdr:col><xdr:colOff>12700</xdr:colOff><xdr:row>{row}</xdr:row><xdr:rowOff>6350</xdr:rowOff></xdr:{tag}>"
    )
}

fn picture(id: u32) -> String {
    format!(
        r#"<xdr:pic><xdr:nvPicPr><xdr:cNvPr id="{id}" name="Picture {id}"/><xdr:cNvPicPr/></xdr:nvPicPr><xdr:blipFill><a:blip r:embed="rIdImage"/></xdr:blipFill><xdr:spPr/></xdr:pic><xdr:clientData/>"#
    )
}

fn shape(id: u32) -> String {
    format!(
        r#"<xdr:sp macro="" textlink=""><xdr:nvSpPr><xdr:cNvPr id="{id}" name="Shape {id}"/><xdr:cNvSpPr/></xdr:nvSpPr><xdr:spPr/></xdr:sp><xdr:clientData/>"#
    )
}

/// A two-cell anchor; `edit_as` is written only when given.
fn two_cell(edit_as: Option<&str>, from: (u32, u32), to: (u32, u32), id: u32) -> String {
    let edit_as = edit_as.map_or(String::new(), |mode| format!(r#" editAs="{mode}""#));
    format!(
        "<xdr:twoCellAnchor{edit_as}>{}{}{}</xdr:twoCellAnchor>",
        marker("from", from.0, from.1),
        marker("to", to.0, to.1),
        picture(id)
    )
}

fn one_cell(from: (u32, u32), id: u32) -> String {
    format!(
        r#"<xdr:oneCellAnchor>{}<xdr:ext cx="914400" cy="457200"/>{}</xdr:oneCellAnchor>"#,
        marker("from", from.0, from.1),
        shape(id)
    )
}

fn absolute(id: u32) -> String {
    format!(
        r#"<xdr:absoluteAnchor><xdr:pos x="1828800" y="914400"/><xdr:ext cx="914400" cy="457200"/>{}</xdr:absoluteAnchor>"#,
        picture(id)
    )
}

/// One sheet with values in A1:A20 and a drawing holding `anchors`.
fn package(anchors: &str) -> Vec<u8> {
    let rows = (1..=20)
        .map(|row| format!(r#"<row r="{row}"><c r="A{row}"><v>{row}</v></c></row>"#))
        .collect::<String>();
    let parts = [
        (
            "[Content_Types].xml",
            r#"<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/></Types>"#.to_owned(),
        ),
        (
            "_rels/.rels",
            relationships(&[("rId1", "officeDocument", "xl/workbook.xml")]),
        ),
        (
            "xl/workbook.xml",
            format!(
                r#"<?xml version="1.0" encoding="UTF-8"?><workbook xmlns="{MAIN}" xmlns:r="{DOC_RELS}"><sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets></workbook>"#
            ),
        ),
        (
            "xl/_rels/workbook.xml.rels",
            relationships(&[("rId1", "worksheet", "worksheets/sheet1.xml")]),
        ),
        (
            "xl/worksheets/sheet1.xml",
            format!(
                r#"<?xml version="1.0" encoding="UTF-8"?><worksheet xmlns="{MAIN}" xmlns:r="{DOC_RELS}"><sheetData>{rows}</sheetData><drawing r:id="rIdDrawing"/></worksheet>"#
            ),
        ),
        (
            "xl/worksheets/_rels/sheet1.xml.rels",
            relationships(&[("rIdDrawing", "drawing", "../drawings/drawing1.xml")]),
        ),
        (
            DRAWING,
            format!(
                r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="{DOC_RELS}">{anchors}</xdr:wsDr>"#
            ),
        ),
        (
            "xl/drawings/_rels/drawing1.xml.rels",
            relationships(&[("rIdImage", "image", "../media/image1.png")]),
        ),
    ];
    let mut parts = parts
        .into_iter()
        .map(|(name, xml)| (name.to_owned(), xml.into_bytes()))
        .collect::<Vec<_>>();
    parts.push((
        "xl/media/image1.png".to_owned(),
        vec![0x89, b'P', b'N', b'G', 0, 1, 2, 3],
    ));
    ooxml_opc::rezip_parts(&parts).unwrap()
}

fn parts(bytes: &[u8]) -> BTreeMap<String, Vec<u8>> {
    ooxml_opc::unzip_parts(bytes).unwrap().into_iter().collect()
}

fn drawing_xml(bytes: &[u8]) -> String {
    String::from_utf8(parts(bytes)[DRAWING].clone()).unwrap()
}

fn source_drawing(anchors: &str) -> String {
    drawing_xml(&package(anchors))
}

fn insert_rows(at: u32, count: u32) -> Op {
    Op::InsertRows {
        sheet: SheetId(0),
        at,
        count,
    }
}

fn delete_rows(at: u32, count: u32) -> Op {
    Op::DeleteRows {
        sheet: SheetId(0),
        at,
        count,
    }
}

fn insert_cols(at: u32, count: u32) -> Op {
    Op::InsertCols {
        sheet: SheetId(0),
        at,
        count,
    }
}

/// Applies `op`, saves, and returns the saved drawing part; the saved file also reopens.
fn after(anchors: &str, op: Op) -> String {
    let mut workbook = Workbook::open(&package(anchors)).unwrap();
    workbook.apply_ops(vec![op], options()).unwrap();
    let saved = workbook.save().unwrap();
    Workbook::open(&saved).unwrap();
    drawing_xml(&saved)
}

#[test]
fn a_row_insert_above_moves_every_cell_anchored_drawing() {
    let anchors = [
        two_cell(None, (1, 4), (3, 8), 2),
        two_cell(Some("oneCell"), (5, 4), (7, 8), 3),
        one_cell((1, 12), 4),
        absolute(5),
    ]
    .concat();
    let expected = source_drawing(
        &[
            two_cell(None, (1, 6), (3, 10), 2),
            two_cell(Some("oneCell"), (5, 6), (7, 10), 3),
            one_cell((1, 14), 4),
            absolute(5),
        ]
        .concat(),
    );
    assert_eq!(after(&anchors, insert_rows(2, 2)), expected);
}

#[test]
fn a_column_insert_moves_anchors_at_and_right_of_it() {
    let anchors = [
        two_cell(None, (1, 4), (3, 8), 2),
        two_cell(None, (6, 4), (8, 8), 3),
    ]
    .concat();
    let expected = source_drawing(
        &[
            two_cell(None, (1, 4), (4, 8), 2),
            two_cell(None, (7, 4), (9, 8), 3),
        ]
        .concat(),
    );
    assert_eq!(after(&anchors, insert_cols(2, 1)), expected);
}

#[test]
fn a_row_insert_inside_a_two_cell_anchor_stretches_it_and_a_one_cell_mode_keeps_its_size() {
    let anchors = [
        two_cell(None, (1, 4), (3, 8), 2),
        two_cell(Some("oneCell"), (5, 4), (7, 8), 3),
    ]
    .concat();
    let expected = source_drawing(
        &[
            two_cell(None, (1, 4), (3, 9), 2),
            two_cell(Some("oneCell"), (5, 4), (7, 8), 3),
        ]
        .concat(),
    );
    assert_eq!(after(&anchors, insert_rows(6, 1)), expected);
}

#[test]
fn a_delete_over_part_of_an_anchor_pulls_its_corner_to_the_deletion_point() {
    let anchors = two_cell(None, (1, 4), (3, 8), 2);
    let expected = source_drawing(&two_cell(None, (1, 3), (3, 5), 2));
    assert_eq!(after(&anchors, delete_rows(3, 3)), expected);
}

#[test]
fn an_edit_below_every_anchor_leaves_the_drawing_untouched() {
    let anchors = [two_cell(None, (1, 4), (3, 8), 2), absolute(3)].concat();
    assert_eq!(
        after(&anchors, insert_rows(15, 2)),
        source_drawing(&anchors)
    );
}

#[test]
fn a_true_absolute_anchor_never_blocks_an_edit() {
    let anchors = absolute(2);
    assert_eq!(after(&anchors, insert_rows(0, 3)), source_drawing(&anchors));
    assert_eq!(after(&anchors, delete_rows(0, 3)), source_drawing(&anchors));
}

#[test]
fn edits_the_first_version_cannot_place_refuse_and_change_nothing() {
    let cases = [
        (
            "a cell-anchored picture that does not move with cells, below the insert",
            two_cell(Some("absolute"), (1, 4), (3, 8), 2),
            insert_rows(2, 1),
        ),
        (
            "a cell-anchored picture that does not move with cells, under a column insert",
            two_cell(Some("absolute"), (1, 4), (3, 8), 2),
            insert_cols(0, 1),
        ),
        (
            "a picture wholly inside deleted rows",
            two_cell(None, (1, 4), (3, 6), 2),
            delete_rows(3, 5),
        ),
        (
            "a shape whose anchor cell is deleted",
            one_cell((1, 4), 2),
            delete_rows(4, 1),
        ),
        (
            "a picture an insert would push past the last row",
            two_cell(None, (1, 1_048_570), (3, 1_048_575), 2),
            insert_rows(2, 1),
        ),
    ];
    for (label, anchors, op) in cases {
        let source = package(&anchors);
        let mut workbook = Workbook::open(&source).unwrap();
        let error = workbook.apply_ops(vec![op], options()).unwrap_err();
        assert_eq!(error.to_string(), REFUSAL, "{label}");
        assert_eq!(
            drawing_xml(&workbook.save().unwrap()),
            drawing_xml(&source),
            "{label}"
        );
    }
}

#[test]
fn a_cell_anchored_picture_that_does_not_move_allows_edits_after_it() {
    let anchors = two_cell(Some("absolute"), (1, 4), (3, 8), 2);
    assert_eq!(after(&anchors, insert_rows(9, 1)), source_drawing(&anchors));
    assert_eq!(after(&anchors, insert_cols(4, 1)), source_drawing(&anchors));
}

#[test]
fn undo_restores_the_source_drawing_and_redo_moves_it_again() {
    let anchors = [two_cell(None, (1, 4), (3, 8), 2), one_cell((1, 12), 3)].concat();
    let mut workbook = Workbook::open(&package(&anchors)).unwrap();
    workbook
        .apply_ops(vec![insert_rows(2, 2)], options())
        .unwrap();
    let moved = drawing_xml(&workbook.save().unwrap());
    workbook.undo(options()).unwrap();
    assert_eq!(
        drawing_xml(&workbook.save().unwrap()),
        source_drawing(&anchors)
    );
    workbook.redo(options()).unwrap();
    assert_eq!(drawing_xml(&workbook.save().unwrap()), moved);
}

#[test]
fn successive_edits_compose() {
    let anchors = two_cell(None, (1, 4), (3, 8), 2);
    let mut workbook = Workbook::open(&package(&anchors)).unwrap();
    workbook
        .apply_ops(vec![insert_rows(0, 3), insert_cols(0, 1)], options())
        .unwrap();
    workbook
        .apply_ops(vec![delete_rows(0, 1)], options())
        .unwrap();
    assert_eq!(
        drawing_xml(&workbook.save().unwrap()),
        source_drawing(&two_cell(None, (2, 6), (4, 10), 2))
    );
}
