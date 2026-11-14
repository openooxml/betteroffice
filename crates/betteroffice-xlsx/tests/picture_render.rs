use betteroffice_xlsx::{
    CalculationOptions, DisplayList, DrawCmd, GridGeometry, Op, SheetId, Viewport, Workbook,
};

const MAIN: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const RELS: &str = "http://schemas.openxmlformats.org/package/2006/relationships";
const DOC_RELS: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const MEDIA: &str = "xl/media/image1.png";

fn relationships(entries: &[(&str, &str, &str)]) -> String {
    let entries: String = entries
        .iter()
        .map(|(id, kind, target)| {
            format!(r#"<Relationship Id="{id}" Type="{DOC_RELS}/{kind}" Target="{target}"/>"#)
        })
        .collect();
    format!(
        r#"<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="{RELS}">{entries}</Relationships>"#
    )
}

fn marker(tag: &str, col: u32, row: u32) -> String {
    format!(
        "<xdr:{tag}><xdr:col>{col}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>{row}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:{tag}>"
    )
}

fn picture(hidden: bool) -> String {
    let hidden = if hidden { r#" hidden="1""# } else { "" };
    format!(
        r#"<xdr:pic><xdr:nvPicPr><xdr:cNvPr id="2" name="Picture 1"{hidden}/><xdr:cNvPicPr/></xdr:nvPicPr><xdr:blipFill><a:blip r:embed="rIdImage"/></xdr:blipFill><xdr:spPr/></xdr:pic><xdr:clientData/>"#
    )
}

fn two_cell(from: (u32, u32), to: (u32, u32), hidden: bool) -> String {
    format!(
        "<xdr:twoCellAnchor>{}{}{}</xdr:twoCellAnchor>",
        marker("from", from.0, from.1),
        marker("to", to.0, to.1),
        picture(hidden)
    )
}

fn red_png() -> Vec<u8> {
    let mut bytes = Vec::new();
    let mut encoder = png::Encoder::new(&mut bytes, 1, 1);
    encoder.set_color(png::ColorType::Rgba);
    encoder.set_depth(png::BitDepth::Eight);
    let mut writer = encoder.write_header().unwrap();
    writer.write_image_data(&[255, 0, 0, 255]).unwrap();
    writer.finish().unwrap();
    bytes
}

/// One sheet whose drawing holds `anchors`, each embedding the red PNG.
fn package(anchors: &str) -> Vec<u8> {
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
                r#"<?xml version="1.0" encoding="UTF-8"?><worksheet xmlns="{MAIN}" xmlns:r="{DOC_RELS}"><sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData><drawing r:id="rIdDrawing"/></worksheet>"#
            ),
        ),
        (
            "xl/worksheets/_rels/sheet1.xml.rels",
            relationships(&[("rIdDrawing", "drawing", "../drawings/drawing1.xml")]),
        ),
        (
            "xl/drawings/drawing1.xml",
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
    parts.push((MEDIA.to_owned(), red_png()));
    ooxml_opc::rezip_parts(&parts).unwrap()
}

const VIEWPORT: Viewport = Viewport {
    x: 0.0,
    y: 0.0,
    width: 640.0,
    height: 400.0,
};

fn geometry(workbook: &Workbook) -> GridGeometry {
    GridGeometry::new(&workbook.model().sheets[0], &workbook.model().styles)
}

fn images(list: &DisplayList) -> Vec<(f32, f32, f32, f32)> {
    list.commands
        .iter()
        .filter_map(|command| match command {
            DrawCmd::Image {
                x,
                y,
                w,
                h,
                asset_id,
                ..
            } => {
                assert_eq!(&**asset_id, MEDIA);
                Some((*x, *y, *w, *h))
            }
            _ => None,
        })
        .collect()
}

#[test]
fn a_picture_draws_over_the_cells_it_is_anchored_to() {
    let workbook = Workbook::open(&package(&two_cell((1, 1), (3, 4), false))).unwrap();
    let geometry = geometry(&workbook);
    let (x, y) = (geometry.col_x(1), geometry.row_y(1));
    assert_eq!(
        images(&workbook.display_list(&VIEWPORT).unwrap()),
        [(x, y, geometry.col_x(3) - x, geometry.row_y(4) - y)]
    );
}

#[test]
fn a_hidden_picture_draws_nothing() {
    let workbook = Workbook::open(&package(&two_cell((1, 1), (3, 4), true))).unwrap();
    assert!(images(&workbook.display_list(&VIEWPORT).unwrap()).is_empty());
}

#[test]
fn a_picture_follows_a_row_insert_above_it() {
    let mut workbook = Workbook::open(&package(&two_cell((1, 1), (3, 4), false))).unwrap();
    let [(_, before, _, height)] = images(&workbook.display_list(&VIEWPORT).unwrap())[..] else {
        panic!("expected one picture");
    };
    workbook
        .apply_ops(
            vec![Op::InsertRows {
                sheet: SheetId(0),
                at: 0,
                count: 2,
            }],
            CalculationOptions::default(),
        )
        .unwrap();
    let geometry = geometry(&workbook);
    let [(_, after, _, moved_height)] = images(&workbook.display_list(&VIEWPORT).unwrap())[..]
    else {
        panic!("expected one picture");
    };
    assert_eq!((after, moved_height), (geometry.row_y(3), height));
    assert!(after > before);
}

#[test]
fn media_bytes_serves_only_images_a_picture_draws() {
    let workbook = Workbook::open(&package(&two_cell((1, 1), (3, 4), false))).unwrap();
    assert_eq!(workbook.media_bytes(MEDIA).unwrap(), red_png());
    assert!(workbook.media_bytes("xl/workbook.xml").is_err());
    let hidden = Workbook::open(&package(&two_cell((1, 1), (3, 4), true))).unwrap();
    assert!(hidden.media_bytes(MEDIA).is_err());
}

#[cfg(feature = "raster")]
#[test]
fn a_png_export_paints_the_picture() {
    let workbook = Workbook::open(&package(&two_cell((1, 1), (3, 4), false))).unwrap();
    let geometry = geometry(&workbook);
    let rendered = workbook
        .render_png(
            betteroffice_xlsx::PngTarget::Viewport {
                sheet: SheetId(0),
                viewport: VIEWPORT,
            },
            &betteroffice_xlsx::RenderOptions::default(),
        )
        .unwrap();
    let mut reader = png::Decoder::new(std::io::Cursor::new(rendered.bytes))
        .read_info()
        .unwrap();
    let mut pixels = vec![0; reader.output_buffer_size().unwrap()];
    let info = reader.next_frame(&mut pixels).unwrap();
    let pixel = |x: f32, y: f32| {
        let index = (y as usize * info.width as usize + x as usize) * 4;
        <[u8; 3]>::try_from(&pixels[index..index + 3]).unwrap()
    };
    let (x, y) = (
        (geometry.col_x(1) + geometry.col_x(3)) / 2.0,
        (geometry.row_y(1) + geometry.row_y(4)) / 2.0,
    );
    assert_eq!(pixel(x, y), [255, 0, 0]);
    assert_eq!(pixel(geometry.col_x(4) + 4.0, y), [255, 255, 255]);
}
