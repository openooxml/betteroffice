use vsdx_export::{
    export_docx, export_docx_with_report, export_pptx, export_pptx_with_report, shape_data,
};
use vsdx_parse::{
    Cell, ForeignData, Relationship, Row, RowChild, Section, SectionChild, Shape, ShapeChild,
    ShapesChild, Sheet, SheetChild, TargetMode, TextToken, VsdxPackage,
};

fn cell(name: &str, value: &str) -> Cell {
    Cell {
        name: name.into(),
        formula: None,
        value: Some(value.into()),
        unit: None,
        del: false,
        other_attrs: Vec::new(),
    }
}

fn formula(name: &str, value: &str) -> Cell {
    Cell {
        name: name.into(),
        formula: Some(value.into()),
        value: None,
        unit: None,
        del: false,
        other_attrs: Vec::new(),
    }
}

fn row(index: u32, row_type: &str, cells: Vec<Cell>) -> Row {
    Row {
        index: Some(index),
        name: None,
        local_name: None,
        row_type: Some(row_type.into()),
        del: false,
        children: cells.into_iter().map(RowChild::Cell).collect(),
        other_attrs: Vec::new(),
    }
}

fn rectangle() -> Section {
    Section {
        name: "Geometry".into(),
        index: None,
        del: false,
        children: vec![
            row(0, "MoveTo", vec![cell("X", "0"), cell("Y", "0")]),
            row(1, "LineTo", vec![cell("X", "1"), cell("Y", "0")]),
            row(2, "LineTo", vec![cell("X", "1"), cell("Y", "1")]),
            row(3, "LineTo", vec![cell("X", "0"), cell("Y", "1")]),
        ]
        .into_iter()
        .map(SectionChild::Row)
        .collect(),
        other_attrs: Vec::new(),
    }
}

fn property_row(name: &str, label: &str, value: &str) -> Row {
    Row {
        index: None,
        name: Some(name.into()),
        local_name: None,
        row_type: None,
        del: false,
        children: vec![
            RowChild::Cell(cell("Label", label)),
            RowChild::Cell(cell("Value", value)),
        ],
        other_attrs: Vec::new(),
    }
}

fn rect_shape(id: u32, pin_x: &str, pin_y: &str) -> Shape {
    Shape {
        id,
        name: Some(format!("Box{id}")),
        name_u: None,
        shape_type: None,
        master: None,
        master_shape: None,
        line_style: None,
        fill_style: None,
        text_style: None,
        del: false,
        other_attrs: Vec::new(),
        children: vec![
            ShapeChild::Cell(cell("Width", "3")),
            ShapeChild::Cell(cell("Height", "2")),
            ShapeChild::Cell(cell("PinX", pin_x)),
            ShapeChild::Cell(cell("PinY", pin_y)),
            ShapeChild::Cell(cell("LocPinX", "0")),
            ShapeChild::Cell(cell("LocPinY", "0")),
            ShapeChild::Cell(cell("FillPattern", "1")),
            ShapeChild::Cell(formula("FillForegnd", "RGB(255,0,0)")),
            ShapeChild::Cell(cell("LinePattern", "1")),
            ShapeChild::Cell(formula("LineColor", "RGB(0,0,255)")),
            ShapeChild::Cell(cell("LineWeight", "0.02")),
            ShapeChild::Section(rectangle()),
            ShapeChild::Section(Section {
                name: "Property".into(),
                index: None,
                del: false,
                children: vec![SectionChild::Row(property_row("Cost", "Cost", "42"))],
                other_attrs: Vec::new(),
            }),
        ],
    }
}

fn text_shape() -> Shape {
    let mut shape = rect_shape(2, "7", "6");
    shape
        .children
        .push(ShapeChild::Text(vec![TextToken::Literal(
            "Hello".to_owned(),
        )]));
    shape
}

fn rotated_text_shape() -> Shape {
    let mut shape = text_shape();
    shape
        .children
        .push(ShapeChild::Cell(cell("Angle", "0.7853981633974483")));
    shape
}

fn package(shapes: Vec<Shape>) -> VsdxPackage {
    let mut package: VsdxPackage = serde_json::from_value(serde_json::json!({
        "documentPartPath": "", "pagesPartPath": null, "mastersPartPath": null,
        "pagePartPaths": ["page"], "masterPartPaths": [], "themePartPaths": [],
        "windowsPartPath": null, "relationships": {}, "documentSheet": null,
        "styleSheets": [], "colors": [], "faceNames": [], "pageSheets": {},
        "masterSheets": {}, "pagePartIds": {"page": 1}, "masterPartIds": {},
        "pageContents": {}, "masterContents": {}
    }))
    .unwrap();
    package.page_sheets.insert(
        1,
        Sheet {
            id: None,
            children: vec![
                SheetChild::Cell(cell("PageWidth", "10")),
                SheetChild::Cell(cell("PageHeight", "8")),
            ],
            other_attrs: Vec::new(),
        },
    );
    package.page_contents.insert(
        "page".into(),
        Sheet {
            id: None,
            children: vec![SheetChild::Shapes(
                shapes.into_iter().map(ShapesChild::Shape).collect(),
            )],
            other_attrs: Vec::new(),
        },
    );
    package
}

fn image_package() -> VsdxPackage {
    let mut diagram = package(vec![rect_shape(1, "2", "2")]);
    diagram.relationships.insert(
        "page".into(),
        vec![Relationship {
            id: "rId1".into(),
            relationship_type: "image".into(),
            target: "media/image1.png".into(),
            target_mode: TargetMode::Internal,
            resolved_target: Some("visio/media/image1.png".into()),
        }],
    );
    diagram.add_part(
        "visio/media/image1.png",
        vec![0x89, b'P', b'N', b'G', 0x0D, 0x0A, 0x1A, 0x0A, 0, 1, 2, 3],
    );
    let content = diagram.page_contents.get_mut("page").unwrap();
    let SheetChild::Shapes(shapes) = content.children.last_mut().unwrap() else {
        panic!("shapes expected");
    };
    shapes.push(ShapesChild::Shape(Shape {
        id: 5,
        name: Some("Photo".into()),
        name_u: None,
        shape_type: None,
        master: None,
        master_shape: None,
        line_style: None,
        fill_style: None,
        text_style: None,
        del: false,
        other_attrs: Vec::new(),
        children: vec![
            ShapeChild::Cell(cell("Width", "2")),
            ShapeChild::Cell(cell("Height", "1")),
            ShapeChild::Cell(cell("PinX", "5")),
            ShapeChild::Cell(cell("PinY", "4")),
            ShapeChild::Cell(cell("LocPinX", "0")),
            ShapeChild::Cell(cell("LocPinY", "0")),
            ShapeChild::ForeignData(ForeignData {
                foreign_type: None,
                compression_type: None,
                relationship_id: Some("rId1".into()),
                other_attrs: Vec::new(),
            }),
        ],
    }));
    diagram
}

#[test]
fn pptx_round_trips_through_engine() {
    let diagram = package(vec![rect_shape(1, "2", "2"), text_shape()]);
    let bytes = export_pptx(&diagram).unwrap();
    ooxml_opc::sanitize_package_for_format(&bytes, "pptx").unwrap();
    let parsed = pptx_parse::parse_pptx(&bytes).unwrap();
    assert_eq!(parsed.slides.len(), 1);
    assert_eq!(parsed.presentation.width_emu, 9_144_000);
    assert_eq!(parsed.presentation.height_emu, 7_315_200);
    let mut custom_paths = 0;
    let mut texts = Vec::new();
    for shape in &parsed.slides[0].shapes {
        if let pptx_parse::ShapeNode::Shape(shape) = shape {
            custom_paths += shape.paths.len();
            for paragraph in shape.text.iter().flat_map(|text| &text.paragraphs) {
                for run in &paragraph.runs {
                    texts.push(run.text.clone());
                }
            }
        }
    }
    assert!(custom_paths >= 1, "diagram geometry survives");
    assert!(
        texts.iter().any(|text| text.contains("Hello")),
        "diagram text survives"
    );
    let parts = ooxml_opc::unzip_parts(&bytes).unwrap();
    let slide = parts
        .iter()
        .find(|(path, _)| path == "ppt/slides/slide1.xml")
        .map(|(_, bytes)| String::from_utf8(bytes.clone()).unwrap())
        .unwrap();
    assert!(
        slide.contains("<a:custGeom>"),
        "geometry stays native DrawingML, not raster"
    );
}

#[test]
fn pptx_carries_images_as_media() {
    let diagram = image_package();
    let bytes = export_pptx(&diagram).unwrap();
    let parsed = pptx_parse::parse_pptx(&bytes).unwrap();
    assert_eq!(parsed.media.len(), 1);
    assert_eq!(
        parsed.media[0].bytes,
        vec![0x89, b'P', b'N', b'G', 0x0D, 0x0A, 0x1A, 0x0A, 0, 1, 2, 3]
    );
    let pictures = parsed.slides[0]
        .shapes
        .iter()
        .filter(|shape| matches!(shape, pptx_parse::ShapeNode::Picture(_)))
        .count();
    assert_eq!(pictures, 1);
}

#[test]
fn pptx_scales_mixed_pages_and_preserves_metafiles() {
    let diagram = vsdx_parse::parse_vsdx(include_bytes!(
        "../../vsdx-parse/tests/fixtures/export-mixed-pages.vsdx"
    ))
    .unwrap();
    let bytes = export_pptx(&diagram).unwrap();
    ooxml_opc::sanitize_package_for_format(&bytes, "pptx").unwrap();
    let parts = ooxml_opc::unzip_parts(&bytes).unwrap();
    let presentation = String::from_utf8(
        parts
            .iter()
            .find(|(path, _)| path == "ppt/presentation.xml")
            .unwrap()
            .1
            .clone(),
    )
    .unwrap();
    let slide = String::from_utf8(
        parts
            .iter()
            .find(|(path, _)| path == "ppt/slides/slide2.xml")
            .unwrap()
            .1
            .clone(),
    )
    .unwrap();
    let types = String::from_utf8(
        parts
            .iter()
            .find(|(path, _)| path == "[Content_Types].xml")
            .unwrap()
            .1
            .clone(),
    )
    .unwrap();
    assert!(presentation.contains("<p:sldSz cx=\"9144000\" cy=\"7315200\"/>"));
    assert!(slide.contains("<a:rPr sz=\"2400\""));
    assert!(slide.contains("Unsupported image:"));
    assert!(types.contains("Extension=\"emf\" ContentType=\"image/x-emf\""));
    assert!(types.contains("Extension=\"wmf\" ContentType=\"image/x-wmf\""));
    assert!(parts.iter().any(|(path, _)| path.ends_with(".emf")));
    assert!(parts.iter().any(|(path, _)| path.ends_with(".wmf")));
    assert!(!parts.iter().any(|(path, _)| path.ends_with(".tiff")));

    let docx = export_docx(&diagram).unwrap();
    let docx_parts = ooxml_opc::unzip_parts(&docx).unwrap();
    let document = String::from_utf8(
        docx_parts
            .iter()
            .find(|(path, _)| path == "word/document.xml")
            .unwrap()
            .1
            .clone(),
    )
    .unwrap();
    assert!(document.contains("Unsupported image:"));
    assert!(docx_parts.iter().any(|(path, _)| path.ends_with(".emf")));
    assert!(docx_parts.iter().any(|(path, _)| path.ends_with(".wmf")));
    assert!(!docx_parts.iter().any(|(path, _)| path.ends_with(".tiff")));
}

#[test]
fn shape_data_lists_property_rows() {
    let diagram = package(vec![rect_shape(1, "2", "2")]);
    let data = shape_data(&diagram).unwrap();
    assert_eq!(data.len(), 1);
    assert_eq!(data[0].shape_id, 1);
    assert_eq!(data[0].shape, "Box1");
    assert_eq!(data[0].label, "Cost");
    assert_eq!(data[0].value, "42");
}

#[test]
fn docx_keeps_metadata_separate_when_pages_share_a_name() {
    let mut diagram = package(vec![rect_shape(1, "2", "2")]);
    diagram.page_part_paths.push("second".into());
    diagram.page_part_ids.insert("second".into(), 2);
    diagram
        .page_sheets
        .insert(2, diagram.page_sheets[&1].clone());
    diagram
        .page_contents
        .insert("second".into(), diagram.page_contents["page"].clone());
    diagram.page_names.insert(1, "Same name".into());
    diagram.page_names.insert(2, "Same name".into());
    let bytes = export_docx(&diagram).unwrap();
    let parts = ooxml_opc::unzip_parts(&bytes).unwrap();
    let xml = std::str::from_utf8(
        &parts
            .iter()
            .find(|(path, _)| path == "word/document.xml")
            .unwrap()
            .1,
    )
    .unwrap();
    assert_eq!(xml.matches("<w:tbl>").count(), 2);
    assert_eq!(xml.matches(">42</w:t>").count(), 2);
}

#[test]
fn docx_fits_tall_pages_inside_the_printable_height() {
    let diagram = vsdx_parse::parse_vsdx(include_bytes!(
        "../../vsdx-parse/tests/fixtures/export-mixed-pages.vsdx"
    ))
    .unwrap();
    let bytes = export_docx(&diagram).unwrap();
    let parts = ooxml_opc::unzip_parts(&bytes).unwrap();
    let xml = std::str::from_utf8(
        &parts
            .iter()
            .find(|(path, _)| path == "word/document.xml")
            .unwrap()
            .1,
    )
    .unwrap();
    assert!(xml.contains("<wp:extent cx=\"2438400\" cy=\"7315200\"/>"));
}

#[test]
fn docx_round_trips_through_engine() {
    let diagram = package(vec![rect_shape(1, "2", "2"), text_shape()]);
    let bytes = export_docx(&diagram).unwrap();
    ooxml_opc::sanitize_package_for_format(&bytes, "docx").unwrap();
    let envelope =
        docx_parse::parse_docx_s9_wire(&bytes, docx_parse::S9ParseOptions::default()).unwrap();
    let content = &envelope.document.package.document.content;
    let tables: Vec<&docx_parse::Table> = content
        .iter()
        .filter_map(|block| match block {
            docx_parse::BlockContent::Table(table) => Some(table.as_ref()),
            _ => None,
        })
        .collect();
    assert_eq!(tables.len(), 1);
    let expected = shape_data(&diagram).unwrap().len();
    assert_eq!(tables[0].rows.len(), expected + 1);
    let labels: Vec<String> = tables[0]
        .rows
        .iter()
        .flat_map(|row| row.cells.iter())
        .flat_map(|cell| cell.content.iter())
        .filter_map(|block| match block {
            docx_parse::BlockContent::Paragraph(paragraph) => {
                Some(docx_parse::get_paragraph_text(paragraph))
            }
            _ => None,
        })
        .collect();
    assert!(labels.iter().any(|label| label == "Cost"));
    assert!(labels.iter().any(|label| label == "42"));
    let parts = ooxml_opc::unzip_parts(&bytes).unwrap();
    let document = parts
        .iter()
        .find(|(path, _)| path == "word/document.xml")
        .map(|(_, bytes)| String::from_utf8(bytes.clone()).unwrap())
        .unwrap();
    assert!(
        !document.contains("wps:cNvPr"),
        "Word rejects the unknown wps:cNvPr element"
    );
    assert!(
        document.contains("<wpg:cNvGrpSpPr/>"),
        "Word requires the group non-visual element"
    );
    assert_eq!(
        document.matches("<wps:wsp>").count(),
        document.matches("<wps:txbx>").count(),
        "Word requires a text box on every shape"
    );
}

#[test]
fn docx_carries_images_as_media() {
    let diagram = image_package();
    let bytes = export_docx(&diagram).unwrap();
    let envelope =
        docx_parse::parse_docx_s9_wire(&bytes, docx_parse::S9ParseOptions::default()).unwrap();
    let media = &envelope.document.package.media_entries;
    assert!(!media.is_empty());
}

#[test]
fn rotated_text_keeps_its_pre_rotation_frame() {
    let diagram = package(vec![rotated_text_shape()]);
    let bytes = export_pptx(&diagram).unwrap();
    let parts = ooxml_opc::unzip_parts(&bytes).unwrap();
    let slide = parts
        .iter()
        .find(|(path, _)| path == "ppt/slides/slide1.xml")
        .map(|(_, bytes)| String::from_utf8(bytes.clone()).unwrap())
        .unwrap();
    assert!(
        slide.contains("rot=\"-2700000\""),
        "rotation survives as a DrawingML rotation: {slide}"
    );
}

#[test]
fn degraded_export_returns_its_report() {
    let diagram = vsdx_parse::parse_vsdx(include_bytes!(
        "../../vsdx-parse/tests/fixtures/export-mixed-pages.vsdx"
    ))
    .unwrap();
    let pptx = export_pptx_with_report(&diagram).unwrap();
    assert_eq!(pptx.report.unsupported_images.len(), 1);
    assert_eq!(pptx.report.summary(), "1 image unsupported");
    let docx = export_docx_with_report(&diagram).unwrap();
    assert_eq!(docx.report.unsupported_images.len(), 1);
    assert!(!pptx.bytes.is_empty() && !docx.bytes.is_empty());
}

#[test]
fn clean_export_reports_nothing() {
    let diagram = package(vec![rect_shape(1, "2", "2"), text_shape()]);
    let pptx = export_pptx_with_report(&diagram).unwrap();
    assert_eq!(pptx.report.degraded_shapes(), 0);
    assert!(pptx.report.summary().is_empty());
}

#[test]
fn export_rejects_diagram_without_pages() {
    let mut diagram = package(Vec::new());
    diagram.page_part_paths.clear();
    assert!(matches!(
        export_pptx(&diagram),
        Err(vsdx_export::ExportError::NoPages)
    ));
}
