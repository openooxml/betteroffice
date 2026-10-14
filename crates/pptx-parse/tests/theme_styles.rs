use pptx_parse::{ShapeNode, ThemePart, parse_pptx};

const FIXTURE: &[u8] = include_bytes!("fixtures/style-matrix-deck.pptx");

#[test]
fn parses_theme_matrix_slots_references_and_explicit_fill_barriers() {
    let package = parse_pptx(FIXTURE).unwrap();
    let scheme = &package.themes[0].format_scheme;
    assert_eq!(scheme.fills.len(), 3);
    assert_eq!(scheme.fills[0], None);
    assert_eq!(scheme.fills[1].as_ref().unwrap().fill_type, "gradient");
    assert_eq!(scheme.fills[2].as_ref().unwrap().fill_type, "solid");
    assert_eq!(scheme.lines.len(), 3);
    assert!(scheme.lines[1].as_ref().unwrap().color.is_none());
    assert_eq!(scheme.lines[2].as_ref().unwrap().width, Some(28575.0));
    assert_eq!(scheme.background_fills.len(), 2);
    let shape = |id| match package.slides[0]
        .shapes
        .iter()
        .find(|shape| shape.id() == id)
        .unwrap()
    {
        ShapeNode::Shape(shape) => shape,
        _ => panic!("expected shape"),
    };
    let reference = shape(2).style.as_ref().unwrap().fill.as_ref().unwrap();
    assert_eq!(reference.index, 3);
    let color = reference.color.as_ref().unwrap();
    assert_eq!(color.rgb.as_deref(), Some("808080"));
    assert_eq!(color.luminance_modulation, Some(0.5));
    assert_eq!(color.luminance_offset, Some(0.1));
    assert_eq!(
        shape(6)
            .fill
            .as_ref()
            .unwrap()
            .color
            .as_ref()
            .unwrap()
            .rgb
            .as_deref(),
        Some("112233")
    );
    assert_eq!(
        shape(14)
            .style
            .as_ref()
            .unwrap()
            .line
            .as_ref()
            .unwrap()
            .index,
        3
    );
    assert!(shape(15).style.as_ref().unwrap().line_disabled);
    assert!(shape(17).style.as_ref().unwrap().fill_disabled);
    assert!(shape(17).style.as_ref().unwrap().line_disabled);
    let picture = package.slides[0]
        .shapes
        .iter()
        .find(|s| s.id() == 23)
        .unwrap();
    let ShapeNode::Picture(picture) = picture else {
        panic!("expected picture")
    };
    assert_eq!(
        picture.style.as_ref().unwrap().line.as_ref().unwrap().index,
        3
    );
}

#[test]
fn absent_style_fields_deserialize_and_serialize_without_new_defaults() {
    let package = parse_pptx(FIXTURE).unwrap();
    let mut legacy = serde_json::to_value(&package.themes[0]).unwrap();
    legacy.as_object_mut().unwrap().remove("formatScheme");
    let theme: ThemePart = serde_json::from_value(legacy.clone()).unwrap();
    assert!(theme.format_scheme.is_empty());
    assert_eq!(serde_json::to_value(theme).unwrap(), legacy);
    let shape = package.slides[0]
        .shapes
        .iter()
        .find(|shape| shape.id() == 2)
        .unwrap();
    let mut legacy = serde_json::to_value(shape).unwrap();
    legacy.as_object_mut().unwrap().remove("style");
    let shape: ShapeNode = serde_json::from_value(legacy.clone()).unwrap();
    assert_eq!(serde_json::to_value(shape).unwrap(), legacy);
    let style = ooxml_drawingml::ShapeStyle::default();
    assert_eq!(serde_json::to_string(&style).unwrap(), "{}");
    assert_eq!(
        serde_json::from_str::<ooxml_drawingml::ShapeStyle>("{}").unwrap(),
        style
    );
}
