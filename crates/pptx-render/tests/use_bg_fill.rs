use ooxml_drawingml::GeometryPathCommand as C;
use pptx_edit::{DeckSession, EditCtx};
use pptx_render::{
    GradientStop, GradientType, Paint, Primitive, RenderedSlide, SlideRenderer, Transform,
};

const FONT: &[u8] = include_bytes!("../../ooxml-text/tests/fonts/LiberationSans-Regular.ttf");
const GRADIENT: &str = r#"<a:gradFill><a:gsLst><a:gs pos="0"><a:srgbClr val="FF0000"/></a:gs><a:gs pos="100000"><a:srgbClr val="0000FF"/></a:gs></a:gsLst><a:lin ang="0" scaled="1"/></a:gradFill>"#;

const CONTENT_TYPES: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
<Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
<Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>
<Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>
<Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
</Types>"#;

const ROOT_RELS: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>
</Relationships>"#;

const PRESENTATION: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
<p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst>
<p:sldIdLst><p:sldId id="256" r:id="rId2"/></p:sldIdLst>
<p:sldSz cx="12192000" cy="6858000"/><p:notesSz cx="6858000" cy="9144000"/>
</p:presentation>"#;

const PRESENTATION_RELS: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>
</Relationships>"#;

const SLIDE_RELS: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
</Relationships>"#;

const LAYOUT: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" type="blank">
<p:cSld><p:spTree>
<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
<p:grpSpPr/>
</p:spTree></p:cSld></p:sldLayout>"#;

const LAYOUT_RELS: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/>
</Relationships>"#;

const MASTER: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
<p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld>
<p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst></p:sldMaster>"#;

const MASTER_RELS: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme1.xml"/>
</Relationships>"#;

const THEME: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Fixture">
<a:themeElements><a:clrScheme name="Fixture">
<a:dk1><a:srgbClr val="000000"/></a:dk1><a:lt1><a:srgbClr val="FFFFFF"/></a:lt1>
<a:dk2><a:srgbClr val="222222"/></a:dk2><a:lt2><a:srgbClr val="EEEEEE"/></a:lt2>
<a:accent1><a:srgbClr val="4472C4"/></a:accent1><a:accent2><a:srgbClr val="ED7D31"/></a:accent2>
<a:accent3><a:srgbClr val="A5A5A5"/></a:accent3><a:accent4><a:srgbClr val="FFC000"/></a:accent4>
<a:accent5><a:srgbClr val="5B9BD5"/></a:accent5><a:accent6><a:srgbClr val="70AD47"/></a:accent6>
<a:hlink><a:srgbClr val="0563C1"/></a:hlink><a:folHlink><a:srgbClr val="954F72"/></a:folHlink>
</a:clrScheme></a:themeElements></a:theme>"#;

fn slide(background: &str) -> String {
    format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
<p:cSld><p:bg><p:bgPr>{background}<a:effectLst/></p:bgPr></p:bg><p:spTree>
<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
<p:grpSpPr/>
<p:sp useBgFill="1"><p:nvSpPr><p:cNvPr id="2" name="Backdrop"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
<p:spPr><a:xfrm><a:off x="6096000" y="0"/><a:ext cx="6096000" cy="6858000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>
<p:style><a:lnRef idx="0"><a:schemeClr val="accent1"/></a:lnRef><a:fillRef idx="1"><a:schemeClr val="accent1"/></a:fillRef><a:effectRef idx="0"><a:schemeClr val="accent1"/></a:effectRef><a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef></p:style></p:sp>
</p:spTree></p:cSld></p:sld>"#
    )
}

fn fixture(background: &str) -> Vec<u8> {
    fixture_with_parts(&slide(background), LAYOUT)
}

fn fixture_with_parts(slide_xml: &str, layout_xml: &str) -> Vec<u8> {
    let parts: Vec<(String, Vec<u8>)> = [
        ("[Content_Types].xml", CONTENT_TYPES.to_owned()),
        ("_rels/.rels", ROOT_RELS.to_owned()),
        ("ppt/presentation.xml", PRESENTATION.to_owned()),
        (
            "ppt/_rels/presentation.xml.rels",
            PRESENTATION_RELS.to_owned(),
        ),
        ("ppt/slides/slide1.xml", slide_xml.to_owned()),
        ("ppt/slides/_rels/slide1.xml.rels", SLIDE_RELS.to_owned()),
        ("ppt/slideLayouts/slideLayout1.xml", layout_xml.to_owned()),
        (
            "ppt/slideLayouts/_rels/slideLayout1.xml.rels",
            LAYOUT_RELS.to_owned(),
        ),
        ("ppt/slideMasters/slideMaster1.xml", MASTER.to_owned()),
        (
            "ppt/slideMasters/_rels/slideMaster1.xml.rels",
            MASTER_RELS.to_owned(),
        ),
        ("ppt/theme/theme1.xml", THEME.to_owned()),
    ]
    .into_iter()
    .map(|(path, body)| (path.to_owned(), body.into_bytes()))
    .collect();
    ooxml_opc::rezip_parts(&parts).unwrap()
}

fn backdrop(background: &str, seed: u64) -> Primitive {
    let bytes = fixture(background);
    let session = DeckSession::open(&bytes, seed).unwrap();
    rendered(&session)
        .display_list
        .primitives
        .into_iter()
        .find(|primitive| {
            matches!(
                primitive,
                Primitive::Shape { object_id: 2, .. } | Primitive::Image { object_id: 2, .. }
            )
        })
        .expect("the backdrop shape is drawn")
}

fn rendered(session: &DeckSession) -> RenderedSlide {
    let snapshot = session.snapshot().unwrap();
    let mut renderer = SlideRenderer::new();
    renderer.register_font("Arial", false, false, FONT).unwrap();
    renderer
        .layout_slide(session.package(), &snapshot, 0)
        .unwrap()
}

#[test]
fn a_use_bg_fill_shape_takes_the_white_slide_background_over_its_style_fill() {
    let shape = backdrop(
        r#"<a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>"#,
        9_101,
    );
    let Primitive::Shape { fill, .. } = shape else {
        panic!("expected a shape primitive");
    };
    assert_eq!(
        fill,
        Some(Paint::Solid {
            color: "#FFFFFF".to_owned()
        })
    );
}

#[test]
fn a_use_bg_fill_shape_takes_a_gradient_background_in_slide_coordinates() {
    let shape = backdrop(
        r#"<a:gradFill><a:gsLst><a:gs pos="0"><a:srgbClr val="FF0000"/></a:gs><a:gs pos="100000"><a:srgbClr val="0000FF"/></a:gs></a:gsLst><a:lin ang="0" scaled="1"/></a:gradFill>"#,
        9_102,
    );
    let Primitive::Shape {
        x, w, fill, path, ..
    } = shape
    else {
        panic!("expected a shape primitive");
    };
    assert_eq!(
        fill,
        Some(Paint::Gradient {
            gradient_type: GradientType::Linear,
            angle_deg: Some(0.0),
            stops: vec![
                GradientStop {
                    position: 0.0,
                    color: "#FF0000".to_owned()
                },
                GradientStop {
                    position: 1.0,
                    color: "#0000FF".to_owned()
                },
            ],
        })
    );
    assert_eq!((x, w), (0.0, 1280.0));
    assert!(path.iter().any(|command| matches!(command, ooxml_drawingml::GeometryPathCommand::Move { x, .. } if (*x - 0.5).abs() < 1e-6)));
}

#[test]
fn background_fill_keeps_the_object_frame_for_hits_and_selection() {
    let session = DeckSession::open(&fixture(GRADIENT), 9_103).unwrap();
    let shape_id = session.snapshot().unwrap().slides[0].shapes[0].id.clone();
    let rendered = rendered(&session);
    let frames: Vec<_> = rendered
        .display_list
        .primitives
        .iter()
        .filter_map(|primitive| match primitive {
            Primitive::Shape {
                shape_id: Some(id),
                x,
                y,
                w,
                h,
                ..
            }
            | Primitive::Image {
                shape_id: Some(id),
                x,
                y,
                w,
                h,
                ..
            } if id == &shape_id => Some((*x, *y, *w, *h)),
            _ => None,
        })
        .collect();
    assert_eq!(frames, vec![(640.0, 0.0, 640.0, 720.0)]);
    assert!(rendered.hit_test(320.0, 360.0).is_none());
    assert!(rendered.hit_test(960.0, 360.0).is_some());
}

fn transformed_slide(background: &str, attributes: &str, geometry: &str) -> String {
    slide(background)
        .replace(
            r#"<a:xfrm><a:off x="6096000" y="0"/><a:ext cx="6096000" cy="6858000"/></a:xfrm>"#,
            &format!(
                r#"<a:xfrm {attributes}><a:off x="952500" y="1905000"/><a:ext cx="3810000" cy="1905000"/></a:xfrm>"#
            ),
        )
        .replace(r#"prst="rect""#, &format!(r#"prst="{geometry}""#))
}

fn assert_slide_path(primitive: &Primitive, expected: &[(f64, f64)]) {
    let (frame, transform, path) = match primitive {
        Primitive::Shape {
            x,
            y,
            w,
            h,
            transform,
            path,
            ..
        } => ((*x, *y, *w, *h), transform, path.as_slice()),
        Primitive::Image {
            x,
            y,
            w,
            h,
            transform,
            path: Some(path),
            ..
        } => ((*x, *y, *w, *h), transform, path.as_slice()),
        _ => panic!("expected a masked background fill"),
    };
    assert_eq!(frame, (0.0, 0.0, 1280.0, 720.0));
    assert!(transform.is_identity());
    let points: Vec<_> = path
        .iter()
        .filter_map(|command| match command {
            C::Move { x, y } | C::Line { x, y } => Some((x * 1280.0, y * 720.0)),
            _ => None,
        })
        .collect();
    assert_eq!(points.len(), expected.len(), "{path:?}");
    for (actual, expected) in points.iter().zip(expected) {
        assert!((actual.0 - expected.0).abs() < 1e-4, "{points:?}");
        assert!((actual.1 - expected.1).abs() < 1e-4, "{points:?}");
    }
}

#[test]
fn rotated_and_flipped_shapes_sample_the_slide_gradient() {
    for (attributes, expected) in [
        (
            r#"rot="5400000""#,
            [(400.0, 300.0), (200.0, 500.0), (200.0, 100.0)],
        ),
        (
            r#"rot="5400000" flipH="1" flipV="1""#,
            [(200.0, 300.0), (400.0, 100.0), (400.0, 500.0)],
        ),
    ] {
        let xml = transformed_slide(GRADIENT, attributes, "triangle");
        let session = DeckSession::open(&fixture_with_parts(&xml, LAYOUT), 9_104).unwrap();
        let rendered = rendered(&session);
        let fill = rendered
            .display_list
            .primitives
            .iter()
            .find(|primitive| matches!(primitive, Primitive::Shape { fill: Some(_), .. }))
            .unwrap();
        assert_slide_path(fill, &expected);
        let Primitive::Shape { fill, .. } = fill else {
            unreachable!()
        };
        assert_eq!(fill, &rendered.display_list.background);
    }
}

#[test]
fn rotated_picture_and_tile_backgrounds_keep_the_slide_frame() {
    for placement in [
        r#"<a:stretch><a:fillRect/></a:stretch>"#,
        r#"<a:tile sx="50000" sy="75000"/>"#,
    ] {
        let background = format!(
            r#"<a:blipFill><a:blip r:embed="rId2"/><a:srcRect l="10000"/>{placement}</a:blipFill>"#
        );
        let xml = transformed_slide(
            &background,
            r#"rot="5400000" flipH="1" flipV="1""#,
            "triangle",
        );
        let bytes = fixture_with_parts(&xml, LAYOUT);
        let mut parts = ooxml_opc::unzip_parts(&bytes).unwrap();
        for (path, bytes) in &mut parts {
            let xml = String::from_utf8(bytes.clone()).unwrap();
            if path == "ppt/slides/_rels/slide1.xml.rels" {
                *bytes = xml.replace("</Relationships>", r#"<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/></Relationships>"#).into_bytes();
            } else if path == "[Content_Types].xml" {
                *bytes = xml
                    .replace(
                        "</Types>",
                        r#"<Default Extension="png" ContentType="image/png"/></Types>"#,
                    )
                    .into_bytes();
            }
        }
        let image = ooxml_opc::unzip_parts(include_bytes!(
            "../../pptx-parse/tests/fixtures/picture-fill.pptx"
        ))
        .unwrap()
        .into_iter()
        .find(|(path, _)| path == "ppt/media/image1.png")
        .unwrap();
        parts.push(image);
        let session = DeckSession::open(&ooxml_opc::rezip_parts(&parts).unwrap(), 9_105).unwrap();
        let rendered = rendered(&session);
        let images: Vec<_> = rendered
            .display_list
            .primitives
            .iter()
            .filter(|primitive| matches!(primitive, Primitive::Image { .. }))
            .collect();
        assert_eq!(images.len(), 2);
        assert_slide_path(images[1], &[(200.0, 300.0), (400.0, 100.0), (400.0, 500.0)]);
        let Primitive::Image {
            asset_id,
            crop,
            tile,
            shape_id,
            stroke,
            ..
        } = images[1]
        else {
            unreachable!()
        };
        let Primitive::Image {
            asset_id: bg_asset,
            crop: bg_crop,
            tile: bg_tile,
            ..
        } = images[0]
        else {
            unreachable!()
        };
        assert_eq!((asset_id, crop, tile), (bg_asset, bg_crop, bg_tile));
        assert_eq!(tile.is_some(), placement.contains("tile"));
        assert!(shape_id.is_none());
        assert!(stroke.is_none());
    }
}

#[test]
fn grouped_background_clips_follow_the_group_transform() {
    let xml = transformed_slide(GRADIENT, "", "triangle")
        .replace(
            "</p:spPr>",
            r#"<a:ln w="38100"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln></p:spPr>"#,
        )
        .replace(
            r#"<p:sp useBgFill="1">"#,
            r#"<p:grpSp><p:nvGrpSpPr><p:cNvPr id="3" name="Group"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm rot="5400000" flipH="1"><a:off x="1905000" y="952500"/><a:ext cx="3810000" cy="3810000"/><a:chOff x="952500" y="1905000"/><a:chExt cx="3810000" cy="1905000"/></a:xfrm></p:grpSpPr><p:sp useBgFill="1">"#,
        )
        .replace("</p:sp>\n</p:spTree>", "</p:sp></p:grpSp>\n</p:spTree>");
    let start = xml.find("<p:grpSp>").unwrap();
    let end = xml.find("</p:grpSp>").unwrap() + "</p:grpSp>".len();
    let group = &xml[start..end];
    let layout = LAYOUT.replace("</p:spTree>", &format!("{group}</p:spTree>"));
    for bytes in [
        fixture_with_parts(&xml, LAYOUT),
        fixture_with_parts(&xml.replace(group, ""), &layout),
    ] {
        let session = DeckSession::open(&bytes, 9_106).unwrap();
        let rendered = rendered(&session);
        let fill = rendered
            .display_list
            .primitives
            .iter()
            .find(|primitive| matches!(primitive, Primitive::Shape { fill: Some(_), .. }))
            .unwrap();
        let (x, y, w, h, path, transform) = rendered
            .display_list
            .primitives
            .iter()
            .find_map(|primitive| match primitive {
                Primitive::Shape {
                    x,
                    y,
                    w,
                    h,
                    path,
                    stroke: Some(_),
                    transform,
                    ..
                } => Some((*x, *y, *w, *h, path, transform)),
                _ => None,
            })
            .unwrap();
        assert!(transform.is_identity());
        let points: Vec<_> = path
            .iter()
            .filter_map(|command| match command {
                C::Move { x: px, y: py } | C::Line { x: px, y: py } => Some((
                    f64::from(x) + px * f64::from(w),
                    f64::from(y) + py * f64::from(h),
                )),
                _ => None,
            })
            .collect();
        assert_slide_path(fill, &points);
    }
}

#[test]
fn rotated_background_fill_matches_outline_in_an_unequally_scaled_slide_group() {
    assert_rotated_group_background_matches_outline(false);
}

#[test]
fn rotated_background_fill_matches_outline_in_an_unequally_scaled_layout_group() {
    assert_rotated_group_background_matches_outline(true);
}

fn assert_rotated_group_background_matches_outline(in_layout: bool) {
    let xml = transformed_slide(GRADIENT, r#"rot="5400000""#, "rect")
        .replace(
            "</p:spPr>",
            r#"<a:ln w="38100"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln></p:spPr>"#,
        )
        .replace(
            r#"<p:sp useBgFill="1">"#,
            r#"<p:grpSp><p:nvGrpSpPr><p:cNvPr id="3" name="Group"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="1905000" y="952500"/><a:ext cx="7620000" cy="1905000"/><a:chOff x="952500" y="1905000"/><a:chExt cx="3810000" cy="1905000"/></a:xfrm></p:grpSpPr><p:sp useBgFill="1">"#,
        )
        .replace("</p:sp>\n</p:spTree>", "</p:sp></p:grpSp>\n</p:spTree>");
    let bytes = if in_layout {
        let start = xml.find("<p:grpSp>").unwrap();
        let end = xml.find("</p:grpSp>").unwrap() + "</p:grpSp>".len();
        let group = &xml[start..end];
        let layout = LAYOUT.replace("</p:spTree>", &format!("{group}</p:spTree>"));
        fixture_with_parts(&xml.replace(group, ""), &layout)
    } else {
        fixture_with_parts(&xml, LAYOUT)
    };
    let session = DeckSession::open(&bytes, 9_112).unwrap();
    let rendered = rendered(&session);
    let (x, y, w, h, transform) = rendered
        .display_list
        .primitives
        .iter()
        .find_map(|primitive| match primitive {
            Primitive::Shape {
                x,
                y,
                w,
                h,
                stroke: Some(_),
                transform,
                ..
            } => Some((*x, *y, *w, *h, *transform)),
            _ => None,
        })
        .unwrap();
    assert_eq!((x, y, w, h), (200.0, 100.0, 800.0, 200.0));
    assert_eq!(transform.rotation_deg, 90.0);
    let fill = rendered
        .display_list
        .primitives
        .iter()
        .find(|primitive| matches!(primitive, Primitive::Shape { fill: Some(_), .. }))
        .unwrap();
    assert_slide_path(
        fill,
        &[
            (700.0, -200.0),
            (700.0, 600.0),
            (500.0, 600.0),
            (500.0, -200.0),
        ],
    );
}

#[test]
fn background_fill_keeps_the_original_gradient_outline_frame() {
    let outline = format!(r#"<a:ln w="38100">{GRADIENT}</a:ln>"#);
    let outlines = |session: &DeckSession| {
        rendered(session)
            .display_list
            .primitives
            .into_iter()
            .filter_map(|primitive| match primitive {
                Primitive::Shape {
                    x,
                    y,
                    w,
                    h,
                    path,
                    stroke: Some(stroke),
                    transform,
                    ..
                } => {
                    assert!(matches!(
                        stroke.paint.as_ref(),
                        Some(Paint::Gradient { .. })
                    ));
                    Some((x, y, w, h, path, stroke, transform))
                }
                _ => None,
            })
            .collect::<Vec<_>>()
    };
    for attributes in ["", r#"rot="1800000" flipH="1""#] {
        let xml = transformed_slide(GRADIENT, attributes, "rect")
            .replace("</p:spPr>", &format!("{outline}</p:spPr>"));
        let normal = xml.replace(r#" useBgFill="1""#, "");
        let marked = DeckSession::open(&fixture_with_parts(&xml, LAYOUT), 9_107).unwrap();
        let unmarked = DeckSession::open(&fixture_with_parts(&normal, LAYOUT), 9_107).unwrap();
        assert_eq!(outlines(&marked), outlines(&unmarked));
        assert_eq!(outlines(&marked).len(), 1);
    }
}

#[test]
fn an_unmarked_shape_keeps_its_original_display_list() {
    let xml = slide(GRADIENT).replace(r#" useBgFill="1""#, "");
    let session = DeckSession::open(&fixture_with_parts(&xml, LAYOUT), 9_108).unwrap();
    let shape_id = session.snapshot().unwrap().slides[0].shapes[0].id.clone();
    let rendered = rendered(&session);
    assert_eq!(
        rendered.display_list.primitives,
        vec![Primitive::Shape {
            object_id: 2,
            shape_id: Some(shape_id),
            name: "Backdrop".to_owned(),
            x: 640.0,
            y: 0.0,
            w: 640.0,
            h: 720.0,
            geometry: "rect".to_owned(),
            path: vec![
                C::Move { x: 0.0, y: 0.0 },
                C::Line { x: 1.0, y: 0.0 },
                C::Line { x: 1.0, y: 1.0 },
                C::Line { x: 0.0, y: 1.0 },
                C::Close,
            ],
            geometry_fallback: false,
            clip: None,
            even_odd: false,
            adjust_values: Default::default(),
            fill: Some(Paint::Solid {
                color: "#4472C4".to_owned()
            }),
            stroke: None,
            shadow: None,
            transform: Transform::default(),
        }]
    );
}

#[test]
fn a_background_fill_shape_uses_the_layout_background() {
    let xml = slide("").replace("<p:bg><p:bgPr><a:effectLst/></p:bgPr></p:bg>", "");
    let layout = LAYOUT.replace(
        "<p:cSld>",
        &format!("<p:cSld><p:bg><p:bgPr>{GRADIENT}</p:bgPr></p:bg>"),
    );
    let session = DeckSession::open(&fixture_with_parts(&xml, &layout), 9_109).unwrap();
    let rendered = rendered(&session);
    let fill = rendered
        .display_list
        .primitives
        .iter()
        .find_map(|primitive| match primitive {
            Primitive::Shape {
                fill: Some(fill), ..
            } => Some(fill),
            _ => None,
        })
        .unwrap();
    assert!(matches!(fill, Paint::Gradient { .. }));
    assert_eq!(Some(fill), rendered.display_list.background.as_ref());
}

#[test]
fn a_background_fill_attribute_survives_an_edit_and_save_reopen() {
    let session = DeckSession::open(&fixture(GRADIENT), 9_110).unwrap();
    let snapshot = session.snapshot().unwrap();
    session
        .move_shape(
            &EditCtx::local("background"),
            &snapshot.slides[0].id,
            &snapshot.slides[0].shapes[0].id,
            952_500,
            1_905_000,
        )
        .unwrap();
    let saved = session.save().unwrap();
    let parts = ooxml_opc::unzip_parts(&saved).unwrap();
    let xml = &parts
        .iter()
        .find(|(path, _)| path == "ppt/slides/slide1.xml")
        .unwrap()
        .1;
    assert!(String::from_utf8_lossy(xml).contains(r#"useBgFill="1""#));
    let reopened = DeckSession::open(&saved, 9_111).unwrap();
    let pptx_parse::ShapeNode::Shape(shape) = &reopened.package().slides[0].shapes[0] else {
        panic!("expected a shape");
    };
    assert!(shape.use_bg_fill);
    assert_eq!(
        (shape.base.transform.x, shape.base.transform.y),
        (952_500, 1_905_000)
    );
    let rendered = rendered(&reopened);
    assert!(
        rendered
            .display_list
            .primitives
            .iter()
            .any(|primitive| matches!(
                primitive,
                Primitive::Shape {
                    fill: Some(Paint::Gradient { .. }),
                    ..
                }
            ))
    );
}
