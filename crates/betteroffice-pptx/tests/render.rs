#![cfg(feature = "raster")]

use std::collections::HashSet;
use std::io::Cursor;

use betteroffice_pptx::{ColorValue, EditCtx, Paint, Presentation, RenderOptions, ShapeFill};

const FIXTURE: &[u8] = include_bytes!("../../../apps/demo/public/betteroffice-demo.pptx");
const FONT: &[u8] = include_bytes!("../../ooxml-text/tests/fonts/LiberationSans-Regular.ttf");

const PNG_MAGIC: [u8; 8] = [0x89, b'P', b'N', b'G', 0x0d, 0x0a, 0x1a, 0x0a];

fn png_dimensions(png: &[u8]) -> (u32, u32) {
    let read = |offset: usize| u32::from_be_bytes(png[offset..offset + 4].try_into().unwrap());
    (read(16), read(20))
}

fn deck() -> Presentation {
    let mut presentation = Presentation::open(FIXTURE).unwrap();
    for bold in [false, true] {
        presentation
            .register_font("Arial", bold, false, FONT)
            .unwrap();
    }
    presentation
}

#[test]
fn a_written_background_reaches_the_display_list() {
    let presentation = deck();
    let slide_id = presentation.snapshot().unwrap().slides[0].id.clone();
    let fill = ShapeFill {
        fill_type: "solid".to_owned(),
        color: Some(ColorValue {
            rgb: Some("0F1720".to_owned()),
            ..ColorValue::default()
        }),
        gradient: None,
    };
    presentation
        .set_slide_background(&EditCtx::local("render-test"), &slide_id, Some(&fill))
        .unwrap();
    let saved = presentation.save().unwrap();

    let mut reopened = Presentation::open(&saved).unwrap();
    for bold in [false, true] {
        reopened.register_font("Arial", bold, false, FONT).unwrap();
    }
    let rendered = reopened.render_slide(0).unwrap();
    let Some(Paint::Solid { color }) = rendered.display_list.background else {
        panic!("expected a solid background paint");
    };
    assert_eq!(color.trim_start_matches('#').to_ascii_uppercase(), "0F1720");
}

#[test]
fn renders_every_slide_of_the_demo_deck() {
    let presentation = deck();
    for index in 0..presentation.slides().len() {
        let rendered = presentation
            .render_png(index, &RenderOptions::default())
            .unwrap();
        assert_eq!(rendered.bytes[..8], PNG_MAGIC);
        assert_eq!((rendered.width, rendered.height), (1280, 720));
        assert_eq!(png_dimensions(&rendered.bytes), (1280, 720));
        assert_eq!(rendered.skipped_images, 0);
    }
}

#[test]
fn scale_multiplies_the_output_dimensions() {
    let rendered = deck()
        .render_png(
            0,
            &RenderOptions {
                scale: 2.0,
                ..RenderOptions::default()
            },
        )
        .unwrap();
    assert_eq!(png_dimensions(&rendered.bytes), (2560, 1440));
}

#[test]
fn a_slide_past_the_deck_is_refused() {
    let error = deck()
        .render_png(99, &RenderOptions::default())
        .unwrap_err();
    assert!(error.to_string().contains("outside the deck"), "{error}");
}

#[test]
fn text_without_a_registered_font_needs_one_for_layout_and_png() {
    let presentation = Presentation::open(FIXTURE).unwrap();
    let Err(error) = presentation.render_slide(0) else {
        panic!("slide text laid out without a registered font");
    };
    assert!(error.to_string().contains("no font"), "{error}");
    let error = presentation
        .render_png(0, &RenderOptions::default())
        .unwrap_err();
    assert!(error.to_string().contains("no font"), "{error}");
}

fn text_container_deck(content: &str) -> Vec<u8> {
    let parts = [
        ("[Content_Types].xml", r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/><Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/></Types>"#.to_owned()),
        ("_rels/.rels", r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="doc" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/></Relationships>"#.to_owned()),
        ("ppt/presentation.xml", r#"<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:sldIdLst><p:sldId id="256" r:id="slide"/></p:sldIdLst><p:sldSz cx="9144000" cy="5143500"/></p:presentation>"#.to_owned()),
        ("ppt/_rels/presentation.xml.rels", r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="slide" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/></Relationships>"#.to_owned()),
        ("ppt/slides/slide1.xml", format!(r#"<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>{content}</p:spTree></p:cSld></p:sld>"#)),
    ];
    let parts: Vec<_> = parts
        .into_iter()
        .map(|(path, xml)| (path.to_owned(), xml.into_bytes()))
        .collect();
    ooxml_opc::rezip_parts(&parts).unwrap()
}

#[test]
fn nested_text_without_fonts_is_refused() {
    let table = r#"<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="2" name="Table"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="100000" y="100000"/><a:ext cx="3000000" cy="1000000"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table"><a:tbl><a:tblPr/><a:tblGrid><a:gridCol w="3000000"/></a:tblGrid><a:tr h="1000000"><a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr sz="2400"/><a:t>Cell text</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc></a:tr></a:tbl></a:graphicData></a:graphic></p:graphicFrame>"#;
    let group = r#"<p:grpSp><p:nvGrpSpPr><p:cNvPr id="2" name="Group"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="3000000" cy="1000000"/><a:chOff x="0" y="0"/><a:chExt cx="3000000" cy="1000000"/></a:xfrm></p:grpSpPr><p:sp><p:nvSpPr><p:cNvPr id="3" name="Text"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="3000000" cy="1000000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr sz="2400"/><a:t>Grouped text</a:t></a:r></a:p></p:txBody></p:sp></p:grpSp>"#;
    for content in [table, group] {
        let bytes = text_container_deck(content);
        let mut presentation = Presentation::open(&bytes).unwrap();
        let Err(error) = presentation.render_slide(0) else {
            panic!("nested text rendered without a registered font");
        };
        assert!(error.to_string().contains("no font"), "{error}");
        let error = presentation
            .render_png(0, &RenderOptions::default())
            .unwrap_err();
        assert!(error.to_string().contains("no font"), "{error}");
        presentation
            .register_font("Arial", false, false, FONT)
            .unwrap();
        let rendered = presentation.render_slide(0).unwrap();
        assert!(rendered.display_list.has_glyphs());
        if content == table {
            assert!(rendered.display_list.primitives.iter().any(|primitive| {
                matches!(primitive, betteroffice_pptx::Primitive::Table { .. })
            }));
            assert!(!rendered.display_list.primitives.iter().any(|primitive| {
                matches!(primitive, betteroffice_pptx::Primitive::TextBox { .. })
            }));
        } else {
            assert!(rendered.display_list.primitives.iter().any(|primitive| {
                matches!(primitive, betteroffice_pptx::Primitive::TextBox { .. })
            }));
        }
    }
}

#[test]
fn a_text_free_slide_renders_png_without_fonts() {
    let source = include_bytes!("../../pptx-render/tests/fixtures/blip-lum.pptx");
    let presentation = Presentation::open(source).unwrap();
    let rendered = presentation.render_slide(0).unwrap();
    let has_text = rendered.display_list.primitives.iter().any(|primitive| {
        matches!(primitive, betteroffice_pptx::Primitive::TextBox { lines, .. }
            if lines.iter().flat_map(|line| &line.runs).any(|run| !run.glyphs.is_empty()))
    });
    assert!(!has_text);
    let png = presentation
        .render_png(0, &RenderOptions::default())
        .unwrap();
    assert_eq!(png.bytes[..8], PNG_MAGIC);
}

#[test]
fn a_scale_that_is_not_positive_is_refused() {
    let error = deck()
        .render_png(
            0,
            &RenderOptions {
                scale: 0.0,
                ..RenderOptions::default()
            },
        )
        .unwrap_err();
    assert!(error.to_string().contains("finite and positive"), "{error}");
}

/// A wiring bug that resolves no primitives still encodes a valid PNG, so the
/// bytes alone prove nothing: the pixels have to carry slide content.
#[test]
fn the_render_is_not_a_blank_surface() {
    let rendered = deck().render_png(0, &RenderOptions::default()).unwrap();
    let decoder = png::Decoder::new(Cursor::new(&rendered.bytes));
    let mut reader = decoder.read_info().unwrap();
    let mut pixels = vec![0; reader.output_buffer_size().unwrap()];
    let info = reader.next_frame(&mut pixels).unwrap();
    let distinct: HashSet<[u8; 4]> = pixels[..info.buffer_size()]
        .as_chunks::<4>()
        .0
        .iter()
        .copied()
        .collect();
    assert!(
        distinct.len() > 16,
        "slide painted only {} distinct colors",
        distinct.len()
    );
}

/// The glyph cache lives on the deck, so a second export of the same slide must
/// still produce the same bytes.
#[test]
fn repeated_renders_are_byte_identical() {
    let presentation = deck();
    let first = presentation
        .render_png(1, &RenderOptions::default())
        .unwrap();
    let second = presentation
        .render_png(1, &RenderOptions::default())
        .unwrap();
    assert_eq!(first.bytes, second.bytes);
}

/// Writes every slide as a png when `PPTX_RENDER_DUMP` names a directory, so a
/// change can be eyeballed rather than only byte-compared.
#[test]
fn dumps_slides_when_asked() {
    let Ok(directory) = std::env::var("PPTX_RENDER_DUMP") else {
        return;
    };
    let presentation = deck();
    for index in 0..presentation.slides().len() {
        let rendered = presentation
            .render_png(
                index,
                &RenderOptions {
                    scale: 2.0,
                    ..RenderOptions::default()
                },
            )
            .unwrap();
        std::fs::write(format!("{directory}/slide-{index}.png"), &rendered.bytes).unwrap();
    }
}

#[test]
fn lum_brightness_and_contrast_reach_the_bitmap() {
    use betteroffice_pptx::{ImageEffect, Primitive};
    let source = include_bytes!("../../pptx-render/tests/fixtures/blip-lum.pptx");
    let deck = Presentation::open(source).unwrap();
    let list = deck.render_slide(0).unwrap().display_list;
    let effects: Vec<_> = list
        .primitives
        .iter()
        .filter_map(|primitive| match primitive {
            Primitive::Image { name, effects, .. } => Some((name.as_str(), effects.as_slice())),
            _ => None,
        })
        .collect();
    assert_eq!(effects.len(), 5);
    assert_eq!(effects[0], ("Control", [].as_slice()));
    assert_eq!(
        effects[1].1,
        [ImageEffect::Luminance {
            brightness: 0.7,
            contrast: -0.7
        }]
    );

    let png = deck.render_png(0, &RenderOptions::default()).unwrap();
    let mut reader = png::Decoder::new(Cursor::new(&png.bytes))
        .read_info()
        .unwrap();
    let mut pixels = vec![0; reader.output_buffer_size().unwrap()];
    let info = reader.next_frame(&mut pixels).unwrap();
    let pixel =
        |x, y| &pixels[(y * info.width as usize + x) * 4..(y * info.width as usize + x) * 4 + 4];
    for (x, y, expected) in [
        (40, 40, [0, 0, 0, 255]),
        (280, 40, [3, 167, 223, 255]),
        (40, 100, [205, 205, 205, 255]),
        (88, 100, [225, 225, 225, 255]),
        (280, 100, [206, 255, 255, 255]),
        (40, 160, [128, 128, 128, 255]),
        (40, 220, [64, 64, 64, 255]),
        (232, 220, [192, 192, 192, 255]),
        (280, 220, [65, 148, 176, 255]),
        (136, 280, [148, 148, 148, 255]),
    ] {
        assert_eq!(pixel(x, y), expected, "pixel ({x}, {y})");
    }
}

#[test]
fn blip_effects_render_on_slides_layouts_and_masters() {
    use betteroffice_pptx::{ImageEffect, Primitive};
    let source = include_bytes!("../../pptx-render/tests/fixtures/blip-effects.pptx");
    let deck = Presentation::open(source).unwrap();
    let list = deck.render_slide(0).unwrap().display_list;
    let effects: Vec<_> = list
        .primitives
        .iter()
        .filter_map(|primitive| match primitive {
            Primitive::Image { name, effects, .. } => Some((name.as_str(), effects.as_slice())),
            _ => None,
        })
        .collect();
    assert_eq!(effects.len(), 8);
    let duo = ImageEffect::Duotone {
        shadow: "#B3B3B3FF".into(),
        highlight: "#FFFFFFFF".into(),
    };
    let bilevel = ImageEffect::BiLevel { threshold: 0.5 };
    assert_eq!(effects[0], ("Master duotone", std::slice::from_ref(&duo)));
    assert_eq!(
        effects[1],
        ("Layout biLevel", std::slice::from_ref(&bilevel))
    );
    assert_eq!(effects[2], ("Control", [].as_slice()));
    assert_eq!(effects[3], ("BiLevel", std::slice::from_ref(&bilevel)));
    assert_eq!(effects[4], ("Duotone", std::slice::from_ref(&duo)));
    assert!(matches!(
        effects[6].1,
        [ImageEffect::ColorChange {
            use_alpha: false,
            ..
        }]
    ));
    let png = deck.render_png(0, &RenderOptions::default()).unwrap();
    let mut reader = png::Decoder::new(Cursor::new(&png.bytes))
        .read_info()
        .unwrap();
    let mut pixels = vec![0; reader.output_buffer_size().unwrap()];
    let info = reader.next_frame(&mut pixels).unwrap();
    assert_eq!(info.color_type, png::ColorType::Rgba);
    let pixel =
        |x, y| &pixels[(y * info.width as usize + x) * 4..(y * info.width as usize + x) * 4 + 4];
    for (x, y, expected) in [
        (40, 40, [3, 167, 223, 255]),
        (40, 100, [0, 0, 0, 255]),
        (40, 160, [216, 216, 216, 255]),
        (136, 220, [208, 208, 208, 255]),
        (136, 280, [255, 0, 0, 255]),
        (232, 280, [232, 104, 104, 255]),
        (40, 340, [124, 124, 124, 255]),
        (360, 100, [0, 0, 0, 255]),
        (360, 160, [216, 216, 216, 255]),
    ] {
        assert_eq!(pixel(x, y), expected, "pixel ({x}, {y})");
    }
}

#[test]
fn pictures_received_from_a_peer_render_before_save() {
    let presentation = deck();
    let before = presentation
        .render_png(0, &RenderOptions::default())
        .unwrap();
    let peer = pptx_edit::DeckSession::open(FIXTURE, 12).unwrap();
    let slide_id = peer.snapshot().unwrap().slides[0].id.clone();
    let mut bytes = Vec::new();
    {
        let mut encoder = png::Encoder::new(&mut bytes, 1, 1);
        encoder.set_color(png::ColorType::Rgba);
        encoder.set_depth(png::BitDepth::Eight);
        encoder
            .write_header()
            .unwrap()
            .write_image_data(&[255, 0, 0, 255])
            .unwrap();
    }
    let added = peer
        .add_picture(
            &pptx_edit::EditCtx::local("peer"),
            &slide_id,
            &pptx_edit::PictureDraft {
                name: "Red square".to_owned(),
                rect: pptx_edit::ShapeRect {
                    x: 0,
                    y: 0,
                    width: 1_000_000,
                    height: 1_000_000,
                },
                content_type: "image/png".to_owned(),
                media_bytes: bytes.clone(),
            },
        )
        .unwrap();
    presentation
        .apply_update_v1(&peer.encode_state_as_update_v1())
        .unwrap();
    assert_eq!(
        presentation
            .media_bytes(&format!("pending-media:{}", added.shape_id))
            .unwrap(),
        bytes
    );
    let rendered = presentation
        .render_png(0, &RenderOptions::default())
        .unwrap();
    assert_eq!(rendered.skipped_images, 0);
    assert_ne!(rendered.bytes, before.bytes);
}
