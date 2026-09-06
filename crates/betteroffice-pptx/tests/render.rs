#![cfg(feature = "raster")]

use std::collections::HashSet;
use std::io::Cursor;

use betteroffice_pptx::{Presentation, RenderOptions};

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
fn text_without_a_registered_font_is_refused_by_layout() {
    let presentation = Presentation::open(FIXTURE).unwrap();
    let error = presentation
        .render_png(0, &RenderOptions::default())
        .unwrap_err();
    assert!(error.to_string().contains("no font"), "{error}");
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
        shadow: "#737373FF".into(),
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
        (40, 160, [183, 183, 183, 255]),
        (136, 220, [208, 208, 208, 255]),
        (136, 280, [255, 0, 0, 255]),
        (232, 280, [232, 104, 104, 255]),
        (40, 340, [124, 124, 124, 255]),
        (360, 100, [0, 0, 0, 255]),
        (360, 160, [183, 183, 183, 255]),
    ] {
        assert_eq!(pixel(x, y), expected, "pixel ({x}, {y})");
    }
}
