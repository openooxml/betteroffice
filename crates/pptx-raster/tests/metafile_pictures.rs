use ooxml_text::FontStore;
use pptx_edit::DeckSession;
use pptx_raster::{
    AssetMap, GlyphCache, ImageCache, RenderOptions, RenderResources, render_png,
    render_slide_shared,
};
use pptx_render::SlideRenderer;
use tiny_skia::Pixmap;

#[test]
fn text_and_bitmap_with_line_paint_inside_their_picture_frames() {
    let fixture = include_bytes!("../../pptx-render/tests/fixtures/metafile-svg-pictures.pptx");
    let session = DeckSession::open(fixture, 320).unwrap();
    let saved = session.save().unwrap();
    let list = SlideRenderer::new()
        .layout_slide(session.package(), &session.snapshot().unwrap(), 0)
        .unwrap()
        .display_list;
    let mut fonts = FontStore::new();
    fonts
        .register(
            include_bytes!("../../../packages/fonts/assets/LiberationSans-Regular.ttf").to_vec(),
        )
        .unwrap();
    let images: AssetMap<'_> = session
        .package()
        .media
        .iter()
        .map(|part| (part.part_path.as_str(), part.bytes.as_slice()))
        .collect();
    let png = render_png(&list, &RenderResources::new(&fonts, &images)).unwrap();
    let pixels = Pixmap::decode_png(&png).unwrap();
    for left in [96, 216] {
        assert!(
            (96..192).any(|y| (left..left + 96).any(|x| {
                let pixel = pixels.pixel(x, y).unwrap();
                pixel.red() != 255 || pixel.green() != 255 || pixel.blue() != 255
            })),
            "picture at {left} is blank"
        );
    }
    assert_eq!(pixels.pixel(95, 95).unwrap().red(), 255);
    assert_eq!(session.save().unwrap(), saved);
}

#[test]
fn a_metafile_picture_rerenders_when_a_font_is_registered() {
    let fixture = include_bytes!("../../pptx-render/tests/fixtures/metafile-svg-pictures.pptx");
    let session = DeckSession::open(fixture, 320).unwrap();
    let list = SlideRenderer::new()
        .layout_slide(session.package(), &session.snapshot().unwrap(), 0)
        .unwrap()
        .display_list;
    let images: AssetMap<'_> = session
        .package()
        .media
        .iter()
        .map(|part| (part.part_path.as_str(), part.bytes.as_slice()))
        .collect();
    let mut fonts = FontStore::new();
    let mut cache = ImageCache::default();
    let first = render_slide_shared(
        &list,
        &RenderResources::new(&fonts, &images),
        &RenderOptions::default(),
        &mut GlyphCache::default(),
        &mut cache,
    )
    .unwrap();
    fonts
        .register(
            include_bytes!("../../../packages/fonts/assets/LiberationSans-Regular.ttf").to_vec(),
        )
        .unwrap();
    let second = render_slide_shared(
        &list,
        &RenderResources::new(&fonts, &images),
        &RenderOptions::default(),
        &mut GlyphCache::default(),
        &mut cache,
    )
    .unwrap();
    assert_ne!(first.bytes, second.bytes);
}
