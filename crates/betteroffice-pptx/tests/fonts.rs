#[cfg(feature = "fonts")]
use betteroffice_pptx::Primitive;
#[cfg(feature = "fonts-latin")]
use betteroffice_pptx::fonts::{BundledFonts, FontProvider};
#[cfg(feature = "fonts")]
use betteroffice_pptx::fonts::{FontDefinition, FontSource};
use betteroffice_pptx::{OpenOptions, Presentation};

#[cfg(feature = "fonts")]
const APP_FONT: &[u8] = include_bytes!("../../ooxml-text/tests/fonts/LiberationSans-Regular.ttf");

fn backstop() -> OpenOptions {
    let options = OpenOptions::default();
    #[cfg(feature = "fonts")]
    let options = OpenOptions {
        font_source: FontSource::Backstop,
        ..options
    };
    options
}

fn fixture(embedded: Option<&[u8]>) -> Vec<u8> {
    let embedded_xml = if embedded.is_some() {
        r#"<p:embeddedFontLst><p:embeddedFont><p:font typeface="Calibri"/><p:regular r:id="font"/><p:bold r:id="font"/></p:embeddedFont></p:embeddedFontLst>"#
    } else {
        ""
    };
    let mut parts = vec![
        ("[Content_Types].xml".into(), br#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/><Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/></Types>"#.to_vec()),
        ("_rels/.rels".into(), br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="doc" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/></Relationships>"#.to_vec()),
        ("ppt/presentation.xml".into(), format!(r#"<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:sldIdLst><p:sldId id="256" r:id="slide"/></p:sldIdLst><p:sldSz cx="9144000" cy="5143500"/>{embedded_xml}</p:presentation>"#).into_bytes()),
        ("ppt/_rels/presentation.xml.rels".into(), br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="slide" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/><Relationship Id="font" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/font" Target="fonts/face.fntdata"/></Relationships>"#.to_vec()),
        ("ppt/slides/slide1.xml".into(), br#"<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/><p:sp><p:nvSpPr><p:cNvPr id="2" name="Text"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="8000000" cy="3000000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr sz="2400"><a:latin typeface="Calibri"/></a:rPr><a:t>Office regular</a:t></a:r></a:p><a:p><a:r><a:rPr sz="2400" b="1"><a:latin typeface="Calibri"/></a:rPr><a:t>Office bold</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:sld>"#.to_vec()),
    ];
    if let Some(bytes) = embedded {
        parts.push(("ppt/fonts/face.fntdata".into(), bytes.to_vec()));
    }
    ooxml_opc::rezip_parts(&parts).unwrap()
}

#[cfg(feature = "fonts")]
fn glyphs(presentation: &Presentation) -> Vec<Vec<(u32, f32)>> {
    presentation
        .render_slide(0)
        .unwrap()
        .display_list
        .primitives
        .iter()
        .filter_map(|primitive| match primitive {
            Primitive::TextBox { lines, .. } => Some(lines),
            _ => None,
        })
        .flatten()
        .flat_map(|line| &line.runs)
        .map(|run| {
            run.glyphs
                .iter()
                .map(|glyph| (glyph.glyph_id, glyph.advance))
                .collect()
        })
        .collect()
}

#[cfg(feature = "fonts-latin")]
fn private_use_font() -> Vec<u8> {
    private_use_font_with_glyph(2)
}

#[cfg(feature = "fonts-latin")]
fn private_use_font_with_glyph(glyph: u32) -> Vec<u8> {
    let mut bytes =
        include_bytes!("../../betteroffice-fonts/tests/fixtures/Synthetic-Regular.ttf").to_vec();
    let tables = u16::from_be_bytes(bytes[4..6].try_into().unwrap()) as usize;
    let record = (0..tables)
        .map(|index| 12 + index * 16)
        .find(|offset| &bytes[*offset..*offset + 4] == b"cmap")
        .unwrap();
    let mut cmap = vec![0, 0, 0, 1, 0, 3, 0, 10, 0, 0, 0, 12, 0, 12, 0, 0];
    for value in [40_u32, 0, 2, 0x44, 0x44, 3, 0xe000, 0xe000, glyph] {
        cmap.extend_from_slice(&value.to_be_bytes());
    }
    bytes.resize(bytes.len().next_multiple_of(4), 0);
    let offset = bytes.len() as u32;
    let checksum = cmap.as_chunks::<4>().0.iter().fold(0_u32, |sum, word| {
        sum.wrapping_add(u32::from_be_bytes(*word))
    });
    bytes[record + 4..record + 8].copy_from_slice(&checksum.to_be_bytes());
    bytes[record + 8..record + 12].copy_from_slice(&offset.to_be_bytes());
    bytes[record + 12..record + 16].copy_from_slice(&(cmap.len() as u32).to_be_bytes());
    bytes.extend_from_slice(&cmap);
    let head = (0..tables)
        .map(|index| 12 + index * 16)
        .find(|offset| &bytes[*offset..*offset + 4] == b"head")
        .unwrap();
    let head = u32::from_be_bytes(bytes[head + 8..head + 12].try_into().unwrap()) as usize;
    bytes[head + 8..head + 12].fill(0);
    let checksum = bytes.as_chunks::<4>().0.iter().fold(0_u32, |sum, word| {
        sum.wrapping_add(u32::from_be_bytes(*word))
    });
    bytes[head + 8..head + 12]
        .copy_from_slice(&0xb1b0_afba_u32.wrapping_sub(checksum).to_be_bytes());
    bytes
}

#[cfg(feature = "fonts-latin")]
fn private_use_fixture(embedded: Option<&[u8]>) -> Vec<u8> {
    let mut parts = ooxml_opc::unzip_parts(&fixture(embedded)).unwrap();
    for (name, bytes) in &mut parts {
        if name == "ppt/slides/slide1.xml" {
            *bytes = String::from_utf8(bytes.clone())
                .unwrap()
                .replace("Office regular", "&#xE000;")
                .replace("Office bold", "&#xE000;")
                .into_bytes();
        } else if name == "ppt/presentation.xml" {
            *bytes = String::from_utf8(bytes.clone())
                .unwrap()
                .replace(r#"<p:bold r:id="font"/>"#, "")
                .into_bytes();
        }
    }
    ooxml_opc::rezip_parts(&parts).unwrap()
}

#[cfg(feature = "fonts-latin")]
fn two_family_fixture() -> Vec<u8> {
    let slide = |family: &str| {
        format!(r#"<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/><p:sp><p:nvSpPr><p:cNvPr id="2" name="Text"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="8000000" cy="3000000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr sz="2400" b="1"><a:latin typeface="{family}"/></a:rPr><a:t>&#xE000;</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:sld>"#).into_bytes()
    };
    let parts = vec![
        ("[Content_Types].xml".into(), br#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/><Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/><Override PartName="/ppt/slides/slide2.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/></Types>"#.to_vec()),
        ("_rels/.rels".into(), br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="doc" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/></Relationships>"#.to_vec()),
        ("ppt/presentation.xml".into(), br#"<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:sldIdLst><p:sldId id="256" r:id="first"/><p:sldId id="257" r:id="second"/></p:sldIdLst><p:sldSz cx="9144000" cy="5143500"/></p:presentation>"#.to_vec()),
        ("ppt/_rels/presentation.xml.rels".into(), br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="first" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/><Relationship Id="second" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide2.xml"/></Relationships>"#.to_vec()),
        ("ppt/slides/slide1.xml".into(), slide("Calibri")),
        ("ppt/slides/slide2.xml".into(), slide("Arial")),
    ];
    ooxml_opc::rezip_parts(&parts).unwrap()
}

#[cfg(feature = "fonts-latin")]
fn slide_glyphs(presentation: &Presentation, index: usize) -> Vec<u32> {
    presentation
        .render_slide(index)
        .unwrap()
        .display_list
        .primitives
        .iter()
        .filter_map(|primitive| match primitive {
            Primitive::TextBox { lines, .. } => Some(lines),
            _ => None,
        })
        .flatten()
        .flat_map(|line| &line.runs)
        .flat_map(|run| run.glyphs.iter().map(|glyph| glyph.glyph_id))
        .collect()
}

#[cfg(feature = "fonts-latin")]
#[test]
fn another_family_keeps_each_family_coverage_order() {
    let mut presentation = Presentation::open(&two_family_fixture()).unwrap();
    for (family, glyph) in [("Calibri", 2), ("Arial", 3)] {
        presentation
            .register_font(family, false, false, &private_use_font_with_glyph(glyph))
            .unwrap();
        presentation
            .register_font(family, true, false, APP_FONT)
            .unwrap();
    }
    let first = slide_glyphs(&presentation, 0);
    assert_eq!(first, [2]);
    assert_eq!(slide_glyphs(&presentation, 1), [3]);
    assert_eq!(slide_glyphs(&presentation, 0), first);
}

#[cfg(feature = "fonts-latin")]
#[test]
fn registered_regular_face_handles_bold_private_use_text() {
    let bytes = private_use_fixture(None);
    let font = private_use_font();
    let mut expected = Presentation::open_with_options(&bytes, &backstop()).unwrap();
    expected
        .register_font("Calibri", false, false, &font)
        .unwrap();
    let expected = glyphs(&expected);
    assert_eq!(expected.len(), 2);
    assert!(expected.iter().all(|run| run.len() == 1 && run[0].0 == 2));
    for render_before_registration in [false, true] {
        let mut presentation = Presentation::open(&bytes).unwrap();
        if render_before_registration {
            presentation.render_slide(0).unwrap();
        }
        let id = presentation
            .register_font("Calibri", false, false, &font)
            .unwrap();
        assert_eq!(glyphs(&presentation), expected);
        let rendered = presentation.render_slide(0).unwrap();
        for primitive in &rendered.display_list.primitives {
            if let Primitive::TextBox { lines, .. } = primitive {
                for run in lines.iter().flat_map(|line| &line.runs) {
                    assert_eq!(run.font_id, id);
                }
            }
        }
    }
}

#[cfg(feature = "fonts-latin")]
#[test]
fn embedded_regular_face_handles_bold_private_use_text() {
    let font = private_use_font();
    let presentation = Presentation::open(&private_use_fixture(Some(&font))).unwrap();
    let mut expected =
        Presentation::open_with_options(&private_use_fixture(None), &backstop()).unwrap();
    expected
        .register_font("Calibri", false, false, &font)
        .unwrap();
    let actual = glyphs(&presentation);
    assert_eq!(actual.len(), 2);
    assert!(actual.iter().all(|run| run.len() == 1 && run[0].0 == 2));
    assert_eq!(actual, glyphs(&expected));
}

#[cfg(feature = "fonts-latin")]
#[test]
fn application_regular_coverage_precedes_bundled_bold() {
    let embedded = include_bytes!("../../betteroffice-fonts/tests/fixtures/Synthetic-Regular.ttf");
    let app = private_use_font();
    let mut parts = ooxml_opc::unzip_parts(&private_use_fixture(Some(embedded))).unwrap();
    for (name, bytes) in &mut parts {
        if name == "ppt/slides/slide1.xml" {
            *bytes = String::from_utf8(bytes.clone())
                .unwrap()
                .replace("&#xE000;", "&#xE000;D")
                .into_bytes();
        }
    }
    let bytes = ooxml_opc::rezip_parts(&parts).unwrap();
    for render_before_registration in [false, true] {
        let mut presentation = Presentation::open(&bytes).unwrap();
        if render_before_registration {
            presentation.render_slide(0).unwrap();
        }
        let app_id = presentation
            .register_font("Calibri", false, false, &app)
            .unwrap();
        let rendered = presentation.render_slide(0).unwrap();
        let mut glyphs = Vec::new();
        for primitive in &rendered.display_list.primitives {
            if let Primitive::TextBox { lines, .. } = primitive {
                for run in lines.iter().flat_map(|line| &line.runs) {
                    assert_eq!(run.font_id, app_id);
                    glyphs.extend(run.glyphs.iter().map(|glyph| glyph.glyph_id));
                }
            }
        }
        assert_eq!(glyphs, [2, 3, 2, 3]);
    }
}

#[cfg(feature = "fonts-latin")]
#[test]
fn replaced_primary_face_does_not_remain_in_coverage() {
    let bytes = private_use_fixture(None);
    let mut presentation = Presentation::open_with_options(&bytes, &backstop()).unwrap();
    presentation
        .register_font("Calibri", false, false, &private_use_font())
        .unwrap();
    let first = glyphs(&presentation);
    presentation
        .register_font("Calibri", false, false, APP_FONT)
        .unwrap();
    let mut expected = Presentation::open_with_options(&bytes, &backstop()).unwrap();
    expected
        .register_font("Calibri", false, false, APP_FONT)
        .unwrap();
    assert_ne!(first, glyphs(&expected));
    assert_eq!(glyphs(&presentation), glyphs(&expected));
}

#[cfg(feature = "fonts")]
#[test]
fn single_eot_header_keeps_embedded_font_rendering() {
    let mut font = vec![0; 82];
    font[4..8].copy_from_slice(&(APP_FONT.len() as u32).to_le_bytes());
    font[34..36].copy_from_slice(&[0x4c, 0x50]);
    font.extend_from_slice(APP_FONT);
    let presentation = Presentation::open_with_options(&fixture(Some(&font)), &backstop()).unwrap();
    let mut expected = Presentation::open_with_options(&fixture(None), &backstop()).unwrap();
    for bold in [false, true] {
        expected
            .register_font("Calibri", bold, false, APP_FONT)
            .unwrap();
    }
    let actual = glyphs(&presentation);
    assert_eq!(actual.len(), 2);
    assert!(actual.iter().all(|run| !run.is_empty()));
    assert_eq!(actual, glyphs(&expected));
}

#[cfg(feature = "fonts")]
#[test]
fn nested_embedded_font_headers_are_skipped_without_failing_open() {
    let depth = 4096;
    let mut font = vec![0; depth * 82];
    for index in 0..depth {
        let offset = index * 82;
        let length = (depth - index - 1) * 82 + APP_FONT.len();
        font[offset + 4..offset + 8].copy_from_slice(&(length as u32).to_le_bytes());
        font[offset + 34..offset + 36].copy_from_slice(&[0x4c, 0x50]);
    }
    font.extend_from_slice(APP_FONT);
    let presentation = Presentation::open_with_options(&fixture(Some(&font)), &backstop()).unwrap();
    assert!(matches!(
        presentation.render_slide(0),
        Err(betteroffice_pptx::Error::Render(
            betteroffice_pptx::RenderError::NoFont
        ))
    ));
}

#[cfg(feature = "fonts-latin")]
#[test]
fn default_calibri_uses_regular_and_bold_carlito() {
    let bytes = fixture(None);
    let presentation = Presentation::open(&bytes).unwrap();
    let mut expected = Presentation::open_with_options(&bytes, &backstop()).unwrap();
    for bold in [false, true] {
        let face = BundledFonts::default()
            .resolve_family("Calibri", bold, false)
            .unwrap();
        assert_eq!(face.family, "Carlito");
        expected
            .register_font("Calibri", bold, false, &face.load().unwrap())
            .unwrap();
    }
    let actual = glyphs(&presentation);
    assert_eq!(actual.len(), 2);
    assert!(actual.iter().all(|glyphs| !glyphs.is_empty()));
    assert_eq!(actual, glyphs(&expected));
}

#[cfg(feature = "fonts")]
#[test]
fn registered_and_option_faces_win_over_the_source() {
    let bytes = fixture(None);
    let mut presentation = Presentation::open(&bytes).unwrap();
    #[cfg(feature = "fonts-latin")]
    assert!(!glyphs(&presentation).is_empty());
    let mut expected = Presentation::open_with_options(&bytes, &backstop()).unwrap();
    for bold in [false, true] {
        presentation
            .register_font("Calibri", bold, false, APP_FONT)
            .unwrap();
        expected
            .register_font("Calibri", bold, false, APP_FONT)
            .unwrap();
    }
    assert_eq!(glyphs(&presentation), glyphs(&expected));
    let presentation = Presentation::open_with_options(
        &bytes,
        &OpenOptions {
            fonts: [false, true]
                .into_iter()
                .map(|bold| FontDefinition {
                    family: "Calibri".into(),
                    data: APP_FONT.to_vec(),
                    bold,
                    italic: false,
                })
                .collect(),
            ..Default::default()
        },
    )
    .unwrap();
    assert_eq!(glyphs(&presentation), glyphs(&expected));
}

#[cfg(feature = "fonts")]
#[test]
fn embedded_faces_win_over_registered_faces_and_the_source() {
    let embedded = include_bytes!("../../docx-raster/tests/assets/Carlito-Regular.ttf");
    let mut presentation = Presentation::open(&fixture(Some(embedded))).unwrap();
    let mut expected = Presentation::open_with_options(&fixture(None), &backstop()).unwrap();
    for bold in [false, true] {
        presentation
            .register_font("Calibri", bold, false, APP_FONT)
            .unwrap();
        expected
            .register_font("Calibri", bold, false, embedded)
            .unwrap();
    }
    assert_eq!(glyphs(&presentation), glyphs(&expected));
}

#[test]
fn backstop_keeps_the_fontless_error() {
    let presentation = Presentation::open_with_options(&fixture(None), &backstop()).unwrap();
    assert!(matches!(
        presentation.render_slide(0),
        Err(betteroffice_pptx::Error::Render(
            betteroffice_pptx::RenderError::NoFont
        ))
    ));
}

#[cfg(not(any(feature = "fonts-latin", feature = "fonts-intl", feature = "fonts-cjk")))]
#[test]
fn open_without_pack_features_keeps_the_fontless_error() {
    let presentation = Presentation::open(&fixture(None)).unwrap();
    assert!(matches!(
        presentation.render_slide(0),
        Err(betteroffice_pptx::Error::Render(
            betteroffice_pptx::RenderError::NoFont
        ))
    ));
}
