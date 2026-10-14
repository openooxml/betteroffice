use pptx_edit::DeckSession;
use pptx_render::{Paint, Primitive, SlideRenderer, SurfaceDisplayList};

const FONT: &[u8] = include_bytes!("../../ooxml-text/tests/fonts/LiberationSans-Regular.ttf");

const RELATIONSHIPS: &str = "http://schemas.openxmlformats.org/package/2006/relationships";
const OFFICE: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const DRAWING: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";
const PRESENTATION: &str = "http://schemas.openxmlformats.org/presentationml/2006/main";

/// The `Blue` scheme and inverted `p:clrMap` measured in the
/// `pptarena-054-original` corpus deck, whose slides PowerPoint paints on
/// `dk2` = `#17406D` with `lt1` = white title text.
fn deck(color_map: &str, layout_mapping: &str, override_mapping: &str) -> Vec<u8> {
    let parts: Vec<(String, Vec<u8>)> = vec![
        (
            "[Content_Types].xml".to_owned(),
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/><Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/><Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/><Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/><Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/></Types>"#.to_owned()
            .into_bytes(),
        ),
        (
            "_rels/.rels".to_owned(),
            format!(
                r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="{RELATIONSHIPS}"><Relationship Id="rId1" Type="{OFFICE}/officeDocument" Target="ppt/presentation.xml"/></Relationships>"#
            )
            .into_bytes(),
        ),
        (
            "ppt/presentation.xml".to_owned(),
            format!(
                r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:presentation xmlns:a="{DRAWING}" xmlns:r="{OFFICE}" xmlns:p="{PRESENTATION}"><p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst><p:sldIdLst><p:sldId id="256" r:id="rId2"/></p:sldIdLst><p:sldSz cx="12192000" cy="6858000"/><p:notesSz cx="6858000" cy="9144000"/></p:presentation>"#
            )
            .into_bytes(),
        ),
        (
            "ppt/_rels/presentation.xml.rels".to_owned(),
            format!(
                r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="{RELATIONSHIPS}"><Relationship Id="rId1" Type="{OFFICE}/slideMaster" Target="slideMasters/slideMaster1.xml"/><Relationship Id="rId2" Type="{OFFICE}/slide" Target="slides/slide1.xml"/></Relationships>"#
            )
            .into_bytes(),
        ),
        (
            "ppt/slides/slide1.xml".to_owned(),
            format!(
                r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:sld xmlns:a="{DRAWING}" xmlns:r="{OFFICE}" xmlns:p="{PRESENTATION}"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/><p:sp><p:nvSpPr><p:cNvPr id="2" name="Band"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="6096000" cy="3429000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:schemeClr val="bg1"/></a:solidFill></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US" sz="4000"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="Arial"/></a:rPr><a:t>Cinematic Studies</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld>{override_mapping}</p:sld>"#
            )
            .into_bytes(),
        ),
        (
            "ppt/slides/_rels/slide1.xml.rels".to_owned(),
            format!(
                r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="{RELATIONSHIPS}"><Relationship Id="rId1" Type="{OFFICE}/slideLayout" Target="../slideLayouts/slideLayout1.xml"/></Relationships>"#
            )
            .into_bytes(),
        ),
        (
            "ppt/slideLayouts/slideLayout1.xml".to_owned(),
            format!(
                r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:sldLayout xmlns:a="{DRAWING}" xmlns:r="{OFFICE}" xmlns:p="{PRESENTATION}" type="title"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld>{layout_mapping}</p:sldLayout>"#
            )
            .into_bytes(),
        ),
        (
            "ppt/slideLayouts/_rels/slideLayout1.xml.rels".to_owned(),
            format!(
                r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="{RELATIONSHIPS}"><Relationship Id="rId1" Type="{OFFICE}/slideMaster" Target="../slideMasters/slideMaster1.xml"/></Relationships>"#
            )
            .into_bytes(),
        ),
        (
            "ppt/slideMasters/slideMaster1.xml".to_owned(),
            format!(
                r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:sldMaster xmlns:a="{DRAWING}" xmlns:r="{OFFICE}" xmlns:p="{PRESENTATION}"><p:cSld><p:bg><p:bgRef idx="1001"><a:schemeClr val="bg2"/></p:bgRef></p:bg><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld>{color_map}<p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst></p:sldMaster>"#
            )
            .into_bytes(),
        ),
        (
            "ppt/slideMasters/_rels/slideMaster1.xml.rels".to_owned(),
            format!(
                r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="{RELATIONSHIPS}"><Relationship Id="rId1" Type="{OFFICE}/slideLayout" Target="../slideLayouts/slideLayout1.xml"/><Relationship Id="rId2" Type="{OFFICE}/theme" Target="../theme/theme1.xml"/></Relationships>"#
            )
            .into_bytes(),
        ),
        (
            "ppt/theme/theme1.xml".to_owned(),
            format!(
                r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><a:theme xmlns:a="{DRAWING}" name="Blue"><a:themeElements><a:clrScheme name="Blue"><a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="17406D"/></a:dk2><a:lt2><a:srgbClr val="DBEFF9"/></a:lt2><a:accent1><a:srgbClr val="0F6FC6"/></a:accent1><a:accent2><a:srgbClr val="009DD9"/></a:accent2><a:accent3><a:srgbClr val="0BD0D9"/></a:accent3><a:accent4><a:srgbClr val="10CF9B"/></a:accent4><a:accent5><a:srgbClr val="7CCA62"/></a:accent5><a:accent6><a:srgbClr val="A5C249"/></a:accent6><a:hlink><a:srgbClr val="F49100"/></a:hlink><a:folHlink><a:srgbClr val="85DFD0"/></a:folHlink></a:clrScheme><a:fontScheme name="Blue"><a:majorFont><a:latin typeface="Corbel"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont><a:minorFont><a:latin typeface="Corbel"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont></a:fontScheme><a:fmtScheme name="Blue"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst><a:lnStyleLst><a:ln><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst></a:fmtScheme></a:themeElements></a:theme>"#
            )
            .into_bytes(),
        ),
    ];
    ooxml_opc::rezip_parts(&parts).unwrap()
}

fn render(bytes: &[u8]) -> SurfaceDisplayList {
    let session = DeckSession::open(bytes, 900).unwrap();
    let mut renderer = SlideRenderer::new();
    renderer.register_font("Arial", false, false, FONT).unwrap();
    renderer
        .layout_slide(session.package(), &session.snapshot().unwrap(), 0)
        .unwrap()
        .display_list
}

fn band_fill(list: &SurfaceDisplayList) -> Option<&Paint> {
    list.primitives
        .iter()
        .find_map(|primitive| match primitive {
            Primitive::Shape { name, fill, .. } if name == "Band" => fill.as_ref(),
            _ => None,
        })
}

fn title_color(list: &SurfaceDisplayList) -> String {
    list.primitives
        .iter()
        .find_map(|primitive| match primitive {
            Primitive::TextBox { lines, .. } => lines
                .iter()
                .flat_map(|line| &line.runs)
                .find(|run| run.text.contains("Cinematic"))
                .map(|run| run.color.clone()),
            _ => None,
        })
        .unwrap()
}

const INVERTED: &str = r#"<p:clrMap bg1="dk1" tx1="lt1" bg2="dk2" tx2="lt2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>"#;
const IDENTITY: &str = r#"<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>"#;
const MASTER_MAPPING: &str = "<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>";

#[test]
fn an_inverted_master_colour_map_swaps_background_and_text_slots() {
    let list = render(&deck(INVERTED, MASTER_MAPPING, MASTER_MAPPING));
    assert_eq!(
        list.background,
        Some(Paint::Solid {
            color: "#17406D".to_owned()
        })
    );
    assert_eq!(
        band_fill(&list),
        Some(&Paint::Solid {
            color: "#000000".to_owned()
        })
    );
    assert_eq!(title_color(&list), "#FFFFFF");
}

#[test]
fn an_identity_master_colour_map_leaves_every_slot_alone() {
    let list = render(&deck(IDENTITY, MASTER_MAPPING, MASTER_MAPPING));
    assert_eq!(
        list.background,
        Some(Paint::Solid {
            color: "#DBEFF9".to_owned()
        })
    );
    assert_eq!(
        band_fill(&list),
        Some(&Paint::Solid {
            color: "#FFFFFF".to_owned()
        })
    );
    assert_eq!(title_color(&list), "#000000");
}

#[test]
fn a_slide_override_mapping_wins_over_the_master() {
    let list = render(&deck(
        INVERTED,
        MASTER_MAPPING,
        r#"<p:clrMapOvr><a:overrideClrMapping bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/></p:clrMapOvr>"#,
    ));
    assert_eq!(
        list.background,
        Some(Paint::Solid {
            color: "#DBEFF9".to_owned()
        })
    );
    assert_eq!(title_color(&list), "#000000");
}

const INVERTED_OVERRIDE: &str = r#"<p:clrMapOvr><a:overrideClrMapping bg1="dk1" tx1="lt1" bg2="dk2" tx2="lt2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/></p:clrMapOvr>"#;

/// `pptarena-031-original`'s shape: an identity master, the inversion declared
/// on the layout, and a slide that inherits it.
#[test]
fn a_layout_override_reaches_a_slide_that_inherits_it() {
    let list = render(&deck(IDENTITY, INVERTED_OVERRIDE, MASTER_MAPPING));
    assert_eq!(
        list.background,
        Some(Paint::Solid {
            color: "#17406D".to_owned()
        })
    );
    assert_eq!(
        band_fill(&list),
        Some(&Paint::Solid {
            color: "#000000".to_owned()
        })
    );
    assert_eq!(title_color(&list), "#FFFFFF");
}

/// The save projection resolves colours too, so an untouched mapped deck must
/// still re-serialize byte-identically.
#[test]
fn a_mapped_deck_saves_byte_identically_when_untouched() {
    let bytes = deck(IDENTITY, INVERTED_OVERRIDE, MASTER_MAPPING);
    let session = DeckSession::open(&bytes, 900).unwrap();
    let saved = session.save().unwrap();
    let part = |zip: &[u8], name: &str| {
        ooxml_opc::unzip_parts(zip)
            .unwrap()
            .into_iter()
            .find(|(path, _)| path == name)
            .map(|(_, body)| body)
            .unwrap()
    };
    for name in [
        "ppt/slides/slide1.xml",
        "ppt/slideLayouts/slideLayout1.xml",
        "ppt/slideMasters/slideMaster1.xml",
    ] {
        assert_eq!(
            part(&saved, name),
            part(&bytes, name),
            "{name} changed on save"
        );
    }
}
