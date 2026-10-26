use pptx_edit::DeckSession;
use pptx_render::{Primitive, Shadow, SlideRenderer};

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

const THEME: &str = r#"<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Fixture">
<a:themeElements><a:fmtScheme name="Fixture"><a:effectStyleLst>
<a:effectStyle><a:effectLst/></a:effectStyle>
<a:effectStyle><a:effectLst><a:outerShdw blurRad="76200" dist="38100" dir="5400000"><a:srgbClr val="000000"><a:alpha val="38000"/></a:srgbClr></a:outerShdw></a:effectLst></a:effectStyle>
<a:effectStyle><a:effectLst><a:outerShdw blurRad="38100" dist="76200" dir="0"><a:schemeClr val="phClr"><a:alpha val="50000"/></a:schemeClr></a:outerShdw></a:effectLst></a:effectStyle>
</a:effectStyleLst></a:fmtScheme></a:themeElements></a:theme>"#;

fn shape(kind: &str, index: u32, effects: &str) -> String {
    let nv = match kind {
        "sp" => "nvSpPr",
        "cxnSp" => "nvCxnSpPr",
        "pic" => "nvPicPr",
        _ => unreachable!(),
    };
    format!(
        r#"<p:{kind}><p:{nv}><p:cNvPr id="2" name="Styled"/><p:nvPr/></p:{nv}>
<p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm><a:prstGeom prst="rect"/><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill><a:ln w="9525"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln>{effects}</p:spPr>
<p:style><a:effectRef idx="{index}"><a:srgbClr val="2468AC"/></a:effectRef></p:style></p:{kind}>"#
    )
}

fn fixture(shape: &str, layout: &str) -> Vec<u8> {
    fixture_with_master(shape, layout, MASTER)
}

fn fixture_with_master(shape: &str, layout: &str, master: &str) -> Vec<u8> {
    let slide = format!(
        r#"<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:spTree>{shape}</p:spTree></p:cSld></p:sld>"#
    );
    let parts: Vec<(String, Vec<u8>)> = [
        ("[Content_Types].xml", CONTENT_TYPES),
        ("_rels/.rels", ROOT_RELS),
        ("ppt/presentation.xml", PRESENTATION),
        ("ppt/_rels/presentation.xml.rels", PRESENTATION_RELS),
        ("ppt/slides/slide1.xml", slide.as_str()),
        ("ppt/slides/_rels/slide1.xml.rels", SLIDE_RELS),
        ("ppt/slideLayouts/slideLayout1.xml", layout),
        ("ppt/slideLayouts/_rels/slideLayout1.xml.rels", LAYOUT_RELS),
        ("ppt/slideMasters/slideMaster1.xml", master),
        ("ppt/slideMasters/_rels/slideMaster1.xml.rels", MASTER_RELS),
        ("ppt/theme/theme1.xml", THEME),
    ]
    .into_iter()
    .map(|(path, xml)| (path.to_owned(), xml.as_bytes().to_vec()))
    .collect();
    ooxml_opc::rezip_parts(&parts).unwrap()
}

fn rendered_shadow(bytes: &[u8], prefix: &str) -> Option<Shadow> {
    let session = DeckSession::open(bytes, 37).unwrap();
    let list = SlideRenderer::new()
        .layout_slide(session.package(), &session.snapshot().unwrap(), 0)
        .unwrap()
        .display_list;
    list.primitives
        .into_iter()
        .find_map(|primitive| match primitive {
            Primitive::Shape {
                shape_id: Some(id),
                shadow,
                ..
            }
            | Primitive::Image {
                shape_id: Some(id),
                shadow,
                ..
            } if id.starts_with(prefix) => Some(shadow),
            _ => None,
        })
        .expect("styled object is drawn")
}

#[test]
fn theme_effect_refs_resolve_shadows_for_shapes_connectors_and_pictures() {
    for kind in ["sp", "cxnSp", "pic"] {
        let bytes = fixture(&shape(kind, 2, ""), LAYOUT);
        let shadow = rendered_shadow(&bytes, "slide:").unwrap();
        assert_eq!(shadow.color, "#00000061", "{kind}");
        assert!((shadow.blur - 8.0).abs() < 0.01, "{kind}");
        assert!(shadow.dx.abs() < 0.01, "{kind}");
        assert!((shadow.dy - 4.0).abs() < 0.01, "{kind}");
    }
}

#[test]
fn theme_effect_refs_respect_empty_lists_and_zero_or_missing_indices() {
    for kind in ["sp", "cxnSp", "pic"] {
        for (index, effects) in [(2, "<a:effectLst/>"), (0, ""), (1, ""), (4, "")] {
            let bytes = fixture(&shape(kind, index, effects), LAYOUT);
            assert_eq!(
                rendered_shadow(&bytes, "slide:"),
                None,
                "{kind}, {index}, {effects}"
            );
        }
    }
}

#[test]
fn theme_effect_refs_substitute_the_reference_color_for_ph_clr() {
    for kind in ["sp", "cxnSp", "pic"] {
        let bytes = fixture(&shape(kind, 3, ""), LAYOUT);
        let shadow = rendered_shadow(&bytes, "slide:").unwrap();
        assert_eq!(shadow.color, "#2468AC80", "{kind}");
        assert!((shadow.blur - 4.0).abs() < 0.01, "{kind}");
        assert!((shadow.dx - 8.0).abs() < 0.01, "{kind}");
        assert!(shadow.dy.abs() < 0.01, "{kind}");
    }
}

#[test]
fn theme_effect_refs_resolve_on_layout_decorations() {
    for kind in ["sp", "cxnSp", "pic"] {
        let layout = LAYOUT.replace("</p:spTree>", &format!("{}</p:spTree>", shape(kind, 2, "")));
        let bytes = fixture("", &layout);
        assert_eq!(
            rendered_shadow(&bytes, "layout:").unwrap().color,
            "#00000061",
            "{kind}"
        );
    }
}

#[test]
fn theme_effect_refs_leave_source_parts_unchanged_on_save() {
    let bytes = fixture(&shape("sp", 2, ""), LAYOUT);
    let session = DeckSession::open(&bytes, 38).unwrap();
    let saved = session.save().unwrap();
    let original_parts = ooxml_opc::unzip_parts(&bytes).unwrap();
    let saved_parts = ooxml_opc::unzip_parts(&saved).unwrap();
    for path in ["ppt/slides/slide1.xml", "ppt/theme/theme1.xml"] {
        let original = original_parts
            .iter()
            .find(|(name, _)| name == path)
            .unwrap();
        let saved = saved_parts.iter().find(|(name, _)| name == path).unwrap();
        assert_eq!(original.1, saved.1, "{path}");
    }
}

#[test]
fn theme_effect_refs_follow_explicit_and_placeholder_effects() {
    let explicit =
        r#"<a:effectLst><a:outerShdw><a:srgbClr val="123456"/></a:outerShdw></a:effectLst>"#;
    let placeholder = |index, effects| {
        shape("sp", index, effects).replace(
            "<p:nvPr/>",
            r#"<p:nvPr><p:ph type="body" idx="1"/></p:nvPr>"#,
        )
    };
    for (own, layout_effects, master_effects, expected) in [
        (explicit, "<a:effectLst/>", "", Some("#123456")),
        ("<a:effectLst/>", explicit, "", None),
        ("", explicit, "<a:effectLst/>", Some("#123456")),
        ("", "<a:effectLst/>", explicit, None),
        ("", "", explicit, Some("#123456")),
        ("", "", "<a:effectLst/>", None),
        ("", "", "", Some("#00000061")),
    ] {
        let layout = LAYOUT.replace(
            "</p:spTree>",
            &format!("{}</p:spTree>", placeholder(3, layout_effects)),
        );
        let master = MASTER.replace(
            "</p:spTree>",
            &format!("{}</p:spTree>", placeholder(3, master_effects)),
        );
        let bytes = fixture_with_master(&placeholder(2, own), &layout, &master);
        let shadow = rendered_shadow(&bytes, "slide:");
        assert_eq!(
            shadow.as_ref().map(|shadow| shadow.color.as_str()),
            expected,
            "{own}, {layout_effects}, {master_effects}"
        );
    }
}
