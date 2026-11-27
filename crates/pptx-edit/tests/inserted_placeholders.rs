use std::collections::BTreeMap;

use pptx_edit::{DeckSession, EditCtx, InheritedGeometry, SlideSnapshot};

const REL: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const NS: &str = r#"xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships""#;
const MASTER: &str = "ppt/slideMasters/slideMaster1.xml";
const SOURCE_SLIDE: &str = "ppt/slides/slide1.xml";
const CONTENT_LAYOUT: usize = 2;
const TITLE_PROPERTIES: &str = r#"<a:xfrm flipH="1" rot="600000"><a:off x="800000" y="450000"/><a:ext cx="10000000" cy="1100000"/></a:xfrm>"#;
const TITLE_GEOMETRY: &str =
    r#"<a:prstGeom prst="roundRect"><a:avLst><a:gd fmla="val 25000" name="adj"/></a:avLst></a:prstGeom>"#;
const FOOTERS: &[&str] = &[
    r#"type="dt" idx="10""#,
    r#"type="ftr" idx="11""#,
    r#"type="sldNum" idx="12""#,
    r#"type="hdr" idx="13""#,
];
const TITLE: &str = r#"type="title""#;
const BODY: &str = r#"type="body" idx="1""#;
const PICTURE: &str = r#"type="pic" idx="2""#;
const LAYOUTS: &[(&str, &str, &[&str])] = &[
    (
        "Title Slide",
        "title",
        &[r#"type="ctrTitle""#, r#"type="subTitle" idx="1""#],
    ),
    ("Title and Content", "obj", &[TITLE, r#"idx="1""#]),
    (
        "Two Content",
        "twoObj",
        &[TITLE, r#"idx="1""#, r#"type="obj" idx="2""#],
    ),
    ("Picture with Caption", "picTx", &[TITLE, BODY, PICTURE]),
    (
        "Title and Vertical Text",
        "vertTx",
        &[TITLE, r#"type="body" idx="1" orient="vert""#],
    ),
    ("Blank", "blank", &[]),
];

fn layout_path(number: usize) -> String {
    format!("ppt/slideLayouts/slideLayout{number}.xml")
}

fn relationship(id: usize, kind: &str, target: &str) -> String {
    format!(r#"<Relationship Id="rId{id}" Type="{REL}/{kind}" Target="{target}"/>"#)
}

fn relationships(entries: &str) -> String {
    format!(r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">{entries}</Relationships>"#)
}

fn common_slide(name: &str, shapes: &str) -> String {
    format!(r#"<p:cSld name="{name}"><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>{shapes}</p:spTree></p:cSld>"#)
}

fn placeholder(id: usize, attributes: &str, properties: &str) -> String {
    let properties = if properties.is_empty() {
        "<p:spPr/>".to_owned()
    } else {
        format!("<p:spPr>{properties}</p:spPr>")
    };
    format!(r#"<p:sp><p:nvSpPr><p:cNvPr id="{id}" name="Placeholder {id}"/><p:cNvSpPr/><p:nvPr><p:ph {attributes}/></p:nvPr></p:nvSpPr>{properties}<p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody></p:sp>"#)
}

fn deck() -> Vec<u8> {
    let mut parts = BTreeMap::new();
    let mut overrides = String::new();
    for (path, kind) in [
        ("ppt/presentation.xml", "presentation.main"),
        (MASTER, "slideMaster"),
        (SOURCE_SLIDE, "slide"),
    ] {
        overrides += &format!(r#"<Override PartName="/{path}" ContentType="application/vnd.openxmlformats-officedocument.presentationml.{kind}+xml"/>"#);
    }
    let mut layout_ids = String::new();
    let mut master_rels = String::new();
    for (index, &(name, kind, placeholders)) in LAYOUTS.iter().enumerate() {
        let number = index + 1;
        let path = layout_path(number);
        overrides += &format!(r#"<Override PartName="/{path}" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>"#);
        let id = 2_147_483_648 + number;
        layout_ids += &format!(r#"<p:sldLayoutId id="{id}" r:id="rId{number}"/>"#);
        master_rels += &relationship(
            number,
            "slideLayout",
            &format!("../slideLayouts/slideLayout{number}.xml"),
        );
        let shapes = placeholders
            .iter()
            .chain(FOOTERS)
            .enumerate()
            .map(|(i, attributes)| {
                let properties = if i == 0 && !placeholders.is_empty() {
                    TITLE_PROPERTIES.to_owned()
                } else {
                    let y = 1_750_000 + i * 450_000;
                    format!(r#"<a:xfrm><a:off x="900000" y="{y}"/><a:ext cx="4100000" cy="1500000"/></a:xfrm>"#)
                };
                placeholder(i + 2, attributes, &properties)
            })
            .collect::<String>();
        let common = common_slide(name, &shapes);
        parts.insert(
            path,
            format!(r#"<p:sldLayout {NS} type="{kind}">{common}</p:sldLayout>"#),
        );
        parts.insert(
            format!("ppt/slideLayouts/_rels/slideLayout{number}.xml.rels"),
            relationships(&relationship(1, "slideMaster", "../slideMasters/slideMaster1.xml")),
        );
    }
    parts.insert("[Content_Types].xml".to_owned(), format!(r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>{overrides}</Types>"#));
    parts.insert(
        "_rels/.rels".to_owned(),
        relationships(&relationship(1, "officeDocument", "ppt/presentation.xml")),
    );
    parts.insert("ppt/presentation.xml".to_owned(), format!(r#"<p:presentation {NS}><p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst><p:sldIdLst><p:sldId id="256" r:id="rId2"/></p:sldIdLst><p:sldSz cx="12192000" cy="6858000"/><p:notesSz cx="6858000" cy="9144000"/></p:presentation>"#));
    let presentation_rels = relationship(1, "slideMaster", "slideMasters/slideMaster1.xml")
        + &relationship(2, "slide", "slides/slide1.xml");
    parts.insert(
        "ppt/_rels/presentation.xml.rels".to_owned(),
        relationships(&presentation_rels),
    );
    let master = common_slide("Master", &placeholder(2, TITLE, TITLE_GEOMETRY));
    parts.insert(
        MASTER.to_owned(),
        format!(r#"<p:sldMaster {NS}>{master}<p:sldLayoutIdLst>{layout_ids}</p:sldLayoutIdLst></p:sldMaster>"#),
    );
    parts.insert(
        "ppt/slideMasters/_rels/slideMaster1.xml.rels".to_owned(),
        relationships(&master_rels),
    );
    let slide = common_slide(
        "Existing",
        &(placeholder(2, TITLE, "") + &placeholder(3, BODY, "")),
    );
    parts.insert(
        SOURCE_SLIDE.to_owned(),
        format!(r#"<p:sld {NS}>{slide}</p:sld>"#),
    );
    parts.insert(
        "ppt/slides/_rels/slide1.xml.rels".to_owned(),
        relationships(&relationship(1, "slideLayout", "../slideLayouts/slideLayout2.xml")),
    );
    let parts: Vec<_> = parts.into_iter().collect();
    let bytes = ooxml_opc::rezip_parts_borrowed(&parts).unwrap();
    ooxml_opc::rezip_parts_preserving(&parts, &bytes).unwrap()
}

fn inserted_content() -> (DeckSession, SlideSnapshot) {
    let session = DeckSession::open(&deck(), 7).unwrap();
    session
        .insert_slide_with_placeholders(
            &EditCtx::local("test"),
            1,
            Some(&layout_path(CONTENT_LAYOUT)),
        )
        .unwrap();
    let slide = session.snapshot().unwrap().slides.remove(1);
    (session, slide)
}

fn reopened_content(bytes: &[u8]) -> (SlideSnapshot, String) {
    let session = DeckSession::open(bytes, 8).unwrap();
    let slide = session.snapshot().unwrap().slides.remove(1);
    let bytes = session
        .package()
        .part_bytes(slide.source_part_path.as_deref().unwrap())
        .unwrap();
    let xml = std::str::from_utf8(bytes).unwrap().to_owned();
    (slide, xml)
}

#[test]
fn insertion_copies_only_text_placeholders_across_standard_layouts() {
    let session = DeckSession::open(&deck(), 7).unwrap();
    let title = (Some("title"), None, None);
    let default_content = (None, Some(1), None);
    let body = (Some("body"), Some(1), None);
    let expected: &[&[_]] = &[
        &[(Some("ctrTitle"), None, None), (Some("subTitle"), Some(1), None)],
        &[title, default_content],
        &[title, default_content, (Some("obj"), Some(2), None)],
        &[title, body],
        &[title, (Some("body"), Some(1), Some("vert"))],
        &[],
    ];
    for (index, expected) in expected.iter().enumerate() {
        let number = index + 1;
        session
            .insert_slide_with_placeholders(
                &EditCtx::local("test"),
                number as u32,
                Some(&layout_path(number)),
            )
            .unwrap();
        let snapshot = session.snapshot().unwrap();
        let shapes = &snapshot.slides[number].shapes;
        let actual: Vec<_> = shapes
            .iter()
            .map(|shape| {
                let placeholder = shape.placeholder.as_ref().unwrap();
                assert_eq!(shape.text_stories.len(), 1);
                assert_eq!(shape.text_stories[0].paragraphs.len(), 1);
                assert_eq!(shape.text_stories[0].plain_text(), "");
                (
                    placeholder.placeholder_type.as_deref(),
                    placeholder.index,
                    placeholder.orientation.as_deref(),
                )
            })
            .collect();
        assert_eq!(actual.as_slice(), *expected, "{}", LAYOUTS[index].0);
    }
}

#[test]
fn untouched_inserted_placeholders_save_empty_properties_and_reopen_inherited() {
    let (session, _) = inserted_content();
    let (slide, xml) = reopened_content(&session.save().unwrap());
    assert_eq!(xml.matches("<p:spPr/>").count(), 2, "{xml}");
    assert_eq!(xml.matches("<a:xfrm").count(), 1, "{xml}");
    assert_eq!(xml.matches("<a:bodyPr/><a:lstStyle/><a:p/>").count(), 2, "{xml}");
    assert!(!xml.contains("<a:prstGeom"), "{xml}");
    assert_eq!(
        slide.shapes[0].inherited,
        Some(InheritedGeometry {
            x: 800_000,
            y: 450_000,
            width: 10_000_000,
            height: 1_100_000,
            rotation_deg: 10.0,
            flip_h: true,
            flip_v: false,
        }),
    );
}

#[test]
fn untouched_inserted_placeholder_follows_a_repositioned_layout() {
    let (session, _) = inserted_content();
    let saved = session.save().unwrap();
    let mut parts = ooxml_opc::unzip_parts(&saved).unwrap();
    let path = layout_path(CONTENT_LAYOUT);
    let (_, bytes) = parts.iter_mut().find(|(name, _)| name == &path).unwrap();
    *bytes = std::str::from_utf8(bytes)
        .unwrap()
        .replace("x=\"800000\"", "x=\"2100000\"")
        .into_bytes();
    let (slide, _) = reopened_content(&ooxml_opc::rezip_parts(&parts).unwrap());
    let title = &slide.shapes[0];
    assert_eq!(title.inherited.map_or(title.x, |rect| rect.x), 2_100_000);
}

#[test]
fn a_moved_inserted_placeholder_saves_an_explicit_transform() {
    let (session, slide) = inserted_content();
    session
        .move_shape(
            &EditCtx::local("test"),
            &slide.id,
            &slide.shapes[0].id,
            2_100_000,
            450_000,
        )
        .unwrap();
    let (slide, xml) = reopened_content(&session.save().unwrap());
    assert!(xml.contains(&TITLE_PROPERTIES.replace("800000", "2100000")), "{xml}");
    assert_eq!(slide.shapes[0].x, 2_100_000);
    assert!(slide.shapes[0].inherited.is_none());
}

#[test]
fn adjusting_an_inserted_placeholder_saves_explicit_geometry() {
    let (session, slide) = inserted_content();
    session
        .set_shape_adjust(
            &EditCtx::local("test"),
            &slide.id,
            &slide.shapes[0].id,
            &BTreeMap::from([("adj".to_owned(), 0.4)]),
        )
        .unwrap();
    let (_, xml) = reopened_content(&session.save().unwrap());
    assert!(xml.contains(TITLE_PROPERTIES), "{xml}");
    assert!(xml.contains(&TITLE_GEOMETRY.replace("25000", "40000")), "{xml}");
}

#[test]
fn existing_empty_title_and_body_placeholders_save_byte_identically() {
    let bytes = deck();
    let session = DeckSession::open(&bytes, 7).unwrap();
    let snapshot = session.snapshot().unwrap();
    let shapes = &snapshot.slides[0].shapes;
    assert_eq!(shapes.len(), 2);
    assert!(shapes.iter().all(|shape| {
        shape.inherited.is_some() && shape.text_stories[0].plain_text().is_empty()
    }));
    assert_eq!(session.save().unwrap(), bytes);
}
