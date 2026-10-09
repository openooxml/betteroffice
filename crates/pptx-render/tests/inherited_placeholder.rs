//! A placeholder draws at the same rect and orientation live, after its first
//! geometry edit, and once that edit is saved and the deck reopened.

use pptx_edit::{DeckSession, EditCtx, ShapeSnapshot};
use pptx_render::{Primitive, SlideRenderer, Transform};

const FONT: &[u8] = include_bytes!("../../ooxml-text/tests/fonts/LiberationSans-Regular.ttf");

const NS: &str = r#"xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main""#;

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

const PRESENTATION_RELS: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>
</Relationships>"#;

const SLIDE_RELS: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
</Relationships>"#;

const LAYOUT_RELS: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/>
</Relationships>"#;

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

/// A drawn shape: its rect in EMU and the transform it is drawn with.
type Drawn = ((i64, i64, i64, i64), Transform);

fn sp(id: u32, name: &str, placeholder: &str, properties: &str) -> String {
    format!(
        r#"<p:sp><p:nvSpPr><p:cNvPr id="{id}" name="{name}"/><p:cNvSpPr/><p:nvPr>{placeholder}</p:nvPr></p:nvSpPr>{properties}<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr><a:latin typeface="Arial"/></a:rPr><a:t>{name}</a:t></a:r></a:p></p:txBody></p:sp>"#
    )
}

fn filled(transform: &str) -> String {
    format!(r#"<p:spPr>{transform}<a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></p:spPr>"#)
}

fn tree(shapes: &str) -> String {
    format!(
        r#"<p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>{shapes}</p:spTree></p:cSld>"#
    )
}

fn deck(slide: &str, layout: &str, master: &str) -> Vec<u8> {
    let presentation = format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation {NS}><p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst><p:sldIdLst><p:sldId id="256" r:id="rId2"/></p:sldIdLst><p:sldSz cx="12192000" cy="6858000"/><p:notesSz cx="6858000" cy="9144000"/></p:presentation>"#
    );
    let slide = format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld {NS}>{}</p:sld>"#,
        tree(slide)
    );
    let layout = format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldLayout {NS} type="obj">{}</p:sldLayout>"#,
        tree(layout)
    );
    let master = format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldMaster {NS}>{}<p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst></p:sldMaster>"#,
        tree(master)
    );
    let parts: Vec<(String, Vec<u8>)> = [
        ("[Content_Types].xml", CONTENT_TYPES.to_owned()),
        ("_rels/.rels", ROOT_RELS.to_owned()),
        ("ppt/presentation.xml", presentation),
        (
            "ppt/_rels/presentation.xml.rels",
            PRESENTATION_RELS.to_owned(),
        ),
        ("ppt/slides/slide1.xml", slide),
        ("ppt/slides/_rels/slide1.xml.rels", SLIDE_RELS.to_owned()),
        ("ppt/slideLayouts/slideLayout1.xml", layout),
        (
            "ppt/slideLayouts/_rels/slideLayout1.xml.rels",
            LAYOUT_RELS.to_owned(),
        ),
        ("ppt/slideMasters/slideMaster1.xml", master),
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

fn context() -> EditCtx {
    EditCtx::local("inherited")
}

fn emu(px: f32) -> i64 {
    (f64::from(px) * 9525.0).round() as i64
}

fn drawn(session: &DeckSession, shape_id: &str) -> Drawn {
    let mut renderer = SlideRenderer::new();
    renderer.register_font("Arial", false, false, FONT).unwrap();
    let rendered = renderer
        .layout_slide(session.package(), &session.snapshot().unwrap(), 0)
        .unwrap();
    rendered
        .display_list
        .primitives
        .iter()
        .find_map(|primitive| match primitive {
            Primitive::Shape {
                shape_id: Some(id),
                x,
                y,
                w,
                h,
                transform,
                ..
            } if id == shape_id => Some(((emu(*x), emu(*y), emu(*w), emu(*h)), *transform)),
            _ => None,
        })
        .unwrap_or_else(|| panic!("{shape_id} was not drawn"))
}

fn shape(session: &DeckSession, index: usize) -> ShapeSnapshot {
    session.snapshot().unwrap().slides[0].shapes[index].clone()
}

fn slide_id(session: &DeckSession) -> String {
    session.snapshot().unwrap().slides[0].id.clone()
}

fn reopened(session: &DeckSession) -> DeckSession {
    DeckSession::open(&session.save().unwrap(), 2).unwrap()
}

fn turned(rotation_deg: f32, flip_h: bool) -> Transform {
    Transform {
        rotation_deg,
        flip_h,
        flip_v: false,
    }
}

/// A body placeholder over a turned and flipped layout body at
/// `952500,1905000 7620000x2857500`.
fn body_deck(properties: &str) -> Vec<u8> {
    deck(
        &sp(3, "Body", r#"<p:ph type="body"/>"#, properties),
        &sp(
            3,
            "Body Placeholder",
            r#"<p:ph type="body"/>"#,
            &filled(
                r#"<a:xfrm rot="1800000" flipH="1"><a:off x="952500" y="1905000"/><a:ext cx="7620000" cy="2857500"/></a:xfrm>"#,
            ),
        ),
        "",
    )
}

const LAYOUT_BODY: (i64, i64, i64, i64) = (952_500, 1_905_000, 7_620_000, 2_857_500);

#[test]
fn a_move_to_the_slide_corner_draws_there_live_and_reopened() {
    let bytes = deck(
        &sp(2, "Title", r#"<p:ph type="title"/>"#, "<p:spPr/>"),
        &sp(
            2,
            "Title Placeholder",
            r#"<p:ph type="title"/>"#,
            &filled(
                r#"<a:xfrm><a:off x="762000" y="457200"/><a:ext cx="10668000" cy="1143000"/></a:xfrm>"#,
            ),
        ),
        "",
    );
    let session = DeckSession::open(&bytes, 1).unwrap();
    let title = shape(&session, 0);
    assert_eq!(
        drawn(&session, &title.id),
        (
            (762_000, 457_200, 10_668_000, 1_143_000),
            Transform::default()
        )
    );

    session
        .move_shape(&context(), &slide_id(&session), &title.id, 0, 0)
        .unwrap();

    let live = drawn(&session, &title.id);
    assert_eq!(live, ((0, 0, 10_668_000, 1_143_000), Transform::default()));
    assert_eq!(drawn(&reopened(&session), &title.id), live);
}

#[test]
fn an_offset_only_transform_is_not_inherited_and_draws_with_its_own_orientation() {
    let session = DeckSession::open(
        &body_deck(r#"<p:spPr><a:xfrm><a:off x="123825" y="657225"/></a:xfrm></p:spPr>"#),
        1,
    )
    .unwrap();
    let body = shape(&session, 0);

    assert_eq!(body.inherited, None);
    assert_eq!(
        drawn(&session, &body.id),
        (LAYOUT_BODY, Transform::default())
    );
}

#[test]
fn an_api_move_of_an_offset_only_transform_moves_the_rect_it_draws_at() {
    let session = DeckSession::open(
        &body_deck(r#"<p:spPr><a:xfrm><a:off x="123825" y="657225"/></a:xfrm></p:spPr>"#),
        1,
    )
    .unwrap();
    let body = shape(&session, 0);

    session
        .move_shape(
            &context(),
            &slide_id(&session),
            &body.id,
            1_905_000,
            2_857_500,
        )
        .unwrap();

    let live = drawn(&session, &body.id);
    assert_eq!(
        live,
        (
            (1_905_000, 2_857_500, LAYOUT_BODY.2, LAYOUT_BODY.3),
            Transform::default()
        )
    );
    assert_eq!(drawn(&reopened(&session), &body.id), live);
}

#[test]
fn a_rotation_only_transform_keeps_its_own_rotation_through_a_resize() {
    let session =
        DeckSession::open(&body_deck(r#"<p:spPr><a:xfrm rot="1200000"/></p:spPr>"#), 1).unwrap();
    let body = shape(&session, 0);

    assert_eq!(body.inherited, None);
    assert_eq!(
        drawn(&session, &body.id),
        (LAYOUT_BODY, turned(20.0, false))
    );

    session
        .resize_shape(
            &context(),
            &slide_id(&session),
            &body.id,
            7_620_000,
            2_857_500,
        )
        .unwrap();

    let live = drawn(&session, &body.id);
    assert_eq!(live, ((0, 0, 7_620_000, 2_857_500), turned(20.0, false)));
    assert_eq!(drawn(&reopened(&session), &body.id), live);
}

/// A slide number and a date whose indices name the layout's content
/// placeholders, while the master's own slide number and date sit elsewhere.
fn footer_deck() -> Vec<u8> {
    deck(
        &[
            sp(
                4,
                "Number",
                r#"<p:ph type="sldNum" sz="quarter" idx="12"/>"#,
                "<p:spPr/>",
            ),
            sp(
                5,
                "Date",
                r#"<p:ph type="dt" sz="half" idx="10"/>"#,
                "<p:spPr/>",
            ),
        ]
        .concat(),
        &[
            sp(
                6,
                "Content 12",
                r#"<p:ph sz="quarter" idx="12"/>"#,
                &filled(
                    r#"<a:xfrm><a:off x="952500" y="952500"/><a:ext cx="3810000" cy="1905000"/></a:xfrm>"#,
                ),
            ),
            sp(
                7,
                "Content 10",
                r#"<p:ph sz="half" idx="10"/>"#,
                &filled(
                    r#"<a:xfrm><a:off x="952500" y="3810000"/><a:ext cx="3810000" cy="1905000"/></a:xfrm>"#,
                ),
            ),
        ]
        .concat(),
        &[
            sp(
                8,
                "Slide Number",
                r#"<p:ph type="sldNum" sz="quarter" idx="4"/>"#,
                &filled(
                    r#"<a:xfrm><a:off x="8610600" y="6286500"/><a:ext cx="2743200" cy="381000"/></a:xfrm>"#,
                ),
            ),
            sp(
                9,
                "Date",
                r#"<p:ph type="dt" sz="half" idx="2"/>"#,
                &filled(
                    r#"<a:xfrm><a:off x="838200" y="6286500"/><a:ext cx="2743200" cy="381000"/></a:xfrm>"#,
                ),
            ),
        ]
        .concat(),
    )
}

const MASTER_NUMBER: (i64, i64, i64, i64) = (8_610_600, 6_286_500, 2_743_200, 381_000);
const MASTER_DATE: (i64, i64, i64, i64) = (838_200, 6_286_500, 2_743_200, 381_000);

#[test]
fn a_slide_number_and_date_inherit_the_rect_they_are_drawn_at() {
    let session = DeckSession::open(&footer_deck(), 1).unwrap();
    for (index, expected) in [(0, MASTER_NUMBER), (1, MASTER_DATE)] {
        let footer = shape(&session, index);
        let inherited = footer
            .inherited
            .expect("the master placeholder has geometry");

        assert_eq!(
            (inherited.x, inherited.y, inherited.width, inherited.height),
            expected
        );
        assert_eq!(
            drawn(&session, &footer.id),
            (expected, Transform::default())
        );
    }
}

#[test]
fn a_first_nudge_keeps_the_drawn_size_of_a_slide_number_and_date() {
    let session = DeckSession::open(&footer_deck(), 1).unwrap();
    let slide_id = slide_id(&session);
    for (index, (x, y, width, height)) in [(0, MASTER_NUMBER), (1, MASTER_DATE)] {
        let footer = shape(&session, index);
        session
            .move_shape(&context(), &slide_id, &footer.id, x + 95_250, y)
            .unwrap();

        let live = drawn(&session, &footer.id);
        assert_eq!(live, ((x + 95_250, y, width, height), Transform::default()));
        assert_eq!(drawn(&reopened(&session), &footer.id), live);
    }
}

#[test]
fn an_all_zero_transform_reopens_with_the_orientation_it_inherits() {
    let bytes = deck(
        &sp(
            7,
            "Zeroed",
            r#"<p:ph type="pic" idx="20"/>"#,
            r#"<p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></a:xfrm></p:spPr>"#,
        ),
        &sp(
            8,
            "Picture Placeholder",
            r#"<p:ph type="pic" idx="20"/>"#,
            &filled(
                r#"<a:xfrm rot="5400000" flipH="1"><a:off x="1905000" y="952500"/><a:ext cx="2857500" cy="1905000"/></a:xfrm>"#,
            ),
        ),
        "",
    );
    let session = DeckSession::open(&bytes, 1).unwrap();
    let zeroed = shape(&session, 0);
    assert_eq!(
        drawn(&session, &zeroed.id),
        (
            (1_905_000, 952_500, 2_857_500, 1_905_000),
            turned(90.0, true)
        )
    );

    session
        .move_shape(
            &context(),
            &slide_id(&session),
            &zeroed.id,
            2_000_250,
            952_500,
        )
        .unwrap();

    let live = drawn(&session, &zeroed.id);
    assert_eq!(
        live,
        (
            (2_000_250, 952_500, 2_857_500, 1_905_000),
            turned(90.0, true)
        )
    );
    assert_eq!(drawn(&reopened(&session), &zeroed.id), live);
}
