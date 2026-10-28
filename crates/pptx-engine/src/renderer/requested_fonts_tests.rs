use super::*;
use pptx_edit::{DeckSession, EditCtx, TextStylePatch};
use serde_json::{Value, json};

const NS: &str = r#"xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main""#;

fn relationships(entries: &str) -> String {
    format!(
        r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">{entries}</Relationships>"#
    )
}

fn relationship(id: &str, kind: &str, target: &str) -> String {
    format!(
        r#"<Relationship Id="{id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/{kind}" Target="{target}"/>"#
    )
}

fn tree(shapes: &str) -> String {
    format!(
        r#"<p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>{shapes}</p:spTree></p:cSld>"#
    )
}

fn run(text: &str, properties: &str) -> String {
    format!(r#"<a:r><a:rPr {properties}/><a:t>{text}</a:t></a:r>"#)
}

fn font_run(text: &str, family: &str) -> String {
    format!(r#"<a:r><a:rPr><a:latin typeface="{family}"/></a:rPr><a:t>{text}</a:t></a:r>"#)
}

fn shape(id: u32, placeholder: &str, list: &str, paragraphs: &str) -> String {
    format!(
        r#"<p:sp><p:nvSpPr><p:cNvPr id="{id}" name="Text {id}"/><p:cNvSpPr/><p:nvPr>{placeholder}</p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="4000000" cy="2000000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr/><a:lstStyle>{list}</a:lstStyle>{paragraphs}</p:txBody></p:sp>"#
    )
}

fn text_shape(id: u32, family: &str) -> String {
    shape(
        id,
        "",
        "",
        &format!("<a:p>{}</a:p>", font_run("Text", family)),
    )
}

#[derive(Default)]
struct Fixture<'a> {
    slides: &'a [&'a str],
    slide_attributes: &'a str,
    layout: &'a str,
    layout_attributes: &'a str,
    master: &'a str,
    title_style: &'a str,
    other_style: &'a str,
    defaults: &'a str,
}

impl Fixture<'_> {
    fn session(&self) -> DeckSession {
        let mut ids = String::new();
        let mut slide_rels = String::new();
        let mut parts = Vec::new();
        let mut overrides = String::new();
        for (index, shapes) in self.slides.iter().enumerate() {
            let number = index + 1;
            ids.push_str(&format!(
                r#"<p:sldId id="{}" r:id="slide{number}"/>"#,
                256 + index
            ));
            slide_rels.push_str(&relationship(
                &format!("slide{number}"),
                "slide",
                &format!("slides/slide{number}.xml"),
            ));
            overrides.push_str(&format!(
                r#"<Override PartName="/ppt/slides/slide{number}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>"#
            ));
            parts.push((
                format!("ppt/slides/slide{number}.xml"),
                format!(
                    r#"<p:sld {NS} {}>{}</p:sld>"#,
                    self.slide_attributes,
                    tree(shapes)
                ),
            ));
            parts.push((
                format!("ppt/slides/_rels/slide{number}.xml.rels"),
                relationships(&relationship(
                    "layout",
                    "slideLayout",
                    "../slideLayouts/slideLayout1.xml",
                )),
            ));
        }
        parts.extend([
            ("[Content_Types].xml".to_owned(), format!(
                r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/><Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/><Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/><Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>{overrides}</Types>"#
            )),
            ("_rels/.rels".to_owned(), relationships(&relationship("office", "officeDocument", "ppt/presentation.xml"))),
            ("ppt/presentation.xml".to_owned(), format!(
                r#"<p:presentation {NS}><p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="master"/></p:sldMasterIdLst><p:sldIdLst>{ids}</p:sldIdLst><p:sldSz cx="12192000" cy="6858000"/><p:defaultTextStyle>{}</p:defaultTextStyle></p:presentation>"#, self.defaults
            )),
            ("ppt/_rels/presentation.xml.rels".to_owned(), relationships(&format!("{}{slide_rels}", relationship("master", "slideMaster", "slideMasters/slideMaster1.xml")))),
            ("ppt/slideLayouts/slideLayout1.xml".to_owned(), format!(r#"<p:sldLayout {NS} {}>{}</p:sldLayout>"#, self.layout_attributes, tree(self.layout))),
            ("ppt/slideLayouts/_rels/slideLayout1.xml.rels".to_owned(), relationships(&relationship("master", "slideMaster", "../slideMasters/slideMaster1.xml"))),
            ("ppt/slideMasters/slideMaster1.xml".to_owned(), format!(
                r#"<p:sldMaster {NS}>{}<p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="layout"/></p:sldLayoutIdLst><p:txStyles><p:titleStyle>{}</p:titleStyle><p:otherStyle>{}</p:otherStyle></p:txStyles></p:sldMaster>"#, tree(self.master), self.title_style, self.other_style
            )),
            ("ppt/slideMasters/_rels/slideMaster1.xml.rels".to_owned(), relationships(&format!("{}{}", relationship("layout", "slideLayout", "../slideLayouts/slideLayout1.xml"), relationship("theme", "theme", "../theme/theme1.xml")))),
            ("ppt/theme/theme1.xml".to_owned(), format!(
                r#"<a:theme {NS} name="Fixture"><a:themeElements><a:fontScheme name="Fixture"><a:majorFont><a:latin typeface="Major Latin"/><a:ea typeface="Major EA"/><a:cs typeface="Major CS"/></a:majorFont><a:minorFont><a:latin typeface="Minor Latin"/><a:ea typeface="Minor EA"/><a:cs typeface="Minor CS"/></a:minorFont></a:fontScheme></a:themeElements></a:theme>"#
            )),
        ]);
        let parts = parts
            .into_iter()
            .map(|(path, xml)| (path, xml.into_bytes()))
            .collect::<Vec<_>>();
        let bytes = ooxml_opc::rezip_parts(&parts).unwrap();
        DeckSession::open(&bytes, 17101).unwrap()
    }
}

fn faces(renderer: &Renderer, session: &DeckSession, id: Option<&str>) -> Value {
    serde_json::from_str::<Value>(&renderer.requested_font_faces_json(session, id).unwrap())
        .unwrap()["faces"]
        .clone()
}

fn assert_agreement(session: &DeckSession) {
    let renderer = Renderer::new();
    let layout = pptx_render::SlideRenderer::new();
    for (index, id) in session.slide_ids().unwrap().iter().enumerate() {
        let collected = faces(&renderer, session, Some(id));
        let collected = collected
            .as_array()
            .unwrap()
            .iter()
            .map(|face| face["family"].as_str().unwrap().trim().to_lowercase())
            .collect::<HashSet<_>>();
        let scope = session.slide_scope(index).unwrap();
        let rendered = layout
            .layout_scoped_slide(session.package(), &scope)
            .unwrap();
        let requested = rendered
            .font_substitutions
            .iter()
            .map(|substitution| substitution.requested_family.trim().to_lowercase())
            .collect::<HashSet<_>>();
        assert_eq!(collected, requested, "{id}");
    }
    assert_eq!(renderer.layout_count, 0);
    assert!(renderer.rendered.is_empty());
    assert!(renderer.layout_keys.is_empty());
}

#[test]
fn theme_explicit_script_and_field_requests_agree_with_layout() {
    let scripts = r#"<a:r><a:rPr><a:latin typeface="Latin For Scripts"/><a:ea typeface="Unused EA"/><a:cs typeface="Unused CS"/></a:rPr><a:t>日本語 العربية</a:t></a:r>"#;
    let field = r#"<a:fld id="number" type="slidenum"><a:rPr i="1"><a:latin typeface="Field Face"/></a:rPr><a:t>1</a:t></a:fld>"#;
    let slide = [
        text_shape(2, "+mn-lt"),
        text_shape(3, "+mj-lt"),
        text_shape(4, " Explicit Latin "),
        text_shape(5, "+mn-ea"),
        text_shape(6, "+mj-cs"),
        shape(7, "", "", &format!("<a:p>{scripts}{field}</a:p>")),
        text_shape(8, "Symbol"),
    ]
    .concat();
    let session = Fixture {
        slides: &[&slide],
        ..Fixture::default()
    }
    .session();
    assert_agreement(&session);
    let requests = faces(&Renderer::new(), &session, None);
    let families = requests
        .as_array()
        .unwrap()
        .iter()
        .map(|face| face["family"].as_str().unwrap())
        .collect::<Vec<_>>();
    assert_eq!(
        families,
        [
            "Minor Latin",
            "Major Latin",
            "Explicit Latin",
            "Minor EA",
            "Major CS",
            "Latin For Scripts",
            "Field Face",
            "Symbol"
        ]
    );
}

#[test]
fn bullets_tables_and_groups_agree_with_layout() {
    let bullet = shape(
        2,
        "",
        r#"<a:lvl1pPr><a:buFont typeface="Bullet Face"/><a:buChar char="*"/><a:defRPr i="1"><a:latin typeface="List Face"/></a:defRPr></a:lvl1pPr>"#,
        &format!("<a:p>{}</a:p>", run("Bullet", "")),
    );
    let table = r#"<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="3" name="Table"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="0" y="0"/><a:ext cx="4000000" cy="1000000"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table"><a:tbl><a:tblPr/><a:tblGrid><a:gridCol w="4000000"/></a:tblGrid><a:tr h="1000000"><a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr><a:latin typeface="Cell Face"/></a:rPr><a:t>Cell</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc></a:tr></a:tbl></a:graphicData></a:graphic></p:graphicFrame>"#;
    let group = format!(
        r#"<p:grpSp><p:nvGrpSpPr><p:cNvPr id="4" name="Group"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="4000000" cy="2000000"/><a:chOff x="0" y="0"/><a:chExt cx="4000000" cy="2000000"/></a:xfrm></p:grpSpPr>{}</p:grpSp>"#,
        text_shape(5, "Group Face")
    );
    let slide = format!("{bullet}{table}{group}");
    let session = Fixture {
        slides: &[&slide],
        ..Fixture::default()
    }
    .session();
    assert_agreement(&session);
    assert_eq!(
        faces(&Renderer::new(), &session, None),
        json!([
            {"family":"List Face","bold":false,"italic":true},
            {"family":"Bullet Face","bold":false,"italic":true},
            {"family":"Cell Face","bold":false,"italic":false},
            {"family":"Group Face","bold":false,"italic":false},
        ])
    );
}

#[test]
fn placeholder_inheritance_and_parent_text_agree_with_layout() {
    let title = shape(
        2,
        r#"<p:ph type="title"/>"#,
        "",
        &format!(
            "<a:p>{}{}</a:p>",
            run("Title", ""),
            run("Italic", r#"i="1""#)
        ),
    );
    let master_title = shape(
        10,
        r#"<p:ph type="title"/>"#,
        "",
        &format!(
            "<a:p>{}</a:p>",
            font_run("Not painted", "Unused Master Placeholder")
        ),
    );
    let master = format!("{master_title}{}", text_shape(11, "Master Face"));
    let layout = text_shape(12, "Layout Face");
    let session = Fixture {
        slides: &[&title], master: &master, layout: &layout,
        title_style: r#"<a:lvl1pPr><a:defRPr b="1"><a:latin typeface="+mj-lt"/></a:defRPr></a:lvl1pPr>"#,
        ..Fixture::default()
    }.session();
    assert_agreement(&session);
    assert_eq!(
        faces(&Renderer::new(), &session, None),
        json!([
            {"family":"Master Face","bold":false,"italic":false},
            {"family":"Layout Face","bold":false,"italic":false},
            {"family":"Major Latin","bold":true,"italic":false},
            {"family":"Major Latin","bold":true,"italic":true},
        ])
    );
}

#[test]
fn slide_scopes_and_whole_deck_keep_first_face_order() {
    let first = format!(
        "{}{}{}",
        text_shape(2, " First "),
        text_shape(3, "SHARED"),
        shape(
            4,
            "",
            "",
            &format!("<a:p>{}</a:p>", run("Italic", r#"i="1""#))
        )
    );
    let second = format!("{}{}", text_shape(2, "shared"), text_shape(3, "Second"));
    let session = Fixture {
        slides: &[&first, &second],
        ..Fixture::default()
    }
    .session();
    assert_agreement(&session);
    let renderer = Renderer::new();
    let ids = session.slide_ids().unwrap();
    assert_eq!(
        faces(&renderer, &session, Some(&ids[1])),
        json!([
            {"family":"shared","bold":false,"italic":false},
            {"family":"Second","bold":false,"italic":false},
        ])
    );
    let expected = json!([
        {"family":"First","bold":false,"italic":false},
        {"family":"SHARED","bold":false,"italic":false},
        {"family":"Minor Latin","bold":false,"italic":true},
        {"family":"Second","bold":false,"italic":false},
    ]);
    assert_eq!(faces(&renderer, &session, None), expected);
    assert_eq!(faces(&renderer, &session, None), expected);
    let error = renderer
        .requested_font_faces_json(&session, Some("missing"))
        .unwrap_err();
    assert_eq!(
        error,
        engine_error(pptx_edit::EditError::SlideNotFound("missing".to_owned()))
    );
}

#[test]
fn collecting_keeps_layouts_keys_and_font_epoch_unchanged() {
    let slide = text_shape(2, "Cache Face");
    let session = Fixture {
        slides: &[&slide],
        ..Fixture::default()
    }
    .session();
    let mut renderer = Renderer::new();
    let id = session.slide_ids().unwrap().remove(0);
    let before = renderer.layout_slide_json(&session, 0).unwrap();
    let key = renderer.slide_layout_key(&session, 0).unwrap();
    let counts = (
        renderer.layout_count,
        renderer.key_count,
        renderer.font_epoch,
        renderer.clock,
        renderer.generation,
        renderer.next_token,
    );
    for slide_id in [Some(id.as_str()), None] {
        faces(&renderer, &session, slide_id);
        assert_eq!(
            (
                renderer.layout_count,
                renderer.key_count,
                renderer.font_epoch,
                renderer.clock,
                renderer.generation,
                renderer.next_token
            ),
            counts
        );
    }
    assert_eq!(renderer.layout_slide_json(&session, 0).unwrap(), before);
    assert_eq!(renderer.slide_layout_key(&session, 0).unwrap(), key);
    assert_eq!(renderer.layout_count, counts.0);
    assert_eq!(renderer.key_count, counts.1);
}

#[test]
fn collecting_uses_live_run_styles_without_a_layout() {
    let slide = text_shape(2, "Original Face");
    let session = Fixture {
        slides: &[&slide],
        ..Fixture::default()
    }
    .session();
    let scope = session.slide_scope(0).unwrap();
    let story = &scope.slide.shapes[0].text_stories[0];
    session
        .format_text(
            &EditCtx::local("fonts"),
            &story.id,
            0,
            4,
            &TextStylePatch {
                font_family: Some("Edited Face".to_owned()),
                bold: Some(true),
                italic: Some(true),
                ..TextStylePatch::default()
            },
        )
        .unwrap();
    assert_agreement(&session);
    assert_eq!(
        faces(&Renderer::new(), &session, None),
        json!([
            {"family":"Edited Face","bold":true,"italic":true},
        ])
    );
}

#[test]
fn parent_visibility_and_hidden_shapes_match_layout() {
    let hidden =
        text_shape(3, "Hidden Face").replace(r#"name="Text 3""#, r#"name="Text 3" hidden="1""#);
    let slide = format!("{}{hidden}", text_shape(2, "Slide Face"));
    let layout = text_shape(10, "Layout Face");
    let master = text_shape(11, "Master Face");
    for (slide_attributes, layout_attributes, expected) in [
        (r#"showMasterSp="0""#, "", vec!["Slide Face"]),
        ("", r#"showMasterSp="0""#, vec!["Layout Face", "Slide Face"]),
        ("", "", vec!["Master Face", "Layout Face", "Slide Face"]),
    ] {
        let session = Fixture {
            slides: &[&slide],
            slide_attributes,
            layout_attributes,
            layout: &layout,
            master: &master,
            ..Fixture::default()
        }
        .session();
        assert_agreement(&session);
        let requests = faces(&Renderer::new(), &session, None);
        assert_eq!(
            requests
                .as_array()
                .unwrap()
                .iter()
                .map(|face| face["family"].as_str().unwrap())
                .collect::<Vec<_>>(),
            expected
        );
    }
}

#[test]
fn defaults_list_levels_and_empty_paragraphs_match_layout() {
    let slide = [
        shape(2, "", "", &format!("<a:p>{}</a:p>", run("Default", ""))),
        shape(
            3,
            "",
            r#"<a:lvl1pPr><a:defRPr><a:latin typeface="Local Face"/></a:defRPr></a:lvl1pPr>"#,
            &format!("<a:p>{}</a:p>", run("Local", "")),
        ),
        shape(
            4,
            "",
            "",
            &format!(r#"<a:p><a:pPr lvl="1"/>{}</a:p>"#, run("Level", "")),
        ),
        shape(
            5,
            "",
            "",
            r#"<a:p><a:endParaRPr i="0"><a:latin typeface="Empty Face"/></a:endParaRPr></a:p>"#,
        ),
    ]
    .concat();
    let session = Fixture {
        slides: &[&slide],
        defaults: r#"<a:defPPr><a:defRPr i="1"><a:latin typeface="Default Face"/></a:defRPr></a:defPPr><a:lvl2pPr><a:defRPr><a:latin typeface="Level Face"/></a:defRPr></a:lvl2pPr>"#,
        other_style: r#"<a:lvl1pPr><a:defRPr b="1"><a:latin typeface="Other Face"/></a:defRPr></a:lvl1pPr>"#,
        ..Fixture::default()
    }.session();
    assert_agreement(&session);
    assert_eq!(
        faces(&Renderer::new(), &session, None),
        json!([
            {"family":"Default Face","bold":true,"italic":true},
            {"family":"Local Face","bold":true,"italic":true},
            {"family":"Level Face","bold":false,"italic":true},
            {"family":"Empty Face","bold":true,"italic":false},
        ])
    );
}

#[test]
fn layout_placeholder_styles_and_explicit_weight_override_master_defaults() {
    let placeholder = r#"<p:ph type="title"/>"#;
    let slide = shape(
        2,
        placeholder,
        "",
        &format!("<a:p>{}</a:p>", run("Title", r#"b="0""#)),
    );
    let layout = shape(
        10,
        placeholder,
        r#"<a:lvl1pPr><a:defRPr i="1"><a:latin typeface="Layout Title"/></a:defRPr></a:lvl1pPr>"#,
        &format!(
            "<a:p>{}</a:p>",
            font_run("Not painted", "Unused Layout Run")
        ),
    );
    let session = Fixture {
        slides: &[&slide], layout: &layout,
        title_style: r#"<a:lvl1pPr><a:defRPr b="1"><a:latin typeface="+mj-lt"/></a:defRPr></a:lvl1pPr>"#,
        ..Fixture::default()
    }.session();
    assert_agreement(&session);
    assert_eq!(
        faces(&Renderer::new(), &session, None),
        json!([
            {"family":"Layout Title","bold":false,"italic":true},
        ])
    );
}

#[test]
fn emulated_symbol_bullets_still_request_the_named_face() {
    let slide = shape(
        2,
        "",
        "",
        &format!(
            r#"<a:p><a:pPr><a:buFont typeface="Wingdings"/><a:buChar char="§"/></a:pPr>{}</a:p>"#,
            font_run("Bullet", "Text Face"),
        ),
    );
    let session = Fixture {
        slides: &[&slide],
        ..Fixture::default()
    }
    .session();
    assert_eq!(
        faces(&Renderer::new(), &session, None),
        json!([
            {"family":"Text Face","bold":false,"italic":false},
            {"family":"Wingdings","bold":false,"italic":false},
        ])
    );
}

#[test]
fn empty_deck_has_no_requests() {
    let session = Fixture::default().session();
    assert_eq!(faces(&Renderer::new(), &session, None), json!([]));
}

#[test]
fn requests_do_not_depend_on_registered_faces() {
    let slide = text_shape(2, "Requested Face");
    let session = Fixture {
        slides: &[&slide],
        ..Fixture::default()
    }
    .session();
    let mut renderer = Renderer::new();
    let expected = faces(&renderer, &session, None);
    renderer
        .register_fallback_font(
            "Fallback",
            false,
            false,
            include_bytes!("../../../ooxml-text/tests/fonts/LiberationSans-Regular.ttf"),
        )
        .unwrap();
    assert_eq!(faces(&renderer, &session, None), expected);
    assert_eq!(renderer.layout_count, 0);
}

fn chart_session(chart: pptx_parse::ChartSpace, width: i64) -> DeckSession {
    use pptx_parse::{ChartPart, GraphicFrame, GraphicFrameData, ShapeNode};

    let slide = text_shape(2, "Unused Text Face");
    let original = Fixture {
        slides: &[&slide],
        ..Fixture::default()
    }
    .session();
    let mut package = original.package().clone();
    let ShapeNode::Shape(shape) = &package.slides[0].shapes[0] else {
        panic!("expected a text shape")
    };
    let mut base = shape.base.clone();
    base.transform.width = width;
    base.transform.height = 6_000_000;
    package.slides[0].shapes[0] = ShapeNode::GraphicFrame(GraphicFrame {
        base,
        data: GraphicFrameData::Chart {
            relationship_id: "chart".to_owned(),
            part_path: Some("ppt/charts/chart1.xml".to_owned()),
        },
    });
    package.charts.push(ChartPart {
        part_path: "ppt/charts/chart1.xml".to_owned(),
        theme_part_path: Some(package.themes[0].part_path.clone()),
        chart,
    });
    DeckSession::from_package_with_source(package, b"font chart", 17102).unwrap()
}

#[test]
fn chart_fonts_agree_with_layout() {
    use pptx_parse::{ChartLegend, ChartSeries, ChartSpace, ChartTextProperties};

    let session = chart_session(
        ChartSpace {
            chart_type: "column".to_owned(),
            title: Some("Title".to_owned()),
            text: Some(ChartTextProperties {
                font: Some("Chart Face".to_owned()),
                ..ChartTextProperties::default()
            }),
            title_text: Some(ChartTextProperties {
                font: Some("Chart Title".to_owned()),
                italic: Some(true),
                ..ChartTextProperties::default()
            }),
            legend: Some(ChartLegend {
                position: Some("right".to_owned()),
                visible: true,
                overlay: false,
                text: Some(ChartTextProperties {
                    font: Some("Chart Legend".to_owned()),
                    ..ChartTextProperties::default()
                }),
            }),
            series: vec![ChartSeries {
                name: Some("Series".to_owned()),
                categories: vec!["First".to_owned(), "Second".to_owned()],
                values: vec![1.0, 2.0],
                color: "#112233".to_owned(),
                ..ChartSeries::default()
            }],
            ..ChartSpace::default()
        },
        10_000_000,
    );
    assert_agreement(&session);
    let requests = faces(&Renderer::new(), &session, None);
    assert!(requests.as_array().unwrap().contains(&json!({
        "family":"Chart Title","bold":true,"italic":true,
    })));
}

#[test]
fn narrow_chart_requests_may_be_a_superset_of_layout() {
    use pptx_parse::{ChartAxis, ChartPlotGroup, ChartSeries, ChartSpace, ChartTextProperties};

    let session = chart_session(
        ChartSpace {
            chart_type: "column".to_owned(),
            axis_list: Some(vec![
                ChartAxis {
                    id: Some("category".to_owned()),
                    axis_type: "category".to_owned(),
                    text: Some(ChartTextProperties {
                        font: Some("Category Face".to_owned()),
                        ..ChartTextProperties::default()
                    }),
                    ..ChartAxis::default()
                },
                ChartAxis {
                    id: Some("value".to_owned()),
                    axis_type: "value".to_owned(),
                    number_format: Some("$#,##0.00".to_owned()),
                    text: Some(ChartTextProperties {
                        font: Some("Value Face".to_owned()),
                        ..ChartTextProperties::default()
                    }),
                    ..ChartAxis::default()
                },
            ]),
            plot_groups: vec![ChartPlotGroup {
                chart_type: Some("column".to_owned()),
                axis_ids: vec!["category".to_owned(), "value".to_owned()],
                series: vec![ChartSeries {
                    categories: (0..10).map(|index| format!("Category {index}")).collect(),
                    values: vec![1_000_000.0; 10],
                    color: "#112233".to_owned(),
                    ..ChartSeries::default()
                }],
                ..ChartPlotGroup::default()
            }],
            ..ChartSpace::default()
        },
        952_500,
    );
    let collected = faces(&Renderer::new(), &session, None);
    let collected = collected
        .as_array()
        .unwrap()
        .iter()
        .map(|face| face["family"].as_str().unwrap().to_lowercase())
        .collect::<HashSet<_>>();
    let scope = session.slide_scope(0).unwrap();
    let rendered = pptx_render::SlideRenderer::new()
        .layout_scoped_slide(session.package(), &scope)
        .unwrap();
    let requested = rendered
        .font_substitutions
        .iter()
        .map(|substitution| substitution.requested_family.to_lowercase())
        .collect::<HashSet<_>>();
    assert!(requested.is_subset(&collected));
    assert!(collected.contains("category face"));
    assert!(requested.contains("value face"));
    assert!(!requested.contains("category face"));
}

#[test]
fn diagram_drawing_text_agrees_with_layout() {
    use pptx_parse::{DiagramDrawing, GraphicFrame, GraphicFrameData, ShapeNode};

    let slide = text_shape(2, "Diagram Face");
    let original = Fixture {
        slides: &[&slide],
        ..Fixture::default()
    }
    .session();
    let mut package = original.package().clone();
    let child = package.slides[0].shapes[0].clone();
    let ShapeNode::Shape(shape) = &child else {
        panic!("expected a text shape")
    };
    package.slides[0].shapes[0] = ShapeNode::GraphicFrame(GraphicFrame {
        base: shape.base.clone(),
        data: GraphicFrameData::Diagram {
            relationship_ids: Vec::new(),
            drawing_part_path: Some("ppt/diagrams/drawing1.xml".to_owned()),
        },
    });
    package.diagram_drawings.push(DiagramDrawing {
        part_path: "ppt/diagrams/drawing1.xml".to_owned(),
        synthesis_extent: Some([4_000_000, 2_000_000]),
        shapes: vec![child],
    });
    let session = DeckSession::from_package_with_source(package, b"font diagram", 17103).unwrap();
    assert_agreement(&session);
    assert_eq!(
        faces(&Renderer::new(), &session, None),
        json!([
            {"family":"Diagram Face","bold":false,"italic":false},
        ])
    );
}
