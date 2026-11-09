use super::*;
use crate::test_support::zip;

const P: &str = "http://schemas.openxmlformats.org/presentationml/2006/main";
const A: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";
const R: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

fn rels(entries: &[(&str, &str, &str)]) -> String {
    let mut xml = String::from("<Relationships>");
    for (id, kind, target) in entries {
        xml.push_str(&format!(
            r#"<Relationship Id="{id}" Type="{R}/{kind}" Target="{target}"/>"#
        ));
    }
    xml.push_str("</Relationships>");
    xml
}

fn part(kind: &str, shapes: &str, tail: &str) -> String {
    format!(
        r#"<p:{kind} xmlns:p="{P}" xmlns:a="{A}" xmlns:r="{R}"><p:cSld><p:spTree>{shapes}</p:spTree></p:cSld>{tail}</p:{kind}>"#
    )
}

fn shape(ph: &str, body: &str) -> String {
    format!(
        r#"<p:sp><p:nvSpPr><p:cNvPr id="2"/><p:nvPr>{ph}</p:nvPr></p:nvSpPr><p:txBody><a:lstStyle/>{body}</p:txBody></p:sp>"#
    )
}

fn run(family: &str, text: &str) -> String {
    format!(
        r#"<a:p><a:r><a:rPr><a:latin typeface="{family}"/></a:rPr><a:t>{text}</a:t></a:r></a:p>"#
    )
}

fn fixture(
    slide: &str,
    layout: &str,
    master: &str,
    extra: &[(&str, &str)],
    deflate: bool,
) -> Vec<u8> {
    let table_styles = format!(
        r#"<a:tblStyleLst xmlns:a="{A}"><a:tblStyle styleId="table"><a:wholeTbl><a:tcTxStyle i="on"/></a:wholeTbl><a:firstRow><a:tcTxStyle b="on"/></a:firstRow><a:lastRow><a:tcTxStyle b="off"/></a:lastRow></a:tblStyle></a:tblStyleLst>"#
    );
    fixture_with_table_styles(slide, layout, master, extra, deflate, &table_styles)
}

fn fixture_with_table_styles(
    slide: &str,
    layout: &str,
    master: &str,
    extra: &[(&str, &str)],
    deflate: bool,
    table_styles: &str,
) -> Vec<u8> {
    let presentation = format!(
        r#"<p:presentation xmlns:p="{P}" xmlns:r="{R}"><p:sldIdLst><p:sldId id="256" r:id="first"/><p:sldId id="257" r:id="second"/></p:sldIdLst><p:defaultTextStyle><a:lvl1pPr xmlns:a="{A}"><a:defRPr><a:latin typeface="Default"/></a:defRPr></a:lvl1pPr></p:defaultTextStyle></p:presentation>"#
    );
    let presentation_rels = rels(&[
        ("first", "slide", "slides/one.xml"),
        ("second", "slide", "slides/two.xml"),
        ("styles", "tableStyles", "tableStyles.xml"),
    ]);
    let slide_rels = rels(&[("layout", "slideLayout", "../slideLayouts/one.xml")]);
    let layout_rels = rels(&[("master", "slideMaster", "../slideMasters/one.xml")]);
    let master_rels = rels(&[("theme", "theme", "../theme/one.xml")]);
    let theme = format!(
        r#"<a:theme xmlns:a="{A}"><a:themeElements><a:fontScheme><a:majorFont><a:latin typeface="Heading"/><a:ea typeface="Heading EA"/><a:cs typeface="Heading CS"/></a:majorFont><a:minorFont><a:latin typeface="Body"/><a:ea typeface=""/><a:cs typeface="Body CS"/><a:font script="Hans" typeface="Unused"/></a:minorFont></a:fontScheme></a:themeElements></a:theme>"#
    );
    let mut parts = vec![
        ("ppt/presentation.xml", presentation.as_bytes()),
        (
            "ppt/_rels/presentation.xml.rels",
            presentation_rels.as_bytes(),
        ),
        ("ppt/slides/one.xml", slide.as_bytes()),
        ("ppt/slides/_rels/one.xml.rels", slide_rels.as_bytes()),
        ("ppt/slideLayouts/one.xml", layout.as_bytes()),
        (
            "ppt/slideLayouts/_rels/one.xml.rels",
            layout_rels.as_bytes(),
        ),
        ("ppt/slideMasters/one.xml", master.as_bytes()),
        (
            "ppt/slideMasters/_rels/one.xml.rels",
            master_rels.as_bytes(),
        ),
        ("ppt/theme/one.xml", theme.as_bytes()),
        ("ppt/tableStyles.xml", table_styles.as_bytes()),
    ];
    parts.extend(extra.iter().map(|(name, xml)| (*name, xml.as_bytes())));
    zip(&parts, deflate)
}

fn keys(result: &FontPreflight) -> Vec<(&str, bool, bool)> {
    result
        .faces
        .iter()
        .map(|face| (face.family.as_str(), face.bold, face.italic))
        .collect()
}

#[test]
fn active_slide_excludes_unused_slides_and_theme_lists() {
    for deflate in [false, true] {
        let slide = part("sld", &shape("", &run("+mn-lt", "Text")), "");
        let layout = part("sldLayout", "", "");
        let master = part("sldMaster", "", "");
        let bytes = fixture(
            &slide,
            &layout,
            &master,
            &[("ppt/slides/two.xml", "broken")],
            deflate,
        );
        let result = pptx_font_preflight(&bytes, 0).unwrap();
        assert!(result.complete);
        assert_eq!(keys(&result), [("Body", false, false)]);
        assert_eq!(
            pptx_font_preflight(&bytes, 1).unwrap_err(),
            PreflightError::Xml
        );
        assert_eq!(
            pptx_font_preflight(&bytes, 2).unwrap_err(),
            PreflightError::MissingPart
        );
    }
}

#[test]
fn placeholder_inherits_layout_master_and_text_styles() {
    let ph = r#"<p:ph type="body" idx="3"/>"#;
    let slide = part(
        "sld",
        &shape(ph, "<a:p><a:r><a:rPr i=\"1\"/><a:t>Text</a:t></a:r></a:p>"),
        "",
    );
    let layout = part(
        "sldLayout",
        &shape(ph, "<a:p><a:pPr><a:defRPr b=\"1\"/></a:pPr></a:p>"),
        "",
    );
    let master = part(
        "sldMaster",
        &shape(ph, "<a:p/>"),
        r#"<p:txStyles><p:bodyStyle><a:lvl1pPr><a:defRPr><a:latin typeface="+mj-lt"/></a:defRPr></a:lvl1pPr></p:bodyStyle></p:txStyles>"#,
    );
    let result = pptx_font_preflight(&fixture(&slide, &layout, &master, &[], false), 0).unwrap();
    assert!(result.complete);
    assert_eq!(keys(&result), [("Heading", true, true)]);
}

#[test]
fn singleton_title_placeholder_matches_by_type() {
    let slide = part(
        "sld",
        &shape(r#"<p:ph type="ctrTitle" idx="8"/>"#, "<a:p/>"),
        "",
    );
    let layout = part(
        "sldLayout",
        &shape(
            r#"<p:ph type="title" idx="1"/>"#,
            r#"<a:p><a:endParaRPr b="1"><a:latin typeface="Empty"/></a:endParaRPr></a:p>"#,
        ),
        "",
    );
    let master = part("sldMaster", "", "");
    let result = pptx_font_preflight(&fixture(&slide, &layout, &master, &[], false), 0).unwrap();
    assert_eq!(keys(&result), [("Empty", true, false)]);
}

#[test]
fn inherited_shapes_hidden_groups_and_show_master_shapes() {
    let hidden = shape("", &run("Hidden", "Text")).replace("id=\"2\"", "id=\"2\" hidden=\"1\"");
    let slide = part(
        "sld",
        &format!("{}{}", shape("", &run("Own", "Text")), hidden),
        "",
    );
    let layout = part("sldLayout", &shape("", &run("Layout", "Text")), "");
    let master = part(
        "sldMaster",
        &format!(
            "{}{}",
            shape("", &run("Master", "Text")),
            shape("<p:ph/>", &run("Unused Placeholder", "Text"))
        ),
        "",
    );
    let result = pptx_font_preflight(&fixture(&slide, &layout, &master, &[], false), 0).unwrap();
    assert!(result.complete);
    assert_eq!(
        keys(&result),
        [
            ("Layout", false, false),
            ("Master", false, false),
            ("Own", false, false)
        ]
    );
    let slide = slide.replace("<p:sld ", "<p:sld showMasterSp=\"0\" ");
    let result = pptx_font_preflight(&fixture(&slide, &layout, &master, &[], false), 0).unwrap();
    assert_eq!(keys(&result), [("Own", false, false)]);
}

#[test]
fn theme_slots_and_bullet_fonts_carry_engine_scripts() {
    let body = r#"<a:p><a:pPr><a:buFont typeface="+mj-lt"/><a:buChar char="•"/></a:pPr><a:r><a:rPr b="1"><a:latin typeface="+mn-cs"/></a:rPr><a:t>العربية</a:t></a:r></a:p><a:p><a:r><a:rPr lang="ja-JP"><a:latin typeface="+mn-ea"/><a:ea typeface="Unused EA"/><a:cs typeface="Unused CS"/></a:rPr><a:t>漢字</a:t></a:r></a:p>"#;
    let slide = part("sld", &shape("", body), "");
    let result = pptx_font_preflight(
        &fixture(
            &slide,
            &part("sldLayout", "", ""),
            &part("sldMaster", "", ""),
            &[],
            false,
        ),
        0,
    )
    .unwrap();
    assert!(result.complete);
    assert_eq!(
        keys(&result),
        [
            ("Body", false, false),
            ("Body CS", true, false),
            ("Heading", true, false)
        ]
    );
    assert_eq!(result.faces[0].scripts, ["cjk-jp"]);
    assert_eq!(result.faces[1].scripts, ["arabic"]);
    assert_eq!(result.faces[2].scripts, ["arabic"]);
}

#[test]
fn table_styles_apply_only_to_painted_cells() {
    let cell = |text| {
        format!(
            "<a:tc><a:txBody><a:lstStyle/>{}</a:txBody></a:tc>",
            run("+mn-lt", text)
        )
    };
    let table = format!(
        r#"<p:graphicFrame><a:graphic><a:graphicData><a:tbl><a:tblPr firstRow="1"><a:tableStyleId>table</a:tableStyleId></a:tblPr><a:tblGrid><a:gridCol/></a:tblGrid><a:tr>{}</a:tr><a:tr>{}</a:tr></a:tbl></a:graphicData></a:graphic></p:graphicFrame>"#,
        cell("Header"),
        cell("Value")
    );
    let slide = part("sld", &table, "");
    let result = pptx_font_preflight(
        &fixture(
            &slide,
            &part("sldLayout", "", ""),
            &part("sldMaster", "", ""),
            &[],
            false,
        ),
        0,
    )
    .unwrap();
    assert!(result.complete);
    assert_eq!(keys(&result), [("Body", false, true), ("Body", true, true)]);
}

#[test]
fn table_style_def_preserves_inherited_bold_and_italic() {
    let slide = part(
        "sld",
        &format!(
            r#"<p:graphicFrame><a:graphic><a:graphicData><a:tbl><a:tblPr firstRow="1"><a:tableStyleId>table</a:tableStyleId></a:tblPr><a:tr><a:tc><a:txBody>{}</a:txBody></a:tc></a:tr></a:tbl></a:graphicData></a:graphic></p:graphicFrame>"#,
            run("+mn-lt", "Text")
        ),
        "",
    );
    for (value, expected) in [
        ("def", true),
        ("unknown", true),
        ("off", false),
        ("0", false),
        ("false", false),
        ("on", true),
        ("1", true),
        ("true", true),
    ] {
        for (attribute, bold, italic) in [("b", expected, false), ("i", false, expected)] {
            let styles = format!(
                r#"<a:tblStyleLst xmlns:a="{A}"><a:tblStyle styleId="table"><a:wholeTbl><a:tcTxStyle {attribute}="on"/></a:wholeTbl><a:firstRow><a:tcTxStyle {attribute}="{value}"/></a:firstRow></a:tblStyle></a:tblStyleLst>"#
            );
            let result = pptx_font_preflight(
                &fixture_with_table_styles(
                    &slide,
                    &part("sldLayout", "", ""),
                    &part("sldMaster", "", ""),
                    &[],
                    false,
                    &styles,
                ),
                0,
            )
            .unwrap();
            assert!(result.complete);
            assert_eq!(
                keys(&result),
                [("Body", bold, italic)],
                "{attribute}={value}"
            );
        }
    }
}

#[test]
fn table_without_properties_exhausts_the_shared_node_budget() {
    let rows = "<a:tr><a:tc><a:txBody><a:p/></a:txBody></a:tc></a:tr>".repeat(50_000);
    let slide = part(
        "sld",
        &format!(
            "<p:graphicFrame><a:graphic><a:graphicData><a:tbl>{rows}</a:tbl></a:graphicData></a:graphic></p:graphicFrame>"
        ),
        "",
    );
    let bytes = fixture(
        &slide,
        &part("sldLayout", "", ""),
        &part("sldMaster", "", ""),
        &[],
        false,
    );
    let mut parts = Parts {
        zip: Zip::new(&bytes).unwrap(),
        nodes: xml::MAX_NODES,
    };
    let mut complete = true;
    assert!(
        parts
            .part("ppt/slides/one.xml", &mut complete)
            .unwrap()
            .is_some()
    );
    assert!(complete);
    assert_eq!(
        pptx_font_preflight(&bytes, 0).unwrap_err(),
        PreflightError::Limit
    );
}

#[test]
fn missing_table_style_regions_charge_attribute_free_sibling_searches() {
    let rows = "<a:tr><a:tc><a:txBody><a:p/></a:txBody></a:tc></a:tr>".repeat(256);
    let slide = part(
        "sld",
        &format!(
            "<p:graphicFrame><a:graphic><a:graphicData><a:tbl><a:tblPr><a:tableStyleId>table</a:tableStyleId></a:tblPr>{rows}</a:tbl></a:graphicData></a:graphic></p:graphicFrame>"
        ),
        "",
    );
    let regions = "<a:firstRow/>".repeat(1024);
    let styles = format!(
        r#"<a:tblStyleLst xmlns:a="{A}"><a:tblStyle styleId="table">{regions}</a:tblStyle></a:tblStyleLst>"#
    );
    let bytes = fixture_with_table_styles(
        &slide,
        &part("sldLayout", "", ""),
        &part("sldMaster", "", ""),
        &[],
        false,
        &styles,
    );
    assert_eq!(
        pptx_font_preflight(&bytes, 0).unwrap_err(),
        PreflightError::Limit
    );
}

#[test]
fn missing_master_styles_charge_attribute_free_sibling_searches() {
    let slide = part("sld", &shape("", &"<a:p/>".repeat(256)), "");
    let master = part("sldMaster", "", &"<p:extLst/>".repeat(1024));
    let bytes = fixture(&slide, &part("sldLayout", "", ""), &master, &[], false);
    assert_eq!(
        pptx_font_preflight(&bytes, 0).unwrap_err(),
        PreflightError::Limit
    );
}

#[test]
fn missing_dependencies_and_unresolved_graphics_are_incomplete() {
    let slide = part("sld", &shape("", &run("Explicit", "Text")), "");
    let presentation = format!(
        r#"<p:presentation xmlns:p="{P}" xmlns:r="{R}"><p:sldIdLst><p:sldId r:id="slide"/></p:sldIdLst></p:presentation>"#
    );
    let rels = rels(&[("slide", "slide", "slides/one.xml")]);
    let bytes = zip(
        &[
            ("ppt/presentation.xml", presentation.as_bytes()),
            ("ppt/_rels/presentation.xml.rels", rels.as_bytes()),
            ("ppt/slides/one.xml", slide.as_bytes()),
        ],
        false,
    );
    let result = pptx_font_preflight(&bytes, 0).unwrap();
    assert!(!result.complete);
    assert_eq!(keys(&result), [("Explicit", false, false)]);
    let slide = part(
        "sld",
        r#"<p:graphicFrame><a:graphic><a:graphicData><c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" r:id="chart"/></a:graphicData></a:graphic></p:graphicFrame>"#,
        "",
    );
    assert!(
        !pptx_font_preflight(
            &fixture(
                &slide,
                &part("sldLayout", "", ""),
                &part("sldMaster", "", ""),
                &[],
                false
            ),
            0
        )
        .unwrap()
        .complete
    );
}

#[test]
fn malformed_zip_xml_and_missing_parts_are_errors() {
    assert_eq!(
        pptx_font_preflight(b"invalid", 0).unwrap_err(),
        PreflightError::Zip
    );
    assert_eq!(
        pptx_font_preflight(&zip(&[], false), 0).unwrap_err(),
        PreflightError::MissingPart
    );
    let bytes = fixture(
        "<p:sld>",
        &part("sldLayout", "", ""),
        &part("sldMaster", "", ""),
        &[],
        false,
    );
    assert_eq!(
        pptx_font_preflight(&bytes, 0).unwrap_err(),
        PreflightError::Xml
    );
    for end in 0..bytes.len() {
        assert!(pptx_font_preflight(&bytes[..end], 0).is_err());
    }
}

#[test]
fn font_budget_returns_partial_results_and_work_budget_errors() {
    let body: String = (0..513)
        .map(|index| run(&format!("Font{index}"), "x"))
        .collect();
    let slide = part("sld", &shape("", &body), "");
    let result = pptx_font_preflight(
        &fixture(
            &slide,
            &part("sldLayout", "", ""),
            &part("sldMaster", "", ""),
            &[],
            false,
        ),
        0,
    )
    .unwrap();
    assert!(!result.complete);
    assert_eq!(result.faces.len(), 512);
    let mut nodes = 20;
    let root = xml::parse(part("sld", "", "").as_bytes(), &mut nodes).unwrap();
    let mut scan = Scan {
        presentation: &root,
        layout: None,
        master: None,
        table_styles: None,
        theme: DrawingTheme::parse(None),
        faces: Faces::default(),
        complete: true,
        work: 0,
        limited: false,
    };
    assert!(!scan.charge());
    assert!(!scan.complete);
    assert_eq!(scan.finish().unwrap_err(), PreflightError::Limit);
}
