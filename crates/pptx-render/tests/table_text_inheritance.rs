use pptx_edit::DeckSession;
use pptx_render::{PositionedTextLine, Primitive, SlideRenderer, SurfaceDisplayList};

const DECK: &[u8] = include_bytes!("fixtures/table-basic.pptx");
const FONT: &[u8] = include_bytes!("../../ooxml-text/tests/fonts/LiberationSans-Regular.ttf");
const MONO: &[u8] = include_bytes!("../../../packages/fonts/assets/LiberationMono-Regular.ttf");

#[derive(Default)]
struct Fixture<'a> {
    defaults: &'a str,
    other: &'a str,
    list: &'a str,
    paragraph: &'a str,
    run: &'a str,
    autofit: &'a str,
    table_style: &'a str,
    table_style_run: &'a str,
}

impl Fixture<'_> {
    fn bytes(&self) -> Vec<u8> {
        let mut parts = ooxml_opc::unzip_parts(DECK).unwrap();
        for (path, bytes) in &mut parts {
            let xml = String::from_utf8(bytes.clone());
            let replacement = match path.as_str() {
                "ppt/presentation.xml" => {
                    let xml = xml.unwrap();
                    if self.defaults.is_empty() {
                        xml
                    } else {
                        xml.replace(
                            "</p:presentation>",
                            &format!(
                                "<p:defaultTextStyle>{}</p:defaultTextStyle></p:presentation>",
                                self.defaults
                            ),
                        )
                    }
                }
                "ppt/slideMasters/slideMaster1.xml" => {
                    let mut xml = xml.unwrap();
                    let start = xml.find("<p:otherStyle>").unwrap();
                    let end = xml.find("</p:otherStyle>").unwrap() + "</p:otherStyle>".len();
                    xml.replace_range(
                        start..end,
                        &format!("<p:otherStyle>{}</p:otherStyle>", self.other),
                    );
                    xml
                }
                "ppt/slides/slide1.xml" => {
                    let mut xml = xml.unwrap();
                    let start = xml.find("<p:graphicFrame>").unwrap();
                    let end = xml.find("</p:graphicFrame>").unwrap() + "</p:graphicFrame>".len();
                    xml.replace_range(
                        start..end,
                        &format!(
                            r#"<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="2" name="Table"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>
<p:xfrm><a:off x="609600" y="609600"/><a:ext cx="5715000" cy="952500"/></p:xfrm>
<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table"><a:tbl>
<a:tblPr><a:tableStyleId>fixture</a:tableStyleId></a:tblPr><a:tblGrid><a:gridCol w="5715000"/></a:tblGrid>
<a:tr h="952500"><a:tc><a:txBody><a:bodyPr>{}</a:bodyPr><a:lstStyle>{}</a:lstStyle>
<a:p>{}<a:r><a:rPr {}></a:rPr><a:t>Cell</a:t></a:r></a:p>
</a:txBody><a:tcPr/></a:tc></a:tr></a:tbl></a:graphicData></a:graphic></p:graphicFrame>"#,
                            self.autofit, self.list, self.paragraph, self.run
                        ),
                    );
                    xml
                }
                "ppt/tableStyles.xml" => format!(
                    r#"<a:tblStyleLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" def="fixture"><a:tblStyle styleId="fixture" styleName="Fixture"><a:wholeTbl><a:tcTxStyle {}>{}</a:tcTxStyle></a:wholeTbl></a:tblStyle></a:tblStyleLst>"#,
                    self.table_style_run, self.table_style
                ),
                _ => continue,
            };
            *bytes = replacement.into_bytes();
        }
        ooxml_opc::rezip_parts(&parts).unwrap()
    }

    fn render(&self) -> SurfaceDisplayList {
        render(&DeckSession::open(&self.bytes(), 4201).unwrap())
    }
}

fn render(session: &DeckSession) -> SurfaceDisplayList {
    let mut renderer = SlideRenderer::new();
    for bold in [false, true] {
        for italic in [false, true] {
            renderer.register_font("Arial", bold, italic, FONT).unwrap();
            renderer
                .register_font("Courier New", bold, italic, MONO)
                .unwrap();
        }
    }
    renderer
        .layout_slide(session.package(), &session.snapshot().unwrap(), 0)
        .unwrap()
        .display_list
}

fn line(list: &SurfaceDisplayList) -> &PositionedTextLine {
    let cells = list
        .primitives
        .iter()
        .find_map(|primitive| match primitive {
            Primitive::Table { primitives, .. } => Some(primitives),
            _ => None,
        })
        .unwrap();
    cells
        .iter()
        .find_map(|primitive| match primitive {
            Primitive::TextBox { lines, .. } => lines.first(),
            _ => None,
        })
        .unwrap()
}

fn assert_size(list: &SurfaceDisplayList, points: f32) {
    let run = &line(list).runs[0];
    assert_eq!(run.text, "Cell");
    assert!((run.font_size_px - points * 96.0 / 72.0).abs() < 0.001);
}

const FOURTEEN: &str = r#"<a:lvl1pPr><a:defRPr sz="1400"/></a:lvl1pPr>"#;

#[test]
fn unsized_table_cell_inherits_fourteen_points_when_presentation_and_master_agree() {
    assert_size(
        &Fixture {
            defaults: FOURTEEN,
            other: FOURTEEN,
            ..Fixture::default()
        }
        .render(),
        14.0,
    );
}

#[test]
fn table_cells_use_other_style_before_presentation_defaults_per_iso_29500_19_3_1_35() {
    assert_size(
        &Fixture {
            defaults: r#"<a:lvl1pPr><a:defRPr sz="1200"/></a:lvl1pPr>"#,
            other: r#"<a:lvl1pPr><a:defRPr sz="2000"/></a:lvl1pPr>"#,
            ..Fixture::default()
        }
        .render(),
        20.0,
    );
}

#[test]
fn table_cell_master_level_blocks_presentation_properties() {
    let list = Fixture {
        defaults: r#"<a:defPPr><a:defRPr b="1" i="1"/></a:defPPr><a:lvl1pPr><a:defRPr sz="1400" b="1"><a:latin typeface="Courier New"/></a:defRPr></a:lvl1pPr>"#,
        other: r#"<a:lvl1pPr><a:defRPr sz="1800"><a:latin typeface="+mn-lt"/><a:solidFill><a:srgbClr val="654321"/></a:solidFill></a:defRPr></a:lvl1pPr>"#,
        paragraph: r#"<a:pPr algn="l"/>"#,
        ..Fixture::default()
    }
    .render();
    assert_size(&list, 18.0);
    let run = &line(&list).runs[0];
    assert!(!run.bold);
    assert!(!run.italic);
    assert_eq!(run.font_family, "Arial");
    assert_eq!(run.color, "#654321");
}

#[test]
fn table_cell_missing_master_level_uses_presentation_properties() {
    for other in ["", r#"<a:lvl2pPr><a:defRPr sz="2000"/></a:lvl2pPr>"#] {
        let list = Fixture {
            defaults: r#"<a:lvl1pPr><a:defRPr sz="1400" b="1"><a:latin typeface="Courier New"/></a:defRPr></a:lvl1pPr>"#,
            other,
            ..Fixture::default()
        }
        .render();
        assert_size(&list, 14.0);
        assert!(line(&list).runs[0].bold);
        assert_eq!(line(&list).runs[0].font_family, "Courier New");
    }
}

#[test]
fn table_cell_empty_master_level_blocks_presentation_defaults() {
    let list = Fixture {
        defaults: r#"<a:defPPr><a:defRPr i="1"/></a:defPPr><a:lvl1pPr><a:defRPr sz="1400" b="1"/></a:lvl1pPr>"#,
        other: "<a:lvl1pPr/>",
        ..Fixture::default()
    }
    .render();
    assert_size(&list, 18.0);
    assert!(!line(&list).runs[0].bold);
    assert!(!line(&list).runs[0].italic);
}

#[test]
fn table_cell_missing_second_master_level_keeps_presentation_fallback() {
    let list = Fixture {
        defaults: r#"<a:defPPr><a:defRPr i="1"/></a:defPPr><a:lvl2pPr><a:defRPr sz="1400" b="1"/></a:lvl2pPr>"#,
        other: r#"<a:lvl1pPr><a:defRPr sz="1800"/></a:lvl1pPr>"#,
        paragraph: r#"<a:pPr lvl="1"/>"#,
        ..Fixture::default()
    }
    .render();
    assert_size(&list, 14.0);
    assert!(line(&list).runs[0].bold);
    assert!(line(&list).runs[0].italic);
}

#[test]
fn table_cell_table_style_and_local_bold_override_master_level() {
    for (table_style, list_style, paragraph, run, bold) in [
        (r#"b="1""#, "", "", "", true),
        (
            r#"b="1""#,
            r#"<a:lvl1pPr><a:defRPr b="0"/></a:lvl1pPr>"#,
            "",
            "",
            false,
        ),
        ("", "", r#"<a:pPr><a:defRPr b="1"/></a:pPr>"#, "", true),
        (
            r#"b="1""#,
            "",
            r#"<a:pPr><a:defRPr b="1"/></a:pPr>"#,
            r#"b="0""#,
            false,
        ),
    ] {
        let display = Fixture {
            defaults: r#"<a:lvl1pPr><a:defRPr sz="1400" b="1"/></a:lvl1pPr>"#,
            other: r#"<a:lvl1pPr><a:defRPr sz="1800" b="0"/></a:lvl1pPr>"#,
            list: list_style,
            paragraph,
            run,
            table_style_run: table_style,
            ..Fixture::default()
        }
        .render();
        assert_size(&display, 18.0);
        assert_eq!(line(&display).runs[0].bold, bold);
    }
}

#[test]
fn table_cell_explicit_run_size_wins_over_inherited_sizes() {
    assert_size(
        &Fixture {
            defaults: FOURTEEN,
            other: FOURTEEN,
            list: r#"<a:lvl1pPr><a:defRPr sz="2200"/></a:lvl1pPr>"#,
            paragraph: r#"<a:pPr><a:defRPr sz="2400"/></a:pPr>"#,
            run: r#"sz="1600""#,
            ..Fixture::default()
        }
        .render(),
        16.0,
    );
}

#[test]
fn table_cell_paragraph_and_list_sizes_override_master_and_presentation_sizes() {
    for (paragraph, expected) in [
        ("", 22.0),
        (r#"<a:pPr><a:defRPr sz="2400"/></a:pPr>"#, 24.0),
    ] {
        assert_size(
            &Fixture {
                defaults: FOURTEEN,
                other: FOURTEEN,
                list: r#"<a:lvl1pPr><a:defRPr sz="2200"/></a:lvl1pPr>"#,
                paragraph,
                ..Fixture::default()
            }
            .render(),
            expected,
        );
    }
}

#[test]
fn second_level_table_cell_inherits_second_level_presentation_defaults() {
    assert_size(
        &Fixture {
            defaults: r#"<a:lvl1pPr><a:defRPr sz="1400"/></a:lvl1pPr><a:lvl2pPr><a:defRPr sz="1100"/></a:lvl2pPr>"#,
            paragraph: r#"<a:pPr lvl="1"/>"#,
            ..Fixture::default()
        }
        .render(),
        11.0,
    );
}

#[test]
fn second_level_table_cell_inherits_second_level_master_other_style() {
    assert_size(
        &Fixture {
            defaults: r#"<a:lvl1pPr><a:defRPr sz="1400"/></a:lvl1pPr><a:lvl2pPr><a:defRPr sz="1100"/></a:lvl2pPr>"#,
            other: r#"<a:lvl1pPr><a:defRPr sz="2000"/></a:lvl1pPr><a:lvl2pPr><a:defRPr sz="1500"/></a:lvl2pPr>"#,
            paragraph: r#"<a:pPr lvl="1"/>"#,
            ..Fixture::default()
        }
        .render(),
        15.0,
    );
}

#[test]
fn table_cell_without_default_text_style_or_other_style_keeps_eighteen_points() {
    assert_size(&Fixture::default().render(), 18.0);
}

#[test]
fn table_cell_without_default_text_style_still_inherits_master_other_style() {
    assert_size(
        &Fixture {
            other: FOURTEEN,
            ..Fixture::default()
        }
        .render(),
        14.0,
    );
}

#[test]
fn table_cell_inherits_presentation_default_paragraph_when_levels_are_absent() {
    assert_size(
        &Fixture {
            defaults: r#"<a:defPPr><a:defRPr sz="1300"/></a:defPPr>"#,
            ..Fixture::default()
        }
        .render(),
        13.0,
    );
}

#[test]
fn table_cell_inherits_font_color_tracking_and_paragraph_spacing_from_defaults() {
    let list = Fixture {
        defaults: r#"<a:lvl1pPr algn="r"><a:lnSpc><a:spcPts val="3000"/></a:lnSpc><a:defRPr sz="1400" spc="200"><a:latin typeface="Courier New"/><a:solidFill><a:srgbClr val="123456"/></a:solidFill></a:defRPr></a:lvl1pPr>"#,
        ..Fixture::default()
    }
    .render();
    assert_size(&list, 14.0);
    let inherited = line(&list);
    assert_eq!(inherited.runs[0].font_family, "Courier New");
    assert_eq!(inherited.runs[0].color, "#123456");
    assert!((inherited.height - 40.0).abs() < 0.001);
    let plain = Fixture {
        defaults: r#"<a:lvl1pPr><a:defRPr sz="1400"><a:latin typeface="Courier New"/></a:defRPr></a:lvl1pPr>"#,
        ..Fixture::default()
    }
    .render();
    assert!(inherited.width > line(&plain).width);
    assert!(inherited.x > line(&plain).x);
}

#[test]
fn table_cell_master_other_style_supplies_font_tracking_and_paragraph_spacing() {
    let list = Fixture {
        defaults: FOURTEEN,
        other: r#"<a:lvl1pPr algn="r"><a:lnSpc><a:spcPts val="3000"/></a:lnSpc><a:defRPr spc="200"><a:latin typeface="Courier New"/><a:solidFill><a:srgbClr val="654321"/></a:solidFill></a:defRPr></a:lvl1pPr>"#,
        ..Fixture::default()
    }
    .render();
    assert_size(&list, 18.0);
    assert_eq!(line(&list).runs[0].font_family, "Courier New");
    assert_eq!(line(&list).runs[0].color, "#654321");
    assert!((line(&list).height - 40.0).abs() < 0.001);
    let plain = Fixture {
        defaults: r#"<a:lvl1pPr><a:defRPr sz="1800"><a:latin typeface="Courier New"/></a:defRPr></a:lvl1pPr>"#,
        ..Fixture::default()
    }
    .render();
    assert!(line(&list).width > line(&plain).width);
    assert!(line(&list).x > line(&plain).x);
}

#[test]
fn table_style_color_overrides_master_and_presentation_colors() {
    let list = Fixture {
        defaults: r#"<a:lvl1pPr><a:defRPr sz="1400"><a:solidFill><a:srgbClr val="123456"/></a:solidFill></a:defRPr></a:lvl1pPr>"#,
        other: r#"<a:lvl1pPr><a:defRPr><a:solidFill><a:srgbClr val="654321"/></a:solidFill></a:defRPr></a:lvl1pPr>"#,
        table_style: r#"<a:srgbClr val="ABCDEF"/>"#,
        ..Fixture::default()
    }
    .render();
    assert_size(&list, 18.0);
    assert_eq!(line(&list).runs[0].color, "#ABCDEF");
}

#[test]
fn table_cell_normal_autofit_scales_the_inherited_size() {
    assert_size(
        &Fixture {
            defaults: FOURTEEN,
            other: FOURTEEN,
            autofit: r#"<a:normAutofit fontScale="50000"/>"#,
            ..Fixture::default()
        }
        .render(),
        7.0,
    );
}

#[test]
fn table_text_defaults_render_after_source_attach_without_changing_stored_state() {
    let bytes = Fixture {
        defaults: FOURTEEN,
        other: "<a:lvl1pPr/>",
        ..Fixture::default()
    }
    .bytes();
    let session = DeckSession::open(&bytes, 4202).unwrap();
    let stored = session.encode_state_as_update_v1();
    let stored_only = DeckSession::open_from_update(&stored, 4205).unwrap();
    assert_size(&render(&stored_only), 14.0);
    assert_eq!(stored_only.encode_state_as_update_v1(), stored);
    for client in [4203, 4204] {
        let attached = DeckSession::open_from_update_with_source(&stored, &bytes, client).unwrap();
        assert_size(&render(&attached), 18.0);
        assert_eq!(attached.encode_state_as_update_v1(), stored);
    }
}

fn cell_formatting(session: &DeckSession) -> pptx_edit::TextStyle {
    let snapshot = session.snapshot().unwrap();
    let slide = &snapshot.slides[0];
    let table = slide
        .shapes
        .iter()
        .find(|shape| shape.kind == pptx_edit::ShapeKind::GraphicFrame)
        .unwrap();
    let story = &table.text_stories[0];
    let formatted = session
        .story_formatting(&slide.id, &table.id, &story.id)
        .unwrap();
    formatted.paragraphs[0].runs[0].style.clone()
}

#[test]
fn table_cell_formatting_reads_the_size_and_weight_it_renders() {
    for (fixture, size, bold) in [
        (
            Fixture {
                other: r#"<a:lvl1pPr><a:defRPr sz="1400" b="1"/></a:lvl1pPr>"#,
                ..Fixture::default()
            },
            14.0,
            true,
        ),
        (
            Fixture {
                defaults: r#"<a:lvl1pPr><a:defRPr sz="1400" b="1"/></a:lvl1pPr>"#,
                other: r#"<a:lvl2pPr><a:defRPr sz="2000"/></a:lvl2pPr>"#,
                ..Fixture::default()
            },
            14.0,
            true,
        ),
        (
            Fixture {
                defaults: r#"<a:lvl1pPr><a:defRPr sz="1400" b="1"/></a:lvl1pPr>"#,
                other: "<a:lvl1pPr/>",
                ..Fixture::default()
            },
            18.0,
            false,
        ),
    ] {
        let session = DeckSession::open(&fixture.bytes(), 4201).unwrap();
        let list = render(&session);
        assert_size(&list, size);
        assert_eq!(line(&list).runs[0].bold, bold);
        let style = cell_formatting(&session);
        assert_eq!(style.font_size_pt, Some(f64::from(size)));
        assert_eq!(style.bold, Some(bold));
    }
}
