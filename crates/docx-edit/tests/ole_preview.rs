use docx_edit::bridge::{RenderEnv, yrs_doc_to_layout_blocks};
use docx_edit::structured::{BlockKind, ExportOptions, InlineKind, RevisionView};
use docx_edit::{EditingDoc, EngineSession, seed_from_docx};
use docx_layout::types::{LayoutBlock, Run};
use serde_json::{Value, json};

const REL: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const NS: &str = r#"xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:w10="urn:schemas-microsoft-com:office:word""#;
const INLINE_OLE: &str = r#"<w:object><v:shape style="width:72pt;height:36pt"><v:imagedata r:id="rIdImage"/></v:shape><o:OLEObject Type="Embed" ProgID="Equation.DSMT4" r:id="rIdOle"/></w:object>"#;
const FLOATING_OLE: &str = r#"<w:object><v:shape style="position:absolute;margin-left:0pt;margin-top:42.75pt;width:505pt;height:133pt;mso-position-horizontal-relative:text;mso-position-vertical-relative:text"><v:imagedata r:id="rIdImage"/><w10:wrap type="square"/></v:shape><o:OLEObject Type="Embed" ProgID="Word.Picture.8" r:id="rIdOle"/></w:object>"#;

fn ole_document(header: bool, inline_control: bool) -> Vec<u8> {
    let body = if header {
        r#"<w:p w14:paraId="11111111"><w:r><w:t>First body line below the header object</w:t></w:r></w:p><w:p w14:paraId="22222222"><w:r><w:br w:type="page"/><w:t>Second page body</w:t></w:r></w:p>"#.to_owned()
    } else {
        let run = format!(r#"<w:r><w:t>A</w:t>{INLINE_OLE}<w:t>B</w:t></w:r>"#);
        let content = if inline_control {
            format!(
                r#"<w:sdt><w:sdtPr><w:id w:val="17"/></w:sdtPr><w:sdtContent>{run}</w:sdtContent></w:sdt>"#
            )
        } else {
            run
        };
        format!(r#"<w:p w14:paraId="11111111">{content}</w:p>"#)
    };
    let header_ref = if header {
        r#"<w:headerReference w:type="first" r:id="rIdHeader"/><w:titlePg/>"#
    } else {
        ""
    };
    let header_type = if header {
        r#"<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>"#
    } else {
        ""
    };
    let header_rel = if header {
        format!(r#"<Relationship Id="rIdHeader" Type="{REL}/header" Target="header1.xml"/>"#)
    } else {
        String::new()
    };
    let mut parts = vec![
        (
            "[Content_Types].xml".to_owned(),
            format!(r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>{header_type}</Types>"#).into_bytes(),
        ),
        (
            "_rels/.rels".to_owned(),
            format!(r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdDoc" Type="{REL}/officeDocument" Target="word/document.xml"/></Relationships>"#).into_bytes(),
        ),
        (
            "word/_rels/document.xml.rels".to_owned(),
            format!(r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdImage" Type="{REL}/image" Target="media/image.png"/>{header_rel}</Relationships>"#).into_bytes(),
        ),
        (
            "word/document.xml".to_owned(),
            format!(r#"<w:document {NS}><w:body>{body}<w:sectPr>{header_ref}<w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1120" w:right="900" w:bottom="1440" w:left="900" w:header="240" w:footer="720"/></w:sectPr></w:body></w:document>"#).into_bytes(),
        ),
        (
            "word/media/image.png".to_owned(),
            b"\x89PNG\r\n\x1a\n".to_vec(),
        ),
    ];
    if header {
        parts.push((
            "word/header1.xml".to_owned(),
            format!(
                r#"<w:hdr {NS}><w:p w14:paraId="33333333"><w:r>{FLOATING_OLE}</w:r></w:p></w:hdr>"#
            )
            .into_bytes(),
        ));
        parts.push((
            "word/_rels/header1.xml.rels".to_owned(),
            format!(r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdImage" Type="{REL}/image" Target="media/image.png"/></Relationships>"#).into_bytes(),
        ));
    }
    ooxml_opc::rezip_parts(&parts).unwrap()
}

#[test]
fn ole_preview_lowers_as_an_image_without_expanding_the_atom() {
    let doc = EditingDoc::new(74312);
    seed_from_docx(&doc, &ole_document(false, false)).unwrap();
    let blocks = yrs_doc_to_layout_blocks(&doc, "body", &RenderEnv::default()).unwrap();
    let LayoutBlock::Paragraph(paragraph) = &blocks[0] else {
        panic!("expected body paragraph");
    };
    assert_eq!(paragraph.runs.len(), 3);
    let Run::Image(image) = &paragraph.runs[1] else {
        panic!("expected OLE preview image");
    };
    assert_eq!((image.width, image.height), (96.0, 48.0));
    assert_eq!((image.pm_start, image.pm_end), (Some(2.0), Some(3.0)));
    let Run::Text(tail) = &paragraph.runs[2] else {
        panic!("expected text after the OLE atom");
    };
    assert_eq!(tail.text, "B");
    assert_eq!((tail.pm_start, tail.pm_end), (Some(3.0), Some(4.0)));
}

#[test]
fn inline_control_ole_preview_lowers_as_an_image_atom() {
    let doc = EditingDoc::new(74316);
    seed_from_docx(&doc, &ole_document(false, true)).unwrap();
    let blocks = yrs_doc_to_layout_blocks(&doc, "body", &RenderEnv::default()).unwrap();
    let LayoutBlock::Paragraph(paragraph) = &blocks[0] else {
        panic!("expected body paragraph");
    };
    let image = paragraph
        .runs
        .iter()
        .find_map(|run| match run {
            Run::Image(image) => Some(image),
            _ => None,
        })
        .expect("inline control must draw its OLE preview");
    assert_eq!((image.width, image.height), (96.0, 48.0));
    assert_eq!((image.pm_start, image.pm_end), (Some(3.0), Some(4.0)));
    let Run::Text(tail) = paragraph.runs.last().unwrap() else {
        panic!("expected text after the OLE atom");
    };
    assert_eq!(tail.text, "B");
    assert_eq!((tail.pm_start, tail.pm_end), (Some(4.0), Some(5.0)));
}

#[test]
fn structured_export_describes_the_ole_preview_as_one_image() {
    let doc = EditingDoc::new(74315);
    seed_from_docx(&doc, &ole_document(false, false)).unwrap();
    let options = ExportOptions::new(RevisionView::Accepted);
    let exported = doc.export_structured(&options).unwrap();
    let BlockKind::Paragraph { paragraph } = &exported.content.stories[0].blocks[0].content else {
        panic!("expected body paragraph");
    };
    assert_eq!(paragraph.inlines.len(), 3);
    let InlineKind::Image {
        relationship_id,
        part,
        ..
    } = &paragraph.inlines[1].content
    else {
        panic!("expected OLE preview image");
    };
    assert_eq!(relationship_id.as_deref(), Some("rIdImage"));
    assert_eq!(part.as_deref(), Some("word/media/image.png"));
}

fn vml_group_document(inline_control: bool) -> Vec<u8> {
    let group = r#"<w:pict><v:group coordsize="1000,1000" style="width:72pt;height:36pt"><v:shape style="left:0;top:0;width:500;height:500"><v:imagedata r:id="rIdImage"/></v:shape><v:shape style="left:500;top:500;width:500;height:500"><v:imagedata r:id="rIdImage"/></v:shape></v:group></w:pict>"#;
    let mut parts = ooxml_opc::unzip_parts(&ole_document(false, inline_control)).unwrap();
    let (_, xml) = parts
        .iter_mut()
        .find(|(path, _)| path == "word/document.xml")
        .unwrap();
    *xml = std::str::from_utf8(xml)
        .unwrap()
        .replace(INLINE_OLE, group)
        .into_bytes();
    ooxml_opc::rezip_parts(&parts).unwrap()
}

#[test]
fn vml_group_previews_share_one_atom_in_body_and_inline_control() {
    for inline_control in [false, true] {
        let doc = EditingDoc::new(74317);
        seed_from_docx(&doc, &vml_group_document(inline_control)).unwrap();
        let blocks = yrs_doc_to_layout_blocks(&doc, "body", &RenderEnv::default()).unwrap();
        let LayoutBlock::Paragraph(paragraph) = &blocks[0] else {
            panic!("expected body paragraph");
        };
        let images: Vec<_> = paragraph
            .runs
            .iter()
            .filter_map(|run| match run {
                Run::Image(image) => Some(image),
                _ => None,
            })
            .collect();
        assert_eq!(images.len(), 2);
        let start = if inline_control { 3.0 } else { 2.0 };
        for image in images {
            assert_eq!((image.width, image.height), (48.0, 24.0));
            assert_eq!(
                (image.pm_start, image.pm_end),
                (Some(start), Some(start + 1.0))
            );
        }
        let Run::Text(tail) = paragraph.runs.last().unwrap() else {
            panic!("expected text after the group");
        };
        assert_eq!(tail.text, "B");
        assert_eq!(
            (tail.pm_start, tail.pm_end),
            (Some(start + 1.0), Some(start + 2.0))
        );
    }
}

#[test]
fn negative_vml_group_z_index_keeps_its_position_behind_text() {
    let mut parts = ooxml_opc::unzip_parts(&vml_group_document(false)).unwrap();
    let (_, xml) = parts
        .iter_mut()
        .find(|(path, _)| path == "word/document.xml")
        .unwrap();
    *xml = std::str::from_utf8(xml)
        .unwrap()
        .replace(
            "style=\"width:72pt;height:36pt\"",
            "style=\"position:absolute;z-index:-10;margin-left:12pt;width:72pt;height:36pt\"",
        )
        .into_bytes();
    let doc = EditingDoc::new(74319);
    seed_from_docx(&doc, &ooxml_opc::rezip_parts(&parts).unwrap()).unwrap();
    let blocks = yrs_doc_to_layout_blocks(&doc, "body", &RenderEnv::default()).unwrap();
    let LayoutBlock::Paragraph(paragraph) = &blocks[0] else {
        panic!("expected body paragraph");
    };
    let images: Vec<_> = paragraph
        .runs
        .iter()
        .filter_map(|run| match run {
            Run::Image(image) => Some(image),
            _ => None,
        })
        .collect();
    assert_eq!(images.len(), 2);
    for image in images {
        assert_eq!(image.wrap_type.as_deref(), Some("behind"));
        let position = image.position.as_ref().unwrap();
        assert_eq!(position.behind_doc, Some(true));
        assert_eq!(position.relative_height, Some(0));
    }
}

#[test]
fn structured_export_keeps_a_vml_group_image_preview() {
    let doc = EditingDoc::new(74318);
    seed_from_docx(&doc, &vml_group_document(false)).unwrap();
    let exported = doc
        .export_structured(&ExportOptions::new(RevisionView::Accepted))
        .unwrap();
    let BlockKind::Paragraph { paragraph } = &exported.content.stories[0].blocks[0].content else {
        panic!("expected body paragraph");
    };
    assert_eq!(paragraph.inlines.len(), 3);
    assert!(
        matches!(&paragraph.inlines[1].content, InlineKind::Image { relationship_id, .. }
        if relationship_id.as_deref() == Some("rIdImage"))
    );
}

#[test]
fn first_page_floating_ole_header_keeps_body_below_its_preview() {
    let font = docx_layout::register_measure_font_bytes(include_bytes!(
        "../../ooxml-text/tests/fonts/LiberationSans-Regular.ttf"
    ))
    .unwrap();
    let engine = EngineSession::new(74313);
    seed_from_docx(engine.doc(), &ole_document(true, false)).unwrap();
    let blocks =
        yrs_doc_to_layout_blocks(engine.doc(), "hf:rIdHeader", &RenderEnv::default()).unwrap();
    let LayoutBlock::Paragraph(paragraph) = &blocks[0] else {
        panic!("expected header paragraph");
    };
    let Run::Image(preview) = &paragraph.runs[0] else {
        panic!("expected floating OLE preview");
    };
    assert_eq!(preview.wrap_type.as_deref(), Some("square"));
    let position = preview.position.as_ref().unwrap();
    let vertical = position.vertical.as_ref().unwrap();
    assert_eq!(vertical.relative_to.as_deref(), Some("paragraph"));
    assert_eq!(vertical.pos_offset, Some(542_925.0));
    let request = json!({
        "bodyStory": "body",
        "regions": {"sections": [{
            "sectionId": "main",
            "pageSize": {"w": 11906.0 / 15.0, "h": 16838.0 / 15.0},
            "margins": {"top": 1120.0 / 15.0, "right": 60.0, "bottom": 96.0, "left": 60.0, "header": 16.0, "footer": 48.0},
            "headerDistance": 16.0,
            "titlePg": true,
            "headerFooterRefs": {"headerFirst": "rIdHeader"}
        }]},
        "measurement": {
            "fontChains": {"liberation sans|0|0": [font]},
            "defaults": {"fontSize": 11, "fontFamily": "Liberation Sans"},
            "authoritativeShaping": true
        },
        "renderEnv": {}
    });
    let layout = engine
        .layout_document_with_regions_json(&request.to_string())
        .unwrap();
    let display: Value =
        serde_json::from_str(&engine.build_display_list_json(&layout).unwrap()).unwrap();
    let pages = display["pages"].as_array().unwrap();
    assert_eq!(pages.len(), 2);
    let header = pages[0]["header"]["primitives"].as_array().unwrap();
    let image = header
        .iter()
        .find(|primitive| primitive["kind"] == "image")
        .expect("first-page header must draw its OLE preview");
    assert!((image["w"].as_f64().unwrap() - 505.0 * 96.0 / 72.0).abs() < 0.01);
    assert!((image["h"].as_f64().unwrap() - 133.0 * 96.0 / 72.0).abs() < 0.01);
    let bottom = image["y"].as_f64().unwrap() + image["h"].as_f64().unwrap();
    let layout: Value = serde_json::from_str(&layout).unwrap();
    let first_body = &layout["layout"]["pages"][0]["fragments"][0];
    assert!(first_body["y"].as_f64().unwrap() >= bottom - 0.01);
    assert!(pages[1]["header"].is_null());
    assert!(
        layout["layout"]["pages"][1]["fragments"][0]["y"]
            .as_f64()
            .unwrap()
            < bottom
    );
}
