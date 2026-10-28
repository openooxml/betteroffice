#[allow(dead_code)]
#[path = "support/page_fixture.rs"]
mod fixture;

use docx_edit::{EngineSession, seed_from_docx};
use fixture::{FONT, p, r, region_request, with_body};
use serde_json::{Value, json};

#[test]
fn wrapped_cover_reserves_shared_default_header_with_own_geometry() {
    let font = docx_layout::register_measure_font(FONT).unwrap();
    for (index, (width, height, top)) in [(7200, 9000, 900), (12000, 8400, 1800)]
        .into_iter()
        .enumerate()
    {
        let cover = p(
            "00000001",
            &format!(
                "<w:pPr>{}</w:pPr>{}",
                section(7200, 9000, 900, 1200, ""),
                r("Cover")
            ),
        );
        let body = format!(
            concat!(
                "<w:sdt><w:sdtPr><w:id w:val=\"7\"/></w:sdtPr>",
                "<w:sdtContent>{cover}</w:sdtContent></w:sdt>{last}{section}"
            ),
            cover = cover,
            last = p("00000002", &r("Last")),
            section = section(
                width,
                height,
                top,
                1200,
                "<w:headerReference w:type=\"default\" r:id=\"rIdHeader2\"/>"
            ),
        );
        let document = format!(
            concat!(
                "<w:document ",
                "xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" ",
                "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" ",
                "xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\">",
                "<w:body>{body}</w:body></w:document>"
            ),
            body = body,
        );
        let header = concat!(
            "<w:hdr xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">",
            "<w:p><w:pPr><w:spacing w:before=\"0\" w:after=\"0\" ",
            "w:line=\"1200\" w:lineRule=\"exact\"/></w:pPr>",
            "<w:r><w:t>Shared header</w:t></w:r></w:p></w:hdr>"
        );
        let mut parts = fixture::page_parts(false);
        for (name, bytes) in &mut parts {
            match name.as_str() {
                "word/document.xml" => *bytes = document.as_bytes().to_vec(),
                "word/header2.xml" => *bytes = header.as_bytes().to_vec(),
                _ => {}
            }
        }
        let bytes = ooxml_opc::rezip_parts(&parts).unwrap();
        let engine = EngineSession::new(75211 + index as u64);
        seed_from_docx(engine.doc(), &bytes).unwrap();
        let request = region_request(&engine, &bytes, font);
        assert_eq!(request["regions"]["sections"].as_array().unwrap().len(), 1);
        let output: Value = serde_json::from_str(
            &engine
                .layout_document_with_regions_json(&request.to_string())
                .unwrap(),
        )
        .unwrap();
        let pages = output["layout"]["pages"].as_array().unwrap();
        assert_eq!(pages.len(), 2, "case {index}");
        assert_eq!(pages[0]["size"], json!({"w": 480.0, "h": 600.0}));
        let body_margins = |page: &Value| {
            page.get("bodyMargins")
                .filter(|margins| !margins.is_null())
                .unwrap_or(&page["margins"])
                .clone()
        };
        assert_eq!(body_margins(&pages[0])["top"], 100.0, "case {index}");
        assert_eq!(pages[0]["fragments"][0]["y"], 100.0, "case {index}");
        assert_eq!(
            body_margins(&pages[1])["top"],
            (top as f64 / 15.0).max(100.0),
            "case {index}"
        );
    }
}

#[test]
fn consecutive_wrapped_sections_inherit_authored_margins_after_a_tall_header() {
    let first = p(
        "00000001",
        &format!(
            "<w:pPr>{}</w:pPr>{}",
            section(
                9000,
                9000,
                1440,
                1440,
                "<w:headerReference w:type=\"default\" r:id=\"rIdHeader2\"/>"
            ),
            r("First")
        ),
    );
    let wrapped = |control: u32, id: &str, text: &str| {
        format!(
            concat!(
                "<w:sdt><w:sdtPr><w:id w:val=\"{control}\"/></w:sdtPr><w:sdtContent>",
                "{paragraph}</w:sdtContent></w:sdt>"
            ),
            control = control,
            paragraph = p(
                id,
                &format!(
                    concat!(
                        "<w:pPr><w:sectPr><w:type w:val=\"nextPage\"/>",
                        "</w:sectPr></w:pPr>{text}"
                    ),
                    text = r(text),
                )
            ),
        )
    };
    let body = format!(
        "{first}{wrapped_first}{wrapped_second}{last}{section}",
        first = first,
        wrapped_first = wrapped(7, "00000002", "Wrapped first"),
        wrapped_second = wrapped(8, "00000003", "Wrapped second"),
        last = p("00000004", &r("Last")),
        section = section(
            9000,
            9000,
            1440,
            1440,
            "<w:headerReference w:type=\"default\" r:id=\"rIdHeader1\"/>"
        ),
    );
    let mut parts = fixture::page_parts(false);
    for (name, bytes) in &mut parts {
        let xml = match name.as_str() {
            "word/document.xml" => Some(format!(
                concat!(
                    "<w:document ",
                    "xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" ",
                    "xmlns:r=\"",
                    "http://schemas.openxmlformats.org/officeDocument/2006/relationships\" ",
                    "xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\">",
                    "<w:body>{body}</w:body></w:document>"
                ),
                body = body,
            )),
            "word/header2.xml" => Some(
                concat!(
                    "<w:hdr xmlns:w=\"",
                    "http://schemas.openxmlformats.org/wordprocessingml/2006/main\">",
                    "<w:p><w:pPr><w:spacing w:before=\"0\" w:after=\"0\" ",
                    "w:line=\"6000\" w:lineRule=\"exact\"/></w:pPr>",
                    "<w:r><w:t>Tall header</w:t></w:r></w:p></w:hdr>"
                )
                .to_owned(),
            ),
            "word/header1.xml" => Some(
                concat!(
                    "<w:hdr xmlns:w=\"",
                    "http://schemas.openxmlformats.org/wordprocessingml/2006/main\">",
                    "<w:p/></w:hdr>"
                )
                .to_owned(),
            ),
            _ => None,
        };
        if let Some(xml) = xml {
            *bytes = xml.into_bytes();
        }
    }
    let bytes = ooxml_opc::rezip_parts(&parts).unwrap();
    let font = docx_layout::register_measure_font(FONT).unwrap();
    let engine = EngineSession::new(75213);
    seed_from_docx(engine.doc(), &bytes).unwrap();
    let mut request = region_request(&engine, &bytes, font);
    let package = docx_parse::parse_docx_s9_wire(&bytes, Default::default())
        .unwrap()
        .document
        .package;
    let sections = request["regions"]["sections"].as_array_mut().unwrap();
    assert_eq!(sections.len(), 2);
    sections.last_mut().unwrap()["properties"] = json!(package.document.final_section_properties);
    let output: Value = serde_json::from_str(
        &engine
            .layout_document_with_regions_json(&request.to_string())
            .unwrap(),
    )
    .unwrap();
    let pages = output["layout"]["pages"].as_array().unwrap();
    assert_eq!(pages.len(), 4);
    assert_eq!(pages[0]["margins"]["top"], 420.0);
    for page in &pages[1..] {
        let margins = page
            .get("bodyMargins")
            .filter(|margins| !margins.is_null())
            .unwrap_or(&page["margins"]);
        assert_eq!(margins["top"], 96.0);
        assert_eq!(
            page["size"]["h"].as_f64().unwrap()
                - margins["top"].as_f64().unwrap()
                - margins["bottom"].as_f64().unwrap(),
            408.0
        );
    }
}

fn short_page_floating_footer_output(last: &str, client_id: u64) -> Value {
    let lines = (0..22)
        .map(|index| {
            let line = r(&format!("Cover line {index}"));
            if index == 21 {
                line
            } else {
                format!("{line}<w:r><w:br/></w:r>")
            }
        })
        .collect::<String>();
    let cover = p(
        "00000001",
        &format!(
            concat!(
                "<w:pPr><w:spacing w:before=\"0\" w:after=\"0\" ",
                "w:line=\"300\" w:lineRule=\"exact\"/>{section}</w:pPr>{lines}"
            ),
            section = section(9000, 9000, 900, 1200, ""),
            lines = lines,
        ),
    );
    let body = format!(
        concat!(
            "<w:sdt><w:sdtPr><w:id w:val=\"7\"/></w:sdtPr>",
            "<w:sdtContent>{cover}</w:sdtContent></w:sdt>{last}{section}"
        ),
        cover = cover,
        last = last,
        section = section(
            9000,
            15000,
            900,
            1200,
            "<w:footerReference w:type=\"default\" r:id=\"rIdFooter1\"/>"
        ),
    );
    let footer = concat!(
        "<w:ftr ",
        "xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" ",
        "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" ",
        "xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\" ",
        "xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" ",
        "xmlns:pic=\"http://schemas.openxmlformats.org/drawingml/2006/picture\">",
        "<w:p><w:r><w:drawing><wp:anchor distT=\"0\" distB=\"0\" distL=\"0\" ",
        "distR=\"0\" simplePos=\"0\" relativeHeight=\"1\" behindDoc=\"0\" ",
        "locked=\"0\" layoutInCell=\"1\" allowOverlap=\"1\">",
        "<wp:simplePos x=\"0\" y=\"0\"/>",
        "<wp:positionH relativeFrom=\"margin\"><wp:posOffset>0</wp:posOffset></wp:positionH>",
        "<wp:positionV relativeFrom=\"margin\"><wp:align>bottom</wp:align></wp:positionV>",
        "<wp:extent cx=\"1143000\" cy=\"1143000\"/>",
        "<wp:effectExtent l=\"0\" t=\"0\" r=\"0\" b=\"0\"/>",
        "<wp:wrapTopAndBottom/><wp:docPr id=\"1\" name=\"Footer image\"/>",
        "<wp:cNvGraphicFramePr/><a:graphic><a:graphicData ",
        "uri=\"http://schemas.openxmlformats.org/drawingml/2006/picture\">",
        "<pic:pic><pic:nvPicPr><pic:cNvPr id=\"1\" name=\"Footer image\"/>",
        "<pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed=\"rIdImage\"/>",
        "<a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm>",
        "<a:off x=\"0\" y=\"0\"/><a:ext cx=\"1143000\" cy=\"1143000\"/></a:xfrm>",
        "<a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom></pic:spPr></pic:pic>",
        "</a:graphicData></a:graphic></wp:anchor></w:drawing></w:r></w:p></w:ftr>"
    );
    let mut parts = fixture::page_parts(false);
    for (name, bytes) in &mut parts {
        match name.as_str() {
            "word/document.xml" => {
                *bytes = format!(
                    concat!(
                        "<w:document ",
                        "xmlns:w=\"",
                        "http://schemas.openxmlformats.org/wordprocessingml/2006/main\" ",
                        "xmlns:r=\"",
                        "http://schemas.openxmlformats.org/officeDocument/2006/relationships\" ",
                        "xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\">",
                        "<w:body>{body}</w:body></w:document>"
                    ),
                    body = body,
                )
                .into_bytes();
            }
            "word/footer1.xml" => *bytes = footer.as_bytes().to_vec(),
            "[Content_Types].xml" => {
                *bytes = String::from_utf8(std::mem::take(bytes))
                    .unwrap()
                    .replace(
                        "</Types>",
                        "<Default Extension=\"png\" ContentType=\"image/png\"/></Types>",
                    )
                    .into_bytes();
            }
            _ => {}
        }
    }
    parts.push((
        "word/_rels/footer1.xml.rels".into(),
        concat!(
            "<Relationships ",
            "xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">",
            "<Relationship Id=\"rIdImage\" ",
            "Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/image\" ",
            "Target=\"media/footer.png\"/></Relationships>"
        )
        .as_bytes()
        .to_vec(),
    ));
    parts.push((
        "word/media/footer.png".into(),
        vec![
            137, 80, 78, 71, 13, 10, 26, 10, 0, 0, 0, 13, 73, 72, 68, 82, 0, 0, 0, 1, 0, 0, 0, 1,
            8, 4, 0, 0, 0, 181, 28, 12, 2, 0, 0, 0, 11, 73, 68, 65, 84, 120, 218, 99, 252, 255,
            31, 0, 3, 3, 2, 0, 239, 163, 133, 45, 0, 0, 0, 0, 73, 69, 78, 68, 174, 66, 96, 130,
        ],
    ));
    let bytes = ooxml_opc::rezip_parts(&parts).unwrap();
    let font = docx_layout::register_measure_font(FONT).unwrap();
    let engine = EngineSession::new(client_id);
    seed_from_docx(engine.doc(), &bytes).unwrap();
    let request = region_request(&engine, &bytes, font);
    serde_json::from_str(
        &engine
            .layout_document_with_regions_json(&request.to_string())
            .unwrap(),
    )
    .unwrap()
}

#[test]
fn wrapped_short_page_reserves_its_floating_footer_image_at_the_rendered_position() {
    let output = short_page_floating_footer_output(&p("00000002", &r("Last")), 75214);
    let pages = output["layout"]["pages"].as_array().unwrap();
    assert_eq!(pages.len(), 3);
    for page in &pages[..2] {
        assert_eq!(page["size"]["h"], 600.0);
        assert_eq!(page["margins"]["bottom"], 80.0);
        assert_eq!(page["bodyMargins"]["bottom"], 200.0);
        for fragment in page["fragments"].as_array().unwrap() {
            let bottom = fragment["y"].as_f64().unwrap() + fragment["height"].as_f64().unwrap();
            assert!(bottom <= 400.0, "body overlaps the footer: {fragment}");
        }
    }
    assert_eq!(pages[2]["size"]["h"], 1000.0);
    assert_eq!(pages[2]["bodyMargins"]["bottom"], 200.0);
}

#[test]
fn first_tall_page_after_wrapped_section_refreshes_its_footer_exclusion() {
    let lines = (0..43)
        .map(|index| {
            let line = r(&format!("Following line {index}"));
            if index == 42 {
                line
            } else {
                format!("{line}<w:r><w:br/></w:r>")
            }
        })
        .collect::<String>();
    let last = p(
        "00000002",
        &format!(
            concat!(
                "<w:pPr><w:spacing w:before=\"0\" w:after=\"0\" ",
                "w:line=\"300\" w:lineRule=\"exact\"/></w:pPr>{lines}"
            ),
            lines = lines,
        ),
    );
    let output = short_page_floating_footer_output(&last, 75215);
    let pages = output["layout"]["pages"].as_array().unwrap();
    assert_eq!(pages.len(), 4);
    let first_tall = &pages[2];
    assert_eq!(first_tall["size"]["h"], 1000.0);
    assert_eq!(first_tall["sectionPageIndex"], 2);
    assert_eq!(first_tall["margins"]["bottom"], 80.0);
    assert_eq!(first_tall["bodyMargins"]["bottom"], 200.0);
    let fragments = first_tall["fragments"].as_array().unwrap();
    let mut body_bottom = 0.0_f64;
    for fragment in fragments {
        let bottom = fragment["y"].as_f64().unwrap() + fragment["height"].as_f64().unwrap();
        assert!(bottom <= 800.0, "body overlaps the footer: {fragment}");
        body_bottom = body_bottom.max(bottom);
    }
    assert_eq!(body_bottom, 800.0);
}

fn section(width: u32, height: u32, top: u32, bottom: u32, extra: &str) -> String {
    format!(
        concat!(
            "<w:sectPr><w:type w:val=\"nextPage\"/>",
            "<w:pgSz w:w=\"{width}\" w:h=\"{height}\"/>",
            "<w:pgMar w:top=\"{top}\" w:right=\"720\" w:bottom=\"{bottom}\" ",
            "w:left=\"720\" w:header=\"300\" w:footer=\"300\"/>",
            "{extra}</w:sectPr>"
        ),
        width = width,
        height = height,
        top = top,
        bottom = bottom,
        extra = extra,
    )
}

#[test]
fn resident_sdt_cover_section_preserves_each_sections_geometry() {
    let cover = p(
        "00000001",
        &format!(
            "<w:pPr>{}</w:pPr>{}",
            section(7200, 9000, 900, 1200, ""),
            r("Cover")
        ),
    );
    let middle = p(
        "00000002",
        &format!(
            "<w:pPr>{}</w:pPr>{}",
            section(
                12000,
                8400,
                1440,
                1080,
                concat!(
                    "<w:headerReference w:type=\"default\" r:id=\"rIdHeader2\"/>",
                    "<w:docGrid w:type=\"lines\" w:linePitch=\"300\"/>"
                )
            ),
            r("Middle")
        ),
    );
    let body = format!(
        concat!(
            "<w:sdt><w:sdtPr><w:id w:val=\"7\"/></w:sdtPr>",
            "<w:sdtContent>{cover}</w:sdtContent></w:sdt>{middle}{last}{section}"
        ),
        cover = cover,
        middle = middle,
        last = p("00000003", &r("Last")),
        section = section(
            7800,
            10200,
            1800,
            1500,
            "<w:docGrid w:type=\"lines\" w:linePitch=\"450\"/>"
        ),
    );
    let bytes = with_body(&body);
    let font = docx_layout::register_measure_font(FONT).unwrap();
    let engine = EngineSession::new(75210);
    seed_from_docx(engine.doc(), &bytes).unwrap();
    let mut request = region_request(&engine, &bytes, font);
    let package = docx_parse::parse_docx_s9_wire(&bytes, Default::default())
        .unwrap()
        .document
        .package;
    let sections = request["regions"]["sections"].as_array_mut().unwrap();
    assert_eq!(sections.len(), 2);
    sections.last_mut().unwrap()["properties"] = json!(package.document.final_section_properties);
    let request = request.to_string();
    for output in [
        engine.layout_document_with_regions_json(&request).unwrap(),
        engine.layout_document_with_regions_json(&request).unwrap(),
    ] {
        let output: Value = serde_json::from_str(&output).unwrap();
        let pages = output["layout"]["pages"].as_array().unwrap();
        assert_eq!(pages.len(), 3);
        for (index, (width, height, top, bottom)) in [
            (480.0, 600.0, 60.0, 80.0),
            (800.0, 560.0, 96.0, 72.0),
            (520.0, 680.0, 120.0, 100.0),
        ]
        .into_iter()
        .enumerate()
        {
            assert_eq!(pages[index]["size"], json!({"w": width, "h": height}));
            assert_eq!(pages[index]["margins"]["top"], top);
            assert_eq!(pages[index]["margins"]["bottom"], bottom);
        }
        let measured = output["measured"].as_array().unwrap();
        let paragraphs: Vec<_> = measured
            .iter()
            .filter(|measured| measured["block"]["kind"] == "paragraph")
            .collect();
        assert_eq!(paragraphs[1]["block"]["attrs"]["docGridPitchPx"], 20.0);
        assert_eq!(paragraphs[2]["block"]["attrs"]["docGridPitchPx"], 30.0);
    }
}
