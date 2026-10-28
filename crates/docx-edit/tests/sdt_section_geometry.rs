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
        assert_eq!(request["regions"]["sections"].as_array().unwrap().len(), 2);
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
    assert_eq!(sections.len(), 3);
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
    render_section_parts(floating_footer_parts(&body), client_id)
}

fn floating_footer_parts(body: &str) -> Vec<(String, Vec<u8>)> {
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
            8, 4, 0, 0, 0, 181, 28, 12, 2, 0, 0, 0, 11, 73, 68, 65, 84, 120, 218, 99, 252, 255, 31,
            0, 3, 3, 2, 0, 239, 163, 133, 45, 0, 0, 0, 0, 73, 69, 78, 68, 174, 66, 96, 130,
        ],
    ));
    parts
}

fn render_section_parts(parts: Vec<(String, Vec<u8>)>, client_id: u64) -> Value {
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

#[test]
fn final_section_after_two_wrapped_sections_keeps_its_limits_after_a_split_table() {
    let side_margins = |xml: String| {
        xml.replace("w:right=\"720\"", "w:right=\"1440\"")
            .replace("w:left=\"720\"", "w:left=\"1440\"")
    };
    let wrapped = |id, properties: String, text: &str| {
        format!(
            concat!(
                "<w:sdt><w:sdtPr><w:id w:val=\"{id}\"/></w:sdtPr><w:sdtContent>",
                "{paragraph}</w:sdtContent></w:sdt>"
            ),
            id = id,
            paragraph = p(
                &format!("{id:08X}"),
                &format!("<w:pPr>{properties}</w:pPr>{}", r(text))
            ),
        )
    };
    let cover = wrapped(
        7_u32,
        side_margins(section(12240, 15840, 3600, 1440, ""))
            .replace("w:right=\"1440\"", "w:right=\"450\""),
        "Cover",
    );
    let contents = wrapped(
        8_u32,
        side_margins(section(12240, 15840, 1170, 1440, "")),
        "Contents",
    );
    let rows = (0..20)
        .map(|index| {
            format!(
                concat!(
                    "<w:tr><w:trPr><w:cantSplit/><w:trHeight w:val=\"1500\" ",
                    "w:hRule=\"exact\"/></w:trPr><w:tc><w:tcPr>",
                    "<w:tcW w:w=\"9360\" w:type=\"dxa\"/></w:tcPr>{paragraph}</w:tc></w:tr>"
                ),
                paragraph = p(&format!("{:08X}", 0x10000 + index), &r("Table row")),
            )
        })
        .collect::<String>();
    let table = format!(
        concat!(
            "<w:tbl><w:tblPr><w:tblW w:w=\"9360\" w:type=\"dxa\"/>",
            "<w:tblLayout w:type=\"fixed\"/></w:tblPr>",
            "<w:tblGrid><w:gridCol w:w=\"9360\"/></w:tblGrid>{rows}</w:tbl>"
        ),
        rows = rows,
    );
    let paragraph = |id, content: &str, keep| {
        p(
            &format!("{id:08X}"),
            &format!(
                concat!(
                    "<w:pPr>{keep}<w:spacing w:before=\"0\" w:after=\"0\" ",
                    "w:line=\"300\" w:lineRule=\"exact\"/></w:pPr>{content}"
                ),
                keep = if keep { "<w:keepNext/>" } else { "" },
                content = content,
            ),
        )
    };
    let tails = (0..18)
        .map(|index| paragraph(0x20000 + index, &r("Following body text"), false))
        .collect::<String>();
    let body = format!(
        "{cover}{contents}{table}{anchor}{caption}{tails}{section}",
        cover = cover,
        contents = contents,
        table = table,
        anchor = paragraph(0xF001, "BODY_FLOAT_IMAGE", true),
        caption = paragraph(0xF002, &r("Caption after the table"), false),
        tails = tails,
        section = side_margins(section(12240, 15840, 1440, 1440, "")),
    );
    let mut parts = floating_footer_parts(&body);
    let (_, footer_bytes) = parts
        .iter()
        .find(|(name, _)| name == "word/footer1.xml")
        .unwrap();
    let footer = std::str::from_utf8(footer_bytes).unwrap();
    let drawing = footer
        .split_once("<w:drawing>")
        .unwrap()
        .1
        .split_once("</w:drawing>")
        .unwrap()
        .0
        .replace("<wp:align>bottom</wp:align>", "<wp:align>top</wp:align>")
        .replace(
            "<wp:wrapTopAndBottom/>",
            "<wp:wrapSquare wrapText=\"bothSides\"/>",
        )
        .replace("cx=\"1143000\"", "cx=\"5852160\"")
        .replace("cy=\"1143000\"", "cy=\"4876800\"")
        .replace("distL=\"0\"", "distL=\"114300\"")
        .replace("distR=\"0\"", "distR=\"114300\"");
    let image = format!(
        concat!(
            "<w:r ",
            "xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\" ",
            "xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" ",
            "xmlns:pic=\"http://schemas.openxmlformats.org/drawingml/2006/picture\">",
            "<w:drawing>{drawing}</w:drawing></w:r>"
        ),
        drawing = drawing,
    );
    for (name, bytes) in &mut parts {
        let xml = std::str::from_utf8(bytes).ok();
        let replacement = match name.as_str() {
            "word/document.xml" => Some(xml.unwrap().replace("BODY_FLOAT_IMAGE", &image)),
            "word/_rels/document.xml.rels" => Some(xml.unwrap().replace(
                "</Relationships>",
                concat!(
                    "<Relationship Id=\"rIdImage\" ",
                    "Type=\"",
                    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image\" ",
                    "Target=\"media/footer.png\"/></Relationships>"
                ),
            )),
            _ => None,
        };
        if let Some(replacement) = replacement {
            *bytes = replacement.into_bytes();
        }
    }
    let output = render_section_parts(parts, 75216);
    let pages = output["layout"]["pages"].as_array().unwrap();
    assert_eq!(pages.len(), 5);
    for page in &pages[2..] {
        assert_eq!(page["size"], json!({"w": 816.0, "h": 1056.0}));
        let margins = page
            .get("bodyMargins")
            .filter(|margins| !margins.is_null())
            .unwrap_or(&page["margins"]);
        assert_eq!(margins["top"], 96.0);
        assert_eq!(margins["bottom"], 96.0);
        for fragment in page["fragments"].as_array().unwrap() {
            let bottom = fragment["y"].as_f64().unwrap() + fragment["height"].as_f64().unwrap();
            assert!(
                bottom <= 960.0,
                "body exceeds final-section limits: {fragment}"
            );
        }
    }
    assert_eq!(pages[4]["fragments"][0]["kind"], "table");
    assert_eq!(pages[4]["fragments"][0]["carriedFromPrev"], true);
    let measured = output["measured"].as_array().unwrap();
    let anchor = measured
        .iter()
        .find(|measured| {
            measured["block"]["runs"]
                .as_array()
                .is_some_and(|runs| runs.iter().any(|run| run["kind"] == "image"))
        })
        .unwrap();
    assert_eq!(anchor["measure"]["totalHeight"], 20.0);
    assert!(
        pages[4]["fragments"]
            .as_array()
            .unwrap()
            .iter()
            .any(|fragment| fragment["blockId"] == anchor["block"]["id"])
    );
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
    assert_eq!(sections.len(), 3);
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

fn wrapped_float_edit_document(cover_width: u32, body_width: u32, relative_to: &str) -> Vec<u8> {
    let paragraph = |id: u32, extra: &str, content: &str| {
        p(
            &format!("{id:08X}"),
            &format!(
                concat!(
                    "<w:pPr>{extra}<w:spacing w:before=\"0\" w:after=\"0\" ",
                    "w:line=\"300\" w:lineRule=\"exact\"/></w:pPr>{content}"
                ),
                extra = extra,
                content = content,
            ),
        )
    };
    let cover = paragraph(
        1,
        &section(cover_width * 15 + 1440, 15840, 1440, 1440, ""),
        &r("Cover"),
    );
    let before = if cover_width > body_width {
        format!("SMALL_FLOAT{}", r("Before the image"))
    } else {
        r("Before the image")
    };
    let tails = (0..40)
        .map(|index| paragraph(0x1000 + index, "", &r("Following body text")))
        .collect::<String>();
    let body = format!(
        concat!(
            "<w:sdt><w:sdtPr><w:id w:val=\"7\"/></w:sdtPr><w:sdtContent>",
            "{cover}</w:sdtContent></w:sdt>{before}{anchor}{near}{tails}{section}"
        ),
        cover = cover,
        before = paragraph(2, "", &before),
        anchor = paragraph(3, "<w:pageBreakBefore/>", "BODY_FLOAT"),
        near = paragraph(4, "", &r(&"Body words beside the image. ".repeat(12))),
        tails = tails,
        section = section(body_width * 15 + 1440, 15840, 1440, 1440, ""),
    );
    let mut parts = floating_footer_parts(&body);
    let footer = parts
        .iter()
        .find(|(name, _)| name == "word/footer1.xml")
        .unwrap();
    let drawing = std::str::from_utf8(&footer.1)
        .unwrap()
        .split_once("<w:drawing>")
        .unwrap()
        .1
        .split_once("</w:drawing>")
        .unwrap()
        .0
        .replace(
            "<wp:positionV relativeFrom=\"margin\"><wp:align>bottom</wp:align></wp:positionV>",
            &format!(
                concat!(
                    "<wp:positionV relativeFrom=\"{relative_to}\">",
                    "<wp:posOffset>0</wp:posOffset></wp:positionV>"
                ),
                relative_to = relative_to,
            ),
        )
        .replace(
            "<wp:wrapTopAndBottom/>",
            "<wp:wrapSquare wrapText=\"bothSides\"/>",
        )
        .replace("cx=\"1143000\"", "cx=\"5852160\"")
        .replace("cy=\"1143000\"", "cy=\"4876800\"")
        .replace("distL=\"0\"", "distL=\"114300\"")
        .replace("distR=\"0\"", "distR=\"114300\"");
    let image = format!(
        concat!(
            "<w:r ",
            "xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\" ",
            "xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" ",
            "xmlns:pic=\"http://schemas.openxmlformats.org/drawingml/2006/picture\">",
            "<w:drawing>{drawing}</w:drawing></w:r>"
        ),
        drawing = drawing,
    );
    let small_image = image
        .replace("5852160", "381000")
        .replace("4876800", "381000")
        .replace("id=\"1\"", "id=\"2\"");
    for (name, bytes) in &mut parts {
        let xml = std::str::from_utf8(bytes).ok();
        let replacement = match name.as_str() {
            "word/document.xml" => Some(
                xml.unwrap()
                    .replace("BODY_FLOAT", &image)
                    .replace("SMALL_FLOAT", &small_image),
            ),
            "word/_rels/document.xml.rels" => Some(xml.unwrap().replace(
                "</Relationships>",
                concat!(
                    "<Relationship Id=\"rIdImage\" ",
                    "Type=\"",
                    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image\" ",
                    "Target=\"media/footer.png\"/></Relationships>"
                ),
            )),
            _ => None,
        };
        if let Some(replacement) = replacement {
            *bytes = replacement.into_bytes();
        }
    }
    ooxml_opc::rezip_parts(&parts).unwrap()
}

fn assert_wrapped_float_measurements(input: &Value, wraps: bool) {
    let measured = input["measured"].as_array().unwrap();
    let image = measured
        .iter()
        .find(|entry| {
            entry["block"]["runs"].as_array().is_some_and(|runs| {
                runs.iter().any(|run| {
                    run["kind"] == "image"
                        && run["width"].as_f64().is_some_and(|width| width > 600.0)
                })
            })
        })
        .unwrap();
    assert_eq!(image["measure"]["totalHeight"], 20.0);
    assert_eq!(
        image["measure"]["lines"][0]["floatSkipBefore"]
            .as_f64()
            .unwrap_or(0.0),
        0.0
    );
    let text = measured
        .iter()
        .find(|entry| {
            entry["block"]["runs"].as_array().is_some_and(|runs| {
                runs.iter().any(|run| {
                    run["text"]
                        .as_str()
                        .is_some_and(|text| text.starts_with("Body words beside the image."))
                })
            })
        })
        .unwrap();
    let line = &text["measure"]["lines"][0];
    assert_eq!(line["floatSkipBefore"].as_f64().unwrap_or(0.0), 0.0);
    assert_eq!(line["leftOffset"].as_f64().unwrap_or(0.0) > 600.0, wraps);
}

fn assert_wrapped_float_edit_matches_fresh(cover_width: u32, body_width: u32, relative_to: &str) {
    let bytes = wrapped_float_edit_document(cover_width, body_width, relative_to);
    let font = docx_layout::register_measure_font(FONT).unwrap();
    for host_update in [false, true] {
        let engine = EngineSession::new(75220);
        seed_from_docx(engine.doc(), &bytes).unwrap();
        let mut request = region_request(&engine, &bytes, font);
        let package = docx_parse::parse_docx_s9_wire(&bytes, Default::default())
            .unwrap()
            .document
            .package;
        request["regions"]["sections"] = json!([{
            "properties": package.document.final_section_properties,
        }]);
        request["notes"]["contents"] = json!([]);
        let extras = json!({"fontChains": request["measurement"]["fontChains"]}).to_string();
        let request = request.to_string();
        engine
            .layout_document_with_regions_retained(&request)
            .unwrap();
        engine.build_display_list_frame(&extras, 0).unwrap();
        let initial: Value =
            serde_json::from_str(&engine.retained_kernel_inputs_json().unwrap()).unwrap();
        assert_wrapped_float_measurements(&initial, body_width == 690);
        let before = engine.stats();
        let edit = |doc: &docx_edit::EditingDoc| {
            doc.insert_text(
                &docx_edit::EditCtx::local("", ""),
                doc.paragraph_mark_position("00000004").unwrap(),
                "x",
                docx_edit::FormatPolicy::Inherit,
            )
            .unwrap();
        };
        if host_update {
            let peer = docx_edit::EditingDoc::new(75221);
            peer.apply_update_v1(&engine.doc().encode_state_as_update_v1())
                .unwrap();
            edit(&peer);
            engine
                .doc()
                .apply_host_update_v1(&peer.encode_state_as_update_v1())
                .unwrap();
            engine
                .layout_document_with_regions_retained(&request)
                .unwrap();
        } else {
            edit(engine.doc());
            engine.apply_and_layout("body", 0).unwrap();
        }
        let after = engine.stats();
        if relative_to == "paragraph" {
            assert!(after.resident_measure_calls - before.resident_measure_calls > 1);
            assert!(after.resident_reused_blocks > before.resident_reused_blocks);
        } else {
            assert_eq!(after.resident_reused_blocks, before.resident_reused_blocks);
        }
        assert!(after.incremental_pagination_calls > before.incremental_pagination_calls);
        let incremental: Value =
            serde_json::from_str(&engine.retained_kernel_inputs_json().unwrap()).unwrap();
        assert_wrapped_float_measurements(&incremental, body_width == 690);
        let fresh = EngineSession::new(75222);
        fresh
            .doc()
            .apply_update_v1(&engine.doc().encode_state_as_update_v1())
            .unwrap();
        fresh
            .layout_document_with_regions_retained(&request)
            .unwrap();
        let full: Value =
            serde_json::from_str(&fresh.retained_kernel_inputs_json().unwrap()).unwrap();
        assert_eq!(incremental["measured"], full["measured"]);
        assert_eq!(after.retained_pages, fresh.stats().retained_pages);
        assert_eq!(
            engine.retained_layout_json().unwrap(),
            fresh.retained_layout_json().unwrap()
        );
    }
}

#[test]
fn wrapped_cover_float_skip_stays_clamped_after_incremental_text_edits() {
    assert_wrapped_float_edit_matches_fresh(690, 624, "paragraph");
}

#[test]
fn wrapped_cover_body_float_wrapping_survives_incremental_text_edits() {
    for relative_to in ["paragraph", "margin", "page"] {
        assert_wrapped_float_edit_matches_fresh(624, 690, relative_to);
    }
}

fn wrapped_cover_header_document(paragraphs: u32, page_variants: bool) -> Vec<u8> {
    let cover = (0..paragraphs)
        .map(|index| {
            let properties = if index + 1 == paragraphs {
                section(7200, 9000, 900, 1200, "")
            } else {
                String::new()
            };
            p(
                &format!("{:08X}", 0x2000_0000 + index),
                &format!(
                    concat!(
                        "<w:pPr><w:spacing w:before=\"0\" w:after=\"0\" ",
                        "w:line=\"300\" w:lineRule=\"exact\"/>{properties}</w:pPr>{text}"
                    ),
                    properties = properties,
                    text = r("Cover line"),
                ),
            )
        })
        .collect::<String>();
    let refs = if page_variants {
        concat!(
            "<w:headerReference w:type=\"first\" r:id=\"rIdHeader2\"/>",
            "<w:headerReference w:type=\"even\" r:id=\"rIdHeader1\"/><w:titlePg/>"
        )
    } else {
        "<w:headerReference w:type=\"default\" r:id=\"rIdHeader2\"/>"
    };
    let body = format!(
        concat!(
            "<w:sdt><w:sdtPr><w:id w:val=\"7\"/></w:sdtPr><w:sdtContent>",
            "{cover}</w:sdtContent></w:sdt>{last}{section}"
        ),
        cover = cover,
        last = p("00000001", &r("Last")),
        section = section(12000, 8400, 2700, 1200, refs),
    );
    let header = |height| {
        format!(
            concat!(
                "<w:hdr xmlns:w=\"",
                "http://schemas.openxmlformats.org/wordprocessingml/2006/main\">",
                "<w:p><w:pPr><w:spacing w:before=\"0\" w:after=\"0\" ",
                "w:line=\"{height}\" w:lineRule=\"exact\"/></w:pPr>",
                "<w:r><w:t>Shared header</w:t></w:r></w:p></w:hdr>"
            ),
            height = height,
        )
    };
    let mut parts = fixture::page_parts(false);
    for (name, bytes) in &mut parts {
        let replacement = match name.as_str() {
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
            "word/header1.xml" => Some(header(1500)),
            "word/header2.xml" => Some(header(1200)),
            "word/footnotes.xml" => Some(
                concat!(
                    "<w:footnotes xmlns:w=\"",
                    "http://schemas.openxmlformats.org/wordprocessingml/2006/main\"/>"
                )
                .to_owned(),
            ),
            "word/endnotes.xml" => Some(
                concat!(
                    "<w:endnotes xmlns:w=\"",
                    "http://schemas.openxmlformats.org/wordprocessingml/2006/main\"/>"
                )
                .to_owned(),
            ),
            _ => None,
        };
        if let Some(replacement) = replacement {
            *bytes = replacement.into_bytes();
        }
    }
    ooxml_opc::rezip_parts(&parts).unwrap()
}

fn wrapped_cover_header_request(engine: &EngineSession, bytes: &[u8], font: u32) -> String {
    let mut request = region_request(engine, bytes, font);
    let package = docx_parse::parse_docx_s9_wire(bytes, Default::default())
        .unwrap()
        .document
        .package;
    let properties = json!(package.document.final_section_properties);
    let sections = request["regions"]["sections"].as_array_mut().unwrap();
    assert_eq!(sections.len(), 2);
    for section in sections {
        section["properties"] = properties.clone();
    }
    request["regions"]["evenAndOddHeaders"] = json!(true);
    request.to_string()
}

#[test]
fn wrapped_cover_first_and_even_headers_reserve_its_own_margins() {
    let bytes = wrapped_cover_header_document(50, true);
    let font = docx_layout::register_measure_font(FONT).unwrap();
    let engine = EngineSession::new(75230);
    seed_from_docx(engine.doc(), &bytes).unwrap();
    let request = wrapped_cover_header_request(&engine, &bytes, font);
    let output: Value = serde_json::from_str(
        &engine
            .layout_document_with_regions_retained_json(&request)
            .unwrap(),
    )
    .unwrap();
    let pages = output["layout"]["pages"].as_array().unwrap();
    assert_eq!(pages.len(), 4);
    for (index, top) in [100.0, 120.0, 60.0].into_iter().enumerate() {
        let page = &pages[index];
        assert_eq!(page["size"], json!({"w": 480.0, "h": 600.0}));
        assert_eq!(page["margins"]["top"], top);
        assert_eq!(page["fragments"][0]["y"], top);
        assert_eq!(page["sectionIndex"], 0);
        assert_eq!(page["sectionPageIndex"], index);
    }
    assert_eq!(pages[3]["size"], json!({"w": 800.0, "h": 560.0}));
    assert_eq!(pages[3]["margins"]["top"], 180.0);
}

#[test]
fn a_prefix_inside_a_wrapped_cover_with_a_shared_header_matches_full_layout() {
    let bytes = wrapped_cover_header_document(256, false);
    let font = docx_layout::register_measure_font(FONT).unwrap();
    let full = EngineSession::new(75231);
    seed_from_docx(full.doc(), &bytes).unwrap();
    let request = wrapped_cover_header_request(&full, &bytes, font);
    let expected = full
        .layout_document_with_regions_retained_json(&request)
        .unwrap();
    let expected_value: Value = serde_json::from_str(&expected).unwrap();
    let pages = expected_value["layout"]["pages"].as_array().unwrap();
    assert!(pages.len() > 10);
    assert_eq!(pages[0]["size"], json!({"w": 480.0, "h": 600.0}));
    assert_eq!(pages[0]["fragments"][0]["y"], 100.0);
    for exact in [false, true] {
        let prefix = EngineSession::new(75232);
        seed_from_docx(prefix.doc(), &bytes).unwrap();
        if exact {
            let result = prefix
                .layout_document_with_regions_exact_prefix_retained(&request, 1)
                .unwrap();
            assert!(!result.provisional);
            assert_eq!(result.pages, pages.len());
        } else {
            let result = prefix
                .layout_document_with_regions_prefix_retained_json(&request, 1)
                .unwrap();
            assert_eq!(result, expected);
        }
        assert_eq!(prefix.retained_layout_json().unwrap(), expected);
        assert_eq!(
            prefix.retained_kernel_inputs_json().unwrap(),
            full.retained_kernel_inputs_json().unwrap()
        );
    }
}

fn late_wrapped_section_float_document() -> Vec<u8> {
    let mut parts =
        ooxml_opc::unzip_parts(&wrapped_float_edit_document(690, 624, "paragraph")).unwrap();
    let document = parts
        .iter()
        .find(|(name, _)| name == "word/document.xml")
        .unwrap();
    let images: Vec<_> = std::str::from_utf8(&document.1)
        .unwrap()
        .split("<w:r xmlns:wp=")
        .skip(1)
        .map(|image| {
            let drawing = image.split_once("</w:r>").unwrap().0;
            format!("<w:r xmlns:wp={drawing}</w:r>")
        })
        .collect();
    assert_eq!(images.len(), 2);
    let paragraph = |id: u32, properties: &str, content: &str| {
        p(
            &format!("{id:08X}"),
            &format!(
                concat!(
                    "<w:pPr><w:spacing w:before=\"0\" w:after=\"0\" ",
                    "w:line=\"300\" w:lineRule=\"exact\"/>{properties}</w:pPr>{content}"
                ),
                properties = properties,
                content = content,
            ),
        )
    };
    let tails = (0..24)
        .map(|index| paragraph(0x3000 + index, "", &r("Following body text")))
        .collect::<String>();
    let body = format!(
        concat!(
            "{cover}{image}{near}{tails}{reset}<w:sdt><w:sdtPr><w:id w:val=\"7\"/>",
            "</w:sdtPr><w:sdtContent>{wrapped}</w:sdtContent></w:sdt>{outside}{last}{section}"
        ),
        cover = paragraph(
            1,
            &section(11790, 15840, 1440, 1440, ""),
            &format!("{}{}", images[0], r("Cover")),
        ),
        image = paragraph(3, "<w:pageBreakBefore/>", &images[1]),
        near = paragraph(4, "", &r(&"Body words beside the image. ".repeat(12))),
        tails = tails,
        reset = paragraph(5, "<w:pageBreakBefore/>", &r("Late section")),
        wrapped = paragraph(
            6,
            &section(10800, 15840, 1440, 1440, ""),
            &images[0].replace("id=\"2\"", "id=\"3\""),
        ),
        outside = paragraph(7, "", &images[0].replace("id=\"2\"", "id=\"4\"")),
        last = paragraph(8, "", &r("After the moved section")),
        section = section(10800, 15840, 1440, 1440, ""),
    );
    for (name, bytes) in &mut parts {
        if name == "word/document.xml" {
            *bytes = format!(
                concat!(
                    "<w:document ",
                    "xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" ",
                    "xmlns:r=\"",
                    "http://schemas.openxmlformats.org/officeDocument/2006/relationships\" ",
                    "xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\">",
                    "<w:body>{body}</w:body></w:document>"
                ),
                body = body,
            )
            .into_bytes();
        }
    }
    ooxml_opc::rezip_parts(&parts).unwrap()
}

fn retained_float_width_policy(input: &Value) -> docx_layout::measure_blocks::FloatWidthPolicy {
    let blocks: Vec<docx_layout::types::LayoutBlock> = input["measured"]
        .as_array()
        .unwrap()
        .iter()
        .map(|entry| serde_json::from_value(entry["block"].clone()).unwrap())
        .collect();
    docx_layout::measure_blocks::FloatWidthPolicy::for_document(blocks.iter())
}

#[test]
fn a_host_batch_moving_a_wrapped_section_break_remeasures_the_document_float_policy() {
    use docx_layout::measure_blocks::FloatWidthPolicy;

    let bytes = late_wrapped_section_float_document();
    let font = docx_layout::register_measure_font(FONT).unwrap();
    let engine = EngineSession::new(75240);
    seed_from_docx(engine.doc(), &bytes).unwrap();
    let mut request = region_request(&engine, &bytes, font);
    let package = docx_parse::parse_docx_s9_wire(&bytes, Default::default())
        .unwrap()
        .document
        .package;
    let cover_properties = request["regions"]["sections"][0]["properties"].clone();
    let body_properties = json!(package.document.final_section_properties);
    request["regions"]["sections"] = json!([
        {"properties": cover_properties},
        {"properties": body_properties},
        {"properties": body_properties},
    ]);
    request["notes"]["contents"] = json!([]);
    let extras = json!({"fontChains": request["measurement"]["fontChains"]}).to_string();
    let request = request.to_string();
    engine
        .layout_document_with_regions_retained(&request)
        .unwrap();
    engine.build_display_list_frame(&extras, 0).unwrap();
    let peer = docx_edit::EditingDoc::new(75241);
    peer.apply_update_v1(&engine.doc().encode_state_as_update_v1())
        .unwrap();
    peer.insert_text(
        &docx_edit::EditCtx::local("", ""),
        peer.paragraph_mark_position("00000004").unwrap(),
        "x",
        docx_edit::FormatPolicy::Inherit,
    )
    .unwrap();
    engine
        .doc()
        .apply_host_update_v1(&peer.encode_state_as_update_v1())
        .unwrap();
    engine
        .layout_document_with_regions_retained(&request)
        .unwrap();
    engine.build_display_list_frame(&extras, 0).unwrap();
    let warm: Value =
        serde_json::from_str(&engine.retained_kernel_inputs_json().unwrap()).unwrap();
    assert_eq!(retained_float_width_policy(&warm), FloatWidthPolicy::PerBlock);
    assert_wrapped_float_measurements(&warm, false);
    let owner = peer.paragraph_mark_position("00000006").unwrap().story;
    let properties = peer
        .paragraphs(&owner)
        .unwrap()
        .into_iter()
        .find(|paragraph| paragraph.para_id == "00000006")
        .unwrap()
        .properties["sectPr"]
        .clone();
    for (from, to, policy, interactive) in [
        ("00000006", "00000007", FloatWidthPolicy::FirstBlock, false),
        ("00000007", "00000006", FloatWidthPolicy::PerBlock, false),
        ("00000006", "00000007", FloatWidthPolicy::FirstBlock, true),
        ("00000007", "00000006", FloatWidthPolicy::PerBlock, true),
    ] {
        peer.set_paragraph_attr(from, "sectPr", yrs::Any::Null)
            .unwrap();
        peer.set_paragraph_attr(from, "sectionBreakType", yrs::Any::Null)
            .unwrap();
        peer.set_paragraph_attr(to, "sectPr", properties.clone())
            .unwrap();
        peer.set_paragraph_attr(to, "sectionBreakType", yrs::Any::from("nextPage"))
            .unwrap();
        engine
            .doc()
            .apply_host_update_v1(&peer.encode_state_as_update_v1())
            .unwrap();
        let before = engine.stats();
        if interactive {
            engine.apply_and_layout("body", 0).unwrap();
        } else {
            engine
                .layout_document_with_regions_retained(&request)
                .unwrap();
        }
        assert_eq!(engine.stats().resident_reused_blocks, before.resident_reused_blocks);
        let incremental: Value =
            serde_json::from_str(&engine.retained_kernel_inputs_json().unwrap()).unwrap();
        assert_eq!(retained_float_width_policy(&incremental), policy);
        assert_eq!(
            incremental["measured"].as_array().unwrap().len(),
            warm["measured"].as_array().unwrap().len()
        );
        let image = incremental["measured"]
            .as_array()
            .unwrap()
            .iter()
            .find(|entry| entry["block"]["id"] == "00000003")
            .unwrap();
        let skip = image["measure"]["lines"][0]["floatSkipBefore"]
            .as_f64()
            .unwrap_or(0.0);
        let expected_skip = if policy == FloatWidthPolicy::FirstBlock {
            512.0
        } else {
            0.0
        };
        assert_eq!(skip, expected_skip);
        let fresh = EngineSession::new(75242);
        fresh
            .doc()
            .apply_update_v1(&engine.doc().encode_state_as_update_v1())
            .unwrap();
        fresh
            .layout_document_with_regions_retained(&request)
            .unwrap();
        assert_eq!(
            engine.retained_kernel_inputs_json().unwrap(),
            fresh.retained_kernel_inputs_json().unwrap()
        );
        assert_eq!(
            engine.retained_layout_json().unwrap(),
            fresh.retained_layout_json().unwrap()
        );
        engine.build_display_list_frame(&extras, 0).unwrap();
    }
}
