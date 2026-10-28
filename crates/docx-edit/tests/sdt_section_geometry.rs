#[allow(dead_code)]
#[path = "support/page_fixture.rs"]
mod fixture;

use docx_edit::{EngineSession, seed_from_docx};
use fixture::{FONT, p, r, region_request, with_body};
use serde_json::{Value, json};

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
