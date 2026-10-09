use std::fmt::Write;
use std::time::Instant;

use docx_edit::{EngineSession, seed_from_docx};
use serde_json::{Value, json};

const PAGES: usize = 900;
const FONT: &[u8] = include_bytes!("../../ooxml-text/tests/fonts/LiberationSans-Regular.ttf");
const TEXT: [&str; 3] = [
    "The operations team records each inspection and reviews the results at the end of the month. This section describes the equipment, the observations collected during the visit, and the actions assigned to the next shift.",
    "A regional coordinator checks the inventory against the delivery schedule before approving a transfer. The report includes a description of the materials, their destination, and the supporting records retained for the annual review.",
    "The maintenance plan brings together routine checks and scheduled repairs. Each entry identifies the responsible team and the expected completion date. Supervisors compare these records with the previous period to track progress.",
];
const PNG: &[u8] = &[
    137, 80, 78, 71, 13, 10, 26, 10, 0, 0, 0, 13, 73, 72, 68, 82, 0, 0, 0, 1, 0, 0, 0, 1, 8, 2, 0,
    0, 0, 144, 119, 83, 222, 0, 0, 0, 12, 73, 68, 65, 84, 120, 156, 99, 96, 96, 248, 15, 0, 1, 3,
    1, 0, 8, 137, 194, 236, 0, 0, 0, 0, 73, 69, 78, 68, 174, 66, 96, 130,
];

fn synthetic_docx() -> Vec<u8> {
    let mut body = String::new();
    for page in 0..PAGES {
        let page_break = if page == 0 {
            ""
        } else {
            "<w:pageBreakBefore/>"
        };
        write!(
            body,
            r#"<w:p><w:pPr>{page_break}<w:keepNext/></w:pPr><w:r><w:rPr><w:sz w:val="28"/></w:rPr><w:t>Operations report {}</w:t></w:r></w:p>"#,
            page + 1
        )
        .unwrap();
        for paragraph in 0..(5 + page % 3) {
            write!(
                body,
                r#"<w:p><w:r><w:t>Record {}.{}: {}</w:t></w:r></w:p>"#,
                page + 1,
                paragraph + 1,
                TEXT[(page + paragraph) % TEXT.len()]
            )
            .unwrap();
        }
        if page.is_multiple_of(5) {
            body.push_str(r#"<w:tbl><w:tblPr><w:tblW w:w="9360" w:type="dxa"/><w:tblLayout w:type="fixed"/></w:tblPr><w:tblGrid><w:gridCol w:w="3120"/><w:gridCol w:w="3120"/><w:gridCol w:w="3120"/></w:tblGrid>"#);
            for row in 0..3 {
                body.push_str("<w:tr>");
                for column in 0..3 {
                    write!(
                        body,
                        r#"<w:tc><w:tcPr><w:tcW w:w="3120" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>Item {}.{}.{}</w:t></w:r></w:p></w:tc>"#,
                        page + 1,
                        row + 1,
                        column + 1
                    )
                    .unwrap();
                }
                body.push_str("</w:tr>");
            }
            body.push_str("</w:tbl>");
        }
        if page.is_multiple_of(7) {
            write!(
                body,
                r#"<w:p><w:r><w:drawing><wp:inline><wp:extent cx="914400" cy="457200"/><wp:docPr id="{}" name="Inspection image"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="{}" name="Inspection image"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="rIdImage"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="457200"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>"#,
                page + 1,
                page + 1
            )
            .unwrap();
        }
    }
    body.push_str(r#"<w:sectPr><w:headerReference w:type="default" r:id="rIdHeader"/><w:footerReference w:type="default" r:id="rIdFooter"/><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720"/></w:sectPr>"#);
    let document = format!(
        r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><w:body>{body}</w:body></w:document>"#
    );
    let parts = [
        (
            "[Content_Types].xml",
            r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/><Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/></Types>"#,
        ),
        (
            "_rels/.rels",
            r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="doc" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"#,
        ),
        (
            "word/_rels/document.xml.rels",
            r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdStyles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rIdHeader" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/><Relationship Id="rIdFooter" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/><Relationship Id="rIdImage" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image.png"/></Relationships>"#,
        ),
        (
            "word/styles.xml",
            r#"<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="22"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:before="0" w:after="80" w:line="280" w:lineRule="exact"/></w:pPr></w:pPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style></w:styles>"#,
        ),
        (
            "word/header1.xml",
            r#"<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:r><w:t>Synthetic operations report</w:t></w:r></w:p></w:hdr>"#,
        ),
        (
            "word/footer1.xml",
            r#"<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:r><w:t>Internal review copy</w:t></w:r></w:p></w:ftr>"#,
        ),
        ("word/document.xml", document.as_str()),
    ];
    let mut parts: Vec<_> = parts
        .into_iter()
        .map(|(name, xml)| (name.to_owned(), xml.as_bytes().to_vec()))
        .collect();
    parts.push(("word/media/image.png".to_owned(), PNG.to_vec()));
    ooxml_opc::rezip_parts(&parts).unwrap()
}

#[test]
#[ignore = "large DOCX benchmark; run with --release --ignored --nocapture"]
fn large_document_page_count_is_stable() {
    let started = Instant::now();
    let bytes = synthetic_docx();
    let generated = started.elapsed();
    let font = docx_layout::register_measure_font(FONT).unwrap();
    let package = docx_parse::parse_docx_s9_wire(&bytes, Default::default())
        .unwrap()
        .document
        .package;
    let engine = EngineSession::new(95001);
    seed_from_docx(engine.doc(), &bytes).unwrap();
    let request = json!({
        "bodyStory": "body",
        "renderEnv": {},
        "regions": {
            "sections": [{"properties": package.document.final_section_properties}],
            "settings": package.settings,
        },
        "measurement": {
            "fontChains": {"arial|0|0": [font]},
            "defaults": {"fontFamily": "Arial", "fontSize": 11},
            "authoritativeShaping": true,
        },
    });
    let output = engine
        .layout_document_with_regions_retained_json(&request.to_string())
        .unwrap();
    let elapsed = started.elapsed();
    let output: Value = serde_json::from_str(&output).unwrap();
    let pages = output["layout"]["pages"].as_array().unwrap();
    println!(
        "synthetic_docx pages={} docx_bytes={} generation_s={:.3} generation_and_layout_s={:.3}",
        pages.len(),
        bytes.len(),
        generated.as_secs_f64(),
        elapsed.as_secs_f64(),
    );
    assert!(
        (800..=1000).contains(&pages.len()),
        "expected 800-1,000 pages"
    );
    assert_eq!(pages.len(), PAGES, "synthetic pagination changed");
    assert_eq!(engine.stats().retained_pages, PAGES);
    let tables = pages
        .iter()
        .flat_map(|page| page["fragments"].as_array().unwrap())
        .filter(|fragment| fragment["kind"] == "table")
        .count();
    assert_eq!(tables, PAGES.div_ceil(5), "tables survive pagination");
    assert!(pages.iter().all(|page| {
        page["headerFooterRefs"]["headerDefault"] == "rIdHeader"
            && page["headerFooterRefs"]["footerDefault"] == "rIdFooter"
    }));
    let variants = output["headersFooters"]["variants"].as_array().unwrap();
    for kind in ["header", "footer"] {
        assert!(variants.iter().any(|variant| {
            variant["kind"] == kind && variant["flowHeight"].as_f64().unwrap() > 0.0
        }));
    }
}
