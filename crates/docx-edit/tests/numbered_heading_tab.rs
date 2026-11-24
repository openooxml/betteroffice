//! Styled DOCX headings paint their numbers and text at the authored stops.

use docx_edit::layout_request::{measurement, region_layout_request};
use docx_edit::{EngineSession, seed_from_docx};
use ooxml_text::{FontId, shape};
use serde_json::{Value, json};

const FONT: &[u8] = include_bytes!("../../ooxml-text/tests/fonts/LiberationSans-Regular.ttf");
const W: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const REL: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const PACKAGE_REL: &str = "http://schemas.openxmlformats.org/package/2006/relationships";
const CONTENT: &str = "application/vnd.openxmlformats-officedocument.wordprocessingml";
const DEFAULT_TAB: f64 = 709.0 / 15.0;
const TEMPLATES: [&str; 3] = ["%1", "%1.%2", "%1.%2.%3"];
const NUMBERS: [&str; 3] = ["2", "2.1", "2.1.1"];

#[derive(Clone, Copy)]
struct Heading {
    level: usize,
    points: u32,
    left: u32,
    hanging: u32,
    tab: Option<u32>,
    suffix: Option<&'static str>,
    text: &'static str,
    rtl: bool,
}

const HEADING2: Heading = Heading {
    level: 1,
    points: 16,
    left: 0,
    hanging: 0,
    tab: None,
    suffix: None,
    text: "Beta section",
    rtl: false,
};

fn document(heading: Heading) -> Vec<u8> {
    let mut body = String::new();
    let mut styles = format!(r#"<w:styles xmlns:w="{W}">"#);
    let mut numbering = format!(
        r#"<w:numbering xmlns:w="{W}"><w:abstractNum w:abstractNumId="0"><w:multiLevelType w:val="multilevel"/>"#
    );
    for (level, &template) in TEMPLATES[..=heading.level].iter().enumerate() {
        let current = if level == heading.level {
            heading
        } else {
            Heading {
                level,
                points: if level == 0 { 20 } else { 16 },
                text: if level == 0 { "Alpha" } else { "Beta section" },
                ..HEADING2
            }
        };
        let index = level + 1;
        let style = format!("Heading{index}");
        let para_id = format!("{index:08X}");
        let half_points = current.points * 2;
        let start = if level == 0 { 2 } else { 1 };
        let left = current.left;
        let indent = if current.hanging > 0 {
            format!(r#"w:hanging="{}""#, current.hanging)
        } else {
            r#"w:firstLine="0""#.to_owned()
        };
        let tabs = current.tab.map_or_else(String::new, |position| {
            format!(r#"<w:tabs><w:tab w:val="num" w:pos="{position}"/></w:tabs>"#)
        });
        let suffix = current.suffix.map_or_else(String::new, |suffix| {
            format!(r#"<w:suff w:val="{suffix}"/>"#)
        });
        let (bidi, rtl) = if current.rtl {
            ("<w:bidi/>", "<w:rPr><w:rtl/></w:rPr>")
        } else {
            ("", "")
        };
        let text = current.text;
        body.push_str(&format!(
            r#"<w:p w14:paraId="{para_id}"><w:pPr><w:pStyle w:val="{style}"/>{bidi}</w:pPr><w:r>{rtl}<w:t>{text}</w:t></w:r></w:p>"#
        ));
        styles.push_str(&format!(
            r#"<w:style w:type="paragraph" w:styleId="{style}"><w:name w:val="heading {index}"/><w:pPr><w:numPr><w:ilvl w:val="{level}"/><w:numId w:val="1"/></w:numPr></w:pPr><w:rPr><w:rFonts w:ascii="Liberation Sans" w:hAnsi="Liberation Sans" w:cs="Liberation Sans"/><w:b/><w:bCs/><w:sz w:val="{half_points}"/><w:szCs w:val="{half_points}"/></w:rPr></w:style>"#
        ));
        numbering.push_str(&format!(
            r#"<w:lvl w:ilvl="{level}"><w:start w:val="{start}"/><w:numFmt w:val="decimal"/><w:pStyle w:val="{style}"/>{suffix}<w:lvlText w:val="{template}"/><w:lvlJc w:val="left"/><w:pPr>{tabs}<w:ind w:left="{left}" {indent}/></w:pPr></w:lvl>"#
        ));
    }
    styles.push_str("</w:styles>");
    numbering.push_str(
        r#"</w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>"#,
    );
    let mut types = format!(
        r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="{CONTENT}.document.main+xml"/>"#
    );
    let mut relationships = format!(r#"<Relationships xmlns="{PACKAGE_REL}">"#);
    for part in ["styles", "numbering", "settings"] {
        types.push_str(&format!(
            r#"<Override PartName="/word/{part}.xml" ContentType="{CONTENT}.{part}+xml"/>"#
        ));
        relationships.push_str(&format!(
            r#"<Relationship Id="{part}" Type="{REL}/{part}" Target="{part}.xml"/>"#
        ));
    }
    types.push_str("</Types>");
    relationships.push_str("</Relationships>");
    let parts = [
        ("[Content_Types].xml", types),
        (
            "_rels/.rels",
            format!(r#"<Relationships xmlns="{PACKAGE_REL}"><Relationship Id="main" Type="{REL}/officeDocument" Target="word/document.xml"/></Relationships>"#),
        ),
        ("word/_rels/document.xml.rels", relationships),
        (
            "word/document.xml",
            format!(r#"<w:document xmlns:w="{W}" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:body>{body}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr></w:body></w:document>"#),
        ),
        ("word/styles.xml", styles),
        ("word/numbering.xml", numbering),
        (
            "word/settings.xml",
            format!(r#"<w:settings xmlns:w="{W}"><w:defaultTabStop w:val="709"/></w:settings>"#),
        ),
    ];
    ooxml_opc::rezip_parts_borrowed(&parts.map(|(name, xml)| (name.to_owned(), xml))).unwrap()
}

struct Geometry {
    fragment_x: f64,
    fragment_width: f64,
    marker_x: f64,
    marker_width: f64,
    space_width: f64,
    text_x: f64,
    text_right: f64,
    second_line_x: Option<f64>,
}

fn x(run: &Value) -> f64 {
    run["x"]
        .as_f64()
        .or_else(|| run["glyphs"][0]["x"].as_f64())
        .unwrap()
}

fn right(run: &Value) -> f64 {
    if let Some(glyphs) = run["glyphs"].as_array() {
        let glyph = glyphs.last().unwrap();
        glyph["x"].as_f64().unwrap() + glyph["advance"].as_f64().unwrap()
    } else {
        x(run) + run["width"].as_f64().unwrap()
    }
}

fn render(heading: Heading) -> Geometry {
    docx_layout::with_private_measure_fonts(|| {
        let font = docx_layout::register_measure_font_bytes(FONT).unwrap();
        let bytes = document(heading);
        let package = docx_parse::parse_docx_s9_wire(&bytes, Default::default())
            .unwrap()
            .document
            .package;
        let package = serde_json::to_value(package).unwrap();
        let font_chains = json!({"liberation sans|0|0": [font], "liberation sans|1|0": [font]});
        let mut request = region_layout_request(&package, 0.0);
        request["measurement"] = measurement(&package);
        request["measurement"]["fontChains"] = font_chains.clone();
        request["measurement"]["defaults"] = json!({"fontFamily": "Liberation Sans", "fontSize": 16});
        let engine = EngineSession::new(79102);
        seed_from_docx(engine.doc(), &bytes).unwrap();
        let output = engine
            .layout_document_with_regions_json(&request.to_string())
            .unwrap();
        let mut layout: Value = serde_json::from_str(&output).unwrap();
        layout["fontChains"] = font_chains;
        let display = engine.build_display_list_json(&layout.to_string()).unwrap();
        let display: Value = serde_json::from_str(&display).unwrap();
        let block = &layout["measured"][heading.level]["block"];
        let fragment = layout["layout"]["pages"][0]["fragments"]
            .as_array()
            .unwrap()
            .iter()
            .find(|fragment| fragment["blockId"] == block["id"])
            .unwrap();
        let para_id = format!("{:08X}", heading.level + 1);
        let runs = display["pages"][0]["primitives"]
            .as_array()
            .unwrap()
            .iter()
            .filter(|run| {
                run["paraId"].as_str() == Some(para_id.as_str())
                    && run["text"].as_str().is_some_and(|text| !text.is_empty())
            });
        let marker = runs
            .clone()
            .find(|run| run["listMarker"].as_bool() == Some(true))
            .expect("numbering from paragraph style");
        let number = NUMBERS[heading.level];
        assert_eq!(marker["text"].as_str(), Some(number));
        let body = |line| {
            runs.clone().filter(move |run| {
                run["listMarker"].as_bool() != Some(true) && run["lineIndex"].as_u64() == Some(line)
            })
        };
        let width = |text| {
            docx_layout::with_measure_fonts_ref(|store| {
                let size = heading.points as f32 * 4.0 / 3.0;
                shape(store, FontId::from_u32(font), text, size, &[])
                    .unwrap()
                    .iter()
                    .map(|glyph| f64::from(glyph.x_advance))
                    .sum()
            })
        };
        Geometry {
            fragment_x: fragment["x"].as_f64().unwrap(),
            fragment_width: fragment["width"].as_f64().unwrap(),
            marker_x: x(marker),
            marker_width: width(number),
            space_width: width(" "),
            text_x: x(body(0).next().expect("first line text")),
            text_right: body(0).map(right).reduce(f64::max).unwrap(),
            second_line_x: body(1).next().map(x),
        }
    })
}

fn near(actual: f64, expected: f64, name: &str) {
    assert!((actual - expected).abs() < 0.5, "{name}: {actual} vs {expected}");
}

fn assert_ltr_positions(heading: Heading, geometry: Geometry) {
    let left = f64::from(heading.left) / 15.0;
    let first = left - f64::from(heading.hanging) / 15.0;
    let number_end = first + geometry.marker_width;
    let text_start = match (heading.suffix, heading.tab) {
        (Some("space"), _) => number_end + geometry.space_width,
        (Some("nothing"), _) => number_end,
        (_, Some(tab)) => f64::from(tab) / 15.0,
        _ if heading.hanging > 0 && left > number_end => left,
        _ => ((number_end / DEFAULT_TAB).floor() + 1.0) * DEFAULT_TAB,
    };
    assert_eq!(geometry.marker_x, geometry.fragment_x + first, "marker origin");
    near(geometry.text_x, geometry.fragment_x + text_start, "text origin");
    assert!(
        geometry.text_x + 0.001 >= geometry.marker_x + geometry.marker_width,
        "text at {} overlaps number ending at {}",
        geometry.text_x,
        geometry.marker_x + geometry.marker_width
    );
}

#[test]
fn heading1_zero_hanging_uses_the_default_tab() {
    let heading = Heading {
        level: 0,
        points: 20,
        text: "Alpha",
        ..HEADING2
    };
    assert_ltr_positions(heading, render(heading));
}

#[test]
fn heading2_zero_hanging_uses_the_default_tab() {
    assert_ltr_positions(HEADING2, render(HEADING2));
}

#[test]
fn heading_number_that_fits_uses_the_hanging_indent() {
    let heading = Heading {
        left: 720,
        hanging: 720,
        ..HEADING2
    };
    assert_ltr_positions(heading, render(heading));
}

#[test]
fn heading_number_that_overruns_the_hanging_indent_uses_the_default_tab() {
    let heading = Heading {
        left: 240,
        hanging: 240,
        ..HEADING2
    };
    assert_ltr_positions(heading, render(heading));
}

#[test]
fn heading_zero_hanging_uses_the_numbering_level_tab() {
    let heading = Heading {
        tab: Some(1134),
        ..HEADING2
    };
    assert_ltr_positions(heading, render(heading));
}

#[test]
fn heading_zero_hanging_with_space_suffix_advances_one_space() {
    let heading = Heading {
        suffix: Some("space"),
        ..HEADING2
    };
    assert_ltr_positions(heading, render(heading));
}

#[test]
fn heading_zero_hanging_with_nothing_suffix_advances_only_the_number() {
    let heading = Heading {
        suffix: Some("nothing"),
        ..HEADING2
    };
    assert_ltr_positions(heading, render(heading));
}

#[test]
fn three_part_heading_number_at_zero_hanging_uses_the_next_default_tab() {
    let heading = Heading {
        level: 2,
        points: 20,
        text: "Gamma section",
        ..HEADING2
    };
    assert_ltr_positions(heading, render(heading));
}

#[test]
fn large_heading_number_at_zero_hanging_uses_the_next_default_tab() {
    let heading = Heading {
        points: 28,
        ..HEADING2
    };
    assert_ltr_positions(heading, render(heading));
}

#[test]
fn wrapped_zero_hanging_heading_returns_to_the_left_indent() {
    let heading = Heading {
        left: 360,
        text: "Beta section describes the footpath across the meadow and follows the stream past the old mill to the bridge at the edge of town",
        ..HEADING2
    };
    let geometry = render(heading);
    near(
        geometry.second_line_x.expect("the heading wraps"),
        geometry.fragment_x + f64::from(heading.left) / 15.0,
        "second line origin",
    );
    assert_ltr_positions(heading, geometry);
}

#[test]
fn bidi_heading_at_zero_hanging_keeps_number_and_text_apart() {
    let heading = Heading {
        text: "שלום",
        rtl: true,
        ..HEADING2
    };
    let geometry = render(heading);
    let fragment_right = geometry.fragment_x + geometry.fragment_width;
    assert!(
        geometry.text_right <= geometry.marker_x + 0.001,
        "text ending at {} overlaps number at {}",
        geometry.text_right,
        geometry.marker_x
    );
    near(
        geometry.marker_x + geometry.marker_width,
        fragment_right,
        "marker right edge",
    );
    let tab = (geometry.marker_width / DEFAULT_TAB).floor() + 1.0;
    near(
        geometry.text_right,
        fragment_right - tab * DEFAULT_TAB,
        "text right edge",
    );
}
