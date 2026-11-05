#[path = "common/strikethrough.rs"]
mod fixture;

use pptx_parse::{ShapeNode, parse_pptx, write_pptx};

#[test]
fn parses_strikes_and_preserves_untouched_slide_xml() {
    let bytes = fixture::deck();
    let package = parse_pptx(&bytes).unwrap();
    let ShapeNode::Shape(shape) = &package.slides[0].shapes[0] else {
        panic!("expected shape");
    };
    let body = shape.text.as_ref().unwrap();
    let strikes = body.paragraphs[0]
        .runs
        .iter()
        .map(|run| run.properties.strike.as_deref())
        .collect::<Vec<_>>();
    assert_eq!(
        strikes,
        [Some("sngStrike"), Some("dblStrike"), Some("noStrike"), None]
    );
    assert_eq!(
        body.list_style[0]
            .default_run
            .as_ref()
            .unwrap()
            .strike
            .as_deref(),
        Some("dblStrike")
    );
    assert_eq!(
        body.list_style[1]
            .default_run
            .as_ref()
            .unwrap()
            .strike
            .as_deref(),
        Some("noStrike")
    );
    let saved = write_pptx(&package).unwrap();
    let slide = |bytes: &[u8]| {
        ooxml_opc::unzip_parts(bytes)
            .unwrap()
            .into_iter()
            .find(|(path, _)| path == "ppt/slides/slide1.xml")
            .unwrap()
            .1
    };
    assert_eq!(slide(&saved), slide(&bytes));
}
