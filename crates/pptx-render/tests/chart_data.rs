use ooxml_drawingml::chart::{ChartData, ChartDataSeries};
use pptx_edit::DeckSession;
use pptx_render::{Primitive, SlideRenderer};

const FIXTURE: &[u8] = include_bytes!("fixtures/data-labels.pptx");
const FONT: &[u8] = include_bytes!("../../ooxml-text/tests/fonts/LiberationSans-Regular.ttf");
const NORTH_CACHE: &str = r#"<c:numCache>
            <c:pt idx="0"><c:v>12</c:v></c:pt><c:pt idx="1"><c:v>19</c:v></c:pt><c:pt idx="2"><c:v>7</c:v></c:pt>
          </c:numCache>"#;

fn chart(bytes: &[u8], id: u64) -> Primitive {
    let session = DeckSession::open(bytes, id).unwrap();
    let mut renderer = SlideRenderer::new();
    renderer.register_font("Arial", false, false, FONT).unwrap();
    renderer
        .layout_slide(session.package(), &session.snapshot().unwrap(), 0)
        .unwrap()
        .display_list
        .primitives
        .into_iter()
        .find(|primitive| matches!(primitive, Primitive::Chart { .. }))
        .expect("the first slide has a chart")
}

fn data(primitive: &Primitive) -> &ChartData {
    let Primitive::Chart { data, .. } = primitive else {
        unreachable!()
    };
    data.as_ref().expect("a plotted chart carries its data")
}

fn series(name: &str, values: &[Option<&str>]) -> ChartDataSeries {
    ChartDataSeries {
        name: Some(name.to_owned()),
        values: values
            .iter()
            .map(|value| value.map(str::to_owned))
            .collect(),
    }
}

/// The fixture with the first series' cache replaced by `cache`.
fn with_north_cache(cache: &str) -> Vec<u8> {
    let parts = ooxml_opc::unzip_parts(FIXTURE)
        .unwrap()
        .into_iter()
        .map(|(path, body)| {
            if path != "ppt/charts/chart1.xml" {
                return (path, body);
            }
            let xml = String::from_utf8(body).unwrap();
            assert_eq!(xml.matches(NORTH_CACHE).count(), 1);
            (path, xml.replace(NORTH_CACHE, cache).into_bytes())
        })
        .collect::<Vec<_>>();
    ooxml_opc::rezip_parts(&parts).unwrap()
}

/// The fixture with every series removed from its first chart.
fn without_series() -> Vec<u8> {
    let parts = ooxml_opc::unzip_parts(FIXTURE)
        .unwrap()
        .into_iter()
        .map(|(path, body)| {
            if path != "ppt/charts/chart1.xml" {
                return (path, body);
            }
            let mut xml = String::from_utf8(body).unwrap();
            while let Some(start) = xml.find("<c:ser>") {
                let end = xml[start..].find("</c:ser>").unwrap() + start + "</c:ser>".len();
                xml.replace_range(start..end, "");
            }
            (path, xml.into_bytes())
        })
        .collect::<Vec<_>>();
    ooxml_opc::rezip_parts(&parts).unwrap()
}

#[test]
fn a_chart_that_caches_no_points_carries_empty_data() {
    assert_eq!(
        data(&chart(&without_series(), 7)),
        &ChartData {
            categories: Vec::new(),
            series: Vec::new(),
        }
    );
}

#[test]
fn a_chart_carries_its_categories_and_series_values() {
    let chart = chart(FIXTURE, 4201);
    assert_eq!(
        data(&chart),
        &ChartData {
            categories: vec!["Q1".into(), "Q2".into(), "Q3".into()],
            series: vec![
                series("North", &[Some("12"), Some("19"), Some("7")]),
                series("South", &[Some("8"), Some("14"), Some("21")]),
            ],
        }
    );
}

#[test]
fn values_read_through_their_cached_format_and_blanks_stay_empty() {
    let bytes = with_north_cache(
        r#"<c:numCache><c:formatCode>0.0</c:formatCode><c:ptCount val="3"/><c:pt idx="0"><c:v>12</c:v></c:pt><c:pt idx="2"><c:v>7</c:v></c:pt></c:numCache>"#,
    );
    let chart = chart(&bytes, 4202);
    assert_eq!(
        data(&chart).series[0],
        series("North", &[Some("12.0"), None, Some("7.0")])
    );
    assert_eq!(
        data(&chart).series[1],
        series("South", &[Some("8"), Some("14"), Some("21")]),
        "the second series keeps the general format"
    );
}

#[test]
fn the_data_serializes_beside_the_summary_and_an_old_frame_still_reads() {
    let chart = chart(FIXTURE, 4203);
    let mut json = serde_json::to_value(&chart).unwrap();
    assert!(
        json["label"]
            .as_str()
            .unwrap()
            .contains("2 series, 3 categories")
    );
    assert_eq!(json["data"]["categories"][0], "Q1");
    assert_eq!(json["data"]["series"][0]["name"], "North");
    assert_eq!(json["data"]["series"][0]["values"][1], "19");
    json.as_object_mut().unwrap().remove("data");
    let Primitive::Chart { data, .. } = serde_json::from_value::<Primitive>(json).unwrap() else {
        unreachable!()
    };
    assert!(data.is_none());
}

#[test]
fn a_blank_value_serializes_as_null() {
    let bytes = with_north_cache(
        r#"<c:numCache><c:ptCount val="3"/><c:pt idx="0"><c:v>12</c:v></c:pt><c:pt idx="2"><c:v>7</c:v></c:pt></c:numCache>"#,
    );
    let json = serde_json::to_value(chart(&bytes, 4204)).unwrap();
    assert!(json["data"]["series"][0]["values"][1].is_null());
}
