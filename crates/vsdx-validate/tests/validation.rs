use vsdx_validate::{
    RULE_CONNECTOR_CROSSING, RULE_DANGLING_CONNECTOR, RULE_EMPTY_SHAPE_DATA, RULE_ISOLATED_SHAPE,
    RULE_OVERLAPPING_SHAPES, validate_package,
};

fn issue_keys() -> Vec<String> {
    let bytes = include_bytes!("../../vsdx-parse/tests/fixtures/validation.vsdx");
    let package = vsdx_parse::parse_vsdx(bytes).unwrap();
    let report = validate_package(&package);
    assert_eq!(report, validate_package(&package));
    report
        .issues
        .iter()
        .map(|issue| {
            format!(
                "{}:{}:{}:{}",
                issue.rule,
                issue.shape_id,
                issue.other_shape_id.unwrap_or(0),
                issue
                    .endpoint
                    .clone()
                    .or(issue.row.clone())
                    .unwrap_or_default()
            )
        })
        .collect()
}

#[test]
fn validation_fixture_reports_each_default_rule_once() {
    assert_eq!(
        issue_keys(),
        [
            format!("{RULE_CONNECTOR_CROSSING}:6:7:"),
            format!("{RULE_DANGLING_CONNECTOR}:4:0:end"),
            format!("{RULE_DANGLING_CONNECTOR}:5:0:"),
            format!("{RULE_EMPTY_SHAPE_DATA}:1:0:Owner"),
            format!("{RULE_ISOLATED_SHAPE}:7:0:"),
            format!("{RULE_OVERLAPPING_SHAPES}:1:2:"),
        ]
    );
}
