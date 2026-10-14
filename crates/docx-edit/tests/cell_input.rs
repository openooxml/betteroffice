#[allow(dead_code)]
#[path = "support/page_fixture.rs"]
mod fixture;

use docx_edit::{EditCtx, EngineSession, FormatPolicy, Position};

fn table_body() -> String {
    let cell = |id: &str, text: &str| {
        format!(
            r#"<w:tc><w:tcPr><w:tcW w:w="2400" w:type="dxa"/></w:tcPr>{}</w:tc>"#,
            fixture::p(id, &fixture::r(text))
        )
    };
    format!(
        concat!(
            "{}",
            r#"<w:tbl><w:tblPr><w:tblW w:w="4800" w:type="dxa"/></w:tblPr>"#,
            r#"<w:tblGrid><w:gridCol w:w="2400"/><w:gridCol w:w="2400"/></w:tblGrid>"#,
            "<w:tr>{}{}</w:tr></w:tbl>{}"
        ),
        fixture::p("10000001", &fixture::r("Before the table")),
        cell("10000002", "First cell"),
        cell("10000003", "Second cell"),
        fixture::p("10000004", &fixture::r("After the table")),
    )
}

#[test]
fn typing_in_a_table_cell_lays_out_residently_as_a_full_pass_would() {
    let (engine, request) = fixture::laid_out(&fixture::with_body(&table_body()), 9401);
    engine.build_display_list_frame("{}", 0).unwrap();
    let cell = "body:t0:r0c1";
    let paragraph = engine.doc().paragraphs(cell).unwrap()[0].para_id.clone();
    assert!(engine.can_apply_input(cell, &paragraph));

    engine
        .doc()
        .insert_text(
            &EditCtx::local("", ""),
            Position::new(cell, 6),
            " typed",
            FormatPolicy::Inherit,
        )
        .unwrap();
    let before = engine.stats();
    engine.apply_and_layout(cell, 1).unwrap();
    let after = engine.stats();
    assert_eq!(
        after.resident_measure_calls - before.resident_measure_calls,
        1,
        "only the edited table is measured again"
    );
    let resident = engine.retained_kernel_inputs_json().unwrap();
    let resident_display = engine
        .with_display_list(|display| serde_json::to_value(display).unwrap())
        .unwrap();
    assert!(resident.contains("Second typed cell"));

    engine.layout_document_with_regions_json(&request).unwrap();
    engine.build_display_list_frame("{}", 2).unwrap();
    assert_eq!(engine.retained_kernel_inputs_json().unwrap(), resident);
    assert_eq!(
        engine
            .with_display_list(|display| serde_json::to_value(display).unwrap())
            .unwrap(),
        resident_display
    );
}

#[test]
fn a_table_cell_is_not_resident_input_without_region_layout() {
    let engine = EngineSession::new(9402);
    docx_edit::seed_from_docx(engine.doc(), &fixture::with_body(&table_body())).unwrap();
    engine
        .layout_document_json(
            r#"{"measured": [], "options": {"pageSize": {"w": 816, "h": 1056},
                "margins": {"top": 96, "right": 96, "bottom": 96, "left": 96}}}"#,
        )
        .unwrap();
    engine.build_display_list_frame("{}", 0).unwrap();
    let paragraph = engine.doc().paragraphs("body:t0:r0c1").unwrap()[0]
        .para_id
        .clone();
    assert!(!engine.can_apply_input("body:t0:r0c1", &paragraph));
}

#[test]
fn a_cell_edit_beside_contextual_spacing_lays_out_as_the_region_pass_would() {
    let paragraph = |id: &str, text: &str| {
        format!(
            concat!(
                r#"<w:p w14:paraId="{}"><w:pPr><w:contextualSpacing/>"#,
                r#"<w:spacing w:before="240" w:after="240"/></w:pPr>{}</w:p>"#
            ),
            id,
            fixture::r(text)
        )
    };
    let body = format!(
        concat!(
            "{}",
            r#"<w:tbl><w:tblPr><w:tblW w:w="4800" w:type="dxa"/></w:tblPr>"#,
            r#"<w:tblGrid><w:gridCol w:w="4800"/></w:tblGrid><w:tr><w:tc><w:tcPr>"#,
            r#"<w:tcW w:w="4800" w:type="dxa"/></w:tcPr>{}{}</w:tc></w:tr></w:tbl>{}"#
        ),
        fixture::p("10000001", &fixture::r("Before")),
        paragraph("10000002", "One"),
        paragraph("10000003", "Two"),
        fixture::p("10000004", &fixture::r("After")),
    );
    let (engine, request) = fixture::laid_out(&fixture::with_body(&body), 9403);
    // without note contents the region fast path owns plain edits
    let mut request: serde_json::Value = serde_json::from_str(&request).unwrap();
    request["notes"]["contents"] = serde_json::json!([]);
    let request = request.to_string();
    engine.layout_document_with_regions_json(&request).unwrap();
    engine.build_display_list_frame("{}", 0).unwrap();
    let cell = "body:t0:r0c0";
    engine
        .doc()
        .insert_text(
            &EditCtx::local("", ""),
            Position::new(cell, 1),
            "x",
            FormatPolicy::Inherit,
        )
        .unwrap();
    engine.apply_and_layout(cell, 1).unwrap();
    let resident = engine.retained_kernel_inputs_json().unwrap();
    engine.layout_document_with_regions_json(&request).unwrap();
    assert_eq!(engine.retained_kernel_inputs_json().unwrap(), resident);
}
