//! Layout reservations a host passes in the render environment: inline boxes and blocks that
//! take room without painting, and ranges hidden from render only.

#[allow(dead_code)]
#[path = "support/page_fixture.rs"]
mod fixture;

use std::time::Instant;

use docx_edit::structured::{PageExportOptions, RevisionView};
use docx_edit::{EngineSession, seed_from_docx};
use serde_json::{Value, json};

const SECTION: &str = r#"<w:sectPr><w:pgSz w:w="7200" w:h="5760"/><w:pgMar w:top="720" w:right="720" w:bottom="720" w:left="720" w:header="300" w:footer="300" w:gutter="0"/></w:sectPr>"#;
const LEFT: f64 = 48.0;
const COLUMN: f64 = 384.0;
const BODY: f64 = 288.0;

#[derive(Debug, Clone, Copy)]
struct Rect {
    page: usize,
    x: f64,
    y: f64,
    width: f64,
    height: f64,
}

#[derive(Debug, Clone)]
struct Reserved {
    page: usize,
    owner: String,
    id: String,
    rect: Rect,
}

fn document(paragraphs: &[(&str, &str)]) -> Vec<u8> {
    let body: String = paragraphs
        .iter()
        .map(|(id, text)| fixture::p(id, &fixture::r(text)))
        .collect();
    fixture::with_body(&format!("{body}{SECTION}"))
}

fn inline(story: &str, para: &str, at: u32, width: f64, height: f64) -> Value {
    json!({
        "owner": "plugin", "id": format!("{para}@{at}"), "story": story, "paraId": para,
        "start": at, "end": at, "kind": "inline", "width": width, "height": height,
    })
}

fn block(para: &str, side: &str, height: f64) -> Value {
    json!({
        "owner": "plugin", "id": format!("{para}-{side}"), "story": "body", "paraId": para,
        "start": 0, "end": 0, "kind": "block", "side": side, "height": height,
    })
}

fn hide(para: &str, start: u32, end: u32) -> Value {
    json!({
        "owner": "plugin", "id": format!("{para}:{start}-{end}"), "story": "body",
        "paraId": para, "start": start, "end": end, "kind": "hide",
    })
}

fn set(items: Vec<Value>) -> Value {
    Value::Object(
        items
            .into_iter()
            .enumerate()
            .map(|(index, item)| (format!("k{index}"), item))
            .collect(),
    )
}

fn request(engine: &EngineSession, bytes: &[u8], reservations: &Value) -> Value {
    let font = docx_layout::register_measure_font(fixture::FONT).unwrap();
    let mut request = fixture::region_request(engine, bytes, font);
    request["renderEnv"]["reservations"] = reservations.clone();
    request
}

/// A laid-out session with the display extras a host passes.
struct Laid {
    engine: EngineSession,
    extras: String,
}

impl std::ops::Deref for Laid {
    type Target = EngineSession;

    fn deref(&self) -> &EngineSession {
        &self.engine
    }
}

fn laid_out(bytes: &[u8], reservations: Value) -> Laid {
    let engine = EngineSession::new(7101);
    seed_from_docx(engine.doc(), bytes).unwrap();
    let request = request(&engine, bytes, &reservations);
    let output: Value = serde_json::from_str(
        &engine
            .layout_document_with_regions_retained_json(&request.to_string())
            .unwrap(),
    )
    .unwrap();
    let extras = json!({
        "headersFooters": output["headersFooters"],
        "fontChains": request["measurement"]["fontChains"],
    })
    .to_string();
    engine.build_display_list_frame(&extras, 0).unwrap();
    Laid { engine, extras }
}

fn display(engine: &EngineSession) -> Value {
    engine
        .with_display_list(|list| serde_json::to_value(list).unwrap())
        .unwrap()
}

fn number(value: &Value) -> f64 {
    value.as_f64().unwrap()
}

fn reserved(engine: &EngineSession) -> Vec<Reserved> {
    let display = display(engine);
    let mut found = Vec::new();
    for (page, value) in display["pages"].as_array().unwrap().iter().enumerate() {
        for item in value["reservations"].as_array().into_iter().flatten() {
            found.push(Reserved {
                page,
                owner: item["owner"].as_str().unwrap().to_owned(),
                id: item["id"].as_str().unwrap().to_owned(),
                rect: Rect {
                    page,
                    x: number(&item["x"]),
                    y: number(&item["y"]),
                    width: number(&item["width"]),
                    height: number(&item["height"]),
                },
            });
        }
    }
    found
}

fn rects(engine: &EngineSession, from: i64, to: i64) -> Vec<Rect> {
    let rects: Value =
        serde_json::from_str(&engine.display_range_rects_json(from, to).unwrap()).unwrap();
    rects
        .as_array()
        .unwrap()
        .iter()
        .filter(|rect| number(&rect["width"]) > 0.0)
        .map(|rect| Rect {
            page: rect["pageIndex"].as_u64().unwrap() as usize,
            x: number(&rect["x"]),
            y: number(&rect["y"]),
            width: number(&rect["width"]),
            height: number(&rect["height"]),
        })
        .collect()
}

fn rect(engine: &EngineSession, from: i64, to: i64) -> Rect {
    *rects(engine, from, to)
        .first()
        .unwrap_or_else(|| panic!("no rect for {from}..{to}"))
}

/// The vertical middle of a text rect, which overhangs its line box by the glyph ascent.
fn middle(rect: Rect) -> f64 {
    rect.y + rect.height / 2.0
}

fn assert_near(actual: f64, expected: f64, what: &str) {
    assert!(
        (actual - expected).abs() < 0.5,
        "{what}: {actual} vs {expected}"
    );
}

#[test]
fn an_inline_reservation_takes_its_width_and_height_at_its_anchor() {
    let bytes = document(&[("10000001", "Alpha beta gamma"), ("10000002", "Next")]);
    let plain = laid_out(&bytes, json!({}));
    let engine = laid_out(&bytes, set(vec![inline("body", "10000001", 6, 40.0, 60.0)]));

    let beta = rect(&plain, 7, 11);
    assert_near(rect(&engine, 7, 11).x, beta.x + 40.0, "text after the box");
    assert_near(
        rect(&engine, 1, 6).x,
        rect(&plain, 1, 6).x,
        "text before the box",
    );

    let boxes = reserved(&engine);
    assert_eq!(boxes.len(), 1, "{boxes:?}");
    let reservation = &boxes[0];
    assert_eq!(
        (
            reservation.page,
            reservation.owner.as_str(),
            reservation.id.as_str()
        ),
        (0, "plugin", "10000001@6")
    );
    assert_near(reservation.rect.x, beta.x, "box x");
    assert_near(reservation.rect.width, 40.0, "box width");
    assert_near(reservation.rect.height, 60.0, "box height");
    let next = rect(&engine, 19, 23);
    assert!(
        reservation.rect.y + reservation.rect.height <= middle(next),
        "{reservation:?} overlaps {next:?}"
    );
    assert!(
        next.y >= rect(&plain, 19, 23).y + 40.0,
        "the line grows to the box"
    );
    assert!(reserved(&plain).is_empty());
}

#[test]
fn a_block_reservation_pushes_the_following_text_by_its_height() {
    let bytes = document(&[("10000001", "Alpha"), ("10000002", "Next")]);
    let plain = laid_out(&bytes, json!({}));
    let engine = laid_out(
        &bytes,
        set(vec![
            block("10000001", "before", 30.0),
            block("10000001", "after", 100.0),
        ]),
    );
    let alpha = rect(&engine, 1, 6);
    let pitch = rect(&plain, 8, 12).y - rect(&plain, 1, 6).y;
    assert_near(alpha.y, rect(&plain, 1, 6).y + 30.0, "before block");
    let next = rect(&engine, 8, 12);
    assert_near(next.y, rect(&plain, 8, 12).y + 130.0, "both blocks");

    let boxes = reserved(&engine);
    assert_eq!(boxes.len(), 2, "{boxes:?}");
    let before = boxes
        .iter()
        .find(|item| item.id == "10000001-before")
        .unwrap();
    let after = boxes
        .iter()
        .find(|item| item.id == "10000001-after")
        .unwrap();
    for (reservation, height) in [(before, 30.0), (after, 100.0)] {
        assert_near(reservation.rect.x, LEFT, "block x");
        assert_near(reservation.rect.width, COLUMN, "block width");
        assert_near(reservation.rect.height, height, "block height");
    }
    assert_near(before.rect.y, 48.0, "before block at the top of the body");
    assert_near(
        after.rect.y,
        48.0 + 30.0 + pitch,
        "after block below its paragraph",
    );
    assert!(
        after.rect.y + after.rect.height <= middle(next),
        "{after:?} {next:?}"
    );
}

#[test]
fn a_wide_block_reservation_is_clipped_to_its_column() {
    let bytes = document(&[("10000001", "Alpha")]);
    let mut wide = block("10000001", "before", 30.0);
    wide["width"] = json!(1000.0);
    let engine = laid_out(&bytes, set(vec![wide]));
    let boxes = reserved(&engine);
    assert_eq!(boxes.len(), 1);
    assert_near(boxes[0].rect.width, COLUMN, "clipped block width");
    assert!(boxes[0].rect.x >= LEFT);
    assert!(boxes[0].rect.x + boxes[0].rect.width <= LEFT + COLUMN);
}

#[test]
fn a_block_taller_than_the_page_starts_the_next_page_clipped_to_the_body() {
    let bytes = document(&[("10000001", "Alpha"), ("10000002", "Next")]);
    let engine = laid_out(&bytes, set(vec![block("10000001", "after", 5000.0)]));
    let boxes = reserved(&engine);
    assert_eq!(boxes.len(), 1, "{boxes:?}");
    assert_eq!(boxes[0].page, 1);
    assert_near(boxes[0].rect.y, 48.0, "top of the body");
    assert_near(boxes[0].rect.height, BODY, "clipped to the body");
    assert_eq!(rect(&engine, 8, 12).page, 2);
}

#[test]
fn an_inline_box_wider_than_the_column_takes_its_own_line_and_overflows() {
    let bytes = document(&[("10000001", "Alpha beta gamma")]);
    let engine = laid_out(
        &bytes,
        set(vec![inline("body", "10000001", 6, 1000.0, 20.0)]),
    );
    let boxes = reserved(&engine);
    assert_eq!(boxes.len(), 1, "{boxes:?}");
    let reservation = boxes[0].rect;
    assert_near(reservation.width, 1000.0, "never shrunk or split");
    let alpha = rect(&engine, 1, 6);
    let beta = *rects(&engine, 7, 11).last().unwrap();
    assert!(reservation.y >= middle(alpha), "{reservation:?} {alpha:?}");
    assert!(
        middle(beta) >= reservation.y + reservation.height,
        "{beta:?} {reservation:?}"
    );
}

#[test]
fn hiding_removes_a_range_and_a_whole_paragraph_from_render_only() {
    let bytes = document(&[
        ("10000001", "Alpha beta gamma"),
        ("10000002", "Hidden"),
        ("10000003", "Last"),
    ]);
    let plain = laid_out(&bytes, json!({}));
    let engine = laid_out(
        &bytes,
        set(vec![hide("10000001", 6, 11), hide("10000002", 0, 7)]),
    );
    assert!(rects(&engine, 7, 11).is_empty());
    assert_near(
        rect(&engine, 12, 17).x,
        rect(&plain, 7, 11).x,
        "gamma takes beta's place",
    );
    assert!(rects(&engine, 19, 25).is_empty());
    assert_near(
        rect(&engine, 27, 31).y,
        rect(&plain, 19, 25).y,
        "the paragraph is gone",
    );
    assert!(reserved(&engine).is_empty());
}

#[test]
fn a_reserved_box_owns_no_document_positions() {
    let bytes = document(&[("10000001", "Alpha beta gamma")]);
    let engine = laid_out(&bytes, set(vec![inline("body", "10000001", 6, 40.0, 20.0)]));
    let reservation = reserved(&engine)[0].rect;
    for x in [
        reservation.x + 2.0,
        reservation.x + reservation.width / 2.0,
        reservation.x + reservation.width - 2.0,
    ] {
        let hit: Value = serde_json::from_str(
            &engine
                .display_hit_test_regions_json(0, x, reservation.y + reservation.height - 4.0)
                .unwrap(),
        )
        .unwrap();
        assert_eq!(hit["pos"], json!(7), "{hit} at {x}");
    }
    let caret = engine
        .with_display_list(|list| docx_layout::hit::caret_rect(list, 7))
        .unwrap()
        .unwrap();
    let caret: Value = serde_json::to_value(caret).unwrap();
    let x = number(&caret["x"]);
    assert!(
        x <= reservation.x + 0.5 || x >= reservation.x + reservation.width - 0.5,
        "caret {x} inside {reservation:?}"
    );
}

#[test]
fn reservations_lay_out_in_table_cells_and_headers() {
    let table = r#"<w:tbl><w:tblPr><w:tblW w:w="4800" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="4800"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="4800" w:type="dxa"/></w:tcPr><w:p w14:paraId="30000001"><w:r><w:t>Cell text</w:t></w:r></w:p></w:tc></w:tr></w:tbl>"#;
    let bytes = fixture::with_body(&format!(
        "{table}{}{SECTION}",
        fixture::p("30000002", &fixture::r("After"))
    ));
    let plain = laid_out(&bytes, json!({}));
    let mut cell = inline("body:t0:r0c0", "30000001", 0, 30.0, 25.0);
    cell["id"] = json!("cell");
    let engine = laid_out(&bytes, set(vec![cell]));
    let boxes = reserved(&engine);
    assert_eq!(boxes.len(), 1, "{boxes:?}");
    assert_eq!(boxes[0].id, "cell");
    assert!(boxes[0].rect.x >= LEFT - 0.5 && boxes[0].rect.x < LEFT + 20.0);
    assert_near(boxes[0].rect.width, 30.0, "cell box width");
    let after = rects(&plain, 0, 1000)
        .into_iter()
        .map(|rect| rect.y)
        .fold(0.0, f64::max);
    let moved = rects(&engine, 0, 1000)
        .into_iter()
        .map(|rect| rect.y)
        .fold(0.0, f64::max);
    assert!(moved >= after + 10.0, "the row grows: {after} -> {moved}");

    let header = fixture::headed_docx(&fixture::p("21000001", &fixture::r("Running header")));
    let mut running = inline("hf:rIdHeader2", "21000001", 0, 12.0, 12.0);
    running["id"] = json!("running");
    let mut first = inline("hf:rIdHeader1", "20000001", 0, 12.0, 12.0);
    first["id"] = json!("first");
    let engine = laid_out(&header, set(vec![running, first]));
    let boxes = reserved(&engine);
    let pages = display(&engine)["pages"].as_array().unwrap().len();
    assert!(pages > 2, "{pages} pages");
    for page in 0..pages {
        let on_page: Vec<_> = boxes.iter().filter(|item| item.page == page).collect();
        assert_eq!(on_page.len(), 1, "page {page}: {boxes:?}");
        assert_eq!(on_page[0].id, if page == 0 { "first" } else { "running" });
        assert!(on_page[0].rect.y < 48.0, "in the header band: {on_page:?}");
    }
}

#[test]
fn wide_block_reservations_are_clipped_in_cells_and_headers() {
    let table = r#"<w:tbl><w:tblPr><w:tblW w:w="4800" w:type="dxa"/><w:tblLayout w:type="fixed"/></w:tblPr><w:tblGrid><w:gridCol w:w="4800"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="4800" w:type="dxa"/></w:tcPr><w:p w14:paraId="30000001"><w:r><w:t>Cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl>"#;
    let bytes = fixture::with_body(&format!("{table}{SECTION}"));
    let mut cell = block("30000001", "before", 10.0);
    cell["story"] = json!("body:t0:r0c0");
    let plain = laid_out(&bytes, set(vec![cell.clone()]));
    cell["width"] = json!(1000.0);
    let wide = laid_out(&bytes, set(vec![cell]));
    assert_near(
        reserved(&wide)[0].rect.width,
        reserved(&plain)[0].rect.width,
        "cell content width",
    );
    assert!(reserved(&wide)[0].rect.width < COLUMN);

    let bytes = fixture::headed_docx(&fixture::p("21000001", &fixture::r("Running")));
    let mut running = block("21000001", "before", 10.0);
    running["story"] = json!("hf:rIdHeader2");
    let plain = laid_out(&bytes, set(vec![running.clone()]));
    running["width"] = json!(1000.0);
    let wide = laid_out(&bytes, set(vec![running]));
    let expected = reserved(&plain);
    let actual = reserved(&wide);
    assert!(actual.len() > 1);
    assert_eq!(actual.len(), expected.len());
    for (wide, plain) in actual.iter().zip(expected.iter()) {
        assert_near(wide.rect.width, plain.rect.width, "header band width");
    }
}

#[test]
fn an_inline_box_in_a_right_to_left_paragraph_sits_in_visual_order() {
    let paragraph = r#"<w:p w14:paraId="40000001"><w:pPr><w:bidi/></w:pPr><w:r><w:rPr><w:rtl/></w:rPr><w:t xml:space="preserve">שלום עולם</w:t></w:r></w:p>"#;
    let bytes = fixture::with_body(&format!("{paragraph}{SECTION}"));
    let engine = laid_out(&bytes, set(vec![inline("body", "40000001", 5, 30.0, 10.0)]));
    let boxes = reserved(&engine);
    assert_eq!(boxes.len(), 1, "{boxes:?}");
    let reservation = boxes[0].rect;
    assert_near(reservation.width, 30.0, "box width");
    assert!(reservation.x >= LEFT - 0.5 && reservation.x + 30.0 <= LEFT + COLUMN + 0.5);
    let first_word = rect(&engine, 1, 5);
    assert!(
        reservation.x + reservation.width <= first_word.x + 0.5,
        "{reservation:?} is not left of {first_word:?}"
    );
}

#[test]
fn an_anchor_past_its_paragraph_or_in_a_missing_one_is_skipped() {
    let bytes = document(&[("10000001", "Alpha")]);
    let engine = laid_out(
        &bytes,
        set(vec![
            inline("body", "10000001", 50, 40.0, 20.0),
            inline("body", "missing", 0, 40.0, 20.0),
            hide("10000001", 3, 50),
        ]),
    );
    assert!(reserved(&engine).is_empty());
    assert_eq!(rects(&engine, 1, 6).len(), 1);
}

#[test]
fn a_paged_export_lays_out_without_reservations() {
    let bytes = document(&[("10000001", "Alpha"), ("10000002", "Next")]);
    let options = PageExportOptions::new(RevisionView::Markup);
    let export = |engine: &EngineSession| {
        let read = engine.export_structured_with_pages(&options).unwrap();
        let layout = serde_json::to_value(&read.content.layout).unwrap();
        (layout["pages"].clone(), layout["fragments"].clone())
    };
    let expected = export(&laid_out(&bytes, json!({})));
    let engine = laid_out(
        &bytes,
        set(vec![
            block("10000001", "after", 5000.0),
            hide("10000002", 0, 2),
        ]),
    );
    assert_eq!(rect(&engine, 8, 12).page, 2);
    assert_eq!(export(&engine), expected);
    let layout_version = engine
        .export_structured_with_pages(&options)
        .unwrap()
        .content
        .layout
        .layout_version;
    let pinned = PageExportOptions {
        expect_layout_version: Some(layout_version.clone()),
        ..options.clone()
    };
    let again = engine.export_structured_with_pages(&pinned).unwrap();
    assert_eq!(again.content.layout.layout_version, layout_version);
    engine.build_display_list_frame(&engine.extras, 0).unwrap();
    assert_eq!(reserved(&engine).len(), 1);
    assert_eq!(rect(&engine, 8, 12).page, 2);
}

#[test]
fn a_thousand_inline_reservations_each_report_a_box() {
    let paragraphs: Vec<(String, String)> = (0..1000)
        .map(|index| (format!("5{index:07}"), format!("Item {index} text")))
        .collect();
    let borrowed: Vec<(&str, &str)> = paragraphs
        .iter()
        .map(|(id, text)| (id.as_str(), text.as_str()))
        .collect();
    let bytes = document(&borrowed);
    let items = set(paragraphs
        .iter()
        .map(|(id, _)| inline("body", id, 0, 10.0, 10.0))
        .collect());
    let started = Instant::now();
    laid_out(&bytes, json!({}));
    let plain = started.elapsed();
    let started = Instant::now();
    let engine = laid_out(&bytes, items);
    let reserved_time = started.elapsed();
    eprintln!("1000 inline reservations: plain {plain:?}, reserved {reserved_time:?}");
    assert_eq!(reserved(&engine).len(), 1000);
}
