#[allow(dead_code)]
#[path = "support/page_fixture.rs"]
mod fixture;

use docx_edit::{EditCtx, EngineSession, FormatPolicy, Position};
use serde_json::Value;

fn staggered_row(exact: bool) -> Vec<u8> {
    let cells: String = [("left", 90, 240), ("right", 85, 260)]
        .into_iter()
        .enumerate()
        .map(|(cell, (prefix, count, spacing))| {
            let paragraphs: String = (0..count)
                .map(|index| {
                    fixture::p(
                        &format!("{:08X}", 0x7100_0000 + cell * 0x100 + index),
                        &format!(
                            r#"<w:pPr><w:widowControl w:val="0"/><w:spacing w:before="0" w:after="0" w:line="{spacing}" w:lineRule="exact"/></w:pPr>{}"#,
                            fixture::r(&format!("{prefix}{index:03}"))
                        ),
                    )
                })
                .collect();
            format!(
                r#"<w:tc><w:tcPr><w:tcW w:w="4320" w:type="dxa"/><w:vAlign w:val="center"/></w:tcPr>{paragraphs}</w:tc>"#
            )
        })
        .collect();
    let height = if exact {
        r#"<w:trPr><w:trHeight w:val="7200" w:hRule="exact"/></w:trPr>"#
    } else {
        ""
    };
    fixture::with_body(&format!(
        r#"<w:tbl><w:tblPr><w:tblW w:w="8640" w:type="dxa"/><w:tblLayout w:type="fixed"/><w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/></w:tblCellMar></w:tblPr><w:tblGrid><w:gridCol w:w="4320"/><w:gridCol w:w="4320"/></w:tblGrid><w:tr>{height}{cells}</w:tr></w:tbl><w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>"#
    ))
}

fn staggered_row_layout(exact: bool) -> (Value, Value) {
    let (engine, request) = fixture::laid_out(&staggered_row(exact), 9403);
    let output = engine.layout_document_with_regions_json(&request).unwrap();
    let display = engine.build_display_list_json(&output).unwrap();
    (
        serde_json::from_str(&output).unwrap(),
        serde_json::from_str(&display).unwrap(),
    )
}

#[test]
fn centred_cells_with_staggered_lines_split_within_the_page_body() {
    let (output, display) = staggered_row_layout(false);
    let pages = output["layout"]["pages"].as_array().unwrap();
    assert_eq!(pages.len(), display["pages"].as_array().unwrap().len());
    let mut fragment_count = 0;
    let mut texts = [String::new(), String::new()];
    for (page, painted) in pages.iter().zip(display["pages"].as_array().unwrap()) {
        let margins = page.get("bodyMargins").unwrap_or(&page["margins"]);
        let body_top = margins["top"].as_f64().unwrap();
        let body_bottom = page["size"]["h"].as_f64().unwrap() - margins["bottom"].as_f64().unwrap();
        let fragments: Vec<_> = page["fragments"]
            .as_array()
            .unwrap()
            .iter()
            .filter(|fragment| fragment["kind"] == "table")
            .collect();
        if fragments.is_empty() {
            continue;
        }
        assert_eq!(fragments.len(), 1);
        fragment_count += 1;
        let fragment = fragments[0];
        let top = fragment["y"].as_f64().unwrap();
        let bottom = top + fragment["height"].as_f64().unwrap();
        assert!(top >= body_top - 0.02);
        assert!(bottom <= body_bottom + 0.02, "{fragment}");
        let clips = fragment["cellClips"].as_array().unwrap();
        assert_eq!(clips.len(), 2);
        assert_ne!(clips[0]["bottom"], clips[1]["bottom"]);
        for row in fragment["rowStart"].as_u64().unwrap()..fragment["rowEnd"].as_u64().unwrap() {
            assert!(
                painted["primitives"]
                    .as_array()
                    .unwrap()
                    .iter()
                    .any(|primitive| {
                        primitive["cell"]["row"].as_u64() == Some(row)
                            && primitive["text"]
                                .as_str()
                                .is_some_and(|text| !text.is_empty())
                    }),
                "row {row} has no line text in {fragment}"
            );
        }
        for primitive in painted["primitives"].as_array().unwrap() {
            let Some(text) = primitive["text"].as_str().filter(|text| !text.is_empty()) else {
                continue;
            };
            let cell = primitive["cell"]["col"].as_u64().unwrap() as usize;
            let measured_cell = &output["measured"][0]["measure"]["rows"][0]["cells"][cell];
            let line = &measured_cell["blocks"][0]["lines"][0];
            let baseline = primitive["baselineY"]
                .as_f64()
                .or_else(|| primitive["glyphs"][0]["y"].as_f64())
                .unwrap();
            let line_top = baseline - line["ascent"].as_f64().unwrap();
            let line_bottom = line_top + line["lineHeight"].as_f64().unwrap();
            let clip = &primitive["clipGroup"]["clip"];
            let clip_top = clip["y"].as_f64().unwrap();
            let clip_bottom = clip_top + clip["h"].as_f64().unwrap();
            assert!(line_top >= top.max(clip_top) - 0.02, "{text}: {line_top}");
            assert!(
                line_bottom <= bottom.min(clip_bottom) + 0.02,
                "{text}: {line_bottom}"
            );
            texts[cell].push_str(text);
        }
    }
    assert!(fragment_count > 1);
    for (cell, (prefix, count)) in [("left", 90), ("right", 85)].into_iter().enumerate() {
        assert_eq!(
            texts[cell],
            (0..count)
                .map(|index| format!("{prefix}{index:03}"))
                .collect::<String>()
        );
    }
}

#[test]
fn a_short_bottom_aligned_cell_stays_on_the_last_row_fragment() {
    assert_short_aligned_cells(&["top", "top", "bottom"]);
}

#[test]
fn short_centred_and_bottom_cells_keep_their_whole_row_alignment() {
    assert_short_aligned_cells(&["top", "top", "center", "bottom"]);
}

fn assert_short_aligned_cells(alignments: &[&str]) {
    let width = 8640 / alignments.len();
    let counts = [225, 216, 1, 1];
    let cells: String = alignments
        .iter()
        .enumerate()
        .map(|(cell, alignment)| {
            let spacing = if cell == 1 { 250 } else { 240 };
            let paragraphs: String = (0..counts[cell])
                .map(|index| {
                    fixture::p(
                        &format!("{:08X}", 0x7200_0000 + cell * 0x100 + index),
                        &format!(
                            r#"<w:pPr><w:widowControl w:val="0"/><w:spacing w:before="0" w:after="0" w:line="{spacing}" w:lineRule="exact"/></w:pPr>{}"#,
                            fixture::r(&format!("cell{cell}line{index:03}"))
                        ),
                    )
                })
                .collect();
            format!(
                r#"<w:tc><w:tcPr><w:tcW w:w="{width}" w:type="dxa"/><w:vAlign w:val="{alignment}"/></w:tcPr>{paragraphs}</w:tc>"#
            )
        })
        .collect();
    let grid = format!(r#"<w:gridCol w:w="{width}"/>"#).repeat(alignments.len());
    let bytes = fixture::with_body(&format!(
        r#"<w:tbl><w:tblPr><w:tblW w:w="8640" w:type="dxa"/><w:tblLayout w:type="fixed"/><w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/></w:tblCellMar></w:tblPr><w:tblGrid>{grid}</w:tblGrid><w:tr>{cells}</w:tr></w:tbl><w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>"#
    ));
    let (engine, request) = fixture::laid_out(&bytes, 9404);
    let output = engine.layout_document_with_regions_json(&request).unwrap();
    let display = engine.build_display_list_json(&output).unwrap();
    let output: Value = serde_json::from_str(&output).unwrap();
    let display: Value = serde_json::from_str(&display).unwrap();
    let pages = output["layout"]["pages"].as_array().unwrap();
    assert_eq!(pages.len(), 5);
    assert_eq!(pages.len(), display["pages"].as_array().unwrap().len());
    let row_height = output["measured"][0]["measure"]["rows"][0]["height"]
        .as_f64()
        .unwrap();
    let mut consumed = 0.0;
    let mut texts = vec![String::new(); alignments.len()];
    for (page_index, (page, painted)) in pages
        .iter()
        .zip(display["pages"].as_array().unwrap())
        .enumerate()
    {
        let margins = page.get("bodyMargins").unwrap_or(&page["margins"]);
        let body_top = margins["top"].as_f64().unwrap();
        let body_bottom = page["size"]["h"].as_f64().unwrap() - margins["bottom"].as_f64().unwrap();
        let fragments: Vec<_> = page["fragments"]
            .as_array()
            .unwrap()
            .iter()
            .filter(|fragment| fragment["kind"] == "table")
            .collect();
        assert_eq!(fragments.len(), 1);
        let fragment = fragments[0];
        let top = fragment["y"].as_f64().unwrap();
        let bottom = top + fragment["height"].as_f64().unwrap();
        assert!(top >= body_top - 0.02);
        assert!(bottom <= body_bottom + 0.02, "{fragment}");
        for primitive in painted["primitives"].as_array().unwrap() {
            let Some(text) = primitive["text"].as_str().filter(|text| !text.is_empty()) else {
                continue;
            };
            let cell = primitive["cell"]["col"].as_u64().unwrap() as usize;
            if cell > 1 {
                let expected_page = if alignments[cell] == "center" { 2 } else { 4 };
                assert_eq!(page_index, expected_page, "{text}");
            }
            let line = &output["measured"][0]["measure"]["rows"][0]["cells"][cell]["blocks"][0]["lines"]
                [0];
            let baseline = primitive["baselineY"]
                .as_f64()
                .or_else(|| primitive["glyphs"][0]["y"].as_f64())
                .unwrap();
            let line_top = baseline - line["ascent"].as_f64().unwrap();
            let line_height = line["lineHeight"].as_f64().unwrap();
            let line_bottom = line_top + line_height;
            if cell > 1 {
                let slack = row_height - line_height;
                let offset = if alignments[cell] == "center" {
                    slack / 2.0
                } else {
                    slack
                };
                let expected_top = top + offset - consumed;
                assert!((line_top - expected_top).abs() < 0.02, "{text}: {line_top}");
            }
            let clip = &primitive["clipGroup"]["clip"];
            let clip_top = clip["y"].as_f64().unwrap();
            let clip_bottom = clip_top + clip["h"].as_f64().unwrap();
            assert!(line_top >= top.max(clip_top) - 0.02, "{text}: {line_top}");
            assert!(
                line_bottom <= bottom.min(clip_bottom) + 0.02,
                "{text}: {line_bottom}"
            );
            texts[cell].push_str(text);
        }
        consumed += fragment["height"].as_f64().unwrap();
    }
    for (cell, text) in texts.iter().enumerate() {
        assert_eq!(
            *text,
            (0..counts[cell])
                .map(|index| format!("cell{cell}line{index:03}"))
                .collect::<String>()
        );
    }
}

#[test]
fn an_exact_height_row_with_staggered_lines_stays_whole_and_clips() {
    let (output, display) = staggered_row_layout(true);
    let pages = output["layout"]["pages"].as_array().unwrap();
    assert_eq!(pages.len(), 1);
    let fragments: Vec<_> = pages[0]["fragments"]
        .as_array()
        .unwrap()
        .iter()
        .filter(|fragment| fragment["kind"] == "table")
        .collect();
    assert_eq!(fragments.len(), 1);
    let fragment = fragments[0];
    assert_eq!(fragment["rowStart"], 0);
    assert_eq!(fragment["rowEnd"], 1);
    assert_eq!(fragment["height"], 480.0);
    assert_ne!(fragment["carriedToNext"], true);
    assert!(fragment["cellClips"].is_null());
    let top = fragment["y"].as_f64().unwrap();
    let bottom = top + fragment["height"].as_f64().unwrap();
    let mut text = String::new();
    for primitive in display["pages"][0]["primitives"].as_array().unwrap() {
        let Some(run) = primitive["text"].as_str().filter(|text| !text.is_empty()) else {
            continue;
        };
        let clip = &primitive["clipGroup"]["clip"];
        assert!((clip["y"].as_f64().unwrap() - top).abs() < 0.02);
        assert!(clip["y"].as_f64().unwrap() + clip["h"].as_f64().unwrap() <= bottom + 0.02);
        text.push_str(run);
    }
    assert!(text.contains("left000"));
    assert!(text.contains("right000"));
    assert!(!text.contains("left089"));
    assert!(!text.contains("right084"));
}

#[test]
fn an_edit_before_page_breaks_repaginates_only_its_own_page() {
    let paragraph = |id: u32| fixture::p(&format!("{id:08X}"), &fixture::r("text on its own page"));
    let page_break =
        |id: u32| fixture::p(&format!("{id:08X}"), r#"<w:r><w:br w:type="page"/></w:r>"#);
    let body: String = (0..6_u32)
        .map(|index| {
            format!(
                "{}{}",
                paragraph(0x7000_0000 + index * 2),
                page_break(0x7000_0001 + index * 2)
            )
        })
        .collect();
    let (engine, request) = fixture::laid_out(&fixture::with_body(&body), 9401);
    let before = engine.stats();
    assert!(before.retained_pages >= 6);

    engine
        .doc()
        .insert_text(
            &EditCtx::local("", ""),
            Position::new("body", 0),
            "x",
            FormatPolicy::Inherit,
        )
        .unwrap();
    let layout = engine.layout_document_with_regions_json(&request).unwrap();
    let after = engine.stats();
    assert_eq!(
        after.incremental_pagination_calls - before.incremental_pagination_calls,
        1
    );
    assert_eq!(after.rebuilt_pages, 1, "later pages converge");

    let fresh = EngineSession::new(9402);
    fresh
        .doc()
        .apply_update_v1(&engine.doc().encode_state_as_update_v1())
        .unwrap();
    assert_eq!(
        fresh.layout_document_with_regions_json(&request).unwrap(),
        layout
    );
}
