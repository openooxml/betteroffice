//! A picture anchored to the page beside text in an equal-column continuous
//! section that opens below a one-column section.

use super::*;
use crate::seed::fixture::{Package, para, run};

const PAGE: &str = r#"<w:pgSz w:w="5760" w:h="5760"/>"#;
const LEAD: &str = "00002600";
const KEPT: &str = "00002700";
const RELEASED_PAGES: usize = 4;
const RELEASED_DIGEST: u64 = 0x6302_e410_5b28_5ae3;
const MARGIN_CHANGE_PAGES: usize = 0;
const MARGIN_CHANGE_DIGEST: u64 = 0;

fn margins(left: u32, right: u32) -> String {
    format!(
        r#"<w:pgMar w:top="360" w:right="{right}" w:bottom="360" w:left="{left}" w:header="200" w:footer="200" w:gutter="0"/>"#
    )
}

fn first_section() -> String {
    para(
        "00002000",
        &format!(
            "<w:pPr><w:sectPr>{PAGE}{}</w:sectPr></w:pPr>{}",
            margins(360, 360),
            run("First section ends")
        ),
    )
}

fn words(count: usize, from: usize) -> String {
    (from..from + count)
        .map(|index| format!("word{index}"))
        .collect::<Vec<_>>()
        .join(" ")
}

fn picture(x_px: u32, width_px: u32, height_px: u32) -> String {
    let x = u64::from(x_px) * 9525;
    let cx = u64::from(width_px) * 9525;
    let cy = u64::from(height_px) * 9525;
    format!(
        r#"<w:r><w:drawing><wp:anchor simplePos="0" relativeHeight="1" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="page"><wp:posOffset>{x}</wp:posOffset></wp:positionH><wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV><wp:extent cx="{cx}" cy="{cy}"/><wp:wrapSquare wrapText="bothSides"/><wp:docPr id="7" name="picture7"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="7" name="picture7"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="image"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>"#
    )
}

/// A one-column section, then a lead paragraph of `lead` words, the
/// picture's keep-lines paragraph and filler in a continuous section whose
/// properties are `section`.
fn package(picture_x: u32, lead: usize, section: &str) -> Vec<u8> {
    let kept = para(
        KEPT,
        &format!(
            "<w:pPr><w:keepLines/></w:pPr>{}{}",
            picture(picture_x, 70, 150),
            run(&format!("Beside the picture {}", words(24, 500)))
        ),
    );
    let body = format!(
        "{}{}{kept}{}{section}",
        first_section(),
        para(LEAD, &run(&words(lead, 0))),
        filler_paragraphs(0x2800, 30),
    );
    Package::new(&body)
        .rel("image", "image", "media/image1.png")
        .bytes()
}

/// `columns` equal columns with the picture over the second one.
fn equal_columns(columns: u32, lead: usize) -> Vec<u8> {
    let section = format!(
        r#"<w:sectPr><w:type w:val="continuous"/>{PAGE}{}<w:cols w:num="{columns}" w:space="360"/></w:sectPr>"#,
        margins(360, 360)
    );
    package(if columns == 2 { 212 } else { 152 }, lead, &section)
}

/// One column whose continuous section widens both side margins.
fn side_margin_change() -> Vec<u8> {
    let section = format!(
        r#"<w:sectPr><w:type w:val="continuous"/>{PAGE}{}</w:sectPr>"#,
        margins(720, 1080)
    );
    package(120, 9, &section)
}

fn opened<'a>(engine: &'a EngineSession, bytes: &[u8]) -> CommandHost<'a> {
    let font = docx_layout::register_measure_font_bytes(lowering_pages::FONT).unwrap();
    let request = lowering_pages::region_request(engine, bytes, font).to_string();
    CommandHost::open_with(engine, request)
}

fn lead_end(engine: &EngineSession) -> u32 {
    engine
        .doc()
        .paragraph_index("body")
        .unwrap()
        .para_span(LEAD)
        .unwrap()
        .1
        - 1
}

fn fnv(hash: &mut u64, bytes: &[u8]) {
    for byte in bytes {
        *hash ^= u64::from(*byte);
        *hash = hash.wrapping_mul(0x0100_0000_01b3);
    }
}

fn layout_pages(value: &serde_json::Value) -> Option<&Vec<serde_json::Value>> {
    let mut queue = std::collections::VecDeque::from([value]);
    while let Some(value) = queue.pop_front() {
        match value {
            serde_json::Value::Object(map) => {
                if let Some(serde_json::Value::Array(pages)) = map.get("pages")
                    && pages.iter().any(|page| page.get("fragments").is_some())
                {
                    return Some(pages);
                }
                queue.extend(map.values());
            }
            serde_json::Value::Array(items) => queue.extend(items.iter()),
            _ => {}
        }
    }
    None
}

/// The page count and an FNV-1a digest of every page's size and every
/// fragment's kind, position, size and line window, in hundredths of a pixel.
fn layout_digest(engine: &EngineSession) -> (usize, u64) {
    let json = engine.retained_layout_json().unwrap();
    let value = serde_json::from_str::<serde_json::Value>(&json).unwrap();
    let pages = layout_pages(&value).unwrap();
    let number = |value: Option<&serde_json::Value>| {
        value.and_then(serde_json::Value::as_f64).map_or_else(
            || "-".to_owned(),
            |number| format!("{:.0}", (number * 100.0).round()),
        )
    };
    let mut hash = 0xcbf2_9ce4_8422_2325_u64;
    for page in pages {
        let size = page.get("size");
        fnv(
            &mut hash,
            format!(
                "P{}x{}|",
                number(size.and_then(|size| size.get("w"))),
                number(size.and_then(|size| size.get("h")))
            )
            .as_bytes(),
        );
        for fragment in page
            .get("fragments")
            .and_then(serde_json::Value::as_array)
            .into_iter()
            .flatten()
        {
            let kind = fragment
                .get("kind")
                .and_then(serde_json::Value::as_str)
                .unwrap_or("-");
            let mut line = format!("F{kind}");
            for key in ["x", "y", "width", "height", "fromLine", "toLine"] {
                line.push(',');
                line.push_str(&number(fragment.get(key)));
            }
            line.push('|');
            fnv(&mut hash, line.as_bytes());
        }
    }
    (pages.len(), hash)
}

#[test]
fn picture_over_second_of_three_equal_columns_lays_out_at_open() {
    let fonts = docx_layout::MeasureFonts::default();
    let _scope = fonts.enter();
    let bytes = equal_columns(3, 60);
    let engine = seeded_engine(9801, &bytes);
    opened(&engine, &bytes);
}

#[test]
fn picture_over_second_of_two_equal_columns_lays_out_after_typing_before_it() {
    let fonts = docx_layout::MeasureFonts::default();
    let _scope = fonts.enter();
    let bytes = equal_columns(2, 9);
    let engine = seeded_engine(9802, &bytes);
    let mut host = opened(&engine, &bytes);
    let at = lead_end(&engine);
    host.command("body", at, at, Some(&format!(" {}", words(40, 900))));
}

#[test]
fn picture_over_second_equal_column_lays_out_through_growing_inserts() {
    let fonts = docx_layout::MeasureFonts::default();
    let _scope = fonts.enter();
    let bytes = equal_columns(2, 9);
    let engine = seeded_engine(9803, &bytes);
    let mut host = opened(&engine, &bytes);
    for count in [8, 20, 40] {
        let text = format!(" {}", words(count, 900));
        let at = lead_end(&engine);
        host.command("body", at, at, Some(&text));
        if count < 40 {
            host.command("body", at, at + text.chars().count() as u32, None);
        }
    }
}

#[test]
fn picture_over_second_equal_column_lays_out_as_the_release_did() {
    let fonts = docx_layout::MeasureFonts::default();
    let _scope = fonts.enter();
    let bytes = equal_columns(3, 60);
    let engine = seeded_engine(10032, &bytes);
    opened(&engine, &bytes);
    assert_eq!(layout_digest(&engine), (RELEASED_PAGES, RELEASED_DIGEST));
}

#[test]
fn picture_below_a_side_margin_change_keeps_its_layout() {
    let fonts = docx_layout::MeasureFonts::default();
    let _scope = fonts.enter();
    let bytes = side_margin_change();
    let engine = seeded_engine(9804, &bytes);
    opened(&engine, &bytes);
    assert_eq!(
        layout_digest(&engine),
        (MARGIN_CHANGE_PAGES, MARGIN_CHANGE_DIGEST)
    );
}

#[test]
fn wrapping_that_reaches_the_phase_limit_without_repeating_fails() {
    let fonts = docx_layout::MeasureFonts::default();
    let _scope = fonts.enter();
    let bytes = equal_columns(3, 60);
    let engine = seeded_engine(9805, &bytes);
    let font = docx_layout::register_measure_font_bytes(lowering_pages::FONT).unwrap();
    let request = lowering_pages::region_request(&engine, &bytes, font).to_string();
    engine.set_local_lowering(true);
    crate::engine::WRAP_PHASE_LIMIT.with(|limit| limit.set(Some(1)));
    let laid_out = engine.layout_document_with_regions_retained(&request);
    crate::engine::WRAP_PHASE_LIMIT.with(|limit| limit.set(None));
    assert!(
        laid_out
            .unwrap_err()
            .contains("anchored shape wrapping did not converge")
    );
}
