//! A picture anchored to the page beside text in an equal-column continuous
//! section that opens below a one-column section.

use super::*;
use crate::seed::fixture::{Package, image, para, run};

const PAGE: &str = r#"<w:pgSz w:w="5760" w:h="5760"/>"#;
const LEAD: &str = "00002600";
const KEPT: &str = "00002700";
const RELEASED_PAGES: usize = 4;
const RELEASED_DIGEST: u64 = 0x6302_e410_5b28_5ae3;
const MARGIN_CHANGE_PAGES: usize = 4;
const MARGIN_CHANGE_DIGEST: u64 = 0x1ac3_1774_7ad7_87fa;

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

/// Three columns of equal widths that the section declares unequal, with the
/// picture over the second one.
fn unequal_marked_columns(lead: usize) -> Vec<u8> {
    let section = format!(
        r#"<w:sectPr><w:type w:val="continuous"/>{PAGE}{}<w:cols w:num="3" w:space="360" w:equalWidth="0"><w:col w:w="1440" w:space="360"/><w:col w:w="1440" w:space="360"/><w:col w:w="1440"/></w:cols></w:sectPr>"#,
        margins(360, 360)
    );
    package(152, lead, &section)
}

/// A shape placed by page side and an inline picture in a continuous section.
fn page_side_shape_with_inline_picture() -> Vec<u8> {
    let section = format!(
        r#"<w:sectPr><w:type w:val="continuous"/>{PAGE}{}</w:sectPr>"#,
        margins(360, 360)
    );
    let body = format!(
        "{}{}{}{}{section}",
        first_section(),
        para("00002900", &format!("{}{INSIDE_SHAPE}", run("Anchor "))),
        para(
            "00002901",
            &format!("{}{}", run("Inline "), image("image", "inline"))
        ),
        filler_paragraphs(0x2A00, 12),
    );
    Package::new(&body)
        .rel("image", "image", "media/image1.png")
        .bytes()
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

fn text_end(engine: &EngineSession, id: &str) -> u32 {
    engine
        .doc()
        .paragraph_index("body")
        .unwrap()
        .para_span(id)
        .unwrap()
        .1
        - 1
}

/// The measured body of `bytes` as an open lays it out, before wrapping settles.
fn measured_body(engine: &EngineSession, bytes: &[u8]) -> crate::engine::PreparedRegionLayout {
    let font = docx_layout::register_measure_font_bytes(lowering_pages::FONT).unwrap();
    let request = lowering_pages::region_request(engine, bytes, font).to_string();
    engine.set_local_lowering(true);
    let mut prepared = engine
        .prepare_region_layout(&request, None, RelayoutTrigger::Open)
        .unwrap();
    prepared.measure(usize::MAX).unwrap();
    prepared
}

fn picture_index(input: &LayoutInput) -> Option<usize> {
    input.measured.iter().position(|measured| {
        matches!(
            &measured.block,
            LayoutBlock::Paragraph(paragraph)
                if paragraph.runs.iter().any(|run| matches!(run, Run::Image(_)))
        )
    })
}

/// Every block of `prepared` measured afresh in one pass, with no picture
/// held to a horizontal frame.
fn fresh_measures(prepared: &crate::engine::PreparedRegionLayout) -> Vec<serde_json::Value> {
    let input = &prepared.input;
    let mut blocks = input
        .measured
        .iter()
        .map(|measured| measured.block.clone())
        .collect::<Vec<_>>();
    let (widths, frames) =
        crate::engine::region_table_measurement_frames(blocks.iter(), input, &prepared.regions);
    let geometry = crate::engine::initial_float_page_geometry(input, &prepared.regions);
    docx_layout::measure_blocks::measure_blocks_with_column_frames(
        &mut blocks,
        &widths,
        &frames,
        &prepared.measurement,
        Some(&geometry),
        &BTreeMap::new(),
        &BTreeMap::new(),
    )
    .unwrap()
    .iter()
    .map(|measure| serde_json::to_value(measure).unwrap())
    .collect()
}

/// Checks the open session against a fresh session opened on the same
/// document: the frame, and the measured blocks and options it paginates.
fn assert_input_matches_fresh(
    host: &mut CommandHost<'_>,
    client_id: u64,
    bytes: &[u8],
    label: &str,
) {
    host.assert_matches_cold(label);
    let cold = cold_session(host.engine, client_id, Some((client_id, bytes)));
    cold.set_local_lowering(true);
    cold.layout_regions_for_trigger(&host.request, None, RelayoutTrigger::Open)
        .unwrap();
    let inputs = |engine: &EngineSession| {
        serde_json::from_str::<serde_json::Value>(&engine.retained_kernel_inputs_json().unwrap())
            .unwrap()
    };
    assert_eq!(inputs(host.engine), inputs(&cold), "{label}");
}

#[test]
fn a_cycling_page_side_shape_beside_an_inline_picture_keeps_the_convergence_error() {
    let fonts = docx_layout::MeasureFonts::default();
    let _scope = fonts.enter();
    let bytes = page_side_shape_with_inline_picture();
    let engine = seeded_engine(9806, &bytes);
    let mut prepared = measured_body(&engine, &bytes);
    let page_side = prepared
        .input
        .measured
        .iter()
        .any(|measured| match &measured.block {
            LayoutBlock::Shape(shape) => crate::engine::wraps_by_page_side(shape),
            _ => false,
        });
    assert!(page_side);
    assert!(picture_index(&prepared.input).is_some());
    let laid_out = crate::engine::stabilize_flow_wrapping(
        &mut prepared.input,
        &prepared.regions,
        &prepared.measurement,
        |input, phase| {
            let mut layout = docx_layout::place::layout_document(input)?;
            for fragment in layout
                .pages
                .iter_mut()
                .flat_map(|page| page.fragments.iter_mut())
            {
                if let docx_layout::types::Fragment::Shape(shape) = fragment {
                    shape.wrap_offset_x = Some(if phase % 2 == 0 { 12.0 } else { 36.0 });
                }
            }
            Ok(layout)
        },
    );
    assert!(matches!(
        laid_out,
        Err(docx_layout::LayoutError::Invalid(message))
            if message == "anchored shape wrapping did not converge"
    ));
}

/// Places every fragment of the `picture` paragraph in the column at `column_x`.
fn put_in_column(layout: &mut docx_layout::types::Layout, picture: &LayoutBlock, column_x: f64) {
    let LayoutBlock::Paragraph(picture) = picture else {
        panic!("the picture block is a paragraph");
    };
    for fragment in layout
        .pages
        .iter_mut()
        .flat_map(|page| page.fragments.iter_mut())
    {
        if let docx_layout::types::Fragment::Paragraph(fragment) = fragment
            && fragment.block_id == picture.id
        {
            fragment.horizontal_frame = Some(docx_layout::types::HorizontalFrame {
                margin_left: 24.0,
                content_width: 336.0,
                column_x,
                column_width: 96.0,
            });
        }
    }
}

#[test]
fn a_picture_pinned_on_the_last_wrap_phase_still_lays_out() {
    let fonts = docx_layout::MeasureFonts::default();
    let _scope = fonts.enter();
    let bytes = equal_columns(3, 60);
    let engine = seeded_engine(9814, &bytes);
    let mut prepared = measured_body(&engine, &bytes);
    let picture = picture_index(&prepared.input).unwrap();
    let block = prepared.input.measured[picture].block.clone();
    crate::engine::WRAP_PHASE_LIMIT.with(|limit| limit.set(Some(3)));
    let settled = crate::engine::stabilize_flow_wrapping(
        &mut prepared.input,
        &prepared.regions,
        &prepared.measurement,
        |input, phase| {
            let mut layout = docx_layout::place::layout_document(input)?;
            let column_x = if phase % 2 == 0 { 144.0 } else { 264.0 };
            put_in_column(&mut layout, &block, column_x);
            Ok(layout)
        },
    );
    crate::engine::WRAP_PHASE_LIMIT.with(|limit| limit.set(None));
    assert!(settled.unwrap());
}

#[test]
fn wrapping_that_keeps_moving_after_a_pin_ends_in_the_convergence_error() {
    let fonts = docx_layout::MeasureFonts::default();
    let _scope = fonts.enter();
    let bytes = page_side_shape_with_inline_picture();
    let engine = seeded_engine(9815, &bytes);
    let mut prepared = measured_body(&engine, &bytes);
    let picture = picture_index(&prepared.input).unwrap();
    let block = prepared.input.measured[picture].block.clone();
    let mut moved = 40.0;
    let laid_out = crate::engine::stabilize_flow_wrapping(
        &mut prepared.input,
        &prepared.regions,
        &prepared.measurement,
        |input, phase| {
            let mut layout = docx_layout::place::layout_document(input)?;
            let column_x = if phase % 2 == 0 { 144.0 } else { 264.0 };
            put_in_column(&mut layout, &block, column_x);
            let offset = match phase {
                0 | 2 => 12.0,
                1 => 36.0,
                _ => {
                    moved += 1.0;
                    moved
                }
            };
            for fragment in layout
                .pages
                .iter_mut()
                .flat_map(|page| page.fragments.iter_mut())
            {
                if let docx_layout::types::Fragment::Shape(shape) = fragment {
                    shape.wrap_offset_x = Some(offset);
                }
            }
            Ok(layout)
        },
    );
    assert!(matches!(
        laid_out,
        Err(docx_layout::LayoutError::Invalid(message))
            if message == "anchored shape wrapping did not converge"
    ));
}

#[test]
fn flow_after_a_settled_picture_measures_as_a_fresh_layout_does() {
    let fonts = docx_layout::MeasureFonts::default();
    let _scope = fonts.enter();
    let bytes = equal_columns(3, 60);
    let engine = seeded_engine(9807, &bytes);
    let mut prepared = measured_body(&engine, &bytes);
    let picture = picture_index(&prepared.input).unwrap();
    let settled = crate::engine::stabilize_flow_wrapping(
        &mut prepared.input,
        &prepared.regions,
        &prepared.measurement,
        |input, _| docx_layout::place::layout_document(input),
    );
    assert!(settled.unwrap());
    let fresh = fresh_measures(&prepared);
    for (index, measured) in prepared.input.measured.iter().enumerate().skip(picture) {
        assert_eq!(
            serde_json::to_value(&measured.measure).unwrap(),
            fresh[index],
            "block {index}"
        );
    }
}

#[test]
fn equal_column_picture_edits_leave_the_input_a_fresh_open_builds() {
    let fonts = docx_layout::MeasureFonts::default();
    let _scope = fonts.enter();
    for (client_id, columns, lead) in [(9808, 2, 9), (9809, 3, 60)] {
        let bytes = equal_columns(columns, lead);
        let engine = seeded_engine(client_id, &bytes);
        let mut host = opened(&engine, &bytes).seeded(client_id, bytes.clone());
        let check = |host: &mut CommandHost<'_>, label: &str| {
            assert_input_matches_fresh(
                host,
                client_id,
                &bytes,
                &format!("{columns} columns: {label}"),
            );
        };
        check(&mut host, "open");
        let at = lead_end(&engine);
        let typed = format!(" {}", words(40, 900));
        host.command("body", at, at, Some(&typed));
        check(&mut host, "typing before the picture");
        host.command("body", at, at + typed.chars().count() as u32, None);
        check(&mut host, "shrinking the lead");
        let grown = format!(" {}", words(20, 900));
        host.command("body", at, at, Some(&grown));
        check(&mut host, "growing the lead");
        let split = at + grown.chars().count() as u32 / 2;
        engine
            .split_resident_paragraph(crate::Position::new("body", split))
            .unwrap();
        host.relayout(Some("body"));
        check(&mut host, "enter in the lead");
    }
}

#[test]
fn edits_in_the_preceding_section_leave_the_input_a_fresh_open_builds() {
    let fonts = docx_layout::MeasureFonts::default();
    let _scope = fonts.enter();
    for (client_id, columns, lead) in [(9810, 2, 9), (9811, 3, 60)] {
        let bytes = equal_columns(columns, lead);
        let engine = seeded_engine(client_id, &bytes);
        let mut host = opened(&engine, &bytes).seeded(client_id, bytes.clone());
        for count in [6, 14, 30] {
            let at = text_end(&engine, "00002000");
            host.command("body", at, at, Some(&format!(" {}", words(count, 700))));
            assert_input_matches_fresh(
                &mut host,
                client_id,
                &bytes,
                &format!("{columns} columns: {count} words typed in the section above"),
            );
        }
    }
}

#[test]
fn a_picture_cycling_in_unequal_columns_keeps_the_convergence_error() {
    let fonts = docx_layout::MeasureFonts::default();
    let _scope = fonts.enter();
    let font = docx_layout::register_measure_font_bytes(lowering_pages::FONT).unwrap();
    for (client_id, bytes, settles) in [
        (9812, equal_columns(3, 60), true),
        (9813, unequal_marked_columns(60), false),
    ] {
        let engine = seeded_engine(client_id, &bytes);
        let request = lowering_pages::region_request(&engine, &bytes, font).to_string();
        engine.set_local_lowering(true);
        let laid_out = engine.layout_document_with_regions_retained(&request);
        if settles {
            laid_out.unwrap();
        } else {
            assert!(
                laid_out
                    .unwrap_err()
                    .contains("anchored shape wrapping did not converge")
            );
        }
    }
}
