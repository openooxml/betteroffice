//! A continuous section that moves both side margins by the same amount, so
//! its content keeps the width of the section above, with a picture anchored
//! to the page inside the new left margin beside its first paragraph.

use super::equal_column_pictures::{
    PAGE, assert_input_matches_fresh, first_section, margins, opened, picture, text_end, words,
};
use super::*;
use crate::seed::fixture::{Package, para, run};
use docx_layout::types::{BlockExtent, Fragment};

const FIRST: &str = "00002000";
const PICTURE: &str = "00003000";
const BELOW_PICTURE: &str = "00003080";
const NEW_LEFT_MARGIN: f64 = 48.0;

fn shifted_side_margins() -> Vec<u8> {
    side_margins_below(720, 1440)
}

/// The picture's continuous section, with side margins of `left` and `right`
/// twips, below a section whose side margins are 1440 and 720 twips.
fn side_margins_below(left: u32, right: u32) -> Vec<u8> {
    section_below(left, right, &picture(6, 24, 60))
}

/// [`side_margins_below`] with a picture whose wrap edge lies past the text's
/// right edge under a right margin of 1440 twips and inside it under 720.
fn reaching_picture_below(left: u32, right: u32) -> Vec<u8> {
    section_below(left, right, &picture(310, 30, 60))
}

/// [`side_margins_below`] with a picture 270 px from the page's left edge
/// beside a long paragraph.
fn picture_beside_long_text_below(left: u32, right: u32) -> Vec<u8> {
    section_holding(
        left,
        right,
        &para(
            PICTURE,
            &format!("{}{}", picture(270, 30, 60), run(&"a ".repeat(240))),
        ),
    )
}

/// [`reaching_picture_below`] with the picture alone in a paragraph above the
/// text it reaches into.
fn picture_above_text_below(left: u32, right: u32) -> Vec<u8> {
    let paragraphs = format!(
        "{}{}",
        para(PICTURE, &picture(310, 30, 60)),
        para(BELOW_PICTURE, &run(&words(30, 300)))
    );
    section_holding(left, right, &paragraphs)
}

fn section_below(left: u32, right: u32, picture: &str) -> Vec<u8> {
    section_holding(
        left,
        right,
        &para(PICTURE, &format!("{picture}{}", run(&words(30, 300)))),
    )
}

fn section_holding(left: u32, right: u32, paragraphs: &str) -> Vec<u8> {
    let section = format!(
        r#"<w:sectPr><w:type w:val="continuous"/>{PAGE}{}</w:sectPr>"#,
        margins(left, right)
    );
    let body = format!(
        "{}{paragraphs}{}{section}",
        first_section(1440, 720),
        filler_paragraphs(0x3100, 12),
    );
    Package::new(&body)
        .rel("image", "image", "media/image1.png")
        .bytes()
}

/// The page index and x of the first line of the picture's paragraph.
fn first_picture_line(engine: &EngineSession) -> (usize, f64) {
    let pagination = engine.pagination.borrow();
    let (paragraph, extent) = pagination
        .input
        .as_ref()
        .unwrap()
        .measured
        .iter()
        .find_map(|measured| match (&measured.block, &measured.measure) {
            (LayoutBlock::Paragraph(paragraph), BlockExtent::Paragraph(extent))
                if paragraph
                    .runs
                    .iter()
                    .any(|run| matches!(run, Run::Image(_))) =>
            {
                Some((paragraph, extent))
            }
            _ => None,
        })
        .unwrap();
    pagination
        .layout
        .as_ref()
        .unwrap()
        .pages
        .iter()
        .enumerate()
        .find_map(|(page, placed)| {
            placed.fragments.iter().find_map(|fragment| match fragment {
                Fragment::Paragraph(fragment)
                    if fragment.block_id == paragraph.id && fragment.from_line == 0 =>
                {
                    Some((
                        page,
                        fragment.x + extent.lines[0].left_offset.unwrap_or(0.0),
                    ))
                }
                _ => None,
            })
        })
        .unwrap()
}

/// The width and side offsets of each measured line of the picture's
/// paragraph.
fn picture_lines(engine: &EngineSession) -> Vec<(f64, Option<f64>, Option<f64>)> {
    let pagination = engine.pagination.borrow();
    pagination
        .input
        .as_ref()
        .unwrap()
        .measured
        .iter()
        .find_map(|measured| match (&measured.block, &measured.measure) {
            (LayoutBlock::Paragraph(paragraph), BlockExtent::Paragraph(extent))
                if paragraph
                    .runs
                    .iter()
                    .any(|run| matches!(run, Run::Image(_))) =>
            {
                Some(
                    extent
                        .lines
                        .iter()
                        .map(|line| (line.width, line.left_offset, line.right_offset))
                        .collect(),
                )
            }
            _ => None,
        })
        .unwrap()
}

fn print_picture_lines(engine: &EngineSession, label: &str) {
    let lines = picture_lines(engine);
    println!();
    println!("LINES {label}: {} lines {lines:?}", lines.len());
}

fn assert_at_new_left_margin(engine: &EngineSession, label: &str) {
    let (page, x) = first_picture_line(engine);
    assert_eq!(page, 0, "{label}");
    assert!(
        (x - NEW_LEFT_MARGIN).abs() < 0.01,
        "{label}: first line at {x}"
    );
}

#[test]
fn text_beside_a_picture_in_the_shifted_left_margin_starts_at_that_margin() {
    let fonts = docx_layout::MeasureFonts::default();
    let _scope = fonts.enter();
    let bytes = shifted_side_margins();
    let engine = seeded_engine(9820, &bytes);
    opened(&engine, &bytes);
    assert_at_new_left_margin(&engine, "open");
}

#[test]
fn edits_above_a_shifted_side_margin_section_leave_the_input_a_fresh_open_builds() {
    let fonts = docx_layout::MeasureFonts::default();
    let _scope = fonts.enter();
    let client_id = 9821;
    let bytes = shifted_side_margins();
    let engine = seeded_engine(client_id, &bytes);
    let mut host = opened(&engine, &bytes).seeded(client_id, bytes.clone());
    for count in [6, 14] {
        let label = format!("{count} words typed in the section above");
        let at = text_end(&engine, FIRST);
        host.command("body", at, at, Some(&format!(" {}", words(count, 700))));
        assert_input_matches_fresh(&mut host, client_id, &bytes, &label);
        assert_at_new_left_margin(&engine, &label);
    }
}

/// Sets the picture's section side margins to `left` and `right` twips and
/// lays out with the regions of the package `document` builds with them.
fn set_side_margins(
    host: &mut CommandHost<'_>,
    document: fn(u32, u32) -> Vec<u8>,
    left: u32,
    right: u32,
) {
    let engine = host.engine;
    let margins = format!(r#"{{"marginLeft":{left},"marginRight":{right}}}"#);
    assert!(
        engine
            .doc()
            .set_section_margins(&crate::EditCtx::local("", ""), "final", &margins)
            .unwrap()
    );
    let font = docx_layout::register_measure_font_bytes(lowering_pages::FONT).unwrap();
    let bytes = document(left, right);
    host.request = lowering_pages::region_request(engine, &bytes, font).to_string();
    host.relayout(Some("body"));
}

#[test]
fn side_margins_set_to_the_section_above_lay_out_like_a_fresh_open() {
    let fonts = docx_layout::MeasureFonts::default();
    let _scope = fonts.enter();
    let client_id = 9822;
    let bytes = reaching_picture_below(720, 1440);
    let engine = seeded_engine(client_id, &bytes);
    let mut host = opened(&engine, &bytes).seeded(client_id, bytes.clone());
    assert_at_new_left_margin(&engine, "open");
    set_side_margins(&mut host, reaching_picture_below, 1440, 720);
    assert_input_matches_fresh(&mut host, client_id, &bytes, "margins of the section above");
}

#[test]
fn side_margins_shifted_from_the_section_above_lay_out_like_a_fresh_open() {
    let fonts = docx_layout::MeasureFonts::default();
    let _scope = fonts.enter();
    let client_id = 9823;
    let bytes = reaching_picture_below(1440, 720);
    let engine = seeded_engine(client_id, &bytes);
    let mut host = opened(&engine, &bytes).seeded(client_id, bytes.clone());
    set_side_margins(&mut host, reaching_picture_below, 720, 1440);
    assert_input_matches_fresh(&mut host, client_id, &bytes, "shifted margins");
    assert_at_new_left_margin(&engine, "shifted margins");
}

#[test]
fn side_margins_set_to_the_section_above_lay_out_text_below_a_picture_like_a_fresh_open() {
    let fonts = docx_layout::MeasureFonts::default();
    let _scope = fonts.enter();
    let client_id = 9824;
    let bytes = picture_above_text_below(720, 1440);
    let engine = seeded_engine(client_id, &bytes);
    let mut host = opened(&engine, &bytes).seeded(client_id, bytes.clone());
    set_side_margins(&mut host, picture_above_text_below, 1440, 720);
    assert_input_matches_fresh(&mut host, client_id, &bytes, "margins of the section above");
}

#[test]
fn side_margins_shifted_from_the_section_above_lay_out_text_below_a_picture_like_a_fresh_open() {
    let fonts = docx_layout::MeasureFonts::default();
    let _scope = fonts.enter();
    let client_id = 9825;
    let bytes = picture_above_text_below(1440, 720);
    let engine = seeded_engine(client_id, &bytes);
    let mut host = opened(&engine, &bytes).seeded(client_id, bytes.clone());
    set_side_margins(&mut host, picture_above_text_below, 720, 1440);
    assert_input_matches_fresh(&mut host, client_id, &bytes, "shifted margins");
}

#[test]
fn side_margins_set_back_after_frame_narrowing_lay_out_like_a_fresh_open() {
    let fonts = docx_layout::MeasureFonts::default();
    let _scope = fonts.enter();
    let client_id = 9826;
    let bytes = picture_beside_long_text_below(720, 1440);
    let engine = seeded_engine(client_id, &bytes);
    let mut host = opened(&engine, &bytes).seeded(client_id, bytes.clone());
    print_picture_lines(&engine, "before");
    set_side_margins(&mut host, picture_beside_long_text_below, 1440, 720);
    print_picture_lines(&engine, "after");
    let fresh_bytes = picture_beside_long_text_below(1440, 720);
    let fresh = seeded_engine(9827, &fresh_bytes);
    opened(&fresh, &fresh_bytes);
    print_picture_lines(&fresh, "fresh");
    assert_input_matches_fresh(&mut host, client_id, &bytes, "margins of the section above");
}
