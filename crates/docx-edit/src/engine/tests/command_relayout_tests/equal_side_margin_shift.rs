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
const NEW_LEFT_MARGIN: f64 = 48.0;

fn shifted_side_margins() -> Vec<u8> {
    let section = format!(
        r#"<w:sectPr><w:type w:val="continuous"/>{PAGE}{}</w:sectPr>"#,
        margins(720, 1440)
    );
    let body = format!(
        "{}{}{}{section}",
        first_section(1440, 720),
        para(
            PICTURE,
            &format!("{}{}", picture(6, 24, 60), run(&words(30, 300)))
        ),
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
                if paragraph.runs.iter().any(|run| matches!(run, Run::Image(_))) =>
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
