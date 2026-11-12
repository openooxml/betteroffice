#![cfg(all(feature = "fonts-latin", feature = "fonts-cjk"))]

mod common;

use std::collections::{BTreeSet, HashMap};

use betteroffice_docx::fonts::{FontLoading, FontManager, FontSource, LastResortOffice};
use betteroffice_docx::{Document, LayoutInput, LayoutResult, Primitive};
use docx_layout::types::{Fragment, Layout};
use ooxml_text::family_metrics::family_metrics;
use ooxml_text::word_fonts::requested_line_metrics;
use serde_json::{Value, json};

const LATIN: &str = "Hamburgevons0123456789";
const KANA: &str = "ひらがなとカタカナ";
const JAPANESE: &str = "日本語の文章です。";
/// Word's `ＭＳ 明朝` line box at 14px: 14.126px ascent over 4.074px descent.
const MINCHO_PITCH_AT_14PX: f64 = 14.126 + 4.074;

fn quoted(source: &'static str) -> impl Iterator<Item = &'static str> {
    source
        .split("#[cfg(test)]")
        .next()
        .unwrap_or_default()
        .lines()
        .map(|line| line.split("//").next().unwrap_or_default())
        .filter(|code| code.matches('"').count().is_multiple_of(2))
        .flat_map(|code| code.split('"').skip(1).step_by(2))
}

fn table_families() -> BTreeSet<String> {
    let widths = quoted(include_str!("../../ooxml-text/src/family_metrics.rs"))
        .filter(|name| family_metrics(name, false, false).is_some());
    let lines = quoted(include_str!("../../ooxml-text/src/word_fonts.rs"))
        .filter(|name| requested_line_metrics(name).is_some());
    widths.chain(lines).map(str::to_owned).collect()
}

fn input(mut run: Value) -> LayoutInput {
    let length = run["text"].as_str().unwrap_or_default().chars().count();
    run["kind"] = json!("text");
    run["pmStart"] = json!(1);
    run["pmEnd"] = json!(length + 1);
    serde_json::from_value(json!({
        "measured": [{
            "block": {"kind": "paragraph", "id": 1, "pmStart": 0, "pmEnd": length + 2, "runs": [run]},
            "measure": {"kind": "paragraph", "lines": [], "totalHeight": 0}
        }],
        "options": {"pageSize": {"w": 816, "h": 1056},
                    "margins": {"top": 96, "right": 96, "bottom": 96, "left": 96}}
    }))
    .unwrap()
}

fn pitch(layout: &Layout) -> Option<f64> {
    let (height, lines) = layout
        .pages
        .iter()
        .flat_map(|page| &page.fragments)
        .filter_map(|fragment| match fragment {
            Fragment::Paragraph(paragraph) => {
                Some((paragraph.height, paragraph.to_line - paragraph.from_line))
            }
            _ => None,
        })
        .fold((0.0, 0), |(height, lines), (h, l)| (height + h, lines + l));
    (lines > 0).then(|| height / lines as f64)
}

/// Every laid-out glyph as (character, font id, glyph id, advance).
fn glyphs(result: &LayoutResult) -> Vec<(char, u32, u32, f64)> {
    result
        .display_list
        .pages
        .iter()
        .flat_map(|page| &page.primitives)
        .filter_map(|primitive| match primitive {
            Primitive::GlyphRun(run) => Some(run),
            _ => None,
        })
        .flat_map(|run| {
            run.glyphs.iter().filter_map(move |glyph| {
                let character = run.text.get(glyph.cluster as usize..)?.chars().next()?;
                Some((character, run.font_id, glyph.id, glyph.advance))
            })
        })
        .collect()
}

#[test]
fn a_family_without_a_bundled_face_keeps_its_own_line_box_and_latin_face_next_to_cjk_text() {
    let manager = FontManager::new(&FontSource::default(), FontLoading::default()).unwrap();
    let package = common::sample_docx();
    let japanese = ["cjk-jp".to_owned()];
    let mut checked = Vec::new();
    let mut mismatches = Vec::new();
    for family in table_families() {
        let own = manager.resolve(&family, false, false, &[], LastResortOffice::Word);
        let scripted = manager.resolve(&family, false, false, &japanese, LastResortOffice::Word);
        let script_face_first = scripted.first().is_some_and(|face| {
            !own.iter().any(|known| {
                known.family == face.family
                    && known.bold == face.bold
                    && known.italic == face.italic
            })
        });
        if !script_face_first {
            continue;
        }
        checked.push(family.clone());

        let latin = Document::open(&package)
            .unwrap()
            .layout(input(
                json!({"text": LATIN, "fontFamily": family.as_str(), "fontSize": 11}),
            ))
            .unwrap();
        let mixed = Document::open(&package)
            .unwrap()
            .layout(input(json!({"text": format!("{LATIN}{KANA}"),
                                 "fontFamily": family.as_str(), "fontSize": 11})))
            .unwrap();

        match (pitch(&latin.layout), pitch(&mixed.layout)) {
            (Some(expected), Some(actual)) if (actual - expected).abs() > 0.01 => {
                mismatches.push(format!(
                    "{family}: line pitch {actual:.3}px with Japanese text, \
                     {expected:.3}px with its Latin text alone"
                ));
            }
            (Some(_), Some(_)) => {}
            _ => mismatches.push(format!("{family}: the paragraph was not laid out")),
        }

        let own_glyphs: HashMap<char, (u32, f64)> = glyphs(&latin)
            .into_iter()
            .map(|(character, _, id, advance)| (character, (id, advance)))
            .collect();
        let mut latin_fonts = BTreeSet::new();
        let mut kana_fonts = BTreeSet::new();
        let mut wrong = None;
        for (character, font, id, advance) in glyphs(&mixed) {
            if KANA.contains(character) {
                kana_fonts.insert(font);
                continue;
            }
            let Some(&(own_id, own_advance)) = own_glyphs.get(&character) else {
                continue;
            };
            latin_fonts.insert(font);
            // The last Latin letter carries the Latin to East Asian spacing.
            let spaced = LATIN.ends_with(character);
            let advance_differs = !spaced && (advance - own_advance).abs() > 0.02;
            if wrong.is_none() && (id != own_id || advance_differs) {
                wrong = Some(format!(
                    "{family}: {character:?} draws glyph {id} advancing {advance:.3}px, \
                     the family's own face draws glyph {own_id} advancing {own_advance:.3}px"
                ));
            }
        }
        if latin_fonts.is_empty() || kana_fonts.is_empty() {
            mismatches.push(format!(
                "{family}: {} Latin and {} Japanese fonts drew glyphs",
                latin_fonts.len(),
                kana_fonts.len()
            ));
        }
        if let Some(shared) = latin_fonts.intersection(&kana_fonts).next() {
            mismatches.push(format!(
                "{family}: Latin and Japanese text both draw with font {shared}"
            ));
        }
        mismatches.extend(wrong);
    }

    let century = Document::open(&package)
        .unwrap()
        .layout(input(json!({"text": JAPANESE, "fontFamily": "Century",
                             "fontSlots": {"eastAsia": "ＭＳ 明朝"},
                             "language": {"eastAsia": "ja-JP"}, "fontSize": 10.5})))
        .unwrap();
    checked.push("Century with ＭＳ 明朝".to_owned());
    match pitch(&century.layout) {
        Some(actual) if (actual - MINCHO_PITCH_AT_14PX).abs() <= 0.01 => {}
        Some(actual) => mismatches.push(format!(
            "Century with ＭＳ 明朝: line pitch {actual:.3}px, \
             the ＭＳ 明朝 row gives {MINCHO_PITCH_AT_14PX:.3}px"
        )),
        None => mismatches.push("Century with ＭＳ 明朝: the paragraph was not laid out".into()),
    }

    println!("checked {} rows: {}", checked.len(), checked.join(", "));
    assert!(checked.len() > 1);
    assert!(
        mismatches.is_empty(),
        "{} mismatches over {} rows:\n{}",
        mismatches.len(),
        checked.len(),
        mismatches.join("\n")
    );
}
