#![cfg(feature = "fonts-latin")]

mod common;

use std::collections::{BTreeMap, BTreeSet, HashMap};

use betteroffice_docx::fonts::{FontLoading, FontManager, FontSource, LastResortOffice};
use betteroffice_docx::{Document, LayoutInput, Primitive};
use docx_layout::measure_blocks::{FontRequirementCollector, MeasurementConfig, measure_block};
use docx_layout::types::{Fragment, Layout};
use ooxml_text::family_metrics::{family_advance, family_metrics};
use ooxml_text::word_fonts::requested_line_metrics;
use ooxml_text::{FontId, FontStore};
use serde_json::json;

const LINE: &str = "Hamburgevons0123456789";
const SENTENCE: &str =
    "Substituted faces measure with the widths and line metrics of the family the document names. ";
const COLUMN_WIDTH: f64 = 624.0;

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

fn input(family: &str) -> LayoutInput {
    let paragraph = |id: u64, text: &str| {
        json!({"kind": "paragraph", "id": id, "pmStart": 0, "pmEnd": text.len() + 2,
               "runs": [{"kind": "text", "text": text, "fontFamily": family, "fontSize": 11,
                         "pmStart": 1, "pmEnd": text.len() + 1}]})
    };
    let unmeasured = json!({"kind": "paragraph", "lines": [], "totalHeight": 0});
    serde_json::from_value(json!({
        "measured": [
            {"block": paragraph(1, LINE), "measure": unmeasured},
            {"block": paragraph(2, &SENTENCE.repeat(2)), "measure": unmeasured}
        ],
        "options": {"pageSize": {"w": 816, "h": 1056},
                    "margins": {"top": 96, "right": 96, "bottom": 96, "left": 96}}
    }))
    .unwrap()
}

fn chain_key(family: &str, bold: bool, italic: bool) -> String {
    format!(
        "{}|{}|{}",
        family.trim().to_lowercase(),
        u8::from(bold),
        u8::from(italic)
    )
}

fn pitches(layout: &Layout) -> BTreeMap<String, (f64, usize)> {
    let mut blocks: BTreeMap<String, (f64, usize)> = BTreeMap::new();
    for fragment in layout.pages.iter().flat_map(|page| &page.fragments) {
        if let Fragment::Paragraph(paragraph) = fragment {
            let block = serde_json::to_string(&paragraph.block_id).unwrap();
            let entry = blocks.entry(block).or_default();
            entry.0 += paragraph.height;
            entry.1 += paragraph.to_line - paragraph.from_line;
        }
    }
    blocks
        .into_iter()
        .map(|(block, (height, lines))| (block, (height / lines.max(1) as f64, lines)))
        .collect()
}

#[test]
fn substituted_faces_measure_with_the_requested_family_metrics() {
    let manager = FontManager::new(&FontSource::default(), FontLoading::default()).unwrap();
    let document = Document::open(&common::sample_docx()).unwrap();
    let package = json!({
        "settings": { "compatibilityFlags": &document.model().settings.compatibility_flags }
    });
    let mut store = FontStore::new();
    let mut bases: HashMap<(String, bool, bool), FontId> = HashMap::new();
    let mut checked = Vec::new();
    let mut mismatches = Vec::new();
    for family in table_families() {
        let mut reference = input(&family);
        let mut collector = FontRequirementCollector::default();
        collector.collect(
            reference.measured.iter().map(|measured| &measured.block),
            "Calibri",
        );
        let mut chains = BTreeMap::new();
        let mut head = None;
        let mut substituted = false;
        for request in collector.finish().into_values() {
            let faces = manager.resolve(
                &request.family,
                request.bold,
                request.italic,
                &request.scripts,
                LastResortOffice::Word,
            );
            let requested = request.family.trim().to_lowercase() == family
                && !request.bold
                && !request.italic;
            let real = faces
                .first()
                .is_some_and(|face| face.family.to_lowercase() == family);
            if requested && real {
                continue;
            }
            let mut chain: Vec<u32> = Vec::new();
            for face in faces {
                let base = *bases
                    .entry((face.family.clone(), face.bold, face.italic))
                    .or_insert_with(|| store.register_slice(&face.load().unwrap()).unwrap());
                let measured_as = face.measure_as.as_deref().unwrap_or(&request.family);
                let view = store.register_named_substitute(base, measured_as).unwrap();
                if requested {
                    substituted |= view != base;
                    head = head.or(Some(view));
                }
                if !chain.contains(&view.to_u32()) {
                    chain.push(view.to_u32());
                }
            }
            chains.insert(chain_key(&request.family, request.bold, request.italic), chain);
        }
        if !substituted {
            continue;
        }
        checked.push(family.clone());

        let mut config: MeasurementConfig =
            serde_json::from_value(docx_edit::layout_request::measurement(&package)).unwrap();
        config.font_chains = chains;
        docx_layout::with_measure_font_store(&mut store, || {
            for measured in &mut reference.measured {
                measured.measure =
                    measure_block(&mut measured.block, COLUMN_WIDTH, &config).unwrap();
            }
        });
        let expected = docx_layout::compute_layout_input(&mut reference).unwrap();
        let actual = document.layout(input(&family)).unwrap();

        if let Some(widths) = head.and_then(|id| store.family_widths(id).unwrap()) {
            let mut compared = 0;
            let mut wrong = None;
            let runs = actual
                .display_list
                .pages
                .iter()
                .flat_map(|page| &page.primitives)
                .filter_map(|primitive| match primitive {
                    Primitive::GlyphRun(run) if !run.text.is_empty() && LINE.contains(&run.text) => {
                        Some(run)
                    }
                    _ => None,
                });
            for run in runs {
                for glyph in &run.glyphs {
                    let Some(character) = run
                        .text
                        .get(glyph.cluster as usize..)
                        .and_then(|rest| rest.chars().next())
                    else {
                        continue;
                    };
                    let Some(advance) = family_advance(widths, character) else {
                        continue;
                    };
                    let wanted = f64::from((advance * 1000.0).round()) * run.size / 1000.0;
                    compared += 1;
                    if wrong.is_none() && (glyph.advance - wanted).abs() > 0.02 {
                        wrong = Some(format!(
                            "{character:?} advances {:.3}px, the width table gives {wanted:.3}px",
                            glyph.advance
                        ));
                    }
                }
            }
            if compared == 0 {
                mismatches.push(format!("{family}: no glyph of {LINE:?} was laid out"));
            }
            if let Some(wrong) = wrong {
                mismatches.push(format!("{family}: {wrong}"));
            }
        }

        let actual_pitches = pitches(&actual.layout);
        for (block, (pitch, lines)) in pitches(&expected) {
            match actual_pitches.get(&block) {
                Some((actual_pitch, _)) if (actual_pitch - pitch).abs() <= 0.01 => {}
                Some((actual_pitch, actual_lines)) => mismatches.push(format!(
                    "{family}: block {block} line pitch {actual_pitch:.3}px over {actual_lines} lines, \
                     the requested metrics give {pitch:.3}px over {lines} lines"
                )),
                None => mismatches.push(format!("{family}: block {block} was not laid out")),
            }
        }
    }
    println!(
        "checked {} substituted families: {}",
        checked.len(),
        checked.join(", ")
    );
    assert!(!checked.is_empty());
    assert!(
        mismatches.is_empty(),
        "{} mismatches over {} families:\n{}",
        mismatches.len(),
        checked.len(),
        mismatches.join("\n")
    );
}
