#[allow(dead_code)]
#[path = "support/page_fixture.rs"]
mod fixture;

use std::collections::{BTreeMap, BTreeSet};

use docx_edit::structured::{
    PageAnnotations, PageExportOptions, PageMarkdownOptions, RevisionView, StorySelection,
    render_docx_markdown_with_pages,
};
use docx_edit::{EditCtx, FormatPolicy, Position};
use serde_json::Value;

const NS: &str = concat!(
    r#"xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" "#,
    r#"xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" "#,
    r#"xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml""#
);
const NARROW_PAGE: &str = r#"<w:pgSz w:w="8400" w:h="6000"/><w:pgMar w:top="720" w:right="720" w:bottom="720" w:left="720" w:header="300" w:footer="300" w:gutter="0"/>"#;
/// Room left for a line that did not fit at a page bottom.
const LINE_SLACK: f64 = 24.0;

fn words(note: usize, count: usize) -> String {
    (0..count)
        .map(|index| format!("n{note}t{index}"))
        .collect::<Vec<_>>()
        .join(" ")
}

fn reference(id: usize) -> String {
    format!(r#"<w:r><w:endnoteReference w:id="{id}"/></w:r>"#)
}

/// A package whose body is `body` and whose endnotes `1..` hold `notes`, one paragraph each: a
/// note starting with `<` is the paragraph's own content, any other its text.
fn endnotes_docx(body: &str, notes: &[String], settings: Option<&str>) -> Vec<u8> {
    let mut xml = format!(
        concat!(
            r#"<w:endnotes {}><w:endnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p>"#,
            r#"</w:endnote><w:endnote w:type="continuationSeparator" w:id="0"><w:p><w:r>"#,
            r#"<w:continuationSeparator/></w:r></w:p></w:endnote>"#
        ),
        NS
    );
    for (index, text) in notes.iter().enumerate() {
        let content = if text.starts_with('<') {
            text.clone()
        } else {
            fixture::r(text)
        };
        xml.push_str(&format!(
            r#"<w:endnote w:id="{}">{}</w:endnote>"#,
            index + 1,
            fixture::p(&format!("{:08X}", 0x5000_0000 + index), &content)
        ));
    }
    xml.push_str("</w:endnotes>");
    let mut parts = fixture::page_parts(false);
    for (name, bytes) in &mut parts {
        match name.as_str() {
            "word/document.xml" => {
                *bytes = format!(r#"<w:document {NS}><w:body>{body}</w:body></w:document>"#)
                    .into_bytes();
            }
            "word/endnotes.xml" => *bytes = xml.clone().into_bytes(),
            _ => {}
        }
    }
    if let Some(settings) = settings {
        parts = fixture::with_settings(parts, settings);
    }
    ooxml_opc::rezip_parts(&parts).unwrap()
}

/// A short body whose first twelve paragraphs reference endnotes that together fill pages.
fn long_endnotes() -> Vec<u8> {
    let mut body = String::new();
    for index in 0..40 {
        let note = if index < 12 {
            reference(index + 1)
        } else {
            String::new()
        };
        body.push_str(&fixture::p(
            &format!("{:08X}", 0x3000_0000 + index),
            &format!("{}{note}", fixture::r(&format!("Body line {index}"))),
        ));
    }
    body.push_str(&format!("<w:sectPr>{NARROW_PAGE}</w:sectPr>"));
    let notes: Vec<String> = (1..=12).map(|note| words(note, 70)).collect();
    endnotes_docx(&body, &notes, None)
}

fn pages(layout: &Value) -> &Vec<Value> {
    layout["layout"]["pages"].as_array().unwrap()
}

/// The bottom of the page's body text; flowed endnotes keep fragments that paint nothing.
fn fragment_bottom(page: &Value) -> Option<f64> {
    page["fragments"]
        .as_array()?
        .iter()
        .filter(|fragment| fragment["kind"] != "reserved")
        .map(|fragment| fragment["y"].as_f64().unwrap() + fragment["height"].as_f64().unwrap())
        .reduce(f64::max)
}

fn content_bottom(page: &Value) -> f64 {
    let margins = if page["bodyMargins"].is_object() {
        &page["bodyMargins"]
    } else {
        &page["margins"]
    };
    page["size"]["h"].as_f64().unwrap() - margins["bottom"].as_f64().unwrap()
}

fn endnote_area(page: &Value) -> Option<&Value> {
    page["noteAreas"]
        .as_array()?
        .iter()
        .find(|area| area["kind"] == "endnote")
}

fn area_bottom(area: &Value) -> f64 {
    area["y"].as_f64().unwrap() + area["height"].as_f64().unwrap()
}

fn note_ids(area: &Value) -> Vec<i64> {
    area["notes"]
        .as_array()
        .unwrap()
        .iter()
        .map(|note| note["id"].as_i64().unwrap())
        .collect()
}

fn laid_out(bytes: &[u8], client: u64) -> (docx_edit::EngineSession, String, Value) {
    let (engine, request) = fixture::laid_out(bytes, client);
    let layout =
        serde_json::from_str(&engine.layout_document_with_regions_json(&request).unwrap()).unwrap();
    (engine, request, layout)
}

#[test]
fn endnotes_taller_than_a_page_follow_the_body_instead_of_pushing_it_down() {
    let (_, _, layout) = laid_out(&long_endnotes(), 61);
    let pages = pages(&layout);
    let last_body = pages
        .iter()
        .rposition(|page| fragment_bottom(page).is_some())
        .unwrap();
    assert!(
        pages.len() >= last_body + 3,
        "the endnotes continue over pages of their own: {} pages, body ends on {last_body}",
        pages.len()
    );
    for (index, page) in pages.iter().enumerate() {
        let bottom = content_bottom(page);
        let filled = fragment_bottom(page)
            .into_iter()
            .chain(endnote_area(page).map(area_bottom))
            .fold(f64::MIN, f64::max);
        if index + 1 < pages.len() {
            assert!(
                filled >= bottom - LINE_SLACK,
                "page {index} ends at {filled}, its text area at {bottom}"
            );
        }
        if index < last_body {
            assert!(endnote_area(page).is_none(), "page {index} holds endnotes");
        }
        if index >= last_body {
            assert!(endnote_area(page).is_some(), "page {index} has no endnotes");
        }
        if index > last_body {
            assert!(
                fragment_bottom(page).is_none(),
                "page {index} holds body text"
            );
        }
    }
}

#[test]
fn flowed_endnotes_never_overlap_the_body_and_split_across_pages() {
    let (_, _, layout) = laid_out(&long_endnotes(), 62);
    let mut pages_of: BTreeMap<i64, Vec<usize>> = BTreeMap::new();
    let mut order = Vec::new();
    for (index, page) in pages(&layout).iter().enumerate() {
        let Some(area) = endnote_area(page) else {
            continue;
        };
        let y = area["y"].as_f64().unwrap();
        if let Some(body) = fragment_bottom(page) {
            assert!(
                y + 0.5 >= body,
                "page {index}: endnotes at {y}, body to {body}"
            );
        }
        assert!(area_bottom(area) <= content_bottom(page) + 0.5);
        for id in note_ids(area) {
            pages_of.entry(id).or_default().push(index);
            if order.last() != Some(&id) {
                order.push(id);
            }
        }
    }
    assert_eq!(order, (1..=12).collect::<Vec<_>>(), "notes flow in order");
    let split: Vec<_> = pages_of
        .iter()
        .filter(|(_, pages)| pages.len() > 1)
        .collect();
    assert!(!split.is_empty(), "some note continues on the next page");
    for (id, on) in split {
        assert_eq!(on.len(), 2, "endnote {id} on {on:?}");
        assert_eq!(on[1], on[0] + 1);
        for &index in on {
            let area = endnote_area(&pages(&layout)[index]).unwrap();
            let note = area["notes"]
                .as_array()
                .unwrap()
                .iter()
                .find(|note| note["id"] == *id)
                .unwrap();
            assert!(note["lineEnd"].as_u64().unwrap() > note["lineStart"].as_u64().unwrap());
        }
    }
}

#[test]
fn the_separator_opens_the_endnotes_and_continuation_pages_show_their_own() {
    let (engine, request, layout) = laid_out(&long_endnotes(), 63);
    let areas: Vec<(usize, &Value)> = pages(&layout)
        .iter()
        .enumerate()
        .filter_map(|(index, page)| endnote_area(page).map(|area| (index, area)))
        .collect();
    let (_, first) = areas[0];
    assert!(first.get("continuation").is_none());
    assert!(first["separator"]["height"].as_f64().unwrap() > 0.0);
    let opening = &first["notes"][0];
    assert!(
        opening.get("lineEnd").is_none()
            || opening["lineEnd"].as_u64() > opening["lineStart"].as_u64(),
        "the separator stands over a line of the first note"
    );
    for (index, area) in &areas[1..] {
        assert_eq!(area["continuation"], true, "page {index}");
        assert!(area["separator"]["height"].as_f64().unwrap() > 0.0);
    }

    let display: Value = serde_json::from_str(
        &engine
            .build_display_list_json(&layout_json(&engine, &request))
            .unwrap(),
    )
    .unwrap();
    let separator_width = |page: usize| {
        let area = display["pages"][page]["noteAreas"]
            .as_array()
            .unwrap()
            .iter()
            .find(|area| area["kind"] == "endnote")
            .unwrap();
        let line = &area["separatorPrimitives"][0];
        line["x2"].as_f64().unwrap() - line["x1"].as_f64().unwrap()
    };
    let opening = separator_width(areas[0].0);
    let continued = separator_width(areas[1].0);
    assert!(
        (continued - 3.0 * opening).abs() < 1.0,
        "a continuation separator spans the column: {continued} vs {opening}"
    );
}

fn layout_json(engine: &docx_edit::EngineSession, request: &str) -> String {
    engine.layout_document_with_regions_json(request).unwrap()
}

/// The text each endnote paints on each page, in paint order and without white space.
fn painted_notes(display: &Value) -> BTreeMap<String, Vec<(usize, String)>> {
    fn texts(value: &Value, out: &mut String) {
        match value {
            Value::Object(map) => {
                if let Some(Value::String(text)) = map.get("text") {
                    out.extend(text.chars().filter(|ch| !ch.is_whitespace()));
                }
                for (key, value) in map {
                    if key != "text" {
                        texts(value, out);
                    }
                }
            }
            Value::Array(items) => items.iter().for_each(|value| texts(value, out)),
            _ => {}
        }
    }
    let mut painted: BTreeMap<String, Vec<(usize, String)>> = BTreeMap::new();
    for (index, page) in display["pages"].as_array().unwrap().iter().enumerate() {
        for area in page["noteAreas"].as_array().into_iter().flatten() {
            for primitive in area["primitives"].as_array().into_iter().flatten() {
                let Some(group) = primitive["groupId"].as_str() else {
                    continue;
                };
                let parts = painted.entry(group.to_owned()).or_default();
                if parts.last().is_none_or(|(page, _)| *page != index) {
                    parts.push((index, String::new()));
                }
                texts(primitive, &mut parts.last_mut().unwrap().1);
            }
        }
    }
    painted
}

/// How often `word` stands as a whole word in `text`.
fn word_count(text: &str, word: &str) -> usize {
    text.match_indices(word)
        .filter(|(at, _)| {
            let before = text[..*at].chars().next_back();
            let after = text[at + word.len()..].chars().next();
            !before.is_some_and(char::is_alphanumeric) && !after.is_some_and(char::is_alphanumeric)
        })
        .count()
}

#[test]
fn a_split_endnote_is_painted_extracted_and_linked_back_once() {
    let (engine, request, layout) = laid_out(&long_endnotes(), 64);
    let display: Value = serde_json::from_str(
        &engine
            .build_display_list_json(&layout_json(&engine, &request))
            .unwrap(),
    )
    .unwrap();
    let painted = painted_notes(&display);
    let mut split = BTreeSet::new();
    for note in 1..=12 {
        let text: String = words(note, 70).split_whitespace().collect();
        let parts = &painted[&format!("endnote-{note}")];
        let shown: String = parts.iter().map(|(_, text)| text.as_str()).collect();
        assert!(shown.ends_with(&text), "endnote {note} paints {shown}");
        assert_eq!(shown.matches(&text).count(), 1);
        assert!(
            shown.len() < text.len() + 8,
            "endnote {note} paints {shown}"
        );
        if parts.len() > 1 {
            split.insert(note);
        }
    }
    assert!(!split.is_empty());
    for note in &split {
        let backlinks: usize = display["pages"]
            .as_array()
            .unwrap()
            .iter()
            .flat_map(|page| page["noteAreas"].as_array().into_iter().flatten())
            .flat_map(|area| area["notes"].as_array().into_iter().flatten())
            .filter(|entry| entry["id"] == *note)
            .count();
        assert_eq!(backlinks, 1, "endnote {note} links back once");
    }
    assert!(pages(&layout).len() > 2);

    let paged = engine
        .export_structured_with_pages(&PageExportOptions {
            stories: Some(vec![StorySelection::Body, StorySelection::Endnotes]),
            ..PageExportOptions::new(RevisionView::Markup)
        })
        .unwrap()
        .content;
    let markdown = render_docx_markdown_with_pages(
        &paged.structured,
        PageAnnotations::from(&paged.layout),
        &PageMarkdownOptions {
            page_markers: Some(true),
            ..PageMarkdownOptions::default()
        },
    )
    .unwrap()
    .markdown;
    for note in &split {
        for word in words(*note, 70).split_whitespace() {
            assert_eq!(
                word_count(&markdown, word),
                1,
                "{word} in the extracted text"
            );
        }
        let story = format!("en:{note}");
        let occurrences: Vec<_> = paged
            .layout
            .occurrences
            .iter()
            .filter(|occurrence| occurrence.story == story)
            .collect();
        assert_eq!(occurrences.len(), 2, "endnote {note} shows on two pages");
        let fragments: Vec<_> = paged
            .layout
            .fragments
            .iter()
            .filter(|fragment| occurrences.iter().any(|o| o.id == fragment.occurrence_id))
            .filter(|fragment| {
                matches!(fragment.slice, docx_edit::structured::FragmentSlice::Block)
            })
            .collect();
        assert_eq!(fragments.len(), 2);
        assert!(fragments[0].continued_on_next && !fragments[0].continued_from_previous);
        assert!(fragments[1].continued_from_previous && !fragments[1].continued_on_next);
    }
}

#[test]
fn section_end_endnotes_flow_before_the_next_section() {
    let body = format!(
        "{}{}{}{}<w:sectPr>{NARROW_PAGE}</w:sectPr>",
        fixture::p(
            "30000001",
            &format!("{}{}", fixture::r("First section"), reference(1))
        ),
        fixture::p(
            "30000002",
            &format!(
                r#"<w:pPr><w:sectPr><w:type w:val="nextPage"/>{NARROW_PAGE}</w:sectPr></w:pPr>{}{}"#,
                fixture::r("First section ends"),
                reference(2)
            )
        ),
        fixture::p(
            "30000003",
            &format!("{}{}", fixture::r("Second section"), reference(3))
        ),
        fixture::p("30000004", &fixture::r("Second section ends")),
    );
    let notes = [words(1, 8), words(2, 8), words(3, 8)];
    let bytes = endnotes_docx(
        &body,
        &notes,
        Some(r#"<w:endnotePr><w:pos w:val="sectEnd"/></w:endnotePr>"#),
    );
    let (_, _, layout) = laid_out(&bytes, 65);
    let pages = pages(&layout);
    let section_of = |page: &Value| page["sectionIndex"].as_u64().unwrap();
    let second = pages.iter().position(|page| section_of(page) == 1).unwrap();
    let mut placed = BTreeMap::new();
    for (index, page) in pages.iter().enumerate() {
        if let Some(area) = endnote_area(page) {
            assert_eq!(area["placement"], "sectEnd");
            for id in note_ids(area) {
                placed.insert(id, index);
            }
        }
    }
    assert!(placed[&1] < second && placed[&2] < second, "{placed:?}");
    assert_eq!(section_of(&pages[placed[&1]]), 0);
    assert_eq!(placed[&3], pages.len() - 1);
    assert_eq!(section_of(&pages[placed[&3]]), 1);
    for note in [2, 3] {
        let page = &pages[placed[&note]];
        let y = endnote_area(page).unwrap()["y"].as_f64().unwrap();
        let body = fragment_bottom(page).unwrap();
        assert!(
            y + 0.5 >= body && y < body + LINE_SLACK,
            "endnote {note} at {y} follows its section's text, which ends at {body}"
        );
    }
}

#[test]
fn an_edit_before_flowed_endnotes_lays_out_as_a_fresh_session_does() {
    let bytes = long_endnotes();
    let (engine, request) = fixture::laid_out(&bytes, 66);
    for text in ["Inserted words ", "and more "] {
        engine
            .doc()
            .insert_text(
                &EditCtx::local("", ""),
                Position::new("body", 0),
                text,
                FormatPolicy::Inherit,
            )
            .unwrap();
        let edited = engine
            .layout_document_with_regions_retained_json(&request)
            .unwrap();
        let fresh = docx_edit::EngineSession::new(67);
        fresh
            .doc()
            .apply_update_v1(&engine.doc().encode_state_as_update_v1())
            .unwrap();
        fresh
            .doc()
            .set_note_separator_state(engine.doc().note_separator_state().unwrap());
        assert_eq!(
            edited,
            fresh
                .layout_document_with_regions_retained_json(&request)
                .unwrap()
        );
    }
}

fn endnote_areas(page: &Value) -> Vec<&Value> {
    page["noteAreas"]
        .as_array()
        .into_iter()
        .flatten()
        .filter(|area| area["kind"] == "endnote")
        .collect()
}

/// The body fragments of `page` an endnote area paints over: both overlap down the page and,
/// when the area names its column, across it. A paragraph fragment states the text area's
/// width; its lines fill no more than one column.
fn overlapped_body(page: &Value, area: &Value) -> Vec<String> {
    let top = area["y"].as_f64().unwrap();
    let bottom = area_bottom(area);
    let span = area["x"].as_f64().zip(area["width"].as_f64());
    page["fragments"]
        .as_array()
        .into_iter()
        .flatten()
        .filter(|fragment| fragment["kind"] != "reserved")
        .filter(|fragment| {
            let y = fragment["y"].as_f64().unwrap();
            let height = fragment["height"].as_f64().unwrap();
            let x = fragment["x"].as_f64().unwrap();
            let width = fragment["width"]
                .as_f64()
                .unwrap()
                .min(span.map_or(f64::INFINITY, |(_, across)| across));
            y + height > top + 0.5
                && y + 0.5 < bottom
                && span
                    .is_none_or(|(left, across)| x + width > left + 0.5 && x + 0.5 < left + across)
        })
        .map(|fragment| fragment["blockId"].to_string())
        .collect()
}

#[test]
fn section_end_endnotes_after_a_continuous_break_get_an_area_each() {
    let body = format!(
        "{}{}{}<w:sectPr><w:type w:val=\"continuous\"/>{NARROW_PAGE}</w:sectPr>",
        fixture::p(
            "31000001",
            &format!("{}{}", fixture::r("First section"), reference(1))
        ),
        fixture::p(
            "31000002",
            &format!(
                r#"<w:pPr><w:sectPr>{NARROW_PAGE}</w:sectPr></w:pPr>{}"#,
                fixture::r("First section ends")
            )
        ),
        fixture::p(
            "31000003",
            &format!("{}{}", fixture::r("Second section"), reference(2))
        ),
    );
    let bytes = endnotes_docx(
        &body,
        &[words(1, 6), words(2, 6)],
        Some(r#"<w:endnotePr><w:pos w:val="sectEnd"/></w:endnotePr>"#),
    );
    let (_, _, layout) = laid_out(&bytes, 68);
    let page = &pages(&layout)[0];
    let areas = endnote_areas(page);
    assert_eq!(areas.len(), 2, "one area for each section's endnotes");
    assert_eq!(note_ids(areas[0]), [1]);
    assert_eq!(note_ids(areas[1]), [2]);
    for area in &areas {
        assert!(area["separator"]["height"].as_f64().unwrap() > 0.0);
        assert!(area.get("continuation").is_none());
        let over = overlapped_body(page, area);
        assert!(
            over.is_empty(),
            "endnotes {:?} paint over {over:?}",
            note_ids(area)
        );
    }
    assert!(areas[1]["y"].as_f64().unwrap() >= area_bottom(areas[0]));
}

#[test]
fn endnotes_in_two_columns_paint_in_the_column_they_flow_into() {
    let body = format!(
        r#"{}<w:sectPr><w:cols w:num="2" w:space="720"/>{NARROW_PAGE}</w:sectPr>"#,
        (1..=4)
            .map(|id| fixture::p(
                &format!("{:08X}", 0x3200_0000 + id),
                &format!("{}{}", fixture::r(&format!("Body {id}")), reference(id))
            ))
            .collect::<String>()
    );
    let notes: Vec<String> = (1..=4).map(|note| words(note, 50)).collect();
    let (engine, request, layout) = laid_out(&endnotes_docx(&body, &notes, None), 69);
    let page = &pages(&layout)[0];
    let areas = endnote_areas(page);
    assert!(areas.len() >= 2, "{} areas", areas.len());
    let lefts: Vec<f64> = areas
        .iter()
        .map(|area| area["x"].as_f64().unwrap())
        .collect();
    assert!(
        lefts[1] > lefts[0] + areas[0]["width"].as_f64().unwrap() - 0.5,
        "{lefts:?}"
    );
    assert_eq!(areas[1]["continuation"], true);
    for area in &areas {
        let over = overlapped_body(page, area);
        assert!(
            over.is_empty(),
            "endnotes {:?} paint over {over:?}",
            note_ids(area)
        );
    }
    let display: Value = serde_json::from_str(
        &engine
            .build_display_list_json(&layout_json(&engine, &request))
            .unwrap(),
    )
    .unwrap();
    let painted = display["pages"][0]["noteAreas"]
        .as_array()
        .unwrap()
        .iter()
        .filter(|area| area["kind"] == "endnote")
        .collect::<Vec<_>>();
    for (area, left) in painted.iter().zip(&lefts) {
        let width = area["width"].as_f64().unwrap();
        for primitive in area["primitives"].as_array().unwrap() {
            if let Some(x) = primitive["x"].as_f64() {
                assert!(
                    x + 0.5 >= *left && x < left + width,
                    "a primitive at {x} outside its column at {left}"
                );
            }
        }
    }
}

#[test]
fn a_kept_together_endnote_taller_than_a_page_continues_on_the_next() {
    let text = words(1, 400);
    let note = format!("<w:pPr><w:keepLines/></w:pPr>{}", fixture::r(&text));
    let body = format!(
        "{}<w:sectPr>{NARROW_PAGE}</w:sectPr>",
        fixture::p(
            "33000001",
            &format!("{}{}", fixture::r("Body"), reference(1))
        )
    );
    let (engine, request, layout) = laid_out(&endnotes_docx(&body, &[note], None), 70);
    let display: Value = serde_json::from_str(
        &engine
            .build_display_list_json(&layout_json(&engine, &request))
            .unwrap(),
    )
    .unwrap();
    let parts = &painted_notes(&display)["endnote-1"];
    assert!(parts.len() > 1, "the note paints on {} page", parts.len());
    let shown: String = parts.iter().map(|(_, text)| text.as_str()).collect();
    let squeezed: String = text.split_whitespace().collect();
    assert!(shown.ends_with(&squeezed));
    assert_eq!(shown.matches(&squeezed).count(), 1);
    for (index, page) in pages(&layout).iter().enumerate() {
        for area in endnote_areas(page) {
            assert!(
                area_bottom(area) <= content_bottom(page) + 0.5,
                "page {index}: endnotes end at {} below the text area",
                area_bottom(area)
            );
        }
    }
    let paged = engine
        .export_structured_with_pages(&PageExportOptions {
            stories: Some(vec![StorySelection::Body, StorySelection::Endnotes]),
            ..PageExportOptions::new(RevisionView::Markup)
        })
        .unwrap()
        .content;
    let markdown = render_docx_markdown_with_pages(
        &paged.structured,
        PageAnnotations::from(&paged.layout),
        &PageMarkdownOptions {
            page_markers: Some(true),
            ..PageMarkdownOptions::default()
        },
    )
    .unwrap()
    .markdown;
    for word in text.split_whitespace() {
        assert_eq!(
            word_count(&markdown, word),
            1,
            "{word} in the extracted text"
        );
    }
}
