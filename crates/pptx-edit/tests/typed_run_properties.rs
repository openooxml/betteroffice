use pptx_edit::{DeckSession, EditCtx, TextCaps, TextStyle, TextStylePatch};

const NS: &str = r#"xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships""#;
const REL: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const EXPLICIT_DEFAULTS: &str = r#"lang="en-US" sz="1800" b="0" i="0" u="none" strike="noStrike" cap="none" baseline="0" kern="0" spc="0" dirty="0" noProof="0" normalizeH="0" kumimoji="0" err="0""#;

fn slide(paragraphs: &str) -> String {
    format!(
        r#"<p:sld {NS}><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/><p:sp><p:nvSpPr><p:cNvPr id="2" name="Text"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="4267200" cy="1371600"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/>{paragraphs}</p:txBody></p:sp></p:spTree></p:cSld></p:sld>"#
    )
}

fn deck(slide: &str) -> Vec<u8> {
    let parts = [
        ("[Content_Types].xml", r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/><Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/></Types>"#.to_owned()),
        ("_rels/.rels", format!(r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="doc" Type="{REL}/officeDocument" Target="ppt/presentation.xml"/></Relationships>"#)),
        ("ppt/presentation.xml", format!(r#"<p:presentation {NS}><p:sldIdLst><p:sldId id="256" r:id="slide1"/></p:sldIdLst><p:sldSz cx="6096000" cy="3429000"/></p:presentation>"#)),
        ("ppt/_rels/presentation.xml.rels", format!(r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="slide1" Type="{REL}/slide" Target="slides/slide1.xml"/></Relationships>"#)),
        ("ppt/slides/slide1.xml", slide.to_owned()),
    ]
    .into_iter()
    .map(|(path, xml)| (path.to_owned(), xml.into_bytes()))
    .collect::<Vec<_>>();
    ooxml_opc::rezip_parts(&parts).unwrap()
}

fn slide_xml(bytes: &[u8]) -> String {
    String::from_utf8(
        ooxml_opc::unzip_parts(bytes)
            .unwrap()
            .into_iter()
            .find(|(path, _)| path == "ppt/slides/slide1.xml")
            .unwrap()
            .1,
    )
    .unwrap()
}

/// Types `text` at `index` with `style`, or with the first run's own style as
/// the roundtrip check does, and returns the source and saved slide XML.
fn type_into(
    paragraphs: &str,
    index: u32,
    text: &str,
    style: Option<TextStyle>,
) -> (String, String) {
    let source = slide(paragraphs);
    let session = DeckSession::open(&deck(&source), 7101).unwrap();
    let snapshot = session.snapshot().unwrap();
    let shapes = &snapshot.slides[0].shapes;
    let story = shapes
        .iter()
        .flat_map(|shape| &shape.text_stories)
        .next()
        .unwrap();
    let style = style.unwrap_or_else(|| story.paragraphs[0].runs[0].style.clone());
    session
        .insert_text(&EditCtx::local("typing"), &story.id, index, text, &style)
        .unwrap();
    (source, slide_xml(&session.save().unwrap()))
}

fn one_run(properties: &str) -> String {
    let run = format!("<a:r><a:rPr {properties}/><a:t>Hello</a:t></a:r>");
    format!("<a:p>{run}</a:p>")
}

fn paragraph(xml: &str) -> &str {
    xml.split_once("<a:p>")
        .and_then(|(_, rest)| rest.split_once("</a:p>"))
        .map_or(xml, |(paragraph, _)| paragraph)
}

fn runs(xml: &str) -> Vec<&str> {
    paragraph(xml)
        .split("<a:r>")
        .skip(1)
        .map(|run| run.split("</a:r>").next().unwrap())
        .collect()
}

#[test]
fn text_typed_at_the_end_of_a_no_strike_run_stays_in_that_run() {
    let paragraphs = r#"<a:p><a:r><a:rPr lang="en-US" sz="1800" b="0" i="0" u="none" strike="noStrike" baseline="0" dirty="0"><a:latin typeface="Arial"/></a:rPr><a:t>Hello</a:t></a:r><a:endParaRPr lang="en-GB" dirty="0"/></a:p>"#;
    let (source, saved) = type_into(paragraphs, 5, "_X_", None);
    assert_eq!(source.matches("<a:t>Hello</a:t>").count(), 1);
    assert_eq!(
        saved,
        source.replace("<a:t>Hello</a:t>", "<a:t>Hello_X_</a:t>")
    );
}

#[test]
fn east_asian_typeface_persists_through_font_picks_and_typing() {
    let paragraphs = r#"<a:p><a:r><a:rPr><a:latin typeface="Arial"/><a:ea typeface="SimSun"/></a:rPr><a:t>中文</a:t></a:r></a:p>"#;
    let source = slide(paragraphs);
    let session = DeckSession::open(&deck(&source), 7101).unwrap();
    let snapshot = session.snapshot().unwrap();
    let story = snapshot.slides[0]
        .shapes
        .iter()
        .flat_map(|shape| &shape.text_stories)
        .next()
        .unwrap();
    let updated_ea = r#"<a:ea typeface="Microsoft YaHei"/>"#;

    session
        .format_text(
            &EditCtx::local("font-family"),
            &story.id,
            0,
            2,
            &TextStylePatch {
                font_family: Some("Microsoft YaHei".to_owned()),
                ..TextStylePatch::default()
            },
        )
        .unwrap();
    let saved = slide_xml(&session.save().unwrap());
    let formatted = runs(&saved);
    assert_eq!(formatted.len(), 1, "{}", paragraph(&saved));
    let (properties, text) = formatted[0].split_once("<a:t>").unwrap();
    assert_eq!(text, "中文</a:t>");
    assert!(
        properties.contains(r#"<a:latin typeface="Microsoft YaHei"/>"#),
        "{properties}"
    );
    assert!(properties.contains(updated_ea), "{properties}");

    let formatted = session.story(&story.id).unwrap();
    let style = &formatted.paragraphs[0].runs[0].style;
    session
        .insert_text(&EditCtx::local("typing"), &story.id, 2, "字", style)
        .unwrap();
    let saved = slide_xml(&session.save().unwrap());
    let typed = runs(&saved)
        .into_iter()
        .find(|run| run.ends_with("字</a:t>"))
        .unwrap();
    assert!(typed.contains(updated_ea), "{typed}");

    let snapshot = session.snapshot().unwrap();
    let story = snapshot.slides[0]
        .shapes
        .iter()
        .flat_map(|shape| &shape.text_stories)
        .find(|current| current.id == story.id)
        .unwrap();
    assert_eq!(story.plain_text(), "中文字");
    for run in &story.paragraphs[0].runs {
        let style = serde_json::to_value(&run.style).unwrap();
        let mut keys = style
            .as_object()
            .unwrap()
            .keys()
            .map(String::as_str)
            .collect::<Vec<_>>();
        keys.sort_unstable();
        assert_eq!(
            keys,
            [
                "bold",
                "color",
                "fontFamily",
                "fontSizePt",
                "italic",
                "underline"
            ]
        );
    }
}

#[test]
fn non_font_edits_and_untouched_runs_preserve_east_asian_typeface() {
    let paragraphs = concat!(
        r#"<a:p><a:r><a:rPr><a:latin typeface="Arial"/><a:ea typeface="SimSun"/></a:rPr><a:t>中文</a:t></a:r></a:p>"#,
        r#"<a:p><a:r><a:rPr><a:latin typeface="Arial"/><a:ea typeface="+mn-ea"/></a:rPr><a:t>保留</a:t></a:r></a:p>"#,
    );
    let session = DeckSession::open(&deck(&slide(paragraphs)), 7102).unwrap();
    let story_id = session.snapshot().unwrap().slides[0].shapes[0].text_stories[0]
        .id
        .clone();
    session
        .format_text(
            &EditCtx::local("bold"),
            &story_id,
            0,
            2,
            &TextStylePatch {
                bold: Some(true),
                font_size_pt: Some(24.0),
                ..TextStylePatch::default()
            },
        )
        .unwrap();
    let saved = slide_xml(&session.save().unwrap());
    assert!(
        saved.contains(r#"<a:rPr b="1" sz="2400"><a:latin typeface="Arial"/><a:ea typeface="SimSun"/></a:rPr><a:t>中文</a:t>"#),
        "{saved}"
    );
    assert!(
        saved.contains(
            r#"<a:rPr><a:latin typeface="Arial"/><a:ea typeface="+mn-ea"/></a:rPr><a:t>保留</a:t>"#
        ),
        "{saved}"
    );
}

#[test]
fn latin_only_selection_preserves_east_asian_typeface_while_mixed_updates_it() {
    let paragraphs = r#"<a:p><a:r><a:rPr><a:latin typeface="Arial"/><a:ea typeface="SimSun"/></a:rPr><a:t>Hello中文</a:t></a:r></a:p>"#;
    let session = DeckSession::open(&deck(&slide(paragraphs)), 7103).unwrap();
    let story_id = session.snapshot().unwrap().slides[0].shapes[0].text_stories[0]
        .id
        .clone();

    session
        .format_text(
            &EditCtx::local("latin-only"),
            &story_id,
            0,
            5,
            &TextStylePatch {
                font_family: Some("Georgia".to_owned()),
                ..TextStylePatch::default()
            },
        )
        .unwrap();
    let saved = slide_xml(&session.save().unwrap());
    let saved_runs = runs(&saved);
    assert_eq!(saved_runs.len(), 2, "{}", paragraph(&saved));
    assert!(
        saved_runs[0].contains(r#"<a:latin typeface="Georgia"/>"#)
            && saved_runs[0].contains(r#"<a:ea typeface="SimSun"/>"#),
        "{}",
        saved_runs[0]
    );
    assert!(
        saved_runs[1].contains(r#"<a:latin typeface="Arial"/>"#)
            && saved_runs[1].contains(r#"<a:ea typeface="SimSun"/>"#),
        "{}",
        saved_runs[1]
    );

    session
        .format_text(
            &EditCtx::local("ea-only"),
            &story_id,
            5,
            7,
            &TextStylePatch {
                font_family: Some("Microsoft YaHei".to_owned()),
                ..TextStylePatch::default()
            },
        )
        .unwrap();
    let saved = slide_xml(&session.save().unwrap());
    let saved_runs = runs(&saved);
    assert_eq!(saved_runs.len(), 2, "{}", paragraph(&saved));
    assert!(
        saved_runs[0].contains(r#"<a:ea typeface="SimSun"/>"#),
        "{}",
        saved_runs[0]
    );
    assert!(
        saved_runs[1].contains(r#"<a:latin typeface="Microsoft YaHei"/>"#)
            && saved_runs[1].contains(r#"<a:ea typeface="Microsoft YaHei"/>"#),
        "{}",
        saved_runs[1]
    );
}

#[test]
fn font_pick_inserts_ea_when_absent_and_replaces_theme_ea_with_metrics() {
    let paragraphs = concat!(
        r#"<a:p><a:r><a:rPr><a:latin typeface="Arial"/><a:cs typeface="Arial"/></a:rPr><a:t>无东亚</a:t></a:r></a:p>"#,
        r#"<a:p><a:r><a:rPr><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea" panose="02010600030101010101" pitchFamily="34" charset="-122"/></a:rPr><a:t>主题字</a:t></a:r></a:p>"#,
    );
    let session = DeckSession::open(&deck(&slide(paragraphs)), 7104).unwrap();
    let story_id = session.snapshot().unwrap().slides[0].shapes[0].text_stories[0]
        .id
        .clone();
    session
        .format_text(
            &EditCtx::local("font-family"),
            &story_id,
            0,
            7,
            &TextStylePatch {
                font_family: Some("Microsoft YaHei".to_owned()),
                ..TextStylePatch::default()
            },
        )
        .unwrap();
    let saved = slide_xml(&session.save().unwrap());
    let saved_runs: Vec<&str> = saved
        .split("<a:r>")
        .skip(1)
        .map(|part| part.split("</a:r>").next().unwrap())
        .collect();
    assert_eq!(saved_runs.len(), 2, "{saved}");

    let first = &saved_runs[0];
    let latin_pos = first
        .find(r#"<a:latin typeface="Microsoft YaHei"/>"#)
        .unwrap();
    let ea_pos = first.find(r#"<a:ea typeface="Microsoft YaHei"/>"#).unwrap();
    let cs_pos = first.find(r#"<a:cs typeface="Arial"/>"#).unwrap();
    assert!(latin_pos < ea_pos && ea_pos < cs_pos, "{first}");

    let second = &saved_runs[1];
    assert!(
        second.contains(r#"<a:ea typeface="Microsoft YaHei"/>"#),
        "{second}"
    );
    assert!(!second.contains("panose="), "{second}");
    assert!(!second.contains("pitchFamily="), "{second}");
    assert!(!second.contains("charset="), "{second}");
}

#[test]
fn picking_same_latin_face_for_east_asian_text_still_updates_ea_on_save() {
    let paragraphs = r#"<a:p><a:r><a:rPr><a:latin typeface="Arial"/><a:ea typeface="SimSun"/></a:rPr><a:t>中文</a:t></a:r></a:p>"#;
    let session = DeckSession::open(&deck(&slide(paragraphs)), 7105).unwrap();
    let story_id = session.snapshot().unwrap().slides[0].shapes[0].text_stories[0]
        .id
        .clone();
    session
        .format_text(
            &EditCtx::local("same-latin"),
            &story_id,
            0,
            2,
            &TextStylePatch {
                font_family: Some("Arial".to_owned()),
                ..TextStylePatch::default()
            },
        )
        .unwrap();
    let saved = slide_xml(&session.save().unwrap());
    let saved_runs = runs(&saved);
    assert_eq!(saved_runs.len(), 1, "{}", paragraph(&saved));
    assert!(
        saved_runs[0].contains(r#"<a:latin typeface="Arial"/>"#)
            && saved_runs[0].contains(r#"<a:ea typeface="Arial"/>"#),
        "{}",
        saved_runs[0]
    );
}

#[test]
fn typing_in_unedited_east_asian_run_keeps_authored_ea_while_new_font_pick_sets_it() {
    let paragraphs = r#"<a:p><a:r><a:rPr><a:latin typeface="Arial"/><a:ea typeface="SimSun"/></a:rPr><a:t>中文</a:t></a:r></a:p>"#;
    let session = DeckSession::open(&deck(&slide(paragraphs)), 7106).unwrap();
    let story_id = session.snapshot().unwrap().slides[0].shapes[0].text_stories[0]
        .id
        .clone();
    let inherited = session.story(&story_id).unwrap().paragraphs[0].runs[0]
        .style
        .clone();
    session
        .insert_text(&EditCtx::local("typing"), &story_id, 2, "字", &inherited)
        .unwrap();
    let saved = slide_xml(&session.save().unwrap());
    let saved_runs = runs(&saved);
    assert_eq!(saved_runs.len(), 1, "{}", paragraph(&saved));
    assert!(
        saved_runs[0].contains(r#"<a:ea typeface="SimSun"/>"#),
        "{}",
        saved_runs[0]
    );

    let new_font = TextStyle {
        font_family: Some("Microsoft YaHei".to_owned()),
        ..inherited
    };
    session
        .insert_text(&EditCtx::local("new-font"), &story_id, 3, "体", &new_font)
        .unwrap();
    let saved = slide_xml(&session.save().unwrap());
    let typed = runs(&saved)
        .into_iter()
        .find(|run| run.ends_with("体</a:t>"))
        .unwrap();
    assert!(
        typed.contains(r#"<a:latin typeface="Microsoft YaHei"/>"#)
            && typed.contains(r#"<a:ea typeface="Microsoft YaHei"/>"#),
        "{typed}"
    );
}

#[test]
fn east_asian_font_pick_syncs_across_collaborative_peers_and_reopens() {
    let paragraphs = r#"<a:p><a:r><a:rPr><a:latin typeface="Arial"/><a:ea typeface="SimSun"/></a:rPr><a:t>中文</a:t></a:r></a:p>"#;
    let source_bytes = deck(&slide(paragraphs));
    let peer_a = DeckSession::open(&source_bytes, 7107).unwrap();
    let peer_b = DeckSession::open(&source_bytes, 7108).unwrap();
    let story_id = peer_a.snapshot().unwrap().slides[0].shapes[0].text_stories[0]
        .id
        .clone();

    peer_a
        .format_text(
            &EditCtx::local("peer-a"),
            &story_id,
            0,
            2,
            &TextStylePatch {
                font_family: Some("Microsoft YaHei".to_owned()),
                ..TextStylePatch::default()
            },
        )
        .unwrap();
    let update = peer_a.encode_state_as_update_v1();
    peer_b.apply_update_v1(&update).unwrap();

    let saved_bytes = peer_b.save().unwrap();
    let saved = slide_xml(&saved_bytes);
    assert!(
        saved.contains(r#"<a:ea typeface="Microsoft YaHei"/>"#),
        "{saved}"
    );

    let reopened = DeckSession::open(&saved_bytes, 7109).unwrap();
    reopened
        .format_text(
            &EditCtx::local("bold-after-reopen"),
            &story_id,
            0,
            2,
            &TextStylePatch {
                bold: Some(true),
                ..TextStylePatch::default()
            },
        )
        .unwrap();
    let resaved = slide_xml(&reopened.save().unwrap());
    assert!(
        resaved.contains(r#"<a:ea typeface="Microsoft YaHei"/>"#),
        "{resaved}"
    );
}

#[test]
fn batch_format_text_updates_east_asian_typeface_even_when_latin_matches() {
    let paragraphs = r#"<a:p><a:r><a:rPr><a:latin typeface="Arial"/><a:ea typeface="SimSun"/></a:rPr><a:t>中文</a:t></a:r></a:p>"#;
    let session = DeckSession::open(&deck(&slide(paragraphs)), 7110).unwrap();
    let snapshot = session.snapshot().unwrap();
    let slide_id = snapshot.slides[0].id.clone();
    let shape_id = snapshot.slides[0].shapes[0].id.clone();
    let story_id = snapshot.slides[0].shapes[0].text_stories[0].id.clone();
    let request: pptx_edit::EditRequest = serde_json::from_value(serde_json::json!({
        "expectVersion": session.version(),
        "steps": [
            {
                "op": "formatText",
                "target": {
                    "kind": "range",
                    "slideId": slide_id,
                    "shapeId": shape_id,
                    "storyId": story_id,
                    "start": 0,
                    "end": 2
                },
                "patch": {
                    "fontFamily": "Arial"
                }
            }
        ]
    }))
    .unwrap();
    let applied = session.apply_edits(&request).unwrap().unwrap();
    assert!(applied.applied);
    let saved = slide_xml(&session.save().unwrap());
    assert!(saved.contains(r#"<a:ea typeface="Arial"/>"#), "{saved}");
}

#[test]
fn text_typed_at_the_end_of_a_run_keeps_each_explicit_default() {
    let rows = [
        ("strike", r#"strike="noStrike""#),
        ("u", r#"u="none""#),
        ("b", r#"b="0""#),
        ("i", r#"i="0""#),
        ("cap", r#"cap="none""#),
        ("baseline", r#"baseline="0""#),
        ("kern", r#"kern="0""#),
        ("spc", r#"spc="0""#),
        ("lang", r#"lang="en-US""#),
        ("dirty", r#"dirty="0""#),
        ("noProof", r#"noProof="0""#),
        ("normalizeH", r#"normalizeH="0""#),
        ("kumimoji", r#"kumimoji="0""#),
        ("err", r#"err="0""#),
    ];
    let mut mismatches = Vec::new();
    for (name, attribute) in rows {
        let (source, saved) = type_into(&one_run(attribute), 5, "_X_", None);
        if saved != source.replace("<a:t>Hello</a:t>", "<a:t>Hello_X_</a:t>") {
            mismatches.push(format!("{name}: {}", paragraph(&saved)));
        }
    }
    assert!(mismatches.is_empty(), "{}", mismatches.join("\n"));
}

#[test]
fn text_typed_in_the_middle_of_a_run_with_explicit_defaults_stays_in_that_run() {
    let (source, saved) = type_into(&one_run(EXPLICIT_DEFAULTS), 2, "_X_", None);
    assert_eq!(
        saved,
        source.replace("<a:t>Hello</a:t>", "<a:t>He_X_llo</a:t>")
    );
}

#[test]
fn text_typed_at_the_start_of_a_run_with_explicit_defaults_stays_in_that_run() {
    let (source, saved) = type_into(&one_run(EXPLICIT_DEFAULTS), 0, "_X_", None);
    assert_eq!(
        saved,
        source.replace("<a:t>Hello</a:t>", "<a:t>_X_Hello</a:t>")
    );
}

#[test]
fn text_typed_before_a_differently_formatted_run_joins_the_run_before_the_caret() {
    let paragraphs = format!(
        r#"<a:p><a:r><a:rPr {EXPLICIT_DEFAULTS}/><a:t>Hello</a:t></a:r><a:r><a:rPr lang="en-US" b="1" dirty="0"/><a:t> World</a:t></a:r></a:p>"#
    );
    let (source, saved) = type_into(&paragraphs, 5, "_X_", None);
    assert_eq!(
        saved,
        source.replace("<a:t>Hello</a:t>", "<a:t>Hello_X_</a:t>")
    );
}

#[test]
fn text_typed_into_an_empty_paragraph_keeps_its_explicit_defaults_but_no_strike_value() {
    let end = format!("<a:endParaRPr {EXPLICIT_DEFAULTS}/>");
    let paragraphs = format!("<a:p>{end}</a:p>");
    let style = TextStyle {
        bold: Some(false),
        italic: Some(false),
        font_size_pt: Some(18.0),
        underline: Some("none".to_owned()),
        strike: Some("noStrike".to_owned()),
        spacing_pt: Some(0.0),
        baseline_pct: Some(0.0),
        kern_pt: Some(0.0),
        caps: Some(TextCaps::None),
        ..TextStyle::default()
    };
    let (_, saved) = type_into(&paragraphs, 0, "_X_", Some(style));
    let typed = runs(&saved);
    assert_eq!(typed.len(), 1, "{}", paragraph(&saved));
    let (properties, text) = typed[0].split_once("<a:t>").unwrap();
    assert_eq!(text, "_X_</a:t>");
    for attribute in [
        r#"sz="1800""#,
        r#"b="0""#,
        r#"i="0""#,
        r#"u="none""#,
        r#"cap="none""#,
        r#"baseline="0""#,
        r#"kern="0""#,
        r#"spc="0""#,
    ] {
        assert!(
            properties.contains(attribute),
            "{attribute} missing from {properties}"
        );
    }
    assert!(!properties.contains("strike="), "{properties}");
    assert!(paragraph(&saved).ends_with(&end), "{}", paragraph(&saved));
}

#[test]
fn text_typed_after_struck_text_stays_struck_unless_strike_is_turned_off() {
    let paragraphs = one_run(r#"lang="en-US" strike="sngStrike""#);
    let (source, saved) = type_into(&paragraphs, 5, "_X_", None);
    assert_eq!(
        saved,
        source.replace("<a:t>Hello</a:t>", "<a:t>Hello_X_</a:t>")
    );
    let off = TextStyle {
        strike: Some("noStrike".to_owned()),
        ..TextStyle::default()
    };
    let (_, saved) = type_into(&paragraphs, 5, "_X_", Some(off));
    let typed = runs(&saved);
    assert_eq!(typed.len(), 2, "{}", paragraph(&saved));
    assert_eq!(
        typed[0],
        r#"<a:rPr lang="en-US" strike="sngStrike"/><a:t>Hello</a:t>"#
    );
    assert!(typed[1].ends_with("<a:t>_X_</a:t>"), "{}", typed[1]);
    assert!(!typed[1].contains("strike="), "{}", typed[1]);
}

#[test]
fn text_typed_without_a_strike_of_its_own_stays_in_a_no_strike_run() {
    let paragraphs = one_run(r#"lang="en-US" u="none" strike="noStrike""#);
    let style = TextStyle {
        underline: Some("none".to_owned()),
        ..TextStyle::default()
    };
    let (source, saved) = type_into(&paragraphs, 5, "_X_", Some(style));
    assert_eq!(
        saved,
        source.replace("<a:t>Hello</a:t>", "<a:t>Hello_X_</a:t>")
    );
}
