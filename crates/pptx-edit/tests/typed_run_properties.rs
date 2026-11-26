use pptx_edit::{DeckSession, EditCtx, TextCaps, TextStyle, TextStylePatch};
use serde_json::json;
use yrs::types::text::YChange;
use yrs::{Any, Map, Out, ReadTxn, Text, Transact};

const NS: &str = r#"xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships""#;
const REL: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const EXPLICIT_DEFAULTS: &str = r#"lang="en-US" sz="1800" b="0" i="0" u="none" strike="noStrike" cap="none" baseline="0" kern="0" spc="0" dirty="0" noProof="0" normalizeH="0" kumimoji="0" err="0""#;
const SOURCE_FONT_SLOTS: &str = r#"<a:latin panose="020B0604020202020204" typeface="Arial"/><a:ea charset="-122" panose="02010600030101010101" pitchFamily="2" typeface="SimSun"/>"#;
const PICKED_FONT: &str = "Microsoft YaHei";

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

fn format_east_asian_run(patch: &TextStylePatch) -> String {
    let source = slide(&format!(
        "<a:p><a:r><a:rPr>{SOURCE_FONT_SLOTS}</a:rPr><a:t>中文字体</a:t></a:r></a:p>"
    ));
    let session = DeckSession::open(&deck(&source), 7101).unwrap();
    let snapshot = session.snapshot().unwrap();
    let story = &snapshot.slides[0].shapes[0].text_stories[0];
    let receipt = session
        .format_text(
            &EditCtx::local("test"),
            &story.id,
            0,
            story.length - 1,
            patch,
        )
        .unwrap();
    assert_eq!(receipt.text, "中文字体");
    let saved = slide_xml(&session.save().unwrap());
    let saved_runs = runs(&saved);
    assert_eq!(saved_runs.len(), 1, "{saved}");
    let (properties, text) = saved_runs[0].split_once("<a:t>").unwrap();
    assert_eq!(text, "中文字体</a:t>");
    properties.to_owned()
}

#[test]
fn explicit_font_updates_latin_and_east_asian_typefaces() {
    let properties = format_east_asian_run(&TextStylePatch {
        font_family: Some(PICKED_FONT.into()),
        ..Default::default()
    });
    assert!(
        properties.contains(r#"<a:latin typeface="Microsoft YaHei"/>"#),
        "{properties}"
    );
    assert!(
        properties.contains(r#"<a:ea typeface="Microsoft YaHei"/>"#),
        "{properties}"
    );
}

#[test]
fn bold_only_preserves_the_east_asian_typeface() {
    let properties = format_east_asian_run(&TextStylePatch {
        bold: Some(true),
        ..Default::default()
    });
    assert!(properties.contains(r#" b="1""#), "{properties}");
    assert!(properties.contains(SOURCE_FONT_SLOTS), "{properties}");
}

#[test]
fn a_mixed_selection_updates_east_asian_typefaces_only_in_east_asian_runs() {
    let source = slide(&format!(
        "<a:p><a:r><a:rPr>{SOURCE_FONT_SLOTS}</a:rPr><a:t>abc</a:t></a:r><a:r><a:rPr>{SOURCE_FONT_SLOTS}</a:rPr><a:t>中文</a:t></a:r></a:p>"
    ));
    let session = DeckSession::open(&deck(&source), 7101).unwrap();
    let snapshot = session.snapshot().unwrap();
    let story = &snapshot.slides[0].shapes[0].text_stories[0];
    session
        .format_text(
            &EditCtx::local("test"),
            &story.id,
            0,
            story.length - 1,
            &TextStylePatch {
                font_family: Some(PICKED_FONT.into()),
                ..Default::default()
            },
        )
        .unwrap();
    let saved = slide_xml(&session.save().unwrap());
    let saved_runs = runs(&saved);
    assert_eq!(saved_runs.len(), 2, "{saved}");
    assert_eq!(
        saved_runs[0],
        r#"<a:rPr><a:latin typeface="Microsoft YaHei"/><a:ea charset="-122" panose="02010600030101010101" pitchFamily="2" typeface="SimSun"/></a:rPr><a:t>abc</a:t>"#
    );
    assert_eq!(
        saved_runs[1],
        r#"<a:rPr><a:latin typeface="Microsoft YaHei"/><a:ea typeface="Microsoft YaHei"/></a:rPr><a:t>中文</a:t>"#
    );
}

#[test]
fn an_east_asian_font_pick_survives_reopen_and_an_untouched_save() {
    let untouched = r#"<p:sp><p:nvSpPr><p:cNvPr name="Untouched" id="3"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:p><a:r><a:rPr><a:latin typeface="Arial"/><a:ea typeface="SimSun"/></a:rPr><a:t>未改</a:t></a:r></a:p></p:txBody></p:sp>"#;
    let source = slide(&format!(
        "<a:p><a:r><a:rPr>{SOURCE_FONT_SLOTS}</a:rPr><a:t>中文</a:t></a:r></a:p>"
    ))
    .replace("</p:spTree>", &format!("{untouched}</p:spTree>"));
    let session = DeckSession::open(&deck(&source), 7101).unwrap();
    let snapshot = session.snapshot().unwrap();
    let story = &snapshot.slides[0].shapes[0].text_stories[0];
    session
        .format_text(
            &EditCtx::local("test"),
            &story.id,
            0,
            story.length - 1,
            &TextStylePatch {
                font_family: Some(PICKED_FONT.into()),
                ..Default::default()
            },
        )
        .unwrap();
    let saved = session.save().unwrap();
    let saved_xml = slide_xml(&saved);
    assert!(saved_xml.contains(untouched), "{saved_xml}");
    let reopened = DeckSession::open(&saved, 7102).unwrap();
    let reopened_xml = slide_xml(&reopened.save().unwrap());
    assert_eq!(
        runs(&reopened_xml),
        [r#"<a:rPr><a:latin typeface="Microsoft YaHei"/><a:ea typeface="Microsoft YaHei"/></a:rPr><a:t>中文</a:t>"#]
    );
    assert_eq!(reopened_xml, saved_xml);
}

#[test]
fn replayed_font_picks_keep_the_existing_story_and_package_attributes() {
    let source = deck(&slide(&format!(
        "<a:p><a:r><a:rPr>{SOURCE_FONT_SLOTS}</a:rPr><a:t>中文</a:t></a:r></a:p>"
    )));
    let session = DeckSession::open_replay_baseline(&source, Some(7101), None).unwrap();
    session.peer_identity(Vec::new()).unwrap();
    let snapshot = session.snapshot().unwrap();
    let story_id = &snapshot.slides[0].shapes[0].text_stories[0].id;
    let attributes = || {
        let txn = session.yrs_doc().transact();
        let stories = txn.get_map("pptx:stories").unwrap();
        let Some(Out::YText(story)) = stories.get(&txn, story_id) else {
            panic!("missing story");
        };
        story
            .diff(&txn, YChange::identity)
            .into_iter()
            .find(|run| matches!(run.insert, Out::Any(Any::String(_))))
            .and_then(|run| run.attributes)
            .unwrap()
    };
    let original = attributes();
    assert_eq!(original.len(), 1);
    assert_eq!(original.get("fontFamily"), Some(&Any::from("Arial")));
    session
        .replay_json(
            &json!({
                "sequence": 1,
                "baseVersion": session.version(),
                "op": {"method": "formatText", "args": [story_id, 0, 2, {"fontFamily": PICKED_FONT}]}
            })
            .to_string(),
        )
        .unwrap();
    let live = attributes();
    assert_eq!(live.len(), 1);
    assert_eq!(live.get("fontFamily"), Some(&Any::from(PICKED_FONT)));
    assert!(session
        .replay_json(
            &json!({
                "sequence": 2,
                "baseVersion": session.version(),
                "op": {"method": "formatText", "args": [story_id, 0, 2, {"eastAsianFontFamily": "SimSun"}]}
            })
            .to_string(),
        )
        .is_err());
    let txn = session.yrs_doc().transact();
    let meta = txn.get_map("pptx:meta").unwrap();
    let Some(Out::Any(Any::Buffer(package))) = meta.get(&txn, "packageJson") else {
        panic!("missing package JSON");
    };
    let package = std::str::from_utf8(&package).unwrap();
    assert!(!package.contains("SimSun"));
    assert!(!package.contains("eastAsianFontFamily"));
    drop(txn);
    assert!(!serde_json::to_string(&session.snapshot().unwrap())
        .unwrap()
        .contains("eastAsianFontFamily"));
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
