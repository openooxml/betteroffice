use pptx_edit::{DeckSession, EditCtx, TextParagraphDraft, TextRunDraft, TextStyle};

const SOURCE: &[u8] = include_bytes!("../../../apps/demo/public/betteroffice-demo.pptx");

fn formatted_source() -> (Vec<u8>, String) {
    let deck = DeckSession::open(SOURCE, 7).unwrap();
    let path = deck.snapshot().unwrap().slides[0]
        .source_part_path
        .clone()
        .unwrap();
    let mut parts = ooxml_opc::unzip_parts(SOURCE).unwrap();
    let (_, bytes) = parts.iter_mut().find(|(part, _)| part == &path).unwrap();
    let mut xml = String::from_utf8(bytes.clone()).unwrap();
    let start = xml.find("<p:txBody>").unwrap();
    let end = start + xml[start..].find("</p:txBody>").unwrap() + "</p:txBody>".len();
    xml.replace_range(start..end, r#"<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:pPr marL="3000" indent="-1000"/><a:r><a:rPr lang="en-US" strike="sngStrike"/><a:t>Source text</a:t></a:r><a:endParaRPr lang="de-DE" sz="1700"/></a:p><a:p><a:r><a:t>Second</a:t></a:r></a:p></p:txBody>"#);
    *bytes = xml.into_bytes();
    (ooxml_opc::rezip_parts(&parts).unwrap(), path)
}

fn slide_xml(bytes: &[u8], path: &str) -> String {
    String::from_utf8(
        ooxml_opc::unzip_parts(bytes)
            .unwrap()
            .into_iter()
            .find(|(part, _)| part == path)
            .unwrap()
            .1,
    )
    .unwrap()
}

fn formatted_story(deck: &DeckSession, slide: usize) -> pptx_edit::StorySnapshot {
    deck.snapshot().unwrap().slides[slide]
        .shapes
        .iter()
        .flat_map(|shape| &shape.text_stories)
        .find(|story| story.plain_text().starts_with("Source"))
        .unwrap()
        .clone()
}

fn assert_duplicate_markup(bytes: &[u8], source: &[u8], original_path: &str, text: &str) {
    let reopened = DeckSession::open(bytes, 9).unwrap();
    let snapshot = reopened.snapshot().unwrap();
    let copy_path = snapshot.slides[1].source_part_path.as_deref().unwrap();
    assert_ne!(copy_path, original_path);
    let xml = slide_xml(bytes, copy_path);
    let paragraph = xml
        .split("<a:p>")
        .find(|paragraph| paragraph.contains(text))
        .unwrap()
        .split("</a:p>")
        .next()
        .unwrap();
    for markup in [
        "marL=\"3000\"",
        "indent=\"-1000\"",
        "lang=\"en-US\"",
        "strike=\"sngStrike\"",
        "<a:endParaRPr lang=\"de-DE\" sz=\"1700\"/>",
    ] {
        assert!(paragraph.contains(markup), "missing {markup}: {paragraph}");
    }
    assert_eq!(
        slide_xml(bytes, original_path),
        slide_xml(source, original_path)
    );
}

#[test]
fn duplicate_preserves_source_paragraph_markup_across_edits_and_reopen() {
    let (source, path) = formatted_source();
    let deck = DeckSession::open(&source, 7).unwrap();
    let ctx = EditCtx::local("test");
    let original = deck.snapshot().unwrap().slides[0].clone();
    let story = formatted_story(&deck, 0);
    deck.duplicate_slide(&ctx, &original.id, 1).unwrap();
    let copy = formatted_story(&deck, 1);
    assert_ne!(copy.id, story.id);
    assert_eq!(deck.story(&story.id).unwrap(), story);
    let untouched = deck.save().unwrap();
    assert_duplicate_markup(&untouched, &source, &path, "Source text");
    deck.insert_text(&ctx, &copy.id, 6, "!", &TextStyle::default())
        .unwrap();
    let edited = deck.save().unwrap();
    assert_duplicate_markup(&edited, &source, &path, "Source! text");
    let reopened = DeckSession::open(&edited, 8).unwrap();
    let story = formatted_story(&reopened, 1);
    reopened
        .insert_text(&ctx, &story.id, 7, "!", &TextStyle::default())
        .unwrap();
    assert_duplicate_markup(&reopened.save().unwrap(), &source, &path, "Source!! text");
}

#[test]
fn duplicate_preserves_shapes_and_notes_and_can_outlive_its_source() {
    let deck = DeckSession::open(SOURCE, 7).unwrap();
    let ctx = EditCtx::local("test");
    let before = deck.snapshot().unwrap();
    let source = &before.slides[0];
    deck.set_slide_notes(&ctx, &source.id, "speaker notes")
        .unwrap();
    let clone = deck.duplicate_slide(&ctx, &source.id, 1).unwrap();
    let snapshot = deck.snapshot().unwrap();
    let copied = &snapshot.slides[1];
    assert_eq!(copied.notes, "speaker notes");
    assert_eq!(copied.shapes.len(), source.shapes.len());
    for (left, right) in source.shapes.iter().zip(&copied.shapes) {
        assert_ne!(left.id, right.id);
        assert_eq!(left.name, right.name);
        assert_eq!(left.source_id, right.source_id);
        assert_eq!(left.kind, right.kind);
    }
    deck.delete_slide(&ctx, &source.id).unwrap();
    let saved = deck.save().unwrap();
    let restored = DeckSession::open(&saved, 8).unwrap().snapshot().unwrap();
    assert_eq!(restored.slides.len(), before.slides.len());
    assert_eq!(restored.slides[0].notes, "speaker notes");
    assert_eq!(restored.slides[0].shapes.len(), copied.shapes.len());
    assert_eq!(deck.snapshot().unwrap().slides[0].id, clone.slide_id);
}

#[test]
fn duplicate_is_undoable_and_refuses_invalid_targets_without_updates() {
    let deck = DeckSession::open(SOURCE, 7).unwrap();
    let ctx = EditCtx::local("test");
    let baseline = deck.snapshot().unwrap();
    let source = &baseline.slides[0];
    let version = deck.version();
    assert!(deck.duplicate_slide(&ctx, "missing", 0).is_err());
    assert!(deck.duplicate_slide(&ctx, &source.id, 99999).is_err());
    assert_eq!(deck.version(), version);
    deck.duplicate_slide(&ctx, &source.id, 1).unwrap();
    assert!(deck.undo());
    assert_eq!(deck.snapshot().unwrap(), baseline);
}

fn notes_source(with_notes: bool) -> Vec<u8> {
    let mut parts: Vec<(String, Vec<u8>)> = [
        ("[Content_Types].xml", r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/><Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/></Types>"#),
        ("_rels/.rels", r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/></Relationships>"#),
        ("ppt/presentation.xml", r#"<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:sldIdLst><p:sldId id="256" r:id="rId1"/></p:sldIdLst><p:sldSz cx="9144000" cy="6858000"/></p:presentation>"#),
        ("ppt/_rels/presentation.xml.rels", r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/></Relationships>"#),
        ("ppt/slides/slide1.xml", r#"<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld></p:sld>"#),
    ].into_iter().map(|(path, xml)| (path.into(), xml.as_bytes().to_vec())).collect();
    if with_notes {
        let types = String::from_utf8(parts[0].1.clone()).unwrap().replace("</Types>", r#"<Override PartName="/ppt/notesSlides/notesSlide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml"/><Override PartName="/ppt/notesMasters/notesMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.notesMaster+xml"/></Types>"#);
        parts[0].1 = types.into_bytes();
        for (path, xml) in [
            (
                "ppt/slides/_rels/slide1.xml.rels",
                r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide" Target="../notesSlides/notesSlide1.xml"/></Relationships>"#,
            ),
            (
                "ppt/notesSlides/notesSlide1.xml",
                r#"<p:notes xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/><p:sp><p:nvSpPr><p:cNvPr id="2" name="Notes"/><p:cNvSpPr/><p:nvPr><p:ph type="body"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:pPr marL="1234"/><a:r><a:rPr b="1" lang="en-US"/><a:t>NOTES_MARKER</a:t></a:r><a:endParaRPr lang="de-DE"/></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:notes>"#,
            ),
            (
                "ppt/notesSlides/_rels/notesSlide1.xml.rels",
                r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="../slides/slide1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesMaster" Target="../notesMasters/notesMaster1.xml"/></Relationships>"#,
            ),
            (
                "ppt/notesMasters/notesMaster1.xml",
                r#"<p:notesMaster xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:spTree/></p:cSld></p:notesMaster>"#,
            ),
        ] {
            parts.push((path.into(), xml.as_bytes().to_vec()));
        }
    }
    ooxml_opc::rezip_parts(&parts).unwrap()
}

fn assert_duplicate_notes(bytes: &[u8], source: &[u8], text: &str) {
    let parts = ooxml_opc::unzip_parts(bytes).unwrap();
    let notes: Vec<_> = parts
        .iter()
        .filter(|(path, _)| {
            path.starts_with("ppt/notesSlides/notesSlide") && path.ends_with(".xml")
        })
        .collect();
    assert_eq!(notes.len(), 2);
    let (copy_path, copy_bytes) = notes
        .into_iter()
        .find(|(path, _)| path != "ppt/notesSlides/notesSlide1.xml")
        .unwrap();
    let xml = String::from_utf8(copy_bytes.clone()).unwrap();
    for markup in [
        "marL=\"1234\"",
        "b=\"1\"",
        "lang=\"en-US\"",
        "<a:endParaRPr lang=\"de-DE\"/>",
        text,
    ] {
        assert!(xml.contains(markup), "missing {markup}: {xml}");
    }
    let deck = DeckSession::open(bytes, 9).unwrap().snapshot().unwrap();
    assert_eq!(deck.slides[1].notes, text);
    let copy_slide = deck.slides[1]
        .source_part_path
        .as_deref()
        .unwrap()
        .rsplit('/')
        .next()
        .unwrap();
    let notes_name = copy_path.rsplit('/').next().unwrap();
    let rels = slide_xml(bytes, &format!("ppt/notesSlides/_rels/{notes_name}.rels"));
    assert!(rels.contains(&format!("Target=\"../slides/{copy_slide}\"")));
    assert!(rels.contains("Target=\"../notesMasters/notesMaster1.xml\""));
    let rels = slide_xml(bytes, &format!("ppt/slides/_rels/{copy_slide}.rels"));
    assert!(rels.contains(&format!("Target=\"../notesSlides/{notes_name}\"")));
    assert!(!rels.contains("/comments"));
    assert!(
        slide_xml(bytes, "[Content_Types].xml").contains(&format!("PartName=\"/{copy_path}\""))
    );
    assert_eq!(
        slide_xml(bytes, "ppt/notesSlides/notesSlide1.xml"),
        slide_xml(source, "ppt/notesSlides/notesSlide1.xml")
    );
}

#[test]
fn duplicate_notes_preserves_markup_without_edits() {
    let source = notes_source(true);
    let deck = DeckSession::open(&source, 7).unwrap();
    let slide = deck.snapshot().unwrap().slides[0].id.clone();
    deck.duplicate_slide(&EditCtx::local("test"), &slide, 1)
        .unwrap();
    assert_duplicate_notes(&deck.save().unwrap(), &source, "NOTES_MARKER");
}

#[test]
fn duplicate_notes_preserves_markup_after_edit() {
    let source = notes_source(true);
    let deck = DeckSession::open(&source, 7).unwrap();
    let ctx = EditCtx::local("test");
    let slide = deck.snapshot().unwrap().slides[0].id.clone();
    let copy = deck.duplicate_slide(&ctx, &slide, 1).unwrap();
    deck.set_slide_notes(&ctx, &copy.slide_id, "NOTES_MARKER appended")
        .unwrap();
    assert_duplicate_notes(&deck.save().unwrap(), &source, "NOTES_MARKER appended");
}

#[test]
fn duplicate_notes_preserves_markup_after_reopen_and_edit() {
    let source = notes_source(true);
    let deck = DeckSession::open(&source, 7).unwrap();
    let ctx = EditCtx::local("test");
    let slide = deck.snapshot().unwrap().slides[0].id.clone();
    let copy = deck.duplicate_slide(&ctx, &slide, 1).unwrap();
    deck.set_slide_notes(&ctx, &copy.slide_id, "NOTES_MARKER appended")
        .unwrap();
    let saved = deck.save().unwrap();
    let reopened = DeckSession::open(&saved, 8).unwrap();
    let copy = reopened.snapshot().unwrap().slides[1].id.clone();
    reopened
        .set_slide_notes(&ctx, &copy, "NOTES_MARKER appended again")
        .unwrap();
    assert_duplicate_notes(
        &reopened.save().unwrap(),
        &source,
        "NOTES_MARKER appended again",
    );
}

#[test]
fn duplicate_notes_absent_source_mints_no_notes() {
    let deck = DeckSession::open(&notes_source(false), 7).unwrap();
    let slide = deck.snapshot().unwrap().slides[0].id.clone();
    deck.duplicate_slide(&EditCtx::local("test"), &slide, 1)
        .unwrap();
    let parts = ooxml_opc::unzip_parts(&deck.save().unwrap()).unwrap();
    assert!(
        !parts
            .iter()
            .any(|(path, _)| path.starts_with("ppt/notesSlides/"))
    );
}

#[test]
fn duplicate_notes_undo_leaves_no_orphan_part() {
    let source = notes_source(true);
    let deck = DeckSession::open(&source, 7).unwrap();
    let before = deck.snapshot().unwrap();
    deck.duplicate_slide(&EditCtx::local("test"), &before.slides[0].id, 1)
        .unwrap();
    deck.save().unwrap();
    assert!(deck.undo());
    assert_eq!(deck.snapshot().unwrap(), before);
    let saved = deck.save().unwrap();
    let parts = ooxml_opc::unzip_parts(&saved).unwrap();
    assert_eq!(
        parts
            .iter()
            .filter(|(path, _)| path.starts_with("ppt/notesSlides/"))
            .count(),
        2
    );
    assert_eq!(
        slide_xml(&saved, "ppt/notesSlides/notesSlide1.xml"),
        slide_xml(&source, "ppt/notesSlides/notesSlide1.xml")
    );
}

#[test]
fn set_story_paragraphs_replaces_multiple_paragraphs_and_preserves_story_identity() {
    let deck = DeckSession::open(SOURCE, 7).unwrap();
    let ctx = EditCtx::local("test");
    let snapshot = deck.snapshot().unwrap();
    let story = snapshot
        .slides
        .iter()
        .flat_map(|s| &s.shapes)
        .flat_map(|s| &s.text_stories)
        .next()
        .unwrap();
    let paragraphs = vec![
        TextParagraphDraft {
            bullet: None,
            alignment: Some("ctr".into()),
            runs: vec![TextRunDraft {
                text: "New title 😀".into(),
                style: TextStyle {
                    bold: Some(true),
                    font_size_pt: Some(24.0),
                    ..TextStyle::default()
                },
            }],
        },
        TextParagraphDraft {
            bullet: None,
            alignment: None,
            runs: vec![TextRunDraft {
                text: "Second paragraph".into(),
                style: TextStyle::default(),
            }],
        },
    ];
    let receipt = deck
        .set_story_paragraphs(&ctx, &story.id, &paragraphs)
        .unwrap();
    assert_eq!(receipt.id, story.id);
    assert_eq!(receipt.paragraphs.len(), 2);
    assert_eq!(receipt.plain_text(), "New title 😀\nSecond paragraph");
    let restored = DeckSession::open(&deck.save().unwrap(), 9).unwrap();
    let current = restored.snapshot().unwrap();
    let text = current
        .slides
        .iter()
        .flat_map(|s| &s.shapes)
        .flat_map(|s| &s.text_stories)
        .find(|s| s.plain_text().starts_with("New title"))
        .unwrap();
    assert_eq!(text.paragraphs[0].alignment.as_deref(), Some("ctr"));
    assert_eq!(text.paragraphs[0].runs[0].style.bold, Some(true));
    let version = deck.version();
    let invalid = vec![TextParagraphDraft {
        bullet: None,
        alignment: None,
        runs: vec![TextRunDraft {
            text: "invalid\nrun".into(),
            style: TextStyle::default(),
        }],
    }];
    assert!(
        deck.set_story_paragraphs(&ctx, &story.id, &invalid)
            .is_err()
    );
    assert_eq!(deck.version(), version);
}
