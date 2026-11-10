use docx_edit::{EngineSession, StoryInfoKind, seed_from_docx};
use serde_json::{Value, json};

const W: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const REL: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

fn integral_numbers(value: &mut Value) {
    match value {
        Value::Number(number) => {
            if let Some(float) = number.as_f64().filter(|_| number.is_f64())
                && float.fract() == 0.0
                && float.abs() < 9.0e15
            {
                *value = json!(float as i64);
            }
        }
        Value::Array(items) => items.iter_mut().for_each(integral_numbers),
        Value::Object(fields) => fields.values_mut().for_each(integral_numbers),
        _ => {}
    }
}

fn engine_faces(bytes: &[u8]) -> Vec<(String, bool, bool)> {
    let engine = EngineSession::new(1);
    seed_from_docx(engine.doc(), bytes).unwrap();
    let sections: Vec<_> = engine
        .doc()
        .list_sections()
        .unwrap()
        .into_iter()
        .map(|section| {
            json!({"sectionId": section.closing_paragraph_id, "properties": section.authored})
        })
        .collect();
    let mut notes = Vec::new();
    for story in engine.doc().list_stories().stories {
        let (prefix, kind) = match story.kind {
            StoryInfoKind::Footnote => ("fn:", "footnote"),
            StoryInfoKind::Endnote => ("en:", "endnote"),
            _ => continue,
        };
        let id: i64 = story.story.strip_prefix(prefix).unwrap().parse().unwrap();
        notes.push(json!({"id": id, "noteKind": kind, "height": 0}));
    }
    let parts =
        ooxml_opc::unzip_parts_where(bytes, u64::MAX, |name| name == "word/settings.xml").unwrap();
    let limits = docx_parse::xml::ParseLimits::default();
    let settings = docx_parse::parse_settings(
        parts.first().map(|(_, bytes)| bytes.as_slice()),
        "word/settings.xml",
        &mut docx_parse::xml::ParseBudget::new(&limits),
    )
    .unwrap();
    let mut request = json!({
        "bodyStory": "body",
        "regions": {"sections": sections, "settings": settings},
        "notes": {"contents": notes},
        "renderEnv": {},
    });
    integral_numbers(&mut request);
    let requirements: Vec<Value> = serde_json::from_str(
        &engine
            .layout_font_requirements_json(&request.to_string())
            .unwrap(),
    )
    .unwrap();
    requirements
        .iter()
        .map(|face| {
            (
                face["family"].as_str().unwrap().to_owned(),
                face["bold"].as_bool().unwrap(),
                face["italic"].as_bool().unwrap(),
            )
        })
        .collect()
}

fn has(faces: &[(String, bool, bool)], family: &str) -> bool {
    faces
        .iter()
        .any(|(name, bold, italic)| name.eq_ignore_ascii_case(family) && !bold && !italic)
}

fn assert_parity(bytes: &[u8], wanted: &str) {
    let engine = engine_faces(bytes);
    assert!(
        has(&engine, wanted),
        "engine does not request {wanted}: {engine:?}"
    );
    let preflight = ooxml_font_preflight::docx_font_preflight(bytes).unwrap();
    let answered: Vec<_> = preflight
        .faces
        .iter()
        .map(|face| (face.family.clone(), face.bold, face.italic))
        .collect();
    if preflight.complete {
        for (family, bold, italic) in &engine {
            assert!(
                answered
                    .iter()
                    .any(|(name, b, i)| name.eq_ignore_ascii_case(family)
                        && b == bold
                        && i == italic),
                "complete preflight misses ({family}, {bold}, {italic}) the engine requests: {answered:?}"
            );
        }
    }
    assert!(
        has(&answered, wanted),
        "preflight does not name {wanted}: {answered:?}"
    );
}

fn rels(entries: &[(&str, &str)]) -> String {
    let body: String = entries
        .iter()
        .enumerate()
        .map(|(index, (kind, target))| {
            format!(r#"<Relationship Id="rId{index}" Type="{REL}/{kind}" Target="{target}"/>"#)
        })
        .collect();
    format!(
        r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">{body}</Relationships>"#
    )
}

fn document(body: &str) -> String {
    format!(r#"<w:document xmlns:w="{W}" xmlns:r="{REL}"><w:body>{body}</w:body></w:document>"#)
}

fn styles(defaults: &str, body: &str) -> String {
    format!(
        r#"<w:styles xmlns:w="{W}"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="{defaults}" w:hAnsi="{defaults}"/></w:rPr></w:rPrDefault></w:docDefaults>{body}</w:styles>"#
    )
}

fn package(parts: &[(&str, String)]) -> Vec<u8> {
    let mut entries = vec![(
        "[Content_Types].xml".to_owned(),
        br#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/></Types>"#.to_vec(),
    )];
    if !parts.iter().any(|(name, _)| *name == "_rels/.rels") {
        entries.push((
            "_rels/.rels".to_owned(),
            rels(&[("officeDocument", "word/document.xml")]).into_bytes(),
        ));
    }
    entries.extend(
        parts
            .iter()
            .map(|(name, xml)| ((*name).to_owned(), xml.as_bytes().to_vec())),
    );
    ooxml_opc::rezip_parts(&entries).unwrap()
}

fn plain(text: &str) -> String {
    format!("<w:p><w:r><w:t>{text}</w:t></w:r></w:p>")
}

#[test]
fn notes_at_the_conventional_path_without_a_relationship() {
    let notes = |root: &str, item: &str, family: &str| {
        format!(
            r#"<w:{root} xmlns:w="{W}"><w:{item} w:id="-1" w:type="separator"><w:p/></w:{item}><w:{item} w:id="0" w:type="continuationSeparator"><w:p/></w:{item}><w:{item} w:id="1"><w:p><w:r><w:rPr><w:rFonts w:ascii="{family}" w:hAnsi="{family}"/></w:rPr><w:t>Note</w:t></w:r></w:p></w:{item}></w:{root}>"#
        )
    };
    for (path, xml, family) in [
        (
            "word/footnotes.xml",
            notes("footnotes", "footnote", "NoteFont"),
            "NoteFont",
        ),
        (
            "word/endnotes.xml",
            notes("endnotes", "endnote", "EndFont"),
            "EndFont",
        ),
    ] {
        let bytes = package(&[
            ("word/document.xml", document(&plain("Body"))),
            (
                "word/_rels/document.xml.rels",
                rels(&[("styles", "styles.xml")]),
            ),
            ("word/styles.xml", styles("Base", "")),
            (path, xml),
        ]);
        assert_parity(&bytes, family);
    }
}

#[test]
fn main_part_named_by_the_package_relationship() {
    let bytes = package(&[
        (
            "_rels/.rels",
            rels(&[("officeDocument", "word/document2.xml")]),
        ),
        (
            "word/document2.xml",
            document(
                r#"<w:p><w:r><w:rPr><w:rFonts w:ascii="MissingFont" w:hAnsi="MissingFont"/></w:rPr><w:t>Real</w:t></w:r></w:p>"#,
            ),
        ),
        (
            "word/_rels/document2.xml.rels",
            rels(&[("styles", "styles.xml")]),
        ),
        ("word/document.xml", document(&plain("Decoy"))),
        (
            "word/_rels/document.xml.rels",
            rels(&[("styles", "styles.xml")]),
        ),
        ("word/styles.xml", styles("Base", "")),
    ]);
    assert_parity(&bytes, "MissingFont");
}

#[test]
fn duplicate_numbering_ids_take_the_last_definition() {
    let level = |id: u32, family: &str| {
        format!(
            r#"<w:abstractNum w:abstractNumId="{id}"><w:lvl w:ilvl="0"><w:numFmt w:val="bullet"/><w:lvlText w:val="*"/><w:rPr><w:rFonts w:ascii="{family}" w:hAnsi="{family}"/></w:rPr></w:lvl></w:abstractNum>"#
        )
    };
    let num = |num: u32, abs: u32| {
        format!(r#"<w:num w:numId="{num}"><w:abstractNumId w:val="{abs}"/></w:num>"#)
    };
    for (definitions, wanted) in [
        (
            format!(
                "{}{}{}{}",
                level(0, "BulletA"),
                level(1, "BulletB"),
                num(1, 0),
                num(1, 1)
            ),
            "BulletB",
        ),
        (
            format!(
                "{}{}{}",
                level(0, "BulletA"),
                level(0, "BulletC"),
                num(1, 0)
            ),
            "BulletC",
        ),
    ] {
        let bytes = package(&[
            (
                "word/document.xml",
                document(
                    r#"<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>Item</w:t></w:r></w:p>"#,
                ),
            ),
            (
                "word/_rels/document.xml.rels",
                rels(&[("styles", "styles.xml"), ("numbering", "numbering.xml")]),
            ),
            ("word/styles.xml", styles("Base", "")),
            (
                "word/numbering.xml",
                format!(r#"<w:numbering xmlns:w="{W}">{definitions}</w:numbering>"#),
            ),
        ]);
        assert_parity(&bytes, wanted);
    }
}

#[test]
fn duplicate_theme_script_fonts_take_the_last_entry() {
    let theme = r#"<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:themeElements><a:fontScheme><a:majorFont><a:latin typeface="Heading"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont><a:minorFont><a:latin typeface="Body"/><a:ea typeface=""/><a:cs typeface=""/><a:font script="Jpan" typeface="FirstFont"/><a:font script="Jpan" typeface="SecondFont"/></a:minorFont></a:fontScheme></a:themeElements></a:theme>"#;
    let bytes = package(&[
        (
            "word/document.xml",
            document(
                r#"<w:p><w:r><w:rPr><w:rFonts w:eastAsiaTheme="minorEastAsia"/></w:rPr><w:t>漢字</w:t></w:r></w:p>"#,
            ),
        ),
        (
            "word/_rels/document.xml.rels",
            rels(&[
                ("styles", "styles.xml"),
                ("theme", "theme/theme1.xml"),
                ("settings", "settings.xml"),
            ]),
        ),
        ("word/styles.xml", styles("Base", "")),
        ("word/theme/theme1.xml", theme.to_owned()),
        (
            "word/settings.xml",
            format!(
                r#"<w:settings xmlns:w="{W}"><w:themeFontLang w:val="en-US" w:eastAsia="ja-JP"/></w:settings>"#
            ),
        ),
    ]);
    assert_parity(&bytes, "SecondFont");
}

#[test]
fn duplicate_style_id_takes_the_default_flag_of_the_last_definition() {
    let bytes = package(&[
        ("word/document.xml", document(&plain("Text"))),
        (
            "word/_rels/document.xml.rels",
            rels(&[("styles", "styles.xml")]),
        ),
        (
            "word/styles.xml",
            styles(
                "Base",
                r#"<w:style w:type="paragraph" w:default="1" w:styleId="A"><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/></w:rPr></w:style><w:style w:type="paragraph" w:styleId="A"><w:rPr><w:rFonts w:ascii="Courier" w:hAnsi="Courier"/></w:rPr></w:style>"#,
            ),
        ),
    ]);
    assert_parity(&bytes, "Base");
}

#[test]
fn table_without_a_grid_counts_grid_before_in_its_columns() {
    let bytes = package(&[
        (
            "word/document.xml",
            document(
                r#"<w:tbl><w:tblPr><w:tblStyle w:val="T"/><w:tblLook w:val="0180"/></w:tblPr><w:tr><w:trPr><w:gridBefore w:val="1"/></w:trPr><w:tc><w:p><w:r><w:t>Cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:p/>"#,
            ),
        ),
        (
            "word/_rels/document.xml.rels",
            rels(&[("styles", "styles.xml")]),
        ),
        (
            "word/styles.xml",
            styles(
                "Base",
                r#"<w:style w:type="table" w:styleId="T"><w:tblStylePr w:type="lastCol"><w:rPr><w:rFonts w:ascii="LastColFont" w:hAnsi="LastColFont"/></w:rPr></w:tblStylePr></w:style>"#,
            ),
        ),
    ]);
    assert_parity(&bytes, "LastColFont");
}
