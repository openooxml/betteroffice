use std::collections::{BTreeMap, HashMap};

use docx_edit::bridge::{RenderEnv, yrs_doc_to_layout_blocks};
use docx_edit::{
    EditCtx, EditingDoc, LIST_DEF, LIST_DEFINITIONS, ListIdentity, ListKind, resolve_list_identity,
    seed_from_docx,
};
use docx_layout::types::LayoutBlock;
use yrs::{Any, Map, Transact};

const W: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const NUMBERING: &str = concat!(
    r#"<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">"#,
    r#"<w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl></w:abstractNum>"#,
    r#"<w:abstractNum w:abstractNumId="1"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl></w:abstractNum>"#,
    r#"<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>"#,
    r#"<w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num></w:numbering>"#
);
const STYLED_TYPES: &str = concat!(
    r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>"#,
    r#"<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>"#,
    r#"<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>"#
);
const STYLED_RELATIONSHIPS: &str = concat!(
    r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">"#,
    r#"<Relationship Id="numbering" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>"#,
    r#"<Relationship Id="styles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>"#
);
/// A Heading 1 style numbered by list 1.
const NUMBERED_HEADING_STYLES: &str = concat!(
    r#"<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">"#,
    r#"<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>"#,
    r#"<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/>"#,
    r#"<w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr><w:outlineLvl w:val="0"/></w:pPr></w:style></w:styles>"#
);
const SHAPES: &str =
    include_str!("../../../packages/docx/src/yrs/__fixtures__/created-list-shapes.json");

fn docx(paragraphs: &[(&str, Option<u32>)]) -> Vec<u8> {
    let body: String = paragraphs
        .iter()
        .map(|(text, num_id)| {
            let properties = num_id.map_or_else(String::new, |num_id| {
                format!(
                    r#"<w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="{num_id}"/></w:numPr></w:pPr>"#
                )
            });
            format!("<w:p>{properties}<w:r><w:t>{text}</w:t></w:r></w:p>")
        })
        .collect();
    styled_docx(&body, None)
}

/// A package holding `body` and [`NUMBERING`], with the styles part `styles` when given.
fn styled_docx(body: &str, styles: Option<&str>) -> Vec<u8> {
    let mut parts = vec![
        (
            "[Content_Types].xml".to_owned(),
            br#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/></Types>"#.to_vec(),
        ),
        (
            "_rels/.rels".to_owned(),
            br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="main" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"#.to_vec(),
        ),
        (
            "word/_rels/document.xml.rels".to_owned(),
            br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="numbering" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/></Relationships>"#.to_vec(),
        ),
        (
            "word/document.xml".to_owned(),
            format!(
                r#"<w:document xmlns:w="{W}"><w:body>{body}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>"#
            )
            .into_bytes(),
        ),
        (
            "word/numbering.xml".to_owned(),
            NUMBERING.as_bytes().to_vec(),
        ),
    ];
    if let Some(styles) = styles {
        parts[0].1 = STYLED_TYPES.as_bytes().to_vec();
        parts[2].1 = STYLED_RELATIONSHIPS.as_bytes().to_vec();
        parts.push(("word/styles.xml".to_owned(), styles.as_bytes().to_vec()));
    }
    ooxml_opc::rezip_parts(&parts).unwrap()
}

fn seeded(client: u64, paragraphs: &[(&str, Option<u32>)]) -> EditingDoc {
    let doc = EditingDoc::new(client);
    seed_from_docx(&doc, &docx(paragraphs)).unwrap();
    doc
}

fn paragraph_id(doc: &EditingDoc, text: &str) -> String {
    doc.paragraphs("body")
        .unwrap()
        .into_iter()
        .find(|paragraph| paragraph.text == text)
        .unwrap()
        .para_id
}

fn attrs(doc: &EditingDoc, text: &str) -> BTreeMap<String, Any> {
    doc.paragraphs("body")
        .unwrap()
        .into_iter()
        .find(|paragraph| paragraph.text == text)
        .unwrap()
        .properties
}

fn created_key(doc: &EditingDoc, text: &str) -> Option<String> {
    match resolve_list_identity(&attrs(doc, text)) {
        Some(ListIdentity::Created { key, .. }) => Some(key),
        _ => None,
    }
}

/// Puts the paragraph `text` on a list of `kind` as a command would, returning the attributes.
fn create(doc: &EditingDoc, text: &str, kind: ListKind) -> BTreeMap<String, Any> {
    let id = paragraph_id(doc, text);
    let attrs = doc
        .ensure_list_definition(&EditCtx::local("", ""), "body", &id, kind)
        .unwrap();
    for (key, value) in attrs.clone() {
        doc.set_paragraph_attr(&id, key, value).unwrap();
    }
    attrs
}

fn num_pr(num_id: f64, level: f64) -> Any {
    Any::from(HashMap::from([
        ("numId".to_owned(), Any::Number(num_id)),
        ("ilvl".to_owned(), Any::Number(level)),
    ]))
}

fn markers(doc: &EditingDoc) -> Vec<Option<String>> {
    yrs_doc_to_layout_blocks(doc, "body", &RenderEnv::default())
        .unwrap()
        .into_iter()
        .filter_map(|block| match block {
            LayoutBlock::Paragraph(paragraph) => Some(paragraph),
            _ => None,
        })
        .map(|paragraph| paragraph.attrs.and_then(|attrs| attrs.list_marker))
        .collect()
}

#[test]
fn an_indented_numbered_item_shows_the_marker_of_its_level() {
    let doc = seeded(701, &[("Alpha", None), ("Beta", None)]);
    create(&doc, "Alpha", ListKind::Numbered);
    create(&doc, "Beta", ListKind::Numbered);
    assert_eq!(created_key(&doc, "Beta"), created_key(&doc, "Alpha"));
    doc.set_paragraph_attr(&paragraph_id(&doc, "Beta"), "numPr", num_pr(2.0, 1.0))
        .unwrap();
    assert!(created_key(&doc, "Beta").is_some());
    assert_eq!(
        markers(&doc),
        [Some("1.".to_owned()), Some("a.".to_owned())]
    );
}

#[test]
fn an_indented_bullet_takes_the_bullet_of_its_level_and_none_past_level_eight() {
    let doc = seeded(702, &[("Alpha", None), ("Beta", None)]);
    create(&doc, "Alpha", ListKind::Bullet);
    create(&doc, "Beta", ListKind::Bullet);
    let beta = paragraph_id(&doc, "Beta");
    doc.set_paragraph_attr(&beta, "numPr", num_pr(1.0, 1.0))
        .unwrap();
    let bullets = doc.created_list_definitions()[0].level_texts.clone();
    let shown = |level: usize| Some(docx_parse::convert_bullet_to_unicode(&bullets[level]));
    assert_eq!(markers(&doc), [shown(0), shown(1)]);
    assert_ne!(shown(0), shown(1));
    doc.set_paragraph_attr(&beta, "numPr", num_pr(1.0, 9.0))
        .unwrap();
    assert_eq!(markers(&doc), [shown(0), None]);
}

#[test]
fn a_key_whose_definition_is_absent_counts_with_the_list_its_num_id_names() {
    let doc = seeded(703, &[("Source", Some(2)), ("Orphan", None)]);
    let orphan = paragraph_id(&doc, "Orphan");
    let reference = HashMap::from([
        ("key".to_owned(), Any::from("77:0")),
        ("numId".to_owned(), Any::Number(2.0)),
    ]);
    for (key, value) in [
        ("numPr", num_pr(2.0, 0.0)),
        (LIST_DEF, Any::from(reference)),
        ("listMarker", Any::from("%1.")),
        ("listNumFmt", Any::from("decimal")),
        ("listIsBullet", Any::Bool(false)),
    ] {
        doc.set_paragraph_attr(&orphan, key, value).unwrap();
    }
    assert!(doc.created_list_definitions().is_empty());
    assert_eq!(
        markers(&doc),
        [Some("1.".to_owned()), Some("2.".to_owned())]
    );
}

#[test]
fn ensuring_the_list_a_paragraph_is_on_creates_nothing() {
    let doc = seeded(704, &[("Lead", None), ("Alpha", None), ("Gap", None)]);
    let first = create(&doc, "Alpha", ListKind::Bullet);
    let again = create(&doc, "Alpha", ListKind::Bullet);
    assert_eq!(again.get(LIST_DEF), first.get(LIST_DEF));
    assert_eq!(doc.created_list_definitions().len(), 1);
    let bullet = created_key(&doc, "Alpha");
    create(&doc, "Alpha", ListKind::Numbered);
    let definitions = doc.created_list_definitions();
    assert_eq!(definitions.len(), 2);
    let numbered = created_key(&doc, "Alpha");
    assert!(numbered.is_some());
    assert_ne!(numbered, bullet);
    assert_eq!(
        definitions
            .iter()
            .find(|definition| Some(&definition.key) == numbered.as_ref())
            .map(|definition| definition.kind),
        Some(ListKind::Numbered)
    );
}

#[test]
fn ensuring_a_list_where_a_key_has_no_definition_treats_the_paragraph_by_its_num_id() {
    let doc = seeded(706, &[("Lead", None), ("Orphan", None), ("Tail", None)]);
    let orphan = paragraph_id(&doc, "Orphan");
    let reference = HashMap::from([
        ("key".to_owned(), Any::from("88:0")),
        ("numId".to_owned(), Any::Number(2.0)),
    ]);
    for (key, value) in [
        ("numPr", num_pr(2.0, 0.0)),
        (LIST_DEF, Any::from(reference)),
        ("listIsBullet", Any::Bool(false)),
    ] {
        doc.set_paragraph_attr(&orphan, key, value).unwrap();
    }
    create(&doc, "Orphan", ListKind::Numbered);
    let legacy = ListIdentity::Source {
        num_id: 2.0,
        level: 0.0,
    };
    assert_eq!(resolve_list_identity(&attrs(&doc, "Orphan")), Some(legacy));
    assert!(doc.created_list_definitions().is_empty());
}

#[test]
fn definitions_from_a_peer_are_read_as_nine_levels_of_xml_text() {
    let doc = EditingDoc::new(705);
    let root = doc.yrs_doc().get_or_insert_map(LIST_DEFINITIONS);
    {
        let mut txn = doc.yrs_doc().transact_mut();
        let entries = [
            (
                "705:1",
                r#"{"kind":"bullet","levelTexts":["a\u0001","b","c","d","e","f","g","h","i","j","k","l"]}"#,
            ),
            ("705:2", r#"{"kind":"square","levelTexts":["a"]}"#),
            ("705:3", r#"{"kind":"bullet","levelTexts":"a"}"#),
            ("705:4", r#"{"kind":"numbered","levelTexts":["%1)"]}"#),
            ("705:5", "5"),
        ];
        for (key, json) in entries {
            root.insert(&mut txn, key, Any::from_json(json).unwrap());
        }
    }
    let definitions = doc.created_list_definitions();
    let keys: Vec<&str> = definitions
        .iter()
        .map(|definition| definition.key.as_str())
        .collect();
    assert_eq!(keys, ["705:1", "705:4"]);
    assert_eq!(definitions[0].level_texts.len(), 9);
    assert_eq!(definitions[0].level_texts[0], "a");
    assert_eq!(definitions[0].level_texts[8], "i");
    assert_eq!(definitions[1].level_texts.len(), 9);
    assert_eq!(definitions[1].level_texts[0], "%1)");
    assert_eq!(definitions[1].level_texts[1], "%2.");
}

#[test]
fn a_neighbour_whose_format_is_bullet_is_a_bullet_list() {
    let doc = seeded(707, &[("Source", Some(1)), ("Next", None), ("Tail", None)]);
    let source = paragraph_id(&doc, "Source");
    doc.set_paragraph_attr(&source, "listIsBullet", Any::Null)
        .unwrap();
    doc.set_paragraph_attr(&source, "listNumFmt", Any::from("bullet"))
        .unwrap();
    create(&doc, "Next", ListKind::Bullet);
    let joined = ListIdentity::Source {
        num_id: 1.0,
        level: 0.0,
    };
    assert_eq!(resolve_list_identity(&attrs(&doc, "Next")), Some(joined));
    assert!(doc.created_list_definitions().is_empty());
}

#[test]
fn a_created_list_resets_the_marker_formatting_of_the_list_the_paragraph_left() {
    let doc = seeded(708, &[("Lead", None), ("Alpha", Some(1)), ("Gap", None)]);
    let alpha = paragraph_id(&doc, "Alpha");
    doc.set_paragraph_attr(&alpha, "listMarkerBold", Any::Bool(true))
        .unwrap();
    let ensured = create(&doc, "Alpha", ListKind::Bullet);
    for key in [
        "listMarkerHidden",
        "listMarkerBold",
        "listMarkerItalic",
        "listMarkerColor",
        "listMarkerFontSize",
    ] {
        assert_eq!(ensured.get(key), Some(&Any::Null), "{key}");
    }
}

#[test]
fn a_style_numbered_neighbour_is_not_joined_and_a_direct_one_is() {
    let paragraph = |properties: &str, text: &str| {
        format!("<w:p><w:pPr>{properties}</w:pPr><w:r><w:t>{text}</w:t></w:r></w:p>")
    };
    let body = [
        paragraph(r#"<w:pStyle w:val="Heading1"/>"#, "Heading"),
        paragraph("", "Body"),
        paragraph("", "Gap"),
        paragraph(r#"<w:numPr><w:ilvl w:val="0"/><w:numId w:val="2"/></w:numPr>"#, "Item"),
        paragraph("", "Below"),
    ]
    .concat();
    let doc = EditingDoc::new(709);
    seed_from_docx(&doc, &styled_docx(&body, Some(NUMBERED_HEADING_STYLES)))
        .unwrap();
    let heading = attrs(&doc, "Heading");
    let shapes = format!(
        "numPr={:?} numPrFromStyle={:?} _originalFormatting={:?}",
        heading.get("numPr"),
        heading.get("numPrFromStyle"),
        heading.get("_originalFormatting")
    );
    assert!(heading.contains_key("numPr"), "{shapes}");
    create(&doc, "Body", ListKind::Numbered);
    assert!(created_key(&doc, "Body").is_some(), "{shapes}");
    create(&doc, "Below", ListKind::Numbered);
    let joined = ListIdentity::Source {
        num_id: 2.0,
        level: 0.0,
    };
    assert_eq!(resolve_list_identity(&attrs(&doc, "Below")), Some(joined));
    assert_eq!(doc.created_list_definitions().len(), 1);
}

#[test]
fn the_created_list_rule_over_odd_attribute_shapes() {
    let cases: Vec<serde_json::Value> = serde_json::from_str(SHAPES).unwrap();
    for case in cases {
        let name = case["name"].as_str().unwrap();
        let values: BTreeMap<String, Any> = case["attrs"]
            .as_object()
            .unwrap()
            .iter()
            .map(|(key, value)| (key.clone(), Any::from_json(&value.to_string()).unwrap()))
            .collect();
        let key = match resolve_list_identity(&values) {
            Some(ListIdentity::Created { key, .. }) => Some(key),
            _ => None,
        };
        assert_eq!(key.as_deref(), case["key"].as_str(), "{name}");
    }
}
