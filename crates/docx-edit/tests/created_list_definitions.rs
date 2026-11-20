use std::collections::{BTreeMap, HashMap};

use docx_edit::bridge::{RenderEnv, yrs_doc_to_layout_blocks};
use docx_edit::{
    CreatedListDefinition, EditCtx, EditingDoc, LIST_DEF, ListIdentity, ListKind, UndoCaptureMode,
    resolve_list_identity, seed_from_docx,
};
use docx_layout::types::LayoutBlock;
use yrs::Any;

const W: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const LISTS_NUMBERING: &str = include_str!(
    "../../../packages/docx/src/yrs/__fixtures__/seed-parity/lists/word/numbering.xml"
);
const DECIMAL_NUMBERING: &str = concat!(
    r#"<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">"#,
    r#"<w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl></w:abstractNum>"#,
    r#"<w:abstractNum w:abstractNumId="1"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl></w:abstractNum>"#,
    r#"<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>"#,
    r#"<w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num></w:numbering>"#
);
const LIST_KEYS: [&str; 7] = [
    "numPr",
    LIST_DEF,
    "listMarker",
    "listIsBullet",
    "listNumFmt",
    "listLevelNumFmts",
    "listAbstractNumId",
];

fn docx(body: &str, numbering: &str) -> Vec<u8> {
    let parts = vec![
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
            numbering.as_bytes().to_vec(),
        ),
    ];
    ooxml_opc::rezip_parts(&parts).unwrap()
}

fn paragraph(text: &str, num_id: Option<u32>) -> String {
    let properties = num_id.map_or_else(String::new, |num_id| {
        format!(
            r#"<w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="{num_id}"/></w:numPr></w:pPr>"#
        )
    });
    format!("<w:p>{properties}<w:r><w:t>{text}</w:t></w:r></w:p>")
}

fn lists_docx() -> Vec<u8> {
    let body = [
        paragraph("Lead", None),
        paragraph("Alpha", None),
        paragraph("Gap", None),
        paragraph("Beta", None),
        paragraph("Tail", None),
        paragraph("Decimal source", Some(1)),
        paragraph("Bullet source", Some(2)),
    ]
    .concat();
    docx(&body, LISTS_NUMBERING)
}

fn seeded(client: u64, bytes: &[u8]) -> EditingDoc {
    let doc = EditingDoc::new(client);
    seed_from_docx(&doc, bytes).unwrap();
    doc
}

fn replica(of: &EditingDoc, client: u64) -> EditingDoc {
    let doc = EditingDoc::new(client);
    doc.apply_update_v1(&of.encode_state_as_update_v1())
        .unwrap();
    doc
}

fn sync(left: &EditingDoc, right: &EditingDoc) {
    let left_update = left.encode_state_as_update_v1();
    let right_update = right.encode_state_as_update_v1();
    left.apply_update_v1(&right_update).unwrap();
    right.apply_update_v1(&left_update).unwrap();
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

fn list_attrs(doc: &EditingDoc, text: &str) -> BTreeMap<String, Any> {
    attrs(doc, text)
        .into_iter()
        .filter(|(key, _)| LIST_KEYS.contains(&key.as_str()))
        .collect()
}

fn created_key(doc: &EditingDoc, text: &str) -> Option<String> {
    match resolve_list_identity(&attrs(doc, text)) {
        Some(ListIdentity::Created { key, .. }) => Some(key),
        _ => None,
    }
}

fn kind_of(definitions: &[CreatedListDefinition], key: &str) -> Option<ListKind> {
    definitions
        .iter()
        .find(|definition| definition.key == key)
        .map(|definition| definition.kind)
}

/// Puts the paragraph `text` on a list of `kind` as a command would: the attributes the engine
/// hands back are written on the paragraph.
fn create(doc: &EditingDoc, text: &str, kind: ListKind) {
    let id = paragraph_id(doc, text);
    let attrs = doc
        .ensure_list_definition(&EditCtx::local("", ""), "body", &id, kind)
        .unwrap();
    for (key, value) in attrs {
        doc.set_paragraph_attr(&id, key, value).unwrap();
    }
}

fn num_pr(num_id: f64) -> Any {
    Any::from(HashMap::from([
        ("numId".to_owned(), Any::Number(num_id)),
        ("ilvl".to_owned(), Any::Number(0.0)),
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
fn lists_created_apart_keep_their_own_definitions_once_merged() {
    for (left_kind, right_kind) in [
        (ListKind::Bullet, ListKind::Numbered),
        (ListKind::Bullet, ListKind::Bullet),
    ] {
        let left = seeded(401, &lists_docx());
        let right = replica(&left, 402);
        create(&left, "Alpha", left_kind);
        create(&right, "Beta", right_kind);
        sync(&left, &right);
        for doc in [&left, &right] {
            let definitions = doc.created_list_definitions();
            assert_eq!(definitions.len(), 2, "{left_kind:?} and {right_kind:?}");
            let alpha = created_key(doc, "Alpha").unwrap_or_default();
            let beta = created_key(doc, "Beta").unwrap_or_default();
            assert_ne!(alpha, beta);
            assert_eq!(kind_of(&definitions, &alpha), Some(left_kind));
            assert_eq!(kind_of(&definitions, &beta), Some(right_kind));
        }
        assert_eq!(
            left.created_list_definitions(),
            right.created_list_definitions()
        );
    }
}

#[test]
fn undoing_a_creation_removes_the_definition_and_the_reference_in_one_step() {
    let doc = seeded(301, &lists_docx());
    let mut undo = doc.undo_manager();
    undo.set_capture_mode(UndoCaptureMode::Manual);
    let before = list_attrs(&doc, "Alpha");
    create(&doc, "Alpha", ListKind::Bullet);
    undo.add_undo_barrier();
    let created = list_attrs(&doc, "Alpha");
    assert_eq!(doc.created_list_definitions().len(), 1);
    assert!(created.contains_key(LIST_DEF));
    assert_eq!(undo.undo_depth(), 1);
    assert!(undo.undo());
    assert!(doc.created_list_definitions().is_empty());
    assert_eq!(list_attrs(&doc, "Alpha"), before);
    assert!(undo.redo());
    assert_eq!(doc.created_list_definitions().len(), 1);
    assert_eq!(list_attrs(&doc, "Alpha"), created);
}

#[test]
fn a_numbering_change_from_an_older_peer_makes_the_paragraph_a_legacy_list_paragraph() {
    for replacement in [Some(3.0), None] {
        let doc = seeded(501, &lists_docx());
        create(&doc, "Alpha", ListKind::Numbered);
        assert!(created_key(&doc, "Alpha").is_some(), "{replacement:?}");
        let older = replica(&doc, 502);
        let alpha = paragraph_id(&older, "Alpha");
        older
            .set_paragraph_attr(&alpha, "numPr", replacement.map_or(Any::Null, num_pr))
            .unwrap();
        sync(&doc, &older);
        let toggled = attrs(&doc, "Alpha");
        assert!(toggled.contains_key(LIST_DEF), "{replacement:?}");
        assert_eq!(
            resolve_list_identity(&toggled),
            replacement.map(|num_id| ListIdentity::Source { num_id, level: 0.0 })
        );
    }
}

#[test]
fn created_lists_mark_and_count_their_own_paragraphs() {
    let body = [
        paragraph("Source two first", Some(2)),
        paragraph("Gap one", None),
        paragraph("Beta", None),
        paragraph("Tail", None),
        paragraph("Gap two", None),
        paragraph("Alpha", None),
        paragraph("Gap three", None),
        paragraph("Source two second", Some(2)),
        paragraph("Source one", Some(1)),
    ]
    .concat();
    let doc = seeded(601, &docx(&body, DECIMAL_NUMBERING));
    create(&doc, "Beta", ListKind::Numbered);
    create(&doc, "Tail", ListKind::Numbered);
    create(&doc, "Alpha", ListKind::Bullet);
    let definitions = doc.created_list_definitions();
    assert_eq!(definitions.len(), 2);
    for definition in &definitions {
        assert_eq!(definition.level_texts.len(), 9, "{}", definition.key);
    }
    assert!(created_key(&doc, "Beta").is_some());
    assert_eq!(created_key(&doc, "Tail"), created_key(&doc, "Beta"));
    let alpha = created_key(&doc, "Alpha").unwrap_or_default();
    assert_eq!(kind_of(&definitions, &alpha), Some(ListKind::Bullet));
    let bullet = definitions
        .iter()
        .find(|definition| definition.key == alpha)
        .and_then(|definition| definition.level_texts.first())
        .map(String::as_str)
        .map(docx_parse::convert_bullet_to_unicode);
    assert!(bullet.as_deref().is_some_and(|bullet| !bullet.is_empty()));
    assert_eq!(
        markers(&doc),
        [
            Some("1.".to_owned()),
            None,
            Some("1.".to_owned()),
            Some("2.".to_owned()),
            None,
            bullet,
            None,
            Some("2.".to_owned()),
            Some("1.".to_owned()),
        ]
    );
}
