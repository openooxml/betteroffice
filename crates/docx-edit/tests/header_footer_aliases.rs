use docx_edit::structured::{
    ExportOptions, PageExportOptions, RevisionView, StoryKind, StorySelection,
};
use docx_edit::{
    EditCtx, EditingDoc, EngineSession, FormatPolicy, Position, seed_from_docx_with_generation,
};
use serde_json::{Value, json};
use yrs::{Map, ReadTxn, Transact};

#[path = "support/header_footer_alias_fixture.rs"]
mod fixture;

const FONT: &[u8] = include_bytes!("../../ooxml-text/tests/fonts/LiberationSans-Regular.ttf");

fn seeded(bytes: &[u8]) -> EditingDoc {
    let doc = EditingDoc::new(41);
    seed_from_docx_with_generation(&doc, bytes, "aliases").unwrap();
    doc
}

fn story_ids(doc: &EditingDoc) -> Vec<String> {
    let txn = doc.yrs_doc().transact();
    let mut ids: Vec<_> = txn
        .get_map("stories")
        .unwrap()
        .keys(&txn)
        .map(str::to_owned)
        .collect();
    ids.sort();
    ids
}

fn alias_package() -> Vec<u8> {
    fixture::package(&[("rId7", "header1.xml"), ("rId9", "header1.xml")], &[])
}

fn text(value: &Value) -> String {
    match value {
        Value::Array(values) => values.iter().map(text).collect(),
        Value::Object(fields) => fields
            .iter()
            .map(|(key, value)| {
                if key == "text" {
                    value.as_str().unwrap_or_default().to_owned()
                } else {
                    text(value)
                }
            })
            .collect(),
        _ => String::new(),
    }
}

fn assert_shared_content(headers: &[(&str, &str)], footers: &[(&str, &str)], kind: StoryKind) {
    let bytes = fixture::package(headers, footers);
    let engine = EngineSession::new(41);
    seed_from_docx_with_generation(engine.doc(), &bytes, "aliases").unwrap();
    let entries = if kind == StoryKind::Header {
        headers
    } else {
        footers
    };
    let root = format!("hf:{}", entries[0].0);
    assert_eq!(story_ids(engine.doc()), ["body", &root]);
    for (id, _) in entries {
        assert_eq!(engine.doc().header_footer_story(id), root);
    }
    engine
        .doc()
        .insert_text(
            &EditCtx::local("Ada", "2030-01-02T03:04:05Z"),
            Position::new(root.clone(), 0),
            "Edited ",
            FormatPolicy::Plain,
        )
        .unwrap();
    let lowered = engine.lower_story_json(&root, &Default::default()).unwrap();
    assert!(text(&serde_json::from_str(&lowered).unwrap()).contains("Edited Shared band"));
    let structured = engine
        .doc()
        .export_structured(&ExportOptions {
            stories: Some(vec![
                StorySelection::Body,
                StorySelection::Headers,
                StorySelection::Footers,
            ]),
            ..ExportOptions::new(RevisionView::Accepted)
        })
        .unwrap()
        .content;
    let stories: Vec<_> = structured
        .stories
        .iter()
        .filter(|story| story.kind == kind)
        .collect();
    assert_eq!(stories.len(), 1);
    assert_eq!(stories[0].story, root);
    assert_eq!(stories[0].uses.len(), entries.len());
    assert_eq!(
        stories[0]
            .uses
            .iter()
            .map(|used| used.section_index)
            .collect::<Vec<_>>(),
        (0..entries.len() as u32).collect::<Vec<_>>()
    );
    assert!(
        text(&serde_json::to_value(&stories[0].blocks).unwrap()).contains("Edited Shared band")
    );
    let package = docx_parse::parse_docx_s9_wire(&bytes, Default::default())
        .unwrap()
        .document
        .package;
    let sections: Vec<_> = package
        .document
        .sections
        .unwrap()
        .into_iter()
        .map(|section| json!({"properties": section.properties}))
        .collect();
    let font = docx_layout::register_measure_font(FONT).unwrap();
    let mut request = json!({
        "bodyStory": "body", "renderEnv": {},
        "regions": {"sections": sections, "settings": package.settings},
        "measurement": {
            "fontChains": {"arial|0|0": [font]},
            "defaults": {"fontFamily": "Arial", "fontSize": 12},
            "authoritativeShaping": true
        }
    });
    let requirements: Vec<Value> = serde_json::from_str(
        &engine
            .layout_font_requirements_json(&request.to_string())
            .unwrap(),
    )
    .unwrap();
    for requirement in requirements {
        request["measurement"]["fontChains"][requirement["key"].as_str().unwrap()] = json!([font]);
    }
    let request = request.to_string();
    let output: Value =
        serde_json::from_str(&engine.layout_document_with_regions_json(&request).unwrap()).unwrap();
    let variants = output["headersFooters"]["variants"].as_array().unwrap();
    assert_eq!(variants.len(), entries.len());
    for (index, (id, _)) in entries.iter().enumerate() {
        let variant = variants
            .iter()
            .find(|variant| variant["rId"] == *id)
            .unwrap();
        assert_eq!(variant["sectionIndex"], index);
        assert!(text(variant).contains("Edited Shared band"));
    }
    let pages = output["layout"]["pages"].as_array().unwrap();
    assert_eq!(pages.len(), entries.len());
    let reference = if kind == StoryKind::Header {
        "headerDefault"
    } else {
        "footerDefault"
    };
    for (page, (id, _)) in pages.iter().zip(entries) {
        assert_eq!(page["headerFooterRefs"][reference], *id);
    }
    let paged = engine
        .export_structured_with_pages_for(
            &PageExportOptions {
                stories: Some(vec![
                    StorySelection::Body,
                    StorySelection::Headers,
                    StorySelection::Footers,
                ]),
                ..PageExportOptions::new(RevisionView::Markup)
            },
            &request,
        )
        .unwrap()
        .content;
    let occurrences: Vec<_> = paged
        .layout
        .occurrences
        .iter()
        .filter(|occurrence| occurrence.story == root)
        .collect();
    assert_eq!(occurrences.len(), entries.len());
    assert!(
        paged.layout.diagnostics.is_empty(),
        "{:?}",
        paged.layout.diagnostics
    );
    for index in 0..entries.len() {
        assert!(
            occurrences
                .iter()
                .any(|occurrence| occurrence.section_index == Some(index as u32))
        );
    }
    let identities = engine.doc().paragraph_identities();
    let band_identities: Vec<_> = identities.paragraphs.iter().filter(|identity| {
        matches!(&identity.paragraph, docx_edit::ParagraphRef::Session { story, .. } if story == &root)
    }).collect();
    assert_eq!(band_identities.len(), 1);
    let plan = engine.doc().paragraph_save_plan();
    let band_parts: Vec<_> = plan
        .spliced_parts
        .iter()
        .filter(|part| {
            part.part.contains(if kind == StoryKind::Header {
                "header1.xml"
            } else {
                "footer1.xml"
            })
        })
        .collect();
    assert_eq!(band_parts.len(), 1);
    assert_eq!(band_parts[0].paragraphs.len(), 1);
}

#[test]
fn header_aliases_share_edits_and_keep_section_relationships() {
    assert_shared_content(
        &[("rId7", "header1.xml"), ("rId9", "header1.xml")],
        &[],
        StoryKind::Header,
    );
}

#[test]
fn three_headers_share_a_part_with_a_dot_target() {
    assert_shared_content(
        &[
            ("rId7", "header1.xml"),
            ("rId9", "./header1.xml"),
            ("rId11", "header1.xml"),
        ],
        &[],
        StoryKind::Header,
    );
}

#[test]
fn footer_aliases_share_edits() {
    assert_shared_content(
        &[],
        &[("rId7", "footer1.xml"), ("rId9", "footer1.xml")],
        StoryKind::Footer,
    );
}

#[test]
fn mapping_survives_hydration_without_source_bytes() {
    let doc = seeded(&alias_package());
    assert_eq!(story_ids(&doc), ["body", "hf:rId7"]);
    let peer = EditingDoc::new(42);
    peer.apply_update_v1(&doc.encode_state_as_update_v1())
        .unwrap();
    assert_eq!(peer.header_footer_story("rId9"), "hf:rId7");
    assert_eq!(peer.header_footer_story("rId7"), "hf:rId7");
    assert_eq!(peer.header_footer_story("rIdOther"), "hf:rIdOther");
    assert_eq!(
        peer.header_footer_aliases(),
        [("rId9".to_owned(), "rId7".to_owned())]
    );
    assert!(peer.paragraph_identities().package_sha256.is_none());
}

#[test]
fn legacy_stories_deactivate_the_whole_group_after_later_updates() {
    for alias in ["hf:rId9", "hf:rId9:t0:r0c0", "hf:rId9:sdt0"] {
        let doc = seeded(&fixture::package(
            &[
                ("rId7", "header1.xml"),
                ("rId9", "header1.xml"),
                ("rId11", "header1.xml"),
            ],
            &[],
        ));
        let peer = EditingDoc::new(42);
        peer.apply_update_v1(&doc.encode_state_as_update_v1())
            .unwrap();
        assert_eq!(peer.header_footer_story("rId9"), "hf:rId7");
        doc.create_story(alias, "Legacy", "Normal", "left").unwrap();
        assert_eq!(doc.header_footer_story("rId9"), "hf:rId9");
        assert_eq!(doc.header_footer_story("rId11"), "hf:rId11");
        assert!(doc.header_footer_aliases().is_empty());
        peer.apply_update_v1(&doc.encode_state_as_update_v1())
            .unwrap();
        assert_eq!(peer.header_footer_story("rId9"), "hf:rId9");
        assert!(peer.header_footer_aliases().is_empty());
        let hydrated = EditingDoc::new(43);
        hydrated
            .apply_update_v1(&doc.encode_state_as_update_v1())
            .unwrap();
        assert_eq!(hydrated.header_footer_story("rId9"), "hf:rId9");
        doc.delete_story(alias).unwrap();
        assert_eq!(doc.header_footer_story("rId9"), "hf:rId7");
        peer.apply_update_v1(&doc.encode_state_as_update_v1())
            .unwrap();
        assert_eq!(peer.header_footer_story("rId9"), "hf:rId7");
    }
}

#[test]
fn similar_story_prefix_does_not_deactivate_aliases() {
    let doc = seeded(&alias_package());
    doc.create_story("hf:rId90", "Other", "Normal", "left")
        .unwrap();
    assert_eq!(doc.header_footer_story("rId9"), "hf:rId7");
}

#[test]
fn header_and_footer_groups_activate_independently() {
    let doc = seeded(&fixture::package(
        &[("rId7", "header1.xml"), ("rId9", "header1.xml")],
        &[("rId11", "footer1.xml"), ("rId13", "footer1.xml")],
    ));
    assert_eq!(story_ids(&doc), ["body", "hf:rId11", "hf:rId7"]);
    assert_eq!(doc.header_footer_aliases().len(), 2);
    doc.create_story("hf:rId9", "Legacy", "Normal", "left")
        .unwrap();
    assert_eq!(doc.header_footer_story("rId9"), "hf:rId9");
    assert_eq!(doc.header_footer_story("rId13"), "hf:rId11");
    assert_eq!(
        doc.header_footer_aliases(),
        [("rId13".to_owned(), "rId11".to_owned())]
    );
}

#[test]
fn nested_header_content_is_seeded_and_indexed_only_for_the_canonical_story() {
    let mut parts = ooxml_opc::unzip_parts(&alias_package()).unwrap();
    let (_, header) = parts
        .iter_mut()
        .find(|(path, _)| path == "word/header1.xml")
        .unwrap();
    *header = String::from_utf8(header.clone()).unwrap().replace(
        "</w:hdr>",
        r#"<w:tbl><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid><w:tr><w:tc><w:p w14:paraId="20000002"><w:r><w:t>Cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:sdt><w:sdtPr><w:id w:val="17"/><w:text/></w:sdtPr><w:sdtContent><w:p w14:paraId="20000003"><w:r><w:t>Control</w:t></w:r></w:p></w:sdtContent></w:sdt></w:hdr>"#,
    ).into_bytes();
    let doc = seeded(&ooxml_opc::rezip_parts(&parts).unwrap());
    let ids = story_ids(&doc);
    assert!(ids.iter().any(|id| id.starts_with("hf:rId7:t")));
    assert!(ids.iter().any(|id| id.starts_with("hf:rId7:sdt")));
    assert!(
        ids.iter()
            .all(|id| id == "body" || id == "hf:rId7" || id.starts_with("hf:rId7:"))
    );
    let identities = doc.paragraph_identities();
    let header_sources: Vec<_> = identities
        .paragraphs
        .iter()
        .filter_map(|identity| identity.source.as_ref())
        .filter(|source| source.part_uri == "/word/header1.xml")
        .collect();
    assert_eq!(header_sources.len(), 3);
    let plan = doc.paragraph_save_plan();
    let header = plan
        .spliced_parts
        .iter()
        .find(|part| part.part == "word/header1.xml")
        .unwrap();
    assert_eq!(header.paragraphs.len(), 3);
}

#[test]
fn incoming_alias_story_ids_are_not_routed_to_edit_operations() {
    let doc = seeded(&alias_package());
    assert!(
        doc.insert_text(
            &EditCtx::local("Ada", "2030-01-02T03:04:05Z"),
            Position::new("hf:rId9", 0),
            "Alias edit",
            FormatPolicy::Plain,
        )
        .is_err()
    );
    assert_eq!(doc.paragraphs("hf:rId7").unwrap()[0].text, "Shared band");
}
