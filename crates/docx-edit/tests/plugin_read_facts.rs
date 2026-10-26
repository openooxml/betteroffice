use std::collections::{BTreeMap, HashSet};

use docx_edit::structured::{ExportOptions, RevisionView, StorySelection};
use docx_edit::*;
use serde_json::{Value, json};

const NS: &str = r#"xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml""#;
const R: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

fn p(id: usize, properties: &str) -> String {
    format!(
        r#"<w:p w14:paraId="{id:08X}"><w:pPr>{properties}</w:pPr><w:r><w:t>item{id}</w:t></w:r></w:p>"#
    )
}

fn num(id: usize, level: usize) -> String {
    format!(r#"<w:numPr><w:ilvl w:val="{level}"/><w:numId w:val="{id}"/></w:numPr>"#)
}

fn numbering(restart: &str) -> String {
    let levels: String = (0..4).map(|level| {
        let template = (1..=level + 1).map(|n| format!("%{n}")).collect::<Vec<_>>().join(".");
        format!(r#"<w:lvl w:ilvl="{level}"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:pStyle w:val="Heading{}"/><w:lvlText w:val="{template}."/>{}</w:lvl>"#, level + 1, if level == 1 { restart } else { "" })
    }).collect();
    format!(
        r#"<w:abstractNum w:abstractNumId="0"><w:multiLevelType w:val="multilevel"/>{levels}</w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num><w:num w:numId="2"><w:abstractNumId w:val="0"/><w:lvlOverride w:ilvl="0"><w:startOverride w:val="1"/></w:lvlOverride></w:num><w:num w:numId="3"><w:abstractNumId w:val="0"/></w:num><w:abstractNum w:abstractNumId="1"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="•"/></w:lvl></w:abstractNum><w:num w:numId="4"><w:abstractNumId w:val="1"/></w:num>"#
    )
}

fn package(body: &str, styles: &str, numbering: &str) -> Vec<u8> {
    let mut parts = vec![
        ("[Content_Types].xml".to_owned(), br#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>"#.to_vec()),
        ("_rels/.rels".to_owned(), format!(r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="main" Type="{R}/officeDocument" Target="word/document.xml"/></Relationships>"#).into_bytes()),
        ("word/document.xml".to_owned(), format!(r#"<w:document {NS}><w:body>{body}</w:body></w:document>"#).into_bytes()),
        ("word/styles.xml".to_owned(), format!(r#"<w:styles {NS}><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>{styles}</w:styles>"#).into_bytes()),
        ("word/numbering.xml".to_owned(), format!(r#"<w:numbering {NS}>{numbering}</w:numbering>"#).into_bytes()),
        ("word/settings.xml".to_owned(), format!(r#"<w:settings {NS}/>"#).into_bytes()),
        ("word/footnotes.xml".to_owned(), format!(r#"<w:footnotes {NS}><w:footnote w:id="1">{}</w:footnote></w:footnotes>"#, p(90, "")).into_bytes()),
    ];
    let mut overrides = String::new();
    for (name, kind) in [
        ("styles", "styles"),
        ("numbering", "numbering"),
        ("settings", "settings"),
        ("footnotes", "footnotes"),
        ("headerA", "header"),
        ("headerB", "header"),
        ("headerC", "header"),
        ("headerE", "header"),
        ("footerF", "footer"),
    ] {
        overrides.push_str(&format!(r#"<Override PartName="/word/{name}.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.{kind}+xml"/>"#));
    }
    let content_types = &mut parts
        .iter_mut()
        .find(|(name, _)| name == "[Content_Types].xml")
        .unwrap()
        .1;
    *content_types = String::from_utf8(content_types.clone())
        .unwrap()
        .replace("</Types>", &format!("{overrides}</Types>"))
        .into_bytes();
    let mut rels = String::new();
    for (id, kind, target) in [
        ("styles", "styles", "styles.xml"),
        ("numbering", "numbering", "numbering.xml"),
        ("settings", "settings", "settings.xml"),
        ("notes", "footnotes", "footnotes.xml"),
        ("A", "header", "headerA.xml"),
        ("B", "header", "headerB.xml"),
        ("C", "header", "headerC.xml"),
        ("E", "header", "headerE.xml"),
        ("F", "footer", "footerF.xml"),
    ] {
        rels.push_str(&format!(
            r#"<Relationship Id="{id}" Type="{R}/{kind}" Target="{target}"/>"#
        ));
        if kind == "header" || kind == "footer" {
            let tag = if kind == "header" { "hdr" } else { "ftr" };
            parts.push((
                format!("word/{target}"),
                format!(
                    "<w:{tag} {NS}>{}</w:{tag}>",
                    p(100 + id.as_bytes()[0] as usize, "")
                )
                .into_bytes(),
            ));
        }
    }
    parts.push(("word/_rels/document.xml.rels".to_owned(), format!(r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">{rels}</Relationships>"#).into_bytes()));
    ooxml_opc::rezip_parts(&parts).unwrap()
}

fn open(bytes: &[u8]) -> EditingDoc {
    let doc = EditingDoc::new(7181);
    seed_from_docx(&doc, bytes).unwrap();
    doc
}

fn read(doc: &EditingDoc, story: &str, view: EditTextView) -> Vec<ParagraphText> {
    doc.read_paragraphs(&ReadParagraphsRequest {
        story: Some(story.into()),
        para_ids: None,
        view,
    })
    .unwrap()
    .paragraphs
}

fn export_facts(value: &Value, facts: &mut BTreeMap<(String, String), Value>) {
    match value {
        Value::Array(values) => {
            for value in values {
                export_facts(value, facts);
            }
        }
        Value::Object(fields) => {
            if matches!(
                value["kind"].as_str(),
                Some("paragraph" | "heading" | "listItem")
            ) && value["anchor"]["kind"] == "paragraph"
            {
                let list = &value["list"];
                facts.insert((value["anchor"]["story"].as_str().unwrap().into(), value["anchor"]["paraId"].as_str().unwrap().into()), json!({
                    "outlineLevel": value["heading"]["outlineLevel"],
                    "list": if list.is_null() { Value::Null } else { json!({"marker": list["marker"], "level": list["level"], "numId": list["numId"], "format": list["format"]}) }
                }));
            }
            for value in fields.values() {
                export_facts(value, facts);
            }
        }
        _ => {}
    }
}

fn consistent(doc: &EditingDoc, view: EditTextView) {
    consistent_with_literals(doc, view, &BTreeMap::new());
}

fn consistent_with_literals(
    doc: &EditingDoc,
    view: EditTextView,
    literals: &BTreeMap<(String, String), Value>,
) {
    let revision = match view {
        EditTextView::Accepted => RevisionView::Accepted,
        EditTextView::Original => RevisionView::Original,
    };
    let content = doc
        .export_structured(&ExportOptions {
            stories: Some(vec![
                StorySelection::Body,
                StorySelection::Headers,
                StorySelection::Footers,
                StorySelection::Footnotes,
            ]),
            ..ExportOptions::new(revision)
        })
        .unwrap()
        .content;
    let mut facts = BTreeMap::new();
    export_facts(&serde_json::to_value(content).unwrap(), &mut facts);
    let stories = doc
        .read_stories(&ReadStoriesRequest {
            stories: None,
            by_root: false,
            view,
            expect_version: None,
        })
        .unwrap();
    let story_reads: BTreeMap<_, _> = stories
        .stories
        .into_iter()
        .map(|story| match story {
            StoryText::Read { story, paragraphs } => (story, paragraphs),
            StoryText::Refused { failure, .. } => panic!("{failure:?}"),
        })
        .collect();
    for info in doc.list_stories().stories {
        let paragraphs = read(doc, &info.story, view);
        assert_eq!(paragraphs, story_reads[&info.story]);
        for paragraph in paragraphs {
            let key = (info.story.clone(), paragraph.para_id.clone());
            let expected = literals.get(&key).or_else(|| facts.get(&key));
            assert!(
                expected.is_some(),
                "missing export paragraph {} {}",
                info.story,
                paragraph.para_id
            );
            assert_eq!(
                json!({"outlineLevel": paragraph.outline_level, "list": paragraph.list}),
                *expected.unwrap(),
                "{} {}",
                info.story,
                paragraph.para_id
            );
        }
    }
}

fn markers(doc: &EditingDoc, expected: &[&str]) {
    consistent(doc, EditTextView::Accepted);
    let paragraphs = read(doc, "body", EditTextView::Accepted);
    assert_eq!(
        paragraphs
            .iter()
            .map(|p| p.list.as_ref().and_then(|l| l.marker.as_deref()))
            .collect::<Vec<_>>(),
        expected.iter().map(|s| Some(*s)).collect::<Vec<_>>()
    );
}

#[test]
fn outline_levels_resolve_styles_direct_values_and_body_text_like_export() {
    let styles = r#"<w:style w:type="paragraph" w:styleId="Heading1"><w:pPr><w:outlineLvl w:val="0"/></w:pPr></w:style><w:style w:type="paragraph" w:styleId="Heading2"><w:basedOn w:val="Heading1"/><w:pPr><w:outlineLvl w:val="1"/></w:pPr></w:style>"#;
    let body = [
        p(1, r#"<w:pStyle w:val="Heading1"/>"#),
        p(2, r#"<w:pStyle w:val="Heading2"/>"#),
        p(3, r#"<w:pStyle w:val="Normal"/><w:outlineLvl w:val="2"/>"#),
        p(4, r#"<w:pStyle w:val="Normal"/>"#),
        p(5, r#"<w:outlineLvl w:val="9"/>"#),
    ]
    .concat();
    let doc = open(&package(&body, styles, ""));
    assert_eq!(
        read(&doc, "body", EditTextView::Accepted)
            .iter()
            .map(|p| p.outline_level)
            .collect::<Vec<_>>(),
        [Some(0), Some(1), Some(2), None, None]
    );
    consistent(&doc, EditTextView::Accepted);
}

#[test]
fn direct_multilevel_and_bullet_numbering_reports_labels_levels_ids_and_formats() {
    let body = [(1, 0), (1, 0), (1, 1), (1, 0), (4, 0)]
        .iter()
        .enumerate()
        .map(|(i, (id, level))| p(i + 1, &num(*id, *level)))
        .collect::<String>();
    let doc = open(&package(&body, "", &numbering("")));
    markers(&doc, &["1.", "2.", "2.1.", "3.", "•"]);
    let paragraphs = read(&doc, "body", EditTextView::Accepted);
    for (index, paragraph) in paragraphs.iter().enumerate() {
        let list = paragraph.list.as_ref().unwrap();
        assert_eq!(list.level, if index == 2 { 1 } else { 0 });
        assert_eq!(list.num_id, if index == 4 { "4" } else { "1" });
        assert_eq!(list.format, if index == 4 { "bullet" } else { "decimal" });
    }
}

#[test]
fn style_only_heading_numbering_resolves_all_four_levels() {
    let styles = (0..4).map(|level| format!(r#"<w:style w:type="paragraph" w:styleId="Heading{}"><w:pPr><w:outlineLvl w:val="{level}"/>{}</w:pPr></w:style>"#, level+1, num(1,level))).collect::<String>();
    let levels = [0, 1, 2, 3, 1, 0, 1];
    let body = levels
        .iter()
        .enumerate()
        .map(|(i, level)| {
            p(
                i + 1,
                &format!(r#"<w:pStyle w:val="Heading{}"/>"#, level + 1),
            )
        })
        .collect::<String>();
    let numbering = numbering("")
        .replace("%1.\"", "%1\"")
        .replace("%1.%2.\"", "%1.%2\"")
        .replace("%1.%2.%3.\"", "%1.%2.%3\"")
        .replace("%1.%2.%3.%4.\"", "%1.%2.%3.%4\"");
    let doc = open(&package(&body, &styles, &numbering));
    markers(&doc, &["1", "1.1", "1.1.1", "1.1.1.1", "1.2", "2", "2.1"]);
    consistent(&doc, EditTextView::Original);
    assert_eq!(
        read(&doc, "body", EditTextView::Accepted),
        read(&doc, "body", EditTextView::Original)
    );
    assert_eq!(
        read(&doc, "body", EditTextView::Accepted)
            .iter()
            .map(|p| p.outline_level)
            .collect::<Vec<_>>(),
        levels.map(|v| Some(v as u8))
    );
}

#[test]
fn numbering_instances_restart_with_overrides_and_continue_without_them() {
    let body = [(1, 0), (1, 0), (2, 0), (2, 0), (3, 0)]
        .iter()
        .enumerate()
        .map(|(i, (id, level))| p(i + 1, &num(*id, *level)))
        .collect::<String>();
    markers(
        &open(&package(&body, "", &numbering(""))),
        &["1.", "2.", "1.", "2.", "3."],
    );
}

#[test]
fn lvl_restart_zero_preserves_child_counter_across_parent_items() {
    let body = [0, 1, 0, 1]
        .iter()
        .enumerate()
        .map(|(i, level)| p(i + 1, &num(1, *level)))
        .collect::<String>();
    markers(
        &open(&package(
            &body,
            "",
            &numbering(r#"<w:lvlRestart w:val="0"/>"#),
        )),
        &["1.", "1.1.", "2.", "2.2."],
    );
}

fn cell(content: &str) -> String {
    format!(
        r#"<w:tbl><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid><w:tr><w:tc>{content}</w:tc></w:tr></w:tbl>"#
    )
}

#[test]
fn cell_and_subset_reads_count_from_the_root_story() {
    let body = format!(
        "{}{}{}",
        p(1, &num(1, 0)),
        cell(&p(2, &num(1, 0))),
        p(3, &num(1, 0))
    );
    let doc = open(&package(&body, "", &numbering("")));
    let body = read(&doc, "body", EditTextView::Accepted);
    let nested = read(&doc, "body:t0:r0c0", EditTextView::Accepted);
    assert_eq!(
        body[0].list.as_ref().and_then(|l| l.marker.as_deref()),
        Some("1.")
    );
    assert_eq!(
        nested[0].list.as_ref().and_then(|l| l.marker.as_deref()),
        Some("2.")
    );
    assert_eq!(
        body[1].list.as_ref().and_then(|l| l.marker.as_deref()),
        Some("3.")
    );
    let subset = doc
        .read_paragraphs(&ReadParagraphsRequest {
            story: None,
            para_ids: Some(vec![body[1].para_id.clone()]),
            view: EditTextView::Accepted,
        })
        .unwrap();
    assert_eq!(subset.paragraphs, vec![body[1].clone()]);
    let subset = doc
        .read_paragraphs(&ReadParagraphsRequest {
            story: Some("body:t0:r0c0".into()),
            para_ids: Some(vec![nested[0].para_id.clone()]),
            view: EditTextView::Accepted,
        })
        .unwrap();
    assert_eq!(subset.paragraphs, nested);
    consistent(&doc, EditTextView::Accepted);
}

#[test]
fn property_revision_numbering_and_heading_facts_follow_each_read_view() {
    let change = |old: &str| {
        format!(
            r#"<w:pPrChange w:id="1" w:author="Author" w:date="2026-01-01T00:00:00Z"><w:pPr>{old}</w:pPr></w:pPrChange>"#
        )
    };
    let body = format!(
        "{}{}{}{}",
        p(1, &format!("{}{}", num(1, 0), change(""))),
        p(
            2,
            &format!(
                r#"<w:pStyle w:val="Normal"/>{}"#,
                change(r#"<w:pStyle w:val="Heading2"/>"#)
            )
        ),
        p(3, &change(&num(1, 0))),
        p(
            4,
            &format!(
                r#"<w:pStyle w:val="Heading1"/>{}"#,
                change(r#"<w:pStyle w:val="Normal"/>"#)
            )
        )
    );
    let styles = r#"<w:style w:type="paragraph" w:styleId="Heading1"><w:pPr><w:outlineLvl w:val="0"/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr></w:style><w:style w:type="paragraph" w:styleId="Heading2"><w:pPr><w:outlineLvl w:val="1"/></w:pPr></w:style>"#;
    let doc = open(&package(&body, styles, &numbering("")));
    let accepted = read(&doc, "body", EditTextView::Accepted);
    let original = read(&doc, "body", EditTextView::Original);
    assert!(accepted[0].list.is_some());
    assert!(original[0].list.is_none());
    assert_eq!(accepted[1].outline_level, None);
    assert_eq!(original[1].outline_level, Some(1));
    assert!(accepted[2].list.is_none());
    assert!(original[2].list.is_some());
    assert_eq!(
        accepted[0]
            .list
            .as_ref()
            .and_then(|list| list.marker.as_deref()),
        Some("1.")
    );
    assert_eq!(
        original[2]
            .list
            .as_ref()
            .and_then(|list| list.marker.as_deref()),
        Some("1.")
    );
    assert_eq!(accepted[3].outline_level, Some(0));
    assert_eq!(original[3].outline_level, None);
    assert_eq!(
        accepted[3]
            .list
            .as_ref()
            .and_then(|list| list.marker.as_deref()),
        Some("2.")
    );
    assert!(original[3].list.is_none());
    consistent(&doc, EditTextView::Accepted);
    consistent_with_literals(
        &doc,
        EditTextView::Original,
        &BTreeMap::from([
            (
                ("body".into(), original[0].para_id.clone()),
                json!({"outlineLevel": null, "list": null}),
            ),
            (
                ("body".into(), original[2].para_id.clone()),
                json!({"outlineLevel": null, "list": {"marker": "1.", "level": 0, "numId": "1", "format": "decimal"}}),
            ),
            (
                ("body".into(), original[3].para_id.clone()),
                json!({"outlineLevel": null, "list": null}),
            ),
        ]),
    );
}

#[test]
fn original_export_resolves_numbering_from_the_original_paragraph_style() {
    let change = r#"<w:pPrChange w:id="1" w:author="Author" w:date="2026-01-01T00:00:00Z"><w:pPr><w:pStyle w:val="Heading1"/></w:pPr></w:pPrChange>"#;
    let body = format!(
        "{}{}",
        p(1, &format!(r#"<w:pStyle w:val="Normal"/>{change}"#)),
        p(2, r#"<w:pStyle w:val="Heading1"/>"#)
    );
    let styles = r#"<w:style w:type="paragraph" w:styleId="Heading1"><w:pPr><w:outlineLvl w:val="0"/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr></w:style>"#;
    let doc = open(&package(&body, styles, &numbering("")));
    let original = read(&doc, "body", EditTextView::Original);
    assert_eq!(
        original
            .iter()
            .map(|paragraph| paragraph.outline_level)
            .collect::<Vec<_>>(),
        [Some(0), Some(0)]
    );
    assert_eq!(
        original
            .iter()
            .map(|paragraph| paragraph
                .list
                .as_ref()
                .and_then(|list| list.marker.as_deref()))
            .collect::<Vec<_>>(),
        [Some("1."), Some("2.")]
    );
    let accepted = read(&doc, "body", EditTextView::Accepted);
    assert!(accepted[0].list.is_none());
    assert_eq!(
        accepted[1]
            .list
            .as_ref()
            .and_then(|list| list.marker.as_deref()),
        Some("1.")
    );
    consistent(&doc, EditTextView::Accepted);
    consistent(&doc, EditTextView::Original);
}

#[test]
fn original_numbering_changes_advance_later_items_before_subset_filtering() {
    let change = format!(
        r#"<w:pPrChange w:id="1" w:author="Author" w:date="2026-01-01T00:00:00Z"><w:pPr>{}</w:pPr></w:pPrChange>"#,
        num(1, 0)
    );
    let body = format!("{}{}", p(1, &change), p(2, &num(1, 0)));
    let doc = open(&package(&body, "", &numbering("")));
    let original = read(&doc, "body", EditTextView::Original);
    consistent_with_literals(
        &doc,
        EditTextView::Original,
        &BTreeMap::from([
            (
                ("body".into(), original[0].para_id.clone()),
                json!({"outlineLevel": null, "list": {"marker": "1.", "level": 0, "numId": "1", "format": "decimal"}}),
            ),
            (
                ("body".into(), original[1].para_id.clone()),
                json!({"outlineLevel": null, "list": {"marker": "2.", "level": 0, "numId": "1", "format": "decimal"}}),
            ),
        ]),
    );
    let subset = doc
        .read_paragraphs(&ReadParagraphsRequest {
            story: None,
            para_ids: Some(vec![original[1].para_id.clone()]),
            view: EditTextView::Original,
        })
        .unwrap();
    assert_eq!(subset.paragraphs.as_slice(), &original[1..]);
    let accepted = read(&doc, "body", EditTextView::Accepted);
    assert!(accepted[0].list.is_none());
    assert_eq!(
        accepted[1]
            .list
            .as_ref()
            .and_then(|list| list.marker.as_deref()),
        Some("1.")
    );
    consistent(&doc, EditTextView::Accepted);
}

#[test]
fn paragraph_fact_cache_refreshes_both_views_after_a_document_edit() {
    let body = format!("{}{}", p(1, &num(1, 0)), p(2, &num(1, 0)));
    let doc = open(&package(&body, "", &numbering("")));
    let first = read(&doc, "body", EditTextView::Accepted)[0]
        .para_id
        .clone();
    for view in [EditTextView::Accepted, EditTextView::Original] {
        assert_eq!(
            read(&doc, "body", view)[1]
                .list
                .as_ref()
                .and_then(|list| list.marker.as_deref()),
            Some("2.")
        );
    }
    doc.set_paragraph_attr(
        &first,
        "numPr",
        yrs::Any::Map(std::sync::Arc::new(std::collections::HashMap::from([(
            "numId".to_owned(),
            yrs::Any::Number(0.0),
        )]))),
    )
    .unwrap();
    doc.set_paragraph_attr(&first, "outlineLevel", yrs::Any::Number(3.0))
        .unwrap();
    for view in [EditTextView::Accepted, EditTextView::Original] {
        let paragraphs = read(&doc, "body", view);
        assert!(paragraphs[0].list.is_none());
        assert_eq!(paragraphs[0].outline_level, Some(3));
        assert_eq!(
            paragraphs[1]
                .list
                .as_ref()
                .and_then(|list| list.marker.as_deref()),
            Some("1.")
        );
        consistent(&doc, view);
    }
}

fn all_reads(doc: &EditingDoc) -> Vec<ParagraphText> {
    doc.read_stories(&ReadStoriesRequest {
        stories: None,
        by_root: false,
        view: EditTextView::Accepted,
        expect_version: None,
    })
    .unwrap()
    .stories
    .into_iter()
    .flat_map(|story| match story {
        StoryText::Read { paragraphs, .. } => paragraphs,
        StoryText::Refused { failure, .. } => panic!("{failure:?}"),
    })
    .collect()
}

#[test]
fn read_ids_match_export_search_and_edit_targets_in_every_story_with_duplicate_source_ids() {
    let body = format!(
        r#"{}{}{}<w:sectPr><w:headerReference w:type="default" r:id="A"/></w:sectPr>"#,
        p(1, ""),
        cell(&p(1, "")),
        p(1, "")
    );
    let doc = open(&package(&body, "", ""));
    let paragraphs = all_reads(&doc);
    assert!(paragraphs.iter().any(|p| p.story == "body:t0:r0c0"));
    assert!(paragraphs.iter().any(|p| p.story == "hf:A"));
    assert!(paragraphs.iter().any(|p| p.story == "fn:1"));
    let mut ids = HashSet::new();
    for paragraph in &paragraphs {
        assert!(
            ids.insert(paragraph.para_id.clone()),
            "duplicate read id {}",
            paragraph.para_id
        );
        let single = read(&doc, &paragraph.story, EditTextView::Accepted);
        assert!(single.contains(paragraph));
    }
    consistent(&doc, EditTextView::Accepted);
    for paragraph in paragraphs {
        let hits = doc
            .find_text(&FindTextRequest {
                text: paragraph.text.clone(),
                within: SearchScope::Paragraph(ParagraphTarget {
                    story: paragraph.story.clone(),
                    para_id: paragraph.para_id.clone(),
                }),
                view: EditTextView::Accepted,
                limit: None,
            })
            .unwrap();
        assert_eq!(hits.matches.len(), 1);
        assert_eq!(hits.matches[0].range.start.para_id, paragraph.para_id);
        assert_eq!(hits.matches[0].range.end.para_id, paragraph.para_id);
        assert_eq!(hits.matches[0].range.story, paragraph.story);
        let request = EditRequest {
            expect_version: doc.version(),
            source: EditSource::Host,
            history: EditHistory::Separate,
            steps: vec![EditStep::new(EditOperation::ReplaceText {
                target: TextTarget::Paragraph(ParagraphTarget {
                    story: paragraph.story.clone(),
                    para_id: paragraph.para_id.clone(),
                }),
                text: format!("{} edited", paragraph.text),
            })],
        };
        doc.apply_edits(&request, &UndoSession::new())
            .unwrap()
            .unwrap();
        assert!(
            read(&doc, &paragraph.story, EditTextView::Accepted)
                .iter()
                .any(|p| p.para_id == paragraph.para_id && p.text.ends_with(" edited"))
        );
    }
}

fn layout_request(engine: &EngineSession, bytes: &[u8]) -> String {
    let package = docx_parse::parse_docx_s9_wire(bytes, Default::default())
        .unwrap()
        .document
        .package;
    let sections = package
        .document
        .sections
        .unwrap()
        .into_iter()
        .map(|section| json!({"properties":section.properties}))
        .collect::<Vec<_>>();
    let font = docx_layout::register_measure_font(include_bytes!(
        "../../ooxml-text/tests/fonts/LiberationSans-Regular.ttf"
    ))
    .unwrap();
    let mut request = json!({"bodyStory":"body","renderEnv":{},"regions":{"sections":sections,"settings":package.settings},"measurement":{"fontChains":{},"defaults":{"fontFamily":"Liberation Sans","fontSize":11},"authoritativeShaping":true}});
    let requirements: Vec<Value> = serde_json::from_str(
        &engine
            .layout_font_requirements_json(&request.to_string())
            .unwrap(),
    )
    .unwrap();
    for requirement in requirements {
        request["measurement"]["fontChains"][requirement["key"].as_str().unwrap()] = json!([font]);
    }
    request.to_string()
}

fn uses_json(engine: &EngineSession) -> BTreeMap<(String, u64, String), Value> {
    let response = serde_json::to_value(engine.list_stories_with_layout()).unwrap();
    let mut uses = BTreeMap::new();
    for story in response["stories"].as_array().unwrap() {
        if let Some(entries) = story["uses"].as_array() {
            for usage in entries {
                uses.insert(
                    (
                        story["story"].as_str().unwrap().into(),
                        usage["sectionIndex"].as_u64().unwrap(),
                        usage["variant"].as_str().unwrap().into(),
                    ),
                    usage.clone(),
                );
            }
        }
    }
    assert!(!uses.is_empty());
    uses
}

#[test]
fn shown_uses_follow_painted_section_bands_and_are_absent_without_current_complete_layout() {
    let geometry = r#"<w:type w:val="nextPage"/><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720"/>"#;
    let first = format!(
        r#"<w:sectPr><w:headerReference w:type="first" r:id="A"/><w:headerReference w:type="default" r:id="B"/><w:headerReference w:type="even" r:id="E"/><w:footerReference w:type="default" r:id="F"/><w:footerReference w:type="first" r:id="F"/><w:titlePg/>{geometry}</w:sectPr>"#
    );
    let body = format!(
        r#"{}{}<w:sectPr><w:headerReference w:type="default" r:id="C"/>{geometry}</w:sectPr>"#,
        p(1, &first),
        p(2, "")
    );
    let bytes = package(&body, "", "");
    let engine = EngineSession::new(7182);
    seed_from_docx(engine.doc(), &bytes).unwrap();
    assert!(
        uses_json(&engine)
            .values()
            .all(|usage| usage.get("shown").is_none())
    );
    let request = layout_request(&engine, &bytes);
    let layout: Value =
        serde_json::from_str(&engine.layout_document_with_regions_json(&request).unwrap()).unwrap();
    assert_eq!(layout["layout"]["pages"].as_array().unwrap().len(), 2);
    let uses = uses_json(&engine);
    assert!(uses.values().all(|usage| usage["shown"].is_boolean()));
    for (story, section, variant, shown) in [
        ("hf:A", 0, "first", true),
        ("hf:B", 0, "default", false),
        ("hf:C", 1, "default", true),
        ("hf:F", 0, "first", true),
        ("hf:F", 0, "default", true),
        ("hf:F", 1, "default", true),
        ("hf:E", 0, "even", false),
        ("hf:E", 1, "even", false),
    ] {
        assert_eq!(
            uses[&(story.into(), section, variant.into())]["shown"],
            shown,
            "{story} {section} {variant}"
        );
    }
    let export = serde_json::to_value(
        engine
            .doc()
            .export_structured(&ExportOptions {
                stories: Some(vec![StorySelection::Headers, StorySelection::Footers]),
                ..ExportOptions::new(RevisionView::Accepted)
            })
            .unwrap()
            .content,
    )
    .unwrap();
    for story in export["stories"].as_array().unwrap() {
        for usage in story["uses"].as_array().unwrap() {
            assert!(usage.get("shown").is_none());
        }
    }
    let before = engine.doc().version();
    engine
        .doc()
        .insert_text(
            &EditCtx::local("Author", "2026-01-01T00:00:00Z"),
            Position::new("body", 0),
            "Edited ",
            FormatPolicy::Plain,
        )
        .unwrap();
    assert_ne!(before, engine.doc().version());
    assert!(
        uses_json(&engine)
            .values()
            .all(|usage| usage.get("shown").is_none())
    );
    engine.layout_document_with_regions_json(&request).unwrap();
    assert_eq!(uses_json(&engine), uses);
}

#[test]
fn paragraph_facts_serialize_null_and_deserialize_legacy_payloads() {
    let paragraph: ParagraphText =
        serde_json::from_value(json!({"story":"body","paraId":"p","text":"text","atoms":[]}))
            .unwrap();
    let value = serde_json::to_value(paragraph).unwrap();
    assert_eq!(value.get("outlineLevel"), Some(&Value::Null));
    assert_eq!(value.get("list"), Some(&Value::Null));
}

#[test]
fn read_paragraphs_and_read_stories_match_export_facts_across_nested_stories() {
    let body = format!(
        r#"{}{}<w:sdt><w:sdtPr><w:id w:val="10"/></w:sdtPr><w:sdtContent>{}</w:sdtContent></w:sdt>{}"#,
        p(1, &num(1, 0)),
        cell(&p(2, &num(1, 0))),
        p(3, &num(1, 0)),
        p(4, &num(1, 0))
    );
    let doc = open(&package(&body, "", &numbering("")));
    consistent(&doc, EditTextView::Accepted);
    consistent(&doc, EditTextView::Original);
    for (story, text, marker) in [
        ("body", "item1", "1."),
        ("body:t0:r0c0", "item2", "2."),
        ("body:sdt0", "item3", "3."),
        ("body", "item4", "4."),
    ] {
        let paragraphs = read(&doc, story, EditTextView::Accepted);
        let paragraph = paragraphs.iter().find(|p| p.text == text).unwrap();
        assert_eq!(
            paragraph.list.as_ref().and_then(|l| l.marker.as_deref()),
            Some(marker)
        );
    }
}

#[test]
fn header_footer_and_note_roots_start_their_own_list_counters() {
    let body = format!(
        r#"{}{}<w:sectPr><w:headerReference w:type="default" r:id="A"/><w:footerReference w:type="default" r:id="F"/></w:sectPr>"#,
        p(1, &num(1, 0)),
        p(2, &num(1, 0))
    );
    let mut parts = ooxml_opc::unzip_parts(&package(&body, "", &numbering(""))).unwrap();
    for (name, bytes) in &mut parts {
        if matches!(
            name.as_str(),
            "word/headerA.xml" | "word/footerF.xml" | "word/footnotes.xml"
        ) {
            *bytes = String::from_utf8(bytes.clone())
                .unwrap()
                .replace("<w:pPr></w:pPr>", &format!("<w:pPr>{}</w:pPr>", num(1, 0)))
                .into_bytes();
        }
    }
    let doc = open(&ooxml_opc::rezip_parts(&parts).unwrap());
    consistent(&doc, EditTextView::Accepted);
    for story in ["body", "hf:A", "hf:F", "fn:1"] {
        let paragraphs = read(&doc, story, EditTextView::Accepted);
        assert_eq!(
            paragraphs[0]
                .list
                .as_ref()
                .and_then(|list| list.marker.as_deref()),
            Some("1."),
            "{story}"
        );
    }
    assert_eq!(
        read(&doc, "body", EditTextView::Accepted)[1]
            .list
            .as_ref()
            .and_then(|list| list.marker.as_deref()),
        Some("2.")
    );
}
