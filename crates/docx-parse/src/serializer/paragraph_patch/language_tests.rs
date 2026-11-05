use serde_json::{Value, json};

use super::language::patch_style_languages;
use super::properties::{paragraph_properties, run_properties};
use super::{Failure, PatchContext, patch_inner};
use crate::paragraph_patch::{
    AttrDeltaV1, Attrs, InheritanceV1, ParagraphPatchV1, RunEditV1, SeedV1, SlotV1, UnitV1,
    result_sha256, unit_sha256,
};

fn delta(slot: &str, value: Value) -> AttrDeltaV1 {
    AttrDeltaV1 {
        path: vec!["language".to_owned(), slot.to_owned()],
        before: SlotV1::Missing,
        after: SlotV1::Present { value },
    }
}

fn ok(result: Result<String, Failure>) -> String {
    result.unwrap_or_else(|error| panic!("{}", error.detail))
}

#[test]
fn run_language_preserves_other_properties_and_revision_bytes() {
    let source = r#"<w:rPr><w:rFonts w:ascii='Arial'/><w:b/><w:sz w:val="24"/><w:lang w:val='en-US' w:eastAsia="ja-JP" w:bidi="ar-SA"/><w:rPrChange w:id="1" w:author="A"><w:rPr><w:i/></w:rPr></w:rPrChange></w:rPr>"#;
    let after = ok(run_properties(
        source,
        &[delta("latin", json!("de-DE"))],
        &Attrs::new(),
        &Attrs::new(),
    ));
    assert_eq!(after, source.replace("'en-US'", "'de-DE'"));
    let after = ok(run_properties(
        source,
        &[delta("eastAsia", Value::Null)],
        &Attrs::new(),
        &Attrs::new(),
    ));
    assert_eq!(after, source.replace(" w:eastAsia=\"ja-JP\"", ""));
    let after = ok(run_properties(
        source,
        &[
            delta("latin", Value::Null),
            delta("eastAsia", Value::Null),
            delta("bidi", Value::Null),
        ],
        &Attrs::new(),
        &Attrs::new(),
    ));
    assert_eq!(
        after,
        source.replace(
            r#"<w:lang w:val='en-US' w:eastAsia="ja-JP" w:bidi="ar-SA"/>"#,
            ""
        )
    );
}

#[test]
fn missing_language_is_inserted_at_schema_rank() {
    let source = r#"<w:rPr><w:b/><w:sz w:val="24"/><w:eastAsianLayout/><w:rPrChange w:id="1"><w:rPr/></w:rPrChange></w:rPr>"#;
    let after = ok(run_properties(
        source,
        &[delta("latin", json!("de-DE"))],
        &Attrs::new(),
        &Attrs::new(),
    ));
    assert_eq!(
        after,
        source.replace(
            "<w:eastAsianLayout/>",
            r#"<w:lang w:val="de-DE"/><w:eastAsianLayout/>"#
        )
    );
    assert_eq!(
        ok(run_properties(
            "",
            &[delta("latin", json!("de-DE"))],
            &Attrs::new(),
            &Attrs::new()
        )),
        r#"<w:rPr><w:lang w:val="de-DE"/></w:rPr>"#
    );
}

#[test]
fn paragraph_mark_language_preserves_other_paragraph_properties() {
    let source = r#"<w:pPr><w:jc w:val="center"/><w:rPr><w:b/><w:lang w:val="en-US" w:bidi="ar-SA"/></w:rPr><w:sectPr/></w:pPr>"#;
    let after = ok(paragraph_properties(
        source,
        &[delta("latin", json!("de-DE"))],
        &Attrs::new(),
        &Attrs::new(),
        &Attrs::new(),
    ));
    assert_eq!(after, source.replace("en-US", "de-DE"));
    let source = "<w:pPr><w:jc w:val=\"center\"/><w:sectPr/></w:pPr>";
    let after = ok(paragraph_properties(
        source,
        &[delta("latin", json!("de-DE"))],
        &Attrs::new(),
        &Attrs::new(),
        &Attrs::new(),
    ));
    assert_eq!(
        after,
        source.replace(
            "<w:sectPr/>",
            r#"<w:rPr><w:lang w:val="de-DE"/></w:rPr><w:sectPr/>"#
        )
    );
    assert_eq!(
        ok(paragraph_properties(
            "",
            &[delta("latin", Value::Null)],
            &Attrs::new(),
            &Attrs::new(),
            &Attrs::new()
        )),
        ""
    );
}

#[test]
fn language_only_paragraph_with_run_revision_stays_patchable() {
    let source = r#"<w:p><w:r><w:rPr><w:lang w:val="en-US"/><w:rPrChange w:author="A" w:id="7"><w:rPr><w:b/></w:rPr></w:rPrChange></w:rPr><w:t>Text</w:t></w:r></w:p>"#;
    let units = vec![UnitV1::Text {
        text: "Text".to_owned(),
        attrs: Attrs::new(),
    }];
    let target = Attrs::from([("language".to_owned(), json!({"latin": "de-DE"}))]);
    let result = vec![UnitV1::Text {
        text: "Text".to_owned(),
        attrs: target.clone(),
    }];
    let patch = ParagraphPatchV1 {
        ordinal: 0,
        session_key: "source:0".to_owned(),
        seed: SeedV1 {
            unit_sha256: unit_sha256(&units),
            units,
            properties: Attrs::new(),
            inherit: InheritanceV1::default(),
        },
        text_edits: vec![],
        run_edits: vec![RunEditV1 {
            from: 0,
            to: 4,
            deltas: vec![delta("latin", json!("de-DE"))],
            target,
            inherit: Attrs::new(),
        }],
        para_deltas: vec![],
        para_inherit: Attrs::new(),
        mark_inherit: Attrs::new(),
        result_sha256: result_sha256(&result, &Attrs::new()),
    };
    assert_eq!(
        ok(patch_inner(source, &patch, &PatchContext::word_document())),
        source.replace("en-US", "de-DE")
    );
}

const NS: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

fn patch_styles(body: &str, entries: Value) -> String {
    patch_style_languages(
        &format!(r#"<w:styles xmlns:w="{NS}">{body}</w:styles>"#),
        &entries.to_string(),
    )
    .unwrap()
}

#[test]
fn style_language_patches_only_the_direct_run_properties() {
    let body = r#"<w:style w:styleId="Heading1"><w:name w:val="Heading"/><w:pPr><w:rPr><w:lang w:val="en-US"/></w:rPr></w:pPr><w:rPr><w:rFonts w:ascii='Arial'/><w:lang w:val="en-US" w:eastAsia="ja-JP" w:bidi="ar-SA"/><w:rPrChange w:id="4"><w:rPr/></w:rPrChange></w:rPr></w:style>"#;
    let after = patch_styles(body, json!({"Heading1": {"lang": "de-DE", "bidi": null}}));
    let expected = body.replace(
        r#"<w:lang w:val="en-US" w:eastAsia="ja-JP" w:bidi="ar-SA"/>"#,
        r#"<w:lang w:val="de-DE" w:eastAsia="ja-JP"/>"#,
    );
    assert_eq!(
        after,
        format!(r#"<w:styles xmlns:w="{NS}">{expected}</w:styles>"#)
    );
}

#[test]
fn style_language_creates_missing_properties_and_expands_empty_styles() {
    for body in [
        r#"<w:style w:styleId="A"><w:name w:val="A"/><w:tblPr/></w:style>"#,
        r#"<w:style w:styleId="A"/>"#,
        r#"<w:style w:styleId="A"><w:rPr/></w:style>"#,
        r#"<w:style w:styleId="A"><w:rPr><w:b/></w:rPr></w:style>"#,
    ] {
        let after = patch_styles(body, json!({"A": {"lang": "de-DE"}}));
        assert!(after.contains(r#"<w:lang w:val="de-DE"/>"#));
        if body.contains("<w:tblPr/>") {
            assert!(after.contains(r#"<w:rPr><w:lang w:val="de-DE"/></w:rPr><w:tblPr/>"#));
        }
        if body.contains("<w:b/>") {
            assert!(after.contains(r#"<w:b/><w:lang w:val="de-DE"/>"#));
        }
    }
}

#[test]
fn style_language_clears_all_slots_without_touching_other_bytes() {
    let body = r#"<w:style w:styleId="A"><w:rPr><w:b/><w:lang w:val="en-US" w:eastAsia="ja-JP" w:bidi="ar-SA"/><w:sz w:val="24"/></w:rPr></w:style>"#;
    let expected = body.replace(
        r#"<w:lang w:val="en-US" w:eastAsia="ja-JP" w:bidi="ar-SA"/>"#,
        "",
    );
    assert_eq!(
        patch_styles(
            body,
            json!({"A": {"lang": null, "eastAsia": null, "bidi": null}})
        ),
        format!(r#"<w:styles xmlns:w="{NS}">{expected}</w:styles>"#)
    );
    assert_eq!(
        patch_styles(r#"<w:style w:styleId="A"/>"#, json!({"A": {"lang": null}})),
        format!(r#"<w:styles xmlns:w="{NS}"><w:style w:styleId="A"/></w:styles>"#)
    );
}

#[test]
fn styles_with_another_prefix_keep_that_prefix() {
    let source = format!(
        r#"<x:styles xmlns:x="{NS}"><x:style x:styleId="A"><x:rPr><x:b/><x:lang x:val='en-US' x:eastAsia="ja-JP"/></x:rPr></x:style><x:style x:styleId="B"/></x:styles>"#
    );
    let after =
        patch_style_languages(&source, r#"{"A":{"lang":"de-DE"},"B":{"bidi":"he-IL"}}"#).unwrap();
    assert_eq!(
        after,
        source.replace("en-US", "de-DE").replace(
            r#"<x:style x:styleId="B"/>"#,
            r#"<x:style x:styleId="B"><x:rPr><x:lang x:bidi="he-IL"/></x:rPr></x:style>"#
        )
    );
}

#[test]
fn defaults_create_missing_containers_in_schema_order() {
    for body in [
        r#"<w:docDefaults><w:rPrDefault><w:rPr><w:b/></w:rPr></w:rPrDefault></w:docDefaults>"#,
        r#"<w:docDefaults><w:pPrDefault><w:pPr/></w:pPrDefault></w:docDefaults>"#,
        r#"<w:latentStyles/><w:style w:styleId="A"/>"#,
        r#"<w:docDefaults/>"#,
    ] {
        let after = patch_styles(
            body,
            json!({"": {"lang": "de-DE", "eastAsia": "zh-Hant", "bidi": "ar-SA"}}),
        );
        assert!(after.contains(r#"<w:lang w:val="de-DE" w:eastAsia="zh-Hant" w:bidi="ar-SA"/>"#));
        if after.contains("<w:pPrDefault>") {
            assert!(after.find("<w:rPrDefault>") < after.find("<w:pPrDefault>"));
        }
        if after.contains("<w:latentStyles/>") {
            assert!(after.find("<w:docDefaults>") < after.find("<w:latentStyles/>"));
        }
    }
}

#[test]
fn defaults_patch_existing_language_without_reserializing_other_xml() {
    let body = r#"<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii='Arial'/><w:lang w:val='en-US' w:eastAsia="ja-JP"/></w:rPr></w:rPrDefault><w:pPrDefault/></w:docDefaults>"#;
    let after = patch_styles(body, json!({"": {"lang": "de-DE"}}));
    assert_eq!(
        after,
        format!(
            r#"<w:styles xmlns:w="{NS}">{}</w:styles>"#,
            body.replace("en-US", "de-DE")
        )
    );
    assert_eq!(
        patch_style_languages(&after, r#"{"":{"lang":"de-DE"}}"#).unwrap(),
        after
    );
}
