use super::{PatchContext, patch_paragraph_v1};
use crate::paragraph_patch::{
    AttrDeltaV1, Attrs, DonorV1, InsertedTextV1, ParagraphPatchV1, PatchErrorCode,
    RunEditV1, SeedV1, SlotV1, TextEditV1, UnitV1, result_sha256, unit_sha256,
};
use serde_json::{Value, json};

fn attrs(value: Value) -> Attrs {
    serde_json::from_value(value).unwrap()
}

fn text(value: &str, marks: Value) -> UnitV1 {
    UnitV1::Text {
        text: value.to_owned(),
        attrs: attrs(marks),
    }
}

fn patch(seed: Vec<UnitV1>, result: Vec<UnitV1>) -> ParagraphPatchV1 {
    ParagraphPatchV1 {
        ordinal: 0,
        session_key: "source:0".to_owned(),
        seed: SeedV1 {
            unit_sha256: unit_sha256(&seed),
            units: seed,
            properties: Attrs::new(),
        },
        text_edits: Vec::new(),
        run_edits: Vec::new(),
        para_deltas: Vec::new(),
        para_inherit: Attrs::new(),
        result_sha256: result_sha256(&result, &Attrs::new()),
    }
}

fn replace(plan: &mut ParagraphPatchV1, from: u32, to: u32, value: &str, marks: Value) {
    plan.text_edits.push(TextEditV1 {
        from,
        to,
        insert: vec![InsertedTextV1 {
            text: value.to_owned(),
            attrs: attrs(marks),
            inherit: Attrs::new(),
        }],
        donor: DonorV1::At { at: from },
    });
}

fn delta(key: &str, before: Option<Value>, after: Option<Value>) -> AttrDeltaV1 {
    let slot = |value: Option<Value>| match value {
        Some(value) => SlotV1::Present { value },
        None => SlotV1::Missing,
    };
    AttrDeltaV1 {
        path: vec![key.to_owned()],
        before: slot(before),
        after: slot(after),
    }
}

fn write(
    source: &str,
    plan: &ParagraphPatchV1,
) -> Result<String, crate::paragraph_patch::ParagraphPatchError> {
    patch_paragraph_v1(source, plan, &PatchContext::word_document())
}

#[test]
fn paragraph_patch_preserves_run_properties_and_entities() {
    let source = r#"<w:p><w:r><w:rPr><w:rFonts w:asciiTheme="minorHAnsi" w:hAnsiTheme="minorHAnsi"/><w:noProof/><w:color w:val="auto"/></w:rPr><w:t>😀 It&#x2019;s old &amp; fine</w:t></w:r></w:p>"#;
    let marks = json!({"fontFamily": {
        "ascii": null, "hAnsi": null, "eastAsia": null, "cs": null,
        "asciiTheme": "minorHAnsi", "hAnsiTheme": "minorHAnsi",
        "eastAsiaTheme": null, "csTheme": null
    }});
    let mut plan = patch(
        vec![text("😀 It’s old & fine", marks.clone())],
        vec![text("😀 It’s new & fine", marks.clone())],
    );
    replace(&mut plan, 8, 11, "new", marks);
    assert_eq!(write(source, &plan).unwrap(), source.replace("old", "new"));
}

#[test]
fn paragraph_patch_keeps_note_hyperlink_and_div_id() {
    let source = r#"<w:p><w:pPr><w:divId w:val="12"/></w:pPr><w:r><w:rPr><w:rStyle w:val="FootnoteReference"/><w:bCs/></w:rPr><w:footnoteReference w:id="2"/></w:r><w:hyperlink r:id="rId5" w:history="1"><w:r><w:t>old</w:t></w:r></w:hyperlink></w:p>"#;
    let note = UnitV1::Atom {
        atom_kind: "noteRef".to_owned(),
        payload: attrs(json!({"footnoteRefId": 2})),
        attrs: attrs(json!({"runStyle": {"styleId": "FootnoteReference"}})),
    };
    let link = json!({"hyperlink": {"href": "", "tooltip": null, "rId": "rId5"}});
    let mut plan = patch(
        vec![note.clone(), text("old", link.clone())],
        vec![note, text("new", link.clone())],
    );
    replace(&mut plan, 1, 4, "new", link);
    assert_eq!(write(source, &plan).unwrap(), source.replace("old", "new"));
}

#[test]
fn paragraph_patch_splits_run_and_inserts_bold_in_schema_order() {
    let source = r#"<w:p><w:r w:rsidR="12345678"><w:rPr><w:rFonts w:ascii="A"/><w:noProof/></w:rPr><w:t>abc</w:t></w:r></w:p>"#;
    let font = json!({"fontFamily": {
        "ascii": "A", "hAnsi": null, "eastAsia": null, "cs": null,
        "asciiTheme": null, "hAnsiTheme": null, "eastAsiaTheme": null, "csTheme": null
    }});
    let mut bold = font.clone();
    bold["bold"] = json!(true);
    let mut plan = patch(
        vec![text("abc", font.clone())],
        vec![
            text("a", font.clone()),
            text("b", bold.clone()),
            text("c", font),
        ],
    );
    plan.run_edits.push(RunEditV1 {
        from: 1,
        to: 2,
        deltas: vec![delta("bold", None, Some(json!(true)))],
        target: attrs(bold),
        inherit: Attrs::new(),
    });
    let run = |value: &str, extra: &str| {
        format!(
            r#"<w:r w:rsidR="12345678"><w:rPr><w:rFonts w:ascii="A"/>{extra}<w:noProof/></w:rPr><w:t>{value}</w:t></w:r>"#
        )
    };
    let expected = format!(
        "<w:p>{}{}{}</w:p>",
        run("a", ""),
        run("b", "<w:b/>"),
        run("c", ""),
    );
    assert_eq!(write(source, &plan).unwrap(), expected);
}

#[test]
fn paragraph_patch_replacement_uses_first_run_donor() {
    let source = "<w:p><w:r><w:rPr><w:b/><w:bCs/></w:rPr><w:t>one</w:t></w:r><w:r><w:t>two</w:t></w:r></w:p>";
    let mut plan = patch(
        vec![text("one", json!({"bold": true})), text("two", json!({}))],
        vec![text("new", json!({"bold": true}))],
    );
    replace(&mut plan, 0, 6, "new", json!({"bold": true}));
    assert_eq!(write(source, &plan).unwrap(),
        "<w:p><w:r><w:rPr><w:b/><w:bCs/></w:rPr><w:t>new</w:t></w:r><w:r><w:t></w:t></w:r></w:p>");
}

#[test]
fn paragraph_patch_start_and_empty_paragraph_donors() {
    let mut plan = patch(
        vec![text("tail", json!({"bold": true}))],
        vec![text("head tail", json!({"bold": true}))],
    );
    replace(&mut plan, 0, 0, "head ", json!({"bold": true}));
    assert_eq!(write("<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>tail</w:t></w:r></w:p>", &plan).unwrap(),
        "<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>head tail</w:t></w:r></w:p>");
    let mut plan = patch(Vec::new(), vec![text("new", json!({"italic": true}))]);
    replace(&mut plan, 0, 0, "new", json!({"italic": true}));
    plan.text_edits[0].donor = DonorV1::Pilcrow;
    assert_eq!(write("<w:p><w:pPr><w:rPr><w:i/></w:rPr></w:pPr></w:p>", &plan).unwrap(),
        "<w:p><w:pPr><w:rPr><w:i/></w:rPr></w:pPr><w:r><w:rPr><w:i/></w:rPr><w:t>new</w:t></w:r></w:p>");
}

#[test]
fn paragraph_patch_patches_ppr_children_without_rebuilding() {
    let source = r#"<w:p><w:pPr><w:pStyle w:val='Old'/><w:keepNext/><w:divId w:val="12"/><x:opaque data='keep'/></w:pPr><w:r><w:t>a</w:t></w:r></w:p>"#;
    let mut plan = patch(vec![text("a", json!({}))], vec![text("a", json!({}))]);
    plan.seed.properties = attrs(json!({"pStyle": "Old", "keepNext": true}));
    plan.para_deltas = vec![
        delta("pStyle", Some(json!("Old")), Some(json!("New"))),
        delta("alignment", None, Some(json!("center"))),
    ];
    plan.result_sha256 = result_sha256(&plan.seed.units, &attrs(json!({
        "pStyle": "New", "keepNext": true, "alignment": "center"
    })));
    assert_eq!(write(source, &plan).unwrap(), source
        .replace("'Old'", "'New'")
        .replace("<w:divId", "<w:jc w:val=\"center\"/><w:divId"));
}

#[test]
fn paragraph_patch_errors_have_stable_codes_and_display() {
    let source = "<w:p><w:r><w:t>a</w:t></w:r></w:p>";
    let plan = patch(vec![text("b", json!({}))], vec![text("b", json!({}))]);
    let error = write(source, &plan).unwrap_err();
    assert_eq!(error.code, PatchErrorCode::SeedMismatch);
    assert_eq!(error.to_string(),
        "DOCX_MINIMAL_PATCH [SEED_MISMATCH] word/document.xml p#0: source units differ from seed");
    let mut plan = patch(vec![text("a", json!({}))], vec![text("a", json!({}))]);
    assert_eq!(write(&source.replace("<w:t>", "<w:fldChar w:fldCharType=\"begin\"/><w:t>"), &plan).unwrap_err().code,
        PatchErrorCode::UnsupportedStructure);
    plan.run_edits.push(RunEditV1 {
        from: 0,
        to: 1,
        target: attrs(json!({"arbitrary": true})),
        inherit: Attrs::new(),
        deltas: vec![delta("arbitrary", None, Some(json!(true)))],
    });
    assert_eq!(write(source, &plan).unwrap_err().code, PatchErrorCode::UnsupportedKey);
}

#[test]
fn paragraph_patch_writes_explicit_inherited_off() {
    let source = "<w:p><w:r><w:rPr><w:b/><w:noProof/></w:rPr><w:t>x</w:t></w:r></w:p>";
    let mut plan = patch(vec![text("x", json!({"bold": true}))], vec![text("x", json!({}))]);
    plan.run_edits.push(RunEditV1 {
        from: 0,
        to: 1,
        deltas: vec![delta("bold", Some(json!(true)), None)],
        target: Attrs::new(),
        inherit: attrs(json!({"bold": true})),
    });
    assert_eq!(write(source, &plan).unwrap(), source.replace("<w:b/>", "<w:b w:val=\"0\"/>"));
}

#[test]
fn paragraph_patch_preserves_grouped_attributes() {
    let source = r#"<w:p><w:pPr><w:spacing w:before='20' w:after='40' x:keep='yes'/></w:pPr><w:r><w:t>x</w:t></w:r></w:p>"#;
    let mut plan = patch(vec![text("x", json!({}))], vec![text("x", json!({}))]);
    plan.seed.properties = attrs(json!({"spaceBefore": 20, "spaceAfter": 40}));
    plan.para_deltas = vec![delta("spaceBefore", Some(json!(20)), Some(json!(30)))];
    plan.result_sha256 = result_sha256(&plan.seed.units, &attrs(json!({"spaceBefore": 30, "spaceAfter": 40})));
    assert_eq!(write(source, &plan).unwrap(), source.replace("'20'", "'30'"));
}

#[test]
fn paragraph_patch_refuses_invalid_boundaries_markers_and_digest() {
    let source = "<w:p><w:r><w:t>😀x</w:t></w:r></w:p>";
    let mut plan = patch(vec![text("😀x", json!({}))], vec![text("😀y", json!({}))]);
    replace(&mut plan, 1, 3, "y", json!({}));
    assert_eq!(write(source, &plan).unwrap_err().code, PatchErrorCode::InvalidBoundary);
    plan.text_edits[0].from = 2;
    plan.seed.unit_sha256 = "0".repeat(64);
    assert_eq!(write(source, &plan).unwrap_err().code, PatchErrorCode::SeedMismatch);
    plan.seed.unit_sha256 = unit_sha256(&plan.seed.units);
    let marked = source.replace("<w:r>", "<w:bookmarkStart w:id=\"1\" w:name=\"mark\"/><w:r>");
    plan.text_edits[0].from = 0;
    assert_eq!(write(&marked, &plan).unwrap_err().code, PatchErrorCode::UnsupportedStructure);
}

#[test]
fn paragraph_patch_lowers_special_inline_units_without_changing_them() {
    let source = "<w:p><w:r><w:footnoteRef/><w:tab/><w:softHyphen/><w:noBreakHyphen/><w:sym w:font=\"A\" w:char=\"0041\"/><w:br/><w:t>old</w:t></w:r></w:p>";
    let line_break = UnitV1::Atom {
        atom_kind: "break".to_owned(),
        payload: Attrs::new(),
        attrs: Attrs::new(),
    };
    let symbol = text("A", json!({"fontFamily": {"ascii": "A", "hAnsi": "A", "eastAsia": "A", "cs": "A", "asciiTheme": null, "hAnsiTheme": null, "eastAsiaTheme": null, "csTheme": null}}));
    let mut plan = patch(
        vec![
            text("\t\u{ad}\u{2011}", json!({})),
            symbol.clone(),
            line_break.clone(),
            text("old", json!({})),
        ],
        vec![
            text("\t\u{ad}\u{2011}", json!({})),
            symbol,
            line_break,
            text("new", json!({})),
        ],
    );
    replace(&mut plan, 5, 8, "new", json!({}));
    assert_eq!(write(source, &plan).unwrap(), source.replace("old", "new"));
}
