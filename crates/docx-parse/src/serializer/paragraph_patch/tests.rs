use super::{PatchContext, patch_paragraph_v1};
use crate::paragraph_patch::{
    AttrDeltaV1, Attrs, DonorV1, InsertedTextV1, ParagraphPatchV1, PatchErrorCode, RunEditV1,
    SeedV1, SlotV1, TextEditV1, UnitV1, result_sha256, unit_sha256,
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
fn paragraph_patch_rejects_seed_colors_absent_from_explicit_source() {
    for (color, marks) in [
        (r#"w:val="auto""#, json!({"textColor": {"rgb": "FF0000"}})),
        (
            r#"w:val="FF0000""#,
            json!({"textColor": {"rgb": "FF0000", "themeColor": "accent1"}}),
        ),
        (
            r#"w:themeColor="accent1""#,
            json!({"textColor": {"rgb": "FF0000", "themeColor": "accent1"}}),
        ),
    ] {
        let source =
            format!("<w:p><w:r><w:rPr><w:color {color}/></w:rPr><w:t>a</w:t></w:r></w:p>");
        let mut plan = patch(
            vec![text("a", marks.clone())],
            vec![text("b", marks.clone())],
        );
        replace(&mut plan, 0, 1, "b", marks);
        assert_eq!(
            write(&source, &plan).unwrap_err().code,
            PatchErrorCode::SeedMismatch
        );
    }
}

#[test]
fn paragraph_patch_pilcrow_donation_strips_identifiers_from_clones() {
    let properties = r#"<w:rPr xml:id="D" w14:paraId="A" w14:textId="B"><w:b xml:id="E"/><w:noProof/></w:rPr>"#;
    let clean = "<w:rPr><w:b/><w:noProof/></w:rPr>";
    for count in [1, 2] {
        let mut plan = patch(
            Vec::new(),
            vec![text(&"x".repeat(count), json!({"bold": true}))],
        );
        plan.seed.properties = attrs(json!({"defaultTextFormatting": {"bold": true}}));
        plan.result_sha256 = result_sha256(
            &[text(&"x".repeat(count), json!({"bold": true}))],
            &plan.seed.properties,
        );
        replace(&mut plan, 0, 0, "x", json!({"bold": true}));
        plan.text_edits[0].donor = DonorV1::Pilcrow;
        plan.text_edits[0].insert = vec![plan.text_edits[0].insert[0].clone(); count];
        let source = format!("<w:p><w:pPr>{properties}</w:pPr></w:p>");
        assert_eq!(
            write(&source, &plan).unwrap(),
            format!(
                "<w:p><w:pPr>{properties}</w:pPr>{}</w:p>",
                format!("<w:r>{clean}<w:t>x</w:t></w:r>").repeat(count)
            )
        );
    }
    let source = r#"<w:p><w:r xml:id="R" w14:paraId="A" w14:textId="B"><w:rPr xml:id="D"><w:b/></w:rPr><w:t xml:id="T"/></w:r></w:p>"#;
    let mut plan = patch(Vec::new(), vec![text("x", json!({"bold": true}))]);
    replace(&mut plan, 0, 0, "x", json!({"bold": true}));
    plan.text_edits[0].donor = DonorV1::Pilcrow;
    assert_eq!(
        write(source, &plan).unwrap(),
        "<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>x</w:t></w:r></w:p>"
    );
}

#[test]
fn paragraph_patch_empty_donor_clears_shading_highlight() {
    let source = r#"<w:p><w:r><w:rPr><w:shd w:val="clear" w:fill="FF0000" w:themeFill="accent1" x:keep="yes"/></w:rPr><w:t/></w:r></w:p>"#;
    let mut plan = patch(Vec::new(), vec![text("x", json!({}))]);
    replace(&mut plan, 0, 0, "x", json!({}));
    plan.text_edits[0].donor = DonorV1::Pilcrow;
    assert_eq!(
        write(source, &plan).unwrap(),
        r#"<w:p><w:r><w:rPr><w:highlight w:val="none"/><w:shd w:val="clear" x:keep="yes"/></w:rPr><w:t>x</w:t></w:r></w:p>"#
    );
}

#[test]
fn paragraph_patch_empty_donors_stay_in_the_paragraph_scope() {
    for wrapper in [
        r#"<w:hyperlink r:id="rId1"><w:r><w:t/></w:r></w:hyperlink>"#,
        r#"<w:smartTag w:element="tag"><w:r><w:t/></w:r></w:smartTag>"#,
    ] {
        for empty in ["", "<w:r><w:t/></w:r>"] {
            let source = format!("<w:p>{wrapper}{empty}</w:p>");
            let mut plan = patch(Vec::new(), vec![text("x", json!({}))]);
            replace(&mut plan, 0, 0, "x", json!({}));
            plan.text_edits[0].donor = DonorV1::Pilcrow;
            assert_eq!(
                write(&source, &plan).unwrap(),
                format!("<w:p>{wrapper}<w:r><w:t>x</w:t></w:r></w:p>")
            );
        }
    }
}

#[test]
fn paragraph_patch_theme_colors_always_have_a_value() {
    let source = "<w:p><w:r><w:t>x</w:t></w:r></w:p>";
    for rgb in [None, Some(""), Some("123456")] {
        let marks = json!({"textColor": {"rgb": rgb, "themeColor": "accent1"}});
        let mut plan = patch(vec![text("x", json!({}))], vec![text("x", marks.clone())]);
        plan.run_edits.push(RunEditV1 {
            from: 0,
            to: 1,
            deltas: vec![delta("textColor", None, Some(marks["textColor"].clone()))],
            target: attrs(marks.clone()),
            inherit: Attrs::new(),
        });
        let expected = format!(
            r#"<w:p><w:r><w:rPr><w:color w:val="{}" w:themeColor="accent1"/></w:rPr><w:t>x</w:t></w:r></w:p>"#,
            rgb.filter(|value| !value.is_empty()).unwrap_or("auto")
        );
        assert_eq!(write(source, &plan).unwrap(), expected);
        let mut plan = patch(vec![text("x", json!({}))], vec![text("y", marks.clone())]);
        replace(&mut plan, 0, 1, "y", marks);
        assert_eq!(write(source, &plan).unwrap(), expected.replace(">x<", ">y<"));
    }
}

#[test]
fn paragraph_patch_comment_reference_keeps_empty_seed_attributes() {
    let source = r#"<w:p><w:r><w:rPr><w:rStyle w:val="CommentReference"/><w:b/></w:rPr><w:commentReference w:id="7"/></w:r><w:r><w:t>old</w:t></w:r></w:p>"#;
    let atom = UnitV1::Atom {
        atom_kind: "field".to_owned(),
        payload: attrs(json!({
            "fieldType": "COMMENT", "instruction": "", "displayText": "",
            "fieldKind": "simple", "fldLock": false, "dirty": false,
            "displayMode": "result", "hasCachedResult": false,
            "modelKind": "commentReference", "commentId": 7
        })),
        attrs: Attrs::new(),
    };
    let mut plan = patch(
        vec![atom.clone(), text("old", json!({}))],
        vec![atom, text("new", json!({}))],
    );
    replace(&mut plan, 1, 4, "new", json!({}));
    assert_eq!(write(source, &plan).unwrap(), source.replace("old", "new"));
}

#[test]
fn paragraph_patch_verifies_seeded_paragraph_keys_and_formatting_carriers() {
    for (properties, xml) in [
        (json!({"styleId": "Heading1"}), r#"<w:pStyle w:val="Heading1"/>"#),
        (
            json!({"_originalFormatting": {"indentLeftChars": 100, "indentRightChars": 200}}),
            r#"<w:ind w:leftChars="100" w:rightChars="200"/>"#,
        ),
        (
            json!({"_originalFormatting": {"indentFirstLineChars": -100, "hangingIndentChars": true}}),
            r#"<w:ind w:hangingChars="100"/>"#,
        ),
        (
            json!({"_originalFormatting": {"contextualSpacing": true}}),
            "<w:contextualSpacing/>",
        ),
    ] {
        let source = format!("<w:p><w:pPr>{xml}</w:pPr><w:r><w:t>a</w:t></w:r></w:p>");
        let mut plan = patch(vec![text("a", json!({}))], vec![text("b", json!({}))]);
        plan.seed.properties = attrs(properties);
        plan.result_sha256 = result_sha256(&[text("b", json!({}))], &plan.seed.properties);
        replace(&mut plan, 0, 1, "b", json!({}));
        assert_eq!(write(&source, &plan).unwrap(), source.replace(">a<", ">b<"));
    }
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
    assert_eq!(
        write(source, &plan).unwrap(),
        "<w:p><w:r><w:rPr><w:b/><w:bCs/></w:rPr><w:t>new</w:t></w:r><w:r><w:t></w:t></w:r></w:p>"
    );
}

#[test]
fn paragraph_patch_start_and_empty_paragraph_donors() {
    let mut plan = patch(
        vec![text("tail", json!({"bold": true}))],
        vec![text("head tail", json!({"bold": true}))],
    );
    replace(&mut plan, 0, 0, "head ", json!({"bold": true}));
    assert_eq!(
        write(
            "<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>tail</w:t></w:r></w:p>",
            &plan
        )
        .unwrap(),
        "<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>head tail</w:t></w:r></w:p>"
    );
    let mut plan = patch(Vec::new(), vec![text("new", json!({"italic": true}))]);
    replace(&mut plan, 0, 0, "new", json!({"italic": true}));
    plan.text_edits[0].donor = DonorV1::Pilcrow;
    plan.seed.properties = attrs(json!({"defaultTextFormatting": {"italic": true}}));
    plan.result_sha256 = result_sha256(
        &[text("new", json!({"italic": true}))],
        &plan.seed.properties,
    );
    assert_eq!(
        write("<w:p><w:pPr><w:rPr><w:i/></w:rPr></w:pPr></w:p>", &plan).unwrap(),
        "<w:p><w:pPr><w:rPr><w:i/></w:rPr></w:pPr><w:r><w:rPr><w:i/></w:rPr><w:t>new</w:t></w:r></w:p>"
    );
}

#[test]
fn paragraph_patch_patches_ppr_children_without_rebuilding() {
    let source = r#"<w:p><w:pPr><w:pStyle w:val='Old'/><w:keepNext/><w:divId w:val="12"/><x:opaque data='keep'/></w:pPr><w:r><w:t>a</w:t></w:r></w:p>"#;
    let mut plan = patch(vec![text("a", json!({}))], vec![text("a", json!({}))]);
    plan.seed.properties = attrs(json!({"styleId": "Old", "keepNext": true}));
    plan.para_deltas = vec![
        delta("styleId", Some(json!("Old")), Some(json!("New"))),
        delta("alignment", None, Some(json!("center"))),
    ];
    plan.result_sha256 = result_sha256(
        &plan.seed.units,
        &attrs(json!({
            "styleId": "New", "keepNext": true, "alignment": "center"
        })),
    );
    assert_eq!(
        write(source, &plan).unwrap(),
        source
            .replace("'Old'", "'New'")
            .replace("<w:divId", "<w:jc w:val=\"center\"/><w:divId")
    );
}

#[test]
fn paragraph_patch_errors_have_stable_codes_and_display() {
    let source = "<w:p><w:r><w:t>a</w:t></w:r></w:p>";
    let plan = patch(vec![text("b", json!({}))], vec![text("b", json!({}))]);
    let error = write(source, &plan).unwrap_err();
    assert_eq!(error.code, PatchErrorCode::SeedMismatch);
    assert_eq!(
        error.to_string(),
        "DOCX_MINIMAL_PATCH [SEED_MISMATCH] word/document.xml p#0: source units differ from seed"
    );
    let mut plan = patch(vec![text("a", json!({}))], vec![text("a", json!({}))]);
    assert_eq!(
        write(
            &source.replace("<w:t>", "<w:fldChar w:fldCharType=\"begin\"/><w:t>"),
            &plan
        )
        .unwrap_err()
        .code,
        PatchErrorCode::UnsupportedStructure
    );
    plan.run_edits.push(RunEditV1 {
        from: 0,
        to: 1,
        target: attrs(json!({"arbitrary": true})),
        inherit: Attrs::new(),
        deltas: vec![delta("arbitrary", None, Some(json!(true)))],
    });
    assert_eq!(
        write(source, &plan).unwrap_err().code,
        PatchErrorCode::UnsupportedKey
    );
}

#[test]
fn paragraph_patch_writes_explicit_inherited_off() {
    let source = "<w:p><w:r><w:rPr><w:b/><w:noProof/></w:rPr><w:t>x</w:t></w:r></w:p>";
    let mut plan = patch(
        vec![text("x", json!({"bold": true}))],
        vec![text("x", json!({}))],
    );
    plan.run_edits.push(RunEditV1 {
        from: 0,
        to: 1,
        deltas: vec![delta("bold", Some(json!(true)), None)],
        target: Attrs::new(),
        inherit: attrs(json!({"bold": true})),
    });
    assert_eq!(
        write(source, &plan).unwrap(),
        source.replace("<w:b/>", "<w:b w:val=\"0\"/>")
    );
}

#[test]
fn paragraph_patch_preserves_grouped_attributes() {
    let source = r#"<w:p><w:pPr><w:spacing w:before='20' w:after='40' x:keep='yes'/></w:pPr><w:r><w:t>x</w:t></w:r></w:p>"#;
    let mut plan = patch(vec![text("x", json!({}))], vec![text("x", json!({}))]);
    plan.seed.properties = attrs(json!({"spaceBefore": 20, "spaceAfter": 40}));
    plan.para_deltas = vec![delta("spaceBefore", Some(json!(20)), Some(json!(30)))];
    plan.result_sha256 = result_sha256(
        &plan.seed.units,
        &attrs(json!({"spaceBefore": 30, "spaceAfter": 40})),
    );
    assert_eq!(
        write(source, &plan).unwrap(),
        source.replace("'20'", "'30'")
    );
}

#[test]
fn paragraph_patch_refuses_invalid_boundaries_markers_and_digest() {
    let source = "<w:p><w:r><w:t>😀x</w:t></w:r></w:p>";
    let mut plan = patch(vec![text("😀x", json!({}))], vec![text("😀y", json!({}))]);
    replace(&mut plan, 1, 3, "y", json!({}));
    assert_eq!(
        write(source, &plan).unwrap_err().code,
        PatchErrorCode::InvalidBoundary
    );
    plan.text_edits[0].from = 2;
    plan.seed.unit_sha256 = "0".repeat(64);
    assert_eq!(
        write(source, &plan).unwrap_err().code,
        PatchErrorCode::SeedMismatch
    );
    plan.seed.unit_sha256 = unit_sha256(&plan.seed.units);
    let marked = source.replace(
        "<w:r>",
        "<w:bookmarkStart w:id=\"1\" w:name=\"mark\"/><w:r>",
    );
    plan.text_edits[0].from = 0;
    assert_eq!(
        write(&marked, &plan).unwrap_err().code,
        PatchErrorCode::UnsupportedStructure
    );
}

#[test]
fn paragraph_patch_lowers_special_inline_units_without_changing_them() {
    let source = "<w:p><w:r><w:footnoteRef/><w:tab/><w:softHyphen/><w:noBreakHyphen/><w:sym w:font=\"A\" w:char=\"0041\"/><w:br/><w:t>old</w:t></w:r></w:p>";
    let line_break = UnitV1::Atom {
        atom_kind: "break".to_owned(),
        payload: Attrs::new(),
        attrs: Attrs::new(),
    };
    let symbol = text(
        "A",
        json!({"fontFamily": {"ascii": "A", "hAnsi": "A", "eastAsia": "A", "cs": "A", "asciiTheme": null, "hAnsiTheme": null, "eastAsiaTheme": null, "csTheme": null}}),
    );
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

#[test]
fn paragraph_patch_rejects_wrong_direct_run_and_paragraph_seeds() {
    for (source, marks, properties) in [
        (
            "<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>a</w:t></w:r></w:p>",
            json!({}),
            json!({}),
        ),
        (
            "<w:p><w:r><w:rPr><w:b w:val='0'/></w:rPr><w:t>a</w:t></w:r></w:p>",
            json!({"bold": true}),
            json!({}),
        ),
        (
            "<w:p><w:pPr><w:jc w:val='center'/></w:pPr><w:r><w:t>a</w:t></w:r></w:p>",
            json!({}),
            json!({"alignment": "left"}),
        ),
        (
            "<w:p><w:pPr><w:spacing w:before='20'/></w:pPr><w:r><w:t>a</w:t></w:r></w:p>",
            json!({}),
            json!({"spaceBefore": 30}),
        ),
        (
            "<w:p><w:r><w:rPr><w:rFonts w:ascii='A'/></w:rPr><w:t>a</w:t></w:r></w:p>",
            json!({"fontFamily": {"ascii": "B"}}),
            json!({}),
        ),
        (
            "<w:p><w:pPr><w:rPr><w:i/></w:rPr></w:pPr><w:r><w:t>a</w:t></w:r></w:p>",
            json!({}),
            json!({"defaultTextFormatting": {"italic": false}}),
        ),
    ] {
        let mut plan = patch(
            vec![text("a", marks.clone())],
            vec![text("b", marks.clone())],
        );
        plan.seed.properties = attrs(properties);
        replace(&mut plan, 0, 1, "b", marks);
        assert_eq!(
            write(source, &plan).unwrap_err().code,
            PatchErrorCode::SeedMismatch
        );
    }
}

#[test]
fn paragraph_patch_normalizes_source_line_endings_like_seeding() {
    for raw in ["a\nb", "a\rb", "a\r\nb", "a&#13;&#10;b"] {
        let source = format!("<w:p><w:r><w:t>{raw}x</w:t></w:r></w:p>");
        let mut plan = patch(vec![text("a bx", json!({}))], vec![text("a by", json!({}))]);
        replace(&mut plan, 3, 4, "y", json!({}));
        assert_eq!(write(&source, &plan).unwrap(), source.replace("bx", "by"));
    }
}

#[test]
fn paragraph_patch_refuses_inserted_line_endings() {
    for value in ["a\nb", "a\rb", "a\r\nb"] {
        let mut plan = patch(vec![text("x", json!({}))], vec![text(value, json!({}))]);
        replace(&mut plan, 0, 1, value, json!({}));
        let error = write("<w:p><w:r><w:t>x</w:t></w:r></w:p>", &plan).unwrap_err();
        assert_eq!(error.code, PatchErrorCode::UnsupportedStructure);
        assert_eq!(error.detail, "ambiguous: newline in inserted text");
    }
}

#[test]
fn paragraph_patch_color_removal_keeps_required_attributes_valid() {
    let source = "<w:p><w:r><w:rPr><w:color w:val='FF0000'/></w:rPr><w:t>x</w:t></w:r></w:p>";
    let color = json!({"textColor": {"rgb": "FF0000"}});
    let mut plan = patch(vec![text("x", color.clone())], vec![text("x", json!({}))]);
    plan.run_edits.push(RunEditV1 {
        from: 0,
        to: 1,
        deltas: vec![delta("textColor", Some(color["textColor"].clone()), None)],
        target: Attrs::new(),
        inherit: Attrs::new(),
    });
    assert_eq!(
        write(source, &plan).unwrap(),
        source.replace("<w:color w:val='FF0000'/>", "")
    );
    let opaque = source.replace("w:val='FF0000'", "w:val='FF0000' x:keep='yes'");
    assert_eq!(
        write(&opaque, &plan).unwrap_err().code,
        PatchErrorCode::PropertyConflict
    );
}

#[test]
fn paragraph_patch_indentation_uses_source_aliases() {
    for (key, spelling) in [("indentLeft", "start"), ("indentRight", "end")] {
        for after in [None, Some(json!(360))] {
            let source = format!("<w:p><w:pPr><w:ind w:{spelling}='720'/></w:pPr><w:r><w:t>x</w:t></w:r></w:p>");
            let mut plan = patch(vec![text("x", json!({}))], vec![text("x", json!({}))]);
            plan.seed.properties.insert(key.to_owned(), json!(720));
            plan.para_deltas
                .push(delta(key, Some(json!(720)), after.clone()));
            let target = after
                .clone()
                .map(|value| Attrs::from([(key.to_owned(), value)]))
                .unwrap_or_default();
            plan.result_sha256 = result_sha256(&plan.seed.units, &target);
            let expected = if after.is_some() {
                source.replace("'720'", "'360'")
            } else {
                source.replace(&format!("<w:ind w:{spelling}='720'/>"), "<w:ind/>")
            };
            assert_eq!(write(&source, &plan).unwrap(), expected);
        }
    }
}

#[test]
fn paragraph_patch_first_line_removal_checks_inherited_pair() {
    for (key, hanging, attribute) in [
        ("indentFirstLine", "hangingIndent", "firstLine"),
        ("indentFirstLineChars", "hangingIndentChars", "firstLineChars"),
    ] {
        let source = format!("<w:p><w:pPr><w:ind w:{attribute}='360'/></w:pPr><w:r><w:t>x</w:t></w:r></w:p>");
        let mut plan = patch(vec![text("x", json!({}))], vec![text("x", json!({}))]);
        plan.seed.properties.insert(key.to_owned(), json!(360));
        plan.para_inherit.insert(key.to_owned(), json!(720));
        plan.para_inherit.insert(hanging.to_owned(), json!(true));
        plan.para_deltas.push(delta(key, Some(json!(360)), None));
        assert_eq!(
            write(&source, &plan).unwrap_err().code,
            PatchErrorCode::PropertyConflict
        );
    }
}

#[test]
fn paragraph_patch_clear_highlight_removes_theme_fill() {
    let source = "<w:p><w:r><w:rPr><w:shd w:val='clear' w:fill='FF0000' w:themeFill='accent1' w:themeFillTint='80' w:themeFillShade='40'/></w:rPr><w:t>x</w:t></w:r></w:p>";
    let highlight = json!({"highlight": {"color": "#FF0000"}});
    let mut plan = patch(vec![text("x", highlight.clone())], vec![text("x", json!({}))]);
    plan.run_edits.push(RunEditV1 {
        from: 0,
        to: 1,
        deltas: vec![delta("highlight", Some(highlight["highlight"].clone()), None)],
        target: Attrs::new(),
        inherit: Attrs::new(),
    });
    let output = write(source, &plan).unwrap();
    assert!(!output.contains("themeFill"), "{output}");
    assert!(!output.contains("FF0000"), "{output}");
}

#[test]
fn paragraph_patch_empty_donor_resets_inherited_bold() {
    let mut plan = patch(Vec::new(), vec![text("x", json!({}))]);
    replace(&mut plan, 0, 0, "x", json!({}));
    plan.text_edits[0].donor = DonorV1::Pilcrow;
    plan.text_edits[0].insert[0].inherit = attrs(json!({"bold": true}));
    assert_eq!(
        write("<w:p/>", &plan).unwrap(),
        "<w:p><w:r><w:rPr><w:b w:val=\"0\"/></w:rPr><w:t>x</w:t></w:r></w:p>"
    );
}

#[test]
fn paragraph_patch_empty_donor_keeps_inherited_font_slots_in_reset_checks() {
    let source = "<w:p><w:r><w:rPr><w:rFonts w:ascii='A'/></w:rPr><w:t/></w:r></w:p>";
    let mut plan = patch(Vec::new(), vec![text("x", json!({}))]);
    replace(&mut plan, 0, 0, "x", json!({}));
    plan.text_edits[0].donor = DonorV1::Pilcrow;
    plan.text_edits[0].insert[0].inherit = attrs(json!({"fontFamily": {"hAnsi": "B"}}));
    assert_eq!(
        write(source, &plan).unwrap_err().code,
        PatchErrorCode::PropertyConflict
    );
}

#[test]
fn paragraph_patch_text_identifier_cannot_be_duplicated() {
    let source = "<w:p><w:r><w:t xml:id='T'>abc</w:t></w:r></w:p>";
    let mut plan = patch(
        vec![text("abc", json!({}))],
        vec![
            text("a", json!({})),
            text("b", json!({"bold": true})),
            text("c", json!({})),
        ],
    );
    plan.run_edits.push(RunEditV1 {
        from: 1,
        to: 2,
        deltas: vec![delta("bold", None, Some(json!(true)))],
        target: attrs(json!({"bold": true})),
        inherit: Attrs::new(),
    });
    assert_eq!(
        write(source, &plan).unwrap_err().code,
        PatchErrorCode::UnsupportedStructure
    );
    let mut plan = patch(
        Vec::new(),
        vec![text("a", json!({})), text("b", json!({"bold": true}))],
    );
    replace(&mut plan, 0, 0, "a", json!({}));
    plan.text_edits[0].donor = DonorV1::Pilcrow;
    plan.text_edits[0].insert.push(InsertedTextV1 {
        text: "b".to_owned(),
        attrs: attrs(json!({"bold": true})),
        inherit: Attrs::new(),
    });
    assert_eq!(
        write("<w:p><w:r><w:t xml:id='T'/></w:r></w:p>", &plan).unwrap(),
        "<w:p><w:r><w:t>a</w:t></w:r><w:r><w:rPr><w:b/></w:rPr><w:t>b</w:t></w:r></w:p>"
    );
}

#[test]
fn paragraph_patch_preserves_proof_errors_and_smart_tags() {
    for source in [
        "<w:p><w:proofErr w:type='spellStart'/><w:r><w:rPr><w:noProof/></w:rPr><w:t>old</w:t></w:r><w:proofErr w:type='spellEnd'/></w:p>",
        "<w:p><w:smartTag w:element='name'><w:r><w:rPr><w:noProof/></w:rPr><w:t>old</w:t></w:r></w:smartTag></w:p>",
        "<w:p><w:smartTag w:element='name'><w:smartTagPr><w:attr w:name='k' w:val='v'/></w:smartTagPr><w:r><w:rPr><w:noProof/></w:rPr><w:t>old</w:t></w:r></w:smartTag></w:p>",
    ] {
        let mut plan = patch(vec![text("old", json!({}))], vec![text("new", json!({}))]);
        replace(&mut plan, 0, 3, "new", json!({}));
        assert_eq!(write(source, &plan).unwrap(), source.replace("old", "new"));
    }
}

#[test]
fn paragraph_patch_writer_advertises_version_one() {
    assert_eq!(crate::paragraph_patch::PARAGRAPH_PATCH_WRITER_VERSION, 1);
}

#[test]
fn paragraph_patch_first_line_pair_needs_an_amount() {
    let source = "<w:p><w:pPr><w:ind w:hanging='360'/></w:pPr><w:r><w:t>x</w:t></w:r></w:p>";
    let mut plan = patch(vec![text("x", json!({}))], vec![text("x", json!({}))]);
    plan.seed.properties = attrs(json!({"indentFirstLine": -360, "hangingIndent": true}));
    plan.para_deltas
        .push(delta("indentFirstLine", Some(json!(-360)), None));
    plan.result_sha256 = result_sha256(
        &plan.seed.units,
        &attrs(json!({"hangingIndent": true})),
    );
    assert_eq!(
        write(source, &plan).unwrap_err().code,
        PatchErrorCode::PropertyConflict
    );
}

#[cfg(feature = "wasm")]
#[test]
fn paragraph_patch_wasm_capability_matches_writer() {
    assert_eq!(
        crate::paragraph_patch_writer_version(),
        crate::paragraph_patch::PARAGRAPH_PATCH_WRITER_VERSION
    );
}
