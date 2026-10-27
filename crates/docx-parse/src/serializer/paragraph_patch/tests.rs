use super::{PatchContext, patch_paragraph_v1};
use crate::paragraph_patch::{
    AttrDeltaV1, Attrs, DonorV1, InheritanceV1, InsertedTextV1, ParagraphPatchV1, PatchErrorCode,
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
            inherit: InheritanceV1::default(),
        },
        text_edits: Vec::new(),
        run_edits: Vec::new(),
        para_deltas: Vec::new(),
        para_inherit: Attrs::new(),
        mark_inherit: Attrs::new(),
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

fn marker_properties(start: u32, end: u32) -> Attrs {
    attrs(json!({
        "bookmarks": [
            {"id": 7, "name": "mark", "kind": "start", "offset": start, "colFirst": 0},
            {"id": 7, "kind": "end", "offset": end}
        ],
        "markers": [[start, "1", false], [end, "1", true]]
    }))
}

const MARKED_PARAGRAPH: &str = concat!(
    "<w:p><w:r><w:t>😀a</w:t></w:r>",
    "<w:bookmarkStart w:id=\"7\" w:name=\"mark\" w:colFirst=\"0\"/>",
    "<w:commentRangeStart w:id=\"1\"/><w:r><w:t>bc</w:t></w:r>",
    "<w:commentRangeEnd w:id=\"1\"/><w:bookmarkEnd w:id=\"7\"/>",
    "<w:r><w:t>d</w:t></w:r></w:p>"
);

#[test]
fn paragraph_patch_hashes_transformed_bookmark_and_comment_positions() {
    let target = vec![text("Zabcd", json!({}))];
    let mut plan = patch(vec![text("😀abcd", json!({}))], target.clone());
    plan.seed.properties = marker_properties(3, 5);
    replace(&mut plan, 0, 2, "Z", json!({}));
    plan.result_sha256 = result_sha256(&target, &marker_properties(2, 4));
    assert_eq!(
        write(MARKED_PARAGRAPH, &plan).unwrap(),
        MARKED_PARAGRAPH.replace("😀a", "Za")
    );
    plan.result_sha256 = result_sha256(&target, &plan.seed.properties);
    assert_eq!(
        write(MARKED_PARAGRAPH, &plan).unwrap_err().code,
        PatchErrorCode::ResultMismatch
    );
}

#[test]
fn paragraph_patch_checks_seed_marker_positions_and_identities() {
    for key in ["bookmarks", "markers"] {
        for identity in [false, true] {
            let mut plan = patch(
                vec![text("😀abcd", json!({}))],
                vec![text("😀abcd", json!({}))],
            );
            plan.seed.properties = marker_properties(3, 5);
            if key == "bookmarks" {
                let field = if identity { "id" } else { "offset" };
                plan.seed.properties.get_mut(key).unwrap()[0][field] = json!(9);
            } else {
                let field = usize::from(identity);
                plan.seed.properties.get_mut(key).unwrap()[0][field] =
                    if identity { json!("9") } else { json!(9) };
            }
            plan.result_sha256 = result_sha256(&plan.seed.units, &plan.seed.properties);
            assert_eq!(
                write(MARKED_PARAGRAPH, &plan).unwrap_err().code,
                PatchErrorCode::SeedMismatch
            );
        }
    }
}

#[test]
fn paragraph_patch_seed_color_disagreements_request_ambiguous_replace() {
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
        let source = format!("<w:p><w:r><w:rPr><w:color {color}/></w:rPr><w:t>a</w:t></w:r></w:p>");
        let mut plan = patch(
            vec![text("a", marks.clone())],
            vec![text("b", marks.clone())],
        );
        replace(&mut plan, 0, 1, "b", marks);
        let error = write(&source, &plan).unwrap_err();
        assert_eq!(error.code, PatchErrorCode::UnsupportedStructure);
        assert!(error.detail.starts_with("ambiguous:"));
    }
}

#[test]
fn paragraph_patch_pilcrow_donation_strips_identifiers_from_clones() {
    let properties =
        r#"<w:rPr xml:id="D" w14:paraId="A" w14:textId="B"><w:b xml:id="E"/><w:noProof/></w:rPr>"#;
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
        assert_eq!(
            write(source, &plan).unwrap(),
            expected.replace(">x<", ">y<")
        );
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
        (
            json!({"styleId": "Heading1"}),
            r#"<w:pStyle w:val="Heading1"/>"#,
        ),
        (
            json!({"pStyle": "Heading1"}),
            r#"<w:pStyle w:val="Heading1"/>"#,
        ),
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

const RUN_START: &str = concat!(
    "<w:r  w:rsidR='00112233' w:rsidRPr=\"00575701\"\n",
    " w:rsidDel = '00947A26' xmlns:x='urn:run-metadata'",
    " x:keep = 'A&#x26;B' w:unknown=\"opaque\">"
);

const FOREIGN_ID_RUN_START: &str = concat!(
    "<w:r w:rsidR='00112233' w:rsidRPr=\"00575701\" w:rsidDel='00947A26'",
    " xmlns:x='urn:run-metadata' x:id = 'foreign' x:paraId='P' x:textId=\"T\">"
);

#[test]
fn paragraph_patch_run_start_attributes_survive_on_untouched_runs() {
    let cases = [("keep", "old", "keepold", 4), ("old", "keep", "oldkeep", 0)];
    for (before, after, seed, from) in cases {
        let source = format!(
            "<w:p>{RUN_START}<w:t>{before}</w:t></w:r>{RUN_START}<w:t>{after}</w:t></w:r></w:p>"
        );
        let mut plan = patch(
            vec![text(seed, json!({}))],
            vec![text(&seed.replace("old", "new"), json!({}))],
        );
        replace(&mut plan, from, from + 3, "new", json!({}));
        let expected = source.replace("old", "new");
        assert_eq!(
            write(&source, &plan).unwrap().as_bytes(),
            expected.as_bytes()
        );
    }
}

#[test]
fn paragraph_patch_run_start_attributes_survive_in_place_text_edits() {
    for start in [RUN_START, FOREIGN_ID_RUN_START] {
        let source = format!("<w:p>{start}<w:t>old</w:t></w:r></w:p>");
        let mut plan = patch(vec![text("old", json!({}))], vec![text("new", json!({}))]);
        replace(&mut plan, 0, 3, "new", json!({}));
        let expected = source.replace("old", "new");
        assert_eq!(
            write(&source, &plan).unwrap().as_bytes(),
            expected.as_bytes()
        );
    }
}

#[test]
fn paragraph_patch_run_start_attributes_survive_both_halves_of_text_splits() {
    for start in [RUN_START, FOREIGN_ID_RUN_START] {
        let source = format!("<w:p>{start}<w:t>abc</w:t></w:r></w:p>");
        let mut plan = patch(
            vec![text("abc", json!({}))],
            vec![
                text("a", json!({})),
                text("B", json!({"bold": true})),
                text("c", json!({})),
            ],
        );
        replace(&mut plan, 1, 2, "B", json!({"bold": true}));
        let expected = format!(
            "<w:p>{start}<w:t>a</w:t></w:r>\
             {start}<w:rPr><w:b/></w:rPr><w:t>B</w:t></w:r>\
             {start}<w:t>c</w:t></w:r></w:p>"
        );
        assert_eq!(
            write(&source, &plan).unwrap().as_bytes(),
            expected.as_bytes()
        );
    }
}

#[test]
fn paragraph_patch_run_start_attributes_are_cloned_verbatim_for_inserted_runs() {
    for start in [RUN_START, FOREIGN_ID_RUN_START] {
        let source = format!("<w:p>{start}<w:t>ab</w:t></w:r></w:p>");
        let mut plan = patch(
            vec![text("ab", json!({}))],
            vec![
                text("a", json!({})),
                text("X", json!({"bold": true})),
                text("b", json!({})),
            ],
        );
        replace(&mut plan, 1, 1, "X", json!({"bold": true}));
        plan.text_edits[0].donor = DonorV1::At { at: 0 };
        let expected = format!(
            "<w:p>{start}<w:t>a</w:t></w:r>\
             {start}<w:rPr><w:b/></w:rPr><w:t>X</w:t></w:r>\
             {start}<w:t>b</w:t></w:r></w:p>"
        );
        assert_eq!(
            write(&source, &plan).unwrap().as_bytes(),
            expected.as_bytes()
        );
    }
}

#[test]
fn paragraph_patch_run_start_attributes_are_cloned_from_empty_donors() {
    for start in [RUN_START, FOREIGN_ID_RUN_START] {
        let start = start.replace(" xmlns:x='urn:run-metadata'", "");
        let source = format!("<w:p xmlns:x='urn:run-metadata'>{start}<w:t/></w:r></w:p>");
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
        let expected = format!(
            "<w:p xmlns:x='urn:run-metadata'>{start}<w:t>a</w:t></w:r>\
             {start}<w:rPr><w:b/></w:rPr><w:t>b</w:t></w:r></w:p>"
        );
        assert_eq!(
            write(&source, &plan).unwrap().as_bytes(),
            expected.as_bytes()
        );
    }
}

#[test]
fn paragraph_patch_run_start_attributes_survive_next_to_note_atoms() {
    for reference in [false, true] {
        let marker = if reference {
            "<w:footnoteReference w:id=\"3\"/>"
        } else {
            "<w:footnoteRef/>"
        };
        let source = format!(
            "<w:p>{RUN_START}<w:rPr><w:rStyle w:val=\"FootnoteReference\"/></w:rPr>\
             {marker}</w:r>{RUN_START}<w:t xml:space=\"preserve\"> Some note text</w:t></w:r></w:p>"
        );
        let units = |value| {
            let mut units = Vec::new();
            if reference {
                units.push(UnitV1::Atom {
                    atom_kind: "noteRef".to_owned(),
                    payload: attrs(json!({"footnoteRefId": 3})),
                    attrs: attrs(json!({"runStyle": {"styleId": "FootnoteReference"}})),
                });
            }
            units.push(text(value, json!({})));
            units
        };
        let mut plan = patch(units(" Some note text"), units(" Some edited text"));
        let offset = u32::from(reference);
        replace(&mut plan, 6 + offset, 10 + offset, "edited", json!({}));
        let expected = source.replace(" Some note text", " Some edited text");
        assert_eq!(
            write(&source, &plan).unwrap().as_bytes(),
            expected.as_bytes()
        );
    }
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
fn paragraph_patch_drops_only_emptied_text_runs() {
    for (tail, kept) in [
        ("<w:t>]</w:t>", ""),
        ("<w:t>]suffix</w:t>", "<w:t>suffix</w:t>"),
        ("<w:t>]</w:t><w:tab/>", "<w:t></w:t><w:tab/>"),
        (
            "<w:t>]</w:t><w:lastRenderedPageBreak/>",
            "<w:t></w:t><w:lastRenderedPageBreak/>",
        ),
    ] {
        let empty = "<w:r><w:rPr/><w:t/></w:r>";
        let source = format!(
            "<w:p>{empty}<w:r><w:rPr/><w:t>[</w:t></w:r><w:r><w:rPr/><w:t>abc</w:t></w:r><w:r>{tail}</w:r></w:p>"
        );
        let suffix = if tail.contains("suffix") {
            "suffix"
        } else {
            ""
        };
        let special = if tail.contains("tab") { "\t" } else { "" };
        let mut plan = patch(
            vec![text(&format!("[abc]{suffix}{special}"), json!({}))],
            vec![text(&format!("abc{suffix}{special}"), json!({}))],
        );
        replace(&mut plan, 0, 1, "", json!({}));
        replace(&mut plan, 4, 5, "", json!({}));
        let tail = if kept.is_empty() {
            String::new()
        } else {
            format!("<w:r>{kept}</w:r>")
        };
        assert_eq!(
            write(&source, &plan).unwrap(),
            format!("<w:p>{empty}<w:r><w:rPr/><w:t>abc</w:t></w:r>{tail}</w:p>")
        );
    }
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
        "<w:p><w:r><w:rPr><w:b/><w:bCs/></w:rPr><w:t>new</w:t></w:r></w:p>"
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
    plan.para_deltas = vec![delta("alignment", None, Some(json!("center")))];
    plan.result_sha256 = result_sha256(
        &plan.seed.units,
        &attrs(json!({
            "styleId": "Old", "keepNext": true, "alignment": "center"
        })),
    );
    assert_eq!(
        write(source, &plan).unwrap(),
        source.replace("<w:divId", "<w:jc w:val=\"center\"/><w:divId")
    );
    plan.para_deltas
        .insert(0, delta("styleId", Some(json!("Old")), Some(json!("New"))));
    plan.result_sha256 = result_sha256(
        &plan.seed.units,
        &attrs(json!({
            "styleId": "New", "keepNext": true, "alignment": "center"
        })),
    );
    let error = write(source, &plan).unwrap_err();
    assert_eq!(error.code, PatchErrorCode::UnsupportedStructure);
    assert!(error.detail.starts_with("restyle:"), "{}", error.detail);
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
    plan.seed.inherit.runs = vec![attrs(json!({"bold": true}))];
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
fn paragraph_patch_seed_property_disagreements_request_ambiguous_replace() {
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
        let error = write(source, &plan).unwrap_err();
        assert_eq!(error.code, PatchErrorCode::UnsupportedStructure);
        assert!(error.detail.starts_with("ambiguous:"));
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
            let source = format!(
                "<w:p><w:pPr><w:ind w:{spelling}='720'/></w:pPr><w:r><w:t>x</w:t></w:r></w:p>"
            );
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
fn paragraph_patch_negative_first_line_indent_is_ambiguous() {
    for (key, attribute) in [
        ("indentFirstLine", "firstLine"),
        ("indentFirstLineChars", "firstLineChars"),
    ] {
        let source = format!(
            "<w:p><w:pPr><w:ind w:{attribute}='100'/></w:pPr><w:r><w:t>x</w:t></w:r></w:p>"
        );
        let mut plan = patch(vec![text("x", json!({}))], vec![text("x", json!({}))]);
        plan.seed.properties.insert(key.to_owned(), json!(100));
        plan.para_deltas
            .push(delta(key, Some(json!(100)), Some(json!(-200))));
        plan.result_sha256 = result_sha256(
            &plan.seed.units,
            &Attrs::from([(key.to_owned(), json!(-200))]),
        );
        let error = write(&source, &plan).unwrap_err();
        assert_eq!(error.code, PatchErrorCode::AmbiguousAlignment);
        assert_eq!(error.detail, format!("ambiguous: indent kind {key}"));
    }
}

#[test]
fn paragraph_patch_first_line_removal_checks_inherited_pair() {
    for (key, hanging, attribute) in [
        ("indentFirstLine", "hangingIndent", "firstLine"),
        (
            "indentFirstLineChars",
            "hangingIndentChars",
            "firstLineChars",
        ),
    ] {
        let source = format!(
            "<w:p><w:pPr><w:ind w:{attribute}='360'/></w:pPr><w:r><w:t>x</w:t></w:r></w:p>"
        );
        let mut plan = patch(vec![text("x", json!({}))], vec![text("x", json!({}))]);
        plan.seed.properties.insert(key.to_owned(), json!(360));
        plan.para_inherit.insert(key.to_owned(), json!(720));
        plan.para_inherit.insert(hanging.to_owned(), json!(true));
        plan.seed.inherit.paragraph = plan.para_inherit.clone();
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
    let mut plan = patch(
        vec![text("x", highlight.clone())],
        vec![text("x", json!({}))],
    );
    plan.run_edits.push(RunEditV1 {
        from: 0,
        to: 1,
        deltas: vec![delta(
            "highlight",
            Some(highlight["highlight"].clone()),
            None,
        )],
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
    plan.mark_inherit = plan.text_edits[0].insert[0].inherit.clone();
    plan.seed.inherit.paragraph_mark = plan.mark_inherit.clone();
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
    plan.mark_inherit = plan.text_edits[0].insert[0].inherit.clone();
    plan.seed.inherit.paragraph_mark = plan.mark_inherit.clone();
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
    plan.result_sha256 = result_sha256(&plan.seed.units, &attrs(json!({"hangingIndent": true})));
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

fn seed_inheritance(plan: &mut ParagraphPatchV1, paragraph: Value, runs: Value, mark: Value) {
    plan.seed.inherit = InheritanceV1 {
        paragraph: attrs(paragraph),
        runs: serde_json::from_value(runs).unwrap(),
        paragraph_mark: attrs(mark),
    };
}

fn assert_restyle_failure(source: &str, plan: &ParagraphPatchV1) {
    let error = write(source, plan).unwrap_err();
    assert_eq!(error.code, PatchErrorCode::UnsupportedStructure);
    assert!(error.detail.starts_with("restyle:"), "{error}");
}

#[test]
fn paragraph_patch_changed_run_inheritance_requests_restyle_replace() {
    for (source_properties, marks, old, new) in [
        (
            r#"<w:rFonts w:ascii='Courier New'/><w:noProof/>"#,
            json!({"fontFamily": {"ascii": "Courier New", "hAnsi": "Arial"}}),
            json!({"fontFamily": {"ascii": "Arial", "hAnsi": "Arial"}}),
            json!({"fontFamily": {"ascii": "Times", "hAnsi": "Times"}}),
        ),
        (
            r#"<w:noProof/><w:sz w:val='30'/>"#,
            json!({"fontSize": {"size": 30, "sizeCs": 24}}),
            json!({"fontSize": {"size": 24, "sizeCs": 24}}),
            json!({"fontSize": {"size": 40, "sizeCs": 40}}),
        ),
        (
            "<w:noProof/>",
            json!({"textColor": {"rgb": "123456"}, "bold": true, "italic": true}),
            json!({"textColor": {"rgb": "123456"}, "bold": true, "italic": true}),
            json!({"textColor": {"rgb": "ABCDEF"}}),
        ),
        ("<w:noProof/>", json!({}), json!({}), json!({"bold": true})),
        (
            "<w:dstrike w:val='0'/><w:noProof/>",
            json!({}),
            json!({"strike": {"double": true}}),
            json!({"strike": {"double": false}}),
        ),
        (
            "<w:noProof/>",
            json!({"strike": {"double": false}}),
            json!({"strike": {"double": false}}),
            json!({}),
        ),
        (
            "<w:dstrike w:val='0'/><w:noProof/>",
            json!({"strike": {"double": false}}),
            json!({"strike": {"double": false}}),
            json!({"strike": {"double": true}}),
        ),
        (
            "<w:bCs/>",
            json!({}),
            json!({"boldCs": true}),
            json!({"boldCs": false}),
        ),
        (
            "<w:iCs/>",
            json!({}),
            json!({"italicCs": true}),
            json!({"italicCs": false}),
        ),
        (
            "<w:shd w:val='clear' w:fill='FF0000'/>",
            json!({"highlight": {"color": "#FF0000"}}),
            json!({}),
            json!({"highlight": {"color": "yellow"}}),
        ),
        ("", json!({}), json!({}), json!({"unknown": "new"})),
        (
            "",
            json!({}),
            json!({}),
            json!({"fontSize": {"sizeCs": 40}}),
        ),
    ] {
        let source =
            format!("<w:p><w:r><w:rPr>{source_properties}</w:rPr><w:t>x</w:t></w:r></w:p>");
        let units = vec![text("x", marks.clone())];
        let mut plan = patch(units.clone(), units);
        seed_inheritance(&mut plan, json!({}), json!([old]), json!({}));
        plan.run_edits.push(RunEditV1 {
            from: 0,
            to: 1,
            deltas: Vec::new(),
            target: attrs(marks),
            inherit: attrs(new),
        });
        assert_restyle_failure(&source, &plan);
    }
}

#[test]
fn paragraph_patch_text_donor_changed_inheritance_requests_restyle_replace() {
    let source = concat!(
        "<w:p><w:r><w:rPr><w:dstrike w:val='0'/><w:noProof/></w:rPr>",
        "<w:t>x</w:t></w:r></w:p>",
    );
    let mut plan = patch(vec![text("x", json!({}))], vec![text("y", json!({}))]);
    replace(&mut plan, 0, 1, "y", json!({}));
    plan.text_edits[0].insert[0].inherit = attrs(json!({"strike": {"double": false}}));
    assert_restyle_failure(source, &plan);
}

#[test]
fn paragraph_patch_style_and_numbering_deltas_request_restyle_replace() {
    for (key, before, after) in [
        ("pStyle", json!("A"), json!("B")),
        (
            "numPr",
            json!({"numId": 1, "ilvl": 0}),
            json!({"numId": 2, "ilvl": 0}),
        ),
    ] {
        let units = vec![text("x", json!({}))];
        let mut plan = patch(units.clone(), units);
        plan.seed.properties.insert(key.to_owned(), before.clone());
        plan.para_deltas = vec![delta(key, Some(before), Some(after))];
        assert_restyle_failure("<w:p><w:r><w:t>x</w:t></w:r></w:p>", &plan);
    }
    let mut plan = patch(vec![text("x", json!({}))], vec![text("x", json!({}))]);
    plan.run_edits.push(RunEditV1 {
        from: 0,
        to: 1,
        deltas: vec![delta("runStyle", None, Some(json!({"styleId": "C"})))],
        target: attrs(json!({"runStyle": {"styleId": "C"}})),
        inherit: Attrs::new(),
    });
    assert_restyle_failure("<w:p><w:r><w:t>x</w:t></w:r></w:p>", &plan);
}

#[test]
fn paragraph_patch_retained_paragraph_inheritance_requests_restyle_replace() {
    let source = concat!(
        "<w:p><w:pPr><w:pStyle w:val='A'/><w:spacing w:before='120'/>",
        "<w:ind w:left='360'/><w:divId w:val='7'/></w:pPr><w:r><w:t>x</w:t></w:r></w:p>",
    );
    for (old, new) in [
        (
            json!({"spaceAfter": 240, "indentRight": 720}),
            json!({"spaceAfter": 480, "indentRight": 1080}),
        ),
        (json!({"widowControl": false}), json!({})),
        (
            json!({"widowControl": false}),
            json!({"widowControl": null}),
        ),
        (json!({}), json!({"bidi": true})),
    ] {
        let units = vec![text("x", json!({}))];
        let mut plan = patch(units.clone(), units);
        seed_inheritance(&mut plan, old, json!([{}]), json!({}));
        plan.para_inherit = attrs(new);
        assert_restyle_failure(source, &plan);
    }
}

#[test]
fn paragraph_patch_changed_paragraph_mark_inheritance_requests_restyle_replace() {
    let source = "<w:p><w:pPr><w:rPr><w:rFonts w:ascii='Courier New'/></w:rPr></w:pPr></w:p>";
    let mut plan = patch(Vec::new(), Vec::new());
    seed_inheritance(
        &mut plan,
        json!({}),
        json!([]),
        json!({"fontFamily": {"ascii": "Arial", "hAnsi": "Arial"}}),
    );
    plan.mark_inherit = attrs(json!({
        "fontFamily": {"ascii": "Times New Roman", "hAnsi": "Times New Roman"}
    }));
    assert_restyle_failure(source, &plan);
}

#[test]
fn paragraph_patch_direct_bold_delta_keeps_unchanged_inheritance_and_source_bytes() {
    let source = concat!(
        "<w:p><w:r w:rsidR='12345678'><w:rPr><w:rFonts w:ascii='Courier New'/>",
        "<w:noProof/></w:rPr><w:t>x</w:t></w:r><w:r><w:t>keep</w:t></w:r></w:p>",
    );
    let marks = json!({"fontFamily": {"ascii": "Courier New", "hAnsi": "Arial"}});
    let target = json!({"fontFamily": {"ascii": "Courier New", "hAnsi": "Arial"}, "bold": true});
    let inherited = json!({"fontFamily": {"ascii": "Arial", "hAnsi": "Arial"}});
    let mut plan = patch(
        vec![text("x", marks), text("keep", json!({}))],
        vec![text("x", target.clone()), text("keep", json!({}))],
    );
    seed_inheritance(&mut plan, json!({}), json!([inherited, {}]), json!({}));
    plan.run_edits.push(RunEditV1 {
        from: 0,
        to: 1,
        deltas: vec![delta("bold", None, Some(json!(true)))],
        target: attrs(target),
        inherit: attrs(inherited),
    });
    assert_eq!(
        write(source, &plan).unwrap(),
        source.replace("<w:noProof/>", "<w:b/><w:noProof/>")
    );
}

fn bold_every_run(runs: usize) -> (String, String, ParagraphPatchV1) {
    let source = format!(
        "<w:p>{}</w:p>",
        "<w:r><w:t>abcdefgh</w:t></w:r>".repeat(runs)
    );
    let expected = format!(
        "<w:p>{}</w:p>",
        "<w:r><w:rPr><w:b/></w:rPr><w:t>abcdefgh</w:t></w:r>".repeat(runs)
    );
    let all = "abcdefgh".repeat(runs);
    let mut plan = patch(
        vec![text(&all, json!({}))],
        vec![text(&all, json!({"bold": true}))],
    );
    plan.run_edits.push(RunEditV1 {
        from: 0,
        to: all.len() as u32,
        deltas: vec![delta("bold", None, Some(json!(true)))],
        target: attrs(json!({"bold": true})),
        inherit: Attrs::new(),
    });
    (source, expected, plan)
}

#[test]
fn paragraph_patch_formats_every_run_of_a_paragraph_in_place() {
    let (source, expected, plan) = bold_every_run(3);
    assert_eq!(
        write(&source, &plan).unwrap().as_bytes(),
        expected.as_bytes()
    );
}

fn linear_work_growth(small: u64, large: u64) -> bool {
    assert!(small > 0, "work counter must measure the writer");
    large * 4 <= small * 8 * 5
}

#[test]
fn paragraph_patch_formats_many_short_runs_with_linear_work() {
    let measured = |runs: usize| {
        let (source, expected, plan) = bold_every_run(runs);
        super::work::reset();
        let output = write(&source, &plan).unwrap();
        let work = super::work::get();
        assert_eq!(output, expected);
        work
    };
    let (small, large) = (measured(400), measured(3200));
    assert!(
        linear_work_growth(small, large),
        "8x the runs exceeded 10x the work: {small} -> {large}"
    );
}

#[test]
fn paragraph_patch_linear_work_check_rejects_quadratic_scans() {
    let measured = |runs: usize| {
        super::work::reset();
        for _ in 0..runs {
            for _ in 0..runs {
                super::work::add(1);
            }
        }
        super::work::get()
    };
    let (small, large) = (measured(40), measured(320));
    assert_eq!((small, large), (1600, 102400));
    assert!(!linear_work_growth(small, large));
}

fn model_run(text: &str) -> String {
    use crate::inline::{Run, RunContent, RunType};
    let mut content = Vec::new();
    let mut plain = String::new();
    let flush = |plain: &mut String, content: &mut Vec<RunContent>| {
        if !plain.is_empty() {
            content.push(RunContent::Text {
                text: std::mem::take(plain),
                preserve_space: None,
            });
        }
    };
    for ch in text.chars() {
        let special = match ch {
            '\t' => RunContent::Tab,
            '\u{ad}' => RunContent::SoftHyphen,
            '\u{2011}' => RunContent::NoBreakHyphen,
            _ => {
                plain.push(ch);
                continue;
            }
        };
        flush(&mut plain, &mut content);
        content.push(special);
    }
    flush(&mut plain, &mut content);
    let run = Run {
        node_type: RunType::Run,
        formatting: None,
        property_changes: None,
        content,
    };
    let mut context =
        crate::serializer::SerializerContext::new(&crate::serializer::s10::SerializerDeterminism {
            seed: "0".repeat(64),
            now: "2000-01-01T00:00:00.000Z".to_owned(),
        })
        .unwrap();
    crate::serializer::serialize_run(&run, &mut context).unwrap()
}

#[test]
fn paragraph_patch_writes_special_characters_as_the_model_writer_does() {
    for special in ["\t", "\u{ad}", "\u{2011}"] {
        for (source_text, from, to, insert) in [
            ("ab", 1, 1, special.to_owned()),
            ("ab", 0, 2, special.to_owned()),
            ("ab", 2, 2, format!("x{special}{special}y")),
            ("ab", 0, 0, format!("{special} c")),
            ("ab", 0, 2, format!("x{special}\u{a0}y")),
            ("ab", 1, 1, "\u{a0}".to_owned()),
        ] {
            let source = format!("<w:p><w:r><w:t>{source_text}</w:t></w:r></w:p>");
            let (from, to) = (from as usize, to as usize);
            let result = format!("{}{insert}{}", &source_text[..from], &source_text[to..]);
            let mut plan = patch(
                vec![text(source_text, json!({}))],
                vec![text(&result, json!({}))],
            );
            replace(&mut plan, from as u32, to as u32, &insert, json!({}));
            if from == source_text.len() {
                plan.text_edits[0].donor = DonorV1::At {
                    at: from as u32 - 1,
                };
            }
            assert_eq!(
                write(&source, &plan).unwrap(),
                format!("<w:p>{}</w:p>", model_run(&result)),
                "{result:?}"
            );
        }
        let source = "<w:p><w:r><w:t/></w:r></w:p>";
        let inserted = format!("a{special}b");
        let mut plan = patch(Vec::new(), vec![text(&inserted, json!({}))]);
        replace(&mut plan, 0, 0, &inserted, json!({}));
        plan.text_edits[0].donor = DonorV1::Pilcrow;
        assert_eq!(
            write(source, &plan).unwrap(),
            format!("<w:p>{}</w:p>", model_run(&inserted)),
            "{inserted:?}"
        );
    }
}

#[test]
fn paragraph_patch_refuses_inserted_breaks_and_symbol_scalars() {
    for insert in ["a\nb", "a\u{b}b", "\u{f04a}"] {
        let source = "<w:p><w:r><w:t>ab</w:t></w:r></w:p>";
        let result = format!("a{insert}b");
        let mut plan = patch(vec![text("ab", json!({}))], vec![text(&result, json!({}))]);
        replace(&mut plan, 1, 1, insert, json!({}));
        assert!(write(source, &plan).is_err(), "{insert:?}");
    }
}
