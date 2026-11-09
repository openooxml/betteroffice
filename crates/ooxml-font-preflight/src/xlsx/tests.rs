use super::*;
use crate::test_support::zip;

const S: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const R: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const MC: &str = "http://schemas.openxmlformats.org/markup-compatibility/2006";
const STYLES: &str = r#"<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts><font><name val="Base"/></font><font><name val="Styled"/><b/><i/></font><font><name val="Unused"/></font></fonts><cellXfs><xf fontId="0"/><xf fontId="1"/><xf fontId="2"/><xf fontId="1" applyFont="0"/></cellXfs></styleSheet>"#;

fn worksheet(rows: &str, tail: &str) -> String {
    format!(r#"<worksheet xmlns="{S}"><sheetData>{rows}</sheetData>{tail}</worksheet>"#)
}

fn fixture(
    sheet: &str,
    styles: &str,
    shared: Option<&str>,
    extra: &[(&str, &str)],
    deflate: bool,
) -> Vec<u8> {
    let workbook = format!(
        r#"<workbook xmlns="{S}" xmlns:r="{R}"><bookViews><workbookView activeTab="1"/></bookViews><sheets><sheet name="Hidden" state="hidden" sheetId="1" r:id="hidden"/><sheet name="Active" sheetId="2" r:id="active"/><sheet name="Other" sheetId="3" r:id="other"/></sheets></workbook>"#
    );
    let rels = format!(
        r#"<Relationships><Relationship Id="hidden" Type="{R}/worksheet" Target="worksheets/hidden.xml"/><Relationship Id="active" Type="{R}/worksheet" Target="worksheets/active.xml"/><Relationship Id="other" Type="{R}/worksheet" Target="worksheets/other.xml"/><Relationship Id="styles" Type="{R}/styles" Target="styles.xml"/><Relationship Id="shared" Type="{R}/sharedStrings" Target="sharedStrings.xml"/><Relationship Id="theme" Type="{R}/theme" Target="theme/unused.xml"/></Relationships>"#
    );
    let mut parts = vec![
        ("xl/workbook.xml", workbook.as_bytes()),
        ("xl/_rels/workbook.xml.rels", rels.as_bytes()),
        ("xl/worksheets/active.xml", sheet.as_bytes()),
        ("xl/worksheets/hidden.xml", b"invalid".as_slice()),
        ("xl/worksheets/other.xml", b"invalid".as_slice()),
        ("xl/styles.xml", styles.as_bytes()),
        ("xl/theme/unused.xml", b"invalid".as_slice()),
    ];
    if let Some(shared) = shared {
        parts.push(("xl/sharedStrings.xml", shared.as_bytes()));
    }
    parts.extend(extra.iter().map(|(name, xml)| (*name, xml.as_bytes())));
    zip(&parts, deflate)
}

fn keys(result: &FontPreflight) -> Vec<(&str, bool, bool)> {
    result
        .faces
        .iter()
        .map(|face| (face.family.as_str(), face.bold, face.italic))
        .collect()
}

fn alternate(choice: &str, fallback: &str) -> String {
    format!(
        r#"<mc:AlternateContent xmlns:mc="{MC}" xmlns:future="urn:future"><mc:Choice Requires="future">{choice}</mc:Choice><mc:Fallback>{fallback}</mc:Fallback></mc:AlternateContent>"#
    )
}

#[test]
fn alternate_content_cells_follow_worksheet_read_order() {
    let first = r#"<c r="A1" s="2"><v>1</v></c>"#;
    let last = r#"<c r="A1" s="1"><v>2</v></c>"#;
    for wrapped in [
        alternate(first, last),
        alternate("", last),
        alternate(last, ""),
        alternate(first, &alternate("", last)),
    ] {
        let sheet = worksheet(&format!(r#"<row r="1">{wrapped}</row>"#), "");
        let result = xlsx_font_preflight(&fixture(&sheet, STYLES, None, &[], false)).unwrap();
        assert!(result.complete);
        assert_eq!(keys(&result), [("Styled", true, true)]);
    }
    let wrapped = alternate(r#"<c r="B1" s="2"><v>1</v></c>"#, last);
    let sheet = worksheet(&format!(r#"<row r="1">{wrapped}</row>"#), "");
    let result = xlsx_font_preflight(&fixture(&sheet, STYLES, None, &[], false)).unwrap();
    assert!(result.complete);
    assert_eq!(
        keys(&result),
        [("Styled", true, true), ("Unused", false, false)]
    );
}

#[test]
fn alternate_content_default_cell_keeps_earlier_value() {
    let cells = alternate(
        r#"<c r="A1" s="1"><v>1</v></c>"#,
        r#"<c r="A1" t="inlineStr"/>"#,
    );
    let sheet = worksheet(&format!(r#"<row r="1">{cells}</row>"#), "");
    let result = xlsx_font_preflight(&fixture(&sheet, STYLES, None, &[], false)).unwrap();
    assert!(result.complete);
    assert_eq!(keys(&result), [("Styled", true, true)]);
}

#[test]
fn alternate_content_stored_cell_replaces_earlier_value() {
    let cells = alternate(
        r#"<c r="A1" s="1"><v>1</v></c>"#,
        r#"<c r="A1" t="inlineStr"><is><t>Text</t></is></c>"#,
    );
    let sheet = worksheet(&format!(r#"<row r="1">{cells}</row>"#), "");
    let result = xlsx_font_preflight(&fixture(&sheet, STYLES, None, &[], false)).unwrap();
    assert!(result.complete);
    assert_eq!(keys(&result), [("Base", false, false)]);
}

#[test]
fn alternate_content_rows_follow_worksheet_read_order() {
    let rows = alternate(
        r#"<row r="1"><c r="A1" s="2"><v>1</v></c></row>"#,
        r#"<row r="1"><c r="A1" s="1"><v>2</v></c></row>"#,
    );
    let sheet = worksheet(&rows, "");
    let result = xlsx_font_preflight(&fixture(&sheet, STYLES, None, &[], false)).unwrap();
    assert!(result.complete);
    assert_eq!(keys(&result), [("Styled", true, true)]);
    let rows = alternate(
        r#"<row r="1" hidden="1"><c r="A1" s="1"><v>1</v></c></row>"#,
        r#"<row r="1"><c r="A1" s="1"><v>2</v></c></row>"#,
    );
    let sheet = worksheet(&rows, "");
    let result = xlsx_font_preflight(&fixture(&sheet, STYLES, None, &[], false)).unwrap();
    assert!(result.complete);
    assert!(result.faces.is_empty());
}

#[test]
fn alternate_content_cell_values_use_the_last_value() {
    let values = alternate("<v/>", "<v>1</v>");
    let sheet = worksheet(
        &format!(r#"<row r="1"><c r="A1" s="1">{values}</c></row>"#),
        "",
    );
    let result = xlsx_font_preflight(&fixture(&sheet, STYLES, None, &[], false)).unwrap();
    assert!(result.complete);
    assert_eq!(keys(&result), [("Styled", true, true)]);
}

#[test]
fn alternate_content_style_pools_take_the_first_branch() {
    let fonts = r#"<fonts><font><name val="Base"/></font><font><name val="Styled"/><b/><i/></font></fonts>"#;
    let xfs = r#"<cellXfs><xf fontId="0"/><xf fontId="1"/></cellXfs>"#;
    let unused_pools =
        r#"<fonts><font><name val="Unused"/></font></fonts><cellXfs><xf fontId="0"/></cellXfs>"#;
    let wrapped_pools = alternate(&format!("{fonts}{xfs}"), unused_pools);
    let wrapped_fonts = alternate(
        r#"<font><name val="Styled"/><b/><i/></font>"#,
        r#"<font><name val="Unused"/></font>"#,
    );
    let wrapped_xfs = alternate(r#"<xf fontId="1"/>"#, r#"<xf fontId="0"/>"#);
    let wrapped_properties =
        alternate(r#"<name val="Styled"/><b/><i/>"#, r#"<name val="Unused"/>"#);
    let pools = [
        wrapped_pools,
        format!(r#"<fonts><font><name val="Base"/></font>{wrapped_fonts}</fonts>{xfs}"#),
        format!(r#"{fonts}<cellXfs><xf fontId="0"/>{wrapped_xfs}</cellXfs>"#),
        format!(
            r#"<fonts><font><name val="Base"/></font><font>{wrapped_properties}</font></fonts>{xfs}"#
        ),
        format!(
            r#"<fonts><font><name val="Base"/></font><font><name val="Unused"/>{wrapped_properties}</font></fonts>{xfs}"#
        ),
        format!(
            r#"<mc:AlternateContent xmlns:mc="{MC}"><mc:Fallback>{fonts}{xfs}</mc:Fallback><mc:Choice Requires="future">{unused_pools}</mc:Choice></mc:AlternateContent>"#
        ),
    ];
    let sheet = worksheet(r#"<row r="1"><c r="A1" s="1"><v>1</v></c></row>"#, "");
    for pools in pools {
        let styles = format!(r#"<styleSheet xmlns="{S}">{pools}</styleSheet>"#);
        let result = xlsx_font_preflight(&fixture(&sheet, &styles, None, &[], false)).unwrap();
        assert!(result.complete);
        assert_eq!(keys(&result), [("Styled", true, true)]);
    }
    let formats = alternate(
        r#"<numFmt numFmtId="164" formatCode=";;;"/>"#,
        r#"<numFmt numFmtId="164" formatCode="0"/>"#,
    );
    let styles = STYLES
        .replace(
            "<cellXfs>",
            &format!("<numFmts>{formats}</numFmts><cellXfs>"),
        )
        .replace("<xf fontId=\"2\"/>", "<xf fontId=\"2\" numFmtId=\"164\"/>");
    let sheet = worksheet(
        r#"<row r="1"><c r="A1" s="1"><v>1</v></c><c r="B1" s="2"><v>2</v></c></row>"#,
        "",
    );
    let result = xlsx_font_preflight(&fixture(&sheet, &styles, None, &[], false)).unwrap();
    assert!(result.complete);
    assert_eq!(keys(&result), [("Styled", true, true)]);
}

#[test]
fn unhandled_alternate_content_is_incomplete() {
    let columns = alternate("", r#"<col min="1" max="1"/>"#);
    let sheet = worksheet(
        r#"<row r="1"><c r="A1" s="1"><v>1</v></c></row>"#,
        &format!("<cols>{columns}</cols>"),
    );
    let result = xlsx_font_preflight(&fixture(&sheet, STYLES, None, &[], false)).unwrap();
    assert!(!result.complete);
    assert_eq!(keys(&result), [("Styled", true, true)]);
}

#[test]
fn active_visible_sheet_values_and_cell_styles_only() {
    let sheet = worksheet(
        r#"<row r="1"><c r="A1" t="inlineStr"><is><t>Plain</t></is></c><c r="B1" s="1"><v>12</v></c><c r="C1" s="2"/></row>"#,
        "",
    );
    for deflate in [false, true] {
        let result = xlsx_font_preflight(&fixture(&sheet, STYLES, None, &[], deflate)).unwrap();
        assert!(result.complete);
        assert_eq!(
            keys(&result),
            [("Base", false, false), ("Styled", true, true)]
        );
    }
}

#[test]
fn opening_viewport_excludes_far_cells_and_hidden_tracks() {
    let sheet = worksheet(
        r#"<row r="1"><c r="A1"><v>1</v></c><c r="U1" s="2"><v>2</v></c><c r="B1" s="2"><v>3</v></c></row><row r="2" hidden="1"><c r="A2" s="2"><v>4</v></c></row><row r="51"><c r="A51" s="2"><v>5</v></c></row>"#,
        r#"<cols><col min="2" max="2" hidden="1"/></cols>"#,
    );
    let result = xlsx_font_preflight(&fixture(&sheet, STYLES, None, &[], false)).unwrap();
    assert!(result.complete);
    assert_eq!(keys(&result), [("Base", false, false)]);
    assert_eq!((OPENING_VIEWPORT_ROWS, OPENING_VIEWPORT_COLUMNS), (50, 20));
}

#[test]
fn existing_unstyled_cells_use_normal_font_in_styled_rows_and_columns() {
    let sheet = worksheet(
        r#"<row r="1" s="2" customFormat="1"><c r="A1"><v>1</v></c></row>"#,
        r#"<cols><col min="1" max="1" style="2"/></cols>"#,
    );
    let result = xlsx_font_preflight(&fixture(&sheet, STYLES, None, &[], false)).unwrap();
    assert!(result.complete);
    assert_eq!(keys(&result), [("Base", false, false)]);
}

#[test]
fn shared_and_inline_strings_attach_only_used_scripts() {
    let sheet = worksheet(
        r#"<row r="1"><c r="A1" t="s"><v>1</v></c><c r="B1" s="1" t="inlineStr"><is><r><rPr><rFont val="Unused"/></rPr><t>العربية</t></r><r><t>אב</t></r><rPh><t>かな</t></rPh></is></c></row>"#,
        "",
    );
    let strings = format!(
        r#"<sst xmlns="{S}"><si><t>한글</t></si><si><r><t>漢</t></r><r><t>かな</t></r></si></sst>"#
    );
    let result = xlsx_font_preflight(&fixture(&sheet, STYLES, Some(&strings), &[], false)).unwrap();
    assert!(result.complete);
    assert_eq!(result.faces[0].scripts, ["cjk-jp"]);
    assert_eq!(result.faces[1].scripts, ["cjk-jp", "arabic", "hebrew"]);
    assert_eq!(
        keys(&result),
        [("Base", false, false), ("Styled", true, true)]
    );
}

#[test]
fn cached_theme_scheme_name_matches_the_painter() {
    let styles = STYLES.replace(
        "<name val=\"Base\"/>",
        "<name val=\"Cached\"/><scheme val=\"minor\"/>",
    );
    let sheet = worksheet(r#"<row r="1"><c r="A1"><v>1</v></c></row>"#, "");
    let result = xlsx_font_preflight(&fixture(&sheet, &styles, None, &[], false)).unwrap();
    assert!(result.complete);
    assert_eq!(keys(&result), [("Cached", false, false)]);
    let styles = styles.replace("<name val=\"Cached\"/>", "");
    let result = xlsx_font_preflight(&fixture(&sheet, &styles, None, &[], false)).unwrap();
    assert!(!result.complete);
    assert_eq!(keys(&result), [("Calibri", false, false)]);
}

#[test]
fn apply_font_false_and_missing_font_fall_back_without_bold() {
    let sheet = worksheet(r#"<row r="1"><c r="A1" s="3"><v>1</v></c></row>"#, "");
    let result = xlsx_font_preflight(&fixture(&sheet, STYLES, None, &[], false)).unwrap();
    assert!(result.complete);
    assert_eq!(keys(&result), [("Base", false, false)]);
    let styles = STYLES.replace("fontId=\"1\" applyFont=\"0\"", "fontId=\"99\"");
    let result = xlsx_font_preflight(&fixture(&sheet, &styles, None, &[], false)).unwrap();
    assert!(!result.complete);
    assert_eq!(keys(&result), [("Base", false, false)]);
}

#[test]
fn merged_cells_exclude_covered_values() {
    let sheet = worksheet(
        r#"<row r="1"><c r="A1"><v>1</v></c><c r="B1" s="2"><v>2</v></c></row>"#,
        r#"<mergeCells><mergeCell ref="A1:B1"/></mergeCells>"#,
    );
    let result = xlsx_font_preflight(&fixture(&sheet, STYLES, None, &[], false)).unwrap();
    assert!(result.complete);
    assert_eq!(keys(&result), [("Base", false, false)]);
}

#[test]
fn overlapping_merge_ranges_keep_covered_anchors_hidden() {
    let sheet = worksheet(
        r#"<row r="1"><c r="A1"><v>1</v></c><c r="B1" s="1"><v>2</v></c><c r="C1" s="2"><v>3</v></c></row>"#,
        r#"<mergeCells><mergeCell ref="A1:B1"/><mergeCell ref="B1:C1"/><mergeCell ref="C1:C1"/></mergeCells>"#,
    );
    let result = xlsx_font_preflight(&fixture(&sheet, STYLES, None, &[], false)).unwrap();
    assert!(result.complete);
    assert_eq!(keys(&result), [("Base", false, false)]);
}

#[test]
fn hidden_merge_anchors_stay_incomplete_with_single_cell_ranges() {
    let sheet = worksheet(
        r#"<row r="1" hidden="1"><c r="A1" s="2"><v>1</v></c></row><row r="2"><c r="A2"><v>2</v></c></row>"#,
        r#"<mergeCells><mergeCell ref="A1:B1"/><mergeCell ref="A1:A1"/></mergeCells>"#,
    );
    let result = xlsx_font_preflight(&fixture(&sheet, STYLES, None, &[], false)).unwrap();
    assert!(!result.complete);
    assert_eq!(keys(&result), [("Base", false, false)]);
}

#[test]
fn missing_shared_strings_and_unknown_styles_are_incomplete() {
    let sheet = worksheet(
        r#"<row r="1"><c r="A1" s="99"><v>1</v></c><c r="B1" t="s"><v>5</v></c></row>"#,
        "",
    );
    let result = xlsx_font_preflight(&fixture(&sheet, STYLES, None, &[], false)).unwrap();
    assert!(!result.complete);
    assert_eq!(keys(&result), [("Base", false, false)]);
}

#[test]
fn unsupported_opening_paint_rules_are_incomplete() {
    for tail in [
        r#"<sheetViews><sheetView topLeftCell="C9"/></sheetViews>"#,
        r#"<sheetViews><sheetView><pane xSplit="1" ySplit="2"/></sheetView></sheetViews>"#,
        "<conditionalFormatting/>",
        "<hyperlinks/>",
    ] {
        let sheet = worksheet(r#"<row r="1"><c r="A1"><v>1</v></c></row>"#, tail);
        assert!(
            !xlsx_font_preflight(&fixture(&sheet, STYLES, None, &[], false))
                .unwrap()
                .complete
        );
    }
}

#[test]
fn malformed_inputs_and_shared_xml_fail_without_panicking() {
    assert_eq!(
        xlsx_font_preflight(b"invalid").unwrap_err(),
        PreflightError::Zip
    );
    assert_eq!(
        xlsx_font_preflight(&zip(&[], false)).unwrap_err(),
        PreflightError::MissingPart
    );
    let sheet = worksheet(r#"<row r="1"><c r="A1" t="s"><v>0</v></c></row>"#, "");
    let bytes = fixture(&sheet, STYLES, Some("<sst>"), &[], false);
    assert_eq!(
        xlsx_font_preflight(&bytes).unwrap_err(),
        PreflightError::Xml
    );
    for end in 0..bytes.len() {
        assert!(xlsx_font_preflight(&bytes[..end]).is_err());
    }
}

#[test]
fn xml_byte_and_face_budgets_are_bounded() {
    let sheet = worksheet(r#"<row r="1"><c r="A1"><v>1</v></c></row>"#, "");
    let bytes = fixture(&sheet, STYLES, None, &[], false);
    let mut parts = Parts {
        zip: Zip::new(&bytes).unwrap(),
        nodes: 1,
    };
    let mut complete = true;
    assert!(matches!(
        parts.part("xl/workbook.xml", &mut complete),
        Err(PreflightError::Limit)
    ));
    drop(parts);
    let mut oversized = bytes;
    let central = oversized
        .windows(4)
        .position(|bytes| bytes == b"PK\x01\x02")
        .unwrap();
    oversized[central + 24..central + 28].copy_from_slice(&(33u32 * 1024 * 1024).to_le_bytes());
    assert_eq!(
        xlsx_font_preflight(&oversized).unwrap_err(),
        PreflightError::Limit
    );
    let mut faces = Faces::default();
    for index in 0..513 {
        faces.record(&format!("Font{index}"), false, false, &[]);
    }
    let result = faces.finish(true);
    assert!(!result.complete);
    assert_eq!(result.faces.len(), 512);
}

#[test]
fn escaped_text_and_suppressed_values_match_painting() {
    let styles = STYLES
        .replace(
            "<cellXfs>",
            "<numFmts><numFmt numFmtId=\"164\" formatCode=\";;;\"/></numFmts><cellXfs>",
        )
        .replace("<xf fontId=\"2\"/>", "<xf fontId=\"2\" numFmtId=\"164\"/>");
    let sheet = worksheet(
        r#"<row r="1"><c r="A1" t="inlineStr"><is><t>_x6F22_</t></is></c><c r="B1" s="2"><v>1</v></c></row>"#,
        "",
    );
    let result = xlsx_font_preflight(&fixture(&sheet, &styles, None, &[], false)).unwrap();
    assert!(result.complete);
    assert_eq!(keys(&result), [("Base", false, false)]);
    assert_eq!(result.faces[0].scripts, ["cjk-sc"]);
}

#[test]
fn styles_fall_back_to_the_canonical_part_without_a_relationship() {
    let workbook = format!(
        r#"<workbook xmlns="{S}"><sheets><sheet name="Sheet" sheetId="1"/></sheets></workbook>"#
    );
    let sheet = worksheet(r#"<row r="1"><c r="A1"><v>1</v></c></row>"#, "");
    let bytes = zip(
        &[
            ("xl/workbook.xml", workbook.as_bytes()),
            ("xl/worksheets/sheet1.xml", sheet.as_bytes()),
            ("xl/styles.xml", STYLES.as_bytes()),
        ],
        false,
    );
    let result = xlsx_font_preflight(&bytes).unwrap();
    assert!(result.complete);
    assert_eq!(keys(&result), [("Base", false, false)]);
}
