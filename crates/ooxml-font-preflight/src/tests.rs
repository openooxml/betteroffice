use super::*;

fn zip(parts: &[(&str, &[u8])], deflate: bool) -> Vec<u8> {
    let mut output = Vec::new();
    let mut central = Vec::new();
    let method = u16::from(deflate) * 8;
    for (name, data) in parts {
        let offset = output.len() as u32;
        let crc = zip::crc32(data);
        let compressed = if deflate {
            let mut compressed = Vec::new();
            let count = data.len().max(1).div_ceil(65535);
            for index in 0..count {
                let start = index * 65535;
                let end = (start + 65535).min(data.len());
                let len = (end - start) as u16;
                compressed.push(u8::from(index + 1 == count));
                compressed.extend_from_slice(&len.to_le_bytes());
                compressed.extend_from_slice(&(!len).to_le_bytes());
                compressed.extend_from_slice(&data[start..end]);
            }
            compressed
        } else {
            data.to_vec()
        };
        output.extend_from_slice(b"PK\x03\x04");
        for value in [20u16, 0, method, 0, 0] {
            output.extend_from_slice(&value.to_le_bytes());
        }
        for value in [crc, compressed.len() as u32, data.len() as u32] {
            output.extend_from_slice(&value.to_le_bytes());
        }
        output.extend_from_slice(&(name.len() as u16).to_le_bytes());
        output.extend_from_slice(&0u16.to_le_bytes());
        output.extend_from_slice(name.as_bytes());
        output.extend_from_slice(&compressed);
        central.extend_from_slice(b"PK\x01\x02");
        for value in [20u16, 20, 0, method, 0, 0] {
            central.extend_from_slice(&value.to_le_bytes());
        }
        for value in [crc, compressed.len() as u32, data.len() as u32] {
            central.extend_from_slice(&value.to_le_bytes());
        }
        for value in [name.len() as u16, 0, 0, 0, 0] {
            central.extend_from_slice(&value.to_le_bytes());
        }
        central.extend_from_slice(&0u32.to_le_bytes());
        central.extend_from_slice(&offset.to_le_bytes());
        central.extend_from_slice(name.as_bytes());
    }
    let offset = output.len() as u32;
    output.extend_from_slice(&central);
    output.extend_from_slice(b"PK\x05\x06");
    for value in [0u16, 0, parts.len() as u16, parts.len() as u16] {
        output.extend_from_slice(&value.to_le_bytes());
    }
    output.extend_from_slice(&(central.len() as u32).to_le_bytes());
    output.extend_from_slice(&offset.to_le_bytes());
    output.extend_from_slice(&0u16.to_le_bytes());
    output
}

fn document(body: &str) -> String {
    format!(
        r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body>{body}</w:body></w:document>"#
    )
}

fn styles(body: &str) -> String {
    format!(
        r#"<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Base" w:hAnsi="Base"/></w:rPr></w:rPrDefault></w:docDefaults>{body}</w:styles>"#
    )
}

fn fixture(style: &str, body: &str, extra: &[(&str, &str)], deflate: bool) -> Vec<u8> {
    let document = document(body);
    let styles = styles(style);
    let mut parts = vec![
        ("word/document.xml", document.as_bytes()),
        ("word/styles.xml", styles.as_bytes()),
    ];
    parts.extend(extra.iter().map(|(name, xml)| (*name, xml.as_bytes())));
    zip(&parts, deflate)
}

fn scan(style: &str, body: &str) -> FontPreflight {
    docx_font_preflight(&fixture(style, body, &[], false)).unwrap()
}

fn keys(result: &FontPreflight) -> Vec<(&str, bool, bool)> {
    result
        .faces
        .iter()
        .map(|face| (face.family.as_str(), face.bold, face.italic))
        .collect()
}

#[test]
fn doc_defaults_stored_and_deflated() {
    for deflate in [false, true] {
        let result = docx_font_preflight(&fixture(
            "",
            "<w:p><w:r><w:t>Text</w:t></w:r></w:p>",
            &[],
            deflate,
        ))
        .unwrap();
        assert!(result.complete);
        assert_eq!(keys(&result), [("Base", false, false)]);
    }
}

#[test]
fn based_on_chain_and_unused_styles() {
    let result = scan(
        r#"<w:style w:type="paragraph" w:styleId="Parent"><w:rPr><w:rFonts w:ascii="Inherited" w:hAnsi="Inherited"/><w:b/></w:rPr></w:style><w:style w:type="paragraph" w:styleId="Child"><w:basedOn w:val="Parent"/><w:rPr><w:i/></w:rPr></w:style><w:style w:styleId="Unused"><w:rPr><w:rFonts w:ascii="Unused"/></w:rPr></w:style>"#,
        r#"<w:p><w:pPr><w:pStyle w:val="Child"/></w:pPr><w:r><w:t>Text</w:t></w:r></w:p>"#,
    );
    assert!(result.complete);
    assert_eq!(
        keys(&result),
        [("Inherited", false, false), ("Inherited", true, true)]
    );
}

#[test]
fn run_styles_direct_faces_and_style_toggles() {
    let result = scan(
        r#"<w:style w:type="character" w:styleId="RunBase"><w:rPr><w:rFonts w:ascii="Run" w:hAnsi="Run"/><w:b/></w:rPr></w:style><w:style w:type="character" w:styleId="RunChild"><w:basedOn w:val="RunBase"/></w:style>"#,
        r#"<w:p><w:r><w:rPr><w:rStyle w:val="RunChild"/></w:rPr><w:t>A</w:t></w:r><w:r><w:rPr><w:rStyle w:val="RunChild"/><w:b w:val="0"/><w:i/></w:rPr><w:t>B</w:t></w:r><w:r><w:rPr><w:b/></w:rPr><w:t>C</w:t></w:r></w:p>"#,
    );
    assert_eq!(
        keys(&result),
        [
            ("Base", false, false),
            ("Base", true, false),
            ("Run", false, false),
            ("Run", false, true),
            ("Run", true, false)
        ]
    );
    let toggled = scan(
        r#"<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:rPr><w:b/></w:rPr></w:style><w:style w:type="character" w:styleId="Toggle"><w:rPr><w:b/></w:rPr></w:style>"#,
        r#"<w:p><w:r><w:rPr><w:rStyle w:val="Toggle"/></w:rPr><w:t>A</w:t></w:r></w:p>"#,
    );
    assert_eq!(keys(&toggled), [("Base", false, false)]);
}

const THEME: &str = r#"<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:themeElements><a:fontScheme><a:majorFont><a:latin typeface="Heading"/><a:ea typeface="Major CJK"/><a:cs typeface="Major CS"/></a:majorFont><a:minorFont><a:latin typeface="Body"/><a:ea typeface=""/><a:cs typeface=""/><a:font script="Jpan" typeface="Japanese"/><a:font script="Arab" typeface="Arabic"/></a:minorFont></a:fontScheme></a:themeElements></a:theme>"#;
const THEME_RELS: &str = r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="theme" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="./theme/custom.xml"/></Relationships>"#;

#[test]
fn canonical_theme_is_complete() {
    let rels = THEME_RELS.replace("./theme/custom.xml", "theme/theme1.xml");
    let result = docx_font_preflight(&fixture("",
        r#"<w:p><w:r><w:rPr><w:rFonts w:asciiTheme="minorHAnsi" w:hAnsiTheme="minorHAnsi" w:eastAsiaTheme="majorEastAsia" w:cstheme="majorBidi"/></w:rPr><w:t>漢字 العربية</w:t></w:r></w:p>"#,
        &[("word/_rels/document.xml.rels", &rels), ("word/theme/theme1.xml", THEME)], true,
    )).unwrap();
    assert!(result.complete);
    assert_eq!(
        keys(&result),
        [
            ("Base", false, false),
            ("Body", false, false),
            ("Major CJK", false, false),
            ("Major CS", false, false)
        ]
    );
}

#[test]
fn theme_references_and_explicit_overrides() {
    let result = docx_font_preflight(&fixture(
        r#"<w:style w:type="paragraph" w:styleId="Theme"><w:rPr><w:rFonts w:asciiTheme="majorHAnsi" w:hAnsiTheme="majorHAnsi"/></w:rPr></w:style>"#,
        r#"<w:p><w:pPr><w:pStyle w:val="Theme"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Explicit" w:asciiTheme="minorHAnsi" w:hAnsiTheme="minorHAnsi" w:eastAsiaTheme="majorEastAsia" w:cstheme="majorBidi"/></w:rPr><w:t>漢字 العربية</w:t></w:r></w:p>"#,
        &[("word/_rels/document.xml.rels", THEME_RELS), ("word/theme/custom.xml", THEME)], false,
    )).unwrap();
    assert!(!result.complete);
    assert_eq!(
        keys(&result),
        [
            ("Body", false, false),
            ("Explicit", false, false),
            ("Heading", false, false),
            ("Major CJK", false, false),
            ("Major CS", false, false)
        ]
    );
}

#[test]
fn theme_language_fills_empty_slots() {
    let result = docx_font_preflight(&fixture("",
        r#"<w:p><w:r><w:rPr><w:rFonts w:eastAsiaTheme="minorEastAsia" w:cstheme="minorBidi"/></w:rPr><w:t>かな العربية</w:t></w:r></w:p>"#,
        &[("word/_rels/document.xml.rels", THEME_RELS), ("word/theme/custom.xml", THEME),
            ("word/settings.xml", r#"<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:themeFontLang w:eastAsia="ja-JP" w:bidi="ar-SA"/></w:settings>"#)], false,
    )).unwrap();
    assert_eq!(
        keys(&result),
        [
            ("Arabic", false, false),
            ("Base", false, false),
            ("Japanese", false, false)
        ]
    );
}

#[test]
fn east_asia_requires_text_and_hint_respects_combining_marks() {
    for (text, hint, expected) in [
        ("Latin", "", false),
        ("漢字", "", true),
        ("é", "", false),
        ("é\u{301}", "eastAsia", true),
        ("A\u{301}", "eastAsia", false),
    ] {
        let body = format!(
            r#"<w:p><w:r><w:rPr><w:rFonts w:eastAsia="CJK" w:hint="{hint}"/></w:rPr><w:t>{text}</w:t></w:r></w:p>"#
        );
        let result = scan("", &body);
        assert_eq!(
            result.faces.iter().any(|face| face.family == "CJK"),
            expected,
            "{text}"
        );
    }
}

#[test]
fn complex_slot_styles_are_per_run() {
    let result = scan(
        "",
        r#"<w:p><w:r><w:rPr><w:rFonts w:cs="Complex"/><w:b/><w:iCs/></w:rPr><w:t>العربية</w:t></w:r><w:r><w:rPr><w:rFonts w:cs="Unused"/></w:rPr><w:t>Latin</w:t></w:r></w:p>"#,
    );
    assert_eq!(
        keys(&result),
        [
            ("Base", false, false),
            ("Base", true, false),
            ("Complex", false, false),
            ("Complex", true, false),
            ("Complex", true, true)
        ]
    );
    assert!(result.faces.iter().all(|face| face.scripts == ["arabic"]));
    assert!(!result.complete);
    let forced = scan(
        "",
        r#"<w:p><w:r><w:rPr><w:cs/><w:bCs/><w:rFonts w:cs="Forced"/></w:rPr><w:t>ASCII</w:t></w:r></w:p>"#,
    );
    assert_eq!(
        keys(&forced),
        [
            ("Base", false, false),
            ("Forced", false, false),
            ("Forced", true, false)
        ]
    );
}

#[test]
fn paragraph_script_keys_and_order() {
    let result = scan(
        "",
        "<w:p><w:r><w:t>漢かな한 العربية אב</w:t></w:r></w:p><w:p><w:r><w:t>漢</w:t></w:r></w:p>",
    );
    assert_eq!(
        result.faces[0].scripts,
        ["cjk-jp", "cjk-kr", "arabic", "hebrew", "cjk-sc"]
    );
}

#[test]
fn normal_without_default_and_paragraph_mark_fonts() {
    let result = scan(
        r#"<w:style w:type="paragraph" w:styleId="Normal"><w:rPr><w:rFonts w:ascii="Normal" w:hAnsi="Normal"/></w:rPr></w:style>"#,
        "<w:p><w:r><w:t>Normal</w:t></w:r></w:p>",
    );
    assert_eq!(keys(&result), [("Normal", false, false)]);
    let result = scan(
        "",
        r#"<w:p><w:pPr><w:rPr><w:rFonts w:ascii="Mark" w:hAnsi="Mark"/></w:rPr></w:pPr><w:r><w:t>Base</w:t></w:r></w:p>"#,
    );
    assert_eq!(
        keys(&result),
        [("Base", false, false), ("Mark", false, false)]
    );
}

#[test]
fn math_names_the_engine_fallback_and_inherited_slots() {
    let result = scan(
        "",
        r#"<w:p><m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"><m:r><m:t>x=1</m:t></m:r></m:oMath></w:p>"#,
    );
    assert_eq!(
        keys(&result),
        [
            ("Base", false, false),
            ("Base", false, true),
            ("Cambria Math", false, false),
            ("Cambria Math", false, true)
        ]
    );
}

#[test]
fn normal_notes_and_separators_match_the_editor_request() {
    let result = docx_font_preflight(&fixture("",
        r#"<w:p><w:r><w:t>Body</w:t><w:footnoteReference w:id="1"/></w:r></w:p>"#,
        &[("word/_rels/document.xml.rels", r#"<Relationships><Relationship Id="notes" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/></Relationships>"#),
            ("word/footnotes.xml", r#"<w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:footnote w:id="-1" w:type="separator"><w:p/></w:footnote><w:footnote w:id="1"><w:p><w:r><w:rPr><w:rFonts w:ascii="Note"/></w:rPr><w:t>Note</w:t></w:r></w:p></w:footnote><w:footnote w:id="2"><w:p><w:r><w:rPr><w:rFonts w:ascii="Unused"/></w:rPr><w:t>Unused</w:t></w:r></w:p></w:footnote></w:footnotes>"#)], true,
    )).unwrap();
    assert!(result.complete);
    assert_eq!(
        keys(&result),
        [
            ("Base", false, false),
            ("Note", false, false),
            ("Unused", false, false)
        ]
    );
}

#[test]
fn header_relationship_and_unused_header() {
    let result = docx_font_preflight(&fixture("",
        r#"<w:p/><w:sectPr><w:headerReference w:type="default" r:id="header"/></w:sectPr>"#,
        &[("word/_rels/document.xml.rels", r#"<Relationships><Relationship Id="header" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="../word/header.xml"/><Relationship Id="unused" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="unused.xml"/></Relationships>"#),
            ("word/header.xml", r#"<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:r><w:rPr><w:rFonts w:ascii="Header"/></w:rPr><w:t>Header</w:t></w:r></w:p></w:hdr>"#),
            ("word/unused.xml", "broken XML")], true,
    )).unwrap();
    assert!(result.complete);
    assert_eq!(
        keys(&result),
        [("Base", false, false), ("Header", false, false)]
    );
}

#[test]
fn numbering_reads_only_used_levels_and_instances() {
    let result = docx_font_preflight(&fixture("",
        r#"<w:p><w:pPr><w:numPr><w:numId w:val="1"/><w:ilvl w:val="0"/></w:numPr></w:pPr><w:r><w:t>List</w:t></w:r></w:p>"#,
        &[("word/numbering.xml", r#"<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:numFmt w:val="bullet"/><w:lvlText w:val="•"/><w:rPr><w:rFonts w:ascii="Bullet"/><w:b/></w:rPr></w:lvl><w:lvl w:ilvl="1"><w:lvlText w:val="•"/><w:rPr><w:rFonts w:ascii="Unused"/></w:rPr></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>"#)], false,
    )).unwrap();
    assert!(result.complete);
    assert_eq!(
        keys(&result),
        [("Base", false, false), ("Bullet", true, false)]
    );
    assert!(
        docx_font_preflight(&fixture(
            "",
            "<w:p/>",
            &[("word/numbering.xml", "broken")],
            false
        ))
        .unwrap()
        .complete
    );
}

#[test]
fn table_conditionals_apply_only_to_active_cells() {
    let result = scan(
        r#"<w:style w:type="table" w:styleId="Table"><w:rPr><w:rFonts w:ascii="Table" w:hAnsi="Table"/></w:rPr><w:tblStylePr w:type="firstRow"><w:rPr><w:b/></w:rPr></w:tblStylePr><w:tblStylePr w:type="lastRow"><w:rPr><w:rFonts w:ascii="Unused"/></w:rPr></w:tblStylePr></w:style>"#,
        r#"<w:tbl><w:tblPr><w:tblStyle w:val="Table"/></w:tblPr><w:tr><w:tc><w:p><w:r><w:t>Cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl>"#,
    );
    assert!(result.complete);
    assert_eq!(
        keys(&result),
        [
            ("Calibri", false, false),
            ("Table", false, false),
            ("Table", true, false)
        ]
    );
}

#[test]
fn marker_faces_inherit_paragraph_bold_and_italic() {
    for (level_flags, expected_bold) in [("", true), (r#"<w:b w:val="0"/>"#, false)] {
        let numbering = format!(
            r#"<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:numFmt w:val="bullet"/><w:lvlText w:val="•"/><w:rPr><w:rFonts w:ascii="Bullet"/>{level_flags}</w:rPr></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>"#
        );
        let result = docx_font_preflight(&fixture(
            r#"<w:style w:type="paragraph" w:styleId="BoldItalic"><w:rPr><w:b/><w:i/></w:rPr></w:style>"#,
            r#"<w:p><w:pPr><w:pStyle w:val="BoldItalic"/><w:numPr><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>List</w:t></w:r></w:p>"#,
            &[("word/numbering.xml", &numbering)], false,
        )).unwrap();
        let faces: Vec<_> = result
            .faces
            .iter()
            .filter(|face| face.family == "Bullet")
            .collect();
        assert_eq!(faces.len(), 1);
        assert_eq!((faces[0].bold, faces[0].italic), (expected_bold, true));
    }
}

#[test]
fn text_boxes_use_direct_fonts_and_the_shape_default() {
    let result = scan(
        r#"<w:style w:type="paragraph" w:styleId="BoxStyle"><w:rPr><w:rFonts w:ascii="Unused" w:hAnsi="Unused"/><w:b/></w:rPr></w:style>"#,
        r#"<w:p><w:r><w:pict><v:shape xmlns:v="urn:synthetic"><v:textbox><w:txbxContent><w:p><w:pPr><w:pStyle w:val="BoxStyle"/></w:pPr><w:r><w:rPr><w:b/></w:rPr><w:t>Box</w:t></w:r></w:p></w:txbxContent></v:textbox></v:shape></w:pict></w:r></w:p>"#,
    );
    assert_eq!(
        keys(&result),
        [
            ("Base", false, false),
            ("Calibri", false, false),
            ("Calibri", true, false)
        ]
    );
}

#[test]
fn table_cells_inside_controls_keep_their_runs() {
    let result = scan(
        "",
        r#"<w:tbl><w:tr><w:sdt><w:sdtContent><w:tc><w:p><w:r><w:rPr><w:rFonts w:ascii="Glyph" w:hAnsi="Glyph"/><w:b/></w:rPr><w:t>Cell</w:t></w:r></w:p></w:tc></w:sdtContent></w:sdt></w:tr></w:tbl>"#,
    );
    assert_eq!(
        keys(&result),
        [
            ("Base", false, false),
            ("Calibri", false, false),
            ("Glyph", false, false),
            ("Glyph", true, false)
        ]
    );
}

#[test]
fn family_deduplication_folds_ascii_and_fullwidth_latin() {
    let result = scan(
        "",
        r#"<w:p><w:r><w:rPr><w:rFonts w:ascii="ＭＳ Gothic" w:hAnsi="ＭＳ Gothic"/></w:rPr><w:t>A</w:t></w:r><w:r><w:rPr><w:rFonts w:ascii="ｍｓ GOTHIC" w:hAnsi="ｍｓ GOTHIC"/></w:rPr><w:t>B</w:t></w:r></w:p>"#,
    );
    assert_eq!(
        keys(&result),
        [("Base", false, false), ("ＭＳ Gothic", false, false)]
    );
}

#[test]
fn namespace_scopes_and_entities_keep_text_slots() {
    let doc = r#"<x:document xmlns:x="http://purl.oclc.org/ooxml/wordprocessingml/main"><x:body><x:p><x:r xmlns:x="urn:ignored"><x:t>ignored</x:t></x:r><x:r><x:rPr><x:rFonts x:ascii="A &amp; B" x:eastAsia="CJK"/></x:rPr><x:t>&#x6f22;&amp;</x:t></x:r></x:p></x:body></x:document>"#;
    let styles = styles("");
    let result = docx_font_preflight(&zip(
        &[
            ("word/document.xml", doc.as_bytes()),
            ("word/styles.xml", styles.as_bytes()),
        ],
        false,
    ))
    .unwrap();
    assert_eq!(
        keys(&result),
        [
            ("A & B", false, false),
            ("Base", false, false),
            ("CJK", false, false)
        ]
    );
    assert!(result.faces.iter().all(|face| face.scripts == ["cjk-sc"]));
}

#[test]
fn incomplete_missing_styles_and_style_cycles() {
    let result = docx_font_preflight(&zip(
        &[("word/document.xml", document("<w:p/>").as_bytes())],
        false,
    ))
    .unwrap();
    assert!(!result.complete);
    assert_eq!(keys(&result), [("Calibri", false, false)]);
    assert!(
        !scan(
            r#"<w:style w:styleId="Loop"><w:basedOn w:val="Loop"/></w:style>"#,
            "<w:p/>"
        )
        .complete
    );
}

#[test]
fn malformed_inputs_return_errors_without_panics() {
    assert!(docx_font_preflight(b"not a zip").is_err());
    let bytes = fixture("", "<w:p/>", &[], false);
    for end in 0..bytes.len() {
        assert!(docx_font_preflight(&bytes[..end]).is_err());
    }
    let mut corrupt = bytes.clone();
    corrupt[30] ^= 1;
    assert!(docx_font_preflight(&corrupt).is_err());
    let mut offset = bytes;
    let len = offset.len();
    offset[len - 6..len - 2].copy_from_slice(&u32::MAX.to_le_bytes());
    assert!(docx_font_preflight(&offset).is_err());
    assert!(
        docx_font_preflight(&zip(&[("word/document.xml", b"<!DOCTYPE x><x/>")], false)).is_err()
    );
    assert!(docx_font_preflight(&zip(&[("word/document.xml", b"<w:document>")], false)).is_err());
    let body = format!(
        r#"<w:tbl><w:tr><w:tc><w:tcPr><w:gridSpan w:val="{}"/></w:tcPr><w:p/></w:tc><w:tc><w:p/></w:tc></w:tr></w:tbl>"#,
        usize::MAX
    );
    assert!(!scan("", &body).faces.is_empty());
}

#[test]
fn json_escapes_families() {
    let result = FontPreflight {
        complete: false,
        faces: vec![PreflightFace {
            family: "A\"\\\n\t\u{0000}é".into(),
            bold: true,
            italic: false,
            scripts: vec!["arabic".into()],
        }],
    };
    let parsed: serde_json::Value = serde_json::from_str(&result.to_json()).unwrap();
    assert_eq!(parsed["faces"][0]["family"], result.faces[0].family);
    assert_eq!(parsed["complete"], false);
}
