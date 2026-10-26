use docx_parse::xml::{ParseBudget, ParseLimits, parse_xml};
use docx_parse::{
    FontFamily, RunProjection, parse_run, parse_style_definitions, parse_styles, parse_theme,
};

#[test]
fn direct_run_fonts_remain_authored_while_the_projection_resolves_theme_fonts() {
    let limits = ParseLimits::default();
    let mut budget = ParseBudget::new(&limits);
    let theme = parse_theme(
        Some(br#"<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:themeElements><a:fontScheme name="Synthetic"><a:minorFont><a:latin typeface="Theme Sans"/></a:minorFont></a:fontScheme></a:themeElements></a:theme>"#),
        "word/theme/theme1.xml",
        &mut budget,
    )
    .unwrap();
    let xml = parse_xml(
        br#"<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:rPr><w:rFonts w:asciiTheme="minorHAnsi" w:hAnsiTheme="minorHAnsi" w:cstheme="minorHAnsi"/><w:lang w:eastAsia="en-GB"/></w:rPr><w:t>Sample</w:t></w:r>"#,
        "word/document.xml",
        &mut budget,
    )
    .unwrap();
    let projection = parse_run(xml.root().unwrap(), Some(&theme), None, None);
    let direct = projection.run.formatting.unwrap();
    let fonts = direct.font_family.unwrap();
    assert_eq!(fonts.ascii, None);
    assert_eq!(fonts.h_ansi, None);
    assert_eq!(fonts.east_asia, None);
    assert_eq!(fonts.cs, None);
    assert_eq!(fonts.east_asia_theme, None);
    assert_eq!(fonts.ascii_theme.as_deref(), Some("minorHAnsi"));
    assert_eq!(fonts.h_ansi_theme.as_deref(), Some("minorHAnsi"));
    assert_eq!(fonts.cs_theme.as_deref(), Some("minorHAnsi"));
    assert_eq!(direct.language.unwrap().east_asia.as_deref(), Some("en-GB"));
    let resolved = projection.resolved_formatting.unwrap().font_family.unwrap();
    assert_eq!(resolved.ascii.as_deref(), Some("Theme Sans"));
    assert_eq!(resolved.h_ansi.as_deref(), Some("Theme Sans"));
    assert_eq!(resolved.cs.as_deref(), Some("Theme Sans"));
}

fn parse_inherited_fonts(
    default_fonts: &str,
    character_fonts: &str,
    run_fonts: &str,
) -> RunProjection {
    let limits = ParseLimits::default();
    let mut budget = ParseBudget::new(&limits);
    let theme = parse_theme(
        Some(br#"<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:themeElements><a:fontScheme name="Synthetic"><a:majorFont><a:latin typeface="Theme Serif"/></a:majorFont><a:minorFont><a:latin typeface="Theme Sans"/></a:minorFont></a:fontScheme></a:themeElements></a:theme>"#),
        "word/theme/theme1.xml",
        &mut budget,
    )
    .unwrap();
    let styles_xml = format!(
        r#"<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts {default_fonts}/></w:rPr></w:rPrDefault></w:docDefaults><w:style w:type="character" w:styleId="CharacterFont"><w:rPr><w:rFonts {character_fonts}/></w:rPr></w:style></w:styles>"#,
    );
    let styles = parse_styles(
        styles_xml.as_bytes(),
        Some(&theme),
        "word/styles.xml",
        &mut budget,
    )
    .unwrap();
    let definitions = parse_style_definitions(
        styles_xml.as_bytes(),
        Some(&theme),
        "word/styles.xml",
        &mut budget,
    )
    .unwrap();
    let run_xml = format!(
        r#"<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:rPr><w:rStyle w:val="CharacterFont"/><w:rFonts {run_fonts}/></w:rPr><w:t>Sample</w:t></w:r>"#,
    );
    let xml = parse_xml(run_xml.as_bytes(), "word/document.xml", &mut budget).unwrap();
    parse_run(
        xml.root().unwrap(),
        Some(&theme),
        Some(&styles),
        Some(definitions.doc_defaults.as_ref().unwrap()),
    )
}

fn assert_theme_override(projection: RunProjection) {
    let authored = projection.run.formatting.unwrap().font_family.unwrap();
    assert_eq!(
        authored,
        FontFamily {
            ascii_theme: Some("majorHAnsi".to_owned()),
            h_ansi_theme: Some("majorHAnsi".to_owned()),
            east_asia_theme: Some("majorHAnsi".to_owned()),
            cs_theme: Some("majorHAnsi".to_owned()),
            ..FontFamily::default()
        }
    );
    let resolved = projection.resolved_formatting.unwrap().font_family.unwrap();
    for font in [
        resolved.ascii,
        resolved.h_ansi,
        resolved.east_asia,
        resolved.cs,
    ] {
        assert_eq!(font.as_deref(), Some("Theme Serif"));
    }
}

const EXPLICIT_INHERITED_FONTS: &str = r#"w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:eastAsia="Times New Roman" w:cs="Times New Roman""#;
const MINOR_THEME_FONTS: &str = r#"w:asciiTheme="minorHAnsi" w:hAnsiTheme="minorHAnsi" w:eastAsiaTheme="minorHAnsi" w:cstheme="minorHAnsi""#;
const MAJOR_THEME_FONTS: &str = r#"w:asciiTheme="majorHAnsi" w:hAnsiTheme="majorHAnsi" w:eastAsiaTheme="majorHAnsi" w:cstheme="majorHAnsi""#;

#[test]
fn direct_theme_fonts_override_explicit_doc_defaults() {
    assert_theme_override(parse_inherited_fonts(
        EXPLICIT_INHERITED_FONTS,
        "",
        MAJOR_THEME_FONTS,
    ));
}

#[test]
fn direct_theme_fonts_override_theme_doc_defaults() {
    assert_theme_override(parse_inherited_fonts(
        MINOR_THEME_FONTS,
        "",
        MAJOR_THEME_FONTS,
    ));
}

#[test]
fn direct_theme_fonts_override_explicit_character_style() {
    assert_theme_override(parse_inherited_fonts(
        MINOR_THEME_FONTS,
        EXPLICIT_INHERITED_FONTS,
        MAJOR_THEME_FONTS,
    ));
}

#[test]
fn direct_explicit_fonts_override_inherited_theme_fonts() {
    let projection = parse_inherited_fonts(
        MINOR_THEME_FONTS,
        "",
        r#"w:ascii="Arial" w:hAnsi="Arial" w:eastAsia="Arial" w:cs="Arial""#,
    );
    let authored = projection.run.formatting.unwrap().font_family.unwrap();
    assert_eq!(
        authored,
        FontFamily {
            ascii: Some("Arial".to_owned()),
            h_ansi: Some("Arial".to_owned()),
            east_asia: Some("Arial".to_owned()),
            cs: Some("Arial".to_owned()),
            ..FontFamily::default()
        }
    );
    let resolved = projection.resolved_formatting.unwrap().font_family.unwrap();
    for font in [
        resolved.ascii,
        resolved.h_ansi,
        resolved.east_asia,
        resolved.cs,
    ] {
        assert_eq!(font.as_deref(), Some("Arial"));
    }
}
