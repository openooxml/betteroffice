use docx_parse::xml::{ParseBudget, ParseLimits, parse_xml};
use docx_parse::{parse_run, parse_theme};

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
