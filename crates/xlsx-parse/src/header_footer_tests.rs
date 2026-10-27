use crate::{HeaderFooterPair, HeaderFooterText, parse_header_footer, parse_header_footer_text};

#[test]
fn header_footer_text_matches_excel_sections() {
    let cases = [
        (r#"&L&"Arial,Bold"&14Left &A&C&KFF0000Center &P of &N&RRight &F"#, "Left &[Tab]", "Center &[Page] of &[Pages]", "Right &[File]"),
        ("&LTotal 50&&&C&1#Label&RTail&12 2024 x", "Total 50&", "#Label", "Tail 2024 x"),
        ("&CBeta&LLeftOne&RRightOne", "LeftOne", "Beta", "RightOne"),
        ("NoCode text&E&S&X&Y&U&B&I&O&HStyled_x000D_Next line", "", "NoCode textStyled\nNext line", ""),
        ("Bare header &P", "", "Bare header &[Page]", ""),
        ("&L&D &T&C&Z&F&R&G", "&[Date] &[Time]", "&[Path]&[File]", "&[Picture]"),
        ("&K01+033Theme", "", "Theme", ""),
        ("&C&\"-,Bold\"&16TITLE\r\nSecond", "", "TITLE\nSecond", ""),
        ("&Qodd", "", "&Qodd", ""),
        ("&Lone&Cmiddle&Ltwo", "onetwo", "middle", ""),
        ("A\rB\r\nC\nD_x000A_E", "", "A\nB\nC\nD\nE", ""),
        ("_x005F_x000D_", "", "_x000D_", ""),
    ];
    for (source, left, center, right) in cases {
        let actual = parse_header_footer_text(source);
        assert_eq!((&*actual.left, &*actual.center, &*actual.right), (left, center, right), "{source:?}");
        assert_eq!(actual.raw, source.replace("_x005F_x000D_", "_x000D_").replace("Styled_x000D_", "Styled\r").replace("_x000A_", "\n"));
    }
}

fn pair(text: &str) -> HeaderFooterPair {
    HeaderFooterPair {
        header: Some(HeaderFooterText {
            raw: text.to_owned(), center: text.to_owned(), ..Default::default()
        }),
        footer: None,
    }
}

#[test]
fn header_footer_element_defaults_and_inactive_variants() {
    let actual = parse_header_footer(br#"<headerFooter><oddHeader>Odd</oddHeader><evenHeader>Ignored</evenHeader><firstFooter>Ignored</firstFooter></headerFooter>"#).unwrap();
    assert!(!actual.different_first);
    assert!(!actual.different_odd_even);
    assert!(actual.scale_with_doc);
    assert!(actual.align_with_margins);
    assert_eq!(actual.odd, pair("Odd"));
    assert_eq!(actual.even, None);
    assert_eq!(actual.first, None);
}

#[test]
fn header_footer_element_boolean_spellings_and_empty_parts() {
    for (yes, no) in [("1", "0"), ("true", "false")] {
        let xml = format!(r#"<headerFooter differentFirst="{yes}" differentOddEven="{yes}" scaleWithDoc="{no}" alignWithMargins="{no}"><oddHeader/><oddFooter></oddFooter><evenHeader>Even</evenHeader><firstHeader>First</firstHeader></headerFooter>"#);
        let actual = parse_header_footer(xml.as_bytes()).unwrap();
        assert!(actual.different_first && actual.different_odd_even);
        assert!(!actual.scale_with_doc && !actual.align_with_margins);
        assert_eq!(actual.odd, HeaderFooterPair::default());
        assert_eq!(actual.even, Some(pair("Even")));
        assert_eq!(actual.first, Some(pair("First")));
        let xml = format!(r#"<headerFooter differentFirst="{no}" differentOddEven="{no}" scaleWithDoc="{yes}" alignWithMargins="{yes}"><oddHeader>A &amp; B</oddHeader></headerFooter>"#);
        let actual = parse_header_footer(xml.as_bytes()).unwrap();
        assert!(!actual.different_first && !actual.different_odd_even);
        assert!(actual.scale_with_doc && actual.align_with_margins);
        assert_eq!(actual.odd, pair("A & B"));
    }
}

#[test]
fn header_footer_element_all_empty_is_absent() {
    for xml in ["<headerFooter/>", "<headerFooter><oddHeader></oddHeader></headerFooter>", "<headerFooter differentFirst=\"true\" differentOddEven=\"true\"><oddHeader/><oddFooter/><evenHeader/><evenFooter/><firstHeader/><firstFooter/></headerFooter>"] {
        assert_eq!(parse_header_footer(xml.as_bytes()), None);
    }
}
