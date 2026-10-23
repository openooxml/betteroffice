use crate::mask::TextMasker;
use crate::styles::StyleMap;
use crate::{Format, RedactionOptions, RedactionReport, xml};

const W: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

fn rewrite(body: &str, random_characters: bool) -> Result<String, crate::RedactError> {
    let styles = format!(
        r#"<w:styles xmlns:w="{W}"><w:style w:type="paragraph" w:customStyle="1" w:styleId="SecretId"><w:name w:val="Secret Style"/></w:style></w:styles>"#
    );
    let styles = StyleMap::collect("word/styles.xml", styles.as_bytes()).unwrap();
    let source = format!(
        r#"<w:document xmlns:w="{W}" xmlns:b="http://schemas.openxmlformats.org/wordprocessingml/2006/mai&#110;" xmlns:x="urn:ignored"><w:body><w:p>{body}</w:p></w:body></w:document>"#
    );
    xml::redact_xml_with_styles(
        Format::Docx,
        "word/document.xml",
        source.as_bytes(),
        &mut RedactionReport::default(),
        &styles,
        &mut TextMasker::new(&RedactionOptions { random_characters }),
    )
    .map(|output| String::from_utf8(output).unwrap())
}

#[test]
fn shared_fields_rewrite_split_styles_with_decoded_word_markers() {
    let body = r#"<w:r><b:fldChar x:fldCharType="end" b:fldCharType="beg&#105;n"/></w:r>
<w:r><w:instrText>STYLEREF &quot;Secret</w:instrText></w:r>
<w:r><w:instrText><![CDATA[ Style]]>&quot;</w:instrText></w:r>
<w:r><b:fldChar b:fldCharType="separ&#97;te"/></w:r>
<w:r><w:t>Result</w:t></w:r><w:r><b:fldChar b:fldCharType="e&#110;d"/></w:r>
<w:r><b:fldChar b:fldCharType="beg&#105;n"/></w:r>
<w:r><w:instrText>TOC \t &quot;Secret</w:instrText></w:r>
<w:r><w:delInstrText> Style,1&quot;</w:delInstrText></w:r>
<w:r><b:fldChar b:fldCharType="e&#110;d"/></w:r>"#;
    for random in [false, true] {
        let output = rewrite(body, random)
            .unwrap()
            .replace("</w:instrText></w:r>\n<w:r><w:instrText>", "")
            .replace("</w:instrText></w:r>\n<w:r><w:delInstrText>", "");
        for expected in [
            r#"STYLEREF "RedactedStyle1""#,
            r#"TOC \t "RedactedStyle1,1""#,
        ] {
            assert!(
                output.contains(quick_xml::escape::escape(expected).as_ref()),
                "{output}"
            );
        }
        assert!(!output.contains("Secret"), "{output}");
    }
}

#[test]
fn shared_fields_keep_bookmark_and_style_instructions_consistent() {
    let body = r#"<w:bookmarkStart w:id="1" w:name="SecretBookmark"/><w:bookmarkEnd w:id="1"/>
<w:r><w:fldChar w:fldCharType="begin"/></w:r>
<w:r><w:instrText>REF Secret</w:instrText></w:r>
<w:r><w:instrText>Bookmark</w:instrText></w:r>
<w:r><w:fldChar w:fldCharType="end"/></w:r>
<w:fldSimple b:instr="REF SecretBookmark"/>
<w:fldSimple b:instr="STYLEREF &quot;Secret Style&quot;"/>
<w:fldSimple b:instr="TOC \t &quot;Secret Style,1&quot;"/>"#;
    for random in [false, true] {
        let output = rewrite(body, random).unwrap();
        let name = output
            .split("w:name=\"")
            .nth(1)
            .unwrap()
            .split('"')
            .next()
            .unwrap();
        assert!(output.contains(&format!("REF {name}")), "{output}");
        assert_eq!(
            output.matches(&format!("REF {name}")).count(),
            2,
            "{output}"
        );
        assert_eq!(output.matches("RedactedStyle1").count(), 2, "{output}");
        assert!(!output.contains("Secret"), "{output}");
    }
}

#[test]
fn shared_fields_refuse_unknown_styles_in_complex_and_simple_fields() {
    for body in [
        r#"<w:r><b:fldChar b:fldCharType="beg&#105;n"/></w:r><w:r><w:instrText>STYLEREF Missing</w:instrText></w:r><w:r><b:fldChar b:fldCharType="e&#110;d"/></w:r>"#,
        r#"<w:fldSimple b:instr="STYLEREF Missing"/>"#,
        r#"<w:r><b:fldChar b:fldCharType="beg&#105;n"/></w:r><w:r><w:instrText>TOC \t &quot;Missing,1&quot;</w:instrText></w:r><w:r><b:fldChar b:fldCharType="e&#110;d"/></w:r>"#,
        r#"<w:fldSimple b:instr="TOC \t &quot;Missing,1&quot;"/>"#,
    ] {
        assert!(rewrite(body, false).is_err(), "{body}");
    }
}
