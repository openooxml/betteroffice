//! S10 legacy VML watermark serializer.

use crate::vml::Watermark;

use super::xml_writer::{XmlWriter, js_number};

const EMU_PER_PT: f64 = 12_700.0;

/// Serialize a text or resolved picture watermark as the header paragraph Word
/// expects. An unresolved picture relationship intentionally emits nothing.
pub fn serialize_watermark(watermark: &Watermark) -> String {
    match watermark {
        Watermark::Text {
            text,
            font,
            color,
            semitransparent,
            layout,
            font_size,
        } => serialize_text_watermark(text, font, color, *semitransparent, layout, *font_size),
        Watermark::Picture {
            relationship_id,
            scale,
            washout,
            width_emu,
            height_emu,
            ..
        } => serialize_picture_watermark(
            relationship_id.as_deref(),
            *scale,
            *washout,
            *width_emu,
            *height_emu,
        ),
    }
}

fn serialize_text_watermark(
    text: &str,
    font: &str,
    color: &str,
    semitransparent: bool,
    layout: &str,
    font_size: Option<f64>,
) -> String {
    let characters = javascript_trim(text).encode_utf16().count().max(1);
    let width = (characters * 26).clamp(120, 468) as f64;
    let height = width / 2.0;
    let rotation = if layout == "diagonal" {
        ";rotation:315"
    } else {
        ""
    };
    let style = format!(
        "position:absolute;margin-left:0;margin-top:0;width:{}pt;height:{}pt{};z-index:-251658240;mso-position-horizontal:center;mso-position-horizontal-relative:margin;mso-position-vertical:center;mso-position-vertical-relative:margin",
        js_number(width),
        js_number(height),
        rotation
    );
    let text_path_style = format!(
        "font-family:\"{font}\";font-size:{}pt",
        js_number(font_size.unwrap_or(1.0))
    );

    let mut writer = XmlWriter::with_capacity(1_600);
    write_watermark_prefix(&mut writer);
    write_text_shapetype(&mut writer);
    writer
        .start_element("v:shape")
        .attribute("id", "PowerPlusWaterMarkObject1")
        .attribute("o:spid", "_x0000_s2049")
        .attribute("type", "#_x0000_t136")
        .attribute("style", &style)
        .attribute("o:allowincell", "f")
        .attribute("fillcolor", color)
        .attribute("stroked", "f");
    if semitransparent {
        writer
            .start_element("v:fill")
            .attribute("opacity", ".5")
            .end_element();
    }
    writer
        .start_element("v:textpath")
        .attribute("style", &text_path_style)
        .attribute("string", text)
        .end_element()
        .end_element();
    write_watermark_suffix(&mut writer);
    writer.finish()
}

fn serialize_picture_watermark(
    relationship_id: Option<&str>,
    scale: f64,
    washout: bool,
    width_emu: Option<f64>,
    height_emu: Option<f64>,
) -> String {
    let Some(relationship_id) = relationship_id.filter(|value| !value.is_empty()) else {
        return String::new();
    };
    let width = width_emu.map(|value| value / EMU_PER_PT).unwrap_or(311.4);
    let height = height_emu.map(|value| value / EMU_PER_PT).unwrap_or(width);
    let scale = if scale == 0.0 { 1.0 } else { scale };
    let style = format!(
        "position:absolute;margin-left:0;margin-top:0;width:{}pt;height:{}pt;z-index:-251657216;mso-position-horizontal:center;mso-position-horizontal-relative:margin;mso-position-vertical:center;mso-position-vertical-relative:margin",
        js_number(width * scale),
        js_number(height * scale)
    );

    let mut writer = XmlWriter::with_capacity(1_400);
    write_watermark_prefix(&mut writer);
    write_picture_shapetype(&mut writer);
    writer
        .start_element("v:shape")
        .attribute("id", "WordPictureWatermark1")
        .attribute("o:spid", "_x0000_s2050")
        .attribute("type", "#_x0000_t75")
        .attribute("style", &style)
        .attribute("o:allowincell", "f")
        .start_element("v:imagedata")
        .attribute("r:id", relationship_id)
        .attribute("o:title", "watermark");
    if washout {
        writer
            .attribute("gain", "19661f")
            .attribute("blacklevel", "22938f");
    }
    writer.end_element().end_element();
    write_watermark_suffix(&mut writer);
    writer.finish()
}

fn write_watermark_prefix(writer: &mut XmlWriter) {
    writer
        .start_element("w:p")
        .start_element("w:r")
        .start_element("w:rPr")
        .start_element("w:noProof")
        .end_element()
        .end_element()
        .start_element("w:pict");
}

fn write_watermark_suffix(writer: &mut XmlWriter) {
    writer.end_element().end_element().end_element();
}

fn write_text_shapetype(writer: &mut XmlWriter) {
    writer
        .start_element("v:shapetype")
        .attribute("id", "_x0000_t136")
        .attribute("coordsize", "21600,21600")
        .attribute("o:spt", "136")
        .attribute("adj", "10800")
        .attribute("path", "m@7,l@8,m@5,21600l@6,21600e")
        .start_element("v:formulas");
    for equation in [
        "sum #0 0 10800",
        "prod #0 2 1",
        "sum 21600 0 @1",
        "sum 0 0 @2",
        "sum 21600 0 @3",
        "if @0 @3 0",
        "if @0 21600 @1",
        "if @0 0 @2",
        "if @0 @4 21600",
        "mid @5 @6",
        "mid @8 @5",
        "mid @7 @8",
        "mid @6 @7",
        "sum @6 0 @5",
    ] {
        writer
            .start_element("v:f")
            .attribute("eqn", equation)
            .end_element();
    }
    writer
        .end_element()
        .start_element("v:path")
        .attribute("textpathok", "t")
        .attribute("o:connecttype", "custom")
        .attribute("o:connectlocs", "@9,0;@10,10800;@9,21600;@11,10800")
        .attribute("o:connectangles", "270,180,90,0")
        .end_element()
        .start_element("v:textpath")
        .attribute("on", "t")
        .attribute("fitshape", "t")
        .end_element()
        .start_element("v:handles")
        .start_element("v:h")
        .attribute("position", "#0,bottomRight")
        .attribute("xrange", "6629,14971")
        .end_element()
        .end_element()
        .start_element("o:lock")
        .attribute("v:ext", "edit")
        .attribute("text", "t")
        .attribute("shapetype", "t")
        .end_element()
        .end_element();
}

fn write_picture_shapetype(writer: &mut XmlWriter) {
    writer
        .start_element("v:shapetype")
        .attribute("id", "_x0000_t75")
        .attribute("coordsize", "21600,21600")
        .attribute("o:spt", "75")
        .attribute("o:preferrelative", "t")
        .attribute("path", "m@4@5l@4@11@9@11@9@5xe")
        .attribute("filled", "f")
        .attribute("stroked", "f")
        .start_element("v:stroke")
        .attribute("joinstyle", "miter")
        .end_element()
        .start_element("v:formulas");
    for equation in [
        "if lineDrawn pixelLineWidth 0",
        "sum @0 1 0",
        "sum 0 0 @1",
        "prod @2 1 2",
        "prod @3 21600 pixelWidth",
        "prod @3 21600 pixelHeight",
        "sum @0 0 1",
        "prod @6 1 2",
        "prod @7 21600 pixelWidth",
        "sum @8 21600 0",
        "prod @7 21600 pixelHeight",
        "sum @10 21600 0",
    ] {
        writer
            .start_element("v:f")
            .attribute("eqn", equation)
            .end_element();
    }
    writer
        .end_element()
        .start_element("v:path")
        .attribute("o:extrusionok", "f")
        .attribute("gradientshapeok", "t")
        .attribute("o:connecttype", "rect")
        .end_element()
        .start_element("o:lock")
        .attribute("v:ext", "edit")
        .attribute("aspectratio", "t")
        .end_element()
        .end_element();
}

fn javascript_trim(value: &str) -> &str {
    value.trim_matches(is_ecmascript_whitespace)
}

fn is_ecmascript_whitespace(character: char) -> bool {
    matches!(
        character,
        '\u{0009}'
            ..='\u{000D}'
                | '\u{0020}'
                | '\u{00A0}'
                | '\u{1680}'
                | '\u{2000}'..='\u{200A}'
                | '\u{2028}'
                | '\u{2029}'
                | '\u{202F}'
                | '\u{205F}'
                | '\u{3000}'
                | '\u{FEFF}'
    )
}

#[cfg(test)]
mod tests {
    use crate::xml::{ParseBudget, ParseLimits, parse_xml};
    use sha2::{Digest, Sha256};

    use super::*;

    #[test]
    fn text_watermark_matches_incumbent_geometry_and_order() {
        let watermark = Watermark::Text {
            text: "DRAFT".to_owned(),
            font: "Calibri".to_owned(),
            color: "#C0C0C0".to_owned(),
            semitransparent: true,
            layout: "diagonal".to_owned(),
            font_size: Some(42.5),
        };
        let xml = serialize_watermark(&watermark);
        assert_eq!(xml.len(), 1_335);
        assert_eq!(
            format!("{:x}", Sha256::digest(xml.as_bytes())),
            "d90d90638898524669dc4c4705f5743396b482d825fd23de2e82b98ca634a8e9"
        );
        assert!(xml.starts_with(
            "<w:p><w:r><w:rPr><w:noProof/></w:rPr><w:pict><v:shapetype id=\"_x0000_t136\""
        ));
        assert!(xml.contains("width:130pt;height:65pt;rotation:315"));
        assert!(xml.contains("<v:fill opacity=\".5\"/>"));
        assert!(xml.contains(
            "<v:textpath style=\"font-family:&quot;Calibri&quot;;font-size:42.5pt\" string=\"DRAFT\"/>"
        ));
        assert!(xml.ends_with("</v:shape></w:pict></w:r></w:p>"));
    }

    #[test]
    fn text_size_counts_javascript_utf16_units_after_trim() {
        let watermark = Watermark::Text {
            text: "\u{FEFF}😀😀😀😀😀\u{3000}".to_owned(),
            font: "Calibri".to_owned(),
            color: "black".to_owned(),
            semitransparent: false,
            layout: "horizontal".to_owned(),
            font_size: None,
        };
        let xml = serialize_watermark(&watermark);
        assert!(xml.contains("width:260pt;height:130pt;z-index"));
    }

    #[test]
    fn picture_watermark_matches_scaling_washout_and_unresolved_omission() {
        let unresolved = Watermark::Picture {
            relationship_id: None,
            media_path: None,
            content_type: None,
            data_url: None,
            scale: 1.0,
            washout: false,
            width_emu: None,
            height_emu: None,
        };
        assert_eq!(serialize_watermark(&unresolved), "");

        let watermark = Watermark::Picture {
            relationship_id: Some("rId7".to_owned()),
            media_path: None,
            content_type: None,
            data_url: None,
            scale: 0.5,
            washout: true,
            width_emu: Some(3_954_780.0),
            height_emu: Some(1_977_390.0),
        };
        let xml = serialize_watermark(&watermark);
        assert_eq!(xml.len(), 1_166);
        assert_eq!(
            format!("{:x}", Sha256::digest(xml.as_bytes())),
            "b8cbeb6d07e256a41ee355bcf6bad629b3c04c736d435cac21feb22c83c2d58d"
        );
        assert!(xml.contains("width:155.7pt;height:77.85pt;z-index"));
        assert!(xml.contains(
            "<v:imagedata r:id=\"rId7\" o:title=\"watermark\" gain=\"19661f\" blacklevel=\"22938f\"/>"
        ));
    }

    #[test]
    fn escapes_every_watermark_string_without_double_escaping() {
        let attack = "\"/><evil attr='&";
        let watermark = Watermark::Text {
            text: attack.to_owned(),
            font: attack.to_owned(),
            color: attack.to_owned(),
            semitransparent: false,
            layout: "horizontal".to_owned(),
            font_size: None,
        };
        let xml = serialize_watermark(&watermark);
        assert!(!xml.contains("<evil"));
        assert!(!xml.contains("attr='"));
        assert_eq!(
            xml.matches("&quot;/&gt;&lt;evil attr=&apos;&amp;").count(),
            3
        );
        assert!(xml.contains("font-family:&quot;&quot;/&gt;&lt;evil"));

        let picture = Watermark::Picture {
            relationship_id: Some(attack.to_owned()),
            media_path: None,
            content_type: None,
            data_url: None,
            scale: 1.0,
            washout: false,
            width_emu: None,
            height_emu: None,
        };
        assert!(!serialize_watermark(&picture).contains("<evil"));
    }

    #[test]
    fn emitted_text_watermark_parses_back() {
        let original = Watermark::Text {
            text: "CONFIDENTIAL".to_owned(),
            font: "Aptos".to_owned(),
            color: "#D8D8D8".to_owned(),
            semitransparent: true,
            layout: "diagonal".to_owned(),
            font_size: None,
        };
        let fragment = serialize_watermark(&original);
        let xml = format!(
            "<w:hdr xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:v=\"urn:schemas-microsoft-com:vml\" xmlns:o=\"urn:schemas-microsoft-com:office:office\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">{fragment}</w:hdr>"
        );
        let limits = ParseLimits::default();
        let parsed = parse_xml(
            xml.as_bytes(),
            "word/header1.xml",
            &mut ParseBudget::new(&limits),
        )
        .unwrap();
        assert_eq!(
            crate::vml::extract_watermark(parsed.root(), None, None),
            Some(original)
        );
    }
}
