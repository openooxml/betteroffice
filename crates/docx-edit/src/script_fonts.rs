//! The fonts a seeded document names only for East Asian or complex-script
//! text it does not contain.

use std::collections::{BTreeMap, HashSet};

use ooxml_text::{FontSlotUse, font_slot_use};
use serde_json::Value;

/// `fontTable.xml` character sets (hex) of East Asian and of complex-script fonts.
const EAST_ASIAN_CHARSETS: [&str; 5] = ["80", "81", "82", "86", "88"];
const COMPLEX_SCRIPT_CHARSETS: [&str; 2] = ["b1", "b2"];
/// Longer strings in an embed's payload are not taken for font names.
const MAX_FONT_NAME_BYTES: usize = 128;
/// The families layout and display draw text with when a document names
/// none, such as the paragraph mark of a tracked insertion, a chart without
/// text properties, or math.
const BUILT_IN_FONTS: [&str; 2] = ["calibri", "cambria math"];

/// What a document's seeded runs measure with, gathered unit by unit: the
/// lowered run takes its `w:rFonts` slots and `w:cs` from its own attributes,
/// or else from its paragraph's run defaults. Embeds count conservatively:
/// any short string in a payload may name the font something is drawn with.
#[derive(Default)]
pub(crate) struct ScriptFontUse {
    /// Trimmed names some text may be measured or drawn with whatever its script.
    latin: HashSet<String>,
    east_asian: HashSet<String>,
    complex: HashSet<String>,
    measured: FontSlotUse,
    /// Text of runs with no fonts of their own, measured without and with an
    /// `eastAsia` hint from the paragraph.
    unfonted: FontSlotUse,
    unfonted_hinted_east_asia: bool,
    /// Whether some paragraph's run defaults hint `eastAsia`, or set `w:cs`.
    default_hint: bool,
    default_cs: bool,
}

fn lowercase(name: &str) -> Option<String> {
    let name = name.trim();
    (!name.is_empty()).then(|| name.to_lowercase())
}

/// Adds `name`, trimmed, allocating only for a name not seen yet.
fn note(names: &mut HashSet<String>, name: &str) {
    let name = name.trim();
    if !name.is_empty() && !names.contains(name) {
        names.insert(name.to_owned());
    }
}

fn lowercased(names: &HashSet<String>) -> HashSet<String> {
    names.iter().map(|name| name.to_lowercase()).collect()
}

fn slot<'a>(fonts: &'a serde_json::Map<String, Value>, key: &str) -> Option<&'a str> {
    fonts
        .get(key)
        .and_then(Value::as_str)
        .filter(|name| !name.trim().is_empty())
}

fn is_rtl(value: Option<&Value>) -> bool {
    value.and_then(Value::as_bool) == Some(true)
}

fn complex_script_flag(value: &Value) -> Option<bool> {
    match value {
        Value::Null => None,
        Value::Bool(flag) => Some(*flag),
        Value::Object(map) => Some(map.get("enabled").and_then(Value::as_bool).unwrap_or(true)),
        _ => Some(true),
    }
}

impl ScriptFontUse {
    /// A seeded text unit and its attributes.
    pub(crate) fn text(&mut self, text: &str, attrs: &BTreeMap<String, Value>) {
        self.run_fonts(attrs);
        if text.is_empty() {
            return;
        }
        let complex_script = attrs
            .get("complexScript")
            .and_then(complex_script_flag)
            .unwrap_or(false);
        match attrs.get("fontFamily").and_then(Value::as_object) {
            Some(fonts) => {
                let hint = fonts.get("hint").and_then(Value::as_str);
                merge(
                    &mut self.measured,
                    font_slot_use(text, complex_script, hint),
                );
            }
            None => {
                merge(
                    &mut self.unfonted,
                    font_slot_use(text, complex_script, None),
                );
                self.unfonted_hinted_east_asia = self.unfonted_hinted_east_asia
                    || font_slot_use(text, complex_script, Some("eastAsia")).east_asia;
            }
        }
    }

    /// A seeded embed's payload and attributes, including JSON-encoded parts
    /// such as `shapeJson`. Text inside a payload counts as text that may take
    /// its fonts from the paragraph, and as hinted or `w:cs` text when the
    /// payload holds such formatting anywhere.
    pub(crate) fn embed(
        &mut self,
        payload: &BTreeMap<String, Value>,
        attrs: &BTreeMap<String, Value>,
    ) {
        self.run_fonts(attrs);
        let mut embedded = Embedded::default();
        self.object(payload.iter(), is_rtl(payload.get("rtl")), &mut embedded);
        let hint = embedded.hinted.then_some("eastAsia");
        let complex_script: &[bool] = if embedded.complex_script {
            &[false, true]
        } else {
            &[false]
        };
        for text in embedded.texts.iter().filter(|text| !text.is_empty()) {
            for &complex_script in complex_script {
                merge(
                    &mut self.measured,
                    font_slot_use(text, complex_script, hint),
                );
                merge(
                    &mut self.unfonted,
                    font_slot_use(text, complex_script, None),
                );
                self.unfonted_hinted_east_asia = self.unfonted_hinted_east_asia
                    || font_slot_use(text, complex_script, Some("eastAsia")).east_asia;
            }
        }
    }

    /// The fonts and character sets of `fontTable.xml`.
    pub(crate) fn font_table(&mut self, fonts: &[docx_parse::FontInfo]) {
        for font in fonts {
            let charset = font.charset.as_deref().map(str::to_ascii_lowercase);
            let target = match charset.as_deref() {
                Some(charset) if EAST_ASIAN_CHARSETS.contains(&charset) => &mut self.east_asian,
                Some(charset) if COMPLEX_SCRIPT_CHARSETS.contains(&charset) => &mut self.complex,
                _ => continue,
            };
            for name in std::iter::once(font.name.as_str()).chain(font.alt_name.as_deref()) {
                note(target, name);
            }
        }
    }

    /// The names in `referenced` given only East Asian or complex-script roles
    /// whose text the document lacks, so nothing is measured or drawn with them.
    pub(crate) fn unused<'a>(
        &self,
        referenced: impl IntoIterator<Item = &'a String>,
    ) -> Vec<String> {
        let east_asian_text = self.measured.east_asia
            || self.unfonted.east_asia
            || (self.default_hint && self.unfonted_hinted_east_asia);
        let complex_text =
            self.measured.complex_script || self.unfonted.complex_script || self.default_cs;
        let latin = lowercased(&self.latin);
        let east_asian_names = lowercased(&self.east_asian);
        let complex_names = lowercased(&self.complex);
        referenced
            .into_iter()
            .filter(|name| {
                let Some(key) = lowercase(name) else {
                    return false;
                };
                let east_asian = east_asian_names.contains(&key);
                let complex = complex_names.contains(&key);
                !latin.contains(&key)
                    && !BUILT_IN_FONTS.contains(&key.as_str())
                    && (east_asian || complex)
                    && !(east_asian && east_asian_text)
                    && !(complex && complex_text)
            })
            .cloned()
            .collect()
    }

    fn run_fonts(&mut self, attrs: &BTreeMap<String, Value>) {
        if let Some(fonts) = attrs.get("fontFamily") {
            self.fonts(fonts, is_rtl(attrs.get("rtl")));
        }
    }

    fn object<'a>(
        &mut self,
        entries: impl Iterator<Item = (&'a String, &'a Value)>,
        rtl: bool,
        embedded: &mut Embedded,
    ) {
        for (key, value) in entries {
            // Copies of the source formatting kept for saving; layout reads
            // no fonts from them.
            if key.starts_with("_original") {
                continue;
            }
            self.entry(key, value, rtl);
            match (key.as_str(), value) {
                ("text" | "plainText", Value::String(text)) => embedded.texts.push(text.clone()),
                (key, Value::String(encoded)) if key.ends_with("Json") => {
                    if let Ok(decoded) = serde_json::from_str::<Value>(encoded) {
                        self.walk(&decoded, embedded);
                    }
                }
                ("cs" | "complexScript", Value::Bool(true)) => embedded.complex_script = true,
                ("hint", Value::String(hint)) => embedded.hinted |= hint == "eastAsia",
                ("eastAsia" | "ea", Value::String(name)) => note(&mut self.east_asian, name),
                ("cs", Value::String(name)) => note(&mut self.complex, name),
                (_, Value::String(name)) if name.len() <= MAX_FONT_NAME_BYTES => {
                    note(&mut self.latin, name);
                }
                _ => {}
            }
            self.walk(value, embedded);
        }
    }

    fn walk(&mut self, value: &Value, embedded: &mut Embedded) {
        match value {
            Value::Array(values) => {
                for value in values {
                    self.walk(value, embedded);
                }
            }
            Value::Object(values) => {
                self.object(values.iter(), is_rtl(values.get("rtl")), embedded);
            }
            _ => {}
        }
    }

    fn entry(&mut self, key: &str, value: &Value, rtl: bool) {
        match key {
            "fontFamily" | "listMarkerFontFamily" | "markerFontFamily" => self.fonts(value, rtl),
            "defaultTextFormatting" => {
                let Some(defaults) = value.as_object() else {
                    return;
                };
                self.default_cs |= defaults.get("cs").and_then(Value::as_bool) == Some(true);
                self.default_hint |= defaults
                    .get("fontFamily")
                    .and_then(|fonts| fonts.get("hint"))
                    .and_then(Value::as_str)
                    == Some("eastAsia");
            }
            _ => {}
        }
    }

    /// A `fontFamily` value as the render bridge lowers it: the run's own
    /// family is its `cs` font when right-to-left, else the first of its
    /// `ascii`, `hAnsi`, `eastAsia` and `cs` fonts.
    fn fonts(&mut self, value: &Value, rtl: bool) {
        match value {
            Value::String(name) => note(&mut self.latin, name),
            Value::Object(fonts) => {
                let ascii = slot(fonts, "ascii");
                let h_ansi = slot(fonts, "hAnsi");
                let east_asia = slot(fonts, "eastAsia");
                let cs = slot(fonts, "cs");
                let family = if rtl { cs } else { None }
                    .or(ascii)
                    .or(h_ansi)
                    .or(east_asia)
                    .or(cs);
                for name in [ascii, h_ansi, family].into_iter().flatten() {
                    note(&mut self.latin, name);
                }
                if let Some(name) = east_asia {
                    note(&mut self.east_asian, name);
                }
                if let Some(name) = cs {
                    note(&mut self.complex, name);
                }
            }
            _ => {}
        }
    }
}

#[derive(Default)]
struct Embedded {
    texts: Vec<String>,
    hinted: bool,
    complex_script: bool,
}

fn merge(target: &mut FontSlotUse, used: FontSlotUse) {
    target.east_asia |= used.east_asia;
    target.complex_script |= used.complex_script;
}

#[cfg(test)]
mod tests {
    use super::*;
    use serde_json::json;

    fn attrs(value: Value) -> BTreeMap<String, Value> {
        serde_json::from_value(value).unwrap()
    }

    fn font(name: &str, alt_name: Option<&str>, charset: &str) -> docx_parse::FontInfo {
        docx_parse::FontInfo {
            name: name.to_owned(),
            alt_name: alt_name.map(str::to_owned),
            panose1: None,
            charset: Some(charset.to_owned()),
            family: None,
            pitch: None,
            embed_regular: None,
            embed_bold: None,
            embed_italic: None,
            embed_bold_italic: None,
        }
    }

    fn unused(scan: &ScriptFontUse, names: &[&str]) -> Vec<String> {
        let names: Vec<String> = names.iter().map(|name| (*name).to_owned()).collect();
        scan.unused(&names)
    }

    const NAMES: [&str; 7] = [
        "Aptos",
        "SimSun",
        "宋体",
        "Batang",
        "Traditional Arabic",
        "Shared",
        "Scalar",
    ];

    fn latin_document() -> ScriptFontUse {
        let mut scan = ScriptFontUse::default();
        scan.font_table(&[
            font("SimSun", Some("宋体"), "86"),
            font("Batang", None, "81"),
        ]);
        let fonts = json!({"fontFamily": {
            "ascii": "Aptos", "hAnsi": "Aptos", "eastAsia": "Shared", "cs": "Traditional Arabic"
        }});
        scan.text("Latin \u{201c}quoted\u{201d} a\u{301}", &attrs(fonts));
        scan.text("abc", &attrs(json!({"fontFamily": {"eastAsia": "Scalar"}})));
        scan.embed(
            &attrs(json!({
                "defaultTextFormatting": {"fontFamily": {"ascii": "Aptos", "cs": "Shared"}},
                "_originalRunBoundaries": [{"formatting": {"fontFamily": {"eastAsia": "SimSun"}}}]
            })),
            &BTreeMap::new(),
        );
        scan
    }

    #[test]
    fn names_only_script_fonts_whose_text_is_absent() {
        let scan = latin_document();
        assert_eq!(
            unused(&scan, &NAMES),
            ["SimSun", "宋体", "Batang", "Traditional Arabic", "Shared"]
        );
    }

    #[test]
    fn keeps_script_fonts_the_text_is_measured_with() {
        let mut east_asian = latin_document();
        east_asian.text("漢字", &attrs(json!({"fontFamily": {"ascii": "Aptos"}})));
        assert_eq!(unused(&east_asian, &NAMES), ["Traditional Arabic"]);

        let mut complex = latin_document();
        complex.text("abc", &attrs(json!({"complexScript": true})));
        assert_eq!(unused(&complex, &NAMES), ["SimSun", "宋体", "Batang"]);
    }

    #[test]
    fn keeps_script_fonts_a_paragraph_default_applies_to_its_text() {
        let mut hinted = latin_document();
        hinted.text("\u{201c}", &BTreeMap::new());
        assert_eq!(unused(&hinted, &NAMES).len(), 5);
        hinted.embed(
            &attrs(json!({"defaultTextFormatting": {"fontFamily": {"hint": "eastAsia"}}})),
            &BTreeMap::new(),
        );
        assert_eq!(unused(&hinted, &NAMES), ["Traditional Arabic"]);

        let mut complex = latin_document();
        complex.embed(
            &attrs(json!({"defaultTextFormatting": {"cs": true}})),
            &BTreeMap::new(),
        );
        assert_eq!(unused(&complex, &NAMES), ["SimSun", "宋体", "Batang"]);
    }

    #[test]
    fn keeps_script_fonts_for_text_inside_an_embed() {
        let mut scan = latin_document();
        scan.embed(
            &attrs(
                json!({"runs": [{"formatting": {"fontFamily": {"hint": "eastAsia"}},
                "content": [{"type": "text", "text": "\u{2014}"}]}]}),
            ),
            &BTreeMap::new(),
        );
        assert_eq!(unused(&scan, &NAMES), ["Traditional Arabic"]);

        let mut math = latin_document();
        math.embed(&attrs(json!({"plainText": "\u{6f22}"})), &BTreeMap::new());
        assert_eq!(unused(&math, &NAMES), ["Traditional Arabic"]);

        let mut fallbacks = ScriptFontUse::default();
        let defaults =
            json!({"fontFamily": {"ascii": "Arial", "cs": "Calibri", "eastAsia": "Cambria Math"}});
        fallbacks.embed(
            &attrs(json!({"defaultTextFormatting": defaults})),
            &BTreeMap::new(),
        );
        assert!(unused(&fallbacks, &["Calibri", "Cambria Math"]).is_empty());

        let mut shape = latin_document();
        let shape_json = json!({"paragraphs": [{"runs": [{"fontFamily": {"ascii": "SimSun"}}]}]});
        shape.embed(
            &attrs(json!({"shapeJson": shape_json.to_string()})),
            &BTreeMap::new(),
        );
        assert_eq!(
            unused(&shape, &NAMES),
            ["宋体", "Batang", "Traditional Arabic", "Shared"]
        );

        let mut chart = latin_document();
        let chart_json = json!({"legend": {"text": {"font": "SimSun"}}});
        chart.embed(
            &attrs(json!({"chartJson": chart_json.to_string()})),
            &BTreeMap::new(),
        );
        assert_eq!(
            unused(&chart, &NAMES),
            ["宋体", "Batang", "Traditional Arabic", "Shared"]
        );

        let mut runs = latin_document();
        runs.embed(
            &attrs(json!({"runs": [
                {"cs": false, "text": "\u{6f22}"},
                {"cs": true, "text": "abc"}
            ]})),
            &BTreeMap::new(),
        );
        assert!(unused(&runs, &NAMES).is_empty());
    }

    #[test]
    fn keeps_the_complex_script_font_a_right_to_left_run_is_drawn_with() {
        let mut scan = latin_document();
        scan.text(
            "abc",
            &attrs(
                json!({"rtl": true, "fontFamily": {"ascii": "Aptos", "cs": "Traditional Arabic"}}),
            ),
        );
        assert_eq!(
            unused(&scan, &NAMES),
            ["SimSun", "宋体", "Batang", "Shared"]
        );
    }
}
