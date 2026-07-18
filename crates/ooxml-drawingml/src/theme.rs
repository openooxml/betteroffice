use indexmap::{IndexMap, IndexSet};
use serde::{Deserialize, Serialize};

#[derive(Clone, Debug, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ThemeColorScheme {
    pub dk1: String,
    pub lt1: String,
    pub dk2: String,
    pub lt2: String,
    pub accent1: String,
    pub accent2: String,
    pub accent3: String,
    pub accent4: String,
    pub accent5: String,
    pub accent6: String,
    pub hlink: String,
    pub fol_hlink: String,
}

impl Default for ThemeColorScheme {
    fn default() -> Self {
        Self {
            dk1: "000000".to_owned(),
            lt1: "FFFFFF".to_owned(),
            dk2: "44546A".to_owned(),
            lt2: "E7E6E6".to_owned(),
            accent1: "4472C4".to_owned(),
            accent2: "ED7D31".to_owned(),
            accent3: "A5A5A5".to_owned(),
            accent4: "FFC000".to_owned(),
            accent5: "5B9BD5".to_owned(),
            accent6: "70AD47".to_owned(),
            hlink: "0563C1".to_owned(),
            fol_hlink: "954F72".to_owned(),
        }
    }
}

impl ThemeColorScheme {
    pub fn set(&mut self, slot: &str, value: String) {
        match slot {
            "dk1" => self.dk1 = value,
            "lt1" => self.lt1 = value,
            "dk2" => self.dk2 = value,
            "lt2" => self.lt2 = value,
            "accent1" => self.accent1 = value,
            "accent2" => self.accent2 = value,
            "accent3" => self.accent3 = value,
            "accent4" => self.accent4 = value,
            "accent5" => self.accent5 = value,
            "accent6" => self.accent6 = value,
            "hlink" => self.hlink = value,
            "folHlink" => self.fol_hlink = value,
            _ => {}
        }
    }

    pub fn get(&self, slot: &str) -> Option<&str> {
        match slot {
            "dk1" => Some(&self.dk1),
            "lt1" => Some(&self.lt1),
            "dk2" => Some(&self.dk2),
            "lt2" => Some(&self.lt2),
            "accent1" => Some(&self.accent1),
            "accent2" => Some(&self.accent2),
            "accent3" => Some(&self.accent3),
            "accent4" => Some(&self.accent4),
            "accent5" => Some(&self.accent5),
            "accent6" => Some(&self.accent6),
            "hlink" => Some(&self.hlink),
            "folHlink" => Some(&self.fol_hlink),
            _ => None,
        }
        .map(String::as_str)
    }
}

#[derive(Clone, Debug, PartialEq, Eq, Serialize, Deserialize)]
pub struct ThemeFont {
    pub latin: String,
    pub ea: String,
    pub cs: String,
    pub fonts: IndexMap<String, String>,
}

impl ThemeFont {
    pub fn default_major() -> Self {
        Self {
            latin: "Calibri Light".to_owned(),
            ea: String::new(),
            cs: String::new(),
            fonts: IndexMap::new(),
        }
    }

    pub fn default_minor() -> Self {
        Self {
            latin: "Calibri".to_owned(),
            ea: String::new(),
            cs: String::new(),
            fonts: IndexMap::new(),
        }
    }

    pub fn empty() -> Self {
        Self {
            latin: String::new(),
            ea: String::new(),
            cs: String::new(),
            fonts: IndexMap::new(),
        }
    }
}

#[derive(Clone, Debug, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ThemeFontScheme {
    pub major_font: ThemeFont,
    pub minor_font: ThemeFont,
}

impl Default for ThemeFontScheme {
    fn default() -> Self {
        Self {
            major_font: ThemeFont::default_major(),
            minor_font: ThemeFont::default_minor(),
        }
    }
}

#[derive(Clone, Debug, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct Theme {
    pub name: String,
    pub color_scheme: ThemeColorScheme,
    pub font_scheme: ThemeFontScheme,
}

impl Default for Theme {
    fn default() -> Self {
        Self {
            name: "Office Theme".to_owned(),
            color_scheme: ThemeColorScheme::default(),
            font_scheme: ThemeFontScheme::default(),
        }
    }
}

pub fn get_theme_color(theme: Option<&Theme>, slot: &str) -> String {
    if let Some(value) = theme.and_then(|theme| theme.color_scheme.get(slot)) {
        return value.to_owned();
    }
    let defaults = ThemeColorScheme::default();
    defaults.get(slot).unwrap_or("000000").to_owned()
}

pub fn get_major_font(theme: Option<&Theme>, script: &str) -> String {
    get_font(
        theme.map(|theme| &theme.font_scheme.major_font),
        script,
        "Calibri Light",
    )
}

pub fn get_minor_font(theme: Option<&Theme>, script: &str) -> String {
    get_font(
        theme.map(|theme| &theme.font_scheme.minor_font),
        script,
        "Calibri",
    )
}

fn get_font(font: Option<&ThemeFont>, script: &str, latin_default: &str) -> String {
    let Some(font) = font else {
        return latin_default.to_owned();
    };
    match script {
        "latin" => nonempty(Some(&font.latin)).unwrap_or_else(|| latin_default.to_owned()),
        "ea" => font.ea.clone(),
        "cs" => font.cs.clone(),
        script => font
            .fonts
            .get(script)
            .cloned()
            .or_else(|| nonempty(Some(&font.latin)))
            .unwrap_or_else(|| latin_default.to_owned()),
    }
}

pub fn resolve_theme_font_ref(theme: Option<&Theme>, reference: &str) -> String {
    if reference.is_empty() {
        return "Calibri".to_owned();
    }
    let lower = reference.to_ascii_lowercase();
    let script = if lower.contains("eastasia") {
        "ea"
    } else if lower.contains("bidi") || lower.contains("cs") {
        "cs"
    } else {
        "latin"
    };
    if lower.contains("major") {
        get_major_font(theme, script)
    } else {
        get_minor_font(theme, script)
    }
}

pub fn get_theme_fonts(theme: Option<&Theme>) -> Vec<String> {
    let mut fonts = IndexSet::new();
    if let Some(theme) = theme {
        for font in [&theme.font_scheme.major_font, &theme.font_scheme.minor_font] {
            for value in [&font.latin, &font.ea, &font.cs] {
                if !value.is_empty() {
                    fonts.insert(value.clone());
                }
            }
        }
        for value in theme
            .font_scheme
            .major_font
            .fonts
            .values()
            .chain(theme.font_scheme.minor_font.fonts.values())
        {
            if !value.is_empty() {
                fonts.insert(value.clone());
            }
        }
    }
    fonts.into_iter().collect()
}

pub fn get_default_theme() -> Theme {
    Theme::default()
}

fn nonempty(value: Option<&str>) -> Option<String> {
    value.filter(|value| !value.is_empty()).map(str::to_owned)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn defaults_and_font_resolution_match_office_theme() {
        let theme = Theme::default();
        assert_eq!(get_theme_color(Some(&theme), "accent1"), "4472C4");
        assert_eq!(get_major_font(Some(&theme), "latin"), "Calibri Light");
        assert_eq!(get_minor_font(Some(&theme), "latin"), "Calibri");
    }
}
