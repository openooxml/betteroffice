use std::collections::BTreeMap;

use crate::xml::Node;
use crate::{
    FontPreflight, Parts, PreflightError, PreflightFace, Relationship, relationships_from,
};

const MAX_FACES: usize = 512;

#[derive(Default)]
pub(crate) struct Faces {
    values: BTreeMap<(String, bool, bool), PreflightFace>,
    pub(crate) limited: bool,
}

impl Faces {
    pub(crate) fn record(&mut self, family: &str, bold: bool, italic: bool, scripts: &[String]) {
        let family = family.trim();
        if family.is_empty() {
            return;
        }
        if family.len() > crate::xml::MAX_VALUE_BYTES {
            self.limited = true;
            return;
        }
        let key = (family.to_lowercase(), bold, italic);
        if self.values.len() == MAX_FACES && !self.values.contains_key(&key) {
            self.limited = true;
            return;
        }
        let face = self.values.entry(key).or_insert_with(|| PreflightFace {
            family: family.to_owned(),
            bold,
            italic,
            scripts: Vec::new(),
        });
        for script in scripts {
            if !face.scripts.contains(script) {
                face.scripts.push(script.clone());
            }
        }
    }

    pub(crate) fn finish(self, complete: bool) -> FontPreflight {
        FontPreflight {
            faces: self.values.into_values().collect(),
            complete: complete && !self.limited,
        }
    }
}

impl Parts<'_> {
    pub(crate) fn rels(
        &mut self,
        source: &str,
        complete: &mut bool,
    ) -> Result<Vec<Relationship>, PreflightError> {
        let (directory, file) = source.rsplit_once('/').unwrap_or(("", source));
        let path = format!("{directory}/_rels/{file}.rels");
        let root = self.part(&path, complete)?;
        if root
            .as_ref()
            .is_some_and(|root| root.name != "Relationships")
        {
            return Err(PreflightError::Xml);
        }
        Ok(relationships_from(root.as_ref(), directory, complete))
    }

    #[cfg(feature = "pptx")]
    pub(crate) fn related(
        &mut self,
        rels: &[Relationship],
        kind: &str,
        complete: &mut bool,
    ) -> Result<Option<(String, Node)>, PreflightError> {
        let Some(rel) = rels.iter().find(|rel| rel.kind == kind) else {
            return Ok(None);
        };
        let root = self.part(&rel.target, complete)?;
        if root.is_none() {
            *complete = false;
        }
        Ok(root.map(|root| (rel.target.clone(), root)))
    }
}

#[cfg(feature = "pptx")]
pub(crate) struct DrawingTheme {
    fonts: [[String; 3]; 2],
}

#[cfg(feature = "pptx")]
impl DrawingTheme {
    pub(crate) fn parse(root: Option<&Node>) -> Self {
        let mut theme = Self {
            fonts: [
                ["Calibri Light".into(), String::new(), String::new()],
                ["Calibri".into(), String::new(), String::new()],
            ],
        };
        for (index, name) in ["a:majorFont", "a:minorFont"].into_iter().enumerate() {
            if let Some(font) =
                root.and_then(|root| root.path(&["a:themeElements", "a:fontScheme", name]))
            {
                for (slot, name) in ["a:latin", "a:ea", "a:cs"].into_iter().enumerate() {
                    theme.fonts[index][slot] = font
                        .child(name)
                        .and_then(|node| node.attr("typeface"))
                        .unwrap_or_default()
                        .to_owned();
                }
            }
        }
        theme
    }

    pub(crate) fn resolve(&self, family: &str) -> String {
        if !family.starts_with('+') {
            return family.to_owned();
        }
        let lower = family.trim_start_matches('+').to_ascii_lowercase();
        let (major, slot, drawing) = match lower.split_once('-') {
            Some((kind @ ("mj" | "mn"), slot)) => (kind == "mj", slot, true),
            _ => (
                lower.starts_with("major"),
                lower
                    .strip_prefix("major")
                    .or_else(|| lower.strip_prefix("minor"))
                    .unwrap_or(&lower),
                false,
            ),
        };
        let fonts = &self.fonts[usize::from(!major)];
        let slot = match slot {
            "ea" | "eastasia" => 1,
            "cs" | "bidi" => 2,
            _ => 0,
        };
        let latin = || {
            if fonts[0].is_empty() {
                if major { "Calibri Light" } else { "Calibri" }.to_owned()
            } else {
                fonts[0].clone()
            }
        };
        if slot == 0 || (drawing && fonts[slot].is_empty()) {
            latin()
        } else {
            fonts[slot].clone()
        }
    }
}
