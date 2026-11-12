use std::collections::HashMap;
#[cfg(feature = "fonts")]
use std::collections::HashSet;
use std::sync::PoisonError;

#[cfg(feature = "fonts")]
use betteroffice_fonts::{FontManager, LastResortOffice};
#[cfg(all(feature = "fonts", feature = "raster"))]
use docx_layout::display_list::{DisplayList, Primitive};
#[cfg(feature = "fonts")]
use docx_layout::measure_blocks::FontRequirement;
use ooxml_text::FontStore;

use crate::{Document, Error, Result};

#[derive(Default)]
pub(crate) struct FontRegistry {
    pub(crate) store: FontStore,
    pub(crate) chains: HashMap<String, Vec<ooxml_text::FontId>>,
    #[cfg(feature = "raster")]
    pub(crate) glyphs: docx_raster::GlyphCache,
    embedded: HashMap<String, usize>,
    application: HashMap<String, usize>,
    #[cfg(feature = "fonts")]
    resolved: HashSet<(String, Vec<String>)>,
    #[cfg(feature = "fonts")]
    pub(crate) cjk_fallbacks: bool,
    #[cfg(feature = "fonts")]
    pub(crate) pdfa_fallbacks: bool,
    views: HashMap<(ooxml_text::FontId, String), ooxml_text::FontId>,
    registrations: usize,
}

/// Bundled monochrome Noto Emoji (OFL-1.1) for PDF/A painter fallback.
///
/// A variable font at its default `wght` instance. The PDF writer subsets it
/// through `subsetter`'s default-location TrueType path, which keeps the
/// default outlines and drops `fvar`/`gvar`/`HVAR`, so the embedded program is
/// a static TrueType face.
#[cfg(feature = "fonts")]
const PDFA_EMOJI_FALLBACK: &[u8] = include_bytes!("../assets/NotoEmoji.ttf");

impl FontRegistry {
    /// Ordered face chains keyed `family|bold|italic`, exactly as layout
    /// resolved them. PDF text painting follows this order so application and
    /// embedded aliases keep their priority over bundled faces.
    pub(crate) fn chains(&self) -> &HashMap<String, Vec<ooxml_text::FontId>> {
        &self.chains
    }

    fn register(
        &mut self,
        family: &str,
        bold: bool,
        italic: bool,
        bytes: &[u8],
        priority: u8,
    ) -> Result<u32> {
        if family.trim().is_empty() {
            return Err(Error::Font("font family is empty".into()));
        }
        self.register_keyed(chain_key(family, bold, italic), bytes, priority)
    }

    /// Registers a provider fallback face for export without adding a chain.
    ///
    /// Applies the same per-font, registration-count, and distinct-face limits
    /// as [`Self::register_keyed`], so a provider cannot bypass the facade's
    /// resource contract through PDF export.
    #[cfg(feature = "fonts")]
    pub(crate) fn register_export_fallback(&mut self, bytes: &[u8]) -> Result<()> {
        if bytes.len() > 32 * 1024 * 1024 {
            return Err(Error::ResourceLimit("font exceeds 33554432 bytes".into()));
        }
        if self.registrations >= 4_096 {
            return Err(Error::ResourceLimit(
                "more than 4096 font registrations".into(),
            ));
        }
        if self.store.face_count() >= 256 && self.store.find_face(bytes).is_none() {
            return Err(Error::ResourceLimit("more than 256 font faces".into()));
        }
        self.store
            .register_slice(bytes)
            .map_err(|error| Error::Font(error.to_string()))?;
        self.registrations += 1;
        Ok(())
    }

    /// Registers the bundled Noto Emoji face as a PDF/A painter fallback.
    ///
    /// The face stays out of [`Self::chains`]: the painter resolves unmatched
    /// scalars against the store in registration order, so calling this after
    /// document and CJK registration leaves their priority unchanged. The
    /// helper is idempotent per registry and goes through
    /// [`Self::register_export_fallback`], so its resource limits apply.
    #[cfg(feature = "fonts")]
    pub(crate) fn register_pdfa_fallbacks(&mut self) -> Result<()> {
        if self.pdfa_fallbacks {
            return Ok(());
        }
        self.register_export_fallback(PDFA_EMOJI_FALLBACK)?;
        self.pdfa_fallbacks = true;
        Ok(())
    }

    /// Drops the export-local PDF/A face and its registration budget.
    ///
    /// `font_count` is the count captured before
    /// [`Self::register_pdfa_fallbacks`]. Removing the face without returning
    /// its one registration would make repeated PDF/A exports of the same
    /// document consume the 4096-registration budget.
    #[cfg(feature = "fonts")]
    pub(crate) fn release_pdfa_fallbacks(&mut self, font_count: usize) {
        if self.pdfa_fallbacks {
            self.store.truncate_fonts(font_count);
            self.registrations = self.registrations.saturating_sub(1);
            self.pdfa_fallbacks = false;
        }
    }

    /// Registers `bytes` into the chain under `key`, ahead of bundled faces for
    /// embedded (`0`) and application (`1`) priorities.
    fn register_keyed(&mut self, key: String, bytes: &[u8], priority: u8) -> Result<u32> {
        let id = self.register_bytes(bytes)?;
        Ok(self.insert(key, id, priority))
    }

    fn register_bytes(&mut self, bytes: &[u8]) -> Result<ooxml_text::FontId> {
        if bytes.len() > 32 * 1024 * 1024 {
            return Err(Error::ResourceLimit("font exceeds 33554432 bytes".into()));
        }
        if self.registrations >= 4_096 {
            return Err(Error::ResourceLimit(
                "more than 4096 font registrations".into(),
            ));
        }
        if self.store.face_count() >= 256 && self.store.find_face(bytes).is_none() {
            return Err(Error::ResourceLimit("more than 256 font faces".into()));
        }
        self.store
            .register_slice(bytes)
            .map_err(|error| Error::Font(error.to_string()))
    }

    fn insert(&mut self, key: String, id: ooxml_text::FontId, priority: u8) -> u32 {
        let chain = self.chains.entry(key.clone()).or_default();
        if priority == 2 && chain.contains(&id) {
            return id.to_u32();
        }
        let embedded = *self.embedded.get(&key).unwrap_or(&0);
        let application = *self.application.get(&key).unwrap_or(&0);
        match priority {
            0 => {
                chain.insert(embedded, id);
                *self.embedded.entry(key).or_default() += 1;
            }
            1 => {
                chain.insert(embedded + application, id);
                *self.application.entry(key).or_default() += 1;
            }
            _ => chain.push(id),
        }
        self.registrations += 1;
        id.to_u32()
    }

    /// `base` measured as Word measures `family`; `base` itself when `family` has no metrics.
    #[cfg(feature = "fonts")]
    fn measured_as(&mut self, base: ooxml_text::FontId, family: &str) -> ooxml_text::FontId {
        let owner = self.store.byte_owner(base).unwrap_or(base);
        let key = (owner, family.trim().to_lowercase());
        if let Some(view) = self.views.get(&key) {
            return *view;
        }
        let view = self
            .store
            .register_named_substitute(base, family)
            .unwrap_or(base);
        self.views.insert(key, view);
        view
    }

    #[cfg(feature = "fonts")]
    pub(crate) fn resolve(
        &mut self,
        manager: &FontManager,
        requirements: &[FontRequirement],
    ) -> Result<()> {
        for request in requirements {
            let key = chain_key(&request.family, request.bold, request.italic);
            let resolved_key = (key.clone(), request.scripts.clone());
            if self.resolved.contains(&resolved_key) {
                continue;
            }
            let family = if request.family.trim().is_empty() {
                DEFAULT_LAYOUT_FAMILY
            } else {
                request.family.as_str()
            };
            let faces = manager.resolve(
                family,
                request.bold,
                request.italic,
                &request.scripts,
                LastResortOffice::Word,
            );
            let own = (!request.scripts.is_empty()).then(|| {
                manager.resolve(
                    family,
                    request.bold,
                    request.italic,
                    &[],
                    LastResortOffice::Word,
                )
            });
            for face in faces {
                let base = self.register_bytes(&face.load()?)?;
                let script = own.as_ref().is_some_and(|own| {
                    !own.iter().any(|known| {
                        known.family == face.family
                            && known.bold == face.bold
                            && known.italic == face.italic
                    })
                });
                let id = if script {
                    base
                } else {
                    self.measured_as(base, face.measure_as.as_deref().unwrap_or(family))
                };
                self.insert(key.clone(), id, 2);
            }
            self.resolved.insert(resolved_key);
        }
        Ok(())
    }
}

/// The family layout measures a run with when it names none.
#[cfg(feature = "fonts")]
const DEFAULT_LAYOUT_FAMILY: &str = "Calibri";

fn chain_key(family: &str, bold: bool, italic: bool) -> String {
    format!(
        "{}|{}|{}",
        family.trim().to_lowercase(),
        u8::from(bold),
        u8::from(italic)
    )
}

#[cfg(feature = "fonts")]
pub(crate) fn measurement_widths(input: &docx_layout::types::Input) -> Result<Vec<(f64, bool)>> {
    use docx_layout::prescan::{SectionLayoutConfig, collect_section_configs, default_columns};
    let options = &input.options;
    let body = SectionLayoutConfig {
        page_size: options
            .page_size
            .clone()
            .unwrap_or(docx_layout::types::Size {
                w: 816.0,
                h: 1056.0,
            }),
        margins: docx_layout::section_breaks::resolve_page_margins(options.margins.as_ref()),
        columns: options.columns.clone(),
    };
    let final_section = SectionLayoutConfig {
        page_size: options
            .final_page_size
            .clone()
            .unwrap_or_else(|| body.page_size.clone()),
        margins: options
            .final_margins
            .clone()
            .unwrap_or_else(|| body.margins.clone()),
        columns: options.columns.clone(),
    };
    let (sections, ends) = collect_section_configs(&input.measured, &body, final_section);
    let widths = sections
        .into_iter()
        .map(|section| {
            let columns = section.columns.unwrap_or_else(default_columns);
            let count =
                (columns.count.max(1.0) as usize).max(columns.columns.as_ref().map_or(0, Vec::len));
            let paginator = docx_layout::page_flow::Paginator::new(
                section.page_size,
                section.margins,
                columns,
                None,
            )?;
            let width = paginator.get_column_width(0);
            Ok((
                width,
                (1..count).all(|index| (paginator.get_column_width(index) - width).abs() < 0.01),
            ))
        })
        .collect::<Result<Vec<_>>>()?;
    Ok((0..input.measured.len())
        .map(|index| widths[ends.partition_point(|end| *end < index)])
        .collect())
}

impl Document {
    pub fn register_font(
        &mut self,
        family: &str,
        bold: bool,
        italic: bool,
        bytes: &[u8],
    ) -> Result<u32> {
        self.fonts
            .get_mut()
            .unwrap_or_else(PoisonError::into_inner)
            .register(family, bold, italic, bytes, 1)
    }

    #[cfg(feature = "fonts")]
    pub(crate) fn load_embedded_fonts(&mut self) -> Result<()> {
        let part = |path: &str| {
            self.original_parts
                .iter()
                .find(|(name, _)| name == path)
                .map(|(_, bytes)| bytes.as_slice())
        };
        let fonts = betteroffice_fonts::embedded_font_definitions(
            "word/fontTable.xml",
            part("word/fontTable.xml").unwrap_or_default(),
            part("word/_rels/fontTable.xml.rels").unwrap_or_default(),
            part,
        );
        for font in fonts {
            match self
                .fonts
                .get_mut()
                .unwrap_or_else(PoisonError::into_inner)
                .register(&font.family, font.bold, font.italic, &font.data, 0)
            {
                Ok(_) | Err(Error::Font(_)) => {}
                Err(error) => return Err(error),
            }
        }
        Ok(())
    }

    #[cfg(all(feature = "fonts", feature = "raster"))]
    pub(crate) fn prepare_display_fonts(&self, display: &DisplayList, page: usize) -> Result<()> {
        let Some(page) = display.pages.get(page) else {
            return Ok(());
        };
        let mut requests = Vec::new();
        let mut collect = |primitives: &[Primitive]| {
            requests.extend(primitives.iter().filter_map(|primitive| match primitive {
                Primitive::Text(run) => text_requirement(run),
                _ => None,
            }));
        };
        collect(&page.primitives);
        for parts in [&page.header, &page.footer].into_iter().flatten() {
            collect(&parts.primitives);
        }
        for area in &page.note_areas {
            collect(&area.primitives);
        }
        self.fonts
            .lock()
            .unwrap_or_else(PoisonError::into_inner)
            .resolve(&self.font_manager, &requests)
    }
}

#[cfg(all(feature = "fonts", feature = "raster"))]
fn text_requirement(run: &docx_layout::display_list::TextRunPrimitive) -> Option<FontRequirement> {
    let (style, family) = run.font.split_once("px")?;
    let family = family
        .trim()
        .split(',')
        .next()
        .unwrap_or_default()
        .trim_matches(['\'', '"']);
    let bold = style.split_whitespace().any(|word| {
        matches!(word, "bold" | "bolder") || word.parse::<u16>().is_ok_and(|weight| weight >= 600)
    });
    let italic = style
        .split_whitespace()
        .any(|word| matches!(word, "italic" | "oblique"));
    Some(FontRequirement {
        key: chain_key(family, bold, italic),
        family: family.to_owned(),
        bold,
        italic,
        scripts: text_scripts(run),
        characters: Default::default(),
    })
}

#[cfg(all(feature = "fonts", feature = "raster"))]
fn text_scripts(run: &docx_layout::display_list::TextRunPrimitive) -> Vec<String> {
    let own = ooxml_text::paragraph_scripts_with_han_fallback([run.text.as_str()], false);
    let Some(paragraph) = run.scripts.as_deref().filter(|scripts| !scripts.is_empty()) else {
        return own;
    };
    let regional = paragraph
        .iter()
        .any(|script| matches!(script.as_str(), "cjk-jp" | "cjk-kr" | "cjk-tc"));
    let mut scripts = paragraph.clone();
    for script in own {
        if (regional && script == "cjk-sc") || scripts.contains(&script) {
            continue;
        }
        scripts.push(script);
    }
    scripts
}

#[cfg(all(test, feature = "fonts", feature = "raster"))]
mod tests {
    use super::text_requirement;
    use docx_layout::display_list::TextRunPrimitive;

    #[test]
    fn a_requirement_without_a_family_resolves_through_the_default_family() {
        use super::{FontRegistry, chain_key};
        use docx_layout::measure_blocks::FontRequirement;
        let options = crate::OpenOptions::default();
        let manager =
            betteroffice_fonts::FontManager::new(&options.font_source, options.font_loading)
                .unwrap();
        let requirement = |family: &str| FontRequirement {
            key: chain_key(family, false, false),
            family: family.to_owned(),
            bold: false,
            italic: false,
            scripts: Vec::new(),
            characters: Default::default(),
        };
        let mut empty = FontRegistry::default();
        empty.resolve(&manager, &[requirement("")]).unwrap();
        let mut calibri = FontRegistry::default();
        calibri
            .resolve(&manager, &[requirement("Calibri")])
            .unwrap();
        let chain = &empty.chains[&chain_key("", false, false)];
        assert!(!chain.is_empty());
        assert_eq!(
            chain.len(),
            calibri.chains[&chain_key("Calibri", false, false)].len()
        );
    }

    #[test]
    fn pdfa_fallbacks_register_the_bundled_emoji_face_once() {
        use super::FontRegistry;
        let mut fonts = FontRegistry::default();
        let before = fonts.store.face_count();
        let font_count = fonts.store.font_count();
        fonts.register_pdfa_fallbacks().unwrap();
        assert_eq!(fonts.store.face_count(), before + 1);
        assert!(fonts.pdfa_fallbacks);
        assert!(fonts.chains.is_empty());
        fonts.register_pdfa_fallbacks().unwrap();
        assert_eq!(fonts.store.face_count(), before + 1);
        let registrations = fonts.registrations;
        fonts.release_pdfa_fallbacks(font_count);
        assert!(!fonts.pdfa_fallbacks);
        assert_eq!(fonts.store.face_count(), before);
        assert_eq!(fonts.registrations, registrations.saturating_sub(1));
    }

    fn run(text: &str, scripts: &[&str]) -> TextRunPrimitive {
        serde_json::from_value(serde_json::json!({
            "text": text, "x": 0.0, "baselineY": 0.0, "width": 10.0,
            "font": "bold 16px 'MS Mincho', serif", "color": "#000000", "scripts": scripts
        }))
        .unwrap()
    }

    #[test]
    fn text_requirements_keep_their_own_scripts_next_to_the_paragraph_choice() {
        let field = text_requirement(&run("مرحبا", &["cjk-jp"])).unwrap();
        assert_eq!(field.scripts, ["cjk-jp", "arabic"]);
        let marker = text_requirement(&run("١.", &["cjk-jp"])).unwrap();
        assert_eq!(marker.scripts, ["cjk-jp", "arabic"]);
        assert_eq!(
            text_requirement(&run("骨", &["cjk-jp"])).unwrap().scripts,
            ["cjk-jp"]
        );
        assert_eq!(
            text_requirement(&run("骨", &["cjk-sc"])).unwrap().scripts,
            ["cjk-sc"]
        );
    }

    #[test]
    fn text_requirements_take_the_paragraph_scripts() {
        let japanese = text_requirement(&run("骨", &["cjk-jp"])).unwrap();
        assert_eq!(
            (japanese.family.as_str(), japanese.bold),
            ("MS Mincho", true)
        );
        assert_eq!(japanese.scripts, ["cjk-jp"]);
        assert_eq!(
            text_requirement(&run("漢字", &[])).unwrap().scripts,
            ["cjk-sc"]
        );
        assert!(
            text_requirement(&run("Latin", &[]))
                .unwrap()
                .scripts
                .is_empty()
        );
    }
}
