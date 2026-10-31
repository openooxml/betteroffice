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
    registrations: usize,
}

impl FontRegistry {
    fn register(
        &mut self,
        family: &str,
        bold: bool,
        italic: bool,
        bytes: &[u8],
        priority: u8,
    ) -> Result<u32> {
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
        if family.trim().is_empty() {
            return Err(Error::Font("font family is empty".into()));
        }
        let id = self
            .store
            .register_slice(bytes)
            .map_err(|error| Error::Font(error.to_string()))?;
        let key = chain_key(family, bold, italic);
        let chain = self.chains.entry(key.clone()).or_default();
        if priority == 2 && chain.contains(&id) {
            return Ok(id.to_u32());
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
        Ok(id.to_u32())
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
            let faces = manager.resolve(
                &request.family,
                request.bold,
                request.italic,
                &request.scripts,
                LastResortOffice::Word,
            );
            for face in faces {
                self.register(
                    &request.family,
                    request.bold,
                    request.italic,
                    &face.load()?,
                    2,
                )?;
            }
            self.resolved.insert(resolved_key);
        }
        Ok(())
    }
}

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
            for primitive in primitives {
                if let Primitive::Text(run) = primitive
                    && let Some((style, family)) = run.font.split_once("px")
                {
                    let family = family
                        .trim()
                        .split(',')
                        .next()
                        .unwrap_or_default()
                        .trim_matches(['\'', '"']);
                    let bold = style.split_whitespace().any(|word| {
                        matches!(word, "bold" | "bolder")
                            || word.parse::<u16>().is_ok_and(|weight| weight >= 600)
                    });
                    let italic = style
                        .split_whitespace()
                        .any(|word| matches!(word, "italic" | "oblique"));
                    requests.push(FontRequirement {
                        key: chain_key(family, bold, italic),
                        family: family.to_owned(),
                        bold,
                        italic,
                        scripts: ooxml_text::paragraph_scripts_with_han_fallback(
                            [run.text.as_str()],
                            false,
                        ),
                    });
                }
            }
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
