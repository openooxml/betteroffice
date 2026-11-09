use std::collections::{HashMap, HashSet};
use std::sync::{Arc, MutexGuard, PoisonError};

use betteroffice_fonts::LastResortOffice;
use xlsx_raster::RasterFonts;
use xlsx_render::{DisplayList, DrawCmd};

use crate::{Error, PngFontFace, Result, Workbook};

const MAX_FONT_BYTES: usize = 32 * 1024 * 1024;
const MAX_FONT_FACES: usize = 256;
const MAX_FONT_REGISTRATIONS: usize = 4_096;

struct ApplicationFont {
    family: String,
    bold: bool,
    italic: bool,
    data: Arc<[u8]>,
}

#[derive(Default)]
pub(crate) struct FontRegistry {
    fonts: RasterFonts,
    application: Vec<ApplicationFont>,
    source: HashMap<(String, bool, bool), Arc<[u8]>>,
    resolved: HashSet<(String, bool, bool, Vec<String>)>,
}

impl Workbook {
    pub fn register_font(
        &mut self,
        family: &str,
        bold: bool,
        italic: bool,
        bytes: &[u8],
    ) -> Result<()> {
        if bytes.len() > MAX_FONT_BYTES {
            return Err(Error::Raster(format!(
                "font exceeds {MAX_FONT_BYTES} bytes"
            )));
        }
        let registry = self.fonts.get_mut().unwrap_or_else(PoisonError::into_inner);
        if registry.application.len() >= MAX_FONT_REGISTRATIONS {
            return Err(Error::Raster(format!(
                "more than {MAX_FONT_REGISTRATIONS} font registrations"
            )));
        }
        let data = registry
            .application
            .iter()
            .find(|font| font.data.as_ref() == bytes)
            .map_or_else(|| Arc::from(bytes), |font| Arc::clone(&font.data));
        let definition = ApplicationFont {
            family: family.to_owned(),
            data,
            bold,
            italic,
        };
        let mut fonts = RasterFonts::default();
        for font in registry
            .application
            .iter()
            .chain(std::iter::once(&definition))
        {
            fonts
                .register_shared_font(
                    &font.family,
                    font.bold,
                    font.italic,
                    Arc::clone(&font.data),
                    font.bold,
                    font.italic,
                )
                .map_err(Error::Raster)?;
        }
        registry.application.push(definition);
        registry.fonts = fonts;
        registry.source.clear();
        registry.resolved.clear();
        Ok(())
    }

    pub(crate) fn raster_fonts(
        &self,
        display: &DisplayList,
    ) -> Result<MutexGuard<'_, FontRegistry>> {
        let mut registry = self.fonts.lock().unwrap_or_else(PoisonError::into_inner);
        for command in &display.commands {
            if let DrawCmd::Text {
                font_family,
                bold,
                italic,
                text,
                ..
            } = command
            {
                let family = font_family.as_deref().unwrap_or("Calibri");
                let requested = primary_family(family);
                let scripts =
                    ooxml_text::paragraph_scripts_with_han_fallback([text.as_ref()], false);
                let key = (
                    family.trim().to_lowercase(),
                    *bold,
                    *italic,
                    scripts.clone(),
                );
                if registry.resolved.contains(&key) {
                    continue;
                }
                let faces = self.font_manager.resolve(
                    requested,
                    *bold,
                    *italic,
                    &scripts,
                    LastResortOffice::Word,
                );
                for face in faces {
                    if registry.fonts.registration_count() >= MAX_FONT_REGISTRATIONS {
                        break;
                    }
                    let cache_key = (face.family.clone(), face.bold, face.italic);
                    let data = if let Some(data) = registry.source.get(&cache_key) {
                        Arc::clone(data)
                    } else {
                        if registry.fonts.face_count() >= MAX_FONT_FACES {
                            continue;
                        }
                        let bytes = face.load()?;
                        if bytes.len() > MAX_FONT_BYTES {
                            continue;
                        }
                        Arc::from(bytes)
                    };
                    if registry
                        .fonts
                        .register_shared_font(
                            family,
                            *bold,
                            *italic,
                            Arc::clone(&data),
                            face.bold,
                            face.italic,
                        )
                        .is_ok()
                    {
                        let data = registry
                            .fonts
                            .shared_font_data(&data)
                            .expect("registered font");
                        registry.source.entry(cache_key).or_insert(data);
                    }
                }
                registry.resolved.insert(key);
            }
        }
        Ok(registry)
    }
}

/// The family a CSS-style list such as a chart's `"Calibri", sans-serif` asks for first, as the canvas reads it.
fn primary_family(family: &str) -> &str {
    let head = family.split(',').next().unwrap_or(family).trim();
    let head = head
        .strip_prefix('"')
        .and_then(|inner| inner.strip_suffix('"'))
        .or_else(|| {
            head.strip_prefix('\'')
                .and_then(|inner| inner.strip_suffix('\''))
        })
        .unwrap_or(head);
    if head.is_empty() { family } else { head }
}

/// The family a run can take a supplied face for: its own name, or a chart's first named family.
fn suppliable_family(family: &str, chart: bool) -> Option<&str> {
    const GENERIC: [&str; 9] = [
        "serif",
        "sans-serif",
        "monospace",
        "cursive",
        "fantasy",
        "system-ui",
        "ui-sans-serif",
        "ui-serif",
        "ui-monospace",
    ];
    if family.is_empty() {
        return None;
    }
    if !chart {
        return Some(family);
    }
    let head = primary_family(family);
    (!head.is_empty() && !GENERIC.contains(&head.to_lowercase().as_str())).then_some(head)
}

/// The supplied face for a font: its family, dropping bold before italic, the last of equal faces winning.
fn supplied_face<'a, 'b>(
    faces: &'b [PngFontFace<'a>],
    family: &str,
    bold: bool,
    italic: bool,
) -> Option<&'b PngFontFace<'a>> {
    let family = family.trim().to_lowercase();
    [
        (bold, italic),
        (bold, false),
        (false, italic),
        (false, false),
    ]
    .into_iter()
    .find_map(|(bold, italic)| {
        faces.iter().rev().find(|face| {
            face.family.trim().to_lowercase() == family
                && face.bold == bold
                && face.italic == italic
        })
    })
}

/// The fonts for an export given `faces`: those alone, as the canvas export paints without the bundled faces.
pub(crate) fn supplied_fonts(
    display: &DisplayList,
    faces: &[PngFontFace<'_>],
) -> Result<RasterFonts> {
    let mut fonts = RasterFonts::default();
    let mut seen = HashSet::new();
    for command in &display.commands {
        let DrawCmd::Text {
            font_family: Some(family),
            bold,
            italic,
            chart,
            ..
        } = command
        else {
            continue;
        };
        if !seen.insert((family.trim().to_lowercase(), *bold, *italic)) {
            continue;
        }
        if let Some(face) = suppliable_family(family, *chart)
            .and_then(|name| supplied_face(faces, name, *bold, *italic))
        {
            fonts
                .register_font(
                    family,
                    *bold,
                    *italic,
                    face.bytes.to_vec(),
                    face.bold,
                    face.italic,
                )
                .map_err(|_| Error::FontUnavailable {
                    family: face.family.to_owned(),
                })?;
        }
    }
    Ok(fonts)
}

impl FontRegistry {
    pub(crate) fn fonts(&self) -> &RasterFonts {
        &self.fonts
    }
}

#[cfg(test)]
mod tests {
    #[test]
    fn supplied_faces_match_as_the_canvas_export_does() {
        let face = |family, bold, italic, bytes| PngFontFace {
            family,
            bold,
            italic,
            bytes,
        };
        let faces = [
            face("Body", false, false, &b"regular"[..]),
            face(" body ", false, true, &b"italic"[..]),
            face("BODY", false, false, &b"last regular"[..]),
        ];
        let chosen = |family, bold, italic| {
            supplied_face(&faces, family, bold, italic).map(|face| face.bytes)
        };
        assert_eq!(chosen("body", false, false), Some(&b"last regular"[..]));
        assert_eq!(chosen("Body", true, true), Some(&b"italic"[..]));
        assert_eq!(chosen("Body", true, false), Some(&b"last regular"[..]));
        assert_eq!(chosen("Other", false, false), None);
    }

    #[test]
    fn charts_supply_their_first_named_family_only() {
        assert_eq!(
            suppliable_family("\"Aptos\", sans-serif", true),
            Some("Aptos")
        );
        assert_eq!(suppliable_family("sans-serif, Aptos", true), None);
        assert_eq!(suppliable_family("A, B", false), Some("A, B"));
        assert_eq!(suppliable_family("", false), None);
    }

    use super::*;
    use crate::{Cell, CellRef, CellValue, OpenOptions, Sheet, SheetId, Viewport, WorkbookModel};
    #[cfg(feature = "fonts-latin")]
    use betteroffice_fonts::{BundledFonts, FontProvider};
    use betteroffice_fonts::{FontSource, OpenOptions as FontOptions};
    use xlsx_model::{CellFormat, Font};

    fn fixture() -> Vec<u8> {
        let mut model = WorkbookModel::default();
        let style = model
            .styles
            .intern_cell_format(&CellFormat {
                font: Font {
                    name: Some("Calibri".into()),
                    bold: true,
                    ..Default::default()
                },
                ..Default::default()
            })
            .unwrap();
        let mut sheet = Sheet::new("Text");
        sheet.set_cell(
            CellRef::new(0, 0),
            Cell {
                value: CellValue::Text {
                    value: "Office".into(),
                },
                style,
                ..Default::default()
            },
        );
        model.sheets.push(sheet);
        Workbook::from_model(model).unwrap().save().unwrap()
    }

    fn viewport() -> Viewport {
        Viewport {
            x: 0.0,
            y: 0.0,
            width: 160.0,
            height: 48.0,
        }
    }

    #[cfg(feature = "fonts-latin")]
    #[test]
    fn bold_calibri_uses_the_real_carlito_bold_face() {
        let workbook = Workbook::open(&fixture()).unwrap();
        let display = workbook.display_list(&viewport()).unwrap();
        assert!(
            display
                .commands
                .iter()
                .any(|command| matches!(command, DrawCmd::Text { bold: true, .. }))
        );
        let fonts = workbook.raster_fonts(&display).unwrap();
        let expected = BundledFonts::default()
            .resolve_family("Calibri", true, false)
            .unwrap();
        assert_eq!(expected.family, "Carlito");
        assert!(expected.bold);
        let bold = expected.load().unwrap();
        let regular = BundledFonts::default()
            .resolve_family("Calibri", false, false)
            .unwrap()
            .load()
            .unwrap();
        assert_ne!(bold, regular);
        assert_eq!(fonts.fonts.font_data("Calibri", true, false).unwrap(), bold);
    }

    #[cfg(feature = "fonts-latin")]
    #[test]
    fn a_family_list_resolves_its_first_family() {
        let workbook = Workbook::open(&fixture()).unwrap();
        let mut display = workbook.display_list(&viewport()).unwrap();
        for command in &mut display.commands {
            if let DrawCmd::Text { font_family, .. } = command {
                *font_family = Some("\"Calibri\", sans-serif".into());
            }
        }
        let fonts = workbook.raster_fonts(&display).unwrap();
        let bold = BundledFonts::default()
            .resolve_family("Calibri", true, false)
            .unwrap()
            .load()
            .unwrap();
        assert_eq!(
            fonts
                .fonts
                .font_data("\"Calibri\", sans-serif", true, false)
                .unwrap(),
            bold
        );
    }

    #[test]
    fn backstop_preserves_the_embedded_carlito_raster() {
        let workbook = Workbook::open_with_options(
            &fixture(),
            &OpenOptions {
                font_source: FontSource::Backstop,
                ..FontOptions::default()
            },
        )
        .unwrap();
        let display = workbook.display_list(&viewport()).unwrap();
        assert!(
            workbook
                .raster_fonts(&display)
                .unwrap()
                .fonts
                .font_data("Calibri", true, false)
                .is_none()
        );
        assert_eq!(
            workbook
                .render_png(
                    crate::PngTarget::Viewport {
                        sheet: SheetId(0),
                        viewport: viewport()
                    },
                    &crate::RenderOptions::default()
                )
                .unwrap()
                .bytes,
            xlsx_raster::render_png(&display).unwrap()
        );
    }

    #[test]
    fn supplied_font_wins_over_the_pack() {
        let bytes = include_bytes!("../../ooxml-text/tests/fonts/LiberationSans-Regular.ttf");
        let mut workbook = Workbook::open(&fixture()).unwrap();
        let display = workbook.display_list(&viewport()).unwrap();
        drop(workbook.raster_fonts(&display).unwrap());
        workbook
            .register_font("Calibri", true, false, bytes)
            .unwrap();
        let display = workbook.display_list(&viewport()).unwrap();
        let fonts = workbook.raster_fonts(&display).unwrap();
        assert_eq!(
            fonts.fonts.font_data("Calibri", true, false).unwrap(),
            bytes
        );
    }
}
