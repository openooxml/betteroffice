#[cfg(feature = "fonts")]
use betteroffice_xlsx::fonts::FontSource;
#[cfg(all(feature = "fonts", not(feature = "fonts-latin")))]
use betteroffice_xlsx::fonts::{BundledFontOptions, FontPack};
use betteroffice_xlsx::{OpenOptions, Sheet, Workbook, WorkbookModel};

fn fixture() -> Vec<u8> {
    Workbook::from_model(WorkbookModel {
        sheets: vec![Sheet::new("Text")],
        ..Default::default()
    })
    .unwrap()
    .save()
    .unwrap()
}

#[test]
fn opens_with_backstop_options() {
    let options = OpenOptions::default();
    #[cfg(feature = "fonts")]
    let options = OpenOptions {
        font_source: FontSource::Backstop,
        ..options
    };
    let workbook = Workbook::open_with_options(&fixture(), &options).unwrap();
    assert_eq!(workbook.model().sheets.len(), 1);
}

#[cfg(all(feature = "fonts", not(feature = "fonts-latin")))]
#[test]
fn explicitly_missing_pack_returns_the_typed_error() {
    let result = Workbook::open_with_options(
        &fixture(),
        &OpenOptions {
            font_source: FontSource::Bundled(BundledFontOptions {
                packs: vec![FontPack::Latin],
                base_url: None,
            }),
            ..Default::default()
        },
    );
    assert!(matches!(
        result,
        Err(betteroffice_xlsx::Error::FontSource(
            betteroffice_xlsx::fonts::FontError::MissingFontPack {
                pack: FontPack::Latin,
                ..
            }
        ))
    ));
}

#[cfg(not(any(feature = "fonts-latin", feature = "fonts-intl", feature = "fonts-cjk")))]
#[test]
fn open_without_pack_features_still_opens() {
    assert_eq!(Workbook::open(&fixture()).unwrap().model().sheets.len(), 1);
}

#[cfg(all(feature = "fonts-cjk", feature = "raster"))]
mod shared_faces {
    use super::*;
    use betteroffice_xlsx::fonts::{BundledFonts, FontFaceHandle, FontProvider, LastResortOffice};
    use betteroffice_xlsx::{Cell, CellRef, CellValue, SheetId, Viewport};
    use std::sync::{
        Arc,
        atomic::{AtomicUsize, Ordering},
    };
    use xlsx_model::{CellFormat, Font};

    struct SharedFace {
        bytes: Arc<[u8]>,
        loads: Arc<AtomicUsize>,
    }

    impl SharedFace {
        fn handle(&self) -> FontFaceHandle {
            let bytes = Arc::clone(&self.bytes);
            let loads = Arc::clone(&self.loads);
            FontFaceHandle::new("Noto Sans SC", false, false, move || {
                loads.fetch_add(1, Ordering::Relaxed);
                Ok(bytes.to_vec())
            })
        }
    }

    impl FontProvider for SharedFace {
        fn resolve_family(&self, _: &str, _: bool, _: bool) -> Option<FontFaceHandle> {
            None
        }
        fn resolve_script_fallback(&self, _: &str, _: bool, _: bool) -> Option<FontFaceHandle> {
            Some(self.handle())
        }
        fn resolve_glyph_fallback(
            &self,
            _: &str,
            _: bool,
            _: bool,
            _: LastResortOffice,
        ) -> Option<FontFaceHandle> {
            None
        }
        fn resolve_last_resort(
            &self,
            _: &str,
            _: bool,
            _: bool,
            _: LastResortOffice,
        ) -> FontFaceHandle {
            self.handle()
        }
    }

    #[test]
    fn cjk_family_aliases_load_one_shared_source_face() {
        let source = BundledFonts::default()
            .resolve_script_fallback("cjk-sc", false, false)
            .unwrap();
        assert_eq!(source.family, "Noto Sans SC");
        let bytes: Arc<[u8]> = source.load().unwrap().into();
        let loads = Arc::new(AtomicUsize::new(0));
        let mut model = WorkbookModel::default();
        let mut sheet = Sheet::new("Aliases");
        for row in 0..300 {
            let style = model
                .styles
                .intern_cell_format(&CellFormat {
                    font: Font {
                        name: Some(format!("Alias {row}")),
                        ..Default::default()
                    },
                    ..Default::default()
                })
                .unwrap();
            sheet.set_cell(
                CellRef::new(row, 0),
                Cell {
                    value: CellValue::Text {
                        value: "汉".into()
                    },
                    style,
                    ..Default::default()
                },
            );
        }
        model.sheets.push(sheet);
        let input = Workbook::from_model(model).unwrap().save().unwrap();
        let workbook = Workbook::open_with_options(
            &input,
            &OpenOptions {
                font_source: FontSource::Provider(Arc::new(SharedFace {
                    bytes: Arc::clone(&bytes),
                    loads: Arc::clone(&loads),
                })),
                ..Default::default()
            },
        )
        .unwrap();
        let viewport = Viewport {
            x: 0.0,
            y: 0.0,
            width: 160.0,
            height: 7000.0,
        };
        let display = workbook.display_list(&viewport).unwrap();
        let families: std::collections::HashSet<_> = display
            .commands
            .iter()
            .filter_map(|command| match command {
                xlsx_render::DrawCmd::Text {
                    font_family, text, ..
                } if text.as_ref() == "汉" => font_family.as_deref(),
                _ => None,
            })
            .collect();
        assert_eq!(families.len(), 300);
        assert!(
            !workbook
                .render_png_for(SheetId(0), &viewport)
                .unwrap()
                .bytes
                .is_empty()
        );
        assert_eq!(loads.load(Ordering::Relaxed), 1);
        let mut raster = xlsx_raster::RasterFonts::default();
        for family in &families {
            raster
                .register_font(family, false, false, bytes.to_vec(), false, false)
                .unwrap();
        }
        let retained: std::collections::HashSet<_> = families
            .iter()
            .map(|family| {
                let data = raster.font_data(family, false, false).unwrap();
                (data.as_ptr(), data.len())
            })
            .collect();
        assert_eq!(retained.len(), 1);
        assert_eq!(
            retained.iter().map(|(_, len)| len).sum::<usize>(),
            bytes.len()
        );
        assert!(
            !xlsx_raster::render_png_with_fonts(&display, &raster)
                .unwrap()
                .is_empty()
        );
    }
}
