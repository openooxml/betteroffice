#![cfg(feature = "manager")]

use std::sync::Arc;

use betteroffice_fonts::{
    FontFaceHandle, FontLoading, FontManager, FontProvider, FontSource, LastResortOffice,
    set_default_font_source,
};

struct Fixed;

impl FontProvider for Fixed {
    fn resolve_family(&self, family: &str, bold: bool, italic: bool) -> Option<FontFaceHandle> {
        Some(FontFaceHandle::new(
            format!("fixed {family}"),
            bold,
            italic,
            || Ok(vec![1]),
        ))
    }

    fn resolve_script_fallback(&self, _: &str, _: bool, _: bool) -> Option<FontFaceHandle> {
        None
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
        bold: bool,
        italic: bool,
        _: LastResortOffice,
    ) -> FontFaceHandle {
        FontFaceHandle::new("fixed last resort", bold, italic, || Ok(vec![2]))
    }
}

#[test]
fn the_first_process_default_wins_for_every_default_source() {
    assert!(matches!(FontSource::default(), FontSource::Bundled(_)));
    assert!(set_default_font_source(FontSource::Provider(Arc::new(Fixed))).is_ok());
    assert!(matches!(
        set_default_font_source(FontSource::Backstop),
        Err(FontSource::Backstop)
    ));
    let manager = FontManager::new(&FontSource::default(), FontLoading::OnDemand).unwrap();
    let faces = manager.resolve("Calibri", false, false, &[], LastResortOffice::Word);
    let families: Vec<_> = faces.iter().map(|face| face.family.as_str()).collect();
    assert_eq!(families, ["fixed Calibri", "fixed last resort"]);
}
