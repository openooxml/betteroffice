#![cfg(feature = "manager")]

use std::path::PathBuf;
use std::sync::Arc;

use betteroffice_fonts::{
    BundledFontOptions, BundledFonts, FontError, FontFaceHandle, FontLoading, FontManager,
    FontPack, FontProvider, FontSource, LastResortOffice, resolve_metric_compat_face,
};

fn local_provider(root: PathBuf) -> BundledFonts {
    BundledFonts::new(BundledFontOptions {
        packs: vec![FontPack::Latin],
        base_url: Some(root),
    })
    .unwrap()
}

fn deploy(root: &std::path::Path, file: &str) {
    let face = resolve_metric_compat_face("Calibri", false, false).unwrap();
    let dir = root.join("latin").join(&face.version);
    std::fs::create_dir_all(&dir).unwrap();
    std::fs::write(
        dir.join(file),
        betteroffice_fonts_latin::face(file).unwrap(),
    )
    .unwrap();
}

#[test]
fn loads_local_faces_and_dictionary_payloads_into_owned_buffers() {
    let root = tempfile::tempdir().unwrap();
    for file in ["Carlito-Regular.ttf.zst", "Carlito-Bold.ttf.zst"] {
        deploy(root.path(), file);
    }
    let provider = local_provider(root.path().to_path_buf());
    for bold in [false, true] {
        let handle = provider.resolve_family("Calibri", bold, false).unwrap();
        let mut first = handle.load().unwrap();
        assert_eq!(&first[..4], &[0, 1, 0, 0]);
        assert_eq!(
            first.len(),
            resolve_metric_compat_face("Calibri", bold, false)
                .unwrap()
                .byte_length
        );
        let second = handle.load().unwrap();
        first[0] = 42;
        assert_eq!(second[0], 0);
    }
}

#[cfg(feature = "fonts-latin")]
#[test]
fn decodes_embedded_carlito_and_its_dictionary_faces() {
    let provider = BundledFonts::default();
    for bold in [false, true] {
        let handle = provider.resolve_family("Calibri", bold, false).unwrap();
        let bytes = handle.load().unwrap();
        assert_eq!(&bytes[..4], &[0, 1, 0, 0]);
        assert_eq!(
            bytes.len(),
            resolve_metric_compat_face("Calibri", bold, false)
                .unwrap()
                .byte_length
        );
    }
}

#[cfg(not(feature = "fonts-latin"))]
#[test]
fn absent_embedded_pack_names_the_cargo_feature() {
    assert!(
        BundledFonts::default()
            .resolve_family("Calibri", false, false)
            .is_none()
    );
    match BundledFonts::new(BundledFontOptions {
        packs: vec![FontPack::Latin],
        base_url: None,
    })
    .unwrap_err()
    {
        FontError::MissingFontPack { pack, action } => {
            assert_eq!(pack, FontPack::Latin);
            assert!(action.contains("fonts-latin"));
        }
        error => panic!("{error}"),
    }
}

#[test]
fn local_root_never_falls_back_to_embedded_assets() {
    let root = tempfile::tempdir().unwrap();
    let provider = local_provider(root.path().to_path_buf());
    let handle = provider.resolve_family("Calibri", false, false).unwrap();
    match handle.load().unwrap_err() {
        FontError::MissingFontPack { pack, action } => {
            assert_eq!(pack, FontPack::Latin);
            assert!(action.contains(&root.path().display().to_string()));
            assert!(action.contains("Carlito-Regular.ttf.zst"));
        }
        error => panic!("{error}"),
    }
    deploy(root.path(), "Carlito-Regular.ttf.zst");
    assert!(handle.load().is_ok());
}

#[test]
fn roots_and_pack_selections_do_not_share_cached_bytes() {
    let first = tempfile::tempdir().unwrap();
    let second = tempfile::tempdir().unwrap();
    deploy(first.path(), "Carlito-Regular.ttf.zst");
    let provider = local_provider(first.path().to_path_buf());
    assert!(
        provider
            .resolve_family("Calibri", false, false)
            .unwrap()
            .load()
            .is_ok()
    );
    let empty = local_provider(second.path().to_path_buf());
    assert!(matches!(
        empty
            .resolve_family("Calibri", false, false)
            .unwrap()
            .load(),
        Err(FontError::MissingFontPack { .. })
    ));
    let disabled = BundledFonts::new(BundledFontOptions {
        packs: vec![],
        base_url: Some(first.path().to_path_buf()),
    })
    .unwrap();
    assert!(disabled.resolve_family("Calibri", false, false).is_none());
    assert!(matches!(
        disabled.load_face(resolve_metric_compat_face("Calibri", false, false).unwrap()),
        Err(FontError::MissingFontPack { .. })
    ));
    assert!(
        provider
            .resolve_family("Noto Sans Thai", false, false)
            .is_none()
    );
    assert!(
        provider
            .resolve_script_fallback("thai", false, false)
            .is_none()
    );
}

#[test]
fn rejects_invalid_options_and_reports_corrupt_local_data() {
    assert!(matches!(
        BundledFonts::new(BundledFontOptions {
            packs: vec![FontPack::Latin, FontPack::Latin],
            base_url: None,
        }),
        Err(FontError::InvalidFontOptions(_))
    ));
    for root in ["", "https://invalid/fonts"] {
        assert!(matches!(
            BundledFonts::new(BundledFontOptions {
                packs: vec![FontPack::Latin],
                base_url: Some(root.into()),
            }),
            Err(FontError::InvalidFontOptions(_))
        ));
    }
    let root = tempfile::tempdir().unwrap();
    deploy(root.path(), "Carlito-Regular.ttf.zst");
    let face = resolve_metric_compat_face("Calibri", false, false).unwrap();
    std::fs::write(
        root.path()
            .join("latin")
            .join(&face.version)
            .join(&face.file),
        [0],
    )
    .unwrap();
    assert!(matches!(
        local_provider(root.path().to_path_buf())
            .resolve_family("Calibri", false, false)
            .unwrap()
            .load(),
        Err(FontError::InvalidFontData(_))
    ));
}

struct SuppliedProvider;

impl FontProvider for SuppliedProvider {
    fn resolve_family(&self, family: &str, bold: bool, italic: bool) -> Option<FontFaceHandle> {
        Some(FontFaceHandle::new(family, bold, italic, || {
            Ok(vec![1, 2, 3])
        }))
    }

    fn resolve_script_fallback(
        &self,
        script: &str,
        bold: bool,
        italic: bool,
    ) -> Option<FontFaceHandle> {
        self.resolve_family(script, bold, italic)
    }

    fn resolve_glyph_fallback(
        &self,
        family: &str,
        bold: bool,
        italic: bool,
        _office: LastResortOffice,
    ) -> Option<FontFaceHandle> {
        self.resolve_family(family, bold, italic)
    }

    fn resolve_last_resort(
        &self,
        family: &str,
        bold: bool,
        italic: bool,
        _office: LastResortOffice,
    ) -> FontFaceHandle {
        self.resolve_family(family, bold, italic).unwrap()
    }
}

#[test]
fn provider_source_is_used_verbatim() {
    let provider: Arc<dyn FontProvider> = Arc::new(SuppliedProvider);
    let resolved = FontSource::Provider(Arc::clone(&provider))
        .provider()
        .unwrap()
        .unwrap();
    assert!(Arc::ptr_eq(&provider, &resolved));
    let handle = resolved.resolve_family("App Face", true, true).unwrap();
    assert_eq!(handle.family, "App Face");
    assert!(handle.bold && handle.italic);
    assert_eq!(handle.load().unwrap(), [1, 2, 3]);
    assert_eq!(
        resolved
            .resolve_script_fallback("thai", false, false)
            .unwrap()
            .load()
            .unwrap(),
        [1, 2, 3]
    );
    assert_eq!(
        resolved
            .resolve_glyph_fallback("App Face", false, false, LastResortOffice::Word)
            .unwrap()
            .load()
            .unwrap(),
        [1, 2, 3]
    );
    assert_eq!(
        resolved
            .resolve_last_resort("App Face", false, false, LastResortOffice::Word)
            .load()
            .unwrap(),
        [1, 2, 3]
    );
}

#[test]
fn defaults_select_compiled_local_packs() {
    assert!(
        matches!(FontSource::default(), FontSource::Bundled(options) if options == BundledFontOptions::default())
    );
    assert!(FontSource::Backstop.provider().unwrap().is_none());
    assert_eq!(
        BundledFontOptions::default().packs,
        [FontPack::Latin, FontPack::Intl, FontPack::Cjk]
            .into_iter()
            .filter(|pack| pack.is_compiled_in())
            .collect::<Vec<_>>()
    );
}

#[test]
fn local_dictionary_must_be_deployed_even_with_embedded_fonts() {
    let root = tempfile::tempdir().unwrap();
    deploy(root.path(), "Carlito-Bold.ttf.zst");
    let provider = local_provider(root.path().to_path_buf());
    let handle = provider.resolve_family("Calibri", true, false).unwrap();
    assert!(
        matches!(handle.load(), Err(FontError::MissingFontPack { action, .. }) if action.contains("Carlito-Regular.ttf.zst"))
    );
    deploy(root.path(), "Carlito-Regular.ttf.zst");
    assert!(handle.load().is_ok());
}

#[test]
fn concurrent_loaders_share_a_face_and_return_owned_bytes() {
    let root = tempfile::tempdir().unwrap();
    deploy(root.path(), "Carlito-Regular.ttf.zst");
    let provider = local_provider(root.path().to_path_buf());
    let handles: Vec<_> = (0..4)
        .map(|_| {
            let face = provider.resolve_family("Calibri", false, false).unwrap();
            std::thread::spawn(move || face.load().unwrap())
        })
        .collect();
    let bytes: Vec<_> = handles
        .into_iter()
        .map(|handle| handle.join().unwrap())
        .collect();
    assert!(
        bytes
            .windows(2)
            .all(|pair| pair[0] == pair[1] && pair[0].as_ptr() != pair[1].as_ptr())
    );
}

#[test]
fn decode_errors_are_reported_and_retryable() {
    let root = tempfile::tempdir().unwrap();
    deploy(root.path(), "Carlito-Regular.ttf.zst");
    let face = resolve_metric_compat_face("Calibri", false, false).unwrap();
    let path = root
        .path()
        .join("latin")
        .join(&face.version)
        .join(&face.file);
    let mut bytes = std::fs::read(&path).unwrap();
    *bytes.last_mut().unwrap() ^= 1;
    std::fs::write(&path, bytes).unwrap();
    let provider = local_provider(root.path().to_path_buf());
    let handle = provider.resolve_family("Calibri", false, false).unwrap();
    assert!(matches!(handle.load(), Err(FontError::Decode(_))));
    deploy(root.path(), "Carlito-Regular.ttf.zst");
    assert!(handle.load().is_ok());
}

#[test]
fn preload_obeys_the_selected_packs() {
    let provider = BundledFonts::new(BundledFontOptions {
        packs: vec![],
        base_url: None,
    })
    .unwrap();
    assert!(provider.preload().is_ok());
    let root = tempfile::tempdir().unwrap();
    let provider = local_provider(root.path().to_path_buf());
    assert!(matches!(
        provider.preload(),
        Err(FontError::MissingFontPack {
            pack: FontPack::Latin,
            ..
        })
    ));
}

#[test]
fn on_demand_defers_local_reads_and_preload_reads_at_open() {
    let root = tempfile::tempdir().unwrap();
    let source = FontSource::Bundled(BundledFontOptions {
        packs: vec![FontPack::Latin],
        base_url: Some(root.path().to_owned()),
    });
    let manager = FontManager::new(&source, FontLoading::OnDemand).unwrap();
    let faces = manager.resolve("Calibri", false, false, &[], LastResortOffice::Word);
    assert_eq!(faces[0].family, "Carlito");
    assert!(matches!(
        faces[0].load(),
        Err(FontError::MissingFontPack {
            pack: FontPack::Latin,
            ..
        })
    ));
    assert!(matches!(
        FontManager::new(&source, FontLoading::Preload),
        Err(FontError::MissingFontPack {
            pack: FontPack::Latin,
            ..
        })
    ));
}
