#![cfg(feature = "manager")]

use std::path::PathBuf;
use std::sync::Arc;

use betteroffice_fonts::{
    BundledFontOptions, BundledFonts, FontError, FontFaceHandle, FontLoading, FontManager,
    FontPack, FontPackDiagnostic, FontPackProblem, FontProvider, FontSource, LastResortOffice,
    resolve_metric_compat_face,
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

fn deploy_flat(root: &std::path::Path) {
    let face = resolve_metric_compat_face("Calibri", false, false).unwrap();
    std::fs::write(
        root.join("manifest.json"),
        serde_json::to_vec(&serde_json::json!({ "version": face.version })).unwrap(),
    )
    .unwrap();
    for face in &betteroffice_fonts::catalog::catalog().faces {
        if face.pack == FontPack::Latin {
            std::fs::write(
                root.join(&face.file),
                betteroffice_fonts_latin::face(&face.file).unwrap(),
            )
            .unwrap();
        }
    }
}

#[test]
fn flat_pack_directories_load_faces_and_dictionaries() {
    let root = tempfile::tempdir().unwrap();
    deploy_flat(root.path());
    let provider =
        BundledFonts::from_pack_directories(vec![(FontPack::Latin, root.path().into())]).unwrap();
    provider.preload().unwrap();
    for bold in [false, true] {
        let face = resolve_metric_compat_face("Calibri", bold, false).unwrap();
        let bytes = provider
            .resolve_family("Calibri", bold, false)
            .unwrap()
            .load()
            .unwrap();
        assert_eq!(&bytes[..4], &[0, 1, 0, 0]);
        assert_eq!(bytes.len(), face.byte_length);
        assert_eq!(provider.load_face(face).unwrap(), bytes);
    }
}

#[test]
fn flat_pack_directories_skip_missing_or_mismatched_manifests() {
    let root = tempfile::tempdir().unwrap();
    deploy_flat(root.path());
    let manifest = root.path().join("manifest.json");
    std::fs::write(&manifest, br#"{"version":"0.0.0-mismatch"}"#).unwrap();
    for missing in [false, true] {
        if missing {
            std::fs::remove_file(&manifest).unwrap();
        }
        let provider =
            BundledFonts::from_pack_directories(vec![(FontPack::Latin, root.path().into())])
                .unwrap();
        assert!(provider.resolve_family("Calibri", false, false).is_none());
        assert!(matches!(
            provider.load_face(resolve_metric_compat_face("Calibri", false, false).unwrap()),
            Err(FontError::MissingFontPack { .. })
        ));
        let expected = if missing {
            FontPackProblem::MissingFile {
                file: "manifest.json".into(),
                missing: 1,
            }
        } else {
            FontPackProblem::VersionMismatch {
                expected: resolve_metric_compat_face("Calibri", false, false)
                    .unwrap()
                    .version
                    .clone(),
                found: Some("0.0.0-mismatch".into()),
            }
        };
        assert_eq!(
            provider.pack_diagnostics(),
            [FontPackDiagnostic {
                pack: FontPack::Latin,
                directory: root.path().into(),
                problem: expected,
            }]
        );
    }
}

#[test]
fn flat_pack_directories_skip_unreadable_manifests_and_incomplete_packs() {
    let unreadable = tempfile::tempdir().unwrap();
    deploy_flat(unreadable.path());
    std::fs::remove_file(unreadable.path().join("manifest.json")).unwrap();
    std::fs::create_dir(unreadable.path().join("manifest.json")).unwrap();
    let incomplete = tempfile::tempdir().unwrap();
    deploy_flat(incomplete.path());
    let face = resolve_metric_compat_face("Calibri", true, false).unwrap();
    std::fs::remove_file(incomplete.path().join(&face.file)).unwrap();
    for root in [unreadable.path(), incomplete.path()] {
        let provider =
            BundledFonts::from_pack_directories(vec![(FontPack::Latin, root.into())]).unwrap();
        assert!(provider.resolve_family("Calibri", false, false).is_none());
        let [diagnostic] = provider.pack_diagnostics() else {
            panic!("one diagnostic: {:?}", provider.pack_diagnostics());
        };
        assert_eq!(diagnostic.pack, FontPack::Latin);
        if root == unreadable.path() {
            assert_eq!(diagnostic.reason(), "invalid-manifest");
        } else {
            assert_eq!(
                diagnostic.problem,
                FontPackProblem::MissingFile {
                    file: face.file.clone(),
                    missing: 1
                }
            );
            assert_eq!(
                diagnostic.to_string(),
                format!(
                    "@betteroffice/fonts-latin is missing {}; its faces are substituted",
                    face.file
                )
            );
        }
    }
}

#[test]
fn pack_directories_that_do_not_exist_report_every_file_missing() {
    let root = tempfile::tempdir().unwrap();
    let absent = root.path().join("absent");
    let provider =
        BundledFonts::from_pack_directories(vec![(FontPack::Latin, absent.clone())]).unwrap();
    assert!(provider.resolve_family("Calibri", false, false).is_none());
    let mut files: Vec<&str> = betteroffice_fonts::catalog::catalog()
        .faces
        .iter()
        .filter(|face| face.pack == FontPack::Latin)
        .map(|face| face.file.as_str())
        .chain(["manifest.json"])
        .collect();
    files.sort_unstable();
    assert_eq!(
        provider.pack_diagnostics(),
        [FontPackDiagnostic {
            pack: FontPack::Latin,
            directory: absent,
            problem: FontPackProblem::MissingFile {
                file: files[0].to_string(),
                missing: files.len(),
            },
        }]
    );
    deploy_flat(root.path());
    let provider =
        BundledFonts::from_pack_directories(vec![(FontPack::Latin, root.path().into())]).unwrap();
    assert!(provider.pack_diagnostics().is_empty());
}

#[test]
fn provider_managers_offer_only_faces_of_found_packs() {
    let none = BundledFonts::from_pack_directories(Vec::new()).unwrap();
    let manager =
        FontManager::new(&FontSource::Provider(Arc::new(none)), FontLoading::OnDemand).unwrap();
    for (family, scripts) in [("Calibri", vec![]), ("MS Mincho", vec!["Hani".to_string()])] {
        assert!(
            manager
                .resolve(family, false, false, &scripts, LastResortOffice::Word)
                .is_empty()
        );
    }
    let root = tempfile::tempdir().unwrap();
    deploy_flat(root.path());
    let latin =
        BundledFonts::from_pack_directories(vec![(FontPack::Latin, root.path().into())]).unwrap();
    let manager = FontManager::new(
        &FontSource::Provider(Arc::new(latin)),
        FontLoading::OnDemand,
    )
    .unwrap();
    for family in ["Calibri", "MS Mincho", "Unknown Family"] {
        let faces = manager.resolve(
            family,
            false,
            false,
            &["Hani".to_string(), "Deva".to_string()],
            LastResortOffice::Word,
        );
        assert!(family != "Calibri" || !faces.is_empty());
        for face in faces {
            assert!(face.load().is_ok(), "{} failed to load", face.family);
        }
    }
}

#[test]
fn flat_pack_directories_reject_duplicate_packs_and_empty_paths() {
    let root = tempfile::tempdir().unwrap();
    assert!(matches!(
        BundledFonts::from_pack_directories(vec![
            (FontPack::Latin, root.path().into()),
            (FontPack::Latin, root.path().into()),
        ]),
        Err(FontError::InvalidFontOptions(_))
    ));
    assert!(matches!(
        BundledFonts::from_pack_directories(vec![(FontPack::Latin, PathBuf::new())]),
        Err(FontError::InvalidFontOptions(_))
    ));
}

#[test]
fn flat_pack_directories_resolve_relative_paths() {
    let current = std::env::current_dir().unwrap();
    let root = tempfile::tempdir_in(&current).unwrap();
    deploy_flat(root.path());
    let relative = root.path().strip_prefix(&current).unwrap().to_path_buf();
    let provider = BundledFonts::from_pack_directories(vec![(FontPack::Latin, relative)]).unwrap();
    assert!(
        provider
            .resolve_family("Calibri", false, false)
            .unwrap()
            .load()
            .is_ok()
    );
}

#[test]
fn flat_pack_directories_do_not_share_cached_bytes() {
    let first = tempfile::tempdir().unwrap();
    let second = tempfile::tempdir().unwrap();
    deploy_flat(first.path());
    deploy_flat(second.path());
    let face = resolve_metric_compat_face("Calibri", false, false).unwrap();
    let provider =
        BundledFonts::from_pack_directories(vec![(FontPack::Latin, first.path().into())]).unwrap();
    assert!(provider.load_face(face).is_ok());
    std::fs::write(second.path().join(&face.file), [0]).unwrap();
    let corrupt =
        BundledFonts::from_pack_directories(vec![(FontPack::Latin, second.path().into())]).unwrap();
    assert!(matches!(
        corrupt.load_face(face),
        Err(FontError::InvalidFontData(_))
    ));
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

#[test]
fn naskh_styled_families_chain_noto_naskh_arabic_as_their_arabic_fallback() {
    let root = tempfile::tempdir().unwrap();
    deploy_flat(root.path());
    let latin =
        BundledFonts::from_pack_directories(vec![(FontPack::Latin, root.path().into())]).unwrap();
    let manager = FontManager::new(
        &FontSource::Provider(Arc::new(latin)),
        FontLoading::OnDemand,
    )
    .unwrap();
    let arabic = |family: &str, bold: bool| {
        manager
            .resolve(
                family,
                bold,
                false,
                &["arabic".to_string()],
                LastResortOffice::Word,
            )
            .into_iter()
            .filter(|face| face.family.contains("Arabic"))
            .map(|face| (face.family, face.bold))
            .collect::<Vec<_>>()
    };
    for family in ["Simplified Arabic", "Traditional Arabic"] {
        for bold in [false, true] {
            assert_eq!(
                arabic(family, bold),
                [("Noto Naskh Arabic".to_string(), false)]
            );
        }
    }
    for family in ["Arial", "Times New Roman"] {
        for bold in [false, true] {
            assert_eq!(
                arabic(family, bold),
                [("Noto Sans Arabic".to_string(), bold)]
            );
        }
    }
}

#[test]
fn native_chains_match_the_shared_arabic_fallback_cases() {
    let root = tempfile::tempdir().unwrap();
    deploy_flat(root.path());
    let latin =
        BundledFonts::from_pack_directories(vec![(FontPack::Latin, root.path().into())]).unwrap();
    let manager = FontManager::new(
        &FontSource::Provider(Arc::new(latin)),
        FontLoading::OnDemand,
    )
    .unwrap();
    let cases: Vec<serde_json::Value> =
        serde_json::from_str(include_str!("../data/script-fallback-chains.json")).unwrap();
    for case in &cases {
        let family = case["family"].as_str().unwrap();
        let bold = case["bold"].as_bool().unwrap();
        let expected = betteroffice_fonts::catalog::catalog()
            .faces
            .iter()
            .find(|face| face.file == case["arabic"].as_str().unwrap())
            .unwrap();
        for office in [LastResortOffice::Word, LastResortOffice::PowerPoint] {
            let arabic: Vec<_> = manager
                .resolve(family, bold, false, &["arabic".to_string()], office)
                .into_iter()
                .filter(|face| face.family.contains("Arabic"))
                .map(|face| (face.family, face.bold))
                .collect();
            assert_eq!(
                arabic,
                [(expected.family.clone(), expected.weight == 700)],
                "{case}"
            );
        }
    }
}
