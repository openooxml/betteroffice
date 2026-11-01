#[test]
fn standalone_consumers_keep_facade_defaults_disabled() {
    for (name, manifest, reads_installed_packs) in [
        (
            "native viewer",
            include_str!("../../../apps/native-viewer/Cargo.toml"),
            false,
        ),
        ("fuzz", include_str!("../../../fuzz/Cargo.toml"), false),
        (
            "native document binding",
            include_str!("../../../bindings/docx-native/Cargo.toml"),
            false,
        ),
        (
            "native presentation binding",
            include_str!("../../../bindings/pptx-native/Cargo.toml"),
            false,
        ),
        (
            "native spreadsheet binding",
            include_str!("../../../bindings/xlsx-native/Cargo.toml"),
            false,
        ),
        (
            "document wheel binding",
            include_str!("../../../bindings/python-docx/Cargo.toml"),
            true,
        ),
        (
            "presentation wheel binding",
            include_str!("../../../bindings/python-pptx/Cargo.toml"),
            true,
        ),
        (
            "spreadsheet wheel binding",
            include_str!("../../../bindings/python-xlsx/Cargo.toml"),
            true,
        ),
        (
            "diagram wheel binding",
            include_str!("../../../bindings/python-vsdx/Cargo.toml"),
            false,
        ),
    ] {
        for dependency in manifest.split("\nbetteroffice-").skip(1) {
            let (crate_name, declaration) = dependency.split_once(" = ").unwrap();
            if !matches!(crate_name, "docx" | "pptx" | "xlsx" | "vsdx") {
                continue;
            }
            let declaration = declaration.split_once('}').unwrap().0;
            assert!(
                declaration.contains("default-features = false"),
                "{name} enables default features of {crate_name}"
            );
            for feature in ["\"fonts-latin\"", "\"fonts-intl\"", "\"fonts-cjk\""] {
                assert!(
                    !declaration.contains(feature),
                    "{name} embeds font packs through {crate_name}"
                );
            }
            assert!(
                reads_installed_packs || !declaration.contains("\"fonts"),
                "{name} enables fonts of {crate_name}"
            );
        }
    }
}
