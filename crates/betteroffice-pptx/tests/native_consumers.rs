#[test]
fn standalone_consumers_keep_facade_defaults_disabled() {
    for (name, manifest) in [
        (
            "native viewer",
            include_str!("../../../apps/native-viewer/Cargo.toml"),
        ),
        ("fuzz", include_str!("../../../fuzz/Cargo.toml")),
        (
            "native document binding",
            include_str!("../../../bindings/docx-native/Cargo.toml"),
        ),
        (
            "native presentation binding",
            include_str!("../../../bindings/pptx-native/Cargo.toml"),
        ),
        (
            "native spreadsheet binding",
            include_str!("../../../bindings/xlsx-native/Cargo.toml"),
        ),
        (
            "document wheel binding",
            include_str!("../../../bindings/python-docx/Cargo.toml"),
        ),
        (
            "presentation wheel binding",
            include_str!("../../../bindings/python-pptx/Cargo.toml"),
        ),
        (
            "spreadsheet wheel binding",
            include_str!("../../../bindings/python-xlsx/Cargo.toml"),
        ),
        (
            "diagram wheel binding",
            include_str!("../../../bindings/python-vsdx/Cargo.toml"),
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
            assert!(
                !declaration.contains("\"fonts"),
                "{name} enables fonts of {crate_name}"
            );
        }
    }
}
