use std::env;
use std::path::PathBuf;

fn main() {
    let root = PathBuf::from(env::var_os("CARGO_MANIFEST_DIR").expect("CARGO_MANIFEST_DIR"));
    let packaged = root.join("assets");
    let assets = if packaged.join("manifest.json").is_file() {
        packaged
    } else {
        root.join("../../packages/fonts-cjk/assets")
            .canonicalize()
            .expect("font pack assets")
    };
    println!("cargo:rerun-if-changed={}", assets.display());
    println!("cargo:rustc-env=FONT_PACK_ASSETS={}", assets.display());
}
