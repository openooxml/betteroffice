use std::env;
use std::fs;
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
    let mut manifest: serde_json::Value = serde_json::from_slice(
        &fs::read(assets.join("manifest.json")).expect("font pack manifest"),
    )
    .expect("valid font pack manifest");
    manifest["faces"]
        .as_array_mut()
        .expect("font faces")
        .retain(|face| face["family"].as_str() == Some("NotoSansSC"));
    let output = PathBuf::from(env::var_os("OUT_DIR").expect("OUT_DIR")).join("manifest.json");
    fs::write(
        &output,
        serde_json::to_vec(&manifest).expect("serialize manifest"),
    )
    .expect("write manifest");
    println!("cargo:rustc-env=FONT_PACK_MANIFEST={}", output.display());
    println!("cargo:rustc-env=FONT_PACK_ASSETS={}", assets.display());
}
