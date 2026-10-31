use std::env;
use std::fs;
use std::path::PathBuf;

fn main() {
    let target = env::var("TARGET").expect("TARGET");
    // wasm32 decodes with ruzstd, so a wasm build needs no C compiler with a wasm target.
    // No link arguments here: Cargo hands a build script's cdylib link arguments to every
    // dependent cdylib, so a memory cap set here would also cap each engine that links this
    // crate. The standalone decoder's cap is set by scripts/build-fonts-decoder.ts.
    if target.starts_with("wasm32") {
        return;
    }
    println!("cargo:rerun-if-changed=vendor/zstd/zstddeclib.c");

    let manifest = PathBuf::from(env::var_os("CARGO_MANIFEST_DIR").expect("CARGO_MANIFEST_DIR"));
    let source = fs::read_to_string(manifest.join("vendor/zstd/zstddeclib.c"))
        .expect("vendor zstd first with bash scripts/fonts/vendor-zstd.sh");
    let output = PathBuf::from(env::var_os("OUT_DIR").expect("OUT_DIR"));
    let source_path = output.join("fontpack_zstd.c");
    fs::write(&source_path, namespace(&source)).expect("write namespaced zstd decoder");

    cc::Build::new()
        .file(source_path)
        .define("ZSTD_LEGACY_SUPPORT", "0")
        .define("ZSTD_MULTITHREAD", "0")
        .warnings(false)
        .compile("fontpack_zstd");
}

// Isolate decoder symbols from the full zstd compressor linked by the build tool.
fn namespace(source: &str) -> String {
    let mut output = String::with_capacity(source.len());
    let mut start = 0;
    for (index, ch) in source.char_indices() {
        if ch.is_ascii_alphanumeric() || ch == '_' {
            continue;
        }
        push_token(&mut output, &source[start..index]);
        output.push(ch);
        start = index + ch.len_utf8();
    }
    push_token(&mut output, &source[start..]);
    output
}

fn push_token(output: &mut String, token: &str) {
    if ["ZSTD_", "FSE_", "HUF_", "HIST_", "ERR_"]
        .iter()
        .any(|prefix| token.starts_with(prefix))
        && token.bytes().any(|byte| byte.is_ascii_lowercase())
    {
        output.push_str("FONTPACK_");
    }
    output.push_str(token);
}
