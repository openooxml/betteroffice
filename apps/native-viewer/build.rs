use std::env;
use std::fs;
use std::path::PathBuf;

fn numeric_product(source: &str, name: &str) -> u64 {
    let declaration = source
        .lines()
        .find(|line| line.contains(name) && line.contains('='))
        .unwrap_or_else(|| panic!("missing {name}"));
    declaration
        .split_once('=')
        .unwrap()
        .1
        .trim()
        .trim_end_matches(';')
        .split('*')
        .map(|part| part.trim().parse::<u64>().unwrap())
        .product()
}

fn main() {
    let limits_path = "../../shared/collaboration-limits.ts";
    let protocol_path = "../../packages/docx/src/collaboration/protocol.ts";
    println!("cargo::rerun-if-changed={limits_path}");
    println!("cargo::rerun-if-changed={protocol_path}");
    let limits = fs::read_to_string(limits_path).unwrap();
    let protocol = fs::read_to_string(protocol_path).unwrap();
    let frame = numeric_product(&limits, "MAX_COLLABORATION_FRAME_BYTES");
    let messages = numeric_product(&protocol, "DEFAULT_MAX_MESSAGES_PER_FRAME");
    let generated = format!(
        "pub const MAX_FRAME_BYTES: usize = {frame};\npub const MAX_MESSAGES_PER_FRAME: usize = {messages};\n"
    );
    let output = PathBuf::from(env::var_os("OUT_DIR").unwrap()).join("collaboration_limits.rs");
    fs::write(output, generated).unwrap();
}
