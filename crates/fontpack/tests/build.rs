#![cfg(feature = "build")]

mod common;

use std::fs;
use std::process::Command;

use serde_json::{Value, json};
use sha2::{Digest, Sha256};

#[test]
fn builder_emits_a_sorted_manifest_and_verified_faces() {
    let temp = tempfile::tempdir().unwrap();
    let src = temp.path().join("src");
    let out = temp.path().join("out");
    fs::create_dir(&src).unwrap();
    let regular = common::synthetic(42);
    let mut bold = regular.clone();
    bold[100] ^= 1;
    let standalone = common::synthetic(43);
    for (name, bytes) in [
        ("B-Regular.otf", &standalone),
        ("A-Regular.ttf", &regular),
        ("A-Bold.ttf", &bold),
    ] {
        fs::write(src.join(name), bytes).unwrap();
    }
    fs::write(src.join("ignored.txt"), b"ignored").unwrap();
    let result = Command::new(env!("CARGO_BIN_EXE_fontpack"))
        .args(["build", "--src"])
        .arg(&src)
        .arg("--out")
        .arg(&out)
        .args(["--pack", "synthetic", "--version", "1.2.3"])
        .output()
        .unwrap();
    assert!(
        result.status.success(),
        "{}",
        String::from_utf8_lossy(&result.stderr)
    );
    let manifest: Value =
        serde_json::from_slice(&fs::read(out.join("manifest.json")).unwrap()).unwrap();
    let mut expected = Vec::new();
    for (name, family, bytes, dictionary) in [
        ("A-Bold.ttf", "A", &bold, Some("A-Regular.ttf.zst")),
        ("A-Regular.ttf", "A", &regular, None),
        ("B-Regular.otf", "B", &standalone, None),
    ] {
        let file = format!("{name}.zst");
        let encoded = fs::read(out.join(&file)).unwrap();
        let dict = dictionary.map(|_| regular.as_slice());
        assert_eq!(
            fontpack::decode(&encoded, dict, bytes.len()).unwrap(),
            *bytes
        );
        if dictionary.is_some() {
            assert!(encoded.len() < common::compress(bytes, None).len());
            assert!(fontpack::decode(&encoded, None, bytes.len()).is_err());
        }
        expected.push(json!({
            "file": file,
            "source": name,
            "family": family,
            "codec": "zstd",
            "encodedLength": encoded.len(),
            "byteLength": bytes.len(),
            "sha256": format!("{:x}", Sha256::digest(bytes)),
            "dict": dictionary,
        }));
    }
    assert_eq!(
        manifest,
        json!({ "format": 1, "pack": "synthetic", "version": "1.2.3", "faces": expected })
    );
}

fn woff2t_build(src: &std::path::Path, out: &std::path::Path) -> std::process::Output {
    Command::new(env!("CARGO_BIN_EXE_fontpack"))
        .args(["build", "--src"])
        .arg(src)
        .arg("--out")
        .arg(out)
        .args([
            "--pack",
            "synthetic",
            "--version",
            "1.2.3",
            "--codec",
            "woff2t",
        ])
        .output()
        .unwrap()
}

fn copy_fixtures(src: &std::path::Path) {
    fs::create_dir(src).unwrap();
    let fixtures = std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("tests/fixtures");
    for style in ["Regular", "Bold"] {
        for extension in ["ttf", "uw2", "woff2"] {
            let name = format!("Synthetic-{style}.{extension}");
            fs::copy(fixtures.join(&name), src.join(name)).unwrap();
        }
    }
}

#[test]
fn builder_verifies_woff2t_and_hashes_reconstructed_fonts() {
    let temp = tempfile::tempdir().unwrap();
    let src = temp.path().join("src");
    let out = temp.path().join("out");
    copy_fixtures(&src);
    let result = woff2t_build(&src, &out);
    assert!(
        result.status.success(),
        "{}",
        String::from_utf8_lossy(&result.stderr)
    );
    let manifest: Value =
        serde_json::from_slice(&fs::read(out.join("manifest.json")).unwrap()).unwrap();
    assert_eq!(manifest["format"], 1);
    let regular_frame = fs::read(out.join("Synthetic-Regular.ttf.zst")).unwrap();
    let regular =
        fontpack::decode_face(&regular_frame, None, fontpack::Codec::Woff2tZstd, 65536).unwrap();
    assert_eq!(
        regular.payload.as_deref().unwrap(),
        fs::read(src.join("Synthetic-Regular.uw2")).unwrap()
    );
    for face in manifest["faces"].as_array().unwrap() {
        let file = face["file"].as_str().unwrap();
        let source = face["source"].as_str().unwrap();
        let dict = if source.contains("Bold") {
            regular.payload.as_deref()
        } else {
            None
        };
        let decoded = fontpack::decode_face(
            &fs::read(out.join(file)).unwrap(),
            dict,
            fontpack::Codec::Woff2tZstd,
            65536,
        )
        .unwrap();
        let standard = fs::read(src.join(source).with_extension("woff2")).unwrap();
        let expected = wuff::decompress_woff2(&standard).unwrap();
        assert_eq!(decoded.font, expected);
        assert_eq!(face["codec"], "woff2t+zstd");
        assert_eq!(face["byteLength"], expected.len());
        assert_eq!(face["sha256"], format!("{:x}", Sha256::digest(&expected)));
        assert_eq!(
            face["dict"],
            if dict.is_some() {
                json!("Synthetic-Regular.ttf.zst")
            } else {
                Value::Null
            }
        );
    }
}

#[test]
fn builder_rejects_non_ttf_and_missing_companions_for_woff2t() {
    let temp = tempfile::tempdir().unwrap();
    let src = temp.path().join("src");
    let out = temp.path().join("out");
    fs::create_dir(&src).unwrap();
    fs::write(src.join("Cff-Regular.otf"), b"OTTO").unwrap();
    let result = woff2t_build(&src, &out);
    assert!(!result.status.success());
    assert!(String::from_utf8_lossy(&result.stderr).contains("TTF glyf/loca"));
    fs::remove_file(src.join("Cff-Regular.otf")).unwrap();
    fs::write(src.join("Cff-Regular.ttf"), b"OTTO").unwrap();
    let result = woff2t_build(&src, &out);
    assert!(!result.status.success());
    assert!(String::from_utf8_lossy(&result.stderr).contains("TTF glyf/loca"));
    fs::remove_file(src.join("Cff-Regular.ttf")).unwrap();
    fs::write(
        src.join("Synthetic-Regular.ttf"),
        include_bytes!("fixtures/Synthetic-Regular.ttf"),
    )
    .unwrap();
    assert!(!woff2t_build(&src, &out).status.success());
}

#[test]
fn builder_rejects_reference_and_glyph_count_mismatches() {
    let temp = tempfile::tempdir().unwrap();
    let src = temp.path().join("src");
    let out = temp.path().join("out");
    copy_fixtures(&src);
    let bold_reference = fs::read(src.join("Synthetic-Bold.woff2")).unwrap();
    fs::copy(
        src.join("Synthetic-Regular.woff2"),
        src.join("Synthetic-Bold.woff2"),
    )
    .unwrap();
    let result = woff2t_build(&src, &out);
    assert!(!result.status.success());
    assert!(String::from_utf8_lossy(&result.stderr).contains("round-trip mismatch"));
    fs::write(src.join("Synthetic-Bold.woff2"), bold_reference).unwrap();
    let path = src.join("Synthetic-Bold.ttf");
    let mut bytes = fs::read(&path).unwrap();
    let count = u16::from_be_bytes([bytes[4], bytes[5]]) as usize;
    for index in 0..count {
        let start = 12 + index * 16;
        if &bytes[start..start + 4] == b"maxp" {
            let offset =
                u32::from_be_bytes(bytes[start + 8..start + 12].try_into().unwrap()) as usize;
            bytes[offset + 4..offset + 6].copy_from_slice(&6u16.to_be_bytes());
        }
    }
    fs::write(path, bytes).unwrap();
    let result = woff2t_build(&src, &out);
    assert!(!result.status.success());
    assert!(String::from_utf8_lossy(&result.stderr).contains("glyph count mismatch"));
}
