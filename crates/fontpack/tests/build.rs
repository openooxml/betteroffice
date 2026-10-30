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
        json!({ "format": 1, "schema": 1, "pack": "synthetic", "version": "1.2.3", "faces": expected })
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
        for extension in ["ttf", "uw2", "woff2", "source.sha256"] {
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

fn refresh_source_hash(src: &std::path::Path, stem: &str) {
    let lines: String = ["ttf", "uw2", "woff2"]
        .iter()
        .map(|extension| {
            let name = format!("{stem}.{extension}");
            let digest = Sha256::digest(fs::read(src.join(&name)).unwrap());
            format!("{digest:x}  {name}\n")
        })
        .collect();
    fs::write(src.join(format!("{stem}.source.sha256")), lines).unwrap();
}

#[test]
fn builder_rejects_companions_swapped_from_another_face() {
    let temp = tempfile::tempdir().unwrap();
    let src = temp.path().join("src");
    let out = temp.path().join("out");
    copy_fixtures(&src);
    for extension in ["uw2", "woff2"] {
        fs::copy(
            src.join(format!("Synthetic-Regular.{extension}")),
            src.join(format!("Synthetic-Bold.{extension}")),
        )
        .unwrap();
    }
    let result = woff2t_build(&src, &out);
    assert!(!result.status.success());
    let stderr = String::from_utf8_lossy(&result.stderr);
    assert!(stderr.contains("source SHA-256 mismatch for Synthetic-Bold.ttf"));
    assert!(!out.exists());
}

#[test]
fn builder_rejects_stale_woff2t_companions_with_unchanged_glyph_count() {
    let temp = tempfile::tempdir().unwrap();
    let src = temp.path().join("src");
    let out = temp.path().join("out");
    copy_fixtures(&src);
    let path = src.join("Synthetic-Bold.ttf");
    let mut bytes = fs::read(&path).unwrap();
    bytes.extend_from_slice(&[0; 4]);
    fs::write(path, bytes).unwrap();
    let result = woff2t_build(&src, &out);
    assert!(!result.status.success());
    let stderr = String::from_utf8_lossy(&result.stderr);
    assert!(stderr.contains("source SHA-256 mismatch for Synthetic-Bold.ttf"));
    assert!(stderr.contains("regenerate companions with scripts/fonts/woff2-transform.py"));
    assert!(!out.exists());
}

#[test]
fn builder_rejects_missing_woff2t_source_hash() {
    let temp = tempfile::tempdir().unwrap();
    let src = temp.path().join("src");
    let out = temp.path().join("out");
    copy_fixtures(&src);
    fs::remove_file(src.join("Synthetic-Regular.source.sha256")).unwrap();
    let result = woff2t_build(&src, &out);
    assert!(!result.status.success());
    let stderr = String::from_utf8_lossy(&result.stderr);
    assert!(stderr.contains("cannot read source SHA-256 for Synthetic-Regular.ttf"));
    assert!(stderr.contains("regenerate companions with scripts/fonts/woff2-transform.py"));
    assert!(!out.exists());
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
    refresh_source_hash(&src, "Synthetic-Bold");
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
    fs::write(&path, &bytes).unwrap();
    refresh_source_hash(&src, "Synthetic-Bold");
    let result = woff2t_build(&src, &out);
    assert!(!result.status.success());
    assert!(String::from_utf8_lossy(&result.stderr).contains("glyph count mismatch"));
}

#[test]
fn builder_refuses_an_existing_output_and_leaves_no_staging() {
    let temp = tempfile::tempdir().unwrap();
    let src = temp.path().join("src");
    let out = temp.path().join("out");
    fs::create_dir(&src).unwrap();
    fs::write(src.join("A-Regular.ttf"), common::synthetic(42)).unwrap();
    let run = || {
        Command::new(env!("CARGO_BIN_EXE_fontpack"))
            .args(["build", "--src"])
            .arg(&src)
            .arg("--out")
            .arg(&out)
            .args(["--pack", "synthetic", "--version", "1.2.3"])
            .output()
            .unwrap()
    };
    let first = run();
    assert!(
        first.status.success(),
        "{}",
        String::from_utf8_lossy(&first.stderr)
    );
    let manifest = fs::read(out.join("manifest.json")).unwrap();
    let face = fs::read(out.join("A-Regular.ttf.zst")).unwrap();

    fs::write(src.join("A-Regular.ttf"), common::synthetic(43)).unwrap();
    fs::write(src.join("A-Bold.ttf"), common::synthetic(44)).unwrap();
    let second = run();
    assert!(!second.status.success());
    assert!(String::from_utf8_lossy(&second.stderr).contains("already exists"));
    assert_eq!(fs::read(out.join("manifest.json")).unwrap(), manifest);
    assert_eq!(fs::read(out.join("A-Regular.ttf.zst")).unwrap(), face);
    assert!(!out.join("A-Bold.ttf.zst").exists());

    let mut entries: Vec<_> = fs::read_dir(temp.path())
        .unwrap()
        .map(|entry| entry.unwrap().file_name().into_string().unwrap())
        .collect();
    entries.sort();
    assert_eq!(entries, ["out", "src"]);
}

#[test]
fn builder_accepts_a_dot_suffixed_output_and_skips_stale_staging() {
    let temp = tempfile::tempdir().unwrap();
    let src = temp.path().join("src");
    fs::create_dir(&src).unwrap();
    fs::write(src.join("A-Regular.ttf"), common::synthetic(42)).unwrap();
    let stale = temp.path().join(".out.partial-1");
    fs::create_dir(&stale).unwrap();
    let result = Command::new(env!("CARGO_BIN_EXE_fontpack"))
        .args(["build", "--src"])
        .arg(&src)
        .arg("--out")
        .arg(temp.path().join("out").join("."))
        .args(["--pack", "synthetic", "--version", "1.2.3"])
        .output()
        .unwrap();
    assert!(
        result.status.success(),
        "{}",
        String::from_utf8_lossy(&result.stderr)
    );
    assert!(temp.path().join("out/manifest.json").is_file());
    assert!(stale.is_dir());
    assert_eq!(fs::read_dir(temp.path()).unwrap().count(), 3);
}

#[test]
fn builder_auto_selects_the_smaller_family_encoding() {
    let temp = tempfile::tempdir().unwrap();
    let src = temp.path().join("src");
    copy_fixtures(&src);
    let mut manifests = Vec::new();
    for codec in ["zstd", "woff2t", "auto"] {
        let out = temp.path().join(codec);
        let result = Command::new(env!("CARGO_BIN_EXE_fontpack"))
            .args(["build", "--src"])
            .arg(&src)
            .arg("--out")
            .arg(&out)
            .args([
                "--pack",
                "synthetic",
                "--version",
                "1.2.3",
                "--codec",
                codec,
            ])
            .output()
            .unwrap();
        assert!(
            result.status.success(),
            "{}",
            String::from_utf8_lossy(&result.stderr)
        );
        let manifest: Value =
            serde_json::from_slice(&fs::read(out.join("manifest.json")).unwrap()).unwrap();
        manifests.push(manifest);
    }
    let total = |manifest: &Value| {
        manifest["faces"]
            .as_array()
            .unwrap()
            .iter()
            .map(|face| face["encodedLength"].as_u64().unwrap())
            .sum::<u64>()
    };
    let expected = if total(&manifests[1]) < total(&manifests[0]) {
        &manifests[1]
    } else {
        &manifests[0]
    };
    assert_eq!(manifests[2], *expected);
}
