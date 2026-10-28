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
