use std::collections::BTreeMap;
use std::fs;
use std::path::PathBuf;

use betteroffice_fonts::slices::{Skeleton, chunk_entries};
use sha2::{Digest, Sha256};

fn assets() -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("../../packages/fonts-cjk/assets")
}

fn decode(file: &str) -> Vec<u8> {
    let encoded = fs::read(assets().join(file)).unwrap();
    zstd::stream::decode_all(encoded.as_slice()).unwrap()
}

fn tables(font: &[u8]) -> BTreeMap<[u8; 4], &[u8]> {
    let count = u16::from_be_bytes([font[4], font[5]]) as usize;
    (0..count)
        .map(|index| {
            let record = &font[12 + index * 16..28 + index * 16];
            let offset = u32::from_be_bytes(record[8..12].try_into().unwrap()) as usize;
            let length = u32::from_be_bytes(record[12..16].try_into().unwrap()) as usize;
            (
                record[..4].try_into().unwrap(),
                &font[offset..offset + length],
            )
        })
        .collect()
}

#[test]
fn shipped_cjk_slices_reassemble_to_their_recorded_faces() {
    let manifest: serde_json::Value =
        serde_json::from_str(&fs::read_to_string(assets().join("manifest.json")).unwrap()).unwrap();
    let faces = manifest["faces"].as_array().unwrap();
    assert_eq!(faces.len(), 5);
    for face in faces {
        let pointer = &face["slices"];
        let index = fs::read(assets().join(pointer["index"].as_str().unwrap())).unwrap();
        assert_eq!(
            format!("{:x}", Sha256::digest(&index)),
            pointer["indexSha256"].as_str().unwrap()
        );
        let slices: serde_json::Value = serde_json::from_slice(&index).unwrap();
        assert_eq!(slices["sha256"], pointer["sha256"]);
        let skeleton_bytes = decode(slices["skeleton"]["file"].as_str().unwrap());
        let skeleton = Skeleton::parse(&skeleton_bytes).unwrap();
        let mut files = vec![decode(slices["core"]["file"].as_str().unwrap())];
        for (index, chunk) in slices["chunks"].as_array().unwrap().iter().enumerate() {
            assert_eq!(chunk["id"].as_u64().unwrap(), index as u64 + 1);
            let bytes = decode(chunk["file"].as_str().unwrap());
            assert_eq!(bytes.len() as u64, chunk["byteLength"].as_u64().unwrap());
            for (glyph, _) in chunk_entries(&bytes).unwrap() {
                assert!(
                    skeleton
                        .chunk_of_glyph(glyph)
                        .is_some_and(|owner| owner as usize <= index + 1)
                );
            }
            files.push(bytes);
        }
        let chunks: Vec<&[u8]> = files.iter().map(Vec::as_slice).collect();
        let full = skeleton.reassemble(&chunks).unwrap();
        assert_eq!(full.len() as u64, slices["byteLength"].as_u64().unwrap());
        assert_eq!(
            format!("{:x}", Sha256::digest(&full)),
            slices["sha256"].as_str().unwrap()
        );

        let whole = decode(face["file"].as_str().unwrap());
        let (original, rebuilt) = (tables(&whole), tables(&full));
        assert_eq!(
            original.keys().collect::<Vec<_>>(),
            rebuilt.keys().collect::<Vec<_>>()
        );
        for (tag, table) in &original {
            match tag {
                b"CFF " => {}
                b"head" => {
                    assert_eq!(&table[..8], &rebuilt[tag][..8]);
                    assert_eq!(&table[12..], &rebuilt[tag][12..]);
                }
                _ => assert_eq!(table, &rebuilt[tag], "{}", String::from_utf8_lossy(tag)),
            }
        }

        let skeleton_only = skeleton.reassemble(&[]).unwrap();
        assert_eq!(tables(&skeleton_only)[b"hmtx"], original[b"hmtx"]);
        assert_eq!(skeleton.chunk_of_codepoint(0x20), Some(0));
    }
}
