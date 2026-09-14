use std::time::Instant;

use docx_edit::bridge::RenderEnv;
use docx_edit::{EngineSession, seed_from_docx};
use serde_json::json;
use sha2::{Digest, Sha256};
use yrs::{Map, ReadTxn, Transact};

fn fingerprint(bytes: &[u8]) -> String {
    format!("{:x}", Sha256::digest(bytes))
}

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut args = std::env::args().skip(1);
    let path = args
        .next()
        .ok_or("usage: open_probe input.docx [samples]")?;
    let samples: usize = args.next().as_deref().unwrap_or("3").parse()?;
    if !(1..=10).contains(&samples) || args.next().is_some() {
        return Err("usage: open_probe input.docx [samples: 1..10]".into());
    }
    let bytes = std::fs::read(&path)?;
    let env = RenderEnv::default();
    for sample in 0..=samples {
        let engine = EngineSession::new(7);
        let start = Instant::now();
        seed_from_docx(engine.doc(), &bytes)?;
        let open_ms = start.elapsed().as_secs_f64() * 1000.0;
        let txn = engine.doc().yrs_doc().transact();
        let mut stories: Vec<_> = txn
            .get_map("stories")
            .ok_or("seeded document has no stories")?
            .keys(&txn)
            .map(str::to_owned)
            .collect();
        stories.sort();
        drop(txn);
        let start = Instant::now();
        let body_blocks = engine.with_lowered_story("body", &env, <[_]>::len)?;
        let lower_body_ms = start.elapsed().as_secs_f64() * 1000.0;
        let mut lowered = Vec::new();
        for story in &stories {
            lowered.push((story, engine.lower_story_json(story, &env)?));
        }
        println!(
            "{}",
            json!({
                "source": path,
                "sample": sample,
                "warmup": sample == 0,
                "openMs": open_ms,
                "lowerBodyMs": lower_body_ms,
                "bodyBlocks": body_blocks,
                "stories": stories.len(),
                "stateFingerprint": fingerprint(&engine.doc().encode_state_as_update_v1()),
                "blocksFingerprint": fingerprint(&serde_json::to_vec(&lowered)?),
            })
        );
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    #[test]
    fn fingerprints_use_sha256() {
        assert_eq!(
            super::fingerprint(b"abc"),
            "ba7816bf8f01cfea414140de5dae2223b00361a396177a9cb410ff61f20015ad"
        );
    }
}
