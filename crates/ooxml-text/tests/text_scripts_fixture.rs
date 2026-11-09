use ooxml_text::paragraph_scripts_with_han_fallback;

/// The script buckets of every code point up to the last Han plane, as ranges of equal buckets, and of seeded
/// random mixes of them. The canvas PNG exports pick fallback faces from the same buckets and read this file.
#[test]
fn text_scripts_match_the_shared_fixture() {
    let path = concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/tests/fixtures/text-scripts.json"
    );
    let mut ranges: Vec<(u32, u32, Vec<String>)> = Vec::new();
    for point in 0..=0x4_0000_u32 {
        let Some(character) = char::from_u32(point) else {
            continue;
        };
        let scripts = paragraph_scripts_with_han_fallback([character.to_string().as_str()], false);
        match ranges.last_mut() {
            Some((_, end, last)) if *last == scripts && *end + 1 == point => *end = point,
            _ => ranges.push((point, point, scripts)),
        }
    }
    let samples: Vec<char> = ranges
        .iter()
        .flat_map(|(start, end, _)| [*start, *end])
        .filter_map(char::from_u32)
        .chain(['a', '1', ' ', '\u{301}', '\u{200d}'])
        .collect();
    let mut seed = 0x9e37_79b9_7f4a_7c15_u64;
    let mut next = move || {
        seed ^= seed << 13;
        seed ^= seed >> 7;
        seed ^= seed << 17;
        (seed >> 11) as f64 / (1u64 << 53) as f64
    };
    let mut mixes = Vec::new();
    for _ in 0..600 {
        let length = 1 + (next() * 8.0) as usize;
        let text: String = (0..length)
            .map(|_| samples[(next() * samples.len() as f64) as usize % samples.len()])
            .collect();
        let scripts = paragraph_scripts_with_han_fallback([text.as_str()], false);
        mixes.push(serde_json::json!({ "text": text, "scripts": scripts }));
    }
    let fixture = serde_json::to_string_pretty(&serde_json::json!({
        "ranges": ranges
            .iter()
            .map(|(start, end, scripts)| serde_json::json!([start, end, scripts]))
            .collect::<Vec<_>>(),
        "mixes": mixes,
    }))
    .unwrap()
        + "\n";
    if std::env::var_os("BO_WRITE_SCRIPT_FIXTURE").is_some() {
        std::fs::write(path, &fixture).unwrap();
    }
    let committed = std::fs::read_to_string(path)
        .expect("run with BO_WRITE_SCRIPT_FIXTURE=1 to write the fixture");
    assert_eq!(
        committed, fixture,
        "script buckets changed: rerun with BO_WRITE_SCRIPT_FIXTURE=1 and check the canvas test"
    );
    let mixed = mixes
        .iter()
        .filter(|mix| {
            mix["scripts"]
                .as_array()
                .is_some_and(|scripts| scripts.len() > 1)
        })
        .count();
    assert!(mixed > 100, "the mixes combine scripts: {mixed}");
}
