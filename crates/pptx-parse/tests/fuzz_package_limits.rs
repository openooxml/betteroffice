mod fuzz_harness;

use std::path::Path;

use pptx_parse::PptxError;

const SLIDE_REFERENCED_TWICE: &[u8] =
    include_bytes!("../../../fuzz/corpus/pptx-package-parse/slide-referenced-twice.records");

#[test]
fn every_committed_seed_passes_the_fuzz_oracle() {
    let corpus = Path::new(env!("CARGO_MANIFEST_DIR")).join("../../fuzz/corpus/pptx-package-parse");
    let mut seeds = 0;
    for entry in std::fs::read_dir(&corpus).unwrap() {
        let path = entry.unwrap().path();
        let bytes = std::fs::read(&path).unwrap();
        assert!(
            fuzz_harness::run(&bytes).is_some(),
            "{} does not decode to a package",
            path.display()
        );
        seeds += 1;
    }
    assert!(seeds > 0, "no seeds in {}", corpus.display());
}

#[test]
fn a_slide_referenced_twice_spends_one_package_shape_budget() {
    assert!(matches!(
        fuzz_harness::run(SLIDE_REFERENCED_TWICE),
        Some(Err(PptxError::ResourceLimit { kind: "shapes", .. }))
    ));
}
