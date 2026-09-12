#![no_main]

#[path = "../../crates/pptx-parse/tests/fuzz_harness/mod.rs"]
mod fuzz_harness;

use libfuzzer_sys::fuzz_target;

fuzz_target!(|data: &[u8]| {
    let _ = fuzz_harness::run(data);
});
