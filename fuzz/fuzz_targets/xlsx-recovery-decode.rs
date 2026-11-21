#![no_main]

use betteroffice_xlsx::fuzzing::recovery_decode;
use libfuzzer_sys::fuzz_target;

fuzz_target!(|data: &[u8]| {
    recovery_decode(data);
});
