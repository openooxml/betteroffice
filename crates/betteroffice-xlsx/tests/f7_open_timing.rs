use std::time::Instant;

use betteroffice_xlsx::Workbook;

#[test]
fn f7_open_timing() {
    let Ok(files) = std::env::var("F7_TIME") else {
        return;
    };
    for path in files.split(',').filter(|path| !path.is_empty()) {
        let bytes = std::fs::read(path).unwrap();
        let mut micros: Vec<u128> = (0..5)
            .map(|_| {
                let started = Instant::now();
                let opened = Workbook::open(&bytes);
                let elapsed = started.elapsed().as_micros();
                drop(opened);
                elapsed
            })
            .collect();
        micros.sort_unstable();
        println!("F7TIME {path} median_micros={} runs={micros:?}", micros[2]);
    }
}
