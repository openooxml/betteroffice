use std::time::Instant;
fn main() {
    let path = std::env::args().nth(1).expect("input.json path");
    let iters: u32 = std::env::args().nth(2).and_then(|s| s.parse().ok()).unwrap_or(50);
    let json = std::fs::read_to_string(&path).expect("read input");
    // warmup + verify
    let out = docx_layout::layout_to_json(&json).expect("layout");
    eprintln!("out bytes: {}", out.len());
    let t = Instant::now();
    for _ in 0..iters {
        let out = docx_layout::layout_to_json(&json).expect("layout");
        std::hint::black_box(out);
    }
    let d = t.elapsed();
    println!("{iters} iters: {:?} total, {:.3} ms/iter", d, d.as_secs_f64() * 1000.0 / iters as f64);
}
