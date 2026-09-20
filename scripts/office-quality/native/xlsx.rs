use std::fs;

use betteroffice_xlsx::{CalculationOptions, Workbook};
use serde_json::json;
use sha2::{Digest, Sha256};

fn main() {
    if let Err(error) = run() {
        eprintln!("{error}");
        std::process::exit(1);
    }
}

fn run() -> Result<(), Box<dyn std::error::Error>> {
    let args: Vec<_> = std::env::args_os().skip(1).collect();
    if args.len() != 2 {
        return Err("usage: xlsx-native-bench input.xlsx output.xlsx".into());
    }
    let bytes = fs::read(&args[0])?;
    let workbook = Workbook::open_recalculated(&bytes, CalculationOptions::default())?;
    let output = workbook.save()?;
    fs::write(&args[1], &output)?;
    println!(
        "{}",
        json!({
            "status": "ok",
            "input_sha256": format!("{:x}", Sha256::digest(&bytes)),
            "output_sha256": format!("{:x}", Sha256::digest(&output)),
        })
    );
    Ok(())
}
