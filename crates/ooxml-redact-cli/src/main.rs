use std::path::{Path, PathBuf};
use std::process::ExitCode;

use ooxml_redact_cli::{
    DEFAULT_UPLOAD_URL, MAX_FILE_BYTES, redact_local, report_line, upload_redacted,
};

const USAGE: &str = "\
betteroffice-redact — locally redact an OOXML repro file

usage:
  betteroffice-redact <file.docx|file.xlsx|file.pptx> [options]

options:
  -o, --output <path>    output path (default: <input>.redacted.<ext>)
      --share            upload only the locally redacted bytes
      --endpoint <url>   upload endpoint (default: BETTEROFFICE_REDACT_UPLOAD_URL or BetterOffice)
  -h, --help             show this help

The original file never leaves this machine. Maximum file size: 64 MiB.";

fn main() -> ExitCode {
    match parse_args(std::env::args().skip(1)) {
        Ok(Command::Help) => {
            println!("{USAGE}");
            ExitCode::SUCCESS
        }
        Ok(Command::Run(options)) => run(options),
        Err(error) => {
            eprintln!("error: {error}\n\n{USAGE}");
            ExitCode::from(2)
        }
    }
}

fn run(options: Options) -> ExitCode {
    match execute(options) {
        Ok(()) => ExitCode::SUCCESS,
        Err(error) => {
            eprintln!("error: {error}");
            ExitCode::from(1)
        }
    }
}

fn execute(options: Options) -> Result<(), String> {
    let metadata = std::fs::metadata(&options.input)
        .map_err(|error| format!("reading {}: {error}", options.input.display()))?;
    if metadata.len() > MAX_FILE_BYTES as u64 {
        return Err(format!(
            "input is {} bytes; maximum is {MAX_FILE_BYTES} bytes",
            metadata.len()
        ));
    }
    let input = std::fs::read(&options.input)
        .map_err(|error| format!("reading {}: {error}", options.input.display()))?;
    let redacted = redact_local(&input)?;
    println!("{}", report_line(redacted.report()));

    let output = options
        .output
        .unwrap_or_else(|| default_output(&options.input, redacted.format().extension()));
    if same_path(&options.input, &output) {
        return Err("output path must not overwrite the original file".to_owned());
    }
    if output.exists() {
        return Err(format!("output already exists: {}", output.display()));
    }
    std::fs::write(&output, redacted.bytes())
        .map_err(|error| format!("writing {}: {error}", output.display()))?;
    println!(
        "wrote {} ({} bytes)",
        output.display(),
        redacted.bytes().len()
    );

    if options.share {
        let endpoint = options
            .endpoint
            .or_else(|| std::env::var("BETTEROFFICE_REDACT_UPLOAD_URL").ok())
            .unwrap_or_else(|| DEFAULT_UPLOAD_URL.to_owned());
        let response = upload_redacted(&redacted, &endpoint)?;
        println!("uploaded redacted bytes only: {}", response.url);
    }
    Ok(())
}

#[derive(Debug, Eq, PartialEq)]
struct Options {
    input: PathBuf,
    output: Option<PathBuf>,
    share: bool,
    endpoint: Option<String>,
}

enum Command {
    Help,
    Run(Options),
}

fn parse_args(args: impl IntoIterator<Item = String>) -> Result<Command, String> {
    let args: Vec<String> = args.into_iter().collect();
    let mut input = None;
    let mut output = None;
    let mut share = false;
    let mut endpoint = None;
    let mut index = 0;
    while index < args.len() {
        match args[index].as_str() {
            "-h" | "--help" => return Ok(Command::Help),
            "-o" | "--output" => {
                index += 1;
                output = Some(PathBuf::from(
                    args.get(index).ok_or("--output needs a path")?,
                ));
            }
            "--share" => share = true,
            "--endpoint" => {
                index += 1;
                endpoint = Some(args.get(index).ok_or("--endpoint needs a URL")?.to_owned());
            }
            flag if flag.starts_with('-') => return Err(format!("unknown flag {flag:?}")),
            value => {
                if input.replace(PathBuf::from(value)).is_some() {
                    return Err("expected a single input file".to_owned());
                }
            }
        }
        index += 1;
    }
    let input = input.ok_or("an input DOCX, XLSX, or PPTX file is required")?;
    if endpoint.is_some() && !share {
        return Err("--endpoint requires --share".to_owned());
    }
    Ok(Command::Run(Options {
        input,
        output,
        share,
        endpoint,
    }))
}

fn default_output(input: &Path, extension: Option<&str>) -> PathBuf {
    let stem = input
        .file_stem()
        .and_then(|stem| stem.to_str())
        .unwrap_or("redacted");
    let extension = extension.unwrap_or("ooxml");
    input.with_file_name(format!("{stem}.redacted.{extension}"))
}

fn same_path(left: &Path, right: &Path) -> bool {
    if left == right {
        return true;
    }
    match (left.canonicalize(), right.canonicalize()) {
        (Ok(left), Ok(right)) => left == right,
        _ => false,
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn parses_local_and_share_options() {
        let Command::Run(options) = parse_args(
            [
                "secret.docx",
                "--share",
                "--endpoint",
                "http://127.0.0.1/upload",
                "-o",
                "safe.docx",
            ]
            .map(str::to_owned),
        )
        .unwrap() else {
            panic!("expected run command")
        };
        assert_eq!(options.input, PathBuf::from("secret.docx"));
        assert_eq!(options.output, Some(PathBuf::from("safe.docx")));
        assert!(options.share);
        assert_eq!(options.endpoint.as_deref(), Some("http://127.0.0.1/upload"));
    }

    #[test]
    fn endpoint_requires_share() {
        assert!(
            parse_args(["secret.docx", "--endpoint", "http://localhost"].map(str::to_owned))
                .is_err()
        );
    }

    #[test]
    fn output_name_preserves_detected_extension() {
        assert_eq!(
            default_output(Path::new("report.docx"), Some("docx")),
            PathBuf::from("report.redacted.docx")
        );
    }
}
