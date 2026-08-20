mod document;
mod fonts;
mod gpu;
mod images;
mod scene;
mod window;

use std::env;
use std::path::{Path, PathBuf};

use anyhow::{Context, Result, bail};

use document::load_document;

fn main() -> Result<()> {
    let options = Options::parse()?;
    let document = load_document(&options.document)?;
    if let Some(output) = options.png {
        gpu::render_comparison(&document, options.page, &output, options.scale)?;
    } else {
        window::run(document)?;
    }
    Ok(())
}

struct Options {
    document: PathBuf,
    png: Option<PathBuf>,
    page: usize,
    scale: f64,
}

impl Options {
    fn parse() -> Result<Self> {
        let mut document = default_document();
        let mut png = None;
        let mut page = 1usize;
        let mut scale = 1.0f64;
        let mut args = env::args_os().skip(1);
        while let Some(arg) = args.next() {
            match arg.to_str() {
                Some("--document") => {
                    document = args
                        .next()
                        .map(PathBuf::from)
                        .context("--document needs a path")?;
                }
                Some("--png") => {
                    png = Some(
                        args.next()
                            .map(PathBuf::from)
                            .context("--png needs a path")?,
                    );
                }
                Some("--page") => {
                    page = args
                        .next()
                        .and_then(|value| value.into_string().ok())
                        .context("--page needs a number")?
                        .parse()
                        .context("invalid --page value")?;
                    if page == 0 {
                        bail!("--page is one-based");
                    }
                }
                Some("--scale") => {
                    scale = args
                        .next()
                        .and_then(|value| value.into_string().ok())
                        .context("--scale needs a number")?
                        .parse()
                        .context("invalid --scale value")?;
                    if !scale.is_finite() || scale <= 0.0 || scale > 8.0 {
                        bail!("--scale must be greater than zero and at most 8");
                    }
                }
                Some("--help" | "-h") => {
                    println!(
                        "Usage: betteroffice-native-viewer [--document FILE] [--png OUT] [--page N] [--scale N]"
                    );
                    std::process::exit(0);
                }
                _ => bail!("unknown argument {}", arg.to_string_lossy()),
            }
        }
        Ok(Self {
            document,
            png,
            page: page - 1,
            scale,
        })
    }
}

fn default_document() -> PathBuf {
    Path::new(env!("CARGO_MANIFEST_DIR")).join("../demo/public/betteroffice-demo.docx")
}
