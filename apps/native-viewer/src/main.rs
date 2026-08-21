#[cfg(test)]
mod build_support;
mod chrome;
mod collaboration;
mod collaboration_document;
mod collaboration_protocol;
mod collaboration_test_peer;
mod document;
#[path = "scene.rs"]
mod docx_scene;
mod editing;
mod fonts;
mod gpu;
mod images;
mod pptx_editing;
mod pptx_scene;
mod scene_shared;
#[cfg(test)]
mod test_fixtures;
mod window;
mod xlsx_editing;
mod xlsx_scene;

use std::env;
use std::ffi::OsString;
use std::path::{Path, PathBuf};

use anyhow::{Context, Result, bail};

use collaboration::{CollaborationConfig, DEFAULT_RELAY_ORIGIN};
use document::{
    DocumentFormat, load_collaborative_document, load_document, load_document_for_export,
};

fn main() -> Result<()> {
    let options = Options::parse()?;
    let format = DocumentFormat::from_path(&options.document)?;
    let (page, sheet) = options.selection(format)?;
    let collaboration = options.collaboration(format)?;
    if options.collaboration_test_peer {
        let config = collaboration.context("--collaboration-test-peer requires --room")?;
        return collaboration_test_peer::run(&options.document, config);
    }
    let (mut context, max_texture_dimension_2d) = gpu::create_render_context()?;
    let document = if options.png.is_some() {
        load_document_for_export(&options.document, sheet, max_texture_dimension_2d)?
    } else if let Some(config) = &collaboration {
        load_collaborative_document(
            &options.document,
            sheet,
            max_texture_dimension_2d,
            config.client_id(),
        )?
    } else {
        load_document(&options.document, sheet, max_texture_dimension_2d)?
    };
    if let Some(output) = &options.png {
        gpu::render_comparison(&mut context, &document, page, output, options.scale)?;
    } else {
        window::run(document, context, collaboration)?;
    }
    Ok(())
}

struct Options {
    document: PathBuf,
    png: Option<PathBuf>,
    page: Option<usize>,
    sheet: Option<usize>,
    slide: Option<usize>,
    scale: f64,
    room: Option<String>,
    relay_origin: String,
    collaboration_test_peer: bool,
}

impl Options {
    fn parse() -> Result<Self> {
        Self::parse_from(env::args_os().skip(1))
    }

    fn parse_from(args: impl IntoIterator<Item = OsString>) -> Result<Self> {
        let mut document = default_document();
        let mut png = None;
        let mut page = None;
        let mut sheet = None;
        let mut slide = None;
        let mut scale = 1.0f64;
        let mut room = None;
        let mut collaboration_test_peer = false;
        let mut relay_origin = match env::var_os("BETTEROFFICE_RELAY_ORIGIN") {
            Some(value) => value
                .into_string()
                .map_err(|_| anyhow::anyhow!("BETTEROFFICE_RELAY_ORIGIN must be Unicode"))?,
            None => DEFAULT_RELAY_ORIGIN.to_owned(),
        };
        let mut args = args.into_iter();
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
                    page = Some(one_based_index(args.next(), "--page")?);
                }
                Some("--sheet") => {
                    sheet = Some(one_based_index(args.next(), "--sheet")?);
                }
                Some("--slide") => {
                    slide = Some(one_based_index(args.next(), "--slide")?);
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
                Some("--room") => {
                    room = Some(
                        args.next()
                            .and_then(|value| value.into_string().ok())
                            .context("--room needs an id")?,
                    );
                }
                Some("--relay-origin") => {
                    relay_origin = args
                        .next()
                        .and_then(|value| value.into_string().ok())
                        .context("--relay-origin needs a URL")?;
                }
                Some("--collaboration-test-peer") => {
                    collaboration_test_peer = true;
                }
                Some("--help" | "-h") => {
                    println!(
                        "Usage: betteroffice-native-viewer [--document FILE] [--png OUT] [--page N | --sheet N | --slide N] [--scale N] [--room ID] [--relay-origin URL]"
                    );
                    std::process::exit(0);
                }
                _ => bail!("unknown argument {}", arg.to_string_lossy()),
            }
        }
        Ok(Self {
            document,
            png,
            page,
            sheet,
            slide,
            scale,
            room,
            relay_origin,
            collaboration_test_peer,
        })
    }

    fn collaboration(&self, _format: DocumentFormat) -> Result<Option<CollaborationConfig>> {
        let Some(room) = &self.room else {
            return Ok(None);
        };
        if self.png.is_some() {
            bail!("--room requires the interactive viewer");
        }
        CollaborationConfig::new(room.clone(), &self.relay_origin).map(Some)
    }

    fn selection(&self, format: DocumentFormat) -> Result<(usize, usize)> {
        match format {
            DocumentFormat::Docx => {
                if self.sheet.is_some() || self.slide.is_some() {
                    bail!("--sheet and --slide do not select DOCX pages");
                }
                Ok((self.page.unwrap_or(0), 0))
            }
            DocumentFormat::Xlsx => {
                if self.page.is_some() || self.slide.is_some() {
                    bail!("--page and --slide do not select XLSX sheets");
                }
                Ok((0, self.sheet.unwrap_or(0)))
            }
            DocumentFormat::Pptx => {
                if self.page.is_some() || self.sheet.is_some() {
                    bail!("--page and --sheet do not select PPTX slides");
                }
                Ok((self.slide.unwrap_or(0), 0))
            }
        }
    }
}

fn one_based_index(value: Option<OsString>, flag: &str) -> Result<usize> {
    let value = value
        .and_then(|value| value.into_string().ok())
        .with_context(|| format!("{flag} needs a number"))?;
    let value = value
        .parse::<usize>()
        .with_context(|| format!("invalid {flag} value"))?;
    if value == 0 {
        bail!("{flag} is one-based");
    }
    Ok(value - 1)
}

fn default_document() -> PathBuf {
    Path::new(env!("CARGO_MANIFEST_DIR")).join("../demo/public/betteroffice-demo.docx")
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn selectors_are_one_based_and_format_specific() {
        let xlsx = Options::parse_from([
            "--document".into(),
            "book.xlsx".into(),
            "--sheet".into(),
            "2".into(),
        ])
        .unwrap();
        assert_eq!(xlsx.selection(DocumentFormat::Xlsx).unwrap(), (0, 1));
        assert!(xlsx.selection(DocumentFormat::Docx).is_err());

        let docx = Options::parse_from([
            "--document".into(),
            "document.docx".into(),
            "--page".into(),
            "3".into(),
        ])
        .unwrap();
        assert_eq!(docx.selection(DocumentFormat::Docx).unwrap(), (2, 0));
        assert!(docx.selection(DocumentFormat::Xlsx).is_err());

        let pptx = Options::parse_from([
            "--document".into(),
            "slides.pptx".into(),
            "--slide".into(),
            "2".into(),
        ])
        .unwrap();
        assert_eq!(pptx.selection(DocumentFormat::Pptx).unwrap(), (1, 0));
        assert!(pptx.selection(DocumentFormat::Docx).is_err());
        assert!(pptx.selection(DocumentFormat::Xlsx).is_err());
    }

    #[test]
    fn rejects_zero_sheet() {
        assert!(Options::parse_from(["--sheet".into(), "0".into()]).is_err());
        assert!(Options::parse_from(["--slide".into(), "0".into()]).is_err());
    }

    #[test]
    fn room_options_enable_interactive_collaboration_for_every_format() {
        let options = Options::parse_from([
            "--document".into(),
            "document.docx".into(),
            "--room".into(),
            "shared".into(),
            "--relay-origin".into(),
            "http://127.0.0.1:8787".into(),
        ])
        .unwrap();
        let collaboration = options
            .collaboration(DocumentFormat::Docx)
            .unwrap()
            .unwrap();
        assert_eq!(collaboration.room(), "shared");
        assert!(options.collaboration(DocumentFormat::Xlsx).is_ok());
        assert!(options.collaboration(DocumentFormat::Pptx).is_ok());

        let export = Options::parse_from([
            "--document".into(),
            "document.docx".into(),
            "--png".into(),
            "out.png".into(),
            "--room".into(),
            "shared".into(),
        ])
        .unwrap();
        assert!(export.collaboration(DocumentFormat::Docx).is_err());
    }
}
