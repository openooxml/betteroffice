use std::collections::BTreeMap;
use std::env;
use std::error::Error;
use std::ffi::OsString;
use std::fs;
use std::io;
use std::path::PathBuf;
use std::process::ExitCode;

use serde::Serialize;
use sha2::{Digest, Sha256};

type Result<T> = std::result::Result<T, Box<dyn Error>>;

struct Options {
    src: PathBuf,
    out: PathBuf,
    pack: String,
    version: String,
    level: i32,
}

#[derive(Serialize)]
struct Manifest {
    format: u32,
    pack: String,
    version: String,
    faces: Vec<Face>,
}

#[derive(Serialize)]
#[serde(rename_all = "camelCase")]
struct Face {
    file: String,
    source: String,
    family: String,
    codec: &'static str,
    encoded_length: usize,
    byte_length: usize,
    sha256: String,
    dict: Option<String>,
}

struct Source {
    file: String,
    family: String,
    regular: bool,
    bytes: Vec<u8>,
}

fn invalid(message: impl Into<String>) -> io::Error {
    io::Error::new(io::ErrorKind::InvalidInput, message.into())
}

fn options() -> Result<Options> {
    let mut args = env::args_os().skip(1);
    if args.next().as_deref() != Some(std::ffi::OsStr::new("build")) {
        return Err(invalid(concat!(
            "usage: fontpack build --src <dir> --out <dir> ",
            "--pack <name> --version <semver> [--level 19]"
        ))
        .into());
    }
    let mut values = BTreeMap::<String, OsString>::new();
    while let Some(arg) = args.next() {
        let key = arg.into_string().map_err(|_| invalid("invalid option"))?;
        if !["--src", "--out", "--pack", "--version", "--level"].contains(&key.as_str()) {
            return Err(invalid(format!("unknown option {key}")).into());
        }
        let value = args
            .next()
            .ok_or_else(|| invalid(format!("missing value for {key}")))?;
        if values.insert(key.clone(), value).is_some() {
            return Err(invalid(format!("duplicate option {key}")).into());
        }
    }
    let mut take = |key: &str| {
        values
            .remove(key)
            .ok_or_else(|| invalid(format!("missing {key}")))
    };
    let src = PathBuf::from(take("--src")?);
    let out = PathBuf::from(take("--out")?);
    let pack = take("--pack")?
        .into_string()
        .map_err(|_| invalid("invalid pack name"))?;
    if pack.is_empty() {
        return Err(invalid("pack name must not be empty").into());
    }
    let version = take("--version")?
        .into_string()
        .map_err(|_| invalid("invalid version"))?;
    semver::Version::parse(&version)?;
    let level = values
        .remove("--level")
        .map(|value| value.into_string().map_err(|_| invalid("invalid level")))
        .transpose()?
        .unwrap_or_else(|| "19".to_owned())
        .parse()?;
    Ok(Options {
        src,
        out,
        pack,
        version,
        level,
    })
}

fn sources(options: &Options) -> Result<Vec<Source>> {
    let mut files = Vec::new();
    for entry in fs::read_dir(&options.src)? {
        let entry = entry?;
        if !entry.file_type()?.is_file() {
            continue;
        }
        let path = entry.path();
        let Some(extension) = path.extension().and_then(|value| value.to_str()) else {
            continue;
        };
        if !extension.eq_ignore_ascii_case("ttf") && !extension.eq_ignore_ascii_case("otf") {
            continue;
        }
        let file = entry
            .file_name()
            .into_string()
            .map_err(|_| invalid("non-UTF-8 face filename"))?;
        let stem = path
            .file_stem()
            .and_then(|value| value.to_str())
            .ok_or_else(|| invalid("invalid face filename"))?;
        let (family, style) = stem.rsplit_once('-').unwrap_or((stem, ""));
        if family.is_empty() {
            return Err(invalid(format!("empty family in {file}")).into());
        }
        files.push(Source {
            file,
            family: family.to_owned(),
            regular: style == "Regular",
            bytes: fs::read(path)?,
        });
    }
    files.sort_by(|left, right| left.file.cmp(&right.file));
    if files.is_empty() {
        return Err(invalid("source directory has no TTF or OTF faces").into());
    }
    Ok(files)
}

fn compress(bytes: &[u8], dict: Option<&[u8]>, level: i32) -> Result<Vec<u8>> {
    let span = bytes
        .len()
        .checked_add(dict.map_or(0, <[u8]>::len))
        .ok_or_else(|| invalid("face and dictionary are too large"))?;
    let window = span
        .max(1024)
        .checked_next_power_of_two()
        .ok_or_else(|| invalid("face and dictionary are too large"))?;
    let window_log = window.ilog2();
    let mut compressor = zstd::bulk::Compressor::default();
    compressor.set_parameter(zstd::zstd_safe::CParameter::CompressionLevel(level))?;
    compressor.include_checksum(true)?;
    compressor.include_contentsize(true)?;
    compressor.window_log(window_log)?;
    if let Some(dict) = dict {
        compressor
            .context_mut()
            .ref_prefix(dict)
            .map_err(|code| io::Error::other(zstd::zstd_safe::get_error_name(code)))?;
    }
    Ok(compressor.compress(bytes)?)
}

fn build(options: Options) -> Result<()> {
    let sources = sources(&options)?;
    let mut families = BTreeMap::<&str, Vec<usize>>::new();
    for (index, face) in sources.iter().enumerate() {
        families.entry(&face.family).or_default().push(index);
    }
    let mut dictionaries = vec![None; sources.len()];
    for (family, indices) in families {
        if indices.len() < 2 {
            continue;
        }
        let regulars: Vec<_> = indices
            .iter()
            .copied()
            .filter(|&index| sources[index].regular)
            .collect();
        let [regular] = regulars.as_slice() else {
            return Err(invalid(format!("family {family} needs exactly one Regular face")).into());
        };
        for index in indices {
            if index != *regular {
                dictionaries[index] = Some(*regular);
            }
        }
    }
    fs::create_dir_all(&options.out)?;
    let mut faces = Vec::with_capacity(sources.len());
    for (index, source) in sources.iter().enumerate() {
        let dictionary = dictionaries[index].map(|index| &sources[index]);
        let dict = dictionary.map(|source| source.bytes.as_slice());
        let encoded = compress(&source.bytes, dict, options.level)?;
        let decoded = fontpack::decode(&encoded, dict, source.bytes.len())?;
        if decoded != source.bytes {
            return Err(
                io::Error::other(format!("round-trip mismatch for {}", source.file)).into(),
            );
        }
        let file = format!("{}.zst", source.file);
        fs::write(options.out.join(&file), &encoded)?;
        faces.push(Face {
            file,
            source: source.file.clone(),
            family: source.family.clone(),
            codec: "zstd",
            encoded_length: encoded.len(),
            byte_length: source.bytes.len(),
            sha256: format!("{:x}", Sha256::digest(&decoded)),
            dict: dictionary.map(|source| format!("{}.zst", source.file)),
        });
    }
    faces.sort_by(|left, right| left.file.cmp(&right.file));
    let manifest = Manifest {
        format: 1,
        pack: options.pack,
        version: options.version,
        faces,
    };
    let mut json = serde_json::to_vec_pretty(&manifest)?;
    json.push(b'\n');
    fs::write(options.out.join("manifest.json"), json)?;
    Ok(())
}

fn main() -> ExitCode {
    match options().and_then(build) {
        Ok(()) => ExitCode::SUCCESS,
        Err(error) => {
            eprintln!("fontpack: {error}");
            ExitCode::FAILURE
        }
    }
}
