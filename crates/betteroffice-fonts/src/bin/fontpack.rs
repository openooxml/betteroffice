use std::collections::BTreeMap;
use std::env;
use std::error::Error;
use std::ffi::OsString;
use std::fs;
use std::io;
use std::path::{Path, PathBuf};
use std::process::ExitCode;
use std::time::{SystemTime, UNIX_EPOCH};

use betteroffice_fonts::Codec;
use serde::Serialize;
use sha2::{Digest, Sha256};

type Result<T> = std::result::Result<T, Box<dyn Error>>;

struct Options {
    src: PathBuf,
    out: PathBuf,
    pack: String,
    version: String,
    level: i32,
    codec: Option<Codec>,
}

#[derive(Serialize)]
struct Manifest {
    format: u32,
    schema: u32,
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
    payload: Option<Vec<u8>>,
    reference: Option<Vec<u8>>,
    codec: Codec,
}

impl Source {
    fn payload(&self) -> &[u8] {
        self.payload.as_deref().unwrap_or(&self.bytes)
    }
}

fn invalid(message: impl Into<String>) -> io::Error {
    io::Error::new(io::ErrorKind::InvalidInput, message.into())
}

fn options() -> Result<Options> {
    let mut args = env::args_os().skip(1);
    if args.next().as_deref() != Some(std::ffi::OsStr::new("build")) {
        return Err(invalid(concat!(
            "usage: fontpack build --src <dir> --out <dir> ",
            "--pack <name> --version <semver> [--level 19] [--codec zstd|woff2t|auto]"
        ))
        .into());
    }
    let mut values = BTreeMap::<String, OsString>::new();
    while let Some(arg) = args.next() {
        let key = arg.into_string().map_err(|_| invalid("invalid option"))?;
        if ![
            "--src",
            "--out",
            "--pack",
            "--version",
            "--level",
            "--codec",
        ]
        .contains(&key.as_str())
        {
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
    let codec = values
        .remove("--codec")
        .map(|value| value.into_string().map_err(|_| invalid("invalid codec")))
        .transpose()?
        .unwrap_or_else(|| "zstd".to_owned());
    let codec = match codec.as_str() {
        "zstd" => Some(Codec::Zstd),
        "woff2t" => Some(Codec::Woff2tZstd),
        "auto" => None,
        _ => return Err(invalid("codec must be zstd, woff2t or auto").into()),
    };
    Ok(Options {
        src,
        out,
        pack,
        version,
        level,
        codec,
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
        let bytes = fs::read(&path)?;
        let (payload, reference) = if options.codec != Some(Codec::Zstd)
            && (options.codec.is_some() || extension.eq_ignore_ascii_case("ttf"))
        {
            if !extension.eq_ignore_ascii_case("ttf") {
                return Err(invalid(format!("woff2t requires TTF glyf/loca faces: {file}")).into());
            }
            glyph_count(&bytes)?;
            let payload = fs::read(path.with_extension("uw2"))?;
            let standard = fs::read(path.with_extension("woff2"))?;
            let regenerate = "regenerate companions with scripts/fonts/woff2-transform.py";
            let hash_path = path.with_extension("source.sha256");
            let recorded = fs::read_to_string(hash_path).map_err(|error| {
                invalid(format!(
                    "cannot read source SHA-256 for {file}: {error}; {regenerate}"
                ))
            })?;
            let expected = [
                (bytes.as_slice(), file.clone()),
                (payload.as_slice(), format!("{stem}.uw2")),
                (standard.as_slice(), format!("{stem}.woff2")),
            ]
            .map(|(data, name)| format!("{:x}  {name}", Sha256::digest(data)));
            if !recorded
                .lines()
                .map(str::trim_end)
                .eq(expected.iter().map(String::as_str))
            {
                return Err(
                    invalid(format!("source SHA-256 mismatch for {file}; {regenerate}")).into(),
                );
            }
            let reference = wuff::decompress_woff2(&standard)
                .map_err(|error| invalid(format!("invalid reference WOFF2 for {file}: {error}")))?;
            (Some(payload), Some(reference))
        } else {
            (None, None)
        };
        files.push(Source {
            file,
            family: family.to_owned(),
            regular: style == "Regular",
            bytes,
            codec: if payload.is_some() {
                Codec::Woff2tZstd
            } else {
                Codec::Zstd
            },
            payload,
            reference,
        });
    }
    files.sort_by(|left, right| left.file.cmp(&right.file));
    if files.is_empty() {
        return Err(invalid("source directory has no TTF or OTF faces").into());
    }
    Ok(files)
}

fn glyph_count(bytes: &[u8]) -> Result<u16> {
    if !bytes.starts_with(b"\0\x01\0\0") {
        return Err(invalid("woff2t requires TTF glyf/loca faces").into());
    }
    let count = bytes
        .get(4..6)
        .ok_or_else(|| invalid("truncated sfnt header"))?;
    let count = u16::from_be_bytes([count[0], count[1]]) as usize;
    let mut glyphs = None;
    let mut glyf = false;
    let mut loca = false;
    for index in 0..count {
        let record = bytes
            .get(12 + index * 16..28 + index * 16)
            .ok_or_else(|| invalid("truncated sfnt directory"))?;
        let offset = u32::from_be_bytes(record[8..12].try_into()?) as usize;
        let length = u32::from_be_bytes(record[12..16].try_into()?) as usize;
        let end = offset
            .checked_add(length)
            .ok_or_else(|| invalid("invalid sfnt table range"))?;
        let table = bytes
            .get(offset..end)
            .ok_or_else(|| invalid("truncated sfnt table"))?;
        match &record[..4] {
            b"glyf" => glyf = true,
            b"loca" => loca = true,
            b"CFF " | b"CFF2" => return Err(invalid("woff2t does not support CFF faces").into()),
            b"maxp" => {
                let value = table.get(4..6).ok_or_else(|| invalid("truncated maxp"))?;
                glyphs = Some(u16::from_be_bytes([value[0], value[1]]));
            }
            _ => {}
        }
    }
    if !glyf || !loca {
        return Err(invalid("woff2t requires TTF glyf/loca faces").into());
    }
    glyphs.ok_or_else(|| invalid("sfnt has no maxp glyph count").into())
}

fn dictionary_payload<'a>(source: &Source, dict: Option<&'a Source>) -> Result<Option<&'a [u8]>> {
    if dict.is_some_and(|dict| source.codec != dict.codec) {
        return Err(invalid("face and dictionary codecs must match").into());
    }
    Ok(dict.map(Source::payload))
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

fn select_codecs(sources: &mut [Source], dictionaries: &[Option<usize>], level: i32) -> Result<()> {
    let mut families = BTreeMap::<String, Vec<usize>>::new();
    for (index, source) in sources.iter().enumerate() {
        families
            .entry(source.family.clone())
            .or_default()
            .push(index);
    }
    for indices in families.values() {
        let mut plain_size = 0;
        let mut transformed_size = 0;
        for &index in indices {
            let source = &sources[index];
            let dict = dictionaries[index].map(|index| &sources[index]);
            plain_size += compress(
                &source.bytes,
                dict.map(|source| source.bytes.as_slice()),
                level,
            )?
            .len();
            transformed_size +=
                compress(source.payload(), dictionary_payload(source, dict)?, level)?.len();
        }
        if plain_size <= transformed_size {
            for &index in indices {
                sources[index].codec = Codec::Zstd;
                sources[index].payload = None;
                sources[index].reference = None;
            }
        }
    }
    Ok(())
}

fn build(options: Options) -> Result<()> {
    let mut sources = sources(&options)?;
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
    if options.codec.is_none() {
        select_codecs(&mut sources, &dictionaries, options.level)?;
    }
    let name = options
        .out
        .file_name()
        .and_then(|value| value.to_str())
        .ok_or_else(|| invalid("invalid output directory name"))?;
    let parent = match options.out.parent() {
        Some(parent) if !parent.as_os_str().is_empty() => parent,
        _ => Path::new("."),
    };
    let out = parent.join(name);
    if fs::symlink_metadata(&out).is_ok() {
        return Err(invalid(format!(
            "output {} already exists; remove it before rebuilding",
            out.display()
        ))
        .into());
    }
    fs::create_dir_all(parent)?;
    let staging = staging_dir(parent, name)?;
    let written = write_pack(&options, &sources, &dictionaries, &staging)
        .and_then(|()| fs::rename(&staging, &out).map_err(Into::into));
    if written.is_err() {
        let _ = fs::remove_dir_all(&staging);
    }
    written
}

fn staging_dir(parent: &Path, name: &str) -> Result<PathBuf> {
    let nanos = SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .map_or(0, |elapsed| elapsed.as_nanos());
    for attempt in 0..100u32 {
        let staging = parent.join(format!(
            ".{name}.partial-{}-{nanos:x}-{attempt}",
            std::process::id()
        ));
        match fs::create_dir(&staging) {
            Ok(()) => return Ok(staging),
            Err(error) if error.kind() == io::ErrorKind::AlreadyExists => continue,
            Err(error) => return Err(error.into()),
        }
    }
    Err(io::Error::other("cannot create a unique staging directory").into())
}

fn verify_limit(parts: &[&[u8]]) -> usize {
    parts.iter().map(|part| part.len()).max().unwrap_or(0)
}

fn write_pack(
    options: &Options,
    sources: &[Source],
    dictionaries: &[Option<usize>],
    staging: &Path,
) -> Result<()> {
    let mut faces = Vec::with_capacity(sources.len());
    for (index, source) in sources.iter().enumerate() {
        let dictionary = dictionaries[index].map(|index| &sources[index]);
        let dict = dictionary_payload(source, dictionary)?;
        let encoded = compress(source.payload(), dict, options.level)?;
        let expected = source.reference.as_deref().unwrap_or(&source.bytes);
        let limit = verify_limit(&[
            expected,
            source.bytes.as_slice(),
            source.payload(),
            dict.unwrap_or_default(),
            encoded.as_slice(),
        ]);
        let decoded = betteroffice_fonts::decode_face(&encoded, dict, source.codec, limit)?;
        if decoded.font != expected {
            return Err(
                io::Error::other(format!("round-trip mismatch for {}", source.file)).into(),
            );
        }
        if source.codec == Codec::Woff2tZstd
            && glyph_count(&decoded.font)? != glyph_count(&source.bytes)?
        {
            return Err(invalid(format!("glyph count mismatch for {}", source.file)).into());
        }
        let file = format!("{}.zst", source.file);
        fs::write(staging.join(&file), &encoded)?;
        faces.push(Face {
            file,
            source: source.file.clone(),
            family: source.family.clone(),
            codec: source.codec.as_str(),
            encoded_length: encoded.len(),
            byte_length: decoded.font.len(),
            sha256: format!("{:x}", Sha256::digest(&decoded.font)),
            dict: dictionary.map(|source| format!("{}.zst", source.file)),
        });
    }
    faces.sort_by(|left, right| left.file.cmp(&right.file));
    let manifest = Manifest {
        format: 1,
        schema: 1,
        pack: options.pack.clone(),
        version: options.version.clone(),
        faces,
    };
    let mut json = serde_json::to_vec_pretty(&manifest)?;
    json.push(b'\n');
    fs::write(staging.join("manifest.json"), json)?;
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

#[cfg(test)]
mod tests {
    use super::{Manifest, Source, dictionary_payload, verify_limit};
    use betteroffice_fonts::Codec;

    fn source(codec: Codec) -> Source {
        Source {
            file: "Synthetic-Regular.ttf".to_owned(),
            family: "Synthetic".to_owned(),
            regular: true,
            bytes: vec![1],
            payload: (codec == Codec::Woff2tZstd).then(|| vec![2]),
            reference: None,
            codec,
        }
    }

    #[test]
    fn manifest_has_independent_format_and_schema_versions() {
        let manifest = Manifest {
            format: 1,
            schema: 1,
            pack: "synthetic".to_owned(),
            version: "1.2.3".to_owned(),
            faces: Vec::new(),
        };
        let json = serde_json::to_value(manifest).unwrap();
        assert_eq!(json["format"], 1);
        assert_eq!(json["schema"], 1);
    }

    #[test]
    fn dictionary_codec_must_match_face() {
        let plain = source(Codec::Zstd);
        let transformed = source(Codec::Woff2tZstd);
        assert!(dictionary_payload(&plain, Some(&transformed)).is_err());
        assert!(dictionary_payload(&transformed, Some(&plain)).is_err());
        assert_eq!(
            dictionary_payload(&transformed, Some(&transformed)).unwrap(),
            Some(&[2][..])
        );
    }

    #[test]
    fn verify_limit_covers_a_dictionary_larger_than_the_face() {
        let face = [0u8; 10];
        let dict = [0u8; 40];
        let frame = [0u8; 12];
        assert_eq!(verify_limit(&[&face, &face, &face, &dict, &frame]), 40);
        assert_eq!(verify_limit(&[&face, &face, &face, &[], &[0u8; 50]]), 50);
    }
}
