use std::collections::HashMap;
use std::fmt;
use std::path::{Path, PathBuf};
use std::sync::{Arc, Mutex, OnceLock};

use serde::Deserialize;

use crate::catalog::{BundledFontFace, catalog};
use crate::{
    Decoded, FaceError, LastResortOffice, decode_font_face, resolve_bundled_family_face,
    resolve_glyph_fallback_face, resolve_last_resort_choice, resolve_script_fallback_face,
};

#[derive(Clone, Copy, Debug, Deserialize, Eq, Hash, PartialEq)]
#[serde(rename_all = "lowercase")]
pub enum FontPack {
    Latin,
    Intl,
    Cjk,
}

impl FontPack {
    pub fn as_str(self) -> &'static str {
        match self {
            Self::Latin => "latin",
            Self::Intl => "intl",
            Self::Cjk => "cjk",
        }
    }

    pub fn feature(self) -> &'static str {
        match self {
            Self::Latin => "fonts-latin",
            Self::Intl => "fonts-intl",
            Self::Cjk => "fonts-cjk",
        }
    }
}

impl fmt::Display for FontPack {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        f.write_str(self.as_str())
    }
}

#[derive(Clone, Copy, Debug, Default, Eq, PartialEq)]
pub enum FontLoading {
    #[default]
    OnDemand,
    Preload,
}

#[derive(Clone, Debug, Eq, PartialEq)]
pub struct BundledFontOptions {
    pub packs: Vec<FontPack>,
    pub base_url: Option<PathBuf>,
}

impl Default for BundledFontOptions {
    fn default() -> Self {
        Self {
            packs: vec![FontPack::Latin, FontPack::Intl, FontPack::Cjk],
            base_url: None,
        }
    }
}

#[derive(Clone)]
pub enum FontSource {
    Bundled(BundledFontOptions),
    Backstop,
    Provider(Arc<dyn FontProvider>),
}

impl Default for FontSource {
    fn default() -> Self {
        Self::Bundled(BundledFontOptions::default())
    }
}

impl fmt::Debug for FontSource {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Bundled(options) => f.debug_tuple("Bundled").field(options).finish(),
            Self::Backstop => f.write_str("Backstop"),
            Self::Provider(_) => f.write_str("Provider"),
        }
    }
}

#[derive(Clone, Debug, Eq, PartialEq)]
pub struct FontDefinition {
    pub family: String,
    pub data: Vec<u8>,
    pub bold: bool,
    pub italic: bool,
}

#[derive(Debug)]
pub enum FontError {
    MissingFontPack {
        pack: FontPack,
        action: String,
    },
    InvalidFontOptions(String),
    Decode(FaceError),
    InvalidFontData(String),
    Io {
        path: PathBuf,
        source: std::io::Error,
    },
}

impl fmt::Display for FontError {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::MissingFontPack { pack, action } => {
                write!(f, "missing font pack {pack}: {action}")
            }
            Self::InvalidFontOptions(message) => write!(f, "invalid font options: {message}"),
            Self::Decode(error) => fmt::Display::fmt(error, f),
            Self::InvalidFontData(message) => write!(f, "invalid font data: {message}"),
            Self::Io { path, source } => write!(f, "cannot read font {}: {source}", path.display()),
        }
    }
}

impl std::error::Error for FontError {
    fn source(&self) -> Option<&(dyn std::error::Error + 'static)> {
        match self {
            Self::Decode(error) => Some(error),
            Self::Io { source, .. } => Some(source),
            _ => None,
        }
    }
}

type FontLoader = dyn Fn() -> Result<Vec<u8>, FontError> + Send + Sync;

#[derive(Clone)]
pub struct FontFaceHandle {
    pub family: String,
    pub bold: bool,
    pub italic: bool,
    pub measure_as: Option<String>,
    pub only_if_primary_fails: bool,
    loader: Arc<FontLoader>,
}

impl fmt::Debug for FontFaceHandle {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        f.debug_struct("FontFaceHandle")
            .field("family", &self.family)
            .field("bold", &self.bold)
            .field("italic", &self.italic)
            .field("measure_as", &self.measure_as)
            .field("only_if_primary_fails", &self.only_if_primary_fails)
            .finish_non_exhaustive()
    }
}

impl FontFaceHandle {
    pub fn new(
        family: impl Into<String>,
        bold: bool,
        italic: bool,
        loader: impl Fn() -> Result<Vec<u8>, FontError> + Send + Sync + 'static,
    ) -> Self {
        Self {
            family: family.into(),
            bold,
            italic,
            measure_as: None,
            only_if_primary_fails: false,
            loader: Arc::new(loader),
        }
    }

    pub fn load(&self) -> Result<Vec<u8>, FontError> {
        (self.loader)()
    }
}

pub trait FontProvider: Send + Sync {
    fn resolve_family(&self, family: &str, bold: bool, italic: bool) -> Option<FontFaceHandle>;
    fn resolve_script_fallback(
        &self,
        script: &str,
        bold: bool,
        italic: bool,
    ) -> Option<FontFaceHandle>;
    fn resolve_glyph_fallback(
        &self,
        family: &str,
        bold: bool,
        italic: bool,
        office: LastResortOffice,
    ) -> Option<FontFaceHandle>;
    fn resolve_last_resort(
        &self,
        family: &str,
        bold: bool,
        italic: bool,
        office: LastResortOffice,
    ) -> FontFaceHandle;
}

impl FontSource {
    pub fn provider(&self) -> Result<Option<Arc<dyn FontProvider>>, FontError> {
        match self {
            Self::Provider(provider) => Ok(Some(Arc::clone(provider))),
            Self::Bundled(options) => Ok(Some(Arc::new(BundledFonts::new(options.clone())?))),
            Self::Backstop => Ok(None),
        }
    }
}

#[derive(Clone, Debug, Default)]
pub struct BundledFonts {
    options: BundledFontOptions,
}

impl BundledFonts {
    pub fn new(mut options: BundledFontOptions) -> Result<Self, FontError> {
        for (index, pack) in options.packs.iter().enumerate() {
            if options.packs[..index].contains(pack) {
                return Err(FontError::InvalidFontOptions(format!(
                    "duplicate {pack} pack"
                )));
            }
        }
        if let Some(base) = &options.base_url {
            if base.as_os_str().is_empty() || base.to_string_lossy().contains("://") {
                return Err(FontError::InvalidFontOptions(
                    "base_url must be a local directory".into(),
                ));
            }
            if base.is_relative() {
                options.base_url = Some(
                    std::env::current_dir()
                        .map_err(|source| FontError::Io {
                            path: base.clone(),
                            source,
                        })?
                        .join(base),
                );
            }
        }
        Ok(Self { options })
    }

    pub fn preload(&self) -> Result<(), FontError> {
        for face in &catalog().faces {
            if self.options.packs.contains(&face.pack) {
                load_decoded(&self.options, face)?;
            }
        }
        Ok(())
    }

    pub fn load_face(&self, face: &BundledFontFace) -> Result<Vec<u8>, FontError> {
        let known = catalog()
            .faces
            .iter()
            .find(|known| *known == face)
            .ok_or_else(|| {
                FontError::InvalidFontOptions("face is outside the bundled catalog".into())
            })?;
        Ok(load_decoded(&self.options, known)?.font.clone())
    }

    fn handle(&self, face: &'static BundledFontFace) -> FontFaceHandle {
        let options = self.options.clone();
        let mut handle = FontFaceHandle::new(
            &face.family,
            face.weight == 700,
            face.style == "italic",
            move || Ok(load_decoded(&options, face)?.font.clone()),
        );
        if face.family == "Selawik" {
            handle.measure_as = face.metric_compat_with.clone();
        }
        handle
    }
}

impl FontProvider for BundledFonts {
    fn resolve_family(&self, family: &str, bold: bool, italic: bool) -> Option<FontFaceHandle> {
        resolve_bundled_family_face(family, bold, italic)
            .filter(|face| self.options.packs.contains(&face.pack))
            .map(|face| self.handle(face))
    }

    fn resolve_script_fallback(
        &self,
        script: &str,
        bold: bool,
        italic: bool,
    ) -> Option<FontFaceHandle> {
        resolve_script_fallback_face(script, bold, italic)
            .filter(|face| self.options.packs.contains(&face.pack))
            .map(|face| self.handle(face))
    }

    fn resolve_glyph_fallback(
        &self,
        family: &str,
        bold: bool,
        italic: bool,
        office: LastResortOffice,
    ) -> Option<FontFaceHandle> {
        resolve_glyph_fallback_face(family, bold, italic, office)
            .filter(|face| self.options.packs.contains(&face.pack))
            .map(|face| self.handle(face))
    }

    fn resolve_last_resort(
        &self,
        family: &str,
        bold: bool,
        italic: bool,
        office: LastResortOffice,
    ) -> FontFaceHandle {
        let choice = resolve_last_resort_choice(family, bold, italic, office);
        let mut handle = self.handle(choice.face);
        handle.only_if_primary_fails = choice.only_if_primary_fails;
        handle
    }
}

#[derive(Eq, Hash, PartialEq)]
struct CacheKey {
    base: Option<PathBuf>,
    pack: FontPack,
    file: &'static str,
}

type FaceCell = Mutex<Option<Arc<Decoded>>>;
type FaceCache = HashMap<CacheKey, Arc<FaceCell>>;

fn missing(pack: FontPack, base: Option<&Path>, file: &str, version: &str) -> FontError {
    let action = match base {
        Some(base) => format!(
            "deploy {}",
            base.join(pack.as_str()).join(version).join(file).display()
        ),
        None => format!("enable the `{}` Cargo feature", pack.feature()),
    };
    FontError::MissingFontPack { pack, action }
}

fn encoded(options: &BundledFontOptions, face: &BundledFontFace) -> Result<Vec<u8>, FontError> {
    if !options.packs.contains(&face.pack) {
        return Err(FontError::MissingFontPack {
            pack: face.pack,
            action: format!(
                "select the {0} pack in BundledFontOptions and enable `{1}` or deploy its base_url files",
                face.pack,
                face.pack.feature()
            ),
        });
    }
    if let Some(base) = &options.base_url {
        let path = base
            .join(face.pack.as_str())
            .join(&face.version)
            .join(&face.file);
        return std::fs::read(&path).map_err(|source| {
            if source.kind() == std::io::ErrorKind::NotFound {
                missing(face.pack, Some(base), &face.file, &face.version)
            } else {
                FontError::Io { path, source }
            }
        });
    }
    let bytes: Option<&[u8]> = match face.pack {
        FontPack::Latin => {
            #[cfg(feature = "fonts-latin")]
            {
                betteroffice_fonts_latin::face(&face.file)
            }
            #[cfg(not(feature = "fonts-latin"))]
            {
                None
            }
        }
        FontPack::Intl => {
            #[cfg(feature = "fonts-intl")]
            {
                betteroffice_fonts_intl::face(&face.file)
            }
            #[cfg(not(feature = "fonts-intl"))]
            {
                None
            }
        }
        FontPack::Cjk => {
            #[cfg(feature = "fonts-cjk")]
            {
                betteroffice_fonts_cjk::face(&face.file)
            }
            #[cfg(not(feature = "fonts-cjk"))]
            {
                None
            }
        }
    };
    bytes
        .map(<[u8]>::to_vec)
        .ok_or_else(|| missing(face.pack, None, &face.file, &face.version))
}

fn load_decoded(
    options: &BundledFontOptions,
    face: &'static BundledFontFace,
) -> Result<Arc<Decoded>, FontError> {
    if !options.packs.contains(&face.pack) {
        encoded(options, face)?;
    }
    static CACHE: OnceLock<Mutex<FaceCache>> = OnceLock::new();
    let cell = {
        let mut cache = CACHE
            .get_or_init(Mutex::default)
            .lock()
            .unwrap_or_else(|error| error.into_inner());
        cache
            .entry(CacheKey {
                base: options.base_url.clone(),
                pack: face.pack,
                file: &face.file,
            })
            .or_default()
            .clone()
    };
    let mut cached = cell.lock().unwrap_or_else(|error| error.into_inner());
    if let Some(decoded) = cached.as_ref() {
        return Ok(Arc::clone(decoded));
    }
    let dictionary = if let Some(file) = &face.dict {
        let dict = catalog()
            .faces
            .iter()
            .find(|candidate| candidate.pack == face.pack && &candidate.file == file)
            .filter(|dict| dict.dict.is_none() && dict.codec == face.codec)
            .ok_or_else(|| {
                FontError::InvalidFontData(format!("invalid dictionary for {}", face.file))
            })?;
        Some(load_decoded(options, dict)?)
    } else {
        None
    };
    let bytes = encoded(options, face)?;
    if bytes.len() != face.encoded_length {
        return Err(FontError::InvalidFontData(format!(
            "encoded length mismatch for {}",
            face.file
        )));
    }
    let dict = dictionary
        .as_ref()
        .map(|dict| dict.payload.as_deref().unwrap_or(&dict.font));
    let decoded = decode_font_face(&bytes, &face.codec, dict, None).map_err(FontError::Decode)?;
    if decoded.font.len() != face.byte_length {
        return Err(FontError::InvalidFontData(format!(
            "decoded length mismatch for {}",
            face.file
        )));
    }
    let decoded = Arc::new(decoded);
    *cached = Some(Arc::clone(&decoded));
    Ok(decoded)
}
