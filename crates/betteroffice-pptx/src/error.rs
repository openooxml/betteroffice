use std::fmt;

#[derive(Debug)]
#[non_exhaustive]
pub enum Error {
    Parse(pptx_parse::PptxError),
    Edit(pptx_edit::EditError),
    Render(pptx_render::RenderError),
    /// The raster backend refused a surface or could not paint a primitive.
    Raster(String),
}

impl fmt::Display for Error {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Parse(error) => error.fmt(formatter),
            Self::Edit(error) => error.fmt(formatter),
            Self::Render(error) => error.fmt(formatter),
            Self::Raster(message) => formatter.write_str(message),
        }
    }
}

impl std::error::Error for Error {
    fn source(&self) -> Option<&(dyn std::error::Error + 'static)> {
        match self {
            Self::Parse(error) => Some(error),
            Self::Edit(error) => Some(error),
            Self::Render(error) => Some(error),
            Self::Raster(_) => None,
        }
    }
}

impl From<pptx_parse::PptxError> for Error {
    fn from(error: pptx_parse::PptxError) -> Self {
        Self::Parse(error)
    }
}

impl From<pptx_edit::EditError> for Error {
    fn from(error: pptx_edit::EditError) -> Self {
        Self::Edit(error)
    }
}

impl From<pptx_render::RenderError> for Error {
    fn from(error: pptx_render::RenderError) -> Self {
        Self::Render(error)
    }
}
