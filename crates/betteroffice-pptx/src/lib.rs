//! Typed native facade for opening, editing, rendering, and saving PPTX files.

mod error;
mod presentation;
mod types;

pub use error::Error;
pub use presentation::Presentation;
pub use types::*;

pub type Result<T> = std::result::Result<T, Error>;
