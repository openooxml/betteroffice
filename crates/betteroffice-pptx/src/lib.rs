//! Typed native facade for opening, editing, rendering, and saving PPTX files.

mod error;
mod presentation;
mod types;

pub use error::Error;
pub use presentation::Presentation;
pub use types::*;

/// Largest collaboration update or state vector the engine will decode.
pub const MAX_COLLABORATION_BYTES: usize = pptx_edit::MAX_UPDATE_BYTES;
/// Largest collaboration client ID the engine accepts; 0 is rejected too.
pub const MAX_COLLABORATION_CLIENT_ID: u64 = pptx_edit::MAX_SAFE_CLIENT_ID;

pub type Result<T> = std::result::Result<T, Error>;
