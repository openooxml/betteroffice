use super::{ProjectionFamily, ProjectionHooks};

pub(super) struct Hooks;
pub(super) type StoryData = ();

impl ProjectionHooks for Hooks {
    fn family(&self) -> ProjectionFamily {
        ProjectionFamily::Tables
    }
}
