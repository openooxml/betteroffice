use super::args::parse_raw_ops;
use super::{EditSessionCore, EngineError, EngineResult};
use crate::EditCtx;
use crate::engine::RelayoutTrigger;

impl EditSessionCore {
    pub fn apply_raw_ops(&self, story: &str, ops_json: &str) -> EngineResult<()> {
        let ops = parse_raw_ops(ops_json)?;
        let ctx = EditCtx::local(String::new(), String::new());
        self.engine
            .doc()
            .apply_raw_ops(story, ops, &ctx)
            .map_err(EngineError::message)
    }

    pub fn apply_seed_raw_ops(&self, story: &str, ops_json: &str) -> EngineResult<()> {
        let ops = parse_raw_ops(ops_json)?;
        let ctx = EditCtx::local(String::new(), String::new());
        self.engine
            .doc()
            .apply_raw_story_batches(vec![(story.to_owned(), ops)], &ctx)
            .map_err(EngineError::message)?;
        self.engine.set_relayout_trigger(RelayoutTrigger::Open);
        Ok(())
    }
}
