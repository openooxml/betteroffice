use serde_json::Value;

use super::args::{edit_ctx, loc_index, parse_any_object};
use super::{EditSessionCore, EngineError, EngineResult};
use crate::{CellLoc, EditCtx, EditingDoc, Position, TableLocator, TableRange};

impl EditSessionCore {
    #[allow(clippy::too_many_arguments)]
    pub fn insert_table(
        &self,
        story: &str,
        para_id: &str,
        offset: u32,
        rows: u32,
        columns: u32,
        author_name: Option<String>,
        author_date: Option<String>,
    ) -> EngineResult<String> {
        let index = loc_index(self.engine.doc(), story, para_id, offset)?;
        let ctx = edit_ctx(author_name, author_date)?;
        let receipt = self
            .engine
            .doc()
            .insert_table(&ctx, Position::new(story, index), rows, columns)
            .map_err(EngineError::message)?;
        serde_json::to_string(&receipt).map_err(EngineError::message)
    }

    pub fn insert_table_fragment(
        &self,
        story: &str,
        para_id: &str,
        offset: u32,
        fragment_json: &str,
        author_name: Option<String>,
        author_date: Option<String>,
    ) -> EngineResult<String> {
        let index = loc_index(self.engine.doc(), story, para_id, offset)?;
        let ctx = edit_ctx(author_name, author_date)?;
        let fragment: Value = serde_json::from_str(fragment_json).map_err(EngineError::message)?;
        let receipt = self
            .engine
            .doc()
            .insert_table_fragment(&ctx, Position::new(story, index), &fragment)
            .map_err(EngineError::message)?;
        serde_json::to_string(&receipt).map_err(EngineError::message)
    }

    pub fn validate_table_fragment(&self, fragment_json: &str) -> EngineResult<()> {
        let fragment: Value = serde_json::from_str(fragment_json).map_err(EngineError::message)?;
        EditingDoc::validate_table_fragment(&fragment).map_err(EngineError::message)
    }

    pub fn insert_row(
        &self,
        at_json: &str,
        after: bool,
        author_name: Option<String>,
        author_date: Option<String>,
    ) -> EngineResult<String> {
        let at: CellLoc = serde_json::from_str(at_json).map_err(EngineError::message)?;
        let ctx = edit_ctx(author_name, author_date)?;
        let receipt = self
            .engine
            .doc()
            .insert_row(&ctx, &at, after)
            .map_err(EngineError::message)?;
        serde_json::to_string(&receipt).map_err(EngineError::message)
    }

    pub fn insert_column(&self, at_json: &str, after: bool) -> EngineResult<String> {
        let at: CellLoc = serde_json::from_str(at_json).map_err(EngineError::message)?;
        let receipt = self
            .engine
            .doc()
            .insert_column(&EditCtx::local("", ""), &at, after)
            .map_err(EngineError::message)?;
        serde_json::to_string(&receipt).map_err(EngineError::message)
    }

    pub fn delete_row(
        &self,
        range_json: &str,
        author_name: Option<String>,
        author_date: Option<String>,
    ) -> EngineResult<String> {
        let range: TableRange = serde_json::from_str(range_json).map_err(EngineError::message)?;
        let ctx = edit_ctx(author_name, author_date)?;
        let receipt = self
            .engine
            .doc()
            .delete_row(&ctx, &range)
            .map_err(EngineError::message)?;
        serde_json::to_string(&receipt).map_err(EngineError::message)
    }

    pub fn delete_column(&self, range_json: &str) -> EngineResult<String> {
        let range: TableRange = serde_json::from_str(range_json).map_err(EngineError::message)?;
        let receipt = self
            .engine
            .doc()
            .delete_column(&EditCtx::local("", ""), &range)
            .map_err(EngineError::message)?;
        serde_json::to_string(&receipt).map_err(EngineError::message)
    }

    pub fn delete_table(&self, table_json: &str) -> EngineResult<String> {
        let table: TableLocator = serde_json::from_str(table_json).map_err(EngineError::message)?;
        let receipt = self
            .engine
            .doc()
            .delete_table(&EditCtx::local("", ""), &table)
            .map_err(EngineError::message)?;
        serde_json::to_string(&receipt).map_err(EngineError::message)
    }

    pub fn merge_cells(&self, range_json: &str) -> EngineResult<String> {
        let range: TableRange = serde_json::from_str(range_json).map_err(EngineError::message)?;
        let receipt = self
            .engine
            .doc()
            .merge_cells(&EditCtx::local("", ""), &range)
            .map_err(EngineError::message)?;
        serde_json::to_string(&receipt).map_err(EngineError::message)
    }

    pub fn split_cell(
        &self,
        at_json: &str,
        rows: Option<u32>,
        columns: Option<u32>,
    ) -> EngineResult<String> {
        let at: CellLoc = serde_json::from_str(at_json).map_err(EngineError::message)?;
        let receipt = self
            .engine
            .doc()
            .split_cell_grid(&EditCtx::local("", ""), &at, rows, columns)
            .map_err(EngineError::message)?;
        serde_json::to_string(&receipt).map_err(EngineError::message)
    }

    pub fn set_cell_shading(
        &self,
        range_json: &str,
        color: Option<String>,
    ) -> EngineResult<String> {
        let range: TableRange = serde_json::from_str(range_json).map_err(EngineError::message)?;
        let receipt = self
            .engine
            .doc()
            .set_cell_shading(&EditCtx::local("", ""), &range, color.as_deref())
            .map_err(EngineError::message)?;
        serde_json::to_string(&receipt).map_err(EngineError::message)
    }

    pub fn set_cell_text_format(&self, range_json: &str, patch_json: &str) -> EngineResult<String> {
        let range: TableRange = serde_json::from_str(range_json).map_err(EngineError::message)?;
        let patch = parse_any_object(patch_json, "cell format patch")?;
        let receipt = self
            .engine
            .doc()
            .set_cell_text_format(&EditCtx::local("", ""), &range, &patch)
            .map_err(EngineError::message)?;
        serde_json::to_string(&receipt).map_err(EngineError::message)
    }

    pub fn set_cell_borders(&self, range_json: &str, borders_json: &str) -> EngineResult<String> {
        let range: TableRange = serde_json::from_str(range_json).map_err(EngineError::message)?;
        let borders = parse_any_object(borders_json, "cell borders")?;
        let receipt = self
            .engine
            .doc()
            .set_cell_borders(&EditCtx::local("", ""), &range, &borders)
            .map_err(EngineError::message)?;
        serde_json::to_string(&receipt).map_err(EngineError::message)
    }

    pub fn set_column_width(&self, at_json: &str, width_twips: f64) -> EngineResult<String> {
        let at: CellLoc = serde_json::from_str(at_json).map_err(EngineError::message)?;
        let receipt = self
            .engine
            .doc()
            .set_column_width(&EditCtx::local("", ""), &at, width_twips)
            .map_err(EngineError::message)?;
        serde_json::to_string(&receipt).map_err(EngineError::message)
    }

    pub fn set_table_width(&self, table_json: &str, width_twips: f64) -> EngineResult<String> {
        let table: TableLocator = serde_json::from_str(table_json).map_err(EngineError::message)?;
        let receipt = self
            .engine
            .doc()
            .set_table_width(&EditCtx::local("", ""), &table, width_twips)
            .map_err(EngineError::message)?;
        serde_json::to_string(&receipt).map_err(EngineError::message)
    }

    pub fn set_table_properties(&self, table_json: &str, patch_json: &str) -> EngineResult<String> {
        let table: TableLocator = serde_json::from_str(table_json).map_err(EngineError::message)?;
        let patch = parse_any_object(patch_json, "table properties patch")?;
        let receipt = self
            .engine
            .doc()
            .set_table_properties(&EditCtx::local("", ""), &table, &patch)
            .map_err(EngineError::message)?;
        serde_json::to_string(&receipt).map_err(EngineError::message)
    }
}
