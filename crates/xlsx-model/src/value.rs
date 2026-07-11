//! cell values. numbers are f64 like excel; dates are numbers + a number
//! format, never a distinct storage type.

use serde::{Deserialize, Serialize};

#[derive(Debug, Clone, PartialEq, Default, Serialize, Deserialize)]
#[serde(tag = "kind", rename_all = "camelCase")]
pub enum CellValue {
    #[default]
    Empty,
    Number {
        value: f64,
    },
    Text {
        value: String,
    },
    Bool {
        value: bool,
    },
    Error {
        value: ErrorValue,
    },
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum ErrorValue {
    #[serde(rename = "#DIV/0!")]
    Div0,
    #[serde(rename = "#N/A")]
    NA,
    #[serde(rename = "#NAME?")]
    Name,
    #[serde(rename = "#NULL!")]
    Null,
    #[serde(rename = "#NUM!")]
    Num,
    #[serde(rename = "#REF!")]
    Ref,
    #[serde(rename = "#VALUE!")]
    Value,
    #[serde(rename = "#SPILL!")]
    Spill,
}

impl ErrorValue {
    pub fn as_str(&self) -> &'static str {
        match self {
            ErrorValue::Div0 => "#DIV/0!",
            ErrorValue::NA => "#N/A",
            ErrorValue::Name => "#NAME?",
            ErrorValue::Null => "#NULL!",
            ErrorValue::Num => "#NUM!",
            ErrorValue::Ref => "#REF!",
            ErrorValue::Value => "#VALUE!",
            ErrorValue::Spill => "#SPILL!",
        }
    }
}
