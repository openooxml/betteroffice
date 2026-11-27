use std::collections::BTreeSet;

use serde_json::Value;

pub fn json_string_paths(value: &Value, skipped: &BTreeSet<&str>) -> BTreeSet<String> {
    fn visit(value: &Value, path: &str, skipped: &BTreeSet<&str>, paths: &mut BTreeSet<String>) {
        if skipped.contains(path) {
            return;
        }
        match value {
            Value::String(_) => {
                paths.insert(path.to_owned());
            }
            Value::Object(record) => {
                for (key, value) in record {
                    let key = key.replace('~', "~0").replace('/', "~1");
                    visit(value, &format!("{path}/{key}"), skipped, paths);
                }
            }
            Value::Array(items) => {
                for (index, value) in items.iter().enumerate() {
                    visit(value, &format!("{path}/{index}"), skipped, paths);
                }
            }
            Value::Null | Value::Bool(_) | Value::Number(_) => {}
        }
    }

    let mut paths = BTreeSet::new();
    visit(value, "", skipped, &mut paths);
    paths
}
