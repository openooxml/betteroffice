use serde_json::{Number, Value};

const MAX_JSON_NESTING: usize = 1024;

pub(crate) enum OrderedValue {
    Null,
    Bool(bool),
    Number(Number),
    String(String),
    Array(Vec<Self>),
    Object(Vec<(String, Self)>),
}

impl OrderedValue {
    pub(crate) fn into_value(self) -> Result<Value, String> {
        enum Step {
            Visit(OrderedValue),
            Array(usize),
            Object(Vec<String>),
        }

        let mut pending = vec![Step::Visit(self)];
        let mut values = Vec::new();
        while let Some(step) = pending.pop() {
            let value = match step {
                Step::Visit(Self::Array(children)) => {
                    pending.push(Step::Array(children.len()));
                    pending.extend(children.into_iter().rev().map(Step::Visit));
                    continue;
                }
                Step::Visit(Self::Object(entries)) => {
                    let (keys, children): (Vec<_>, Vec<_>) = entries.into_iter().unzip();
                    pending.push(Step::Object(keys));
                    pending.extend(children.into_iter().rev().map(Step::Visit));
                    continue;
                }
                Step::Visit(scalar) => scalar.value(),
                Step::Array(count) => Value::Array(values.split_off(values.len() - count)),
                Step::Object(keys) => {
                    let children = values.split_off(values.len() - keys.len());
                    let mut entries = serde_json::Map::new();
                    for (key, child) in keys.into_iter().zip(children) {
                        if entries.contains_key(&key) {
                            return Err(format!("duplicate JSON object key {key:?}"));
                        }
                        entries.insert(key, child);
                    }
                    Value::Object(entries)
                }
            };
            values.push(value);
        }
        Ok(values.pop().expect("one parsed JSON value"))
    }

    pub(crate) fn value(&self) -> Value {
        match self {
            Self::Null => Value::Null,
            Self::Bool(value) => Value::Bool(*value),
            Self::Number(value) => Value::Number(value.clone()),
            Self::String(value) => Value::String(value.clone()),
            Self::Array(values) => Value::Array(values.iter().map(Self::value).collect()),
            Self::Object(entries) => Value::Object(
                entries
                    .iter()
                    .map(|(key, value)| (key.clone(), value.value()))
                    .collect(),
            ),
        }
    }
}

enum Frame {
    Array(Vec<OrderedValue>),
    Object(Vec<(String, OrderedValue)>, String),
}

impl Frame {
    fn end(&self) -> char {
        match self {
            Self::Array(_) => ']',
            Self::Object(..) => '}',
        }
    }

    fn finish(self) -> OrderedValue {
        match self {
            Self::Array(values) => OrderedValue::Array(values),
            Self::Object(entries, _) => OrderedValue::Object(entries),
        }
    }
}

fn skip_space(input: &mut &str) {
    *input = input.trim_start_matches([' ', '\t', '\r', '\n']);
}

fn take(input: &mut &str, expected: char) -> bool {
    skip_space(input);
    if let Some(rest) = input.strip_prefix(expected) {
        *input = rest;
        true
    } else {
        false
    }
}

fn scalar(input: &mut &str) -> Result<OrderedValue, String> {
    skip_space(input);
    let mut bytes = input.bytes().enumerate();
    let end = if input.starts_with('"') {
        bytes.next();
        loop {
            match bytes.next() {
                Some((index, b'"')) => break index + 1,
                Some((_, b'\\')) => {
                    bytes.next();
                }
                Some(_) => {}
                None => return Err("unterminated JSON string".to_owned()),
            }
        }
    } else {
        bytes
            .find(|(_, byte)| {
                matches!(
                    byte,
                    b' ' | b'\t' | b'\r' | b'\n' | b',' | b']' | b'}' | b'|'
                )
            })
            .map_or(input.len(), |(index, _)| index)
    };
    let (token, rest) = input.split_at(end);
    if token.starts_with(['[', '{']) {
        return Err("expected a JSON scalar".to_owned());
    }
    let value = serde_json::from_str::<Value>(token).map_err(|error| error.to_string())?;
    *input = rest;
    Ok(match value {
        Value::Null => OrderedValue::Null,
        Value::Bool(value) => OrderedValue::Bool(value),
        Value::Number(value) => OrderedValue::Number(value),
        Value::String(value) => OrderedValue::String(value),
        _ => unreachable!("collections are parsed separately"),
    })
}

fn object_key(input: &mut &str) -> Result<String, String> {
    let OrderedValue::String(key) = scalar(input)? else {
        return Err("expected a JSON object key".to_owned());
    };
    if !take(input, ':') {
        return Err("expected ':' after a JSON object key".to_owned());
    }
    Ok(key)
}

/// Parses JSON containers on a heap stack and preserves object member order.
pub(crate) fn parse_prefix(mut input: &str) -> Result<(OrderedValue, &str), String> {
    let mut stack = Vec::new();
    loop {
        let mut value = if take(&mut input, '[') {
            if stack.len() == MAX_JSON_NESTING {
                return Err(format!("JSON nesting exceeds {MAX_JSON_NESTING}"));
            }
            if !take(&mut input, ']') {
                stack.push(Frame::Array(Vec::new()));
                continue;
            }
            OrderedValue::Array(Vec::new())
        } else if take(&mut input, '{') {
            if stack.len() == MAX_JSON_NESTING {
                return Err(format!("JSON nesting exceeds {MAX_JSON_NESTING}"));
            }
            if !take(&mut input, '}') {
                stack.push(Frame::Object(Vec::new(), object_key(&mut input)?));
                continue;
            }
            OrderedValue::Object(Vec::new())
        } else {
            scalar(&mut input)?
        };
        loop {
            let Some(frame) = stack.last_mut() else {
                return Ok((value, input));
            };
            match frame {
                Frame::Array(values) => values.push(value),
                Frame::Object(entries, key) => entries.push((std::mem::take(key), value)),
            }
            if take(&mut input, ',') {
                if let Frame::Object(_, key) = frame {
                    *key = object_key(&mut input)?;
                }
                break;
            }
            if !take(&mut input, frame.end()) {
                return Err(format!("expected ',' or '{}' in JSON", frame.end()));
            }
            value = stack.pop().expect("open JSON container").finish();
        }
    }
}

pub(crate) fn parse(input: &str) -> Result<OrderedValue, String> {
    let (value, mut rest) = parse_prefix(input)?;
    skip_space(&mut rest);
    if !rest.is_empty() {
        return Err("trailing characters after JSON value".to_owned());
    }
    Ok(value)
}
