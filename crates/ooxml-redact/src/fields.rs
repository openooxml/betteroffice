use crate::RedactError;
use crate::identifiers::{self, InstructionEdit, instruction_edits};
use crate::mask::TextMasker;
use crate::styles::StyleMap;

/// Rewrites complex-field instructions in their original text slots.
pub(crate) struct Fields<'a> {
    path: &'a str,
    styles: &'a StyleMap,
    slots: Vec<(usize, String)>,
    open: Vec<Field>,
}

#[derive(Default)]
struct Field {
    instruction: String,
    slots: Vec<(usize, std::ops::Range<usize>)>,
    breaks: Vec<usize>,
    closed: bool,
}

impl<'a> Fields<'a> {
    pub(crate) fn new(path: &'a str, styles: &'a StyleMap) -> Self {
        Self {
            path,
            styles,
            slots: Vec::new(),
            open: Vec::new(),
        }
    }

    pub(crate) fn begin(&mut self) {
        if let Some(field) = self.open.last_mut().filter(|field| !field.closed) {
            field.breaks.push(field.instruction.len());
        }
        self.open.push(Field::default());
    }

    pub(crate) fn separate(&mut self, masker: &mut TextMasker) -> Result<(), RedactError> {
        let Some(field) = self.open.last_mut().filter(|field| !field.closed) else {
            return Ok(());
        };
        field.closed = true;
        let field = Field {
            instruction: std::mem::take(&mut field.instruction),
            slots: std::mem::take(&mut field.slots),
            breaks: std::mem::take(&mut field.breaks),
            closed: true,
        };
        self.apply(&field, masker)
    }

    pub(crate) fn end(&mut self, masker: &mut TextMasker) -> Result<(), RedactError> {
        match self.open.pop() {
            Some(field) if !field.closed => self.apply(&field, masker),
            _ => Ok(()),
        }
    }

    pub(crate) fn chunk(
        &mut self,
        offset: usize,
        original: &str,
        fallback: String,
        masker: &mut TextMasker,
    ) -> Result<(), RedactError> {
        let slot = self.slots.len();
        self.slots.push((offset, fallback));
        match self.open.last_mut().filter(|field| !field.closed) {
            Some(field) => {
                let start = field.instruction.len();
                field.instruction.push_str(original);
                field.slots.push((slot, start..field.instruction.len()));
                Ok(())
            }
            None => self.apply(
                &Field {
                    instruction: original.to_owned(),
                    slots: vec![(slot, 0..original.len())],
                    breaks: Vec::new(),
                    closed: true,
                },
                masker,
            ),
        }
    }

    fn apply(&mut self, field: &Field, masker: &mut TextMasker) -> Result<(), RedactError> {
        if let Some(edits) = self.edits(field, masker)? {
            for (slot, range) in &field.slots {
                self.slots[*slot].1 = field.instruction[range.clone()].to_owned();
            }
            for edit in edits.iter().rev() {
                for (slot, range) in &field.slots {
                    let start = edit.offset.max(range.start);
                    let end = (edit.offset + edit.length).min(range.end);
                    if start < end || (edit.length == 0 && range.contains(&edit.offset)) {
                        let replacement = if range.contains(&edit.offset) {
                            edit.replacement.as_str()
                        } else {
                            ""
                        };
                        self.slots[*slot]
                            .1
                            .replace_range(start - range.start..end - range.start, replacement);
                    }
                }
            }
        }
        Ok(())
    }

    /// Bookmark edits, else style edits; nested style fields keep the masked fallback.
    fn edits(
        &self,
        field: &Field,
        masker: &mut TextMasker,
    ) -> Result<Option<Vec<InstructionEdit>>, RedactError> {
        if let Some(edits) = instruction_edits(&field.instruction, &field.breaks, masker)? {
            return Ok(Some(edits));
        }
        if !field.breaks.is_empty() {
            return Ok(None);
        }
        let Some(rewritten) = self
            .styles
            .rewrite_field_instruction(self.path, &field.instruction)?
        else {
            return Ok(None);
        };
        let original = field.instruction.as_str();
        let prefix = original
            .char_indices()
            .zip(rewritten.chars())
            .find(|((_, a), b)| a != b)
            .map_or(original.len().min(rewritten.len()), |((index, _), _)| index);
        let suffix = original[prefix..]
            .chars()
            .rev()
            .zip(rewritten[prefix..].chars().rev())
            .take_while(|(a, b)| a == b)
            .map(|(a, _)| a.len_utf8())
            .sum::<usize>();
        let end = original.len() - suffix;
        Ok(Some(vec![InstructionEdit {
            offset: prefix,
            length: end - prefix,
            replacement: rewritten[prefix..rewritten.len() - suffix].to_owned(),
        }]))
    }

    pub(crate) fn finish(
        mut self,
        output: Vec<u8>,
        masker: &mut TextMasker,
    ) -> Result<Vec<u8>, RedactError> {
        while let Some(field) = self.open.pop() {
            if !field.closed {
                self.apply(&field, masker)?;
            }
        }
        if self.slots.is_empty() {
            return Ok(output);
        }
        let mut spliced = Vec::with_capacity(output.len() + self.slots.len() * 8);
        let mut copied = 0;
        for (offset, text) in &self.slots {
            spliced.extend_from_slice(&output[copied..*offset]);
            spliced.extend_from_slice(quick_xml::escape::escape(text.as_str()).as_bytes());
            copied = *offset;
        }
        spliced.extend_from_slice(&output[copied..]);
        Ok(spliced)
    }
}

pub(crate) fn rewrite_instruction(
    path: &str,
    instruction: &str,
    styles: &StyleMap,
    masker: &mut TextMasker,
) -> Result<Option<String>, RedactError> {
    match identifiers::rewrite_instruction(instruction, masker)? {
        Some(rewritten) => Ok(Some(rewritten)),
        None => styles.rewrite_field_instruction(path, instruction),
    }
}
