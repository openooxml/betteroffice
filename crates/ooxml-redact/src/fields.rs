use crate::RedactError;
use crate::identifiers::instruction_edits;
use crate::mask::TextMasker;

/// Rewrites complex-field instructions in their original text slots.
#[derive(Default)]
pub(crate) struct Fields {
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

impl Fields {
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
        if let Some(edits) = instruction_edits(&field.instruction, &field.breaks, masker)? {
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
