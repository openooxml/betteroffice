use serde::{Deserialize, Serialize};

use crate::{
    DeckSession, DeckSnapshot, Proposal, ProposalResult, ShapeSnapshot, StorySnapshot,
    TextRunSnapshot, TextStyle,
};

#[derive(Clone, Copy, Debug, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum ProposalTextChangeKind {
    Insertion,
    Deletion,
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ProposalTextChange {
    pub story_id: String,
    pub start: u32,
    pub end: u32,
    pub kind: ProposalTextChangeKind,
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ProposalDiffPreview {
    pub proposal: Proposal,
    pub snapshot: DeckSnapshot,
    pub text_changes: Vec<ProposalTextChange>,
}

impl DeckSession {
    /// Builds a render-only snapshot whose offsets are not editable.
    pub fn preview_proposal_diff(&self, id: &str) -> ProposalResult<ProposalDiffPreview> {
        let preview = self.preview_proposal(id)?;
        let mut snapshot = preview.snapshot;
        let mut text_changes = Vec::new();
        for change in &preview.proposal.changes {
            let Some(before) = &change.before else {
                continue;
            };
            let Some(slide) = snapshot
                .slides
                .iter_mut()
                .find(|slide| slide.id == change.slide_id)
            else {
                continue;
            };
            let Some(after) = find_shape_mut(&mut slide.shapes, &before.id) else {
                continue;
            };
            for story in &mut after.text_stories {
                if let Some(original) = before.text_stories.iter().find(|old| old.id == story.id) {
                    diff_story(original, story, &mut text_changes);
                }
            }
        }
        Ok(ProposalDiffPreview {
            proposal: preview.proposal,
            snapshot,
            text_changes,
        })
    }
}

fn find_shape_mut<'a>(shapes: &'a mut [ShapeSnapshot], id: &str) -> Option<&'a mut ShapeSnapshot> {
    for shape in shapes {
        if shape.id == id {
            return Some(shape);
        }
        if let Some(child) = find_shape_mut(&mut shape.children, id) {
            return Some(child);
        }
    }
    None
}

#[derive(Clone, Debug, PartialEq)]
struct Token {
    text: String,
    style: TextStyle,
}

fn tokens(runs: &[TextRunSnapshot]) -> Vec<Token> {
    let mut result: Vec<Token> = Vec::new();
    for run in runs {
        for ch in run.text.chars() {
            if let Some(last) = result.last_mut()
                && last.style == run.style
                && last.text.ends_with(char::is_whitespace) == ch.is_whitespace()
            {
                last.text.push(ch);
            } else {
                result.push(Token {
                    text: ch.to_string(),
                    style: run.style.clone(),
                });
            }
        }
    }
    result
}

fn diff_story(
    before: &StorySnapshot,
    after: &mut StorySnapshot,
    changes: &mut Vec<ProposalTextChange>,
) {
    let mut offset = 0;
    for paragraph in &mut after.paragraphs {
        let old = before.paragraphs.iter().find(|old| old.id == paragraph.id);
        let old = tokens(old.map_or(&[], |paragraph| &paragraph.runs));
        let new = tokens(&paragraph.runs);
        let mut runs: Vec<TextRunSnapshot> = Vec::new();
        for (token, kind) in diff_tokens(&old, &new) {
            let start = offset;
            offset += token.text.encode_utf16().count() as u32;
            let mut style = token.style.clone();
            if let Some(kind) = kind {
                style.color = Some(
                    match kind {
                        ProposalTextChangeKind::Insertion => "#166534",
                        ProposalTextChangeKind::Deletion => "#b91c1c",
                    }
                    .into(),
                );
                if let Some(last) = changes.last_mut()
                    && last.story_id == after.id
                    && last.kind == kind
                    && last.end == start
                {
                    last.end = offset;
                } else {
                    changes.push(ProposalTextChange {
                        story_id: after.id.clone(),
                        start,
                        end: offset,
                        kind,
                    });
                }
            }
            if let Some(last) = runs.last_mut()
                && last.style == style
            {
                last.text.push_str(&token.text);
            } else {
                runs.push(TextRunSnapshot {
                    text: token.text.clone(),
                    style,
                });
            }
        }
        paragraph.runs = runs;
        offset += 1;
    }
    after.length = offset;
}

fn diff_tokens<'a>(
    old: &'a [Token],
    new: &'a [Token],
) -> Vec<(&'a Token, Option<ProposalTextChangeKind>)> {
    use ProposalTextChangeKind::{Deletion, Insertion};
    let prefix = old.iter().zip(new).take_while(|(a, b)| a == b).count();
    let suffix = old[prefix..]
        .iter()
        .rev()
        .zip(new[prefix..].iter().rev())
        .take_while(|(a, b)| a == b)
        .count();
    let mut result: Vec<_> = new[..prefix].iter().map(|token| (token, None)).collect();
    let a = &old[prefix..old.len() - suffix];
    let b = &new[prefix..new.len() - suffix];
    if a.len()
        .saturating_add(1)
        .saturating_mul(b.len().saturating_add(1))
        > 250_000
    {
        result.extend(a.iter().map(|token| (token, Some(Deletion))));
        result.extend(b.iter().map(|token| (token, Some(Insertion))));
    } else {
        let width = b.len() + 1;
        let mut lengths = vec![0u32; (a.len() + 1) * width];
        for i in (0..a.len()).rev() {
            for j in (0..b.len()).rev() {
                lengths[i * width + j] = if a[i] == b[j] {
                    lengths[(i + 1) * width + j + 1] + 1
                } else {
                    lengths[(i + 1) * width + j].max(lengths[i * width + j + 1])
                };
            }
        }
        let (mut i, mut j) = (0, 0);
        while i < a.len() || j < b.len() {
            if i < a.len() && j < b.len() && a[i] == b[j] {
                result.push((&b[j], None));
                i += 1;
                j += 1;
            } else if i < a.len()
                && (j == b.len() || lengths[(i + 1) * width + j] >= lengths[i * width + j + 1])
            {
                result.push((&a[i], Some(Deletion)));
                i += 1;
            } else {
                result.push((&b[j], Some(Insertion)));
                j += 1;
            }
        }
    }
    result.extend(new[new.len() - suffix..].iter().map(|token| (token, None)));
    result
}
