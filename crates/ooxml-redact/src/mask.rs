use unicode_script::{Script, UnicodeScript};

use crate::{RedactError, RedactionOptions};

const KANJI: &str = "日月火水木金土山川田人大小中上下左右前後東西南北本年時分毎今先生学校子女子男女父母友白赤青黒春夏秋冬朝昼夜空雨雪風花草竹林森海池魚鳥犬猫馬牛車電気天文音楽光家店町村国道駅門外内高長新古多少早明広近遠強弱心力手足目耳口首体食飲読書話聞見行来帰入出休立歩走買売思知作使持待会合開閉動止色形円角紙糸米麦茶肉石玉王正直百千万一二三四五六七八九十";

pub(crate) struct TextMasker {
    random: Option<RandomCharacters>,
}

impl TextMasker {
    pub(crate) fn new(options: &RedactionOptions) -> Self {
        Self {
            random: options.random_characters.then(RandomCharacters::new),
        }
    }

    pub(crate) fn is_random(&self) -> bool {
        self.random.is_some()
    }

    pub(crate) fn replace(&mut self, text: &str) -> Result<String, RedactError> {
        let Some(random) = &mut self.random else {
            return Ok(placeholder(text));
        };
        let mut counts = [0usize; 4];
        for character in text.chars() {
            if let Some(index) = script_index(character) {
                counts[index] += 1;
            }
        }
        let dominant = counts
            .iter()
            .enumerate()
            .max_by_key(|(index, count)| (**count, std::cmp::Reverse(*index)))
            .map_or(0, |(index, _)| index);
        text.chars()
            .map(|character| {
                if character.is_whitespace() {
                    return Ok(character);
                }
                let script = script_index(character).unwrap_or(dominant);
                random.character(script, character.is_uppercase())
            })
            .collect()
    }
}

fn script_index(character: char) -> Option<usize> {
    match character.script() {
        Script::Latin => Some(0),
        Script::Hiragana => Some(1),
        Script::Katakana => Some(2),
        Script::Han => Some(3),
        _ => None,
    }
}

struct RandomCharacters {
    bytes: [u8; 4096],
    offset: usize,
    kanji: Vec<char>,
    fill: fn(&mut [u8]) -> Result<(), RedactError>,
}

impl RandomCharacters {
    fn new() -> Self {
        Self {
            bytes: [0; 4096],
            offset: 4096,
            kanji: KANJI.chars().collect(),
            fill: |bytes| {
                getrandom::fill(bytes).map_err(|error| RedactError::Randomness(error.to_string()))
            },
        }
    }

    fn index(&mut self, count: u32) -> Result<usize, RedactError> {
        let limit = u32::MAX - u32::MAX % count;
        loop {
            if self.offset == self.bytes.len() {
                (self.fill)(&mut self.bytes)?;
                self.offset = 0;
            }
            let value =
                u32::from_ne_bytes(self.bytes[self.offset..self.offset + 4].try_into().unwrap());
            self.offset += 4;
            if value < limit {
                return Ok((value % count) as usize);
            }
        }
    }

    fn character(&mut self, script: usize, uppercase: bool) -> Result<char, RedactError> {
        if script == 3 {
            let index = self.index(self.kanji.len() as u32)?;
            return Ok(self.kanji[index]);
        }
        let (first, count) = match script {
            1 => (0x3041, 86),
            2 => (0x30a1, 90),
            _ => (
                if uppercase {
                    u32::from('A')
                } else {
                    u32::from('a')
                },
                26,
            ),
        };
        Ok(char::from_u32(first + self.index(count)? as u32).unwrap())
    }
}

pub(crate) fn placeholder(text: &str) -> String {
    text.chars()
        .map(|character| {
            if character.is_whitespace() {
                character
            } else {
                'x'
            }
        })
        .collect()
}

#[cfg(test)]
mod tests {
    use super::*;

    fn deterministic() -> TextMasker {
        let mut masker = TextMasker::new(&RedactionOptions {
            random_characters: true,
        });
        masker.random.as_mut().unwrap().fill = |bytes| {
            for (index, chunk) in bytes.as_chunks_mut::<4>().0.iter_mut().enumerate() {
                chunk.copy_from_slice(&(index as u32).to_ne_bytes());
            }
            Ok(())
        };
        masker
    }

    #[test]
    fn replacements_preserve_scripts_and_whitespace_without_a_substitution_map() {
        let mut masker = deterministic();
        let source = "ABC éßø ひらがな カタカナ 日本語 \t\r\n\u{3000}";
        let output = masker.replace(source).unwrap();
        for (before, after) in source.chars().zip(output.chars()) {
            if before.is_whitespace() {
                assert_eq!(after, before);
            } else {
                assert_eq!(after.script(), before.script());
                assert!(after.is_alphabetic());
            }
        }
        let first = masker.replace(&"a".repeat(64)).unwrap();
        let second = masker.replace(&"a".repeat(64)).unwrap();
        assert!(
            first
                .chars()
                .collect::<std::collections::HashSet<_>>()
                .len()
                > 20
        );
        assert_ne!(first, second);
        let japanese = masker.replace("日本語！123🙂").unwrap();
        assert_eq!(japanese.chars().count(), 8);
        assert!(
            japanese
                .chars()
                .all(|character| character.script() == Script::Han)
        );
    }

    #[test]
    fn replacement_alphabets_contain_only_assigned_letters() {
        for (first, last, script) in [
            (0x3041, 0x3096, Script::Hiragana),
            (0x30a1, 0x30fa, Script::Katakana),
        ] {
            for code in first..=last {
                let character = char::from_u32(code).unwrap();
                assert!(character.is_alphabetic());
                assert_eq!(character.script(), script);
            }
        }
        assert!(
            KANJI
                .chars()
                .all(|character| character.is_alphabetic() && character.script() == Script::Han)
        );
    }

    #[test]
    fn randomness_failure_is_reported_without_a_fallback() {
        let mut masker = deterministic();
        masker.random.as_mut().unwrap().fill =
            |_| Err(RedactError::Randomness("unavailable".to_owned()));
        assert!(matches!(
            masker.replace("private"),
            Err(RedactError::Randomness(_))
        ));
        assert_eq!(masker.replace(" \t\n").unwrap(), " \t\n");
    }

    #[test]
    fn default_mask_remains_exactly_compatible() {
        let source = "a日あカ1!🙂 \t\n\u{3000}";
        assert_eq!(
            TextMasker::new(&Default::default())
                .replace(source)
                .unwrap(),
            "xxxxxxx \t\n\u{3000}"
        );
    }
}
