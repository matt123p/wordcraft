use thiserror::Error;
use wordcraft_doc::{Block, Document, Pos, StoryRef};

/// Maximum main-story size indexed by the initial Word Range profile.
pub const MAX_MAPPED_CHARACTERS: usize = 4 * 1024 * 1024;

/// Which end of a range should receive a character position when the final paragraph mark makes
/// one model position correspond to two Word positions.
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum PositionAffinity {
    Start,
    End,
}

#[derive(Clone, Debug, Error, Eq, PartialEq)]
pub enum PositionError {
    #[error("the main-story range profile is limited to plain paragraphs")]
    UnsupportedStructure,
    #[error("document exceeds the Word Range character budget")]
    Limit,
    #[error("document position is outside the main story")]
    OutsideStory,
    #[error("document position falls inside a UTF-16 surrogate pair")]
    SplitSurrogate,
    #[error("position mapping allocation failed")]
    Allocation,
}

/// Snapshot mapping between Word's zero-based main-story character positions and WordCraft's
/// UTF-8 byte-offset positions. A paragraph break and the final paragraph mark each occupy one
/// Word position. This profile deliberately refuses tables and inline objects rather than
/// assigning them guessed character widths.
#[derive(Clone, Debug, Eq, PartialEq)]
pub struct MainStoryMap {
    boundaries: Vec<Option<Pos>>,
}

impl MainStoryMap {
    pub fn new(document: &Document) -> Result<Self, PositionError> {
        let mut boundaries = Vec::new();
        boundaries.try_reserve(document.body.len().saturating_add(1)).map_err(|_| PositionError::Allocation)?;

        for (block_index, block) in document.body.iter().enumerate() {
            let Block::Para(paragraph) = &**block else {
                return Err(PositionError::UnsupportedStructure);
            };
            if !paragraph.objects.is_empty() {
                return Err(PositionError::UnsupportedStructure);
            }
            if block_index == 0 {
                boundaries.push(Some(Pos::body(block_index, 0)));
            }

            for (byte, character) in paragraph.text.char_indices() {
                let units = character.len_utf16();
                if boundaries.len().saturating_add(units) > MAX_MAPPED_CHARACTERS.saturating_add(1) {
                    return Err(PositionError::Limit);
                }
                boundaries.try_reserve(units).map_err(|_| PositionError::Allocation)?;
                if units == 1 {
                    boundaries.push(Some(Pos::body(block_index, byte + character.len_utf8())));
                } else {
                    // A UTF-16 index between a surrogate pair has no UTF-8 byte position.
                    boundaries.push(None);
                    boundaries.push(Some(Pos::body(block_index, byte + character.len_utf8())));
                }
            }

            if block_index + 1 < document.body.len() {
                if boundaries.len() >= MAX_MAPPED_CHARACTERS {
                    return Err(PositionError::Limit);
                }
                let next = document.body.get(block_index + 1).ok_or(PositionError::UnsupportedStructure)?;
                let Block::Para(_) = &**next else {
                    return Err(PositionError::UnsupportedStructure);
                };
                // The current end boundary denotes the position before CR. The next boundary
                // denotes the position after CR.
                if boundaries.last() != Some(&Some(Pos::body(block_index, paragraph.text.len()))) {
                    return Err(PositionError::UnsupportedStructure);
                }
                boundaries.push(Some(Pos::body(block_index + 1, 0)));
            }
        }

        if document.body.is_empty() {
            boundaries.push(Some(Pos::body(0, 0)));
        }
        // Word's document content ends after its final paragraph mark. The model position before
        // and after that mark is the same; affinity disambiguates the two character positions.
        let terminal = boundaries.last().cloned().unwrap_or(Some(Pos::body(0, 0)));
        boundaries.try_reserve(1).map_err(|_| PositionError::Allocation)?;
        boundaries.push(terminal);
        if boundaries.len() > MAX_MAPPED_CHARACTERS.saturating_add(2) {
            return Err(PositionError::Limit);
        }
        Ok(Self { boundaries })
    }

    /// Number of addressable Word character positions, including paragraph marks.
    pub fn len(&self) -> usize {
        self.boundaries.len().saturating_sub(1)
    }

    pub fn is_empty(&self) -> bool {
        self.len() == 0
    }

    /// Convert a zero-based Word position into a WordCraft model position.
    pub fn position(&self, index: usize) -> Result<Pos, PositionError> {
        if index > self.len() {
            return Err(PositionError::OutsideStory);
        }
        self.boundaries.get(index).ok_or(PositionError::OutsideStory)?.clone().ok_or(PositionError::SplitSurrogate)
    }

    /// Convert a WordCraft position to a zero-based Word position. The final paragraph mark has
    /// identical model positions before and after it, so callers choose start/end affinity.
    pub fn index(&self, position: &Pos, affinity: PositionAffinity) -> Result<usize, PositionError> {
        if position.story != StoryRef::Body {
            return Err(PositionError::OutsideStory);
        }
        let matches = self.boundaries.iter().enumerate().filter_map(|(index, candidate)| (candidate.as_ref() == Some(position)).then_some(index));
        match affinity {
            PositionAffinity::Start => matches.into_iter().next(),
            PositionAffinity::End => matches.into_iter().next_back(),
        }
        .ok_or(PositionError::OutsideStory)
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::sync::Arc;
    use wordcraft_doc::para::InlineObject;

    #[test]
    fn maps_unicode_scalar_bytes_to_utf16_story_positions() {
        let document = Document::from_text("A😀e\u{301}\nZ");
        let map = MainStoryMap::new(&document).expect("plain paragraphs map");

        // A (1) + the supplementary scalar (2) + e + combining mark + paragraph mark + Z + final mark.
        assert_eq!(map.len(), 8);
        assert_eq!(map.position(0), Ok(Pos::body(0, 0)));
        assert_eq!(map.position(1), Ok(Pos::body(0, 1)));
        assert_eq!(map.position(2), Err(PositionError::SplitSurrogate));
        assert_eq!(map.position(3), Ok(Pos::body(0, 5)));
        assert_eq!(map.position(6), Ok(Pos::body(1, 0)));
        assert_eq!(map.index(&Pos::body(1, 0), PositionAffinity::Start), Ok(6));
        assert_eq!(map.index(&Pos::body(1, 1), PositionAffinity::Start), Ok(7));
        assert_eq!(map.index(&Pos::body(1, 1), PositionAffinity::End), Ok(8));
    }

    #[test]
    fn refuses_tables_and_out_of_range_positions() {
        let document = Document::from_text("a\nb");
        let map = MainStoryMap::new(&document).expect("plain paragraphs map");
        assert_eq!(map.position(map.len() + 1), Err(PositionError::OutsideStory));

        let mut with_table = Document::new();
        with_table.body.push(Arc::new(Block::Table(Default::default())));
        assert_eq!(MainStoryMap::new(&with_table), Err(PositionError::UnsupportedStructure));
    }

    #[test]
    fn refuses_inline_objects_instead_of_guessing_their_width() {
        let mut document = Document::from_text("text");
        if let Some(Block::Para(paragraph)) = document.body.get_mut(0).map(Arc::make_mut) {
            paragraph.objects.push(InlineObject::note_separator(false));
        }
        assert_eq!(MainStoryMap::new(&document), Err(PositionError::UnsupportedStructure));
    }
}
