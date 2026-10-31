use super::{FontFace, FontId, RenderError, SlideRenderer, normalize_family};

impl SlideRenderer {
    #[doc(hidden)]
    pub fn remove_fallback_face(&mut self, id: u32) {
        self.coverage_faces
            .retain(|(_, _, face)| face.id.to_u32() != id);
        for faces in self.family_coverage_faces.values_mut() {
            faces.retain(|(_, _, face)| face.id.to_u32() != id);
        }
        if let Ok(cache) = self.text_layouts.get_mut() {
            cache.clear();
        }
    }

    #[doc(hidden)]
    pub fn register_priority_fallback_face(
        &mut self,
        family: &str,
        bold: bool,
        italic: bool,
        id: u32,
        index: usize,
    ) -> Result<(), RenderError> {
        let id = FontId::from_u32(id);
        self.fonts
            .metrics(id)
            .map_err(|error| RenderError::Font(error.to_string()))?;
        let requested_family = normalize_family(family);
        let faces = self
            .family_coverage_faces
            .entry(requested_family.clone())
            .or_default();
        let current = faces.iter().position(|(face_bold, face_italic, face)| {
            (*face_bold, *face_italic, face.id) == (bold, italic, id)
        });
        if current == Some(index.min(faces.len().saturating_sub(1))) {
            return Ok(());
        }
        if let Some(current) = current {
            faces.remove(current);
        }
        faces.insert(
            index.min(faces.len()),
            (
                bold,
                italic,
                FontFace {
                    id,
                    family: family.to_owned(),
                    requested_family,
                    widths: None,
                    line: None,
                    line_id: id,
                },
            ),
        );
        if let Ok(cache) = self.text_layouts.get_mut() {
            cache.clear();
        }
        Ok(())
    }

    #[doc(hidden)]
    pub fn retain_font_style(&mut self, family: &str, bold: bool, italic: bool) {
        let family = normalize_family(family);
        self.faces.retain(|(name, face_bold, face_italic), _| {
            *name != family || (*face_bold, *face_italic) == (bold, italic)
        });
    }

    #[doc(hidden)]
    pub fn register_priority_fallback_font(
        &mut self,
        family: &str,
        bold: bool,
        italic: bool,
        bytes: &[u8],
        index: usize,
    ) -> Result<u32, RenderError> {
        let id = self.register_fallback_font(family, bold, italic, bytes)?;
        let face = self.coverage_faces.pop().expect("registered fallback face");
        self.coverage_faces
            .insert(index.min(self.coverage_faces.len()), face);
        Ok(id)
    }

    #[doc(hidden)]
    pub fn has_fallback_font(&self, family: &str, bold: bool, italic: bool) -> bool {
        let family = normalize_family(family);
        self.coverage_faces
            .iter()
            .any(|(face_bold, face_italic, face)| {
                *face_bold == bold && *face_italic == italic && face.requested_family == family
            })
    }

    #[doc(hidden)]
    pub fn has_font_face(&self, family: &str, bold: bool, italic: bool) -> bool {
        self.faces
            .contains_key(&(normalize_family(family), bold, italic))
    }
}
