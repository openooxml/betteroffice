use std::sync::atomic::Ordering;

use serde::{Deserialize, Serialize};
use sha2::{Digest, Sha256};
use yrs::{ReadTxn, Transact};

use crate::{DeckSession, UndoCaptureMode, decode_update_v1, doc_with_client_id, hydrate_doc};

#[derive(Debug, thiserror::Error, Serialize)]
#[serde(rename_all = "camelCase")]
#[error("{code}: {message}")]
#[doc(hidden)]
pub struct PeerError {
    pub code: String,
    pub message: String,
}

impl PeerError {
    pub(crate) fn new(code: &str, message: impl ToString) -> Self {
        Self {
            code: code.into(),
            message: message.to_string(),
        }
    }
}

impl From<crate::EditError> for PeerError {
    fn from(error: crate::EditError) -> Self {
        Self::new("engine", error)
    }
}

impl From<crate::ProposalError> for PeerError {
    fn from(error: crate::ProposalError) -> Self {
        Self::new("proposal", error)
    }
}

impl From<serde_json::Error> for PeerError {
    fn from(error: serde_json::Error) -> Self {
        Self::new("invalidJson", error)
    }
}

#[derive(Clone, Debug, Deserialize, Eq, PartialEq, Serialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
#[doc(hidden)]
pub struct PeerFont {
    pub family: String,
    pub bold: bool,
    pub italic: bool,
    pub fallback: bool,
    pub fingerprint: String,
}

#[derive(Clone, Debug, Deserialize, Eq, PartialEq, Serialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub(crate) struct PeerIdentity {
    schema_version: u32,
    client_id: String,
    source_fingerprint: String,
    baseline_fingerprint: String,
    version_nonce: String,
    epoch: String,
    allocator_counter: String,
    proposal_counter: String,
    capture_policy: String,
    fonts: Vec<PeerFont>,
}

pub(crate) struct PeerOpening {
    identity: PeerIdentity,
    fonts_registered: bool,
}

fn counter(value: &str) -> Result<u64, PeerError> {
    let parsed = value
        .parse::<u64>()
        .map_err(|_| PeerError::new("identity", "invalid u64 string"))?;
    if parsed.to_string() != value {
        return Err(PeerError::new("identity", "noncanonical u64 string"));
    }
    Ok(parsed)
}

fn fingerprint(session: &DeckSession) -> Result<String, PeerError> {
    Ok(format!(
        "{:x}",
        Sha256::digest(serde_json::to_vec(&session.snapshot()?)?)
    ))
}

impl PeerIdentity {
    pub(crate) fn parse(json: &str) -> Result<Self, PeerError> {
        let identity: Self = serde_json::from_str(json)?;
        if identity.schema_version != 1 || identity.capture_policy != "manual" {
            return Err(PeerError::new(
                "schema",
                "unsupported peer identity or capture policy",
            ));
        }
        crate::validate_client_id(counter(&identity.client_id)?)?;
        for value in [
            &identity.version_nonce,
            &identity.epoch,
            &identity.allocator_counter,
            &identity.proposal_counter,
        ] {
            counter(value)?;
        }
        for hash in std::iter::once(&identity.source_fingerprint)
            .chain(std::iter::once(&identity.baseline_fingerprint))
            .chain(identity.fonts.iter().map(|font| &font.fingerprint))
        {
            if hash.len() != 64
                || !hash
                    .bytes()
                    .all(|byte| byte.is_ascii_hexdigit() && !byte.is_ascii_uppercase())
            {
                return Err(PeerError::new("identity", "invalid fingerprint"));
            }
        }
        Ok(identity)
    }
}

impl DeckSession {
    #[doc(hidden)]
    pub fn open_replay_baseline(
        source: &[u8],
        client_id: Option<u64>,
        initial_update: Option<&[u8]>,
    ) -> Result<Self, PeerError> {
        let session = if let Some(update) = initial_update {
            if update.len() > crate::MAX_UPDATE_BYTES {
                return Err(PeerError::new(
                    "initialUpdate",
                    "update exceeds the engine limit",
                ));
            }
            decode_update_v1(update).map_err(|error| PeerError::new("initialUpdate", error))?;
            let incoming = doc_with_client_id(crate::BOOTSTRAP_CLIENT_ID);
            hydrate_doc(&incoming, update)?;
            let vector = incoming.transact().state_vector();
            let client_id = match client_id {
                Some(id) if vector.get(&yrs::ClientID::new(id)) != 0 => {
                    return Err(PeerError::new(
                        "clientIdReuse",
                        "initial update already uses this client id",
                    ));
                }
                Some(id) => id,
                None => (1..=crate::MAX_SAFE_CLIENT_ID)
                    .find(|id| vector.get(&yrs::ClientID::new(*id)) == 0)
                    .ok_or_else(|| PeerError::new("clientIdReuse", "no unused client id"))?,
            };
            Self::open_from_update_with_source(update, source, client_id)
                .map_err(|error| PeerError::new("initialUpdateSource", error))?
        } else {
            Self::open(source, client_id.unwrap_or(1))?
        };
        if !session.package().has_parts() || session.has_pending_updates() {
            return Err(PeerError::new(
                "baseline",
                "a complete save-capable baseline is required",
            ));
        }
        session.set_undo_capture_mode(UndoCaptureMode::Manual);
        Ok(session)
    }

    fn verify_pristine(&self) -> Result<(), PeerError> {
        if self.epoch() != self.baseline_epoch
            || self.id_counter.load(Ordering::Relaxed) != 0
            || self.undo.borrow().diagnostics() != (0, 0)
            || self.proposals.borrow().counters().1 != 0
            || self.has_pending_updates()
            || !self.package().has_parts()
        {
            return Err(PeerError::new(
                "baseline",
                "peer hydration requires an unedited save-capable baseline without history or proposals",
            ));
        }
        Ok(())
    }

    #[doc(hidden)]
    pub fn peer_identity(&self, fonts: Vec<PeerFont>) -> Result<String, PeerError> {
        self.verify_pristine()?;
        self.undo.borrow().assert_replay_capture()?;
        let identity = PeerIdentity {
            schema_version: 1,
            client_id: self.client_id.to_string(),
            source_fingerprint: crate::deck::fingerprint_from_doc(&self.doc)?,
            baseline_fingerprint: fingerprint(self)?,
            version_nonce: self.version_nonce.load(Ordering::Relaxed).to_string(),
            epoch: self.epoch().to_string(),
            allocator_counter: self.id_counter.load(Ordering::Relaxed).to_string(),
            proposal_counter: self.proposals.borrow().counters().0.to_string(),
            capture_policy: "manual".into(),
            fonts,
        };
        let json = serde_json::to_string(&identity)?;
        PeerIdentity::parse(&json)?;
        let mut replay = self.replay_state.borrow_mut();
        if !replay.can_capture() {
            return Err(PeerError::new(
                "stage",
                "peer is already opening or replaying",
            ));
        }
        replay.ready = true;
        Ok(json)
    }

    #[doc(hidden)]
    pub fn open_peer_deck(
        source: &[u8],
        identity: &str,
        initial_update: Option<&[u8]>,
    ) -> Result<Self, PeerError> {
        let identity = PeerIdentity::parse(identity)?;
        if format!("{:x}", Sha256::digest(source)) != identity.source_fingerprint {
            return Err(PeerError::new(
                "sourceMismatch",
                "source fingerprint differs",
            ));
        }
        let session = Self::open_replay_baseline(
            source,
            Some(counter(&identity.client_id)?),
            initial_update,
        )?;
        session.replay_state.borrow_mut().opening = Some(PeerOpening {
            identity,
            fonts_registered: false,
        });
        Ok(session)
    }

    #[doc(hidden)]
    pub fn register_peer_fonts(&self, fonts: Vec<PeerFont>) -> Result<(), PeerError> {
        let mut state = self.replay_state.borrow_mut();
        let opening = state
            .opening
            .as_mut()
            .ok_or_else(|| PeerError::new("stage", "deck open must precede font registration"))?;
        if opening.fonts_registered || opening.identity.fonts != fonts {
            return Err(PeerError::new(
                "fontOrder",
                "font and fallback inputs differ or were already registered",
            ));
        }
        opening.fonts_registered = true;
        Ok(())
    }

    #[doc(hidden)]
    pub fn adopt_peer_identity(&self) -> Result<(), PeerError> {
        self.verify_pristine()?;
        self.undo.borrow().assert_replay_capture()?;
        let mut state = self.replay_state.borrow_mut();
        let opening = state
            .opening
            .as_ref()
            .ok_or_else(|| PeerError::new("stage", "no staged deck open"))?;
        let identity = &opening.identity;
        if !opening.fonts_registered {
            return Err(PeerError::new(
                "stage",
                "font registration must precede adoption",
            ));
        }
        if identity.baseline_fingerprint != fingerprint(self)?
            || counter(&identity.epoch)? != self.epoch()
            || counter(&identity.allocator_counter)? != self.id_counter.load(Ordering::Relaxed)
        {
            return Err(PeerError::new(
                "baselineMismatch",
                "reconstructed baseline differs",
            ));
        }
        let nonce = counter(&identity.version_nonce)?;
        let proposals = counter(&identity.proposal_counter)?;
        self.version_nonce.store(nonce, Ordering::Relaxed);
        self.proposals.borrow_mut().adopt_counter(proposals);
        state.opening = None;
        state.ready = true;
        Ok(())
    }
}

#[cfg(test)]
pub(crate) mod tests {
    use super::*;
    use crate::{EditCtx, TextStyle};
    use serde_json::{Value, json};

    pub(crate) fn source() -> Vec<u8> {
        let parts = [
            (
                "[Content_Types].xml",
                r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/><Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/></Types>"#,
            ),
            (
                "_rels/.rels",
                r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/></Relationships>"#,
            ),
            (
                "ppt/presentation.xml",
                r#"<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:sldIdLst><p:sldId id="256" r:id="rId1"/></p:sldIdLst><p:sldSz cx="9144000" cy="6858000"/><p:notesSz cx="6858000" cy="9144000"/></p:presentation>"#,
            ),
            (
                "ppt/_rels/presentation.xml.rels",
                r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/></Relationships>"#,
            ),
            (
                "ppt/slides/slide1.xml",
                r#"<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/><p:sp><p:nvSpPr><p:cNvPr id="2" name="Text"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1000000" cy="1000000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Hello</a:t></a:r><a:endParaRPr/></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:sld>"#,
            ),
        ];
        ooxml_opc::rezip_parts(
            &parts
                .into_iter()
                .map(|(path, xml)| (path.into(), xml.as_bytes().to_vec()))
                .collect::<Vec<_>>(),
        )
        .unwrap()
    }

    pub(crate) fn pair() -> (DeckSession, DeckSession) {
        let source = source();
        let worker = DeckSession::open_replay_baseline(&source, Some(71), None).unwrap();
        let identity = worker.peer_identity(Vec::new()).unwrap();
        let peer = DeckSession::open_peer_deck(&source, &identity, None).unwrap();
        peer.register_peer_fonts(Vec::new()).unwrap();
        peer.adopt_peer_identity().unwrap();
        (worker, peer)
    }

    pub(crate) fn replay(session: &DeckSession, op: &Value, expected: Option<Value>) -> Value {
        let sequence = session.replay_state.borrow().sequence + 1;
        let json = session
            .replay_json(
                &json!({
                    "sequence": sequence, "baseVersion": session.version(), "op": op,
                    "expectedOutcome": expected,
                })
                .to_string(),
            )
            .unwrap_or_else(|error| panic!("replay prefix {sequence}, operation {op}: {error}"));
        serde_json::from_str(&json).unwrap()
    }

    fn javascript_round_trip(value: &Value) -> Value {
        match value {
            Value::Number(number) => {
                let float = number.as_f64().unwrap() + 0.0;
                if float.fract() == 0.0 && float.abs() < 1e21 {
                    serde_json::from_str(&format!("{float:.0}")).unwrap()
                } else {
                    json!(float)
                }
            }
            Value::Array(items) => items.iter().map(javascript_round_trip).collect(),
            Value::Object(object) => object
                .iter()
                .map(|(key, value)| (key.clone(), javascript_round_trip(value)))
                .collect(),
            other => other.clone(),
        }
    }

    pub(crate) fn apply_pair(worker: &DeckSession, peer: &DeckSession, op: &Value) -> Value {
        let result = replay(worker, op, None);
        assert_eq!(
            result,
            replay(peer, op, Some(javascript_round_trip(&result["outcome"]))),
            "operation {op}"
        );
        result
    }

    #[test]
    fn identity_u64_values_are_lossless_strings() {
        let bytes = source();
        let worker = DeckSession::open_replay_baseline(&bytes, Some(71), None).unwrap();
        worker.version_nonce.store(u64::MAX, Ordering::Relaxed);
        worker
            .proposals
            .borrow_mut()
            .adopt_counter((1_u64 << 53) + 7);
        let identity = worker.peer_identity(Vec::new()).unwrap();
        let payload: Value = serde_json::from_str(&identity).unwrap();
        assert_eq!(payload["versionNonce"], u64::MAX.to_string());
        assert_eq!(payload["proposalCounter"], ((1_u64 << 53) + 7).to_string());
        let peer = DeckSession::open_peer_deck(&bytes, &identity, None).unwrap();
        peer.register_peer_fonts(Vec::new()).unwrap();
        peer.adopt_peer_identity().unwrap();
        assert_eq!(worker.version(), peer.version());
        assert_eq!(
            worker.proposals.borrow().counters(),
            peer.proposals.borrow().counters()
        );
    }

    #[test]
    fn effective_client_id_and_generated_ids_match() {
        let (worker, peer) = pair();
        assert_eq!(worker.client_id(), peer.client_id());
        let result = apply_pair(&worker, &peer, &json!({"method":"insertSlide", "args":[1]}));
        assert_eq!(result["outcome"]["result"]["slideId"], "slide:71:0");
        assert_eq!(
            worker.id_counter.load(Ordering::Relaxed),
            peer.id_counter.load(Ordering::Relaxed)
        );
    }

    #[test]
    fn peer_versions_match_independent_versions_differ() {
        let (worker, peer) = pair();
        assert_eq!(worker.version(), peer.version());
        let independent = DeckSession::open(&source(), 71).unwrap();
        assert_ne!(worker.version(), independent.version());
        assert_ne!(
            independent.version(),
            DeckSession::open(&source(), 71).unwrap().version()
        );
    }

    #[test]
    fn initial_update_requires_matching_source_and_fresh_client() {
        let bytes = source();
        let original = DeckSession::open(&bytes, 71).unwrap();
        let slide = original.slide_ids().unwrap()[0].clone();
        original
            .set_slide_notes(&EditCtx::local("test"), &slide, "updated")
            .unwrap();
        let update = original.encode_state_as_update_v1();
        let reused = DeckSession::open_replay_baseline(&bytes, Some(71), Some(&update));
        assert!(matches!(reused, Err(error) if error.code == "clientIdReuse"));
        let unrelated =
            ooxml_opc::rezip_parts(&[("ppt/presentation.xml".into(), b"<p/>".to_vec())]).unwrap();
        assert!(
            matches!(DeckSession::open_replay_baseline(&unrelated, Some(72), Some(&update)), Err(error) if error.code == "initialUpdateSource")
        );
        let opened = DeckSession::open_replay_baseline(&bytes, None, Some(&update)).unwrap();
        assert_ne!(opened.client_id(), 71);
        assert_eq!(opened.snapshot().unwrap(), original.snapshot().unwrap());
        assert_eq!(opened.save().unwrap(), original.save().unwrap());
        let identity = opened.peer_identity(Vec::new()).unwrap();
        let peer = DeckSession::open_peer_deck(&bytes, &identity, Some(&update)).unwrap();
        peer.register_peer_fonts(Vec::new()).unwrap();
        peer.adopt_peer_identity().unwrap();
        assert_eq!(opened.version(), peer.version());
        assert_eq!(opened.snapshot().unwrap(), peer.snapshot().unwrap());
        assert_eq!(opened.save().unwrap(), peer.save().unwrap());
    }

    #[test]
    fn staged_hydration_refuses_without_partial_state() {
        let bytes = source();
        let worker = DeckSession::open_replay_baseline(&bytes, Some(71), None).unwrap();
        let identity = worker.peer_identity(Vec::new()).unwrap();
        let mut mismatched: Value = serde_json::from_str(&identity).unwrap();
        mismatched["epoch"] = json!("1");
        mismatched["proposalCounter"] = json!("99");
        let mismatch = DeckSession::open_peer_deck(&bytes, &mismatched.to_string(), None).unwrap();
        mismatch.register_peer_fonts(Vec::new()).unwrap();
        let mismatch_version = mismatch.version();
        assert!(
            matches!(mismatch.adopt_peer_identity(), Err(error) if error.code == "baselineMismatch")
        );
        assert_eq!(mismatch.version(), mismatch_version);
        assert_eq!(mismatch.proposals.borrow().counters(), (0, 0));
        assert!(!mismatch.replay_state.borrow().ready);
        let peer = DeckSession::open_peer_deck(&bytes, &identity, None).unwrap();
        let version = peer.version();
        assert!(peer.adopt_peer_identity().is_err());
        assert_eq!(peer.version(), version);
        assert!(!peer.replay_state.borrow().ready);
        let mut bad: Value = serde_json::from_str(&identity).unwrap();
        bad["schemaVersion"] = json!(2);
        assert!(DeckSession::open_peer_deck(&bytes, &bad.to_string(), None).is_err());
        peer.register_peer_fonts(Vec::new()).unwrap();
        assert!(peer.register_peer_fonts(Vec::new()).is_err());
        let story = peer.snapshot().unwrap().slides[0].shapes[0].text_stories[0]
            .id
            .clone();
        peer.insert_text(
            &EditCtx::local("test"),
            &story,
            0,
            "edited",
            &TextStyle::default(),
        )
        .unwrap();
        let version = peer.version();
        assert!(peer.adopt_peer_identity().is_err());
        assert_eq!(peer.version(), version);
        assert!(!peer.replay_state.borrow().ready);
        assert_ne!(peer.version(), worker.version());
    }

    #[test]
    fn font_and_fallback_order_match() {
        let bytes = source();
        let worker = DeckSession::open_replay_baseline(&bytes, Some(71), None).unwrap();
        let fonts = vec![
            PeerFont {
                family: "A".into(),
                bold: false,
                italic: true,
                fallback: false,
                fingerprint: format!("{:x}", Sha256::digest(b"a")),
            },
            PeerFont {
                family: "B".into(),
                bold: true,
                italic: false,
                fallback: true,
                fingerprint: format!("{:x}", Sha256::digest(b"b")),
            },
        ];
        let identity = worker.peer_identity(fonts.clone()).unwrap();
        let peer = DeckSession::open_peer_deck(&bytes, &identity, None).unwrap();
        let version = peer.version();
        let mut reversed = fonts.clone();
        reversed.reverse();
        assert!(peer.register_peer_fonts(reversed).is_err());
        assert_eq!(peer.version(), version);
        assert!(peer.adopt_peer_identity().is_err());
        let mut changed = fonts.clone();
        changed[0].fallback = true;
        assert!(peer.register_peer_fonts(changed).is_err());
        peer.register_peer_fonts(fonts).unwrap();
        peer.adopt_peer_identity().unwrap();
        assert_eq!(worker.version(), peer.version());
    }
}
