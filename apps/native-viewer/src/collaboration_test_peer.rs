use std::fmt::Write as _;
use std::io::{self, BufRead};
use std::path::Path;
use std::sync::mpsc::{Receiver, Sender, TryRecvError, channel};
use std::thread;
use std::time::Duration;

use anyhow::{Context, Result, bail};
use docx_edit::story_checksum;
use serde::{Deserialize, Serialize};
use sha2::{Digest, Sha256};
use yrs::{Map, ReadTxn, Transact};

use crate::collaboration::{CollaborationClient, CollaborationConfig, TransportEvent};
use crate::collaboration_document;
use crate::document::{DocumentView, ReferenceDocument, load_collaborative_docx};
use crate::editing::TextLoc;

#[derive(Deserialize)]
#[serde(tag = "command", rename_all = "camelCase")]
enum PeerCommand {
    Snapshot {
        id: u64,
    },
    Insert {
        id: u64,
        #[serde(rename = "paraId")]
        para_id: String,
        offset: u32,
        text: String,
    },
    Disconnect {
        id: u64,
    },
    Reconnect {
        id: u64,
    },
    Shutdown {
        id: u64,
    },
}

impl PeerCommand {
    fn id(&self) -> u64 {
        match self {
            Self::Snapshot { id }
            | Self::Insert { id, .. }
            | Self::Disconnect { id }
            | Self::Reconnect { id }
            | Self::Shutdown { id } => *id,
        }
    }
}

enum PeerInput {
    Command(PeerCommand),
    Invalid(String),
    Closed,
}

#[derive(Serialize)]
#[serde(rename_all = "camelCase")]
struct PeerResponse {
    id: u64,
    ok: bool,
    snapshot: Option<PeerSnapshot>,
    error: Option<String>,
}

#[derive(Serialize)]
#[serde(rename_all = "camelCase")]
struct PeerSnapshot {
    connected: bool,
    synced: bool,
    state_vector: String,
    canonical_checksum: String,
    stories: Vec<StoryState>,
    paragraphs: Vec<ParagraphState>,
}

#[derive(Serialize)]
#[serde(rename_all = "camelCase")]
struct StoryState {
    story_id: String,
    checksum: String,
}

#[derive(Serialize)]
#[serde(rename_all = "camelCase")]
struct ParagraphState {
    para_id: String,
    text: String,
}

struct Connection {
    client: CollaborationClient,
    events: Receiver<TransportEvent>,
}

pub fn run(document_path: &Path, config: CollaborationConfig) -> Result<()> {
    let mut document = load_collaborative_docx(document_path, 8_192, config.client_id())?;
    let (inputs, input_receiver) = channel();
    read_commands(inputs)?;
    let mut connection = Some(connect(&config)?);
    let stdout = io::stdout();
    let mut output = stdout.lock();

    'running: loop {
        if let Some(active) = &mut connection {
            match active.events.recv_timeout(Duration::from_millis(10)) {
                Ok(event) => {
                    collaboration_document::handle_transport_event(
                        &mut document,
                        &mut active.client,
                        event,
                    )?;
                }
                Err(std::sync::mpsc::RecvTimeoutError::Timeout) => {}
                Err(std::sync::mpsc::RecvTimeoutError::Disconnected) => {
                    bail!("collaboration event channel stopped")
                }
            }
            while let Ok(event) = active.events.try_recv() {
                collaboration_document::handle_transport_event(
                    &mut document,
                    &mut active.client,
                    event,
                )?;
            }
            collaboration_document::forward_local_updates(&document, &mut active.client);
        } else {
            thread::sleep(Duration::from_millis(10));
        }

        loop {
            let input = match input_receiver.try_recv() {
                Ok(input) => input,
                Err(TryRecvError::Empty) => break,
                Err(TryRecvError::Disconnected) => PeerInput::Closed,
            };
            match input {
                PeerInput::Command(command) => {
                    let id = command.id();
                    let shutdown = matches!(command, PeerCommand::Shutdown { .. });
                    let result = apply_command(command, &config, &mut document, &mut connection)
                        .and_then(|()| snapshot(&document, connection.as_ref()));
                    write_response(&mut output, id, result)?;
                    if shutdown {
                        break 'running;
                    }
                }
                PeerInput::Invalid(error) => {
                    write_response(&mut output, 0, Err(anyhow::anyhow!(error)))?;
                }
                PeerInput::Closed => break 'running,
            }
        }
    }
    Ok(())
}

fn read_commands(sender: Sender<PeerInput>) -> Result<()> {
    thread::Builder::new()
        .name("collaboration-test-commands".to_owned())
        .spawn(move || {
            let stdin = io::stdin();
            for line in stdin.lock().lines() {
                let input = match line {
                    Ok(line) => match serde_json::from_str(&line) {
                        Ok(command) => PeerInput::Command(command),
                        Err(error) => PeerInput::Invalid(error.to_string()),
                    },
                    Err(error) => PeerInput::Invalid(error.to_string()),
                };
                if sender.send(input).is_err() {
                    return;
                }
            }
            let _ = sender.send(PeerInput::Closed);
        })
        .context("start collaboration test command reader")?;
    Ok(())
}

fn connect(config: &CollaborationConfig) -> Result<Connection> {
    let (sender, events) = channel();
    let client =
        CollaborationClient::start(config.clone(), move |event| sender.send(event).is_ok())?;
    Ok(Connection { client, events })
}

fn apply_command(
    command: PeerCommand,
    config: &CollaborationConfig,
    document: &mut DocumentView,
    connection: &mut Option<Connection>,
) -> Result<()> {
    match command {
        PeerCommand::Snapshot { .. } | PeerCommand::Shutdown { .. } => {}
        PeerCommand::Insert {
            para_id,
            offset,
            text,
            ..
        } => {
            document.docx_select_point(TextLoc { para_id, offset }, false, false)?;
            if !document.docx_insert_text(&text)? {
                bail!("native insert made no change");
            }
            if let Some(active) = connection {
                collaboration_document::forward_local_updates(document, &mut active.client);
            }
        }
        PeerCommand::Disconnect { .. } => {
            if connection.take().is_none() {
                bail!("native peer is already disconnected");
            }
        }
        PeerCommand::Reconnect { .. } => {
            if connection.is_some() {
                bail!("native peer is already connected");
            }
            *connection = Some(connect(config)?);
        }
    }
    Ok(())
}

fn snapshot(document: &DocumentView, connection: Option<&Connection>) -> Result<PeerSnapshot> {
    let ReferenceDocument::Docx(reference) = &document.reference else {
        bail!("collaboration test peer requires DOCX");
    };
    let editing = reference.editor.engine().doc();
    let story_ids = {
        let transaction = editing.yrs_doc().transact();
        let stories = transaction
            .get_map("stories")
            .context("collaborative document has no stories")?;
        let mut ids = stories
            .keys(&transaction)
            .map(|id| id.to_string())
            .collect::<Vec<_>>();
        ids.sort();
        ids
    };
    let stories = story_ids
        .into_iter()
        .map(|story_id| {
            Ok(StoryState {
                checksum: story_checksum(editing, &story_id)?.to_string(),
                story_id,
            })
        })
        .collect::<Result<Vec<_>>>()?;
    let paragraphs = editing
        .paragraphs("body")?
        .into_iter()
        .map(|paragraph| ParagraphState {
            para_id: paragraph.para_id,
            text: paragraph.text,
        })
        .collect();
    Ok(PeerSnapshot {
        connected: connection.is_some_and(|active| active.client.is_connected()),
        synced: connection.is_some_and(|active| active.client.is_synced()),
        state_vector: hex(&document.docx_state_vector()?),
        canonical_checksum: canonical_checksum(&stories),
        stories,
        paragraphs,
    })
}

fn canonical_checksum(stories: &[StoryState]) -> String {
    let mut checksum = Sha256::new();
    checksum.update(b"canonical-document-checksums-v1\n");
    for story in stories {
        checksum.update(story.story_id.as_bytes());
        checksum.update([0]);
        checksum.update(story.checksum.as_bytes());
        checksum.update(b"\n");
    }
    hex(&checksum.finalize())
}

fn hex(bytes: &[u8]) -> String {
    let mut encoded = String::with_capacity(bytes.len() * 2);
    for byte in bytes {
        write!(encoded, "{byte:02x}").expect("writing to a string cannot fail");
    }
    encoded
}

fn write_response(
    output: &mut impl io::Write,
    id: u64,
    result: Result<PeerSnapshot>,
) -> Result<()> {
    let response = match result {
        Ok(snapshot) => PeerResponse {
            id,
            ok: true,
            snapshot: Some(snapshot),
            error: None,
        },
        Err(error) => PeerResponse {
            id,
            ok: false,
            snapshot: None,
            error: Some(format!("{error:#}")),
        },
    };
    serde_json::to_writer(&mut *output, &response)?;
    writeln!(output)?;
    output.flush()?;
    Ok(())
}
