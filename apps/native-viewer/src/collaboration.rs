use std::fmt;
use std::sync::mpsc::{
    Receiver as EventReceiver, SyncSender as EventSender, TrySendError as EventTrySendError,
    sync_channel,
};
use std::thread;
use std::time::Duration;

use anyhow::{Context, Result, bail};
use futures_util::{SinkExt, StreamExt};
use serde::Deserialize;
use tokio::runtime::Builder;
use tokio::sync::mpsc::{Receiver, Sender, channel, error::TrySendError};
use tokio_tungstenite::connect_async_with_config;
use tokio_tungstenite::tungstenite::Message;
use tokio_tungstenite::tungstenite::protocol::WebSocketConfig;
use url::Url;

use crate::collaboration_protocol::MAX_FRAME_BYTES;

pub const DEFAULT_RELAY_ORIGIN: &str = "https://betteroffice-collaboration-relay.elia7.workers.dev";
pub const BROWSER_SEED_CLIENT_ID: u64 = 1;
const INITIAL_RETRY: Duration = Duration::from_millis(250);
const MAX_RETRY: Duration = Duration::from_secs(5);
const COMMAND_QUEUE_CAPACITY: usize = 64;
const EVENT_QUEUE_CAPACITY: usize = 16;

#[derive(Clone, Debug)]
pub struct CollaborationConfig {
    room: String,
    url: String,
    client_id: u64,
}

impl CollaborationConfig {
    pub fn new(room: String, relay_origin: &str) -> Result<Self> {
        if room.is_empty() {
            bail!("--room cannot be empty");
        }
        if room.encode_utf16().count() > 128 {
            bail!("--room cannot exceed 128 characters");
        }
        let url = room_url(relay_origin, &room)?;
        let client_id = random_client_id()?;
        Ok(Self {
            room,
            url,
            client_id,
        })
    }

    pub fn room(&self) -> &str {
        &self.room
    }

    pub fn client_id(&self) -> u64 {
        self.client_id
    }

    #[cfg(test)]
    pub(crate) fn for_test(room: &str, client_id: u64) -> Self {
        Self {
            room: room.to_owned(),
            url: "ws://127.0.0.1:9".to_owned(),
            client_id,
        }
    }
}

#[derive(Debug)]
pub enum TransportEvent {
    Connecting,
    Connected,
    Binary(Vec<u8>),
    PeerCount(usize),
    Reconnecting { delay: Duration, reason: String },
    Failed(String),
}

#[derive(Clone)]
pub struct TransportEventSender {
    sender: EventSender<TransportEvent>,
}

pub struct TransportEventReceiver {
    receiver: EventReceiver<TransportEvent>,
}

pub fn transport_event_channel() -> (TransportEventSender, TransportEventReceiver) {
    let (sender, receiver) = sync_channel(EVENT_QUEUE_CAPACITY);
    (
        TransportEventSender { sender },
        TransportEventReceiver { receiver },
    )
}

impl TransportEventSender {
    pub fn try_send(&self, event: TransportEvent) -> bool {
        match self.sender.try_send(event) {
            Ok(()) => true,
            Err(EventTrySendError::Full(_) | EventTrySendError::Disconnected(_)) => false,
        }
    }
}

impl TransportEventReceiver {
    pub fn try_recv(&self) -> Option<TransportEvent> {
        self.receiver.try_recv().ok()
    }
}

pub(crate) enum TransportCommand {
    Send(Vec<u8>),
    Shutdown,
}

pub struct CollaborationClient {
    config: CollaborationConfig,
    commands: Sender<TransportCommand>,
    status: ConnectionStatus,
    peer_count: Option<usize>,
    protocol_error: Option<String>,
}

#[derive(Clone, Debug, Eq, PartialEq)]
enum ConnectionStatus {
    Connecting,
    Connected { synced: bool },
    Reconnecting { delay: Duration, reason: String },
    Failed(String),
    Denied(String),
}

impl CollaborationClient {
    pub fn start(
        config: CollaborationConfig,
        notify: impl Fn(TransportEvent) -> bool + Send + 'static,
    ) -> Result<Self> {
        let (commands, receiver) = channel(COMMAND_QUEUE_CAPACITY);
        let url = config.url.clone();
        thread::Builder::new()
            .name("document-collaboration".to_owned())
            .spawn(move || {
                let runtime = Builder::new_current_thread().enable_all().build();
                match runtime {
                    Ok(runtime) => runtime.block_on(run_transport(url, receiver, notify)),
                    Err(error) => {
                        notify(TransportEvent::Failed(format!(
                            "create collaboration runtime: {error}"
                        )));
                    }
                }
            })
            .context("start collaboration transport")?;
        Ok(Self {
            config,
            commands,
            status: ConnectionStatus::Connecting,
            peer_count: None,
            protocol_error: None,
        })
    }

    #[cfg(test)]
    pub(crate) fn detached(config: CollaborationConfig) -> (Self, Receiver<TransportCommand>) {
        let (commands, receiver) = channel(COMMAND_QUEUE_CAPACITY);
        (
            Self {
                config,
                commands,
                status: ConnectionStatus::Connecting,
                peer_count: None,
                protocol_error: None,
            },
            receiver,
        )
    }

    pub fn room(&self) -> &str {
        self.config.room()
    }

    pub fn client_id(&self) -> u64 {
        self.config.client_id()
    }

    pub fn is_connected(&self) -> bool {
        matches!(self.status, ConnectionStatus::Connected { .. })
    }

    pub fn is_synced(&self) -> bool {
        matches!(self.status, ConnectionStatus::Connected { synced: true })
    }

    pub fn send(&self, frame: Vec<u8>) -> Result<()> {
        if frame.len() > MAX_FRAME_BYTES {
            bail!("collaboration frame exceeds {MAX_FRAME_BYTES} bytes");
        }
        self.commands
            .try_send(TransportCommand::Send(frame))
            .map_err(|error| match error {
                TrySendError::Full(_) => anyhow::anyhow!("collaboration command queue is full"),
                TrySendError::Closed(_) => anyhow::anyhow!("collaboration transport stopped"),
            })
    }

    pub fn apply_transport_event(&mut self, event: &TransportEvent) {
        match event {
            TransportEvent::Connecting => {
                self.status = ConnectionStatus::Connecting;
                self.peer_count = None;
            }
            TransportEvent::Connected => {
                self.status = ConnectionStatus::Connected { synced: false };
                self.protocol_error = None;
            }
            TransportEvent::PeerCount(count) => self.peer_count = Some(*count),
            TransportEvent::Reconnecting { delay, reason } => {
                self.status = ConnectionStatus::Reconnecting {
                    delay: *delay,
                    reason: reason.clone(),
                };
                self.peer_count = None;
            }
            TransportEvent::Failed(error) => {
                self.status = ConnectionStatus::Failed(error.clone());
                self.peer_count = None;
            }
            TransportEvent::Binary(_) => {}
        }
    }

    pub fn mark_synced(&mut self) {
        if matches!(self.status, ConnectionStatus::Connected { .. }) {
            self.status = ConnectionStatus::Connected { synced: true };
            self.protocol_error = None;
        }
    }

    pub fn report_protocol_error(&mut self, error: impl fmt::Display) {
        self.protocol_error = Some(error.to_string());
    }

    pub fn deny(&mut self, reason: String) {
        self.status = ConnectionStatus::Denied(reason);
        let _ = self.commands.try_send(TransportCommand::Shutdown);
    }

    pub fn status_line(&self) -> String {
        let peers = self
            .peer_count
            .map(|count| format!(", {count} peer{}", if count == 1 { "" } else { "s" }))
            .unwrap_or_default();
        let detail = if let Some(error) = &self.protocol_error {
            format!("protocol error: {error}")
        } else {
            match &self.status {
                ConnectionStatus::Connecting => "connecting".to_owned(),
                ConnectionStatus::Connected { synced: false } => "connected, syncing".to_owned(),
                ConnectionStatus::Connected { synced: true } => "synced".to_owned(),
                ConnectionStatus::Reconnecting { delay, reason } => format!(
                    "offline, reconnecting in {} ms: {reason}",
                    delay.as_millis()
                ),
                ConnectionStatus::Failed(error) => format!("stopped: {error}"),
                ConnectionStatus::Denied(reason) => format!("denied: {reason}"),
            }
        };
        format!("Room {}: {detail}{peers}", self.room())
    }
}

impl Drop for CollaborationClient {
    fn drop(&mut self) {
        let _ = self.commands.try_send(TransportCommand::Shutdown);
    }
}

#[derive(Deserialize)]
struct PeerMessage {
    #[serde(rename = "type")]
    message_type: String,
    count: usize,
}

async fn run_transport(
    url: String,
    mut commands: Receiver<TransportCommand>,
    notify: impl Fn(TransportEvent) -> bool,
) {
    let mut retry = INITIAL_RETRY;
    loop {
        if !notify(TransportEvent::Connecting) {
            return;
        }
        let config = WebSocketConfig::default()
            .max_message_size(Some(MAX_FRAME_BYTES))
            .max_frame_size(Some(MAX_FRAME_BYTES));
        let connection = connect_or_shutdown(&url, config, &mut commands).await;
        let Some(connection) = connection else {
            return;
        };
        let socket = match connection {
            Ok((socket, _)) => socket,
            Err(error) => {
                let delay = retry;
                if !notify(TransportEvent::Reconnecting {
                    delay,
                    reason: error.to_string(),
                }) {
                    return;
                }
                if !wait_for_retry(delay, &mut commands).await {
                    return;
                }
                retry = doubled_retry(retry);
                continue;
            }
        };
        retry = INITIAL_RETRY;
        if !notify(TransportEvent::Connected) {
            return;
        }
        let (mut writer, mut reader) = socket.split();
        let reason = loop {
            tokio::select! {
                command = commands.recv() => match command {
                    Some(TransportCommand::Send(frame)) => {
                        if frame.len() > MAX_FRAME_BYTES {
                            if !notify(TransportEvent::Failed(format!(
                                "outgoing frame exceeds {MAX_FRAME_BYTES} bytes"
                            ))) {
                                return;
                            }
                            continue;
                        }
                        if let Err(error) = writer.send(Message::Binary(frame.into())).await {
                            break error.to_string();
                        }
                    }
                    Some(TransportCommand::Shutdown) | None => {
                        let _ = writer.close().await;
                        return;
                    }
                },
                message = reader.next() => match message {
                    Some(Ok(Message::Binary(bytes))) => {
                        if bytes.len() <= MAX_FRAME_BYTES {
                            if !notify(TransportEvent::Binary(bytes.to_vec())) {
                                let _ = writer.close().await;
                                return;
                            }
                        } else {
                            break format!("incoming frame exceeds {MAX_FRAME_BYTES} bytes");
                        }
                    }
                    Some(Ok(Message::Text(text))) => {
                        if let Ok(message) = serde_json::from_str::<PeerMessage>(&text)
                            && message.message_type == "peers"
                            && !notify(TransportEvent::PeerCount(message.count))
                        {
                            let _ = writer.close().await;
                            return;
                        }
                    }
                    Some(Ok(Message::Close(frame))) => {
                        break frame.map_or_else(
                            || "connection closed".to_owned(),
                            |frame| {
                                if frame.reason.is_empty() {
                                    format!("connection closed ({})", u16::from(frame.code))
                                } else {
                                    frame.reason.to_string()
                                }
                            },
                        );
                    }
                    Some(Ok(Message::Ping(_) | Message::Pong(_) | Message::Frame(_))) => {}
                    Some(Err(error)) => break error.to_string(),
                    None => break "connection ended".to_owned(),
                }
            }
        };
        let delay = retry;
        if !notify(TransportEvent::Reconnecting { delay, reason }) {
            return;
        }
        if !wait_for_retry(delay, &mut commands).await {
            return;
        }
        retry = doubled_retry(retry);
    }
}

async fn connect_or_shutdown(
    url: &str,
    config: WebSocketConfig,
    commands: &mut Receiver<TransportCommand>,
) -> Option<
    Result<
        (
            tokio_tungstenite::WebSocketStream<
                tokio_tungstenite::MaybeTlsStream<tokio::net::TcpStream>,
            >,
            tokio_tungstenite::tungstenite::handshake::client::Response,
        ),
        tokio_tungstenite::tungstenite::Error,
    >,
> {
    let connection = connect_async_with_config(url, Some(config), false);
    tokio::pin!(connection);
    loop {
        tokio::select! {
            result = &mut connection => return Some(result),
            command = commands.recv() => match command {
                Some(TransportCommand::Send(_)) => {}
                Some(TransportCommand::Shutdown) | None => return None,
            }
        }
    }
}

async fn wait_for_retry(delay: Duration, commands: &mut Receiver<TransportCommand>) -> bool {
    let sleep = tokio::time::sleep(delay);
    tokio::pin!(sleep);
    loop {
        tokio::select! {
            () = &mut sleep => return true,
            command = commands.recv() => match command {
                Some(TransportCommand::Send(_)) => {}
                Some(TransportCommand::Shutdown) | None => return false,
            }
        }
    }
}

fn doubled_retry(delay: Duration) -> Duration {
    delay.saturating_mul(2).min(MAX_RETRY)
}

fn room_url(relay_origin: &str, room: &str) -> Result<String> {
    let mut url = Url::parse(relay_origin).context("invalid collaboration relay origin")?;
    let scheme = match url.scheme() {
        "http" => "ws",
        "https" => "wss",
        "ws" => "ws",
        "wss" => "wss",
        scheme => bail!("unsupported collaboration relay scheme {scheme}"),
    };
    url.set_scheme(scheme)
        .map_err(|_| anyhow::anyhow!("invalid collaboration relay scheme"))?;
    url.set_query(None);
    url.set_fragment(None);
    url.path_segments_mut()
        .map_err(|_| anyhow::anyhow!("collaboration relay origin cannot hold a room path"))?
        .clear()
        .push("room")
        .push(room);
    Ok(url.into())
}

fn random_client_id() -> Result<u64> {
    loop {
        let mut bytes = [0_u8; 8];
        getrandom::fill(&mut bytes).map_err(|error| anyhow::anyhow!(error))?;
        let client_id = u64::from_le_bytes(bytes) & MAX_VAR_UINT;
        if client_id > BROWSER_SEED_CLIENT_ID {
            return Ok(client_id);
        }
    }
}

const MAX_VAR_UINT: u64 = 9_007_199_254_740_991;

#[cfg(test)]
mod tests {
    use std::net::TcpListener;
    use std::sync::mpsc;

    use tokio_tungstenite::accept_async;

    use super::*;

    #[test]
    fn websocket_transport_connects_and_round_trips_binary_frames() {
        let listener = TcpListener::bind("127.0.0.1:0").unwrap();
        let address = listener.local_addr().unwrap();
        listener.set_nonblocking(true).unwrap();
        let server = thread::spawn(move || {
            Builder::new_current_thread()
                .enable_all()
                .build()
                .unwrap()
                .block_on(async move {
                    let listener = tokio::net::TcpListener::from_std(listener).unwrap();
                    let (stream, _) = listener.accept().await.unwrap();
                    let mut socket = accept_async(stream).await.unwrap();
                    let Some(Ok(Message::Binary(frame))) = socket.next().await else {
                        panic!("native client did not send a binary frame");
                    };
                    socket.send(Message::Binary(frame)).await.unwrap();
                    socket.close(None).await.unwrap();
                });
        });
        let config =
            CollaborationConfig::new("transport".to_owned(), &format!("http://{address}")).unwrap();
        let (events, received) = mpsc::channel();
        let client =
            CollaborationClient::start(config, move |event| events.send(event).is_ok()).unwrap();
        loop {
            match received.recv_timeout(Duration::from_secs(3)).unwrap() {
                TransportEvent::Connected => break,
                TransportEvent::Connecting => {}
                event => panic!("unexpected connection event: {event:?}"),
            }
        }
        client.send(vec![3]).unwrap();
        loop {
            match received.recv_timeout(Duration::from_secs(3)).unwrap() {
                TransportEvent::Binary(frame) => {
                    assert_eq!(frame, [3]);
                    break;
                }
                TransportEvent::PeerCount(_) => {}
                event => panic!("unexpected transport event: {event:?}"),
            }
        }
        drop(client);
        server.join().unwrap();
    }

    #[test]
    fn room_url_matches_the_browser_transport() {
        assert_eq!(
            room_url("https://relay.example/base?old=1#hash", "a room/雪").unwrap(),
            "wss://relay.example/room/a%20room%2F%E9%9B%AA"
        );
        assert_eq!(
            room_url("http://127.0.0.1:8787", "native").unwrap(),
            "ws://127.0.0.1:8787/room/native"
        );
    }

    #[test]
    fn room_validation_matches_the_relay() {
        assert!(CollaborationConfig::new(String::new(), DEFAULT_RELAY_ORIGIN).is_err());
        assert!(CollaborationConfig::new("x".repeat(129), DEFAULT_RELAY_ORIGIN).is_err());
        assert!(CollaborationConfig::new("x".to_owned(), "file:///tmp/relay").is_err());
    }

    #[test]
    fn reconnect_backoff_caps_at_five_seconds() {
        assert_eq!(
            doubled_retry(Duration::from_millis(250)),
            Duration::from_millis(500)
        );
        assert_eq!(
            doubled_retry(Duration::from_secs(4)),
            Duration::from_secs(5)
        );
        assert_eq!(
            doubled_retry(Duration::from_secs(5)),
            Duration::from_secs(5)
        );
    }

    #[test]
    fn command_queue_rejects_overflow() {
        let config = CollaborationConfig::new("bounded".to_owned(), "ws://127.0.0.1:9").unwrap();
        let (client, _commands) = CollaborationClient::detached(config);
        for _ in 0..COMMAND_QUEUE_CAPACITY {
            client.send(vec![1]).unwrap();
        }
        assert!(
            client
                .send(vec![1])
                .unwrap_err()
                .to_string()
                .contains("queue is full")
        );
    }

    #[test]
    fn transport_event_queue_rejects_overflow() {
        let (events, _receiver) = transport_event_channel();
        for _ in 0..EVENT_QUEUE_CAPACITY {
            assert!(events.try_send(TransportEvent::Binary(vec![1])));
        }
        assert!(!events.try_send(TransportEvent::Binary(vec![1])));
    }
}
