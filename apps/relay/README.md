# @betteroffice/collaboration-relay

Reference collaboration relay: one Cloudflare Durable Object per room speaks the
Yjs sync-v1 wire protocol. It replays retained document frames to joining sockets
and broadcasts new binary frames. It never parses Office files.

## Run

```bash
bun run --filter @betteroffice/collaboration-relay dev
bun run --filter @betteroffice/collaboration-relay deploy
```

WebSockets connect at `/room/:roomId`. `GET /` returns `ok` for health checks.

## Limits

- 16 MiB maximum frame (`MAX_COLLABORATION_FRAME_BYTES`).
- 512 frames or 64 MiB of retained history per room; join replay capped at 64 MiB.
- 32 KiB awareness payload per frame; frames carrying awareness rate-limited
  per socket (1009 close on oversize, 1008 close on rate abuse).

## Threat model

Room IDs are bearer credentials: anyone holding one gets read, write, and
retained-state replay. The relay performs no authentication and validates framing
only; CRDT validity is enforced inside the engine. Generate unguessable room IDs,
treat a shared link as shared access, and add authentication before resolving a
room in production.
