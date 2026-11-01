# BetterOffice WOPI app

Bun and TypeScript app that exposes BetterOffice to WOPI hosts such as Nextcloud, ownCloud, OpenCloud and Seafile.

Run with `bun src/server.ts`. Configuration comes from the environment: `WOPI_PUBLIC_URL` and `WOPI_HOSTS` are required, see `src/config.ts`. `WOPI_RECONNECT_GRACE_MS` is a positive integer in milliseconds, defaulting to `30000`, that keeps a disconnected handle alive for reconnection.

Endpoints in this first slice: `GET /hosting/discovery`, `GET /hosting/capabilities` and `GET /healthz`.
