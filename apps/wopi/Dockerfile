FROM oven/bun:1.4.3@sha256:ec06c3b6cea04192ae6770c434f668ca41d343ad19fa6472216c7b48be39c598 AS build

RUN apt-get update \
  && apt-get install -y --no-install-recommends build-essential ca-certificates curl pkg-config python3 \
  && rm -rf /var/lib/apt/lists/*

RUN curl --proto '=https' --tlsv1.2 -sSf https://sh.rustup.rs | sh -s -- -y --profile minimal --default-toolchain stable --target wasm32-unknown-unknown
ENV PATH="/root/.cargo/bin:${PATH}"

RUN curl -fsSL -o /tmp/wasm-pack.tar.gz https://github.com/rustwasm/wasm-pack/releases/download/v0.15.0/wasm-pack-v0.15.0-x86_64-unknown-linux-musl.tar.gz \
  && tar -xzf /tmp/wasm-pack.tar.gz -C /tmp \
  && install -m 0755 /tmp/wasm-pack-v0.15.0-x86_64-unknown-linux-musl/wasm-pack /usr/local/bin/wasm-pack \
  && wasm-pack --version

RUN curl -fsSL --retry 5 --retry-all-errors -o /tmp/binaryen.tar.gz https://github.com/WebAssembly/binaryen/releases/download/version_132/binaryen-version_132-x86_64-linux.tar.gz \
  && echo "195ddc94f9bc89f45abdabb0b9eea86023d727ba90eac8b35b80f2544fc30572  /tmp/binaryen.tar.gz" | sha256sum -c - \
  && mkdir -p /opt/binaryen \
  && tar -xzf /tmp/binaryen.tar.gz -C /opt/binaryen --strip-components=1
ENV PATH="/opt/binaryen/bin:${PATH}"

WORKDIR /app
COPY . .
RUN bun install --frozen-lockfile
RUN bun run build:xlsx-wasm && bun run build:docx-wasm && bun run build:pptx-wasm
RUN bun run build:fonts && bun run build:xlsx && bun run build:docx && bun run build:pptx
RUN bun run --filter @betteroffice/wopi build:web

FROM oven/bun:1.4.3-slim@sha256:14a223fa35747ec728bce300ef85888bc5f7958164d15bbba6a54b22b52a3288 AS runtime

ENV NODE_ENV=production
ENV PORT=3000
ENV WOPI_PROOF_KEY_FILE=/data/proof-key.pem

WORKDIR /app

COPY --from=build --chown=bun:bun /app/apps/wopi/package.json ./apps/wopi/package.json
COPY --from=build --chown=bun:bun /app/apps/wopi/src ./apps/wopi/src
COPY --from=build --chown=bun:bun /app/apps/wopi/web/dist ./apps/wopi/web/dist
COPY --from=build --chown=bun:bun /app/node_modules ./node_modules
COPY --from=build --chown=bun:bun /app/packages/docx/src ./packages/docx/src
COPY --from=build --chown=bun:bun /app/packages/xlsx/src ./packages/xlsx/src
COPY --from=build --chown=bun:bun /app/packages/pptx/src ./packages/pptx/src
COPY --from=build --chown=bun:bun /app/packages/fonts/assets ./packages/fonts/assets
COPY --from=build --chown=bun:bun /app/packages/fonts-cjk/assets ./packages/fonts-cjk/assets
COPY --from=build --chown=bun:bun /app/shared/collaboration-limits.ts ./shared/collaboration-limits.ts
COPY --from=build --chown=bun:bun /app/shared/collaboration-pipe ./shared/collaboration-pipe

RUN mkdir -p /data && chown bun:bun /data

USER bun
EXPOSE 3000
VOLUME ["/data"]
HEALTHCHECK --interval=30s --timeout=5s --start-period=20s --retries=3 CMD bun -e "fetch('http://127.0.0.1:'+(process.env.PORT||3000)+'/healthz').then((r)=>{if(!r.ok)process.exit(1)}).catch(()=>process.exit(1))"

CMD ["bun", "apps/wopi/src/server.ts"]
