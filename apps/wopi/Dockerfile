FROM --platform=$BUILDPLATFORM oven/bun:1.4.3@sha256:ec06c3b6cea04192ae6770c434f668ca41d343ad19fa6472216c7b48be39c598 AS build

ARG BUILDARCH

RUN apt-get update \
  && apt-get install -y --no-install-recommends build-essential ca-certificates curl pkg-config python3 \
  && rm -rf /var/lib/apt/lists/*

RUN curl --proto '=https' --tlsv1.2 -sSf https://sh.rustup.rs | sh -s -- -y --profile minimal --default-toolchain stable --target wasm32-unknown-unknown
ENV PATH="/root/.cargo/bin:${PATH}"

RUN case "$BUILDARCH" in \
      amd64) wasm_pack_arch=x86_64; wasm_pack_sha256=c09f971ecaed9a2efc80fdcea7a00ef6b53c7fadc8c57d1f61b53a6aa66b668a ;; \
      arm64) wasm_pack_arch=aarch64; wasm_pack_sha256=e17ef0806381c3a0acb9c9ddad643a49facaa5a2ecf657a421d4d8f3357a24b7 ;; \
      *) echo "Unsupported build architecture: $BUILDARCH" >&2; exit 1 ;; \
    esac \
  && curl -fsSL --retry 5 --retry-all-errors -o /tmp/wasm-pack.tar.gz "https://github.com/rustwasm/wasm-pack/releases/download/v0.15.0/wasm-pack-v0.15.0-${wasm_pack_arch}-unknown-linux-musl.tar.gz" \
  && echo "${wasm_pack_sha256}  /tmp/wasm-pack.tar.gz" | sha256sum -c - \
  && mkdir -p /tmp/wasm-pack \
  && tar -xzf /tmp/wasm-pack.tar.gz -C /tmp/wasm-pack --strip-components=1 \
  && install -m 0755 /tmp/wasm-pack/wasm-pack /usr/local/bin/wasm-pack \
  && rm -rf /tmp/wasm-pack /tmp/wasm-pack.tar.gz \
  && wasm-pack --version

RUN case "$BUILDARCH" in \
      amd64) binaryen_arch=x86_64; binaryen_sha256=195ddc94f9bc89f45abdabb0b9eea86023d727ba90eac8b35b80f2544fc30572 ;; \
      arm64) binaryen_arch=aarch64; binaryen_sha256=c58562417836c5d0493d89bdefc434933bdc097db641b483df86bcfa557a107f ;; \
      *) echo "Unsupported build architecture: $BUILDARCH" >&2; exit 1 ;; \
    esac \
  && curl -fsSL --retry 5 --retry-all-errors -o /tmp/binaryen.tar.gz "https://github.com/WebAssembly/binaryen/releases/download/version_132/binaryen-version_132-${binaryen_arch}-linux.tar.gz" \
  && echo "${binaryen_sha256}  /tmp/binaryen.tar.gz" | sha256sum -c - \
  && mkdir -p /opt/binaryen \
  && tar -xzf /tmp/binaryen.tar.gz -C /opt/binaryen --strip-components=1 \
  && rm -f /tmp/binaryen.tar.gz
ENV PATH="/opt/binaryen/bin:${PATH}"

WORKDIR /app
COPY . .
RUN bun install --frozen-lockfile
RUN bun run build:xlsx-wasm && bun run build:docx-wasm && bun run build:pptx-wasm
RUN bun run build:packages
RUN bun run --filter @betteroffice/wopi build:web
RUN mkdir -p /manifests \
  && find apps packages bindings crates -name package.json -not -path "*/node_modules/*" -exec cp --parents {} /manifests/ \; \
  && cp package.json bun.lock /manifests/
RUN find /app/apps/wopi/src /app/packages/docx/src /app/packages/xlsx/src /app/packages/pptx/src /app/shared -type f \( -name "*.test.ts" -o -name "*.test.tsx" -o -name "*.test.d.ts" \) -delete

FROM oven/bun:1.4.3-slim@sha256:14a223fa35747ec728bce300ef85888bc5f7958164d15bbba6a54b22b52a3288 AS runtime

ENV NODE_ENV=production
ENV PORT=3000
ENV WOPI_PROOF_KEY_FILE=/data/proof-key.pem

WORKDIR /app

COPY --from=build --chown=root:root /manifests ./
RUN bun install --production --frozen-lockfile --filter @betteroffice/wopi

COPY --from=build --chown=root:root /app/apps/wopi/package.json ./apps/wopi/package.json
COPY --from=build --chown=root:root /app/apps/wopi/src ./apps/wopi/src
COPY --from=build --chown=root:root /app/apps/wopi/web/dist ./apps/wopi/web/dist
COPY --from=build --chown=root:root /app/packages/docx/src ./packages/docx/src
COPY --from=build --chown=root:root /app/packages/xlsx/src ./packages/xlsx/src
COPY --from=build --chown=root:root /app/packages/pptx/src ./packages/pptx/src
COPY --from=build --chown=root:root /app/packages/fonts/assets ./packages/fonts/assets
COPY --from=build --chown=root:root /app/packages/fonts-cjk/assets ./packages/fonts-cjk/assets
COPY --from=build --chown=root:root /app/shared/collaboration-limits.ts ./shared/collaboration-limits.ts
COPY --from=build --chown=root:root /app/shared/collaboration-pipe ./shared/collaboration-pipe
COPY --from=build --chown=root:root /app/shared/media.ts ./shared/media.ts
COPY --from=build --chown=root:root /app/shared/office-session ./shared/office-session

RUN chown -R root:root /app \
  && chmod -R a-w /app \
  && mkdir -p /data && chown bun:bun /data

USER bun
EXPOSE 3000
VOLUME ["/data"]
HEALTHCHECK --interval=30s --timeout=5s --start-period=20s --retries=3 CMD bun -e "fetch('http://127.0.0.1:'+(process.env.PORT||3000)+'/healthz').then((r)=>{if(!r.ok)process.exit(1)}).catch(()=>process.exit(1))"

CMD ["bun", "apps/wopi/src/server.ts"]
