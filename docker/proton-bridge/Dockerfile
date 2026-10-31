# Proton Mail Bridge built from the official source, for local IMAP testing.
# Based on shenxn/protonmail-bridge-docker's `build` recipe (entrypoint.sh and
# gpgparams vendored from commit 5874fde), plus the libfido2/libcbor/openssl
# dependencies Bridge 3.2x links against, which that recipe lacks.
FROM debian:sid-slim AS build

ARG BRIDGE_VERSION

RUN apt-get update && apt-get install -y --no-install-recommends \
      ca-certificates git golang build-essential pkg-config \
      libsecret-1-dev libfido2-dev libcbor-dev libssl-dev \
    && rm -rf /var/lib/apt/lists/*

ADD https://github.com/ProtonMail/proton-bridge.git#${BRIDGE_VERSION} /build/
WORKDIR /build/
RUN make build-nogui vault-editor

FROM debian:sid-slim

EXPOSE 25/tcp
EXPOSE 143/tcp

RUN apt-get update && apt-get install -y --no-install-recommends \
      socat pass libsecret-1-0 libfido2-1 libcbor0.10 ca-certificates procps \
    && rm -rf /var/lib/apt/lists/*

COPY gpgparams entrypoint.sh /protonmail/
COPY --from=build /build/bridge /build/proton-bridge /build/vault-editor /protonmail/

ENTRYPOINT ["bash", "/protonmail/entrypoint.sh"]
