#!/usr/bin/env bash
# Deploy the Outlook AI Draft backend next to n8n on a Hostinger-style
# Docker + Traefik VPS. Run as root ON THE VPS:
#
#   bash <(curl -fsSL https://raw.githubusercontent.com/DestroTechnologiesGit/outlook-reply-ai-draft/main/deploy/vps/deploy.sh)
#
# Re-running is safe: it pulls the latest code and rebuilds the container.
set -euo pipefail

DOMAIN="${DOMAIN:-outlook-ai.srv1747149.hstgr.cloud}"
APP_DIR=/opt/outlook-ai
REPO=https://github.com/DestroTechnologiesGit/outlook-reply-ai-draft.git

# --- find the running Traefik container ---
TRAEFIK_ID=$(docker ps --format '{{.ID}} {{.Image}}' | awk 'tolower($2) ~ /traefik/ {print $1; exit}')
if [ -z "$TRAEFIK_ID" ]; then
  echo "ERROR: no running Traefik container found. Is this the VPS that runs n8n?" >&2
  exit 1
fi

# --- detect the certresolver + entrypoint the existing n8n router uses ---
ALL_LABELS=$(docker ps -q | xargs -r docker inspect \
  --format '{{range $k, $v := .Config.Labels}}{{$k}}={{$v}}{{"\n"}}{{end}}' 2>/dev/null || true)
CERTRESOLVER=$(printf '%s' "$ALL_LABELS" | grep -m1 -oE 'tls\.certresolver=[^ ]+' | cut -d= -f2 || true)
ENTRYPOINT=$(printf '%s' "$ALL_LABELS" | grep -m1 -oE '\.entrypoints=[^ ]+' | cut -d= -f2 || true)
CERTRESOLVER="${CERTRESOLVER:-mytlschallenge}"
ENTRYPOINT="${ENTRYPOINT:-websecure}"

# --- pick a network Traefik can route on ---
# Prefer the user-defined network of a container that already carries traefik
# router labels (e.g. n8n) — Traefik demonstrably reaches that one. "host",
# "bridge" and "none" are unusable as compose external networks; if nothing
# user-defined is found (e.g. Traefik runs in host network mode), we skip the
# networks section entirely — a host-mode Traefik reaches container IPs on the
# default compose network just fine.
usable_network_of() {
  docker inspect "$1" \
    --format '{{range $k, $v := .NetworkSettings.Networks}}{{$k}}{{"\n"}}{{end}}' 2>/dev/null |
    grep -vE '^(host|bridge|none)?$' | head -n1
}
NETWORK=""
for id in $(docker ps -q); do
  if docker inspect "$id" --format '{{range $k, $v := .Config.Labels}}{{$k}}{{"\n"}}{{end}}' 2>/dev/null |
    grep -q '^traefik\.http\.routers\.'; then
    NETWORK=$(usable_network_of "$id")
    [ -n "$NETWORK" ] && break
  fi
done
[ -z "$NETWORK" ] && NETWORK=$(usable_network_of "$TRAEFIK_ID")

echo "Domain:          $DOMAIN"
echo "Traefik network: ${NETWORK:-<none — using default compose network>}"
echo "Cert resolver:   $CERTRESOLVER"
echo "Entrypoint:      $ENTRYPOINT"

# --- get / update the code ---
mkdir -p "$APP_DIR"
if [ -d "$APP_DIR/app/.git" ]; then
  git -C "$APP_DIR/app" pull --ff-only
else
  git clone "$REPO" "$APP_DIR/app"
fi

# --- secrets: prompt only on first run, kept in /opt/outlook-ai/.env ---
if [ ! -f "$APP_DIR/.env" ]; then
  read -rp "OpenAI API key (sk-...): " OPENAI_KEY
  cat > "$APP_DIR/.env" <<EOF
OPENAI_API_KEY=$OPENAI_KEY
OPENAI_MODEL=gpt-4o-mini
# n8n KB search webhook (no auth). Comment out to run without RAG.
KB_SEARCH_URL=https://n8n-jp6p.srv1747149.hstgr.cloud/webhook/kb-search
KB_TIMEOUT_MS=8000
EOF
  chmod 600 "$APP_DIR/.env"
  echo "Wrote $APP_DIR/.env (edit it later to change keys, then re-run this script)."
fi

# --- compose file: joins Traefik's network (when one exists) so it can route to us ---
if [ -n "$NETWORK" ]; then
  NET_SERVICE_LINE="    networks: [traefik-net]"
  NET_BLOCK="
networks:
  traefik-net:
    external: true
    name: $NETWORK"
else
  NET_SERVICE_LINE=""
  NET_BLOCK=""
fi

cat > "$APP_DIR/docker-compose.yml" <<EOF
services:
  outlook-ai:
    build: ./app
    restart: unless-stopped
    env_file: .env
$NET_SERVICE_LINE
    labels:
      - traefik.enable=true
      - traefik.http.routers.outlook-ai.rule=Host(\`$DOMAIN\`)
      - traefik.http.routers.outlook-ai.entrypoints=$ENTRYPOINT
      - traefik.http.routers.outlook-ai.tls=true
      - traefik.http.routers.outlook-ai.tls.certresolver=$CERTRESOLVER
      - traefik.http.services.outlook-ai.loadbalancer.server.port=3000
$NET_BLOCK
EOF

docker compose -f "$APP_DIR/docker-compose.yml" up -d --build

echo
echo "Deployed. First HTTPS cert issuance can take ~30-60 seconds, then test:"
echo "  https://$DOMAIN/taskpane.html"
