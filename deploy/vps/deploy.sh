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

# --- detect the Docker network Traefik routes on ---
NETWORK=$(docker inspect "$TRAEFIK_ID" \
  --format '{{range $k, $v := .NetworkSettings.Networks}}{{$k}}{{"\n"}}{{end}}' | head -n1)

# --- detect the certresolver + entrypoint the existing n8n router uses ---
ALL_LABELS=$(docker ps -q | xargs -r docker inspect \
  --format '{{range $k, $v := .Config.Labels}}{{$k}}={{$v}}{{"\n"}}{{end}}' 2>/dev/null || true)
CERTRESOLVER=$(printf '%s' "$ALL_LABELS" | grep -m1 -oE 'tls\.certresolver=[^ ]+' | cut -d= -f2 || true)
ENTRYPOINT=$(printf '%s' "$ALL_LABELS" | grep -m1 -oE '\.entrypoints=[^ ]+' | cut -d= -f2 || true)
CERTRESOLVER="${CERTRESOLVER:-mytlschallenge}"
ENTRYPOINT="${ENTRYPOINT:-websecure}"

echo "Domain:          $DOMAIN"
echo "Traefik network: $NETWORK"
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
  read -rp "KB search token (must match the n8n 'KB Search Header Auth' credential; leave empty to run without RAG): " KB_TOKEN
  cat > "$APP_DIR/.env" <<EOF
OPENAI_API_KEY=$OPENAI_KEY
OPENAI_MODEL=gpt-4o-mini
KB_SEARCH_URL=https://n8n-jp6p.srv1747149.hstgr.cloud/webhook/kb-search
KB_SEARCH_TOKEN=$KB_TOKEN
KB_TIMEOUT_MS=8000
EOF
  chmod 600 "$APP_DIR/.env"
  echo "Wrote $APP_DIR/.env (edit it later to change keys, then re-run this script)."
fi

# --- compose file: joins Traefik's network so it can route to us ---
cat > "$APP_DIR/docker-compose.yml" <<EOF
services:
  outlook-ai:
    build: ./app
    restart: unless-stopped
    env_file: .env
    networks: [traefik-net]
    labels:
      - traefik.enable=true
      - traefik.http.routers.outlook-ai.rule=Host(\`$DOMAIN\`)
      - traefik.http.routers.outlook-ai.entrypoints=$ENTRYPOINT
      - traefik.http.routers.outlook-ai.tls=true
      - traefik.http.routers.outlook-ai.tls.certresolver=$CERTRESOLVER
      - traefik.http.services.outlook-ai.loadbalancer.server.port=3000

networks:
  traefik-net:
    external: true
    name: $NETWORK
EOF

docker compose -f "$APP_DIR/docker-compose.yml" up -d --build

echo
echo "Deployed. First HTTPS cert issuance can take ~30-60 seconds, then test:"
echo "  https://$DOMAIN/taskpane.html"
