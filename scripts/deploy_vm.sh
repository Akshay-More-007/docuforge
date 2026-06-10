#!/usr/bin/env bash
# DocuForge — one-shot VM bootstrap + deploy (Ubuntu/Debian or Oracle Linux).
# Run ON the VM:  bash deploy_vm.sh
# Expects a real .env to exist at ~/docuforge.env (scp it there first).
set -euo pipefail

REPO_URL="https://github.com/Akshay-More-007/docuforge.git"
APP_DIR="$HOME/docuforge"
ENV_SRC="$HOME/docuforge.env"

echo "==> Installing Docker (if missing)"
if ! command -v docker >/dev/null 2>&1; then
    curl -fsSL https://get.docker.com | sh
    sudo usermod -aG docker "$USER"
fi

echo "==> Cloning/updating repo"
if [ -d "$APP_DIR/.git" ]; then
    git -C "$APP_DIR" pull --ff-only
else
    git clone "$REPO_URL" "$APP_DIR"
fi

echo "==> Installing .env"
if [ ! -f "$ENV_SRC" ] && [ ! -f "$APP_DIR/.env" ]; then
    echo "ERROR: no .env found. scp your .env to $ENV_SRC first." >&2
    exit 1
fi
[ -f "$ENV_SRC" ] && cp "$ENV_SRC" "$APP_DIR/.env"

echo "==> Building and starting DocuForge"
cd "$APP_DIR"
sudo docker compose -f compose.docuforge.yaml up -d --build

echo "==> Waiting for health"
for i in $(seq 1 30); do
    if curl -fsS http://localhost:8501/healthz >/dev/null 2>&1; then
        echo "DocuForge is UP: http://$(curl -fsS ifconfig.me 2>/dev/null || hostname -I | awk '{print $1}'):8501"
        exit 0
    fi
    sleep 2
done
echo "ERROR: app did not become healthy; check: sudo docker compose -f compose.docuforge.yaml logs" >&2
exit 1
