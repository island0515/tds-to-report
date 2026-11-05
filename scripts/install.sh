#!/bin/sh
set -eu
cd "$(dirname "$0")/.."
mkdir -p data/uploads data/output data/temp data/logs blank-templates standards
for d in data/uploads data/output data/temp data/logs; do touch "$d/.gitkeep"; done
if [ ! -f .env ]; then
  cp .env.example .env
  token=""
  if command -v openssl >/dev/null 2>&1; then token="$(openssl rand -hex 32)"; fi
  if [ -z "$token" ]; then token="$(python3 -c 'import secrets; print(secrets.token_hex(32))' 2>/dev/null || true)"; fi
  if [ -n "$token" ]; then
    sed -i "s/REPLACE_WITH_A_LONG_RANDOM_SECRET/$token/" .env
  fi
  echo "Created .env. Edit APP_IMAGE before starting."
fi
echo "Put your blank DOCX in blank-templates/ and standards files in standards/."
echo "Then run: docker compose pull && docker compose up -d"
