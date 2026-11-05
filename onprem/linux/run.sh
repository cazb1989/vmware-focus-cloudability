#!/usr/bin/env bash
# Ejecutado por cron: inventario de cada vCenter -> CSV (formato vInfo) -> S3 inbox.
set -uo pipefail
ENV_FILE="${ENV_FILE:-/etc/vmware-focus/env}"
# shellcheck disable=SC1090
set -a; . "$ENV_FILE"; set +a
APP_DIR="${APP_DIR:-/opt/vmware-focus}"
DATA_DIR="${DATA_DIR:-/var/lib/vmware-focus}"
DAY=$(date -u +%F)                      # la foto se registra con la fecha UTC
OUT="$DATA_DIR/$DAY"; mkdir -p "$OUT"
PY="$APP_DIR/venv/bin/python"
rc=0
for HOST in $VCENTER_HOSTS; do
  FILE="$OUT/vinfo_${HOST}_${DAY}.csv"
  echo "$(date -u +%FT%TZ) $HOST: recolectando..."
  if VCENTER_HOST="$HOST" "$PY" "$APP_DIR/vcenter_collector.py" --out "$FILE" ${INCLUDE_TAGS:+--include-tags}; then
    KEY="${INBOX_PREFIX:-vmware-inbox}/$HOST/$DAY/$(basename "$FILE")"
    if "$PY" -c 'import boto3,sys; boto3.client("s3").upload_file(sys.argv[1], sys.argv[2], sys.argv[3]); print("subido s3://%s/%s" % (sys.argv[2], sys.argv[3]))' \
         "$FILE" "$S3_BUCKET" "$KEY"; then
      sleep "${UPLOAD_SPACING_SECONDS:-60}"   # escalona los vCenter para que cada drop los incluya a todos
    else
      echo "ERROR subiendo $FILE"; rc=1
    fi
  else
    echo "ERROR recolectando $HOST"; rc=1
  fi
done
find "$DATA_DIR" -mindepth 1 -maxdepth 1 -type d -mtime +"${KEEP_DAYS:-30}" -exec rm -rf {} + 2>/dev/null
exit $rc
