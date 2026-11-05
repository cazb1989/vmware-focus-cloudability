"""
Lambda: RVTools (vInfo) en S3 -> fotos de inventario -> drop FOCUS 1.0 para Cloudability.

Disparadores:
  - S3 "Object Created" en INBOX_PREFIX (vía EventBridge): guarda la foto y regenera el mes.
  - Programado (EventBridge, diario): regenera el mes en curso (rellena días sin foto con la
    última foto, hasta max_carry_forward_days) y, los primeros días del mes, el mes anterior.
  - Manual (aws lambda invoke):
      {"action": "rebuild", "months": ["2026-11"]}
      {"action": "backfill", "from": "2026-11-01", "to": "2026-11-04"}   (carga inicial)

Variables de entorno:
  BUCKET, SUB_DIR (vmware), INBOX_PREFIX (vmware-inbox), STATE_PREFIX (vmware-state)
"""
import json
import os
import urllib.parse
from datetime import date, datetime, timedelta, timezone

import boto3
import rvtools_focus as rf

s3 = boto3.client("s3")


def _env():
    return (os.environ["BUCKET"], os.environ.get("SUB_DIR", "vmware"),
            os.environ.get("INBOX_PREFIX", "vmware-inbox").strip("/"),
            os.environ.get("STATE_PREFIX", "vmware-state").strip("/"))


def load_config(store, state_prefix):
    raw = store.get(f"{state_prefix}/config.json")
    return rf.merged_config(json.loads(raw) if raw else None)


def rebuild(store, state_prefix, months, cfg, bucket, sub_dir, today):
    out = []
    for m in sorted(set(months)):
        end = min(today, rf.month_bounds(m)[1] - timedelta(days=1))
        out.append(rf.build_month(store, state_prefix, m, end, cfg, bucket, sub_dir, store))
    return out


def ingest(store, state_prefix, bucket, key, event_time, cfg):
    name = key.rsplit("/", 1)[-1]
    snap = rf.date_from_name(key) or event_time.date()
    try:
        data = s3.get_object(Bucket=bucket, Key=key)["Body"].read()
    except s3.exceptions.NoSuchKey:
        # reintento de un evento ya procesado (el original se movió a processed/)
        return snap, {"note": "ya procesado"}
    try:
        by_vc, stats = rf.parse_rvtools(data, name, cfg)
    except Exception:
        s3.copy_object(Bucket=bucket, Key=f"{state_prefix}/failed/{snap}/{name}",
                       CopySource={"Bucket": bucket, "Key": key})
        raise
    rf.save_snapshots(store, state_prefix, by_vc, snap, key)
    # archivo original: se archiva fuera del inbox (y fuera de la carpeta que lee Cloudability)
    s3.copy_object(Bucket=bucket, Key=f"{state_prefix}/processed/{snap}/{name}",
                   CopySource={"Bucket": bucket, "Key": key})
    s3.delete_object(Bucket=bucket, Key=key)
    return snap, stats


def lambda_handler(event, context):
    event = event or {}
    bucket, sub_dir, inbox, state = _env()
    store = rf.S3Store(bucket, s3)
    cfg = load_config(store, state)
    today = (date.fromisoformat(event["today"]) if event.get("today")
             else datetime.now(timezone.utc).date())
    result = {"today": today.isoformat()}

    if event.get("source") == "aws.s3":                                   # archivo nuevo
        key = urllib.parse.unquote_plus(event["detail"]["object"]["key"])
        if not key.startswith(inbox + "/") or key.endswith("/"):
            return {"skipped": key}
        t = datetime.fromisoformat(event.get("time", today.isoformat() + "T00:00:00Z")
                                   .replace("Z", "+00:00"))
        snap, stats = ingest(store, state, bucket, key, t, cfg)
        result.update(ingested=key, snapshot_date=snap.isoformat(), stats=stats)
        months = [snap.replace(day=1)]
        if snap.replace(day=1) != today.replace(day=1):
            months.append(today.replace(day=1))
        result["builds"] = rebuild(store, state, months, cfg, bucket, sub_dir, today)

    elif event.get("action") == "backfill":                                # carga inicial
        start = date.fromisoformat(event["from"])
        end = date.fromisoformat(event["to"]) if event.get("to") else today
        created = rf.backfill(store, state, start, end)
        months, d = set(), start
        while d <= end:
            months.add(d.replace(day=1))
            d += timedelta(days=1)
        result.update(backfilled=len(created),
                      builds=rebuild(store, state, months, cfg, bucket, sub_dir, today))

    elif event.get("action") == "rebuild":                                 # manual
        months = [date.fromisoformat(m + "-01") for m in event.get("months") or [f"{today:%Y-%m}"]]
        result["builds"] = rebuild(store, state, months, cfg, bucket, sub_dir, today)

    else:                                                                  # diario
        months = [today.replace(day=1)]
        if today.day <= int(cfg["max_carry_forward_days"]) + 1:
            months.append((today.replace(day=1) - timedelta(days=1)).replace(day=1))
        result["builds"] = rebuild(store, state, months, cfg, bucket, sub_dir, today)

    print(json.dumps(result, ensure_ascii=False, default=str))
    return result
