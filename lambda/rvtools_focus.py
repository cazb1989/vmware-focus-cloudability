#!/usr/bin/env python3
"""
RVTools (pestaña vInfo) -> FOCUS 1.0 -> IBM Cloudability (Connect Custom Data / FOCUS Ingress).

Modelo:
  1. INGEST: cada export de RVTools es una "foto" del inventario en una fecha. Se normaliza
     y se guarda como snapshot por vCenter:  <state>/snapshots/<vcenter>/<AAAA-MM-DD>.json.gz
  2. BUILD:  se reconstruye el MES A LA FECHA con todas las fotos del mes (de todos los vCenter)
     aplicando el tarifario, y se escribe el drop que exige Cloudability:
       <sub_dir>/<vendor>/<AAAAMMDD-AAAAMMDD>/<epoch>/<prefix>-00001.csv.gz
       <sub_dir>/<vendor>/<AAAAMMDD-AAAAMMDD>/<epoch>/<prefix>-Manifest.json
     Si falta la foto de un día, se usa la última foto anterior (hasta max_carry_forward_days).

Funciona igual en local (carpetas) y en AWS Lambda (S3), ver handler.py.

Uso local:
  python3 rvtools_focus.py ingest  RVTools_export.xlsx [--date 2026-11-05] --state ./state
  python3 rvtools_focus.py build   --state ./state --out ./out --root-dir <bucket> [--month 2026-11]
  python3 rvtools_focus.py backfill --state ./state --from 2026-11-01 [--to 2026-11-04]
"""
import argparse
import calendar
import csv
import gzip
import hashlib
import io
import json
import os
import re
import sys
import uuid
from datetime import date, datetime, timedelta, timezone

FOCUS_VERSION = "1.0"

DEFAULT_CONFIG = {
    "currency": "USD",                      # ISO 4217 del tarifario (USD, MXN, PEN, COP, CLP...)
    "invoice_issuer": "TI Corporativa",     # área que "factura" la nube privada
    "provider_name": "VMware",
    "publisher_name": "VMware",
    "vendor": "VMware",                     # carpeta <vendor> y campo Vendor en Cloudability
    "file_prefix": "vmware-focus",          # Manifest Prefix en Cloudability
    # Tarifario MENSUAL por unidad; se prorratea por las horas reales del mes.
    "rates_monthly": {"vcpu": 15.00, "memory_gb": 5.00, "storage_gb": 0.10},
    # Fracción del costo de cómputo (vCPU + memoria) por estado. El disco se cobra siempre.
    "compute_factor": {"poweredOn": 1.0, "poweredOff": 0.0, "suspended": 0.0},
    "storage_basis": "provisioned",         # "provisioned" (Provisioned MiB) o "in_use" (In Use MiB)
    "templates": "storage_only",            # "storage_only" | "exclude" | "charge"
    "srm_placeholders": "exclude",          # "exclude" | "storage_only"  (réplicas DR de SRM)
    "max_carry_forward_days": 3,            # días sin foto que se rellenan con la última foto
    # Atributos personalizados / tags de vSphere (columnas de RVTools) que van a Tags.
    "tags": {"include": [], "exclude_prefixes": ["FW_"], "add_inventory_tags": True},
}

FOCUS_COLUMNS = [
    "BilledCost", "BillingAccountId", "BillingAccountName", "BillingCurrency",
    "BillingPeriodEnd", "BillingPeriodStart", "ChargeCategory", "ChargeClass",
    "ChargeDescription", "ChargeFrequency", "ChargePeriodEnd", "ChargePeriodStart",
    "ConsumedQuantity", "ConsumedUnit", "ContractedCost", "ContractedUnitPrice",
    "EffectiveCost", "InvoiceIssuerName", "ListCost", "ListUnitPrice",
    "PricingQuantity", "PricingUnit", "ProviderName", "PublisherName",
    "RegionId", "RegionName", "ResourceId", "ResourceName", "ResourceType",
    "ServiceCategory", "ServiceName", "SubAccountId", "SubAccountName", "Tags",
]
CUSTOM_COLUMNS = ["x_vCenter", "x_VMwareMoRef", "x_PowerState", "x_IsTemplate", "x_VCPUCount",
                  "x_MemoryGB", "x_StorageGB", "x_GuestOS", "x_Host", "x_SnapshotDate"]
NOT_NULL = ["BilledCost", "BillingAccountId", "BillingAccountName", "BillingCurrency",
            "BillingPeriodEnd", "BillingPeriodStart", "ChargeCategory", "ChargeDescription",
            "ChargePeriodEnd", "ChargePeriodStart", "ContractedCost", "EffectiveCost",
            "InvoiceIssuerName", "ListCost", "PricingQuantity", "PricingUnit",
            "ProviderName", "PublisherName", "ServiceCategory", "ServiceName"]

# Nombres de columna de RVTools (con alias de versiones anteriores).
COLS = {
    "vm": ["VM"], "power": ["Powerstate"], "template": ["Template"],
    "srm": ["SRM Placeholder"], "cpus": ["CPUs"], "memory_mib": ["Memory"],
    "prov_mib": ["Provisioned MiB", "Provisioned MB"], "inuse_mib": ["In Use MiB", "In Use MB"],
    "datacenter": ["Datacenter"], "cluster": ["Cluster"], "host": ["Host"],
    "os_cfg": ["OS according to the configuration file", "OS"],
    "os_tools": ["OS according to the VMware Tools"],
    "moref": ["VM ID", "Object ID"], "uuid": ["VM UUID", "UUID"],
    "vcenter": ["VI SDK Server"], "vcenter_uuid": ["VI SDK UUID"],
    "folder": ["Folder"], "resource_pool": ["Resource pool"],
}
REQUIRED = ["vm", "power", "cpus", "memory_mib", "prov_mib", "vcenter"]


def merged_config(user_cfg=None):
    cfg = json.loads(json.dumps(DEFAULT_CONFIG))
    for k, v in (user_cfg or {}).items():
        if isinstance(v, dict) and isinstance(cfg.get(k), dict):
            cfg[k].update(v)
        else:
            cfg[k] = v
    return cfg


# =============================================================================
# 1) LECTURA DEL EXPORT DE RVTOOLS
# =============================================================================
def _read_table(data: bytes, filename: str):
    """Devuelve (header, filas) de la pestaña vInfo. Acepta .xlsx o .csv."""
    name = filename.lower()
    if name.endswith((".xlsx", ".xlsm")):
        import openpyxl  # empaquetado en la Lambda
        wb = openpyxl.load_workbook(io.BytesIO(data), read_only=True, data_only=True)
        ws = wb["vInfo"] if "vInfo" in wb.sheetnames else wb.worksheets[0]
        rows = ws.iter_rows(values_only=True)
        header = [str(h).strip() if h is not None else "" for h in next(rows)]
        return header, [r for r in rows if any(c not in (None, "") for c in r)]
    text = data.decode("utf-8-sig", errors="replace")
    dialect = csv.Sniffer().sniff(text[:5000], delimiters=",;\t")
    reader = csv.reader(io.StringIO(text), dialect)
    header = [h.strip() for h in next(reader)]
    return header, [r for r in reader if any(c.strip() for c in r)]


def _num(v):
    if v in (None, ""):
        return 0.0
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip().replace(" ", "")
    if "," in s and "." not in s:      # "1234,5" (CSV regional)
        s = s.replace(",", ".")
    return float(s.replace(",", ""))


def _bool(v):
    return str(v).strip().lower() in ("true", "1", "yes", "si", "sí")


def slug(s):
    return re.sub(r"[^A-Za-z0-9._-]+", "_", str(s or "unknown")).strip("_") or "unknown"


def parse_rvtools(data: bytes, filename: str, cfg):
    """Normaliza el vInfo. Devuelve {vcenter: [vm, ...]} y un resumen."""
    header, rows = _read_table(data, filename)
    idx = {}
    for key, names in COLS.items():
        for n in names:
            if n in header:
                idx[key] = header.index(n)
                break
    missing = [COLS[k][0] for k in REQUIRED if k not in idx]
    if missing:
        raise ValueError("El archivo no parece un vInfo de RVTools; faltan columnas: " + ", ".join(missing))

    # Atributos personalizados / tags: columnas entre "VI SDK API Version" y "VI SDK Server".
    tag_cols = []
    if "VI SDK API Version" in header and "VI SDK Server" in header:
        a, b = header.index("VI SDK API Version"), header.index("VI SDK Server")
        tag_cols = [(i, header[i]) for i in range(a + 1, b) if header[i]]
    tcfg = cfg["tags"]
    if tcfg.get("include"):
        tag_cols = [(i, n) for i, n in tag_cols if n in tcfg["include"]]
    tag_cols = [(i, n) for i, n in tag_cols
                if not any(n.startswith(p) for p in tcfg.get("exclude_prefixes", []))]

    def g(r, key, default=""):
        i = idx.get(key)
        v = r[i] if i is not None and i < len(r) else None
        return default if v is None else v

    by_vc, stats = {}, {"rows": 0, "templates": 0, "srm": 0}
    for r in rows:
        name = str(g(r, "vm")).strip()
        if not name:
            continue
        stats["rows"] += 1
        tags = {}
        for i, n in tag_cols:
            v = r[i] if i < len(r) else None
            if v not in (None, ""):
                tags[n] = str(v).strip()
        vm = {
            "name": name,
            "power": str(g(r, "power")).strip() or "poweredOn",
            "template": _bool(g(r, "template", False)),
            "srm": _bool(g(r, "srm", False)),
            "vcpu": int(_num(g(r, "cpus"))),
            "memory_gb": round(_num(g(r, "memory_mib")) / 1024, 4),
            "prov_gb": round(_num(g(r, "prov_mib")) / 1024, 4),
            "inuse_gb": round(_num(g(r, "inuse_mib")) / 1024, 4),
            "datacenter": str(g(r, "datacenter")).strip() or "Sin datacenter",
            "cluster": str(g(r, "cluster")).strip() or "Sin cluster",
            "host": str(g(r, "host")).strip(),
            "os": str(g(r, "os_tools") or g(r, "os_cfg")).strip(),
            "moref": str(g(r, "moref")).strip(),
            "uuid": str(g(r, "uuid")).strip(),
            "folder": str(g(r, "folder")).strip(),
            "vcenter": str(g(r, "vcenter")).strip(),
            "vcenter_uuid": str(g(r, "vcenter_uuid")).strip(),
            "tags": tags,
        }
        stats["templates"] += vm["template"]
        stats["srm"] += vm["srm"]
        by_vc.setdefault(vm["vcenter"] or "vcenter", []).append(vm)
    stats["vcenters"] = {k: len(v) for k, v in by_vc.items()}
    stats["tag_columns"] = [n for _, n in tag_cols]
    return by_vc, stats


# =============================================================================
# 2) ALMACÉN DE ESTADO (carpeta local o S3)
# =============================================================================
class LocalStore:
    def __init__(self, root):
        self.root = root

    def put(self, key, data: bytes, content_type=None):
        p = os.path.join(self.root, *key.split("/"))
        os.makedirs(os.path.dirname(p), exist_ok=True)
        with open(p, "wb") as f:
            f.write(data)

    def get(self, key):
        p = os.path.join(self.root, *key.split("/"))
        if not os.path.exists(p):
            return None
        with open(p, "rb") as f:
            return f.read()

    def list(self, prefix):
        base = os.path.join(self.root, *prefix.rstrip("/").split("/"))
        out = []
        for dirpath, _, files in os.walk(base):
            for fn in files:
                out.append(os.path.relpath(os.path.join(dirpath, fn), self.root).replace(os.sep, "/"))
        return sorted(out)


class S3Store:
    def __init__(self, bucket, client=None):
        import boto3
        self.bucket, self.s3 = bucket, client or boto3.client("s3")

    def put(self, key, data: bytes, content_type="application/octet-stream"):
        self.s3.put_object(Bucket=self.bucket, Key=key, Body=data, ContentType=content_type)

    def get(self, key):
        try:
            return self.s3.get_object(Bucket=self.bucket, Key=key)["Body"].read()
        except self.s3.exceptions.NoSuchKey:
            return None

    def list(self, prefix):
        out, token = [], None
        while True:
            kw = {"Bucket": self.bucket, "Prefix": prefix}
            if token:
                kw["ContinuationToken"] = token
            page = self.s3.list_objects_v2(**kw)
            out += [o["Key"] for o in page.get("Contents", [])]
            if not page.get("IsTruncated"):
                return sorted(out)
            token = page["NextContinuationToken"]


def _k(prefix, rest):
    return f"{prefix.strip('/')}/{rest}" if prefix and prefix.strip("/") else rest


SNAP_RE = re.compile(r"snapshots/([^/]+)/(\d{4}-\d{2}-\d{2})\.json\.gz$")


def save_snapshots(store, state_prefix, by_vc, snap_date: date, source, backfilled=False):
    keys = []
    for vc, vms in by_vc.items():
        doc = {"vcenter": vc, "snapshot_date": snap_date.isoformat(), "source": source,
               "backfilled": backfilled, "vms": vms}
        key = _k(state_prefix, f"snapshots/{slug(vc)}/{snap_date.isoformat()}.json.gz")
        store.put(key, gzip.compress(json.dumps(doc, ensure_ascii=False).encode(), mtime=0),
                  "application/gzip")
        keys.append(key)
    return keys


def snapshot_index(store, state_prefix):
    """{vcenter_slug: {fecha: key}}"""
    idx = {}
    for key in store.list(_k(state_prefix, "snapshots/")):
        m = SNAP_RE.search(key)
        if m:
            idx.setdefault(m.group(1), {})[date.fromisoformat(m.group(2))] = key
    return idx


def backfill(store, state_prefix, start: date, end: date):
    """Copia la primera foto disponible hacia atrás (carga inicial). Marca backfilled=True."""
    created = []
    for vc, snaps in snapshot_index(store, state_prefix).items():
        first = min(snaps)
        doc = json.loads(gzip.decompress(store.get(snaps[first])))
        d = start
        while d <= min(end, first - timedelta(days=1)):
            if d not in snaps:
                created += save_snapshots(store, state_prefix, {doc["vcenter"]: doc["vms"]}, d,
                                          f"backfill de {first.isoformat()}", backfilled=True)
            d += timedelta(days=1)
    return created


# =============================================================================
# 3) TRANSFORMACIÓN A FOCUS
# =============================================================================
def _iso(d):
    return f"{d:%Y-%m-%d}T00:00:00Z"


def month_bounds(d: date):
    start = d.replace(day=1)
    return start, (start + timedelta(days=32)).replace(day=1)


def _f(v, n):
    return f"{v:.{n}f}"


def rows_for_day(day: date, vms, cfg, snap_date: date):
    bp_s, bp_e = month_bounds(day)
    hours = calendar.monthrange(day.year, day.month)[1] * 24
    rates = cfg["rates_monthly"]
    for vm in vms:
        if vm["srm"] and cfg["srm_placeholders"] == "exclude":
            continue
        if vm["template"] and cfg["templates"] == "exclude":
            continue
        storage_only = ((vm["srm"] and cfg["srm_placeholders"] == "storage_only")
                        or (vm["template"] and cfg["templates"] == "storage_only"))
        factor = 0.0 if storage_only else float(cfg["compute_factor"].get(vm["power"], 1.0))
        storage_gb = vm["inuse_gb"] if cfg["storage_basis"] == "in_use" else vm["prov_gb"]

        lines = []
        if factor > 0:
            lines.append(("vcpu", vm["vcpu"] * 24 * factor, "vCPU-Hours", "Compute",
                          "VMware vSphere VM", f"vCPU de la VM {vm['name']}"))
            lines.append(("memory_gb", vm["memory_gb"] * 24 * factor, "GB-Hours", "Compute",
                          "VMware vSphere VM", f"Memoria de la VM {vm['name']}"))
        lines.append(("storage_gb", storage_gb * 24, "GB-Hours", "Storage",
                      "VMware vSphere Storage", f"Almacenamiento de la VM {vm['name']}"))

        tags = dict(vm["tags"])
        if cfg["tags"].get("add_inventory_tags", True):
            tags.update({"vcenter": vm["vcenter"], "cluster": vm["cluster"],
                         "power_state": vm["power"]})
        tags_json = json.dumps(tags, ensure_ascii=False, separators=(",", ":"))
        vc_id = vm["vcenter_uuid"] or vm["vcenter"]
        resource_id = vm["uuid"] or f"{vm['vcenter']}/{vm['moref'] or vm['name']}"

        for key, qty, unit, category, service, desc in lines:
            if qty <= 0:
                continue
            unit_price = rates[key] / hours
            cost = qty * unit_price
            yield {
                "BilledCost": _f(cost, 6), "BillingAccountId": vc_id,
                "BillingAccountName": vm["vcenter"], "BillingCurrency": cfg["currency"],
                "BillingPeriodEnd": _iso(bp_e), "BillingPeriodStart": _iso(bp_s),
                "ChargeCategory": "Usage", "ChargeClass": "", "ChargeDescription": desc,
                "ChargeFrequency": "Usage-Based",
                "ChargePeriodEnd": _iso(day + timedelta(days=1)), "ChargePeriodStart": _iso(day),
                "ConsumedQuantity": _f(qty, 4), "ConsumedUnit": unit,
                "ContractedCost": _f(cost, 6), "ContractedUnitPrice": _f(unit_price, 8),
                "EffectiveCost": _f(cost, 6), "InvoiceIssuerName": cfg["invoice_issuer"],
                "ListCost": _f(cost, 6), "ListUnitPrice": _f(unit_price, 8),
                "PricingQuantity": _f(qty, 4), "PricingUnit": unit,
                "ProviderName": cfg["provider_name"], "PublisherName": cfg["publisher_name"],
                "RegionId": vm["datacenter"], "RegionName": vm["datacenter"],
                "ResourceId": resource_id, "ResourceName": vm["name"],
                "ResourceType": "VM Template" if vm["template"] else "Virtual Machine",
                "ServiceCategory": category, "ServiceName": service,
                "SubAccountId": f"{vm['vcenter']}/{vm['cluster']}", "SubAccountName": vm["cluster"],
                "Tags": tags_json,
                "x_vCenter": vm["vcenter"], "x_VMwareMoRef": vm["moref"],
                "x_PowerState": vm["power"], "x_IsTemplate": str(vm["template"]).lower(),
                "x_VCPUCount": str(vm["vcpu"]), "x_MemoryGB": _f(vm["memory_gb"], 2),
                "x_StorageGB": _f(storage_gb, 2), "x_GuestOS": vm["os"], "x_Host": vm["host"],
                "x_SnapshotDate": snap_date.isoformat(),
            }


def validate_row(n, row):
    errs = [f"fila {n}: {c} vacío" for c in NOT_NULL if row.get(c) in ("", None)]
    for c in ("BilledCost", "EffectiveCost", "ListCost", "ContractedCost", "PricingQuantity"):
        try:
            if float(row[c]) < 0:
                errs.append(f"fila {n}: {c} negativo")
        except ValueError:
            errs.append(f"fila {n}: {c} no numérico")
    return errs


def build_month(store, state_prefix, month_start: date, through: date, cfg,
                root_dir, sub_dir, out_store, epoch=None):
    """Genera el drop del mes con los días month_start..through. Devuelve un resumen."""
    bp_s, bp_e = month_bounds(month_start)
    through = min(through, bp_e - timedelta(days=1))
    max_gap = int(cfg["max_carry_forward_days"])
    index = snapshot_index(store, state_prefix)
    cache = {}

    def load(key):
        if key not in cache:
            cache[key] = json.loads(gzip.decompress(store.get(key)))["vms"]
        return cache[key]

    buf = io.BytesIO()
    gz = gzip.GzipFile(fileobj=buf, mode="wb", mtime=0, filename="")
    text = io.TextIOWrapper(gz, encoding="utf-8", newline="")   # UTF-8 sin BOM
    writer = csv.DictWriter(text, fieldnames=FOCUS_COLUMNS + CUSTOM_COLUMNS, lineterminator="\n")
    writer.writeheader()

    n, total, errors = 1, 0.0, []
    coverage = {}   # vcenter -> {"days": n, "carried": n, "missing": [fechas]}
    d = bp_s
    while d <= through:
        for vc, snaps in index.items():
            cov = coverage.setdefault(vc, {"days": 0, "carried": 0, "missing": []})
            past = [s for s in snaps if s <= d]
            if not past or (d - max(past)).days > max_gap:
                cov["missing"].append(d.isoformat())
                continue
            snap = max(past)
            cov["days"] += 1
            cov["carried"] += snap != d
            for row in rows_for_day(d, load(snaps[snap]), cfg, snap):
                n += 1
                if len(errors) < 20:
                    errors += validate_row(n, row)
                total += float(row["BilledCost"])
                writer.writerow(row)
        cache.clear()   # libera memoria día a día (las fotos se recargan si hacen falta)
        d += timedelta(days=1)
    text.flush()
    text.detach()
    gz.close()

    rows = n - 1
    if errors:
        raise ValueError("Validación FOCUS falló:\n  " + "\n  ".join(errors[:20]))
    if rows == 0:
        return {"month": f"{bp_s:%Y-%m}", "rows": 0, "note": "sin fotos para este mes"}

    epoch = epoch or int(datetime.now(timezone.utc).timestamp())
    rel = "/".join(p for p in (sub_dir, cfg["vendor"], f"{bp_s:%Y%m%d}-{bp_e:%Y%m%d}", str(epoch)) if p)
    data_key = f"{rel}/{cfg['file_prefix']}-00001.csv.gz"
    manifest = {
        "report_id": str(uuid.UUID(hashlib.md5(data_key.encode()).hexdigest())),
        "root_dir": root_dir, "content_type": "CSV", "compression": "GZIP",
        "all_report_keys": [data_key],
        "updated_at": datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ"),
        "focus_version": FOCUS_VERSION,
    }
    out_store.put(data_key, buf.getvalue(), "application/gzip")      # datos primero,
    out_store.put(f"{rel}/{cfg['file_prefix']}-Manifest.json",          # manifest al final
                  json.dumps(manifest, indent=2).encode(), "application/json")
    return {"month": f"{bp_s:%Y-%m}", "through": through.isoformat(), "rows": rows,
            "total": round(total, 2), "currency": cfg["currency"], "path": rel,
            "coverage": {k: {**v, "missing": v["missing"][:10]} for k, v in coverage.items()}}


DATE_IN_NAME = re.compile(r"(20\d{2})[-_]?(\d{2})[-_]?(\d{2})")


def date_from_name(name):
    m = DATE_IN_NAME.search(name)
    if m:
        try:
            return date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
        except ValueError:
            return None
    return None


# =============================================================================
# CLI local
# =============================================================================
def main():
    ap = argparse.ArgumentParser(description="RVTools vInfo -> FOCUS 1.0 para Cloudability")
    sub = ap.add_subparsers(dest="cmd", required=True)
    for name in ("ingest", "build", "backfill"):
        p = sub.add_parser(name)
        p.add_argument("--state", default="state", help="carpeta de estado (fotos)")
        p.add_argument("--config", help="config.json (tarifario y políticas)")
    sub.choices["ingest"].add_argument("files", nargs="+")
    sub.choices["ingest"].add_argument("--date", help="fecha de la foto (UTC); por defecto del nombre o hoy")
    b = sub.choices["build"]
    b.add_argument("--out", default="out")
    b.add_argument("--root-dir", required=True, help="bucket (Root Directory en Cloudability)")
    b.add_argument("--sub-dir", default="vmware")
    b.add_argument("--month", help="AAAA-MM; por defecto el mes actual")
    b.add_argument("--through", help="último día AAAA-MM-DD; por defecto hoy o fin de mes")
    bf = sub.choices["backfill"]
    bf.add_argument("--from", dest="start", required=True)
    bf.add_argument("--to", dest="end")
    a = ap.parse_args()

    cfg = merged_config(json.load(open(a.config)) if a.config else None)
    store = LocalStore(a.state)
    if a.cmd == "ingest":
        for f in a.files:
            d = date.fromisoformat(a.date) if a.date else (
                date_from_name(os.path.basename(f)) or datetime.now(timezone.utc).date())
            by_vc, stats = parse_rvtools(open(f, "rb").read(), f, cfg)
            save_snapshots(store, "", by_vc, d, os.path.basename(f))
            print(f"{os.path.basename(f)} -> foto {d}: {stats}")
    elif a.cmd == "backfill":
        end = date.fromisoformat(a.end) if a.end else date.max
        print(f"{len(backfill(store, '', date.fromisoformat(a.start), end))} fotos de backfill creadas")
    else:
        today = datetime.now(timezone.utc).date()
        m = date.fromisoformat(a.month + "-01") if a.month else today.replace(day=1)
        through = date.fromisoformat(a.through) if a.through else today
        print(json.dumps(build_month(store, "", m, through, cfg, a.root_dir, a.sub_dir,
                                     LocalStore(a.out)), indent=2, ensure_ascii=False))


if __name__ == "__main__":
    sys.exit(main())
