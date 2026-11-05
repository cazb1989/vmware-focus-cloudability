"""
Pruebas sin AWS ni vCenter reales (S3 simulado con moto, export de RVTools sintético).
  pip install -r requirements-dev.txt && python -m pytest -q
"""
import csv, gzip, io, json, sys
from datetime import date
from pathlib import Path

import boto3
import openpyxl
import pytest
from moto import mock_aws

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "lambda"))
import rvtools_focus as rf  # noqa: E402

BUCKET = "test-vmware-focus"
HEADER = ["VM", "Powerstate", "Template", "SRM Placeholder", "CPUs", "Memory", "Provisioned MiB",
          "In Use MiB", "Datacenter", "Cluster", "Host", "OS according to the configuration file",
          "VM ID", "VM UUID", "VI SDK API Version", "Ambiente", "FW_PA_DMZ", "VI SDK Server", "VI SDK UUID"]
VMS = [  # nombre, estado, template, srm, vCPU, MiB RAM, MiB disco
    ("web-01", "poweredOn", "False", "False", 4, 8192, 102400),
    ("db-01", "poweredOff", "False", "False", 8, 16384, 204800),
    ("tpl-rhel", "poweredOff", "True", "False", 2, 4096, 51200),
    ("dr-web-01", "poweredOff", "False", "True", 4, 8192, 102400),
]


def make_xlsx(vcenter="vc01.lab.local"):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "vInfo"
    ws.append(HEADER)
    for i, (n, p, t, s, c, m, d) in enumerate(VMS):
        ws.append([n, p, t, s, c, m, d, d // 2, "DC1", "Prod", "esx01", "RHEL 9", f"vm-{i}",
                   f"5000-{i}", "8.0.3.0", "PRODUCCION", "grupo-fw", vcenter, "VC-UUID-1"])
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def test_parse_y_reglas_de_costo():
    cfg = rf.merged_config()
    by_vc, stats = rf.parse_rvtools(make_xlsx(), "x.xlsx", cfg)
    assert stats["templates"] == 1 and stats["srm"] == 1
    assert stats["tag_columns"] == ["Ambiente"]                       # FW_ excluido
    rows = list(rf.rows_for_day(date(2026, 10, 15), by_vc["vc01.lab.local"], cfg, date(2026, 10, 15)))
    names = {(r["ResourceName"], r["PricingUnit"]) for r in rows}
    assert ("dr-web-01", "GB-Hours") not in names                     # SRM placeholder excluido
    assert ("tpl-rhel", "vCPU-Hours") not in names                    # template: solo disco
    assert ("db-01", "vCPU-Hours") not in names                       # apagada: solo disco
    vcpu = next(r for r in rows if r["ResourceName"] == "web-01" and r["PricingUnit"] == "vCPU-Hours")
    assert vcpu["PricingQuantity"] == "96.0000"                       # 4 vCPU x 24 h
    assert vcpu["BilledCost"] == "1.935484"                           # 96 x 15/744 (guía, 5.3)
    assert vcpu["BillingPeriodStart"] == "2026-10-01T00:00:00Z"
    assert json.loads(vcpu["Tags"])["Ambiente"] == "PRODUCCION"
    assert all(not rf.validate_row(i, r) for i, r in enumerate(rows))


def test_carry_forward_y_backfill(tmp_path):
    cfg = rf.merged_config({"max_carry_forward_days": 2})
    st, out = rf.LocalStore(tmp_path / "s"), rf.LocalStore(tmp_path / "o")
    by_vc, _ = rf.parse_rvtools(make_xlsx(), "x.xlsx", cfg)
    rf.save_snapshots(st, "", by_vc, date(2026, 11, 3), "x")
    r = rf.build_month(st, "", date(2026, 11, 1), date(2026, 11, 7), cfg, BUCKET, "vmware", out)
    cov = r["coverage"]["vc01.lab.local"]
    assert cov["days"] == 3 and cov["carried"] == 2                   # 3 real, 4-5 rellenados
    assert cov["missing"] == ["2026-11-01", "2026-11-02", "2026-11-06", "2026-11-07"]
    rf.backfill(st, "", date(2026, 11, 1), date(2026, 11, 7))
    r = rf.build_month(st, "", date(2026, 11, 1), date(2026, 11, 5), cfg, BUCKET, "vmware", out)
    assert r["coverage"]["vc01.lab.local"]["days"] == 5


def test_varios_vcenter_en_un_drop(tmp_path):
    cfg = rf.merged_config()
    st = rf.LocalStore(tmp_path / "s")
    for vc in ("vc01.lab.local", "vc02.lab.local"):
        rf.save_snapshots(st, "", rf.parse_rvtools(make_xlsx(vc), "x.xlsx", cfg)[0], date(2026, 11, 1), "x")
    r = rf.build_month(st, "", date(2026, 11, 1), date(2026, 11, 1), cfg, BUCKET, "vmware",
                       rf.LocalStore(tmp_path / "o"))
    assert set(r["coverage"]) == {"vc01.lab.local", "vc02.lab.local"}


def test_lambda_s3_de_punta_a_punta(monkeypatch):
    monkeypatch.setenv("AWS_DEFAULT_REGION", "us-east-1")
    monkeypatch.setenv("AWS_ACCESS_KEY_ID", "x")
    monkeypatch.setenv("AWS_SECRET_ACCESS_KEY", "x")
    monkeypatch.setenv("BUCKET", BUCKET)
    with mock_aws():
        s3 = boto3.client("s3")
        s3.create_bucket(Bucket=BUCKET)
        s3.put_object(Bucket=BUCKET, Key="vmware-state/config.json",
                      Body=json.dumps({"currency": "MXN", "rates_monthly": {"vcpu": 300}}))
        key = "vmware-inbox/vc01/2026-11-05/vInfo_vc01_2026-11-05.xlsx"
        s3.put_object(Bucket=BUCKET, Key=key, Body=make_xlsx())
        import importlib, handler
        importlib.reload(handler)
        ev = {"source": "aws.s3", "time": "2026-11-05T22:00:00Z",
              "detail": {"bucket": {"name": BUCKET}, "object": {"key": key}}}
        res = handler.lambda_handler({**ev, "today": "2026-11-05"}, None)
        b = res["builds"][0]
        assert res["snapshot_date"] == "2026-11-05" and b["currency"] == "MXN"
        keys = [o["Key"] for o in s3.list_objects_v2(Bucket=BUCKET)["Contents"]]
        assert key not in keys and any(k.startswith("vmware-state/processed/") for k in keys)
        man_key = next(k for k in keys if k.endswith("vmware-focus-Manifest.json"))
        assert man_key.startswith("vmware/VMware/20261101-20261201/")
        man = json.loads(s3.get_object(Bucket=BUCKET, Key=man_key)["Body"].read())
        assert man["focus_version"] == "1.0" and man["root_dir"] == BUCKET
        data = gzip.decompress(s3.get_object(Bucket=BUCKET, Key=man["all_report_keys"][0])["Body"].read())
        assert not data.startswith(b"\xef\xbb\xbf")                    # UTF-8 sin BOM
        rows = list(csv.DictReader(io.StringIO(data.decode())))
        assert len(rows) == b["rows"] and {r["BillingCurrency"] for r in rows} == {"MXN"}
        # reintento del mismo evento: no falla ni duplica fotos
        assert handler.lambda_handler({**ev, "today": "2026-11-05"}, None)["stats"] == {"note": "ya procesado"}


def test_rechaza_archivo_que_no_es_vinfo():
    with pytest.raises(ValueError):
        rf.parse_rvtools(b"a,b\n1,2\n", "otro.csv", rf.merged_config())
