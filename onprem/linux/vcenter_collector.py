#!/usr/bin/env python3
"""
vcenter_collector.py - Inventario de VMs desde vCenter (pyVmomi) en formato vInfo de RVTools.

Genera un CSV con los mismos nombres de columna que la pestaña vInfo de RVTools, incluidos
los atributos personalizados (y opcionalmente las etiquetas de vSphere) como columnas. Así la
Lambda procesa igual un export de RVTools o la salida de este colector.

Usa un solo PropertyCollector (rápido aun con miles de VMs) y un usuario de solo lectura.

Uso:
  export VCENTER_PASSWORD='...'
  python3 vcenter_collector.py --host vcenter.empresa.com --user cloudability-ro@vsphere.local \
      --ca-bundle /etc/vmware-focus/vcenter-ca.pem --out /var/lib/vmware-focus/vinfo.csv

Requiere: pip install pyvmomi   (y requests si se usa --include-tags)
"""
import argparse
import csv
import os
import ssl
import sys

from pyVim.connect import Disconnect, SmartConnect
from pyVmomi import vim, vmodl

MIB = 1024 * 1024
VM_PROPS = [
    "name", "config.template", "config.hardware.numCPU", "config.hardware.memoryMB",
    "config.instanceUuid", "config.guestFullName", "config.managedBy", "guest.guestFullName",
    "runtime.powerState", "runtime.host", "summary.storage.committed",
    "summary.storage.uncommitted", "customValue", "parent",
]
BASE_COLUMNS = [
    "VM", "Powerstate", "Template", "SRM Placeholder", "CPUs", "Memory",
    "Provisioned MiB", "In Use MiB", "Folder", "Datacenter", "Cluster", "Host",
    "OS according to the configuration file", "OS according to the VMware Tools",
    "VM ID", "VM UUID", "VI SDK Server type", "VI SDK API Version",
]
TAIL_COLUMNS = ["VI SDK Server", "VI SDK UUID"]   # los atributos van entre API Version y Server


def connect(host, user, password, ca_bundle, port=443):
    if ca_bundle and ca_bundle.lower() == "false":           # solo para pruebas
        ctx = ssl._create_unverified_context()
    else:
        ctx = ssl.create_default_context(cafile=ca_bundle or None)
    return SmartConnect(host=host, user=user, pwd=password, port=port, sslContext=ctx)


def retrieve(content, obj_type, props):
    """Lee props de todos los objetos de un tipo en una sola llamada (con paginación)."""
    view = content.viewManager.CreateContainerView(content.rootFolder, [obj_type], True)
    try:
        spec = vmodl.query.PropertyCollector.FilterSpec(
            objectSet=[vmodl.query.PropertyCollector.ObjectSpec(
                obj=view, skip=True,
                selectSet=[vmodl.query.PropertyCollector.TraversalSpec(
                    name="t", path="view", skip=False, type=vim.view.ContainerView)])],
            propSet=[vmodl.query.PropertyCollector.PropertySpec(type=obj_type, pathSet=props)])
        pc, out = content.propertyCollector, {}
        res = pc.RetrievePropertiesEx([spec], vmodl.query.PropertyCollector.RetrieveOptions(maxObjects=500))
        while res:
            for o in res.objects:
                out[o.obj] = {p.name: p.val for p in (o.propSet or [])}
            res = pc.ContinueRetrievePropertiesEx(res.token) if res.token else None
        return out
    finally:
        view.Destroy()


def datacenter_of(obj):
    while obj is not None and not isinstance(obj, vim.Datacenter):
        obj = getattr(obj, "parent", None)
    return obj.name if obj is not None else ""


def folder_path(obj):
    """Ruta de carpetas de la VM dentro de 'vm' (p. ej. /Produccion/Web)."""
    parts = []
    while isinstance(obj, vim.Folder) and not isinstance(obj.parent, vim.Datacenter):
        parts.append(obj.name)
        obj = obj.parent
    return "/" + "/".join(reversed(parts)) if parts else ""


def collect_tags(host, user, password, ca_bundle, morefs):
    """Etiquetas de vSphere (API REST de tagging). Devuelve {moref: {categoria: valor}}."""
    import requests
    base, http = f"https://{host}/api", requests.Session()
    http.verify = False if (ca_bundle or "").lower() == "false" else (ca_bundle or True)
    r = http.post(base + "/session", auth=(user, password), timeout=60)
    r.raise_for_status()
    http.headers["vmware-api-session-id"] = r.json()
    result, names, cats = {}, {}, {}
    try:
        for i in range(0, len(morefs), 500):
            body = {"object_ids": [{"type": "VirtualMachine", "id": m} for m in morefs[i:i + 500]]}
            r = http.post(base + "/cis/tagging/tag-association",
                          params={"action": "list-attached-tags-on-objects"}, json=body, timeout=120)
            r.raise_for_status()
            for item in r.json():
                for tid in item.get("tag_ids", []):
                    if tid not in names:
                        t = http.get(f"{base}/cis/tagging/tag/{tid}", timeout=60).json()
                        if t["category_id"] not in cats:
                            cats[t["category_id"]] = http.get(
                                f"{base}/cis/tagging/category/{t['category_id']}", timeout=60).json()["name"]
                        names[tid] = (cats[t["category_id"]], t["name"])
                    k, v = names[tid]
                    d = result.setdefault(item["object_id"]["id"], {})
                    d[k] = v if k not in d else d[k] + ", " + v
    finally:
        http.delete(base + "/session", timeout=30)
    return result


def collect(si, host, args):
    content = si.RetrieveContent()
    about = content.about
    fields = {f.key: f.name for f in (content.customFieldsManager.field or [])
              if f.managedObjectType in (None, vim.VirtualMachine)} if content.customFieldsManager else {}

    hosts = retrieve(content, vim.HostSystem, ["name", "parent"])
    host_info = {}
    for h, p in hosts.items():
        parent = p.get("parent")
        cluster = parent.name if isinstance(parent, vim.ClusterComputeResource) else ""
        host_info[h] = (p.get("name", ""), cluster, datacenter_of(parent))

    vms = retrieve(content, vim.VirtualMachine, VM_PROPS)
    tags = {}
    if args.include_tags:
        try:
            tags = collect_tags(host, args.user, os.environ["VCENTER_PASSWORD"], args.ca_bundle,
                                [vm._moId for vm in vms])
        except Exception as exc:  # noqa: BLE001
            print(f"AVISO: no se pudieron leer las etiquetas ({exc}); se continúa sin ellas.",
                  file=sys.stderr)

    rows, attr_names = [], set()
    for vm, p in vms.items():
        h = p.get("runtime.host")
        hname, cluster, dc = host_info.get(h, ("", "", ""))
        if not dc:
            dc = datacenter_of(p.get("parent"))
        managed = p.get("config.managedBy")
        srm = bool(managed and managed.extensionKey == "com.vmware.vcDr"
                   and getattr(managed, "type", "") == "placeholderVm")
        committed = p.get("summary.storage.committed") or 0
        uncommitted = p.get("summary.storage.uncommitted") or 0
        row = {
            "VM": p.get("name", ""),
            "Powerstate": str(p.get("runtime.powerState", "")),
            "Template": str(bool(p.get("config.template"))),
            "SRM Placeholder": str(srm),
            "CPUs": p.get("config.hardware.numCPU") or 0,
            "Memory": p.get("config.hardware.memoryMB") or 0,
            "Provisioned MiB": round((committed + uncommitted) / MIB),
            "In Use MiB": round(committed / MIB),
            "Folder": folder_path(p.get("parent")),
            "Datacenter": dc, "Cluster": cluster, "Host": hname,
            "OS according to the configuration file": p.get("config.guestFullName") or "",
            "OS according to the VMware Tools": p.get("guest.guestFullName") or "",
            "VM ID": vm._moId, "VM UUID": p.get("config.instanceUuid") or "",
            "VI SDK Server type": about.fullName, "VI SDK API Version": about.apiVersion,
            "VI SDK Server": host, "VI SDK UUID": about.instanceUuid,
        }
        for cv in p.get("customValue") or []:
            name = fields.get(cv.key)
            if name and getattr(cv, "value", ""):
                row[name] = cv.value
                attr_names.add(name)
        for k, v in tags.get(vm._moId, {}).items():
            if k not in row or not row[k]:
                row[k] = v
                attr_names.add(k)
        rows.append(row)
    columns = BASE_COLUMNS + sorted(attr_names) + TAIL_COLUMNS
    return columns, sorted(rows, key=lambda r: r["VM"].lower())


def main():
    ap = argparse.ArgumentParser(description="vCenter -> CSV con formato vInfo de RVTools")
    ap.add_argument("--host", default=os.environ.get("VCENTER_HOST"), required=not os.environ.get("VCENTER_HOST"))
    ap.add_argument("--user", default=os.environ.get("VCENTER_USER"), required=not os.environ.get("VCENTER_USER"))
    ap.add_argument("--port", type=int, default=int(os.environ.get("VCENTER_PORT", "443")))
    ap.add_argument("--ca-bundle", default=os.environ.get("VCENTER_CA_BUNDLE", ""),
                    help="PEM con la CA de vCenter ('false' desactiva la validación TLS: solo pruebas)")
    ap.add_argument("--include-tags", action="store_true", help="leer también etiquetas de vSphere")
    ap.add_argument("--out", required=True)
    args = ap.parse_args()
    if "VCENTER_PASSWORD" not in os.environ:
        sys.exit("ERROR: defina VCENTER_PASSWORD.")

    si = connect(args.host, args.user, os.environ["VCENTER_PASSWORD"], args.ca_bundle, args.port)
    try:
        columns, rows = collect(si, args.host, args)
    finally:
        Disconnect(si)
    os.makedirs(os.path.dirname(os.path.abspath(args.out)), exist_ok=True)
    tmp = args.out + ".tmp"
    with open(tmp, "w", newline="", encoding="utf-8") as f:     # UTF-8 sin BOM, coma, punto decimal
        w = csv.DictWriter(f, fieldnames=columns, extrasaction="ignore")
        w.writeheader()
        w.writerows(rows)
    os.replace(tmp, args.out)
    print(f"{args.host}: {len(rows)} VMs -> {args.out}")


if __name__ == "__main__":
    main()
