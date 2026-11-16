# VMware vSphere → FOCUS 1.0 → IBM Cloudability

Esta solución lleva el costo de la **nube privada VMware** a **IBM Cloudability** con el estándar
[FOCUS](https://focus.finops.org/), de forma diaria y automática. 

VMware no factura por uso. El costo se calcula con el **tarifario interno** (por vCPU, GB de RAM
y GB de disco) aplicado al inventario de cada día, como describe la *Guía de transformación de
datos VMware a FOCUS*.

```
  Centro de datos del cliente                         Cuenta AWS                          IBM Cloudability
┌─────────────────────────────────┐     HTTPS     ┌───────────────────────────────────┐
│ Linux + cron (recomendado)      │    (solo      │ S3  vmware-inbox/  ──evento──┐    │
│   vcenter_collector.py (pyVmomi)├───subida)───► │                              ▼    │
│   ──► vInfo.csv                 │               │   Lambda vmware-rvtools-focus     │   ┌──────────────────┐
│  o bien                         │               │   (fotos diarias por vCenter)     │   │   Custom Data    │
│ Windows + RVTools (alternativa) │               │              │                    │   │ (FOCUS Ingress)  │
│   ──► vInfo.xlsx                │               │              ▼                    │   └────────┬─────────┘
└─────────────────────────────────┘               │ S3  vmware/VMware/<mes>/<epoch>/  │◄───────────┘
                                                  │   + EventBridge diario 23:30 UTC  │  lee 1 vez al día
                                                  └───────────────────────────────────┘  (rol de solo lectura)
```

**Características:**
- **Sin acceso entrante a la red del cliente.** El servidor on-prem solo **sube** archivos a S3, con
  un usuario IAM que únicamente puede escribir en el inbox. AWS nunca se conecta a vCenter.
- **Acepta dos formatos de entrada.** Lee igual la salida del colector Linux (CSV) y un export de
  **RVTools** (vInfo, `.xlsx` o `.csv`). También sirve para cargas manuales: se sube el archivo y listo.
- **Varios vCenter en un solo drop.** Cloudability toma solo la carpeta epoch más reciente de cada mes,
  así que todos los vCenter tienen que ir juntos.
- **Tolera días perdidos.** Si un día falla la recolección, usa la última foto disponible, hasta 3 días
  (configurable). La columna `x_SnapshotDate` indica qué foto se usó.
- **Tarifario y políticas en un JSON del bucket.** Se cambian sin volver a desplegar.
- **Reglas de costo según la guía**, con dos mejoras para datos reales: los **SRM placeholders** (réplicas
  de DR) se excluyen para no contar dos veces, y los **templates** solo pagan almacenamiento.

---

## Contenido

| Ruta | Qué es |
|---|---|
| `lambda/rvtools_focus.py` | Motor: lee vInfo (RVTools o colector), guarda fotos y genera el drop FOCUS 1.0. También es una CLI local |
| `lambda/handler.py` | Lambda: archivo nuevo en el inbox, ejecución diaria, reconstrucción y backfill |
| `template.yaml` | CloudFormation: Lambda, reglas de EventBridge, rol con permisos mínimos, usuario de subida y alarma |
| `deploy.sh` | Despliegue en un comando (recomendado desde AWS CloudShell) |
| `config.example.json` | Tarifario y políticas |
| `onprem/linux/` | Colector pyVmomi + `run.sh` (cron) + `install.sh` + `env.example` |
| `onprem/windows/Export-RVToolsToS3.ps1` | Alternativa con RVTools y el Programador de tareas |
| `tests/` | Pruebas con S3 simulado y un vInfo sintético |

---

## Requisitos

| Dónde | Qué |
|---|---|
| vCenter 6.7+ | Usuario de **solo lectura** (p. ej. `cloudability-ro@vsphere.local`), rol *Read-only* en la raíz con *Propagar a objetos hijos* |
| On-prem | **Linux** (RHEL/Rocky/Alma 8-9, Ubuntu 22.04+) con Python 3.8+, red a vCenter:443 y salida HTTPS a S3. **O** Windows con RVTools 4.x y AWS CLI |
| AWS | Cuenta con permisos para S3, IAM, Lambda, EventBridge y CloudFormation |
| Cloudability | Rol de administrador en *Vendor Credentials* |
| Finanzas | **Tarifario mensual** por vCPU, GB de RAM y GB de disco, en la moneda acordada |

---

## Paso 1: Definir el tarifario y las políticas

Copia `config.example.json` a `config.json` y ajústalo con Finanzas:

| Clave | Default | Significado |
|---|---|---|
| `currency` | `USD` | Moneda ISO 4217 del tarifario (`MXN`, `PEN`, `COP`, `CLP`…) |
| `rates_monthly` | `vcpu 15`, `memory_gb 5`, `storage_gb 0.10` | Costo **mensual** por unidad. Se prorratea por las horas reales del mes. **Los valores de ejemplo son ilustrativos** |
| `compute_factor` | on 1.0 / off 0 / suspended 0 | Fracción del cómputo que se cobra según el estado. El disco se cobra siempre |
| `storage_basis` | `provisioned` | `provisioned` (aprovisionado) o `in_use` (usado, útil con thin provisioning) |
| `templates` | `storage_only` | `storage_only`, `exclude` o `charge` |
| `srm_placeholders` | `exclude` | `exclude` o `storage_only`. Son réplicas de SRM: cobrarlas duplica el costo de la VM protegida |
| `max_carry_forward_days` | `3` | Días sin foto que se rellenan con la última foto |
| `tags.include` | `[]` (todas) | Lista de atributos o etiquetas que van a `Tags`. Vacío = todas |
| `tags.exclude_prefixes` | `["FW_"]` | Atributos que se excluyen (p. ej. grupos de firewall) |
| `tags.add_inventory_tags` | `true` | Agrega `vcenter`, `cluster` y `power_state` a `Tags` |
| `invoice_issuer` | `TI Corporativa` | Área interna que "factura" la nube privada |

## Paso 2: Desplegar en AWS

En **AWS CloudShell** (ícono `>_` en la consola, región `us-east-1`), sube el repo o clónalo y ejecuta:

```bash
cd vmware-focus-cloudability
cp config.example.json config.json   # y edítalo (nano config.json)
BUCKET=<bucket-unico> REGION=us-east-1 ALERT_EMAIL=<email-opcional> ./deploy.sh
```

El script hace todo esto:
- Crea el bucket privado y cifrado si no existe.
- Activa las notificaciones a EventBridge, **sin borrar las que ya tenga** el bucket.
- Sube `config.json` y el código con `openpyxl`.
- Despliega el stack `vmware-rvtools-focus`.
- Al final imprime los valores para Cloudability.

Crea la credencial del servidor on-prem, que **solo puede subir al inbox**:
```bash
aws iam create-access-key --user-name vmware-inventory-uploader
```

## Paso 3a (recomendado): Colector en Linux con cron

```bash
# en el servidor Linux, como root
cd vmware-focus-cloudability/onprem/linux
sudo ./install.sh                       # usuario de servicio, venv (pyvmomi, boto3), cron diario
sudo nano /etc/vmware-focus/env         # vCenter(s), usuario RO, contraseña, CA, bucket y access key
curl -sk https://<vcenter>/certs/download.zip -o ca.zip && unzip ca.zip   # CA de vCenter → /etc/vmware-focus/vcenter-ca.pem
sudo -u vmwarefocus /opt/vmware-focus/run.sh                              # prueba manual
```

- El cron corre **todos los días a las 22:00 UTC** (16:00 CDMX / 17:00 Lima), según `/etc/cron.d/vmware-focus`.
- Por cada vCenter genera `vinfo_<vcenter>_<fecha>.csv`, con las mismas columnas que el vInfo de RVTools,
  incluidos los **atributos personalizados**. Con `INCLUDE_TAGS=1` agrega también las **etiquetas de vSphere**.
- Sube cada archivo a `s3://<bucket>/vmware-inbox/<vcenter>/<fecha>/`. La Lambda se dispara sola.
- Log: `/var/log/vmware-focus/collector.log`. El archivo `env` (con la contraseña) queda con permisos `640`.

## Paso 3b (alternativa): RVTools en Windows

Si el cliente ya usa RVTools, programa `onprem/windows/Export-RVToolsToS3.ps1` en el Programador de
tareas. El script exporta vInfo de cada vCenter y lo sube al inbox con el AWS CLI.

**Carga manual**, por ejemplo un export puntual de RVTools:
```bash
aws s3 cp RVTools_tabvInfo.xlsx s3://<bucket>/vmware-inbox/<vcenter>/2026-11-05/RVTools_tabvInfo.xlsx
```
La fecha de la foto se toma de la ruta o del nombre del archivo (`AAAA-MM-DD`). Si no aparece en
ninguno, se usa la fecha de subida (UTC).

## Paso 4: Conectar Cloudability

Ve a **Settings → Vendor Credentials → Add Datasource → Other Data Sources → Add a Credential**:

| Campo | Valor |
|---|---|
| Storage | AWS Simple Storage Service |
| Vendor | `VMware` |
| AWS Account ID | cuenta del bucket |
| Root Directory | nombre del bucket |
| Sub Directories | `vmware` |
| Manifest Prefix | `vmware-focus` |

El *Sample path* debe quedar así: `<bucket>/vmware/VMware/<AAAAMMDD-AAAAMMDD>/<epoch>/vmware-focus-Manifest.json`.

Luego pulsa **Generate Template**, despliega ese template en CloudFormation (crea `CloudabilityCustomDataRole`)
y pulsa **Verify Credential**. Cloudability descarga los datos una vez al día, así que pueden tardar
**hasta 24 h** en aparecer.

## Paso 5: Carga inicial (opcional)

El inventario es una foto: no se puede reconstruir el pasado. Si el proceso arranca a mitad de mes y
se acepta suponer que el inventario no cambió, se puede **rellenar hacia atrás** con la primera foto.
Las fotos rellenadas quedan marcadas como `backfilled`.

```bash
aws lambda invoke --function-name vmware-rvtools-focus --cli-binary-format raw-in-base64-out \
  --payload '{"action":"backfill","from":"2026-11-01"}' out.json && cat out.json
```

---

## Cómo se calcula

Cada fila de FOCUS representa **una VM, un día y un componente**:

| Componente | Cuándo se cobra | PricingQuantity | ServiceName / ServiceCategory |
|---|---|---|---|
| vCPU | Encendida (× `compute_factor`) | vCPU × 24 (`vCPU-Hours`) | VMware vSphere VM / Compute |
| Memoria | Encendida (× `compute_factor`) | GB × 24 (`GB-Hours`) | VMware vSphere VM / Compute |
| Disco | Siempre (salvo SRM excluido) | GB × 24 (`GB-Hours`) | VMware vSphere Storage / Storage |

`tarifa_hora = tarifa_mensual / (días_del_mes × 24)`. Con esto, un mes completo cuesta exactamente la tarifa mensual.

| Origen (vInfo) | FOCUS |
|---|---|
| `VI SDK UUID` / `VI SDK Server` | BillingAccountId / BillingAccountName (el vCenter) |
| `Datacenter` | RegionId / RegionName |
| `VI SDK Server` + `Cluster` | SubAccountId / SubAccountName |
| `VM UUID` / `VM` | ResourceId / ResourceName (`Virtual Machine` o `VM Template`) |
| Atributos personalizados (`Aplicaciones`, `Ambiente`, `Empresa`…) | `Tags` (JSON), para Business Mappings y showback |
| `VM ID`, `Powerstate`, `Template`, `CPUs`, `Memory`, disco, SO, `Host` | `x_VMwareMoRef`, `x_PowerState`, `x_IsTemplate`, `x_VCPUCount`, `x_MemoryGB`, `x_StorageGB`, `x_GuestOS`, `x_Host` |
| Fecha de la foto usada | `x_SnapshotDate` |

Fechas en UTC con `Z`, punto decimal, UTF-8 sin BOM. `BillingPeriod*` = mes calendario, `ChargePeriod*` = día.

## Estructura en S3

```
<bucket>/
  vmware-inbox/<vcenter>/<fecha>/archivo      ← aquí se suben los exports (se vacía solo)
  vmware-state/config.json                    ← tarifario y políticas
  vmware-state/snapshots/<vcenter>/<fecha>.json.gz   ← fotos diarias normalizadas
  vmware-state/processed/<fecha>/archivo      ← originales ya procesados
  vmware-state/failed/<fecha>/archivo         ← archivos que no se pudieron leer
  vmware/VMware/20261101-20261201/<epoch>/vmware-focus-00001.csv.gz + vmware-focus-Manifest.json
```

## Operación

```bash
# Cambiar el tarifario: editar config.json y subirlo (aplica desde la siguiente regeneración)
aws s3 cp config.json s3://<bucket>/vmware-state/config.json
aws lambda invoke --function-name vmware-rvtools-focus --cli-binary-format raw-in-base64-out \
  --payload '{"action":"rebuild","months":["2026-11"]}' out.json

aws logs tail /aws/lambda/vmware-rvtools-focus --since 1d     # logs
```

Uso local, sin AWS:
```bash
python3 lambda/rvtools_focus.py ingest export.xlsx --date 2026-11-05 --state ./state --config config.json
python3 lambda/rvtools_focus.py build --state ./state --out ./out --root-dir <bucket> --month 2026-11 --config config.json
```

Pruebas: `pip install -r requirements-dev.txt && python -m pytest -q && cfn-lint template.yaml`

## Troubleshooting

| Síntoma | Causa probable | Acción |
|---|---|---|
| El archivo aparece en `vmware-state/failed/` | No es un vInfo (faltan columnas `VM`, `CPUs`, `Memory`, `Provisioned MiB`, `VI SDK Server`) | Exportar la pestaña **vInfo** completa |
| `coverage.missing` con días | No llegó foto y se superó `max_carry_forward_days` | Revisar el cron o la tarea programada. Recargar la foto con la fecha en la ruta |
| Falta un vCenter en Cloudability | Su foto no llegó ese día | Revisar el log del colector. El drop diario de las 23:30 UTC incluye todas las fotos disponibles |
| Colector: `CERTIFICATE_VERIFY_FAILED` | La CA de vCenter no es de confianza | Descargar `https://<vcenter>/certs/download.zip` y apuntar `VCENTER_CA_BUNDLE` |
| Colector: `InvalidLogin` | Usuario o contraseña incorrectos, o falta el dominio SSO | Usar `usuario@vsphere.local` o el dominio configurado |
| `Tags` sin `Ambiente`/`Aplicaciones` | La VM no tiene el atributo, o se excluyó en `tags` | Medir la cobertura con el cliente; revisar `tags.include` / `exclude_prefixes` |
| Costos muy altos o muy bajos | Tarifario de ejemplo, o moneda distinta | Ajustar `rates_monthly` y `currency`, subir la config y regenerar |
| Check verde en Cloudability pero sin datos | *Sub Directories* o *Manifest Prefix* distintos | Comparar el *Sample path* con la ruta real en S3 |

## Limitaciones

- No se validó con el `focus-validator` oficial de la FinOps Foundation. El formato sigue FOCUS 1.0 y la
  estructura de *Connect Custom Data* de IBM.
- El costo es una **asignación** según el tarifario interno, no una factura. Su exactitud depende del tarifario.
- El colector se probó contra el simulador oficial de vCenter (`vcsim`, govmomi). Antes de la puesta en marcha,
  conviene hacer una corrida en el vCenter del cliente y comparar el conteo de VMs con RVTools.
- Los parámetros de línea de comandos de RVTools pueden cambiar entre versiones; verifícalos en su documentación.

## Referencias

- [IBM Cloudability: Connect Custom Data (FOCUS Ingress)](https://www.ibm.com/docs/en/cloudability-commercial/cloudability-enterprise/saas?topic=cloudability-connect-custom-data-focus-ingress)
- [FOCUS Specification](https://focus.finops.org/)
- [pyVmomi](https://github.com/vmware/pyvmomi) · [RVTools](https://www.robware.net/)
