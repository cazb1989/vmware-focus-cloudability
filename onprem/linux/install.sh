#!/usr/bin/env bash
# Instala el colector en un servidor Linux con acceso de red a vCenter y salida HTTPS a S3.
# Probado en RHEL/Rocky/Alma 8-9 y Ubuntu 22.04+. Ejecutar como root:  sudo ./install.sh
set -euo pipefail
APP_DIR=/opt/vmware-focus; DATA_DIR=/var/lib/vmware-focus; ETC=/etc/vmware-focus; SVC=vmwarefocus
cd "$(dirname "$0")"
command -v python3 >/dev/null || { echo "Instala python3 (3.8+) primero"; exit 1; }
id "$SVC" >/dev/null 2>&1 || useradd --system --home-dir "$APP_DIR" --shell /sbin/nologin "$SVC"
install -d -o "$SVC" -g "$SVC" -m 750 "$APP_DIR" "$DATA_DIR" /var/log/vmware-focus
install -d -m 750 -g "$SVC" "$ETC"
install -m 755 vcenter_collector.py run.sh "$APP_DIR/"
python3 -m venv "$APP_DIR/venv"
"$APP_DIR/venv/bin/pip" install --quiet --upgrade pip
"$APP_DIR/venv/bin/pip" install --quiet pyvmomi requests boto3
chown -R "$SVC:$SVC" "$APP_DIR"
if [ ! -f "$ETC/env" ]; then
  install -m 640 -g "$SVC" env.example "$ETC/env"
  echo "==> Edita $ETC/env (vCenter, usuario, contraseña, bucket y access key)"
fi
cat > /etc/cron.d/vmware-focus <<CRON
# Inventario diario de VMware -> S3 (22:00 UTC = 16:00 CDMX / 17:00 Lima). La Lambda regenera a las 23:30 UTC.
0 22 * * * $SVC $APP_DIR/run.sh >> /var/log/vmware-focus/collector.log 2>&1
CRON
chmod 644 /etc/cron.d/vmware-focus
echo "Listo. Prueba manual:  sudo -u $SVC $APP_DIR/run.sh"
