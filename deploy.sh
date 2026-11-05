#!/usr/bin/env bash
# Despliega RVTools -> FOCUS -> S3 en tu cuenta AWS (recomendado: AWS CloudShell).
#
# Uso:
#   BUCKET=<bucket-dedicado> REGION=us-east-1 ALERT_EMAIL=<email-opcional> ./deploy.sh
#   UPDATE_CONFIG=yes ./deploy.sh      # sube de nuevo config.json (tarifario) al bucket
set -euo pipefail

: "${BUCKET:?Define BUCKET (bucket dedicado para FOCUS)}"
REGION="${REGION:-us-east-1}"
SUB_DIR="${SUB_DIR:-vmware}"
INBOX="${INBOX:-vmware-inbox}"
STATE="${STATE:-vmware-state}"
ALERT_EMAIL="${ALERT_EMAIL:-}"
STACK="${STACK:-vmware-rvtools-focus}"
CONFIG="${CONFIG:-config.json}"
cd "$(dirname "$0")"
[ -f "$CONFIG" ] || { cp config.example.json "$CONFIG"; echo "==> Se creó $CONFIG desde config.example.json (revisa el tarifario)"; }
python3 -c "import json,sys; json.load(open(sys.argv[1]))" "$CONFIG" || { echo "ERROR: $CONFIG no es JSON válido"; exit 1; }

ACCOUNT=$(aws sts get-caller-identity --query Account --output text)
echo "==> Cuenta: $ACCOUNT  Región: $REGION"

# 1) Bucket dedicado, privado y cifrado
if ! aws s3api head-bucket --bucket "$BUCKET" 2>/dev/null; then
  echo "==> Creando bucket $BUCKET"
  if [ "$REGION" = "us-east-1" ]; then
    aws s3api create-bucket --bucket "$BUCKET" --region "$REGION" >/dev/null
  else
    aws s3api create-bucket --bucket "$BUCKET" --region "$REGION" \
      --create-bucket-configuration LocationConstraint="$REGION" >/dev/null
  fi
  aws s3api put-public-access-block --bucket "$BUCKET" --public-access-block-configuration \
    BlockPublicAcls=true,IgnorePublicAcls=true,BlockPublicPolicy=true,RestrictPublicBuckets=true
  aws s3api put-bucket-encryption --bucket "$BUCKET" --server-side-encryption-configuration \
    '{"Rules":[{"ApplyServerSideEncryptionByDefault":{"SSEAlgorithm":"AES256"}}]}'
fi

# 2) Notificaciones del bucket a EventBridge (conserva las notificaciones que ya existan)
CURRENT=$(aws s3api get-bucket-notification-configuration --bucket "$BUCKET" --output json 2>/dev/null || echo '{}')
NEWCONF=$(printf '%s' "$CURRENT" | python3 -c 'import json,sys; c=json.loads(sys.stdin.read() or "{}"); c["EventBridgeConfiguration"]={}; print(json.dumps(c))')
aws s3api put-bucket-notification-configuration --bucket "$BUCKET" --notification-configuration "$NEWCONF"

# 3) Configuración (tarifario y políticas) en el bucket
if [ "${UPDATE_CONFIG:-no}" = "yes" ] || ! aws s3api head-object --bucket "$BUCKET" --key "$STATE/config.json" >/dev/null 2>&1; then
  aws s3 cp "$CONFIG" "s3://$BUCKET/$STATE/config.json" --region "$REGION" >/dev/null
  echo "==> config.json subido a s3://$BUCKET/$STATE/config.json"
fi

# 4) Empaquetar el código + openpyxl (para leer .xlsx) y subirlo
rm -rf build build.zip && mkdir build
python3 -m pip install --quiet --target build openpyxl
cp lambda/handler.py lambda/rvtools_focus.py build/
(cd build && zip -qr ../build.zip .)
HASH=$( (sha256sum build.zip 2>/dev/null || shasum -a 256 build.zip) | cut -c1-12)
CODE_KEY="_lambda/vmware-rvtools-focus-$HASH.zip"
aws s3 cp build.zip "s3://$BUCKET/$CODE_KEY" --region "$REGION" >/dev/null

# 5) Stack
aws cloudformation deploy --region "$REGION" --stack-name "$STACK" \
  --template-file template.yaml --capabilities CAPABILITY_NAMED_IAM \
  --parameter-overrides BucketName="$BUCKET" CodeKey="$CODE_KEY" SubDir="$SUB_DIR" \
    InboxPrefix="$INBOX" StatePrefix="$STATE" AlertEmail="$ALERT_EMAIL"

echo
echo "Listo."
echo "1) Sube un export de RVTools para probar:"
echo "     aws s3 cp <archivo>.xlsx s3://$BUCKET/$INBOX/<vcenter>/\$(date -u +%F)/<archivo>.xlsx"
echo "     aws logs tail /aws/lambda/vmware-rvtools-focus --follow"
echo "2) Credencial para el servidor on-prem (solo puede subir al inbox):"
echo "     aws iam create-access-key --user-name vmware-inventory-uploader"
echo "3) En Cloudability: Settings > Vendor Credentials > Add Datasource > Other Data Sources"
echo "     Storage        : AWS Simple Storage Service"
echo "     Vendor         : VMware"
echo "     AWS Account ID : $ACCOUNT"
echo "     Root Directory : $BUCKET"
echo "     Sub Directories: $SUB_DIR"
echo "     Manifest Prefix: $(python3 -c 'import json,sys; print(json.load(open(sys.argv[1])).get("file_prefix","vmware-focus"))' "$CONFIG")"
