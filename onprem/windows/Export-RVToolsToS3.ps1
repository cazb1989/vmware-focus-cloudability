<#
.SYNOPSIS
  Alternativa Windows: exporta la pestaña vInfo con RVTools y la sube al inbox de S3.
  Programar con el Programador de tareas (diario, 16:00 CDMX / 22:00 UTC).

.EXAMPLE
  .\Export-RVToolsToS3.ps1 -VCenters vcenter01.empresa.local,vcenter02.empresa.local `
      -Bucket mi-bucket-focus -AwsProfile vmware-uploader

.NOTES
  - Requiere RVTools 4.x y AWS CLI v2 (perfil con la access key de vmware-inventory-uploader).
  - Autenticación: -passthroughAuth (cuenta de Windows con permiso de lectura en vCenter) o
    -User + -EncryptedPassword (generado con RVToolsPasswordEncryption.exe).
  - Verifique los parámetros de línea de comandos en la documentación de su versión de RVTools.
#>
param(
  [Parameter(Mandatory)] [string[]] $VCenters,
  [Parameter(Mandatory)] [string]   $Bucket,
  [string] $InboxPrefix = "vmware-inbox",
  [string] $AwsProfile  = "vmware-uploader",
  [string] $RVTools     = "C:\Program Files (x86)\Dell\RVTools\RVTools.exe",
  [string] $OutDir      = "C:\RVToolsExports",
  [string] $User,
  [string] $EncryptedPassword
)
$ErrorActionPreference = "Stop"
$day = (Get-Date).ToUniversalTime().ToString("yyyy-MM-dd")   # la foto se registra con la fecha UTC
$dir = Join-Path $OutDir $day
New-Item -ItemType Directory -Force -Path $dir | Out-Null
$failed = 0

foreach ($vc in $VCenters) {
  $file = "vInfo_${vc}_${day}.xlsx"
  $auth = if ($User) { @("-u", $User, "-p", $EncryptedPassword) } else { @("-passthroughAuth") }
  Write-Host "$(Get-Date -Format s) ${vc}: exportando vInfo..."
  $p = Start-Process -FilePath $RVTools -Wait -PassThru -NoNewWindow `
         -ArgumentList ($auth + @("-s", $vc, "-c", "ExportvInfo2xlsx", "-d", $dir, "-f", $file))
  $path = Join-Path $dir $file
  if ($p.ExitCode -ne 0 -or -not (Test-Path $path)) { Write-Warning "RVTools falló para $vc"; $failed++; continue }

  aws s3 cp $path "s3://$Bucket/$InboxPrefix/$vc/$day/$file" --profile $AwsProfile
  if ($LASTEXITCODE -ne 0) { Write-Warning "Falló la subida de $file"; $failed++ }
  Start-Sleep -Seconds 60   # escalona los vCenter para que cada drop los incluya a todos
}

Get-ChildItem $OutDir -Directory | Where-Object { $_.LastWriteTime -lt (Get-Date).AddDays(-30) } |
  Remove-Item -Recurse -Force
exit $failed
