<#
.SYNOPSIS
  Registra el complemento "Agente Office" en Excel y Word para Windows (solo para el usuario actual).

.EXAMPLE
  # Sin descargar nada (el script baja el manifiesto solo):
  irm https://raw.githubusercontent.com/rudelgadov-ctrl/excel-agent/main/scripts/instalar-windows.ps1 | iex
  & ([scriptblock]::Create((irm https://raw.githubusercontent.com/rudelgadov-ctrl/excel-agent/main/scripts/instalar-windows.ps1))) -Desinstalar

.EXAMPLE
  # Desde un clon del repo:
  powershell -ExecutionPolicy Bypass -File scripts\instalar-windows.ps1
  powershell -ExecutionPolicy Bypass -File scripts\instalar-windows.ps1 -Dev
  powershell -ExecutionPolicy Bypass -File scripts\instalar-windows.ps1 -Desinstalar
#>
param(
  [switch]$Dev,
  [switch]$Desinstalar
)
$ErrorActionPreference = "Stop"
$baseUrl = "https://rudelgadov-ctrl.github.io/excel-agent"
$file = if ($Dev) { "manifest.dev.xml" } else { "manifest.xml" }

# Si el script se ejecuta desde un clon del repo se usa su manifiesto. Si no (irm | iex),
# el manifiesto de produccion se descarga de GitHub Pages a una carpeta fija del usuario.
$manifest = $null
if ($PSScriptRoot) {
  $candidate = Join-Path (Split-Path -Parent $PSScriptRoot) $file
  if (Test-Path $candidate) { $manifest = $candidate }
}
$descargado = $false
if (-not $manifest) {
  if ($Dev) { throw "-Dev necesita un clon del repo (manifest.dev.xml no se publica)." }
  $dir = Join-Path $env:LOCALAPPDATA "AgenteOffice"
  New-Item -ItemType Directory -Path $dir -Force | Out-Null
  $manifest = Join-Path $dir "manifest.xml"
  # Windows PowerShell 5.1 puede arrancar sin TLS 1.2.
  [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
  Write-Host "Descargando $baseUrl/manifest.xml ..."
  Invoke-WebRequest -UseBasicParsing -Uri "$baseUrl/manifest.xml" -OutFile $manifest
  $descargado = $true
}

[xml]$xml = Get-Content -Raw -Encoding UTF8 $manifest
$id = $xml.OfficeApp.Id

# Misma clave que usa office-addin-dev-settings para cargar complementos de desarrollo.
# Es comun a todas las aplicaciones de Office: el manifiesto indica que el complemento es para Excel y Word.
$key = "HKCU:\Software\Microsoft\Office\16.0\WEF\Developer"

if ($Desinstalar) {
  if (Test-Path $key) { Remove-ItemProperty -Path $key -Name $id -ErrorAction SilentlyContinue }
  if ($descargado) { Remove-Item -Recurse -Force (Split-Path -Parent $manifest) -ErrorAction SilentlyContinue }
  Write-Host "Complemento desregistrado ($file)."
  return
}

# Ojo: New-Item -Force sobre una clave existente borraria sus valores; solo se crea si falta.
if (-not (Test-Path $key)) { New-Item -Path $key -Force | Out-Null }
Set-ItemProperty -Path $key -Name $id -Value $manifest

Write-Host "Listo: $file registrado."
Write-Host "Cierra Excel y Word por completo y vuelve a abrirlos. El boton 'Agente Excel' o 'Agente Word' aparece en la ficha Inicio."
