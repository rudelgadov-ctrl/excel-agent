<#
.SYNOPSIS
  Registra el complemento "Agente Excel" en Excel para Windows (solo para el usuario actual).

.EXAMPLE
  powershell -ExecutionPolicy Bypass -File scripts\instalar-windows.ps1
  powershell -ExecutionPolicy Bypass -File scripts\instalar-windows.ps1 -Dev
  powershell -ExecutionPolicy Bypass -File scripts\instalar-windows.ps1 -Desinstalar
#>
param(
  [switch]$Dev,
  [switch]$Desinstalar
)
$ErrorActionPreference = "Stop"

$root = Split-Path -Parent $PSScriptRoot
$file = if ($Dev) { "manifest.dev.xml" } else { "manifest.xml" }
$manifest = Join-Path $root $file
if (-not (Test-Path $manifest)) { throw "No se encontró $manifest" }

[xml]$xml = Get-Content -Raw -Encoding UTF8 $manifest
$id = $xml.OfficeApp.Id

# Misma clave que usa office-addin-dev-settings para cargar complementos de desarrollo.
$key = "HKCU:\Software\Microsoft\Office\16.0\WEF\Developer"

if ($Desinstalar) {
  if (Test-Path $key) { Remove-ItemProperty -Path $key -Name $id -ErrorAction SilentlyContinue }
  Write-Host "Complemento desregistrado ($file)."
  return
}

# Ojo: New-Item -Force sobre una clave existente borraría sus valores; solo se crea si falta.
if (-not (Test-Path $key)) { New-Item -Path $key -Force | Out-Null }
Set-ItemProperty -Path $key -Name $id -Value $manifest

Write-Host "Listo: $file registrado."
Write-Host "Cierra Excel por completo y vuelve a abrirlo. El botón 'Agente Excel' aparece en la pestaña Inicio."
