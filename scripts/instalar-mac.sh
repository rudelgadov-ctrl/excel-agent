#!/bin/bash
# Instala el complemento "Agente Office" en Excel y Word para Mac (solo para el usuario actual).
#
# Sin descargar nada (el script baja el manifiesto solo):
#   curl -fsSL https://raw.githubusercontent.com/rudelgadov-ctrl/excel-agent/main/scripts/instalar-mac.sh | bash
#   curl -fsSL https://raw.githubusercontent.com/rudelgadov-ctrl/excel-agent/main/scripts/instalar-mac.sh | bash -s -- --desinstalar
#
# Desde un clon del repo:
#   ./scripts/instalar-mac.sh               # producción (GitHub Pages)
#   ./scripts/instalar-mac.sh --dev         # desarrollo (https://localhost:3000)
#   ./scripts/instalar-mac.sh --desinstalar
set -euo pipefail

BASE_URL="https://rudelgadov-ctrl.github.io/excel-agent"
MANIFEST="manifest.xml"
TARGET_NAME="excel-agent.xml"
DEV=0
UNINSTALL=0
for arg in "$@"; do
  case "$arg" in
    --dev) DEV=1; MANIFEST="manifest.dev.xml"; TARGET_NAME="excel-agent-dev.xml" ;;
    --desinstalar) UNINSTALL=1 ;;
    *) echo "Opción desconocida: $arg" >&2; exit 1 ;;
  esac
done

# Cada aplicación lee los manifiestos de su propia carpeta "wef".
APPS="Excel Word"

if [ "$UNINSTALL" = 0 ]; then
  # Si el script se ejecuta desde un clon del repo se usa su manifiesto;
  # si no (curl | bash), el manifiesto de producción se descarga de GitHub Pages.
  SRC=""
  SCRIPT="${BASH_SOURCE[0]:-}"
  if [ -n "$SCRIPT" ] && [ -f "$SCRIPT" ]; then
    ROOT="$(cd "$(dirname "$SCRIPT")/.." && pwd)"
    [ -f "$ROOT/$MANIFEST" ] && SRC="$ROOT/$MANIFEST"
  fi
  if [ -z "$SRC" ]; then
    if [ "$DEV" = 1 ]; then
      echo "--dev necesita un clon del repo (manifest.dev.xml no se publica)." >&2
      exit 1
    fi
    SRC="$(mktemp)"
    trap 'rm -f "$SRC"' EXIT
    echo "Descargando $BASE_URL/manifest.xml ..."
    curl -fsSL "$BASE_URL/manifest.xml" -o "$SRC"
  fi
fi

for APP in $APPS; do
  WEF="$HOME/Library/Containers/com.microsoft.$APP/Data/Documents/wef"
  if [ "$UNINSTALL" = 1 ]; then
    rm -f "$WEF/$TARGET_NAME"
  else
    mkdir -p "$WEF"
    cp "$SRC" "$WEF/$TARGET_NAME"
    echo "Listo: manifiesto copiado a $WEF"
  fi
done

if [ "$UNINSTALL" = 1 ]; then
  echo "Complemento desinstalado de Excel y Word ($TARGET_NAME)."
  exit 0
fi

echo "Cierra Excel y Word (Cmd+Q) y vuelve a abrirlos. Luego: Insertar > Complementos > Mis complementos > Agente Office."
echo "Desde ahí queda el botón 'Agente Excel' o 'Agente Word' en la pestaña Inicio."
