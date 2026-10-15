#!/bin/bash
# Instala el complemento "Agente Office" en Excel y Word para Mac (solo para el usuario actual).
#
# Uso:
#   ./scripts/instalar-mac.sh               # producción (GitHub Pages)
#   ./scripts/instalar-mac.sh --dev         # desarrollo (https://localhost:3000)
#   ./scripts/instalar-mac.sh --desinstalar
set -euo pipefail
cd "$(dirname "$0")/.."

MANIFEST="manifest.xml"
TARGET_NAME="excel-agent.xml"
UNINSTALL=0
for arg in "$@"; do
  case "$arg" in
    --dev) MANIFEST="manifest.dev.xml"; TARGET_NAME="excel-agent-dev.xml" ;;
    --desinstalar) UNINSTALL=1 ;;
    *) echo "Opción desconocida: $arg" >&2; exit 1 ;;
  esac
done

# Cada aplicación lee los manifiestos de su propia carpeta "wef".
APPS="Excel Word"

for APP in $APPS; do
  WEF="$HOME/Library/Containers/com.microsoft.$APP/Data/Documents/wef"
  if [ "$UNINSTALL" = 1 ]; then
    rm -f "$WEF/$TARGET_NAME"
  else
    mkdir -p "$WEF"
    cp "$MANIFEST" "$WEF/$TARGET_NAME"
    echo "Listo: $MANIFEST copiado a $WEF"
  fi
done

if [ "$UNINSTALL" = 1 ]; then
  echo "Complemento desinstalado de Excel y Word ($TARGET_NAME)."
  exit 0
fi

echo "Cierra Excel y Word (Cmd+Q) y vuelve a abrirlos. Luego: Insertar > Complementos > Mis complementos > Agente Office."
echo "Desde ahí queda el botón 'Agente Excel' o 'Agente Word' en la pestaña Inicio."
