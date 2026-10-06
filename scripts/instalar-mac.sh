#!/bin/bash
# Instala el complemento "Agente Excel" en Excel para Mac (solo para el usuario actual).
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

WEF="$HOME/Library/Containers/com.microsoft.Excel/Data/Documents/wef"

if [ "$UNINSTALL" = 1 ]; then
  rm -f "$WEF/$TARGET_NAME"
  echo "Complemento desinstalado ($TARGET_NAME)."
  exit 0
fi

mkdir -p "$WEF"
cp "$MANIFEST" "$WEF/$TARGET_NAME"
echo "Listo: $MANIFEST copiado a $WEF"
echo "Cierra Excel (Cmd+Q) y vuelve a abrirlo. Luego: Insertar > Complementos > Mis complementos > Agente Excel."
echo "Desde ahí queda el botón 'Agente Excel' en la pestaña Inicio."
