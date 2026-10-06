# Agente Excel

Complemento de Excel con un agente (Claude) que ejecuta requerimientos sobre el libro abierto.
Abres una plantilla, pegas las instrucciones o adjuntas el PDF, revisas el plan que propone y
presionas **Aprobar y ejecutar**. Funciona en Excel para **Windows** y **Mac** (y Excel en la web).

- Panel publicado en: <https://rudelgadov-ctrl.github.io/excel-agent/>
- Cada `push` a `main` lo vuelve a publicar automáticamente (GitHub Actions).

## Requisitos

- **Excel de Microsoft 365** (recomendado) o Excel 2021/2024. Para ver tu versión:
  Archivo > Cuenta (Windows) o Excel > Acerca de Excel (Mac). Con Excel 2019 o anterior
  algunas funciones usan alternativas o no están disponibles.
- **API key de Anthropic** con crédito: <https://console.anthropic.com/settings/keys>.
  La suscripción Claude Pro/Max **no** incluye uso de la API; se paga aparte por consumo.

## Instalación (una vez por computadora)

### Windows

```powershell
git clone https://github.com/rudelgadov-ctrl/excel-agent.git
cd excel-agent
powershell -ExecutionPolicy Bypass -File scripts\instalar-windows.ps1
```

Cierra Excel por completo y vuelve a abrirlo: aparece el botón **Agente Excel** en la pestaña Inicio.

### Mac

```bash
git clone https://github.com/rudelgadov-ctrl/excel-agent.git
cd excel-agent
./scripts/instalar-mac.sh
```

Sin clonar el repo también sirve:

```bash
mkdir -p ~/Library/Containers/com.microsoft.Excel/Data/Documents/wef
curl -o ~/Library/Containers/com.microsoft.Excel/Data/Documents/wef/excel-agent.xml \
  https://rudelgadov-ctrl.github.io/excel-agent/manifest.xml
```

Cierra Excel (Cmd+Q), ábrelo y ve a **Insertar > Complementos > Mis complementos > Agente Excel**.

### Primera vez en el panel

Abre **Ajustes**, pega tu API key y guarda. Se guarda solo en esa computadora (no viaja al repo
ni a GitHub), así que hay que ponerla en cada máquina.

## Uso

1. Abre la plantilla (mejor una **copia**: algunos cambios del agente no se deshacen con Ctrl+Z).
2. Pega las instrucciones o **adjunta el PDF** (arrástralo al cuadro de texto). Con el PDF, el agente
   ve también las figuras, como el formato de una factura o un gráfico de ejemplo.
3. El agente inspecciona el libro y propone un **plan** numerado. Si quieres cambios, escríbelos;
   si está bien, presiona **Aprobar y ejecutar**.
4. Al terminar muestra un resumen por punto. Revisa el libro.

En **Ajustes** puedes activar el **modo automático** (ejecuta sin pedir aprobación), cambiar el modelo
(Opus 5.5 recomendado; Sonnet 5.5 cuesta la mitad) y el nivel de esfuerzo.

### Macros

Un complemento de Office no puede crear macros VBA ni guardar archivos `.xlsm`. Cuando una
instrucción pide "grabar una macro", el agente:

1. Implementa el mismo comportamiento con fórmulas, para que funcione de una vez.
2. Te muestra el código VBA completo con un botón **Copiar** y los pasos para pegarlo
   (Alt+F11 en Windows; Herramientas > Macro > Editor de Visual Basic en Mac) y guardar como `.xlsm`.

## Qué sabe hacer

| Herramienta | Para qué |
|---|---|
| `leer_libro`, `leer_rango` | Inspeccionar hojas, celdas, fórmulas y formatos |
| `escribir_rango` | Valores y fórmulas (detecta celdas con error) |
| `formatear_rango` | Moneda, fechas, %, alineación, fuente, relleno, bordes, combinar, anchos |
| `crear_nombre` | Rangos con nombre |
| `rellenar_serie` | Relleno automático (series, fechas mensuales, copiar fórmulas) |
| `dividir_texto` | Texto en columnas |
| `copiar_rango`, `gestionar_hojas` | Copiar datos entre hojas; crear/copiar/renombrar hojas |
| `validacion_datos` | Listas desplegables (también contra otra hoja), números, fechas |
| `ordenar` | Ordenar por una o más columnas |
| `crear_grafico` | Columnas, barras, líneas, circular, radial, 3D… con series a elección |
| `entregar_vba` | Código de macros listo para pegar |
| `ejecutar_office_js` | Último recurso para lo que no cubren las demás |

## Costo aproximado

El panel muestra el costo estimado de la conversación. Como referencia, una práctica completa
suele costar del orden de centavos a uno o dos dólares con Opus 5.5, según el tamaño del libro,
los PDF adjuntos y cuántas correcciones haga. Con Sonnet 5.5 cuesta aproximadamente la mitad.

## Privacidad y seguridad

- La API key se guarda en el almacenamiento local del complemento en cada computadora.
  Ese almacenamiento se comparte con otras páginas que publiques en `rudelgadov-ctrl.github.io`.
- El contenido del libro que el agente lee y los archivos que adjuntas se envían a la API de Anthropic.
- `ejecutar_office_js` ejecuta código generado por el agente dentro del libro abierto. Solo puede
  actuar sobre ese libro, pero revisa el plan antes de aprobarlo.

## Desarrollo

```bash
npm install
npm test            # pruebas unitarias
npm run build       # tipos + build de producción en dist/
npm run validate    # valida los manifiestos con el servicio de Microsoft
```

Probar cambios sin publicar (panel desde `https://localhost:3000`):

```bash
npm run certs       # una vez: crea y confía el certificado HTTPS de desarrollo (pide confirmación)
npm run dev
```

Y registra el manifiesto de desarrollo, que aparece como **Agente (dev)**:

- Windows: `powershell -ExecutionPolicy Bypass -File scripts\instalar-windows.ps1 -Dev`
- Mac: `./scripts/instalar-mac.sh --dev`

Para desinstalar, los mismos scripts aceptan `-Desinstalar` / `--desinstalar`.

### Estructura

```
manifest.xml / manifest.dev.xml   Manifiestos del complemento (producción / desarrollo)
taskpane.html, src/taskpane/      Panel: chat, adjuntos, plan, tarjetas VBA, ajustes
src/agent/                        Bucle con la API de Claude, prompt de sistema, ajustes
src/tools/                        Herramientas Office.js que usa el agente
src/excel/                        Helpers (direcciones, series, texto, Office.js)
tests/                            Pruebas (vitest)
practicas/                        Material de prueba local (no se sube a GitHub)
```

## Solución de problemas

- **El botón no aparece**: cierra Excel por completo (también en segundo plano) y vuelve a abrirlo.
  En Mac, entra a Insertar > Complementos > Mis complementos.
- **El panel no se actualiza tras publicar**: Office guarda caché. Windows: cierra Excel y borra
  `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\`. Mac: borra
  `~/Library/Containers/com.microsoft.Excel/Data/Library/Caches/` y reabre Excel.
- **"La API key no es válida"**: revisa la key en Ajustes y que tenga crédito en la consola.
