# Agente Excel

Complemento de Excel con un agente (Claude) que resuelve **cualquier práctica, tarea o requerimiento de
Excel** sobre el libro abierto, incluidas prácticas nuevas que nunca ha visto. Abres la plantilla, pegas
las instrucciones o adjuntas el archivo, revisas el plan que propone y presionas **Aprobar y ejecutar**.
Funciona en Excel para **Windows** y **Mac** (y Excel en la web).

> **Alcance:** no está programado para prácticas específicas. Las prácticas de `practicas/` solo sirven
> de ejemplo del tipo de requerimientos; el agente lee las instrucciones de cada práctica nueva, inspecciona
> el libro y decide qué hacer. Hay una prueba automática que falla si el prompt o las herramientas
> mencionan una práctica concreta.

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
2. Pega las instrucciones o **adjunta el archivo** (arrástralo al cuadro de texto): PDF, Word (.docx),
   texto (.txt/.md/.csv) o imagen (captura). Con PDF o imagen el agente ve también las figuras
   (por ejemplo, el formato de una factura o un gráfico de ejemplo); de Word solo toma el texto, así que
   si la práctica tiene figuras, exporta a PDF. Si las instrucciones ya están dentro del libro
   (una hoja o un cuadro de texto), basta con escribir "resuelve las instrucciones del libro".
3. El agente arma la lista de requerimientos, inspecciona el libro y propone un **plan** con la misma
   numeración de la práctica. Si quieres cambios, escríbelos; si está bien, presiona **Aprobar y ejecutar**.
4. Al terminar muestra una tabla con **todos** los puntos: hecho, hecho con alternativa o pasos manuales.

En **Ajustes** puedes activar el **modo automático** (ejecuta sin pedir aprobación), cambiar el modelo
(Opus 5.5 recomendado; Sonnet 5.5 cuesta la mitad) y el nivel de esfuerzo.

## Qué sabe hacer

| Herramienta | Para qué |
|---|---|
| `leer_libro`, `leer_rango` | Inspeccionar hojas, celdas, fórmulas, formatos, tablas, cuadros de texto |
| `escribir_rango` | Valores y fórmulas de cualquier tipo (detecta celdas con error) |
| `rellenar_serie` | Relleno automático (series, fechas, copiar fórmulas) |
| `copiar_rango`, `dividir_texto` | Copiar y pegado especial (también entre hojas); texto en columnas |
| `insertar_eliminar`, `quitar_duplicados` | Filas, columnas y celdas; duplicados |
| `ordenar`, `filtrar` | Ordenar por varias columnas; autofiltro (valores, criterios, top N) |
| `crear_nombre`, `validacion_datos` | Rangos con nombre; listas desplegables, números, fechas, fórmulas |
| `formatear_rango` | Moneda, fechas, %, alineación, fuente, relleno, bordes, combinar, anchos |
| `formato_condicional` | Reglas, texto, top/bottom, fórmulas, duplicados, escalas, barras, íconos |
| `crear_tabla`, `tabla_dinamica` | Dar formato como tabla; tablas dinámicas con filas, columnas, valores y filtros |
| `crear_grafico` | Columnas, barras, líneas, circular, anillo, área, dispersión, radial, 3D, histograma… |
| `gestionar_hojas`, `configurar_hoja` | Crear/copiar/renombrar hojas; inmovilizar, pestaña, proteger, impresión |
| `entregar_vba` | Código de macros listo para pegar |
| `ejecutar_office_js` | Cualquier otra cosa que permita la API de Excel (hipervínculos, comentarios, imágenes…) |

### Lo que un complemento no puede hacer directamente

Para estas funciones Excel no ofrece API a los complementos. El agente no las omite: usa la mejor
alternativa y lo indica en el plan y en el resumen.

- **Macros** (grabar macros, botones con macros): implementa el comportamiento con fórmulas para que
  funcione de una vez y además te muestra el código VBA con botón **Copiar** y los pasos para pegarlo
  (Alt+F11 en Windows; Herramientas > Macro > Editor de Visual Basic en Mac) y guardar como `.xlsm`.
- **Buscar objetivo, Solver, escenarios**: calcula la solución con fórmulas, escribe el resultado y
  entrega el VBA o los pasos manuales si la práctica exige usar la herramienta.
- **Tablas de datos (análisis Y si)**: las reproduce con fórmulas equivalentes y explica cómo crear la original.
- **Minigráficos, Power Query, guardar como otro formato, imprimir**: pasos manuales precisos.

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
