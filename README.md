# Agente Office (Excel y Word)

Complemento de **Excel y Word** con un agente (Claude) que resuelve **cualquier práctica, tarea o
requerimiento** sobre el libro o el documento abierto, incluidas prácticas nuevas que nunca ha visto.
Abres la plantilla, pegas las instrucciones o adjuntas el archivo, revisas el plan que propone y presionas
**Aprobar y ejecutar**. Es el mismo complemento en las dos aplicaciones: el panel detecta si se abrió en
Excel o en Word y usa las herramientas de esa aplicación.
Funciona en **Windows** y **Mac** (y en Excel y Word en la web).

> **Alcance:** no está programado para prácticas específicas. Las prácticas de `practicas/` solo sirven
> de ejemplo del tipo de requerimientos; el agente lee las instrucciones de cada práctica nueva, inspecciona
> el archivo y decide qué hacer. Hay una prueba automática que falla si el prompt o las herramientas
> mencionan una práctica concreta.

- Panel publicado en: <https://rudelgadov-ctrl.github.io/excel-agent/>
- Cada `push` a `main` lo vuelve a publicar automáticamente (GitHub Actions).

## Requisitos

- **Excel o Word de Microsoft 365** (recomendado) o 2021/2024. Para ver tu versión:
  Archivo > Cuenta (Windows) o Excel/Word > Acerca de (Mac).
  - Excel 2019 o anterior: algunas funciones usan alternativas o no están disponibles.
  - Word 2016/2019: lo básico funciona (texto, estilos, listas, tablas, encabezados y pies). Configurar la
    página (márgenes, orientación, columnas), el ajuste de texto de imágenes y algunos formatos de fuente
    requieren Word de escritorio de Microsoft 365 actualizado; si faltan, el agente entrega VBA o los pasos
    manuales.
- **API key de Anthropic** con crédito: <https://console.anthropic.com/settings/keys>.
  La suscripción Claude Pro/Max **no** incluye uso de la API; se paga aparte por consumo.

## Instalación (una vez por computadora)

### Windows

```powershell
git clone https://github.com/rudelgadov-ctrl/excel-agent.git
cd excel-agent
powershell -ExecutionPolicy Bypass -File scripts\instalar-windows.ps1
```

Cierra Excel y Word por completo y vuelve a abrirlos: aparece el botón **Agente Excel** (o **Agente Word**)
en la pestaña Inicio.

### Mac

```bash
git clone https://github.com/rudelgadov-ctrl/excel-agent.git
cd excel-agent
./scripts/instalar-mac.sh
```

El script instala el complemento en Excel y en Word. Sin clonar el repo también sirve:

```bash
for APP in Excel Word; do
  mkdir -p ~/Library/Containers/com.microsoft.$APP/Data/Documents/wef
  curl -o ~/Library/Containers/com.microsoft.$APP/Data/Documents/wef/excel-agent.xml \
    https://rudelgadov-ctrl.github.io/excel-agent/manifest.xml
done
```

Cierra Excel y Word (Cmd+Q), ábrelos y ve a **Insertar > Complementos > Mis complementos > Agente Office**.

### Primera vez en el panel

Abre **Ajustes**, pega tu API key y guarda. Se guarda solo en esa computadora (no viaja al repo
ni a GitHub), así que hay que ponerla en cada máquina.

## Uso

1. Abre la plantilla (mejor una **copia**: algunos cambios del agente no se deshacen con Ctrl+Z).
2. Pega las instrucciones o **adjunta el archivo** (arrástralo al cuadro de texto): PDF, Word (.docx),
   texto (.txt/.md/.csv) o imagen (captura). Con PDF o imagen el agente ve también las figuras
   (por ejemplo, el formato de una factura, de una carta o un gráfico de ejemplo); de Word solo toma el
   texto, así que si la práctica tiene figuras, exporta a PDF. Si las instrucciones ya están dentro del
   archivo (una hoja, un cuadro de texto o el propio documento), basta con escribir "resuelve las
   instrucciones del archivo".
3. El agente arma la lista de requerimientos, inspecciona el archivo y propone un **plan** con la misma
   numeración de la práctica. Si quieres cambios, escríbelos; si está bien, presiona **Aprobar y ejecutar**.
4. Al terminar muestra una tabla con **todos** los puntos: hecho, hecho con alternativa o pasos manuales.

En Word, las imágenes que adjuntes en la conversación también se pueden insertar en el documento
("inserta el logo adjunto en el encabezado").

En **Ajustes** puedes activar el **modo automático** (ejecuta sin pedir aprobación), cambiar el modelo
(Opus 5.5 recomendado; Sonnet 5.5 cuesta la mitad) y el nivel de esfuerzo.

## Qué sabe hacer

### En Excel

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

### En Word

| Herramienta | Para qué |
|---|---|
| `leer_documento`, `leer_tabla` | Párrafos numerados con estilo y formato, tablas, secciones, encabezados y pies, página, campos |
| `insertar_contenido` | Redactar párrafos con estilo, texto en línea, buscar y reemplazar, HTML con formato mixto, OOXML |
| `eliminar_contenido` | Borrar párrafos, tablas o texto; limpiar párrafos vacíos |
| `formato_texto` | Fuente, tamaño, negrita, cursiva, subrayado, color, resaltado, mayúsculas, versalitas… |
| `formato_parrafo` | Estilos, alineación, sangrías (también francesa), espaciado, interlineado, sombreado |
| `aplicar_lista`, `modificar_estilo` | Viñetas y numeración con niveles; modificar o crear estilos (Título 1, Normal…) |
| `insertar_tabla`, `editar_tabla`, `formato_tabla` | Tablas con estilo; filas, columnas, combinar celdas; sombreado, bordes, anchos |
| `configurar_pagina`, `insertar_salto` | Márgenes, orientación, tamaño, columnas, primera página diferente; saltos de página y sección |
| `insertar_campo` | Número de página, "Página X de Y", fecha, tabla de contenido, SEQ para figuras, REF, índice, MERGEFIELD |
| `insertar_vinculo_o_nota` | Hipervínculos, notas al pie y al final, comentarios, marcadores |
| `insertar_imagen` | Imágenes adjuntadas en la conversación, con tamaño, alineación y ajuste de texto |
| `entregar_vba` | Código de macros listo para pegar |
| `ejecutar_office_js` | Cualquier otra cosa que permita la API de Word (propiedades, control de cambios, formularios…) |

Encabezados y pies de página se trabajan con las mismas herramientas, indicando la zona.

### Lo que un complemento no puede hacer directamente

Para estas funciones Office no ofrece API a los complementos. El agente no las omite: usa la mejor
alternativa y lo indica en el plan y en el resumen.

- **Macros** (grabar macros, botones con macros): en Excel implementa el comportamiento con fórmulas para
  que funcione de una vez; en ambos casos muestra el código VBA con botón **Copiar** y los pasos para pegarlo
  (Alt+F11 en Windows; Herramientas > Macro > Editor de Visual Basic en Mac) y guardar como `.xlsm` o `.docm`.
- **Excel: Buscar objetivo, Solver, escenarios**: calcula la solución con fórmulas, escribe el resultado y
  entrega el VBA o los pasos manuales si la práctica exige usar la herramienta.
- **Excel: tablas de datos (análisis Y si)**: las reproduce con fórmulas equivalentes y explica cómo crear la original.
- **Excel: minigráficos, Power Query, guardar como otro formato, imprimir**: pasos manuales precisos.
- **Word: combinación de correspondencia**: inserta los campos MERGEFIELD, prepara la lista de datos y da los
  pasos para conectar los destinatarios y combinar.
- **Word: citas y bibliografía, sobres y etiquetas, SmartArt, WordArt, marca de agua, color o bordes de página,
  letra capital, proteger, guardar como PDF**: lo aproxima cuando se puede y entrega pasos manuales (y VBA si aplica).

## Costo aproximado

El panel muestra el costo estimado de la conversación. Como referencia, una práctica completa
suele costar del orden de centavos a uno o dos dólares con Opus 5.5, según el tamaño del archivo,
los PDF adjuntos y cuántas correcciones haga. Con Sonnet 5.5 cuesta aproximadamente la mitad.

## Privacidad y seguridad

- La API key se guarda en el almacenamiento local del complemento en cada computadora.
  Ese almacenamiento se comparte con otras páginas que publiques en `rudelgadov-ctrl.github.io`.
- El contenido del libro o documento que el agente lee y los archivos que adjuntas se envían a la API de Anthropic.
- `ejecutar_office_js` ejecuta código generado por el agente dentro del archivo abierto. Solo puede
  actuar sobre ese archivo, pero revisa el plan antes de aprobarlo.

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

Y registra el manifiesto de desarrollo, que aparece como **Agente (dev)** en Excel y Word:

- Windows: `powershell -ExecutionPolicy Bypass -File scripts\instalar-windows.ps1 -Dev`
- Mac: `./scripts/instalar-mac.sh --dev`

Para desinstalar, los mismos scripts aceptan `-Desinstalar` / `--desinstalar`.

### Estructura

```
manifest.xml / manifest.dev.xml   Manifiestos del complemento para Excel y Word (producción / desarrollo)
taskpane.html, src/taskpane/      Panel: chat, adjuntos, plan, tarjetas VBA, ajustes
src/agent/                        Bucle con la API de Claude, prompts de sistema, perfil por aplicación
src/tools/excel/, src/tools/word/ Herramientas Office.js de cada aplicación
src/excel/, src/word/             Helpers (direcciones, series, estilos, OOXML, Office.js)
tests/                            Pruebas (vitest)
practicas/                        Material de prueba local (no se sube a GitHub)
```

## Solución de problemas

- **El botón no aparece**: cierra Excel o Word por completo (también en segundo plano) y vuelve a abrirlo.
  En Mac, entra a Insertar > Complementos > Mis complementos.
- **No aparece en Word después de actualizar el repo**: vuelve a ejecutar el instalador (en Mac copia el
  manifiesto también a la carpeta de Word).
- **El panel no se actualiza tras publicar**: Office guarda caché. Windows: cierra Excel y Word y borra
  `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\`. Mac: borra
  `~/Library/Containers/com.microsoft.Excel/Data/Library/Caches/` (o `com.microsoft.Word`) y reabre la aplicación.
- **"La API key no es válida"**: revisa la key en Ajustes y que tenga crédito en la consola.
