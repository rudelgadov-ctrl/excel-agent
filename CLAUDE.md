# CLAUDE.md

Complemento de Office (Excel y Word) con un panel de chat que llama a la API de Claude desde el navegador
y ejecuta herramientas Office.js sobre el libro o documento abierto. Un solo manifiesto declara los hosts
`Workbook` y `Document`; el panel detecta la aplicación al arrancar y usa su perfil (`src/agent/hosts.ts`). TypeScript + Vite; publicado en GitHub Pages
(`https://rudelgadov-ctrl.github.io/excel-agent/`) por `.github/workflows/deploy.yml` en cada push a `main`.

## Requisito central: generalidad

El agente debe resolver **cualquier** práctica o requerimiento de Excel o Word, incluidas prácticas futuras que
nadie ha visto. Las prácticas de `practicas/` son solo ejemplos del tipo de trabajo, nunca el objetivo:

- No escribir lógica, ejemplos ni reglas atadas a una práctica concreta (nombres de hojas, celdas,
  macros, textos o datos de ejemplo). `tests/agent.test.ts` ("generalidad") falla si algún prompt o
  herramienta menciona una.
- Cuando una práctica nueva falle, corregir la capacidad general que faltó (una herramienta, una regla
  del prompt), no el caso particular.
- Lo que la API de Office.js no permite (macros, Buscar objetivo, Solver, tablas de datos, minigráficos,
  modelo de datos/Power Pivot y relaciones; en Word combinación de correspondencia, SmartArt, marca de agua…)
  se resuelve con alternativa + VBA o pasos manuales; nunca se omite en silencio.

## Comandos

- `npm test`: vitest (helpers puros, esquemas de herramientas, bucle, markdown).
- `npm run build`: `tsc --noEmit` + `vite build` a `dist/` (base `/excel-agent/`).
- `npm run dev`: panel en `https://localhost:3000` (usa `manifest.dev.xml`). Requiere `npm run certs` una vez.
- `npm run validate`: valida ambos manifiestos (servicio en línea de Microsoft). Si falla con
  "Cannot find module 'fs-extra'", es una dependencia faltante del validador: `npm install --no-save fs-extra`.

## Arquitectura

- `src/agent/hosts.ts`: perfil por aplicación (prompt de sistema + `Toolset`). `OfficeAgent` recibe el perfil;
  como no cambia durante la sesión, el prefijo cacheado (prompt + herramientas) es estable.
- `src/agent/loop.ts`: bucle manual con `client.beta.messages.stream` + `finalMessage()`,
  `dangerouslyAllowBrowser`, pensamiento adaptativo con `display: "updates"`, `fallbacks: "default"`,
  caché automática. El historial es solo de agregado: no editar mensajes anteriores (pensamiento preservado).
- Sin aprobación del plan: el agente siempre es automático (decisión del usuario); no hay ajuste ni botón
  «Aprobar». El bucle pasa `allowWrites: true` a `Toolset.execute`, que conserva el bloqueo `BLOQUEADO` para
  las herramientas `writes: true` por si algún día se reactiva. El prompt (`planStep`) pide anunciar en pocas
  líneas lo que se hará y seguir de inmediato, registrando las ambigüedades como supuestos.
- `src/taskpane/styles.css`: `[hidden]` lleva `display: none !important` porque los componentes con
  `display: flex` anulaban el atributo (Ajustes y la barra de acción no se ocultaban). Ajustes se superpone
  al área de la conversación (`.stage`) en vez de apilarse, para que la respuesta conserve todo el alto.
- `systemPrompt.ts`: `EXCEL_SYSTEM_PROMPT` y `WORD_SYSTEM_PROMPT` comparten los pasos de ejecución,
  cierre y estilo. Cada prompt debe nombrar todas sus herramientas y ninguna de la otra aplicación (lo prueba
  `tests/agent.test.ts`).
- Adjuntos (`src/taskpane/attachments.ts`): PDF e imágenes van en base64; Word (.docx, con fflate), Excel
  (.xlsx, valores por hoja con fechas ISO) y texto (UTF-8, UTF-16 o Windows-1252) se envían como documentos de texto. Cada imagen va precedida de su nombre y el panel la guarda en
  la sesión (`ToolEnv.images()`) para que `insertar_imagen` (Word) pueda insertarla.
- `src/tools/`: cada herramienta es `defineTool({ name, description, schema (zod), writes, run })`.
  `createToolset` (`tools/toolset.ts`) arma el JSON Schema para la API con `z.toJSONSchema(..., { io: "input" })`,
  valida con Zod y aplica el bloqueo de escritura. Registrar herramientas nuevas en `tools/excel/index.ts`
  (`EXCEL_TOOLSET`) o `tools/word/index.ts` (`WORD_TOOLSET`). `entregar_vba` (`tools/vba.ts`) es compartida.
- Herramientas de Word: casi todas reciben `objetivo` (`tools/word/common.ts`: zona cuerpo/encabezado/pie,
  párrafos por número de `leer_documento`, tabla, texto buscado, selección), que `resolveTarget`
  (`src/word/office.ts`) convierte en rangos. Las unidades son cm (sangrías, márgenes) y puntos (fuente, espaciado).
- `src/excel/`: helpers puros testeables (`address`, `series`, `text`, `clean` para limpiar_datos, `pivot` con
  las definiciones de tablas dinámicas) y helpers Office.js (`office.ts`).
- Modelo de datos: Office.js no tiene API para Power Pivot, relaciones, agrupar fechas en tablas dinámicas,
  columnas de segmentadores, conexiones de informe ni vincular formas a celdas. `macro_modelo_datos`
  (`tools/excel/model.ts`) valida contra el libro abierto y `src/excel/macro.ts` genera un módulo VBA
  determinista (un `PasoNN` por acción con su propio manejo de errores y un resumen final). Las llamadas al
  modelo de objetos se validaron en Excel 365 por COM; si se cambia el VBA generado, volver a probarlo así
  (requiere activar temporalmente "Confiar en el acceso al modelo de objetos de proyectos de VBA" y
  restaurarlo después de que Excel cierre, porque Excel reescribe ese valor al salir). Detalles aprendidos:
  una tabla que se llama igual que su hoja no se puede agregar al modelo (la macro renombra la hoja un
  momento); `GetMeasure` reutiliza la medida implícita existente, así que los campos de valores se
  referencian por posición (`pt.DataFields(k)`); las tablas del modelo creadas por VBA no traen subtotales.
- `src/word/`: helpers puros (`spec.ts`: índices, unidades, estilos integrados, códigos de campo; `ooxml.ts`)
  y helpers Office.js (`office.ts`).

## Convenciones

- Interfaz, nombres de herramientas y parámetros en español; valores de enums de Office.js en inglés.
- Las fórmulas se escriben en inglés con `range.formulas` (Excel las traduce); los formatos numéricos usan códigos en-US.
- Funciones que dependen de requirement sets nuevos: comprobar con `supports(n)` (Excel) o
  `wordSupports(n, set)` / `requireWord` (Word, incluido `WordApiDesktop`) y ofrecer alternativa.
  El manifiesto no exige versiones (para cargarse en ambas aplicaciones); el panel avisa si la aplicación
  no llega a ExcelApi 1.7 o WordApi 1.3.
- No llamar a `getHttpsServerOptions()` de `office-addin-dev-certs` desde la config: instala certificados
  y abre un diálogo del sistema. `vite.config.ts` solo lee los archivos existentes.
- `scripts/instalar-windows.ps1` debe seguir siendo ASCII puro y sin BOM: se ejecuta con `irm ... | iex` y un BOM
  (o un acento mal leído) rompe el comando en Windows PowerShell 5.1. Ambos scripts descargan el manifiesto de
  GitHub Pages si no se ejecutan desde un clon; la instalación publicada es un comando de una línea (ver README).
- `practicas/` (PDF y plantillas del curso) está en `.gitignore`: no publicarlo (repo público). El material
  nuevo a veces llega en otras carpetas de la raíz; `.gitignore` también excluye PDF, Excel, Word y .txt.
