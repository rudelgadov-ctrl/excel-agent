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
- Lo que la API de Office.js no permite (macros, Buscar objetivo, Solver, tablas de datos, minigráficos;
  en Word combinación de correspondencia, SmartArt, marca de agua…) se resuelve con alternativa + VBA o
  pasos manuales; nunca se omite en silencio.

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
- Aprobación del plan: las herramientas con `writes: true` devuelven `BLOQUEADO` hasta que el usuario
  aprueba (o modo automático). El estado viaja como nota al final de cada mensaje de usuario
  (`STATUS` en `systemPrompt.ts`), no en el prompt de sistema, para no romper la caché.
- `systemPrompt.ts`: `EXCEL_SYSTEM_PROMPT` y `WORD_SYSTEM_PROMPT` comparten los pasos de aprobación,
  cierre y estilo. Cada prompt debe nombrar todas sus herramientas y ninguna de la otra aplicación (lo prueba
  `tests/agent.test.ts`).
- Adjuntos (`src/taskpane/attachments.ts`): PDF e imágenes van en base64; Word (.docx, con fflate) y
  texto se envían como documentos de texto. Cada imagen va precedida de su nombre y el panel la guarda en
  la sesión (`ToolEnv.images()`) para que `insertar_imagen` (Word) pueda insertarla.
- `src/tools/`: cada herramienta es `defineTool({ name, description, schema (zod), writes, run })`.
  `createToolset` (`tools/toolset.ts`) arma el JSON Schema para la API con `z.toJSONSchema(..., { io: "input" })`,
  valida con Zod y aplica el bloqueo de escritura. Registrar herramientas nuevas en `tools/excel/index.ts`
  (`EXCEL_TOOLSET`) o `tools/word/index.ts` (`WORD_TOOLSET`). `entregar_vba` (`tools/vba.ts`) es compartida.
- Herramientas de Word: casi todas reciben `objetivo` (`tools/word/common.ts`: zona cuerpo/encabezado/pie,
  párrafos por número de `leer_documento`, tabla, texto buscado, selección), que `resolveTarget`
  (`src/word/office.ts`) convierte en rangos. Las unidades son cm (sangrías, márgenes) y puntos (fuente, espaciado).
- `src/excel/`: helpers puros testeables (`address`, `series`, `text`) y helpers Office.js (`office.ts`).
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
- `practicas/` (PDF y plantillas del curso) está en `.gitignore`: no publicarlo (repo público).
