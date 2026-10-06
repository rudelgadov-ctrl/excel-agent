# CLAUDE.md

Complemento de Office (Excel) con un panel de chat que llama a la API de Claude desde el navegador
y ejecuta herramientas Office.js sobre el libro abierto. TypeScript + Vite; publicado en GitHub Pages
(`https://rudelgadov-ctrl.github.io/excel-agent/`) por `.github/workflows/deploy.yml` en cada push a `main`.

## Comandos

- `npm test`: vitest (helpers puros, esquemas de herramientas, bucle, markdown).
- `npm run build`: `tsc --noEmit` + `vite build` a `dist/` (base `/excel-agent/`).
- `npm run dev`: panel en `https://localhost:3000` (usa `manifest.dev.xml`). Requiere `npm run certs` una vez.
- `npm run validate`: valida ambos manifiestos (servicio en línea de Microsoft).

## Arquitectura

- `src/agent/loop.ts`: bucle manual con `client.beta.messages.stream` + `finalMessage()`,
  `dangerouslyAllowBrowser`, pensamiento adaptativo con `display: "updates"`, `fallbacks: "default"`,
  caché automática. El historial es solo de agregado: no editar mensajes anteriores (pensamiento preservado).
- Aprobación del plan: las herramientas con `writes: true` devuelven `BLOQUEADO` hasta que el usuario
  aprueba (o modo automático). El estado viaja como nota al final de cada mensaje de usuario
  (`STATUS` en `systemPrompt.ts`), no en el prompt de sistema, para no romper la caché.
- `src/tools/`: cada herramienta es `defineTool({ name, description, schema (zod), writes, run })`.
  El JSON Schema para la API sale de `z.toJSONSchema(..., { io: "input" })` en `tools/index.ts`.
  Registrar herramientas nuevas en `TOOLS`.
- `src/excel/`: helpers puros testeables (`address`, `series`, `text`) y helpers Office.js (`office.ts`).

## Convenciones

- Interfaz, nombres de herramientas y parámetros en español; valores de enums de Office.js en inglés.
- Las fórmulas se escriben en inglés con `range.formulas` (Excel las traduce); los formatos numéricos usan códigos en-US.
- Funciones que dependen de requirement sets nuevos: comprobar con `supports(n)` y ofrecer alternativa.
- No llamar a `getHttpsServerOptions()` de `office-addin-dev-certs` desde la config: instala certificados
  y abre un diálogo del sistema. `vite.config.ts` solo lee los archivos existentes.
- `practicas/` (PDF y plantillas del curso) está en `.gitignore`: no publicarlo (repo público).
