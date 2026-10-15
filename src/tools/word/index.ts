import { describeWordError } from "../../word/office";
import { createToolset } from "../toolset";
import { entregarVba } from "../vba";
import {
  eliminarContenido,
  insertarCampo,
  insertarContenido,
  insertarImagen,
  insertarSalto,
  insertarVinculoONota,
} from "./content";
import { aplicarLista, formatoParrafo, formatoTexto, modificarEstilo } from "./format";
import { configurarPagina } from "./layout";
import { ejecutarOfficeJs } from "./misc";
import { leerDocumento, leerTabla } from "./read";
import { editarTabla, formatoTabla, insertarTabla } from "./table";

export const WORD_TOOLSET = createToolset(
  [
    leerDocumento,
    leerTabla,
    insertarContenido,
    eliminarContenido,
    formatoTexto,
    formatoParrafo,
    aplicarLista,
    modificarEstilo,
    insertarTabla,
    editarTabla,
    formatoTabla,
    configurarPagina,
    insertarSalto,
    insertarCampo,
    insertarVinculoONota,
    insertarImagen,
    entregarVba,
    ejecutarOfficeJs,
  ],
  describeWordError,
);
