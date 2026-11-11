import { describeOfficeError } from "../../excel/office";
import { createToolset } from "../toolset";
import { entregarVba } from "../vba";
import {
  configurarHoja,
  crearTabla,
  filtrar,
  formatoCondicional,
  insertarEliminar,
  quitarDuplicados,
  tablaDinamica,
} from "./advanced";
import { crearGrafico } from "./chart";
import { gestionarFormas, segmentacion } from "./dashboard";
import { limpiarDatos, modificarTabla } from "./data";
import { formatearRango } from "./format";
import { ejecutarOfficeJs } from "./misc";
import { macroModeloDatos } from "./model";
import { leerLibro, leerRango } from "./read";
import { crearNombre, gestionarHojas, ordenar, validacionDatos } from "./structure";
import { copiarRango, dividirTexto, escribirRango, rellenarSerie } from "./write";

export const EXCEL_TOOLSET = createToolset(
  [
    leerLibro,
    leerRango,
    escribirRango,
    formatearRango,
    crearNombre,
    rellenarSerie,
    dividirTexto,
    copiarRango,
    gestionarHojas,
    validacionDatos,
    ordenar,
    insertarEliminar,
    quitarDuplicados,
    limpiarDatos,
    formatoCondicional,
    crearTabla,
    modificarTabla,
    tablaDinamica,
    segmentacion,
    filtrar,
    configurarHoja,
    crearGrafico,
    gestionarFormas,
    macroModeloDatos,
    entregarVba,
    ejecutarOfficeJs,
  ],
  describeOfficeError,
);
