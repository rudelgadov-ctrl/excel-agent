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
import { formatearRango } from "./format";
import { ejecutarOfficeJs } from "./misc";
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
    formatoCondicional,
    crearTabla,
    tablaDinamica,
    filtrar,
    configurarHoja,
    crearGrafico,
    entregarVba,
    ejecutarOfficeJs,
  ],
  describeOfficeError,
);
