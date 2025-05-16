"use strict";

const prompt = require("prompt-sync")();
const XLSX = require("xlsx");

const fs = require("fs");
const ruta = "Datos/datos.json";
const datosjson = fs.readFileSync(ruta,"utf8");
const parsear= JSON.parse(datos);


const traer_seleccion = require("FuncionesCRDU");
const llamar_seleccion = traer_seleccion();
