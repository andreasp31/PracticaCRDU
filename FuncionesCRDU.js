"use strict";

const prompt = require("prompt-sync")();
const XLSX = require("xlsx");

const fs = require("fs");

let nombre_hoja;
let nombre_excel;

const datosjson = [
    {   
        "id": 1, 
        "nombre": "Andrea", 
        "edad": 22, 
        "ciudad": "Ourense"
    },
    {   
        "id": 2, 
        "nombre": "Jacobo", 
        "edad": 24, 
        "ciudad": "Vigo"
    },
    {   
        "id": 3, 
        "nombre": "Santiago", 
        "edad": 26, 
        "ciudad": "Vigo"
    }
];


function menu(){
    console.log("--------TRABAJAR CON EXCEL--------");
    console.log("1. Crear un archivo de excel");
    console.log("2. Leer un excel");
    console.log("3. Eliminar un excel");
    console.log("4. Editar un excel");
    console.log("5. Salir de la aplicación");
    console.log("..................................");
}

menu();
seleccionar_menu();

function seleccionar_menu(){
    let opcion = parseInt(prompt("Selecciona una opción: "));
    switch(opcion){
        case 1:{
            console.clear();
            console.log("-------CREAR UN ARCHIVO DE EXCEL-------");
            crear_excel();
            otra_accion();
            break;
        }
        case 2:{
            console.clear();
            console.log("-------LEER UN ARCHIVO DE EXCEL-------");
            leer_excel();
            otra_accion();
            break;
        }
        case 3:{
            console.clear();
            console.log("-------BORRAR UN ARCHIVO DE EXCEL-------");
            borrar_excel();
            otra_accion();
            break;
        }
        case 4:{
            console.clear();
            console.log("-------EDITAR UN ARCHIVO DE EXCEL-------");
            editar_excel();
            otra_accion();
            break;
        }
        case 5:{
            console.clear();
            console.log("Has salido de la aplicacón");
            break;
        }
        default:{
            console.log("Opción inválida prueba de nuevo.");
            menu();
            seleccionar_menu();
        }
    }
}

function otra_accion(){
    console.log("----------------------------------------");
    console.log("Quieres realizar otra acción?: Si o No");
    let mas = prompt("");
    if(mas.toLowerCase() == "si"){
        console.clear();
        menu();
        seleccionar_menu();
    }
    else if(mas.toLowerCase() == "no"){
        console.clear();
        console.log("Has salido de la aplicación.");
    }
    else{
        otra_accion();
    }
}

function crear_excel(){
    const crear = XLSX.utils.book_new();
    const datos = XLSX.utils.json_to_sheet(datosjson);
    nombre_hoja = prompt("Introduce el nombre de la hoja: ");
    nombre_excel = prompt("Introduce el nombre del archivo: ");
    XLSX.utils.book_append_sheet(crear,datos,nombre_hoja);
    XLSX.writeFile(crear,"Excel/"+nombre_excel+".xlsx");
}

function leer_excel(){
    nombre_excel = prompt("Introduce el nombre del archivo: ");
    const leer = XLSX.readFile("Excel/"+nombre_excel+".xlsx");
    const hoja = leer.SheetNames[0];
    const archivo = leer.Sheets[hoja];
    const convertir = XLSX.utils.sheet_to_json(archivo);
    console.log("Lista de personas:");
        convertir.forEach(p => {
            console.log(`${p.id}. ${p.nombre} - ${p.edad} - ${p.ciudad}`);
        });
}

function borrar_excel(){
    nombre_excel = prompt("Introduce el nombre del archivo: ");
    fs.unlinkSync("Excel/"+nombre_excel+".xlsx");
    console.log("Archivo eliminado exitosamente.")
    
}

function editar_excel(){

    class estudiante{
        constructor(id,nombre,edad,ciudad){
            this.id = id;
            this.nombre = nombre;
            this.edad = edad;
            this.ciudad = ciudad;
        }

        nuevos_datos(){
            let id = (datosjson.length)+1;
            nombre_excel = prompt("Introduce el nombre del archivo: ");
            console.log("---------------------------------");
            console.log("Nuevo estudiante:");
            let nombre = prompt("Introduce el nombre del estudiante: ");
            let edad = parseInt(prompt("Introduce la edad: "));
            let ciudad = prompt("Introduce la ciudad: ");
            let nuevo_estudiante = new estudiante(id,nombre,edad,ciudad);

            const leer = XLSX.readFile("Excel/"+nombre_excel+".xlsx");
            const hoja = leer.SheetNames[0];
            const archivo = leer.Sheets[hoja];

            XLSX.utils.sheet_add_json(archivo, [nuevo_estudiante],{ origin: -1 });
            XLSX.writeFile(leer,"Excel/"+nombre_excel+".xlsx");
            console.log("Datos añadidos correctamente."); 
        }

        borrar_datos(){
            leer_excel();
            let borrar = parseInt(prompt("Selecciona el id del alumno a borrar: "));
        }

        modificar_datos(){
            nombre_excel = prompt("Introduce el nombre del excel: ");
            console.log("\n");
            const archivo = `Excel/${nombre_excel}.xlsx`;
            const leer = XLSX.readFile(archivo);


            console.log("Hojas disponibles: ", leer.SheetNames.join(", "));


            nombre_hoja = prompt("Introduce el nombre de la hoja: ");
            console.log("\n");
            const trabajo = leer.Sheets[nombre_hoja];


            let convertir_datos = XLSX.utils.sheet_to_json(trabajo, {defval: 1});
            console.log("\n");
            console.log("Lista de personas:");
            convertir_datos.forEach(p => {
                console.log(`${p.id}. ${p.nombre} - ${p.edad} - ${p.ciudad}`);
            });


            // Aqui leer el archivo


            let nombre =(prompt("Introduce el nombre de la persona a editar: "));
            console.log("\n");


            let persona_encontrada = false;
            for (let i = 0; i < convertir_datos.length; i++) {
                if (convertir_datos[i].nombre === nombre) {
                    persona_encontrada = true;


                    const campo = prompt("¿Qué quieres editar? (nombre / edad / ciudad): ");
                    const nuevoValor = prompt("Nuevo valor para ",campo,": ");
                    datos[i][campo] = nuevoValor;


                    break;
                }
            }


            const nueva_hoja = XLSX.utils.json_to_sheet(convertir_datos);
            leer.Sheets[nombre_hoja] = nueva_hoja;


            //Guardar el archivo
            XLSX.writeFile(leer,archivo);
            console.log("Datos modificados correctamente.");
        }

    }

    const gestion = new estudiante();

    function menu(){
        console.log("Qué quieres editar?","\n");
        console.log("1. Añadir datos al excel");
        console.log("2. Modificar datos del excel");
        console.log("3. Borrar datos del excel");
    }
    function seleccionar_menu(){
        let opcion = parseInt(prompt("Selecciona la opción deseada: "));
        switch(opcion){
            case 1:{
                console.clear();
                console.log("--Añadir datos al excel--");
                gestion.nuevos_datos();
                break;
            }
            case 2:{
                console.clear();
                console.log("--Modificar datos al excel--");
                gestion.modificar_datos();
                break;
            }
            case 3:{
                console.clear();
                console.log("--Borrar datos al excel--");
                gestion.borrar_datos();
                break;
            }
            default:{
                console.log("No ha seleccionado una opción valida, vuelva a intentarlo");
                editar_excel();
                break;
            }
        }
    }

    menu();
    seleccionar_menu();
}