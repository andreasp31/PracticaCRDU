"use strict";

const prompt = require("prompt-sync")();
const XLSX = require("xlsx");

const fs = require("fs");

let nombre_hoja;
let nombre_excel;

//Datos que se van a añadir automaticamente al crear el archivo excel
const datosjson = [
    {  
        "id": 1,
        "nombre": "Andrea",
        "edad": "22",
        "ciudad": "Vigo"
    },
    {  
        "id": 2,
        "nombre": "Jacobo",
        "edad": "24",
        "ciudad": "Vigo"
    },
    {   
        "id": 3, 
        "nombre": "Santiago", 
        "edad": 26, 
        "ciudad": "Vigo"
    }
];

//Clase para crear los objetos que se van a añadir, editar y modificar en el excel
class estudiante{
    constructor(id,nombre,edad,ciudad){
        this.id = id;
        this.nombre = nombre;
        this.edad = edad;
        this.ciudad = ciudad;
    }

    //Añadir datos al excel preguntando por terminal los datos
    nuevos_datos(){
        //Datos nuevo usuario
        let id = (datosjson.length)+1;
        nombre_excel = prompt("Introduce el nombre del archivo: ");
        console.log("---------------------------------");
        console.log("Nuevo estudiante:");
        let nombre = prompt("Introduce el nombre del estudiante: ");
        let edad = parseInt(prompt("Introduce la edad: "));
        let ciudad = prompt("Introduce la ciudad: ");
        let nuevo_estudiante = new estudiante(id,nombre,edad,ciudad);
    
        //Se lee el archivo excel y se selecciona
        const leer = XLSX.readFile("Excel/"+nombre_excel+".xlsx");
        const hoja = leer.SheetNames[0];
        const archivo = leer.Sheets[hoja];
    
        //Ese estudiante se añade al json de datos y se escriben en el excel
        XLSX.utils.sheet_add_json(archivo,[nuevo_estudiante],{ origin: -1 });
        XLSX.writeFile(leer,"Excel/"+nombre_excel+".xlsx");
        console.log("Datos añadidos correctamente."); 
        
        //Borrar cabecera de los objetos que se añadan
        

    }

    //Modificar los datos seleccionando el nombre del usuario de cada fila
    modificar_datos(){
        //Seleccionar el nombre del archivo de excel
        nombre_excel = prompt("Introduce el nombre del excel: ");
        console.log("\n");
        const archivo = `Excel/${nombre_excel}.xlsx`;
        const leer = XLSX.readFile(archivo);
        //Enseña las hojas que tiene ese excel para seleccionar una mediante el nombre
        console.log("Hojas disponibles: ", leer.SheetNames.join(", "));

        nombre_hoja = prompt("Introduce el nombre de la hoja: ");
        console.log("\n");
        //Se lee el archivo excel
        const trabajo = leer.Sheets[nombre_hoja];

        //Convertimos los datos a un json para trabajar con ellos
        let convertir_datos = XLSX.utils.sheet_to_json(trabajo, {defval: 1});
        console.log("\n");
        //Primero enseño la lista de los datos del excel
        console.log("Lista de personas:");
        convertir_datos.forEach(p => {
            console.log(`${p.id}. ${p.nombre} - ${p.edad} - ${p.ciudad}`);
        });

        //Seleccionamos mediante el nombre la persona que queremos editar
        let nombre =(prompt("Introduce el nombre de la persona a editar: "));
        console.log("\n");

        //Recorremos el archivo convertido a json para seleccionar la persona a editar
        let persona_encontrada = false;
        for (let i = 0; i < convertir_datos.length; i++) {
            if (convertir_datos[i].nombre === nombre) {
                persona_encontrada = true;
                //Al tener a la persona preguntamos que se quiere cambiar
                const campo = prompt("¿Qué quieres editar? (nombre / edad / ciudad): ");
                const nuevoValor = prompt("Nuevo valor: ");
                convertir_datos[i][campo] = nuevoValor;
                break;
            }
        }

        //Ahora esos datos se vuelven a convertir para ir al excel
        const nueva_hoja = XLSX.utils.json_to_sheet(convertir_datos);
        leer.Sheets[nombre_hoja] = nueva_hoja;

        //Se guarda el archivo
        XLSX.writeFile(leer,archivo);
        console.log("Datos modificados correctamente.");
    }

    //Borrar una fila seleccionando el id de la fila
    borrar_datos(){
        //Seleccionar el nombre del archivo de excel
        nombre_excel = prompt("Introduce el nombre del excel: ");
        console.log("\n");
        const archivo = `Excel/${nombre_excel}.xlsx`;
        const leer = XLSX.readFile(archivo);

        //Enseña las hojas que tiene ese excel para seleccionar una mediante el nombre
        console.log("Hojas disponibles: ", leer.SheetNames.join(", "));

        nombre_hoja = prompt("Introduce el nombre de la hoja: ");
        console.log("\n");
        const trabajo = leer.Sheets[nombre_hoja];

        //Convertimos los datos a un json para trabajar con ellos
        let convertir_datos = XLSX.utils.sheet_to_json(trabajo, {header: 1});

        //Seleccionamos mediante el nombre la persona que queremos liminar
        let indice = parseInt(prompt("Introduce el índice de la persona a eliminar: "));
        console.log("\n");
        
        //Borramos los datos de ese json
        convertir_datos.splice(indice,1);

        //Se crea una nueva hoja y se sobreescribe sobre la otra (igualando la vieja por la nueva con los datos borrados)
        const nueva_hoja = XLSX.utils.aoa_to_sheet(convertir_datos);
        leer.Sheets[nombre_hoja] = nueva_hoja;

        //Se guarda el archivo
        XLSX.writeFile(leer,archivo);
        console.log("Datos eliminados correctamente.");
    }
}

const gestion = new estudiante();

//Menú principal de la aplicación
function menu(){
    console.log("--------TRABAJAR CON EXCEL--------");
    console.log("1. Crear un archivo de excel");
    console.log("2. Leer un excel")
    console.log("3. Eliminar un excel");
    console.log("4. Editar un excel");
    console.log("5. Salir de la aplicación")
    console.log("..................................")
}

//Aquí lleva a las funciones principales que realizan las acciones seleccionadas
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

//Función para crear el excel usando la librería XLSX
function crear_excel(){
    const crear = XLSX.utils.book_new();
    //convierte los datos que tengo arriba del todo a un tipo que puede entrar en el excel
    const datos = XLSX.utils.json_to_sheet(datosjson);
    //Datos del excel
    nombre_hoja = prompt("Introduce el nombre de la hoja: ");
    nombre_excel = prompt("Introduce el nombre del archivo: ");
    //Crea el archivo en la ruta seleccionada y añade los datos anteriores
    XLSX.utils.book_append_sheet(crear,datos,nombre_hoja);
    XLSX.writeFile(crear,"Excel/"+nombre_excel+".xlsx");
}

//Función para leer el excel usando la librería XLSX
function leer_excel(){
    //Selecciono el archivo de excel a leer
    nombre_excel = prompt("Introduce el nombre del excel: ");
        console.log("\n");
        const archivo = `Excel/${nombre_excel}.xlsx`;
        const leer = XLSX.readFile(archivo);

        //Enseño las hojas que hay de ese archivo y la selecciono
        console.log("Hojas disponibles: ", leer.SheetNames.join(", "));

        nombre_hoja = prompt("Introduce el nombre de la hoja: ");
        console.log("\n");
        const trabajo = leer.Sheets[nombre_hoja];

        //Se pasan los datos de excel a json
        let convertir_datos = XLSX.utils.sheet_to_json(trabajo, {defval: 1});
        //Y enseño con un forEach los datos de ese excel
        console.log("Lista de personas:");
        convertir_datos.forEach(p => {
            console.log(`${p.id}. ${p.nombre} - ${p.edad} - ${p.ciudad}`);
        });
}

//Función para borrar el excel usando la librería XLSX
function borrar_excel(){
    //Seleccionamos el nombre del archivo
    nombre_excel = prompt("Introduce el nombre del archivo: ");
        //Y con el el método unlinkSync lo desvinculo de la carpeta por lo que se elimina ¡
        fs.unlinkSync("Excel/"+nombre_excel+".xlsx");
        console.log("Archivo eliminado exitosamente.");
}

//Función para editar el excel que permite (añadir datos, modificar datos, eliminar datos)
function editar_excel(){
    console.log("Que quieres editar?")
    console.log("1. Añadir datos en el excel");
    console.log("2. Modificar los datos ya exisstentes excel");
    console.log("3. Eliminar datos del excel");
    console.log("..................................");
    let opcion = parseInt(prompt("Selecciona una opción: "));

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
            console.log("--Eliminar datos al excel--");
            gestion.borrar_datos();
            break;
        }
        default:{
            console.log("No has seleccionado una de las opciones, vuelve a intentarlo.");
            editar_excel();
        }
    }
}

//Función que al acabar cada apartado pregunta si quieres realizar otra y ahí se sale o no de la aplicación
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

menu();
seleccionar_menu();
