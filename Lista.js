const fs = require('fs');
const XLSX = require('xlsx');
const {Lista_hi, Lista_conectadas,Lista_cierre_de_brechas, Apptitude_Mt,Igualdapp_Mt,Igualdapp_cierre_de_brechas}=require('./Prefijos_a_eliminar')
// Función para convertir números de serie de Excel a fechas, solo para la columna 'Birthdate'
function convertirNumeroASFecha(celda, columnaNombre) {
    if (columnaNombre === 'Birthdate' && typeof celda === 'number' && celda > 0) {
        const fechaBase = new Date(1899, 11, 30);
        const fecha = new Date(fechaBase.getTime() + (celda * 86400000));
        return fecha.toLocaleDateString(); // Devuelve la fecha en formato local
    }
    return celda; // Retorna la celda sin cambios si no es la columna 'Birthdate'
}

// Función para mover una columna a una posición específica
function moverColumna(datos, columnaNombre, nuevaPosicion) {
    const encabezados = datos[0];
    const indiceColumna = encabezados.indexOf(columnaNombre);

    if (indiceColumna === -1) {
        throw new Error(`La columna "${columnaNombre}" no se encontró.`);
    }

    const nuevaPosicionIndex = nuevaPosicion - 1;

    if (nuevaPosicionIndex >= encabezados.length) {
        throw new Error(`La nueva posición (${nuevaPosicion}) es mayor que el número de columnas.`);
    }

    // Mueve la columna a la nueva posición
    const encabezadoMovido = encabezados.splice(indiceColumna, 1)[0];
    encabezados.splice(nuevaPosicionIndex, 0, encabezadoMovido);

    datos.forEach((fila, index) => {
        if (index === 0) return; // No modificar los encabezados
        const valorMovido = fila.splice(indiceColumna, 1)[0];
        fila.splice(nuevaPosicionIndex, 0, valorMovido);
    });

    return datos;
}

// Función para ordenar alfabéticamente las columnas después de una columna específica
function ordenarColumnasDespues(datos, columnaNombre) {
    const encabezados = datos[0];
    const indiceColumna = encabezados.indexOf(columnaNombre);

    if (indiceColumna === -1) {
        throw new Error(`La columna "${columnaNombre}" no se encontró.`);
    }

    const columnasDespues = encabezados.slice(indiceColumna + 1).map(String);
    const indicesColumnasDespues = encabezados.map((nombre, index) => index)
        .filter(index => index > indiceColumna);

    const columnasOrdenadas = [...columnasDespues].sort((a, b) => a.localeCompare(b));

    const nuevosIndicesColumnas = columnasOrdenadas.map(nombre => encabezados.indexOf(nombre));
    const indicesColumnasReordenadas = [...indicesColumnasDespues]
        .sort((a, b) => columnasDespues.indexOf(encabezados[a]) - columnasDespues.indexOf(encabezados[b]));

    const nuevosEncabezados = encabezados.slice(0, indiceColumna + 1)
        .concat(columnasOrdenadas);

    datos.forEach((fila, index) => {
        if (index === 0) return; // No modificar los encabezados
        const filaReordenada = fila.slice(0, indiceColumna + 1)
            .concat(indicesColumnasReordenadas.map(i => fila[i]));
        datos[index] = filaReordenada;
    });

    datos[0] = nuevosEncabezados;

    return datos;
}

// Función para eliminar columnas que comienzan con ciertos prefijos
function eliminarColumnasConPrefijos(datos, columnaNombre, prefijos) {
    const encabezados = datos[0];
    const indiceColumna = encabezados.indexOf(columnaNombre);

    if (indiceColumna === -1) {
        throw new Error(`La columna "${columnaNombre}" no se encontró.`);
    }

    const indicesAEliminar = encabezados.map((nombre, index) => index)
        .filter(indice => {
            if (indice > indiceColumna) {
                return prefijos.some(prefijo => encabezados[indice].startsWith(prefijo));
            }
            return false;
        });

    if (indicesAEliminar.length > 0) {
        datos.forEach((fila, index) => {
            if (index === 0) {
                indicesAEliminar.sort((a, b) => b - a).forEach(indice => fila.splice(indice, 1));
            } else {
                indicesAEliminar.sort((a, b) => b - a).forEach(indice => fila.splice(indice, 1));
            }
        });
    }

    return datos;
}

// Función para procesar el archivo de Excel según las especificaciones
function procesarArchivoExcel(nombreArchivo, prefijosAEliminar, nombreArchivoModificado) {
    if (!fs.existsSync(nombreArchivo)) {
        throw new Error(`El archivo "${nombreArchivo}" no existe.`);
    }

    const libro = XLSX.readFile(nombreArchivo);
    const hojaOriginal = libro.Sheets[libro.SheetNames[0]];
    let datos = XLSX.utils.sheet_to_json(hojaOriginal, { header: 1 });

    const encabezados = datos[0];
    const datosFiltrados = datos.slice(1).filter(row => {
        const nombre = String(row[0] || ''); // Asumiendo que el nombre está en la primera columna
        const apellido = String(row[1] || ''); // Asumiendo que el apellido está en la segunda columna
        return !nombre.toLowerCase().includes('Test') && !apellido.toLowerCase().includes('Test');
    });

    // Convertir los números de serie a fechas solo en la columna 'Birthdate'
    const datosConvertidos = datosFiltrados.map(row =>
        row.map((celda, colIndex) => convertirNumeroASFecha(celda, encabezados[colIndex]))
    );

    datosConvertidos.unshift(encabezados.map((celda, colIndex) => {
        return convertirNumeroASFecha(celda, encabezados[colIndex]);
    }));

    const datosConColumnaMovida = moverColumna(datosConvertidos, 'logged_at', 14);
    const datosOrdenados = ordenarColumnasDespues(datosConColumnaMovida, 'logged_at');

    const datosFinales = eliminarColumnasConPrefijos(datosOrdenados, 'logged_at', prefijosAEliminar);

    // Transponer los datos
    const datosTranspuestos = datosFinales[0].map((_, colIndex) =>
        datosFinales.map(row => row[colIndex])
    );

    const hojaTranspuesta = XLSX.utils.aoa_to_sheet(datosTranspuestos, { origin: 'A1' });

    const nuevoLibro = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(nuevoLibro, XLSX.utils.aoa_to_sheet(datosFinales, { origin: 'A1' }), 'Hoja Procesada');
    XLSX.utils.book_append_sheet(nuevoLibro, hojaTranspuesta, 'Hoja Transpuesta');

    // Guardar el archivo modificado
    XLSX.writeFile(nuevoLibro, nombreArchivoModificado);
}

// Llamada a la función con parámetros
const nombreArchivo = 'Reporte/2024-01-30 - 2024-07-22 mujeres.xlsx'

const nombreArchivoModificado = 'datos_procesados_Igualdapp_mujeres.xlsx';

procesarArchivoExcel(nombreArchivo, Igualdapp_Mt, nombreArchivoModificado);





