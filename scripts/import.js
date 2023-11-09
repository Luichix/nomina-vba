const fs = require('fs');
const path = require('path');

function buscarArchivoRecursivo(nombreArchivo, directorio) {
  const archivos = fs.readdirSync(directorio);

  for (const archivo of archivos) {
    const rutaCompleta = path.join(directorio, archivo);
    const stats = fs.statSync(rutaCompleta);

    if (stats.isDirectory()) {
      const resultadoRecursivo = buscarArchivoRecursivo(
        nombreArchivo,
        rutaCompleta
      );
      if (resultadoRecursivo) {
        return resultadoRecursivo;
      }
    } else if (archivo === nombreArchivo) {
      return rutaCompleta;
    }
  }

  return null;
}

function copiarYpegar(origen, destino) {
  try {
    const contenido = fs.readFileSync(origen);
    const nombreArchivo = path.basename(origen);
    const destinoCompleto = path.join(destino, nombreArchivo);
    fs.writeFileSync(destinoCompleto, contenido);
    return { success: true, message: 'Changes made successfully! 🥳' };
  } catch (error) {
    console.error('Error to copy file ☠️:', error.message);
    // Devuelve un objeto con información de error
    return { success: false, message: `Error: ${error.message}` };
  }
}

// Obtén el nombre del archivo desde los argumentos de la línea de comandos
const args = process.argv.slice(2);
const nombreArchivo = args[0];

// Ruta de origen (búsqueda recursiva)
const rutaOrigen = buscarArchivoRecursivo(
  nombreArchivo,
  path.join('vba-files', 'Module')
);

if (rutaOrigen) {
  // Ruta de destino
  const destino = 'vba-build';

  // Llamar a la función para copiar y pegar
  const resultadoCopia = copiarYpegar(rutaOrigen, destino);
  if (resultadoCopia.success) {
    process.stdout.write(JSON.stringify(resultadoCopia.message));
  } else {
    process.stderr.write(JSON.stringify(resultadoCopia.message));
  }
} else {
  console.error(`File ${nombreArchivo} no found. 🤡`);
}
