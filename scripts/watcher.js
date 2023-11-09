const chokidar = require('chokidar');
const { spawn } = require('child_process');
const path = require('path');

// Ruta del script de importación
const importScriptPath = path.join(__dirname, 'import.js');

// Carpeta a observar
const carpetaAObservar = 'vba-files/Module';

// Configuración de chokidar
const watcher = chokidar.watch(carpetaAObservar, {
  ignored: /(^|[/\\])\../, // Ignorar archivos ocultos
  persistent: true,
});

console.log(`Watching changes in VBA Folder: ${carpetaAObservar} 🧐`);

// Evento de cambio detectado
watcher
  .on('add', (ruta) => ejecutarImportScript(ruta))
  .on('change', (ruta) => ejecutarImportScript(ruta))
  .on('unlink', (ruta) => console.log(`Deleted File: ${ruta} 👻`));

// Función para ejecutar el script de importación
function ejecutarImportScript(ruta) {
  const nombreArchivo = path.basename(ruta);
  const comando = `npm run import ${nombreArchivo}`;

  console.log(`Change detected in ${nombreArchivo.split('.')[0]} 😝`);
  console.log(`Importing changes in source 🔧`);

  const proceso = spawn('node', [importScriptPath, nombreArchivo]);

  proceso.stdout.on('data', (data) => console.log(data.toString()));
  proceso.stderr.on('data', (data) => console.error(data.toString()));
}
