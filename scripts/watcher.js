const chokidar = require('chokidar');
const { spawn } = require('child_process');
const path = require('path');

let primerArranque = true;

// Ruta del script de importación
const importScriptPath = path.join(__dirname, 'import.js');

// Carpeta a observar
const carpetaAObservar = 'vba-files/Module';

// Configuración de chokidar
const watcher = chokidar.watch(carpetaAObservar, {
  ignored: /(^|[/\\])\../, // Ignorar archivos ocultos
  persistent: true,
});
if (primerArranque) {
  console.log(`Importing files to VBA Build!!! 🚀`);
  console.log(`Watching changes in VBA Folder: ${carpetaAObservar} 🧐`);
}
// Evento de cambio detectado
watcher
  .on('add', (ruta) => ejecutarImportScript(ruta))
  .on('change', (ruta) => ejecutarImportScript(ruta))
  .on('unlink', (ruta) => console.log(`Deleted File: ${ruta} 👻`));

// Función para ejecutar el script de importación
function ejecutarImportScript(ruta) {
  const nombreArchivo = path.basename(ruta);
  const comando = `npm run import ${nombreArchivo}`;
  if (!primerArranque) {
    console.log(`Change detected in ${nombreArchivo.split('.')[0]} 😝`);
    console.log(`Importing changes to vba build 🔧`);
  }
  const proceso = spawn('node', [importScriptPath, nombreArchivo]);

  if (!primerArranque) {
    proceso.stdout.on('data', (data) => {
      if (data) {
        console.log(data.toString());
      }
    });
  }
  proceso.stderr.on('data', (data) => {
    if (data) {
      console.error(data.toString());
    }
  });
}

// Marcar el primer arranque como completado después de un breve retraso
setTimeout(() => {
  primerArranque = false;
}, 3000);
