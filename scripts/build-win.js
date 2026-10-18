const { spawnSync } = require('child_process');
const path = require('path');

const projectDirectory = path.resolve(__dirname, '..');
const outputDirectory = 'release';

/**
 * Ejecuta una herramienta Node local con el mismo runtime que inició la compilación.
 * @param {string} modulePath Módulo o CLI instalado en el proyecto.
 * @param {string[]} args Argumentos enviados sin interpolación de comandos de shell.
 * @returns {void} Retorna al completar la etapa correctamente.
 * @throws Si el proceso no puede iniciarse o termina con un código distinto de cero.
 */
function runTool(modulePath, args) {
  const result = spawnSync(process.execPath, [require.resolve(modulePath), ...args], {
    cwd: projectDirectory,
    stdio: 'inherit',
    shell: false,
    env: {
      ...process.env,
      PATH: `${path.dirname(process.execPath)}${path.delimiter}${process.env.PATH || ''}`,
    },
  });
  if (result.error) throw result.error;
  if (result.status !== 0) {
    throw new Error(`${modulePath} terminó con código ${result.status ?? 'desconocido'}.`);
  }
}

/**
 * Genera Electron, Angular y el instalador NSIS de Windows en orden.
 * @returns {void} Escribe el instalador en release; no recibe parámetros ni cierra otras aplicaciones.
 */
function buildWindows() {
  if (process.platform !== 'win32')
    throw new Error('El instalador de Windows debe compilarse en Windows.');

  console.log('Compilando Electron...');
  runTool('typescript/bin/tsc', ['-p', 'electron/tsconfig.json']);
  console.log('Compilando Angular...');
  runTool('@angular/cli/bin/ng.js', ['build', '--configuration', 'production']);
  console.log('Generando el instalador de Windows...');
  runTool('electron-builder/cli.js', [
    '--win',
    '--publish',
    'never',
    `--config.directories.output=${outputDirectory}`,
  ]);
  console.log(`Compilación completada. Instalador disponible en ${outputDirectory}.`);
}

try {
  buildWindows();
} catch (error) {
  console.error('Falló la compilación:', error.message);
  process.exitCode = 1;
}
