const { spawnSync } = require('child_process');
const path = require('path');

const requestedPlatform = process.argv[2];
const requestedArch = process.argv[3] || process.arch;
const hostForTarget = { linux: 'linux', mac: 'darwin' };
const supportedArchitectures = new Set(['x64', 'arm64']);

if (!hostForTarget[requestedPlatform]) {
  console.error('Uso: node scripts/build-platform.js <linux|mac> [x64|arm64]');
  process.exit(1);
}
if (!supportedArchitectures.has(requestedArch)) {
  console.error(`Arquitectura no soportada: ${requestedArch}. Use x64 o arm64.`);
  process.exit(1);
}
if (process.platform !== hostForTarget[requestedPlatform]) {
  console.error(
    `El paquete ${requestedPlatform} debe generarse en ${hostForTarget[requestedPlatform]} ` +
      'porque better-sqlite3 contiene código nativo. Use el workflow de CI o el sistema destino.',
  );
  process.exit(1);
}

const outputDirectory = path.join('release', `${requestedPlatform}-${requestedArch}`);
const builderCli = require.resolve('electron-builder/cli.js');

function run(command, args) {
  const result = spawnSync(command, args, { stdio: 'inherit', shell: false });
  if (result.error) throw result.error;
  if (result.status !== 0) process.exit(result.status ?? 1);
}

console.log(`Compilando Helper Notes para ${requestedPlatform}/${requestedArch}...`);
run(process.execPath, [require.resolve('typescript/bin/tsc'), '-p', 'electron/tsconfig.json']);
run(process.execPath, [require.resolve('@angular/cli/bin/ng.js'), 'build', '--configuration', 'production']);

// electron-builder exige al menos 512 px para el icono de macOS. `sips` está
// incluido en macOS y conserva el recurso generado fuera del código fuente.
if (requestedPlatform === 'mac') {
  run('/usr/bin/sips', [
    '-z',
    '1024',
    '1024',
    'assets/favicon.png',
    '--out',
    'build/icon-mac.png',
  ]);
}

run(process.execPath, [
  builderCli,
  `--${requestedPlatform}`,
  `--${requestedArch}`,
  `--config.directories.output=${outputDirectory}`,
]);
console.log(`Paquetes disponibles en ${outputDirectory}`);
