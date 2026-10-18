const { execFile } = require('child_process');
const path = require('path');
const { promisify } = require('util');
const { getRceditBundle } = require('app-builder-lib/out/toolsets/windows');

const executeFile = promisify(execFile);

/**
 * Actualiza los recursos del ejecutable con el paquete RCEdit exclusivo de Windows.
 * Evita el paquete heredado winCodeSign-2.6.0, que contiene enlaces simbólicos
 * de macOS cuya extracción requiere privilegios adicionales en Windows.
 * @param {import('app-builder-lib').AfterPackContext} context Resultado del empaquetado de Electron.
 * @returns {Promise<void>} Resuelve al guardar icono, nombre y versiones; no modifica otros sistemas.
 * @throws Si la descarga verificada o la edición del ejecutable falla.
 */
module.exports = async function editWindowsExecutable(context) {
  if (context.electronPlatformName !== 'win32') return;

  const { appInfo } = context.packager;
  const tools = await getRceditBundle(context.packager.config.toolsets.winCodeSign);
  const executablePath = path.join(context.appOutDir, `${appInfo.productFilename}.exe`);
  const iconPath = path.resolve(context.packager.projectDir, 'assets/favicon.ico');
  const editorPath = process.arch === 'ia32' ? tools.x86 : tools.x64;

  await executeFile(editorPath, [
    executablePath,
    '--set-icon',
    iconPath,
    '--set-version-string',
    'FileDescription',
    appInfo.productName,
    '--set-version-string',
    'ProductName',
    appInfo.productName,
    '--set-version-string',
    'LegalCopyright',
    appInfo.copyright,
    '--set-version-string',
    'InternalName',
    appInfo.productFilename,
    '--set-version-string',
    'OriginalFilename',
    `${appInfo.productFilename}.exe`,
    '--set-file-version',
    appInfo.shortVersion || appInfo.buildVersion,
    '--set-product-version',
    appInfo.shortVersionWindows || appInfo.getVersionInWeirdWindowsForm(),
  ]);

  // Conserva la firma opcional cuando el equipo de compilación dispone de un certificado.
  await context.packager.signIf(executablePath);
};
