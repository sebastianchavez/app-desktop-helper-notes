const { spawnSync } = require('child_process');
const path = require('path');

const platform = process.platform === 'darwin' ? 'mac' : process.platform;
const script = platform === 'win32' ? 'build-win.js' : 'build-platform.js';
const args = platform === 'win32' ? [] : [platform, process.arch];
const result = spawnSync(process.execPath, [path.join(__dirname, script), ...args], {
  stdio: 'inherit',
  shell: false,
});

if (result.error) throw result.error;
process.exit(result.status ?? 1);
