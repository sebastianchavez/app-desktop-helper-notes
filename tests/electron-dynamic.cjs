const { app, BrowserWindow, session, dialog } = require('electron');
const { mkdirSync, writeFileSync, readFileSync } = require('fs');
const path = require('path');
const assert = require('node:assert/strict');
const XLSX = require('xlsx');
const output = path.resolve(__dirname, '../release/dynamic-qa');
mkdirSync(output, { recursive: true });
app.setPath('userData', path.join(output, 'profile'));
app.disableHardwareAcceleration();

/** Recibe ventana y expresión; espera a que Angular/Electron complete una condición o rechaza por tiempo agotado. */
async function wait(window, expression) {
  for (let i = 0; i < 200; i++) {
    if (await window.webContents.executeJavaScript(expression)) return;
    await new Promise((r) => setTimeout(r, 100));
  }
  throw Error('No se cumplió: ' + expression);
}

/** Verifica matrices nativas externas y edición real con datos sintéticos y SQLite aislado. */
async function verify() {
  require('../electron-dist/db/db').initDb();
  require('../electron-dist/ipc-handlers').registerIpcHandlers();
  const { createNote, updateNote, getNoteById } = require('../electron-dist/db/note.repository');
  const { readSpreadsheetStructure } = require('../electron-dist/spreadsheet-structure');
  const { strFromU8, unzipSync } = require('fflate');
  const original = readFileSync(path.resolve(__dirname, 'fixtures/dynamic-source.xlsx'));
  const note = createNote(null, 'Dynamic QA ' + Date.now());
  updateNote(note.id, { file_type: 'xlsx', file_data: original.toString('base64') });
  const other = createNote(null, 'Other dynamic QA ' + Date.now());
  session.defaultSession.webRequest.onBeforeRequest((details, callback) =>
    callback({ cancel: /^https?:/.test(details.url) }),
  );
  const window = new BrowserWindow({
    show: false,
    width: 1900,
    height: 1100,
    webPreferences: {
      preload: path.resolve(__dirname, '../electron-dist/preload.js'),
      contextIsolation: true,
      nodeIntegration: false,
      sandbox: false,
      offscreen: true,
      backgroundThrottling: false,
    },
  });
  const run = (code) => window.webContents.executeJavaScript(code);
  /** Recibe título; selecciona una nota mediante su fila real del listado. */
  const selectNote = async (title) => {
    await run(
      `Array.from(document.querySelectorAll('app-note-list h3')).find(n=>n.textContent.includes(${JSON.stringify(title)})).closest('[draggable="true"]').click()`,
    );
    await new Promise((r) => setTimeout(r, 150));
  };
  /** Recibe nombre; activa una hoja mediante su pestaña. */
  const sheet = async (name) => {
    await run(
      `Array.from(document.querySelectorAll('.xlsx-sheet-tab')).find(b=>b.textContent.trim()===${JSON.stringify(name)}).click()`,
    );
    await new Promise((r) => setTimeout(r, 70));
  };
  /** Recibe fila, columna y expresión; edita mediante barra real y espera la escritura. */
  const edit = async (row, col, value) => {
    await run(
      `(()=>{const cell=document.querySelectorAll('.xlsx-grid tbody tr')[${row}].querySelectorAll('.xlsx-cell')[${col}];cell.dispatchEvent(new MouseEvent('mousedown',{bubbles:true,button:0}));document.dispatchEvent(new MouseEvent('mouseup',{bubbles:true}));const input=document.querySelector('.xlsx-formula-input');input.value=${JSON.stringify(value)};input.dispatchEvent(new Event('input',{bubbles:true}));input.dispatchEvent(new KeyboardEvent('keydown',{key:'Enter',bubbles:true}));})()`,
    );
    await new Promise((r) => setTimeout(r, 200));
    await wait(window, `document.querySelector('.xlsx-app')?.getAttribute('aria-busy')==='false'`);
  };
  const saved = () =>
    XLSX.read(Buffer.from(getNoteById(note.id).file_data, 'base64'), {
      type: 'buffer',
      xlfn: true,
    });
  try {
    await window.loadFile(
      path.resolve(__dirname, '../dist/app-desktop-helper-notes/browser/index.html'),
    );
    await wait(
      window,
      `Array.from(document.querySelectorAll('app-note-list h3')).some(n=>n.textContent.includes(${JSON.stringify(note.title)}))`,
    );
    await selectNote(note.title);
    await wait(window, `!!document.querySelector('.xlsx-grid .xlsx-cell')`);
    await edit(1, 1, '101');
    assert.equal(saved().Sheets.Data.D2.v, 202);
    await sheet('Dashboard');
    let data = saved().Sheets.Dashboard;
    assert.equal(data.J1.v, 351);
    assert.equal(data.K1.v, 1048572);
    assert.equal(data.L1.v, 351);
    assert.equal(data.M1.v, 135);
    assert.equal(data.N1.v, 21);
    assert.equal(data.A1.v, 'North');
    assert.equal(data.A2.v, 'South');
    assert.equal(data.E2.v, 'South');
    await edit(1, 6, '999');
    assert.equal(saved().Sheets.Dashboard.G2.v, 20);
    assert(await run(`document.querySelector('.xlsx-error')?.textContent.includes('matriz')`));
    await edit(1, 15, '=SECUENCIA(3;2)');
    let structure = readSpreadsheetStructure(getNoteById(note.id).file_data);
    assert.equal(structure.Dashboard.errors.P2, '#SPILL!');
    assert.equal(saved().Sheets.Dashboard.Q3.v, 77);
    await edit(2, 16, '');
    data = saved().Sheets.Dashboard;
    assert.equal(data.P2.v, 1);
    assert.equal(data.Q3.v, 4);
    assert.equal(data.Q4.v, 6);
    await edit(0, 6, '=SECUENCIA(2;2;10;5)');
    data = saved().Sheets.Dashboard;
    assert.equal(data.M1.v, 70);
    assert(!data.G3 || data.G3.v === undefined);
    assert(!data.H3 || data.H3.v === undefined);
    structure = readSpreadsheetStructure(getNoteById(note.id).file_data);
    assert.equal(structure.Dashboard.arrays.G1.ref, 'G1:H2');
    assert.equal(structure.Dashboard.arrays.G1.dynamic, true);
    assert.equal(structure.Data.tables[0].name, 'Sales');
    const archive = unzipSync(Buffer.from(getNoteById(note.id).file_data, 'base64')),
      before = unzipSync(original);
    assert.equal(
      strFromU8(archive['xl/tables/table1.xml']),
      strFromU8(before['xl/tables/table1.xml']),
    );
    /** Recibe coordenadas; selecciona una celda sin cambiar su contenido. */
    const select = async (row, col) =>
      run(
        `(()=>{const cell=document.querySelectorAll('.xlsx-grid tbody tr')[${row}].querySelectorAll('.xlsx-cell')[${col}];cell.dispatchEvent(new MouseEvent('mousedown',{bubbles:true,button:0}));document.dispatchEvent(new MouseEvent('mouseup',{bubbles:true}));})()`,
      );
    /** Recibe título de herramienta; pulsa el botón y espera al guardado. */
    const tool = async (title) => {
      await run(
        `Array.from(document.querySelectorAll('.xlsx-tb-btn')).find(b=>b.title===${JSON.stringify(title)}).click()`,
      );
      await new Promise((r) => setTimeout(r, 200));
      await wait(
        window,
        `document.querySelector('.xlsx-app')?.getAttribute('aria-busy')==='false'`,
      );
    };
    await select(0, 0);
    await tool('Insertar fila');
    data = saved().Sheets.Dashboard;
    assert.equal(data.G2.v, 20); // Insertar dentro del desbordamiento regenera los resultados, no los desplaza como literales.
    assert(!data.G3 || data.G3.v === undefined);
    assert.equal(data.P3.v, 1);
    assert.equal(
      readSpreadsheetStructure(getNoteById(note.id).file_data).Dashboard.arrays.P3.ref,
      'P3:Q5',
    );
    await select(0, 0);
    await tool('Insertar columna');
    data = saved().Sheets.Dashboard;
    assert.equal(data.H1.v, 10);
    assert.equal(data.I2.v, 25);
    assert.equal(data.N1.v, 70);
    assert.equal(data.N1.f, 'SUM(H1#)'.replace('H1#', '_xlfn.ANCHORARRAY(H1)'));
    // Volver al ancla inicial también comprueba que deshacer/restaurar importa las matrices nativas.
    await run(
      `Array.from(document.querySelectorAll('.xlsx-tb-btn')).find(b=>b.querySelector('span')?.textContent.trim()==='undo').click()`,
    );
    await new Promise((r) => setTimeout(r, 250));
    await wait(window, `document.querySelector('.xlsx-app')?.getAttribute('aria-busy')==='false'`);
    assert.equal(saved().Sheets.Dashboard.G1.v, 10);
    await select(1, 0);
    await run(
      `document.querySelectorAll('.xlsx-grid tbody tr')[1].querySelector('.xlsx-cell').dispatchEvent(new MouseEvent('contextmenu',{bubbles:true,clientX:700,clientY:300}))`,
    );
    await run(
      `Array.from(document.querySelectorAll('.xlsx-ctx-item')).find(b=>b.textContent.includes('Eliminar fila(s)')).click()`,
    );
    await new Promise((r) => setTimeout(r, 250));
    await wait(window, `document.querySelector('.xlsx-app')?.getAttribute('aria-busy')==='false'`);
    data = saved().Sheets.Dashboard;
    assert.equal(data.G2.v, 20);
    assert.equal(data.P2.v, 1);
    assert.equal(
      readSpreadsheetStructure(getNoteById(note.id).file_data).Dashboard.arrays.P2.ref,
      'P2:Q4',
    );
    await select(0, 0);
    await tool('Ordenar filas de menor a mayor');
    assert(
      await run(
        `document.querySelector('.xlsx-error')?.textContent.includes('atraviesa una matriz')`,
      ),
    );
    assert.equal(saved().Sheets.Dashboard.G1.v, 10);
    await run(
      `document.querySelectorAll('.xlsx-grid tbody tr')[0].querySelector('.xlsx-row-header').click()`,
    );
    await tool('Mover fila(s) hacia abajo');
    data = saved().Sheets.Dashboard;
    assert.equal(data.G2.v, 10);
    assert.equal(data.G3.v, 20);
    assert.equal(data.P1.v, 1);
    assert.equal(data.M2.v, 70);
    await run(
      `Array.from(document.querySelectorAll('.xlsx-tb-btn')).find(b=>b.querySelector('span')?.textContent.trim()==='undo').click()`,
    );
    await new Promise((r) => setTimeout(r, 250));
    await wait(window, `document.querySelector('.xlsx-app')?.getAttribute('aria-busy')==='false'`);
    assert.equal(saved().Sheets.Dashboard.G1.v, 10);
    await select(0, 6);
    await run(
      `document.querySelectorAll('.xlsx-grid tbody tr')[0].querySelectorAll('.xlsx-cell')[6].dispatchEvent(new MouseEvent('contextmenu',{bubbles:true,clientX:700,clientY:300}))`,
    );
    await run(
      `Array.from(document.querySelectorAll('.xlsx-ctx-item')).find(b=>b.textContent.includes('Eliminar columna(s)')).click()`,
    );
    await new Promise((r) => setTimeout(r, 250));
    await wait(window, `document.querySelector('.xlsx-app')?.getAttribute('aria-busy')==='false'`);
    assert(!readSpreadsheetStructure(getNoteById(note.id).file_data).Dashboard.arrays.G1);
    assert.equal(saved().Sheets.Dashboard.O3.v, 3);
    await run(
      `Array.from(document.querySelectorAll('.xlsx-tb-btn')).find(b=>b.querySelector('span')?.textContent.trim()==='undo').click()`,
    );
    await new Promise((r) => setTimeout(r, 250));
    await wait(window, `document.querySelector('.xlsx-app')?.getAttribute('aria-busy')==='false'`);
    assert.equal(saved().Sheets.Dashboard.G1.v, 10);
    await sheet('Data');
    await edit(4, 0, 'East');
    await edit(4, 1, '30');
    await edit(4, 2, '5');
    await select(1, 1);
    await run(
      `Array.from(document.querySelectorAll('.xlsx-tb-btn')).find(b=>b.textContent.includes('Redimensionar tabla')).click()`,
    );
    await run(
      `(()=>{const input=document.querySelector('.xlsx-table-range');input.value='A1:E5';input.dispatchEvent(new Event('input',{bubbles:true}));input.dispatchEvent(new KeyboardEvent('keydown',{key:'Enter',bubbles:true}));})()`,
    );
    await new Promise((r) => setTimeout(r, 300));
    await wait(window, `document.querySelector('.xlsx-app')?.getAttribute('aria-busy')==='false'`);
    assert.equal(
      readSpreadsheetStructure(getNoteById(note.id).file_data).Data.tables[0].ref,
      'A1:E5',
    );
    assert.equal(saved().Sheets.Data.D5.v, 150);
    assert.equal(saved().Sheets.Data.E1.v, 'Column5');
    assert.equal(saved().Sheets.Dashboard.L1.v, 381);
    await selectNote(other.title);
    await selectNote(note.title);
    await wait(window, `!!document.querySelector('.xlsx-grid .xlsx-cell')`);
    await sheet('Dashboard');
    assert(await run(`document.querySelectorAll('.xlsx-spill-cell').length>=10`));
    writeFileSync(
      path.join(output, 'result.xlsx'),
      Buffer.from(getNoteById(note.id).file_data, 'base64'),
    );
    writeFileSync(
      path.join(output, 'result.png'),
      (await window.webContents.capturePage()).toPNG(),
    );
    console.log(
      JSON.stringify({
        externalDynamicArrays: true,
        filter: true,
        sortUnique: true,
        sequence: true,
        spillDependencies: true,
        structuredTables: true,
        fullColumns: true,
        readonlyChildren: true,
        blockedSpillPreservesData: true,
        unblockedSpill: true,
        shrinkClearsChildren: true,
        nativeMetadata: true,
        tableXmlPreserved: true,
        structuralArrayEdits: true,
        nativeTableResize: true,
        reopened: true,
      }),
    );
  } catch (error) {
    console.error(
      await run(
        `document.querySelector('.xlsx-error')?.textContent??document.body.textContent.slice(-600)`,
      ),
    );
    writeFileSync(
      path.join(output, 'failure.png'),
      (await window.webContents.capturePage()).toPNG(),
    );
    throw error;
  } finally {
    window.destroy();
  }
}
app
  .whenReady()
  .then(verify)
  .then(() => app.exit(0))
  .catch((error) => {
    console.error(error);
    app.exit(1);
  });
