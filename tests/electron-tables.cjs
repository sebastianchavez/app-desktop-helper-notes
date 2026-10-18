const { app, BrowserWindow, session } = require('electron');
const { mkdirSync, readFileSync, writeFileSync } = require('fs');
const path = require('path');
const assert = require('node:assert/strict');
const XLSX = require('xlsx');
const output = path.resolve(__dirname, '../release/tables-qa');
mkdirSync(output, { recursive: true });
app.setPath('userData', path.join(output, 'profile'));
app.disableHardwareAcceleration();

/** Recibe ventana y condición de interfaz; espera su cumplimiento o rechaza sin tocar documentos del usuario. */
async function wait(win, expression) {
  for (let i = 0; i < 200; i++) {
    if (await win.webContents.executeJavaScript(expression)) return;
    await new Promise((r) => setTimeout(r, 100));
  }
  throw Error('No se cumplió: ' + expression);
}

/** Crea libros sintéticos y verifica por interfaz la vida completa de tablas, fórmulas, totales y reapertura. */
async function verify() {
  require('../electron-dist/db/db').initDb();
  require('../electron-dist/ipc-handlers').registerIpcHandlers();
  const { createNote, updateNote, getNoteById } = require('../electron-dist/db/note.repository');
  const { readSpreadsheetStructure } = require('../electron-dist/spreadsheet-structure');
  const imported = createNote(null, 'Imported table QA ' + Date.now());
  updateNote(imported.id, {
    file_type: 'xlsx',
    file_data: readFileSync(path.join(__dirname, 'fixtures/dynamic-source.xlsx')).toString(
      'base64',
    ),
  });
  const book = XLSX.utils.book_new();
  const plain = XLSX.utils.aoa_to_sheet([
    ['Product', 'Qty', 'Price', 'Total', null, 0],
    ['A', 2, 5, 10, null, 101],
    ['B', 1, 9, 9, null, 202],
    ['C', 3, 2, 6, null, 303],
  ]);
  for (let row = 2; row <= 4; row++)
    plain['D' + row] = { t: 'n', v: plain['D' + row].v, f: 'B' + row + '*C' + row };
  plain.F1 = { t: 'n', v: 0, f: 'SUM(Orders[Total])' };
  XLSX.utils.book_append_sheet(book, plain, 'Data');
  const created = createNote(null, 'New table QA ' + Date.now());
  updateNote(created.id, {
    file_type: 'xlsx',
    file_data: Buffer.from(XLSX.write(book, { type: 'buffer', bookType: 'xlsx' })).toString(
      'base64',
    ),
  });
  session.defaultSession.webRequest.onBeforeRequest((details, callback) =>
    callback({ cancel: /^https?:/.test(details.url) }),
  );
  const win = new BrowserWindow({
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
  const run = (code) => win.webContents.executeJavaScript(code);
  /** Espera cálculo, renderizado y escritura SQLite de la acción anterior. */
  const saved = async () => {
    await new Promise((r) => setTimeout(r, 220));
    await wait(win, `document.querySelector('.xlsx-app')?.getAttribute('aria-busy')==='false'`);
  };
  /** Recibe nota; la abre desde su fila real y espera la cuadrícula. */
  const note = async (n) => {
    await run(
      `Array.from(document.querySelectorAll('app-note-list h3')).find(n=>n.textContent.includes(${JSON.stringify(n.title)})).closest('[draggable="true"]').click()`,
    );
    await new Promise((r) => setTimeout(r, 250));
    await wait(win, `!!document.querySelector('.xlsx-grid .xlsx-cell')`);
  };
  /** Recibe coordenadas y selección final opcional; selecciona un rango mediante eventos reales. */
  const select = async (row, col, endRow = row, endCol = col) =>
    run(
      `(()=>{const at=(r,c)=>document.querySelectorAll('.xlsx-grid tbody tr')[r].querySelectorAll('.xlsx-cell')[c];at(${row},${col}).dispatchEvent(new MouseEvent('mousedown',{bubbles:true,button:0}));at(${endRow},${endCol}).dispatchEvent(new MouseEvent('mouseenter',{bubbles:true}));document.dispatchEvent(new MouseEvent('mouseup',{bubbles:true}));})()`,
    );
  /** Recibe celda y contenido; edita con la barra de fórmulas de la aplicación. */
  const edit = async (row, col, value) => {
    await select(row, col);
    await run(
      `(()=>{const input=document.querySelector('.xlsx-formula-input');input.value=${JSON.stringify(value)};input.dispatchEvent(new Event('input',{bubbles:true}));input.dispatchEvent(new KeyboardEvent('keydown',{key:'Enter',bubbles:true}));})()`,
    );
    await saved();
  };
  /** Recibe selector, valor y evento; cambia un control nativo del formulario. */
  const change = async (selector, value, event = 'input') =>
    run(
      `(()=>{const input=document.querySelector(${JSON.stringify(selector)});input.value=${JSON.stringify(value)};input.dispatchEvent(new Event(${JSON.stringify(event)},{bubbles:true}));})()`,
    );
  /** Recibe nombre visible; abre el editor o creador de tablas por su botón. */
  const open = async (label) => {
    await run(
      `Array.from(document.querySelectorAll('.xlsx-chart-toolbar button')).find(b=>b.textContent.includes(${JSON.stringify(label)})).click()`,
    );
    await wait(win, `!!document.querySelector('.xlsx-table-panel')`);
  };
  /** Guarda el formulario y verifica que no haya validaciones pendientes. */
  const save = async () => {
    await run(`document.querySelector('.xlsx-table-panel .pivot-save-button').click()`);
    await saved();
    assert(
      !(await run(`!!document.querySelector('.xlsx-table-panel')`)),
      await run(`document.querySelector('.xlsx-table-panel [role="alert"]')?.textContent`),
    );
  };
  /** Recibe título de herramienta; pulsa y espera a que el libro quede guardado. */
  const tool = async (title) => {
    await run(
      `Array.from(document.querySelectorAll('.xlsx-tb-btn')).find(b=>b.title===${JSON.stringify(title)}).click()`,
    );
    await saved();
  };
  /** Recibe celda y acción; ejecuta una operación del menú contextual. */
  const context = async (row, col, label) => {
    await select(row, col);
    await run(
      `document.querySelectorAll('.xlsx-grid tbody tr')[${row}].querySelectorAll('.xlsx-cell')[${col}].dispatchEvent(new MouseEvent('contextmenu',{bubbles:true,clientX:700,clientY:300}))`,
    );
    await run(
      `Array.from(document.querySelectorAll('.xlsx-ctx-item')).find(b=>b.textContent.includes(${JSON.stringify(label)})).click()`,
    );
    await saved();
  };
  const data = (n) =>
    XLSX.read(Buffer.from(getNoteById(n.id).file_data, 'base64'), { type: 'buffer', xlfn: true });
  const table = (n) => readSpreadsheetStructure(getNoteById(n.id).file_data).Data.tables[0];
  try {
    await win.loadFile(
      path.resolve(__dirname, '../dist/app-desktop-helper-notes/browser/index.html'),
    );
    await wait(
      win,
      `Array.from(document.querySelectorAll('app-note-list h3')).some(n=>n.textContent.includes(${JSON.stringify(imported.title)}))`,
    );
    await note(imported);
    await select(1, 1);
    await open('Redimensionar tabla');
    await change('.xlsx-table-name', 'Ventas');
    await run(`document.querySelector('.xlsx-table-totals-toggle input').click()`);
    await save();
    assert.equal(table(imported).name, 'Ventas');
    assert.equal(table(imported).id, 1);
    assert.equal(table(imported).ref, 'A1:D5');
    assert.equal(data(imported).Sheets.Data.D5.v, 1000);
    assert.equal(data(imported).Sheets.Dashboard.L1.f, 'SUM(Ventas[Amount])');
    await select(1, 1);
    await open('Redimensionar tabla');
    await change('.xlsx-table-range', 'A1:D7');
    await save();
    assert.equal(data(imported).Sheets.Data.D7.v, 1000);
    assert.equal(data(imported).Sheets.Data.D5.v, 0);
    await select(1, 1);
    await tool('Insertar fila');
    assert.equal(table(imported).ref, 'A1:D8');
    await edit(2, 1, '10');
    await edit(2, 2, '2');
    assert.equal(data(imported).Sheets.Data.D3.v, 20);
    assert.equal(data(imported).Sheets.Data.D8.v, 1020);
    await select(1, 0);
    await tool('Insertar columna');
    assert.deepEqual(table(imported).columnIds, [1, 5, 2, 3, 4]);
    assert.equal(table(imported).ref, 'A1:E8');
    assert.equal(data(imported).Sheets.Data.E8.v, 1020);
    await context(1, 1, 'Eliminar columna(s)');
    assert.deepEqual(table(imported).columnIds, [1, 2, 3, 4]);
    assert.equal(data(imported).Sheets.Data.D8.v, 1020);
    await edit(0, 1, 'Precio');
    assert.equal(table(imported).columns[1], 'Precio');
    assert.deepEqual(table(imported).columnIds, [1, 2, 3, 4]);
    assert.equal(data(imported).Sheets.Data.D3.v, 20);
    assert.equal(data(imported).Sheets.Dashboard.L1.f, 'SUM(Ventas[Precio])');
    await context(0, 0, 'Eliminar fila(s)');
    assert.equal(table(imported).ref, 'A1:D8');
    assert(await run(`document.querySelector('.xlsx-error')?.textContent.includes('encabezados')`));
    writeFileSync(
      path.join(output, 'imported.xlsx'),
      Buffer.from(getNoteById(imported.id).file_data, 'base64'),
    );
    await note(created);
    await select(0, 0, 3, 3);
    await open('Crear tabla');
    await change('.xlsx-table-name', 'Orders');
    await run(`document.querySelector('.xlsx-table-totals-toggle input').click()`);
    await save();
    assert.equal(table(created).ref, 'A1:D5');
    assert.equal(data(created).Sheets.Data.D5.v, 25);
    assert.equal(data(created).Sheets.Data.F1.v, 25);
    await edit(4, 4, 'Conservar');
    await select(1, 0);
    await open('Redimensionar tabla');
    await change('.xlsx-table-range', 'A1:E5');
    await run(`document.querySelector('.xlsx-table-panel .pivot-save-button').click()`);
    await new Promise((r) => setTimeout(r, 120));
    assert(
      await run(
        `document.querySelector('.xlsx-table-panel [role="alert"]')?.textContent.includes('contiene datos')`,
      ),
    );
    assert.equal(table(created).ref, 'A1:D5');
    assert.equal(data(created).Sheets.Data.E5.v, 'Conservar');
    await run(`document.querySelector('.xlsx-table-panel footer button').click()`);
    await select(1, 1);
    await tool('Ordenar filas de mayor a menor');
    assert.equal(data(created).Sheets.Data.B2.v, 3);
    assert.equal(data(created).Sheets.Data.F2.v, 101);
    assert.equal(data(created).Sheets.Data.F3.v, 202);
    assert.equal(data(created).Sheets.Data.D5.v, 25);
    await edit(0, 3, 'Importe');
    assert.deepEqual(table(created).columnIds, [1, 2, 3, 4]);
    assert.equal(data(created).Sheets.Data.F1.f, 'SUM(Orders[Importe])');
    assert.equal(data(created).Sheets.Data.F1.v, 25);
    await select(1, 0);
    await open('Redimensionar tabla');
    await change('.xlsx-table-name', 'Archive');
    await save();
    assert.equal(data(created).Sheets.Data.F1.f, 'SUM(Archive[Importe])');
    await edit(4, 3, '=SUM(D2:D4)*2');
    assert.equal(data(created).Sheets.Data.D5.v, 50);
    assert.equal(table(created).totals[3], 'custom');
    await select(1, 0);
    await open('Redimensionar tabla');
    await run(
      `(()=>{const label=Array.from(document.querySelectorAll('.xlsx-table-aggregates label')).find(n=>n.textContent.includes('Importe'));const input=label.querySelector('select');input.value='average';input.dispatchEvent(new Event('change',{bubbles:true}));})()`,
    );
    await save();
    assert(Math.abs(data(created).Sheets.Data.D5.v - 25 / 3) < 1e-10);
    await select(1, 0);
    await open('Redimensionar tabla');
    await run(`document.querySelector('.xlsx-table-totals-toggle input').click()`);
    await save();
    assert.equal(table(created).totalRows, 0);
    assert(!data(created).Sheets.Data.D5 || data(created).Sheets.Data.D5.v === undefined);
    await run(
      `Array.from(document.querySelectorAll('.xlsx-tb-btn')).find(b=>b.querySelector('span')?.textContent.trim()==='undo').click()`,
    );
    await saved();
    assert.equal(table(created).totalRows, 1);
    assert(Math.abs(data(created).Sheets.Data.D5.v - 25 / 3) < 1e-10);
    await note(imported);
    await note(created);
    assert.equal(table(created).name, 'Archive');
    assert.equal(table(created).totals[3], 'average');
    writeFileSync(
      path.join(output, 'created.xlsx'),
      Buffer.from(getNoteById(created.id).file_data, 'base64'),
    );
    writeFileSync(path.join(output, 'result.png'), (await win.webContents.capturePage()).toPNG());
    await select(1, 0);
    await open('Redimensionar tabla');
    await new Promise((r) => setTimeout(r, 150));
    assert.deepEqual(
      await run(
        `Array.from(document.querySelectorAll('.xlsx-table-aggregates select')).map(s=>s.value)`,
      ),
      ['', 'sum', 'sum', 'average'],
    );
    writeFileSync(
      path.join(output, 'table-editor.png'),
      (await win.webContents.capturePage()).toPNG(),
    );
    console.log(
      JSON.stringify({
        create: true,
        rename: true,
        headers: true,
        stableIds: true,
        totals: true,
        relocateTotals: true,
        calculatedColumns: true,
        insertRows: true,
        insertColumns: true,
        deleteColumns: true,
        sortPreservesOutsideCells: true,
        protectedHeaders: true,
        undoRedo: true,
        reopened: true,
      }),
    );
  } catch (error) {
    console.error(
      await run(
        `document.querySelector('.xlsx-error')?.textContent??document.body.textContent.slice(-800)`,
      ),
    );
    writeFileSync(path.join(output, 'failure.png'), (await win.webContents.capturePage()).toPNG());
    throw error;
  } finally {
    win.destroy();
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
