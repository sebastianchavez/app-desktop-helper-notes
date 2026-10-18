const { app, BrowserWindow, session, dialog } = require('electron');
const { mkdirSync, writeFileSync, readFileSync } = require('fs');
const path = require('path');
const assert = require('node:assert/strict');
const XLSX = require('xlsx');
const output = path.resolve(__dirname, '../release/chart-qa');
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

/** Recorre los controles reales de gráficos con datos y SQLite aislados; no recibe ni toca documentos del usuario. */
async function verify() {
  require('../electron-dist/db/db').initDb();
  require('../electron-dist/ipc-handlers').registerIpcHandlers();
  const { createNote, updateNote, getNoteById } = require('../electron-dist/db/note.repository');
  const { readSpreadsheetCharts } = require('../electron-dist/spreadsheet-chart');
  const book = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(
    book,
    XLSX.utils.aoa_to_sheet([
      ['Region', 'Product', 'Seller', 'Amount', 'Units'],
      ['North', 'A', 'Ana', 100, 2],
      ['North', 'B', 'Luis', 50, 3],
      ['South', 'A', 'Ana', 200, 4],
      ['North', 'A', 'Ana', 150, 6],
      ['South', 'B', null, null, 1],
    ]),
    'Sales',
  );
  const note = createNote(null, `Chart QA ${Date.now()}`);
  updateNote(note.id, {
    file_type: 'xlsx',
    file_data: Buffer.from(XLSX.write(book, { type: 'buffer', bookType: 'xlsx' })).toString(
      'base64',
    ),
  });
  const other = createNote(null, `Otra chart QA ${Date.now()}`);
  session.defaultSession.webRequest.onBeforeRequest((details, callback) =>
    callback({ cancel: /^https?:/.test(details.url) }),
  );
  const window = new BrowserWindow({
    show: false,
    width: 1700,
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
  window.webContents.on('console-message', (event) => {
    if (/error|exception/i.test(event.message)) console.log(event.message);
  });
  const run = (code) => window.webContents.executeJavaScript(code);
  /** Recibe selector y valor; actualiza un control real con su evento nativo y espera el siguiente render. */
  const change = async (selector, value, event = 'change') => {
    await run(
      `(()=>{const element=document.querySelector(${JSON.stringify(selector)});if(!element)throw Error('Control ausente');element.value=${JSON.stringify(String(value))};element.dispatchEvent(new Event(${JSON.stringify(event)},{bubbles:true}));})()`,
    );
    await new Promise((r) => setTimeout(r, 60));
  };
  /** Pulsa un botón existente por selector; recibe el selector y retorna tras actualizar Angular. */
  const click = async (selector) => {
    await run(`document.querySelector(${JSON.stringify(selector)}).click()`);
    await new Promise((r) => setTimeout(r, 80));
  };
  /** Recupera el último libro guardado; no recibe argumentos y permite verificar sin depender del DOM. */
  const saved = () =>
    XLSX.read(Buffer.from(getNoteById(note.id).file_data, 'base64'), { type: 'buffer' });
  /** Recibe título de nota; la selecciona mediante el listado visible y retorna tras el evento. */
  const selectNote = async (title) => {
    await run(
      `Array.from(document.querySelectorAll('app-note-list h3')).find(n=>n.textContent.includes(${JSON.stringify(title)})).closest('[draggable="true"]').click()`,
    );
    await new Promise((r) => setTimeout(r, 180));
  };

  /** Retorna el gráfico nativo guardado, leyendo SQLite en el proceso principal. */
  const charts = () => readSpreadsheetCharts(getNoteById(note.id).file_data);
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
    await click('.xlsx-chart-toolbar button[title="Insertar gráfico"]');
    await wait(window, `!!document.querySelector('.xlsx-chart-panel')`);
    await change('.chart-title-input', 'Sales chart', 'input');
    await change('.chart-range', 'A1:E6');
    // El origen tiene texto en B/C: elegir Amount y Units como series reales.
    await run(
      `(()=>{const checks=document.querySelectorAll('.xlsx-chart-panel .chart-check input');for(let i=2;i<=3;i++)checks[i].click();})()`,
    );
    await click('.chart-preview-button');
    await wait(window, `!!document.querySelector('.chart-modal-preview svg rect[fill="#2563eb"]')`);
    assert(
      await run(
        `document.querySelector('.chart-modal-preview svg').textContent.includes('Amount')`,
      ),
    );
    writeFileSync(
      path.join(output, 'configuration.png'),
      (await window.webContents.capturePage()).toPNG(),
    );
    await click('.chart-save-button');
    await wait(
      window,
      `!document.querySelector('.xlsx-chart-panel')&&!!document.querySelector('.xlsx-chart-object svg')`,
    );
    assert.equal(charts().length, 1);
    assert.deepEqual(charts()[0].series, [3, 4]);
    const id = charts()[0].id;
    // Cambiar Amount de la primera fila desde los controles reales y comprobar datos y SVG.
    await run(
      `(()=>{const c=document.querySelectorAll('.xlsx-grid tbody tr')[1].querySelectorAll('.xlsx-cell')[3];c.dispatchEvent(new MouseEvent('mousedown',{bubbles:true,button:0}));document.dispatchEvent(new MouseEvent('mouseup',{bubbles:true}));})()`,
    );
    await change('.xlsx-formula-input', '777', 'input');
    await run(
      `document.querySelector('.xlsx-formula-input').dispatchEvent(new KeyboardEvent('keydown',{key:'Enter',bubbles:true}))`,
    );
    await new Promise((resolve) => setTimeout(resolve, 150));
    await wait(window, `document.querySelector('.xlsx-app')?.getAttribute('aria-busy')==='false'`);
    assert.equal(charts()[0].data.series[0].values[0], 777);
    assert(
      await run(`document.querySelector('.xlsx-chart-object svg').textContent.includes('777')`),
    );
    // Mover y redimensionar mediante los mismos eventos que usa el ratón.
    const previous = { ...charts()[0].anchor };
    await run(
      `document.querySelector('.xlsx-chart-handle').dispatchEvent(new MouseEvent('mousedown',{bubbles:true,button:0,clientX:700,clientY:300}));document.dispatchEvent(new MouseEvent('mousemove',{bubbles:true,clientX:800,clientY:380}));document.dispatchEvent(new MouseEvent('mouseup',{bubbles:true}));`,
    );
    await new Promise((resolve) => setTimeout(resolve, 150));
    await wait(window, `document.querySelector('.xlsx-app')?.getAttribute('aria-busy')==='false'`);
    assert.notDeepEqual(charts()[0].anchor, previous);
    await run(
      `document.querySelector('.xlsx-chart-resize').dispatchEvent(new MouseEvent('mousedown',{bubbles:true,button:0,clientX:700,clientY:300}));document.dispatchEvent(new MouseEvent('mousemove',{bubbles:true,clientX:800,clientY:350}));document.dispatchEvent(new MouseEvent('mouseup',{bubbles:true}));`,
    );
    await new Promise((resolve) => setTimeout(resolve, 150));
    await wait(window, `document.querySelector('.xlsx-app')?.getAttribute('aria-busy')==='false'`);
    assert.equal(charts()[0].anchor.width, previous.width + 100);
    assert.equal(charts()[0].anchor.height, previous.height + 50);
    // Cambiar tipo y reabrir no crea un segundo objeto.
    await click('.chart-edit');
    await wait(window, `!!document.querySelector('.xlsx-chart-panel')`);
    await change('.chart-type', 'pie');
    await click('.chart-save-button');
    await wait(window, `!document.querySelector('.xlsx-chart-panel')`);
    assert.equal(charts().length, 1);
    assert.equal(charts()[0].id, id);
    assert.equal(charts()[0].type, 'pie');
    await selectNote(other.title);
    await selectNote(note.title);
    await wait(window, `!!document.querySelector('.xlsx-chart-object svg path')`);
    assert.equal(charts()[0].type, 'pie');
    writeFileSync(
      path.join(output, 'result.png'),
      (await window.webContents.capturePage()).toPNG(),
    );
    writeFileSync(
      path.join(output, 'chart.xlsx'),
      Buffer.from(getNoteById(note.id).file_data, 'base64'),
    );
    const exportedPath = path.join(output, 'exported-chart.xlsx');
    dialog.showSaveDialog = async () => ({ canceled: false, filePath: exportedPath });
    const exported = await run(
      `(async()=>window.electronAPI.fileExportXlsxRaw({defaultName:'Chart.xlsx',base64:(await window.electronAPI.notesGetById(${note.id})).file_data,format:'xlsx'}))()`,
    );
    assert(exported.success);
    assert(
      readFileSync(exportedPath).equals(Buffer.from(getNoteById(note.id).file_data, 'base64')),
    );
    dialog.showMessageBox = async () => ({ response: 1 });
    await click('.chart-delete');
    await wait(window, `!document.querySelector('.xlsx-chart-object')`);
    assert.equal(charts().length, 0);
    await run(
      `Array.from(document.querySelectorAll('.xlsx-toolbar button')).find(b=>b.textContent.includes('undo')).click()`,
    );
    await wait(window, `!!document.querySelector('.xlsx-chart-object svg')`);
    assert.equal(charts().length, 1);
    await run(
      `Array.from(document.querySelectorAll('.xlsx-toolbar button')).find(b=>b.textContent.includes('redo')).click()`,
    );
    await wait(window, `!document.querySelector('.xlsx-chart-object')`);
    assert.equal(charts().length, 0);
    console.log(
      JSON.stringify({
        preview: true,
        native: true,
        linkedRefresh: true,
        move: true,
        resize: true,
        edit: true,
        reopen: true,
        export: true,
        delete: true,
        undoRedo: true,
      }),
    );
  } catch (error) {
    console.error(
      await run(
        `document.querySelector('.xlsx-pivot-error')?.innerText??document.querySelector('.xlsx-error')?.innerText??document.body.innerText.slice(-700)`,
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
