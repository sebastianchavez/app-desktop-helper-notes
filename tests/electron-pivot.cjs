const { app, BrowserWindow, session, dialog } = require('electron');
const { mkdirSync, writeFileSync, readFileSync } = require('fs');
const path = require('path');
const assert = require('node:assert/strict');
const XLSX = require('xlsx');
const output = path.resolve(__dirname, '../release/pivot-qa');
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

/** Recorre los controles reales de tablas dinámicas con datos y SQLite aislados; no recibe ni toca documentos del usuario. */
async function verify() {
  require('../electron-dist/db/db').initDb();
  require('../electron-dist/ipc-handlers').registerIpcHandlers();
  const { createNote, updateNote, getNoteById } = require('../electron-dist/db/note.repository');
  const { readSpreadsheetPivots } = require('../electron-dist/spreadsheet-pivot');
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
  const note = createNote(null, `Pivot QA ${Date.now()}`);
  updateNote(note.id, {
    file_type: 'xlsx',
    file_data: Buffer.from(XLSX.write(book, { type: 'buffer', bookType: 'xlsx' })).toString(
      'base64',
    ),
  });
  const other = createNote(null, `Otra pivot QA ${Date.now()}`);
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
  /** Retorna la fila de totales de Summary desde SQLite, sin usar datos de vista previa. */
  const total = () =>
    XLSX.utils
      .sheet_to_json(saved().Sheets.Summary, { header: 1 })
      .find((row) => row[0] === 'Total general');
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
    await run(
      `Array.from(document.querySelectorAll('.xlsx-toolbar button')).find(b=>b.textContent.includes('pivot_table_chart')).click()`,
    );
    await wait(window, `!!document.querySelector('.xlsx-pivot-panel')`);
    await change('.pivot-name', 'Summary', 'input');
    await change('.pivot-value-field', 3);
    await click('.pivot-column-field[data-field="1"]');
    await click('.pivot-add-value');
    await run(
      `(()=>{const fields=document.querySelectorAll('.pivot-value-field'),aggregates=document.querySelectorAll('.pivot-aggregate');fields[1].value='3';fields[1].dispatchEvent(new Event('change',{bubbles:true}));aggregates[1].value='average';aggregates[1].dispatchEvent(new Event('change',{bubbles:true}));})()`,
    );
    await click('.pivot-preview-button');
    await wait(window, `!!document.querySelector('.xlsx-pivot-preview table')`);
    assert(await run(`document.querySelector('.xlsx-pivot-preview').textContent.includes('500')`));
    await click('.pivot-add-filter');
    await click('.pivot-preview-button');
    await wait(window, `document.querySelector('.xlsx-pivot-preview').textContent.includes('450')`);
    writeFileSync(
      path.join(output, 'configuration.png'),
      (await window.webContents.capturePage()).toPNG(),
    );
    await click('.pivot-save-button');
    await wait(
      window,
      `!document.querySelector('.xlsx-pivot-panel') && !!document.querySelector('.xlsx-pivot-notice')`,
    );
    assert.equal(total().at(-2), 450);
    assert.equal(total().at(-1), 150);
    assert.equal(readSpreadsheetPivots(getNoteById(note.id).file_data).length, 1);
    // La tabla se actualiza automáticamente al editar el origen mediante la barra de fórmulas.
    await run(
      `Array.from(document.querySelectorAll('.xlsx-sheet-tab')).find(b=>b.textContent.trim()==='Sales').click()`,
    );
    await run(
      `(()=>{const c=document.querySelectorAll('.xlsx-grid tbody tr')[1].querySelectorAll('.xlsx-cell')[3];c.dispatchEvent(new MouseEvent('mousedown',{bubbles:true,button:0}));document.dispatchEvent(new MouseEvent('mouseup',{bubbles:true}));})()`,
    );
    await change('.xlsx-formula-input', '200', 'input');
    await run(
      `document.querySelector('.xlsx-formula-input').dispatchEvent(new KeyboardEvent('keydown',{key:'Enter',bubbles:true}))`,
    );
    await wait(window, `document.querySelector('.xlsx-app')?.getAttribute('aria-busy')==='false'`);
    assert.equal(total().at(-2), 550);
    await run(
      `Array.from(document.querySelectorAll('.xlsx-sheet-tab')).find(b=>b.textContent.trim()==='Summary').click()`,
    );
    await wait(window, `document.querySelector('.xlsx-grid').innerText.includes('550')`);
    // Cerrar la nota y volver obliga a reconstruir el estado desde SQLite, incluidas las áreas de campos.
    await selectNote(other.title);
    await selectNote(note.title);
    await wait(window, `!!document.querySelector('.xlsx-grid .xlsx-cell')`);
    await run(
      `Array.from(document.querySelectorAll('.xlsx-sheet-tab')).find(b=>b.textContent.trim()==='Summary').click()`,
    );
    await run(`document.querySelector('.xlsx-pivot-notice button').click()`);
    await wait(window, `!!document.querySelector('.xlsx-pivot-panel')`);
    assert.equal(await run(`document.querySelectorAll('.pivot-aggregate')[1].value`), 'average');
    await change('.pivot-aggregate', 'count');
    await click('.pivot-save-button');
    await wait(window, `!document.querySelector('.xlsx-pivot-panel')`);
    assert.equal(total().at(-2), 3);
    assert.equal(saved().SheetNames.length, 2);
    writeFileSync(
      path.join(output, 'pivot-table.xlsx'),
      Buffer.from(getNoteById(note.id).file_data, 'base64'),
    );
    const exportedPath = path.join(output, 'exported-pivot.xlsx');
    dialog.showSaveDialog = async () => ({ canceled: false, filePath: exportedPath });
    const exported = await run(
      `(async()=>window.electronAPI.fileExportXlsxRaw({defaultName:'Pivot.xlsx',base64:(await window.electronAPI.notesGetById(${note.id})).file_data,format:'xlsx'}))()`,
    );
    assert(exported.success);
    assert(
      readFileSync(exportedPath).equals(Buffer.from(getNoteById(note.id).file_data, 'base64')),
    );
    writeFileSync(
      path.join(output, 'result.png'),
      (await window.webContents.capturePage()).toPNG(),
    );
    console.log(
      JSON.stringify({
        preview: true,
        filters: true,
        multipleValues: true,
        nativeCache: true,
        sourceRefresh: true,
        visibleRefresh: true,
        reopened: true,
        editedWithoutDuplicateSheets: true,
        exportPreservesNativeParts: true,
      }),
    );
  } catch (error) {
    console.error(
      JSON.stringify({
        pivots: readSpreadsheetPivots(getNoteById(note.id).file_data),
        total: saved().Sheets.Summary ? total() : null,
      }),
    );
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
