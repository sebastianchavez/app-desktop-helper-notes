const { app, BrowserWindow, session } = require('electron');
const { mkdirSync, writeFileSync } = require('fs');
const path = require('path');
const assert = require('node:assert/strict');
const XLSX = require('xlsx');
const output = path.resolve(__dirname, '../release/spreadsheet-qa');
mkdirSync(output, { recursive: true });
app.setPath('userData', path.join(output, 'profile'));
app.disableHardwareAcceleration();

/** Espera una condición del renderer; recibe ventana/expresión y rechaza si no se cumple. */
async function wait(window, expression) {
  for (let i = 0; i < 200; i++) {
    if (await window.webContents.executeJavaScript(expression)) return;
    await new Promise((r) => setTimeout(r, 100));
  }
  throw Error('No se cumplió: ' + expression);
}

/** Verifica controles reales y SQLite en un perfil aislado; no recibe ni modifica archivos del usuario. */
async function verify() {
  require('../electron-dist/db/db').initDb();
  require('../electron-dist/ipc-handlers').registerIpcHandlers();
  const { createNote, updateNote, getNoteById } = require('../electron-dist/db/note.repository');
  const book = XLSX.utils.book_new();
  const data = XLSX.utils.aoa_to_sheet([
    [10, 20, 30],
    ['Nombre', 'Apellido'],
    ['Conservar', 'Cubierta'],
  ]);
  data.C1 = { t: 'n', v: 30, f: 'SUM(A1:B1)' };
  XLSX.utils.book_append_sheet(book, data, 'Datos');
  const summary = XLSX.utils.aoa_to_sheet([[60]]);
  summary.A1 = { t: 'n', v: 60, f: 'Datos!C1*2' };
  XLSX.utils.book_append_sheet(book, summary, 'Resumen');
  const note = createNote(null, 'Excel QA');
  updateNote(note.id, {
    file_type: 'xlsx',
    file_data: Buffer.from(XLSX.write(book, { type: 'buffer', bookType: 'xlsx' })).toString(
      'base64',
    ),
  });
  createNote(null, 'Otra QA');
  session.defaultSession.webRequest.onBeforeRequest((details, callback) =>
    callback({ cancel: /^https?:/.test(details.url) }),
  );
  const window = new BrowserWindow({
    show: false,
    width: 1800,
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
  /** Selecciona por eventos de ratón un rango; recibe coordenadas base cero y retorna al actualizar Angular. */
  const select = async (r, c, er = r, ec = c) => {
    await run(
      `(()=>{const cell=(r,c)=>document.querySelectorAll('.xlsx-grid tbody tr')[r].querySelectorAll('.xlsx-cell')[c];cell(${r},${c}).dispatchEvent(new MouseEvent('mousedown',{bubbles:true,button:0}));cell(${er},${ec}).dispatchEvent(new MouseEvent('mouseenter'));document.dispatchEvent(new MouseEvent('mouseup',{bubbles:true}));})()`,
    );
    await new Promise((resolve) => setTimeout(resolve, 100));
  };
  /** Pulsa el control por su icono y espera al guardado; recibe nombre del icono y retorna al finalizar IPC. */
  const button = async (icon) => {
    await run(
      `Array.from(document.querySelectorAll('.xlsx-toolbar button')).find(b=>b.querySelector('.material-symbols-outlined')?.textContent.trim()===${JSON.stringify(icon)}).click()`,
    );
    await new Promise((resolve) => setTimeout(resolve, 100));
    await wait(window, `document.querySelector('.xlsx-app')?.getAttribute('aria-busy')==='false'`);
    assert.equal(await run(`document.querySelector('.xlsx-error')?.innerText??''`), '');
  };
  /** Recupera los bytes guardados en SQLite como libro; no recibe argumentos ni usa el renderer para comprobar resultados. */
  const saved = () =>
    XLSX.read(Buffer.from(getNoteById(note.id).file_data, 'base64'), {
      type: 'buffer',
      cellStyles: true,
    });
  try {
    await window.loadFile(
      path.resolve(__dirname, '../dist/app-desktop-helper-notes/browser/index.html'),
    );
    await wait(
      window,
      `Array.from(document.querySelectorAll('app-note-list h3')).some(n=>n.textContent.includes('Excel QA'))`,
    );
    await run(
      `Array.from(document.querySelectorAll('app-note-list h3')).find(n=>n.textContent.includes('Excel QA')).closest('[draggable="true"]').click()`,
    );
    await wait(window, `!!document.querySelector('.xlsx-grid .xlsx-cell')`);
    await select(2, 0, 2, 1);
    await button('cell_merge');
    assert.equal(saved().Sheets.Datos['!merges'].length, 1);
    assert.equal(saved().Sheets.Datos.B3.v, 'Cubierta');
    assert.equal(
      await run(
        `document.querySelectorAll('.xlsx-grid tbody tr')[2].querySelector('.xlsx-cell').colSpan`,
      ),
      2,
    );
    await button('cell_split');
    assert.equal(saved().Sheets.Datos['!merges']?.length ?? 0, 0);
    await select(0, 0);
    await button('format_bold');
    await run(
      `(()=>{const select=document.querySelector('select[aria-label="Formato numérico"],select[aria-label="Number format"]');select.value='0.00%';select.dispatchEvent(new Event('change',{bubbles:true}));})()`,
    );
    await wait(window, `document.querySelector('.xlsx-app')?.getAttribute('aria-busy')==='false'`);
    assert.equal(saved().Sheets.Datos.A1.w, '1000.00%');
    await select(0, 0);
    await run(
      `(()=>{const input=document.querySelector('.xlsx-formula-input');input.focus();input.value='40';input.dispatchEvent(new Event('input',{bubbles:true}));input.dispatchEvent(new KeyboardEvent('keydown',{key:'Enter',bubbles:true}));})()`,
    );
    await new Promise((resolve) => setTimeout(resolve, 150));
    await wait(window, `document.querySelector('.xlsx-app')?.getAttribute('aria-busy')==='false'`);
    assert.equal(saved().Sheets.Datos.A1.v, 40);
    assert.equal(saved().Sheets.Datos.C1.v, 60);
    assert.equal(saved().Sheets.Resumen.A1.v, 120);
    await select(1, 0);
    await button('playlist_add');
    assert.equal(saved().Sheets.Datos.A4.v, 'Conservar');
    await button('undo');
    await button('redo');
    await select(0, 0);
    await button('view_column');
    assert.equal(saved().Sheets.Datos.D1.f, 'SUM(A1:C1)');
    assert.equal(saved().Sheets.Resumen.A1.f, 'Datos!D1*2');
    await button('undo');
    // Portapapeles sintético exclusivo de esta ventana: no toca el del usuario.
    await run(
      `Object.defineProperty(navigator,'clipboard',{configurable:true,value:{writeText:async text=>{window.qaClipboard=text},readText:async()=>window.qaClipboard??''}})`,
    );
    await select(0, 2);
    await button('content_copy');
    await select(0, 3);
    await button('content_paste');
    assert.equal(saved().Sheets.Datos.D1.f, 'SUM(B1:C1)');
    assert.equal(saved().Sheets.Datos.D1.v, 80);
    await select(1, 0, 3, 0);
    await button('sort_by_alpha');
    assert.equal(saved().Sheets.Datos.A2.v, 'Conservar');
    assert.equal(saved().Sheets.Datos.B2.v, 'Cubierta');
    await button('undo');
    await run(
      `document.querySelectorAll('.xlsx-grid tbody tr')[3].querySelector('.xlsx-row-header').click()`,
    );
    await button('move_up');
    assert.equal(saved().Sheets.Datos.A3.v, 'Conservar');
    await button('undo');
    await select(3, 5);
    await run(
      `(()=>{const input=document.querySelector('.xlsx-formula-input');input.value='Texto muy largo para comprobar el ajuste de la celda';input.dispatchEvent(new Event('input',{bubbles:true}));input.dispatchEvent(new KeyboardEvent('keydown',{key:'Enter',bubbles:true}));})()`,
    );
    await new Promise((resolve) => setTimeout(resolve, 150));
    await wait(window, `document.querySelector('.xlsx-app')?.getAttribute('aria-busy')==='false'`);
    await select(3, 5);
    await button('wrap_text');
    assert(
      await run(
        `(()=>{const c=document.querySelectorAll('.xlsx-grid tbody tr')[3].querySelectorAll('.xlsx-cell')[5];return getComputedStyle(c.querySelector('.xlsx-cell-value')).whiteSpace==='pre-wrap'&&c.getBoundingClientRect().height>25})()`,
      ),
    );
    assert.equal(
      await run(`getComputedStyle(document.querySelector('.xlsx-cell')).backgroundColor`),
      'rgb(255, 255, 255)',
    );
    // Cambiar de nota y volver vuelve a leer SQLite; el formato no depende de mapas en memoria.
    for (const title of ['Otra QA', 'Excel QA']) {
      await run(
        `Array.from(document.querySelectorAll('app-note-list h3')).find(n=>n.textContent.includes(${JSON.stringify(title)})).closest('[draggable="true"]').click()`,
      );
      await new Promise((resolve) => setTimeout(resolve, 250));
    }
    await wait(window, `!!document.querySelector('.xlsx-grid .xlsx-cell')`);
    assert(await run(`document.querySelector('.xlsx-grid .xlsx-cell').innerText.includes('4000')`));
    writeFileSync(
      path.join(output, 'edited.xlsx'),
      Buffer.from(getNoteById(note.id).file_data, 'base64'),
    );
    writeFileSync(
      path.join(output, 'spreadsheet.png'),
      (await window.webContents.capturePage()).toPNG(),
    );
    console.log(
      JSON.stringify({
        merge: true,
        unmerge: true,
        coveredValuesPreserved: true,
        formatPersisted: true,
        formulas: true,
        crossSheetRecalculation: true,
        insertRow: true,
        insertColumn: true,
        relativeFormulaPaste: true,
        sortRows: true,
        moveRows: true,
        wrapText: true,
        readableInDarkMode: true,
        undoRedo: true,
        reopened: true,
      }),
    );
  } catch (error) {
    writeFileSync(
      path.join(output, 'failure.png'),
      (await window.webContents.capturePage()).toPNG(),
    );
    console.error(await run(`document.querySelector('.xlsx-app')?.innerText.slice(0,1400)`));
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
