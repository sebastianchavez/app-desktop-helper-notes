const { app, BrowserWindow, session, dialog } = require('electron');
const { mkdirSync, writeFileSync, readFileSync } = require('fs');
const path = require('path');
const assert = require('node:assert/strict');
const XLSX = require('xlsx');
const output = path.resolve(__dirname, '../release/formula-qa');
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

/** Comprueba fórmulas, nombres, fechas y cachés con SQLite y documentos generados exclusivamente para la prueba. */
async function verify() {
  require('../electron-dist/db/db').initDb();
  require('../electron-dist/ipc-handlers').registerIpcHandlers();
  const { createNote, updateNote, getNoteById } = require('../electron-dist/db/note.repository');
  const { strFromU8, unzipSync } = require('fflate');
  const book = XLSX.utils.book_new(),
    data = XLSX.utils.aoa_to_sheet([
      ['Region', 'Product', 'Amount', 'Units'],
      ['North', 'A', 100, 2],
      ['North', 'B', 200, 3],
      ['South', 'A', 50, 4],
    ]);
  const formulas = {
    E1: ['SUMIFS(C2:C4,A2:A4,"North")', 300],
    F1: ['_xlfn.XLOOKUP("South",A2:A4,C2:C4)', 50],
    G1: ['DATE(2026,10,15)', (Date.UTC(2026, 9, 15) - Date.UTC(1904, 0, 1)) / 86400000],
    H1: ['YEAR(G1)', 2026],
    I1: ['SUM(Sales)*Rate', 70],
    J1: ['_xlfn.UNKNOWNARRAY(C2:C4,A2:A4="North")', 777],
    M1: ['J1+1', 778],
  };
  for (const [address, [f, v]] of Object.entries(formulas)) data[address] = { t: 'n', v, f };
  data.J1.F = 'J1:J3';
  data.K1 = { t: 'e', v: 7 };
  data.L1 = { t: 'e', v: 7, f: '1/0' };
  data['!ref'] = 'A1:M4';
  data.G1.z = 'yyyy-mm-dd';
  XLSX.utils.book_append_sheet(book, data, 'Data');
  XLSX.utils.book_append_sheet(book, XLSX.utils.aoa_to_sheet([[35]]), 'Other');
  book.Sheets.Other.A1 = { t: 'n', v: 35, f: 'SUM(Data!C2:C4)*Rate' };
  book.Workbook = {
    WBProps: { date1904: true },
    Names: [
      { Name: 'Rate', Ref: '0.1' },
      { Name: 'Sales', Ref: "'Data'!$C$2:$C$4" },
      { Name: 'Rate', Ref: '0.2', Sheet: 0 },
    ],
  };
  const note = createNote(null, 'Formula QA ' + Date.now());
  updateNote(note.id, {
    file_type: 'xlsx',
    file_data: Buffer.from(XLSX.write(book, { type: 'buffer', bookType: 'xlsx' })).toString(
      'base64',
    ),
  });
  const other = createNote(null, 'Other formula QA ' + Date.now());
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
  /** Recibe título; selecciona una nota en el listado real. */
  const selectNote = async (title) => {
    await run(
      `Array.from(document.querySelectorAll('app-note-list h3')).find(n=>n.textContent.includes(${JSON.stringify(title)})).closest('[draggable="true"]').click()`,
    );
    await new Promise((r) => setTimeout(r, 150));
  };
  /** Recibe coordenadas y texto; edita desde la barra de fórmulas y espera el guardado real. */
  const edit = async (row, col, value) => {
    await run(
      `(()=>{const c=document.querySelectorAll('.xlsx-grid tbody tr')[${row}].querySelectorAll('.xlsx-cell')[${col}];c.dispatchEvent(new MouseEvent('mousedown',{bubbles:true,button:0}));document.dispatchEvent(new MouseEvent('mouseup',{bubbles:true}));})()`,
    );
    await run(
      `(()=>{const input=document.querySelector('.xlsx-formula-input');input.value=${JSON.stringify(value)};input.dispatchEvent(new Event('input',{bubbles:true}));input.dispatchEvent(new KeyboardEvent('keydown',{key:'Enter',bubbles:true}));})()`,
    );
    await new Promise((r) => setTimeout(r, 220));
    await wait(window, `document.querySelector('.xlsx-app')?.getAttribute('aria-busy')==='false'`);
  };
  /** Retorna el último libro persistido con prefijos futuros intactos. */
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
    assert(
      await run(`document.querySelector('.xlsx-formula-compatibility')?.textContent.includes('1')`),
    );
    assert(await run(`document.querySelector('.xlsx-grid').textContent.includes('2026-10-15')`));
    await edit(1, 2, '400');
    const current = saved();
    assert.equal(current.Sheets.Data.E1.v, 600);
    assert.equal(current.Sheets.Data.F1.v, 50);
    assert.equal(current.Sheets.Data.H1.v, 2026);
    assert.equal(current.Sheets.Data.I1.v, 130);
    assert.equal(current.Sheets.Other.A1.v, 65);
    assert.equal(current.Sheets.Data.J1.v, 777);
    assert.equal(current.Sheets.Data.M1.v, 778);
    assert.equal(current.Workbook.WBProps.date1904, true);
    assert.equal(current.Workbook.Names.length, 3);
    assert.equal(current.Sheets.Data.K1.t, 'e');
    assert.equal(current.Sheets.Data.K1.v, 7);
    assert.equal(current.Sheets.Data.L1.t, 'e');
    assert.equal(current.Sheets.Data.L1.v, 7);
    const xml = strFromU8(
      unzipSync(Buffer.from(getNoteById(note.id).file_data, 'base64'))['xl/worksheets/sheet1.xml'],
    );
    assert.match(xml, /_xlfn.XLOOKUP/);
    assert.match(xml, /<f[^>]*t="array"[^>]*>/);
    assert.match(xml, /ref="J1:J3"/);
    await edit(0, 14, '=BUSCARX("South";A2:A4;C2:C4)');
    assert.equal(saved().Sheets.Data.O1.v, 50);
    assert.match(saved().Sheets.Data.O1.f, /_xlfn.XLOOKUP/);
    await selectNote(other.title);
    await selectNote(note.title);
    await wait(window, `!!document.querySelector('.xlsx-grid .xlsx-cell')`);
    await wait(window, `!!document.querySelector('.xlsx-formula-compatibility')`);
    writeFileSync(
      path.join(output, 'compatible.xlsx'),
      Buffer.from(getNoteById(note.id).file_data, 'base64'),
    );
    writeFileSync(
      path.join(output, 'result.png'),
      (await window.webContents.capturePage()).toPNG(),
    );
    await edit(0, 9, '=UNKNOWN(1)');
    assert.equal(saved().Sheets.Data.J1.t, 'e');
    assert.equal(saved().Sheets.Data.J1.v, 29);
    console.log(
      JSON.stringify({
        multiCriteria: true,
        modernLookup: true,
        localizedInput: true,
        namedRangesAndScope: true,
        date1904: true,
        dateDisplay: true,
        typedErrors: true,
        cachedUnsupported: true,
        arrayDefinitionPreserved: true,
        changedUnsupportedDoesNotReuseOldValue: true,
        reopened: true,
      }),
    );
  } catch (error) {
    console.error(
      await run(
        `document.querySelector('.xlsx-error')?.textContent??document.body.textContent.slice(-500)`,
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
