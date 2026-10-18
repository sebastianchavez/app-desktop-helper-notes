const { app, BrowserWindow, dialog, session } = require('electron');
const {
  mkdirSync,
  writeFileSync,
  readFileSync,
  readdirSync,
  existsSync,
  unlinkSync,
} = require('fs');
const path = require('path');
const assert = require('node:assert/strict');
const { zipSync, strToU8, unzipSync, strFromU8 } = require('fflate');

const output = path.resolve(__dirname, '../release/word-editor-qa');
mkdirSync(output, { recursive: true });
app.setPath('userData', path.join(output, 'profile'));
app.disableHardwareAcceleration();

/** Espera una condición asíncrona de QA; recibe función y descripción y retorna al cumplirse. */
async function waitFor(check, label) {
  const deadline = Date.now() + 25000;
  while (Date.now() < deadline) {
    if (await check()) return;
    await new Promise((resolve) => setTimeout(resolve, 100));
  }
  throw new Error(`No se cumplió: ${label}`);
}

/** Genera un Word sintético con tabla; no recibe datos ni retorna archivos del usuario. */
function fixtureBytes() {
  return zipSync({
    '[Content_Types].xml': strToU8(
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>',
    ),
    '_rels/.rels': strToU8(
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>',
    ),
    'word/document.xml': strToU8(
      '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>DOCUMENTO EDITABLE</w:t></w:r></w:p><w:tbl><w:tblPr><w:tblW w:w="9000" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="4500"/><w:gridCol w:w="4500"/></w:tblGrid><w:tr><w:tc><w:p><w:r><w:t>Celda editable</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>Otra celda</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:p/><w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1080" w:bottom="1080" w:left="1080" w:right="1080"/></w:sectPr></w:body></w:document>',
    ),
  });
}

/** Prueba tabla, imagen y página con el adaptador de producción; retorna DOCX exportado para verificar XML. */
async function verifyEditingCommands(window) {
  const directory = path.resolve(__dirname, '../dist/app-desktop-helper-notes/browser');
  const chunk = readdirSync(directory).find(
    (name) =>
      name.startsWith('chunk-') &&
      name.endsWith('.js') &&
      readFileSync(path.join(directory, name), 'utf8').includes('as mountWordEditor'),
  );
  assert(chunk, 'No se encontró el adaptador compilado.');
  const result = await window.webContents.executeJavaScript(`(async()=>{
    const {mountWordEditor}=await import('./${chunk}');
    const host=document.createElement('div');host.style.cssText='position:fixed;inset:0;z-index:9999;background:white';document.body.append(host);
    const bytes=Uint8Array.from(atob(${JSON.stringify(Buffer.from(fixtureBytes()).toString('base64'))}),c=>c.charCodeAt(0));
    const mounted=await mountWordEditor(host,bytes,{title:'QA commands',language:'es',save:()=>{},open:()=>{}});
    try{
      const editor=await mounted.ready;
      const first=editor.snapshot().selection.from;
      const table=editor.exec({type:'insertTable',target:{paraId:first.paraId},rows:2,cols:2});
      const page=editor.exec({type:'setPageSetup',orientation:'landscape',marginLeft:1800});
      editor.exec({type:'setSelection',anchor:{paraId:first.paraId}});
      const canvas=document.createElement('canvas');canvas.width=20;canvas.height=20;canvas.getContext('2d').fillRect(0,0,20,20);
      const png=Uint8Array.from(atob(canvas.toDataURL('image/png').split(',')[1]),c=>c.charCodeAt(0));
      const image=await editor.executeImageCommand({type:'insertImage',data:png,mime:'image/png',widthPoints:30,heightPoints:30});
      const resize=editor.exec({type:'setImageProperties',widthEmu:914400,heightEmu:457200});
      const output=Array.from(new Uint8Array(await editor.save()));
      return {table,page,image,resize,output};
    }finally{mounted.destroy();host.remove();}
  })()`);
  for (const name of ['table', 'page', 'image', 'resize'])
    assert(result[name].ok, `${name}: ${JSON.stringify(result[name])}`);
  const entries = unzipSync(Uint8Array.from(result.output));
  const xml = strFromU8(entries['word/document.xml']);
  assert((xml.match(/<w:tbl[\s>]/g) || []).length >= 2, 'La tabla nueva no se guardó.');
  assert(xml.includes('landscape'), 'La orientación no se guardó.');
  assert(xml.includes('914400') && xml.includes('457200'), 'El tamaño de la imagen no se guardó.');
  assert(
    Object.keys(entries).some((name) => name.startsWith('word/media/')),
    'La imagen no se guardó.',
  );
  writeFileSync(path.join(output, 'table-image-page.docx'), Buffer.from(result.output));
  console.log(
    JSON.stringify({
      tableInserted: true,
      imageInserted: true,
      imageResized: true,
      pageChanged: true,
    }),
  );
}

/**
 * Comprueba edición nativa, SQLite, cambio de nota, reapertura y exportación del DOCX.
 * Recibe opcionalmente la ruta del CV en argv[2]; trabaja únicamente en un perfil de prueba.
 * @returns Promesa resuelta después de comprobar bytes e imágenes y guardar una captura local.
 */
async function verifyEditor() {
  const { initDb } = require('../electron-dist/db/db');
  const { registerIpcHandlers } = require('../electron-dist/ipc-handlers');
  const { createNote, updateNote, getNoteById } = require('../electron-dist/db/note.repository');
  initDb();
  registerIpcHandlers();
  const original = process.argv[2] ? readFileSync(process.argv[2]) : fixtureBytes();
  const note = createNote(null, 'Word editor QA');
  updateNote(note.id, {
    file_type: 'docx',
    file_data: Buffer.from(original).toString('base64'),
    content: '<p>QA</p>',
  });
  const other = createNote(null, 'Otra nota QA');
  const exportedPath = path.join(output, 'edited.docx');
  if (existsSync(exportedPath)) unlinkSync(exportedPath);
  dialog.showSaveDialog = async () => ({ canceled: false, filePath: exportedPath });
  session.defaultSession.webRequest.onBeforeRequest((details, callback) =>
    callback({ cancel: /^https?:/.test(details.url) }),
  );
  const window = new BrowserWindow({
    show: false,
    width: 1920,
    height: 1080,
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
  const execute = (code) => window.webContents.executeJavaScript(code);
  const clickNote = (title) =>
    execute(
      `Array.from(document.querySelectorAll('app-note-list h3')).find(el=>el.textContent.includes(${JSON.stringify(title)})).closest('[draggable="true"]').click()`,
    );
  try {
    await window.loadFile(
      path.resolve(__dirname, '../dist/app-desktop-helper-notes/browser/index.html'),
    );
    await waitFor(
      () =>
        execute(
          `Array.from(document.querySelectorAll('app-note-list h3')).some(el=>el.textContent.includes('Word editor QA'))`,
        ),
      'listado',
    );
    await clickNote('Word editor QA');
    await waitFor(
      () =>
        execute(
          `document.querySelector('.word-host[aria-busy="false"] .layout-run-text') !== null`,
        ),
      'editor listo',
    );
    const target = process.argv[2] ? 'SEBASTIAN' : 'DOCUMENTO';
    const point = await execute(
      `(()=>{const el=Array.from(document.querySelectorAll('.word-host .layout-run-text')).find(el=>el.textContent.includes(${JSON.stringify(target)}));if(!el)throw Error('Texto no visible');const box=el.getBoundingClientRect();return {x:Math.round(box.left+30),y:Math.round(box.top+box.height/2),width:box.width,height:box.height};})()`,
    );
    assert(point.width > 0 && point.height > 0);
    window.webContents.sendInputEvent({
      type: 'mouseDown',
      x: point.x,
      y: point.y,
      button: 'left',
      clickCount: 1,
    });
    window.webContents.sendInputEvent({
      type: 'mouseUp',
      x: point.x,
      y: point.y,
      button: 'left',
      clickCount: 1,
    });
    await new Promise((resolve) => setTimeout(resolve, 200));
    await window.webContents.insertText(' HNQA ');
    await waitFor(() => getNoteById(note.id).content.includes('HNQA'), 'autosave SQLite');
    const saved = unzipSync(Buffer.from(getNoteById(note.id).file_data, 'base64'));
    assert(strFromU8(saved['word/document.xml']).includes('HNQA'));
    const source = unzipSync(original);
    for (const name of Object.keys(source).filter((name) => name.startsWith('word/media/')))
      assert.deepEqual(saved[name], source[name], `Imagen alterada: ${name}`);
    await clickNote('Otra nota QA');
    await waitFor(() => execute(`document.querySelector('.word-host') === null`), 'otra nota');
    await clickNote('Word editor QA');
    await waitFor(
      () =>
        execute(
          `document.querySelector('.word-host[aria-busy="false"]')?.textContent.includes('HNQA')`,
        ),
      'reapertura editada',
    );
    await execute(
      `Array.from(document.querySelectorAll('.document-toolbar button')).find(el=>el.textContent.includes('Guardar archivo')||el.textContent.includes('Save file')).click()`,
    );
    await waitFor(() => {
      try {
        return readFileSync(exportedPath).length > 0;
      } catch {
        return false;
      }
    }, 'exportación');
    assert(strFromU8(unzipSync(readFileSync(exportedPath))['word/document.xml']).includes('HNQA'));
    await window.webContents.capturePage();
    await new Promise((resolve) => setTimeout(resolve, 400));
    writeFileSync(
      path.join(output, 'word-editor.png'),
      (await window.webContents.capturePage()).toPNG(),
    );
    if (process.argv[2]) assert.deepEqual(readFileSync(process.argv[2]), original);
    await verifyEditingCommands(window);
    console.log(
      JSON.stringify({
        edited: true,
        autosaved: true,
        reopened: true,
        exported: true,
        mediaPreserved: true,
        originalUntouched: true,
      }),
    );
  } catch (error) {
    writeFileSync(
      path.join(output, 'failure.png'),
      (await window.webContents.capturePage()).toPNG(),
    );
    console.error(await execute('document.body.innerText.slice(-2500)'));
    throw error;
  } finally {
    window.destroy();
  }
}

app
  .whenReady()
  .then(verifyEditor)
  .then(() => app.exit(0))
  .catch((error) => {
    console.error(error);
    app.exit(1);
  });
