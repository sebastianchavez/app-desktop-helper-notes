const { app, BrowserWindow } = require('electron');
const { readFileSync, readdirSync, mkdirSync, writeFileSync } = require('fs');
const path = require('path');
const { zipSync, strToU8 } = require('fflate');
const assert = require('node:assert/strict');
const output = path.resolve(__dirname, '../release/word-clipboard-qa');
mkdirSync(output, { recursive: true });
app.setPath('userData', path.join(output, 'profile'));
app.disableHardwareAcceleration();

/** Crea un Word aislado con cuerpo, tabla y cuadro flotante; retorna bytes sin datos del usuario. */
function fixture() {
  return zipSync({
    '[Content_Types].xml': strToU8(
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>',
    ),
    '_rels/.rels': strToU8(
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>',
    ),
    'word/document.xml': strToU8(
      '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:v="urn:schemas-microsoft-com:vml"><w:body><w:p><w:r><w:t>BODY_ANCHOR</w:t></w:r><w:r><w:pict><v:rect id="Box" filled="f" stroked="f" style="position:absolute;margin-left:40pt;margin-top:180pt;width:400pt;height:200pt;z-index:2;mso-position-horizontal-relative:page;mso-position-vertical-relative:page"><v:textbox inset="0,0,0,0"><w:txbxContent><w:p><w:r><w:t>BOX_TARGET</w:t></w:r></w:p></w:txbxContent></v:textbox></v:rect></w:pict></w:r></w:p><w:p><w:r><w:t>BODY_TARGET</w:t></w:r></w:p><w:tbl><w:tblPr><w:tblW w:w="6000" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="6000"/></w:tblGrid><w:tr><w:tc><w:p><w:r><w:t>CELL_TARGET</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="720" w:bottom="720" w:left="720" w:right="720"/></w:sectPr></w:body></w:document>',
    ),
  });
}

/** Espera la carga del adaptador; recibe ventana y expresión y retorna cuando se cumple. */
async function waitFor(window, expression) {
  for (let i = 0; i < 150; i++) {
    if (await window.webContents.executeJavaScript(expression)) return;
    await new Promise((resolve) => setTimeout(resolve, 100));
  }
  throw new Error(`No se cumplió: ${expression}`);
}

/** Hace clic real en el texto indicado para fijar la historia activa del editor. */
async function clickText(window, text, button = 'left') {
  const point = await window.webContents.executeJavaScript(
    `(()=>{const el=Array.from(document.querySelectorAll('.layout-run-text')).find(el=>el.textContent.includes(${JSON.stringify(text)}));if(!el)throw Error('Texto no visible');const r=el.getBoundingClientRect();return {x:Math.round(r.left+20),y:Math.round(r.top+r.height/2)};})()`,
  );
  for (const type of ['mouseDown', 'mouseUp'])
    window.webContents.sendInputEvent({ type, ...point, button, clickCount: 1 });
  await new Promise((resolve) => setTimeout(resolve, 100));
}

/** Reproduce eventos de portapapeles con payload propio, sin leer ni modificar el portapapeles del usuario. */
async function verifyClipboard() {
  const window = new BrowserWindow({
    show: false,
    width: 1500,
    height: 1100,
    webPreferences: { offscreen: true, backgroundThrottling: false, contextIsolation: true },
  });
  window.webContents.on('console-message', (event) => {
    if (/error|ACTIVE/i.test(event.message)) console.log(event.message);
  });
  try {
    const directory = path.resolve(__dirname, '../dist/app-desktop-helper-notes/browser');
    const { pathToFileURL } = require('url');
    const style = readdirSync(directory).find(
      (name) => name.startsWith('styles-') && name.endsWith('.css'),
    );
    const page = path.join(output, 'test.html');
    writeFileSync(
      page,
      `<!doctype html><html><head><meta charset="utf-8"><link rel="stylesheet" href="${pathToFileURL(path.join(directory, style)).href}"></head><body></body></html>`,
    );
    await window.loadFile(page);
    const chunk = readdirSync(directory).find(
      (name) =>
        name.startsWith('chunk-') &&
        name.endsWith('.js') &&
        readFileSync(path.join(directory, name), 'utf8').includes('as mountWordEditor'),
    );
    await window.webContents.executeJavaScript(`(async()=>{
      window.qaErrors=[];
      window.qaTransfer=new DataTransfer();
      window.electronAPI={onWordClipboardCommand:callback=>{window.qaNativeCommand=callback;return ()=>{}},wordClipboardExecute:async command=>{
        await new Promise(resolve=>setTimeout(resolve,30));
        document.activeElement.dispatchEvent(new ClipboardEvent(command==='pastePlain'?'paste':command,{clipboardData:window.qaTransfer,bubbles:true,cancelable:true}));
      },wordClipboardWriteImage:async base64=>{window.qaCopiedImage=base64;}};
      window.qaPaste=(text,html='')=>{const transfer=new DataTransfer();transfer.setData('text/plain',text);if(html)transfer.setData('text/html',html);document.activeElement.dispatchEvent(new ClipboardEvent('paste',{clipboardData:transfer,bubbles:true,cancelable:true}));};
      window.qaPasteImage=async()=>{const canvas=document.createElement('canvas');canvas.width=24;canvas.height=16;canvas.getContext('2d').fillRect(0,0,24,16);const blob=await new Promise(resolve=>canvas.toBlob(resolve,'image/png'));const transfer=new DataTransfer();transfer.items.add(new File([blob],'clipboard.png',{type:'image/png'}));document.activeElement.dispatchEvent(new ClipboardEvent('paste',{clipboardData:transfer,bubbles:true,cancelable:true}));};
      const {mountWordEditor}=await import(${JSON.stringify(pathToFileURL(path.join(directory, chunk)).href)});
      const host=document.createElement('div');host.style.cssText='position:fixed;inset:0;z-index:9999;background:white';document.body.append(host);
      window.qaMount=await mountWordEditor(host,Uint8Array.from(atob(${JSON.stringify(Buffer.from(fixture()).toString('base64'))}),c=>c.charCodeAt(0)),{title:'Clipboard QA',language:'es',save:()=>{},print:()=>{},open:()=>{},clipboardError:key=>{if(key)window.qaErrors.push(key)}});
      window.qaEditor=await window.qaMount.ready;
    })()`);
    await waitFor(window, '!!window.qaEditor');
    await clickText(window, 'BOX_TARGET');
    await window.webContents.executeJavaScript(
      `window.qaPaste('PASTE_BOX','<p><b>PASTE_BOX</b></p>')`,
    );
    await waitFor(window, `document.body.innerText.includes('PASTE_BOX')`);
    // Clic derecho: el menú puede quitar el foco del documento, pero el rango de Word debe conservarse.
    await clickText(window, 'PASTE_BOX', 'right');
    await waitFor(window, `!!document.querySelector('.docx-contextmenu')`);
    await window.webContents.executeJavaScript(
      `(()=>{window.qaTransfer=new DataTransfer();window.qaTransfer.setData('text/plain','CONTEXT_BOX');const menu=document.querySelector('.docx-contextmenu');const item=Array.from(menu.querySelectorAll('[role="menuitem"],button')).find(el=>el.textContent.trim().startsWith('Pegar')&&!el.textContent.includes('formato'));if(!item)throw Error('Falta Pegar en el menú');item.click();})()`,
    );
    await waitFor(window, `document.body.innerText.includes('CONTEXT_BOX')`);
    await clickText(window, 'BODY_TARGET');
    await window.webContents.executeJavaScript(
      `window.qaPaste('PASTE_BODY','<p><b>PASTE_BODY</b></p>')`,
    );
    await waitFor(window, `document.body.innerText.includes('PASTE_BODY')`);
    // Copiar desde el mismo Word conserva texto y fragmento enriquecido, sin leer el portapapeles real.
    for (const type of ['keyDown', 'keyUp'])
      window.webContents.sendInputEvent({ type, keyCode: 'HOME' });
    for (const type of ['keyDown', 'keyUp'])
      window.webContents.sendInputEvent({ type, keyCode: 'END', modifiers: ['shift'] });
    const copied = await window.webContents.executeJavaScript(
      `(()=>{const data=new DataTransfer();document.activeElement.dispatchEvent(new ClipboardEvent('copy',{clipboardData:data,bubbles:true,cancelable:true}));window.qaCopiedText=data;return data.getData('text/plain')})()`,
    );
    assert(copied.includes('PASTE_BODY'), 'No se copió el texto seleccionado del Word.');
    await clickText(window, 'CELL_TARGET');
    await window.webContents.executeJavaScript(
      `document.activeElement.dispatchEvent(new ClipboardEvent('paste',{clipboardData:window.qaCopiedText,bubbles:true,cancelable:true}))`,
    );
    await waitFor(window, `document.body.innerText.split('PASTE_BODY').length >= 3`);
    await window.webContents.executeJavaScript(`window.qaPaste('PASTE_CELL')`);
    await waitFor(window, `document.body.innerText.includes('PASTE_CELL')`);
    await clickText(window, 'BODY_ANCHOR');
    await window.webContents.executeJavaScript(
      `window.qaPaste('TABLE_A\tTABLE_B','<table><tr><td>TABLE_A</td><td>TABLE_B</td></tr></table>')`,
    );
    await waitFor(window, `document.body.innerText.includes('TABLE_A')`);
    await clickText(window, 'PASTE_BODY');
    window.webContents.sendInputEvent({ type: 'keyDown', keyCode: 'END' });
    window.webContents.sendInputEvent({ type: 'keyUp', keyCode: 'END' });
    await window.webContents.executeJavaScript('window.qaPasteImage()');
    await waitFor(
      window,
      `!!document.querySelector('.docx-paginated-surface img')?.naturalWidth || window.qaErrors.length>0`,
    );
    assert.deepEqual(await window.webContents.executeJavaScript('window.qaErrors'), []);
    const imagePoint = await window.webContents.executeJavaScript(
      `(()=>{const r=document.querySelector('.docx-paginated-surface img').getBoundingClientRect();return {x:Math.round(r.left+r.width/2),y:Math.round(r.top+r.height/2)}})()`,
    );
    for (const type of ['mouseDown', 'mouseUp'])
      window.webContents.sendInputEvent({ type, ...imagePoint, button: 'left', clickCount: 1 });
    await waitFor(window, '!!window.qaEditor.getSelectedImage()');
    await window.webContents.executeJavaScript(
      `document.activeElement.dispatchEvent(new ClipboardEvent('copy',{clipboardData:new DataTransfer(),bubbles:true,cancelable:true}))`,
    );
    await waitFor(window, `!!window.qaCopiedImage || window.qaErrors.length>0`);
    assert(
      await window.webContents.executeJavaScript('!!window.qaCopiedImage'),
      'No se pudo copiar la imagen del Word.',
    );
    const result = await window.webContents.executeJavaScript(
      `(async()=>({data:Array.from(new Uint8Array(await window.qaEditor.save())),visible:document.body.innerText.includes('PASTE_BOX'),errors:window.qaErrors}))()`,
    );
    const { unzipSync, strFromU8 } = require('fflate');
    const xml = strFromU8(unzipSync(Uint8Array.from(result.data))['word/document.xml']);
    assert(xml.includes('PASTE_BODY') && xml.includes('PASTE_CELL'));
    assert(xml.includes('TABLE_A') && xml.includes('TABLE_B'));
    assert((xml.match(/<w:tbl[\s>]/g) || []).length >= 2, 'La tabla pegada perdió su estructura.');
    assert(
      /<w:txbxContent[\s\S]*?CONTEXT_BOX[\s\S]*?<\/w:txbxContent>/.test(xml),
      'El pegado contextual salió del cuadro.',
    );
    assert.deepEqual(result.errors, []);
    writeFileSync(path.join(output, 'paste.png'), (await window.webContents.capturePage()).toPNG());
    assert(result.visible, 'El pegado no quedó visible.');
    assert(
      /<w:txbxContent[\s\S]*?PASTE_BOX[\s\S]*?<\/w:txbxContent>/.test(xml),
      'El pegado salió del cuadro de texto.',
    );
    writeFileSync(path.join(output, 'clipboard.docx'), Buffer.from(result.data));
    await window.webContents.executeJavaScript(
      `window.qaEditor.load(Uint8Array.from(${JSON.stringify(result.data)}))`,
    );
    await waitFor(
      window,
      `document.body.innerText.includes('CONTEXT_BOX') && document.body.innerText.includes('PASTE_BODY') && document.body.innerText.includes('PASTE_CELL') && document.body.innerText.includes('TABLE_B') && !!document.querySelector('.docx-paginated-surface img')?.naturalWidth`,
    );
    console.log(
      JSON.stringify({
        bodyPaste: true,
        cellPaste: true,
        textBoxPaste: true,
        contextMenuPaste: true,
        tablePaste: true,
        imagePaste: true,
        imageCopy: true,
        internalTextCopyPaste: true,
        saved: true,
        reopened: true,
      }),
    );
  } catch (error) {
    console.error(
      await window.webContents.executeJavaScript(
        `JSON.stringify({errors:window.qaErrors,scope:window.qaEditor?.snapshot().scope,active:document.activeElement?.className,images:document.querySelectorAll('img').length,text:document.body.innerText.slice(0,1000)})`,
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
  .then(verifyClipboard)
  .then(() => app.exit(0))
  .catch((error) => {
    console.error(error);
    app.exit(1);
  });
