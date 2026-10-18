import { unzipSync, zipSync, strFromU8, strToU8 } from 'fflate';

const WORD = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const VML = 'urn:schemas-microsoft-com:vml';
const WP = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing';
const DRAWING = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const PICTURE = 'http://schemas.openxmlformats.org/drawingml/2006/picture';
const SHAPE = 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape';
const RELATIONSHIP = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const XMLNS = 'http://www.w3.org/2000/xmlns/';
const EMUS_PER_POINT = 12700;

/**
 * Convierte rectángulos VML simples de compatibilidad en imágenes y cuadros DrawingML editables.
 * @param bytes DOCX completo; el archivo almacenado no se cambia hasta que el usuario edita.
 * @returns DOCX con posiciones, tamaños, texto, relaciones y medios conservados. Formas no soportadas siguen intactas.
 * @throws Si el ZIP supera 200 MB descomprimidos o el XML principal no es válido.
 */
export function prepareWordForEditing(bytes: Uint8Array): Uint8Array {
  let totalBytes = 0;
  const entries = unzipSync(bytes, {
    filter: (entry) => {
      totalBytes += entry.originalSize;
      if (totalBytes > 200 * 1024 * 1024)
        throw new Error('El documento Word descomprimido supera los 200 MB.');
      return true;
    },
  });
  if (!entries['word/document.xml']) return bytes;
  const xml = new DOMParser().parseFromString(
    strFromU8(entries['word/document.xml']),
    'application/xml',
  );
  if (xml.querySelector('parsererror') || xml.documentElement.namespaceURI !== WORD)
    throw new Error('El documento Word contiene XML no válido.');
  let id = Math.max(
    0,
    ...Array.from(
      xml.getElementsByTagNameNS(WP, 'docPr'),
      (node) => Number(node.getAttribute('id')) || 0,
    ),
  );
  let converted = 0;
  for (const pict of Array.from(xml.getElementsByTagNameNS(WORD, 'pict'))) {
    if (pict.children.length !== 1) continue;
    const rect = pict.children[0];
    if (rect.namespaceURI !== VML || rect.localName !== 'rect' || rect.hasAttribute('href'))
      continue;
    const style = Object.fromEntries(
      (rect.getAttribute('style') ?? '').split(';').map((part) => {
        const colon = part.indexOf(':');
        return [
          part.slice(0, colon).trim().toLowerCase(),
          part
            .slice(colon + 1)
            .trim()
            .toLowerCase(),
        ];
      }),
    );
    if (
      style['position'] !== 'absolute' ||
      style['rotation'] ||
      style['flip'] ||
      style['mso-position-horizontal-relative'] !== 'page' ||
      style['mso-position-vertical-relative'] !== 'page'
    )
      continue;
    const left = lengthInEmus(style['margin-left'] ?? '0pt');
    const top = lengthInEmus(style['margin-top'] ?? '0pt');
    const width = lengthInEmus(style['width']);
    const height = lengthInEmus(style['height']);
    if ([left, top, width, height].some((value) => value === null) || width! <= 0 || height! <= 0)
      continue;
    const textbox = Array.from(rect.children).find(
      (node) => node.namespaceURI === VML && node.localName === 'textbox',
    );
    const image = Array.from(rect.children).find(
      (node) => node.namespaceURI === VML && node.localName === 'imagedata',
    );
    const content = textbox?.getElementsByTagNameNS(WORD, 'txbxContent')[0];
    const relation = image?.getAttributeNS(RELATIONSHIP, 'id');
    if (
      !!content === !!relation ||
      rect.getAttribute('filled') !== 'f' ||
      rect.getAttribute('stroked') !== 'f'
    )
      continue;
    if (
      image &&
      Array.from(image.attributes).some(
        (attribute) => /^crop/i.test(attribute.localName) && !/^0f?$/.test(attribute.value),
      )
    )
      continue;
    if (rect.children.length !== 1) continue;
    const inset = (textbox?.getAttribute('inset') ?? '7.2pt,3.6pt,7.2pt,3.6pt')
      .split(',')
      .map(lengthInEmus);
    if (content && (inset.length !== 4 || inset.some((value) => value === null))) continue;
    id++;
    const drawing = xml.createElementNS(WORD, 'w:drawing');
    const zIndex = Number(style['z-index']) || 0;
    const anchor = child(xml, drawing, WP, 'wp:anchor', {
      distT: '0',
      distB: '0',
      distL: '0',
      distR: '0',
      simplePos: '0',
      relativeHeight: String(Math.max(0, zIndex)),
      behindDoc: zIndex < 0 ? '1' : '0',
      locked: '0',
      layoutInCell: '0',
      allowOverlap: '1',
    });
    child(xml, anchor, WP, 'wp:simplePos', { x: '0', y: '0' });
    const horizontal = child(xml, anchor, WP, 'wp:positionH', { relativeFrom: 'page' });
    child(xml, horizontal, WP, 'wp:posOffset').textContent = String(left);
    const vertical = child(xml, anchor, WP, 'wp:positionV', { relativeFrom: 'page' });
    child(xml, vertical, WP, 'wp:posOffset').textContent = String(top);
    child(xml, anchor, WP, 'wp:extent', { cx: String(width), cy: String(height) });
    child(xml, anchor, WP, 'wp:effectExtent', { l: '0', t: '0', r: '0', b: '0' });
    child(xml, anchor, WP, 'wp:wrapNone');
    child(xml, anchor, WP, 'wp:docPr', {
      id: String(id),
      name: rect.getAttribute('id') || `Objeto ${id}`,
    });
    child(xml, anchor, WP, 'wp:cNvGraphicFramePr');
    const graphic = child(xml, anchor, DRAWING, 'a:graphic');
    const data = child(xml, graphic, DRAWING, 'a:graphicData', { uri: content ? SHAPE : PICTURE });
    if (content) {
      const shape = child(xml, data, SHAPE, 'wps:wsp');
      child(xml, shape, SHAPE, 'wps:cNvSpPr', { txBox: '1' });
      shapeProperties(xml, shape, SHAPE, 'wps:spPr', width!, height!);
      child(xml, shape, SHAPE, 'wps:txbx').append(content.cloneNode(true));
      const body = child(xml, shape, SHAPE, 'wps:bodyPr', {
        wrap: 'square',
        lIns: String(inset[0]),
        tIns: String(inset[1]),
        rIns: String(inset[2]),
        bIns: String(inset[3]),
        anchor: 't',
      });
      child(xml, body, DRAWING, 'a:noAutofit');
    } else {
      const picture = child(xml, data, PICTURE, 'pic:pic');
      const properties = child(xml, picture, PICTURE, 'pic:nvPicPr');
      child(xml, properties, PICTURE, 'pic:cNvPr', {
        id: String(id),
        name: image?.getAttribute('o:title') || 'Imagen',
      });
      child(xml, properties, PICTURE, 'pic:cNvPicPr');
      const fill = child(xml, picture, PICTURE, 'pic:blipFill');
      child(xml, fill, DRAWING, 'a:blip').setAttributeNS(RELATIONSHIP, 'r:embed', relation!);
      child(xml, child(xml, fill, DRAWING, 'a:stretch'), DRAWING, 'a:fillRect');
      shapeProperties(xml, picture, PICTURE, 'pic:spPr', width!, height!);
    }
    pict.replaceWith(drawing);
    converted++;
  }
  if (!converted) return bytes;
  for (const [prefix, uri] of Object.entries({
    w: WORD,
    wp: WP,
    a: DRAWING,
    pic: PICTURE,
    wps: SHAPE,
    r: RELATIONSHIP,
  }))
    xml.documentElement.setAttributeNS(XMLNS, `xmlns:${prefix}`, uri);
  entries['word/document.xml'] = strToU8(new XMLSerializer().serializeToString(xml));
  return zipSync(entries);
}

/**
 * Proyecta cuadros DrawingML rectangulares a VML únicamente para el visor de impresión.
 * @param bytes Exportación DOCX actual; el documento editable no se modifica.
 * @returns Copia de impresión con texto flotante que docx-preview puede representar.
 */
export function prepareWordForPrint(bytes: Uint8Array): Uint8Array {
  const entries = unzipSync(bytes);
  if (!entries['word/document.xml']) return bytes;
  const xml = new DOMParser().parseFromString(
    strFromU8(entries['word/document.xml']),
    'application/xml',
  );
  let converted = false;
  for (const drawing of Array.from(xml.getElementsByTagNameNS(WORD, 'drawing'))) {
    const shape = drawing.getElementsByTagNameNS(SHAPE, 'wsp')[0];
    const content = shape?.getElementsByTagNameNS(WORD, 'txbxContent')[0];
    const anchor = drawing.getElementsByTagNameNS(WP, 'anchor')[0];
    const extent = anchor?.getElementsByTagNameNS(WP, 'extent')[0];
    const horizontal = anchor?.getElementsByTagNameNS(WP, 'positionH')[0];
    const vertical = anchor?.getElementsByTagNameNS(WP, 'positionV')[0];
    if (!shape || !content || !anchor || !extent || !horizontal || !vertical) continue;
    const left = Number(horizontal.getElementsByTagNameNS(WP, 'posOffset')[0]?.textContent);
    const top = Number(vertical.getElementsByTagNameNS(WP, 'posOffset')[0]?.textContent);
    const width = Number(extent.getAttribute('cx'));
    const height = Number(extent.getAttribute('cy'));
    if (![left, top, width, height].every(Number.isFinite) || width <= 0 || height <= 0) continue;
    const pict = xml.createElementNS(WORD, 'w:pict');
    const rect = child(xml, pict, VML, 'v:rect', {
      filled: 'f',
      stroked: 'f',
      style: `position:absolute;margin-left:${left / EMUS_PER_POINT}pt;margin-top:${top / EMUS_PER_POINT}pt;width:${width / EMUS_PER_POINT}pt;height:${height / EMUS_PER_POINT}pt`,
    });
    const body = shape.getElementsByTagNameNS(SHAPE, 'bodyPr')[0];
    const insets = ['lIns', 'tIns', 'rIns', 'bIns'].map(
      (name) => `${Number(body?.getAttribute(name) || 0) / EMUS_PER_POINT}pt`,
    );
    const textbox = child(xml, rect, VML, 'v:textbox', { inset: insets.join(',') });
    textbox.append(content.cloneNode(true));
    drawing.replaceWith(pict);
    converted = true;
  }
  if (!converted) return bytes;
  entries['word/document.xml'] = strToU8(new XMLSerializer().serializeToString(xml));
  return zipSync(entries);
}

/** Recibe una longitud VML en puntos o cero y devuelve EMU enteros; devuelve null si no se puede preservar. */
function lengthInEmus(value: string | undefined): number | null {
  if (value === undefined || !/^-?(?:\d+(?:\.\d*)?|\.\d+)(?:pt)?$/.test(value.trim())) return null;
  const result = Math.round(Number.parseFloat(value) * EMUS_PER_POINT);
  return Number.isSafeInteger(result) ? result : null;
}

/** Crea un hijo XML con namespace y atributos; recibe padre/nombre y retorna el elemento insertado. */
function child(
  xml: XMLDocument,
  parent: Element,
  namespace: string,
  name: string,
  attributes: Record<string, string> = {},
): Element {
  const element = xml.createElementNS(namespace, name);
  for (const [key, value] of Object.entries(attributes)) element.setAttribute(key, value);
  parent.append(element);
  return element;
}

/** Añade geometría rectangular transparente; recibe tamaño EMU y no retorna valor. */
function shapeProperties(
  xml: XMLDocument,
  parent: Element,
  namespace: string,
  name: string,
  width: number,
  height: number,
): void {
  const properties = child(xml, parent, namespace, name);
  const transform = child(xml, properties, DRAWING, 'a:xfrm');
  child(xml, transform, DRAWING, 'a:off', { x: '0', y: '0' });
  child(xml, transform, DRAWING, 'a:ext', { cx: String(width), cy: String(height) });
  child(xml, child(xml, properties, DRAWING, 'a:prstGeom', { prst: 'rect' }), DRAWING, 'a:avLst');
  child(xml, properties, DRAWING, 'a:noFill');
  child(xml, child(xml, properties, DRAWING, 'a:ln'), DRAWING, 'a:noFill');
}
