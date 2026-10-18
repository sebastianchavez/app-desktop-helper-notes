import * as XLSX from 'xlsx';
import { DOMParser, XMLSerializer, Element as XmlElement } from '@xmldom/xmldom';
import { unzipSync, zipSync, strFromU8, strToU8 } from 'fflate';
import path from 'node:path';
import { randomUUID } from 'node:crypto';

export type ChartType = 'column' | 'bar' | 'line' | 'area' | 'pie' | 'doughnut' | 'scatter';
export type ChartDefinition = {
  id?: string;
  hostSheet: string;
  sourceSheet: string;
  sourceRange: string;
  title: string;
  type: ChartType;
  headers: boolean;
  category: number;
  series: number[];
  legend: boolean;
  labels: boolean;
  stacked: boolean;
  xTitle: string;
  yTitle: string;
  colors: string[];
  anchor: {
    row: number;
    col: number;
    offsetX: number;
    offsetY: number;
    width: number;
    height: number;
  };
};
export type ChartData = {
  categories: string[];
  x: (number | null)[];
  series: { name: string; values: (number | null)[]; color: string }[];
};
export type SavedChart = ChartDefinition & {
  id: string;
  chartPath: string;
  drawingPath: string;
  relationshipId: string;
  managed?: boolean;
  data: ChartData;
  error?: string;
};
const NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const C = 'http://schemas.openxmlformats.org/drawingml/2006/chart';
const A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const XDR = 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing';
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const PKG = 'http://schemas.openxmlformats.org/package/2006/relationships';
const TYPES = 'http://schemas.openxmlformats.org/package/2006/content-types';
const META = 'helperNotes/charts.json';
const PALETTE = [
  '#2563eb',
  '#16a34a',
  '#ea580c',
  '#9333ea',
  '#0891b2',
  '#dc2626',
  '#ca8a04',
  '#db2777',
];
const KINDS = new Set<ChartType>(['column', 'bar', 'line', 'area', 'pie', 'doughnut', 'scatter']);

/** Recibe un texto; retorna XML escapado, eliminando caracteres de control inválidos. */
function xml(value: unknown): string {
  return String(value ?? '')
    .replace(/[\u0000-\u0008\u000b\u000c\u000e-\u001f]/g, '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}
/** Recibe un nodo y nombre local opcional; retorna sus hijos XML directos. */
function children(node: XmlElement | undefined, name?: string): XmlElement[] {
  return Array.from(node?.childNodes ?? [])
    .filter((n) => n.nodeType === 1)
    .map((n) => n as XmlElement)
    .filter((n) => !name || n.localName === name);
}
/** Recibe Base64 XLSX o formato antiguo; retorna su ZIP nativo, normalizando XLS/CSV una sola vez. */
function archiveOf(base64: string): Record<string, Uint8Array> {
  let b = Buffer.from(base64, 'base64');
  if (b[0] !== 80 || b[1] !== 75)
    b = Buffer.from(
      XLSX.write(XLSX.read(b, { type: 'buffer' }), { type: 'buffer', bookType: 'xlsx' }),
    );
  return unzipSync(b);
}
/** Recibe ZIP/ruta y XML inicial opcional; retorna un documento editable en memoria. */
function parse(archive: Record<string, Uint8Array>, name: string, fallback = '') {
  return new DOMParser().parseFromString(
    archive[name] ? strFromU8(archive[name]) : fallback,
    'application/xml',
  );
}
/** Recibe ZIP/ruta/documento; serializa sólo esa parte y no retorna valor. */
function store(
  archive: Record<string, Uint8Array>,
  name: string,
  doc: ReturnType<typeof parse>,
): void {
  archive[name] = strToU8(new XMLSerializer().serializeToString(doc));
}
/** Recibe una parte OPC; retorna la ruta de sus relaciones. */
function relPath(part: string): string {
  return path.posix.join(path.posix.dirname(part), '_rels', path.posix.basename(part) + '.rels');
}
/** Recibe parte propietaria y destino relativo; retorna la ruta dentro del XLSX. */
function resolve(part: string, target: string): string {
  return target.startsWith('/')
    ? target.slice(1)
    : path.posix.normalize(path.posix.join(path.posix.dirname(part), target));
}
/** Recibe documento de relaciones, tipo y destino; agrega una relación con ID único y retorna ese ID. */
function addRelation(doc: ReturnType<typeof parse>, type: string, target: string): string {
  let n = 1;
  while (children(doc.documentElement!).some((r) => r.getAttribute('Id') === `rId${n}`)) n++;
  const r = doc.createElementNS(PKG, 'Relationship');
  r.setAttribute('Id', `rId${n}`);
  r.setAttribute('Type', `${REL}/${type}`);
  r.setAttribute('Target', target);
  doc.documentElement!.appendChild(r);
  return `rId${n}`;
}
/** Recibe XML OPC de tipos, parte y tipo MIME; agrega el registro si falta y no retorna valor. */
function addType(doc: ReturnType<typeof parse>, part: string, type: string): void {
  if (children(doc.documentElement!).some((n) => n.getAttribute('PartName') === `/${part}`)) return;
  const n = doc.createElementNS(TYPES, 'Override');
  n.setAttribute('PartName', `/${part}`);
  n.setAttribute('ContentType', type);
  doc.documentElement!.appendChild(n);
}
/** Recibe ZIP y nombre de hoja; retorna la parte worksheet y rechaza una hoja inexistente. */
function sheetPath(archive: Record<string, Uint8Array>, name: string): string {
  const book = parse(archive, 'xl/workbook.xml');
  const sheet = Array.from(book.getElementsByTagNameNS(NS, 'sheet')).find(
    (s) => s.getAttribute('name') === name,
  );
  if (!sheet) throw Error('CHART_SOURCE');
  const relation = children(parse(archive, 'xl/_rels/workbook.xml.rels').documentElement!).find(
    (r) => r.getAttribute('Id') === sheet.getAttributeNS(REL, 'id'),
  );
  if (!relation) throw Error('CHART_SOURCE');
  return resolve('xl/workbook.xml', relation.getAttribute('Target')!);
}

/**
 * Valida la configuración y obtiene categorías/series sin interpretar texto como números.
 * @param base64 Libro completo con sus resultados calculados actuales.
 * @param definition Rango rectangular, columnas relativas y presentación del gráfico.
 * @returns Datos tipados para SVG y cachés nativas; rechaza rangos o series no compatibles.
 */
export function buildSpreadsheetChartData(base64: string, definition: ChartDefinition): ChartData {
  const d = definition;
  if (
    !d ||
    !KINDS.has(d.type) ||
    typeof d.title !== 'string' ||
    d.title.length > 160 ||
    typeof d.xTitle !== 'string' ||
    d.xTitle.length > 160 ||
    typeof d.yTitle !== 'string' ||
    d.yTitle.length > 160
  )
    throw Error('CHART_FIELDS');
  if (
    typeof d.sourceRange !== 'string' ||
    !/^\$?[A-Z]{1,3}\$?\d+:\$?[A-Z]{1,3}\$?\d+$/i.test(d.sourceRange)
  )
    throw Error('CHART_RANGE');
  const range = XLSX.utils.decode_range(d.sourceRange.replace(/\$/g, '').toUpperCase()),
    count = range.e.r - range.s.r + (d.headers ? 0 : 1),
    width = range.e.c - range.s.c + 1;
  if (
    range.s.r < 0 ||
    range.s.c < 0 ||
    range.e.r >= 1048576 ||
    range.e.c >= 16384 ||
    count < 1 ||
    count > 3000 ||
    width < 2 ||
    width > 256
  )
    throw Error('CHART_RANGE');
  if (
    !Number.isInteger(d.category) ||
    d.category < 0 ||
    d.category >= width ||
    !Array.isArray(d.series) ||
    !d.series.length ||
    d.series.length > 12 ||
    new Set(d.series).size !== d.series.length ||
    d.series.some((s) => !Number.isInteger(s) || s < 0 || s >= width || s === d.category) ||
    count * d.series.length > 12000
  )
    throw Error('CHART_FIELDS');
  if ((d.type === 'pie' || d.type === 'doughnut') && d.series.length !== 1)
    throw Error('CHART_PIE_SERIES');
  if (
    !Array.isArray(d.colors) ||
    d.colors.length > 64 ||
    d.colors.some((c) => !/^#[\da-f]{6}$/i.test(c))
  )
    throw Error('CHART_FIELDS');
  const a = d.anchor;
  if (
    !a ||
    ![a.row, a.col].every(Number.isInteger) ||
    a.row < 0 ||
    a.row >= 1048576 ||
    a.col < 0 ||
    a.col >= 16384 ||
    ![a.offsetX, a.offsetY, a.width, a.height].every(Number.isFinite) ||
    a.offsetX < 0 ||
    a.offsetY < 0 ||
    a.width < 280 ||
    a.width > 1600 ||
    a.height < 200 ||
    a.height > 1000
  )
    throw Error('CHART_POSITION');
  const book = XLSX.read(Buffer.from(base64, 'base64'), { type: 'buffer', cellStyles: true });
  const sheet = book.Sheets[d.sourceSheet];
  if (!sheet || !book.Sheets[d.hostSheet]) throw Error('CHART_SOURCE');
  const first = range.s.r + (d.headers ? 1 : 0),
    categories: string[] = [],
    x: (number | null)[] = [];
  const series = d.series.map((column, i) => ({
    name: d.headers
      ? String(
          sheet[XLSX.utils.encode_cell({ r: range.s.r, c: range.s.c + column })]?.v ??
            `Serie ${i + 1}`,
        )
      : `Serie ${i + 1}`,
    values: [] as (number | null)[],
    color: d.colors[i] ?? PALETTE[i % PALETTE.length],
  }));
  for (let row = first; row <= range.e.r; row++) {
    const category = sheet[XLSX.utils.encode_cell({ r: row, c: range.s.c + d.category })];
    categories.push(category ? XLSX.utils.format_cell(category) : '');
    x.push(
      category?.t === 'n' && typeof category.v === 'number' && Number.isFinite(category.v)
        ? category.v
        : null,
    );
    d.series.forEach((column, i) => {
      const cell = sheet[XLSX.utils.encode_cell({ r: row, c: range.s.c + column })],
        v = cell?.v;
      series[i].values.push(
        cell?.t === 'n' && typeof v === 'number' && Number.isFinite(v) ? v : null,
      );
    });
  }
  if (!series.some((s) => s.values.some((v) => v !== null))) throw Error('CHART_NUMBERS');
  if (
    d.type === 'scatter' &&
    x.some((v, i) => v === null && series.some((s) => s.values[i] !== null))
  )
    throw Error('CHART_SCATTER');
  if (
    (d.type === 'pie' || d.type === 'doughnut') &&
    (series[0].values.some((v) => v !== null && v < 0) ||
      !series[0].values.some((v) => v !== null && v > 0))
  )
    throw Error('CHART_PIE_VALUES');
  return { categories, x, series };
}

/** Recibe referencia de hoja escapada y rango; retorna fórmula de gráfico absoluta, sin ejecutar fórmulas del usuario. */
function formula(sheet: string, col: number, start: number, end: number): string {
  return `'${sheet.replace(/'/g, "''")}'!$${XLSX.utils.encode_col(col)}$${start + 1}:$${XLSX.utils.encode_col(col)}$${end + 1}`;
}
/** Recibe valores de una serie; retorna la caché numérica con huecos y cantidad total conservados. */
function numberCache(values: (number | null)[]): string {
  return `<c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="${values.length}"/>${values.map((v, i) => (v === null ? '' : `<c:pt idx="${i}"><c:v>${v}</c:v></c:pt>`)).join('')}</c:numCache>`;
}
/** Recibe título visible; retorna el elemento c:title con texto escapado. */
function titleXml(title: string): string {
  return `<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="es-CL"/><a:t>${xml(title)}</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title>`;
}
/** Recibe configuración y datos; retorna un chartSpace nativo vinculado a rangos, con cachés y ejes. */
function chartXml(d: ChartDefinition, data: ChartData): string {
  const range = XLSX.utils.decode_range(d.sourceRange.replace(/\$/g, '').toUpperCase()),
    first = range.s.r + (d.headers ? 1 : 0),
    pie = d.type === 'pie' || d.type === 'doughnut';
  const series = data.series
    .map((s, i) => {
      const tx = d.headers
        ? `<c:tx><c:strRef><c:f>${xml(formula(d.sourceSheet, range.s.c + d.series[i], range.s.r, range.s.r).split(':')[0])}</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>${xml(s.name)}</c:v></c:pt></c:strCache></c:strRef></c:tx>`
        : `<c:tx><c:v>${xml(s.name)}</c:v></c:tx>`;
      const color = s.color.slice(1),
        sp = `<c:spPr><a:solidFill><a:srgbClr val="${color}"/></a:solidFill><a:ln w="25400"><a:solidFill><a:srgbClr val="${color}"/></a:solidFill></a:ln></c:spPr>`;
      const cats = `<c:cat><c:strRef><c:f>${xml(formula(d.sourceSheet, range.s.c + d.category, first, range.e.r))}</c:f><c:strCache><c:ptCount val="${data.categories.length}"/>${data.categories.map((v, n) => `<c:pt idx="${n}"><c:v>${xml(v)}</c:v></c:pt>`).join('')}</c:strCache></c:strRef></c:cat>`;
      const values = `<c:numRef><c:f>${xml(formula(d.sourceSheet, range.s.c + d.series[i], first, range.e.r))}</c:f>${numberCache(s.values)}</c:numRef>`;
      const points = pie
        ? s.values
            .map(
              (_, n) =>
                `<c:dPt><c:idx val="${n}"/><c:spPr><a:solidFill><a:srgbClr val="${(d.colors[n] ?? PALETTE[n % PALETTE.length]).slice(1)}"/></a:solidFill></c:spPr></c:dPt>`,
            )
            .join('')
        : '';
      return `<c:ser><c:idx val="${i}"/><c:order val="${i}"/>${tx}${sp}${d.type === 'line' || d.type === 'scatter' ? `<c:marker><c:symbol val="circle"/><c:size val="5"/></c:marker>` : ''}${points}${d.type === 'scatter' ? `<c:xVal><c:numRef><c:f>${xml(formula(d.sourceSheet, range.s.c + d.category, first, range.e.r))}</c:f>${numberCache(data.x)}</c:numRef></c:xVal><c:yVal>${values}</c:yVal>` : `${cats}<c:val>${values}</c:val>`}</c:ser>`;
    })
    .join('');
  const labels = d.labels
    ? `<c:dLbls><c:showLegendKey val="0"/><c:showVal val="${pie ? 0 : 1}"/><c:showCatName val="0"/><c:showSerName val="0"/><c:showPercent val="${pie ? 1 : 0}"/><c:showLeaderLines val="1"/></c:dLbls>`
    : '';
  const grouping = d.stacked
    ? 'stacked'
    : d.type === 'column' || d.type === 'bar'
      ? 'clustered'
      : 'standard';
  let plot: string;
  if (d.type === 'column' || d.type === 'bar')
    plot = `<c:barChart><c:barDir val="${d.type === 'bar' ? 'bar' : 'col'}"/><c:grouping val="${grouping}"/><c:varyColors val="0"/>${series}${labels}<c:gapWidth val="150"/><c:overlap val="${d.stacked ? 100 : 0}"/><c:axId val="100"/><c:axId val="200"/></c:barChart>`;
  else if (pie)
    plot = `<c:${d.type === 'pie' ? 'pieChart' : 'doughnutChart'}><c:varyColors val="1"/>${series}${labels}<c:firstSliceAng val="270"/>${d.type === 'doughnut' ? '<c:holeSize val="55"/>' : ''}</c:${d.type === 'pie' ? 'pieChart' : 'doughnutChart'}>`;
  else if (d.type === 'scatter')
    plot = `<c:scatterChart><c:scatterStyle val="marker"/><c:varyColors val="0"/>${series}${labels}<c:axId val="100"/><c:axId val="200"/></c:scatterChart>`;
  else
    plot = `<c:${d.type}Chart><c:grouping val="${grouping}"/><c:varyColors val="0"/>${series}${labels}<c:axId val="100"/><c:axId val="200"/></c:${d.type}Chart>`;
  const axis = (id: number, other: number, value: boolean, pos: string, title: string) =>
    `<c:${value ? 'valAx' : 'catAx'}><c:axId val="${id}"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="${pos}"/>${value ? '<c:majorGridlines/>' : ''}${title ? titleXml(title) : ''}<c:numFmt formatCode="General" sourceLinked="1"/><c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:crossAx val="${other}"/><c:crosses val="autoZero"/>${value ? '<c:crossBetween val="between"/>' : '<c:auto val="1"/><c:lblAlgn val="ctr"/><c:lblOffset val="100"/>'}</c:${value ? 'valAx' : 'catAx'}>`;
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><c:chartSpace xmlns:c="${C}" xmlns:a="${A}" xmlns:r="${REL}"><c:date1904 val="0"/><c:lang val="es-CL"/><c:roundedCorners val="0"/><c:chart>${d.title ? titleXml(d.title) : '<c:autoTitleDeleted val="1"/>'}<c:plotArea><c:layout/>${plot}${pie ? '' : axis(100, 200, d.type === 'scatter', d.type === 'bar' ? 'l' : 'b', d.xTitle) + axis(200, 100, true, d.type === 'bar' ? 'b' : 'l', d.yTitle)}</c:plotArea>${d.legend ? '<c:legend><c:legendPos val="b"/><c:layout/><c:overlay val="0"/></c:legend>' : ''}<c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/><c:showDLblsOverMax val="0"/></c:chart><c:spPr><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill><a:ln><a:noFill/></a:ln></c:spPr></c:chartSpace>`;
}

/** Recibe fórmula de rango de gráfico; retorna hoja y rango sin referencias externas o expresiones calculadas. */
function reference(text: string): { sheet: string; range: XLSX.Range } | null {
  const match = text.match(
    /^(?:'((?:[^']|'')+)'|([^!]+))!(\$?[A-Z]{1,3}\$?\d+)(?::(\$?[A-Z]{1,3}\$?\d+))?$/i,
  );
  if (!match) return null;
  try {
    return {
      sheet: match[1]?.replace(/''/g, "'") ?? match[2],
      range: XLSX.utils.decode_range(
        (match[3] + (match[4] ? ':' + match[4] : '')).replace(/\$/g, ''),
      ),
    };
  } catch {
    return null;
  }
}
/** Recibe un nodo y nombre de descendiente; retorna su primer texto o una cadena vacía. */
function text(node: XmlElement | undefined, name: string): string {
  return node
    ? (Array.from(node.getElementsByTagName('*')).find((n) => n.localName === name)?.textContent ??
        '')
    : '';
}

/**
 * Lee gráficos básicos nativos, incluidos los importados, sin modificar gráficos o dibujos no soportados.
 * @param base64 Libro completo.
 * @returns Definiciones, anclajes y datos actuales para renderizar; los orígenes inválidos publican error y una caché vacía.
 */
export function readSpreadsheetCharts(base64: string): SavedChart[] {
  const archive = archiveOf(base64),
    book = XLSX.read(Buffer.from(base64, 'base64'), { type: 'buffer', cellStyles: true });
  let saved: SavedChart[] = [];
  try {
    saved = archive[META] ? JSON.parse(strFromU8(archive[META])) : [];
    if (!Array.isArray(saved)) saved = [];
  } catch {
    saved = [];
  }
  const results: SavedChart[] = [];
  for (const hostSheet of book.SheetNames) {
    const owner = sheetPath(archive, hostSheet),
      sheet = parse(archive, owner).documentElement!,
      drawingNode = children(sheet, 'drawing')[0];
    if (!drawingNode || !archive[relPath(owner)]) continue;
    const relation = children(parse(archive, relPath(owner)).documentElement!).find(
      (r) => r.getAttribute('Id') === drawingNode.getAttributeNS(REL, 'id'),
    );
    if (!relation) continue;
    const drawingPath = resolve(owner, relation.getAttribute('Target')!);
    if (!archive[drawingPath] || !archive[relPath(drawingPath)]) continue;
    const drawing = parse(archive, drawingPath),
      rels = children(parse(archive, relPath(drawingPath)).documentElement!);
    for (const anchor of children(drawing.documentElement!)) {
      const chartNode = anchor.getElementsByTagNameNS(C, 'chart')[0];
      if (!chartNode) continue;
      const relationshipId = chartNode.getAttributeNS(REL, 'id')!,
        rel = rels.find((r) => r.getAttribute('Id') === relationshipId);
      if (!rel) continue;
      const chartPath = resolve(drawingPath, rel.getAttribute('Target')!);
      if (!archive[chartPath]) continue;
      const root = parse(archive, chartPath).documentElement!,
        plot = root.getElementsByTagNameNS(C, 'plotArea')[0];
      if (!plot) continue;
      const types = children(plot).filter((n) =>
        [
          'barChart',
          'lineChart',
          'areaChart',
          'pieChart',
          'doughnutChart',
          'scatterChart',
        ].includes(n.localName!),
      );
      if (
        types.length !== 1 ||
        children(plot).some((n) => n.localName?.endsWith('Chart') && !types.includes(n))
      )
        continue;
      const chart = types[0],
        seriesNodes = children(chart, 'ser');
      if (!seriesNodes.length) continue;
      // Los porcentajes y ejes secundarios necesitan una geometría distinta; se conserva su XML sin aproximarlos.
      if (children(chart, 'grouping')[0]?.getAttribute('val') === 'percentStacked') continue;
      const piePlot = chart.localName === 'pieChart' || chart.localName === 'doughnutChart';
      if (
        !piePlot &&
        children(plot).filter((n) => ['catAx', 'dateAx', 'valAx'].includes(n.localName!)).length !==
          2
      )
        continue;
      const categoryRef = reference(
        text(children(seriesNodes[0], chart.localName === 'scatterChart' ? 'xVal' : 'cat')[0], 'f'),
      );
      const refs = seriesNodes.map((s) =>
        reference(text(children(s, chart.localName === 'scatterChart' ? 'yVal' : 'val')[0], 'f')),
      );
      if (
        !categoryRef ||
        refs.some(
          (r) =>
            !r ||
            r.sheet !== categoryRef.sheet ||
            r.range.s.r !== categoryRef.range.s.r ||
            r.range.e.r !== categoryRef.range.e.r ||
            r.range.s.c !== r.range.e.c,
        )
      )
        continue;
      if (categoryRef.range.s.c !== categoryRef.range.e.c) continue;
      if (
        seriesNodes.some((s) => {
          const r = reference(
            text(children(s, chart.localName === 'scatterChart' ? 'xVal' : 'cat')[0], 'f'),
          );
          return (
            !r ||
            r.sheet !== categoryRef.sheet ||
            XLSX.utils.encode_range(r.range) !== XLSX.utils.encode_range(categoryRef.range)
          );
        })
      )
        continue;
      const headerRefs = seriesNodes.map((s) => reference(text(children(s, 'tx')[0], 'f')));
      const headers = headerRefs.every(
        (r) => r && r.sheet === categoryRef.sheet && r.range.s.r === categoryRef.range.s.r - 1,
      );
      let startCol = Math.min(categoryRef.range.s.c, ...refs.map((r) => r!.range.s.c)),
        endCol = Math.max(categoryRef.range.e.c, ...refs.map((r) => r!.range.e.c));
      const prior = saved.find(
        (s) =>
          s.chartPath === chartPath &&
          s.drawingPath === drawingPath &&
          s.relationshipId === relationshipId,
      );
      if (prior?.sourceSheet === categoryRef.sheet) {
        const range = XLSX.utils.decode_range(prior.sourceRange.replace(/\$/g, '').toUpperCase());
        if (
          range.s.r === categoryRef.range.s.r - (headers ? 1 : 0) &&
          range.e.r === categoryRef.range.e.r &&
          range.s.c <= startCol &&
          range.e.c >= endCol
        ) {
          startCol = range.s.c;
          endCol = range.e.c;
        }
      }
      const from = children(anchor, 'from')[0],
        to = children(anchor, 'to')[0],
        ext = children(anchor, 'ext')[0],
        pos = children(anchor, 'pos')[0];
      const read = (n: XmlElement | undefined, key: string) =>
        Number(children(n, key)[0]?.textContent ?? 0);
      let col = read(from, 'col'),
        row = read(from, 'row'),
        offsetX = read(from, 'colOff') / 9525,
        offsetY = read(from, 'rowOff') / 9525;
      const ws = book.Sheets[hostSheet],
        colWidth = (i: number) => ws['!cols']?.[i]?.wpx ?? 80,
        rowHeight = (i: number) => ws['!rows']?.[i]?.hpx ?? 20;
      if (pos) {
        let x = Number(pos.getAttribute('x')) / 9525,
          y = Number(pos.getAttribute('y')) / 9525;
        while (col < 16383 && x >= colWidth(col)) x -= colWidth(col++);
        while (row < 1048575 && y >= rowHeight(row)) y -= rowHeight(row++);
        offsetX = x;
        offsetY = y;
      }
      const width = ext
        ? Number(ext.getAttribute('cx')) / 9525
        : to
          ? Array.from({ length: Math.max(0, read(to, 'col') - col) }, (_, i) =>
              colWidth(col + i),
            ).reduce((a, b) => a + b, 0) +
            read(to, 'colOff') / 9525 -
            offsetX
          : 600;
      const height = ext
        ? Number(ext.getAttribute('cy')) / 9525
        : to
          ? Array.from({ length: Math.max(0, read(to, 'row') - row) }, (_, i) =>
              rowHeight(row + i),
            ).reduce((a, b) => a + b, 0) +
            read(to, 'rowOff') / 9525 -
            offsetY
          : 360;
      const titleNode = children(root.getElementsByTagNameNS(C, 'chart')[0], 'title')[0];
      const title = titleNode
        ? Array.from(titleNode.getElementsByTagNameNS(A, 't'))
            .map((n) => n.textContent)
            .join('')
        : '';
      const axes = children(plot).filter((n) => n.localName === 'catAx' || n.localName === 'valAx');
      const axisTitle = (index: number) => {
        const t = children(axes[index], 'title')[0];
        return t
          ? Array.from(t.getElementsByTagNameNS(A, 't'))
              .map((n) => n.textContent)
              .join('')
          : '';
      };
      const type: ChartType =
        chart.localName === 'barChart'
          ? children(chart, 'barDir')[0]?.getAttribute('val') === 'bar'
            ? 'bar'
            : 'column'
          : (chart.localName!.replace('Chart', '') as ChartType);
      const colors =
        type === 'pie' || type === 'doughnut'
          ? children(seriesNodes[0], 'dPt')
              .map((n) => n.getElementsByTagNameNS(A, 'srgbClr')[0]?.getAttribute('val'))
              .map((v, i) => (v ? '#' + v : PALETTE[i % PALETTE.length]))
          : seriesNodes.map((s, i) => {
              const c = children(s, 'spPr')[0]
                ?.getElementsByTagNameNS(A, 'srgbClr')[0]
                ?.getAttribute('val');
              return c ? '#' + c : PALETTE[i % PALETTE.length];
            });
      const definition: ChartDefinition = {
        id: prior?.id ?? `native:${drawingPath}:${relationshipId}`,
        hostSheet,
        sourceSheet: categoryRef.sheet,
        sourceRange: XLSX.utils.encode_range({
          s: { r: categoryRef.range.s.r - (headers ? 1 : 0), c: startCol },
          e: { r: categoryRef.range.e.r, c: endCol },
        }),
        headers,
        category: categoryRef.range.s.c - startCol,
        series: refs.map((r) => r!.range.s.c - startCol),
        type,
        title,
        legend: !!root.getElementsByTagNameNS(C, 'legend')[0],
        labels:
          !!children(chart, 'dLbls')[0] &&
          Array.from(children(chart, 'dLbls')[0].getElementsByTagName('*')).some(
            (n) =>
              ['showVal', 'showPercent'].includes(n.localName!) && n.getAttribute('val') === '1',
          ),
        stacked: children(chart, 'grouping')[0]?.getAttribute('val') === 'stacked',
        xTitle: axisTitle(0),
        yTitle: axisTitle(1),
        colors: colors.slice(0, 64),
        anchor: {
          row,
          col,
          offsetX,
          offsetY,
          width: Math.max(280, Math.min(1600, width || 600)),
          height: Math.max(200, Math.min(1000, height || 360)),
        },
      };
      try {
        results.push({
          ...definition,
          id: definition.id!,
          managed: prior?.managed === true,
          chartPath,
          drawingPath,
          relationshipId,
          data: buildSpreadsheetChartData(base64, definition),
        });
      } catch (error) {
        results.push({
          ...definition,
          id: definition.id!,
          managed: prior?.managed === true,
          chartPath,
          drawingPath,
          relationshipId,
          data: { categories: [], x: [], series: [] },
          error: error instanceof Error ? error.message : 'CHART_SOURCE',
        });
      }
    }
  }
  return results;
}

/**
 * Crea o edita un gráfico nativo y su anclaje, conservando los dibujos, imágenes y relaciones restantes.
 * @param base64 Libro completo.
 * @param definition Configuración y posición; un id existente edita sólo ese objeto.
 * @returns XLSX completo y gráfico actualizado con datos actuales.
 */
export function saveSpreadsheetChart(
  base64: string,
  definition: ChartDefinition,
): { base64: string; chart: SavedChart } {
  const data = buildSpreadsheetChartData(base64, definition),
    archive = archiveOf(base64),
    charts = readSpreadsheetCharts(base64);
  const previous = definition.id ? charts.find((c) => c.id === definition.id) : undefined;
  if (definition.id && !previous) throw Error('CHART_NOT_FOUND');
  if (charts.length >= 50 && !previous) throw Error('CHART_LIMIT');
  if (previous && previous.hostSheet !== definition.hostSheet) throw Error('CHART_POSITION');
  const owner = sheetPath(archive, definition.hostSheet),
    sheet = parse(archive, owner),
    sheetRoot = sheet.documentElement!,
    sheetRels = parse(archive, relPath(owner), `<Relationships xmlns="${PKG}"/>`),
    types = parse(archive, '[Content_Types].xml');
  let drawingNode = children(sheetRoot, 'drawing')[0],
    drawingPath: string;
  if (drawingNode) {
    const rel = children(sheetRels.documentElement!).find(
      (r) => r.getAttribute('Id') === drawingNode.getAttributeNS(REL, 'id'),
    );
    if (!rel) throw Error('CHART_POSITION');
    drawingPath = resolve(owner, rel.getAttribute('Target')!);
  } else {
    let n = 1;
    while (archive[`xl/drawings/drawing${n}.xml`]) n++;
    drawingPath = `xl/drawings/drawing${n}.xml`;
    drawingNode = sheet.createElementNS(NS, 'drawing');
    drawingNode.setAttributeNS(
      REL,
      'r:id',
      addRelation(
        sheetRels,
        'drawing',
        path.posix.relative(path.posix.dirname(owner), drawingPath),
      ),
    );
    const after = children(sheetRoot).find((n) =>
      [
        'legacyDrawing',
        'legacyDrawingHF',
        'picture',
        'oleObjects',
        'controls',
        'webPublishItems',
        'tableParts',
        'pivotParts',
        'extLst',
      ].includes(n.localName!),
    );
    sheetRoot.insertBefore(drawingNode, after ?? null);
  }
  const drawing = parse(archive, drawingPath, `<xdr:wsDr xmlns:xdr="${XDR}" xmlns:a="${A}"/>`),
    rels = parse(archive, relPath(drawingPath), `<Relationships xmlns="${PKG}"/>`);
  let chartPath = previous?.chartPath;
  if (!chartPath) {
    let n = 1;
    while (archive[`xl/charts/chart${n}.xml`]) n++;
    chartPath = `xl/charts/chart${n}.xml`;
  }
  const relationshipId =
    previous?.relationshipId ??
    addRelation(rels, 'chart', path.posix.relative(path.posix.dirname(drawingPath), chartPath));
  const id = previous?.id ?? randomUUID();
  const old = children(drawing.documentElement!).find((n) =>
    Array.from(n.getElementsByTagNameNS(C, 'chart')).some(
      (c) => c.getAttributeNS(REL, 'id') === relationshipId,
    ),
  );
  const ids = Array.from(drawing.getElementsByTagNameNS(XDR, 'cNvPr')).map((n) =>
    Number(n.getAttribute('id')),
  );
  const a = definition.anchor;
  const fragment = new DOMParser().parseFromString(
    `<xdr:oneCellAnchor xmlns:xdr="${XDR}" xmlns:a="${A}" xmlns:c="${C}" xmlns:r="${REL}"><xdr:from><xdr:col>${a.col}</xdr:col><xdr:colOff>${Math.round(a.offsetX * 9525)}</xdr:colOff><xdr:row>${a.row}</xdr:row><xdr:rowOff>${Math.round(a.offsetY * 9525)}</xdr:rowOff></xdr:from><xdr:ext cx="${Math.round(a.width * 9525)}" cy="${Math.round(a.height * 9525)}"/><xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="${Math.max(0, ...ids) + 1}" name="HelperChart_${xml(id)}"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm><a:graphic><a:graphicData uri="${C}"><c:chart r:id="${relationshipId}"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:oneCellAnchor>`,
    'application/xml',
  ).documentElement!;
  const anchor = drawing.importNode(fragment, true);
  if (old) drawing.documentElement!.replaceChild(anchor, old);
  else drawing.documentElement!.appendChild(anchor);
  archive[chartPath] = strToU8(chartXml(definition, data));
  store(archive, owner, sheet);
  store(archive, relPath(owner), sheetRels);
  store(archive, drawingPath, drawing);
  store(archive, relPath(drawingPath), rels);
  addType(types, drawingPath, 'application/vnd.openxmlformats-officedocument.drawing+xml');
  addType(types, chartPath, 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml');
  addType(types, META, 'application/json');
  store(archive, '[Content_Types].xml', types);
  const chart: SavedChart = {
    ...definition,
    id,
    chartPath,
    drawingPath,
    relationshipId,
    data,
    managed: true,
  };
  archive[META] = strToU8(
    JSON.stringify([
      ...charts.filter((c) => c.id !== id && c.managed).map(({ data, ...c }) => c),
      (({ data, ...c }) => c)(chart),
    ]),
  );
  return { base64: Buffer.from(zipSync(archive)).toString('base64'), chart };
}

/** Recibe libro e ID; elimina sólo el anclaje y la relación de ese gráfico, conservando otros objetos y celdas. */
export function deleteSpreadsheetChart(base64: string, id: string): string {
  const charts = readSpreadsheetCharts(base64),
    chart = charts.find((c) => c.id === id);
  if (!chart) throw Error('CHART_NOT_FOUND');
  const archive = archiveOf(base64),
    drawing = parse(archive, chart.drawingPath),
    rels = parse(archive, relPath(chart.drawingPath));
  const anchor = children(drawing.documentElement!).find((n) =>
    Array.from(n.getElementsByTagNameNS(C, 'chart')).some(
      (c) => c.getAttributeNS(REL, 'id') === chart.relationshipId,
    ),
  );
  if (!anchor) throw Error('CHART_NOT_FOUND');
  drawing.documentElement!.removeChild(anchor);
  children(rels.documentElement!)
    .filter((r) => r.getAttribute('Id') === chart.relationshipId)
    .forEach((r) => rels.documentElement!.removeChild(r));
  store(archive, chart.drawingPath, drawing);
  store(archive, relPath(chart.drawingPath), rels);
  archive[META] = strToU8(
    JSON.stringify(charts.filter((c) => c.id !== id && c.managed).map(({ data, ...c }) => c)),
  );
  return Buffer.from(zipSync(archive)).toString('base64');
}

/** Recibe el libro después de editar celdas; actualiza cachés de gráficos soportados y conserva los inválidos con sus últimos datos. */
export function refreshSpreadsheetCharts(base64: string): { base64: string; warnings: string[] } {
  const warnings: string[] = [];
  for (const chart of readSpreadsheetCharts(base64)) {
    if (!chart.managed) continue;
    try {
      base64 = saveSpreadsheetChart(base64, chart).base64;
    } catch (error) {
      warnings.push(`${chart.title}: ${error instanceof Error ? error.message : 'CHART_SOURCE'}`);
    }
  }
  return { base64, warnings };
}
