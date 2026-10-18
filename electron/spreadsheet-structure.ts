import { DOMParser, XMLSerializer, Element as XmlElement } from '@xmldom/xmldom';
import { unzipSync, strFromU8, strToU8 } from 'fflate';
import path from 'node:path';
const NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const PKG = 'http://schemas.openxmlformats.org/package/2006/relationships';
const DYNAMIC = 'http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray';
type Archive = Record<string, Uint8Array>;
export type NativeTable = {
  name: string;
  id?: number;
  columnIds?: number[];
  sheet: string;
  ref: string;
  columns: string[];
  headerRows: number;
  totalRows: number;
  sourceName?: string;
  totals?: string[];
  totalFormulas?: Record<string, string>;
  calculated?: Record<string, string>;
  columnSources?: Record<string, string>;
};
export type NativeSheetStructure = {
  tables: NativeTable[];
  arrays: Record<string, { ref: string; dynamic: boolean }>;
  errors: Record<string, string>;
  hiddenRows: number[];
  hasActiveFilters?: boolean;
};
/** Recibe nodo y nombre local opcional; retorna sus hijos XML directos. */
function children(node: XmlElement | undefined, name?: string): XmlElement[] {
  return Array.from(node?.childNodes ?? [])
    .filter((n) => n.nodeType === 1)
    .map((n) => n as XmlElement)
    .filter((n) => !name || n.localName === name);
}
/** Recibe paquete, ruta y XML inicial opcional; retorna documento XML independiente. */
function parse(a: Archive, p: string, initial?: string) {
  return new DOMParser().parseFromString(
    a[p] ? strFromU8(a[p]) : (initial ?? ''),
    'application/xml',
  );
}
/** Recibe documento y ruta; actualiza sólo esa parte del paquete, sin retornar valor. */
function store(a: Archive, p: string, doc: ReturnType<typeof parse>): void {
  a[p] = strToU8(new XMLSerializer().serializeToString(doc));
}
/** Recibe ruta de parte; retorna su archivo de relaciones. */
function relPath(p: string): string {
  return path.posix.join(path.posix.dirname(p), '_rels', path.posix.basename(p) + '.rels');
}
/** Recibe parte propietaria y destino; retorna ruta normalizada dentro del paquete. */
function target(owner: string, value: string): string {
  return value.startsWith('/')
    ? value.slice(1)
    : path.posix.normalize(path.posix.join(path.posix.dirname(owner), value));
}

/** Recibe XLSX Base64; retorna tablas, matrices y errores nativos por hoja, sin modificar archivos. Los formatos no ZIP retornan un mapa vacío. */
export function readSpreadsheetStructure(base64: string): Record<string, NativeSheetStructure> {
  const buffer = Buffer.from(base64, 'base64');
  if (buffer[0] !== 80 || buffer[1] !== 75) return {};
  const a = unzipSync(buffer),
    book = parse(a, 'xl/workbook.xml'),
    rels = children(parse(a, 'xl/_rels/workbook.xml.rels').documentElement!);
  const metadataRel = rels.find((r) => r.getAttribute('Type')?.endsWith('/sheetMetadata'));
  const dynamicIndices = new Set<number>();
  if (metadataRel) {
    const metadata = parse(a, target('xl/workbook.xml', metadataRel.getAttribute('Target')!)),
      types = children(children(metadata.documentElement!, 'metadataTypes')[0], 'metadataType'),
      typeIndex = types.findIndex((t) => t.getAttribute('name') === 'XLDAPR') + 1;
    const future = children(metadata.documentElement!, 'futureMetadata').find(
      (f) => f.getAttribute('name') === 'XLDAPR',
    );
    const dynamicRecords = new Set(
      children(future, 'bk').flatMap((bk, i) =>
        Array.from(bk.getElementsByTagNameNS(DYNAMIC, 'dynamicArrayProperties')).some(
          (n) => n.getAttribute('fDynamic') === '1',
        )
          ? [i]
          : [],
      ),
    );
    children(children(metadata.documentElement!, 'cellMetadata')[0], 'bk').forEach((bk, i) => {
      if (
        children(bk, 'rc').some(
          (rc) =>
            Number(rc.getAttribute('t')) === typeIndex &&
            dynamicRecords.has(Number(rc.getAttribute('v'))),
        )
      )
        dynamicIndices.add(i + 1);
    });
  }
  const result: Record<string, NativeSheetStructure> = {};
  for (const sheet of Array.from(book.getElementsByTagNameNS(NS, 'sheet'))) {
    const relation = rels.find((r) => r.getAttribute('Id') === sheet.getAttributeNS(REL, 'id'));
    if (!relation) continue;
    const owner = target('xl/workbook.xml', relation.getAttribute('Target')!),
      doc = parse(a, owner),
      name = sheet.getAttribute('name')!,
      out: NativeSheetStructure = {
        tables: [],
        arrays: {},
        errors: {},
        hiddenRows: Array.from(doc.getElementsByTagNameNS(NS, 'row'))
          .filter((r) => r.getAttribute('hidden') === '1')
          .map((r) => Number(r.getAttribute('r')) - 1),
      };
    out.hasActiveFilters = doc.getElementsByTagNameNS(NS, 'filterColumn').length > 0;
    result[name] = out;
    for (const cell of Array.from(doc.getElementsByTagNameNS(NS, 'c'))) {
      const address = cell.getAttribute('r')!,
        f = children(cell, 'f')[0];
      if (cell.getAttribute('t') === 'e')
        out.errors[address] = children(cell, 'v')[0]?.textContent ?? '#VALUE!';
      if (f?.getAttribute('t') === 'array' && f.getAttribute('ref'))
        out.arrays[address] = {
          ref: f.getAttribute('ref')!,
          dynamic: dynamicIndices.has(Number(cell.getAttribute('cm'))),
        };
    }
    if (!a[relPath(owner)]) continue;
    const sheetRels = children(parse(a, relPath(owner)).documentElement!);
    for (const part of Array.from(doc.getElementsByTagNameNS(NS, 'tablePart'))) {
      const r = sheetRels.find((r) => r.getAttribute('Id') === part.getAttributeNS(REL, 'id'));
      if (!r) continue;
      const table = parse(a, target(owner, r.getAttribute('Target')!)).documentElement!;
      if (table.getElementsByTagNameNS(NS, 'filterColumn').length) out.hasActiveFilters = true;
      out.tables.push({
        id: Number(table.getAttribute('id')),
        columnIds: Array.from(table.getElementsByTagNameNS(NS, 'tableColumn')).map((c) =>
          Number(c.getAttribute('id')),
        ),
        name: table.getAttribute('displayName') ?? table.getAttribute('name')!,
        sheet: name,
        ref: table.getAttribute('ref')!,
        headerRows: Number(table.getAttribute('headerRowCount') ?? 1),
        totalRows: Number(table.getAttribute('totalsRowCount') ?? 0),
        ...(Number(table.getAttribute('totalsRowCount') ?? 0)
          ? {
              totals: Array.from(table.getElementsByTagNameNS(NS, 'tableColumn')).map((c) =>
                c.getAttribute('totalsRowFunction') === 'none'
                  ? ''
                  : (c.getAttribute('totalsRowFunction') ?? ''),
              ),
            }
          : {}),
        columns: Array.from(table.getElementsByTagNameNS(NS, 'tableColumn')).map(
          (c) => c.getAttribute('name') ?? '',
        ),
      });
    }
  }
  return result;
}

/**
 * Sincroniza las tablas finales de una hoja dentro del paquete, sin reconstruir estilos ajenos.
 * @param archive Paquete XLSX mutable; recibe tablas nuevas, renombradas, desplazadas o eliminadas.
 * @param owner Parte de la hoja propietaria.
 * @param tables Definiciones finales; sourceName vincula un cambio de nombre con su tabla nativa.
 * @returns Nada; mantiene IDs de columnas por nombre y crea relaciones/content types cuando corresponde.
 */
export function patchSpreadsheetTables(
  archive: Archive,
  owner: string,
  tables: NativeTable[],
): void {
  const doc = parse(archive, owner),
    root = doc.documentElement!;
  const relationships = parse(archive, relPath(owner), '<Relationships xmlns="' + PKG + '"/>');
  const types = parse(archive, '[Content_Types].xml');
  let parts = children(root, 'tableParts')[0];
  const existing = children(relationships.documentElement!).filter((r) =>
    r.getAttribute('Type')?.endsWith('/table'),
  );
  const handled = new Set<NativeTable>();
  const renames = new Map<string, string>();
  const changedNames = new Set<string>();
  const entries: { table: NativeTable; part: string; xml: ReturnType<typeof parse> }[] = [];
  for (const relation of existing) {
    const part = target(owner, relation.getAttribute('Target')!),
      xml = parse(archive, part),
      old =
        xml.documentElement!.getAttribute('displayName') ??
        xml.documentElement!.getAttribute('name')!;
    const table = tables.find((t) =>
      t.id
        ? t.id === Number(xml.documentElement!.getAttribute('id'))
        : t.name === old || t.sourceName === old,
    );
    if (table) {
      handled.add(table);
      entries.push({ table, part, xml });
      if (old !== table.name) renames.set(old.toLowerCase(), table.name);
    } else {
      changedNames.add(old.toLowerCase());
      relationships.documentElement!.removeChild(relation);
      children(parts, 'tablePart')
        .filter((p) => p.getAttributeNS(REL, 'id') === relation.getAttribute('Id'))
        .forEach((p) => parts!.removeChild(p));
      children(types.documentElement!, 'Override')
        .filter((p) => p.getAttribute('PartName') === '/' + part)
        .forEach((p) => types.documentElement!.removeChild(p));
      delete archive[part];
    }
  }
  const allParts = Object.keys(archive).filter((p) => /^xl\/tables\/[^/]+\.xml$/.test(p));
  let nextId =
    Math.max(
      0,
      ...allParts.map((p) => Number(parse(archive, p).documentElement!.getAttribute('id')) || 0),
    ) + 1;
  for (const table of tables.filter((t) => !handled.has(t))) {
    const tableId = table.id ?? nextId++;
    if (
      !Number.isInteger(tableId) ||
      tableId < 1 ||
      allParts.some(
        (p) => Number(parse(archive, p).documentElement!.getAttribute('id')) === tableId,
      )
    )
      throw Error('Identificador de tabla duplicado.');
    let n = tableId,
      part = 'xl/tables/table' + n + '.xml';
    while (archive[part] || entries.some((e) => e.part === part)) {
      n = nextId++;
      part = 'xl/tables/table' + n + '.xml';
    }
    const xml = parse(
      archive,
      part,
      '<table xmlns="' +
        NS +
        '" id="' +
        tableId +
        '"><autoFilter/><tableColumns/><tableStyleInfo name="TableStyleMedium2" showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/></table>',
    );
    entries.push({ table, part, xml });
    const ids = new Set(children(relationships.documentElement!).map((r) => r.getAttribute('Id')));
    let id = 1;
    while (ids.has('rId' + id)) id++;
    const relation = relationships.createElementNS(PKG, 'Relationship');
    relation.setAttribute('Id', 'rId' + id);
    relation.setAttribute('Type', REL + '/table');
    relation.setAttribute('Target', path.posix.relative(path.posix.dirname(owner), part));
    relationships.documentElement!.appendChild(relation);
    if (!parts) {
      parts = doc.createElementNS(NS, 'tableParts');
      root.insertBefore(parts, children(root, 'extLst')[0] ?? null);
    }
    const child = doc.createElementNS(NS, 'tablePart');
    child.setAttributeNS(REL, 'r:id', 'rId' + id);
    parts.appendChild(child);
    const content = types.createElementNS(types.documentElement!.namespaceURI!, 'Override');
    content.setAttribute('PartName', '/' + part);
    content.setAttribute(
      'ContentType',
      'application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml',
    );
    types.documentElement!.appendChild(content);
  }
  for (const { table, part, xml } of entries) {
    const element = xml.documentElement!,
      columns = children(element, 'tableColumns')[0],
      previous = children(columns, 'tableColumn');
    if (
      element.getAttribute('ref') === table.ref &&
      element.getAttribute('displayName') === table.name &&
      Number(element.getAttribute('totalsRowCount') ?? 0) === table.totalRows &&
      previous.map((c) => c.getAttribute('name')).join('\0') === table.columns.join('\0') &&
      (!table.totals ||
        previous.every(
          (c, i) => (c.getAttribute('totalsRowFunction') ?? '') === (table.totals![i] ?? ''),
        )) &&
      (!table.totalFormulas ||
        previous.every(
          (c) =>
            table.totalFormulas![c.getAttribute('name')!] === undefined ||
            table.totalFormulas![c.getAttribute('name')!]?.replace(/^=/, '') ===
              children(c, 'totalsRowFormula')[0]?.textContent,
        )) &&
      (!table.calculated ||
        previous.every(
          (c) =>
            table.calculated![c.getAttribute('name')!] === undefined ||
            table.calculated![c.getAttribute('name')!]?.replace(/^=/, '') ===
              children(c, 'calculatedColumnFormula')[0]?.textContent,
        ))
    )
      continue;
    const match = /^([A-Z]{1,3})([1-9]\d*):([A-Z]{1,3})([1-9]\d*)$/.exec(table.ref);
    /** Recibe etiqueta A1; retorna columna base uno para validar límites nativos. */
    const columnNumber = (label: string): number =>
      [...label].reduce((n, c) => n * 26 + c.charCodeAt(0) - 64, 0);
    if (
      !match ||
      !columns ||
      !table.columns.length ||
      !/^[\p{L}_\\][\p{L}\p{N}_.\\]*$/u.test(table.name) ||
      columnNumber(match[3]) > 16384 ||
      Number(match[4]) > 1048576 ||
      Number(match[4]) - Number(match[2]) < table.headerRows + table.totalRows ||
      columnNumber(match[3]) - columnNumber(match[1]) + 1 !== table.columns.length ||
      table.columns.some((c) => !c || c.length > 255) ||
      new Set(table.columns.map((c) => c.toLowerCase())).size !== table.columns.length
    )
      throw Error('Definición de tabla inválida.');
    changedNames.add(table.name.toLowerCase());
    element.setAttribute('name', table.name);
    element.setAttribute('displayName', table.name);
    element.setAttribute('ref', table.ref);
    element.setAttribute('headerRowCount', String(table.headerRows));
    element.setAttribute('totalsRowCount', String(table.totalRows));
    element.setAttribute('totalsRowShown', String(table.totalRows));
    const filter = children(element, 'autoFilter')[0];
    if (filter) {
      filter.setAttribute(
        'ref',
        match[1] + match[2] + ':' + match[3] + (Number(match[4]) - table.totalRows),
      );
      for (const field of children(filter, 'filterColumn')) {
        const oldName = previous[Number(field.getAttribute('colId'))]?.getAttribute('name');
        const index = table.columns.findIndex(
          (name) => name === oldName || table.columnSources?.[name] === oldName,
        );
        if (index < 0) filter.removeChild(field);
        else field.setAttribute('colId', String(index));
      }
      children(filter, 'sortState').forEach((c) => filter.removeChild(c));
    }
    children(element, 'sortState').forEach((c) => element.removeChild(c));
    if (
      table.columnIds &&
      (new Set(table.columnIds).size !== table.columnIds.length ||
        table.columnIds.some((id) => !Number.isInteger(id) || id < 1))
    )
      throw Error('Identificador de columna inválido.');
    let id =
      Math.max(
        0,
        ...(table.columnIds ?? []),
        ...previous.map((c) => Number(c.getAttribute('id'))),
      ) + 1;
    previous.forEach((c) => columns.removeChild(c));
    table.columns.forEach((name, i) => {
      const column =
        previous.find((c) =>
          table.columnIds?.[i]
            ? Number(c.getAttribute('id')) === table.columnIds[i]
            : c.getAttribute('name') === (table.columnSources?.[name] ?? name),
        ) ?? xml.createElementNS(NS, 'tableColumn');
      if (!column.hasAttribute('id'))
        column.setAttribute('id', String(table.columnIds?.[i] ?? id++));
      column.setAttribute('name', name);
      if (table.totals || !table.totalRows) {
        const fn = table.totalRows ? (table.totals?.[i] ?? '') : '';
        if (fn) column.setAttribute('totalsRowFunction', fn);
        else column.removeAttribute('totalsRowFunction');
        if (fn !== 'custom')
          children(column, 'totalsRowFormula').forEach((c) => column.removeChild(c));
        if (!table.totalRows) column.removeAttribute('totalsRowLabel');
      }
      if (table.totalFormulas?.[name]) {
        let formula = children(column, 'totalsRowFormula')[0];
        if (!formula) {
          formula = xml.createElementNS(NS, 'totalsRowFormula');
          column.appendChild(formula);
        }
        formula.textContent = table.totalFormulas[name].replace(/^=/, '');
      }
      if (table.totalRows && i === 0 && !table.totals?.[i])
        column.setAttribute('totalsRowLabel', 'Total');
      if (table.calculated?.[name] !== undefined) {
        let formula = children(column, 'calculatedColumnFormula')[0];
        if (!formula) {
          formula = xml.createElementNS(NS, 'calculatedColumnFormula');
          column.insertBefore(formula, column.firstChild);
        }
        formula.textContent = table.calculated[name].replace(/^=/, '');
      }
      columns.appendChild(column);
    });
    columns.setAttribute('count', String(table.columns.length));
    store(archive, part, xml);
  }
  if (parts) {
    if (!children(parts).length) root.removeChild(parts);
    else parts.setAttribute('count', String(children(parts).length));
  }
  store(archive, owner, doc);
  store(archive, relPath(owner), relationships);
  store(archive, '[Content_Types].xml', types);
  // Las cachés importadas por nombre se conservan; Excel puede refrescarlas contra el nombre/rango actualizado.
  if (changedNames.size || renames.size)
    for (const part of Object.keys(archive).filter((p) =>
      /^xl\/pivotCache\/pivotCacheDefinition[^/]*\.xml$/.test(p),
    )) {
      const xml = parse(archive, part);
      let changed = false;
      for (const source of Array.from(xml.getElementsByTagNameNS(NS, 'worksheetSource'))) {
        const old = source.getAttribute('name')?.toLowerCase();
        if (!old) continue;
        const next = renames.get(old);
        if (next) {
          source.setAttribute('name', next);
          changed = true;
        }
        if (changedNames.has(old) || (next && changedNames.has(next.toLowerCase()))) {
          xml.documentElement!.setAttribute('refreshOnLoad', '1');
          changed = true;
        }
      }
      if (changed) store(archive, part, xml);
    }
  // Las fórmulas de columnas calculadas son partes independientes de las celdas reconstruidas.
  if (renames.size)
    for (const part of Object.keys(archive).filter((p) => /^xl\/tables\/[^/]+\.xml$/.test(p))) {
      const xml = parse(archive, part);
      let changed = false;
      for (const tag of ['calculatedColumnFormula', 'totalsRowFormula'])
        for (const f of Array.from(xml.getElementsByTagNameNS(NS, tag))) {
          const old = f.textContent ?? '',
            value = old.replace(/"(?:[^"]|"")*"|[\p{L}_\\][\p{L}\p{N}_.\\]*(?=\[)/gu, (t) =>
              t.startsWith('"') ? t : (renames.get(t.toLowerCase()) ?? t),
            );
          if (value !== old) {
            f.textContent = value;
            changed = true;
          }
        }
      if (changed) store(archive, part, xml);
    }
}

/**
 * Agrega o reutiliza un registro dinámico XLDAPR sin reemplazar metadatos de imágenes u otras características.
 * @param archive Paquete XLSX mutable.
 * @returns Índice cm de base uno para asignar a las celdas ancla dinámicas.
 */
export function ensureDynamicSpreadsheetMetadata(archive: Archive): number {
  const owner = 'xl/workbook.xml',
    relationships = parse(archive, relPath(owner)),
    rels = children(relationships.documentElement!);
  let relation = rels.find((r) => r.getAttribute('Type')?.endsWith('/sheetMetadata'));
  let part: string;
  if (relation) part = target(owner, relation.getAttribute('Target')!);
  else {
    part = 'xl/metadata.xml';
    let n = 1;
    while (archive[part]) part = `xl/metadata${n++}.xml`;
    const ids = new Set(rels.map((r) => r.getAttribute('Id')));
    let id = 1;
    while (ids.has('rId' + id)) id++;
    relation = relationships.createElementNS(PKG, 'Relationship');
    relation.setAttribute('Id', 'rId' + id);
    relation.setAttribute('Type', REL + '/sheetMetadata');
    relation.setAttribute('Target', path.posix.relative('xl', part));
    relationships.documentElement!.appendChild(relation);
    store(archive, relPath(owner), relationships);
  }
  const doc = parse(archive, part, `<metadata xmlns="${NS}"/>`),
    root = doc.documentElement!;
  let types = children(root, 'metadataTypes')[0];
  if (!types) {
    types = doc.createElementNS(NS, 'metadataTypes');
    root.insertBefore(types, root.firstChild);
  }
  let typeIndex =
    children(types, 'metadataType').findIndex((t) => t.getAttribute('name') === 'XLDAPR') + 1;
  if (!typeIndex) {
    const type = doc.createElementNS(NS, 'metadataType');
    type.setAttribute('name', 'XLDAPR');
    type.setAttribute('minSupportedVersion', '120000');
    for (const name of [
      'copy',
      'pasteAll',
      'pasteValues',
      'merge',
      'splitFirst',
      'rowColShift',
      'clearFormats',
      'clearComments',
      'assign',
      'coerce',
      'cellMeta',
    ])
      type.setAttribute(name, '1');
    types.appendChild(type);
    typeIndex = children(types, 'metadataType').length;
    types.setAttribute('count', String(typeIndex));
  }
  let future = children(root, 'futureMetadata').find((f) => f.getAttribute('name') === 'XLDAPR');
  if (!future) {
    future = doc.createElementNS(NS, 'futureMetadata');
    future.setAttribute('name', 'XLDAPR');
    const after = children(root).find((n) =>
      ['cellMetadata', 'valueMetadata', 'extLst'].includes(n.localName!),
    );
    root.insertBefore(future, after ?? null);
  }
  let futureIndex = children(future, 'bk').findIndex((b) =>
    Array.from(b.getElementsByTagNameNS(DYNAMIC, 'dynamicArrayProperties')).some(
      (n) => n.getAttribute('fDynamic') === '1',
    ),
  );
  if (futureIndex < 0) {
    const fragment = new DOMParser().parseFromString(
      `<bk xmlns="${NS}"><extLst><ext uri="{bdbb8cdc-fa1e-496e-a857-3c3f30c029c3}"><xda:dynamicArrayProperties xmlns:xda="${DYNAMIC}" fDynamic="1" fCollapsed="0"/></ext></extLst></bk>`,
      'application/xml',
    ).documentElement!;
    future.appendChild(doc.importNode(fragment, true));
    futureIndex = children(future, 'bk').length - 1;
    future.setAttribute('count', String(futureIndex + 1));
  }
  let cells = children(root, 'cellMetadata')[0];
  if (!cells) {
    cells = doc.createElementNS(NS, 'cellMetadata');
    const after = children(root).find((n) => ['valueMetadata', 'extLst'].includes(n.localName!));
    root.insertBefore(cells, after ?? null);
  }
  let index = children(cells, 'bk').findIndex(
    (b) =>
      children(b, 'rc').length === 1 &&
      children(b, 'rc')[0].getAttribute('t') === String(typeIndex) &&
      children(b, 'rc')[0].getAttribute('v') === String(futureIndex),
  );
  if (index < 0) {
    const bk = doc.createElementNS(NS, 'bk'),
      rc = doc.createElementNS(NS, 'rc');
    rc.setAttribute('t', String(typeIndex));
    rc.setAttribute('v', String(futureIndex));
    bk.appendChild(rc);
    cells.appendChild(bk);
    index = children(cells, 'bk').length - 1;
    cells.setAttribute('count', String(index + 1));
  }
  store(archive, part, doc);
  const content = parse(archive, '[Content_Types].xml');
  if (!children(content.documentElement!).some((n) => n.getAttribute('PartName') === '/' + part)) {
    const type = content.createElementNS(
      'http://schemas.openxmlformats.org/package/2006/content-types',
      'Override',
    );
    type.setAttribute('PartName', '/' + part);
    type.setAttribute(
      'ContentType',
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml',
    );
    content.documentElement!.appendChild(type);
    store(archive, '[Content_Types].xml', content);
  }
  return index + 1;
}
