/**
 * Recupera los cuadros de texto VML que el visor coloca dentro de formas SVG no contenedoras.
 * Chromium no dibuja un foreignObject hijo de rect/ellipse, aunque su texto exista en el DOM.
 * @param container Vista DOCX recién renderizada, antes de montarla en el documento.
 * @returns Cantidad de cuadros recuperados; conserva los estilos, las posiciones y los gráficos.
 */
export function restoreWordTextBoxes(container: HTMLElement): number {
  let restored = 0;
  const invalidBoxes = Array.from(container.querySelectorAll('svg foreignObject')).filter((box) =>
    ['rect', 'ellipse', 'line', 'path', 'image', 'polygon', 'polyline'].includes(
      box.parentElement?.localName ?? '',
    ),
  );
  const roots = new Set(invalidBoxes.map((box) => box.closest('svg')!));
  for (const svg of roots) {
    const boxes = invalidBoxes.filter((box) => box.closest('svg') === svg);
    const overlay = document.createElement('div');
    overlay.className = 'word-text-box';
    // Las medidas de VML están en puntos y sus márgenes se refieren a la misma página que el SVG.
    overlay.style.cssText = svg.getAttribute('style') ?? '';
    overlay.style.boxSizing = 'border-box';
    overlay.style.overflow = 'visible';
    for (const box of boxes) {
      const content = document.createElement('div');
      content.style.cssText = box.getAttribute('style') ?? '';
      content.append(...Array.from(box.childNodes));
      overlay.append(content);
      box.remove();
      restored++;
    }
    // Las imágenes y formas con relleno explícito siguen en SVG; el texto usa un contenedor HTML.
    const hasGraphics = svg.querySelector(
      'image, path, polygon, polyline, line, ellipse, rect[fill]',
    );
    if (hasGraphics) svg.after(overlay);
    else svg.replaceWith(overlay);
  }
  return restored;
}
