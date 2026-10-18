import { restoreWordTextBoxes } from './word-text-boxes';

describe('Word VML text boxes', () => {
  /**
   * Reproduce la estructura SVG inválida generada por el visor para un cuadro VML de Word.
   * @returns Contenedor con posición, fuente y color originales del texto de prueba.
   */
  function textBoxFixture(): HTMLDivElement {
    const container = document.createElement('div');
    container.innerHTML =
      '<svg style="position:absolute;margin-left:240pt;margin-top:33pt;width:360pt;height:40pt;z-index:2"><rect width="100%" height="100%"><foreignObject width="100%" height="100%"><p xmlns="http://www.w3.org/1999/xhtml"><span style="color:rgb(255,255,255);font-size:28pt">Visible heading</span></p></foreignObject></rect></svg>';
    return container;
  }

  it('restores text as positioned HTML while preserving font and color', () => {
    const container = textBoxFixture();
    expect(restoreWordTextBoxes(container)).toBe(1);
    const overlay = container.querySelector<HTMLElement>('.word-text-box')!;
    expect(overlay.style.position).toBe('absolute');
    expect(overlay.style.marginLeft).toBe('240pt');
    expect(overlay.style.marginTop).toBe('33pt');
    expect(overlay.style.width).toBe('360pt');
    expect(overlay.textContent).toBe('Visible heading');
    expect(overlay.querySelector('span')!.style.fontSize).toBe('28pt');
    expect(container.querySelector('foreignObject')).toBeNull();
    expect(container.querySelector('svg')).toBeNull();
  });

  it('keeps accompanying graphics and produces no duplicate text on repeated passes', () => {
    const container = textBoxFixture();
    const image = document.createElementNS('http://www.w3.org/2000/svg', 'image');
    image.setAttribute('href', 'data:image/png;base64,AA==');
    container.querySelector('svg')!.append(image);
    expect(restoreWordTextBoxes(container)).toBe(1);
    expect(container.querySelector('image')).toBe(image);
    expect(restoreWordTextBoxes(container)).toBe(0);
    expect(container.querySelectorAll('.word-text-box')).toHaveLength(1);
  });

  it('leaves ordinary paragraphs and tables untouched', () => {
    const container = document.createElement('div');
    container.innerHTML = '<p>Paragraph</p><table><tr><td>Cell</td></tr></table>';
    const original = container.innerHTML;
    expect(restoreWordTextBoxes(container)).toBe(0);
    expect(container.innerHTML).toBe(original);
  });

  it('leaves valid foreign objects inside SVG containers unchanged', () => {
    const container = document.createElement('div');
    container.innerHTML =
      '<svg><foreignObject width="100" height="40"><p>Valid text</p></foreignObject></svg>';
    const original = container.innerHTML;
    expect(restoreWordTextBoxes(container)).toBe(0);
    expect(container.innerHTML).toBe(original);
  });
});
