import { Component, output } from '@angular/core';

@Component({
  selector: 'app-resize-handle',
  host: {
    '(mousedown)': 'onMouseDown($event)',
    'class': 'resize-handle',
  },
  template: ``,
  styles: `
    :host {
      display: block;
      width: 4px;
      cursor: col-resize;
      background: transparent;
      transition: background 0.15s;
      flex-shrink: 0;
      position: relative;
      z-index: 10;
      height: 100%;
    }
    :host:hover,
    :host:active {
      background: rgba(19, 91, 236, 0.4);
    }
  `,
})
/**
 * Tirador horizontal reutilizable.
 * Captura el puntero a nivel de documento y emite desplazamientos incrementales
 * para que el componente padre decida límites y dirección.
 */
export class ResizeHandle {
  readonly resizing = output<number>();

  private startX = 0;
  private dragging = false;

  /** Inicia el arrastre y registra listeners temporales hasta liberar el botón. */
  onMouseDown(event: MouseEvent): void {
    event.preventDefault();
    this.startX = event.clientX;
    this.dragging = true;

    const onMouseMove = (e: MouseEvent) => {
      if (!this.dragging) return;
      const delta = e.clientX - this.startX;
      this.startX = e.clientX;
      this.resizing.emit(delta);
    };

    const onMouseUp = () => {
      this.dragging = false;
      document.removeEventListener('mousemove', onMouseMove);
      document.removeEventListener('mouseup', onMouseUp);
      document.body.style.cursor = '';
      document.body.style.userSelect = '';
    };

    document.body.style.cursor = 'col-resize';
    document.body.style.userSelect = 'none';
    document.addEventListener('mousemove', onMouseMove);
    document.addEventListener('mouseup', onMouseUp);
  }
}
