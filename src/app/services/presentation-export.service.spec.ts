import { TestBed } from '@angular/core/testing';
import { PresentationExportService } from './presentation-export.service';

describe('PresentationExportService', () => {
  it('generates a valid PPTX ZIP from rich-text blocks', async () => {
    const service = TestBed.inject(PresentationExportService);

    const bytes = await service.generate('Nota de prueba', [
      { kind: 'heading', text: 'Introducción', level: 1 },
      { kind: 'paragraph', text: 'Contenido de la nota.' },
      { kind: 'bullet', text: 'Primer elemento' },
    ]);

    expect(bytes.length).toBeGreaterThan(1000);
    expect(Array.from(bytes.slice(0, 4))).toEqual([0x50, 0x4b, 0x03, 0x04]);
  });

  it('generates an editable PPTX from explicit slides and speaker notes', async () => {
    const service = TestBed.inject(PresentationExportService);
    const bytes = await service.generateSlides('Demo', [
      { title: 'Inicio', body: 'Primera diapositiva', notes: 'Saludar al público.' },
      { title: 'Cierre', body: 'Resumen final' },
    ]);

    expect(bytes.length).toBeGreaterThan(1000);
    expect(Array.from(bytes.slice(0, 4))).toEqual([0x50, 0x4b, 0x03, 0x04]);
  });
});
