import {
  createPresentationDocument,
  parsePresentationDocument,
} from './presentation.model';

describe('presentation model', () => {
  it('creates a document with one editable slide', () => {
    const document = createPresentationDocument();
    expect(document.kind).toBe('helper-notes-presentation');
    expect(document.slides).toHaveLength(1);
  });

  it('normalizes imported serialized presentation data', () => {
    const document = parsePresentationDocument(JSON.stringify({
      kind: 'helper-notes-presentation',
      slides: [{ id: 'one', title: 'Título', body: 'Contenido', notes: 'Notas' }],
    }));
    expect(document.slides[0]).toMatchObject({ id: 'one', title: 'Título', body: 'Contenido', notes: 'Notas', layout: 'title-body' });
    expect(document.slides[0].style.titleBold).toBe(true);
    expect(document.slides[0].transition).toEqual({ type: 'fade', duration: 0.45 });
  });

  it('preserves transition and element animation settings', () => {
    const document = parsePresentationDocument(JSON.stringify({
      kind: 'helper-notes-presentation',
      slides: [{
        id: 'animated', title: 'Animada', body: 'Contenido', notes: '',
        transition: { type: 'wipe', duration: 1.2 },
        animations: {
          title: { type: 'zoom', duration: 0.8, delay: 0.2 },
          body: { type: 'float', duration: 1, delay: 0.5 },
        },
      }],
    }));

    expect(document.slides[0].transition).toEqual({ type: 'wipe', duration: 1.2 });
    expect(document.slides[0].animations.title).toEqual({ type: 'zoom', duration: 0.8, delay: 0.2 });
    expect(document.slides[0].animations.body).toEqual({ type: 'float', duration: 1, delay: 0.5 });
  });

  it('preserves supported layouts and visual styles', () => {
    const document = parsePresentationDocument(JSON.stringify({
      kind: 'helper-notes-presentation',
      slides: [{
        id: 'styled', title: 'Sección', body: 'Contenido', notes: '', layout: 'section',
        style: { backgroundColor: '#0f172a', accentColor: '#fb923c', textColor: '#e2e8f0', titleAlign: 'center', titleFontSize: 54 },
      }],
    }));

    expect(document.slides[0].layout).toBe('section');
    expect(document.slides[0].style).toMatchObject({
      backgroundColor: '#0f172a', accentColor: '#fb923c', textColor: '#e2e8f0', titleAlign: 'center', titleFontSize: 54,
    });
  });
});
