import { Injectable } from '@angular/core';

export type PresentationBlockKind = 'heading' | 'paragraph' | 'bullet' | 'number' | 'quote';

export interface PresentationBlock {
  kind: PresentationBlockKind;
  text: string;
  level?: number;
}

export interface EditablePresentationSlide {
  title: string;
  body: string;
  notes?: string;
  layout?: 'title-body' | 'title-only' | 'section';
  style?: {
    backgroundColor?: string;
    accentColor?: string;
    textColor?: string;
    titleFontSize?: number;
    bodyFontSize?: number;
    titleAlign?: 'left' | 'center' | 'right';
    bodyAlign?: 'left' | 'center' | 'right';
    titleBold?: boolean;
    bodyBold?: boolean;
    titleItalic?: boolean;
    bodyItalic?: boolean;
    titleUnderline?: boolean;
    bodyUnderline?: boolean;
  };
}

@Injectable({ providedIn: 'root' })
/** Convierte notas y presentaciones editables en archivos PPTX generados en memoria. */
export class PresentationExportService {
  /**
   * Exporta el modelo visual de presentaciones respetando layout, colores,
   * estilos de texto, notas del orador y límites defensivos de contenido.
   */
  async generateSlides(titleValue: string, sourceSlides: EditablePresentationSlide[]): Promise<Uint8Array> {
    const { default: PptxGenJS } = await import('pptxgenjs');
    const pptx = new PptxGenJS();
    pptx.layout = 'LAYOUT_WIDE';
    pptx.author = 'Helper Notes';
    pptx.company = 'Helper Notes';
    pptx.subject = 'Presentación creada con Helper Notes';
    pptx.title = titleValue.trim().slice(0, 250) || 'Presentación';
    pptx.theme = { headFontFace: 'Aptos Display', bodyFontFace: 'Aptos' };

    const slides = sourceSlides.length > 0 ? sourceSlides.slice(0, 500) : [{ title: titleValue, body: '' }];
    slides.forEach((source, index) => {
      const slide = pptx.addSlide();
      const style = source.style ?? {};
      const backgroundColor = (style.backgroundColor ?? (index === 0 ? '#fff7ed' : '#ffffff')).replace('#', '').toUpperCase();
      const accentColor = (style.accentColor ?? '#9a3412').replace('#', '').toUpperCase();
      const textColor = (style.textColor ?? '#334155').replace('#', '').toUpperCase();
      const centeredLayout = source.layout === 'section';
      const titleOnly = source.layout === 'title-only';
      slide.background = { color: backgroundColor };
      const title = String(source.title ?? '').trim().slice(0, 500) || (index === 0 ? titleValue : `Diapositiva ${index + 1}`);
      slide.addText(title, {
        x: 0.75, y: centeredLayout || titleOnly ? 2.45 : 0.55,
        w: 11.85, h: centeredLayout || titleOnly ? 1.2 : 0.8,
        fontFace: 'Aptos Display', fontSize: style.titleFontSize ?? (centeredLayout || titleOnly ? 36 : 28),
        bold: style.titleBold ?? true, italic: style.titleItalic ?? false,
        ...(style.titleUnderline ? { underline: { style: 'sng' as const } } : {}),
        color: accentColor, align: centeredLayout ? 'center' : (style.titleAlign ?? (titleOnly ? 'center' : 'left')),
        valign: 'middle', margin: 0.05, fit: 'shrink',
      });
      const body = String(source.body ?? '').trim().slice(0, 50_000);
      if (body && !titleOnly) {
        slide.addText(body, {
          x: 0.9, y: centeredLayout ? 3.75 : 1.65, w: 11.45, h: centeredLayout ? 2.2 : 4.95,
          fontFace: 'Aptos', fontSize: style.bodyFontSize ?? 21, color: textColor,
          bold: style.bodyBold ?? false, italic: style.bodyItalic ?? false,
          ...(style.bodyUnderline ? { underline: { style: 'sng' as const } } : {}),
          align: centeredLayout ? 'center' : (style.bodyAlign ?? 'left'),
          breakLine: false, valign: 'top', margin: 0.08, fit: 'shrink',
          paraSpaceAfter: 10,
        });
      }
      slide.addText(`${index + 1} / ${slides.length}`, {
        x: 11.7, y: 7.05, w: 0.8, h: 0.2,
        fontSize: 9, color: '94A3B8', align: 'right', margin: 0,
      });
      if (source.notes?.trim()) slide.addNotes(source.notes.trim().slice(0, 50_000));
    });

    const output = await pptx.write({ outputType: 'arraybuffer', compression: true });
    if (output instanceof ArrayBuffer) return new Uint8Array(output);
    if (output instanceof Uint8Array) return output;
    if (output instanceof Blob) return new Uint8Array(await output.arrayBuffer());
    throw new Error('No se pudo generar el archivo PowerPoint');
  }

  /**
   * Convierte una secuencia de bloques semánticos de una nota de texto en una
   * presentación paginada automáticamente.
   */
  async generate(titleValue: string, sourceBlocks: PresentationBlock[]): Promise<Uint8Array> {
    const { default: PptxGenJS } = await import('pptxgenjs');
    const pptx = new PptxGenJS();
    pptx.layout = 'LAYOUT_WIDE';
    pptx.author = 'Helper Notes';
    pptx.company = 'Helper Notes';
    pptx.subject = 'Nota exportada desde Helper Notes';
    pptx.title = titleValue;
    pptx.theme = { headFontFace: 'Aptos Display', bodyFontFace: 'Aptos' };

    const title = titleValue.trim().slice(0, 250) || 'Nota';
    const titleSlide = pptx.addSlide();
    titleSlide.background = { color: 'F8FAFC' };
    titleSlide.addText(title, {
      x: 0.85, y: 2.25, w: 11.65, h: 1.25,
      fontFace: 'Aptos Display', fontSize: 34, bold: true,
      color: '0F172A', align: 'center', valign: 'middle', margin: 0.05,
    });
    titleSlide.addText('Helper Notes', {
      x: 0.85, y: 3.75, w: 11.65, h: 0.35,
      fontSize: 13, color: '64748B', align: 'center', margin: 0,
    });

    const blocks = this.normalizedBlocks(sourceBlocks).flatMap(block =>
      this.splitLongText(block.text).map(text => ({ ...block, text })),
    );
    if (blocks.length > 0) {
      let slide = pptx.addSlide();
      let slideTitle = title;
      let y = 1.25;
      let slideNumber = 2;

      const decorateSlide = (): void => {
        slide.background = { color: 'FFFFFF' };
        slide.addText(slideTitle, {
          x: 0.65, y: 0.35, w: 11.8, h: 0.55,
          fontFace: 'Aptos Display', fontSize: 24, bold: true,
          color: '1D4ED8', margin: 0,
        });
        slide.addText(String(slideNumber), {
          x: 12.1, y: 7.1, w: 0.55, h: 0.18,
          fontSize: 9, color: '94A3B8', align: 'right', margin: 0,
        });
      };

      decorateSlide();
      const nextSlide = (): void => {
        slide = pptx.addSlide();
        slideNumber += 1;
        y = 1.25;
        decorateSlide();
      };

      for (const block of blocks) {
        const isHeading = block.kind === 'heading';
        const fontSize = isHeading ? Math.max(20, 27 - (block.level ?? 1) * 2) : 17;
        const estimatedLines = Math.max(1, Math.ceil(block.text.length / (isHeading ? 65 : 92)));
        const height = Math.max(isHeading ? 0.48 : 0.36, estimatedLines * (isHeading ? 0.34 : 0.27) + 0.08);
        if (y + height > 6.8) nextSlide();

        slide.addText(block.text, {
          x: block.kind === 'bullet' || block.kind === 'number' ? 1.05 : 0.8,
          y,
          w: block.kind === 'bullet' || block.kind === 'number' ? 11.35 : 11.7,
          h: height,
          fontFace: 'Aptos', fontSize, bold: isHeading,
          italic: block.kind === 'quote',
          color: isHeading ? '0F172A' : block.kind === 'quote' ? '475569' : '334155',
          valign: 'top', margin: 0.03,
          ...(block.kind === 'bullet' ? { bullet: { indent: 18 } } : {}),
        });
        y += height + (isHeading ? 0.18 : 0.1);
        if (isHeading) slideTitle = block.text.slice(0, 100);
      }
    }

    const output = await pptx.write({ outputType: 'arraybuffer', compression: true });
    if (output instanceof ArrayBuffer) return new Uint8Array(output);
    if (output instanceof Uint8Array) return output;
    if (output instanceof Blob) return new Uint8Array(await output.arrayBuffer());
    throw new Error('No se pudo generar el archivo PowerPoint');
  }

  /** Normaliza tipos, niveles y longitudes antes de calcular el layout. */
  private normalizedBlocks(blocks: PresentationBlock[]): PresentationBlock[] {
    return blocks.slice(0, 500).flatMap(block => {
      const text = String(block.text ?? '').replace(/\s+/g, ' ').trim().slice(0, 5000);
      if (!text) return [];
      const kind: PresentationBlockKind = ['heading', 'paragraph', 'bullet', 'number', 'quote'].includes(block.kind)
        ? block.kind
        : 'paragraph';
      return [{ kind, text, level: Math.max(1, Math.min(3, Number(block.level ?? 1))) }];
    });
  }

  /** Divide párrafos extensos procurando cortar primero en límites naturales. */
  private splitLongText(text: string, maxLength = 420): string[] {
    if (text.length <= maxLength) return [text];
    const chunks: string[] = [];
    let remaining = text;
    while (remaining.length > maxLength) {
      const candidate = remaining.slice(0, maxLength + 1);
      const splitAt = Math.max(candidate.lastIndexOf('. '), candidate.lastIndexOf('; '), candidate.lastIndexOf(' '));
      const safeSplit = splitAt > maxLength * 0.55 ? splitAt + 1 : maxLength;
      chunks.push(remaining.slice(0, safeSplit).trim());
      remaining = remaining.slice(safeSplit).trim();
    }
    if (remaining) chunks.push(remaining);
    return chunks;
  }
}
