export type PresentationTextAlign = 'left' | 'center' | 'right';
export type PresentationLayout = 'title-body' | 'title-only' | 'section';
export type PresentationTransitionType = 'none' | 'fade' | 'push' | 'zoom' | 'wipe';
export type PresentationAnimationType = 'none' | 'appear' | 'fade' | 'float' | 'zoom';

export interface PresentationTransition {
  type: PresentationTransitionType;
  duration: number;
}

export interface PresentationElementAnimation {
  type: PresentationAnimationType;
  duration: number;
  delay: number;
}

export interface PresentationSlideStyle {
  backgroundColor: string;
  accentColor: string;
  textColor: string;
  titleFontSize: number;
  bodyFontSize: number;
  titleAlign: PresentationTextAlign;
  bodyAlign: PresentationTextAlign;
  titleBold: boolean;
  bodyBold: boolean;
  titleItalic: boolean;
  bodyItalic: boolean;
  titleUnderline: boolean;
  bodyUnderline: boolean;
}

export interface PresentationSlide {
  id: string;
  title: string;
  body: string;
  notes: string;
  layout: PresentationLayout;
  style: PresentationSlideStyle;
  transition: PresentationTransition;
  animations: {
    title: PresentationElementAnimation;
    body: PresentationElementAnimation;
  };
}

/** Crea la apariencia predeterminada usada por una diapositiva nueva. */
export function createPresentationSlideStyle(): PresentationSlideStyle {
  return {
    backgroundColor: '#ffffff',
    accentColor: '#9a3412',
    textColor: '#334155',
    titleFontSize: 38,
    bodyFontSize: 22,
    titleAlign: 'left',
    bodyAlign: 'left',
    titleBold: true,
    bodyBold: false,
    titleItalic: false,
    bodyItalic: false,
    titleUnderline: false,
    bodyUnderline: false,
  };
}

/** Crea una animación neutral con tiempos válidos para el editor. */
export function createPresentationElementAnimation(): PresentationElementAnimation {
  return { type: 'none', duration: 0.6, delay: 0 };
}

export interface PresentationDocument {
  kind: 'helper-notes-presentation';
  version: 1;
  slides: PresentationSlide[];
}

/**
 * Crea una diapositiva con identificador único y valores seguros para todos
 * los controles del editor.
 */
export function createPresentationSlide(index = 0): PresentationSlide {
  return {
    id: typeof crypto !== 'undefined' && crypto.randomUUID
      ? crypto.randomUUID()
      : `slide-${Date.now()}-${index}`,
    title: '',
    body: '',
    notes: '',
    layout: 'title-body',
    style: createPresentationSlideStyle(),
    transition: { type: 'fade', duration: 0.45 },
    animations: {
      title: createPresentationElementAnimation(),
      body: createPresentationElementAnimation(),
    },
  };
}

/** Crea un documento de presentación versión 1 con una diapositiva inicial. */
export function createPresentationDocument(): PresentationDocument {
  return {
    kind: 'helper-notes-presentation',
    version: 1,
    slides: [createPresentationSlide()],
  };
}

/**
 * Deserializa y normaliza una presentación guardada.
 *
 * Además de validar el discriminante y la versión implícita, limita tamaños,
 * colores, tipografías y duraciones para impedir que contenido antiguo o
 * manipulado deje el editor en un estado inválido. Ante cualquier error se
 * devuelve un documento nuevo utilizable.
 */
export function parsePresentationDocument(content: string): PresentationDocument {
  try {
    const parsed = JSON.parse(content) as Partial<PresentationDocument>;
    if (parsed.kind !== 'helper-notes-presentation' || !Array.isArray(parsed.slides)) {
      return createPresentationDocument();
    }
    const slides: PresentationSlide[] = parsed.slides.slice(0, 500).map((slide, index): PresentationSlide => {
      const fallback = createPresentationSlide(index);
      const source = slide as Partial<PresentationSlide>;
      const rawStyle = source.style ?? fallback.style;
      const transitionTypes: PresentationTransitionType[] = ['none', 'fade', 'push', 'zoom', 'wipe'];
      const animationTypes: PresentationAnimationType[] = ['none', 'appear', 'fade', 'float', 'zoom'];
      const duration = (value: unknown, defaultValue: number): number =>
        Math.max(0.1, Math.min(5, Number(value) || defaultValue));
      const delay = (value: unknown, defaultValue: number): number =>
        Math.max(0, Math.min(10, Number(value) || defaultValue));
      const color = (value: unknown, defaultValue: string): string =>
        typeof value === 'string' && /^#[0-9a-f]{6}$/i.test(value) ? value : defaultValue;
      const size = (value: unknown, defaultValue: number, min: number, max: number): number =>
        Math.max(min, Math.min(max, Number(value) || defaultValue));
      const align = (value: unknown, defaultValue: PresentationTextAlign): PresentationTextAlign =>
        value === 'center' || value === 'right' || value === 'left' ? value : defaultValue;
      return {
        id: String(source.id || fallback.id),
        title: String(source.title ?? '').slice(0, 500),
        body: String(source.body ?? '').slice(0, 50_000),
        notes: String(source.notes ?? '').slice(0, 50_000),
        layout: source.layout === 'title-only' || source.layout === 'section' ? source.layout : 'title-body',
        style: {
          backgroundColor: color(rawStyle.backgroundColor, fallback.style.backgroundColor),
          accentColor: color(rawStyle.accentColor, fallback.style.accentColor),
          textColor: color(rawStyle.textColor, fallback.style.textColor),
          titleFontSize: size(rawStyle.titleFontSize, fallback.style.titleFontSize, 18, 72),
          bodyFontSize: size(rawStyle.bodyFontSize, fallback.style.bodyFontSize, 10, 48),
          titleAlign: align(rawStyle.titleAlign, fallback.style.titleAlign),
          bodyAlign: align(rawStyle.bodyAlign, fallback.style.bodyAlign),
          titleBold: rawStyle.titleBold ?? fallback.style.titleBold,
          bodyBold: rawStyle.bodyBold ?? fallback.style.bodyBold,
          titleItalic: rawStyle.titleItalic ?? fallback.style.titleItalic,
          bodyItalic: rawStyle.bodyItalic ?? fallback.style.bodyItalic,
          titleUnderline: rawStyle.titleUnderline ?? fallback.style.titleUnderline,
          bodyUnderline: rawStyle.bodyUnderline ?? fallback.style.bodyUnderline,
        },
        transition: {
          type: transitionTypes.includes(source.transition?.type as PresentationTransitionType)
            ? source.transition!.type
            : fallback.transition.type,
          duration: duration(source.transition?.duration, fallback.transition.duration),
        },
        animations: {
          title: {
            type: animationTypes.includes(source.animations?.title?.type as PresentationAnimationType)
              ? source.animations!.title.type
              : fallback.animations.title.type,
            duration: duration(source.animations?.title?.duration, fallback.animations.title.duration),
            delay: delay(source.animations?.title?.delay, fallback.animations.title.delay),
          },
          body: {
            type: animationTypes.includes(source.animations?.body?.type as PresentationAnimationType)
              ? source.animations!.body.type
              : fallback.animations.body.type,
            duration: duration(source.animations?.body?.duration, fallback.animations.body.duration),
            delay: delay(source.animations?.body?.delay, fallback.animations.body.delay),
          },
        },
      };
    });
    return {
      kind: 'helper-notes-presentation',
      version: 1,
      slides: slides.length > 0 ? slides : [createPresentationSlide()],
    };
  } catch {
    return createPresentationDocument();
  }
}
