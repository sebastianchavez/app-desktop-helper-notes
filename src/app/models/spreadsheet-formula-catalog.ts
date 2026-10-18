/** Ayuda de fórmulas en español e inglés; las firmas documentan el subconjunto que calcula este editor. */
export const ADDITIONAL_FORMULA_SUGGESTIONS: {
  name: string;
  syntax: string;
  description: string;
}[] = [
  {
    name: 'SUMAR.SI.CONJUNTO',
    syntax: 'SUMAR.SI.CONJUNTO(rango_suma; rango1; criterio1; …)',
    description: 'Suma valores que cumplen varios criterios.',
  },
  {
    name: 'SUMIFS',
    syntax: 'SUMIFS(rango_suma; rango1; criterio1; …)',
    description: 'Suma valores que cumplen varios criterios.',
  },
  {
    name: 'CONTAR.SI.CONJUNTO',
    syntax: 'CONTAR.SI.CONJUNTO(rango1; criterio1; …)',
    description: 'Cuenta filas que cumplen todos los criterios.',
  },
  {
    name: 'COUNTIFS',
    syntax: 'COUNTIFS(rango1; criterio1; …)',
    description: 'Cuenta filas que cumplen todos los criterios.',
  },
  {
    name: 'PROMEDIO.SI',
    syntax: 'PROMEDIO.SI(rango; criterio; rango_promedio)',
    description: 'Promedia valores que cumplen un criterio.',
  },
  {
    name: 'AVERAGEIF',
    syntax: 'AVERAGEIF(rango; criterio; rango_promedio)',
    description: 'Promedia valores que cumplen un criterio.',
  },
  {
    name: 'PROMEDIO.SI.CONJUNTO',
    syntax: 'PROMEDIO.SI.CONJUNTO(rango_promedio; rango1; criterio1; …)',
    description: 'Promedia valores con varios criterios.',
  },
  {
    name: 'AVERAGEIFS',
    syntax: 'AVERAGEIFS(rango_promedio; rango1; criterio1; …)',
    description: 'Promedia valores con varios criterios.',
  },
  {
    name: 'MIN.SI.CONJUNTO',
    syntax: 'MIN.SI.CONJUNTO(rango_mínimo; rango1; criterio1; …)',
    description: 'Obtiene el mínimo con criterios.',
  },
  {
    name: 'MINIFS',
    syntax: 'MINIFS(rango_mínimo; rango1; criterio1; …)',
    description: 'Obtiene el mínimo con criterios.',
  },
  {
    name: 'MAX.SI.CONJUNTO',
    syntax: 'MAX.SI.CONJUNTO(rango_máximo; rango1; criterio1; …)',
    description: 'Obtiene el máximo con criterios.',
  },
  {
    name: 'MAXIFS',
    syntax: 'MAXIFS(rango_máximo; rango1; criterio1; …)',
    description: 'Obtiene el máximo con criterios.',
  },
  {
    name: 'SUMAPRODUCTO',
    syntax: 'SUMAPRODUCTO(matriz1; matriz2; …)',
    description: 'Suma productos de matrices del mismo tamaño.',
  },
  {
    name: 'SUMPRODUCT',
    syntax: 'SUMPRODUCT(matriz1; matriz2; …)',
    description: 'Suma productos de matrices del mismo tamaño.',
  },
  {
    name: 'BUSCARX',
    syntax: 'BUSCARX(valor; rango_búsqueda; rango_resultado; [si_no]; [modo]; [dirección])',
    description: 'Busca por coincidencia exacta, aproximada o comodines; dirección 1 o -1.',
  },
  {
    name: 'XLOOKUP',
    syntax: 'XLOOKUP(valor; rango_búsqueda; rango_resultado; [si_no]; [modo]; [dirección])',
    description: 'Busca por coincidencia exacta, aproximada o comodines; dirección 1 o -1.',
  },
  {
    name: 'COINCIDIRX',
    syntax: 'COINCIDIRX(valor; rango; [modo]; [dirección])',
    description: 'Devuelve la posición de una coincidencia.',
  },
  {
    name: 'XMATCH',
    syntax: 'XMATCH(valor; rango; [modo]; [dirección])',
    description: 'Devuelve la posición de una coincidencia.',
  },
  {
    name: 'BUSCARV',
    syntax: 'BUSCARV(valor; tabla; columna; [aproximada])',
    description: 'Busca en la primera columna de una tabla.',
  },
  {
    name: 'VLOOKUP',
    syntax: 'VLOOKUP(valor; tabla; columna; [aproximada])',
    description: 'Busca en la primera columna de una tabla.',
  },
  {
    name: 'BUSCARH',
    syntax: 'BUSCARH(valor; tabla; fila; [aproximada])',
    description: 'Busca en la primera fila de una tabla.',
  },
  {
    name: 'HLOOKUP',
    syntax: 'HLOOKUP(valor; tabla; fila; [aproximada])',
    description: 'Busca en la primera fila de una tabla.',
  },
  {
    name: 'INDICE',
    syntax: 'INDICE(tabla; fila; [columna])',
    description: 'Obtiene un valor por fila y columna.',
  },
  {
    name: 'INDEX',
    syntax: 'INDEX(tabla; fila; [columna])',
    description: 'Obtiene un valor por fila y columna.',
  },
  {
    name: 'COINCIDIR',
    syntax: 'COINCIDIR(valor; rango; [modo])',
    description: 'Busca una posición; modo 0 exacto, 1 o -1 aproximado.',
  },
  {
    name: 'MATCH',
    syntax: 'MATCH(valor; rango; [modo])',
    description: 'Busca una posición; modo 0 exacto, 1 o -1 aproximado.',
  },
  {
    name: 'SI.ERROR',
    syntax: 'SI.ERROR(valor; alternativa)',
    description: 'Sustituye un resultado de error.',
  },
  {
    name: 'IFERROR',
    syntax: 'IFERROR(valor; alternativa)',
    description: 'Sustituye un resultado de error.',
  },
  {
    name: 'SI.ND',
    syntax: 'SI.ND(valor; alternativa)',
    description: 'Sustituye únicamente el error #N/A.',
  },
  {
    name: 'IFNA',
    syntax: 'IFNA(valor; alternativa)',
    description: 'Sustituye únicamente el error #N/A.',
  },
  {
    name: 'SI.CONJUNTO',
    syntax: 'SI.CONJUNTO(condición1; resultado1; …)',
    description: 'Evalúa condiciones en orden sin calcular ramas innecesarias.',
  },
  {
    name: 'IFS',
    syntax: 'IFS(condición1; resultado1; …)',
    description: 'Evalúa condiciones en orden sin calcular ramas innecesarias.',
  },
  {
    name: 'CAMBIAR',
    syntax: 'CAMBIAR(valor; caso1; resultado1; …; [predeterminado])',
    description: 'Selecciona el primer caso coincidente.',
  },
  {
    name: 'SWITCH',
    syntax: 'SWITCH(valor; caso1; resultado1; …; [predeterminado])',
    description: 'Selecciona el primer caso coincidente.',
  },
  {
    name: 'ESERROR',
    syntax: 'ESERROR(valor)',
    description: 'Indica si un valor es un error.',
  },
  {
    name: 'ISERROR',
    syntax: 'ISERROR(valor)',
    description: 'Indica si un valor es un error.',
  },
  {
    name: 'ESNUMERO',
    syntax: 'ESNUMERO(valor)',
    description: 'Indica si un valor es numérico.',
  },
  {
    name: 'ISNUMBER',
    syntax: 'ISNUMBER(valor)',
    description: 'Indica si un valor es numérico.',
  },
  {
    name: 'ESTEXTO',
    syntax: 'ESTEXTO(valor)',
    description: 'Indica si un valor es texto.',
  },
  {
    name: 'ISTEXT',
    syntax: 'ISTEXT(valor)',
    description: 'Indica si un valor es texto.',
  },
  {
    name: 'ESBLANCO',
    syntax: 'ESBLANCO(valor)',
    description: 'Detecta celdas vacías.',
  },
  {
    name: 'ISBLANK',
    syntax: 'ISBLANK(valor)',
    description: 'Detecta celdas vacías.',
  },
  {
    name: 'CONTAR.BLANCO',
    syntax: 'CONTAR.BLANCO(rango)',
    description: 'Cuenta celdas vacías y cadenas vacías.',
  },
  {
    name: 'COUNTBLANK',
    syntax: 'COUNTBLANK(rango)',
    description: 'Cuenta celdas vacías y cadenas vacías.',
  },
  {
    name: 'MEDIANA',
    syntax: 'MEDIANA(número1; número2; …)',
    description: 'Calcula la mediana.',
  },
  {
    name: 'MEDIAN',
    syntax: 'MEDIAN(número1; número2; …)',
    description: 'Calcula la mediana.',
  },
  {
    name: 'K.ESIMO.MAYOR',
    syntax: 'K.ESIMO.MAYOR(rango; k)',
    description: 'Obtiene el k-ésimo mayor.',
  },
  {
    name: 'LARGE',
    syntax: 'LARGE(rango; k)',
    description: 'Obtiene el k-ésimo mayor.',
  },
  {
    name: 'K.ESIMO.MENOR',
    syntax: 'K.ESIMO.MENOR(rango; k)',
    description: 'Obtiene el k-ésimo menor.',
  },
  {
    name: 'SMALL',
    syntax: 'SMALL(rango; k)',
    description: 'Obtiene el k-ésimo menor.',
  },
  {
    name: 'DESVEST.M',
    syntax: 'DESVEST.M(número1; …)',
    description: 'Calcula desviación estándar muestral.',
  },
  {
    name: 'STDEV.S',
    syntax: 'STDEV.S(número1; …)',
    description: 'Calcula desviación estándar muestral.',
  },
  {
    name: 'DESVEST.P',
    syntax: 'DESVEST.P(número1; …)',
    description: 'Calcula desviación estándar poblacional.',
  },
  {
    name: 'STDEV.P',
    syntax: 'STDEV.P(número1; …)',
    description: 'Calcula desviación estándar poblacional.',
  },
  {
    name: 'VAR.S',
    syntax: 'VAR.S(número1; …)',
    description: 'Calcula varianza muestral.',
  },
  {
    name: 'VAR.P',
    syntax: 'VAR.P(número1; …)',
    description: 'Calcula varianza poblacional.',
  },
  {
    name: 'UNIRCADENAS',
    syntax: 'UNIRCADENAS(separador; ignorar_vacíos; texto1; …)',
    description: 'Une textos o rangos con separador.',
  },
  {
    name: 'TEXTJOIN',
    syntax: 'TEXTJOIN(separador; ignorar_vacíos; texto1; …)',
    description: 'Une textos o rangos con separador.',
  },
  {
    name: 'SUSTITUIR',
    syntax: 'SUSTITUIR(texto; anterior; nuevo; [ocurrencia])',
    description: 'Sustituye coincidencias de texto.',
  },
  {
    name: 'SUBSTITUTE',
    syntax: 'SUBSTITUTE(texto; anterior; nuevo; [ocurrencia])',
    description: 'Sustituye coincidencias de texto.',
  },
  {
    name: 'REEMPLAZAR',
    syntax: 'REEMPLAZAR(texto; inicio; cantidad; nuevo)',
    description: 'Reemplaza una parte del texto.',
  },
  {
    name: 'REPLACE',
    syntax: 'REPLACE(texto; inicio; cantidad; nuevo)',
    description: 'Reemplaza una parte del texto.',
  },
  {
    name: 'ENCONTRAR',
    syntax: 'ENCONTRAR(buscar; texto; [inicio])',
    description: 'Busca texto distinguiendo mayúsculas.',
  },
  {
    name: 'FIND',
    syntax: 'FIND(buscar; texto; [inicio])',
    description: 'Busca texto distinguiendo mayúsculas.',
  },
  {
    name: 'HALLAR',
    syntax: 'HALLAR(buscar; texto; [inicio])',
    description: 'Busca sin distinguir mayúsculas y admite comodines.',
  },
  {
    name: 'SEARCH',
    syntax: 'SEARCH(buscar; texto; [inicio])',
    description: 'Busca sin distinguir mayúsculas y admite comodines.',
  },
  {
    name: 'NOMPROPIO',
    syntax: 'NOMPROPIO(texto)',
    description: 'Pone iniciales en mayúscula.',
  },
  {
    name: 'PROPER',
    syntax: 'PROPER(texto)',
    description: 'Pone iniciales en mayúscula.',
  },
  {
    name: 'REPETIR',
    syntax: 'REPETIR(texto; veces)',
    description: 'Repite un texto.',
  },
  {
    name: 'REPT',
    syntax: 'REPT(texto; veces)',
    description: 'Repite un texto.',
  },
  {
    name: 'LIMPIAR',
    syntax: 'LIMPIAR(texto)',
    description: 'Quita controles ASCII no imprimibles.',
  },
  {
    name: 'CLEAN',
    syntax: 'CLEAN(texto)',
    description: 'Quita controles ASCII no imprimibles.',
  },
  {
    name: 'FECHA',
    syntax: 'FECHA(año; mes; día)',
    description: 'Construye un serial usando la época del libro.',
  },
  {
    name: 'DATE',
    syntax: 'DATE(año; mes; día)',
    description: 'Construye un serial usando la época del libro.',
  },
  {
    name: 'HOY',
    syntax: 'HOY()',
    description: 'Devuelve la fecha local actual.',
  },
  {
    name: 'TODAY',
    syntax: 'TODAY()',
    description: 'Devuelve la fecha local actual.',
  },
  {
    name: 'AHORA',
    syntax: 'AHORA()',
    description: 'Devuelve fecha y hora local actuales.',
  },
  {
    name: 'NOW',
    syntax: 'NOW()',
    description: 'Devuelve fecha y hora local actuales.',
  },
  {
    name: 'AÑO',
    syntax: 'AÑO(fecha)',
    description: 'Extrae el año de un serial.',
  },
  {
    name: 'YEAR',
    syntax: 'YEAR(fecha)',
    description: 'Extrae el año de un serial.',
  },
  {
    name: 'MES',
    syntax: 'MES(fecha)',
    description: 'Extrae el mes de un serial.',
  },
  {
    name: 'MONTH',
    syntax: 'MONTH(fecha)',
    description: 'Extrae el mes de un serial.',
  },
  {
    name: 'DIA',
    syntax: 'DIA(fecha)',
    description: 'Extrae el día de un serial.',
  },
  {
    name: 'DAY',
    syntax: 'DAY(fecha)',
    description: 'Extrae el día de un serial.',
  },
  {
    name: 'DIAS',
    syntax: 'DIAS(final; inicial)',
    description: 'Calcula diferencia de días.',
  },
  {
    name: 'DAYS',
    syntax: 'DAYS(final; inicial)',
    description: 'Calcula diferencia de días.',
  },
  {
    name: 'FECHA.MES',
    syntax: 'FECHA.MES(fecha; meses)',
    description: 'Desplaza una fecha por meses.',
  },
  {
    name: 'EDATE',
    syntax: 'EDATE(fecha; meses)',
    description: 'Desplaza una fecha por meses.',
  },
  {
    name: 'FIN.MES',
    syntax: 'FIN.MES(fecha; meses)',
    description: 'Obtiene el último día del mes.',
  },
  {
    name: 'EOMONTH',
    syntax: 'EOMONTH(fecha; meses)',
    description: 'Obtiene el último día del mes.',
  },
  {
    name: 'DIASEM',
    syntax: 'DIASEM(fecha; [tipo])',
    description: 'Obtiene el día de semana.',
  },
  {
    name: 'WEEKDAY',
    syntax: 'WEEKDAY(fecha; [tipo])',
    description: 'Obtiene el día de semana.',
  },
  {
    name: 'DIA.LAB',
    syntax: 'DIA.LAB(fecha; días; [feriados])',
    description: 'Desplaza por días laborables de lunes a viernes.',
  },
  {
    name: 'WORKDAY',
    syntax: 'WORKDAY(fecha; días; [feriados])',
    description: 'Desplaza por días laborables de lunes a viernes.',
  },
  {
    name: 'DIAS.LAB',
    syntax: 'DIAS.LAB(inicial; final; [feriados])',
    description: 'Cuenta días laborables inclusivos.',
  },
  {
    name: 'NETWORKDAYS',
    syntax: 'NETWORKDAYS(inicial; final; [feriados])',
    description: 'Cuenta días laborables inclusivos.',
  },
  {
    name: 'NSHORA',
    syntax: 'NSHORA(hora; minuto; segundo)',
    description: 'Construye una fracción de día.',
  },
  {
    name: 'TIME',
    syntax: 'TIME(hora; minuto; segundo)',
    description: 'Construye una fracción de día.',
  },
  {
    name: 'HORA',
    syntax: 'HORA(hora)',
    description: 'Extrae la hora.',
  },
  {
    name: 'HOUR',
    syntax: 'HOUR(hora)',
    description: 'Extrae la hora.',
  },
  {
    name: 'MINUTO',
    syntax: 'MINUTO(hora)',
    description: 'Extrae el minuto.',
  },
  {
    name: 'MINUTE',
    syntax: 'MINUTE(hora)',
    description: 'Extrae el minuto.',
  },
  {
    name: 'SEGUNDO',
    syntax: 'SEGUNDO(hora)',
    description: 'Extrae el segundo.',
  },
  {
    name: 'SECOND',
    syntax: 'SECOND(hora)',
    description: 'Extrae el segundo.',
  },
  {
    name: 'FILA',
    syntax: 'FILA([referencia])',
    description: 'Obtiene número de fila.',
  },
  {
    name: 'ROW',
    syntax: 'ROW([referencia])',
    description: 'Obtiene número de fila.',
  },
  {
    name: 'COLUMNA',
    syntax: 'COLUMNA([referencia])',
    description: 'Obtiene número de columna.',
  },
  {
    name: 'COLUMN',
    syntax: 'COLUMN([referencia])',
    description: 'Obtiene número de columna.',
  },
  {
    name: 'FILAS',
    syntax: 'FILAS(rango)',
    description: 'Cuenta filas del rango.',
  },
  {
    name: 'ROWS',
    syntax: 'ROWS(rango)',
    description: 'Cuenta filas del rango.',
  },
  {
    name: 'COLUMNAS',
    syntax: 'COLUMNAS(rango)',
    description: 'Cuenta columnas del rango.',
  },
  {
    name: 'COLUMNS',
    syntax: 'COLUMNS(rango)',
    description: 'Cuenta columnas del rango.',
  },
];

/** Ayuda de matrices dinámicas: los resultados se expanden desde una única celda ancla. */
ADDITIONAL_FORMULA_SUGGESTIONS.push(
  {
    name: 'FILTRAR',
    syntax: 'FILTRAR(matriz; incluir; [si_vacío])',
    description: 'Filtra filas o columnas y expande el resultado.',
  },
  {
    name: 'FILTER',
    syntax: 'FILTER(matriz; incluir; [si_vacío])',
    description: 'Filtra filas o columnas y expande el resultado.',
  },
  {
    name: 'ORDENAR',
    syntax: 'ORDENAR(matriz; [índice]; [orden]; [por_columna])',
    description: 'Ordena una matriz con expansión automática.',
  },
  {
    name: 'SORT',
    syntax: 'SORT(matriz; [índice]; [orden]; [por_columna])',
    description: 'Ordena una matriz con expansión automática.',
  },
  {
    name: 'UNICOS',
    syntax: 'UNICOS(matriz; [por_columna]; [solo_una_vez])',
    description: 'Devuelve filas o columnas únicas.',
  },
  {
    name: 'UNIQUE',
    syntax: 'UNIQUE(matriz; [por_columna]; [solo_una_vez])',
    description: 'Devuelve filas o columnas únicas.',
  },
  {
    name: 'SECUENCIA',
    syntax: 'SECUENCIA(filas; [columnas]; [inicio]; [paso])',
    description: 'Genera una secuencia rectangular.',
  },
  {
    name: 'SEQUENCE',
    syntax: 'SEQUENCE(filas; [columnas]; [inicio]; [paso])',
    description: 'Genera una secuencia rectangular.',
  },
  {
    name: 'TRANSPONER',
    syntax: 'TRANSPONER(matriz)',
    description: 'Intercambia filas por columnas.',
  },
  {
    name: 'TRANSPOSE',
    syntax: 'TRANSPOSE(matriz)',
    description: 'Intercambia filas por columnas.',
  },
);

/** Firmas nuevas compartidas entre nombres ingleses y alias españoles para evitar duplicar descripciones en el paquete. */
const EXTRA_FORMULA_HELP = [
  {
    name: 'SUBTOTAL',
    syntax: 'SUBTOTAL(código; rango; …)',
    description:
      'Agrega datos excluyendo subtotales anidados. Los códigos 101–111 omiten filas ocultas.',
  },
  {
    name: 'PMT',
    syntax: 'PMT(tasa; períodos; valor_actual; valor_futuro; tipo)',
    description: 'Pago periódico de una anualidad.',
  },
  {
    name: 'PV',
    syntax: 'PV(tasa; períodos; pago; valor_futuro; tipo)',
    description: 'Valor actual de pagos periódicos.',
  },
  {
    name: 'FV',
    syntax: 'FV(tasa; períodos; pago; valor_actual; tipo)',
    description: 'Valor futuro de pagos periódicos.',
  },
  {
    name: 'NPER',
    syntax: 'NPER(tasa; pago; valor_actual; valor_futuro; tipo)',
    description: 'Cantidad de períodos.',
  },
  {
    name: 'IPMT',
    syntax: 'IPMT(tasa; período; períodos; valor_actual; valor_futuro; tipo)',
    description: 'Interés correspondiente a un período.',
  },
  {
    name: 'PPMT',
    syntax: 'PPMT(tasa; período; períodos; valor_actual; valor_futuro; tipo)',
    description: 'Amortización de capital de un período.',
  },
  {
    name: 'NPV',
    syntax: 'NPV(tasa; flujos; …)',
    description: 'Valor actual neto de flujos al final de cada período.',
  },
  {
    name: 'MROUND',
    syntax: 'MROUND(número; múltiplo)',
    description: 'Redondea al múltiplo más cercano.',
  },
  {
    name: 'LN',
    syntax: 'LN(número)',
    description: 'Logaritmo natural.',
  },
  {
    name: 'LOG10',
    syntax: 'LOG10(número)',
    description: 'Logaritmo decimal.',
  },
  {
    name: 'LOG',
    syntax: 'LOG(número; base)',
    description: 'Logaritmo en la base indicada.',
  },
  {
    name: 'EXP',
    syntax: 'EXP(número)',
    description: 'Potencia de e.',
  },
  {
    name: 'FACT',
    syntax: 'FACT(número)',
    description: 'Factorial de un entero truncado.',
  },
  {
    name: 'GCD',
    syntax: 'GCD(números; …)',
    description: 'Máximo común divisor.',
  },
  {
    name: 'LCM',
    syntax: 'LCM(números; …)',
    description: 'Mínimo común múltiplo.',
  },
  {
    name: 'T',
    syntax: 'T(valor)',
    description: 'Conserva texto y errores; convierte otros tipos en texto vacío.',
  },
  {
    name: 'NUMBERVALUE',
    syntax: 'NUMBERVALUE(texto; decimal; grupo)',
    description: 'Convierte texto numérico con separadores explícitos.',
  },
  {
    name: 'TEXTBEFORE',
    syntax:
      'TEXTBEFORE(texto; delimitador; instancia; ignorar_mayúsculas; coincidir_final; no_encontrado)',
    description: 'Texto anterior a una aparición del delimitador.',
  },
  {
    name: 'TEXTAFTER',
    syntax:
      'TEXTAFTER(texto; delimitador; instancia; ignorar_mayúsculas; coincidir_final; no_encontrado)',
    description: 'Texto posterior a una aparición del delimitador.',
  },
];
const EXTRA_FORMULA_ALIASES: Record<string, string> = {
  SUBTOTALES: 'SUBTOTAL',
  PAGO: 'PMT',
  VA: 'PV',
  VF: 'FV',
  PAGOINT: 'IPMT',
  PAGOPRIN: 'PPMT',
  VNA: 'NPV',
  'REDOND.MULT': 'MROUND',
  'M.C.D': 'GCD',
  'M.C.M': 'LCM',
  'VALOR.NUMERO': 'NUMBERVALUE',
  TEXTOANTES: 'TEXTBEFORE',
  TEXTODESPUES: 'TEXTAFTER',
};
ADDITIONAL_FORMULA_SUGGESTIONS.push(...EXTRA_FORMULA_HELP);
for (const [alias, name] of Object.entries(EXTRA_FORMULA_ALIASES)) {
  const help = EXTRA_FORMULA_HELP.find((h) => h.name === name)!;
  ADDITIONAL_FORMULA_SUGGESTIONS.push({
    ...help,
    name: alias,
    syntax: help.syntax.replace(name, alias),
  });
}
