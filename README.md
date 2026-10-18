# Helper Notes

Helper Notes es una aplicación de escritorio para crear y organizar notas por carpetas. Su experiencia toma como referencia la aplicación Notas de macOS —barra lateral, lista de notas y editor—, pero amplía ese concepto con documentos enriquecidos, hojas de cálculo, importación y exportación de archivos y un asistente de IA local.

Está desarrollada con Angular, Electron y SQLite. La información se guarda localmente en el equipo y el asistente funciona mediante Ollama, sin depender de un servicio de IA en la nube.

> Helper Notes es un proyecto independiente y no está afiliado a Apple.

## Diferencias frente a Notas de macOS

- Incluye un editor de hojas de cálculo dentro de la propia aplicación.
- Importa archivos TXT, HTML, DOCX, PDF, XLSX, XLS, CSV y PPTX, además de respaldos JSON.
- Al soltar documentos compatibles sobre cualquier zona de la aplicación, los importa en la carpeta activa y abre el último. Los originales permanecen intactos.
- Los DOCX se editan directamente en páginas mediante DOCX Editor (núcleo Apache 2.0): texto, formato, imágenes, tablas, márgenes y orientación. Los cambios se guardan en el DOCX completo y «Guardar archivo» exporta la versión editada. No hace falta convertir el documento en una nota aparte.
- TXT se abre como texto plano editable, conservando espacios, tabulaciones y saltos de línea. PDF usa el visor del archivo original y XLSX muestra fuentes, colores, bordes, celdas combinadas y valores formateados del libro importado.
- El instalador de Windows agrega «Abrir con Helper Notes» al menú contextual y Helper Notes a «Abrir con» para TXT, HTML/HTM, DOCX, PDF, XLSX, XLS, CSV y PPTX. Las asociaciones se registran al instalar o actualizar y se retiran al desinstalar, sin cambiar la aplicación predeterminada.
- Exporta notas como TXT, HTML, DOCX, PPTX, PDF, XLSX o JSON.
- Imprime notas, PDF, la hoja de cálculo activa y presentaciones mediante el diálogo nativo (`Ctrl+P`).
- Integra modelos locales de Ollama que pueden conversar sobre la nota abierta y proponer cambios.
- Permite aplicar o rechazar individualmente las acciones sugeridas por la IA.
- Usa SQLite local en lugar de iCloud. Actualmente no incluye sincronización entre dispositivos ni colaboración en línea.
- Genera instaladores nativos para Windows, Linux y macOS.

## Funcionalidades

### Notas y carpetas

- Crear, renombrar y eliminar carpetas y subcarpetas organizadas como árbol.
- Menú contextual con clic derecho sobre «Todas las notas» o una carpeta para crear carpetas, subcarpetas, notas, hojas de cálculo y presentaciones en la ubicación elegida.
- Ver todas las notas o filtrar por carpeta.
- Buscar por título y contenido.
- Mover notas entre carpetas mediante arrastrar y soltar.
- Abrir una nota en una ventana independiente con doble clic.
- Sincronizar los cambios entre las ventanas abiertas de la aplicación.
- Eliminación suave de notas y carpetas en la base de datos.
- Paneles laterales redimensionables y plegables.

### Editor de texto enriquecido

- Títulos, párrafos y encabezados H1, H2 y H3.
- Negrita, cursiva, subrayado, tachado y limpieza de formato.
- Siete tamaños de fuente, color de texto y resaltado.
- Alineación izquierda, centrada, derecha y justificada.
- Listas ordenadas y no ordenadas, con sangría y reducción de sangría.
- Imágenes desde archivo, portapapeles o arrastrar y soltar, almacenadas dentro de la nota.
- Ajuste de línea opcional.
- Conteo de palabras y caracteres y fecha de última edición.
- Guardado automático 500 ms después del último cambio, conservando correctamente los cambios pendientes al cambiar de nota o cerrar el editor.

### Hojas de cálculo

- Crear una hoja en blanco o importar archivos XLSX, XLS y CSV.
- Navegar entre las hojas de un libro.
- Editar celdas desde la cuadrícula o la barra de fórmulas; el cambio se guarda al confirmar o al salir de la celda.
- Comenzar a escribir directamente sobre una celda seleccionada sin perder el primer carácter.
- Confirmar con `Enter` y avanzar a la fila siguiente (`Shift+Enter` retrocede); durante la edición, las flechas guardan el cambio y mueven la selección a la celda contigua.
- Seleccionar rangos, filas o columnas.
- Cortar, copiar, pegar, insertar y eliminar filas o columnas.
- Mover una o varias filas completas hacia arriba o abajo y columnas completas hacia la izquierda o derecha desde la barra de herramientas o el menú contextual.
- Deshacer y rehacer cambios durante la edición.
- Cálculos de selección en la barra de estado: suma, promedio y cantidad.
- Fórmulas básicas con referencias y rangos, entre ellas `SUM`, `AVERAGE`, `COUNT`, `MIN`, `MAX`, `IF`, `COUNTIF` y `SUMIF`, junto con varios alias en español.
- Autocompletado contextual de fórmulas y evaluación aritmética con paréntesis, precedencia, potencias y porcentajes.
- Ajustar mediante arrastre el ancho de las columnas y la altura de las filas; las dimensiones se guardan en el archivo XLSX.
- Guardar en el archivo tanto la expresión de cada fórmula como su último valor calculado: la cuadrícula muestra el resultado y la barra de fórmulas muestra la expresión editable.
- Recalcular fórmulas simples al editar datos y ajustar sus referencias al insertar, eliminar o mover filas y columnas dentro de la hoja activa.
- Descargar el libro como XLSX o XLS, o exportar la hoja activa como CSV.

El evaluador integrado cubre fórmulas simples y referencias dentro de la hoja activa; no pretende sustituir todavía el motor completo de Microsoft Excel. Al abrir el archivo exportado, Excel, LibreOffice u otra aplicación compatible puede recalcular las fórmulas OOXML conservadas.

### Edición de documentos Word

- El motor DOCX funciona localmente en el renderer; se carga al abrir un Word y no requiere servidor, cuentas ni servicios externos. Se incluyen únicamente `@docx-editor.dev/core`, `vue` e `i18n`, bajo Apache 2.0, y Vue bajo MIT. No incluye SuperDoc ni paquetes Pro. Los avisos de licencia se distribuyen en `assets/licenses/word-editor`.
- El guardado automático espera 700 ms después de editar, serializa las revisiones y actualiza el binario y el texto de búsqueda juntos. Cambiar de nota termina el guardado de la sesión anterior antes de liberar su motor. Un fallo conserva la sesión para reintentar y muestra un aviso global; el cierre de la ventana se bloquea mientras quedan cambios por confirmar.
- Los rectángulos VML simples anclados a página se convierten en imágenes/cuadros DrawingML editables manteniendo coordenadas, dimensiones, relaciones e imágenes. Esta preparación ocurre en memoria; abrir sin editar no reemplaza el adjunto. Las formas antiguas fuera de ese subconjunto permanecen intactas.
- El soporte no equivale a todas las funciones de Microsoft Office: SmartArt, gráficos avanzados, formas agrupadas/rotadas y algunos cuadros en encabezados o con tablas pueden conservarse sin poder editarse. La fidelidad de fuentes y funciones avanzadas debe comprobarse con documentos representativos. Consulta la [matriz del motor](https://www.docx-editor.dev/docs/2.x/word-fidelity).
- `tests/electron-word-editor.cjs` comprueba escritura real, guardado SQLite, reapertura, exportación y conservación de imágenes en un perfil aislado. Acepta una ruta DOCX opcional para QA local; no modifica el archivo de entrada ni lo añade al repositorio.

### Presentaciones

- Crear una presentación en blanco como un tipo de nota independiente.
- Importar archivos PPTX y recuperar su secuencia de diapositivas, títulos, bloques de texto y notas del presentador.
- Editar el título y contenido de cada diapositiva desde un lienzo 16:9 con miniaturas laterales.
- Agregar, duplicar, eliminar y reordenar diapositivas.
- Usar una cinta compacta con pestañas Inicio, Insertar, Diseño, Transiciones, Animaciones y Presentación.
- Aplicar formato al título o contenido: tamaño, negrita, cursiva, subrayado, alineación y viñetas.
- Elegir diseños de título/contenido, solo título o sección, además de temas de color que se conservan al exportar.
- Configurar transiciones por diapositiva —desvanecer, empuje, zoom o barrido—, ajustar su duración, previsualizarlas y aplicarlas a toda la presentación.
- Animar título, contenido o ambos con aparición, desvanecimiento, flotación o zoom, incluyendo duración y retraso independientes.
- Añadir notas del presentador que no aparecen en el modo de exposición.
- Iniciar una presentación a pantalla completa desde la diapositiva seleccionada.
- Navegar con las flechas, `Page Up`, `Page Down`, espacio, `Home` y `End`, y salir con `Esc`.
- Exportar la presentación editada como un archivo PPTX compatible con PowerPoint y LibreOffice Impress.
- Conservar el archivo PPTX original dentro de la nota importada y dentro de los respaldos de Helper Notes.

El editor interno utiliza un modelo de diapositivas enfocado en títulos, texto y notas. Al importar, el archivo original conserva temas, imágenes, gráficos, animaciones y otros elementos avanzados, pero esos elementos todavía no se reconstruyen en el editor ni en el modo de presentación de Helper Notes. La exportación genera una nueva versión editable con el contenido manejado por la aplicación. El formato binario heredado `.ppt` no está soportado; debe convertirse previamente a `.pptx`.

### Importación y exportación

| Formato | Importar | Exportar | Comportamiento |
| --- | :---: | :---: | --- |
| TXT | Sí | Sí | Convierte cada línea en un párrafo al importar y elimina el formato al exportar. |
| HTML / HTM | Sí | Sí | Usa el contenido de `body` cuando está disponible. |
| DOCX | Sí | Sí | Lee y genera OOXML con soporte para el formato principal de las notas. |
| PPTX | Sí | Sí | Permite crear y editar presentaciones, importar texto y notas por diapositiva, presentar a pantalla completa y volver a exportar. Las notas de texto también pueden convertirse automáticamente a PPTX. |
| PDF | Sí | Sí | Los PDF importados se conservan y visualizan embebidos; las notas se renderizan a PDF al exportar. |
| XLSX / XLS / CSV | Sí | XLSX / XLS / CSV | Se abren en el editor de hojas de cálculo. XLSX y XLS exportan el libro; CSV exporta la hoja activa. Una nota de texto también puede convertirse a XLSX. |
| JSON | Sí | Sí | Permite respaldar y restaurar una nota, incluidos sus datos de archivo cuando corresponda. |

La importación admite selección múltiple. Las exportaciones muestran su progreso y permiten revelar el archivo guardado en el explorador del sistema.

### Respaldos de carpetas y notas

- Pantalla dedicada para seleccionar simultáneamente varias carpetas y notas individuales.
- Exportación de una carpeta completa, incluidas todas sus notas, contenido enriquecido, hojas de cálculo, PDF y presentaciones.
- Exportación como JSON versionado para conservar la estructura interna o como ZIP con carpetas y archivos físicos: HTML, XLSX, PDF y PPTX.
- Los nombres repetidos dentro del ZIP reciben un sufijo y los caracteres incompatibles con nombres de archivo se reemplazan de forma segura.
- Importación de uno o varios respaldos JSON, ZIP físicos y ZIP antiguos con JSON interno.
- Importación desde una carpeta del sistema, buscando respaldos compatibles en sus subcarpetas.
- Los nombres de carpetas que ya existen se conservan creando una copia con sufijo, sin sobrescribir información local.
- Lectura segura de ZIP sin extracción al disco, con límites de tamaño, cantidad de archivos y profundidad de búsqueda.

El modelo de datos conserva la relación entre carpetas y subcarpetas. Tanto el respaldo JSON como el ZIP preservan esa jerarquía; el ZIP contiene directorios y archivos físicos que pueden abrirse directamente fuera de Helper Notes.

### Asistente de IA local

Cada asistente es un perfil configurable con nombre, rol, instrucciones, proveedor, modelo, temperatura y acciones disponibles. La aplicación puede conectarse a Ollama o a un `llama-server` de llama.cpp mediante HTTP.

La integración permite:

- Crear, duplicar, editar y eliminar perfiles desde la pantalla **Asistentes**.
- Asignar roles como escritor, profesor, diseñador u otro rol definido por el usuario.
- Configurar instrucciones diferentes para notas de texto, hojas de cálculo, presentaciones y PDF.
- Definir acciones, sus instrucciones, el tipo de ejecución y los archivos en los que están habilitadas.
- Detectar si Ollama está instalado, iniciar o detener su servicio y seleccionar sus modelos instalados mediante una lista desplegable.
- Instalar modelos locales recomendados desde la propia pantalla, mostrando el progreso y seleccionándolos al terminar.
- Agregar modelos de Ollama Cloud mediante inicio de sesión. El catálogo sugerido incluye `gemma4:31b-cloud`, `gpt-oss:20b-cloud`, `gpt-oss:120b-cloud`, `nemotron-3-nano:30b-cloud`, `nemotron-3-super:cloud` y `nemotron-3-ultra:cloud`; su disponibilidad y sus límites dependen del plan de Ollama.
- Instalar automáticamente el paquete CPU oficial más reciente de llama.cpp para Windows x64, incluso cuando la release estable enlaza a una publicación nocturna separada.
- Buscar modelos GGUF públicos en Hugging Face, elegir su cuantización y descargarlos desde la aplicación.
- Administrar los modelos locales e iniciar o detener `llama-server` sin salir de Helper Notes.
- Iniciar automáticamente el servidor administrado al abrir el chat cuando el perfil tiene un modelo local seleccionado.
- Conectarse también a un servidor llama.cpp externo compatible con la API de chat de OpenAI.
- Elegir entre los modelos conversacionales instalados y recordar el último modelo utilizado.
- Recibir la respuesta mediante streaming y cancelar una generación en curso.
- Enviar nuevos mensajes mientras el asistente responde; se ejecutan secuencialmente desde una cola visible.
- Mantener varias conversaciones y guardar su historial en SQLite.
- Incorporar como contexto el título y contenido de la nota abierta o los datos de la hoja de cálculo activa.
- Responder en el idioma configurado en la aplicación.
- Proponer acciones para analizar, editar la nota actual, reemplazar texto, crear una nota, crear una carpeta o modificar celdas de una hoja de cálculo.
- Mostrar una vista previa de los cambios antes de aplicar o rechazar cada acción.

Ollama y llama.cpp son opcionales: las notas, carpetas, archivos y el resto de la aplicación funcionan sin ellos. Helper Notes puede descargar llama.cpp y modelos GGUF bajo el directorio `userData/llama.cpp`, o conectarse a una instalación externa configurando otro endpoint.

### Apariencia e idiomas

- Modos claro, oscuro y sistema.
- Temas de color azul, verde, rosa, ámbar y violeta, además de un color personalizado.
- Cambio de idioma sin reiniciar la aplicación.
- Traducciones incluidas para español, inglés, chino, alemán, francés, japonés y portugués.

### Integración de escritorio

- Una sola instancia principal de la aplicación.
- Bandeja del sistema: cerrar la ventana la oculta; la opción **Salir** del tray termina el proceso.
- Menú contextual nativo para deshacer, rehacer, cortar, copiar y pegar.
- Diálogos nativos para confirmar acciones, abrir archivos y guardar exportaciones.
- Instaladores nativos: NSIS para Windows, AppImage/DEB para Linux y DMG/ZIP para macOS.

## Stack tecnológico

| Tecnología | Versión del proyecto | Uso |
| --- | --- | --- |
| Angular | 21 | Interfaz con componentes standalone, signals y effects. |
| Electron | 40 | Ventanas, bandeja del sistema, diálogos e IPC. |
| TypeScript | 5.9 | Código del frontend y del proceso principal. |
| better-sqlite3 | 12 | Persistencia local con SQLite y modo WAL. |
| SheetJS (`xlsx`) | 0.18 | Lectura, edición y escritura de hojas de cálculo. |
| Tailwind CSS | 3 | Utilidades de estilos y temas. |
| ngx-translate | 17 | Internacionalización. |
| electron-builder | 26 | Empaquetado NSIS, AppImage, DEB, DMG y ZIP. |
| Vitest | 4 | Pruebas unitarias de Angular. |

La comunicación entre Angular y Electron se realiza a través de un `preload` con `contextIsolation` habilitado y una API IPC explícita; el renderer no tiene acceso directo a Node.js.

## Almacenamiento local

La base de datos se crea automáticamente como `helpernote.db` dentro del directorio `userData` asignado por Electron y utiliza el modo WAL de SQLite.

| Tabla | Contenido |
| --- | --- |
| `folders` | Carpetas y fecha de eliminación suave. |
| `notes` | Título, HTML, carpeta, archivos embebidos y fecha de eliminación suave. |
| `settings` | Idioma, tema, color y otras preferencias. |
| `ai_conversations` | Sesiones, modelo y perfil de asistente seleccionado. |
| `ai_messages` | Mensajes asociados a cada conversación. |

Las tablas y migraciones necesarias se ejecutan al iniciar la aplicación. Los PDF y libros de cálculo importados se almacenan codificados en la propia base de datos.

## Actualizaciones automáticas

La aplicación consulta la API de versiones después de iniciarse y luego cada seis horas. La consulta incluye la versión instalada, el sistema (`win32`, `darwin` o `linux`), la arquitectura y el canal. Cuando existe una versión superior publicada:

- descarga el instalador desde la URL registrada;
- valida su hash SHA-256 cuando está disponible;
- solicita instalarlo; ejecuta NSIS silenciosamente en Windows, abre DMG en macOS y puede iniciar una AppImage descargada en Linux;
- expone controles IPC para comprobar, descargar o instalar manualmente.

Copie `.env.example` como `.env` para desarrollo o defina estas variables en el entorno del proceso:

| Variable | Valor mock | Descripción |
| --- | --- | --- |
| `HELPER_NOTES_UPDATE_API_URL` | `http://127.0.0.1:3000` | URL base de la API. |
| `HELPER_NOTES_UPDATE_CHANNEL` | `stable` | Canal de publicación. |
| `HELPER_NOTES_AUTO_UPDATE_ENABLED` | `true` | Activa las comprobaciones periódicas. |
| `HELPER_NOTES_AUTO_DOWNLOAD_UPDATES` | `true` | Descarga automáticamente en builds empaquetados. |
| `HELPER_NOTES_UPDATE_INITIAL_DELAY_MS` | `5000` | Espera inicial. |
| `HELPER_NOTES_UPDATE_INTERVAL_MS` | `21600000` | Intervalo entre comprobaciones. |

El número publicado debe ser superior al campo `version` de `package.json`. Antes de distribuir una compilación nueva, actualice ese campo y genere el instalador correspondiente a la plataforma.

## Requisitos

- Windows para generar NSIS.
- Linux para generar AppImage y DEB.
- macOS para generar DMG y ZIP. La firma y notarización requieren certificados de Apple; las compilaciones locales y el workflow quedan sin firmar si no se configuran credenciales.
- Node.js 22 y npm.

Los paquetes se construyen en su sistema destino porque `better-sqlite3` contiene código nativo. No se debe generar el paquete de macOS o Linux desde Windows.
- [Node.js](https://nodejs.org/) `20.19` o superior compatible con Angular 21 (`22.12+` o `24+` también son compatibles).
- npm 10 o superior.
- [Ollama](https://ollama.com/download) con un modelo conversacional, o un `llama-server` de llama.cpp accesible por HTTP, únicamente para usar el asistente de IA.

## Desarrollo

Desde el directorio `helper-notes-app-desktop`:

```bash
npm install
npm run electron:dev
```

Este comando inicia Angular en `http://localhost:4200`, observa la compilación del proceso principal y abre Electron cuando el servidor está disponible.

Otros comandos útiles:

| Comando | Descripción |
| --- | --- |
| `npm start` | Inicia solamente el servidor de desarrollo de Angular. Las APIs de Electron no estarán disponibles en un navegador normal. |
| `npm run build` | Genera el build de producción de Angular en `dist/`. |
| `npm run electron:build` | Compila el proceso principal y el preload de Electron en `electron-dist/`. |
| `npm run electron:run` | Ejecuta el proceso de Electron ya compilado. |
| `npm run electron` | Compila Angular y Electron, y luego abre la aplicación. |
| `npm run build:electron` | Compila y empaqueta para la plataforma y arquitectura del equipo actual. |
| `npm run dist:win` | Genera el instalador de Windows en `release/`. |
| `npm run dist:linux` | Genera AppImage y DEB para la arquitectura del equipo Linux. |
| `npm run dist:linux:x64` | Genera AppImage y DEB para Linux x64. |
| `npm run dist:linux:arm64` | Genera AppImage y DEB para Linux ARM64. |
| `npm run dist:mac` | Genera DMG y ZIP para la arquitectura del Mac. |
| `npm run dist:mac:x64` | Genera DMG y ZIP para macOS Intel. |
| `npm run dist:mac:arm64` | Genera DMG y ZIP para Apple Silicon. |
| `npm test -- --watch=false` | Ejecuta las pruebas una vez. |

Las salidas se separan en `release/linux-<arquitectura>/` y `release/mac-<arquitectura>/`. El workflow `.github/workflows/build-desktop.yml` ejecuta estas compilaciones en runners nativos al publicar una etiqueta `v*` o manualmente desde GitHub Actions.

La compilación de Windows usa el paquete RCEdit exclusivo de Windows de `win-codesign@1.1.0` para conservar el icono y las versiones del ejecutable. El hook `scripts/edit-windows-executable.js` reemplaza la edición heredada de electron-builder, evitando extraer los enlaces simbólicos de macOS incluidos en `winCodeSign-2.6.0`. El script no cierra otras aplicaciones y ejecuta las herramientas locales con el mismo runtime de Node que inició la compilación.
| `npx ng test` | Ejecuta las pruebas unitarias con Vitest. |

## Estructura del proyecto

```text
helper-notes-app-desktop/
├── assets/                    # Iconos de la aplicación y del instalador
├── build/                     # Entitlements y recursos generados para macOS
├── electron/
│   ├── db/                    # Inicialización de SQLite y repositorios
│   ├── ipc-handlers.ts        # Archivos, ventanas, hojas de cálculo y Ollama
│   ├── main.ts                # Ciclo de vida, ventana principal y tray
│   ├── update-manager.ts      # Descarga, verificación e instalación de versiones
│   └── preload.ts             # API segura expuesta a Angular
├── public/
│   └── i18n/                  # Traducciones JSON
├── scripts/                   # Utilidades de compilación por plataforma
├── src/
│   └── app/
│       ├── components/        # Sidebar, lista, editor, IA y configuración
│       ├── models/            # Modelos de carpetas y notas
│       ├── services/          # Estado, persistencia, temas, idioma e IA
│       └── types/             # Tipos de la API expuesta por Electron
├── angular.json
└── package.json
```

## Edición de hojas de cálculo

El editor permite combinar/separar celdas (conserva los valores cubiertos), combinar y centrar, ajustar texto, cambiar colores y formatos numéricos, insertar/eliminar/mover filas y columnas, ordenar filas completas, copiar fórmulas con referencias relativas/absolutas y deshacer/rehacer. Los cambios de formato se guardan en el XLSX y sobreviven al reabrirlo.

El parser local admite operadores, funciones anidadas, referencias entre hojas, detección de ciclos y funciones de agregación, lógica, texto, criterios y búsqueda (`SUM`, `AVERAGE`, `IF`, `IFERROR`, `COUNTIF`, `SUMIF`, `VLOOKUP`, `INDEX`, `MATCH`, entre otras). Acepta alias españoles y punto y coma; guarda nombres ingleses y comas de OOXML. Conserva el resultado importado de funciones desconocidas y solicita recálculo al abrir en Excel/Calc; no calcula todas las funciones de Office. No implementa macros ni todas las funciones avanzadas de una suite como FreeOffice.

`electron/spreadsheet-edit.ts` modifica datos, combinaciones y estilos dentro del ZIP original, conservando los recursos y relaciones restantes. `src/app/models/spreadsheet-formula.ts` contiene el cálculo seguro, sin ejecutar JavaScript de las fórmulas. `tests/electron-spreadsheet.cjs` verifica los controles reales y la reapertura con SQLite en un perfil aislado; no utiliza documentos ni portapapeles del usuario.

### Fórmulas y compatibilidad ampliadas

El catálogo de cálculo local incluye funciones matriciales, además de SUMIFS/COUNTIFS/AVERAGEIFS, AVERAGEIF, MINIFS/MAXIFS, SUMPRODUCT, XLOOKUP/XMATCH, IFNA/IFS/SWITCH, inspección de errores/tipos, estadísticas muestrales/poblacionales y más operaciones de texto. Se aceptan alias españoles (por ejemplo, `SUMAR.SI.CONJUNTO`, `BUSCARX`, `SI.CONJUNTO` y `UNIRCADENAS`) y nombres ingleses; las sugerencias incluyen firmas y descripciones. Las condiciones admiten comodines `*`/`?` y escape `~`; la igualdad normal de operadores no trata el texto como patrón. Las búsquedas aproximadas de MATCH/VLOOKUP y la precedencia de signos/potencias siguen las [reglas de cálculo de Excel](https://support.microsoft.com/es-es/excel/calculation-operators-and-precedence-in-excel).

DATE, YEAR/MONTH/DAY, TODAY/NOW, DAYS, EDATE/EOMONTH, WEEKDAY, TIME/HOUR/MINUTE/SECOND y WORKDAY/NETWORKDAYS calculan seriales con la [época 1900 o 1904 del libro](https://support.microsoft.com/en-us/excel/date-systems-in-excel), incluyendo el día ficticio 29/02/1900. Las fechas y horas conservan su formato original; para una fórmula nueva se puede elegir el formato de fecha de la barra. WORKDAY y NETWORKDAYS usan lunes–viernes y feriados opcionales, con un límite de 100000 días recorridos.

Se resuelven nombres definidos simples de celdas/rangos A1, constantes y expresiones escalares, con prioridad de nombres locales sobre globales y detección de ciclos. Se conserva su definición en el archivo. Las funciones modernas conservan o reciben prefijos OOXML `_xlfn.` al guardar. Las tablas dinámicas conservan el último resultado y muestran un aviso si una medida numérica de su origen contiene errores. Los errores de fórmulas se exportan como errores nativos; los errores importados sin fórmula no se convierten en texto. Las fórmulas matriciales no soportadas conservan sus atributos si la expresión no cambia.

Una fórmula importada que el parser o el catálogo no pueda interpretar conserva su expresión y su resultado importado, y muestra un aviso de posible desactualización. Una expresión nueva o modificada no reutiliza el resultado de otra fórmula. El archivo solicita recálculo completo en Excel/Calc al abrirlo.

El motor sigue siendo un subconjunto: no calcula vínculos externos, fórmulas 3D ni macros. XLOOKUP puede devolver una fila o columna de resultados; XLOOKUP/XMATCH calculan búsquedas lineales de principio a fin o al revés (modos 1/-1), con coincidencia exacta, aproximada o comodines; no implementan búsqueda binaria. Los rangos evaluados se limitan a 100000 celdas. Los nombres importados no tienen todavía un administrador de edición ni traducción de referencias relativas/estructuradas; sus definiciones se conservan tal como se importan. No se garantiza equivalencia con todas las reglas regionales o funciones de Office.

La lógica está separada entre `spreadsheet-formula.ts` (parser/dependencias por referencia), `spreadsheet-function-library.ts` (funciones adicionales), `spreadsheet-date.ts` (seriales y calendario) y `spreadsheet-formula-catalog.ts` (ayuda). Las pruebas de modelos y `tests/spreadsheet-compatibility.test.ts` cubren cálculo y conservación del paquete; `tests/electron-formula.cjs` verifica controles reales, fórmulas localizadas, nombres locales/globales, época 1904, errores nativos, cachés y reapertura en SQLite aislado.

### Matrices dinámicas, tablas y rangos completos

`FILTRAR/FILTER`, `ORDENAR/SORT`, `UNICOS/UNIQUE`, `SECUENCIA/SEQUENCE` y `TRANSPONER/TRANSPOSE` devuelven matrices desde una celda ancla. Los operadores comparan y calculan rangos por elemento; FILTER acepta máscaras combinadas y un resultado alternativo si no hay coincidencias. IF/IFERROR y XLOOKUP admiten resultados matriciales en el subconjunto implementado. `A1#` referencia el desbordamiento completo y se guarda mediante `_xlfn.ANCHORARRAY`.

El cálculo por dependencias protege celdas ocupadas, combinaciones y tablas. Si el resultado no cabe, el ancla muestra `#SPILL!` y no reemplaza los datos existentes; al quitar el obstáculo se recalcula. Los resultados derivados están marcados y protegidos: se modifica la fórmula de origen. Reducir o eliminar la fórmula limpia sus antiguos resultados, y deshacer/rehacer restaura expresión y estructura. Los resultados de funciones dinámicas externas aún no soportadas se conservan completos con un aviso.

Las tablas nativas importadas y nuevas se resuelven por nombre y columnas: `Ventas[Importe]`, `Ventas[[Importe]:[Unidades]]`, `[@Importe]`, `#Data`, `#Headers`, `#Totals`, `#All` y `#This Row`. **Crear tabla** usa la selección como rango inicial; **Redimensionar tabla** permite editar nombre, tamaño y fila de totales. Hay suma, promedio, conteo de valores/números, mínimo, máximo y desviación/varianza muestral o poblacional por columna. El rango incluye encabezados y totales. Los destinos ocupados para una nueva fila de totales se rechazan antes de modificar datos.

La creación genera partes, relaciones, estilos y tipos de contenido nativos XLSX. Renombrar tablas o encabezados actualiza referencias explícitas e implícitas en fórmulas y nombres definidos, sin modificar literales. Los identificadores estables de tabla/columna conservan fórmulas calculadas y filtros válidos; las fórmulas de columna también se guardan como metadatos nativos. Se pueden insertar/eliminar filas de datos y columnas, mover bloques completos y ordenar los datos de una tabla sin mover sus encabezados, totales ni celdas externas. Las filas calculadas nuevas vacías heredan su fórmula. Reducir el rango conserva las celdas externas; desplazar una fila de totales limpia sólo su posición anterior. Se rechazan tablas partidas, eliminación parcial de encabezados y pérdida de todas las filas de datos. Para quitar totales se usa su casilla. Se admiten hasta 100.000 celdas por tabla; las uniones estructuradas no contiguas y varios selectores combinados siguen pendientes.

Las hojas con matrices calculables permiten insertar, eliminar y mover filas o columnas: se limpian sólo resultados derivados, se remapean fórmulas/nombres y se regeneran las matrices. Eliminar un ancla elimina su matriz; encontrar datos en el desbordamiento produce `#SPILL!` sin sobrescribirlos. Ordenar filas que atraviesan una matriz se rechaza; usa `ORDENAR` o modifica su origen. Las matrices importadas con funciones no calculables conservan su protección estructural.

Al borrar extremos de rangos, sus referencias se contraen a las celdas supervivientes, preservando signos absolutos y nombres de hoja; un rango eliminado completamente produce `#REF!`. Esto incluye columnas/filas completas, que mantienen independencia del otro eje.

Se añaden 20 funciones: `SUBTOTAL`, `PMT`, `PV`, `FV`, `NPER`, `IPMT`, `PPMT`, `NPV`, `MROUND`, `LN`, `LOG10`, `LOG`, `EXP`, `FACT`, `GCD`, `LCM`, `T`, `NUMBERVALUE`, `TEXTBEFORE` y `TEXTAFTER`, con alias españoles y ayuda contextual. Las funciones modernas mantienen sus prefijos OOXML y las alternativas de texto se evalúan sólo cuando hacen falta. `SUBTOTAL` implementa los códigos 1–11/101–111 y evita subtotales anidados; los códigos 101–111 usan las filas ocultas importadas. Con filtros activos, el archivo no distingue siempre ocultación manual/filtrada: los códigos 1–11 conservan el resultado importado y muestran el aviso de cálculo pendiente. Este editor aún no reaplica filtros avanzados ni incorpora todas las funciones de Office. La semántica de subtotales sigue la [documentación de Microsoft](https://support.microsoft.com/en-us/excel/functions/subtotal-function).

Las cachés de tablas dinámicas importadas cuyo origen es el nombre de una tabla conservan esa vinculación al renombrarla; se solicita su actualización al abrir en Excel. El cálculo local de esos orígenes por nombre sigue pendiente; las tablas dinámicas creadas por Helper Notes mantienen su actualización por rango.

`tests/electron-tables.cjs` verifica creación, nombres, encabezados, IDs, totales, desplazamiento de totales, columnas calculadas, inserciones, eliminación de columnas, ordenación sin afectar celdas externas, protección de encabezados, deshacer y reapertura en una base SQLite aislada.


El guardado escribe `f t="array"`, los rangos resultantes y metadata nativa XLDAPR, reutilizando sus registros y conservando metadatos ajenos. No necesita servicios externos. Los detalles siguen la [documentación oficial de matrices y desbordamientos de Excel](https://support.microsoft.com/en-US/Excel/dynamic-array-formulas-and-spilled-array-behavior) y la estructura [DynamicArrayProperties de Open XML](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.office2019.excel.dynamicarray.dynamicarrayproperties?view=openxml-3.0.1).

`spreadsheet-array.ts` administra matrices y rangos compactos, `spreadsheet-workbook.ts` dependencias/propietarios, `spreadsheet-table.ts` referencias estructuradas y `electron/spreadsheet-structure.ts` la estructura nativa. Las pruebas incluyen un XLSX sintético de XlsxWriter como productor independiente (`tests/fixtures/dynamic-source.xlsx`), verificación del XML y `tests/electron-dynamic.cjs` para editar, bloquear, desbloquear, reducir y reabrir resultados en un perfil aislado.

### Tablas dinámicas

El botón **Tabla dinámica** abre el configurador sobre el rango seleccionado (primera fila = encabezados). Con una sola celda seleccionada se propone el rango usado de la hoja. Permite hasta cinco niveles de filas, tres de columnas y ocho valores, con suma, conteo de valores, conteo numérico, promedio, mínimo y máximo. Los filtros comparan un valor tipado por campo. Los totales se calculan desde los registros originales, incluyendo los promedios, y se puede consultar una vista previa antes de guardar.

Cada tabla obtiene una hoja de resultado propia. **Editar campos** modifica su configuración sin duplicarla; **Actualizar tablas dinámicas** fuerza el recálculo, que también ocurre al guardar cambios en el origen. El resultado calculado se protege contra edición directa de valores. Si el origen queda inválido, se conserva el último resultado y se muestra un aviso; el rango de origen sigue siendo el rango A1 explícito configurado y puede ajustarse desde el panel.

`electron/spreadsheet-pivot.ts` escribe las partes `pivotTableDefinition`, `pivotCacheDefinition` y `pivotCacheRecords` y sus relaciones, siguiendo la [estructura de tablas dinámicas de SpreadsheetML](https://learn.microsoft.com/en-us/office/open-xml/spreadsheet/working-with-pivottables). La exportación XLSX conserva el paquete completo; XLS y CSV son conversiones de valores y no conservan tablas dinámicas nativas. Las tablas creadas aquí pueden reabrirse y editarse; las tablas externas con configuraciones avanzadas se conservan, pero su edición aún no está implementada. No hay agrupación de fechas, campos calculados, segmentadores, DAX ni orígenes OLAP.

Las pruebas de cálculo/caché están en `tests/spreadsheet-pivot.test.ts`; `tests/electron-pivot.cjs` comprueba el flujo real con SQLite aislado, la actualización del origen, la reapertura y la conservación de las partes nativas al exportar.

### Gráficos

**Insertar gráfico** propone el rango seleccionado; con una sola celda usa el área ocupada. El configurador permite columnas, barras, líneas, áreas, circular, anillo y dispersión XY, categorías, hasta doce series, títulos de ejes, colores, leyenda, etiquetas y series apiladas en columnas/barras/áreas. Una vista previa permite comprobar la selección antes de guardar. Circular y anillo usan una sola serie no negativa; dispersión requiere X numérico. Los huecos se conservan y las cadenas no se convierten silenciosamente en números.

Los gráficos aparecen sobre la hoja: arrastrar el encabezado los mueve; arrastrar la esquina inferior cambia su tamaño. El botón de edición permite introducir ancho/alto y modificar la configuración; eliminar un gráfico conserva sus celdas. Creación, edición, posición y eliminación participan en deshacer/rehacer. Se actualizan al guardar cambios de datos, fórmulas o tablas dinámicas.

`electron/spreadsheet-chart.ts` escribe partes nativas DrawingML/ChartML con referencias absolutas y cachés, según la [estructura de gráficos de SpreadsheetML](https://learn.microsoft.com/en-us/office/open-xml/spreadsheet/how-to-insert-a-chart-into-a-spreadsheet). Conserva las imágenes y dibujos restantes. XLSX exporta el paquete original completo; XLS/CSV no conservan gráficos. El renderizador SVG local no requiere servicios externos ni nuevas licencias de motor.

Se pueden leer y editar gráficos básicos externos con series de columnas contiguas en filas y un único origen. El guardado ordinario conserva su XML original; editar explícitamente uno aplica las opciones soportadas por este editor. Combinados, 3D, ejes secundarios y otros tipos avanzados se conservan dentro del archivo, pero no se renderizan ni editan todavía. El rango A1 es explícito: si cambia la estructura o crece la tabla, debe ajustarse en el configurador. Límites: 50 gráficos editables por libro, 3000 filas por rango y 12000 puntos por gráfico; se simplifican etiquetas/leyenda en vistas densas para evitar superposición.

`tests/spreadsheet-chart.test.ts` verifica XML, referencias, cachés, validación y conservación de tablas dinámicas. `spreadsheet-chart-view.spec.ts` verifica negativos, huecos, anillos, apilados y escala XY. `tests/electron-chart.cjs` comprueba los controles reales, arrastre, tamaño, actualización, reapertura, eliminación y exportación con un perfil SQLite aislado. Los siete tipos se verificaron además con lectura y reescritura mediante openpyxl.

## Autor y licencia

Creado por **Schavez**.

Proyecto privado. Todos los derechos reservados.
