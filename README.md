# Helper Notes

Helper Notes es una aplicación de escritorio para crear y organizar notas por carpetas. Su experiencia toma como referencia la aplicación Notas de macOS —barra lateral, lista de notas y editor—, pero amplía ese concepto con documentos enriquecidos, hojas de cálculo, importación y exportación de archivos y un asistente de IA local.

Está desarrollada con Angular, Electron y SQLite. La información se guarda localmente en el equipo y el asistente funciona mediante Ollama, sin depender de un servicio de IA en la nube.

> Helper Notes es un proyecto independiente y no está afiliado a Apple.

## Diferencias frente a Notas de macOS

- Incluye un editor de hojas de cálculo dentro de la propia aplicación.
- Importa archivos TXT, HTML, DOCX, PDF, XLSX, XLS, CSV y PPTX, además de respaldos JSON.
- Exporta notas como TXT, HTML, DOCX, PPTX, PDF, XLSX o JSON.
- Integra modelos locales de Ollama que pueden conversar sobre la nota abierta y proponer cambios.
- Permite aplicar o rechazar individualmente las acciones sugeridas por la IA.
- Usa SQLite local en lugar de iCloud. Actualmente no incluye sincronización entre dispositivos ni colaboración en línea.
- Su distribución empaquetada actual está orientada a Windows mediante un instalador NSIS.

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
- Instalador de Windows con selección del directorio de instalación.

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
| electron-builder | 26 | Empaquetado e instalador NSIS. |
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
- solicita instalarlo y, en Windows, ejecuta silenciosamente el instalador NSIS;
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

El número publicado debe ser superior al campo `version` de `package.json`. Antes de distribuir una compilación nueva, actualice ese campo y genere el instalador con `npm run dist:win`.

## Requisitos

- Windows para generar y utilizar el instalador NSIS configurado por el proyecto.
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
| `npm run build:electron` | Compila Angular y Electron y empaqueta la aplicación con electron-builder. |
| `npm run dist:win` | Genera el instalador de Windows en `release/`. |
| `npx ng test` | Ejecuta las pruebas unitarias con Vitest. |

## Estructura del proyecto

```text
helper-notes-app-desktop/
├── assets/                    # Iconos de la aplicación y del instalador
├── electron/
│   ├── db/                    # Inicialización de SQLite y repositorios
│   ├── ipc-handlers.ts        # Archivos, ventanas, hojas de cálculo y Ollama
│   ├── main.ts                # Ciclo de vida, ventana principal y tray
│   ├── update-manager.ts      # Descarga, verificación e instalación de versiones
│   └── preload.ts             # API segura expuesta a Angular
├── public/
│   └── i18n/                  # Traducciones JSON
├── scripts/                   # Utilidades de compilación para Windows
├── src/
│   └── app/
│       ├── components/        # Sidebar, lista, editor, IA y configuración
│       ├── models/            # Modelos de carpetas y notas
│       ├── services/          # Estado, persistencia, temas, idioma e IA
│       └── types/             # Tipos de la API expuesta por Electron
├── angular.json
└── package.json
```

## Autor y licencia

Creado por **Schavez**.

Proyecto privado. Todos los derechos reservados.
