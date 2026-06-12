SISTEMA DE RADICADO - SERVENTEGRAL S.A. E.S.P
═══════════════════════════════════════════════

ANTES DE EMPEZAR:
  1. Pon el logo en:              logo.png         (raíz del proyecto)
  2. Pon los documentos Word en:  documentos/
  3. Llena el Excel con tus datos: radicados.xlsx
     (usa radicados_plantilla.xlsx como guía de columnas)

COLUMNAS DEL EXCEL (radicados.xlsx):
  NUMERO_RADICADO   → número del radicado (ej. 0001)
  FECHA_RECIBIDO    → fecha en formato DD/MM/AAAA
  RECIBIDO_POR      → nombre de quien recibe
  ENTREGADO_POR     → nombre de quien entrega
  ASUNTO            → descripción del asunto
  ANEXOS            → número o descripción de anexos
  AREA_RESPONSABLE  → área a la que va dirigido
  ARCHIVO_DOCUMENTO → nombre del archivo .docx en documentos/

PARA PREVISUALIZAR EL SELLO:
  - Abre preview.html en Chrome o Edge
  - Llena los campos del formulario
  - Ajusta la posición con los controles
  - Haz clic en "Guardar config.json" para guardar la posición

PARA PROCESAR DOCUMENTOS:
  - Abre una terminal en esta carpeta
  - Ejecuta:  node sello_radicado.js
  - Los documentos sellados quedan en:  salida/
  - El preview del sello queda en:      salida/preview_sello.png

ARCHIVOS DEL PROYECTO:
  sello_radicado.js       → script principal
  preview.html            → app visual de preview
  config.json             → posición del sello (se genera con preview.html)
  radicados.xlsx          → datos a procesar
  radicados_plantilla.xlsx → plantilla de ejemplo con columnas correctas
  logo.png                → logo de la empresa
  documentos/             → documentos Word originales
  salida/                 → documentos procesados (se crea automáticamente)
