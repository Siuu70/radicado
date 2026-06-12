# 📋 Sistema Automatizado de Sello de Radicado — Guía para Claude Code

## ¿Qué hace este sistema?

Toma un documento (Word o PDF), lee los datos de un Excel (número de radicado, fecha, recibido por, asunto, anexos) y estampa un **sello de radicado visual** con el logo de la empresa directamente sobre el documento.

---

## 📁 Estructura de archivos que debes tener listos

```
proyecto-radicado/
├── logo.png                  ← Logo de Serventegral (tú lo provees)
├── radicados.xlsx            ← Excel con los datos de radicado
├── documentos/               ← Carpeta con los docs a radicar
│   ├── carta_proveedor.docx
│   └── solicitud.pdf
├── salida/                   ← Aquí quedan los documentos sellados
└── sello_radicado.js         ← Script principal (Claude Code lo genera)
```

---

## 📊 Estructura del Excel (`radicados.xlsx`)

El Excel debe tener estas columnas exactas en la **primera hoja**:

| NUMERO_RADICADO | RECIBIDO_POR | ASUNTO | ANEXOS | FECHA_RECIBIDO | ARCHIVO_DOCUMENTO |
|----------------|--------------|--------|--------|----------------|-------------------|
| 0914 | YUDY DELGADO | Solicitud de servicio | 3 folios | 03/06/2026 | carta_proveedor.docx |
| 0915 | CARLOS RUIZ | Cotización | Ninguno | 04/06/2026 | solicitud.pdf |

> **Importante:** La columna `ARCHIVO_DOCUMENTO` debe tener el nombre exacto del archivo que está en la carpeta `documentos/`.

---

## 🛠️ Paso a paso en Claude Code

### PASO 1 — Dile a Claude Code qué necesitas

Abre Claude Code y escribe exactamente esto (ajusta las rutas):

```
Necesito un script en Node.js que haga lo siguiente:

1. Lea el archivo radicados.xlsx y extraiga los datos de cada fila
2. Para cada fila, abra el documento Word (.docx) indicado en la columna ARCHIVO_DOCUMENTO
3. Inserte un sello de radicado visual al final del documento con estos campos:
   - Logo de la empresa (archivo logo.png)
   - Nombre de la empresa: Serventegral S.A. E.S.P.
   - NIT: 828.002.229-2
   - Recibido por: [RECIBIDO_POR]
   - Asunto: [ASUNTO]
   - Anexos: [ANEXOS]
   - Número de radicado: [NUMERO_RADICADO]
   - Fecha de recibido: [FECHA_RECIBIDO]
4. Guarde el documento modificado en la carpeta salida/ con el mismo nombre
5. El sello debe verse como un recuadro con borde, similar a un sello oficial

Usa las librerías: xlsx, docx, fs
Estructura de carpetas:
- logo.png está en la raíz del proyecto
- Los documentos están en documentos/
- La salida va en salida/
```

---

### PASO 2 — Instalar dependencias

Claude Code ejecutará esto (o puedes hacerlo tú):

```bash
npm init -y
npm install xlsx docx fs-extra
```

---

### PASO 3 — Proveer el logo

Coloca tu archivo `logo.png` en la carpeta raíz del proyecto.

> Si el logo está en otro formato (JPG, SVG), dile a Claude Code:
> *"El logo está en formato JPG, ajusta el script para usar logo.jpg"*

---

### PASO 4 — Ejecutar el script

```bash
node sello_radicado.js
```

El script procesará **todas las filas del Excel** automáticamente y generará un documento sellado por cada una en la carpeta `salida/`.

---

### PASO 5 — Verificar el resultado

Abre uno de los archivos de la carpeta `salida/` y verifica que el sello aparezca correctamente con todos los campos.

---

## 🎨 Diseño del sello

El sello tendrá este aspecto dentro del documento:

```
┌─────────────────────────────────────────────────┐
│  [LOGO]   SERVENTEGRAL S.A. E.S.P.              │
│           NIT: 828.002.229-2                     │
│  ─────────────────────────────────────────────  │
│  RECIBIDO POR:  YUDY DELGADO                    │
│  ASUNTO:        Solicitud de servicio           │
│  ANEXOS:        3 folios                        │
│  N° RADICADO:   0914                            │
│  FECHA:         03/06/2026                      │
└─────────────────────────────────────────────────┘
```

---

## 🔧 Variaciones útiles — dile a Claude Code

### Si quieres procesar UN solo documento manualmente:
```
Modifica el script para que acepte argumentos por línea de comandos:
node sello_radicado.js --radicado 0914 --archivo carta.docx
```

### Si los documentos son PDF en vez de Word:
```
Adapta el script para documentos PDF usando la librería pdf-lib
en lugar de docx
```

### Si quieres una interfaz visual (formulario web):
```
Crea una interfaz web simple en HTML donde pueda:
- Subir el Excel
- Subir el documento
- Hacer clic en "Generar sello"
- Descargar el documento sellado
```

### Si quieres radicado automático (número consecutivo):
```
Agrega lógica para que el número de radicado se genere 
automáticamente de forma consecutiva, guardando el último 
número en un archivo radicado_counter.json
```

---

## ⚠️ Posibles errores y cómo resolverlos

| Error | Causa | Solución |
|-------|-------|----------|
| `Cannot find module 'xlsx'` | Falta instalar dependencias | Ejecuta `npm install xlsx docx` |
| `ENOENT: no such file` | El archivo del documento no existe en `documentos/` | Verifica que el nombre en el Excel coincida exactamente |
| Logo no aparece | Ruta del logo incorrecta | Verifica que `logo.png` esté en la raíz del proyecto |
| El sello aparece en página en blanco | Comportamiento normal en Word | El sello va al final, puede crear nueva página |

---

## 📌 Prompt completo listo para copiar en Claude Code

Copia y pega esto directamente:

```
Tengo un proyecto para automatizar sellos de radicado en documentos Word.

Crea un script Node.js completo llamado sello_radicado.js que:

1. Lea radicados.xlsx (columnas: NUMERO_RADICADO, RECIBIDO_POR, ASUNTO, ANEXOS, FECHA_RECIBIDO, ARCHIVO_DOCUMENTO)
2. Por cada fila, abra el .docx de la carpeta documentos/
3. Agregue al final del documento una tabla-sello con:
   - Imagen logo.png (esquina superior izquierda, aprox 2cm)
   - Texto "SERVENTEGRAL S.A. E.S.P." en negrita
   - NIT: 828.002.229-2
   - Línea separadora
   - Campo "RECIBIDO POR:" con valor de la columna RECIBIDO_POR
   - Campo "ASUNTO:" con valor de la columna ASUNTO  
   - Campo "ANEXOS:" con valor de la columna ANEXOS
   - Campo "N° RADICADO:" con valor de la columna NUMERO_RADICADO (en negrita y grande)
   - Campo "FECHA:" con valor de la columna FECHA_RECIBIDO
4. Guarde el resultado en salida/[nombre_original]_radicado.docx
5. Muestre en consola: "✓ Procesado: [archivo] — Radicado [número]"

Usa: npm xlsx, npm docx
La tabla del sello debe tener borde visible y fondo levemente gris en los labels.
```

---

## 📞 Información del sello (referencia del logo)

Basado en el sello de Serventegral:
- **Empresa:** Serventegral S.A. E.S.P.
- **NIT:** 828.002.229-2
- **Slogan:** *"Se quiere a tu ciudad limpia"*
- **Colores sugeridos:** Verde (#2E7D32) y Rojo (#C62828) — colores del logo

---

*Guía generada para implementación en Claude Code — Sistema de Radicado Automatizado*
