# Diseño: Voest Paquets — App de Gestión de Packing Lists

**Fecha:** 2026-05-15  
**Estado:** Aprobado

---

## Objetivo

Aplicación Python con interfaz Gradio para cargar PDFs de packing lists (siempre con el mismo formato de Brausa), extraer sus datos mediante un LLM de OpenRouter, almacenarlos en una base de datos SQLite local, organizarlos por proyectos y exportarlos a Excel.

---

## Formato del Packing List

El PDF de entrada tiene siempre el siguiente formato (Brausa):

**Cabecera del documento:**
- Fecha del documento, Número OF (ej: OF26004215)
- N.Pedido, Cliente (código + nombre), Artículo (código + descripción)
- Ref.Pedido, Desarrollo, Espesor, Calidad, Nota
- Total Piezas, Kilos teóricos

**Cuerpo — Paquetes:**
- PAQUETE N/Total con sus KILOS
- Filas: Línea, Piezas, Longitud, Marca

---

## Arquitectura

Módulos separados por responsabilidad:

```
voest_paquets/
  app.py          ← UI Gradio (página única + sidebar)
  database.py     ← Operaciones SQLite
  extractor.py    ← LLM OpenRouter + conversión PDF a imagen
  exporter.py     ← Generación de Excel con openpyxl/pandas
  config.py       ← Constantes: API key, modelos, paths
  OPEN_KEY.txt    ← API key (gitignored)
  data/
    paquets.db    ← Base de datos SQLite
```

---

## Base de Datos (SQLite)

### Tabla `projects`
| Columna     | Tipo    | Notas              |
|-------------|---------|---------------------|
| id          | INTEGER | PK autoincrement    |
| name        | TEXT    | UNIQUE NOT NULL     |
| created_at  | TEXT    | ISO 8601            |

### Tabla `packing_lists`
| Columna       | Tipo    | Notas                              |
|---------------|---------|-------------------------------------|
| id            | INTEGER | PK autoincrement                   |
| project_id    | INTEGER | FK → projects.id                   |
| of_number     | TEXT    | Número OF (ej: OF26004215)         |
| n_pedido      | TEXT    |                                    |
| cliente       | TEXT    | Código + nombre concatenados       |
| articulo      | TEXT    | Código + descripción concatenados  |
| ref_pedido    | TEXT    |                                    |
| desarrollo    | TEXT    |                                    |
| espesor       | TEXT    |                                    |
| calidad       | TEXT    |                                    |
| nota          | TEXT    |                                    |
| total_piezas  | INTEGER |                                    |
| kilos_teoricos| REAL    |                                    |
| fecha_doc     | TEXT    | Fecha del documento PDF            |
| imported_at   | TEXT    | ISO 8601, fecha de importación     |

**Constraint:** `UNIQUE(project_id, of_number)` — para detectar y reemplazar duplicados.

### Tabla `lines`
| Columna          | Tipo    | Notas                   |
|------------------|---------|--------------------------|
| id               | INTEGER | PK autoincrement        |
| packing_list_id  | INTEGER | FK → packing_lists.id   |
| paquete_num      | INTEGER | Número de paquete       |
| kilos_paquete    | REAL    | Kilos del paquete       |
| linea            | INTEGER |                         |
| piezas           | INTEGER |                         |
| longitud         | REAL    |                         |
| marca            | TEXT    |                         |

---

## Flujo de Extracción LLM

1. Usuario sube PDF → `extractor.py` convierte cada página a imagen PNG en base64 (via `pdf2image`, DPI=150)
2. Imágenes + prompt estructurado → OpenRouter (modelo configurable, visión)
3. El prompt solicita respuesta en JSON con el esquema exacto de la BD
4. `extractor.py` valida y parsea el JSON
5. Si `of_number` ya existe en el proyecto → se eliminan los registros anteriores (packing_list + lines) y se insertan los nuevos (**reemplazar**)
6. UI muestra tabla preview de las líneas extraídas + mensaje de éxito/reemplazo

**Estructura JSON esperada del LLM:**
```json
{
  "of_number": "OF26004215",
  "n_pedido": "2026000508 / 1",
  "cliente": "C001526 VOESTALPINE ALEMAN",
  "articulo": "P517780300GL1951 PERFIL C 160x67,5x67,5x19x19 Esp. 3 mm",
  "ref_pedido": "4500096097 CUBE WAREHOUSE T. D11",
  "desarrollo": "309",
  "espesor": "3,00",
  "calidad": "GL1951 A GALVANIZADO S450GD +Z140 MAC",
  "nota": "-63000",
  "total_piezas": 28,
  "kilos_teoricos": 2284.0,
  "fecha_doc": "2026-05-13",
  "paquetes": [
    {
      "paquete_num": 1,
      "kilos_paquete": 1285.0,
      "lineas": [
        {"linea": 1, "piezas": 3, "longitud": 12.713, "marca": "11889"},
        {"linea": 2, "piezas": 3, "longitud": 12.646, "marca": "11937"},
        {"linea": 3, "piezas": 8, "longitud": 12.561, "marca": "11917"}
      ]
    },
    {
      "paquete_num": 2,
      "kilos_paquete": 999.0,
      "lineas": [
        {"linea": 2, "piezas": 8, "longitud": 9.803, "marca": "11919"},
        {"linea": 3, "piezas": 3, "longitud": 9.803, "marca": "11939"},
        {"linea": 1, "piezas": 3, "longitud": 9.801, "marca": "11891"}
      ]
    }
  ]
}
```

---

## Interfaz Gradio

### Sidebar
- Campo de texto para API key de OpenRouter (se guarda en `OPEN_KEY.txt`)
- Dropdown de modelos disponibles + campo para añadir nuevo modelo + botón eliminar

### Página principal — flujo vertical

**Sección 1: Proyecto activo**
- Dropdown con proyectos existentes
- Campo de texto + botón "Crear proyecto"
- Indicador del proyecto activo seleccionado

**Sección 2: Importar nuevo packing list**
- File uploader (PDF)
- Botón "Procesar con LLM"
- Spinner durante el procesamiento
- Resultado: tabla de líneas extraídas + mensaje (nuevo / reemplazado / error)

**Sección 3: Packing lists del proyecto**
- Tabla con columnas: OF, N.Pedido, Cliente, Artículo, Fecha doc, Importado
- Botón "Exportar proyecto a Excel" → descarga `{nombre_proyecto}.xlsx`

---

## Excel de Exportación

Una fila por línea de paquete. Columnas:

| Proyecto | OF | N.Pedido | Cliente | Artículo | Ref.Pedido | Desarrollo | Espesor | Calidad | Nota | Total Piezas | Kilos Teóricos | Fecha Doc | N.Paquete | Kilos Paquete | Línea | Piezas | Longitud | Marca |

---

## Gestión de Duplicados

- Al importar un PDF: se detecta si `(project_id, of_number)` ya existe en la BD
- Si existe: se borran `lines` y `packing_list` anteriores, se insertan los nuevos
- Mensaje en UI: "OF26004215 reemplazado correctamente"
- Si no existe: inserción normal, mensaje "Importado correctamente"

---

## Dependencias Python

```
gradio>=4.0
openai              # cliente OpenRouter
pdf2image           # PDF → imágenes
Pillow              # manejo de imágenes
pandas              # tablas y Excel
openpyxl            # escritura Excel
```

---

## Configuración

- `OPEN_KEY.txt` en la raíz del proyecto (gitignored): API key de OpenRouter
- Modelo por defecto configurable en `config.py`
- Base de datos en `data/paquets.db` (creada automáticamente al primer arranque)
