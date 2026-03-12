
# Agregar Formato a Excel (PHP + PhpSpreadsheet)

Este proyecto permite **recibir un archivo Excel (.xlsx)** desde un formulario HTML, procesarlo con **PhpSpreadsheet**, aplicar estilos básicos y generar un nuevo archivo Excel con formato mejorado.

El objetivo es proporcionar una forma sencilla de **subir un archivo Excel**, aplicarle formateo automático y descargar la versión estilizada.

---
## 🚀 ¿Qué hace este proyecto?

1. El usuario **sube un archivo .xlsx** mediante `index.php`.
2. El archivo se envía a `procesar_excel_html.php`.
3. El script carga el Excel usando `PhpOffice\PhpSpreadsheet\IOFactory`.
4. Se aplican estilos automáticos:
   - Negritas en la primera fila
   - Autoajuste de columnas
   - Bordes
   - Alineaciones
5. Se genera y descarga un archivo Excel actualizado.

---
## 📂 Archivos principales

### `index.php`
Formulario HTML que contiene:
- Campo `<input type="file">` que permite subir un archivo `.xlsx`
- Envío mediante POST a `procesar_excel_html.php`

### `procesar_excel_html.php`
Script que:
1. Recibe el archivo subido en `$_FILES`.
2. Lo carga con `IOFactory::load()`.
3. Obtiene la hoja activa.
4. Aplica estilos según las reglas del proyecto.
5. Descarga el nuevo archivo.

---
## ▶️ Cómo usarlo

### 1) Instalar dependencias
```bash
composer require phpoffice/phpspreadsheet
```

### 2) Requisitos
- PHP 8.1+
- Extensiones: `zip`, `xml`, `mbstring`

### 3) Ejecutar
1. Subir proyecto a un servidor PHP.
2. Abrir `index.php`.
3. Seleccionar un archivo `.xlsx`.
4. Descargar archivo procesado.

---
## 📑 Estructura del proyecto
```
agregar-formato-excel-php/
├── index.php
├── procesar_excel_html.php
├── vendor/ (Composer)
└── README.md
```

---
## 🧩 Limitaciones
- Procesa solo una hoja a la vez, salvo modificaciones.
- Estilos básicos pero expandibles.

---
## 📄 Licencia
MIT
