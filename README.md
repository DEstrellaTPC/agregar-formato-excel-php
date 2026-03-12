
# Agregar Formato a Tablas HTML Exportadas a Excel (PHP + PhpSpreadsheet)

Este proyecto permite **tomar una tabla HTML enviada mediante POST**, procesarla con **PhpSpreadsheet** y generar un archivo Excel con formateos básicos para mejorar su presentación.
El objetivo es proporcionar una forma sencilla de **convertir una tabla HTML en un archivo .xlsx con estilos**, sin necesidad de procesar celda por celda manualmente.

---

## 🚀 ¿Qué hace este proyecto?

A partir de un formulario donde se envía HTML crudo (una tabla completa), el script:

- Limpia el HTML recibido.
- Convierte la tabla HTML en un documento de Excel mediante `Html::load()`.
- Aplica estilos de formato básicos:
  - Negritas en encabezados
  - Autoajuste del ancho de columnas
  - Bordes externos
  - Alineación
- Genera y descarga el archivo final `formato_aplicado.xlsx`.

El resultado es un Excel más limpio, legible y presentable.

---

## 📂 Archivos principales

### `index.php`
Formulario HTML minimalista que contiene un `<textarea>` para pegar código HTML (normalmente tablas con `thead` + `tbody`), y lo envía vía POST al procesador.

Incluye:
- Validación simple
- Vista previa del contenido enviado
- Enlace para convertir el HTML en Excel

### `procesar_excel_html.php`
Script principal que:

1. Recibe el contenido HTML desde `POST["contenido_html"]`.
2. Utiliza `PhpOffice\PhpSpreadsheet\Reader\Html` para interpretarlo.
3. Inserta la tabla en un nuevo libro de Excel.
4. Aplica estilos y autoajustes.
5. Envía el archivo Excel al navegador con headers correctos.

Fragmento de ejemplo del procesamiento:

```php
$reader = new Html();
$spreadsheet = $reader->loadFromString($contenido_html);
$sheet = $spreadsheet->getActiveSheet();

// Autoajustar columnas
foreach(range('A', $sheet->getHighestColumn()) as $col) {
    $sheet->getColumnDimension($col)->setAutoSize(true);
}

// Encabezados en negritas
$sheet->getStyle('1:1')->getFont()->setBold(true);

// Generar archivo
$writer = new Xlsx($spreadsheet);
$writer->save('php://output');
```

---

## ▶️ Cómo usarlo

### 1) Instalar dependencias

```bash
composer require phpoffice/phpspreadsheet
```

### 2) Requisitos

- PHP 8.1 o superior
- Extensiones: `zip`, `xml`, `mbstring`

### 3) Ejecutar

1. Subir el proyecto a un servidor con PHP.
2. Abrir `index.php` en el navegador.
3. Pegar una tabla HTML como:

```html
<table>
  <thead>
    <tr><th>Producto</th><th>Precio</th></tr>
  </thead>
  <tbody>
    <tr><td>Café</td><td>35</td></tr>
    <tr><td>Pan</td><td>12</td></tr>
  </tbody>
</table>
```

4. Hacer clic en **"Procesar Excel"** para descargar el `.xlsx` formateado.

---

## 📑 Estructura del proyecto

```
agregar-formato-excel-php/
├── index.php
├── procesar_excel_html.php
├── vendor/ (generado por Composer)
└── README.md
```

---

## 🧩 Limitaciones conocidas

- Procesa una tabla HTML por ejecución.
- `rowspan`/`colspan` complejos pueden no preservarse.
- Los estilos aplicados son básicos, pero pueden ampliarse.

---

## 🤝 Contribuciones

PRs y sugerencias son bienvenidos.

---

## 📄 Licencia

MIT
