<?php
// procesar_excel_html.php
// Procesa un Excel subido por formulario, aplicando etiquetas HTML según formato de texto.
// Compatible con PHP 8.5 y PhpSpreadsheet 5.3.x

declare(strict_types=1);

// Evita warning de JIT en algunos sistemas (macOS, etc.)
ini_set('pcre.jit', '0');

use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;
use PhpOffice\PhpSpreadsheet\Cell\Coordinate;
use PhpOffice\PhpSpreadsheet\RichText\RichText;
use PhpOffice\PhpSpreadsheet\RichText\Run;
use PhpOffice\PhpSpreadsheet\Style\Font;
use PhpOffice\PhpSpreadsheet\Cell\Cell;
use PhpOffice\PhpSpreadsheet\Cell\DataType;

require 'vendor/autoload.php';

// ---------------------------------------------------------------------
// 1) Gestión del archivo subido
// ---------------------------------------------------------------------
if (
    !isset($_FILES['excel_file']) ||
    $_FILES['excel_file']['error'] !== UPLOAD_ERR_OK
) {
    $mensaje = 'No se recibió ningún archivo o hubo un error en la subida.';
    header('Location: index.php?tipo=error&mensaje=' . urlencode($mensaje));
    exit;
}

$originalName = $_FILES['excel_file']['name'];
$tmpPath      = $_FILES['excel_file']['tmp_name'];

$ext = strtolower(pathinfo($originalName, PATHINFO_EXTENSION));
if ($ext !== 'xlsx') {
    $mensaje = 'El archivo debe ser un .xlsx (Excel moderno).';
    header('Location: index.php?tipo=error&mensaje=' . urlencode($mensaje));
    exit;
}

// Carpeta para entrada/salida
$uploadsDir = __DIR__ . '/uploads';
if (!is_dir($uploadsDir)) {
    mkdir($uploadsDir, 0775, true);
}

// Nombre seguro para el archivo de entrada
$inputFileSanitized = preg_replace('/[^A-Za-z0-9_\-\.]/', '_', $originalName);
$inputPath  = $uploadsDir . '/in_' . uniqid() . '_' . $inputFileSanitized;
if (!move_uploaded_file($tmpPath, $inputPath)) {
    $mensaje = 'No se pudo guardar el archivo subido.';
    header('Location: index.php?tipo=error&mensaje=' . urlencode($mensaje));
    exit;
}

// Nombre y ruta del archivo de salida
$outputBaseName = preg_replace('/\.xlsx$/i', '', $inputFileSanitized);
$outputBaseName .= '_html.xlsx';
$outputPath = $uploadsDir . '/out_' . uniqid() . '_' . $outputBaseName;

// ---------------------------------------------------------------------
// 2) Carga del Excel origen y creación del libro de salida
// ---------------------------------------------------------------------
$spreadsheetIn = IOFactory::load($inputPath);

$spreadsheetOut = new Spreadsheet();
$spreadsheetOut->removeSheetByIndex(0); // fuera hoja por defecto

$sheetIndexOut = 0;

// ---------------------------------------------------------------------
// 3) Recorremos solo las pestañas visibles
// ---------------------------------------------------------------------
foreach ($spreadsheetIn->getWorksheetIterator() as $sheetIn) {
    if ($sheetIn->getSheetState() !== Worksheet::SHEETSTATE_VISIBLE) {
        continue; // saltar hojas ocultas
    }

    /** @var Worksheet $sheetIn */
    $sheetOut = new Worksheet($spreadsheetOut, $sheetIn->getTitle());
    $spreadsheetOut->addSheet($sheetOut, $sheetIndexOut);
    $sheetIndexOut++;

    // Copiar dimensiones de columnas y filas para facilitar inspección visual
    foreach ($sheetIn->getColumnDimensions() as $colDim) {
        $col = $colDim->getColumnIndex();
        $targetColDim = $sheetOut->getColumnDimension($col);
        $targetColDim->setWidth($colDim->getWidth());
        $targetColDim->setVisible($colDim->getVisible());
    }

    foreach ($sheetIn->getRowDimensions() as $rowIndex => $rowDim) {
        $targetRowDim = $sheetOut->getRowDimension($rowIndex);
        $targetRowDim->setRowHeight($rowDim->getRowHeight());
        $targetRowDim->setVisible($rowDim->getVisible());
    }

    $highestRow    = $sheetIn->getHighestRow();
    $highestColumn = $sheetIn->getHighestColumn();
    $highestColIdx = Coordinate::columnIndexFromString($highestColumn);

    // -------------------------------------------------------------
    // 3.1) Detectar columnas cuyo encabezado sea exactamente "Language"
    // -------------------------------------------------------------
    $languageColumns = [];
    for ($col = 1; $col <= $highestColIdx; $col++) {
        $colLetter   = Coordinate::stringFromColumnIndex($col);
        $coord       = $colLetter . '1';
        $cell        = $sheetIn->getCell($coord);
        $headerValue = trim((string) $cell->getValue());
        if ($headerValue === 'Language') {
            $languageColumns[$col] = true;
        }
    }

    $saltarSiguienteFilaPorGeoloc = false;

    // -------------------------------------------------------------
    // 3.2) Recorremos filas visibles
    // -------------------------------------------------------------
    for ($row = 1; $row <= $highestRow; $row++) {
        $rowDimIn = $sheetIn->getRowDimension($row);
        if ($rowDimIn && $rowDimIn->getVisible() === false) {
            continue; // fila oculta
        }

        // Primera celda de la fila para detectar Geoloc
        $primeraCeldaCoord  = 'A' . $row; // col 1 = A
        $primeraCelda       = $sheetIn->getCell($primeraCeldaCoord);
        $primeraCeldaTexto  = trim((string) $primeraCelda->getValue());
        $filaSinTransformHTML = false;

        if ($primeraCeldaTexto === 'Geoloc') {
            $filaSinTransformHTML       = true;
            $saltarSiguienteFilaPorGeoloc = true;
        } elseif ($saltarSiguienteFilaPorGeoloc) {
            $filaSinTransformHTML       = true;
            $saltarSiguienteFilaPorGeoloc = false;
        }

        // Asegurar visibilidad de la fila de salida
        $sheetOut->getRowDimension($row)->setVisible(true);

        // ---------------------------------------------------------
        // 3.3) Recorremos columnas visibles
        // ---------------------------------------------------------
        for ($col = 1; $col <= $highestColIdx; $col++) {
            $colLetter = Coordinate::stringFromColumnIndex($col);
            $colDimIn  = $sheetIn->getColumnDimension($colLetter);
            if ($colDimIn && $colDimIn->getVisible() === false) {
                continue; // columna oculta
            }

            $coord = $colLetter . $row;

            $cellIn  = $sheetIn->getCell($coord);
            $cellOut = $sheetOut->getCell($coord);

            // Copiar estilo visual
            $sheetOut->duplicateStyle($cellIn->getStyle(), $cellOut->getCoordinate());

            $rawValue = $cellIn->getValue();

            // Columnas Language o filas Geoloc (y la siguiente): copiar texto tal cual
            if (isset($languageColumns[$col]) || $filaSinTransformHTML) {
                $cellOut->setValue($rawValue);
                continue;
            }

            // Si no es string ni RichText, copiar tal cual (números, fechas, etc.)
            if (!is_string($rawValue) && !($rawValue instanceof RichText)) {
                $cellOut->setValue($rawValue);
                continue;
            }

            // Procesar texto / rich text → HTML
            $textoHTML = procesarTextoCeldaAHtml($cellIn);

            $cellOut->setValueExplicit($textoHTML, DataType::TYPE_STRING);
        }
    }
}

// ---------------------------------------------------------------------
// 4) Guardar archivo de salida y enviarlo al navegador
// ---------------------------------------------------------------------
$writer = IOFactory::createWriter($spreadsheetOut, 'Xlsx');
$writer->save($outputPath);

header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
header('Content-Disposition: attachment; filename="' . basename($outputBaseName) . '"');
header('Content-Length: ' . filesize($outputPath));
readfile($outputPath);

// Opcional: limpiar archivos temporales si quieres
// @unlink($inputPath);
// @unlink($outputPath);

exit;

// =====================================================================
//  FUNCIONES AUXILIARES
// =====================================================================

/**
 * Procesa una celda (string o RichText) y devuelve HTML con:
 * - Respeto a rich text (fragmentos con negrita/cursiva/subrayado).
 * - Sin romper palabras (toda la palabra hereda el formato).
 * - Puntuación fuera de las etiquetas.
 * - Manejo de líneas que empiezan con "- ".
 */
function procesarTextoCeldaAHtml(Cell $cell): string
{
    $value = $cell->getValue();
    $runs = [];

    // Caso RichText: leer cada elemento de texto con su formato, si lo tiene
    if ($value instanceof RichText) {
        /** @var RichText $value */
        foreach ($value->getRichTextElements() as $element) {
            $text = $element->getText();
            if ($text === '') {
                continue;
            }

            // Valores por defecto (sin formato)
            $bold = false;
            $italic = false;
            $underline = false;

            // En PhpSpreadsheet 5.x, los fragmentos con formato parcial suelen ser de tipo Run
            // y exponen getFont(). Otros elementos (texto plano) pueden no tener font propia.
            $font = null;
            if (method_exists($element, 'getFont')) {
                $font = $element->getFont();
            }

            if ($font instanceof Font) {
                $bold      = $font->getBold();
                $italic    = $font->getItalic();
                $underline = $font->getUnderline() !== Font::UNDERLINE_NONE;
            } else {
                // Sin font propia en el elemento: heredamos el estilo de la celda
                $cellFont = $cell->getStyle()->getFont();
                $bold      = $cellFont->getBold();
                $italic    = $cellFont->getItalic();
                $underline = $cellFont->getUnderline() !== Font::UNDERLINE_NONE;
            }

            $runs[] = [
                'text'      => $text,
                'bold'      => $bold,
                'italic'    => $italic,
                'underline' => $underline,
            ];
        }
    } else {
        // Texto plano: un único run con el estilo de la celda
        $text = (string) $value;
        $font = $cell->getStyle()->getFont();

        $bold      = $font->getBold();
        $italic    = $font->getItalic();
        $underline = $font->getUnderline() !== Font::UNDERLINE_NONE;

        $runs[] = [
            'text'      => $text,
            'bold'      => $bold,
            'italic'    => $italic,
            'underline' => $underline,
        ];
    }

    return generarHtmlDesdeRuns($runs);
}

/**
 * A partir de la lista de runs (texto + flags de formato), genera HTML:
 * - Construye caracteres con formato.
 * - Separa por líneas (respeta CR/LF).
 * - Detecta "- " al inicio de cada línea y lo elimina.
 * - Tokeniza en palabras/espacios/puntuación.
 * - Aplica formato solo a las palabras, sin incluir puntuación.
 */
function generarHtmlDesdeRuns(array $runs): string
{
    // 1) Convertir runs a array de caracteres con formato
    $chars = [];

    foreach ($runs as $run) {
        $text      = $run['text'];
        $bold      = (bool)$run['bold'];
        $italic    = (bool)$run['italic'];
        $underline = (bool)$run['underline'];

        $length = mb_strlen($text, 'UTF-8');
        for ($i = 0; $i < $length; $i++) {
            $ch = mb_substr($text, $i, 1, 'UTF-8');
            $chars[] = [
                'char'      => $ch,
                'bold'      => $bold,
                'italic'    => $italic,
                'underline' => $underline,
            ];
        }
    }

    // 2) Separar en líneas (respeta \r, \n y \r\n)
    $lines = [];
    $currentLine = [];

    $totalChars = count($chars);
    for ($i = 0; $i < $totalChars; $i++) {
        $ch = $chars[$i]['char'];

        if ($ch === "\r" || $ch === "\n") {
            // Combinar CRLF como un solo salto
            if ($ch === "\r" && $i + 1 < $totalChars && $chars[$i + 1]['char'] === "\n") {
                $i++;
            }
            $lines[]    = $currentLine;
            $currentLine = [];
        } else {
            $currentLine[] = $chars[$i];
        }
    }
    // Última línea (aunque esté vacía)
    $lines[] = $currentLine;

    // 3) Procesar cada línea: quitar "- " inicial y tokenizar
    $htmlLines = [];

    foreach ($lines as $lineChars) {
        // Eliminar marcador "- " al inicio (si existe)
        if (count($lineChars) >= 2
            && $lineChars[0]['char'] === '-'
            && $lineChars[1]['char'] === ' '
        ) {
            array_shift($lineChars);
            array_shift($lineChars);
        }

        $tokens   = tokenizarLineaPorPalabrasYPuntuacion($lineChars);
        $lineHtml = generarHtmlDesdeTokens($tokens);

        $htmlLines[] = $lineHtml;
    }

    // Unir líneas con \n (mismo separador que usábamos en reglas)
    return implode("\n", $htmlLines);
}

/**
 * Tokeniza una línea de caracteres en:
 *  - 'word'  (letras/números)
 *  - 'space' (espacios/tabulaciones)
 *  - 'punct' (resto de signos, puntuación, etc.)
 *
 * Cada token agrega el formato de todos sus caracteres:
 *  - Si alguna parte de la palabra está en negrita, la palabra entera se marca en negrita.
 *  - Lo mismo para cursiva y subrayado.
 */
function tokenizarLineaPorPalabrasYPuntuacion(array $lineChars): array
{
    $tokens = [];
    $current = null;

    foreach ($lineChars as $chInfo) {
        $ch = $chInfo['char'];

        // Categoría básica
        if (preg_match('/\s/u', $ch)) {
            $type = 'space';
        } elseif (preg_match('/[\p{L}\p{N}]/u', $ch)) {
            $type = 'word';
        } else {
            $type = 'punct';
        }

        if ($current === null) {
            $current = [
                'type'      => $type,
                'text'      => $ch,
                'bold'      => (bool)$chInfo['bold'],
                'italic'    => (bool)$chInfo['italic'],
                'underline' => (bool)$chInfo['underline'],
            ];
            continue;
        }

        // Si el tipo coincide, agregamos al token actual
        if ($current['type'] === $type) {
            $current['text']     .= $ch;
            $current['bold']      = $current['bold']      || (bool)$chInfo['bold'];
            $current['italic']    = $current['italic']    || (bool)$chInfo['italic'];
            $current['underline'] = $current['underline'] || (bool)$chInfo['underline'];
        } else {
            // Cerramos token actual y comenzamos uno nuevo
            $tokens[] = $current;
            $current = [
                'type'      => $type,
                'text'      => $ch,
                'bold'      => (bool)$chInfo['bold'],
                'italic'    => (bool)$chInfo['italic'],
                'underline' => (bool)$chInfo['underline'],
            ];
        }
    }

    if ($current !== null) {
        $tokens[] = $current;
    }

    return $tokens;
}

/**
 * Genera HTML a partir de tokens:
 * - Agrupa palabras contiguas con el mismo formato en un solo bloque,
 *   incluyendo los espacios intermedios.
 * - La puntuación queda siempre fuera de las etiquetas.
 */
function generarHtmlDesdeTokens(array $tokens): string
{
    $html = '';

    $n = count($tokens);
    $i = 0;

    while ($i < $n) {
        $token = $tokens[$i];
        $type  = $token['type'];
        $text  = $token['text'];
        $bold  = (bool)$token['bold'];
        $italic = (bool)$token['italic'];
        $underline = (bool)$token['underline'];

        // Solo las palabras con algún formato inician un bloque formateado
        if ($type === 'word' && ($bold || $italic || $underline)) {
            $blockBold = $bold;
            $blockItalic = $italic;
            $blockUnderline = $underline;

            $start = $i;
            $j = $i + 1;

            // Extendemos el bloque mientras:
            // - Los tokens sean space/punct/word
            // - Cualquier token 'word' del bloque tenga exactamente el mismo formato
            while ($j < $n) {
                $t = $tokens[$j];
                if ($t['type'] === 'word') {
                    if ((bool)$t['bold'] === $blockBold
                        && (bool)$t['italic'] === $blockItalic
                        && (bool)$t['underline'] === $blockUnderline) {
                        $j++;
                        continue;
                    }
                    break;
                } elseif ($t['type'] === 'space' || $t['type'] === 'punct') {
                    $j++;
                    continue;
                } else {
                    break;
                }
            }

            // Ahora tenemos un bloque candidato [start, j)
            // Queremos excluir solo la puntuación inicial/final, pero mantener
            // la puntuación interna dentro del bloque formateado.

            $firstContent = $start;
            while ($firstContent < $j && $tokens[$firstContent]['type'] === 'punct') {
                $firstContent++;
            }

            $lastContent = $j - 1;
            while ($lastContent >= $start && $tokens[$lastContent]['type'] === 'punct') {
                $lastContent--;
            }

            // Puntuación inicial (antes del contenido) → siempre fuera de etiquetas
            for ($k = $start; $k < $firstContent; $k++) {
                $html .= $tokens[$k]['text'];
            }

            if ($firstContent <= $lastContent) {
                $innerText = '';
                for ($k = $firstContent; $k <= $lastContent; $k++) {
                    $innerText .= $tokens[$k]['text'];
                }
                $html .= aplicarFormatoHtmlBasico($innerText, $blockBold, $blockItalic, $blockUnderline);
            }

            // Puntuación final (después del contenido) → siempre fuera de etiquetas
            for ($k = $lastContent + 1; $k < $j; $k++) {
                $html .= $tokens[$k]['text'];
            }

            $i = $j;
        } else {
            // Cualquier cosa que no sea inicio de bloque formateado se emite tal cual
            $html .= $text;
            $i++;
        }
    }

    return $html;
}


/**
 * Aplica la combinación de <strong>, <i>, <u> al texto según negrita/cursiva/subrayado.
 * Orden de anidación:
 *   <strong> → <i> → <u>
 */
function aplicarFormatoHtmlBasico(string $texto, bool $bold, bool $italic, bool $underline): string
{
    if (!$bold && !$italic && !$underline) {
        return $texto;
    }

    $prefix = '';
    $suffix = '';

    if ($bold) {
        $prefix .= '<strong>';
        $suffix  = '</strong>' . $suffix;
    }
    if ($italic) {
        $prefix .= '<i>';
        $suffix  = '</i>' . $suffix;
    }
    if ($underline) {
        $prefix .= '<u>';
        $suffix  = '</u>' . $suffix;
    }

    return $prefix . $texto . $suffix;
}
