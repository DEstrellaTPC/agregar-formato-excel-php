<?php
// index.php
// Página inicial con formulario para subir un Excel y procesarlo con procesar_excel_html.php

declare(strict_types=1);
?>
<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <title>Procesar Excel a HTML</title>
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <style>
        body {
            font-family: system-ui, -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
            margin: 0;
            padding: 0;
            background: #f5f5f5;
        }
        .container {
            max-width: 480px;
            margin: 40px auto;
            background: #ffffff;
            padding: 24px 28px;
            border-radius: 8px;
            box-shadow: 0 4px 10px rgba(0,0,0,0.06);
        }
        h1 {
            font-size: 1.4rem;
            margin-top: 0;
            margin-bottom: 0.5rem;
        }
        p.desc {
            margin-top: 0;
            color: #555;
            font-size: 0.9rem;
        }
        label {
            display: block;
            margin-bottom: 0.4rem;
            font-weight: 600;
        }
        input[type="file"] {
            display: block;
            width: 100%;
            margin-bottom: 1rem;
        }
        button {
            display: inline-block;
            padding: 0.55rem 1.2rem;
            border-radius: 4px;
            border: none;
            cursor: pointer;
            font-weight: 600;
            font-size: 0.95rem;
        }
        .btn-primary {
            background: #007bff;
            color: #fff;
        }
        .btn-primary:hover {
            filter: brightness(1.05);
        }
        .alert {
            padding: 0.7rem 0.9rem;
            border-radius: 4px;
            margin-bottom: 1rem;
            font-size: 0.9rem;
        }
        .alert-success {
            background: #e6f6e9;
            color: #216e39;
        }
        .alert-error {
            background: #fde2e1;
            color: #b3261e;
        }
        .download-link {
            margin-top: 0.7rem;
            font-size: 0.9rem;
        }
        .download-link a {
            text-decoration: none;
        }
    </style>
</head>
<body>
<div class="container">
    <h1>Procesar Excel</h1>
    <p class="desc">
        Selecciona un archivo <strong>.xlsx</strong> para aplicar las etiquetas HTML
        (<code>&lt;strong&gt;</code>, <code>&lt;i&gt;</code>, <code>&lt;u&gt;</code>) según el formato del texto.
    </p>

    <?php
    $mensaje = $_GET["mensaje"] ?? null;
    $tipo    = $_GET["tipo"] ?? null;
    if ($mensaje): ?>
        <div class="alert <?= $tipo === 'ok' ? 'alert-success' : 'alert-error' ?>">
            <?= htmlspecialchars($mensaje, ENT_QUOTES, 'UTF-8') ?>
        </div>
    <?php endif; ?>

    <form action="procesar_excel_html.php" method="post" enctype="multipart/form-data">
        <label for="excel_file">Archivo Excel (.xlsx)</label>
        <input
            type="file"
            name="excel_file"
            id="excel_file"
            accept=".xlsx,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            required
        >
        <button type="submit" class="btn-primary">Procesar</button>
    </form>
</div>
</body>
</html>
