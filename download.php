<?php

declare(strict_types=1);

session_start();

require __DIR__ . '/functions.php';

$resultado = $_SESSION['resultado'] ?? null;

if (
    !is_array($resultado) ||
    empty($resultado['path']) ||
    empty($resultado['filename'])
) {
    http_response_code(404);

    exit(
        'Arquivo processado não encontrado ou sessão expirada.'
    );
}

$path = (string)$resultado['path'];
$filename = basename((string)$resultado['filename']);

if (!is_file($path)) {
    unset($_SESSION['resultado']);

    http_response_code(404);

    exit(
        'O arquivo processado não está mais disponível.'
    );
}

/*
 * Proteção adicional:
 * garante que o arquivo está dentro da pasta output.
 */
$realOutputDir = realpath(storagePath('output'));
$realFilePath = realpath($path);

if (
    $realOutputDir === false ||
    $realFilePath === false ||
    !str_starts_with(
        $realFilePath,
        $realOutputDir . DIRECTORY_SEPARATOR
    )
) {
    http_response_code(403);

    exit('Arquivo inválido.');
}

$size = filesize($path);

if ($size === false) {
    http_response_code(500);

    exit('Não foi possível obter o tamanho do arquivo.');
}

header(
    'Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
);

header(
    'Content-Disposition: attachment; filename="' .
    rawurlencode($filename) .
    '"'
);

header('Content-Length: ' . $size);

header('Cache-Control: private, no-store, no-cache, must-revalidate');
header('Pragma: no-cache');
header('Expires: 0');

if (ob_get_level() > 0) {
    ob_end_clean();
}

$handle = fopen($path, 'rb');

if ($handle === false) {
    http_response_code(500);

    exit('Não foi possível abrir o arquivo para download.');
}

while (!feof($handle)) {
    echo fread($handle, 1024 * 1024);
    flush();
}

fclose($handle);

/*
 * Depois do download, remove o resultado do servidor.
 */
cleanupFile($path);

unset($_SESSION['resultado']);
