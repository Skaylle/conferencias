<?php

declare(strict_types=1);

use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Cell\DataType;

const MAX_UPLOAD_SIZE = 50 * 1024 * 1024; // 50 MB

const ALLOWED_EXTENSIONS = [
    'xlsx',
    'xls',
];

const ALLOWED_MIMES = [
    'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    'application/vnd.ms-excel',
    'application/zip',
    'application/octet-stream',
];

const ARQUIVO1_COLUNA_CODIGO = 'B';
const ARQUIVO1_COLUNA_FALLBACK = 'A';
const ARQUIVO2_COLUNA_CODIGO = 'C';
const ARQUIVO2_COLUNA_DESCRICAO = 'D';

function storagePath(string $directory): string
{
    return __DIR__ . DIRECTORY_SEPARATOR . trim($directory, '/');
}

function ensureStorageDirectories(): void
{
    $directories = [
        storagePath('tmp'),
        storagePath('output'),
    ];

    foreach ($directories as $directory) {
        if (!is_dir($directory)) {
            if (!mkdir($directory, 0700, true) && !is_dir($directory)) {
                throw new RuntimeException(
                    'Não foi possível criar o diretório: ' . $directory
                );
            }
        }

        @chmod($directory, 0700);
    }
}

function generateToken(): string
{
    return bin2hex(random_bytes(24));
}

function extensionFromFilename(string $filename): string
{
    return strtolower(pathinfo($filename, PATHINFO_EXTENSION));
}

function validateUploadedExcel(
    array $file,
    string $fieldName
): void {
    if (!isset($file['error'])) {
        throw new RuntimeException(
            "O campo {$fieldName} não foi recebido corretamente."
        );
    }

    if ($file['error'] !== UPLOAD_ERR_OK) {
        $messages = [
            UPLOAD_ERR_INI_SIZE   => 'O arquivo excede o limite configurado pelo servidor.',
            UPLOAD_ERR_FORM_SIZE  => 'O arquivo excede o limite permitido pelo formulário.',
            UPLOAD_ERR_PARTIAL    => 'O upload foi realizado apenas parcialmente.',
            UPLOAD_ERR_NO_FILE    => 'Nenhum arquivo foi enviado.',
            UPLOAD_ERR_NO_TMP_DIR => 'O servidor não possui diretório temporário disponível.',
            UPLOAD_ERR_CANT_WRITE => 'Não foi possível gravar o arquivo temporário.',
            UPLOAD_ERR_EXTENSION  => 'Uma extensão do PHP interrompeu o upload.',
        ];

        $message = $messages[$file['error']] ?? 'Erro desconhecido no upload.';

        throw new RuntimeException($message);
    }

    if (!isset($file['size']) || $file['size'] <= 0) {
        throw new RuntimeException('O arquivo enviado está vazio.');
    }

    if ($file['size'] > MAX_UPLOAD_SIZE) {
        throw new RuntimeException(
            'O arquivo excede o limite de ' .
            (MAX_UPLOAD_SIZE / 1024 / 1024) .
            ' MB.'
        );
    }

    $originalName = (string)($file['name'] ?? '');
    $extension = extensionFromFilename($originalName);

    if (!in_array($extension, ALLOWED_EXTENSIONS, true)) {
        throw new RuntimeException(
            'Formato inválido. Envie um arquivo .xlsx ou .xls.'
        );
    }

    if (
        !isset($file['tmp_name']) ||
        !is_uploaded_file($file['tmp_name'])
    ) {
        throw new RuntimeException(
            'O arquivo enviado não foi reconhecido como upload válido.'
        );
    }

    $finfo = new finfo(FILEINFO_MIME_TYPE);
    $mime = $finfo->file($file['tmp_name']);

    if (
        $mime === false ||
        !in_array($mime, ALLOWED_MIMES, true)
    ) {
        throw new RuntimeException(
            'O arquivo enviado não parece ser um arquivo Excel válido.'
        );
    }

    try {
        IOFactory::createReaderForFile($file['tmp_name']);
    } catch (Throwable $e) {
        throw new RuntimeException(
            'O arquivo enviado não pôde ser reconhecido como uma planilha Excel válida.'
        );
    }
}

function saveUploadedFile(
    array $file,
    string $token,
    string $prefix
): string {
    $extension = extensionFromFilename((string)$file['name']);

    $safeName = $prefix . '_' . $token . '.' . $extension;

    $destination = storagePath('tmp') . DIRECTORY_SEPARATOR . $safeName;

    if (!move_uploaded_file($file['tmp_name'], $destination)) {
        throw new RuntimeException(
            'Não foi possível armazenar o arquivo temporariamente.'
        );
    }

    @chmod($destination, 0600);

    return $destination;
}

function loadSpreadsheet(
    string $path,
    ?string $sheetName = null,
    bool $readDataOnly = false
) {
    if (!is_file($path)) {
        throw new RuntimeException('Arquivo temporário não encontrado.');
    }

    try {
        $reader = IOFactory::createReaderForFile($path);

        if (method_exists($reader, 'setReadDataOnly')) {
            $reader->setReadDataOnly($readDataOnly);
        }

        if (
            $sheetName !== null &&
            method_exists($reader, 'setLoadSheetsOnly')
        ) {
            $reader->setLoadSheetsOnly([$sheetName]);
        }

        return $reader->load($path);
    } catch (Throwable $e) {
        throw new RuntimeException(
            'Não foi possível abrir a planilha. Verifique se o arquivo não está corrompido.'
        );
    }
}

function getWorksheetNames(string $path): array
{
    if (!is_file($path)) {
        throw new RuntimeException('Arquivo temporário não encontrado.');
    }

    try {
        $reader = IOFactory::createReaderForFile($path);

        if (method_exists($reader, 'listWorksheetNames')) {
            return $reader->listWorksheetNames($path);
        }

        $spreadsheet = $reader->load($path);

        return $spreadsheet->getSheetNames();
    } catch (Throwable $e) {
        throw new RuntimeException(
            'Não foi possível obter as abas do arquivo Excel.'
        );
    }
}

function normalizeCode(mixed $value): string
{
    if ($value === null) {
        return '';
    }

    if (is_bool($value)) {
        return $value ? '1' : '0';
    }

    $value = (string)$value;

    $value = str_replace("\xC2\xA0", ' ', $value);
    $value = trim($value);

    if ($value === '') {
        return '';
    }

    if (is_numeric($value)) {
        $number = (float)$value;

        if (is_finite($number)) {
            $normalized = sprintf('%.15g', $number);

            if (str_contains(strtolower($normalized), 'e')) {
                $normalized = rtrim(
                    rtrim(
                        sprintf('%.15f', $number),
                        '0'
                    ),
                    '.'
                );
            }

            return $normalized;
        }
    }

    return $value;
}

function htmlEscape(mixed $value): string
{
    return htmlspecialchars(
        (string)$value,
        ENT_QUOTES | ENT_SUBSTITUTE,
        'UTF-8'
    );
}

function cleanupFile(?string $path): void
{
    if ($path !== null && is_file($path)) {
        @unlink($path);
    }
}

function cleanupOldFiles(string $directory, int $maxAgeSeconds = 3600): void
{
    $path = storagePath($directory);

    if (!is_dir($path)) {
        return;
    }

    $files = glob($path . DIRECTORY_SEPARATOR . '*');

    if ($files === false) {
        return;
    }

    $now = time();

    foreach ($files as $file) {
        if (!is_file($file)) {
            continue;
        }

        $mtime = @filemtime($file);

        if ($mtime !== false && ($now - $mtime) > $maxAgeSeconds) {
            @unlink($file);
        }
    }
}

function setSpreadsheetCellAsString(
    \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet $worksheet,
    string $coordinate,
    string $value
): void {
    $worksheet
        ->getCell($coordinate)
        ->setValueExplicit($value, DataType::TYPE_STRING);
}

function getDefaultRepeatedTreatmentConfig(?array $resultadoBase): array {
    return [
        'base_compare_column' => 'A',
        'base_start_row' => 4,
        'reference_compare_column' => 'A',
        'reference_start_row' => 4,
        'result_column' => '',
        'rules' => [
            [
                'column' => '',
                'text' => '',
                'append_value' => false
            ]
        ]
    ];
}

function getSubmittedRepeatedTreatmentConfig(array $input): array {
    return [
        'base_compare_column' => strtoupper(trim((string)($input['base_compare_column'] ?? 'A'))),
        'base_start_row' => max(4, (int)($input['base_start_row'] ?? 4)),
        'reference_compare_column' => strtoupper(trim((string)($input['reference_compare_column'] ?? 'A'))),
        'reference_start_row' => max(4, (int)($input['reference_start_row'] ?? 4)),
        'result_column' => strtoupper(trim((string)($input['result_column'] ?? ''))),
    ];
}

function getSubmittedTreatmentRules(array $input): array {
    $rawRules = $input['rules'] ?? [];
    $rules = [];

    foreach ($rawRules as $idx => $rule) {
        $col = strtoupper(trim((string)($rule['column'] ?? '')));
        if ($col !== '') {
            $rules[] = [
                'index' => $idx,
                'column' => $col,
                'text' => (string)($rule['text'] ?? ''),
                'append_value' => !empty($rule['append_value'])
            ];
        }
    }

    return $rules;
}

function evaluateTreatmentRules($worksheet, int $rowIndex, array $rules): array {
    $matches = [];

    foreach ($rules as $rule) {
        $cellCoordinate = $rule['column'] . $rowIndex;
        if ($worksheet->cellExists($cellCoordinate)) {
            $val = trim((string)$worksheet->getCell($cellCoordinate)->getValue());
            if ($val !== '') {
                $matches[] = [
                    'rule' => $rule,
                    'cell_value' => $val
                ];
            }
        }
    }

    if (empty($matches)) {
        return ['status' => 'no_rule', 'matches' => []];
    }

    return ['status' => 'success', 'matches' => $matches];
}

function generateTreatmentText(array $rule, string $cellValue): string {
    $text = $rule['text'];
    if ($rule['append_value']) {
        $text .= $cellValue;
    }
    return $text;
}

function buildRuleStatistics(array $rules): array {
    $stats = [];
    foreach ($rules as $rule) {
        $stats[$rule['index']] = [
            'column' => $rule['column'],
            'count' => 0
        ];
    }
    return $stats;
}

function cleanupStageTwoSessionFiles(): void {
    if (isset($_SESSION['arquivo3_path'])) {
        cleanupFile($_SESSION['arquivo3_path']);
        unset($_SESSION['arquivo3_path'], $_SESSION['arquivo3_original_name']);
    }
}

function findWorksheetOrFail(\PhpOffice\PhpSpreadsheet\Spreadsheet $spreadsheet, string $sheetName): \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet
{
    $sheet = $spreadsheet->getSheetByName($sheetName);

    if ($sheet === null) {
        throw new RuntimeException("A aba '{$sheetName}' não foi encontrada na planilha.");
    }

    return $sheet;
}

function clearSession(): void
{
    unset($_SESSION['arquivo1_path'], $_SESSION['arquivo1_original_name']);
    unset($_SESSION['arquivo2_path'], $_SESSION['arquivo2_original_name']);
    unset($_SESSION['arquivo3_path'], $_SESSION['arquivo3_original_name']);
    unset($_SESSION['treatment_config']);
    unset($_SESSION['treatment_rules']);
}

/**
 * Cria um índice (hash map) de uma aba com base em uma coluna chave ou lista de códigos.
 *
 * @param \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet $sheet
 * @param string $keyColumn Coluna chave para comparação (ex: 'A')
 * @param int $startRow Linha inicial para varredura
 * @param array $filterKeys Lista de códigos específicos (se vazia, lê toda a coluna)
 * @return array<string, int>
 */
function buildWorksheetIndex(
    \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet $sheet,
    string $keyColumn,
    int $startRow = 1,
    array $filterKeys = []
): array {
    $index = [];
    $highestRow = $sheet->getHighestRow();

    $allowedKeys = [];
    if (!empty($filterKeys)) {
        foreach ($filterKeys as $fk) {
            $cleaned = trim((string)$fk);
            if ($cleaned !== '') {
                $allowedKeys[$cleaned] = true;
            }
        }
    }

    for ($row = $startRow; $row <= $highestRow; $row++) {
        $cellValue = $sheet->getCell($keyColumn . $row)->getValue();
        
        if ($cellValue === null) {
            continue;
        }

        $key = trim((string)$cellValue);

        if ($key === '') {
            continue;
        }

        if (!empty($allowedKeys) && !isset($allowedKeys[$key])) {
            continue;
        }

        if (!isset($index[$key])) {
            $index[$key] = $row;
        }
    }

    return $index;
}
