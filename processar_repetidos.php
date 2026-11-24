<?php

declare(strict_types=1);

session_start();

require __DIR__ . '/vendor/autoload.php';
require __DIR__ . '/functions.php';

ensureStorageDirectories();
@ini_set('memory_limit', '1024M');
@set_time_limit(300);

$error = null;

try {
    if ($_SERVER['REQUEST_METHOD'] !== 'POST') {
        throw new RuntimeException('Método de acesso inválido.');
    }

    // Resolver caminho do arquivo base
    $basePath = $_SESSION['resultado']['path'] ?? null;
    if (!is_string($basePath) || !is_file($basePath)) {
        $arquivos = glob(__DIR__ . '/output/*.xlsx');
        if (!empty($arquivos)) {
            usort($arquivos, fn($a, $b) => filemtime($b) <=> filemtime($a));
            $basePath = $arquivos[0];
        } else {
            throw new RuntimeException('Planilha base não encontrada.');
        }
    }

    $arquivo3Path = $_SESSION['arquivo3_path'] ?? null;
    if (!is_string($arquivo3Path) || !is_file($arquivo3Path)) {
        throw new RuntimeException('A terceira planilha não está disponível. Faça o upload novamente.');
    }

    $abaBase = trim((string)($_POST['aba_base'] ?? ''));
    $abaArquivo3 = trim((string)($_POST['aba_arquivo3'] ?? ''));

    if ($abaBase === '') {
        throw new RuntimeException('Selecione a aba da planilha base.');
    }

    if ($abaArquivo3 === '') {
        throw new RuntimeException('Selecione a aba da terceira planilha.');
    }

    // Captura o texto para registros não tratados (se informado pelo usuário)
    $textoSemTratamento = trim((string)($_POST['texto_sem_tratamento'] ?? ''));

    // Processar códigos manuais/específicos informados no textarea (se houver)
    $rawCodigos = $_POST['codigos_especificos'] ?? '';
    $codigosEspecificos = [];

    if (trim((string)$rawCodigos) !== '') {
        $elementos = preg_split('/[\n\r,;]+/', (string)$rawCodigos);
        if (is_array($elementos)) {
            foreach ($elementos as $elem) {
                $limpo = trim($elem);
                if ($limpo !== '') {
                    $codigosEspecificos[] = $limpo;
                }
            }
        }
    }

    $processingConfig = getSubmittedRepeatedTreatmentConfig($_POST);
    $treatmentRules = getSubmittedTreatmentRules($_POST);

    $spreadsheet3 = loadSpreadsheet($arquivo3Path, $abaArquivo3, true);
    $worksheet3 = findWorksheetOrFail($spreadsheet3, $abaArquivo3, 'terceira planilha');

    // Monta o índice considerando o filtro de códigos manuais (se informado) ou a coluna toda
    $indiceReferencia = buildWorksheetIndex(
        $worksheet3,
        $processingConfig['reference_compare_column'],
        (int)$processingConfig['reference_start_row'],
        $codigosEspecificos
    );

    $spreadsheetBase = loadSpreadsheet($basePath);
    $worksheetBase = findWorksheetOrFail($spreadsheetBase, $abaBase, 'planilha base');

    $ultimaLinhaBase = $worksheetBase->getHighestDataRow(
        $processingConfig['base_compare_column']
    );

    $totalLinhas = 0;
    $correspondencias = 0;
    $semCorrespondencia = 0;
    $celulasVazias = 0;
    $semRegraAplicavel = 0;
    $inconsistencias = 0;
    $estatisticasRegras = buildRuleStatistics($treatmentRules);

    for ($linha = (int)$processingConfig['base_start_row']; $linha <= $ultimaLinhaBase; $linha++) {
        $totalLinhas++;
        $coordinate = $processingConfig['base_compare_column'] . $linha;

        if (!$worksheetBase->cellExists($coordinate)) {
            $celulasVazias++;
            continue;
        }

        $codigoNormalizado = normalizeCode(
            $worksheetBase->getCell($coordinate)->getValue()
        );

        if ($codigoNormalizado === '') {
            $celulasVazias++;
            continue;
        }

        if (!array_key_exists($codigoNormalizado, $indiceReferencia)) {
            $semCorrespondencia++;

            if ($textoSemTratamento !== '') {
                setSpreadsheetCellAsString(
                    $worksheetBase,
                    $processingConfig['result_column'] . $linha,
                    $textoSemTratamento
                );
            }
            continue;
        }

        $correspondencias++;

        $resultadoRegras = evaluateTreatmentRules(
            $worksheet3,
            $indiceReferencia[$codigoNormalizado],
            $treatmentRules
        );

        if ($resultadoRegras['status'] === 'no_rule' || $resultadoRegras['status'] === 'inconsistent') {
            if ($resultadoRegras['status'] === 'no_rule') {
                $semRegraAplicavel++;
            } else {
                $inconsistencias++;
            }

            if ($textoSemTratamento !== '') {
                setSpreadsheetCellAsString(
                    $worksheetBase,
                    $processingConfig['result_column'] . $linha,
                    $textoSemTratamento
                );
            }
            continue;
        }

        $regraAplicada = $resultadoRegras['matches'][0];
        $textoTratamento = generateTreatmentText(
            $regraAplicada['rule'],
            $regraAplicada['cell_value']
        );

        setSpreadsheetCellAsString(
            $worksheetBase,
            $processingConfig['result_column'] . $linha,
            $textoTratamento
        );

        if (isset($estatisticasRegras[$regraAplicada['rule']['index']])) {
            $estatisticasRegras[$regraAplicada['rule']['index']]['count']++;
        }
    }

    $outputFilename = 'arquivo_processado_etapa2_' . date('Ymd_His') . '_' . generateToken() . '.xlsx';
    $outputPath = storagePath('output') . DIRECTORY_SEPARATOR . $outputFilename;

    $writer = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheetBase, 'Xlsx');
    if (method_exists($writer, 'setIncludeCharts')) {
        $writer->setIncludeCharts(true);
    }

    $writer->save($outputPath);
    @chmod($outputPath, 0600);

    $spreadsheet3->disconnectWorksheets();
    unset($spreadsheet3);

    $spreadsheetBase->disconnectWorksheets();
    unset($spreadsheetBase);

    cleanupStageTwoSessionFiles();

    // Atualiza a sessão com os dados do resultado da Etapa 2
    $_SESSION['resultado'] = [
        'stage' => 'etapa_2',
        'filename' => $outputFilename,
        'path' => $outputPath,
        'total_linhas' => $totalLinhas,
        'correspondencias' => $correspondencias,
        'sem_correspondencia' => $semCorrespondencia,
        'celulas_vazias' => $celulasVazias,
        'sem_regra_aplicavel' => $semRegraAplicavel,
        'inconsistencias' => $inconsistencias,
        'aba_arquivo1' => $abaBase,
        'aba_arquivo2' => $abaArquivo3,
        'stage2' => true,
    ];

} catch (Throwable $e) {
    $error = $e->getMessage();
}

$resultado = $_SESSION['resultado'] ?? null;
?>
<!DOCTYPE html>
<html lang="pt-BR">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Resultado - Etapa 2</title>
    <link rel="stylesheet" href="assets/style.css">
</head>
<body>

<div class="container">
    <?php if ($error !== null): ?>
        <div class="card">
            <div class="alert alert-error">
                <strong>Não foi possível processar a etapa 2.</strong>
                <p><?= htmlEscape($error) ?></p>
            </div>
            <div class="actions">
                <a href="tratar_repetidos.php" class="button">Voltar</a>
                <a href="index.php" class="button button-secondary">Início</a>
            </div>
        </div>
    <?php elseif ($resultado !== null): ?>
        <div class="card">
            <div class="success-icon">✓</div>
            <h1>Processamento de Repetidos Concluído</h1>

            <div class="summary">
                <div class="summary-item">
                    <span>Total de linhas analisadas</span>
                    <strong><?= number_format($resultado['total_linhas'], 0, ',', '.') ?></strong>
                </div>
                <div class="summary-item">
                    <span>Tratamentos aplicados</span>
                    <strong><?= number_format($resultado['correspondencias'], 0, ',', '.') ?></strong>
                </div>
                <div class="summary-item">
                    <span>Sem correspondência</span>
                    <strong><?= number_format($resultado['sem_correspondencia'], 0, ',', '.') ?></strong>
                </div>
            </div>

            <div class="actions">
                <a href="download.php" class="button button-success">Baixar arquivo atualizado</a>
                <a href="index.php" class="button button-secondary">Novo processamento</a>
            </div>
        </div>
    <?php endif; ?>
</div>

</body>
</html>
