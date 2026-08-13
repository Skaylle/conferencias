<?php

declare(strict_types=1);

session_start();

require __DIR__ . '/vendor/autoload.php';
require __DIR__ . '/functions.php';

ensureStorageDirectories();

/*
 * Processamento de planilhas grandes pode exigir mais memória/tempo.
 * Em hosts que permitem ini_set, elevamos os limites desta requisição.
 */
@ini_set('memory_limit', '1024M');
@set_time_limit(300);

$error = null;

try {
    if ($_SERVER['REQUEST_METHOD'] !== 'POST') {
        throw new RuntimeException(
            'Método de acesso inválido.'
        );
    }

    $abaArquivo1 = trim(
        (string)($_POST['aba_arquivo1'] ?? '')
    );

    $abaArquivo2 = trim(
        (string)($_POST['aba_arquivo2'] ?? '')
    );

    if ($abaArquivo1 === '') {
        throw new RuntimeException(
            'Selecione a aba do Arquivo 1.'
        );
    }

    if ($abaArquivo2 === '') {
        throw new RuntimeException(
            'Selecione a aba do Arquivo 2.'
        );
    }

    $arquivo1Path = $_SESSION['arquivo1_path'] ?? null;
    $arquivo2Path = $_SESSION['arquivo2_path'] ?? null;

    if (
        !is_string($arquivo1Path) ||
        !is_file($arquivo1Path)
    ) {
        throw new RuntimeException(
            'O Arquivo 1 não está mais disponível. Faça o upload novamente.'
        );
    }

    if (
        !is_string($arquivo2Path) ||
        !is_file($arquivo2Path)
    ) {
        throw new RuntimeException(
            'O Arquivo 2 não está mais disponível. Faça o upload novamente.'
        );
    }


    /*
     * ============================================================
     * 1. CARREGAR ARQUIVO 2
     * ============================================================
     */

    $spreadsheet2 = loadSpreadsheet(
        $arquivo2Path,
        $abaArquivo2,
        true
    );

    try {
        $worksheet2 = $spreadsheet2->getSheetByName($abaArquivo2);
    } catch (Throwable $e) {
        $worksheet2 = null;
    }

    if ($worksheet2 === null) {
        throw new RuntimeException(
            'A aba selecionada do Arquivo 2 não foi encontrada.'
        );
    }


    /*
     * ============================================================
     * 2. CONSTRUIR ÍNDICE DO ARQUIVO 2
     * ============================================================
     *
     * C = COD_ASSUNTO
     * D = DESCR_ASSUNTO
     */

    $indiceAssuntos = [];

    $ultimaLinhaArquivo2 =
        $worksheet2->getHighestDataRow('C');

    for (
        $linha = 1;
        $linha <= $ultimaLinhaArquivo2;
        $linha++
    ) {
        if (!$worksheet2->cellExists(ARQUIVO2_COLUNA_CODIGO . $linha)) {
            continue;
        }

        $codigo = $worksheet2
            ->getCell(ARQUIVO2_COLUNA_CODIGO . $linha)
            ->getValue();

        $descricao = $worksheet2
            ->getCell(ARQUIVO2_COLUNA_DESCRICAO . $linha)
            ->getValue();

        $codigoNormalizado = normalizeCode($codigo);

        if ($codigoNormalizado === '') {
            continue;
        }

        /*
         * Se houver código duplicado na referência,
         * a primeira ocorrência será utilizada.
         */
        if (!array_key_exists($codigoNormalizado, $indiceAssuntos)) {
            $indiceAssuntos[$codigoNormalizado] =
                (string)$descricao;
        }
    }


    /*
     * O Arquivo 2 não será mais necessário.
     */
    $spreadsheet2->disconnectWorksheets();
    unset($spreadsheet2);


    /*
     * ============================================================
     * 3. CARREGAR ARQUIVO 1
     * ============================================================
     */

    $spreadsheet1 = loadSpreadsheet($arquivo1Path);

    try {
        $worksheet1 = $spreadsheet1->getSheetByName($abaArquivo1);
    } catch (Throwable $e) {
        $worksheet1 = null;
    }

    if ($worksheet1 === null) {
        throw new RuntimeException(
            'A aba selecionada do Arquivo 1 não foi encontrada.'
        );
    }


    /*
     * ============================================================
     * 4. PROCESSAR COLUNA B
     * ============================================================
     *
     * Regra adicional:
     * quando B estiver vazia, tentamos usar o código da coluna A.
     */

    $ultimaLinhaArquivo1 = max(
        $worksheet1->getHighestDataRow(ARQUIVO1_COLUNA_CODIGO),
        $worksheet1->getHighestDataRow(ARQUIVO1_COLUNA_FALLBACK)
    );

    $totalLinhas = 0;
    $correspondencias = 0;
    $semCorrespondencia = 0;
    $celulasVazias = 0;

    for (
        $linha = 1;
        $linha <= $ultimaLinhaArquivo1;
        $linha++
    ) {
        $totalLinhas++;

        $coordinate =
            ARQUIVO1_COLUNA_CODIGO . $linha;

        $valorOriginal = null;

        if ($worksheet1->cellExists($coordinate)) {
            $valorOriginal = $worksheet1
                ->getCell($coordinate)
                ->getValue();
        }

        $usaFallback =
            $valorOriginal === null ||
            trim((string)$valorOriginal) === '';

        $codigoParaBusca = $valorOriginal;

        if ($usaFallback) {
            $fallbackCoordinate =
                ARQUIVO1_COLUNA_FALLBACK . $linha;

            $codigoParaBusca = null;

            if ($worksheet1->cellExists($fallbackCoordinate)) {
                $codigoParaBusca = $worksheet1
                    ->getCell($fallbackCoordinate)
                    ->getValue();
            }
        }

        /*
         * 1. Se não houver código nem em B nem em A, ignora.
         */
        if (
            $codigoParaBusca === null ||
            trim((string)$codigoParaBusca) === ''
        ) {
            $celulasVazias++;
            continue;
        }

        /*
         * 2. Normalizar código para busca na planilha 2.
         */
        $codigoNormalizado =
            normalizeCode($codigoParaBusca);

        if ($codigoNormalizado === '') {
            $celulasVazias++;
            continue;
        }

        /*
         * 3. Procurar no índice do Arquivo 2.
         */
        if (
            array_key_exists(
                $codigoNormalizado,
                $indiceAssuntos
            )
        ) {
            $descricao =
                $indiceAssuntos[$codigoNormalizado];

            /*
             * 4. Substituir SOMENTE a célula B.
             *
             * O estilo da célula permanece.
             * Nenhuma outra célula é modificada.
             */
            setSpreadsheetCellAsString(
                $worksheet1,
                $coordinate,
                $descricao
            );

            $correspondencias++;
        } else {
            /*
             * 5. Não encontrou:
             * mantém o valor original.
             */
            $semCorrespondencia++;
        }
    }


    /*
     * ============================================================
     * 5. GERAR NOVO ARQUIVO
     * ============================================================
     */

    $token = generateToken();

    $outputFilename =
        'arquivo_processado_' .
        date('Ymd_His') .
        '_' .
        $token .
        '.xlsx';

    $outputPath =
        storagePath('output') .
        DIRECTORY_SEPARATOR .
        $outputFilename;


    /*
     * Nunca usamos o caminho do Arquivo 1 como saída.
     */
    $writer = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter(
        $spreadsheet1,
        'Xlsx'
    );

    /*
     * Se a planilha possuir gráficos suportados pelo PhpSpreadsheet,
     * tentamos preservá-los também.
     */
    if (method_exists($writer, 'setIncludeCharts')) {
        $writer->setIncludeCharts(true);
    }

    $writer->save($outputPath);

    @chmod($outputPath, 0600);


    /*
     * Liberar memória.
     */
    $spreadsheet1->disconnectWorksheets();
    unset($spreadsheet1);

    /*
     * Apagar os arquivos temporários dos uploads.
     */
    cleanupFile($arquivo1Path);
    cleanupFile($arquivo2Path);

    unset(
        $_SESSION['arquivo1_path'],
        $_SESSION['arquivo2_path'],
        $_SESSION['arquivo1_original_name'],
        $_SESSION['arquivo2_original_name']
    );


    /*
     * Guardamos o resultado na sessão.
     */
    $_SESSION['resultado'] = [
        'filename' => $outputFilename,
        'path' => $outputPath,
        'total_linhas' => $totalLinhas,
        'correspondencias' => $correspondencias,
        'sem_correspondencia' => $semCorrespondencia,
        'celulas_vazias' => $celulasVazias,
        'aba_arquivo1' => $abaArquivo1,
        'aba_arquivo2' => $abaArquivo2,
    ];

} catch (Throwable $e) {
    $error = $e->getMessage();

    cleanupFile(
        $_SESSION['arquivo1_path'] ?? null
    );

    cleanupFile(
        $_SESSION['arquivo2_path'] ?? null
    );

    unset(
        $_SESSION['arquivo1_path'],
        $_SESSION['arquivo2_path'],
        $_SESSION['arquivo1_original_name'],
        $_SESSION['arquivo2_original_name']
    );
}

$resultado = $_SESSION['resultado'] ?? null;
?>
<!DOCTYPE html>
<html lang="pt-BR">
<head>
    <meta charset="UTF-8">

    <meta
        name="viewport"
        content="width=device-width, initial-scale=1.0"
    >

    <title>Resultado - Cruzamento de Excel</title>

    <link
        rel="stylesheet"
        href="assets/style.css"
    >
</head>

<body>

<div class="container">

    <?php if ($error !== null): ?>

        <div class="card">

            <div class="alert alert-error">

                <strong>Não foi possível processar o arquivo.</strong>

                <p>
                    <?= htmlEscape($error) ?>
                </p>

            </div>

            <a
                href="index.php"
                class="button"
            >
                Voltar
            </a>

        </div>

    <?php elseif ($resultado !== null): ?>

        <div class="card">

            <div class="success-icon">
                ✓
            </div>

            <h1>Processamento concluído.</h1>

            <p class="subtitle">
                O arquivo original não foi sobrescrito.
            </p>


            <div class="summary">

                <div class="summary-item">

                    <span>
                        Total de linhas verificadas
                    </span>

                    <strong>
                        <?= number_format(
                            $resultado['total_linhas'],
                            0,
                            ',',
                            '.'
                        ) ?>
                    </strong>

                </div>


                <div class="summary-item">

                    <span>
                        Correspondências encontradas
                    </span>

                    <strong>
                        <?= number_format(
                            $resultado['correspondencias'],
                            0,
                            ',',
                            '.'
                        ) ?>
                    </strong>

                </div>


                <div class="summary-item">

                    <span>
                        Sem correspondência
                    </span>

                    <strong>
                        <?= number_format(
                            $resultado['sem_correspondencia'],
                            0,
                            ',',
                            '.'
                        ) ?>
                    </strong>

                </div>


                <div class="summary-item">

                    <span>
                        Células vazias ignoradas
                    </span>

                    <strong>
                        <?= number_format(
                            $resultado['celulas_vazias'],
                            0,
                            ',',
                            '.'
                        ) ?>
                    </strong>

                </div>

            </div>


            <div class="details">

                <p>
                    <strong>Aba do Arquivo 1:</strong>
                    <?= htmlEscape($resultado['aba_arquivo1']) ?>
                </p>

                <p>
                    <strong>Aba do Arquivo 2:</strong>
                    <?= htmlEscape($resultado['aba_arquivo2']) ?>
                </p>

            </div>


            <div class="actions">

                <a
                    href="download.php"
                    class="button button-success"
                >
                    Baixar arquivo processado
                </a>

                <a 
                    href="tratar_repetidos.php" 
                    class="button button-primary"
                >
                    Avançar para Etapa 2 (Tratar Repetidos)
                </a>

                <a
                    href="index.php"
                    class="button button-secondary"
                >
                    Processar outro arquivo
                </a>

            </div>

        </div>

    <?php endif; ?>

</div>

</body>
</html>
