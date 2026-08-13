<?php

declare(strict_types=1);

session_start();

require __DIR__ . '/vendor/autoload.php';
require __DIR__ . '/functions.php';

ensureStorageDirectories();

cleanupOldFiles('tmp');
cleanupOldFiles('output');

$error = null;

$arquivo1 = null;
$arquivo2 = null;

$abasArquivo1 = [];
$abasArquivo2 = [];

// Limpa todos os arquivos das pastas tmp e output ao acessar a index
cleanupOldFiles('tmp', 0);
cleanupOldFiles('output', 0);

// Limpar também os dados de resultado gravados na sessão
clearSession();

if ($_SERVER['REQUEST_METHOD'] === 'POST') {
    try {
        $acao = $_POST['acao'] ?? '';

        if ($acao === 'upload') {
            if (
                !isset($_FILES['arquivo1']) ||
                !isset($_FILES['arquivo2'])
            ) {
                throw new RuntimeException(
                    'Selecione os dois arquivos Excel.'
                );
            }

            validateUploadedExcel(
                $_FILES['arquivo1'],
                'Arquivo 1'
            );

            validateUploadedExcel(
                $_FILES['arquivo2'],
                'Arquivo 2'
            );

            /*
             * Limpa arquivos temporários de uma sessão anterior.
             */
            cleanupFile($_SESSION['arquivo1_path'] ?? null);
            cleanupFile($_SESSION['arquivo2_path'] ?? null);

            $token = generateToken();

            $arquivo1Path = saveUploadedFile(
                $_FILES['arquivo1'],
                $token,
                'arquivo1'
            );

            try {
                $arquivo2Path = saveUploadedFile(
                    $_FILES['arquivo2'],
                    $token,
                    'arquivo2'
                );
            } catch (Throwable $e) {
                cleanupFile($arquivo1Path);
                throw $e;
            }

            try {
                $abasArquivo1 = getWorksheetNames($arquivo1Path);
                $abasArquivo2 = getWorksheetNames($arquivo2Path);
            } catch (Throwable $e) {
                cleanupFile($arquivo1Path);
                cleanupFile($arquivo2Path);
                throw $e;
            }

            if (count($abasArquivo1) === 0) {
                throw new RuntimeException(
                    'O Arquivo 1 não possui abas disponíveis.'
                );
            }

            if (count($abasArquivo2) === 0) {
                throw new RuntimeException(
                    'O Arquivo 2 não possui abas disponíveis.'
                );
            }

            $_SESSION['arquivo1_path'] = $arquivo1Path;
            $_SESSION['arquivo2_path'] = $arquivo2Path;
            $_SESSION['arquivo1_original_name'] =
                basename((string)$_FILES['arquivo1']['name']);
            $_SESSION['arquivo2_original_name'] =
                basename((string)$_FILES['arquivo2']['name']);

            $arquivo1 = [
                'name' => $_SESSION['arquivo1_original_name'],
            ];

            $arquivo2 = [
                'name' => $_SESSION['arquivo2_original_name'],
            ];
        }
    } catch (Throwable $e) {
        $error = $e->getMessage();
    }
}

if (
    isset($_SESSION['arquivo1_path']) &&
    is_file($_SESSION['arquivo1_path'])
) {
    $arquivo1 = [
        'name' => $_SESSION['arquivo1_original_name'] ?? 'Arquivo 1',
    ];

    if (empty($abasArquivo1)) {
        try {
            $abasArquivo1 = getWorksheetNames(
                $_SESSION['arquivo1_path']
            );
        } catch (Throwable $e) {
            $error = $e->getMessage();
        }
    }
}

if (
    isset($_SESSION['arquivo2_path']) &&
    is_file($_SESSION['arquivo2_path'])
) {
    $arquivo2 = [
        'name' => $_SESSION['arquivo2_original_name'] ?? 'Arquivo 2',
    ];

    if (empty($abasArquivo2)) {
        try {
            $abasArquivo2 = getWorksheetNames(
                $_SESSION['arquivo2_path']
            );
        } catch (Throwable $e) {
            $error = $e->getMessage();
        }
    }
}
?>
<!DOCTYPE html>
<html lang="pt-BR">
<head>
    <meta charset="UTF-8">

    <meta
        name="viewport"
        content="width=device-width, initial-scale=1.0"
    >

    <title>Cruzamento de Excel</title>

    <link
        rel="stylesheet"
        href="assets/style.css"
    >
</head>

<body>

<div class="container">

    <div class="card">

        <h1>Cruzamento de códigos Excel</h1>

        <p class="subtitle">
            Substitui os códigos da coluna B do Arquivo 1
            pelas descrições encontradas no Arquivo 2.
        </p>

        <?php if ($error !== null): ?>

            <div class="alert alert-error">
                <?= htmlEscape($error) ?>
            </div>

        <?php endif; ?>

        <form
            method="post"
            enctype="multipart/form-data"
        >

            <input
                type="hidden"
                name="acao"
                value="upload"
            >

            <div class="form-group">

                <label for="arquivo1">
                    Arquivo 1
                </label>

                <input
                    type="file"
                    id="arquivo1"
                    name="arquivo1"
                    accept=".xlsx,.xls"
                    required
                >

                <small>
                    Templete já carregado com os dados do Eproc.
                    O original não será sobrescrito.
                </small>

            </div>

            <div class="form-group">

                <label for="arquivo2">
                    Arquivo 2
                </label>

                <input
                    type="file"
                    id="arquivo2"
                    name="arquivo2"
                    accept=".xlsx,.xls"
                    required
                >

                <small>
                    Planilha com a extração do DCP
                    COD_XXX na coluna C e DESCR_XXX na coluna D.
                </small>

            </div>

            <button
                type="submit"
                class="button"
            >
                Carregar arquivos
            </button>

        </form>

    </div>


    <?php if (
        !empty($abasArquivo1) &&
        !empty($abasArquivo2) &&
        $arquivo1 !== null &&
        $arquivo2 !== null
    ): ?>

        <div class="card">

            <h2>Selecionar abas</h2>

            <div class="file-info">

                <strong>Arquivo 1:</strong>

                <?= htmlEscape($arquivo1['name']) ?>

            </div>

            <div class="file-info">

                <strong>Arquivo 2:</strong>

                <?= htmlEscape($arquivo2['name']) ?>

            </div>

            <form
                method="post"
                action="processar.php"
            >

                <div class="form-group">

                    <label for="aba_arquivo1">
                        Aba do Arquivo 1
                    </label>

                    <select
                        id="aba_arquivo1"
                        name="aba_arquivo1"
                        required
                    >

                        <option value="">
                            Selecione uma aba
                        </option>

                        <?php foreach ($abasArquivo1 as $aba): ?>

                            <option
                                value="<?= htmlEscape($aba) ?>"
                            >
                                <?= htmlEscape($aba) ?>
                            </option>

                        <?php endforeach; ?>

                    </select>

                    <small>
                        Nesta aba será processada a coluna B.
                        Se B estiver vazia, será usado o código da coluna A.
                    </small>

                </div>


                <div class="form-group">

                    <label for="aba_arquivo2">
                        Aba do Arquivo 2
                    </label>

                    <select
                        id="aba_arquivo2"
                        name="aba_arquivo2"
                        required
                    >

                        <option value="">
                            Selecione uma aba
                        </option>

                        <?php foreach ($abasArquivo2 as $aba): ?>

                            <option
                                value="<?= htmlEscape($aba) ?>"
                            >
                                <?= htmlEscape($aba) ?>
                            </option>

                        <?php endforeach; ?>

                    </select>

                    <small>
                        Nesta aba serão utilizadas as colunas C e D.
                    </small>

                </div>


                <div class="rules">

                    <h3>Regras utilizadas</h3>

                    <ul>
                        <li>
                            Arquivo 1: coluna B (com fallback na coluna A quando B estiver vazia).
                        </li>

                        <li>
                            Arquivo 2: coluna C = Código da competência Ex: COD_CLASSE.
                        </li>

                        <li>
                            Arquivo 2: coluna D = Descrição da competência Ex: DESCR_CLASSE.
                        </li>

                        <li>
                            Células vazias serão ignoradas.
                        </li>

                        <li>
                            Códigos sem correspondência permanecerão inalterados.
                        </li>

                        <li>
                            Nenhuma outra coluna será modificada.
                        </li>

                        <li>
                            O Arquivo 1 original nunca será sobrescrito.
                        </li>

                    </ul>

                </div>


                <button
                    type="submit"
                    class="button button-success"
                >
                    Processar arquivo
                </button>

            </form>

        </div>

    <?php endif; ?>

</div>

</body>
</html>
