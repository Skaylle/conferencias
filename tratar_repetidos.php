<?php

declare(strict_types=1);

session_start();

require __DIR__ . '/vendor/autoload.php';
require __DIR__ . '/functions.php';

ensureStorageDirectories();
cleanupOldFiles('tmp');
cleanupOldFiles('output');

$error = null;
$resultadoBase = $_SESSION['resultado'] ?? null;
$configEtapa2 = [];
$arquivo3 = null;
$abasArquivoBase = [];
$abasArquivo3 = [];
$basePath = null;
$nomeArquivo = null;

try {
    // 1. Localizar o arquivo base (Sessão ou arquivo mais recente em output/)
    if (is_array($resultadoBase) && !empty($resultadoBase['path']) && is_file((string)$resultadoBase['path'])) {
        $basePath = (string)$resultadoBase['path'];
        $nomeArquivo = basename($basePath);
    } else {
        $arquivos = glob(__DIR__ . '/output/*.xlsx');
        if (!empty($arquivos)) {
            usort($arquivos, fn($a, $b) => filemtime($b) <=> filemtime($a));
            $basePath = $arquivos[0];
            $nomeArquivo = basename($basePath);
        } else {
            $error = 'Nenhuma planilha base encontrada no sistema. Processe a Etapa 1 primeiro.';
        }
    }

    if ($basePath !== null) {
        $abasArquivoBase = getWorksheetNames($basePath);
        $configEtapa2 = getDefaultRepeatedTreatmentConfig($resultadoBase);
    }

    // 2. Processar Upload do Arquivo 3
    if ($_SERVER['REQUEST_METHOD'] === 'POST') {
        $acao = $_POST['acao'] ?? '';

        if ($acao === 'upload_arquivo3') {
            if (!isset($_FILES['arquivo3'])) {
                throw new RuntimeException('Selecione a terceira planilha.');
            }

            validateUploadedExcel($_FILES['arquivo3'], 'Terceira planilha');

            $token = generateToken();
            $arquivo3Path = saveUploadedFile($_FILES['arquivo3'], $token, 'arquivo3');

            try {
                $abasArquivo3 = getWorksheetNames($arquivo3Path);
            } catch (Throwable $e) {
                cleanupFile($arquivo3Path);
                throw $e;
            }

            if ($abasArquivo3 === []) {
                cleanupFile($arquivo3Path);
                throw new RuntimeException('A terceira planilha não possui abas disponíveis.');
            }

            $_SESSION['arquivo3_path'] = $arquivo3Path;
            $_SESSION['arquivo3_original_name'] = basename((string)$_FILES['arquivo3']['name']);
        }
    }

    // 3. Restaurar estado do Arquivo 3 se estiver na sessão
    if (
        isset($_SESSION['arquivo3_path']) &&
        is_file($_SESSION['arquivo3_path'])
    ) {
        $arquivo3 = [
            'name' => $_SESSION['arquivo3_original_name'] ?? 'Terceira planilha',
        ];

        if ($abasArquivo3 === []) {
            $abasArquivo3 = getWorksheetNames($_SESSION['arquivo3_path']);
        }
    }
} catch (Throwable $e) {
    $error = $e->getMessage();
}
?>

<!DOCTYPE html>
<html lang="pt-BR">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Tratar Repetidos - Etapa 2</title>
    <link rel="stylesheet" href="assets/style.css">
    <style>
        .rule-card-header {
            display: flex;
            justify-content: space-between;
            align-items: center;
            margin-bottom: 12px;
            font-weight: bold;
        }
        .btn-remove-rule {
            background-color: #dc3545;
            color: #fff;
            border: none;
            padding: 4px 10px;
            border-radius: 4px;
            cursor: pointer;
            font-size: 0.85rem;
        }
        .btn-remove-rule:hover {
            background-color: #bd2130;
        }
        .btn-add-rule {
            margin-top: 10px;
            margin-bottom: 20px;
        }
        .form-group-full {
            grid-column: 1 / -1;
        }
        textarea.code-input {
            width: 100%;
            min-height: 90px;
            font-family: monospace;
            padding: 8px;
            border: 1px solid #ccc;
            border-radius: 4px;
            resize: vertical;
        }
    </style>
</head>
<body>

<div class="container">

    <div class="card">
        <h1>Etapa 2: Tratar repetidos</h1>
        <p class="subtitle">
            Usa o arquivo gerado na etapa 1 como base na nova planilha (PJE ou DCP) para verificar se o registro já foi tratado.
        </p>

        <?php if ($error !== null): ?>
            <div class="alert alert-error">
                <?= htmlEscape($error) ?>
            </div>
        <?php endif; ?>

        <?php if ($nomeArquivo !== null): ?>
            <div class="file-info">
                <strong>Planilha base identificada:</strong>
                <?= htmlEscape($nomeArquivo) ?>
            </div>
        <?php endif; ?>

        <form method="post" enctype="multipart/form-data">
            <input type="hidden" name="acao" value="upload_arquivo3">

            <div class="form-group">
                <label for="arquivo3">Nova planilha (PJE / DCP)</label>
                <input type="file" id="arquivo3" name="arquivo3" accept=".xlsx,.xls" required>
                <small>Esta planilha será consultada para verificar os tratamentos das duplicidades.</small>
            </div>

            <button class="button button-success" type="submit">
                Carregar planilha
            </button>

            <a href="index.php" class="button button-secondary">Sair</a>
        </form>
    </div>

    <?php if (
        $basePath !== null &&
        !empty($abasArquivoBase) &&
        !empty($abasArquivo3) &&
        $arquivo3 !== null
    ): ?>

        <div class="card">
            <h2>Configuração do Cruzamento</h2>

            <div class="file-info">
                <strong>Nova planilha carregada:</strong> <?= htmlEscape($arquivo3['name']) ?>
            </div>

            <form method="post" action="processar_repetidos.php">

                <div class="form-group">
                    <label for="aba_base">Aba da planilha base</label>
                    <select id="aba_base" name="aba_base" required>
                        <option value="">Selecione uma aba</option>
                        <?php foreach ($abasArquivoBase as $aba): ?>
                            <option value="<?= htmlEscape($aba) ?>" <?= $aba === ($resultadoBase['aba_arquivo1'] ?? '') ? 'selected' : '' ?>>
                                <?= htmlEscape($aba) ?>
                            </option>
                        <?php endforeach; ?>
                    </select>
                </div>

                <div class="form-group">
                    <label for="aba_arquivo3">Aba da nova planilha</label>
                    <select id="aba_arquivo3" name="aba_arquivo3" required>
                        <option value="">Selecione uma aba</option>
                        <?php foreach ($abasArquivo3 as $aba): ?>
                            <option value="<?= htmlEscape($aba) ?>">
                                <?= htmlEscape($aba) ?>
                            </option>
                        <?php endforeach; ?>
                    </select>
                </div>

                <div class="rules">
                    <h3>Mapeamento de Colunas</h3>

                    <div class="config-grid">
                        <div class="form-group">
                            <label for="base_compare_column">Coluna chave da planilha base</label>
                            <input type="text" id="base_compare_column" name="base_compare_column" value="<?= htmlEscape($configEtapa2['base_compare_column'] ?? 'B') ?>" maxlength="4" required>
                        </div>

                        <div class="form-group">
                            <label for="base_start_row">Linha inicial da planilha base</label>
                            <input type="number" id="base_start_row" name="base_start_row" value="<?= htmlEscape((string)($configEtapa2['base_start_row'] ?? 1)) ?>" min="1" required>
                        </div>

                        <div class="form-group">
                            <label for="reference_compare_column">Coluna chave da nova planilha</label>
                            <input type="text" id="reference_compare_column" name="reference_compare_column" value="<?= htmlEscape($configEtapa2['reference_compare_column'] ?? 'A') ?>" maxlength="4" required>
                        </div>

                        <div class="form-group">
                            <label for="reference_start_row">Linha inicial da nova planilha</label>
                            <input type="number" id="reference_start_row" name="reference_start_row" value="<?= htmlEscape((string)($configEtapa2['reference_start_row'] ?? 1)) ?>" min="1" required>
                        </div>

                        <div class="form-group">
                            <label for="result_column">Coluna de destino do resultado</label>
                            <input type="text" id="result_column" name="result_column" value="<?= htmlEscape($configEtapa2['result_column'] ?? 'C') ?>" maxlength="4" required>
                        </div>

                        <div class="form-group form-group-full">
                            <label for="codigos_especificos">Filtro por Códigos Específicos (Opcional)</label>
                            <textarea id="codigos_especificos" name="codigos_especificos" class="code-input" placeholder="Cole os códigos aqui separando por linha, vírgula ou ponto e vírgula..."></textarea>
                            <small>Se informado, apenas estes códigos serão buscados no cruzamento com a nova planilha.</small>
                        </div>
                    </div>

                    <h3>Regras de Tratamento</h3>

                    <div id="rules-container">
                        <?php 
                        $rules = $configEtapa2['rules'] ?? [['column' => 'D', 'text' => 'Tratado em ', 'append_value' => true]];
                        foreach ($rules as $index => $rule): 
                        ?>
                            <div class="rule-card" data-index="<?= $index ?>">
                                <div class="rule-card-header">
                                    <span class="rule-title">Regra <?= $index + 1 ?></span>
                                    <button type="button" class="btn-remove-rule" onclick="removeRule(this)">Remover</button>
                                </div>

                                <div class="config-grid">
                                    <div class="form-group">
                                        <label>Coluna de verificação</label>
                                        <input type="text" name="rules[<?= $index ?>][column]" value="<?= htmlEscape($rule['column']) ?>" maxlength="4" required>
                                    </div>

                                    <div class="form-group form-group-wide">
                                        <label>Texto a ser gravado</label>
                                        <input type="text" name="rules[<?= $index ?>][text]" value="<?= htmlEscape($rule['text']) ?>">
                                    </div>
                                </div>

                                <label class="checkbox-field">
                                    <input type="checkbox" name="rules[<?= $index ?>][append_value]" value="1" <?= !empty($rule['append_value']) ? 'checked' : '' ?>>
                                    Anexar o valor da célula ao texto
                                </label>
                            </div>
                        <?php endforeach; ?>
                    </div>

                    <button type="button" class="button button-secondary btn-add-rule" onclick="addRule()">
                        + Adicionar Nova Regra
                    </button>
                </div>

                <button type="submit" class="button button-success">
                    Processar Repetidos
                </button>

            </form>
        </div>

    <?php endif; ?>

</div>

<script>
function addRule() {
    const container = document.getElementById('rules-container');
    const ruleCards = container.getElementsByClassName('rule-card');
    const newIndex = ruleCards.length;

    const card = document.createElement('div');
    card.className = 'rule-card';
    card.setAttribute('data-index', newIndex);

    card.innerHTML = `
        <div class="rule-card-header">
            <span class="rule-title">Regra ${newIndex + 1}</span>
            <button type="button" class="btn-remove-rule" onclick="removeRule(this)">Remover</button>
        </div>
        <div class="config-grid">
            <div class="form-group">
                <label>Coluna de verificação</label>
                <input type="text" name="rules[${newIndex}][column]" value="" maxlength="4" placeholder="Ex: E" required>
            </div>
            <div class="form-group form-group-wide">
                <label>Texto a ser gravado</label>
                <input type="text" name="rules[${newIndex}][text]" value="" placeholder="Ex: Arquivado em ">
            </div>
        </div>
        <label class="checkbox-field">
            <input type="checkbox" name="rules[${newIndex}][append_value]" value="1" checked>
            Anexar o valor da célula ao texto
        </label>
    `;

    container.appendChild(card);
    reindexRules();
}

function removeRule(btn) {
    const container = document.getElementById('rules-container');
    if (container.getElementsByClassName('rule-card').length <= 1) {
        alert('Você precisa ter pelo menos uma regra de tratamento.');
        return;
    }
    btn.closest('.rule-card').remove();
    reindexRules();
}

function reindexRules() {
    const cards = document.querySelectorAll('#rules-container .rule-card');
    cards.forEach((card, index) => {
        card.setAttribute('data-index', index);
        card.querySelector('.rule-title').innerText = `Regra ${index + 1}`;
        
        const columnInput = card.querySelector('input[name*="[column]"]');
        if (columnInput) columnInput.name = `rules[${index}][column]`;
        
        const textInput = card.querySelector('input[name*="[text]"]');
        if (textInput) textInput.name = `rules[${index}][text]`;
        
        const checkInput = card.querySelector('input[name*="[append_value]"]');
        if (checkInput) checkInput.name = `rules[${index}][append_value]`;
    });
}
</script>

</body>
</html>
