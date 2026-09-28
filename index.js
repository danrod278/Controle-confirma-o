const inputPlanilhaAntiga = document.getElementById("input_planilha_anterior");
const inputPlanilhaNova = document.getElementById("input_planilha_atualizada");

// Lê uma planilha XLSX
async function lerWorkbook(input) {
    const arquivo = input.files[0];

    if (!arquivo) {
        throw new Error("Selecione uma planilha.");
    }

    if (!arquivo.name.toLowerCase().endsWith(".xlsx")) {
        throw new Error("Apenas arquivos .xlsx são permitidos.");
    }

    const buffer = await arquivo.arrayBuffer();

    return XLSX.read(buffer, {
        type: "array"
    });
}

// Normaliza os valores para comparação
function normalizar(valor) {
    return String(valor ?? "").trim();
}

// Cria a chave usando:
// Nº doc + CNPJ/CPF Cedente
function criarChave(linha) {
    return (
        normalizar(linha["Nº doc"]) +
        "|" +
        normalizar(linha["CNPJ/CPF Cedente"])
    );
}

async function atualizarPlanilha() {
    try {
        // =====================================================
        // LÊ AS DUAS PLANILHAS
        // =====================================================

        const wbAntiga = await lerWorkbook(inputPlanilhaAntiga);
        const wbNova = await lerWorkbook(inputPlanilhaNova);

        // Pega a primeira aba de cada arquivo
        const wsAntiga = wbAntiga.Sheets[wbAntiga.SheetNames[0]];
        const wsNova = wbNova.Sheets[wbNova.SheetNames[0]];

        // Converte as planilhas para objetos JavaScript
        const antiga = XLSX.utils.sheet_to_json(wsAntiga, {
            defval: ""
        });

        const nova = XLSX.utils.sheet_to_json(wsNova, {
            defval: ""
        });

        // =====================================================
        // CRIA ÍNDICE DA PLANILHA ANTIGA
        // =====================================================

        const indice = new Map();

        for (const linha of antiga) {
            const chave = criarChave(linha);

            indice.set(chave, {
                obs: linha["Obs"]
            });
        }

        // =====================================================
        // ATUALIZA A PLANILHA NOVA
        // =====================================================

        let atualizados = 0;

        for (const linha of nova) {

            // Cria a chave:
            // Nº doc + CNPJ/CPF Cedente
            const chave = criarChave(linha);

            // Procura o título na planilha antiga
            const dadosAntigos = indice.get(chave);

            // Se não encontrou, não altera a linha
            if (!dadosAntigos) {
                continue;
            }

            // =================================================
            // STATUS DE CONFIRMAÇÃO
            // =================================================
            // NÃO ALTERAMOS.
            // O valor permanece exatamente como está
            // na planilha NOVA.
            //
            // =================================================

            // =================================================
            // OBS
            // =================================================
            // Copia a Obs da planilha antiga para a nova.
            //
            linha["Obs"] = dadosAntigos.obs;

            atualizados++;
        }

        // =====================================================
        // GERA A PLANILHA FINAL
        // =====================================================

        const wsResultado = XLSX.utils.json_to_sheet(nova);

        const wbResultado = XLSX.utils.book_new();

        XLSX.utils.book_append_sheet(
            wbResultado,
            wsResultado,
            "Atualizada"
        );

        // Faz o download
        XLSX.writeFile(
            wbResultado,
            "planilha_atualizada.xlsx"
        );

        alert(
            `Concluído! ${atualizados} registros atualizados.`
        );

    } catch (erro) {
        console.error(erro);

        alert(
            erro.message ||
            "Ocorreu um erro ao processar as planilhas."
        );
    }
}

