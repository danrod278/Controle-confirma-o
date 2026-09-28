javascript
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

// Remove espaços extras
function normalizar(valor) {
    return String(valor ?? "").trim();
}

// Cria uma chave única usando Nº doc + CNPJ/CPF Cedente
function criarChave(linha) {
    return (
        normalizar(linha["Nº doc"]) +
        "|" +
        normalizar(linha["CNPJ/CPF Cedente"])
    );
}

async function atualizarPlanilha() {
    try {
        const wbAntiga = await lerWorkbook(inputPlanilhaAntiga);
        const wbNova = await lerWorkbook(inputPlanilhaNova);

        const wsAntiga = wbAntiga.Sheets[wbAntiga.SheetNames[0]];
        const wsNova = wbNova.Sheets[wbNova.SheetNames[0]];

        const antiga = XLSX.utils.sheet_to_json(wsAntiga, {
            defval: ""
        });

        const nova = XLSX.utils.sheet_to_json(wsNova, {
            defval: ""
        });

        // Índice da planilha antiga
        const indice = new Map();

        for (const linha of antiga) {
            const chave = criarChave(linha);

            indice.set(chave, {
                status: linha["Status de confirmação"],
                posicao: linha["POSICAO"]
            });
        }

        let atualizados = 0;

        // Atualiza a planilha nova
        for (const linha of nova) {

            const chave = criarChave(linha);

            const dadosAntigos = indice.get(chave);

            if (!dadosAntigos) {
                continue;
            }

            linha["Status de confirmação"] =
                dadosAntigos.status;

            linha["POSICAO"] =
                dadosAntigos.posicao;

            atualizados++;
        }

        // Salva a nova planilha
        const wsResultado = XLSX.utils.json_to_sheet(nova);

        const wbResultado = XLSX.utils.book_new();

        XLSX.utils.book_append_sheet(
            wbResultado,
            wsResultado,
            "Atualizada"
        );

        XLSX.writeFile(
            wbResultado,
            "planilha_atualizada.xlsx"
        );

        alert(
            `Concluído! ${atualizados} registros atualizados.`
        );

    } catch (erro) {
        console.error(erro);
        alert(erro.message);
    }
}

