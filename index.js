const inputPlanilhaAntiga = document.getElementById("input_planilha_anterior");
const inputPlanilhaNova = document.getElementById("input_planilha_atualizada");

async function lerWorkbook(input) {
    const arquivo = input.files[0];

    if (!arquivo) {
        throw new Error("Selecione uma planilha.");
    }

    const buffer = await arquivo.arrayBuffer();

    return XLSX.read(buffer, {
        type: "array"
    });
}

function normalizar(valor) {
    return String(valor ?? "").trim();
}

async function atualizarPlanilha() {

    try {

        // Lê os workbooks
        const wbAntiga = await lerWorkbook(inputPlanilhaAntiga);
        const wbNova = await lerWorkbook(inputPlanilhaNova);

        // Primeira aba
        const wsAntiga =
            wbAntiga.Sheets[wbAntiga.SheetNames[0]];

        const wsNova =
            wbNova.Sheets[wbNova.SheetNames[0]];

        // Converte para JSON
        const antiga =
            XLSX.utils.sheet_to_json(wsAntiga, {
                defval: ""
            });

        const nova =
            XLSX.utils.sheet_to_json(wsNova, {
                defval: ""
            });

        // Índice da planilha antiga
        const indice = new Map();

        for (const linha of antiga) {

            const doc = normalizar(linha["Nº doc"]);

            if (!doc) continue;

            indice.set(doc, {
                status: linha["Status de confirmação"],
                posicao: linha["POSICAO"]
            });

        }

        let atualizados = 0;

        // Atualiza a nova
        for (const linha of nova) {

            const doc = normalizar(linha["Nº doc"]);

            if (!indice.has(doc)) {
                continue;
            }

            const antigo = indice.get(doc);

            linha["Status de confirmação"] =
                antigo.status;

            linha["POSICAO"] =
                antigo.posicao;

            atualizados++;

        }

        // Gera a planilha preservando a nova
        const wsResultado =
            XLSX.utils.json_to_sheet(nova);

        const wbResultado =
            XLSX.utils.book_new();

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
            `${atualizados} registros atualizados com sucesso.`
        );

    } catch (e) {

        console.error(e);

        alert(e.message);

    }

}