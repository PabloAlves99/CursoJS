function compararEstoqueCopia() {
    let ss = SpreadsheetApp.getActiveSpreadsheet();
    let abaEstoque = ss.getSheetByName("estoque"); // Aba de Estoque com os clientes válidos
    let abaCopiaInicial = ss.getSheetByName("INICIAL"); // Aba que precisa ser comparada

    if (!abaEstoque || !abaCopiaInicial) {
        Logger.log("❌ Verifique se as abas 'Estoque' e 'Cópia de INICIAL' existem.");
        return;
    };

    let dadosEstoque = abaEstoque.getDataRange().getValues(); // Dados da aba Estoque
    let dadosCopia = abaCopiaInicial.getDataRange().getValues(); // Dados da aba Cópia de INICIAL

    let clientesCopia = new Set(); // Conjunto para armazenar clientes da Cópia de INICIAL
    let naoEncontrados = []; // Lista para armazenar clientes que não estão na Cópia de INICIAL

    // Adiciona todos os clientes da aba Cópia de INICIAL ao conjunto (coluna K)
    for (let i = 1; i < dadosCopia.length; i++) {
        let clienteCopia = limparTexto(dadosCopia[i][10]); // Coluna K da aba Cópia de INICIAL
        if (clienteCopia) {
            clientesCopia.add(clienteCopia);
        };
    };

    // Compara cada cliente da aba Estoque
    for (let i = 1; i < dadosEstoque.length; i++) {
        let clienteEstoque = limparTexto(dadosEstoque[i][0]); // Coluna A da aba Estoque
        if (clienteEstoque && !clientesCopia.has(clienteEstoque)) {
            naoEncontrados.push(clienteEstoque); // Adiciona ao array os clientes não encontrados
        };
    };

    // Exibe o log com os resultados
    if (naoEncontrados.length > 0) {
        Logger.log(`❌ Clientes encontrados em Estoque, mas não em Cópia de INICIAL:`);
        naoEncontrados.forEach(cliente => Logger.log(cliente));
    } else {
        Logger.log("✅ Todos os clientes de Estoque estão presentes em Cópia de INICIAL.");
    };
};

// 🔹 Função para limpar texto (remover espaços extras e padronizar letras)
function limparTexto(texto) {
    return texto ? texto.toString().trim().replace(/\s+/g, ' ').toLowerCase() : "";
};
