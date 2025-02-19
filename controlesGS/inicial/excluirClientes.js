function excluirClientesNaoEncontrados() {
    let ss = SpreadsheetApp.getActiveSpreadsheet();
    let abaEstoque = ss.getSheetByName("estoque"); // Aba de Estoque com os clientes válidos
    let abaCopiaInicial = ss.getSheetByName("INICIAL"); // Aba que precisa ser filtrada

    if (!abaEstoque || !abaCopiaInicial) {
        Logger.log("❌ Verifique se as abas 'Estoque' e 'Cópia de INICIAL' existem.");
        return;
    };

    let dadosEstoque = abaEstoque.getDataRange().getValues(); // Dados da aba Estoque
    let dadosCopia = abaCopiaInicial.getDataRange().getValues(); // Dados da aba Cópia de INICIAL

    let clientesEstoque = new Set(); // Criamos um conjunto para armazenar clientes válidos da aba Estoque

    // Adicionamos todos os clientes da aba Estoque ao conjunto (coluna A)
    for (let i = 1; i < dadosEstoque.length; i++) {
        let clienteEstoque = limparTexto(dadosEstoque[i][0]); // Coluna A da aba Estoque
        if (clienteEstoque) {
            clientesEstoque.add(clienteEstoque);
        };
    };

    let linhasParaExcluir = []; // Lista das linhas que precisam ser excluídas
    let mantidas = 0; // Contador de linhas mantidas
    let removidas = 0; // Contador de linhas removidas

    // Verifica cada cliente na aba Cópia de INICIAL
    for (let i = 1; i < dadosCopia.length; i++) {
        let clienteCopia = limparTexto(dadosCopia[i][10]); // Coluna K da aba Cópia de INICIAL (Cliente)
        if (!clientesEstoque.has(clienteCopia)) {
            linhasParaExcluir.push(i + 1); // Armazena as linhas para exclusão
            removidas++; // Incrementa o contador de removidos
        } else {
            mantidas++; // Incrementa o contador de mantidos
        };
    };

    for (let i = linhasParaExcluir.length - 1; i >= 0; i--) {
        abaCopiaInicial.deleteRow(linhasParaExcluir[i]);
    };

    Logger.log(`✅ Filtragem concluída! ${mantidas} linhas mantidas.`);
    Logger.log(`❌ ${removidas} linhas removidas (clientes não encontrados no Estoque).`);
};

function limparTexto(texto) {
    return texto ? texto.toString().trim().replace(/\s+/g, ' ').toLowerCase() : "";
};
