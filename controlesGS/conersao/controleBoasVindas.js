function ControleBoasVindas() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const abaControle = ss.getSheetByName("CONTROLE - BOAS VINDAS");

    let funcionarios = listaDeFuncionarios();
    let funcionariosPrescricao = ["JADE", "MARIA EDUARDA"];
    let compl = "BOAS-VINDAS ";

    let funcionariosInfo = {};
    let motivosPorFase = {};
    let logs = [];

    let clientesPerdidos = [
        "PAROU DE RESPONDER",
        "NÃO POSSUI INTERESSE",
        "NÃO VIÁVEL",
        "RESOLVEU",
        "COMARCA PROIBIDA",
        "SEM PROVAS",
    ];

    processarAba();
    escreverLogsBoasVindas();

    // console.log(funcionariosInfo);
    // console.log(motivosPorFase);
    console.log(logs);


    function processarAba() {
        funcionarios.forEach((funcionario) => {
            let tabBoasVindas = ss
                .getSheetByName(compl + funcionario)
                .getDataRange()
                .getValues();

            if (!tabBoasVindas) {
                logs.push(`Aba não encontrada para ${funcionario}`);
                return;
            };

            if (!funcionariosInfo[funcionario]) funcionariosInfo[funcionario] = {};

            let cabecalho = tabBoasVindas[0];
            let indiceData = cabecalho.indexOf("DATA");
            let indiceNome = cabecalho.indexOf("NOME");
            let indicefaseQueParou = cabecalho.indexOf("FASE QUE PAROU");
            let indiceLeads = cabecalho.indexOf("LEADS");
            let indiceNViavelRelato = cabecalho.indexOf("MOTIVO NÃO VIÁVEL RELATO");
            let indiceDuvidaViavel = cabecalho.indexOf("DÚVIDA SOBRE VIABILIDADE");
            let indiceStatus = cabecalho.indexOf("STATUS");
            let indiceViavel = cabecalho.indexOf("VIÁVEL?");
            let indiceTrigger = cabecalho.indexOf("TRIGGER");

            let lsp = 0

            for (let index = 1; index < tabBoasVindas.length; index++) {
                let data = tabBoasVindas[index][indiceData];
                let nome = tabBoasVindas[index][indiceNome];
                let leadType = tabBoasVindas[index][indiceLeads];
                let viavel = tabBoasVindas[index][indiceViavel];
                let faseQueParou = tabBoasVindas[index][indicefaseQueParou];
                let motivoNaoViavelRelato = tabBoasVindas[index][indiceNViavelRelato];
                let duvidaSobreViabilidade = tabBoasVindas[index][indiceDuvidaViavel];
                let status = tabBoasVindas[index][indiceStatus];
                let dataTrigger = tabBoasVindas[index][indiceTrigger];

                if (nome) {

                    if (!data) lsp++;

                    // QUADRO PRINCIPAL + QUADROS POR CAMPANHA

                    if (leadType) {
                        if (!funcionariosInfo[funcionario][leadType])
                            funcionariosInfo[funcionario][leadType] = {};
                        if (!funcionariosInfo[funcionario]["GERAL"])
                            funcionariosInfo[funcionario]["GERAL"] = {};

                        funcionariosInfo[funcionario]["GERAL"]["CLIENTES"] =
                            (funcionariosInfo[funcionario]["GERAL"]["CLIENTES"] || 0) + 1;

                        funcionariosInfo[funcionario][leadType]["CLIENTES"] =
                            (funcionariosInfo[funcionario][leadType]["CLIENTES"] || 0) + 1;

                        if (motivoNaoViavelRelato !== "") {
                            funcionariosInfo[funcionario]["GERAL"]["DESCARTADO"] =
                                (funcionariosInfo[funcionario]["GERAL"]["DESCARTADO"] || 0) + 1;

                            funcionariosInfo[funcionario][leadType]["DESCARTADO"] =
                                (funcionariosInfo[funcionario][leadType]["DESCARTADO"] || 0) + 1;
                        };

                        if (data && !motivoNaoViavelRelato) {
                            funcionariosInfo[funcionario]["GERAL"]["BOAS-VINDAS"] =
                                (funcionariosInfo[funcionario]["GERAL"]["BOAS-VINDAS"] || 0) + 1;

                            funcionariosInfo[funcionario][leadType]["BOAS-VINDAS"] =
                                (funcionariosInfo[funcionario][leadType]["BOAS-VINDAS"] || 0) + 1;
                        };

                        if (viavel == "VERDADEIRO") {
                            funcionariosInfo[funcionario]["GERAL"]["VIAVEL"] =
                                (funcionariosInfo[funcionario]["GERAL"]["VIAVEL"] || 0) + 1;

                            funcionariosInfo[funcionario][leadType]["VIAVEL"] =
                                (funcionariosInfo[funcionario][leadType]["VIAVEL"] || 0) + 1;
                        };

                        if (clientesPerdidos.includes(status)) {
                            funcionariosInfo[funcionario]["GERAL"]["CLIENTES PERDIDOS"] =
                                (funcionariosInfo[funcionario]["GERAL"]["CLIENTES PERDIDOS"] ||
                                    0) + 1;

                            funcionariosInfo[funcionario][leadType]["CLIENTES PERDIDOS"] =
                                (funcionariosInfo[funcionario][leadType]["CLIENTES PERDIDOS"] ||
                                    0) + 1;

                            funcionariosInfo[funcionario]["GERAL"][status] =
                                (funcionariosInfo[funcionario]["GERAL"][status] || 0) + 1;

                            funcionariosInfo[funcionario][leadType][status] =
                                (funcionariosInfo[funcionario][leadType][status] || 0) + 1;
                        };

                        if (status == "EM TRATATIVA") {
                            funcionariosInfo[funcionario]["GERAL"]["TRATATIVAS"] =
                                (funcionariosInfo[funcionario]["GERAL"]["TRATATIVAS"] || 0) + 1;

                            funcionariosInfo[funcionario][leadType]["TRATATIVAS"] =
                                (funcionariosInfo[funcionario][leadType]["TRATATIVAS"] || 0) + 1;

                            if (viavel == "VERDADEIRO") {
                                funcionariosInfo[funcionario]["GERAL"]["VIAVEIS EM TRATATIVA"] =
                                    (funcionariosInfo[funcionario]["GERAL"][
                                        "VIAVEIS EM TRATATIVA"
                                    ] || 0) + 1;

                                funcionariosInfo[funcionario][leadType]["VIAVEIS EM TRATATIVA"] =
                                    (funcionariosInfo[funcionario][leadType][
                                        "VIAVEIS EM TRATATIVA"
                                    ] || 0) + 1;
                            };
                        } else if (status == "DOCUMENTAÇÃO") {
                            funcionariosInfo[funcionario]["GERAL"]["CONVERTIDO"] =
                                (funcionariosInfo[funcionario]["GERAL"]["CONVERTIDO"] || 0) + 1;

                            funcionariosInfo[funcionario][leadType]["CONVERTIDO"] =
                                (funcionariosInfo[funcionario][leadType]["CONVERTIDO"] || 0) + 1;
                        };

                        if (duvidaSobreViabilidade == 'VERDADEIRO') {
                            funcionariosInfo[funcionario]["GERAL"]["DUVIDA SOBRE VIABILIDADE"] =
                                (funcionariosInfo[funcionario]["GERAL"][
                                    "DUVIDA SOBRE VIABILIDADE"
                                ] || 0) + 1;

                            funcionariosInfo[funcionario][leadType][
                                "DUVIDA SOBRE VIABILIDADE"
                            ] =
                                (funcionariosInfo[funcionario][leadType][
                                    "DUVIDA SOBRE VIABILIDADE"
                                ] || 0) + 1;
                        };
                    };

                    // FIM DO QUADRO PRINCIPAL + QUADROS POR CAMPANHA

                    if (motivoNaoViavelRelato) {
                        funcionariosInfo[funcionario]["GERAL"]["MOTIVO NÃO VIÁVEL RELATO"] =
                            (funcionariosInfo[funcionario]["GERAL"][
                                "MOTIVO NÃO VIÁVEL RELATO"
                            ] || 0) + 1;

                        if (faseQueParou) {
                            if (!motivosPorFase[faseQueParou])
                                motivosPorFase[faseQueParou] = {};

                            motivosPorFase[faseQueParou][motivoNaoViavelRelato] =
                                (motivosPorFase[faseQueParou][motivoNaoViavelRelato] || 0) + 1;
                        };
                    };

                    if (faseQueParou) {
                        funcionariosInfo[funcionario]["GERAL"]["FASE QUE PAROU"] =
                            (funcionariosInfo[funcionario]["GERAL"]["FASE QUE PAROU"] || 0) + 1;
                    };

                    if (viavel == 'VERDADEIRO' && motivoNaoViavelRelato) {
                        logs.push(
                            `Descartado relato e viavel: ${funcionario}: linha ${index + 2} `
                        );
                    };
                } else {
                    logs.push(`aba de ${funcionario} não possui data em ${lsp} linhas`);
                };
            };

        });
    };

    function escreverHora() {
        const horaAtual = Utilities.formatDate(
            new Date(),
            Session.getScriptTimeZone(),
            "HH:mm"
        );

        abaControle
            .getRange("A1")
            .setValue("ULTIMA EXECUÇÃO")
            .setBackground("#668CD9");
        abaControle.getRange("A2").setValue(horaAtual).setBackground("#FFF");
    };


    function escreverLogsBoasVindas() {
        let abaLogs = ss.getSheetByName("LOGS");
        let linha = 3;
        let ultimaLinha = abaLogs.getLastRow();
        abaLogs.getRange(linha, 6, ultimaLinha, 3).clearContent();

        // for (i = linha; i <= 18; i++) {
        //     abaLogs.getRange(i, 2, 1, 3).merge();
        // };

        logs.forEach(log => {
            abaLogs.getRange(linha, 6, 1, 3).merge();
            abaLogs.getRange(linha, 6).setValue(log);
            linha++;
        });

    };
};
