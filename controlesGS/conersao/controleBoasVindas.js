function ControleBoasVindas() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const abaControle = ss.getSheetByName("CONTROLE - BOAS VINDAS");

    var funcionarios = listaDeFuncionarios()
    var funcionariosPrescricao = [
        "JADE",
        "MARIA EDUARDA",
    ];
    let compl = "BOAS-VINDAS ";

    let logs = [];

    let funcionariosInfo = {};
    let motivosPorFase = {};

    const dataHoje = new Date().toDateString();
    let linhaPrimeirosQuadros = 5;
    let linhaPrescricao = 19;
    let linhaInicioAtvDia = 26;
    let colInicioAtvDia = 2;

    const cabecalhoProcon = [
        "FUNCIONÁRIO",
        "TOTAL DE CLIENTES",
        "DESCARTADO RELATO",
        "BOAS-VINDAS",
        "EM TRATATIVA",
        "VIÁVEIS",
    ];
    const cabecalhoPrescricao = [
        "FUNCIONÁRIO",
        "TOTAL DE CLIENTES",
        "BOAS-VINDAS",
        "EM TRATATIVA",
        "VIÁVEIS",
    ];
    const cabecalhoLeads = [
        "FUNCIONÁRIOS",
        "ROBÔ/CONSUMIDOR",
        "PROCON",
        "INDICAÇÃO",
        "RECONVERSÃO",
        "CONSIGNADO",
        "COMPANHIA DE ÔNIBUS",
        "BANCOS",
    ];
    const cabecalhoNaoViavel = [
        "FUNCIONARIOS",
        "NÃO RESPONDEU",
        "EMPRESA PEQUENA",
        "EMPRESA NÃO VIÁVEL",
        "RELATO NÃO VIÁVEL",
        "NÃO ATUAMOS",
    ];
    const cabecalhoFaseQueParou = [
        "FUNCIORARIOS",
        "NOVOS",
        "QUER PROSSEGUIR",
        "FALAR COM ATENDENTE",
        "PROVAS",
    ];
    const cabecalhoTotalFaseQueParou = [
        "MOTIVOS",
        "NOVOS",
        "QUER PROSSEGUIR",
        "FALAR COM ATENDENTE",
        "PROVAS",
    ];
    const cabecalhoStatus = [
        "FUNCIONÁRIO",
        "PAROU DE RESPONDER",
        "NÃO POSSUI INTERESSE",
        "NÃO VIÁVEL",
        "RESOLVEU",
        "COMARCA PROIBIDA"
    ];

    const clientesPerdidos = [
        "PAROU DE RESPONDER",
        "NÃO POSSUI INTERESSE",
        "NÃO VIÁVEL",
        "RESOLVEU",
        "COMARCA PROIBIDA",
        "SEM PROVAS"
    ]


    processarAba()



    abaControle.getRange(1, 1, 150, abaControle.getMaxColumns()).clearContent();
    escreverCabecalhos();
    escreverDados();
    escreverTotal();
    escreverQuadrosLeads();



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
    

    function processarAba() {

        funcionarios.forEach((funcionario) => {

            var tabBoasVindas = ss.getSheetByName(compl + funcionario);

            if (!funcionariosInfo[funcionario]) {
                funcionariosInfo[funcionario] = [];
            }

            if (tabBoasVindas) {

                var lastRow = tabBoasVindas.getLastRow();
                var valores = tabBoasVindas.getRange(2, 1, lastRow - 1, 13).getValues();

                var cabecalho = tabBoasVindas
                    .getRange(1, 1, 1, tabBoasVindas.getLastColumn())
                    .getValues()[0];

                var indiceData = obterIndiceCabecalho(cabecalho, "DATA");
                var indiceNome = obterIndiceCabecalho(cabecalho, "NOME");
                var indicefaseQueParou = obterIndiceCabecalho(cabecalho, "FASE QUE PAROU");
                var indiceLeads = obterIndiceCabecalho(cabecalho, "LEADS");
                var indiceNViavelRelato = obterIndiceCabecalho(cabecalho, "MOTIVO NÃO VIÁVEL RELATO");
                var indiceDuvidaViavel = obterIndiceCabecalho(cabecalho, "DÚVIDA SOBRE VIABILIDADE");
                var indiceStatus = obterIndiceCabecalho(cabecalho, "STATUS");
                var indiceViavel = obterIndiceCabecalho(cabecalho, "VIÁVEL?");
                var indiceTrigger = obterIndiceCabecalho(cabecalho, "TRIGGER");

                var countViavel = 0;
                var countDescartado = 0;
                var countBoasVindas = 0;
                var countNome = 0;
                var countDuvidaSobreViabilidade = 0;
                var countConvertido = 0;
                var countHoje = 0;
                var countPerda = 0;
                var countViaveisPerdido = 0;
                var countViaveisEmTratativa = 0;
                var countTratativa = 0;
                var aguardando = 0;
                var novos = 0;
                var fasesQuePararam = [];
                var motivoPerda = [];
                var viaveisPorCampanha = {};
                var naoViaveisPorCampanha = {};
                var leads = {
                    "FUNCIONARIOS": 0,
                    "PIPE/CONSUMIDOR": 0,
                    "ROBÔ/CONSUMIDOR": 0,
                    "PROCON": 0,
                    "INDICAÇÃO": 0,
                    "RECONVERSÃO": 0,
                    "CONSIGNADO": 0,
                    "COMPANHIA DE ÔNIBUS": 0,
                    "BANCOS": 0,
                    "VAZIO": 0,
                };
                var countMotivoNaoViavelRelato = {
                    "NÃO RESPONDEU": 0,
                    "EMPRESA PEQUENA": 0,
                    "EMPRESA NÃO VIÁVEL": 0,
                    "RELATO NÃO VIÁVEL": 0,
                    "NÃO ATUAMOS": 0,
                    "OUTROS": 0,
                };
                var countFaseQueParou = {
                    "NOVOS": 0,
                    "QUER PROSSEGUIR": 0,
                    "FALAR COM ATENDENTE ": 0,
                    "PROVAS": 0,
                };

                valores.forEach((linha, index) => {
                    var data = valores[index][indiceData];
                    var nome = valores[index][indiceNome];
                    var leadType = valores[index][indiceLeads];
                    var viavel = valores[index][indiceViavel];
                    var faseQueParou = valores[index][indicefaseQueParou];
                    var motivoNaoViavelRelato = valores[index][indiceNViavelRelato];
                    var duvidaSobreViabilidade = valores[index][indiceDuvidaViavel];
                    var status = valores[index][indiceStatus];
                    var dataTrigger = valores[index][indiceTrigger];

                    if (typeof nome === "string" && nome.trim() !== "") {

                        countNome++;


                        var booViavel = Boolean(viavel) == true;
                        var booDuvidaViavel = Boolean(duvidaSobreViabilidade) == true;


                        if (data && !motivoNaoViavelRelato && !booDuvidaViavel && !status && !booViavel) {
                            aguardando++;
                        };



                        // PRIMEIRO QUADRO


                        if (data && !motivoNaoViavelRelato) countBoasVindas++;

                        if (motivoNaoViavelRelato.trim() !== "") {
                            countDescartado++;
                        };

                        if (status == "EM TRATATIVA") {
                            countTratativa++;
                            if (booViavel) countViaveisEmTratativa++;
                        } else if (status == "DOCUMENTAÇÃO") {
                            countConvertido++;
                        };

                        if (booViavel) countViavel++;

                        if (clientesPerdidos.includes(status)) {
                            motivoPerda.push(status.trim());
                            countPerda++;
                        };
                        // FIM DO PRIMEIRO QUADRO




                        if (booDuvidaViavel) countDuvidaSobreViabilidade++;
                        if (faseQueParou) fasesQuePararam.push(faseQueParou.trim());


                        if (new Date(dataTrigger).toDateString() == dataHoje) {
                            countHoje++;
                        };

                        if (leadType.trim()) {
                            var upperCaseLeadType = leadType.toUpperCase();

                            if (leads.hasOwnProperty(upperCaseLeadType)) {
                                leads[upperCaseLeadType]++;
                            }

                            
                        } else {
                            leads["VAZIO"]++
                        };

                        if (motivoNaoViavelRelato) {
                            const upperCaseMotivoNV = motivoNaoViavelRelato.toUpperCase().trim();

                            // Incrementar o contador de motivos não viáveis
                            countMotivoNaoViavelRelato[upperCaseMotivoNV] = (countMotivoNaoViavelRelato[upperCaseMotivoNV] || 0) + 1;

                            // Garantir que `motivosPorFase` tem a estrutura necessária
                            if (!motivosPorFase[upperCaseMotivoNV]) {
                                motivosPorFase[upperCaseMotivoNV] = {};
                            };

                            // Incrementar o contador de fases para o motivo específico
                            motivosPorFase[upperCaseMotivoNV][faseQueParou] =
                                (motivosPorFase[upperCaseMotivoNV][faseQueParou] || 0) + 1;

                        };

                        if (faseQueParou) {
                            var upperCaseFase = faseQueParou.toUpperCase();

                            if (countFaseQueParou.hasOwnProperty(upperCaseFase)) {
                                countFaseQueParou[upperCaseFase]++;
                            };
                        };



                        // INFO POR CAMPANHA


                        if (Boolean(viavel) && clientesPerdidos.includes(status)) {
                            countViaveisPerdido++;
                            const upperCaseMomentoDP = status.toUpperCase().trim();

                            if (!viaveisPorCampanha[leadType]) {
                                viaveisPorCampanha[leadType] = {};
                            };

                            viaveisPorCampanha[leadType][upperCaseMomentoDP] = (viaveisPorCampanha[leadType][upperCaseMomentoDP] || 0) + 1;

                        } else if (!Boolean(viavel) && clientesPerdidos.includes(status)) {
                            const upperCaseMomentoDP = status.toUpperCase().trim();

                            if (!naoViaveisPorCampanha[leadType]) {
                                naoViaveisPorCampanha[leadType] = {};
                            };

                            naoViaveisPorCampanha[leadType][upperCaseMomentoDP] = (naoViaveisPorCampanha[leadType][upperCaseMomentoDP] || 0) + 1;
                        };

                        // FIM INFO POR CAMPANHA


                        if (booViavel && motivoNaoViavelRelato) {
                            logs.push(`Descartado relato e viavel: ${funcionario}: linha ${index + 2}`)
                        };

                    };
                });

                if (aguardando > 0) {
                    logs.push(`Sem preenchimento: ${funcionario}: quantidade ${aguardando}`)
                };

                funcionariosInfo[funcionario].push({
                    NOME: countNome,
                    MOTIVOS_NAO_VIAVEIS: countMotivoNaoViavelRelato,
                    MOMENTOS_DA_PERDA: countPerda,
                    CHAMADOS_HOJE: countHoje,
                    TRATATIVAS: countTratativa,
                    BOAS_VINDAS: countBoasVindas,
                    VIAVEL: countViavel,
                    VIAVEL_POR_CAMPANHA: viaveisPorCampanha,
                    NAO_VIAVEL_POR_CAMPANHA: naoViaveisPorCampanha,
                    CONVERTIDO: countConvertido,
                    DESCARTADO: countDescartado,
                    DUVIDA_SOBRE_VIABILIDADE: countDuvidaSobreViabilidade,
                    VIAVEIS_PERDIDOS: countViaveisPerdido,
                    VIAVEIS_EM_TRATATIVA: countViaveisEmTratativa,
                    FASES_QUE_PARARAM: countFaseQueParou,
                    LEADS: leads,
                });

            } else {
                logs.push(`Aba não encontrada para ${funcionario}`);
                return;
            }
        });
    };

    function escreverCabecalhos() {
        abaControle
            .getRange("B2:G3")
            .merge()
            .setValue("CONTROLE NÃO PRESCRIÇÃO - BOAS-VINDAS")
            .setHorizontalAlignment("center")
            .setBackground("#668CD9");
        abaControle
            .getRange(4, 2, 1, cabecalhoProcon.length)
            .setValues([cabecalhoProcon])
            .setBackground("#fff");

        abaControle
            .getRange("B17:F17")
            .merge()
            .setValue("CONTROLE PRESCRIÇÃO - BOAS-VINDAS")
            .setHorizontalAlignment("center")
            .setBackground("#668CD9");
        abaControle
            .getRange(18, 2, 1, cabecalhoPrescricao.length)
            .setValues([cabecalhoPrescricao])
            .setBackground("#FFF");

        abaControle
            .getRange("O2:V3")
            .merge()
            .setValue("TOTAL DE LEADS")
            .setHorizontalAlignment("center")
            .setBackground("#668CD9");
        abaControle
            .getRange(4, 15, 1, cabecalhoLeads.length)
            .setValues([cabecalhoLeads])
            .setBackground("#fff");

        abaControle
            .getRange("X2:AB3")
            .merge()
            .setValue("TOTAL NÃO VIÁVEIS POR FASE QUE PAROU")
            .setHorizontalAlignment("center")
            .setBackground("#668CD9");
        abaControle
            .getRange(4, 24, 1, cabecalhoFaseQueParou.length)
            .setValues([cabecalhoTotalFaseQueParou])
            .setBackground("#fff");

        abaControle
            .getRange("AD2:AH3")
            .merge()
            .setValue("FASE QUE PAROU")
            .setHorizontalAlignment("center")
            .setBackground("#668CD9");
        abaControle
            .getRange(4, 30, 1, cabecalhoFaseQueParou.length)
            .setValues([cabecalhoFaseQueParou])
            .setBackground("#fff");

        abaControle
            .getRange("AJ2:AO3")
            .merge()
            .setValue("MOTIVO NÃO VIAVEIS RELATO")
            .setHorizontalAlignment("center")
            .setBackground("#668CD9");
        abaControle
            .getRange(4, 36, 1, cabecalhoNaoViavel.length)
            .setValues([cabecalhoNaoViavel])
            .setBackground("#fff");
    };

    function escreverDados() {
        for (let motivo in motivosPorFase) {
            const fases = motivosPorFase[motivo];
            const totalNovos = fases["NOVOS"] || 0;
            const totalQuerProsseguir = fases["QUER PROSSEGUIR"] || 0;
            const totalProvas = fases["PROVAS"] || 0;
            const totalFalarComAtendente = fases["FALAR COM ATENDENTE "] || 0;
            abaControle
                .getRange(linhaPrimeirosQuadros + 1, 24, 1, cabecalhoTotalFaseQueParou.length)
                .setValues([
                    [
                        motivo,
                        totalNovos || 0,
                        totalQuerProsseguir || 0,
                        totalFalarComAtendente || 0,
                        totalProvas || 0,
                    ],
                ])
                .setBackground("#D9E5FF");

            linhaPrimeirosQuadros++;
        };

        linhaPrimeirosQuadros = 5;

        for (let funcionario in funcionariosInfo) {
            if (!funcionariosPrescricao.includes(funcionario)) {
                const dados = funcionariosInfo[funcionario][0];

                if (dados) {
                    abaControle
                        .getRange(linhaPrimeirosQuadros, 2, 1, cabecalhoProcon.length)
                        .setValues([
                            [
                                funcionario,
                                dados.NOME,
                                dados.DESCARTADO,
                                dados.BOAS_VINDAS,
                                dados.TRATATIVAS,
                                dados.VIAVEL,
                            ],
                        ])
                        .setBackground("#D9E5FF");

                    abaControle
                        .getRange(linhaPrimeirosQuadros, 15, 1, cabecalhoLeads.length)
                        .setValues([
                            [
                                funcionario,
                                dados.LEADS["ROBÔ/CONSUMIDOR"] || 0,
                                dados.LEADS["PROCON"] || 0,
                                dados.LEADS["INDICAÇÃO"] || 0,
                                dados.LEADS["RECONVERSÃO"] || 0,
                                dados.LEADS["CONSIGNADO"] || 0,
                                dados.LEADS["COMPANHIA DE ÔNIBUS"] || 0,
                                dados.LEADS["BANCOS"] || 0,
                            ],
                        ])
                        .setBackground("#D9E5FF");

                    abaControle
                        .getRange(linhaPrimeirosQuadros, 36, 1, cabecalhoNaoViavel.length)
                        .setValues([
                            [
                                funcionario,
                                dados.MOTIVOS_NAO_VIAVEIS["NÃO RESPONDEU"] || 0,
                                dados.MOTIVOS_NAO_VIAVEIS["EMPRESA PEQUENA"] || 0,
                                dados.MOTIVOS_NAO_VIAVEIS["EMPRESA NÃO VIÁVEL"] || 0,
                                dados.MOTIVOS_NAO_VIAVEIS["RELATO NÃO VIÁVEL"] || 0,
                                dados.MOTIVOS_NAO_VIAVEIS["NÃO ATUAMOS"] || 0,
                            ],
                        ])
                        .setBackground("#D9E5FF");

                    abaControle
                        .getRange(linhaPrimeirosQuadros, 30, 1, cabecalhoFaseQueParou.length)
                        .setValues([
                            [
                                funcionario,
                                dados.FASES_QUE_PARARAM["NOVOS"] || 0,
                                dados.FASES_QUE_PARARAM["QUER PROSSEGUIR"] || 0,
                                dados.FASES_QUE_PARARAM["FALAR COM ATENDENTE "] || 0,
                                dados.FASES_QUE_PARARAM["PROVAS"] || 0,
                            ],
                        ])
                        .setBackground("#D9E5FF");
                    linhaPrimeirosQuadros++;
                };
            } else {
                const dados = funcionariosInfo[funcionario][0];

                abaControle
                    .getRange(linhaPrescricao, 2, 1, cabecalhoPrescricao.length)
                    .setValues([
                        [
                            funcionario,
                            dados.NOME,
                            dados.BOAS_VINDAS,
                            dados.TRATATIVAS,
                            dados.VIAVEL,
                        ],
                    ])
                    .setBackground("#D9E5FF");
                linhaPrescricao++;
            };
        };
    };

    function escreverTotal() {
        let colunas = [
            3, 4, 5, 6, 7, 16, 17, 18, 19, 20, 21, 22, 31, 32, 33, 34,
            37, 38, 39, 40, 41
        ];
        let colunasFase = [25, 26, 27, 28];
        let colunasPrescricao = [3, 4, 5, 6];
        let colunasVProcon = [3, 4, 5, 6, 7]
        const quantidadeChaves = Object.keys(motivosPorFase).length;

        let ultimaLinha = 5 + funcionarios.length - 1 - funcionariosPrescricao.length;
        let ultimaLinhaPrescricao = 19 + funcionariosPrescricao.length - 1;
        let ultimaLinhaPorFase = 6 + quantidadeChaves - 1;
        let ultimaLinhaVProcon = 18 + funcionarios.length - 1 - funcionariosPrescricao.length;

        abaControle
            .getRange(ultimaLinha + 1, 2)
            .setValue("TOTAL")
            .setBackground("#0cf25d");
        abaControle
            .getRange(ultimaLinhaPrescricao + 1, 2)
            .setValue("TOTAL")
            .setBackground("#0cf25d");
        abaControle
            .getRange(ultimaLinha + 1, 15)
            .setValue("TOTAL")
            .setBackground("#0cf25d");
        abaControle
            .getRange(ultimaLinha + 1, 30)
            .setValue("TOTAL")
            .setBackground("#0cf25d");
        abaControle
            .getRange(ultimaLinha + 1, 36)
            .setValue("TOTAL")
            .setBackground("#0cf25d");
        abaControle
            .getRange(ultimaLinhaPorFase + 1, 24)
            .setValue("TOTAL")
            .setBackground("#0cf25d");

        colunasFase.forEach((coluna) => {
            let nomeColuna = obterNomeColuna(coluna);

            abaControle
                .getRange(ultimaLinhaPorFase + 1, coluna)
                .setFormula(`=SUM(${nomeColuna}5:${nomeColuna}${ultimaLinhaPorFase})`)
                .setBackground("#0cf25d");
        });

        colunas.forEach((coluna) => {
            let nomeColuna = obterNomeColuna(coluna);

            abaControle
                .getRange(ultimaLinha + 1, coluna)
                .setFormula(`=SUM(${nomeColuna}5:${nomeColuna}${ultimaLinha})`)
                .setBackground("#0cf25d");
        });

        colunasPrescricao.forEach((coluna) => {
            let nomeColuna = obterNomeColuna(coluna);

            abaControle
                .getRange(ultimaLinhaPrescricao + 1, coluna)
                .setFormula(`=SUM(${nomeColuna}19:${nomeColuna}${ultimaLinhaPrescricao})`)
                .setBackground("#0cf25d");
        });

    };

    function escreverQuadrosLeads() {
        const campanhas = [
            "PROCON",
            "CONSIGNADO",
            "ROBÔ/CONSUMIDOR",
            "COMPANHIA DE ÔNIBUS",
            "RECONVERSÃO",
            "INDICAÇÃO",
            "BANCOS"
        ];

        const colunasTotais = [3, 4, 5, 6, 7]; // Colunas para aplicar as fórmulas de total
        let linhaInicialQLead = 29; // Linha inicial para o primeiro quadro

        campanhas.forEach((campanha) => {
            // Escrever título do quadro (VIÁVEIS)
            abaControle
                .getRange(linhaInicialQLead, 2, 1, 6)
                .merge()
                .setValue(`VIÁVEIS POR CAMPANHA - ${campanha}`)
                .setHorizontalAlignment("center")
                .setBackground("#668CD9");

            // Escrever cabeçalho do quadro (VIÁVEIS)
            abaControle
                .getRange(linhaInicialQLead + 1, 2, 1, cabecalhoStatus.length)
                .setValues([cabecalhoStatus])
                .setBackground("#fff");

            // Inicializar linha para dados dos funcionários (VIÁVEIS)
            let linhaDados = linhaInicialQLead + 2;

            for (let funcionario in funcionariosInfo) {
                if (!funcionariosPrescricao.includes(funcionario)) {
                    const dados = funcionariosInfo[funcionario][0];

                    if (dados) {
                        const campanhaData = dados.VIAVEL_POR_CAMPANHA[campanha] || {};

                        abaControle
                            .getRange(linhaDados, 2, 1, 6)
                            .setValues([
                                [
                                    funcionario,
                                    campanhaData["PAROU DE RESPONDER"] || 0,
                                    campanhaData["NÃO POSSUI INTERESSE"] || 0,
                                    campanhaData["NÃO VIÁVEL"] || 0,
                                    campanhaData["RESOLVEU"] || 0,
                                    campanhaData["COMARCA PROIBIDA"] || 0,
                                ]
                            ])
                            .setBackground("#D9E5FF");

                        linhaDados++;
                    }
                }
            }

            // Calcular totais para a campanha atual (VIÁVEIS)
            abaControle
                .getRange(linhaDados, 2)
                .setValue("TOTAL")
                .setBackground("#0cf25d");

            colunasTotais.forEach((coluna) => {
                const nomeColuna = obterNomeColuna(coluna);

                abaControle
                    .getRange(linhaDados, coluna)
                    .setFormula(`=SUM(${nomeColuna}${linhaInicialQLead + 2}:${nomeColuna}${linhaDados - 1})`)
                    .setBackground("#0cf25d");
            });

            // Escrever título do quadro (NÃO VIÁVEIS)
            const colunaInicialNaoViaveis = 9; // Coluna inicial para NÃO VIÁVEIS (ao lado)
            abaControle
                .getRange(linhaInicialQLead, colunaInicialNaoViaveis, 1, 6)
                .merge()
                .setValue(`NÃO VIÁVEIS POR CAMPANHA - ${campanha}`)
                .setHorizontalAlignment("center")
                .setBackground("#668CD9");

            // Escrever cabeçalho do quadro (NÃO VIÁVEIS)
            abaControle
                .getRange(linhaInicialQLead + 1, colunaInicialNaoViaveis, 1, cabecalhoStatus.length)
                .setValues([cabecalhoStatus])
                .setBackground("#fff");

            // Inicializar linha para dados (NÃO VIÁVEIS)
            let linhaDadosNaoViaveis = linhaInicialQLead + 2;

            for (let funcionario in funcionariosInfo) {
                if (!funcionariosPrescricao.includes(funcionario)) {
                    const dados = funcionariosInfo[funcionario][0];

                    if (dados) {
                        const campanhaData = dados.NAO_VIAVEL_POR_CAMPANHA[campanha] || {};

                        abaControle
                            .getRange(linhaDadosNaoViaveis, colunaInicialNaoViaveis, 1, 6)
                            .setValues([
                                [
                                    funcionario,
                                    campanhaData["PAROU DE RESPONDER"] || 0,
                                    campanhaData["NÃO POSSUI INTERESSE"] || 0,
                                    campanhaData["NÃO VIÁVEL"] || 0,
                                    campanhaData["RESOLVEU"] || 0,
                                    campanhaData["COMARCA PROIBIDA"] || 0,
                                ]
                            ])
                            .setBackground("#D9E5FF");

                        linhaDadosNaoViaveis++;
                    }
                }
            }

            // Calcular totais para a campanha atual (NÃO VIÁVEIS)
            abaControle
                .getRange(linhaDadosNaoViaveis, colunaInicialNaoViaveis)
                .setValue("TOTAL")
                .setBackground("#0cf25d");

            colunasTotais.forEach((coluna, index) => {
                const nomeColuna = obterNomeColuna(colunaInicialNaoViaveis + index + 1);

                abaControle
                    .getRange(linhaDadosNaoViaveis, colunaInicialNaoViaveis + index + 1)
                    .setFormula(`=SUM(${nomeColuna}${linhaInicialQLead + 2}:${nomeColuna}${linhaDadosNaoViaveis - 1})`)
                    .setBackground("#0cf25d");
            });

            linhaInicialQLead = Math.max(linhaDados, linhaDadosNaoViaveis) + 3;
        });
    };


};

function contarValores(array) {
    return array.reduce((acc, curr) => {
        acc[curr] = (acc[curr] || 0) + 1;
        return acc;
    }, {});
};

function obterNomeColuna(indice) {
    let coluna = "";
    while (indice > 0) {
        let resto = (indice - 1) % 26;
        coluna = String.fromCharCode(65 + resto) + coluna;
        indice = Math.floor((indice - 1) / 26);
    }
    return coluna;
};

function obterIndiceCabecalho(allCabecalho, nomeColuna) {
    return allCabecalho.indexOf(nomeColuna);

};
