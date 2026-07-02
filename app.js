// Domino's DRE Analyzer - Lógica do Aplicativo
// Processamento 100% Client-Side com SheetJS (XLSX.js)

// Parâmetros de Referência Ideais pré-incorporados (como fallbacks se a ref não for subida)
const REFERENCIA_CLUSTERS = {
    "ATE_100K": {
        nome: "Até R$ 100K",
        meta_cmv: -39.2,
        meta_pessoal: -27.7,
        meta_ocupacao: -9.8,
        meta_utilidades: -6.8,
        meta_ebitda: -6.8
    },
    "ENTRE_100K_150K": {
        nome: "Entre R$ 100K e R$ 150K",
        meta_cmv: -35.5,
        meta_pessoal: -25.7,
        meta_ocupacao: -8.1,
        meta_utilidades: -5.5,
        meta_ebitda: 1.0
    },
    "ENTRE_150K_200K": {
        nome: "Entre R$ 150K e R$ 200K",
        meta_cmv: -33.5,
        meta_pessoal: -22.2,
        meta_ocupacao: -5.1,
        meta_utilidades: -6.7,
        meta_ebitda: 10.3
    },
    "ENTRE_200K_250K": {
        nome: "Entre R$ 200K e R$ 250K",
        meta_cmv: -33.5,
        meta_pessoal: -22.0,
        meta_ocupacao: -5.1,
        meta_utilidades: -5.8,
        meta_ebitda: 11.5
    },
    "ACIMA_250K": {
        nome: "Acima de R$ 250K",
        meta_cmv: -33.5,
        meta_pessoal: -20.4,
        meta_ocupacao: -5.0,
        meta_utilidades: -4.8,
        meta_ebitda: 15.8
    }
};

// Dados locais processados das lojas
let lojasProcessadas = {};
let currentLoja = "";
let currentPeriod = "";

// Elementos DOM
const dropZone = document.getElementById('dre-dropzone');
const fileInput = document.getElementById('dre-file-input');
const fileInfoBar = document.getElementById('file-info-bar');
const uploadedFileName = document.getElementById('uploaded-file-name');
const removeFileBtn = document.getElementById('remove-file-btn');
const analyzeBtn = document.getElementById('analyze-btn');
const resultsSection = document.getElementById('results-section');
const storeSelect = document.getElementById('store-select');
const selectorsContainer = document.getElementById('selectors-container');

// Elementos do Relatório
const kpiFaturamento = document.getElementById('kpi-faturamento');
const kpiReceitaLiquida = document.getElementById('kpi-receita-liquida');
const kpiEbitda = document.getElementById('kpi-ebitda');
const kpiEbitdaStatus = document.getElementById('kpi-ebitda-status');
const kpiEbitdaCard = document.getElementById('kpi-ebitda-card');
const diagnosticText = document.getElementById('diagnostic-text');
const tableBody = document.getElementById('comparison-table-body');
const actionPlanContainer = document.getElementById('action-plan-container');

// Eventos de Arrastar e Soltar
dropZone.addEventListener('dragover', (e) => {
    e.preventDefault();
    dropZone.classList.add('dragover');
});

dropZone.addEventListener('dragleave', () => {
    dropZone.classList.remove('dragover');
});

dropZone.addEventListener('drop', (e) => {
    e.preventDefault();
    dropZone.classList.remove('dragover');
    if (e.dataTransfer.files.length > 0) {
        handleFileSelect(e.dataTransfer.files[0]);
    }
});

fileInput.addEventListener('change', (e) => {
    if (e.target.files.length > 0) {
        handleFileSelect(e.target.files[0]);
    }
});

removeFileBtn.addEventListener('remove', clearFile);
removeFileBtn.addEventListener('click', clearFile);

function handleFileSelect(file) {
    if (!file.name.endsWith('.xlsx') && !file.name.endsWith('.xls')) {
        alert('Por favor, selecione apenas arquivos do Excel (.xlsx ou .xls).');
        return;
    }
    
    uploadedFileName.textContent = file.name;
    dropZone.style.display = 'none';
    fileInfoBar.style.display = 'flex';
    analyzeBtn.disabled = false;
    
    // Guardar o arquivo no botão
    analyzeBtn.fileData = file;
    analyzeBtn.uploadedFileName = file.name; // Salva o nome do arquivo para pré-seleção
}

function clearFile() {
    fileInput.value = '';
    dropZone.style.display = 'block';
    fileInfoBar.style.display = 'none';
    analyzeBtn.disabled = true;
    analyzeBtn.fileData = null;
    analyzeBtn.uploadedFileName = null;
    resultsSection.style.display = 'none';
    lojasProcessadas = {};
}

// Ação de Análise
analyzeBtn.addEventListener('click', () => {
    const file = analyzeBtn.fileData;
    if (!file) return;
    
    try {
        const reader = new FileReader();
        reader.onload = function(e) {
            try {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, {type: 'array'});
                
                processDREWorkbook(workbook);
            } catch (err) {
                console.error(err);
                alert("Erro ao ler dados da planilha: " + err.message + "\n\nPor favor, envie este erro ao suporte.");
            }
        };
        reader.onerror = function(err) {
            alert("Erro de leitura do arquivo: " + err.message);
        };
        reader.readAsArrayBuffer(file);
    } catch (err) {
        alert("Erro ao iniciar leitor de arquivos: " + err.message);
    }
});

// Processamento da Planilha DRE
function processDREWorkbook(workbook) {
    try {
        lojasProcessadas = {};
        const sheetNames = workbook.SheetNames;
        
        let allPeriods = new Set();
        
        sheetNames.forEach(sheetName => {
            // Ignorar abas gerais que não são lojas individuais
            if (sheetName === "Resumo" || sheetName === "Base" || sheetName.startsWith("DRE ")) {
                return;
            }
            
            const worksheet = workbook.Sheets[sheetName];
            const rows = XLSX.utils.sheet_to_json(worksheet, {header: 1});
            
            let lojaData = parseStoreDRE(rows);
            
            // Verificar se a loja tem faturamento em pelo menos um período
            let temFaturamento = false;
            lojaData.periods.forEach(p => {
                if (lojaData.values[p].receitaBruta > 0) {
                    temFaturamento = true;
                    allPeriods.add(p);
                }
            });
            
            if (temFaturamento) {
                lojasProcessadas[sheetName] = lojaData;
            }
        });
        
        const listLojas = Object.keys(lojasProcessadas);
        
        if (listLojas.length === 0) {
            alert("Não foi possível identificar nenhuma aba de loja válida com faturamento positivo na planilha.");
            return;
        }
        
        // Tentar pré-selecionar a loja baseando-se no nome do arquivo
        const fileNameLower = (analyzeBtn.uploadedFileName || "").toLowerCase();
        let matchedStore = listLojas[0];
        
        for (let i = 0; i < listLojas.length; i++) {
            const storeName = listLojas[i].toLowerCase();
            // Remove acentos e caracteres especiais para comparação
            const cleanStoreName = storeName.normalize("NFD").replace(/[\u0300-\u036f]/g, "");
            const cleanFileName = fileNameLower.normalize("NFD").replace(/[\u0300-\u036f]/g, "");
            
            if (cleanFileName.includes(cleanStoreName)) {
                matchedStore = listLojas[i];
                break;
            }
        }
        
        // Atualizar seletor de lojas
        storeSelect.innerHTML = "";
        listLojas.forEach(loja => {
            const option = document.createElement('option');
            option.value = loja;
            option.textContent = loja;
            if (loja === matchedStore) {
                option.selected = true;
            }
            storeSelect.appendChild(option);
        });
        
        // Atualizar seletor de períodos
        const periodSelect = document.getElementById('period-select');
        periodSelect.innerHTML = "";
        
        const sortedPeriods = Array.from(allPeriods);
        sortedPeriods.forEach(p => {
            const option = document.createElement('option');
            option.value = p;
            option.textContent = p;
            periodSelect.appendChild(option);
        });
        
        // Mostrar os seletores (loja e período)
        selectorsContainer.style.display = 'flex';
        
        // Definir padrões de inicialização
        currentLoja = matchedStore;
        
        // Selecionar "Maio" por padrão se disponível, senão o último mês
        let defaultPeriod = sortedPeriods[sortedPeriods.length - 1];
        if (sortedPeriods.includes("Maio")) {
            defaultPeriod = "Maio";
        }
        
        periodSelect.value = defaultPeriod;
        currentPeriod = defaultPeriod;
        
        // Mostrar seção de resultados e carregar primeira loja
        resultsSection.style.display = 'block';
        renderAnalysis(currentLoja, currentPeriod);
        
        // Rolar até os resultados
        resultsSection.scrollIntoView({ behavior: 'smooth' });
    } catch (err) {
        console.error(err);
        alert("Erro no processamento da DRE: " + err.message);
    }
}

// Analisar linhas da DRE da Loja e extrair todos os períodos
function parseStoreDRE(rows) {
    let result = {
        periods: [],
        values: {}
    };
    
    // Encontrar a linha de cabeçalho e a coluna dos meses
    let monthRowIndex = -1;
    
    for (let r = 0; r < Math.min(rows.length, 15); r++) {
        const row = rows[r];
        if (!row) continue;
        for (let c = 0; c < row.length; c++) {
            const cellVal = String(row[c]).trim();
            if (["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"].includes(cellVal)) {
                monthRowIndex = r;
                break;
            }
        }
        if (monthRowIndex !== -1) break;
    }
    
    if (monthRowIndex === -1) {
        // Sem cabeçalho de meses, assume coluna 1 como valores fixos
        result.periods.push("Geral");
        result.values["Geral"] = parseDREColumn(rows, 1);
        return result;
    }
    
    const headerRow = rows[monthRowIndex];
    // Adicionar cada mês encontrado como um período independente
    for (let c = 0; c < headerRow.length; c++) {
        const cellVal = String(headerRow[c]).trim();
        if (["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"].includes(cellVal)) {
            result.periods.push(cellVal);
            result.values[cellVal] = parseDREColumn(rows, c);
        }
    }
    
    return result;
}

// Processa uma coluna específica de dados para retornar as contas gerenciais daquele mês
function parseDREColumn(rows, colIndex) {
    let data = {
        receitaBruta: 0,
        receitaLiquida: 0,
        devolucoes: 0,
        cmvTotal: 0,
        cmvBebidas: 0,
        cmvMassas: 0,
        cmvLaticinios: 0,
        cmvAlimentos: 0,
        pessoalTotal: 0,
        salarios: 0,
        horasExtras: 0,
        encargos: 0,
        beneficios: 0,
        rescisao: 0,
        ocupacaoTotal: 0,
        aluguel: 0,
        energia: 0,
        gas: 0,
        agua: 0,
        despComerciais: 0,
        lucroOperacional: 0
    };
    
    let sumPessoal = 0;
    
    rows.forEach(row => {
        if (!row || row.length <= colIndex) return;
        const conta = String(row[0]).trim();
        const valorVal = parseCurrency(row[colIndex]);
        const absVal = Math.abs(valorVal);
        
        switch (true) {
            case conta.toUpperCase().includes("RECEITA") && conta.toUpperCase().includes("BRUTA"):
                data.receitaBruta = valorVal;
                break;
            case conta.toUpperCase().includes("DEVOLU") || conta.toUpperCase().includes("CANCEL"):
                data.devolucoes = valorVal;
                break;
            case conta.toUpperCase().includes("RECEITA") && (conta.toUpperCase().includes("LIQ") || conta.toUpperCase().includes("LÍQ") || conta.toUpperCase().includes("LQ") || conta.toUpperCase().includes("LÝQ") || conta.toUpperCase().includes("LQ")):
                data.receitaLiquida = valorVal;
                break;
            case conta.toUpperCase().includes("CMV") || conta.toUpperCase().includes("CUSTO DE MERCADORIA VENDIDA") || conta.toUpperCase().includes("CUSTO DA MERCADORIA VENDIDA"):
                data.cmvTotal = valorVal;
                break;
            case conta.toUpperCase().includes("BEBIDAS"):
                data.cmvBebidas = valorVal;
                break;
            case conta.toUpperCase().includes("MASSAS"):
                data.cmvMassas = valorVal;
                break;
            case conta.toUpperCase().includes("LATIC") || conta.toUpperCase().includes("LATÍC") || conta.toUpperCase().includes("MUSSAR"):
                data.cmvLaticinios = valorVal;
                break;
            case conta.toUpperCase().includes("ALIMENT") || conta.toUpperCase().includes("INSUMO"):
                data.cmvAlimentos = valorVal;
                break;
            case conta.toUpperCase().includes("SALÁR") || conta.toUpperCase().includes("SALAR") || conta.toUpperCase().includes("ORDENAD"):
                data.salarios = valorVal;
                sumPessoal += absVal;
                break;
            case conta.toUpperCase().includes("HORAS EXTRAS"):
                data.horasExtras = valorVal;
                sumPessoal += absVal;
                break;
            case conta.toUpperCase().includes("ENCARGO"):
                data.encargos = valorVal;
                sumPessoal += absVal;
                break;
            case conta.toUpperCase().includes("BENEF") || conta.toUpperCase().includes("VALE TRANSP") || conta.toUpperCase().includes("VALE REFEI") || conta.toUpperCase().includes("LANCHES") || conta.toUpperCase().includes("ASSISTENCIA M") || conta.toUpperCase().includes("ASSISTÊNCIA M"):
                data.beneficios += valorVal;
                sumPessoal += absVal;
                break;
            case conta.toUpperCase().includes("RECIS") || conta.toUpperCase().includes("RESCIS") || conta.toUpperCase().includes("FGTS") || conta.toUpperCase().includes("FÉRIAS") || conta.toUpperCase().includes("FERIAS"):
                data.rescisao += valorVal;
                sumPessoal += absVal;
                break;
            case conta.toUpperCase().includes("PRÓ-LABORE") || conta.toUpperCase().includes("PRO-LABORE") || conta.toUpperCase().includes("OUTROS - PESSOAL"):
                sumPessoal += absVal;
                break;
            case conta.toUpperCase().includes("ALUGUEL") || conta.toUpperCase().includes("ALUGUÉL"):
                data.aluguel = valorVal;
                data.ocupacaoTotal = valorVal;
                break;
            case conta.toUpperCase().includes("ENERGIA"):
                data.energia = valorVal;
                break;
            case conta.toUpperCase().includes("GÁS") || conta.toUpperCase().includes("GAS"):
                data.gas = valorVal;
                break;
            case conta.toUpperCase().includes("ÁGUA") || conta.toUpperCase().includes("AGUA"):
                data.agua = valorVal;
                break;
            case conta.toUpperCase().includes("CUSTO UTILIDADES") || conta.toUpperCase().includes("= CUSTO UTILIDADES"):
                data.ocupacaoTotal = valorVal; // Usamos utilidades + aluguel como total de ocupação no app
                break;
            case conta.toUpperCase().includes("DESPESAS COMERCIAIS") || conta.toUpperCase().includes("MARKETING") || conta.toUpperCase().includes("PROPAGANDA"):
                data.despComerciais += valorVal;
                break;
            case conta.toUpperCase().includes("RESULTADO OPERACIONAL") || conta.toUpperCase().includes("EBITDA") || conta.toUpperCase().includes("LUCRO OPERACIONAL"):
                data.lucroOperacional = valorVal;
                break;
        }
    });
    
    // Atualizar pessoalTotal com a soma das subcontas se não tiver sido extraído diretamente
    data.pessoalTotal = -sumPessoal;
    
    // Se o faturamento líquido ou bruto estiver zerado, tenta forçar leitura de outras abas
    if (data.receitaLiquida === 0 && data.receitaBruta > 0) {
        data.receitaLiquida = data.receitaBruta - Math.abs(data.devolucoes);
    }
    
    // Se o lucro operacional (EBITDA) não foi lido diretamente da DRE, calcula por diferença
    if (data.lucroOperacional === 0) {
        data.lucroOperacional = data.receitaLiquida - Math.abs(data.cmvTotal) - Math.abs(data.pessoalTotal) - Math.abs(data.ocupacaoTotal) - Math.abs(data.despComerciais);
    }
    
    return data;
}

// Parser Numérico de Valores Excel
function parseCurrency(val) {
    if (val === undefined || val === null) return 0;
    if (typeof val === 'number') return val;
    
    let clean = String(val).replace(/R\$\s?/g, '').replace(/\s/g, '');
    
    // Tratar formatos de parênteses negativos (ex: (1.500,00))
    if (clean.startsWith('(') && clean.endsWith(')')) {
        clean = '-' + clean.substring(1, clean.length - 1);
    }
    
    // Remover pontos de milhar e trocar vírgula por ponto decimal
    clean = clean.replace(/\./g, '').replace(',', '.');
    let parsed = parseFloat(clean);
    
    return isNaN(parsed) ? 0 : parsed;
}

// Enquadramento de Cluster de Faturamento
function getClusterInfo(faturamentoBruto) {
    if (faturamentoBruto < 100000) return REFERENCIA_CLUSTERS.ATE_100K;
    if (faturamentoBruto < 150000) return REFERENCIA_CLUSTERS.ENTRE_100K_150K;
    if (faturamentoBruto < 200000) return REFERENCIA_CLUSTERS.ENTRE_150K_200K;
    if (faturamentoBruto < 250000) return REFERENCIA_CLUSTERS.ENTRE_200K_250K;
    return REFERENCIA_CLUSTERS.ACIMA_250K;
}

// Renderizar a Análise da Loja Selecionada
function renderAnalysis(loja, period) {
    if (!period) period = currentPeriod;
    const storeData = lojasProcessadas[loja];
    if (!storeData) return;
    const data = storeData.values[period];
    if (!data) return;
    
    const ref = getClusterInfo(data.receitaBruta);
    
    // Evitar divisões por zero se a receita líquida for zero
    const recLiquidaDiv = data.receitaLiquida > 0 ? data.receitaLiquida : 1;
    
    // 1. Preencher KPIs principais
    kpiFaturamento.textContent = formatCurrencyBRL(data.receitaBruta);
    kpiReceitaLiquida.textContent = formatCurrencyBRL(data.receitaLiquida);
    
    const pctEbitdaReal = (data.lucroOperacional / recLiquidaDiv) * 100;
    kpiEbitda.textContent = formatCurrencyBRL(data.lucroOperacional) + ` (${pctEbitdaReal.toFixed(2)}%)`;
    
    // Status do EBITDA
    kpiEbitdaStatus.className = "kpi-status-badge";
    if (pctEbitdaReal >= ref.meta_ebitda) {
        kpiEbitdaStatus.textContent = "Saudável";
        kpiEbitdaStatus.classList.add('healthy');
    } else {
        kpiEbitdaStatus.textContent = "Crítico";
        kpiEbitdaStatus.classList.add('critical');
    }
    
    // 2. Preencher Diagnóstico
    let diagnostic = "";
    if (pctEbitdaReal >= ref.meta_ebitda) {
        diagnostic = `A unidade **${loja}** apresentou uma performance **muito saudável** neste mês. O Lucro Operacional (EBITDA) real atingiu **${pctEbitdaReal.toFixed(2)}%**, superando a referência ideal de **${ref.meta_ebitda.toFixed(2)}%** estabelecida para o cluster **${ref.nome}**. O principal fator de sucesso foi o controle rigoroso do CMV (CMV real de ${((Math.abs(data.cmvTotal) / recLiquidaDiv) * 100).toFixed(2)}%), neutralizando desvios menores na folha de pessoal.`;
    } else {
        const gap = ref.meta_ebitda - pctEbitdaReal;
        diagnostic = `A unidade **${loja}** encontra-se em cenário **crítico de rentabilidade**, com EBITDA de **${pctEbitdaReal.toFixed(2)}%** (um gap negativo de **${gap.toFixed(2)} p.p.** em relação à meta de **${ref.meta_ebitda.toFixed(2)}%**). Os principais ralos identificados na DRE que explicam essa queda de margem são: `;
        
        let ralos = [];
        const pctRealCMV = (Math.abs(data.cmvTotal) / recLiquidaDiv) * 100;
        if (pctRealCMV > Math.abs(ref.meta_cmv)) {
            ralos.push(`estouro no CMV (${pctRealCMV.toFixed(2)}% vs. ${Math.abs(ref.meta_cmv).toFixed(2)}% ideal)`);
        }
        
        const pctRealPessoal = (Math.abs(data.pessoalTotal) / recLiquidaDiv) * 100;
        if (pctRealPessoal > Math.abs(ref.meta_pessoal)) {
            ralos.push(`descontrole de pessoal (${pctRealPessoal.toFixed(2)}% vs. ${Math.abs(ref.meta_pessoal).toFixed(2)}% ideal) com alta dependência de horas extras`);
        }
        
        diagnostic += ralos.join(" e ") + ". Ações imediatas de controle operacional são indispensáveis para reverter o cenário.";
    }
    
    // Converter Markdown básico para HTML para exibição
    diagnosticText.innerHTML = diagnostic.replace(/\*\*(.*?)\*\*/g, '<strong>$1</strong>');
    
    // Atualizar título do EBITDA para incluir o período
    const kpiEbitdaLabel = kpiEbitdaCard.querySelector('.kpi-label');
    if (kpiEbitdaLabel) {
        kpiEbitdaLabel.textContent = `EBITDA / Lucro Operacional (${period})`;
    }
    
    // 3. Montar Tabela Comparativa
    tableBody.innerHTML = "";
    
    // Contas principais para comparação
    const contasComparar = [
        { nome: "Custo de Mercadoria Vendida (CMV)", valorReal: -Math.abs(data.cmvTotal), meta: ref.meta_cmv },
        { nome: "Custo de Pessoal (Folha)", valorReal: -Math.abs(data.pessoalTotal), meta: ref.meta_pessoal },
        { nome: "Ocupação (Aluguel)", valorReal: -Math.abs(data.aluguel), meta: ref.meta_ocupacao },
        { nome: "Utilidades (Energia, Gás e Água)", valorReal: -Math.abs(data.energia + data.gas + data.agua), meta: ref.meta_utilidades }
    ];
    
    contasComparar.forEach(conta => {
        const pctRealVal = Math.abs((conta.valorReal / recLiquidaDiv) * 100);
        const metaVal = Math.abs(conta.meta);
        const desvio = pctRealVal - metaVal; // Diferença em p.p. (positivo se o custo real exceder a meta)
        
        // Impacto financeiro: Receita Líquida * (Desvio / 100)
        const impactoFinanceiro = data.receitaLiquida * (desvio / 100);
        
        // Formatar classe de status do desvio (valores positivos indicam estouro de custos)
        let desvioClass = "";
        let statusBadge = "";
        
        if (desvio > 1.5) {
            desvioClass = "up-critical";
            statusBadge = `<span class="status-badge danger">Crítico</span>`;
        } else if (desvio > 0) {
            desvioClass = "up-warning";
            statusBadge = `<span class="status-badge warning">Atenção</span>`;
        } else {
            desvioClass = "down-healthy";
            statusBadge = `<span class="status-badge success">Saudável</span>`;
        }
        
        const tr = document.createElement('tr');
        tr.innerHTML = `
            <td><strong>${conta.nome}</strong></td>
            <td class="text-right">${formatCurrencyBRL(Math.abs(conta.valorReal))}</td>
            <td class="text-right">${pctRealVal.toFixed(2)}%</td>
            <td class="text-right">${metaVal.toFixed(2)}%</td>
            <td class="text-right desvio-indicator ${desvioClass}">${desvio > 0 ? '+' : ''}${desvio.toFixed(2)}%</td>
            <td class="text-right ${desvio > 0 ? 'up-critical' : 'down-healthy'}">${desvio > 0 ? '-' : '+'}${formatCurrencyBRL(Math.abs(impactoFinanceiro))}</td>
            <td class="text-center">${statusBadge}</td>
        `;
        tableBody.appendChild(tr);
    });
    
    // 4. Montar Plano de Ação
    actionPlanContainer.innerHTML = "";
    
    const pctRealCMV = (Math.abs(data.cmvTotal) / recLiquidaDiv) * 100;
    const pctRealPessoal = (Math.abs(data.pessoalTotal) / recLiquidaDiv) * 100;
    const pctRealHorasExtras = (Math.abs(data.horasExtras) / recLiquidaDiv) * 100;
    
    // Ação para CMV
    if (pctRealCMV > Math.abs(ref.meta_cmv)) {
        addActionCard("Ação Corretiva para CMV", "critical", [
            `**Repesagem obrigatória da Mussarela:** Implementar em 100% das pizzas preparadas o porcionamento padronizado com copos medidores oficiais Domino's ou balança.`,
            `**Variance Diária:** Contar fisicamente no início e final de cada dia a mussarela, caixas, molho e proteínas para achar e eliminar o desvio (variance) de inventário.`,
            `**Auditoria de Pizza:** O gerente deve auditar a mesa de montagem de pizzas nos horários de pico para evitar desperdício de insumos no chão/mesa.`
        ]);
    } else {
        addActionCard("Manutenção Preventiva de CMV", "success", [
            `Manter os padrões saudáveis de CMV acompanhando a planilha de sugestão de pedidos para evitar obsolescência ou falta de produtos.`,
            `Auditar quinzenalmente o porcionamento na montagem para garantir a padronização das pizzas.`
        ]);
    }
    
    // Ação para Pessoal
    if (pctRealPessoal > Math.abs(ref.meta_pessoal)) {
        const excessoHorasExtras = pctRealHorasExtras > 2;
        addActionCard("Ação Corretiva para Custo de Pessoal", "critical", [
            `**Otimização de Escala de Trabalho:** Desenhar a grade de horários dos especialistas concentrando a equipe nas janelas de pico (18:00 às 22:30), reduzindo horas paradas à tarde.`,
            excessoHorasExtras ? `**Controle de Horas Extras:** Limitar horas extras a no máximo 1.5% do faturamento líquido. Toda hora extra deve ter autorização prévia por escrito da gerência regional.` : `Ajustar a escala para equilibrar a jornada dos colaboradores, reduzindo horas de ociosidade no início da semana.`,
            `**Absenteísmo:** Revisar o banco de folgas e criar política rígida de atestados para reduzir dobras e plantões de última hora.`
        ]);
    } else {
        addActionCard("Gestão de Escalas e Pessoal", "success", [
            `Manter a escala alinhada à curva diária de vendas para evitar sobrecarga de trabalho e manter o clima organizacional saudável.`,
            `Garantir que as Horas Extras não excedam o limite técnico de 1,50% do faturamento da loja.`
        ]);
    }
    
    // Ação para Utilidades e Ocupação
    addActionCard("Ações de Eficiência para Utilidades e Ocupação", "warning", [
        `**Gestão de Fornos de Esteira:** Programar o forno em modo econômico (ou desligar se o forno for duplo) durante os períodos de baixa venda no meio da tarde.`,
        `**Boracha e Vedação de Câmaras:** Agendar manutenção corretiva para verificar as portas das câmaras frias e do freezer de forma a evitar perda de temperatura e estouro na conta de energia.`,
        `**Negociação de Aluguel:** Avaliar renegociação com o locador ou expansão das vendas online para melhor diluição dos custos fixos de ocupação.`
    ]);
}

// Helper para Adicionar Cards de Ação
function addActionCard(titulo, tipo, itens) {
    const card = document.createElement('div');
    card.className = `action-card ${tipo}`;
    
    const badgeText = tipo === 'critical' ? 'Crítico' : (tipo === 'warning' ? 'Atenção' : 'Saudável');
    const badgeClass = tipo === 'critical' ? 'badge-critical' : (tipo === 'warning' ? 'status-badge warning' : 'status-badge success');
    
    let listItems = "";
    itens.forEach(item => {
        const cleanItem = item.replace(/\*\*(.*?)\*\*/g, '<strong>$1</strong>');
        listItems += `<li>${cleanItem}</li>`;
    });
    
    card.innerHTML = `
        <h4>${titulo} <span class="${badgeClass}">${badgeText}</span></h4>
        <ul>${listItems}</ul>
    `;
    actionPlanContainer.appendChild(card);
}

// Formatar moeda BRL
function formatCurrencyBRL(val) {
    return new Intl.NumberFormat('pt-BR', { style: 'currency', currency: 'BRL' }).format(val);
}

// Listener para Mudança de Loja no Select
storeSelect.addEventListener('change', (e) => {
    currentLoja = e.target.value;
    renderAnalysis(currentLoja, currentPeriod);
});

// Listener para Mudança de Período no Select
document.getElementById('period-select').addEventListener('change', (e) => {
    currentPeriod = e.target.value;
    renderAnalysis(currentLoja, currentPeriod);
});

// Ação de Impressão / PDF
document.getElementById('print-btn').addEventListener('click', () => {
    window.print();
});

// Ação de Exportar Markdown
document.getElementById('export-md-btn').addEventListener('click', () => {
    const storeData = lojasProcessadas[currentLoja];
    if (!storeData) return;
    const data = storeData.values[currentPeriod];
    if (!data) return;
    
    const ref = getClusterInfo(data.receitaBruta);
    const recLiquidaDiv = data.receitaLiquida > 0 ? data.receitaLiquida : 1;
    const pctEbitdaReal = (data.lucroOperacional / recLiquidaDiv) * 100;
    
    const pctRealCMV = (Math.abs(data.cmvTotal) / recLiquidaDiv) * 100;
    const pctRealPessoal = (Math.abs(data.pessoalTotal) / recLiquidaDiv) * 100;
    const pctRealHorasExtras = (Math.abs(data.horasExtras) / recLiquidaDiv) * 100;
    const pctRealOcupacao = (Math.abs(data.aluguel) / recLiquidaDiv) * 100;
    const pctRealUtilidades = (Math.abs(data.energia + data.gas + data.agua) / recLiquidaDiv) * 100;
    
    const desvioCMV = pctRealCMV - Math.abs(ref.meta_cmv);
    const desvioPessoal = pctRealPessoal - Math.abs(ref.meta_pessoal);
    
    const markdown = `# 📊 RELATÓRIO DE ANÁLISE GERENCIAL E OPORTUNIDADES
**Loja Analisada:** ${currentLoja}
**Mês de Referência:** ${currentPeriod}

### 1. Diagnóstico Geral (Resumo Executivo)
A unidade **${currentLoja}** fechou o mês com faturamento bruto de **${formatCurrencyBRL(data.receitaBruta)}** e receita líquida de **${formatCurrencyBRL(data.receitaLiquida)}**. O EBITDA real obtido foi de **${formatCurrencyBRL(data.lucroOperacional)} (${pctEbitdaReal.toFixed(2)}%)**, em comparação com a meta de referência de **${ref.meta_ebitda.toFixed(2)}%**. 

### 2. Análise de Desvios Críticos (Real vs. Referência)
| Conta Gerencial | % Real na Loja | % Ideal/Referência | Impacto Financeiro (R$) |
| :--- | :---: | :---: | :---: |
| **Custo de Mercadoria Vendida (CMV)** | ${pctRealCMV.toFixed(2)}% | ${Math.abs(ref.meta_cmv).toFixed(2)}% | -${formatCurrencyBRL(data.receitaLiquida * (desvioCMV / 100))} |
| **Custo de Pessoal (Total)** | ${pctRealPessoal.toFixed(2)}% | ${Math.abs(ref.meta_pessoal).toFixed(2)}% | -${formatCurrencyBRL(data.receitaLiquida * (desvioPessoal / 100))} |
| *-- Horas Extras (inclusa em Pessoal)* | ${pctRealHorasExtras.toFixed(2)}% | 1.50% | - |
| **Ocupação (Aluguel)** | ${pctRealOcupacao.toFixed(2)}% | ${Math.abs(ref.meta_ocupacao).toFixed(2)}% | - |
| **Utilidades (Energia/Gás/Água)** | ${pctRealUtilidades.toFixed(2)}% | ${Math.abs(ref.meta_utilidades).toFixed(2)}% | - |

### 3. Principais Oportunidades Identificadas
* **Oportunidade 1 (CMV):** Desvio de CMV de ${desvioCMV.toFixed(2)}% em relação ao ideal de cluster, acarretando perda líquida estimada em ${formatCurrencyBRL(data.receitaLiquida * (desvioCMV / 100))}.
* **Oportunidade 2 (Pessoal):** A conta de horas extras representou ${pctRealHorasExtras.toFixed(2)}% da receita líquida, pressionando o custo total de pessoal para ${pctRealPessoal.toFixed(2)}%.
* **Oportunidade 3 (Ocupação/Utilidades):** A fatia gasta em aluguel e concessionárias (luz, gás, água) totalizou ${(pctRealOcupacao + pctRealUtilidades).toFixed(2)}% da receita.

### 4. Plano de Ação Recomendado (Foco em Resultado)
* **Ação para CMV:** Garantir repesagem total na mesa de montagem para controle de mussarela e insumos do TOP 5, reduzindo desperdícios na produção.
* **Ação para Pessoal:** Otimizar escalas de trabalho concentrando a equipe nos picos operacionais de final de semana, diminuindo dependência de horas extras.
* **Ação para Ocupação/Utilidades:** Controlar acendimento e temperatura do forno de esteira nos períodos de vale operacional.
`;

    const blob = new Blob([markdown], { type: 'text/markdown;charset=utf-8;' });
    const link = document.createElement("a");
    link.href = URL.createObjectURL(blob);
    link.setAttribute("download", `relatorio_dre_${currentLoja.replace(/\s+/g, '_').toLowerCase()}_${currentPeriod.toLowerCase()}.md`);
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
});
