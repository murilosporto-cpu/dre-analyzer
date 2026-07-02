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
let activeWorkbook = null;
let activePeriods = [];

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

// Prevenir comportamento padrão do navegador de abrir arquivos arrastados para fora da zona
window.addEventListener("dragover", (e) => {
    e.preventDefault();
}, false);
window.addEventListener("drop", (e) => {
    e.preventDefault();
}, false);

// Eventos de Arrastar e Soltar na Zona de Upload
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

// Clique na Zona de Upload redireciona para o seletor invisível
dropZone.addEventListener('click', () => {
    fileInput.click();
});

fileInput.addEventListener('change', (e) => {
    if (e.target.files.length > 0) {
        handleFileSelect(e.target.files[0]);
    }
});

removeFileBtn.addEventListener('remove', clearFile);
removeFileBtn.addEventListener('click', clearFile);

function handleFileSelect(file) {
    const fileNameLower = file.name.toLowerCase();
    const validExtensions = ['.xlsx', '.xls', '.xlsm', '.xlsb', '.csv', '.ods'];
    const isValid = validExtensions.some(ext => fileNameLower.endsWith(ext));
    
    if (!isValid) {
        alert('Por favor, selecione um arquivo de planilha válido (.xlsx, .xls, .xlsm, .xlsb ou .csv).');
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
    activeWorkbook = null;
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
                activeWorkbook = workbook;
                
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

// Detecta o nome da loja a partir do conteúdo das primeiras linhas da aba caso o nome da aba seja genérico
function detectStoreNameFromCells(rows, defaultName) {
    const validStores = [
        { key: "ASA NORTE", patterns: ["asa norte", "33 norte", "33_norte"] },
        { key: "ASA SUL", patterns: ["asa sul"] },
        { key: "SUDOESTE", patterns: ["sudoeste"] },
        { key: "GOIÂNIA", patterns: ["goiania", "goiania i", "goiania ii", "goiânia"] },
        { key: "GUARÁ", patterns: ["guara", "guará"] },
        { key: "FIGUEIRAS", patterns: ["figueiras"] },
        { key: "RIO BRANCO", patterns: ["rio branco"] },
        { key: "SÃO LUIS", patterns: ["sao luis", "são luis", "são luís"] },
        { key: "AMICO", patterns: ["amico"] }
    ];
    
    // Varre as primeiras 15 linhas
    for (let r = 0; r < Math.min(rows.length, 15); r++) {
        const row = rows[r];
        if (!row) continue;
        for (let c = 0; c < row.length; c++) {
            const cellVal = String(row[c]).trim().toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "");
            if (!cellVal) continue;
            
            // Verifica se bate com algum dos nossos padrões de lojas
            for (let store of validStores) {
                for (let pattern of store.patterns) {
                    if (cellVal.includes(pattern) || pattern.includes(cellVal)) {
                        // Retorna o nome formatado da loja
                        if (cellVal.includes("goiania ii") || cellVal.includes("goiania 2") || cellVal.includes("goiânia ii")) {
                            return "GOIÂNIA II";
                        }
                        if (cellVal.includes("goiania i") || cellVal.includes("goiania 1") || cellVal.includes("goiânia i") || cellVal.includes("goiania")) {
                            return "GOIÂNIA I";
                        }
                        return store.key;
                    }
                }
            }
        }
    }
    return defaultName;
}

// Processamento da Planilha DRE
function processDREWorkbook(workbook) {
    try {
        // Apaga as informações de planilhas carregadas anteriormente
        lojasProcessadas = {};
        const sheetNames = workbook.SheetNames;
        
        const yearSelect = document.getElementById('year-select');
        const selectedYear = yearSelect ? parseInt(yearSelect.value) : new Date().getFullYear();
        const currentDate = new Date();
        const curYear = currentDate.getFullYear();
        const curMonthIdx = currentDate.getMonth(); // 0 = Janeiro, 6 = Julho
        
        const monthOrder = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"];
        let allPeriods = new Set();
        
        sheetNames.forEach(sheetName => {
            const worksheet = workbook.Sheets[sheetName];
            const rows = XLSX.utils.sheet_to_json(worksheet, {header: 1});
            
            // Tentar identificar o nome da loja
            let finalStoreName = sheetName.trim();
            
            // Se o nome da aba for genérico, tenta ler a partir das primeiras linhas da planilha
            const isGenericTab = ["sheet", "plan", "dre", "dados", "tabela", "excel", "abas", "geral", "aba"].some(g => finalStoreName.toLowerCase().includes(g)) || finalStoreName.length <= 3;
            if (isGenericTab) {
                const detectedName = detectStoreNameFromCells(rows, sheetName);
                if (detectedName) {
                    finalStoreName = detectedName;
                }
            }
            
            // Ignorar abas gerais que não são lojas individuais
            if (finalStoreName === "Resumo" || finalStoreName === "Base" || finalStoreName.startsWith("DRE ")) {
                return;
            }
            
            let lojaData = parseStoreDRE(rows);
            
            // Verificar se a loja tem faturamento em pelo menos um período
            let temFaturamento = false;
            lojaData.periods.forEach(p => {
                const monthIdx = monthOrder.indexOf(p);
                
                // Se for o ano atual (2026), ignorar meses futuros (Julho em diante, pois Julho ainda está em curso)
                if (selectedYear === curYear && monthIdx !== -1 && monthIdx >= curMonthIdx) {
                    return; // Pula mês futuro
                }
                
                if (lojaData.values[p].receitaBruta > 0) {
                    temFaturamento = true;
                    allPeriods.add(p);
                }
            });
            
            if (temFaturamento) {
                lojasProcessadas[finalStoreName] = lojaData;
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
        
        // Tentar pré-selecionar o Ano baseando-se no nome do arquivo
        if (yearSelect) {
            if (fileNameLower.includes("2024")) {
                yearSelect.value = "2024";
            } else if (fileNameLower.includes("2025")) {
                yearSelect.value = "2025";
            } else if (fileNameLower.includes("2026")) {
                yearSelect.value = "2026";
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
        
        activePeriods = Array.from(allPeriods).sort((a, b) => {
            return monthOrder.indexOf(a) - monthOrder.indexOf(b);
        });
        
        activePeriods.forEach(p => {
            const option = document.createElement('option');
            option.value = p;
            option.textContent = p;
            periodSelect.appendChild(option);
        });
        
        // Mostrar os seletores (loja e período)
        if (selectorsContainer) {
            selectorsContainer.style.display = 'flex';
        }
        
        // Definir padrões de inicialização
        currentLoja = matchedStore;
        
        // Selecionar "Maio" por padrão se disponível, senão o último mês
        let defaultPeriod = activePeriods[activePeriods.length - 1];
        if (activePeriods.includes("Maio")) {
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
    
    const validMonths = ["janeiro", "fevereiro", "marco", "abril", "maio", "junho", "julho", "agosto", "setembro", "outubro", "novembro", "dezembro"];
    const monthDisplayNames = {
        "janeiro": "Janeiro",
        "fevereiro": "Fevereiro",
        "marco": "Março",
        "abril": "Abril",
        "maio": "Maio",
        "junho": "Junho",
        "julho": "Julho",
        "agosto": "Agosto",
        "setembro": "Setembro",
        "outubro": "Outubro",
        "novembro": "Novembro",
        "dezembro": "Dezembro"
    };
    
    // Encontrar a linha de cabeçalho e a coluna dos meses
    let monthRowIndex = -1;
    
    for (let r = 0; r < Math.min(rows.length, 15); r++) {
        const row = rows[r];
        if (!row) continue;
        for (let c = 0; c < row.length; c++) {
            const cellVal = String(row[c]).trim().toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "");
            if (validMonths.includes(cellVal)) {
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
        const rawCell = String(headerRow[c]).trim();
        const cleanCell = rawCell.toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "");
        if (validMonths.includes(cleanCell)) {
            const displayName = monthDisplayNames[cleanCell];
            result.periods.push(displayName);
            result.values[displayName] = parseDREColumn(rows, c);
        }
    }
    
    return result;
}

// Processa uma coluna específica de dados para retornar as contas gerenciais daquele mês
function parseDREColumn(rows, colIndex) {
    let data = {
        receitaBruta: 0,
        receitaLiquida: 0,
        receitaServicos: 0, // Receita de Serviços (Taxa de entrega, merchandising)
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
        lucroOperacional: 0,
        linhasDRE: []
    };
    
    // Detecta se o arquivo do Excel utiliza o padrão de "=" na frente das contas totais/gerenciais
    let hasEqualsPrefix = false;
    for (let i = 0; i < Math.min(rows.length, 100); i++) {
        const row = rows[i];
        if (row && row[0] && String(row[0]).trim().startsWith("=")) {
            hasEqualsPrefix = true;
            break;
        }
    }

    let sumPessoal = 0;
    
    rows.forEach(row => {
        if (!row || row.length <= colIndex) return;
        const rawConta = String(row[0]).trim();
        const contaUpper = rawConta.toUpperCase();
        const startsWithEquals = rawConta.startsWith("=");
        const cleanContaUpper = rawConta.replace(/^=\s*/, "").toUpperCase();
        const valorVal = parseCurrency(row[colIndex]);
        const absVal = Math.abs(valorVal);
        
        // Se for uma conta gerencial total (=) ou subcategoria (-), adiciona na lista
        if (rawConta.startsWith("=") || rawConta.startsWith("-")) {
            data.linhasDRE.push({
                nome: rawConta,
                valor: valorVal,
                isMain: rawConta.startsWith("="),
                isSub: rawConta.startsWith("-")
            });
        }
        
        switch (true) {
            // 1. Receitas (Lê por similaridade, pois não há risco de conflito com subcontas)
            case contaUpper.includes("RECEITA") && contaUpper.includes("BRUTA"):
            case contaUpper.includes("FATURAMENTO BRUTO") || contaUpper.includes("VENDAS BRUTAS") || contaUpper.includes("RECEITA DE VENDAS"):
                data.receitaBruta = valorVal;
                break;
                
            case contaUpper.includes("DEVOLU") || contaUpper.includes("CANCEL"):
                data.devolucoes = valorVal;
                break;
                
            case contaUpper.includes("RECEITA") && (contaUpper.includes("LIQ") || contaUpper.includes("LÍQ") || contaUpper.includes("LQ")):
            case contaUpper.includes("RECEITA OPERACIONAL LIQUIDA") || contaUpper.includes("RECEITA OPERACIONAL LÍQUIDA"):
                data.receitaLiquida = valorVal;
                break;

            // 2. CMV Total (Prioriza linha com "=" ou faz correspondência exata para evitar CMV - Bebidas etc.)
            case startsWithEquals && (cleanContaUpper === "CMV" || cleanContaUpper.includes("CUSTO DE MERCADORIA VENDIDA") || cleanContaUpper.includes("CUSTO DA MERCADORIA VENDIDA")):
            case !startsWithEquals && (contaUpper === "CMV" || contaUpper === "CUSTO DE MERCADORIA VENDIDA (CMV)" || contaUpper === "CUSTO DA MERCADORIA VENDIDA (CMV)" || contaUpper === "CUSTO DE MERCADORIA VENDIDA" || contaUpper === "CUSTO DA MERCADORIA VENDIDA"):
                data.cmvTotal = valorVal;
                break;

            // 3. CMV Subcontas
            case contaUpper.includes("BEBIDAS"):
                data.cmvBebidas = valorVal;
                break;
            case contaUpper.includes("MASSAS"):
                data.cmvMassas = valorVal;
                break;
            case contaUpper.includes("LATIC") || contaUpper.includes("LATÍC") || contaUpper.includes("MUSSAR"):
                data.cmvLaticinios = valorVal;
                break;
            case contaUpper.includes("ALIMENT") || contaUpper.includes("INSUMO"):
                data.cmvAlimentos = valorVal;
                break;

            // 4. Pessoal Total
            case startsWithEquals && (cleanContaUpper === "PESSOAL" || cleanContaUpper.includes("CUSTO DE PESSOAL") || cleanContaUpper === "FOLHA" || cleanContaUpper.includes("CUSTO DE PESSOAL (FOLHA)")):
            case !startsWithEquals && (contaUpper === "PESSOAL" || contaUpper === "CUSTO DE PESSOAL" || contaUpper === "CUSTO DE PESSOAL (FOLHA)" || contaUpper === "FOLHA"):
                data.pessoalTotal = valorVal;
                break;

            // 5. Pessoal Subcontas
            case contaUpper.includes("SALÁR") || contaUpper.includes("SALAR") || contaUpper.includes("ORDENAD"):
                data.salarios = valorVal;
                sumPessoal += absVal;
                break;
            case contaUpper.includes("HORAS EXTRAS"):
                data.horasExtras = valorVal;
                sumPessoal += absVal;
                break;
            case contaUpper.includes("ENCARGO"):
                data.encargos = valorVal;
                sumPessoal += absVal;
                break;
            case contaUpper.includes("BENEF") || contaUpper.includes("VALE TRANSP") || contaUpper.includes("VALE REFEI") || contaUpper.includes("LANCHES") || contaUpper.includes("ASSISTENCIA M") || contaUpper.includes("ASSISTÊNCIA M"):
                data.beneficios += valorVal;
                sumPessoal += absVal;
                break;
            case contaUpper.includes("RECIS") || contaUpper.includes("RESCIS") || contaUpper.includes("FGTS") || contaUpper.includes("FÉRIAS") || contaUpper.includes("FERIAS"):
                data.rescisao += valorVal;
                sumPessoal += absVal;
                break;
            case contaUpper.includes("PRÓ-LABORE") || contaUpper.includes("PRO-LABORE") || contaUpper.includes("OUTROS - PESSOAL"):
                sumPessoal += absVal;
                break;

            // 6. Ocupação & Utilidades
            case startsWithEquals && (cleanContaUpper.includes("OCUPAÇ") || cleanContaUpper.includes("OCUPAC") || cleanContaUpper.includes("ALUGUEL") || cleanContaUpper.includes("ALUGUÉL")):
            case !startsWithEquals && (contaUpper === "ALUGUEL" || contaUpper === "ALUGUÉL" || contaUpper === "OCUPAÇÃO" || contaUpper === "CUSTO DE OCUPAÇÃO" || contaUpper === "OCUPAÇÃO (ALUGUEL)" || contaUpper === "OCUPACAO (ALUGUEL)"):
                data.aluguel = valorVal;
                data.ocupacaoTotal = valorVal;
                break;
            case contaUpper.includes("ENERGIA") && !startsWithEquals:
                data.energia = valorVal;
                break;
            case (contaUpper.includes("GÁS") || contaUpper.includes("GAS")) && !startsWithEquals:
                data.gas = valorVal;
                break;
            case (contaUpper.includes("ÁGUA") || contaUpper.includes("AGUA")) && !startsWithEquals:
                data.agua = valorVal;
                break;
            case contaUpper.includes("CUSTO UTILIDADES") || contaUpper.includes("= CUSTO UTILIDADES"):
                data.ocupacaoTotal = valorVal;
                break;

            // 7. Despesas Comerciais & EBITDA
            case contaUpper.includes("DESPESAS COMERCIAIS") || contaUpper.includes("MARKETING") || contaUpper.includes("PROPAGANDA"):
                data.despComerciais += valorVal;
                break;
            case startsWithEquals && (cleanContaUpper === "EBITDA" || cleanContaUpper.includes("RESULTADO OPERACIONAL") || cleanContaUpper.includes("LUCRO OPERACIONAL")):
            case !startsWithEquals && (contaUpper === "EBITDA" || contaUpper === "RESULTADO OPERACIONAL" || contaUpper === "LUCRO OPERACIONAL" || contaUpper.includes("LUCRO OPERACIONAL (EBITDA)") || contaUpper.includes("= RESULTADO OPERACIONAL") || contaUpper.includes("= EBITDA")):
                data.lucroOperacional = valorVal;
                break;
                
            // 8. Receita de Serviços
            case startsWithEquals && (cleanContaUpper.includes("RECEITA SERVIÇO") || cleanContaUpper.includes("RECEITA SERVICO") || cleanContaUpper.includes("TAXA DE ENTREGA") || cleanContaUpper.includes("MERCHANDISING")):
            case !startsWithEquals && (contaUpper.includes("RECEITA SERVIÇO") || contaUpper.includes("RECEITA SERVICO") || contaUpper.includes("TAXA DE ENTREGA") || contaUpper.includes("MERCHANDISING")):
                data.receitaServicos = valorVal;
                break;
        }
    });
    
    // Acrescenta a Receita de Serviços aos totais de faturamento bruto, receita líquida e ebitda
    if (data.receitaServicos !== 0) {
        data.receitaBruta += data.receitaServicos;
        data.receitaLiquida += data.receitaServicos;
        data.lucroOperacional += data.receitaServicos;
    }
    
    // Atualizar pessoalTotal com a soma das subcontas se não tiver sido extraído diretamente
    if (data.pessoalTotal === 0) {
        data.pessoalTotal = -sumPessoal;
    }
    
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
    
    // 2. Preencher Diagnóstico
    let diagnostic = "";
    if (pctEbitdaReal >= ref.meta_ebitda) {
        diagnostic = `A unidade **${loja}** apresentou uma performance **muito saudável** neste mês. O Lucro Operacional (EBITDA) real atingiu **${pctEbitdaReal.toFixed(2)}%**, superando a referência ideal de **${ref.meta_ebitda.toFixed(2)}%** estabelecida para o cluster **${ref.nome}**. O principal fator de sucesso foi o controle rigoroso do CMV (CMV real de ${((Math.abs(data.cmvTotal) / recLiquidaDiv) * 100).toFixed(2)}%), neutralizando desvios menores na folha de pessoal.`;
    } else if (pctEbitdaReal > 0) {
        const gap = ref.meta_ebitda - pctEbitdaReal;
        diagnostic = `A unidade **${loja}** apresentou resultado **positivo, mas em atenção**, com EBITDA de **${pctEbitdaReal.toFixed(2)}%** (um gap de **${gap.toFixed(2)} p.p.** abaixo da meta de **${ref.meta_ebitda.toFixed(2)}%** estabelecida para o cluster **${ref.nome}**). Os desvios que explicam esse resultado são: `;
        
        let ralos = [];
        const pctRealCMV = (Math.abs(data.cmvTotal) / recLiquidaDiv) * 100;
        if (pctRealCMV > Math.abs(ref.meta_cmv)) {
            ralos.push(`estouro no CMV (${pctRealCMV.toFixed(2)}% vs. ${Math.abs(ref.meta_cmv).toFixed(2)}% ideal)`);
        }
        
        const pctRealPessoal = (Math.abs(data.pessoalTotal) / recLiquidaDiv) * 100;
        if (pctRealPessoal > Math.abs(ref.meta_pessoal)) {
            ralos.push(`descontrole de pessoal (${pctRealPessoal.toFixed(2)}% vs. ${Math.abs(ref.meta_pessoal).toFixed(2)}% ideal)`);
        }
        
        diagnostic += ralos.join(" e ") + ". É necessário focar nas oportunidades de economia para atingir a meta do cluster.";
    } else {
        const gap = ref.meta_ebitda - pctEbitdaReal;
        diagnostic = `A unidade **${loja}** encontra-se em cenário **crítico de rentabilidade** com EBITDA negativo de **${pctEbitdaReal.toFixed(2)}%** (um gap de **${gap.toFixed(2)} p.p.** em relação à meta de **${ref.meta_ebitda.toFixed(2)}%**). Os principais ralos identificados na DRE que explicam essa queda de margem são: `;
        
        let ralos = [];
        const pctRealCMV = (Math.abs(data.cmvTotal) / recLiquidaDiv) * 100;
        if (pctRealCMV > Math.abs(ref.meta_cmv)) {
            ralos.push(`estouro no CMV (${pctRealCMV.toFixed(2)}% vs. ${Math.abs(ref.meta_cmv).toFixed(2)}% ideal)`);
        }
        
        const pctRealPessoal = (Math.abs(data.pessoalTotal) / recLiquidaDiv) * 100;
        if (pctRealPessoal > Math.abs(ref.meta_pessoal)) {
            ralos.push(`descontrole de pessoal (${pctRealPessoal.toFixed(2)}% vs. ${Math.abs(ref.meta_pessoal).toFixed(2)}% ideal)`);
        }
        
        diagnostic += ralos.join(" e ") + ". Ações imediatas de controle operacional são indispensáveis para reverter o cenário.";
    }
    
    // Converter Markdown básico para HTML para exibição
    diagnosticText.innerHTML = diagnostic.replace(/\*\*(.*?)\*\*/g, '<strong>$1</strong>');
    
    // 3. Montar Tabela Comparativa
    tableBody.innerHTML = "";
    
    // Exibe estritamente as contas gerenciais oficiais que constam na planilha de referência
    const contasExibir = [
        { nome: "Fat. Bruto", valorReal: data.receitaBruta, meta: null, isMain: true, isSub: false },
        { nome: "Rec. Líquida", valorReal: data.receitaLiquida, meta: null, isMain: true, isSub: false },
        { nome: "Rec. Serviços", valorReal: data.receitaServicos, meta: null, isMain: true, isSub: false },
        { nome: "CMV", valorReal: -Math.abs(data.cmvTotal), meta: ref.meta_cmv, isMain: true, isSub: false },
        { nome: "Pessoal", valorReal: -Math.abs(data.pessoalTotal), meta: ref.meta_pessoal, isMain: true, isSub: false },
        { nome: "Ocupação", valorReal: -Math.abs(data.aluguel), meta: ref.meta_ocupacao, isMain: true, isSub: false },
        { nome: "Utilidades", valorReal: -Math.abs(data.energia + data.gas + data.agua), meta: ref.meta_utilidades, isMain: true, isSub: false },
        { nome: "EBITDA", valorReal: data.lucroOperacional, meta: ref.meta_ebitda, isMain: true, isSub: false }
    ];
    
    contasExibir.forEach(conta => {
        // Ignora contas gerenciais que estão zeradas no relatório do período selecionado
        if (conta.valorReal === 0) return;
        
        const pctRealVal = Math.abs((conta.valorReal / recLiquidaDiv) * 100);
        
        let metaValStr = "-";
        let desvioStr = "-";
        let desvioClass = "";
        let impactoStr = "-";
        let statusBadge = "-";
        
        if (conta.meta !== null) {
            const metaVal = Math.abs(conta.meta);
            let desvio = 0;
            let impactoFinanceiro = 0;
            
            if (conta.nome.toUpperCase().includes("EBITDA") || conta.nome.toUpperCase().includes("RESULTADO OPERACIONAL")) {
                const realEbitdaPct = (conta.valorReal / recLiquidaDiv) * 100;
                desvio = realEbitdaPct - conta.meta;
                impactoFinanceiro = data.receitaLiquida * (desvio / 100);
                
                metaValStr = `${conta.meta.toFixed(2)}%`;
                desvioStr = `${desvio > 0 ? '+' : ''}${desvio.toFixed(2)}%`;
                
                if (desvio < -1.5) {
                    desvioClass = "up-critical";
                    statusBadge = `<span class="status-badge danger">Crítico</span>`;
                } else if (desvio < 0) {
                    desvioClass = "up-warning";
                    statusBadge = `<span class="status-badge warning">Atenção</span>`;
                } else {
                    desvioClass = "down-healthy";
                    statusBadge = `<span class="status-badge success">Saudável</span>`;
                }
                
                impactoStr = `${impactoFinanceiro > 0 ? '+' : ''}${formatCurrencyBRL(impactoFinanceiro)}`;
            } else {
                desvio = pctRealVal - metaVal;
                impactoFinanceiro = data.receitaLiquida * (desvio / 100);
                
                metaValStr = `${metaVal.toFixed(2)}%`;
                desvioStr = `${desvio > 0 ? '+' : ''}${desvio.toFixed(2)}%`;
                
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
                
                impactoStr = `${desvio > 0 ? '-' : '+'}${formatCurrencyBRL(Math.abs(impactoFinanceiro))}`;
            }
        }
        
        let displayPctReal = `${pctRealVal.toFixed(2)}%`;
        if (conta.nome.toUpperCase().includes("EBITDA") || conta.nome.toUpperCase().includes("RESULTADO OPERACIONAL")) {
            const realEbitdaPct = (conta.valorReal / recLiquidaDiv) * 100;
            displayPctReal = `${realEbitdaPct.toFixed(2)}%`;
        }
        
        const tr = document.createElement('tr');
        if (conta.isMain) {
            tr.style.fontWeight = "bold";
            tr.style.backgroundColor = "rgba(0, 100, 145, 0.03)";
            if (conta.nome.toUpperCase().includes("EBITDA") || conta.nome.toUpperCase().includes("RESULTADO OPERACIONAL")) {
                tr.style.backgroundColor = "rgba(0, 100, 145, 0.08)";
            }
        } else if (conta.isSub) {
            tr.style.fontSize = "0.85rem";
            tr.style.color = "var(--text-muted)";
        }
        
        tr.innerHTML = `
            <td style="${conta.isSub ? 'padding-left: 2rem;' : ''}"><strong>${conta.nome}</strong></td>
            <td class="text-right">${formatCurrencyBRL(Math.abs(conta.valorReal))}</td>
            <td class="text-right">${displayPctReal}</td>
            <td class="text-right">${metaValStr}</td>
            <td class="text-right desvio-indicator ${desvioClass}">${desvioStr}</td>
            <td class="text-right ${desvioClass}">${impactoStr}</td>
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

    // 5. Montar Comparativo Mensal (Real)
    const sortedActivePeriods = activePeriods;

    const monthAbbreviations = {
        "Janeiro": "Jan",
        "Fevereiro": "Fev",
        "Março": "Mar",
        "Abril": "Abr",
        "Maio": "Mai",
        "Junho": "Jun",
        "Julho": "Jul",
        "Agosto": "Ago",
        "Setembro": "Set",
        "Outubro": "Out",
        "Novembro": "Nov",
        "Dezembro": "Dez"
    };

    const headersRow = document.getElementById('monthly-table-headers');
    if (headersRow) {
        headersRow.innerHTML = "<th>Conta</th>";
        sortedActivePeriods.forEach(p => {
            const abbrev = monthAbbreviations[p] || p;
            headersRow.innerHTML += `<th class="text-right">${abbrev} (R$)</th><th class="text-right">%</th>`;
        });
    }

    const rowsConfig = [
        { label: "Fat. Bruto", getValue: (d) => d.receitaBruta, format: "currency" },
        { label: "Rec. Líquida", getValue: (d) => d.receitaLiquida, format: "currency" },
        { label: "Rec. Serviços", getValue: (d) => d.receitaServicos, format: "currency" },
        { label: "CMV", getValue: (d) => -Math.abs(d.cmvTotal), format: "currency" },
        { label: "Pessoal", getValue: (d) => -Math.abs(d.pessoalTotal), format: "currency" },
        { label: "Ocupação", getValue: (d) => -Math.abs(d.aluguel), format: "currency" },
        { label: "Utilidades", getValue: (d) => -Math.abs(d.energia + d.gas + d.agua), format: "currency" },
        { label: "EBITDA", getValue: (d) => d.lucroOperacional, format: "currency" },
        { 
            label: "Margem EBITDA", 
            getValue: (d) => {
                const div = d.receitaLiquida > 0 ? d.receitaLiquida : 1;
                return (d.lucroOperacional / div) * 100;
            }, 
            format: "percent" 
        }
    ];

    const monthlyBody = document.getElementById('monthly-comparison-body');
    if (monthlyBody) {
        monthlyBody.innerHTML = "";
        
        rowsConfig.forEach(row => {
            const tr = document.createElement('tr');
            
            if (row.label.includes("EBITDA")) {
                tr.style.fontWeight = "bold";
                tr.style.backgroundColor = "rgba(0, 100, 145, 0.05)";
            }
            
            tr.innerHTML = `<td><strong>${row.label}</strong></td>`;
            
            sortedActivePeriods.forEach(p => {
                const dataVal = storeData.values[p];
                let formattedVal = "-";
                let formattedPct = "-";
                
                if (dataVal) {
                    const val = row.getValue(dataVal);
                    const denom = dataVal.receitaBruta > 0 ? dataVal.receitaBruta : 1;
                    const pctVal = (val / denom) * 100;
                    
                    if (row.format === "currency") {
                        formattedVal = formatCurrencyBRL(val);
                        if (val < 0) {
                            formattedVal = `<span class="negative-value">${formattedVal}</span>`;
                        }
                        
                        if (row.label === "Faturamento Bruto") {
                            formattedPct = "100.0%";
                        } else {
                            formattedPct = `${pctVal.toFixed(1)}%`;
                            if (pctVal < 0) {
                                formattedPct = `<span class="negative-value">${formattedPct}</span>`;
                            }
                        }
                    } else {
                        // Linha que já é percentual (Margem EBITDA)
                        formattedVal = `${val.toFixed(2)}%`;
                        if (val < 0) {
                            formattedVal = `<span class="negative-value">${formattedVal}</span>`;
                        }
                        formattedPct = ""; // não se aplica percentual redundante
                    }
                }
                
                tr.innerHTML += `<td class="text-right">${formattedVal}</td>`;
                if (row.format === "currency") {
                    tr.innerHTML += `<td class="text-right">${formattedPct}</td>`;
                } else {
                    tr.innerHTML += `<td class="text-right">-</td>`;
                }
            });
            
            tr.style.borderBottom = "1px solid var(--border-color)";
            monthlyBody.appendChild(tr);
        });
    }

    // 6. Montar Comparativo Mensal (%)
    const headersPctRow = document.getElementById('monthly-pct-table-headers');
    if (headersPctRow) {
        headersPctRow.innerHTML = "<th>Conta</th>";
        sortedActivePeriods.forEach(p => {
            const abbrev = monthAbbreviations[p] || p;
            headersPctRow.innerHTML += `<th class="text-right">${abbrev}</th>`;
        });
    }

    const rowsPctConfig = [
        { label: "Fat. Bruto", getValue: (d) => d.receitaBruta, format: "currency" },
        { label: "Rec. Líquida", getValue: (d) => d.receitaLiquida, format: "currency" },
        { label: "Rec. Serviços", getValue: (d) => d.receitaServicos, format: "currency" },
        { label: "CMV", getValue: (d) => -Math.abs(d.cmvTotal), format: "percent_of_sales" },
        { label: "Pessoal", getValue: (d) => -Math.abs(d.pessoalTotal), format: "percent_of_sales" },
        { label: "Ocupação", getValue: (d) => -Math.abs(d.aluguel), format: "percent_of_sales" },
        { label: "Utilidades", getValue: (d) => -Math.abs(d.energia + d.gas + d.agua), format: "percent_of_sales" },
        { label: "EBITDA", getValue: (d) => d.lucroOperacional, format: "percent_of_sales" }
    ];

    const monthlyPctBody = document.getElementById('monthly-pct-comparison-body');
    if (monthlyPctBody) {
        monthlyPctBody.innerHTML = "";
        
        rowsPctConfig.forEach(row => {
            const tr = document.createElement('tr');
            
            if (row.label.includes("EBITDA")) {
                tr.style.fontWeight = "bold";
                tr.style.backgroundColor = "rgba(0, 100, 145, 0.05)";
            }
            
            tr.innerHTML = `<td><strong>${row.label}</strong></td>`;
            
            sortedActivePeriods.forEach(p => {
                const dataVal = storeData.values[p];
                let formattedVal = "-";
                
                if (dataVal) {
                    const val = row.getValue(dataVal);
                    
                    if (row.format === "currency") {
                        formattedVal = formatCurrencyBRL(val);
                        if (val < 0) {
                            formattedVal = `<span class="negative-value">${formattedVal}</span>`;
                        }
                    } else {
                        // Calcula % baseado no Faturamento Bruto (Vendas)
                        const denom = dataVal.receitaBruta > 0 ? dataVal.receitaBruta : 1;
                        const pctVal = (val / denom) * 100;
                        formattedVal = `${pctVal.toFixed(1)}%`;
                        if (pctVal < 0) {
                            formattedVal = `<span class="negative-value">${formattedVal}</span>`;
                        }
                    }
                }
                
                tr.innerHTML += `<td class="text-right">${formattedVal}</td>`;
            });
            
            tr.style.borderBottom = "1px solid var(--border-color)";
            monthlyPctBody.appendChild(tr);
        });
    }
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

// Listener para Mudança de Ano no Select
const yearSelectElement = document.getElementById('year-select');
if (yearSelectElement) {
    yearSelectElement.addEventListener('change', (e) => {
        if (activeWorkbook) {
            processDREWorkbook(activeWorkbook);
        }
    });
}

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
