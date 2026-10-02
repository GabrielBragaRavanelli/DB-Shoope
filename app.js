// ==========================================================================
// AJBorges • Dashboard de Ocorrências Linehaul - Lógica da Aplicação (app.js)
// Conexão Read-Only em Tempo Real, Contagem Diária, Top Motoristas e Filtros
// ==========================================================================

let dadosOcorrenciasCompletos = [];
let dadosFiltrados = [];

// Instâncias dos gráficos do Chart.js
let chartDiarioInstance = null;
let chartMotivosInstance = null;
let chartMotoristasInstance = null;
let chart5DiasStackedInstance = null;

// Paleta harmoniosa e consistente com o gráfico circular
const PALETA_CORES_CIRCULAR = [
    '#ef4444', // Vermelho (ex: Transit time com gordura)
    '#f59e0b', // Laranja / Âmbar (ex: Saída antecipada CPT)
    '#2563eb', // Azul forte (ex: Transit time madrugada)
    '#10b981', // Verde Esmeralda
    '#8b5cf6', // Roxo (ex: Posto Fiscal / Multa)
    '#ec4899', // Rosa (ex: Morosidade carregamento)
    '#06b6d4', // Ciano (ex: Falta de aderência)
    '#f97316', // Laranja escuro (ex: Parada indevida)
    '#64748b', // Cinza Slate (Outros Motivos)
    '#6366f1', // Índigo
    '#14b8a6', // Teal
    '#a855f7', // Violeta
    '#eab308', // Amarelo Dourado
    '#f43f5e', // Rose
    '#0284c7'  // Sky Blue
];
const mapaCoresMotivos = {};

function obterCorMotivo(motivo) {
    if (!motivo) return '#94a3b8';
    if (!mapaCoresMotivos[motivo]) {
        const qtd = Object.keys(mapaCoresMotivos).length;
        mapaCoresMotivos[motivo] = PALETA_CORES_CIRCULAR[qtd % PALETA_CORES_CIRCULAR.length];
    }
    return mapaCoresMotivos[motivo];
}

// Intervalo de auto-refresh (2 minutos)
let intervaloAutoRefresh = null;
const TEMPO_REFRESH_MS = 120000;

// Elementos da Interface
const fileInput = document.getElementById('excel-file-input');
const dropZone = document.getElementById('drop-zone');
const statusBadge = document.getElementById('data-status-badge');
const btnAtualizarAuto = document.getElementById('btn-atualizar-auto');

const kpiTotalOcorrencias = document.getElementById('kpi-total-ocorrencias');
const kpiHojeOcorrencias = document.getElementById('kpi-hoje-ocorrencias');
const kpiHojeSub = document.getElementById('kpi-hoje-sub');
const kpiTopMotorista = document.getElementById('kpi-top-motorista');
const kpiTopMotoristaSub = document.getElementById('kpi-top-motorista-sub');
const kpiTotalDelay = document.getElementById('kpi-total-delay');
const kpiTotalParadas = document.getElementById('kpi-total-paradas');

const filtroPesquisa = document.getElementById('filtro-pesquisa');
const filtroMotorista = document.getElementById('filtro-motorista');
const filtroData = document.getElementById('filtro-data');
const filtroMotivo = document.getElementById('filtro-motivo');
const filtroStatus = document.getElementById('filtro-status');
const btnLimparFiltros = document.getElementById('btn-limpar-filtros');
const btnExportarCsv = document.getElementById('btn-exportar-csv');

const tabelaCorpo = document.getElementById('tabela-corpo');
const contadorTabela = document.getElementById('contador-tabela');

// Elementos do Gráfico e Planilha dos Últimos 5 Dias
const btnToggleGrafico5Dias = document.getElementById('btn-toggle-grafico-5dias');
const btnToggleTabela5Dias = document.getElementById('btn-toggle-tabela-5dias');
const containerGrafico5Dias = document.getElementById('container-grafico-5dias');
const containerTabela5Dias = document.getElementById('container-tabela-5dias');

const tabela5DiasCorpo = document.getElementById('tabela-5dias-corpo');
const tabela5DiasFoot = document.getElementById('tabela-5dias-foot');
const totalUltimos5DiasEl = document.getElementById('total-ultimos-5-dias');
let listaUltimas5Datas = [];

// Elementos da Modal de Ocorrências do Dia
const modalDiaOcorrencias = document.getElementById('modal-dia-ocorrencias');
const modalTituloData = document.getElementById('modal-titulo-data');
const modalSubtituloData = document.getElementById('modal-subtitulo-data');
const modalBadgeTotal = document.getElementById('modal-badge-total');
const modalMotivosResumo = document.getElementById('modal-motivos-resumo');
const modalFiltroBusca = document.getElementById('modal-filtro-busca');
const modalBtnFiltrarPainel = document.getElementById('modal-btn-filtrar-painel');
const modalTabelaCorpo = document.getElementById('modal-tabela-corpo');
const modalContadorRegistros = document.getElementById('modal-contador-registros');
const modalBtnFechar = document.getElementById('modal-btn-fechar');
const modalBtnFecharRodape = document.getElementById('modal-btn-fechar-rodape');

let ocorrenciasDiaModalAtual = [];
let dataModalAtual = '';

// ==========================================================================
// ESTADO E ELEMENTOS DA BASE GERAL DBLH (33 COLUNAS)
// ==========================================================================
let dadosDblhCompletos = [];
let dadosDblhFiltrados = [];
let paginaAtualDblh = 1;
let itensPorPaginaDblh = 100;
let totalPaginasDblh = 1;

// Abas de Navegação
const tabNavOcorrencias = document.getElementById('tab-nav-ocorrencias');
const tabNavDblh = document.getElementById('tab-nav-dblh');
const viewOcorrencias = document.getElementById('view-ocorrencias');
const viewDblh = document.getElementById('view-dblh');
const badgeContadorOcorrencias = document.getElementById('badge-contador-ocorrencias');
const badgeContadorDblh = document.getElementById('badge-contador-dblh');

// KPIs da Base DBLH
const dblhKpiTotal = document.getElementById('dblh-kpi-total');
const dblhKpiFiltradas = document.getElementById('dblh-kpi-filtradas');
const dblhKpiAbertas = document.getElementById('dblh-kpi-abertas');
const dblhKpiFechadas = document.getElementById('dblh-kpi-fechadas');
const dblhKpiOcorrencias = document.getElementById('dblh-kpi-ocorrencias');
const dblhKpiDelays = document.getElementById('dblh-kpi-delays');

// Filtros da Base DBLH
const dblhBtnLimparFiltros = document.getElementById('dblh-btn-limpar-filtros');
const dblhBtnExportarExcel = document.getElementById('dblh-btn-exportar-excel');
const dblhFiltroBusca = document.getElementById('dblh-filtro-busca');
const dblhBtnClearSearch = document.getElementById('dblh-btn-clear-search');
const dblhFiltroData = document.getElementById('dblh-filtro-data');
const dblhFiltroStatusViagem = document.getElementById('dblh-filtro-status-viagem');
const dblhFiltroStatusDestino = document.getElementById('dblh-filtro-status-destino');
const dblhFiltroStatusCpt = document.getElementById('dblh-filtro-status-cpt');
const dblhFiltroStatusEta = document.getElementById('dblh-filtro-status-eta');
const dblhFiltroOcorrencia = document.getElementById('dblh-filtro-ocorrencia');
const dblhFiltroOrigem = document.getElementById('dblh-filtro-origem');
const dblhFiltroDestino = document.getElementById('dblh-filtro-destino');
const dblhFiltroVeiculo = document.getElementById('dblh-filtro-veiculo');
const dblhFiltroMotorista = document.getElementById('dblh-filtro-motorista');
const dblhFiltroAgencia = document.getElementById('dblh-filtro-agencia');
const dblhFiltroSolicitacao = document.getElementById('dblh-filtro-solicitacao');

// Tabela e Paginação DBLH
const dblhContadorInfo = document.getElementById('dblh-contador-info');
const dblhContadorInfoBottom = document.getElementById('dblh-contador-info-bottom');
const dblhPaginaAtualEl = document.getElementById('dblh-pagina-atual');
const dblhTotalPaginasEl = document.getElementById('dblh-total-paginas');
const dblhPaginaAtualBottomEl = document.getElementById('dblh-pagina-atual-bottom');
const dblhTotalPaginasBottomEl = document.getElementById('dblh-total-paginas-bottom');
const dblhItensPorPagina = document.getElementById('dblh-itens-por-pagina');
const dblhTabelaCorpo = document.getElementById('dblh-tabela-corpo');

const dblhBtnPrimeira = document.getElementById('dblh-btn-primeira');
const dblhBtnAnterior = document.getElementById('dblh-btn-anterior');
const dblhBtnProxima = document.getElementById('dblh-btn-proxima');
const dblhBtnUltima = document.getElementById('dblh-btn-ultima');
const dblhBtnPrimeiraBottom = document.getElementById('dblh-btn-primeira-bottom');
const dblhBtnAnteriorBottom = document.getElementById('dblh-btn-anterior-bottom');
const dblhBtnProximaBottom = document.getElementById('dblh-btn-proxima-bottom');
const dblhBtnUltimaBottom = document.getElementById('dblh-btn-ultima-bottom');

// ==========================================================================
// 1. INICIALIZAÇÃO E SINCRONIZAÇÃO AUTOMÁTICA (READ-ONLY)
// ==========================================================================

document.addEventListener('DOMContentLoaded', () => {
    // Inicializa alternância de abas e eventos da DBLH
    inicializarNavegacaoAbas();
    inicializarEventosDblh();

    // Tenta carregar automaticamente ao abrir
    tentarCarregarPlanilhaOnline();

    // Configura o botão manual de atualização
    if (btnAtualizarAuto) {
        btnAtualizarAuto.addEventListener('click', () => {
            tentarCarregarPlanilhaOnline(true);
        });
    }

    // Configura polling periódico para atualizar sozinho
    intervaloAutoRefresh = setInterval(() => {
        tentarCarregarPlanilhaOnline(false);
    }, TEMPO_REFRESH_MS);
});

async function tentarCarregarPlanilhaOnline(isManualClick = false) {
    if (typeof CONFIG === 'undefined' || !CONFIG.URL_PLANILHA_OCORRENCIAS) {
        console.warn('CONFIG.URL_PLANILHA_OCORRENCIAS não definida em config.js');
        if (isManualClick) alert('Configure o link da planilha no arquivo config.js.');
        return;
    }

    if (btnAtualizarAuto) btnAtualizarAuto.classList.add('loading');
    statusBadge.textContent = 'Sincronizando...';
    statusBadge.className = 'badge-status-offline';

    try {
        // Cache-busting para garantir sempre dados frescos sem cache do navegador
        const urlComCacheBuster = CONFIG.URL_PLANILHA_OCORRENCIAS + 
            (CONFIG.URL_PLANILHA_OCORRENCIAS.includes('?') ? '&' : '?') + '_t=' + Date.now();

        const response = await fetch(urlComCacheBuster);
        if (!response.ok) {
            throw new Error(`HTTP ${response.status} - ${response.statusText}`);
        }

        const buffer = await response.arrayBuffer();
        const workbook = XLSX.read(new Uint8Array(buffer), { type: 'array' });

        // Seleciona preferencialmente a aba 'Base Tratada' ou 'BASE' ou a primeira disponível
        const sheetName = workbook.SheetNames.includes('Base Tratada') ? 'Base Tratada' :
                          (workbook.SheetNames.includes('BASE') ? 'BASE' : workbook.SheetNames[0]);

        const worksheet = workbook.Sheets[sheetName];
        const jsonData = XLSX.utils.sheet_to_json(worksheet, { defval: '' });

        const agora = new Date().toLocaleTimeString('pt-BR', { hour: '2-digit', minute: '2-digit', second: '2-digit' });
        processarDadosPlanilha(jsonData, `Google Sheets [${sheetName}] (às ${agora})`);

    } catch (error) {
        console.warn('Erro ao sincronizar:', error.message);
        if (dadosOcorrenciasCompletos.length === 0) {
            statusBadge.textContent = `Aviso: ${error.message} (Clique Atualizar)`;
            statusBadge.className = 'badge-status-offline';
        }
        if (isManualClick) {
            alert('Aviso ao sincronizar planilha: ' + error.message);
        }
    } finally {
        if (btnAtualizarAuto) btnAtualizarAuto.classList.remove('loading');
    }
}

// ==========================================================================
// 2. UPLOAD MANUAL VIA ARQUIVO (DRAG AND DROP OU INPUT)
// ==========================================================================

fileInput.addEventListener('change', (e) => {
    const file = e.target.files[0];
    if (file) processarArquivoLocal(file);
});

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
    const file = e.dataTransfer.files[0];
    if (file) processarArquivoLocal(file);
});

function processarArquivoLocal(file) {
    statusBadge.textContent = 'Processando arquivo...';
    statusBadge.className = 'badge-status-offline';

    const reader = new FileReader();
    reader.onload = (e) => {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            const sheetName = workbook.SheetNames.includes('Base Tratada') ? 'Base Tratada' :
                              (workbook.SheetNames.includes('BASE') ? 'BASE' : workbook.SheetNames[0]);
            const worksheet = workbook.Sheets[sheetName];
            const jsonData = XLSX.utils.sheet_to_json(worksheet, { defval: '' });

            processarDadosPlanilha(jsonData, `${file.name} [${sheetName}]`);
        } catch (err) {
            console.error('Erro ao ler Excel:', err);
            alert('Não foi possível ler o arquivo. Certifique-se de que é um .xlsx ou .csv válido.');
            statusBadge.textContent = 'Erro no Arquivo';
        }
    };
    reader.readAsArrayBuffer(file);
}

// ==========================================================================
// 3. EXTRAÇÃO E FILTRO RIGOROSO DAS OCORRÊNCIAS
// ==========================================================================

// Função auxiliar para normalizar e formatar datas
function normalizarData(valorData) {
    if (!valorData) return '';
    const str = String(valorData).trim();

    // Se vier com D/M/YYYY ou DD/MM/YYYY (ex: 8/10/2026 ou 08/10/2026)
    if (/^\d{1,2}\/\d{1,2}\/\d{4}/.test(str)) {
        const partes = str.substring(0, 10).split('/');
        const dia = partes[0].padStart(2, '0');
        const mes = partes[1].padStart(2, '0');
        const ano = partes[2].substring(0, 4);
        return `${dia}/${mes}/${ano}`;
    }

    // Se vier no formato YYYY-MM-DD
    if (/^\d{4}-\d{2}-\d{2}/.test(str)) {
        const partes = str.substring(0, 10).split('-');
        return `${partes[2]}/${partes[1]}/${partes[0]}`; // DD/MM/YYYY
    }

    // Se for número serial do Excel (ex: 46063)
    if (!isNaN(str) && Number(str) > 30000) {
        const d = new Date(Math.round((Number(str) - 25569) * 86400 * 1000));
        const dia = String(d.getUTCDate()).padStart(2, '0');
        const mes = String(d.getUTCMonth() + 1).padStart(2, '0');
        const ano = d.getUTCFullYear();
        return `${dia}/${mes}/${ano}`;
    }

    return str.substring(0, 10);
}

function processarDadosPlanilha(linhas, origemNome) {
    dadosOcorrenciasCompletos = [];

    linhas.forEach(linha => {
        const getVal = (colNames) => {
            for (let name of colNames) {
                for (let key in linha) {
                    if (key.trim().toLowerCase() === name.toLowerCase()) {
                        return String(linha[key]).trim();
                    }
                }
            }
            return '';
        };

        const tripNumber = getVal(['trip_number', 'trip', 'lt', 'numero_viagem']);
        const dataBruta = getVal(['data', 'eta_origin_date', 'sta_origin_date', 'data_saida', 'horario_de_descarga']);
        const statusAgrupado = getVal(['status_agrupado', 'status_viagem', 'tipo']);
        const veiculo = getVal(['vehicle_number', 'placa', 'placas']);
        const origem = getVal(['origin_station_code', 'origem']);
        const destino = getVal(['destination_station_code', 'destino']);
        const rotaFormatada = getVal(['rota']) || `${origem || 'N/D'} ➔ ${destino || 'N/D'}`;

        // Status de horários e atrasos
        const statusEtaDestino = getVal(['status_eta_destino', 'status eta destino']);
        const statusCpt = getVal(['status_cpt', 'status cpt']);
        const statusEta = getVal(['status_eta', 'status eta']);

        // Motivos de ocorrência
        const ocorrenciaDestino = getVal(['ocorrencia_eta_destino', 'ocorrência eta destino', 'ocorrencia eta destino']);
        const ocorrenciaCpt = getVal(['ocorrencia_cpt', 'ocorrência cpt', 'ocorrencia cpt']);
        const ocorrenciaEta = getVal(['ocorrencia_eta', 'ocorrência no eta', 'ocorrencia no eta']);

        // Motorista e ID do Motorista
        let motoristaBruto = getVal(['motorista', 'checkin_origin_operator', 'driver_name', 'checkout_origin_operator']);
        let driverId = getVal(['id driver', 'driver_id', 'id_driver', 'id_motorista', 'id']);
        if (!driverId && motoristaBruto) {
            const matchId = motoristaBruto.match(/\[(\d+)\]/);
            if (matchId) driverId = matchId[1];
        }

        // Limpa código entre colchetes: "[2519958]MAX EMILIANO MOURA" -> "MAX EMILIANO MOURA"
        let motoristaLimpo = motoristaBruto.replace(/\[\d+\]\s*/g, '').trim();

        // Se o nome for apenas números ou vazio, formata com o ID
        let motoristaExibicao = motoristaLimpo;
        if (!motoristaExibicao || motoristaExibicao === '-' || /^\d+$/.test(motoristaExibicao)) {
            motoristaExibicao = driverId ? `Condutor [ID ${driverId}]` : (motoristaLimpo || 'Motorista Não Informado');
        }

        // Determina se há ocorrência real e extrai o motivo principal
        let motivoFinal = '';

        if (ocorrenciaDestino && ocorrenciaDestino !== '-' && ocorrenciaDestino !== 'None' && ocorrenciaDestino !== 'null') {
            motivoFinal = ocorrenciaDestino;
        } else if (ocorrenciaCpt && ocorrenciaCpt !== '-' && ocorrenciaCpt !== 'None' && ocorrenciaCpt !== 'null') {
            motivoFinal = ocorrenciaCpt;
        } else if (ocorrenciaEta && ocorrenciaEta !== '-' && ocorrenciaEta !== 'None' && ocorrenciaEta !== 'null') {
            motivoFinal = ocorrenciaEta;
        } else if (statusEtaDestino === 'DELAY' || statusCpt === 'DELAY' || statusEta === 'DELAY') {
            motivoFinal = 'Atraso na Viagem (DELAY)';
        }

        // Se encontrou alguma ocorrência de fato
        if (motivoFinal) {
            const dataFormatada = normalizarData(dataBruta);
            dadosOcorrenciasCompletos.push({
                dataOriginal: dataBruta,
                data: dataFormatada,
                trip: tripNumber || 'Sem Trip',
                driverId: driverId ? String(driverId) : '-',
                motorista: motoristaExibicao,
                placa: veiculo || 'Sem Placa',
                origem: origem || 'N/D',
                destino: destino || 'N/D',
                rota: rotaFormatada,
                motivo: motivoFinal,
                statusDestino: statusEtaDestino || 'N/D',
                statusViagem: statusAgrupado || '-'
            });
        }
    });

    statusBadge.textContent = `🟢 Ao Vivo: ${dadosOcorrenciasCompletos.length} ocorrências (${origemNome})`;
    statusBadge.className = 'badge-status-online';

    popularFiltrosSelect();
    atualizarPlanilha5Dias(dadosOcorrenciasCompletos);
    aplicarFiltros();

    // Processamento da Base Geral DBLH (Todas as 33 Colunas)
    dadosDblhCompletos = linhas.map(l => normalizarLinhaDblh(l));
    dadosDblhFiltrados = [...dadosDblhCompletos];

    if (badgeContadorDblh) badgeContadorDblh.textContent = dadosDblhCompletos.length;
    if (badgeContadorOcorrencias) badgeContadorOcorrencias.textContent = dadosOcorrenciasCompletos.length;

    popularFiltrosDblh();
    aplicarFiltrosDblh();
}

// ==========================================================================
// 4. POPULAR FILTROS DINÂMICOS (DATA E MOTIVO)
// ==========================================================================

function popularFiltrosSelect() {
    // 1. Filtro de Datas
    const parseDate = (str) => {
        if (!str) return 0;
        const parts = str.split('/');
        return parts.length === 3 ? new Date(parts[2], parts[1] - 1, parts[0]).getTime() : 0;
    };

    const datasSet = new Set(dadosOcorrenciasCompletos.map(d => d.data).filter(Boolean));
    const datasOrdenadas = Array.from(datasSet).sort((a, b) => parseDate(b) - parseDate(a));

    listaUltimas5Datas = datasOrdenadas.slice(0, 5);

    filtroData.innerHTML = '<option value="">Todas as Datas</option>';
    if (listaUltimas5Datas.length > 1) {
        filtroData.innerHTML += '<option value="__ULTIMOS_5_DIAS__">📅 Últimos 5 Dias (Consolidado)</option>';
    }

    datasOrdenadas.forEach(data => {
        const opt = document.createElement('option');
        opt.value = data;
        opt.textContent = data;
        filtroData.appendChild(opt);
    });

    // 2. Filtro de Motivos
    const motivos = [...new Set(dadosOcorrenciasCompletos.map(d => d.motivo))].sort();
    filtroMotivo.innerHTML = '<option value="">Todos os Motivos</option>';
    motivos.forEach(motivo => {
        const opt = document.createElement('option');
        opt.value = motivo;
        opt.textContent = motivo;
        filtroMotivo.appendChild(opt);
    });

    // 3. Filtro de Motoristas (com contagem de ocorrências e ordenado de A a Z)
    if (filtroMotorista) {
        const contagemMotoristas = {};
        dadosOcorrenciasCompletos.forEach(d => {
            if (d.motorista && d.motorista !== 'Motorista Não Informado' && d.motorista !== '-') {
                contagemMotoristas[d.motorista] = (contagemMotoristas[d.motorista] || 0) + 1;
            }
        });

        const motoristasOrdenados = Object.keys(contagemMotoristas).sort((a, b) => a.localeCompare(b, 'pt-BR'));

        filtroMotorista.innerHTML = '<option value="">Todos os Motoristas</option>';
        motoristasOrdenados.forEach(motorista => {
            const qtd = contagemMotoristas[motorista];
            const opt = document.createElement('option');
            opt.value = motorista;
            opt.textContent = `${motorista} (${qtd})`;
            filtroMotorista.appendChild(opt);
        });
    }
}

// ==========================================================================
// 5. RESUMO / GRÁFICO EMPILHADO DOS ÚLTIMOS 5 DIAS E MODAL DE DETALHES
// ==========================================================================

function formatarDiaSemana(dataStr) {
    if (!dataStr) return '';
    const parts = dataStr.split('/');
    if (parts.length !== 3) return '';
    const d = new Date(Number(parts[2]), Number(parts[1]) - 1, Number(parts[0]));
    const dias = ['Dom', 'Seg', 'Ter', 'Qua', 'Qui', 'Sex', 'Sáb'];
    return dias[d.getDay()] || '';
}

function atualizarPlanilha5Dias(dados) {
    const parseDate = (str) => {
        if (!str) return 0;
        const parts = str.split('/');
        return parts.length === 3 ? new Date(parts[2], parts[1] - 1, parts[0]).getTime() : 0;
    };

    const todasDatas = [...new Set(dados.map(d => d.data).filter(Boolean))].sort((a, b) => parseDate(b) - parseDate(a));
    listaUltimas5Datas = todasDatas.slice(0, 5);

    if (listaUltimas5Datas.length === 0) {
        if (tabela5DiasCorpo) {
            tabela5DiasCorpo.innerHTML = `
                <tr>
                    <td colspan="4" class="empty-state">Nenhuma ocorrência encontrada para compor os últimos 5 dias.</td>
                </tr>
            `;
        }
        if (tabela5DiasFoot) tabela5DiasFoot.innerHTML = '';
        if (totalUltimos5DiasEl) totalUltimos5DiasEl.textContent = '0';
        if (chart5DiasStackedInstance) {
            chart5DiasStackedInstance.destroy();
            chart5DiasStackedInstance = null;
        }
        return;
    }

    const dataHoje = obterDataHojeFormatada();
    let totalAcumulado5Dias = 0;
    const consolidadoMotivos5Dias = {};

    // ----------------------------------------------------------------------
    // A) PREENCHER A TABELA RESUMO (VISÃO ALTERNATIVA)
    // ----------------------------------------------------------------------
    if (tabela5DiasCorpo) {
        tabela5DiasCorpo.innerHTML = '';

        listaUltimas5Datas.forEach(data => {
            const ocorrenciasDoDia = dados.filter(d => d.data === data);
            const qtdDia = ocorrenciasDoDia.length;
            totalAcumulado5Dias += qtdDia;

            const motivosDia = {};
            ocorrenciasDoDia.forEach(d => {
                const m = d.motivo || 'Outros';
                motivosDia[m] = (motivosDia[m] || 0) + 1;
                consolidadoMotivos5Dias[m] = (consolidadoMotivos5Dias[m] || 0) + 1;
            });

            const motivosOrdenados = Object.entries(motivosDia).sort((a, b) => b[1] - a[1]);
            const motivosHtml = motivosOrdenados.map(([motivo, count]) => {
                const cor = obterCorMotivo(motivo);
                return `
                    <span class="tag-motivo-item" title="${motivo}: ${count} registro(s)" style="border-left: 3px solid ${cor};">
                        ${motivo} <strong>${count}</strong>
                    </span>
                `;
            }).join('');

            const ehHoje = data === dataHoje;
            const diaSemana = formatarDiaSemana(data);

            const tr = document.createElement('tr');
            tr.innerHTML = `
                <td>
                    <strong style="color: #1e1e5c; font-size: 13.5px;">${data}</strong>
                    ${ehHoje ? '<span class="badge-hoje">Hoje</span>' : (diaSemana ? `<span class="badge-semana">(${diaSemana})</span>` : '')}
                </td>
                <td style="text-align: center;">
                    <span class="badge-day-count">${qtdDia}</span>
                </td>
                <td>
                    <div class="tags-motivos-list">
                        ${motivosHtml || '<span style="color: #94a3b8; font-size: 12px;">Sem detalhes</span>'}
                    </div>
                </td>
                <td style="text-align: center;">
                    <button class="btn-filter-single-day" data-date="${data}" title="Abrir motoristas e ocorrências de ${data}">
                        🔍 Abrir Detalhes
                    </button>
                </td>
            `;
            tabela5DiasCorpo.appendChild(tr);
        });

        if (totalUltimos5DiasEl) {
            totalUltimos5DiasEl.textContent = totalAcumulado5Dias;
        }

        if (tabela5DiasFoot) {
            const topMotivos5Dias = Object.entries(consolidadoMotivos5Dias)
                .sort((a, b) => b[1] - a[1])
                .slice(0, 3)
                .map(([m, c]) => `<strong>${m}</strong> (${c})`)
                .join(' • ');

            tabela5DiasFoot.innerHTML = `
                <tr class="row-total-5days">
                    <td><strong style="color: #1e1e5c; font-size: 13.5px;">TOTAL (5 DIAS)</strong></td>
                    <td style="text-align: center;">
                        <span class="total-highlight-number">${totalAcumulado5Dias}</span>
                    </td>
                    <td>
                        <div style="font-size: 12.5px; color: #334155;">
                            Principais desvios do período: ${topMotivos5Dias || '-'}
                        </div>
                    </td>
                    <td style="text-align: center;">
                        <button class="btn-filter-single-day" data-date="__ULTIMOS_5_DIAS__" title="Filtrar todos os 5 dias simultaneamente no painel">
                            🎯 Filtrar os 5 Dias
                        </button>
                    </td>
                </tr>
            `;
        }

        const botoesFiltrarDia = tabela5DiasCorpo.querySelectorAll('.btn-filter-single-day');
        botoesFiltrarDia.forEach(btn => {
            btn.addEventListener('click', () => {
                const dataEscolhida = btn.getAttribute('data-date');
                abrirModalDiaOcorrencias(dataEscolhida);
            });
        });

        const btnFooter = tabela5DiasFoot.querySelector('.btn-filter-single-day');
        if (btnFooter) {
            btnFooter.addEventListener('click', () => {
                aplicarFiltroPorDataRapida('__ULTIMOS_5_DIAS__');
            });
        }
    }

    // ----------------------------------------------------------------------
    // B) CONSTRUIR O GRÁFICO DE BARRAS EMPILHADAS (STACKED BAR CHART)
    // ----------------------------------------------------------------------
    const canvas5Dias = document.getElementById('chart-5dias-stacked');
    if (canvas5Dias) {
        // Ordena os 5 dias em ordem cronológica (da esquerda para a direita)
        const datasCronologicas = [...listaUltimas5Datas].sort((a, b) => parseDate(a) - parseDate(b));

        const labelsX = datasCronologicas.map(dia => {
            const ds = formatarDiaSemana(dia);
            return dia === dataHoje ? `${dia} (Hoje)` : `${dia} (${ds})`;
        });

        // Identifica todos os motivos únicos presentes nestes 5 dias
        const ocorrencias5Dias = dados.filter(d => listaUltimas5Datas.includes(d.data));
        const contagemGeralMotivos = {};
        ocorrencias5Dias.forEach(d => {
            contagemGeralMotivos[d.motivo] = (contagemGeralMotivos[d.motivo] || 0) + 1;
        });

        // Ordena os motivos por frequência para empilhar de forma elegante
        const motivosOrdenados = Object.keys(contagemGeralMotivos).sort((a, b) => contagemGeralMotivos[b] - contagemGeralMotivos[a]);

        const datasets = motivosOrdenados.map(motivo => {
            const cor = obterCorMotivo(motivo);
            const dataCounts = datasCronologicas.map(dia => {
                return dados.filter(d => d.data === dia && d.motivo === motivo).length;
            });

            return {
                label: motivo,
                data: dataCounts,
                backgroundColor: cor,
                borderRadius: 4,
                borderSkipped: false,
                stack: 'stack5dias'
            };
        });

        if (chart5DiasStackedInstance) {
            chart5DiasStackedInstance.destroy();
        }

        const ctx5Dias = canvas5Dias.getContext('2d');
        chart5DiasStackedInstance = new Chart(ctx5Dias, {
            type: 'bar',
            data: {
                labels: labelsX,
                datasets: datasets
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                scales: {
                    x: {
                        stacked: true,
                        grid: { display: false },
                        ticks: {
                            font: { size: 12, weight: '600' },
                            color: '#1e1e5c'
                        }
                    },
                    y: {
                        stacked: true,
                        beginAtZero: true,
                        ticks: { precision: 0 }
                    }
                },
                interaction: {
                    mode: 'index',
                    intersect: false
                },
                plugins: {
                    legend: {
                        display: true,
                        position: 'bottom',
                        labels: {
                            boxWidth: 12,
                            padding: 10,
                            font: { size: 11.5 }
                        }
                    },
                    tooltip: {
                        callbacks: {
                            footer: (tooltipItems) => {
                                const total = tooltipItems.reduce((acc, cur) => acc + cur.raw, 0);
                                return `\nTOTAL DO DIA: ${total} ocorrências\n👉 Clique nesta barra para abrir todos os motoristas`;
                            }
                        }
                    }
                },
                onClick: (event, elements) => {
                    if (elements && elements.length > 0) {
                        const index = elements[0].index;
                        const diaClicado = datasCronologicas[index];
                        abrirModalDiaOcorrencias(diaClicado);
                    }
                },
                onHover: (event, chartElement) => {
                    const canvas = event.native.target;
                    canvas.style.cursor = (chartElement && chartElement.length > 0) ? 'pointer' : 'default';
                }
            }
        });
    }
}

// --------------------------------------------------------------------------
// C) ALTERNADOR DE VISÃO (GRÁFICO EMPILHADO VS TABELA)
// --------------------------------------------------------------------------
if (btnToggleGrafico5Dias && btnToggleTabela5Dias) {
    btnToggleGrafico5Dias.addEventListener('click', () => {
        btnToggleGrafico5Dias.classList.add('active');
        btnToggleTabela5Dias.classList.remove('active');
        if (containerGrafico5Dias) containerGrafico5Dias.style.display = 'block';
        if (containerTabela5Dias) containerTabela5Dias.style.display = 'none';
        if (chart5DiasStackedInstance) chart5DiasStackedInstance.resize();
    });

    btnToggleTabela5Dias.addEventListener('click', () => {
        btnToggleTabela5Dias.classList.add('active');
        btnToggleGrafico5Dias.classList.remove('active');
        if (containerGrafico5Dias) containerGrafico5Dias.style.display = 'none';
        if (containerTabela5Dias) containerTabela5Dias.style.display = 'block';
    });
}

// --------------------------------------------------------------------------
// D) MODAL INTERATIVA COM TODAS AS OCORRÊNCIAS E MOTORISTAS DO DIA
// --------------------------------------------------------------------------
function abrirModalDiaOcorrencias(dia) {
    if (!modalDiaOcorrencias) return;

    dataModalAtual = dia;
    ocorrenciasDiaModalAtual = dadosOcorrenciasCompletos.filter(d => d.data === dia);

    const ds = formatarDiaSemana(dia);
    if (modalTituloData) {
        modalTituloData.textContent = `Ocorrências de ${dia} ${ds ? `(${ds})` : ''}`;
    }
    if (modalSubtituloData) {
        modalSubtituloData.textContent = `Listagem detalhada de todos os motoristas, viagens e desvios nesta data`;
    }
    if (modalBadgeTotal) {
        modalBadgeTotal.textContent = `${ocorrenciasDiaModalAtual.length} Ocorrências`;
    }

    // Monta o resumo dos motivos com as mesmas cores da barra
    if (modalMotivosResumo) {
        const contagemMotivos = {};
        ocorrenciasDiaModalAtual.forEach(d => {
            contagemMotivos[d.motivo] = (contagemMotivos[d.motivo] || 0) + 1;
        });

        const motivosSorted = Object.entries(contagemMotivos).sort((a, b) => b[1] - a[1]);
        modalMotivosResumo.innerHTML = motivosSorted.map(([motivo, count]) => {
            const cor = obterCorMotivo(motivo);
            return `
                <div class="modal-motivo-pill" title="${motivo}: ${count}">
                    <span class="modal-motivo-dot" style="background-color: ${cor};"></span>
                    <span>${motivo}</span>
                    <strong>${count}</strong>
                </div>
            `;
        }).join('');
    }

    if (modalFiltroBusca) modalFiltroBusca.value = '';
    renderizarLinhasModal(ocorrenciasDiaModalAtual);

    modalDiaOcorrencias.style.display = 'flex';
}

function renderizarLinhasModal(lista) {
    if (!modalTabelaCorpo) return;

    if (modalContadorRegistros) {
        modalContadorRegistros.textContent = `Exibindo ${lista.length} registro(s) de ocorrência`;
    }

    modalTabelaCorpo.innerHTML = '';

    if (lista.length === 0) {
        modalTabelaCorpo.innerHTML = `
            <tr>
                <td colspan="7" class="empty-state">Nenhum motorista ou viagem encontrado nesta pesquisa.</td>
            </tr>
        `;
        return;
    }

    lista.forEach(item => {
        const tr = document.createElement('tr');
        let badgeClasse = 'badge-default';
        if (item.statusDestino === 'DELAY') badgeClasse = 'badge-delay';
        else if (item.statusDestino === 'EARLY') badgeClasse = 'badge-early';
        else if (item.statusDestino === 'ON TIME') badgeClasse = 'badge-ontime';

        const corMotivo = obterCorMotivo(item.motivo);

        tr.innerHTML = `
            <td><span class="badge-driver-id">${item.driverId}</span></td>
            <td style="font-weight: 600; color: #1e1e5c;">${item.motorista}</td>
            <td><code>${item.trip}</code></td>
            <td>${item.rota}</td>
            <td>
                <span style="display: inline-flex; align-items: center; gap: 6px; font-weight: 600; color: #1e293b;">
                    <span style="width: 8px; height: 8px; border-radius: 50%; background: ${corMotivo}; display: inline-block;"></span>
                    ${item.motivo}
                </span>
            </td>
            <td><span class="badge-status ${badgeClasse}">${item.statusDestino}</span></td>
        `;
        modalTabelaCorpo.appendChild(tr);
    });
}

function fecharModal() {
    if (modalDiaOcorrencias) {
        modalDiaOcorrencias.style.display = 'none';
    }
}

// Configura eventos da Modal
if (modalBtnFechar) modalBtnFechar.addEventListener('click', fecharModal);
if (modalBtnFecharRodape) modalBtnFecharRodape.addEventListener('click', fecharModal);

if (modalDiaOcorrencias) {
    modalDiaOcorrencias.addEventListener('click', (e) => {
        if (e.target === modalDiaOcorrencias) fecharModal();
    });
}

window.addEventListener('keydown', (e) => {
    if (e.key === 'Escape' && modalDiaOcorrencias && modalDiaOcorrencias.style.display !== 'none') {
        fecharModal();
    }
});

if (modalFiltroBusca) {
    modalFiltroBusca.addEventListener('input', () => {
        const termo = modalFiltroBusca.value.toLowerCase().trim();
        const filtrados = ocorrenciasDiaModalAtual.filter(item => {
            return !termo ||
                item.motorista.toLowerCase().includes(termo) ||
                item.driverId.toLowerCase().includes(termo) ||
                item.trip.toLowerCase().includes(termo) ||
                item.placa.toLowerCase().includes(termo) ||
                item.motivo.toLowerCase().includes(termo);
        });
        renderizarLinhasModal(filtrados);
    });
}

if (modalBtnFiltrarPainel) {
    modalBtnFiltrarPainel.addEventListener('click', () => {
        fecharModal();
        aplicarFiltroPorDataRapida(dataModalAtual);
    });
}

function aplicarFiltroPorDataRapida(dataValor) {
    if (filtroData) {
        filtroData.value = dataValor;
        aplicarFiltros();
        const secaoTabela = document.querySelector('.table-section');
        if (secaoTabela) {
            secaoTabela.scrollIntoView({ behavior: 'smooth' });
        }
    }
}

// ==========================================================================
// 6. APLICAÇÃO DE FILTROS E ATUALIZAÇÃO DA TELA
// ==========================================================================

filtroPesquisa.addEventListener('input', aplicarFiltros);
if (filtroMotorista) filtroMotorista.addEventListener('change', aplicarFiltros);
filtroData.addEventListener('change', aplicarFiltros);
filtroMotivo.addEventListener('change', aplicarFiltros);
filtroStatus.addEventListener('change', aplicarFiltros);

btnLimparFiltros.addEventListener('click', () => {
    filtroPesquisa.value = '';
    if (filtroMotorista) filtroMotorista.value = '';
    filtroData.value = '';
    filtroMotivo.value = '';
    filtroStatus.value = '';
    aplicarFiltros();
});

function aplicarFiltros() {
    const termo = filtroPesquisa.value.toLowerCase().trim();
    const motoristaSel = filtroMotorista ? filtroMotorista.value : '';
    const dataSel = filtroData.value;
    const motivoSel = filtroMotivo.value;
    const statusSel = filtroStatus.value;

    dadosFiltrados = dadosOcorrenciasCompletos.filter(item => {
        const matchesTermo = !termo ||
            item.motorista.toLowerCase().includes(termo) ||
            item.driverId.toLowerCase().includes(termo) ||
            item.placa.toLowerCase().includes(termo) ||
            item.trip.toLowerCase().includes(termo) ||
            item.motivo.toLowerCase().includes(termo);

        const matchesMotorista = !motoristaSel || item.motorista === motoristaSel;
        const matchesData = !dataSel || 
            (dataSel === '__ULTIMOS_5_DIAS__' ? listaUltimas5Datas.includes(item.data) : item.data === dataSel);
        const matchesMotivo = !motivoSel || item.motivo === motivoSel;
        const matchesStatus = !statusSel || item.statusDestino === statusSel;

        return matchesTermo && matchesMotorista && matchesData && matchesMotivo && matchesStatus;
    });

    atualizarKPIs(dadosFiltrados);
    renderizarGraficos(dadosFiltrados);
    renderizarTabela(dadosFiltrados);
}

// ==========================================================================
// 6. ATUALIZAÇÃO DE KPIS (TOTAL, HOJE, TOP MOTORISTA, DELAYS, PARADAS)
// ==========================================================================

function obterDataHojeFormatada() {
    const hoje = new Date();
    const dia = String(hoje.getDate()).padStart(2, '0');
    const mes = String(hoje.getMonth() + 1).padStart(2, '0');
    const ano = hoje.getFullYear();
    return `${dia}/${mes}/${ano}`;
}

function atualizarKPIs(dados) {
    // 1. Total Geral Filtrado
    kpiTotalOcorrencias.textContent = dados.length;

    // 2. Ocorrências na Data de Hoje
    const dataHoje = obterDataHojeFormatada();
    const ocorrenciasHoje = dados.filter(d => d.data === dataHoje).length;
    kpiHojeOcorrencias.textContent = ocorrenciasHoje;
    kpiHojeSub.textContent = `Registradas em ${dataHoje}`;

    // 3. Motorista com Mais Ocorrências
    const contagemMotoristas = {};
    dados.forEach(d => {
        if (d.motorista && d.motorista !== 'Motorista Não Informado') {
            contagemMotoristas[d.motorista] = (contagemMotoristas[d.motorista] || 0) + 1;
        }
    });

    const motoristasSorted = Object.entries(contagemMotoristas).sort((a, b) => b[1] - a[1]);
    if (motoristasSorted.length > 0) {
        const topDriver = motoristasSorted[0];
        kpiTopMotorista.textContent = topDriver[0];
        kpiTopMotoristaSub.textContent = `${topDriver[1]} ocorrências registradas`;
        kpiTopMotorista.title = `${topDriver[0]} (${topDriver[1]} ocorrências)`;
    } else {
        kpiTopMotorista.textContent = '-';
        kpiTopMotoristaSub.textContent = 'Sem registros no momento';
    }

    // 4. Atrasos DELAY
    const totalDelay = dados.filter(d => 
        d.statusDestino === 'DELAY' || 
        d.motivo.toLowerCase().includes('delay') || 
        d.motivo.toLowerCase().includes('atraso')
    ).length;
    kpiTotalDelay.textContent = totalDelay;

    // 5. Paradas Indevidas / Longas / Abastecimento
    const totalParadas = dados.filter(d => 
        d.motivo.toLowerCase().includes('parada') || 
        d.motivo.toLowerCase().includes('abastecimento') ||
        d.motivo.toLowerCase().includes('descanso')
    ).length;
    kpiTotalParadas.textContent = totalParadas;
}

// ==========================================================================
// 7. RENDERIZAR GRÁFICOS (CHART.JS)
// ==========================================================================

function renderizarGraficos(dados) {
    const dataHoje = obterDataHojeFormatada();

    // ----------------------------------------------------------------------
    // Gráfico 1: Ocorrências por Dia (Evolução Diária / Linha do Tempo)
    // ----------------------------------------------------------------------
    const contagemPorDia = {};
    dados.forEach(d => {
        const dia = d.data || 'Sem Data';
        contagemPorDia[dia] = (contagemPorDia[dia] || 0) + 1;
    });

    // Ordenar as datas cronologicamente (da mais antiga para a mais recente)
    const datasCronologicas = Object.keys(contagemPorDia).sort((a, b) => {
        if (a === 'Sem Data') return -1;
        if (b === 'Sem Data') return 1;
        const pA = a.split('/');
        const pB = b.split('/');
        return new Date(pA[2], pA[1] - 1, pA[0]) - new Date(pB[2], pB[1] - 1, pB[0]);
    });

    const labelsDias = datasCronologicas;
    const valoresDias = datasCronologicas.map(dia => contagemPorDia[dia]);
    // Destaca a barra de hoje com cor chamativa (laranja ou azul forte)
    const coresBarrasDias = datasCronologicas.map(dia => dia === dataHoje ? '#f97316' : '#1e1e5c');

    if (chartDiarioInstance) chartDiarioInstance.destroy();
    const ctxDiario = document.getElementById('chart-diario').getContext('2d');
    chartDiarioInstance = new Chart(ctxDiario, {
        type: 'bar',
        data: {
            labels: labelsDias,
            datasets: [{
                label: 'Total de Ocorrências no Dia',
                data: valoresDias,
                backgroundColor: coresBarrasDias,
                borderRadius: 6,
                borderSkipped: false
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: { display: false },
                tooltip: {
                    callbacks: {
                        label: (context) => ` ${context.raw} ocorrências registradas`
                    }
                }
            },
            scales: {
                y: {
                    beginAtZero: true,
                    ticks: { precision: 0 }
                },
                x: {
                    grid: { display: false }
                }
            }
        }
    });

    // ----------------------------------------------------------------------
    // Gráfico 2: Principais Motivos de Ocorrência (Doughnut)
    // ----------------------------------------------------------------------
    const contagemMotivos = {};
    dados.forEach(d => {
        contagemMotivos[d.motivo] = (contagemMotivos[d.motivo] || 0) + 1;
    });

    // Ordena os 8 principais e agrupa o restante em "Outros"
    const motivosOrdenados = Object.entries(contagemMotivos).sort((a, b) => b[1] - a[1]);
    const topMotivos = motivosOrdenados.slice(0, 8);
    const outrosTotal = motivosOrdenados.slice(8).reduce((acc, cur) => acc + cur[1], 0);

    const labelsMotivos = topMotivos.map(m => m[0]);
    const dataMotivos = topMotivos.map(m => m[1]);

    if (outrosTotal > 0) {
        labelsMotivos.push('Outros Motivos');
        dataMotivos.push(outrosTotal);
    }

    if (chartMotivosInstance) chartMotivosInstance.destroy();
    const ctxMotivos = document.getElementById('chart-motivos').getContext('2d');
    chartMotivosInstance = new Chart(ctxMotivos, {
        type: 'doughnut',
        data: {
            labels: labelsMotivos,
            datasets: [{
                data: dataMotivos,
                backgroundColor: labelsMotivos.map(m => obterCorMotivo(m))
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: { position: 'bottom', labels: { boxWidth: 12, font: { size: 11 } } }
            }
        }
    });

    // ----------------------------------------------------------------------
    // Gráfico 3: Top Motoristas com Mais Ocorrências (Barras Horizontais)
    // ----------------------------------------------------------------------
    const contagemMotoristas = {};
    dados.forEach(d => {
        if (d.motorista && d.motorista !== 'Motorista Não Informado') {
            contagemMotoristas[d.motorista] = (contagemMotoristas[d.motorista] || 0) + 1;
        }
    });

    const motoristasTop10 = Object.entries(contagemMotoristas)
        .sort((a, b) => b[1] - a[1])
        .slice(0, 10);

    const labelsMotoristas = motoristasTop10.map(m => m[0]);
    const dataMotoristas = motoristasTop10.map(m => m[1]);

    if (chartMotoristasInstance) chartMotoristasInstance.destroy();
    const ctxMotoristas = document.getElementById('chart-motoristas').getContext('2d');
    chartMotoristasInstance = new Chart(ctxMotoristas, {
        type: 'bar',
        data: {
            labels: labelsMotoristas,
            datasets: [{
                label: 'Ocorrências Registradas',
                data: dataMotoristas,
                backgroundColor: '#3d3d90',
                borderRadius: 6
            }]
        },
        options: {
            indexAxis: 'y',
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: { display: false }
            },
            scales: {
                x: { ticks: { precision: 0 } },
                y: { ticks: { font: { size: 11 } } }
            }
        }
    });
}

// ==========================================================================
// 8. RENDERIZAR TABELA DE RESULTADOS
// ==========================================================================

function renderizarTabela(dados) {
    contadorTabela.textContent = dados.length;
    tabelaCorpo.innerHTML = '';

    if (dados.length === 0) {
        tabelaCorpo.innerHTML = `
            <tr>
                <td colspan="7" class="empty-state">
                    Nenhuma ocorrência encontrada com os filtros selecionados.
                </td>
            </tr>
        `;
        return;
    }

    dados.forEach(item => {
        const tr = document.createElement('tr');

        let badgeClasse = 'badge-default';
        if (item.statusDestino === 'DELAY') badgeClasse = 'badge-delay';
        else if (item.statusDestino === 'EARLY') badgeClasse = 'badge-early';
        else if (item.statusDestino === 'ON TIME') badgeClasse = 'badge-ontime';

        tr.innerHTML = `
            <td style="font-weight: 600;">${item.data || item.dataOriginal || '-'}</td>
            <td><code>${item.trip}</code></td>
            <td><span class="badge-driver-id">${item.driverId}</span></td>
            <td style="font-weight: 600; color: #1e1e5c;">${item.motorista}</td>
            <td>${item.rota}</td>
            <td style="color: #b91c1c; font-weight: 500;">${item.motivo}</td>
            <td><span class="badge-status ${badgeClasse}">${item.statusDestino}</span></td>
        `;

        tabelaCorpo.appendChild(tr);
    });
}

// ==========================================================================
// 9. EXPORTAÇÃO PARA EXCEL / CSV
// ==========================================================================

btnExportarCsv.addEventListener('click', () => {
    if (dadosFiltrados.length === 0) {
        alert('Não há dados para exportar.');
        return;
    }

    const ws = XLSX.utils.json_to_sheet(dadosFiltrados.map(d => ({
        'Data': d.data,
        'Nº Viagem (Trip)': d.trip,
        'ID Motorista': d.driverId,
        'Motorista': d.motorista,
        'Placas': d.placa,
        'Rota': d.rota,
        'Ocorrência / Motivo': d.motivo,
        'Status Destino': d.statusDestino,
        'Status Viagem': d.statusViagem
    })));

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Ocorrencias');

    // Adiciona também a aba com o resumo dos últimos 5 dias
    if (listaUltimas5Datas && listaUltimas5Datas.length > 0) {
        const linhas5DiasExcel = [];
        let totalSoma5Dias = 0;

        listaUltimas5Datas.forEach(data => {
            const itensDia = dadosOcorrenciasCompletos.filter(d => d.data === data);
            totalSoma5Dias += itensDia.length;

            const motivos = {};
            itensDia.forEach(d => {
                const m = d.motivo || 'Outros';
                motivos[m] = (motivos[m] || 0) + 1;
            });
            const motivosStr = Object.entries(motivos)
                .map(([m, qtd]) => `${m} (${qtd})`)
                .join('; ');

            linhas5DiasExcel.push({
                'Data': data,
                'Qtd Ocorrências': itensDia.length,
                'Quais foram as Ocorrências (Detalhamento)': motivosStr
            });
        });

        linhas5DiasExcel.push({
            'Data': 'TOTAL (ÚLTIMOS 5 DIAS)',
            'Qtd Ocorrências': totalSoma5Dias,
            'Quais foram as Ocorrências (Detalhamento)': 'Total consolidado do período de 5 dias'
        });

        const ws5Dias = XLSX.utils.json_to_sheet(linhas5DiasExcel);
        XLSX.utils.book_append_sheet(wb, ws5Dias, 'Resumo_Ultimos_5_Dias');
    }

    XLSX.writeFile(wb, `Relatorio_Ocorrencias_AJBorges_${obterDataHojeFormatada().replace(/\//g, '-')}.xlsx`);
});

// ==========================================================================
// 10. GESTÃO E EXPLORADOR DA BASE COMPLETA DBLH (TODAS AS 33 COLUNAS)
// ==========================================================================

// Alternador entre as Abas (Ocorrências vs DBLH Geral)
function inicializarNavegacaoAbas() {
    if (tabNavOcorrencias && tabNavDblh && viewOcorrencias && viewDblh) {
        tabNavOcorrencias.addEventListener('click', () => {
            tabNavOcorrencias.classList.add('active');
            tabNavDblh.classList.remove('active');
            viewOcorrencias.style.display = 'flex';
            viewDblh.style.display = 'none';
        });

        tabNavDblh.addEventListener('click', () => {
            tabNavDblh.classList.add('active');
            tabNavOcorrencias.classList.remove('active');
            viewDblh.style.display = 'flex';
            viewOcorrencias.style.display = 'none';
        });
    }
}

// Função auxiliar para formatar datas e horas do Excel
function formatarDataHoraExcel(val) {
    if (!val || val === '-' || val === 'null' || val === 'None') return '-';
    const str = String(val).trim();
    if (!isNaN(str) && Number(str) > 30000) {
        const d = new Date(Math.round((Number(str) - 25569) * 86400 * 1000));
        const dia = String(d.getUTCDate()).padStart(2, '0');
        const mes = String(d.getUTCMonth() + 1).padStart(2, '0');
        const ano = d.getUTCFullYear();
        const hora = String(d.getUTCHours()).padStart(2, '0');
        const min = String(d.getUTCMinutes()).padStart(2, '0');
        return `${dia}/${mes}/${ano} ${hora}:${min}`;
    }
    return str;
}

// Normalizador de cada linha da aba BASE DBLH para um objeto padronizado com as 33 colunas
function normalizarLinhaDblh(linha) {
    const getVal = (colNames) => {
        for (let name of colNames) {
            for (let key in linha) {
                if (key.trim().toLowerCase() === name.toLowerCase()) {
                    const val = linha[key];
                    return (val !== undefined && val !== null) ? String(val).trim() : '';
                }
            }
        }
        return '';
    };

    const trip_number = getVal(['trip_number', 'trip', 'lt', 'numero_viagem']);
    const status_agrupado = getVal(['status_agrupado', 'status_viagem', 'status']);
    const solicitation_by = getVal(['solicitation_by', 'solicitado_por']);
    const planned_vehicle = getVal(['planned_vehicle', 'veiculo_planejado']);
    const used_vehicle = getVal(['used_vehicle', 'veiculo_utilizado']);
    const used_agency_name = getVal(['used_agency_name', 'agencia']);

    let driver_id = getVal(['driver_id', 'id_driver', 'id driver', 'id_motorista', 'id']);
    let driver_name = getVal(['driver_name', 'motorista', 'nome_motorista']);

    if (!driver_id && driver_name) {
        const matchId = driver_name.match(/\[(\d+)\]/);
        if (matchId) driver_id = matchId[1];
    }
    let driver_name_limpo = driver_name.replace(/\[\d+\]\s*/g, '').trim();
    if (!driver_name_limpo || driver_name_limpo === '-' || /^\d+$/.test(driver_name_limpo)) {
        driver_name_limpo = (driver_id && driver_id !== '0') ? `Condutor [ID ${driver_id}]` : (driver_name_limpo || '-');
    }

    const vehicle_number = getVal(['vehicle_number', 'placa', 'placas']);
    const origin_station_code = getVal(['origin_station_code', 'origem']);
    const destination_station_code = getVal(['destination_station_code', 'destino']);
    const eta_scheduled_origin_edited = getVal(['eta_scheduled_origin_edited']);
    const cpt_scheduled_origin_edited = getVal(['cpt_scheduled_origin_edited']);
    const eta_destination_edited = getVal(['eta_destination_edited']);
    const id_rota = getVal(['id_rota', 'rota']);
    const eta_realizado = getVal(['eta_realizado']);
    const status_eta = getVal(['status_eta']);
    const ocorrencia_eta = getVal(['ocorrencia_eta', 'ocorrencia no eta']);
    const cpt_realizado = getVal(['cpt_realizado']);
    const status_cpt = getVal(['status_cpt']);
    const ocorrencia_cpt = getVal(['ocorrencia_cpt', 'ocorrencia no cpt']);
    const eta_destino_realizado = getVal(['eta_destino_realizado']);
    const status_eta_destino = getVal(['status_eta_destino']);
    const ocorrencia_eta_destino = getVal(['ocorrencia_eta_destino']);
    const horario_de_descarga = getVal(['horario_de_descarga']);
    const sum_orders = getVal(['sum_orders', 'pedidos']);
    const checkin_origin_operator = getVal(['checkin_origin_operator']);
    const checkout_origin_operator = getVal(['checkout_origin_operator']);
    const checkin_destination_operator = getVal(['checkin_destination_operator']);
    const eta_origin_realized = getVal(['eta_origin_realized']);
    const cpt_origin_realized = getVal(['cpt_origin_realized']);
    const eta_destination_realized = getVal(['eta_destination_realized']);
    const atualizacao = getVal(['atualizacao', 'atualizado']);

    // Data normalizada para filtro de dia (DD/MM/YYYY)
    const dataNormalizada = normalizarData(eta_scheduled_origin_edited) || 
                            normalizarData(cpt_scheduled_origin_edited) || 
                            normalizarData(atualizacao);

    // Tem ocorrência registrada?
    const temOcorrencia = Boolean(
        (ocorrencia_eta_destino && ocorrencia_eta_destino !== '-' && ocorrencia_eta_destino !== 'None' && ocorrencia_eta_destino !== 'null') ||
        (ocorrencia_cpt && ocorrencia_cpt !== '-' && ocorrencia_cpt !== 'None' && ocorrencia_cpt !== 'null') ||
        (ocorrencia_eta && ocorrencia_eta !== '-' && ocorrencia_eta !== 'None' && ocorrencia_eta !== 'null') ||
        status_eta_destino === 'DELAY' ||
        status_cpt === 'DELAY' ||
        status_eta === 'DELAY'
    );

    // Texto consolidado para busca rápida
    const textoBusca = [
        trip_number, status_agrupado, solicitation_by, planned_vehicle, used_vehicle,
        used_agency_name, driver_id, driver_name_limpo, vehicle_number, origin_station_code,
        destination_station_code, id_rota, status_eta, ocorrencia_eta, status_cpt,
        ocorrencia_cpt, status_eta_destino, ocorrencia_eta_destino, checkin_origin_operator,
        checkout_origin_operator, checkin_destination_operator, atualizacao
    ].join(' ').toLowerCase();

    return {
        trip_number: trip_number || '-',
        status_agrupado: status_agrupado || '-',
        solicitation_by: solicitation_by || '-',
        planned_vehicle: planned_vehicle || '-',
        used_vehicle: used_vehicle || '-',
        used_agency_name: used_agency_name || '-',
        driver_id: driver_id || '-',
        driver_name: driver_name_limpo,
        vehicle_number: vehicle_number || '-',
        origin_station_code: origin_station_code || '-',
        destination_station_code: destination_station_code || '-',
        eta_scheduled_origin_edited: formatarDataHoraExcel(eta_scheduled_origin_edited),
        cpt_scheduled_origin_edited: formatarDataHoraExcel(cpt_scheduled_origin_edited),
        eta_destination_edited: formatarDataHoraExcel(eta_destination_edited),
        id_rota: id_rota || '-',
        eta_realizado: formatarDataHoraExcel(eta_realizado),
        status_eta: status_eta || '-',
        ocorrencia_eta: ocorrencia_eta || '-',
        cpt_realizado: formatarDataHoraExcel(cpt_realizado),
        status_cpt: status_cpt || '-',
        ocorrencia_cpt: ocorrencia_cpt || '-',
        eta_destino_realizado: formatarDataHoraExcel(eta_destino_realizado),
        status_eta_destino: status_eta_destino || '-',
        ocorrencia_eta_destino: ocorrencia_eta_destino || '-',
        horario_de_descarga: formatarDataHoraExcel(horario_de_descarga),
        sum_orders: sum_orders || '-',
        checkin_origin_operator: checkin_origin_operator || '-',
        checkout_origin_operator: checkout_origin_operator || '-',
        checkin_destination_operator: checkin_destination_operator || '-',
        eta_origin_realized: formatarDataHoraExcel(eta_origin_realized),
        cpt_origin_realized: formatarDataHoraExcel(cpt_origin_realized),
        eta_destination_realized: formatarDataHoraExcel(eta_destination_realized),
        atualizacao: formatarDataHoraExcel(atualizacao),
        dataNormalizada,
        temOcorrencia,
        textoBusca
    };
}

// Popular todos os dropdowns de filtros da Base DBLH
function popularFiltrosDblh() {
    if (!dadosDblhCompletos || dadosDblhCompletos.length === 0) return;

    const preencherSelect = (el, lista, placeholder = 'Todos') => {
        if (!el) return;
        const valorAtual = el.value;
        el.innerHTML = `<option value="">${placeholder}</option>`;
        lista.forEach(val => {
            if (!val || val === '-') return;
            const opt = document.createElement('option');
            opt.value = val;
            opt.textContent = val;
            el.appendChild(opt);
        });
        if (valorAtual) el.value = valorAtual;
    };

    // 1. Datas
    const parseDate = (str) => {
        if (!str) return 0;
        const p = str.split('/');
        return p.length === 3 ? new Date(p[2], p[1] - 1, p[0]).getTime() : 0;
    };
    const datas = [...new Set(dadosDblhCompletos.map(d => d.dataNormalizada).filter(Boolean))]
        .sort((a, b) => parseDate(b) - parseDate(a));
    preencherSelect(dblhFiltroData, datas, 'Todas as Datas');

    // 2. Status Viagem
    const statusViagem = [...new Set(dadosDblhCompletos.map(d => d.status_agrupado).filter(v => v && v !== '-'))].sort();
    preencherSelect(dblhFiltroStatusViagem, statusViagem, 'Todos os Status');

    // 3. Status ETA Destino
    const statusDestino = [...new Set(dadosDblhCompletos.map(d => d.status_eta_destino).filter(v => v && v !== '-'))].sort();
    preencherSelect(dblhFiltroStatusDestino, statusDestino, 'Todos os Status Destino');

    // 4. Status CPT
    const statusCpt = [...new Set(dadosDblhCompletos.map(d => d.status_cpt).filter(v => v && v !== '-'))].sort();
    preencherSelect(dblhFiltroStatusCpt, statusCpt, 'Todos os Status CPT');

    // 5. Status ETA
    const statusEta = [...new Set(dadosDblhCompletos.map(d => d.status_eta).filter(v => v && v !== '-'))].sort();
    preencherSelect(dblhFiltroStatusEta, statusEta, 'Todos os Status ETA');

    // 6. Hub Origem
    const origens = [...new Set(dadosDblhCompletos.map(d => d.origin_station_code).filter(v => v && v !== '-'))].sort();
    preencherSelect(dblhFiltroOrigem, origens, 'Todas as Origens');

    // 7. Hub Destino
    const destinos = [...new Set(dadosDblhCompletos.map(d => d.destination_station_code).filter(v => v && v !== '-'))].sort();
    preencherSelect(dblhFiltroDestino, destinos, 'Todos os Destinos');

    // 8. Veículos (planejado ou usado)
    const veiculos = [...new Set([
        ...dadosDblhCompletos.map(d => d.planned_vehicle),
        ...dadosDblhCompletos.map(d => d.used_vehicle)
    ].filter(v => v && v !== '-'))].sort();
    preencherSelect(dblhFiltroVeiculo, veiculos, 'Todos os Veículos');

    // 9. Motoristas (com contagem de viagens)
    if (dblhFiltroMotorista) {
        const contagemMot = {};
        dadosDblhCompletos.forEach(d => {
            if (d.driver_name && d.driver_name !== '-') {
                contagemMot[d.driver_name] = (contagemMot[d.driver_name] || 0) + 1;
            }
        });
        const motoristas = Object.keys(contagemMot).sort((a, b) => a.localeCompare(b, 'pt-BR'));
        dblhFiltroMotorista.innerHTML = '<option value="">Todos os Motoristas</option>';
        motoristas.forEach(m => {
            const opt = document.createElement('option');
            opt.value = m;
            opt.textContent = `${m} (${contagemMot[m]} viagens)`;
            dblhFiltroMotorista.appendChild(opt);
        });
    }

    // 10. Agências
    const agencias = [...new Set(dadosDblhCompletos.map(d => d.used_agency_name).filter(v => v && v !== '-'))].sort();
    preencherSelect(dblhFiltroAgencia, agencias, 'Todas as Agências');

    // 11. Solicitado Por
    const solicitantes = [...new Set(dadosDblhCompletos.map(d => d.solicitation_by).filter(v => v && v !== '-'))].sort();
    preencherSelect(dblhFiltroSolicitacao, solicitantes, 'Todos os Solicitantes');
}

// Inicializar eventos de filtros e paginação da Base DBLH
function inicializarEventosDblh() {
    // Busca livre em tempo real
    if (dblhFiltroBusca) {
        dblhFiltroBusca.addEventListener('input', () => {
            if (dblhBtnClearSearch) {
                dblhBtnClearSearch.style.display = dblhFiltroBusca.value ? 'inline-block' : 'none';
            }
            paginaAtualDblh = 1;
            aplicarFiltrosDblh();
        });
    }

    if (dblhBtnClearSearch) {
        dblhBtnClearSearch.addEventListener('click', () => {
            dblhFiltroBusca.value = '';
            dblhBtnClearSearch.style.display = 'none';
            paginaAtualDblh = 1;
            aplicarFiltrosDblh();
        });
    }

    // Eventos de mudança em todos os selects da Base DBLH
    const selectsDblh = [
        dblhFiltroData, dblhFiltroStatusViagem, dblhFiltroStatusDestino,
        dblhFiltroStatusCpt, dblhFiltroStatusEta, dblhFiltroOcorrencia,
        dblhFiltroOrigem, dblhFiltroDestino, dblhFiltroVeiculo,
        dblhFiltroMotorista, dblhFiltroAgencia, dblhFiltroSolicitacao
    ];

    selectsDblh.forEach(sel => {
        if (sel) {
            sel.addEventListener('change', () => {
                paginaAtualDblh = 1;
                aplicarFiltrosDblh();
            });
        }
    });

    // Botão Limpar Filtros DBLH
    if (dblhBtnLimparFiltros) {
        dblhBtnLimparFiltros.addEventListener('click', () => {
            if (dblhFiltroBusca) dblhFiltroBusca.value = '';
            if (dblhBtnClearSearch) dblhBtnClearSearch.style.display = 'none';
            selectsDblh.forEach(sel => {
                if (sel) sel.value = '';
            });
            paginaAtualDblh = 1;
            aplicarFiltrosDblh();
        });
    }

    // Seletor de itens por página
    if (dblhItensPorPagina) {
        dblhItensPorPagina.addEventListener('change', () => {
            const val = dblhItensPorPagina.value;
            itensPorPaginaDblh = val === 'TODOS' ? 'TODOS' : Number(val);
            paginaAtualDblh = 1;
            renderizarTabelaDblh();
        });
    }

    // Botões de Paginação Superior
    if (dblhBtnPrimeira) dblhBtnPrimeira.addEventListener('click', () => mudarPaginaDblh(1));
    if (dblhBtnAnterior) dblhBtnAnterior.addEventListener('click', () => mudarPaginaDblh(paginaAtualDblh - 1));
    if (dblhBtnProxima) dblhBtnProxima.addEventListener('click', () => mudarPaginaDblh(paginaAtualDblh + 1));
    if (dblhBtnUltima) dblhBtnUltima.addEventListener('click', () => mudarPaginaDblh(totalPaginasDblh));

    // Botões de Paginação Inferior
    if (dblhBtnPrimeiraBottom) dblhBtnPrimeiraBottom.addEventListener('click', () => mudarPaginaDblh(1));
    if (dblhBtnAnteriorBottom) dblhBtnAnteriorBottom.addEventListener('click', () => mudarPaginaDblh(paginaAtualDblh - 1));
    if (dblhBtnProximaBottom) dblhBtnProximaBottom.addEventListener('click', () => mudarPaginaDblh(paginaAtualDblh + 1));
    if (dblhBtnUltimaBottom) dblhBtnUltimaBottom.addEventListener('click', () => mudarPaginaDblh(totalPaginasDblh));

    // Exportação Excel da Base DBLH (Todas as 33 Colunas)
    if (dblhBtnExportarExcel) {
        dblhBtnExportarExcel.addEventListener('click', exportarDblhExcel);
    }
}

function mudarPaginaDblh(novaPagina) {
    if (novaPagina < 1 || novaPagina > totalPaginasDblh) return;
    paginaAtualDblh = novaPagina;
    renderizarTabelaDblh();
    const tabelaSecao = document.querySelector('.dblh-table-section');
    if (tabelaSecao) {
        tabelaSecao.scrollIntoView({ behavior: 'smooth' });
    }
}

// Aplicar filtros avançados na Base DBLH
function aplicarFiltrosDblh() {
    const termo = dblhFiltroBusca ? dblhFiltroBusca.value.toLowerCase().trim() : '';
    const dataSel = dblhFiltroData ? dblhFiltroData.value : '';
    const statusViagemSel = dblhFiltroStatusViagem ? dblhFiltroStatusViagem.value : '';
    const statusDestinoSel = dblhFiltroStatusDestino ? dblhFiltroStatusDestino.value : '';
    const statusCptSel = dblhFiltroStatusCpt ? dblhFiltroStatusCpt.value : '';
    const statusEtaSel = dblhFiltroStatusEta ? dblhFiltroStatusEta.value : '';
    const ocorrenciaSel = dblhFiltroOcorrencia ? dblhFiltroOcorrencia.value : '';
    const origemSel = dblhFiltroOrigem ? dblhFiltroOrigem.value : '';
    const destinoSel = dblhFiltroDestino ? dblhFiltroDestino.value : '';
    const veiculoSel = dblhFiltroVeiculo ? dblhFiltroVeiculo.value : '';
    const motoristaSel = dblhFiltroMotorista ? dblhFiltroMotorista.value : '';
    const agenciaSel = dblhFiltroAgencia ? dblhFiltroAgencia.value : '';
    const solicitacaoSel = dblhFiltroSolicitacao ? dblhFiltroSolicitacao.value : '';

    dadosDblhFiltrados = dadosDblhCompletos.filter(item => {
        // Busca livre
        if (termo && !item.textoBusca.includes(termo)) return false;

        // Filtro Data
        if (dataSel && item.dataNormalizada !== dataSel) return false;

        // Filtro Status Viagem
        if (statusViagemSel && item.status_agrupado !== statusViagemSel) return false;

        // Filtro Status ETA Destino
        if (statusDestinoSel && item.status_eta_destino !== statusDestinoSel) return false;

        // Filtro Status CPT
        if (statusCptSel && item.status_cpt !== statusCptSel) return false;

        // Filtro Status ETA
        if (statusEtaSel && item.status_eta !== statusEtaSel) return false;

        // Filtro Ocorrência
        if (ocorrenciaSel === 'COM_OCORRENCIA' && !item.temOcorrencia) return false;
        if (ocorrenciaSel === 'SEM_OCORRENCIA' && item.temOcorrencia) return false;

        // Filtro Origem
        if (origemSel && item.origin_station_code !== origemSel) return false;

        // Filtro Destino
        if (destinoSel && item.destination_station_code !== destinoSel) return false;

        // Filtro Veículo
        if (veiculoSel && item.planned_vehicle !== veiculoSel && item.used_vehicle !== veiculoSel) return false;

        // Filtro Motorista
        if (motoristaSel && item.driver_name !== motoristaSel) return false;

        // Filtro Agência
        if (agenciaSel && item.used_agency_name !== agenciaSel) return false;

        // Filtro Solicitado Por
        if (solicitacaoSel && item.solicitation_by !== solicitacaoSel) return false;

        return true;
    });

    // Atualiza KPIs da Base DBLH
    if (dblhKpiTotal) dblhKpiTotal.textContent = dadosDblhCompletos.length;
    if (dblhKpiFiltradas) dblhKpiFiltradas.textContent = dadosDblhFiltrados.length;
    if (dblhKpiAbertas) {
        dblhKpiAbertas.textContent = dadosDblhFiltrados.filter(d => d.status_agrupado === 'ABERTA').length;
    }
    if (dblhKpiFechadas) {
        dblhKpiFechadas.textContent = dadosDblhFiltrados.filter(d => d.status_agrupado === 'FECHADA').length;
    }
    if (dblhKpiOcorrencias) {
        dblhKpiOcorrencias.textContent = dadosDblhFiltrados.filter(d => d.temOcorrencia).length;
    }
    if (dblhKpiDelays) {
        dblhKpiDelays.textContent = dadosDblhFiltrados.filter(d => 
            d.status_eta_destino === 'DELAY' || d.status_cpt === 'DELAY' || d.status_eta === 'DELAY'
        ).length;
    }

    renderizarTabelaDblh();
}

// Renderizar Tabela das 33 Colunas com Paginação
function renderizarTabelaDblh() {
    if (!dblhTabelaCorpo) return;

    const totalRegistros = dadosDblhFiltrados.length;

    if (totalRegistros === 0) {
        dblhTabelaCorpo.innerHTML = `
            <tr>
                <td colspan="33" class="empty-state">
                    Nenhuma viagem encontrada para os filtros selecionados na Base DBLH.
                </td>
            </tr>
        `;
        if (dblhContadorInfo) dblhContadorInfo.textContent = '0 viagens encontradas';
        if (dblhContadorInfoBottom) dblhContadorInfoBottom.textContent = '0 viagens encontradas';
        atualizarBotoesPaginacao(1, 1);
        return;
    }

    // Calcula páginas
    if (itensPorPaginaDblh === 'TODOS') {
        totalPaginasDblh = 1;
        paginaAtualDblh = 1;
    } else {
        totalPaginasDblh = Math.max(1, Math.ceil(totalRegistros / itensPorPaginaDblh));
        if (paginaAtualDblh > totalPaginasDblh) paginaAtualDblh = totalPaginasDblh;
    }

    const inicio = itensPorPaginaDblh === 'TODOS' ? 0 : (paginaAtualDblh - 1) * itensPorPaginaDblh;
    const fim = itensPorPaginaDblh === 'TODOS' ? totalRegistros : Math.min(totalRegistros, inicio + itensPorPaginaDblh);
    const registrosPagina = dadosDblhFiltrados.slice(inicio, fim);

    // Texto descritivo de registros
    const textoContador = `Exibindo ${inicio + 1}–${fim} de ${totalRegistros} viagens (Total Base: ${dadosDblhCompletos.length})`;
    if (dblhContadorInfo) dblhContadorInfo.textContent = textoContador;
    if (dblhContadorInfoBottom) dblhContadorInfoBottom.textContent = textoContador;

    atualizarBotoesPaginacao(paginaAtualDblh, totalPaginasDblh);

    // Constrói linhas da tabela
    const fragment = document.createDocumentFragment();

    registrosPagina.forEach(row => {
        const tr = document.createElement('tr');

        // Badge Status Viagem
        let badgeViagemHtml = row.status_agrupado;
        if (row.status_agrupado === 'ABERTA') {
            badgeViagemHtml = `<span class="badge-viagem-aberta">ABERTA</span>`;
        } else if (row.status_agrupado === 'FECHADA') {
            badgeViagemHtml = `<span class="badge-viagem-fechada">FECHADA</span>`;
        }

        // Helper para badge de status
        const formatStatusBadge = (st) => {
            if (!st || st === '-') return '-';
            if (st === 'DELAY') return `<span class="badge-status badge-delay">DELAY</span>`;
            if (st === 'ON TIME') return `<span class="badge-status badge-ontime">ON TIME</span>`;
            if (st === 'EARLY') return `<span class="badge-status badge-early">EARLY</span>`;
            return `<span class="badge-status badge-default">${st}</span>`;
        };

        // Helper para destacar ocorrências
        const formatOcorrencia = (oc) => {
            if (!oc || oc === '-' || oc === 'None' || oc === 'null') return '-';
            return `<span class="badge-ocorrencia-destaque">${oc}</span>`;
        };

        tr.innerHTML = `
            <td class="col-sticky-trip"><span class="badge-trip">${row.trip_number}</span></td>
            <td>${badgeViagemHtml}</td>
            <td>${row.solicitation_by}</td>
            <td>${row.planned_vehicle}</td>
            <td>${row.used_vehicle}</td>
            <td>${row.used_agency_name}</td>
            <td><span class="badge-driver-id">${row.driver_id}</span></td>
            <td style="font-weight: 600; color: #1e1e5c;">${row.driver_name}</td>
            <td style="font-weight: 600;">${row.vehicle_number}</td>
            <td><strong>${row.origin_station_code}</strong></td>
            <td><strong>${row.destination_station_code}</strong></td>
            <td>${row.eta_scheduled_origin_edited}</td>
            <td>${row.cpt_scheduled_origin_edited}</td>
            <td>${row.eta_destination_edited}</td>
            <td>${row.id_rota}</td>
            <td>${row.eta_realizado}</td>
            <td>${formatStatusBadge(row.status_eta)}</td>
            <td>${formatOcorrencia(row.ocorrencia_eta)}</td>
            <td>${row.cpt_realizado}</td>
            <td>${formatStatusBadge(row.status_cpt)}</td>
            <td>${formatOcorrencia(row.ocorrencia_cpt)}</td>
            <td>${row.eta_destino_realizado}</td>
            <td>${formatStatusBadge(row.status_eta_destino)}</td>
            <td>${formatOcorrencia(row.ocorrencia_eta_destino)}</td>
            <td>${row.horario_de_descarga}</td>
            <td style="text-align: center;">${row.sum_orders}</td>
            <td>${row.checkin_origin_operator}</td>
            <td>${row.checkout_origin_operator}</td>
            <td>${row.checkin_destination_operator}</td>
            <td>${row.eta_origin_realized}</td>
            <td>${row.cpt_origin_realized}</td>
            <td>${row.eta_destination_realized}</td>
            <td style="font-size: 12px; color: var(--text-muted);">${row.atualizacao}</td>
        `;

        fragment.appendChild(tr);
    });

    dblhTabelaCorpo.innerHTML = '';
    dblhTabelaCorpo.appendChild(fragment);
}

function atualizarBotoesPaginacao(atual, total) {
    if (dblhPaginaAtualEl) dblhPaginaAtualEl.textContent = atual;
    if (dblhTotalPaginasEl) dblhTotalPaginasEl.textContent = total;
    if (dblhPaginaAtualBottomEl) dblhPaginaAtualBottomEl.textContent = atual;
    if (dblhTotalPaginasBottomEl) dblhTotalPaginasBottomEl.textContent = total;

    const desabilitarVoltar = atual <= 1;
    const desabilitarAvancar = atual >= total;

    if (dblhBtnPrimeira) dblhBtnPrimeira.disabled = desabilitarVoltar;
    if (dblhBtnAnterior) dblhBtnAnterior.disabled = desabilitarVoltar;
    if (dblhBtnProxima) dblhBtnProxima.disabled = desabilitarAvancar;
    if (dblhBtnUltima) dblhBtnUltima.disabled = desabilitarAvancar;

    if (dblhBtnPrimeiraBottom) dblhBtnPrimeiraBottom.disabled = desabilitarVoltar;
    if (dblhBtnAnteriorBottom) dblhBtnAnteriorBottom.disabled = desabilitarVoltar;
    if (dblhBtnProximaBottom) dblhBtnProximaBottom.disabled = desabilitarAvancar;
    if (dblhBtnUltimaBottom) dblhBtnUltimaBottom.disabled = desabilitarAvancar;
}

// Exportação Completa das 33 Colunas da Base DBLH para Excel
function exportarDblhExcel() {
    if (!dadosDblhFiltrados || dadosDblhFiltrados.length === 0) {
        alert('Não há viagens filtradas para exportar.');
        return;
    }

    const linhasExcel = dadosDblhFiltrados.map(d => ({
        'trip_number': d.trip_number,
        'status_agrupado': d.status_agrupado,
        'solicitation_by': d.solicitation_by,
        'planned_vehicle': d.planned_vehicle,
        'used_vehicle': d.used_vehicle,
        'used_agency_name': d.used_agency_name,
        'driver_id': d.driver_id,
        'driver_name': d.driver_name,
        'vehicle_number': d.vehicle_number,
        'origin_station_code': d.origin_station_code,
        'destination_station_code': d.destination_station_code,
        'eta_scheduled_origin_edited': d.eta_scheduled_origin_edited,
        'cpt_scheduled_origin_edited': d.cpt_scheduled_origin_edited,
        'eta_destination_edited': d.eta_destination_edited,
        'id_rota': d.id_rota,
        'eta_realizado': d.eta_realizado,
        'status_eta': d.status_eta,
        'ocorrencia_eta': d.ocorrencia_eta,
        'cpt_realizado': d.cpt_realizado,
        'status_cpt': d.status_cpt,
        'ocorrencia_cpt': d.ocorrencia_cpt,
        'eta_destino_realizado': d.eta_destino_realizado,
        'status_eta_destino': d.status_eta_destino,
        'ocorrencia_eta_destino': d.ocorrencia_eta_destino,
        'horario_de_descarga': d.horario_de_descarga,
        'sum_orders': d.sum_orders,
        'checkin_origin_operator': d.checkin_origin_operator,
        'checkout_origin_operator': d.checkout_origin_operator,
        'checkin_destination_operator': d.checkin_destination_operator,
        'eta_origin_realized': d.eta_origin_realized,
        'cpt_origin_realized': d.cpt_origin_realized,
        'eta_destination_realized': d.eta_destination_realized,
        'atualizacao': d.atualizacao
    }));

    const ws = XLSX.utils.json_to_sheet(linhasExcel);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'BASE_DBLH');

    const dataHoje = obterDataHojeFormatada().replace(/\//g, '-');
    XLSX.writeFile(wb, `Base_DBLH_AJBorges_Completa_${dataHoje}.xlsx`);
}
