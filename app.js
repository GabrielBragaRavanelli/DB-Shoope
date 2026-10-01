// ==========================================================================
// AJBorges • Dashboard de Ocorrências Linehaul - Lógica da Aplicação (app.js)
// ==========================================================================

let dadosOcorrenciasCompletos = [];
let dadosFiltrados = [];

let chartMotivosInstance = null;
let chartMotoristasInstance = null;

// Elementos da Interface
const fileInput = document.getElementById('excel-file-input');
const dropZone = document.getElementById('drop-zone');
const statusBadge = document.getElementById('data-status-badge');

const kpiTotalOcorrencias = document.getElementById('kpi-total-ocorrencias');
const kpiTotalDelay = document.getElementById('kpi-total-delay');
const kpiTotalParadas = document.getElementById('kpi-total-paradas');
const kpiTotalMotoristas = document.getElementById('kpi-total-motoristas');

const filtroPesquisa = document.getElementById('filtro-pesquisa');
const filtroMotivo = document.getElementById('filtro-motivo');
const filtroStatus = document.getElementById('filtro-status');
const btnLimparFiltros = document.getElementById('btn-limpar-filtros');
const btnExportarCsv = document.getElementById('btn-exportar-csv');

const tabelaCorpo = document.getElementById('tabela-corpo');
const contadorTabela = document.getElementById('contador-tabela');

// ==========================================================================
// 1. UPLOAD E LEITURA DO EXCEL (SheetJS)
// ==========================================================================

fileInput.addEventListener('change', (e) => {
    const file = e.target.files[0];
    if (file) processarArquivoExcel(file);
});

// Drag and drop
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
    if (file) processarArquivoExcel(file);
});

function processarArquivoExcel(file) {
    statusBadge.textContent = 'Processando...';
    statusBadge.className = 'badge-status-offline';

    const reader = new FileReader();
    reader.onload = (e) => {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });

            // Pega a primeira aba da planilha
            const firstSheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[firstSheetName];

            // Converte para JSON
            const jsonData = XLSX.utils.sheet_to_json(worksheet, { defval: '' });

            extrairOcorrencias(jsonData, file.name);
        } catch (err) {
            console.error('Erro ao ler Excel:', err);
            alert('Não foi possível ler a planilha. Verifique se é um arquivo .xlsx ou .csv válido.');
            statusBadge.textContent = 'Erro ao Carregar';
        }
    };
    reader.readAsArrayBuffer(file);
}

// ==========================================================================
// 2. EXTRAÇÃO E FILTRO DAS OCORRÊNCIAS
// ==========================================================================

function extrairOcorrencias(linhas, fileName) {
    dadosOcorrenciasCompletos = [];

    linhas.forEach(linha => {
        // Mapeia colunas da planilha DBLH_AJ BORGES de forma flexível (sem se preocupar com maiúsculas/minúsculas)
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

        const tripNumber = getVal(['trip_number', 'trip', 'numero_viagem']);
        const dataOrigem = getVal(['eta_origin_date', 'data', 'data_saida', 'horario_de_descarga']);
        const statusAgrupado = getVal(['status_agrupado', 'status_viagem']);
        const veiculo = getVal(['vehicle_number', 'placa', 'placas']);
        const origem = getVal(['origin_station_code', 'origem']);
        const destino = getVal(['destination_station_code', 'destino']);

        // Status de horários e atrasos
        const statusEtaDestino = getVal(['status_eta_destino']);
        const statusCpt = getVal(['status_cpt']);

        // Motivos de ocorrência
        const ocorrenciaDestino = getVal(['ocorrencia_eta_destino']);
        const ocorrenciaCpt = getVal(['ocorrencia_cpt']);
        const ocorrenciaEta = getVal(['ocorrencia_eta']);

        // Motorista (procura na coluna AB de operador ou driver_name)
        let motoristaBruto = getVal(['checkin_origin_operator', 'driver_name', 'motorista', 'checkout_origin_operator']);
        // Limpa código entre colchetes ex: "[2866581]DOUGLAS DOS SANTOS" -> "DOUGLAS DOS SANTOS"
        let motoristaLimpo = motoristaBruto.replace(/\[\d+\]\s*/g, '').trim();

        // Determina se a linha tem ocorrência de verdade
        let motivoFinal = '';

        if (ocorrenciaDestino && ocorrenciaDestino !== '-' && ocorrenciaDestino !== 'None') {
            motivoFinal = ocorrenciaDestino;
        } else if (statusEtaDestino === 'DELAY' || statusCpt === 'DELAY') {
            motivoFinal = 'Atraso na Viagem (DELAY)';
        } else if (ocorrenciaCpt && ocorrenciaCpt !== '-' && ocorrenciaCpt !== 'None') {
            motivoFinal = ocorrenciaCpt;
        } else if (ocorrenciaEta && ocorrenciaEta !== '-' && ocorrenciaEta !== 'None') {
            motivoFinal = ocorrenciaEta;
        }

        // Se encontrou alguma ocorrência ou atraso real
        if (motivoFinal) {
            dadosOcorrenciasCompletos.push({
                data: dataOrigem,
                trip: tripNumber,
                motorista: motoristaLimpo || 'Não Identificado',
                placa: veiculo || 'Sem Placa',
                origem: origem || 'N/D',
                destino: destino || 'N/D',
                rota: `${origem || 'N/D'} ➔ ${destino || 'N/D'}`,
                motivo: motivoFinal,
                statusDestino: statusEtaDestino || 'N/D',
                statusViagem: statusAgrupado
            });
        }
    });

    statusBadge.textContent = `Carregado: ${fileName} (${dadosOcorrenciasCompletos.length} ocorrências)`;
    statusBadge.className = 'badge-status-online';

    popularFiltrosSelect();
    aplicarFiltros();
}

// ==========================================================================
// 3. POPULAR SELECTS DE FILTRO
// ==========================================================================

function popularFiltrosSelect() {
    const motivos = [...new Set(dadosOcorrenciasCompletos.map(d => d.motivo))].sort();

    filtroMotivo.innerHTML = '<option value="">Todos os Motivos</option>';
    motivos.forEach(motivo => {
        const opt = document.createElement('option');
        opt.value = motivo;
        opt.textContent = motivo;
        filtroMotivo.appendChild(opt);
    });
}

// ==========================================================================
// 4. APLICAÇÃO DE FILTROS E ATUALIZAÇÃO DA TELA
// ==========================================================================

filtroPesquisa.addEventListener('input', aplicarFiltros);
filtroMotivo.addEventListener('change', aplicarFiltros);
filtroStatus.addEventListener('change', aplicarFiltros);

btnLimparFiltros.addEventListener('click', () => {
    filtroPesquisa.value = '';
    filtroMotivo.value = '';
    filtroStatus.value = '';
    aplicarFiltros();
});

function aplicarFiltros() {
    const termo = filtroPesquisa.value.toLowerCase().trim();
    const motivoSel = filtroMotivo.value;
    const statusSel = filtroStatus.value;

    dadosFiltrados = dadosOcorrenciasCompletos.filter(item => {
        const matchesTermo = !termo ||
            item.motorista.toLowerCase().includes(termo) ||
            item.placa.toLowerCase().includes(termo) ||
            item.trip.toLowerCase().includes(termo) ||
            item.motivo.toLowerCase().includes(termo);

        const matchesMotivo = !motivoSel || item.motivo === motivoSel;
        const matchesStatus = !statusSel || item.statusDestino === statusSel;

        return matchesTermo && matchesMotivo && matchesStatus;
    });

    atualizarKPIs(dadosFiltrados);
    renderizarGraficos(dadosFiltrados);
    renderizarTabela(dadosFiltrados);
}

// ==========================================================================
// 5. ATUALIZAR INDICADORES (KPIS)
// ==========================================================================

function atualizarKPIs(dados) {
    kpiTotalOcorrencias.textContent = dados.length;

    const totalDelay = dados.filter(d => d.statusDestino === 'DELAY' || d.motivo.toLowerCase().includes('delay')).length;
    kpiTotalDelay.textContent = totalDelay;

    const totalParadas = dados.filter(d => d.motivo.toLowerCase().includes('parada') || d.motivo.toLowerCase().includes('abastecimento')).length;
    kpiTotalParadas.textContent = totalParadas;

    const motoristasUnicos = new Set(dados.map(d => d.motorista).filter(m => m !== 'Não Identificado'));
    kpiTotalMotoristas.textContent = motoristasUnicos.size;
}

// ==========================================================================
// 6. RENDERIZAR GRÁFICOS (CHART.JS)
// ==========================================================================

function renderizarGraficos(dados) {
    // 1. Contagem por Motivo
    const contagemMotivos = {};
    dados.forEach(d => {
        contagemMotivos[d.motivo] = (contagemMotivos[d.motivo] || 0) + 1;
    });

    const labelsMotivos = Object.keys(contagemMotivos);
    const dataMotivos = Object.values(contagemMotivos);

    if (chartMotivosInstance) chartMotivosInstance.destroy();
    const ctxMotivos = document.getElementById('chart-motivos').getContext('2d');
    chartMotivosInstance = new Chart(ctxMotivos, {
        type: 'doughnut',
        data: {
            labels: labelsMotivos,
            datasets: [{
                data: dataMotivos,
                backgroundColor: [
                    '#ef4444', '#f59e0b', '#3b82f6', '#10b981', '#8b5cf6',
                    '#ec4899', '#6366f1', '#14b8a6', '#f97316', '#64748b'
                ]
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

    // 2. Contagem por Motorista (Top 10)
    const contagemMotoristas = {};
    dados.forEach(d => {
        if (d.motorista && d.motorista !== 'Não Identificado') {
            contagemMotoristas[d.motorista] = (contagemMotoristas[d.motorista] || 0) + 1;
        }
    });

    const motoristasOrdenados = Object.entries(contagemMotoristas)
        .sort((a, b) => b[1] - a[1])
        .slice(0, 10);

    const labelsMotoristas = motoristasOrdenados.map(m => m[0]);
    const dataMotoristas = motoristasOrdenados.map(m => m[1]);

    if (chartMotoristasInstance) chartMotoristasInstance.destroy();
    const ctxMotoristas = document.getElementById('chart-motoristas').getContext('2d');
    chartMotoristasInstance = new Chart(ctxMotoristas, {
        type: 'bar',
        data: {
            labels: labelsMotoristas,
            datasets: [{
                label: 'Ocorrências Registradas',
                data: dataMotoristas,
                backgroundColor: '#1e1e5c',
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
// 7. RENDERIZAR TABELA
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

        // Formatação do badge de status
        let badgeClasse = 'badge-default';
        if (item.statusDestino === 'DELAY') badgeClasse = 'badge-delay';
        else if (item.statusDestino === 'EARLY') badgeClasse = 'badge-early';
        else if (item.statusDestino === 'ON TIME') badgeClasse = 'badge-ontime';

        tr.innerHTML = `
            <td style="font-weight: 500;">${item.data || '-'}</td>
            <td><code>${item.trip}</code></td>
            <td style="font-weight: 600; color: #1e1e5c;">${item.motorista}</td>
            <td><strong>${item.placa}</strong></td>
            <td>${item.rota}</td>
            <td style="color: #b91c1c; font-weight: 500;">${item.motivo}</td>
            <td><span class="badge-status ${badgeClasse}">${item.statusDestino}</span></td>
        `;

        tabelaCorpo.appendChild(tr);
    });
}

// ==========================================================================
// 8. EXPORTAÇÃO PARA EXCEL / CSV
// ==========================================================================

btnExportarCsv.addEventListener('click', () => {
    if (dadosFiltrados.length === 0) {
        alert('Não há dados filtrados para exportar.');
        return;
    }

    const ws = XLSX.utils.json_to_sheet(dadosFiltrados.map(d => ({
        'Data': d.data,
        'Nº Viagem (Trip)': d.trip,
        'Motorista': d.motorista,
        'Placas': d.placa,
        'Rota': d.rota,
        'Ocorrência / Motivo': d.motivo,
        'Status Destino': d.statusDestino,
        'Status Viagem': d.statusViagem
    })));

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Ocorrencias');
    XLSX.writeFile(wb, 'Relatorio_Ocorrencias_AJBorges.xlsx');
});
