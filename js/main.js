/* ═══════════════════════════════════════════════════════════
   CONFIGURAÇÃO — URLs das abas do Google Sheets (CSV)
   ═══════════════════════════════════════════════════════════ */
const CSV_URLS = [
    'https://docs.google.com/spreadsheets/d/e/2PACX-1vRy08WRcxagknb0ucDbTkvUyUW7hjqR2uAyHtaSqyZpIJIq8ejzTL-1F2ZC0M4spg8XUQIxBqNz38s_/pub?gid=352996843&single=true&output=csv',
    'https://docs.google.com/spreadsheets/d/e/2PACX-1vRy08WRcxagknb0ucDbTkvUyUW7hjqR2uAyHtaSqyZpIJIq8ejzTL-1F2ZC0M4spg8XUQIxBqNz38s_/pub?gid=1439413758&single=true&output=csv',
    'https://docs.google.com/spreadsheets/d/e/2PACX-1vRy08WRcxagknb0ucDbTkvUyUW7hjqR2uAyHtaSqyZpIJIq8ejzTL-1F2ZC0M4spg8XUQIxBqNz38s_/pub?gid=1819995490&single=true&output=csv',
    'https://docs.google.com/spreadsheets/d/e/2PACX-1vRy08WRcxagknb0ucDbTkvUyUW7hjqR2uAyHtaSqyZpIJIq8ejzTL-1F2ZC0M4spg8XUQIxBqNz38s_/pub?gid=2022041798&single=true&output=csv',
    'https://docs.google.com/spreadsheets/d/e/2PACX-1vRy08WRcxagknb0ucDbTkvUyUW7hjqR2uAyHtaSqyZpIJIq8ejzTL-1F2ZC0M4spg8XUQIxBqNz38s_/pub?gid=1117196734&single=true&output=csv',
];

/* ═══════════════════════════════════════════════════════════
   UTILITÁRIOS
   ═══════════════════════════════════════════════════════════ */
function normKey(s) {
    return String(s).toLowerCase()
        .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
        .replace(/[^a-z0-9]/g, '');
}

function parseDate(s) {
    if (!s) return null;
    s = s.trim();
    let m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
    if (m) return new Date(+m[3], +m[2] - 1, +m[1]);
    m = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
    if (m) return new Date(+m[1], +m[2] - 1, +m[3]);
    return null;
}

function toISO(d) {
    if (!d) return '';
    return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
}

function fmtDate(d) {
    if (!d) return '—';
    return `${String(d.getDate()).padStart(2, '0')}/${String(d.getMonth() + 1).padStart(2, '0')}/${d.getFullYear()}`;
}

function daysDiff(a, b) {
    if (!a || !b) return null;
    const diff = Math.round((b - a) / 86400000);
    return diff;
}

/* ═══════════════════════════════════════════════════════════
   MAPEAMENTO DE COLUNAS (com aliases)
   ═══════════════════════════════════════════════════════════ */
const COL = {
    projeto: ['projeto'],
    dataCal: ['datacalibracao', 'datadacalibracao', 'calibracao', 'dtcalibracao', 'dtcal'],
    dataAud: ['dataauditoria', 'datadaauditoria', 'dtauditoria', 'dtaud'],
    chassi: ['chassi', 'chassis'],
    auditor: ['auditor'],
    correta: ['auditoriaestacorreta', 'auditoriaestacoreta', 'estacorreta', 'correta', 'auditoriaestaokay', 'auditoriaestaok'],
    obs: ['observacao', 'observacoes', 'obs', 'observacao1'],
    gravidade: ['gravidadeimpacto', 'gravidade', 'impacto', 'impactogravidade', 'classificacaoimpacto'],
};

function findCol(hmap, key) {
    for (const alias of (COL[key] || [key])) {
        if (hmap[alias] !== undefined) return hmap[alias];
    }
    // fuzzy: check if any header contains the alias
    for (const alias of (COL[key] || [key])) {
        const found = Object.keys(hmap).find(h => h.includes(alias) || alias.includes(h));
        if (found !== undefined) return hmap[found];
    }
    return undefined;
}

function buildHmap(headers) {
    const m = {};
    headers.forEach((h, i) => { m[normKey(h)] = i; });
    return m;
}

function getCell(raw, hmap, key) {
    const idx = findCol(hmap, key);
    return idx !== undefined ? String(raw[idx] || '').trim() : '';
}

function normCorreta(s) {
    const n = s.toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '');
    if (n.startsWith('sim') || n === 's') return 'sim';
    if (n.startsWith('nao') || n.startsWith('nã') || n === 'n' || n === 'nao') return 'nao';
    return null;
}

function normGravidade(s) {
    const n = String(s || '').toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '').trim();
    if (!n) return null;
    if (n.includes('passivel') || n.includes('ajuste')) return 'passivel';
    if (n.includes('grave') || n.includes('prejuizo')) return 'grave';
    if (n.includes('omoda')) return 'omoda';
    return null;
}

/* ═══════════════════════════════════════════════════════════
   DATA LAYER
   ═══════════════════════════════════════════════════════════ */
let allRows = [];
let filteredRows = [];
let charts = {};
let selectedProjects = new Set();
let topObsLimit = 10;
let topObsGravidade = 'all';
let selectedAuditorEvolution = 'all';
let auditorDrillAuditor = null;
let auditorDrillGravidade = null;
const URL_PARAMS = new URLSearchParams(window.location.search);
const FORCED_PROJECT = (URL_PARAMS.get('projeto') || '').trim();

function normalizeRow(raw, hmap) {
    const corr = normCorreta(getCell(raw, hmap, 'correta'));
    if (!corr) return null;
    const projeto = getCell(raw, hmap, 'projeto');
    if (!projeto) return null;
    const dataCal = parseDate(getCell(raw, hmap, 'dataCal'));
    const dataAud = parseDate(getCell(raw, hmap, 'dataAud'));
    const chassi = getCell(raw, hmap, 'chassi');
    const auditor = getCell(raw, hmap, 'auditor');
    const obs = getCell(raw, hmap, 'obs');
    const gravidadeRaw = getCell(raw, hmap, 'gravidade');
    const gravidade = normGravidade(gravidadeRaw);
    return {
        projeto,
        dataCal,
        dataAud,
        chassi,
        auditor,
        correto: corr === 'sim',
        obs,
        gravidade,
        gravidadeRaw,
        sla: daysDiff(dataCal, dataAud),
        isoCal: dataCal ? toISO(dataCal) : '',
        isoAud: dataAud ? toISO(dataAud) : '',
    };
}

async function fetchOne(url) {
    const r = await fetch(url + '&_ts=' + Date.now(), { cache: 'no-store' });
    if (!r.ok) throw new Error('HTTP ' + r.status);
    return r.text();
}

function parseText(text) {
    const firstLine = text.split('\n')[0];
    const delim = (firstLine.match(/;/g) || []).length > (firstLine.match(/,/g) || []).length ? ';' : ',';
    return Papa.parse(text, { delimiter: delim, skipEmptyLines: true, header: false }).data;
}

async function loadAllData() {
    // Show loading
    document.getElementById('dashboard').style.display = 'none';
    document.getElementById('stateError').style.display = 'none';
    document.getElementById('stateLoading').style.display = 'flex';

    // Destroy old charts
    Object.values(charts).forEach(c => { try { c.destroy(); } catch (e) { } });
    charts = {};
    allRows = [];

    const erros = [];
    for (const url of CSV_URLS) {
        try {
            const text = await fetchOne(url);
            const data = parseText(text);
            if (data.length < 2) { erros.push('Aba vazia'); continue; }
            const hmap = buildHmap(data[0]);
            for (let i = 1; i < data.length; i++) {
                const row = normalizeRow(data[i], hmap);
                if (row) allRows.push(row);
            }
        } catch (e) {
            erros.push(e.message);
            console.warn('Erro ao carregar aba:', e);
        }
    }

    if (allRows.length === 0) {
        document.getElementById('stateLoading').style.display = 'none';
        document.getElementById('errTitle').textContent = 'Nenhum dado encontrado';
        document.getElementById('errDesc').textContent = erros.length
            ? 'Erros: ' + erros.join(' | ')
            : 'Verifique se as planilhas estão publicadas como CSV publicamente.';
        document.getElementById('stateError').style.display = 'flex';
        return;
    }

    try {
        // Timestamp
        const now = new Date();
        document.getElementById('tsEl').textContent =
            fmtDate(now) + ' ' + now.toLocaleTimeString('pt-BR', { hour: '2-digit', minute: '2-digit' });

        populateFilters();
        renderProjectPagesNav();
        applyFilters();

        document.getElementById('stateLoading').style.display = 'none';
        document.getElementById('dashboard').style.display = 'block';
    } catch (e) {
        console.error('Erro ao renderizar dashboard:', e);
        document.getElementById('stateLoading').style.display = 'none';
        document.getElementById('errTitle').textContent = 'Erro ao montar o dashboard';
        document.getElementById('errDesc').textContent = e?.message || 'Falha inesperada ao renderizar os gráficos.';
        document.getElementById('stateError').style.display = 'flex';
    }
}

function renderProjectPagesNav() {
    const wrap = document.getElementById('projectPagesWrap');
    const nav = document.getElementById('projectPagesNav');
    if (!wrap || !nav) return;

    const projects = [...new Set(allRows.map(r => r.projeto).filter(Boolean))].sort();
    if (!projects.length) {
        wrap.style.display = 'none';
        return;
    }

    const targetPage = 'projeto.html';
    const geralLink = `<a class="btn btn-ghost project-page-link ${!FORCED_PROJECT ? 'active' : ''}" href="index.html">Geral</a>`;
    const projectLinks = projects.map(p => {
        const active = FORCED_PROJECT && p === FORCED_PROJECT;
        const href = `${targetPage}?projeto=${encodeURIComponent(p)}`;
        return `<a class="btn btn-ghost project-page-link ${active ? 'active' : ''}" href="${href}">${p}</a>`;
    }).join('');

    nav.innerHTML = geralLink + projectLinks;

    wrap.style.display = 'block';

    if (FORCED_PROJECT) {
        document.title = `QA Intelligence - ${FORCED_PROJECT}`;
        const sub = document.querySelector('.header-sub');
        if (sub) sub.textContent = `Painel de Indicadores de Qualidade · Projeto: ${FORCED_PROJECT}`;
    }
}

/* ═══════════════════════════════════════════════════════════
   FILTERS
   ═══════════════════════════════════════════════════════════ */
function populateFilters() {
    // Projects multiselect
    const projects = [...new Set(allRows.map(r => r.projeto).filter(Boolean))].sort();
    selectedProjects = new Set(projects);
    const dd = document.getElementById('msDD');
    dd.innerHTML = projects.map(p =>
        `<label class="ms-opt"><input type="checkbox" value="${p}" checked onchange="onProjChange()"> ${p}</label>`
    ).join('');
    updateMSText();

    if (FORCED_PROJECT && projects.includes(FORCED_PROJECT)) {
        const checks = document.querySelectorAll('#msDD input[type=checkbox]');
        checks.forEach(c => {
            c.checked = c.value === FORCED_PROJECT;
            c.disabled = true;
        });
        selectedProjects = new Set([FORCED_PROJECT]);
        updateMSText();

        const msDisplay = document.getElementById('msDisplay');
        if (msDisplay) {
            msDisplay.style.pointerEvents = 'none';
            msDisplay.style.opacity = '0.7';
        }
    }

    // Auditors
    const auditors = [...new Set(allRows.map(r => r.auditor).filter(Boolean))].sort();
    const sel = document.getElementById('fAuditor');
    sel.innerHTML = '<option value="">Todos</option>' +
        auditors.map(a => `<option value="${a}">${a}</option>`).join('');

    // Gravidade / Impacto
    const gravSel = document.getElementById('fGravidade');
    if (gravSel) {
        gravSel.innerHTML = [
            '<option value="">Todos</option>',
            '<option value="passivel">Passível de Ajuste</option>',
            '<option value="grave">Grave (Prejuízo Financeiro)</option>',
            '<option value="omoda">AVALIAÇÃO OMODA</option>',
            '<option value="naoClassificado">Não Classificado</option>'
        ].join('');
    }

    // Date range
    const dates = allRows.map(r => r.isoAud).filter(Boolean).sort();
    if (dates.length) {
        document.getElementById('fDateS').value = dates[0];
        document.getElementById('fDateE').value = dates[dates.length - 1];
    }
}

let msOpen = false;
function toggleMS() {
    msOpen = !msOpen;
    document.getElementById('msDD').classList.toggle('open', msOpen);
    document.getElementById('msDisplay').classList.toggle('open', msOpen);
    if (msOpen) {
        setTimeout(() => {
            const close = (e) => {
                if (!document.getElementById('msDD').contains(e.target) &&
                    !document.getElementById('msDisplay').contains(e.target)) {
                    msOpen = false;
                    document.getElementById('msDD').classList.remove('open');
                    document.getElementById('msDisplay').classList.remove('open');
                    document.removeEventListener('click', close);
                }
            };
            document.addEventListener('click', close);
        }, 0);
    }
}

function onProjChange() {
    const checks = document.querySelectorAll('#msDD input[type=checkbox]');
    selectedProjects = new Set([...checks].filter(c => c.checked).map(c => c.value));
    updateMSText();
    applyFilters();
}

function updateMSText() {
    const all = document.querySelectorAll('#msDD input[type=checkbox]').length;
    const sel = selectedProjects.size;
    document.getElementById('msText').textContent =
        sel === 0 ? 'Nenhum' : sel === all ? 'Todos' : `${sel} selecionado${sel > 1 ? 's' : ''}`;
}

function resetFilters() {
    document.getElementById('fChassi').value = '';
    document.getElementById('fAuditor').value = '';
    document.getElementById('fGravidade').value = '';
    const dates = allRows.map(r => r.isoAud).filter(Boolean).sort();
    if (dates.length) {
        document.getElementById('fDateS').value = dates[0];
        document.getElementById('fDateE').value = dates[dates.length - 1];
    } else {
        document.getElementById('fDateS').value = '';
        document.getElementById('fDateE').value = '';
    }
    const checks = document.querySelectorAll('#msDD input[type=checkbox]');
    if (FORCED_PROJECT) {
        checks.forEach(c => {
            c.checked = c.value === FORCED_PROJECT;
            c.disabled = true;
        });
        selectedProjects = new Set([FORCED_PROJECT]);
    } else {
        checks.forEach(c => c.checked = true);
        selectedProjects = new Set([...checks].map(c => c.value));
    }
    updateMSText();
    applyFilters();
}

function applyFilters() {
    const chassi = document.getElementById('fChassi').value.trim().toLowerCase();
    const dateCal = document.getElementById('fDateCal').value;
    const dateS = document.getElementById('fDateS').value;
    const dateE = document.getElementById('fDateE').value;
    const auditor = document.getElementById('fAuditor').value;
    const gravidade = document.getElementById('fGravidade').value;

    filteredRows = allRows.filter(r => {
        if (FORCED_PROJECT && r.projeto !== FORCED_PROJECT) return false;
        if (chassi && !r.chassi.toLowerCase().includes(chassi)) return false;
        if (dateCal && r.isoCal && r.isoCal !== dateCal) return false;
        if (dateS && r.isoAud && r.isoAud < dateS) return false;
        if (dateE && r.isoAud && r.isoAud > dateE) return false;
        if (selectedProjects.size > 0 && !selectedProjects.has(r.projeto)) return false;
        if (auditor && r.auditor !== auditor) return false;
        if (gravidade) {
            if (gravidade === 'naoClassificado' && r.gravidade) return false;
            if (gravidade !== 'naoClassificado' && r.gravidade !== gravidade) return false;
        }
        return true;
    });

    auditorDrillAuditor = null;
    auditorDrillGravidade = null;
    renderKPIs();
    renderCharts();
}

/* ═══════════════════════════════════════════════════════════
    KPIs
   ═══════════════════════════════════════════════════════════ */
function renderKPIs() {
    const total = filteredRows.length;
    const sim = filteredRows.filter(r => r.correto).length;
    const nao = total - sim;
    const pct = total > 0 ? sim / total * 100 : 0;

    // Volume
    document.getElementById('kTotal').textContent = total.toLocaleString('pt-BR');

    // Qualidade
    document.getElementById('kQual').textContent = pct.toFixed(1) + '%';
    const above = pct >= 95;
    document.getElementById('kQualAccent').className = 'kpi-accent ' + (above ? 'accent-green' : 'accent-red');
    document.getElementById('kQualIcon').className = 'kpi-icon ' + (above ? 'icon-green' : 'icon-red');
    document.getElementById('kQualIcon').textContent = above ? '✅' : '⚠️';
    const chip = document.getElementById('kQualChip');
    chip.className = 'kpi-chip ' + (above ? 'chip-green' : 'chip-red');
    chip.textContent = above ? '✓ Acima da meta (95%)' : `▼ ${(95 - pct).toFixed(1)}% abaixo da meta`;

    // Inconsistência
    document.getElementById('kNao').textContent = nao.toLocaleString('pt-BR');

    // Volume de Conformes
    document.getElementById('kConformes').textContent = sim.toLocaleString('pt-BR');

    // SLA médio
    const slas = filteredRows.map(r => r.sla).filter(v => v !== null && v >= 0);
    const slaAvg = slas.length > 0 ? (slas.reduce((a, b) => a + b, 0) / slas.length).toFixed(1) : '—';
    document.getElementById('kSla').textContent = slaAvg;

    // Auditor destaque
    const aMap = {};
    filteredRows.forEach(r => { if (r.auditor) aMap[r.auditor] = (aMap[r.auditor] || 0) + 1; });
    const top = Object.entries(aMap).sort((a, b) => b[1] - a[1])[0];
    if (top) {
        document.getElementById('kAud').textContent = top[0].split(' ')[0];
        document.getElementById('kAudChip').textContent = top[1] + ' auditorias';
    } else {
        document.getElementById('kAud').textContent = '—';
        document.getElementById('kAudChip').textContent = '—';
    }
}

/* ═══════════════════════════════════════════════════════════
   CHARTS — helpers
   ═══════════════════════════════════════════════════════════ */
const C = {
    sim: '#059669',
    simBg: 'rgba(5,150,105,0.95)',
    nao: '#DC2626',
    naoBg: 'rgba(220,38,38,0.95)',
    blue: '#0369A1',
    blueBg: 'rgba(3,105,161,0.95)',
    ref: '#F59E0B',
    grid: '#E8EEF5',
    txt: '#1e293b',
    font: 'DM Sans',
};

const CHART_FONT = { family: C.font, size: 12, weight: '500' };

// Register DataLabels plugin
Chart.register(ChartDataLabels);

function destroyC(id) {
    if (charts[id]) { try { charts[id].destroy(); } catch (e) { } delete charts[id]; }
}

function renderCharts() {
    renderGauge();
    renderTrend();
    renderProject();
    renderGravidadeImpacto();
    if (FORCED_PROJECT) {
        renderGravidadePorAuditor();
    } else {
        renderGravidadePorProjeto();
    }
    renderGravidadePorObservacao();
    renderPareto();
    renderAuditor();
    renderAuditorEvolution();
    renderProjectRanking();
}

function onTopObsLimitChange() {
    const sel = document.getElementById('topObsLimit');
    const v = parseInt(sel?.value, 10);
    topObsLimit = [5, 10, 15].includes(v) ? v : 10;
    renderGravidadePorObservacao();
}

function onTopObsGravidadeChange() {
    const sel = document.getElementById('topObsGravidade');
    const v = sel?.value || 'all';
    topObsGravidade = ['all', 'passivel', 'grave', 'omoda', 'naoClassificado'].includes(v) ? v : 'all';
    renderGravidadePorObservacao();
}

/* ── GAUGE ───────────────────────────────────────────────── */
function renderGauge() {
    destroyC('gauge');
    const total = filteredRows.length;
    const sim = filteredRows.filter(r => r.correto).length;
    const pct = total > 0 ? sim / total * 100 : 0;
    const above = pct >= 95;
    const color = above ? C.sim : C.nao;

    document.getElementById('gaugePct').textContent = pct.toFixed(1) + '%';
    document.getElementById('gaugePct').style.color = color;

    const badge = document.getElementById('gaugeBadge');
    badge.textContent = above ? '✓ META ATINGIDA' : '⚠ ABAIXO DA META';
    badge.style.background = above ? '#D1FAE5' : '#FEE2E2';
    badge.style.color = above ? '#047857' : '#B91C1C';

    const ctx = document.getElementById('gaugeChart').getContext('2d');
    const val = Math.min(Math.max(pct, 0), 100);
    const remaining = 100 - val;

    charts.gauge = new Chart(ctx, {
        type: 'doughnut',
        data: {
            datasets: [{
                data: [val, remaining],
                backgroundColor: [color, '#EFF3F8'],
                borderWidth: 0,
                borderRadius: [6, 0],
                hoverBorderWidth: 0,
            }]
        },
        options: {
            rotation: -90,
            circumference: 180,
            cutout: '70%',
            responsive: true,
            maintainAspectRatio: true,
            animation: { duration: 900, easing: 'easeInOutQuart' },
            plugins: {
                legend: { display: false },
                datalabels: {
                    display: false
                },
                tooltip: {
                    callbacks: {
                        label: (ctx) => ctx.dataIndex === 0
                            ? `Qualidade: ${val.toFixed(1)}%`
                            : `Restante: ${remaining.toFixed(1)}%`
                    }
                }
            }
        }
    });
}

/* ── TREND ───────────────────────────────────────────────── */
function renderTrend() {
    destroyC('trend');

    // Verificar se há filtro de data específico (intervalo <= 30 dias)
    const dateS = document.getElementById('fDateS').value;
    const dateE = document.getElementById('fDateE').value;
    const hasSpecificDateFilter = dateS && dateE;

    let isDailyView = false;
    if (hasSpecificDateFilter) {
        const startDate = new Date(dateS);
        const endDate = new Date(dateE);
        const diffTime = Math.abs(endDate - startDate);
        const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));
        isDailyView = diffDays <= 30; // Se intervalo <= 30 dias, mostrar visão diária
    }

    let byPeriod = {};
    let periodLabels = [];

    if (isDailyView) {
        // Visão diária quando filtro específico
        filteredRows.forEach(r => {
            if (!r.isoAud) return;
            if (!byPeriod[r.isoAud]) byPeriod[r.isoAud] = { sim: 0, total: 0 };
            byPeriod[r.isoAud].total++;
            if (r.correto) byPeriod[r.isoAud].sim++;
        });
        periodLabels = Object.keys(byPeriod).sort();
    } else {
        // Visão mensal por campanha (padrão)
        filteredRows.forEach(r => {
            if (!r.isoAud) return;
            const [year, month] = r.isoAud.split('-');
            const periodKey = `${year}-${month}`;

            if (!byPeriod[periodKey]) byPeriod[periodKey] = { sim: 0, total: 0 };
            byPeriod[periodKey].total++;
            if (r.correto) byPeriod[periodKey].sim++;
        });
        periodLabels = Object.keys(byPeriod).sort();
    }

    const pcts = periodLabels.map(p => byPeriod[p].total > 0
        ? +(byPeriod[p].sim / byPeriod[p].total * 100).toFixed(1) : 0);
    const refs = periodLabels.map(() => 95);

    let fmtL;
    if (isDailyView) {
        fmtL = periodLabels.map(d => { const [y, m, dd] = d.split('-'); return `${dd}/${m}`; });
        document.getElementById('trendBadge').textContent = periodLabels.length + ' dias';
    } else {
        fmtL = periodLabels.map(p => {
            const [y, m] = p.split('-');
            const monthNames = ['Jan', 'Fev', 'Mar', 'Abr', 'Mai', 'Jun', 'Jul', 'Ago', 'Set', 'Out', 'Nov', 'Dez'];
            return `${monthNames[parseInt(m) - 1]}/${y.slice(2)}`;
        });
        document.getElementById('trendBadge').textContent = periodLabels.length + ' meses';
    }

    const ctx = document.getElementById('trendChart').getContext('2d');
    charts.trend = new Chart(ctx, {
        type: 'line',
        data: {
            labels: fmtL,
            datasets: [
                {
                    label: '% Qualidade',
                    data: pcts,
                    borderColor: C.blue,
                    backgroundColor: 'rgba(3,105,161,0.15)',
                    fill: true,
                    tension: 0.42,
                    pointRadius: periodLabels.length > 60 ? 2 : 4,
                    pointHoverRadius: 7,
                    pointBackgroundColor: (ctx) => ctx.parsed.y >= 95 ? '#10B981' : C.blue,
                    pointBorderColor: (ctx) => ctx.parsed.y >= 95 ? '#10B981' : C.blue,
                    borderWidth: 2.5,
                    order: 1,
                },
                {
                    label: 'Meta (95%)',
                    data: refs,
                    borderColor: C.ref,
                    borderDash: [7, 4],
                    borderWidth: 2,
                    pointRadius: 0,
                    fill: false,
                    tension: 0,
                    order: 2,
                }
            ]
        },
        options: {
            responsive: true,
            maintainAspectRatio: true,
            layout: {
                padding: {
                    top: 40,
                    bottom: 20,
                    left: 20,
                    right: 20
                }
            },
            interaction: { mode: 'index', intersect: false },
            plugins: {
                legend: { 
                    display: true,
                    position: 'bottom',
                    align: 'center'
                },
                datalabels: {
                    display: true,
                    clip: false,
                    font: { ...CHART_FONT, weight: '600', size: 11 },
                    color: '#000',
                    formatter: (value, ctx) => {
                        if (ctx.datasetIndex === 1) return '';
                        return value + '%';
                    },
                    offset: 8,
                    anchor: 'end',
                    align: 'top'
                },
                tooltip: {
                    callbacks: {
                        label: ctx => `${ctx.dataset.label}: ${ctx.parsed.y}%`,
                        afterLabel: ctx => {
                            if (ctx.datasetIndex === 0 && periodLabels[ctx.dataIndex]) {
                                const d = byPeriod[periodLabels[ctx.dataIndex]];
                                return `  Sim: ${d.sim}  |  Total: ${d.total}`;
                            }
                            return '';
                        }
                    }
                }
            },
            scales: {
                x: {
                    grid: { color: C.grid },
                    ticks: { font: CHART_FONT, color: C.txt, maxRotation: 45, maxTicksLimit: 20, padding: 5 }
                },
                y: {
                    min: 0, max: 100,
                    grid: { color: C.grid },
                    ticks: { font: CHART_FONT, color: C.txt, callback: v => v + '%', padding: 5 }
                }
            }
        }
    });

    // Legend is now at the bottom of the chart
}

/* ── PROJECT RANKING ────────────────────────────────────── */
function renderProjectRanking() {
    destroyC('projectRanking');

    // Agrupar por projeto e calcular % de qualidade
    let projectData = {};
    filteredRows.forEach(r => {
        if (!r.projeto) return;
        if (!projectData[r.projeto]) projectData[r.projeto] = { sim: 0, total: 0 };
        projectData[r.projeto].total++;
        if (r.correto) projectData[r.projeto].sim++;
    });

    // Calcular percentuais e ordenar do melhor para o pior
    let rankings = Object.entries(projectData).map(([projeto, data]) => ({
        projeto,
        percentage: data.total > 0 ? +(data.sim / data.total * 100).toFixed(1) : 0,
        sim: data.sim,
        total: data.total
    })).sort((a, b) => b.percentage - a.percentage);

    if (rankings.length === 0) {
        return;
    }

    const labels = rankings.map(r => r.projeto);
    const data = rankings.map(r => r.percentage);
    const colors = rankings.map(r => {
        if (r.percentage >= 95) return '#10B981'; // Verde - dentro da meta
        if (r.percentage >= 80) return '#F59E0B'; // Amarelo - atenção
        return '#DC2626'; // Vermelho - crítico
    });

    const ctx = document.getElementById('projectRankingChart').getContext('2d');
    charts.projectRanking = new Chart(ctx, {
        type: 'bar',
        data: {
            labels,
            datasets: [{
                label: '% Qualidade',
                data,
                backgroundColor: colors,
                borderColor: colors,
                borderWidth: 1,
                borderRadius: 5
            }]
        },
        options: {
            indexAxis: 'y',
            responsive: true,
            maintainAspectRatio: true,
            aspectRatio: 2.5,
            layout: {
                padding: { top: 20, bottom: 20, left: 20, right: 20 }
            },
            plugins: {
                legend: {
                    display: true,
                    position: 'top',
                    labels: { font: CHART_FONT, color: C.txt, boxWidth: 12, padding: 10 }
                },
                datalabels: {
                    display: true,
                    clip: false,
                    font: { ...CHART_FONT, weight: '600', size: 11 },
                    color: '#000',
                    formatter: (value) => value + '%',
                    offset: 6,
                    anchor: 'end',
                    align: 'right'
                },
                tooltip: {
                    callbacks: {
                        label: ctx => {
                            const ranking = rankings[ctx.dataIndex];
                            return `${ranking.sim}/${ranking.total} (${ctx.parsed.x}%)`;
                        }
                    }
                }
            },
            scales: {
                x: {
                    beginAtZero: true,
                    max: 100,
                    grid: { color: C.grid },
                    ticks: {
                        font: CHART_FONT,
                        color: C.txt,
                        callback: value => value + '%'
                    }
                },
                y: {
                    grid: { display: false },
                    ticks: { font: CHART_FONT, color: C.txt }
                }
            }
        }
    });
}

/* ── PROJECT STACKED ─────────────────────────────────────── */
function renderProject() {
    destroyC('project');
    const projs = [...new Set(filteredRows.map(r => r.projeto).filter(Boolean))].sort();
    const simD = projs.map(p => filteredRows.filter(r => r.projeto === p && r.correto).length);
    const naoD = projs.map(p => filteredRows.filter(r => r.projeto === p && !r.correto).length);

    const ctx = document.getElementById('projectChart').getContext('2d');
    charts.project = new Chart(ctx, {
        type: 'bar',
        data: {
            labels: projs,
            datasets: [
                { label: 'Conformes (Sim)', data: simD, backgroundColor: C.simBg, borderColor: C.sim, borderWidth: 1, borderRadius: 5, borderSkipped: false },
                { label: 'Não Conformes (Não)', data: naoD, backgroundColor: C.naoBg, borderColor: C.nao, borderWidth: 1, borderRadius: 5, borderSkipped: false },
            ]
        },
        options: {
            responsive: true,
            maintainAspectRatio: true,
            interaction: { mode: 'index' },
            plugins: {
                legend: { labels: { font: CHART_FONT, color: C.txt, boxWidth: 12, padding: 14 } },
                datalabels: {
                    display: true,
                    font: { ...CHART_FONT, weight: '600', size: 11 },
                    color: '#000',
                    formatter: (value) => value,
                    offset: 4,
                    anchor: 'center',
                    align: 'center'
                },
                tooltip: {
                    callbacks: {
                        label: ctx => {
                            const tot = simD[ctx.dataIndex] + naoD[ctx.dataIndex];
                            const p = tot > 0 ? (ctx.parsed.y / tot * 100).toFixed(1) : 0;
                            return ` ${ctx.dataset.label}: ${ctx.parsed.y}  (${p}%)`;
                        }
                    }
                }
            },
            scales: {
                x: { stacked: true, grid: { display: false }, ticks: { font: CHART_FONT, color: C.txt } },
                y: { stacked: true, grid: { color: C.grid }, ticks: { font: CHART_FONT, color: C.txt } }
            }
        }
    });
}

/* ── GRAVIDADE / IMPACTO ────────────────────────────────── */
function renderGravidadeImpacto() {
    destroyC('gravidadeImpacto');

    const naoConformes = filteredRows.filter(r => !r.correto);
    const passivel = naoConformes.filter(r => r.gravidade === 'passivel').length;
    const grave = naoConformes.filter(r => r.gravidade === 'grave').length;
    const omoda = naoConformes.filter(r => r.gravidade === 'omoda').length;
    const naoClassificado = naoConformes.filter(r => !r.gravidade).length;

    const labels = ['Passível de Ajuste', 'Grave (Prejuízo Financeiro)', 'AVALIAÇÃO OMODA', 'Não Classificado'];
    const data = [passivel, grave, omoda, naoClassificado];
    const colors = ['#F59E0B', '#DC2626', '#0EA5E9', '#94A3B8'];
    const total = data.reduce((acc, v) => acc + v, 0);

    const ctx = document.getElementById('gravidadeImpactoChart').getContext('2d');
    charts.gravidadeImpacto = new Chart(ctx, {
        type: 'doughnut',
        data: {
            labels,
            datasets: [{
                label: 'Quantidade',
                data,
                backgroundColor: colors,
                borderColor: colors,
                borderWidth: 1,
                hoverOffset: 8,
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: true,
            cutout: '58%',
            plugins: {
                legend: {
                    display: true,
                    position: 'bottom',
                    labels: { font: CHART_FONT, color: C.txt, boxWidth: 12, padding: 14 }
                },
                datalabels: {
                    display: true,
                    font: { ...CHART_FONT, weight: '700', size: 12 },
                    color: '#000',
                    formatter: (value) => {
                        if (!value || !total) return '';
                        const pct = (value / total * 100).toFixed(1);
                        return `${value} (${pct}%)`;
                    }
                },
                tooltip: {
                    callbacks: {
                        label: (ctx) => {
                            const value = ctx.parsed;
                            const pct = total > 0 ? (value / total * 100).toFixed(1) : '0.0';
                            return ` ${ctx.label}: ${value} (${pct}%)`;
                        },
                        afterLabel: () => ` Total de não conformes: ${total}`
                    }
                }
            }
        }
    });
}

/* ── GRAVIDADE / IMPACTO POR PROJETO ───────────────────── */
function renderGravidadePorProjeto() {
    destroyC('gravidadeProjeto');

    const ncRows = filteredRows.filter(r => !r.correto && r.projeto);
    const map = {};
    ncRows.forEach(r => {
        if (!map[r.projeto]) {
            map[r.projeto] = { passivel: 0, grave: 0, omoda: 0, naoClassificado: 0 };
        }
        if (r.gravidade === 'passivel') map[r.projeto].passivel++;
        else if (r.gravidade === 'grave') map[r.projeto].grave++;
        else if (r.gravidade === 'omoda') map[r.projeto].omoda++;
        else map[r.projeto].naoClassificado++;
    });

    const projetos = Object.keys(map).sort();
    const passivelData = projetos.map(p => map[p].passivel);
    const graveData = projetos.map(p => map[p].grave);
    const omodaData = projetos.map(p => map[p].omoda);
    const naoClassifData = projetos.map(p => map[p].naoClassificado);

    const ctx = document.getElementById('gravidadeProjetoChart').getContext('2d');
    charts.gravidadeProjeto = new Chart(ctx, {
        type: 'bar',
        data: {
            labels: projetos,
            datasets: [
                {
                    label: 'Passível de Ajuste',
                    data: passivelData,
                    backgroundColor: '#F59E0B',
                    borderColor: '#D97706',
                    borderWidth: 1,
                    borderRadius: 4,
                    borderSkipped: false,
                },
                {
                    label: 'Grave (Prejuízo Financeiro)',
                    data: graveData,
                    backgroundColor: '#DC2626',
                    borderColor: '#B91C1C',
                    borderWidth: 1,
                    borderRadius: 4,
                    borderSkipped: false,
                },
                {
                    label: 'AVALIAÇÃO OMODA',
                    data: omodaData,
                    backgroundColor: '#0EA5E9',
                    borderColor: '#0284C7',
                    borderWidth: 1,
                    borderRadius: 4,
                    borderSkipped: false,
                },
                {
                    label: 'Não Classificado',
                    data: naoClassifData,
                    backgroundColor: '#94A3B8',
                    borderColor: '#64748B',
                    borderWidth: 1,
                    borderRadius: 4,
                    borderSkipped: false,
                }
            ]
        },
        options: {
            responsive: true,
            maintainAspectRatio: true,
            interaction: { mode: 'index', intersect: false },
            plugins: {
                legend: {
                    labels: { font: CHART_FONT, color: C.txt, boxWidth: 12, padding: 14 }
                },
                datalabels: {
                    display: true,
                    font: { ...CHART_FONT, weight: '700', size: 12 },
                    color: '#000',
                    formatter: (value) => value > 0 ? value : '',
                    anchor: 'center',
                    align: 'center',
                    clamp: true
                },
                tooltip: {
                    callbacks: {
                        label: (ctx) => ` ${ctx.dataset.label}: ${ctx.parsed.y}`,
                        afterBody: (items) => {
                            if (!items || !items.length) return '';
                            const idx = items[0].dataIndex;
                            const total = passivelData[idx] + graveData[idx] + omodaData[idx] + naoClassifData[idx];
                            return `Total no projeto: ${total}`;
                        }
                    }
                }
            },
            scales: {
                x: {
                    stacked: true,
                    grid: { display: false },
                    ticks: { font: CHART_FONT, color: C.txt }
                },
                y: {
                    stacked: true,
                    beginAtZero: true,
                    grid: { color: C.grid },
                    ticks: { font: CHART_FONT, color: C.txt, precision: 0 }
                }
            }
        }
    });
}

/* ── GRAVIDADE / IMPACTO POR AUDITOR ───────────────────── */
function clearAuditorDrill() {
    auditorDrillAuditor = null;
    auditorDrillGravidade = null;
    renderGravidadePorAuditor();
    renderGravidadePorObservacao();
}

function renderGravidadePorAuditor() {
    destroyC('gravidadeAuditor');

    const ncRows = filteredRows.filter(r => !r.correto && r.auditor);
    const map = {};
    ncRows.forEach(r => {
        if (!map[r.auditor]) {
            map[r.auditor] = { passivel: 0, grave: 0, omoda: 0, naoClassificado: 0 };
        }
        if (r.gravidade === 'passivel') map[r.auditor].passivel++;
        else if (r.gravidade === 'grave') map[r.auditor].grave++;
        else if (r.gravidade === 'omoda') map[r.auditor].omoda++;
        else map[r.auditor].naoClassificado++;
    });

    const auditores = Object.keys(map).sort();
    const passivelData = auditores.map(a => map[a].passivel);
    const graveData = auditores.map(a => map[a].grave);
    const omodaData = auditores.map(a => map[a].omoda);
    const naoClassifData = auditores.map(a => map[a].naoClassificado);

    // Cores com destaque: esmaece não selecionados quando drill ativo
    function barBg(base, gravKey) {
        if (!auditorDrillAuditor) return auditores.map(() => base);
        return auditores.map(a =>
            (a === auditorDrillAuditor && auditorDrillGravidade === gravKey) ? base : base + '33'
        );
    }

    const el = document.getElementById('gravidadeAuditorChart');
    if (!el) return;
    const ctx = el.getContext('2d');
    if (!ctx) return;
    charts.gravidadeAuditor = new Chart(ctx, {
        type: 'bar',
        data: {
            labels: auditores,
            datasets: [
                {
                    label: 'Passível de Ajuste',
                    data: passivelData,
                    backgroundColor: barBg('#F59E0B', 'passivel'),
                    borderColor: barBg('#D97706', 'passivel'),
                    borderWidth: 1,
                    borderRadius: 4,
                    borderSkipped: false,
                },
                {
                    label: 'Grave (Prejuízo Financeiro)',
                    data: graveData,
                    backgroundColor: barBg('#DC2626', 'grave'),
                    borderColor: barBg('#B91C1C', 'grave'),
                    borderWidth: 1,
                    borderRadius: 4,
                    borderSkipped: false,
                },
                {
                    label: 'AVALIAÇÃO OMODA',
                    data: omodaData,
                    backgroundColor: barBg('#0EA5E9', 'omoda'),
                    borderColor: barBg('#0284C7', 'omoda'),
                    borderWidth: 1,
                    borderRadius: 4,
                    borderSkipped: false,
                },
                {
                    label: 'Não Classificado',
                    data: naoClassifData,
                    backgroundColor: barBg('#94A3B8', 'naoClassificado'),
                    borderColor: barBg('#64748B', 'naoClassificado'),
                    borderWidth: 1,
                    borderRadius: 4,
                    borderSkipped: false,
                }
            ]
        },
        options: {
            responsive: true,
            maintainAspectRatio: true,
            interaction: { mode: 'index', intersect: false },
            onClick: (event) => {
                const clickedElements = charts.gravidadeAuditor
                    ? charts.gravidadeAuditor.getElementsAtEventForMode(event, 'nearest', { intersect: true }, true)
                    : [];

                if (!clickedElements.length) {
                    if (auditorDrillAuditor) clearAuditorDrill();
                    return;
                }
                const gravKeyMap = ['passivel', 'grave', 'omoda', 'naoClassificado'];
                const clickedAuditor = auditores[clickedElements[0].index];
                const clickedGrav = gravKeyMap[clickedElements[0].datasetIndex];
                if (auditorDrillAuditor === clickedAuditor && auditorDrillGravidade === clickedGrav) {
                    auditorDrillAuditor = null;
                    auditorDrillGravidade = null;
                } else {
                    auditorDrillAuditor = clickedAuditor;
                    auditorDrillGravidade = clickedGrav;
                }
                renderGravidadePorAuditor();
                renderGravidadePorObservacao();
            },
            onHover: (event) => {
                const hoveredElements = charts.gravidadeAuditor
                    ? charts.gravidadeAuditor.getElementsAtEventForMode(event, 'nearest', { intersect: true }, true)
                    : [];
                if (event.native) event.native.target.style.cursor = hoveredElements.length ? 'pointer' : 'default';
            },
            plugins: {
                legend: {
                    labels: { font: CHART_FONT, color: C.txt, boxWidth: 12, padding: 14 }
                },
                datalabels: {
                    display: true,
                    font: { ...CHART_FONT, weight: '700', size: 12 },
                    color: '#000',
                    formatter: (value) => value > 0 ? value : '',
                    anchor: 'center',
                    align: 'center',
                    clamp: true
                },
                tooltip: {
                    callbacks: {
                        label: (ctx) => ` ${ctx.dataset.label}: ${ctx.parsed.y}`,
                        afterBody: (items) => {
                            if (!items || !items.length) return '';
                            const idx = items[0].dataIndex;
                            const total = passivelData[idx] + graveData[idx] + omodaData[idx] + naoClassifData[idx];
                            return `Total do auditor: ${total} — clique para filtrar observações`;
                        }
                    }
                }
            },
            scales: {
                x: {
                    stacked: true,
                    grid: { display: false },
                    ticks: { font: CHART_FONT, color: C.txt }
                },
                y: {
                    stacked: true,
                    beginAtZero: true,
                    grid: { color: C.grid },
                    ticks: { font: CHART_FONT, color: C.txt, precision: 0 }
                }
            }
        }
    });
}

/* ── GRAVIDADE / IMPACTO POR OBSERVAÇÃO ────────────────── */
function renderGravidadePorObservacao() {
    destroyC('gravidadeObs');

    // Atualizar badge de drill-down
    const badge = document.getElementById('auditorDrillBadge');
    if (badge) {
        if (auditorDrillAuditor) {
            const gravLabels = { passivel: 'Passível de Ajuste', grave: 'Grave', omoda: 'AVALIAÇÃO OMODA', naoClassificado: 'Não Classificado' };
            badge.style.display = 'inline-flex';
            const drillText = badge.querySelector('.drill-text');
            if (drillText) {
                drillText.textContent = `🔍 ${auditorDrillAuditor} · ${gravLabels[auditorDrillGravidade] || ''}`;
            }
        } else {
            badge.style.display = 'none';
        }
    }

    const ncRows = filteredRows.filter(r => {
        if (r.correto || !r.obs || !r.obs.trim()) return false;
        // Drill-down do gráfico por auditor tem prioridade
        if (auditorDrillAuditor) {
            if (r.auditor !== auditorDrillAuditor) return false;
            if (auditorDrillGravidade === 'naoClassificado') return !r.gravidade;
            return r.gravidade === auditorDrillGravidade;
        }
        if (topObsGravidade === 'all') return true;
        if (topObsGravidade === 'naoClassificado') return !r.gravidade;
        return r.gravidade === topObsGravidade;
    });
    const byObs = {};

    ncRows.forEach(r => {
        const obs = r.obs.trim();
        if (!byObs[obs]) {
            byObs[obs] = { passivel: 0, grave: 0, omoda: 0, naoClassificado: 0, total: 0 };
        }
        byObs[obs].total++;
        if (r.gravidade === 'passivel') byObs[obs].passivel++;
        else if (r.gravidade === 'grave') byObs[obs].grave++;
        else if (r.gravidade === 'omoda') byObs[obs].omoda++;
        else byObs[obs].naoClassificado++;
    });

    const topObs = Object.entries(byObs)
        .sort((a, b) => b[1].total - a[1].total)
        .slice(0, topObsLimit);

    const labels = topObs.map(([obs]) => obs.length > 62 ? `${obs.slice(0, 59)}...` : obs);
    const passivelData = topObs.map(([, v]) => v.passivel);
    const graveData = topObs.map(([, v]) => v.grave);
    const omodaData = topObs.map(([, v]) => v.omoda);
    const naoClassifData = topObs.map(([, v]) => v.naoClassificado);

    const obsCanvas = document.getElementById('gravidadeObsChart');
    if (!obsCanvas) return;
    const ctx = obsCanvas.getContext('2d');
    if (!ctx) return;
    charts.gravidadeObs = new Chart(ctx, {
        type: 'bar',
        data: {
            labels,
            datasets: [
                {
                    label: 'Passível de Ajuste',
                    data: passivelData,
                    backgroundColor: '#F59E0B',
                    borderColor: '#D97706',
                    borderWidth: 1,
                    borderRadius: 4,
                    borderSkipped: false,
                },
                {
                    label: 'Grave (Prejuízo Financeiro)',
                    data: graveData,
                    backgroundColor: '#DC2626',
                    borderColor: '#B91C1C',
                    borderWidth: 1,
                    borderRadius: 4,
                    borderSkipped: false,
                },
                {
                    label: 'AVALIAÇÃO OMODA',
                    data: omodaData,
                    backgroundColor: '#0EA5E9',
                    borderColor: '#0284C7',
                    borderWidth: 1,
                    borderRadius: 4,
                    borderSkipped: false,
                },
                {
                    label: 'Não Classificado',
                    data: naoClassifData,
                    backgroundColor: '#94A3B8',
                    borderColor: '#64748B',
                    borderWidth: 1,
                    borderRadius: 4,
                    borderSkipped: false,
                }
            ]
        },
        options: {
            indexAxis: 'y',
            responsive: true,
            maintainAspectRatio: true,
            interaction: { mode: 'index', intersect: false },
            plugins: {
                legend: {
                    labels: { font: CHART_FONT, color: C.txt, boxWidth: 12, padding: 12 }
                },
                datalabels: {
                    display: true,
                    font: { ...CHART_FONT, weight: '700', size: 11 },
                    color: '#000',
                    formatter: (value) => value > 0 ? value : '',
                    anchor: 'center',
                    align: 'center',
                    clamp: true
                },
                tooltip: {
                    callbacks: {
                        title: (items) => {
                            if (!items || !items.length) return '';
                            return topObs[items[0].dataIndex]?.[0] || '';
                        },
                        label: (ctx) => ` ${ctx.dataset.label}: ${ctx.parsed.x}`,
                        afterBody: (items) => {
                            if (!items || !items.length) return '';
                            const idx = items[0].dataIndex;
                            const total = passivelData[idx] + graveData[idx] + omodaData[idx] + naoClassifData[idx];
                            return `Total desta observação: ${total}`;
                        }
                    }
                }
            },
            scales: {
                x: {
                    stacked: true,
                    beginAtZero: true,
                    grid: { color: C.grid },
                    ticks: { font: CHART_FONT, color: C.txt, precision: 0 }
                },
                y: {
                    stacked: true,
                    grid: { display: false },
                    ticks: { font: { ...CHART_FONT, size: 11 }, color: C.txt }
                }
            }
        }
    });
}

/* ── PARETO ──────────────────────────────────────────────── */
const STOPWORDS = new Set([
    'de', 'a', 'o', 'que', 'e', 'do', 'da', 'em', 'um', 'para', 'uma', 'os', 'no', 'se', 'na', 'por',
    'mais', 'as', 'dos', 'como', 'mas', 'foi', 'ao', 'ele', 'das', 'tem', 'seu', 'sua', 'ou', 'ser',
    'quando', 'muito', 'nos', 'ja', 'esta', 'eu', 'tambem', 'so', 'pelo', 'pela', 'ate', 'isso',
    'ela', 'entre', 'era', 'depois', 'sem', 'mesmo', 'aos', 'ter', 'seus', 'quem', 'nas', 'me',
    'esse', 'eles', 'estao', 'voce', 'tinha', 'foram', 'essa', 'num', 'nem', 'suas', 'meu', 'as',
    'minha', 'pelos', 'elas', 'havia', 'seja', 'qual', 'sera', 'nos', 'lhe', 'essas', 'esses',
    'pelas', 'este', 'dele', 'tu', 'te', 'ok', 'sim', 'nao', 'obs', 'item', 'campo', 'numero',
    'informado', 'informada', 'preenchido', 'preenchida', 'incorreto', 'incorreta', 'falta',
    'esta', 'nesta', 'aqui', 'com', 'sobre', 'entre', 'esse', 'essa', 'pelo', 'pela', 'para',
    'data', 'valor', 'ano', 'dias', 'dia', 'mes', 'numero', 'nro', 'nr', 'ref', 'esta', 'nao',
    'sim', 'corr', 'err', 'tipo', 'cod', 'codigo', 'linha',
]);

function renderPareto() {
    destroyC('pareto');
    
    // Get original observations without normalization
    const naoObs = filteredRows
        .filter(r => !r.correto && r.obs && r.obs.trim().length > 0)
        .map(r => r.obs.trim());

    if (naoObs.length === 0) {
        const ctx = document.getElementById('paretoChart').getContext('2d');
        charts.pareto = new Chart(ctx, {
            type: 'bar',
            data: { labels: ['Sem ocorrências registradas'], datasets: [{ data: [0], backgroundColor: '#E8EEF5' }] },
            options: { plugins: { legend: { display: false } } }
        });
        return;
    }

    // Count exact observation values
    const obsCounts = {};
    naoObs.forEach(obs => {
        obsCounts[obs] = (obsCounts[obs] || 0) + 1;
    });

    // Get top 8 observations
    const cats = Object.entries(obsCounts)
        .map(([label, count]) => ({ label, count }))
        .sort((a, b) => b.count - a.count)
        .slice(0, 8);

    // Color ramp: blue gradient
    const palette = [
        '#0369A1', '#0284C7', '#0EA5E9', '#38BDF8', '#7DD3FC',
        '#BAE6FD', '#2563EB', '#3B82F6'
    ];

    const ctx = document.getElementById('paretoChart').getContext('2d');
    charts.pareto = new Chart(ctx, {
        type: 'bar',
        data: {
            labels: cats.map(c => c.label),
            datasets: [{
                label: 'Ocorrências',
                data: cats.map(c => c.count),
                backgroundColor: cats.map((_, i) => palette[i % palette.length]),
                borderColor: cats.map((_, i) => palette[i % palette.length]),
                borderWidth: 1,
                borderRadius: 6,
                borderSkipped: false,
            }]
        },
        options: {
            indexAxis: 'y',
            responsive: true,
            maintainAspectRatio: true,
            plugins: {
                legend: { display: false },
                datalabels: {
                    display: true,
                    font: { ...CHART_FONT, weight: '600', size: 11 },
                    color: '#000',
                    formatter: (value) => value,
                    offset: 4,
                    anchor: 'center',
                    align: 'center'
                },
                tooltip: {
                    callbacks: {
                        label: ctx => {
                            const pct = naoObs.length > 0
                                ? (ctx.parsed.x / naoObs.length * 100).toFixed(1) : 0;
                            return ` ${ctx.parsed.x} ocorrências  (${pct}% dos erros)`;
                        }
                    }
                }
            },
            scales: {
                x: { grid: { color: C.grid }, ticks: { font: CHART_FONT, color: C.txt } },
                y: { grid: { display: false }, ticks: { font: { ...CHART_FONT, size: 12 }, color: C.txt } }
            }
        }
    });
}

/* ── AUDITOR ─────────────────────────────────────────────── */
function renderAuditor() {
    destroyC('auditor');
    const amap = {};
    filteredRows.forEach(r => {
        if (!r.auditor) return;
        if (!amap[r.auditor]) amap[r.auditor] = { sim: 0, nao: 0 };
        if (r.correto) amap[r.auditor].sim++; else amap[r.auditor].nao++;
    });
    const auds = Object.keys(amap).sort((a, b) =>
        (amap[b].sim + amap[b].nao) - (amap[a].sim + amap[a].nao)
    );
    const simD = auds.map(a => amap[a].sim);
    const naoD = auds.map(a => amap[a].nao);

    const ctx = document.getElementById('auditorChart').getContext('2d');
    charts.auditor = new Chart(ctx, {
        type: 'bar',
        data: {
            labels: auds,
            datasets: [
                { label: 'Conformes (Sim)', data: simD, backgroundColor: C.simBg, borderColor: C.sim, borderWidth: 1, borderRadius: 5, borderSkipped: false },
                { label: 'Não Conformes (Não)', data: naoD, backgroundColor: C.naoBg, borderColor: C.nao, borderWidth: 1, borderRadius: 5, borderSkipped: false },
            ]
        },
        options: {
            responsive: true,
            maintainAspectRatio: true,
            interaction: { mode: 'index' },
            plugins: {
                legend: { labels: { font: CHART_FONT, color: C.txt, boxWidth: 12, padding: 14 } },
                datalabels: {
                    display: true,
                    font: { ...CHART_FONT, weight: '600', size: 11 },
                    color: '#000',
                    formatter: (value) => value,
                    offset: 4,
                    anchor: 'center',
                    align: 'center'
                },
                tooltip: {
                    callbacks: {
                        label: ctx => {
                            const a = auds[ctx.dataIndex];
                            const tot = amap[a].sim + amap[a].nao;
                            const p = tot > 0 ? (ctx.parsed.y / tot * 100).toFixed(1) : 0;
                            return ` ${ctx.dataset.label}: ${ctx.parsed.y}  (${p}%)`;
                        }
                    }
                }
            },
            scales: {
                x: { grid: { display: false }, ticks: { font: CHART_FONT, color: C.txt } },
                y: { grid: { color: C.grid }, ticks: { font: CHART_FONT, color: C.txt } }
            }
        }
    });
}

/* ── AUDITOR EVOLUTION ──────────────────────────────────── */
function onAuditorEvolutionChange() {
    const sel = document.getElementById('fAuditorEvolution');
    selectedAuditorEvolution = sel?.value || 'all';
    renderAuditorEvolution();
}

function renderAuditorEvolution() {
    destroyC('auditorEvolution');

    // Agrupar por auditor e mês
    let auditorData = {};
    filteredRows.forEach(r => {
        if (!r.auditor || !r.isoAud) return;
        const month = r.isoAud.slice(0, 7); // YYYY-MM
        if (!auditorData[r.auditor]) auditorData[r.auditor] = {};
        if (!auditorData[r.auditor][month]) auditorData[r.auditor][month] = { sim: 0, total: 0 };
        auditorData[r.auditor][month].total++;
        if (r.correto) auditorData[r.auditor][month].sim++;
    });

    // Popular o select de auditores
    const sel = document.getElementById('fAuditorEvolution');
    if (sel) {
        const auditoresList = Object.keys(auditorData).sort();
        const currentVal = selectedAuditorEvolution;
        sel.innerHTML = '<option value="all">Todos os auditores</option>' +
            auditoresList.map(a => `<option value="${a}" ${a === currentVal ? 'selected' : ''}>${a}</option>`).join('');
        if (currentVal !== 'all' && !auditoresList.includes(currentVal)) {
            selectedAuditorEvolution = 'all';
            sel.value = 'all';
        }
    }

    // Filtrar auditores conforme seleção
    const auditoresFiltrados = selectedAuditorEvolution === 'all'
        ? Object.keys(auditorData)
        : Object.keys(auditorData).filter(a => a === selectedAuditorEvolution);

    // Obter todos os meses únicos e ordenar
    const allMonths = [...new Set(filteredRows.map(r => r.isoAud?.slice(0, 7)).filter(Boolean))].sort();

    // Cores para os auditores
    const colors = ['#0369A1', '#059669', '#DC2626', '#F59E0B', '#6D28D9', '#EC4899', '#10B981', '#3B82F6', '#F97316', '#8B5CF6'];

    // Criar datasets para cada auditor
    const allAuditores = Object.keys(auditorData);
    const datasets = auditoresFiltrados.map((auditor) => {
        const index = allAuditores.indexOf(auditor);
        const data = allMonths.map(month => {
            const d = auditorData[auditor][month];
            return d && d.total > 0 ? +(d.sim / d.total * 100).toFixed(1) : null;
        });
        const color = colors[index % colors.length];
        return {
            label: auditor,
            data,
            borderColor: color,
            backgroundColor: color + '1a',
            fill: false,
            tension: 0.4,
            pointRadius: 3,
            pointHoverRadius: 5,
            borderWidth: 2,
        };
    });

    // Labels formatados
    const fmtLabels = allMonths.map(m => {
        const [y, mm] = m.split('-');
        const monthNames = ['Jan', 'Fev', 'Mar', 'Abr', 'Mai', 'Jun', 'Jul', 'Ago', 'Set', 'Out', 'Nov', 'Dez'];
        return `${monthNames[parseInt(mm) - 1]}/${y.slice(2)}`;
    });

    const ctx = document.getElementById('auditorEvolutionChart').getContext('2d');
    charts.auditorEvolution = new Chart(ctx, {
        type: 'line',
        data: {
            labels: fmtLabels,
            datasets
        },
        options: {
            responsive: true,
            maintainAspectRatio: true,
            aspectRatio: 2.5,
            layout: {
                padding: {
                    top: 30,
                    bottom: 20,
                    left: 20,
                    right: 20
                }
            },
            interaction: { mode: 'index', intersect: false },
            plugins: {
                legend: {
                    display: true,
                    position: 'top',
                    align: 'center',
                    labels: { font: CHART_FONT, color: C.txt, boxWidth: 12, padding: 16 }
                },
                datalabels: {
                    display: true,
                    clip: false,
                    font: { ...CHART_FONT, weight: '600', size: 10 },
                    color: '#000',
                    formatter: (value) => value !== null ? value + '%' : '',
                    offset: 6,
                    anchor: 'end',
                    align: 'top'
                },
                tooltip: {
                    callbacks: {
                        label: ctx => {
                            const value = ctx.parsed.y;
                            const auditor = ctx.dataset.label;
                            const month = allMonths[ctx.dataIndex];
                            const d = auditorData[auditor][month];
                            if (d) {
                                return `${auditor}: ${value}% (${d.sim}/${d.total})`;
                            }
                            return `${auditor}: ${value}%`;
                        }
                    }
                }
            },
            scales: {
                x: {
                    grid: { color: C.grid },
                    ticks: { font: CHART_FONT, color: C.txt, maxRotation: 45 }
                },
                y: {
                    beginAtZero: true,
                    max: 100,
                    grid: { color: C.grid },
                    ticks: { font: CHART_FONT, color: C.txt, callback: value => value + '%' }
                }
            }
        }
    });
}

/* ═══════════════════════════════════════════════════════════
   EXPORTAR CSV
   ═══════════════════════════════════════════════════════════ */
function exportCSV() {
    const header = ['Projeto', 'Data Calibração', 'Data Auditoria', 'Chassi', 'Auditor', 'Auditoria Correta?', 'Observação', 'Gravidade / Impacto', 'SLA (dias)'];
    const rows = filteredRows.map(r => [
        r.projeto,
        fmtDate(r.dataCal),
        fmtDate(r.dataAud),
        r.chassi,
        r.auditor,
        r.correto ? 'Sim' : 'Não',
        r.obs,
        r.gravidadeRaw,
        r.sla !== null ? r.sla : '',
    ]);
    const csv = [header, ...rows]
        .map(row => row.map(v => `"${String(v || '').replace(/"/g, '""')}"`).join(','))
        .join('\r\n');
    const blob = new Blob(['\uFEFF' + csv], { type: 'text/csv;charset=utf-8' });
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = `qa_auditoria_${new Date().toISOString().slice(0, 10)}.csv`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
}

/* ═══════════════════════════════════════════════════════════
   INIT
   ═══════════════════════════════════════════════════════════ */
loadAllData();
