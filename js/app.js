/*
 * Simulador de metas e comissões para assessorias.
 * Todo o processamento do arquivo acontece no navegador. Não há chamadas
 * de rede com os dados da planilha nem persistência do histórico em storage.
 */
(() => {
  'use strict';

  const MONTHS = [
    { n: 1, short: 'JAN', name: 'Janeiro' },
    { n: 2, short: 'FEV', name: 'Fevereiro' },
    { n: 3, short: 'MAR', name: 'Março' },
    { n: 4, short: 'ABR', name: 'Abril' },
    { n: 5, short: 'MAI', name: 'Maio' },
    { n: 6, short: 'JUN', name: 'Junho' },
    { n: 7, short: 'JUL', name: 'Julho' },
    { n: 8, short: 'AGO', name: 'Agosto' },
    { n: 9, short: 'SET', name: 'Setembro' },
    { n: 10, short: 'OUT', name: 'Outubro' },
    { n: 11, short: 'NOV', name: 'Novembro' },
    { n: 12, short: 'DEZ', name: 'Dezembro' }
  ];
  const MONTH_LOOKUP = Object.fromEntries(MONTHS.map(m => [m.short, m.n]));
  const $ = (selector) => document.querySelector(selector);

  const state = {
    rows: [],
    policy: [],
    fileName: '',
    year: null,
    agencies: [],
    startMonth: 1,
    endMonth: 9,
    includePartial: false,
    agencyFilter: 'ALL',
    baselineMode: 'average',
    targetBasis: 'recovered',
    growth: 10,
    globalRules: { mode: 'progressive', tiers: [{ min: 0, rate: 0 }] },
    rulesByAgency: {},
    ruleScope: 'ALL',
    activeTab: 'overview'
  };

  const money = (value, compact = false) => {
    const n = Number.isFinite(Number(value)) ? Number(value) : 0;
    if (compact && Math.abs(n) >= 1000000) {
      return n.toLocaleString('pt-BR', { style: 'currency', currency: 'BRL', notation: 'compact', maximumFractionDigits: 1 });
    }
    return n.toLocaleString('pt-BR', { style: 'currency', currency: 'BRL', maximumFractionDigits: 2 });
  };
  const number = (value, digits = 0) => (Number(value) || 0).toLocaleString('pt-BR', { minimumFractionDigits: digits, maximumFractionDigits: digits });
  const percent = (value, digits = 2) => `${number((Number(value) || 0) * 100, digits)}%`;
  const percentPoints = (value, digits = 2) => `${number(Number(value) || 0, digits)}%`;
  const esc = (value) => String(value ?? '').replace(/[&<>"']/g, char => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[char]));
  const normalized = (value) => String(value ?? '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').trim().toUpperCase();

  function parseOptionalNumeric(value) {
    if (typeof value === 'number' && Number.isFinite(value)) return value;
    if (value === null || value === undefined) return null;
    const text = String(value).trim();
    if (!text || text === '-' || text === '—' || text.includes('#VALUE')) return null;
    let cleaned = text.replace(/R\$/gi, '').replace(/\s/g, '');
    if (cleaned.includes(',')) cleaned = cleaned.replace(/\./g, '').replace(',', '.');
    cleaned = cleaned.replace(/[^0-9.-]/g, '');
    if (!cleaned) return null;
    const parsed = Number(cleaned);
    return Number.isFinite(parsed) ? parsed : null;
  }

  function parseNumeric(value) {
    return parseOptionalNumeric(value) ?? 0;
  }

  const sumField = (rows, field) => rows.reduce((sum, row) => sum + (Number.isFinite(row[field]) ? row[field] : 0), 0);
  const displayMoney = (value) => Number.isFinite(value) ? money(value) : '—';
  const displayPercent = (value) => Number.isFinite(value) ? percent(value) : '—';

  function inferYear(fileName) {
    const match = String(fileName).match(/(?:^|\D)(\d{2})[-_.](\d{2})[-_.](\d{2,4})(?:\D|$)/);
    if (!match) return null;
    const raw = Number(match[3]);
    return raw < 100 ? 2000 + raw : raw;
  }

  function findHeaderIndex(header, patterns) {
    return header.findIndex(cell => {
      const value = normalized(cell);
      return patterns.some(pattern => value.includes(pattern));
    });
  }

  function parseHistoryRows(rows) {
    let headerIndex = -1;
    let columns = null;
    for (let i = 0; i < Math.min(rows.length, 12); i += 1) {
      const candidate = rows[i] || [];
      const agency = findHeaderIndex(candidate, ['ASSESSORIA']);
      const month = findHeaderIndex(candidate, ['MES']);
      const recovered = findHeaderIndex(candidate, ['RECUPERADO']);
      if (agency >= 0 && month >= 0 && recovered >= 0) {
        headerIndex = i;
        columns = {
          agency,
          month,
          collection: findHeaderIndex(candidate, ['COBRANCA']),
          recovered,
          active: findHeaderIndex(candidate, ['RECUPERADO ATIVO']),
          canceled: findHeaderIndex(candidate, ['RECUPERADO CANCELADO']),
          commission: findHeaderIndex(candidate, ['COMISSAO']),
          conversion: findHeaderIndex(candidate, ['CONVERSAO']),
          costRatio: findHeaderIndex(candidate, ['CUSTO COMISSAO'])
        };
        break;
      }
    }
    if (headerIndex < 0) throw new Error('Não encontrei os cabeçalhos ASSESSORIA, MÊS e RECUPERADO na aba de histórico.');

    const data = [];
    for (let i = headerIndex + 1; i < rows.length; i += 1) {
      const row = rows[i] || [];
      const agencyRaw = String(row[columns.agency] ?? '').trim();
      const monthRaw = String(row[columns.month] ?? '').trim();
      const agencyKey = normalized(agencyRaw);
      const monthKey = normalized(monthRaw.split(/[\s(]/)[0]);
      const month = MONTH_LOOKUP[monthKey];
      if (!agencyRaw || agencyKey === '0' || agencyKey.includes('TOTAL') || normalized(monthRaw).includes('TOTAL') || !month) continue;

      const get = (index) => index >= 0 ? row[index] : null;
      const partial = normalized(monthRaw).includes('PARCIAL');
      data.push({
        agency: agencyRaw.trim(),
        month,
        monthLabel: MONTHS[month - 1].short,
        monthName: MONTHS[month - 1].name,
        partial,
        collection: parseNumeric(get(columns.collection)),
        recovered: parseNumeric(get(columns.recovered)),
        active: parseOptionalNumeric(get(columns.active)),
        canceled: parseOptionalNumeric(get(columns.canceled)),
        commission: parseOptionalNumeric(get(columns.commission)),
        conversion: parseOptionalNumeric(get(columns.conversion)),
        costRatio: parseOptionalNumeric(get(columns.costRatio)),
        sourceRow: i + 1
      });
    }
    if (!data.length) throw new Error('A aba de histórico foi localizada, mas não há registros mensais válidos para as assessorias.');
    return data;
  }

  function parsePolicyRows(rows) {
    let headerRow = -1;
    let startCol = -1;
    for (let i = 0; i < rows.length; i += 1) {
      const row = rows[i] || [];
      const idx = row.findIndex(cell => normalized(cell).includes('FAIXA DE ATRASO'));
      if (idx >= 0) {
        headerRow = i;
        startCol = idx;
        break;
      }
    }
    if (headerRow < 0) return [];
    const result = [];
    for (let i = headerRow + 1; i < rows.length; i += 1) {
      const row = rows[i] || [];
      const label = String(row[startCol] ?? '').trim();
      if (!label) break;
      const b2c = parseNumeric(row[startCol + 1]);
      const b2b = parseNumeric(row[startCol + 2]);
      if (typeof row[startCol + 1] === 'string' && normalized(row[startCol + 1]).includes('COMISSAO')) break;
      if (!Number.isFinite(b2c) || !Number.isFinite(b2b)) continue;
      result.push({ range: label, b2c, b2b });
    }
    return result;
  }

  function getFilteredRows(ignoreAgency = false) {
    return state.rows.filter(row => {
      if (row.month < state.startMonth || row.month > state.endMonth) return false;
      if (!state.includePartial && row.partial) return false;
      if (!ignoreAgency && state.agencyFilter !== 'ALL' && row.agency !== state.agencyFilter) return false;
      return true;
    });
  }

  function groupRows(rows) {
    const groups = new Map();
    rows.forEach(row => {
      if (!groups.has(row.agency)) {
        groups.set(row.agency, {
          agency: row.agency, records: [], totalCollection: 0, totalRecovered: 0,
          totalActive: 0, totalCanceled: 0, totalCommission: 0
        });
      }
      const group = groups.get(row.agency);
      group.records.push(row);
      group.totalCollection += row.collection;
      group.totalRecovered += row.recovered;
      if (Number.isFinite(row.active)) {
        group.totalActive += row.active;
        group.activeCount = (group.activeCount || 0) + 1;
      }
      if (Number.isFinite(row.canceled)) group.totalCanceled += row.canceled;
      if (Number.isFinite(row.commission)) {
        group.totalCommission += row.commission;
        group.commissionCount = (group.commissionCount || 0) + 1;
        group.commissionBaseRecovered = (group.commissionBaseRecovered || 0) + row.recovered;
      }
    });
    return [...groups.values()].map(group => {
      const complete = group.records.filter(row => !row.partial).sort((a, b) => a.month - b.month);
      const allSorted = [...group.records].sort((a, b) => a.month - b.month);
      const count = group.records.length;
      group.monthCount = count;
      group.avgRecovered = count ? group.totalRecovered / count : 0;
      group.activeCount = group.activeCount || 0;
      group.avgActive = group.activeCount ? group.totalActive / group.activeCount : 0;
      group.commissionCount = group.commissionCount || 0;
      group.commissionBaseRecovered = group.commissionBaseRecovered || 0;
      group.avgCommission = group.commissionCount ? group.totalCommission / group.commissionCount : 0;
      group.effectiveRate = group.commissionBaseRecovered > 0 ? group.totalCommission / group.commissionBaseRecovered : 0;
      group.lastComplete = complete.length ? complete[complete.length - 1] : null;
      group.lastRecord = allSorted.length ? allSorted[allSorted.length - 1] : null;
      return group;
    }).sort((a, b) => a.agency.localeCompare(b.agency, 'pt-BR'));
  }

  function getBaseline(group) {
    if (state.baselineMode === 'last') {
      const row = group.lastComplete;
      if (!row) return 0;
      return state.targetBasis === 'active' ? row.active : row.recovered;
    }
    return state.targetBasis === 'active' ? group.avgActive : group.avgRecovered;
  }

  function getTargetRows() {
    return groupRows(getFilteredRows()).map(group => {
      const base = getBaseline(group);
      const target = base * (1 + state.growth / 100);
      return { ...group, base, target, increase: target - base, annualTarget: target * 12 };
    });
  }

  function getBaselineRate(agency) {
    const source = state.rows.filter(row => row.agency === agency && row.month >= state.startMonth && row.month <= state.endMonth && !row.partial && Number.isFinite(row.commission));
    const recovery = source.reduce((sum, row) => sum + row.recovered, 0);
    const commission = source.reduce((sum, row) => sum + row.commission, 0);
    return recovery > 0 ? commission / recovery : 0;
  }

  function getGlobalBaselineRate() {
    const source = state.rows.filter(row => row.month >= state.startMonth && row.month <= state.endMonth && !row.partial && Number.isFinite(row.commission));
    const recovery = source.reduce((sum, row) => sum + row.recovered, 0);
    const commission = source.reduce((sum, row) => sum + row.commission, 0);
    return recovery > 0 ? commission / recovery : 0;
  }

  function initializeRules() {
    state.globalRules = { mode: 'progressive', tiers: [{ min: 0, rate: getGlobalBaselineRate() * 100 }] };
    state.rulesByAgency = {};
    state.ruleScope = 'ALL';
  }

  function cloneRules(rules) {
    return { mode: rules.mode, tiers: rules.tiers.map(tier => ({ min: tier.min, rate: tier.rate })) };
  }

  function getEditorRules() {
    if (state.ruleScope !== 'ALL' && state.rulesByAgency[state.ruleScope]) return state.rulesByAgency[state.ruleScope];
    return state.globalRules;
  }

  function getRulesForAgency(agency) {
    return state.rulesByAgency[agency] || state.globalRules;
  }

  function ensureEditableRules() {
    if (state.ruleScope !== 'ALL' && !state.rulesByAgency[state.ruleScope]) {
      state.rulesByAgency[state.ruleScope] = cloneRules(state.globalRules);
    }
    return getEditorRules();
  }

  function calculateCommission(gross, rules) {
    const value = Math.max(0, Number(gross) || 0);
    const tiers = [...rules.tiers].sort((a, b) => a.min - b.min);
    if (!tiers.length || value <= 0) return 0;
    if (rules.mode === 'attained') {
      let selected = tiers[0];
      tiers.forEach(tier => { if (value >= tier.min) selected = tier; });
      return value * selected.rate / 100;
    }
    let total = 0;
    for (let i = 0; i < tiers.length; i += 1) {
      const tier = tiers[i];
      const next = tiers[i + 1];
      const low = Math.max(0, tier.min);
      const high = next ? Math.min(value, next.min) : value;
      const slice = Math.max(0, high - low);
      total += slice * tier.rate / 100;
      if (value <= (next ? next.min : value)) break;
    }
    return total;
  }

  function setStatus(message, type = '') {
    const line = $('#statusLine');
    line.classList.remove('success', 'error');
    if (type) line.classList.add(type);
    $('#statusText').textContent = message;
  }

  let toastTimer = null;
  function toast(message, type = '') {
    const node = $('#toast');
    node.textContent = message;
    node.className = `toast show ${type}`.trim();
    window.clearTimeout(toastTimer);
    toastTimer = window.setTimeout(() => node.className = 'toast', 3000);
  }

  function populateFilters() {
    const agencyOptions = ['<option value="ALL">Todas as assessorias</option>', ...state.agencies.map(name => `<option value="${esc(name)}">${esc(name)}</option>`)].join('');
    $('#agencyFilter').innerHTML = agencyOptions;
    $('#ruleScope').innerHTML = ['<option value="ALL">Todas as assessorias</option>', ...state.agencies.map(name => `<option value="${esc(name)}">${esc(name)}</option>`)].join('');

    const months = MONTHS.map(month => `<option value="${month.n}">${month.short} · ${month.name}</option>`).join('');
    $('#startMonth').innerHTML = months;
    $('#endMonth').innerHTML = months;
    const observedMonths = state.rows.map(row => row.month);
    const firstMonth = Math.min(...observedMonths);
    const completeMonths = state.rows.filter(row => !row.partial).map(row => row.month);
    const lastObserved = Math.max(...observedMonths);
    state.startMonth = firstMonth;
    state.endMonth = lastObserved;
    $('#startMonth').value = String(firstMonth);
    $('#endMonth').value = String(lastObserved);
    $('#agencyFilter').value = 'ALL';
    $('#ruleScope').value = 'ALL';
    $('#includePartial').checked = false;
    state.agencyFilter = 'ALL';
    state.includePartial = false;
  }

  async function processWorkbookBuffer(buffer, fileName) {
    if (!window.XLSX) {
      setStatus('A biblioteca de leitura do Excel não foi carregada. Verifique a conexão e atualize a página.', 'error');
      toast('Não foi possível carregar o leitor de Excel.', 'error');
      return;
    }
    setStatus('Lendo a planilha localmente…');
    $('#statusFile').textContent = fileName;
    try {
        const workbook = window.XLSX.read(buffer, { type: 'array', cellDates: true });
        const sheetName = workbook.SheetNames.find(name => normalized(name) === 'HISTORICO ASSESSORIAS');
        if (!sheetName) throw new Error('A aba “Histórico Assessorias” não foi encontrada. Verifique se selecionou a planilha correta.');
        const historySheet = workbook.Sheets[sheetName];
        const historyArray = window.XLSX.utils.sheet_to_json(historySheet, { header: 1, raw: true, defval: null, blankrows: false });
        const parsedRows = parseHistoryRows(historyArray);
        const policy = parsePolicyRows(historyArray);
        state.rows = parsedRows;
        state.policy = policy;
        state.fileName = fileName;
        state.year = inferYear(fileName);
        state.agencies = [...new Set(parsedRows.map(row => row.agency))].sort((a, b) => a.localeCompare(b, 'pt-BR'));
        populateFilters();
        initializeRules();
        $('#emptyState').classList.add('hidden');
        $('#workspace').classList.remove('hidden');
        $('#clearDataBtn').classList.remove('hidden');
        $('#statusFile').textContent = `${fileName} · ${parsedRows.length} registros`;
        const partialCount = parsedRows.filter(row => row.partial).length;
        const partialMessage = partialCount ? ` ${partialCount} registro(s) de mês parcial foram identificados e estão excluídos por padrão.` : '';
        setStatus(`Histórico importado: ${state.agencies.length} assessorias e ${parsedRows.length} registros mensais.${partialMessage}`, 'success');
        const first = Math.min(...parsedRows.map(row => row.month));
        const last = Math.max(...parsedRows.map(row => row.month));
        const yearText = state.year ? ` · referência ${state.year} (inferida do nome do arquivo)` : ' · ano não informado no nome do arquivo';
        $('#dataContext').innerHTML = `<strong>${esc(fileName)}</strong> · ${state.agencies.length} assessorias · ${MONTHS[first - 1].short} a ${MONTHS[last - 1].short}${yearText} · ${partialCount} linha(s) parcial(is) identificada(s)`;
        $('#policyFoot').textContent = policy.length ? `A planilha forneceu ${policy.length} faixas de atraso. Os percentuais são exibidos como referência da política registrada no arquivo, sem aplicação automática no simulador por volume.` : 'Não foi localizada uma tabela de política de comissão por atraso na aba de histórico.';
        renderAll();
        activateTab('overview');
        toast('Histórico carregado. As simulações estão prontas.');
    } catch (error) {
      console.error(error);
      setStatus(error.message || 'Ocorreu um erro ao interpretar a planilha.', 'error');
      toast(error.message || 'Erro ao processar a planilha.', 'error');
    }
  }

  async function importWorkbook(file) {
    if (!file) return;
    setStatus('Lendo a planilha localmente…');
    $('#statusFile').textContent = file.name;
    try {
      const buffer = await file.arrayBuffer();
      await processWorkbookBuffer(buffer, file.name);
    } catch (error) {
      setStatus('Não foi possível ler o arquivo selecionado.', 'error');
      toast(error.message || 'Falha ao abrir o arquivo.', 'error');
    }
  }

  async function tryLoadLocalHistory() {
    // Em ambiente local, procura automaticamente o arquivo ignorado pelo Git em /dados.
    // O primeiro nome é o padrão recomendado; o segundo aceita o nome do arquivo de origem enviado.
    // No GitHub Pages, os caminhos não existirão porque a planilha não deve ser publicada.
    const candidates = [
      './dados/historico.xlsx',
      './dados/analise_assessorias_14-10-26(1).xlsx'
    ];
    for (const path of candidates) {
      try {
        const response = await fetch(path, { cache: 'no-store' });
        if (!response.ok) continue;
        const buffer = await response.arrayBuffer();
        await processWorkbookBuffer(buffer, path.replace('./', ''));
        return;
      } catch (_) {
        // Acesso via file:// ou publicação sem o arquivo privado: mantém a tela de importação.
        return;
      }
    }
  }

  function updateMetrics(groups) {
    const rows = getFilteredRows();
    const recovered = rows.reduce((sum, row) => sum + row.recovered, 0);
    const commissionRows = rows.filter(row => Number.isFinite(row.commission));
    const commission = sumField(commissionRows, 'commission');
    const commissionBaseRecovered = commissionRows.reduce((sum, row) => sum + row.recovered, 0);
    const averageSum = groups.reduce((sum, group) => sum + group.avgRecovered, 0);
    $('#metricRecovered').textContent = money(recovered, true);
    $('#metricRecoveredFoot').textContent = `${rows.length} registro(s) no recorte`;
    $('#metricMonthly').textContent = money(averageSum, true);
    $('#metricCommission').textContent = money(commission, true);
    $('#metricCommission').nextElementSibling.textContent = `${commissionRows.length} registro(s) com comissão apurada`;
    $('#metricCost').textContent = percent(commissionBaseRecovered > 0 ? commission / commissionBaseRecovered : 0, 2);
    $('#metricCost').nextElementSibling.textContent = 'Somente meses com comissão informada';
  }

  function renderOverview(groups) {
    const rows = getFilteredRows();
    const total = groups.reduce((sum, group) => sum + group.totalRecovered, 0);
    const ranking = [...groups].sort((a, b) => b.totalRecovered - a.totalRecovered);
    const max = ranking.length ? Math.max(...ranking.map(group => group.totalRecovered), 1) : 1;
    $('#overviewPeriodChip').textContent = `${MONTHS[state.startMonth - 1].short} – ${MONTHS[state.endMonth - 1].short}${state.includePartial ? ' · parciais incluídos' : ' · parciais excluídos'}`;
    $('#rankingChart').innerHTML = ranking.length ? ranking.map(group => `
      <div class="ranking-row">
        <div class="ranking-name">${esc(group.agency)}</div>
        <div class="ranking-track" aria-label="${esc(group.agency)}: ${money(group.totalRecovered)}"><div class="ranking-fill" style="width:${Math.max(1, group.totalRecovered / max * 100)}%"></div></div>
        <div class="ranking-value">${money(group.totalRecovered, true)}</div>
      </div>`).join('') : '<p class="muted-empty">Sem registros neste recorte.</p>';

    const leader = ranking[0];
    const lowestCost = [...groups].filter(group => group.totalRecovered > 0).sort((a, b) => a.effectiveRate - b.effectiveRate)[0];
    const knownCommissionRows = rows.filter(row => Number.isFinite(row.commission));
    const commissionTotal = sumField(knownCommissionRows, 'commission');
    const commissionRecovery = knownCommissionRows.reduce((sum, row) => sum + row.recovered, 0);
    const avgCost = commissionRecovery > 0 ? commissionTotal / commissionRecovery : 0;
    const averageCount = groups.reduce((sum, group) => sum + group.monthCount, 0);
    $('#insightContent').innerHTML = `
      <div class="insight-item"><span>Maior recuperação</span><strong>${leader ? esc(leader.agency) : '—'}</strong><small>${leader ? money(leader.totalRecovered, true) : 'Sem dados'}</small></div>
      <div class="insight-item"><span>Menor custo efetivo</span><strong>${lowestCost ? esc(lowestCost.agency) : '—'}</strong><small>${lowestCost ? percent(lowestCost.effectiveRate) + ' sobre o recuperado' : 'Sem dados'}</small></div>
      <div class="insight-message">${groups.length ? `A recuperação consolidada foi de <strong>${money(total, true)}</strong>, com custo de comissão ponderado de <strong>${percent(avgCost)}</strong> nos registros em que a comissão foi informada. A média mensal individual considera somente os meses com registro para cada assessoria.` : 'Não há dados suficientes para gerar uma leitura do período.'}</div>`;

    $('#overviewTableBody').innerHTML = groups.length ? groups.map(group => `
      <tr><td>${esc(group.agency)}</td><td class="numeric">${group.monthCount}</td><td class="numeric">${money(group.totalRecovered)}</td><td class="numeric">${money(group.avgRecovered)}</td><td class="numeric">${group.commissionCount ? money(group.avgCommission) : '—'}</td><td class="numeric">${group.commissionCount ? percent(group.effectiveRate) : '—'}</td></tr>`).join('') : '<tr><td colspan="6">Sem dados para os filtros selecionados.</td></tr>';
  }

  function renderTargets() {
    const targets = getTargetRows();
    const baselineTotal = targets.reduce((sum, row) => sum + row.base, 0);
    const targetTotal = targets.reduce((sum, row) => sum + row.target, 0);
    const increaseTotal = targetTotal - baselineTotal;
    const annualTotal = targets.reduce((sum, row) => sum + row.annualTarget, 0);
    const basisLabel = state.targetBasis === 'active' ? 'Recuperado ativo' : 'Recuperado bruto';
    $('#baselineHint').textContent = state.baselineMode === 'last' ? 'Usa o último mês completo de cada assessoria dentro do recorte.' : 'A média considera apenas os meses em que a assessoria possui registro.';
    $('#targetHighlight').innerHTML = `
      <div class="target-intro"><span>Meta conjunta · ${esc(basisLabel)}</span><strong>${number(state.growth, 0)}% acima da referência</strong></div>
      <div class="target-kpi"><span>Meta mensal somada</span><strong>${money(targetTotal, true)}</strong></div>
      <div class="target-kpi"><span>Adicional necessário / mês</span><strong class="accent">${money(increaseTotal, true)}</strong></div>
      <div class="target-kpi"><span>Projeção em 12 meses</span><strong>${money(annualTotal, true)}</strong></div>`;
    $('#targetsTableBody').innerHTML = targets.length ? targets.map(row => `
      <tr><td>${esc(row.agency)}</td><td class="numeric">${money(row.base)}</td><td class="numeric"><strong>${money(row.target)}</strong></td><td class="numeric positive">${money(row.increase)}</td><td class="numeric">${money(row.annualTarget)}</td><td class="numeric">${number(state.growth, 1)}%</td></tr>`).join('') : '<tr><td colspan="6">Sem dados para os filtros selecionados.</td></tr>';
  }

  function renderRuleEditor() {
    const rules = getEditorRules();
    $('#commissionMode').value = rules.mode;
    $('#modeExplainer').textContent = rules.mode === 'progressive'
      ? 'Na forma progressiva, cada percentual incide apenas sobre a parcela do valor dentro daquela faixa.'
      : 'Na faixa atingida, o percentual da última faixa alcançada incide sobre todo o valor recuperado.';
    const sorted = [...rules.tiers].sort((a, b) => a.min - b.min);
    $('#tiersTableBody').innerHTML = sorted.map((tier, index) => `
      <tr>
        <td><input type="number" min="0" step="1000" value="${Number(tier.min).toFixed(2)}" data-tier-index="${index}" data-tier-field="min" aria-label="Valor inicial da faixa ${index + 1}"></td>
        <td><input class="rate-input" type="number" min="0" max="100" step="0.05" value="${Number(tier.rate).toFixed(3)}" data-tier-index="${index}" data-tier-field="rate" aria-label="Percentual da faixa ${index + 1}"></td>
        <td><button class="tier-delete" type="button" data-delete-tier="${index}" aria-label="Remover faixa ${index + 1}" ${sorted.length === 1 ? 'disabled' : ''}>×</button></td>
      </tr>`).join('');
    const refRate = state.ruleScope === 'ALL' ? getGlobalBaselineRate() : getBaselineRate(state.ruleScope);
    const label = state.ruleScope === 'ALL' ? 'Média ponderada de todas as assessorias' : `Referência histórica de ${esc(state.ruleScope)}`;
    $('#referenceRate').innerHTML = `${label}: <strong>${percent(refRate)}</strong> de comissão efetiva sobre a recuperação bruta em meses completos. ${state.ruleScope !== 'ALL' && !state.rulesByAgency[state.ruleScope] ? 'A tabela individual ainda herda o modelo geral.' : ''}`;
  }

  function renderCommissionComparison(targets) {
    const result = targets.map(row => {
      const rules = getRulesForAgency(row.agency);
      const simulated = calculateCommission(row.target, rules);
      const delta = simulated - row.avgCommission;
      return { ...row, simulated, delta, net: row.target - simulated };
    });
    $('#commissionTableBody').innerHTML = result.length ? result.map(row => `
      <tr><td>${esc(row.agency)}</td><td class="numeric">${money(row.target)}</td><td class="numeric">${money(row.avgCommission)}</td><td class="numeric"><strong>${money(row.simulated)}</strong></td><td class="numeric ${row.delta > 0 ? 'negative' : 'positive'}">${row.delta > 0 ? '+' : ''}${money(row.delta)}</td><td class="numeric">${money(row.net)}</td></tr>`).join('') : '<tr><td colspan="6">Sem dados para os filtros selecionados.</td></tr>';

    const selected = state.ruleScope === 'ALL' ? result : result.filter(row => row.agency === state.ruleScope);
    const gross = selected.reduce((sum, row) => sum + row.target, 0);
    const historical = selected.reduce((sum, row) => sum + row.avgCommission, 0);
    const simulated = selected.reduce((sum, row) => sum + row.simulated, 0);
    const delta = simulated - historical;
    const net = gross - simulated;
    $('#commissionPreviewMetrics').innerHTML = `
      <div class="preview-line"><span>Recuperação bruta na meta</span><strong>${money(gross, true)}</strong></div>
      <div class="preview-line"><span>Comissão histórica média</span><strong>${money(historical, true)}</strong></div>
      <div class="preview-line primary"><span>Comissão no novo cenário</span><strong>${money(simulated, true)}</strong></div>
      <div class="preview-line"><span>Variação de custo</span><strong class="${delta > 0 ? 'negative' : 'positive'}">${delta > 0 ? '+' : ''}${money(delta, true)}</strong></div>
      <div class="preview-line"><span>Recuperação líquida após comissão</span><strong>${money(net, true)}</strong></div>`;
    const scopeLabel = state.ruleScope === 'ALL' ? 'Todas as assessorias' : state.ruleScope;
    $('#commissionPreviewFoot').textContent = `${scopeLabel}: ${gross > 0 ? percent(simulated / gross) : '0,00%'} de custo de comissão no cenário. Os valores usam a meta projetada, não o pagamento efetivamente realizado.`;
    return result;
  }

  function renderPolicy() {
    $('#policyTableBody').innerHTML = state.policy.length ? state.policy.map(row => `
      <tr><td>${esc(row.range)}</td><td class="numeric">${percentPoints(row.b2c * 100, 2)}</td><td class="numeric">${percentPoints(row.b2b * 100, 2)}</td></tr>`).join('') : '<tr><td colspan="3">Tabela de política não encontrada no arquivo.</td></tr>';
  }

  function renderHistory() {
    const search = normalized($('#historySearch').value);
    const rows = getFilteredRows().filter(row => !search || normalized(row.agency).includes(search) || normalized(row.monthLabel).includes(search) || normalized(row.monthName).includes(search));
    rows.sort((a, b) => a.month - b.month || a.agency.localeCompare(b.agency, 'pt-BR'));
    $('#historyCount').textContent = `${rows.length} registro(s) exibido(s)`;
    $('#historyTableBody').innerHTML = rows.length ? rows.map(row => `
      <tr><td>${esc(row.agency)}</td><td>${esc(row.monthLabel)}${row.partial ? ' · parcial' : ''}</td><td class="numeric">${money(row.collection)}</td><td class="numeric">${money(row.recovered)}</td><td class="numeric">${displayMoney(row.active)}</td><td class="numeric">${displayMoney(row.canceled)}</td><td class="numeric">${displayMoney(row.commission)}</td><td class="numeric">${displayPercent(row.conversion)}</td><td class="numeric">${displayPercent(row.costRatio)}</td></tr>`).join('') : '<tr><td colspan="9">Nenhum registro corresponde à busca e aos filtros atuais.</td></tr>';
  }

  function renderAll() {
    const groups = groupRows(getFilteredRows());
    updateMetrics(groups);
    renderOverview(groups);
    renderTargets();
    renderRuleEditor();
    renderCommissionComparison(getTargetRows());
    renderPolicy();
    renderHistory();
  }

  function activateTab(tab) {
    state.activeTab = tab;
    document.querySelectorAll('.tab-button').forEach(button => {
      const active = button.dataset.tab === tab;
      button.classList.toggle('active', active);
      button.setAttribute('aria-selected', String(active));
    });
    document.querySelectorAll('.tab-panel').forEach(panel => panel.classList.toggle('hidden', panel.id !== `panel-${tab}`));
  }

  function csvValue(value) {
    if (value === null || value === undefined) return '';
    const text = String(value).replace(/\r?\n/g, ' ');
    return /[;"\n]/.test(text) ? `"${text.replace(/"/g, '""')}"` : text;
  }
  function csvNumber(value) {
    if (value === null || value === undefined || value === '') return '';
    const numericValue = Number(value);
    return Number.isFinite(numericValue) ? numericValue.toFixed(2).replace('.', ',') : '';
  }
  function downloadCsv(filename, headers, rows) {
    const csv = [headers, ...rows].map(row => row.map(csvValue).join(';')).join('\r\n');
    const blob = new Blob(['\ufeff', csv], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const anchor = document.createElement('a');
    anchor.href = url;
    anchor.download = filename;
    document.body.appendChild(anchor);
    anchor.click();
    anchor.remove();
    URL.revokeObjectURL(url);
    toast('Arquivo CSV preparado para download.');
  }

  function exportTargets() {
    const rows = getTargetRows().map(row => [row.agency, csvNumber(row.base), csvNumber(row.target), csvNumber(row.increase), csvNumber(row.annualTarget), csvNumber(state.growth)]);
    downloadCsv('simulacao-metas-assessorias.csv', ['Assessoria', 'Base mensal', 'Meta mensal', 'Aumento necessario', 'Projecao anual', 'Crescimento percentual'], rows);
  }

  function exportCommissions() {
    const rows = getTargetRows().map(row => {
      const simulated = calculateCommission(row.target, getRulesForAgency(row.agency));
      return [row.agency, csvNumber(row.target), csvNumber(row.avgCommission), csvNumber(simulated), csvNumber(simulated - row.avgCommission), csvNumber(row.target - simulated)];
    });
    downloadCsv('comparativo-comissoes-assessorias.csv', ['Assessoria', 'Meta bruta', 'Comissao historica media', 'Comissao simulada', 'Variacao de custo', 'Liquido apos comissao'], rows);
  }

  function exportHistory() {
    const rows = getFilteredRows().map(row => [row.agency, row.monthLabel + (row.partial ? ' (PARCIAL)' : ''), ...[
      row.collection, row.recovered, row.active, row.canceled, row.commission,
      Number.isFinite(row.conversion) ? row.conversion * 100 : null,
      Number.isFinite(row.costRatio) ? row.costRatio * 100 : null
    ].map(csvNumber)]);
    downloadCsv('historico-assessorias-filtrado.csv', ['Assessoria', 'Mes', 'Cobranca', 'Recuperado bruto', 'Recuperado ativo', 'Recuperado cancelado', 'Comissao', 'Conversao percentual', 'Custo percentual'], rows);
  }

  function syncGrowth(value, source) {
    const normalizedValue = Math.min(300, Math.max(0, Number(value) || 0));
    state.growth = normalizedValue;
    $('#growthNumber').value = String(normalizedValue);
    $('#growthSlider').value = String(Math.min(100, normalizedValue));
    if (source === 'slider' && normalizedValue > 100) $('#growthSlider').value = '100';
    if (state.rows.length) {
      renderTargets();
      renderCommissionComparison(getTargetRows());
    }
  }

  function handleRuleInput(event) {
    const target = event.target;
    const field = target.dataset.tierField;
    if (!field) return;
    const index = Number(target.dataset.tierIndex);
    const rules = ensureEditableRules();
    const tier = rules.tiers[index];
    if (!tier) return;
    const rawValue = Math.max(0, Number(target.value) || 0);
    const value = field === 'rate' ? Math.min(100, rawValue) : rawValue;
    tier[field] = value;
    rules.tiers.sort((a, b) => a.min - b.min);
    renderRuleEditor();
    renderCommissionComparison(getTargetRows());
  }

  function resetActiveRules() {
    const rate = state.ruleScope === 'ALL' ? getGlobalBaselineRate() * 100 : getBaselineRate(state.ruleScope) * 100;
    const initial = { mode: 'progressive', tiers: [{ min: 0, rate }] };
    if (state.ruleScope === 'ALL') state.globalRules = initial;
    else state.rulesByAgency[state.ruleScope] = initial;
    renderRuleEditor();
    renderCommissionComparison(getTargetRows());
    toast('Tabela restaurada à taxa efetiva histórica de referência.');
  }

  function addTier() {
    const rules = ensureEditableRules();
    const sorted = [...rules.tiers].sort((a, b) => a.min - b.min);
    const last = sorted[sorted.length - 1] || { min: 0, rate: 0 };
    rules.tiers.push({ min: Math.max(100000, last.min + 100000), rate: Math.min(100, last.rate + 0.5) });
    renderRuleEditor();
    renderCommissionComparison(getTargetRows());
  }

  function bindEvents() {
    const input = $('#fileInput');
    ['importTopBtn', 'importMainBtn', 'emptyImportBtn'].forEach(id => $(`#${id}`).addEventListener('click', () => input.click()));
    input.addEventListener('change', event => {
      const file = event.target.files && event.target.files[0];
      importWorkbook(file);
      input.value = '';
    });

    $('#clearDataBtn').addEventListener('click', () => {
      state.rows = [];
      state.policy = [];
      state.fileName = '';
      state.agencies = [];
      state.rulesByAgency = {};
      $('#workspace').classList.add('hidden');
      $('#emptyState').classList.remove('hidden');
      $('#clearDataBtn').classList.add('hidden');
      $('#statusFile').textContent = '';
      setStatus('Dados removidos da memória desta sessão. Importe um histórico para continuar.');
      toast('Dados removidos da sessão.');
    });

    $('#agencyFilter').addEventListener('change', event => { state.agencyFilter = event.target.value; renderAll(); });
    $('#startMonth').addEventListener('change', event => {
      state.startMonth = Number(event.target.value);
      if (state.startMonth > state.endMonth) { state.endMonth = state.startMonth; $('#endMonth').value = String(state.endMonth); }
      renderAll();
    });
    $('#endMonth').addEventListener('change', event => {
      state.endMonth = Number(event.target.value);
      if (state.endMonth < state.startMonth) { state.startMonth = state.endMonth; $('#startMonth').value = String(state.startMonth); }
      renderAll();
    });
    $('#includePartial').addEventListener('change', event => { state.includePartial = event.target.checked; renderAll(); });

    document.querySelectorAll('.tab-button').forEach(button => button.addEventListener('click', () => activateTab(button.dataset.tab)));
    $('#goTargetsBtn').addEventListener('click', () => activateTab('targets'));
    $('#baselineMode').addEventListener('change', event => { state.baselineMode = event.target.value; renderTargets(); renderCommissionComparison(getTargetRows()); });
    $('#targetBasis').addEventListener('change', event => { state.targetBasis = event.target.value; renderTargets(); renderCommissionComparison(getTargetRows()); });
    $('#growthNumber').addEventListener('input', event => syncGrowth(event.target.value, 'number'));
    $('#growthSlider').addEventListener('input', event => syncGrowth(event.target.value, 'slider'));

    $('#ruleScope').addEventListener('change', event => {
      state.ruleScope = event.target.value;
      renderRuleEditor();
      renderCommissionComparison(getTargetRows());
    });
    $('#commissionMode').addEventListener('change', event => {
      const rules = ensureEditableRules();
      rules.mode = event.target.value;
      renderRuleEditor();
      renderCommissionComparison(getTargetRows());
    });
    $('#tiersTableBody').addEventListener('change', handleRuleInput);
    $('#tiersTableBody').addEventListener('click', event => {
      const button = event.target.closest('[data-delete-tier]');
      if (!button) return;
      const rules = ensureEditableRules();
      if (rules.tiers.length <= 1) return;
      rules.tiers.splice(Number(button.dataset.deleteTier), 1);
      renderRuleEditor();
      renderCommissionComparison(getTargetRows());
    });
    $('#addTierBtn').addEventListener('click', addTier);
    $('#resetRulesBtn').addEventListener('click', resetActiveRules);
    $('#historySearch').addEventListener('input', renderHistory);

    $('#exportTargetsBtn').addEventListener('click', exportTargets);
    $('#exportCommissionBtn').addEventListener('click', exportCommissions);
    $('#exportHistoryBtn').addEventListener('click', exportHistory);

    // Drag-and-drop na área principal, além do seletor tradicional.
    const dropTargets = [$('#emptyState'), $('.page-heading')];
    dropTargets.forEach(target => {
      target.addEventListener('dragover', event => { event.preventDefault(); target.classList.add('drag-over'); });
      target.addEventListener('dragleave', () => target.classList.remove('drag-over'));
      target.addEventListener('drop', event => {
        event.preventDefault();
        target.classList.remove('drag-over');
        const file = event.dataTransfer.files && event.dataTransfer.files[0];
        if (file) importWorkbook(file);
      });
    });
  }

  bindEvents();
  tryLoadLocalHistory();
})();
