/*
 * Central de Recuperação — módulo Assessorias.
 * O Excel e as simulações são processados somente na memória do navegador.
 */
(() => {
  'use strict';

  const MONTHS = [
    { n: 1, short: 'JAN', name: 'Janeiro' }, { n: 2, short: 'FEV', name: 'Fevereiro' },
    { n: 3, short: 'MAR', name: 'Março' }, { n: 4, short: 'ABR', name: 'Abril' },
    { n: 5, short: 'MAI', name: 'Maio' }, { n: 6, short: 'JUN', name: 'Junho' },
    { n: 7, short: 'JUL', name: 'Julho' }, { n: 8, short: 'AGO', name: 'Agosto' },
    { n: 9, short: 'SET', name: 'Setembro' }, { n: 10, short: 'OUT', name: 'Outubro' },
    { n: 11, short: 'NOV', name: 'Novembro' }, { n: 12, short: 'DEZ', name: 'Dezembro' }
  ];
  const MONTH_LOOKUP = Object.fromEntries(MONTHS.map(month => [month.short, month.n]));
  const $ = selector => document.querySelector(selector);

  const BASIS_LABELS = {
    recovered: 'Recuperado bruto', active: 'Recuperado ativo',
    canceled: 'Recuperado cancelado', combined: 'Ativo + cancelado', target: 'Mesmo indicador da meta'
  };
  const MODE_LABELS = {
    history_growth: 'Crescimento sobre o histórico',
    manual: 'Valor manual por assessoria',
    monthly_growth: 'Aumento percentual a cada mês'
  };

  const state = {
    rows: [], policy: [], fileName: '', year: null, agencies: [], hasHistory: false,
    startMonth: 1, endMonth: 9, includePartial: false, agencyFilter: 'ALL',
    baselineMode: 'average', targetBasis: 'recovered', targetMode: 'history_growth',
    growth: 10, monthlyGrowthRate: 2, monthlyInitialGrowth: 0, projectionMonths: 12,
    manualTargets: {}, manualDefault: null,
    globalRules: null, rulesByAgency: {}, ruleScope: 'ALL', activeTab: 'overview'
  };

  const money = (value, compact = false) => {
    if (value === null || value === undefined || value === '' || !Number.isFinite(Number(value))) return '—';
    const n = Number(value);
    if (compact && Math.abs(n) >= 1000000) return n.toLocaleString('pt-BR', { style: 'currency', currency: 'BRL', notation: 'compact', maximumFractionDigits: 1 });
    return n.toLocaleString('pt-BR', { style: 'currency', currency: 'BRL', maximumFractionDigits: 2 });
  };
  const number = (value, digits = 0) => (Number(value) || 0).toLocaleString('pt-BR', { minimumFractionDigits: digits, maximumFractionDigits: digits });
  const percent = (value, digits = 2) => Number.isFinite(Number(value)) ? `${number(Number(value) * 100, digits)}%` : '—';
  const percentPoints = (value, digits = 2) => `${number(Number(value) || 0, digits)}%`;
  const esc = value => String(value ?? '').replace(/[&<>"']/g, char => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[char]));
  const normalized = value => String(value ?? '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').trim().toUpperCase();
  const sum = values => values.reduce((total, value) => total + (Number.isFinite(value) ? value : 0), 0);
  const sumOrNull = values => values.some(Number.isFinite) ? sum(values) : null;
  const clamp = (value, min, max) => Math.min(max, Math.max(min, Number(value) || 0));

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
  const parseNumeric = value => parseOptionalNumeric(value) ?? 0;

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
    for (let index = 0; index < Math.min(rows.length, 12); index += 1) {
      const candidate = rows[index] || [];
      const agency = findHeaderIndex(candidate, ['ASSESSORIA']);
      const month = findHeaderIndex(candidate, ['MES']);
      const recovered = findHeaderIndex(candidate, ['RECUPERADO']);
      if (agency >= 0 && month >= 0 && recovered >= 0) {
        headerIndex = index;
        columns = {
          agency, month,
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
    for (let index = headerIndex + 1; index < rows.length; index += 1) {
      const row = rows[index] || [];
      const agencyRaw = String(row[columns.agency] ?? '').trim();
      const monthRaw = String(row[columns.month] ?? '').trim();
      const agencyKey = normalized(agencyRaw);
      const monthKey = normalized(monthRaw.split(/[\s(]/)[0]);
      const month = MONTH_LOOKUP[monthKey];
      if (!agencyRaw || agencyKey === '0' || agencyKey.includes('TOTAL') || normalized(monthRaw).includes('TOTAL') || !month) continue;
      const get = col => col >= 0 ? row[col] : null;
      data.push({
        agency: agencyRaw, month, monthLabel: MONTHS[month - 1].short, monthName: MONTHS[month - 1].name,
        partial: normalized(monthRaw).includes('PARCIAL'),
        collection: parseNumeric(get(columns.collection)), recovered: parseNumeric(get(columns.recovered)),
        active: parseOptionalNumeric(get(columns.active)), canceled: parseOptionalNumeric(get(columns.canceled)),
        commission: parseOptionalNumeric(get(columns.commission)), conversion: parseOptionalNumeric(get(columns.conversion)),
        costRatio: parseOptionalNumeric(get(columns.costRatio)), sourceRow: index + 1
      });
    }
    if (!data.length) throw new Error('A aba de histórico foi localizada, mas não há registros mensais válidos para as assessorias.');
    return data;
  }

  function parsePolicyRows(rows) {
    let headerRow = -1;
    let startCol = -1;
    for (let index = 0; index < rows.length; index += 1) {
      const col = (rows[index] || []).findIndex(cell => normalized(cell).includes('FAIXA DE ATRASO'));
      if (col >= 0) { headerRow = index; startCol = col; break; }
    }
    if (headerRow < 0) return [];
    const result = [];
    for (let index = headerRow + 1; index < rows.length; index += 1) {
      const row = rows[index] || [];
      const label = String(row[startCol] ?? '').trim();
      if (!label) break;
      if (typeof row[startCol + 1] === 'string' && normalized(row[startCol + 1]).includes('COMISSAO')) break;
      result.push({ range: label, b2c: parseNumeric(row[startCol + 1]), b2b: parseNumeric(row[startCol + 2]) });
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
      if (!groups.has(row.agency)) groups.set(row.agency, { agency: row.agency, records: [], totalCollection: 0, totalRecovered: 0, totalActive: 0, totalCanceled: 0, totalCommission: 0, activeCount: 0, canceledCount: 0, commissionCount: 0, commissionBaseRecovered: 0 });
      const group = groups.get(row.agency);
      group.records.push(row);
      group.totalCollection += row.collection;
      group.totalRecovered += row.recovered;
      if (Number.isFinite(row.active)) { group.totalActive += row.active; group.activeCount += 1; }
      if (Number.isFinite(row.canceled)) { group.totalCanceled += row.canceled; group.canceledCount += 1; }
      if (Number.isFinite(row.commission)) { group.totalCommission += row.commission; group.commissionCount += 1; group.commissionBaseRecovered += row.recovered; }
    });
    return [...groups.values()].map(group => {
      group.records.sort((a, b) => a.month - b.month);
      group.monthCount = group.records.length;
      group.avgRecovered = group.monthCount ? group.totalRecovered / group.monthCount : null;
      group.avgActive = group.activeCount ? group.totalActive / group.activeCount : null;
      group.avgCanceled = group.canceledCount ? group.totalCanceled / group.canceledCount : null;
      group.avgCommission = group.commissionCount ? group.totalCommission / group.commissionCount : null;
      group.effectiveRate = group.commissionBaseRecovered > 0 ? group.totalCommission / group.commissionBaseRecovered : null;
      return group;
    }).sort((a, b) => a.agency.localeCompare(b.agency, 'pt-BR'));
  }

  function getScenarioGroups() {
    const existing = groupRows(getFilteredRows());
    const byAgency = new Map(existing.map(group => [group.agency, group]));
    const names = state.agencyFilter === 'ALL' ? state.agencies : [state.agencyFilter];
    return names.map(agency => byAgency.get(agency) || {
      agency, records: [], totalCollection: 0, totalRecovered: 0, totalActive: 0,
      totalCanceled: 0, totalCommission: 0, activeCount: 0, canceledCount: 0,
      commissionCount: 0, commissionBaseRecovered: 0, monthCount: 0,
      avgRecovered: null, avgActive: null, avgCanceled: null, avgCommission: null, effectiveRate: null
    });
  }

  function manualTargetKey(agency, basis = state.targetBasis) {
    return `${basis}::${agency}`;
  }

  function metricFromRow(row, basis) {
    if (basis === 'recovered') return Number.isFinite(row.recovered) ? row.recovered : null;
    if (basis === 'active') return Number.isFinite(row.active) ? row.active : null;
    if (basis === 'canceled') return Number.isFinite(row.canceled) ? row.canceled : null;
    if (basis === 'combined') return Number.isFinite(row.active) && Number.isFinite(row.canceled) ? row.active + row.canceled : null;
    return null;
  }

  function getBaseline(group, basis, mode = state.baselineMode) {
    const eligible = group.records.map(row => ({ row, value: metricFromRow(row, basis) })).filter(item => Number.isFinite(item.value));
    if (!eligible.length) return { value: null, count: 0, latestMonth: null, total: 0 };
    if (mode === 'last') {
      const latest = eligible[eligible.length - 1];
      return { value: latest.value, count: eligible.length, latestMonth: latest.row.monthLabel, total: eligible.reduce((acc, item) => acc + item.value, 0) };
    }
    const total = eligible.reduce((acc, item) => acc + item.value, 0);
    return { value: total / eligible.length, count: eligible.length, latestMonth: eligible[eligible.length - 1].row.monthLabel, total };
  }

  function makeSchedule(start, months, rate = 0) {
    if (!Number.isFinite(start)) return Array(months).fill(null);
    const schedule = [];
    let current = start;
    for (let month = 0; month < months; month += 1) {
      schedule.push(current);
      current *= 1 + rate / 100;
    }
    return schedule;
  }

  function getTargetRows() {
    return getScenarioGroups().map(group => {
      const baseline = getBaseline(group, state.targetBasis);
      let start = baseline.value;
      const custom = state.manualTargets[manualTargetKey(group.agency)];
      if (state.targetMode === 'history_growth') {
        if (Number.isFinite(start)) start *= 1 + state.growth / 100;
      } else if (state.targetMode === 'monthly_growth') {
        if (Number.isFinite(start)) start *= 1 + state.monthlyInitialGrowth / 100;
        if (Number.isFinite(custom)) start = custom;
      } else if (Number.isFinite(custom)) {
        start = custom;
      }
      const rate = state.targetMode === 'monthly_growth' ? state.monthlyGrowthRate : 0;
      const schedule = makeSchedule(start, state.projectionMonths, rate);
      const finalTarget = schedule.length ? schedule[schedule.length - 1] : null;
      const totalProjection = schedule.every(Number.isFinite) ? sum(schedule) : null;
      const initialIncrease = Number.isFinite(start) && Number.isFinite(baseline.value) ? start - baseline.value : null;
      return { ...group, baseline: baseline.value, baselineCount: baseline.count, baselineMonth: baseline.latestMonth, targetStart: start, targetFinal: finalTarget, target: start, schedule, totalProjection, increase: initialIncrease };
    });
  }

  function projectedMetricForBasis(targetRow, basis, monthIndex = 0) {
    const actualBasis = basis === 'target' ? state.targetBasis : basis;
    if (actualBasis === state.targetBasis) return targetRow.schedule[monthIndex] ?? null;
    const baseline = getBaseline(targetRow, actualBasis);
    if (!Number.isFinite(baseline.value)) return null;
    if (state.targetMode === 'manual') return baseline.value;
    const initialGrowth = state.targetMode === 'monthly_growth' ? state.monthlyInitialGrowth : state.growth;
    const start = baseline.value * (1 + initialGrowth / 100);
    return start * Math.pow(1 + (state.targetMode === 'monthly_growth' ? state.monthlyGrowthRate : 0) / 100, monthIndex);
  }

  function getBaselineRate(agency) {
    const source = state.rows.filter(row => row.agency === agency && row.month >= state.startMonth && row.month <= state.endMonth && !row.partial && Number.isFinite(row.commission));
    const recovery = sum(source.map(row => row.recovered));
    const commission = sum(source.map(row => row.commission));
    return recovery > 0 ? commission / recovery : 0;
  }

  function getGlobalBaselineRate() {
    const source = state.rows.filter(row => row.month >= state.startMonth && row.month <= state.endMonth && !row.partial && Number.isFinite(row.commission));
    const recovery = sum(source.map(row => row.recovered));
    const commission = sum(source.map(row => row.commission));
    return recovery > 0 ? commission / recovery : 0;
  }

  function getGlobalMonthlyRecovered() {
    return groupRows(getFilteredRows(true)).reduce((total, group) => total + (group.avgRecovered || 0), 0);
  }

  function makeDefaultRules(rate, threshold) {
    const ratePoints = rate * 100;
    return {
      mode: 'progressive', tiers: [{ min: 0, rate: ratePoints }],
      commissionBasis: 'target',
      trigger: { basis: 'target', threshold: Math.max(0, threshold || 0), baseRate: ratePoints, bonusRate: 1 }
    };
  }
  function initializeRules() {
    state.globalRules = makeDefaultRules(getGlobalBaselineRate(), getGlobalMonthlyRecovered());
    state.rulesByAgency = {};
    state.ruleScope = 'ALL';
  }
  function cloneRules(rules) {
    return { mode: rules.mode, tiers: rules.tiers.map(tier => ({ min: tier.min, rate: tier.rate })), commissionBasis: rules.commissionBasis || 'target', trigger: { ...rules.trigger } };
  }
  function getEditorRules() {
    if (state.ruleScope !== 'ALL' && state.rulesByAgency[state.ruleScope]) return state.rulesByAgency[state.ruleScope];
    return state.globalRules;
  }
  function getRulesForAgency(agency) { return state.rulesByAgency[agency] || state.globalRules; }
  function ensureEditableRules() {
    if (state.ruleScope !== 'ALL' && !state.rulesByAgency[state.ruleScope]) state.rulesByAgency[state.ruleScope] = cloneRules(state.globalRules);
    return getEditorRules();
  }

  function calculateCommission(baseValue, rules, triggerValue = null) {
    if (!Number.isFinite(baseValue)) return { amount: null, rate: null, reached: false };
    const value = Math.max(0, Number(baseValue) || 0);
    if (rules.mode === 'threshold_bonus') {
      const trigger = rules.trigger || {};
      const threshold = Number(trigger.threshold) || 0;
      const reached = threshold > 0 && Number.isFinite(triggerValue) && triggerValue >= threshold;
      const rate = (Number(trigger.baseRate) || 0) + (reached ? (Number(trigger.bonusRate) || 0) : 0);
      return { amount: value * rate / 100, rate, reached };
    }
    const tiers = [...rules.tiers].sort((a, b) => a.min - b.min);
    if (!tiers.length || value <= 0) return { amount: 0, rate: 0, reached: false };
    if (rules.mode === 'attained') {
      let selected = tiers[0];
      tiers.forEach(tier => { if (value >= tier.min) selected = tier; });
      return { amount: value * selected.rate / 100, rate: selected.rate, reached: false };
    }
    let total = 0;
    for (let index = 0; index < tiers.length; index += 1) {
      const tier = tiers[index];
      const next = tiers[index + 1];
      const low = Math.max(0, tier.min);
      const high = next ? Math.min(value, next.min) : value;
      total += Math.max(0, high - low) * tier.rate / 100;
      if (value <= (next ? next.min : value)) break;
    }
    return { amount: total, rate: value > 0 ? total / value * 100 : 0, reached: false };
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
    toastTimer = window.setTimeout(() => node.className = 'toast', 3200);
  }

  function refreshAgencyOptions() {
    const selectedAgency = state.agencyFilter;
    const selectedScope = state.ruleScope;
    const options = ['<option value="ALL">Todas as assessorias</option>', ...state.agencies.map(name => `<option value="${esc(name)}">${esc(name)}</option>`)].join('');
    $('#agencyFilter').innerHTML = options;
    $('#ruleScope').innerHTML = options;
    $('#agencyFilter').value = state.agencies.includes(selectedAgency) ? selectedAgency : 'ALL';
    $('#ruleScope').value = selectedScope === 'ALL' || state.agencies.includes(selectedScope) ? selectedScope : 'ALL';
    state.agencyFilter = $('#agencyFilter').value;
    state.ruleScope = $('#ruleScope').value;
  }

  function populateFilters() {
    refreshAgencyOptions();
    const months = MONTHS.map(month => `<option value="${month.n}">${month.short} · ${month.name}</option>`).join('');
    $('#startMonth').innerHTML = months;
    $('#endMonth').innerHTML = months;
    const observed = state.rows.map(row => row.month);
    state.startMonth = observed.length ? Math.min(...observed) : 1;
    state.endMonth = observed.length ? Math.max(...observed) : 12;
    $('#startMonth').value = String(state.startMonth);
    $('#endMonth').value = String(state.endMonth);
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
      state.rows = parsedRows;
      state.policy = parsePolicyRows(historyArray);
      state.hasHistory = true;
      state.fileName = fileName;
      state.year = inferYear(fileName);
      state.agencies = [...new Set(parsedRows.map(row => row.agency))].sort((a, b) => a.localeCompare(b, 'pt-BR'));
      state.manualTargets = {};
      state.manualDefault = null;
      populateFilters();
      initializeRules();
      $('#emptyState').classList.add('hidden');
      $('#workspace').classList.remove('hidden');
      $('#clearDataBtn').classList.remove('hidden');
      $('#statusFile').textContent = `${fileName} · ${parsedRows.length} registros`;
      const partialCount = parsedRows.filter(row => row.partial).length;
      const partialMessage = partialCount ? ` ${partialCount} registro(s) parcial(is) foram identificados e estão excluídos por padrão.` : '';
      setStatus(`Histórico importado: ${state.agencies.length} assessorias e ${parsedRows.length} registros mensais.${partialMessage}`, 'success');
      const first = Math.min(...parsedRows.map(row => row.month));
      const last = Math.max(...parsedRows.map(row => row.month));
      const yearText = state.year ? ` · referência ${state.year} (inferida do nome do arquivo)` : ' · ano não informado no nome do arquivo';
      $('#historyFilters').classList.remove('hidden');
      $('#dataContext').innerHTML = `<strong>${esc(fileName)}</strong> · ${state.agencies.length} assessorias · ${MONTHS[first - 1].short} a ${MONTHS[last - 1].short}${yearText} · ${partialCount} linha(s) parcial(is)`;
      $('#policyFoot').textContent = state.policy.length ? `A planilha forneceu ${state.policy.length} faixas de atraso. Os percentuais são exibidos como referência da política registrada no arquivo, sem aplicação automática nos novos cenários.` : 'Não foi localizada uma tabela de política de comissão por atraso na aba de histórico.';
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
    try { await processWorkbookBuffer(await file.arrayBuffer(), file.name); }
    catch (error) {
      setStatus('Não foi possível ler o arquivo selecionado.', 'error');
      toast(error.message || 'Falha ao abrir o arquivo.', 'error');
    }
  }

  async function tryLoadLocalHistory() {
    const candidates = ['./dados/historico.xlsx', './dados/analise_assessorias_14-10-26(1).xlsx'];
    for (const path of candidates) {
      try {
        const response = await fetch(path, { cache: 'no-store' });
        if (!response.ok) continue;
        await processWorkbookBuffer(await response.arrayBuffer(), path.replace('./', ''));
        return;
      } catch (_) { return; }
    }
  }

  function updateMetrics(groups) {
    if (!state.hasHistory) {
      ['metricRecovered', 'metricMonthly', 'metricCommission', 'metricCost'].forEach(id => { $(`#${id}`).textContent = '—'; });
      $('#metricRecoveredFoot').textContent = 'Sem histórico importado';
      $('#metricCommission').nextElementSibling.textContent = 'Sem dados históricos';
      $('#metricCost').nextElementSibling.textContent = 'Sem dados históricos';
      return;
    }
    const rows = getFilteredRows();
    const recovered = sum(rows.map(row => row.recovered));
    const commissionRows = rows.filter(row => Number.isFinite(row.commission));
    const commission = sum(commissionRows.map(row => row.commission));
    const commissionBase = sum(commissionRows.map(row => row.recovered));
    const averageSum = sum(groups.map(group => group.avgRecovered));
    $('#metricRecovered').textContent = money(recovered, true);
    $('#metricRecoveredFoot').textContent = `${rows.length} registro(s) no recorte`;
    $('#metricMonthly').textContent = money(averageSum, true);
    $('#metricCommission').textContent = money(commission, true);
    $('#metricCommission').nextElementSibling.textContent = `${commissionRows.length} registro(s) com comissão apurada`;
    $('#metricCost').textContent = percent(commissionBase > 0 ? commission / commissionBase : null, 2);
    $('#metricCost').nextElementSibling.textContent = 'Somente meses com comissão informada';
  }

  function renderOverview(groups) {
    const rows = getFilteredRows();
    const total = sum(groups.map(group => group.totalRecovered));
    const ranking = [...groups].sort((a, b) => b.totalRecovered - a.totalRecovered);
    const max = ranking.length ? Math.max(...ranking.map(group => group.totalRecovered), 1) : 1;
    $('#overviewPeriodChip').textContent = state.hasHistory ? `${MONTHS[state.startMonth - 1].short} – ${MONTHS[state.endMonth - 1].short}${state.includePartial ? ' · parciais incluídos' : ' · parciais excluídos'}` : 'Sem histórico';
    if (!state.hasHistory) {
      $('#rankingChart').innerHTML = '<p class="muted-empty">Este cenário foi criado sem histórico. Informe metas e regras manualmente na aba “Metas e evolução”.</p>';
      $('#insightContent').innerHTML = '<div class="insight-message">Não há dados históricos carregados. O simulador está livre para testar valores manuais; importe um Excel a qualquer momento para habilitar as comparações históricas.</div>';
      $('#overviewTableBody').innerHTML = groups.map(group => `<tr><td>${esc(group.agency)}</td><td class="numeric">—</td><td class="numeric">—</td><td class="numeric">—</td><td class="numeric">—</td><td class="numeric">—</td></tr>`).join('');
      return;
    }
    $('#rankingChart').innerHTML = ranking.length ? ranking.map(group => `
      <div class="ranking-row"><div class="ranking-name">${esc(group.agency)}</div><div class="ranking-track" aria-label="${esc(group.agency)}: ${money(group.totalRecovered)}"><div class="ranking-fill" style="width:${Math.max(1, group.totalRecovered / max * 100)}%"></div></div><div class="ranking-value">${money(group.totalRecovered, true)}</div></div>`).join('') : '<p class="muted-empty">Sem registros neste recorte.</p>';
    const leader = ranking[0];
    const lowestCost = [...groups].filter(group => Number.isFinite(group.effectiveRate)).sort((a, b) => a.effectiveRate - b.effectiveRate)[0];
    const knownCommission = rows.filter(row => Number.isFinite(row.commission));
    const commissionTotal = sum(knownCommission.map(row => row.commission));
    const commissionRecovery = sum(knownCommission.map(row => row.recovered));
    const avgCost = commissionRecovery > 0 ? commissionTotal / commissionRecovery : null;
    $('#insightContent').innerHTML = `
      <div class="insight-item"><span>Maior recuperação</span><strong>${leader ? esc(leader.agency) : '—'}</strong><small>${leader ? money(leader.totalRecovered, true) : 'Sem dados'}</small></div>
      <div class="insight-item"><span>Menor custo efetivo</span><strong>${lowestCost ? esc(lowestCost.agency) : '—'}</strong><small>${lowestCost ? percent(lowestCost.effectiveRate) + ' sobre o recuperado' : 'Sem dados'}</small></div>
      <div class="insight-message">${groups.length ? `A recuperação consolidada foi de <strong>${money(total, true)}</strong>, com custo de comissão ponderado de <strong>${percent(avgCost)}</strong> nos registros em que a comissão foi informada. As médias individuais consideram somente os meses com registro para cada assessoria.` : 'Não há dados suficientes para gerar uma leitura do período.'}</div>`;
    $('#overviewTableBody').innerHTML = groups.length ? groups.map(group => `
      <tr><td>${esc(group.agency)}</td><td class="numeric">${group.monthCount}</td><td class="numeric">${money(group.totalRecovered)}</td><td class="numeric">${money(group.avgRecovered)}</td><td class="numeric">${money(group.avgCommission)}</td><td class="numeric">${percent(group.effectiveRate)}</td></tr>`).join('') : '<tr><td colspan="6">Sem dados para os filtros selecionados.</td></tr>';
  }

  function setModeControlVisibility() {
    $('#historyGrowthControls').classList.toggle('hidden', state.targetMode !== 'history_growth');
    $('#monthlyGrowthControls').classList.toggle('hidden', state.targetMode !== 'monthly_growth');
    $('#manualControls').classList.toggle('hidden', state.targetMode !== 'manual');
    $('#baselineMode').disabled = state.targetMode === 'manual';
    $('#targetModeNote').textContent = state.targetMode === 'manual'
      ? 'As metas são livres; a referência histórica permanece visível apenas para comparação.'
      : state.targetMode === 'monthly_growth'
        ? (!state.hasHistory ? 'Sem histórico: informe o valor inicial (mês 1) de cada assessoria na tabela; o crescimento composto será aplicado a partir dele.' : 'A meta começa pela referência ajustada e cresce de forma composta mês a mês.')
        : 'A meta mensal é a referência histórica acrescida do percentual inicial.';
  }

  function renderTargets() {
    $('#targetMode').value = state.targetMode;
    $('#targetBasis').value = state.targetBasis;
    $('#baselineMode').value = state.baselineMode;
    setModeControlVisibility();
    const targets = getTargetRows();
    const totalStart = sumOrNull(targets.map(row => row.targetStart));
    const totalFinal = sumOrNull(targets.map(row => row.targetFinal));
    const totalProjection = sumOrNull(targets.map(row => row.totalProjection));
    const baseTotal = sumOrNull(targets.map(row => row.baseline));
    const increase = totalStart - baseTotal;
    const validBaselines = targets.filter(row => Number.isFinite(row.baseline));
    const covered = validBaselines.reduce((acc, row) => acc + row.baselineCount, 0);
    const available = targets.reduce((acc, row) => acc + row.monthCount, 0);
    const enteredTargets = targets.filter(row => Number.isFinite(row.targetStart)).length;
    const basisLabel = BASIS_LABELS[state.targetBasis];
    const modeLabel = MODE_LABELS[state.targetMode];
    $('#targetBasisHint').textContent = state.targetBasis === 'combined'
      ? 'A soma só é calculada quando ativo e cancelado estão preenchidos no mesmo mês.'
      : 'Meses sem valor para este indicador são desconsiderados da média, não tratados como zero.';
    $('#targetCoverageNote').textContent = !state.hasHistory
      ? `${enteredTargets} de ${targets.length} assessoria(s) com meta inicial informada.`
      : `${covered} mês(es) com indicador disponível entre ${available} registro(s) no recorte.`;
    $('#targetHighlight').innerHTML = `
      <div class="target-intro"><span>${esc(basisLabel)} · ${esc(modeLabel)}</span><strong>${state.targetMode === 'monthly_growth' ? `${number(state.monthlyGrowthRate, 2)}% composto ao mês` : state.targetMode === 'history_growth' ? `${number(state.growth, 1)}% sobre a referência` : 'Valores definidos manualmente'}</strong></div>
      <div class="target-kpi"><span>Meta consolidada · mês 1</span><strong>${money(totalStart, true)}</strong></div>
      <div class="target-kpi"><span>Meta no mês final</span><strong>${money(totalFinal, true)}</strong></div>
      <div class="target-kpi"><span>Total projetado · ${state.projectionMonths} meses</span><strong>${money(totalProjection, true)}</strong></div>`;

    $('#targetsTableBody').innerHTML = targets.length ? targets.map(row => `
      <tr>
        <td><strong>${esc(row.agency)}</strong><small class="cell-subtext">${Number.isFinite(row.baseline) ? `${row.baselineCount} mês(es) com dado` : 'Sem histórico para este indicador'}</small></td>
        <td class="numeric">${money(row.baseline)}</td>
        <td class="numeric">${['manual', 'monthly_growth'].includes(state.targetMode) ? `<label class="visually-hidden" for="manual-${esc(row.agency)}">Meta inicial de ${esc(row.agency)}</label><input class="table-number-input" id="manual-${esc(row.agency)}" type="number" min="0" step="100" value="${Number.isFinite(row.targetStart) ? row.targetStart.toFixed(2) : ''}" data-manual-agency="${esc(row.agency)}" placeholder="Informe M1">` : `<strong>${money(row.targetStart)}</strong>`}</td>
        <td class="numeric">${money(row.targetFinal)}</td><td class="numeric">${money(row.totalProjection)}</td><td class="numeric ${Number.isFinite(row.increase) && row.increase >= 0 ? 'positive' : 'negative'}">${money(row.increase)}</td>
      </tr>`).join('') : '<tr><td colspan="6">Sem dados para os filtros selecionados.</td></tr>';

    const first = state.targetMode === 'monthly_growth' ? state.monthlyInitialGrowth : state.growth;
    $('#targetTableNote').textContent = state.targetMode === 'manual'
      ? 'A meta manual é aplicada como valor mensal constante. As referências históricas são exibidas para comparação e não limitam os valores digitados.'
      : state.targetMode === 'monthly_growth'
        ? (!state.hasHistory ? `Informe a meta inicial no mês 1 para cada assessoria. A partir dela, o valor cresce ${number(state.monthlyGrowthRate, 2)}% sobre o mês anterior.` : `O mês 1 começa com ${number(first, 2)}% sobre a referência e os meses seguintes crescem ${number(state.monthlyGrowthRate, 2)}% sobre o mês anterior.`)
        : 'A meta é constante ao longo dos meses projetados; a projeção total corresponde à meta mensal multiplicada pelo número de meses.';

    $('#evolutionPeriodChip').textContent = `${state.projectionMonths} mês(es)`;
    $('#evolutionTableHead').innerHTML = `<th>Indicador</th>${Array.from({ length: state.projectionMonths }, (_, index) => `<th class="numeric">Mês ${index + 1}</th>`).join('')}<th class="numeric">Total</th>`;
    const monthlyTotals = Array.from({ length: state.projectionMonths }, (_, index) => sumOrNull(targets.map(row => row.schedule[index])));
    $('#evolutionTableBody').innerHTML = `<tr><td><strong>${esc(basisLabel)}</strong></td>${monthlyTotals.map(value => `<td class="numeric">${money(value)}</td>`).join('')}<td class="numeric"><strong>${money(totalProjection)}</strong></td></tr>`;
  }

  function renderRuleEditor() {
    const rules = getEditorRules();
    $('#commissionMode').value = rules.mode;
    $('#commissionBasis').value = rules.commissionBasis || 'target';
    const triggerMode = rules.mode === 'threshold_bonus';
    $('#tierSettings').classList.toggle('hidden', triggerMode);
    $('#triggerSettings').classList.toggle('hidden', !triggerMode);
    $('#addTierBtn').disabled = triggerMode;
    $('#modeExplainer').textContent = rules.mode === 'progressive'
      ? 'Cada percentual incide somente sobre a parcela do valor que estiver dentro daquela faixa.'
      : rules.mode === 'attained'
        ? 'O percentual da faixa mais alta alcançada incide sobre todo o valor usado como base de comissão.'
        : 'Ao alcançar o gatilho, o acréscimo percentual é aplicado sobre toda a base de comissão do período.';
    const sorted = [...rules.tiers].sort((a, b) => a.min - b.min);
    $('#tiersTableBody').innerHTML = sorted.map((tier, index) => `
      <tr><td><input type="number" min="0" step="1000" value="${Number(tier.min).toFixed(2)}" data-tier-index="${index}" data-tier-field="min" aria-label="Valor inicial da faixa ${index + 1}"></td>
      <td><input class="rate-input" type="number" min="0" max="100" step="0.05" value="${Number(tier.rate).toFixed(3)}" data-tier-index="${index}" data-tier-field="rate" aria-label="Percentual da faixa ${index + 1}"></td>
      <td><button class="tier-delete" type="button" data-delete-tier="${index}" aria-label="Remover faixa ${index + 1}" ${sorted.length === 1 ? 'disabled' : ''}>×</button></td></tr>`).join('');
    $('#triggerBasis').value = rules.trigger.basis || 'target';
    $('#triggerThreshold').value = Number(rules.trigger.threshold || 0).toFixed(2);
    $('#triggerBaseRate').value = Number(rules.trigger.baseRate || 0).toFixed(3);
    $('#triggerBonusRate').value = Number(rules.trigger.bonusRate || 0).toFixed(3);
    const refRate = state.ruleScope === 'ALL' ? getGlobalBaselineRate() : getBaselineRate(state.ruleScope);
    const label = state.ruleScope === 'ALL' ? 'Média ponderada histórica de todas as assessorias' : `Referência histórica de ${esc(state.ruleScope)}`;
    $('#referenceRate').innerHTML = `${label}: <strong>${percent(refRate)}</strong> de comissão efetiva sobre a recuperação bruta em meses completos. ${state.ruleScope !== 'ALL' && !state.rulesByAgency[state.ruleScope] ? 'A tabela individual ainda herda o modelo geral.' : ''}`;
    $('#triggerReference').innerHTML = `A taxa-base pode partir da referência histórica de <strong>${percent(refRate)}</strong>, mas é independente dela e pode ser alterada livremente.`;
  }

  function buildCommissionResults(targets) {
    return targets.map(row => {
      const rules = getRulesForAgency(row.agency);
      const commissionBasisValue = projectedMetricForBasis(row, rules.commissionBasis || 'target', 0);
      const triggerBasisValue = projectedMetricForBasis(row, rules.trigger.basis || 'target', 0);
      const calculation = calculateCommission(commissionBasisValue, rules, triggerBasisValue);
      const historical = row.avgCommission;
      const delta = Number.isFinite(calculation.amount) && Number.isFinite(historical) ? calculation.amount - historical : null;
      return {
        ...row, rules, commissionBasisValue, triggerBasisValue,
        simulated: calculation.amount, effectiveRate: calculation.rate, triggerReached: calculation.reached,
        historical, delta, net: Number.isFinite(commissionBasisValue) && Number.isFinite(calculation.amount) ? commissionBasisValue - calculation.amount : null
      };
    });
  }

  function renderCommissionComparison(targets) {
    const result = buildCommissionResults(targets);
    $('#commissionTableBody').innerHTML = result.length ? result.map(row => `
      <tr><td><strong>${esc(row.agency)}</strong></td><td class="numeric">${money(row.commissionBasisValue)}</td><td class="numeric">${row.rules.mode === 'threshold_bonus' ? `<span class="trigger-status ${row.triggerReached ? 'reached' : 'not-reached'}">${row.triggerReached ? 'Atingido' : 'Não atingido'}</span><small class="cell-subtext">${money(row.triggerBasisValue)} / ${money(row.rules.trigger.threshold)}</small>` : 'Não aplicável'}</td>
      <td class="numeric">${percentPoints(row.effectiveRate)}</td><td class="numeric"><strong>${money(row.simulated)}</strong></td><td class="numeric ${Number.isFinite(row.delta) && row.delta > 0 ? 'negative' : 'positive'}">${Number.isFinite(row.delta) && row.delta > 0 ? '+' : ''}${money(row.delta)}</td><td class="numeric">${money(row.net)}</td></tr>`).join('') : '<tr><td colspan="7">Sem dados para os filtros selecionados.</td></tr>';

    const selected = state.ruleScope === 'ALL' ? result : result.filter(row => row.agency === state.ruleScope);
    const base = sumOrNull(selected.map(row => row.commissionBasisValue));
    const historical = sumOrNull(selected.map(row => row.historical));
    const simulated = sumOrNull(selected.map(row => row.simulated));
    const delta = Number.isFinite(simulated) && Number.isFinite(historical) ? simulated - historical : null;
    const net = Number.isFinite(base) && Number.isFinite(simulated) ? base - simulated : null;
    const anyMissing = selected.some(row => !Number.isFinite(row.commissionBasisValue));
    $('#commissionPreviewMetrics').innerHTML = `
      <div class="preview-line"><span>Base de comissão projetada</span><strong>${money(base)}</strong></div>
      <div class="preview-line"><span>Comissão histórica média</span><strong>${money(historical)}</strong></div>
      <div class="preview-line primary"><span>Comissão no novo cenário</span><strong>${money(simulated)}</strong></div>
      <div class="preview-line"><span>Variação frente ao histórico</span><strong class="${delta > 0 ? 'negative' : 'positive'}">${delta > 0 ? '+' : ''}${money(delta)}</strong></div>
      <div class="preview-line"><span>Base líquida após comissão</span><strong>${money(net)}</strong></div>`;
    const scope = state.ruleScope === 'ALL' ? 'Todas as assessorias' : state.ruleScope;
    const mode = getEditorRules().mode === 'threshold_bonus' ? 'gatilho por resultado' : 'faixas de comissão';
    $('#commissionPreviewFoot').textContent = `${scope}: ${Number.isFinite(base) && base > 0 ? percent(simulated / base) : Number.isFinite(base) ? '0,00%' : '—'} de custo efetivo simulado, usando ${mode}. ${anyMissing ? 'Há indicadores sem histórico suficiente para alguma assessoria; confira as células exibidas como “—”.' : 'Os resultados representam o mês 1 da projeção.'}`;
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
      <tr><td>${esc(row.agency)}</td><td>${esc(row.monthLabel)}${row.partial ? ' · parcial' : ''}</td><td class="numeric">${money(row.collection)}</td><td class="numeric">${money(row.recovered)}</td><td class="numeric">${money(row.active)}</td><td class="numeric">${money(row.canceled)}</td><td class="numeric">${money(row.commission)}</td><td class="numeric">${percent(row.conversion)}</td><td class="numeric">${percent(row.costRatio)}</td></tr>`).join('') : '<tr><td colspan="9">Nenhum registro corresponde à busca e aos filtros atuais.</td></tr>';
  }

  function renderAll() {
    $('#historyFilters').classList.toggle('hidden', !state.hasHistory);
    $('#dataContext').innerHTML = state.hasHistory
      ? $('#dataContext').innerHTML
      : '<strong>Cenário manual sem histórico.</strong> Metas e comissões serão calculadas a partir dos valores informados; não há comparação histórica.';
    const groups = getScenarioGroups();
    updateMetrics(groups);
    renderOverview(groups);
    renderTargets();
    renderRuleEditor();
    renderCommissionComparison(getTargetRows());
    renderPolicy();
    renderHistory();
    $('#studyScopeLabel').textContent = `${state.agencyFilter === 'ALL' ? 'Todas as assessorias' : state.agencyFilter} · ${state.hasHistory ? `${MONTHS[state.startMonth - 1].short} a ${MONTHS[state.endMonth - 1].short}` : 'cenário manual sem histórico'}`;
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

  function getExportRows() {
    const targets = getTargetRows();
    const commissions = buildCommissionResults(targets);
    const filteredHistory = getFilteredRows().sort((a, b) => a.month - b.month || a.agency.localeCompare(b.agency, 'pt-BR'));
    const ruleRows = [];
    const scopes = [{ name: 'Todas as assessorias', rules: state.globalRules }, ...Object.entries(state.rulesByAgency).map(([name, rules]) => ({ name, rules }))];
    scopes.forEach(({ name, rules }) => {
      if (rules.mode === 'threshold_bonus') {
        ruleRows.push([name, rules.mode, rules.commissionBasis, 'Gatilho', rules.trigger.basis, rules.trigger.threshold, rules.trigger.baseRate, rules.trigger.bonusRate]);
      } else {
        rules.tiers.forEach(tier => ruleRows.push([name, rules.mode, rules.commissionBasis, tier.min, '', '', tier.rate, '']));
      }
    });
    return { targets, commissions, filteredHistory, ruleRows };
  }

  function buildWorkbook() {
    if (!window.XLSX) throw new Error('O leitor/exportador Excel não está disponível. Verifique a conexão com a biblioteca SheetJS.');
    const { targets, commissions, filteredHistory, ruleRows } = getExportRows();
    const wb = window.XLSX.utils.book_new();
    const period = `${MONTHS[state.startMonth - 1].name} a ${MONTHS[state.endMonth - 1].name}${state.year ? `/${state.year}` : ''}`;
    const summary = [
      ['ESTUDO DE METAS E COMISSÕES — ASSESSORIAS'],
      ['Arquivo de origem', state.fileName], ['Período do histórico', period],
      ['Filtro de assessoria', state.agencyFilter === 'ALL' ? 'Todas' : state.agencyFilter],
      ['Inclusão de mês parcial', state.includePartial ? 'Sim' : 'Não'],
      ['Indicador da meta', BASIS_LABELS[state.targetBasis]], ['Método de meta', MODE_LABELS[state.targetMode]],
      ['Referência', state.baselineMode === 'average' ? 'Média mensal' : 'Último mês com dado'],
      ['Aumento inicial (%)', state.targetMode === 'monthly_growth' ? state.monthlyInitialGrowth : state.growth],
      ['Crescimento mensal composto (%)', state.targetMode === 'monthly_growth' ? state.monthlyGrowthRate : 'Não aplicável'],
      ['Meses projetados', state.projectionMonths], ['Gerado em', new Date().toLocaleString('pt-BR')],
      [], ['Indicadores consolidados'], ['Meta total mês 1', sumOrNull(targets.map(row => row.targetStart))],
      ['Meta total mês final', sumOrNull(targets.map(row => row.targetFinal))], ['Total projetado', sumOrNull(targets.map(row => row.totalProjection))],
      ['Comissão simulada mês 1', sumOrNull(commissions.map(row => row.simulated))]
    ];
    appendSheet(wb, 'Resumo', summary);

    const targetData = [['Assessoria', 'Indicador', 'Referência mensal', 'Meses com dado', 'Meta mês 1', 'Meta mês final', 'Total projetado', 'Aumento inicial']];
    targets.forEach(row => targetData.push([row.agency, BASIS_LABELS[state.targetBasis], row.baseline, row.baselineCount, row.targetStart, row.targetFinal, row.totalProjection, row.increase]));
    appendSheet(wb, 'Metas', targetData);

    const evolutionData = [['Mês', ...targets.map(row => row.agency), 'Total consolidado']];
    for (let index = 0; index < state.projectionMonths; index += 1) {
      evolutionData.push([`Mês ${index + 1}`, ...targets.map(row => row.schedule[index]), sum(targets.map(row => row.schedule[index]))]);
    }
    appendSheet(wb, 'Evolução mensal', evolutionData);

    const commissionData = [['Assessoria', 'Base de comissão', 'Indicador do gatilho', 'Gatilho atingido', 'Taxa efetiva (%)', 'Comissão simulada', 'Comissão histórica média', 'Variação de custo', 'Líquido após comissão']];
    commissions.forEach(row => commissionData.push([row.agency, row.commissionBasisValue, row.triggerBasisValue, row.rules.mode === 'threshold_bonus' ? (row.triggerReached ? 'Sim' : 'Não') : 'N/A', row.effectiveRate, row.simulated, row.historical, row.delta, row.net]));
    appendSheet(wb, 'Comissões', commissionData);
    appendSheet(wb, 'Regras de comissão', [['Escopo', 'Modo', 'Base de cálculo', 'Faixa mínima (R$)', 'Indicador do gatilho', 'Gatilho (R$)', 'Taxa-base (%)', 'Bônus (%)'], ...ruleRows]);

    const historyData = [['Assessoria', 'Mês', 'Parcial', 'Cobrança', 'Recuperado bruto', 'Recuperado ativo', 'Recuperado cancelado', 'Comissão', 'Conversão', 'Custo / recuperação']];
    filteredHistory.forEach(row => historyData.push([row.agency, row.monthName, row.partial ? 'Sim' : 'Não', row.collection, row.recovered, row.active, row.canceled, row.commission, Number.isFinite(row.conversion) ? row.conversion * 100 : null, Number.isFinite(row.costRatio) ? row.costRatio * 100 : null]));
    appendSheet(wb, 'Histórico filtrado', historyData);

    if (state.policy.length) appendSheet(wb, 'Política atual', [['Faixa de atraso', 'Comissão B2C (%)', 'Comissão B2B (%)'], ...state.policy.map(row => [row.range, row.b2c * 100, row.b2b * 100])]);
    return wb;
  }

  function appendSheet(workbook, name, rows) {
    const sheet = window.XLSX.utils.aoa_to_sheet(rows);
    sheet['!cols'] = (rows[0] || []).map((_, index) => ({ wch: Math.min(34, Math.max(14, ...rows.slice(0, 30).map(row => String(row[index] ?? '').length + 2))) }));
    window.XLSX.utils.book_append_sheet(workbook, sheet, name);
  }

  function exportStudyExcel() {
    try {
      const workbook = buildWorkbook();
      const dateTag = new Date().toISOString().slice(0, 10);
      window.XLSX.writeFile(workbook, `estudo-assessorias-${dateTag}.xlsx`);
      toast('Estudo Excel gerado com abas de resumo, metas, evolução, comissões e histórico.');
    } catch (error) { toast(error.message || 'Não foi possível gerar o Excel.', 'error'); }
  }

  function buildPrintReport() {
    const { targets, commissions } = getExportRows();
    const totals = {
      start: sumOrNull(targets.map(row => row.targetStart)), final: sumOrNull(targets.map(row => row.targetFinal)),
      projection: sumOrNull(targets.map(row => row.totalProjection)), commission: sumOrNull(commissions.map(row => row.simulated))
    };
    const period = `${MONTHS[state.startMonth - 1].name} a ${MONTHS[state.endMonth - 1].name}${state.year ? `/${state.year}` : ''}`;
    const monthHeaders = Array.from({ length: state.projectionMonths }, (_, index) => `<th>M${index + 1}</th>`).join('');
    const evolutionRows = Array.from({ length: state.projectionMonths }, (_, monthIndex) => {
      const monthly = sum(targets.map(row => row.schedule[monthIndex]));
      return `<tr><td>Mês ${monthIndex + 1}</td><td>${money(monthly)}</td></tr>`;
    }).join('');
    $('#printReport').innerHTML = `
      <header class="print-header"><div><div class="print-kicker">CENTRAL DE RECUPERAÇÃO · MÓDULO ASSESSORIAS</div><h1>Estudo de metas e comissões</h1><p>Relatório gerado em ${new Date().toLocaleString('pt-BR')}</p></div><div class="print-period">${esc(period)}</div></header>
      <section class="print-meta"><div><span>Arquivo de origem</span><strong>${esc(state.fileName)}</strong></div><div><span>Indicador da meta</span><strong>${esc(BASIS_LABELS[state.targetBasis])}</strong></div><div><span>Método</span><strong>${esc(MODE_LABELS[state.targetMode])}</strong></div><div><span>Filtro</span><strong>${esc(state.agencyFilter === 'ALL' ? 'Todas as assessorias' : state.agencyFilter)}</strong></div></section>
      <section class="print-kpis"><div><span>Meta no mês 1</span><strong>${money(totals.start)}</strong></div><div><span>Meta no mês final</span><strong>${money(totals.final)}</strong></div><div><span>Total projetado</span><strong>${money(totals.projection)}</strong></div><div><span>Comissão simulada · mês 1</span><strong>${money(totals.commission)}</strong></div></section>
      <h2>Metas por assessoria</h2><table class="print-table"><thead><tr><th>Assessoria</th><th>Referência</th><th>Meta mês 1</th><th>Meta mês final</th><th>Total projetado</th></tr></thead><tbody>${targets.map(row => `<tr><td>${esc(row.agency)}</td><td>${money(row.baseline)}</td><td>${money(row.targetStart)}</td><td>${money(row.targetFinal)}</td><td>${money(row.totalProjection)}</td></tr>`).join('')}</tbody></table>
      <h2>Evolução mensal consolidada</h2><table class="print-table"><thead><tr><th>Período</th><th>Meta total</th></tr></thead><tbody>${evolutionRows}</tbody></table>
      <h2>Comissões por assessoria</h2><table class="print-table"><thead><tr><th>Assessoria</th><th>Base comissão</th><th>Gatilho</th><th>Taxa efetiva</th><th>Comissão simulada</th><th>Líquido</th></tr></thead><tbody>${commissions.map(row => `<tr><td>${esc(row.agency)}</td><td>${money(row.commissionBasisValue)}</td><td>${row.rules.mode === 'threshold_bonus' ? (row.triggerReached ? 'Atingido' : 'Não atingido') : 'N/A'}</td><td>${percentPoints(row.effectiveRate)}</td><td>${money(row.simulated)}</td><td>${money(row.net)}</td></tr>`).join('')}</tbody></table>
      <footer class="print-footer"><strong>Premissas:</strong> resultados são simulações agregadas com base nos campos disponíveis no histórico. Não representam apuração definitiva de comissão por título. Meses parciais ${state.includePartial ? 'foram incluídos' : 'foram excluídos'}.</footer>`;
  }

  function exportStudyPdf() {
    buildPrintReport();
    toast('Na janela de impressão, selecione “Salvar como PDF”.');
    window.print();
  }

  function syncGrowth(value, source) {
    state.growth = clamp(value, 0, 1000);
    $('#growthNumber').value = String(state.growth);
    $('#growthSlider').value = String(Math.min(100, state.growth));
    if (state.rows.length) { renderTargets(); renderCommissionComparison(getTargetRows()); }
  }

  function updateManualTarget(agency, value) {
    const parsed = parseOptionalNumeric(value);
    const key = manualTargetKey(agency);
    if (parsed === null || parsed < 0) delete state.manualTargets[key];
    else state.manualTargets[key] = parsed;
    renderTargets();
    renderCommissionComparison(getTargetRows());
  }

  function handleRuleInput(event) {
    const target = event.target;
    const field = target.dataset.tierField;
    if (!field) return;
    const index = Number(target.dataset.tierIndex);
    const rules = ensureEditableRules();
    const tier = rules.tiers[index];
    if (!tier) return;
    const raw = Math.max(0, Number(target.value) || 0);
    tier[field] = field === 'rate' ? Math.min(100, raw) : raw;
    rules.tiers.sort((a, b) => a.min - b.min);
    renderRuleEditor();
    renderCommissionComparison(getTargetRows());
  }

  function handleTriggerInput() {
    const rules = ensureEditableRules();
    rules.trigger.basis = $('#triggerBasis').value;
    rules.trigger.threshold = clamp($('#triggerThreshold').value, 0, 1e12);
    rules.trigger.baseRate = clamp($('#triggerBaseRate').value, 0, 100);
    rules.trigger.bonusRate = clamp($('#triggerBonusRate').value, 0, 100);
    renderCommissionComparison(getTargetRows());
  }

  function resetActiveRules() {
    const rate = state.ruleScope === 'ALL' ? getGlobalBaselineRate() : getBaselineRate(state.ruleScope);
    const initial = makeDefaultRules(rate, getGlobalMonthlyRecovered());
    if (state.ruleScope === 'ALL') state.globalRules = initial;
    else state.rulesByAgency[state.ruleScope] = initial;
    renderRuleEditor();
    renderCommissionComparison(getTargetRows());
    toast('Regra restaurada à referência histórica.');
  }

  function addTier() {
    const rules = ensureEditableRules();
    const sorted = [...rules.tiers].sort((a, b) => a.min - b.min);
    const last = sorted[sorted.length - 1] || { min: 0, rate: 0 };
    rules.tiers.push({ min: Math.max(100000, last.min + 100000), rate: Math.min(100, last.rate + 0.5) });
    renderRuleEditor();
    renderCommissionComparison(getTargetRows());
  }

  function startManualScenario() {
    const names = $('#manualAgencyNames').value.split(/\r?\n/).map(value => value.trim()).filter(Boolean);
    const unique = [...new Map(names.map(name => [normalized(name), name])).values()];
    if (!unique.length) {
      toast('Informe ao menos o nome de uma assessoria para criar o cenário.', 'error');
      $('#manualAgencyNames').focus();
      return;
    }
    state.rows = [];
    state.policy = [];
    state.fileName = 'Cenário manual (sem histórico)';
    state.year = null;
    state.hasHistory = false;
    state.agencies = unique.sort((a, b) => a.localeCompare(b, 'pt-BR'));
    state.startMonth = 1;
    state.endMonth = 12;
    state.includePartial = false;
    state.agencyFilter = 'ALL';
    state.targetMode = 'manual';
    state.manualTargets = {};
    state.baselineMode = 'average';
    state.targetBasis = 'recovered';
    state.projectionMonths = 12;
    populateFilters();
    initializeRules();
    $('#historyFilters').classList.add('hidden');
    $('#manualScenarioSetup').classList.add('hidden');
    $('#emptyState').classList.add('hidden');
    $('#workspace').classList.remove('hidden');
    $('#clearDataBtn').classList.remove('hidden');
    $('#dataContext').innerHTML = '<strong>Cenário manual sem histórico.</strong> Metas e comissões serão calculadas a partir dos valores informados; não há comparação histórica.';
    $('#statusFile').textContent = '';
    setStatus(`Cenário livre criado com ${state.agencies.length} assessoria(s). Informe as metas na tabela.`, 'success');
    renderAll();
    activateTab('targets');
    toast('Cenário manual pronto.');
  }

  function addScenarioAgency() {
    const input = $('#newAgencyName');
    const name = input.value.trim();
    if (!name) { toast('Informe o nome da nova assessoria.', 'error'); input.focus(); return; }
    if (state.agencies.some(existing => normalized(existing) === normalized(name))) {
      toast('Essa assessoria já existe no cenário.', 'error'); input.focus(); return;
    }
    state.agencies.push(name);
    state.agencies.sort((a, b) => a.localeCompare(b, 'pt-BR'));
    refreshAgencyOptions();
    input.value = '';
    renderAll();
    toast(`Assessoria “${name}” adicionada ao cenário.`);
  }

  function bindEvents() {
    const fileInput = $('#fileInput');
    ['importTopBtn', 'importMainBtn', 'emptyImportBtn'].forEach(id => $(`#${id}`).addEventListener('click', () => fileInput.click()));
    fileInput.addEventListener('change', event => {
      const file = event.target.files && event.target.files[0];
      importWorkbook(file);
      fileInput.value = '';
    });
    $('#createManualScenarioBtn').addEventListener('click', () => {
      $('#manualScenarioSetup').classList.remove('hidden');
      $('#manualAgencyNames').focus();
    });
    $('#confirmManualScenarioBtn').addEventListener('click', startManualScenario);
    $('#addAgencyBtn').addEventListener('click', addScenarioAgency);
    $('#newAgencyName').addEventListener('keydown', event => { if (event.key === 'Enter') addScenarioAgency(); });

    $('#clearDataBtn').addEventListener('click', () => {
      state.rows = []; state.policy = []; state.fileName = ''; state.agencies = []; state.rulesByAgency = {}; state.manualTargets = {}; state.hasHistory = false;
      $('#workspace').classList.add('hidden'); $('#emptyState').classList.remove('hidden'); $('#clearDataBtn').classList.add('hidden');
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

    $('#targetBasis').addEventListener('change', event => { state.targetBasis = event.target.value; renderTargets(); renderCommissionComparison(getTargetRows()); });
    $('#baselineMode').addEventListener('change', event => { state.baselineMode = event.target.value; renderTargets(); renderCommissionComparison(getTargetRows()); });
    $('#targetMode').addEventListener('change', event => {
      state.targetMode = event.target.value;
      renderTargets(); renderCommissionComparison(getTargetRows());
    });
    $('#growthNumber').addEventListener('input', event => syncGrowth(event.target.value));
    $('#growthSlider').addEventListener('input', event => syncGrowth(event.target.value, 'slider'));
    $('#monthlyGrowthRate').addEventListener('input', event => { state.monthlyGrowthRate = clamp(event.target.value, 0, 100); renderTargets(); renderCommissionComparison(getTargetRows()); });
    $('#growthNumberMonthly').addEventListener('input', event => { state.monthlyInitialGrowth = clamp(event.target.value, 0, 1000); renderTargets(); renderCommissionComparison(getTargetRows()); });
    $('#projectionMonths').addEventListener('change', event => { state.projectionMonths = clamp(Math.round(Number(event.target.value) || 12), 1, 36); event.target.value = String(state.projectionMonths); renderTargets(); renderCommissionComparison(getTargetRows()); });
    $('#applyManualAllBtn').addEventListener('click', () => {
      const value = parseOptionalNumeric($('#manualFillValue').value);
      if (value === null || value < 0) { toast('Informe um valor válido para preencher as metas.', 'error'); return; }
      getTargetRows().forEach(row => { state.manualTargets[manualTargetKey(row.agency)] = value; });
      state.manualDefault = value;
      renderTargets(); renderCommissionComparison(getTargetRows());
      toast('Valor aplicado a todas as assessorias do recorte.');
    });
    $('#targetsTableBody').addEventListener('change', event => {
      const input = event.target.closest('[data-manual-agency]');
      if (input) updateManualTarget(input.dataset.manualAgency, input.value);
    });

    $('#ruleScope').addEventListener('change', event => { state.ruleScope = event.target.value; renderRuleEditor(); renderCommissionComparison(getTargetRows()); });
    $('#commissionMode').addEventListener('change', event => { ensureEditableRules().mode = event.target.value; renderRuleEditor(); renderCommissionComparison(getTargetRows()); });
    $('#commissionBasis').addEventListener('change', event => { ensureEditableRules().commissionBasis = event.target.value; renderCommissionComparison(getTargetRows()); });
    $('#tiersTableBody').addEventListener('change', handleRuleInput);
    $('#tiersTableBody').addEventListener('click', event => {
      const button = event.target.closest('[data-delete-tier]');
      if (!button) return;
      const rules = ensureEditableRules();
      if (rules.tiers.length <= 1) return;
      rules.tiers.splice(Number(button.dataset.deleteTier), 1);
      renderRuleEditor(); renderCommissionComparison(getTargetRows());
    });
    $('#addTierBtn').addEventListener('click', addTier);
    $('#resetRulesBtn').addEventListener('click', resetActiveRules);
    ['triggerBasis', 'triggerThreshold', 'triggerBaseRate', 'triggerBonusRate'].forEach(id => $(`#${id}`).addEventListener('change', handleTriggerInput));

    $('#historySearch').addEventListener('input', renderHistory);
    ['exportStudyExcelBtn', 'exportTargetsBtn', 'exportCommissionBtn', 'exportHistoryBtn'].forEach(id => $(`#${id}`).addEventListener('click', exportStudyExcel));
    $('#exportStudyPdfBtn').addEventListener('click', exportStudyPdf);

    const dropTargets = [$('#emptyState'), $('.page-heading')];
    dropTargets.forEach(target => {
      target.addEventListener('dragover', event => { event.preventDefault(); target.classList.add('drag-over'); });
      target.addEventListener('dragleave', () => target.classList.remove('drag-over'));
      target.addEventListener('drop', event => {
        event.preventDefault(); target.classList.remove('drag-over');
        const file = event.dataTransfer.files && event.dataTransfer.files[0];
        if (file) importWorkbook(file);
      });
    });
  }

  bindEvents();
  tryLoadLocalHistory();
})();
