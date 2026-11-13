/*
 * Simulador Cobranças Master — módulo Assessorias.
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
  const FULL_MONTH_LOOKUP = { JANEIRO: 1, FEVEREIRO: 2, MARCO: 3, ABRIL: 4, MAIO: 5, JUNHO: 6, JULHO: 7, AGOSTO: 8, SETEMBRO: 9, OUTUBRO: 10, NOVEMBRO: 11, DEZEMBRO: 12 };
  const $ = selector => document.querySelector(selector);

  const BASIS_LABELS = {
    recovered: 'Recuperado bruto', active: 'Recuperado ativo', canceled: 'Recuperado cancelado'
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
    activeTab: 'overview', activeModule: 'assessorias',
    agingByAgency: {}, commissionScenarios: [], activeCommissionScenarioId: null, commissionScenarioCounter: 1
  };

  const money = (value, compact = false) => {
    if (value === null || value === undefined || value === '' || !Number.isFinite(Number(value))) return '—';
    const n = Number(value);
    if (compact && Math.abs(n) >= 1000000) return n.toLocaleString('pt-BR', { style: 'currency', currency: 'BRL', notation: 'compact', maximumFractionDigits: 1 });
    return n.toLocaleString('pt-BR', { style: 'currency', currency: 'BRL', maximumFractionDigits: 2 });
  };
  const number = (value, digits = 0) => (Number(value) || 0).toLocaleString('pt-BR', { minimumFractionDigits: digits, maximumFractionDigits: digits });
  const percent = (value, digits = 2) => value === null || value === undefined || value === '' || !Number.isFinite(Number(value)) ? '—' : `${number(Number(value) * 100, digits)}%`;
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


  function mapPolicyAgingBucket(label) {
    const value = normalized(label).replace(/\s+/g, ' ');
    if (value.includes('NO VENCIMENTO') || value.includes('PAGAMENTO EM DIA') || value.startsWith('00')) return '00';
    if (value.includes('01-20') || value.includes('01 - 20') || value.includes('01 A 20')) return '01-20';
    if (value.includes('ACIMA DE 1800') || value.includes('MAIOR QUE 1800')) return '1800+';
    const range = value.match(/(\d{1,4})\s*[-–]\s*(\d{1,4})/);
    return range ? `${Number(range[1])}-${Number(range[2])}` : null;
  }

  function mapDetailAgingBucket(label) {
    const value = normalized(label);
    if (value.includes('PAGAMENTO EM DIA') || value.startsWith('00')) return '00';
    if (value.includes('MAIOR QUE 1800') || value.includes('ACIMA DE 1800')) return '1800+';
    if (value.includes('01 A 14') || value.includes('15 A 20')) return '01-20';
    const range = value.match(/^\s*\d+\s*[-–]\s*(\d{1,4})\s*(?:A|[-–])\s*(\d{1,4})/);
    return range ? `${Number(range[1])}-${Number(range[2])}` : null;
  }

  function parseAgencyAgingRows(rows) {
    const byMonth = new Map();
    (rows || []).forEach(row => {
      const monthRaw = normalized(row?.[0]).split(/[\s(]/)[0];
      const month = MONTH_LOOKUP[monthRaw] || FULL_MONTH_LOOKUP[monthRaw];
      const bucket = mapDetailAgingBucket(row?.[1]);
      if (!month) return;
      if (!byMonth.has(month)) byMonth.set(month, { month, amounts: {}, cpf: 0, cnpj: 0, hasCpf: false, hasCnpj: false });
      const data = byMonth.get(month);
      if (bucket) data.amounts[bucket] = (data.amounts[bucket] || 0) + parseNumeric(row?.[2]);
      const contractType = normalized(row?.[5]);
      if (contractType === 'CPF') {
        data.cpf += parseNumeric(row?.[6]);
        data.hasCpf = true;
      } else if (contractType === 'CNPJ') {
        data.cnpj += parseNumeric(row?.[6]);
        data.hasCnpj = true;
      }
    });
    return [...byMonth.values()].sort((a, b) => a.month - b.month);
  }

  function parseAgencyAging(workbook, agencies) {
    const parsed = {};
    agencies.forEach(agency => {
      const sheetName = workbook.SheetNames.find(name => normalized(name) === normalized(agency));
      if (!sheetName) return;
      const sheetRows = window.XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { header: 1, raw: true, defval: null, blankrows: false });
      const months = parseAgencyAgingRows(sheetRows);
      if (months.length) parsed[normalized(agency)] = months;
    });
    return parsed;
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
      result.push({ range: label, key: mapPolicyAgingBucket(label), b2c: parseNumeric(row[startCol + 1]), b2b: parseNumeric(row[startCol + 2]) });
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
    const options = ['<option value="ALL">Todas as assessorias</option>', ...state.agencies.map(name => `<option value="${esc(name)}">${esc(name)}</option>`)].join('');
    $('#agencyFilter').innerHTML = options;
    $('#agencyFilter').value = state.agencies.includes(selectedAgency) ? selectedAgency : 'ALL';
    state.agencyFilter = $('#agencyFilter').value;
    const commissionSelect = $('#commissionAgency');
    if (commissionSelect) {
      const previous = commissionSelect.value;
      commissionSelect.innerHTML = state.agencies.map(name => `<option value="${esc(name)}">${esc(name)}</option>`).join('');
      if (state.agencies.includes(previous)) commissionSelect.value = previous;
    }
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
      state.agingByAgency = parseAgencyAging(workbook, state.agencies);
      state.commissionScenarios = [];
      state.activeCommissionScenarioId = null;
      state.commissionScenarioCounter = 1;
      state.manualTargets = {};
      state.manualDefault = null;
      populateFilters();
      initializeCommissionScenarios();
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
      renderAll();
      activateModule('assessorias');
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

  function buildTargetEvolutionSvg(targets) {
    const months = state.projectionMonths;
    const planned = Array.from({ length: months }, (_, index) => sumOrNull(targets.map(row => row.schedule[index])));
    const historicReference = state.hasHistory ? sumOrNull(targets.map(row => row.baseline)) : null;
    const initialReference = sumOrNull(targets.map(row => row.schedule[0]));
    const reference = Number.isFinite(historicReference) ? historicReference : initialReference;
    const validValues = [...planned.filter(Number.isFinite), ...(Number.isFinite(reference) ? [reference] : [])];
    if (!validValues.length) return '<div class="chart-empty">Informe uma meta inicial para visualizar a projeção.</div>';

    const width = 1000;
    const height = 320;
    const left = 92;
    const right = 34;
    const top = 24;
    const bottom = 54;
    const plotW = width - left - right;
    const plotH = height - top - bottom;
    const maxValue = Math.max(1, ...validValues) * 1.12;
    const x = index => months <= 1 ? left + plotW / 2 : left + (index / (months - 1)) * plotW;
    const y = value => top + plotH - (Math.max(0, value) / maxValue) * plotH;
    const pathFor = values => values.map((value, index) => `${index ? 'L' : 'M'} ${x(index).toFixed(1)} ${y(Number.isFinite(value) ? value : 0).toFixed(1)}`).join(' ');
    const targetPath = pathFor(planned);
    const grid = Array.from({ length: 5 }, (_, index) => {
      const value = maxValue * (4 - index) / 4;
      const yy = y(value);
      const label = value.toLocaleString('pt-BR', { style: 'currency', currency: 'BRL', notation: 'compact', maximumFractionDigits: 1 });
      return `<g><line x1="${left}" y1="${yy}" x2="${width-right}" y2="${yy}" stroke="#E4EAF1" stroke-width="1"/><text x="${left-12}" y="${yy+4}" text-anchor="end" fill="#7C8998" font-size="12">${label}</text></g>`;
    }).join('');
    const xLabels = Array.from({ length: months }, (_, index) => {
      const show = months <= 12 || index === 0 || index === months - 1 || index % Math.ceil(months / 10) === 0;
      return show ? `<text x="${x(index)}" y="${height-20}" text-anchor="middle" fill="#7C8998" font-size="12">M${index+1}</text>` : '';
    }).join('');
    const targetPoints = planned.map((value, index) => Number.isFinite(value) ? `<circle cx="${x(index)}" cy="${y(value)}" r="4.5" fill="#1F5DA8" stroke="#FFFFFF" stroke-width="2"><title>Mês ${index+1}: ${money(value)}</title></circle>` : '').join('');
    const referenceLine = Number.isFinite(reference)
      ? `<line x1="${left}" y1="${y(reference)}" x2="${width-right}" y2="${y(reference)}" stroke="#E2B52D" stroke-width="3" stroke-dasharray="9 7"/><text x="${width-right}" y="${Math.max(14, y(reference)-9)}" text-anchor="end" fill="#A27B0B" font-size="12" font-weight="700">Referência ${money(reference, true)}</text>`
      : '';
    const referenceLabel = Number.isFinite(historicReference) ? 'Referência histórica mensal' : 'Patamar inicial';
    return `<div class="chart-legend"><span><i class="legend-blue"></i>Meta projetada</span><span><i class="legend-yellow"></i>${referenceLabel}</span></div>
      <svg viewBox="0 0 ${width} ${height}" class="target-chart-svg" role="img" aria-label="Meta projetada ao longo de ${months} meses, comparada à referência mensal">
        ${grid}
        ${referenceLine}
        <path d="${targetPath}" fill="none" stroke="#1F5DA8" stroke-width="4" stroke-linecap="round" stroke-linejoin="round"/>
        ${targetPoints}
        ${xLabels}
      </svg>`;
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
    $('#targetBasisHint').textContent = 'Meses sem valor para este indicador são desconsiderados da média, não tratados como zero.';
    $('#targetCoverageNote').textContent = !state.hasHistory
      ? `${enteredTargets} de ${targets.length} assessoria(s) com meta inicial informada.`
      : `${covered} mês(es) com indicador disponível entre ${available} registro(s) no recorte.`;
    $('#targetHighlight').innerHTML = `
      <div class="target-intro"><span>${esc(basisLabel)} · ${esc(modeLabel)}</span><strong>${state.targetMode === 'monthly_growth' ? `${number(state.monthlyGrowthRate, 2)}% composto ao mês` : state.targetMode === 'history_growth' ? `${number(state.growth, 1)}% sobre a referência` : 'Valores definidos manualmente'}</strong></div>
      <div class="target-kpi"><span>Meta consolidada · mês 1</span><strong>${money(totalStart, true)}</strong></div>
      <div class="target-kpi"><span>Meta no mês final</span><strong>${money(totalFinal, true)}</strong></div>
      <div class="target-kpi"><span>Total projetado · ${state.projectionMonths} meses</span><strong>${money(totalProjection, true)}</strong></div>`;

    $('#targetEvolutionChart').innerHTML = buildTargetEvolutionSvg(targets);

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

  const DEFAULT_COMMISSION_SCENARIO = {
    basis: 'canceled', targetMode: 'portfolio_percent', portfolioValue: null, targetPercent: 10, targetAmount: null
  };

  function getActiveCommissionScenario() {
    return state.commissionScenarios.find(item => item.id === state.activeCommissionScenarioId) || state.commissionScenarios[0] || null;
  }

  function makeCommissionScenario(agency, name = null) {
    const id = `fee-${state.commissionScenarioCounter++}`;
    const label = name || `${agency || 'Assessoria'} — nova tabela`;
    return {
      id, name: label, agency: agency || state.agencies[0] || '',
      basis: DEFAULT_COMMISSION_SCENARIO.basis,
      targetMode: DEFAULT_COMMISSION_SCENARIO.targetMode,
      portfolioValue: null, targetPercent: DEFAULT_COMMISSION_SCENARIO.targetPercent, targetAmount: null,
      rates: state.policy.map(row => ({
        key: row.key || mapPolicyAgingBucket(row.range), range: row.range,
        newB2C: Number(row.b2c || 0) * 100, newB2B: Number(row.b2b || 0) * 100
      }))
    };
  }

  function initializeCommissionScenarios() {
    state.commissionScenarios = [];
    state.activeCommissionScenarioId = null;
    if (state.agencies.length) {
      const scenario = makeCommissionScenario(state.agencies[0], `${state.agencies[0]} — incentivo cancelado 10%`);
      state.commissionScenarios.push(scenario);
      state.activeCommissionScenarioId = scenario.id;
    }
    refreshAgencyOptions();
    renderCommissionModule();
  }


  function getAgingProfile(agency) {
    const months = state.agingByAgency[normalized(agency)] || [];
    const totals = {};
    let cpf = 0, cnpj = 0, cpfFound = false, cnpjFound = false, includedMonths = 0;
    months.forEach(item => {
      if (item.month < state.startMonth || item.month > state.endMonth) return;
      const isPartial = state.rows.some(row => row.agency === agency && row.month === item.month && row.partial);
      if (isPartial && !state.includePartial) return;
      includedMonths += 1;
      Object.entries(item.amounts || {}).forEach(([key, value]) => { totals[key] = (totals[key] || 0) + value; });
      if (item.hasCpf) { cpf += item.cpf; cpfFound = true; }
      if (item.hasCnpj) { cnpj += item.cnpj; cnpjFound = true; }
    });
    const totalRecovered = sum(Object.values(totals));
    if (!includedMonths || totalRecovered <= 0) return null;
    const mixTotal = (cpfFound ? cpf : 0) + (cnpjFound ? cnpj : 0);
    const cpfShare = mixTotal > 0 && cpfFound ? cpf / mixTotal : (cpfFound ? 1 : 0);
    const cnpjShare = mixTotal > 0 && cnpjFound ? cnpj / mixTotal : (cnpjFound ? 1 : 0);
    return { totals, totalRecovered, cpfShare, cnpjShare, includedMonths, hasContractMix: mixTotal > 0 };
  }

  function calculateScenarioFees(scenario, recoveryValue, profile) {
    if (!profile || !profile.hasContractMix || !state.policy.length) return { amount: null, currentAmount: null, rows: [] };
    const rates = scenario.rates || [];
    const rows = state.policy.map(policyRow => {
      const key = policyRow.key || mapPolicyAgingBucket(policyRow.range);
      const agingAmount = profile.totals[key] || 0;
      const share = profile.totalRecovered > 0 ? agingAmount / profile.totalRecovered : 0;
      const allocated = Number.isFinite(recoveryValue) ? recoveryValue * share : null;
      const currentB2C = Number(policyRow.b2c || 0) * 100;
      const currentB2B = Number(policyRow.b2b || 0) * 100;
      const override = rates.find(rate => rate.key === key);
      const newB2C = override ? Number(override.newB2C) : currentB2C;
      const newB2B = override ? Number(override.newB2B) : currentB2B;
      const currentRate = currentB2C * profile.cpfShare + currentB2B * profile.cnpjShare;
      const newRate = newB2C * profile.cpfShare + newB2B * profile.cnpjShare;
      const currentFee = Number.isFinite(allocated) ? allocated * currentRate / 100 : null;
      const newFee = Number.isFinite(allocated) ? allocated * newRate / 100 : null;
      return { key, range: policyRow.range, share, allocated, currentB2C, currentB2B, newB2C, newB2B, currentRate, newRate, currentFee, newFee, delta: Number.isFinite(newFee) && Number.isFinite(currentFee) ? newFee - currentFee : null };
    });
    return { amount: sumOrNull(rows.map(row => row.newFee)), currentAmount: sumOrNull(rows.map(row => row.currentFee)), rows };
  }

  function buildCommissionScenarioResult(scenario) {
    if (!scenario) return null;
    const records = groupRows(getFilteredRows(true));
    const group = records.find(row => row.agency === scenario.agency);
    const baselineInfo = group ? getBaseline(group, scenario.basis, 'average') : { value: null, count: 0 };
    const baseline = baselineInfo.value;
    const target = scenario.targetMode === 'portfolio_percent'
      ? (Number.isFinite(Number(scenario.portfolioValue)) && Number(scenario.portfolioValue) > 0 ? Number(scenario.portfolioValue) * Number(scenario.targetPercent || 0) / 100 : null)
      : (Number.isFinite(Number(scenario.targetAmount)) && Number(scenario.targetAmount) > 0 ? Number(scenario.targetAmount) : null);
    const alreadyAtGoal = Number.isFinite(baseline) && Number.isFinite(target) && baseline >= target;
    const additionalRecovery = Number.isFinite(target) && Number.isFinite(baseline) ? Math.max(0, target - baseline) : null;
    const evaluationValue = Number.isFinite(target) && Number.isFinite(baseline) ? Math.max(target, baseline) : target;
    const profile = getAgingProfile(scenario.agency);
    const evaluationFees = calculateScenarioFees(scenario, evaluationValue, profile);
    const baselineFees = calculateScenarioFees(scenario, baseline, profile);
    const additionalCommissionVsToday = Number.isFinite(evaluationFees.amount) && Number.isFinite(baselineFees.currentAmount)
      ? evaluationFees.amount - baselineFees.currentAmount : null;
    const extraCommissionForNewRates = Number.isFinite(evaluationFees.amount) && Number.isFinite(evaluationFees.currentAmount)
      ? evaluationFees.amount - evaluationFees.currentAmount : null;
    const netIncremental = Number.isFinite(additionalRecovery) && Number.isFinite(additionalCommissionVsToday)
      ? additionalRecovery - additionalCommissionVsToday : null;
    return {
      scenario, baseline, baselineCount: baselineInfo.count, target, evaluationValue, additionalRecovery, profile,
      currentFeeAtTarget: evaluationFees.currentAmount, newFeeAtTarget: evaluationFees.amount,
      currentFeeAtBaseline: baselineFees.currentAmount, additionalCommissionVsToday,
      extraCommissionForNewRates, netIncremental, agingRows: evaluationFees.rows,
      alreadyAtGoal
    };
  }

  function renderCurrentPolicyTable() {
    $('#currentPolicyTableBody').innerHTML = state.policy.length
      ? state.policy.map(row => `<tr><td>${esc(row.range)}</td><td class="numeric">${percentPoints(row.b2c * 100)}</td><td class="numeric">${percentPoints(row.b2b * 100)}</td></tr>`).join('')
      : '<tr><td colspan="3">Importe o Excel com a política vigente para exibir as taxas por aging.</td></tr>';
    $('#currentPolicySource').textContent = state.policy.length ? 'Origem: Política de Comissionamento do Excel' : 'Sem política importada';
  }

  function renderCommissionScenarioSelect() {
    const select = $('#commissionScenarioSelect');
    const scenario = getActiveCommissionScenario();
    select.innerHTML = state.commissionScenarios.map(item => `<option value="${esc(item.id)}">${esc(item.name)}</option>`).join('');
    select.disabled = !state.commissionScenarios.length;
    if (scenario) select.value = scenario.id;
    $('#deleteCommissionScenarioBtn').disabled = state.commissionScenarios.length <= 1;
  }

  function renderCommissionRatesTable(result) {
    const scenario = result?.scenario || getActiveCommissionScenario();
    const profile = result?.profile || (scenario ? getAgingProfile(scenario.agency) : null);
    const rows = state.policy.map(policyRow => {
      const key = policyRow.key || mapPolicyAgingBucket(policyRow.range);
      const override = scenario?.rates?.find(rate => rate.key === key);
      const newB2C = override ? Number(override.newB2C) : Number(policyRow.b2c || 0) * 100;
      const newB2B = override ? Number(override.newB2B) : Number(policyRow.b2b || 0) * 100;
      const share = profile && profile.totalRecovered > 0 ? (profile.totals[key] || 0) / profile.totalRecovered : null;
      return { range: policyRow.range, newB2C, newB2B, currentB2C: Number(policyRow.b2c || 0) * 100, currentB2B: Number(policyRow.b2b || 0) * 100, share };
    });
    $('#commissionRatesTableBody').innerHTML = rows.length ? rows.map((row, index) => `
      <tr><td><strong>${esc(row.range)}</strong></td><td class="numeric">${percent(row.share, 1)}</td><td class="numeric">${percentPoints(row.currentB2C)}</td><td class="numeric">${percentPoints(row.currentB2B)}</td>
      <td class="numeric"><input class="commission-rate-input" type="number" min="0" max="100" step="0.01" value="${row.newB2C.toFixed(3)}" data-rate-index="${index}" data-rate-field="newB2C" aria-label="Nova taxa B2C para ${esc(row.range)}"></td>
      <td class="numeric"><input class="commission-rate-input" type="number" min="0" max="100" step="0.01" value="${row.newB2B.toFixed(3)}" data-rate-index="${index}" data-rate-field="newB2B" aria-label="Nova taxa B2B para ${esc(row.range)}"></td></tr>`).join('')
      : '<tr><td colspan="6">Não há política por aging carregada para editar.</td></tr>';
  }

  function renderCommissionScenarioMetrics(result) {
    const metrics = $('#commissionScenarioMetrics');
    if (!result) {
      metrics.innerHTML = '<div class="scenario-empty">Crie um estudo para começar a simulação.</div>';
      $('#commissionScenarioDisclaimer').textContent = '';
      $('#commissionScenarioSentence').textContent = 'Selecione uma assessoria e informe a condição para desbloquear a nova tabela.';
      $('#commissionFeeChart').innerHTML = '<p class="muted-empty">A comparação aparecerá após configurar a meta e importar a distribuição por aging.</p>';
      $('#commissionAgingBreakdownBody').innerHTML = '<tr><td colspan="6">Nenhum cenário selecionado.</td></tr>';
      renderCommissionRatesTable(null);
      return;
    }
    const scenario = result.scenario;
    const validTarget = Number.isFinite(result.target);
    const targetLabel = scenario.targetMode === 'portfolio_percent'
      ? `${number(scenario.targetPercent, 2)}% da carteira cancelada de ${money(scenario.portfolioValue)}`
      : `meta de ${money(result.target)}`;
    $('#commissionScenarioSentence').innerHTML = `<strong>${esc(scenario.agency || 'Assessoria')}</strong> deverá atingir <strong>${esc(targetLabel)}</strong> em ${esc(BASIS_LABELS[scenario.basis] || scenario.basis)} para liberar a nova tabela de comissão.`;

    const values = [
      ['Recuperação média atual', result.baseline, 'Patamar mensal do histórico'],
      ['Meta para desbloqueio', result.target, scenario.targetMode === 'portfolio_percent' ? 'Percentual do saldo cancelado informado' : 'Valor bruto definido manualmente'],
      ['Recuperação adicional estimada', result.additionalRecovery, result.alreadyAtGoal ? 'A média atual já alcança a meta' : 'Diferença entre meta e média atual'],
      ['Comissão na tabela atual', result.currentFeeAtTarget, result.alreadyAtGoal ? 'Custo no patamar médio atual' : 'Custo estimado ao atingir a meta'],
      ['Comissão na nova tabela', result.newFeeAtTarget, result.alreadyAtGoal ? 'Custo no patamar médio atual' : 'Custo estimado ao atingir a meta'],
      ['Custo incremental vs. hoje', result.additionalCommissionVsToday, 'Nova comissão na meta menos comissão atual no patamar histórico'],
      ['Saldo incremental após custo de comissão', result.netIncremental, 'Recuperação adicional menos o custo incremental de comissão']
    ];
    metrics.innerHTML = values.map((item, index) => `<div class="scenario-metric ${index === 2 || index === 6 ? 'highlight' : ''}"><span>${item[0]}</span><strong class="${index === 2 || index === 6 ? (Number.isFinite(item[1]) && item[1] >= 0 ? 'positive' : 'negative') : ''}">${money(item[1], true)}</strong><small>${item[2]}</small></div>`).join('');

    const issues = [];
    if (!validTarget) issues.push(scenario.targetMode === 'portfolio_percent' ? 'Informe o valor atualizado da carteira cancelada e um percentual válido.' : 'Informe o valor bruto que libera a nova tabela.');
    if (!Number.isFinite(result.baseline)) issues.push('Não há média histórica disponível para este indicador e período.');
    if (!result.profile) issues.push('Não foi encontrada distribuição por aging para esta assessoria no período selecionado.');
    if (result.profile && !result.profile.hasContractMix) issues.push('A composição CPF/CNPJ não foi encontrada; por isso o custo de comissão não será estimado até que exista essa informação.');
    if (!state.policy.length) issues.push('A tabela vigente por aging não foi encontrada no arquivo.');
    if (result.alreadyAtGoal) issues.push('A média histórica já está no nível ou acima desta meta; a recuperação adicional necessária é zero e o custo é avaliado sobre a média atual.');
    $('#commissionScenarioDisclaimer').textContent = issues.length ? issues.join(' ') : `Mix calculado com ${result.profile.includedMonths} mês(es) de distribuição detalhada. A comissão é estimada pela composição média CPF/CNPJ (CPF tratado como B2C e CNPJ como B2B, hipótese a validar internamente). A nova tabela incide sobre todo o valor recuperado no mês avaliado.`;

    const chart = $('#commissionFeeChart');
    const recoveryValues = [
      { label: 'Média mensal atual', value: result.baseline, style: 'current' },
      { label: 'Meta de desbloqueio', value: result.target, style: 'new' }
    ];
    const feeValues = [
      { label: 'Comissão pela tabela atual', value: result.currentFeeAtTarget, style: 'current' },
      { label: 'Comissão pela nova tabela', value: result.newFeeAtTarget, style: 'new' }
    ];
    const renderBarGroup = (heading, rows, footer) => {
      const maxValue = Math.max(1, ...rows.map(row => Number.isFinite(row.value) ? row.value : 0));
      if (!rows.every(row => Number.isFinite(row.value))) return `<div class="fee-chart-section"><h4>${heading}</h4><p class="muted-empty">Aguardando valores válidos.</p></div>`;
      return `<div class="fee-chart-section"><h4>${heading}</h4>${rows.map(row => `<div class="fee-chart-row"><div class="fee-chart-label"><span>${row.label}</span><strong>${money(row.value)}</strong></div><div class="fee-chart-track"><div class="fee-chart-fill ${row.style}" style="width:${Math.max(0, row.value / maxValue * 100)}%"></div></div></div>`).join('')}<div class="fee-chart-delta">${footer}</div></div>`;
    };
    chart.innerHTML = renderBarGroup('Recuperação mensal', recoveryValues,
      `Recuperação adicional necessária: <strong>${money(result.additionalRecovery)}</strong>${result.alreadyAtGoal ? ' · meta já alcançada pela média atual' : ''}`) +
      renderBarGroup('Custo de comissão no patamar avaliado', feeValues,
      `Diferença entre a nova e a atual tabela: <strong>${money(result.extraCommissionForNewRates)}</strong>`);

    $('#commissionAgingBreakdownBody').innerHTML = result.agingRows.length ? result.agingRows.map(row => `<tr><td>${esc(row.range)}</td><td class="numeric">${percent(row.share, 1)}</td><td class="numeric">${money(row.allocated)}</td><td class="numeric">${money(row.currentFee)}</td><td class="numeric"><strong>${money(row.newFee)}</strong></td><td class="numeric ${row.delta > 0 ? 'negative' : 'positive'}">${row.delta > 0 ? '+' : ''}${money(row.delta)}</td></tr>`).join('') : '<tr><td colspan="6">Sem distribuição histórica por aging para a assessoria selecionada.</td></tr>';
    renderCommissionRatesTable(result);
  }

  function renderCommissionModule() {
    renderCurrentPolicyTable();
    renderCommissionScenarioSelect();
    const scenario = getActiveCommissionScenario();
    if (!scenario) { renderCommissionScenarioMetrics(null); return; }
    $('#commissionScenarioName').value = scenario.name;
    $('#commissionAgency').value = scenario.agency;
    $('#commissionTriggerBasis').value = scenario.basis;
    $('#commissionTargetMode').value = scenario.targetMode;
    $('#commissionPortfolioValue').value = scenario.portfolioValue ?? '';
    $('#commissionTargetPercent').value = scenario.targetPercent ?? 10;
    $('#commissionTargetAmount').value = scenario.targetAmount ?? '';
    $('#commissionPortfolioField').classList.toggle('hidden', scenario.targetMode !== 'portfolio_percent');
    $('#commissionPercentField').classList.toggle('hidden', scenario.targetMode !== 'portfolio_percent');
    $('#commissionFixedField').classList.toggle('hidden', scenario.targetMode !== 'fixed_amount');
    renderCommissionScenarioMetrics(buildCommissionScenarioResult(scenario));
  }

  // Alias used by the existing goal controls: every change refreshes the new commission lab.
  function renderCommissionComparison() { renderCommissionModule(); }

  function createCommissionScenario() {
    if (!state.agencies.length) { toast('Adicione uma assessoria antes de criar uma tabela.', 'error'); return; }
    const currentScenario = getActiveCommissionScenario();
    const agency = state.agencyFilter !== 'ALL' ? state.agencyFilter : (currentScenario?.agency || state.agencies[0]);
    const scenario = makeCommissionScenario(agency, `${agency} — novo incentivo ${state.commissionScenarioCounter}`);
    state.commissionScenarios.push(scenario);
    state.activeCommissionScenarioId = scenario.id;
    renderCommissionModule();
    toast('Novo estudo de comissão criado.');
  }

  function deleteCommissionScenario() {
    if (state.commissionScenarios.length <= 1) return;
    state.commissionScenarios = state.commissionScenarios.filter(item => item.id !== state.activeCommissionScenarioId);
    state.activeCommissionScenarioId = state.commissionScenarios[0]?.id || null;
    renderCommissionModule();
    toast('Estudo de comissão excluído.');
  }

  function updateCommissionRate(event) {
    const input = event.target.closest('[data-rate-index][data-rate-field]');
    if (!input) return;
    const scenario = getActiveCommissionScenario();
    if (!scenario) return;
    const index = Number(input.dataset.rateIndex);
    const field = input.dataset.rateField;
    if (!scenario.rates[index] || !['newB2C', 'newB2B'].includes(field)) return;
    scenario.rates[index][field] = clamp(input.value, 0, 100);
    renderCommissionScenarioMetrics(buildCommissionScenarioResult(scenario));
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
    renderCommissionModule();
    renderHistory();
    $('#studyScopeLabel').textContent = `${state.agencyFilter === 'ALL' ? 'Todas as assessorias' : state.agencyFilter} · ${state.hasHistory ? `${MONTHS[state.startMonth - 1].short} a ${MONTHS[state.endMonth - 1].short}` : 'cenário manual sem histórico'}`;
  }

  function activateModule(moduleKey) {
    const isAssessorias = moduleKey === 'assessorias';
    state.activeModule = moduleKey;
    document.querySelectorAll('.module-button').forEach(button => {
      const active = button.dataset.module === moduleKey;
      button.classList.toggle('active', active);
      if (active) button.setAttribute('aria-current', 'page');
      else button.removeAttribute('aria-current');
    });

    $('#assessoriasHeading').classList.toggle('hidden', !isAssessorias);
    $('#privacyStrip').classList.toggle('hidden', !isAssessorias);
    $('#statusLine').classList.toggle('hidden', !isAssessorias);
    $('#emptyState').classList.toggle('hidden', !isAssessorias || state.agencies.length > 0);
    $('#workspace').classList.toggle('hidden', !isAssessorias || state.agencies.length === 0);
    $('#modulePlaceholder').classList.toggle('hidden', isAssessorias);

    const moduleInfo = {
      b2b: { title: 'B2B', text: 'Estudos de recuperação, metas e regras de comissão B2B serão adicionados nesta página.' },
      cobranca: { title: 'Cobrança interna', text: 'Indicadores e simulações da operação interna serão adicionados nesta página.' },
      carteira: { title: 'Carteira', text: 'Análises de composição, aging e evolução da carteira serão adicionadas nesta página.' }
    };
    if (!isAssessorias) {
      const info = moduleInfo[moduleKey] || moduleInfo.b2b;
      $('#modulePlaceholderTitle').textContent = info.title;
      $('#modulePlaceholderText').textContent = info.text;
    }
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
    const commissionScenarios = state.commissionScenarios.map(scenario => buildCommissionScenarioResult(scenario)).filter(Boolean);
    const filteredHistory = getFilteredRows().sort((a, b) => a.month - b.month || a.agency.localeCompare(b.agency, 'pt-BR'));
    return { targets, commissionScenarios, filteredHistory, performance: groupRows(filteredHistory) };
  }

  function buildWorkbook() {
    if (!window.XLSX) throw new Error('O leitor/exportador Excel não está disponível. Verifique a conexão com a biblioteca SheetJS.');
    const { targets, commissionScenarios, filteredHistory, performance } = getExportRows();
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
      ['Meses projetados', state.projectionMonths], ['Cenários de comissão salvos', commissionScenarios.length],
      ['Gerado em', new Date().toLocaleString('pt-BR')], [], ['Indicadores consolidados'],
      ['Meta total mês 1', sumOrNull(targets.map(row => row.targetStart))],
      ['Meta total mês final', sumOrNull(targets.map(row => row.targetFinal))],
      ['Total projetado', sumOrNull(targets.map(row => row.totalProjection))]
    ];
    appendSheet(wb, 'Resumo', summary);

    const performanceData = [['Assessoria', 'Meses com dados', 'Cobrança total (R$)', 'Recuperado bruto (R$)', 'Média mensal bruto (R$)', 'Recuperado ativo (R$)', 'Média mensal ativo (R$)', 'Recuperado cancelado (R$)', 'Média mensal cancelado (R$)', 'Comissão total (R$)', 'Comissão média (R$)', 'Custo efetivo (%)']];
    performance.forEach(group => performanceData.push([
      group.agency, group.monthCount, group.totalCollection, group.totalRecovered, group.avgRecovered,
      group.activeCount ? group.totalActive : null, group.activeCount ? group.avgActive : null,
      group.canceledCount ? group.totalCanceled : null, group.canceledCount ? group.avgCanceled : null, group.totalCommission,
      group.commissionCount ? group.avgCommission : null, Number.isFinite(group.effectiveRate) ? group.effectiveRate * 100 : null
    ]));
    appendSheet(wb, 'Desempenho por assessoria', performanceData);

    const targetData = [['Assessoria', 'Indicador', 'Referência mensal', 'Meses com dado', 'Meta mês 1', 'Meta mês final', 'Total projetado', 'Aumento inicial']];
    targets.forEach(row => targetData.push([row.agency, BASIS_LABELS[state.targetBasis], row.baseline, row.baselineCount, row.targetStart, row.targetFinal, row.totalProjection, row.increase]));
    appendSheet(wb, 'Metas', targetData);

    const evolutionData = [['Mês', ...targets.map(row => row.agency), 'Total consolidado']];
    for (let index = 0; index < state.projectionMonths; index += 1) {
      evolutionData.push([`Mês ${index + 1}`, ...targets.map(row => row.schedule[index]), sumOrNull(targets.map(row => row.schedule[index]))]);
    }
    appendSheet(wb, 'Evolução mensal', evolutionData);

    const commissionData = [[
      'Estudo', 'Assessoria', 'Indicador do gatilho', 'Condição', 'Carteira cancelada informada',
      'Percentual da carteira (%)', 'Meta fixa (R$)', 'Média histórica mensal', 'Meta de desbloqueio (R$)',
      'Patamar avaliado para comissão (R$)', 'Recuperação adicional (R$)', 'Comissão atual no patamar avaliado (R$)', 'Nova comissão no patamar avaliado (R$)',
      'Aumento da comissão pela nova tabela (R$)', 'Custo incremental vs. hoje (R$)', 'Saldo incremental após custo de comissão (R$)',
      'Meses de aging considerados', 'Composição CPF (%)', 'Composição CNPJ (%)'
    ]];
    const agingRows = [['Estudo', 'Assessoria', 'Faixa de aging', 'Participação histórica (%)', 'Recuperação estimada (R$)', 'Taxa atual B2C (%)', 'Taxa atual B2B (%)', 'Nova taxa B2C (%)', 'Nova taxa B2B (%)', 'Comissão atual estimada (R$)', 'Nova comissão estimada (R$)', 'Diferença (R$)']];
    commissionScenarios.forEach(result => {
      const scenario = result.scenario;
      const condition = scenario.targetMode === 'portfolio_percent' ? 'Percentual da carteira cancelada' : 'Valor fixo';
      commissionData.push([
        scenario.name, scenario.agency, BASIS_LABELS[scenario.basis], condition, scenario.portfolioValue,
        scenario.targetMode === 'portfolio_percent' ? scenario.targetPercent : null,
        scenario.targetMode === 'fixed_amount' ? scenario.targetAmount : null,
        result.baseline, result.target, result.evaluationValue, result.additionalRecovery,
        result.currentFeeAtTarget, result.newFeeAtTarget, result.extraCommissionForNewRates,
        result.additionalCommissionVsToday, result.netIncremental, result.profile?.includedMonths,
        result.profile ? result.profile.cpfShare * 100 : null, result.profile ? result.profile.cnpjShare * 100 : null
      ]);
      (result.agingRows || []).forEach(row => agingRows.push([
        scenario.name, scenario.agency, row.range, row.share * 100, row.allocated,
        row.currentB2C, row.currentB2B, row.newB2C, row.newB2B,
        row.currentFee, row.newFee, row.delta
      ]));
    });
    appendSheet(wb, 'Cenários de comissão', commissionData);
    const simulatedRates = [['Estudo', 'Assessoria', 'Faixa de aging', 'Mix histórico (%)', 'Taxa atual B2C (%)', 'Taxa atual B2B (%)', 'Nova taxa B2C (%)', 'Nova taxa B2B (%)']];
    commissionScenarios.forEach(result => {
      const scenario = result.scenario;
      state.policy.forEach(policyRow => {
        const key = policyRow.key || mapPolicyAgingBucket(policyRow.range);
        const rate = scenario.rates?.find(item => item.key === key);
        const share = result.profile?.totalRecovered > 0 ? (result.profile.totals[key] || 0) / result.profile.totalRecovered * 100 : null;
        simulatedRates.push([
          scenario.name, scenario.agency, policyRow.range, share,
          Number(policyRow.b2c || 0) * 100, Number(policyRow.b2b || 0) * 100,
          rate ? Number(rate.newB2C) : Number(policyRow.b2c || 0) * 100,
          rate ? Number(rate.newB2B) : Number(policyRow.b2b || 0) * 100
        ]);
      });
    });
    appendSheet(wb, 'Tabelas simuladas', simulatedRates);
    appendSheet(wb, 'Custo por aging', agingRows);

    const historyData = [['Assessoria', 'Mês', 'Parcial', 'Cobrança', 'Recuperado bruto', 'Recuperado ativo', 'Recuperado cancelado', 'Comissão', 'Conversão', 'Custo / recuperação']];
    filteredHistory.forEach(row => historyData.push([row.agency, row.monthName, row.partial ? 'Sim' : 'Não', row.collection, row.recovered, row.active, row.canceled, row.commission, Number.isFinite(row.conversion) ? row.conversion * 100 : null, Number.isFinite(row.costRatio) ? row.costRatio * 100 : null]));
    appendSheet(wb, 'Histórico filtrado', historyData);

    if (state.policy.length) appendSheet(wb, 'Política atual', [['Faixa de atraso', 'Comissão B2C (%)', 'Comissão B2B (%)'], ...state.policy.map(row => [row.range, row.b2c * 100, row.b2b * 100])]);
    return wb;
  }

  function appendSheet(workbook, name, rows) {
    const sheet = window.XLSX.utils.aoa_to_sheet(rows);
    sheet['!cols'] = (rows[0] || []).map((_, index) => ({ wch: Math.min(38, Math.max(14, ...rows.slice(0, 35).map(row => String(row[index] ?? '').length + 2))) }));
    window.XLSX.utils.book_append_sheet(workbook, sheet, name);
  }

  function exportStudyExcel() {
    try {
      const workbook = buildWorkbook();
      const dateTag = new Date().toISOString().slice(0, 10);
      window.XLSX.writeFile(workbook, `estudo-assessorias-${dateTag}.xlsx`);
      closeExportMenu();
      toast('Excel gerado com resumo, metas, cenários de comissão, custos por aging, política e histórico.');
    } catch (error) { toast(error.message || 'Não foi possível gerar o Excel.', 'error'); }
  }

  function buildPrintReport() {
    const { targets, commissionScenarios, filteredHistory, performance } = getExportRows();
    const totals = {
      start: sumOrNull(targets.map(row => row.targetStart)), final: sumOrNull(targets.map(row => row.targetFinal)),
      projection: sumOrNull(targets.map(row => row.totalProjection))
    };
    const period = `${MONTHS[state.startMonth - 1].name} a ${MONTHS[state.endMonth - 1].name}${state.year ? `/${state.year}` : ''}`;
    const evolutionRows = Array.from({ length: state.projectionMonths }, (_, monthIndex) => {
      const monthly = sumOrNull(targets.map(row => row.schedule[monthIndex]));
      return `<tr><td>Mês ${monthIndex + 1}</td><td>${money(monthly)}</td></tr>`;
    }).join('');
    const performanceRows = performance.map(group => `<tr>
      <td>${esc(group.agency)}</td><td>${group.monthCount}</td><td>${money(group.totalRecovered)}</td><td>${money(group.avgRecovered)}</td>
      <td>${money(group.activeCount ? group.totalActive : null)}</td><td>${money(group.avgActive)}</td><td>${money(group.canceledCount ? group.totalCanceled : null)}</td><td>${money(group.avgCanceled)}</td>
      <td>${money(group.totalCommission)}</td><td>${percent(group.effectiveRate)}</td>
    </tr>`).join('');
    const scenarioSummary = commissionScenarios.map(result => `<tr>
      <td>${esc(result.scenario.name)}</td><td>${esc(result.scenario.agency)}</td><td>${esc(BASIS_LABELS[result.scenario.basis])}</td>
      <td>${money(result.target)}</td><td>${money(result.evaluationValue)}</td><td>${money(result.additionalRecovery)}</td><td>${money(result.currentFeeAtTarget)}</td>
      <td>${money(result.newFeeAtTarget)}</td><td>${money(result.additionalCommissionVsToday)}</td><td>${money(result.netIncremental)}</td>
    </tr>`).join('');
    const scenarioDetails = commissionScenarios.map(result => {
      const scenario = result.scenario;
      const rateRows = state.policy.map(policyRow => {
        const key = policyRow.key || mapPolicyAgingBucket(policyRow.range);
        const rate = scenario.rates?.find(item => item.key === key);
        const share = result.profile?.totalRecovered > 0 ? (result.profile.totals[key] || 0) / result.profile.totalRecovered : null;
        return `<tr><td>${esc(policyRow.range)}</td><td>${percent(share, 1)}</td><td>${percentPoints(Number(policyRow.b2c || 0) * 100)}</td><td>${percentPoints(Number(policyRow.b2b || 0) * 100)}</td><td>${percentPoints(rate ? Number(rate.newB2C) : Number(policyRow.b2c || 0) * 100)}</td><td>${percentPoints(rate ? Number(rate.newB2B) : Number(policyRow.b2b || 0) * 100)}</td></tr>`;
      }).join('');
      const agingDetail = result.agingRows?.map(row => `<tr><td>${esc(row.range)}</td><td>${percent(row.share, 1)}</td><td>${money(row.allocated)}</td><td>${money(row.currentFee)}</td><td>${money(row.newFee)}</td><td>${money(row.delta)}</td></tr>`).join('') || '<tr><td colspan="6">Sem custo calculável por aging: confira a disponibilidade do histórico e da composição CPF/CNPJ.</td></tr>';
      return `<section class="print-subsection"><h3>${esc(scenario.name)} — ${esc(scenario.agency)}</h3>
        <p class="print-scenario-meta">Indicador: ${esc(BASIS_LABELS[scenario.basis])} · Meta: ${money(result.target)} · Patamar avaliado: ${money(result.evaluationValue)} · Recuperação adicional: ${money(result.additionalRecovery)} · Comissão atual: ${money(result.currentFeeAtTarget)} · Nova comissão: ${money(result.newFeeAtTarget)} · Custo incremental vs. hoje: ${money(result.additionalCommissionVsToday)} · Saldo incremental após custo de comissão: ${money(result.netIncremental)}</p>
        <h4>Tabela nova por aging</h4><table class="print-table"><thead><tr><th>Faixa</th><th>Mix histórico</th><th>Atual B2C</th><th>Atual B2B</th><th>Nova B2C</th><th>Nova B2B</th></tr></thead><tbody>${rateRows || '<tr><td colspan="6">Política por aging não importada.</td></tr>'}</tbody></table>
        <h4>Custo projetado por aging</h4><table class="print-table"><thead><tr><th>Faixa</th><th>Mix</th><th>Recuperação estimada</th><th>Comissão atual</th><th>Nova comissão</th><th>Diferença</th></tr></thead><tbody>${agingDetail}</tbody></table></section>`;
    }).join('');
    const policyRows = state.policy.map(row => `<tr><td>${esc(row.range)}</td><td>${percentPoints(row.b2c * 100)}</td><td>${percentPoints(row.b2b * 100)}</td></tr>`).join('');
    const historyRows = filteredHistory.map(row => `<tr><td>${esc(row.agency)}</td><td>${esc(row.monthName)}${row.partial ? ' · parcial' : ''}</td><td>${money(row.collection)}</td><td>${money(row.recovered)}</td><td>${money(row.active)}</td><td>${money(row.canceled)}</td><td>${money(row.commission)}</td><td>${percent(row.conversion)}</td><td>${percent(row.costRatio)}</td></tr>`).join('');
    $('#printReport').innerHTML = `
      <header class="print-header"><div><div class="print-kicker">SIMULADOR COBRANÇAS MASTER · ASSESSORIAS</div><h1>Simulador Cobranças Master</h1><p>Relatório gerado em ${new Date().toLocaleString('pt-BR')}</p></div><div class="print-period">${esc(period)}</div></header>
      <section class="print-meta"><div><span>Arquivo de origem</span><strong>${esc(state.fileName || 'Cenário manual')}</strong></div><div><span>Indicador da meta</span><strong>${esc(BASIS_LABELS[state.targetBasis])}</strong></div><div><span>Método</span><strong>${esc(MODE_LABELS[state.targetMode])}</strong></div><div><span>Filtro</span><strong>${esc(state.agencyFilter === 'ALL' ? 'Todas as assessorias' : state.agencyFilter)}</strong></div></section>
      <section class="print-kpis"><div><span>Meta total mês 1</span><strong>${money(totals.start)}</strong></div><div><span>Meta mês final</span><strong>${money(totals.final)}</strong></div><div><span>Total projetado</span><strong>${money(totals.projection)}</strong></div><div><span>Cenários de comissão</span><strong>${commissionScenarios.length}</strong></div></section>
      <h2>1. Visão geral por assessoria</h2><table class="print-table"><thead><tr><th>Assessoria</th><th>Meses</th><th>Bruto total</th><th>Média bruto</th><th>Ativo total</th><th>Média ativo</th><th>Cancelado total</th><th>Média cancelado</th><th>Comissão</th><th>Custo efetivo</th></tr></thead><tbody>${performanceRows || '<tr><td colspan="10">Sem histórico disponível.</td></tr>'}</tbody></table>
      <h2>2. Metas por assessoria</h2><table class="print-table"><thead><tr><th>Assessoria</th><th>Indicador</th><th>Referência</th><th>Meta mês 1</th><th>Meta mês final</th><th>Total projetado</th></tr></thead><tbody>${targets.map(row => `<tr><td>${esc(row.agency)}</td><td>${esc(BASIS_LABELS[state.targetBasis])}</td><td>${money(row.baseline)}</td><td>${money(row.targetStart)}</td><td>${money(row.targetFinal)}</td><td>${money(row.totalProjection)}</td></tr>`).join('') || '<tr><td colspan="6">Nenhuma meta configurada.</td></tr>'}</tbody></table>
      <h2>3. Evolução mensal consolidada</h2><div class="print-chart">${buildTargetEvolutionSvg(targets)}</div><table class="print-table"><thead><tr><th>Período</th><th>Meta total</th></tr></thead><tbody>${evolutionRows}</tbody></table>
      <h2>4. Comparação dos cenários de comissão</h2><table class="print-table"><thead><tr><th>Estudo</th><th>Assessoria</th><th>Indicador</th><th>Meta</th><th>Patamar avaliado</th><th>Recuperação adicional</th><th>Comissão atual</th><th>Nova comissão</th><th>Custo incremental</th><th>Saldo após comissão</th></tr></thead><tbody>${scenarioSummary || '<tr><td colspan="10">Nenhum cenário criado.</td></tr>'}</tbody></table>
      <h2>5. Memória de cálculo por aging de todos os cenários</h2>${scenarioDetails || '<p>Nenhum estudo de comissão salvo nesta sessão.</p>'}
      <h2>6. Política atual por aging</h2><table class="print-table"><thead><tr><th>Faixa de atraso</th><th>Comissão B2C</th><th>Comissão B2B</th></tr></thead><tbody>${policyRows || '<tr><td colspan="3">Política atual não importada.</td></tr>'}</tbody></table>
      <h2>7. Histórico filtrado utilizado</h2><table class="print-table"><thead><tr><th>Assessoria</th><th>Mês</th><th>Cobrança</th><th>Bruto</th><th>Ativo</th><th>Cancelado</th><th>Comissão</th><th>Conversão</th><th>Custo / recuperação</th></tr></thead><tbody>${historyRows || '<tr><td colspan="9">Sem histórico importado ou registros no recorte.</td></tr>'}</tbody></table>
      <footer class="print-footer"><strong>Premissas e limitações:</strong> os custos são estimados com base na distribuição histórica por aging e na composição geral CPF/CNPJ da assessoria. A simulação associa CPF a B2C e CNPJ a B2B como hipótese operacional, a validar com a regra interna. A planilha não cruza aging, CPF/CNPJ e situação ativa/cancelada por título e não informa o saldo da carteira cancelada por assessoria; para gatilhos percentuais, esse saldo é informado manualmente. O resultado é estimativa gerencial, não apuração individual por título. Meses parciais ${state.includePartial ? 'foram incluídos' : 'foram excluídos'}.</footer>`;
  }

  function exportStudyPdf() {
    buildPrintReport();
    closeExportMenu();
    toast('Na janela de impressão, selecione “Salvar como PDF”.');
    window.print();
  }

  function closeExportMenu() {
    $('#exportMenu').classList.add('hidden');
    $('#exportMenuToggle').setAttribute('aria-expanded', 'false');
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
    state.agingByAgency = {};
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
    initializeCommissionScenarios();
    $('#historyFilters').classList.add('hidden');
    $('#manualScenarioSetup').classList.add('hidden');
    $('#emptyState').classList.add('hidden');
    $('#workspace').classList.remove('hidden');
    $('#clearDataBtn').classList.remove('hidden');
    $('#dataContext').innerHTML = '<strong>Cenário manual sem histórico.</strong> Metas e comissões serão calculadas a partir dos valores informados; não há comparação histórica.';
    $('#statusFile').textContent = '';
    setStatus(`Cenário livre criado com ${state.agencies.length} assessoria(s). Informe as metas na tabela.`, 'success');
    renderAll();
    activateModule('assessorias');
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
      state.rows = []; state.policy = []; state.agingByAgency = {}; state.commissionScenarios = []; state.activeCommissionScenarioId = null; state.fileName = ''; state.agencies = []; state.manualTargets = {}; state.hasHistory = false;
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
    document.querySelectorAll('.module-button').forEach(button => button.addEventListener('click', () => activateModule(button.dataset.module)));
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

    $('#commissionScenarioSelect').addEventListener('change', event => { state.activeCommissionScenarioId = event.target.value; renderCommissionModule(); });
    $('#newCommissionScenarioBtn').addEventListener('click', createCommissionScenario);
    $('#deleteCommissionScenarioBtn').addEventListener('click', deleteCommissionScenario);
    $('#commissionScenarioName').addEventListener('change', event => {
      const scenario = getActiveCommissionScenario();
      if (!scenario) return;
      scenario.name = event.target.value.trim() || 'Estudo de comissão sem nome';
      renderCommissionScenarioSelect();
    });
    $('#commissionAgency').addEventListener('change', event => { const scenario = getActiveCommissionScenario(); if (scenario) { scenario.agency = event.target.value; renderCommissionModule(); } });
    $('#commissionTriggerBasis').addEventListener('change', event => { const scenario = getActiveCommissionScenario(); if (scenario) { scenario.basis = event.target.value; renderCommissionModule(); } });
    $('#commissionTargetMode').addEventListener('change', event => { const scenario = getActiveCommissionScenario(); if (scenario) { scenario.targetMode = event.target.value; renderCommissionModule(); } });
    $('#commissionPortfolioValue').addEventListener('change', event => { const scenario = getActiveCommissionScenario(); if (scenario) { scenario.portfolioValue = parseOptionalNumeric(event.target.value); renderCommissionScenarioMetrics(buildCommissionScenarioResult(scenario)); } });
    $('#commissionTargetPercent').addEventListener('change', event => { const scenario = getActiveCommissionScenario(); if (scenario) { scenario.targetPercent = clamp(event.target.value, 0, 100); renderCommissionScenarioMetrics(buildCommissionScenarioResult(scenario)); } });
    $('#commissionTargetAmount').addEventListener('change', event => { const scenario = getActiveCommissionScenario(); if (scenario) { scenario.targetAmount = parseOptionalNumeric(event.target.value); renderCommissionScenarioMetrics(buildCommissionScenarioResult(scenario)); } });
    $('#commissionRatesTableBody').addEventListener('change', updateCommissionRate);

    $('#historySearch').addEventListener('input', renderHistory);
    $('#exportMenuToggle').addEventListener('click', () => {
      const menu = $('#exportMenu');
      const open = menu.classList.contains('hidden');
      menu.classList.toggle('hidden', !open);
      $('#exportMenuToggle').setAttribute('aria-expanded', String(open));
    });
    $('#exportExcelOption').addEventListener('click', exportStudyExcel);
    $('#exportPdfOption').addEventListener('click', exportStudyPdf);
    document.addEventListener('click', event => {
      if (!event.target.closest('.export-menu-wrap')) closeExportMenu();
    });

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
  activateModule('assessorias');
  tryLoadLocalHistory();
})();
