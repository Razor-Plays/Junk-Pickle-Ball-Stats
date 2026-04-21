(() => {
  const SHEET_ID = '1pe2SvzjhPTlWaRfMWkuXmErbjPWHGIEraELYhZxgAHA';
  const BUILD_TAG = '2026-04-21.2';
  const CSV_URL = (sheetName) => `https://docs.google.com/spreadsheets/d/${SHEET_ID}/gviz/tq?tqx=out:csv&sheet=${encodeURIComponent(sheetName)}`;
  const SNAPSHOT_ATTENDANCE_URL = './data/attendance_export_snapshot.csv';
  const SNAPSHOT_META_URL = './data/source_meta_snapshot.csv';

  const GUEST_REGEX = /^guest\s+\d+$/i;
  const PLACEHOLDER_NAMES = new Set(['', '-', 'na', 'n/a', 'null', 'none', 'unknown', 'tbd']);

  const state = {
    rawRows: [],
    sessions: [],
    players: [],
    syncLabel: '—',
    dataSource: 'loading',
    charts: {},
  };

  const $ = (s) => document.querySelector(s);

  function showLoading(show) {
    $('#loadingState').classList.toggle('show', show);
  }

  function showError(message = '') {
    const el = $('#errorState');
    el.textContent = message;
    el.classList.toggle('show', Boolean(message));
  }

  function setDataSourceStatus(mode, detail = '') {
    state.dataSource = mode;
    const el = $('#dataSourceStatus');
    el.classList.remove('live', 'snapshot', 'error');
    if (mode === 'loading') {
      el.textContent = 'Data source: loading…';
      return;
    }
    if (mode === 'live') {
      el.classList.add('live');
      el.textContent = 'Data source: Live Google Sheet';
      return;
    }
    if (mode === 'snapshot') {
      el.classList.add('snapshot');
      el.textContent = detail ? `Data source: Local snapshot fallback (${detail})` : 'Data source: Local snapshot fallback';
      return;
    }
    el.classList.add('error');
    el.textContent = 'Data source: Error loading data';
  }

  function parseCsv(url) {
    return new Promise((resolve, reject) => {
      Papa.parse(url, {
        download: true,
        header: true,
        skipEmptyLines: true,
        complete: (result) => {
          if (result.errors?.length) {
            reject(new Error(result.errors[0].message || 'CSV parse error'));
            return;
          }
          resolve(result.data || []);
        },
        error: (err) => reject(err),
      });
    });
  }

  function pickColumn(row, candidates) {
    const entries = Object.entries(row);
    for (const c of candidates) {
      const found = entries.find(([k]) => k.trim().toLowerCase() === c);
      if (found) return found[1];
    }
    return '';
  }

  function cleanPlayerName(value) {
    const n = String(value || '').trim();
    if (PLACEHOLDER_NAMES.has(n.toLowerCase())) return '';
    if (GUEST_REGEX.test(n)) return '';
    return n;
  }

  function toIsoDate(value) {
    const d = new Date(value);
    if (Number.isNaN(d.getTime())) return '';
    return d.toISOString().slice(0, 10);
  }

  function formatDate(value) {
    if (!value) return '—';
    const d = new Date(value);
    if (Number.isNaN(d.getTime())) return String(value);
    return d.toLocaleDateString(undefined, { year: 'numeric', month: 'short', day: '2-digit' });
  }

  const pct = (n) => `${(n * 100).toFixed(1)}%`;

  function formatNames(names) {
    if (!names.length) return '—';
    if (names.length === 1) return names[0];
    if (names.length === 2) return `${names[0]} & ${names[1]}`;
    return `${names.slice(0, -1).join(', ')} & ${names[names.length - 1]}`;
  }

  function topBy(players, getValue, { minValue = -Infinity, maxLeaders = 2 } = {}) {
    if (!players.length) return { leaders: [], value: null };
    const ranked = [...players].sort((a, b) => getValue(b) - getValue(a));
    const best = getValue(ranked[0]);
    if (best == null || best < minValue) return { leaders: [], value: null };
    const leaders = ranked.filter((p) => getValue(p) === best).slice(0, maxLeaders);
    return { leaders, value: best };
  }

  function gradient(ctx) {
    const g = ctx.createLinearGradient(0, 0, 0, 280);
    g.addColorStop(0, 'rgba(34,197,94,0.95)');
    g.addColorStop(1, 'rgba(134,239,172,0.25)');
    return g;
  }

  function destroyCharts() {
    Object.values(state.charts).forEach((ch) => ch?.destroy());
    state.charts = {};
  }

  function inferRows(attendanceRows) {
    const normalized = [];

    for (const row of attendanceRows) {
      const sessionDateRaw = pickColumn(row, [
        'session_date',
        'date',
        'attendance_date',
        'session day',
      ]);
      const playerRaw = pickColumn(row, [
        'player_name',
        'player',
        'name',
        'participant_name',
        'participant',
      ]);
      const sessionIdRaw = pickColumn(row, ['session_id', 'session_key', 'session']);

      const sessionDate = toIsoDate(sessionDateRaw);
      const playerName = cleanPlayerName(playerRaw);
      if (!sessionDate || !playerName) continue;

      normalized.push({
        sessionDate,
        sessionId: String(sessionIdRaw || '').trim(),
        playerName,
      });
    }

    return normalized;
  }

  function buildSessions(rows, fromDate, toDate) {
    const map = new Map();

    for (const r of rows) {
      if (fromDate && r.sessionDate < fromDate) continue;
      if (toDate && r.sessionDate > toDate) continue;

      const key = r.sessionId ? `${r.sessionDate}__${r.sessionId}` : r.sessionDate;
      if (!map.has(key)) {
        map.set(key, {
          key,
          date: r.sessionDate,
          players: new Set(),
        });
      }
      map.get(key).players.add(r.playerName);
    }

    return [...map.values()]
      .map((s) => ({ ...s, players: [...s.players].sort((a, b) => a.localeCompare(b)) }))
      .sort((a, b) => (a.date === b.date ? a.key.localeCompare(b.key) : a.date.localeCompare(b.date)));
  }

  function buildPlayerStats(sessions) {
    const totalSessions = sessions.length;
    const playerIndices = new Map();

    sessions.forEach((session, idx) => {
      for (const p of session.players) {
        if (!playerIndices.has(p)) playerIndices.set(p, []);
        playerIndices.get(p).push(idx);
      }
    });

    const players = [];

    for (const [name, indices] of playerIndices.entries()) {
      const attendedSet = new Set(indices);
      const firstIdx = indices[0];
      const lastIdx = indices[indices.length - 1];

      const eligibleSessions = Math.max(1, totalSessions - firstIdx);
      const totalAttended = indices.length;
      const attendanceRate = totalAttended / eligibleSessions;

      let longestStreak = 0;
      let run = 0;
      for (let i = firstIdx; i < totalSessions; i += 1) {
        if (attendedSet.has(i)) {
          run += 1;
          if (run > longestStreak) longestStreak = run;
        } else {
          run = 0;
        }
      }

      let currentStreak = 0;
      for (let i = totalSessions - 1; i >= 0; i -= 1) {
        if (attendedSet.has(i)) currentStreak += 1;
        else break;
      }

      const start = Math.max(0, totalSessions - 10);
      let last10 = 0;
      for (let i = start; i < totalSessions; i += 1) {
        if (attendedSet.has(i)) last10 += 1;
      }

      const prevStart = Math.max(0, totalSessions - 20);
      const prevEnd = Math.max(0, totalSessions - 10);
      let prev10 = 0;
      for (let i = prevStart; i < prevEnd; i += 1) {
        if (attendedSet.has(i)) prev10 += 1;
      }

      players.push({
        playerName: name,
        totalAttended,
        firstAttendanceDate: sessions[firstIdx]?.date || '',
        lastAttendanceDate: sessions[lastIdx]?.date || '',
        attendanceRate,
        currentStreak,
        longestStreak,
        last10,
        prev10,
        improveDelta: last10 - prev10,
      });
    }

    players.sort((a, b) => b.totalAttended - a.totalAttended || a.playerName.localeCompare(b.playerName));
    return players;
  }

  function readLastSyncLabel(metaRows) {
    if (!metaRows?.length) return '—';

    const candidates = ['last_sync_date', 'last_sync', 'sync_date', 'updated_at', 'timestamp'];

    for (const row of metaRows) {
      for (const [k, v] of Object.entries(row)) {
        const key = k.trim().toLowerCase();
        if (candidates.includes(key) && String(v || '').trim()) {
          const parsed = toIsoDate(v);
          return parsed ? formatDate(parsed) : String(v);
        }
      }

      const keyField = pickColumn(row, ['key', 'name', 'metric']);
      const valueField = pickColumn(row, ['value', 'metric_value']);
      if (String(keyField).trim().toLowerCase().includes('sync') && String(valueField).trim()) {
        const parsed = toIsoDate(valueField);
        return parsed ? formatDate(parsed) : String(valueField);
      }
    }

    return '—';
  }

  function renderKpis(sessions, players, syncLabel) {
    $('#kpiSessions').textContent = String(sessions.length);
    $('#kpiPlayers').textContent = String(players.length);
    const avg = sessions.length ? (players.reduce((acc, p) => acc + p.totalAttended, 0) / sessions.length) : 0;
    $('#kpiAvg').textContent = avg.toFixed(1);
    $('#kpiSync').textContent = syncLabel || '—';
  }

  function renderHighlights(players, sessions) {
    const el = $('#highlightsGrid');
    const recentWindow = Math.min(10, sessions.length);
    const regulars = players.filter((p) => p.totalAttended >= 8);
    const latestIndex = sessions.length - 1;

    const streak = topBy(players, (p) => p.currentStreak, { minValue: 1 });
    const recent = topBy(players, (p) => p.last10, { minValue: 1 });
    const rate = topBy(regulars, (p) => Number((p.attendanceRate * 100).toFixed(1)), { minValue: 1 });

    // Assumption: comeback watch = meaningful prior contributor (8+ sessions) absent for last 6 sessions.
    const comebackPool = players
      .map((p) => ({
        ...p,
        lastSeenIdx: sessions.findIndex((s) => s.date === p.lastAttendanceDate),
      }))
      .filter((p) => p.totalAttended >= 8 && p.lastSeenIdx >= 0 && (latestIndex - p.lastSeenIdx) >= 6)
      .sort((a, b) => b.totalAttended - a.totalAttended || b.longestStreak - a.longestStreak);
    const comeback = comebackPool[0];

    const cards = [];
    if (streak.leaders.length) {
      cards.push({
        title: '🔥 Hot Streak',
        main: `${formatNames(streak.leaders.map((p) => p.playerName))} — ${streak.value} straight`,
        sub: 'Carrying the best active run right now.',
      });
    }
    if (recent.leaders.length) {
      cards.push({
        title: '🎯 Most Active (Last 10)',
        main: `${formatNames(recent.leaders.map((p) => p.playerName))} — ${recent.value}/${recentWindow}`,
        sub: 'Setting the pace in recent sessions.',
      });
    }
    if (rate.leaders.length) {
      cards.push({
        title: '📈 Best Attendance Rate',
        main: `${formatNames(rate.leaders.map((p) => p.playerName))} — ${rate.value.toFixed(1)}%`,
        sub: 'Best rate among established regulars.',
      });
    }
    if (comeback) {
      cards.push({
        title: '👀 Comeback Watch',
        main: `${comeback.playerName} — ${comeback.totalAttended} total sessions`,
        sub: `Proven presence; last seen ${formatDate(comeback.lastAttendanceDate)}.`,
      });
    }

    el.innerHTML = cards.slice(0, 4).map((c) => `
      <article class="story-card">
        <div class="story-title">${c.title}</div>
        <div class="story-main">${c.main}</div>
        <div class="story-sub">${c.sub}</div>
      </article>
    `).join('');
  }

  function renderAwards(players, sessions) {
    const el = $('#awardsGrid');
    const recentWindow = Math.min(10, sessions.length);
    const hasHistory20 = sessions.length >= 20;
    const regulars = players.filter((p) => p.totalAttended >= 8);

    const mostRegular = topBy(players, (p) => p.totalAttended, { minValue: 1 });
    const mostConsistent = topBy(regulars, (p) => Number((p.attendanceRate * 100).toFixed(1)), { minValue: 1 });
    const ironStreak = topBy(players, (p) => p.longestStreak, { minValue: 1 });
    const hotStreak = topBy(players, (p) => p.currentStreak, { minValue: 1 });
    const improvedPool = hasHistory20 ? players.filter((p) => p.improveDelta > 0) : [];
    const mostImproved = topBy(improvedPool, (p) => p.improveDelta, { minValue: 1 });
    const maxTotal = Math.max(1, ...players.map((p) => p.totalAttended));
    // Assumption: reliable regular combines strong volume + consistency, min 12 sessions, avoid duplicating Most Regular if possible.
    const reliablePool = players
      .filter((p) => p.totalAttended >= 12)
      .map((p) => ({
        ...p,
        reliableScore: 0.6 * p.attendanceRate + 0.4 * (p.totalAttended / maxTotal),
      }))
      .sort((a, b) => b.reliableScore - a.reliableScore);
    const mostRegularName = mostRegular.leaders[0]?.playerName;
    const reliableWinner = reliablePool.find((p) => p.playerName !== mostRegularName) || reliablePool[0];

    const awards = [
      {
        name: 'Most Regular Player',
        winner: formatNames(mostRegular.leaders.map((p) => p.playerName)),
        stat: `${mostRegular.value || 0} sessions`,
        note: 'Still setting the pace.',
      },
      {
        name: 'Most Consistent Presence',
        winner: formatNames(mostConsistent.leaders.map((p) => p.playerName)),
        stat: `${mostConsistent.value ? mostConsistent.value.toFixed(1) : '0.0'}% rate`,
        note: 'Best rate among regulars (8+ sessions).',
      },
      {
        name: 'Iron Streak',
        winner: formatNames(ironStreak.leaders.map((p) => p.playerName)),
        stat: `${ironStreak.value || 0} in a row`,
        note: 'Longest all-time consecutive run.',
      },
      {
        name: 'Current Hot Streak',
        winner: formatNames(hotStreak.leaders.map((p) => p.playerName)),
        stat: `${hotStreak.value || 0} straight`,
        note: 'Active streak through the latest session.',
      },
      {
        name: 'Most Improved',
        winner: formatNames(mostImproved.leaders.map((p) => p.playerName)),
        stat: `${mostImproved.value ? `+${mostImproved.value}` : '+0'} vs previous 10`,
        note: 'Best jump from previous 10 to last 10.',
      },
      {
        name: 'Reliable Regular',
        winner: reliableWinner?.playerName || '—',
        stat: reliableWinner ? `${pct(reliableWinner.attendanceRate)} rate • ${reliableWinner.totalAttended} sessions` : '—',
        note: 'Trusted volume and consistency over time.',
      },
    ];

    el.innerHTML = awards.map((a) => `
      <article class="award-card">
        <div class="award-name">${a.name}</div>
        <div class="award-winner">${a.winner}</div>
        <div class="award-stat">${a.stat}</div>
        <div class="award-note">${a.note}</div>
      </article>
    `).join('');
  }

  function renderAllTimeGreats(players) {
    const el = $('#allTimeGrid');
    const streakLegend = topBy(players, (p) => p.longestStreak, { minValue: 1 });
    const grinder = topBy(players, (p) => p.totalAttended, { minValue: 1 });
    const core5 = [...players].sort((a, b) => b.totalAttended - a.totalAttended).slice(0, 5);

    el.innerHTML = `
      <article class="legacy-card">
        <div class="legacy-title">Iron Legend</div>
        <div class="legacy-name">${formatNames(streakLegend.leaders.map((p) => p.playerName))}</div>
        <div class="legacy-sub">${streakLegend.value || 0} straight — longest run in group history.</div>
      </article>
      <article class="legacy-card">
        <div class="legacy-title">The Grinder</div>
        <div class="legacy-name">${formatNames(grinder.leaders.map((p) => p.playerName))}</div>
        <div class="legacy-sub">${grinder.value || 0} total sessions — always in the mix.</div>
      </article>
      <article class="core-card">
        <div class="legacy-title">Active Core 5</div>
        <div class="legacy-name">The backbone of the group</div>
        <ol class="core-list">
          ${core5.map((p, i) => `<li><span>${i + 1}. ${p.playerName}</span><strong>${p.totalAttended}</strong></li>`).join('')}
        </ol>
      </article>
    `;
  }

  function renderCharts(players, sessions) {
    destroyCharts();

    Chart.defaults.color = '#9ca3af';
    Chart.defaults.borderColor = '#334155';

    const leaders = [...players].sort((a, b) => b.totalAttended - a.totalAttended).slice(0, 12);
    const fairness = [...players]
      .filter((p) => p.totalAttended >= 5)
      .sort((a, b) => b.attendanceRate - a.attendanceRate)
      .slice(0, 12);
    const form = [...players].sort((a, b) => b.last10 - a.last10 || b.totalAttended - a.totalAttended).slice(0, 12);
    const dowCounts = new Map([['Sun', 0], ['Mon', 0], ['Tue', 0], ['Wed', 0], ['Thu', 0], ['Fri', 0], ['Sat', 0]]);
    sessions.forEach((s) => {
      const key = new Date(s.date).toLocaleDateString('en-US', { weekday: 'short', timeZone: 'UTC' });
      dowCounts.set(key, (dowCounts.get(key) || 0) + 1);
    });

    const leadersCtx = document.getElementById('leadersChart').getContext('2d');
    const rateCtx = document.getElementById('rateChart').getContext('2d');
    const formCtx = document.getElementById('formChart').getContext('2d');
    const dowCtx = document.getElementById('dowChart').getContext('2d');

    state.charts.leaders = new Chart(leadersCtx, {
      type: 'bar',
      data: {
        labels: leaders.map((x) => x.playerName),
        datasets: [{
          label: 'Sessions Attended',
          data: leaders.map((x) => x.totalAttended),
          backgroundColor: gradient(leadersCtx),
          borderColor: '#34c77b',
          borderWidth: 1,
        }],
      },
      options: {
        indexAxis: 'y',
        responsive: true,
        maintainAspectRatio: false,
        scales: { x: { beginAtZero: true, ticks: { precision: 0 } } },
      },
    });

    state.charts.rate = new Chart(rateCtx, {
      type: 'bar',
      data: {
        labels: fairness.map((x) => x.playerName),
        datasets: [{
          label: 'Attendance Rate %',
          data: fairness.map((x) => +(x.attendanceRate * 100).toFixed(1)),
          backgroundColor: gradient(rateCtx),
          borderColor: '#34c77b',
          borderWidth: 1,
        }],
      },
      options: {
        indexAxis: 'y',
        responsive: true,
        maintainAspectRatio: false,
        scales: { x: { beginAtZero: true, max: 100 } },
      },
    });

    state.charts.form = new Chart(formCtx, {
      type: 'bar',
      data: {
        labels: form.map((x) => x.playerName),
        datasets: [{
          label: 'Last 10 Sessions',
          data: form.map((x) => x.last10),
          backgroundColor: '#60a5fa99',
          borderColor: '#60a5fa',
          borderWidth: 1,
        }],
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        scales: { x: { beginAtZero: true, max: 10, ticks: { precision: 0 } } },
      },
    });

    state.charts.dow = new Chart(dowCtx, {
      type: 'doughnut',
      data: {
        labels: [...dowCounts.keys()],
        datasets: [{
          data: [...dowCounts.values()],
          backgroundColor: ['#60a5fa', '#34c77b', '#f6c453', '#fb923c', '#818cf8', '#22d3ee', '#f472b6'],
          borderColor: '#1b2433',
          borderWidth: 2,
        }],
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: { legend: { position: 'bottom' } },
      },
    });
  }

  function renderTable(players) {
    const body = $('#playerTableBody');
    body.innerHTML = '';

    for (const p of players) {
      const tr = document.createElement('tr');
      tr.innerHTML = `
        <td>${p.playerName}</td>
        <td>${p.totalAttended}</td>
        <td>${(p.attendanceRate * 100).toFixed(1)}%</td>
        <td>${formatDate(p.lastAttendanceDate)}</td>
        <td>${p.currentStreak}</td>
        <td>${p.longestStreak}</td>
      `;
      body.appendChild(tr);
    }
  }

  async function loadSheetsData() {
    try {
      const [attendanceRows, metaRows] = await Promise.all([
        parseCsv(CSV_URL('attendance_export')),
        parseCsv(CSV_URL('source_meta')).catch(() => []),
      ]);
      return { attendanceRows, metaRows, sourceMode: 'live' };
    } catch (liveErr) {
      const [attendanceRows, metaRows] = await Promise.all([
        parseCsv(SNAPSHOT_ATTENDANCE_URL),
        parseCsv(SNAPSHOT_META_URL).catch(() => []),
      ]);
      return {
        attendanceRows,
        metaRows,
        sourceMode: 'snapshot',
        sourceDetail: liveErr?.message ? `live failed: ${liveErr.message}` : 'live fetch failed',
      };
    }
  }

  function getDateFilters() {
    const from = $('#fromDate').value || '';
    const to = $('#toDate').value || '';
    return { from, to };
  }

  function applyAndRender() {
    const { from, to } = getDateFilters();
    state.sessions = buildSessions(state.rawRows, from, to);
    state.players = buildPlayerStats(state.sessions);

    renderKpis(state.sessions, state.players, state.syncLabel);
    renderHighlights(state.players, state.sessions);
    renderAwards(state.players, state.sessions);
    renderAllTimeGreats(state.players);
    renderCharts(state.players, state.sessions);
    renderTable(state.players);
  }

  async function refresh() {
    showError('');
    showLoading(true);
    setDataSourceStatus('loading');

    try {
      const { attendanceRows, metaRows, sourceMode, sourceDetail } = await loadSheetsData();
      state.rawRows = inferRows(attendanceRows);
      state.syncLabel = readLastSyncLabel(metaRows);
      setDataSourceStatus(sourceMode, sourceDetail);

      if (!state.rawRows.length) {
        throw new Error('No usable rows found in attendance_export. Verify public sheet + expected columns (session_date/player_name).');
      }

      const dates = [...new Set(state.rawRows.map((r) => r.sessionDate))].sort((a, b) => a.localeCompare(b));
      if (dates.length) {
        if (!$('#fromDate').value) $('#fromDate').value = dates[0];
        if (!$('#toDate').value) $('#toDate').value = dates[dates.length - 1];
      }

      applyAndRender();
    } catch (err) {
      showError(`Unable to load dashboard data: ${err.message || err}`);
      setDataSourceStatus('error');
      destroyCharts();
      $('#playerTableBody').innerHTML = '';
      $('#highlightsGrid').innerHTML = '';
      $('#awardsGrid').innerHTML = '';
      $('#allTimeGrid').innerHTML = '';
      ['#kpiSessions', '#kpiPlayers', '#kpiAvg'].forEach((sel) => { $(sel).textContent = '—'; });
      $('#kpiSync').textContent = state.syncLabel || '—';
    } finally {
      showLoading(false);
    }
  }

  function boot() {
    $('#buildTag').textContent = `Build: ${BUILD_TAG}`;
    $('#refreshBtn').addEventListener('click', refresh);
    $('#fromDate').addEventListener('change', applyAndRender);
    $('#toDate').addEventListener('change', applyAndRender);
    refresh();
  }

  boot();
})();
