<!DOCTYPE html>
<html lang="th">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title><H1>SYS-Type R Dashboard</H1></title>
  <link href="https://fonts.googleapis.com/css2?family=IBM+Plex+Sans+Thai:wght@300;400;500;600&display=swap" rel="stylesheet">
  <style>
    /* ─── CSS Variables ─────────────────────────────────────── */
    :root {
      --bg:          #f1f5f9;
      --card:        #ffffff;
      --text:        #1e293b;
      --muted:       #64748b;
      --border:      #e2e8f0;
      --hover:       #f8fafc;
      --primary:     #0ea5e9;
      --primary-dim: #e0f2fe;
      --danger:      #ef4444;
      --warning:     #f59e0b;
      --success:     #22c55e;
      --font:        'IBM Plex Sans Thai', sans-serif;
      --radius:      12px;
      --shadow:      0 1px 3px rgba(0,0,0,.08), 0 4px 16px rgba(0,0,0,.04);
    }
    [data-theme="dark"] {
      --bg:          #0f172a;
      --card:        #1e293b;
      --text:        #f1f5f9;
      --muted:       #94a3b8;
      --border:      #334155;
      --hover:       #263348;
      --primary-dim: #0c4a6e;
    }

    /* ─── Reset / Base ──────────────────────────────────────── */
    *, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }
    body  { font-family: var(--font); background: var(--bg); color: var(--text); padding: 24px 20px; transition: background .25s, color .25s; }
    a     { color: var(--primary); text-decoration: none; }

    /* ─── Layout ────────────────────────────────────────────── */
    .wrap  { max-width: 1280px; margin: 0 auto; }

    /* ─── Top Bar ───────────────────────────────────────────── */
    .topbar {
      display: flex; align-items: center; justify-content: space-between;
      flex-wrap: wrap; gap: 12px; margin-bottom: 24px;
    }
    .topbar-left  { display: flex; align-items: center; gap: 12px; }
    .logo         { font-size: 22px; font-weight: 700; color: var(--primary); letter-spacing: -.5px; }
    .logo span    { font-weight: 300; color: var(--muted); font-size: 14px; display: block; margin-top: 2px; }
    .topbar-right { display: flex; align-items: center; gap: 10px; }
    .updated-at   { font-size: 12px; color: var(--muted); }

    /* ─── Buttons ───────────────────────────────────────────── */
    .btn {
      display: inline-flex; align-items: center; gap: 6px;
      padding: 8px 14px; border-radius: 8px; font-size: 13px; font-family: var(--font);
      border: 1px solid var(--border); background: var(--card); color: var(--text);
      cursor: pointer; transition: background .15s, border-color .15s;
    }
    .btn:hover         { background: var(--hover); }
    .btn.btn-primary   { background: var(--primary); color: #fff; border-color: var(--primary); }
    .btn.btn-primary:hover { background: #0284c7; }
    .btn.active        { background: var(--primary); color: #fff; border-color: var(--primary); }
    .btn.btn-sm        { padding: 6px 10px; font-size: 12px; }
    .btn:disabled      { opacity: .5; cursor: not-allowed; }
    .spin              { display: inline-block; animation: rotate .7s linear infinite; }
    @keyframes rotate  { to { transform: rotate(360deg); } }

    /* ─── Summary Cards ─────────────────────────────────────── */
    .cards { display: grid; grid-template-columns: repeat(auto-fit, minmax(180px, 1fr)); gap: 16px; margin-bottom: 24px; }
    .card  {
      background: var(--card); padding: 20px 22px; border-radius: var(--radius);
      border: 1px solid var(--border); box-shadow: var(--shadow);
      border-top: 4px solid var(--border);
    }
    .card.c-total   { border-top-color: var(--muted); }
    .card.c-pending { border-top-color: var(--primary); }
    .card.c-warn    { border-top-color: var(--warning); }
    .card.c-danger  { border-top-color: var(--danger); }
    .card-label { font-size: 12px; color: var(--muted); margin-bottom: 8px; font-weight: 500; }
    .card-val   { font-size: 32px; font-weight: 700; line-height: 1; }

    /* ─── Toolbar ───────────────────────────────────────────── */
    .toolbar {
      display: flex; align-items: center; flex-wrap: wrap;
      gap: 10px; margin-bottom: 16px;
    }
    .filter-group { display: flex; gap: 6px; flex-wrap: wrap; }
    .search-wrap  { margin-left: auto; position: relative; }
    .search-wrap input {
      padding: 8px 12px 8px 34px; border-radius: 8px;
      border: 1px solid var(--border); background: var(--card);
      color: var(--text); font-family: var(--font); font-size: 13px;
      width: 220px; outline: none; transition: border-color .15s;
    }
    .search-wrap input:focus { border-color: var(--primary); }
    .search-wrap .search-icon {
      position: absolute; left: 10px; top: 50%; transform: translateY(-50%);
      color: var(--muted); pointer-events: none; font-size: 14px;
    }
    .row-count { font-size: 12px; color: var(--muted); white-space: nowrap; }

    /* ─── Table ─────────────────────────────────────────────── */
    .table-wrap {
      background: var(--card); border-radius: var(--radius);
      border: 1px solid var(--border); box-shadow: var(--shadow); overflow: auto;
    }
    table { width: 100%; border-collapse: collapse; min-width: 820px; }
    thead th {
      background: var(--hover); padding: 12px 14px; text-align: left;
      font-size: 12px; color: var(--muted); font-weight: 600;
      white-space: nowrap; user-select: none; cursor: pointer;
      border-bottom: 1px solid var(--border);
    }
    thead th:hover { color: var(--primary); }
    thead th .sort-arrow { margin-left: 4px; opacity: .4; }
    thead th.asc  .sort-arrow::after { content: "▲"; }
    thead th.desc .sort-arrow::after { content: "▼"; }
    thead th:not(.asc):not(.desc) .sort-arrow::after { content: "⇅"; }
    tbody tr { transition: background .1s; }
    tbody tr:hover { background: var(--hover); }
    tbody td { padding: 12px 14px; border-bottom: 1px solid var(--border); font-size: 13px; vertical-align: middle; }
    tbody tr:last-child td { border-bottom: none; }

    /* ─── Badges ────────────────────────────────────────────── */
    .badge {
      display: inline-block; padding: 3px 10px; border-radius: 99px;
      font-size: 11px; font-weight: 700; white-space: nowrap;
    }
    .badge-danger  { background: #fee2e2; color: #b91c1c; animation: blink 2s infinite; }
    .badge-warning { background: #fef3c7; color: #92400e; }
    .badge-normal  { background: #f1f5f9; color: #64748b; }
    @keyframes blink { 0%,100%{opacity:1} 50%{opacity:.6} }

    .status-pill {
      display: inline-block; padding: 3px 10px; border-radius: 99px;
      font-size: 11px; font-weight: 600; background: var(--primary-dim); color: var(--primary);
    }

    /* ─── Empty / Loading ───────────────────────────────────── */
    .empty {
      text-align: center; padding: 60px 20px; color: var(--muted);
    }
    .empty .emoji { font-size: 36px; display: block; margin-bottom: 10px; }
    .skeleton { background: linear-gradient(90deg, var(--border) 25%, var(--hover) 50%, var(--border) 75%); background-size: 400%; animation: shimmer 1.2s infinite; height: 16px; border-radius: 4px; }
    @keyframes shimmer { from{background-position:100%} to{background-position:-100%} }

    /* ─── Toast ─────────────────────────────────────────────── */
    #toast {
      position: fixed; bottom: 24px; right: 24px; z-index: 999;
      background: #1e293b; color: #fff; padding: 10px 18px;
      border-radius: 8px; font-size: 13px; box-shadow: 0 4px 20px rgba(0,0,0,.3);
      transform: translateY(80px); opacity: 0; transition: transform .3s, opacity .3s;
    }
    #toast.show { transform: translateY(0); opacity: 1; }

    /* ─── Responsive ────────────────────────────────────────── */
    @media (max-width: 600px) {
      body { padding: 16px 12px; }
      .search-wrap input { width: 160px; }
      .topbar { flex-direction: column; align-items: flex-start; }
    }
  </style>
</head>
<body>
<div class="wrap">

  <!-- Top Bar -->
  <div class="topbar">
    <div class="topbar-left">
      <div class="logo">📦 SYS-Type R Dashboard<span>Return Tracking Dashboard</span></div>
    </div>
    <div class="topbar-right">
      <span class="updated-at" id="updatedAt">–</span>
      <button class="btn btn-sm" id="themeBtn" onclick="toggleTheme()">🌙 Dark</button>
      <button class="btn btn-sm btn-primary" id="refreshBtn" onclick="refreshData()">⟳ Refresh</button>
    </div>
  </div>

  <!-- Summary Cards -->
  <div class="cards">
    <div class="card c-total">
      <div class="card-label">รายการทั้งหมด</div>
      <div class="card-val" id="totalCount">–</div>
    </div>
    <div class="card c-pending">
      <div class="card-label">รอดำเนินการ</div>
      <div class="card-val" id="waitingCount">–</div>
    </div>
    <div class="card c-warn">
      <div class="card-label">ใกล้กำหนด (7–13 วัน)</div>
      <div class="card-val" id="nearCount">–</div>
    </div>
    <div class="card c-danger">
      <div class="card-label">เกินกำหนด (14 วัน+)</div>
      <div class="card-val" id="overdueCount">–</div>
    </div>
  </div>

  <!-- Toolbar -->
  <div class="toolbar">
    <div class="filter-group">
      <button class="btn active" data-filter="pending" onclick="applyFilter('pending', this)">🕐 ค้างทั้งหมด</button>
      <button class="btn" data-filter="danger"  onclick="applyFilter('danger', this)">🚨 เกินกำหนด</button>
      <button class="btn" data-filter="warning" onclick="applyFilter('warning', this)">⚠️ ใกล้กำหนด</button>
      <button class="btn" data-filter="all"     onclick="applyFilter('all', this)">📋 ทั้งหมด</button>
    </div>
    <span class="row-count" id="rowCount"></span>
    <div class="search-wrap">
      <span class="search-icon">🔍</span>
      <input type="text" id="searchBox" placeholder="ค้นหา Order / Applicant / Item…" oninput="onSearch()">
    </div>
  </div>

  <!-- Table -->
  <div class="table-wrap">
    <table>
      <thead>
        <tr>
          <th onclick="sortTable('orderNo')"    data-col="orderNo">Order No.<span class="sort-arrow"></span></th>
          <th onclick="sortTable('applicant')"  data-col="applicant">Applicant<span class="sort-arrow"></span></th>
          <th onclick="sortTable('province')"   data-col="province">Province<span class="sort-arrow"></span></th>
          <th onclick="sortTable('domain')"     data-col="domain">Domain<span class="sort-arrow"></span></th>
          <th onclick="sortTable('itemName')"   data-col="itemName">Item Name<span class="sort-arrow"></span></th>
          <th onclick="sortTable('borrowDays')" data-col="borrowDays">Borrow Days<span class="sort-arrow"></span></th>
          <th onclick="sortTable('status')"     data-col="status">Status<span class="sort-arrow"></span></th>
        </tr>
      </thead>
      <tbody id="tableBody">
        <tr><td colspan="7"><div style="padding:60px;text-align:center;color:var(--muted)">กำลังโหลด…</div></td></tr>
      </tbody>
    </table>
  </div>

</div><!-- /wrap -->

<div id="toast"></div>

<script>
  /* ── State ─────────────────────────────────────────────── */
  var rawData    = [];
  var activeFilter = 'pending';
  var searchText   = '';
  var sortCol      = 'borrowDays';
  var sortDir      = 'desc'; // asc | desc

  /* ── Init ───────────────────────────────────────────────── */
  window.onload = function() {
    // Restore theme
    if (localStorage.getItem('theme') === 'dark') applyTheme('dark');
    loadData(false);
  };

  /* ── Data Loading ───────────────────────────────────────── */
  function loadData(forceRefresh) {
    setRefreshing(true);
    var fn = forceRefresh ? 'clearCache' : 'getDashboardData';
    google.script.run
      .withSuccessHandler(onDataLoaded)
      .withFailureHandler(function(err) {
        setRefreshing(false);
        showToast('❌ Error: ' + err.message);
        document.getElementById('tableBody').innerHTML =
          '<tr><td colspan="7"><div class="empty"><span class="emoji">⚠️</span>' + err.message + '</div></td></tr>';
      })
      [fn]();
  }

  function refreshData() {
    loadData(true);
  }

  function onDataLoaded(res) {
    setRefreshing(false);
    if (res.status === 'error') {
      showToast('❌ ' + res.message);
      return;
    }
    rawData = res.recentOrders;
    document.getElementById('totalCount').textContent   = res.summary.totalAll;
    document.getElementById('waitingCount').textContent = res.summary.totalPending;
    document.getElementById('nearCount').textContent    = res.summary.nearDeadline;
    document.getElementById('overdueCount').textContent = res.summary.overdue;
    document.getElementById('updatedAt').textContent    = 'อัปเดต ' + res.updatedAt;
    renderTable();
  }

  /* ── Filter / Search / Sort ─────────────────────────────── */
  function applyFilter(type, btn) {
    activeFilter = type;
    document.querySelectorAll('.filter-group .btn').forEach(function(b) { b.classList.remove('active'); });
    if (btn) btn.classList.add('active');
    renderTable();
  }

  function onSearch() {
    searchText = document.getElementById('searchBox').value.toLowerCase().trim();
    renderTable();
  }

  function sortTable(col) {
    if (sortCol === col) {
      sortDir = sortDir === 'asc' ? 'desc' : 'asc';
    } else {
      sortCol = col;
      sortDir = col === 'borrowDays' ? 'desc' : 'asc';
    }
    // Update header arrows
    document.querySelectorAll('thead th').forEach(function(th) {
      th.classList.remove('asc', 'desc');
      if (th.dataset.col === sortCol) th.classList.add(sortDir);
    });
    renderTable();
  }

  function getFilteredData() {
    var data = rawData.slice();

    // Filter
    if (activeFilter === 'pending') {
      data = data.filter(function(r) { return r.isPending; });
    } else if (activeFilter === 'danger') {
      data = data.filter(function(r) { return r.alert === 'danger' && r.isPending; });
    } else if (activeFilter === 'warning') {
      data = data.filter(function(r) { return r.alert === 'warning' && r.isPending; });
    }
    // 'all' → no filter

    // Search
    if (searchText) {
      data = data.filter(function(r) {
        return (r.orderNo    + r.applicant + r.itemName +
                r.province   + r.domain   + r.status).toLowerCase().indexOf(searchText) > -1;
      });
    }

    // Sort
    data.sort(function(a, b) {
      var va = sortCol === 'borrowDays' ? parseFloat(a[sortCol]) : (a[sortCol] || '').toLowerCase();
      var vb = sortCol === 'borrowDays' ? parseFloat(b[sortCol]) : (b[sortCol] || '').toLowerCase();
      if (va < vb) return sortDir === 'asc' ? -1 : 1;
      if (va > vb) return sortDir === 'asc' ?  1 : -1;
      return 0;
    });

    return data;
  }

  /* ── Render ─────────────────────────────────────────────── */
  function renderTable() {
    var data = getFilteredData();
    var tbody = document.getElementById('tableBody');
    document.getElementById('rowCount').textContent = data.length + ' รายการ';

    if (data.length === 0) {
      tbody.innerHTML = '<tr><td colspan="7"><div class="empty"><span class="emoji">❌</span>ไม่พบรายการที่ตรงกับเงื่อนไข ลองใช้ Order NO. หรือตรวจสอบข้อมูลให้ถูกต้องครับ </div></td></tr>';
      return;
    }

    var html = '';
    data.forEach(function(r) {
      var bc  = r.alert === 'danger' ? 'badge-danger' : r.alert === 'warning' ? 'badge-warning' : 'badge-normal';
      var days = parseFloat(r.borrowDays);
      var daysStr = isNaN(days) ? r.borrowDays : (days % 1 === 0 ? days.toFixed(0) : r.borrowDays);
      html += '<tr>' +
        '<td><b style="font-size:12px">' + esc(r.orderNo) + '</b></td>' +
        '<td>' + esc(r.applicant) + '</td>' +
        '<td><span style="font-size:12px;color:var(--muted)">' + esc(r.province) + '</span></td>' +
        '<td><span style="font-size:12px">' + esc(r.domain) + '</span></td>' +
        '<td style="max-width:280px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap" title="' + esc(r.itemName) + '">' + esc(r.itemName) + '</td>' +
        '<td><span class="badge ' + bc + '">' + daysStr + ' วัน</span></td>' +
        '<td><span class="status-pill">' + esc(r.status) + '</span></td>' +
        '</tr>';
    });
    tbody.innerHTML = html;
  }

  /* ── Utils ──────────────────────────────────────────────── */
  function esc(str) {
    return (str || '').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
  }

  function setRefreshing(on) {
    var btn = document.getElementById('refreshBtn');
    btn.disabled = on;
    btn.innerHTML = on ? '<span class="spin">⟳</span> Loading…' : '⟳ Refresh';
  }

  function showToast(msg) {
    var t = document.getElementById('toast');
    t.textContent = msg;
    t.classList.add('show');
    setTimeout(function() { t.classList.remove('show'); }, 3000);
  }

  /* ── Theme ──────────────────────────────────────────────── */
  function toggleTheme() {
    var cur = document.documentElement.getAttribute('data-theme') === 'dark' ? 'light' : 'dark';
    applyTheme(cur);
    localStorage.setItem('theme', cur);
  }

  function applyTheme(t) {
    if (t === 'dark') {
      document.documentElement.setAttribute('data-theme', 'dark');
      document.getElementById('themeBtn').textContent = '☀️ Light';
    } else {
      document.documentElement.removeAttribute('data-theme');
      document.getElementById('themeBtn').textContent = '🌙 Dark';
    }
  }

  /* ── Auto-refresh every 5 min ───────────────────────────── */
  setInterval(function() { loadData(false); }, 5 * 60 * 1000);
</script>
</body>
</html>