let allData              = [];
let filteredPeserta      = [];
let currentView          = 'overview';
let prevView             = 'peserta';
let charts               = {};
let currentPage          = 1;
const PAGE_SIZE          = 50;

// Multi-select filter for Per Jabatan
let activeJenisFilters     = new Set();
// Multi-select filter for Lulus-per-Jenis chart
let lulusJenisActiveFilters = new Set();

// Pendidikan view state
let activePendLevel        = 'semua';
let activePendidikanItem   = null;

// Column name aliases (for XLSX upload)
const COL = {
  NAMA:       ['nama', 'nama peserta', 'name'],
  JABATAN:    ['jabatan', 'jabatan_formasi', 'jabatan formasi', 'formasi'],
  PENDIDIKAN: ['pendidikan', 'education', 'pendidikan terakhir'],
  TAHUN:      ['tahunskd', 'tahun skd', 'tahun', 'year'],
  INSTANSI:   ['instansi'],
  LOKASI:     ['lokasi', 'unit_kerja', 'unit kerja', 'unit'],
  JENIS:      ['jenis', 'jenis_formasi', 'jenis formasi'],
  TWK:        ['twk', 'nilai_twk', 'nilai twk'],
  TIU:        ['tiu', 'nilai_tiu', 'nilai tiu'],
  TKP:        ['tkp', 'nilai_tkp', 'nilai tkp'],
  TOTAL:      ['total', 'nilai_skd', 'nilai skd', 'skd', 'nilai total'],
  KETERANGAN: ['keterangan', 'status', 'hasil'],
  NOPES:      ['nopes', 'no peserta', 'nomor peserta'],
  SYARAT:     ['syarat', 'syarat jabatan', 'kualifikasi', 'persyaratan'],
};

// ==========================================
// INIT – Load from dataSKD.js on start
// ==========================================
document.addEventListener('DOMContentLoaded', () => {
  if (typeof dataSKD !== 'undefined' && Array.isArray(dataSKD) && dataSKD.length > 0) {
    allData = transformDataSKD(dataSKD);
    document.getElementById('loadingState').style.display = 'none';
    document.getElementById('errorState').style.display   = 'none';
    const badge = document.getElementById('dataSourceBadge');
    badge.innerHTML = `<svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><path d="M14 2H6a2 2 0 00-2 2v16a2 2 0 002 2h12a2 2 0 002-2V8z"/><polyline points="14 2 14 8 20 8"/></svg> dataSKD.js`;
    document.getElementById('uploadInfo').textContent = `${allData.length.toLocaleString('id')} baris dimuat`;
    initDashboard();
  } else {
    document.getElementById('loadingState').style.display = 'none';
    document.getElementById('errorState').style.display   = '';
  }
});

// ==========================================
// TRANSFORM dataSKD.js FORMAT → internal
// ==========================================
function transformDataSKD(data) {
  return data.map(row => {
    const twk   = parseFloat(row.twk)   || 0;
    const tiu   = parseFloat(row.tiu)   || 0;
    const tkp   = parseFloat(row.tkp)   || 0;
    const total = parseFloat(row.total) || 0;
    const skd   = total > 0 ? total : (twk + tiu + tkp);

    const rawKet = String(row.keterangan || '').toUpperCase().trim();
    let ketCode = 'TH';
    if      (rawKet === 'P/L' || rawKet === 'P/L (LULUS SKB)')     ketCode = 'P/L';
    else if (rawKet === 'P'   || rawKet === 'LULUS' || rawKet === 'PASS' || rawKet === 'P (PASSING GRADE)') ketCode = 'P';
    else if (rawKet === 'TL'  || rawKet === 'TIDAK LULUS')          ketCode = 'TL';
    else if (rawKet === 'TH'  || rawKet === 'TIDAK HADIR')          ketCode = 'TH';
    else if (rawKet === 'TMS' || rawKet === 'GUGUR')                ketCode = 'TMS';
    else if (rawKet === 'DIS' || rawKet === 'DISKUALIFIKASI')       ketCode = 'DIS';

    const status = (ketCode === 'P/L' || ketCode === 'P') ? 'LULUS'
                 : ketCode === 'TL' ? 'TL'
                 : ketCode === 'TH' ? 'TH'
                 : ketCode;

    // Clean jenis: "1 - UMUM" → "UMUM"
    let jenisRaw = String(row.jenis || '').replace(/\s+/g, ' ').trim();
    const jenisM = jenisRaw.match(/^\d+\s*-\s*(.+)$/);
    const jenis  = jenisM ? jenisM[1].trim() : (jenisRaw || 'UMUM');

    return {
      nama:       String(row.nama       || `Peserta`).trim(),
      jabatan:    String(row.jabatan    || '—').trim(),
      pendidikan: String(row.pendidikan || '—').trim(),
      instansi:   String(row.instansi   || '—').trim(),
      lokasi:     String(row.lokasi     || '—').trim(),
      jenis,
      tahun:      String(row.tahunSKD   || row.tahun || '2024').replace(/\.0$/, ''),
      twk, tiu, tkp, skd,
      ket:    ketCode,
      status,
      nopes:  String(row.nopes  || '').trim(),
      syarat: String(row.syarat || '').trim(),
    };
  }).filter(r => r.nama !== '—' && r.nama !== '');
}

// ==========================================
// HELPERS
// ==========================================
function truncate(str, n) {
  return str.length > n ? str.slice(0, n) + '…' : str;
}
function escapeHTML(str) {
  return String(str)
    .replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;')
    .replace(/"/g,'&quot;').replace(/'/g,'&#39;');
}
function escapeAttr(str) {
  return String(str).replace(/'/g,"\\'").replace(/"/g,'&quot;');
}
function avg(arr) {
  const v = arr.filter(x => x > 0);
  return v.length ? v.reduce((a,b) => a+b, 0) / v.length : 0;
}

function parseJabatan(jabatan) {
  const raw = String(jabatan || '').trim();
  const sep = raw.indexOf(' - ');
  if (sep > 0) return { code: raw.substring(0, sep).trim(), name: raw.substring(sep + 3).trim() };
  return { code: '', name: raw };
}

// ==========================================
// XLSX PARSING (for file upload)
// ==========================================
function findCol(headers, candidates) {
  const lower = headers.map(h => String(h).trim().toLowerCase());
  for (const c of candidates) {
    const idx = lower.indexOf(c.toLowerCase());
    if (idx !== -1) return idx;
  }
  return -1;
}

function parseXLSXData(jsonRows) {
  if (!jsonRows || jsonRows.length < 1) return [];
  const rawHeaders = jsonRows[0];
  const headers    = rawHeaders.map(h => String(h || '').trim());

  const idx = {};
  for (const [key, cands] of Object.entries(COL)) {
    idx[key] = findCol(headers, cands);
  }

  const rows = [];
  for (let i = 1; i < jsonRows.length; i++) {
    const cells = jsonRows[i];
    if (!cells || cells.every(c => c === null || c === undefined || c === '')) continue;

    const get = key => {
      const ci = idx[key];
      if (ci === -1 || ci === undefined) return '';
      const val = cells[ci];
      return val === null || val === undefined ? '' : String(val).trim();
    };

    const twk   = parseFloat(get('TWK'))   || 0;
    const tiu   = parseFloat(get('TIU'))   || 0;
    const tkp   = parseFloat(get('TKP'))   || 0;
    const total = parseFloat(get('TOTAL')) || 0;
    const skd   = total > 0 ? total : (twk + tiu + tkp);

    const rawKet = get('KETERANGAN').toUpperCase().trim();
    let ketCode = 'TH';
    if      (rawKet === 'P/L' || rawKet === 'P/L (LULUS SKB)')     ketCode = 'P/L';
    else if (rawKet === 'P'   || rawKet === 'LULUS' || rawKet === 'PASS' || rawKet === 'P (PASSING GRADE)') ketCode = 'P';
    else if (rawKet === 'TL'  || rawKet === 'TIDAK LULUS')          ketCode = 'TL';
    else if (rawKet === 'TH'  || rawKet === 'TIDAK HADIR')          ketCode = 'TH';
    else if (rawKet === 'TMS' || rawKet === 'GUGUR')                ketCode = 'TMS';
    else if (rawKet === 'DIS' || rawKet === 'DISKUALIFIKASI')       ketCode = 'DIS';

    const status = (ketCode === 'P/L' || ketCode === 'P') ? 'LULUS'
                 : ketCode === 'TL' ? 'TL'
                 : ketCode === 'TH' ? 'TH'
                 : ketCode;

    let jenisRaw = get('JENIS').replace(/\s+/g, ' ').trim();
    const jenisM = jenisRaw.match(/^\d+\s*-\s*(.+)$/);
    const jenis  = jenisM ? jenisM[1].trim() : (jenisRaw || 'UMUM');

    const tahunRaw = get('TAHUN');
    const tahun    = tahunRaw ? String(tahunRaw).replace(/\.0$/, '') : '2024';

    rows.push({
      nama:       get('NAMA')       || `Peserta ${i}`,
      jabatan:    get('JABATAN')    || '—',
      pendidikan: get('PENDIDIKAN') || '—',
      instansi:   get('INSTANSI')   || '—',
      lokasi:     get('LOKASI')     || '—',
      jenis,
      tahun,
      twk, tiu, tkp, skd,
      ket: ketCode,
      status,
      nopes:  get('NOPES')  || '',
      syarat: get('SYARAT') || '',
    });
  }
  return rows;
}

// ==========================================
// XLSX FILE UPLOAD
// ==========================================
document.getElementById('xlsxInput').addEventListener('change', function(e) {
  const file = e.target.files[0];
  if (!file) return;

  const reader = new FileReader();
  reader.onload = function(ev) {
    try {
      const data     = new Uint8Array(ev.target.result);
      const wb       = XLSX.read(data, { type: 'array' });
      const ws       = wb.Sheets[wb.SheetNames[0]];
      const jsonRows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });

      allData = parseXLSXData(jsonRows);
      if (allData.length === 0) {
        alert('File Excel tidak valid atau kosong. Pastikan baris pertama adalah header kolom.');
        return;
      }

      const badge = document.getElementById('dataSourceBadge');
      badge.innerHTML = `<svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><path d="M14 2H6a2 2 0 00-2 2v16a2 2 0 002 2h12a2 2 0 002-2V8z"/><polyline points="14 2 14 8 20 8"/></svg> ${file.name}`;
      document.getElementById('uploadInfo').textContent =
        `${allData.length.toLocaleString('id')} baris dimuat`;
      document.getElementById('errorState').style.display = 'none';
      initDashboard();
    } catch (err) {
      alert('Gagal membaca file Excel: ' + err.message);
    }
  };
  reader.readAsArrayBuffer(file);
  this.value = '';
});

// ==========================================
// INIT DASHBOARD
// ==========================================
function initDashboard() {
  document.getElementById('loadingState').style.display = 'none';
  document.getElementById('errorState').style.display   = 'none';

  // Reset filters
  activeJenisFilters.clear();
  lulusJenisActiveFilters.clear();
  activePendLevel       = 'semua';
  activePendidikanItem  = null;

  const jabatanSet = [...new Set(allData.map(d => d.jabatan))].sort();
  const selJab     = document.getElementById('filterJabatanSelect');
  selJab.innerHTML = '<option value="">Semua Jabatan</option>';
  jabatanSet.forEach(j => {
    const { name } = parseJabatan(j);
    const o = document.createElement('option');
    o.value = j; o.textContent = truncate(name || j, 60);
    selJab.appendChild(o);
  });

  const jenisSet = [...new Set(allData.map(d => d.jenis))].filter(Boolean).sort();
  const selJenis = document.getElementById('filterJenisSelect');
  selJenis.innerHTML = '<option value="">Semua Jenis Formasi</option>';
  jenisSet.forEach(j => {
    const o = document.createElement('option');
    o.value = j; o.textContent = j;
    selJenis.appendChild(o);
  });

  const pendSet = [...new Set(allData.map(d => d.pendidikan))].filter(v => v && v !== '—').sort();
  const selPend = document.getElementById('filterPendidikanSelect');
  selPend.innerHTML = '<option value="">Semua Pendidikan</option>';
  pendSet.forEach(p => {
    const o = document.createElement('option');
    o.value = p; o.textContent = p;
    selPend.appendChild(o);
  });

  updateJenisChips(jenisSet);

  filteredPeserta = [...allData];
  switchView('overview');
}

// ==========================================
// UPDATE JENIS CHIPS
// ==========================================
function updateJenisChips(jenisSet) {
  const container = document.getElementById('jabatanFilterChips');
  if (!container) return;

  const hasDisabilitas = jenisSet.some(j => j.toUpperCase().includes('DISABILITAS'));
  const hasTerbaik     = jenisSet.some(j => j.toUpperCase().includes('TERBAIK'));
  const hasUmum        = jenisSet.some(j => j.toUpperCase().includes('UMUM'));
  const isSemua        = activeJenisFilters.size === 0;

  let html = `<button class="filter-chip${isSemua ? ' active' : ''}" data-jenis="semua" onclick="setJenisFilter('semua',this)">Semua</button>`;
  if (hasUmum) {
    const isActive = [...activeJenisFilters].some(f => f.toUpperCase().includes('UMUM'));
    html += `<button class="filter-chip${isActive ? ' active' : ''}" data-jenis="UMUM" onclick="setJenisFilter('UMUM',this)"><span class="chip-dot"></span>Umum</button>`;
  }
  if (hasDisabilitas) {
    const isActive = [...activeJenisFilters].some(f => f.toUpperCase().includes('DISABILITAS'));
    html += `<button class="filter-chip${isActive ? ' active' : ''}" data-jenis="DISABILITAS" onclick="setJenisFilter('DISABILITAS',this)"><span class="chip-dot"></span>Disabilitas</button>`;
  }
  if (hasTerbaik) {
    const isActive = [...activeJenisFilters].some(f => f.toUpperCase().includes('TERBAIK'));
    html += `<button class="filter-chip${isActive ? ' active' : ''}" data-jenis="LULUSAN TERBAIK" onclick="setJenisFilter('LULUSAN TERBAIK',this)"><span class="chip-dot"></span>Lulusan Terbaik</button>`;
  }
  container.innerHTML = html;
}

// ==========================================
// VIEW SWITCHING
// ==========================================
function switchView(view) {
  currentView = view;
  document.querySelectorAll('.nav-btn').forEach(b => {
    b.classList.toggle('active', b.dataset.view === view);
  });

  const allViews = ['viewOverview', 'viewPerJabatan', 'viewPeserta', 'viewDetailPeserta', 'viewPendidikan'];
  allViews.forEach(id => {
    const el = document.getElementById(id);
    if (el) el.style.display = 'none';
  });

  if (allData.length === 0) return;

  if (view === 'overview') {
    document.getElementById('viewOverview').style.display = '';
    renderOverview();
  } else if (view === 'per-jabatan') {
    document.getElementById('viewPerJabatan').style.display = '';
    renderPerJabatan();
  } else if (view === 'peserta') {
    document.getElementById('viewPeserta').style.display = '';
    renderPeserta();
  } else if (view === 'pendidikan') {
    document.getElementById('viewPendidikan').style.display = '';
    renderPendidikan();
  }
}

// ==========================================
// OVERVIEW & ANALISIS
// ==========================================
function renderOverview() {
  const total  = allData.length;
  const pl     = allData.filter(d => d.ket === 'P/L').length;
  const p      = allData.filter(d => d.ket === 'P').length;
  const tl     = allData.filter(d => d.ket === 'TL').length;
  const th     = allData.filter(d => d.ket === 'TH').length;
  const jabSet = [...new Set(allData.map(d => d.jabatan))];
  const avgSKD = total > 0 ? (allData.reduce((s,d) => s+d.skd, 0)/total).toFixed(1) : 0;
  const maxSKD = allData.length ? Math.max(...allData.map(d => d.skd)) : 0;

  document.getElementById('overviewDesc').textContent =
    `${total.toLocaleString('id')} peserta dari ${jabSet.length} jabatan formasi – SKD CPNS 2024 Kota Semarang`;

  const pct = v => total > 0 ? ((v/total)*100).toFixed(1) : '0';

  document.getElementById('statCards').innerHTML = `
    <div class="stat-card total-peserta" role="button" tabindex="0" onclick="showStatModal('total')" title="Lihat detail peserta">
      <div class="stat-card-top">
        <div class="stat-icon">
          <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
            <path d="M17 21v-2a4 4 0 00-4-4H5a4 4 0 00-4 4v2"/><circle cx="9" cy="7" r="4"/>
            <path d="M23 21v-2a4 4 0 00-3-3.87"/><path d="M16 3.13a4 4 0 010 7.75"/>
          </svg>
        </div>
        <span class="stat-trend">${jabSet.length} Jabatan</span>
      </div>
      <div class="stat-label">Total Peserta</div>
      <div class="stat-value">${total.toLocaleString('id')}</div>
      <div class="stat-sub">Seluruh formasi CPNS 2024</div>
      <div class="stat-bar"><div class="stat-bar-fill" style="width:100%"></div></div>
    </div>

    <div class="stat-card pl-lulus" role="button" tabindex="0" onclick="showStatModal('pl')" title="Lihat detail P/L">
      <div class="stat-card-top">
        <div class="stat-icon">
          <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
            <polyline points="20 6 9 17 4 12"/>
          </svg>
        </div>
        <span class="stat-trend">${pct(pl)}%</span>
      </div>
      <div class="stat-label">P/L – Lulus SKB</div>
      <div class="stat-value">${pl.toLocaleString('id')}</div>
      <div class="stat-sub">Lanjut ke Seleksi Kompetensi Bidang</div>
      <div class="stat-bar"><div class="stat-bar-fill" style="width:${pct(pl)}%"></div></div>
    </div>

    <div class="stat-card p-grade" role="button" tabindex="0" onclick="showStatModal('p')" title="Lihat detail Passing Grade">
      <div class="stat-card-top">
        <div class="stat-icon">
          <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
            <path d="M22 11.08V12a10 10 0 11-5.93-9.14"/>
            <polyline points="22 4 12 14.01 9 11.01"/>
          </svg>
        </div>
        <span class="stat-trend">${pct(p)}%</span>
      </div>
      <div class="stat-label">P – Passing Grade</div>
      <div class="stat-value">${p.toLocaleString('id')}</div>
      <div class="stat-sub">Memenuhi nilai ambang batas</div>
      <div class="stat-bar"><div class="stat-bar-fill" style="width:${pct(p)}%"></div></div>
    </div>

    <div class="stat-card tl-card" role="button" tabindex="0" onclick="showStatModal('tl')" title="Lihat detail Tidak Lulus">
      <div class="stat-card-top">
        <div class="stat-icon">
          <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
            <line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/>
          </svg>
        </div>
        <span class="stat-trend">${pct(tl)}%</span>
      </div>
      <div class="stat-label">TL – Tidak Lulus</div>
      <div class="stat-value">${tl.toLocaleString('id')}</div>
      <div class="stat-sub">Di bawah nilai ambang batas</div>
      <div class="stat-bar"><div class="stat-bar-fill" style="width:${pct(tl)}%"></div></div>
    </div>

    <div class="stat-card th-card" role="button" tabindex="0" onclick="showStatModal('th')" title="Lihat detail Tidak Hadir">
      <div class="stat-card-top">
        <div class="stat-icon">
          <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
            <circle cx="12" cy="12" r="10"/>
            <line x1="12" y1="8" x2="12" y2="12"/><line x1="12" y1="16" x2="12.01" y2="16"/>
          </svg>
        </div>
        <span class="stat-trend">${pct(th)}%</span>
      </div>
      <div class="stat-label">TH – Tidak Hadir</div>
      <div class="stat-value">${th.toLocaleString('id')}</div>
      <div class="stat-sub">Tidak mengikuti ujian SKD</div>
      <div class="stat-bar"><div class="stat-bar-fill" style="width:${pct(th)}%"></div></div>
    </div>

    <div class="stat-card skd-tertinggi" role="button" tabindex="0" onclick="showStatModal('skd')" title="Lihat top peserta SKD">
      <div class="stat-card-top">
        <div class="stat-icon">
          <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
            <polygon points="12 2 15.09 8.26 22 9.27 17 14.14 18.18 21.02 12 17.77 5.82 21.02 7 14.14 2 9.27 8.91 8.26 12 2"/>
          </svg>
        </div>
        <span class="stat-trend">Ø ${avgSKD}</span>
      </div>
      <div class="stat-label">SKD Tertinggi</div>
      <div class="stat-value">${maxSKD}</div>
      <div class="stat-sub">Nilai terbaik dari seluruh peserta</div>
      <div class="stat-bar"><div class="stat-bar-fill" style="width:${maxSKD > 0 ? Math.min(100, (maxSKD/600)*100).toFixed(1) : 0}%"></div></div>
    </div>
  `;

  renderCharts();
}

// ==========================================
// CHART HELPERS
// ==========================================
const CC = {
  PALETTE: ['#1D4ED8','#059669','#B91C1C','#64748B','#D97706','#6D28D9','#0F766E','#C2410C','#BE185D','#0E7490'],
};

function destroyChart(key) {
  if (charts[key]) { charts[key].destroy(); delete charts[key]; }
}

function fontCfg(size = 11) {
  return { family: "'Plus Jakarta Sans', 'Segoe UI', sans-serif", size };
}

// ==========================================
// CHARTS
// ==========================================
function renderCharts() {
  const total = allData.length;
  const pl    = allData.filter(d => d.ket === 'P/L').length;
  const p     = allData.filter(d => d.ket === 'P').length;
  const tl    = allData.filter(d => d.ket === 'TL').length;
  const th    = allData.filter(d => d.ket === 'TH').length;

  // 1. Doughnut: status kelulusan
  destroyChart('status');
  charts.status = new Chart(document.getElementById('chartStatus').getContext('2d'), {
    type: 'doughnut',
    data: {
      labels: ['P/L – Lulus SKB', 'P – Passing Grade', 'TL – Tidak Lulus', 'TH – Tidak Hadir'],
      datasets: [{
        data: [pl, p, tl, th],
        backgroundColor: ['#1D4ED8', '#059669', '#DC2626', '#64748B'],
        borderWidth: 3, borderColor: '#fff', hoverOffset: 10,
      }],
    },
    options: {
      responsive: true, maintainAspectRatio: false,
      cutout: '60%',
      plugins: {
        legend: { position: 'bottom', labels: { font: fontCfg(11), padding: 12, boxWidth: 12 } },
        tooltip: { callbacks: { label: c => ` ${c.label}: ${c.raw} peserta (${total>0?((c.raw/total)*100).toFixed(1):0}%)` } },
      },
    },
  });

  // 2. Pie: distribusi jenis formasi
  const jenisCount = {};
  allData.forEach(d => { jenisCount[d.jenis] = (jenisCount[d.jenis]||0)+1; });
  const jenisEntries = Object.entries(jenisCount).sort((a,b) => b[1]-a[1]);

  const PIE_COLORS = [
    '#1D4ED8', '#059669', '#B91C1C', '#D97706',
    '#7C3AED', '#0891B2', '#BE185D', '#065F46',
  ];

  destroyChart('jenis');
  charts.jenis = new Chart(document.getElementById('chartJenis').getContext('2d'), {
    type: 'pie',
    data: {
      labels: jenisEntries.map(e => e[0]),
      datasets: [{
        data: jenisEntries.map(e => e[1]),
        backgroundColor: jenisEntries.map((_,i) => PIE_COLORS[i % PIE_COLORS.length]),
        borderColor: '#fff',
        borderWidth: 3,
        hoverOffset: 16,
      }],
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: {
          position: 'bottom',
          labels: { font: fontCfg(11), padding: 12, boxWidth: 12, boxHeight: 12 },
        },
        tooltip: {
          callbacks: {
            label: c => {
              const total2 = jenisEntries.reduce((s,e) => s+e[1], 0);
              return ` ${c.label}: ${c.raw} (${((c.raw/total2)*100).toFixed(1)}%)`;
            },
          },
        },
      },
    },
  });

  // 3. Lulus per Jenis – build filter row first
  const allJenisNames = jenisEntries.map(e => e[0]);
  buildLulusJenisFilterRow(allJenisNames);
  renderLulusJenisChart(lulusJenisActiveFilters);

  // 4. Top 10 jabatan by avg SKD
  const jabMap = {};
  allData.forEach(d => {
    const { name } = parseJabatan(d.jabatan);
    const k = name || d.jabatan;
    (jabMap[k] = jabMap[k] || []).push(d.skd);
  });
  const jabRanked = Object.entries(jabMap)
    .filter(([,v]) => v.some(x => x > 0))
    .map(([name,vals]) => ({ name, avg: avg(vals) }))
    .sort((a,b) => b.avg-a.avg).slice(0,10);

  destroyChart('jabatan');
  charts.jabatan = new Chart(document.getElementById('chartJabatan').getContext('2d'), {
    type: 'bar',
    data: {
      labels: jabRanked.map(j => truncate(j.name, 38)),
      datasets: [{
        label: 'Rata-rata SKD',
        data: jabRanked.map(j => +j.avg.toFixed(1)),
        backgroundColor: jabRanked.map((_,i) => i===0?'#B91C1C':i<3?'#F87171':'rgba(185,28,28,.35)'),
        borderRadius: 6,
      }],
    },
    options: {
      indexAxis: 'y', responsive: true, maintainAspectRatio: false,
      plugins: { legend: { display: false } },
      scales: {
        x: { grid: { color: 'rgba(0,0,0,.04)' }, ticks: { font: fontCfg(11) } },
        y: { ticks: { font: fontCfg(10), color: '#4a5568' } },
      },
    },
  });

  // 5. Top 10 jabatan by count
  const jabCount = Object.entries(jabMap)
    .map(([name,vals]) => ({ name, count: vals.length }))
    .sort((a,b) => b.count-a.count).slice(0,10);

  destroyChart('jabatanCount');
  charts.jabatanCount = new Chart(document.getElementById('chartJabatanCount').getContext('2d'), {
    type: 'bar',
    data: {
      labels: jabCount.map(j => truncate(parseJabatan(j.name).name || j.name, 38)),
      datasets: [{
        label: 'Jumlah Peserta',
        data: jabCount.map(j => j.count),
        backgroundColor: jabCount.map((_,i) => i===0?'#6D28D9':i<3?'#A78BFA':'rgba(109,40,217,.35)'),
        borderRadius: 6,
      }],
    },
    options: {
      indexAxis: 'y', responsive: true, maintainAspectRatio: false,
      plugins: { legend: { display: false } },
      scales: {
        x: { grid: { color: 'rgba(0,0,0,.04)' }, ticks: { font: fontCfg(11) } },
        y: { ticks: { font: fontCfg(10), color: '#4a5568' } },
      },
    },
  });

  // 6. Histogram: distribusi nilai SKD
  const allSkd = allData.map(d => d.skd).filter(v => v > 0);
  if (allSkd.length > 0) {
    const minS  = Math.floor(Math.min(...allSkd)/10)*10;
    const maxS  = Math.ceil(Math.max(...allSkd)/10)*10;
    const binSz = Math.max(10, Math.ceil((maxS-minS)/14));
    const dbins = [];
    for (let b = minS; b < maxS; b += binSz)
      dbins.push({ label: `${b}–${b+binSz-1}`, min: b, max: b+binSz, count: 0 });
    allSkd.forEach(v => { const bn = dbins.find(b=>v>=b.min&&v<b.max); if(bn) bn.count++; });

    destroyChart('distribusi');
    charts.distribusi = new Chart(document.getElementById('chartDistribusi').getContext('2d'), {
      type: 'bar',
      data: {
        labels: dbins.map(b => b.label),
        datasets: [{
          label: 'Jumlah Peserta',
          data: dbins.map(b => b.count),
          backgroundColor: 'rgba(29,78,216,.55)',
          borderColor: '#1D4ED8', borderWidth: 1.5, borderRadius: 4,
        }],
      },
      options: {
        responsive: true, maintainAspectRatio: false,
        plugins: { legend: { display: false } },
        scales: {
          x: { ticks: { font: fontCfg(10), maxRotation: 45 } },
          y: { ticks: { font: fontCfg(11) }, grid: { color: 'rgba(0,0,0,.04)' } },
        },
      },
    });
  }

  // 7. Bar: rata-rata TWK · TIU · TKP
  const hasSubtes = allData.some(d => d.twk>0||d.tiu>0||d.tkp>0);
  destroyChart('subtes');
  if (hasSubtes) {
    charts.subtes = new Chart(document.getElementById('chartSubtes').getContext('2d'), {
      type: 'bar',
      data: {
        labels: ['TWK','TIU','TKP'],
        datasets: [{
          label: 'Rata-rata Nilai',
          data: [+avg(allData.map(d=>d.twk)).toFixed(1), +avg(allData.map(d=>d.tiu)).toFixed(1), +avg(allData.map(d=>d.tkp)).toFixed(1)],
          backgroundColor: ['#B91C1C','#D97706','#0F766E'],
          borderRadius: 8,
        }],
      },
      options: {
        responsive: true, maintainAspectRatio: false,
        plugins: {
          legend: { display: false },
          tooltip: { callbacks: { title: ctx => ({TWK:'Wawasan Kebangsaan',TIU:'Intelejensi Umum',TKP:'Karakteristik Pribadi'}[ctx[0].label]||ctx[0].label) } },
        },
        scales: {
          x: { ticks: { font: fontCfg(12) }, grid: { display: false } },
          y: { ticks: { font: fontCfg(11) }, grid: { color: 'rgba(0,0,0,.04)' } },
        },
      },
    });
  }

  // 8. Top 10 Pendidikan
  const pendCount = {};
  allData.forEach(d => { if (d.pendidikan && d.pendidikan !== '—') pendCount[d.pendidikan] = (pendCount[d.pendidikan]||0)+1; });
  const pendRanked = Object.entries(pendCount).sort((a,b) => b[1]-a[1]).slice(0,10);

  destroyChart('pendidikan');
  charts.pendidikan = new Chart(document.getElementById('chartPendidikan').getContext('2d'), {
    type: 'bar',
    data: {
      labels: pendRanked.map(p => p[0]),
      datasets: [{
        label: 'Jumlah Peserta',
        data: pendRanked.map(p => p[1]),
        backgroundColor: pendRanked.map((_,i) => CC.PALETTE[i%CC.PALETTE.length]),
        borderRadius: 6,
      }],
    },
    options: {
      indexAxis: 'y', responsive: true, maintainAspectRatio: false,
      plugins: { legend: { display: false } },
      scales: {
        x: { ticks: { font: fontCfg(11) }, grid: { color: 'rgba(0,0,0,.04)' } },
        y: { ticks: { font: fontCfg(10.5), color: '#4a5568' } },
      },
    },
  });
}

// ==========================================
// LULUS PER JENIS – Multi-select Filter
// ==========================================
function buildLulusJenisFilterRow(jenisNames) {
  const row = document.getElementById('lulusJenisFilterRow');
  if (!row) return;

  const dotColors = { umum:'#D97706', terbaik:'#16A34A', dis:'#1D4ED8', khusus:'#B91C1C' };
  function jenisClass(j) {
    const u = j.toUpperCase();
    if (u.includes('DISABILITAS')) return 'dis';
    if (u.includes('TERBAIK'))     return 'terbaik';
    if (u.includes('UMUM'))        return 'umum';
    return 'khusus';
  }

  let html = `<span class="cfc-label">Tampilkan:</span>
    <button class="cfc-chip active semua" data-lj-jenis="SEMUA">
      <svg class="cfc-check" viewBox="0 0 12 12" fill="none" stroke="currentColor" stroke-width="2.5"><polyline points="2 6 5 9 10 3"/></svg>
      Semua
    </button>`;

  jenisNames.forEach(j => {
    const cls = jenisClass(j);
    const dot = dotColors[cls] || '#64748B';
    html += `<button class="cfc-chip inactive ${cls}" data-lj-jenis="${escapeAttr(j)}">
      <span class="cfc-dot" style="background:${dot}"></span>
      <svg class="cfc-check" viewBox="0 0 12 12" fill="none" stroke="currentColor" stroke-width="2.5"><polyline points="2 6 5 9 10 3"/></svg>
      ${escapeHTML(j)}
    </button>`;
  });
  row.innerHTML = html;
}

function toggleLulusJenisFilter(jenis) {
  const row = document.getElementById('lulusJenisFilterRow');
  if (!row) return;

  if (jenis === 'SEMUA') {
    lulusJenisActiveFilters.clear();
  } else {
    lulusJenisActiveFilters.has(jenis) ? lulusJenisActiveFilters.delete(jenis) : lulusJenisActiveFilters.add(jenis);
  }

  row.querySelectorAll('.cfc-chip').forEach(chip => {
    const cj = chip.dataset.ljJenis;
    chip.classList.remove('active', 'inactive');
    chip.classList.add(cj === 'SEMUA' ? (lulusJenisActiveFilters.size===0?'active':'inactive') : (lulusJenisActiveFilters.has(cj)?'active':'inactive'));
  });
  renderLulusJenisChart(lulusJenisActiveFilters);
}

document.addEventListener('click', function(e) {
  const chip = e.target.closest('.cfc-chip[data-lj-jenis]');
  if (chip) toggleLulusJenisFilter(chip.dataset.ljJenis);
});

function renderLulusJenisChart(activeFilters) {
  const jenisCount = {};
  allData.forEach(d => { jenisCount[d.jenis] = (jenisCount[d.jenis]||0)+1; });
  const allJenisNames = Object.entries(jenisCount).sort((a,b) => b[1]-a[1]).map(e => e[0]);
  const filteredJenis = activeFilters.size === 0 ? allJenisNames : allJenisNames.filter(j => activeFilters.has(j));

  destroyChart('lulusPerJenis');
  const ctx = document.getElementById('chartLulusPerJenis');
  if (!ctx) return;

  charts.lulusPerJenis = new Chart(ctx.getContext('2d'), {
    type: 'bar',
    data: {
      labels: filteredJenis,
      datasets: [
        { label:'P/L',         data: filteredJenis.map(j=>allData.filter(d=>d.jenis===j&&d.ket==='P/L').length), backgroundColor:'#1D4ED8', borderRadius:5 },
        { label:'P (PG)',      data: filteredJenis.map(j=>allData.filter(d=>d.jenis===j&&d.ket==='P').length),   backgroundColor:'#059669', borderRadius:5 },
        { label:'Tidak Lulus', data: filteredJenis.map(j=>allData.filter(d=>d.jenis===j&&d.ket==='TL').length),  backgroundColor:'#B91C1C', borderRadius:5 },
        { label:'Tidak Hadir', data: filteredJenis.map(j=>allData.filter(d=>d.jenis===j&&d.ket==='TH').length),  backgroundColor:'#64748B', borderRadius:5 },
      ],
    },
    options: {
      responsive: true, maintainAspectRatio: false,
      plugins: { legend: { position:'top', labels: { font: fontCfg(12), padding: 16 } } },
      scales: {
        x: { ticks: { font: fontCfg(10.5), maxRotation: 20 }, grid: { display: false } },
        y: { ticks: { font: fontCfg(11) }, grid: { color: 'rgba(0,0,0,.04)' } },
      },
    },
  });
}

// ==========================================
// PER JABATAN – Grouped by jabatan+lokasi+jenis
// Each unique jabatan/lokasi/jenis gets its own card
// ==========================================
function setJenisFilter(jenis, btnEl) {
  if (jenis === 'semua') {
    activeJenisFilters.clear();
  } else {
    activeJenisFilters.has(jenis) ? activeJenisFilters.delete(jenis) : activeJenisFilters.add(jenis);
  }

  document.querySelectorAll('#jabatanFilterChips .filter-chip').forEach(b => {
    const dj = b.dataset.jenis;
    b.classList.toggle('active', dj === 'semua' ? activeJenisFilters.size === 0 : activeJenisFilters.has(dj));
  });

  renderPerJabatan(document.getElementById('searchJabatan')?.value || '');
}

/**
 * Group by jabatan + lokasi + jenis — so UMUM and DISABILITAS
 * entries for the same jabatan appear as SEPARATE cards.
 */
function groupByJabatanLokasiJenis(data) {
  return data.reduce((map, d) => {
    const key = d.jabatan + '||' + d.lokasi + '||' + d.jenis;
    (map[key] = map[key] || []).push(d);
    return map;
  }, {});
}

function renderPerJabatan(filter = '') {
  const allJabMap = groupByJabatanLokasiJenis(allData);
  let keys = Object.keys(allJabMap);

  // Filter by active jenis chips
  if (activeJenisFilters.size > 0) {
    keys = keys.filter(key => {
      const jenisPart = key.split('||')[2] || '';
      return [...activeJenisFilters].some(f => jenisPart.toUpperCase().includes(f.toUpperCase()));
    });
  }

  // Filter by search text
  if (filter) {
    keys = keys.filter(k => k.toLowerCase().includes(filter.toLowerCase()));
  }

  keys = keys.sort();

  let filterLabel = activeJenisFilters.size > 0 ? ` · Filter: ${[...activeJenisFilters].join(', ')}` : '';
  document.getElementById('jabatanDesc').textContent =
    `${keys.length} jabatan formasi ditampilkan${filterLabel}`;

  const grid = document.getElementById('jabatanGrid');
  if (!keys.length) {
    grid.innerHTML = `<div class="no-results" style="grid-column:1/-1">Tidak ada jabatan yang cocok dengan filter ini.</div>`;
    return;
  }

  grid.innerHTML = keys.map(key => {
    const [jabatan, lokasi, jenis] = key.split('||');
    const peserta   = allJabMap[key];
    const plC  = peserta.filter(p => p.ket==='P/L').length;
    const pC   = peserta.filter(p => p.ket==='P').length;
    const tlC  = peserta.filter(p => p.ket==='TL').length;
    const thC  = peserta.filter(p => p.ket==='TH').length;
    const avgS = (peserta.reduce((s,p) => s+p.skd, 0)/peserta.length).toFixed(0);
    const sample = peserta[0];

    const { code, name: displayName } = parseJabatan(jabatan);
    const displayLokasi = lokasi && lokasi !== '—' ? lokasi : (sample.instansi||'Pemerintah Kota Semarang');

    // Single jenis badge per card (since we split by jenis)
    const ju  = (jenis||'').toUpperCase();
    const cls = ju.includes('DISABILITAS') ? 'disabilitas' : ju.includes('TERBAIK') ? 'terbaik' : ju.includes('UMUM') ? 'umum' : 'khusus';
    const badge = `<span class="badge-formasi ${cls}">${escapeHTML(jenis||'UMUM')}</span>`;

    const pendSet    = [...new Set(peserta.map(p => p.pendidikan).filter(p => p&&p!=='—'))];
    const syaratText = sample.syarat || (pendSet.length ? pendSet.join(', ') : '');

    return `
      <div class="jabatan-card" data-key="${escapeAttr(key)}" tabindex="0" role="button" aria-label="${escapeAttr(displayName || jabatan)}">
        <div class="jabatan-card-header">
          <div class="jabatan-card-hint">
            <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><line x1="12" y1="5" x2="12" y2="19"/><line x1="5" y1="12" x2="19" y2="12"/></svg>
          </div>
          ${code ? `<div class="jabatan-code">${escapeHTML(code)}</div>` : ''}
          <div class="jabatan-name">${escapeHTML(displayName || jabatan)}</div>
          <div class="jabatan-meta">
            <svg width="11" height="11" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><path d="M21 10c0 7-9 13-9 13S3 17 3 10a9 9 0 0118 0z"/><circle cx="12" cy="10" r="3"/></svg>
            ${escapeHTML(truncate(displayLokasi, 80))}
          </div>
          ${syaratText ? `<div class="jabatan-syarat">
            <svg width="11" height="11" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><path d="M22 10v6M2 10l10-5 10 5-10 5z"/><path d="M6 12v5c3 3 9 3 12 0v-5"/></svg>
            ${escapeHTML(truncate(syaratText, 80))}
          </div>` : ''}
          <div style="margin-top:8px">${badge}</div>
        </div>
        <div class="jabatan-card-stats">
          <div class="jabatan-stat"><div class="jabatan-stat-val pl">${plC}</div><div class="jabatan-stat-lbl">P/L</div></div>
          <div class="jabatan-divider"></div>
          <div class="jabatan-stat"><div class="jabatan-stat-val lulus">${pC}</div><div class="jabatan-stat-lbl">P</div></div>
          <div class="jabatan-divider"></div>
          <div class="jabatan-stat"><div class="jabatan-stat-val tl">${tlC}</div><div class="jabatan-stat-lbl">TL</div></div>
          <div class="jabatan-divider"></div>
          <div class="jabatan-stat"><div class="jabatan-stat-val th">${thC}</div><div class="jabatan-stat-lbl">TH</div></div>
          <div class="jabatan-divider"></div>
          <div class="jabatan-stat"><div class="jabatan-stat-val rata">${avgS}</div><div class="jabatan-stat-lbl">Rata SKD</div></div>
        </div>
        <div class="jabatan-card-footer">
          <span style="display:flex;align-items:center;gap:5px">
            <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M17 21v-2a4 4 0 00-4-4H5a4 4 0 00-4 4v2"/><circle cx="9" cy="7" r="4"/></svg>
            ${peserta.length} peserta
          </span>
          <span class="jabatan-card-footer-right">
            Detail
            <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><polyline points="9 18 15 12 9 6"/></svg>
          </span>
        </div>
      </div>`;
  }).join('');
}

function filterJabatan() {
  renderPerJabatan(document.getElementById('searchJabatan').value);
}

// Delegated click for jabatan cards
document.addEventListener('click', function(e) {
  const card = e.target.closest('.jabatan-card[data-key]');
  if (card) showJabatanModal(card.dataset.key);
});

// ==========================================
// PENDIDIKAN VIEW
// ==========================================

/** Get education level abbreviation for icon */
function getPendLevel(pend) {
  const u = pend.toUpperCase();
  if (u.startsWith('S-3') || u.startsWith('S3'))  return 'S3';
  if (u.startsWith('S-2') || u.startsWith('S2'))  return 'S2';
  if (u.startsWith('S-1') || u.startsWith('S1') || u.startsWith('SARJANA')) return 'S1';
  if (u.startsWith('D-4') || u.startsWith('D4'))  return 'D4';
  if (u.startsWith('D-3') || u.startsWith('D3') || u.startsWith('DIPLOMA III') || u.startsWith('DIII')) return 'D3';
  if (u.startsWith('D-2') || u.startsWith('D2'))  return 'D2';
  if (u.startsWith('D-1') || u.startsWith('D1'))  return 'D1';
  if (u.includes('PROFESI') || u.startsWith('PROF'))  return 'PROFESI';
  if (u.includes('SMA') || u.includes('SMK') || u.includes('SLTA')) return 'SMA';
  return '—';
}

function getPendLevelGroup(pend) {
  const u = pend.toUpperCase();
  if (u.startsWith('S-3') || u.startsWith('S3'))  return 'S3';
  if (u.startsWith('S-2') || u.startsWith('S2'))  return 'S2';
  if (u.startsWith('S-1') || u.startsWith('S1') || u.startsWith('SARJANA')) return 'S1';
  if (u.startsWith('D-4') || u.startsWith('D4'))  return 'D4';
  if (u.startsWith('D-3') || u.startsWith('D3') || u.startsWith('DIPLOMA III') || u.startsWith('DIII')) return 'D3';
  if (u.startsWith('D-2') || u.startsWith('D2'))  return 'D2';
  if (u.startsWith('D-1') || u.startsWith('D1'))  return 'D1';
  if (u.includes('PROFESI') || u.startsWith('PROF')) return 'PROFESI';
  if (u.includes('SMA') || u.includes('SMK') || u.includes('SLTA')) return 'SMA';
  return 'LAIN';
}

function renderPendidikan(filterText = '') {
  // Build summary per unique pendidikan — filter out NaN/invalid entries
  const pendMap = {};
  allData.forEach(d => {
    const pend = d.pendidikan;
    if (!pend || pend === '—' || pend === 'NaN' || pend.toLowerCase() === 'nan' || pend.trim() === '') return;
    if (!pendMap[pend]) {
      pendMap[pend] = {
        nama: pend,
        peserta: [],
        jabatanSet: new Set(),
        level: getPendLevelGroup(pend),
      };
    }
    pendMap[pend].peserta.push(d);
    pendMap[pend].jabatanSet.add(d.jabatan);
  });

  // Build level chips from available data
  const levelsAvail = [...new Set(Object.values(pendMap).map(p => p.level))].sort();
  const levelOrder  = ['S3','S2','S1','D4','D3','D2','D1','SMA','LAIN'];
  const sortedLevels = levelOrder.filter(l => levelsAvail.includes(l));
  buildPendLevelChips(sortedLevels);

  // Apply filter
  let pendList = Object.values(pendMap).sort((a,b) => b.peserta.length - a.peserta.length);

  if (activePendLevel && activePendLevel !== 'semua') {
    pendList = pendList.filter(p => p.level === activePendLevel);
  }
  if (filterText) {
    const q = filterText.toLowerCase();
    pendList = pendList.filter(p =>
      p.nama.toLowerCase().includes(q) ||
      [...p.jabatanSet].some(j => j.toLowerCase().includes(q))
    );
  }

  document.getElementById('pendidikanDesc').textContent =
    `${pendList.length} jenis pendidikan dari ${Object.keys(pendMap).length} total – ${allData.length.toLocaleString('id')} peserta`;

  // Render sidebar
  const sidebar = document.getElementById('pendidikanSidebar');
  if (!pendList.length) {
    sidebar.innerHTML = `<div class="no-results">Tidak ada data pendidikan yang cocok.</div>`;
    return;
  }

  sidebar.innerHTML = pendList.map(p => {
    const isActive = activePendidikanItem === p.nama;
    const levelAbbr = getPendLevel(p.nama);
    const plCount = p.peserta.filter(d => d.ket==='P/L').length;
    const pCount  = p.peserta.filter(d => d.ket==='P').length;
    return `
      <div class="pend-sidebar-item${isActive ? ' active' : ''}" data-pend="${escapeAttr(p.nama)}" onclick="selectPendidikan('${escapeAttr(p.nama)}')">
        <div class="pend-item-icon">${escapeHTML(levelAbbr)}</div>
        <div class="pend-item-info">
          <div class="pend-item-name">${escapeHTML(p.nama)}</div>
          <div class="pend-item-sub">${p.jabatanSet.size} jabatan · ${plCount+pCount} lulus</div>
        </div>
        <div class="pend-item-count">${p.peserta.length}</div>
      </div>`;
  }).join('');

  // If an item was active and still in list, show detail
  if (activePendidikanItem && pendList.some(p => p.nama === activePendidikanItem)) {
    renderPendidikanDetail(pendMap[activePendidikanItem]);
  }
}

function buildPendLevelChips(levels) {
  const chips = document.getElementById('pendLevelChips');
  if (!chips) return;
  const levelLabels = { S3:'S-3', S2:'S-2', S1:'S-1', D4:'D-IV', D3:'D-III', PROFESI:'Profesi', D2:'D-2', D1:'D-1', SMA:'SMA/SMK', LAIN:'Lainnya' };
  let html = `<button class="pend-level-chip${activePendLevel==='semua'?' active':''}" onclick="setPendLevel('semua')">Semua</button>`;
  const levelOrder = ['S3','S2','S1','D4','D3','PROFESI','D2','D1','SMA','LAIN'];
  const sorted = levelOrder.filter(l => levels.includes(l));
  sorted.forEach(l => {
    html += `<button class="pend-level-chip${activePendLevel===l?' active':''}" onclick="setPendLevel('${l}')">${levelLabels[l]||l}</button>`;
  });
  chips.innerHTML = html;
}

function setPendLevel(level) {
  activePendLevel = level;
  const q = document.getElementById('searchPendidikanView')?.value || '';
  renderPendidikan(q);
}

function filterPendidikanView() {
  const q = document.getElementById('searchPendidikanView').value;
  renderPendidikan(q);
}

function selectPendidikan(nama) {
  activePendidikanItem = nama;
  // Update sidebar active state
  document.querySelectorAll('.pend-sidebar-item').forEach(el => {
    el.classList.toggle('active', el.dataset.pend === nama);
  });

  // Build detail for selected pendidikan
  const pendMap = {};
  allData.forEach(d => {
    if (!d.pendidikan || d.pendidikan === '—') return;
    if (!pendMap[d.pendidikan]) pendMap[d.pendidikan] = { nama: d.pendidikan, peserta: [], jabatanSet: new Set() };
    pendMap[d.pendidikan].peserta.push(d);
    pendMap[d.pendidikan].jabatanSet.add(d.jabatan);
  });

  if (pendMap[nama]) renderPendidikanDetail(pendMap[nama]);
}

function renderPendidikanDetail(pData) {
  const panel = document.getElementById('pendidikanDetailPanel');
  const peserta = pData.peserta;
  const plC  = peserta.filter(d => d.ket==='P/L').length;
  const pC   = peserta.filter(d => d.ket==='P').length;
  const tlC  = peserta.filter(d => d.ket==='TL').length;
  const thC  = peserta.filter(d => d.ket==='TH').length;
  const avgS = peserta.some(d=>d.skd>0) ? (peserta.reduce((s,d)=>s+d.skd,0)/peserta.filter(d=>d.skd>0).length).toFixed(1) : '—';

  // Jabatan list grouped
  const jabMap2 = {};
  peserta.forEach(d => {
    const key = d.jabatan + '||' + d.jenis;
    (jabMap2[key] = jabMap2[key] || []).push(d);
  });
  const jabEntries = Object.entries(jabMap2).sort((a,b) => b[1].length - a[1].length);

  const jabatanHTML = jabEntries.map(([key, ps]) => {
    const [jabatan, jenis] = key.split('||');
    const { code, name: dname } = parseJabatan(jabatan);
    const ju  = (jenis||'').toUpperCase();
    const cls = ju.includes('DISABILITAS') ? 'disabilitas' : ju.includes('TERBAIK') ? 'terbaik' : ju.includes('UMUM') ? 'umum' : 'khusus';
    const lokSet = [...new Set(ps.map(p => p.lokasi).filter(Boolean))];
    return `
      <div class="pend-jabatan-item">
        <div class="pend-jabatan-item-info">
          ${code ? `<div class="pend-jabatan-item-code">${escapeHTML(code)}</div>` : ''}
          <div class="pend-jabatan-item-name">${escapeHTML(dname || jabatan)}</div>
          <div class="pend-jabatan-item-meta">${escapeHTML(truncate(lokSet[0]||'',80))}</div>
        </div>
        <div class="pend-jabatan-item-right">
          <div class="pend-jabatan-peserta-count">${ps.length}</div>
          <div class="pend-jabatan-badges">
            <span class="badge-formasi ${cls}">${escapeHTML(jenis||'UMUM')}</span>
          </div>
        </div>
      </div>`;
  }).join('');

  // Peserta list (top 100 by SKD)
  const sortedPeserta = [...peserta].sort((a,b) => b.skd-a.skd).slice(0, 100);
  const pesertaHTML = sortedPeserta.map((d,i) => {
    const ridx = filteredPeserta.indexOf(d);
    const ki   = getKetInfo(d.ket);
    const { name: jabName } = parseJabatan(d.jabatan);
    return `
      <div class="pend-peserta-row">
        <div class="pend-peserta-no">${i+1}</div>
        <div class="pend-peserta-nama" ${ridx>=0?`onclick="switchView('peserta');setTimeout(()=>showDetailPage(${ridx}),50)"`:''}>
          ${escapeHTML(d.nama)}
        </div>
        <div class="pend-peserta-jabatan">${escapeHTML(truncate(jabName||d.jabatan,50))}</div>
        <div class="pend-peserta-skd">${d.skd > 0 ? d.skd : '—'}</div>
        <span class="badge-status ${ki.cls}" style="font-size:9px;padding:2px 8px">${ki.label}</span>
      </div>`;
  }).join('');

  panel.innerHTML = `
    <div class="pend-detail-header">
      <div class="pend-detail-title">${escapeHTML(pData.nama)}</div>
      <div class="pend-detail-stats">
        <div class="pend-detail-stat">
          <div class="pend-detail-stat-num" style="color:var(--text-primary)">${peserta.length}</div>
          <div class="pend-detail-stat-lbl">Total<br>Peserta</div>
        </div>
        <div class="pend-detail-stat">
          <div class="pend-detail-stat-num" style="color:var(--blue)">${plC}</div>
          <div class="pend-detail-stat-lbl">Lulus<br>SKB</div>
        </div>
        <div class="pend-detail-stat">
          <div class="pend-detail-stat-num" style="color:var(--green)">${pC}</div>
          <div class="pend-detail-stat-lbl">Passing<br>Grade</div>
        </div>
        <div class="pend-detail-stat">
          <div class="pend-detail-stat-num" style="color:var(--red)">${tlC}</div>
          <div class="pend-detail-stat-lbl">Tidak<br>Lulus</div>
        </div>
        <div class="pend-detail-stat">
          <div class="pend-detail-stat-num" style="color:#64748B">${thC}</div>
          <div class="pend-detail-stat-lbl">Tidak<br>Hadir</div>
        </div>
        <div class="pend-detail-stat">
          <div class="pend-detail-stat-num" style="color:var(--yellow-light)">${avgS}</div>
          <div class="pend-detail-stat-lbl">Rata-rata<br>SKD</div>
        </div>
        <div class="pend-detail-stat">
          <div class="pend-detail-stat-num" style="color:var(--text-primary)">${jabEntries.length}</div>
          <div class="pend-detail-stat-lbl">Jabatan<br>Dilamar</div>
        </div>
      </div>
    </div>
    <div class="pend-detail-body">
      <div class="pend-jabatan-section-title">Jabatan yang Dilamar (${jabEntries.length})</div>
      <div class="pend-jabatan-list">${jabatanHTML}</div>
      <div class="pend-jabatan-section-title">Daftar Peserta${peserta.length>100?' (Top 100 nilai tertinggi)':' ('+peserta.length+')'}</div>
      <div class="pend-peserta-list">${pesertaHTML}</div>
    </div>
  `;
}

// ==========================================
// DAFTAR PESERTA
// ==========================================
function renderPeserta() {
  currentPage     = 1;
  filteredPeserta = [...allData];
  document.getElementById('pesertaDesc').textContent =
    `${allData.length.toLocaleString('id')} peserta terdaftar`;
  renderPesertaTable();
}

function filterPeserta() {
  currentPage = 1;
  const q     = document.getElementById('searchPeserta').value.toLowerCase();
  const ket   = document.getElementById('filterKeterangan').value;
  const jab   = document.getElementById('filterJabatanSelect').value;
  const jenis = document.getElementById('filterJenisSelect').value;
  const pend  = document.getElementById('filterPendidikanSelect').value;

  filteredPeserta = allData.filter(d => {
    const mQ  = !q    || d.nama.toLowerCase().includes(q) || d.jabatan.toLowerCase().includes(q) || d.pendidikan.toLowerCase().includes(q);
    const mK  = !ket  || d.ket === ket;
    const mJ  = !jab  || d.jabatan === jab;
    const mJn = !jenis || d.jenis === jenis;
    const mP  = !pend || d.pendidikan === pend;
    return mQ && mK && mJ && mJn && mP;
  });

  document.getElementById('pesertaDesc').textContent =
    `Menampilkan ${filteredPeserta.length.toLocaleString('id')} dari ${allData.length.toLocaleString('id')} peserta`;
  renderPesertaTable();
}

function getKetInfo(ket) {
  const map = {
    'P/L': { label:'P/L', cls:'pl'   },
    'P':   { label:'P',   cls:'lulus' },
    'TL':  { label:'TL',  cls:'tl'   },
    'TH':  { label:'TH',  cls:'th'   },
    'TMS': { label:'TMS', cls:'tms'  },
    'DIS': { label:'DIS', cls:'dis'  },
  };
  return map[ket] || { label: ket || 'TH', cls: 'th' };
}

function renderPesertaTable() {
  const tbody = document.getElementById('pesertaTbody');
  const total = filteredPeserta.length;
  const pages = Math.max(1, Math.ceil(total/PAGE_SIZE));
  if (currentPage > pages) currentPage = pages;

  if (total === 0) {
    tbody.innerHTML = `<tr><td colspan="11" class="no-results">Tidak ada data yang cocok.</td></tr>`;
    document.getElementById('tableInfo').textContent = '';
    document.getElementById('pagination').innerHTML  = '';
    return;
  }

  const start = (currentPage-1)*PAGE_SIZE;
  const slice = filteredPeserta.slice(start, start+PAGE_SIZE);

  tbody.innerHTML = slice.map((d,i) => {
    const absIdx = start+i;
    const ki     = getKetInfo(d.ket);
    const jenisClass = d.jenis.toUpperCase().includes('DISABILITAS') ? 'disabilitas'
                     : d.jenis.toUpperCase().includes('TERBAIK') ? 'terbaik'
                     : d.jenis.toUpperCase().includes('UMUM') ? 'umum' : 'khusus';
    const { name: jabName } = parseJabatan(d.jabatan);
    return `<tr>
      <td>${start+i+1}</td>
      <td><span class="link-nama" onclick="showDetailPage(${absIdx})">${escapeHTML(d.nama)}</span></td>
      <td><span class="jabatan-chip" title="${escapeAttr(jabName||d.jabatan)}">${escapeHTML(truncate(jabName||d.jabatan, 36))}</span></td>
      <td class="cell-pendidikan">${escapeHTML(d.pendidikan)}</td>
      <td><span class="badge-formasi ${jenisClass}">${escapeHTML(d.jenis)}</span></td>
      <td class="cell-tahun">${escapeHTML(d.tahun)}</td>
      <td>${d.twk > 0 ? d.twk : '—'}</td>
      <td>${d.tiu > 0 ? d.tiu : '—'}</td>
      <td>${d.tkp > 0 ? d.tkp : '—'}</td>
      <td class="nilai-skd">${d.skd > 0 ? d.skd : '—'}</td>
      <td><span class="badge-status ${ki.cls}">${ki.label}</span></td>
    </tr>`;
  }).join('');

  document.getElementById('tableInfo').textContent =
    `Menampilkan ${start+1}–${Math.min(start+PAGE_SIZE, total)} dari ${total.toLocaleString('id')} peserta`;

  renderPagination(pages);
}

function renderPagination(pages) {
  const pg = document.getElementById('pagination');
  if (pages <= 1) { pg.innerHTML = ''; return; }

  const MAX_BTNS = 7;
  let btns = [];

  btns.push(`<button class="pg-btn${currentPage===1?' disabled':''}" onclick="goPage(${currentPage-1})">‹</button>`);

  if (pages <= MAX_BTNS) {
    for (let i=1; i<=pages; i++)
      btns.push(`<button class="pg-btn${i===currentPage?' active':''}" onclick="goPage(${i})">${i}</button>`);
  } else {
    btns.push(`<button class="pg-btn${currentPage===1?' active':''}" onclick="goPage(1)">1</button>`);
    if (currentPage > 3) btns.push(`<span class="pg-ellipsis">…</span>`);
    for (let i=Math.max(2,currentPage-1); i<=Math.min(pages-1,currentPage+1); i++)
      btns.push(`<button class="pg-btn${i===currentPage?' active':''}" onclick="goPage(${i})">${i}</button>`);
    if (currentPage < pages-2) btns.push(`<span class="pg-ellipsis">…</span>`);
    btns.push(`<button class="pg-btn${currentPage===pages?' active':''}" onclick="goPage(${pages})">${pages}</button>`);
  }

  btns.push(`<button class="pg-btn${currentPage===pages?' disabled':''}" onclick="goPage(${currentPage+1})">›</button>`);
  pg.innerHTML = btns.join('');
}

function goPage(n) {
  const pages = Math.max(1, Math.ceil(filteredPeserta.length/PAGE_SIZE));
  if (n < 1 || n > pages) return;
  currentPage = n;
  renderPesertaTable();
  window.scrollTo({ top: 300, behavior: 'smooth' });
}

// ==========================================
// DETAIL PAGE – Peserta
// ==========================================
function showDetailPage(idx) {
  prevView = currentView;
  const d  = filteredPeserta[idx];
  if (!d) return;

  const ki       = getKetInfo(d.ket);
  const hasSub   = d.twk > 0 || d.tiu > 0 || d.tkp > 0;
  const ketLabel = { 'P/L':'P/L – Lulus SKB','P':'P – Passing Grade','TL':'TL – Tidak Lulus','TH':'TH – Tidak Hadir','TMS':'TMS – Gugur Syarat','DIS':'DIS – Diskualifikasi' };
  const ketText  = ketLabel[d.ket] || d.ket;
  const hasilCls = { 'P/L':'pl','P':'lulus','TL':'tl','TH':'th','TMS':'tms','DIS':'dis' };
  const hasilBtnCls = 'detail-hasil-' + (hasilCls[d.ket]||'th');

  ['viewOverview','viewPerJabatan','viewPeserta','viewPendidikan'].forEach(id => {
    const el = document.getElementById(id);
    if (el) el.style.display = 'none';
  });

  const { code, name: jabName } = parseJabatan(d.jabatan);
  const displayLokasi = d.lokasi && d.lokasi !== '—' ? d.lokasi : (d.instansi||'Pemerintah Kota Semarang');

  document.getElementById('detailPageContent').innerHTML = `
    <div class="detail-hero">
      <div class="detail-hero-left">
        ${d.nopes ? `<div class="detail-nopes">No. Peserta: ${escapeHTML(d.nopes)}</div>` : ''}
        <div class="detail-nama">${escapeHTML(d.nama)}</div>
        <div class="detail-tahun-badge">SKD ${escapeHTML(d.tahun||'2024')}</div>
      </div>
      <div class="detail-hero-right">
        <div class="detail-total-label">Nilai SKD</div>
        <div class="detail-total-val">${d.skd > 0 ? d.skd : '—'}</div>
      </div>
    </div>

    <div class="detail-body">
      <div class="detail-section">
        <div class="detail-section-title">Nilai Sub-Tes</div>
        ${hasSub ? `
          <div class="detail-scores">
            <div class="detail-score-row">
              <span class="detail-score-label">TWK – Wawasan Kebangsaan</span>
              <span class="detail-score-val">${d.twk > 0 ? d.twk : '—'}</span>
            </div>
            <div class="detail-score-row">
              <span class="detail-score-label">TIU – Intelejensi Umum</span>
              <span class="detail-score-val">${d.tiu > 0 ? d.tiu : '—'}</span>
            </div>
            <div class="detail-score-row">
              <span class="detail-score-label">TKP – Karakteristik Pribadi</span>
              <span class="detail-score-val">${d.tkp > 0 ? d.tkp : '—'}</span>
            </div>
          </div>
          <div class="detail-score-chart-wrap">
            <canvas id="detailScoreChart"></canvas>
          </div>` : `<p style="color:var(--text-muted);font-style:italic;font-size:13px">Data nilai sub-tes tidak tersedia.</p>`}
      </div>

      <div class="detail-section">
        <div class="detail-section-title">Informasi Jabatan</div>
        <div class="detail-info-grid">
          <div class="detail-info-item">
            <div class="detail-info-label">Jabatan Formasi</div>
            <div class="detail-info-val">${escapeHTML(jabName || d.jabatan)}</div>
          </div>
          ${code ? `<div class="detail-info-item">
            <div class="detail-info-label">Kode Jabatan</div>
            <div class="detail-info-val" style="font-family:var(--font-mono)">${escapeHTML(code)}</div>
          </div>` : ''}
          <div class="detail-info-item">
            <div class="detail-info-label">Unit Kerja / Lokasi</div>
            <div class="detail-info-val">${escapeHTML(displayLokasi)}</div>
          </div>
          <div class="detail-info-item">
            <div class="detail-info-label">Pendidikan</div>
            <div class="detail-info-val">${escapeHTML(d.pendidikan)}</div>
          </div>
          <div class="detail-info-item">
            <div class="detail-info-label">Jenis Formasi</div>
            <div class="detail-info-val">${escapeHTML(d.jenis)}</div>
          </div>
          <div class="detail-info-item">
            <div class="detail-info-label">Tahun SKD</div>
            <div class="detail-info-val">${escapeHTML(d.tahun||'2024')}</div>
          </div>
        </div>
      </div>
    </div>

    <div class="detail-footer">
      <button class="detail-hasil-btn ${hasilBtnCls}">${ketText}</button>
    </div>
  `;

  document.getElementById('viewDetailPeserta').style.display = '';

  if (hasSub) {
    setTimeout(() => {
      const ctx = document.getElementById('detailScoreChart');
      if (!ctx) return;
      new Chart(ctx.getContext('2d'), {
        type: 'bar',
        data: {
          labels: ['TWK','TIU','TKP'],
          datasets: [{ data: [d.twk,d.tiu,d.tkp], backgroundColor: ['#B91C1C','#D97706','#0F766E'], borderRadius: 8 }],
        },
        options: {
          responsive: true, maintainAspectRatio: false,
          plugins: { legend: { display: false }, tooltip: { callbacks: { label: c=>`Nilai: ${c.raw}` } } },
          scales: {
            x: { ticks: { font: fontCfg(12), color: '#1A202C' }, grid: { display: false } },
            y: { ticks: { font: fontCfg(11) }, grid: { color: 'rgba(0,0,0,.06)' }, beginAtZero: true },
          },
        },
      });
    }, 100);
  }

  window.scrollTo({ top: 0, behavior: 'smooth' });
}

function closeDetailPage() {
  document.getElementById('viewDetailPeserta').style.display = 'none';
  switchView('peserta');
}

// ==========================================
// MODAL: DETAIL JABATAN
// ==========================================
function showJabatanModal(key) {
  const parts   = key.split('||');
  const jabatan = parts[0];
  const lokasi  = parts[1];
  const jenis   = parts[2];

  const peserta = allData.filter(d => d.jabatan===jabatan && d.lokasi===lokasi && d.jenis===jenis);
  if (!peserta.length) return;

  const plC  = peserta.filter(p => p.ket==='P/L').length;
  const pC   = peserta.filter(p => p.ket==='P').length;
  const tlC  = peserta.filter(p => p.ket==='TL').length;
  const thC  = peserta.filter(p => p.ket==='TH').length;
  const avgS = (peserta.reduce((s,p) => s+p.skd, 0)/peserta.length).toFixed(1);
  const sample = peserta[0];

  const { code, name: displayName } = parseJabatan(jabatan);
  const displayLokasi = lokasi&&lokasi!=='—' ? lokasi : (sample.instansi||'Pemerintah Kota Semarang');
  const n = peserta.length;

  const pctRows = [
    { label:'P/L – Lulus SKB',   count:plC, color:'#1D4ED8' },
    { label:'P – Passing Grade', count:pC,  color:'#059669' },
    { label:'TL – Tidak Lulus',  count:tlC, color:'#DC2626' },
    { label:'TH – Tidak Hadir',  count:thC, color:'#64748B' },
  ].filter(r => r.count > 0);

  const pctHTML = pctRows.map(r => `
    <div class="modal-pct-row">
      <span class="modal-pct-dot" style="background:${r.color}"></span>
      <span class="modal-pct-label">${r.label}</span>
      <span class="modal-pct-val">${r.count}</span>
      <span class="modal-pct-pct">${((r.count/n)*100).toFixed(1)}%</span>
    </div>`).join('');

  const rows = [...peserta].sort((a,b) => b.skd-a.skd).map((p,i) => {
    const ridx = filteredPeserta.indexOf(p);
    const ki   = getKetInfo(p.ket);
    return `<tr>
      <td>${i+1}</td>
      <td>${ridx>=0 ? `<span class="link-nama" onclick="closeJabatanModal();showDetailPage(${ridx})">${escapeHTML(p.nama)}</span>` : escapeHTML(p.nama)}</td>
      <td class="cell-pendidikan">${escapeHTML(p.pendidikan)}</td>
      <td>${p.skd>0?p.skd:'—'}</td>
      <td>${p.twk>0?p.twk:'—'}</td>
      <td>${p.tiu>0?p.tiu:'—'}</td>
      <td>${p.tkp>0?p.tkp:'—'}</td>
      <td><span class="badge-status ${ki.cls}" style="font-size:9.5px">${ki.label}</span></td>
    </tr>`;
  }).join('');

  const pendSet    = [...new Set(peserta.map(p => p.pendidikan).filter(p => p&&p!=='—'))];
  const syaratText = sample.syarat || (pendSet.length ? pendSet.join(', ') : '');
  const ju  = (jenis||'').toUpperCase();
  const jenCls = ju.includes('DISABILITAS') ? 'disabilitas' : ju.includes('TERBAIK') ? 'terbaik' : ju.includes('UMUM') ? 'umum' : 'khusus';

  document.getElementById('modalJabatanContent').innerHTML = `
    <div class="modal-jabatan-code">${escapeHTML(code || '')}</div>
    <div class="modal-jabatan-title">${escapeHTML(displayName || jabatan)}</div>
    <div class="modal-jabatan-meta">
      <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" style="vertical-align:middle;margin-right:4px"><path d="M21 10c0 7-9 13-9 13S3 17 3 10a9 9 0 0118 0z"/><circle cx="12" cy="10" r="3"/></svg>
      ${escapeHTML(truncate(displayLokasi,120))}
      &nbsp;·&nbsp; <span class="badge-formasi ${jenCls}">${escapeHTML(jenis||'UMUM')}</span>
      ${syaratText ? ` &nbsp;·&nbsp; Kualifikasi: ${escapeHTML(truncate(syaratText,100))}` : ''}
    </div>

    <div class="modal-stats-row">
      <div class="modal-stat-box"><div class="val pl">${plC}</div><div class="lbl">P/L</div></div>
      <div class="modal-stat-box"><div class="val lulus">${pC}</div><div class="lbl">P</div></div>
      <div class="modal-stat-box"><div class="val tl">${tlC}</div><div class="lbl">TL</div></div>
      <div class="modal-stat-box"><div class="val th">${thC}</div><div class="lbl">TH</div></div>
      <div class="modal-stat-box"><div class="val rata">${avgS}</div><div class="lbl">Rata SKD</div></div>
      <div class="modal-stat-box"><div class="val" style="color:var(--text-primary)">${n}</div><div class="lbl">Total</div></div>
    </div>

    <h4 class="modal-section-title">Diagram Persentase Kelulusan</h4>
    <div class="modal-chart-area">
      <div class="modal-chart-wrap modal-chart-donut">
        <canvas id="modalJabatanChart"></canvas>
      </div>
      <div class="modal-pct-legend">${pctHTML}</div>
    </div>

    <h4 class="modal-section-title" style="margin-top:8px">Daftar Peserta (${n} orang)</h4>
    <div class="modal-table-wrap">
      <table class="modal-table">
        <thead><tr>
          <th>No</th><th>Nama</th><th>Pendidikan</th>
          <th>SKD</th><th>TWK</th><th>TIU</th><th>TKP</th><th>Ket.</th>
        </tr></thead>
        <tbody>${rows}</tbody>
      </table>
    </div>
  `;

  document.getElementById('modalJabatanOverlay').classList.add('open');
  document.body.style.overflow = 'hidden';

  setTimeout(() => {
    const ctx = document.getElementById('modalJabatanChart');
    if (!ctx) return;
    new Chart(ctx.getContext('2d'), {
      type: 'doughnut',
      data: {
        labels: pctRows.map(r => r.label.split(' – ')[0]),
        datasets: [{ data: pctRows.map(r=>r.count), backgroundColor: pctRows.map(r=>r.color), borderWidth: 3, borderColor: '#fff', hoverOffset: 10 }],
      },
      options: {
        responsive: true, maintainAspectRatio: false,
        cutout: '62%',
        plugins: {
          legend: { display: false },
          tooltip: { callbacks: { label: c=>`${c.label}: ${c.raw} (${((c.raw/n)*100).toFixed(1)}%)` } },
        },
      },
    });
  }, 120);
}

function closeJabatanModal() {
  document.getElementById('modalJabatanOverlay').classList.remove('open');
  document.body.style.overflow = '';
}

// ==========================================
// STAT CARD MODAL
// ==========================================
function showStatModal(type) {
  const total = allData.length;
  const pl    = allData.filter(d => d.ket==='P/L').length;
  const p     = allData.filter(d => d.ket==='P').length;
  const tl    = allData.filter(d => d.ket==='TL').length;
  const th    = allData.filter(d => d.ket==='TH').length;
  const avgSKD = total > 0 ? (allData.reduce((s,d)=>s+d.skd,0)/total).toFixed(1) : 0;

  // Count by jabatan for specific status
  function topByStatus(ket, n=10) {
    const m = {};
    allData.filter(d=>d.ket===ket).forEach(d=>{
      const {name} = parseJabatan(d.jabatan);
      const k = name||d.jabatan;
      m[k] = (m[k]||0)+1;
    });
    return Object.entries(m).sort((a,b)=>b[1]-a[1]).slice(0,n);
  }

  function topByJenis() {
    const m = {};
    allData.forEach(d=>{ m[d.jenis]=(m[d.jenis]||0)+1; });
    return Object.entries(m).sort((a,b)=>b[1]-a[1]);
  }

  const configs = {
    total: {
      title: 'Total Peserta',
      kode: 'SEMUA FORMASI',
      bigVal: total.toLocaleString('id'),
      iconBg: 'linear-gradient(135deg,#0F172A,#1E3A8A,#2563EB)',
      accentColor: '#2563EB',
      icon: `<path d="M17 21v-2a4 4 0 00-4-4H5a4 4 0 00-4 4v2"/><circle cx="9" cy="7" r="4"/><path d="M23 21v-2a4 4 0 00-3-3.87"/><path d="M16 3.13a4 4 0 010 7.75"/>`,
      sectionTitle: 'Sebaran per Jenis Formasi',
      rows: () => {
        const entries = topByJenis();
        const max = entries[0]?.[1]||1;
        return entries.map(([label,count])=>({label,count,max}));
      },
      barColor: '#2563EB',
    },
    pl: {
      title: 'P/L – Lulus SKB',
      kode: 'STATUS KELULUSAN',
      bigVal: pl.toLocaleString('id'),
      iconBg: 'linear-gradient(135deg,#0F2167,#1D4ED8,#60A5FA)',
      accentColor: '#1D4ED8',
      icon: `<polyline points="20 6 9 17 4 12"/>`,
      sectionTitle: 'Top 10 Jabatan dengan Peserta P/L Terbanyak',
      rows: () => { const e=topByStatus('P/L'); const m=e[0]?.[1]||1; return e.map(([label,count])=>({label,count,max:m})); },
      barColor: '#1D4ED8',
    },
    p: {
      title: 'P – Passing Grade',
      kode: 'STATUS KELULUSAN',
      bigVal: p.toLocaleString('id'),
      iconBg: 'linear-gradient(135deg,#064E3B,#059669,#34D399)',
      accentColor: '#059669',
      icon: `<path d="M22 11.08V12a10 10 0 11-5.93-9.14"/><polyline points="22 4 12 14.01 9 11.01"/>`,
      sectionTitle: 'Top 10 Jabatan dengan Peserta Passing Grade Terbanyak',
      rows: () => { const e=topByStatus('P'); const m=e[0]?.[1]||1; return e.map(([label,count])=>({label,count,max:m})); },
      barColor: '#059669',
    },
    tl: {
      title: 'TL – Tidak Lulus',
      kode: 'STATUS KELULUSAN',
      bigVal: tl.toLocaleString('id'),
      iconBg: 'linear-gradient(135deg,#7F1D1D,#B91C1C,#F87171)',
      accentColor: '#B91C1C',
      icon: `<line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/>`,
      sectionTitle: 'Top 10 Jabatan dengan Peserta TL Terbanyak',
      rows: () => { const e=topByStatus('TL'); const m=e[0]?.[1]||1; return e.map(([label,count])=>({label,count,max:m})); },
      barColor: '#B91C1C',
    },
    th: {
      title: 'TH – Tidak Hadir',
      kode: 'STATUS KELULUSAN',
      bigVal: th.toLocaleString('id'),
      iconBg: 'linear-gradient(135deg,#1E293B,#475569,#94A3B8)',
      accentColor: '#475569',
      icon: `<circle cx="12" cy="12" r="10"/><line x1="12" y1="8" x2="12" y2="12"/><line x1="12" y1="16" x2="12.01" y2="16"/>`,
      sectionTitle: 'Top 10 Jabatan dengan Peserta Tidak Hadir Terbanyak',
      rows: () => { const e=topByStatus('TH'); const m=e[0]?.[1]||1; return e.map(([label,count])=>({label,count,max:m})); },
      barColor: '#475569',
    },
    skd: {
      title: 'Top 10 Peserta SKD Tertinggi',
      kode: 'NILAI SKD',
      bigVal: avgSKD,
      iconBg: 'linear-gradient(135deg,#78350F,#D97706,#FCD34D)',
      accentColor: '#D97706',
      icon: `<polygon points="12 2 15.09 8.26 22 9.27 17 14.14 18.18 21.02 12 17.77 5.82 21.02 7 14.14 2 9.27 8.91 8.26 12 2"/>`,
      sectionTitle: null,
      rows: null,
      barColor: '#D97706',
    },
  };

  const cfg = configs[type];
  if (!cfg) return;

  let bodyHTML = '';
  if (type === 'skd') {
    const top10 = [...allData].filter(d=>d.skd>0).sort((a,b)=>b.skd-a.skd).slice(0,10);
    const maxS  = top10[0]?.skd || 1;
    bodyHTML = `<div class="stat-modal-section-title">Top 10 Peserta Nilai SKD Tertinggi</div>`;
    bodyHTML += top10.map((d,i) => {
      const {name} = parseJabatan(d.jabatan);
      const rankCls = i===0?'top1':i===1?'top2':i===2?'top3':'';
      return `<div class="stat-modal-peserta-row">
        <div class="stat-modal-peserta-rank ${rankCls}">${i+1}</div>
        <div class="stat-modal-peserta-name">${escapeHTML(d.nama)}</div>
        <div class="stat-modal-peserta-jabatan">${escapeHTML(truncate(name||d.jabatan,40))}</div>
        <div class="stat-modal-peserta-skd" style="color:${cfg.accentColor}">${d.skd}</div>
      </div>`;
    }).join('');

    // Extra summary row
    bodyHTML += `<div class="stat-modal-section-title" style="margin-top:20px">Ringkasan Nilai SKD</div>`;
    const skdArr = allData.map(d=>d.skd).filter(v=>v>0);
    const minS   = skdArr.length ? Math.min(...skdArr) : 0;
    const maxAll = skdArr.length ? Math.max(...skdArr) : 0;
    bodyHTML += `
      <div class="stat-modal-bar-row"><span class="stat-modal-bar-label">Rata-rata SKD</span><span class="stat-modal-bar-count" style="color:${cfg.accentColor}">${avgSKD}</span></div>
      <div class="stat-modal-bar-row"><span class="stat-modal-bar-label">Nilai Tertinggi</span><span class="stat-modal-bar-count" style="color:${cfg.accentColor}">${maxAll}</span></div>
      <div class="stat-modal-bar-row"><span class="stat-modal-bar-label">Nilai Terendah (hadir)</span><span class="stat-modal-bar-count">${minS}</span></div>
      <div class="stat-modal-bar-row"><span class="stat-modal-bar-label">Total Peserta Hadir</span><span class="stat-modal-bar-count">${skdArr.length.toLocaleString('id')}</span></div>
    `;
  } else {
    const rows = cfg.rows();
    const max  = rows.reduce((m,r)=>Math.max(m,r.max),1);
    bodyHTML = `<div class="stat-modal-section-title">${cfg.sectionTitle}</div>`;
    if (rows.length) {
      bodyHTML += rows.map(r=>`
        <div class="stat-modal-bar-row">
          <span class="stat-modal-bar-label">${escapeHTML(r.label)}</span>
          <div class="stat-modal-bar-track"><div class="stat-modal-bar-fill" style="width:${((r.count/max)*100).toFixed(1)}%;background:${cfg.barColor}"></div></div>
          <span class="stat-modal-bar-count">${r.count}</span>
        </div>`).join('');
    } else {
      bodyHTML += `<p style="color:var(--text-muted);font-style:italic;font-size:13px;padding:12px 0">Tidak ada data.</p>`;
    }
    // Additional context
    if (type === 'total') {
      const bigNum = pl+p;
      const pctLulus = total>0?((bigNum/total)*100).toFixed(1):'0';
      bodyHTML += `<div class="stat-modal-section-title" style="margin-top:20px">Ringkasan Kelulusan</div>
        <div class="stat-modal-bar-row"><span class="stat-modal-bar-label">Total Lulus (P/L + P)</span><div class="stat-modal-bar-track"><div class="stat-modal-bar-fill" style="width:${pctLulus}%;background:#1D4ED8"></div></div><span class="stat-modal-bar-count">${bigNum}</span></div>
        <div class="stat-modal-bar-row"><span class="stat-modal-bar-label">Tidak Lulus (TL)</span><div class="stat-modal-bar-track"><div class="stat-modal-bar-fill" style="width:${total>0?((tl/total)*100).toFixed(1):0}%;background:#B91C1C"></div></div><span class="stat-modal-bar-count">${tl}</span></div>
        <div class="stat-modal-bar-row"><span class="stat-modal-bar-label">Tidak Hadir (TH)</span><div class="stat-modal-bar-track"><div class="stat-modal-bar-fill" style="width:${total>0?((th/total)*100).toFixed(1):0}%;background:#64748B"></div></div><span class="stat-modal-bar-count">${th}</span></div>`;
    }
  }

  document.getElementById('statModalContent').innerHTML = `
    <div class="stat-modal-header">
      <div class="stat-modal-icon-wrap" style="background:${cfg.iconBg}">
        <svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">${cfg.icon}</svg>
      </div>
      <div class="stat-modal-title-group">
        <div class="stat-modal-kode">${cfg.kode}</div>
        <div class="stat-modal-title">${cfg.title}</div>
      </div>
      <div class="stat-modal-big" style="color:${cfg.accentColor}">${cfg.bigVal}</div>
    </div>
    ${bodyHTML}
  `;

  document.getElementById('statCardModalOverlay').classList.add('open');
  document.body.style.overflow = 'hidden';
}

function closeStatModal() {
  document.getElementById('statCardModalOverlay').classList.remove('open');
  document.body.style.overflow = '';
}

// ==========================================
// KEYBOARD
// ==========================================
document.addEventListener('keydown', e => {
  if (e.key === 'Escape') { closeJabatanModal(); closeStatModal(); }
});