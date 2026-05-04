/* ==========================================
   DASHBOARD CPNS 2024 – KOTA SEMARANG
   script.js  (Revised v4 – All revisions applied)
   ========================================== */

// ==========================================
// GLOBAL STATE
// ==========================================
let allData         = [];
let filteredPeserta = [];
let currentView     = 'overview';
let prevView        = 'peserta';
let charts          = {};
let currentPage     = 1;
const PAGE_SIZE     = 50;

// Multi-select filter for Per Jabatan (Set; empty = semua)
let activeJenisFilters = new Set();

// Multi-select filter for the Lulus-per-Jenis chart (Set; empty = semua)
let lulusJenisActiveFilters = new Set();

// Column name aliases
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
// INIT
// ==========================================
document.addEventListener('DOMContentLoaded', () => {
  document.getElementById('loadingState').style.display = 'none';
  document.getElementById('errorState').style.display   = '';
});

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

/**
 * Parse jabatan string into { code, name }.
 * Strategy: split on first " - " separator.
 * If no separator found, the whole string is the name.
 * Ensures code and name are ALWAYS in the correct order.
 */
function parseJabatan(jabatan) {
  const raw = String(jabatan || '').trim();
  const sep = raw.indexOf(' - ');
  if (sep > 0) {
    return {
      code: raw.substring(0, sep).trim(),
      name: raw.substring(sep + 3).trim(),
    };
  }
  // Fallback: no " - " separator
  return { code: '', name: raw };
}

// ==========================================
// XLSX PARSING
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
  const headers = rawHeaders.map(h => String(h || '').trim());

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
    if (rawKet === 'P/L' || rawKet === 'P/L (LULUS SKB)') {
      ketCode = 'P/L';
    } else if (rawKet === 'P' || rawKet === 'LULUS' || rawKet === 'PASS' || rawKet === 'P (PASSING GRADE)') {
      ketCode = 'P';
    } else if (rawKet === 'TL' || rawKet === 'TIDAK LULUS') {
      ketCode = 'TL';
    } else if (rawKet === 'TH' || rawKet === 'TIDAK HADIR') {
      ketCode = 'TH';
    } else if (rawKet === 'TMS' || rawKet === 'GUGUR') {
      ketCode = 'TMS';
    } else if (rawKet === 'DIS' || rawKet === 'DISKUALIFIKASI') {
      ketCode = 'DIS';
    }

    const status = (ketCode === 'P/L' || ketCode === 'P') ? 'LULUS'
                 : ketCode === 'TL' ? 'TL'
                 : ketCode === 'TH' ? 'TH'
                 : ketCode;

    let jenisRaw = get('JENIS').replace(/\s+/g, ' ').trim();
    const jenisM = jenisRaw.match(/^\d+\s*-\s*(.+)$/);
    const jenis  = jenisM ? jenisM[1].trim() : (jenisRaw || 'UMUM');

    const tahunRaw = get('TAHUN');
    const tahun = tahunRaw ? String(tahunRaw).replace(/\.0$/, '') : '2024';

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

  // Reset filters on new upload
  activeJenisFilters.clear();
  lulusJenisActiveFilters.clear();

  const jabatanSet = [...new Set(allData.map(d => d.jabatan))].sort();
  const selJab = document.getElementById('filterJabatanSelect');
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
// UPDATE JENIS CHIPS (multi-select aware)
// ==========================================
function updateJenisChips(jenisSet) {
  const container = document.getElementById('jabatanFilterChips');
  if (!container) return;

  const hasDisabilitas = jenisSet.some(j => j.toUpperCase().includes('DISABILITAS'));
  const hasTerbaik     = jenisSet.some(j => j.toUpperCase().includes('TERBAIK'));
  const hasUmum        = jenisSet.some(j => j.toUpperCase().includes('UMUM'));

  const isSemua = activeJenisFilters.size === 0;

  let html = `<button class="filter-chip${isSemua ? ' active' : ''}" data-jenis="semua" onclick="setJenisFilter('semua',this)">
    Semua
  </button>`;

  if (hasUmum) {
    const isActive = [...activeJenisFilters].some(f => f.toUpperCase().includes('UMUM'));
    html += `<button class="filter-chip${isActive ? ' active' : ''}" data-jenis="UMUM" onclick="setJenisFilter('UMUM',this)">
      <span class="chip-dot"></span>Umum
    </button>`;
  }
  if (hasDisabilitas) {
    const isActive = [...activeJenisFilters].some(f => f.toUpperCase().includes('DISABILITAS'));
    html += `<button class="filter-chip${isActive ? ' active' : ''}" data-jenis="DISABILITAS" onclick="setJenisFilter('DISABILITAS',this)">
      <span class="chip-dot"></span>Disabilitas
    </button>`;
  }
  if (hasTerbaik) {
    const isActive = [...activeJenisFilters].some(f => f.toUpperCase().includes('TERBAIK'));
    html += `<button class="filter-chip${isActive ? ' active' : ''}" data-jenis="LULUSAN TERBAIK" onclick="setJenisFilter('LULUSAN TERBAIK',this)">
      <span class="chip-dot"></span>Lulusan Terbaik
    </button>`;
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
  ['viewOverview', 'viewPerJabatan', 'viewPeserta', 'viewDetailPeserta'].forEach(id => {
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
  }
}

// ==========================================
// OVERVIEW & ANALISIS
// ==========================================
function renderOverview() {
  const total  = allData.length;
  const pl     = allData.filter(d => d.ket === 'P/L').length;
  const p      = allData.filter(d => d.ket === 'P').length;
  const lulus  = pl + p;
  const tl     = allData.filter(d => d.ket === 'TL').length;
  const th     = allData.filter(d => d.ket === 'TH').length;
  const jabSet = [...new Set(allData.map(d => d.jabatan))];
  const avgSKD = total > 0 ? (allData.reduce((s,d) => s+d.skd, 0)/total).toFixed(1) : 0;
  const maxSKD = allData.length ? Math.max(...allData.map(d => d.skd)) : 0;

  document.getElementById('overviewDesc').textContent =
    `${total.toLocaleString('id')} peserta dari ${jabSet.length} jabatan formasi – SKD CPNS 2024 Kota Semarang`;

  document.getElementById('statCards').innerHTML = `
    <div class="stat-card blue">
      <div class="stat-icon"><svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M17 21v-2a4 4 0 00-4-4H5a4 4 0 00-4 4v2"/><circle cx="9" cy="7" r="4"/><path d="M23 21v-2a4 4 0 00-3-3.87"/><path d="M16 3.13a4 4 0 010 7.75"/></svg></div>
      <div class="stat-info">
        <div class="stat-label">Total Peserta</div>
        <div class="stat-value">${total.toLocaleString('id')}</div>
        <div class="stat-sub">${jabSet.length} jabatan formasi</div>
      </div>
    </div>
    <div class="stat-card blue">
      <div class="stat-icon"><svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><polyline points="20 6 9 17 4 12"/></svg></div>
      <div class="stat-info">
        <div class="stat-label">P/L – Lulus SKB</div>
        <div class="stat-value">${pl.toLocaleString('id')}</div>
        <div class="stat-sub">${total>0?((pl/total)*100).toFixed(1):0}% dari total</div>
      </div>
    </div>
    <div class="stat-card plg">
      <div class="stat-icon"><svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M22 11.08V12a10 10 0 11-5.93-9.14"/><polyline points="22 4 12 14.01 9 11.01"/></svg></div>
      <div class="stat-info">
        <div class="stat-label">P – Passing Grade</div>
        <div class="stat-value">${p.toLocaleString('id')}</div>
        <div class="stat-sub">${total>0?((p/total)*100).toFixed(1):0}% dari total</div>
      </div>
    </div>
    <div class="stat-card red">
      <div class="stat-icon"><svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/></svg></div>
      <div class="stat-info">
        <div class="stat-label">TL – Tidak Lulus</div>
        <div class="stat-value">${tl.toLocaleString('id')}</div>
        <div class="stat-sub">${total>0?((tl/total)*100).toFixed(1):0}% dari total</div>
      </div>
    </div>
    <div class="stat-card grey">
      <div class="stat-icon"><svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="12" cy="12" r="10"/><line x1="12" y1="8" x2="12" y2="12"/><line x1="12" y1="16" x2="12.01" y2="16"/></svg></div>
      <div class="stat-info">
        <div class="stat-label">TH – Tidak Hadir</div>
        <div class="stat-value">${th.toLocaleString('id')}</div>
        <div class="stat-sub">${total>0?((th/total)*100).toFixed(1):0}% dari total</div>
      </div>
    </div>
    <div class="stat-card yellow">
      <div class="stat-icon"><svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><polygon points="12 2 15.09 8.26 22 9.27 17 14.14 18.18 21.02 12 17.77 5.82 21.02 7 14.14 2 9.27 8.91 8.26 12 2"/></svg></div>
      <div class="stat-info">
        <div class="stat-label">SKD Tertinggi</div>
        <div class="stat-value">${maxSKD}</div>
        <div class="stat-sub">Rata-rata: ${avgSKD}</div>
      </div>
    </div>
  `;

  renderCharts();
}

// ==========================================
// CHART HELPERS
// ==========================================
const CC = {
  PALETTE: ['#2563EB','#16A34A','#B91C1C','#64748B','#D97706','#6D28D9','#0F766E','#C2410C','#BE185D','#0E7490'],
};

function destroyChart(key) {
  if (charts[key]) { charts[key].destroy(); delete charts[key]; }
}

function fontCfg(size=11) {
  return { family: "'Jost', 'Segoe UI', sans-serif", size };
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
  const tms   = allData.filter(d => d.ket === 'TMS').length;
  const dis   = allData.filter(d => d.ket === 'DIS').length;

  // 1. Pie: Status Kelulusan
  destroyChart('status');
  charts.status = new Chart(document.getElementById('chartStatus').getContext('2d'), {
    type: 'pie',
    data: {
      labels: ['P/L (Lulus SKB)', 'P (Passing Grade)', 'TL (Tidak Lulus)', 'TH (Tidak Hadir)', 'TMS', 'DIS'],
      datasets: [{
        data: [pl, p, tl, th, tms, dis].filter(v => v > 0),
        backgroundColor: ['#2563EB','#16A34A','#B91C1C','#64748B','#D97706','#831843']
                          .slice(0, [pl,p,tl,th,tms,dis].filter(v=>v>0).length),
        borderWidth: 2,
        borderColor: '#fff',
        hoverOffset: 12,
      }],
    },
    options: {
      responsive: true, maintainAspectRatio: false,
      plugins: {
        legend: {
          position: 'bottom',
          labels: {
            padding: 16, font: fontCfg(12),
            generateLabels: (chart) => {
              const rawLabels = ['P/L (Lulus SKB)', 'P (Passing Grade)', 'TL (Tidak Lulus)', 'TH (Tidak Hadir)', 'TMS', 'DIS'];
              const rawData   = [pl, p, tl, th, tms, dis];
              const colors    = ['#2563EB','#16A34A','#B91C1C','#64748B','#D97706','#831843'];
              return rawLabels.map((label, i) => ({
                text: `${label.split(' ')[0]}: ${rawData[i]}`,
                fillStyle: colors[i],
                hidden: rawData[i] === 0,
                index: i,
              })).filter((_,i) => rawData[i] > 0);
            },
          },
        },
        tooltip: {
          callbacks: {
            label: ctx => ` ${ctx.label}: ${ctx.raw.toLocaleString('id')} (${((ctx.raw/total)*100).toFixed(1)}%)`
          }
        },
      },
    },
  });

  // 2. Doughnut: Jenis Formasi
  const jenisCount = {};
  allData.forEach(d => { jenisCount[d.jenis] = (jenisCount[d.jenis]||0)+1; });
  const jenisEntries = Object.entries(jenisCount).sort((a,b) => b[1]-a[1]);

  destroyChart('jenis');
  charts.jenis = new Chart(document.getElementById('chartJenis').getContext('2d'), {
    type: 'doughnut',
    data: {
      labels: jenisEntries.map(e => e[0]),
      datasets: [{ data: jenisEntries.map(e => e[1]), backgroundColor: CC.PALETTE, borderWidth:0, hoverOffset:10 }],
    },
    options: {
      responsive:true, maintainAspectRatio:false, cutout:'60%',
      plugins: {
        legend: { position:'bottom', labels:{ padding:14, font:fontCfg(11) } },
        tooltip: { callbacks:{ label: ctx => ` ${ctx.label}: ${ctx.raw.toLocaleString('id')} peserta` } },
      },
    },
  });

  // 3. Lulus per Jenis chart + interactive multi-select filter row
  const jenisNames = jenisEntries.map(e => e[0]);
  buildLulusJenisFilterRow(jenisNames);
  lulusJenisActiveFilters.clear(); // reset on re-render
  renderLulusJenisChart(new Set()); // show all initially

  // 4. Horizontal Bar: Top 10 jabatan by avg SKD
  const jabMap = {};
  allData.forEach(d => { if (!jabMap[d.jabatan]) jabMap[d.jabatan]=[]; jabMap[d.jabatan].push(d.skd); });
  const jabRanked = Object.entries(jabMap)
    .map(([name,vals]) => ({ name, avg: vals.reduce((a,b)=>a+b,0)/vals.length }))
    .sort((a,b) => b.avg-a.avg).slice(0,10);

  destroyChart('jabatan');
  charts.jabatan = new Chart(document.getElementById('chartJabatan').getContext('2d'), {
    type:'bar',
    data: {
      labels: jabRanked.map(j => truncate(parseJabatan(j.name).name || j.name, 38)),
      datasets: [{
        label:'Rata-rata SKD',
        data: jabRanked.map(j => +j.avg.toFixed(1)),
        backgroundColor: jabRanked.map((_,i) => i===0?'#2563EB':i<3?'#60A5FA':'rgba(37,99,235,.35)'),
        borderRadius:6,
      }],
    },
    options: {
      indexAxis:'y', responsive:true, maintainAspectRatio:false,
      plugins:{ legend:{ display:false } },
      scales: {
        x:{ grid:{ color:'rgba(0,0,0,.04)' }, ticks:{ font:fontCfg(11) } },
        y:{ ticks:{ font:fontCfg(10), color:'#4a3730' } },
      },
    },
  });

  // 5. Horizontal Bar: Top 10 jabatan by count
  const jabCount = Object.entries(jabMap)
    .map(([name,vals]) => ({ name, count:vals.length }))
    .sort((a,b) => b.count-a.count).slice(0,10);

  destroyChart('jabatanCount');
  charts.jabatanCount = new Chart(document.getElementById('chartJabatanCount').getContext('2d'), {
    type:'bar',
    data: {
      labels: jabCount.map(j => truncate(parseJabatan(j.name).name || j.name, 38)),
      datasets: [{
        label:'Jumlah Peserta',
        data: jabCount.map(j => j.count),
        backgroundColor: jabCount.map((_,i) => i===0?'#6D28D9':i<3?'#A78BFA':'rgba(109,40,217,.35)'),
        borderRadius:6,
      }],
    },
    options: {
      indexAxis:'y', responsive:true, maintainAspectRatio:false,
      plugins:{ legend:{ display:false } },
      scales: {
        x:{ grid:{ color:'rgba(0,0,0,.04)' }, ticks:{ font:fontCfg(11) } },
        y:{ ticks:{ font:fontCfg(10), color:'#4a3730' } },
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
    for (let b=minS; b<maxS; b+=binSz) {
      dbins.push({ label:`${b}–${b+binSz-1}`, min:b, max:b+binSz, count:0 });
    }
    allSkd.forEach(v => { const bn=dbins.find(b=>v>=b.min&&v<b.max); if(bn) bn.count++; });

    destroyChart('distribusi');
    charts.distribusi = new Chart(document.getElementById('chartDistribusi').getContext('2d'), {
      type:'bar',
      data: {
        labels: dbins.map(b => b.label),
        datasets: [{
          label:'Jumlah Peserta',
          data: dbins.map(b => b.count),
          backgroundColor: 'rgba(37,99,235,.60)',
          borderColor:'#2563EB', borderWidth:1.5, borderRadius:4,
        }],
      },
      options: {
        responsive:true, maintainAspectRatio:false,
        plugins:{ legend:{ display:false } },
        scales: {
          x:{ ticks:{ font:fontCfg(10), maxRotation:45 } },
          y:{ ticks:{ font:fontCfg(11) }, grid:{ color:'rgba(0,0,0,.04)' } },
        },
      },
    });
  }

  // 7. Bar: rata-rata TWK · TIU · TKP
  const hasSubtes = allData.some(d => d.twk>0||d.tiu>0||d.tkp>0);
  destroyChart('subtes');
  if (hasSubtes) {
    const avgTWK = avg(allData.map(d => d.twk));
    const avgTIU = avg(allData.map(d => d.tiu));
    const avgTKP = avg(allData.map(d => d.tkp));
    charts.subtes = new Chart(document.getElementById('chartSubtes').getContext('2d'), {
      type:'bar',
      data: {
        labels: ['TWK (Wawasan Kebangsaan)','TIU (Intelejensi Umum)','TKP (Karakteristik Pribadi)'],
        datasets: [{
          label:'Rata-rata Nilai',
          data: [+avgTWK.toFixed(1), +avgTIU.toFixed(1), +avgTKP.toFixed(1)],
          backgroundColor: ['#B91C1C','#D97706','#0F766E'],
          borderRadius:8,
        }],
      },
      options: {
        responsive:true, maintainAspectRatio:false,
        plugins:{ legend:{ display:false } },
        scales: {
          x:{ ticks:{ font:fontCfg(10.5), maxRotation:20 }, grid:{ display:false } },
          y:{ ticks:{ font:fontCfg(11) }, grid:{ color:'rgba(0,0,0,.04)' } },
        },
      },
    });
  }

  // 8. Horizontal Bar: Top 10 Pendidikan
  const pendCount = {};
  allData.forEach(d => {
    if (d.pendidikan && d.pendidikan !== '—') {
      pendCount[d.pendidikan] = (pendCount[d.pendidikan]||0)+1;
    }
  });
  const pendRanked = Object.entries(pendCount).sort((a,b) => b[1]-a[1]).slice(0,10);

  destroyChart('pendidikan');
  charts.pendidikan = new Chart(document.getElementById('chartPendidikan').getContext('2d'), {
    type:'bar',
    data: {
      labels: pendRanked.map(p => p[0]),
      datasets: [{
        label:'Jumlah Peserta',
        data: pendRanked.map(p => p[1]),
        backgroundColor: pendRanked.map((_,i) => CC.PALETTE[i%CC.PALETTE.length]),
        borderRadius:6,
      }],
    },
    options: {
      indexAxis:'y', responsive:true, maintainAspectRatio:false,
      plugins:{ legend:{ display:false } },
      scales: {
        x:{ ticks:{ font:fontCfg(11) }, grid:{ color:'rgba(0,0,0,.04)' } },
        y:{ ticks:{ font:fontCfg(10.5), color:'#4a3730' } },
      },
    },
  });
}

// ==========================================
// LULUS PER JENIS – Multi-select Filter + Chart
// ==========================================

/**
 * Build the interactive filter row for the "Lulus per Jenis" chart.
 * Chips support multi-select: click to toggle individual jenis.
 * "Semua" clears selection (shows all).
 */
function buildLulusJenisFilterRow(jenisNames) {
  const row = document.getElementById('lulusJenisFilterRow');
  if (!row) return;

  const dotColors = {
    umum:    '#D97706',
    terbaik: '#16A34A',
    dis:     '#2563EB',
    khusus:  '#B91C1C',
  };

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

/**
 * Toggle a jenis in the lulusJenisActiveFilters Set and re-render the chart.
 * "SEMUA" clears all filters.
 */
function toggleLulusJenisFilter(jenis) {
  const row = document.getElementById('lulusJenisFilterRow');
  if (!row) return;

  if (jenis === 'SEMUA') {
    lulusJenisActiveFilters.clear();
  } else {
    if (lulusJenisActiveFilters.has(jenis)) {
      lulusJenisActiveFilters.delete(jenis);
    } else {
      lulusJenisActiveFilters.add(jenis);
    }
  }

  // Update chip states
  row.querySelectorAll('.cfc-chip').forEach(chip => {
    const cj = chip.dataset.ljJenis;
    const cls = chip.classList;
    cls.remove('active', 'inactive');
    if (cj === 'SEMUA') {
      cls.add(lulusJenisActiveFilters.size === 0 ? 'active' : 'inactive');
    } else {
      cls.add(lulusJenisActiveFilters.has(cj) ? 'active' : 'inactive');
    }
  });

  renderLulusJenisChart(lulusJenisActiveFilters);
}

// Delegated click for lulus-jenis chart filter chips
document.addEventListener('click', function(e) {
  const chip = e.target.closest('.cfc-chip[data-lj-jenis]');
  if (chip) {
    toggleLulusJenisFilter(chip.dataset.ljJenis);
  }
});

/**
 * Render (or re-render) the "Lulus per Jenis" grouped bar chart.
 * activeFilters: Set of jenis strings to show; empty = show all.
 */
function renderLulusJenisChart(activeFilters) {
  const jenisCount = {};
  allData.forEach(d => { jenisCount[d.jenis] = (jenisCount[d.jenis]||0)+1; });
  const allJenisNames = Object.entries(jenisCount).sort((a,b) => b[1]-a[1]).map(e => e[0]);

  const filteredJenis = activeFilters.size === 0
    ? allJenisNames
    : allJenisNames.filter(j => activeFilters.has(j));

  const plPerJ  = filteredJenis.map(j => allData.filter(d => d.jenis===j && d.ket==='P/L').length);
  const pPerJ   = filteredJenis.map(j => allData.filter(d => d.jenis===j && d.ket==='P').length);
  const tlPerJ  = filteredJenis.map(j => allData.filter(d => d.jenis===j && d.ket==='TL').length);
  const thPerJ  = filteredJenis.map(j => allData.filter(d => d.jenis===j && d.ket==='TH').length);

  destroyChart('lulusPerJenis');
  const ctx = document.getElementById('chartLulusPerJenis');
  if (!ctx) return;

  charts.lulusPerJenis = new Chart(ctx.getContext('2d'), {
    type: 'bar',
    data: {
      labels: filteredJenis,
      datasets: [
        { label:'P/L',          data:plPerJ,  backgroundColor:'#2563EB', borderRadius:5 },
        { label:'P (PG)',       data:pPerJ,   backgroundColor:'#16A34A', borderRadius:5 },
        { label:'Tidak Lulus',  data:tlPerJ,  backgroundColor:'#B91C1C', borderRadius:5 },
        { label:'Tidak Hadir',  data:thPerJ,  backgroundColor:'#64748B', borderRadius:5 },
      ],
    },
    options: {
      responsive:true, maintainAspectRatio:false,
      plugins: {
        legend:{ position:'top', labels:{ font:fontCfg(12), padding:16 } },
        tooltip: {
          callbacks: {
            label: ctx => ` ${ctx.dataset.label}: ${ctx.raw} peserta`
          }
        },
      },
      scales: {
        x:{ ticks:{ font:fontCfg(10.5), maxRotation:20 }, grid:{ display:false } },
        y:{ ticks:{ font:fontCfg(11) }, grid:{ color:'rgba(0,0,0,.04)' } },
      },
    },
  });
}

// ==========================================
// PER JABATAN – Multi-select filter by jenis
// ==========================================

/**
 * Toggle a jenis in activeJenisFilters and re-render jabatan grid.
 * "semua" resets all filters.
 */
function setJenisFilter(jenis, btnEl) {
  if (jenis === 'semua') {
    activeJenisFilters.clear();
  } else {
    if (activeJenisFilters.has(jenis)) {
      activeJenisFilters.delete(jenis);
    } else {
      activeJenisFilters.add(jenis);
    }
  }

  // Update all chip visual states
  document.querySelectorAll('#jabatanFilterChips .filter-chip').forEach(b => {
    const dj = b.dataset.jenis;
    if (dj === 'semua') {
      b.classList.toggle('active', activeJenisFilters.size === 0);
    } else {
      b.classList.toggle('active', activeJenisFilters.has(dj));
    }
  });

  renderPerJabatan(document.getElementById('searchJabatan')?.value || '');
}

function groupByJabatanLokasi(data) {
  return data.reduce((map, d) => {
    const key = d.jabatan + '||' + d.lokasi;
    (map[key] = map[key] || []).push(d);
    return map;
  }, {});
}

function renderPerJabatan(filter='') {
  // Multi-select AND logic: show jabatan that have ALL selected jenis types
  // First build jabatan groups from all data, then filter groups by jenis presence
  const allJabMap = groupByJabatanLokasi(allData);

  let keys = Object.keys(allJabMap);

  if (activeJenisFilters.size > 0) {
    keys = keys.filter(key => {
      const pesertaInJab = allJabMap[key];
      const jenisInJab = new Set(pesertaInJab.map(p => p.jenis.toUpperCase()));
      // AND logic: every selected filter must be represented in this jabatan
      return [...activeJenisFilters].every(f =>
        [...jenisInJab].some(j => j.includes(f.toUpperCase()))
      );
    });
  }

  // For display, build jabMap only for the filtered keys
  const jabMap = {};
  keys.forEach(k => { jabMap[k] = allJabMap[k]; });

  keys = keys
    .filter(k => !filter || k.toLowerCase().includes(filter.toLowerCase()))
    .sort();

  // Build active filter label for description
  let filterLabel = '';
  if (activeJenisFilters.size > 0) {
    filterLabel = ` · Filter: ${[...activeJenisFilters].join(', ')}`;
  }

  document.getElementById('jabatanDesc').textContent =
    `${keys.length} jabatan formasi ditampilkan${filterLabel}`;

  const grid = document.getElementById('jabatanGrid');
  if (!keys.length) {
    grid.innerHTML = `<div class="no-results" style="grid-column:1/-1">Tidak ada jabatan yang cocok dengan filter ini.</div>`;
    return;
  }

  grid.innerHTML = keys.map(key => {
    const [jabatan, lokasi] = key.split('||');
    const peserta = jabMap[key];
    const plC     = peserta.filter(p => p.ket==='P/L').length;
    const pC      = peserta.filter(p => p.ket==='P').length;
    const tlC     = peserta.filter(p => p.ket==='TL').length;
    const thC     = peserta.filter(p => p.ket==='TH').length;
    const avgS    = (peserta.reduce((s,p) => s+p.skd, 0)/peserta.length).toFixed(0);
    const sample  = peserta[0];

    // Always parse code and name consistently
    const { code, name: displayName } = parseJabatan(jabatan);

    const displayLokasi = lokasi && lokasi !== '—' ? lokasi : (sample.instansi||'Pemerintah Kota Semarang');

    const jenisSet = [...new Set(peserta.map(p => p.jenis))];
    const badges   = jenisSet.map(j => {
      const ju = j.toUpperCase();
      const cls = ju.includes('DISABILITAS') ? 'disabilitas'
                : ju.includes('TERBAIK')     ? 'terbaik'
                : ju.includes('UMUM')        ? 'umum'
                : 'khusus';
      return `<span class="badge-formasi ${cls}">${j}</span>`;
    }).join('');

    const pendSet    = [...new Set(peserta.map(p => p.pendidikan).filter(p => p&&p!=='—'))];
    const syaratText = sample.syarat || (pendSet.length ? pendSet.join(', ') : '');

    return `
      <div class="jabatan-card" data-key="${escapeAttr(key)}" tabindex="0" role="button" aria-label="Lihat detail ${escapeAttr(displayName || jabatan)}">
        <div class="jabatan-card-header">
          ${code ? `<div class="jabatan-code">${escapeHTML(code)}</div>` : ''}
          <div class="jabatan-name">${escapeHTML(displayName || jabatan)}</div>
          <div class="jabatan-meta">
            <svg width="11" height="11" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><path d="M21 10c0 7-9 13-9 13S3 17 3 10a9 9 0 0118 0z"/><circle cx="12" cy="10" r="3"/></svg>
            ${escapeHTML(displayLokasi)}
          </div>
          ${syaratText ? `
          <div class="jabatan-syarat">
            <svg width="11" height="11" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><path d="M22 10v6M2 10l10-5 10 5-10 5z"/><path d="M6 12v5c3 3 9 3 12 0v-5"/></svg>
            Kualifikasi: ${escapeHTML(truncate(syaratText, 80))}
          </div>` : ''}
          <div style="margin-top:8px">${badges}</div>
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
          <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M17 21v-2a4 4 0 00-4-4H5a4 4 0 00-4 4v2"/><circle cx="9" cy="7" r="4"/></svg>
          ${peserta.length} peserta · Klik untuk detail
        </div>
      </div>
    `;
  }).join('');
}

function filterJabatan() {
  renderPerJabatan(document.getElementById('searchJabatan').value);
}

// Delegated click + keyboard handler for jabatan cards (data-key approach avoids inline JSON escaping issues)
document.addEventListener('click', function(e) {
  const card = e.target.closest('.jabatan-card[data-key]');
  if (card) {
    showJabatanModal(card.dataset.key);
  }
});


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
    'P/L': { label: 'P/L',  cls: 'pl'    },
    'P':   { label: 'P',    cls: 'lulus'  },
    'TL':  { label: 'TL',   cls: 'tl'    },
    'TH':  { label: 'TH',   cls: 'th'    },
    'TMS': { label: 'TMS',  cls: 'tms'   },
    'DIS': { label: 'DIS',  cls: 'dis'   },
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
    const absIdx     = start+i;
    const ki         = getKetInfo(d.ket);
    const jenisClass = d.jenis.toUpperCase().includes('DISABILITAS') ? 'disabilitas'
                     : d.jenis.toUpperCase().includes('TERBAIK')     ? 'terbaik'
                     : d.jenis.toUpperCase().includes('UMUM')        ? 'umum'
                     : 'khusus';
    const { name: displayJab } = parseJabatan(d.jabatan);
    return `
      <tr>
        <td>${absIdx+1}</td>
        <td><span class="link-nama" onclick="showDetailPage(${absIdx})">${escapeHTML(d.nama)}</span></td>
        <td><span class="jabatan-chip" title="${escapeAttr(d.jabatan)}">${escapeHTML(displayJab || d.jabatan)}</span></td>
        <td class="cell-pendidikan">${escapeHTML(d.pendidikan)}</td>
        <td><span class="badge-formasi ${jenisClass}" style="font-size:9px;padding:2px 8px">${escapeHTML(d.jenis)}</span></td>
        <td class="cell-tahun">${escapeHTML(d.tahun||'—')}</td>
        <td>${d.twk>0?d.twk:'—'}</td>
        <td>${d.tiu>0?d.tiu:'—'}</td>
        <td>${d.tkp>0?d.tkp:'—'}</td>
        <td class="nilai-skd">${d.skd>0?d.skd:'—'}</td>
        <td><span class="badge-status ${ki.cls}">${ki.label}</span></td>
      </tr>
    `;
  }).join('');

  document.getElementById('tableInfo').textContent =
    `Halaman ${currentPage} dari ${pages} · ${total.toLocaleString('id')} hasil`;
  renderPagination(pages);
}

function renderPagination(pages) {
  const pg  = document.getElementById('pagination');
  if (pages <= 1) { pg.innerHTML=''; return; }
  const cur = currentPage;
  let html  = '';
  const btn = (p,label,disabled=false,active=false) =>
    `<button class="pg-btn${active?' active':''}${disabled?' disabled':''}"
      ${disabled?'disabled':`onclick="goPage(${p})"`}>${label}</button>`;

  html += btn(cur-1,'‹ Prev',cur===1);
  const show = new Set([1,pages,cur,cur-1,cur+1].filter(p=>p>=1&&p<=pages));
  let prev = 0;
  [...show].sort((a,b)=>a-b).forEach(p => {
    if (prev&&p-prev>1) html+=`<span class="pg-ellipsis">…</span>`;
    html+=btn(p,p,false,p===cur);
    prev=p;
  });
  html += btn(cur+1,'Next ›',cur===pages);
  pg.innerHTML = html;
}

function goPage(p) {
  currentPage = p;
  renderPesertaTable();
  document.querySelector('.table-wrap')?.scrollIntoView({ behavior:'smooth', block:'start' });
}

// ==========================================
// DETAIL PAGE
// ==========================================
function showDetailPage(absIdx) {
  const d = filteredPeserta[absIdx];
  if (!d) return;

  prevView = 'peserta';

  ['viewOverview','viewPerJabatan','viewPeserta'].forEach(id => {
    const el = document.getElementById(id);
    if (el) el.style.display='none';
  });
  document.querySelectorAll('.nav-btn').forEach(b=>b.classList.remove('active'));

  const hasSub   = d.twk>0||d.tiu>0||d.tkp>0;
  const ki       = getKetInfo(d.ket);

  const btnClsMap = {
    'P/L':'detail-hasil-pl','P':'detail-hasil-lulus',
    'TL':'detail-hasil-tl','TH':'detail-hasil-th',
    'TMS':'detail-hasil-tms','DIS':'detail-hasil-dis',
  };
  const hasilBtnCls = btnClsMap[d.ket] || 'detail-hasil-th';

  const hasilTextMap = {
    'P/L':'HASIL AKHIR: P/L – Lulus SKB',
    'P':'HASIL AKHIR: P – Passing Grade',
    'TL':'HASIL AKHIR: TL – Tidak Lulus',
    'TH':'HASIL AKHIR: TH – Tidak Hadir',
    'TMS':'HASIL AKHIR: TMS – Gugur Syarat',
    'DIS':'HASIL AKHIR: DIS – Diskualifikasi',
  };
  const ketText = hasilTextMap[d.ket] || 'HASIL AKHIR: ' + (d.ket||'—');

  const scoreRows = hasSub ? `
    <div class="detail-score-row"><span class="detail-score-label">TWK – Wawasan Kebangsaan</span><span class="detail-score-val">${d.twk}</span></div>
    <div class="detail-score-row"><span class="detail-score-label">TIU – Intelejensi Umum</span><span class="detail-score-val">${d.tiu}</span></div>
    <div class="detail-score-row"><span class="detail-score-label">TKP – Karakteristik Pribadi</span><span class="detail-score-val">${d.tkp}</span></div>
  ` : '<p style="color:var(--text-muted);font-style:italic;font-size:13px;font-family:var(--font-serif)">Data nilai sub-tes tidak tersedia.</p>';

  const { code, name: displayName } = parseJabatan(d.jabatan);

  document.getElementById('detailPageContent').innerHTML = `
    <div class="detail-hero">
      <div class="detail-hero-left">
        <h2 class="detail-nama">${escapeHTML(d.nama)}</h2>
        ${d.nopes ? `<div class="detail-nopes">No. Peserta: ${escapeHTML(d.nopes)}</div>` : ''}
        <div class="detail-tahun-badge">SKD ${escapeHTML(d.tahun||'2024')}</div>
      </div>
      <div class="detail-hero-right">
        <div class="detail-total-label">Total Skor SKD</div>
        <div class="detail-total-val">${d.skd>0?d.skd:'—'}</div>
      </div>
    </div>

    <div class="detail-body">
      <div class="detail-section">
        <h3 class="detail-section-title">Rincian Nilai SKD</h3>
        <div class="detail-scores">${scoreRows}</div>
        ${hasSub ? `<div class="detail-score-chart-wrap"><canvas id="detailScoreChart"></canvas></div>` : ''}
      </div>

      <div class="detail-section">
        <h3 class="detail-section-title">Informasi Formasi</h3>
        <div class="detail-info-grid">
          <div class="detail-info-item">
            <div class="detail-info-label">Jabatan Formasi</div>
            <div class="detail-info-val">${escapeHTML(displayName || d.jabatan)}</div>
          </div>
          ${code ? `<div class="detail-info-item">
            <div class="detail-info-label">Kode Formasi</div>
            <div class="detail-info-val">${escapeHTML(code)}</div>
          </div>` : ''}
          <div class="detail-info-item">
            <div class="detail-info-label">Instansi</div>
            <div class="detail-info-val">${escapeHTML(d.instansi)}</div>
          </div>
          <div class="detail-info-item">
            <div class="detail-info-label">Lokasi / Unit Kerja</div>
            <div class="detail-info-val">${escapeHTML(d.lokasi)}</div>
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
        type:'bar',
        data: {
          labels:['TWK','TIU','TKP'],
          datasets:[{
            data:[d.twk,d.tiu,d.tkp],
            backgroundColor:['#B91C1C','#D97706','#0F766E'],
            borderRadius:8,
          }],
        },
        options: {
          responsive:true, maintainAspectRatio:false,
          plugins:{ legend:{ display:false }, tooltip:{ callbacks:{ label: c=>`Nilai: ${c.raw}` } } },
          scales: {
            x:{ ticks:{ font:fontCfg(12), color:'#1C1410' }, grid:{ display:false } },
            y:{ ticks:{ font:fontCfg(11) }, grid:{ color:'rgba(0,0,0,.06)' }, beginAtZero:true },
          },
        },
      });
    }, 100);
  }

  window.scrollTo({ top:0, behavior:'smooth' });
}

function closeDetailPage() {
  document.getElementById('viewDetailPeserta').style.display = 'none';
  switchView('peserta');
}

// ==========================================
// MODAL: DETAIL JABATAN
// Requirement: Doughnut % chart + Peserta list
// ==========================================
function showJabatanModal(key) {
  const [jabatan, lokasi] = key.split('||');
  const peserta = allData.filter(d => d.jabatan===jabatan && d.lokasi===lokasi);
  if (!peserta.length) return;

  const plC  = peserta.filter(p => p.ket==='P/L').length;
  const pC   = peserta.filter(p => p.ket==='P').length;
  const tlC  = peserta.filter(p => p.ket==='TL').length;
  const thC  = peserta.filter(p => p.ket==='TH').length;
  const avgS = (peserta.reduce((s,p) => s+p.skd, 0)/peserta.length).toFixed(1);
  const sample = peserta[0];

  const { code, name: displayName } = parseJabatan(jabatan);
  const displayLokasi = lokasi&&lokasi!=='—' ? lokasi : (sample.instansi||'Pemerintah Kota Semarang');

  // Percentage legend rows
  const n = peserta.length;
  const pctRows = [
    { label:'P/L – Lulus SKB',    count:plC, color:'#2563EB' },
    { label:'P – Passing Grade',  count:pC,  color:'#16A34A' },
    { label:'TL – Tidak Lulus',   count:tlC, color:'#B91C1C' },
    { label:'TH – Tidak Hadir',   count:thC, color:'#64748B' },
  ].filter(r => r.count > 0);

  const pctHTML = pctRows.map(r => `
    <div class="modal-pct-row">
      <span class="modal-pct-dot" style="background:${r.color}"></span>
      <span class="modal-pct-label">${r.label}</span>
      <span class="modal-pct-val">${r.count}</span>
      <span class="modal-pct-pct">${((r.count/n)*100).toFixed(1)}%</span>
    </div>
  `).join('');

  // Peserta list sorted by SKD desc
  const rows = [...peserta].sort((a,b) => b.skd-a.skd).map((p,i) => {
    const ridx = filteredPeserta.indexOf(p);
    const ki   = getKetInfo(p.ket);
    const clickable = ridx !== -1;
    return `<tr>
      <td>${i+1}</td>
      <td>${clickable
        ? `<span class="link-nama" onclick="closeJabatanModal();showDetailPage(${ridx})">${escapeHTML(p.nama)}</span>`
        : escapeHTML(p.nama)}</td>
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

  document.getElementById('modalJabatanContent').innerHTML = `
    <div class="modal-jabatan-code">${escapeHTML(code || jabatan.match(/^([A-Z0-9]+-\d*)/)?.[0] || '')}</div>
    <div class="modal-jabatan-title">${escapeHTML(displayName || jabatan)}</div>
    <div class="modal-jabatan-meta">
      <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" style="vertical-align:middle;margin-right:4px"><path d="M21 10c0 7-9 13-9 13S3 17 3 10a9 9 0 0118 0z"/><circle cx="12" cy="10" r="3"/></svg>
      ${escapeHTML(displayLokasi)}
      ${syaratText ? ` &nbsp;·&nbsp; Kualifikasi: ${escapeHTML(truncate(syaratText, 100))}` : ''}
    </div>

    <div class="modal-stats-row">
      <div class="modal-stat-box"><div class="val pl">${plC}</div><div class="lbl">P/L</div></div>
      <div class="modal-stat-box"><div class="val lulus">${pC}</div><div class="lbl">P</div></div>
      <div class="modal-stat-box"><div class="val tl">${tlC}</div><div class="lbl">TL</div></div>
      <div class="modal-stat-box"><div class="val th">${thC}</div><div class="lbl">TH</div></div>
      <div class="modal-stat-box"><div class="val rata">${avgS}</div><div class="lbl">Rata SKD</div></div>
      <div class="modal-stat-box"><div class="val" style="color:var(--text-primary)">${n}</div><div class="lbl">Total</div></div>
    </div>

    <!-- Diagram persentase + legend -->
    <h4 class="modal-section-title">Diagram Persentase Kelulusan</h4>
    <div class="modal-chart-area">
      <div class="modal-chart-wrap modal-chart-donut">
        <canvas id="modalJabatanChart"></canvas>
      </div>
      <div class="modal-pct-legend">
        ${pctHTML}
      </div>
    </div>

    <!-- Daftar Peserta -->
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

  // Render doughnut chart for percentage breakdown
  setTimeout(() => {
    const ctx = document.getElementById('modalJabatanChart');
    if (!ctx) return;

    const chartLabels = pctRows.map(r => r.label.split(' – ')[0]); // short label e.g. "P/L"
    const chartData   = pctRows.map(r => r.count);
    const chartColors = pctRows.map(r => r.color);

    new Chart(ctx.getContext('2d'), {
      type: 'doughnut',
      data: {
        labels: chartLabels,
        datasets: [{
          data: chartData,
          backgroundColor: chartColors,
          borderWidth: 3,
          borderColor: '#fff',
          hoverOffset: 10,
        }],
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        cutout: '62%',
        plugins: {
          legend: { display: false },
          tooltip: {
            callbacks: {
              label: ctx => ` ${ctx.label}: ${ctx.raw} peserta (${((ctx.raw/n)*100).toFixed(1)}%)`
            }
          },
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
// KEYBOARD
// ==========================================
document.addEventListener('keydown', e => {
  if (e.key === 'Escape') { closeJabatanModal(); }
});