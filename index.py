<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8"/>
<meta name="viewport" content="width=device-width,initial-scale=1.0"/>
<title>EduAnalytics Pro</title>
<link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.3/dist/css/bootstrap.min.css" rel="stylesheet"/>
<link href="https://cdn.jsdelivr.net/npm/bootstrap-icons@1.11.3/font/bootstrap-icons.min.css" rel="stylesheet"/>
<link href="https://fonts.googleapis.com/css2?family=DM+Sans:wght@300;400;500;600;700&family=Syne:wght@700;800&display=swap" rel="stylesheet"/>
<script src="https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js"></script>
<style>
:root{
  --blue-900:#0a1628;--blue-800:#0f2044;--blue-700:#1a3a6b;
  --blue-600:#1d4ed8;--blue-500:#2563eb;--blue-400:#3b82f6;
  --blue-100:#dbeafe;--blue-50:#eff6ff;
  --accent:#0ea5e9;--success:#10b981;--warning:#f59e0b;--danger:#ef4444;
  --gray-50:#f8fafc;--gray-100:#f1f5f9;--gray-200:#e2e8f0;
  --gray-300:#cbd5e1;--gray-400:#94a3b8;--gray-500:#64748b;
  --gray-600:#475569;--gray-700:#334155;--gray-800:#1e293b;
  --shadow-sm:0 1px 3px rgba(0,0,0,.08);
  --shadow-md:0 4px 16px rgba(0,0,0,.08),0 2px 6px rgba(0,0,0,.04);
  --shadow-lg:0 10px 40px rgba(0,0,0,.10);
}
*,*::before,*::after{box-sizing:border-box;margin:0;padding:0}
body{font-family:'DM Sans',sans-serif;background:var(--gray-50);color:var(--gray-800);min-height:100vh}

/* NAV */
.top-nav{position:sticky;top:0;z-index:1000;background:#fff;border-bottom:1px solid var(--gray-200);box-shadow:var(--shadow-sm);height:64px;display:flex;align-items:center;padding:0 18px;gap:12px;overflow-x:auto}
.brand{display:flex;align-items:center;gap:9px;text-decoration:none;flex-shrink:0}
.brand-icon{width:34px;height:34px;background:linear-gradient(135deg,var(--blue-600),var(--accent));border-radius:9px;display:flex;align-items:center;justify-content:center;color:#fff;font-size:16px}
.brand-name{font-family:'Syne',sans-serif;font-weight:800;font-size:16px;color:var(--blue-900);white-space:nowrap}
.brand-name span{color:var(--blue-500)}
.nav-div{width:1px;height:26px;background:var(--gray-200);flex-shrink:0}
.nav-tabs-c{display:flex;gap:2px;flex-shrink:0}
.ntb{display:flex;align-items:center;gap:6px;padding:7px 13px;border:none;background:transparent;border-radius:8px;font-family:'DM Sans',sans-serif;font-weight:500;font-size:13px;color:var(--gray-500);cursor:pointer;transition:all .2s;white-space:nowrap}
.ntb:hover{background:var(--gray-100);color:var(--gray-800)}
.ntb.active{background:var(--blue-50);color:var(--blue-600);font-weight:600}
.nav-right{margin-left:auto;display:flex;align-items:center;gap:10px;flex-shrink:0}
.rfbtn{display:flex;align-items:center;gap:6px;padding:7px 13px;border:1.5px solid var(--blue-300);background:var(--blue-50);color:var(--blue-600);border-radius:8px;font-size:12px;font-weight:600;cursor:pointer;transition:all .2s;white-space:nowrap}
.rfbtn:hover{background:var(--blue-100)}
.rfbtn.spinning i{animation:spin .7s linear infinite}
.lbl{font-size:11px;color:var(--gray-400);white-space:nowrap}
@keyframes spin{to{transform:rotate(360deg)}}

/* PAGES */
.page{display:none}.page.active{display:block}

/* PAGE HEADER */
.ph{background:linear-gradient(120deg,var(--blue-800),var(--blue-700) 60%,var(--accent));padding:22px 26px;color:#fff}
.ph h1{font-family:'Syne',sans-serif;font-size:21px;font-weight:800}
.ph p{font-size:12px;opacity:.75;margin-top:2px}
.stat-pill{display:inline-flex;align-items:center;gap:5px;background:rgba(255,255,255,.15);border:1px solid rgba(255,255,255,.25);border-radius:99px;padding:4px 12px;font-size:12px;font-weight:600;margin-top:9px}

/* FILTERS */
.fbar{background:#fff;border-bottom:1px solid var(--gray-200);padding:11px 18px;display:flex;flex-wrap:wrap;gap:8px;align-items:center;position:sticky;top:64px;z-index:900;box-shadow:var(--shadow-sm)}
.sw{position:relative;flex:1;min-width:170px;max-width:260px}
.sw i{position:absolute;left:9px;top:50%;transform:translateY(-50%);color:var(--gray-400);font-size:12px}
.sw input{width:100%;padding:7px 9px 7px 28px;border:1.5px solid var(--gray-200);border-radius:8px;font-family:'DM Sans',sans-serif;font-size:12px;color:var(--gray-700);background:var(--gray-50);outline:none;transition:border-color .2s}
.sw input:focus{border-color:var(--blue-400);background:#fff}
.fsel{padding:7px 26px 7px 9px;border:1.5px solid var(--gray-200);border-radius:8px;font-family:'DM Sans',sans-serif;font-size:12px;color:var(--gray-700);background:var(--gray-50) url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='10' height='10' viewBox='0 0 10 10'%3E%3Cpath fill='%2394a3b8' d='M5 7L1 2h8z'/%3E%3C/svg%3E") no-repeat right 7px center;cursor:pointer;outline:none;appearance:none;min-width:130px;transition:border-color .2s}
.fsel:focus{border-color:var(--blue-400)}
.clrbtn{padding:6px 11px;border:1.5px solid var(--gray-200);background:#fff;border-radius:8px;font-size:12px;font-weight:500;color:var(--gray-500);cursor:pointer;transition:all .2s;white-space:nowrap}
.clrbtn:hover{border-color:var(--danger);color:var(--danger);background:#fef2f2}

/* TABLE */
.tsec{padding:16px 18px}
.tcard{background:#fff;border-radius:14px;box-shadow:var(--shadow-md);overflow:hidden;border:1px solid var(--gray-200)}
.tcard-hdr{padding:11px 16px;border-bottom:1px solid var(--gray-100);display:flex;align-items:center;justify-content:space-between;background:var(--gray-50)}
.tcard-hdr .ttl{font-weight:600;font-size:13px;color:var(--gray-700)}
.sl{font-size:11px;color:var(--gray-400)}
.tw{overflow-x:auto}
table.dt{width:100%;border-collapse:collapse;font-size:12px}
table.dt thead tr{background:var(--gray-50);border-bottom:2px solid var(--gray-200)}
table.dt thead th{padding:9px 12px;font-weight:600;color:var(--gray-600);white-space:nowrap;cursor:pointer;user-select:none;transition:background .15s;text-align:left}
table.dt thead th:hover{background:var(--blue-50);color:var(--blue-600)}
table.dt thead th.sa{color:var(--blue-600)}
.si{font-size:9px;margin-left:3px;opacity:.4}
table.dt thead th.sa .si{opacity:1}
table.dt tbody tr{border-bottom:1px solid var(--gray-100);transition:background .12s}
table.dt tbody tr:hover{background:var(--blue-50)}
table.dt tbody td{padding:8px 12px;color:var(--gray-700);white-space:nowrap;max-width:200px;overflow:hidden;text-overflow:ellipsis}
.lst{padding:52px 20px;text-align:center;color:var(--gray-400)}
.spr{width:32px;height:32px;border:3px solid var(--gray-200);border-top-color:var(--blue-500);border-radius:50%;animation:spin .7s linear infinite;margin:0 auto 10px}
.empty{padding:52px 20px;text-align:center;color:var(--gray-400)}
.empty i{font-size:36px;margin-bottom:8px;display:block}

/* PAGINATION */
.pbar{padding:9px 16px;border-top:1px solid var(--gray-100);display:flex;align-items:center;justify-content:space-between;flex-wrap:wrap;gap:6px;background:var(--gray-50)}
.psz{display:flex;align-items:center;gap:5px;font-size:11px;color:var(--gray-500)}
.psz select{padding:3px 6px;border:1px solid var(--gray-200);border-radius:6px;font-size:11px;color:var(--gray-700);outline:none}
.pbtns{display:flex;gap:3px}
.pgb{width:28px;height:28px;display:flex;align-items:center;justify-content:center;border:1.5px solid var(--gray-200);border-radius:6px;background:#fff;font-size:11px;color:var(--gray-600);cursor:pointer;transition:all .15s}
.pgb:hover:not(:disabled){border-color:var(--blue-400);color:var(--blue-600);background:var(--blue-50)}
.pgb.active{border-color:var(--blue-500);background:var(--blue-500);color:#fff;font-weight:600}
.pgb:disabled{opacity:.3;cursor:default}

/* EXAM LAYOUT */
.exlay{display:flex;min-height:calc(100vh - 64px)}
.exsb{width:270px;background:#fff;border-right:1px solid var(--gray-200);padding:16px 12px;position:sticky;top:64px;height:calc(100vh - 64px);overflow-y:auto;flex-shrink:0}
.sblbl{font-size:10px;font-weight:700;letter-spacing:1px;text-transform:uppercase;color:var(--gray-400);margin-bottom:5px}
.rswrap{position:relative}
.rswrap i{position:absolute;left:9px;top:50%;transform:translateY(-50%);color:var(--gray-400);font-size:12px;pointer-events:none}
.rsinput{width:100%;padding:7px 9px 7px 28px;border:1.5px solid var(--gray-200);border-radius:8px;font-family:'DM Sans',sans-serif;font-size:12px;outline:none;transition:border-color .2s}
.rsinput:focus{border-color:var(--blue-400)}
.rdrop{max-height:calc(100vh - 220px);overflow-y:auto;border:1.5px solid var(--gray-200);border-radius:8px;background:#fff;margin-top:5px}
.ritem{padding:7px 11px;font-size:12px;cursor:pointer;border-bottom:1px solid var(--gray-100);transition:background .12s;color:var(--gray-700)}
.ritem:last-child{border-bottom:none}
.ritem:hover{background:var(--blue-50);color:var(--blue-600)}
.ritem.selected{background:var(--blue-600);color:#fff;font-weight:600}
.exmain{flex:1;padding:20px;overflow-x:hidden}

/* SUMMARY */
.sumcard{background:linear-gradient(120deg,var(--blue-800),var(--blue-700) 50%,var(--accent));border-radius:14px;padding:20px;color:#fff;margin-bottom:16px;box-shadow:var(--shadow-lg);display:none}
.sumcard.vis{display:block}
.snm{font-family:'Syne',sans-serif;font-size:19px;font-weight:800;margin-bottom:2px}
.srid{font-size:11px;opacity:.7}
.smeta{display:flex;flex-wrap:wrap;gap:6px;margin-top:8px}
.mchip{background:rgba(255,255,255,.15);border:1px solid rgba(255,255,255,.25);border-radius:99px;padding:3px 10px;font-size:11px;font-weight:500}

/* KPI */
.kgrid{display:grid;grid-template-columns:repeat(auto-fill,minmax(145px,1fr));gap:10px;margin-bottom:18px}
.kcard{background:#fff;border-radius:12px;padding:13px;box-shadow:var(--shadow-md);border:1px solid var(--gray-200);border-top:3px solid var(--blue-400);transition:transform .2s}
.kcard:hover{transform:translateY(-2px)}
.kcard:nth-child(2){border-top-color:var(--accent)}
.kcard:nth-child(3){border-top-color:var(--success)}
.kcard:nth-child(4){border-top-color:var(--warning)}
.kcard:nth-child(5){border-top-color:#a78bfa}
.kcard:nth-child(6){border-top-color:#f472b6}
.kcard:nth-child(7){border-top-color:var(--accent)}
.kcard:nth-child(8){border-top-color:var(--danger)}
.kcard:nth-child(9){border-top-color:#34d399}
.kcard:nth-child(10){border-top-color:#fb923c}
.kcard:nth-child(11){border-top-color:#818cf8}
.kcard:nth-child(12){border-top-color:#e879f9}
.ktit{font-size:10px;font-weight:600;color:var(--gray-400);text-transform:uppercase;letter-spacing:.4px;margin-bottom:6px}
.kval{font-family:'Syne',sans-serif;font-size:19px;font-weight:700;color:var(--gray-800);line-height:1}

/* SECTION TITLE */
.stit{font-family:'Syne',sans-serif;font-size:15px;font-weight:700;color:var(--gray-800);display:flex;align-items:center;gap:6px}
.stit i{color:var(--blue-500)}

/* EXPORT */
.expbtn{display:flex;align-items:center;gap:5px;padding:6px 12px;border-radius:8px;font-size:11px;font-weight:600;cursor:pointer;border:1.5px solid;transition:all .15s;white-space:nowrap}
.expbtn.csv{border-color:var(--success);color:var(--success);background:#f0fdf4}
.expbtn.csv:hover{background:var(--success);color:#fff}
.expbtn.excel{border-color:#16a34a;color:#16a34a;background:#f0fdf4}
.expbtn.excel:hover{background:#16a34a;color:#fff}

/* STATUS BADGES */
.bp{background:#d1fae5;color:#065f46;padding:2px 8px;border-radius:99px;font-size:10px;font-weight:700}
.bf{background:#fee2e2;color:#991b1b;padding:2px 8px;border-radius:99px;font-size:10px;font-weight:700}
.bo{background:#fef3c7;color:#92400e;padding:2px 8px;border-radius:99px;font-size:10px;font-weight:700}

/* ATTENDANCE PROGRESS */
.attn-bar{display:inline-flex;align-items:center;gap:6px;width:100%}
.attn-track{flex:1;height:6px;background:var(--gray-200);border-radius:99px;min-width:60px}
.attn-fill{height:100%;border-radius:99px;transition:width .3s}

/* PLACEHOLDER */
.pholder{display:flex;flex-direction:column;align-items:center;justify-content:center;min-height:240px;color:var(--gray-300);text-align:center}
.pholder i{font-size:48px;margin-bottom:12px}
.pholder p{font-size:12px;color:var(--gray-400)}

/* TOAST */
.toast-c{position:fixed;bottom:16px;right:16px;z-index:9999;display:flex;flex-direction:column;gap:6px}
.tmsg{background:var(--gray-800);color:#fff;padding:10px 14px;border-radius:8px;font-size:12px;box-shadow:var(--shadow-lg);animation:slideIn .3s ease;display:flex;align-items:center;gap:6px;max-width:300px}
.tmsg.success{background:var(--success)}
.tmsg.error{background:var(--danger)}
.tmsg.info{background:var(--blue-600)}
@keyframes slideIn{from{transform:translateX(30px);opacity:0}to{transform:translateX(0);opacity:1}}

/* ATT COLOR */
.att-h{color:#10b981;font-weight:600}
.att-m{color:#f59e0b;font-weight:600}
.att-l{color:#ef4444;font-weight:600}
.att-na{color:var(--gray-400)}

@media(max-width:768px){
  .exlay{flex-direction:column}
  .exsb{width:100%;height:auto;position:static;border-right:none;border-bottom:1px solid var(--gray-200)}
  .ntb span{display:none}
  .kgrid{grid-template-columns:repeat(2,1fr)}
}
</style>
</head>
<body>

<!-- NAV -->
<nav class="top-nav">
  <a class="brand" href="#">
    <div class="brand-icon"><i class="bi bi-mortarboard-fill"></i></div>
    <span class="brand-name">Edu<span>Analytics</span></span>
  </a>
  <div class="nav-div"></div>
  <div class="nav-tabs-c">
    <button class="ntb active" id="tab-raw" onclick="switchTab('raw')"><i class="bi bi-people-fill"></i><span>Student List</span></button>
    <button class="ntb" id="tab-neet" onclick="switchTab('neet')"><i class="bi bi-capsule"></i><span>NEET</span></button>
    <button class="ntb" id="tab-jee" onclick="switchTab('jee')"><i class="bi bi-calculator-fill"></i><span>JEE</span></button>
    <button class="ntb" id="tab-foundation" onclick="switchTab('foundation')"><i class="bi bi-book-fill"></i><span>Foundation</span></button>
  </div>
  <div class="nav-right">
    <span class="lbl" id="lbl">–</span>
    <button class="rfbtn" id="rfbtn" onclick="loadAll()"><i class="bi bi-arrow-clockwise"></i>Refresh</button>
  </div>
</nav>

<!-- ====== STUDENT LIST ====== -->
<div class="page active" id="page-raw">
  <div class="ph">
    <h1><i class="bi bi-people-fill me-2"></i>Student List</h1>
    <p>All enrolled students — live from Google Sheets</p>
    <div class="stat-pill"><i class="bi bi-person-check-fill"></i><span id="rawTotal">Loading…</span></div>
  </div>
  <div class="fbar">
    <div class="sw"><i class="bi bi-search"></i><input type="text" id="rawSearch" placeholder="Search students…" oninput="rawApply()"/></div>
    <select class="fsel" id="rawF1" onchange="rawApply()"><option value="">All Centers</option></select>
    <select class="fsel" id="rawF2" onchange="rawApply()"><option value="">All Courses</option></select>
    <select class="fsel" id="rawF3" onchange="rawApply()"><option value="">All Batches</option></select>
    <select class="fsel" id="rawF4" onchange="rawApply()"><option value="">All Streams</option></select>
    <button class="clrbtn" onclick="rawClear()"><i class="bi bi-x-circle me-1"></i>Clear</button>
  </div>
  <div class="tsec">
    <div class="tcard">
      <div class="tcard-hdr"><span class="ttl"><i class="bi bi-table me-1"></i>Students</span><span class="sl" id="rawSl">–</span></div>
      <div class="tw">
        <div class="lst" id="rawLd"><div class="spr"></div><p>Fetching from Google Sheets…</p></div>
        <table class="dt" id="rawTbl" style="display:none"><thead id="rawTH"></thead><tbody id="rawTB"></tbody></table>
        <div class="empty" id="rawEm" style="display:none"><i class="bi bi-inbox"></i><p>No students match.</p></div>
      </div>
      <div class="pbar">
        <div class="psz">Rows:<select id="rawPS" onchange="rawSz()"><option value="10">10</option><option value="25" selected>25</option><option value="50">50</option><option value="100">100</option></select></div>
        <div class="pbtns" id="rawPB"></div>
      </div>
    </div>
  </div>
</div>

<!-- ====== NEET ====== -->
<div class="page" id="page-neet">
  <div class="exlay">
    <aside class="exsb">
      <div class="sblbl">Registration No</div>
      <div class="rswrap"><i class="bi bi-search"></i><input type="text" class="rsinput" id="neetRI" placeholder="Search Reg No…" oninput="filterReg('neet')" autocomplete="off"/></div>
      <div class="rdrop" id="neetRD"></div>
    </aside>
    <div class="exmain">
      <div class="sumcard" id="neetSum">
        <div class="snm" id="neetNm">–</div>
        <div class="srid" id="neetRg">–</div>
        <div class="smeta" id="neetMt"></div>
      </div>
      <div class="kgrid" id="neetKPI"></div>
      <div class="pholder" id="neetPH"><i class="bi bi-arrow-left-circle"></i><p>Select a Registration Number</p></div>
      <div id="neetPS" style="display:none">
        <div style="display:flex;align-items:center;justify-content:space-between;flex-wrap:wrap;gap:8px;margin-bottom:12px">
          <div class="stit"><i class="bi bi-clipboard-data-fill"></i>Test Performance</div>
          <div style="display:flex;gap:6px;flex-wrap:wrap;align-items:center">
            <div class="sw" style="max-width:190px"><i class="bi bi-search"></i><input type="text" id="neetTS" placeholder="Search tests…" oninput="applyPF('neet')"/></div>
            <button class="expbtn csv" onclick="exportD('neet','csv')"><i class="bi bi-filetype-csv"></i>CSV</button>
            <button class="expbtn excel" onclick="exportD('neet','xlsx')"><i class="bi bi-file-earmark-excel"></i>Excel</button>
          </div>
        </div>
        <div class="tcard">
          <div class="tw"><table class="dt" id="neetPT"><thead id="neetPH2"></thead><tbody id="neetPB"></tbody></table><div class="empty" id="neetPE" style="display:none"><i class="bi bi-inbox"></i><p>No records.</p></div></div>
          <div class="pbar"><div class="psz">Rows:<select id="neetSZ" onchange="pfSz('neet')"><option value="10" selected>10</option><option value="25">25</option><option value="50">50</option></select></div><div class="pbtns" id="neetPPB"></div></div>
        </div>
      </div>
    </div>
  </div>
</div>

<!-- ====== JEE ====== -->
<div class="page" id="page-jee">
  <div class="exlay">
    <aside class="exsb">
      <div class="sblbl">Registration No</div>
      <div class="rswrap"><i class="bi bi-search"></i><input type="text" class="rsinput" id="jeeRI" placeholder="Search Reg No…" oninput="filterReg('jee')" autocomplete="off"/></div>
      <div class="rdrop" id="jeeRD"></div>
    </aside>
    <div class="exmain">
      <div class="sumcard" id="jeeSum">
        <div class="snm" id="jeeNm">–</div>
        <div class="srid" id="jeeRg">–</div>
        <div class="smeta" id="jeeMt"></div>
      </div>
      <div class="kgrid" id="jeeKPI"></div>
      <div class="pholder" id="jeePH"><i class="bi bi-arrow-left-circle"></i><p>Select a Registration Number</p></div>
      <div id="jeePS" style="display:none">
        <div style="display:flex;align-items:center;justify-content:space-between;flex-wrap:wrap;gap:8px;margin-bottom:12px">
          <div class="stit"><i class="bi bi-clipboard-data-fill"></i>Test Performance</div>
          <div style="display:flex;gap:6px;flex-wrap:wrap;align-items:center">
            <div class="sw" style="max-width:190px"><i class="bi bi-search"></i><input type="text" id="jeeTS" placeholder="Search tests…" oninput="applyPF('jee')"/></div>
            <button class="expbtn csv" onclick="exportD('jee','csv')"><i class="bi bi-filetype-csv"></i>CSV</button>
            <button class="expbtn excel" onclick="exportD('jee','xlsx')"><i class="bi bi-file-earmark-excel"></i>Excel</button>
          </div>
        </div>
        <div class="tcard">
          <div class="tw"><table class="dt" id="jeePT"><thead id="jeePH2"></thead><tbody id="jeePB"></tbody></table><div class="empty" id="jeePE" style="display:none"><i class="bi bi-inbox"></i><p>No records.</p></div></div>
          <div class="pbar"><div class="psz">Rows:<select id="jeeSZ" onchange="pfSz('jee')"><option value="10" selected>10</option><option value="25">25</option><option value="50">50</option></select></div><div class="pbtns" id="jeePPB"></div></div>
        </div>
      </div>
    </div>
  </div>
</div>

<!-- ====== FOUNDATION ====== -->
<div class="page" id="page-foundation">
  <div class="exlay">
    <aside class="exsb">
      <div class="sblbl">Registration No</div>
      <div class="rswrap"><i class="bi bi-search"></i><input type="text" class="rsinput" id="foundationRI" placeholder="Search Reg No…" oninput="filterReg('foundation')" autocomplete="off"/></div>
      <div class="rdrop" id="foundationRD"></div>
    </aside>
    <div class="exmain">
      <div class="sumcard" id="foundationSum">
        <div class="snm" id="foundationNm">–</div>
        <div class="srid" id="foundationRg">–</div>
        <div class="smeta" id="foundationMt"></div>
      </div>
      <div class="kgrid" id="foundationKPI"></div>
      <div class="pholder" id="foundationPH"><i class="bi bi-arrow-left-circle"></i><p>Select a Registration Number</p></div>
      <div id="foundationPS" style="display:none">
        <div style="display:flex;align-items:center;justify-content:space-between;flex-wrap:wrap;gap:8px;margin-bottom:12px">
          <div class="stit"><i class="bi bi-clipboard-data-fill"></i>Test Performance</div>
          <div style="display:flex;gap:6px;flex-wrap:wrap;align-items:center">
            <div class="sw" style="max-width:190px"><i class="bi bi-search"></i><input type="text" id="foundationTS" placeholder="Search tests…" oninput="applyPF('foundation')"/></div>
            <button class="expbtn csv" onclick="exportD('foundation','csv')"><i class="bi bi-filetype-csv"></i>CSV</button>
            <button class="expbtn excel" onclick="exportD('foundation','xlsx')"><i class="bi bi-file-earmark-excel"></i>Excel</button>
          </div>
        </div>
        <div class="tcard">
          <div class="tw"><table class="dt" id="foundationPT"><thead id="foundationPH2"></thead><tbody id="foundationPB"></tbody></table><div class="empty" id="foundationPE" style="display:none"><i class="bi bi-inbox"></i><p>No records.</p></div></div>
          <div class="pbar"><div class="psz">Rows:<select id="foundationSZ" onchange="pfSz('foundation')"><option value="10" selected>10</option><option value="25">25</option><option value="50">50</option></select></div><div class="pbtns" id="foundationPPB"></div></div>
        </div>
      </div>
    </div>
  </div>
</div>

<div class="toast-c" id="toastC"></div>

<script>
// ================================================================
// CONFIG
// ================================================================
const API = 'https://script.google.com/macros/s/AKfycbzyo5Jfw2Et6s3Cnm1-ks43_jREQwpVX6MZPUaEfKSYURdbRiExt5VNV9RSByaDqzIM/exec';

// ================================================================
// DATA
// ================================================================
const DS = { raw:{headers:[],rows:[]}, neet:{headers:[],rows:[]}, jee:{headers:[],rows:[]}, foundation:{headers:[],rows:[]}, results:{headers:[],rows:[]} };

// Raw table state
const RS = { filtered:[], page:1, ps:25, sc:null, sd:'asc', fc:'', cc:'', bc:'', stc:'' };

// Exam states
const ES = {};
['neet','jee','foundation'].forEach(ex=>{
  ES[ex] = { sel:null, pf:[], page:1, ps:10, sc:null, sd:'asc' };
});

// ================================================================
// FETCH
// ================================================================
async function fetchSheet(tab){
  const r = await fetch(`${API}?sheet=${encodeURIComponent(tab)}`,{redirect:'follow'});
  if(!r.ok) throw new Error(`HTTP ${r.status}`);
  const j = await r.json();
  if(j.error && !j.available_sheets) throw new Error(j.error);
  const headers = (j.headers||[]).map(h=>String(h??'').trim()).filter(h=>h);
  const rows = (j.rows||[]).map(row=>{
    const o={};
    headers.forEach(h=>{o[h]=row[h]??'';});
    return o;
  }).filter(r=>Object.values(r).some(v=>String(v).trim()!==''));
  return {headers,rows};
}

async function loadAll(){
  spin(true);
  try{
    const [rd,nd,jd,fd,resd]=await Promise.allSettled([
      fetchSheet('Raw'), fetchSheet('NEET'), fetchSheet('JEE'),
      fetchSheet('Foundation'), fetchSheet('Results')
    ]);
    if(rd.status==='fulfilled'){DS.raw=rd.value; initRaw();}
    else showToast('Raw sheet: '+rd.reason?.message,'error');
    if(nd.status==='fulfilled'){DS.neet=nd.value; initReg('neet');}
    if(jd.status==='fulfilled'){DS.jee=jd.value; initReg('jee');}
    if(fd.status==='fulfilled'){DS.foundation=fd.value; initReg('foundation');}
    if(resd.status==='fulfilled') DS.results=resd.value;
    updateLbl();
    showToast('Data refreshed','success');
  }catch(e){showToast('Refresh failed: '+e.message,'error');}
  finally{spin(false);}
}

function spin(on){document.getElementById('rfbtn').classList.toggle('spinning',on);}
function updateLbl(){document.getElementById('lbl').textContent='Updated '+new Date().toLocaleTimeString([],{hour:'2-digit',minute:'2-digit'});}

// ================================================================
// RAW TABLE
// ================================================================
function initRaw(){
  const h=DS.raw.headers;
  RS.fc = h.find(x=>/center/i.test(x))||'';
  RS.cc = h.find(x=>/course|class/i.test(x))||'';
  RS.bc = h.find(x=>/batch/i.test(x))||'';
  RS.stc= h.find(x=>/stream/i.test(x))||'';
  fillSel('rawF1',RS.fc,'All Centers');
  fillSel('rawF2',RS.cc,'All Courses');
  fillSel('rawF3',RS.bc,'All Batches');
  fillSel('rawF4',RS.stc,'All Streams');
  document.getElementById('rawTotal').textContent=DS.raw.rows.length+' Total Students';
  rawApply();
}

function fillSel(id,col,ph){
  const s=document.getElementById(id);
  while(s.options.length>1) s.remove(1);
  if(!col) return;
  [...new Set(DS.raw.rows.map(r=>String(r[col]||'')).filter(Boolean))].sort()
    .forEach(v=>{const o=document.createElement('option');o.value=o.textContent=v;s.appendChild(o);});
}

function rawApply(){
  const q=(document.getElementById('rawSearch').value||'').toLowerCase();
  const f1=document.getElementById('rawF1').value;
  const f2=document.getElementById('rawF2').value;
  const f3=document.getElementById('rawF3').value;
  const f4=document.getElementById('rawF4').value;
  RS.filtered=DS.raw.rows.filter(r=>{
    if(f1&&String(r[RS.fc]||'')!==f1) return false;
    if(f2&&String(r[RS.cc]||'')!==f2) return false;
    if(f3&&String(r[RS.bc]||'')!==f3) return false;
    if(f4&&String(r[RS.stc]||'')!==f4) return false;
    if(q&&!Object.values(r).join(' ').toLowerCase().includes(q)) return false;
    return true;
  });
  if(RS.sc){
    RS.filtered.sort((a,b)=>{
      const va=String(a[RS.sc]||''),vb=String(b[RS.sc]||'');
      const n=isNaN(va)||isNaN(vb)?va.localeCompare(vb):Number(va)-Number(vb);
      return RS.sd==='asc'?n:-n;
    });
  }
  RS.page=1; renderRaw();
}

function rawClear(){
  document.getElementById('rawSearch').value='';
  ['rawF1','rawF2','rawF3','rawF4'].forEach(id=>document.getElementById(id).value='');
  rawApply();
}
function rawSz(){RS.ps=+document.getElementById('rawPS').value;RS.page=1;renderRaw();}
function rawSort(col){
  RS.sc===col?(RS.sd=RS.sd==='asc'?'desc':'asc'):(RS.sc=col,RS.sd='asc');
  rawApply();
}

// Columns to SHOW in the student list (exclude junk SQL columns)
const RAW_SHOW = ['regno','student_name','JD','center','class_course','batch','acad_year','Stream','overall_class_attn','l_15d_attn'];

function renderRaw(){
  const ld=document.getElementById('rawLd');
  const tbl=document.getElementById('rawTbl');
  const em=document.getElementById('rawEm');
  const th=document.getElementById('rawTH');
  const tb=document.getElementById('rawTB');
  ld.style.display='none';
  const {filtered,page,ps,sc,sd}=RS;
  // Use only valid display columns that exist in the sheet
  const headers=RAW_SHOW.filter(h=>DS.raw.headers.includes(h));
  if(!headers.length||!filtered.length){
    tbl.style.display='none'; em.style.display='block';
    document.getElementById('rawSl').textContent='0 results';
    document.getElementById('rawPB').innerHTML='';
    return;
  }
  em.style.display='none'; tbl.style.display='table';
  // Head
  th.innerHTML='<tr>'+headers.map(h=>{
    const act=sc===h;
    const ic=act?(sd==='asc'?'↑':'↓'):'⇅';
    return`<th class="${act?'sa':''}" onclick="rawSort('${esc(h)}')">${fmtHeader(h)} <span class="si">${ic}</span></th>`;
  }).join('')+'</tr>';
  // Body
  const start=(page-1)*ps,end=start+ps;
  tb.innerHTML=filtered.slice(start,end).map(row=>'<tr>'+headers.map(h=>{
    const v=String(row[h]??'');
    if(h==='overall_class_attn'||h==='l_15d_attn'){
      const pct=parseFloat(v);
      if(isNaN(pct)||v==='') return`<td><span class="att-na">–</span></td>`;
      const cls=pct>=75?'att-h':pct>=50?'att-m':'att-l';
      const fill=pct>=75?'#10b981':pct>=50?'#f59e0b':'#ef4444';
      return`<td><div class="attn-bar"><span class="${cls}" style="min-width:36px">${pct}%</span><div class="attn-track"><div class="attn-fill" style="width:${Math.min(pct,100)}%;background:${fill}"></div></div></div></td>`;
    }
    return`<td title="${esc(v)}">${esc(v)||'<span style="color:var(--gray-300)">–</span>'}</td>`;
  }).join('')+'</tr>').join('');
  document.getElementById('rawSl').textContent=`${start+1}–${Math.min(end,filtered.length)} of ${filtered.length}`;
  renderPag('rawPB',filtered.length,page,ps,p=>{RS.page=p;renderRaw();});
}

function fmtHeader(h){
  const map={regno:'Reg No',student_name:'Student Name',JD:'Join Date',center:'Center',class_course:'Class/Course',batch:'Batch',acad_year:'Acad Year',Stream:'Stream',overall_class_attn:'Overall Attn %',l_15d_attn:'Last 15D Attn %'};
  return map[h]||h;
}

// ================================================================
// EXAM SIDEBAR
// ================================================================
function initReg(ex){
  filterReg(ex);
}

function filterReg(ex){
  const inpId = ex==='neet'?'neetRI':ex==='jee'?'jeeRI':'foundationRI';
  const dropId = ex+'RD';
  const q=(document.getElementById(inpId)?.value||'').toLowerCase();
  const data=DS[ex];
  const drop=document.getElementById(dropId);
  if(!drop) return;
  if(!data.rows.length){drop.innerHTML='<div class="ritem" style="color:var(--gray-400)">No data loaded</div>';return;}
  const rc=data.headers.find(h=>/regno/i.test(h))||data.headers[0]||'';
  const sel=ES[ex].sel;
  const vals=[...new Set(data.rows.map(r=>String(r[rc]||'').trim()).filter(Boolean))]
    .filter(v=>!q||v.toLowerCase().includes(q)).slice(0,200);
  if(!vals.length){drop.innerHTML='<div class="ritem" style="color:var(--gray-400)">No matches</div>';return;}
  drop.innerHTML=vals.map(v=>`<div class="ritem ${v===sel?'selected':''}" onclick="selectReg('${ex}','${esc(v)}')">${esc(v)}</div>`).join('');
}

function selectReg(ex,reg){
  ES[ex].sel=reg;
  const inpId=ex==='neet'?'neetRI':ex==='jee'?'jeeRI':'foundationRI';
  document.getElementById(inpId).value=reg;
  filterReg(ex);
  renderExamStudent(ex);
}

// ================================================================
// EXAM STUDENT SUMMARY + KPI
// ================================================================
// KPI columns per exam type
const KPI_COLS = {
  neet:['Total Exam','Exam_Attended','Exam_Attn%','Avg_Marks%','avg_physics_marks','avg_chemistry_marks','avg_botany_marks','avg_zoology_marks'],
  jee: ['Total Exam','Exam_Attended','Exam_Attn%','Avg_Marks%','avg_physics_marks','avg_chemistry_marks'],
  foundation:['Total Exam','Exam_Attended','Exam_Attn%','Avg_Marks%','avg_Maths_marks','avg_biology_marks','avg_socialstudies_marks','avg_english_marks','avg_mat_marks','avg_history_marks','avg_civics_marks','avg_geography_marks','avg_science_marks']
};

const KPI_LABELS = {
  'Total Exam':'Total Exams','Exam_Attended':'Attended','Exam_Attn%':'Exam Attn %','Avg_Marks%':'Avg Marks %',
  'avg_physics_marks':'Avg Physics','avg_chemistry_marks':'Avg Chemistry','avg_botany_marks':'Avg Botany',
  'avg_zoology_marks':'Avg Zoology','avg_Maths_marks':'Avg Maths','avg_biology_marks':'Avg Biology',
  'avg_socialstudies_marks':'Avg Soc Studies','avg_english_marks':'Avg English','avg_mat_marks':'Avg MAT',
  'avg_history_marks':'Avg History','avg_civics_marks':'Avg Civics','avg_geography_marks':'Avg Geography',
  'avg_science_marks':'Avg Science'
};

function renderExamStudent(ex){
  const reg=ES[ex].sel;
  if(!reg) return;
  const data=DS[ex];
  const rc=data.headers.find(h=>/regno/i.test(h))||data.headers[0]||'';
  const student=data.rows.find(r=>String(r[rc]||'').trim()===reg);

  // Show sections
  document.getElementById(ex+'PH').style.display='none';
  document.getElementById(ex+'PS').style.display='block';
  document.getElementById(ex+'Sum').classList.add('vis');

  if(!student){
    document.getElementById(ex+'Nm').textContent=reg;
    document.getElementById(ex+'Rg').textContent='Not found in '+ex.toUpperCase()+' sheet';
    document.getElementById(ex+'Mt').innerHTML='';
    document.getElementById(ex+'KPI').innerHTML='';
    renderPerfTable(ex);
    return;
  }

  const h=data.headers;
  const nc=h.find(k=>/student_name/i.test(k))||'';
  const cc=h.find(k=>/center/i.test(k))||'';
  const coc=h.find(k=>/course|class/i.test(k))||'';
  const bc=h.find(k=>/batch/i.test(k))||'';
  const stc=h.find(k=>/stream/i.test(k))||'';

  document.getElementById(ex+'Nm').textContent=student[nc]||'(No Name)';
  document.getElementById(ex+'Rg').textContent='Reg No: '+reg;
  document.getElementById(ex+'Mt').innerHTML=[
    {icon:'bi-building',v:student[cc],l:'Center'},
    {icon:'bi-book',v:student[coc],l:'Course'},
    {icon:'bi-calendar-week',v:student[bc],l:'Batch'},
    {icon:'bi-tag',v:student[stc],l:'Stream'}
  ].filter(m=>m.v).map(m=>`<span class="mchip"><i class="bi ${m.icon} me-1"></i>${esc(String(m.v))}</span>`).join('');

  // KPI cards
  const kpiCols=KPI_COLS[ex]||[];
  document.getElementById(ex+'KPI').innerHTML=kpiCols.map(col=>{
    let val=String(student[col]??'–');
    if(val===''||val==='undefined') val='–';
    // Format percentage
    if((col==='Exam_Attn%'||col==='Avg_Marks%')&&val!=='-'&&val!=='–'&&!isNaN(val)){
      val=(parseFloat(val)*100).toFixed(1)+'%';
    }
    const lbl=KPI_LABELS[col]||col;
    return`<div class="kcard"><div class="ktit">${esc(lbl)}</div><div class="kval">${esc(val)}</div></div>`;
  }).join('');

  renderPerfTable(ex);
}

// ================================================================
// PERFORMANCE TABLE
// ================================================================
// Result columns to display
const PERF_HEADERS = ['test_date','test_name','test_pattern','paper_type','totalmarks','userscore','markspercent','AIR','physics_marks','chemistry_marks','maths_marks','botany_marks','zoology_marks','biology_marks','resultstatus'];
const PERF_LABELS  = {test_date:'Date',test_name:'Test Name',test_pattern:'Pattern',paper_type:'Paper',totalmarks:'Total',userscore:'Score',markspercent:'%',AIR:'AIR',physics_marks:'Physics',chemistry_marks:'Chemistry',maths_marks:'Maths',botany_marks:'Botany',zoology_marks:'Zoology',biology_marks:'Biology',resultstatus:'Status'};

function renderPerfTable(ex){
  const st=ES[ex];
  const reg=st.sel;
  const resData=DS.results;
  const rrc=resData.headers.find(h=>/regno/i.test(h))||resData.headers[0]||'';
  const streamCol=resData.headers.find(h=>/stream/i.test(h))||'';

  // Map exam to stream filter value
  const streamMap={neet:'NEET',jee:'JEE',foundation:'BOARD_EXAM'};
  const streamFilter=streamMap[ex];

  let rows=resData.rows.filter(r=>{
    const rMatch=String(r[rrc]||'').trim()===reg;
    if(!rMatch) return false;
    if(streamFilter&&streamCol){
      return String(r[streamCol]||'').toUpperCase()===streamFilter;
    }
    return true;
  });

  const q=(document.getElementById(ex+'TS')?.value||'').toLowerCase();
  if(q) rows=rows.filter(r=>Object.values(r).join(' ').toLowerCase().includes(q));

  if(st.sc){
    rows.sort((a,b)=>{
      const va=String(a[st.sc]||''),vb=String(b[st.sc]||'');
      const n=isNaN(va)||isNaN(vb)?va.localeCompare(vb):Number(va)-Number(vb);
      return st.sd==='asc'?n:-n;
    });
  }
  st.pf=rows;

  // Determine which perf columns to show based on stream
  let showCols=PERF_HEADERS.filter(h=>resData.headers.includes(h));
  // Filter subject columns by stream
  if(streamFilter==='NEET') showCols=showCols.filter(h=>!['maths_marks','biology_marks','socialstudies_marks','english_marks'].includes(h));
  if(streamFilter==='JEE') showCols=showCols.filter(h=>!['botany_marks','zoology_marks','biology_marks'].includes(h));

  const headEl=document.getElementById(ex+'PH2');
  if(headEl){
    headEl.innerHTML='<tr>'+showCols.map(h=>{
      const act=st.sc===h;
      const ic=act?(st.sd==='asc'?'↑':'↓'):'⇅';
      return`<th class="${act?'sa':''}" onclick="pfSort('${ex}','${esc(h)}')">${PERF_LABELS[h]||h} <span class="si">${ic}</span></th>`;
    }).join('')+'</tr>';
  }

  const body=document.getElementById(ex+'PB');
  const empty=document.getElementById(ex+'PE');
  if(!rows.length){
    if(body) body.innerHTML='';
    if(empty) empty.style.display='flex';
    document.getElementById(ex+'PPB').innerHTML='';
    return;
  }
  if(empty) empty.style.display='none';

  const start=(st.page-1)*st.ps,end=start+st.ps;
  if(body){
    body.innerHTML=rows.slice(start,end).map(row=>'<tr>'+showCols.map(h=>{
      let v=String(row[h]??'');
      if(h==='resultstatus'){
        const lv=v.toLowerCase();
        if(lv.includes('present')) return`<td><span class="bp">${esc(v)}</span></td>`;
        if(lv.includes('absent')) return`<td><span class="bf">${esc(v)}</span></td>`;
        return`<td><span class="bo">${esc(v)||'–'}</span></td>`;
      }
      if(h==='markspercent'&&v&&!isNaN(v)) v=parseFloat(v).toFixed(1)+'%';
      return`<td>${esc(v)||'<span style="color:var(--gray-300)">–</span>'}</td>`;
    }).join('')+'</tr>').join('');
  }
  renderPag(ex+'PPB',rows.length,st.page,st.ps,p=>{ES[ex].page=p;renderPerfTable(ex);});
}

function applyPF(ex){ES[ex].page=1;renderPerfTable(ex);}
function pfSz(ex){ES[ex].ps=+document.getElementById(ex+'SZ').value;ES[ex].page=1;renderPerfTable(ex);}
function pfSort(ex,col){
  const st=ES[ex];
  st.sc===col?(st.sd=st.sd==='asc'?'desc':'asc'):(st.sc=col,st.sd='asc');
  renderPerfTable(ex);
}

// ================================================================
// EXPORT
// ================================================================
function exportD(ex,fmt){
  const data=ES[ex].pf;
  if(!data.length){showToast('No data to export','error');return;}
  const reg=ES[ex].sel||ex;
  const headers=DS.results.headers;
  if(fmt==='csv'){
    const lines=[headers.join(','),...data.map(r=>headers.map(h=>`"${String(r[h]??'').replace(/"/g,'""')}"`).join(','))];
    const blob=new Blob([lines.join('\n')],{type:'text/csv'});
    const a=document.createElement('a');a.href=URL.createObjectURL(blob);a.download=`${ex}_${reg}.csv`;a.click();URL.revokeObjectURL(a.href);
    showToast('CSV downloaded!','success');
  } else {
    const ws=XLSX.utils.json_to_sheet(data);
    const wb=XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb,ws,'Results');
    XLSX.writeFile(wb,`${ex}_${reg}.xlsx`);
    showToast('Excel downloaded!','success');
  }
}

// ================================================================
// PAGINATION
// ================================================================
function renderPag(cid,total,current,size,onPage){
  const el=document.getElementById(cid);if(!el) return;
  const tp=Math.ceil(total/size);
  if(tp<=1){el.innerHTML='';return;}
  let s=Math.max(1,current-2),e=Math.min(tp,s+4);s=Math.max(1,e-4);
  const btn=(lbl,pg,cls='')=>`<button class="pgb ${cls}" ${pg===null?'disabled':''} onclick="(${onPage.toString()})(${pg})">${lbl}</button>`;
  let h=btn('‹',current>1?current-1:null);
  if(s>1) h+=btn('1',1);
  if(s>2) h+=`<span class="pgb" style="cursor:default;border:none">…</span>`;
  for(let i=s;i<=e;i++) h+=btn(i,i,i===current?'active':'');
  if(e<tp-1) h+=`<span class="pgb" style="cursor:default;border:none">…</span>`;
  if(e<tp) h+=btn(tp,tp);
  h+=btn('›',current<tp?current+1:null);
  el.innerHTML=h;
}

// ================================================================
// TAB
// ================================================================
function switchTab(tab){
  document.querySelectorAll('.page').forEach(p=>p.classList.remove('active'));
  document.querySelectorAll('.ntb').forEach(b=>b.classList.remove('active'));
  document.getElementById('page-'+tab).classList.add('active');
  document.getElementById('tab-'+tab).classList.add('active');
}

// ================================================================
// TOAST
// ================================================================
function showToast(msg,type='info'){
  const icons={success:'bi-check-circle-fill',error:'bi-exclamation-circle-fill',info:'bi-info-circle-fill'};
  const d=document.createElement('div');
  d.className=`tmsg ${type}`;
  d.innerHTML=`<i class="bi ${icons[type]||icons.info}"></i> ${esc(msg)}`;
  document.getElementById('toastC').appendChild(d);
  setTimeout(()=>d.remove(),3500);
}

// ================================================================
// UTILS
// ================================================================
function esc(s){return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');}

// ================================================================
// INIT
// ================================================================
document.addEventListener('DOMContentLoaded',()=>{
  loadAll();
  setInterval(loadAll,5*60*1000);
});
</script>
<script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.3/dist/js/bootstrap.bundle.min.js"></script>
</body>
</html>
