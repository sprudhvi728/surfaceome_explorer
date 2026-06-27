# surfaceome_explorer
An interactive web tool designed to streamline cell-surface protein identification from quantitative proteomics datasets.
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Surfaceome Explorer</title>
<script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.1/chart.umd.min.js"></script>
<style>
/* ── RESET ── */
*,*::before,*::after{box-sizing:border-box;margin:0;padding:0}
body{font-family:Arial,sans-serif;background:#f8fafc;color:#0f172a;font-size:14px;line-height:1.6;min-height:100vh}
a{color:#2563eb;text-decoration:none}a:hover{text-decoration:underline;color:#1d4ed8}

/* ── TOPBAR ── */
.topbar{position:sticky;top:0;z-index:50;height:54px;background:#0f172a;display:flex;align-items:center;gap:12px;padding:0 32px;box-shadow:0 1px 6px rgba(0,0,0,.25)}
.topbar-logo{width:32px;height:32px;background:#2563eb;border-radius:6px;display:flex;align-items:center;justify-content:center;color:#fff;font-weight:800;font-size:15px;flex-shrink:0;letter-spacing:-.01em}
.topbar-name{font-weight:700;font-size:16px;color:#f1f5f9;letter-spacing:-.01em}
.topbar-ver{font-size:11px;font-weight:700;background:#334155;color:#94a3b8;padding:2px 7px;border-radius:3px;letter-spacing:.05em}
.topbar-sub{font-size:12px;color:#64748b;margin-left:auto}

/* ── HERO ── */
#hero{background:#fff;border-bottom:1px solid #e2e8f0}
.hero-content{max-width:680px;margin:0 auto;padding:60px 28px 44px;text-align:center}
.hero-eyebrow{font-size:12px;font-weight:700;letter-spacing:.12em;text-transform:uppercase;color:#2563eb;margin-bottom:14px}
.hero-title{font-size:clamp(30px,5vw,52px);font-weight:800;letter-spacing:-.03em;line-height:1.05;margin-bottom:12px;color:#0f172a}
.hero-subtitle{font-size:17px;font-weight:500;color:#3b5380;margin-bottom:16px;line-height:1.5}
.hero-desc{font-size:13.5px;color:#64748b;line-height:1.8;max-width:560px;margin:0 auto 32px}
.hero-cta{display:inline-block;padding:12px 32px;background:#2563eb;color:#fff;font-weight:700;font-size:14px;border-radius:8px;cursor:pointer;text-decoration:none;border:none;transition:background .14s,box-shadow .14s;box-shadow:0 2px 8px rgba(37,99,235,.25)}
.hero-cta:hover{background:#1d4ed8;text-decoration:none;color:#fff;box-shadow:0 4px 14px rgba(37,99,235,.35)}
.hero-membrane{width:100%;overflow:hidden;border-top:1px solid #e8edf4}
.hero-svg{width:100%;height:auto;display:block;max-height:280px}
.hero-tagline{text-align:center;font-size:15px;font-weight:700;letter-spacing:.06em;color:#475569;padding:14px 0 18px;border-top:1px solid #e8edf4;background:#f8fafc}

/* ── TAB NAV ── */
.tab-nav{position:sticky;top:54px;z-index:40;background:#fff;border-bottom:1px solid #e2e8f0;padding:0 28px;overflow-x:auto;-webkit-overflow-scrolling:touch;box-shadow:0 1px 3px rgba(0,0,0,.05)}
.tab-list{display:flex;align-items:center;height:50px;min-width:max-content}
.tab-item{display:flex;align-items:center;gap:7px;padding:0 16px;height:50px;cursor:pointer;user-select:none;border-bottom:2px solid transparent;transition:border-color .14s,color .14s,background .12s;white-space:nowrap}
.tab-item:hover:not(.tab-locked){background:#f0f6ff}
.tab-item.tab-active{border-bottom-color:#2563eb;color:#1d4ed8}
.tab-item.tab-done{color:#475569}
.tab-item.tab-locked{color:#cbd5e1;cursor:default;pointer-events:none}
.tab-num{width:22px;height:22px;border:1.5px solid currentColor;border-radius:50%;display:flex;align-items:center;justify-content:center;font-size:12px;font-weight:700;flex-shrink:0;transition:background .14s,border-color .14s}
.tab-item.tab-active .tab-num{background:#2563eb;border-color:#2563eb;color:#fff}
.tab-item.tab-done .tab-num{background:#2563eb;border-color:#2563eb;color:#fff;font-size:13px}
.tab-lbl{font-size:13px;font-weight:700}
.tab-item.tab-done .tab-lbl{font-weight:400}
.tab-conn{flex:0 0 20px;height:1px;background:#e2e8f0;margin:0 2px}
.tab-conn-done{background:#93c5fd}

/* ── MAIN LAYOUT ── */
.main-wrap{max-width:920px;margin:0 auto;padding:36px 28px 100px}
@media(max-width:640px){.main-wrap{padding:24px 16px 80px}}

/* ── TAB HEADER ── */
.tab-hd{margin-bottom:28px}
.tab-title{font-size:22px;font-weight:700;letter-spacing:-.02em;margin-bottom:6px;color:#0f172a}
.tab-desc{font-size:13px;color:#64748b;line-height:1.75;max-width:660px}

/* ── CARD ── */
.card{background:#fff;border:1px solid #e2e8f0;border-radius:10px;padding:24px;margin-bottom:20px;box-shadow:0 1px 4px rgba(15,23,42,.05)}
.card-title{font-size:15px;font-weight:700;margin-bottom:4px;color:#0f172a}
.card-sub{font-size:13px;color:#64748b;margin-bottom:20px;line-height:1.6}
.card-meta{font-size:12px;font-weight:400;color:#94a3b8;margin-left:8px}

/* ── BUTTONS ── */
.btn{font-family:Arial,sans-serif;font-weight:700;font-size:13px;padding:8px 20px;border:1.5px solid #2563eb;background:#fff;color:#2563eb;cursor:pointer;display:inline-flex;align-items:center;gap:7px;white-space:nowrap;transition:background .12s,color .12s,box-shadow .12s;border-radius:7px;text-decoration:none}
.btn:hover{background:#2563eb;color:#fff;text-decoration:none}
.btn-primary{background:#2563eb;color:#fff;border-color:#2563eb;box-shadow:0 2px 6px rgba(37,99,235,.20)}
.btn-primary:hover{background:#1d4ed8;border-color:#1d4ed8}
.btn-ghost{background:#fff;color:#2563eb;border-color:#cbd5e1}
.btn-ghost:hover{background:#f0f6ff;color:#1d4ed8;border-color:#93c5fd}
.btn-danger{background:#dc2626;color:#fff;border-color:#dc2626}
.btn-danger:hover{background:#b91c1c;border-color:#b91c1c}
.btn-lg{font-size:14px;padding:11px 28px;border-radius:8px}
.btn-sm{font-size:12px;padding:5px 13px}
.btn-xs{font-size:12px;padding:3px 10px;border-color:#e2e8f0;color:#64748b}
.btn-xs:hover{background:#2563eb;color:#fff;border-color:#2563eb}
.btn-row{display:flex;gap:10px;flex-wrap:wrap;align-items:center}

/* ── NOTICES ── */
.notice{padding:10px 14px;border:1px solid;font-size:13px;line-height:1.5;border-radius:6px}
.notice-info{background:#eff6ff;border-color:#bfdbfe;color:#1e40af}
.notice-ok{background:#f0fdf4;border-color:#bbf7d0;color:#166534}
.notice-warn{background:#fffbeb;border-color:#fde68a;color:#92400e}
.notice-err{background:#fef2f2;border-color:#fecaca;color:#991b1b}

/* ── TOAST ── */
#toast{position:fixed;bottom:24px;left:50%;transform:translateX(-50%) translateY(80px);background:#0f172a;color:#f1f5f9;padding:9px 22px;font-size:13px;font-weight:700;z-index:999;transition:transform .2s;pointer-events:none;border-radius:8px;box-shadow:0 4px 16px rgba(0,0,0,.20);white-space:nowrap;max-width:90vw;overflow:hidden;text-overflow:ellipsis}
#toast.show{transform:translateX(-50%) translateY(0)}

/* ── DROP ZONE ── */
.drop-zone{position:relative;border:1.5px dashed #cbd5e1;border-radius:10px;padding:52px 24px;text-align:center;cursor:pointer;transition:background .12s,border-color .12s;background:#fafbff}
.drop-zone:hover,.drop-zone.drag-over{background:#eff6ff;border-color:#2563eb;border-style:solid}
.drop-icon{font-size:28px;margin-bottom:10px;color:#93c5fd}
.drop-primary{font-size:15px;font-weight:700;margin-bottom:6px;color:#1e3a5f}
.drop-secondary{font-size:13px;color:#94a3b8}
/* Sample data row */
.t1-sample-row{display:flex;align-items:center;flex-wrap:wrap;gap:10px;margin-top:14px;padding:12px 14px;background:#fafafa;border:1px dashed #e2e8f0;border-radius:8px}
.t1-sample-label{font-size:12px;color:#94a3b8;font-weight:600;text-transform:uppercase;letter-spacing:.04em}
.t1-sample-btn{font-size:12px;font-weight:600;padding:6px 14px;background:#fff;border:1.5px solid #2563eb;border-radius:6px;color:#2563eb;cursor:pointer;transition:.12s}
.t1-sample-btn:hover{background:#eff6ff}
.t1-sample-hint{font-size:11px;color:#94a3b8;margin-left:4px}

/* ── FILE BADGE ── */
.file-badge{display:flex;align-items:center;flex-wrap:wrap;gap:8px;margin-top:12px;padding:11px 16px;border:1.5px solid #bbf7d0;background:#f0fdf4;border-radius:8px}
.fbadge-icon{font-size:15px;font-weight:700;color:#16a34a}
.fbadge-name{font-weight:700;font-size:13px;color:#0f172a}
.fbadge-meta{font-size:12px;color:#64748b}
.fbadge-upid{font-size:12px;font-weight:700;padding:2px 9px;border:1px solid;border-radius:4px}
.fbadge-ok{background:#f0fdf4;border-color:#bbf7d0;color:#166534}
.fbadge-warn{background:#fffbeb;border-color:#fde68a;color:#92400e}
.fbadge-err{background:#fef2f2;border-color:#fecaca;color:#991b1b}

/* ── PREVIEW TABLE ── */
.preview-wrap{overflow-x:auto;border:1px solid #e2e8f0;border-radius:8px;margin-bottom:10px}
.ptable{width:100%;border-collapse:collapse;font-size:13px;font-family:'Courier New',monospace}
.ptable th{padding:7px 11px;text-align:left;font-weight:700;white-space:nowrap;font-family:Arial;font-size:12px;border-bottom:2px solid #e2e8f0}
.ptable td{padding:4px 11px;border-bottom:1px solid #f1f5f9;white-space:nowrap;max-width:130px;overflow:hidden;text-overflow:ellipsis;color:#334155}
.ptable tr:last-child td{border-bottom:none}
.ptable tr:nth-child(even) td{background:#fafbff}
.th-id{background:#0f172a;color:#f1f5f9}
.th-quant{background:#334155;color:#f1f5f9}
.th-other{background:#94a3b8;color:#fff}
.th-more{background:#f1f5f9;color:#94a3b8;font-weight:400;font-style:italic}
.preview-legend{display:flex;gap:14px;flex-wrap:wrap;margin-top:6px}
.leg-item{display:flex;align-items:center;gap:5px;font-size:12px;color:#64748b}
.leg-dot{width:10px;height:10px;display:inline-block;flex-shrink:0;border-radius:2px}
.leg-id{background:#0f172a}.leg-quant{background:#334155}.leg-other{background:#94a3b8}

/* ── COLUMN SUMMARY (replaces picker) ── */
.mbadge{font-size:12px;font-weight:700;padding:3px 9px;border:1px solid;border-radius:4px}
.mbadge-ok{background:#f0fdf4;border-color:#bbf7d0;color:#166534}
.mbadge-warn{background:#fffbeb;border-color:#fde68a;color:#92400e}
.mbadge-err{background:#fef2f2;border-color:#fecaca;color:#991b1b}

/* ── STUB CARDS ── */
.stub-locked-banner{display:flex;align-items:flex-start;gap:14px;padding:16px;background:#f8fafc;border:1px solid #e2e8f0;border-radius:8px;margin-bottom:20px}
.stub-lock-icon{font-size:22px;flex-shrink:0;margin-top:2px}
.stub-lock-title{font-size:14px;font-weight:700;margin-bottom:3px;color:#0f172a}
.stub-lock-title em{font-style:normal;color:#2563eb}
.stub-lock-sub{font-size:12px;color:#94a3b8}
.stub-plan-title{font-size:12px;font-weight:700;text-transform:uppercase;letter-spacing:.07em;color:#94a3b8;margin-bottom:8px}
.stub-list{list-style:none;display:flex;flex-direction:column;gap:5px}
.stub-item{display:flex;align-items:flex-start;gap:8px;font-size:13px;color:#475569}
.stub-bullet{color:#cbd5e1;flex-shrink:0;margin-top:2px}
.stub-note{font-size:12px;color:#94a3b8;margin-top:16px;padding:8px 12px;border-left:2px solid #bfdbfe;background:#eff6ff;border-radius:0 4px 4px 0}

/* ── STAT CARDS (Tab 2) ── */
.stat-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(140px,1fr));gap:12px;margin-bottom:20px}
.stat-card{background:#fff;border:1px solid #e2e8f0;border-radius:10px;padding:20px 16px;text-align:center;box-shadow:0 1px 3px rgba(15,23,42,.04)}
.stat-val{font-size:26px;font-weight:800;color:#0f172a;line-height:1.1;margin-bottom:6px}
.stat-lbl{font-size:12px;color:#64748b;font-weight:500;line-height:1.4}
.stat-card-warn .stat-val{color:#b45309}
.stat-card-err  .stat-val{color:#dc2626}

/* ── CHECK LIST (Tab 2) ── */
.chk-list{display:flex;flex-direction:column;gap:10px;margin-top:16px}
.chk-item{display:flex;align-items:flex-start;gap:12px;padding:14px 16px;border-radius:8px;border:1px solid}
.chk-item-pass{background:#f0fdf4;border-color:#bbf7d0}
.chk-item-warn{background:#fffbeb;border-color:#fde68a}
.chk-item-err {background:#fef2f2;border-color:#fecaca}
.chk-badge{font-size:12px;font-weight:700;padding:3px 10px;border-radius:4px;white-space:nowrap;flex-shrink:0;border:1px solid;margin-top:1px}
.chk-badge-pass{background:#dcfce7;color:#166534;border-color:#bbf7d0}
.chk-badge-warn{background:#fef9c3;color:#854d0e;border-color:#fde047}
.chk-badge-err {background:#fee2e2;color:#991b1b;border-color:#fca5a5}
.chk-item-body{flex:1;min-width:0}
.chk-item-title{font-size:13px;font-weight:700;margin-bottom:4px;color:#0f172a}
.chk-item-desc{font-size:13px;color:#475569;line-height:1.65}

/* ── FILTER TAB (Tab 3) ── */
.t3-sumbar{display:flex;align-items:center;justify-content:center;gap:0;background:#fff;border:1px solid #e2e8f0;border-radius:10px;padding:22px 32px;margin-bottom:20px;box-shadow:0 1px 4px rgba(15,23,42,.05)}
.t3-sum-item{text-align:center;min-width:120px}
.t3-sum-sep{font-size:22px;font-weight:300;color:#cbd5e1;margin:0 20px;user-select:none}
.t3-sum-n{display:block;font-size:32px;font-weight:800;line-height:1.1;margin-bottom:5px;color:#0f172a;transition:color .2s}
.t3-sum-lbl{display:block;font-size:12px;color:#64748b;font-weight:500}
.t3-sum-red{color:#dc2626}

.t3-fcard{margin-bottom:16px}
.t3-fhead{display:flex;align-items:flex-start;flex-wrap:wrap;gap:8px 16px;padding-bottom:14px;border-bottom:1px solid #e2e8f0;margin-bottom:18px}
.t3-en-lbl{display:flex;align-items:center;gap:8px;font-size:14px;font-weight:700;color:#0f172a;cursor:pointer;user-select:none}
.t3-en-lbl input[type=checkbox]{width:16px;height:16px;accent-color:#2563eb;cursor:pointer;flex-shrink:0}
.t3-opt-ttl{font-size:14px;font-weight:700;color:#94a3b8}
.t3-fhead-sub{font-size:12px;color:#94a3b8;align-self:center}

.t3-fbody{display:grid;grid-template-columns:1fr 1fr;gap:24px;align-items:start}
.t3-fbody-simple{display:block}
.t3-ctrl{}
.t3-ctrl-lbl{font-size:13px;font-weight:600;color:#334155;margin-bottom:8px}
.t3-radios{display:flex;flex-direction:column;gap:8px}
.t3-radio{display:flex;align-items:center;gap:8px;font-size:13px;color:#475569;cursor:pointer;user-select:none}
.t3-radio input[type=radio]{width:14px;height:14px;accent-color:#2563eb;cursor:pointer;flex-shrink:0}
/* Compact pill radios for dashboard */
.t3-pills{display:flex;flex-wrap:wrap;gap:6px;margin-bottom:10px}
.t3-pill{display:inline-flex;align-items:center;gap:4px;font-size:12px;font-weight:600;color:#475569;cursor:pointer;padding:4px 10px;border:1.5px solid #cbd5e1;border-radius:20px;background:#f8fafc;user-select:none;transition:all .15s}
.t3-pill:hover{border-color:#2563eb;color:#2563eb;background:#eff6ff}
.t3-pill-sel{border-color:#2563eb;color:#fff;background:#2563eb}
.t3-pill input[type=radio]{display:none}
/* 2-column dashboard grid */
.t3-dash-grid{display:grid;grid-template-columns:1fr 1fr;gap:16px;margin-bottom:16px}
@media(max-width:860px){.t3-dash-grid{grid-template-columns:1fr}}
.t3-dash-card{margin:0!important}
.t3-dash-ctrl{padding:10px 0 6px;border-bottom:1px solid #f1f5f9;margin-bottom:8px}
/* Info grid inside dash card */
.t3-info-grid{display:grid;grid-template-columns:auto 1fr;gap:4px 12px;font-size:12px;margin-top:10px}
.t3-ig-lbl{color:#64748b;font-weight:600}
.t3-ig-val{color:#0f172a}
.t3-sl-row{display:flex;align-items:center;gap:12px;margin-top:4px}
.t3-sl{flex:1;accent-color:#2563eb;cursor:pointer;height:4px}
.t3-sl-val{font-size:13px;font-weight:700;color:#2563eb;min-width:44px;text-align:right;white-space:nowrap}

.t3-chart-box{position:relative;height:200px;max-height:200px;overflow:hidden}
.t3-chart-key{display:flex;gap:16px;justify-content:center;margin-top:6px;font-size:12px}
.t3-key-rem{color:#dc2626;font-weight:600}
.t3-key-kpt{color:#2563eb;font-weight:600}

.t3-disabled{font-size:13px;color:#94a3b8;padding:10px 0}
.t3-opt-card{opacity:0.82}

.t3-rpt-lines{display:flex;flex-direction:column;gap:7px;margin-top:14px}
.t3-rpt-line{font-size:13px;color:#475569;line-height:1.7}
.t3-rpt-final{font-size:14px;color:#0f172a;font-weight:600;padding-top:12px;margin-top:6px;border-top:1px solid #e2e8f0}

button:disabled{opacity:0.45;cursor:not-allowed}

@media(max-width:640px){
  .t3-fbody{grid-template-columns:1fr}
  .t3-sumbar{flex-direction:column;gap:14px;padding:18px}
  .t3-sum-sep{display:none}
}

/* ── NORMALIZATION (Tab 5) ── */
.t5-3col{display:grid;grid-template-columns:repeat(3,1fr);gap:16px;margin-bottom:4px}
.t5-scard{margin-bottom:0}
.t5-sdesc{font-size:12px;color:#64748b;margin:0 0 14px;line-height:1.5}
.t5-opts{display:flex;flex-direction:column;gap:6px}
.t5-opt-row{display:flex;align-items:center;gap:9px;cursor:pointer;user-select:none;padding:8px 10px;border-radius:7px;border:1.5px solid transparent;transition:.12s}
.t5-opt-row:hover{background:#f0f9ff;border-color:#bfdbfe}
.t5-opt-row input[type=radio]{width:14px;height:14px;accent-color:#2563eb;cursor:pointer;flex-shrink:0;margin:0}
.t5-opt-row input:checked ~ .t5-opt-txt{font-weight:700;color:#1d4ed8}
.t5-opt-row:has(input:checked){background:#eff6ff;border-color:#93c5fd}
.t5-opt-txt{font-size:13px;color:#334155;line-height:1.3}
.t5-tag{font-size:10px;font-weight:700;background:#2563eb;color:#fff;padding:1px 6px;border-radius:20px;margin-left:4px;vertical-align:middle}
.t5-pseudo-note{font-size:12px;color:#b45309;background:#fffbeb;border:1px solid #fde68a;border-radius:6px;padding:6px 10px;margin-top:10px}
.t5-bp-grid{display:grid;grid-template-columns:1fr 1fr;gap:20px}
.t5-bp-wrap{overflow-x:auto;height:250px;max-height:250px;overflow:hidden}
.t5-bdg{font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:.04em;padding:2px 8px;border-radius:20px}
.t5-bdg-b{background:#f1f5f9;color:#64748b}
.t5-bdg-a{background:#eff6ff;color:#2563eb}
.t5-trunc-note{font-size:11px;font-weight:400;color:#94a3b8}
.t5-sumtbl{overflow-x:auto;margin-top:4px;max-height:320px;overflow-y:auto}
.t5-sumtbl table{border-collapse:collapse;font-size:12px;white-space:nowrap;width:100%}
.t5-sumtbl th,.t5-sumtbl td{padding:6px 12px;border:1px solid #e2e8f0;text-align:right}
.t5-sumtbl td:first-child{text-align:left;position:sticky;left:0;background:#fff;font-weight:600;color:#334155;z-index:1}
.t5-sumtbl thead th{background:#f8fafc;font-weight:700;color:#475569;position:sticky;top:0;z-index:2}
.t5-sumtbl thead th:first-child{z-index:3}
.t5-ghd{border-bottom:2px solid #e2e8f0;text-align:center}
.t5-ghd-b{background:#f8fafc;color:#64748b}
.t5-ghd-a{background:#eff6ff;color:#2563eb}
.t5-rpt-grid{display:grid;grid-template-columns:auto 1fr;gap:6px 16px;margin-top:10px;font-size:13px}
.t5-rpt-k{font-weight:600;color:#334155;white-space:nowrap}
.t5-rpt-v{color:#475569}
@media(max-width:700px){.t5-3col{grid-template-columns:1fr}.t5-bp-grid{grid-template-columns:1fr}}

/* ── MISSING VALUE ESTIMATION (Tab 4) ── */
.t4-viz-row{display:grid;grid-template-columns:1fr 1fr;gap:24px}
.t4-viz-cell{}
.t4-chart-box{position:relative;height:190px}
.t4-heatmap-wrap{overflow-x:auto;margin-top:8px}
.t4-hm-note{font-size:11px;color:#94a3b8;font-weight:400;margin-left:6px}
.t4-hm-legend{display:flex;align-items:center;gap:8px;margin-top:8px;font-size:12px;font-weight:600}
.t4-hm-miss{color:#ef4444}.t4-hm-lo{color:#bfdbfe}.t4-hm-hi{color:#1e3a8a}
.t4-cat-table td.num,.t4-cat-table th.num{text-align:right}
.t4-methods{display:grid;grid-template-columns:repeat(2,1fr);gap:12px;margin-top:6px}
.t4-mcard{display:block;padding:16px;border:2px solid #e2e8f0;border-radius:10px;cursor:pointer;transition:border-color .15s,background .15s;background:#fff}
.t4-mcard:hover{border-color:#93c5fd;background:#f0f9ff}
.t4-mcard-sel{border-color:#2563eb!important;background:#eff6ff!important}
.t4-mcard-hd{display:flex;align-items:center;gap:8px;margin-bottom:6px}
.t4-mcard-name{font-size:14px;font-weight:700;color:#0f172a}
.t4-mcard-tag{font-size:11px;font-weight:700;background:#2563eb;color:#fff;padding:2px 8px;border-radius:20px;flex-shrink:0}
.t4-mcard-desc{font-size:12px;color:#64748b;line-height:1.55}
.t4-preview-tbl td.t4-val-ok{font-family:'Courier New',monospace;font-size:12px;color:#334155}
.t4-preview-tbl td.t4-val-miss{font-size:12px}
.t4-miss-dash{color:#94a3b8;font-weight:500;margin-right:4px}
.t4-imputed{color:#16a34a;font-weight:700;font-family:'Courier New',monospace;font-size:11px;background:#f0fdf4;padding:1px 4px;border-radius:3px}
.t4-knn-est{color:#d97706;background:#fffbeb}
.t4-pid{font-family:'Courier New',monospace;font-size:12px;color:#1e40af;font-weight:600}
.t4-miss-pct{font-weight:700;color:#dc2626}
.t4-dim{color:#94a3b8;font-size:11px}
.t4-preview-note{font-size:12px;color:#94a3b8;margin-top:8px}
.t4-rpt{display:flex;flex-direction:column;gap:10px;margin-top:10px}
.t4-rpt-row{display:flex;gap:12px;align-items:baseline;flex-wrap:wrap}
.t4-rpt-lbl{font-size:13px;font-weight:600;color:#334155;min-width:170px;flex-shrink:0}
.t4-rpt-val{font-size:13px;color:#475569}
@media(max-width:640px){.t4-viz-row{grid-template-columns:1fr}.t4-methods{grid-template-columns:1fr}}

/* ── SUMMARY LOG (Tab 2) ── */
.log-block{background:#1e293b;border-radius:8px;padding:20px 24px;margin-top:16px}
.log-title{font-size:13px;font-weight:700;color:#93c5fd;letter-spacing:.04em;margin-bottom:10px;font-family:'Courier New',monospace}
.log-text{font-family:'Courier New',monospace;font-size:13px;line-height:1.9;color:#cbd5e1}

/* ── SURFACEOME FILTER (Tab 6) ── */
.t6-loading{font-size:14px;color:#64748b;padding:40px 0;text-align:center;font-style:italic}
/* 4-category summary cards */
.t6-4cards{display:grid;grid-template-columns:repeat(4,1fr);gap:14px;margin-bottom:20px}
@media(max-width:800px){.t6-4cards{grid-template-columns:repeat(2,1fr)}}
.t6-sum-card{background:#fff;border-radius:10px;border:1px solid #e2e8f0;padding:18px 16px;display:flex;flex-direction:column;gap:3px}
.t6-sc-count{font-size:28px;font-weight:800;line-height:1.1}
.t6-sc-pct{font-size:12px;color:#64748b;margin:1px 0 2px}
.t6-sc-label{font-size:12px;font-weight:600;color:#334155}
/* Top row: metrics | bar chart (2 columns) */
.t6-charts-top{display:grid;grid-template-columns:200px 1fr;gap:14px;margin-bottom:16px;align-items:start}
@media(max-width:700px){.t6-charts-top{grid-template-columns:1fr}}
.t6-metrics-panel{background:#fff;border:1px solid #e2e8f0;border-radius:10px;padding:16px}
.t6-mp-hd{font-size:12px;font-weight:700;text-transform:uppercase;letter-spacing:.05em;color:#94a3b8;margin-bottom:12px}
.t6-mp-row{display:flex;flex-direction:column;padding:7px 0;border-bottom:1px solid #f1f5f9}
.t6-mp-lbl{font-size:11px;color:#64748b}
.t6-mp-val{font-size:14px;font-weight:700;color:#0f172a;margin-top:1px}
.t6-chart-card{margin:0!important}
.t6-sec-hd{font-size:13px;font-weight:700;color:#0f172a;margin-bottom:12px;display:flex;align-items:center;flex-wrap:wrap;gap:8px}
.t6-sec-sub{font-size:11px;font-weight:400;color:#94a3b8}
.t6-bar-wrap{height:200px;max-height:200px;overflow:hidden}
/* PNG download button */
.t6-png-btn{margin-left:auto;font-size:11px;font-weight:600;padding:3px 10px;border:1px solid #e2e8f0;border-radius:5px;background:#f8fafc;color:#475569;cursor:pointer;transition:.12s;white-space:nowrap}
.t6-png-btn:hover{background:#eff6ff;border-color:#93c5fd;color:#1d4ed8}
/* Venn — full-width card */
.t6-venn-card{margin-bottom:16px}
.t6-venn-outer{display:flex;gap:24px;align-items:flex-start;flex-wrap:wrap}
.t6-venn-wrap{position:relative;flex:0 0 auto;width:100%;max-width:560px}
.t6-venn-wrap canvas{width:100%;height:auto;display:block}
/* Venn region download panel */
.t6-venn-dl{flex:1;min-width:200px}
.t6-venn-dl-hd{font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:.05em;color:#94a3b8;margin-bottom:10px}
.t6-venn-dl-grid{display:flex;flex-direction:column;gap:6px}
.t6-venn-tip{position:absolute;background:#1e293b;color:#e2e8f0;border-radius:8px;padding:10px 14px;font-size:12px;max-width:220px;z-index:100;pointer-events:none;box-shadow:0 4px 16px rgba(0,0,0,.25)}
.t6-tip-hd{font-weight:700;color:#fff;margin-bottom:2px;font-size:12px}
.t6-tip-sub{font-size:11px;color:#94a3b8;margin-bottom:6px}
.t6-tip-row{display:flex;gap:6px;align-items:baseline;padding:2px 0;border-bottom:1px solid rgba(255,255,255,.08)}
.t6-tip-pid{font-family:'Courier New',monospace;font-size:11px;color:#93c5fd;font-weight:600}
.t6-tip-gene{font-size:11px;color:#a3e635}
.t6-tip-ab{font-size:10px;color:#64748b;margin-left:auto}
.t6-tip-more{font-size:11px;color:#64748b;margin-top:5px;font-style:italic}
/* Section heading between cards */
.t6-section-hd{font-size:12px;font-weight:700;text-transform:uppercase;letter-spacing:.06em;color:#94a3b8;margin:28px 0 12px;padding-top:4px;border-top:1px solid #e2e8f0}
/* Group abundance grid */
.t6-group-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(340px,1fr));gap:16px;margin-bottom:16px}
.t6-gtbl-card{margin:0!important}
/* Tables */
.t6-tbl-wrap{overflow-x:auto;max-height:340px;overflow-y:auto}
.t6-tbl{border-collapse:collapse;font-size:12px;white-space:nowrap;width:100%}
.t6-tbl th,.t6-tbl td{padding:7px 10px;border-bottom:1px solid #f1f5f9;text-align:left;vertical-align:middle}
.t6-tbl thead th{background:#f8fafc;font-weight:700;color:#475569;position:sticky;top:0;z-index:2;font-size:11px;text-transform:uppercase;letter-spacing:.04em;border-bottom:1px solid #e2e8f0}
.t6-tbl .t6-num{text-align:right;font-variant-numeric:tabular-nums}
.t6-prow:hover{background:#f8fafc;cursor:pointer}
.t6-pid{color:#1d4ed8;font-family:'Courier New',monospace;font-size:11px;font-weight:600}
.t6-tbl-note{font-size:11px;color:#94a3b8;margin:0 0 8px;font-style:italic}
/* Search input */
.t6-search-row{margin-bottom:10px}
.t6-search{width:100%;max-width:300px;padding:6px 10px;border:1px solid #e2e8f0;border-radius:6px;font-size:12px;color:#0f172a;outline:none}
.t6-search:focus{border-color:#2563eb;box-shadow:0 0 0 2px rgba(37,99,235,.1)}
/* Diff analysis */
.t6-diff-grid{display:flex;flex-direction:column;gap:16px;margin-bottom:16px}
.t6-diff-card{margin:0!important}
.t6-up-hd{color:#dc2626}
.t6-dn-hd{color:#1d4ed8}
.t6-sig{background:#fef9f9}
.t6-fc{font-weight:700;font-variant-numeric:tabular-nums}
/* Comparison selector */
.t6-comp-sel{display:flex;align-items:center;gap:10px;margin-bottom:16px;padding:12px 16px;background:#f8fafc;border:1px solid #e2e8f0;border-radius:8px}
.t6-comp-lbl{font-size:13px;font-weight:600;color:#334155;white-space:nowrap}
.t6-comp-dd{font-size:13px;padding:5px 10px;border:1px solid #e2e8f0;border-radius:6px;background:#fff;color:#0f172a;cursor:pointer}
/* Volcano */
.t6-volcano-wrap{height:300px;max-height:300px;overflow:hidden}
/* Condition-specific */
.t6-spec-tag{font-size:11px;font-weight:700;padding:2px 8px;border-radius:20px;background:#fef3c7;color:#92400e;white-space:nowrap}
/* Downloads */
.t6-dl-row{display:flex;flex-wrap:wrap;gap:10px;margin-top:4px}
.t6-dlbtn{font-size:12px!important;padding:7px 14px!important}
/* Drawer */
.t6-overlay{position:fixed;inset:0;background:rgba(15,23,42,.3);z-index:200}
.t6-drawer{position:fixed;right:0;top:0;bottom:0;width:340px;max-width:90vw;background:#fff;z-index:201;overflow-y:auto;box-shadow:-4px 0 20px rgba(15,23,42,.12);display:flex;flex-direction:column}
.t6-dr-hd{display:flex;justify-content:space-between;align-items:center;padding:18px 20px 12px;border-bottom:1px solid #e2e8f0;position:sticky;top:0;background:#fff;z-index:1}
.t6-dr-title{font-size:14px;font-weight:800;color:#0f172a;font-family:'Courier New',monospace}
.t6-dr-close{background:none;border:none;font-size:17px;cursor:pointer;color:#64748b;padding:4px;border-radius:4px;line-height:1}
.t6-dr-close:hover{background:#f1f5f9}
.t6-dr-body{padding:14px 20px;display:flex;flex-direction:column}
.t6-dr-row{display:flex;justify-content:space-between;align-items:flex-start;padding:7px 0;border-bottom:1px solid #f8fafc;gap:8px;font-size:13px;color:#0f172a}
.t6-dr-lbl{font-weight:600;color:#64748b;flex-shrink:0;font-size:11px;text-transform:uppercase;letter-spacing:.04em;padding-top:1px}
.t6-dr-sep{height:1px;background:#e2e8f0;margin:8px 0}
.t6-dr-links{display:flex;flex-wrap:wrap;gap:8px;margin-top:12px}
.t6-extlink{color:#1d4ed8;font-weight:600;text-decoration:none;font-size:12px}
.t6-extlink:hover{text-decoration:underline}
.t6-extbtn{background:#f8fafc;border:1px solid #e2e8f0;border-radius:6px;padding:5px 11px;font-size:12px;font-weight:600;color:#334155;text-decoration:none;transition:.12s}
.t6-extbtn:hover{background:#eff6ff;border-color:#93c5fd;color:#1d4ed8}
/* DB inline tags */
.t6-db-ss{display:inline-block;font-size:11px;font-weight:700;padding:2px 7px;border-radius:20px;background:#dbeafe;color:#1e40af;margin-right:4px}
.t6-db-pm{display:inline-block;font-size:11px;font-weight:700;padding:2px 7px;border-radius:20px;background:#d1fae5;color:#065f46;margin-right:4px}
.t6-db-ma{display:inline-block;font-size:11px;font-weight:700;padding:2px 7px;border-radius:20px;background:#ede9fe;color:#4c1d95;margin-right:4px}
/* Biological interpretation summary */
.t6-bio-summary{background:#f0f9ff;border:1px solid #bae6fd;border-left:4px solid #0284c7;border-radius:8px;padding:18px 22px;margin-bottom:20px}
.t6-bio-sum-hd{font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:.07em;color:#0369a1;margin-bottom:8px}
.t6-bio-sum-body{font-size:13px;color:#0f172a;line-height:1.75}
.t6-bio-sum-body strong{color:#0369a1}
/* Section divider label */
.t6-bio-section-label{font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:.07em;color:#94a3b8;margin:28px 0 6px;padding-bottom:6px;border-bottom:2px solid #e2e8f0}
/* Database membership matrix */
.t6-matrix-tbl th.t6-ck-cell,.t6-matrix-tbl td.t6-ck-cell{text-align:center;min-width:90px}
.t6-ck-yes{color:#059669;font-weight:700;font-size:15px}
.t6-ck-no{color:#cbd5e1;font-size:14px}
.t6-db3{display:inline-block;background:#dcfce7;color:#166534;font-size:11px;font-weight:700;padding:2px 7px;border-radius:20px}
.t6-db2{display:inline-block;background:#fef9c3;color:#854d0e;font-size:11px;font-weight:700;padding:2px 7px;border-radius:20px}
.t6-db1{display:inline-block;background:#f1f5f9;color:#475569;font-size:11px;font-weight:700;padding:2px 7px;border-radius:20px}
/* Analysis Report call-to-action block */
.t6-report-cta{display:flex;gap:18px;align-items:center;background:#f0fdf4;border:1px solid #bbf7d0;border-left:4px solid #16a34a;border-radius:8px;padding:16px 18px;margin-bottom:18px;flex-wrap:wrap}
.t6-report-cta-body{flex:1;min-width:200px}
.t6-report-cta-title{font-size:13px;font-weight:700;color:#14532d;margin-bottom:4px}
.t6-report-cta-desc{font-size:12px;color:#166534;line-height:1.6}
.t6-report-btn{white-space:nowrap;font-size:12px;padding:8px 16px;background:#16a34a;border-color:#16a34a;flex-shrink:0}
.t6-report-btn:hover{background:#15803d;border-color:#15803d}
.t6-dl-subhd{font-size:11px;font-weight:600;color:#94a3b8;text-transform:uppercase;letter-spacing:.05em;margin-bottom:8px}
/* References */
.t6-refs{border-top:1px solid #e2e8f0;margin-top:36px;padding-top:20px}
.t6-refs-hd{font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:.07em;color:#94a3b8;margin-bottom:12px}
.t6-refs-list{margin:0;padding-left:20px;display:flex;flex-direction:column;gap:8px}
.t6-refs-list li{font-size:12px;color:#64748b;line-height:1.7}
.t6-refs-list em{font-style:italic}

/* ── MISC ── */
.sdiv{border:none;border-top:1px solid #e2e8f0;margin:24px 0}
.sec-title{font-size:12px;font-weight:700;text-transform:uppercase;letter-spacing:.07em;color:#94a3b8;margin-bottom:10px;padding-bottom:6px;border-bottom:1px solid #e2e8f0}
@media(max-width:600px){.cp-row{flex-direction:column;gap:10px}.v2-sel{min-width:100%}.hero-content{padding:40px 20px 32px}}
</style>
</head>
<body>

<!-- ── TOPBAR ── -->
<header class="topbar">
  <div class="topbar-logo">S</div>
  <span class="topbar-name">Surfaceome Explorer</span>
  <span class="topbar-ver">v2</span>
  <span class="topbar-sub">Cell-surface protein annotation &amp; prioritization</span>
</header>

<!-- ── HERO ── -->
<section id="hero">
  <div class="hero-content">
    <div class="hero-eyebrow">Proteomics &middot; Cell Surface Biology</div>
    <h1 class="hero-title">Surfaceome Explorer</h1>
    <p class="hero-subtitle">Prioritize Cell Surface Proteins from Quantitative Proteomics Data</p>
    <p class="hero-desc">Upload quantitative proteomics datasets, annotate proteins using curated surfaceome databases, visualize cell surface protein abundance, and export filtered target lists for downstream analysis.</p>
    <a class="hero-cta" href="#workflow"
       onclick="document.getElementById('workflow').scrollIntoView({behavior:'smooth'});return false">
      Start Analysis &darr;
    </a>
  </div>
  <div class="hero-membrane">
    <svg class="hero-svg" viewBox="0 0 1600 340" xmlns="http://www.w3.org/2000/svg"
     preserveAspectRatio="xMidYMid meet" role="img"
     aria-label="Plasma membrane cross-section showing diverse membrane protein types">
  <title>Plasma membrane with membrane proteins</title>
  <defs>
    <radialGradient id="hg" cx="34%" cy="28%" r="70%">
      <stop offset="0%" stop-color="#ffffff"/>
      <stop offset="100%" stop-color="#cccccc"/>
    </radialGradient>
    <!-- Outer leaflet: heads y=128, tails to y=185 -->
    <pattern id="lo" x="0" y="128" width="15" height="57" patternUnits="userSpaceOnUse">
      <line x1="5"  y1="15" x2="4.5" y2="55" stroke="#bebebe" stroke-width="0.9"/>
      <line x1="10" y1="15" x2="10.5" y2="55" stroke="#bebebe" stroke-width="0.9"/>
      <circle cx="7.5" cy="7.5" r="7" fill="url(#hg)" stroke="#aaaaaa" stroke-width="0.9"/>
    </pattern>
    <!-- Inner leaflet: tails y=185, heads y=242 -->
    <pattern id="li" x="0" y="185" width="15" height="57" patternUnits="userSpaceOnUse">
      <line x1="5"  y1="2"  x2="4.5" y2="42" stroke="#bebebe" stroke-width="0.9"/>
      <line x1="10" y1="2"  x2="10.5" y2="42" stroke="#bebebe" stroke-width="0.9"/>
      <circle cx="7.5" cy="49.5" r="7" fill="url(#hg)" stroke="#aaaaaa" stroke-width="0.9"/>
    </pattern>
  </defs>

  <!-- Backgrounds -->
  <rect width="1600" height="340" fill="#ffffff"/>
  <rect y="242" width="1600" height="98"  fill="#f4f6ff"/>
  <rect y="128" width="1600" height="114" fill="#f4efe6"/>

  <!-- Bilayer fill patterns -->
  <rect y="128" width="1600" height="57" fill="url(#lo)"/>
  <rect y="185" width="1600" height="57" fill="url(#li)"/>

  <!-- Boundary lines -->
  <line x1="0" y1="128" x2="1600" y2="128" stroke="#c0c0c0" stroke-width="0.8"/>
  <line x1="0" y1="242" x2="1600" y2="242" stroke="#c0c0c0" stroke-width="0.8"/>

  <!-- ================================================================
       PROTEINS  (L→R)  EC = y<128  |  TM = y 128-242  |  IC = y>242
       ================================================================ -->

  <!-- 1. Small yellow secreted blob (EC, x=72) -->
  <g transform="translate(72,0)">
    <ellipse cx="0" cy="99"  rx="22" ry="18" fill="#f8d060" opacity="0.92"/>
    <ellipse cx="-6" cy="93" rx="11" ry="9"  fill="#fce890" opacity="0.55"/>
  </g>

  <!-- 2. Salmon large IC blob (x=150, IC side) -->
  <g transform="translate(150,0)">
    <ellipse cx="4"   cy="283" rx="40" ry="30" fill="#f0a888" opacity="0.82"/>
    <ellipse cx="-12" cy="270" rx="22" ry="18" fill="#f8c8b0" opacity="0.58"/>
    <ellipse cx="18"  cy="292" rx="18" ry="12" fill="#fcd8c0" opacity="0.45"/>
  </g>

  <!-- 3. Orange floating EC blob (x=132) -->
  <g transform="translate(132,0)">
    <ellipse cx="0" cy="87"  rx="24" ry="20" fill="#f8a860" opacity="0.88"/>
    <ellipse cx="-7" cy="81" rx="12" ry="10" fill="#fcc890" opacity="0.55"/>
  </g>

  <!-- 4. Blue two-lobe TM receptor #1 (x=215) -->
  <g transform="translate(215,0)">
    <ellipse cx="0" cy="88" rx="26" ry="30" fill="#76b4e4" opacity="0.88"/>
    <ellipse cx="-7" cy="81" rx="14" ry="15" fill="#a8d0f5" opacity="0.50"/>
    <ellipse cx="0" cy="40" rx="20" ry="24" fill="#8ac4ec" opacity="0.88"/>
    <ellipse cx="-5" cy="34" rx="10" ry="12" fill="#b4dcf8" opacity="0.52"/>
    <rect x="-7" y="62"  width="14" height="28" fill="#7ab8e8" rx="5"/>
    <rect x="-7" y="118" width="14" height="126" fill="#5a9fd8" rx="6" opacity="0.88"/>
    <ellipse cx="0" cy="261" rx="14" ry="9" fill="#4a8cc8" opacity="0.72"/>
  </g>

  <!-- 5. Pink mushroom GPI #1 (x=350) -->
  <g transform="translate(350,0)">
    <ellipse cx="0" cy="103" rx="34" ry="20" fill="#f4a0b8" opacity="0.90"/>
    <ellipse cx="-10" cy="99" rx="17" ry="10" fill="#fcc8d8" opacity="0.55"/>
    <rect x="-5" y="123" width="10" height="14" fill="#e888a8" rx="4"/>
  </g>

  <!-- 6. Green flat peripheral (outer leaflet, x=422) -->
  <g transform="translate(422,0)">
    <ellipse cx="0" cy="121" rx="32" ry="13" fill="#80c878" opacity="0.88"/>
    <ellipse cx="-9" cy="118" rx="15" ry="7"  fill="#a8e0a0" opacity="0.58"/>
  </g>

  <!-- 7. Large yellow/gold EC protein (x=498) -->
  <g transform="translate(498,0)">
    <ellipse cx="0" cy="88"  rx="28" ry="32" fill="#f8d060" opacity="0.90"/>
    <ellipse cx="-9" cy="80" rx="14" ry="16" fill="#fce890" opacity="0.55"/>
  </g>

  <!-- 8. Small yellow companion (x=560) -->
  <g transform="translate(560,0)">
    <ellipse cx="0" cy="101" rx="16" ry="13" fill="#f4c840" opacity="0.88"/>
    <ellipse cx="-5" cy="97" rx="7"  ry="6"  fill="#f8e090" opacity="0.55"/>
  </g>

  <!-- 9. Pink single-pass TM (slender, x=626) -->
  <g transform="translate(626,0)">
    <ellipse cx="0" cy="78"  rx="15" ry="44" fill="#f4a0b0" opacity="0.88"/>
    <ellipse cx="-4" cy="66" rx="7"  ry="22" fill="#fcc8d4" opacity="0.50"/>
    <rect x="-6" y="122" width="12" height="122" fill="#e07898" rx="5" opacity="0.88"/>
    <ellipse cx="0" cy="261" rx="12" ry="9" fill="#c86080" opacity="0.70"/>
  </g>

  <!-- 10. Large pink mushroom GPI (x=716) -->
  <g transform="translate(716,0)">
    <ellipse cx="0" cy="96"  rx="44" ry="25" fill="#fcaac0" opacity="0.90"/>
    <ellipse cx="-13" cy="91" rx="22" ry="13" fill="#fcd4e4" opacity="0.55"/>
    <rect x="-7" y="121" width="14" height="19" fill="#f088a8" rx="5"/>
  </g>

  <!-- 11. Purple dome integral (x=788) -->
  <g transform="translate(788,0)">
    <ellipse cx="0" cy="114" rx="28" ry="19" fill="#b898d8" opacity="0.88"/>
    <ellipse cx="-8" cy="109" rx="13" ry="9"  fill="#d8c4f4" opacity="0.55"/>
    <rect x="-13" y="133" width="26" height="64" fill="#9870cc" rx="6" opacity="0.70"/>
  </g>

  <!-- 12. Teal asymmetric TM protein (x=858) -->
  <g transform="translate(858,0)">
    <ellipse cx="-5" cy="96"  rx="21" ry="24" fill="#70c8b8" opacity="0.88"/>
    <ellipse cx="9"  cy="107" rx="15" ry="17" fill="#70c8b8" opacity="0.75"/>
    <ellipse cx="-2" cy="88"  rx="10" ry="9"  fill="#a4ddd4" opacity="0.55"/>
    <rect x="-6" y="128" width="12" height="116" fill="#50a898" rx="5" opacity="0.88"/>
    <ellipse cx="0"  cy="260" rx="19" ry="14" fill="#449080" opacity="0.78"/>
    <ellipse cx="9"  cy="268" rx="10" ry="8"  fill="#68b0a0" opacity="0.50"/>
  </g>

  <!-- 13. Compact blue-ish integral (x=920) -->
  <g transform="translate(920,0)">
    <ellipse cx="0" cy="119" rx="14" ry="12" fill="#92b8e2" opacity="0.84"/>
    <rect x="-8" y="131" width="16" height="105" fill="#7098cc" rx="5" opacity="0.80"/>
  </g>

  <!-- 14. Small IC yellow blob (x=972) -->
  <g transform="translate(972,0)">
    <ellipse cx="0" cy="262" rx="18" ry="13" fill="#f8d060" opacity="0.82"/>
    <ellipse cx="-6" cy="257" rx="8"  ry="6"  fill="#fce890" opacity="0.55"/>
  </g>

  <!-- 15. Blue two-lobe TM receptor #2 (x=1034) -->
  <g transform="translate(1034,0)">
    <ellipse cx="0" cy="88" rx="26" ry="30" fill="#76b4e4" opacity="0.88"/>
    <ellipse cx="-7" cy="81" rx="14" ry="15" fill="#a8d0f5" opacity="0.50"/>
    <ellipse cx="0" cy="40" rx="20" ry="24" fill="#8ac4ec" opacity="0.88"/>
    <ellipse cx="-5" cy="34" rx="10" ry="12" fill="#b4dcf8" opacity="0.52"/>
    <rect x="-7" y="62"  width="14" height="28" fill="#7ab8e8" rx="5"/>
    <rect x="-7" y="118" width="14" height="126" fill="#5a9fd8" rx="6" opacity="0.88"/>
    <ellipse cx="0" cy="261" rx="14" ry="9" fill="#4a8cc8" opacity="0.72"/>
  </g>

  <!-- 16. Green multi-pass transporter (x=1108) -->
  <g transform="translate(1108,0)">
    <ellipse cx="0" cy="114" rx="22" ry="12" fill="#82cc80" opacity="0.85"/>
    <rect x="-19" y="126" width="38" height="118" fill="#60b060" rx="8" opacity="0.82"/>
    <line x1="-6" y1="126" x2="-6" y2="244" stroke="#9ee09c" stroke-width="2" opacity="0.50"/>
    <line x1="6"  y1="126" x2="6"  y2="244" stroke="#9ee09c" stroke-width="2" opacity="0.50"/>
    <ellipse cx="0" cy="261" rx="21" ry="13" fill="#509050" opacity="0.75"/>
  </g>

  <!-- 17. Pink mushroom GPI #2 (x=1176) -->
  <g transform="translate(1176,0)">
    <ellipse cx="0" cy="104" rx="31" ry="18" fill="#f4a0b8" opacity="0.90"/>
    <ellipse cx="-9" cy="100" rx="15" ry="9"  fill="#fcc8d8" opacity="0.55"/>
    <rect x="-5" y="122" width="10" height="14" fill="#e888a8" rx="4"/>
  </g>

  <!-- 18. Small purple EC dome (x=1238) -->
  <g transform="translate(1238,0)">
    <ellipse cx="0" cy="110" rx="21" ry="15" fill="#c0a0e0" opacity="0.88"/>
    <ellipse cx="-6" cy="106" rx="9"  ry="7"  fill="#dcc8f4" opacity="0.55"/>
  </g>

  <!-- 19. Large orange IC blob (x=1308) -->
  <g transform="translate(1308,0)">
    <ellipse cx="5"   cy="278" rx="38" ry="30" fill="#f0a870" opacity="0.84"/>
    <ellipse cx="-10" cy="264" rx="22" ry="18" fill="#f8c090" opacity="0.65"/>
    <ellipse cx="18"  cy="286" rx="18" ry="12" fill="#fcd4a8" opacity="0.48"/>
  </g>

  <!-- 20. Blue two-lobe TM receptor #3 (x=1378) -->
  <g transform="translate(1378,0)">
    <ellipse cx="0" cy="88" rx="26" ry="30" fill="#76b4e4" opacity="0.88"/>
    <ellipse cx="-7" cy="81" rx="14" ry="15" fill="#a8d0f5" opacity="0.50"/>
    <ellipse cx="0" cy="40" rx="20" ry="24" fill="#8ac4ec" opacity="0.88"/>
    <ellipse cx="-5" cy="34" rx="10" ry="12" fill="#b4dcf8" opacity="0.52"/>
    <rect x="-7" y="62"  width="14" height="28" fill="#7ab8e8" rx="5"/>
    <rect x="-7" y="118" width="14" height="126" fill="#5a9fd8" rx="6" opacity="0.88"/>
    <ellipse cx="0" cy="261" rx="14" ry="9" fill="#4a8cc8" opacity="0.72"/>
  </g>

  <!-- 21. Pink mushroom GPI #3 (x=1466) -->
  <g transform="translate(1466,0)">
    <ellipse cx="0" cy="103" rx="36" ry="21" fill="#fcaac0" opacity="0.90"/>
    <ellipse cx="-11" cy="98" rx="18" ry="11" fill="#fcd4e4" opacity="0.55"/>
    <rect x="-6" y="124" width="12" height="18" fill="#f088a8" rx="5"/>
  </g>

  <!-- 22. Red compact IC protein (x=1542) -->
  <g transform="translate(1542,0)">
    <ellipse cx="0" cy="258" rx="19" ry="14" fill="#e84840" opacity="0.88"/>
    <ellipse cx="-6" cy="252" rx="8"  ry="7"  fill="#f09090" opacity="0.55"/>
  </g>

  <!-- Small IC green blob (x=1104) -->
  <g transform="translate(1104,0)">
    <ellipse cx="0" cy="262" rx="14" ry="9" fill="#72c072" opacity="0.72"/>
  </g>

</svg>
    <div class="hero-tagline">Explore &nbsp;&middot;&nbsp; Filter &nbsp;&middot;&nbsp; Prioritize</div>
  </div>
</section>

<!-- ── WORKFLOW ── -->
<div id="workflow">
  <nav class="tab-nav" id="tab-nav"></nav>
  <main class="main-wrap">
    <div id="main-content"></div>
  </main>
</div>

<!-- ── TOAST ── -->
<div id="toast"></div>

<script>
/* ====== ANNOTATION DATABASES ====== */
const _SS_RAW='A0AV02,A0FGR9,A0PJK1,A0PK11,A0ZSE6,A1A5B4,A1A5C7,A1L157,A2RRL7,A2VDJ0,A3KFT3,A4D2G3,A4IF30,A5X5Y0,A6BM72,A6H8M9,A6NC51,A6NCV1,A6ND01,A6ND48,A6NDA9,A6NDH6,A6NDL8,A6NDP7,A6NDV4,A6NET4,A6NF34,A6NF89,A6NFA1,A6NFC5,A6NFX1,A6NGU5,A6NGY5,A6NH00,A6NH21,A6NHA9,A6NHG9,A6NHS7,A6NI73,A6NIJ9,A6NIM6,A6NJW9,A6NJZ3,A6NK97,A6NKB5,A6NKC4,A6NKK0,A6NL08,A6NL26,A6NL88,A6NM03,A6NM11,A6NM45,A6NM76,A6NMB1,A6NMS3,A6NMS7,A6NMU1,A6NMZ5,A6NND4,A6NNN8,A7MBM2,A8MPY1,A8MUP6,A8MVS5,A8MVW0,A8MVW5,A8MVZ5,A8MWY0,A8MXK1,B0FP48,B2RN74,B3SHH9,B4DS77,B6A8C7,B6SEH8,B6SEH9,B8ZZ34,D3W0D1,E2RYF6,F2Z333,F5H4A9,G3V0H7,H3BS89,I3L273,O00144,O00155,O00206,O00220,O00222,O00241,O00254,O00270,O00322,O00337,O00341,O00391,O00398,O00400,O00421,O00451,O00478,O00481,O00526,O00533,O00548,O00574,O00590,O00591,O00592,O14493,O14494,O14495,O14511,O14514,O14522,O14524,O14525,O14581,O14626,O14672,O14718,O14764,O14786,O14788,O14798,O14804,O14817,O14836,O14842,O14843,O14894,O14917,O14931,O14944,O15031,O15118,O15146,O15165,O15197,O15218,O15244,O15245,O15303,O15321,O15354,O15374,O15375,O15389,O15394,O15399,O15403,O15431,O15438,O15439,O15440,O15455,O15529,O15547,O15551,O15552,O42043,O43155,O43157,O43184,O43193,O43194,O43246,O43280,O43291,O43300,O43306,O43315,O43424,O43490,O43493,O43497,O43506,O43511,O43556,O43567,O43570,O43603,O43613,O43614,O43653,O43657,O43688,O43699,O43749,O43826,O43868,O43869,O43895,O43921,O43934,O60235,O60241,O60242,O60245,O60266,O60279,O60309,O60330,O60353,O60359,O60391,O60403,O60404,O60412,O60431,O60449,O60462,O60469,O60478,O60486,O60487,O60500,O60503,O60602,O60603,O60609,O60635,O60636,O60637,O60669,O60706,O60755,O60779,O60883,O60895,O60896,O60931,O60939,O71037,O75015,O75019,O75022,O75023,O75051,O75054,O75074,O75077,O75078,O75084,O75096,O75121,O75129,O75144,O75197,O75309,O75311,O75325,O75326,O75355,O75387,O75388,O75443,O75445,O75473,O75487,O75508,O75509,O75578,O75581,O75631,O75712,O75751,O75841,O75871,O75882,O75899,O75954,O75976,O76000,O76001,O76002,O76036,O76082,O76099,O76100,O94772,O94778,O94779,O94856,O94886,O94898,O94910,O94911,O94933,O94956,O94985,O94991,O95006,O95007,O95013,O95047,O95136,O95150,O95185,O95196,O95206,O95221,O95222,O95256,O95264,O95274,O95279,O95297,O95342,O95371,O95377,O95436,O95452,O95471,O95477,O95484,O95490,O95497,O95498,O95528,O95622,O95665,O95727,O95754,O95800,O95832,O95838,O95857,O95858,O95866,O95867,O95868,O95907,O95918,O95944,O95971,O95976,O95977,O95980,P00533,P01130,P01133,P01135,P01589,P01730,P01732,P01833,P01889,P01891,P01892,P01893,P01903,P01906,P01909,P01911,P01912,P01920,P02708,P02724,P02730,P03986,P03989,P03999,P04000,P04001,P04156,P04201,P04216,P04222,P04229,P04233,P04234,P04439,P04440,P04626,P04629,P04839,P04843,P04921,P05023,P05026,P05067,P05106,P05107,P05186,P05187,P05362,P05534,P05538,P05556,P06126,P06127,P06213,P06340,P06729,P06731,P06756,P06858,P07202,P07204,P07306,P07307,P07333,P07359,P07510,P07550,P07911,P07949,P08034,P08069,P08100,P08138,P08172,P08173,P08174,P08183,P08195,P08247,P08473,P08514,P08571,P08575,P08581,P08582,P08588,P08637,P08648,P08842,P08887,P08908,P08912,P08913,P08922,P08962,P08F94,P09131,P09326,P09564,P09603,P09619,P09693,P09758,P09848,P09923,P09958,P0C604,P0C617,P0C623,P0C626,P0C628,P0C629,P0C645,P0C646,P0C6S8,P0C7N1,P0C7N5,P0C7N8,P0C7T2,P0C7T3,P0C7U0,P0C7U3,P0CG37,P0DKB5,P10314,P10316,P10319,P10321,P10586,P10646,P10696,P10721,P10747,P10912,P10966,P11049,P11117,P11166,P11168,P11169,P11215,P11229,P11230,P11279,P11362,P11717,P11836,P11912,P12314,P12318,P12319,P12821,P12830,P13224,P13385,P13473,P13591,P13598,P13612,P13637,P13688,P13726,P13746,P13747,P13760,P13761,P13762,P13765,P13866,P13945,P13987,P14151,P14207,P14384,P14415,P14416,P14616,P14672,P14679,P14778,P14784,P14867,P15144,P15151,P15260,P15328,P15391,P15509,P15514,P15529,P15812,P15813,P15941,P16066,P16070,P16109,P16144,P16150,P16188,P16189,P16190,P16234,P16284,P16410,P16422,P16444,P16471,P16473,P16581,P16671,P16871,P17181,P17301,P17302,P17342,P17643,P17693,P17787,P17813,P17927,P17948,P18084,P18089,P18433,P18462,P18463,P18464,P18465,P18505,P18507,P18564,P18627,P18825,P18827,P19021,P19022,P19075,P19235,P19256,P19320,P19397,P19438,P19440,P19634,P20023,P20036,P20039,P20138,P20273,P20309,P20333,P20594,P20645,P20701,P20702,P20827,P20916,P21439,P21452,P21453,P21462,P21554,P21583,P21589,P21709,P21728,P21730,P21731,P21754,P21757,P21802,P21860,P21917,P21918,P21926,P22001,P22223,P22303,P22413,P22455,P22607,P22732,P22748,P22794,P22888,P22897,P23229,P23276,P23415,P23416,P23467,P23468,P23470,P23471,P23510,P23515,P23634,P23942,P23945,P23975,P24046,P24071,P24394,P24530,P25021,P25024,P25025,P25063,P25089,P25090,P25092,P25100,P25101,P25103,P25105,P25106,P25116,P25189,P25445,P25929,P25942,P26006,P26010,P26012,P26715,P26717,P26718,P26842,P26951,P26992,P27037,P27487,P27701,P27930,P28067,P28068,P28221,P28222,P28223,P28335,P28336,P28472,P28476,P28566,P28827,P28906,P28907,P28908,P29016,P29017,P29033,P29274,P29275,P29317,P29320,P29322,P29323,P29371,P29376,P29965,P29972,P30203,P30301,P30408,P30411,P30443,P30447,P30450,P30453,P30455,P30456,P30457,P30459,P30460,P30461,P30462,P30464,P30466,P30475,P30479,P30480,P30481,P30483,P30484,P30485,P30486,P30487,P30488,P30490,P30491,P30492,P30493,P30495,P30498,P30499,P30501,P30504,P30505,P30508,P30510,P30511,P30512,P30518,P30530,P30531,P30532,P30542,P30550,P30556,P30559,P30685,P30825,P30872,P30874,P30926,P30939,P30953,P30954,P30968,P30988,P30989,P31358,P31391,P31639,P31641,P31644,P31645,P31785,P31994,P31995,P31997,P32004,P32238,P32239,P32241,P32245,P32246,P32247,P32248,P32249,P32297,P32302,P32418,P32745,P32926,P32927,P32942,P32970,P32971,P33032,P33151,P33527,P33681,P33765,P34741,P34810,P34903,P34910,P34925,P34969,P34972,P34981,P34982,P34995,P34998,P35052,P35070,P35212,P35346,P35348,P35367,P35368,P35372,P35408,P35410,P35414,P35462,P35498,P35499,P35590,P35613,P35916,P35968,P36544,P36888,P36894,P36896,P36897,P36941,P37023,P37088,P37173,P37288,P38484,P38567,P38570,P39086,P40126,P40189,P40197,P40198,P40199,P40200,P40238,P40259,P40879,P40967,P41143,P41145,P41146,P41180,P41181,P41217,P41231,P41440,P41586,P41587,P41594,P41595,P41597,P41732,P41968,P42081,P42261,P42262,P42263,P42658,P42701,P42702,P42892,P43003,P43004,P43005,P43007,P43088,P43115,P43116,P43119,P43121,P43146,P43220,P43307,P43489,P43626,P43627,P43628,P43629,P43630,P43631,P43632,P43657,P43681,P46059,P46089,P46091,P46092,P46093,P46094,P46095,P46098,P46531,P46663,P46721,P47211,P47775,P47804,P47869,P47870,P47871,P47872,P47881,P47883,P47884,P47887,P47888,P47890,P47893,P47898,P47900,P47901,P48023,P48029,P48039,P48050,P48058,P48060,P48065,P48066,P48067,P48145,P48146,P48167,P48169,P48230,P48357,P48509,P48546,P48549,P48551,P48664,P48764,P48960,P49019,P49146,P49190,P49238,P49279,P49281,P49286,P49682,P49683,P49685,P49768,P49771,P49810,P49961,P50052,P50281,P50391,P50406,P50443,P50895,P50993,P51164,P51168,P51170,P51172,P51512,P51575,P51582,P51654,P51674,P51677,P51679,P51681,P51684,P51685,P51686,P51693,P51801,P51805,P51810,P51828,P52569,P52797,P52798,P52799,P52803,P52961,P53708,P53794,P53801,P53985,P54219,P54709,P54753,P54756,P54760,P54762,P54764,P54826,P54849,P54851,P54852,P55011,P55017,P55064,P55082,P55085,P55087,P55259,P55283,P55285,P55286,P55287,P55289,P55290,P55291,P55344,P55899,P56159,P56199,P56373,P56746,P56747,P56748,P56749,P56817,P56856,P56880,P57057,P57087,P57103,P57739,P58170,P58173,P58180,P58181,P58182,P58335,P58400,P58401,P58418,P58658,P58743,P59533,P59534,P59540,P59541,P59542,P59543,P59901,P59922,P60201,P60507,P60508,P60509,P60852,P60893,P61073,P61550,P61565,P61566,P61567,P61570,P62079,P62955,P78310,P78324,P78325,P78333,P78334,P78348,P78357,P78363,P78369,P78380,P78410,P78423,P78504,P78536,P78552,P78562,P79483,P80370,P81408,P82251,P82279,P98073,P98153,P98155,P98161,P98164,P98172,Q01113,Q01118,Q01151,Q01344,Q01638,Q01650,Q01718,Q01726,Q01814,Q01959,Q01973,Q01974,Q02094,Q02223,Q02246,Q02297,Q02413,Q02487,Q02505,Q02643,Q02763,Q03167,Q03405,Q03431,Q04609,Q04721,Q04771,Q04826,Q04844,Q04900,Q04912,Q05586,Q05901,Q05940,Q05996,Q06418,Q06432,Q06481,Q06495,Q07000,Q07001,Q07011,Q07075,Q07108,Q07444,Q07699,Q07837,Q07954,Q08174,Q08334,Q08345,Q08357,Q08462,Q08554,Q08708,Q08722,Q08AI6,Q08ET2,Q09160,Q0D2K0,Q0P6H9,Q10588,Q10589,Q12770,Q12836,Q12860,Q12864,Q12866,Q12879,Q12884,Q12891,Q12907,Q12908,Q12913,Q12918,Q12999,Q13002,Q13003,Q13018,Q13145,Q13183,Q13224,Q13255,Q13258,Q13261,Q13286,Q13291,Q13304,Q13308,Q13324,Q13332,Q13336,Q13349,Q13410,Q13421,Q13433,Q13443,Q13444,Q13449,Q13467,Q13478,Q13488,Q13491,Q13508,Q13530,Q13563,Q13585,Q13586,Q13591,Q13606,Q13607,Q13621,Q13634,Q13635,Q13639,Q13641,Q13651,Q13683,Q13705,Q13733,Q13740,Q13797,Q13873,Q13936,Q14002,Q14108,Q14114,Q14118,Q14126,Q14162,Q14210,Q14242,Q14246,Q14330,Q14332,Q14392,Q14416,Q14439,Q14500,Q14517,Q14524,Q14542,Q14574,Q14626,Q14627,Q14714,Q14773,Q14831,Q14832,Q14833,Q14916,Q14943,Q14952,Q14953,Q14954,Q14956,Q14957,Q14973,Q14982,Q14C87,Q14CN2,Q14CZ8,Q14DG7,Q15043,Q15077,Q15109,Q15116,Q15125,Q15223,Q15256,Q15262,Q15303,Q15375,Q15391,Q15399,Q15612,Q15614,Q15615,Q15617,Q15619,Q15620,Q15622,Q15722,Q15743,Q15758,Q15760,Q15761,Q15762,Q15768,Q15822,Q15825,Q15849,Q15858,Q16099,Q16288,Q16348,Q16445,Q16478,Q16538,Q16553,Q16558,Q16563,Q16570,Q16572,Q16581,Q16585,Q16586,Q16602,Q16620,Q16651,Q16653,Q16671,Q16720,Q16819,Q16820,Q16827,Q16832,Q16849,Q16853,Q16880,Q17R55,Q17RY6,Q19T08,Q1EHB4,Q1HG43,Q24JP5,Q29718,Q29836,Q29865,Q29940,Q29960,Q29963,Q29974,Q29980,Q29983,Q2HXU8,Q2KHT4,Q2M385,Q2M3G0,Q2M3M2,Q2VWP7,Q2Y0W8,Q30134,Q30154,Q30167,Q30201,Q31610,Q31612,Q32ZL2,Q3KNS1,Q3KNT9,Q3KNW5,Q3KPI0,Q3KR37,Q3MIW9,Q3SXP7,Q3SXY7,Q3ZCQ3,Q401N2,Q495A1,Q495M3,Q495N2,Q495T6,Q496F6,Q496H8,Q496J9,Q49SQ1,Q4G0T1,Q4KMG0,Q4KMQ2,Q4KMZ8,Q4U2R8,Q4VNC0,Q504Y0,Q50LG9,Q53EL9,Q53GD3,Q53R12,Q58EX2,Q5BIV9,Q5BKX6,Q5DID0,Q5DX21,Q5FWE3,Q5GH77,Q5HYA8,Q5IJ48,Q5JQS5,Q5JRS4,Q5JRV8,Q5JXA9,Q5JXX5,Q5JZY3,Q5M7Z0,Q5NUL3,Q5PT55,Q5QGZ9,Q5R3F8,Q5SQ64,Q5SSG8,Q5SY80,Q5SZK8,Q5T1S8,Q5T2D2,Q5T3F8,Q5T3U5,Q5T442,Q5T4F4,Q5T601,Q5T6X5,Q5T848,Q5TF39,Q5TFQ8,Q5TZ20,Q5UAW9,Q5VT99,Q5VU13,Q5VU65,Q5VU97,Q5VUB5,Q5VV43,Q5VV63,Q5VW38,Q5VWK5,Q5VX71,Q5VXU1,Q5VY43,Q5VY80,Q5VYJ5,Q5VZ72,Q5XXA6,Q5Y7A7,Q5ZPR3,Q685J3,Q687X5,Q68D85,Q68DH5,Q68DV7,Q69384,Q695T7,Q6DN72,Q6DWJ6,Q6EMK4,Q6GTX8,Q6GV28,Q6H3X3,Q6IA17,Q6ICI0,Q6IEE7,Q6IEU7,Q6IEV9,Q6IEY1,Q6IEZ7,Q6IF00,Q6IF42,Q6IF63,Q6IF82,Q6IF99,Q6IFG1,Q6IFH4,Q6IFN5,Q6ISU1,Q6IWH7,Q6J4K2,Q6MZM0,Q6N022,Q6N075,Q6NUJ2,Q6NUS6,Q6NUT3,Q6NV75,Q6NVV3,Q6NW40,Q6P1J6,Q6P499,Q6P4Q7,Q6P5W5,Q6P995,Q6P9G4,Q6PCB8,Q6PEY0,Q6PI73,Q6PIZ9,Q6PJF5,Q6PJG9,Q6PRD1,Q6PXP3,Q6Q8B3,Q6QNK2,Q6T423,Q6U736,Q6U841,Q6UQ28,Q6UVK1,Q6UW56,Q6UW88,Q6UWB1,Q6UWI2,Q6UWJ1,Q6UWJ8,Q6UWL2,Q6UWL6,Q6UWN0,Q6UWN5,Q6UWR7,Q6UWV2,Q6UX01,Q6UX15,Q6UX27,Q6UX71,Q6UX82,Q6UXB3,Q6UXB4,Q6UXB8,Q6UXC1,Q6UXD5,Q6UXD7,Q6UXE8,Q6UXF1,Q6UXG2,Q6UXG3,Q6UXG8,Q6UXK2,Q6UXK5,Q6UXL0,Q6UXM1,Q6UXN8,Q6UXU4,Q6UXV0,Q6UXZ0,Q6UXZ3,Q6UXZ4,Q6UY09,Q6UY11,Q6UY18,Q6V0I7,Q6V1P9,Q6W5P4,Q6YBV0,Q6YHK3,Q6ZMB5,Q6ZMC9,Q6ZMD2,Q6ZMH5,Q6ZMI3,Q6ZMJ2,Q6ZN44,Q6ZNA5,Q6ZP29,Q6ZP80,Q6ZQN7,Q6ZRH7,Q6ZRP7,Q6ZS10,Q6ZSJ9,Q6ZSM3,Q6ZSS7,Q6ZTQ4,Q6ZUK4,Q6ZVL6,Q6ZVN8,Q6ZW05,Q70Z44,Q71RS6,Q75T13,Q75V66,Q7KYR7,Q7L0J3,Q7L0X0,Q7L1I2,Q7L985,Q7LBE3,Q7RTM1,Q7RTS6,Q7RTT9,Q7RTW8,Q7RTX0,Q7RTX1,Q7RTY9,Q7Z2D5,Q7Z2H8,Q7Z2K6,Q7Z3B1,Q7Z3C6,Q7Z3D4,Q7Z3F1,Q7Z3Q1,Q7Z3T1,Q7Z402,Q7Z408,Q7Z418,Q7Z442,Q7Z443,Q7Z4F1,Q7Z553,Q7Z5H5,Q7Z5N4,Q7Z601,Q7Z602,Q7Z692,Q7Z6A9,Q7Z6M3,Q7Z7D3,Q7Z7J7,Q7Z7M0,Q7Z7M1,Q7Z7N9,Q86SJ2,Q86SJ6,Q86SM5,Q86SM8,Q86SP6,Q86SQ3,Q86SQ4,Q86SQ6,Q86SU0,Q86T13,Q86T26,Q86TG1,Q86TL2,Q86TY3,Q86UF1,Q86UG4,Q86UK0,Q86UK5,Q86UN2,Q86UN3,Q86UP0,Q86UP6,Q86UQ4,Q86UW1,Q86UW2,Q86V24,Q86V40,Q86V85,Q86VB7,Q86VE9,Q86VH4,Q86VH5,Q86VR7,Q86VW1,Q86VZ1,Q86VZ4,Q86W33,Q86W47,Q86WB7,Q86WC4,Q86WI0,Q86WI1,Q86WK6,Q86WK7,Q86XK7,Q86XM0,Q86XR5,Q86XS8,Q86XT9,Q86XX4,Q86Y34,Q86YC3,Q86YD3,Q86YD5,Q86YT9,Q8IU57,Q8IU80,Q8IUA7,Q8IUH8,Q8IUK5,Q8IUW5,Q8IV16,Q8IVJ1,Q8IVM8,Q8IVU1,Q8IVW8,Q8IW00,Q8IW52,Q8IWA5,Q8IWD5,Q8IWK6,Q8IWT1,Q8IWV2,Q8IX05,Q8IXE1,Q8IXH8,Q8IY34,Q8IYL9,Q8IYR6,Q8IYV9,Q8IZD6,Q8IZF0,Q8IZF2,Q8IZF3,Q8IZF4,Q8IZF5,Q8IZF6,Q8IZF7,Q8IZJ1,Q8IZP9,Q8IZU9,Q8IZY2,Q8J025,Q8N0W4,Q8N0Y3,Q8N0Y5,Q8N0Z9,Q8N109,Q8N126,Q8N127,Q8N130,Q8N131,Q8N139,Q8N146,Q8N148,Q8N149,Q8N158,Q8N162,Q8N1C3,Q8N1N2,Q8N271,Q8N2G4,Q8N2Q7,Q8N349,Q8N370,Q8N386,Q8N387,Q8N3F9,Q8N3J6,Q8N3T6,Q8N423,Q8N434,Q8N441,Q8N468,Q8N4K4,Q8N4M1,Q8N5C1,Q8N5U1,Q8N608,Q8N628,Q8N695,Q8N697,Q8N6C5,Q8N6F1,Q8N6P7,Q8N6Q3,Q8N6U8,Q8N6Y1,Q8N743,Q8N7C0,Q8N7C4,Q8N7P1,Q8N7P3,Q8N7X8,Q8N8D7,Q8N8F7,Q8N8Z6,Q8N967,Q8NA29,Q8NAC3,Q8NAU1,Q8NBI5,Q8NBJ9,Q8NBL3,Q8NBM4,Q8NBN3,Q8NBR0,Q8NBT3,Q8NBW4,Q8NC01,Q8NC42,Q8NC54,Q8NC67,Q8NCC5,Q8NCG7,Q8NCL8,Q8NCS7,Q8NCW0,Q8ND94,Q8NDV2,Q8NDX2,Q8NE00,Q8NE01,Q8NE79,Q8NEA5,Q8NER5,Q8NET5,Q8NFF2,Q8NFJ5,Q8NFJ6,Q8NFK1,Q8NFM7,Q8NFN8,Q8NFP4,Q8NFR9,Q8NFT8,Q8NFU0,Q8NFU1,Q8NFY4,Q8NFZ3,Q8NFZ4,Q8NFZ6,Q8NFZ8,Q8NG11,Q8NG75,Q8NG76,Q8NG77,Q8NG78,Q8NG80,Q8NG81,Q8NG83,Q8NG84,Q8NG85,Q8NG92,Q8NG94,Q8NG95,Q8NG97,Q8NG98,Q8NG99,Q8NGA0,Q8NGA1,Q8NGA2,Q8NGA4,Q8NGA5,Q8NGA6,Q8NGA8,Q8NGB2,Q8NGB4,Q8NGB6,Q8NGB8,Q8NGB9,Q8NGC0,Q8NGC1,Q8NGC2,Q8NGC3,Q8NGC4,Q8NGC5,Q8NGC6,Q8NGC7,Q8NGC8,Q8NGC9,Q8NGD0,Q8NGD1,Q8NGD2,Q8NGD3,Q8NGD4,Q8NGD5,Q8NGE0,Q8NGE1,Q8NGE2,Q8NGE3,Q8NGE5,Q8NGE7,Q8NGE8,Q8NGE9,Q8NGF0,Q8NGF1,Q8NGF3,Q8NGF4,Q8NGF6,Q8NGF7,Q8NGF8,Q8NGF9,Q8NGG0,Q8NGG1,Q8NGG2,Q8NGG3,Q8NGG4,Q8NGG5,Q8NGG6,Q8NGG7,Q8NGG8,Q8NGH3,Q8NGH5,Q8NGH6,Q8NGH7,Q8NGH8,Q8NGH9,Q8NGI0,Q8NGI1,Q8NGI2,Q8NGI3,Q8NGI4,Q8NGI6,Q8NGI7,Q8NGI8,Q8NGI9,Q8NGJ0,Q8NGJ1,Q8NGJ2,Q8NGJ3,Q8NGJ4,Q8NGJ5,Q8NGJ6,Q8NGJ7,Q8NGJ8,Q8NGJ9,Q8NGK0,Q8NGK1,Q8NGK2,Q8NGK3,Q8NGK4,Q8NGK5,Q8NGK6,Q8NGK9,Q8NGL0,Q8NGL1,Q8NGL2,Q8NGL3,Q8NGL4,Q8NGL6,Q8NGL7,Q8NGL9,Q8NGM1,Q8NGM8,Q8NGM9,Q8NGN0,Q8NGN1,Q8NGN2,Q8NGN3,Q8NGN4,Q8NGN5,Q8NGN6,Q8NGN7,Q8NGN8,Q8NGP0,Q8NGP2,Q8NGP3,Q8NGP4,Q8NGP6,Q8NGP8,Q8NGP9,Q8NGQ1,Q8NGQ2,Q8NGQ3,Q8NGQ4,Q8NGQ5,Q8NGQ6,Q8NGR1,Q8NGR2,Q8NGR3,Q8NGR4,Q8NGR5,Q8NGR6,Q8NGR8,Q8NGR9,Q8NGS0,Q8NGS1,Q8NGS2,Q8NGS3,Q8NGS4,Q8NGS5,Q8NGS6,Q8NGS7,Q8NGS8,Q8NGS9,Q8NGT0,Q8NGT1,Q8NGT2,Q8NGT5,Q8NGT7,Q8NGT9,Q8NGU1,Q8NGU2,Q8NGU4,Q8NGU9,Q8NGV0,Q8NGV5,Q8NGV6,Q8NGV7,Q8NGW1,Q8NGW6,Q8NGX0,Q8NGX1,Q8NGX2,Q8NGX3,Q8NGX5,Q8NGX6,Q8NGX8,Q8NGX9,Q8NGY0,Q8NGY1,Q8NGY2,Q8NGY3,Q8NGY5,Q8NGY6,Q8NGY7,Q8NGY9,Q8NGZ0,Q8NGZ2,Q8NGZ3,Q8NGZ4,Q8NGZ5,Q8NGZ6,Q8NGZ9,Q8NH00,Q8NH01,Q8NH02,Q8NH03,Q8NH04,Q8NH05,Q8NH06,Q8NH07,Q8NH08,Q8NH09,Q8NH10,Q8NH16,Q8NH18,Q8NH19,Q8NH21,Q8NH37,Q8NH40,Q8NH41,Q8NH42,Q8NH43,Q8NH48,Q8NH49,Q8NH50,Q8NH51,Q8NH53,Q8NH54,Q8NH55,Q8NH56,Q8NH57,Q8NH59,Q8NH60,Q8NH61,Q8NH63,Q8NH64,Q8NH67,Q8NH69,Q8NH70,Q8NH72,Q8NH73,Q8NH74,Q8NH76,Q8NH79,Q8NH80,Q8NH81,Q8NH83,Q8NH85,Q8NH87,Q8NH89,Q8NH90,Q8NH92,Q8NH93,Q8NH94,Q8NH95,Q8NHA4,Q8NHA6,Q8NHA8,Q8NHB1,Q8NHB7,Q8NHB8,Q8NHC4,Q8NHC5,Q8NHC6,Q8NHC7,Q8NHC8,Q8NHK3,Q8NHL6,Q8NHS3,Q8NHV5,Q8NI17,Q8NI32,Q8TAB3,Q8TAF8,Q8TB96,Q8TBB6,Q8TBE3,Q8TBG9,Q8TBJ4,Q8TBP5,Q8TC27,Q8TCB6,Q8TCC7,Q8TCJ2,Q8TCT7,Q8TCT8,Q8TCT9,Q8TCU5,Q8TCW7,Q8TCW9,Q8TD07,Q8TD20,Q8TD46,Q8TD84,Q8TDB8,Q8TDF5,Q8TDM5,Q8TDN1,Q8TDN2,Q8TDQ0,Q8TDQ1,Q8TDS4,Q8TDS5,Q8TDS7,Q8TDT2,Q8TDU6,Q8TDU9,Q8TDV0,Q8TDV2,Q8TDV5,Q8TDW7,Q8TDX9,Q8TDY8,Q8TE23,Q8TEB7,Q8TED4,Q8TEM1,Q8TEQ8,Q8TF66,Q8WTR4,Q8WTV0,Q8WUD6,Q8WUG5,Q8WUT4,Q8WUX1,Q8WV15,Q8WVE6,Q8WVN6,Q8WVP7,Q8WVV5,Q8WWA0,Q8WWB7,Q8WWF5,Q8WWG1,Q8WWI5,Q8WWQ8,Q8WWT9,Q8WWV6,Q8WWX8,Q8WWZ7,Q8WXA8,Q8WXD0,Q8WXG9,Q8WXI7,Q8WXS4,Q8WXS5,Q8WY07,Q8WY21,Q8WYK1,Q8WZ71,Q8WZ84,Q8WZ92,Q8WZ94,Q8WZA6,Q902F8,Q902F9,Q92508,Q92536,Q92537,Q92542,Q92544,Q92581,Q92629,Q92633,Q92637,Q92673,Q92692,Q92729,Q92823,Q92824,Q92838,Q92847,Q92854,Q92859,Q92887,Q92911,Q92932,Q92956,Q92959,Q93033,Q93038,Q93070,Q95365,Q95460,Q95604,Q95IE3,Q969F8,Q969I6,Q969N2,Q969N4,Q969P0,Q969V1,Q969W9,Q969Z4,Q96A25,Q96A28,Q96AM1,Q96AP7,Q96B86,Q96BD0,Q96BF3,Q96CE8,Q96CH1,Q96CW9,Q96D42,Q96DD7,Q96DU3,Q96E93,Q96EP9,Q96F05,Q96F46,Q96F81,Q96FE5,Q96FE7,Q96FL8,Q96FT7,Q96FV3,Q96G91,Q96GP6,Q96GW7,Q96GX1,Q96GZ6,Q96H15,Q96HA4,Q96IQ7,Q96J42,Q96J65,Q96J66,Q96J84,Q96JA1,Q96JJ7,Q96JP9,Q96JQ0,Q96JW4,Q96K49,Q96K78,Q96KG7,Q96KJ4,Q96KK3,Q96KK4,Q96KV6,Q96L08,Q96LA5,Q96LA6,Q96LA9,Q96LB0,Q96LB1,Q96LB2,Q96LC7,Q96LD1,Q96MC6,Q96MS0,Q96MU8,Q96N19,Q96N87,Q96NI6,Q96NR3,Q96NT5,Q96NU0,Q96NY8,Q96P31,Q96P65,Q96P66,Q96P67,Q96P68,Q96P69,Q96P88,Q96PB1,Q96PB8,Q96PD2,Q96PE1,Q96PE5,Q96PJ5,Q96PL2,Q96PL5,Q96PQ0,Q96PQ1,Q96PS8,Q96PX8,Q96PZ7,Q96QD8,Q96QE2,Q96QE4,Q96QU1,Q96R08,Q96R09,Q96R27,Q96R28,Q96R30,Q96R45,Q96R47,Q96R48,Q96R54,Q96R67,Q96R69,Q96R72,Q96R84,Q96RA2,Q96RB7,Q96RC9,Q96RD0,Q96RD1,Q96RD2,Q96RD3,Q96RD6,Q96RD7,Q96RD9,Q96RI0,Q96RI8,Q96RI9,Q96RJ0,Q96RL6,Q96RN1,Q96RV3,Q96S37,Q96S97,Q96SA4,Q96SJ8,Q96SL1,Q96T83,Q99062,Q99075,Q99102,Q99250,Q99445,Q99463,Q99466,Q99467,Q99500,Q99523,Q99527,Q99571,Q99572,Q99650,Q99665,Q99677,Q99678,Q99679,Q99680,Q99705,Q99706,Q99758,Q99788,Q99795,Q99805,Q99808,Q99835,Q99884,Q99928,Q99965,Q9BPV8,Q9BQ51,Q9BQS7,Q9BQT9,Q9BRK3,Q9BS91,Q9BSA4,Q9BSN7,Q9BT76,Q9BTN0,Q9BUF7,Q9BWQ8,Q9BWV1,Q9BX67,Q9BX97,Q9BXA5,Q9BXB1,Q9BXC0,Q9BXC1,Q9BXE9,Q9BXJ7,Q9BXN2,Q9BXP2,Q9BXR5,Q9BXS9,Q9BXT2,Q9BY07,Q9BY10,Q9BY14,Q9BY15,Q9BY21,Q9BY67,Q9BY79,Q9BYE2,Q9BYE9,Q9BYF1,Q9BYT1,Q9BYT9,Q9BYW1,Q9BZ11,Q9BZ76,Q9BZA7,Q9BZA8,Q9BZC7,Q9BZD2,Q9BZG2,Q9BZG9,Q9BZJ6,Q9BZJ7,Q9BZJ8,Q9BZM4,Q9BZM5,Q9BZM6,Q9BZR6,Q9BZV2,Q9BZV3,Q9BZW2,Q9BZW8,Q9BZZ2,Q9C0A0,Q9C0B5,Q9C0C4,Q9C0H2,Q9C0I4,Q9C0K1,Q9GIY3,Q9GZK3,Q9GZK4,Q9GZK6,Q9GZK7,Q9GZM6,Q9GZN0,Q9GZN6,Q9GZP7,Q9GZQ4,Q9GZQ6,Q9GZU1,Q9GZV3,Q9GZZ6,Q9GZZ7,Q9H013,Q9H015,Q9H0V9,Q9H156,Q9H158,Q9H159,Q9H172,Q9H195,Q9H1C0,Q9H1C4,Q9H1D0,Q9H1E5,Q9H1U4,Q9H1V8,Q9H1Y3,Q9H205,Q9H207,Q9H208,Q9H209,Q9H210,Q9H222,Q9H228,Q9H244,Q9H251,Q9H255,Q9H295,Q9H2A7,Q9H2B4,Q9H2C5,Q9H2C8,Q9H2E6,Q9H2H9,Q9H2J7,Q9H2U9,Q9H2W1,Q9H2X3,Q9H2X9,Q9H2Y9,Q9H313,Q9H330,Q9H339,Q9H340,Q9H341,Q9H342,Q9H343,Q9H344,Q9H346,Q9H3N8,Q9H3R2,Q9H3S1,Q9H3S3,Q9H3T2,Q9H3T3,Q9H3W5,Q9H461,Q9H4A9,Q9H4B8,Q9H4D0,Q9H598,Q9H5I5,Q9H5V8,Q9H5Y7,Q9H665,Q9H6B4,Q9H6D8,Q9H6L2,Q9H6X2,Q9H6Y7,Q9H756,Q9H7F0,Q9H7M9,Q9H841,Q9H8J5,Q9H8M5,Q9H8X9,Q9H9K5,Q9H9P2,Q9H9S5,Q9HA72,Q9HAB3,Q9HAR2,Q9HAS3,Q9HAV5,Q9HB29,Q9HB89,Q9HBA0,Q9HBB8,Q9HBE5,Q9HBG7,Q9HBJ8,Q9HBL6,Q9HBT6,Q9HBV2,Q9HBW0,Q9HBW1,Q9HBW9,Q9HBX8,Q9HBX9,Q9HC56,Q9HC58,Q9HC73,Q9HC97,Q9HCC8,Q9HCJ1,Q9HCJ2,Q9HCK4,Q9HCL0,Q9HCM2,Q9HCN3,Q9HCN6,Q9HCU0,Q9HCU4,Q9HD20,Q9HD43,Q9HD45,Q9HDB5,Q9N2J8,Q9N2K0,Q9NP59,Q9NP60,Q9NP78,Q9NP91,Q9NP94,Q9NP99,Q9NPA1,Q9NPA2,Q9NPB9,Q9NPC1,Q9NPD5,Q9NPD7,Q9NPF0,Q9NPG1,Q9NPG4,Q9NPH3,Q9NPH5,Q9NPR2,Q9NPR9,Q9NPU4,Q9NPY3,Q9NQ11,Q9NQ25,Q9NQ34,Q9NQ40,Q9NQ60,Q9NQ84,Q9NQ90,Q9NQA5,Q9NQN1,Q9NQS3,Q9NQS5,Q9NQX7,Q9NR16,Q9NR61,Q9NR96,Q9NR97,Q9NRA2,Q9NRD8,Q9NRD9,Q9NRJ7,Q9NRM0,Q9NRM6,Q9NRR2,Q9NRX5,Q9NS62,Q9NS64,Q9NS66,Q9NS67,Q9NS68,Q9NS75,Q9NS82,Q9NS93,Q9NSA0,Q9NSD5,Q9NSD7,Q9NSI5,Q9NT68,Q9NT99,Q9NTG1,Q9NTN9,Q9NTQ9,Q9NU53,Q9NUM3,Q9NUM4,Q9NUN5,Q9NV12,Q9NV96,Q9NWF4,Q9NX52,Q9NX61,Q9NX77,Q9NXL6,Q9NY15,Q9NY25,Q9NY35,Q9NY37,Q9NY46,Q9NY64,Q9NY72,Q9NY84,Q9NY91,Q9NYB5,Q9NYJ7,Q9NYK1,Q9NYM4,Q9NYQ6,Q9NYQ7,Q9NYQ8,Q9NYV7,Q9NYV8,Q9NYW0,Q9NYW1,Q9NYW2,Q9NYW3,Q9NYW5,Q9NYW6,Q9NYW7,Q9NYX4,Q9NYZ4,Q9NZ53,Q9NZ94,Q9NZC2,Q9NZD1,Q9NZH0,Q9NZM1,Q9NZN1,Q9NZP0,Q9NZP2,Q9NZP5,Q9NZQ7,Q9NZR2,Q9NZS2,Q9NZU0,Q9NZU1,Q9NZV1,Q9P0K1,Q9P0L9,Q9P0T7,Q9P0V8,Q9P0X4,Q9P121,Q9P126,Q9P1P5,Q9P1Q5,Q9P1W3,Q9P1W8,Q9P232,Q9P244,Q9P273,Q9P283,Q9P296,Q9P2B2,Q9P2E7,Q9P2J2,Q9P2K9,Q9P2S2,Q9P2U7,Q9P2U8,Q9P2V4,Q9TNN7,Q9TQE0,Q9UBD6,Q9UBG0,Q9UBL9,Q9UBN1,Q9UBN6,Q9UBS5,Q9UBS9,Q9UBY0,Q9UBY5,Q9UEF7,Q9UF02,Q9UF33,Q9UGF5,Q9UGF6,Q9UGF7,Q9UGH3,Q9UGM1,Q9UGN4,Q9UGQ3,Q9UGT4,Q9UHC6,Q9UHC9,Q9UHF4,Q9UHI7,Q9UHM6,Q9UHP7,Q9UHW9,Q9UHX3,Q9UI33,Q9UI40,Q9UIB8,Q9UIG8,Q9UIK5,Q9UIQ6,Q9UIW2,Q9UJ14,Q9UJ42,Q9UJ99,Q9UJA9,Q9UJQ1,Q9UK23,Q9UKB5,Q9UKF2,Q9UKF5,Q9UKG4,Q9UKH3,Q9UKJ0,Q9UKJ1,Q9UKJ8,Q9UKL2,Q9UKL4,Q9UKN1,Q9UKP6,Q9UKQ2,Q9UKU6,Q9UKX5,Q9UKY0,Q9UKZ4,Q9UL52,Q9ULB1,Q9ULB4,Q9ULB5,Q9ULC0,Q9ULF5,Q9ULH4,Q9ULI3,Q9ULK0,Q9ULK6,Q9ULL4,Q9ULQ1,Q9ULV1,Q9ULW2,Q9ULX7,Q9ULZ9,Q9UM44,Q9UM47,Q9UM73,Q9UMF0,Q9UMX9,Q9UMZ3,Q9UN42,Q9UN66,Q9UN67,Q9UN70,Q9UN71,Q9UN72,Q9UN73,Q9UN74,Q9UN75,Q9UN76,Q9UN88,Q9UNE0,Q9UNG2,Q9UNN8,Q9UNQ0,Q9UNW8,Q9UP38,Q9UP95,Q9UPC5,Q9UPI3,Q9UPR5,Q9UPU3,Q9UPX0,Q9UPZ6,Q9UQ52,Q9UQC9,Q9UQD0,Q9UQF0,Q9UQQ1,Q9UQV4,Q9Y219,Q9Y226,Q9Y267,Q9Y271,Q9Y275,Q9Y286,Q9Y287,Q9Y289,Q9Y2C9,Q9Y2I2,Q9Y2T5,Q9Y2T6,Q9Y336,Q9Y345,Q9Y3B3,Q9Y3N9,Q9Y3P8,Q9Y3Q7,Q9Y487,Q9Y493,Q9Y4A9,Q9Y4C0,Q9Y4D2,Q9Y4D7,Q9Y561,Q9Y585,Q9Y5E1,Q9Y5E2,Q9Y5E3,Q9Y5E4,Q9Y5E5,Q9Y5E6,Q9Y5E7,Q9Y5E8,Q9Y5E9,Q9Y5F0,Q9Y5F1,Q9Y5F2,Q9Y5F3,Q9Y5F6,Q9Y5F7,Q9Y5F8,Q9Y5F9,Q9Y5G0,Q9Y5G1,Q9Y5G2,Q9Y5G3,Q9Y5G4,Q9Y5G5,Q9Y5G6,Q9Y5G7,Q9Y5G8,Q9Y5G9,Q9Y5H0,Q9Y5H1,Q9Y5H2,Q9Y5H3,Q9Y5H4,Q9Y5H5,Q9Y5H6,Q9Y5H7,Q9Y5H8,Q9Y5H9,Q9Y5I0,Q9Y5I1,Q9Y5I2,Q9Y5I3,Q9Y5I4,Q9Y5I7,Q9Y5N1,Q9Y5P0,Q9Y5P1,Q9Y5Q5,Q9Y5S1,Q9Y5U5,Q9Y5X5,Q9Y5Y0,Q9Y5Y3,Q9Y5Y4,Q9Y5Y7,Q9Y5Y9,Q9Y5Z0,Q9Y624,Q9Y625,Q9Y639,Q9Y653,Q9Y666,Q9Y691,Q9Y694,Q9Y698,Q9Y6F6,Q9Y6H8,Q9Y6L6,Q9Y6M0,Q9Y6M5,Q9Y6M7,Q9Y6N7,Q9Y6N8,Q9Y6Q6,Q9Y6R1,Q9Y6W8,Q9Y6X5';
const _MA_RAW='A0A5B9,A0AV02,A0AVI2,A0AVI4,A0FGR8,A0FGR9,A0PJK1,A0PJW6,A0PJX4,A0PJX8,A0PJZ3,A0PK00,A0PK05,A0PK11,A0ZSE6,A1A5B4,A1A5C7,A1KXE4,A1L0T0,A1L157,A1L1A6,A1L3X0,A1L4H1,A2A2V5,A2A2Y4,A2CJ06,A2RRL7,A2RU14,A2RU48,A2RU67,A2RUG3,A2RUT3,A2VDJ0,A2VEC9,A3KFT3,A4D0S4,A4D0T2,A4D1S0,A4D1S5,A4D256,A4D2G3,A4D2H0,A4FU28,A4IF30,A5D6W6,A5PKW4,A5PLK6,A5PLL7,A5X5Y0,A6BM72,A6H8M9,A6NC51,A6NC97,A6NCI5,A6NCQ9,A6NCV1,A6ND48,A6NDA9,A6NDD5,A6NDH6,A6NDL8,A6NDP7,A6NDV4,A6NEH6,A6NET4,A6NF34,A6NF89,A6NFA1,A6NFC5,A6NFC9,A6NFE2,A6NFR6,A6NFU0,A6NFX1,A6NFY4,A6NGA9,A6NGB0,A6NGB7,A6NGC4,A6NGU5,A6NGY5,A6NGZ8,A6NH00,A6NH21,A6NH52,A6NHA9,A6NHG9,A6NHS7,A6NI61,A6NI73,A6NIJ9,A6NIL9,A6NIM6,A6NJW4,A6NJW9,A6NJY1,A6NJZ3,A6NK97,A6NKB5,A6NKC4,A6NKF7,A6NKG5,A6NKK0,A6NKL6,A6NKW6,A6NKX4,A6NL05,A6NL08,A6NL26,A6NL88,A6NL99,A6NLE4,A6NLU5,A6NLX4,A6NM03,A6NM10,A6NM11,A6NM45,A6NM62,A6NM76,A6NMB1,A6NMD0,A6NML5,A6NMS3,A6NMS7,A6NMU1,A6NMY6,A6NMZ5,A6NMZ7,A6NN92,A6NNB3,A6NND4,A6NNE9,A6NNN8,A6NNS2,A6PVL3,A6QL63,A6XGL0,A7E2Y1,A7MBM2,A8CG34,A8K4G0,A8K7I4,A8MPY1,A8MV81,A8MVG2,A8MVS5,A8MVW0,A8MVW5,A8MVZ5,A8MWK0,A8MWL7,A8MWY0,A8MXK1,A8MXV6,A8MYB1,A8MYU2,A8MZ97,A8TX70,A9Z1Z3,B0FP48,B0I1T2,B0YJ81,B1AH88,B1AL88,B2RN74,B2RTY4,B2RUZ4,B2RXF0,B3SHH9,B4DJY2,B4DS77,B4DYI2,B6A8C7,B6SEH8,B6SEH9,B7U540,B7Z8K6,B7ZAQ6,B8ZZ34,B9EJG8,C9JDP6,C9JH25,C9JI98,C9JQI7,C9JVW0,D3W0D1,E2RYF6,E7ERA6,E9PQ53,F2Z333,F5H4A9,F7VJQ1,G3V0H7,H0YL14,H3BR10,H3BS89,H3BV60,O00144,O00155,O00161,O00168,O00180,O00194,O00198,O00206,O00212,O00213,O00219,O00220,O00222,O00237,O00238,O00241,O00254,O00258,O00264,O00270,O00299,O00322,O00337,O00341,O00391,O00398,O00400,O00408,O00421,O00443,O00445,O00451,O00453,O00461,O00468,O00476,O00478,O00481,O00501,O00515,O00519,O00526,O00533,O00548,O00555,O00559,O00574,O00587,O00590,O00591,O00592,O00623,O00624,O00631,O00634,O00744,O00755,O00767,O14490,O14493,O14494,O14495,O14511,O14514,O14520,O14521,O14522,O14523,O14524,O14525,O14526,O14569,O14581,O14609,O14610,O14626,O14638,O14640,O14641,O14649,O14653,O14654,O14662,O14668,O14669,O14672,O14678,O14681,O14683,O14684,O14718,O14735,O14763,O14764,O14786,O14788,O14798,O14804,O14807,O14817,O14828,O14836,O14842,O14843,O14863,O14880,O14894,O14904,O14905,O14910,O14917,O14925,O14931,O14939,O14944,O14957,O14966,O14967,O14975,O14983,O15031,O15033,O15034,O15056,O15072,O15079,O15118,O15120,O15121,O15126,O15127,O15146,O15155,O15162,O15165,O15173,O15197,O15218,O15229,O15230,O15239,O15243,O15244,O15245,O15247,O15258,O15260,O15269,O15270,O15303,O15321,O15335,O15342,O15354,O15374,O15375,O15389,O15393,O15394,O15399,O15400,O15403,O15427,O15431,O15432,O15438,O15439,O15440,O15455,O15466,O15482,O15503,O15529,O15533,O15547,O15551,O15552,O15554,O43155,O43157,O43169,O43173,O43184,O43193,O43194,O43246,O43280,O43286,O43291,O43292,O43295,O43300,O43306,O43315,O43405,O43424,O43451,O43462,O43464,O43490,O43493,O43497,O43505,O43506,O43508,O43511,O43520,O43525,O43526,O43529,O43556,O43557,O43559,O43561,O43567,O43570,O43581,O43603,O43613,O43614,O43653,O43657,O43674,O43676,O43677,O43688,O43699,O43731,O43736,O43739,O43749,O43752,O43759,O43760,O43761,O43768,O43772,O43808,O43825,O43826,O43861,O43868,O43869,O43889,O43895,O43908,O43909,O43914,O43916,O43921,O43934,O60235,O60238,O60241,O60242,O60243,O60245,O60262,O60266,O60279,O60309,O60312,O60313,O60320,O60330,O60331,O60337,O60344,O60353,O60359,O60391,O60403,O60404,O60412,O60423,O60427,O60431,O60449,O60462,O60469,O60476,O60478,O60486,O60487,O60488,O60499,O60500,O60503,O60507,O60512,O60513,O60602,O60603,O60609,O60610,O60635,O60636,O60637,O60641,O60656,O60667,O60669,O60704,O60706,O60721,O60725,O60733,O60741,O60755,O60774,O60779,O60830,O60831,O60840,O60844,O60858,O60882,O60883,O60884,O60894,O60895,O60896,O60906,O60909,O60928,O60931,O60938,O60939,O75015,O75019,O75022,O75023,O75027,O75044,O75051,O75054,O75056,O75063,O75069,O75071,O75072,O75074,O75077,O75078,O75084,O75096,O75106,O75110,O75121,O75128,O75129,O75144,O75161,O75173,O75185,O75192,O75197,O75204,O75264,O75298,O75309,O75310,O75311,O75324,O75325,O75326,O75339,O75340,O75352,O75354,O75355,O75356,O75365,O75379,O75387,O75388,O75396,O75398,O75420,O75425,O75427,O75438,O75443,O75445,O75452,O75460,O75473,O75477,O75487,O75508,O75509,O75558,O75578,O75581,O75592,O75631,O75695,O75712,O75718,O75746,O75747,O75751,O75752,O75762,O75781,O75783,O75787,O75795,O75829,O75841,O75844,O75845,O75871,O75882,O75899,O75900,O75907,O75908,O75911,O75915,O75916,O75923,O75949,O75954,O75955,O75970,O75976,O76000,O76001,O76002,O76024,O76036,O76050,O76062,O76081,O76082,O76087,O76090,O76095,O76099,O76100,O94759,O94766,O94769,O94772,O94777,O94778,O94779,O94804,O94823,O94826,O94856,O94876,O94886,O94898,O94901,O94905,O94910,O94911,O94921,O94923,O94933,O94956,O94966,O94973,O94985,O94991,O95006,O95007,O95013,O95047,O95049,O95057,O95069,O95070,O95136,O95139,O95140,O95150,O95159,O95164,O95167,O95168,O95169,O95180,O95183,O95185,O95196,O95202,O95206,O95210,O95214,O95221,O95222,O95237,O95249,O95255,O95256,O95258,O95259,O95264,O95274,O95279,O95292,O95297,O95298,O95342,O95371,O95377,O95395,O95406,O95415,O95425,O95427,O95436,O95450,O95452,O95461,O95470,O95471,O95473,O95476,O95477,O95484,O95490,O95497,O95498,O95500,O95502,O95528,O95562,O95563,O95573,O95619,O95622,O95631,O95661,O95665,O95672,O95674,O95711,O95716,O95727,O95754,O95772,O95782,O95800,O95803,O95807,O95832,O95833,O95838,O95847,O95857,O95858,O95859,O95864,O95866,O95867,O95868,O95870,O95876,O95886,O95907,O95918,O95944,O95971,O95976,O95977,O95980,O95992,O96002,O96005,O96009,O96011,O96014,O96017,O96024,P00156,P00167,P00395,P00403,P00414,P00451,P00533,P00734,P00750,P00813,P00846,P01008,P01042,P01111,P01112,P01116,P01130,P01133,P01135,P01137,P01375,P01589,P01730,P01732,P01833,P01848,P01850,P01889,P01891,P01892,P01893,P01903,P01906,P01909,P01911,P01912,P01920,P02452,P02458,P02461,P02462,P02708,P02724,P02730,P02748,P02751,P02786,P03886,P03891,P03897,P03901,P03905,P03915,P03923,P03928,P03950,P03956,P03986,P03989,P03999,P04000,P04001,P04004,P04035,P04156,P04201,P04216,P04222,P04229,P04233,P04234,P04264,P04439,P04440,P04626,P04628,P04629,P04746,P04839,P04843,P04844,P04920,P04921,P05023,P05026,P05067,P05093,P05106,P05107,P05141,P05154,P05156,P05164,P05186,P05187,P05362,P05496,P05534,P05538,P05556,P05981,P05997,P06028,P06126,P06127,P06133,P06213,P06340,P06729,P06731,P06734,P06756,P06858,P07093,P07099,P07202,P07204,P07306,P07307,P07333,P07355,P07359,P07477,P07478,P07510,P07550,P07585,P07766,P07942,P07947,P07948,P07949,P07988,P08034,P08069,P08100,P08123,P08134,P08138,P08172,P08173,P08174,P08183,P08195,P08247,P08253,P08254,P08294,P08473,P08514,P08571,P08572,P08574,P08575,P08581,P08582,P08588,P08637,P08648,P08684,P08842,P08865,P08887,P08908,P08910,P08912,P08913,P08922,P08962,P08F94,P09131,P09172,P09237,P09238,P09326,P09382,P09486,P09543,P09544,P09564,P09603,P09619,P09669,P09693,P09758,P09769,P09848,P09912,P09923,P09958,P0C091,P0C0E4,P0C2L3,P0C2S0,P0C604,P0C617,P0C623,P0C626,P0C628,P0C629,P0C645,P0C646,P0C672,P0C6S8,P0C6T2,P0C7M8,P0C7N1,P0C7N4,P0C7N5,P0C7N8,P0C7Q5,P0C7Q6,P0C7T2,P0C7T3,P0C7T8,P0C7U0,P0C7U3,P0C7U9,P0C7V7,P0C838,P0C851,P0C874,P0CF51,P0CG08,P0CG37,P0CG41,P0CK96,P0CK97,P0DI80,P0DJ07,P0DJ93,P0DKB5,P0DKB6,P0DKV0,P0DKX4,P0DL12,P0DMS8,P0DPA2,P0DPA3,P10176,P10301,P10314,P10316,P10319,P10321,P10415,P10586,P10620,P10644,P10696,P10721,P10745,P10747,P10767,P10912,P10915,P10966,P11047,P11049,P11117,P11166,P11168,P11169,P11215,P11229,P11230,P11234,P11279,P11362,P11511,P11532,P11597,P11686,P11717,P11836,P11912,P12107,P12109,P12110,P12111,P12235,P12236,P12314,P12318,P12319,P12644,P12821,P12829,P12830,P12931,P13164,P13224,P13385,P13473,P13498,P13569,P13591,P13598,P13611,P13612,P13637,P13688,P13726,P13746,P13747,P13760,P13761,P13762,P13765,P13866,P13942,P13945,P13987,P14060,P14151,P14207,P14209,P14317,P14384,P14410,P14415,P14416,P14543,P14555,P14616,P14672,P14679,P14770,P14778,P14780,P14784,P14867,P15144,P15151,P15169,P15260,P15291,P15328,P15382,P15391,P15421,P15502,P15509,P15514,P15529,P15812,P15813,P15907,P15941,P15954,P16066,P16070,P16109,P16112,P16144,P16150,P16188,P16189,P16190,P16234,P16260,P16284,P16389,P16410,P16422,P16442,P16452,P16471,P16473,P16499,P16581,P16615,P16662,P16671,P16871,P16989,P17152,P17181,P17301,P17302,P17342,P17538,P17643,P17658,P17677,P17693,P17787,P17813,P17927,P17948,P18084,P18089,P18405,P18433,P18462,P18463,P18464,P18465,P18505,P18507,P18564,P18577,P18627,P18825,P18827,P18850,P19021,P19022,P19075,P19086,P19224,P19235,P19256,P19320,P19397,P19438,P19440,P19526,P19634,P19801,P20020,P20023,P20036,P20039,P20138,P20273,P20292,P20309,P20333,P20336,P20337,P20338,P20339,P20594,P20645,P20648,P20701,P20702,P20774,P20827,P20849,P20908,P20916,P20963,P21145,P21217,P21397,P21439,P21452,P21453,P21462,P21554,P21579,P21583,P21589,P21709,P21728,P21730,P21731,P21754,P21757,P21796,P21802,P21810,P21817,P21854,P21860,P21917,P21918,P21926,P21941,P21964,P22001,P22079,P22083,P22105,P22223,P22309,P22310,P22413,P22455,P22459,P22460,P22607,P22732,P22735,P22748,P22760,P22794,P22888,P22897,P23142,P23229,P23276,P23352,P23415,P23416,P23467,P23468,P23469,P23470,P23471,P23508,P23510,P23515,P23634,P23763,P23942,P23945,P23975,P24043,P24046,P24071,P24310,P24311,P24347,P24390,P24394,P24530,P24557,P24588,P24821,P25021,P25024,P25025,P25063,P25067,P25089,P25090,P25092,P25100,P25101,P25103,P25105,P25106,P25116,P25189,P25391,P25445,P25874,P25929,P25940,P25942,P26006,P26010,P26012,P26038,P26045,P26439,P26572,P26678,P26715,P26717,P26718,P26842,P26951,P26992,P27037,P27105,P27169,P27216,P27338,P27448,P27449,P27487,P27544,P27658,P27701,P27824,P27930,P28067,P28068,P28221,P28222,P28223,P28288,P28300,P28328,P28335,P28336,P28472,P28476,P28566,P28827,P28845,P28906,P28907,P28908,P29016,P29017,P29033,P29074,P29274,P29275,P29279,P29317,P29320,P29322,P29323,P29371,P29376,P29400,P29474,P29475,P29965,P29972,P29973,P29992,P30203,P30273,P30301,P30408,P30411,P30414,P30443,P30447,P30450,P30453,P30455,P30456,P30457,P30459,P30460,P30461,P30462,P30464,P30466,P30475,P30479,P30480,P30481,P30483,P30484,P30485,P30486,P30487,P30488,P30490,P30491,P30492,P30493,P30495,P30498,P30499,P30501,P30504,P30505,P30508,P30510,P30511,P30512,P30518,P30530,P30531,P30532,P30536,P30542,P30550,P30556,P30559,P30685,P30825,P30872,P30874,P30926,P30939,P30953,P30954,P30968,P30988,P30989,P31213,P31321,P31358,P31391,P31431,P31512,P31639,P31641,P31644,P31645,P31689,P31785,P31994,P31995,P31997,P32004,P32238,P32239,P32241,P32245,P32246,P32247,P32248,P32249,P32297,P32302,P32418,P32745,P32856,P32926,P32927,P32942,P32970,P32971,P33032,P33121,P33151,P33241,P33527,P33681,P33897,P33908,P33947,P34741,P34810,P34903,P34910,P34925,P34969,P34972,P34981,P34982,P34995,P34998,P35052,P35070,P35125,P35212,P35241,P35247,P35346,P35348,P35367,P35368,P35372,P35408,P35410,P35414,P35462,P35498,P35499,P35503,P35504,P35523,P35555,P35556,P35575,P35590,P35610,P35613,P35625,P35670,P35858,P35913,P35916,P35968,P36021,P36222,P36268,P36269,P36382,P36383,P36537,P36544,P36888,P36894,P36896,P36897,P36941,P36952,P36956,P37023,P37059,P37088,P37173,P37268,P37287,P37288,P38435,P38484,P38567,P38570,P39059,P39060,P39086,P39210,P39656,P39880,P39900,P40123,P40126,P40145,P40189,P40197,P40198,P40199,P40200,P40238,P40259,P40305,P40617,P40879,P40967,P41143,P41145,P41146,P41180,P41181,P41217,P41220,P41221,P41231,P41273,P41440,P41586,P41587,P41594,P41595,P41597,P41732,P41968,P42081,P42166,P42167,P42261,P42262,P42263,P42658,P42701,P42702,P42857,P42892,P43003,P43004,P43005,P43007,P43080,P43088,P43115,P43116,P43119,P43121,P43146,P43220,P43250,P43251,P43307,P43308,P43405,P43489,P43626,P43627,P43628,P43629,P43630,P43631,P43632,P43657,P43681,P45452,P45844,P45880,P46019,P46020,P46059,P46089,P46091,P46092,P46093,P46094,P46095,P46098,P46531,P46663,P46695,P46721,P46940,P46977,P47211,P47775,P47804,P47869,P47870,P47871,P47872,P47881,P47883,P47884,P47887,P47888,P47890,P47893,P47898,P47900,P47901,P47985,P48023,P48029,P48039,P48048,P48050,P48051,P48058,P48060,P48065,P48066,P48067,P48145,P48146,P48165,P48167,P48169,P48201,P48230,P48357,P48426,P48509,P48544,P48546,P48547,P48549,P48551,P48651,P48664,P48751,P48764,P48960,P48995,P49019,P49069,P49146,P49190,P49238,P49257,P49279,P49281,P49286,P49327,P49447,P49641,P49682,P49683,P49685,P49747,P49755,P49757,P49768,P49771,P49788,P49795,P49810,P49895,P49903,P49908,P49961,P50052,P50150,P50151,P50281,P50391,P50402,P50406,P50416,P50443,P50591,P50851,P50876,P50895,P50993,P51148,P51151,P51157,P51159,P51160,P51164,P51168,P51170,P51172,P51451,P51511,P51512,P51571,P51572,P51575,P51582,P51648,P51654,P51674,P51677,P51679,P51681,P51684,P51685,P51686,P51693,P51787,P51788,P51790,P51793,P51795,P51797,P51798,P51800,P51801,P51805,P51809,P51810,P51811,P51828,P51841,P51864,P51884,P51888,P51993,P52429,P52569,P52757,P52797,P52798,P52799,P52803,P52848,P52849,P53007,P53420,P53680,P53708,P53794,P53801,P53805,P53816,P53985,P54219,P54289,P54707,P54709,P54710,P54753,P54756,P54760,P54762,P54764,P54792,P54826,P54829,P54849,P54851,P54852,P54855,P54920,P55001,P55011,P55017,P55040,P55042,P55061,P55064,P55073,P55081,P55082,P55083,P55085,P55087,P55160,P55259,P55268,P55283,P55285,P55286,P55287,P55289,P55290,P55291,P55344,P55808,P55851,P55899,P55916,P56134,P56159,P56180,P56199,P56373,P56378,P56557,P56589,P56696,P56703,P56704,P56705,P56706,P56746,P56747,P56748,P56749,P56750,P56817,P56856,P56880,P56937,P56962,P56975,P57054,P57057,P57087,P57088,P57103,P57105,P57679,P57727,P57729,P57735,P57738,P57739,P57773,P57789,P58170,P58173,P58180,P58181,P58182,P58215,P58335,P58397,P58400,P58401,P58418,P58511,P58549,P58550,P58658,P58743,P58753,P58872,P59025,P59190,P59510,P59533,P59534,P59535,P59536,P59537,P59538,P59539,P59540,P59541,P59542,P59543,P59544,P59551,P59646,P59768,P59773,P59901,P59922,P60033,P60059,P60201,P60323,P60468,P60508,P60602,P60606,P60852,P60893,P60953,P61006,P61009,P61018,P61020,P61073,P61077,P61165,P61224,P61266,P61421,P61586,P61619,P61647,P61803,P61952,P62070,P62079,P62491,P62699,P62834,P62952,P62955,P63000,P63010,P63027,P63146,P63211,P63215,P63218,P63244,P63252,P67812,P69849,P78310,P78324,P78325,P78333,P78334,P78348,P78352,P78357,P78363,P78369,P78380,P78381,P78382,P78383,P78410,P78423,P78504,P78508,P78509,P78536,P78552,P78562,P79483,P80370,P80723,P81408,P82251,P82279,P82987,P84095,P84157,P98073,P98095,P98153,P98155,P98160,P98161,P98164,P98172,P98187,P98194,P98196,P98198,Q00013,Q00325,Q00765,Q00973,Q00975,Q00LT1,Q01113,Q01118,Q01151,Q01344,Q01362,Q01453,Q01518,Q01628,Q01629,Q01638,Q01650,Q01668,Q01718,Q01726,Q01814,Q01955,Q01959,Q01970,Q01973,Q01974,Q02094,Q02127,Q02161,Q02223,Q02246,Q02297,Q02388,Q02413,Q02487,Q02505,Q02641,Q02643,Q02742,Q02763,Q02846,Q02978,Q03113,Q03167,Q03395,Q03405,Q03431,Q03518,Q03519,Q03692,Q03721,Q04609,Q04656,Q04671,Q04721,Q04771,Q04826,Q04844,Q04900,Q04912,Q04941,Q05469,Q05586,Q05707,Q05901,Q05940,Q05996,Q06055,Q06136,Q06418,Q06432,Q06481,Q06495,Q06643,Q06828,Q07000,Q07001,Q07011,Q07065,Q07075,Q07092,Q07108,Q07157,Q07326,Q07444,Q07507,Q07654,Q07699,Q07812,Q07817,Q07820,Q07837,Q07912,Q07954,Q08174,Q08289,Q08334,Q08345,Q08357,Q08397,Q08431,Q08462,Q08477,Q08554,Q08629,Q08708,Q08722,Q08828,Q08AI6,Q08ET2,Q09013,Q09160,Q09327,Q09328,Q09428,Q09470,Q0D2K0,Q0GE19,Q0JRZ9,Q0P670,Q0P6D2,Q0P6H9,Q0VAQ4,Q0VD83,Q0VDE8,Q0VDI3,Q10469,Q10471,Q10472,Q10588,Q10589,Q10981,Q11128,Q11130,Q11201,Q11203,Q11206,Q12767,Q12770,Q12772,Q12791,Q12797,Q12805,Q12809,Q12829,Q12836,Q12846,Q12860,Q12864,Q12866,Q12879,Q12884,Q12887,Q12891,Q12893,Q12907,Q12908,Q12912,Q12913,Q12918,Q12934,Q12959,Q12974,Q12981,Q12983,Q12999,Q13002,Q13003,Q13018,Q13021,Q13061,Q13093,Q13112,Q13113,Q13114,Q13145,Q13183,Q13190,Q13224,Q13241,Q13255,Q13258,Q13261,Q13277,Q13286,Q13291,Q13303,Q13304,Q13308,Q13319,Q13323,Q13324,Q13326,Q13332,Q13336,Q13349,Q13361,Q13370,Q13410,Q13421,Q13423,Q13424,Q13425,Q13433,Q13443,Q13444,Q13445,Q13449,Q13454,Q13467,Q13477,Q13478,Q13488,Q13491,Q13492,Q13505,Q13507,Q13508,Q13515,Q13520,Q13530,Q13557,Q13563,Q13564,Q13571,Q13585,Q13586,Q13591,Q13606,Q13607,Q13613,Q13621,Q13634,Q13635,Q13639,Q13641,Q13651,Q13683,Q13685,Q13698,Q13702,Q13705,Q13724,Q13733,Q13740,Q13751,Q13753,Q13797,Q13873,Q13884,Q13887,Q13936,Q13948,Q14002,Q14003,Q14028,Q14031,Q14050,Q14055,Q14088,Q14108,Q14112,Q14114,Q14118,Q14126,Q14156,Q14160,Q14162,Q14165,Q14168,Q14210,Q14242,Q14246,Q14254,Q14318,Q14330,Q14332,Q14344,Q14392,Q14416,Q14432,Q14435,Q14439,Q14494,Q14500,Q14512,Q14515,Q14517,Q14524,Q14534,Q14542,Q14571,Q14573,Q14574,Q14626,Q14627,Q14642,Q14643,Q14644,Q14654,Q14656,Q14680,Q14699,Q14703,Q14714,Q14721,Q14728,Q14739,Q14761,Q14773,Q14789,Q147U7,Q14802,Q14831,Q14832,Q14833,Q14849,Q14916,Q14940,Q14943,Q14952,Q14953,Q14954,Q14956,Q14957,Q14964,Q14973,Q14982,Q14993,Q14BN4,Q14C86,Q14C87,Q14CN2,Q14CX5,Q14CZ8,Q14D04,Q14D33,Q14DG7,Q15005,Q15011,Q15012,Q15029,Q15035,Q15041,Q15043,Q15049,Q15070,Q15077,Q15078,Q15109,Q15116,Q15125,Q15139,Q15155,Q15165,Q15166,Q15223,Q15256,Q15262,Q15286,Q15303,Q15311,Q15334,Q15363,Q15375,Q15382,Q15388,Q15391,Q15392,Q15399,Q15413,Q15438,Q15506,Q15526,Q15546,Q15582,Q15612,Q15615,Q15617,Q15619,Q15620,Q15622,Q15629,Q15700,Q15722,Q15738,Q15743,Q15751,Q15758,Q15760,Q15761,Q15762,Q15768,Q15771,Q15800,Q15822,Q15825,Q15835,Q15836,Q15842,Q15849,Q15858,Q15878,Q15884,Q15904,Q15907,Q16099,Q16206,Q16280,Q16281,Q16288,Q16322,Q16348,Q16363,Q16394,Q16445,Q16478,Q16515,Q16538,Q16549,Q16553,Q16558,Q16563,Q16570,Q16572,Q16581,Q16585,Q16586,Q16602,Q16610,Q16611,Q16617,Q16620,Q16623,Q16625,Q16635,Q16647,Q16651,Q16653,Q16655,Q16671,Q16706,Q16720,Q16739,Q16787,Q16790,Q16799,Q16819,Q16820,Q16821,Q16827,Q16832,Q16842,Q16849,Q16850,Q16853,Q16873,Q16880,Q16891,Q17R55,Q17RQ9,Q17RW2,Q18PE1,Q19T08,Q1AE95,Q1EHB4,Q1HG43,Q1HG44,Q24JP5,Q24JQ0,Q29718,Q29836,Q29865,Q29940,Q29960,Q29963,Q29974,Q29980,Q29983,Q2HXU8,Q2I0M4,Q2KHT4,Q2LD37,Q2M2E3,Q2M2I8,Q2M385,Q2M3C6,Q2M3G0,Q2M3M2,Q2M3R5,Q2M3T9,Q2MJR0,Q2PZI1,Q2QL34,Q2T9K0,Q2TAA5,Q2TAL6,Q2TBF2,Q2UY09,Q2VWP7,Q2VYF4,Q2WGJ8,Q2WGJ9,Q2Y0W8,Q30134,Q30154,Q30167,Q30201,Q31610,Q31612,Q32M45,Q32ZL2,Q3B7J2,Q3B7S5,Q3B7T3,Q3C1V0,Q3KNS1,Q3KNT9,Q3KNW5,Q3KPI0,Q3KQZ1,Q3KR37,Q3MIP1,Q3MIR4,Q3MIW9,Q3MIX3,Q3MUY2,Q3SXP7,Q3SXY7,Q3SY17,Q3SY77,Q3SYC2,Q3T906,Q3V5L5,Q3V6T2,Q3YBM2,Q3ZAQ7,Q3ZCQ3,Q3ZCQ8,Q401N2,Q494W8,Q495A1,Q495M3,Q495N2,Q495T6,Q495W5,Q496F6,Q496H8,Q496J9,Q49A17,Q49SQ1,Q4ADV7,Q4G0I0,Q4G0N0,Q4G0N8,Q4G0S4,Q4G148,Q4G1C9,Q4KMG0,Q4KMG9,Q4KMQ2,Q4KMZ8,Q4LDR2,Q4U2R8,Q4V9L6,Q4VC39,Q4VNC0,Q4VNC1,Q4VXA5,Q4VXF1,Q4ZG55,Q4ZIN3,Q4ZJI4,Q504Y0,Q50LG9,Q52LC2,Q52LD8,Q53EL9,Q53EP0,Q53EU6,Q53F39,Q53FP2,Q53FV1,Q53GD3,Q53GL0,Q53GQ0,Q53HI1,Q53R12,Q53RD9,Q53RT3,Q53RY4,Q53S58,Q53TN4,Q567V2,Q587I9,Q58DX5,Q58EX2,Q58HT5,Q5BIV9,Q5BJD5,Q5BJF2,Q5BJH2,Q5BJH7,Q5BKT4,Q5BKX6,Q5BVD1,Q5DID0,Q5DX21,Q5EB52,Q5FWE3,Q5FYA8,Q5GH70,Q5GH72,Q5GH73,Q5GH76,Q5GH77,Q5H8A4,Q5H8C1,Q5H943,Q5H9E4,Q5H9R4,Q5H9S7,Q5HYA8,Q5HYC2,Q5HYJ1,Q5HYL7,Q5I7T1,Q5IJ48,Q5J8M3,Q5J8X5,Q5JPE7,Q5JQS5,Q5JRA6,Q5JRM2,Q5JRS4,Q5JRV8,Q5JTH9,Q5JTV8,Q5JUK3,Q5JW98,Q5JX69,Q5JX71,Q5JXA9,Q5JXX5,Q5JXX7,Q5JZY3,Q5K4E3,Q5K4L6,Q5KU26,Q5M7Z0,Q5M8T2,Q5MY95,Q5NUL3,Q5PT55,Q5QFB9,Q5QGT7,Q5QGZ9,Q5QJU3,Q5R3F8,Q5R3K3,Q5RGS3,Q5RI15,Q5SGD2,Q5SNT2,Q5SQ64,Q5SR56,Q5SRD1,Q5SRI9,Q5SRN2,Q5SSG8,Q5SV17,Q5SVS4,Q5SWH9,Q5SWX8,Q5SY80,Q5SZI1,Q5SZK8,Q5T0T0,Q5T197,Q5T1A1,Q5T1C6,Q5T1Q4,Q5T1S8,Q5T292,Q5T2D2,Q5T2E6,Q5T2T1,Q5T3F8,Q5T3U5,Q5T442,Q5T4D3,Q5T4F4,Q5T4J0,Q5T4S7,Q5T4T1,Q5T5W8,Q5T601,Q5T6L9,Q5T6X4,Q5T6X5,Q5T700,Q5T7M9,Q5T7P6,Q5T7P8,Q5T7R7,Q5T848,Q5T8D3,Q5T9L3,Q5T9Z0,Q5TAH2,Q5TAT6,Q5TCQ9,Q5TEA6,Q5TF21,Q5TF39,Q5TFQ8,Q5TGI0,Q5TGU0,Q5TGY1,Q5TGZ0,Q5TH69,Q5TZ20,Q5TZJ5,Q5U3C3,Q5U4P2,Q5UAW9,Q5UCC4,Q5VSG8,Q5VT66,Q5VT99,Q5VTJ3,Q5VTT2,Q5VTY9,Q5VU36,Q5VU65,Q5VU97,Q5VUB5,Q5VUD6,Q5VUY2,Q5VV42,Q5VV43,Q5VV63,Q5VVB8,Q5VVP1,Q5VW32,Q5VW36,Q5VW38,Q5VWC8,Q5VWK5,Q5VX71,Q5VXT5,Q5VXU1,Q5VXU3,Q5VY43,Q5VY80,Q5VYJ5,Q5VYP0,Q5VZ72,Q5VZI3,Q5VZR4,Q5VZY2,Q5W0B7,Q5W0N0,Q5W0Z9,Q5XG99,Q5XXA6,Q5Y7A7,Q5ZPR3,Q60682,Q629K1,Q63HM2,Q63HQ2,Q63ZE4,Q643R3,Q658N2,Q658P3,Q66K14,Q66K66,Q66K79,Q674R7,Q67FW5,Q685J3,Q687X5,Q68CJ9,Q68CP4,Q68CQ1,Q68CQ7,Q68CR1,Q68CR7,Q68D42,Q68D85,Q68DH5,Q68DV7,Q68EM7,Q68G75,Q69383,Q69384,Q695T7,Q69YG0,Q69YW2,Q69YZ2,Q6AI14,Q6AZY7,Q6BAA4,Q6DD88,Q6DKI7,Q6DN12,Q6DN14,Q6DN72,Q6DWJ6,Q6E213,Q6EIG7,Q6EMK4,Q6GMR7,Q6GPH6,Q6GPI1,Q6GTX8,Q6GV28,Q6H3X3,Q6IA17,Q6IAN0,Q6IC98,Q6ICH7,Q6ICI0,Q6ICL7,Q6IEE7,Q6IEE8,Q6IEU7,Q6IEV9,Q6IEY1,Q6IEZ7,Q6IF00,Q6IF36,Q6IF42,Q6IF63,Q6IF82,Q6IF99,Q6IFG1,Q6IFH4,Q6IFN5,Q6IQ20,Q6IS24,Q6ISU1,Q6IWH7,Q6J4K2,Q6J9G0,Q6KCM7,Q6L9W6,Q6MZM0,Q6N022,Q6N075,Q6NSJ0,Q6NSJ5,Q6NT16,Q6NTF9,Q6NUI2,Q6NUI6,Q6NUJ2,Q6NUK1,Q6NUK4,Q6NUQ4,Q6NUS6,Q6NUS8,Q6NUT2,Q6NUT3,Q6NV75,Q6NVV3,Q6NW34,Q6NW40,Q6NXN4,Q6NXT4,Q6NXT6,Q6NZ63,Q6NZI2,Q6P0A1,Q6P179,Q6P1A2,Q6P1J6,Q6P1K1,Q6P1M0,Q6P1Q0,Q6P1S2,Q6P2H8,Q6P499,Q6P4A7,Q6P4E1,Q6P4F1,Q6P4H8,Q6P4Q7,Q6P531,Q6P5S7,Q6P5W5,Q6P5X7,Q6P7N7,Q6P995,Q6P9A2,Q6P9B9,Q6P9F7,Q6P9G4,Q6PCB0,Q6PCB7,Q6PCB8,Q6PEX7,Q6PEY0,Q6PEY1,Q6PEZ8,Q6PHW0,Q6PI25,Q6PI73,Q6PI78,Q6PIL6,Q6PIS1,Q6PIU1,Q6PIU2,Q6PIV7,Q6PIZ9,Q6PJF5,Q6PJG9,Q6PJW8,Q6PK18,Q6PKC3,Q6PL45,Q6PML9,Q6PP77,Q6PRD1,Q6PXP3,Q6Q0C1,Q6Q4G3,Q6Q8B3,Q6QAJ8,Q6QHC5,Q6QNK2,Q6RW13,Q6T423,Q6T4P5,Q6TCH4,Q6TCH7,Q6U736,Q6U841,Q6UE05,Q6UQ28,Q6UVK1,Q6UVM3,Q6UVW9,Q6UVY6,Q6UW02,Q6UW56,Q6UW60,Q6UW68,Q6UW88,Q6UWB1,Q6UWB4,Q6UWD8,Q6UWF3,Q6UWH4,Q6UWH6,Q6UWI2,Q6UWI4,Q6UWJ1,Q6UWJ8,Q6UWL2,Q6UWL6,Q6UWM7,Q6UWM9,Q6UWN0,Q6UWN5,Q6UWP7,Q6UWR7,Q6UWT2,Q6UWU4,Q6UWV2,Q6UWV6,Q6UWV7,Q6UWW9,Q6UX01,Q6UX06,Q6UX15,Q6UX27,Q6UX34,Q6UX40,Q6UX41,Q6UX65,Q6UX68,Q6UX71,Q6UX72,Q6UX82,Q6UX98,Q6UXA7,Q6UXB3,Q6UXB4,Q6UXB8,Q6UXC1,Q6UXD1,Q6UXD5,Q6UXD7,Q6UXE8,Q6UXF1,Q6UXG2,Q6UXG3,Q6UXG8,Q6UXI7,Q6UXI9,Q6UXK2,Q6UXK5,Q6UXL0,Q6UXM1,Q6UXN7,Q6UXN8,Q6UXP3,Q6UXU4,Q6UXU6,Q6UXV0,Q6UXV1,Q6UXY1,Q6UXY8,Q6UXZ0,Q6UXZ3,Q6UXZ4,Q6UY09,Q6UY11,Q6UY14,Q6UY18,Q6V0I7,Q6V0L0,Q6V1P9,Q6W3E5,Q6W5P4,Q6XPR3,Q6XPS3,Q6XR72,Q6XYQ8,Q6XZB0,Q6Y1H2,Q6Y288,Q6Y2X3,Q6YBV0,Q6YHK3,Q6YI46,Q6ZMB0,Q6ZMB5,Q6ZMC9,Q6ZMD2,Q6ZMG9,Q6ZMH5,Q6ZMI3,Q6ZMJ2,Q6ZMP0,Q6ZMQ8,Q6ZMR5,Q6ZMU5,Q6ZMZ0,Q6ZMZ3,Q6ZN44,Q6ZN68,Q6ZNA5,Q6ZNB6,Q6ZNB7,Q6ZNC8,Q6ZNI0,Q6ZNR0,Q6ZP29,Q6ZP80,Q6ZPB5,Q6ZPD8,Q6ZPD9,Q6ZQN7,Q6ZQQ2,Q6ZRH7,Q6ZRP7,Q6ZRR5,Q6ZS10,Q6ZS62,Q6ZS81,Q6ZS82,Q6ZSA7,Q6ZSJ9,Q6ZSM3,Q6ZSS7,Q6ZSY5,Q6ZT12,Q6ZT21,Q6ZT89,Q6ZTQ4,Q6ZTR5,Q6ZU45,Q6ZU64,Q6ZU69,Q6ZUB0,Q6ZUB1,Q6ZUK4,Q6ZUT9,Q6ZUX7,Q6ZV29,Q6ZVE7,Q6ZVK1,Q6ZVL6,Q6ZVN8,Q6ZVX9,Q6ZW05,Q6ZWK4,Q6ZWK6,Q6ZWL3,Q6ZWT7,Q6ZXV5,Q70CQ3,Q70E73,Q70HW3,Q70JA7,Q70SY1,Q70UQ0,Q70Z44,Q71H61,Q71RC9,Q71RG4,Q71RH2,Q71RS6,Q75N90,Q75NE6,Q75T13,Q75V66,Q76EJ3,Q76KD6,Q76KP1,Q76M96,Q76MJ5,Q7KYR7,Q7KZI7,Q7KZN9,Q7L0J3,Q7L0Q8,Q7L0X0,Q7L1I2,Q7L1S5,Q7L1W4,Q7L211,Q7L311,Q7L4E1,Q7L4S7,Q7L513,Q7L5A8,Q7L5L3,Q7L5N7,Q7L804,Q7L8C5,Q7L985,Q7LBE3,Q7LDI9,Q7LFX5,Q7LGA3,Q7LGC8,Q7RTM1,Q7RTP0,Q7RTR8,Q7RTS5,Q7RTS6,Q7RTT9,Q7RTX0,Q7RTX1,Q7RTX7,Q7RTX9,Q7RTY0,Q7RTY1,Q7RTY8,Q7RTY9,Q7Z2D5,Q7Z2H8,Q7Z2K6,Q7Z2K8,Q7Z2Q7,Q7Z2W7,Q7Z304,Q7Z388,Q7Z3B0,Q7Z3B1,Q7Z3C6,Q7Z3D4,Q7Z3F1,Q7Z3J2,Q7Z3Q1,Q7Z3S7,Q7Z3T1,Q7Z402,Q7Z403,Q7Z404,Q7Z407,Q7Z408,Q7Z410,Q7Z412,Q7Z418,Q7Z419,Q7Z429,Q7Z434,Q7Z442,Q7Z443,Q7Z444,Q7Z449,Q7Z460,Q7Z4F1,Q7Z4J2,Q7Z4L0,Q7Z4N2,Q7Z4T8,Q7Z4W1,Q7Z553,Q7Z5A7,Q7Z5B4,Q7Z5H4,Q7Z5H5,Q7Z5J8,Q7Z5L7,Q7Z5M5,Q7Z5N4,Q7Z5R6,Q7Z5S9,Q7Z601,Q7Z602,Q7Z614,Q7Z692,Q7Z695,Q7Z698,Q7Z699,Q7Z6A9,Q7Z6B0,Q7Z6J6,Q7Z6L0,Q7Z6M3,Q7Z6P3,Q7Z6W1,Q7Z739,Q7Z769,Q7Z7B1,Q7Z7D3,Q7Z7G2,Q7Z7H5,Q7Z7J7,Q7Z7M0,Q7Z7M1,Q7Z7M8,Q7Z7M9,Q7Z7N9,Q86SF2,Q86SJ2,Q86SJ6,Q86SK9,Q86SM5,Q86SM8,Q86SP6,Q86SQ3,Q86SQ4,Q86SQ6,Q86SR1,Q86SS6,Q86SU0,Q86T03,Q86T13,Q86T26,Q86T96,Q86TG1,Q86TL2,Q86TM6,Q86TY3,Q86U02,Q86U90,Q86UB9,Q86UD3,Q86UD5,Q86UE4,Q86UE6,Q86UF1,Q86UF2,Q86UG4,Q86UK0,Q86UK5,Q86UL3,Q86UN2,Q86UN3,Q86UP0,Q86UP2,Q86UP6,Q86UP9,Q86UQ4,Q86UQ5,Q86UR5,Q86UT5,Q86UW1,Q86UW2,Q86UW8,Q86V24,Q86V35,Q86V40,Q86V85,Q86VB7,Q86VD7,Q86VD9,Q86VE9,Q86VF5,Q86VH4,Q86VH5,Q86VI4,Q86VL8,Q86VR2,Q86VR7,Q86VU5,Q86VW1,Q86VW2,Q86VY9,Q86VZ1,Q86VZ4,Q86VZ5,Q86W10,Q86W33,Q86W47,Q86W74,Q86WA9,Q86WB7,Q86WC4,Q86WI0,Q86WI1,Q86WK6,Q86WK7,Q86WK9,Q86WS5,Q86WV6,Q86X19,Q86X29,Q86X52,Q86XA0,Q86XE3,Q86XJ0,Q86XK7,Q86XL3,Q86XM0,Q86XQ3,Q86XR5,Q86XS8,Q86XT9,Q86XX4,Q86Y07,Q86Y22,Q86Y34,Q86Y38,Q86Y39,Q86Y82,Q86YC3,Q86YD1,Q86YD3,Q86YD5,Q86YJ5,Q86YL7,Q86YN1,Q86YR6,Q86YT5,Q86YT9,Q86YW5,Q86Z14,Q8IU57,Q8IU68,Q8IU80,Q8IU89,Q8IU99,Q8IUA7,Q8IUC8,Q8IUH4,Q8IUH5,Q8IUH8,Q8IUK5,Q8IUL8,Q8IUN9,Q8IUR5,Q8IUS5,Q8IUW5,Q8IUX1,Q8IUX8,Q8IUY3,Q8IV01,Q8IV08,Q8IV16,Q8IV31,Q8IV45,Q8IV77,Q8IVB4,Q8IVI9,Q8IVJ1,Q8IVJ8,Q8IVM8,Q8IVN8,Q8IVP5,Q8IVQ6,Q8IVU1,Q8IVV8,Q8IVW8,Q8IVY1,Q8IW00,Q8IW52,Q8IW70,Q8IWA4,Q8IWA5,Q8IWB1,Q8IWB4,Q8IWB9,Q8IWD5,Q8IWK6,Q8IWL1,Q8IWL2,Q8IWR1,Q8IWT1,Q8IWT6,Q8IWU2,Q8IWU4,Q8IWV1,Q8IWV2,Q8IWX5,Q8IWY9,Q8IX05,Q8IX19,Q8IX94,Q8IXA5,Q8IXB3,Q8IXE1,Q8IXF9,Q8IXH8,Q8IXI1,Q8IXI2,Q8IXK2,Q8IXM6,Q8IXS6,Q8IXU6,Q8IXX5,Q8IY17,Q8IY26,Q8IY34,Q8IY49,Q8IY50,Q8IY95,Q8IYB5,Q8IYJ0,Q8IYJ3,Q8IYK8,Q8IYL9,Q8IYP9,Q8IYR6,Q8IYS0,Q8IYS2,Q8IYT2,Q8IYV9,Q8IZ08,Q8IZ52,Q8IZ57,Q8IZ96,Q8IZA0,Q8IZC6,Q8IZD6,Q8IZF0,Q8IZF2,Q8IZF3,Q8IZF4,Q8IZF5,Q8IZF6,Q8IZF7,Q8IZJ1,Q8IZJ4,Q8IZK6,Q8IZM9,Q8IZN3,Q8IZP7,Q8IZP9,Q8IZQ1,Q8IZR5,Q8IZS7,Q8IZS8,Q8IZT8,Q8IZU8,Q8IZU9,Q8IZV2,Q8IZV5,Q8IZY2,Q8J025,Q8MH63,Q8N0U2,Q8N0U8,Q8N0V5,Q8N0W4,Q8N0W7,Q8N0Y3,Q8N0Y5,Q8N0Z9,Q8N109,Q8N111,Q8N112,Q8N114,Q8N118,Q8N126,Q8N127,Q8N130,Q8N131,Q8N138,Q8N139,Q8N144,Q8N146,Q8N148,Q8N149,Q8N158,Q8N162,Q8N1C3,Q8N1L4,Q8N1M1,Q8N1N0,Q8N1N2,Q8N1S5,Q8N201,Q8N205,Q8N271,Q8N292,Q8N2A8,Q8N2C7,Q8N2F6,Q8N2G4,Q8N2H4,Q8N2K0,Q8N2K1,Q8N2M4,Q8N2Q7,Q8N2S1,Q8N2U0,Q8N2U9,Q8N326,Q8N349,Q8N350,Q8N357,Q8N370,Q8N386,Q8N387,Q8N394,Q8N3E9,Q8N3F9,Q8N3G9,Q8N3J6,Q8N3R9,Q8N3T1,Q8N3T6,Q8N3Y3,Q8N3Y7,Q8N413,Q8N414,Q8N423,Q8N428,Q8N434,Q8N441,Q8N468,Q8N490,Q8N4A0,Q8N4C7,Q8N4C9,Q8N4F4,Q8N4F7,Q8N4H5,Q8N4K4,Q8N4L1,Q8N4L2,Q8N4L4,Q8N4M1,Q8N4S7,Q8N4S9,Q8N4T0,Q8N4T4,Q8N4V1,Q8N4V2,Q8N4W6,Q8N4Z0,Q8N511,Q8N539,Q8N5B7,Q8N5C1,Q8N5D6,Q8N5G0,Q8N5G2,Q8N5K1,Q8N5M9,Q8N5S1,Q8N5U1,Q8N5Y8,Q8N608,Q8N609,Q8N614,Q8N628,Q8N661,Q8N682,Q8N695,Q8N697,Q8N6C5,Q8N6D2,Q8N6F1,Q8N6G5,Q8N6G6,Q8N6I4,Q8N6K0,Q8N6L0,Q8N6L1,Q8N6L7,Q8N6M3,Q8N6P7,Q8N6Q1,Q8N6Q3,Q8N6R1,Q8N6S5,Q8N6U8,Q8N6Y1,Q8N6Y2,Q8N743,Q8N755,Q8N766,Q8N7C0,Q8N7C4,Q8N7C7,Q8N7J2,Q8N7P1,Q8N7P3,Q8N7S2,Q8N7S6,Q8N7X8,Q8N808,Q8N816,Q8N8D7,Q8N8F6,Q8N8F7,Q8N8J7,Q8N8N0,Q8N8Q1,Q8N8Q3,Q8N8Q8,Q8N8Q9,Q8N8R3,Q8N8V2,Q8N8V8,Q8N8Z6,Q8N912,Q8N944,Q8N966,Q8N967,Q8N9A8,Q8N9F0,Q8N9F7,Q8N9I0,Q8N9I5,Q8N9M5,Q8N9R8,Q8N9X5,Q8NA29,Q8NA58,Q8NAC3,Q8NAN2,Q8NAU1,Q8NB49,Q8NB59,Q8NBD8,Q8NBF6,Q8NBI2,Q8NBI5,Q8NBI6,Q8NBJ4,Q8NBJ9,Q8NBL3,Q8NBM4,Q8NBN3,Q8NBP5,Q8NBQ7,Q8NBR0,Q8NBS3,Q8NBT3,Q8NBV4,Q8NBV8,Q8NBW4,Q8NBZ7,Q8NC01,Q8NC24,Q8NC42,Q8NC44,Q8NC54,Q8NC56,Q8NC67,Q8NCB2,Q8NCC5,Q8NCG5,Q8NCG7,Q8NCH0,Q8NCK7,Q8NCL4,Q8NCL8,Q8NCL9,Q8NCM2,Q8NCQ3,Q8NCR0,Q8NCR9,Q8NCS4,Q8NCS7,Q8NCU8,Q8NCW0,Q8NCW6,Q8ND61,Q8ND76,Q8ND94,Q8NDA2,Q8NDB6,Q8NDV1,Q8NDV2,Q8NDX1,Q8NDX2,Q8NDY8,Q8NDZ6,Q8NE00,Q8NE01,Q8NE79,Q8NE86,Q8NEA5,Q8NEB5,Q8NEC5,Q8NEQ5,Q8NER1,Q8NER5,Q8NES3,Q8NET5,Q8NET6,Q8NET8,Q8NEW0,Q8NEW7,Q8NF37,Q8NF91,Q8NFA0,Q8NFA2,Q8NFB2,Q8NFF2,Q8NFJ5,Q8NFJ6,Q8NFK1,Q8NFL0,Q8NFM4,Q8NFM7,Q8NFN8,Q8NFP4,Q8NFQ8,Q8NFR3,Q8NFR9,Q8NFT2,Q8NFT8,Q8NFU0,Q8NFU1,Q8NFW1,Q8NFY4,Q8NFZ3,Q8NFZ4,Q8NFZ6,Q8NFZ8,Q8NG04,Q8NG11,Q8NG75,Q8NG76,Q8NG77,Q8NG78,Q8NG80,Q8NG81,Q8NG83,Q8NG84,Q8NG85,Q8NG92,Q8NG94,Q8NG95,Q8NG97,Q8NG98,Q8NG99,Q8NGA0,Q8NGA1,Q8NGA2,Q8NGA5,Q8NGA6,Q8NGA8,Q8NGB2,Q8NGB4,Q8NGB6,Q8NGB8,Q8NGB9,Q8NGC0,Q8NGC1,Q8NGC2,Q8NGC3,Q8NGC4,Q8NGC5,Q8NGC6,Q8NGC7,Q8NGC8,Q8NGC9,Q8NGD0,Q8NGD1,Q8NGD2,Q8NGD3,Q8NGD4,Q8NGD5,Q8NGE0,Q8NGE1,Q8NGE2,Q8NGE3,Q8NGE5,Q8NGE7,Q8NGE8,Q8NGE9,Q8NGF0,Q8NGF1,Q8NGF3,Q8NGF4,Q8NGF6,Q8NGF7,Q8NGF8,Q8NGF9,Q8NGG0,Q8NGG1,Q8NGG2,Q8NGG3,Q8NGG4,Q8NGG5,Q8NGG6,Q8NGG7,Q8NGG8,Q8NGH3,Q8NGH5,Q8NGH6,Q8NGH7,Q8NGH8,Q8NGH9,Q8NGI0,Q8NGI1,Q8NGI2,Q8NGI3,Q8NGI4,Q8NGI6,Q8NGI7,Q8NGI8,Q8NGI9,Q8NGJ0,Q8NGJ1,Q8NGJ2,Q8NGJ3,Q8NGJ4,Q8NGJ5,Q8NGJ6,Q8NGJ7,Q8NGJ8,Q8NGJ9,Q8NGK0,Q8NGK1,Q8NGK2,Q8NGK3,Q8NGK4,Q8NGK5,Q8NGK6,Q8NGK9,Q8NGL0,Q8NGL1,Q8NGL2,Q8NGL3,Q8NGL4,Q8NGL6,Q8NGL7,Q8NGL9,Q8NGM1,Q8NGM8,Q8NGM9,Q8NGN0,Q8NGN1,Q8NGN2,Q8NGN3,Q8NGN4,Q8NGN5,Q8NGN6,Q8NGN7,Q8NGN8,Q8NGP0,Q8NGP2,Q8NGP3,Q8NGP4,Q8NGP6,Q8NGP8,Q8NGP9,Q8NGQ1,Q8NGQ2,Q8NGQ3,Q8NGQ4,Q8NGQ5,Q8NGQ6,Q8NGR1,Q8NGR2,Q8NGR3,Q8NGR4,Q8NGR5,Q8NGR6,Q8NGR8,Q8NGR9,Q8NGS0,Q8NGS1,Q8NGS2,Q8NGS3,Q8NGS4,Q8NGS5,Q8NGS6,Q8NGS7,Q8NGS8,Q8NGS9,Q8NGT0,Q8NGT1,Q8NGT2,Q8NGT5,Q8NGT7,Q8NGT9,Q8NGU1,Q8NGU2,Q8NGU4,Q8NGU9,Q8NGV0,Q8NGV5,Q8NGV6,Q8NGV7,Q8NGW1,Q8NGW6,Q8NGX0,Q8NGX1,Q8NGX2,Q8NGX3,Q8NGX5,Q8NGX6,Q8NGX8,Q8NGX9,Q8NGY0,Q8NGY1,Q8NGY2,Q8NGY3,Q8NGY5,Q8NGY6,Q8NGY7,Q8NGY9,Q8NGZ0,Q8NGZ2,Q8NGZ3,Q8NGZ4,Q8NGZ5,Q8NGZ6,Q8NGZ9,Q8NH00,Q8NH01,Q8NH02,Q8NH03,Q8NH04,Q8NH05,Q8NH06,Q8NH07,Q8NH08,Q8NH09,Q8NH10,Q8NH16,Q8NH18,Q8NH19,Q8NH21,Q8NH37,Q8NH40,Q8NH41,Q8NH42,Q8NH43,Q8NH48,Q8NH49,Q8NH50,Q8NH51,Q8NH53,Q8NH54,Q8NH55,Q8NH56,Q8NH57,Q8NH59,Q8NH60,Q8NH61,Q8NH63,Q8NH64,Q8NH67,Q8NH69,Q8NH70,Q8NH72,Q8NH73,Q8NH74,Q8NH76,Q8NH79,Q8NH80,Q8NH81,Q8NH83,Q8NH85,Q8NH87,Q8NH89,Q8NH90,Q8NH92,Q8NH93,Q8NH94,Q8NHA4,Q8NHA6,Q8NHA8,Q8NHB1,Q8NHB7,Q8NHB8,Q8NHC4,Q8NHC5,Q8NHC6,Q8NHC7,Q8NHC8,Q8NHE4,Q8NHH9,Q8NHJ6,Q8NHK3,Q8NHL6,Q8NHP6,Q8NHS1,Q8NHS3,Q8NHU3,Q8NHX9,Q8NHY0,Q8NI17,Q8NI28,Q8NI32,Q8NI35,Q8TAA9,Q8TAB3,Q8TAC9,Q8TAD4,Q8TAE7,Q8TAF8,Q8TAI7,Q8TAQ9,Q8TAV3,Q8TAV4,Q8TAZ6,Q8TB36,Q8TB61,Q8TB68,Q8TB96,Q8TBA6,Q8TBB6,Q8TBE1,Q8TBE3,Q8TBE7,Q8TBF5,Q8TBG9,Q8TBJ4,Q8TBM7,Q8TBM8,Q8TBP5,Q8TBP6,Q8TBQ9,Q8TBR7,Q8TC12,Q8TC26,Q8TC27,Q8TC36,Q8TC41,Q8TC92,Q8TCB6,Q8TCC7,Q8TCG1,Q8TCG5,Q8TCJ2,Q8TCP9,Q8TCQ1,Q8TCT6,Q8TCT7,Q8TCT8,Q8TCT9,Q8TCU3,Q8TCU5,Q8TCW7,Q8TCW9,Q8TCY5,Q8TCZ2,Q8TD07,Q8TD20,Q8TD22,Q8TD43,Q8TD46,Q8TD84,Q8TDB4,Q8TDB8,Q8TDD5,Q8TDF5,Q8TDI7,Q8TDI8,Q8TDM5,Q8TDN1,Q8TDN2,Q8TDN7,Q8TDQ0,Q8TDQ1,Q8TDS4,Q8TDS5,Q8TDS7,Q8TDT2,Q8TDU5,Q8TDU6,Q8TDU9,Q8TDV0,Q8TDV2,Q8TDV5,Q8TDW0,Q8TDW4,Q8TDW5,Q8TDW7,Q8TDX6,Q8TDX9,Q8TDY8,Q8TE23,Q8TE54,Q8TE56,Q8TE57,Q8TE58,Q8TE59,Q8TE60,Q8TEB7,Q8TEB9,Q8TED1,Q8TED4,Q8TEF2,Q8TEM1,Q8TEQ8,Q8TEY5,Q8TEZ7,Q8TF08,Q8TF62,Q8TF66,Q8TF71,Q8WTQ7,Q8WTR4,Q8WTT0,Q8WTV0,Q8WTX9,Q8WU17,Q8WU67,Q8WUD1,Q8WUD6,Q8WUG5,Q8WUH6,Q8WUM9,Q8WUS8,Q8WUT4,Q8WUT9,Q8WUU8,Q8WUX1,Q8WUY8,Q8WV15,Q8WV19,Q8WV48,Q8WV83,Q8WVE6,Q8WVE7,Q8WVF2,Q8WVH0,Q8WVI0,Q8WVN6,Q8WVP7,Q8WVQ1,Q8WVV5,Q8WVX3,Q8WVX9,Q8WVZ1,Q8WVZ7,Q8WW22,Q8WW34,Q8WW43,Q8WW52,Q8WW62,Q8WWA0,Q8WWA1,Q8WWB7,Q8WWF3,Q8WWF5,Q8WWG1,Q8WWG9,Q8WWI5,Q8WWM7,Q8WWP7,Q8WWQ0,Q8WWQ2,Q8WWQ8,Q8WWR8,Q8WWT9,Q8WWU5,Q8WWV6,Q8WWX8,Q8WWZ4,Q8WWZ7,Q8WXA8,Q8WXB1,Q8WXD0,Q8WXF7,Q8WXG6,Q8WXG9,Q8WXH0,Q8WXH2,Q8WXH6,Q8WXI7,Q8WXI8,Q8WXS4,Q8WXS5,Q8WXS8,Q8WY07,Q8WY21,Q8WY22,Q8WY41,Q8WY98,Q8WYK1,Q8WZ04,Q8WZ33,Q8WZ55,Q8WZ59,Q8WZ71,Q8WZ84,Q8WZ92,Q8WZ94,Q8WZA1,Q8WZA6,Q92185,Q92186,Q92187,Q92478,Q92482,Q92504,Q92508,Q92521,Q92523,Q92535,Q92536,Q92537,Q92542,Q92544,Q92545,Q92563,Q92581,Q92604,Q92611,Q92617,Q92626,Q92629,Q92633,Q92637,Q92643,Q92673,Q92681,Q92685,Q92692,Q92729,Q92730,Q92736,Q92737,Q92752,Q92781,Q92806,Q92813,Q92819,Q92820,Q92823,Q92824,Q92832,Q92838,Q92839,Q92847,Q92854,Q92859,Q92887,Q92896,Q92903,Q92911,Q92928,Q92930,Q92932,Q92935,Q92952,Q92953,Q92956,Q92959,Q92963,Q92968,Q92982,Q93033,Q93038,Q93050,Q93063,Q93070,Q93084,Q93086,Q93096,Q93097,Q93098,Q93100,Q95365,Q95460,Q95604,Q95IE3,Q969E2,Q969F0,Q969F2,Q969F8,Q969G9,Q969I6,Q969K7,Q969L2,Q969M2,Q969M3,Q969N2,Q969N4,Q969P0,Q969R2,Q969S0,Q969S6,Q969V1,Q969V3,Q969V5,Q969V6,Q969W0,Q969W1,Q969W9,Q969X1,Q969X2,Q969X5,Q969Z4,Q96A11,Q96A25,Q96A26,Q96A28,Q96A29,Q96A33,Q96A46,Q96A54,Q96A57,Q96A59,Q96A83,Q96A84,Q96AA3,Q96AD5,Q96AG3,Q96AG4,Q96AJ9,Q96AM1,Q96AN5,Q96AP7,Q96AQ2,Q96AQ8,Q96AW1,Q96AZ1,Q96B21,Q96B33,Q96B42,Q96B77,Q96B86,Q96B96,Q96BA8,Q96BD0,Q96BF3,Q96BI1,Q96BI3,Q96BM0,Q96BY9,Q96BZ4,Q96BZ9,Q96C03,Q96C19,Q96C24,Q96CC6,Q96CE8,Q96CG8,Q96CH1,Q96CP6,Q96CP7,Q96CQ1,Q96CU9,Q96CW1,Q96CW9,Q96D05,Q96D21,Q96D31,Q96D42,Q96D53,Q96D59,Q96D71,Q96D96,Q96DA2,Q96DA6,Q96DB9,Q96DC7,Q96DD7,Q96DL1,Q96DS6,Q96DU3,Q96DW6,Q96DX8,Q96DZ7,Q96DZ9,Q96E16,Q96E17,Q96E22,Q96E52,Q96E66,Q96E93,Q96EC8,Q96EP9,Q96ER9,Q96ES6,Q96ET8,Q96EU7,Q96EX1,Q96EX2,Q96F05,Q96F15,Q96F25,Q96F46,Q96F81,Q96FB5,Q96FC9,Q96FE5,Q96FE7,Q96FL8,Q96FL9,Q96FM1,Q96FT7,Q96FV3,Q96FX8,Q96FZ5,Q96G23,Q96G79,Q96G91,Q96G97,Q96GC9,Q96GE9,Q96GF1,Q96GL9,Q96GM1,Q96GP6,Q96GQ5,Q96GR4,Q96GW7,Q96GX1,Q96GZ6,Q96H15,Q96H72,Q96H78,Q96H96,Q96HA1,Q96HA4,Q96HA9,Q96HD1,Q96HE8,Q96HG1,Q96HH4,Q96HH6,Q96HJ5,Q96HP8,Q96HR9,Q96HS1,Q96HU8,Q96HV5,Q96I34,Q96I36,Q96I45,Q96I82,Q96IK0,Q96IQ7,Q96IV6,Q96IW7,Q96IX5,Q96IZ2,Q96J02,Q96J42,Q96J65,Q96J66,Q96J84,Q96J86,Q96JA1,Q96JA4,Q96JB6,Q96JF0,Q96JI7,Q96JJ6,Q96JJ7,Q96JN2,Q96JP9,Q96JQ0,Q96JQ2,Q96JQ5,Q96JT2,Q96JW4,Q96JX3,Q96K12,Q96K19,Q96K37,Q96K49,Q96K78,Q96KA5,Q96KC8,Q96KF7,Q96KG7,Q96KJ4,Q96KK3,Q96KK4,Q96KN9,Q96KR6,Q96KT7,Q96KV6,Q96KW2,Q96L08,Q96L33,Q96L42,Q96L58,Q96LA5,Q96LA6,Q96LA9,Q96LB0,Q96LB1,Q96LB2,Q96LC7,Q96LD1,Q96LL3,Q96LR9,Q96LT4,Q96LU7,Q96LZ7,Q96M19,Q96MC6,Q96MH6,Q96MM7,Q96MP8,Q96MS0,Q96MT1,Q96MU8,Q96MV1,Q96MV8,Q96MX0,Q96N19,Q96N35,Q96N66,Q96N68,Q96N87,Q96NA8,Q96NB2,Q96ND0,Q96NF6,Q96NI6,Q96NL1,Q96NR3,Q96NT5,Q96NU0,Q96NY7,Q96NY8,Q96P31,Q96P44,Q96P56,Q96P65,Q96P66,Q96P67,Q96P68,Q96P69,Q96P88,Q96PB1,Q96PB8,Q96PD2,Q96PD6,Q96PD7,Q96PE1,Q96PE5,Q96PG1,Q96PG2,Q96PH1,Q96PJ5,Q96PL2,Q96PL5,Q96PN6,Q96PQ0,Q96PQ1,Q96PR1,Q96PS6,Q96PS8,Q96PX8,Q96PZ7,Q96Q04,Q96Q06,Q96Q45,Q96Q80,Q96Q91,Q96QD8,Q96QE2,Q96QE4,Q96QI5,Q96QK8,Q96QS1,Q96QT4,Q96QU1,Q96QV1,Q96QZ0,Q96R08,Q96R09,Q96R27,Q96R28,Q96R30,Q96R45,Q96R47,Q96R48,Q96R54,Q96R67,Q96R69,Q96R72,Q96R84,Q96RA2,Q96RB7,Q96RC9,Q96RD0,Q96RD1,Q96RD2,Q96RD3,Q96RD6,Q96RD7,Q96RD9,Q96RI0,Q96RI8,Q96RI9,Q96RJ0,Q96RJ3,Q96RL6,Q96RN1,Q96RP7,Q96RP8,Q96RQ1,Q96RT6,Q96RV3,Q96RW7,Q96S06,Q96S21,Q96S37,Q96S52,Q96S66,Q96S79,Q96S86,Q96S97,Q96SA4,Q96SE0,Q96SJ8,Q96SK2,Q96SL1,Q96SN7,Q96T49,Q96T52,Q96T53,Q96T54,Q96T55,Q96T83,Q96TA0,Q96TA2,Q96TC7,Q99062,Q99075,Q99102,Q99217,Q99218,Q99250,Q99418,Q99435,Q99437,Q99442,Q99445,Q99463,Q99466,Q99467,Q99500,Q99523,Q99527,Q99541,Q99542,Q99547,Q99571,Q99572,Q99595,Q99624,Q99643,Q99645,Q99650,Q99665,Q99677,Q99678,Q99679,Q99680,Q99705,Q99706,Q99712,Q99715,Q99720,Q99726,Q99735,Q99747,Q99755,Q99758,Q99788,Q99795,Q99797,Q99805,Q99808,Q99828,Q99829,Q99835,Q99884,Q99928,Q99941,Q99942,Q99943,Q99946,Q99965,Q99983,Q99999,Q9BPV8,Q9BPX6,Q9BPZ7,Q9BQ16,Q9BQ31,Q9BQ49,Q9BQ51,Q9BQA9,Q9BQB4,Q9BQB6,Q9BQD7,Q9BQE4,Q9BQG1,Q9BQI5,Q9BQI7,Q9BQJ4,Q9BQQ7,Q9BQS2,Q9BQS7,Q9BQT8,Q9BQT9,Q9BR10,Q9BR26,Q9BR39,Q9BRB3,Q9BRC7,Q9BRI3,Q9BRK0,Q9BRK3,Q9BRL7,Q9BRN9,Q9BRQ5,Q9BRQ8,Q9BRR3,Q9BRV3,Q9BRY0,Q9BS91,Q9BSA4,Q9BSA9,Q9BSE2,Q9BSE4,Q9BSF0,Q9BSJ5,Q9BSJ8,Q9BSK0,Q9BSK2,Q9BSN7,Q9BSR8,Q9BSW7,Q9BT22,Q9BT67,Q9BT76,Q9BT88,Q9BTD3,Q9BTN0,Q9BTV4,Q9BTX1,Q9BTX3,Q9BU23,Q9BU79,Q9BUB7,Q9BUD6,Q9BUF7,Q9BUJ0,Q9BUM1,Q9BUN8,Q9BUR5,Q9BV10,Q9BV23,Q9BV35,Q9BV40,Q9BV81,Q9BV87,Q9BVA6,Q9BVC6,Q9BVG9,Q9BVH7,Q9BVI4,Q9BVK2,Q9BVK6,Q9BVK8,Q9BVT8,Q9BVV7,Q9BVV8,Q9BVW6,Q9BVX2,Q9BW60,Q9BW72,Q9BWL3,Q9BWM7,Q9BWQ6,Q9BWQ8,Q9BWV1,Q9BWV2,Q9BX59,Q9BX67,Q9BX73,Q9BX74,Q9BX79,Q9BX84,Q9BX95,Q9BX97,Q9BXA5,Q9BXB1,Q9BXC0,Q9BXC1,Q9BXE9,Q9BXI2,Q9BXJ7,Q9BXJ8,Q9BXK5,Q9BXM7,Q9BXN1,Q9BXN2,Q9BXP2,Q9BXR3,Q9BXR5,Q9BXS0,Q9BXS4,Q9BXS9,Q9BXT2,Q9BXU9,Q9BXX0,Q9BY07,Q9BY08,Q9BY10,Q9BY14,Q9BY15,Q9BY19,Q9BY21,Q9BY50,Q9BY64,Q9BY67,Q9BY71,Q9BY78,Q9BY79,Q9BYC5,Q9BYE2,Q9BYE9,Q9BYF1,Q9BYG0,Q9BYH1,Q9BYJ0,Q9BYT1,Q9BYT9,Q9BYW1,Q9BZ11,Q9BZ76,Q9BZ97,Q9BZA7,Q9BZA8,Q9BZC7,Q9BZD2,Q9BZD6,Q9BZD7,Q9BZG2,Q9BZH6,Q9BZJ4,Q9BZJ6,Q9BZJ7,Q9BZJ8,Q9BZL3,Q9BZL6,Q9BZM4,Q9BZM5,Q9BZM6,Q9BZR6,Q9BZV2,Q9BZV3,Q9BZW2,Q9BZW4,Q9BZW5,Q9BZW8,Q9BZZ2,Q9C091,Q9C0A0,Q9C0B5,Q9C0B7,Q9C0C4,Q9C0D9,Q9C0E8,Q9C0H2,Q9C0I4,Q9C0J1,Q9C0K0,Q9C0K1,Q9GIP4,Q9GIY3,Q9GZK3,Q9GZK4,Q9GZK6,Q9GZK7,Q9GZM5,Q9GZM6,Q9GZN0,Q9GZN6,Q9GZP1,Q9GZP7,Q9GZP9,Q9GZQ4,Q9GZQ6,Q9GZR5,Q9GZS9,Q9GZT5,Q9GZT6,Q9GZU1,Q9GZU3,Q9GZU5,Q9GZV3,Q9GZV5,Q9GZV7,Q9GZW8,Q9GZX3,Q9GZY4,Q9GZY6,Q9GZY8,Q9GZZ6,Q9GZZ7,Q9H013,Q9H015,Q9H061,Q9H0A3,Q9H0C2,Q9H0C3,Q9H0H0,Q9H0Q3,Q9H0R3,Q9H0U3,Q9H0U4,Q9H0V1,Q9H0V9,Q9H0X4,Q9H0X9,Q9H115,Q9H156,Q9H158,Q9H159,Q9H172,Q9H195,Q9H1B5,Q9H1C0,Q9H1C3,Q9H1C4,Q9H1C7,Q9H1D0,Q9H1E5,Q9H1J5,Q9H1J7,Q9H1K0,Q9H1K4,Q9H1N7,Q9H1U4,Q9H1U9,Q9H1V8,Q9H1X3,Q9H1Y3,Q9H1Z9,Q9H205,Q9H207,Q9H208,Q9H209,Q9H210,Q9H221,Q9H222,Q9H223,Q9H228,Q9H237,Q9H239,Q9H244,Q9H251,Q9H252,Q9H255,Q9H295,Q9H2A7,Q9H2A9,Q9H2B2,Q9H2B4,Q9H2C2,Q9H2C5,Q9H2C8,Q9H2D1,Q9H2E6,Q9H2F3,Q9H2H9,Q9H2J7,Q9H2L4,Q9H2Q1,Q9H2S1,Q9H2S6,Q9H2U9,Q9H2V7,Q9H2W1,Q9H2X3,Q9H2X8,Q9H2X9,Q9H2Y9,Q9H300,Q9H306,Q9H310,Q9H313,Q9H324,Q9H330,Q9H339,Q9H340,Q9H341,Q9H342,Q9H343,Q9H344,Q9H346,Q9H3H5,Q9H3K2,Q9H3M0,Q9H3N1,Q9H3N8,Q9H3Q3,Q9H3R1,Q9H3R2,Q9H3S1,Q9H3S3,Q9H3S5,Q9H3T2,Q9H3T3,Q9H3U1,Q9H3U5,Q9H3U7,Q9H3V2,Q9H3W5,Q9H3Z4,Q9H400,Q9H427,Q9H461,Q9H490,Q9H496,Q9H4A6,Q9H4A9,Q9H4B8,Q9H4D0,Q9H4E5,Q9H4F1,Q9H4F8,Q9H4I9,Q9H4M9,Q9H553,Q9H598,Q9H5I5,Q9H5J4,Q9H5K3,Q9H5V8,Q9H5Y7,Q9H665,Q9H6A9,Q9H6B4,Q9H6D3,Q9H6D8,Q9H6F2,Q9H6H4,Q9H6L2,Q9H6L5,Q9H6R6,Q9H6U8,Q9H6X2,Q9H6X4,Q9H6Y7,Q9H720,Q9H756,Q9H799,Q9H7F0,Q9H7F4,Q9H7M9,Q9H7T0,Q9H7V2,Q9H7X2,Q9H7Z7,Q9H813,Q9H819,Q9H841,Q9H8J5,Q9H8L6,Q9H8M5,Q9H8M9,Q9H8P0,Q9H8X9,Q9H902,Q9H920,Q9H936,Q9H9B4,Q9H9K5,Q9H9P2,Q9H9S3,Q9H9S5,Q9H9V4,Q9HA72,Q9HA82,Q9HAB3,Q9HAP6,Q9HAR2,Q9HAS3,Q9HAT1,Q9HAV5,Q9HAW7,Q9HAW8,Q9HAW9,Q9HB03,Q9HB14,Q9HB15,Q9HB29,Q9HB63,Q9HB89,Q9HBA0,Q9HBB8,Q9HBE5,Q9HBG4,Q9HBG7,Q9HBH0,Q9HBI6,Q9HBJ8,Q9HBL6,Q9HBL7,Q9HBM0,Q9HBR0,Q9HBT6,Q9HBU9,Q9HBV1,Q9HBV2,Q9HBW0,Q9HBW1,Q9HBW9,Q9HBX8,Q9HBX9,Q9HBY0,Q9HC07,Q9HC10,Q9HC21,Q9HC24,Q9HC47,Q9HC56,Q9HC58,Q9HC73,Q9HC97,Q9HCB6,Q9HCC8,Q9HCE9,Q9HCF6,Q9HCJ1,Q9HCJ2,Q9HCK4,Q9HCL0,Q9HCL2,Q9HCM2,Q9HCM3,Q9HCN3,Q9HCN6,Q9HCP6,Q9HCQ5,Q9HCS2,Q9HCU0,Q9HCU4,Q9HCU5,Q9HCX4,Q9HD20,Q9HD23,Q9HD36,Q9HD43,Q9HD45,Q9HD87,Q9HDB5,Q9HDC5,Q9HDC9,Q9HDD0,Q9NNX6,Q9NNZ3,Q9NP31,Q9NP58,Q9NP59,Q9NP60,Q9NP70,Q9NP72,Q9NP78,Q9NP80,Q9NP84,Q9NP85,Q9NP90,Q9NP91,Q9NP94,Q9NP99,Q9NPA0,Q9NPA1,Q9NPA2,Q9NPB0,Q9NPB9,Q9NPC1,Q9NPC2,Q9NPC4,Q9NPD5,Q9NPD7,Q9NPE6,Q9NPF0,Q9NPF2,Q9NPG1,Q9NPG4,Q9NPG8,Q9NPH3,Q9NPH5,Q9NPI0,Q9NPI9,Q9NPL8,Q9NPR2,Q9NPR9,Q9NPU4,Q9NPY3,Q9NPZ5,Q9NQ11,Q9NQ25,Q9NQ34,Q9NQ40,Q9NQ60,Q9NQ76,Q9NQ79,Q9NQ84,Q9NQ90,Q9NQA5,Q9NQC3,Q9NQG1,Q9NQG6,Q9NQN1,Q9NQQ7,Q9NQR9,Q9NQS3,Q9NQS5,Q9NQW8,Q9NQX5,Q9NQX7,Q9NQZ7,Q9NR16,Q9NR22,Q9NR34,Q9NR61,Q9NR71,Q9NR77,Q9NR82,Q9NR96,Q9NR97,Q9NRA2,Q9NRB3,Q9NRC1,Q9NRD8,Q9NRD9,Q9NRE1,Q9NRJ7,Q9NRK6,Q9NRM0,Q9NRM1,Q9NRM6,Q9NRP0,Q9NRQ2,Q9NRQ5,Q9NRR2,Q9NRS4,Q9NRU3,Q9NRX5,Q9NRX6,Q9NRY6,Q9NRY7,Q9NRZ5,Q9NRZ7,Q9NS00,Q9NS40,Q9NS61,Q9NS62,Q9NS64,Q9NS66,Q9NS67,Q9NS68,Q9NS69,Q9NS75,Q9NS82,Q9NS84,Q9NS93,Q9NSA0,Q9NSA2,Q9NSC7,Q9NSD5,Q9NSD7,Q9NSE7,Q9NSI5,Q9NSK7,Q9NST1,Q9NT22,Q9NT68,Q9NT99,Q9NTG1,Q9NTI2,Q9NTJ5,Q9NTN3,Q9NTN9,Q9NTQ9,Q9NU53,Q9NUB4,Q9NUD9,Q9NUE0,Q9NUH8,Q9NUM3,Q9NUM4,Q9NUN5,Q9NUN7,Q9NUP9,Q9NUQ2,Q9NUR3,Q9NUT2,Q9NUV7,Q9NV12,Q9NV29,Q9NV58,Q9NV64,Q9NV92,Q9NV96,Q9NVA4,Q9NVC3,Q9NVI7,Q9NVM1,Q9NVV0,Q9NVV5,Q9NW15,Q9NW97,Q9NWC5,Q9NWD8,Q9NWF4,Q9NWH2,Q9NWH9,Q9NWM0,Q9NWQ8,Q9NWR8,Q9NWS6,Q9NWW5,Q9NWW9,Q9NX00,Q9NX14,Q9NX47,Q9NX52,Q9NX61,Q9NX62,Q9NX76,Q9NX78,Q9NX94,Q9NX95,Q9NXB9,Q9NXE4,Q9NXF1,Q9NXF8,Q9NXG6,Q9NXH8,Q9NXI6,Q9NXJ0,Q9NXK6,Q9NXL6,Q9NXS2,Q9NXW2,Q9NY15,Q9NY25,Q9NY26,Q9NY28,Q9NY35,Q9NY37,Q9NY46,Q9NY47,Q9NY64,Q9NY72,Q9NY84,Q9NY91,Q9NY97,Q9NY99,Q9NYB5,Q9NYG2,Q9NYG8,Q9NYJ7,Q9NYJ8,Q9NYK1,Q9NYL4,Q9NYM4,Q9NYM9,Q9NYP7,Q9NYQ6,Q9NYQ7,Q9NYQ8,Q9NYR8,Q9NYV7,Q9NYV8,Q9NYV9,Q9NYW0,Q9NYW1,Q9NYW2,Q9NYW3,Q9NYW4,Q9NYW5,Q9NYW6,Q9NYW7,Q9NYX4,Q9NYZ1,Q9NYZ2,Q9NYZ4,Q9NZ01,Q9NZ08,Q9NZ42,Q9NZ43,Q9NZ45,Q9NZ53,Q9NZ94,Q9NZA1,Q9NZC2,Q9NZC3,Q9NZD1,Q9NZG7,Q9NZH0,Q9NZI2,Q9NZJ5,Q9NZJ7,Q9NZM1,Q9NZM6,Q9NZN1,Q9NZN3,Q9NZN4,Q9NZP0,Q9NZP2,Q9NZP5,Q9NZQ7,Q9NZQ8,Q9NZR2,Q9NZS2,Q9NZS9,Q9NZU0,Q9NZU1,Q9NZV1,Q9NZV8,Q9NZW4,Q9NZW5,Q9P003,Q9P035,Q9P055,Q9P0B6,Q9P0G3,Q9P0I2,Q9P0J0,Q9P0K1,Q9P0K9,Q9P0L0,Q9P0L2,Q9P0L9,Q9P0N5,Q9P0N8,Q9P0S2,Q9P0S3,Q9P0S9,Q9P0T7,Q9P0U1,Q9P0V3,Q9P0V8,Q9P0X4,Q9P109,Q9P121,Q9P126,Q9P1A6,Q9P1P4,Q9P1P5,Q9P1Q5,Q9P1W3,Q9P1W8,Q9P1Z3,Q9P1Z9,Q9P218,Q9P232,Q9P241,Q9P244,Q9P246,Q9P273,Q9P283,Q9P291,Q9P296,Q9P298,Q9P2B2,Q9P2C4,Q9P2D8,Q9P2E5,Q9P2E7,Q9P2E8,Q9P2E9,Q9P2J2,Q9P2K9,Q9P2N4,Q9P2P1,Q9P2S2,Q9P2U7,Q9P2U8,Q9P2V4,Q9P2W3,Q9P2W7,Q9P2W9,Q9P2X0,Q9TNN7,Q9TQE0,Q9UBC2,Q9UBD6,Q9UBF9,Q9UBG0,Q9UBH6,Q9UBI4,Q9UBI6,Q9UBJ2,Q9UBK5,Q9UBL9,Q9UBM1,Q9UBM4,Q9UBM7,Q9UBM8,Q9UBN1,Q9UBN4,Q9UBN6,Q9UBP0,Q9UBQ6,Q9UBR5,Q9UBS5,Q9UBS9,Q9UBU6,Q9UBV2,Q9UBV4,Q9UBV7,Q9UBX3,Q9UBX8,Q9UBY0,Q9UBY5,Q9UBY8,Q9UDX4,Q9UDX5,Q9UEF7,Q9UEU0,Q9UEW3,Q9UF02,Q9UF11,Q9UF33,Q9UF47,Q9UG22,Q9UGC6,Q9UGF5,Q9UGF6,Q9UGF7,Q9UGH3,Q9UGI6,Q9UGM1,Q9UGN4,Q9UGP8,Q9UGQ2,Q9UGQ3,Q9UGT4,Q9UH62,Q9UH99,Q9UHC3,Q9UHC6,Q9UHC9,Q9UHE5,Q9UHE8,Q9UHF3,Q9UHF4,Q9UHI5,Q9UHI7,Q9UHI8,Q9UHJ9,Q9UHM6,Q9UHN6,Q9UHP7,Q9UHQ4,Q9UHQ9,Q9UHW9,Q9UHX3,Q9UI14,Q9UI33,Q9UI40,Q9UIA0,Q9UIB8,Q9UIG8,Q9UIJ5,Q9UIK5,Q9UIQ6,Q9UIR0,Q9UIW2,Q9UIX4,Q9UJ14,Q9UJ37,Q9UJ42,Q9UJ71,Q9UJ90,Q9UJ96,Q9UJ99,Q9UJA2,Q9UJA9,Q9UJG1,Q9UJQ1,Q9UJS0,Q9UJW2,Q9UK00,Q9UK08,Q9UK17,Q9UK23,Q9UK28,Q9UK39,Q9UK41,Q9UKB5,Q9UKC9,Q9UKF2,Q9UKF5,Q9UKG4,Q9UKJ0,Q9UKJ1,Q9UKJ5,Q9UKJ8,Q9UKL2,Q9UKL4,Q9UKM7,Q9UKN1,Q9UKP4,Q9UKP5,Q9UKP6,Q9UKQ2,Q9UKR5,Q9UKR8,Q9UKU0,Q9UKU6,Q9UKV5,Q9UKX5,Q9UKY0,Q9UKY4,Q9UKZ4,Q9UL01,Q9UL19,Q9UL51,Q9UL52,Q9UL54,Q9UL62,Q9ULB1,Q9ULB4,Q9ULB5,Q9ULC0,Q9ULC3,Q9ULC5,Q9ULC8,Q9ULD8,Q9ULF5,Q9ULG6,Q9ULH0,Q9ULH4,Q9ULI3,Q9ULK0,Q9ULK5,Q9ULK6,Q9ULL4,Q9ULM6,Q9ULQ1,Q9ULS5,Q9ULS6,Q9ULT6,Q9ULV1,Q9ULW2,Q9ULX3,Q9ULX5,Q9ULX7,Q9ULY5,Q9ULZ9,Q9UM00,Q9UM01,Q9UM11,Q9UM21,Q9UM44,Q9UM47,Q9UM73,Q9UMD9,Q9UMF0,Q9UMR7,Q9UMX5,Q9UMX6,Q9UMX9,Q9UMY4,Q9UMZ3,Q9UN42,Q9UN66,Q9UN67,Q9UN70,Q9UN71,Q9UN72,Q9UN73,Q9UN74,Q9UN75,Q9UN76,Q9UN88,Q9UNA0,Q9UNA3,Q9UNE0,Q9UNG2,Q9UNK0,Q9UNL2,Q9UNN8,Q9UNP4,Q9UNQ0,Q9UNU6,Q9UNW8,Q9UNX9,Q9UP38,Q9UP52,Q9UP65,Q9UP79,Q9UP95,Q9UPC5,Q9UPI3,Q9UPQ8,Q9UPR5,Q9UPU3,Q9UPX0,Q9UPX6,Q9UPY5,Q9UPZ6,Q9UQ05,Q9UQ26,Q9UQ49,Q9UQ52,Q9UQ53,Q9UQ90,Q9UQC9,Q9UQD0,Q9UQF0,Q9UQP3,Q9UQQ1,Q9UQV4,Q9Y210,Q9Y219,Q9Y225,Q9Y226,Q9Y227,Q9Y228,Q9Y231,Q9Y241,Q9Y256,Q9Y257,Q9Y267,Q9Y271,Q9Y272,Q9Y274,Q9Y275,Q9Y276,Q9Y277,Q9Y278,Q9Y279,Q9Y282,Q9Y284,Q9Y286,Q9Y287,Q9Y289,Q9Y2A7,Q9Y2A9,Q9Y2B1,Q9Y2B2,Q9Y2C2,Q9Y2C3,Q9Y2C5,Q9Y2C9,Q9Y2D2,Q9Y2E8,Q9Y2G1,Q9Y2G3,Q9Y2G8,Q9Y2H0,Q9Y2H6,Q9Y2H9,Q9Y2I1,Q9Y2I2,Q9Y2I9,Q9Y2P4,Q9Y2P5,Q9Y2Q0,Q9Y2R0,Q9Y2T5,Q9Y2T6,Q9Y2U2,Q9Y2U8,Q9Y2W3,Q9Y2Y6,Q9Y320,Q9Y328,Q9Y336,Q9Y342,Q9Y345,Q9Y385,Q9Y397,Q9Y3A6,Q9Y3B3,Q9Y3D6,Q9Y3E0,Q9Y3N9,Q9Y3P4,Q9Y3P8,Q9Y3Q0,Q9Y3Q3,Q9Y3Q4,Q9Y3Q7,Q9Y442,Q9Y487,Q9Y493,Q9Y4A9,Q9Y4C0,Q9Y4C5,Q9Y4D2,Q9Y4D7,Q9Y4D8,Q9Y4F1,Q9Y4G8,Q9Y4K0,Q9Y4W6,Q9Y4X1,Q9Y519,Q9Y548,Q9Y561,Q9Y584,Q9Y585,Q9Y5E1,Q9Y5E2,Q9Y5E3,Q9Y5E4,Q9Y5E5,Q9Y5E6,Q9Y5E7,Q9Y5E8,Q9Y5E9,Q9Y5F0,Q9Y5F1,Q9Y5F2,Q9Y5F3,Q9Y5F6,Q9Y5F7,Q9Y5F8,Q9Y5F9,Q9Y5G0,Q9Y5G1,Q9Y5G2,Q9Y5G3,Q9Y5G4,Q9Y5G5,Q9Y5G6,Q9Y5G7,Q9Y5G8,Q9Y5G9,Q9Y5H0,Q9Y5H1,Q9Y5H2,Q9Y5H3,Q9Y5H4,Q9Y5H5,Q9Y5H6,Q9Y5H7,Q9Y5H8,Q9Y5H9,Q9Y5I0,Q9Y5I1,Q9Y5I2,Q9Y5I3,Q9Y5I4,Q9Y5I7,Q9Y5K8,Q9Y5L2,Q9Y5L3,Q9Y5M8,Q9Y5N1,Q9Y5P0,Q9Y5P1,Q9Y5Q0,Q9Y5Q5,Q9Y5R2,Q9Y5S1,Q9Y5S8,Q9Y5T4,Q9Y5U4,Q9Y5U5,Q9Y5U8,Q9Y5U9,Q9Y5W7,Q9Y5W9,Q9Y5X5,Q9Y5Y0,Q9Y5Y3,Q9Y5Y4,Q9Y5Y5,Q9Y5Y6,Q9Y5Y7,Q9Y5Y9,Q9Y5Z0,Q9Y5Z6,Q9Y5Z9,Q9Y619,Q9Y624,Q9Y625,Q9Y639,Q9Y644,Q9Y653,Q9Y661,Q9Y662,Q9Y663,Q9Y666,Q9Y672,Q9Y673,Q9Y679,Q9Y691,Q9Y693,Q9Y694,Q9Y696,Q9Y698,Q9Y6A1,Q9Y6A2,Q9Y6A9,Q9Y6C2,Q9Y6C5,Q9Y6C9,Q9Y6F6,Q9Y6F9,Q9Y6G1,Q9Y6H6,Q9Y6H8,Q9Y6I0,Q9Y6I8,Q9Y6J6,Q9Y6K0,Q9Y6L6,Q9Y6M0,Q9Y6M5,Q9Y6M7,Q9Y6N1,Q9Y6N6,Q9Y6N7,Q9Y6N8,Q9Y6Q6,Q9Y6Q9,Q9Y6R1,Q9Y6U7,Q9Y6W8,Q9Y6X1,Q9Y6X5,Q9Y6Y9';
const _PM_RAW='A0A075B6H7,A0A075B6H8,A0A075B6H9,A0A075B6I0,A0A075B6I1,A0A075B6I3,A0A075B6I4,A0A075B6I6,A0A075B6I7,A0A075B6I9,A0A075B6J1,A0A075B6J2,A0A075B6J6,A0A075B6J9,A0A075B6K0,A0A075B6K2,A0A075B6K4,A0A075B6K5,A0A075B6K6,A0A075B6L2,A0A075B6L6,A0A075B6N1,A0A075B6N2,A0A075B6N3,A0A075B6N4,A0A075B6P5,A0A075B6Q5,A0A075B6R0,A0A075B6R2,A0A075B6R9,A0A075B6S0,A0A075B6S2,A0A075B6S4,A0A075B6S5,A0A075B6S6,A0A075B6S9,A0A075B6T6,A0A075B6T7,A0A075B6T8,A0A075B6U4,A0A075B6V5,A0A075B6W5,A0A075B6X5,A0A075B6Y3,A0A075B6Y9,A0A075B700,A0A075B706,A0A087WSX0,A0A087WSY4,A0A087WSY6,A0A087WSZ0,A0A087WSZ9,A0A087WT01,A0A087WT02,A0A087WT03,A0A087WV62,A0A087WVF3,A0A087WW87,A0A087WXS9,A0A087X0K7,A0A087X0M5,A0A087X179,A0A087X1G2,A0A0A0MRZ7,A0A0A0MRZ8,A0A0A0MRZ9,A0A0A0MS00,A0A0A0MS01,A0A0A0MS02,A0A0A0MS03,A0A0A0MS04,A0A0A0MS05,A0A0A0MS06,A0A0A0MS14,A0A0A0MS15,A0A0A0MT36,A0A0A0MT70,A0A0A0MT76,A0A0A0MT78,A0A0A0MT87,A0A0A0MT89,A0A0A0MT94,A0A0A0MTA4,A0A0A0MTA7,A0A0A6YYC5,A0A0A6YYD4,A0A0A6YYG2,A0A0A6YYG3,A0A0A6YYJ7,A0A0A6YYK1,A0A0A6YYK4,A0A0A6YYK6,A0A0A6YYK7,A0A0B4J1U3,A0A0B4J1U4,A0A0B4J1U6,A0A0B4J1U7,A0A0B4J1V0,A0A0B4J1V1,A0A0B4J1V2,A0A0B4J1V6,A0A0B4J1V7,A0A0B4J1X5,A0A0B4J1X8,A0A0B4J1Y8,A0A0B4J1Y9,A0A0B4J1Z2,A0A0B4J200,A0A0B4J234,A0A0B4J235,A0A0B4J237,A0A0B4J238,A0A0B4J240,A0A0B4J241,A0A0B4J244,A0A0B4J245,A0A0B4J248,A0A0B4J249,A0A0B4J262,A0A0B4J263,A0A0B4J264,A0A0B4J265,A0A0B4J266,A0A0B4J268,A0A0B4J271,A0A0B4J272,A0A0B4J273,A0A0B4J274,A0A0B4J275,A0A0B4J276,A0A0B4J277,A0A0B4J279,A0A0B4J280,A0A0B4J2D9,A0A0B4J2E0,A0A0B4J2H0,A0A0C4DH24,A0A0C4DH25,A0A0C4DH26,A0A0C4DH27,A0A0C4DH28,A0A0C4DH29,A0A0C4DH30,A0A0C4DH31,A0A0C4DH32,A0A0C4DH33,A0A0C4DH34,A0A0C4DH35,A0A0C4DH36,A0A0C4DH38,A0A0C4DH39,A0A0C4DH41,A0A0C4DH42,A0A0C4DH43,A0A0C4DH55,A0A0C4DH59,A0A0C4DH62,A0A0C4DH67,A0A0C4DH68,A0A0C4DH69,A0A0C4DH72,A0A0C4DH73,A0A0G2JMI3,A0A0G2JS06,A0A0J9YVY3,A0A0J9YWP8,A0A0J9YWX3,A0A0J9YX06,A0A0J9YX35,A0A0J9YX75,A0A0J9YXA8,A0A0J9YXG5,A0A0J9YXM7,A0A0J9YXX1,A0A0J9YXY3,A0A0K0K1A3,A0A0K0K1A5,A0A0K0K1B3,A0A0K0K1C0,A0A0K0K1C4,A0A0K0K1D8,A0A0K0K1E9,A0A0K0K1G6,A0A0K0K1G8,A0A1B0GTQ4,A0A1B0GX31,A0A1B0GX49,A0A1B0GX51,A0A1B0GX56,A0A1B0GX68,A0A1B0GX78,A0A1B0GX95,A0A1B0GXF2,A0A1W2PPG7,A0A539,A0A576,A0A577,A0A578,A0A584,A0A589,A0A597,A0A599,A0A5A2,A0A5A6,A0A5B0,A0A5B6,A0A5B7,A0A5B9,A0AV02,A0AVI2,A0AVI4,A0FGR8,A0FGR9,A0JD32,A0JD36,A0JD37,A0M8Q6,A0PJK1,A0PJW6,A0PJX4,A0PJX8,A0PJZ3,A0PK00,A0PK05,A0PK11,A0ZSE6,A1A4S6,A1A4Y4,A1A5B4,A1A5C7,A1KXE4,A1L0T0,A1L157,A1L1A6,A1L3X0,A1L4H1,A2A2V5,A2A2Y4,A2CJ06,A2IDD5,A2NJV5,A2RRL7,A2RU14,A2RU48,A2RU67,A2RUG3,A2RUT3,A2VDJ0,A2VEC9,A3KFT3,A4D0S4,A4D0T2,A4D1S0,A4D1S5,A4D256,A4D2G3,A4D2H0,A4D2P6,A4FU28,A4IF30,A5D6W6,A5PKW4,A5PLK6,A5PLL7,A5X5Y0,A6BM72,A6H8M9,A6NC51,A6NC97,A6NCI5,A6NCQ9,A6NCV1,A6ND01,A6ND48,A6NDA9,A6NDB9,A6NDD5,A6NDH6,A6NDL8,A6NDP7,A6NDS4,A6NDV4,A6NEH6,A6NER0,A6NET4,A6NF34,A6NF89,A6NFA1,A6NFC5,A6NFC9,A6NFE2,A6NFQ2,A6NFR6,A6NFU0,A6NFX1,A6NFY4,A6NGA9,A6NGB0,A6NGB7,A6NGC4,A6NGU5,A6NGY5,A6NGZ8,A6NH00,A6NH21,A6NH52,A6NHA9,A6NHG9,A6NHS7,A6NI61,A6NI73,A6NIJ9,A6NIL9,A6NIM6,A6NIZ1,A6NJW4,A6NJW9,A6NJY1,A6NJZ3,A6NK97,A6NKB5,A6NKC4,A6NKF7,A6NKG5,A6NKK0,A6NKL6,A6NKW6,A6NKX4,A6NL05,A6NL08,A6NL26,A6NL88,A6NL99,A6NLE4,A6NLU5,A6NLX4,A6NM03,A6NM10,A6NM11,A6NM45,A6NM62,A6NM76,A6NMB1,A6NMD0,A6NML5,A6NMS3,A6NMS7,A6NMU1,A6NMY6,A6NMZ5,A6NMZ7,A6NN92,A6NNB3,A6NND4,A6NNE9,A6NNN8,A6NNS2,A6PVL3,A6QL63,A6XGL0,A7E2Y1,A7MBM2,A8CG34,A8K4G0,A8K7I4,A8K855,A8MPY1,A8MUP6,A8MV81,A8MVG2,A8MVS5,A8MVW0,A8MVW5,A8MVZ5,A8MWK0,A8MWL7,A8MWY0,A8MXK1,A8MXV6,A8MYB1,A8MYU2,A8MZ97,A8TX70,A9Z1Z3,B0FP48,B0I1T2,B0YJ81,B1AH88,B1AL88,B2RN74,B2RTY4,B2RUZ4,B2RXF0,B3SHH9,B4DJY2,B4DS77,B4DYI2,B6A8C7,B6SEH8,B6SEH9,B7U540,B7Z8K6,B7ZAQ6,B8ZZ34,B9A6J9,B9EJG8,C9J069,C9J798,C9JDP6,C9JH25,C9JI98,C9JQI7,C9JVW0,D3W0D1,E2RYF6,E7ERA6,E9PQ53,F2Z333,F5H4A9,F7VJQ1,G3V0H7,H0YL14,H3BQJ8,H3BR10,H3BS89,H3BV60,H7C241,I3L273,O00141,O00144,O00155,O00159,O00161,O00165,O00168,O00180,O00186,O00194,O00198,O00206,O00212,O00213,O00219,O00220,O00222,O00237,O00238,O00241,O00254,O00258,O00264,O00270,O00294,O00299,O00322,O00337,O00341,O00391,O00398,O00400,O00408,O00421,O00443,O00445,O00451,O00453,O00461,O00468,O00476,O00478,O00481,O00499,O00501,O00515,O00519,O00522,O00526,O00533,O00548,O00555,O00559,O00560,O00571,O00574,O00587,O00590,O00591,O00592,O00602,O00623,O00624,O00631,O00634,O00744,O00750,O00755,O00767,O14490,O14492,O14493,O14494,O14495,O14511,O14512,O14514,O14520,O14521,O14522,O14523,O14524,O14525,O14526,O14569,O14581,O14609,O14610,O14626,O14638,O14640,O14641,O14649,O14653,O14654,O14662,O14668,O14669,O14672,O14678,O14681,O14683,O14684,O14713,O14718,O14735,O14745,O14763,O14764,O14786,O14788,O14795,O14798,O14804,O14807,O14817,O14827,O14828,O14836,O14842,O14843,O14863,O14880,O14894,O14904,O14905,O14907,O14910,O14917,O14925,O14931,O14936,O14939,O14944,O14957,O14966,O14967,O14975,O14983,O15031,O15033,O15034,O15056,O15068,O15072,O15079,O15118,O15120,O15121,O15126,O15127,O15146,O15155,O15162,O15165,O15169,O15173,O15197,O15218,O15229,O15230,O15239,O15243,O15244,O15245,O15247,O15255,O15258,O15260,O15269,O15270,O15294,O15296,O15303,O15321,O15335,O15342,O15354,O15374,O15375,O15389,O15393,O15394,O15399,O15400,O15403,O15427,O15431,O15432,O15438,O15439,O15440,O15455,O15466,O15482,O15503,O15529,O15530,O15533,O15537,O15547,O15551,O15552,O15554,O42043,O43150,O43155,O43157,O43164,O43169,O43173,O43184,O43193,O43194,O43246,O43280,O43286,O43291,O43292,O43295,O43300,O43306,O43315,O43318,O43323,O43374,O43405,O43424,O43451,O43462,O43464,O43490,O43491,O43493,O43497,O43505,O43506,O43508,O43511,O43520,O43525,O43526,O43529,O43556,O43557,O43559,O43561,O43566,O43567,O43570,O43581,O43586,O43597,O43603,O43613,O43614,O43653,O43657,O43674,O43676,O43677,O43687,O43688,O43699,O43731,O43736,O43739,O43745,O43749,O43752,O43759,O43760,O43761,O43768,O43772,O43808,O43813,O43820,O43825,O43826,O43861,O43865,O43868,O43869,O43889,O43895,O43900,O43908,O43909,O43914,O43916,O43921,O43934,O60235,O60238,O60241,O60242,O60243,O60245,O60262,O60266,O60279,O60291,O60292,O60309,O60312,O60313,O60320,O60330,O60331,O60337,O60344,O60346,O60353,O60359,O60391,O60403,O60404,O60412,O60423,O60427,O60431,O60437,O60443,O60449,O60462,O60469,O60476,O60478,O60486,O60487,O60488,O60494,O60499,O60500,O60503,O60507,O60512,O60513,O60602,O60603,O60609,O60610,O60635,O60636,O60637,O60641,O60656,O60667,O60669,O60704,O60706,O60711,O60716,O60721,O60725,O60733,O60741,O60755,O60774,O60779,O60830,O60831,O60840,O60844,O60858,O60882,O60883,O60884,O60894,O60895,O60896,O60906,O60909,O60928,O60931,O60938,O60939,O71037,O75015,O75019,O75022,O75023,O75027,O75038,O75044,O75051,O75054,O75056,O75063,O75069,O75071,O75072,O75074,O75077,O75078,O75084,O75096,O75106,O75110,O75116,O75121,O75122,O75128,O75129,O75131,O75144,O75161,O75173,O75185,O75192,O75197,O75204,O75264,O75298,O75309,O75310,O75311,O75324,O75325,O75326,O75339,O75340,O75352,O75354,O75355,O75356,O75365,O75379,O75386,O75387,O75388,O75396,O75398,O75420,O75425,O75427,O75438,O75443,O75445,O75452,O75460,O75473,O75477,O75487,O75508,O75509,O75558,O75575,O75578,O75581,O75592,O75608,O75631,O75695,O75712,O75718,O75746,O75747,O75751,O75752,O75762,O75781,O75783,O75787,O75795,O75800,O75829,O75838,O75841,O75844,O75845,O75871,O75882,O75899,O75900,O75907,O75908,O75911,O75915,O75916,O75923,O75949,O75954,O75955,O75970,O75976,O76000,O76001,O76002,O76024,O76036,O76050,O76062,O76081,O76082,O76083,O76087,O76090,O76095,O76099,O76100,O94759,O94766,O94768,O94769,O94772,O94777,O94778,O94779,O94804,O94812,O94823,O94826,O94827,O94850,O94856,O94868,O94875,O94876,O94886,O94898,O94901,O94905,O94910,O94911,O94921,O94923,O94933,O94956,O94966,O94973,O94985,O94991,O95006,O95007,O95013,O95047,O95049,O95057,O95069,O95070,O95136,O95139,O95140,O95150,O95159,O95164,O95167,O95168,O95169,O95180,O95183,O95185,O95196,O95202,O95206,O95210,O95214,O95221,O95222,O95237,O95249,O95255,O95256,O95258,O95259,O95264,O95267,O95274,O95278,O95279,O95292,O95297,O95298,O95342,O95371,O95377,O95395,O95406,O95415,O95425,O95427,O95436,O95450,O95452,O95461,O95466,O95470,O95471,O95473,O95476,O95477,O95484,O95490,O95497,O95498,O95500,O95502,O95528,O95562,O95563,O95573,O95619,O95622,O95631,O95661,O95665,O95672,O95674,O95711,O95716,O95721,O95727,O95741,O95754,O95772,O95782,O95786,O95800,O95803,O95807,O95832,O95833,O95838,O95847,O95857,O95858,O95859,O95864,O95866,O95867,O95868,O95870,O95876,O95886,O95907,O95918,O95944,O95971,O95976,O95977,O95980,O95990,O95992,O96002,O96005,O96009,O96011,O96014,O96017,O96024,P00156,P00167,P00395,P00403,P00414,P00451,P00505,P00533,P00734,P00750,P00813,P00846,P00918,P01008,P01042,P01111,P01112,P01116,P01130,P01133,P01135,P01137,P01375,P01589,P01593,P01594,P01597,P01599,P01601,P01602,P01611,P01614,P01615,P01619,P01624,P01699,P01700,P01701,P01703,P01704,P01705,P01706,P01709,P01714,P01715,P01717,P01718,P01721,P01730,P01732,P01733,P01737,P01742,P01743,P01762,P01763,P01764,P01766,P01767,P01768,P01772,P01780,P01782,P01814,P01817,P01824,P01825,P01833,P01834,P01848,P01850,P01854,P01857,P01859,P01860,P01861,P01871,P01876,P01877,P01880,P01889,P01891,P01892,P01893,P01903,P01906,P01909,P01911,P01912,P01920,P02452,P02458,P02461,P02462,P02686,P02708,P02724,P02730,P02748,P02751,P02786,P03372,P03886,P03891,P03897,P03901,P03905,P03915,P03923,P03928,P03950,P03956,P03979,P03986,P03989,P03999,P04000,P04001,P04004,P04035,P04049,P04083,P04156,P04201,P04211,P04216,P04222,P04229,P04233,P04234,P04264,P04430,P04432,P04433,P04435,P04437,P04439,P04440,P04626,P04628,P04629,P04632,P04746,P04839,P04843,P04844,P04899,P04920,P04921,P05023,P05026,P05067,P05093,P05106,P05107,P05109,P05129,P05141,P05154,P05156,P05164,P05186,P05187,P05362,P05496,P05534,P05538,P05556,P05981,P05997,P06028,P06126,P06127,P06133,P06213,P06239,P06241,P06310,P06312,P06315,P06331,P06340,P06702,P06703,P06729,P06731,P06733,P06734,P06756,P06858,P07093,P07099,P07202,P07204,P07237,P07306,P07307,P07332,P07333,P07355,P07357,P07359,P07384,P07477,P07478,P07510,P07550,P07585,P07711,P07766,P07858,P07900,P07911,P07942,P07947,P07948,P07949,P07988,P08034,P08069,P08100,P08123,P08134,P08138,P08172,P08173,P08174,P08183,P08195,P08238,P08247,P08253,P08254,P08294,P08473,P08514,P08571,P08572,P08574,P08575,P08581,P08582,P08588,P08631,P08637,P08648,P08670,P08684,P08754,P08842,P08865,P08887,P08908,P08910,P08912,P08913,P08922,P08962,P08F94,P09104,P09131,P09172,P09237,P09238,P09326,P09382,P09429,P09471,P09486,P09543,P09544,P09564,P09603,P09619,P09669,P09693,P09758,P09769,P09848,P09912,P09923,P09958,P0C091,P0C0E4,P0C2L3,P0C2S0,P0C2W1,P0C604,P0C617,P0C623,P0C626,P0C628,P0C629,P0C645,P0C646,P0C672,P0C6S8,P0C6T2,P0C7M8,P0C7N1,P0C7N4,P0C7N5,P0C7N8,P0C7Q5,P0C7Q6,P0C7T2,P0C7T3,P0C7T8,P0C7U0,P0C7U3,P0C7U9,P0C7V7,P0C7X1,P0C838,P0C851,P0C874,P0CAP1,P0CF51,P0CF74,P0CG04,P0CG08,P0CG37,P0CG41,P0CK96,P0CK97,P0DI80,P0DJ07,P0DJ93,P0DKB5,P0DKB6,P0DKV0,P0DKX4,P0DL12,P0DMS8,P0DMU2,P0DN77,P0DN78,P0DN80,P0DN81,P0DN82,P0DOX2,P0DOX3,P0DOX4,P0DOX5,P0DOX6,P0DOX7,P0DOX8,P0DOY2,P0DOY3,P0DOY5,P0DP01,P0DP02,P0DP03,P0DP04,P0DP06,P0DP07,P0DP08,P0DP09,P0DP58,P0DPA2,P0DPA3,P0DPF4,P0DPF7,P0DPI4,P0DPR3,P0DQD5,P0DSE1,P0DSE2,P0DSN7,P0DTE1,P0DTE2,P0DTU3,P0DTU4,P0DTW3,P10176,P10301,P10314,P10316,P10319,P10321,P10415,P10586,P10620,P10636,P10644,P10646,P10696,P10721,P10745,P10747,P10767,P10912,P10915,P10966,P11047,P11049,P11117,P11142,P11166,P11168,P11169,P11215,P11229,P11230,P11233,P11234,P11279,P11309,P11362,P11511,P11532,P11597,P11686,P11717,P11801,P11836,P11912,P12107,P12109,P12110,P12111,P12235,P12236,P12314,P12318,P12319,P12644,P12814,P12821,P12829,P12830,P12931,P13164,P13224,P13385,P13473,P13498,P13569,P13591,P13598,P13611,P13612,P13637,P13688,P13726,P13746,P13747,P13760,P13761,P13762,P13765,P13796,P13861,P13866,P13942,P13945,P13987,P14060,P14151,P14207,P14209,P14222,P14317,P14384,P14410,P14415,P14416,P14543,P14550,P14555,P14616,P14672,P14679,P14735,P14770,P14778,P14780,P14784,P14867,P15056,P15104,P15144,P15151,P15169,P15260,P15291,P15309,P15311,P15328,P15382,P15391,P15421,P15502,P15509,P15514,P15529,P15812,P15813,P15907,P15924,P15941,P15954,P16050,P16066,P16070,P16109,P16112,P16144,P16150,P16188,P16189,P16190,P16234,P16260,P16284,P16389,P16410,P16422,P16442,P16444,P16452,P16471,P16473,P16499,P16581,P16591,P16615,P16662,P16671,P16871,P16989,P17081,P17152,P17181,P17252,P17301,P17302,P17342,P17538,P17612,P17643,P17655,P17658,P17661,P17677,P17693,P17706,P17787,P17813,P17927,P17948,P18084,P18089,P18206,P18405,P18433,P18462,P18463,P18464,P18465,P18505,P18507,P18564,P18577,P18627,P18825,P18827,P18848,P18850,P19021,P19022,P19075,P19086,P19224,P19235,P19256,P19320,P19397,P19438,P19440,P19526,P19634,P19801,P20020,P20023,P20036,P20039,P20138,P20273,P20292,P20309,P20333,P20336,P20337,P20338,P20339,P20594,P20645,P20648,P20701,P20702,P20774,P20827,P20849,P20908,P20916,P20963,P21145,P21217,P21397,P21439,P21452,P21453,P21462,P21554,P21579,P21583,P21589,P21709,P21728,P21730,P21731,P21754,P21757,P21796,P21802,P21810,P21817,P21854,P21860,P21917,P21918,P21926,P21941,P21964,P22001,P22079,P22083,P22105,P22223,P22303,P22309,P22310,P22413,P22455,P22459,P22460,P22607,P22681,P22694,P22732,P22735,P22748,P22760,P22794,P22888,P22897,P23083,P23142,P23229,P23276,P23352,P23415,P23416,P23435,P23467,P23468,P23469,P23470,P23471,P23508,P23510,P23515,P23528,P23634,P23763,P23942,P23945,P23975,P24043,P24046,P24071,P24158,P24310,P24311,P24347,P24390,P24394,P24530,P24557,P24588,P24821,P25021,P25024,P25025,P25054,P25063,P25067,P25089,P25090,P25092,P25098,P25100,P25101,P25103,P25105,P25106,P25116,P25189,P25391,P25445,P25705,P25815,P25874,P25929,P25940,P25942,P26006,P26010,P26012,P26038,P26045,P26232,P26439,P26572,P26678,P26715,P26717,P26718,P26842,P26951,P26992,P27037,P27105,P27169,P27216,P27338,P27448,P27449,P27487,P27544,P27658,P27701,P27815,P27824,P27930,P28067,P28068,P28221,P28222,P28223,P28288,P28300,P28328,P28335,P28336,P28472,P28476,P28566,P28827,P28845,P28906,P28907,P28908,P29016,P29017,P29033,P29074,P29274,P29275,P29279,P29317,P29320,P29322,P29323,P29371,P29376,P29400,P29466,P29474,P29475,P29965,P29972,P29973,P29992,P30153,P30203,P30273,P30301,P30408,P30411,P30414,P30443,P30447,P30450,P30453,P30455,P30456,P30457,P30459,P30460,P30461,P30462,P30464,P30466,P30475,P30479,P30480,P30481,P30483,P30484,P30485,P30486,P30487,P30488,P30490,P30491,P30492,P30493,P30495,P30498,P30499,P30501,P30504,P30505,P30508,P30510,P30511,P30512,P30518,P30530,P30531,P30532,P30536,P30542,P30550,P30556,P30559,P30685,P30825,P30872,P30874,P30926,P30939,P30953,P30954,P30968,P30988,P30989,P31213,P31321,P31323,P31358,P31391,P31431,P31512,P31639,P31641,P31644,P31645,P31689,P31749,P31751,P31785,P31994,P31995,P31997,P32004,P32121,P32238,P32239,P32241,P32245,P32246,P32247,P32248,P32249,P32297,P32302,P32418,P32455,P32745,P32856,P32926,P32927,P32942,P32970,P32971,P33032,P33121,P33151,P33241,P33527,P33681,P33765,P33897,P33908,P33947,P34741,P34810,P34903,P34910,P34925,P34947,P34969,P34972,P34981,P34982,P34995,P34998,P35052,P35070,P35125,P35212,P35221,P35222,P35232,P35240,P35241,P35247,P35346,P35348,P35367,P35368,P35372,P35408,P35410,P35414,P35462,P35498,P35499,P35503,P35504,P35523,P35555,P35556,P35575,P35590,P35610,P35611,P35612,P35613,P35625,P35670,P35858,P35913,P35916,P35968,P36021,P36222,P36268,P36269,P36382,P36383,P36537,P36544,P36888,P36894,P36896,P36897,P36941,P36952,P36956,P37023,P37059,P37088,P37173,P37268,P37287,P37288,P38435,P38484,P38567,P38570,P39059,P39060,P39086,P39210,P39656,P39877,P39880,P39900,P40123,P40126,P40145,P40189,P40197,P40198,P40199,P40200,P40238,P40259,P40305,P40617,P40818,P40879,P40967,P41143,P41145,P41146,P41180,P41181,P41217,P41220,P41221,P41231,P41240,P41273,P41440,P41586,P41587,P41594,P41595,P41597,P41732,P41968,P42081,P42166,P42167,P42261,P42262,P42263,P42356,P42566,P42658,P42680,P42681,P42701,P42702,P42857,P42892,P43003,P43004,P43005,P43007,P43026,P43080,P43088,P43115,P43116,P43119,P43121,P43146,P43220,P43235,P43250,P43251,P43307,P43308,P43353,P43403,P43405,P43489,P43626,P43627,P43628,P43629,P43630,P43631,P43632,P43657,P43681,P45452,P45844,P45880,P46019,P46020,P46059,P46089,P46091,P46092,P46093,P46094,P46095,P46098,P46108,P46531,P46663,P46695,P46721,P46934,P46939,P46940,P46977,P47211,P47775,P47804,P47869,P47870,P47871,P47872,P47881,P47883,P47884,P47887,P47888,P47890,P47893,P47898,P47900,P47901,P47985,P48023,P48029,P48039,P48048,P48050,P48051,P48058,P48059,P48060,P48065,P48066,P48067,P48145,P48146,P48165,P48167,P48169,P48201,P48230,P48357,P48426,P48509,P48544,P48546,P48547,P48549,P48551,P48651,P48664,P48730,P48736,P48751,P48764,P48960,P48995,P49006,P49019,P49069,P49146,P49190,P49238,P49257,P49279,P49281,P49286,P49327,P49407,P49447,P49641,P49682,P49683,P49685,P49703,P49747,P49755,P49757,P49758,P49768,P49771,P49788,P49795,P49796,P49802,P49810,P49841,P49895,P49903,P49908,P49961,P50052,P50148,P50150,P50151,P50281,P50391,P50402,P50406,P50416,P50443,P50552,P50591,P50607,P50851,P50876,P50895,P50993,P51148,P51151,P51153,P51157,P51159,P51160,P51164,P51168,P51170,P51172,P51451,P51511,P51512,P51571,P51572,P51575,P51582,P51636,P51648,P51654,P51674,P51677,P51679,P51681,P51684,P51685,P51686,P51693,P51787,P51788,P51790,P51793,P51795,P51797,P51798,P51800,P51801,P51805,P51809,P51810,P51811,P51828,P51841,P51864,P51884,P51888,P51993,P52429,P52569,P52757,P52797,P52798,P52799,P52803,P52824,P52848,P52849,P52961,P53007,P53041,P53420,P53618,P53680,P53708,P53794,P53801,P53805,P53816,P53985,P54219,P54289,P54578,P54707,P54709,P54710,P54753,P54756,P54760,P54762,P54764,P54792,P54826,P54829,P54849,P54851,P54852,P54855,P54920,P55001,P55008,P55011,P55017,P55040,P55042,P55061,P55064,P55073,P55081,P55082,P55083,P55085,P55087,P55160,P55259,P55268,P55283,P55285,P55286,P55287,P55289,P55290,P55291,P55344,P55808,P55851,P55899,P55916,P56134,P56159,P56180,P56199,P56373,P56378,P56539,P56557,P56559,P56589,P56696,P56703,P56704,P56705,P56706,P56746,P56747,P56748,P56749,P56750,P56817,P56856,P56880,P56937,P56962,P56975,P57054,P57057,P57087,P57088,P57103,P57105,P57679,P57727,P57729,P57735,P57738,P57739,P57764,P57771,P57773,P57789,P58107,P58170,P58173,P58180,P58181,P58182,P58215,P58335,P58397,P58400,P58401,P58418,P58511,P58549,P58550,P58658,P58743,P58753,P58872,P59025,P59044,P59190,P59510,P59533,P59534,P59535,P59536,P59537,P59538,P59539,P59540,P59541,P59542,P59543,P59544,P59551,P59646,P59768,P59773,P59901,P59922,P60033,P60059,P60201,P60323,P60468,P60507,P60508,P60509,P60602,P60606,P60763,P60852,P60880,P60893,P60953,P61006,P61009,P61018,P61020,P61073,P61077,P61165,P61224,P61266,P61421,P61550,P61565,P61566,P61567,P61570,P61586,P61619,P61647,P61803,P61952,P62070,P62079,P62166,P62330,P62491,P62683,P62684,P62685,P62699,P62745,P62834,P62952,P62955,P63000,P63010,P63027,P63092,P63096,P63098,P63126,P63128,P63130,P63145,P63146,P63165,P63211,P63215,P63218,P63244,P63252,P67812,P68104,P69849,P78310,P78324,P78325,P78333,P78334,P78348,P78352,P78356,P78357,P78363,P78369,P78380,P78381,P78382,P78383,P78395,P78410,P78423,P78504,P78508,P78509,P78524,P78536,P78552,P78562,P79483,P80370,P80511,P80723,P80748,P81274,P81408,P82251,P82279,P82987,P84095,P84157,P84996,P87889,P98073,P98095,P98153,P98155,P98160,P98161,P98164,P98172,P98187,P98194,P98196,P98198,Q00013,Q00325,Q00535,Q00536,Q00765,Q00973,Q00975,Q00LT1,Q01082,Q01113,Q01118,Q01151,Q01344,Q01362,Q01453,Q01484,Q01518,Q01628,Q01629,Q01638,Q01650,Q01668,Q01718,Q01726,Q01814,Q01844,Q01955,Q01959,Q01970,Q01973,Q01974,Q02094,Q02127,Q02156,Q02161,Q02223,Q02246,Q02297,Q02388,Q02413,Q02487,Q02505,Q02641,Q02643,Q02742,Q02763,Q02846,Q02978,Q03001,Q03113,Q03135,Q03167,Q03395,Q03405,Q03431,Q03518,Q03519,Q03692,Q03721,Q04609,Q04656,Q04671,Q04721,Q04759,Q04771,Q04826,Q04844,Q04900,Q04912,Q04941,Q05329,Q05397,Q05469,Q05586,Q05655,Q05707,Q05901,Q05940,Q05996,Q06055,Q06136,Q06187,Q06418,Q06432,Q06481,Q06495,Q06643,Q06787,Q06828,Q07000,Q07001,Q07011,Q07021,Q07065,Q07075,Q07092,Q07108,Q07157,Q07326,Q07444,Q07507,Q07654,Q07699,Q07812,Q07817,Q07820,Q07837,Q07912,Q07954,Q08116,Q08174,Q08209,Q08289,Q08334,Q08345,Q08357,Q08397,Q08431,Q08462,Q08477,Q08495,Q08499,Q08554,Q08629,Q08708,Q08722,Q08828,Q08AE8,Q08AI6,Q08ET2,Q09013,Q09160,Q09327,Q09328,Q09428,Q09470,Q0D2K0,Q0GE19,Q0JRZ9,Q0P670,Q0P6D2,Q0P6H9,Q0VAQ4,Q0VD83,Q0VDE8,Q0VDI3,Q10469,Q10471,Q10472,Q10588,Q10589,Q10981,Q11128,Q11130,Q11201,Q11203,Q11206,Q12767,Q12770,Q12772,Q12791,Q12797,Q12805,Q12809,Q12829,Q12836,Q12846,Q12851,Q12852,Q12860,Q12864,Q12866,Q12879,Q12884,Q12887,Q12891,Q12893,Q12907,Q12908,Q12912,Q12913,Q12918,Q12929,Q12934,Q12955,Q12959,Q12974,Q12981,Q12983,Q12999,Q13002,Q13003,Q13009,Q13017,Q13018,Q13021,Q13061,Q13093,Q13112,Q13113,Q13114,Q13145,Q13153,Q13183,Q13190,Q13224,Q13237,Q13241,Q13255,Q13258,Q13261,Q13277,Q13286,Q13291,Q13303,Q13304,Q13308,Q13319,Q13323,Q13324,Q13325,Q13326,Q13332,Q13336,Q13349,Q13361,Q13370,Q13410,Q13418,Q13421,Q13423,Q13424,Q13425,Q13433,Q13443,Q13444,Q13445,Q13449,Q13454,Q13464,Q13467,Q13474,Q13477,Q13478,Q13488,Q13491,Q13492,Q13496,Q13505,Q13507,Q13508,Q13515,Q13520,Q13530,Q13546,Q13557,Q13563,Q13564,Q13571,Q13574,Q13585,Q13586,Q13591,Q13606,Q13607,Q13613,Q13621,Q13634,Q13635,Q13639,Q13641,Q13651,Q13683,Q13685,Q13698,Q13702,Q13705,Q13724,Q13733,Q13740,Q13751,Q13753,Q13796,Q13797,Q13873,Q13884,Q13887,Q13936,Q13948,Q14002,Q14003,Q14028,Q14031,Q14050,Q14055,Q14088,Q14108,Q14112,Q14114,Q14118,Q14126,Q14156,Q14160,Q14162,Q14165,Q14168,Q14184,Q14210,Q14242,Q14244,Q14246,Q14247,Q14254,Q14289,Q14318,Q14330,Q14332,Q14344,Q14392,Q14416,Q14432,Q14435,Q14439,Q14451,Q14494,Q14500,Q14512,Q14515,Q14517,Q14524,Q14534,Q14542,Q14571,Q14573,Q14574,Q14623,Q14626,Q14627,Q14642,Q14643,Q14644,Q14654,Q14656,Q14678,Q14680,Q14699,Q14703,Q14714,Q14721,Q14722,Q14728,Q14739,Q14761,Q14773,Q14789,Q147U7,Q14802,Q14831,Q14832,Q14833,Q14849,Q14916,Q14940,Q14943,Q14952,Q14953,Q14954,Q14956,Q14957,Q14964,Q14973,Q14980,Q14982,Q14993,Q14BN4,Q14C86,Q14C87,Q14CN2,Q14CX5,Q14CZ8,Q14D04,Q14D33,Q14DG7,Q15005,Q15011,Q15012,Q15019,Q15029,Q15035,Q15040,Q15041,Q15043,Q15046,Q15049,Q15070,Q15077,Q15078,Q15084,Q15109,Q15116,Q15125,Q15139,Q15155,Q15165,Q15166,Q15223,Q15256,Q15262,Q15286,Q15303,Q15311,Q15334,Q15363,Q15375,Q15382,Q15388,Q15391,Q15392,Q15399,Q15413,Q15438,Q15464,Q15465,Q15506,Q15526,Q15546,Q15582,Q15599,Q15612,Q15614,Q15615,Q15617,Q15619,Q15620,Q15622,Q15629,Q15642,Q15669,Q15700,Q15722,Q15738,Q15743,Q15751,Q15758,Q15760,Q15761,Q15762,Q15768,Q15771,Q15800,Q15811,Q15822,Q15825,Q15835,Q15836,Q15842,Q15849,Q15858,Q15878,Q15884,Q15904,Q15907,Q16099,Q16206,Q16280,Q16281,Q16288,Q16322,Q16348,Q16363,Q16394,Q16445,Q16478,Q16512,Q16515,Q16518,Q16538,Q16549,Q16553,Q16558,Q16563,Q16570,Q16572,Q16581,Q16585,Q16586,Q16602,Q16610,Q16611,Q16617,Q16620,Q16623,Q16625,Q16635,Q16647,Q16651,Q16653,Q16655,Q16671,Q16706,Q16720,Q16739,Q16760,Q16787,Q16790,Q16799,Q16819,Q16820,Q16821,Q16827,Q16832,Q16842,Q16849,Q16850,Q16853,Q16873,Q16880,Q16891,Q17R55,Q17RQ9,Q17RW2,Q17RY6,Q18PE1,Q19T08,Q1AE95,Q1EHB4,Q1HG43,Q1HG44,Q24JP5,Q24JQ0,Q29718,Q29836,Q29865,Q29940,Q29960,Q29963,Q29974,Q29980,Q29983,Q2HXU8,Q2I0M4,Q2KHT4,Q2LD37,Q2M2E3,Q2M2I8,Q2M385,Q2M3C6,Q2M3G0,Q2M3M2,Q2M3R5,Q2M3T9,Q2MJR0,Q2PZI1,Q2QL34,Q2T9K0,Q2TAA5,Q2TAL6,Q2TBF2,Q2UY09,Q2VWP7,Q2VYF4,Q2WGJ8,Q2WGJ9,Q2Y0W8,Q30134,Q30154,Q30167,Q30201,Q31610,Q31612,Q32M45,Q32ZL2,Q3B7J2,Q3B7S5,Q3B7T3,Q3C1V0,Q3KNS1,Q3KNT9,Q3KNW5,Q3KPI0,Q3KQZ1,Q3KR37,Q3MIP1,Q3MIR4,Q3MIW9,Q3MIX3,Q3MJ16,Q3MUY2,Q3SXP7,Q3SXY7,Q3SXY8,Q3SY17,Q3SY77,Q3SYC2,Q3SYG4,Q3T906,Q3V5L5,Q3V6T2,Q3YBM2,Q3ZAQ7,Q3ZCQ3,Q3ZCQ8,Q401N2,Q494U1,Q494W8,Q495A1,Q495M3,Q495M9,Q495N2,Q495T6,Q495W5,Q496F6,Q496H8,Q496J9,Q49A17,Q49SQ1,Q4ADV7,Q4G0I0,Q4G0N0,Q4G0N8,Q4G0S4,Q4G0T1,Q4G148,Q4G1C9,Q4KMG0,Q4KMG9,Q4KMQ2,Q4KMZ8,Q4LDR2,Q4U2R8,Q4V9L6,Q4VC39,Q4VNC0,Q4VNC1,Q4VXA5,Q4VXF1,Q4ZG55,Q4ZIN3,Q4ZJI4,Q504Y0,Q50LG9,Q52LC2,Q52LD8,Q53EL9,Q53EP0,Q53EU6,Q53F39,Q53FP2,Q53FV1,Q53GD3,Q53GL0,Q53GQ0,Q53HI1,Q53R12,Q53RD9,Q53RT3,Q53RY4,Q53S58,Q53TN4,Q567V2,Q587I9,Q58DX5,Q58EX2,Q58HT5,Q5BIV9,Q5BJD5,Q5BJF2,Q5BJH2,Q5BJH7,Q5BKT4,Q5BKX6,Q5BKX8,Q5BVD1,Q5DID0,Q5DX21,Q5EB52,Q5EBL8,Q5FWE3,Q5FYA8,Q5GH70,Q5GH72,Q5GH73,Q5GH76,Q5GH77,Q5GJ75,Q5H8A4,Q5H8C1,Q5H943,Q5H9E4,Q5H9R4,Q5H9S7,Q5HYA8,Q5HYC2,Q5HYJ1,Q5HYL7,Q5I7T1,Q5IJ48,Q5J8M3,Q5J8X5,Q5JPE7,Q5JQS5,Q5JRA6,Q5JRM2,Q5JRS4,Q5JRV8,Q5JS13,Q5JTC6,Q5JTH9,Q5JTV8,Q5JUK3,Q5JW98,Q5JWF2,Q5JX69,Q5JX71,Q5JXA9,Q5JXX5,Q5JXX7,Q5JZY3,Q5K4E3,Q5K4L6,Q5KU26,Q5M7Z0,Q5M8T2,Q5MY95,Q5NUL3,Q5PT55,Q5QFB9,Q5QGT7,Q5QGZ9,Q5QJU3,Q5R3F8,Q5R3K3,Q5RGS3,Q5RI15,Q5SGD2,Q5SNT2,Q5SQ64,Q5SR56,Q5SRD1,Q5SRI9,Q5SRN2,Q5SSG8,Q5SV17,Q5SVS4,Q5SWH9,Q5SWX8,Q5SY80,Q5SYB0,Q5SZI1,Q5SZK8,Q5T0N5,Q5T0T0,Q5T197,Q5T1A1,Q5T1C6,Q5T1Q4,Q5T1S8,Q5T292,Q5T2D2,Q5T2E6,Q5T2T1,Q5T2W1,Q5T3F8,Q5T3U5,Q5T442,Q5T4D3,Q5T4F4,Q5T4J0,Q5T4S7,Q5T4T1,Q5T5C0,Q5T5W8,Q5T601,Q5T6L9,Q5T6X4,Q5T6X5,Q5T700,Q5T7M9,Q5T7P6,Q5T7P8,Q5T7R7,Q5T848,Q5T8D3,Q5T9L3,Q5T9Z0,Q5TA50,Q5TAH2,Q5TAT6,Q5TCQ9,Q5TEA6,Q5TF21,Q5TF39,Q5TFQ8,Q5TGI0,Q5TGU0,Q5TGY1,Q5TGZ0,Q5TH69,Q5TZ20,Q5TZJ5,Q5U3C3,Q5U4P2,Q5UAW9,Q5UCC4,Q5UE93,Q5VSG8,Q5VST6,Q5VST9,Q5VT66,Q5VT99,Q5VTH2,Q5VTJ3,Q5VTT2,Q5VTY9,Q5VU13,Q5VU36,Q5VU65,Q5VU97,Q5VUB5,Q5VUD6,Q5VUY2,Q5VV42,Q5VV43,Q5VV63,Q5VVB8,Q5VVP1,Q5VW32,Q5VW36,Q5VW38,Q5VWC8,Q5VWK5,Q5VWP3,Q5VWQ8,Q5VX71,Q5VXT5,Q5VXU1,Q5VXU3,Q5VY43,Q5VY80,Q5VYJ5,Q5VYP0,Q5VZ72,Q5VZ89,Q5VZI3,Q5VZK9,Q5VZR4,Q5VZY2,Q5W0B7,Q5W0N0,Q5W0Z9,Q5XG99,Q5XXA6,Q5Y7A7,Q5ZPR3,Q60682,Q629K1,Q63HM2,Q63HQ2,Q63HR2,Q63ZE4,Q641Q2,Q643R3,Q658N2,Q658P3,Q66K14,Q66K66,Q66K79,Q674R7,Q67FW5,Q685J3,Q687X5,Q68CJ9,Q68CP4,Q68CQ1,Q68CQ7,Q68CR1,Q68CR7,Q68D42,Q68D85,Q68DA7,Q68DD2,Q68DH5,Q68DU8,Q68DV7,Q68DX3,Q68EM7,Q68G75,Q69383,Q69384,Q695T7,Q69YG0,Q69YW2,Q69YZ2,Q6AI14,Q6AZY7,Q6BAA4,Q6DD88,Q6DHY5,Q6DKI7,Q6DN12,Q6DN14,Q6DN72,Q6DWJ6,Q6E213,Q6EIG7,Q6EMK4,Q6F5E8,Q6GMR7,Q6GPH6,Q6GPI1,Q6GTX8,Q6GV28,Q6H3X3,Q6HA08,Q6IA17,Q6IAA8,Q6IAN0,Q6IC98,Q6ICH7,Q6ICI0,Q6ICL7,Q6IEE7,Q6IEE8,Q6IEU7,Q6IEV9,Q6IEY1,Q6IEZ7,Q6IF00,Q6IF36,Q6IF42,Q6IF63,Q6IF82,Q6IF99,Q6IFG1,Q6IFH4,Q6IFN5,Q6IMN6,Q6IPM2,Q6IPX1,Q6IQ20,Q6IS24,Q6ISU1,Q6IWH7,Q6J4K2,Q6J9G0,Q6KCM7,Q6L9W6,Q6MZM0,Q6MZT1,Q6N022,Q6N075,Q6NSJ0,Q6NSJ5,Q6NT16,Q6NTF9,Q6NUI2,Q6NUI6,Q6NUJ2,Q6NUK1,Q6NUK4,Q6NUQ4,Q6NUS6,Q6NUS8,Q6NUT2,Q6NUT3,Q6NV75,Q6NVV3,Q6NW34,Q6NW40,Q6NXN4,Q6NXT4,Q6NXT6,Q6NZ63,Q6NZI2,Q6P0A1,Q6P0Q8,Q6P179,Q6P1A2,Q6P1J6,Q6P1K1,Q6P1M0,Q6P1Q0,Q6P1S2,Q6P2H8,Q6P499,Q6P4A7,Q6P4E1,Q6P4F1,Q6P4H8,Q6P4Q7,Q6P531,Q6P5S7,Q6P5W5,Q6P5X7,Q6P7N7,Q6P995,Q6P9A2,Q6P9B9,Q6P9F7,Q6P9G4,Q6P9H4,Q6PCB0,Q6PCB6,Q6PCB7,Q6PCB8,Q6PEX7,Q6PEY0,Q6PEY1,Q6PEZ8,Q6PFW1,Q6PHW0,Q6PI25,Q6PI73,Q6PI78,Q6PIL6,Q6PIS1,Q6PIU1,Q6PIU2,Q6PIV7,Q6PIZ9,Q6PJF5,Q6PJG9,Q6PJW8,Q6PK18,Q6PKC3,Q6PL45,Q6PML9,Q6PP77,Q6PRD1,Q6PXP3,Q6Q0C1,Q6Q4G3,Q6Q8B3,Q6QAJ8,Q6QHC5,Q6QNK2,Q6RW13,Q6S5L8,Q6T423,Q6T4P5,Q6T4R5,Q6TCH4,Q6TCH7,Q6U736,Q6U841,Q6UE05,Q6UQ28,Q6UVK1,Q6UVM3,Q6UVW9,Q6UVY6,Q6UW02,Q6UW56,Q6UW60,Q6UW68,Q6UW88,Q6UWB1,Q6UWB4,Q6UWD8,Q6UWF3,Q6UWH4,Q6UWH6,Q6UWI2,Q6UWI4,Q6UWJ1,Q6UWJ8,Q6UWL2,Q6UWL6,Q6UWM5,Q6UWM7,Q6UWM9,Q6UWN0,Q6UWN5,Q6UWP7,Q6UWR7,Q6UWT2,Q6UWU4,Q6UWV2,Q6UWV6,Q6UWV7,Q6UWW9,Q6UX01,Q6UX06,Q6UX15,Q6UX27,Q6UX34,Q6UX40,Q6UX41,Q6UX65,Q6UX68,Q6UX71,Q6UX72,Q6UX82,Q6UX98,Q6UXA7,Q6UXB3,Q6UXB4,Q6UXB8,Q6UXC1,Q6UXD1,Q6UXD5,Q6UXD7,Q6UXE8,Q6UXF1,Q6UXG2,Q6UXG3,Q6UXG8,Q6UXI7,Q6UXI9,Q6UXK2,Q6UXK5,Q6UXL0,Q6UXM1,Q6UXN7,Q6UXN8,Q6UXP3,Q6UXU4,Q6UXU6,Q6UXV0,Q6UXV1,Q6UXY1,Q6UXY8,Q6UXZ0,Q6UXZ3,Q6UXZ4,Q6UY09,Q6UY11,Q6UY14,Q6UY18,Q6V0I7,Q6V0L0,Q6V1P9,Q6W3E5,Q6W5P4,Q6X4W1,Q6XPR3,Q6XPS3,Q6XR72,Q6XUX3,Q6XYQ8,Q6XZB0,Q6Y1H2,Q6Y288,Q6Y2X3,Q6YBV0,Q6YHK3,Q6YI46,Q6ZMB0,Q6ZMB5,Q6ZMC9,Q6ZMD2,Q6ZMG9,Q6ZMH5,Q6ZMI3,Q6ZMJ2,Q6ZMP0,Q6ZMQ8,Q6ZMR5,Q6ZMT1,Q6ZMU5,Q6ZMZ0,Q6ZMZ3,Q6ZN44,Q6ZN68,Q6ZNA5,Q6ZNB6,Q6ZNB7,Q6ZNC8,Q6ZNI0,Q6ZNL6,Q6ZNR0,Q6ZP29,Q6ZP80,Q6ZPB5,Q6ZPD8,Q6ZPD9,Q6ZQN7,Q6ZQQ2,Q6ZRH7,Q6ZRI0,Q6ZRP7,Q6ZRR5,Q6ZS10,Q6ZS62,Q6ZS81,Q6ZS82,Q6ZSA7,Q6ZSJ9,Q6ZSM3,Q6ZSS7,Q6ZSY5,Q6ZSZ5,Q6ZT12,Q6ZT21,Q6ZT62,Q6ZT89,Q6ZTN6,Q6ZTQ4,Q6ZTR5,Q6ZU45,Q6ZU64,Q6ZU69,Q6ZUB0,Q6ZUB1,Q6ZUJ8,Q6ZUK4,Q6ZUT9,Q6ZUX7,Q6ZV29,Q6ZVE7,Q6ZVK1,Q6ZVL6,Q6ZVN8,Q6ZVX9,Q6ZW05,Q6ZWB6,Q6ZWE6,Q6ZWK4,Q6ZWK6,Q6ZWL3,Q6ZWT7,Q6ZXV5,Q70CQ3,Q70E73,Q70HW3,Q70JA7,Q70SY1,Q70UQ0,Q70Z44,Q717R9,Q71H61,Q71RC9,Q71RG4,Q71RH2,Q71RS6,Q75N90,Q75NE6,Q75T13,Q75V66,Q76EJ3,Q76KD6,Q76KP1,Q76M96,Q76MJ5,Q7KYR7,Q7KZI7,Q7KZN9,Q7L0J3,Q7L0Q8,Q7L0X0,Q7L1I2,Q7L1S5,Q7L1W4,Q7L211,Q7L311,Q7L4E1,Q7L4S7,Q7L513,Q7L591,Q7L5A8,Q7L5L3,Q7L5N7,Q7L5Y9,Q7L804,Q7L8C5,Q7L985,Q7LBE3,Q7LC44,Q7LDG7,Q7LDI9,Q7LFX5,Q7LGA3,Q7LGC8,Q7RTM1,Q7RTP0,Q7RTR8,Q7RTS5,Q7RTS6,Q7RTT9,Q7RTW8,Q7RTX0,Q7RTX1,Q7RTX7,Q7RTX9,Q7RTY0,Q7RTY1,Q7RTY8,Q7RTY9,Q7Z2D5,Q7Z2H8,Q7Z2K6,Q7Z2K8,Q7Z2Q7,Q7Z2W7,Q7Z304,Q7Z388,Q7Z3B0,Q7Z3B1,Q7Z3C6,Q7Z3D4,Q7Z3F1,Q7Z3J2,Q7Z3Q1,Q7Z3S7,Q7Z3T1,Q7Z402,Q7Z403,Q7Z404,Q7Z407,Q7Z408,Q7Z410,Q7Z412,Q7Z418,Q7Z419,Q7Z429,Q7Z434,Q7Z442,Q7Z443,Q7Z444,Q7Z449,Q7Z460,Q7Z4F1,Q7Z4I7,Q7Z4J2,Q7Z4L0,Q7Z4N2,Q7Z4T8,Q7Z4W1,Q7Z553,Q7Z5A4,Q7Z5A7,Q7Z5B4,Q7Z5H4,Q7Z5H5,Q7Z5J8,Q7Z5L7,Q7Z5M5,Q7Z5N4,Q7Z5R6,Q7Z5S9,Q7Z601,Q7Z602,Q7Z614,Q7Z692,Q7Z695,Q7Z698,Q7Z699,Q7Z6A9,Q7Z6B0,Q7Z6G3,Q7Z6J2,Q7Z6J4,Q7Z6J6,Q7Z6L0,Q7Z6M3,Q7Z6P3,Q7Z6W1,Q7Z739,Q7Z769,Q7Z7A4,Q7Z7B1,Q7Z7D3,Q7Z7G2,Q7Z7H3,Q7Z7H5,Q7Z7J7,Q7Z7M0,Q7Z7M1,Q7Z7M8,Q7Z7M9,Q7Z7N9,Q86SF2,Q86SJ2,Q86SJ6,Q86SK9,Q86SM5,Q86SM8,Q86SP6,Q86SQ3,Q86SQ4,Q86SQ6,Q86SR1,Q86SS6,Q86SU0,Q86T03,Q86T13,Q86T26,Q86T96,Q86TB3,Q86TG1,Q86TL2,Q86TM6,Q86TV6,Q86TY3,Q86U02,Q86U90,Q86UB9,Q86UD0,Q86UD3,Q86UD5,Q86UE4,Q86UE6,Q86UF1,Q86UF2,Q86UG4,Q86UK0,Q86UK5,Q86UL3,Q86UL8,Q86UN2,Q86UN3,Q86UP0,Q86UP2,Q86UP6,Q86UP9,Q86UQ4,Q86UQ5,Q86UR1,Q86UR5,Q86UT5,Q86UW1,Q86UW2,Q86UW8,Q86V24,Q86V35,Q86V40,Q86V85,Q86VB7,Q86VD7,Q86VD9,Q86VE9,Q86VF5,Q86VH4,Q86VH5,Q86VI4,Q86VL8,Q86VR2,Q86VR7,Q86VU5,Q86VW1,Q86VW2,Q86VY9,Q86VZ1,Q86VZ4,Q86VZ5,Q86W10,Q86W26,Q86W33,Q86W47,Q86W74,Q86WA9,Q86WB7,Q86WC4,Q86WI0,Q86WI1,Q86WK6,Q86WK7,Q86WK9,Q86WS5,Q86WV1,Q86WV6,Q86X19,Q86X27,Q86X29,Q86X52,Q86XA0,Q86XE3,Q86XJ0,Q86XK7,Q86XL3,Q86XM0,Q86XP1,Q86XQ3,Q86XR5,Q86XR7,Q86XS8,Q86XT9,Q86XX4,Q86Y07,Q86Y22,Q86Y34,Q86Y38,Q86Y39,Q86Y78,Q86Y82,Q86YC3,Q86YD1,Q86YD3,Q86YD5,Q86YJ5,Q86YJ7,Q86YL7,Q86YN1,Q86YR5,Q86YR6,Q86YS7,Q86YT5,Q86YT6,Q86YT9,Q86YW5,Q86Z14,Q8IU57,Q8IU68,Q8IU80,Q8IU89,Q8IU99,Q8IUA7,Q8IUC8,Q8IUD2,Q8IUH4,Q8IUH5,Q8IUH8,Q8IUK5,Q8IUL8,Q8IUN9,Q8IUR5,Q8IUS5,Q8IUW5,Q8IUX1,Q8IUX8,Q8IUY3,Q8IV01,Q8IV08,Q8IV16,Q8IV31,Q8IV45,Q8IV77,Q8IVB4,Q8IVE3,Q8IVF7,Q8IVI9,Q8IVJ1,Q8IVJ8,Q8IVL8,Q8IVM8,Q8IVN8,Q8IVP5,Q8IVQ6,Q8IVT5,Q8IVU1,Q8IVV8,Q8IVW8,Q8IVY1,Q8IW00,Q8IW52,Q8IW70,Q8IWA4,Q8IWA5,Q8IWB1,Q8IWB4,Q8IWB9,Q8IWD5,Q8IWE4,Q8IWK6,Q8IWL1,Q8IWL2,Q8IWR1,Q8IWT1,Q8IWT6,Q8IWU2,Q8IWU4,Q8IWV1,Q8IWV2,Q8IWX5,Q8IWY4,Q8IWY9,Q8IWZ6,Q8IX03,Q8IX05,Q8IX19,Q8IX94,Q8IXA5,Q8IXB3,Q8IXE1,Q8IXF9,Q8IXH8,Q8IXI1,Q8IXI2,Q8IXJ6,Q8IXK2,Q8IXM6,Q8IXS6,Q8IXS8,Q8IXU6,Q8IXX5,Q8IY17,Q8IY26,Q8IY33,Q8IY34,Q8IY49,Q8IY50,Q8IY95,Q8IYB5,Q8IYJ0,Q8IYJ3,Q8IYK8,Q8IYL9,Q8IYP9,Q8IYR6,Q8IYS0,Q8IYS2,Q8IYS5,Q8IYT2,Q8IYV9,Q8IZ07,Q8IZ08,Q8IZ52,Q8IZ57,Q8IZ96,Q8IZA0,Q8IZC6,Q8IZD2,Q8IZD6,Q8IZF0,Q8IZF2,Q8IZF3,Q8IZF4,Q8IZF5,Q8IZF6,Q8IZF7,Q8IZJ1,Q8IZJ3,Q8IZJ4,Q8IZK6,Q8IZM9,Q8IZN3,Q8IZP1,Q8IZP7,Q8IZP9,Q8IZQ1,Q8IZR5,Q8IZS7,Q8IZS8,Q8IZT8,Q8IZU8,Q8IZU9,Q8IZV2,Q8IZV5,Q8IZY2,Q8J025,Q8MH63,Q8N0U2,Q8N0U8,Q8N0V5,Q8N0W4,Q8N0W7,Q8N0Y3,Q8N0Y5,Q8N0Z9,Q8N109,Q8N111,Q8N112,Q8N114,Q8N118,Q8N126,Q8N127,Q8N130,Q8N131,Q8N138,Q8N139,Q8N144,Q8N146,Q8N148,Q8N149,Q8N158,Q8N162,Q8N1C3,Q8N1I0,Q8N1L4,Q8N1M1,Q8N1N0,Q8N1N2,Q8N1S5,Q8N1W1,Q8N201,Q8N205,Q8N271,Q8N292,Q8N2A8,Q8N2C7,Q8N2F6,Q8N2G4,Q8N2H4,Q8N2K0,Q8N2K1,Q8N2M4,Q8N2Q7,Q8N2S1,Q8N2U0,Q8N2U9,Q8N307,Q8N326,Q8N349,Q8N350,Q8N357,Q8N370,Q8N386,Q8N387,Q8N394,Q8N3E9,Q8N3F9,Q8N3G9,Q8N3I7,Q8N3J6,Q8N3R9,Q8N3T1,Q8N3T6,Q8N3Y3,Q8N3Y7,Q8N413,Q8N414,Q8N423,Q8N428,Q8N434,Q8N441,Q8N468,Q8N490,Q8N4A0,Q8N4C7,Q8N4C9,Q8N4F4,Q8N4F7,Q8N4H5,Q8N4K4,Q8N4L1,Q8N4L2,Q8N4L4,Q8N4M1,Q8N4S7,Q8N4S9,Q8N4T0,Q8N4T4,Q8N4V1,Q8N4V2,Q8N4W6,Q8N4Z0,Q8N511,Q8N539,Q8N565,Q8N5B7,Q8N5C1,Q8N5D6,Q8N5G0,Q8N5G2,Q8N5H7,Q8N5I2,Q8N5K1,Q8N5M9,Q8N5S1,Q8N5U1,Q8N5Y8,Q8N608,Q8N609,Q8N614,Q8N628,Q8N661,Q8N682,Q8N695,Q8N697,Q8N6C5,Q8N6D2,Q8N6F1,Q8N6F7,Q8N6G5,Q8N6G6,Q8N6I4,Q8N6K0,Q8N6L0,Q8N6L1,Q8N6L7,Q8N6M3,Q8N6P7,Q8N6Q1,Q8N6Q3,Q8N6R1,Q8N6S5,Q8N6U8,Q8N6Y1,Q8N6Y2,Q8N743,Q8N755,Q8N766,Q8N7C0,Q8N7C4,Q8N7C7,Q8N7J2,Q8N7P1,Q8N7P3,Q8N7S2,Q8N7S6,Q8N7X8,Q8N808,Q8N816,Q8N884,Q8N8D7,Q8N8F6,Q8N8F7,Q8N8J7,Q8N8N0,Q8N8Q1,Q8N8Q3,Q8N8Q8,Q8N8Q9,Q8N8R3,Q8N8V2,Q8N8V8,Q8N8Z6,Q8N912,Q8N944,Q8N966,Q8N967,Q8N9A8,Q8N9F0,Q8N9F7,Q8N9I0,Q8N9I5,Q8N9M5,Q8N9R8,Q8N9X5,Q8NA29,Q8NA58,Q8NAC3,Q8NAN2,Q8NAU1,Q8NB16,Q8NB49,Q8NB59,Q8NB66,Q8NBD8,Q8NBF6,Q8NBI2,Q8NBI5,Q8NBI6,Q8NBJ4,Q8NBJ9,Q8NBL3,Q8NBM4,Q8NBN3,Q8NBP5,Q8NBQ7,Q8NBR0,Q8NBS3,Q8NBT3,Q8NBU5,Q8NBV4,Q8NBV8,Q8NBW4,Q8NBZ7,Q8NC01,Q8NC24,Q8NC42,Q8NC44,Q8NC54,Q8NC56,Q8NC67,Q8NC96,Q8NCB2,Q8NCC5,Q8NCG5,Q8NCG7,Q8NCH0,Q8NCK7,Q8NCL4,Q8NCL8,Q8NCL9,Q8NCM2,Q8NCM8,Q8NCQ3,Q8NCR0,Q8NCR9,Q8NCS4,Q8NCS7,Q8NCT1,Q8NCU8,Q8NCW0,Q8NCW6,Q8ND23,Q8ND61,Q8ND76,Q8ND94,Q8NDA2,Q8NDB6,Q8NDV1,Q8NDV2,Q8NDX1,Q8NDX2,Q8NDY8,Q8NDZ6,Q8NE00,Q8NE01,Q8NE79,Q8NE86,Q8NEA5,Q8NEB5,Q8NEC5,Q8NEQ5,Q8NER1,Q8NER5,Q8NES3,Q8NET5,Q8NET6,Q8NET8,Q8NEU8,Q8NEW0,Q8NEW7,Q8NF37,Q8NF50,Q8NF91,Q8NFA0,Q8NFA2,Q8NFB2,Q8NFF2,Q8NFJ5,Q8NFJ6,Q8NFJ9,Q8NFK1,Q8NFL0,Q8NFM4,Q8NFM7,Q8NFN8,Q8NFP4,Q8NFQ8,Q8NFR3,Q8NFR9,Q8NFT2,Q8NFT8,Q8NFU0,Q8NFU1,Q8NFW1,Q8NFY4,Q8NFZ3,Q8NFZ4,Q8NFZ6,Q8NFZ8,Q8NG04,Q8NG11,Q8NG75,Q8NG76,Q8NG77,Q8NG78,Q8NG80,Q8NG81,Q8NG83,Q8NG84,Q8NG85,Q8NG92,Q8NG94,Q8NG95,Q8NG97,Q8NG98,Q8NG99,Q8NGA0,Q8NGA1,Q8NGA2,Q8NGA4,Q8NGA5,Q8NGA6,Q8NGA8,Q8NGB2,Q8NGB4,Q8NGB6,Q8NGB8,Q8NGB9,Q8NGC0,Q8NGC1,Q8NGC2,Q8NGC3,Q8NGC4,Q8NGC5,Q8NGC6,Q8NGC7,Q8NGC8,Q8NGC9,Q8NGD0,Q8NGD1,Q8NGD2,Q8NGD3,Q8NGD4,Q8NGD5,Q8NGE0,Q8NGE1,Q8NGE2,Q8NGE3,Q8NGE5,Q8NGE7,Q8NGE8,Q8NGE9,Q8NGF0,Q8NGF1,Q8NGF3,Q8NGF4,Q8NGF6,Q8NGF7,Q8NGF8,Q8NGF9,Q8NGG0,Q8NGG1,Q8NGG2,Q8NGG3,Q8NGG4,Q8NGG5,Q8NGG6,Q8NGG7,Q8NGG8,Q8NGH3,Q8NGH5,Q8NGH6,Q8NGH7,Q8NGH8,Q8NGH9,Q8NGI0,Q8NGI1,Q8NGI2,Q8NGI3,Q8NGI4,Q8NGI6,Q8NGI7,Q8NGI8,Q8NGI9,Q8NGJ0,Q8NGJ1,Q8NGJ2,Q8NGJ3,Q8NGJ4,Q8NGJ5,Q8NGJ6,Q8NGJ7,Q8NGJ8,Q8NGJ9,Q8NGK0,Q8NGK1,Q8NGK2,Q8NGK3,Q8NGK4,Q8NGK5,Q8NGK6,Q8NGK9,Q8NGL0,Q8NGL1,Q8NGL2,Q8NGL3,Q8NGL4,Q8NGL6,Q8NGL7,Q8NGL9,Q8NGM1,Q8NGM8,Q8NGM9,Q8NGN0,Q8NGN1,Q8NGN2,Q8NGN3,Q8NGN4,Q8NGN5,Q8NGN6,Q8NGN7,Q8NGN8,Q8NGP0,Q8NGP2,Q8NGP3,Q8NGP4,Q8NGP6,Q8NGP8,Q8NGP9,Q8NGQ1,Q8NGQ2,Q8NGQ3,Q8NGQ4,Q8NGQ5,Q8NGQ6,Q8NGR1,Q8NGR2,Q8NGR3,Q8NGR4,Q8NGR5,Q8NGR6,Q8NGR8,Q8NGR9,Q8NGS0,Q8NGS1,Q8NGS2,Q8NGS3,Q8NGS4,Q8NGS5,Q8NGS6,Q8NGS7,Q8NGS8,Q8NGS9,Q8NGT0,Q8NGT1,Q8NGT2,Q8NGT5,Q8NGT7,Q8NGT9,Q8NGU1,Q8NGU2,Q8NGU4,Q8NGU9,Q8NGV0,Q8NGV5,Q8NGV6,Q8NGV7,Q8NGW1,Q8NGW6,Q8NGX0,Q8NGX1,Q8NGX2,Q8NGX3,Q8NGX5,Q8NGX6,Q8NGX8,Q8NGX9,Q8NGY0,Q8NGY1,Q8NGY2,Q8NGY3,Q8NGY5,Q8NGY6,Q8NGY7,Q8NGY9,Q8NGZ0,Q8NGZ2,Q8NGZ3,Q8NGZ4,Q8NGZ5,Q8NGZ6,Q8NGZ9,Q8NH00,Q8NH01,Q8NH02,Q8NH03,Q8NH04,Q8NH05,Q8NH06,Q8NH07,Q8NH08,Q8NH09,Q8NH10,Q8NH16,Q8NH18,Q8NH19,Q8NH21,Q8NH37,Q8NH40,Q8NH41,Q8NH42,Q8NH43,Q8NH48,Q8NH49,Q8NH50,Q8NH51,Q8NH53,Q8NH54,Q8NH55,Q8NH56,Q8NH57,Q8NH59,Q8NH60,Q8NH61,Q8NH63,Q8NH64,Q8NH67,Q8NH69,Q8NH70,Q8NH72,Q8NH73,Q8NH74,Q8NH76,Q8NH79,Q8NH80,Q8NH81,Q8NH83,Q8NH85,Q8NH87,Q8NH89,Q8NH90,Q8NH92,Q8NH93,Q8NH94,Q8NH95,Q8NHA4,Q8NHA6,Q8NHA8,Q8NHB1,Q8NHB7,Q8NHB8,Q8NHC4,Q8NHC5,Q8NHC6,Q8NHC7,Q8NHC8,Q8NHE4,Q8NHG7,Q8NHG8,Q8NHH9,Q8NHJ6,Q8NHK3,Q8NHL6,Q8NHP6,Q8NHS1,Q8NHS3,Q8NHU3,Q8NHV5,Q8NHX9,Q8NHY0,Q8NHY3,Q8NI17,Q8NI28,Q8NI32,Q8NI35,Q8TAA9,Q8TAB3,Q8TAC9,Q8TAD4,Q8TAE7,Q8TAF8,Q8TAG6,Q8TAI7,Q8TAM2,Q8TAQ9,Q8TAV3,Q8TAV4,Q8TAX9,Q8TAZ6,Q8TB36,Q8TB61,Q8TB68,Q8TB96,Q8TBA6,Q8TBB6,Q8TBE1,Q8TBE3,Q8TBE7,Q8TBF5,Q8TBG9,Q8TBJ4,Q8TBM7,Q8TBM8,Q8TBP5,Q8TBP6,Q8TBQ9,Q8TBR7,Q8TC12,Q8TC26,Q8TC27,Q8TC36,Q8TC41,Q8TC92,Q8TCB6,Q8TCC7,Q8TCG1,Q8TCG5,Q8TCJ2,Q8TCP9,Q8TCQ1,Q8TCT0,Q8TCT6,Q8TCT7,Q8TCT8,Q8TCT9,Q8TCU3,Q8TCU5,Q8TCU6,Q8TCW7,Q8TCW9,Q8TCY5,Q8TCZ2,Q8TD07,Q8TD20,Q8TD22,Q8TD43,Q8TD46,Q8TD84,Q8TDB4,Q8TDB8,Q8TDD5,Q8TDF5,Q8TDF6,Q8TDI7,Q8TDI8,Q8TDM5,Q8TDM6,Q8TDN1,Q8TDN2,Q8TDN7,Q8TDQ0,Q8TDQ1,Q8TDS4,Q8TDS5,Q8TDS7,Q8TDT2,Q8TDU5,Q8TDU6,Q8TDU9,Q8TDV0,Q8TDV2,Q8TDV5,Q8TDW0,Q8TDW4,Q8TDW5,Q8TDW7,Q8TDX6,Q8TDX9,Q8TDY8,Q8TE23,Q8TE54,Q8TE56,Q8TE57,Q8TE58,Q8TE59,Q8TE60,Q8TEB7,Q8TEB9,Q8TED1,Q8TED4,Q8TEF2,Q8TEH3,Q8TEM1,Q8TEQ8,Q8TEU7,Q8TEW0,Q8TEY5,Q8TEZ7,Q8TF08,Q8TF62,Q8TF66,Q8TF71,Q8TF72,Q8WTQ7,Q8WTR4,Q8WTT0,Q8WTV0,Q8WTX9,Q8WU08,Q8WU17,Q8WU67,Q8WUD1,Q8WUD6,Q8WUG5,Q8WUH6,Q8WUJ3,Q8WUM9,Q8WUS8,Q8WUT4,Q8WUT9,Q8WUU8,Q8WUX1,Q8WUY8,Q8WV15,Q8WV19,Q8WV28,Q8WV48,Q8WV83,Q8WVE6,Q8WVE7,Q8WVF1,Q8WVF2,Q8WVH0,Q8WVI0,Q8WVN6,Q8WVP7,Q8WVQ1,Q8WVV5,Q8WVX3,Q8WVX9,Q8WVZ1,Q8WVZ7,Q8WW22,Q8WW34,Q8WW43,Q8WW52,Q8WW62,Q8WWA0,Q8WWA1,Q8WWB7,Q8WWF3,Q8WWF5,Q8WWG1,Q8WWG9,Q8WWI5,Q8WWL2,Q8WWM7,Q8WWN8,Q8WWN9,Q8WWP7,Q8WWQ0,Q8WWQ2,Q8WWQ8,Q8WWR8,Q8WWT9,Q8WWU5,Q8WWV6,Q8WWX8,Q8WWY8,Q8WWZ4,Q8WWZ7,Q8WXA8,Q8WXB1,Q8WXD0,Q8WXF7,Q8WXG6,Q8WXG9,Q8WXH0,Q8WXH2,Q8WXH6,Q8WXI7,Q8WXI8,Q8WXS4,Q8WXS5,Q8WXS8,Q8WXX5,Q8WY07,Q8WY21,Q8WY22,Q8WY41,Q8WY64,Q8WY98,Q8WYK1,Q8WYR1,Q8WZ04,Q8WZ33,Q8WZ55,Q8WZ59,Q8WZ71,Q8WZ73,Q8WZ84,Q8WZ92,Q8WZ94,Q8WZA1,Q8WZA6,Q902F8,Q902F9,Q92185,Q92186,Q92187,Q92478,Q92482,Q92485,Q92504,Q92508,Q92521,Q92523,Q92535,Q92536,Q92537,Q92542,Q92544,Q92545,Q92556,Q92563,Q92581,Q92597,Q92604,Q92611,Q92617,Q92619,Q92626,Q92629,Q92633,Q92637,Q92643,Q92673,Q92681,Q92685,Q92692,Q92729,Q92730,Q92736,Q92737,Q92743,Q92752,Q92781,Q92797,Q92806,Q92813,Q92819,Q92820,Q92823,Q92824,Q92832,Q92835,Q92838,Q92839,Q92847,Q92854,Q92859,Q92887,Q92896,Q92903,Q92911,Q92928,Q92930,Q92932,Q92935,Q92952,Q92953,Q92956,Q92959,Q92963,Q92968,Q92974,Q92982,Q93033,Q93038,Q93050,Q93052,Q93063,Q93070,Q93084,Q93086,Q93096,Q93097,Q93098,Q93100,Q95365,Q95460,Q95604,Q95IE3,Q969E2,Q969F0,Q969F2,Q969F8,Q969G9,Q969I6,Q969K3,Q969K7,Q969L2,Q969M2,Q969M3,Q969N2,Q969N4,Q969P0,Q969R2,Q969S0,Q969S6,Q969V1,Q969V3,Q969V5,Q969V6,Q969W0,Q969W1,Q969W9,Q969X1,Q969X2,Q969X5,Q969Z4,Q96A11,Q96A25,Q96A26,Q96A28,Q96A29,Q96A33,Q96A46,Q96A49,Q96A54,Q96A57,Q96A59,Q96A83,Q96A84,Q96AA3,Q96AC1,Q96AD5,Q96AE7,Q96AG3,Q96AG4,Q96AJ9,Q96AM1,Q96AN5,Q96AP7,Q96AQ2,Q96AQ8,Q96AW1,Q96AZ1,Q96B21,Q96B33,Q96B42,Q96B67,Q96B77,Q96B86,Q96B96,Q96BA8,Q96BD0,Q96BF3,Q96BI1,Q96BI3,Q96BM0,Q96BS2,Q96BY9,Q96BZ4,Q96BZ9,Q96C03,Q96C19,Q96C24,Q96CC6,Q96CE8,Q96CG8,Q96CH1,Q96CP6,Q96CP7,Q96CQ1,Q96CU9,Q96CW1,Q96CW9,Q96CX2,Q96D05,Q96D21,Q96D31,Q96D42,Q96D53,Q96D59,Q96D71,Q96D96,Q96DA2,Q96DA6,Q96DB9,Q96DC7,Q96DD7,Q96DL1,Q96DS6,Q96DU3,Q96DW6,Q96DX8,Q96DZ5,Q96DZ7,Q96DZ9,Q96E16,Q96E17,Q96E22,Q96E35,Q96E52,Q96E66,Q96E93,Q96EC8,Q96EP9,Q96ER9,Q96ES6,Q96ET8,Q96EU7,Q96EV8,Q96EX1,Q96EX2,Q96EY1,Q96F05,Q96F15,Q96F25,Q96F46,Q96F81,Q96FB5,Q96FC9,Q96FE5,Q96FE7,Q96FL8,Q96FL9,Q96FM1,Q96FN4,Q96FT7,Q96FV3,Q96FX8,Q96FZ5,Q96G23,Q96G30,Q96G79,Q96G91,Q96G97,Q96GC9,Q96GD0,Q96GE9,Q96GF1,Q96GL9,Q96GM1,Q96GP6,Q96GQ5,Q96GR2,Q96GR4,Q96GS6,Q96GW7,Q96GX1,Q96GZ6,Q96H15,Q96H72,Q96H78,Q96H96,Q96HA1,Q96HA4,Q96HA9,Q96HD1,Q96HD9,Q96HE8,Q96HG1,Q96HH4,Q96HH6,Q96HJ5,Q96HP8,Q96HR9,Q96HS1,Q96HU8,Q96HV5,Q96I34,Q96I36,Q96I45,Q96I82,Q96ID5,Q96IF1,Q96IK0,Q96IQ7,Q96IV6,Q96IW7,Q96IX5,Q96IZ2,Q96J02,Q96J42,Q96J65,Q96J66,Q96J84,Q96J86,Q96JA1,Q96JA4,Q96JB6,Q96JF0,Q96JI7,Q96JJ6,Q96JJ7,Q96JN2,Q96JP9,Q96JQ0,Q96JQ2,Q96JQ5,Q96JT2,Q96JW4,Q96JX3,Q96K12,Q96K19,Q96K37,Q96K49,Q96K78,Q96KA5,Q96KC8,Q96KF7,Q96KG7,Q96KJ4,Q96KK3,Q96KK4,Q96KN9,Q96KR6,Q96KT7,Q96KV6,Q96KW2,Q96L08,Q96L33,Q96L42,Q96L46,Q96L58,Q96LA5,Q96LA6,Q96LA9,Q96LB0,Q96LB1,Q96LB2,Q96LC7,Q96LD1,Q96LL3,Q96LR9,Q96LT4,Q96LU7,Q96LZ7,Q96M19,Q96MC6,Q96MF2,Q96MH6,Q96MM7,Q96MP8,Q96MS0,Q96MT1,Q96MU8,Q96MV1,Q96MV8,Q96MX0,Q96N19,Q96N35,Q96N66,Q96N68,Q96N87,Q96N96,Q96NA8,Q96NB2,Q96ND0,Q96NE9,Q96NF6,Q96NI6,Q96NL1,Q96NR3,Q96NT5,Q96NU0,Q96NW4,Q96NX5,Q96NY7,Q96NY8,Q96P31,Q96P44,Q96P48,Q96P56,Q96P65,Q96P66,Q96P67,Q96P68,Q96P69,Q96P88,Q96PB1,Q96PB8,Q96PD2,Q96PD6,Q96PD7,Q96PE1,Q96PE3,Q96PE5,Q96PG1,Q96PG2,Q96PH1,Q96PJ5,Q96PL2,Q96PL5,Q96PN6,Q96PQ0,Q96PQ1,Q96PR1,Q96PS6,Q96PS8,Q96PX8,Q96PZ7,Q96Q04,Q96Q06,Q96Q45,Q96Q80,Q96Q91,Q96QA5,Q96QD8,Q96QE2,Q96QE4,Q96QG7,Q96QI5,Q96QK8,Q96QS1,Q96QT4,Q96QU1,Q96QV1,Q96QZ0,Q96QZ7,Q96R08,Q96R09,Q96R27,Q96R28,Q96R30,Q96R45,Q96R47,Q96R48,Q96R54,Q96R67,Q96R69,Q96R72,Q96R84,Q96RA2,Q96RB7,Q96RC9,Q96RD0,Q96RD1,Q96RD2,Q96RD3,Q96RD6,Q96RD7,Q96RD9,Q96RF0,Q96RI0,Q96RI8,Q96RI9,Q96RJ0,Q96RJ3,Q96RK4,Q96RL6,Q96RN1,Q96RP7,Q96RP8,Q96RQ1,Q96RT1,Q96RT6,Q96RU3,Q96RV3,Q96RW7,Q96S06,Q96S21,Q96S37,Q96S52,Q96S59,Q96S66,Q96S79,Q96S86,Q96S97,Q96SA4,Q96SB3,Q96SE0,Q96SJ8,Q96SK2,Q96SL1,Q96SN7,Q96T49,Q96T52,Q96T53,Q96T54,Q96T55,Q96T83,Q96TA0,Q96TA2,Q96TC7,Q99062,Q99075,Q99102,Q99217,Q99218,Q99250,Q99418,Q99435,Q99437,Q99442,Q99445,Q99463,Q99466,Q99467,Q99469,Q99497,Q99500,Q99519,Q99523,Q99527,Q99541,Q99542,Q99547,Q99569,Q99571,Q99572,Q99578,Q99595,Q99603,Q99623,Q99624,Q99643,Q99645,Q99650,Q99653,Q99665,Q99677,Q99678,Q99679,Q99680,Q99689,Q99705,Q99706,Q99712,Q99715,Q99720,Q99726,Q99732,Q99735,Q99747,Q99755,Q99758,Q99788,Q99795,Q99797,Q99805,Q99808,Q99828,Q99829,Q99835,Q99884,Q99928,Q99941,Q99942,Q99943,Q99946,Q99965,Q99983,Q99999,Q9BPV8,Q9BPX6,Q9BPZ7,Q9BQ16,Q9BQ31,Q9BQ49,Q9BQ51,Q9BQA9,Q9BQB4,Q9BQB6,Q9BQD7,Q9BQE4,Q9BQG1,Q9BQI0,Q9BQI5,Q9BQI7,Q9BQJ4,Q9BQL6,Q9BQQ7,Q9BQS2,Q9BQS7,Q9BQT8,Q9BQT9,Q9BR10,Q9BR26,Q9BR39,Q9BRB3,Q9BRC7,Q9BRI3,Q9BRK0,Q9BRK3,Q9BRK5,Q9BRL7,Q9BRN9,Q9BRQ5,Q9BRQ8,Q9BRR3,Q9BRT3,Q9BRV3,Q9BRY0,Q9BS91,Q9BSA4,Q9BSA9,Q9BSE2,Q9BSE4,Q9BSF0,Q9BSG5,Q9BSJ5,Q9BSJ8,Q9BSK0,Q9BSK2,Q9BSN7,Q9BSR8,Q9BSW7,Q9BT22,Q9BT67,Q9BT76,Q9BT88,Q9BT92,Q9BTD3,Q9BTN0,Q9BTU6,Q9BTV4,Q9BTW9,Q9BTX1,Q9BTX3,Q9BU23,Q9BU79,Q9BUB7,Q9BUD6,Q9BUF7,Q9BUJ0,Q9BUL8,Q9BUM1,Q9BUN8,Q9BUR5,Q9BUZ4,Q9BV10,Q9BV23,Q9BV35,Q9BV40,Q9BV57,Q9BV81,Q9BV87,Q9BVA6,Q9BVC6,Q9BVG9,Q9BVH7,Q9BVI4,Q9BVK2,Q9BVK6,Q9BVK8,Q9BVT8,Q9BVV7,Q9BVV8,Q9BVW6,Q9BVX2,Q9BW60,Q9BW72,Q9BWL3,Q9BWM7,Q9BWQ6,Q9BWQ8,Q9BWV1,Q9BWV2,Q9BX59,Q9BX66,Q9BX67,Q9BX73,Q9BX74,Q9BX79,Q9BX84,Q9BX95,Q9BX97,Q9BXA5,Q9BXB1,Q9BXC0,Q9BXC1,Q9BXC9,Q9BXE9,Q9BXI2,Q9BXJ7,Q9BXJ8,Q9BXK5,Q9BXM0,Q9BXM7,Q9BXN1,Q9BXN2,Q9BXP2,Q9BXR3,Q9BXR5,Q9BXS0,Q9BXS4,Q9BXS9,Q9BXT2,Q9BXU9,Q9BXX0,Q9BY07,Q9BY08,Q9BY10,Q9BY11,Q9BY14,Q9BY15,Q9BY19,Q9BY21,Q9BY50,Q9BY64,Q9BY67,Q9BY71,Q9BY78,Q9BY79,Q9BYC5,Q9BYE2,Q9BYE9,Q9BYF1,Q9BYG0,Q9BYG4,Q9BYG5,Q9BYG8,Q9BYH1,Q9BYI3,Q9BYJ0,Q9BYT1,Q9BYT9,Q9BYW1,Q9BZ11,Q9BZ67,Q9BZ76,Q9BZ97,Q9BZA7,Q9BZA8,Q9BZC7,Q9BZD2,Q9BZD6,Q9BZD7,Q9BZF2,Q9BZF3,Q9BZG2,Q9BZG9,Q9BZH6,Q9BZJ4,Q9BZJ6,Q9BZJ7,Q9BZJ8,Q9BZL3,Q9BZL6,Q9BZM2,Q9BZM4,Q9BZM5,Q9BZM6,Q9BZR6,Q9BZV2,Q9BZV3,Q9BZW2,Q9BZW4,Q9BZW5,Q9BZW8,Q9BZZ2,Q9C004,Q9C091,Q9C0A0,Q9C0B5,Q9C0B7,Q9C0C4,Q9C0D9,Q9C0E8,Q9C0H2,Q9C0I4,Q9C0J1,Q9C0K0,Q9C0K1,Q9GIP4,Q9GIY3,Q9GZK3,Q9GZK4,Q9GZK6,Q9GZK7,Q9GZM5,Q9GZM6,Q9GZN0,Q9GZN6,Q9GZP1,Q9GZP7,Q9GZP9,Q9GZQ4,Q9GZQ6,Q9GZR5,Q9GZS9,Q9GZT5,Q9GZT6,Q9GZU1,Q9GZU3,Q9GZU5,Q9GZV3,Q9GZV5,Q9GZV7,Q9GZW8,Q9GZX3,Q9GZY4,Q9GZY6,Q9GZY8,Q9GZZ6,Q9GZZ7,Q9H013,Q9H015,Q9H061,Q9H0A3,Q9H0C2,Q9H0C3,Q9H0F7,Q9H0H0,Q9H0H5,Q9H0M0,Q9H0Q3,Q9H0R3,Q9H0U3,Q9H0U4,Q9H0V1,Q9H0V9,Q9H0X4,Q9H0X9,Q9H115,Q9H156,Q9H158,Q9H159,Q9H172,Q9H190,Q9H195,Q9H1B5,Q9H1C0,Q9H1C3,Q9H1C4,Q9H1C7,Q9H1D0,Q9H1E5,Q9H1J5,Q9H1J7,Q9H1K0,Q9H1K4,Q9H1N7,Q9H1P3,Q9H1R2,Q9H1U4,Q9H1U9,Q9H1V8,Q9H1X3,Q9H1Y3,Q9H1Z8,Q9H1Z9,Q9H205,Q9H207,Q9H208,Q9H209,Q9H210,Q9H221,Q9H222,Q9H223,Q9H228,Q9H237,Q9H239,Q9H244,Q9H251,Q9H252,Q9H255,Q9H295,Q9H2A7,Q9H2A9,Q9H2B2,Q9H2B4,Q9H2C2,Q9H2C5,Q9H2C8,Q9H2D1,Q9H2E6,Q9H2F3,Q9H2H9,Q9H2J7,Q9H2K8,Q9H2L4,Q9H2Q1,Q9H2S1,Q9H2S6,Q9H2U9,Q9H2V7,Q9H2W1,Q9H2X3,Q9H2X8,Q9H2X9,Q9H2Y9,Q9H300,Q9H306,Q9H310,Q9H313,Q9H324,Q9H330,Q9H339,Q9H340,Q9H341,Q9H342,Q9H343,Q9H344,Q9H346,Q9H3H5,Q9H3K2,Q9H3M0,Q9H3N1,Q9H3N8,Q9H3Q3,Q9H3R1,Q9H3R2,Q9H3S1,Q9H3S3,Q9H3S5,Q9H3T2,Q9H3T3,Q9H3U1,Q9H3U5,Q9H3U7,Q9H3V2,Q9H3W5,Q9H3Z4,Q9H400,Q9H427,Q9H461,Q9H490,Q9H496,Q9H4A6,Q9H4A9,Q9H4B8,Q9H4D0,Q9H4E5,Q9H4E7,Q9H4F1,Q9H4F8,Q9H4I9,Q9H4L5,Q9H4M9,Q9H553,Q9H598,Q9H5I5,Q9H5J4,Q9H5K3,Q9H5V8,Q9H5Y7,Q9H665,Q9H6A9,Q9H6B4,Q9H6D3,Q9H6D8,Q9H6F2,Q9H6H4,Q9H6L2,Q9H6L5,Q9H6Q3,Q9H6R6,Q9H6U8,Q9H6X2,Q9H6X4,Q9H6Y7,Q9H720,Q9H756,Q9H799,Q9H7D0,Q9H7F0,Q9H7F4,Q9H7M9,Q9H7T0,Q9H7V2,Q9H7X2,Q9H7Z7,Q9H813,Q9H819,Q9H841,Q9H8J5,Q9H8L6,Q9H8M5,Q9H8M9,Q9H8P0,Q9H8T0,Q9H8X9,Q9H902,Q9H920,Q9H936,Q9H9B4,Q9H9K5,Q9H9P2,Q9H9S3,Q9H9S5,Q9H9V4,Q9HA72,Q9HA82,Q9HAB3,Q9HAP6,Q9HAR2,Q9HAS3,Q9HAT1,Q9HAU4,Q9HAV5,Q9HAW7,Q9HAW8,Q9HAW9,Q9HB03,Q9HB14,Q9HB15,Q9HB19,Q9HB21,Q9HB29,Q9HB63,Q9HB89,Q9HBA0,Q9HBB8,Q9HBE5,Q9HBG4,Q9HBG7,Q9HBH0,Q9HBI0,Q9HBI1,Q9HBI6,Q9HBJ8,Q9HBL6,Q9HBL7,Q9HBM0,Q9HBR0,Q9HBT6,Q9HBU9,Q9HBV1,Q9HBV2,Q9HBW0,Q9HBW1,Q9HBW9,Q9HBX8,Q9HBX9,Q9HBY0,Q9HC07,Q9HC10,Q9HC21,Q9HC24,Q9HC29,Q9HC47,Q9HC56,Q9HC58,Q9HC73,Q9HC97,Q9HCB6,Q9HCC8,Q9HCE7,Q9HCE9,Q9HCF6,Q9HCH5,Q9HCI5,Q9HCJ1,Q9HCJ2,Q9HCK4,Q9HCL0,Q9HCL2,Q9HCM2,Q9HCM3,Q9HCN3,Q9HCN6,Q9HCP6,Q9HCQ5,Q9HCS2,Q9HCU0,Q9HCU4,Q9HCU5,Q9HCX4,Q9HD20,Q9HD23,Q9HD36,Q9HD43,Q9HD45,Q9HD47,Q9HD67,Q9HD87,Q9HDB5,Q9HDB9,Q9HDC5,Q9HDC9,Q9HDD0,Q9N2J8,Q9N2K0,Q9NNX6,Q9NNZ3,Q9NP31,Q9NP58,Q9NP59,Q9NP60,Q9NP70,Q9NP72,Q9NP78,Q9NP80,Q9NP84,Q9NP85,Q9NP90,Q9NP91,Q9NP94,Q9NP99,Q9NPA0,Q9NPA1,Q9NPA2,Q9NPB0,Q9NPB3,Q9NPB6,Q9NPB9,Q9NPC1,Q9NPC2,Q9NPC4,Q9NPD5,Q9NPD7,Q9NPE6,Q9NPF0,Q9NPF2,Q9NPF8,Q9NPG1,Q9NPG4,Q9NPG8,Q9NPH3,Q9NPH5,Q9NPI0,Q9NPI9,Q9NPL8,Q9NPQ8,Q9NPR2,Q9NPR9,Q9NPU4,Q9NPY3,Q9NPZ5,Q9NQ11,Q9NQ25,Q9NQ34,Q9NQ40,Q9NQ60,Q9NQ76,Q9NQ79,Q9NQ84,Q9NQ90,Q9NQA5,Q9NQC3,Q9NQC7,Q9NQG1,Q9NQG6,Q9NQN1,Q9NQQ7,Q9NQR9,Q9NQS3,Q9NQS5,Q9NQW8,Q9NQX3,Q9NQX5,Q9NQX7,Q9NQZ6,Q9NQZ7,Q9NR16,Q9NR22,Q9NR34,Q9NR61,Q9NR71,Q9NR77,Q9NR80,Q9NR82,Q9NR96,Q9NR97,Q9NRA1,Q9NRA2,Q9NRB3,Q9NRC1,Q9NRD8,Q9NRD9,Q9NRE1,Q9NRJ7,Q9NRK6,Q9NRM0,Q9NRM1,Q9NRM6,Q9NRP0,Q9NRQ2,Q9NRQ5,Q9NRR2,Q9NRR3,Q9NRR6,Q9NRR8,Q9NRS4,Q9NRU3,Q9NRX5,Q9NRX6,Q9NRY4,Q9NRY6,Q9NRY7,Q9NRZ5,Q9NRZ7,Q9NS00,Q9NS40,Q9NS61,Q9NS62,Q9NS64,Q9NS66,Q9NS67,Q9NS68,Q9NS69,Q9NS75,Q9NS82,Q9NS84,Q9NS86,Q9NS93,Q9NSA0,Q9NSA2,Q9NSB8,Q9NSC7,Q9NSD5,Q9NSD7,Q9NSE7,Q9NSI5,Q9NSK7,Q9NST1,Q9NT22,Q9NT68,Q9NT99,Q9NTG1,Q9NTI2,Q9NTJ5,Q9NTN3,Q9NTN9,Q9NTQ9,Q9NTU4,Q9NU53,Q9NUB4,Q9NUD9,Q9NUE0,Q9NUH8,Q9NUM3,Q9NUM4,Q9NUN5,Q9NUN7,Q9NUP9,Q9NUQ2,Q9NUR3,Q9NUT2,Q9NUV7,Q9NV12,Q9NV29,Q9NV58,Q9NV64,Q9NV92,Q9NV96,Q9NVA4,Q9NVC3,Q9NVD7,Q9NVI7,Q9NVM1,Q9NVV0,Q9NVV5,Q9NVZ3,Q9NW15,Q9NW97,Q9NWC5,Q9NWD8,Q9NWF4,Q9NWH2,Q9NWH9,Q9NWM0,Q9NWQ8,Q9NWR8,Q9NWS6,Q9NWW5,Q9NWW9,Q9NX00,Q9NX14,Q9NX47,Q9NX52,Q9NX61,Q9NX62,Q9NX76,Q9NX77,Q9NX78,Q9NX94,Q9NX95,Q9NXB9,Q9NXE4,Q9NXF1,Q9NXF8,Q9NXG6,Q9NXH8,Q9NXI6,Q9NXJ0,Q9NXK6,Q9NXL6,Q9NXS2,Q9NXW2,Q9NY15,Q9NY25,Q9NY26,Q9NY28,Q9NY35,Q9NY37,Q9NY46,Q9NY47,Q9NY59,Q9NY64,Q9NY72,Q9NY84,Q9NY91,Q9NY97,Q9NY99,Q9NYA1,Q9NYB5,Q9NYG2,Q9NYG8,Q9NYI0,Q9NYJ7,Q9NYJ8,Q9NYK1,Q9NYL4,Q9NYM4,Q9NYM9,Q9NYP7,Q9NYQ6,Q9NYQ7,Q9NYQ8,Q9NYR8,Q9NYT0,Q9NYV7,Q9NYV8,Q9NYV9,Q9NYW0,Q9NYW1,Q9NYW2,Q9NYW3,Q9NYW4,Q9NYW5,Q9NYW6,Q9NYW7,Q9NYX4,Q9NYZ1,Q9NYZ2,Q9NYZ4,Q9NZ01,Q9NZ08,Q9NZ20,Q9NZ42,Q9NZ43,Q9NZ45,Q9NZ53,Q9NZ56,Q9NZ94,Q9NZA1,Q9NZB2,Q9NZC2,Q9NZC3,Q9NZD1,Q9NZG7,Q9NZH0,Q9NZI2,Q9NZJ5,Q9NZJ7,Q9NZM1,Q9NZM6,Q9NZN1,Q9NZN3,Q9NZN4,Q9NZP0,Q9NZP2,Q9NZP5,Q9NZQ7,Q9NZQ8,Q9NZR2,Q9NZS2,Q9NZS9,Q9NZU0,Q9NZU1,Q9NZU7,Q9NZV1,Q9NZV8,Q9NZW4,Q9NZW5,Q9P003,Q9P035,Q9P055,Q9P0B6,Q9P0G3,Q9P0I2,Q9P0J0,Q9P0K1,Q9P0K9,Q9P0L0,Q9P0L2,Q9P0L9,Q9P0N5,Q9P0N8,Q9P0S2,Q9P0S3,Q9P0S9,Q9P0T7,Q9P0U1,Q9P0V3,Q9P0V8,Q9P0X4,Q9P109,Q9P121,Q9P126,Q9P1A6,Q9P1P4,Q9P1P5,Q9P1Q5,Q9P1W3,Q9P1W8,Q9P1Z3,Q9P1Z9,Q9P212,Q9P218,Q9P232,Q9P241,Q9P244,Q9P246,Q9P258,Q9P273,Q9P283,Q9P291,Q9P296,Q9P298,Q9P2B2,Q9P2C4,Q9P2D8,Q9P2E5,Q9P2E7,Q9P2E8,Q9P2E9,Q9P2J2,Q9P2K9,Q9P2N4,Q9P2P1,Q9P2S2,Q9P2U7,Q9P2U8,Q9P2V4,Q9P2W3,Q9P2W7,Q9P2W9,Q9P2X0,Q9TNN7,Q9TQE0,Q9UBC2,Q9UBD6,Q9UBF9,Q9UBG0,Q9UBH6,Q9UBI4,Q9UBI6,Q9UBJ2,Q9UBK5,Q9UBL6,Q9UBL9,Q9UBM1,Q9UBM4,Q9UBM7,Q9UBM8,Q9UBN1,Q9UBN4,Q9UBN6,Q9UBP0,Q9UBQ6,Q9UBR5,Q9UBS5,Q9UBS9,Q9UBT7,Q9UBU6,Q9UBV2,Q9UBV4,Q9UBV7,Q9UBW5,Q9UBX3,Q9UBX8,Q9UBY0,Q9UBY5,Q9UBY8,Q9UDX4,Q9UDX5,Q9UDY2,Q9UDY4,Q9UEF7,Q9UEU0,Q9UEW3,Q9UEY8,Q9UF02,Q9UF11,Q9UF33,Q9UF47,Q9UG22,Q9UGC6,Q9UGF5,Q9UGF6,Q9UGF7,Q9UGH3,Q9UGI6,Q9UGM1,Q9UGN4,Q9UGP8,Q9UGQ2,Q9UGQ3,Q9UGT4,Q9UH62,Q9UH65,Q9UH99,Q9UHB6,Q9UHC3,Q9UHC6,Q9UHC9,Q9UHE5,Q9UHE8,Q9UHF3,Q9UHF4,Q9UHI5,Q9UHI7,Q9UHI8,Q9UHJ9,Q9UHM6,Q9UHN6,Q9UHP7,Q9UHQ4,Q9UHQ9,Q9UHW9,Q9UHX3,Q9UI14,Q9UI33,Q9UI40,Q9UIA0,Q9UIB8,Q9UIG8,Q9UIJ5,Q9UIK5,Q9UIQ6,Q9UIR0,Q9UIW2,Q9UIX4,Q9UJ14,Q9UJ37,Q9UJ42,Q9UJ71,Q9UJ90,Q9UJ96,Q9UJ99,Q9UJA2,Q9UJA9,Q9UJG1,Q9UJM3,Q9UJQ1,Q9UJS0,Q9UJU6,Q9UJW2,Q9UJZ1,Q9UK00,Q9UK08,Q9UK17,Q9UK23,Q9UK28,Q9UK39,Q9UK41,Q9UKB5,Q9UKC9,Q9UKF2,Q9UKF5,Q9UKG4,Q9UKH3,Q9UKJ0,Q9UKJ1,Q9UKJ5,Q9UKJ8,Q9UKL2,Q9UKL4,Q9UKM7,Q9UKN1,Q9UKP4,Q9UKP5,Q9UKP6,Q9UKQ2,Q9UKR5,Q9UKR8,Q9UKS6,Q9UKU0,Q9UKU6,Q9UKV5,Q9UKX5,Q9UKY0,Q9UKY4,Q9UKZ4,Q9UL01,Q9UL19,Q9UL26,Q9UL51,Q9UL52,Q9UL54,Q9UL62,Q9ULB1,Q9ULB4,Q9ULB5,Q9ULC0,Q9ULC3,Q9ULC5,Q9ULC8,Q9ULD2,Q9ULD8,Q9ULF5,Q9ULG6,Q9ULH0,Q9ULH4,Q9ULI3,Q9ULK0,Q9ULK5,Q9ULK6,Q9ULL4,Q9ULM6,Q9ULP9,Q9ULQ1,Q9ULS5,Q9ULS6,Q9ULT0,Q9ULT6,Q9ULV1,Q9ULV4,Q9ULW2,Q9ULX3,Q9ULX5,Q9ULX7,Q9ULY5,Q9ULZ9,Q9UM00,Q9UM01,Q9UM11,Q9UM21,Q9UM44,Q9UM47,Q9UM54,Q9UM73,Q9UMD9,Q9UMF0,Q9UMR7,Q9UMX0,Q9UMX5,Q9UMX6,Q9UMX9,Q9UMY4,Q9UMZ3,Q9UN42,Q9UN66,Q9UN67,Q9UN70,Q9UN71,Q9UN72,Q9UN73,Q9UN74,Q9UN75,Q9UN76,Q9UN88,Q9UNA0,Q9UNA3,Q9UNE0,Q9UNF0,Q9UNG2,Q9UNK0,Q9UNL2,Q9UNN8,Q9UNP4,Q9UNQ0,Q9UNU6,Q9UNW8,Q9UNX9,Q9UP38,Q9UP52,Q9UP65,Q9UP79,Q9UP95,Q9UPC5,Q9UPI3,Q9UPN3,Q9UPQ8,Q9UPR5,Q9UPT5,Q9UPU3,Q9UPW8,Q9UPX0,Q9UPX6,Q9UPX8,Q9UPY5,Q9UPZ6,Q9UQ05,Q9UQ26,Q9UQ49,Q9UQ52,Q9UQ53,Q9UQ90,Q9UQC2,Q9UQC9,Q9UQD0,Q9UQF0,Q9UQP3,Q9UQQ1,Q9UQV4,Q9Y210,Q9Y217,Q9Y219,Q9Y225,Q9Y226,Q9Y227,Q9Y228,Q9Y231,Q9Y239,Q9Y241,Q9Y250,Q9Y256,Q9Y257,Q9Y267,Q9Y271,Q9Y272,Q9Y274,Q9Y275,Q9Y276,Q9Y277,Q9Y278,Q9Y279,Q9Y282,Q9Y284,Q9Y286,Q9Y287,Q9Y289,Q9Y296,Q9Y2A7,Q9Y2A9,Q9Y2B1,Q9Y2B2,Q9Y2C2,Q9Y2C3,Q9Y2C5,Q9Y2C9,Q9Y2D2,Q9Y2E8,Q9Y2G0,Q9Y2G1,Q9Y2G3,Q9Y2G8,Q9Y2H0,Q9Y2H6,Q9Y2H9,Q9Y2I1,Q9Y2I2,Q9Y2I9,Q9Y2J0,Q9Y2J2,Q9Y2K9,Q9Y2P4,Q9Y2P5,Q9Y2Q0,Q9Y2R0,Q9Y2T5,Q9Y2T6,Q9Y2U2,Q9Y2U8,Q9Y2W3,Q9Y2W7,Q9Y2Y6,Q9Y320,Q9Y328,Q9Y336,Q9Y342,Q9Y345,Q9Y385,Q9Y397,Q9Y3A6,Q9Y3B3,Q9Y3D6,Q9Y3E0,Q9Y3N9,Q9Y3P4,Q9Y3P8,Q9Y3Q0,Q9Y3Q3,Q9Y3Q4,Q9Y3Q7,Q9Y3R0,Q9Y3S1,Q9Y442,Q9Y487,Q9Y490,Q9Y493,Q9Y4A9,Q9Y4B5,Q9Y4C0,Q9Y4C2,Q9Y4C5,Q9Y4D2,Q9Y4D7,Q9Y4D8,Q9Y4E1,Q9Y4F1,Q9Y4F9,Q9Y4G6,Q9Y4G8,Q9Y4J8,Q9Y4K0,Q9Y4W6,Q9Y4X1,Q9Y519,Q9Y548,Q9Y561,Q9Y584,Q9Y585,Q9Y597,Q9Y5E1,Q9Y5E2,Q9Y5E3,Q9Y5E4,Q9Y5E5,Q9Y5E6,Q9Y5E7,Q9Y5E8,Q9Y5E9,Q9Y5F0,Q9Y5F1,Q9Y5F2,Q9Y5F3,Q9Y5F6,Q9Y5F7,Q9Y5F8,Q9Y5F9,Q9Y5G0,Q9Y5G1,Q9Y5G2,Q9Y5G3,Q9Y5G4,Q9Y5G5,Q9Y5G6,Q9Y5G7,Q9Y5G8,Q9Y5G9,Q9Y5H0,Q9Y5H1,Q9Y5H2,Q9Y5H3,Q9Y5H4,Q9Y5H5,Q9Y5H6,Q9Y5H7,Q9Y5H8,Q9Y5H9,Q9Y5I0,Q9Y5I1,Q9Y5I2,Q9Y5I3,Q9Y5I4,Q9Y5I7,Q9Y5K8,Q9Y5L2,Q9Y5L3,Q9Y5M8,Q9Y5N1,Q9Y5P0,Q9Y5P1,Q9Y5Q0,Q9Y5Q5,Q9Y5R2,Q9Y5S1,Q9Y5S2,Q9Y5S8,Q9Y5T4,Q9Y5U4,Q9Y5U5,Q9Y5U8,Q9Y5U9,Q9Y5V3,Q9Y5W7,Q9Y5W9,Q9Y5X1,Q9Y5X3,Q9Y5X5,Q9Y5Y0,Q9Y5Y3,Q9Y5Y4,Q9Y5Y5,Q9Y5Y6,Q9Y5Y7,Q9Y5Y9,Q9Y5Z0,Q9Y5Z6,Q9Y5Z9,Q9Y619,Q9Y624,Q9Y625,Q9Y639,Q9Y644,Q9Y653,Q9Y661,Q9Y662,Q9Y663,Q9Y666,Q9Y672,Q9Y673,Q9Y679,Q9Y691,Q9Y693,Q9Y694,Q9Y696,Q9Y698,Q9Y6A1,Q9Y6A2,Q9Y6A9,Q9Y6C2,Q9Y6C5,Q9Y6C9,Q9Y6D0,Q9Y6F6,Q9Y6F9,Q9Y6G1,Q9Y6H6,Q9Y6H8,Q9Y6I0,Q9Y6I3,Q9Y6I8,Q9Y6J6,Q9Y6K0,Q9Y6L6,Q9Y6M0,Q9Y6M5,Q9Y6M7,Q9Y6N1,Q9Y6N6,Q9Y6N7,Q9Y6N8,Q9Y6Q6,Q9Y6Q9,Q9Y6R1,Q9Y6T7,Q9Y6U7,Q9Y6W5,Q9Y6W8,Q9Y6X1,Q9Y6X5,Q9Y6Y9,Q9YNA8';

/* ====== APP MODULES ====== */

/* ====== _v2_state.js ====== */
/* ==============================================================
   STATE  —  Surfaceome Explorer v2
   Central data store. All modules read/write only through App.
   ============================================================== */

const App = {
  file:      null,   // { name, size, ext }
  raw:       null,   // { headers[], rows[], fileType, delimiter, nCols, nRows }
  config:    null,   // { idCol, geneCol, quantCols[] }
  check:     null,   // populated by Tab 2
  processed: null,   // reserved
  filtered:  null,   // populated by Tab 3
  imputed:    null,   // populated by Tab 4
  normalized: null,   // populated by Tab 5
  activeTab: 1,
  tabUnlocked: { 1:true, 2:false, 3:false, 4:false, 5:false, 6:false },
};

const TAB_META = {
  1: { label:'Upload Data',       short:'Upload'     },   2: { label:'Data Check',        short:'Data Check' },   3: { label:'Data Processing',   short:'Processing' },   4: { label:'Missing Value Estimation', short:'Imputation' },   5: { label:'Normalization',      short:'Normalize'  },   6: { label:'Surfaceome Filter',  short:'Filter'     },
};

function unlockTab(n) {
  App.tabUnlocked[n] = true;
  renderTabBar();
}

function goTab(n) {
  if (!App.tabUnlocked[n]) { showToast('Complete the previous step first.'); return; }
  App.activeTab = n;
  renderTabBar();
  renderTabContent(n);
  /* Only scroll workflow into view if it is above the current viewport —
     do NOT force-scroll when user is already inside the workflow area. */
  const wf = document.getElementById('workflow');
  if (wf) {
    const top = wf.getBoundingClientRect().top;
    if (top < -20 || top > window.innerHeight * 0.6) {
      wf.scrollIntoView({ behavior: 'smooth', block: 'start' });
    }
  }
}

/* ====== _v2_parser.js ====== */
/* ==============================================================
   PARSER  —  CSV / TSV / XLSX  →  { headers, rows, meta }
   ============================================================== */

const UP_RE = /^[A-Z][0-9][A-Z0-9]{3}[0-9]$|^[A-Z]{3}[0-9][A-Z0-9]{3}[0-9]$|^A0A[A-Z0-9]{6,10}$/;

/* Returns a Promise<parsedData> */
function parseFile(file) {
  const ext = file.name.split('.').pop().toLowerCase();
  return new Promise((resolve, reject) => {
    if (ext === 'xlsx' || ext === 'xls') {
      const rd = new FileReader();
      rd.onload = e => {
        try {
          const wb = XLSX.read(e.target.result, { type:'array' });
          const ws = wb.Sheets[wb.SheetNames[0]];
          const data = XLSX.utils.sheet_to_json(ws, { header:1, defval:'' });           if (data.length < 2) throw new Error('Sheet has fewer than 2 rows');
          const headers = data[0].map(h => String(h).trim());
          const rows = data.slice(1)
            .filter(r => r.some(v => String(v).trim() !== ''))
            .map(r => {
              const obj = {};
              headers.forEach((h, i) => { obj[h] = r[i] !== undefined ? r[i] : ''; });
              return obj;
            });
          resolve({ headers, rows, fileType:'xlsx', delimiter:null, nCols:headers.length, nRows:rows.length });
        } catch(err) { reject(err); }
      };
      rd.onerror = () => reject(new Error('Could not read file'));
      rd.readAsArrayBuffer(file);
    } else {
      const rd = new FileReader();
      rd.onload = e => {
        try { resolve(parseTextFile(e.target.result, file.name)); }
        catch(err) { reject(err); }
      };
      rd.onerror = () => reject(new Error('Could not read file'));
      rd.readAsText(file);
    }
  });
}

function parseTextFile(text, filename) {
  const ext = filename.split('.').pop().toLowerCase();   const lines = text.trim().replace(/\r\n/g,'\n').replace(/\r/g,'\n').split('\n').filter(l => l.trim());   if (lines.length < 2) throw new Error('File has fewer than 2 rows');   const delim = ext === 'tsv' ? '\t' : detectDelimiter(lines[0]);   const fileType = delim === '\t' ? 'tsv' : 'csv';

  function splitLine(line) {
    const res = []; let cur = '', q = false;
    for (let i = 0; i < line.length; i++) {
      const c = line[i];
      if (c === '"') { q = !q; continue; }       if (c === delim && !q) { res.push(cur.trim()); cur = ''; continue; }
      cur += c;
    }
    res.push(cur.trim());
    return res;
  }

  const headers = splitLine(lines[0]);
  const rows = [];
  for (let i = 1; i < lines.length; i++) {
    const vals = splitLine(lines[i]);
    if (vals.every(v => !v)) continue;
    const row = {};
    headers.forEach((h, j) => { row[h] = vals[j] !== undefined ? vals[j] : ''; });
    rows.push(row);
  }
  return { headers, rows, fileType, delimiter:delim, nCols:headers.length, nRows:rows.length };
}

function detectDelimiter(line) {
  const tabs   = (line.match(/\t/g) || []).length;
  const commas = (line.match(/,/g)  || []).length;
  const semis  = (line.match(/;/g)  || []).length;
  if (tabs > commas && tabs > semis) return '\t';   if (semis > commas) return ';';   return ',';
}

/* Auto-detect Protein ID, Gene, and quantitative columns */
const MISSING_RE = /^(na|nan|n\/a|#n\/a|null|none|missing|undefined|inf|-inf)$/i;

function autoDetectCols(headers, rows) {
  const sample = rows.slice(0, 50);

  const scored = headers.map(h => {
    const hLow = h.toLowerCase().replace(/[\s_\-]/g, '');
    let numericN = 0, total = 0;
    sample.forEach(r => {
      const v = String(r[h] !== undefined ? r[h] : '').trim();
      if (!v || MISSING_RE.test(v)) return;   // skip empty / NA / NaN
      total++;
      if (!isNaN(parseFloat(v))) numericN++;
    });
    const numFrac = total > 0 ? numericN / total : 0;
    const isGeneLike = /^(gene|genesymbol|genename|symbol|geneid)$/.test(hLow);
    // Metadata columns: text-heavy descriptors unlikely to be abundance values
    const isMetaLike = /description|function|family|pathway|ontology|mw|mass|length|fdr|qval|psm|coverage|sequence/.test(hLow);
    return { h, numFrac, isGeneLike, isMetaLike };
  });

  // Column 1 is always Protein ID
  const idCol = headers[0];
  const geneCol = scored.find(s => s.isGeneLike && s.h !== idCol)?.h || null;

  // Primary: ≥50% of non-missing values are numeric, and not a metadata column
  let quantCols = scored
    .filter(s => s.numFrac >= 0.5 && s.h !== idCol && s.h !== geneCol && !s.isMetaLike)
    .map(s => s.h);

  // Fallback 1: any column with ≥20% numeric (handles highly sparse data)
  if (quantCols.length === 0) {
    quantCols = scored
      .filter(s => s.numFrac >= 0.2 && s.h !== idCol && s.h !== geneCol)
      .map(s => s.h);
  }

  // Fallback 2: include all non-ID columns so the user can always proceed
  // (Data Check tab will flag quality issues)
  if (quantCols.length === 0) {
    quantCols = headers.filter(h => h !== idCol && h !== geneCol);
  }

  return { idCol, geneCol, quantCols };
}

/* Check what fraction of sampled IDs look like UniProt accessions */
function uniprotMatchPct(rows, idCol, sampleN = 50) {
  const ids = rows.slice(0, sampleN).map(r => String(r[idCol] || '').trim().toUpperCase());
  const valid = ids.filter(Boolean);
  if (!valid.length) return 0;
  return Math.round(valid.filter(id => UP_RE.test(id)).length / valid.length * 100);
}

/* Safe string → number (returns null for missing/NA/NaN) */
function toNum(v) {
  if (v === undefined || v === null) return null;
  const s = String(v).trim();
  if (!s || s.toUpperCase() === 'NA' || s.toUpperCase() === 'NAN') return null;
  const n = parseFloat(s);
  return isNaN(n) ? null : n;
}

/* ====== _v2_tab1.js ====== */
/* ==============================================================
   TAB 1  —  Upload Data
   Protein ID is always Column 1. No column picker UI.
   ============================================================== */

/* ── Random proteomics dataset generator ──────────────────────
   Generates a new realistic mass spec dataset each time:
   • 3,000–7,000 proteins (random per call)
   • Surface proteins sampled from the real PM annotation DB
     so they will annotate correctly in Tab 6
   • Remaining proteins get random UniProt-format IDs
   • log₂ LFQ intensities, ~20% missing, ~8% differential
   • 2 groups (Control, Treatment), 3 replicates each
   ─────────────────────────────────────────────────────────── */
function _genSampleRaw() {
  const COLS = ['UniProt_ID',     'Control_Rep1','Control_Rep2','Control_Rep3',     'Treatment_Rep1','Treatment_Rep2','Treatment_Rep3'];

  /* Box-Muller normal sample */
  function rn() {
    const u = 1 - Math.random(), v = Math.random();
    return Math.sqrt(-2 * Math.log(u)) * Math.cos(2 * Math.PI * v);
  }

  /* ── Total protein count: 3,000–7,000 ── */
  const N = Math.floor(Math.random() * 4001) + 3000;

  /* ── Build set of all known annotation IDs (avoid accidental hits) ── */
  const knownIds = new Set();
  if (typeof _PM_RAW !== 'undefined') _PM_RAW.split(',').forEach(id => knownIds.add(id.trim()));   if (typeof _SS_RAW !== 'undefined') _SS_RAW.split(',').forEach(id => knownIds.add(id.trim()));   if (typeof _MA_RAW !== 'undefined') _MA_RAW.split(',').forEach(id => knownIds.add(id.trim()));

  /* ── Sample surface proteins from the real PM database (15–25%) ── */
  const pmList = typeof _PM_RAW !== 'undefined'     ? _PM_RAW.split(',').map(s => s.trim()).filter(Boolean)
    : [];
  const surfFrac = 0.15 + Math.random() * 0.10;
  const nSurf    = Math.min(Math.round(N * surfFrac), pmList.length);
  /* Shuffle in place (Fisher-Yates on a copy) */
  const pmCopy = pmList.slice();
  for (let i = pmCopy.length - 1; i > 0; i--) {
    const j = Math.floor(Math.random() * (i + 1));
    [pmCopy[i], pmCopy[j]] = [pmCopy[j], pmCopy[i]];
  }
  const surfIds = pmCopy.slice(0, nSurf);

  /* ── Generate non-surface protein IDs (no DB collisions) ── */
  const usedIds    = new Set(surfIds);
  const nonSurfIds = [];
  const PFXS = ['P','Q','O','A','B'];
  while (nonSurfIds.length < N - nSurf) {
    const id = PFXS[Math.floor(Math.random() * 5)] +
               String(Math.floor(Math.random() * 99998) + 1).padStart(5, '0');
    if (!usedIds.has(id) && !knownIds.has(id)) {
      usedIds.add(id);
      nonSurfIds.push(id);
    }
  }

  /* ── Shuffle all IDs together ── */
  const allIds = surfIds.concat(nonSurfIds);
  for (let i = allIds.length - 1; i > 0; i--) {
    const j = Math.floor(Math.random() * (i + 1));
    [allIds[i], allIds[j]] = [allIds[j], allIds[i]];
  }

  /* ── Assign regulation class per protein ── */
  // ~4% up, ~4% down, ~1% ctrl-only, ~1% treat-only, rest unchanged
  const regOf = {};
  allIds.forEach(id => {
    const r = Math.random();
    regOf[id] = r < 0.04 ? 'up'               : r < 0.08 ? 'down'               : r < 0.09 ? 'ctrlOnly'               : r < 0.10 ? 'treatOnly'               : 'none';
  });

  /* ── Build rows ── */
  const rows = allIds.map(id => {
    const type      = regOf[id];
    const baseMean  = 20 + Math.random() * 10;   // 20–30 log₂ LFQ
    const withinSD  = 0.12 + Math.random() * 0.22; // replicate noise
    const missPr    = 0.08 + Math.random() * 0.22; // 8–30% missing per protein

    let ctrlBase  = baseMean;
    let treatBase = type === 'up'   ? baseMean + 1.5 + Math.random() * 2.0                   : type === 'down' ? baseMean - 1.5 - Math.random() * 2.0
                  :                   baseMean + (Math.random() - 0.5) * 0.35;

    function rep(base, isCtrl) {
      if (type === 'treatOnly' && isCtrl)  return null;       if (type === 'ctrlOnly'  && !isCtrl) return null;
      if (Math.random() < missPr)          return null;
      return Math.max(14, base + rn() * withinSD);
    }

    return {
      'UniProt_ID':    id,       'Control_Rep1':   rep(ctrlBase,  true),       'Control_Rep2':   rep(ctrlBase,  true),       'Control_Rep3':   rep(ctrlBase,  true),       'Treatment_Rep1': rep(treatBase, false),       'Treatment_Rep2': rep(treatBase, false),       'Treatment_Rep3': rep(treatBase, false),
    };
  });

  return { headers: COLS, rows, nRows: rows.length, nCols: COLS.length, fileType: 'csv' };
}

function loadSampleData() {
  try {
    showToast('Generating sample dataset…');
    setTimeout(() => {
      try {
        const raw = _genSampleRaw();
        App.file   = { name: 'sample_data.csv', size: raw.nRows * 55, ext: 'csv' };
        App.raw    = raw;
        App.config = autoDetectCols(raw.headers, raw.rows);
        renderLoadedState(raw, App.config);
        showToast('Sample dataset loaded — ' + raw.nRows.toLocaleString() +                   ' proteins · 2 groups × 3 replicates');
      } catch(err) {
        showToast('Error generating sample data: ' + err.message);
      }
    }, 10);
  } catch(err) {
    showToast('Error: ' + err.message);
  }
}

function renderTab1() {
  const mc = document.getElementById('main-content');
  mc.innerHTML =
    '<div class="tab-content">' +       '<div class="tab-hd">' +         '<h2 class="tab-title">Upload Data</h2>' +         '<p class="tab-desc">Upload your peptide/protein abundance table in CSV, TSV, or Excel format. ' +         'The <strong>first column</strong> must contain <strong>UniProt Protein IDs</strong> ' +         '(e.g.&nbsp;P00533, Q9Y6K9). All remaining numeric columns will be treated as ' +         'quantitative sample measurements.</p>' +       '</div>' +

      /* ── Drop zone card ── */
      '<div class="card" id="card-upload">' +         '<div class="card-title">Select file</div>' +         '<div class="drop-zone" id="drop-zone">' +           '<input type="file" id="file-input" accept=".csv,.tsv,.txt,.xlsx,.xls" ' +                  'style="position:absolute;inset:0;opacity:0;cursor:pointer;width:100%;height:100%">' +           '<div class="drop-icon">&#x2B06;</div>' +           '<div class="drop-primary">Drop your file here or click to browse</div>' +           '<div class="drop-secondary">CSV &middot; TSV &middot; Excel (.xlsx)</div>' +         '</div>' +         '<div id="file-status"></div>' +         '<div class="t1-sample-row">' +           '<span class="t1-sample-label">No data yet?</span>' +           '<button class="t1-sample-btn" onclick="loadSampleData()">' +             '&#x25B6; Load sample dataset' +           '</button>' +           '<span class="t1-sample-hint">3,000&ndash;7,000 proteins &middot; 2 conditions &middot; 3 replicates &middot; log&#x2082; LFQ &middot; randomly generated each time</span>' +         '</div>' +       '</div>' +

      /* ── Preview card (hidden until loaded) ── */
      '<div class="card" id="card-preview" style="display:none">' +         '<div class="card-title">Data preview <span id="preview-meta" class="card-meta"></span></div>' +         '<div class="preview-wrap"><table class="ptable" id="preview-table"></table></div>' +         '<div class="preview-legend">' +           '<span class="leg-item"><span class="leg-dot leg-id"></span>Protein ID (Column 1)</span>' +           '<span class="leg-item"><span class="leg-dot leg-quant"></span>Quantitative</span>' +           '<span class="leg-item"><span class="leg-dot leg-other"></span>Other / skipped</span>' +         '</div>' +       '</div>' +

      /* ── Proceed card (hidden until loaded) ── */
      '<div class="card" id="card-proceed" style="display:none">' +         '<div id="col-summary"></div>' +         '<div class="notice notice-warn" id="col-warn" style="display:none"></div>' +         '<div class="btn-row" style="margin-top:18px">' +           '<button class="btn btn-primary btn-lg" onclick="proceedTab1()">Proceed to Data Check &#8594;</button>' +           '<button class="btn btn-ghost" onclick="resetUpload()">&#8634; Upload a different file</button>' +         '</div>' +       '</div>' +     '</div>';

  /* Set up drag-and-drop */
  const dz  = document.getElementById('drop-zone');   const inp = document.getElementById('file-input');   dz.addEventListener('dragover',  e => { e.preventDefault(); dz.classList.add('drag-over'); });   dz.addEventListener('dragleave', () => dz.classList.remove('drag-over'));   dz.addEventListener('drop', e => {     e.preventDefault(); dz.classList.remove('drag-over');
    const f = e.dataTransfer.files[0]; if (f) handleFileLoad(f);
  });
  inp.addEventListener('change', e => { if (e.target.files[0]) handleFileLoad(e.target.files[0]); });

  /* Restore state if navigating back */
  if (App.raw) renderLoadedState(App.raw, App.config);
}

/* ── File load pipeline ── */
function handleFileLoad(file) {
  const ext = file.name.split('.').pop().toLowerCase();   if (!['csv','tsv','txt','xlsx','xls'].includes(ext)) {     showToast('Unsupported file. Please upload CSV, TSV, or Excel (.xlsx).');
    return;
  }
  App.file = { name: file.name, size: file.size, ext };
  document.getElementById('file-status').innerHTML =     '<div class="notice notice-info" style="margin-top:12px">&#x23F3; Loading ' + esc(file.name) + '&hellip;</div>';

  parseFile(file).then(raw => {
    App.raw    = raw;
    App.config = autoDetectCols(raw.headers, raw.rows);
    renderLoadedState(raw, App.config);
    showToast('Loaded ' + raw.nRows.toLocaleString() + ' proteins, ' +               App.config.quantCols.length + ' sample columns detected');
  }).catch(err => {
    document.getElementById('file-status').innerHTML =       '<div class="notice notice-err" style="margin-top:12px">&#x2717; Error: ' + esc(err.message) + '</div>';
  });
}

function resetUpload() {
  App.file = null; App.raw = null; App.config = null; App.check = null;
  App.tabUnlocked = { 1:true, 2:false, 3:false, 4:false, 5:false, 6:false };
  /* Clear Tab 3 module globals if they exist */
  if (typeof _t3Metrics !== 'undefined') { _t3Metrics = null; }   if (typeof _t3FS     !== 'undefined') { _t3FS = null; }
  /* Clear Tab 4 module globals if they exist */
  if (typeof _t4MV     !== 'undefined') { _t4MV = null; }   if (typeof _t4Method !== 'undefined') { _t4Method = 'lod'; }
  App.imputed = null;
  /* Clear Tab 5 module globals if they exist */
  if (typeof _t5Cache !== 'undefined') { _t5Cache = null; }   if (typeof _t5State !== 'undefined') { _t5State = null; }
  App.normalized = null;
  renderTabBar();
  renderTab1();
}

/* ── Render state after a successful file parse ── */
function renderLoadedState(raw, config) {
  const pct = uniprotMatchPct(raw.rows, config.idCol);
  const fmtStr = raw.fileType === 'xlsx' ? 'Excel (.xlsx)' :                  raw.fileType === 'tsv'  ? 'Tab-separated (.tsv)' : 'Comma-separated (.csv)';   const badgeClass = pct >= 70 ? 'fbadge-ok' : pct >= 30 ? 'fbadge-warn' : 'fbadge-err';   const badgeText  = pct + '% UniProt ID match';

  document.getElementById('file-status').innerHTML =     '<div class="file-badge">' +       '<span class="fbadge-icon">&#x2713;</span>' +       '<span class="fbadge-name">' + esc(App.file.name) + '</span>' +       '<span class="fbadge-meta">' + raw.nRows.toLocaleString() + ' proteins &middot; ' +         raw.nCols + ' columns &middot; ' + fmtStr + '</span>' +       '<span class="fbadge-upid ' + badgeClass + '">' + badgeText + '</span>' +     '</div>';

  /* Preview */
  document.getElementById('card-preview').style.display = '';   document.getElementById('preview-meta').textContent =     '(first ' + Math.min(raw.nRows, 20) + ' of ' + raw.nRows.toLocaleString() + ' rows)';
  renderPreviewTable(raw, config);

  /* Column summary + proceed */
  document.getElementById('card-proceed').style.display = '';
  const qn = config.quantCols.length;
  const summaryClass = qn === 0 ? 'notice notice-warn' : 'notice notice-ok';   const summaryIcon  = qn === 0 ? '&#x26A0;' : '&#x2713;';
  const summaryText  = qn === 0
    ? 'No numeric sample columns detected. Ensure abundance values are numbers, not text.'     : 'Protein ID: <strong>' + esc(config.idCol) + '</strong> &nbsp;&middot;&nbsp; ' +       '<strong>' + qn + '</strong> quantitative sample column' + (qn === 1 ? '' : 's') + ' detected automatically.';   document.getElementById('col-summary').innerHTML =     '<div class="' + summaryClass + '">' + summaryIcon + ' ' + summaryText + '</div>';
}

/* ── Preview table ── */
function renderPreviewTable(raw, cfg) {
  const MAX_COLS = 10;
  const show  = raw.headers.slice(0, MAX_COLS);
  const extra = raw.headers.length - MAX_COLS;

  let html = '<thead><tr>';
  show.forEach(h => {
    let cls = h === cfg.idCol ? 'th-id'             : cfg.quantCols.includes(h) ? 'th-quant'             : 'th-other';     html += '<th class="' + cls + '">' + esc(h) + '</th>';
  });
  if (extra > 0) html += '<th class="th-more">+' + extra + ' cols</th>';   html += '</tr></thead><tbody>';

  raw.rows.slice(0, 20).forEach(row => {
    html += '<tr>';
    show.forEach(h => {
      html += '<td>' + esc(String(row[h] !== undefined ? row[h] : '')) + '</td>';
    });
    if (extra > 0) html += '<td class="th-more">&hellip;</td>';     html += '</tr>';
  });
  html += '</tbody>';   document.getElementById('preview-table').innerHTML = html;
}

/* ── Proceed to Tab 2 ── */
function proceedTab1() {
  const cfg  = App.config;
  const warn = document.getElementById('col-warn');
  if (!cfg) {
    warn.textContent = 'Please upload a file first.';     warn.style.display = '';
    return;
  }
  warn.style.display = 'none';
  unlockTab(2);
  goTab(2);
}

/* ====== _v2_stubs.js ====== */
/* ==============================================================
   TAB STUBS  —  Tabs 2–6 (placeholders until implemented)
   Each stub shows a preview of what the tab will contain.
   ============================================================== */

const TAB_PLANS = {
  2: {
    title: 'Data Check',     desc:  'Automatic quality checks run against your uploaded dataset.',
    items: [
      'Protein and sample counts',       'Missing value summary (count and percentage)',       'Duplicate protein ID detection',       'UniProt accession match rate',       'Non-numeric value warnings',       'Column type validation',
    ],
    note: 'This step is informational only — no data is modified.',
  },
  3: {
    title: 'Data Processing',     desc:  'Optional preprocessing before surfaceome annotation.',
    items: [
      'Missing value handling (remove, zero, half-minimum)',       'Log₂ transformation',       'Median normalization',       'Z-score scaling',       'Before / after boxplot comparison',
    ],
    note: 'All preprocessing is optional. You can proceed with raw data.',
  },
  4: {
    title: 'Surfaceome Filter',     desc:  'Core feature — annotate proteins against multiple surfaceome databases.',
    items: [
      'SURFY in silico predicted surfaceome (2,886 proteins)',       'UniProt plasma membrane (KW-1003, 6,944 proteins)',       'Membrane-associated proteins (5,973 proteins)',       'High-confidence intersections across databases',       'Custom filter combinations',       'Search by gene symbol, protein name, or UniProt ID',
    ],
    note: null,
  },
  5: {
    title: 'Visualization',     desc:  'Interactive charts for exploring your surfaceome results.',
    items: [
      'Category summary bar chart',       'Surfaceome composition pie chart',       'Top 20 most abundant surface proteins',       'Sample heatmap (multi-replicate datasets)',
    ],
    note: 'All charts can be downloaded as PNG.',
  },
  6: {
    title: 'Download Results',     desc:  'Export your annotated surfaceome data.',
    items: [
      'All detected proteins with annotations',       'SURFY-only proteins',       'UniProt membrane proteins',       'Membrane-associated proteins',       'High-confidence surfaceome intersection',
    ],
    note: 'Export as CSV or Excel (.xlsx).',
  },
};

function renderStub(n) {
  const plan = TAB_PLANS[n];
  const prevLabel = TAB_META[n-1] ? TAB_META[n-1].label : '';
  const isLocked = !App.tabUnlocked[n];
  const mc = document.getElementById('main-content');

  const itemsHTML = plan.items.map(item =>
    '<li class="stub-item"><span class="stub-bullet">&#x25B8;</span>' + esc(item) + '</li>'   ).join('');

  const noteHTML = plan.note
    ? '<div class="stub-note"><span>&#x2139;</span> ' + esc(plan.note) + '</div>'     : '';

  // Only show the lock banner when the tab is genuinely locked
  const lockedBanner = isLocked
    ? '<div class="stub-locked-banner">' +         '<div class="stub-lock-icon">&#x1F512;</div>' +         '<div>' +           '<div class="stub-lock-title">Complete <em>' + esc(prevLabel) + '</em> to unlock this step</div>' +           '<div class="stub-lock-sub">Work through each tab in order to reach this step.</div>' +         '</div>' +       '</div>'     : '<div class="notice notice-info" style="margin-bottom:20px">' +         '&#x1F6A7; This step is coming soon. For now you can see what it will include below.' +       '</div>';

  mc.innerHTML =
    '<div class="tab-content">' +       '<div class="tab-hd">' +         '<h2 class="tab-title">' + esc(plan.title) + '</h2>' +         '<p class="tab-desc">' + esc(plan.desc) + '</p>' +       '</div>' +       '<div class="card stub-card">' +
        lockedBanner +
        '<div class="stub-plan">' +           '<div class="stub-plan-title">This tab will include:</div>' +           '<ul class="stub-list">' + itemsHTML + '</ul>' +         '</div>' +
        noteHTML +
      '</div>' +     '</div>';
}

function renderTab2() { renderStub(2); }
function renderTab3() { renderStub(3); }
function renderTab4() { renderStub(4); }
function renderTab5() { renderStub(5); }
function renderTab6() { renderStub(6); }

/* ====== _v2_tab2.js ====== */
/* ==============================================================
   TAB 2  —  Data Check
   Overrides the stub renderTab2() defined in _v2_stubs.js.
   ============================================================== */

/* ── Sanitize a single sample column name ── */
function sanitizeSampleName(name) {
  return String(name)
    .replace(/\s+/g, '_')     .replace(/[^A-Za-z0-9_\-\/]/g, '');
}

/* ── Apply rename map to App.raw headers and row keys in-place ── */
function applySanitizedNames(renameMap) {
  App.raw.headers = App.raw.headers.map(h =>
    renameMap[h] !== undefined ? renameMap[h] : h
  );
  App.raw.rows = App.raw.rows.map(row => {
    const nr = {};
    Object.keys(row).forEach(k => {
      nr[renameMap[k] !== undefined ? renameMap[k] : k] = row[k];
    });
    return nr;
  });
  /* geneCol might be in quantCols domain — update if present */
  if (App.config.geneCol && renameMap[App.config.geneCol]) {
    App.config.geneCol = renameMap[App.config.geneCol];
  }
}

/* ── Detect group structure from column names ── */
/*  Strips trailing replicate indicators: _1, -1, .1, _rep1, rep1, 1 */
function detectGroups(colNames) {
  const prefixes = colNames.map(n =>
    n.replace(/[\._\-]?rep\d+$/i, '')      .replace(/[\._\-\s]\d+$/, '')      .replace(/\d+$/, '')      .replace(/[\._\-\s]+$/, '')
     .trim() || n
  );
  const map = {};
  prefixes.forEach((p, i) => {
    if (!map[p]) map[p] = [];
    map[p].push(colNames[i]);
  });
  return map;
}

/* ── Count missing values across the quantitative matrix ── */
function countMissing(rows, qCols) {
  let n = 0;
  rows.forEach(r => qCols.forEach(c => {
    const v = String(r[c] !== undefined ? r[c] : '').trim();
    if (!v || MISSING_RE.test(v)) n++;
  }));
  return n;
}

/* ── Count cells that are non-empty, non-missing, and non-numeric ── */
function countNonNumeric(rows, qCols) {
  let n = 0;
  rows.forEach(r => qCols.forEach(c => {
    const v = String(r[c] !== undefined ? r[c] : '').trim();
    if (!v || MISSING_RE.test(v)) return;   /* skip recognized missing */
    if (isNaN(parseFloat(v))) n++;
  }));
  return n;
}

/* ── Find rows where all non-missing values are identical ── */
function findConstantFeatures(rows, qCols) {
  return rows.reduce((acc, row, i) => {
    const vals = qCols
      .map(c => {
        const v = String(row[c] !== undefined ? row[c] : '').trim();
        return (!v || MISSING_RE.test(v)) ? null : parseFloat(v);
      })
      .filter(v => v !== null && !isNaN(v));
    if (vals.length > 0 && new Set(vals).size <= 1) acc.push(i);
    return acc;
  }, []);
}

/* ── Human-readable file type label ── */
function fileTypeLabel(ft) {
  return ({
    csv:  'comma separated values (.csv)',     tsv:  'tab separated values (.tsv)',     xlsx: 'Excel (.xlsx)',
  })[ft] || ft;
}

/* ── Status badge HTML ── */
function chkBadge(status) {
  const icons = { pass: '&#x2714;', warn: '&#x26A0;', err: '&#x2718;' };   const labels = { pass: 'Passed', warn: 'Warning', err: 'Error' };   return '<span class="chk-badge chk-badge-' + status + '">' +     icons[status] + ' ' + labels[status] +   '</span>';
}

/* ──────────────────────────────────────────────────────────────
   MAIN CHECK RUNNER
   ────────────────────────────────────────────────────────────── */
function runDataCheck() {
  const raw   = App.raw;
  const cfg   = App.config;
  let   qCols = cfg.quantCols.slice();   /* working copy; may be updated below */

  const report = {
    nProteins:              raw.nRows,
    nSamples:               qCols.length,
    fileType:               raw.fileType,
    checks:                 [],
    colRename:              {},
    nModified:              0,
    nNonNumeric:            0,
    nMissing:               0,
    missPct:                '0.0',
    totalDataPoints:        raw.nRows * qCols.length,
    nGroups:                null,
    groups:                 null,
    constantFeatureIndices: [],
    overallStatus:          'pass',
  };

  /* ── 1. Data content & orientation ── */
  report.checks.push({
    id: 'content', label: 'Data content & orientation', status: 'pass',
    message:
      'Samples are in columns and features are in rows. ' +       'The uploaded file is in ' + fileTypeLabel(raw.fileType) + ' format. ' +       'The dataset contains ' + raw.nRows.toLocaleString() + ' protein' +       (raw.nRows !== 1 ? 's' : '') + ' and ' + qCols.length + ' sample' +       (qCols.length !== 1 ? 's' : '') + '.',
  });

  /* ── 2. Sample name sanitization ── */
  const renameMap = {};
  let nMod = 0;
  qCols.forEach(col => {
    const s = sanitizeSampleName(col);
    renameMap[col] = s;
    if (s !== col) nMod++;
  });
  report.colRename = renameMap;
  report.nModified = nMod;

  report.checks.push({
    id: 'sampleNames', label: 'Sample names', status: nMod > 0 ? 'warn' : 'pass',
    nModified: nMod,
    message: nMod > 0
      ? nMod + ' sample name' + (nMod > 1 ? 's were' : ' was') + ' modified. ' +         'Only English letters, numbers, underscore (_), hyphen (-), and forward slash (/) are allowed. ' +         'Other special characters or punctuation were stripped.'       : 'All sample names contain only valid characters (letters, numbers, _, -, /).',
  });

  /* Apply sanitized names so downstream checks see clean column keys */
  if (nMod > 0) {
    applySanitizedNames(renameMap);
    cfg.quantCols = qCols.map(c => renameMap[c] !== undefined ? renameMap[c] : c);
    qCols = cfg.quantCols.slice();
  }

  /* ── 3. Group / class label detection ── */
  const gMap   = detectGroups(qCols);
  const gNames = Object.keys(gMap);
  const allUnique = gNames.length === qCols.length;  /* every sample has a unique prefix */
  const oneGroup  = gNames.length === 1;
  let gStatus = 'pass', gMsg;

  if (allUnique || oneGroup) {
    gMsg = 'No distinct group structure detected in sample names. Group-based checks skipped.';
  } else {
    report.nGroups = gNames.length;
    report.groups  = gMap;
    const underRep = gNames.filter(g => gMap[g].length < 3);
    gMsg = gNames.length + ' group' + (gNames.length !== 1 ? 's' : '') + ' detected: ' +            gNames.map(g => esc(g) + ' (n=' + gMap[g].length + ')').join(', ') + '. ';
    if (underRep.length) {
      gMsg   += esc(underRep.join(', ')) + (underRep.length > 1 ? ' have' : ' has') +                 ' fewer than 3 replicates. Each class must contain at least 3 replicates for reliable statistical analysis.';       gStatus = 'warn';
    } else {
      gMsg += 'Each group has at least 3 replicates.';
    }
  }
  report.checks.push({ id: 'groups', label: 'Group / class labels', status: gStatus, message: gMsg });

  /* ── 4. Paired samples ── */
  report.checks.push({
    id: 'paired', label: 'Paired samples', status: 'pass',     message: 'Samples are not paired.',
  });

  /* ── 5. Non-numeric values ── */
  const nNonNum = countNonNumeric(raw.rows, qCols);
  report.nNonNumeric = nNonNum;
  report.checks.push({
    id: 'numeric', label: 'Non-numeric values', status: nNonNum > 0 ? 'warn' : 'pass',
    message: nNonNum > 0
      ? 'Non-numeric values were found and replaced by NA. ' +         nNonNum.toLocaleString() + ' cell' + (nNonNum !== 1 ? 's' : '') + ' affected.'       : 'No non-numeric values detected.',
  });

  /* ── 6. Missing values ── */
  const nMiss = countMissing(raw.rows, qCols);
  report.nMissing = nMiss;
  const pct = report.totalDataPoints > 0
    ? ((nMiss / report.totalDataPoints) * 100).toFixed(1) : '0.0';
  report.missPct = pct;
  report.checks.push({
    id: 'missing', label: 'Missing values', status: nMiss > 0 ? 'warn' : 'pass',
    message: nMiss === 0
      ? 'No missing values detected.'       : 'A total of ' + nMiss.toLocaleString() + ' (' + pct + '%) missing values ' +         'detected across the quantitative matrix. ' +         'Imputation and removal options are available in Data Processing.',
  });

  /* ── 7. Constant / zero features ── */
  const constIdx = findConstantFeatures(raw.rows, qCols);
  report.constantFeatureIndices = constIdx;
  report.checks.push({
    id: 'constant', label: 'Constant / zero-value features', status: constIdx.length > 0 ? 'warn' : 'pass',
    message: constIdx.length === 0
      ? 'No constant or zero-value features detected.'       : constIdx.length.toLocaleString() + ' feature' + (constIdx.length !== 1 ? 's' : '') +         ' with a constant or single value across all samples were found and flagged. ' +         'These can be removed in Data Processing.',
  });

  /* ── Overall status ── */
  if (report.checks.some(c => c.status === 'err'))       report.overallStatus = 'err';   else if (report.checks.some(c => c.status === 'warn')) report.overallStatus = 'warn';

  return report;
}

/* ──────────────────────────────────────────────────────────────
   RENDER
   ────────────────────────────────────────────────────────────── */
function renderTab2() {
  const mc = document.getElementById('main-content');
  mc.innerHTML =
    '<div class="tab-content">' +       '<div class="tab-hd">' +         '<h2 class="tab-title">Data Check</h2>' +         '<p class="tab-desc">Automatic quality checks are run on your uploaded dataset. ' +         'Review the results below before proceeding to Data Processing.</p>' +       '</div>' +       '<div id="check-body"><div class="notice notice-info">&#x23F3; Running checks&hellip;</div></div>' +     '</div>';

  /* Defer so the loading notice renders before the (potentially slow) check */
  setTimeout(() => {
    if (!App.check) App.check = runDataCheck();
    renderCheckReport(App.check);
  }, 30);
}

function renderCheckReport(rpt) {
  const body = document.getElementById('check-body');
  if (!body) return;

  /* ── Stat cards ── */
  function sc(val, lbl, warn) {
    return '<div class="stat-card' + (warn ? ' stat-card-warn' : '') + '">' +       '<div class="stat-val">' + val + '</div>' +       '<div class="stat-lbl">' + lbl + '</div>' +     '</div>';
  }
  const missPctNum = parseFloat(rpt.missPct);
  const statsHTML =
    '<div class="stat-grid">' +       sc(rpt.nProteins.toLocaleString(), 'Proteins / Features', false) +       sc(rpt.nSamples, 'Samples', false) +       sc(rpt.nGroups !== null ? rpt.nGroups : '&mdash;', 'Groups Detected', false) +       sc(rpt.missPct + '%', 'Missing Values', missPctNum > 0) +       sc(rpt.constantFeatureIndices.length.toLocaleString(), 'Constant Features', rpt.constantFeatureIndices.length > 0) +     '</div>';

  /* ── Check items ── */
  const itemsHTML = rpt.checks.map(c =>
    '<div class="chk-item chk-item-' + c.status + '">' +
      chkBadge(c.status) +
      '<div class="chk-item-body">' +         '<div class="chk-item-title">' + esc(c.label) + '</div>' +         '<div class="chk-item-desc">' + c.message + '</div>' +       '</div>' +     '</div>'   ).join('');

  /* ── Summary log ── */
  const logParts = [
    'Checking data content ... ' + rpt.overallStatus + '.',     'Samples are in columns and features are in rows.',     'The uploaded file is in ' + fileTypeLabel(rpt.fileType) + ' format.',     'The dataset contains ' + rpt.nSamples + ' sample' + (rpt.nSamples !== 1 ? 's' : '') +       ' and ' + rpt.nProteins.toLocaleString() + ' protein' + (rpt.nProteins !== 1 ? 's' : '') + '.',     'Samples are not paired.',
    rpt.nGroups
      ? rpt.nGroups + ' group' + (rpt.nGroups !== 1 ? 's' : '') + ' detected in samples.'       : 'No distinct group structure detected in sample names.',
    rpt.nModified > 0
      ? 'Only English letters, numbers, underscore, hyphen and forward slash (/) are allowed. Other special characters or punctuations, if any, were stripped off.'       : 'Sample names are valid.',
    rpt.nNonNumeric > 0
      ? 'Non-numeric values were found and replaced by NA.'       : 'No non-numeric values were detected.',
    rpt.constantFeatureIndices.length > 0
      ? rpt.constantFeatureIndices.length.toLocaleString() + ' features with a constant or single value across samples were found and flagged.'       : 'No constant or zero-value features detected.',
    rpt.nMissing > 0
      ? 'A total of ' + rpt.nMissing.toLocaleString() + ' (' + rpt.missPct + '%) missing values were detected.'       : 'No missing values were detected.',
  ];

  /* ── Overall banner ── */
  const banners = {
    pass: ['notice-ok',   '&#x2714; All checks passed. You can proceed to Data Processing.'],     warn: ['notice-warn', '&#x26A0; Data check complete with warnings. Review items above, then proceed when ready.'],     err:  ['notice-err',  '&#x2718; Critical errors detected. Please fix the issues above before proceeding.'],
  };
  const [bannerCls, bannerMsg] = banners[rpt.overallStatus];
  const canProceed = rpt.overallStatus !== 'err';

  body.innerHTML =
    statsHTML +

    '<div class="card">' +       '<div class="card-title">Check Results</div>' +       '<div class="chk-list">' + itemsHTML + '</div>' +     '</div>' +

    '<div class="card">' +       '<div class="card-title">Summary Log</div>' +       '<div class="log-block">' +         '<div class="log-title">Data processing information</div>' +         '<div class="log-text">' + logParts.join(' ') + '</div>' +       '</div>' +     '</div>' +

    '<div class="card">' +       '<div class="notice ' + bannerCls + '" style="margin-bottom:18px">' + bannerMsg + '</div>' +       '<div class="btn-row">' +
        (canProceed
          ? '<button class="btn btn-primary btn-lg" onclick="proceedTab2()">Proceed to Data Processing &#8594;</button>'           : '') +         '<button class="btn btn-ghost" onclick="App.check=null;goTab(1)">&#8592; Back to Upload</button>' +       '</div>' +     '</div>';
}

/* ── Proceed to Tab 3 ── */
function proceedTab2() {
  if (!App.check) { showToast('Data check not complete.'); return; }   if (App.check.overallStatus === 'err') { showToast('Fix errors before proceeding.'); return; }
  unlockTab(3);
  goTab(3);
}

/* ====== _v2_tab3.js ====== */
/* ==============================================================
   TAB 3  —  Data Filtering
   Overrides stub renderTab3() from _v2_stubs.js.
   ============================================================== */

/* ── Module globals (persist across tab navigations, reset on re-upload) ── */
let _t3Metrics = null;   /* per-protein computed stats */
let _t3FS      = null;   /* filter state */
let _t3Charts  = {};     /* canvasId → Chart instance */

/* ══════════════════════════════════════════════════════════════
   STATISTICS HELPERS
══════════════════════════════════════════════════════════════ */
function _srt(arr) { return arr.slice().sort((a, b) => a - b); }

function _med(s) {   /* expects sorted array */
  const n = s.length;
  if (!n) return NaN;
  return n % 2 ? s[Math.floor(n / 2)] : (s[n / 2 - 1] + s[n / 2]) / 2;
}

function _mn(arr) {
  return arr.length ? arr.reduce((a, v) => a + v, 0) / arr.length : NaN;
}

function _sd(arr) {
  if (arr.length < 2) return 0;
  const m = _mn(arr);
  return Math.sqrt(arr.reduce((s, v) => s + (v - m) ** 2, 0) / (arr.length - 1));
}

function _iqr(arr) {
  const s = _srt(arr);
  const n = s.length;
  if (n < 2) return 0;
  return _med(s.slice(Math.ceil(n / 2))) - _med(s.slice(0, Math.floor(n / 2)));
}

function _mad(arr) {
  if (!arr.length) return 0;
  const med = _med(_srt(arr));
  return _med(_srt(arr.map(v => Math.abs(v - med))));
}

function _rsd(arr) {
  const m = _mn(arr);
  return (!isFinite(m) || m === 0) ? 0 : (_sd(arr) / Math.abs(m)) * 100;
}

/* ══════════════════════════════════════════════════════════════
   PER-PROTEIN METRICS  (computed once, cached in _t3Metrics)
══════════════════════════════════════════════════════════════ */
function computeProteinMetrics(rows, qCols) {
  return rows.map((row, rowIdx) => {
    const raw = qCols.map(c => {
      const v = String(row[c] !== undefined ? row[c] : '').trim();
      return (!v || MISSING_RE.test(v)) ? null : parseFloat(v);
    });
    const nums   = raw.filter(v => v !== null && !isNaN(v));
    const nMiss  = raw.length - nums.length;
    const missPct = raw.length ? (nMiss / raw.length) * 100 : 0;
    return {
      rowIdx,
      missPct,
      iqr:       _iqr(nums),
      sd:        _sd(nums),
      mad:       _mad(nums),
      rsd:       _rsd(nums),
      meanAbund: _mn(nums),
      medAbund:  _med(_srt(nums)),
      isConstant: nums.length > 0 && new Set(nums).size <= 1,
    };
  });
}

/* ══════════════════════════════════════════════════════════════
   FILTER STATE
══════════════════════════════════════════════════════════════ */
function _t3DefaultFS(qCols, nConst) {
  const qcCols  = qCols.filter(c => /qc|quality.?control|pool/i.test(c));
  const blkCols = qCols.filter(c => /^blank|^blk|^buffer/i.test(c));
  return {
    missing:   { enabled: false, threshold: 50 },
    variance:  { enabled: false, metric: 'iqr', pct: 40 },     abundance: { enabled: false, metric: 'mean', pct: 0 },
    constant:  { enabled: nConst > 0 },
    qcCols, blkCols,
  };
}

/* ══════════════════════════════════════════════════════════════
   APPLY ALL FILTERS
══════════════════════════════════════════════════════════════ */
function applyFiltersT3(metrics, fs) {
  const out = { missing: new Set(), variance: new Set(), abundance: new Set(), constant: new Set() };

  /* 1. Missing value */
  if (fs.missing.enabled)
    metrics.forEach(m => { if (m.missPct > fs.missing.threshold) out.missing.add(m.rowIdx); });

  /* 2. Low variance (rank-based, among survivors of step 1) */
  if (fs.variance.enabled && fs.variance.pct > 0) {
    const key  = fs.variance.metric;
    const pool = metrics.filter(m => !out.missing.has(m.rowIdx));
    const cut  = Math.floor(pool.length * fs.variance.pct / 100);
    pool.slice().sort((a, b) => a[key] - b[key]).slice(0, cut)
        .forEach(m => out.variance.add(m.rowIdx));
  }

  /* 3. Low abundance (rank-based, among survivors of steps 1–2) */
  if (fs.abundance.enabled && fs.abundance.pct > 0) {
    const key  = fs.abundance.metric === 'median' ? 'medAbund' : 'meanAbund';
    const skip = idx => out.missing.has(idx) || out.variance.has(idx);
    const pool = metrics.filter(m => !skip(m.rowIdx));
    const cut  = Math.floor(pool.length * fs.abundance.pct / 100);
    pool.slice().sort((a, b) => a[key] - b[key]).slice(0, cut)
        .forEach(m => out.abundance.add(m.rowIdx));
  }

  /* 4. Constant features (among survivors of steps 1–3) */
  if (fs.constant.enabled) {
    const skip = idx => out.missing.has(idx) || out.variance.has(idx) || out.abundance.has(idx);
    metrics.forEach(m => { if (m.isConstant && !skip(m.rowIdx)) out.constant.add(m.rowIdx); });
  }

  const all = new Set([...out.missing, ...out.variance, ...out.abundance, ...out.constant]);
  return {
    nBefore:     metrics.length,
    nRemoved:    all.size,
    nKept:       metrics.length - all.size,
    removed:     out,
    keptIndices: metrics.filter(m => !all.has(m.rowIdx)).map(m => m.rowIdx),
  };
}

/* ══════════════════════════════════════════════════════════════
   HISTOGRAM UTILITIES
══════════════════════════════════════════════════════════════ */
function _histBins(values, nBins) {
  const mn = Math.min(...values), mx = Math.max(...values);
  const rng = (mx - mn) || 1e-9;
  const bsz = rng / nBins;
  const bins  = Array(nBins).fill(0);
  const edges = Array.from({ length: nBins }, (_, i) => mn + i * bsz);
  values.forEach(v => bins[Math.min(Math.floor((v - mn) / bsz), nBins - 1)]++);
  return { bins, edges, bsz };
}

function _histColors(edges, bsz, cutoff, removeRight) {
  /* removeRight=true  → bins whose center > cutoff are red (missing value chart)  */
  /* removeRight=false → bins whose center < cutoff are red (variance/abundance)   */
  return edges.map(e => {
    const c = e + bsz / 2;
    return (removeRight ? c > cutoff : c < cutoff)
      ? 'rgba(220,38,38,0.68)' : 'rgba(37,99,235,0.52)';
  });
}

function _buildHistChart(canvasId, values, cutoff, removeRight, xlabel) {
  if (_t3Charts[canvasId]) { _t3Charts[canvasId].destroy(); delete _t3Charts[canvasId]; }
  const el = document.getElementById(canvasId);
  if (!el || !values.length || typeof Chart === 'undefined') return null;

  const nBins = Math.max(10, Math.min(30, Math.ceil(Math.sqrt(values.length))));
  const { bins, edges, bsz } = _histBins(values, nBins);

  const chart = new Chart(el.getContext('2d'), {     type: 'bar',
    data: {
      labels: edges.map(e => e.toFixed(2)),
      datasets: [{
        data: bins,
        backgroundColor: _histColors(edges, bsz, cutoff, removeRight),
        borderWidth: 0,
        barPercentage: 1.0,
        categoryPercentage: 1.0,
      }],
    },
    options: {
      animation: false,
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { display: false },
        tooltip: {
          callbacks: {
            title: items => xlabel + ': ~' + parseFloat(items[0].label).toFixed(2),             label: item  => 'Proteins: ' + item.raw,
          },
        },
      },
      scales: {
        x: {
          ticks: { maxTicksLimit: 5, font: { size: 12 } },
          title: { display: true, text: xlabel, font: { size: 12 }, color: '#64748b' },
          grid: { display: false },
        },
        y: {
          ticks: { font: { size: 12 } },
          title: { display: true, text: 'Count', font: { size: 12 }, color: '#64748b' },
          beginAtZero: true,
        },
      },
    },
  });
  chart._t3 = { edges, bsz, removeRight };
  _t3Charts[canvasId] = chart;
  return chart;
}

function _updateHistCutoff(canvasId, cutoff) {
  const c = _t3Charts[canvasId];
  if (!c?._t3) return;
  c.data.datasets[0].backgroundColor = _histColors(c._t3.edges, c._t3.bsz, cutoff, c._t3.removeRight);
  c.update('none');
}

/* Compute the value at the Xth percentile of a sorted array */
function _pctVal(sorted, pct) {
  if (!sorted.length || pct <= 0) return -Infinity;
  if (pct >= 100) return Infinity;
  return sorted[Math.floor(sorted.length * pct / 100)] ?? sorted[sorted.length - 1];
}

/* ══════════════════════════════════════════════════════════════
   BUILD ALL 3 CHARTS
══════════════════════════════════════════════════════════════ */
function _t3BuildCharts() {
  if (!_t3Metrics) return;
  const fs = _t3FS;

  /* Missing value chart  (remove right of cutoff) */
  const missPcts = _t3Metrics.map(m => m.missPct);
  _buildHistChart('t3-chart-miss', missPcts, fs.missing.threshold, true, 'Missing Value (%)');

  /* Variance chart  (remove left of cutoff) */
  const varKey  = fs.variance.metric;
  const varVals = _t3Metrics.map(m => m[varKey]).filter(isFinite);
  const varCut  = _pctVal(varVals.slice().sort((a, b) => a - b), fs.variance.pct);
  _buildHistChart('t3-chart-var', varVals, varCut, false, varKey.toUpperCase());

  /* Abundance chart  (remove left of cutoff) */
  const abKey  = fs.abundance.metric === 'median' ? 'medAbund' : 'meanAbund';
  const abVals = _t3Metrics.map(m => m[abKey]).filter(isFinite);
  const abCut  = _pctVal(abVals.slice().sort((a, b) => a - b), fs.abundance.pct);
  _buildHistChart('t3-chart-abund', abVals, abCut, false,     fs.abundance.metric === 'median' ? 'Median Abundance' : 'Mean Abundance');
}

/* ══════════════════════════════════════════════════════════════
   LIVE SUMMARY BAR
══════════════════════════════════════════════════════════════ */
function _t3Sync() {
  if (!_t3Metrics || !_t3FS) return;
  const r  = applyFiltersT3(_t3Metrics, _t3FS);
  const se = id => document.getElementById(id);
  if (se('t3-nbefore'))  se('t3-nbefore').textContent  = r.nBefore.toLocaleString();   if (se('t3-nremoved')) se('t3-nremoved').textContent = r.nRemoved.toLocaleString();   if (se('t3-nkept'))    se('t3-nkept').textContent    = r.nKept.toLocaleString();
  const frac = r.nBefore > 0 ? r.nKept / r.nBefore : 1;
  if (se('t3-nkept')) se('t3-nkept').style.color = frac < 0.3 ? '#dc2626' : frac < 0.6 ? '#b45309' : '#16a34a';
}

/* ══════════════════════════════════════════════════════════════
   EVENT HANDLERS
══════════════════════════════════════════════════════════════ */
function t3MissEn(cb)   { _t3FS.missing.enabled = cb.checked; _t3Sync(); }
function t3MissSl(v)    {
  _t3FS.missing.threshold = +v;
  const l = document.getElementById('t3-miss-lbl'); if (l) l.textContent = v + '%';   _updateHistCutoff('t3-chart-miss', +v);
  _t3Sync();
}

function t3VarEn(cb)    { _t3FS.variance.enabled = cb.checked; _t3Sync(); }
function t3VarMet(v)    {
  _t3FS.variance.metric = v;
  const vals = _t3Metrics.map(m => m[v]).filter(isFinite);
  const cut  = _pctVal(vals.slice().sort((a, b) => a - b), _t3FS.variance.pct);
  _buildHistChart('t3-chart-var', vals, cut, false, v.toUpperCase());
  _t3Sync();
}
function t3VarSl(v)     {
  _t3FS.variance.pct = +v;
  const l = document.getElementById('t3-var-lbl'); if (l) l.textContent = v + '%';
  const key  = _t3FS.variance.metric;
  const vals = _t3Metrics.map(m => m[key]).filter(isFinite);
  const cut  = _pctVal(vals.slice().sort((a, b) => a - b), +v);
  _updateHistCutoff('t3-chart-var', cut);
  _t3Sync();
}

function t3AbEn(cb)     { _t3FS.abundance.enabled = cb.checked; _t3Sync(); }
function t3AbMet(v)     {
  _t3FS.abundance.metric = v;
  const key  = v === 'median' ? 'medAbund' : 'meanAbund';
  const vals = _t3Metrics.map(m => m[key]).filter(isFinite);
  const cut  = _pctVal(vals.slice().sort((a, b) => a - b), _t3FS.abundance.pct);
  _buildHistChart('t3-chart-abund', vals, cut, false, v === 'median' ? 'Median Abundance' : 'Mean Abundance');
  _t3Sync();
}
function t3AbSl(v)      {
  _t3FS.abundance.pct = +v;
  const l = document.getElementById('t3-abund-lbl'); if (l) l.textContent = v + '%';   const key  = _t3FS.abundance.metric === 'median' ? 'medAbund' : 'meanAbund';
  const vals = _t3Metrics.map(m => m[key]).filter(isFinite);
  const cut  = _pctVal(vals.slice().sort((a, b) => a - b), +v);
  _updateHistCutoff('t3-chart-abund', cut);
  _t3Sync();
}

function t3ConstEn(cb)  { _t3FS.constant.enabled = cb.checked; _t3Sync(); }

/* ══════════════════════════════════════════════════════════════
   APPLY / RESET / PROCEED
══════════════════════════════════════════════════════════════ */
function t3Apply() {
  const r  = applyFiltersT3(_t3Metrics, _t3FS);
  const fs = _t3FS;

  /* Reset Tab 4 cache since filtered set changed */
  if (typeof _t4MV !== 'undefined') _t4MV = null;
  App.imputed = null;

  /* Save filtered dataset + filter settings for Methods report */
  App.filtered = {
    rows:      r.keptIndices.map(i => App.raw.rows[i]),
    headers:   App.raw.headers,
    quantCols: App.config.quantCols,
    nRows:     r.nKept,
    filterSettings: {
      missing:   { enabled: fs.missing.enabled,   threshold: fs.missing.threshold },
      variance:  { enabled: fs.variance.enabled,   metric: fs.variance.metric,   pct: fs.variance.pct },
      abundance: { enabled: fs.abundance.enabled,  metric: fs.abundance.metric,  pct: fs.abundance.pct },
      constant:  { enabled: fs.constant.enabled },
    },
    filterResults: {
      nBefore:  r.nBefore,
      nKept:    r.nKept,
      nRemoved: r.nRemoved,
      removedMissing:   r.removed.missing.size,
      removedVariance:  r.removed.variance.size,
      removedAbundance: r.removed.abundance.size,
      removedConstant:  r.removed.constant.size,
    },
  };

  const VL = { iqr:'IQR', sd:'SD', mad:'MAD', rsd:'RSD' };   const AL = { mean:'Mean intensity', median:'Median intensity' };

  const lines = [
    fs.missing.enabled
      ? 'Missing value filter enabled (threshold: ' + fs.missing.threshold + '%) — ' +         r.removed.missing.size.toLocaleString() + ' protein' + (r.removed.missing.size !== 1 ? 's' : '') + ' removed.'       : 'Missing value filter: <em>disabled</em>.',
    fs.variance.enabled
      ? 'Low variance filter (' + VL[fs.variance.metric] + ', bottom ' + fs.variance.pct + '%) — ' +         r.removed.variance.size.toLocaleString() + ' protein' + (r.removed.variance.size !== 1 ? 's' : '') + ' removed.'       : 'Low variance filter: <em>disabled</em>.',
    fs.abundance.enabled
      ? 'Low abundance filter (' + AL[fs.abundance.metric] + ', bottom ' + fs.abundance.pct + '%) — ' +         r.removed.abundance.size.toLocaleString() + ' protein' + (r.removed.abundance.size !== 1 ? 's' : '') + ' removed.'       : 'Low abundance filter: <em>disabled</em>.',
    fs.constant.enabled
      ? 'Constant feature filter — ' + r.removed.constant.size.toLocaleString() +         ' protein' + (r.removed.constant.size !== 1 ? 's' : '') + ' removed.'       : 'Constant feature filter: <em>disabled</em>.',
  ];

  const el = document.getElementById('t3-report');
  if (el) {
    el.style.display = '';
    el.innerHTML =
      '<div class="card-title">Filtering Summary</div>' +       '<div class="t3-rpt-lines">' +         lines.map(l => '<div class="t3-rpt-line">&#x2022; ' + l + '</div>').join('') +         '<div class="t3-rpt-final">' +           '&#x2714; <strong>' + r.nKept.toLocaleString() + ' proteins retained</strong>' +           ' &nbsp;|&nbsp; ' + r.nRemoved.toLocaleString() + ' removed from ' +           r.nBefore.toLocaleString() + ' total.' +         '</div>' +       '</div>';
  }
  showToast(r.nKept.toLocaleString() + ' proteins retained after filtering.');
}

function t3Reset() {
  const nConst = _t3Metrics ? _t3Metrics.filter(m => m.isConstant).length : 0;
  _t3FS = _t3DefaultFS(App.config.quantCols, nConst);
  App.filtered = null;
  _renderTab3Body();
}

function proceedTab3() {
  if (!App.filtered) t3Apply();
  unlockTab(4);
  goTab(4);
}

/* ══════════════════════════════════════════════════════════════
   RENDER
══════════════════════════════════════════════════════════════ */
function renderTab3() {
  document.getElementById('main-content').innerHTML =     '<div class="tab-content">' +       '<div class="tab-hd">' +         '<h2 class="tab-title">Data Filtering</h2>' +         '<p class="tab-desc">Remove proteins unlikely to be informative for downstream analysis. ' +         'Enable filters independently and adjust thresholds. ' +         'Protein counts update in real time.</p>' +       '</div>' +       '<div id="t3-body"><div class="notice notice-info">&#x23F3; Preparing filters&hellip;</div></div>' +     '</div>';

  setTimeout(() => {
    if (!_t3Metrics) _t3Metrics = computeProteinMetrics(App.raw.rows, App.config.quantCols);
    const nConst = _t3Metrics.filter(m => m.isConstant).length;
    if (!_t3FS) _t3FS = _t3DefaultFS(App.config.quantCols, nConst);
    App.filtered = null;
    _renderTab3Body();
  }, 30);
}

function _renderTab3Body() {
  Object.values(_t3Charts).forEach(c => { try { c.destroy(); } catch(e){} });
  _t3Charts = {};

  const fs     = _t3FS;
  const nConst = _t3Metrics.filter(m => m.isConstant).length;
  const hasQC  = fs.qcCols.length > 0;
  const hasBlk = fs.blkCols.length > 0;
  const N      = App.raw.nRows;
  const nS     = App.config.quantCols.length;

  /* ── Local helpers ── */
  function chk(id, checked, handler, label) {
    return '<label class="t3-en-lbl"><input type="checkbox" id="' + id +            '" onchange="' + handler + '(this)"' + (checked ? ' checked' : '') +            '> <span>' + label + '</span></label>';
  }
  function slider(id, lblId, min, max, val, fn) {
    return '<div class="t3-sl-row">' +       '<input type="range" class="t3-sl" id="' + id + '" min="' + min +       '" max="' + max + '" value="' + val + '" step="1" oninput="' + fn + '(this.value)">' +       '<span class="t3-sl-val" id="' + lblId + '">' + val + '%</span></div>';
  }
  /* Compact pill-style radios — horizontal, no vertical stack */
  function pills(name, opts, cur, fn) {
    return '<div class="t3-pills">' +
      opts.map(([v, l]) =>
        '<label class="t3-pill' + (cur === v ? ' t3-pill-sel' : '') + '">' +           '<input type="radio" name="' + name + '" value="' + v + '"' +           (cur === v ? ' checked' : '') + ' onchange="' + fn + '(this.value)">' +           l + '</label>'       ).join('') + '</div>';
  }
  function chartBox(id) {
    return '<div class="t3-chart-box">' +       '<canvas id="' + id + '"></canvas>' +       '<div class="t3-chart-key">' +         '<span class="t3-key-rem">&#x25A0; Removed</span>' +         '<span class="t3-key-kpt">&#x25A0; Retained</span>' +       '</div></div>';
  }

  document.getElementById('t3-body').innerHTML =

    /* ── Live summary bar ── */
    '<div class="t3-sumbar">' +       '<div class="t3-sum-item"><span class="t3-sum-n" id="t3-nbefore">' + N.toLocaleString() + '</span><span class="t3-sum-lbl">Before filtering</span></div>' +       '<div class="t3-sum-sep">&#x2212;</div>' +       '<div class="t3-sum-item"><span class="t3-sum-n t3-sum-red" id="t3-nremoved">0</span><span class="t3-sum-lbl">Removed</span></div>' +       '<div class="t3-sum-sep">&#x3d;</div>' +       '<div class="t3-sum-item"><span class="t3-sum-n" id="t3-nkept" style="color:#16a34a">' + N.toLocaleString() + '</span><span class="t3-sum-lbl">Remaining</span></div>' +     '</div>' +

    /* ── 2×2 Filter Dashboard ── */
    '<div class="t3-dash-grid">' +

      /* Top-left: Missing Value */
      '<div class="card t3-dash-card">' +         '<div class="t3-fhead">' +           chk('t3-miss-en', fs.missing.enabled, 't3MissEn', 'Missing Value Filter') +           '<span class="t3-fhead-sub">Remove if missing in &gt;X% of samples</span>' +         '</div>' +         '<div class="t3-dash-ctrl">' +           '<div class="t3-ctrl-lbl">Threshold:</div>' +           slider('t3-miss-sl', 't3-miss-lbl', 0, 100, fs.missing.threshold, 't3MissSl') +         '</div>' +         chartBox('t3-chart-miss') +       '</div>' +

      /* Top-right: Low Variance */
      '<div class="card t3-dash-card">' +         '<div class="t3-fhead">' +           chk('t3-var-en', fs.variance.enabled, 't3VarEn', 'Low Variance Filter') +           '<span class="t3-fhead-sub">Remove bottom X% by variance metric</span>' +         '</div>' +         '<div class="t3-dash-ctrl">' +           pills('t3vmet', [['iqr','IQR'],['sd','SD'],['mad','MAD'],['rsd','RSD']], fs.variance.metric, 't3VarMet') +           slider('t3-var-sl', 't3-var-lbl', 0, 100, fs.variance.pct, 't3VarSl') +         '</div>' +         chartBox('t3-chart-var') +       '</div>' +

      /* Bottom-left: Low Abundance */
      '<div class="card t3-dash-card">' +         '<div class="t3-fhead">' +           chk('t3-ab-en', fs.abundance.enabled, 't3AbEn', 'Low Abundance Filter') +           '<span class="t3-fhead-sub">Remove bottom X% by intensity</span>' +         '</div>' +         '<div class="t3-dash-ctrl">' +           pills('t3amet', [['mean','Mean'],['median','Median']], fs.abundance.metric, 't3AbMet') +           slider('t3-ab-sl', 't3-abund-lbl', 0, 100, fs.abundance.pct, 't3AbSl') +         '</div>' +         chartBox('t3-chart-abund') +       '</div>' +

      /* Bottom-right: Constant Features + dataset info */
      '<div class="card t3-dash-card">' +         '<div class="t3-fhead">' +           chk('t3-const-en', fs.constant.enabled, 't3ConstEn', 'Constant Feature Filter') +           '<span class="t3-fhead-sub">Remove all-zero or single-value proteins</span>' +         '</div>' +         '<div class="t3-dash-ctrl">' +           '<div class="notice ' + (nConst > 0 ? 'notice-warn' : 'notice-ok') + '" style="margin:0 0 10px">' +
            (nConst > 0
              ? '&#x26A0; ' + nConst.toLocaleString() + ' constant protein' + (nConst !== 1 ? 's' : '') + ' detected'               : '&#x2714; No constant proteins detected') +           '</div>' +           '<div class="t3-info-grid">' +             '<span class="t3-ig-lbl">Dataset</span><span class="t3-ig-val">' + N.toLocaleString() + ' proteins &times; ' + nS + ' samples</span>' +             '<span class="t3-ig-lbl">QC samples</span><span class="t3-ig-val">' + (hasQC ? fs.qcCols.length + ' detected' : 'None') + '</span>' +             '<span class="t3-ig-lbl">Blank samples</span><span class="t3-ig-val">' + (hasBlk ? fs.blkCols.length + ' detected' : 'None') + '</span>' +           '</div>' +         '</div>' +       '</div>' +

    '</div>' + /* end t3-dash-grid */

    /* ── Report (shown after Apply) ── */
    '<div class="card" id="t3-report" style="display:none"></div>' +

    /* ── Action row ── */
    '<div class="card">' +       '<div class="btn-row">' +         '<button class="btn btn-primary btn-lg" onclick="t3Apply()">&#x2713; Apply Filters</button>' +         '<button class="btn btn-primary btn-lg" onclick="proceedTab3()" style="background:#7c3aed;border-color:#7c3aed">Proceed to Missing Value Estimation &#8594;</button>' +         '<button class="btn btn-ghost" onclick="t3Reset()">&#x21BA; Reset</button>' +         '<button class="btn btn-ghost" onclick="goTab(2)">&#8592; Back</button>' +       '</div>' +     '</div>';

  requestAnimationFrame(() => {
    requestAnimationFrame(() => _t3BuildCharts());
  });
  _t3Sync();
}

/* ====== _v2_tab4.js ====== */
/* ==============================================================
   TAB 4  —  Missing Value Estimation
   Overrides stub renderTab4() from _v2_stubs.js.
   ============================================================== */

/* ── Module globals ── */
let _t4Method = 'lod';   /* selected imputation method */
let _t4Charts = {};      /* canvasId → Chart instance */
let _t4MV     = null;    /* missing-value analysis cache */

/* ── Annotation sets (built lazily from global strings) ── */
let _t4SS = null, _t4PM = null, _t4MA = null;
function _t4Annot() {
  if (_t4SS) return;
  _t4SS = new Set(_SS_RAW.split('\n').map(s => s.trim()).filter(Boolean));   _t4PM = new Set(_PM_RAW.split('\n').map(s => s.trim()).filter(Boolean));   _t4MA = new Set(_MA_RAW.split('\n').map(s => s.trim()).filter(Boolean));
}
function _t4Category(pid) {
  _t4Annot();
  const id = pid.split(/[|;\s]/)[0].trim();
  if (_t4SS.has(id)) return 'SURFY';   if (_t4PM.has(id)) return 'UniProt Cell Membrane';   if (_t4MA.has(id)) return 'Membrane-Associated';   return 'Other';
}

/* ══════════════════════════════════════════════════════════════
   ANALYSIS — per-protein missing value stats
══════════════════════════════════════════════════════════════ */
function _t4Analyze() {
  const src   = App.filtered || App.raw;
  const rows  = src.rows;
  const qCols = App.config.quantCols;
  const idCol = App.config.idCol;

  let totalMiss = 0, proteinsWithMiss = 0, maxMissPct = 0;
  const catMiss  = { SURFY: 0, 'UniProt Cell Membrane': 0, 'Membrane-Associated': 0, Other: 0 };   const catTotal = { SURFY: 0, 'UniProt Cell Membrane': 0, 'Membrane-Associated': 0, Other: 0 };
  const perProtein = [];

  _t4Annot();

  rows.forEach((row, rowIdx) => {
    const pid = String(row[idCol] || '').trim();
    const cat = _t4Category(pid);
    let nMiss = 0;
    qCols.forEach(c => {
      const v = String(row[c] !== undefined ? row[c] : '').trim();
      if (!v || MISSING_RE.test(v)) nMiss++;
    });
    totalMiss += nMiss;
    catMiss[cat]  += nMiss;
    catTotal[cat] += 1;
    if (nMiss > 0) proteinsWithMiss++;
    const missPct = qCols.length ? (nMiss / qCols.length) * 100 : 0;
    if (missPct > maxMissPct) maxMissPct = missPct;
    perProtein.push({ pid, missPct, nMiss, category: cat, rowIdx });
  });

  const nTotal    = rows.length * qCols.length;
  const overallPct = nTotal > 0 ? (totalMiss / nTotal * 100).toFixed(1) : '0.0';

  return { totalMiss, overallPct, proteinsWithMiss, maxMissPct: maxMissPct.toFixed(1),
           perProtein, catMiss, catTotal, nRows: rows.length, nCols: qCols.length };
}

/* ══════════════════════════════════════════════════════════════
   IMPUTATION
══════════════════════════════════════════════════════════════ */
function _t4FillVal(method, nums) {
  /* nums: observed (non-null, finite) values for one protein */
  if (!nums.length) return 0;
  if (method === 'lod') {
    const pos  = nums.filter(v => v > 0);
    const minV = pos.length ? Math.min(...pos) : Math.min(...nums.map(Math.abs));
    return minV / 5;
  }
  if (method === 'halfmin') return Math.min(...nums) / 2;   if (method === 'mean')    return nums.reduce((a, v) => a + v, 0) / nums.length;
  return null;  /* KNN handled separately */
}

function _t4ImputeKNN(matrix, rows, qCols, k) {
  k = k || 5;
  const n = matrix.length;
  matrix.forEach((row, i) => {
    const missCols = row.reduce((a, v, j) => { if (v === null) a.push(j); return a; }, []);
    if (!missCols.length) return;
    const obsCols = row.reduce((a, v, j) => { if (v !== null) a.push(j); return a; }, []);
    if (!obsCols.length) return;

    /* Euclidean distance on shared observed columns */
    const dists = [];
    for (let ii = 0; ii < n; ii++) {
      if (ii === i) continue;
      let dist = 0, cnt = 0;
      obsCols.forEach(j => {
        if (matrix[ii][j] !== null) { dist += (row[j] - matrix[ii][j]) ** 2; cnt++; }
      });
      if (cnt > 0) dists.push({ ii, dist: dist / cnt });
    }
    dists.sort((a, b) => a.dist - b.dist);
    const neighbors = dists.slice(0, k);

    missCols.forEach(j => {
      const vals = neighbors.map(d => matrix[d.ii][j]).filter(v => v !== null);
      if (vals.length) {
        const impVal = vals.reduce((a, v) => a + v, 0) / vals.length;
        rows[i][qCols[j]] = impVal;
        matrix[i][j] = impVal;
      }
    });
  });
}

function _t4Impute(method) {
  const src   = App.filtered || App.raw;
  const rows  = src.rows.map(r => Object.assign({}, r));
  const qCols = App.config.quantCols;

  const matrix = rows.map(row =>
    qCols.map(c => {
      const v = String(row[c] !== undefined ? row[c] : '').trim();
      return (!v || MISSING_RE.test(v)) ? null : parseFloat(v);
    })
  );
  /* track which cells were originally missing */
  const missMask = matrix.map(row => row.map(v => v === null));

  if (method === 'knn') {
    _t4ImputeKNN(matrix, rows, qCols);
  } else {
    matrix.forEach((row, i) => {
      const nums = row.filter(v => v !== null && isFinite(v));
      const fv   = _t4FillVal(method, nums);
      if (fv === null || fv === undefined) return;
      row.forEach((v, j) => {
        if (v === null) { rows[i][qCols[j]] = fv; matrix[i][j] = fv; }
      });
    });
  }

  let nImputed = 0, proteinsAffected = 0;
  missMask.forEach((rowMask, i) => {
    let hit = false;
    rowMask.forEach((wasMiss, j) => {
      if (wasMiss && matrix[i][j] !== null) { nImputed++; hit = true; }
    });
    if (hit) proteinsAffected++;
  });

  return { rows, headers: src.headers, quantCols: qCols,
           nRows: rows.length, nImputed, proteinsAffected };
}

/* ══════════════════════════════════════════════════════════════
   RENDER
══════════════════════════════════════════════════════════════ */
function renderTab4() {
  document.getElementById('main-content').innerHTML =     '<div class="tab-content">' +       '<div class="tab-hd">' +         '<h2 class="tab-title">Missing Value Estimation</h2>' +         '<p class="tab-desc">Inspect the missing value landscape across your dataset, ' +         'select an imputation strategy, and preview the effect before applying.</p>' +       '</div>' +       '<div id="t4-body"><div class="notice notice-info">&#x23F3; Analyzing missing values&hellip;</div></div>' +     '</div>';

  setTimeout(() => {
    if (!_t4MV) _t4MV = _t4Analyze();
    _renderTab4Body();
  }, 30);
}

function _renderTab4Body() {
  Object.values(_t4Charts).forEach(c => { try { c.destroy(); } catch(e){} });
  _t4Charts = {};

  const mv         = _t4MV;
  const hasMissing = mv.totalMiss > 0;
  const src        = App.filtered || App.raw;
  const qCols      = App.config.quantCols;

  /* ── Category breakdown table ── */
  const cats = ['SURFY', 'UniProt Cell Membrane', 'Membrane-Associated', 'Other'];
  const catHTML = cats.map(cat => {
    const n   = mv.catTotal[cat] || 0;
    const m   = mv.catMiss[cat]  || 0;
    const pct = n > 0 ? (m / (n * mv.nCols) * 100).toFixed(1) : '0.0';     return '<tr><td>' + esc(cat) + '</td>' +       '<td class="num">' + n.toLocaleString() + '</td>' +       '<td class="num">' + m.toLocaleString() + '</td>' +       '<td class="num">' + pct + '%</td></tr>';   }).join('');

  /* ── Preview table (top 8 most-missing proteins) ── */
  const showCols   = qCols.slice(0, 5);
  const extraCols  = qCols.length - showCols.length;
  const previewTop = mv.perProtein
    .filter(p => p.nMiss > 0)
    .sort((a, b) => b.missPct - a.missPct)
    .slice(0, 8);

  const previewColHdr = '<tr><th>Protein</th><th>Missing</th>' +     showCols.map(c => '<th>' + esc(c.length > 10 ? c.slice(0, 10) + '…' : c) + '</th>').join('') +     (extraCols > 0 ? '<th class="t4-dim">+' + extraCols + ' more</th>' : '') + '</tr>';

  /* Compute preview "after" values per row using selected method (fast; only 8 proteins) */
  const previewBody = previewTop.map(p => {
    const row  = src.rows[p.rowIdx];
    const cols = showCols.map(c => {
      const v    = String(row[c] !== undefined ? row[c] : '').trim();
      const miss = !v || MISSING_RE.test(v);
      return { raw: v, miss, num: miss ? null : parseFloat(v) };
    });
    /* fill value for this protein */
    const allNums = qCols.map(c => {
      const v = String(row[c] !== undefined ? row[c] : '').trim();
      return (!v || MISSING_RE.test(v)) ? null : parseFloat(v);
    }).filter(v => v !== null && isFinite(v));
    const fv = (_t4Method !== 'knn') ? _t4FillVal(_t4Method, allNums) : null;

    const cells = cols.map(cell => {
      if (!cell.miss) return '<td class="t4-val-ok">' + esc(String(cell.num !== null ? cell.num.toPrecision(4) : cell.raw)) + '</td>';       const after = fv !== null ? '<span class="t4-imputed">' + fv.toPrecision(3) + '</span>' : '<span class="t4-imputed t4-knn-est">~KNN</span>';       return '<td class="t4-val-miss"><span class="t4-miss-dash">—</span> ' + after + '</td>';     }).join('');

    return '<tr><td class="t4-pid">' + esc(p.pid) + '</td>' +       '<td class="num t4-miss-pct">' + p.missPct.toFixed(0) + '%</td>' +
      cells +
      (extraCols > 0 ? '<td class="t4-dim">…</td>' : '') + '</tr>';   }).join('');

  document.getElementById('t4-body').innerHTML =

    /* ── Stat cards ── */
    '<div class="stat-grid">' +       _t4Sc(mv.totalMiss.toLocaleString(),          'Total Missing Values',          mv.totalMiss > 0) +       _t4Sc(mv.overallPct + '%',                    'Overall Missing (%)',            parseFloat(mv.overallPct) > 20) +       _t4Sc(mv.proteinsWithMiss.toLocaleString(),   'Proteins with Missing Values',  mv.proteinsWithMiss > 0) +       _t4Sc(mv.maxMissPct + '%',                    'Max Protein Missingness',        parseFloat(mv.maxMissPct) > 50) +     '</div>' +

    /* ── Category table ── */
    '<div class="card">' +       '<div class="card-title">Missing Values by Surfaceome Category</div>' +       '<div class="preview-wrap"><table class="ptable t4-cat-table">' +         '<thead><tr><th>Category</th><th class="num">Proteins</th>' +         '<th class="num">Missing Values</th><th class="num">Missing (%)</th></tr></thead>' +         '<tbody>' + catHTML + '</tbody>' +       '</table></div>' +     '</div>' +

    /* ── Visualizations ── */
    '<div class="card">' +       '<div class="card-title">Missing Value Patterns</div>' +       '<div class="t4-viz-row">' +         '<div class="t4-viz-cell">' +           '<div class="sec-title">Distribution of protein missingness</div>' +           '<div class="t4-chart-box"><canvas id="t4-hist"></canvas></div>' +         '</div>' +         '<div class="t4-viz-cell">' +           '<div class="sec-title">Top proteins by missingness (lollipop)</div>' +           '<div class="t4-chart-box"><canvas id="t4-lollipop"></canvas></div>' +         '</div>' +       '</div>' +       '<div style="margin-top:20px">' +         '<div class="sec-title">Missing value heatmap ' +           '<span class="t4-hm-note">(top 40 proteins × ' + Math.min(qCols.length, 30) + ' samples — red = missing, blue gradient = intensity)</span>' +         '</div>' +         '<div id="t4-heatmap-wrap" class="t4-heatmap-wrap"><canvas id="t4-heatmap"></canvas></div>' +         '<div class="t4-hm-legend">' +           '<span class="t4-hm-miss">&#x25A0; Missing</span>' +           '<span class="t4-hm-lo">&#x25A0; Low</span>' +           '<span style="display:inline-block;width:60px;height:11px;background:linear-gradient(to right,#bfdbfe,#1e3a8a);border-radius:2px;vertical-align:middle;margin:0 4px"></span>' +           '<span class="t4-hm-hi">&#x25A0; High</span>' +         '</div>' +       '</div>' +     '</div>' +

    /* ── Imputation method ── */
    '<div class="card" id="t4-method-card">' +       '<div class="card-title">Imputation Strategy</div>' +
      (hasMissing
        ? '<div class="t4-methods" id="t4-methods">' + _t4MethodCardsHTML() + '</div>'         : '<div class="notice notice-ok" style="margin-top:0">&#x2714; No missing values detected &mdash; no imputation required.</div>') +     '</div>' +

    /* ── Preview table ── */
    (hasMissing && previewBody
      ? '<div class="card">' +           '<div class="card-title">Imputation Preview <span class="card-meta">' +             '(— = missing &nbsp;|&nbsp; <span class="t4-imputed">green</span> = estimated fill value for current method)</span></div>' +           '<div class="preview-wrap"><table class="ptable t4-preview-tbl">' +             '<thead>' + previewColHdr + '</thead>' +             '<tbody>' + previewBody + '</tbody>' +           '</table></div>' +           '<p class="t4-preview-note">Showing up to 8 proteins with highest missingness &middot; ' +           'Preview updates when method changes</p>' +         '</div>'       : '') +

    /* ── Report (after apply) ── */
    '<div class="card" id="t4-report" style="display:none"></div>' +

    /* ── Actions ── */
    '<div class="card">' +       '<div class="btn-row">' +
        (hasMissing
          ? '<button class="btn btn-primary btn-lg" id="t4-apply-btn" onclick="t4Apply()">&#x2713; Apply Imputation</button>'           : '') +         '<button class="btn btn-primary btn-lg" onclick="proceedTab4()" style="background:#7c3aed;border-color:#7c3aed">Proceed to Normalization &#8594;</button>' +         '<button class="btn btn-ghost" onclick="goTab(3)">&#8592; Back</button>' +       '</div>' +     '</div>';

  requestAnimationFrame(() => requestAnimationFrame(() => {
    _t4BuildHist();
    _t4BuildLollipop();
    _t4BuildHeatmap();
  }));
}

/* ── Stat card helper ── */
function _t4Sc(val, lbl, warn) {
  return '<div class="stat-card' + (warn ? ' stat-card-warn' : '') + '">' +     '<div class="stat-val">' + val + '</div><div class="stat-lbl">' + lbl + '</div></div>';
}

/* ── Method selection cards ── */
function _t4MethodCardsHTML() {
  const methods = [
    { id: 'lod',     name: 'Limit of Detection (LoD)',   tag: 'Default',       desc: 'Replace missing values with 1/5 of the minimum positive value per protein. Suited for data missing due to values below the detection limit (MNAR).' },     { id: 'halfmin', name: 'Half Minimum',                tag: null,       desc: 'Replace missing values with half of the minimum observed value per protein. A simple, conservative approach widely used in proteomics.' },     { id: 'mean',    name: 'Mean Imputation',             tag: null,       desc: 'Replace missing values with the row mean (mean of observed values for that protein). Best when data is missing at random (MAR).' },     { id: 'knn',     name: 'K-Nearest Neighbors (KNN)',  tag: null,       desc: 'Estimate missing values from the k=5 most similar proteins by Euclidean distance on observed columns. Preserves local data structure.' },
  ];
  return methods.map(m =>
    '<label class="t4-mcard' + (_t4Method === m.id ? ' t4-mcard-sel' : '') + '">' +       '<input type="radio" name="t4meth" value="' + m.id + '"' +         (_t4Method === m.id ? ' checked' : '') +         ' onchange="t4SelMethod(\'' + m.id + '\')" style="display:none">' +       '<div class="t4-mcard-hd">' +         '<span class="t4-mcard-name">' + esc(m.name) + '</span>' +         (m.tag ? '<span class="t4-mcard-tag">' + m.tag + '</span>' : '') +       '</div>' +       '<div class="t4-mcard-desc">' + esc(m.desc) + '</div>' +     '</label>'   ).join('');
}

function t4SelMethod(id) {
  _t4Method = id;
  document.querySelectorAll('.t4-mcard').forEach(el => el.classList.remove('t4-mcard-sel'));   const inp = document.querySelector('input[name="t4meth"]:checked');   if (inp) inp.closest('.t4-mcard').classList.add('t4-mcard-sel');
  /* Refresh preview rows */
  _t4RefreshPreview();
}

function _t4RefreshPreview() {
  /* Re-render just the preview table body with updated fill values */
  const src    = App.filtered || App.raw;
  const qCols  = App.config.quantCols;
  const showCols = qCols.slice(0, 5);
  const previewTop = _t4MV.perProtein
    .filter(p => p.nMiss > 0)
    .sort((a, b) => b.missPct - a.missPct)
    .slice(0, 8);

  const tbody = document.querySelector('.t4-preview-tbl tbody');
  if (!tbody || !previewTop.length) return;

  const rows = previewTop.map(p => {
    const row  = src.rows[p.rowIdx];
    const cols = showCols.map(c => {
      const v = String(row[c] !== undefined ? row[c] : '').trim();
      const miss = !v || MISSING_RE.test(v);
      return { miss, num: miss ? null : parseFloat(v) };
    });
    const allNums = qCols.map(c => {
      const v = String(row[c] !== undefined ? row[c] : '').trim();
      return (!v || MISSING_RE.test(v)) ? null : parseFloat(v);
    }).filter(v => v !== null && isFinite(v));
    const fv = (_t4Method !== 'knn') ? _t4FillVal(_t4Method, allNums) : null;

    const cells = cols.map(cell => {
      if (!cell.miss) return '<td class="t4-val-ok">' + (cell.num !== null ? cell.num.toPrecision(4) : '') + '</td>';       const after = fv !== null ? '<span class="t4-imputed">' + fv.toPrecision(3) + '</span>' : '<span class="t4-imputed t4-knn-est">~KNN</span>';       return '<td class="t4-val-miss"><span class="t4-miss-dash">—</span> ' + after + '</td>';     }).join('');

    const extraCols = qCols.length - showCols.length;
    return '<tr><td class="t4-pid">' + esc(p.pid) + '</td>' +       '<td class="num t4-miss-pct">' + p.missPct.toFixed(0) + '%</td>' +       cells + (extraCols > 0 ? '<td class="t4-dim">…</td>' : '') + '</tr>';
  });
  tbody.innerHTML = rows.join('');
}

/* ══════════════════════════════════════════════════════════════
   CHARTS
══════════════════════════════════════════════════════════════ */
function _t4BuildHist() {
  const mv  = _t4MV;
  const el  = document.getElementById('t4-hist');   if (!el || typeof Chart === 'undefined') return;

  const vals = mv.perProtein.map(p => p.missPct);
  const nBins = Math.max(10, Math.min(20, Math.ceil(Math.sqrt(vals.length))));
  const bsz   = 100 / nBins;
  const bins  = Array(nBins).fill(0);
  vals.forEach(v => bins[Math.min(Math.floor(v / bsz), nBins - 1)]++);
  const labels = Array.from({ length: nBins }, (_, i) => (i * bsz).toFixed(0) + '%');

  const chart = new Chart(el.getContext('2d'), {     type: 'bar',
    data: {
      labels,
      datasets: [{
        data: bins,
        backgroundColor: labels.map((_, i) => i * bsz > 50 ? 'rgba(220,38,38,0.62)' : 'rgba(37,99,235,0.52)'),
        borderWidth: 0,
        barPercentage: 1.0,
        categoryPercentage: 1.0,
      }],
    },
    options: {
      animation: false, responsive: true, maintainAspectRatio: false,
      plugins: { legend: { display: false },
        tooltip: { callbacks: {
          title: items => 'Missing: ' + items[0].label,           label: item  => 'Proteins: ' + item.raw,
        }},
      },
      scales: {
        x: { ticks: { maxTicksLimit: 6, font: { size: 11 } }, grid: { display: false } },
        y: { beginAtZero: true, ticks: { font: { size: 11 } },
          title: { display: true, text: 'Proteins', font: { size: 11 }, color: '#64748b' } },
      },
    },
  });
  _t4Charts['t4-hist'] = chart;
}

function _t4BuildLollipop() {
  const mv = _t4MV;
  const el = document.getElementById('t4-lollipop');   if (!el || typeof Chart === 'undefined') return;

  const top = mv.perProtein
    .filter(p => p.missPct > 0)
    .sort((a, b) => b.missPct - a.missPct)
    .slice(0, 15);
  if (!top.length) return;

  const chart = new Chart(el.getContext('2d'), {     type: 'bar',
    data: {
      labels: top.map(p => p.pid),
      datasets: [{
        data:            top.map(p => p.missPct),
        backgroundColor: top.map(p => p.missPct > 50 ? 'rgba(220,38,38,0.7)' : 'rgba(124,58,237,0.6)'),
        borderWidth: 0,
        barThickness: 6,
        borderRadius: 3,
      }],
    },
    options: {
      indexAxis: 'y', animation: false, responsive: true, maintainAspectRatio: false,
      plugins: { legend: { display: false },
        tooltip: { callbacks: { label: item => item.raw.toFixed(1) + '% missing' } },
      },
      scales: {
        x: { min: 0, max: 100, ticks: { font: { size: 10 }, callback: v => v + '%' },           title: { display: true, text: 'Missing (%)', font: { size: 11 }, color: '#64748b' } },
        y: { ticks: { font: { size: 10 } } },
      },
    },
  });
  _t4Charts['t4-lollipop'] = chart;
}

function _t4BuildHeatmap() {
  const el = document.getElementById('t4-heatmap');
  if (!el) return;

  const src    = App.filtered || App.raw;
  const qCols  = App.config.quantCols;
  const idCol  = App.config.idCol;

  const maxColsShown = Math.min(qCols.length, 30);
  const showCols = qCols.slice(0, maxColsShown);

  const top = _t4MV.perProtein
    .slice().sort((a, b) => b.missPct - a.missPct)
    .slice(0, 40);

  const nR = top.length;
  const nC = showCols.length;

  const LABEL_W  = 78;
  const HEADER_H = 52;
  const CELL_W   = Math.max(10, Math.min(20, Math.floor(700 / nC)));
  const CELL_H   = 14;
  const W        = LABEL_W + nC * CELL_W;
  const H        = HEADER_H + nR * CELL_H;

  el.width  = W;
  el.height = H;
  el.style.display = 'block';

  const ctx = el.getContext('2d');
  ctx.clearRect(0, 0, W, H);

  /* Column headers (rotated) */
  ctx.fillStyle = '#64748b';   ctx.font = '9px system-ui,sans-serif';
  showCols.forEach((c, ci) => {
    const x = LABEL_W + ci * CELL_W + CELL_W / 2;
    ctx.save();
    ctx.translate(x, HEADER_H - 4);
    ctx.rotate(-Math.PI / 2.5);
    ctx.textAlign = 'left';     ctx.fillText(c.length > 14 ? c.slice(0, 14) + '…' : c, 0, 0);
    ctx.restore();
  });

  /* Data rows */
  top.forEach((p, ri) => {
    const row = src.rows[p.rowIdx];
    if (!row) return;

    const nums = showCols.map(c => {
      const v = String(row[c] !== undefined ? row[c] : '').trim();
      return (!v || MISSING_RE.test(v)) ? null : parseFloat(v);
    });
    const validNums = nums.filter(v => v !== null && isFinite(v));
    const rMin   = validNums.length ? Math.min(...validNums) : 0;
    const rRange = (validNums.length ? Math.max(...validNums) : 1) - rMin || 1;

    /* Protein label */
    ctx.fillStyle = '#475569';     ctx.font = '10px system-ui,sans-serif';     ctx.textAlign = 'right';     ctx.textBaseline = 'middle';     const lbl = p.pid.length > 10 ? p.pid.slice(0, 10) + '…' : p.pid;
    ctx.fillText(lbl, LABEL_W - 4, HEADER_H + ri * CELL_H + CELL_H / 2);

    nums.forEach((v, ci) => {
      const x = LABEL_W + ci * CELL_W;
      const y = HEADER_H + ri * CELL_H;
      if (v === null) {
        ctx.fillStyle = '#ef4444';
      } else {
        const t = Math.max(0, Math.min(1, (v - rMin) / rRange));
        /* gradient: #bfdbfe (low) → #1e3a8a (high) */
        const r = Math.round(191 - t * 161);
        const g = Math.round(219 - t * 161);
        const b = Math.round(254 - t * 116);
        ctx.fillStyle = 'rgb(' + r + ',' + g + ',' + b + ')';
      }
      ctx.fillRect(x + 1, y + 1, CELL_W - 2, CELL_H - 2);
    });
  });
}

/* ══════════════════════════════════════════════════════════════
   APPLY / PROCEED
══════════════════════════════════════════════════════════════ */
function t4Apply() {
  const btn = document.getElementById('t4-apply-btn');   if (btn) { btn.disabled = true; btn.textContent = '⏳ Imputing…'; }

  const isKNN = _t4Method === 'knn';
  setTimeout(() => {
    const result = _t4Impute(_t4Method);
    result.method = _t4Method;   /* persist for Methods report */
    App.imputed     = result;
    App.normalized  = null;
    if (typeof _t5Cache !== 'undefined') _t5Cache = null;     if (typeof _t5State !== 'undefined') _t5State = null;

    const methodLabels = {
      lod:     'Limit of Detection (LoD) — 1/5 of minimum positive value per protein',       halfmin: 'Half Minimum — half of minimum observed value per protein',       mean:    'Mean Imputation — row mean of observed values',       knn:     'K-Nearest Neighbors (KNN, k=5)',
    };

    const el = document.getElementById('t4-report');
    if (el) {
      el.style.display = '';
      el.innerHTML =
        '<div class="card-title">Imputation Report</div>' +         '<div class="t4-rpt">' +           '<div class="t4-rpt-row">' +             '<span class="t4-rpt-lbl">Method:</span>' +             '<span class="t4-rpt-val">' + esc(methodLabels[_t4Method]) + '</span>' +           '</div>' +           '<div class="t4-rpt-row">' +             '<span class="t4-rpt-lbl">Values imputed:</span>' +             '<span class="t4-rpt-val">' + result.nImputed.toLocaleString() + '</span>' +           '</div>' +           '<div class="t4-rpt-row">' +             '<span class="t4-rpt-lbl">Proteins affected:</span>' +             '<span class="t4-rpt-val">' + result.proteinsAffected.toLocaleString() + '</span>' +           '</div>' +           '<div class="t4-rpt-row">' +             '<span class="t4-rpt-lbl">Dataset dimensions:</span>' +             '<span class="t4-rpt-val">' + result.nRows.toLocaleString() + ' proteins &times; ' + result.quantCols.length + ' samples (unchanged)</span>' +           '</div>' +           '<div class="notice notice-ok" style="margin-top:16px">' +             '&#x2714; Imputation complete. The processed dataset will be passed to the Normalization step.' +           '</div>' +         '</div>';
    }
    if (btn) { btn.disabled = false; btn.innerHTML = '&#x2713; Imputation Applied'; }     showToast('Imputed ' + result.nImputed.toLocaleString() + ' values across ' +               result.proteinsAffected.toLocaleString() + ' proteins.');
  }, isKNN ? 80 : 10);
}

function proceedTab4() {
  if (_t4MV && _t4MV.totalMiss > 0 && !App.imputed) {
    /* auto-apply then navigate */
    t4Apply();
    setTimeout(() => { unlockTab(5); goTab(5); }, _t4Method === 'knn' ? 200 : 50);
    return;
  }
  /* No missing values — pass through original data */
  if (!App.imputed) {
    const src = App.filtered || App.raw;
    App.imputed = { rows: src.rows.map(r => Object.assign({}, r)),
      headers: src.headers, quantCols: App.config.quantCols,
      nRows: src.rows.length, nImputed: 0, proteinsAffected: 0,
      method: 'none' };
  }
  unlockTab(5);
  goTab(5);
}

/* ====== _v2_tab5.js ====== */
/* ==============================================================
   TAB 5  —  Normalization
   Overrides stub renderTab5() from _v2_stubs.js.
   Order: sample normalization → data transformation → data scaling
   ============================================================== */

let _t5State   = null;   /* {sampleNorm, transform, scaling} */
let _t5Charts  = {};     /* canvasId → Chart instance */
let _t5Cache   = null;   /* {rawMat, normMat, pseudocount, qCols, srcRows, headers} */
let _t5Timer   = null;   /* debounce handle for live-preview updates */

function _t5DefaultState() {
  return { sampleNorm: 'median', transform: 'log2', scaling: 'none' };
}

/* ══════════════════════════════════════════════════════════════
   MATRIX HELPERS
══════════════════════════════════════════════════════════════ */
function _t5GetMat(rows, qCols) {
  return rows.map(row => qCols.map(c => {
    const v = String(row[c] !== undefined ? row[c] : '').trim();
    return (!v || MISSING_RE.test(v)) ? null : parseFloat(v);
  }));
}

function _t5ColVals(mat, j) {
  return mat.map(r => r[j]).filter(v => v !== null && isFinite(v));
}

function _t5RowVals(mat, i) {
  return mat[i].filter(v => v !== null && isFinite(v));
}

function _t5Clone(mat) { return mat.map(row => row.slice()); }

function _t5MatToRows(normMat, srcRows, qCols) {
  return srcRows.map((row, i) => {
    const nr = Object.assign({}, row);
    qCols.forEach((c, j) => { if (normMat[i][j] !== null) nr[c] = normMat[i][j]; });
    return nr;
  });
}

/* ══════════════════════════════════════════════════════════════
   NORMALIZATION PIPELINE
══════════════════════════════════════════════════════════════ */
function _t5RunPipeline(rawMat, state) {
  let mat = _t5Clone(rawMat);
  const nRows = mat.length;
  const nCols = mat[0] ? mat[0].length : 0;
  let pseudocount = null;

  /* ── 1. Sample normalization (column-wise) ── */
  if (state.sampleNorm === 'median') {
    const meds = Array.from({ length: nCols }, (_, j) => {
      const v = _t5ColVals(mat, j);
      return v.length ? _med(_srt(v)) : null;
    });
    const refMed = _med(_srt(meds.filter(m => m !== null && m > 0)));
    if (isFinite(refMed)) {
      for (let j = 0; j < nCols; j++) {
        if (meds[j] && meds[j] !== 0) {
          const scale = refMed / meds[j];
          for (let i = 0; i < nRows; i++) {
            if (mat[i][j] !== null) mat[i][j] *= scale;
          }
        }
      }
    }
  }

  else if (state.sampleNorm === 'sum') {
    const sums = Array.from({ length: nCols }, (_, j) =>
      _t5ColVals(mat, j).reduce((a, x) => a + x, 0)
    );
    const posSums = sums.filter(s => s > 0);
    if (posSums.length) {
      const refSum = posSums.reduce((a, s) => a + s, 0) / posSums.length;
      for (let j = 0; j < nCols; j++) {
        if (sums[j] > 0) {
          const scale = refSum / sums[j];
          for (let i = 0; i < nRows; i++) {
            if (mat[i][j] !== null) mat[i][j] *= scale;
          }
        }
      }
    }
  }

  else if (state.sampleNorm === 'quantile') {
    /* Sort each column, compute rank means, assign back */
    const sorted = Array.from({ length: nCols }, (_, j) =>
      mat.map((r, i) => ({ v: r[j], i }))
         .filter(x => x.v !== null && isFinite(x.v))
         .sort((a, b) => a.v - b.v)
    );
    const maxN = Math.max(...sorted.map(c => c.length), 0);
    const rankMeans = Array.from({ length: maxN }, (_, rank) => {
      const vals = sorted.map(col => col[rank] ? col[rank].v : null).filter(v => v !== null);
      return vals.length ? vals.reduce((a, v) => a + v, 0) / vals.length : null;
    });
    for (let j = 0; j < nCols; j++) {
      sorted[j].forEach((entry, rank) => {
        if (rankMeans[rank] !== null) mat[entry.i][j] = rankMeans[rank];
      });
    }
  }

  /* ── 2. Data transformation (element-wise) ── */
  if (state.transform === 'log2' || state.transform === 'log10') {
    const allVals = mat.flat().filter(v => v !== null && isFinite(v));
    const minVal  = allVals.length ? Math.min(...allVals) : 1;
    if (minVal <= 0) {
      const pos = allVals.filter(v => v > 0);
      pseudocount = pos.length ? Math.min(...pos) / 2 : 1;
    }
    const pc   = pseudocount || 0;
    const logF = state.transform === 'log2'
      ? (x => Math.log2(x + pc))
      : (x => Math.log10(x + pc));
    for (let i = 0; i < nRows; i++) {
      for (let j = 0; j < nCols; j++) {
        if (mat[i][j] !== null) {
          const v = mat[i][j] + pc;
          mat[i][j] = v > 0 ? logF(mat[i][j]) : null;
        }
      }
    }
  }

  else if (state.transform === 'sqrt') {
    for (let i = 0; i < nRows; i++) {
      for (let j = 0; j < nCols; j++) {
        if (mat[i][j] !== null) {
          mat[i][j] = mat[i][j] >= 0 ? Math.sqrt(mat[i][j]) : null;
        }
      }
    }
  }

  /* ── 3. Data scaling (row-wise, per protein) ── */
  if (state.scaling !== 'none') {
    for (let i = 0; i < nRows; i++) {
      const vals = _t5RowVals(mat, i);
      if (vals.length < 2) continue;
      const mean = _mn(vals);
      const std  = _sd(vals);
      for (let j = 0; j < nCols; j++) {
        if (mat[i][j] === null) continue;
        if (state.scaling === 'center') {
          mat[i][j] -= mean;
        } else if (state.scaling === 'auto') {
          mat[i][j] = std > 0 ? (mat[i][j] - mean) / std : 0;
        } else if (state.scaling === 'pareto') {
          mat[i][j] = std > 0 ? (mat[i][j] - mean) / Math.sqrt(std) : 0;
        }
      }
    }
  }

  return { mat, pseudocount };
}

/* ══════════════════════════════════════════════════════════════
   RENDER
══════════════════════════════════════════════════════════════ */
function renderTab5() {
  document.getElementById('main-content').innerHTML =     '<div class="tab-content">' +       '<div class="tab-hd">' +         '<h2 class="tab-title">Normalization</h2>' +         '<p class="tab-desc">Adjust for systematic technical variation, compress dynamic range, ' +         'and center features. The preview updates as you change settings. ' +         'Only quantitative columns are modified &mdash; protein IDs and annotation columns are preserved.</p>' +       '</div>' +       '<div id="t5-body"><div class="notice notice-info">&#x23F3; Preparing normalization&hellip;</div></div>' +     '</div>';

  setTimeout(() => {
    if (!_t5State) _t5State = _t5DefaultState();
    _t5Recompute();
    _renderTab5Body();
  }, 30);
}

function _t5Recompute() {
  const src    = App.imputed || App.filtered || App.raw;
  const qCols  = App.config.quantCols;
  const rawMat = _t5GetMat(src.rows, qCols);
  const { mat, pseudocount } = _t5RunPipeline(rawMat, _t5State);
  _t5Cache = { rawMat, normMat: mat, pseudocount, qCols, srcRows: src.rows, headers: src.headers };
}

/* ── Triggered whenever any radio changes ── */
function _t5OnChange() {
  /* Show a subtle "recalculating" state instantly */
  const viz = document.getElementById('t5-viz-area');   if (viz) viz.style.opacity = '0.5';
  /* Debounce: 180 ms — avoids heavy recompute on every intermediate click */
  clearTimeout(_t5Timer);
  _t5Timer = setTimeout(() => {
    _t5Recompute();
    /* update charts */
    Object.values(_t5Charts).forEach(c => { try { c.destroy(); } catch(e){} });
    _t5Charts = {};
    if (viz) { viz.style.opacity = ''; viz.innerHTML = _t5VizHTML(); requestAnimationFrame(() => requestAnimationFrame(_t5BuildCharts)); }
    /* update table */
    const tb = document.getElementById('t5-sumtbl-body');
    if (tb) _t5FillTable(tb);
    /* update pseudocount notice */
    const pn = document.getElementById('t5-pseudo-note');
    if (pn) {
      const pc = _t5Cache.pseudocount;
      pn.style.display = pc !== null && pc !== undefined ? '' : 'none';
      if (pc !== null && pc !== undefined)
        pn.innerHTML = '&#x26A0; Pseudocount added before log transformation: <strong>' + pc.toExponential(3) + '</strong>';
    }
  }, 180);
}

/* ── Radio handlers ── */
function t5SetSN(v) { _t5State.sampleNorm = v; _t5OnChange(); }
function t5SetTR(v) { _t5State.transform  = v; _t5OnChange(); }
function t5SetSC(v) { _t5State.scaling    = v; _t5OnChange(); }

function _renderTab5Body() {
  Object.values(_t5Charts).forEach(c => { try { c.destroy(); } catch(e){} });
  _t5Charts = {};

  const st  = _t5State;
  const src = App.imputed || App.filtered || App.raw;
  const N   = src.rows.length;
  const nS  = App.config.quantCols.length;
  const pc  = _t5Cache.pseudocount;

  function optRows(radioName, opts, cur, fn) {
    return opts.map(([v, l, tag]) =>
      '<label class="t5-opt-row">' +         '<input type="radio" name="' + radioName + '" value="' + v + '"' +           (cur === v ? ' checked' : '') + ' onchange="' + fn + '(\'' + v + '\')">' +         '<span class="t5-opt-txt">' + esc(l) + (tag ? ' <span class="t5-tag">' + tag + '</span>' : '') + '</span>' +       '</label>'     ).join('');
  }

  document.getElementById('t5-body').innerHTML =

    /* ── Three section cards ── */
    '<div class="t5-3col">' +

      '<div class="card t5-scard">' +         '<div class="card-title">Sample Normalization</div>' +         '<p class="t5-sdesc">Correct for differences in sample loading or total intensity.</p>' +         '<div class="t5-opts">' +           optRows('t5sn', [             ['none',     'None',                   null],             ['median',   'Median normalization',   'Default'],             ['sum',      'Total intensity',        null],             ['quantile', 'Quantile normalization', null],           ], st.sampleNorm, 't5SetSN') +         '</div>' +       '</div>' +

      '<div class="card t5-scard">' +         '<div class="card-title">Data Transformation</div>' +         '<p class="t5-sdesc">Compress dynamic range and approximate normality.</p>' +         '<div class="t5-opts">' +           optRows('t5tr', [             ['none',  'None',                 null],             ['log2',  'Log₂ transformation', 'Default'],             ['log10', 'Log₁₀ transformation', null],             ['sqrt',  'Square root',          null],           ], st.transform, 't5SetTR') +         '</div>' +         '<div class="t5-pseudo-note" id="t5-pseudo-note"' + (pc !== null && pc !== undefined ? '' : ' style="display:none"') + '>' +           '&#x26A0; Pseudocount: ' + (pc !== null && pc !== undefined ? pc.toExponential(3) : '') +         '</div>' +       '</div>' +

      '<div class="card t5-scard">' +         '<div class="card-title">Data Scaling</div>' +         '<p class="t5-sdesc">Center and scale features (applied per protein, row-wise).</p>' +         '<div class="t5-opts">' +           optRows('t5sc', [             ['none',   'None',                   'Default'],             ['center', 'Mean centering',         null],             ['auto',   'Auto / Z-score scaling', null],             ['pareto', 'Pareto scaling',         null],           ], st.scaling, 't5SetSC') +         '</div>' +       '</div>' +

    '</div>' +

    /* ── Visualizations ── */
    '<div class="card">' +       '<div class="card-title">Before &amp; After Comparison</div>' +       '<div id="t5-viz-area">' + _t5VizHTML() + '</div>' +     '</div>' +

    /* ── Per-sample summary table ── */
    '<div class="card">' +       '<div class="card-title">Per-sample Summary <span class="card-meta">' + N.toLocaleString() + ' proteins &times; ' + nS + ' samples</span></div>' +       '<div class="t5-sumtbl">' +         '<table id="t5-sumtbl"><thead>' +           '<tr>' +             '<th rowspan="2" style="text-align:left">Sample</th>' +             '<th colspan="4" class="t5-ghd t5-ghd-b">Before</th>' +             '<th colspan="4" class="t5-ghd t5-ghd-a">After</th>' +           '</tr>' +           '<tr>' +             '<th>Median</th><th>Mean</th><th>CV%</th><th>Missing</th>' +             '<th>Median</th><th>Mean</th><th>CV%</th><th>Missing</th>' +           '</tr>' +         '</thead><tbody id="t5-sumtbl-body"></tbody></table>' +       '</div>' +     '</div>' +

    /* ── Report (after Apply) ── */
    '<div class="card" id="t5-report" style="display:none"></div>' +

    /* ── Actions ── */
    '<div class="card">' +       '<div class="btn-row">' +         '<button class="btn btn-primary btn-lg" onclick="t5Apply()">&#x2713; Apply Normalization</button>' +         '<button class="btn btn-primary btn-lg" onclick="proceedTab5()" style="background:#7c3aed;border-color:#7c3aed">Proceed to Surfaceome Filter &#8594;</button>' +         '<button class="btn btn-ghost" onclick="t5Reset()">&#x21BA; Reset to Defaults</button>' +         '<button class="btn btn-ghost" onclick="goTab(4)">&#8592; Back</button>' +       '</div>' +     '</div>';

  requestAnimationFrame(() => requestAnimationFrame(() => {
    _t5BuildCharts();
    const tb = document.getElementById('t5-sumtbl-body');
    if (tb) _t5FillTable(tb);
  }));
}

/* ══════════════════════════════════════════════════════════════
   VIZ HTML SKELETON
══════════════════════════════════════════════════════════════ */
function _t5VizHTML() {
  const nS = App.config.quantCols.length;
  const shown = Math.min(nS, 20);
  const note  = nS > 20 ? ' <span class="t5-trunc-note">(first ' + shown + ' of ' + nS + ' shown)</span>' : '';
  return (
    '<div class="t5-bp-grid">' +       '<div>' +         '<div class="sec-title">' +           '<span class="t5-bdg t5-bdg-b">Before</span>' +           ' Sample intensity distributions' + note +         '</div>' +         '<div class="t5-bp-wrap"><canvas id="t5-bp-before"></canvas></div>' +       '</div>' +       '<div>' +         '<div class="sec-title">' +           '<span class="t5-bdg t5-bdg-a">After</span>' +           ' Sample intensity distributions' + note +         '</div>' +         '<div class="t5-bp-wrap"><canvas id="t5-bp-after"></canvas></div>' +       '</div>' +     '</div>' +     '<div class="t5-bp-grid" style="margin-top:20px">' +       '<div>' +         '<div class="sec-title"><span class="t5-bdg t5-bdg-b">Before</span> Value distribution</div>' +         '<div style="position:relative;height:175px"><canvas id="t5-hist-before"></canvas></div>' +       '</div>' +       '<div>' +         '<div class="sec-title"><span class="t5-bdg t5-bdg-a">After</span> Value distribution</div>' +         '<div style="position:relative;height:175px"><canvas id="t5-hist-after"></canvas></div>' +       '</div>' +     '</div>'
  );
}

/* ══════════════════════════════════════════════════════════════
   CANVAS BOXPLOT DRAWING
══════════════════════════════════════════════════════════════ */
function _t5ColStats(mat, j) {
  const vals = _t5ColVals(mat, j).sort((a, b) => a - b);
  if (!vals.length) return null;
  const n   = vals.length;
  const med = _med(vals);
  const q1  = _med(vals.slice(0, Math.floor(n / 2)));
  const q3  = _med(vals.slice(Math.ceil(n / 2)));
  const iqr = q3 - q1;
  return {
    min: vals[0], max: vals[n - 1],
    wLo: Math.max(vals[0], q1 - 1.5 * iqr),
    wHi: Math.min(vals[n - 1], q3 + 1.5 * iqr),
    q1, med, q3,
    mean: _mn(vals),
    cv:   _mn(vals) !== 0 ? Math.abs(_sd(vals) / _mn(vals)) * 100 : 0,
    nMiss: mat.length - n,
  };
}

function _t5DrawBoxplot(canvasId, mat, labels, boxColor, lineColor) {
  const el = document.getElementById(canvasId);
  if (!el || !mat.length || !mat[0]) return;

  const SHOW = Math.min(mat[0].length, 20);
  const stats = Array.from({ length: SHOW }, (_, j) => _t5ColStats(mat, j));
  const valid = stats.filter(Boolean);
  if (!valid.length) return;

  const allW = valid.flatMap(s => [s.wLo, s.wHi]).filter(isFinite);
  if (!allW.length) return;

  const yMin = Math.min(...allW);
  const yMax = Math.max(...allW);
  const yRng = (yMax - yMin) || 1;

  const ML = 52, MR = 10, MT = 10, MB = 62;
  const BW = Math.max(12, Math.min(24, Math.floor(500 / SHOW)));
  const GAP = Math.max(6, Math.floor(BW * 0.4));
  const W  = ML + SHOW * (BW + GAP) + MR;
  const H  = 220;
  const pH = H - MT - MB;

  el.width  = W;
  el.height = H;
  el.style.display = 'block';

  const ctx = el.getContext('2d');
  ctx.clearRect(0, 0, W, H);

  /* Y grid + axis */
  ctx.textAlign = 'right';   ctx.font = '9px system-ui,sans-serif';
  for (let t = 0; t <= 4; t++) {
    const yV = yMin + (yRng * t / 4);
    const y  = MT + pH - (yV - yMin) / yRng * pH;
    ctx.strokeStyle = '#e2e8f0';
    ctx.lineWidth = 0.5;
    ctx.beginPath(); ctx.moveTo(ML, y); ctx.lineTo(W - MR, y); ctx.stroke();
    ctx.fillStyle = '#94a3b8';
    ctx.fillText(yV.toFixed(1), ML - 4, y + 4);
  }

  /* Boxes */
  const yS = v => MT + pH - (v - yMin) / yRng * pH;

  stats.forEach((s, j) => {
    if (!s) return;
    const cx = ML + j * (BW + GAP) + BW / 2;
    const x0 = cx - BW / 2;

    /* Whisker stem */
    ctx.strokeStyle = '#94a3b8';
    ctx.lineWidth = 1;
    ctx.beginPath();
    ctx.moveTo(cx, yS(s.wLo)); ctx.lineTo(cx, yS(s.wHi));
    /* Caps */
    const capW = BW / 3;
    ctx.moveTo(cx - capW, yS(s.wLo)); ctx.lineTo(cx + capW, yS(s.wLo));
    ctx.moveTo(cx - capW, yS(s.wHi)); ctx.lineTo(cx + capW, yS(s.wHi));
    ctx.stroke();

    /* IQR box */
    const bTop = yS(s.q3);
    const bBot = yS(s.q1);
    const bH   = Math.max(bBot - bTop, 2);
    ctx.fillStyle = boxColor;
    ctx.fillRect(x0, bTop, BW, bH);
    ctx.strokeStyle = lineColor;
    ctx.lineWidth = 1.5;
    ctx.strokeRect(x0, bTop, BW, bH);

    /* Median line */
    const yM = yS(s.med);
    ctx.strokeStyle = lineColor;
    ctx.lineWidth = 2.5;
    ctx.beginPath(); ctx.moveTo(x0, yM); ctx.lineTo(x0 + BW, yM); ctx.stroke();

    /* X-axis label */
    ctx.save();
    ctx.translate(cx, H - MB + 8);
    ctx.rotate(-Math.PI / 3.5);
    ctx.textAlign = 'right';     ctx.font = '9px system-ui,sans-serif';     ctx.fillStyle = '#475569';     const lbl = labels[j].length > 14 ? labels[j].slice(0, 14) + '…' : labels[j];
    ctx.fillText(lbl, 0, 0);
    ctx.restore();
  });
}

/* ══════════════════════════════════════════════════════════════
   CHART.JS HISTOGRAMS (before and after — independent x scales)
══════════════════════════════════════════════════════════════ */
function _t5OneHist(canvasId, vals, color, key) {
  const el = document.getElementById(canvasId);
  if (!el || !vals.length || typeof Chart === 'undefined') return;
  const nBins = Math.max(12, Math.min(25, Math.ceil(Math.sqrt(vals.length))));
  const mn = Math.min(...vals), mx = Math.max(...vals);
  const bsz = (mx - mn) / nBins || 1;
  const bins = Array(nBins).fill(0);
  vals.forEach(v => bins[Math.min(Math.floor((v - mn) / bsz), nBins - 1)]++);
  const labels = Array.from({ length: nBins }, (_, i) => (mn + i * bsz).toFixed(2));
  if (_t5Charts[key]) { _t5Charts[key].destroy(); delete _t5Charts[key]; }
  _t5Charts[key] = new Chart(el.getContext('2d'), {     type: 'bar',
    data: { labels, datasets: [{ data: bins, backgroundColor: color, borderWidth: 0, barPercentage: 1.0, categoryPercentage: 1.0 }] },
    options: {
      animation: false, responsive: true, maintainAspectRatio: false,
      plugins: { legend: { display: false }, tooltip: { callbacks: { label: item => item.raw + ' values' } } },
      scales: {
        x: { ticks: { maxTicksLimit: 6, font: { size: 10 } }, grid: { display: false } },
        y: { beginAtZero: true, ticks: { font: { size: 10 } },
          title: { display: true, text: 'Count', font: { size: 11 }, color: '#64748b' } },
      },
    },
  });
}

function _t5BuildHistCmp() {
  const { rawMat, normMat } = _t5Cache;
  _t5OneHist('t5-hist-before',
    rawMat.flat().filter(v => v !== null && isFinite(v)),
    'rgba(148,163,184,0.65)', 't5hb');   _t5OneHist('t5-hist-after',
    normMat.flat().filter(v => v !== null && isFinite(v)),
    'rgba(37,99,235,0.55)', 't5ha');
}

function _t5BuildCharts() {
  const labels = App.config.quantCols;
  _t5DrawBoxplot('t5-bp-before', _t5Cache.rawMat,  labels, 'rgba(148,163,184,0.4)', '#64748b');   _t5DrawBoxplot('t5-bp-after',  _t5Cache.normMat, labels, 'rgba(147,197,253,0.5)', '#2563eb');
  _t5BuildHistCmp();
}

/* ══════════════════════════════════════════════════════════════
   SUMMARY TABLE
══════════════════════════════════════════════════════════════ */
function _t5FillTable(tbody) {
  if (!tbody || !_t5Cache) return;
  const { rawMat, normMat, qCols } = _t5Cache;
  const nCols = (rawMat[0] || []).length;

  function fmtN(v) { return isFinite(v) ? v.toFixed(2) : '—'; }

  let html = '';
  for (let j = 0; j < nCols; j++) {
    const sb = _t5ColStats(rawMat, j);
    const sa = _t5ColStats(normMat, j);
    html += '<tr>' +       '<td style="text-align:left;max-width:140px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap" title="' + esc(qCols[j]) + '">' + esc(qCols[j]) + '</td>' +       '<td>' + (sb ? fmtN(sb.med)  : '—') + '</td>' +       '<td>' + (sb ? fmtN(sb.mean) : '—') + '</td>' +       '<td>' + (sb ? fmtN(sb.cv)   : '—') + '</td>' +       '<td>' + (sb ? sb.nMiss.toLocaleString() : '—') + '</td>' +       '<td>' + (sa ? fmtN(sa.med)  : '—') + '</td>' +       '<td>' + (sa ? fmtN(sa.mean) : '—') + '</td>' +       '<td>' + (sa ? fmtN(sa.cv)   : '—') + '</td>' +       '<td>' + (sa ? sa.nMiss.toLocaleString() : '—') + '</td>' +     '</tr>';
  }
  tbody.innerHTML = html;
}

/* ══════════════════════════════════════════════════════════════
   APPLY / RESET / PROCEED
══════════════════════════════════════════════════════════════ */
function t5Apply() {
  if (!_t5Cache) { showToast('Run normalization first.'); return; }

  const { normMat, pseudocount, srcRows, headers, qCols } = _t5Cache;
  const normalizedRows = _t5MatToRows(normMat, srcRows, qCols);
  const src = App.imputed || App.filtered || App.raw;

  App.normalized = {
    rows:        normalizedRows,
    headers:     headers,
    quantCols:   qCols,
    nRows:       normalizedRows.length,
    state:       Object.assign({}, _t5State),
    pseudocount: pseudocount,
  };

  const SN_L = { none: 'None', median: 'Median normalization',                  sum: 'Sum/total intensity normalization', quantile: 'Quantile normalization' };   const TR_L = { none: 'None', log2: 'Log₂ transformation',                  log10: 'Log₁₀ transformation', sqrt: 'Square root transformation' };   const SC_L = { none: 'None', center: 'Mean centering',                  auto: 'Auto / Z-score scaling', pareto: 'Pareto scaling' };

  const el = document.getElementById('t5-report');
  if (el) {
    el.style.display = '';
    el.innerHTML =
      '<div class="card-title">Normalization Summary</div>' +       '<div class="t5-rpt-grid">' +         '<span class="t5-rpt-k">Sample normalization:</span><span class="t5-rpt-v">' + esc(SN_L[_t5State.sampleNorm]) + '</span>' +         '<span class="t5-rpt-k">Data transformation:</span><span class="t5-rpt-v">' + esc(TR_L[_t5State.transform]) + '</span>' +         '<span class="t5-rpt-k">Data scaling:</span><span class="t5-rpt-v">' + esc(SC_L[_t5State.scaling]) + '</span>' +
        (pseudocount !== null && pseudocount !== undefined
          ? '<span class="t5-rpt-k">Pseudocount added:</span><span class="t5-rpt-v">' + pseudocount.toExponential(4) + '</span>'           : '') +         '<span class="t5-rpt-k">Input matrix:</span><span class="t5-rpt-v">' + src.rows.length.toLocaleString() + ' proteins &times; ' + qCols.length + ' samples</span>' +         '<span class="t5-rpt-k">Output matrix:</span><span class="t5-rpt-v">' + normalizedRows.length.toLocaleString() + ' proteins &times; ' + qCols.length + ' samples (unchanged)</span>' +       '</div>' +       '<div class="notice notice-ok" style="margin-top:14px">' +         '&#x2714; Normalization applied. The processed dataset is ready for Surfaceome Filter.' +       '</div>';
  }
  showToast('Normalization applied to ' + normalizedRows.length.toLocaleString() + ' proteins.');
}

function t5Reset() {
  _t5State = _t5DefaultState();
  App.normalized = null;
  _t5Recompute();
  _renderTab5Body();
  showToast('Normalization settings reset to defaults.');
}

function proceedTab5() {
  if (!App.normalized) t5Apply();
  unlockTab(6);
  goTab(6);
}

/* ====== _v2_tab6.js ====== */
/* ==============================================================
   TAB 6 — Surfaceome Filter & Analysis Dashboard
   Clean scientific design: 4-category summary, Venn, group
   abundance tables, differential analysis, downloads.
   ============================================================== */

/* ═══════════════════════════════════════════════════════════════
   SECTION 1 – Module state
═══════════════════════════════════════════════════════════════ */
let _t6SS = null, _t6PM = null, _t6MA = null;   // annotation Sets
let _t6Results = null;       // full analysis result object
let _t6Charts  = {};         // Chart.js instances
let _t6DrawerPid = null;     // open drawer pid
let _t6ActiveComp = 0;       // active pairwise comparison index
let _t6TblState  = {};       // { tableId: { sortCol, sortDir, query } }

/* ═══════════════════════════════════════════════════════════════
   SECTION 2 – Annotation sets + ID normalisation
═══════════════════════════════════════════════════════════════ */
function _t6Annot() {
  if (_t6SS) return;
  /* Embedded strings are comma-separated UniProt accessions */
  _t6SS = new Set(_SS_RAW.split(',').map(s => s.trim()).filter(Boolean));   _t6PM = new Set(_PM_RAW.split(',').map(s => s.trim()).filter(Boolean));   _t6MA = new Set(_MA_RAW.split(',').map(s => s.trim()).filter(Boolean));
}

/* Handle sp|P12345|GENE_HUMAN, P12345;Q67890, plain P12345 */
const _T6_UP_RE = /^[A-Z][0-9][A-Z0-9]{3}[0-9]$|^[A-Z]{3}[0-9][A-Z0-9]{3}[0-9]$|^A0A[A-Z0-9]{6,10}$/;
function _t6NormId(raw) {
  const parts = String(raw || '').split(/[|;,\s]/);
  for (const p of parts) {
    const t = p.trim();
    if (_T6_UP_RE.test(t)) return t;
  }
  return parts[0].trim();
}

/* ═══════════════════════════════════════════════════════════════
   SECTION 3 – Experimental group detection
═══════════════════════════════════════════════════════════════ */
function _t6DetectGroups(quantCols) {
  if (!quantCols || quantCols.length === 0) return [];
  const groups = {};
  for (const col of quantCols) {
    /* Strip trailing replicate indicator: _1, _Rep2, _Replicate3, .1 etc. */
    const m = col.match(/^(.+?)[\s._-](?:rep(?:licate)?s?|r)?(\d+)$/i);
    const key = m ? m[1].trim() : col;
    if (!groups[key]) groups[key] = [];
    groups[key].push(col);
  }
  const entries = Object.entries(groups);
  /* Only call it "groups" if ≥2 groups each with ≥2 replicates */
  const withReps = entries.filter(([, c]) => c.length >= 2);
  if (entries.length < 2 || withReps.length < 2) {
    return [{ name: 'All Samples', cols: quantCols }];
  }
  return entries.map(([name, cols]) => ({ name, cols }));
}

/* ═══════════════════════════════════════════════════════════════
   SECTION 4 – Statistics
═══════════════════════════════════════════════════════════════ */
function _t6Mean(arr)  { return arr.length ? arr.reduce((a,v)=>a+v,0)/arr.length : 0; }
function _t6Sd(arr) {
  if (arr.length < 2) return NaN;
  const m = _t6Mean(arr), v = arr.reduce((a,x)=>a+(x-m)**2,0)/(arr.length-1);
  return Math.sqrt(v);
}

/* Welch's t-test: returns {t, df, p} */
function _t6Welch(a, b) {
  if (a.length < 2 || b.length < 2) return { t:NaN, df:NaN, p:NaN };
  const [m1,m2] = [_t6Mean(a), _t6Mean(b)];
  const [s1,s2] = [_t6Sd(a), _t6Sd(b)];
  const [v1,v2] = [s1*s1/a.length, s2*s2/b.length];
  if (v1+v2 === 0) return { t:NaN, df:NaN, p:1 };
  const t  = (m1-m2)/Math.sqrt(v1+v2);
  const df = (v1+v2)**2 / (v1**2/(a.length-1) + v2**2/(b.length-1));
  const p  = _t6TwoTailP(t, df);
  return { t, df, p };
}

/* Regularised incomplete beta function for t-distribution p-values */
function _t6LogGamma(n) {
  if (n < 0.5) return Math.log(Math.PI/Math.sin(Math.PI*n)) - _t6LogGamma(1-n);
  n -= 1;
  const x = 0.99999999999980993 + 676.5203681218851/(n+1) - 1259.1392167224028/(n+2)
    + 771.32342877765313/(n+3) - 176.61502916214059/(n+4) + 12.507343278686905/(n+5)
    - 0.13857109526572012/(n+6) + 9.9843695780195716e-6/(n+7) + 1.5056327351493116e-7/(n+8);
  const t = n + 7.5;
  return 0.5*Math.log(2*Math.PI) + (n+0.5)*Math.log(t) - t + Math.log(x);
}
function _t6BetaCF(x, a, b) {
  let f=1, C=1, D=1-(a+b)*x/(a+1);
  if (Math.abs(D)<1e-30) D=1e-30; D=1/D; f=D;
  for (let m=1; m<=200; m++) {
    let d=m*(b-m)*x/((a+2*m-1)*(a+2*m));
    D=1+d*D; if(Math.abs(D)<1e-30) D=1e-30; C=1+d/C; if(Math.abs(C)<1e-30) C=1e-30;
    D=1/D; f*=C*D;
    d=-(a+m)*(a+b+m)*x/((a+2*m)*(a+2*m+1));
    D=1+d*D; if(Math.abs(D)<1e-30) D=1e-30; C=1+d/C; if(Math.abs(C)<1e-30) C=1e-30;
    D=1/D; const delta=C*D; f*=delta;
    if (Math.abs(delta-1)<1e-8) break;
  }
  return f;
}
function _t6RegBeta(x, a, b) {
  if (x<=0) return 0; if (x>=1) return 1;
  const lbeta = _t6LogGamma(a)+_t6LogGamma(b)-_t6LogGamma(a+b);
  const bt = Math.exp(a*Math.log(x)+b*Math.log(1-x)-lbeta);
  if (x < (a+1)/(a+b+2)) return bt * _t6BetaCF(x,a,b) / a;
  return 1 - bt * _t6BetaCF(1-x,b,a) / b;
}
function _t6TwoTailP(t, df) {
  if (!isFinite(df)||df<=0||!isFinite(t)) return 1;
  return _t6RegBeta(df/(df+t*t), df/2, 0.5);
}

/* Benjamini-Hochberg FDR correction */
function _t6BH(pvals) {
  const n = pvals.length;
  const idx = pvals.map((p,i) => [p,i]).sort((a,b)=>a[0]-b[0]);
  const adj = new Array(n);
  let prevAdj = 1;
  for (let k = n-1; k >= 0; k--) {
    const [p, i] = idx[k];
    prevAdj = Math.min(1, Math.min(prevAdj, n*p/(k+1)));
    adj[i] = prevAdj;
  }
  return adj;
}

/* ═══════════════════════════════════════════════════════════════
   SECTION 5 – Core analysis
═══════════════════════════════════════════════════════════════ */
function _t6GetSource() {
  return App.normalized || App.imputed || App.filtered || null;
}

function _t6Analyze() {
  _t6Annot();
  const src = _t6GetSource();
  if (!src) return null;

  const { rows, headers } = src;
  const idCol   = App.config.idCol;
  const geneCol = App.config.geneCol || null;
  const qCols   = App.config.quantCols;

  /* Find a protein name / description column if present */
  const nameCol = headers.find(h => /^(protein.?name|description|gene.?name|gene|full.?name)/i.test(h) && h !== idCol && h !== geneCol) || null;

  /* Detect groups from quantitative column naming */
  const groups = _t6DetectGroups(qCols);
  const hasGroups = groups.length >= 2;

  /* ── Per-protein annotation ── */
  const proteins = rows.map((row, rowIdx) => {
    const rawId = String(row[idCol] || '').trim();
    const pid   = _t6NormId(rawId);
    const gene  = geneCol ? String(row[geneCol] || '').trim() : '';     const name  = nameCol ? String(row[nameCol] || '').trim() : '';

    const inSS = _t6SS.has(pid);
    const inPM = _t6PM.has(pid);
    const inMA = _t6MA.has(pid);
    const dbCount = (inSS?1:0)+(inPM?1:0)+(inMA?1:0);

    /* Category (primary = most specific) */
    let category;
    if      (inSS && inPM && inMA)  category = 'All three databases';     else if (inSS && inPM)          category = 'SURFY + UniProt PM';     else if (inSS && inMA)          category = 'SURFY + Membrane-Assoc.';     else if (inPM && inMA)          category = 'UniProt PM + Membrane-Assoc.';     else if (inSS)                  category = 'SURFY (in-silico)';     else if (inPM)                  category = 'UniProt Cell Membrane';     else if (inMA)                  category = 'Membrane-Associated';     else                            category = 'Non-membrane';

    /* Primary category label (for display) */
    let primaryCat;
    if      (inPM) primaryCat = 'UniProt Cell Membrane';     else if (inSS) primaryCat = 'SURFY (in-silico)';     else if (inMA) primaryCat = 'Membrane-Associated';     else           primaryCat = 'Non-membrane';

    /* Abundance per sample and per group */
    const perSample = {};
    for (const c of qCols) {
      const v = parseFloat(row[c]);
      perSample[c] = isNaN(v) ? null : v;
    }

    const allVals = Object.values(perSample).filter(v => v !== null);
    const meanAbundance = allVals.length ? _t6Mean(allVals) : null;

    const perGroup = {};
    for (const g of groups) {
      const gVals = g.cols.map(c => perSample[c]).filter(v => v !== null);
      perGroup[g.name] = {
        vals: gVals,
        mean: gVals.length ? _t6Mean(gVals) : null,
        sd:   gVals.length >= 2 ? _t6Sd(gVals) : null,
        n:    gVals.length,
        detected: gVals.length > 0,
      };
    }

    return { pid, rawId, gene, name, rowIdx,
             inSS, inPM, inMA, dbCount, category, primaryCat,
             meanAbundance, perSample, perGroup };
  });

  /* ── Venn region membership ── */
  const byRegion = key => {
    const pred = {
      ssOnly:   p => p.inSS && !p.inPM && !p.inMA,
      pmOnly:   p => !p.inSS && p.inPM && !p.inMA,
      maOnly:   p => !p.inSS && !p.inPM && p.inMA,
      ssAndPm:  p => p.inSS && p.inPM && !p.inMA,
      ssAndMa:  p => p.inSS && !p.inPM && p.inMA,
      pmAndMa:  p => !p.inSS && p.inPM && p.inMA,
      allThree: p => p.inSS && p.inPM && p.inMA,
    }[key];
    return proteins.filter(pred).sort((a,b)=>(b.meanAbundance||0)-(a.meanAbundance||0));
  };
  const regions = {
    ssOnly:   byRegion('ssOnly'),     pmOnly:   byRegion('pmOnly'),     maOnly:   byRegion('maOnly'),     ssAndPm:  byRegion('ssAndPm'),     ssAndMa:  byRegion('ssAndMa'),     pmAndMa:  byRegion('pmAndMa'),     allThree: byRegion('allThree'),
  };

  /* ── Category counts ── */
  const surface = proteins.filter(p => p.dbCount > 0)
                          .sort((a,b) => (b.meanAbundance||0)-(a.meanAbundance||0));
  const pmProteins = proteins.filter(p => p.inPM)
                             .sort((a,b) => (b.meanAbundance||0)-(a.meanAbundance||0));

  /* ── Differential analysis (pairwise, UniProt PM only) ── */
  const comparisons = [];
  if (hasGroups) {
    for (let i = 0; i < groups.length-1; i++) {
      for (let j = i+1; j < groups.length; j++) {
        const gA = groups[i], gB = groups[j];
        const diff = pmProteins.map(p => {
          const a = p.perGroup[gA.name], b = p.perGroup[gB.name];
          if (!a || !b || !a.mean || !b.mean) return null;
          const log2fc = Math.log2(b.mean / a.mean);
          const { t, df, p: pval } = _t6Welch(a.vals, b.vals);
          return { pid: p.pid, gene: p.gene, name: p.name, category: p.category,
                   meanA: a.mean, sdA: a.sd, nA: a.n,
                   meanB: b.mean, sdB: b.sd, nB: b.n,
                   log2fc, fc: b.mean/a.mean, pval, fdr: NaN };
        }).filter(Boolean);

        /* BH correction */
        const pvals = diff.map(d => isNaN(d.pval) ? 1 : d.pval);
        const fdrs  = _t6BH(pvals);
        diff.forEach((d, i) => { d.fdr = fdrs[i]; });

        comparisons.push({
          labelA: gA.name, labelB: gB.name,
          label: gA.name + ' vs ' + gB.name,
          diff,
          up:   [...diff].filter(d=>isFinite(d.log2fc)&&d.log2fc>0).sort((a,b)=>b.log2fc-a.log2fc).slice(0,20),
          down: [...diff].filter(d=>isFinite(d.log2fc)&&d.log2fc<0).sort((a,b)=>a.log2fc-b.log2fc).slice(0,20),
        });
      }
    }
  }

  /* ── Condition-specific proteins (UniProt PM, 2-group case) ── */
  let condSpecific = null;
  if (groups.length === 2) {
    const [gA, gB] = groups;
    const detThresh = 0.5; /* fraction of replicates that must be detected */
    condSpecific = pmProteins.map(p => {
      const a = p.perGroup[gA.name], b = p.perGroup[gB.name];
      if (!a || !b) return null;
      const detA = a.n > 0 ? a.n / gA.cols.length : 0;
      const detB = b.n > 0 ? b.n / gB.cols.length : 0;
      let status = null;
      if (detA >= detThresh && detB < detThresh)  status = gA.name + '-specific';       if (detA < detThresh  && detB >= detThresh) status = gB.name + '-specific';
      if (!status) return null;
      return { pid: p.pid, gene: p.gene, name: p.name,
               detA: a.n + '/' + gA.cols.length, detB: b.n + '/' + gB.cols.length,
               meanA: a.mean, meanB: b.mean, status };
    }).filter(Boolean).sort((a,b) => a.status.localeCompare(b.status));
  }

  /* ── Abundance stats for key metrics panel ── */
  const surfaceAbunds = surface.map(p=>p.meanAbundance).filter(v=>v!=null);
  const avgSurfAbund  = surfaceAbunds.length ? _t6Mean(surfaceAbunds) : null;

  return {
    proteins, surface, pmProteins, regions,
    headers, idCol, geneCol, nameCol, qCols,
    groups, hasGroups, comparisons, condSpecific,
    totalN: proteins.length,
    nSS:    proteins.filter(p=>p.inSS).length,
    nPM:    proteins.filter(p=>p.inPM).length,
    nMA:    proteins.filter(p=>p.inMA).length,
    nNone:  proteins.filter(p=>p.dbCount===0).length,
    nSurf:  surface.length,
    avgSurfAbund,
  };
}

/* ═══════════════════════════════════════════════════════════════
   SECTION 6 – Render entry
═══════════════════════════════════════════════════════════════ */
function renderTab6() {
  const mc = document.getElementById('main-content');   mc.innerHTML = '<div class="tab-content"><div class="t6-loading">Analysing surfaceome composition…</div></div>';
  Object.values(_t6Charts).forEach(c => { try { c.destroy(); } catch(e){} });
  _t6Charts = {};
  setTimeout(() => { _t6Results = _t6Analyze(); _renderTab6Body(); }, 0);
}

/* ═══════════════════════════════════════════════════════════════
   SECTION 7 – Main body
═══════════════════════════════════════════════════════════════ */
function _renderTab6Body() {
  const R = _t6Results;
  const mc = document.getElementById('main-content');
  if (!R) {
    mc.innerHTML = '<div class="tab-content"><div class="notice notice-warn">No processed data available. Please complete earlier steps first.</div></div>';
    return;
  }
  const N = R.totalN;
  const pct = v => N > 0 ? (v/N*100).toFixed(1) : '0.0';

  /* Group selector (shown only when ≥3 groups produce multiple comparisons) */
  const compSel = R.comparisons.length > 1
    ? '<div class="t6-comp-sel">' +         '<label class="t6-comp-lbl">Comparison:</label>' +         '<select class="t6-comp-dd" onchange="_t6SetComp(+this.value)">' +           R.comparisons.map((c,i) => '<option value="'+i+'"'+(i===_t6ActiveComp?' selected':'')+'>'+esc(c.label)+'</option>').join('') +         '</select>' +       '</div>'     : '';

  mc.innerHTML =
    '<div class="tab-content">' +

      /* ── Page header ── */
      '<div class="tab-hd">' +         '<h2 class="tab-title">Surfaceome Analysis</h2>' +         '<p class="tab-desc">Proteins were matched against three complementary surfaceome databases: ' +         'SURFY (in-silico predicted surfaceome), UniProt Cell Membrane (KW-1003), and ' +         'a curated membrane-associated protein set. All matching uses UniProt accession numbers.</p>' +       '</div>' +

      /* ── 1. Biological Interpretation Summary ── */
      _t6BioSummaryHTML(R) +

      /* ── 2. Four summary cards ── */
      '<div class="t6-4cards">' +         _t6Card('Non-membrane Proteins', R.nNone, pct(R.nNone), '#64748b') +         _t6Card('UniProt Cell Membrane', R.nPM,   pct(R.nPM),   '#0f766e') +         _t6Card('Membrane-Associated',   R.nMA,   pct(R.nMA),   '#7c3aed') +         _t6Card('SURFY (in-silico)',      R.nSS,   pct(R.nSS),   '#1d4ed8') +       '</div>' +

      /* ── 3. Dataset metrics + category distribution bar chart ── */
      '<div class="t6-charts-top">' +         '<div class="t6-metrics-panel">' +           '<div class="t6-mp-hd">Dataset Overview</div>' +           _t6Metric('Total proteins uploaded',          N.toLocaleString()) +           _t6Metric('Surface-associated proteins',      R.nSurf.toLocaleString()) +           _t6Metric('Non-membrane proteins',            R.nNone.toLocaleString()) +           _t6Metric('Surfaceome coverage',              pct(R.nSurf) + '%') +
          (R.avgSurfAbund != null
            ? _t6Metric('Mean abundance (surface proteins)', R.avgSurfAbund.toExponential(2))             : '') +         '</div>' +         '<div class="card t6-chart-card">' +           '<div class="t6-sec-hd">Surfaceome Composition' +             '<button class="t6-png-btn" onclick="_t6DownloadChartPNG(\'t6-bar-chart\',\'surfaceome_composition.png\')" title="Download as PNG">&#x2193; PNG</button>' +           '</div>' +           '<div class="t6-bar-wrap"><canvas id="t6-bar-chart"></canvas></div>' +         '</div>' +       '</div>' +

      /* ── 4. Per-condition abundance: most abundant surface proteins per group ── */
      (R.hasGroups ? _t6GroupTablesHTML(R) : _t6SingleGroupHTML(R)) +

      /* ── 5. Differential surfaceome analysis (primary section when groups present) ── */
      (R.comparisons.length > 0
        ? '<div class="t6-bio-section-label">Differential Surfaceome Analysis</div>' +
          _t6DiffHTML(R, compSel)
        : '') +

      /* ── 6. Condition-specific surface proteins — biomarker candidates ── */
      (R.condSpecific && R.condSpecific.length > 0 ? _t6CondSpecHTML(R) : '') +

      /* ── 7. Database overlap — Venn diagram ── */
      '<div class="card t6-venn-card">' +         '<div class="t6-sec-hd">Database Overlap' +           '<button class="t6-png-btn" onclick="_t6DownloadVennPNG()" title="Download as PNG">&#x2193; PNG</button>' +         '</div>' +         '<p style="font-size:12px;color:#64748b;margin:-4px 0 14px">Proteins may appear in multiple databases. ' +         'Hover over any region to see the number of proteins and the top 10 by abundance. ' +         'Click a download button to export that region as a CSV file.</p>' +         '<div class="t6-venn-outer">' +           '<div class="t6-venn-wrap" style="position:relative">' +             '<canvas id="t6-venn-canvas" width="640" height="520"></canvas>' +             '<div id="t6-venn-tip" class="t6-venn-tip" style="display:none"></div>' +           '</div>' +           '<div class="t6-venn-dl">' +             '<div class="t6-venn-dl-hd">Download region lists (CSV):</div>' +             '<div class="t6-venn-dl-grid">' +               '<button class="btn btn-ghost t6-dlbtn" onclick="t6Download(\'ssOnly\')">SURFY only</button>' +               '<button class="btn btn-ghost t6-dlbtn" onclick="t6Download(\'pmOnly\')">UniProt PM only</button>' +               '<button class="btn btn-ghost t6-dlbtn" onclick="t6Download(\'maOnly\')">Mem.-Assoc. only</button>' +               '<button class="btn btn-ghost t6-dlbtn" onclick="t6Download(\'ssAndPm\')">SURFY ∩ UniProt PM</button>' +               '<button class="btn btn-ghost t6-dlbtn" onclick="t6Download(\'ssAndMa\')">SURFY ∩ Mem.-Assoc.</button>' +               '<button class="btn btn-ghost t6-dlbtn" onclick="t6Download(\'pmAndMa\')">UniProt PM ∩ Mem.-Assoc.</button>' +               '<button class="btn btn-primary t6-dlbtn" style="background:#0f766e;border-color:#0f766e" onclick="t6Download(\'allThree\')">Supported by all three databases</button>' +             '</div>' +           '</div>' +         '</div>' +       '</div>' +

      /* ── 8. Database membership matrix — evidence per protein ── */
      _t6MembershipMatrixHTML(R) +

      /* ── 9. All surface proteins — full ranked table ── */
      _t6FullTableHTML(R) +

      /* ── 10. Downloads ── */
      '<div class="card">' +         '<div class="t6-sec-hd">Download Results</div>' +         '<p style="font-size:12px;color:#64748b;margin:-4px 0 14px">All exports include original abundance values and surfaceome annotations.</p>' +

        /* Methods report — most prominent export */
        '<div class="t6-report-cta">' +           '<div class="t6-report-cta-body">' +             '<div class="t6-report-cta-title">Comprehensive Analysis Report</div>' +             '<div class="t6-report-cta-desc">12-section scientific report with embedded figures (bar chart, Venn, volcano), top-20 protein tables, all analysis parameters, and database references — opens in a new tab for PDF export or HTML download.</div>' +           '</div>' +           '<button class="btn btn-primary t6-report-btn" onclick="t6DownloadReport()">&#x2197; Open Report (PDF / HTML)</button>' +         '</div>' +

        /* Data exports */
        '<div class="t6-dl-subhd">Export annotated protein lists (CSV):</div>' +         '<div class="t6-dl-row">' +           '<button class="btn btn-ghost t6-dlbtn" onclick="t6Download(\'all\')">All surface proteins</button>' +           '<button class="btn btn-ghost t6-dlbtn" onclick="t6Download(\'pm\')">UniProt Cell Membrane</button>' +           '<button class="btn btn-ghost t6-dlbtn" onclick="t6Download(\'ma\')">Membrane-Associated</button>' +           '<button class="btn btn-ghost t6-dlbtn" onclick="t6Download(\'ss\')">SURFY (in-silico)</button>' +           '<button class="btn btn-ghost t6-dlbtn" onclick="t6Download(\'none\')">Non-membrane proteins</button>' +         '</div>' +       '</div>' +

      /* ── 11. References ── */
      '<div class="t6-refs">' +         '<div class="t6-refs-hd">References</div>' +         '<ol class="t6-refs-list">' +           '<li>Bausch-Fluck D, et al. (2015). A mass spectrometric-derived cell surface protein atlas. <em>PLOS ONE</em> 10(3):e0121314.</li>' +           '<li>Bausch-Fluck D, et al. (2021). The in silico human surfaceome. <em>PNAS</em> 118(1):e2024870118.</li>' +           '<li>UniProt Consortium (2023). UniProt: the Universal Protein Knowledgebase. <em>Nucleic Acids Research</em> 51(D1):D523–D531.</li>' +           '<li>Almén MS, et al. (2009). Mapping the human membrane proteome. <em>BMC Biology</em> 7:50.</li>' +           '<li>Fagerberg L, et al. (2010). Prediction of the human membrane proteome. <em>Proteomics</em> 10(6):1141–1149.</li>' +         '</ol>' +       '</div>' +

      /* Protein detail drawer */
      '<div id="t6-overlay" class="t6-overlay" onclick="t6CloseDrawer()" style="display:none"></div>' +       '<div id="t6-drawer" class="t6-drawer" style="display:none"></div>' +

    '</div>'; /* end tab-content */

  requestAnimationFrame(() => requestAnimationFrame(() => {
    _t6DrawBar(R);
    _t6DrawVenn(R);
  }));
}

/* ── Small helpers ── */
function _t6Card(label, count, pctStr, color) {
  return '<div class="t6-sum-card" style="border-top:3px solid '+color+'">' +     '<div class="t6-sc-count" style="color:'+color+'">' + count.toLocaleString() + '</div>' +     '<div class="t6-sc-pct">' + pctStr + '% of uploaded</div>' +     '<div class="t6-sc-label">' + esc(label) + '</div>' +   '</div>';
}
function _t6Metric(label, val) {
  return '<div class="t6-mp-row"><span class="t6-mp-lbl">'+esc(label)+'</span><span class="t6-mp-val">'+esc(val)+'</span></div>';
}
function _t6FmtAb(v) { return v != null ? v.toExponential(2) : '—'; } function _t6FmtNum(v, dec) { return v != null && isFinite(v) ? v.toFixed(dec) : '—'; }

/* ── Biological interpretation summary ── */
function _t6BioSummaryHTML(R) {
  const N  = R.totalN;
  const pc = v => N > 0 ? (v / N * 100).toFixed(1) : '0.0';

  if (!N) return '';

  /* Largest surface category */
  const cats = [
    { name: 'UniProt Cell Membrane proteins', n: R.nPM },     { name: 'membrane-associated proteins',   n: R.nMA },     { name: 'SURFY in-silico predicted proteins', n: R.nSS },
  ].sort((a, b) => b.n - a.n);

  /* High-confidence = supported by ≥2 databases */
  const highConf = R.regions.ssAndPm.length + R.regions.ssAndMa.length +
                   R.regions.pmAndMa.length + R.regions.allThree.length;

  let s = `Of <strong>${N.toLocaleString()}</strong> detected proteins, ` +
    `<strong>${R.nSurf.toLocaleString()} (${pc(R.nSurf)}%)</strong> were annotated ` +
    `as surface-associated based on three complementary surfaceome databases. `;

  if (R.nSurf > 0) {
    s += `The largest group comprised <strong>${cats[0].n.toLocaleString()} ${cats[0].name}</strong>, ` +
         `with ${cats[1].n.toLocaleString()} ${cats[1].name} and ` +
         `${cats[2].n.toLocaleString()} ${cats[2].name}. `;
    if (highConf > 0) {
      s += `<strong>${highConf.toLocaleString()} protein${highConf !== 1 ? 's were' : ' was'} ` +
           `supported by two or more databases</strong>, representing the highest-confidence ` +
           `surface proteins in this dataset. `;
    }
  } else {
    s += `No proteins matched the surfaceome databases — check that your data contains UniProt accession IDs. `;
  }

  /* Differential */
  if (R.comparisons.length > 0) {
    const comp = R.comparisons[0];
    const sigUp   = comp.diff.filter(d => isFinite(d.log2fc) && d.log2fc >  1 && d.fdr < 0.05).length;
    const sigDown = comp.diff.filter(d => isFinite(d.log2fc) && d.log2fc < -1 && d.fdr < 0.05).length;
    s += `Differential analysis between <strong>${esc(comp.labelA)}</strong> and ` +
         `<strong>${esc(comp.labelB)}</strong> identified ` +
         `<strong>${sigUp} significantly upregulated</strong> and ` +
         `<strong>${sigDown} significantly downregulated</strong> ` +
         `UniProt Cell Membrane proteins (|log₂FC| &gt; 1, FDR &lt; 0.05). `;
  }

  /* Condition-specific */
  if (R.condSpecific && R.condSpecific.length > 0) {
    s += `<strong>${R.condSpecific.length} condition-specific surface protein${R.condSpecific.length !== 1 ? 's were' : ' was'} ` +
         `identified</strong> that may represent candidate biomarkers or therapeutic targets.`;
  }

  return '<div class="t6-bio-summary">' +     '<div class="t6-bio-sum-hd">Biological Interpretation Summary</div>' +     '<div class="t6-bio-sum-body">' + s + '</div>' +   '</div>';
}

/* ═══════════════════════════════════════════════════════════════
   SECTION 8 – Bar chart
═══════════════════════════════════════════════════════════════ */
function _t6DrawBar(R) {
  const ctx = document.getElementById('t6-bar-chart');
  if (!ctx) return;
  if (_t6Charts.bar) { try { _t6Charts.bar.destroy(); } catch(e){} }
  const _barData = [R.nNone, R.nPM, R.nMA, R.nSS];
  /* Inline plugin: draws protein counts at the end of each bar */
  const _barLabelsPlugin = {
    id: 't6-bar-labels',
    afterDatasetsDraw(chart) {
      const { ctx: c2 } = chart;
      const meta = chart.getDatasetMeta(0);
      c2.save();
      c2.font = '600 11px system-ui,sans-serif';       c2.fillStyle = '#475569';       c2.textBaseline = 'middle';       c2.textAlign = 'left';
      meta.data.forEach((bar, i) => {
        const val = _barData[i];
        if (val) c2.fillText(val.toLocaleString(), bar.x + 6, bar.y);
      });
      c2.restore();
    }
  };
  _t6Charts.bar = new Chart(ctx, {
    type: 'bar',
    data: {
      labels: ['Non-membrane', 'UniProt Cell Membrane', 'Membrane-Associated', 'SURFY (in-silico)'],
      datasets: [{ data: _barData,
                   backgroundColor: ['#94a3b8','#0f766e','#7c3aed','#1d4ed8'],
                   borderRadius: 4, borderSkipped: false }]
    },
    plugins: [_barLabelsPlugin],
    options: {
      indexAxis: 'y', responsive: true, maintainAspectRatio: false, animation: false,
      plugins: { legend: { display: false },
                 tooltip: { callbacks: { label: c => ' ' + c.raw.toLocaleString() + ' proteins' } } },
      scales: {
        x: {
          title: { display: true, text: 'Number of Proteins',                    font: { size: 11, weight: '400', family: 'system-ui,sans-serif' }, color: '#64748b' },           grid: { color: '#f1f5f9' },
          ticks: { font: { size: 11 } }
        },
        y: { grid: { display: false }, ticks: { font: { size: 11 }, padding: 4 } }
      }
    }
  });
}

/* ═══════════════════════════════════════════════════════════════
   SECTION 9 – Venn diagram (canvas, 3-set)
═══════════════════════════════════════════════════════════════ */
function _t6DrawVenn(R) {
  const canvas = document.getElementById('t6-venn-canvas');   const tip    = document.getElementById('t6-venn-tip');
  if (!canvas || !R) return;
  const ctx2  = canvas.getContext('2d');
  const W = canvas.width, H = canvas.height;

  /* Circle centres and radius — scaled for 640×520 canvas.
     Circle C bottom = 315+140 = 455, fits within 520px height. */
  const r = 140;
  const A = { x: 210, y: 185, color: 'rgba(29,78,216,0.15)',  stroke: '#1d4ed8', lw: 2.2 };   const B = { x: 430, y: 185, color: 'rgba(15,118,110,0.15)', stroke: '#0f766e', lw: 2.2 };   const C = { x: 320, y: 315, color: 'rgba(124,58,237,0.15)', stroke: '#7c3aed', lw: 2.2 };
  const circles = [A, B, C];

  function inC(px, py, cx, cy) { const dx=px-cx,dy=py-cy; return dx*dx+dy*dy<=r*r; }

  function draw() {
    ctx2.clearRect(0,0,W,H);
    for (const c of circles) {
      ctx2.beginPath();
      ctx2.arc(c.x, c.y, r, 0, Math.PI*2);
      ctx2.fillStyle   = c.color;
      ctx2.fill();
      ctx2.strokeStyle = c.stroke;
      ctx2.lineWidth   = c.lw;
      ctx2.stroke();
    }
    ctx2.textAlign = 'center'; ctx2.textBaseline = 'middle';
    const rg = R.regions;
    function lbl(x, y, n) {
      if (!n) return;
      ctx2.font = 'bold 15px system-ui,sans-serif';       ctx2.fillStyle = '#0f172a';
      ctx2.fillText(n.toLocaleString(), x, y);
    }
    lbl(128, 178,  rg.ssOnly.length);
    lbl(512, 178,  rg.pmOnly.length);
    lbl(320, 404,  rg.maOnly.length);
    lbl(320, 143,  rg.ssAndPm.length);
    lbl(224, 283,  rg.ssAndMa.length);
    lbl(416, 283,  rg.pmAndMa.length);
    lbl(320, 228,  rg.allThree.length);
    /* Bold coloured circle name labels */
    ctx2.font = 'bold 13px system-ui,sans-serif';     ctx2.fillStyle = '#1e3a8a';     ctx2.fillText('SURFY', 110, 68);     ctx2.font = '11px system-ui,sans-serif';     ctx2.fillStyle = '#3b5fc0';     ctx2.fillText('(in-silico)', 110, 84);

    ctx2.font = 'bold 13px system-ui,sans-serif';     ctx2.fillStyle = '#134e4a';     ctx2.fillText('UniProt Cell Membrane', 530, 68);     ctx2.font = '11px system-ui,sans-serif';     ctx2.fillStyle = '#0d7a72';     ctx2.fillText('KW-1003', 530, 84);

    ctx2.font = 'bold 13px system-ui,sans-serif';     ctx2.fillStyle = '#4c1d95';     ctx2.fillText('Membrane-Associated', 320, 476);     ctx2.font = '11px system-ui,sans-serif';     ctx2.fillStyle = '#6d28d9';     ctx2.fillText('(membrane-assoc.)', 320, 493);
  }
  draw();

  /* Region sample points — must satisfy exact circle membership */
  const regionDefs = [
    { key:'ssOnly',   label:'SURFY only',                   cx:128, cy:178 },     { key:'pmOnly',   label:'UniProt PM only',              cx:512, cy:178 },     { key:'maOnly',   label:'Membrane-Assoc. only',         cx:320, cy:400 },     { key:'ssAndPm',  label:'SURFY ∩ UniProt PM',           cx:320, cy:148 },     { key:'ssAndMa',  label:'SURFY ∩ Membrane-Assoc.',      cx:230, cy:283 },     { key:'pmAndMa',  label:'UniProt PM ∩ Membrane-Assoc.', cx:410, cy:283 },     { key:'allThree', label:'All three databases',          cx:320, cy:230 },
  ];

  function regionAt(mx, my) {
    const inA = inC(mx,my,A.x,A.y), inB = inC(mx,my,B.x,B.y), inCC = inC(mx,my,C.x,C.y);
    if (!inA && !inB && !inCC) return null;
    /* Find closest region centre that matches the same membership */
    let best=null, bestD=Infinity;
    for (const rd of regionDefs) {
      const rA=inC(rd.cx,rd.cy,A.x,A.y), rB=inC(rd.cx,rd.cy,B.x,B.y), rCC=inC(rd.cx,rd.cy,C.x,C.y);
      if (rA!==inA || rB!==inB || rCC!==inCC) continue;
      const d=(rd.cx-mx)**2+(rd.cy-my)**2;
      if (d<bestD) { bestD=d; best=rd; }
    }
    return best;
  }

  const wrap = canvas.parentElement;
  canvas.addEventListener('mousemove', e => {
    const rect = canvas.getBoundingClientRect();
    const sx = canvas.width/rect.width, sy = canvas.height/rect.height;
    const mx = (e.clientX-rect.left)*sx, my = (e.clientY-rect.top)*sy;
    const rd = regionAt(mx, my);
    if (!rd) { tip.style.display='none'; return; }
    const prots = R.regions[rd.key] || [];
    if (!prots.length) { tip.style.display='none'; return; }
    const top10 = prots.slice(0,10);
    tip.innerHTML =
      '<div class="t6-tip-hd">'+esc(rd.label)+'</div>' +       '<div class="t6-tip-sub">'+prots.length.toLocaleString()+' proteins · top 10:</div>' +
      top10.map(p =>
        '<div class="t6-tip-row"><span class="t6-tip-pid">'+esc(p.pid)+'</span>'+         (p.gene?'<span class="t6-tip-gene"> '+esc(p.gene)+'</span>':'')+         (p.meanAbundance!=null?'<span class="t6-tip-ab"> '+p.meanAbundance.toExponential(1)+'</span>':'')+         '</div>'       ).join('') +       (prots.length>10?'<div class="t6-tip-more">+'+( prots.length-10)+' more…</div>':'');
    const wRect = wrap.getBoundingClientRect();
    tip.style.left = Math.min(e.clientX-wRect.left+12, wrap.offsetWidth-230)+'px';     tip.style.top  = (e.clientY-wRect.top+12)+'px';     tip.style.display = 'block';
  });
  canvas.addEventListener('mouseleave', () => { tip.style.display='none'; });
}

/* ═══════════════════════════════════════════════════════════════
   SECTION 10 – Group abundance tables
═══════════════════════════════════════════════════════════════ */
function _t6GroupTablesHTML(R) {
  const pmProts = R.pmProteins;
  let html = '<div class="t6-section-hd">Most Abundant Surface Proteins by Experimental Group<span style="font-size:11px;font-weight:400;color:#94a3b8;margin-left:8px">UniProt Cell Membrane · ranked by mean abundance</span></div>';   html += '<div class="t6-group-grid">';
  for (const g of R.groups) {
    const ranked = [...pmProts]
      .filter(p => p.perGroup[g.name] && p.perGroup[g.name].mean != null)
      .sort((a,b) => (b.perGroup[g.name].mean||0) - (a.perGroup[g.name].mean||0))
      .slice(0,20);
    const rows = ranked.map((p,i) => {
      const st = p.perGroup[g.name];
      return '<tr class="t6-prow" onclick="t6OpenDrawer(\''+esc(p.pid)+'\')">'+         '<td>'+(i+1)+'</td>'+         '<td><span class="t6-pid">'+esc(p.pid)+'</span></td>'+         '<td class="t6-num">'+_t6FmtAb(st.mean)+'</td>'+         '<td class="t6-num">'+(st.sd!=null?_t6FmtAb(st.sd):'—')+'</td>'+         '<td class="t6-num">'+st.n+'</td>'+       '</tr>';     }).join('');     html += '<div class="card t6-gtbl-card">'+       '<div class="t6-sec-hd">'+esc(g.name)+'<span class="t6-sec-sub">top 20 most abundant UniProt Cell Membrane proteins · click row for details</span></div>'+       '<div class="t6-search-row"><input class="t6-search" placeholder="Search UniProt ID…" oninput="_t6GrpSearch(this,\'tg-'+_t6SafeId(g.name)+'\')"></div>'+       '<div class="t6-tbl-wrap"><table class="t6-tbl" id="tg-'+_t6SafeId(g.name)+'">'+         '<thead><tr><th>#</th><th>UniProt ID</th>'+         '<th class="t6-num">Mean Abundance</th><th class="t6-num">SD</th><th class="t6-num">n</th></tr></thead>'+         '<tbody>'+rows+'</tbody>'+       '</table></div>'+     '</div>';
  }
  html += '</div>';
  return html;
}

function _t6SingleGroupHTML(R) {
  /* Single group — top 20 UniProt PM proteins */
  const top20 = R.pmProteins.slice(0,20);
  const rows = top20.map((p,i) =>
    '<tr class="t6-prow" onclick="t6OpenDrawer(\''+esc(p.pid)+'\')">'+     '<td>'+(i+1)+'</td>'+     '<td><span class="t6-pid">'+esc(p.pid)+'</span></td>'+     '<td class="t6-num">'+_t6FmtAb(p.meanAbundance)+'</td>'+     '<td>'+esc(p.category)+'</td>'+     '</tr>'   ).join('');   return '<div class="card">'+     '<div class="t6-sec-hd">Most Abundant UniProt Cell Membrane Proteins<span class="t6-sec-sub">top 20 by mean abundance · click row for details</span></div>'+     '<div class="t6-tbl-wrap"><table class="t6-tbl">'+       '<thead><tr><th>#</th><th>UniProt ID</th><th class="t6-num">Mean Abundance</th><th>Category</th></tr></thead>'+       '<tbody>'+rows+'</tbody>'+     '</table></div>'+   '</div>';
}

/* ═══════════════════════════════════════════════════════════════
   SECTION 11 – Differential analysis tables + volcano
═══════════════════════════════════════════════════════════════ */
function _t6DiffHTML(R, compSel) {
  const comp = R.comparisons[_t6ActiveComp];
  if (!comp) return '';
  const { up, down, diff, labelA, labelB } = comp;

  function diffRow(d, i) {
    const sig = !isNaN(d.pval) && d.pval < 0.05;
    return '<tr class="t6-prow'+(sig?' t6-sig':'')+'" onclick="t6OpenDrawer(\''+esc(d.pid)+'\')">'+       '<td>'+(i+1)+'</td>'+       '<td><span class="t6-pid">'+esc(d.pid)+'</span></td>'+       '<td class="t6-num">'+_t6FmtAb(d.meanA)+'</td>'+       '<td class="t6-num">'+_t6FmtAb(d.meanB)+'</td>'+       '<td class="t6-num t6-fc">'+_t6FmtNum(d.log2fc,2)+'</td>'+       '<td class="t6-num">'+(isNaN(d.pval)?'—':d.pval.toExponential(2))+'</td>'+       '<td class="t6-num">'+(isNaN(d.fdr)?'—':d.fdr.toExponential(2))+'</td>'+     '</tr>';
  }

  const thRow = '<th>#</th><th>UniProt ID</th>'+     '<th class="t6-num">Mean ('+esc(labelA)+')</th>'+     '<th class="t6-num">Mean ('+esc(labelB)+')</th>'+     '<th class="t6-num">Log2 FC</th>'+     '<th class="t6-num"><em>p</em>-value</th>'+     '<th class="t6-num">FDR</th>';

  return '<div class="t6-section-hd">Differential Surfaceome Analysis — UniProt Cell Membrane Proteins<span style="font-size:11px;font-weight:400;color:#94a3b8;margin-left:8px">Welch\'s t-test · Benjamini–Hochberg FDR correction</span></div>'+
    compSel +
    '<div class="t6-diff-grid">'+       '<div class="card t6-diff-card">'+         '<div class="t6-sec-hd t6-up-hd">Upregulated in '+esc(labelB)+'<span class="t6-sec-sub">highest log₂FC · p &lt; 0.05 highlighted · click for details</span></div>'+         '<div class="t6-tbl-wrap"><table class="t6-tbl"><thead><tr>'+thRow+'</tr></thead>'+         '<tbody>'+up.map(diffRow).join('')+'</tbody></table></div>'+         '<button class="btn btn-ghost t6-dlbtn" style="margin-top:10px" onclick="t6Download(\'up\')">Download upregulated</button>'+       '</div>'+       '<div class="card t6-diff-card">'+         '<div class="t6-sec-hd t6-dn-hd">Downregulated in '+esc(labelB)+'<span class="t6-sec-sub">lowest log₂FC · p &lt; 0.05 highlighted · click for details</span></div>'+         '<div class="t6-tbl-wrap"><table class="t6-tbl"><thead><tr>'+thRow+'</tr></thead>'+         '<tbody>'+down.map(diffRow).join('')+'</tbody></table></div>'+         '<button class="btn btn-ghost t6-dlbtn" style="margin-top:10px" onclick="t6Download(\'down\')">Download downregulated</button>'+       '</div>'+     '</div>'+
    /* Volcano plot */
    '<div class="card">' +       '<div class="t6-sec-hd">Volcano Plot — UniProt Cell Membrane Proteins<span class="t6-sec-sub">red = upregulated · blue = downregulated (|log₂FC| &gt; 1, p &lt; 0.05) · hover to identify · click to open details</span></div>'+       '<div class="t6-volcano-wrap"><canvas id="t6-volcano"></canvas></div>'+     '</div>'+
    _t6VolcanoJS(comp);
}

/* Inline <script> approach doesn't work in single-file; render volcano via rAF */
function _t6VolcanoJS(comp) {
  /* Schedule volcano draw after DOM update */
  setTimeout(() => _t6DrawVolcano(comp), 50);
  return '';
}

function _t6DrawVolcano(comp) {
  const canvas = document.getElementById('t6-volcano');
  if (!canvas || !comp) return;
  if (_t6Charts.volcano) { try { _t6Charts.volcano.destroy(); } catch(e){} }

  const { diff, labelA, labelB } = comp;
  const pts = diff.filter(d => isFinite(d.log2fc) && isFinite(d.pval) && d.pval > 0).map(d => ({
    x: d.log2fc,
    y: -Math.log10(d.pval),
    pid: d.pid, gene: d.gene, pval: d.pval, fc: d.log2fc,
    color: d.log2fc >  1 && d.pval < 0.05 ? '#ef4444'          : d.log2fc < -1 && d.pval < 0.05 ? '#1d4ed8'          : '#94a3b8',
  }));

  const sigY = -Math.log10(0.05); /* ≈ 1.301 */

  /* Inline plugin — draws dashed threshold lines after each render */
  const threshPlugin = {
    id: 't6-thresh',
    afterDraw(chart) {
      const { ctx, chartArea: { left, right, top, bottom }, scales: { x, y } } = chart;
      ctx.save();
      ctx.setLineDash([5, 4]);
      ctx.strokeStyle = '#94a3b8';
      ctx.lineWidth = 1;
      /* Horizontal: p = 0.05 */
      const yPx = y.getPixelForValue(sigY);
      if (yPx >= top && yPx <= bottom) {
        ctx.beginPath(); ctx.moveTo(left, yPx); ctx.lineTo(right, yPx); ctx.stroke();
      }
      /* Vertical: |log₂FC| = 1 */
      for (const xv of [-1, 1]) {
        const xPx = x.getPixelForValue(xv);
        if (xPx >= left && xPx <= right) {
          ctx.beginPath(); ctx.moveTo(xPx, top); ctx.lineTo(xPx, bottom); ctx.stroke();
        }
      }
      /* Labels */
      ctx.setLineDash([]);
      ctx.font = '10px system-ui,sans-serif';       ctx.fillStyle = '#94a3b8';       ctx.textAlign = 'left';
      if (yPx >= top && yPx <= bottom) {
        ctx.fillText('p = 0.05', left + 4, yPx - 4);
      }
      ctx.restore();
    }
  };

  _t6Charts.volcano = new Chart(canvas, {
    type: 'scatter',
    data: { datasets: [{
      data: pts,
      pointBackgroundColor: pts.map(p => p.color),
      pointRadius: 4,
      pointHoverRadius: 6,
    }]},
    plugins: [threshPlugin],
    options: {
      responsive: true, maintainAspectRatio: false, animation: false,
      plugins: {
        legend: { display: false },
        tooltip: {
          callbacks: {
            title: items => {
              const d = pts[items[0].dataIndex];
              return d.pid + (d.gene ? ' · ' + d.gene : '');
            },
            label: item => {
              const d = pts[item.dataIndex];
              return [
                'Log₂FC: ' + d.fc.toFixed(2),                 'p-value: ' + d.pval.toExponential(2),                 '−log₁₀(p): ' + d.y.toFixed(2),
              ];
            }
          }
        }
      },
      scales: {
        x: {
          title: { display: true, text: 'Log₂ Fold Change (' + labelB + ' / ' + labelA + ')', font: { size: 11 } },           grid: { color: '#f1f5f9' }, ticks: { font: { size: 11 } }
        },
        y: {
          title: { display: true, text: '−log₁₀(p-value)', font: { size: 11 } },           grid: { color: '#f1f5f9' }, ticks: { font: { size: 11 } }
        }
      },
      onClick: (evt, elements) => {
        if (elements.length > 0) {
          const d = pts[elements[0].index];
          if (d) t6OpenDrawer(d.pid);
        }
      }
    }
  });
}

/* ═══════════════════════════════════════════════════════════════
   SECTION 12 – Condition-specific proteins
═══════════════════════════════════════════════════════════════ */
function _t6CondSpecHTML(R) {
  const cs = R.condSpecific;
  const [gA, gB] = R.groups;
  const rows = cs.map(p =>
    '<tr class="t6-prow" onclick="t6OpenDrawer(\''+esc(p.pid)+'\')">'+     '<td><span class="t6-pid">'+esc(p.pid)+'</span></td>'+     '<td class="t6-num">'+p.detA+'</td>'+     '<td class="t6-num">'+p.detB+'</td>'+     '<td class="t6-num">'+_t6FmtAb(p.meanA)+'</td>'+     '<td class="t6-num">'+_t6FmtAb(p.meanB)+'</td>'+     '<td><span class="t6-spec-tag">'+esc(p.status)+'</span></td>'+     '</tr>'   ).join('');

  return '<div class="card">'+     '<div class="t6-sec-hd">Condition-Specific Surface Proteins — Candidate Biomarkers<span class="t6-sec-sub">UniProt Cell Membrane · detected in one group but not the other</span></div>'+     '<div class="t6-spec-note notice notice-info" style="margin-bottom:12px">'+       'Proteins detected in ≥50% of replicates in one group and &lt;50% in the other are flagged as condition-specific. '+       'These are strong biomarker candidates.'+     '</div>'+     '<div class="t6-tbl-wrap"><table class="t6-tbl">'+       '<thead><tr>'+         '<th>UniProt ID</th>'+         '<th class="t6-num">Detected ('+esc(gA.name)+')</th>'+         '<th class="t6-num">Detected ('+esc(gB.name)+')</th>'+         '<th class="t6-num">Mean ('+esc(gA.name)+')</th>'+         '<th class="t6-num">Mean ('+esc(gB.name)+')</th>'+         '<th>Status</th>'+       '</tr></thead>'+       '<tbody>'+rows+'</tbody>'+     '</table></div>'+     '<button class="btn btn-ghost t6-dlbtn" style="margin-top:10px" onclick="t6Download(\'condspec\')">Download condition-specific</button>'+   '</div>';
}

/* ═══════════════════════════════════════════════════════════════
   SECTION 13 – Full ranked protein table
═══════════════════════════════════════════════════════════════ */
function _t6FullTableHTML(R) {
  const top = R.surface.slice(0, 200);
  const note = R.surface.length > 200
    ? '<div class="t6-tbl-note">Showing top 200 of '+R.surface.length.toLocaleString()+' surface proteins by mean abundance. Download full list below.</div>'     : '';
  const rows = top.map((p,i) => {
    const dbs = [p.inSS?'SURFY':'',p.inPM?'UniProt PM':'',p.inMA?'MA':''].filter(Boolean).join(', ');     return '<tr class="t6-prow" onclick="t6OpenDrawer(\''+esc(p.pid)+'\')">'+       '<td>'+(i+1)+'</td>'+       '<td><span class="t6-pid">'+esc(p.pid)+'</span></td>'+       '<td class="t6-num">'+_t6FmtAb(p.meanAbundance)+'</td>'+       '<td>'+esc(p.primaryCat)+'</td>'+       '<td>'+esc(dbs)+'</td>'+     '</tr>';   }).join('');   return '<div class="card">'+     '<div class="t6-sec-hd">All Detected Surface Proteins<span class="t6-sec-sub">all databases · sorted by mean abundance · click row for protein details</span></div>'+     '<div class="t6-search-row"><input class="t6-search" id="t6-full-search" placeholder="Search UniProt ID…" oninput="_t6FullSearch(this)"></div>'+
    note+
    '<div class="t6-tbl-wrap" id="t6-full-wrap"><table class="t6-tbl" id="t6-full-tbl">'+       '<thead><tr><th>#</th><th>UniProt ID</th>'+       '<th class="t6-num">Mean Abundance</th><th>Primary Category</th><th>Databases</th></tr></thead>'+       '<tbody>'+rows+'</tbody>'+     '</table></div>'+   '</div>';
}

/* ═══════════════════════════════════════════════════════════════
   SECTION 14 – Interactive table search
═══════════════════════════════════════════════════════════════ */
function _t6SafeId(name) { return name.replace(/[^a-zA-Z0-9]/g, '_'); }

function _t6GrpSearch(input, tableId) {
  const q = input.value.trim().toLowerCase();
  const tbl = document.getElementById(tableId);
  if (!tbl) return;
  for (const tr of tbl.querySelectorAll('tbody tr')) {
    const text = tr.textContent.toLowerCase();
    tr.style.display = (!q || text.includes(q)) ? '' : 'none';
  }
}

function _t6FullSearch(input) {
  const q = input.value.trim().toLowerCase();
  const tbl = document.getElementById('t6-full-tbl');
  if (!tbl) return;
  for (const tr of tbl.querySelectorAll('tbody tr')) {
    const text = tr.textContent.toLowerCase();
    tr.style.display = (!q || text.includes(q)) ? '' : 'none';
  }
}

function _t6MatrixSearch(input) {
  const q = input.value.trim().toLowerCase();
  const tbl = document.getElementById('t6-matrix-tbl');
  if (!tbl) return;
  for (const tr of tbl.querySelectorAll('tbody tr')) {
    const text = tr.textContent.toLowerCase();
    tr.style.display = (!q || text.includes(q)) ? '' : 'none';
  }
}

/* ── Database membership matrix ── */
function _t6MembershipMatrixHTML(R) {
  const top  = R.surface.slice(0, 100);
  const ck   = v => v
    ? '<span class="t6-ck-yes">&#x2713;</span>'     : '<span class="t6-ck-no">&#x2014;</span>';

  const rows = top.map((p, i) => {
    const dbCount = (p.inSS ? 1 : 0) + (p.inPM ? 1 : 0) + (p.inMA ? 1 : 0);
    const dbBadge = dbCount === 3
      ? '<span class="t6-db3">3/3</span>'
      : dbCount === 2
        ? '<span class="t6-db2">2/3</span>'         : '<span class="t6-db1">1/3</span>';     return '<tr class="t6-prow" onclick="t6OpenDrawer(\'' + esc(p.pid) + '\')">' +       '<td>' + (i + 1) + '</td>' +       '<td><span class="t6-pid">' + esc(p.pid) + '</span></td>' +       '<td class="t6-ck-cell">' + ck(p.inSS) + '</td>' +       '<td class="t6-ck-cell">' + ck(p.inPM) + '</td>' +       '<td class="t6-ck-cell">' + ck(p.inMA) + '</td>' +       '<td class="t6-num">' + dbBadge + '</td>' +       '<td class="t6-num">' + _t6FmtAb(p.meanAbundance) + '</td>' +     '</tr>';   }).join('');

  const note = R.surface.length > 100
    ? '<div class="t6-tbl-note">Showing top 100 of ' + R.surface.length.toLocaleString() +       ' surface proteins by mean abundance. Use Downloads to export the full list.</div>'     : '';

  return '<div class="card">' +     '<div class="t6-sec-hd">Database Membership Matrix' +       '<span class="t6-sec-sub">evidence strength per protein · click row for details</span>' +     '</div>' +     '<div class="notice notice-info" style="margin-bottom:10px;font-size:12px">' +       'A &#x2713; indicates the protein is annotated in that surfaceome database. ' +       'Proteins supported by multiple independent databases represent the strongest evidence ' +       'for genuine cell-surface localisation.' +     '</div>' +     '<div class="t6-search-row">' +       '<input class="t6-search" id="t6-matrix-search" placeholder="Search UniProt ID…" ' +              'oninput="_t6MatrixSearch(this)">' +     '</div>' +
    note +
    '<div class="t6-tbl-wrap"><table class="t6-tbl t6-matrix-tbl" id="t6-matrix-tbl">' +       '<thead><tr>' +         '<th>#</th>' +         '<th>UniProt ID</th>' +         '<th class="t6-ck-cell">SURFY</th>' +         '<th class="t6-ck-cell">UniProt Cell Membrane</th>' +         '<th class="t6-ck-cell">Membrane-Associated</th>' +         '<th class="t6-num">Databases Supported</th>' +         '<th class="t6-num">Mean Abundance</th>' +       '</tr></thead>' +       '<tbody>' + rows + '</tbody>' +     '</table></div>' +   '</div>';
}

function _t6SetComp(idx) {
  _t6ActiveComp = idx;
  renderTab6();
}

/* ═══════════════════════════════════════════════════════════════
   SECTION 15 – Protein detail drawer
═══════════════════════════════════════════════════════════════ */
function t6OpenDrawer(pid) {
  const R = _t6Results;
  if (!R) return;
  const p = R.proteins.find(x => x.pid === pid);
  if (!p) return;

  const src = _t6GetSource();
  const row = src ? src.rows[p.rowIdx] : null;

  /* Max abundance */
  let maxAb = null;
  if (row && R.qCols) {
    let mx = -Infinity;
    for (const c of R.qCols) {
      const v = parseFloat(row[c]);
      if (!isNaN(v) && v > mx) mx = v;
    }
    if (isFinite(mx)) maxAb = mx;
  }

  const dbs = [p.inSS?'<span class="t6-db-ss">SURFY</span>':'',                p.inPM?'<span class="t6-db-pm">UniProt PM</span>':'',                p.inMA?'<span class="t6-db-ma">Mem.-Assoc.</span>':''].filter(Boolean).join(' ') || '—';

  /* Per-group rows */
  let grpRows = '';
  if (R.hasGroups) {
    grpRows = R.groups.map(g => {
      const st = p.perGroup[g.name];
      if (!st) return '';       return '<div class="t6-dr-row"><span class="t6-dr-lbl">'+esc(g.name)+'</span>'+         '<span>'+(st.mean!=null?st.mean.toExponential(2):'n.d.')+         (st.sd?' ± '+st.sd.toExponential(1):'')+' (n='+st.n+')</span></div>';     }).join('');
  }

  const drawer  = document.getElementById('t6-drawer');   const overlay = document.getElementById('t6-overlay');
  drawer.innerHTML =
    '<div class="t6-dr-hd">'+       '<div class="t6-dr-title">'+esc(p.pid)+'</div>'+       '<button class="t6-dr-close" onclick="t6CloseDrawer()">&#x2715;</button>'+     '</div>'+     '<div class="t6-dr-body">'+       '<div class="t6-dr-row"><span class="t6-dr-lbl">UniProt ID</span>'+         '<a class="t6-extlink" href="https://www.uniprot.org/uniprot/'+esc(p.pid)+'" target="_blank" rel="noopener">'+esc(p.pid)+' &#x2197;</a></div>'+       (p.gene?'<div class="t6-dr-row"><span class="t6-dr-lbl">Gene Symbol</span><span>'+esc(p.gene)+'</span></div>':'')+       (p.name?'<div class="t6-dr-row"><span class="t6-dr-lbl">Protein Name</span><span style="font-size:12px;color:#334155">'+esc(p.name)+'</span></div>':'')+       '<div class="t6-dr-row"><span class="t6-dr-lbl">Primary Category</span><span>'+esc(p.primaryCat)+'</span></div>'+       '<div class="t6-dr-row"><span class="t6-dr-lbl">Databases</span><span>'+dbs+'</span></div>'+       '<div class="t6-dr-row"><span class="t6-dr-lbl">DB Count</span><span>'+p.dbCount+' / 3</span></div>'+       '<div class="t6-dr-sep"></div>'+       '<div class="t6-dr-row"><span class="t6-dr-lbl">Mean Abundance</span><span>'+_t6FmtAb(p.meanAbundance)+'</span></div>'+       (maxAb!=null?'<div class="t6-dr-row"><span class="t6-dr-lbl">Max Abundance</span><span>'+maxAb.toExponential(2)+'</span></div>':'')+
      grpRows+
      '<div class="t6-dr-sep"></div>'+       '<div class="t6-dr-links">'+         '<a class="t6-extbtn" href="https://www.uniprot.org/uniprot/'+esc(p.pid)+'" target="_blank" rel="noopener">UniProt</a>'+         '<a class="t6-extbtn" href="https://string-db.org/network/'+esc(p.pid)+'" target="_blank" rel="noopener">STRING</a>'+         '<a class="t6-extbtn" href="https://www.ncbi.nlm.nih.gov/protein/'+esc(p.pid)+'" target="_blank" rel="noopener">NCBI</a>'+         (p.gene?'<a class="t6-extbtn" href="https://www.genecards.org/cgi-bin/carddisp.pl?gene='+esc(p.gene)+'" target="_blank" rel="noopener">GeneCards</a>':'')+       '</div>'+     '</div>';

  drawer.style.display = overlay.style.display = 'block';
  _t6DrawerPid = pid;
}
function t6CloseDrawer() {
  const d=document.getElementById('t6-drawer'), o=document.getElementById('t6-overlay');   if(d) d.style.display='none'; if(o) o.style.display='none';
  _t6DrawerPid=null;
}

/* ═══════════════════════════════════════════════════════════════
   SECTION 16 – Downloads
═══════════════════════════════════════════════════════════════ */
/* ── PNG downloads ── */
function _t6DownloadChartPNG(canvasId, filename) {
  const chart = _t6Charts.bar;
  if (chart) {
    const url = chart.toBase64Image('image/png', 1);     const a = document.createElement('a');
    a.href = url; a.download = filename;
    document.body.appendChild(a); a.click(); document.body.removeChild(a);
  }
}
function _t6DownloadVennPNG() {
  const canvas = document.getElementById('t6-venn-canvas');
  if (!canvas) return;
  /* Draw white background first so PNG isn't transparent */   const tmp = document.createElement('canvas');
  tmp.width = canvas.width; tmp.height = canvas.height;
  const tctx = tmp.getContext('2d');   tctx.fillStyle = '#ffffff';
  tctx.fillRect(0, 0, tmp.width, tmp.height);
  tctx.drawImage(canvas, 0, 0);
  const url = tmp.toDataURL('image/png');   const a = document.createElement('a');   a.href = url; a.download = 'venn_diagram.png';
  document.body.appendChild(a); a.click(); document.body.removeChild(a);
}

function t6Download(type) {
  const R = _t6Results;
  if (!R) return;
  const src = _t6GetSource();
  const origHeaders = src ? src.headers : R.headers || [];

  /* Venn region downloads */
  const regionKeys = ['ssOnly','pmOnly','maOnly','ssAndPm','ssAndMa','pmAndMa','allThree'];
  const regionNames = {
    ssOnly:'SURFY_only', pmOnly:'UniProtPM_only', maOnly:'MemAssoc_only',     ssAndPm:'SURFY_UniProtPM', ssAndMa:'SURFY_MemAssoc',     pmAndMa:'UniProtPM_MemAssoc', allThree:'all_three_databases',
  };
  if (regionKeys.includes(type)) {
    const prots = R.regions[type] || [];
    if (!prots.length) { showToast('No proteins in this region.'); return; }     const hdr = ['UniProt_ID','Mean_Abundance','Category'];     const rows = prots.map(p => [p.pid, p.meanAbundance!=null?p.meanAbundance.toFixed(4):'', p.category].map(_t6CsvEsc).join(','));     _t6DownloadCSV([hdr.join(','),...rows].join('\n'), 'venn_'+regionNames[type]+'.csv');     showToast('Downloaded '+prots.length+' proteins → venn_'+regionNames[type]+'.csv');
    return;
  }

  let proteins, filename;
  switch (type) {
    case 'all':      proteins = R.surface;                                filename = 'surfaceome_all.csv'; break;     case 'pm':       proteins = R.proteins.filter(p=>p.inPM);             filename = 'surfaceome_UniProtPM.csv'; break;     case 'ss':       proteins = R.proteins.filter(p=>p.inSS);             filename = 'surfaceome_SURFY.csv'; break;     case 'ma':       proteins = R.proteins.filter(p=>p.inMA);             filename = 'surfaceome_MembraneAssoc.csv'; break;     case 'none':     proteins = R.proteins.filter(p=>p.dbCount===0);      filename = 'non_membrane.csv'; break;     case 'up': {
      const comp = R.comparisons[_t6ActiveComp];
      if (!comp) return;
      const rows = comp.up.map(d => ({ ...d, rawId:d.pid, gene:d.gene, rowIdx:-1, perSample:{}, perGroup:{} }));
      _t6DownloadDiff(comp.up, comp.labelA, comp.labelB, 'upregulated_UniProtPM.csv');
      return;
    }
    case 'down': {
      const comp = R.comparisons[_t6ActiveComp];
      if (!comp) return;
      _t6DownloadDiff(comp.down, comp.labelA, comp.labelB, 'downregulated_UniProtPM.csv');
      return;
    }
    case 'condspec': {
      if (!R.condSpecific) return;
      const [gA,gB] = R.groups;
      const hdr = ['UniProt_ID','Gene','Detected_'+gA.name,'Detected_'+gB.name,'Mean_'+gA.name,'Mean_'+gB.name,'Status'];
      const rows = R.condSpecific.map(p =>
        [p.pid, p.gene||'', p.detA, p.detB,          p.meanA!=null?p.meanA.toFixed(4):'',          p.meanB!=null?p.meanB.toFixed(4):'', p.status].map(_t6CsvEsc).join(',')
      );
      _t6DownloadCSV([hdr.join(','),...rows].join('\n'), 'condition_specific.csv');
      return;
    }
    default: return;
  }

  if (!proteins || !proteins.length) { showToast('No proteins in this category.'); return; }

  const extraHdrs = ['Surfaceome_Category','In_SURFY','In_UniProt_PM','In_Membrane_Assoc','DB_Count','Mean_Abundance'];
  const csvHdr = [...origHeaders, ...extraHdrs];
  const rows = proteins.map(p => {
    const srcRow = src ? src.rows[p.rowIdx] : null;
    const orig = origHeaders.map(h => _t6CsvEsc(srcRow ? String(srcRow[h]??'') : ''));
    const ext  = [p.category, p.inSS?1:0, p.inPM?1:0, p.inMA?1:0, p.dbCount,
                  p.meanAbundance!=null?p.meanAbundance.toFixed(4):''].map(_t6CsvEsc);     return [...orig,...ext].join(',');
  });
  _t6DownloadCSV([csvHdr.map(_t6CsvEsc).join(','), ...rows].join('\n'), filename);   showToast('Downloaded '+proteins.length.toLocaleString()+' proteins → '+filename);
}

function _t6DownloadDiff(items, labelA, labelB, filename) {
  const hdr = ['UniProt_ID','Gene','Mean_'+labelA,'Mean_'+labelB,'Log2FC','p_value','FDR'];
  const rows = items.map(d =>
    [d.pid,d.gene||'',      d.meanA!=null?d.meanA.toFixed(4):'', d.meanB!=null?d.meanB.toFixed(4):'',      isFinite(d.log2fc)?d.log2fc.toFixed(4):'',      isNaN(d.pval)?'':d.pval.toExponential(4),      isNaN(d.fdr)?'':d.fdr.toExponential(4)].map(_t6CsvEsc).join(',')
  );
  _t6DownloadCSV([hdr.map(_t6CsvEsc).join(','),...rows].join('\n'), filename);   showToast('Downloaded '+items.length+' proteins → '+filename);
}

/* ═══════════════════════════════════════════════════════════════
   SECTION 17 – Comprehensive Analysis Report (12 sections)
   Opens in a new tab for browser print → PDF
═══════════════════════════════════════════════════════════════ */
function t6DownloadReport() {
  const R   = _t6Results;
  const raw = App.raw;
  if (!raw || !R) { showToast('Run the analysis first.'); return; }

  showToast('Building report…');

  /* ── Capture chart images ── */
  const barImg     = (_t6Charts.bar)     ? _t6Charts.bar.toBase64Image('image/png', 1.0)     : null;   const volcanoImg = (_t6Charts.volcano) ? _t6Charts.volcano.toBase64Image('image/png', 1.0) : null;
  let vennImg = null;
  const _vennEl = document.getElementById('t6-venn-canvas');
  if (_vennEl) {
    const _tmp = document.createElement('canvas');
    _tmp.width = _vennEl.width; _tmp.height = _vennEl.height;
    const _tc = _tmp.getContext('2d');     _tc.fillStyle = '#ffffff'; _tc.fillRect(0, 0, _tmp.width, _tmp.height);
    _tc.drawImage(_vennEl, 0, 0);
    vennImg = _tmp.toDataURL('image/png');
  }

  /* ── Date / timestamp ── */
  const now     = new Date();
  const dateStr = now.toLocaleDateString('en-US', { year:'numeric', month:'long', day:'numeric' });   const timeStr = now.toLocaleTimeString('en-US', { hour:'2-digit', minute:'2-digit' });   const ts      = now.getFullYear() + '-' +                   String(now.getMonth()+1).padStart(2,'0') + '-' +                   String(now.getDate()).padStart(2,'0');

  /* ── Helpers ── */
  const pc  = (v, n) => n > 0 ? (v/n*100).toFixed(1) + '%' : '0.0%';   const NA  = '—';
  function rRow(label, value) {
    if (value === undefined || value === null || value === '') return '';     return '<tr><th>' + label + '</th><td>' + value + '</td></tr>';
  }
  function rSec(label) {
    return '<tr class="r-sec"><td colspan="2">' + label + '</td></tr>';
  }

  /* ── 1. Dataset summary ── */
  const fname  = App.file ? App.file.name : 'Unknown';   const ftype  = raw.fileType ? raw.fileType.toUpperCase() : 'Unknown';
  const nProts = raw.nRows;
  const nSamps  = App.config.quantCols.length;
  const groups  = R.groups || [];
  const groupList = groups.map((g, gi) =>
    'Group ' + (gi+1) + ': <strong>' + esc(g.name) + '</strong> (' + g.cols.length + ' replicate' + (g.cols.length !== 1 ? 's' : '') + ')'   ).join('<br>');

  /* ── 2. QC ── */
  const qc = App.check || {};
  const nConstant = (qc.constantFeatureIndices || []).length;
  const qcStatusIcon = s => s === 'pass' ? '<span style="color:#059669;font-weight:700">&#10003; Pass</span>'                            : s === 'warn' ? '<span style="color:#d97706;font-weight:700">&#9888; Warning</span>'                            : '<span style="color:#dc2626;font-weight:700">&#10007; Error</span>';
  const qcChecksHTML = (qc.checks || []).map(c => {
    const icon = c.status === 'pass' ? '&#10003;' : c.status === 'warn' ? '&#9888;' : '&#10007;';     const color = c.status === 'pass' ? '#059669' : c.status === 'warn' ? '#d97706' : '#dc2626';     return '<tr><td style="width:35%">' + esc(c.label) + '</td>' +       '<td style="width:8%;text-align:center;font-weight:700;color:' + color + '">' + icon + '</td>' +       '<td>' + c.message + '</td></tr>';   }).join('');

  /* ── 3. Filtering ── */
  const filt = App.filtered;
  const fs   = filt && filt.filterSettings;
  const fr   = filt && filt.filterResults;
  const VL = { iqr:'IQR', sd:'Standard deviation', mad:'MAD', rsd:'Relative SD' };   const AL = { mean:'Mean intensity', median:'Median intensity' };   let filtRows = '';
  if (fs) {
    filtRows = rRow('Proteins before filtering', fr ? fr.nBefore.toLocaleString() : nProts.toLocaleString()) +       rRow('Missing-value filter', fs.missing.enabled         ? 'Enabled &mdash; threshold: ' + fs.missing.threshold + '% missing per protein' +           (fr ? ' (removed: ' + fr.removedMissing.toLocaleString() + ')' : '')         : 'Not applied') +       rRow('Low-variance filter', fs.variance.enabled         ? 'Enabled &mdash; bottom ' + fs.variance.pct + '% by ' + (VL[fs.variance.metric]||fs.variance.metric) +           (fr ? ' (removed: ' + fr.removedVariance.toLocaleString() + ')' : '')         : 'Not applied') +       rRow('Low-abundance filter', fs.abundance.enabled         ? 'Enabled &mdash; bottom ' + fs.abundance.pct + '% by ' + (AL[fs.abundance.metric]||fs.abundance.metric) +           (fr ? ' (removed: ' + fr.removedAbundance.toLocaleString() + ')' : '')         : 'Not applied') +       rRow('Constant-feature filter', fs.constant.enabled         ? 'Enabled' + (fr ? ' (removed: ' + fr.removedConstant.toLocaleString() + ')' : '')         : 'Not applied') +       rRow('Proteins retained', fr         ? '<strong>' + fr.nKept.toLocaleString() + '</strong> of ' + fr.nBefore.toLocaleString()
        : NA);
  } else {
    filtRows = rRow('Status', 'Not applied &mdash; raw data passed through without modification');
  }

  /* ── 4. Imputation ── */
  const imp = App.imputed;
  const impMethodLabels = {
    lod:     'Limit of Detection (LoD) &mdash; replaced with 1/5 of minimum observed value per protein',     halfmin: 'Half-Minimum &mdash; replaced with half the minimum observed value per protein',     mean:    'Mean imputation &mdash; replaced with row mean of observed values',     knn:     'K-Nearest Neighbours (KNN, k&nbsp;=&nbsp;5) &mdash; estimated from 5 most similar proteins',     none:    'Not applied (no missing values)',
  };
  let impRows = '';
  if (imp) {
    impRows = rRow('Method', impMethodLabels[imp.method] || (imp.method || NA)) +
      (imp.nImputed != null
        ? rRow('Values imputed', imp.nImputed.toLocaleString() + ' values across ' + (imp.proteinsAffected||0).toLocaleString() + ' proteins')         : '');
  } else {
    impRows = rRow('Status', 'Not applied');
  }

  /* ── 5. Normalization ── */
  const norm = App.normalized;
  const SN_L = { none:'None', median:'Median normalization', sum:'Sum/total intensity normalization', quantile:'Quantile normalization' };   const TR_L = { none:'None', log2:'Log&#x2082; transformation', log10:'Log&#x2081;&#x2080; transformation', sqrt:'Square root transformation' };   const SC_L = { none:'None', center:'Mean centering', auto:'Auto / Z-score scaling', pareto:'Pareto scaling' };   let normRows = '';
  if (norm && norm.state) {
    normRows = rRow('Sample normalization', SN_L[norm.state.sampleNorm] || norm.state.sampleNorm || 'None') +       rRow('Data transformation',   TR_L[norm.state.transform]   || norm.state.transform   || 'None') +       rRow('Feature scaling',       SC_L[norm.state.scaling]     || norm.state.scaling     || 'None') +       (norm.pseudocount != null ? rRow('Pseudocount added', norm.pseudocount.toExponential(4)) : '');
  } else {
    normRows = rRow('Status', 'Not applied &mdash; data passed through without normalization');
  }

  /* ── 6. Surfaceome analysis ── */
  const highConf = R.regions.ssAndPm.length + R.regions.ssAndMa.length +
                   R.regions.pmAndMa.length + R.regions.allThree.length;

  /* ── 7. Database overlap region table ── */
  const regionTableRows = [
    ['SURFY only',                                 R.regions.ssOnly.length],     ['UniProt Cell Membrane only',                 R.regions.pmOnly.length],     ['Membrane-Associated only',                   R.regions.maOnly.length],     ['SURFY &#x2229; UniProt Cell Membrane',       R.regions.ssAndPm.length],     ['SURFY &#x2229; Membrane-Associated',         R.regions.ssAndMa.length],     ['UniProt Cell Membrane &#x2229; Membrane-Assoc.', R.regions.pmAndMa.length],     ['All three databases',                        R.regions.allThree.length],
  ].map(([label, n]) =>
    '<tr><td>' + label + '</td>' +     '<td style="text-align:right">' + n + '</td>' +     '<td style="text-align:right">' + pc(n, R.totalN) + '</td></tr>'   ).join('');

  /* ── 8. Top 20 per group ── */
  let top20HTML = '';
  if (groups.length >= 1) {
    top20HTML = groups.map(g => {
      const top = R.pmProteins
        .filter(p => p.perGroup[g.name] && p.perGroup[g.name].mean != null)
        .sort((a, b) => (b.perGroup[g.name].mean || 0) - (a.perGroup[g.name].mean || 0))
        .slice(0, 20);
      if (!top.length) return '';
      const rowsHTML = top.map((p, i) => {
        const st = p.perGroup[g.name];
        return '<tr>' +           '<td style="text-align:center">' + (i+1) + '</td>' +           '<td>' + esc(p.pid) + '</td>' +           '<td>' + esc(p.gene || NA) + '</td>' +           '<td style="text-align:right">' + (st.mean != null ? st.mean.toFixed(2) : NA) + '</td>' +           '<td style="text-align:right">' + (st.sd  != null ? st.sd.toFixed(2)  : NA) + '</td>' +           '</tr>';       }).join('');       return '<h3 style="margin:20px 0 8px">Group: ' + esc(g.name) + '</h3>' +         '<table class="data-tbl">' +           '<thead><tr>' +             '<th style="width:8%;text-align:center">Rank</th>' +             '<th style="width:22%">UniProt ID</th>' +             '<th style="width:22%">Gene</th>' +             '<th style="text-align:right;width:24%">Mean Abundance</th>' +             '<th style="text-align:right;width:24%">SD</th>' +           '</tr></thead>' +           '<tbody>' + rowsHTML + '</tbody>' +         '</table>';     }).join('');     if (!top20HTML) top20HTML = '<p class="no-data">No UniProt Cell Membrane proteins found in this dataset.</p>';
  }

  /* ── 9. Differential analysis ── */
  let diffHTML = '';
  if (R.comparisons.length > 0) {
    diffHTML = R.comparisons.map(comp => {
      const sigUp   = comp.diff.filter(d => isFinite(d.log2fc) && d.log2fc >  1 && d.fdr < 0.05).length;
      const sigDown = comp.diff.filter(d => isFinite(d.log2fc) && d.log2fc < -1 && d.fdr < 0.05).length;
      const upRows  = comp.up.slice(0,10).map((d, i) =>
        '<tr><td style="text-align:center">' + (i+1) + '</td>' +         '<td>' + esc(d.pid) + '</td>' +         '<td>' + esc(d.gene||NA) + '</td>' +         '<td style="text-align:right;color:#dc2626">' + d.log2fc.toFixed(2) + '</td>' +         '<td style="text-align:right">' + d.pval.toExponential(2) + '</td>' +         '<td style="text-align:right">' + d.fdr.toExponential(2) + '</td></tr>'       ).join('');
      const downRows = comp.down.slice(0,10).map((d, i) =>
        '<tr><td style="text-align:center">' + (i+1) + '</td>' +         '<td>' + esc(d.pid) + '</td>' +         '<td>' + esc(d.gene||NA) + '</td>' +         '<td style="text-align:right;color:#1d4ed8">' + d.log2fc.toFixed(2) + '</td>' +         '<td style="text-align:right">' + d.pval.toExponential(2) + '</td>' +         '<td style="text-align:right">' + d.fdr.toExponential(2) + '</td></tr>'       ).join('');       const diffTblHdr = '<thead><tr>' +         '<th style="width:8%;text-align:center">Rank</th>' +         '<th style="width:20%">UniProt ID</th><th style="width:18%">Gene</th>' +         '<th style="text-align:right">Log&#x2082;FC</th>' +         '<th style="text-align:right"><em>p</em>-value</th>' +         '<th style="text-align:right">FDR</th></tr></thead>';       return '<h3 style="margin:20px 0 8px">' + esc(comp.label) + '</h3>' +         '<table class="meta-tbl">' +           rRow('UniProt Cell Membrane proteins tested', comp.diff.length.toLocaleString()) +           rRow('Significantly upregulated (|log&#x2082;FC|&nbsp;&gt;&nbsp;1, FDR&nbsp;&lt;&nbsp;0.05)',                '<strong style="color:#dc2626">' + sigUp + '</strong>') +           rRow('Significantly downregulated (|log&#x2082;FC|&nbsp;&gt;&nbsp;1, FDR&nbsp;&lt;&nbsp;0.05)',                '<strong style="color:#1d4ed8">' + sigDown + '</strong>') +         '</table>' +
        (volcanoImg
          ? '<figure><img src="' + volcanoImg + '" alt="Volcano plot" style="max-width:100%;height:auto;border:1px solid #e2e8f0;border-radius:4px">' +             '<figcaption>Figure 3. Volcano plot &mdash; ' + esc(comp.label) + '. Red: upregulated (log&#x2082;FC&nbsp;&gt;&nbsp;1, FDR&nbsp;&lt;&nbsp;0.05); Blue: downregulated; Gray: not significant. Dashed lines: |log&#x2082;FC|&nbsp;=&nbsp;1 and <em>p</em>&nbsp;=&nbsp;0.05.</figcaption></figure>'           : '') +         (upRows ? '<h4 style="margin:16px 0 6px">Top Upregulated in ' + esc(comp.labelB) + '</h4>' +           '<table class="data-tbl">' + diffTblHdr + '<tbody>' + upRows + '</tbody></table>' : '') +         (downRows ? '<h4 style="margin:16px 0 6px">Top Downregulated in ' + esc(comp.labelB) + '</h4>' +           '<table class="data-tbl">' + diffTblHdr + '<tbody>' + downRows + '</tbody></table>' : '');     }).join('<div style="margin:24px 0;border-top:1px solid #e2e8f0"></div>');
  } else {
    diffHTML = '<p class="no-data">No differential analysis performed &mdash; requires at least two experimental groups with &ge;3 replicates each.</p>';
  }

  /* ── 10. Condition-specific proteins ── */
  let condHTML = '<p class="no-data">No condition-specific proteins detected, or requires exactly two experimental groups.</p>';
  if (R.condSpecific && R.condSpecific.length > 0 && R.groups.length >= 2) {
    const [gA, gB] = R.groups;
    const csRows = R.condSpecific.map(p =>
      '<tr><td>' + esc(p.pid) + '</td><td>' + esc(p.gene||NA) + '</td>' +       '<td style="text-align:center">' + esc(p.detA) + '</td>' +       '<td style="text-align:center">' + esc(p.detB) + '</td>' +       '<td><strong>' + esc(p.status) + '</strong></td></tr>'     ).join('');     condHTML = '<table class="data-tbl"><thead><tr>' +       '<th>UniProt ID</th><th>Gene</th>' +       '<th style="text-align:center">Detected (' + esc(gA.name) + ')</th>' +       '<th style="text-align:center">Detected (' + esc(gB.name) + ')</th>' +       '<th>Status</th></tr></thead><tbody>' + csRows + '</tbody></table>';
  }

  /* ── 11. Analysis parameters ── */
  const sampleList = App.config.quantCols.map((c,i) => (i+1) + '. ' + esc(c)).join('<br>');
  let paramRows =
    rRow('Software', 'Surfaceome Explorer v2 (browser-based, fully client-side)') +     rRow('Analysis date', dateStr + ' at ' + timeStr) +     rRow('File name', esc(fname)) +     rRow('File format', ftype) +     rRow('Protein ID column', esc(App.config.idCol)) +     (App.config.geneCol ? rRow('Gene symbol column', esc(App.config.geneCol)) : '') +     rRow('Total proteins (input)', nProts.toLocaleString()) +     rRow('Quantitative samples', nSamps.toLocaleString()) +     rRow('Sample list', sampleList) +     rRow('Experimental groups', groupList) +     rRow('Missing value rate (input)', qc.missPct != null ? qc.missPct + '%' : NA) +     rSec('Data Filtering') +     rRow('Missing-value threshold', fs && fs.missing.enabled ? fs.missing.threshold + '%' : 'Disabled') +     rRow('Variance filter', fs && fs.variance.enabled ? 'Bottom ' + fs.variance.pct + '% by ' + (VL[fs.variance.metric]||fs.variance.metric||'—') : 'Disabled') +     rRow('Abundance filter', fs && fs.abundance.enabled ? 'Bottom ' + fs.abundance.pct + '% by ' + (AL[fs.abundance.metric]||fs.abundance.metric||'—') : 'Disabled') +     rRow('Constant-feature filter', fs && fs.constant.enabled ? 'Enabled' : 'Disabled') +     (fr ? rRow('Proteins retained', fr.nKept.toLocaleString() + ' of ' + fr.nBefore.toLocaleString()) : '') +     rSec('Imputation') +     rRow('Method', imp ? (impMethodLabels[imp.method] || imp.method || NA) : 'Not applied') +     (imp && imp.nImputed != null ? rRow('Values imputed', imp.nImputed.toLocaleString()) : '') +     rSec('Normalization') +     rRow('Sample normalization', norm && norm.state ? (SN_L[norm.state.sampleNorm] || 'None') : 'Not applied') +     rRow('Data transformation',  norm && norm.state ? (TR_L[norm.state.transform]   || 'None') : 'Not applied') +     rRow('Feature scaling',      norm && norm.state ? (SC_L[norm.state.scaling]     || 'None') : 'Not applied') +     rSec('Surfaceome Databases') +     rRow('SURFY', '2,886 human proteins &mdash; in-silico surfaceome predictor') +     rRow('UniProt Cell Membrane', '6,944 human proteins &mdash; keyword KW-1003') +     rRow('Membrane-Associated', '5,973 human proteins') +     rSec('Statistical Testing') +     rRow('Differential test', 'Welch\'s t-test (unequal variance; Welch&ndash;Satterthwaite df)') +     rRow('Multiple testing correction', 'Benjamini&ndash;Hochberg false discovery rate (FDR)') +     rRow('Significance thresholds', '|log&#x2082;FC| &gt; 1 and FDR &lt; 0.05') +     rRow('Condition-specific threshold', '&ge;50% replicate detection in one group, &lt;50% in the other');

  /* ── Compose full HTML report ── */
  const html = '<!DOCTYPE html>\n<html lang="en">\n<head>\n' + '<meta charset="UTF-8">\n' + '<meta name="viewport" content="width=device-width,initial-scale=1">\n' + '<title>Surfaceome Explorer — Analysis Report — ' + dateStr + '</title>\n' + '<style>\n' + '*{box-sizing:border-box;margin:0;padding:0}\n' + 'body{font-family:Georgia,"Times New Roman",serif;font-size:13.5px;line-height:1.75;color:#1a1a2e;background:#f8fafc}\n' + '.page-wrap{max-width:900px;margin:0 auto;padding:60px 64px 80px;background:#fff;box-shadow:0 0 60px rgba(0,0,0,.08);min-height:100vh}\n' + '.toolbar{position:sticky;top:0;z-index:100;background:#1e293b;color:#f1f5f9;display:flex;align-items:center;gap:12px;padding:11px 20px;font-family:system-ui,sans-serif;font-size:13px}\n' + '.toolbar-title{flex:1;font-weight:600;font-size:13.5px}\n' + '.toolbar-note{font-size:11px;color:#94a3b8}\n' + '.tbtn{padding:7px 16px;border-radius:6px;font-size:12px;font-weight:600;cursor:pointer;border:none;font-family:system-ui,sans-serif}\n' + '.tbtn-print{background:#2563eb;color:#fff}\n' + '.tbtn-print:hover{background:#1d4ed8}\n' + '.tbtn-dl{background:#16a34a;color:#fff}\n' + '.tbtn-dl:hover{background:#15803d}\n' + '.cover-title{font-size:26px;font-weight:700;color:#0f172a;margin-bottom:4px;letter-spacing:-.01em}\n' + '.cover-sub{font-size:14px;color:#64748b;margin-bottom:6px;font-style:italic}\n' + '.cover-meta{font-size:12px;color:#94a3b8;font-family:system-ui,sans-serif;margin-bottom:36px}\n' + '.cover-rule{border:none;border-top:2px solid #0f172a;margin-bottom:36px}\n' + '.sec-num{font-size:11px;font-weight:700;color:#2563eb;font-family:system-ui,sans-serif;text-transform:uppercase;letter-spacing:.08em;margin-bottom:4px}\n' + 'h2{font-size:16px;font-weight:700;color:#0f172a;margin:44px 0 6px;letter-spacing:-.01em}\n' + 'h2.first{margin-top:0}\n' + '.sec-rule{border:none;border-top:1px solid #e2e8f0;margin:0 0 16px}\n' + 'h3{font-size:13.5px;font-weight:700;color:#334155;margin:22px 0 10px}\n' + 'h4{font-size:12.5px;font-weight:700;color:#475569;margin:14px 0 8px}\n' + 'p.caption{font-size:12px;color:#64748b;margin:0 0 12px;font-style:italic;font-family:system-ui,sans-serif;line-height:1.5}\n' + 'p.no-data{font-size:12.5px;color:#94a3b8;font-style:italic;font-family:system-ui,sans-serif;margin:8px 0}\n' + 'table.meta-tbl{width:100%;border-collapse:collapse;font-size:12.5px;margin-bottom:16px;font-family:system-ui,sans-serif}\n' + 'table.meta-tbl th,table.meta-tbl td{padding:7px 12px;border:1px solid #e2e8f0;vertical-align:top}\n' + 'table.meta-tbl th{background:#f8fafc;font-weight:600;color:#475569;width:38%;text-align:left;font-size:11px;text-transform:uppercase;letter-spacing:.04em}\n' + 'table.meta-tbl .r-sec td{background:#1e293b;color:#f1f5f9;font-weight:600;font-size:11px;letter-spacing:.05em;text-transform:uppercase;padding:5px 12px;border-color:#1e293b}\n' + 'table.data-tbl{width:100%;border-collapse:collapse;font-size:12px;margin-bottom:16px;font-family:system-ui,sans-serif}\n' + 'table.data-tbl th{background:#f1f5f9;font-weight:600;color:#374151;padding:6px 10px;border:1px solid #e2e8f0;text-align:left;font-size:11px;text-transform:uppercase;letter-spacing:.04em}\n' + 'table.data-tbl td{padding:5px 10px;border:1px solid #e2e8f0;color:#1a1a2e}\n' + 'table.qc-tbl{width:100%;border-collapse:collapse;font-size:12.5px;margin-bottom:16px;font-family:system-ui,sans-serif}\n' + 'table.qc-tbl th{background:#f8fafc;font-weight:600;color:#475569;padding:6px 12px;border:1px solid #e2e8f0;text-align:left;font-size:11px;text-transform:uppercase;letter-spacing:.04em}\n' + 'table.qc-tbl td{padding:7px 12px;border:1px solid #e2e8f0;vertical-align:top;font-size:12.5px}\n' + 'figure{margin:16px 0 22px}\n' + 'figure img{display:block;max-width:100%;height:auto;border:1px solid #e2e8f0;border-radius:4px}\n' + 'figcaption{font-size:11.5px;color:#64748b;margin-top:8px;font-style:italic;font-family:system-ui,sans-serif;line-height:1.5}\n' + 'ol.refs{padding-left:22px;font-size:12px;line-height:1.75;color:#334155;font-family:system-ui,sans-serif}\n' + 'ol.refs li{margin-bottom:10px}\n' + '.footer{margin-top:48px;padding-top:16px;border-top:1px solid #e2e8f0;font-size:11px;color:#94a3b8;font-family:system-ui,sans-serif;line-height:1.5}\n' + '@media print{\n' + '  .toolbar{display:none!important}\n' + '  body{background:#fff}\n' + '  .page-wrap{box-shadow:none;padding:30px 40px}\n' + '  h2{page-break-after:avoid}\n' + '  h3{page-break-after:avoid}\n' + '  figure{page-break-inside:avoid}\n' + '  table{page-break-inside:avoid}\n' + '}\n' + '</style>\n' + '</head>\n<body>\n' +

'<div class="toolbar">' +   '<div class="toolbar-title">Surfaceome Explorer — Analysis Report</div>' +   '<span class="toolbar-note">Use browser Print to save as PDF</span>' +   '<button class="tbtn tbtn-print" onclick="window.print()">🖨 Print / Save as PDF</button>' +   '<button class="tbtn tbtn-dl" onclick="dlHTML()">⬇ Download HTML</button>' + '</div>\n\n' +

'<div class="page-wrap">\n' +

'<div class="cover-title">Surfaceome Explorer v2</div>\n' + '<div class="cover-sub">Cell Surface Proteomics Analysis Report</div>\n' + '<div class="cover-meta">Generated: ' + dateStr + ' at ' + timeStr + '&nbsp;&nbsp;|&nbsp;&nbsp;File: <strong>' + esc(fname) + '</strong></div>\n' + '<hr class="cover-rule">\n\n' +

'<div class="sec-num">Section 1</div>\n' + '<h2 class="first">Dataset Summary</h2>\n<hr class="sec-rule">\n' + '<table class="meta-tbl">' +   rRow('File name', esc(fname)) +   rRow('File format', ftype) +   rRow('Proteins / features', nProts.toLocaleString()) +   rRow('Quantitative samples', nSamps.toLocaleString()) +   rRow('Protein ID column', esc(App.config.idCol)) +   (App.config.geneCol ? rRow('Gene symbol column', esc(App.config.geneCol)) : '') +   rRow('Experimental groups', groupList) + '</table>\n\n' +

'<div class="sec-num">Section 2</div>\n' + '<h2>Data Quality Control</h2>\n<hr class="sec-rule">\n' + '<p class="caption">Automatic quality checks performed after file upload (Tab 2 — Data Check).</p>\n' + '<table class="meta-tbl">' +   rRow('Overall status', qc.overallStatus === 'pass'     ? '<span style="color:#059669;font-weight:700">&#10003; Pass</span>'     : qc.overallStatus === 'warn'     ? '<span style="color:#d97706;font-weight:700">&#9888; Warning</span>'     : '<span style="color:#dc2626;font-weight:700">&#10007; Error</span>') +   rRow('Total proteins', (qc.nProteins||nProts).toLocaleString()) +   rRow('Total samples', (qc.nSamples||nSamps).toLocaleString()) +   rRow('Missing values', qc.nMissing != null ? qc.nMissing.toLocaleString() + ' (' + qc.missPct + '%)' : NA) +   rRow('Non-numeric values', qc.nNonNumeric != null ? qc.nNonNumeric.toLocaleString() : '0') +   rRow('Constant features', nConstant > 0 ? nConstant.toLocaleString() + ' proteins flagged' : 'None detected') +   rRow('Sample name changes', qc.nModified != null ? qc.nModified.toLocaleString() : '0') + '</table>\n' + (qcChecksHTML ? '<table class="qc-tbl"><thead><tr><th style="width:35%">Check</th><th style="width:8%;text-align:center">Status</th><th>Details</th></tr></thead><tbody>' + qcChecksHTML + '</tbody></table>\n' : '') +

'\n<div class="sec-num">Section 3</div>\n' + '<h2>Data Filtering</h2>\n<hr class="sec-rule">\n' + '<table class="meta-tbl">' + filtRows + '</table>\n\n' +

'<div class="sec-num">Section 4</div>\n' + '<h2>Missing Value Estimation</h2>\n<hr class="sec-rule">\n' + '<table class="meta-tbl">' + impRows + '</table>\n\n' +

'<div class="sec-num">Section 5</div>\n' + '<h2>Normalization</h2>\n<hr class="sec-rule">\n' + '<table class="meta-tbl">' + normRows + '</table>\n\n' +

'<div class="sec-num">Section 6</div>\n' + '<h2>Surfaceome Analysis</h2>\n<hr class="sec-rule">\n' + '<table class="meta-tbl">' +   rRow('Proteins analyzed', R.totalN.toLocaleString()) +   rRow('Surface-associated (any database)', R.nSurf.toLocaleString() + ' (' + pc(R.nSurf, R.totalN) + ')') +   rRow('Non-membrane proteins', R.nNone.toLocaleString() + ' (' + pc(R.nNone, R.totalN) + ')') +   rRow('UniProt Cell Membrane (KW-1003)', R.nPM.toLocaleString() + ' (' + pc(R.nPM, R.totalN) + ')') +   rRow('SURFY in-silico predicted', R.nSS.toLocaleString() + ' (' + pc(R.nSS, R.totalN) + ')') +   rRow('Membrane-Associated', R.nMA.toLocaleString() + ' (' + pc(R.nMA, R.totalN) + ')') +   rRow('Supported by ≥2 databases', highConf.toLocaleString() + ' (' + pc(highConf, R.totalN) + ')') +   rRow('All three databases', R.regions.allThree.length.toLocaleString() + ' (' + pc(R.regions.allThree.length, R.totalN) + ')') + '</table>\n' + (barImg ? '<figure><img src="' + barImg + '" alt="Surfaceome bar chart" style="max-width:540px"><figcaption>Figure 1. Surfaceome composition. Proteins classified by database annotation; a single protein may appear in multiple categories.</figcaption></figure>\n' : '') +

'\n<div class="sec-num">Section 7</div>\n' + '<h2>Database Overlap</h2>\n<hr class="sec-rule">\n' + (vennImg ? '<figure><img src="' + vennImg + '" alt="Venn diagram" style="max-width:500px"><figcaption>Figure 2. Three-database Venn diagram. Numbers indicate proteins unique to each region. SURFY (blue), UniProt Cell Membrane (teal), Membrane-Associated (purple).</figcaption></figure>\n' : '') + '<table class="data-tbl"><thead><tr><th>Region</th><th style="text-align:right">Proteins</th><th style="text-align:right">% of total</th></tr></thead><tbody>' + regionTableRows + '</tbody></table>\n\n' +

'<div class="sec-num">Section 8</div>\n' + '<h2>Most Abundant Cell Surface Proteins by Group</h2>\n<hr class="sec-rule">\n' + '<p class="caption">Top 20 UniProt Cell Membrane (KW-1003) proteins ranked by mean abundance per experimental group. Values are in input data units (e.g., log₂ LFQ intensity).</p>\n' + top20HTML + '\n\n' +

'<div class="sec-num">Section 9</div>\n' + '<h2>Differential Surfaceome Analysis</h2>\n<hr class="sec-rule">\n' + (R.comparisons.length > 0 ? '<p class="caption">Pairwise comparison of UniProt Cell Membrane proteins. Welch’s t-test with Benjamini–Hochberg FDR correction. Significance: |log₂FC| &gt; 1 and FDR &lt; 0.05.</p>\n' : '') + diffHTML + '\n\n' +

'<div class="sec-num">Section 10</div>\n' + '<h2>Condition-Specific Surface Proteins</h2>\n<hr class="sec-rule">\n' +
(R.condSpecific && R.condSpecific.length > 0
  ? '<p class="caption">Proteins detected in ≥50% of replicates in one group and &lt;50% in the other. Strong biomarker candidates.</p>\n'   : '') + condHTML + '\n\n' +

'<div class="sec-num">Section 11</div>\n' + '<h2>Analysis Parameters</h2>\n<hr class="sec-rule">\n' + '<p class="caption">Complete parameter record for computational reproducibility.</p>\n' + '<table class="meta-tbl">' + paramRows + '</table>\n\n' +

'<div class="sec-num">Section 12</div>\n' + '<h2>References</h2>\n<hr class="sec-rule">\n' + '<ol class="refs">\n' + '<li>Bausch-Fluck D, Goldmann U, Bramley G, et al. (2015). A mass spectrometric-derived cell surface protein atlas. <em>PLOS ONE</em> 10(3):e0121314. [SURFY]</li>\n' + '<li>Bausch-Fluck D, Milani ES, Wollscheid B. (2021). The in silico human surfaceome. <em>Proc Natl Acad Sci USA</em> 118(1):e2024870118.</li>\n' + '<li>UniProt Consortium. (2023). UniProt: the Universal Protein Knowledgebase in 2023. <em>Nucleic Acids Res</em> 51(D1):D523–D531. [KW-1003 Cell membrane]</li>\n' + '<li>Alm\xe9n MS, Nordstr\xf6m KJV, Fredriksson R, Schi\xf6th HB. (2009). Mapping the human membrane proteome. <em>BMC Biology</em> 7:50. [Membrane-associated set]</li>\n' + '<li>Fagerberg L, et al. (2010). Prediction of the human membrane proteome. <em>Proteomics</em> 10(6):1141–1149.</li>\n' + '<li>Benjamini Y, Hochberg Y. (1995). Controlling the false discovery rate. <em>J R Stat Soc Series B</em> 57(1):289–300.</li>\n' + '</ol>\n\n' +

'<div class="footer">Report generated by Surfaceome Explorer v2 &nbsp;|&nbsp; ' + dateStr + ' &nbsp;|&nbsp; All analysis performed client-side &mdash; no data transmitted to external servers.</div>\n' + '</div>\n\n' +

'<script>\n' + 'function dlHTML(){\n' + '  var h=document.documentElement.outerHTML;\n' + '  var b=new Blob([h],{type:"text/html"});\n' + '  var u=URL.createObjectURL(b);\n' + '  var a=document.createElement("a");\n' + '  a.href=u; a.download="surfaceome_report_' + ts + '.html";\n' + '  document.body.appendChild(a); a.click(); document.body.removeChild(a);\n' + '  setTimeout(function(){URL.revokeObjectURL(u)},2000);\n' + '}\n' + '<\/script>\n' + '</body>\n</html>';

  const blob = new Blob([html], { type: 'text/html' });
  const url  = URL.createObjectURL(blob);
  const win  = window.open(url, '_blank');
  if (!win) {
    showToast('Pop-up blocked — please allow pop-ups for this page and try again.');
    URL.revokeObjectURL(url);
  } else {
    setTimeout(() => URL.revokeObjectURL(url), 120000);
    showToast('Report opened in new tab — use Print / Save as PDF to export.');
  }
}

/* ── Report table row helper ── */
function _rRow(label, value) {
  if (!value && value !== 0) return '';   return '<tr><th>' + label + '</th><td>' + value + '</td></tr>';
}

function _t6CsvEsc(v) {
  const s = String(v??'');   return s.includes(',')||s.includes('"')||s.includes('\n') ? '"'+s.replace(/"/g,'""')+'"' : s;
}
function _t6DownloadCSV(csv, filename) {
  const blob = new Blob([csv], { type:'text/csv' });
  const url  = URL.createObjectURL(blob);
  const a    = document.createElement('a');
  a.href=url; a.download=filename;
  document.body.appendChild(a); a.click(); document.body.removeChild(a);
  setTimeout(()=>URL.revokeObjectURL(url), 1000);
}

/* ====== _v2_main.js ====== */
/* ==============================================================
   MAIN  —  Shared helpers, tab bar, initialization
   ============================================================== */

/* ── HTML escaping ── */
function esc(s) {
  return String(s || '').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
}

/* ── Toast notifications ── */
let _toastTimer;
function showToast(msg) {
  const el = document.getElementById('toast');
  el.textContent = msg;
  el.classList.add('show');
  clearTimeout(_toastTimer);
  _toastTimer = setTimeout(() => el.classList.remove('show'), 3000);
}

/* ── Tab bar rendering ── */
function renderTabBar() {
  const nav = document.getElementById('tab-nav');   let html = '<div class="tab-list">';
  for (let n = 1; n <= 6; n++) {
    const meta      = TAB_META[n];
    const isActive  = App.activeTab === n;
    const isUnlocked = App.tabUnlocked[n];
    const isDone    = isUnlocked && App.activeTab > n;

    let cls = 'tab-item';     if (isActive)    cls += ' tab-active';     else if (!isUnlocked) cls += ' tab-locked';     else if (isDone) cls += ' tab-done';

    const click  = isUnlocked ? ' onclick="goTab(' + n + ')"' : '';     const numTxt = isDone ? '&#x2713;' : String(n);
    const lbl    = meta.short;

    html += '<div class="' + cls + '"' + click + '>' +       '<span class="tab-num">' + numTxt + '</span>' +       '<span class="tab-lbl">' + esc(lbl) + '</span>' +     '</div>';

    if (n < 6) {
      html += '<div class="tab-conn' + (isUnlocked && App.tabUnlocked[n+1] ? ' tab-conn-done' : '') + '"></div>';
    }
  }
  html += '</div>';
  nav.innerHTML = html;
}

/* ── Tab content dispatcher ── */
function renderTabContent(n) {
  const fns = { 1:renderTab1, 2:renderTab2, 3:renderTab3, 4:renderTab4, 5:renderTab5, 6:renderTab6 };
  if (fns[n]) fns[n]();
}

/* ── Initialize on DOMContentLoaded ── */
document.addEventListener('DOMContentLoaded', () => {
  renderTabBar();
  renderTabContent(1);
});

</script>
</body>
</html>
