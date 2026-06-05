import { useState, useCallback, useMemo, useRef } from "react";
import * as XLSX from "xlsx";

/* ═══════════════════════════════════════════════════════════════
   ENGINE — TMI-2019, PSG/UPMK, PUC, Sensitivity, Export
   (tidak berubah dari versi sebelumnya, sudah sesuai PDF)
   ═══════════════════════════════════════════════════════════════ */
const TMI_2019=[{age:20,male:.0004,female:.0003},{age:25,male:.00052,female:.00038},{age:30,male:.00075,female:.00056},{age:35,male:.00107,female:.0008},{age:40,male:.00173,female:.00118},{age:45,male:.00302,female:.00187},{age:50,male:.00508,female:.00305},{age:55,male:.00789,female:.00483},{age:60,male:.011,female:.007},{age:65,male:.016,female:.0105}];
function getMortalityRate(age,gender="male"){const t=TMI_2019;if(age<=t[0].age)return t[0][gender];if(age>=t[t.length-1].age)return t[t.length-1][gender];for(let i=0;i<t.length-1;i++){if(age>=t[i].age&&age<t[i+1].age){const f=(age-t[i].age)/(t[i+1].age-t[i].age);return t[i][gender]+f*(t[i+1][gender]-t[i][gender])}}return .005}
function getPSG(mk){if(mk<1)return 1;if(mk<2)return 2;if(mk<3)return 3;if(mk<4)return 4;if(mk<5)return 5;if(mk<6)return 6;if(mk<7)return 7;if(mk<8)return 8;return 9}
function getUPMK(mk){if(mk<3)return 0;if(mk<6)return 2;if(mk<9)return 3;if(mk<12)return 4;if(mk<15)return 5;if(mk<18)return 6;if(mk<21)return 7;if(mk<24)return 8;return 10}
function calculatePUC(emp,a){const{discountRate:r,salaryGrowth:g,turnoverRate:q,retirementAge}=a;const age=emp.age,mk=emp.masaKerja,gaji=emp.gajiTahunan,N=Math.max(0,retirementAge-age);if(N<=0)return{...emp,N,projGaji:gaji,projBenefit:0,discFactor:1,probAktif:1,pbo:0,csc:0,interestCost:0,mkLaluKoreksi:mk,cap24:0,psg:0,upmk:0,totalMKPensiun:mk,status:"SUDAH PENSIUN"};const totalMK=mk+N,projGaji=gaji*Math.pow(1+g,N),upahBulanan=projGaji/12,psg=getPSG(totalMK),upmk=getUPMK(totalMK),benefitUUCK=(1.75*psg+1*upmk)*upahBulanan;let projBenefit=0;if(emp.tipeBenefit==="Gratuity"||emp.tipeBenefit==="UUCK")projBenefit=benefitUUCK;else if(emp.tipeBenefit==="Pension")projBenefit=.02*upahBulanan*totalMK;else projBenefit=benefitUUCK+.02*upahBulanan*totalMK;const discFactor=1/Math.pow(1+r,N),qMort=getMortalityRate(age,emp.gender==="Wanita"?"female":"male"),qDis=qMort*.1,qTotal=Math.min(1,q+qMort+qDis),probAktif=Math.pow(1-qTotal,N),cap24=Math.min(totalMK,24),mkLaluKoreksi=totalMK>24?Math.max(0,24-N):mk,pbo=cap24>0?projBenefit*discFactor*probAktif*(mkLaluKoreksi/cap24):0,csc=cap24>0?(projBenefit/cap24)*discFactor*probAktif:0,interestCost=pbo*r;return{...emp,N,projGaji,projBenefit,discFactor,probAktif,pbo,csc,interestCost,mkLaluKoreksi,cap24,psg,upmk,totalMKPensiun:totalMK,status:"Aktif"}}
function runSensitivity(employees,ba){const base=employees.map(e=>calculatePUC(e,ba)),basePBO=base.reduce((s,e)=>s+e.pbo,0);return[{label:"Base Case",...ba},{label:"Discount Rate +1%",...ba,discountRate:ba.discountRate+.01},{label:"Discount Rate −1%",...ba,discountRate:ba.discountRate-.01},{label:"Kenaikan Gaji +0.5%",...ba,salaryGrowth:ba.salaryGrowth+.005},{label:"Kenaikan Gaji −0.5%",...ba,salaryGrowth:ba.salaryGrowth-.005},{label:"Turnover +2%",...ba,turnoverRate:ba.turnoverRate+.02},{label:"Turnover −2%",...ba,turnoverRate:Math.max(0,ba.turnoverRate-.02)}].map(sc=>{const res=employees.map(e=>calculatePUC(e,sc)),totalPBO=res.reduce((s,e)=>s+e.pbo,0),totalCSC=res.reduce((s,e)=>s+e.csc,0),diff=basePBO>0?(totalPBO-basePBO)/basePBO:0;return{label:sc.label,r:sc.discountRate,g:sc.salaryGrowth,q:sc.turnoverRate,totalPBO,totalCSC,diff}})}
function downloadTemplate(){const wb=XLSX.utils.book_new();const headers=["Nama","Usia","Masa Kerja","Gaji Tahunan","Tipe Benefit","Gender"];const samples=[["Budi Santoso",35,8,72000000,"UUCK","Pria"],["Siti Rahayu",42,15,96000000,"Pension","Wanita"],["Ahmad Fauzi",28,3,54000000,"Both","Pria"],["Dewi Lestari",50,22,120000000,"UUCK","Wanita"],["Rudi Hartono",38,10,84000000,"UUCK","Pria"]];const ws=XLSX.utils.aoa_to_sheet([["TEMPLATE IMPORT DATA KARYAWAN — KALKULATOR IMBALAN PASCA KERJA"],["Petunjuk: Isi data mulai baris ke-4. Jangan ubah nama kolom. Tipe Benefit: UUCK / Pension / Both. Gender: Pria / Wanita."],[],headers,...samples]);ws["!cols"]=headers.map(h=>({wch:Math.max(h.length+4,20)}));XLSX.utils.book_append_sheet(wb,ws,"Template_Karyawan");XLSX.writeFile(wb,"Template_Import_Karyawan.xlsx")}
function exportToExcel(results,assumptions,sensitivity){const wb=XLSX.utils.book_new();const ws1=XLSX.utils.aoa_to_sheet([["KALKULASI IMBALAN PASCA KERJA — METODE PUC"],["PSAK 24 / IAS 19 | UUCK/PP35 | Siaran Pers DSAK IAI April 2022"],[],["ASUMSI AKTUARIA"],["Discount Rate",assumptions.discountRate],["Salary Growth Rate",assumptions.salaryGrowth],["Turnover Rate",assumptions.turnoverRate],["Usia Pensiun Normal",assumptions.retirementAge],["Tabel Mortalitas","TMI-2019"],["Tanggal Penilaian",assumptions.tanggalPenilaian||new Date().toISOString().split("T")[0]],[]]);XLSX.utils.book_append_sheet(wb,ws1,"01_ASUMSI");const h=["No","Nama","Usia","Gender","Masa Kerja (n)","Thn ke Pensiun (N)","Gaji Tahunan","Proyeksi Gaji","PSG (bln)","UPMK (bln)","Proyeksi Benefit","Faktor Diskonto","Prob. Aktif","MK Lalu Koreksi","Cap 24","PBO / DBO","Current Service Cost","Interest Cost","Tipe Benefit","Status"];const rows=results.map((r,i)=>[i+1,r.nama,r.age,r.gender||"Pria",r.masaKerja,r.N,r.gajiTahunan,Math.round(r.projGaji),r.psg,r.upmk,Math.round(r.projBenefit),r.discFactor.toFixed(8),r.probAktif.toFixed(8),r.mkLaluKoreksi,r.cap24,Math.round(r.pbo),Math.round(r.csc),Math.round(r.interestCost),r.tipeBenefit,r.status]);const tot=["","TOTAL","","","","",results.reduce((s,r)=>s+r.gajiTahunan,0),"","","","","","","","",Math.round(results.reduce((s,r)=>s+r.pbo,0)),Math.round(results.reduce((s,r)=>s+r.csc,0)),Math.round(results.reduce((s,r)=>s+r.interestCost,0)),"",""];const ws2=XLSX.utils.aoa_to_sheet([["KALKULASI PUC DETAIL PER KARYAWAN"],[],h,...rows,[],tot]);ws2["!cols"]=h.map(x=>({wch:Math.max(x.length+2,14)}));XLSX.utils.book_append_sheet(wb,ws2,"02_KALKULASI_PUC");const tPBO=results.reduce((s,r)=>s+r.pbo,0),tCSC=results.reduce((s,r)=>s+r.csc,0),tIC=results.reduce((s,r)=>s+r.interestCost,0);const ws3=XLSX.utils.aoa_to_sheet([["RINGKASAN EKSEKUTIF — IMBALAN PASCA KERJA"],[],["Jumlah Karyawan",results.length],["Total Gaji Tahunan",results.reduce((s,r)=>s+r.gajiTahunan,0)],["Rata-rata Masa Kerja",(results.reduce((s,r)=>s+r.masaKerja,0)/results.length).toFixed(1)],[],["KOMPONEN BEBAN"],["Current Service Cost",Math.round(tCSC)],["Interest Cost",Math.round(tIC)],["Total Beban P&L",Math.round(tCSC+tIC)],[],["POSISI LIABILITAS"],["Defined Benefit Obligation (DBO)",Math.round(tPBO)],["(-) Fair Value Plan Assets",0],["NET LIABILITAS",Math.round(tPBO)]]);XLSX.utils.book_append_sheet(wb,ws3,"03_RINGKASAN");const sh=["Skenario","Discount Rate","Salary Growth","Turnover","Total PBO","Total CSC","PBO vs Base (%)"];const sr=sensitivity.map(s=>[s.label,(s.r*100).toFixed(2)+"%",(s.g*100).toFixed(2)+"%",(s.q*100).toFixed(2)+"%",Math.round(s.totalPBO),Math.round(s.totalCSC),s.label==="Base Case"?"—":(s.diff*100).toFixed(2)+"%"]);const ws4=XLSX.utils.aoa_to_sheet([["SENSITIVITY ANALYSIS"],[],sh,...sr]);ws4["!cols"]=sh.map(x=>({wch:Math.max(x.length+2,16)}));XLSX.utils.book_append_sheet(wb,ws4,"04_SENSITIVITAS");XLSX.writeFile(wb,`PUC_Imbalan_Pasca_Kerja_${new Date().toISOString().slice(0,10)}.xlsx`)}

const fmtRp=n=>new Intl.NumberFormat("id-ID",{style:"currency",currency:"IDR",maximumFractionDigits:0}).format(n);
const fmtPct=n=>(n*100).toFixed(2)+"%";
const fmtNum=n=>new Intl.NumberFormat("id-ID").format(Math.round(n));

/* ═══════════════════════════════════════════════════════════════
   THEME & CSS VARIABLES
   ═══════════════════════════════════════════════════════════════ */
const themes = {
  light: {
    "--bg-page":"#f5f6f8","--bg-card":"#ffffff","--bg-input":"#f0f2f5","--bg-header":"#ffffff",
    "--bg-table-head":"#f8f9fb","--bg-table-stripe":"#fafbfc","--bg-hover":"#f0f4ff",
    "--border":"#e2e5ea","--border-input":"#d0d5dd","--text-primary":"#1a1f36","--text-secondary":"#4e5d78",
    "--text-muted":"#8492a6","--text-on-accent":"#ffffff","--accent":"#2563eb","--accent-hover":"#1d4ed8",
    "--accent-light":"#eff4ff","--success":"#16a34a","--success-bg":"#f0fdf4","--danger":"#dc2626",
    "--danger-bg":"#fef2f2","--warning":"#d97706","--warning-bg":"#fffbeb",
    "--shadow-sm":"0 1px 2px rgba(0,0,0,0.05)","--shadow-md":"0 2px 8px rgba(0,0,0,0.08)",
  },
  dark: {
    "--bg-page":"#0f1117","--bg-card":"#1a1d2e","--bg-input":"#252839","--bg-header":"#161829",
    "--bg-table-head":"#1e2133","--bg-table-stripe":"#1c1f30","--bg-hover":"#252a40",
    "--border":"#2a2e42","--border-input":"#353952","--text-primary":"#e8eaf0","--text-secondary":"#a0a8c0",
    "--text-muted":"#636b85","--text-on-accent":"#ffffff","--accent":"#4f8ffa","--accent-hover":"#3b7cf5",
    "--accent-light":"#1e2a4a","--success":"#34d399","--success-bg":"#0d2e24","--danger":"#f87171",
    "--danger-bg":"#2e1515","--warning":"#fbbf24","--warning-bg":"#2e2610",
    "--shadow-sm":"0 1px 3px rgba(0,0,0,0.3)","--shadow-md":"0 2px 12px rgba(0,0,0,0.4)",
  },
};

const css = `
  @import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Sans:wght@400;500;600;700&family=IBM+Plex+Mono:wght@400;500&display=swap');
  *, *::before, *::after { box-sizing:border-box; margin:0; padding:0; }
  :root { font-family:'IBM Plex Sans',system-ui,sans-serif; }
  body { background:var(--bg-page); color:var(--text-primary); transition:background .25s,color .25s; }
  .mono { font-family:'IBM Plex Mono',monospace; }
  .puc-app { max-width:1100px; margin:0 auto; padding:0 16px 40px; }
  
  /* Header */
  .puc-header { background:var(--bg-header); border-bottom:1px solid var(--border); padding:14px 0; position:sticky; top:0; z-index:50; }
  .puc-header-inner { max-width:1100px; margin:0 auto; padding:0 16px; display:flex; align-items:center; justify-content:space-between; }
  .puc-logo { display:flex; align-items:center; gap:10px; }
  .puc-logo-icon { width:36px; height:36px; background:var(--accent); border-radius:8px; display:flex; align-items:center; justify-content:center; color:#fff; font-weight:700; font-size:15px; }
  .puc-logo h1 { font-size:16px; font-weight:700; color:var(--text-primary); line-height:1.2; }
  .puc-logo p { font-size:11px; color:var(--text-muted); }
  .theme-toggle { background:var(--bg-input); border:1px solid var(--border); border-radius:8px; padding:7px 12px; cursor:pointer; color:var(--text-secondary); font-size:13px; font-family:inherit; display:flex; align-items:center; gap:6px; transition:all .15s; }
  .theme-toggle:hover { border-color:var(--accent); color:var(--accent); }

  /* Steps nav */
  .steps { display:flex; gap:2px; margin:20px 0 18px; background:var(--bg-input); border-radius:10px; padding:3px; border:1px solid var(--border); }
  .step-btn { flex:1; padding:10px 12px; border:none; border-radius:8px; cursor:pointer; font-family:inherit; font-size:13px; font-weight:600; transition:all .15s; background:transparent; color:var(--text-muted); display:flex; align-items:center; justify-content:center; gap:6px; }
  .step-btn.active { background:var(--accent); color:var(--text-on-accent); box-shadow:var(--shadow-sm); }
  .step-btn:not(.active):hover { color:var(--text-primary); background:var(--bg-hover); }
  .step-num { width:22px; height:22px; border-radius:50%; display:inline-flex; align-items:center; justify-content:center; font-size:11px; font-weight:700; background:var(--bg-card); color:var(--text-muted); border:1.5px solid var(--border); }
  .step-btn.active .step-num { background:rgba(255,255,255,0.2); color:#fff; border-color:transparent; }
  .step-done .step-num { background:var(--success); color:#fff; border-color:transparent; }
  .badge { font-size:10px; background:var(--success); color:#fff; padding:1px 7px; border-radius:10px; font-weight:700; margin-left:4px; }

  /* Cards */
  .card { background:var(--bg-card); border:1px solid var(--border); border-radius:10px; padding:20px; margin-bottom:14px; box-shadow:var(--shadow-sm); }
  .card-title { font-size:13px; font-weight:700; color:var(--accent); margin-bottom:14px; display:flex; align-items:center; gap:8px; text-transform:uppercase; letter-spacing:.03em; }
  .card-title svg { flex-shrink:0; }
  .card-subtitle { font-size:12px; color:var(--text-muted); margin-bottom:16px; margin-top:-8px; }

  /* Forms */
  .field { display:flex; flex-direction:column; gap:4px; }
  .field label { font-size:11px; font-weight:600; color:var(--text-secondary); text-transform:uppercase; letter-spacing:.04em; }
  .field input, .field select { background:var(--bg-input); border:1px solid var(--border-input); border-radius:6px; padding:8px 10px; font-size:14px; color:var(--text-primary); font-family:inherit; outline:none; transition:border .15s; width:100%; }
  .field input:focus, .field select:focus { border-color:var(--accent); box-shadow:0 0 0 2px var(--accent-light); }
  .field input[type=number] { font-family:'IBM Plex Mono',monospace; }
  .field-row { display:grid; grid-template-columns:repeat(auto-fit,minmax(170px,1fr)); gap:12px; }

  /* Mode tabs */
  .mode-tabs { display:flex; gap:8px; margin-bottom:14px; }
  .mode-tab { flex:1; padding:12px 14px; border:1.5px solid var(--border); border-radius:8px; cursor:pointer; background:var(--bg-card); text-align:left; font-family:inherit; transition:all .15s; }
  .mode-tab.active { border-color:var(--accent); background:var(--accent-light); }
  .mode-tab:hover:not(.active) { border-color:var(--text-muted); }
  .mode-tab strong { font-size:13px; display:block; color:var(--text-primary); margin-bottom:2px; }
  .mode-tab span { font-size:11px; color:var(--text-muted); }

  /* Table */
  .tbl-wrap { overflow-x:auto; border:1px solid var(--border); border-radius:8px; }
  table.puc-tbl { width:100%; border-collapse:collapse; font-size:12px; }
  .puc-tbl th { padding:9px 8px; text-align:left; font-weight:600; color:var(--text-muted); background:var(--bg-table-head); border-bottom:1px solid var(--border); white-space:nowrap; font-size:11px; text-transform:uppercase; letter-spacing:.04em; }
  .puc-tbl td { padding:7px 8px; border-bottom:1px solid var(--border); color:var(--text-primary); }
  .puc-tbl tr:nth-child(even) td { background:var(--bg-table-stripe); }
  .puc-tbl tr:hover td { background:var(--bg-hover); }
  .puc-tbl .r { text-align:right; }
  .puc-tbl .mono { font-family:'IBM Plex Mono',monospace; font-size:12px; }
  .puc-tbl tfoot td { font-weight:700; background:var(--bg-table-head)!important; }
  .puc-tbl input, .puc-tbl select { background:var(--bg-input); border:1px solid var(--border); border-radius:4px; padding:5px 6px; font-size:12px; color:var(--text-primary); outline:none; font-family:inherit; }
  .puc-tbl input[type=number] { font-family:'IBM Plex Mono',monospace; }
  .puc-tbl input:focus, .puc-tbl select:focus { border-color:var(--accent); }

  /* Summary cards */
  .summary-grid { display:grid; grid-template-columns:repeat(auto-fit,minmax(220px,1fr)); gap:12px; margin-bottom:18px; }
  .summary-card { background:var(--bg-card); border:1px solid var(--border); border-radius:10px; padding:16px 18px; box-shadow:var(--shadow-sm); border-left:3px solid var(--accent); }
  .summary-card .label { font-size:11px; color:var(--text-muted); font-weight:600; text-transform:uppercase; letter-spacing:.04em; margin-bottom:6px; }
  .summary-card .value { font-size:20px; font-weight:700; font-family:'IBM Plex Mono',monospace; color:var(--text-primary); }

  /* Buttons */
  .btn { padding:10px 18px; border:none; border-radius:8px; font-family:inherit; font-size:13px; font-weight:600; cursor:pointer; display:inline-flex; align-items:center; gap:6px; transition:all .15s; }
  .btn-primary { background:var(--accent); color:var(--text-on-accent); box-shadow:var(--shadow-sm); }
  .btn-primary:hover { background:var(--accent-hover); }
  .btn-outline { background:transparent; border:1.5px solid var(--border); color:var(--text-secondary); }
  .btn-outline:hover { border-color:var(--accent); color:var(--accent); }
  .btn-success { background:var(--success); color:#fff; }
  .btn-success:hover { opacity:.9; }
  .btn-danger-ghost { background:none; border:none; color:var(--danger); cursor:pointer; padding:4px; font-size:12px; }
  .btn-lg { padding:14px 28px; font-size:15px; border-radius:10px; width:100%; justify-content:center; }
  .btn-dashed { background:transparent; border:2px dashed var(--border); color:var(--text-muted); width:100%; justify-content:center; border-radius:8px; padding:10px; }
  .btn-dashed:hover { border-color:var(--accent); color:var(--accent); }

  /* Status pill */
  .pill { font-size:10px; padding:2px 8px; border-radius:4px; font-weight:600; display:inline-block; }
  .pill-active { background:var(--success-bg); color:var(--success); }
  .pill-retired { background:var(--danger-bg); color:var(--danger); }

  /* Sensitivity bar */
  .sens-bar-track { background:var(--bg-input); border-radius:4px; height:8px; overflow:hidden; }
  .sens-bar-fill { height:100%; border-radius:4px; transition:width .5s ease; }

  /* Employee card (individual mode) */
  .emp-card { background:var(--bg-card); border:1px solid var(--border); border-radius:10px; padding:16px; margin-bottom:10px; box-shadow:var(--shadow-sm); }
  .emp-card-head { display:flex; justify-content:space-between; align-items:center; margin-bottom:12px; }
  .emp-card-num { font-size:13px; font-weight:700; color:var(--accent); }

  /* Empty state */
  .empty { text-align:center; padding:50px 20px; color:var(--text-muted); }
  .empty-icon { font-size:40px; margin-bottom:10px; }
  .empty h3 { font-size:16px; font-weight:600; color:var(--text-secondary); margin-bottom:4px; }

  /* Footer */
  .puc-footer { margin-top:30px; padding-top:16px; border-top:1px solid var(--border); text-align:center; }
  .puc-footer p { font-size:11px; color:var(--text-muted); margin-bottom:3px; }

  /* Responsive */
  @media(max-width:640px){
    .field-row{grid-template-columns:1fr 1fr}
    .mode-tabs{flex-direction:column}
    .summary-grid{grid-template-columns:1fr}
    .steps{flex-wrap:wrap}
    .step-btn{font-size:12px;padding:8px}
  }
`;

/* ═══════════════════════════════════════════════════════════════
   ICONS (Lucide-style inline SVGs)
   ═══════════════════════════════════════════════════════════════ */
const Ic=({d,size=16})=><svg width={size} height={size} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d={d}/></svg>;
const IcSun=()=><Ic d="M12 1v2M12 21v2M4.22 4.22l1.42 1.42M18.36 18.36l1.42 1.42M1 12h2M21 12h2M4.22 19.78l1.42-1.42M18.36 5.64l1.42-1.42M12 7a5 5 0 100 10 5 5 0 000-10z"/>;
const IcMoon=()=><Ic d="M21 12.79A9 9 0 1111.21 3 7 7 0 0021 12.79z"/>;
const IcUser=()=><><svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M19 21v-2a4 4 0 00-4-4H9a4 4 0 00-4 4v2"/><circle cx="12" cy="7" r="4"/></svg></>;
const IcUsers=()=><svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M17 21v-2a4 4 0 00-4-4H5a4 4 0 00-4 4v2"/><circle cx="9" cy="7" r="4"/><path d="M23 21v-2a4 4 0 00-3-3.87M16 3.13a4 4 0 010 7.75"/></svg>;
const IcChart=()=><svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M18 20V10M12 20V4M6 20v-6"/></svg>;
const IcDownload=()=><svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M21 15v4a2 2 0 01-2 2H5a2 2 0 01-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" y1="15" x2="12" y2="3"/></svg>;
const IcPlus=()=><svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><line x1="12" y1="5" x2="12" y2="19"/><line x1="5" y1="12" x2="19" y2="12"/></svg>;
const IcTrash=()=><svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><polyline points="3 6 5 6 21 6"/><path d="M19 6v14a2 2 0 01-2 2H7a2 2 0 01-2-2V6m3 0V4a2 2 0 012-2h4a2 2 0 012 2v2"/></svg>;
const IcSettings=()=><svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><circle cx="12" cy="12" r="3"/><path d="M19.4 15a1.65 1.65 0 00.33 1.82l.06.06a2 2 0 010 2.83 2 2 0 01-2.83 0l-.06-.06a1.65 1.65 0 00-1.82-.33 1.65 1.65 0 00-1 1.51V21a2 2 0 01-4 0v-.09A1.65 1.65 0 009 19.4a1.65 1.65 0 00-1.82.33l-.06.06a2 2 0 01-2.83-2.83l.06-.06A1.65 1.65 0 004.68 15a1.65 1.65 0 00-1.51-1H3a2 2 0 010-4h.09A1.65 1.65 0 004.6 9a1.65 1.65 0 00-.33-1.82l-.06-.06a2 2 0 012.83-2.83l.06.06A1.65 1.65 0 009 4.68a1.65 1.65 0 001-1.51V3a2 2 0 014 0v.09a1.65 1.65 0 001 1.51 1.65 1.65 0 001.82-.33l.06-.06a2 2 0 012.83 2.83l-.06.06A1.65 1.65 0 0019.4 9a1.65 1.65 0 001.51 1H21a2 2 0 010 4h-.09a1.65 1.65 0 00-1.51 1z"/></svg>;
const IcUpload=()=><svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M21 15v4a2 2 0 01-2 2H5a2 2 0 01-2-2v-4"/><polyline points="17 8 12 3 7 8"/><line x1="12" y1="3" x2="12" y2="15"/></svg>;
const IcCalc=()=><svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><rect x="4" y="2" width="16" height="20" rx="2"/><line x1="8" y1="6" x2="16" y2="6"/><line x1="8" y1="10" x2="10" y2="10"/><line x1="14" y1="10" x2="16" y2="10"/><line x1="8" y1="14" x2="10" y2="14"/><line x1="14" y1="14" x2="16" y2="14"/><line x1="8" y1="18" x2="10" y2="18"/><line x1="14" y1="18" x2="16" y2="18"/></svg>;

/* ═══════════════════════════════════════════════════════════════
   DEFAULTS & OPTIONS
   ═══════════════════════════════════════════════════════════════ */
const DEFAULT_ASSUMPTIONS={discountRate:.065,salaryGrowth:.035,turnoverRate:.05,retirementAge:56,tanggalPenilaian:"2024-12-31"};
const EMPTY_EMP={nama:"",age:35,masaKerja:5,gajiTahunan:60000000,tipeBenefit:"UUCK",gender:"Pria"};
const BENEFIT_OPTS=[{value:"UUCK",label:"UUCK/PP35 (1.75×PSG + UPMK)"},{value:"Pension",label:"Pensiun Imbalan Pasti (2%×Gaji×MK)"},{value:"Both",label:"Keduanya (UUCK + Pensiun)"}];

/* ═══════════════════════════════════════════════════════════════
   MAIN COMPONENT
   ═══════════════════════════════════════════════════════════════ */
export default function PUCCalculator(){
  const[theme,setTheme]=useState(()=>{try{return localStorage.getItem("puc-theme")||"light"}catch{return"light"}});
  const[mode,setMode]=useState("individual");
  const[assumptions,setAssumptions]=useState(DEFAULT_ASSUMPTIONS);
  const[employees,setEmployees]=useState([{...EMPTY_EMP,id:1}]);
  const[summaryInput,setSummaryInput]=useState({jumlahKaryawan:50,rataUsia:38,rataMasaKerja:10,rataGajiTahunan:84000000,tipeBenefit:"UUCK"});
  const[results,setResults]=useState(null);
  const[sensitivity,setSensitivity]=useState(null);
  const[step,setStep]=useState("input");
  const fileRef=useRef(null);

  const toggleTheme=()=>{const next=theme==="light"?"dark":"light";setTheme(next);try{localStorage.setItem("puc-theme",next)}catch{}};

  // Apply theme CSS vars
  const themeVars=themes[theme];
  const rootStyle=Object.entries(themeVars).reduce((o,[k,v])=>{o[k]=v;return o},{});

  const addEmployee=()=>setEmployees(p=>[...p,{...EMPTY_EMP,id:Date.now()}]);
  const removeEmployee=id=>setEmployees(p=>p.filter(e=>e.id!==id));
  const updateEmployee=(id,f,v)=>setEmployees(p=>p.map(e=>e.id===id?{...e,[f]:v}:e));

  const generateFromSummary=()=>{
    const{jumlahKaryawan:jk,rataUsia:ru,rataMasaKerja:rmk,rataGajiTahunan:rg,tipeBenefit:tb}=summaryInput;
    const gen=[];
    for(let i=0;i<jk;i++){
      const aV=Math.round((Math.random()-.5)*16),mV=Math.round((Math.random()-.5)*10),sV=(Math.random()-.5)*.4;
      const ea=Math.max(21,Math.min(59,ru+aV)),em=Math.max(1,Math.min(ea-18,rmk+mV)),eg=Math.round(rg*(1+sV)/1e6)*1e6;
      gen.push({id:Date.now()+i,nama:`Karyawan ${String(i+1).padStart(3,"0")}`,age:ea,masaKerja:em,gajiTahunan:Math.max(12e6,eg),tipeBenefit:tb,gender:Math.random()>.5?"Pria":"Wanita"});
    }
    setEmployees(gen);setMode("bulk");
  };

const handleFileImport = (e) => {
    const file = e.target.files[0];
    if (!file) return;
    e.target.value = "";

    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const wb = XLSX.read(new Uint8Array(evt.target.result), { type: "array" });

        // Pilih sheet: Template_Karyawan > 02_KALKULASI_PUC > sheet pertama
        let sheetName = wb.SheetNames[0];
        if (wb.SheetNames.includes("Template_Karyawan")) sheetName = "Template_Karyawan";
        else if (wb.SheetNames.includes("02_KALKULASI_PUC")) sheetName = "02_KALKULASI_PUC";

        // Baca sebagai array-of-arrays agar bisa mendeteksi baris header secara dinamis
        const rawRows = XLSX.utils.sheet_to_json(wb.Sheets[sheetName], { header: 1, defval: "" });

        // Kata kunci yang dikenali sebagai nama kolom (sudah di-normalize)
        const KNOWN = ["nama","name","namakaryawan","usia","age","umur","masakerja","masakerjan","mk","masakerjathn","gajitahunan","gaji","salary","upah","tipebenefit","benefit","tipe","gender","jeniskelamin","kelamin"];

        // Cari baris yang mengandung ≥ 3 kolom yang dikenali → itu baris header
        let headerIdx = -1;
        for (let i = 0; i < rawRows.length; i++) {
          const norm = rawRows[i].map(c => String(c).toLowerCase().replace(/[^a-z0-9]/g, ""));
          if (norm.filter(k => KNOWN.includes(k)).length >= 3) { headerIdx = i; break; }
        }

        if (headerIdx === -1) {
          alert("Format file tidak dikenali.\nPastikan file menggunakan Template yang sudah di-download, atau file hasil Export dari kalkulator ini.");
          return;
        }

        // Buat peta kolom: keyword → index kolom
        const headerRow = rawRows[headerIdx].map(c => String(c).toLowerCase().replace(/[^a-z0-9]/g, ""));
        const colIdx = (keywords) => { for (const k of keywords) { const i = headerRow.indexOf(k); if (i !== -1) return i; } return -1; };

        const iNama    = colIdx(["nama","name","namakaryawan"]);
        const iAge     = colIdx(["usia","age","umur"]);
        const iMK      = colIdx(["masakerja","masakerjan","mk","masakerjathn"]);
        const iGaji    = colIdx(["gajitahunan","gaji","salary","upah"]);
        const iBenefit = colIdx(["tipebenefit","benefit","tipe"]);
        const iGender  = colIdx(["gender","jeniskelamin","kelamin"]);

        const imp = rawRows.slice(headerIdx + 1).map((row, i) => {
          const nama  = iNama    >= 0 ? String(row[iNama] ?? "").trim() : "";
          const age   = iAge     >= 0 ? Number(row[iAge])   : NaN;
          const mk    = iMK      >= 0 ? Number(row[iMK])    : NaN;
          const gaji  = iGaji    >= 0 ? Number(row[iGaji])  : NaN;
          const tipe  = iBenefit >= 0 ? String(row[iBenefit] ?? "").trim() : "";
          const gen   = iGender  >= 0 ? String(row[iGender]  ?? "").trim() : "";

          // Lewati baris kosong dan baris TOTAL
          if (!nama && isNaN(age) && isNaN(gaji)) return null;
          if (nama.toUpperCase().includes("TOTAL")) return null;

          return {
            id: Date.now() + i,
            nama: nama || `Karyawan ${i + 1}`,
            age: isNaN(age) || age <= 0 ? 35 : Math.round(age),
            masaKerja: isNaN(mk) || mk < 0 ? 0 : Math.round(mk),
            gajiTahunan: isNaN(gaji) || gaji <= 0 ? 60000000 : Math.round(gaji),
            tipeBenefit: ["UUCK","Pension","Both"].includes(tipe) ? tipe : "UUCK",
            gender: ["Pria","Wanita"].includes(gen) ? gen : "Pria",
          };
        }).filter(Boolean);

        if (imp.length === 0) {
          alert("Tidak ada data karyawan yang ditemukan.\nPastikan ada baris data di bawah baris header kolom.");
          return;
        }

        setEmployees(imp);
        setMode("bulk");
      } catch {
        alert("Gagal membaca file. Pastikan file tidak rusak dan berformat .xlsx / .xls / .csv.");
      }
    };
    reader.readAsArrayBuffer(file);
  };

  const calculate=useCallback(()=>{
    const res=employees.map(e=>calculatePUC(e,assumptions));
    const sens=runSensitivity(employees,assumptions);
    setResults(res);setSensitivity(sens);setStep("results");
  },[employees,assumptions]);

  const totals=useMemo(()=>{
    if(!results)return null;
    return{pbo:results.reduce((s,r)=>s+r.pbo,0),csc:results.reduce((s,r)=>s+r.csc,0),ic:results.reduce((s,r)=>s+r.interestCost,0),gaji:results.reduce((s,r)=>s+r.gajiTahunan,0),count:results.length};
  },[results]);

  /* ═══ RENDER ═══ */
  return(
    <div style={rootStyle}>
      <style>{css}</style>

      {/* HEADER */}
      <header className="puc-header">
        <div className="puc-header-inner">
          <div className="puc-logo">
            <div className="puc-logo-icon">PUC</div>
            <div>
              <h1>Kalkulator Imbalan Pasca Kerja</h1>
              <p>PSAK 24 · UUCK/PP35 · TMI-2019 · Siaran Pers DSAK IAI 2022</p>
            </div>
          </div>
          <button className="theme-toggle" onClick={toggleTheme}>
            {theme==="light"?<IcMoon/>:<IcSun/>}
            {theme==="light"?"Gelap":"Terang"}
          </button>
        </div>
      </header>

      <div className="puc-app">
        {/* STEP NAVIGATION */}
        <nav className="steps">
          {[{key:"input",n:"1",label:"Input Data",icon:<IcSettings/>},
            {key:"results",n:"2",label:"Hasil Kalkulasi",icon:<IcChart/>},
            {key:"sensitivity",n:"3",label:"Sensitivitas",icon:<IcCalc/>}
          ].map(s=>(
            <button key={s.key} className={`step-btn ${step===s.key?"active":""} ${results&&s.key==="input"?"step-done":""}`} onClick={()=>setStep(s.key)}>
              <span className="step-num">{results&&s.key==="input"?"✓":s.n}</span>
              {s.label}
              {s.key==="results"&&results&&<span className="badge">{results.length}</span>}
            </button>
          ))}
        </nav>

        {/* ═══ STEP 1: INPUT ═══ */}
        {step==="input"&&(<>
          {/* Asumsi */}
          <div className="card">
            <div className="card-title"><IcSettings/> Asumsi Aktuaria</div>
            <p className="card-subtitle">Parameter yang mempengaruhi seluruh perhitungan. Pastikan sesuai dengan yield SBN dan data perusahaan.</p>
            <div className="field-row">
              {[{k:"discountRate",l:"Discount Rate (%)",m:100,s:.1},{k:"salaryGrowth",l:"Kenaikan Gaji (%)",m:100,s:.1},{k:"turnoverRate",l:"Turnover Rate (%)",m:100,s:.1},{k:"retirementAge",l:"Usia Pensiun (thn)",m:1,s:1}].map(f=>(
                <div className="field" key={f.k}>
                  <label>{f.l}</label>
                  <input type="number" step={f.s} value={Math.round(assumptions[f.k]*f.m*100)/100} onChange={e=>setAssumptions({...assumptions,[f.k]:Number(e.target.value)/f.m})}/>
                </div>
              ))}
              <div className="field">
                <label>Tanggal Penilaian</label>
                <input type="date" value={assumptions.tanggalPenilaian} onChange={e=>setAssumptions({...assumptions,tanggalPenilaian:e.target.value})}/>
              </div>
            </div>
          </div>

          {/* Mode Tabs */}
          <div className="mode-tabs">
            {[{k:"individual",t:"Input Individu",d:"Isi data satu per satu",ic:<IcUser/>},{k:"bulk",t:"Input Massal",d:"Import Excel/CSV atau edit tabel",ic:<IcUsers/>},{k:"summary",t:"Input Rata-rata",d:"Generate dari data ringkasan",ic:<IcChart/>}].map(m=>(
              <button key={m.k} className={`mode-tab ${mode===m.k?"active":""}`} onClick={()=>setMode(m.k)}>
                <strong style={{display:"flex",alignItems:"center",gap:6}}>{m.ic}{m.t}</strong>
                <span>{m.d}</span>
              </button>
            ))}
          </div>

          {/* MODE: SUMMARY */}
          {mode==="summary"&&(
            <div className="card">
              <div className="card-title"><IcChart/> Input Rata-rata Karyawan</div>
              <p className="card-subtitle">Masukkan data rata-rata perusahaan. Sistem akan men-generate data individu dengan variasi acak yang wajar.</p>
              <div className="field-row">
                {[{k:"jumlahKaryawan",l:"Jumlah Karyawan"},{k:"rataUsia",l:"Rata-rata Usia (thn)"},{k:"rataMasaKerja",l:"Rata-rata Masa Kerja (thn)"},{k:"rataGajiTahunan",l:"Rata-rata Gaji Tahunan (Rp)"}].map(f=>(
                  <div className="field" key={f.k}>
                    <label>{f.l}</label>
                    <input type="number" value={summaryInput[f.k]} onChange={e=>setSummaryInput({...summaryInput,[f.k]:Number(e.target.value)})}/>
                  </div>
                ))}
                <div className="field">
                  <label>Tipe Benefit</label>
                  <select value={summaryInput.tipeBenefit} onChange={e=>setSummaryInput({...summaryInput,tipeBenefit:e.target.value})}>
                    {BENEFIT_OPTS.map(o=><option key={o.value} value={o.value}>{o.label}</option>)}
                  </select>
                </div>
              </div>
              <div style={{marginTop:16}}>
                <button className="btn btn-primary" onClick={generateFromSummary}>
                  <IcUsers/> Generate {summaryInput.jumlahKaryawan} Karyawan
                </button>
              </div>
            </div>
          )}

          {/* MODE: BULK */}
          {mode==="bulk"&&(
            <div className="card">
              <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:14}}>
                <div className="card-title" style={{marginBottom:0}}><IcUsers/> Data Karyawan ({employees.length} orang)</div>
                <div style={{display:"flex",gap:8}}>
                  <button className="btn btn-outline" onClick={downloadTemplate}><IcDownload/> Download Template</button>
                  <button className="btn btn-outline" onClick={()=>fileRef.current?.click()}><IcUpload/> Import Excel/CSV</button>
                  <input ref={fileRef} type="file" accept=".xlsx,.xls,.csv" onChange={handleFileImport} style={{display:"none"}}/>
                  <button className="btn btn-primary" onClick={addEmployee}><IcPlus/> Tambah</button>
                </div>
              </div>
              <div className="tbl-wrap" style={{maxHeight:420,overflowY:"auto"}}>
                <table className="puc-tbl">
                  <thead><tr>
                    {["#","Nama","Usia","MK (thn)","Gaji Tahunan","Tipe Benefit","Gender",""].map((h,i)=><th key={i}>{h}</th>)}
                  </tr></thead>
                  <tbody>{employees.map((emp,i)=>(
                    <tr key={emp.id}>
                      <td style={{color:"var(--text-muted)"}}>{i+1}</td>
                      <td><input type="text" value={emp.nama} onChange={e=>updateEmployee(emp.id,"nama",e.target.value)} style={{width:130}}/></td>
                      <td><input type="number" value={emp.age} onChange={e=>updateEmployee(emp.id,"age",+e.target.value)} style={{width:55}}/></td>
                      <td><input type="number" value={emp.masaKerja} onChange={e=>updateEmployee(emp.id,"masaKerja",+e.target.value)} style={{width:55}}/></td>
                      <td><input type="number" value={emp.gajiTahunan} onChange={e=>updateEmployee(emp.id,"gajiTahunan",+e.target.value)} style={{width:110}}/></td>
                      <td><select value={emp.tipeBenefit} onChange={e=>updateEmployee(emp.id,"tipeBenefit",e.target.value)} style={{width:90}}>{BENEFIT_OPTS.map(o=><option key={o.value} value={o.value}>{o.value}</option>)}</select></td>
                      <td><select value={emp.gender} onChange={e=>updateEmployee(emp.id,"gender",e.target.value)} style={{width:65}}><option value="Pria">Pria</option><option value="Wanita">Wanita</option></select></td>
                      <td><button className="btn-danger-ghost" onClick={()=>removeEmployee(emp.id)}><IcTrash/></button></td>
                    </tr>
                  ))}</tbody>
                </table>
              </div>
            </div>
          )}

          {/* MODE: INDIVIDUAL */}
          {mode==="individual"&&(<>
            {employees.map((emp,i)=>(
              <div className="emp-card" key={emp.id}>
                <div className="emp-card-head">
                  <span className="emp-card-num">Karyawan #{i+1}</span>
                  {employees.length>1&&<button className="btn-danger-ghost" onClick={()=>removeEmployee(emp.id)} style={{display:"flex",alignItems:"center",gap:4}}><IcTrash/> Hapus</button>}
                </div>
                <div className="field-row">
                  {[{f:"nama",l:"Nama Karyawan",t:"text"},{f:"age",l:"Usia (tahun)",t:"number"},{f:"masaKerja",l:"Masa Kerja (tahun)",t:"number"},{f:"gajiTahunan",l:"Gaji Tahunan (Rp)",t:"number"}].map(c=>(
                    <div className="field" key={c.f}>
                      <label>{c.l}</label>
                      <input type={c.t} value={emp[c.f]} onChange={e=>updateEmployee(emp.id,c.f,c.t==="number"?+e.target.value:e.target.value)}/>
                    </div>
                  ))}
                  <div className="field">
                    <label>Tipe Benefit</label>
                    <select value={emp.tipeBenefit} onChange={e=>updateEmployee(emp.id,"tipeBenefit",e.target.value)}>
                      {BENEFIT_OPTS.map(o=><option key={o.value} value={o.value}>{o.label}</option>)}
                    </select>
                  </div>
                  <div className="field">
                    <label>Gender</label>
                    <select value={emp.gender} onChange={e=>updateEmployee(emp.id,"gender",e.target.value)}>
                      <option value="Pria">Pria</option><option value="Wanita">Wanita</option>
                    </select>
                  </div>
                </div>
              </div>
            ))}
            <button className="btn btn-dashed" onClick={addEmployee}><IcPlus/> Tambah Karyawan</button>
          </>)}

          {/* Calculate */}
          <div style={{marginTop:18}}>
            <button className="btn btn-primary btn-lg" onClick={calculate}>
              <IcCalc/> Hitung Imbalan Pasca Kerja ({employees.length} Karyawan)
            </button>
          </div>
        </>)}

        {/* ═══ STEP 2: RESULTS ═══ */}
        {step==="results"&&results&&(<>
          <div className="summary-grid">
            {[{l:"Total DBO / PBO",v:fmtRp(totals.pbo),c:"var(--accent)"},{l:"Current Service Cost",v:fmtRp(totals.csc),c:"#0891b2"},{l:"Interest Cost",v:fmtRp(totals.ic),c:"#7c3aed"},{l:"Total Beban Laba Rugi",v:fmtRp(totals.csc+totals.ic),c:"var(--success)"}].map((c,i)=>(
              <div className="summary-card" key={i} style={{borderLeftColor:c.c}}>
                <div className="label">{c.l}</div>
                <div className="value" style={{color:c.c}}>{c.v}</div>
              </div>
            ))}
          </div>

          <div className="card" style={{padding:0,overflow:"hidden"}}>
            <div style={{padding:"16px 20px 0"}}>
              <div className="card-title"><IcChart/> Detail Kalkulasi Per Karyawan</div>
            </div>
            <div className="tbl-wrap" style={{border:"none",borderRadius:0}}>
              <table className="puc-tbl">
                <thead><tr>
                  {["#","Nama","Usia","MK","N","Gaji Tahunan","PSG","UPMK","Proyeksi Benefit","PBO / DBO","CSC","Interest Cost","Status"].map((h,i)=>(
                    <th key={i} className={i>=5?"r":""}>{h}</th>
                  ))}
                </tr></thead>
                <tbody>{results.map((r,i)=>(
                  <tr key={i}>
                    <td style={{color:"var(--text-muted)"}}>{i+1}</td>
                    <td style={{fontWeight:500,maxWidth:130,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{r.nama}</td>
                    <td>{r.age}</td><td>{r.masaKerja}</td><td>{r.N}</td>
                    <td className="r mono">{fmtNum(r.gajiTahunan)}</td>
                    <td className="r">{r.psg}</td><td className="r">{r.upmk}</td>
                    <td className="r mono">{fmtNum(r.projBenefit)}</td>
                    <td className="r mono" style={{fontWeight:700,color:"var(--accent)"}}>{fmtNum(r.pbo)}</td>
                    <td className="r mono" style={{color:"#0891b2"}}>{fmtNum(r.csc)}</td>
                    <td className="r mono" style={{color:"#7c3aed"}}>{fmtNum(r.interestCost)}</td>
                    <td><span className={`pill ${r.status==="Aktif"?"pill-active":"pill-retired"}`}>{r.status}</span></td>
                  </tr>
                ))}</tbody>
                <tfoot><tr>
                  <td colSpan={5}>TOTAL ({results.length} karyawan)</td>
                  <td className="r mono">{fmtNum(totals.gaji)}</td>
                  <td colSpan={3}></td>
                  <td className="r mono" style={{color:"var(--accent)"}}>{fmtNum(totals.pbo)}</td>
                  <td className="r mono" style={{color:"#0891b2"}}>{fmtNum(totals.csc)}</td>
                  <td className="r mono" style={{color:"#7c3aed"}}>{fmtNum(totals.ic)}</td>
                  <td></td>
                </tr></tfoot>
              </table>
            </div>
          </div>

          <button className="btn btn-success btn-lg" onClick={()=>exportToExcel(results,assumptions,sensitivity)}>
            <IcDownload/> Export ke Excel (.xlsx)
          </button>
        </>)}

        {/* ═══ STEP 3: SENSITIVITY ═══ */}
        {step==="sensitivity"&&sensitivity&&(<>
          <div className="card">
            <div className="card-title"><IcCalc/> Analisis Sensitivitas Asumsi</div>
            <p className="card-subtitle">Dampak perubahan asumsi terhadap total PBO/DBO — sesuai persyaratan pengungkapan PSAK 24.</p>
            <div className="tbl-wrap">
              <table className="puc-tbl">
                <thead><tr>
                  {["Skenario","Disc. Rate","Kenaikan Gaji","Turnover","Total PBO","Total CSC","Δ PBO vs Base"].map((h,i)=>(
                    <th key={i} className={i>=4?"r":""}>{h}</th>
                  ))}
                </tr></thead>
                <tbody>{sensitivity.map((s,i)=>(
                  <tr key={i} style={i===0?{fontWeight:700}:{}}>
                    <td>{s.label}</td>
                    <td className="mono">{fmtPct(s.r)}</td><td className="mono">{fmtPct(s.g)}</td><td className="mono">{fmtPct(s.q)}</td>
                    <td className="r mono" style={{fontWeight:600}}>{fmtRp(s.totalPBO)}</td>
                    <td className="r mono">{fmtRp(s.totalCSC)}</td>
                    <td className="r mono" style={{fontWeight:600,color:i===0?"var(--text-muted)":s.diff>0?"var(--danger)":"var(--success)"}}>
                      {i===0?"—":(s.diff>0?"+":"")+fmtPct(s.diff)}
                    </td>
                  </tr>
                ))}</tbody>
              </table>
            </div>
          </div>

          <div className="card">
            <div className="card-title"><IcChart/> Visualisasi Dampak</div>
            {sensitivity.slice(1).map((s,i)=>{
              const maxAbs=Math.max(...sensitivity.slice(1).map(x=>Math.abs(x.diff)));
              const w=maxAbs>0?Math.abs(s.diff)/maxAbs*100:0;
              const neg=s.diff<0;
              return(<div key={i} style={{marginBottom:12}}>
                <div style={{display:"flex",justifyContent:"space-between",fontSize:13,marginBottom:5}}>
                  <span style={{color:"var(--text-secondary)"}}>{s.label}</span>
                  <span className="mono" style={{fontWeight:600,color:neg?"var(--success)":"var(--danger)"}}>
                    {(s.diff>0?"+":"")+fmtPct(s.diff)}
                  </span>
                </div>
                <div className="sens-bar-track">
                  <div className="sens-bar-fill" style={{width:`${w}%`,background:neg?"var(--success)":"var(--danger)"}}/>
                </div>
              </div>);
            })}
          </div>

          <button className="btn btn-success btn-lg" onClick={()=>exportToExcel(results,assumptions,sensitivity)}>
            <IcDownload/> Export ke Excel (.xlsx)
          </button>
        </>)}

        {/* EMPTY */}
        {(step==="results"||step==="sensitivity")&&!results&&(
          <div className="empty">
            <div className="empty-icon">📊</div>
            <h3>Belum Ada Hasil</h3>
            <p>Isi data karyawan di langkah 1, lalu klik "Hitung Imbalan Pasca Kerja"</p>
          </div>
        )}

        {/* FOOTER */}
        <footer className="puc-footer">
          <p>Kalkulator PUC · PSAK 24 / IAS 19 · Formula UUCK/PP35 · TMI-2019 · Siaran Pers DSAK IAI April 2022</p>
          <p>⚠ Ini adalah tool estimasi — untuk pelaporan final wajib menggunakan laporan aktuaris independen tersertifikasi</p>
        </footer>
      </div>
    </div>
  );
}
