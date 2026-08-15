/************************************************************************
 * PQA RS 日報 (Remote Sensing)  Code.gs
 * 置き場所: UAVスプレッドシート の Apps Script（同じブックでOK）
 * UAV Akuisisi / FlightLog / Perbandingan を集計 → 日報テキスト＋Laporan RS
 ************************************************************************/

var UAV_SHEET_ID = '1VPtbrU8ayY9KQByZTPg_GDbBbPTB4bdoWLB64XIX-Bk';       // 読込元(UAVの現場データ)
var LAPORAN_SHEET_ID = '1UegGIVVUtHaJ-5eOFZt1dgkrgccq_uBODXWZ3wNbz4g';   // 保存先(RS日報の別ブック)
var SR_SHEET_ID  = '1tRluw2z4TbNh1G_Wejb2SMI6wQiheNaq7m14g4c-bwI'; // Wilayah/Staff(当面)
var WIL_NAMA = {'1':'Suban Jeriji','2':'Bunakat','3':'Lematang'};
var LAP_RS = 'Laporan RS';
var LAP_RS_HEADER = ['Tanggal','Wilayah','Penanggung Jawab','HOK','U02','U06','U12','U24','U36','Total',
  'Total Target','Realisasi Kumulatif','Sisa Target','Jam Mulai','Jam Selesai','Catatan','Note','Blocks','Teks','Received At'];

function doGet(){
  return HtmlService.createTemplateFromFile('index').evaluate()
    .setTitle('Laporan Harian Remote Sensing')
    .addMetaTag('viewport','width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/*==================== マスター ====================*/
function getMasters(){
  var ss=SpreadsheetApp.openById(SR_SHEET_ID);
  var wilayah=readCol_(ss,'Wilayah'), staffByWil={}, svByWil={};
  wilayah.forEach(function(w){ var n=wilNum_(w); if(!n) return; var st=readStaff_(ss,'W'+n+' Staff'); staffByWil[n]=st.all; svByWil[n]=st.sv; });
  return { wilayah:wilayah, staffByWil:staffByWil, svByWil:svByWil, wilNama:WIL_NAMA };
}
function readCol_(ss,name){ var sh=ss.getSheetByName(name); if(!sh||sh.getLastRow()<2) return [];
  var v=sh.getRange(2,1,sh.getLastRow()-1,1).getValues(),o=[]; for(var i=0;i<v.length;i++){var s=String(v[i][0]).trim(); if(s)o.push(s);} return o; }
function readStaff_(ss,name){ var sh=ss.getSheetByName(name); if(!sh||sh.getLastRow()<2) return {all:[],sv:[]};
  var lc=Math.min(Math.max(sh.getLastColumn(),1),2); var v=sh.getRange(2,1,sh.getLastRow()-1,lc).getValues(),all=[],sv=[];
  for(var i=0;i<v.length;i++){ var nm=String(v[i][0]).trim(); if(!nm)continue; all.push(nm);
    if(lc>=2 && String(v[i][1]).toUpperCase().indexOf('SV')>=0) sv.push(nm); } return {all:all,sv:sv}; }

/*==================== 集計 ====================*/
function hmap_(sh){ var lc=Math.max(sh.getLastColumn(),1), h=sh.getRange(1,1,1,lc).getValues()[0], m={};
  for(var i=0;i<h.length;i++){ var k=String(h[i]).trim(); if(k && !(k in m)) m[k]=i; } return m; }
function col_(m,names){ for(var i=0;i<names.length;i++) if(names[i] in m) return m[names[i]]; return -1; }

function getLaporanRS(tanggal, wilayah){
  var ss=SpreadsheetApp.openById(UAV_SHEET_ID), wn=wilNum_(wilayah);
  var umur={'02':0,'06':0,'12':0,'24':0,'36':0}, total=0, blocks={}, akuIds={};

  var aku=ss.getSheetByName('Akuisisi');
  if(aku && aku.getLastRow()>1){
    var m=hmap_(aku), v=aku.getRange(2,1,aku.getLastRow()-1,aku.getLastColumn()).getValues();
    var ci={id:col_(m,['ID']),tgl:col_(m,['Tanggal Akuisisi']),wil:col_(m,['Wilayah']),umur:col_(m,['Umur (Bulan)']),
      blok:col_(m,['Blok']),compt:col_(m,['Compt']),stand:col_(m,['Stand']),luas:col_(m,['Luas (Ha)']),
      pilot:col_(m,['Pilot']),copilot:col_(m,['Co Pilot']),perb:col_(m,['Perbandingan Strip']),jml:col_(m,['Jumlah Strip'])};
    for(var r=0;r<v.length;r++){ var row=v[r];
      if(fdate_(g_(row,ci.tgl))!==fdate_(tanggal)) continue;
      if(wilNum_(g_(row,ci.wil))!==wn) continue;
      var um=umurNorm_(g_(row,ci.umur)), luas=num_(g_(row,ci.luas));
      var blok=String(g_(row,ci.blok)||'').trim(), compt=String(g_(row,ci.compt)||'').trim(), stand=String(g_(row,ci.stand)||'').trim();
      total+=luas; if(umur[um]!=null) umur[um]+=luas;
      akuIds[String(g_(row,ci.id))]=1;
      var key=um+'|'+blok+'|'+compt+'|'+stand;
      if(!blocks[key]) blocks[key]={umur:um,blok:blok,compt:compt,stand:stand,luas:0,
        pilot:String(g_(row,ci.pilot)||''),copilot:String(g_(row,ci.copilot)||''),
        perbStrip:String(g_(row,ci.perb)||''),jumlahStrip:g_(row,ci.jml),perbAdd:[]};
      blocks[key].luas+=luas;
    }
  }

  // Jam kerja: FlightLog の該当Akuisisi
  var jm='', js='';
  var fl=ss.getSheetByName('FlightLog');
  if(fl && fl.getLastRow()>1){
    var fm=hmap_(fl), fv=fl.getRange(2,1,fl.getLastRow()-1,fl.getLastColumn()).getValues();
    var fi={id:col_(fm,['Akuisisi ID']),to:col_(fm,['Take Off Time']),ld:col_(fm,['Landing Time'])};
    for(var k=0;k<fv.length;k++){ if(!akuIds[String(g_(fv[k],fi.id))]) continue;
      var to=String(g_(fv[k],fi.to)||'').trim(), ld=String(g_(fv[k],fi.ld)||'').trim();
      if(to && (jm===''||to<jm)) jm=to; if(ld && (js===''||ld>js)) js=ld; }
  }

  // Perbandingan（2Strip・ドローン無し）: 同Umur×Blok×Stand の block に strip を追記、無ければ別リスト
  var perbList=[];
  var pb=ss.getSheetByName('Perbandingan');
  if(pb && pb.getLastRow()>1){
    var pm=hmap_(pb), pv=pb.getRange(2,1,pb.getLastRow()-1,pb.getLastColumn()).getValues();
    var pi={tgl:col_(pm,['Tanggal Audit']),wil:col_(pm,['Wilayah']),umur:col_(pm,['Jenis Audit/Umur']),
      blok:col_(pm,['Blok']),stand:col_(pm,['Comp/Stand']),nostrip:col_(pm,['No Strip'])};
    for(var q=0;q<pv.length;q++){ var pr=pv[q];
      if(fdate_(g_(pr,pi.tgl))!==fdate_(tanggal)) continue;
      if(wilNum_(g_(pr,pi.wil))!==wn) continue;
      var pum=umurNorm_(g_(pr,pi.umur)), pblok=String(g_(pr,pi.blok)||'').trim(), pstand=String(g_(pr,pi.stand)||'').trim(), pno=String(g_(pr,pi.nostrip)||'').trim();
      var matched=false;
      Object.keys(blocks).forEach(function(kk){ var b=blocks[kk];
        if(b.umur===pum && b.blok===pblok && b.stand===pstand){ b.perbAdd.push(pno||'(strip)'); matched=true; } });
      if(!matched) perbList.push({umur:pum,blok:pblok,stand:pstand,nostrip:pno});
    }
  }

  var blockArr=Object.keys(blocks).map(function(k){ var b=blocks[k]; b.luas=round2_(b.luas); return b; });
  return { tanggal:fdate_(tanggal), wilayah:'Wilayah '+wn, wilNama:(WIL_NAMA[wn]||''),
    umur:{'02':round2_(umur['02']),'06':round2_(umur['06']),'12':round2_(umur['12']),'24':round2_(umur['24']),'36':round2_(umur['36'])},
    total:round2_(total), jamMulai:jm, jamSelesai:js, blocks:blockArr, perbList:perbList };
}
function g_(row,i){ return i>=0 ? row[i] : ''; }

/*==================== Laporan RS 保存 ====================*/
function getLapRS_(){ var ss=SpreadsheetApp.openById(LAPORAN_SHEET_ID), sh=ss.getSheetByName(LAP_RS);
  if(!sh){ sh=ss.insertSheet(LAP_RS); sh.getRange(1,1,1,LAP_RS_HEADER.length).setValues([LAP_RS_HEADER]).setFontWeight('bold'); sh.setFrozenRows(1); return sh; }
  var lc=Math.max(sh.getLastColumn(),1); if(lc<LAP_RS_HEADER.length) sh.getRange(1,lc+1,1,LAP_RS_HEADER.length-lc).setValues([LAP_RS_HEADER.slice(lc)]).setFontWeight('bold');
  return sh; }
function setupLaporanRS(){ getLapRS_(); return 'Laporan RS siap'; }
function simpanRS(p){
  var lock=LockService.getScriptLock(); lock.waitLock(30000);
  try{ var sh=getLapRS_();
    var row=[ p.tanggal,p.wilayah,p.pj,p.hok,p.u02,p.u06,p.u12,p.u24,p.u36,p.total,
      p.totalTarget,p.realisasi,p.sisa,p.jamMulai,p.jamSelesai,p.catatan,p.note,p.blocks,p.teks,new Date() ];
    var wn=wilNum_(p.wilayah), tg=fdate_(p.tanggal), found=-1;
    if(sh.getLastRow()>1){ var ex=sh.getRange(2,1,sh.getLastRow()-1,2).getValues();
      for(var i=0;i<ex.length;i++){ if(fdate_(ex[i][0])===tg && wilNum_(ex[i][1])===wn){ found=i+2; break; } } }
    if(found>0) sh.getRange(found,1,1,LAP_RS_HEADER.length).setValues([row]); else sh.appendRow(row);
    return { ok:true, updated:(found>0) };
  } finally { lock.releaseLock(); }
}

/*==================== util ====================*/
function wilNum_(w){ return String(w).replace(/[^0-9]/g,''); }
function umurNorm_(u){ var n=String(u).replace(/[^0-9]/g,''); return !n?'':(n.length<2?'0'+n:n.substring(0,2)); }
function num_(v){ var x=parseFloat(String(v).replace(',','.')); return isFinite(x)?x:0; }
function round2_(x){ return Math.round(x*100)/100; }
function fdate_(v){ if(v instanceof Date) return Utilities.formatDate(v,Session.getScriptTimeZone(),'yyyy-MM-dd');
  var s=String(v).trim(); return s.length>=10?s.substring(0,10):s; }
