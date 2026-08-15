/************************************************************************
 * PQA UAV アプリ  Code.gs  (v1)
 * 置き場所: 新スプレッドシート(下記 UAV_SHEET_ID)の Apps Script
 * 完全オフラインHTMLアプリ(index) → doPost で保存
 *   Home: Akuisisi / Validasi(無効) / Perbandingan
 ************************************************************************/

var UAV_SHEET_ID = '1VPtbrU8ayY9KQByZTPg_GDbBbPTB4bdoWLB64XIX-Bk';
var SR_SHEET_ID  = '1tRluw2z4TbNh1G_Wejb2SMI6wQiheNaq7m14g4c-bwI'; // Wilayah/Staff の読込元(当面)
var MASTER_SHEET_ID = UAV_SHEET_ID; // 将来ここに Wilayah/W1 Staff... を作ったら SR_SHEET_ID から切替
var WIL_NAMA = {'1':'Suban Jeriji','2':'Bunakat','3':'Lematang'};

var APP_USER='pqa', APP_PASS='pqa2026', AUTH_SECRET='uav-rahasia-2026', TOKEN_DAYS=7;

var SHEET_AKU='Akuisisi', SHEET_FLIGHT='FlightLog', SHEET_PERB='Perbandingan';

var AKU_HEADER = ['ID','Received At','ClientID','Wilayah','Unit','Blok','Compt','Stand','Luas (Ha)','Umur (Bulan)','Spesies','Jenis Bibit','UAV ID','Tanggal Akuisisi',
  'Status RTH','Altitude RTH','Flight Altitude','Gimbal Lock','Gimbal Angel','Body Check','Flight Mode','Resolution','Subtitle','Gridlines','Battery Cycle','Voltage','Obstacle Avoidance','Satellite Count','Safety Signal Loss',
  'Keterangan','PQA SV','Assessor 1','Assessor 2','Assessor 3','Updated At'];
var PREP_KEYS = ['statusRTH','altitudeRTH','flightAltitude','gimbalLock','gimbalAngel','bodyCheck','flightMode','resolution','subtitle','gridlines','batteryCycle','voltage','obstacleAvoidance','satelliteCount','safetySignalLoss'];

var FLIGHT_HEADER = ['Akuisisi ID','Flight No','Record Flight Name','Take Off Time','Landing Time','Duration','Distance (m)',
  'Battery Take Off %','Battery Landing %','ID Battery',
  'Coord X (DD)','Coord Y (DD)','Coord X (UTM)','Coord Y (UTM)',
  'Connection Mission','Connection Landing','Weather Take-off','Weather Mission','Weather Landing','UAV Cond Take Off','UAV Cond Landing','Remarks','Created At'];

var PERB_HEADER = (function(){
  var h=['ID','Received At','ClientID','Jenis Audit/Umur','Wilayah','Unit','Blok','Comp/Stand','Luas (Ha)','Species','No Strip','Halaman','Dari','Tanggal Audit'];
  for(var i=1;i<=16;i++) h.push('B1-'+i);
  for(var j=1;j<=16;j++) h.push('B2-'+j);
  h=h.concat(['Start X (DD)','Start Y (DD)','Start X (UTM)','Start Y (UTM)','End X (DD)','End Y (DD)','End X (UTM)','End Y (UTM)',
    'PQA SV','Assessor 1','Assessor 2','Assessor 3','Updated At']);
  return h;
})();

/*==================== エントリ ====================*/
function doGet(){
  var t = HtmlService.createTemplateFromFile('index');
  t.EXEC_URL = ScriptApp.getService().getUrl() || '';
  return t.evaluate().setTitle('PQA UAV')
    .addMetaTag('viewport','width=device-width, initial-scale=1, maximum-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
function doPost(e){
  var out;
  try{
    var req = JSON.parse(e.postData.contents);
    if(req.action==='login'){
      if(req.username===APP_USER && req.password===APP_PASS) out={ok:true, token:makeToken_(), masters:getMasters_()};
      else out={ok:false, error:'invalid'};
    } else if(!checkToken_(req.token)){
      out={ok:false, error:'unauthorized'};
    } else if(req.action==='masters'){
      out={ok:true, masters:getMasters_()};
    } else if(req.action==='akuisisi'){
      out=saveAkuisisi_(req.payload)||{ok:true};
    } else if(req.action==='perbandingan'){
      out=savePerbandingan_(req.payload)||{ok:true};
    } else if(req.action==='ping'){
      out={ok:true};
    } else { out={ok:false, error:'unknown_action'}; }
  }catch(err){ out={ok:false, error:'server_error', detail:String(err)}; }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

/*==================== トークン ====================*/
function makeToken_(){ var payload=Utilities.base64EncodeWebSafe(JSON.stringify({u:APP_USER,exp:Date.now()+TOKEN_DAYS*86400000}));
  return payload+'.'+sign_(payload); }
function sign_(s){ return Utilities.base64EncodeWebSafe(Utilities.computeHmacSha256Signature(s,AUTH_SECRET)); }
function checkToken_(token){ if(!token||token.indexOf('.')<0) return false; var p=token.split('.'); if(p.length!==2||sign_(p[0])!==p[1]) return false;
  var d; try{ d=JSON.parse(Utilities.newBlob(Utilities.base64DecodeWebSafe(p[0])).getDataAsString()); }catch(e){ return false; } return d.exp && Date.now()<=d.exp; }

/*==================== マスター(Wilayah / Staff / SV) ====================*/
function getMasters_(){
  var ss = SpreadsheetApp.openById(SR_SHEET_ID);
  var wilayah = readCol_(ss,'Wilayah');
  var byWil={};
  wilayah.forEach(function(w){ var n=String(w).replace(/[^0-9]/g,''); if(!n) return; var st=readStaff_(ss,'W'+n+' Staff'); byWil[n]={staff:st.all, sv:st.sv}; });
  return { wilayah:wilayah, byWil:byWil, wilNama:WIL_NAMA };
}
function readCol_(ss,name){ var sh=ss.getSheetByName(name); if(!sh||sh.getLastRow()<2) return [];
  var v=sh.getRange(2,1,sh.getLastRow()-1,1).getValues(),out=[]; for(var i=0;i<v.length;i++){var s=String(v[i][0]).trim(); if(s)out.push(s);} return out; }
function readStaff_(ss,name){ var sh=ss.getSheetByName(name); if(!sh||sh.getLastRow()<2) return {all:[],sv:[]};
  var lc=Math.min(Math.max(sh.getLastColumn(),1),2); var v=sh.getRange(2,1,sh.getLastRow()-1,lc).getValues(),all=[],sv=[];
  for(var i=0;i<v.length;i++){ var nm=String(v[i][0]).trim(); if(!nm)continue; all.push(nm);
    if(lc>=2 && String(v[i][1]).toUpperCase().indexOf('SV')>=0) sv.push(nm); } return {all:all,sv:sv}; }

/*==================== 保存: Akuisisi + FlightLog ====================*/
function saveAkuisisi_(d){
  if(!d) return {ok:false,error:'empty'};
  var lock=LockService.getScriptLock(); lock.waitLock(30000);
  try{
    var ss=SpreadsheetApp.openById(UAV_SHEET_ID);
    var aku=ensureSheet_(ss,SHEET_AKU,AKU_HEADER);
    if(d.clientId && clientIdExists_(aku,3,d.clientId)) return {ok:true,dedup:true};
    var now=new Date();
    var id=Utilities.formatDate(now,Session.getScriptTimeZone(),'yyMMddHHmmss')+'-'+Math.floor(Math.random()*1e4);
    var p=d.prep||{}, a=d.assessors||[];
    var row=[ id, now, d.clientId||'', d.wilayah||'', d.unit||'', d.blok||'', d.compt||'', d.stand||'', num_(d.luas), d.umur||'', d.spesies||'', d.jenisBibit||'', d.uavId||'', d.tanggal||'' ];
    PREP_KEYS.forEach(function(k){ row.push(p[k]||''); });
    row=row.concat([ d.keterangan||'', d.pqaSV||'', a[0]||'', a[1]||'', a[2]||'', now ]);
    aku.getRange(aku.getLastRow()+1,1,1,AKU_HEADER.length).setValues([row]);

    var fl=ensureSheet_(ss,SHEET_FLIGHT,FLIGHT_HEADER), rows=[];
    (d.flights||[]).forEach(function(f,i){ var c=f.coord||{};
      rows.push([ id, (f.no||i+1), f.name||'', f.takeoff||'', f.landing||'', f.duration||'', num_(f.distance),
        f.batTO||'', f.batLD||'', f.idBattery||'',
        (c.x!=null?c.x:''), (c.y!=null?c.y:''), (c.z&&c.e?(c.z+' '+c.e):''), (c.z&&c.n?(c.z+' '+c.n):''),
        f.connMission||'', f.connLanding||'', f.wxTakeoff||'', f.wxMission||'', f.wxLanding||'', f.uavTO||'', f.uavLD||'', f.remarks||'', now ]); });
    if(rows.length) fl.getRange(fl.getLastRow()+1,1,rows.length,FLIGHT_HEADER.length).setValues(rows);
    return {ok:true, id:id};
  } finally { lock.releaseLock(); }
}

/*==================== 保存: Perbandingan ====================*/
function savePerbandingan_(d){
  if(!d) return {ok:false,error:'empty'};
  var lock=LockService.getScriptLock(); lock.waitLock(30000);
  try{
    var ss=SpreadsheetApp.openById(UAV_SHEET_ID);
    var sh=ensureSheet_(ss,SHEET_PERB,PERB_HEADER);
    if(d.clientId && clientIdExists_(sh,3,d.clientId)) return {ok:true,dedup:true};
    var now=new Date();
    var id=Utilities.formatDate(now,Session.getScriptTimeZone(),'yyMMddHHmmss')+'-'+Math.floor(Math.random()*1e4);
    var a=d.assessors||[], b1=d.baris1||[], b2=d.baris2||[], s=d.start||{}, e=d.end||{};
    var row=[ id, now, d.clientId||'', d.umur||'', d.wilayah||'', d.unit||'', d.blok||'', d.stand||'', num_(d.luas), d.species||'', d.noStrip||'', d.halaman||'', d.dari||'', d.tanggal||'' ];
    for(var i=0;i<16;i++) row.push(b1[i]!=null&&b1[i]!==''?num_(b1[i]):'');
    for(var j=0;j<16;j++) row.push(b2[j]!=null&&b2[j]!==''?num_(b2[j]):'');
    row=row.concat([ (s.x!=null?s.x:''),(s.y!=null?s.y:''),(s.z&&s.e?s.z+' '+s.e:''),(s.z&&s.n?s.z+' '+s.n:''),
      (e.x!=null?e.x:''),(e.y!=null?e.y:''),(e.z&&e.e?e.z+' '+e.e:''),(e.z&&e.n?e.z+' '+e.n:''),
      d.pqaSV||'', a[0]||'', a[1]||'', a[2]||'', now ]);
    sh.getRange(sh.getLastRow()+1,1,1,PERB_HEADER.length).setValues([row]);
    return {ok:true, id:id};
  } finally { lock.releaseLock(); }
}

/*==================== 補助 ====================*/
function ensureSheet_(ss,name,header){ var sh=ss.getSheetByName(name); if(!sh) sh=ss.insertSheet(name);
  if(sh.getLastRow()===0){ sh.getRange(1,1,1,header.length).setValues([header]).setFontWeight('bold'); sh.setFrozenRows(1); }
  else { var lc=Math.max(sh.getLastColumn(),1); if(lc<header.length) sh.getRange(1,lc+1,1,header.length-lc).setValues([header.slice(lc)]).setFontWeight('bold'); }
  return sh; }
function clientIdExists_(sh,col,cid){ if(sh.getLastRow()<2) return false;
  var v=sh.getRange(2,col,sh.getLastRow()-1,1).getValues(); for(var i=0;i<v.length;i++) if(String(v[i][0])===String(cid)) return true; return false; }
function num_(v){ var x=parseFloat(String(v).replace(',','.')); return isFinite(x)?x:''; }

/*==================== 初回セットアップ ====================*/
function setup(){
  var ss=SpreadsheetApp.openById(UAV_SHEET_ID);
  ensureSheet_(ss,SHEET_AKU,AKU_HEADER);
  ensureSheet_(ss,SHEET_FLIGHT,FLIGHT_HEADER);
  ensureSheet_(ss,SHEET_PERB,PERB_HEADER);
  Logger.log('UAV setup 完了 (v1)');
}
