function doGet() {
  return HtmlService.createTemplateFromFile('Index').evaluate()
    .setTitle(APP_CONFIG.appName + ' | ' + APP_CONFIG.schoolName)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}
function include(filename) { return HtmlService.createHtmlOutputFromFile(filename).getContent(); }
function getDb_() {
  const props = PropertiesService.getScriptProperties();
  let id = props.getProperty('PKKS_SPREADSHEET_ID');
  if (!id) {
    const ss = SpreadsheetApp.create('Database SIP-PKKS 2026 - SMP Negeri 3 Cisarua');
    id = ss.getId(); props.setProperty('PKKS_SPREADSHEET_ID', id);
  }
  return SpreadsheetApp.openById(id);
}
function setupApp() {
  const ss = getDb_();
  const defs = [
    [SHEETS.evidence, EVIDENCE_HEADERS],
    [SHEETS.users, ['user_id','nama','email','peran','status_akun','tanggal_dibuat']],
    [SHEETS.activity, ['log_id','waktu','email_pengguna','aktivitas','kode_bukti','data_sebelum','data_sesudah']],
    [SHEETS.settings, ['kunci','nilai']],
    [SHEETS.maps, ['map_id','nama_map','judul_map','bidang','jumlah_bukti']]
  ];
  defs.forEach(([name, headers]) => {
    let sh = ss.getSheetByName(name); if (!sh) sh = ss.insertSheet(name);
    if (sh.getLastRow() === 0) sh.getRange(1,1,1,headers.length).setValues([headers]);
    sh.setFrozenRows(1);
  });
  const evidenceSh = ss.getSheetByName(SHEETS.evidence);
  if (evidenceSh.getLastRow() < 2) {
    const rows = [];
    SEED_MAPS.forEach(m => m.evidence.forEach(e => rows.push([
      Utilities.getUuid(),m.id,m.nama,m.judul,m.bidang,e.kode,
      e.indikator || '',e.kode,e.deskripsi,e.url,'Belum diperiksa',e.keterangan,
      '', '', '', '',new Date(),new Date()
    ])));
    if (rows.length) evidenceSh.getRange(2,1,rows.length,EVIDENCE_HEADERS.length).setValues(rows);
  }
  const mapSh = ss.getSheetByName(SHEETS.maps);
  if (mapSh.getLastRow() < 2) {
    mapSh.getRange(2,1,SEED_MAPS.length,5).setValues(SEED_MAPS.map(m=>[m.id,m.nama,m.judul,m.bidang,m.evidence.length]));
  }
  const settings = ss.getSheetByName(SHEETS.settings);
  if (settings.getLastRow() < 2) {
    settings.getRange(2,1,5,2).setValues([
      ['nama_sekolah',APP_CONFIG.schoolName],['tahun_pkks',APP_CONFIG.year],
      ['zona_waktu',APP_CONFIG.timezone],['database_id',ss.getId()],
      ['catatan','Status bukti harus diverifikasi petugas; tautan aktif bukan bukti kelengkapan.']
    ]);
  }
  return {ok:true,message:'Inisialisasi selesai.',databaseId:ss.getId(),rows:Math.max(0,evidenceSh.getLastRow()-1)};
}
function currentUser_() {
  const email = (Session.getActiveUser().getEmail() || '').toLowerCase().trim();
  let role = 'viewer';
  if (APP_CONFIG.adminEmails.map(x=>x.toLowerCase()).includes(email) && !email.startsWith('ganti-email')) role='admin';
  else if (APP_CONFIG.principalEmails.map(x=>x.toLowerCase()).includes(email)) role='principal';
  else if (APP_CONFIG.teamEmails.map(x=>x.toLowerCase()).includes(email)) role='team';
  return {email:email || 'Email tidak tersedia',role:role,authorized:!!email && role!=='viewer'};
}
function getAppInit() {
  ensureSetup_();
  const user=currentUser_();
  return {appName:APP_CONFIG.appName,schoolName:APP_CONFIG.schoolName,schoolArea:APP_CONFIG.schoolArea,year:APP_CONFIG.year,user:user,maps:getMaps(),dashboard:getDashboard_()};
}
function ensureSetup_() {
  const ss=getDb_();
  if (!ss.getSheetByName(SHEETS.evidence)) setupApp();
}
function getMaps() {
  ensureSetup_();
  return SEED_MAPS.map(m=>({id:m.id,nama:m.nama,judul:m.judul,bidang:m.bidang,jumlah:m.evidence.length}));
}
function rowsToObjects_() {
  const sh=getDb_().getSheetByName(SHEETS.evidence);
  if (sh.getLastRow()<2) return [];
  const values=sh.getRange(2,1,sh.getLastRow()-1,EVIDENCE_HEADERS.length).getValues();
  return values.map(r=>{const o={}; EVIDENCE_HEADERS.forEach((h,i)=>o[h]=r[i] instanceof Date?Utilities.formatDate(r[i],APP_CONFIG.timezone,'yyyy-MM-dd HH:mm:ss'):r[i]); return o;});
}
function listEvidence(filters) {
  ensureSetup_(); filters=filters||{};
  const q=String(filters.q||'').toLowerCase().trim(), map=String(filters.map||''), status=String(filters.status||'');
  let rows=rowsToObjects_().filter(r=>
    (!map||String(r.map_id)===map) && (!status||r.status_bukti===status) &&
    (!q||[r.kode_bukti,r.nama_map,r.judul_map,r.nama_indikator,r.deskripsi_bukti,r.keterangan,r.pemeriksa].join(' ').toLowerCase().includes(q))
  );
  rows.sort((a,b)=>String(a.kode_bukti).localeCompare(String(b.kode_bukti),undefined,{numeric:true}));
  return rows;
}
function getDashboard_() {
  const rows=rowsToObjects_();
  const statuses=['Belum diperiksa','Belum tersedia','Dalam proses','Tersedia, perlu perbaikan','Lengkap dan terverifikasi'];
  const count=s=>rows.filter(r=>r.status_bukti===s).length;
  const maps=SEED_MAPS.map(m=>{
    const rs=rows.filter(r=>Number(r.map_id)===m.id);
    const complete=rs.filter(r=>r.status_bukti==='Lengkap dan terverifikasi').length;
    return {id:m.id,nama:m.nama,judul:m.judul,jumlah:rs.length,lengkap:complete,persen:rs.length?Math.round(complete/rs.length*100):0};
  });
  const checked=rows.filter(r=>r.status_bukti!=='Belum diperiksa').length;
  const complete=count('Lengkap dan terverifikasi');
  const activeLinks=rows.filter(r=>/^https?:\/\//i.test(String(r.url_dokumen||''))).length;
  const sh=getDb_().getSheetByName(SHEETS.activity);
  let recent=[];
  if(sh && sh.getLastRow()>1) recent=sh.getRange(Math.max(2,sh.getLastRow()-4),1,Math.min(5,sh.getLastRow()-1),7).getDisplayValues().reverse().map(r=>({waktu:r[1],email:r[2],aktivitas:r[3],kode:r[4]}));
  return {total:rows.length,totalMaps:12,complete:complete,missing:count('Belum tersedia'),inProgress:count('Dalam proses'),needsFix:count('Tersedia, perlu perbaikan'),unchecked:count('Belum diperiksa'),activeLinks:activeLinks,percent:rows.length?Math.round(complete/rows.length*100):0,checked:checked,statuses:statuses,maps:maps,recent:recent};
}
function getDashboard(){ensureSetup_();return getDashboard_();}
function saveEvidence(payload) {
  assertRole_(['admin','principal','team']);
  if(!payload || !payload.bukti_id) throw new Error('ID bukti tidak ditemukan.');
  const allowed=['Belum diperiksa','Belum tersedia','Dalam proses','Tersedia, perlu perbaikan','Lengkap dan terverifikasi'];
  if(payload.status_bukti && allowed.indexOf(payload.status_bukti)<0) throw new Error('Status tidak valid.');
  const lock=LockService.getScriptLock(); lock.waitLock(15000);
  try {
    const sh=getDb_().getSheetByName(SHEETS.evidence), data=sh.getDataRange().getValues();
    const headers=data[0], idCol=headers.indexOf('bukti_id');
    let rowIndex=-1;
    for(let i=1;i<data.length;i++) if(String(data[i][idCol])===String(payload.bukti_id)){rowIndex=i+1;break;}
    if(rowIndex<2) throw new Error('Data bukti tidak ditemukan.');
    const old=data[rowIndex-1].map(v=>v instanceof Date?v.toISOString():v);
    const editable=['url_dokumen','status_bukti','keterangan','tanggal_pemeriksaan','pemeriksa','catatan_evaluasi','rekomendasi'];
    editable.forEach(k=>{const col=headers.indexOf(k);if(col>=0 && Object.prototype.hasOwnProperty.call(payload,k)) sh.getRange(rowIndex,col+1).setValue(String(payload[k]||''));});
    sh.getRange(rowIndex,headers.indexOf('tanggal_diperbarui')+1).setValue(new Date());
    const now=sh.getRange(rowIndex,1,1,headers.length).getValues()[0].map(v=>v instanceof Date?v.toISOString():v);
    logActivity_('Memperbarui bukti',payload.kode_bukti||'',old,now);
    return {ok:true,message:'Perubahan berhasil disimpan.',dashboard:getDashboard_()};
  } finally {lock.releaseLock();}
}
function addEvidence(payload) {
  assertRole_(['admin']);
  if(!payload || !payload.map_id || !payload.kode_bukti || !payload.deskripsi_bukti) throw new Error('Map, kode bukti, dan deskripsi wajib diisi.');
  const sh=getDb_().getSheetByName(SHEETS.evidence);
  const existing=rowsToObjects_().some(r=>String(r.map_id)===String(payload.map_id)&&String(r.kode_bukti).toLowerCase()===String(payload.kode_bukti).toLowerCase());
  if(existing) throw new Error('Kode bukti tersebut sudah ada pada Map ini.');
  const m=SEED_MAPS.find(x=>x.id===Number(payload.map_id)); if(!m) throw new Error('Map tidak valid.');
  const row=[Utilities.getUuid(),m.id,m.nama,m.judul,m.bidang,payload.kode_indikator||'',payload.nama_indikator||'',payload.kode_bukti,payload.deskripsi_bukti,payload.url_dokumen||'','Belum diperiksa',payload.keterangan||'','','','','',new Date(),new Date()];
  sh.appendRow(row); logActivity_('Menambah bukti',payload.kode_bukti,'',row);
  return {ok:true,message:'Bukti berhasil ditambahkan.'};
}
function deleteEvidence(id) {
  assertRole_(['admin']);
  const sh=getDb_().getSheetByName(SHEETS.evidence), data=sh.getDataRange().getValues();
  for(let i=1;i<data.length;i++) if(String(data[i][0])===String(id)){
    const old=data[i].map(v=>v instanceof Date?v.toISOString():v);
    sh.deleteRow(i+1); logActivity_('Menghapus bukti',old[7],old,''); return {ok:true,message:'Bukti dihapus.'};
  }
  throw new Error('Data tidak ditemukan.');
}
function getActivity() {
  assertRole_(['admin','principal']);
  const sh=getDb_().getSheetByName(SHEETS.activity);
  if(sh.getLastRow()<2) return [];
  return sh.getRange(2,1,sh.getLastRow()-1,7).getDisplayValues().reverse().slice(0,300).map(r=>({id:r[0],waktu:r[1],email:r[2],aktivitas:r[3],kode:r[4],sebelum:r[5],sesudah:r[6]}));
}
function logActivity_(action,code,before,after) {
  const user=currentUser_(), sh=getDb_().getSheetByName(SHEETS.activity);
  sh.appendRow([Utilities.getUuid(),new Date(),user.email,action,code,JSON.stringify(before||''),JSON.stringify(after||'')]);
}
function assertRole_(roles) {
  const u=currentUser_(); if(!u.authorized || roles.indexOf(u.role)<0) throw new Error('Akses ditolak. Pastikan email Google Anda sudah didaftarkan pada Config.gs.');
}
function exportCsv(filters) {
  const u=currentUser_(); if(!u.authorized) throw new Error('Akses ditolak.');
  const rows=listEvidence(filters||{});
  const headers=['nama_map','bidang','kode_bukti','nama_indikator','deskripsi_bukti','url_dokumen','status_bukti','keterangan','tanggal_pemeriksaan','pemeriksa','catatan_evaluasi','rekomendasi'];
  const csv=[headers.join(','),...rows.map(r=>headers.map(h=>'"'+String(r[h]??'').replace(/"/g,'""')+'"').join(','))].join('\r\n');
  return {filename:'Rekap_PKKS_2026.csv',content:'\uFEFF'+csv};
}
function doHealthCheck(){return {ok:true,app:APP_CONFIG.appName,time:new Date().toISOString(),user:currentUser_()};}

/**
 * Import Excel .xlsx melalui Advanced Google Drive service (Drive API v3).
 * Aktifkan Services > Drive API di editor Apps Script terlebih dahulu.
 * Menggabungkan berdasarkan Map + kode bukti; status dan catatan verifikasi dipertahankan.
 */
function importExcelBase64(fileName, base64Data, updateExisting) {
  assertRole_(['admin']);
  if (!fileName || !/\.xlsx?$/i.test(fileName)) throw new Error('Pilih file Excel .xlsx atau .xls.');
  if (!base64Data) throw new Error('File kosong.');
  const bytes=Utilities.base64Decode(base64Data);
  const blob=Utilities.newBlob(bytes,'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',fileName);
  let tempId='';
  try {
    const converted=Drive.Files.create({name:'TEMP_IMPORT_PKKS_'+Date.now(),mimeType:'application/vnd.google-apps.spreadsheet'},blob);
    tempId=converted.id;
    const source=SpreadsheetApp.openById(tempId);
    const dest=getDb_().getSheetByName(SHEETS.evidence);
    const current=rowsToObjects_();
    const index={}; current.forEach(r=>index[String(r.map_id)+'|'+String(r.kode_bukti).toLowerCase()]=r);
    const destData=dest.getDataRange().getValues(), headers=destData[0];
    const rowIndex={}; for(let i=1;i<destData.length;i++)rowIndex[String(destData[i][1])+'|'+String(destData[i][7]).toLowerCase()]=i+1;
    let added=0,updated=0,skipped=0,errors=[];
    source.getSheets().forEach(sheet=>{
      const match=sheet.getName().match(/Map\s*(\d+)/i); if(!match)return;
      const mapId=Number(match[1]), map=SEED_MAPS.find(m=>m.id===mapId); if(!map)return;
      const vals=sheet.getDataRange().getDisplayValues(); let indicator='';
      vals.forEach((r,ri)=>{
        const code=String(r[0]||'').trim();
        if(/^BUTIR\b/i.test(code)){indicator=code;return;}
        if(!/^\d+\.\d+\.[a-z0-9]+$/i.test(code))return;
        const key=mapId+'|'+code.toLowerCase(), desc=String(r[1]||'').trim();
        if(!desc){errors.push(sheet.getName()+' baris '+(ri+1)+': deskripsi kosong untuk '+code);return;}
        const url=String(r[2]||'').trim(), note=String(r[4]||'').trim();
        if(rowIndex[key]){
          if(updateExisting){
            const rn=rowIndex[key];
            dest.getRange(rn,headers.indexOf('judul_map')+1).setValue(map.judul);
            dest.getRange(rn,headers.indexOf('bidang')+1).setValue(map.bidang);
            dest.getRange(rn,headers.indexOf('nama_indikator')+1).setValue(indicator);
            dest.getRange(rn,headers.indexOf('deskripsi_bukti')+1).setValue(desc);
            dest.getRange(rn,headers.indexOf('url_dokumen')+1).setValue(url);
            dest.getRange(rn,headers.indexOf('keterangan')+1).setValue(note);
            dest.getRange(rn,headers.indexOf('tanggal_diperbarui')+1).setValue(new Date());
            updated++;
          }else skipped++;
        }else{
          const newRow=[Utilities.getUuid(),mapId,map.nama,map.judul,map.bidang,indicator,indicator,code,desc,url,'Belum diperiksa',note,'','','','',new Date(),new Date()];
          dest.appendRow(newRow); rowIndex[key]=dest.getLastRow(); added++;
        }
      });
    });
    logActivity_('Impor Excel',fileName,'',{ditambahkan:added,diperbarui:updated,dilewati:skipped});
    return {ok:true,added:added,updated:updated,skipped:skipped,errors:errors.slice(0,30),message:'Impor selesai. Status dan catatan pemeriksaan yang sudah ada dipertahankan.'};
  } catch(e) {
    throw new Error('Impor Excel gagal. Pastikan Drive API sudah diaktifkan dan file berisi worksheet Map 1–Map 12. Detail: '+e.message);
  } finally {
    if(tempId) { try { Drive.Files.remove(tempId); } catch(e) {} }
  }
}
