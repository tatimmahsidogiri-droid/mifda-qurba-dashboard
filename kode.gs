// ============================================
// MENU SPREADSHEET
// ============================================
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('📊 DASHBOARD')
    .addItem('🔄 Refresh Dashboard', 'refreshDashboard')
    .addSeparator()
    .addItem('📊 Buka Form Laporan', 'bukaFormLaporan')
    .addToUi();
}

function bukaFormLaporan() {
  var url = ScriptApp.getService().getUrl() + '?page=laporan';
  var html = '<script>window.open("' + url + '");google.script.host.close();</script>';
  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(html).setHeight(50).setWidth(200), 'Form Laporan');
}

// ============================================
// ROUTING WEB APP
// ============================================
function doGet(e) {
  var page = e.parameter.page || 'login';
  if (page === 'login') return HtmlService.createHtmlOutputFromFile('Login').setTitle('🔐 Login').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL).addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
  if (page === 'laporan') return HtmlService.createHtmlOutputFromFile('FormLaporan').setTitle('📊 Laporan Bulanan').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL).addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
  if (page === 'admin') return HtmlService.createHtmlOutputFromFile('AdminDashboard').setTitle('📊 Dashboard Admin').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL).addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
  if (page === 'pengurus') return HtmlService.createHtmlOutputFromFile('DashboardPengurus').setTitle('📊 Dashboard Pengurus').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL).addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
  return HtmlService.createHtmlOutputFromFile('Login').setTitle('🔐 Login').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL).addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
}

// ============================================
// LOGIN
// ============================================
function prosesLogin(username, password) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('USERS');
  if (!sheet) return { sukses: false, pesan: 'Sheet USERS tidak ditemukan!' };
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    var u = data[i][0] ? data[i][0].toString().trim() : '';
    var p = data[i][1] ? data[i][1].toString().trim() : '';
    var r = data[i][2] ? data[i][2].toString().trim() : '';
    if (u === username.trim() && p === password.trim()) {
      return { sukses: true, username: username, role: r, token: Utilities.getUuid() };
    }
  }
  return { sukses: false, pesan: '❌ Username atau Password salah!' };
}

// ============================================
// DROPDOWN
// ============================================
function getDropdownData() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('MASTER');
  var data = sheet.getDataRange().getValues();
  var bulanHijri = [], tahunHijri = [], pengurus = [];
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] && data[i][0].toString().trim()) bulanHijri.push(data[i][0].toString().trim());
    if (data[i][2] && data[i][2].toString().trim()) tahunHijri.push(data[i][2].toString().trim());
    if (data[i][4] && data[i][4].toString().trim()) pengurus.push(data[i][4].toString().trim());
  }
  return { bulanHijri: bulanHijri, tahunHijri: tahunHijri, pengurus: pengurus };
}

// ============================================
// UPLOAD GAMBAR
// ============================================
function uploadGambar(base64Data, namaFile) {
  var bagian = base64Data.split(',');
  var mimeType = bagian[0].split(':')[1].split(';')[0];
  var dataGambar = Utilities.base64Decode(bagian[1]);
  if (dataGambar.length > 2097152) throw new Error('Ukuran terlalu besar. Maks 2MB.');
  var blob = Utilities.newBlob(dataGambar, mimeType, namaFile);
  var file = Drive.Files.create({ name: namaFile, mimeType: mimeType }, blob, { fields: 'id,webViewLink' });
  Drive.Permissions.create({ role: 'reader', type: 'anyone' }, file.id);
  return file.webViewLink;
}

// ============================================
// FORM LAPORAN
// ============================================
function simpanLaporan(formData) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('LAPORAN');
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['No','Nama Pengurus','Bulan Hijri','Tahun Hijri','IG Followers','IG Postingan','IG Screenshot','FB Followers','FB Postingan','FB Screenshot','TT Followers','TT Like','TT Screenshot']);
  }
  var data = sheet.getDataRange().getValues();
  var rowDup = -1;
  for (var i = 1; i < data.length; i++) {
    if (data[i][1] && data[i][1].toString().trim() === formData.namaPengurus.trim() &&
        data[i][2] && data[i][2].toString().trim() === formData.bulanHijri.trim() &&
        data[i][3] && data[i][3].toString().trim() === formData.tahunHijri.trim()) {
      rowDup = i + 1; break;
    }
  }
  var ts = new Date();
  function up(base64, platform) {
    if (!base64) return '';
    try { return uploadGambar(base64, formData.namaPengurus.replace(/ /g,'_')+'_'+platform+'_'+formData.bulanHijri+'_'+formData.tahunHijri+'_'+Utilities.formatDate(ts,'Asia/Jakarta','yyyyMMdd_HHmmss')+'.png'); } catch(e) { return 'ERROR: '+e.toString(); }
  }
  var igSS = up(formData.igSS, 'IG');
  var fbSS = up(formData.fbSS, 'FB');
  var ttSS = up(formData.ttSS, 'TT');
  
  var aksi = 'LAPOR';
  
  if (rowDup > -1) {
    sheet.getRange(rowDup,5).setValue(parseInt(formData.igFol)||0);
    sheet.getRange(rowDup,6).setValue(parseInt(formData.igPost)||0);
    if (igSS) sheet.getRange(rowDup,7).setValue(igSS);
    sheet.getRange(rowDup,8).setValue(parseInt(formData.fbFol)||0);
    sheet.getRange(rowDup,9).setValue(parseInt(formData.fbPost)||0);
    if (fbSS) sheet.getRange(rowDup,10).setValue(fbSS);
    sheet.getRange(rowDup,11).setValue(parseInt(formData.ttFol)||0);
    sheet.getRange(rowDup,12).setValue(parseInt(formData.ttLike)||0);
    if (ttSS) sheet.getRange(rowDup,13).setValue(ttSS);
    aksi = 'UPDATE';
  } else {
    sheet.appendRow([sheet.getLastRow(), formData.namaPengurus.trim(), formData.bulanHijri.trim(), formData.tahunHijri.trim(), parseInt(formData.igFol)||0, parseInt(formData.igPost)||0, igSS, parseInt(formData.fbFol)||0, parseInt(formData.fbPost)||0, fbSS, parseInt(formData.ttFol)||0, parseInt(formData.ttLike)||0, ttSS]);
  }
  
  // Tulis log
  tulisLog(formData.namaPengurus, formData.bulanHijri, formData.tahunHijri, aksi);
  
  var msg = '✅ Laporan ' + formData.bulanHijri + ' ' + formData.tahunHijri + ' berhasil ';
  msg += (aksi === 'UPDATE') ? 'DIPERBARUI!' : 'disimpan!';
  return msg;
}

// ============================================
// DATA LAPORAN
// ============================================
function getDaftarPengurus(ss) {
  var sheet = ss.getSheetByName('MASTER');
  if (!sheet) return [];
  var data = sheet.getRange('E2:E').getValues();
  var list = [];
  for (var i = 0; i < data.length; i++) {
    if (data[i][0] && data[i][0].toString().trim()) list.push(data[i][0].toString().trim());
  }
  return list.sort();
}

function ambilDataLaporan(bulan, tahun) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('LAPORAN');
  if (!sheet || sheet.getLastRow() <= 1) return [];
  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 13).getValues();
  var hasil = [];
  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    if (row[2] && row[2].toString().trim() === bulan.trim() && row[3] && row[3].toString().trim() === tahun.trim()) {
      hasil.push({ pengurus: row[1] ? row[1].toString().trim() : '', igFol: Number(row[4])||0, igPost: Number(row[5])||0, fbFol: Number(row[7])||0, fbPost: Number(row[8])||0, ttFol: Number(row[10])||0, ttLike: Number(row[11])||0 });
    }
  }
  return hasil;
}

function ambilDataBulanSebelumnya(bulan, tahun) {
  var bl = ['Muharram','Shafar','Rabi\'ul Awal','Rabi\'ul Akhir','Jumadil Awal','Jumadil Akhir','Rajab','Sya\'ban','Ramadhan','Syawal','Dzulqa\'dah','Dzulhijjah'];
  var idx = bl.indexOf(bulan);
  if (idx === 0) return ambilDataLaporan('Dzulhijjah', String(parseInt(tahun)-1));
  return ambilDataLaporan(bl[idx-1], tahun);
}

// ============================================
// POIN
// ============================================
function hitungPoin(arr, key) {
  var sorted = arr.slice().sort(function(a, b) { return (b[key]||0) - (a[key]||0); });
  var poin = {};
  var rank = 100;
  var lastVal = null;
  for (var i = 0; i < sorted.length; i++) {
    var val = sorted[i][key] || 0;
    if (lastVal !== null && val === lastVal) {
      poin[sorted[i].pengurus] = rank;
    } else {
      rank = 100 - i;
      if (rank < -1) rank = -1;
      poin[sorted[i].pengurus] = rank;
    }
    lastVal = val;
  }
  return poin;
}

// ============================================
// AMBIL SEMUA DATA (BULANAN)
// ============================================
function ambilSemuaData(bulan, tahun) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dataLaporan = ambilDataLaporan(bulan, tahun);
  var dataPrev = ambilDataBulanSebelumnya(bulan, tahun);
  var pengurus = getDaftarPengurus(ss);
  
  var prevMap = {};
  dataPrev.forEach(function(d) { prevMap[d.pengurus] = { ig: d.igFol, fb: d.fbFol, tt: d.ttFol }; });
  
  var dataIG = [], dataFB = [], dataTT = [];
  dataLaporan.forEach(function(d) {
    var pr = prevMap[d.pengurus] || { ig: 0, fb: 0, tt: 0 };
    dataIG.push({ pengurus: d.pengurus, growth: d.igFol - pr.ig, post: d.igPost });
    dataFB.push({ pengurus: d.pengurus, growth: d.fbFol - pr.fb, post: d.fbPost });
    dataTT.push({ pengurus: d.pengurus, growth: d.ttFol - pr.tt, like: d.ttLike });
  });
  
  var pIG_Fol = hitungPoin(dataIG, 'growth');
  var pIG_Post = hitungPoin(dataIG, 'post');
  var pFB_Fol = hitungPoin(dataFB, 'growth');
  var pFB_Post = hitungPoin(dataFB, 'post');
  var pTT_Fol = hitungPoin(dataTT, 'growth');
  var pTT_Like = hitungPoin(dataTT, 'like');
  
  var hasil = {};
  pengurus.forEach(function(p) {
    var dl = dataLaporan.find(function(d){return d.pengurus===p;});
    var pr = prevMap[p] || { ig: 0, fb: 0, tt: 0 };
    var igF = pIG_Fol[p]!==undefined ? pIG_Fol[p] : -1;
    var igP = pIG_Post[p]!==undefined ? pIG_Post[p] : -1;
    var fbF = pFB_Fol[p]!==undefined ? pFB_Fol[p] : -1;
    var fbP = pFB_Post[p]!==undefined ? pFB_Post[p] : -1;
    var ttF = pTT_Fol[p]!==undefined ? pTT_Fol[p] : -1;
    var ttL = pTT_Like[p]!==undefined ? pTT_Like[p] : -1;
    
    hasil[p] = {
      pengurus: p,
      igFolPoin: igF, igPostPoin: igP, totalIG: igF + igP,
      fbFolPoin: fbF, fbPostPoin: fbP, totalFB: fbF + fbP,
      ttFolPoin: ttF, ttLikePoin: ttL, totalTT: ttF + ttL,
      grandTotal: igF + igP + fbF + fbP + ttF + ttL,
      igGrowth: dl ? dl.igFol - pr.ig : 0,
      igPostJml: dl ? dl.igPost : 0,
      fbGrowth: dl ? dl.fbFol - pr.fb : 0,
      fbPostJml: dl ? dl.fbPost : 0,
      ttGrowth: dl ? dl.ttFol - pr.tt : 0,
      ttLikeJml: dl ? dl.ttLike : 0
    };
  });
  
  return { bulan: bulan, tahun: tahun, data: hasil, pengurus: pengurus };
}

// ============================================
// AMBIL DATA RANGE (KLASEMEN AKHIR)
// ============================================
function ambilDataRange(bulanAwal, tahunAwal, bulanAkhir, tahunAkhir) {
  var bl = ['Muharram','Shafar','Rabi\'ul Awal','Rabi\'ul Akhir','Jumadil Awal','Jumadil Akhir','Rajab','Sya\'ban','Ramadhan','Syawal','Dzulqa\'dah','Dzulhijjah'];
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var pengurus = getDaftarPengurus(ss);
  var dataAwal = ambilDataLaporan(bulanAwal, tahunAwal);
  var dataAkhir = ambilDataLaporan(bulanAkhir, tahunAkhir);
  
  // Akumulasi postingan & like
  var akumIG = {}, akumFB = {}, akumTT = {};
  pengurus.forEach(function(p) { akumIG[p]=0; akumFB[p]=0; akumTT[p]=0; });
  
  var idxAwal = bl.indexOf(bulanAwal);
  var idxAkhir = bl.indexOf(bulanAkhir);
  var tahun = parseInt(tahunAwal);
  
  for (var i = idxAwal; i !== idxAkhir + 1; i++) {
    if (i >= 12) { i = 0; tahun++; if (idxAkhir < idxAwal && i > idxAkhir) break; }
    var db = ambilDataLaporan(bl[i], String(tahun));
    db.forEach(function(d) { if (akumIG[d.pengurus]!==undefined) { akumIG[d.pengurus]+=d.igPost; akumFB[d.pengurus]+=d.fbPost; akumTT[d.pengurus]+=d.ttLike; } });
    if (i === idxAkhir) break;
    if (tahun > parseInt(tahunAkhir)) break;
  }
  
  var mapAwal = {}, mapAkhir = {};
  dataAwal.forEach(function(d) { mapAwal[d.pengurus] = { ig: d.igFol, fb: d.fbFol, tt: d.ttFol }; });
  dataAkhir.forEach(function(d) { mapAkhir[d.pengurus] = { ig: d.igFol, fb: d.fbFol, tt: d.ttFol }; });
  
  var dataIG = [], dataFB = [], dataTT = [];
  pengurus.forEach(function(p) {
    var aw = mapAwal[p] || { ig:0, fb:0, tt:0 };
    var ak = mapAkhir[p] || { ig:0, fb:0, tt:0 };
    dataIG.push({ pengurus: p, growth: ak.ig - aw.ig, post: akumIG[p]||0 });
    dataFB.push({ pengurus: p, growth: ak.fb - aw.fb, post: akumFB[p]||0 });
    dataTT.push({ pengurus: p, growth: ak.tt - aw.tt, like: akumTT[p]||0 });
  });
  
  var pIG_Fol = hitungPoin(dataIG, 'growth');
  var pIG_Post = hitungPoin(dataIG, 'post');
  var pFB_Fol = hitungPoin(dataFB, 'growth');
  var pFB_Post = hitungPoin(dataFB, 'post');
  var pTT_Fol = hitungPoin(dataTT, 'growth');
  var pTT_Like = hitungPoin(dataTT, 'like');
  
  var hasil = {};
  pengurus.forEach(function(p) {
    var aw = mapAwal[p] || { ig:0, fb:0, tt:0 };
    var ak = mapAkhir[p] || { ig:0, fb:0, tt:0 };
    var igF = pIG_Fol[p]!==undefined ? pIG_Fol[p] : -1;
    var igP = pIG_Post[p]!==undefined ? pIG_Post[p] : -1;
    var fbF = pFB_Fol[p]!==undefined ? pFB_Fol[p] : -1;
    var fbP = pFB_Post[p]!==undefined ? pFB_Post[p] : -1;
    var ttF = pTT_Fol[p]!==undefined ? pTT_Fol[p] : -1;
    var ttL = pTT_Like[p]!==undefined ? pTT_Like[p] : -1;
    
    hasil[p] = {
      pengurus: p,
      igFolPoin: igF, igPostPoin: igP, totalIG: igF + igP,
      fbFolPoin: fbF, fbPostPoin: fbP, totalFB: fbF + fbP,
      ttFolPoin: ttF, ttLikePoin: ttL, totalTT: ttF + ttL,
      grandTotal: igF + igP + fbF + fbP + ttF + ttL,
      igGrowth: ak.ig - aw.ig, igPostJml: akumIG[p]||0,
      fbGrowth: ak.fb - aw.fb, fbPostJml: akumFB[p]||0,
      ttGrowth: ak.tt - aw.tt, ttLikeJml: akumTT[p]||0
    };
  });
  
  return { data: hasil, pengurus: pengurus };
}

// ============================================
// DATA PENGURUS (DASHBOARD PENGURUS)
// ============================================
function ambilDataPengurus(namaPengurus, bulan, tahun) {
  var allData = ambilSemuaData(bulan, tahun);
  var d = allData.data[namaPengurus] || { igFolPoin:0, igPostPoin:0, totalIG:0, fbFolPoin:0, fbPostPoin:0, totalFB:0, ttFolPoin:0, ttLikePoin:0, totalTT:0, grandTotal:0, igGrowth:0, igPostJml:0, fbGrowth:0, fbPostJml:0, ttGrowth:0, ttLikeJml:0 };
  
  var historis = ambilDataHistoris(SpreadsheetApp.getActiveSpreadsheet(), namaPengurus);
  
  return {
    nama: namaPengurus, bulan: bulan, tahun: tahun,
    igFolPoin: d.igFolPoin, igPostPoin: d.igPostPoin, totalIG: d.totalIG,
    fbFolPoin: d.fbFolPoin, fbPostPoin: d.fbPostPoin, totalFB: d.totalFB,
    ttFolPoin: d.ttFolPoin, ttLikePoin: d.ttLikePoin, totalTT: d.totalTT,
    grandTotal: d.grandTotal,
    igGrowth: d.igGrowth, igPostJml: d.igPostJml,
    fbGrowth: d.fbGrowth, fbPostJml: d.fbPostJml,
    ttGrowth: d.ttGrowth, ttLikeJml: d.ttLikeJml,
    historis: historis
  };
}

function ambilDataHistoris(ss, namaPengurus) {
  var bl = ['Muharram','Shafar','Rabi\'ul Awal','Rabi\'ul Akhir','Jumadil Awal','Jumadil Akhir','Rajab','Sya\'ban','Ramadhan','Syawal','Dzulqa\'dah','Dzulhijjah'];
  var sheet = ss.getSheetByName('LAPORAN');
  var data = sheet.getLastRow() > 1 ? sheet.getRange(2, 1, sheet.getLastRow() - 1, 13).getValues() : [];
  var postingan = {}, follower = {};
  bl.forEach(function(b) { postingan[b]=0; follower[b]=0; });
  data.forEach(function(row) {
    var p = row[1] ? row[1].toString().trim() : '';
    var b = row[2] ? row[2].toString().trim() : '';
    if (p === namaPengurus && b) {
      postingan[b] = (Number(row[5])||0) + (Number(row[8])||0);
      follower[b] = (Number(row[4])||0) + (Number(row[7])||0) + (Number(row[10])||0);
    }
  });
  return { postingan: postingan, follower: follower, bulanList: bl };
}

// ============================================
// BUKA/TUTUP LAPORAN
// ============================================
function getStatusLaporan() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('SETELAN');
  
  if (!sheet) {
    sheet = ss.insertSheet('SETELAN');
    sheet.appendRow(['Bulan Aktif', 'Tahun Aktif', 'Status', 'Dibuka Oleh', 'Waktu Buka']);
    sheet.appendRow(['', '', 'TUTUP', '', '']);
  }
  
  if (sheet.getLastRow() < 2) {
    sheet.appendRow(['', '', 'TUTUP', '', '']);
  }
  
  var data = sheet.getRange(2, 1, 1, 5).getValues()[0];
  
  return {
    bulan: data[0] ? data[0].toString().trim() : '',
    tahun: data[1] ? data[1].toString().trim() : '',
    status: data[2] ? data[2].toString().trim() : 'TUTUP',
    dibukaOleh: data[3] ? data[3].toString().trim() : '',
    waktuBuka: data[4] ? data[4].toString() : ''
  };
}

function bukaLaporan(bulan, tahun) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('SETELAN');
  if (!sheet) {
    sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet('SETELAN');
    sheet.appendRow(['Bulan Aktif', 'Tahun Aktif', 'Status', 'Dibuka Oleh', 'Waktu Buka']);
    sheet.appendRow(['', '', 'TUTUP', '', '']);
  }
  sheet.getRange(2, 1).setValue(bulan);
  sheet.getRange(2, 2).setValue(tahun);
  sheet.getRange(2, 3).setValue('BUKA');
  sheet.getRange(2, 4).setValue('Admin');
  sheet.getRange(2, 5).setValue(new Date());
  return '✅ Laporan ' + bulan + ' ' + tahun + ' berhasil DIBUKA!';
}

function tutupLaporan() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('SETELAN');
  if (sheet && sheet.getLastRow() >= 2) {
    sheet.getRange(2, 3).setValue('TUTUP');
  }
  return '✅ Laporan berhasil DITUTUP!';
}

// ============================================
// LOG AKTIVITAS
// ============================================
function tulisLog(nama, bulan, tahun, aksi) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('LOG');
  if (!sheet) {
    sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet('LOG');
    sheet.appendRow(['Timestamp', 'Nama Pengurus', 'Bulan', 'Tahun', 'Aksi']);
  }
  sheet.appendRow([new Date(), nama, bulan, tahun, aksi]);
}

function getMonitoringData(bulan, tahun) {
  if (!bulan || !tahun) {
    return { totalPengurus: 0, sudahLapor: 0, belumLapor: 0, tigaPertama: [], tigaTerakhir: [], daftarBelum: [] };
  }
  
  bulan = String(bulan).trim();
  tahun = String(tahun).trim();
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetLog = ss.getSheetByName('LOG');
  var pengurus = getDaftarPengurus(ss);
  
  var sudahLapor = [];
  if (sheetLog && sheetLog.getLastRow() > 1) {
    var dataLog = sheetLog.getRange(2, 1, sheetLog.getLastRow() - 1, 5).getValues();
    for (var i = 0; i < dataLog.length; i++) {
      var row = dataLog[i];
      var b = row[2] ? String(row[2]).trim() : '';
      var t = row[3] ? String(row[3]).trim() : '';
      if (b === bulan && t === tahun) {
        sudahLapor.push({
          waktu: row[0] ? row[0].toString() : '',
          nama: row[1] ? String(row[1]).trim() : '',
          aksi: row[4] ? String(row[4]).trim() : ''
        });
      }
    }
  }
  
  sudahLapor.sort(function(a, b) {
    if (!a.waktu || !b.waktu) return 0;
    return new Date(a.waktu) - new Date(b.waktu);
  });
  
  var sudahNama = [];
  for (var i = 0; i < sudahLapor.length; i++) {
    if (sudahNama.indexOf(sudahLapor[i].nama) === -1) {
      sudahNama.push(sudahLapor[i].nama);
    }
  }
  
  var belumLapor = [];
  for (var i = 0; i < pengurus.length; i++) {
    if (sudahNama.indexOf(pengurus[i]) === -1) {
      belumLapor.push(pengurus[i]);
    }
  }
  
  // Konversi waktu ke string untuk memastikan bisa dikirim
  var tigaPertama = [];
  var tigaTerakhir = [];
  
  for (var i = 0; i < sudahLapor.length && i < 3; i++) {
    tigaPertama.push({
      waktu: sudahLapor[i].waktu,
      nama: sudahLapor[i].nama,
      aksi: sudahLapor[i].aksi
    });
  }
  
  var reversed = sudahLapor.slice().reverse();
  for (var i = 0; i < reversed.length && i < 3; i++) {
    tigaTerakhir.push({
      waktu: reversed[i].waktu,
      nama: reversed[i].nama,
      aksi: reversed[i].aksi
    });
  }
  
  var result = {
    totalPengurus: pengurus.length,
    sudahLapor: sudahNama.length,
    belumLapor: belumLapor.length,
    tigaPertama: tigaPertama,
    tigaTerakhir: tigaTerakhir,
    daftarBelum: belumLapor
  };
  
  return result;
}

// ============================================
// FILTER GRUP
// ============================================
function filterMIFDA(list) {
  return list.filter(function(p) {
    return p.toLowerCase().indexOf('mifda') === 0;
  });
}

function filterQURBA(list) {
  return list.filter(function(p) {
    return p.toLowerCase().indexOf('qurba') === 0;
  });
}
