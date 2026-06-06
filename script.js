/**
 * BAHAR BEAUTY STUDIO - MERKEZİ BULUT VERİTABANI API MOTORU
 * Sürüm: 4.8.0 (Canlı CSV Veritabanı ve Sütun Yapısıyla %100 Eşlenmiş Kesin Sürüm)
 */

function doGet(e) {
  var action = e.parameter.action;
  var callback = e.parameter.callback;
  var sonuc = { status: "error", message: "Geçersiz Eylem veya Parametre" };

  try {
    if (action === "login") {
      sonuc = kullaniciGirisKontroluVeLog(e.parameter.username, e.parameter.password);
    }
    else if (action === "getkullanicilar") {
      sonuc = tumVerileriListele("salonKullanicilari");
    }
    else if (action === "kaydetkullanici") {
      sonuc = personelEkleVeyaGuncelle(e.parameter);
    }
    else if (action === "silkullanici") {
      sonuc = satirSil("salonKullanicilari", e.parameter.id);
    }
    else if (action === "getdanisanlar") {
      sonuc = tumVerileriListele("salonDanisanlari");
    }
    else if (action === "ekledanisan" || action === "kaydetdanisan") {
      sonuc = yeniDanisanEkleMetodu(e.parameter.adSoyad || e.parameter.danisanAdiSoyadi);
    }
    else if (action === "guncelledanisan") {
      sonuc = danisanGuncelleMetodu(e.parameter.id, e.parameter.adSoyad || e.parameter.danisanAdiSoyadi);
    }
    else if (action === "sildanisan") {
      sonuc = satirSil("salonDanisanlari", e.parameter.id);
    }
    else if (action === "getkategoriler") {
      sonuc = tumVerileriListele("stokKategori");
    }
    else if (action === "addkategori") {
      sonuc = yeniKategoriEkleMetodu(e.parameter.kategoriAdi);
    }
    else if (action === "silkategori") {
      sonuc = satirSil("stokKategori", e.parameter.id);
    }
    else if (action === "getstok") {
      sonuc = tumVerileriListele("stokDepoHafizasi");
    }
    else if (action === "kaydeturun") {
      sonuc = urunEkleVeyaGuncelleMetodu(e.parameter);
    }
    else if (action === "silurun") {
      sonuc = satirSil("stokDepoHafizasi", e.parameter.id);
    }
    else if (action === "addseans") {
      sonuc = seansEkleVeStokDus(e.parameter);
    }
    else if (action === "getseanslar") {
      sonuc = tumVerileriListele("kabinSeansGecmisi");
    }
    else if (action === "deleteseans" || action === "silseans") {
      var seansId = e.parameter.id || e.parameter.seansid || e.parameter.seansId || e.parameter.islemid || e.parameter.islemId || e.parameter.no;
      sonuc = seansSilVeStokIade(seansId);
    }
    else if (action === "getlogs") {
      sonuc = tumVerileriListele("girisLoglari");
    }
    else if (action === "clearlogs") {
      sonuc = loglariTemizleFonksiyonu();
    }
    else if (action === "gethizmetler" || action === "gethizmetkategori") {
      sonuc = tumVerileriListele("hizmetKatogri"); 
    }
    else if (action === "kaydethizmetkutuphane" || action === "kaydethizmetkategori") {
      sonuc = hizmetKutuphanesineEkle(e.parameter.hizmetAdiid || e.parameter.hizmetAdi, e.parameter.hizmetislemSuresi || e.parameter.islemSuresi);
    }
    else if (action === "getarsiv" || action === "getpaketler") {
      sonuc = tumVerileriListele("hizmetPaketGiris");
    }
    else if (action === "savearsiv" || action === "kaydethizmetpaketgiris") {
      sonuc = arsivKaydiOlustur(e.parameter);
    }
    else if (action === "updatearsiv") {
      sonuc = arsivCiroGuncelle(e.parameter.id, e.parameter.ciro || e.parameter.hizmetAraOdeme);
    }
    else if (action === "updatefullseans") {
      sonuc = arsivSeansVeRandevuGuncelle(e.parameter);
    }
    else if (action === "deletearsiv" || action === "silhizmetpaket") {
      var paketId = e.parameter.id || e.parameter.hizmetNo || e.parameter.no;
      sonuc = satirSil("hizmetPaketGiris", paketId);
    }

  } catch (error) {
    sonuc = { status: "error", message: error.toString() };
  }

  if (callback) {
    var jsOutput = callback + "(" + JSON.stringify(sonuc) + ");";
    return ContentService.createTextOutput(jsOutput)
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  } else {
    return ContentService.createTextOutput(JSON.stringify(sonuc))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function tumVerileriListele(sheetName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) return [];

  var range = sheet.getDataRange();
  var data = range.getValues();
  var displayData = range.getDisplayValues(); 
  
  if (data.length <= 1) return [];

  var headers = data[0].map(function(h) { return String(h).trim(); });
  var liste = [];

  for (var i = 1; i < data.length; i++) {
    var obj = {};
    for (var j = 0; j < headers.length; j++) {
      var hucreVerisi = data[i][j];
      if (hucreVerisi instanceof Date) hucreVerisi = displayData[i][j];
      obj[headers[j]] = hucreVerisi;
      var headerLower = headers[j].toLowerCase().replace(/\s/g, "");

      if (sheetName === "salonDanisanlari") {
        if (headerLower === "danisanadisoyadi" || headerLower === "danisanadsoyad" || headerLower === "adsoyad") obj["danisanAdiSoyadi"] = hucreVerisi;
        if (headerLower === "danisanid" || headerLower === "id") { obj["danisanid"] = hucreVerisi; obj["id"] = hucreVerisi; }
      }
      if (sheetName === "hizmetKatogri") {
        if (headerLower === "hizmetid" || headerLower === "id") { obj["hizmetid"] = hucreVerisi; obj["id"] = hucreVerisi; }
        if (headerLower === "hizmetadiid" || headerLower === "hizmetadi" || headerLower === "hizmetadı") { obj["hizmetAdiid"] = hucreVerisi; obj["hizmetAdi"] = hucreVerisi; }
        if (headerLower === "hizmetislemsuresi" || headerLower === "islemsuresi" || headerLower === "sure") { obj["hizmetislemSuresi"] = hucreVerisi; obj["islemSuresi"] = hucreVerisi; }
      }
      if (sheetName === "hizmetPaketGiris") {
        if (headerLower === "hizmetno" || headerLower === "id") obj["id"] = hucreVerisi;
        if (headerLower === "hizmetuzman" || headerLower === "uzman") obj["uzman"] = hucreVerisi;
        if (headerLower === "hizmetdanisan" || headerLower === "danisan") obj["danisan"] = hucreVerisi;
        if (headerLower === "hizmetpakettercihi" || headerLower === "hizmettercihi") obj["hizmetTercihi"] = hucreVerisi;
        if (headerLower === "hizmetpaketsüresi" || headerLower === "hizmetpaketsuresi" || headerLower === "sure") obj["sure"] = hucreVerisi;
        if (headerLower === "hizmetseanssayisi" || headerLower === "seanssayisi") obj["seansSayisi"] = hucreVerisi;
        if (headerLower === "hizmettoplamodeme" || headerLower === "toplamtutar") obj["toplamTutar"] = hucreVerisi;
        if (headerLower === "hizmetaraodeme" || headerLower === "odenenmiktar") obj["odenenMiktar"] = hucreVerisi;
        if (headerLower === "hizmetkalanborc" || headerLower === "kalanborc") obj["kalanBorc"] = hucreVerisi;
        if (headerLower === "hizmetegelmedurumu" || headerLower === "seansdurumu") obj["seansDurumu"] = hucreVerisi;
        if (headerLower === "kalanseans") obj["kalanSeans"] = hucreVerisi;
        if (headerLower === "pakettarihi") { obj["paketTarihi"] = hucreVerisi; obj["kayitTarihi"] = hucreVerisi; }
        if (headerLower === "sonrakiseanstarihi" || headerLower === "randevutarihi") obj["randevuTarihi"] = hucreVerisi;
        if (headerLower === "sonrakiseanssaati" || headerLower === "randevusaati") obj["randevuSaati"] = hucreVerisi;
      }
    }
    liste.push(obj);
  }
  return liste;
}

function arsivKaydiOlustur(p) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("hizmetPaketGiris");
  if (!sheet) return { status: "error", message: "hizmetPaketGiris sayfası bulunamadı." };
  
  var pId = p.id || p.hizmetNo || "";
  var uzman = p.uzman || p.hizmetUzman || "";
  var danisan = p.danisan || p.hizmetDanisan || "";
  var hizmetTercihi = p.hizmetTercihi || p.hizmetPaketTercihi || "";
  var sure = p.sure || p.hizmetPaketSuresi || "0";
  var seansSayisi = p.seansSayisi || p.hizmetSeansSayisi || "1";
  var toplamTutar = parseFloat(p.toplamTutar || p.hizmetToplamOdeme || "0");
  var odenenMiktar = parseFloat(p.odenenMiktar || p.hizmetAraOdeme || "0");
  var kalanBorc = toplamTutar - odenenMiktar;
  var seansDurumu = p.seansDurumu || p.hizmeteGelmeDurumu || "{}";
  var kalanSeans = p.kalanSeans || p.hizmetSeansSayisi || "1";
  var paketTarihi = p.paketTarihi || new Date().toISOString().slice(0,10);
  var randevuTarihi = p.randevuTarihi || p.sonrakiSeansTarihi || "-";
  var randevuSaati = p.randevuSaati || p.sonrakiSeansSaati || "-";

  sheet.appendRow([pId, uzman, danisan, hizmetTercihi, sure, seansSayisi, toplamTutar, odenenMiktar, kalanBorc, seansDurumu, kalanSeans, paketTarihi, randevuTarihi, randevuSaati]);
  return { status: "success", message: "Paket kartı oluşturuldu." };
}

function arsivCiroGuncelle(id, yeniCiro) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("hizmetPaketGiris");
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(id).trim()) {
      var toplam = parseFloat(data[i][6] || 0); 
      sheet.getRange(i + 1, 8).setValue(yeniCiro); 
      sheet.getRange(i + 1, 9).setValue(toplam - yeniCiro); 
      return { status: "success", message: "Ödeme güncellendi." };
    }
  }
  return { status: "error", message: "Kayıt bulunamadı." };
}

function arsivSeansVeRandevuGuncelle(p) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("hizmetPaketGiris");
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(p.id || p.hizmetNo).trim()) {
      sheet.getRange(i + 1, 10).setValue(p.seansDurumu || p.hizmeteGelmeDurumu);
      sheet.getRange(i + 1, 11).setValue(p.kalanSeans);
      sheet.getRange(i + 1, 13).setValue(p.randevuTarihi || p.sonrakiSeansTarihi);
      sheet.getRange(i + 1, 14).setValue(p.randevuSaati || p.sonrakiSeansSaati);
      return { status: "success", message: "Seans matrisi güncellendi." };
    }
  }
  return { status: "error", message: "Kayıt bulunamadı." };
}

function satirSil(sheetName, id) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) return { status: "error", message: "Sayfa bulunamadı." };
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(id).trim()) {
      sheet.deleteRow(i + 1);
      return { status: "success", message: "Kayıt kalıcı olarak silindi." };
    }
  }
  return { status: "error", message: "Silinecek kayıt bulunamadı." };
}

function kullaniciGirisKontroluVeLog(username, password) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("salonKullanicilari");
  var logSheet = ss.getSheetByName("girisLoglari");
  if (!sheet) return { status: "error", message: "Kullanıcı tablosu bulunamadı." };
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][2]).trim() === String(username).trim() && String(data[i][3]).trim() === String(password).trim()) {
      if (logSheet) logSheet.appendRow(["L-00" + logSheet.getLastRow(), Utilities.formatDate(new Date(), "GMT+3", "yyyy-MM-dd HH:mm:ss"), username, "BAŞARILI", "Giriş Yapıldı"]);
      return { status: "success", adSoyad: data[i][1], rol: data[i][4] };
    }
  }
}

function personelEkleVeyaGuncelle(p) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("salonKullanicilari");
  var data = sheet.getDataRange().getValues();
  var keys = data[0].map(function(k) { return String(k).trim().toLowerCase(); });
  var idIdx = keys.indexOf("id");
  var id = p.id || "";
  if (id !== "") {
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][idIdx]).trim() === String(id).trim()) {
        sheet.getRange(i + 1, keys.indexOf("adsoyad") + 1).setValue(p.adSoyad);
        sheet.getRange(i + 1, keys.indexOf("eposta") + 1).setValue(p.eposta);
        sheet.getRange(i + 1, keys.indexOf("sifre") + 1).setValue(p.sifre);
        sheet.getRange(i + 1, keys.indexOf("rol") + 1).setValue(p.rol);
        return { status: "success", message: "Personel güncellendi." };
      }
    }
  }
  var yeniId = "P-" + (100 + sheet.getLastRow());
  sheet.appendRow([yeniId, p.adSoyad, p.eposta, p.sifre, p.rol]);
  return { status: "success", message: "Personel başarıyla eklendi." };
}

function yeniDanisanEkleMetodu(adSoyad) {
  if (!adSoyad) return { status: "error", message: "Danışan adı boş olamaz!" };
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("salonDanisanlari");
  var yeniId = "D-" + (100 + sheet.getLastRow());
  sheet.appendRow([yeniId, adSoyad.trim()]);
  return { status: "success", message: "Danışan başarıyla eklendi.", id: yeniId };
}

function danisanGuncelleMetodu(id, yeniIsim) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("salonDanisanlari");
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(id).trim()) {
      sheet.getRange(i + 1, 2).setValue(yeniIsim.trim());
      return { status: "success", message: "Danışan güncellendi." };
    }
  }
  return { status: "error", message: "Danışan bulunamadı." };
}

function yeniKategoriEkleMetodu(kategoriAdi) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("stokKategori");
  var yeniKatId = "K-" + (100 + sheet.getLastRow());
  sheet.appendRow([yeniKatId, kategoriAdi.trim()]);
  return { status: "success", message: "Kategori başarıyla eklendi." };
}

function urunEkleVeyaGuncelleMetodu(p) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("stokDepoHafizasi");
  var data = sheet.getDataRange().getValues();
  var headers = data[0].map(function(k) { return String(k).trim().toLowerCase(); });
  var idIdx = headers.indexOf("urunid") > -1 ? headers.indexOf("urunid") : headers.indexOf("id");
  var id = p.id || p.urunid || "";
  var hesapHacim = parseFloat(p.tophamHacim || p.toplamHacim || 0);
  var maliyet = parseFloat(p.toplamMaliyet || 0);
  var birimMaliyet = hesapHacim > 0 ? (maliyet / hesapHacim) : 0;

  if (id !== "") {
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][idIdx]).trim() === String(id).trim()) {
        sheet.getRange(i + 1, headers.indexOf("urunadi") + 1).setValue(p.urunAdi);
        sheet.getRange(i + 1, headers.indexOf("kategori") + 1).setValue(p.kategori);
        sheet.getRange(i + 1, headers.indexOf("toplamhacim") + 1).setValue(hesapHacim);
        sheet.getRange(i + 1, headers.indexOf("kalanhacim") + 1).setValue(hesapHacim);
        sheet.getRange(i + 1, headers.indexOf("toplammaliyet") + 1).setValue(maliyet);
        sheet.getRange(i + 1, headers.indexOf("birimmaliyet") + 1).setValue(birimMaliyet);
        sheet.getRange(i + 1, headers.indexOf("birim") + 1).setValue(p.birim);
        return { status: "success", message: "Ürün güncellendi." };
      }
    }
  }
  var maxId = 0; 
  for (var i = 1; i < data.length; i++) {
    var num = parseInt(String(data[i][idIdx]).replace(/\D/g, ''));
    if (!isNaN(num) && num > maxId) maxId = num;
  }
  var yeniUrunId = "U-" + (maxId + 1).toString().padStart(3, '0');
  sheet.appendRow([yeniUrunId, p.kategori, p.urunAdi, hesapHacim, hesapHacim, p.birim, maliyet, birimMaliyet]);
  return { status: "success", message: "Yeni Ürün eklendi." };
}

function hizmetKutuphanesineEkle(ad, sure) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("hizmetKatogri"); 
  var hId = "H-" + (100 + sheet.getLastRow());
  sheet.appendRow([hId, ad.trim(), sure.toString().trim()]);
  return { status: "success", message: "Hizmet eklendi.", id: hId };
}

// 🎯 KABIN SEANS EKLEME: Canlı CSV Yapısına göre 9 Sütunlu Yapıya Tam Eşlendi!
function seansEkleVeStokDus(p) {
  var seansSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("kabinSeansGecmisi");
  var stokSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("stokDepoHafizasi");
  
  // Canlı tablonuzdaki gibi ID yapısı "D-00XX" formatına uyumlu dinamik üretilir
  var sId = "D-" + String(1000 + seansSheet.getLastRow()).substring(1);
  
  var sTarih = p.tarih || p.islemTarihi || Utilities.formatDate(new Date(), "GMT+3", "yyyy-MM-dd HH:mm:ss");
  var uName = p.uzmanAdi || p.uzmanAdSoyad || p.uzman || "";
  var dName = p.danisanAdi || p.danisanAdSoyad || p.danisan || "";
  var kUrun = p.kullanilanUrun || p.urunAdi || "";
  var kMiktar = parseFloat(p.kullanilanMiktar || p.tuketimMiktari || 0);
  var ciro = parseFloat(p.gelir || p.ciro || 0);
  
  var kategori = "";
  var masraf = 0;
  
  // Stok deposundan ürünün kategorisini ve birim maliyetini canlı oku
  if (kUrun) {
    var stokData = stokSheet.getDataRange().getValues();
    for (var i = 1; i < stokData.length; i++) {
      if (String(stokData[i][2]).trim() === String(kUrun).trim()) {
        kategori = stokData[i][1]; // Sütun 2: Kategori
        var birimMaliyet = parseFloat(stokData[i][7] || 0); // Sütun 8: birimMaliyet
        masraf = kMiktar * birimMaliyet; // Harcanan miktar kadar masraf hesapla
        
        // Stok deposundan hacmi düş
        var mevcutKalan = parseFloat(stokData[i][4] || 0); // Sütun 5: kalanHacim
        stokSheet.getRange(i + 1, 5).setValue(mevcutKalan - kMiktar);
        break;
      }
    }
  }
  
  var kar = ciro - masraf;

  // Canlı CSV Sıralaması: arsivId, islemTarihi, uzman, danisanAdi, kategori, urunAdi, ciro, masraflar, kar
  seansSheet.appendRow([sId, sTarih, uName, dName, kategori, kUrun, ciro, masraf, kar]);
  return { status: "success", message: "Seans işlendi." };
}

// 🎯 KABIN SEANS SILME VE IADE: Canlı 9 Sütunlu İndislere Göre Kilitlendi!
function seansSilVeStokIade(id) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var seansSheet = ss.getSheetByName("kabinSeansGecmisi");
  var stokSheet = ss.getSheetByName("stokDepoHafizasi");
  var data = seansSheet.getDataRange().getValues();
  
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(id).trim()) {
      var urunAdi = data[i][5]; // 🎯 Canlı Veride 6. Sütun (İndis 5) -> urunAdi
      var masrafTutari = parseFloat(data[i][7] || 0); // 🎯 Canlı Veride 8. Sütun (İndis 7) -> masraflar
      
      if (urunAdi) {
        var stokData = stokSheet.getDataRange().getValues();
        for (var j = 1; j < stokData.length; j++) {
          if (String(stokData[j][2]).trim() === String(urunAdi).trim()) { // Stok Ürün Adı (İndis 2)
            var birimMaliyet = parseFloat(stokData[j][7] || 0); // Birim Maliyet (İndis 7)
            
            // Masraf ve birim maliyet üzerinden tüketilen gerçek miktarı tam iade et
            var iadeMiktari = birimMaliyet > 0 ? (masrafTutari / birimMaliyet) : 0;
            
            var mevcutKalan = parseFloat(stokData[j][4] || 0); // Kalan Hacim (İndis 4)
            stokSheet.getRange(j + 1, 4 + 1).setValue(mevcutKalan + iadeMiktari);
            break;
          }
        }
      }
      seansSheet.deleteRow(i + 1);
      return { status: "success", message: "Seans iptal edildi." };
    }
  }
  return { status: "error", message: "Seans kaydı bulunamadı." };
}

function loglariTemizleFonksiyonu() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("girisLoglari");
  var basliklar = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues();
  sheet.clearContents();
  sheet.getRange(1, 1, 1, basliklar[0].length).setValues(basliklar);
  return { status: "success", message: "Loglar sıfırlandı." };
}
