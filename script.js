/**
 * BAHAR BEAUTY STUDIO - MERKEZİ BULUT VERİTABANI API MOTORU
 * Sürüm: 4.0.2 (Arşiv Tarih Senkronizasyon Güncellemesi)
 */

function doGet(e) {
  var action = e.parameter.action;
  var callback = e.parameter.callback;
  var sonuc = { status: "error", message: "Geçersiz Eylem veya Parametre" };

  try {
    // 1. KULLANICI GİRİŞ KONTROLÜ
    if (action === "login") {
      sonuc = kullaniciGirisKontroluVeLog(e.parameter.username, e.parameter.password);
    }
    
    // 2. TÜM KULLANICILARI LİSTELEME
    else if (action === "getkullanicilar") {
      sonuc = tumVerileriListele("salonKullanicilari");
    }

    // 2.B PERSONEL EKLEME VEYA GÜNCELLEME
    else if (action === "kaydetkullanici") {
      sonuc = personelEkleVeyaGuncelle(e.parameter);
    }

    // 2.C PERSONEL SİLME
    else if (action === "silkullanici") {
      sonuc = satirSil("salonKullanicilari", e.parameter.id);
    }
    
    // 3. DANIŞAN LİSTESİNİ ÇEKME
    else if (action === "getdanisanlar") {
      sonuc = tumVerileriListele("salonDanisanlari");
    }

    // 3.B YENİ DANIŞAN EKLEME (HTML'deki ekledanisan ve kaydetdanisan eylemleri için tam uyum)
    else if (action === "ekledanisan" || action === "kaydetdanisan") {
      sonuc = yeniDanisanEkleMetodu(e.parameter.adSoyad || e.parameter.danisanAdiSoyadi);
    }

    // 3.C DANIŞAN GÜNCELLEME
    else if (action === "guncelledanisan") {
      sonuc = danisanGuncelleMetodu(e.parameter.id, e.parameter.adSoyad);
    }

    // 3.D DANIŞAN SİLME
    else if (action === "sildanisan") {
      sonuc = satirSil("salonDanisanlari", e.parameter.id);
    }

    // 4. STOK KATEGORİLERİNİ ÇEKME
    else if (action === "getkategoriler") {
      sonuc = tumVerileriListele("stokKategori");
    }

    // 4.B YENİ KATEGORİ EKLEME
    else if (action === "addkategori") {
      sonuc = yeniKategoriEkleMetodu(e.parameter.kategoriAdi);
    }

    // 4.C KATEGORİ SİLME
    else if (action === "silkategori") {
      sonuc = satirSil("stokKategori", e.parameter.id);
    }

    // 5. STOK VERİLERİNİ ÇEKME
    else if (action === "getstok") {
      sonuc = tumVerileriListele("stokDepoHafizasi");
    }

    // 5.B STOK ÜRÜNÜ EKLEME / GÜNCELLEME
    else if (action === "kaydeturun") {
      sonuc = urunEkleVeyaGuncelleMetodu(e.parameter);
    }

    // 5.C STOK ÜRÜNÜ SİLME
    else if (action === "silurun") {
      sonuc = satirSil("stokDepoHafizasi", e.parameter.id);
    }

    // 6. YENİ SEANS KAYDI & STOKTAN DÜŞÜŞ MİMARİSİ
    else if (action === "addseans") {
      sonuc = seansEkleVeStokDus(e.parameter);
    }

    // 7. SEANS GEÇMİŞİNİ LİSTELEME
    else if (action === "getseanslar") {
      sonuc = tumVerileriListele("kabinSeansGecmisi");
    }

    // 8. SEANS SİLME VE STOK İADE İŞLEMİ
    else if (action === "deleteseans") {
      sonuc = seansSilVeStokIade(e.parameter.id);
    }

    // 9. GÜVENLİK LOGLARINI ÇEKME
    else if (action === "getlogs") {
      sonuc = tumVerileriListele("girisLoglari");
    }

    // 10. GÜVENLİK LOGLARINI SIFIRLAMA
    else if (action === "clearlogs") {
      sonuc = loglariTemizleFonksiyonu();
    }

  } catch (error) {
    sonuc = { status: "error", message: error.toString() };
  }

  // Orijinal JSONP Yapısı
  if (callback) {
    var jsOutput = callback + "(" + JSON.stringify(sonuc) + ");";
    return ContentService.createTextOutput(jsOutput)
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  } else {
    return ContentService.createTextOutput(JSON.stringify(sonuc))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// 🌟 GELİŞTİRİLMİŞ UYUMLU LİSTELEME FONKSİYONU
function tumVerileriListele(sheetName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) return [];

  // 🎯 ARŞİV SAAT KAYMASI ÇÖZÜMÜ: Verileri ham nesne olarak değil, ekranda yazan net metin haliyle (getDisplayValues) çekiyoruz.
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
      
      // Eğer hücre bir tarih nesnesiyse, tablodaki net görünen string değerini (GMT+3 bozulmamış halini) alıyoruz.
      if (hucreVerisi instanceof Date) {
        hucreVerisi = displayData[i][j];
      }
      
      obj[headers[j]] = hucreVerisi;
      
      // Danışan listesi çağrılırken HTML tarafındaki .danisanAdiSoyadi çağrısına tam uyum koruması
      if (sheetName === "salonDanisanlari") {
        var headerLower = headers[j].toLowerCase().replace(/\s/g, "");
        if (headerLower === "danisanadisoyadi" || headerLower === "danisanadsoyad" || headerLower === "adsoyad") {
          obj["danisanAdiSoyadi"] = hucreVerisi;
        }
        if (headerLower === "danisanid" || headerLower === "id") {
          obj["danisanid"] = hucreVerisi;
          obj["id"] = hucreVerisi;
        }
      }
    }
    liste.push(obj);
  }
  return liste;
}

// 🌟 YENİ DANIŞAN EKLEME FONKSİYONU
function yeniDanisanEkleMetodu(adSoyad) {
  if (!adSoyad) return { status: "error", message: "Danışan adı boş olamaz!" };
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("salonDanisanlari");
  if (!sheet) return { status: "error", message: "salonDanisanlari sayfası bulunamadı!" };

  var yeniId = "D-" + (100 + sheet.getLastRow());
  sheet.appendRow([yeniId, adSoyad.trim()]);
  return { status: "success", message: "Danışan başarıyla eklendi." };
}

// 🌟 DANIŞAN GÜNCELLEME FONKSİYONU
function danisanGuncelleMetodu(id, yeniIsim) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("salonDanisanlari");
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(id).trim()) {
      sheet.getRange(i + 1, 2).setValue(yeniIsim.trim());
      return { status: "success" };
    }
  }
  return { status: "error", message: "Danışan bulunamadı." };
}

// 🌟 ORİJİNAL KATEGORİ EKLEME METODU
function yeniKategoriEkleMetodu(kategoriAdi) {
  if (!kategoriAdi) return { status: "error", message: "Kategori adı boş olamaz!" };
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("stokKategori");
  if (!sheet) return { status: "error", message: "stokKategori sayfası bulunamadı!" };

  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][1]).trim().toLowerCase() === String(kategoriAdi).trim().toLowerCase()) {
      return { status: "error", message: "Bu kategori zaten mevcut!" };
    }
  }

  var yeniKatId = "K-" + (100 + sheet.getLastRow());
  sheet.appendRow([yeniKatId, kategoriAdi.trim()]);
  return { status: "success", message: "Kategori başarıyla eklendi." };
}

// 🌟 İZOLE EDİLMİŞ GÜVENLİ ÜRÜN MOTORU
function urunEkleVeyaGuncelleMetodu(p) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("stokDepoHafizasi");
  if (!sheet) return { status: "error", message: "stokDepoHafizasi sayfası bulunamadı!" };

  var data = sheet.getDataRange().getValues();
  var headers = data[0].map(function(k) { return String(k).trim().toLowerCase(); });

  var idIdx = headers.indexOf("urunid") > -1 ? headers.indexOf("urunid") : headers.indexOf("id");
  var urunAdiIdx = headers.indexOf("urunadi");
  var kategoriIdx = headers.indexOf("kategori");
  var toplamHacimIdx = headers.indexOf("toplamhacim");
  var kalanHacimIdx = headers.indexOf("kalanhacim");
  var toplamMaliyetIdx = headers.indexOf("toplammaliyet");
  var birimIdx = headers.indexOf("birim");

  var id = p.id || p.urunid || "";
  var hesaplananHacim = parseFloat(p.toplamHacim || 0);
  var maliyet = parseFloat(p.toplamMaliyet || 0);

  if (id !== "") {
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][idIdx]).trim() === String(id).trim()) {
        if (urunAdiIdx > -1) sheet.getRange(i + 1, urunAdiIdx + 1).setValue(p.urunAdi);
        if (kategoriIdx > -1) sheet.getRange(i + 1, kategoriIdx + 1).setValue(p.kategori);
        if (toplamHacimIdx > -1) sheet.getRange(i + 1, toplamHacimIdx + 1).setValue(hesaplananHacim);
        if (kalanHacimIdx > -1) sheet.getRange(i + 1, kalanHacimIdx + 1).setValue(hesaplananHacim); 
        if (toplamMaliyetIdx > -1) sheet.getRange(i + 1, toplamMaliyetIdx + 1).setValue(maliyet);
        if (birimIdx > -1) sheet.getRange(i + 1, birimIdx + 1).setValue(p.birim);
        return { status: "success", message: "Ürün başarıyla güncellendi." };
      }
    }
  } else {
    var enSonYaziliSatirIndex = 1; 
    var enBuyukIdNumarasi = 0;

    for (var i = 1; i < data.length; i++) {
      var hucreVerisi = String(data[i][idIdx]).trim();
      if (hucreVerisi !== "") {
        enSonYaziliSatirIndex = i + 1;
        var sayisalKisim = parseInt(hucreVerisi.replace(/[^0-9]/g, ""), 10);
        if (!isNaN(sayisalKisim) && sayisalKisim > enBuyukIdNumarasi) {
          enBuyukIdNumarasi = sayisalKisim;
        }
      }
    }

    var hedefYeniSatirNo = enSonYaziliSatirIndex + 1;
    var yeniSayisalId = enBuyukIdNumarasi + 1;
    var yeniIdString = "U-" + (yeniSayisalId < 10 ? "00" + yeniSayisalId : (yeniSayisalId < 100 ? "0" + yeniSayisalId : yeniSayisalId));

    if (idIdx > -1) sheet.getRange(hedefYeniSatirNo, idIdx + 1).setValue(yeniIdString);
    if (urunAdiIdx > -1) sheet.getRange(hedefYeniSatirNo, urunAdiIdx + 1).setValue(p.urunAdi);
    if (kategoriIdx > -1) sheet.getRange(hedefYeniSatirNo, kategoriIdx + 1).setValue(p.kategori);
    if (toplamHacimIdx > -1) sheet.getRange(hedefYeniSatirNo, toplamHacimIdx + 1).setValue(hesaplananHacim);
    if (kalanHacimIdx > -1) sheet.getRange(hedefYeniSatirNo, kalanHacimIdx + 1).setValue(hesaplananHacim); 
    if (toplamMaliyetIdx > -1) sheet.getRange(hedefYeniSatirNo, toplamMaliyetIdx + 1).setValue(maliyet);
    if (birimIdx > -1) sheet.getRange(hedefYeniSatirNo, birimIdx + 1).setValue(p.birim);

    return { status: "success", message: "Yeni ürün başarıyla eklendi." };
  }
  return { status: "error", message: "Sütun eşleşme hatası." };
}

// 🛠️ GARANTİLİ SEANS VE STOK DÜŞÜŞ MOTORU
function seansEkleVeStokDus(p) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var seansSheet = ss.getSheetByName("kabinSeansGecmisi");
  var stokSheet = ss.getSheetByName("stokDepoHafizasi");
  if (!seansSheet || !stokSheet) return { status: "error", message: "Sayfalar bulunamadı!" };

  var yeniSeansId = "D-00" + seansSheet.getLastRow();
  
  // Türkiye Saati (GMT+3) ile kayıt motoru
  var TR_Zamani = Utilities.formatDate(new Date(), "GMT+3", "yyyy-MM-dd HH:mm:ss");
  
  var stokData = stokSheet.getDataRange().getValues();
  var sHeaders = stokData[0].map(function(k) { return String(k).trim().toLowerCase(); });
  var sNameIdx = sHeaders.indexOf("urunadi");
  var sKalanIdx = sHeaders.indexOf("kalanhacim");

  var dusuMiktari = parseFloat(p.tuketimMiktari || 0);
  for (var i = 1; i < stokData.length; i++) {
    if (String(stokData[i][sNameIdx]).trim().toLowerCase() === String(p.urunAdi).trim().toLowerCase()) {
      var mevcutKalan = parseFloat(stokData[i][sKalanIdx] || 0);
      var yeniKalan = Math.max(0, mevcutKalan - dusuMiktari);
      stokSheet.getRange(i + 1, sKalanIdx + 1).setValue(yeniKalan);
      break;
    }
  }

  var gelir = parseFloat(p.gelir || 0);
  var masraf = parseFloat(p.masraf || 0);
  var netKar = gelir - masraf;

  seansSheet.appendRow([
    yeniSeansId, 
    TR_Zamani, 
    p.uzmanAdSoyad, 
    p.danisanAdSoyad, 
    p.islemTipi, 
    p.urunAdi, 
    gelir, 
    masraf, 
    netKar
  ]);

  return { status: "success" };
}

// 🌟 SEANS SİLME VE STOK İADE İŞLEMİ
function seansSilVeStokIade(seansId) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var seansSheet = ss.getSheetByName("kabinSeansGecmisi");
  var stokSheet = ss.getSheetByName("stokDepoHafizasi");
  var seansData = seansSheet.getDataRange().getValues();
  var sHeaders = seansData[0].map(function(k) { return String(k).trim().toLowerCase(); });
  
  var idIdx = sHeaders.indexOf("arsivid") > -1 ? sHeaders.indexOf("arsivid") : 0;
  var uNameIdx = sHeaders.indexOf("urunadi");
  var miktarIdx = sHeaders.indexOf("tuketimmiktari");

  for (var i = 1; i < seansData.length; i++) {
    if (String(seansData[i][idIdx]).trim() === String(seansId).trim()) {
      var urunAdi = seansData[i][uNameIdx];
      var iadeMiktari = parseFloat(seansData[i][miktarIdx] || 0);
      if (urunAdi && iadeMiktari > 0) {
        var stokData = stokSheet.getDataRange().getValues();
        var sNameIdx = stokData[0].map(function(k) { return String(k).trim().toLowerCase(); }).indexOf("urunadi");
        var sKalanIdx = stokData[0].map(function(k) { return String(k).trim().toLowerCase(); }).indexOf("kalanhacim");
        for (var j = 1; j < stokData.length; j++) {
          if (String(stokData[j][sNameIdx]).trim().toLowerCase() === String(urunAdi).trim().toLowerCase()) {
            var mevcutKalan = parseFloat(stokData[j][sKalanIdx] || 0);
            stokSheet.getRange(j + 1, sKalanIdx + 1).setValue(mevcutKalan + iadeMiktari);
            break;
          }
        }
      }
      seansSheet.deleteRow(i + 1);
      return { status: "success" };
    }
  }
  return { status: "error" };
}

// 🌟 SATIR SİLME MOTORU
function satirSil(sheetName, id) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  var data = sheet.getDataRange().getValues();
  var headers = data[0].map(function(k) { return String(k).trim().toLowerCase(); });
  
  var idIndex = headers.indexOf("kategoriid") > -1 ? headers.indexOf("kategoriid") : (headers.indexOf("urunid") > -1 ? headers.indexOf("urunid") : (headers.indexOf("id") > -1 ? headers.indexOf("id") : 0));
  
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][idIndex]).trim() === String(id).trim()) {
      sheet.deleteRow(i + 1);
      return { status: "success" };
    }
  }
  return { status: "error", message: "Kayıt bulunamadı." };
}

// 🌟 KULLANICI GİRİŞ SİSTEMİ
function kullaniciGirisKontroluVeLog(username, password) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var userSheet = ss.getSheetByName("salonKullanicilari");
  var logSheet = ss.getSheetByName("girisLoglari");
  var userData = userSheet.getDataRange().getValues();
  var headers = userData[0].map(function(h) { return String(h).trim().toLowerCase(); });

  var epostaIdx = headers.indexOf("eposta");
  var sifreIdx = headers.indexOf("sifre");
  var adSoyadIdx = headers.indexOf("adsoyad");
  var rolIdx = headers.indexOf("rol");

  for (var i = 1; i < userData.length; i++) {
    if (String(userData[i][epostaIdx]).trim().toLowerCase() === String(username).trim().toLowerCase() && String(userData[i][sifreIdx]).trim() === String(password).trim()) {
      var aktifKullanici = { adSoyad: userData[i][adSoyadIdx], eposta: userData[i][epostaIdx], rol: userData[i][rolIdx] };
      logSheet.appendRow(["L-00" + logSheet.getLastRow(), Utilities.formatDate(new Date(), "GMT+3", "yyyy-MM-dd HH:mm:ss"), username, "BAŞARILI", "giriş yaptı"]);
      return { status: "success", user: aktifKullanici };
    }
  }
  logSheet.appendRow(["L-00" + logSheet.getLastRow(), Utilities.formatDate(new Date(), "GMT+3", "yyyy-MM-dd HH:mm:ss"), username, "HATALI", "deneme"]);
  return { status: "error", message: "Hatalı kimlik." };
}

// 🌟 PERSONEL EKLE/GÜNCELLE MOTORU
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
        return { status: "success" };
      }
    }
  } else {
    sheet.appendRow(["P-" + (100 + sheet.getLastRow()), p.adSoyad, p.eposta, p.sifre, p.rol]);
    return { status: "success" };
  }
}

// 🌟 GÜVENLİK LOGLARINI SIFIRLAMA
function loglariTemizleFonksiyonu() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("girisLoglari");
  if (sheet.getLastRow() > 1) sheet.deleteRows(2, sheet.getLastRow() - 1);
  return { status: "success" };
}