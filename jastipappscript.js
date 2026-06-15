function sendJSON(data) {
  const output = ContentService.createTextOutput(JSON.stringify(data));
  output.setMimeType(ContentService.MimeType.JSON);
  return output;
}

// Ambil data dari Sheet (untuk admin)
function getOrders() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("FormOrder");
  const rows = sheet.getDataRange().getValues();
  rows.shift(); // hapus header
  const data = rows.map(r => ({
    tanggal: r[0],
    kodePesanan: r[1],
    nama: r[2],
    wa: r[3],
    barang: r[4],
    catatan: r[5],
    alamat: r[6],
    status: r[7]
  }));
  return data;
}

function doGet(e) {
  const action = e.parameter.action;

  // === [1] Ambil semua pesanan untuk admin ===
  if (action === "getOrders") {
    const data = getOrders();
    return sendJSON(data);
  }

  // === [2] Tracking berdasarkan kode pesanan ===
  if (action === "track" && e.parameter.kode) {
    const kode = e.parameter.kode;
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("FormOrder");
    const data = sheet.getDataRange().getValues();
    let result = null;

    for (let i = 1; i < data.length; i++) {
      if (data[i][1] === kode) { // kolom ke-2 adalah kodePesanan
        result = {
          kodePesanan: data[i][1],
          nama: data[i][2],
          barang: data[i][4],
          status: data[i][7] || "Belum diproses",
          tanggal: data[i][0],
          alamat: data[i][6]
        };
        break;
      }
    }

    if (!result) {
      return sendJSON({ status: null, msg: "Kode pesanan tidak ditemukan." });
    }

    return sendJSON(result);
  }

  // === [3] Default: test API aktif ===
  return sendJSON({ msg: "Jastip API aktif" });
}

// Update status berdasarkan nama + barang
function updateStatus(nama, barang, status) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("FormOrder");
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][2] === nama && data[i][4] === barang) {
      sheet.getRange(i + 1, 8).setValue(status);
      return true;
    }
  }
  return false;
}

// ✅ Update status berdasarkan kode pesanan
function updateStatusByKode(kodePesanan, status) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("FormOrder");
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][1] === kodePesanan) { // kolom ke-2 = kodePesanan
      sheet.getRange(i + 1, 8).setValue(status);
      return true;
    }
  }
  return false;
}

function doPost(e) {
  try {
    const action = e.parameter.action;

    // === [A] Update status berdasarkan nama + barang ===
    if (action === "updateStatus") {
      const nama = e.parameter.nama;
      const barang = e.parameter.barang;
      const status = e.parameter.status;
      const ok = updateStatus(nama, barang, status);
      return ContentService.createTextOutput(ok ? "OK" : "Not Found");
    }

    // ✅ === [B] Update status berdasarkan kode pesanan ===
    if (action === "updateStatusByKode") {
      const kodePesanan = e.parameter.kodePesanan;
      const status = e.parameter.status;
      const ok = updateStatusByKode(kodePesanan, status);
      return ContentService.createTextOutput(ok ? "OK" : "Not Found");
    }

    // === [C] Tambah order baru ===
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("FormOrder");
    const data = JSON.parse(e.postData.contents);

    const barangList = data.barang || [];
    const jumlahList = data.jumlah || [];
    const barangGabung = barangList.map((b, i) => `${b} (${jumlahList[i] || 1})`).join("; ");

    sheet.appendRow([
      new Date(),
      data.kodePesanan,
      data.nama,
      data.wa,
      barangGabung,
      data.catatan,
      data.alamat,
      "baru"
    ]);

    return ContentService.createTextOutput("OK");

  } catch (err) {
    return ContentService.createTextOutput("Error: " + err.message);
  }
}
