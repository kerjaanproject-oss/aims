// ================= ENTRY POINT & API HANDLER =================
function doGet(e) {
  if (e && e.parameter && e.parameter.action) {
    let args = [];
    if (e.parameter.args) {
      try {
        args = JSON.parse(e.parameter.args);
      } catch (err) {
        args = [e.parameter.args];
      }
    }
    return handleApiRequest(e.parameter.action, args);
  }
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('SMART APP BBB')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function doPost(e) {
  try {
    let contents = {};
    if (e && e.postData && e.postData.contents) {
      contents = JSON.parse(e.postData.contents);
    }
    const action = contents.action || (e && e.parameter ? e.parameter.action : null);
    const args = contents.args || (e && e.parameter && e.parameter.args ? JSON.parse(e.parameter.args) : []);
    return handleApiRequest(action, args);
  } catch (err) {
    return createJsonResponse({ success: false, message: "Server error: " + err.toString() });
  }
}

function handleApiRequest(action, args) {
  try {
    if (!action) {
      return createJsonResponse({ success: false, message: "Action parameter missing." });
    }
    // Safety check to ensure function exists and isn't private
    if (typeof this[action] === 'function' && !action.startsWith('_')) {
      const result = this[action].apply(this, Array.isArray(args) ? args : [args]);
      return createJsonResponse({ success: true, result: result });
    } else {
      return createJsonResponse({ success: false, message: "Function not found: " + action });
    }
  } catch (err) {
    return createJsonResponse({ success: false, message: err.toString() });
  }
}

function createJsonResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// ================= DATABASE INITIALIZATION =================
function setupDatabase() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 1. Sheet Users
  let usersSheet = ss.getSheetByName("Users");
  if (!usersSheet) {
    usersSheet = ss.insertSheet("Users");
    usersSheet.appendRow(["Username", "PasswordHash", "Fullname", "Role", "ID Karyawan"]);
    // Default Credentials: admin / Admin@2026 and staff / Admin@2026
    usersSheet.appendRow(["admin", hashPassword("Admin@2026"), "Administrator Perdana", "Admin", ""]);
    usersSheet.appendRow(["staff", hashPassword("Admin@2026"), "Siti Aminah Staf", "Staff", ""]);
  } else {
    const uHeaders = usersSheet.getRange(1, 1, 1, usersSheet.getLastColumn()).getValues()[0].map(h => String(h).trim());
    if (uHeaders.indexOf("ID Karyawan") === -1) {
      usersSheet.insertColumnAfter(4);
      usersSheet.getRange(1, 5).setValue("ID Karyawan");
    }
  }

  // 2. Sheet MasterData
  let masterSheet = ss.getSheetByName("MasterData");
  if (!masterSheet) {
    masterSheet = ss.insertSheet("MasterData");
    masterSheet.appendRow(["Category", "Value"]);

    // Default values
    const defaults = [
      ["departemen", "IT"], ["departemen", "HRD"], ["departemen", "Creative"], ["departemen", "Finance"], ["departemen", "Operational"],
      ["jabatan", "Lead Developer"], ["jabatan", "Senior Supervisor"], ["jabatan", "HR Specialist"], ["jabatan", "UI/UX Designer"], ["jabatan", "Staff Admin"],
      ["lokasiKerja", "Head Office Jakarta"], ["lokasiKerja", "Bandung Hub"], ["lokasiKerja", "Cabang Surabaya"],
      ["statusKaryawan", "Tetap"], ["statusKaryawan", "Kontrak"], ["statusKaryawan", "Magang"], ["statusKaryawan", "Harian Lepas"],
      ["pendidikan", "SMA/SMK"], ["pendidikan", "D3 Akuntansi"], ["pendidikan", "S1 Teknik Informatika"], ["pendidikan", "S1 Manajemen SDM"], ["pendidikan", "S1 Desain Komunikasi Visual"],
      ["agama", "Islam"], ["agama", "Kristen"], ["agama", "Katolik"], ["agama", "Hindu"], ["agama", "Buddha"], ["agama", "Khonghucu"],
      ["bank", "Bank Mandiri"], ["bank", "Bank BCA"], ["bank", "Bank BNI"], ["bank", "Bank BRI"],
      ["golonganDarah", "A"], ["golonganDarah", "B"], ["golonganDarah", "AB"], ["golonganDarah", "O"], ["golonganDarah", "Belum Tahu"],
      ["lokasiMcu", "Klinik Pratama Jaya"], ["lokasiMcu", "Rumah Sakit Siloam"], ["lokasiMcu", "Lab Klinik Prodia"]
    ];
    defaults.forEach(row => masterSheet.appendRow(row));
  }

  // 3. Sheet Karyawan
  let karyawanSheet = ss.getSheetByName("Karyawan");
  if (!karyawanSheet) {
    karyawanSheet = ss.insertSheet("Karyawan");
    karyawanSheet.appendRow([
      "ID Karyawan", "Nama Lengkap", "Jabatan", "Departemen", "Status Kerja", "Status Aktif", "Lokasi Kerja",
      "Atasan Langsung", "Mulai Kontrak Kerja", "Selesai Kontrak Kerja", "Tempat Lahir", "Tanggal Lahir", "Jenis Kelamin",
      "Agama", "Status Pernikahan", "Pendidikan Terakhir", "Bank Pembayaran Gaji", "Email", "No Handphone",
      "No Emergency", "Alamat Lengkap", "No KTP", "Masa Berlaku KTP", "No SIM", "Jenis SIM", "Masa Berlaku SIM",
      "No Rekening Bank", "Tanggal Mulai Kerja"
    ]);

    // Seed initial demo data
    karyawanSheet.appendRow([
      "EMP001", "Farhan Alfaizi", "Lead Developer", "IT", "Tetap", "Aktif", "Head Office Jakarta",
      "Supriadi Hartono", "2024-01-15", "", "Jakarta", "1997-08-20", "Laki-laki",
      "Islam", "Belum Menikah", "S1 Teknik Informatika", "Bank Mandiri", "farhan@company.com", "081234567890",
      "081298765432 (Ibu Kandung)", "Jl. Margonda Raya No. 45, Depok", "3171012345670001", "Seumur Hidup", "970820123456", "A", "2029-08-20",
      "124000123123", "2024-01-15"
    ]);
  }

  // 4. Sheet MCU
  let mcuSheet = ss.getSheetByName("MCU");
  if (!mcuSheet) {
    mcuSheet = ss.insertSheet("MCU");
    mcuSheet.appendRow([
      "ID MCU", "Nama Karyawan", "Tanggal Periksa", "Departemen", "Jabatan", "No Handphone",
      "Tanggal Lahir", "Usia/Umur", "Jenis Kelamin", "Lokasi Periksa", "Golongan Darah", "Hasil MCU", "BPJS Kesehatan", "Catatan Dokter"
    ]);
  }

  // 5. Sheet Gaji (Payroll)
  let gajiSheet = ss.getSheetByName("Gaji");
  if (!gajiSheet) {
    gajiSheet = ss.insertSheet("Gaji");
    gajiSheet.appendRow(["ID Slip", "Nama Karyawan", "Periode", "Gaji Pokok", "Tunjangan", "Potongan", "Total Gaji", "Status"]);
  }

  // 6. Sheet Kamar (GA)
  let kamarSheet = ss.getSheetByName("Kamar");
  if (!kamarSheet) {
    kamarSheet = ss.insertSheet("Kamar");
    kamarSheet.appendRow([
      "ID Kamar", "Nomor Kamar", "Nama Mess", "Kapasitas", "Jumlah Penghuni", "Status", "Kondisi Kamar", "Keterangan"
    ]);
  } else {
    // Migrasi jika masih menggunakan header lama
    const firstRow = kamarSheet.getRange(1, 1, 1, kamarSheet.getLastColumn()).getValues()[0].map(h => String(h).trim());
    if (firstRow.indexOf("Nama Mess") === -1) {
      // Hapus kolom Lantai & Tipe Kamar, ubah Mess jadi Nama Mess
      const headers = ["ID Kamar", "Nomor Kamar", "Nama Mess", "Kapasitas", "Jumlah Penghuni", "Status", "Kondisi Kamar", "Keterangan"];
      kamarSheet.clear();
      kamarSheet.appendRow(headers);
    }
  }

  // 8. Sheet Penghuni (GA)
  let penghuniSheet = ss.getSheetByName("Penghuni");
  if (!penghuniSheet) {
    penghuniSheet = ss.insertSheet("Penghuni");
    penghuniSheet.appendRow([
      "ID Penghuni", "NIK", "Nama Karyawan", "Departemen", "Jabatan", "Perusahaan", "Status Karyawan", "Nomor Kamar", "Tanggal Check In", "Tanggal Check Out"
    ]);
  }

  // 9. Sheet InventarisKamar (GA)
  let inventarisSheet = ss.getSheetByName("InventarisKamar");
  if (!inventarisSheet) {
    inventarisSheet = ss.insertSheet("InventarisKamar");
    inventarisSheet.appendRow([
      "ID Inventaris", "Kamar", "Kipas AC Tipe", "Kipas AC Qty", "Kipas AC Kondisi", "Kasur Qty", "Kasur Kondisi", "Lemari Qty", "Lemari Kondisi", "Rak Sepatu Qty", "Rak Sepatu Kondisi", "Keterangan"
    ]);
  } else {
    const firstRow = inventarisSheet.getRange(1, 1, 1, inventarisSheet.getLastColumn()).getValues()[0].map(h => String(h).trim());
    if (firstRow.indexOf("Kipas AC Tipe") === -1) {
      inventarisSheet.clear();
      inventarisSheet.appendRow([
        "ID Inventaris", "Kamar", "Kipas AC Tipe", "Kipas AC Qty", "Kipas AC Kondisi", "Kasur Qty", "Kasur Kondisi", "Lemari Qty", "Lemari Kondisi", "Rak Sepatu Qty", "Rak Sepatu Kondisi", "Keterangan"
      ]);
    }
  }

  // 10. Sheet AsetLaptop (IT)
  let laptopSheet = ss.getSheetByName("AsetLaptop");
  if (!laptopSheet) {
    laptopSheet = ss.insertSheet("AsetLaptop");
    laptopSheet.appendRow([
      "ID Aset", "Asset Tag", "Merk", "Model", "Serial Number", "Processor", "RAM", "Storage", "Sistem Operasi", "Pengguna", "Departemen", "Site", "Lokasi", "Tanggal Pembelian", "Masa Garansi", "Kondisi", "Status Asset", "Catatan"
    ]);
  }

  // 11. Sheet AsetKomputer (IT)
  let komputerSheet = ss.getSheetByName("AsetKomputer");
  if (!komputerSheet) {
    komputerSheet = ss.insertSheet("AsetKomputer");
    komputerSheet.appendRow([
      "ID Aset", "Asset Tag", "Merk", "Model", "Serial Number", "Processor", "RAM", "Storage", "Sistem Operasi", "Pengguna", "Departemen", "Site", "Lokasi", "Tanggal Pembelian", "Masa Garansi", "Kondisi", "Status Asset", "Catatan"
    ]);
  }

  // 12. Sheet AsetHT (IT)
  let htSheet = ss.getSheetByName("AsetHT");
  if (!htSheet) {
    htSheet = ss.insertSheet("AsetHT");
    htSheet.appendRow([
      "ID Aset", "Asset Tag", "Merk", "Model", "Serial Number", "Band", "Mode", "Pengguna", "Departemen", "Site", "Lokasi", "Tanggal Pembelian", "Masa Garansi", "Kondisi", "Status Asset", "Catatan"
    ]);
  }

  // 13. Sheet AsetRIG (IT)
  let rigSheet = ss.getSheetByName("AsetRIG");
  if (!rigSheet) {
    rigSheet = ss.insertSheet("AsetRIG");
    rigSheet.appendRow([
      "ID Aset", "Asset Tag", "Merk", "Model", "Serial Number", "Band", "Mode", "Unit Kendaraan", "Departemen", "Site", "Lokasi", "Tanggal Pembelian", "Kondisi", "Status Asset", "Catatan"
    ]);
  }

  // 14. Sheet AsetPerangkat (IT)
  let perangkatSheet = ss.getSheetByName("AsetPerangkat");
  if (!perangkatSheet) {
    perangkatSheet = ss.insertSheet("AsetPerangkat");
    perangkatSheet.appendRow([
      "ID Aset", "Asset Tag", "Jenis Perangkat", "Merk", "Model", "Serial Number", "Departemen", "Site", "Lokasi", "Tanggal Pembelian", "Kondisi", "Status Asset", "Catatan"
    ]);
  }

  // 15. Sheet DataUnitDT (PRODUKSI)
  let unitdtSheet = ss.getSheetByName("DataUnitDT");
  if (!unitdtSheet) {
    unitdtSheet = ss.insertSheet("DataUnitDT");
    unitdtSheet.appendRow([
      "ID Unit", "Nomor Unit", "Jenis Unit", "Merk", "Model", "Tahun Pembelian", "Kapasitas", "Jumlah Roda", "Status"
    ]);
  } else {
    const firstRow = unitdtSheet.getRange(1, 1, 1, unitdtSheet.getLastColumn()).getValues()[0].map(h => String(h).trim());
    if (firstRow.indexOf("Jumlah Roda") === -1) {
      unitdtSheet.insertColumnAfter(7);
      unitdtSheet.getRange(1, 8).setValue("Jumlah Roda");
    }
  }

  // 16. Sheet ProduksiHauling (PRODUKSI)
  let haulingSheet = ss.getSheetByName("ProduksiHauling");
  if (!haulingSheet) {
    haulingSheet = ss.insertSheet("ProduksiHauling");
    haulingSheet.appendRow([
      "ID Hauling", "Tanggal Input Data", "Tanggal Start", "Jam Start", "Tanggal Selesai", "Jam Selesai",
      "Shift", "Nomor Unit", "Jenis DT", "Merk Unit", "Jumlah Roda", "Driver", "Loading Point", "Dumping Point", "Lokasi PIT", "Lokasi Site",
      "Jumlah Ritase", "Volume (Ton)", "Timbangan Kosongan (TARA)", "Timbangan", "Netto (Kg)", "Netto (Ton)",
      "HM Awal", "HM Akhir", "Pengawas", "Keterangan"
    ]);
  } else {
    // Self-migrate headers if they are still using the old columns
    const firstRow = haulingSheet.getRange(1, 1, 1, haulingSheet.getLastColumn()).getValues()[0].map(h => String(h).trim());
    if (firstRow.indexOf("Tanggal Input Data") === -1) {
      haulingSheet.clear();
      haulingSheet.appendRow([
        "ID Hauling", "Tanggal Input Data", "Tanggal Start", "Jam Start", "Tanggal Selesai", "Jam Selesai",
        "Shift", "Nomor Unit", "Jenis DT", "Merk Unit", "Jumlah Roda", "Driver", "Loading Point", "Dumping Point", "Lokasi PIT", "Lokasi Site",
        "Jumlah Ritase", "Volume (Ton)", "Timbangan Kosongan (TARA)", "Timbangan", "Netto (Kg)", "Netto (Ton)",
        "HM Awal", "HM Akhir", "Pengawas", "Keterangan"
      ]);
    }
    // Migrate: add "Merk Unit" column if missing
    if (firstRow.indexOf("Merk Unit") === -1) {
      const colIdx = firstRow.indexOf("Jenis DT") + 1;
      haulingSheet.insertColumnAfter(colIdx);
      haulingSheet.getRange(1, colIdx + 1).setValue("Merk Unit");
      if (haulingSheet.getLastRow() > 1) {
        haulingSheet.getRange(2, colIdx + 1, haulingSheet.getLastRow() - 1, 1).setValue("");
      }
    }
    // Migrate: add "Jumlah Roda" column if missing
    if (firstRow.indexOf("Jumlah Roda") === -1) {
      const colIdx = (firstRow.indexOf("Merk Unit") !== -1 ? firstRow.indexOf("Merk Unit") : firstRow.indexOf("Jenis DT")) + 1;
      haulingSheet.insertColumnAfter(colIdx);
      haulingSheet.getRange(1, colIdx + 1).setValue("Jumlah Roda");
      // Fill existing rows with empty value
      if (haulingSheet.getLastRow() > 1) {
        haulingSheet.getRange(2, colIdx + 1, haulingSheet.getLastRow() - 1, 1).setValue("");
      }
    }
  }



  // 18. Sheet MasterDataProduksi (MASTER DATA PRODUKSI)
  let masterProduksiSheet = ss.getSheetByName("MasterDataProduksi");
  if (!masterProduksiSheet) {
    masterProduksiSheet = ss.insertSheet("MasterDataProduksi");
    masterProduksiSheet.appendRow(["Category", "Value"]);
    // Default values
    const defaultsProd = [
      ["lokasiPit", "PIT Merak"], ["lokasiPit", "PIT Elang"],
      ["lokasiSite", "Site Lati"], ["lokasiSite", "Site Binungan"],
      ["loadingPoint", "Loading Point A"], ["loadingPoint", "Loading Point B"],
      ["dumpingPoint", "Dumping Point A"], ["dumpingPoint", "Dumping Point B"],
      ["timbangan", "Weighbridge 01"], ["timbangan", "Weighbridge 02"],
      ["jenisUnit", "DT (Dump Truck)"], ["jenisUnit", "LV (Light Vehicle)"],
      ["jenisUnit", "Excavator"], ["jenisUnit", "Bulldozer"],
      ["jenisUnit", "Grader"], ["jenisUnit", "Wheel Loader"],
      ["jenisUnit", "Compactor"], ["jenisUnit", "Fuel Truck"],
      ["jenisUnit", "Water Truck"], ["jenisUnit", "Lainnya"]
    ];
    defaultsProd.forEach(row => masterProduksiSheet.appendRow(row));
  }

  // 19. Sheet MasterBarang (LOGISTIK)
  let masterBarangSheet = ss.getSheetByName("MasterBarang");
  if (!masterBarangSheet) {
    masterBarangSheet = ss.insertSheet("MasterBarang");
    masterBarangSheet.appendRow([
      "ID Barang", "Kode Barang", "Nama Barang", "Kategori", "Merk", "Satuan", "Minimum Stok", "Stok Saat Ini", "Lokasi Penyimpanan", "Keterangan"
    ]);
  } else {
    const firstRow = masterBarangSheet.getRange(1, 1, 1, masterBarangSheet.getLastColumn()).getValues()[0].map(h => String(h).trim());
    if (firstRow.indexOf("Stok Saat Ini") === -1) {
      masterBarangSheet.insertColumnAfter(7);
      masterBarangSheet.getRange(1, 8).setValue("Stok Saat Ini");
    }
  }

  // 20. Sheet Supplier (LOGISTIK)
  let supplierSheet = ss.getSheetByName("Supplier");
  if (!supplierSheet) {
    supplierSheet = ss.insertSheet("Supplier");
    supplierSheet.appendRow([
      "ID Supplier", "Nama Supplier", "PIC", "Nomor Telepon", "Alamat", "Status"
    ]);
  }

  // 21. Sheet BarangMasuk (LOGISTIK)
  let barangMasukSheet = ss.getSheetByName("BarangMasuk");
  if (!barangMasukSheet) {
    barangMasukSheet = ss.insertSheet("BarangMasuk");
    barangMasukSheet.appendRow([
      "ID Transaksi", "Tanggal", "No Transaksi", "Kode Barang", "Nama Barang", "Jumlah", "Satuan", "Supplier", "Penerima", "Keterangan"
    ]);
  }

  // 22. Sheet BarangKeluar (LOGISTIK)
  let barangKeluarSheet = ss.getSheetByName("BarangKeluar");
  if (!barangKeluarSheet) {
    barangKeluarSheet = ss.insertSheet("BarangKeluar");
    barangKeluarSheet.appendRow([
      "ID Transaksi", "Tanggal", "No Transaksi", "Kode Barang", "Nama Barang", "Jumlah", "Satuan", "Departemen", "Jabatan Penerima", "Penerima", "Keperluan", "Keterangan"
    ]);
  }

  // 23. Sheet StockOpname (LOGISTIK)
  let stockOpnameSheet = ss.getSheetByName("StockOpname");
  if (!stockOpnameSheet) {
    stockOpnameSheet = ss.insertSheet("StockOpname");
    stockOpnameSheet.appendRow([
      "ID Opname", "Tanggal", "Kode Barang", "Nama Barang", "Stok Sistem", "Stok Fisik", "Selisih", "Satuan", "Petugas", "Keterangan"
    ]);
  }

  // 24. Sheet RoleAkses
  let roleAksesSheet = ss.getSheetByName("RoleAkses");
  if (!roleAksesSheet) {
    roleAksesSheet = ss.insertSheet("RoleAkses");
    roleAksesSheet.appendRow(["Role", "MenuAccess"]);
    const allMenus = ["dashboard", "karyawan", "produksi", "logistik", "plant", "she", "pengaturan"];
    roleAksesSheet.appendRow(["Admin", JSON.stringify(allMenus)]);
    roleAksesSheet.appendRow(["Staff", JSON.stringify(["dashboard", "pengaturan"])]);
  }

  return { success: true, message: "Database HRIS Core v5 berhasil di-setup!" };
}

// ================= SECURITY & SESSION HELPERS =================
function hashPassword(password) {
  const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, password, Utilities.Charset.UTF_8);
  let hash = "";
  for (let i = 0; i < digest.length; i++) {
    let byteVal = digest[i];
    if (byteVal < 0) byteVal += 256;
    let byteString = byteVal.toString(16);
    if (byteString.length == 1) byteString = "0" + byteString;
    hash += byteString;
  }
  return hash;
}

function loginUser(usernameInput, password) {
  const inputVal = (usernameInput || "").toLowerCase().trim();

  if (checkLoginRateLimit(inputVal)) {
    return { success: false, message: "Terlalu banyak kegagalan login. Akun Anda dikunci selama 15 menit." };
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Users");
  if (!sheet) return { success: false, message: "Sheet Users tidak ditemukan." };

  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return { success: false, message: "Data user kosong." };

  const headers = data[0].map(h => String(h).trim());
  const uIdx = headers.indexOf("Username") !== -1 ? headers.indexOf("Username") : 0;
  const pHashIdx = headers.indexOf("PasswordHash") !== -1 ? headers.indexOf("PasswordHash") : 1;
  const fIdx = headers.indexOf("Fullname") !== -1 ? headers.indexOf("Fullname") : 2;
  const rIdx = headers.indexOf("Role") !== -1 ? headers.indexOf("Role") : 3;
  const eIdx = headers.indexOf("ID Karyawan") !== -1 ? headers.indexOf("ID Karyawan") : 4;

  let userFound = null;
  const hash = hashPassword(password);

  for (let i = 1; i < data.length; i++) {
    const rowUser = String(data[i][uIdx] || "").toLowerCase().trim();
    const rowEmpId = eIdx !== -1 ? String(data[i][eIdx] || "").toLowerCase().trim() : "";

    if ((rowUser === inputVal || (rowEmpId && rowEmpId === inputVal)) && data[i][pHashIdx] === hash) {
      userFound = {
        username: data[i][uIdx],
        fullname: data[i][fIdx],
        role: data[i][rIdx],
        idKaryawan: eIdx !== -1 ? data[i][eIdx] : ""
      };
      break;
    }
  }

  if (userFound) {
    clearLoginFailures(inputVal);
    const token = generateSessionToken(userFound.username);
    return {
      success: true,
      user: userFound,
      token: token
    };
  } else {
    recordLoginFailure(inputVal);
    return { success: false, message: "Username / ID Karyawan atau password Anda salah." };
  }
}

function generateSessionToken(username) {
  const token = Utilities.getUuid();
  const cache = CacheService.getScriptCache();
  // Simpan username di cache dengan token sebagai key selama 8 jam (28800 detik)
  cache.put(token, username, 28800);
  return token;
}

function validateSession(token) {
  if (!token) return { valid: false, message: "Sesi tidak ditemukan / hampa." };
  const cache = CacheService.getScriptCache();
  const username = cache.get(token);
  if (!username) {
    return { valid: false, message: "Sesi Anda telah berakhir. Harap login ulang." };
  }
  return { valid: true, username: username };
}

function destroySession(token) {
  if (token) {
    const cache = CacheService.getScriptCache();
    cache.remove(token);
  }
  return { success: true };
}

function checkLoginRateLimit(username) {
  const cache = CacheService.getScriptCache();
  const lock = cache.get("lock_" + username);
  return !!lock;
}

function recordLoginFailure(username) {
  const cache = CacheService.getScriptCache();
  let failures = parseInt(cache.get("fail_" + username)) || 0;
  failures++;
  if (failures >= 5) {
    cache.put("lock_" + username, "locked", 900); // Kunci 15 menit (900 dtk)
    cache.remove("fail_" + username);
  } else {
    cache.put("fail_" + username, failures.toString(), 900);
  }
}

function clearLoginFailures(username) {
  const cache = CacheService.getScriptCache();
  cache.remove("fail_" + username);
  cache.remove("lock_" + username);
}

// ================= USER MANAGEMENT (DATA USER) =================
function getAllUsers(token) {
  const session = validateSession(token);
  if (!session.valid) return { success: false, message: session.message, sessionExpired: true };

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Users");
    if (!sheet) return { success: false, users: [] };

    const data = sheet.getDataRange().getValues();
    const headers = data[0].map(h => String(h).trim());
    const uIdx = headers.indexOf("Username");
    const fIdx = headers.indexOf("Fullname");
    const rIdx = headers.indexOf("Role");
    const eIdx = headers.indexOf("ID Karyawan");

    let users = [];
    for (let i = 1; i < data.length; i++) {
      users.push({
        username: data[i][uIdx] || "",
        fullname: data[i][fIdx] || "",
        role: data[i][rIdx] || "",
        idKaryawan: eIdx !== -1 ? (data[i][eIdx] || "") : ""
      });
    }
    return { success: true, users: users };
  } catch (err) {
    return { success: false, message: err.message };
  }
}

function saveUser(token, payload) {
  const session = validateSession(token);
  if (!session.valid) return { success: false, message: session.message, sessionExpired: true };

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Users");
    if (!sheet) return { success: false, message: "Sheet Users tidak ditemukan." };

    const data = sheet.getDataRange().getValues();
    const headers = data[0].map(h => String(h).trim());
    const uIdx = headers.indexOf("Username");
    const pHashIdx = headers.indexOf("PasswordHash");
    const fIdx = headers.indexOf("Fullname");
    const rIdx = headers.indexOf("Role");
    const eIdx = headers.indexOf("ID Karyawan");

    const username = sanitizeInput(payload.username, 50).toLowerCase().trim();
    const originalUsername = payload.originalUsername ? sanitizeInput(payload.originalUsername, 50).toLowerCase().trim() : "";
    const fullname = sanitizeInput(payload.fullname, 100);
    const role = sanitizeInput(payload.role, 20) || "Staff";
    const idKaryawan = sanitizeInput(payload.idKaryawan, 20);
    const password = payload.password || "";

    if (!username) return { success: false, message: "Username harus diisi." };

    let targetRowIdx = -1;
    let isNew = true;

    if (originalUsername) {
      for (let i = 1; i < data.length; i++) {
        const existingUser = String(data[i][uIdx] || "").toLowerCase().trim();
        if (existingUser === originalUsername) {
          targetRowIdx = i + 1;
          isNew = false;
          break;
        }
      }
      if (isNew) {
        return { success: false, message: "User asal \"" + originalUsername + "\" tidak ditemukan." };
      }
    }

    // Conflict check: ensure new username is not used by another row
    for (let i = 1; i < data.length; i++) {
      const existingUser = String(data[i][uIdx] || "").toLowerCase().trim();
      const currentRow = i + 1;
      if (existingUser === username) {
        if (isNew || currentRow !== targetRowIdx) {
          return { success: false, message: "Username \"" + username + "\" sudah terdaftar." };
        }
      }
    }

    if (isNew) {
      if (!password) return { success: false, message: "Password harus diisi untuk user baru." };
      sheet.appendRow([username, hashPassword(password), fullname, role, idKaryawan]);
    } else {
      if (uIdx !== -1) sheet.getRange(targetRowIdx, uIdx + 1).setValue(username);
      if (password && password.trim() !== "") {
        if (pHashIdx !== -1) sheet.getRange(targetRowIdx, pHashIdx + 1).setValue(hashPassword(password));
      }
      if (fIdx !== -1) sheet.getRange(targetRowIdx, fIdx + 1).setValue(fullname);
      if (rIdx !== -1) sheet.getRange(targetRowIdx, rIdx + 1).setValue(role);
      if (eIdx !== -1) sheet.getRange(targetRowIdx, eIdx + 1).setValue(idKaryawan);
    }

    return { success: true };
  } catch (err) {
    return { success: false, message: err.message };
  } finally {
    lock.releaseLock();
  }
}

function deleteUser(token, username) {
  const session = validateSession(token);
  if (!session.valid) return { success: false, message: session.message, sessionExpired: true };

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Users");
    if (!sheet) return { success: false, message: "Sheet Users tidak ditemukan." };

    const data = sheet.getDataRange().getValues();

    // Prevent deleting yourself
    if (session.username.toLowerCase() === username.toLowerCase()) {
      return { success: false, message: "Tidak dapat menghapus akun Anda sendiri." };
    }

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).toLowerCase() === username.toLowerCase()) {
        sheet.deleteRow(i + 1);
        return { success: true };
      }
    }
    return { success: false, message: "User tidak ditemukan." };
  } catch (err) {
    return { success: false, message: err.message };
  } finally {
    lock.releaseLock();
  }
}

// ================= ROLE ACCESS MANAGEMENT =================
function getRoleAccess(token) {
  const session = validateSession(token);
  if (!session.valid) return { success: false, message: session.message, sessionExpired: true };

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("RoleAkses");
    let access = {};
    if (sheet) {
      const data = sheet.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        const role = String(data[i][0] || "").trim();
        let menus = [];
        try { menus = JSON.parse(data[i][1] || "[]"); } catch (e) { menus = []; }
        access[role] = menus;
      }
    }
    return { success: true, access: access };
  } catch (err) {
    return { success: false, message: err.message };
  }
}

function saveRoleAccess(token, payload) {
  const session = validateSession(token);
  if (!session.valid) return { success: false, message: session.message, sessionExpired: true };

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("RoleAkses");
    if (!sheet) return { success: false, message: "Sheet RoleAkses tidak ditemukan." };

    const role = sanitizeInput(payload.role, 50);
    const menus = payload.menus || [];
    if (!role) return { success: false, message: "Role harus diisi." };

    const originalRole = payload.originalRole ? sanitizeInput(payload.originalRole, 50) : null;
    const data = sheet.getDataRange().getValues();

    if (originalRole && originalRole !== role) {
      // Rename: find original role row and update both name and menus
      let found = false;
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][0]).trim() === originalRole) {
          sheet.getRange(i + 1, 1).setValue(role);
          sheet.getRange(i + 1, 2).setValue(JSON.stringify(menus));
          found = true;
          break;
        }
      }
      if (!found) return { success: false, message: "Role \"" + originalRole + "\" tidak ditemukan." };
    } else {
      // Normal update or add
      let found = false;
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][0]).trim() === role) {
          sheet.getRange(i + 1, 2).setValue(JSON.stringify(menus));
          found = true;
          break;
        }
      }
      if (!found) {
        sheet.appendRow([role, JSON.stringify(menus)]);
      }
    }
    return { success: true };
  } catch (err) {
    return { success: false, message: err.message };
  } finally {
    lock.releaseLock();
  }
}

function deleteRole(token, payload) {
  const session = validateSession(token);
  if (!session.valid) return { success: false, message: session.message, sessionExpired: true };

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("RoleAkses");
    if (!sheet) return { success: false, message: "Sheet RoleAkses tidak ditemukan." };

    const role = sanitizeInput(payload.role, 50);
    if (!role) return { success: false, message: "Role harus diisi." };

    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === role) {
        sheet.deleteRow(i + 1);
        return { success: true };
      }
    }
    return { success: false, message: "Role \"" + role + "\" tidak ditemukan." };
  } catch (err) {
    return { success: false, message: err.message };
  } finally {
    lock.releaseLock();
  }
}

// ================= BACKEND HELPERS & SANITIZATION =================
function sanitizeInput(value, maxLength) {
  if (typeof value !== "string") value = String(value || "");
  value = value.replace(/<\/?[^>]+(>|$)/g, "").trim(); // Strip HTML tags
  if (maxLength && value.length > maxLength) {
    value = value.substring(0, maxLength);
  }
  return value;
}

function getColumnMap(sheet) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  let map = {};
  headers.forEach((h, index) => {
    map[h.trim()] = index + 1;
  });
  return map;
}

function safeDateFormat(val) {
  if (!val) return "";
  if (val instanceof Date) {
    return Utilities.formatDate(val, Session.getScriptTimeZone(), "yyyy-MM-dd");
  }
  return String(val);
}

function safeTimeFormat(val) {
  if (!val) return "";
  if (val instanceof Date) {
    const hours = String(val.getHours()).padStart(2, "0");
    const minutes = String(val.getMinutes()).padStart(2, "0");
    return hours + ":" + minutes;
  }
  const str = String(val).trim();
  const match = str.match(/(\d{2}):(\d{2})/);
  if (match) return match[0];
  return str;
}

function sanitizeDataForSerialization(obj) {
  if (obj === null || obj === undefined) return "";
  if (Array.isArray(obj)) {
    return obj.map(item => sanitizeDataForSerialization(item));
  }
  if (obj instanceof Date) {
    if (obj.getFullYear() === 1899) {
      return safeTimeFormat(obj);
    }
    return safeDateFormat(obj);
  }
  if (typeof obj === "object") {
    let cleanObj = {};
    for (let key in obj) {
      if (obj.hasOwnProperty(key)) {
        let val = obj[key];
        if (val === undefined || val === null) {
          cleanObj[key] = "";
        } else if (typeof val === "object") {
          cleanObj[key] = sanitizeDataForSerialization(val);
        } else {
          cleanObj[key] = val;
        }
      }
    }
    return cleanObj;
  }
  return obj;
}

// ================= SYNC CORE DATABASE IN ONE-CALL =================
function getInitialData(token) {
  try {
    const session = validateSession(token);
    if (!session.valid) return { success: false, message: session.message, sessionExpired: true };

    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // 1. Fetch Master Data
    let master = {};
    try {
      const masterSheet = ss.getSheetByName("MasterData");
      if (masterSheet) {
        const masterData = masterSheet.getDataRange().getValues();
        for (let i = 1; i < masterData.length; i++) {
          let cat = masterData[i][0];
          let val = masterData[i][1];
          if (!cat) continue;
          if (!master[cat]) {
            master[cat] = [];
          }
          if (val) {
            master[cat].push(val);
          }
        }
      }
    } catch (e) { }
    const standardCats = ["departemen", "jabatan", "lokasiKerja", "statusKaryawan", "pendidikan", "agama", "bank", "golonganDarah", "lokasiMcu"];
    standardCats.forEach(c => {
      if (!master[c]) master[c] = [];
    });

    // 2. Fetch Karyawan
    let karyawan = [];
    try {
      const karyawanSheet = ss.getSheetByName("Karyawan");
      if (karyawanSheet) {
        const karyawanData = karyawanSheet.getDataRange().getValues();
        const kMap = getColumnMap(karyawanSheet);
        for (let i = 1; i < karyawanData.length; i++) {
          karyawan.push({
            id: karyawanData[i][kMap["ID Karyawan"] - 1],
            nama: karyawanData[i][kMap["Nama Lengkap"] - 1],
            jabatan: karyawanData[i][kMap["Jabatan"] - 1],
            departemen: karyawanData[i][kMap["Departemen"] - 1],
            statusKerja: karyawanData[i][kMap["Status Kerja"] - 1],
            statusAktif: karyawanData[i][kMap["Status Aktif"] - 1],
            lokasiKerja: karyawanData[i][kMap["Lokasi Kerja"] - 1],
            atasanLangsung: karyawanData[i][kMap["Atasan Langsung"] - 1],
            kontrakMulai: safeDateFormat(karyawanData[i][kMap["Mulai Kontrak Kerja"] - 1]),
            kontrakSelesai: safeDateFormat(karyawanData[i][kMap["Selesai Kontrak Kerja"] - 1]),
            tempatLahir: karyawanData[i][kMap["Tempat Lahir"] - 1],
            tanggalLahir: safeDateFormat(karyawanData[i][kMap["Tanggal Lahir"] - 1]),
            jenisKelamin: karyawanData[i][kMap["Jenis Kelamin"] - 1],
            agama: karyawanData[i][kMap["Agama"] - 1],
            statusPernikahan: karyawanData[i][kMap["Status Pernikahan"] - 1],
            pendidikan: karyawanData[i][kMap["Pendidikan Terakhir"] - 1],
            bank: karyawanData[i][kMap["Bank Pembayaran Gaji"] - 1],
            email: karyawanData[i][kMap["Email"] - 1],
            noHp: karyawanData[i][kMap["No Handphone"] - 1],
            noEmergency: karyawanData[i][kMap["No Emergency"] - 1],
            alamat: karyawanData[i][kMap["Alamat Lengkap"] - 1],
            noKtp: karyawanData[i][kMap["No KTP"] - 1],
            masaKtp: karyawanData[i][kMap["Masa Berlaku KTP"] - 1],
            noSim: karyawanData[i][kMap["No SIM"] - 1],
            jenisSim: karyawanData[i][kMap["Jenis SIM"] - 1],
            masaSim: safeDateFormat(karyawanData[i][kMap["Masa Berlaku SIM"] - 1]),
            noRekening: karyawanData[i][kMap["No Rekening Bank"] - 1],
            tglMulaiKerja: safeDateFormat(karyawanData[i][kMap["Tanggal Mulai Kerja"] - 1])
          });
        }
      }
    } catch (e) { }

    // 3. Fetch MCU
    let mcu = [];
    try {
      const mcuSheet = ss.getSheetByName("MCU");
      if (mcuSheet) {
        const mcuData = mcuSheet.getDataRange().getValues();
        for (let i = 1; i < mcuData.length; i++) {
          mcu.push({
            id: mcuData[i][0],
            nama: mcuData[i][1],
            tanggalMcu: safeDateFormat(mcuData[i][2]),
            departemen: mcuData[i][3],
            jabatan: mcuData[i][4],
            noHp: mcuData[i][5],
            tanggalLahir: safeDateFormat(mcuData[i][6]),
            umur: mcuData[i][7],
            jenisKelamin: mcuData[i][8],
            lokasi: mcuData[i][9],
            golDarah: mcuData[i][10],
            hasil: mcuData[i][11],
            bpjs: mcuData[i][12],
            catatan: mcuData[i][13],
            fileUrl: mcuData[i][14] || ""
          });
        }
      }
    } catch (e) { }

    // 4. Fetch Gaji
    let gaji = [];
    try {
      const gajiSheet = ss.getSheetByName("Gaji");
      if (gajiSheet) {
        const gajiData = gajiSheet.getDataRange().getValues();
        for (let i = 1; i < gajiData.length; i++) {
          gaji.push({
            id: gajiData[i][0],
            nama: gajiData[i][1],
            periode: gajiData[i][2],
            gajiPokok: gajiData[i][3],
            uangKehadiran: gajiData[i][4],
            overTime: gajiData[i][5],
            potongan: gajiData[i][6],
            total: gajiData[i][7],
            status: gajiData[i][8]
          });
        }
      }
    } catch (e) { }

    // 5. Fetch Kamar (GA)
    let kamar = [];
    try {
      const kamarSheet = ss.getSheetByName("Kamar");
      if (kamarSheet) {
        const kamarData = kamarSheet.getDataRange().getValues();
        const kMap = getColumnMap(kamarSheet);
        const getK = (row, header) => {
          const idx = kMap[header];
          return (idx !== undefined && idx > 0 && idx <= row.length) ? row[idx - 1] : "";
        };
        for (let i = 1; i < kamarData.length; i++) {
          kamar.push({
            id: getK(kamarData[i], "ID Kamar"),
            nomor: getK(kamarData[i], "Nomor Kamar"),
            nama_mess: getK(kamarData[i], "Nama Mess"),
            kapasitas: getK(kamarData[i], "Kapasitas"),
            penghuni: getK(kamarData[i], "Jumlah Penghuni"),
            status: getK(kamarData[i], "Status"),
            kondisi: getK(kamarData[i], "Kondisi Kamar"),
            keterangan: getK(kamarData[i], "Keterangan")
          });
        }
      }
    } catch (e) {
      // In case sheet is not initialized yet
    }

    // 7. Fetch Penghuni (GA)
    let penghuni = [];
    try {
      const penghuniSheet = ss.getSheetByName("Penghuni");
      if (penghuniSheet) {
        const penghuniData = penghuniSheet.getDataRange().getValues();
        for (let i = 1; i < penghuniData.length; i++) {
          penghuni.push({
            id: penghuniData[i][0],
            nik: penghuniData[i][1],
            nama: penghuniData[i][2],
            departemen: penghuniData[i][3],
            jabatan: penghuniData[i][4],
            perusahaan: penghuniData[i][5],
            statusKaryawan: penghuniData[i][6],
            nomorKamar: penghuniData[i][7],
            tglCheckIn: safeDateFormat(penghuniData[i][8]),
            tglCheckOut: safeDateFormat(penghuniData[i][9])
          });
        }
      }
    } catch (e) { }

    // 8. Fetch InventarisKamar (GA)
    let inventaris = [];
    try {
      const inventarisSheet = ss.getSheetByName("InventarisKamar");
      if (inventarisSheet) {
        const inventarisData = inventarisSheet.getDataRange().getValues();
        for (let i = 1; i < inventarisData.length; i++) {
          inventaris.push({
            id: inventarisData[i][0],
            kamar: inventarisData[i][1],
            kipasAcTipe: inventarisData[i][2],
            kipasAcQty: inventarisData[i][3],
            kipasAcKondisi: inventarisData[i][4],
            kasurQty: inventarisData[i][5],
            kasurKondisi: inventarisData[i][6],
            lemariQty: inventarisData[i][7],
            lemariKondisi: inventarisData[i][8],
            rakSepatuQty: inventarisData[i][9],
            rakSepatuKondisi: inventarisData[i][10],
            keterangan: inventarisData[i][11]
          });
        }
      }
    } catch (e) { }

    // 9. Fetch AsetLaptop (IT)
    let asetLaptop = [];
    try {
      const laptopSheet = ss.getSheetByName("AsetLaptop");
      if (laptopSheet) {
        const laptopData = laptopSheet.getDataRange().getValues();
        for (let i = 1; i < laptopData.length; i++) {
          asetLaptop.push({
            id: laptopData[i][0],
            tag: laptopData[i][1],
            merk: laptopData[i][2],
            model: laptopData[i][3],
            sn: laptopData[i][4],
            processor: laptopData[i][5],
            ram: laptopData[i][6],
            storage: laptopData[i][7],
            os: laptopData[i][8],
            pengguna: laptopData[i][9],
            departemen: laptopData[i][10],
            site: laptopData[i][11],
            lokasi: laptopData[i][12],
            tglBeli: safeDateFormat(laptopData[i][13]),
            garansi: laptopData[i][14],
            kondisi: laptopData[i][15],
            status: laptopData[i][16],
            catatan: laptopData[i][17]
          });
        }
      }
    } catch (e) { }

    // 10. Fetch AsetKomputer (IT)
    let asetKomputer = [];
    try {
      const komputerSheet = ss.getSheetByName("AsetKomputer");
      if (komputerSheet) {
        const komputerData = komputerSheet.getDataRange().getValues();
        for (let i = 1; i < komputerData.length; i++) {
          asetKomputer.push({
            id: komputerData[i][0],
            tag: komputerData[i][1],
            merk: komputerData[i][2],
            model: komputerData[i][3],
            sn: komputerData[i][4],
            processor: komputerData[i][5],
            ram: komputerData[i][6],
            storage: komputerData[i][7],
            os: komputerData[i][8],
            pengguna: komputerData[i][9],
            departemen: komputerData[i][10],
            site: komputerData[i][11],
            lokasi: komputerData[i][12],
            tglBeli: safeDateFormat(komputerData[i][13]),
            garansi: komputerData[i][14],
            kondisi: komputerData[i][15],
            status: komputerData[i][16],
            catatan: komputerData[i][17]
          });
        }
      }
    } catch (e) { }

    // 11. Fetch AsetHT (IT)
    let asetHT = [];
    try {
      const htSheet = ss.getSheetByName("AsetHT");
      if (htSheet) {
        const htData = htSheet.getDataRange().getValues();
        for (let i = 1; i < htData.length; i++) {
          asetHT.push({
            id: htData[i][0],
            tag: htData[i][1],
            merk: htData[i][2],
            model: htData[i][3],
            sn: htData[i][4],
            band: htData[i][5],
            mode: htData[i][6],
            pengguna: htData[i][7],
            departemen: htData[i][8],
            site: htData[i][9],
            lokasi: htData[i][10],
            tglBeli: safeDateFormat(htData[i][11]),
            garansi: htData[i][12],
            kondisi: htData[i][13],
            status: htData[i][14],
            catatan: htData[i][15]
          });
        }
      }
    } catch (e) { }

    // 12. Fetch AsetRIG (IT)
    let asetRIG = [];
    try {
      const rigSheet = ss.getSheetByName("AsetRIG");
      if (rigSheet) {
        const rigData = rigSheet.getDataRange().getValues();
        for (let i = 1; i < rigData.length; i++) {
          asetRIG.push({
            id: rigData[i][0],
            tag: rigData[i][1],
            merk: rigData[i][2],
            model: rigData[i][3],
            sn: rigData[i][4],
            band: rigData[i][5],
            mode: rigData[i][6],
            unit: rigData[i][7],
            departemen: rigData[i][8],
            site: rigData[i][9],
            lokasi: rigData[i][10],
            tglBeli: safeDateFormat(rigData[i][11]),
            kondisi: rigData[i][12],
            status: rigData[i][13],
            catatan: rigData[i][14]
          });
        }
      }
    } catch (e) { }

    // 13. Fetch AsetPerangkat (IT)
    let asetPerangkat = [];
    try {
      const perangkatSheet = ss.getSheetByName("AsetPerangkat");
      if (perangkatSheet) {
        const perangkatData = perangkatSheet.getDataRange().getValues();
        for (let i = 1; i < perangkatData.length; i++) {
          asetPerangkat.push({
            id: perangkatData[i][0],
            tag: perangkatData[i][1],
            jenis: perangkatData[i][2],
            merk: perangkatData[i][3],
            model: perangkatData[i][4],
            sn: perangkatData[i][5],
            departemen: perangkatData[i][6],
            site: perangkatData[i][7],
            lokasi: perangkatData[i][8],
            tglBeli: safeDateFormat(perangkatData[i][9]),
            kondisi: perangkatData[i][10],
            status: perangkatData[i][11],
            catatan: perangkatData[i][12]
          });
        }
      }
    } catch (e) { }

    // 14. Fetch DataUnitDT (PRODUKSI)
    let unitdt = [];
    try {
      const unitdtSheet = ss.getSheetByName("DataUnitDT");
      if (unitdtSheet) {
        const unitdtData = unitdtSheet.getDataRange().getValues();
        const uMap = getColumnMap(unitdtSheet);
        const getUnitVal = (row, header, defaultIdx) => {
          const idx = uMap[header];
          if (idx !== undefined && idx > 0 && idx <= row.length) {
            const val = row[idx - 1];
            return (val === undefined || val === null) ? "" : val;
          }
          if (defaultIdx !== undefined && defaultIdx >= 0 && defaultIdx < row.length) {
            const val = row[defaultIdx];
            return (val === undefined || val === null) ? "" : val;
          }
          return "";
        };

        for (let i = 1; i < unitdtData.length; i++) {
          unitdt.push({
            id: getUnitVal(unitdtData[i], "ID Unit", 0),
            nomor: getUnitVal(unitdtData[i], "Nomor Unit", 1),
            jenis: getUnitVal(unitdtData[i], "Jenis Unit", 2),
            merk: getUnitVal(unitdtData[i], "Merk", 3),
            model: getUnitVal(unitdtData[i], "Model", 4),
            tahun: getUnitVal(unitdtData[i], "Tahun Pembelian", 5),
            kapasitas: getUnitVal(unitdtData[i], "Kapasitas", 6),
            jumlahRoda: getUnitVal(unitdtData[i], "Jumlah Roda", 7),
            status: getUnitVal(unitdtData[i], "Status", 8)
          });
        }
      }
    } catch (e) { }

    // 15. Fetch ProduksiHauling (PRODUKSI)
    let produksihauling = [];
    try {
      const haulingSheet = ss.getSheetByName("ProduksiHauling");
      if (haulingSheet) {
        const haulingData = haulingSheet.getDataRange().getValues();
        const hMap = getColumnMap(haulingSheet);
        const getVal = (row, map, header) => {
          const idx = map[header];
          return (idx !== undefined && idx > 0 && idx <= row.length) ? row[idx - 1] : "";
        };
        for (let i = 1; i < haulingData.length; i++) {
          produksihauling.push({
            id: getVal(haulingData[i], hMap, "ID Hauling"),
            tanggal: safeDateFormat(getVal(haulingData[i], hMap, "Tanggal Start") || getVal(haulingData[i], hMap, "Tanggal Input Data")),
            tanggalInput: safeDateFormat(getVal(haulingData[i], hMap, "Tanggal Input Data")),
            tanggalStart: safeDateFormat(getVal(haulingData[i], hMap, "Tanggal Start")),
            jamStart: getVal(haulingData[i], hMap, "Jam Start"),
            tanggalSelesai: safeDateFormat(getVal(haulingData[i], hMap, "Tanggal Selesai")),
            jamSelesai: getVal(haulingData[i], hMap, "Jam Selesai"),
            shift: getVal(haulingData[i], hMap, "Shift"),
            loadingPoint: getVal(haulingData[i], hMap, "Loading Point"),
            dumpingPoint: getVal(haulingData[i], hMap, "Dumping Point"),
            unit: getVal(haulingData[i], hMap, "Nomor Unit"),
            jenisDt: getVal(haulingData[i], hMap, "Jenis DT"),
            merkUnit: getVal(haulingData[i], hMap, "Merk Unit"),
            jumlahRoda: getVal(haulingData[i], hMap, "Jumlah Roda"),
            driver: getVal(haulingData[i], hMap, "Driver"),
            ritase: getVal(haulingData[i], hMap, "Jumlah Ritase"),
            lokasiPit: getVal(haulingData[i], hMap, "Lokasi PIT"),
            lokasiSite: getVal(haulingData[i], hMap, "Lokasi Site"),
            volume: getVal(haulingData[i], hMap, "Volume (Ton)"),
            tara: getVal(haulingData[i], hMap, "Timbangan Kosongan (TARA)"),
            timbangan: getVal(haulingData[i], hMap, "Timbangan"),
            nettoKg: getVal(haulingData[i], hMap, "Netto (Kg)"),
            nettoTon: getVal(haulingData[i], hMap, "Netto (Ton)"),
            hmAwal: getVal(haulingData[i], hMap, "HM Awal"),
            hmAkhir: getVal(haulingData[i], hMap, "HM Akhir"),
            pengawas: getVal(haulingData[i], hMap, "Pengawas"),
            keterangan: getVal(haulingData[i], hMap, "Keterangan")
          });
        }
      }
    } catch (e) { }



    // 17. Fetch MasterDataProduksi
    let masterProduksi = {};
    try {
      const mpSheet = ss.getSheetByName("MasterDataProduksi");
      if (mpSheet) {
        const mpData = mpSheet.getDataRange().getValues();
        for (let i = 1; i < mpData.length; i++) {
          let cat = mpData[i][0];
          let val = mpData[i][1];
          if (!cat) continue;
          if (!masterProduksi[cat]) {
            masterProduksi[cat] = [];
          }
          if (val) {
            masterProduksi[cat].push(val);
          }
        }
      }
    } catch (e) { }
    // Ensure standard categories exist
    const standardProdCats = ["loadingPoint", "dumpingPoint", "lokasiPit", "lokasiSite", "timbangan", "jenisUnit"];
    standardProdCats.forEach(c => {
      if (!masterProduksi[c]) masterProduksi[c] = [];
    });

    // 18. Fetch MasterBarang (LOGISTIK)
    let masterBarang = [];
    try {
      const mbSheet = ss.getSheetByName("MasterBarang");
      if (mbSheet) {
        const mbData = mbSheet.getDataRange().getValues();
        const headers = mbData[0].map(h => String(h).trim());
        const idIdx = headers.indexOf("ID Barang") !== -1 ? headers.indexOf("ID Barang") : 0;
        const kodeIdx = headers.indexOf("Kode Barang") !== -1 ? headers.indexOf("Kode Barang") : 1;
        const namaIdx = headers.indexOf("Nama Barang") !== -1 ? headers.indexOf("Nama Barang") : 2;
        const katIdx = headers.indexOf("Kategori") !== -1 ? headers.indexOf("Kategori") : 3;
        const merkIdx = headers.indexOf("Merk") !== -1 ? headers.indexOf("Merk") : 4;
        const satuanIdx = headers.indexOf("Satuan") !== -1 ? headers.indexOf("Satuan") : 5;
        const minStokIdx = headers.indexOf("Minimum Stok") !== -1 ? headers.indexOf("Minimum Stok") : 6;
        const stokIdx = headers.indexOf("Stok Saat Ini");
        const lokasiIdx = headers.indexOf("Lokasi Penyimpanan") !== -1 ? headers.indexOf("Lokasi Penyimpanan") : (stokIdx !== -1 ? 8 : 7);
        const ketIdx = headers.indexOf("Keterangan") !== -1 ? headers.indexOf("Keterangan") : (stokIdx !== -1 ? 9 : 8);

        for (let i = 1; i < mbData.length; i++) {
          masterBarang.push({
            id: mbData[i][idIdx],
            kodeBarang: mbData[i][kodeIdx],
            namaBarang: mbData[i][namaIdx],
            kategori: mbData[i][katIdx],
            merk: mbData[i][merkIdx],
            satuan: mbData[i][satuanIdx],
            minStok: mbData[i][minStokIdx] !== "" ? parseFloat(mbData[i][minStokIdx]) : 0,
            stokSaatIni: stokIdx !== -1 && mbData[i][stokIdx] !== "" ? parseFloat(mbData[i][stokIdx]) : 0,
            lokasi: mbData[i][lokasiIdx],
            keterangan: mbData[i][ketIdx]
          });
        }
      }
    } catch (e) { }

    // 19. Fetch Supplier (LOGISTIK)
    let supplier = [];
    try {
      const supSheet = ss.getSheetByName("Supplier");
      if (supSheet) {
        const supData = supSheet.getDataRange().getValues();
        for (let i = 1; i < supData.length; i++) {
          supplier.push({
            id: supData[i][0],
            namaSupplier: supData[i][1],
            pic: supData[i][2],
            telepon: supData[i][3],
            alamat: supData[i][4],
            status: supData[i][5]
          });
        }
      }
    } catch (e) { }

    // 20. Fetch BarangMasuk (LOGISTIK)
    let barangMasuk = [];
    try {
      const bmSheet = ss.getSheetByName("BarangMasuk");
      if (bmSheet) {
        const bmData = bmSheet.getDataRange().getValues();
        for (let i = 1; i < bmData.length; i++) {
          barangMasuk.push({
            id: bmData[i][0],
            tanggal: bmData[i][1] instanceof Date ? bmData[i][1].toISOString().split('T')[0] : String(bmData[i][1]),
            noTransaksi: bmData[i][2],
            kodeBarang: bmData[i][3],
            namaBarang: bmData[i][4],
            jumlah: parseFloat(bmData[i][5]) || 0,
            satuan: bmData[i][6],
            supplier: bmData[i][7],
            penerima: bmData[i][8],
            keterangan: bmData[i][9]
          });
        }
      }
    } catch (e) { }

    // 21. Fetch BarangKeluar (LOGISTIK)
    let barangKeluar = [];
    try {
      const bkSheet = ss.getSheetByName("BarangKeluar");
      if (bkSheet) {
        const bkData = bkSheet.getDataRange().getValues();
        for (let i = 1; i < bkData.length; i++) {
          barangKeluar.push({
            id: bkData[i][0],
            tanggal: bkData[i][1] instanceof Date ? bkData[i][1].toISOString().split('T')[0] : String(bkData[i][1]),
            noTransaksi: bkData[i][2],
            kodeBarang: bkData[i][3],
            namaBarang: bkData[i][4],
            jumlah: parseFloat(bkData[i][5]) || 0,
            satuan: bkData[i][6],
            departemen: bkData[i][7],
            jabatanPenerima: bkData[i][8],
            penerima: bkData[i][9],
            keperluan: bkData[i][10],
            keterangan: bkData[i][11]
          });
        }
      }
    } catch (e) { }

    // 22. Fetch StockOpname (LOGISTIK)
    let stockOpname = [];
    try {
      const soSheet = ss.getSheetByName("StockOpname");
      if (soSheet) {
        const soData = soSheet.getDataRange().getValues();
        for (let i = 1; i < soData.length; i++) {
          stockOpname.push({
            id: soData[i][0],
            tanggal: soData[i][1] instanceof Date ? soData[i][1].toISOString().split('T')[0] : String(soData[i][1]),
            kodeBarang: soData[i][2],
            namaBarang: soData[i][3],
            stokSistem: parseFloat(soData[i][4]) || 0,
            stokFisik: parseFloat(soData[i][5]) || 0,
            selisih: parseFloat(soData[i][6]) || 0,
            satuan: soData[i][7],
            petugas: soData[i][8],
            keterangan: soData[i][9]
          });
        }
      }
    } catch (e) { }

    // 23. Fetch Users (tanpa hash password)
    let users = [];
    try {
      const usersSheet = ss.getSheetByName("Users");
      if (usersSheet) {
        const usersData = usersSheet.getDataRange().getValues();
        const uHeaders = usersData[0].map(h => String(h).trim());
        const uIdx = uHeaders.indexOf("Username");
        const fIdx = uHeaders.indexOf("Fullname");
        const rIdx = uHeaders.indexOf("Role");
        const eIdx = uHeaders.indexOf("ID Karyawan");
        for (let i = 1; i < usersData.length; i++) {
          users.push({
            username: usersData[i][uIdx] || "",
            fullname: usersData[i][fIdx] || "",
            role: usersData[i][rIdx] || "",
            idKaryawan: eIdx !== -1 ? (usersData[i][eIdx] || "") : ""
          });
        }
      }
    } catch (e) { }

    // 24. Fetch RoleAkses
    let roleAccess = {};
    try {
      const raSheet = ss.getSheetByName("RoleAkses");
      if (raSheet) {
        const raData = raSheet.getDataRange().getValues();
        for (let i = 1; i < raData.length; i++) {
          const role = String(raData[i][0] || "").trim();
          let menus = [];
          try { menus = JSON.parse(raData[i][1] || "[]"); } catch (e) { menus = []; }
          roleAccess[role] = menus;
        }
      }
    } catch (e) { }

    const rawData = { success: true, master: master, karyawan: karyawan, mcu: mcu, gaji: gaji, kamar: kamar, penghuni: penghuni, inventaris: inventaris, asetLaptop: asetLaptop, asetKomputer: asetKomputer, asetHT: asetHT, asetRIG: asetRIG, asetPerangkat: asetPerangkat, unitdt: unitdt, produksihauling: produksihauling, masterProduksi: masterProduksi, masterBarang: masterBarang, supplier: supplier, barangMasuk: barangMasuk, barangKeluar: barangKeluar, stockOpname: stockOpname, users: users, roleAccess: roleAccess };
    return sanitizeDataForSerialization(rawData);
  } catch (err) {
    return { success: false, message: "Gagal memuat dataset: " + err.message };
  }
}

// ================= CRUD: KARYAWAN OPERATIONS =================
function saveKaryawan(token, payload) {
  const session = validateSession(token);
  if (!session.valid) return { success: false, message: session.message, sessionExpired: true };

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000); // 10 detik timeout lock konkurensi

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Karyawan");
    const data = sheet.getDataRange().getValues();
    const kMap = getColumnMap(sheet);

    // Sanitasi data sensitif
    payload.nama = sanitizeInput(payload.nama, 100);
    payload.email = sanitizeInput(payload.email, 100);
    payload.noHp = sanitizeInput(payload.noHp, 25);
    payload.noKtp = sanitizeInput(payload.noKtp, 20);

    let rowIdx = -1;
    if (payload.id) {
      // Edit mode
      for (let i = 1; i < data.length; i++) {
        if (data[i][kMap["ID Karyawan"] - 1] === payload.id) {
          rowIdx = i + 1;
          break;
        }
      }
    }

    if (rowIdx === -1) {
      // Add mode - generate new custom incremental ID
      const nextId = "EMP" + String(data.length).padStart(3, "0");
      payload.id = nextId;
      sheet.appendRow([
        payload.id, payload.nama, payload.jabatan, payload.departemen, payload.statusKerja, payload.statusAktif,
        payload.lokasiKerja, payload.atasanLangsung, payload.kontrakMulai, payload.kontrakSelesai, payload.tempatLahir,
        payload.tanggalLahir, payload.jenisKelamin, payload.agama, payload.statusPernikahan, payload.pendidikan,
        payload.bank, payload.email, payload.noHp, payload.noEmergency, payload.alamat, payload.noKtp,
        payload.masaKtp, payload.noSim, payload.jenisSim, payload.masaSim, payload.noRekening, payload.tglMulaiKerja
      ]);
    } else {
      // Update existing row
      sheet.getRange(rowIdx, kMap["Nama Lengkap"]).setValue(payload.nama);
      sheet.getRange(rowIdx, kMap["Jabatan"]).setValue(payload.jabatan);
      sheet.getRange(rowIdx, kMap["Departemen"]).setValue(payload.departemen);
      sheet.getRange(rowIdx, kMap["Status Kerja"]).setValue(payload.statusKerja);
      sheet.getRange(rowIdx, kMap["Status Aktif"]).setValue(payload.statusAktif);
      sheet.getRange(rowIdx, kMap["Lokasi Kerja"]).setValue(payload.lokasiKerja);
      sheet.getRange(rowIdx, kMap["Atasan Langsung"]).setValue(payload.atasanLangsung);
      sheet.getRange(rowIdx, kMap["Mulai Kontrak Kerja"]).setValue(payload.kontrakMulai);
      sheet.getRange(rowIdx, kMap["Selesai Kontrak Kerja"]).setValue(payload.kontrakSelesai);
      sheet.getRange(rowIdx, kMap["Tempat Lahir"]).setValue(payload.tempatLahir);
      sheet.getRange(rowIdx, kMap["Tanggal Lahir"]).setValue(payload.tanggalLahir);
      sheet.getRange(rowIdx, kMap["Jenis Kelamin"]).setValue(payload.jenisKelamin);
      sheet.getRange(rowIdx, kMap["Agama"]).setValue(payload.agama);
      sheet.getRange(rowIdx, kMap["Status Pernikahan"]).setValue(payload.statusPernikahan);
      sheet.getRange(rowIdx, kMap["Pendidikan Terakhir"]).setValue(payload.pendidikan);
      sheet.getRange(rowIdx, kMap["Bank Pembayaran Gaji"]).setValue(payload.bank);
      sheet.getRange(rowIdx, kMap["Email"]).setValue(payload.email);
      sheet.getRange(rowIdx, kMap["No Handphone"]).setValue(payload.noHp);
      sheet.getRange(rowIdx, kMap["No Emergency"]).setValue(payload.noEmergency);
      sheet.getRange(rowIdx, kMap["Alamat Lengkap"]).setValue(payload.alamat);
      sheet.getRange(rowIdx, kMap["No KTP"]).setValue(payload.noKtp);
      sheet.getRange(rowIdx, kMap["Masa Berlaku KTP"]).setValue(payload.masaKtp);
      sheet.getRange(rowIdx, kMap["No SIM"]).setValue(payload.noSim);
      sheet.getRange(rowIdx, kMap["Jenis SIM"]).setValue(payload.jenisSim);
      sheet.getRange(rowIdx, kMap["Masa Berlaku SIM"]).setValue(payload.masaSim);
      sheet.getRange(rowIdx, kMap["No Rekening Bank"]).setValue(payload.noRekening);
      sheet.getRange(rowIdx, kMap["Tanggal Mulai Kerja"]).setValue(payload.tglMulaiKerja);
    }

    return { success: true };
  } catch (err) {
    return { success: false, message: "Gagal menyimpan data karyawan: " + err.message };
  } finally {
    lock.releaseLock();
  }
}

function deleteKaryawan(token, empId) {
  const session = validateSession(token);
  if (!session.valid) return { success: false, message: session.message, sessionExpired: true };

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Karyawan");
    const data = sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === empId) {
        sheet.deleteRow(i + 1);
        return { success: true };
      }
    }
    return { success: false, message: "ID Karyawan tidak ditemukan." };
  } catch (err) {
    return { success: false, message: err.message };
  } finally {
    lock.releaseLock();
  }
}

// ================= CRUD: MCU OPERATIONS =================
function saveMcu(token, payload) {
  const session = validateSession(token);
  if (!session.valid) return { success: false, message: session.message, sessionExpired: true };

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("MCU");

    // Ensure header has the "File Lampiran" column if sheet has less than 15 columns
    if (sheet.getLastColumn() < 15) {
      sheet.getRange(1, 15).setValue("File Lampiran");
    }

    const data = sheet.getDataRange().getValues();

    payload.nama = sanitizeInput(payload.nama, 100);
    payload.lokasi = sanitizeInput(payload.lokasi, 150);

    let fileUrl = payload.fileUrl || "";

    // Upload file to Google Drive if new file data is provided
    if (payload.fileData && payload.fileData.base64) {
      const folder = getOrCreateFolder("data mcu karyawan");
      const ext = payload.fileData.name.split('.').pop();
      const sanitizedName = payload.nama.replace(/[^a-zA-Z0-9 ]/g, "");
      const filename = sanitizedName + " - MCU " + payload.tanggalMcu + "." + ext;

      fileUrl = saveUploadedFile(folder, payload.fileData.base64, filename);
    }

    let rowIdx = -1;
    if (payload.id) {
      for (let i = 1; i < data.length; i++) {
        if (data[i][0] === payload.id) {
          rowIdx = i + 1;
          break;
        }
      }
    }

    if (rowIdx === -1) {
      payload.id = "MCU" + String(data.length).padStart(3, "0");
      sheet.appendRow([
        payload.id, payload.nama, payload.tanggalMcu, payload.departemen, payload.jabatan,
        payload.noHp, payload.tanggalLahir, payload.umur, payload.jenisKelamin, payload.lokasi,
        payload.golDarah, payload.hasil, payload.bpjs, payload.catatan, fileUrl
      ]);
    } else {
      sheet.getRange(rowIdx, 3).setValue(payload.tanggalMcu);
      sheet.getRange(rowIdx, 10).setValue(payload.lokasi);
      sheet.getRange(rowIdx, 11).setValue(payload.golDarah);
      sheet.getRange(rowIdx, 12).setValue(payload.hasil);
      sheet.getRange(rowIdx, 13).setValue(payload.bpjs);
      sheet.getRange(rowIdx, 14).setValue(payload.catatan);
      sheet.getRange(rowIdx, 15).setValue(fileUrl);
    }

    return { success: true };
  } catch (err) {
    return { success: false, message: err.message };
  } finally {
    lock.releaseLock();
  }
}

function deleteMcu(token, mcuId) {
  const session = validateSession(token);
  if (!session.valid) return { success: false, message: session.message, sessionExpired: true };

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("MCU");
    const data = sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === mcuId) {
        sheet.deleteRow(i + 1);
        return { success: true };
      }
    }
    return { success: false, message: "ID MCU tidak cocok." };
  } catch (err) {
    return { success: false, message: err.message };
  } finally {
    lock.releaseLock();
  }
}

// ================= CRUD: PAYROLL OPERATIONS =================
function savePayroll(token, payload) {
  const session = validateSession(token);
  if (!session.valid) return { success: false, message: session.message, sessionExpired: true };

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName("Gaji");
    if (!sheet) {
      sheet = ss.insertSheet("Gaji");
      sheet.appendRow(["ID Slip", "Nama Karyawan", "Periode", "Gaji Pokok", "Tunjangan", "Potongan", "Total Gaji", "Status"]);
    }
    const data = sheet.getDataRange().getValues();

    payload.nama = sanitizeInput(payload.nama, 100);
    payload.periode = sanitizeInput(payload.periode, 50);
    payload.gajiPokok = parseFloat(payload.gajiPokok) || 0;
    payload.uangKehadiran = parseFloat(payload.uangKehadiran) || 0;
    payload.overTime = parseFloat(payload.overTime) || 0;
    payload.potongan = parseFloat(payload.potongan) || 0;
    payload.total = parseFloat(payload.total) || 0;

    let rowIdx = -1;
    if (payload.id) {
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][0]) === String(payload.id)) {
          rowIdx = i + 1;
          break;
        }
      }
    }

    if (rowIdx === -1) {
      const slipId = "PAY" + String(data.length).padStart(3, "0");
      sheet.appendRow([
        slipId, payload.nama, payload.periode, payload.gajiPokok,
        payload.uangKehadiran, payload.overTime, payload.potongan, payload.total, "Dibayar"
      ]);
    } else {
      sheet.getRange(rowIdx, 2).setValue(payload.nama);
      sheet.getRange(rowIdx, 3).setValue(payload.periode);
      sheet.getRange(rowIdx, 4).setValue(payload.gajiPokok);
      sheet.getRange(rowIdx, 5).setValue(payload.uangKehadiran);
      sheet.getRange(rowIdx, 6).setValue(payload.overTime);
      sheet.getRange(rowIdx, 7).setValue(payload.potongan);
      sheet.getRange(rowIdx, 8).setValue(payload.total);
    }

    return { success: true };
  } catch (err) {
    return { success: false, message: err.message };
  } finally {
    lock.releaseLock();
  }
}

function deletePayroll(token, slipId) {
  const session = validateSession(token);
  if (!session.valid) return { success: false, message: session.message, sessionExpired: true };

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Gaji");
    if (!sheet) return { success: false, message: "Sheet Gaji tidak ditemukan." };
    const data = sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(slipId)) {
        sheet.deleteRow(i + 1);
        return { success: true };
      }
    }
    return { success: false, message: "ID Slip Gaji tidak ditemukan." };
  } catch (err) {
    return { success: false, message: err.message };
  } finally {
    lock.releaseLock();
  }
}

// ================= CRUD: DATA MASTER OPERATIONS =================
function saveMasterItem(token, category, indexVal, value) {
  const session = validateSession(token);
  if (!session.valid) return { success: false, message: session.message, sessionExpired: true };

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("MasterData");
    const data = sheet.getDataRange().getValues();

    value = sanitizeInput(value, 100);

    if (indexVal === "") {
      // Add new item
      sheet.appendRow([category, value]);
    } else {
      // Edit item - Cari letak baris yang sesuai berdasarkan urutan/indeks kategori
      let count = 0;
      let targetRow = -1;
      let targetIdx = parseInt(indexVal);

      for (let i = 1; i < data.length; i++) {
        if (data[i][0] === category) {
          if (count === targetIdx) {
            targetRow = i + 1;
            break;
          }
          count++;
        }
      }

      if (targetRow !== -1) {
        sheet.getRange(targetRow, 2).setValue(value);
      } else {
        return { success: false, message: "Urutan entitas master tidak ditemukan." };
      }
    }
    return { success: true };
  } catch (err) {
    return { success: false, message: err.message };
  } finally {
    lock.releaseLock();
  }
}

function deleteMasterItem(token, category, index) {
  const session = validateSession(token);
  if (!session.valid) return { success: false, message: session.message, sessionExpired: true };

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("MasterData");
    const data = sheet.getDataRange().getValues();

    let count = 0;
    let targetRow = -1;
    let targetIdx = parseInt(index);

    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === category) {
        if (count === targetIdx) {
          targetRow = i + 1;
          break;
        }
        count++;
      }
    }

    if (targetRow !== -1) {
      sheet.deleteRow(targetRow);
      return { success: true };
    }
    return { success: false, message: "Target entitas master tidak valid." };
  } catch (err) {
    return { success: false, message: err.message };
  } finally {
    lock.releaseLock();
  }
}

// ================= MASTER CATEGORY OPERATIONS =================
function renameMasterCategory(token, oldCategory, newCategory) {
  const session = validateSession(token);
  if (!session.valid) return { success: false, message: session.message, sessionExpired: true };

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("MasterData");
    const data = sheet.getDataRange().getValues();

    newCategory = sanitizeInput(newCategory, 100);

    let updatedCount = 0;
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === oldCategory) {
        sheet.getRange(i + 1, 1).setValue(newCategory);
        updatedCount++;
      }
    }
    return { success: true, updatedCount: updatedCount };
  } catch (err) {
    return { success: false, message: err.message };
  } finally {
    lock.releaseLock();
  }
}

function deleteMasterCategory(token, category) {
  const session = validateSession(token);
  if (!session.valid) return { success: false, message: session.message, sessionExpired: true };

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("MasterData");
    const data = sheet.getDataRange().getValues();

    // Loop backwards so indices don't shift unpredictably
    for (let i = data.length - 1; i >= 1; i--) {
      if (data[i][0] === category) {
        sheet.deleteRow(i + 1);
      }
    }
    return { success: true };
  } catch (err) {
    return { success: false, message: err.message };
  } finally {
    lock.releaseLock();
  }
}

function saveMasterCategory(token, category, categoryLabel) {
  return saveMasterItem(token, category, "", categoryLabel);
}

// ================= USER ACCOUNT SETTINGS =================
function updateProfile(token, fullName, newPassword) {
  const session = validateSession(token);
  if (!session.valid) return { success: false, message: session.message, sessionExpired: true };

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Users");
    const data = sheet.getDataRange().getValues();

    fullName = sanitizeInput(fullName, 100);

    for (let i = 1; i < data.length; i++) {
      if (data[i][0].toLowerCase() === session.username.toLowerCase()) {
        const row = i + 1;
        sheet.getRange(row, 3).setValue(fullName);
        if (newPassword && newPassword.trim() !== "") {
          sheet.getRange(row, 2).setValue(hashPassword(newPassword.trim()));
        }

        return {
          success: true,
          user: {
            username: session.username,
            fullname: fullName,
            role: data[i][3]
          }
        };
      }
    }
    return { success: false, message: "User tidak ditemukan di database." };
  } catch (err) {
    return { success: false, message: err.message };
  } finally {
    lock.releaseLock();
  }
}

// ================= DRIVE & FILE UPLOAD HELPERS =================
function getOrCreateFolder(folderName) {
  const folders = DriveApp.getFoldersByName(folderName);
  if (folders.hasNext()) {
    return folders.next();
  }
  return DriveApp.createFolder(folderName);
}

function saveUploadedFile(folder, base64Data, filename) {
  const parts = base64Data.split(',');
  const contentType = parts[0].split(':')[1].split(';')[0];
  const byteData = Utilities.base64Decode(parts[1]);
  const blob = Utilities.newBlob(byteData, contentType, filename);
  const file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return file.getUrl();
}

// ================= CRUD: KAMAR OPERATIONS =================
function saveKamar(token, payload) {
  const session = validateSession(token);
  if (!session.valid) return { success: false, message: session.message, sessionExpired: true };

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName("Kamar");
    if (!sheet) {
      sheet = ss.insertSheet("Kamar");
      sheet.appendRow([
        "ID Kamar", "Nomor Kamar", "Nama Mess", "Kapasitas", "Jumlah Penghuni", "Status", "Kondisi Kamar", "Keterangan"
      ]);
    }
    const data = sheet.getDataRange().getValues();

    payload.nomor = sanitizeInput(payload.nomor, 50);
    payload.nama_mess = sanitizeInput(payload.nama_mess, 100);
    payload.status = sanitizeInput(payload.status, 20);
    payload.kondisi = sanitizeInput(payload.kondisi, 100);
    payload.keterangan = sanitizeInput(payload.keterangan, 500);

    let rowIdx = -1;
    if (payload.id) {
      for (let i = 1; i < data.length; i++) {
        if (data[i][0] === payload.id) {
          rowIdx = i + 1;
          break;
        }
      }
    }

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => String(h).trim());
    const rowValues = headers.map(header => {
      switch (header) {
        case "ID Kamar": return payload.id || "";
        case "Nomor Kamar": return payload.nomor || "";
        case "Nama Mess": return payload.nama_mess || "";
        case "Kapasitas": return payload.kapasitas;
        case "Jumlah Penghuni": return payload.penghuni;
        case "Status": return payload.status || "";
        case "Kondisi Kamar": return payload.kondisi || "";
        case "Keterangan": return payload.keterangan || "";
        default: return "";
      }
    });

    if (rowIdx === -1) {
      payload.id = "RMR" + String(data.length).padStart(3, "0");
      rowValues[0] = payload.id;
      sheet.appendRow(rowValues);
    } else {
      sheet.getRange(rowIdx, 1, 1, rowValues.length).setValues([rowValues]);
    }

    syncKamarPenghuniCount(ss);
    return { success: true };
  } catch (err) {
    return { success: false, message: err.message };
  } finally {
    lock.releaseLock();
  }
}

function deleteKamar(token, kamarId) {
  const session = validateSession(token);
  if (!session.valid) return { success: false, message: session.message, sessionExpired: true };

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Kamar");
    const data = sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === kamarId) {
        sheet.deleteRow(i + 1);
        return { success: true };
      }
    }
    return { success: false, message: "ID Kamar tidak ditemukan." };
  } catch (err) {
    return { success: false, message: err.message };
  } finally {
    lock.releaseLock();
  }
}

// ================= CRUD: PENGHUNI OPERATIONS =================
function savePenghuni(token, payload) {
  const session = validateSession(token);
  if (!session.valid) return { success: false, message: session.message, sessionExpired: true };

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName("Penghuni");
    if (!sheet) {
      sheet = ss.insertSheet("Penghuni");
      sheet.appendRow([
        "ID Penghuni", "NIK", "Nama Karyawan", "Departemen", "Jabatan", "Perusahaan", "Status Karyawan", "Nomor Kamar", "Tanggal Check In", "Tanggal Check Out"
      ]);
    }
    const data = sheet.getDataRange().getValues();

    payload.nik = sanitizeInput(payload.nik, 50);
    payload.nama = sanitizeInput(payload.nama, 100);
    payload.departemen = sanitizeInput(payload.departemen, 100);
    payload.jabatan = sanitizeInput(payload.jabatan, 100);
    payload.perusahaan = sanitizeInput(payload.perusahaan, 100);
    payload.statusKaryawan = sanitizeInput(payload.statusKaryawan, 50);
    payload.nomorKamar = sanitizeInput(payload.nomorKamar, 50);

    let rowIdx = -1;
    if (payload.id) {
      for (let i = 1; i < data.length; i++) {
        if (data[i][0] === payload.id) {
          rowIdx = i + 1;
          break;
        }
      }
    }

    if (rowIdx === -1) {
      payload.id = "PNG" + String(data.length).padStart(3, "0");
      sheet.appendRow([
        payload.id, payload.nik, payload.nama, payload.departemen, payload.jabatan,
        payload.perusahaan, payload.statusKaryawan, payload.nomorKamar, payload.tglCheckIn, payload.tglCheckOut
      ]);
    } else {
      sheet.getRange(rowIdx, 2).setValue(payload.nik);
      sheet.getRange(rowIdx, 3).setValue(payload.nama);
      sheet.getRange(rowIdx, 4).setValue(payload.departemen);
      sheet.getRange(rowIdx, 5).setValue(payload.jabatan);
      sheet.getRange(rowIdx, 6).setValue(payload.perusahaan);
      sheet.getRange(rowIdx, 7).setValue(payload.statusKaryawan);
      sheet.getRange(rowIdx, 8).setValue(payload.nomorKamar);
      sheet.getRange(rowIdx, 9).setValue(payload.tglCheckIn);
      sheet.getRange(rowIdx, 10).setValue(payload.tglCheckOut);
    }

    syncKamarPenghuniCount(ss);
    return { success: true };
  } catch (err) {
    return { success: false, message: err.message };
  } finally {
    lock.releaseLock();
  }
}

function deletePenghuni(token, penghuniId) {
  const session = validateSession(token);
  if (!session.valid) return { success: false, message: session.message, sessionExpired: true };

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Penghuni");
    const data = sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === penghuniId) {
        sheet.deleteRow(i + 1);
        syncKamarPenghuniCount(ss);
        return { success: true };
      }
    }
    return { success: false, message: "ID Penghuni tidak ditemukan." };
  } catch (err) {
    return { success: false, message: err.message };
  } finally {
    lock.releaseLock();
  }
}

// ================= CRUD: INVENTARIS KAMAR OPERATIONS =================
function saveInventaris(token, payload) {
  const session = validateSession(token);
  if (!session.valid) return { success: false, message: session.message, sessionExpired: true };

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName("InventarisKamar");
    if (!sheet) {
      sheet = ss.insertSheet("InventarisKamar");
      sheet.appendRow([
        "ID Inventaris", "Kamar", "Kipas AC Tipe", "Kipas AC Qty", "Kipas AC Kondisi", "Kasur Qty", "Kasur Kondisi", "Lemari Qty", "Lemari Kondisi", "Rak Sepatu Qty", "Rak Sepatu Kondisi", "Keterangan"
      ]);
    }
    const data = sheet.getDataRange().getValues();

    payload.kamar = sanitizeInput(payload.kamar, 100);
    payload.kipasAcTipe = sanitizeInput(payload.kipasAcTipe, 50);
    payload.kipasAcKondisi = sanitizeInput(payload.kipasAcKondisi, 50);
    payload.kasurKondisi = sanitizeInput(payload.kasurKondisi, 50);
    payload.lemariKondisi = sanitizeInput(payload.lemariKondisi, 50);
    payload.rakSepatuKondisi = sanitizeInput(payload.rakSepatuKondisi, 50);
    payload.keterangan = sanitizeInput(payload.keterangan, 500);

    let rowIdx = -1;
    if (payload.id) {
      for (let i = 1; i < data.length; i++) {
        if (data[i][0] === payload.id) {
          rowIdx = i + 1;
          break;
        }
      }
    }

    const rowValues = [
      payload.id || "",
      payload.kamar || "",
      payload.kipasAcTipe || "Tidak Ada",
      parseInt(payload.kipasAcQty) || 0,
      payload.kipasAcKondisi || "Baik",
      parseInt(payload.kasurQty) || 0,
      payload.kasurKondisi || "Baik",
      parseInt(payload.lemariQty) || 0,
      payload.lemariKondisi || "Baik",
      parseInt(payload.rakSepatuQty) || 0,
      payload.rakSepatuKondisi || "Baik",
      payload.keterangan || ""
    ];

    if (rowIdx === -1) {
      payload.id = "INV" + String(data.length).padStart(3, "0");
      rowValues[0] = payload.id;
      sheet.appendRow(rowValues);
    } else {
      sheet.getRange(rowIdx, 1, 1, rowValues.length).setValues([rowValues]);
    }

    return { success: true };
  } catch (err) {
    return { success: false, message: err.message };
  } finally {
    lock.releaseLock();
  }
}

function deleteInventaris(token, inventarisId) {
  const session = validateSession(token);
  if (!session.valid) return { success: false, message: session.message, sessionExpired: true };

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("InventarisKamar");
    const data = sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === inventarisId) {
        sheet.deleteRow(i + 1);
        return { success: true };
      }
    }
    return { success: false, message: "ID Inventaris tidak ditemukan." };
  } catch (err) {
    return { success: false, message: err.message };
  } finally {
    lock.releaseLock();
  }
}

// ================= CRUD: ASET LAPTOP OPERATIONS =================
function saveAsetLaptop(token, payload) {
  const session = validateSession(token);
  if (!session.valid) return { success: false, message: session.message, sessionExpired: true };

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName("AsetLaptop");
    if (!sheet) {
      sheet = ss.insertSheet("AsetLaptop");
      sheet.appendRow([
        "ID Aset", "Asset Tag", "Merk", "Model", "Serial Number", "Processor", "RAM", "Storage", "Sistem Operasi", "Pengguna", "Departemen", "Site", "Lokasi", "Tanggal Pembelian", "Masa Garansi", "Kondisi", "Status Asset", "Catatan"
      ]);
    }
    const data = sheet.getDataRange().getValues();

    payload.tag = sanitizeInput(payload.tag, 100);
    payload.merk = sanitizeInput(payload.merk, 100);
    payload.model = sanitizeInput(payload.model, 100);
    payload.sn = sanitizeInput(payload.sn, 100);
    payload.processor = sanitizeInput(payload.processor, 100);
    payload.ram = sanitizeInput(payload.ram, 50);
    payload.storage = sanitizeInput(payload.storage, 50);
    payload.os = sanitizeInput(payload.os, 100);
    payload.pengguna = sanitizeInput(payload.pengguna, 100);
    payload.departemen = sanitizeInput(payload.departemen, 100);
    payload.site = sanitizeInput(payload.site, 100);
    payload.lokasi = sanitizeInput(payload.lokasi, 150);
    payload.garansi = sanitizeInput(payload.garansi, 100);
    payload.kondisi = sanitizeInput(payload.kondisi, 50);
    payload.status = sanitizeInput(payload.status, 50);
    payload.catatan = sanitizeInput(payload.catatan, 500);

    let rowIdx = -1;
    if (payload.id) {
      for (let i = 1; i < data.length; i++) {
        if (data[i][0] === payload.id) {
          rowIdx = i + 1;
          break;
        }
      }
    }

    if (rowIdx === -1) {
      payload.id = "AST" + String(data.length).padStart(3, "0");
      sheet.appendRow([
        payload.id, payload.tag, payload.merk, payload.model, payload.sn,
        payload.processor, payload.ram, payload.storage, payload.os,
        payload.pengguna, payload.departemen, payload.site, payload.lokasi,
        payload.tglBeli, payload.garansi, payload.kondisi, payload.status, payload.catatan
      ]);
    } else {
      sheet.getRange(rowIdx, 2).setValue(payload.tag);
      sheet.getRange(rowIdx, 3).setValue(payload.merk);
      sheet.getRange(rowIdx, 4).setValue(payload.model);
      sheet.getRange(rowIdx, 5).setValue(payload.sn);
      sheet.getRange(rowIdx, 6).setValue(payload.processor);
      sheet.getRange(rowIdx, 7).setValue(payload.ram);
      sheet.getRange(rowIdx, 8).setValue(payload.storage);
      sheet.getRange(rowIdx, 9).setValue(payload.os);
      sheet.getRange(rowIdx, 10).setValue(payload.pengguna);
      sheet.getRange(rowIdx, 11).setValue(payload.departemen);
      sheet.getRange(rowIdx, 12).setValue(payload.site);
      sheet.getRange(rowIdx, 13).setValue(payload.lokasi);
      sheet.getRange(rowIdx, 14).setValue(payload.tglBeli);
      sheet.getRange(rowIdx, 15).setValue(payload.garansi);
      sheet.getRange(rowIdx, 16).setValue(payload.kondisi);
      sheet.getRange(rowIdx, 17).setValue(payload.status);
      sheet.getRange(rowIdx, 18).setValue(payload.catatan);
    }

    return { success: true };
  } catch (err) {
    return { success: false, message: err.message };
  } finally {
    lock.releaseLock();
  }
}

function deleteAsetLaptop(token, assetId) {
  const session = validateSession(token);
  if (!session.valid) return { success: false, message: session.message, sessionExpired: true };

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("AsetLaptop");
    const data = sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === assetId) {
        sheet.deleteRow(i + 1);
        return { success: true };
      }
    }
    return { success: false, message: "ID Aset tidak ditemukan." };
  } catch (err) {
    return { success: false, message: err.message };
  } finally {
    lock.releaseLock();
  }
}

// ================= CRUD: ASET KOMPUTER OPERATIONS =================
function saveAsetKomputer(token, payload) {
  const session = validateSession(token);
  if (!session.valid) return { success: false, message: session.message, sessionExpired: true };

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName("AsetKomputer");
    if (!sheet) {
      sheet = ss.insertSheet("AsetKomputer");
      sheet.appendRow([
        "ID Aset", "Asset Tag", "Merk", "Model", "Serial Number", "Processor", "RAM", "Storage", "Sistem Operasi", "Pengguna", "Departemen", "Site", "Lokasi", "Tanggal Pembelian", "Masa Garansi", "Kondisi", "Status Asset", "Catatan"
      ]);
    }
    const data = sheet.getDataRange().getValues();

    payload.tag = sanitizeInput(payload.tag, 100);
    payload.merk = sanitizeInput(payload.merk, 100);
    payload.model = sanitizeInput(payload.model, 100);
    payload.sn = sanitizeInput(payload.sn, 100);
    payload.processor = sanitizeInput(payload.processor, 100);
    payload.ram = sanitizeInput(payload.ram, 50);
    payload.storage = sanitizeInput(payload.storage, 50);
    payload.os = sanitizeInput(payload.os, 100);
    payload.pengguna = sanitizeInput(payload.pengguna, 100);
    payload.departemen = sanitizeInput(payload.departemen, 100);
    payload.site = sanitizeInput(payload.site, 100);
    payload.lokasi = sanitizeInput(payload.lokasi, 150);
    payload.garansi = sanitizeInput(payload.garansi, 100);
    payload.kondisi = sanitizeInput(payload.kondisi, 50);
    payload.status = sanitizeInput(payload.status, 50);
    payload.catatan = sanitizeInput(payload.catatan, 500);

    let rowIdx = -1;
    if (payload.id) {
      for (let i = 1; i < data.length; i++) {
        if (data[i][0] === payload.id) {
          rowIdx = i + 1;
          break;
        }
      }
    }

    if (rowIdx === -1) {
      payload.id = "PC" + String(data.length).padStart(3, "0");
      sheet.appendRow([
        payload.id, payload.tag, payload.merk, payload.model, payload.sn,
        payload.processor, payload.ram, payload.storage, payload.os,
        payload.pengguna, payload.departemen, payload.site, payload.lokasi,
        payload.tglBeli, payload.garansi, payload.kondisi, payload.status, payload.catatan
      ]);
    } else {
      sheet.getRange(rowIdx, 2).setValue(payload.tag);
      sheet.getRange(rowIdx, 3).setValue(payload.merk);
      sheet.getRange(rowIdx, 4).setValue(payload.model);
      sheet.getRange(rowIdx, 5).setValue(payload.sn);
      sheet.getRange(rowIdx, 6).setValue(payload.processor);
      sheet.getRange(rowIdx, 7).setValue(payload.ram);
      sheet.getRange(rowIdx, 8).setValue(payload.storage);
      sheet.getRange(rowIdx, 9).setValue(payload.os);
      sheet.getRange(rowIdx, 10).setValue(payload.pengguna);
      sheet.getRange(rowIdx, 11).setValue(payload.departemen);
      sheet.getRange(rowIdx, 12).setValue(payload.site);
      sheet.getRange(rowIdx, 13).setValue(payload.lokasi);
      sheet.getRange(rowIdx, 14).setValue(payload.tglBeli);
      sheet.getRange(rowIdx, 15).setValue(payload.garansi);
      sheet.getRange(rowIdx, 16).setValue(payload.kondisi);
      sheet.getRange(rowIdx, 17).setValue(payload.status);
      sheet.getRange(rowIdx, 18).setValue(payload.catatan);
    }

    return { success: true };
  } catch (err) {
    return { success: false, message: err.message };
  } finally {
    lock.releaseLock();
  }
}

function deleteAsetKomputer(token, assetId) {
  const session = validateSession(token);
  if (!session.valid) return { success: false, message: session.message, sessionExpired: true };

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("AsetKomputer");
    const data = sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === assetId) {
        sheet.deleteRow(i + 1);
        return { success: true };
      }
    }
    return { success: false, message: "ID Aset tidak ditemukan." };
  } catch (err) {
    return { success: false, message: err.message };
  } finally {
    lock.releaseLock();
  }
}

// ================= CRUD: ASET RADIO HT OPERATIONS =================
function saveAsetHT(token, payload) {
  const session = validateSession(token);
  if (!session.valid) return { success: false, message: session.message, sessionExpired: true };

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName("AsetHT");
    if (!sheet) {
      sheet = ss.insertSheet("AsetHT");
      sheet.appendRow([
        "ID Aset", "Asset Tag", "Merk", "Model", "Serial Number", "Band", "Mode", "Pengguna", "Departemen", "Site", "Lokasi", "Tanggal Pembelian", "Masa Garansi", "Kondisi", "Status Asset", "Catatan"
      ]);
    }
    const data = sheet.getDataRange().getValues();

    payload.tag = sanitizeInput(payload.tag, 100);
    payload.merk = sanitizeInput(payload.merk, 100);
    payload.model = sanitizeInput(payload.model, 100);
    payload.sn = sanitizeInput(payload.sn, 100);
    payload.band = sanitizeInput(payload.band, 50);
    payload.mode = sanitizeInput(payload.mode, 50);
    payload.pengguna = sanitizeInput(payload.pengguna, 100);
    payload.departemen = sanitizeInput(payload.departemen, 100);
    payload.site = sanitizeInput(payload.site, 100);
    payload.lokasi = sanitizeInput(payload.lokasi, 150);
    payload.garansi = sanitizeInput(payload.garansi, 100);
    payload.kondisi = sanitizeInput(payload.kondisi, 50);
    payload.status = sanitizeInput(payload.status, 50);
    payload.catatan = sanitizeInput(payload.catatan, 500);

    let rowIdx = -1;
    if (payload.id) {
      for (let i = 1; i < data.length; i++) {
        if (data[i][0] === payload.id) {
          rowIdx = i + 1;
          break;
        }
      }
    }

    if (rowIdx === -1) {
      payload.id = "HT" + String(data.length).padStart(3, "0");
      sheet.appendRow([
        payload.id, payload.tag, payload.merk, payload.model, payload.sn,
        payload.band, payload.mode, payload.pengguna, payload.departemen, payload.site,
        payload.lokasi, payload.tglBeli, payload.garansi, payload.kondisi, payload.status, payload.catatan
      ]);
    } else {
      sheet.getRange(rowIdx, 2).setValue(payload.tag);
      sheet.getRange(rowIdx, 3).setValue(payload.merk);
      sheet.getRange(rowIdx, 4).setValue(payload.model);
      sheet.getRange(rowIdx, 5).setValue(payload.sn);
      sheet.getRange(rowIdx, 6).setValue(payload.band);
      sheet.getRange(rowIdx, 7).setValue(payload.mode);
      sheet.getRange(rowIdx, 8).setValue(payload.pengguna);
      sheet.getRange(rowIdx, 9).setValue(payload.departemen);
      sheet.getRange(rowIdx, 10).setValue(payload.site);
      sheet.getRange(rowIdx, 11).setValue(payload.lokasi);
      sheet.getRange(rowIdx, 12).setValue(payload.tglBeli);
      sheet.getRange(rowIdx, 13).setValue(payload.garansi);
      sheet.getRange(rowIdx, 14).setValue(payload.kondisi);
      sheet.getRange(rowIdx, 15).setValue(payload.status);
      sheet.getRange(rowIdx, 16).setValue(payload.catatan);
    }

    return { success: true };
  } catch (err) {
    return { success: false, message: err.message };
  } finally {
    lock.releaseLock();
  }
}

function deleteAsetHT(token, assetId) {
  const session = validateSession(token);
  if (!session.valid) return { success: false, message: session.message, sessionExpired: true };

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("AsetHT");
    const data = sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === assetId) {
        sheet.deleteRow(i + 1);
        return { success: true };
      }
    }
    return { success: false, message: "ID Aset tidak ditemukan." };
  } catch (err) {
    return { success: false, message: err.message };
  } finally {
    lock.releaseLock();
  }
}

// ================= CRUD: ASET RADIO RIG OPERATIONS =================
function saveAsetRIG(token, payload) {
  const session = validateSession(token);
  if (!session.valid) return { success: false, message: session.message, sessionExpired: true };

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName("AsetRIG");
    if (!sheet) {
      sheet = ss.insertSheet("AsetRIG");
      sheet.appendRow([
        "ID Aset", "Asset Tag", "Merk", "Model", "Serial Number", "Band", "Mode", "Unit Kendaraan", "Departemen", "Site", "Lokasi", "Tanggal Pembelian", "Kondisi", "Status Asset", "Catatan"
      ]);
    }
    const data = sheet.getDataRange().getValues();

    payload.tag = sanitizeInput(payload.tag, 100);
    payload.merk = sanitizeInput(payload.merk, 100);
    payload.model = sanitizeInput(payload.model, 100);
    payload.sn = sanitizeInput(payload.sn, 100);
    payload.band = sanitizeInput(payload.band, 50);
    payload.mode = sanitizeInput(payload.mode, 50);
    payload.unit = sanitizeInput(payload.unit, 100);
    payload.departemen = sanitizeInput(payload.departemen, 100);
    payload.site = sanitizeInput(payload.site, 100);
    payload.lokasi = sanitizeInput(payload.lokasi, 150);
    payload.kondisi = sanitizeInput(payload.kondisi, 50);
    payload.status = sanitizeInput(payload.status, 50);
    payload.catatan = sanitizeInput(payload.catatan, 500);

    let rowIdx = -1;
    if (payload.id) {
      for (let i = 1; i < data.length; i++) {
        if (data[i][0] === payload.id) {
          rowIdx = i + 1;
          break;
        }
      }
    }

    if (rowIdx === -1) {
      payload.id = "RG" + String(data.length).padStart(3, "0");
      sheet.appendRow([
        payload.id, payload.tag, payload.merk, payload.model, payload.sn,
        payload.band, payload.mode, payload.unit, payload.departemen, payload.site,
        payload.lokasi, payload.tglBeli, payload.kondisi, payload.status, payload.catatan
      ]);
    } else {
      sheet.getRange(rowIdx, 2).setValue(payload.tag);
      sheet.getRange(rowIdx, 3).setValue(payload.merk);
      sheet.getRange(rowIdx, 4).setValue(payload.model);
      sheet.getRange(rowIdx, 5).setValue(payload.sn);
      sheet.getRange(rowIdx, 6).setValue(payload.band);
      sheet.getRange(rowIdx, 7).setValue(payload.mode);
      sheet.getRange(rowIdx, 8).setValue(payload.unit);
      sheet.getRange(rowIdx, 9).setValue(payload.departemen);
      sheet.getRange(rowIdx, 10).setValue(payload.site);
      sheet.getRange(rowIdx, 11).setValue(payload.lokasi);
      sheet.getRange(rowIdx, 12).setValue(payload.tglBeli);
      sheet.getRange(rowIdx, 13).setValue(payload.kondisi);
      sheet.getRange(rowIdx, 14).setValue(payload.status);
      sheet.getRange(rowIdx, 15).setValue(payload.catatan);
    }

    return { success: true };
  } catch (err) {
    return { success: false, message: err.message };
  } finally {
    lock.releaseLock();
  }
}

function deleteAsetRIG(token, assetId) {
  const session = validateSession(token);
  if (!session.valid) return { success: false, message: session.message, sessionExpired: true };

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("AsetRIG");
    const data = sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === assetId) {
        sheet.deleteRow(i + 1);
        return { success: true };
      }
    }
    return { success: false, message: "ID Aset tidak ditemukan." };
  } catch (err) {
    return { success: false, message: err.message };
  } finally {
    lock.releaseLock();
  }
}

// ================= CRUD: ASET PERANGKAT PENDUKUNG OPERATIONS =================
function saveAsetPerangkat(token, payload) {
  const session = validateSession(token);
  if (!session.valid) return { success: false, message: session.message, sessionExpired: true };

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName("AsetPerangkat");
    if (!sheet) {
      sheet = ss.insertSheet("AsetPerangkat");
      sheet.appendRow([
        "ID Aset", "Asset Tag", "Jenis Perangkat", "Merk", "Model", "Serial Number", "Departemen", "Site", "Lokasi", "Tanggal Pembelian", "Kondisi", "Status Asset", "Catatan"
      ]);
    }
    const data = sheet.getDataRange().getValues();

    payload.tag = sanitizeInput(payload.tag, 100);
    payload.jenis = sanitizeInput(payload.jenis, 100);
    payload.merk = sanitizeInput(payload.merk, 100);
    payload.model = sanitizeInput(payload.model, 100);
    payload.sn = sanitizeInput(payload.sn, 100);
    payload.departemen = sanitizeInput(payload.departemen, 100);
    payload.site = sanitizeInput(payload.site, 100);
    payload.lokasi = sanitizeInput(payload.lokasi, 150);
    payload.kondisi = sanitizeInput(payload.kondisi, 50);
    payload.status = sanitizeInput(payload.status, 50);
    payload.catatan = sanitizeInput(payload.catatan, 500);

    let rowIdx = -1;
    if (payload.id) {
      for (let i = 1; i < data.length; i++) {
        if (data[i][0] === payload.id) {
          rowIdx = i + 1;
          break;
        }
      }
    }

    if (rowIdx === -1) {
      payload.id = "PR" + String(data.length).padStart(3, "0");
      sheet.appendRow([
        payload.id, payload.tag, payload.jenis, payload.merk, payload.model, payload.sn,
        payload.departemen, payload.site, payload.lokasi, payload.tglBeli, payload.kondisi, payload.status, payload.catatan
      ]);
    } else {
      sheet.getRange(rowIdx, 2).setValue(payload.tag);
      sheet.getRange(rowIdx, 3).setValue(payload.jenis);
      sheet.getRange(rowIdx, 4).setValue(payload.merk);
      sheet.getRange(rowIdx, 5).setValue(payload.model);
      sheet.getRange(rowIdx, 6).setValue(payload.sn);
      sheet.getRange(rowIdx, 7).setValue(payload.departemen);
      sheet.getRange(rowIdx, 8).setValue(payload.site);
      sheet.getRange(rowIdx, 9).setValue(payload.lokasi);
      sheet.getRange(rowIdx, 10).setValue(payload.tglBeli);
      sheet.getRange(rowIdx, 11).setValue(payload.kondisi);
      sheet.getRange(rowIdx, 12).setValue(payload.status);
      sheet.getRange(rowIdx, 13).setValue(payload.catatan);
    }

    return { success: true };
  } catch (err) {
    return { success: false, message: err.message };
  } finally {
    lock.releaseLock();
  }
}

function deleteAsetPerangkat(token, assetId) {
  const session = validateSession(token);
  if (!session.valid) return { success: false, message: session.message, sessionExpired: true };

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("AsetPerangkat");
    const data = sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === assetId) {
        sheet.deleteRow(i + 1);
        return { success: true };
      }
    }
    return { success: false, message: "ID Aset tidak ditemukan." };
  } catch (err) {
    return { success: false, message: err.message };
  } finally {
    lock.releaseLock();
  }
}

// ================= CRUD: DATA UNIT DT (PRODUKSI) =================
function saveUnitDT(token, payload) {
  const session = validateSession(token);
  if (!session.valid) return { success: false, message: session.message, sessionExpired: true };

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName("DataUnitDT");
    const standardHeaders = [
      "ID Unit", "Nomor Unit", "Jenis Unit", "Merk", "Model", "Tahun Pembelian", "Kapasitas", "Jumlah Roda", "Status"
    ];

    if (!sheet) {
      sheet = ss.insertSheet("DataUnitDT");
      sheet.appendRow(standardHeaders);
    } else {
      // Normalize header row 1 to prevent column shifting
      sheet.getRange(1, 1, 1, standardHeaders.length).setValues([standardHeaders]);
    }

    const data = sheet.getDataRange().getValues();
    const uMap = getColumnMap(sheet);

    payload.nomor = sanitizeInput(payload.nomor, 50);
    payload.jenis = sanitizeInput(payload.jenis, 50);
    payload.merk = sanitizeInput(payload.merk, 100);
    payload.model = sanitizeInput(payload.model, 100);
    payload.tahun = sanitizeInput(payload.tahun, 10);
    payload.kapasitas = sanitizeInput(payload.kapasitas, 50);
    payload.jumlahRoda = sanitizeInput(payload.jumlahRoda, 10);
    payload.status = sanitizeInput(payload.status, 20);

    let rowIdx = -1;
    if (payload.id) {
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][0]) === String(payload.id)) {
          rowIdx = i + 1;
          break;
        }
      }
    }

    const setValByHeader = (rIdx, headerName, val, defaultCol) => {
      const col = uMap[headerName] || defaultCol;
      if (col && col > 0) {
        sheet.getRange(rIdx, col).setValue(val);
      }
    };

    if (rowIdx === -1) {
      let maxId = 0;
      for (let i = 1; i < data.length; i++) {
        let val = String(data[i][0]);
        let numStr = val.replace(/\D/g, "");
        let currentId = parseInt(numStr);
        if (!isNaN(currentId) && currentId > maxId) {
          maxId = currentId;
        }
      }
      payload.id = String(maxId + 1);
      rowIdx = data.length + 1;

      setValByHeader(rowIdx, "ID Unit", payload.id, 1);
      setValByHeader(rowIdx, "Nomor Unit", payload.nomor, 2);
      setValByHeader(rowIdx, "Jenis Unit", payload.jenis, 3);
      setValByHeader(rowIdx, "Merk", payload.merk, 4);
      setValByHeader(rowIdx, "Model", payload.model, 5);
      setValByHeader(rowIdx, "Tahun Pembelian", payload.tahun, 6);
      setValByHeader(rowIdx, "Kapasitas", payload.kapasitas, 7);
      setValByHeader(rowIdx, "Jumlah Roda", payload.jumlahRoda, 8);
      setValByHeader(rowIdx, "Status", payload.status, 9);
    } else {
      setValByHeader(rowIdx, "Nomor Unit", payload.nomor, 2);
      setValByHeader(rowIdx, "Jenis Unit", payload.jenis, 3);
      setValByHeader(rowIdx, "Merk", payload.merk, 4);
      setValByHeader(rowIdx, "Model", payload.model, 5);
      setValByHeader(rowIdx, "Tahun Pembelian", payload.tahun, 6);
      setValByHeader(rowIdx, "Kapasitas", payload.kapasitas, 7);
      setValByHeader(rowIdx, "Jumlah Roda", payload.jumlahRoda, 8);
      setValByHeader(rowIdx, "Status", payload.status, 9);
    }

    return { success: true };
  } catch (err) {
    return { success: false, message: err.message };
  } finally {
    lock.releaseLock();
  }
}

function deleteUnitDT(token, unitId) {
  const session = validateSession(token);
  if (!session.valid) return { success: false, message: session.message, sessionExpired: true };

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("DataUnitDT");
    const data = sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(unitId)) {
        sheet.deleteRow(i + 1);
        return { success: true };
      }
    }
    return { success: false, message: "ID Unit DT tidak ditemukan." };
  } catch (err) {
    return { success: false, message: err.message };
  } finally {
    lock.releaseLock();
  }
}

// ================= CRUD: PRODUKSI HAULING (PRODUKSI) =================
function saveProduksiHauling(token, payload) {
  const session = validateSession(token);
  if (!session.valid) return { success: false, message: session.message, sessionExpired: true };

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName("ProduksiHauling");
    if (!sheet) {
      sheet = ss.insertSheet("ProduksiHauling");
      sheet.appendRow([
        "ID Hauling", "Tanggal Input Data", "Tanggal Start", "Jam Start", "Tanggal Selesai", "Jam Selesai",
        "Shift", "Nomor Unit", "Jenis DT", "Jumlah Roda", "Driver", "Loading Point", "Dumping Point", "Lokasi PIT", "Lokasi Site",
        "Jumlah Ritase", "Volume (Ton)", "Timbangan Kosongan (TARA)", "Timbangan", "Netto (Kg)", "Netto (Ton)",
        "HM Awal", "HM Akhir", "Pengawas", "Keterangan"
      ]);
    }
    const data = sheet.getDataRange().getValues();

    // Sanitization & Parsing
    payload.tanggalInput = sanitizeInput(payload.tanggalInput, 20);
    payload.tanggalStart = sanitizeInput(payload.tanggalStart, 20);
    payload.jamStart = sanitizeInput(payload.jamStart, 20);
    payload.tanggalSelesai = sanitizeInput(payload.tanggalSelesai, 20);
    payload.jamSelesai = sanitizeInput(payload.jamSelesai, 20);
    payload.shift = sanitizeInput(payload.shift, 20);
    payload.unit = sanitizeInput(payload.unit, 50);
    payload.jenisDt = sanitizeInput(payload.jenisDt, 50);
    payload.merkUnit = sanitizeInput(payload.merkUnit, 50);
    payload.jumlahRoda = sanitizeInput(payload.jumlahRoda, 20);
    payload.driver = sanitizeInput(payload.driver, 100);
    payload.loadingPoint = sanitizeInput(payload.loadingPoint, 100);
    payload.dumpingPoint = sanitizeInput(payload.dumpingPoint, 100);
    payload.lokasiPit = sanitizeInput(payload.lokasiPit, 100);
    payload.lokasiSite = sanitizeInput(payload.lokasiSite, 100);
    payload.ritase = parseFloat(payload.ritase || 0);
    payload.volume = parseFloat(payload.volume || 0);
    payload.tara = parseFloat(payload.tara || 0);
    payload.timbangan = sanitizeInput(payload.timbangan, 50);
    payload.nettoKg = parseFloat(payload.nettoKg || 0);
    payload.nettoTon = parseFloat(payload.nettoTon || 0);
    payload.hmAwal = sanitizeInput(payload.hmAwal, 20);
    payload.hmAkhir = sanitizeInput(payload.hmAkhir, 20);
    payload.pengawas = sanitizeInput(payload.pengawas, 100);
    payload.keterangan = sanitizeInput(payload.keterangan, 500);

    let rowIdx = -1;
    if (payload.id) {
      for (let i = 1; i < data.length; i++) {
        if (data[i][0] === payload.id) {
          rowIdx = i + 1;
          break;
        }
      }
    }

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => String(h).trim());
    const rowValues = headers.map(header => {
      switch (header) {
        case "ID Hauling": return payload.id || "";
        case "Tanggal Input Data": return payload.tanggalInput || "";
        case "Tanggal Start": return payload.tanggalStart || "";
        case "Jam Start": return payload.jamStart || "";
        case "Tanggal Selesai": return payload.tanggalSelesai || "";
        case "Jam Selesai": return payload.jamSelesai || "";
        case "Shift": return payload.shift || "";
        case "Nomor Unit": return payload.unit || "";
        case "Jenis DT": return payload.jenisDt || "";
        case "Merk Unit": return payload.merkUnit || "";
        case "Jumlah Roda": return payload.jumlahRoda || "";
        case "Driver": return payload.driver || "";
        case "Loading Point": return payload.loadingPoint || "";
        case "Dumping Point": return payload.dumpingPoint || "";
        case "Lokasi PIT": return payload.lokasiPit || "";
        case "Lokasi Site": return payload.lokasiSite || "";
        case "Jumlah Ritase": return payload.ritase;
        case "Volume (Ton)": return payload.volume;
        case "Timbangan Kosongan (TARA)": return payload.tara;
        case "Timbangan": return payload.timbangan;
        case "Netto (Kg)": return payload.nettoKg;
        case "Netto (Ton)": return payload.nettoTon;
        case "HM Awal": return payload.hmAwal || "";
        case "HM Akhir": return payload.hmAkhir || "";
        case "Pengawas": return payload.pengawas || "";
        case "Keterangan": return payload.keterangan || "";
        default: return "";
      }
    });

    if (rowIdx === -1) {
      payload.id = "HLG" + String(data.length).padStart(3, "0");
      const idColIdx = headers.indexOf("ID Hauling");
      if (idColIdx !== -1) {
        rowValues[idColIdx] = payload.id;
      }
      sheet.appendRow(rowValues);
    } else {
      sheet.getRange(rowIdx, 1, 1, rowValues.length).setValues([rowValues]);
    }

    return { success: true, insertedId: payload.id };
  } catch (err) {
    return { success: false, message: err.message };
  } finally {
    lock.releaseLock();
  }
}

function saveBatchProduksiHauling(token, payloadArray) {
  const session = validateSession(token);
  if (!session.valid) return { success: false, message: session.message, sessionExpired: true };
  if (!Array.isArray(payloadArray) || payloadArray.length === 0) {
    return { success: false, message: "Tidak ada data batch import yang dikirim." };
  }

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName("ProduksiHauling");
    if (!sheet) {
      sheet = ss.insertSheet("ProduksiHauling");
      sheet.appendRow([
        "ID Hauling", "Tanggal Input Data", "Tanggal Start", "Jam Start", "Tanggal Selesai", "Jam Selesai",
        "Shift", "Nomor Unit", "Jenis DT", "Merk Unit", "Jumlah Roda", "Driver", "Loading Point", "Dumping Point", "Lokasi PIT", "Lokasi Site",
        "Jumlah Ritase", "Volume (Ton)", "Timbangan Kosongan (TARA)", "Timbangan", "Netto (Kg)", "Netto (Ton)",
        "HM Awal", "HM Akhir", "Pengawas", "Keterangan"
      ]);
    }
    const data = sheet.getDataRange().getValues();
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => String(h).trim());

    const idColIdx = headers.indexOf("ID Hauling");
    const tglStartIdx = headers.indexOf("Tanggal Start");
    const jamStartIdx = headers.indexOf("Jam Start");
    const unitIdx = headers.indexOf("Nomor Unit");
    const shiftIdx = headers.indexOf("Shift");
    const driverIdx = headers.indexOf("Driver");
    const pengawasIdx = headers.indexOf("Pengawas");

    // Build lookup map for existing rows
    const existingMap = {};
    const volColIdx = headers.indexOf("Volume (Ton)");

    for (let r = 1; r < data.length; r++) {
      const u = String(data[r][unitIdx] || "").trim().toLowerCase();
      const ts = safeDateFormat(data[r][tglStartIdx]).trim();
      const js = safeTimeFormat(data[r][jamStartIdx]).trim();
      const sh = String(data[r][shiftIdx] || "").trim().toLowerCase();
      const volStr = String(Math.round(parseFloat(data[r][volColIdx] || 0)));

      if (u && ts) {
        const itemInfo = {
          rowIdx: r + 1, // 1-based row index in sheet
          id: String(data[r][idColIdx] || ""),
          existingDriver: String(data[r][driverIdx] || ""),
          existingPengawas: String(data[r][pengawasIdx] || "")
        };
        const key1 = (u + "_" + ts + "_" + js + "_" + sh).toLowerCase();
        const key2 = (u + "_" + ts + "_" + sh + "_" + volStr).toLowerCase();
        existingMap[key1] = itemInfo;
        existingMap[key2] = itemInfo;
      }
    }

    let lastIdNum = data.length;
    const rowsToAppend = [];
    const savedItems = [];

    payloadArray.forEach((item, index) => {
      item.tanggalInput = sanitizeInput(item.tanggalInput, 20);
      item.tanggalStart = sanitizeInput(item.tanggalStart, 20);
      item.jamStart = sanitizeInput(item.jamStart, 20);
      item.tanggalSelesai = sanitizeInput(item.tanggalSelesai, 20);
      item.jamSelesai = sanitizeInput(item.jamSelesai, 20);
      item.shift = sanitizeInput(item.shift, 20);
      item.unit = sanitizeInput(item.unit, 50);
      item.jenisDt = sanitizeInput(item.jenisDt, 50);
      item.merkUnit = sanitizeInput(item.merkUnit, 50);
      item.jumlahRoda = sanitizeInput(item.jumlahRoda, 20);
      item.loadingPoint = sanitizeInput(item.loadingPoint || "", 100);
      item.dumpingPoint = sanitizeInput(item.dumpingPoint || "", 100);
      item.lokasiPit = sanitizeInput(item.lokasiPit || "", 100);
      item.lokasiSite = sanitizeInput(item.lokasiSite || "", 100);
      item.ritase = parseFloat(item.ritase || 1);
      item.volume = parseFloat(item.volume || 0);
      item.tara = parseFloat(item.tara || 0);
      item.timbangan = sanitizeInput(item.timbangan || "", 50);
      item.nettoKg = parseFloat(item.nettoKg || 0);
      item.nettoTon = parseFloat(item.nettoTon || 0);
      item.hmAwal = sanitizeInput(item.hmAwal || "", 20);
      item.hmAkhir = sanitizeInput(item.hmAkhir || "", 20);
      item.keterangan = sanitizeInput(item.keterangan || "Import Excel", 500);

      const itemTs = safeDateFormat(item.tanggalStart).trim();
      const itemJs = safeTimeFormat(item.jamStart).trim();
      const itemVolStr = String(Math.round(parseFloat(item.volume || 0)));

      const key1 = (item.unit + "_" + itemTs + "_" + itemJs + "_" + item.shift).trim().toLowerCase();
      const key2 = (item.unit + "_" + itemTs + "_" + item.shift + "_" + itemVolStr).trim().toLowerCase();
      const existing = existingMap[key1] || existingMap[key2];

      if (existing) {
        // Overwrite existing record! Preserve existing driver & pengawas if manual input exists
        item.id = existing.id;
        item.driver = sanitizeInput(item.driver || existing.existingDriver || "", 100);
        item.pengawas = sanitizeInput(item.pengawas || existing.existingPengawas || "", 100);

        const rowValues = headers.map(header => {
          switch (header) {
            case "ID Hauling": return item.id;
            case "Tanggal Input Data": return item.tanggalInput || "";
            case "Tanggal Start": return item.tanggalStart || "";
            case "Jam Start": return item.jamStart || "";
            case "Tanggal Selesai": return item.tanggalSelesai || "";
            case "Jam Selesai": return item.jamSelesai || "";
            case "Shift": return item.shift || "";
            case "Nomor Unit": return item.unit || "";
            case "Jenis DT": return item.jenisDt || "";
            case "Merk Unit": return item.merkUnit || "";
            case "Jumlah Roda": return item.jumlahRoda || "";
            case "Driver": return item.driver || "";
            case "Loading Point": return item.loadingPoint || "";
            case "Dumping Point": return item.dumpingPoint || "";
            case "Lokasi PIT": return item.lokasiPit || "";
            case "Lokasi Site": return item.lokasiSite || "";
            case "Jumlah Ritase": return item.ritase;
            case "Volume (Ton)": return item.volume;
            case "Timbangan Kosongan (TARA)": return item.tara;
            case "Timbangan": return item.timbangan;
            case "Netto (Kg)": return item.nettoKg;
            case "Netto (Ton)": return item.nettoTon;
            case "HM Awal": return item.hmAwal || "";
            case "HM Akhir": return item.hmAkhir || "";
            case "Pengawas": return item.pengawas || "";
            case "Keterangan": return item.keterangan || "";
            default: return "";
          }
        });

        sheet.getRange(existing.rowIdx, 1, 1, rowValues.length).setValues([rowValues]);
        savedItems.push(item);
      } else {
        item.driver = sanitizeInput(item.driver || "", 100);
        item.pengawas = sanitizeInput(item.pengawas || "", 100);
        item.id = "HLG" + String(lastIdNum + index).padStart(3, "0");

        const rowValues = headers.map(header => {
          switch (header) {
            case "ID Hauling": return item.id;
            case "Tanggal Input Data": return item.tanggalInput || "";
            case "Tanggal Start": return item.tanggalStart || "";
            case "Jam Start": return item.jamStart || "";
            case "Tanggal Selesai": return item.tanggalSelesai || "";
            case "Jam Selesai": return item.jamSelesai || "";
            case "Shift": return item.shift || "";
            case "Nomor Unit": return item.unit || "";
            case "Jenis DT": return item.jenisDt || "";
            case "Merk Unit": return item.merkUnit || "";
            case "Jumlah Roda": return item.jumlahRoda || "";
            case "Driver": return item.driver || "";
            case "Loading Point": return item.loadingPoint || "";
            case "Dumping Point": return item.dumpingPoint || "";
            case "Lokasi PIT": return item.lokasiPit || "";
            case "Lokasi Site": return item.lokasiSite || "";
            case "Jumlah Ritase": return item.ritase;
            case "Volume (Ton)": return item.volume;
            case "Timbangan Kosongan (TARA)": return item.tara;
            case "Timbangan": return item.timbangan;
            case "Netto (Kg)": return item.nettoKg;
            case "Netto (Ton)": return item.nettoTon;
            case "HM Awal": return item.hmAwal || "";
            case "HM Akhir": return item.hmAkhir || "";
            case "Pengawas": return item.pengawas || "";
            case "Keterangan": return item.keterangan || "";
            default: return "";
          }
        });

        rowsToAppend.push(rowValues);
        savedItems.push(item);
      }
    });

    if (rowsToAppend.length > 0) {
      const nextRow = sheet.getLastRow() + 1;
      sheet.getRange(nextRow, 1, rowsToAppend.length, headers.length).setValues(rowsToAppend);
    }

    return { success: true, count: savedItems.length, items: savedItems };
  } catch (err) {
    return { success: false, message: err.message };
  } finally {
    lock.releaseLock();
  }
}

function deleteProduksiHauling(token, haulingId) {
  const session = validateSession(token);
  if (!session.valid) return { success: false, message: session.message, sessionExpired: true };

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("ProduksiHauling");
    const data = sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === haulingId) {
        sheet.deleteRow(i + 1);
        return { success: true };
      }
    }
    return { success: false, message: "ID Hauling tidak ditemukan." };
  } catch (err) {
    return { success: false, message: err.message };
  } finally {
    lock.releaseLock();
  }
}



// ================= CRUD: MASTER DATA PRODUKSI OPERATIONS =================
function saveMasterProduksiItem(token, category, indexVal, value) {
  const session = validateSession(token);
  if (!session.valid) return { success: false, message: session.message, sessionExpired: true };

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName("MasterDataProduksi");
    if (!sheet) {
      sheet = ss.insertSheet("MasterDataProduksi");
      sheet.appendRow(["Category", "Value"]);
    }
    const data = sheet.getDataRange().getValues();

    value = sanitizeInput(value, 100);

    if (indexVal === "") {
      // Add new item
      sheet.appendRow([category, value]);
    } else {
      // Edit item
      let count = 0;
      let targetRow = -1;
      let targetIdx = parseInt(indexVal);

      for (let i = 1; i < data.length; i++) {
        if (data[i][0] === category) {
          if (count === targetIdx) {
            targetRow = i + 1;
            break;
          }
          count++;
        }
      }

      if (targetRow !== -1) {
        sheet.getRange(targetRow, 2).setValue(value);
      } else {
        return { success: false, message: "Urutan entitas master produksi tidak ditemukan." };
      }
    }
    return { success: true };
  } catch (err) {
    return { success: false, message: err.message };
  } finally {
    lock.releaseLock();
  }
}

function deleteMasterProduksiItem(token, category, index) {
  const session = validateSession(token);
  if (!session.valid) return { success: false, message: session.message, sessionExpired: true };

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("MasterDataProduksi");
    if (!sheet) return { success: false, message: "Sheet MasterDataProduksi tidak ditemukan." };
    const data = sheet.getDataRange().getValues();

    let count = 0;
    let targetRow = -1;
    let targetIdx = parseInt(index);

    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === category) {
        if (count === targetIdx) {
          targetRow = i + 1;
          break;
        }
        count++;
      }
    }

    if (targetRow !== -1) {
      sheet.deleteRow(targetRow);
      return { success: true };
    }
    return { success: false, message: "Target entitas master produksi tidak valid." };
  } catch (err) {
    return { success: false, message: err.message };
  } finally {
    lock.releaseLock();
  }
}

function renameMasterProduksiCategory(token, oldCategory, newCategory) {
  const session = validateSession(token);
  if (!session.valid) return { success: false, message: session.message, sessionExpired: true };

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("MasterDataProduksi");
    if (!sheet) return { success: false, message: "Sheet MasterDataProduksi tidak ditemukan." };
    const data = sheet.getDataRange().getValues();

    newCategory = sanitizeInput(newCategory, 100);

    let updatedCount = 0;
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === oldCategory) {
        sheet.getRange(i + 1, 1).setValue(newCategory);
        updatedCount++;
      }
    }
    return { success: true, count: updatedCount };
  } catch (err) {
    return { success: false, message: err.message };
  } finally {
    lock.releaseLock();
  }
}

function deleteMasterProduksiCategory(token, category) {
  const session = validateSession(token);
  if (!session.valid) return { success: false, message: session.message, sessionExpired: true };

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("MasterDataProduksi");
    if (!sheet) return { success: false, message: "Sheet MasterDataProduksi tidak ditemukan." };
    const data = sheet.getDataRange().getValues();

    let deletedCount = 0;
    for (let i = data.length - 1; i >= 1; i--) {
      if (data[i][0] === category) {
        sheet.deleteRow(i + 1);
        deletedCount++;
      }
    }
    return { success: true, count: deletedCount };
  } catch (err) {
    return { success: false, message: err.message };
  } finally {
    lock.releaseLock();
  }
}

function saveMasterProduksiCategory(token, category) {
  const session = validateSession(token);
  if (!session.valid) return { success: false, message: session.message, sessionExpired: true };

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName("MasterDataProduksi");
    if (!sheet) {
      sheet = ss.insertSheet("MasterDataProduksi");
      sheet.appendRow(["Category", "Value"]);
    }

    category = sanitizeInput(category, 100);
    sheet.appendRow([category, ""]);
    return { success: true };
  } catch (err) {
    return { success: false, message: err.message };
  } finally {
    lock.releaseLock();
  }
}

function syncKamarPenghuniCount(ss) {
  const kamarSheet = ss.getSheetByName("Kamar");
  const penghuniSheet = ss.getSheetByName("Penghuni");
  if (!kamarSheet || !penghuniSheet) return;

  const kamarData = kamarSheet.getDataRange().getValues();
  const penghuniData = penghuniSheet.getDataRange().getValues();

  // Map room number to active occupant count
  const countMap = {};
  for (let i = 1; i < penghuniData.length; i++) {
    const roomNo = String(penghuniData[i][7] || "").trim(); // Nomor Kamar
    const checkOut = penghuniData[i][9]; // Tanggal Check Out
    const isActive = !checkOut || checkOut === "" || String(checkOut).toLowerCase() === "invalid date";
    if (roomNo && isActive) {
      countMap[roomNo] = (countMap[roomNo] || 0) + 1;
    }
  }

  // Update Kamar sheet
  const headers = kamarData[0].map(h => String(h).trim());
  const roomNoIdx = headers.indexOf("Nomor Kamar");
  const countIdx = headers.indexOf("Jumlah Penghuni");

  if (roomNoIdx !== -1 && countIdx !== -1) {
    for (let i = 1; i < kamarData.length; i++) {
      const roomNo = String(kamarData[i][roomNoIdx]).trim();
      const actualCount = countMap[roomNo] || 0;
      kamarSheet.getRange(i + 1, countIdx + 1).setValue(actualCount);
    }
  }
}

// ================= CRUD: MASTER BARANG (LOGISTIK) OPERATIONS =================
function saveMasterBarang(token, payload) {
  const session = validateSession(token);
  if (!session.valid) return { success: false, message: session.message, sessionExpired: true };

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName("MasterBarang");
    if (!sheet) {
      sheet = ss.insertSheet("MasterBarang");
      sheet.appendRow([
        "ID Barang", "Kode Barang", "Nama Barang", "Kategori", "Merk", "Satuan", "Minimum Stok", "Stok Saat Ini", "Lokasi Penyimpanan", "Keterangan"
      ]);
    } else {
      const firstRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => String(h).trim());
      if (firstRow.indexOf("Stok Saat Ini") === -1) {
        sheet.insertColumnAfter(7);
        sheet.getRange(1, 8).setValue("Stok Saat Ini");
      }
    }
    const data = sheet.getDataRange().getValues();

    payload.kodeBarang = sanitizeInput(payload.kodeBarang, 50);
    payload.namaBarang = sanitizeInput(payload.namaBarang, 150);
    payload.kategori = sanitizeInput(payload.kategori, 100);
    payload.merk = sanitizeInput(payload.merk, 100);
    payload.satuan = sanitizeInput(payload.satuan, 50);
    payload.lokasi = sanitizeInput(payload.lokasi, 150);
    payload.keterangan = sanitizeInput(payload.keterangan, 500);
    payload.stokSaatIni = parseFloat(payload.stokSaatIni) || 0;
    payload.minStok = parseFloat(payload.minStok) || 0;

    let rowIdx = -1;
    if (payload.id) {
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][0]) === String(payload.id)) {
          rowIdx = i + 1;
          break;
        }
      }
    }

    if (rowIdx === -1) {
      payload.id = "BRG" + String(data.length).padStart(4, "0");
      sheet.appendRow([
        payload.id, payload.kodeBarang, payload.namaBarang, payload.kategori, payload.merk,
        payload.satuan, payload.minStok, payload.stokSaatIni, payload.lokasi, payload.keterangan
      ]);
    } else {
      sheet.getRange(rowIdx, 2).setValue(payload.kodeBarang);
      sheet.getRange(rowIdx, 3).setValue(payload.namaBarang);
      sheet.getRange(rowIdx, 4).setValue(payload.kategori);
      sheet.getRange(rowIdx, 5).setValue(payload.merk);
      sheet.getRange(rowIdx, 6).setValue(payload.satuan);
      sheet.getRange(rowIdx, 7).setValue(payload.minStok);
      sheet.getRange(rowIdx, 8).setValue(payload.stokSaatIni);
      sheet.getRange(rowIdx, 9).setValue(payload.lokasi);
      sheet.getRange(rowIdx, 10).setValue(payload.keterangan);
    }

    return { success: true };
  } catch (err) {
    return { success: false, message: err.message };
  } finally {
    lock.releaseLock();
  }
}

function deleteMasterBarang(token, barangId) {
  const session = validateSession(token);
  if (!session.valid) return { success: false, message: session.message, sessionExpired: true };

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("MasterBarang");
    if (!sheet) return { success: false, message: "Sheet MasterBarang tidak ditemukan." };
    const data = sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(barangId)) {
        sheet.deleteRow(i + 1);
        return { success: true };
      }
    }
    return { success: false, message: "ID Barang tidak ditemukan." };
  } catch (err) {
    return { success: false, message: err.message };
  } finally {
    lock.releaseLock();
  }
}

// ================= CRUD: SUPPLIER (LOGISTIK) OPERATIONS =================
function saveSupplier(token, payload) {
  const session = validateSession(token);
  if (!session.valid) return { success: false, message: session.message, sessionExpired: true };

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName("Supplier");
    if (!sheet) {
      sheet = ss.insertSheet("Supplier");
      sheet.appendRow([
        "ID Supplier", "Nama Supplier", "PIC", "Nomor Telepon", "Alamat", "Status"
      ]);
    }
    const data = sheet.getDataRange().getValues();

    payload.namaSupplier = sanitizeInput(payload.namaSupplier, 150);
    payload.pic = sanitizeInput(payload.pic, 100);
    payload.telepon = sanitizeInput(payload.telepon, 50);
    payload.alamat = sanitizeInput(payload.alamat, 500);
    payload.status = sanitizeInput(payload.status, 50);

    let rowIdx = -1;
    if (payload.id) {
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][0]) === String(payload.id)) {
          rowIdx = i + 1;
          break;
        }
      }
    }

    if (rowIdx === -1) {
      payload.id = "SUP" + String(data.length).padStart(4, "0");
      sheet.appendRow([
        payload.id, payload.namaSupplier, payload.pic, payload.telepon, payload.alamat, payload.status
      ]);
    } else {
      sheet.getRange(rowIdx, 2).setValue(payload.namaSupplier);
      sheet.getRange(rowIdx, 3).setValue(payload.pic);
      sheet.getRange(rowIdx, 4).setValue(payload.telepon);
      sheet.getRange(rowIdx, 5).setValue(payload.alamat);
      sheet.getRange(rowIdx, 6).setValue(payload.status);
    }

    return { success: true };
  } catch (err) {
    return { success: false, message: err.message };
  } finally {
    lock.releaseLock();
  }
}

function deleteSupplier(token, supplierId) {
  const session = validateSession(token);
  if (!session.valid) return { success: false, message: session.message, sessionExpired: true };

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Supplier");
    if (!sheet) return { success: false, message: "Sheet Supplier tidak ditemukan." };
    const data = sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(supplierId)) {
        sheet.deleteRow(i + 1);
        return { success: true };
      }
    }
    return { success: false, message: "ID Supplier tidak ditemukan." };
  } catch (err) {
    return { success: false, message: err.message };
  } finally {
    lock.releaseLock();
  }
}

// ================= CRUD: BARANG MASUK (LOGISTIK) OPERATIONS =================
function saveBarangMasuk(token, payload) {
  const session = validateSession(token);
  if (!session.valid) return { success: false, message: session.message, sessionExpired: true };

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName("BarangMasuk");
    if (!sheet) {
      sheet = ss.insertSheet("BarangMasuk");
      sheet.appendRow([
        "ID Transaksi", "Tanggal", "No Transaksi", "Kode Barang", "Nama Barang", "Jumlah", "Satuan", "Supplier", "Penerima", "Keterangan"
      ]);
    }
    const data = sheet.getDataRange().getValues();

    payload.tanggal = sanitizeInput(payload.tanggal, 50);
    payload.noTransaksi = sanitizeInput(payload.noTransaksi, 100);
    payload.kodeBarang = sanitizeInput(payload.kodeBarang, 50);
    payload.namaBarang = sanitizeInput(payload.namaBarang, 150);
    payload.jumlah = parseFloat(payload.jumlah) || 0;
    payload.satuan = sanitizeInput(payload.satuan, 50);
    payload.supplier = sanitizeInput(payload.supplier, 150);
    payload.penerima = sanitizeInput(payload.penerima, 100);
    payload.keterangan = sanitizeInput(payload.keterangan, 500);

    let rowIdx = -1;
    let oldJumlah = 0;

    if (payload.id) {
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][0]) === String(payload.id)) {
          rowIdx = i + 1;
          oldJumlah = parseFloat(data[i][5]) || 0;
          break;
        }
      }
    }

    if (rowIdx === -1) {
      payload.id = "IN" + String(data.length).padStart(4, "0");
      sheet.appendRow([
        payload.id, payload.tanggal, payload.noTransaksi, payload.kodeBarang, payload.namaBarang,
        payload.jumlah, payload.satuan, payload.supplier, payload.penerima, payload.keterangan
      ]);
    } else {
      sheet.getRange(rowIdx, 2).setValue(payload.tanggal);
      sheet.getRange(rowIdx, 3).setValue(payload.noTransaksi);
      sheet.getRange(rowIdx, 4).setValue(payload.kodeBarang);
      sheet.getRange(rowIdx, 5).setValue(payload.namaBarang);
      sheet.getRange(rowIdx, 6).setValue(payload.jumlah);
      sheet.getRange(rowIdx, 7).setValue(payload.satuan);
      sheet.getRange(rowIdx, 8).setValue(payload.supplier);
      sheet.getRange(rowIdx, 9).setValue(payload.penerima);
      sheet.getRange(rowIdx, 10).setValue(payload.keterangan);
    }

    // Automatically Update Stok Saat Ini in MasterBarang sheet
    const mbSheet = ss.getSheetByName("MasterBarang");
    if (mbSheet) {
      const mbData = mbSheet.getDataRange().getValues();
      const headers = mbData[0].map(h => String(h).trim());
      const kodeIdx = headers.indexOf("Kode Barang") !== -1 ? headers.indexOf("Kode Barang") : 1;
      const stokIdx = headers.indexOf("Stok Saat Ini") !== -1 ? headers.indexOf("Stok Saat Ini") : 7;

      for (let i = 1; i < mbData.length; i++) {
        if (String(mbData[i][kodeIdx]).trim() === String(payload.kodeBarang).trim()) {
          let currentStok = parseFloat(mbData[i][stokIdx]) || 0;
          let newStok = currentStok + (payload.jumlah - oldJumlah);
          if (newStok < 0) newStok = 0;
          mbSheet.getRange(i + 1, stokIdx + 1).setValue(newStok);
          break;
        }
      }
    }

    return { success: true };
  } catch (err) {
    return { success: false, message: err.message };
  } finally {
    lock.releaseLock();
  }
}

function deleteBarangMasuk(token, id) {
  const session = validateSession(token);
  if (!session.valid) return { success: false, message: session.message, sessionExpired: true };

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("BarangMasuk");
    if (!sheet) return { success: false, message: "Sheet BarangMasuk tidak ditemukan." };
    const data = sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(id)) {
        const kodeBarang = data[i][3];
        const jumlah = parseFloat(data[i][5]) || 0;

        // Revert stock in MasterBarang
        const mbSheet = ss.getSheetByName("MasterBarang");
        if (mbSheet) {
          const mbData = mbSheet.getDataRange().getValues();
          const headers = mbData[0].map(h => String(h).trim());
          const kodeIdx = headers.indexOf("Kode Barang") !== -1 ? headers.indexOf("Kode Barang") : 1;
          const stokIdx = headers.indexOf("Stok Saat Ini") !== -1 ? headers.indexOf("Stok Saat Ini") : 7;

          for (let j = 1; j < mbData.length; j++) {
            if (String(mbData[j][kodeIdx]).trim() === String(kodeBarang).trim()) {
              let currentStok = parseFloat(mbData[j][stokIdx]) || 0;
              let newStok = currentStok - jumlah;
              if (newStok < 0) newStok = 0;
              mbSheet.getRange(j + 1, stokIdx + 1).setValue(newStok);
              break;
            }
          }
        }

        sheet.deleteRow(i + 1);
        return { success: true };
      }
    }
    return { success: false, message: "ID Transaksi tidak ditemukan." };
  } catch (err) {
    return { success: false, message: err.message };
  } finally {
    lock.releaseLock();
  }
}

// ================= CRUD: BARANG KELUAR (LOGISTIK) OPERATIONS =================
function saveBarangKeluar(token, payload) {
  const session = validateSession(token);
  if (!session.valid) return { success: false, message: session.message, sessionExpired: true };

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName("BarangKeluar");
    if (!sheet) {
      sheet = ss.insertSheet("BarangKeluar");
      sheet.appendRow([
        "ID Transaksi", "Tanggal", "No Transaksi", "Kode Barang", "Nama Barang", "Jumlah", "Satuan", "Departemen", "Jabatan Penerima", "Penerima", "Keperluan", "Keterangan"
      ]);
    }
    const data = sheet.getDataRange().getValues();

    payload.tanggal = sanitizeInput(payload.tanggal, 50);
    payload.noTransaksi = sanitizeInput(payload.noTransaksi, 100);
    payload.kodeBarang = sanitizeInput(payload.kodeBarang, 50);
    payload.namaBarang = sanitizeInput(payload.namaBarang, 150);
    payload.jumlah = parseFloat(payload.jumlah) || 0;
    payload.satuan = sanitizeInput(payload.satuan, 50);
    payload.departemen = sanitizeInput(payload.departemen, 100);
    payload.jabatanPenerima = sanitizeInput(payload.jabatanPenerima, 100);
    payload.penerima = sanitizeInput(payload.penerima, 100);
    payload.keperluan = sanitizeInput(payload.keperluan, 250);
    payload.keterangan = sanitizeInput(payload.keterangan, 500);

    let rowIdx = -1;
    let oldJumlah = 0;

    if (payload.id) {
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][0]) === String(payload.id)) {
          rowIdx = i + 1;
          oldJumlah = parseFloat(data[i][5]) || 0;
          break;
        }
      }
    }

    if (rowIdx === -1) {
      payload.id = "OUT" + String(data.length).padStart(4, "0");
      sheet.appendRow([
        payload.id, payload.tanggal, payload.noTransaksi, payload.kodeBarang, payload.namaBarang,
        payload.jumlah, payload.satuan, payload.departemen, payload.jabatanPenerima, payload.penerima, payload.keperluan, payload.keterangan
      ]);
    } else {
      sheet.getRange(rowIdx, 2).setValue(payload.tanggal);
      sheet.getRange(rowIdx, 3).setValue(payload.noTransaksi);
      sheet.getRange(rowIdx, 4).setValue(payload.kodeBarang);
      sheet.getRange(rowIdx, 5).setValue(payload.namaBarang);
      sheet.getRange(rowIdx, 6).setValue(payload.jumlah);
      sheet.getRange(rowIdx, 7).setValue(payload.satuan);
      sheet.getRange(rowIdx, 8).setValue(payload.departemen);
      sheet.getRange(rowIdx, 9).setValue(payload.jabatanPenerima);
      sheet.getRange(rowIdx, 10).setValue(payload.penerima);
      sheet.getRange(rowIdx, 11).setValue(payload.keperluan);
      sheet.getRange(rowIdx, 12).setValue(payload.keterangan);
    }

    // Automatically Update Stok Saat Ini in MasterBarang sheet (deduct stock for barang keluar)
    const mbSheet = ss.getSheetByName("MasterBarang");
    if (mbSheet) {
      const mbData = mbSheet.getDataRange().getValues();
      const headers = mbData[0].map(h => String(h).trim());
      const kodeIdx = headers.indexOf("Kode Barang") !== -1 ? headers.indexOf("Kode Barang") : 1;
      const stokIdx = headers.indexOf("Stok Saat Ini") !== -1 ? headers.indexOf("Stok Saat Ini") : 7;

      for (let i = 1; i < mbData.length; i++) {
        if (String(mbData[i][kodeIdx]).trim() === String(payload.kodeBarang).trim()) {
          let currentStok = parseFloat(mbData[i][stokIdx]) || 0;
          let newStok = currentStok - (payload.jumlah - oldJumlah);
          if (newStok < 0) newStok = 0;
          mbSheet.getRange(i + 1, stokIdx + 1).setValue(newStok);
          break;
        }
      }
    }

    return { success: true };
  } catch (err) {
    return { success: false, message: err.message };
  } finally {
    lock.releaseLock();
  }
}

function deleteBarangKeluar(token, id) {
  const session = validateSession(token);
  if (!session.valid) return { success: false, message: session.message, sessionExpired: true };

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("BarangKeluar");
    if (!sheet) return { success: false, message: "Sheet BarangKeluar tidak ditemukan." };
    const data = sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(id)) {
        const kodeBarang = data[i][3];
        const jumlah = parseFloat(data[i][5]) || 0;

        // Revert stock in MasterBarang (add back stock)
        const mbSheet = ss.getSheetByName("MasterBarang");
        if (mbSheet) {
          const mbData = mbSheet.getDataRange().getValues();
          const headers = mbData[0].map(h => String(h).trim());
          const kodeIdx = headers.indexOf("Kode Barang") !== -1 ? headers.indexOf("Kode Barang") : 1;
          const stokIdx = headers.indexOf("Stok Saat Ini") !== -1 ? headers.indexOf("Stok Saat Ini") : 7;

          for (let j = 1; j < mbData.length; j++) {
            if (String(mbData[j][kodeIdx]).trim() === String(kodeBarang).trim()) {
              let currentStok = parseFloat(mbData[j][stokIdx]) || 0;
              let newStok = currentStok + jumlah;
              mbSheet.getRange(j + 1, stokIdx + 1).setValue(newStok);
              break;
            }
          }
        }

        sheet.deleteRow(i + 1);
        return { success: true };
      }
    }
    return { success: false, message: "ID Transaksi tidak ditemukan." };
  } catch (err) {
    return { success: false, message: err.message };
  } finally {
    lock.releaseLock();
  }
}

// ================= CRUD: STOCK OPNAME (LOGISTIK) OPERATIONS =================
function saveStockOpname(token, payload) {
  const session = validateSession(token);
  if (!session.valid) return { success: false, message: session.message, sessionExpired: true };

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName("StockOpname");
    if (!sheet) {
      sheet = ss.insertSheet("StockOpname");
      sheet.appendRow([
        "ID Opname", "Tanggal", "Kode Barang", "Nama Barang", "Stok Sistem", "Stok Fisik", "Selisih", "Satuan", "Petugas", "Keterangan"
      ]);
    }
    const data = sheet.getDataRange().getValues();

    payload.tanggal = sanitizeInput(payload.tanggal, 50);
    payload.kodeBarang = sanitizeInput(payload.kodeBarang, 50);
    payload.namaBarang = sanitizeInput(payload.namaBarang, 150);
    payload.stokSistem = parseFloat(payload.stokSistem) || 0;
    payload.stokFisik = parseFloat(payload.stokFisik) || 0;
    payload.selisih = payload.stokFisik - payload.stokSistem;
    payload.satuan = sanitizeInput(payload.satuan, 50);
    payload.petugas = sanitizeInput(payload.petugas, 100);
    payload.keterangan = sanitizeInput(payload.keterangan, 500);

    let rowIdx = -1;

    if (payload.id) {
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][0]) === String(payload.id)) {
          rowIdx = i + 1;
          break;
        }
      }
    }

    if (rowIdx === -1) {
      payload.id = "SO" + String(data.length).padStart(4, "0");
      sheet.appendRow([
        payload.id, payload.tanggal, payload.kodeBarang, payload.namaBarang,
        payload.stokSistem, payload.stokFisik, payload.selisih, payload.satuan, payload.petugas, payload.keterangan
      ]);
    } else {
      sheet.getRange(rowIdx, 2).setValue(payload.tanggal);
      sheet.getRange(rowIdx, 3).setValue(payload.kodeBarang);
      sheet.getRange(rowIdx, 4).setValue(payload.namaBarang);
      sheet.getRange(rowIdx, 5).setValue(payload.stokSistem);
      sheet.getRange(rowIdx, 6).setValue(payload.stokFisik);
      sheet.getRange(rowIdx, 7).setValue(payload.selisih);
      sheet.getRange(rowIdx, 8).setValue(payload.satuan);
      sheet.getRange(rowIdx, 9).setValue(payload.petugas);
      sheet.getRange(rowIdx, 10).setValue(payload.keterangan);
    }

    // Adjust Stok Saat Ini in MasterBarang sheet to match Stok Fisik
    const mbSheet = ss.getSheetByName("MasterBarang");
    if (mbSheet) {
      const mbData = mbSheet.getDataRange().getValues();
      const headers = mbData[0].map(h => String(h).trim());
      const kodeIdx = headers.indexOf("Kode Barang") !== -1 ? headers.indexOf("Kode Barang") : 1;
      const stokIdx = headers.indexOf("Stok Saat Ini") !== -1 ? headers.indexOf("Stok Saat Ini") : 7;

      for (let i = 1; i < mbData.length; i++) {
        if (String(mbData[i][kodeIdx]).trim() === String(payload.kodeBarang).trim()) {
          mbSheet.getRange(i + 1, stokIdx + 1).setValue(payload.stokFisik);
          break;
        }
      }
    }

    return { success: true };
  } catch (err) {
    return { success: false, message: err.message };
  } finally {
    lock.releaseLock();
  }
}

function deleteStockOpname(token, id) {
  const session = validateSession(token);
  if (!session.valid) return { success: false, message: session.message, sessionExpired: true };

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("StockOpname");
    if (!sheet) return { success: false, message: "Sheet StockOpname tidak ditemukan." };
    const data = sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(id)) {
        sheet.deleteRow(i + 1);
        return { success: true };
      }
    }
    return { success: false, message: "ID Opname tidak ditemukan." };
  } catch (err) {
    return { success: false, message: err.message };
  } finally {
    lock.releaseLock();
  }
}