/**
 * ================================================================================
 * MANAJEMEN KONTAK — Google Apps Script Backend
 * Sheet: Config_Contacts
 * API: People API (Advanced Service: People)
 * ================================================================================
 *
 * AKTIVASI PEOPLE API:
 *   1. Di Apps Script Editor: Extensions → Apps Script → Services → (+) Add service
 *      → cari "People API" → klik Add.
 *   2. Pastikan appsscript.json sudah memuat scope:
 *      "https://www.googleapis.com/auth/contacts"
 *      "https://www.googleapis.com/auth/contacts.readonly"
 *   3. Re-deploy web app setelah menambahkan service.
 * ================================================================================
 */

var CONTACTS_SHEET_NAME = "Config_Contacts";

var CONTACTS_HEADERS = [
  "Contact_ID",
  "Nama_Display",
  "Email",
  "Contact_Group",
  "Status",
  "Sync_Mode",
  "Google_Resource_Name",
  "Google_ETag",
  "Last_Sync",
  "Sync_Status",
  "Sync_Message"
];

// -----------------------------------------------------------------------
// SETUP
// -----------------------------------------------------------------------

/**
 * Membuat atau memvalidasi sheet Config_Contacts beserta header-nya.
 */
function setupContactsSheet() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(CONTACTS_SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(CONTACTS_SHEET_NAME);
  }

  normalizeContactsSheet_(sheet);

  SpreadsheetApp.flush();
  return { success: true, message: "Sheet Config_Contacts siap digunakan dengan kolom terbaru." };
}

/**
 * Menormalkan Config_Contacts ke skema ringkas:
 * Contact_ID, Nama_Display, Email, group, dan metadata sync.
 * Data email dari skema lama tetap dimigrasikan ke kolom Email.
 */
function normalizeContactsSheet_(sheet) {
  var targetHeaderMap = {};
  for (var h = 0; h < CONTACTS_HEADERS.length; h++) {
    targetHeaderMap[CONTACTS_HEADERS[h]] = h;
  }

  var lastCol = Math.max(sheet.getLastColumn(), CONTACTS_HEADERS.length);
  var firstRow = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var currentHeaders = firstRow.map(function(v) { return String(v || "").trim(); });
  var currentLeanHeaders = currentHeaders.slice(0, CONTACTS_HEADERS.length);
  var alreadyNormalized = currentLeanHeaders.join("|") === CONTACTS_HEADERS.join("|") &&
    sheet.getLastColumn() === CONTACTS_HEADERS.length;

  if (!alreadyNormalized) {
    var oldHeaderMap = {};
    for (var i = 0; i < currentHeaders.length; i++) {
      if (currentHeaders[i]) oldHeaderMap[currentHeaders[i]] = i;
    }

    var outputRows = [];
    var lastRow = sheet.getLastRow();
    if (lastRow > 1 && Object.keys(oldHeaderMap).length > 0) {
      var sourceValues = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
      for (var r = 0; r < sourceValues.length; r++) {
        var row = sourceValues[r];
        var hasData = row.some(function(cell) { return String(cell || "").trim() !== ""; });
        if (!hasData) continue;

        var obj = rowArrayToObject_(row, oldHeaderMap);
        obj["Email"] = getContactEmail_(obj) || getRowContactEmail_(row, oldHeaderMap);
        outputRows.push(buildRowFromObject_(obj, targetHeaderMap, CONTACTS_HEADERS.length));
      }
    }

    sheet.clear();
    if (sheet.getMaxColumns() < CONTACTS_HEADERS.length) {
      sheet.insertColumnsAfter(sheet.getMaxColumns(), CONTACTS_HEADERS.length - sheet.getMaxColumns());
    } else if (sheet.getMaxColumns() > CONTACTS_HEADERS.length) {
      sheet.deleteColumns(CONTACTS_HEADERS.length + 1, sheet.getMaxColumns() - CONTACTS_HEADERS.length);
    }

    sheet.getRange(1, 1, 1, CONTACTS_HEADERS.length).setValues([CONTACTS_HEADERS]);
    if (outputRows.length > 0) {
      sheet.getRange(2, 1, outputRows.length, CONTACTS_HEADERS.length).setValues(outputRows);
    }
  }

  sheet.getRange(1, 1, 1, CONTACTS_HEADERS.length)
    .setFontWeight("bold")
    .setBackground("#015850")
    .setFontColor("#FFEB2F")
    .setHorizontalAlignment("center");
  sheet.setFrozenRows(1);

  var lastSyncColIdx = CONTACTS_HEADERS.indexOf("Last_Sync") + 1;
  if (lastSyncColIdx > 0 && sheet.getMaxRows() > 1) {
    sheet.getRange(2, lastSyncColIdx, sheet.getMaxRows() - 1, 1)
      .setNumberFormat("yyyy-MM-dd HH:mm:ss");
  }

  try {
    sheet.autoResizeColumns(1, CONTACTS_HEADERS.length);
  } catch (e) { /* Abaikan jika gagal */ }
}

// -----------------------------------------------------------------------
// BACA DATA
// -----------------------------------------------------------------------

/**
 * Mengambil semua data kontak dari sheet.
 * @returns {Array<Object>} Array of row objects
 */
function getContactsData() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(CONTACTS_SHEET_NAME);

  if (!sheet) {
    setupContactsSheet();
    sheet = ss.getSheetByName(CONTACTS_SHEET_NAME);
    if (!sheet) {
      return { success: false, message: "Sheet Config_Contacts belum bisa dibuat otomatis." };
    }
  }

  normalizeContactsSheet_(sheet);
  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) {
    return { success: true, data: [] };
  }

  var headerMap = getHeaderMap_(sheet);
  var numCols = CONTACTS_HEADERS.length;
  var dataRange = sheet.getRange(2, 1, lastRow - 1, numCols);
  var values = dataRange.getDisplayValues();

  var rows = [];
  for (var i = 0; i < values.length; i++) {
    var row = values[i];
    var hasData = row.some(function(cell) { return String(cell || "").trim() !== ""; });
    if (!hasData) continue;

    var obj = {};
    for (var col in headerMap) {
      obj[col] = String(row[headerMap[col]] !== undefined ? row[headerMap[col]] : "").trim();
    }
    obj._rowIndex = i + 2;
    rows.push(obj);
  }

  return { success: true, data: rows };
}

// -----------------------------------------------------------------------
// SIMPAN DATA
// -----------------------------------------------------------------------

/**
 * Menyimpan satu kontak (tambah baru atau update).
 * Tidak melakukan sync ke Google Contacts secara langsung.
 * @param {Object} contactObject
 */
function saveContactData(contactObject) {
  if (!contactObject) {
    return { success: false, message: "Data kontak tidak boleh kosong." };
  }

  var validation = validateContactRow_(contactObject);
  if (!validation.valid) {
    return { success: false, message: validation.message };
  }

  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(CONTACTS_SHEET_NAME);
  if (!sheet) {
    setupContactsSheet();
    sheet = ss.getSheetByName(CONTACTS_SHEET_NAME);
  }

  normalizeContactsSheet_(sheet);
  var headerMap = getHeaderMap_(sheet);
  var contactId = String(contactObject["Contact_ID"] || "").trim();
  var contactEmail = normalizeEmail_(getContactEmail_(contactObject));
  var targetRow = 0;
  var existingObject = null;

  if (contactId) {
    // Cari baris berdasarkan Contact_ID
    var lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      var idColIdx = headerMap["Contact_ID"];
      if (idColIdx !== undefined) {
        var ids = sheet.getRange(2, idColIdx + 1, lastRow - 1, 1).getValues();
        for (var i = 0; i < ids.length; i++) {
          if (String(ids[i][0] || "").trim() === contactId) {
            targetRow = i + 2;
            break;
          }
        }
      }
    }
  }

  if (contactEmail) {
    var emailLastRow = sheet.getLastRow();
    if (emailLastRow > 1) {
      var existingRows = sheet.getRange(2, 1, emailLastRow - 1, CONTACTS_HEADERS.length).getValues();
      var existingIdIdx = headerMap["Contact_ID"];
      var existingNameIdx = headerMap["Nama_Display"];
      for (var e = 0; e < existingRows.length; e++) {
        var rowIndexByEmail = e + 2;
        if (rowIndexByEmail === targetRow) continue;

        if (normalizeEmail_(getRowContactEmail_(existingRows[e], headerMap)) === contactEmail) {
          var existingName = existingNameIdx !== undefined ? String(existingRows[e][existingNameIdx] || "").trim() : "";
          var existingId = existingIdIdx !== undefined ? String(existingRows[e][existingIdIdx] || "").trim() : "";
          return {
            success: false,
            duplicateEmail: true,
            message: "Email " + getContactEmail_(contactObject) + " sudah pernah didaftarkan" +
              (existingName ? " untuk kontak " + existingName : "") +
              (existingId ? " (" + existingId + ")" : "") + "."
          };
        }
      }
    }
  }

  var isNew = targetRow === 0;
  if (isNew) {
    contactId = contactId || createRecordId("CTK");
    contactObject["Contact_ID"] = contactId;
    targetRow = sheet.getLastRow() + 1;
  } else {
    var existingValues = sheet.getRange(targetRow, 1, 1, CONTACTS_HEADERS.length).getValues()[0];
    existingObject = rowArrayToObject_(existingValues, headerMap);
    for (var key in contactObject) {
      if (contactObject[key] !== undefined) {
        existingObject[key] = contactObject[key];
      }
    }
    existingObject["Contact_ID"] = contactId;
    contactObject = existingObject;
  }

  contactObject["Status"] = "Aktif";

  // Jika ada perubahan data utama, tandai NEED_SYNC (kecuali kalau sedang di-set dari proses sync)
  var syncMode = String(contactObject["Sync_Mode"] || "AUTO").trim();
  if (syncMode !== "NO_SYNC") {
    contactObject["Sync_Status"] = "NEED_SYNC";
    contactObject["Sync_Message"] = isNew ? "Data baru, belum disync." : "Data diperbarui, perlu sync.";
  } else {
    contactObject["Sync_Status"] = "";
    contactObject["Sync_Message"] = "Mode NO_SYNC: kontak tidak ikut sinkronisasi.";
  }

  // Tulis ke sheet
  var numCols = CONTACTS_HEADERS.length;
  var rowValues = buildRowFromObject_(contactObject, headerMap, numCols);
  sheet.getRange(targetRow, 1, 1, numCols).setValues([rowValues]);
  SpreadsheetApp.flush();

  logAudit(isNew ? "Tambah Kontak" : "Edit Kontak", contactId, "SAVE", "", "", Session.getActiveUser().getEmail());

  return {
    success: true,
    message: isNew ? "Kontak berhasil ditambahkan." : "Kontak berhasil diperbarui.",
    contactId: contactId,
    rowIndex: targetRow
  };
}

/**
 * Menyimpan banyak kontak sekaligus (untuk import bulk).
 * @param {Array<Object>} contactArray
 */
function saveBulkContacts(contactArray) {
  if (!Array.isArray(contactArray) || contactArray.length === 0) {
    return { success: false, message: "Tidak ada data yang dikirim." };
  }

  var results = { success: true, saved: 0, failed: 0, errors: [] };

  for (var i = 0; i < contactArray.length; i++) {
    var res = saveContactData(contactArray[i]);
    if (res.success) {
      results.saved++;
    } else {
      results.failed++;
      results.errors.push("Baris " + (i + 1) + ": " + res.message);
    }
  }

  results.message = "Selesai. Tersimpan: " + results.saved + ", gagal: " + results.failed + ".";
  return results;
}

/**
 * Hapus kontak dari sheet, opsional sekalian hapus dari Google Contacts.
 * @param {string} contactId
 * @param {boolean} deleteInGoogle
 */
function deleteContactData(contactId, deleteInGoogle) {
  contactId = String(contactId || "").trim();
  if (!contactId) return { success: false, message: "Contact_ID tidak boleh kosong." };

  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(CONTACTS_SHEET_NAME);
  if (!sheet) return { success: false, message: "Sheet Config_Contacts tidak ditemukan." };

  normalizeContactsSheet_(sheet);
  var headerMap = getHeaderMap_(sheet);
  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) return { success: false, message: "Tidak ada data kontak." };

  var idColIdx = headerMap["Contact_ID"];
  if (idColIdx === undefined) return { success: false, message: "Kolom Contact_ID tidak ditemukan." };

  var ids = sheet.getRange(2, idColIdx + 1, lastRow - 1, 1).getValues();
  var targetRow = 0;
  for (var i = 0; i < ids.length; i++) {
    if (String(ids[i][0] || "").trim() === contactId) {
      targetRow = i + 2;
      break;
    }
  }
  if (!targetRow) return { success: false, message: "Kontak tidak ditemukan." };

  if (deleteInGoogle) {
    var googleResIdx = headerMap["Google_Resource_Name"];
    if (googleResIdx !== undefined) {
      var resourceName = String(sheet.getRange(targetRow, googleResIdx + 1).getValue() || "").trim();
      if (resourceName) {
        try {
          People.People.deleteContact(resourceName);
        } catch (err) {
          var em = String(err.message || err);
          if (em.indexOf("404") === -1 && em.toLowerCase().indexOf("not found") === -1) {
            return { success: false, message: "Gagal menghapus kontak di Google Contacts: " + em.substring(0, 200) };
          }
        }
      }
    }
  }

  sheet.deleteRow(targetRow);
  SpreadsheetApp.flush();

  logAudit("Hapus Kontak", contactId, "DELETE", "", deleteInGoogle ? "Sheet + Google Contacts" : "Sheet only", Session.getActiveUser().getEmail());
  return { success: true, message: "Kontak berhasil dihapus." };
}

// -----------------------------------------------------------------------
// TANDAI NEED SYNC
// -----------------------------------------------------------------------

/**
 * Tandai satu kontak sebagai NEED_SYNC berdasarkan Contact_ID.
 */
function markContactNeedSync(contactId) {
  contactId = String(contactId || "").trim();
  if (!contactId) return { success: false, message: "Contact_ID tidak boleh kosong." };

  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(CONTACTS_SHEET_NAME);
  if (!sheet) return { success: false, message: "Sheet Config_Contacts tidak ditemukan." };

  normalizeContactsSheet_(sheet);
  var headerMap = getHeaderMap_(sheet);
  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) return { success: false, message: "Tidak ada data kontak." };

  var idColIdx = headerMap["Contact_ID"];
  var syncModeColIdx = headerMap["Sync_Mode"];
  var syncStatusColIdx = headerMap["Sync_Status"];
  var syncMsgColIdx = headerMap["Sync_Message"];

  if (idColIdx === undefined || syncStatusColIdx === undefined) {
    return { success: false, message: "Kolom Contact_ID atau Sync_Status tidak ditemukan." };
  }

  var ids = sheet.getRange(2, idColIdx + 1, lastRow - 1, 1).getValues();
  for (var i = 0; i < ids.length; i++) {
    if (String(ids[i][0] || "").trim() === contactId) {
      var targetRow = i + 2;
      var syncMode = syncModeColIdx !== undefined
        ? String(sheet.getRange(targetRow, syncModeColIdx + 1).getValue() || "AUTO").trim().toUpperCase()
        : "AUTO";
      if (syncMode === "NO_SYNC") {
        if (syncMsgColIdx !== undefined) {
          sheet.getRange(targetRow, syncMsgColIdx + 1).setValue("Tidak ditandai: Sync_Mode NO_SYNC.");
        }
        SpreadsheetApp.flush();
        return { success: false, message: "Kontak memakai Sync_Mode NO_SYNC, jadi tidak bisa ditandai NEED_SYNC." };
      }
      sheet.getRange(targetRow, syncStatusColIdx + 1).setValue("NEED_SYNC");
      if (syncMsgColIdx !== undefined) {
        sheet.getRange(targetRow, syncMsgColIdx + 1).setValue("Ditandai manual perlu sync.");
      }
      SpreadsheetApp.flush();
      return { success: true, message: "Kontak ditandai NEED_SYNC." };
    }
  }

  return { success: false, message: "Kontak dengan ID " + contactId + " tidak ditemukan." };
}

/**
 * Tandai semua kontak (kecuali NO_SYNC) sebagai NEED_SYNC.
 */
function markAllContactsNeedSync() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(CONTACTS_SHEET_NAME);
  if (!sheet) return { success: false, message: "Sheet Config_Contacts tidak ditemukan." };

  normalizeContactsSheet_(sheet);
  var headerMap = getHeaderMap_(sheet);
  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) return { success: true, message: "Tidak ada data kontak.", count: 0 };

  var syncModeColIdx = headerMap["Sync_Mode"];
  var syncStatusColIdx = headerMap["Sync_Status"];
  var syncMsgColIdx = headerMap["Sync_Message"];

  if (syncStatusColIdx === undefined) {
    return { success: false, message: "Kolom Sync_Status tidak ditemukan." };
  }

  var numRows = lastRow - 1;
  var syncModes = syncModeColIdx !== undefined
    ? sheet.getRange(2, syncModeColIdx + 1, numRows, 1).getValues()
    : null;

  var count = 0;
  for (var i = 0; i < numRows; i++) {
    var mode = syncModes ? String(syncModes[i][0] || "AUTO").trim().toUpperCase() : "AUTO";
    if (mode === "NO_SYNC") continue;
    sheet.getRange(i + 2, syncStatusColIdx + 1).setValue("NEED_SYNC");
    if (syncMsgColIdx !== undefined) {
      sheet.getRange(i + 2, syncMsgColIdx + 1).setValue("Tandai semua: perlu sync ulang.");
    }
    count++;
  }

  SpreadsheetApp.flush();
  return { success: true, message: count + " kontak ditandai NEED_SYNC.", count: count };
}

// -----------------------------------------------------------------------
// SYNC KE GOOGLE CONTACTS
// -----------------------------------------------------------------------

/**
 * Sinkronisasi semua kontak berstatus NEED_SYNC ke Google Contacts.
 * Jalankan via tombol UI atau trigger harian.
 */
function syncContactsToGoogleContacts() {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(60000)) {
    return { success: false, message: "Sync kontak sedang berjalan oleh proses lain. Coba lagi nanti." };
  }

  var result = {
    success: true,
    synced: 0,
    skipped: 0,
    failed: 0,
    errors: [],
    message: ""
  };

  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(CONTACTS_SHEET_NAME);
    if (!sheet) {
      setupContactsSheet();
      sheet = ss.getSheetByName(CONTACTS_SHEET_NAME);
      if (!sheet) {
        return { success: false, message: "Sheet Config_Contacts belum bisa dibuat otomatis." };
      }
    }

    normalizeContactsSheet_(sheet);
    var lastRow = sheet.getLastRow();
    if (lastRow <= 1) {
      return { success: true, message: "Tidak ada data kontak untuk disync.", synced: 0, skipped: 0, failed: 0, errors: [] };
    }

    var headerMap = getHeaderMap_(sheet);
    var numCols = CONTACTS_HEADERS.length;
    var dataValues = sheet.getRange(2, 1, lastRow - 1, numCols).getValues();

    for (var i = 0; i < dataValues.length; i++) {
      var rowObj = {};
      for (var col in headerMap) {
        rowObj[col] = String(dataValues[i][headerMap[col]] !== undefined ? dataValues[i][headerMap[col]] : "").trim();
      }
      rowObj._rowIndex = i + 2;

      // Lewati baris kosong
      if (!rowObj["Contact_ID"] && !rowObj["Nama_Display"]) continue;

      var syncMode = (rowObj["Sync_Mode"] || "AUTO").toUpperCase();
      var syncStatus = (rowObj["Sync_Status"] || "").toUpperCase();

      // Sync_Mode NO_SYNC adalah mode pengecualian, bukan Sync_Status.
      if (syncMode === "NO_SYNC") {
        updateSyncStatus_(sheet, rowObj._rowIndex, headerMap, "", "Dilewati karena Sync_Mode NO_SYNC.");
        result.skipped++;
        continue;
      }

      if (syncStatus !== "NEED_SYNC") {
        result.skipped++;
        continue;
      }

      // Validasi wajib
      var validation = validateContactRow_(rowObj);
      if (!validation.valid) {
        updateSyncStatus_(sheet, rowObj._rowIndex, headerMap, "ERROR", validation.message);
        result.failed++;
        result.errors.push("Baris " + rowObj._rowIndex + " (" + rowObj["Nama_Display"] + "): " + validation.message);
        continue;
      }

      try {
        var syncResult;
        if (rowObj["Google_Resource_Name"]) {
          syncResult = updateGoogleContact_(rowObj);
        } else {
          // Cek duplikasi berdasarkan email jabatan
          var existing = findContactByEmail_(getContactEmail_(rowObj));
          if (existing) {
            rowObj["Google_Resource_Name"] = existing.resourceName;
            rowObj["Google_ETag"] = existing.etag || "";
            syncResult = updateGoogleContact_(rowObj);
          } else {
            syncResult = createGoogleContact_(rowObj);
          }
        }

        if (syncResult.success) {
          var groupSync = syncContactGroups_(syncResult.resourceName, rowObj["Contact_Group"]);
          if (!groupSync.success) {
            throw new Error(groupSync.message);
          }
          var syncMessage = groupSync.message;

          var now = new Date();
          var nowStr = Utilities.formatDate(now, "Asia/Jakarta", "yyyy-MM-dd HH:mm:ss");
          updateSyncStatus_(sheet, rowObj._rowIndex, headerMap, "SYNCED", syncMessage, syncResult.resourceName, syncResult.etag, nowStr);
          result.synced++;
        } else {
          updateSyncStatus_(sheet, rowObj._rowIndex, headerMap, "ERROR", syncResult.message);
          result.failed++;
          result.errors.push("Baris " + rowObj._rowIndex + " (" + rowObj["Nama_Display"] + "): " + syncResult.message);
        }

      } catch (syncErr) {
        var errMsg = String(syncErr.message || syncErr.toString()).substring(0, 200);
        updateSyncStatus_(sheet, rowObj._rowIndex, headerMap, "ERROR", errMsg);
        result.failed++;
        result.errors.push("Baris " + rowObj._rowIndex + " (" + rowObj["Nama_Display"] + "): " + errMsg);
        console.error("Sync error baris " + rowObj._rowIndex + ": " + errMsg);
      }
    }

    SpreadsheetApp.flush();
    result.message = "Sync selesai. Berhasil: " + result.synced +
      ", dilewati: " + result.skipped +
      ", gagal: " + result.failed + ".";

    logAudit("Sync Kontak", "-", "SYNC", "", result.message, "SYSTEM");

  } catch (err) {
    result.success = false;
    result.message = "Error kritis saat sync: " + String(err.message || err);
  } finally {
    try { lock.releaseLock(); } catch (e) { /* abaikan */ }
  }

  return result;
}

// -----------------------------------------------------------------------
// GOOGLE CONTACTS CRUD
// -----------------------------------------------------------------------

/**
 * Membuat kontak baru di Google Contacts.
 * @param {Object} rowObject
 * @returns {{success: boolean, resourceName: string, etag: string, message: string}}
 */
function createGoogleContact_(rowObject) {
  try {
    var payload = buildPersonPayload_(rowObject);
    var response = People.People.createContact(payload);
    return {
      success: true,
      resourceName: response.resourceName || "",
      etag: response.etag || "",
      message: "Kontak berhasil dibuat."
    };
  } catch (err) {
    return {
      success: false,
      resourceName: "",
      etag: "",
      message: "Gagal membuat kontak: " + String(err.message || err).substring(0, 200)
    };
  }
}

/**
 * Memperbarui kontak yang sudah ada di Google Contacts.
 * @param {Object} rowObject — harus memiliki Google_Resource_Name
 * @returns {{success: boolean, resourceName: string, etag: string, message: string}}
 */
function updateGoogleContact_(rowObject) {
  var resourceName = String(rowObject["Google_Resource_Name"] || "").trim();
  if (!resourceName) {
    return { success: false, resourceName: "", etag: "", message: "Google_Resource_Name kosong, tidak bisa update." };
  }

  try {
    var payload = buildPersonPayload_(rowObject);
    var updateFields = "names,emailAddresses";

    // Ambil etag terkini jika tidak tersedia
    var etag = String(rowObject["Google_ETag"] || "").trim();
    if (!etag) {
      try {
        var existing = People.People.get(resourceName, { personFields: "metadata" });
        etag = existing.etag || "";
      } catch (e) {
        // Jika resource tidak ditemukan, buat baru
        if (String(e.message || "").indexOf("404") !== -1 || String(e.message || "").indexOf("not found") !== -1) {
          return createGoogleContact_(rowObject);
        }
        throw e;
      }
    }

    payload.etag = etag;

    var response = People.People.updateContact(payload, resourceName, {
      updatePersonFields: updateFields
    });

    return {
      success: true,
      resourceName: response.resourceName || resourceName,
      etag: response.etag || "",
      message: "Kontak berhasil diperbarui."
    };
  } catch (err) {
    var fullErrStr = String(err.message || err);
    // Jika kontak tidak ditemukan di Google, buat baru
    if (fullErrStr.indexOf("404") !== -1 || fullErrStr.toLowerCase().indexOf("not found") !== -1) {
      rowObject["Google_Resource_Name"] = "";
      rowObject["Google_ETag"] = "";
      return createGoogleContact_(rowObject);
    }

    // Jika Google menolak karena etag lokal kedaluwarsa, ambil etag terbaru lalu retry sekali.
    if (fullErrStr.toLowerCase().indexOf("etag") !== -1) {
      try {
        var latest = People.People.get(resourceName, { personFields: "metadata" });
        var latestEtag = latest && latest.etag ? latest.etag : "";
        if (latestEtag) {
          var retryPayload = buildPersonPayload_(rowObject);
          retryPayload.etag = latestEtag;
          var retryResponse = People.People.updateContact(retryPayload, resourceName, {
            updatePersonFields: "names,emailAddresses"
          });
          return {
            success: true,
            resourceName: retryResponse.resourceName || resourceName,
            etag: retryResponse.etag || latestEtag,
            message: "Kontak berhasil diperbarui setelah refresh etag."
          };
        }
      } catch (retryErr) {
        var retryErrStr = String(retryErr.message || retryErr);
        if (retryErrStr.indexOf("404") !== -1 || retryErrStr.toLowerCase().indexOf("not found") !== -1) {
          rowObject["Google_Resource_Name"] = "";
          rowObject["Google_ETag"] = "";
          return createGoogleContact_(rowObject);
        }
        fullErrStr = retryErrStr;
      }
    }

    var errStr = fullErrStr.substring(0, 200);
    return {
      success: false,
      resourceName: resourceName,
      etag: "",
      message: "Gagal update kontak: " + errStr
    };
  }
}

/**
 * Mencari kontak di Google Contacts berdasarkan email jabatan.
 * Mengembalikan objek person jika ditemukan, null jika tidak.
 */
function findContactByEmail_(emailJabatan) {
  emailJabatan = String(emailJabatan || "").trim().toLowerCase();
  if (!emailJabatan) return null;

  try {
    var response = People.People.searchContacts({
      query: emailJabatan,
      readMask: "names,emailAddresses,metadata"
    });

    var results = response.results || [];
    for (var i = 0; i < results.length; i++) {
      var person = results[i].person;
      if (!person) continue;
      var emails = person.emailAddresses || [];
      for (var j = 0; j < emails.length; j++) {
        if (String(emails[j].value || "").trim().toLowerCase() === emailJabatan) {
          return {
            resourceName: person.resourceName,
            etag: person.etag || ""
          };
        }
      }
    }
    return null;
  } catch (err) {
    console.warn("findContactByEmail_ error: " + err);
    return null;
  }
}

/**
 * Ambil daftar group/label Google Contacts untuk UI.
 */
function listGoogleContactGroups() {
  try {
    var groups = listGoogleContactGroups_();
    return { success: true, data: groups };
  } catch (err) {
    return {
      success: false,
      data: [],
      message: "Gagal mengambil grup Google Contacts: " + String(err.message || err).substring(0, 200)
    };
  }
}

/**
 * Ambil daftar group Google Contacts beserta anggota di dalamnya.
 */
function listGoogleContactGroupsWithMembers() {
  try {
    var groups = listGoogleContactGroups_();
    var groupMap = {};
    var output = [];

    for (var g = 0; g < groups.length; g++) {
      var item = {
        name: groups[g].name,
        resourceName: groups[g].resourceName,
        memberCount: groups[g].memberCount || 0,
        visibleMemberCount: 0,
        members: []
      };
      groupMap[item.resourceName] = item;
      output.push(item);
    }

    var pageToken = "";
    var safety = 0;
    do {
      safety++;
      if (safety > 100) break;

      var response = People.People.Connections.list("people/me", {
        personFields: "names,emailAddresses,memberships,metadata",
        pageSize: 500,
        pageToken: pageToken
      });

      var connections = response.connections || [];
      for (var i = 0; i < connections.length; i++) {
        var person = connections[i];
        var memberships = person && person.memberships ? person.memberships : [];
        var member = mapGooglePersonToMemberSummary_(person);

        for (var m = 0; m < memberships.length; m++) {
          var membership = memberships[m] && memberships[m].contactGroupMembership;
          if (!membership) continue;

          var groupResourceName = String(membership.contactGroupResourceName || "").trim();
          if (!groupMap[groupResourceName]) continue;

          groupMap[groupResourceName].members.push(member);
        }
      }

      pageToken = response.nextPageToken || "";
    } while (pageToken);

    for (var x = 0; x < output.length; x++) {
      output[x].members.sort(function(a, b) {
        var aa = String(a.name || a.email || "").toLowerCase();
        var bb = String(b.name || b.email || "").toLowerCase();
        return aa.localeCompare(bb, "id", { sensitivity: "base" });
      });
      output[x].visibleMemberCount = output[x].members.length;
    }

    output.sort(function(a, b) {
      return String(a.name || "").localeCompare(String(b.name || ""), "id", { sensitivity: "base" });
    });

    return {
      success: true,
      data: output,
      totalGroups: output.length,
      message: "Daftar grup dan anggota berhasil dimuat."
    };
  } catch (err) {
    return {
      success: false,
      data: [],
      message: "Gagal mengambil anggota grup Google Contacts: " + String(err.message || err).substring(0, 200)
    };
  }
}

/**
 * Ambil daftar group/label Google Contacts.
 */
function listGoogleContactGroups_() {
  var result = [];
  var pageToken = "";
  do {
    var listRes = People.ContactGroups.list({
      pageSize: 1000,
      pageToken: pageToken,
      groupFields: "name,groupType,memberCount"
    });
    var groups = listRes.contactGroups || [];
    for (var i = 0; i < groups.length; i++) {
      var g = groups[i] || {};
      var name = String(g.name || g.formattedName || "").trim();
      var resourceName = String(g.resourceName || "").trim();
      var groupType = String(g.groupType || "").trim();
      var resourceNameLower = resourceName.toLowerCase();
      if (!name || !resourceName) continue;
      if (groupType === "SYSTEM_CONTACT_GROUP") continue;
      if (resourceNameLower.indexOf("contactgroups/mycontacts") !== -1) continue;
      if (resourceNameLower.indexOf("contactgroups/starred") !== -1) continue;
      if (resourceNameLower.indexOf("contactgroups/friends") !== -1) continue;
      if (resourceNameLower.indexOf("contactgroups/family") !== -1) continue;
      if (resourceNameLower.indexOf("contactgroups/coworkers") !== -1) continue;
      result.push({
        name: name,
        resourceName: resourceName,
        memberCount: g.memberCount || 0
      });
    }
    pageToken = listRes.nextPageToken || "";
  } while (pageToken);

  result.sort(function(a, b) {
    return String(a.name || "").localeCompare(String(b.name || ""), "id", { sensitivity: "base" });
  });
  return result;
}

/**
 * Buat map resourceName -> nama group Google Contacts.
 */
function getContactGroupNameMap_() {
  var map = {};
  var groups = listGoogleContactGroups_();
  for (var i = 0; i < groups.length; i++) {
    map[groups[i].resourceName] = groups[i].name;
  }
  return map;
}

/**
 * Pecah isi kolom Contact_Group menjadi array nama group unik.
 */
function splitContactGroups_(value) {
  var raw = String(value || "").trim();
  if (!raw) return [];

  var seen = {};
  var out = [];
  raw.split(/[,\n;]+/).forEach(function(part) {
    var name = String(part || "").trim();
    var key = name.toLowerCase();
    if (!name || seen[key]) return;
    seen[key] = true;
    out.push(name);
  });
  return out;
}

/**
 * Gabungkan dua daftar group tanpa duplikat.
 */
function mergeContactGroupNames_(currentValue, incomingValue) {
  return splitContactGroups_(String(currentValue || "") + "," + String(incomingValue || "")).join(", ");
}

/**
 * Samakan membership Google Contact dengan isi kolom Contact_Group.
 * Grup user yang tidak ada lagi di sheet akan dilepas dari kontak.
 */
function syncContactGroups_(personResourceName, contactGroupValue) {
  personResourceName = String(personResourceName || "").trim();
  if (!personResourceName) {
    return { success: false, message: "resourceName kontak tidak valid." };
  }

  var desiredNames = splitContactGroups_(contactGroupValue);
  var desiredResourceMap = {};
  var desiredResourceNames = [];
  var desiredNameByResource = {};

  for (var i = 0; i < desiredNames.length; i++) {
    var groupName = desiredNames[i];
    var grp = ensureContactGroupExists_(groupName);
    if (!grp.success) {
      return { success: false, message: "Gagal membuat/mencari grup '" + groupName + "': " + grp.message };
    }
    if (grp.resourceName) {
      desiredResourceMap[grp.resourceName] = true;
      desiredResourceNames.push(grp.resourceName);
      desiredNameByResource[grp.resourceName] = groupName;
    }
  }

  var userGroupNameMap = getContactGroupNameMap_();
  for (var d = 0; d < desiredResourceNames.length; d++) {
    if (!userGroupNameMap[desiredResourceNames[d]]) {
      userGroupNameMap[desiredResourceNames[d]] = desiredNameByResource[desiredResourceNames[d]] || desiredResourceNames[d];
    }
  }

  var currentResourceMap = getPersonContactGroupResourceMap_(personResourceName);
  var added = 0;
  var removed = 0;

  for (var addIdx = 0; addIdx < desiredResourceNames.length; addIdx++) {
    var desiredResourceName = desiredResourceNames[addIdx];
    if (currentResourceMap[desiredResourceName]) continue;

    var addMember = addContactToGroup_(personResourceName, desiredResourceName);
    if (!addMember.success) {
      return { success: false, message: "Gagal menambahkan kontak ke grup '" + (userGroupNameMap[desiredResourceName] || desiredResourceName) + "': " + addMember.message };
    }
    added++;
  }

  for (var currentResourceName in currentResourceMap) {
    if (!userGroupNameMap[currentResourceName] || desiredResourceMap[currentResourceName]) continue;

    var removeMember = removeContactFromGroup_(personResourceName, currentResourceName);
    if (!removeMember.success) {
      return { success: false, message: "Gagal menghapus kontak dari grup '" + (userGroupNameMap[currentResourceName] || currentResourceName) + "': " + removeMember.message };
    }
    removed++;
  }

  var message = desiredNames.length > 0
    ? "OK (grup: " + desiredNames.join(", ") + ")"
    : "OK (tanpa grup)";

  if (added > 0 || removed > 0) {
    message += " tambah: " + added + ", hapus: " + removed + ".";
  }

  return { success: true, message: message, added: added, removed: removed };
}

/**
 * Ambil membership group untuk satu Google Contact.
 */
function getPersonContactGroupResourceMap_(personResourceName) {
  var map = {};
  var person = People.People.get(personResourceName, { personFields: "memberships" });
  var memberships = person && person.memberships ? person.memberships : [];

  for (var i = 0; i < memberships.length; i++) {
    var membership = memberships[i] && memberships[i].contactGroupMembership;
    if (!membership) continue;
    var groupResourceName = String(membership.contactGroupResourceName || "").trim();
    if (groupResourceName) map[groupResourceName] = true;
  }

  return map;
}

/**
 * Buat group baru jika belum ada (match by name, case-insensitive).
 */
function ensureContactGroupExists_(groupName) {
  groupName = String(groupName || "").trim();
  if (!groupName) return { success: false, message: "Nama grup kosong.", resourceName: "" };

  try {
    var pageToken = "";
    do {
      var listRes = People.ContactGroups.list({ pageSize: 1000, pageToken: pageToken });
      var groups = listRes.contactGroups || [];
      for (var i = 0; i < groups.length; i++) {
        var gName = String(groups[i].name || "").trim();
        if (gName && gName.toLowerCase() === groupName.toLowerCase()) {
          return { success: true, message: "Group ditemukan.", resourceName: String(groups[i].resourceName || "").trim() };
        }
      }
      pageToken = listRes.nextPageToken || "";
    } while (pageToken);

    var created = People.ContactGroups.create({ contactGroup: { name: groupName } });
    return {
      success: true,
      message: "Group dibuat.",
      resourceName: String(created && created.resourceName || "").trim()
    };
  } catch (err) {
    return { success: false, message: String(err.message || err).substring(0, 200), resourceName: "" };
  }
}

/**
 * Tambahkan kontak ke group.
 */
function addContactToGroup_(personResourceName, groupResourceName) {
  personResourceName = String(personResourceName || "").trim();
  groupResourceName = String(groupResourceName || "").trim();
  if (!personResourceName || !groupResourceName) {
    return { success: false, message: "resourceName kontak/grup tidak valid." };
  }

  try {
    People.ContactGroups.Members.modify({ resourceNamesToAdd: [personResourceName] }, groupResourceName);
    return { success: true, message: "Kontak ditambahkan ke group." };
  } catch (err) {
    var em = String(err.message || err);
    if (em.toLowerCase().indexOf("already") !== -1 || em.toLowerCase().indexOf("exist") !== -1) {
      return { success: true, message: "Kontak sudah ada di group." };
    }
    return { success: false, message: em.substring(0, 200) };
  }
}

/**
 * Hapus kontak dari group.
 */
function removeContactFromGroup_(personResourceName, groupResourceName) {
  personResourceName = String(personResourceName || "").trim();
  groupResourceName = String(groupResourceName || "").trim();
  if (!personResourceName || !groupResourceName) {
    return { success: false, message: "resourceName kontak/grup tidak valid." };
  }

  try {
    People.ContactGroups.Members.modify({ resourceNamesToRemove: [personResourceName] }, groupResourceName);
    return { success: true, message: "Kontak dihapus dari group." };
  } catch (err) {
    var em = String(err.message || err);
    if (em.toLowerCase().indexOf("not a member") !== -1) {
      return { success: true, message: "Kontak sudah tidak ada di group." };
    }
    return { success: false, message: em.substring(0, 200) };
  }
}

// -----------------------------------------------------------------------
// IMPORT & DEDUP GOOGLE CONTACTS
// -----------------------------------------------------------------------

/**
 * Import Google Contacts ke sheet Config_Contacts.
 * Tidak menghapus data existing; hanya tambah baru atau update baris yang emailnya sama.
 */
function importContactsFromGoogleToSheet() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(CONTACTS_SHEET_NAME);
  if (!sheet) {
    setupContactsSheet();
    sheet = ss.getSheetByName(CONTACTS_SHEET_NAME);
  }

  normalizeContactsSheet_(sheet);
  var headerMap = getHeaderMap_(sheet);
  var numCols = CONTACTS_HEADERS.length;

  // Index existing by email agar import idempotent.
  var emailIndex = {};
  var lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    var values = sheet.getRange(2, 1, lastRow - 1, numCols).getValues();
    for (var i = 0; i < values.length; i++) {
      var existingEmail = normalizeEmail_(getRowContactEmail_(values[i], headerMap));
      if (existingEmail) {
        emailIndex[existingEmail] = {
          rowIndex: i + 2,
          rowValues: values[i]
        };
      }
    }
  }

  var imported = 0;
  var updated = 0;
  var skipped = 0;
  var pageToken = "";
  var safety = 0;
  var groupNameMap = getContactGroupNameMap_();

  do {
    safety++;
    if (safety > 100) break;

    var response = People.People.Connections.list("people/me", {
      personFields: "names,emailAddresses,memberships,metadata",
      pageSize: 500,
      pageToken: pageToken
    });

    var connections = response.connections || [];
    for (var c = 0; c < connections.length; c++) {
      var person = connections[c];
      var mapped = mapGooglePersonToSheetRow_(person, groupNameMap);
      if (!getContactEmail_(mapped) || !mapped["Nama_Display"]) {
        skipped++;
        continue;
      }

      var key = normalizeEmail_(getContactEmail_(mapped));
      if (!key) {
        skipped++;
        continue;
      }

      if (emailIndex[key]) {
        var target = emailIndex[key];
        var rowObj = {};
        for (var colName in headerMap) {
          rowObj[colName] = String(target.rowValues[headerMap[colName]] || "").trim();
        }

        rowObj["Contact_ID"] = rowObj["Contact_ID"] || createRecordId("CTK");
        rowObj["Nama_Display"] = mapped["Nama_Display"];
        rowObj["Email"] = mapped["Email"];
        rowObj["Contact_Group"] = mapped["Contact_Group"];
        rowObj["Status"] = "Aktif";
        rowObj["Sync_Mode"] = rowObj["Sync_Mode"] || "AUTO";
        rowObj["Google_Resource_Name"] = mapped["Google_Resource_Name"];
        rowObj["Google_ETag"] = mapped["Google_ETag"];
        rowObj["Last_Sync"] = mapped["Last_Sync"];

        if (String(rowObj["Sync_Mode"] || "AUTO").trim().toUpperCase() === "NO_SYNC") {
          rowObj["Sync_Status"] = "";
          rowObj["Sync_Message"] = "Diimpor dari Google Contacts. Mode NO_SYNC.";
        } else {
          rowObj["Sync_Status"] = "SYNCED";
          rowObj["Sync_Message"] = "Diimpor dari Google Contacts.";
        }

        var updatedRow = buildRowFromObject_(rowObj, headerMap, numCols);
        sheet.getRange(target.rowIndex, 1, 1, numCols).setValues([updatedRow]);
        updated++;
        emailIndex[key].rowValues = updatedRow;
      } else {
        mapped["Contact_ID"] = createRecordId("CTK");
        mapped["Status"] = mapped["Status"] || "Aktif";
        mapped["Sync_Mode"] = mapped["Sync_Mode"] || "AUTO";
        mapped["Sync_Status"] = "SYNCED";
        mapped["Sync_Message"] = "Diimpor dari Google Contacts.";
        mapped["Last_Sync"] = Utilities.formatDate(new Date(), "Asia/Jakarta", "yyyy-MM-dd HH:mm:ss");

        var newRow = buildRowFromObject_(mapped, headerMap, numCols);
        sheet.appendRow(newRow);
        imported++;

        // Update index in-memory for mencegah duplikat pada batch import yang sama.
        var newRowIndex = sheet.getLastRow();
        emailIndex[key] = { rowIndex: newRowIndex, rowValues: newRow };
      }
    }

    pageToken = response.nextPageToken || "";
  } while (pageToken);

  SpreadsheetApp.flush();
  var msg = "Import selesai. Baru: " + imported + ", update: " + updated + ", dilewati: " + skipped + ".";
  logAudit("Import Kontak", "-", "IMPORT", "", msg, Session.getActiveUser().getEmail());
  return { success: true, imported: imported, updated: updated, skipped: skipped, message: msg };
}

/**
 * Rapikan duplikat pada sheet berdasarkan email.
 * Menyisakan 1 baris terbaik per email dan menggabungkan field yang kosong.
 */
function deduplicateContactsInSheet() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(CONTACTS_SHEET_NAME);
  if (!sheet) return { success: false, message: "Sheet Config_Contacts tidak ditemukan." };

  normalizeContactsSheet_(sheet);
  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) {
    return { success: true, message: "Tidak ada data untuk dideduplikasi.", groups: 0, deleted: 0, updated: 0 };
  }

  var headerMap = getHeaderMap_(sheet);
  var numCols = CONTACTS_HEADERS.length;
  var values = sheet.getRange(2, 1, lastRow - 1, numCols).getValues();

  var groups = {};
  for (var i = 0; i < values.length; i++) {
    var row = values[i];
    var email = normalizeEmail_(getRowContactEmail_(row, headerMap));
    if (!email) continue;
    if (!groups[email]) groups[email] = [];
    groups[email].push({ rowIndex: i + 2, rowValues: row });
  }

  var rowsToDelete = [];
  var updatedRows = 0;
  var dupGroups = 0;

  for (var emailKey in groups) {
    var bucket = groups[emailKey];
    if (!bucket || bucket.length <= 1) continue;
    dupGroups++;

    bucket.sort(function(a, b) {
      var sa = scoreContactRow_(a.rowValues, headerMap);
      var sb = scoreContactRow_(b.rowValues, headerMap);
      if (sb !== sa) return sb - sa;
      return a.rowIndex - b.rowIndex;
    });

    var keeper = bucket[0];
    var mergedObj = rowArrayToObject_(keeper.rowValues, headerMap);

    for (var j = 1; j < bucket.length; j++) {
      var dupObj = rowArrayToObject_(bucket[j].rowValues, headerMap);
      mergedObj = mergeContactObjects_(mergedObj, dupObj);
      rowsToDelete.push(bucket[j].rowIndex);
    }

    if (!mergedObj["Contact_ID"]) mergedObj["Contact_ID"] = createRecordId("CTK");
    if (!mergedObj["Status"]) mergedObj["Status"] = "Aktif";
    if (!mergedObj["Sync_Mode"]) mergedObj["Sync_Mode"] = "AUTO";
    mergedObj["Sync_Status"] = "NEED_SYNC";
    mergedObj["Sync_Message"] = "Duplikat digabung. Perlu sync ulang ke Google Contacts.";

    var mergedRow = buildRowFromObject_(mergedObj, headerMap, numCols);
    sheet.getRange(keeper.rowIndex, 1, 1, numCols).setValues([mergedRow]);
    updatedRows++;
  }

  rowsToDelete.sort(function(a, b) { return b - a; });
  for (var d = 0; d < rowsToDelete.length; d++) {
    sheet.deleteRow(rowsToDelete[d]);
  }

  SpreadsheetApp.flush();
  var message = "Deduplikasi selesai. Grup duplikat: " + dupGroups + ", baris dihapus: " + rowsToDelete.length + ", baris digabung: " + updatedRows + ".";
  logAudit("Deduplikasi Kontak", "-", "DEDUP", "", message, Session.getActiveUser().getEmail());

  return {
    success: true,
    groups: dupGroups,
    deleted: rowsToDelete.length,
    updated: updatedRows,
    message: message
  };
}

// -----------------------------------------------------------------------
// TRIGGER HARIAN
// -----------------------------------------------------------------------

/**
 * Install trigger harian untuk sync kontak otomatis setiap hari pukul 02:00.
 */
function installContactsDailySyncTrigger() {
  // Hapus trigger lama dulu
  deleteContactsSyncTriggers();

  ScriptApp.newTrigger("syncContactsToGoogleContacts")
    .timeBased()
    .everyDays(1)
    .atHour(2)
    .create();

  return { success: true, message: "Trigger sync kontak harian (pukul 02:00) berhasil dipasang." };
}

/**
 * Hapus semua trigger yang memanggil syncContactsToGoogleContacts.
 */
function deleteContactsSyncTriggers() {
  deleteContactTriggersByHandler_("syncContactsToGoogleContacts");
  return { success: true, message: "Trigger sync kontak berhasil dihapus." };
}

/**
 * Hapus trigger kontak berdasarkan nama handler function.
 * Dibuat private agar tidak bentrok dengan helper global di Code.gs.
 * @param {string} handlerName
 */
function deleteContactTriggersByHandler_(handlerName) {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === handlerName) {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
}

// -----------------------------------------------------------------------
// HELPER INTERNAL
// -----------------------------------------------------------------------

/**
 * Mengambil peta header → index kolom (0-based) dari sheet.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @returns {Object} { "Header_Name": colIndex, ... }
 */
function getHeaderMap_(sheet) {
  var headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var map = {};
  for (var i = 0; i < headerRow.length; i++) {
    var h = String(headerRow[i] || "").trim();
    if (h) map[h] = i;
  }
  return map;
}

/**
 * Validasi baris kontak - Email wajib ada.
 */
function validateContactRow_(rowObject) {
  var email = getContactEmail_(rowObject);
  if (!email) {
    return { valid: false, message: "Email wajib diisi untuk sync ke Google Contacts." };
  }
  // Validasi format email sederhana
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) {
    return { valid: false, message: "Format Email tidak valid: " + email };
  }
  var namaDisplay = String(rowObject["Nama_Display"] || "").trim();
  if (!namaDisplay) {
    return { valid: false, message: "Nama_Display wajib diisi." };
  }
  return { valid: true, message: "" };
}

/**
 * Membangun payload Person untuk People API.
 */
function buildPersonPayload_(rowObject) {
  var namaDisplay = String(rowObject["Nama_Display"] || "").trim();
  var email = getContactEmail_(rowObject);

  var payload = {
    names: [
      {
        displayName: namaDisplay,
        givenName: namaDisplay
      }
    ],
    emailAddresses: []
  };

  // Satu email utama kontak.
  if (email) {
    payload.emailAddresses.push({
      value: email,
      type: "work",
      formattedType: "Email"
    });
  }

  return payload;
}

/**
 * Update kolom sync status di sheet tanpa mengubah kolom lain.
 */
function updateSyncStatus_(sheet, rowIndex, headerMap, status, message, resourceName, etag, lastSync) {
  var updates = {};
  if (headerMap["Sync_Status"] !== undefined) updates[headerMap["Sync_Status"]] = status || "";
  if (headerMap["Sync_Message"] !== undefined) updates[headerMap["Sync_Message"]] = String(message || "").substring(0, 500);
  if (resourceName !== undefined && headerMap["Google_Resource_Name"] !== undefined) {
    updates[headerMap["Google_Resource_Name"]] = resourceName || "";
  }
  if (etag !== undefined && headerMap["Google_ETag"] !== undefined) {
    updates[headerMap["Google_ETag"]] = etag || "";
  }
  if (lastSync !== undefined && headerMap["Last_Sync"] !== undefined) {
    updates[headerMap["Last_Sync"]] = lastSync || "";
  }

  for (var colIdx in updates) {
    sheet.getRange(rowIndex, parseInt(colIdx) + 1).setValue(updates[colIdx]);
  }
}

/**
 * Membangun array nilai row dari object berdasarkan headerMap.
 */
function buildRowFromObject_(obj, headerMap, numCols) {
  var row = new Array(numCols).fill("");
  for (var col in headerMap) {
    var idx = headerMap[col];
    if (idx < numCols) {
      var val = obj[col];
      if (val === undefined && col === "Email") {
        val = getContactEmail_(obj);
      }

      if (val !== undefined) {
        row[idx] = String(val !== null ? val : "");
      }
    }
  }
  return row;
}

/**
 * Ambil satu email kontak. Mendukung kolom baru Email dan kolom lama.
 */
function getContactEmail_(rowObject) {
  rowObject = rowObject || {};
  return String(rowObject["Email"] || rowObject["Email_Jabatan"] || rowObject["Email_Pribadi"] || "").trim();
}

/**
 * Ambil satu email kontak dari array row sheet.
 */
function getRowContactEmail_(rowValues, headerMap) {
  if (!rowValues || !headerMap) return "";
  if (headerMap["Email"] !== undefined) return String(rowValues[headerMap["Email"]] || "").trim();
  if (headerMap["Email_Jabatan"] !== undefined && rowValues[headerMap["Email_Jabatan"]]) {
    return String(rowValues[headerMap["Email_Jabatan"]] || "").trim();
  }
  if (headerMap["Email_Pribadi"] !== undefined) return String(rowValues[headerMap["Email_Pribadi"]] || "").trim();
  return "";
}

/**
 * Normalisasi email untuk key lookup.
 */
function normalizeEmail_(val) {
  return String(val || "").trim().toLowerCase();
}

/**
 * Mapping person Google Contacts -> object row sheet.
 */
function mapGooglePersonToSheetRow_(person, groupNameMap) {
  var names = person && person.names ? person.names : [];
  var emails = person && person.emailAddresses ? person.emailAddresses : [];
  var memberships = person && person.memberships ? person.memberships : [];

  var displayName = "";
  if (names.length > 0) {
    displayName = String(names[0].displayName || names[0].givenName || "").trim();
  }

  var email = "";
  for (var i = 0; i < emails.length; i++) {
    var e = String(emails[i].value || "").trim();
    if (!e) continue;
    email = e;
    break;
  }

  var contactGroups = [];
  var seenGroups = {};
  groupNameMap = groupNameMap || {};
  for (var m = 0; m < memberships.length; m++) {
    var membership = memberships[m] && memberships[m].contactGroupMembership;
    if (!membership) continue;
    var groupResourceName = String(membership.contactGroupResourceName || "").trim();
    var groupName = String(groupNameMap[groupResourceName] || "").trim();
    var key = groupName.toLowerCase();
    if (!groupName || seenGroups[key]) continue;
    seenGroups[key] = true;
    contactGroups.push(groupName);
  }

  return {
    Nama_Display: displayName,
    Email: email,
    Contact_Group: contactGroups.join(", "),
    Status: "Aktif",
    Sync_Mode: "AUTO",
    Google_Resource_Name: String(person && person.resourceName || "").trim(),
    Google_ETag: String(person && person.etag || "").trim(),
    Last_Sync: Utilities.formatDate(new Date(), "Asia/Jakarta", "yyyy-MM-dd HH:mm:ss")
  };
}

/**
 * Ringkasan person Google Contacts untuk tampilan anggota group.
 */
function mapGooglePersonToMemberSummary_(person) {
  var names = person && person.names ? person.names : [];
  var emails = person && person.emailAddresses ? person.emailAddresses : [];

  var displayName = "";
  if (names.length > 0) {
    displayName = String(names[0].displayName || names[0].givenName || "").trim();
  }

  var email = "";
  for (var i = 0; i < emails.length; i++) {
    var e = String(emails[i].value || "").trim();
    if (!e) continue;
    email = e;
    break;
  }

  return {
    resourceName: String(person && person.resourceName || "").trim(),
    name: displayName || email || "(Tanpa nama)",
    email: email
  };
}

/**
 * Ubah array row menjadi object berdasarkan header map.
 */
function rowArrayToObject_(rowArray, headerMap) {
  var obj = {};
  for (var col in headerMap) {
    obj[col] = String(rowArray[headerMap[col]] !== undefined ? rowArray[headerMap[col]] : "").trim();
  }
  return obj;
}

/**
 * Skor kualitas row untuk memilih keeper saat deduplikasi.
 */
function scoreContactRow_(rowValues, headerMap) {
  var score = 0;
  var important = ["Nama_Display", "Google_Resource_Name", "Contact_Group"];
  for (var i = 0; i < important.length; i++) {
    var idx = headerMap[important[i]];
    if (idx !== undefined && String(rowValues[idx] || "").trim()) score += 2;
  }
  if (getRowContactEmail_(rowValues, headerMap)) score += 2;
  var syncStatusIdx = headerMap["Sync_Status"];
  if (syncStatusIdx !== undefined) {
    var sync = String(rowValues[syncStatusIdx] || "").trim().toUpperCase();
    if (sync === "SYNCED") score += 2;
    if (sync === "ERROR") score -= 1;
  }
  return score;
}

/**
 * Merge dua object kontak dengan preferensi nilai non-kosong.
 */
function mergeContactObjects_(baseObj, incomingObj) {
  var out = {};
  for (var key in baseObj) out[key] = baseObj[key];

  var preferIncomingIfLonger = ["Nama_Display", "Sync_Message"];
  var i;

  for (i = 0; i < preferIncomingIfLonger.length; i++) {
    var k = preferIncomingIfLonger[i];
    var b = String(out[k] || "").trim();
    var inc = String(incomingObj[k] || "").trim();
    if (!b && inc) {
      out[k] = inc;
    } else if (inc && inc.length > b.length) {
      out[k] = inc;
    }
  }

  var mergedEmail = getContactEmail_(out) || getContactEmail_(incomingObj);
  if (mergedEmail) {
    out["Email"] = mergedEmail;
  }

  var mergedGroups = mergeContactGroupNames_(out["Contact_Group"], incomingObj["Contact_Group"]);
  if (mergedGroups) {
    out["Contact_Group"] = mergedGroups;
  }

  if (!out["Google_Resource_Name"] && incomingObj["Google_Resource_Name"]) {
    out["Google_Resource_Name"] = String(incomingObj["Google_Resource_Name"] || "").trim();
  }
  if (!out["Google_ETag"] && incomingObj["Google_ETag"]) {
    out["Google_ETag"] = String(incomingObj["Google_ETag"] || "").trim();
  }

  return out;
}
