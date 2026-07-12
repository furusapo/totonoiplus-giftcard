function doGet(e) {
  const mode = e.parameter.mode;
  const action = e.parameter.action;
  const callback = e.parameter.callback;

  if (mode === "check") return checkGiftCard(e);
  if (mode === "use") return useGiftCard(e);

  // スタッフ管理画面用
  if (mode === "staff_search") return staffSearch(e);
  if (mode === "staff_use") return staffUse(e);

  if (action === "stock") return getStock(e);
  if (e.parameter.data) return issueGiftCard(e);

  return respond(callback, {
    ok: true,
    result: "success",
    message: "ギフトカードAPIが実行中です"
  });
}

function issueGiftCard(e) {
  const callback = e.parameter.callback;

  try {
    const data = JSON.parse(e.parameter.data || "{}");
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("発行履歴");
    if (!sheet) throw new Error("発行履歴シートが見つかりません");

    ensureStaffColumns(sheet);

    const requests = data.requests || [];
    const now = new Date();

    requests.forEach(item => {
      const design = Number(item.design);
      const count = Number(item.count || 0);

      for (let i = 0; i < count; i++) {
        const serial = getNextSerial(sheet, design);
        sheet.appendRow([
          now,
          serial,
          "デザイン" + design,
          data.issueDate || "",
          data.expireDate || "",
          data.phone || "",
          data.staff || "お客様確認済み",
          "未使用",
          "",
          "",
          ""
        ]);
      }
    });

    return respond(callback, {
      ok: true,
      result: "success",
      message: "発行履歴に保存しました"
    });

  } catch (err) {
    return respond(callback, {
      ok: false,
      result: "error",
      message: err.message
    });
  }
}

function getNextSerial(sheet, design) {
  const values = sheet.getDataRange().getValues();
  const offset = { 1: 0, 2: 50, 3: 100 }[design] || 0;
  const min = offset + 1;
  const max = offset + 50;
  const used = new Set();

  for (let i = 1; i < values.length; i++) {
    const serial = String(values[i][1] || "").trim();
    const m = serial.match(/^TG-(\d{4})$/);
    if (m) used.add(Number(m[1]));
  }

  for (let n = min; n <= max; n++) {
    if (!used.has(n)) return "TG-" + String(n).padStart(4, "0");
  }

  throw new Error("デザイン" + design + "の在庫がありません");
}

function getStock(e) {
  const callback = e.parameter.callback;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("発行履歴");

  if (!sheet) {
    return respond(callback, {
      ok: false,
      result: "error",
      message: "発行履歴シートが見つかりません"
    });
  }

  const values = sheet.getDataRange().getValues();
  const used = { 1: 0, 2: 0, 3: 0 };

  for (let i = 1; i < values.length; i++) {
    const serial = String(values[i][1] || "").trim();
    const m = serial.match(/^TG-(\d{4})$/);
    if (!m) continue;

    const num = Number(m[1]);
    if (num >= 1 && num <= 50) used[1]++;
    if (num >= 51 && num <= 100) used[2]++;
    if (num >= 101 && num <= 150) used[3]++;
  }

  return respond(callback, {
    ok: true,
    result: "success",
    stock: {
      1: 50 - used[1],
      2: 50 - used[2],
      3: 50 - used[3]
    }
  });
}

function checkGiftCard(e) {
  const callback = e.parameter.callback;
  const serial = String(e.parameter.serial || "").trim().toUpperCase();

  if (!serial) {
    return respond(callback, {
      found: false,
      message: "シリアル番号が入力されていません"
    });
  }

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("発行履歴");

  if (!sheet) {
    return respond(callback, {
      found: false,
      message: "発行履歴シートが見つかりません"
    });
  }

  const values = sheet.getDataRange().getValues();
  const headers = values[0].map(h => String(h).trim());

  const serialCol = findHeader(headers, ["シリアル番号", "シリアル", "発行番号", "ギフトカード番号"]);
  const designCol = findHeader(headers, ["デザイン", "デザイン名"]);
  const expiryCol = findHeader(headers, ["有効期限", "期限", "利用期限"]);
  const statusCol = findHeader(headers, ["状態", "ステータス", "使用状態"]);

  if (serialCol === -1 || expiryCol === -1 || statusCol === -1) {
    return respond(callback, {
      found: false,
      message: "必要な列が見つかりません"
    });
  }

  for (let i = 1; i < values.length; i++) {
    const rowSerial = String(values[i][serialCol] || "").trim().toUpperCase();

    if (rowSerial === serial) {
      return respond(callback, {
        found: true,
        serial: rowSerial,
        design: designCol !== -1 ? values[i][designCol] : "",
        expiry: formatDate(values[i][expiryCol]),
        status: normalizeStatus(values[i][statusCol])
      });
    }
  }

  return respond(callback, {
    found: false,
    message: "該当するギフトカードが見つかりません"
  });
}

function useGiftCard(e) {
  const callback = e.parameter.callback;
  const serial = String(e.parameter.serial || "").trim().toUpperCase();

  if (!serial) {
    return respond(callback, {
      ok: false,
      message: "シリアル番号が入力されていません"
    });
  }

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("発行履歴");

  if (!sheet) {
    return respond(callback, {
      ok: false,
      message: "発行履歴シートが見つかりません"
    });
  }

  ensureStaffColumns(sheet);

  const values = sheet.getDataRange().getValues();
  const headers = values[0].map(h => String(h).trim());

  const serialCol = findHeader(headers, ["シリアル番号", "シリアル", "発行番号", "ギフトカード番号"]);
  const statusCol = findHeader(headers, ["状態", "ステータス", "使用状態"]);
  const usedAtCol = findHeader(headers, ["使用日時", "利用日時"]);
  const usedStaffCol = findHeader(headers, ["使用スタッフ", "利用スタッフ", "対応スタッフ"]);
  const memoCol = findHeader(headers, ["使用メモ", "利用メモ", "メモ"]);

  if (serialCol === -1 || statusCol === -1) {
    return respond(callback, {
      ok: false,
      message: "シリアル番号または状態の列が見つかりません"
    });
  }

  for (let i = 1; i < values.length; i++) {
    const rowSerial = String(values[i][serialCol] || "").trim().toUpperCase();

    if (rowSerial === serial) {
      const currentStatus = normalizeStatus(values[i][statusCol]);

      if (currentStatus !== "未使用") {
        return respond(callback, {
          ok: false,
          message: "このギフトカードはすでに「" + currentStatus + "」です"
        });
      }

      sheet.getRange(i + 1, statusCol + 1).setValue("利用済");
      if (usedAtCol !== -1) sheet.getRange(i + 1, usedAtCol + 1).setValue(new Date());
      if (usedStaffCol !== -1) sheet.getRange(i + 1, usedStaffCol + 1).setValue("スタッフ画面");
      if (memoCol !== -1) sheet.getRange(i + 1, memoCol + 1).setValue("");

      return respond(callback, {
        ok: true,
        message: serial + " を利用済みにしました"
      });
    }
  }

  return respond(callback, {
    ok: false,
    message: "該当するギフトカードが見つかりません"
  });
}

function staffSearch(e) {
  const callback = e.parameter.callback;
  const serial = String(e.parameter.serial || "").trim().toUpperCase();

  if (!serial) {
    return respond(callback, {
      ok: false,
      message: "シリアル番号が入力されていません"
    });
  }

  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("発行履歴");
    if (!sheet) throw new Error("発行履歴シートが見つかりません");

    ensureStaffColumns(sheet);

    const values = sheet.getDataRange().getValues();
    const headers = values[0].map(h => String(h).trim());

    const serialCol = findHeader(headers, ["シリアル番号", "シリアル", "発行番号", "ギフトカード番号"]);
    const designCol = findHeader(headers, ["デザイン", "デザイン名"]);
    const issueDateCol = findHeader(headers, ["発行日", "購入日"]);
    const expiryCol = findHeader(headers, ["有効期限", "期限", "利用期限"]);
    const phoneCol = findHeader(headers, ["電話番号", "電話"]);
    const staffCol = findHeader(headers, ["発行担当者", "担当者", "スタッフ"]);
    const statusCol = findHeader(headers, ["状態", "ステータス", "使用状態"]);
    const usedAtCol = findHeader(headers, ["使用日時", "利用日時"]);
    const usedStaffCol = findHeader(headers, ["使用スタッフ", "利用スタッフ", "対応スタッフ"]);
    const memoCol = findHeader(headers, ["使用メモ", "利用メモ", "メモ"]);

    if (serialCol === -1 || statusCol === -1) {
      throw new Error("シリアル番号または状態の列が見つかりません");
    }

    for (let i = 1; i < values.length; i++) {
      const rowSerial = String(values[i][serialCol] || "").trim().toUpperCase();

      if (rowSerial === serial) {
        return respond(callback, {
          ok: true,
          card: {
            serial: rowSerial,
            amount: "5,000円",
            design: designCol !== -1 ? values[i][designCol] : "",
            issueDate: issueDateCol !== -1 ? formatDate(values[i][issueDateCol]) : "",
            expiryDate: expiryCol !== -1 ? formatDate(values[i][expiryCol]) : "",
            phone: phoneCol !== -1 ? values[i][phoneCol] : "",
            issueStaff: staffCol !== -1 ? values[i][staffCol] : "",
            status: normalizeStatus(values[i][statusCol]),
            usedAt: usedAtCol !== -1 ? formatDateTime(values[i][usedAtCol]) : "",
            usedStaff: usedStaffCol !== -1 ? values[i][usedStaffCol] : "",
            memo: memoCol !== -1 ? values[i][memoCol] : ""
          }
        });
      }
    }

    return respond(callback, {
      ok: false,
      message: "該当するギフトカードが見つかりません"
    });

  } catch (err) {
    return respond(callback, {
      ok: false,
      message: err.message
    });
  }
}

function staffUse(e) {
  const callback = e.parameter.callback;
  const serial = String(e.parameter.serial || "").trim().toUpperCase();
  const staff = String(e.parameter.staff || "").trim();
  const memo = String(e.parameter.memo || "").trim();

  if (!serial) {
    return respond(callback, {
      ok: false,
      message: "シリアル番号が入力されていません"
    });
  }

  if (!staff) {
    return respond(callback, {
      ok: false,
      message: "対応スタッフが選択されていません"
    });
  }

  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("発行履歴");
    if (!sheet) throw new Error("発行履歴シートが見つかりません");

    ensureStaffColumns(sheet);

    const values = sheet.getDataRange().getValues();
    const headers = values[0].map(h => String(h).trim());

    const serialCol = findHeader(headers, ["シリアル番号", "シリアル", "発行番号", "ギフトカード番号"]);
    const statusCol = findHeader(headers, ["状態", "ステータス", "使用状態"]);
    const usedAtCol = findHeader(headers, ["使用日時", "利用日時"]);
    const usedStaffCol = findHeader(headers, ["使用スタッフ", "利用スタッフ", "対応スタッフ"]);
    const memoCol = findHeader(headers, ["使用メモ", "利用メモ", "メモ"]);

    if (serialCol === -1 || statusCol === -1) {
      throw new Error("シリアル番号または状態の列が見つかりません");
    }

    for (let i = 1; i < values.length; i++) {
      const rowSerial = String(values[i][serialCol] || "").trim().toUpperCase();

      if (rowSerial === serial) {
        const currentStatus = normalizeStatus(values[i][statusCol]);

        if (currentStatus !== "未使用") {
          return respond(callback, {
            ok: false,
            message: "このギフトカードはすでに「" + currentStatus + "」です"
          });
        }

        sheet.getRange(i + 1, statusCol + 1).setValue("利用済");
        sheet.getRange(i + 1, usedAtCol + 1).setValue(new Date());
        sheet.getRange(i + 1, usedStaffCol + 1).setValue(staff);
        sheet.getRange(i + 1, memoCol + 1).setValue(memo);

        return respond(callback, {
          ok: true,
          message: serial + " を利用済みにしました"
        });
      }
    }

    return respond(callback, {
      ok: false,
      message: "該当するギフトカードが見つかりません"
    });

  } catch (err) {
    return respond(callback, {
      ok: false,
      message: err.message
    });
  }
}

function ensureStaffColumns(sheet) {
  const required = ["使用日時", "使用スタッフ", "使用メモ"];
  const lastCol = sheet.getLastColumn();
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h).trim());

  required.forEach(name => {
    if (headers.indexOf(name) === -1) {
      sheet.getRange(1, sheet.getLastColumn() + 1).setValue(name);
    }
  });
}

function normalizeStatus(value) {
  const status = String(value || "未使用").trim();
  if (status === "利用済") return "利用済";
  if (status === "利用済み") return "利用済";
  if (!status) return "未使用";
  return status;
}

function respond(callback, obj) {
  if (callback) return jsonpOutput(callback, obj);
  return jsonOutput(obj);
}

function jsonOutput(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function jsonpOutput(callback, obj) {
  return ContentService
    .createTextOutput(callback + "(" + JSON.stringify(obj) + ");")
    .setMimeType(ContentService.MimeType.JAVASCRIPT);
}

function findHeader(headers, candidates) {
  for (const name of candidates) {
    const index = headers.indexOf(name);
    if (index !== -1) return index;
  }
  return -1;
}

function formatDate(value) {
  if (!value) return "";
  if (Object.prototype.toString.call(value) === "[object Date]") {
    return Utilities.formatDate(value, "Asia/Tokyo", "yyyy/MM/dd");
  }
  return String(value);
}

function formatDateTime(value) {
  if (!value) return "";
  if (Object.prototype.toString.call(value) === "[object Date]") {
    return Utilities.formatDate(value, "Asia/Tokyo", "yyyy/MM/dd HH:mm");
  }
  return String(value);
}

function doPost(e) {

  // ギフトカードPDFをGoogle Driveへ保存
  var giftcardAction = "";

  try {
    if (e && e.parameter && e.parameter.action) {
      giftcardAction = e.parameter.action;
    } else if (e && e.postData && e.postData.contents) {
      var giftcardRequest = JSON.parse(e.postData.contents);
      giftcardAction = giftcardRequest.action || "";
    }
  } catch (giftcardRouteError) {
    giftcardAction = "";
  }

  if (giftcardAction === "savePdf") {
    return saveGiftcardPdfToDrive_(e);
  }

  return ContentService
    .createTextOutput(JSON.stringify({
      ok: false,
      error: "不明なリクエストです"
    }))
    .setMimeType(ContentService.MimeType.JSON);
}


/**
 * Base64形式のPDFをGoogle Driveの
 * 「ギフトカードPDF」フォルダへ保存する
 */
function saveGiftcardPdfToDrive_(e) {
  try {
    var requestData = {};

    if (e && e.postData && e.postData.contents) {
      requestData = JSON.parse(e.postData.contents);
    }

    var pdfBase64 = String(requestData.pdfBase64 || "")
      .replace(/^data:application\/pdf;base64,/, "");

    if (!pdfBase64) {
      throw new Error("PDFデータがありません");
    }

    var folderName = "ギフトカードPDF";
    var folders = DriveApp.getFoldersByName(folderName);
    var folder = folders.hasNext()
      ? folders.next()
      : DriveApp.createFolder(folderName);

    var requestedName = String(requestData.fileName || "");
    var fileName = requestedName
      .replace(/[\\\/:*?"<>|]/g, "_")
      .trim();

    if (!fileName) {
      fileName =
        "ギフトカード_" +
        Utilities.formatDate(
          new Date(),
          Session.getScriptTimeZone() || "Asia/Tokyo",
          "yyyyMMdd_HHmmss"
        ) +
        ".pdf";
    }

    if (!/\.pdf$/i.test(fileName)) {
      fileName += ".pdf";
    }

    var pdfBytes = Utilities.base64Decode(pdfBase64);
    var pdfBlob = Utilities.newBlob(
      pdfBytes,
      MimeType.PDF,
      fileName
    );

    var savedFile = folder.createFile(pdfBlob);

    return ContentService
      .createTextOutput(JSON.stringify({
        ok: true,
        fileName: savedFile.getName(),
        fileId: savedFile.getId(),
        fileUrl: savedFile.getUrl(),
        folderName: folderName,
        folderId: folder.getId()
      }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService
      .createTextOutput(JSON.stringify({
        ok: false,
        error: String(error && error.message ? error.message : error)
      }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}


/**
 * 初回のみ実行してGoogle Driveへのアクセスを許可する
 */
function authorizeGiftcardDrive() {
  var folderName = "ギフトカードPDF";
  var folders = DriveApp.getFoldersByName(folderName);

  var folder = folders.hasNext()
    ? folders.next()
    : DriveApp.createFolder(folderName);

  Logger.log("Google Drive認証完了: " + folder.getUrl());
}
