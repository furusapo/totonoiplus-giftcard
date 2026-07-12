function doGet(e) {
  const mode = e.parameter.mode;
  const action = e.parameter.action;
  const callback = e.parameter.callback;

  if (mode === "check") return checkGiftCard(e);
  if (mode === "use") return useGiftCard(e);

  // スタッフ管理画面用
  if (mode === "staff_search") return staffSearch(e);
  if (mode === "staff_use") return staffUse(e);
  if (mode === "staff_cancel_use_v2") return staffCancelUseV2(e);

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
    const issuedSerials = [];

    requests.forEach(item => {
      const design = Number(item.design);
      const count = Number(item.count || 0);

      for (let i = 0; i < count; i++) {
        const serial = getNextSerial(sheet, design);
        issuedSerials.push(serial);
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
      message: "発行履歴に保存しました",
      serials: issuedSerials
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


/**
 * 二重チェック付き利用済み取消
 *
 * 必須：
 * ・理由カテゴリ
 * ・20文字以上の詳細理由
 * ・操作担当
 * ・別の確認担当
 * ・2項目の確認同意
 */
function staffCancelUseV2(e) {
  const callback = e.parameter.callback;
  const serial = String(e.parameter.serial || "").trim().toUpperCase();
  const reasonCategory = String(e.parameter.reasonCategory || "").trim();
  const reasonDetail = String(e.parameter.reasonDetail || "").trim();
  const operator = String(e.parameter.operator || "").trim();
  const reviewer = String(e.parameter.reviewer || "").trim();
  const contentConfirmed = String(e.parameter.contentConfirmed || "") === "true";
  const auditConfirmed = String(e.parameter.auditConfirmed || "") === "true";

  const allowedCategories = [
    "操作ミス",
    "シリアル入力ミス",
    "お客様都合",
    "システム障害",
    "その他"
  ];

  if (!serial) {
    return respond(callback, {
      ok: false,
      message: "シリアル番号がありません"
    });
  }

  if (allowedCategories.indexOf(reasonCategory) === -1) {
    return respond(callback, {
      ok: false,
      message: "訂正理由のカテゴリを選択してください"
    });
  }

  if (reasonDetail.length < 20) {
    return respond(callback, {
      ok: false,
      message: "訂正理由の詳細を20文字以上で入力してください"
    });
  }

  if (!operator) {
    return respond(callback, {
      ok: false,
      message: "操作担当者を選択してください"
    });
  }

  if (!reviewer) {
    return respond(callback, {
      ok: false,
      message: "確認担当者を選択してください"
    });
  }

  if (operator === reviewer) {
    return respond(callback, {
      ok: false,
      message: "確認担当者は操作担当者とは別の社員を選択してください"
    });
  }

  if (!contentConfirmed || !auditConfirmed) {
    return respond(callback, {
      ok: false,
      message: "2つの確認項目にチェックしてください"
    });
  }

  const lock = LockService.getScriptLock();

  try {
    lock.waitLock(10000);

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("発行履歴");

    if (!sheet) {
      throw new Error("発行履歴シートが見つかりません");
    }

    ensureStaffColumns(sheet);

    const values = sheet.getDataRange().getValues();
    const headers = values[0].map(function(value) {
      return String(value || "").trim();
    });

    const serialCol = findHeader(
      headers,
      ["シリアル番号", "シリアル", "発行番号", "ギフトカード番号"]
    );

    const statusCol = findHeader(
      headers,
      ["状態", "ステータス", "使用状態"]
    );

    const usedAtCol = findHeader(
      headers,
      ["使用日時", "利用日時"]
    );

    const usedStaffCol = findHeader(
      headers,
      ["使用スタッフ", "利用スタッフ", "対応スタッフ"]
    );

    const memoCol = findHeader(
      headers,
      ["使用メモ", "利用メモ", "メモ"]
    );

    if (serialCol === -1 || statusCol === -1) {
      throw new Error("シリアル番号または状態の列が見つかりません");
    }

    for (let i = 1; i < values.length; i++) {
      const rowSerial = String(
        values[i][serialCol] || ""
      ).trim().toUpperCase();

      if (rowSerial !== serial) continue;

      const currentStatus = normalizeStatus(values[i][statusCol]);

      if (currentStatus !== "利用済") {
        return respond(callback, {
          ok: false,
          message:
            "現在の状態は「" +
            currentStatus +
            "」のため、取消できません"
        });
      }

      const originalUsedAt =
        usedAtCol !== -1 ? values[i][usedAtCol] : "";

      const originalUsedStaff =
        usedStaffCol !== -1 ? values[i][usedStaffCol] : "";

      const originalMemo =
        memoCol !== -1 ? values[i][memoCol] : "";

      const historyId = Utilities.getUuid();
      const logSheet = ensureCancelAuditSheet_();

      // 元情報を先に監査ログへ記録してから状態を変更
      logSheet.appendRow([
        new Date(),
        historyId,
        serial,
        "利用済み取消",
        "利用済",
        "未使用",
        reasonCategory,
        reasonDetail,
        operator,
        reviewer,
        originalUsedAt,
        originalUsedStaff,
        originalMemo,
        "スタッフ管理画面"
      ]);

      sheet.getRange(i + 1, statusCol + 1).setValue("未使用");

      if (usedAtCol !== -1) {
        sheet.getRange(i + 1, usedAtCol + 1).clearContent();
      }

      if (usedStaffCol !== -1) {
        sheet.getRange(i + 1, usedStaffCol + 1).clearContent();
      }

      if (memoCol !== -1) {
        sheet.getRange(i + 1, memoCol + 1).clearContent();
      }

      SpreadsheetApp.flush();

      return respond(callback, {
        ok: true,
        message:
          serial +
          " の利用済み処理を訂正し、未使用へ戻しました",
        historyId: historyId
      });
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

  } finally {
    try {
      lock.releaseLock();
    } catch (ignore) {}
  }
}


/**
 * 利用済み取消の監査ログシートを取得または作成
 */
function ensureCancelAuditSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("操作履歴");

  const headers = [
    "記録日時",
    "履歴ID",
    "シリアル番号",
    "操作",
    "変更前",
    "変更後",
    "理由カテゴリ",
    "理由詳細",
    "操作担当",
    "確認担当",
    "元の使用日時",
    "元の使用スタッフ",
    "元の使用メモ",
    "接続元"
  ];

  if (!sheet) {
    sheet = ss.insertSheet("操作履歴");
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold");
    sheet.autoResizeColumns(1, headers.length);
  }

  return sheet;
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

  if (giftcardAction === "saveIssueRecord") {
    return saveGiftcardIssueRecord_(e);
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


/**
 * ギフトカードの発行記録をGoogle Driveへ保存する
 *
 * 保存内容：
 * ・署名.png
 * ・発行データ.json
 * ・発行記録.pdf
 */
function saveGiftcardIssueRecord_(e) {
  try {
    var data = {};

    if (e && e.postData && e.postData.contents) {
      data = JSON.parse(e.postData.contents);
    }

    var serials = Array.isArray(data.serials)
      ? data.serials.map(function(value) {
          return String(value || "").trim().toUpperCase();
        }).filter(String)
      : [];

    if (!serials.length) {
      throw new Error("確定シリアル番号がありません");
    }

    var signatureImage = String(data.signatureImage || "");
    var signatureBase64 = signatureImage.replace(
      /^data:image\/(?:png|jpeg|jpg);base64,/i,
      ""
    );

    if (!signatureBase64) {
      throw new Error("署名画像がありません");
    }

    var now = new Date();
    var timezone = Session.getScriptTimeZone() || "Asia/Tokyo";
    var monthName = Utilities.formatDate(now, timezone, "yyyy-MM");
    var timestamp = Utilities.formatDate(now, timezone, "yyyyMMdd_HHmmss");

    var rootFolder = getOrCreateGiftcardFolder_(
      DriveApp.getRootFolder(),
      "ギフトカード発行記録"
    );

    var monthFolder = getOrCreateGiftcardFolder_(
      rootFolder,
      monthName
    );

    var firstSerial = serials[0];
    var folderLabel = firstSerial;

    if (serials.length > 1) {
      folderLabel += "_ほか" + (serials.length - 1) + "枚";
    }

    var issueFolder = monthFolder.createFolder(
      folderLabel + "_" + timestamp
    );

    var signatureBytes = Utilities.base64Decode(signatureBase64);
    var signatureBlob = Utilities.newBlob(
      signatureBytes,
      "image/png",
      "署名.png"
    );

    var signatureFile = issueFolder.createFile(signatureBlob);

    var record = {
      savedAt: Utilities.formatDate(
        now,
        timezone,
        "yyyy/MM/dd HH:mm:ss"
      ),
      serials: serials,
      phone: String(data.phone || ""),
      staff: String(data.staff || ""),
      issueDate: String(data.issueDate || ""),
      expireDate: String(data.expireDate || ""),
      requests: Array.isArray(data.requests) ? data.requests : [],
      envelopeCount: Number(data.envelopeCount || 0),
      cardPrice: Number(data.cardPrice || 0),
      issueFee: Number(data.issueFee || 0),
      envelopeFee: Number(data.envelopeFee || 0),
      totalPrice: Number(data.totalPrice || 0),
      signatureAt: String(data.signatureAt || ""),
      signatureFileId: signatureFile.getId(),
      signatureFileUrl: signatureFile.getUrl()
    };

    var jsonBlob = Utilities.newBlob(
      JSON.stringify(record, null, 2),
      "application/json",
      "発行データ.json"
    );

    var jsonFile = issueFolder.createFile(jsonBlob);

    var html = buildGiftcardIssueRecordHtml_(record, signatureImage);

    // HTMLをPDFへ変換してGoogle Driveへ保存
    var pdfBlob = HtmlService
      .createHtmlOutput(html)
      .getAs(MimeType.PDF)
      .setName("発行記録.pdf");

    var pdfFile = issueFolder.createFile(pdfBlob);

    return ContentService
      .createTextOutput(JSON.stringify({
        ok: true,
        folderName: issueFolder.getName(),
        folderId: issueFolder.getId(),
        folderUrl: issueFolder.getUrl(),
        signatureUrl: signatureFile.getUrl(),
        jsonUrl: jsonFile.getUrl(),
        pdfUrl: pdfFile.getUrl()
      }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService
      .createTextOutput(JSON.stringify({
        ok: false,
        error: String(
          error && error.message
            ? error.message
            : error
        )
      }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}


/**
 * 指定した親フォルダ内にフォルダを取得または作成する
 */
function getOrCreateGiftcardFolder_(parentFolder, folderName) {
  var folders = parentFolder.getFoldersByName(folderName);

  if (folders.hasNext()) {
    return folders.next();
  }

  return parentFolder.createFolder(folderName);
}


/**
 * 発行記録確認用HTMLを作成する
 */
function buildGiftcardIssueRecordHtml_(record, signatureImage) {
  var requestsHtml = (record.requests || []).map(function(item) {
    return (
      "<tr>" +
        "<td>Design " +
          escapeGiftcardHtml_(
            String(item.design || "").padStart(2, "0")
          ) +
        "</td>" +
        "<td>" +
          escapeGiftcardHtml_(String(item.count || 0)) +
          "枚</td>" +
      "</tr>"
    );
  }).join("");

  var serialHtml = (record.serials || []).map(function(serial) {
    return "<span class=\"serial\">" +
      escapeGiftcardHtml_(serial) +
      "</span>";
  }).join("");

  return [
    "<!DOCTYPE html>",
    "<html lang=\"ja\">",
    "<head>",
    "<meta charset=\"UTF-8\">",
    "<meta name=\"viewport\" content=\"width=device-width,initial-scale=1\">",
    "<title>ギフトカード発行記録</title>",
    "<style>",
    "body{font-family:-apple-system,BlinkMacSystemFont,'Hiragino Kaku Gothic ProN',sans-serif;background:#f5f2eb;color:#222;margin:0;padding:30px}",
    ".sheet{max-width:760px;margin:auto;background:#fff;padding:36px;border-radius:18px;box-shadow:0 10px 40px rgba(0,0,0,.1)}",
    "h1{margin-top:0;color:#9b773c}",
    "h2{font-size:17px;margin-top:28px;border-bottom:1px solid #ddd;padding-bottom:8px}",
    "table{width:100%;border-collapse:collapse}",
    "td,th{border-bottom:1px solid #eee;padding:10px;text-align:left}",
    ".serial{display:inline-block;margin:4px;padding:7px 12px;border:1px solid #b8975a;border-radius:999px;color:#8a642c}",
    ".signature{max-width:100%;height:auto;border:1px solid #ccc;border-radius:10px;background:#fff}",
    ".meta{line-height:1.9}",
    "</style>",
    "</head>",
    "<body>",
    "<div class=\"sheet\">",
    "<h1>TOTONOI+ ギフトカード発行記録</h1>",
    "<div class=\"meta\">",
    "<strong>保存日時：</strong>" +
      escapeGiftcardHtml_(record.savedAt) + "<br>",
    "<strong>発行日：</strong>" +
      escapeGiftcardHtml_(record.issueDate) + "<br>",
    "<strong>有効期限：</strong>" +
      escapeGiftcardHtml_(record.expireDate) + "<br>",
    "<strong>電話番号：</strong>" +
      escapeGiftcardHtml_(record.phone) + "<br>",
    "<strong>発行担当：</strong>" +
      escapeGiftcardHtml_(record.staff),
    "</div>",
    "<h2>シリアル番号</h2>",
    "<div>" + serialHtml + "</div>",
    "<h2>発行内容</h2>",
    "<table>",
    "<thead><tr><th>デザイン</th><th>枚数</th></tr></thead>",
    "<tbody>" + requestsHtml + "</tbody>",
    "</table>",
    "<h2>金額</h2>",
    "<div class=\"meta\">",
    "カード代：" +
      Number(record.cardPrice || 0).toLocaleString("ja-JP") +
      "円<br>",
    "発行手数料：" +
      Number(record.issueFee || 0).toLocaleString("ja-JP") +
      "円<br>",
    "追加封筒：" +
      Number(record.envelopeFee || 0).toLocaleString("ja-JP") +
      "円<br>",
    "<strong>合計：" +
      Number(record.totalPrice || 0).toLocaleString("ja-JP") +
      "円</strong>",
    "</div>",
    "<h2>お客様署名</h2>",
    "<img class=\"signature\" src=\"" +
      escapeGiftcardHtml_(signatureImage) +
      "\" alt=\"お客様署名\">",
    "</div>",
    "</body>",
    "</html>"
  ].join("");
}


/**
 * HTMLへ安全に文字列を埋め込む
 */
function escapeGiftcardHtml_(value) {
  return String(value == null ? "" : value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}
