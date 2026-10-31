// =========================================================================
// 🔐 GOOGLE APPS SCRIPT WITH TOKEN SECURITY FOR POINT & ACCOUNT LEDGER
// =========================================================================

// 🔴 APNA SECRET TOKEN YAHAN SET KAREIN:
const SECRET_TOKEN = "MY_SECRET_TOKEN_123";

function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      return responseJSON({ status: "error", message: "No data received" });
    }

    var request = JSON.parse(e.postData.contents);
    var userToken = (request.token || "").toString().trim();

    // 🔒 Strictly Validate Token Before Allowing Any Changes
    if (!userToken || userToken !== SECRET_TOKEN) {
      return responseJSON({
        status: "error",
        message: "⛔ Access Denied: Invalid or missing Security Token!"
      });
    }

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    
    if (request.action === "upload") {
      var entries = request.data.entries || [];
      var accounts = request.data.accounts || [];

      // 🛡️ Safety Block: Prevent wiping sheet if empty data is accidentally received from a new device
      var entrySheet = ss.getSheetByName("Entries") || ss.insertSheet("Entries");
      if (entrySheet.getLastRow() > 1 && entries.length === 0 && !request.allowEmpty) {
        return responseJSON({
          status: "error",
          message: "⚠️ Safety Block: Google Sheet already has data, but 0 entries were sent. Read data first to prevent data loss!"
        });
      }

      // 1. Entries Sheet
      entrySheet.clearContents();
      entrySheet.appendRow([
        "Sr", "ID", "Type", "Date", "Account", "Pay In", "Pay Out", 
        "Opening Point Balance", "Closing Point Balance", "Points Profit", 
        "Ac Balance", "Commission", "Exp", "Rate Info", "Remark"
      ]);
      
      for (var i = 0; i < entries.length; i++) {
        var row = entries[i];
        var dateStr = "'" + (row.date || ""); // Prefix with ' to preserve literal text YYYY-MM-DD
        
        if (row.type === "transfer") {
          entrySheet.appendRow([
            i + 1,
            row.id || ("tr_" + (i + 1)),
            "transfer",
            dateStr,
            row.acName || "",
            0,
            0,
            "-",
            "-",
            "Transfer",
            row.acBalance || "",
            0,
            0,
            "-",
            row.remark || ""
          ]);
        } else {
          var rateStr = (row.rateToggle && row.rate) ? ("Rate: " + row.rate + " | Profit: " + (row.rateProfitVal || 0)) : "-";
          entrySheet.appendRow([
            i + 1,
            row.id || ("entry_" + (i + 1)),
            "entry",
            dateStr,
            row.acName || "",
            Number(row.payIn) || 0,
            Number(row.payOut) || 0,
            Number(row.openingPointBal) || 0,
            Number(row.closingPointBal) || 0,
            Number(row.pointsProfit) || 0,
            row.acBalance || 0,
            Number(row.commission) || 0,
            Number(row.exp) || 0,
            rateStr,
            row.remark || ""
          ]);
        }
      }
      
      // 2. Accounts Sheet
      var acSheet = ss.getSheetByName("Accounts") || ss.insertSheet("Accounts");
      acSheet.clearContents();
      acSheet.appendRow(["ID", "Account Name", "Current Balance"]);
      for (var j = 0; j < accounts.length; j++) {
        acSheet.appendRow([
          accounts[j].id || ("ac_" + (j + 1)),
          accounts[j].name || "",
          Number(accounts[j].balance) || 0
        ]);
      }
      
      // 3. Telegram Sheet
      var tgSheet = ss.getSheetByName("Telegram") || ss.insertSheet("Telegram");
      tgSheet.clearContents();
      tgSheet.appendRow(["Configuration", "Value"]);
      
      var tgData = request.data.telegram || {};
      tgSheet.appendRow(["Bot Token", tgData.botToken || ""]);
      tgSheet.appendRow(["Notifications Enabled", tgData.enabled !== false ? "YES" : "NO"]);
      tgSheet.appendRow(["---", "---"]);
      tgSheet.appendRow(["Sr", "Recipient Name", "Chat ID"]);
      
      var chatIds = tgData.chatIds || [];
      for (var k = 0; k < chatIds.length; k++) {
        tgSheet.appendRow([k + 1, chatIds[k].name || "Recipient", String(chatIds[k].chatId || "")]);
      }
      
      return responseJSON({
        status: "success",
        message: "Data uploaded successfully without duplicates!",
        count: entries.length
      });
    }

    return responseJSON({ status: "error", message: "Unknown action requested." });

  } catch (err) {
    return responseJSON({ status: "error", message: err.toString() });
  }
}

function doGet(e) {
  try {
    var userToken = (e && e.parameter && e.parameter.token ? e.parameter.token : "").toString().trim();

    // 🔒 Strictly Validate Token Before Allowing Any Data Read
    if (!userToken || userToken !== SECRET_TOKEN) {
      return responseJSON({
        status: "error",
        message: "⛔ Access Denied: Invalid or missing Security Token!"
      });
    }

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // 1. Read Entries Sheet
    var entrySheet = ss.getSheetByName("Entries");
    var entries = [];
    if (entrySheet && entrySheet.getLastRow() > 1) {
      var rows = entrySheet.getDataRange().getValues();
      var headers = rows[0] || [];
      var isNewFormat = String(headers[1] || "").trim().toUpperCase() === "ID";

      for (var i = 1; i < rows.length; i++) {
        var r = rows[i];
        if (isNewFormat) {
          if (!r[1] && !r[3]) continue;
          var idVal = String(r[1] || ("gs_" + i));
          var typeVal = String(r[2] || "entry").toLowerCase();
          var dateVal = formatCellDate(r[3]);
          var acNameVal = String(r[4] || "");
          var isTransfer = typeVal === "transfer" || String(r[9]).toLowerCase().indexOf("transfer") !== -1;

          entries.push({
            id: idVal,
            type: isTransfer ? "transfer" : "entry",
            date: dateVal,
            acName: acNameVal,
            payIn: isTransfer ? 0 : (Number(r[5]) || 0),
            payOut: isTransfer ? 0 : (Number(r[6]) || 0),
            openingPointBal: isTransfer ? "-" : (Number(r[7]) || 0),
            closingPointBal: isTransfer ? "-" : (Number(r[8]) || 0),
            pointsProfit: isTransfer ? 0 : (Number(r[9]) || 0),
            acBalance: r[10],
            commission: isTransfer ? 0 : (Number(r[11]) || 0),
            exp: isTransfer ? 0 : (Number(r[12]) || 0),
            remark: String(r[14] || "")
          });
        } else {
          // Legacy format fallback: ["Sr", "Date", "Account", "Pay In", "Pay Out", ...]
          if (!r[1]) continue;
          var dateVal = formatCellDate(r[1]);
          var acNameVal = String(r[2] || "");
          var isTransfer = String(r[7]).toLowerCase().indexOf("transfer") !== -1;

          entries.push({
            id: "gs_" + i,
            type: isTransfer ? "transfer" : "entry",
            date: dateVal,
            acName: acNameVal,
            payIn: isTransfer ? 0 : (Number(r[3]) || 0),
            payOut: isTransfer ? 0 : (Number(r[4]) || 0),
            openingPointBal: isTransfer ? "-" : (Number(r[5]) || 0),
            closingPointBal: isTransfer ? "-" : (Number(r[6]) || 0),
            pointsProfit: isTransfer ? 0 : (Number(r[7]) || 0),
            acBalance: r[8],
            commission: isTransfer ? 0 : (Number(r[9]) || 0),
            exp: isTransfer ? 0 : (Number(r[10]) || 0),
            remark: ""
          });
        }
      }
    }
    
    // 2. Read Accounts Sheet
    var acSheet = ss.getSheetByName("Accounts");
    var accounts = [];
    if (acSheet && acSheet.getLastRow() > 1) {
      var aRows = acSheet.getDataRange().getValues();
      for (var j = 1; j < aRows.length; j++) {
        if (aRows[j][1]) {
          accounts.push({
            id: String(aRows[j][0] || ("ac_" + j)),
            name: String(aRows[j][1]).trim(),
            balance: Number(aRows[j][2]) || 0
          });
        }
      }
    }

    // 3. Read Telegram Bot Config Sheet
    var tgSheet = ss.getSheetByName("Telegram");
    var telegramData = { botToken: "", chatIds: [], enabled: true };
    if (tgSheet && tgSheet.getLastRow() > 1) {
      var tRows = tgSheet.getDataRange().getValues();
      var startRecipients = false;
      for (var m = 1; m < tRows.length; m++) {
        var key = String(tRows[m][0]).trim();
        var val = tRows[m][1];
        if (key === "Bot Token") {
          telegramData.botToken = String(val || "").trim();
        } else if (key === "Notifications Enabled") {
          telegramData.enabled = String(val).toUpperCase() === "YES";
        } else if (key === "Sr" && String(val).trim() === "Recipient Name") {
          startRecipients = true;
          continue;
        }

        if (startRecipients && tRows[m][2]) {
          telegramData.chatIds.push({
            id: "cid_gs_" + m,
            name: String(tRows[m][1] || "Recipient").trim(),
            chatId: String(tRows[m][2]).trim()
          });
        }
      }
    }

    return responseJSON({
      status: "success",
      data: {
        entries: entries,
        accounts: accounts,
        telegram: telegramData
      }
    });

  } catch (err) {
    return responseJSON({ status: "error", message: err.toString() });
  }
}

function formatCellDate(val) {
  if (!val) return "";
  if (val instanceof Date) {
    var year = val.getFullYear();
    var month = ("0" + (val.getMonth() + 1)).slice(-2);
    var day = ("0" + val.getDate()).slice(-2);
    return year + "-" + month + "-" + day;
  }
  var s = String(val).trim();
  if (s.indexOf("T") !== -1) {
    return s.split("T")[0];
  }
  return s;
}

function responseJSON(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
