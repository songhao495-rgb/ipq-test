// === 把這段程式碼貼到 Google Apps Script 編輯器裡 ===

function doPost(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // 建立「原始資料」工作表（如果不存在）
  var sheet = ss.getSheetByName('原始資料');
  if (!sheet) {
    sheet = ss.insertSheet('原始資料');
    sheet.appendRow([
      '時間', '年齡', '職業',
      'Q1', 'Q2', 'Q3', 'Q4', 'Q5', 'Q6',
      'Q7', 'Q8', 'Q9', 'Q10', 'Q11', 'Q12',
      'IPQ分數', '等級', '主型', '副型', '弱點區'
    ]);
    sheet.getRange(1, 1, 1, 20).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }

  var raw;
  if (e.parameter && e.parameter.payload) {
    raw = e.parameter.payload;
  } else {
    raw = e.postData.contents;
  }
  var data;
  try {
    data = JSON.parse(raw);
  } catch (err) {
    try { data = JSON.parse(raw.replace(/^\[|\]$/g, '')); } catch (e2) { return ContentService.createTextOutput('parse error'); }
  }

  sheet.appendRow([
    new Date(),
    data.age || '',
    data.job || '',
    data.q1, data.q2, data.q3, data.q4, data.q5, data.q6,
    data.q7, data.q8, data.q9, data.q10, data.q11, data.q12,
    data.ipq,
    data.tier,
    data.primaryType,
    data.secondaryType,
    data.weakest
  ]);

  // 更新統計
  updateStats_(ss);

  return ContentService.createTextOutput(JSON.stringify({status: 'ok'}))
    .setMimeType(ContentService.MimeType.JSON);
}

function updateStats_(ss) {
  var dataSheet = ss.getSheetByName('原始資料');
  if (!dataSheet) return;

  var statsSheet = ss.getSheetByName('統計');
  if (!statsSheet) {
    statsSheet = ss.insertSheet('統計');
  }
  statsSheet.clear();

  var lastRow = dataSheet.getLastRow();
  if (lastRow <= 1) {
    statsSheet.getRange(1, 1).setValue('尚無資料');
    return;
  }

  var total = lastRow - 1;
  var allData = dataSheet.getRange(2, 1, total, 20).getValues();

  // 統計主型
  var typeCounts = { '深海型': 0, '水銀型': 0, '火焰型': 0, '雷霆型': 0 };
  // 統計等級
  var tierCounts = {};
  // 統計年齡
  var ageCounts = {};
  // 統計弱點
  var weakCounts = { '定錨': 0, '煉化': 0, '覺醒': 0, '突變': 0 };
  // IPQ 分數加總
  var ipqSum = 0;

  for (var i = 0; i < allData.length; i++) {
    var row = allData[i];
    // 主型 (column R = index 17)
    var pt = row[17];
    if (typeCounts.hasOwnProperty(pt)) typeCounts[pt]++;

    // 等級 (column Q = index 16)
    var tier = row[16];
    if (tier) tierCounts[tier] = (tierCounts[tier] || 0) + 1;

    // 年齡 (column B = index 1)
    var age = row[1];
    if (age) ageCounts[age] = (ageCounts[age] || 0) + 1;

    // 弱點 (column T = index 19)
    var wk = row[19];
    if (weakCounts.hasOwnProperty(wk)) weakCounts[wk]++;

    // IPQ (column P = index 15)
    var ipq = Number(row[15]);
    if (!isNaN(ipq)) ipqSum += ipq;
  }

  var r = 1;

  // 標題
  statsSheet.getRange(r, 1).setValue('IPQ 測驗統計').setFontWeight('bold').setFontSize(14);
  r++;
  statsSheet.getRange(r, 1).setValue('總測驗人數').setFontWeight('bold');
  statsSheet.getRange(r, 2).setValue(total);
  r++;
  statsSheet.getRange(r, 1).setValue('平均 IPQ 分數').setFontWeight('bold');
  statsSheet.getRange(r, 2).setValue(total > 0 ? Math.round(ipqSum / total * 10) / 10 : 0);
  r += 2;

  // 主型分佈
  statsSheet.getRange(r, 1).setValue('主型分佈').setFontWeight('bold').setFontSize(12);
  r++;
  statsSheet.getRange(r, 1, 1, 3).setValues([['類型', '人數', '百分比']]).setFontWeight('bold');
  r++;
  var typeNames = ['深海型', '水銀型', '火焰型', '雷霆型'];
  for (var t = 0; t < typeNames.length; t++) {
    var cnt = typeCounts[typeNames[t]];
    var pct = total > 0 ? (cnt / total * 100).toFixed(1) + '%' : '0%';
    statsSheet.getRange(r, 1, 1, 3).setValues([[typeNames[t], cnt, pct]]);
    r++;
  }
  r++;

  // 等級分佈
  statsSheet.getRange(r, 1).setValue('等級分佈').setFontWeight('bold').setFontSize(12);
  r++;
  statsSheet.getRange(r, 1, 1, 3).setValues([['等級', '人數', '百分比']]).setFontWeight('bold');
  r++;
  for (var tier in tierCounts) {
    var pct = (tierCounts[tier] / total * 100).toFixed(1) + '%';
    statsSheet.getRange(r, 1, 1, 3).setValues([[tier, tierCounts[tier], pct]]);
    r++;
  }
  r++;

  // 年齡分佈
  statsSheet.getRange(r, 1).setValue('年齡分佈').setFontWeight('bold').setFontSize(12);
  r++;
  statsSheet.getRange(r, 1, 1, 3).setValues([['年齡', '人數', '百分比']]).setFontWeight('bold');
  r++;
  var ageOrder = ['18-24', '25-30', '31-36', '37-45', '46+'];
  for (var a = 0; a < ageOrder.length; a++) {
    if (ageCounts[ageOrder[a]]) {
      var pct = (ageCounts[ageOrder[a]] / total * 100).toFixed(1) + '%';
      statsSheet.getRange(r, 1, 1, 3).setValues([[ageOrder[a], ageCounts[ageOrder[a]], pct]]);
      r++;
    }
  }
  r++;

  // 弱點分佈
  statsSheet.getRange(r, 1).setValue('弱點區分佈').setFontWeight('bold').setFontSize(12);
  r++;
  statsSheet.getRange(r, 1, 1, 3).setValues([['弱點區', '人數', '百分比']]).setFontWeight('bold');
  r++;
  var weakNames = ['定錨', '煉化', '覺醒', '突變'];
  for (var w = 0; w < weakNames.length; w++) {
    var cnt = weakCounts[weakNames[w]];
    var pct = total > 0 ? (cnt / total * 100).toFixed(1) + '%' : '0%';
    statsSheet.getRange(r, 1, 1, 3).setValues([[weakNames[w], cnt, pct]]);
    r++;
  }

  // 自動調整欄寬
  statsSheet.autoResizeColumns(1, 3);
}

// 手動跑一次統計（測試用）
function manualUpdateStats() {
  updateStats_(SpreadsheetApp.getActiveSpreadsheet());
}
