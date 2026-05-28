// ============================================
// 果茶配方记录 — Google Apps Script
// ============================================
// 使用方法：
// 1. 新建一个 Google Sheet
// 2. 在 Sheet 底部创建两个 sheet tab，分别命名为 "recipes" 和 "ingredients"
// 3. 点击菜单 Extensions > Apps Script
// 4. 把这段代码粘贴进去，替换掉默认的 function myFunction(){}
// 5. 点 Deploy > New deployment > Web app
//    - Execute as: Me
//    - Who has access: Anyone
// 6. 复制部署后的 URL，粘贴到 HTML 文件里的 SCRIPT_URL 变量
// ============================================

function doGet(e) {
  var action = e.parameter.action;

  // ---------- RECIPES ----------
  if (action === 'getRecipes') {
    return sendJson(readAll('recipes'));
  }

  if (action === 'saveRecipe') {
    var data = JSON.parse(e.parameter.data);
    var sheet = getSheet('recipes');
    // 找到是否已存在（更新 madeCount 时）
    var rows = sheet.getDataRange().getValues();
    var found = false;
    for (var i = 1; i < rows.length; i++) {
      if (rows[i][0] === data.id) {
        // 更新这一行
        sheet.getRange(i + 1, 1, 1, 6).setValues([[
          data.id,
          data.name,
          data.createdAt,
          data.rating,
          data.madeCount || 0,
          JSON.stringify(data.items)
        ]]);
        // 强制文本格式（避免 Google Sheet 自动转换）
        sheet.getRange(i + 1, 1, 1, 6).setNumberFormat('@');
        found = true;
        break;
      }
    }
    if (!found) {
      sheet.appendRow([
        data.id,
        data.name,
        data.createdAt,
        data.rating,
        data.madeCount || 0,
        JSON.stringify(data.items)
      ]);
      var lastRow = sheet.getLastRow();
      sheet.getRange(lastRow, 1, 1, 6).setNumberFormat('@');
    }
    if (data.notes !== undefined && data.notes !== null) {
      // 备注放第7列
      var rowIdx = findRow(sheet, data.id);
      if (rowIdx > 0) {
        sheet.getRange(rowIdx, 7).setValue(data.notes);
        sheet.getRange(rowIdx, 7).setNumberFormat('@');
      }
    }
    return sendJson({ ok: true });
  }

  if (action === 'deleteRecipe') {
    var id = e.parameter.id;
    var sheet = getSheet('recipes');
    var rows = sheet.getDataRange().getValues();
    for (var i = 1; i < rows.length; i++) {
      if (String(rows[i][0]) === String(id)) {
        sheet.deleteRow(i + 1);
        break;
      }
    }
    return sendJson({ ok: true });
  }

  // ---------- INGREDIENTS ----------
  if (action === 'getIngredients') {
    return sendJson(readAll('ingredients'));
  }

  if (action === 'saveIngredient') {
    var data = JSON.parse(e.parameter.data);
    var sheet = getSheet('ingredients');
    sheet.appendRow([
      data.id,
      data.name,
      data.emoji,
      data.category,
      data.totalSize,
      data.unit,
      data.price,
      data.portionUnit
    ]);
    var lastRow = sheet.getLastRow();
    sheet.getRange(lastRow, 1, 1, 8).setNumberFormat('@');
    return sendJson({ ok: true });
  }

  if (action === 'deleteIngredient') {
    var id = e.parameter.id;
    var sheet = getSheet('ingredients');
    var rows = sheet.getDataRange().getValues();
    for (var i = 1; i < rows.length; i++) {
      if (String(rows[i][0]) === String(id)) {
        sheet.deleteRow(i + 1);
        break;
      }
    }
    return sendJson({ ok: true });
  }

  return sendJson({ error: 'unknown action' });
}

// ── Helpers ──

function getSheet(name) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    if (name === 'recipes') {
      sheet.appendRow(['id', 'name', 'createdAt', 'rating', 'madeCount', 'items', 'notes']);
    } else if (name === 'ingredients') {
      sheet.appendRow(['id', 'name', 'emoji', 'category', 'totalSize', 'unit', 'price', 'portionUnit']);
    }
    // 表头也设为文本
    sheet.getRange(1, 1, 1, 8).setNumberFormat('@');
  }
  return sheet;
}

function readAll(sheetName) {
  var sheet = getSheet(sheetName);
  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];

  var headers = data[0];
  var result = [];
  for (var i = 1; i < data.length; i++) {
    var obj = {};
    for (var j = 0; j < headers.length; j++) {
      var key = headers[j];
      var val = data[i][j];

      // 尝试解析 JSON 字段（items）
      if (key === 'items') {
        try { val = JSON.parse(val); } catch (e) { val = []; }
      }
      // 数字字段转回数字
      if (key === 'rating' || key === 'madeCount' || key === 'totalSize' || key === 'price') {
        val = Number(val) || 0;
      }
      obj[key] = val;
    }
    result.push(obj);
  }
  return result;
}

function findRow(sheet, id) {
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(id)) return i + 1;
  }
  return -1;
}

function sendJson(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
