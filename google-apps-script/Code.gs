/**
 * Google Apps Script — мост между планировщиком (index.html) и Google Таблицей.
 *
 * Установка:
 * 1. Создайте Google Таблицу.
 * 2. Расширения → Apps Script → вставьте этот файл.
 * 3. Секреты — в Script Properties (см. setupSecrets ниже), НЕ в публичном GitHub.
 * 4. Развернуть → Веб-приложение (доступ: «Все» — обязательно для браузера; защита — токен + PIN).
 *
 * Один раз в Apps Script выполните setupSecrets() и удалите/закомментируйте вызов.
 */

/** Только запасной вариант. Реальные значения — в Script Properties. */
const SPREADSHEET_ID = '';
const SECRET_TOKEN = '';

/**
 * Один раз: Run → setupSecrets → разрешить доступ.
 * Потом удалите тело или закомментируйте вызов setProperties с реальными значениями.
 */
function setupSecrets() {
  PropertiesService.getScriptProperties().setProperties({
    SPREADSHEET_ID: 'ВАШ_ID_ТАБЛИЦЫ',
    SECRET_TOKEN: 'длинный-случайный-токен-32-символа',
    SYNC_PIN: 'ваш-pin-5678',
  });
}

const SHEETS = {
  MONTHLY: 'По_месяцам',
  GOALS: 'Цели',
  META: '_meta',
};

var MONTHS_RU = ['Янв', 'Фев', 'Мар', 'Апр', 'Май', 'Июн', 'Июл', 'Авг', 'Сен', 'Окт', 'Ноя', 'Дек'];

function doGet(e) {
  try {
    checkToken(e);
    const action = e.parameter.action || 'export';
    if (action === 'ping') {
      return jsonResponse({ ok: true, message: 'pong' });
    }
    if (action === 'export') {
      return jsonResponse(exportData());
    }
    return jsonResponse({ error: 'Unknown action' });
  } catch (err) {
    return jsonResponse({ error: String(err) });
  }
}

function doOptions() {
  return ContentService.createTextOutput('');
}

function parsePostBody(e) {
  if (e.parameter && e.parameter.payload) {
    return JSON.parse(e.parameter.payload);
  }
  if (e.postData && e.postData.contents) {
    return JSON.parse(e.postData.contents);
  }
  throw new Error('Empty body');
}

function doPost(e) {
  try {
    checkToken(e);
    const body = parsePostBody(e);
    const action = body.action || 'import';
    if (action === 'import') {
      importData(body.data);
      return jsonResponse({ ok: true, message: 'Импортировано' });
    }
    if (action === 'syncMonth') {
      const result = syncMonthSheet(body.data);
      return jsonResponse(result);
    }
    return jsonResponse({ error: 'Unknown action' });
  } catch (err) {
    return jsonResponse({ error: String(err) });
  }
}

function getConfigValue(key, fallback) {
  var fromProps = PropertiesService.getScriptProperties().getProperty(key);
  if (fromProps && String(fromProps).trim()) return String(fromProps).trim();
  if (fallback && String(fallback).trim()) return String(fallback).trim();
  return '';
}

function extractPin(e) {
  var pin = e.parameter && e.parameter.pin;
  if (!pin) {
    try {
      pin = parsePostBody(e).pin;
    } catch (ignore) {}
  }
  return pin ? String(pin).trim() : '';
}

function checkToken(e) {
  var token = e.parameter && e.parameter.token;
  if (!token) {
    try {
      token = parsePostBody(e).token;
    } catch (ignore) {}
  }
  token = token ? String(token).trim() : '';
  var expected = getConfigValue('SECRET_TOKEN', SECRET_TOKEN);
  if (!expected) {
    throw new Error('SECRET_TOKEN не задан — выполните setupSecrets() или Project Settings → Script properties');
  }
  if (!token) {
    throw new Error('Missing token — укажите секретный токен в планировщике');
  }
  if (token !== expected) {
    throw new Error('Invalid token');
  }

  var expectedPin = getConfigValue('SYNC_PIN', '');
  if (expectedPin) {
    var pin = extractPin(e);
    if (!pin || pin !== expectedPin) {
      throw new Error('Invalid PIN — укажите PIN синхронизации в планировщике');
    }
  }
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * Из URL или «d/ID/edit…» оставляет только ID таблицы.
 */
function normalizeSpreadsheetId(raw) {
  var id = String(raw || '').trim();
  if (!id || id === 'ВСТАВЬТЕ_ID_ТАБЛИЦЫ') {
    throw new Error('Задайте SPREADSHEET_ID в Code.gs — только ID из URL таблицы');
  }
  var match = id.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
  if (match) return match[1];
  match = id.match(/(?:^|\/)d\/([a-zA-Z0-9-_]+)/);
  if (match) return match[1];
  // 19Z93.../edit?gid=0 или 19Z93...?gid=0
  match = id.match(/^([a-zA-Z0-9-_]{20,})(?:[/?#]|$)/);
  if (match) return match[1];
  throw new Error('Неверный SPREADSHEET_ID: ' + id + ' — вставьте только ID, например 19Z93po1nU1W4VVfstmLR-RsQU5g6ePqjbfiuQCRl2kM');
}

function getSs() {
  return SpreadsheetApp.openById(normalizeSpreadsheetId(SPREADSHEET_ID));
}

function ensureSheet(name, headers) {
  const ss = getSs();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    if (headers && headers.length) {
      sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    }
  }
  return sh;
}

/** getRange(row, col, numRows, numCols) — 3-й аргумент это КОЛИЧЕСТВО строк, не номер последней. */
function writeRows(sh, startRow, rows, colCount) {
  if (!rows || !rows.length) return;
  var cols = colCount || (rows[0] ? rows[0].length : 0);
  if (cols < 1) return;
  sh.getRange(startRow, 1, rows.length, cols).setValues(rows);
}

function monthKeyToLabelGs(mk) {
  var p = String(mk).split('-');
  var mi = parseInt(p[1], 10) - 1;
  return (MONTHS_RU[mi] || p[1]) + ' ' + p[0];
}

function collectMonthKeys(data) {
  var keys = {};
  if (data.incomeByMonth) {
    Object.keys(data.incomeByMonth).forEach(function(k) { keys[k] = true; });
  }
  if (data.expenseHistory) {
    data.expenseHistory.forEach(function(h) {
      if (h.monthKey) keys[h.monthKey] = true;
    });
  }
  if (data.savingsLedger) {
    data.savingsLedger.forEach(function(e) {
      if (e.monthKey) keys[e.monthKey] = true;
    });
  }
  if (data.monthKey) keys[data.monthKey] = true;
  return Object.keys(keys).sort();
}

function incomeValueForMonth(data, mk, srcId) {
  var row = data.incomeByMonth && data.incomeByMonth[mk];
  if (!row) return 0;
  if (row[srcId] !== undefined) return Number(row[srcId]) || 0;
  if (srcId === 'src_salary') return Number(row.a || row.v) || 0;
  if (srcId === 'src_side') return Number(row.b) || 0;
  return 0;
}

function expenseValuesForMonth(data, mk) {
  var hist = data.expenseHistory || [];
  for (var i = 0; i < hist.length; i++) {
    if (hist[i].monthKey === mk && hist[i].values) return hist[i].values;
  }
  if (mk === data.monthKey && data.expenses) {
    var v = {};
    data.expenses.forEach(function(c) { v[c.id] = c.value || 0; });
    return v;
  }
  if (mk === data.monthKey && data.expenseCategories) {
    var v2 = {};
    data.expenseCategories.forEach(function(c) { v2[c.id] = c.value || 0; });
    return v2;
  }
  return {};
}

function collectExpenseCategories(data) {
  var map = {};
  (data.expenseCategories || []).forEach(function(c) {
    map[c.id] = c.name;
  });
  (data.expenseHistory || []).forEach(function(h) {
    var names = h.valueNames || {};
    Object.keys(h.values || {}).forEach(function(id) {
      if (!map[id]) map[id] = names[id] || id;
    });
  });
  return Object.keys(map).map(function(id) { return { id: id, name: map[id] }; });
}

function ledgerForMonth(data, mk) {
  var ledger = data.savingsLedger || [];
  for (var i = 0; i < ledger.length; i++) {
    if (ledger[i].monthKey === mk) return ledger[i];
  }
  return null;
}

function rebuildMonthlySheet(data) {
  var ss = getSs();
  cleanupOldMonthSheets(ss);

  var months = collectMonthKeys(data);
  if (!months.length && data.monthKey) months = [data.monthKey];

  var sources = (data.incomeSources && data.incomeSources.length)
    ? data.incomeSources
    : [{ id: 'src_salary', name: 'Зарплата (основная)' }, { id: 'src_side', name: 'Подработка' }];
  var expenseCats = collectExpenseCategories(data);
  var goals = data.goals || [];

  var rows = [];
  var sectionRows = [];

  rows.push(['Статья'].concat(months.map(monthKeyToLabelGs)));
  rows.push(['_monthKey'].concat(months));

  sectionRows.push(rows.length);
  rows.push(['▸ ДОХОДЫ'].concat(months.map(function() { return ''; })));

  sources.forEach(function(src) {
    var row = [src.name];
    months.forEach(function(mk) {
      row.push(incomeValueForMonth(data, mk, src.id));
    });
    rows.push(row);
  });

  var incTotal = ['ИТОГО доходов'];
  months.forEach(function(mk) {
    var s = 0;
    sources.forEach(function(src) { s += incomeValueForMonth(data, mk, src.id); });
    incTotal.push(s);
  });
  rows.push(incTotal);

  rows.push([''].concat(months.map(function() { return ''; })));

  sectionRows.push(rows.length);
  rows.push(['▸ РАСХОДЫ'].concat(months.map(function() { return ''; })));

  expenseCats.forEach(function(cat) {
    var row = [cat.name];
    months.forEach(function(mk) {
      var vals = expenseValuesForMonth(data, mk);
      row.push(Number(vals[cat.id]) || 0);
    });
    rows.push(row);
  });

  var expTotal = ['ИТОГО расходов'];
  months.forEach(function(mk) {
    var vals = expenseValuesForMonth(data, mk);
    var s = 0;
    expenseCats.forEach(function(cat) { s += Number(vals[cat.id]) || 0; });
    expTotal.push(s);
  });
  rows.push(expTotal);

  rows.push([''].concat(months.map(function() { return ''; })));

  var freeRow = ['Свободный остаток'];
  months.forEach(function(mk, idx) {
    freeRow.push((incTotal[idx + 1] || 0) - (expTotal[idx + 1] || 0));
  });
  rows.push(freeRow);

  if (goals.length) {
    rows.push([''].concat(months.map(function() { return ''; })));
    sectionRows.push(rows.length);
    rows.push(['▸ ОТЛОЖЕНИЯ'].concat(months.map(function() { return ''; })));
    goals.forEach(function(g) {
      var row = [g.name];
      months.forEach(function(mk) {
        var led = ledgerForMonth(data, mk);
        row.push(led && led.allocations ? (Number(led.allocations[g.id]) || 0) : 0);
      });
      rows.push(row);
    });
    var savTotal = ['ИТОГО отложено'];
    months.forEach(function(mk) {
      var led = ledgerForMonth(data, mk);
      var s = 0;
      if (led && led.allocations) {
        goals.forEach(function(g) { s += Number(led.allocations[g.id]) || 0; });
      }
      savTotal.push(s);
    });
    rows.push(savTotal);
  }

  var sh = ensureSheet(SHEETS.MONTHLY, null);
  sh.clear();
  sh.clearFormats();

  var numRows = rows.length;
  var numCols = rows[0].length;
  sh.getRange(1, 1, numRows, numCols).setValues(rows);
  styleMonthlySheet(sh, numRows, numCols, sectionRows);

  sh.hideRows(2);

  return { sheet: SHEETS.MONTHLY, months: months.length };
}

function styleMonthlySheet(sh, numRows, numCols, sectionRows) {
  sh.setFrozenRows(2);
  sh.setFrozenColumns(1);

  var header = sh.getRange(1, 1, 1, numCols);
  header.setFontWeight('bold').setBackground('#1a73e8').setFontColor('#ffffff')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  sh.getRange(1, 1).setBackground('#174ea6').setHorizontalAlignment('left');

  sectionRows.forEach(function(r) {
    sh.getRange(r, 1, 1, numCols).setFontWeight('bold').setBackground('#e8f0fe')
      .setFontColor('#174ea6');
  });

  if (numCols > 1 && numRows >= 3) {
    sh.getRange(3, 2, numRows - 2, numCols - 1).setNumberFormat('#,##0" ₽"');
  }

  sh.autoResizeColumn(1);
  for (var c = 2; c <= numCols; c++) {
    sh.setColumnWidth(c, 108);
  }

  // Итоговые строки — лёгкий фон
  for (var r = 3; r <= numRows; r++) {
    var label = String(sh.getRange(r, 1).getValue() || '');
    if (label.indexOf('ИТОГО') === 0) {
      sh.getRange(r, 1, 1, numCols).setFontWeight('bold').setBackground('#fef7e0');
    }
    if (label === 'Свободный остаток') {
      sh.getRange(r, 1, 1, numCols).setFontWeight('bold').setBackground('#e6f4ea');
    }
  }

  sh.getRange(1, 1, numRows, numCols).setBorder(true, true, true, true, true, true, '#dadce0', SpreadsheetApp.BorderStyle.SOLID);
}

function cleanupOldMonthSheets(ss) {
  var toDelete = [];
  ss.getSheets().forEach(function(sh) {
    var n = sh.getName();
    if (/^\d{4}-\d{2}$/.test(n) || n === 'Доходы' || n === 'Расходы_текущие' || n === 'Расходы_история') {
      toDelete.push(sh);
    }
  });
  toDelete.forEach(function(sh) { ss.deleteSheet(sh); });
}

function exportData() {
  const ss = getSs();
  const result = { version: 3, exportedAt: new Date().toISOString(), sheets: {} };

  const monthlySh = ss.getSheetByName(SHEETS.MONTHLY);
  if (monthlySh) {
    result.sheets.monthly = sheetToObjects(monthlySh);
  }

  const goalsSh = ss.getSheetByName(SHEETS.GOALS);
  if (goalsSh) {
    result.sheets.goals = sheetToObjects(goalsSh);
  }

  const metaSh = ss.getSheetByName(SHEETS.META);
  if (metaSh && metaSh.getLastRow() > 0) {
    const raw = metaSh.getRange(1, 1, metaSh.getLastRow(), 2).getValues();
    result.meta = Object.fromEntries(raw.filter(r => r[0]).map(r => [r[0], r[1]]));
  }

  return result;
}

function importData(data) {
  if (!data) throw new Error('No data');

  rebuildMonthlySheet(data);

  if (data.goals) {
    const sh = ensureSheet(SHEETS.GOALS, [
      'Название', 'Сумма_цели', 'Накоплено', 'Приоритет', 'Режим', 'Значение', 'Цвет'
    ]);
    sh.clearContents();
    sh.getRange(1, 1, 1, 7).setValues([[
      'Название', 'Сумма_цели', 'Накоплено', 'Приоритет', 'Режим', 'Значение', 'Цвет'
    ]]);
    const rows = data.goals.map(g => [
      g.name,
      g.target || 0,
      g.saved || 0,
      g.priority || 1,
      g.planMode || 'monthly',
      g.planMode === 'date' ? (g.targetDate || '') : (g.monthlyAmount || 0),
      g.color || '',
    ]);
    writeRows(sh, 2, rows, 7);
  }

  const metaSh = ensureSheet(SHEETS.META, ['Ключ', 'Значение']);
  metaSh.clearContents();
  metaSh.getRange(1, 1, 1, 2).setValues([['Ключ', 'Значение']]);
  const metaRows = [
    ['version', 2],
    ['syncedAt', new Date().toISOString()],
    ['allocationMode', (data.settings && data.settings.allocationMode) || 'priority'],
    ['savePercent', (data.settings && data.settings.savePercent) || 30],
    ['incomeSources', JSON.stringify(data.incomeSources || [])],
  ];
  writeRows(metaSh, 2, metaRows, 2);
}

/** Обновляет единый лист «По_месяцам» (колонка = месяц, строка = статья). */
function syncMonthSheet(data) {
  if (!data || !data.monthKey) throw new Error('monthKey required');
  var result = rebuildMonthlySheet(data);
  var label = data.label || monthKeyToLabelGs(data.monthKey);
  return {
    ok: true,
    sheet: result.sheet,
    message: 'Сохранено в «' + result.sheet + '»: ' + label + ' (' + result.months + ' мес.)',
  };
}

function sheetToObjects(sh) {
  const values = sh.getDataRange().getValues();
  if (values.length < 2) return [];
  const headers = values[0];
  return values.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => { obj[h] = row[i]; });
    return obj;
  });
}
