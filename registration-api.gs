/**
 * Обработчик регистраций для квиза «Колесо баланса»
 *
 * Установка:
 * 1. Создайте Google Таблицу → Расширения → Apps Script
 * 2. Вставьте этот код
 * 3. Запустите setup() один раз (разрешите доступ к Gmail и Таблице)
 * 4. В setup() укажите EMAIL, VK_GROUP_ID, VK_TOKEN (ключ сообщества)
 * 5. Развернуть → Новое развёртывание → Веб-приложение → Доступ: «Все»
 * 6. Скопируйте URL в админку теста (вкладка «Уведомления»)
 */

const CONFIG = {
  API_KEY: 'wQz8kP2mN7xR4vL9sYtK6hJ3fNcU5dAe1bM0',
  EMAIL: 'your@email.com',
  VK_GROUP_ID: '123456789',
  VK_TOKEN: 'vk1.a.xxxx',
  SHEET_NAME: 'Регистрации'
};

function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEET_NAME);
    sheet.appendRow(['Дата', 'Имя', 'Телефон', 'Email', 'Результат', 'Название результата', 'Счёт A', 'Счёт B', 'Счёт V', 'Счёт G']);
    sheet.getRange(1, 1, 1, 10).setFontWeight('bold');
  }
  Logger.log('Готово. URL веб-приложения — в меню Развернуть.');
}

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    if (data.apiKey !== CONFIG.API_KEY) {
      return jsonResponse({ ok: false, error: 'Неверный ключ' }, 403);
    }
    const reg = {
      date: data.date || new Date().toISOString(),
      name: data.name || '',
      phone: data.phone || '',
      email: data.email || '',
      result: data.result || '',
      resultTitle: data.resultTitle || '',
      scores: data.scores || {}
    };
    saveToSheet(reg);
    sendEmail(reg);
    postToVK(reg);
    return jsonResponse({ ok: true });
  } catch (err) {
    return jsonResponse({ ok: false, error: String(err) }, 500);
  }
}

function doGet(e) {
  try {
    const params = e.parameter;
    if (params.apiKey !== CONFIG.API_KEY) {
      return jsonResponse({ ok: false, error: 'Неверный ключ' }, 403);
    }
    if (params.action === 'list') {
      return jsonResponse({ ok: true, registrations: getFromSheet() });
    }
    return jsonResponse({ ok: true, message: 'API работает' });
  } catch (err) {
    return jsonResponse({ ok: false, error: String(err) }, 500);
  }
}

function saveToSheet(reg) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) setup();
  sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  const s = reg.scores;
  sheet.appendRow([
    reg.date,
    reg.name,
    reg.phone,
    reg.email,
    reg.result,
    reg.resultTitle,
    s.A || 0,
    s.B || 0,
    s.V || 0,
    s.G || 0
  ]);
}

function getFromSheet() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 2) return [];
    const rows = sheet.getRange(2, 1, sheet.getLastRow(), 10).getValues();
  return rows.map(r => ({
    date: r[0] instanceof Date ? r[0].toISOString() : String(r[0]),
    name: String(r[1]),
    phone: String(r[2]),
    email: String(r[3]),
    result: String(r[4]),
    resultTitle: String(r[5]),
    scores: { A: r[6], B: r[7], V: r[8], G: r[9] }
  })).reverse();
}

function formatMessage(reg) {
  return [
    '🆕 Новая регистрация в квизе',
    '',
    '👤 Имя: ' + reg.name,
    '📞 Телефон: ' + reg.phone,
    '📧 Email: ' + reg.email,
    '🎯 Результат: ' + reg.result + ' — ' + reg.resultTitle,
    '📊 Счёт: A=' + (reg.scores.A || 0) + ' B=' + (reg.scores.B || 0) + ' V=' + (reg.scores.V || 0) + ' G=' + (reg.scores.G || 0),
    '🕐 ' + reg.date
  ].join('\n');
}

function sendEmail(reg) {
  if (!CONFIG.EMAIL || CONFIG.EMAIL === 'your@email.com') return;
  GmailApp.sendEmail(
    CONFIG.EMAIL,
    'Новая регистрация: ' + reg.name + ' (' + reg.result + ')',
    formatMessage(reg)
  );
}

function postToVK(reg) {
  if (!CONFIG.VK_TOKEN || CONFIG.VK_TOKEN === 'vk1.a.xxxx') return;
  if (!CONFIG.VK_GROUP_ID || CONFIG.VK_GROUP_ID === '123456789') return;
  const message = formatMessage(reg);
  const url = 'https://api.vk.com/method/wall.post';
  const payload = {
    owner_id: '-' + CONFIG.VK_GROUP_ID,
    from_group: '1',
    message: message,
    access_token: CONFIG.VK_TOKEN,
    v: '5.131'
  };
  UrlFetchApp.fetch(url, { method: 'post', payload: payload, muteHttpExceptions: true });
}

function jsonResponse(obj, code) {
  const output = ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
  return output;
}
