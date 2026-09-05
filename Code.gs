/**
 * Контроль школьных оценок: Google Таблица + Telegram-бот.
 * Весь код в одном файле — вставляется в Apps Script одним куском.
 *
 * Разметка листа предмета взята из существующей таблицы и НЕ переделывается:
 *   A1 «Предмет:», B1 — название
 *   строка 3 — заголовки
 *   строка 4 «Год», строка 5 «1 Четверть» со средним баллом
 *   оценки — с 6-й строки
 * Скрипт только дописывает справа две служебные колонки: «Тема» и «Статус».
 */

const SHEET_SETTINGS = 'Настройки';
const SHEET_LIST = 'Список';
const SHEET_STATE = '_Состояния';

// Колонки листа предмета. A–E были в таблице изначально, F и G добавляет скрипт.
const COL = { DATE: 1, GRADE: 2, REASON: 3, RETAKE: 4, RESULT: 5, TOPIC: 6, STATUS: 7 };
const HEADER_ROW = 3;   // строка с заголовками
const DATA_ROW = 6;     // первая строка с оценками (строки 4–5 заняты «Год» и «1 Четверть»)
const QUARTER_LAST_ROW = 20; // конец блока 1 четверти — его же охватывает формула среднего

// Статусы ставит только скрипт: по ним он понимает, чего ждёт от следующего сообщения.
const ST = {
  WAIT_REASON: 'Ждём причину',
  WAIT_DATE: 'Ждём дату пересдачи',
  PLANNED: 'Пересдача назначена',
  WAIT_RESULT: 'Ждём результат',
  DONE: 'Закрыто',
  ESCALATED: 'Эскалация родителям'
};

// Роли для самостоятельной регистрации: кнопка → пара полей в «Настройках».
const ROLES = {
  child: { id: 'CHILD_USER_ID', name: 'CHILD_USERNAME', label: 'ребёнок' },
  father: { id: 'PARENT1_USER_ID', name: 'PARENT1_USERNAME', label: 'отец' },
  mother: { id: 'PARENT2_USER_ID', name: 'PARENT2_USERNAME', label: 'мать' }
};

// Значения по умолчанию для листа «Настройки»: [ключ, значение, пояснение].
const DEFAULT_SETTINGS = [
  ['BOT_TOKEN', '', 'Токен от @BotFather. После пункта меню «Сохранить секреты» здесь останется маска.'],
  ['WEBAPP_URL', '', 'URL веб-приложения после развёртывания. Нужен для подключения Telegram.'],
  ['GROUP_CHAT_ID', '', 'Заполнится само: отправьте в группе команду /group.'],
  ['CHILD_USER_ID', '', 'Заполнится само: в группе команда /family, ребёнок жмёт «Я ребёнок».'],
  ['CHILD_USERNAME', '', 'Заполнится само вместе с ID.'],
  ['PARENT1_USER_ID', '', 'Заполнится само: отец жмёт «Я отец».'],
  ['PARENT1_USERNAME', '', 'Заполнится само вместе с ID.'],
  ['PARENT2_USER_ID', '', 'Заполнится само: мать жмёт «Я мать».'],
  ['PARENT2_USERNAME', '', 'Заполнится само вместе с ID.'],
  ['MIN_GRADE', '4', 'Оценка НИЖЕ этого числа считается проблемной.'],
  ['ESCALATE_DAYS', '3', 'Через столько дней без даты пересдачи бот зовёт родителей.'],
  ['MORNING_HOUR', '8', 'Час утреннего напоминания.'],
  ['EVENING_HOUR', '19', 'Час вечернего вопроса «сдала?».'],
  ['SHEET_URL', '', 'Ссылка на эту таблицу — бот прикладывает её к сообщениям. Заполняется сама.']
];

// Кэш на время одного запуска: лист читаем один раз, а не на каждый ключ.
let _settingsCache = null;

function getSettings_() {
  if (_settingsCache) return _settingsCache;
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_SETTINGS);
  if (!sh) throw new Error('Нет листа «Настройки». Запустите меню «Оценки → 1. Настроить таблицу».');
  const values = sh.getRange(2, 1, Math.max(sh.getLastRow() - 1, 1), 2).getValues();
  _settingsCache = {};
  values.forEach(function (r) {
    if (r[0]) _settingsCache[String(r[0]).trim()] = String(r[1]).trim();
  });
  return _settingsCache;
}

function getSetting_(key, fallback) {
  const v = getSettings_()[key];
  return (v === undefined || v === '') ? fallback : v;
}

function getNumSetting_(key, fallback) {
  const n = Number(getSetting_(key, fallback));
  return isNaN(n) ? fallback : n;
}

/** Запись в «Настройки» из кода. Кэш сбрасываем, иначе в этом же запуске прочитаем старое. */
function setSetting_(key, value) {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_SETTINGS);
  const row = findSettingRow_(sh, key);
  if (!row) return false;
  sh.getRange(row, 2).setValue(value);
  _settingsCache = null;
  return true;
}

function getToken_() {
  const p = PropertiesService.getScriptProperties().getProperty('BOT_TOKEN');
  if (p) return p;
  const fromSheet = getSetting_('BOT_TOKEN', '');
  if (!fromSheet || fromSheet.indexOf('•') === 0) {
    throw new Error('Токен бота не задан. Заполните BOT_TOKEN и нажмите «Сохранить секреты».');
  }
  return fromSheet;
}

function tz_() {
  return SpreadsheetApp.getActive().getSpreadsheetTimeZone();
}

function isServiceSheet_(name) {
  return [SHEET_SETTINGS, SHEET_LIST, SHEET_STATE].indexOf(name) !== -1;
}

/* ==================== МЕНЮ И НАСТРОЙКА ==================== */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Оценки')
    .addItem('1. Настроить таблицу', 'setupWorkbook')
    .addItem('2. Сохранить секреты', 'saveSecrets')
    .addItem('3. Подключить Telegram', 'setWebhook')
    .addItem('4. Включить триггеры', 'installTriggers')
    .addItem('5. Проверить связь', 'testConnection')
    .addSeparator()
    .addItem('Отключить бота (аварийно)', 'stopBot')
    .addItem('Состояние бота', 'showWebhookInfo')
    .addSeparator()
    .addItem('Кто зарегистрирован', 'showRegistered')
    .addItem('Сбросить регистрацию', 'resetRegistration')
    .addItem('Обновить ссылки на листах', 'buildLinks')
    .addToUi();
}

function setupWorkbook() {
  const ss = SpreadsheetApp.getActive();
  ensureSettingsSheet_(ss);
  ensureStateSheet_(ss);
  ensureSubjectSheets_(ss);
  buildLinks();
  SpreadsheetApp.getUi().alert('Готово. Заполните в «Настройках» только BOT_TOKEN — остальное бот заполнит сам.');
}

function ensureSettingsSheet_(ss) {
  let sh = ss.getSheetByName(SHEET_SETTINGS);
  if (!sh) sh = ss.insertSheet(SHEET_SETTINGS, ss.getNumSheets());
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, 3).setValues([['Параметр', 'Значение', 'Пояснение']]).setFontWeight('bold');
    sh.getRange(2, 1, DEFAULT_SETTINGS.length, 3).setValues(DEFAULT_SETTINGS);
    sh.setColumnWidth(1, 180);
    sh.setColumnWidth(2, 260);
    sh.setColumnWidth(3, 520);
    sh.setFrozenRows(1);
  }
  const row = findSettingRow_(sh, 'SHEET_URL');
  if (row && !sh.getRange(row, 2).getValue()) sh.getRange(row, 2).setValue(ss.getUrl());
  return sh;
}

function findSettingRow_(sh, key) {
  const keys = sh.getRange(2, 1, Math.max(sh.getLastRow() - 1, 1), 1).getValues();
  for (let i = 0; i < keys.length; i++) if (String(keys[i][0]).trim() === key) return i + 2;
  return 0;
}

function ensureStateSheet_(ss) {
  let sh = ss.getSheetByName(SHEET_STATE);
  if (!sh) {
    sh = ss.insertSheet(SHEET_STATE, ss.getNumSheets());
    sh.getRange(1, 1, 1, 5).setValues([['Ключ', 'Лист', 'Строка', 'Ожидаем', 'Создано']]).setFontWeight('bold');
    sh.hideSheet(); // служебный лист, глазам не нужен
  }
  return sh;
}

function ensureSubjectSheets_(ss) {
  const list = ss.getSheetByName(SHEET_LIST);
  if (!list) throw new Error('Нет листа «Список» с перечнем предметов.');
  const names = list.getRange(1, 1, list.getLastRow(), 1).getValues()
    .map(function (r) { return String(r[0]).trim(); })
    .filter(function (n) { return n && n.toLowerCase() !== 'список'; });

  names.forEach(function (name) {
    let sh = ss.getSheetByName(name);
    if (!sh) { sh = ss.insertSheet(name); buildSubjectTemplate_(sh, name); }
    ensureServiceColumns_(sh);
  });
}

/** Новый лист собираем по образцу существующих, чтобы все предметы выглядели одинаково. */
function buildSubjectTemplate_(sh, name) {
  sh.getRange('A1').setValue('Предмет:').setFontWeight('bold');
  sh.getRange('B1').setValue(name).setFontSize(18).setFontWeight('bold');
  sh.getRange(HEADER_ROW, 1, 1, 5)
    .setValues([['Даты', 'Оценки', 'Причина', 'Дата пересдачи', 'Результат']])
    .setFontWeight('bold').setHorizontalAlignment('center');
  sh.getRange('A4').setValue('Год').setFontWeight('bold');
  sh.getRange('A5').setValue('1 Четверть').setFontWeight('bold');
  sh.getRange('B5').setFormula('=IFERROR(AVERAGE(B' + DATA_ROW + ':B' + QUARTER_LAST_ROW + '); 0)');
  sh.setColumnWidth(COL.REASON, 320);
}

/** Дописывает «Тему» и «Статус» справа, если их ещё нет. Существующие колонки не трогает. */
function ensureServiceColumns_(sh) {
  if (String(sh.getRange(HEADER_ROW, COL.TOPIC).getValue()).trim() !== 'Тема') {
    sh.getRange(HEADER_ROW, COL.TOPIC).setValue('Тема').setFontWeight('bold').setHorizontalAlignment('center');
  }
  if (String(sh.getRange(HEADER_ROW, COL.STATUS).getValue()).trim() !== 'Статус') {
    sh.getRange(HEADER_ROW, COL.STATUS).setValue('Статус').setFontWeight('bold').setHorizontalAlignment('center');
  }
  sh.setColumnWidth(COL.TOPIC, 220);
  sh.setColumnWidth(COL.STATUS, 170);
}

/**
 * Ссылки: со «Списка» на предмет и обратно.
 * Через gid, а не через имя: лист можно переименовать, gid остаётся прежним.
 * Обратную ссылку кладём в I1 — A1 и B1 заняты названием предмета.
 */
function buildLinks() {
  const ss = SpreadsheetApp.getActive();
  const list = ss.getSheetByName(SHEET_LIST);
  const rows = list.getRange(1, 1, list.getLastRow(), 1).getValues();

  rows.forEach(function (r, i) {
    const name = String(r[0]).trim();
    if (!name) return;
    const sh = ss.getSheetByName(name);
    if (!sh) return;
    list.getRange(i + 1, 2).setFormula('=HYPERLINK("#gid=' + sh.getSheetId() + '";"открыть →")');
    sh.getRange('I1').setFormula('=HYPERLINK("#gid=' + list.getSheetId() + '";"← к списку предметов")');
  });
  list.setColumnWidth(2, 120);
}

function saveSecrets() {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_SETTINGS);
  const row = findSettingRow_(sh, 'BOT_TOKEN');
  const val = String(sh.getRange(row, 2).getValue()).trim();
  if (!val || val.indexOf('•') === 0) {
    SpreadsheetApp.getUi().alert('Токен уже сохранён либо поле пустое.');
    return;
  }
  PropertiesService.getScriptProperties().setProperty('BOT_TOKEN', val);
  sh.getRange(row, 2).setValue('••••• сохранён ' + Utilities.formatDate(new Date(), tz_(), 'dd.MM.yyyy'));
  SpreadsheetApp.getUi().alert('Токен сохранён отдельно от таблицы.');
}

function installTriggers() {
  const ss = SpreadsheetApp.getActive();
  ScriptApp.getProjectTriggers().forEach(function (t) { ScriptApp.deleteTrigger(t); });

  // Устанавливаемый onEdit — в отличие от простого, ему разрешён выход в интернет.
  ScriptApp.newTrigger('onEditInstalled').forSpreadsheet(ss).onEdit().create();
  ScriptApp.newTrigger('morningRoutine').timeBased().atHour(getNumSetting_('MORNING_HOUR', 8)).everyDays(1).create();
  ScriptApp.newTrigger('eveningRoutine').timeBased().atHour(getNumSetting_('EVENING_HOUR', 19)).everyDays(1).create();

  SpreadsheetApp.getUi().alert('Триггеры включены.');
}

function testConnection() {
  const chat = getSetting_('GROUP_CHAT_ID', '');
  if (!chat) {
    SpreadsheetApp.getUi().alert('Группа ещё не привязана. Отправьте в группе команду /group.');
    return;
  }
  const r = tgSend_(chat, 'Проверка связи: бот на месте.' + sheetLink_());
  SpreadsheetApp.getUi().alert(r.ok ? 'Сообщение ушло в группу.' : 'Ошибка: ' + JSON.stringify(r));
}

function showRegistered() {
  let out = '';
  Object.keys(ROLES).forEach(function (k) {
    const r = ROLES[k];
    out += r.label + ': ' + (getSetting_(r.name, '') || '—') + ' ' + (getSetting_(r.id, '') || '(нет ID)') + '\n';
  });
  out += '\nГруппа: ' + (getSetting_('GROUP_CHAT_ID', '') || 'не привязана');
  SpreadsheetApp.getUi().alert(out);
}

function resetRegistration() {
  Object.keys(ROLES).forEach(function (k) {
    setSetting_(ROLES[k].id, '');
    setSetting_(ROLES[k].name, '');
  });
  SpreadsheetApp.getUi().alert('Регистрация сброшена. Отправьте в группе /family и нажмите кнопки заново.');
}

/* ==================== СТРОКИ ЖУРНАЛА ==================== */

function getSheetByGid_(gid) {
  const sheets = SpreadsheetApp.getActive().getSheets();
  for (let i = 0; i < sheets.length; i++) {
    if (String(sheets[i].getSheetId()) === String(gid)) return sheets[i];
  }
  return null;
}

function subjectSheets_() {
  return SpreadsheetApp.getActive().getSheets().filter(function (sh) {
    return !isServiceSheet_(sh.getName()) && !sh.isSheetHidden();
  });
}

function setCell_(sh, row, col, value) { sh.getRange(row, col).setValue(value); }
function getCell_(sh, row, col) { return sh.getRange(row, col).getValue(); }
function setStatus_(sh, row, status) { setCell_(sh, row, COL.STATUS, status); }

/** Оценки начинаются с DATA_ROW: выше — шапка, «Год» и средний балл за четверть. */
function eachRow_(sh, fn) {
  const last = sh.getLastRow();
  if (last < DATA_ROW) return;
  const values = sh.getRange(DATA_ROW, 1, last - DATA_ROW + 1, COL.STATUS).getValues();
  values.forEach(function (v, i) { fn(v, i + DATA_ROW); });
}

function fmtDate_(d) {
  if (!(d instanceof Date)) return String(d || '');
  return Utilities.formatDate(d, tz_(), 'dd.MM.yyyy');
}

function sameDay_(a, b) { return fmtDate_(a) === fmtDate_(b); }

function addDays_(d, n) {
  const x = new Date(d.getTime());
  x.setDate(x.getDate() + n);
  return x;
}

/**
 * Разбор даты: «12.03», «12.03.2026», «12/03».
 * Год подставляем текущий; если дата далеко позади — считаем, что речь о следующем годе,
 * иначе «12.01», написанное в декабре, попало бы в прошлое.
 */
function parseDate_(text) {
  const m = String(text).match(/(\d{1,2})[.\-\/](\d{1,2})(?:[.\-\/](\d{2,4}))?/);
  if (!m) return null;
  const day = Number(m[1]);
  const month = Number(m[2]) - 1;
  const now = new Date();
  let year = m[3] ? Number(m[3]) : now.getFullYear();
  if (year < 100) year += 2000;
  const d = new Date(year, month, day);
  if (isNaN(d.getTime())) return null;
  if (!m[3] && d < addDays_(now, -30)) d.setFullYear(year + 1);
  return d;
}

/* ==================== TELEGRAM ==================== */

function tgCall_(method, payload) {
  const url = 'https://api.telegram.org/bot' + getToken_() + '/' + method;
  const res = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true // иначе ошибка Telegram уронит весь триггер
  });
  try { return JSON.parse(res.getContentText()); }
  catch (e) { return { ok: false, raw: res.getContentText() }; }
}

function tgSend_(chatId, text, extra) {
  const payload = Object.assign({ chat_id: chatId, text: text, parse_mode: 'HTML' }, extra || {});
  return tgCall_('sendMessage', payload);
}

/**
 * Вопрос с принудительным ответом. selective: true — окно ответа поднимается
 * только у того, кого упомянули в тексте, остальных в группе не дёргает.
 * Ответы reply бот видит всегда, даже при включённом privacy mode.
 */
function tgAsk_(chatId, text) {
  return tgSend_(chatId, text, { reply_markup: { force_reply: true, selective: true } });
}

function tgButtons_(chatId, text, buttons) {
  return tgSend_(chatId, text, { reply_markup: { inline_keyboard: buttons } });
}

function tgAnswerCallback_(id, text) {
  return tgCall_('answerCallbackQuery', { callback_query_id: id, text: text || '' });
}

/**
 * drop_pending_updates обязателен: если в очереди Telegram зависли старые
 * сообщения, после подключения они прилетят все разом и бот начнёт отвечать на них.
 * max_connections: 1 — Apps Script всё равно выполняет запросы по одному.
 */
function setWebhook() {
  const url = getSetting_('WEBAPP_URL', '');
  if (!url) throw new Error('Сначала разверните веб-приложение и впишите его URL в WEBAPP_URL.');
  const r = tgCall_('setWebhook', {
    url: url,
    allowed_updates: ['message', 'callback_query'],
    drop_pending_updates: true,
    max_connections: 1
  });
  SpreadsheetApp.getUi().alert(r.ok ? 'Telegram подключён.' : 'Ошибка: ' + JSON.stringify(r));
}

/** Аварийная кнопка: бот замолкает сразу, без правки ячеек и переразвёртывания. */
function stopBot() {
  const r = tgCall_('deleteWebhook', { drop_pending_updates: true });
  SpreadsheetApp.getUi().alert(r.ok
    ? 'Бот отключён, очередь очищена. Включить обратно: «3. Подключить Telegram».'
    : 'Ошибка: ' + JSON.stringify(r));
}

/** Диагностика: Telegram сам говорит, сколько сообщений зависло и что ему не нравится. */
function showWebhookInfo() {
  const r = tgCall_('getWebhookInfo', {});
  const i = r.result || {};
  const msg = 'Адрес: ' + (i.url || 'не задан') +
    '\nВ очереди: ' + (i.pending_update_count === undefined ? '?' : i.pending_update_count) +
    '\nПоследняя ошибка: ' + (i.last_error_message || 'нет');
  SpreadsheetApp.getUi().alert(msg);
}

function sheetLink_() {
  const u = getSetting_('SHEET_URL', '');
  return u ? '\n\n<a href="' + u + '">Таблица оценок</a>' : '';
}

/* ==================== ПРИЁМ СООБЩЕНИЙ ==================== */

/**
 * Зачем защита от повторов: если Telegram не счёл ответ успешным, он шлёт то же
 * самое сообщение снова и снова. Без проверки каждый повтор — новый ответ бота
 * в группу, и чат заваливает одинаковыми сообщениями.
 * Номер обработанного сообщения кладём в кэш на 6 часов и дубли молча пропускаем.
 */
function alreadyHandled_(updateId) {
  if (!updateId) return false;
  const cache = CacheService.getScriptCache();
  const key = 'upd_' + updateId;
  if (cache.get(key)) return true;
  cache.put(key, '1', 21600);
  return false;
}

function doPost(e) {
  // Блокировка нужна на случай, когда повторы прилетают одновременно и оба
  // успевают проверить кэш до записи в него.
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
  } catch (err) {
    return ContentService.createTextOutput('busy');
  }
  try {
    const update = JSON.parse(e.postData.contents);
    if (alreadyHandled_(update.update_id)) return ContentService.createTextOutput('dup');
    if (update.message) handleMessage_(update.message);
    else if (update.callback_query) handleCallback_(update.callback_query);
  } catch (err) {
    console.error(err);
  } finally {
    lock.releaseLock();
  }
  return ContentService.createTextOutput('ok');
}

function handleMessage_(msg) {
  const text = String(msg.text || '').trim();
  const isPrivate = msg.chat.type === 'private';

  // Команды только латиницей: Telegram распознаёт командой лишь латинские слова
  // после слеша, а в группе при включённом privacy mode бот видит только команды
  // и ответы на свои сообщения. Русские варианты оставлены для личного чата.
  if (text.indexOf('/group') === 0 || text.indexOf('/группа') === 0) {
    if (isPrivate) {
      tgSend_(msg.chat.id, 'Эту команду нужно отправить внутри семейной группы.');
    } else {
      setSetting_('GROUP_CHAT_ID', String(msg.chat.id));
      tgSend_(msg.chat.id, 'Группа привязана. Сюда я буду писать про оценки и пересдачи.');
    }
    return;
  }

  // Регистрация ролей прямо в группе: нажатие кнопки приносит боту числовой ID
  // так же, как личное сообщение, — отдельный чат для этого не нужен.
  if (text.indexOf('/family') === 0 || text.indexOf('/семья') === 0 || text.indexOf('/start') === 0) {
    askRole_(msg.chat.id);
    return;
  }

  if (!msg.reply_to_message) return; // обычную болтовню в группе игнорируем
  const key = msg.chat.id + ':' + msg.reply_to_message.message_id;
  const state = loadState_(key);
  if (!state) return;

  const sh = getSheetByGid_(state.gid);
  if (!sh) return;

  if (state.expects === 'reason') {
    saveReason_(sh, state.row, text);
    dropState_(key);
    askRetakeDate_(sh, state.row, msg.chat.id);

  } else if (state.expects === 'date') {
    const d = parseDate_(text);
    if (!d) {
      const again = tgAsk_(msg.chat.id, 'Не понял дату. Напиши в формате 12.03');
      if (again.ok) saveState_(msg.chat.id + ':' + again.result.message_id, state.gid, state.row, 'date');
      return;
    }
    setCell_(sh, state.row, COL.RETAKE, d);
    setStatus_(sh, state.row, ST.PLANNED);
    dropState_(key);
    tgSend_(msg.chat.id, 'Записал: пересдача «' + sh.getName() + '» — ' + fmtDate_(d) + '. Напомню накануне.');
  }
}

/**
 * Числовой ID приходит боту в каждом сообщении сам — спрашивать его у человека
 * или гонять к стороннему боту не нужно. Неизвестна только роль, её и спрашиваем.
 */
function askRole_(chatId) {
  tgButtons_(chatId, 'Кто есть кто? Каждый нажимает свою кнопку — я запомню и больше не спрошу.', [
    [{ text: 'Я мать', callback_data: 'reg|mother' }, { text: 'Я отец', callback_data: 'reg|father' }],
    [{ text: 'Я ребёнок', callback_data: 'reg|child' }]
  ]);
}

function registerRole_(cb, roleKey) {
  const role = ROLES[roleKey];
  if (!role) { tgAnswerCallback_(cb.id, 'Неизвестная роль'); return; }

  const taken = getSetting_(role.id, '');

  // Роль занимается один раз. Иначе любой, кто найдёт бота, объявит себя родителем.
  if (taken && String(taken) !== String(cb.from.id)) {
    tgAnswerCallback_(cb.id, 'Эта роль уже занята');
    return;
  }

  setSetting_(role.id, String(cb.from.id));
  setSetting_(role.name, cb.from.username ? '@' + cb.from.username : (cb.from.first_name || ''));
  tgAnswerCallback_(cb.id, 'Записал: ' + role.label);
}

function handleCallback_(cb) {
  const parts = String(cb.data).split('|');
  const action = parts[0];

  if (action === 'reg') { registerRole_(cb, parts[1]); return; }

  const gid = parts[1];
  const row = Number(parts[2]);
  const sh = getSheetByGid_(gid);
  const chatId = cb.message.chat.id;
  if (!sh) { tgAnswerCallback_(cb.id, 'Лист не найден'); return; }

  const today = new Date();

  if (action === 'd1' || action === 'd3') {
    const d = addDays_(today, action === 'd1' ? 1 : 3);
    setCell_(sh, row, COL.RETAKE, d);
    setStatus_(sh, row, ST.PLANNED);
    tgAnswerCallback_(cb.id, 'Записал');
    tgSend_(chatId, 'Пересдача «' + sh.getName() + '» — ' + fmtDate_(d) + '. Напомню накануне.');

  } else if (action === 'dx') {
    tgAnswerCallback_(cb.id, '');
    const m = tgAsk_(chatId, getSetting_('CHILD_USERNAME', '') + ', напиши дату пересдачи по предмету «' + sh.getName() + '» в формате 12.03');
    if (m.ok) saveState_(chatId + ':' + m.result.message_id, sh.getSheetId(), row, 'date');

  } else if (action === 'd0') {
    setStatus_(sh, row, ST.WAIT_DATE);
    tgAnswerCallback_(cb.id, 'Хорошо');
    tgSend_(chatId, 'Ок. Спроси у учителя и напиши дату. Если через ' + getNumSetting_('ESCALATE_DAYS', 3) + ' дня даты не будет — подключу родителей.');

  } else if (action === 'ok') {
    setCell_(sh, row, COL.RESULT, 'Пересдача прошла — проверить в журнале');
    setStatus_(sh, row, ST.DONE);
    tgAnswerCallback_(cb.id, 'Отлично!');
    tgSend_(chatId, 'Записал: «' + sh.getName() + '» пересдана. Родители проверят оценку в журнале.');

  } else if (action === 'no') {
    setCell_(sh, row, COL.RESULT, 'Не сдала ' + fmtDate_(today));
    setStatus_(sh, row, ST.ESCALATED);
    tgAnswerCallback_(cb.id, 'Понял');
    escalate_(sh, row, 'пересдача не состоялась');

  } else if (action === 'mv') {
    setCell_(sh, row, COL.RETAKE, '');
    tgAnswerCallback_(cb.id, '');
    askRetakeDate_(sh, row, chatId);
  }
}

/**
 * Тема и причина приходят одним сообщением, чтобы не гонять ребёнка по двум вопросам.
 * Первая строка (или часть до «—») считается темой, остальное — причиной.
 */
function saveReason_(sh, row, text) {
  let topic = '';
  let reason = text;
  const parts = text.split(/\n|—|--/);
  if (parts.length > 1) {
    topic = parts[0].trim();
    reason = parts.slice(1).join(' ').trim();
  }
  if (topic) setCell_(sh, row, COL.TOPIC, topic);
  setCell_(sh, row, COL.REASON, reason);
}

function askRetakeDate_(sh, row, chatId) {
  setStatus_(sh, row, ST.WAIT_DATE);
  const tag = sh.getSheetId() + '|' + row;
  tgButtons_(chatId, 'Когда пересдача по предмету «' + sh.getName() + '»?', [
    [{ text: 'Завтра', callback_data: 'd1|' + tag }, { text: 'Через 3 дня', callback_data: 'd3|' + tag }],
    [{ text: 'Ввести дату', callback_data: 'dx|' + tag }, { text: 'Ещё не знаю', callback_data: 'd0|' + tag }]
  ]);
}

/* ---------- состояние диалога ---------- */

function saveState_(key, gid, row, expects) {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_STATE);
  sh.appendRow([key, String(gid), row, expects, new Date()]);
}

function loadState_(key) {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_STATE);
  const last = sh.getLastRow();
  if (last < 2) return null;
  const v = sh.getRange(2, 1, last - 1, 4).getValues();
  for (let i = v.length - 1; i >= 0; i--) {
    if (String(v[i][0]) === key) return { gid: v[i][1], row: Number(v[i][2]), expects: v[i][3] };
  }
  return null;
}

function dropState_(key) {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_STATE);
  const last = sh.getLastRow();
  if (last < 2) return;
  const v = sh.getRange(2, 1, last - 1, 1).getValues();
  for (let i = v.length - 1; i >= 0; i--) {
    if (String(v[i][0]) === key) sh.deleteRow(i + 2);
  }
}

/* ==================== ТРИГГЕРЫ ==================== */

/**
 * Почему onEditInstalled, а не onEdit: простой onEdit запускается без прав
 * пользователя и не может обращаться к внешним сервисам — сообщение в Telegram
 * из него не уйдёт.
 */
function onEditInstalled(e) {
  const sh = e.range.getSheet();
  if (isServiceSheet_(sh.getName())) return;
  if (e.range.getColumn() !== COL.GRADE || e.range.getRow() < DATA_ROW) return;

  const grade = Number(e.range.getValue());
  if (!grade) return;

  const row = e.range.getRow();
  if (!getCell_(sh, row, COL.DATE)) setCell_(sh, row, COL.DATE, new Date());

  if (grade >= getNumSetting_('MIN_GRADE', 4)) {
    setStatus_(sh, row, ST.DONE);
    return;
  }

  setStatus_(sh, row, ST.WAIT_REASON);
  const chatId = getSetting_('GROUP_CHAT_ID', '');
  if (!chatId) return;
  const text = getSetting_('CHILD_USERNAME', '') + ', по предмету <b>' + sh.getName() + '</b> оценка <b>' +
    grade + '</b>.\nОтветь на это сообщение одним текстом:\n<i>тема работы — почему так вышло</i>';
  const m = tgAsk_(chatId, text);
  if (m.ok) saveState_(chatId + ':' + m.result.message_id, sh.getSheetId(), row, 'reason');
}

/** Утро: напоминание о завтрашней пересдаче и подталкивание тех, кто тянет с датой. */
function morningRoutine() {
  const chatId = getSetting_('GROUP_CHAT_ID', '');
  if (!chatId) return;
  const today = new Date();
  const tomorrow = addDays_(today, 1);
  const limit = getNumSetting_('ESCALATE_DAYS', 3);

  subjectSheets_().forEach(function (sh) {
    eachRow_(sh, function (v, row) {
      const status = v[COL.STATUS - 1];
      const retake = v[COL.RETAKE - 1];

      if (status === ST.PLANNED && retake instanceof Date && sameDay_(retake, tomorrow)) {
        tgSend_(chatId, getSetting_('CHILD_USERNAME', '') + ', завтра пересдача по предмету <b>' + sh.getName() +
          '</b>. Тема: ' + (v[COL.TOPIC - 1] || 'не указана') + '.');
      }

      if (status === ST.WAIT_DATE) {
        const since = v[COL.DATE - 1];
        if (since instanceof Date && addDays_(since, limit) <= today) {
          setStatus_(sh, row, ST.ESCALATED);
          escalate_(sh, row, 'уже ' + limit + ' дня нет даты пересдачи');
        }
      }
    });
  });
}

/** Вечер: у кого сегодня была пересдача — спрашиваем результат. */
function eveningRoutine() {
  const chatId = getSetting_('GROUP_CHAT_ID', '');
  if (!chatId) return;
  const today = new Date();

  subjectSheets_().forEach(function (sh) {
    eachRow_(sh, function (v, row) {
      const retake = v[COL.RETAKE - 1];
      if (v[COL.STATUS - 1] !== ST.PLANNED) return;
      if (!(retake instanceof Date) || !sameDay_(retake, today)) return;

      setStatus_(sh, row, ST.WAIT_RESULT);
      const tag = sh.getSheetId() + '|' + row;
      tgButtons_(chatId, getSetting_('CHILD_USERNAME', '') + ', как прошла пересдача по предмету <b>' + sh.getName() + '</b>?', [
        [{ text: 'Сдала', callback_data: 'ok|' + tag }, { text: 'Не сдала', callback_data: 'no|' + tag }],
        [{ text: 'Перенесли', callback_data: 'mv|' + tag }]
      ]);
    });
  });
}

/** Зовём родителей: сначала в группу с упоминанием, затем лично — если известен ID. */
function escalate_(sh, row, why) {
  const p1 = getSetting_('PARENT1_USERNAME', '');
  const p2 = getSetting_('PARENT2_USERNAME', '');
  const text = '⚠️ ' + p1 + ' ' + p2 + '\nПредмет <b>' + sh.getName() + '</b>, оценка ' +
    getCell_(sh, row, COL.GRADE) + ': ' + why + '. Нужно подключиться.' + sheetLink_();

  const group = getSetting_('GROUP_CHAT_ID', '');
  if (group) tgSend_(group, text);
  [getSetting_('PARENT1_USER_ID', ''), getSetting_('PARENT2_USER_ID', '')].forEach(function (id) {
    if (id) tgSend_(id, text);
  });
}
