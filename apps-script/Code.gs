// =============================================
// АКАДЕМИЯ ЛИЧНОСТИ — Apps Script Backend
// =============================================

var SPREADSHEET_ID = '1iQiTU6wn7TwtvLDrsje4syicDItaSJIt9VlA1c1eCl0';
var NOTIFY_EMAIL   = 'zotmn87@gmail.com';

// ---------- WEB APP ----------

function doGet() {
  return HtmlService.createHtmlOutputFromFile('form')
    .setTitle('Анкета эксперта — Академия Личности')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ---------- FORM SUBMISSION ----------

function submitExpertForm(data) {
  var id;
  try {
    id = writeToSheets(data);
  } catch (e) {
    return { success: false, error: 'Ошибка записи в таблицу: ' + e.toString() };
  }

  var pdf = null;
  try {
    pdf = generatePDF(data);
  } catch (e) {
    Logger.log('PDF error: ' + e.toString());
  }

  try {
    sendEmail(data, pdf);
  } catch (e) {
    Logger.log('Email error: ' + e.toString());
  }

  return { success: true, id: id };
}

// ---------- TEST FUNCTION (run once in editor to authorize all permissions) ----------

function testAuth() {
  var testData = {
    name: 'Тест Авторизации', city: 'Москва', telegram: '@test',
    phone: '', email: '', direction: 'Тест', years_practice: '3–5 лет',
    education: '', teaching_exp: '', methodology: 'Тест',
    has_product: 'Да, живой формат (группы, потоки)',
    program_name: 'Тестовая программа', hours: '40', format: '', price: '20000',
    cohorts_count: '1', students_total: '10', student_result: 'Тест',
    reviews: '', other_programs: '',
    traffic_sources: ['Рекомендации / сарафан'],
    aud_telegram: '500', aud_instagram: '0', aud_vk: '0',
    aud_email: '', aud_other: '', new_clients_month: '4–10', active_base: 'Да, 20–50 человек',
    income: '150 000 – 300 000 руб.', desired_document: 'Сертификат',
    investment_ready: '', time_per_week: '1–3 часа (только вести курсы)',
    partnership_goals: '', objections: '', exclusivity: '', source: 'Тест'
  };
  var pdf = generatePDF(testData);
  Logger.log('PDF OK: ' + pdf.getName());
  MailApp.sendEmail({
    to: NOTIFY_EMAIL,
    subject: '[ТЕСТ] Проверка авторизации Apps Script',
    body: 'Тест успешен. PDF прикреплён.',
    attachments: [pdf]
  });
  Logger.log('Email sent OK');
}

// ---------- WRITE TO SHEETS ----------

function writeToSheets(data) {
  var ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName('Эксперты');
  var id    = sheet.getLastRow() - 1;

  var sources = Array.isArray(data.traffic_sources)
    ? data.traffic_sources.join(', ')
    : (data.traffic_sources || '');

  var audience = [
    data.aud_telegram  ? 'TG: ' + data.aud_telegram  : '',
    data.aud_instagram ? 'IG: ' + data.aud_instagram : '',
    data.aud_vk        ? 'VK: ' + data.aud_vk        : '',
    data.aud_email     ? 'Email: ' + data.aud_email   : '',
    data.aud_other     ? 'Др: ' + data.aud_other      : ''
  ].filter(Boolean).join(' | ');

  var row = [
    id,
    data.name         || '',
    data.telegram     || '',
    data.phone        || '',
    data.email        || '',
    data.city         || '',
    data.direction    || '',
    'Анкета (онлайн)',
    new Date().toLocaleDateString('ru-RU'),
    'Анкета получена',
    'Алексей',
    '',
    'Изучить анкету и принять решение',
    '',
    '',
    data.years_practice   || '',
    data.has_product      || '',
    audience,
    data.income           || '',
    (function(s){ return s==='green'?'Зелёный':s==='yellow'?'Жёлтый':'Красный'; })(autoScore(data)),
    '',
    data.program_name     || '',
    data.hours            || '',
    data.price            || '',
    [
      'Правовой статус: ' + (data.legal_status || '—'),
      'Методология: ' + (data.methodology || '—'),
      'Образование: ' + (data.education || '—'),
      'Результат студента: ' + (data.student_result || '—'),
      'Откуда узнал: ' + (data.source || '—')
    ].join(' | ')
  ];

  sheet.getRange(sheet.getLastRow() + 1, 1, 1, row.length).setValues([row]);
  return id;
}

// ---------- GENERATE PDF ----------

function generatePDF(data) {
  var title = 'Анкета_' + (data.name || 'эксперт') + '_' + new Date().toLocaleDateString('ru-RU');
  var doc   = DocumentApp.create(title);
  var body  = doc.getBody();
  body.setPageWidth(595).setPageHeight(842)
    .setMarginTop(40).setMarginBottom(40)
    .setMarginLeft(50).setMarginRight(50);

  function h(text, level) {
    var p = body.appendParagraph(text);
    p.setHeading(level || DocumentApp.ParagraphHeading.HEADING2);
    return p;
  }
  function row(label, value) {
    body.appendParagraph(label + ': ' + (value || '—'));
  }

  h('АКАДЕМИЯ ЛИЧНОСТИ', DocumentApp.ParagraphHeading.HEADING1);
  h('Анкета предварительного отбора эксперта');
  body.appendParagraph('Дата заполнения: ' + new Date().toLocaleDateString('ru-RU'));
  body.appendParagraph('');

  h('БЛОК 1 — О вас и вашей экспертизе', DocumentApp.ParagraphHeading.HEADING3);
  row('Правовой статус', data.legal_status);
  row('Имя', data.name);
  row('Город', data.city);
  row('Telegram', data.telegram);
  row('Телефон', data.phone);
  row('Email', data.email);
  row('Направление', data.direction);
  row('Опыт практики', data.years_practice);
  row('Образование', data.education);
  row('Педагогический опыт', data.teaching_exp);
  row('Методология', data.methodology);
  body.appendParagraph('');

  h('БЛОК 2 — Ваш продукт', DocumentApp.ParagraphHeading.HEADING3);
  row('Есть курс/программа', data.has_product);
  row('Название программы', data.program_name);
  row('Академических часов', data.hours);
  row('Формат', data.format);
  row('Цена за поток/чел.', data.price);
  row('Потоков проведено', data.cohorts_count);
  row('Студентов суммарно', data.students_total);
  row('Результат студента', data.student_result);
  row('Отзывы/кейсы', data.reviews);
  body.appendParagraph('');

  h('БЛОК 3 — Аудитория', DocumentApp.ParagraphHeading.HEADING3);
  row('Источники клиентов', Array.isArray(data.traffic_sources) ? data.traffic_sources.join(', ') : data.traffic_sources);
  row('Telegram', data.aud_telegram);
  row('Instagram', data.aud_instagram);
  row('VK', data.aud_vk);
  row('Email-база', data.aud_email);
  row('Новых клиентов/мес', data.new_clients_month);
  row('Активная клиентская база', data.active_base);
  body.appendParagraph('');

  h('БЛОК 4 — Финансы и готовность', DocumentApp.ParagraphHeading.HEADING3);
  row('Доход от экспертизы/мес', data.income);
  row('Желаемый документ для студентов', data.desired_document);
  row('Готовность инвестировать', data.investment_ready);
  row('Время в неделю для Академии', data.time_per_week);
  body.appendParagraph('');

  h('БЛОК 5 — Ожидания от партнёрства', DocumentApp.ParagraphHeading.HEADING3);
  row('Главное в партнёрстве', data.partnership_goals);
  row('Что могло бы остановить', data.objections);
  row('Параллельная работа с другими', data.exclusivity);
  row('Откуда узнали об Академии', data.source);

  doc.saveAndClose();

  var pdf = DriveApp.getFileById(doc.getId()).getAs('application/pdf');
  pdf.setName(title + '.pdf');
  DriveApp.getFileById(doc.getId()).setTrashed(true);
  return pdf;
}

// ---------- SEND EMAIL ----------

function sendEmail(data, pdf) {
  var score = autoScore(data);
  var scoreLabel = score === 'green'  ? '🟢 Зелёный коридор — встреча в приоритете'
                 : score === 'yellow' ? '🟡 Условно — нужен уточняющий звонок'
                 : '🔴 Стоп-сигнал — не соответствует критериям';

  var emailConfig = {
    to: NOTIFY_EMAIL,
    subject: '📋 Анкета эксперта: ' + (data.name || '—') + ' | ' + scoreLabel,
    body: [
      'Новая анкета получена от ' + (data.name || '—'),
      '',
      'Направление: ' + (data.direction || '—'),
      'Telegram: ' + (data.telegram || '—'),
      'Город: ' + (data.city || '—'),
      'Доход/мес: ' + (data.income || '—'),
      'Аудитория: ' + (data.aud_telegram || '0') + ' TG / ' + (data.aud_instagram || '0') + ' IG',
      '',
      '─────────────',
      scoreLabel,
      '─────────────',
      '',
      'Данные автоматически внесены в таблицу Эксперты.',
      pdf ? 'Анкета в PDF — во вложении.' : '⚠️ PDF не сформирован — данные в таблице.'
    ].join('\n')
  };

  if (pdf) emailConfig.attachments = [pdf];
  MailApp.sendEmail(emailConfig);
}

// ---------- AUTO-SCORE ----------

function autoScore(data) {
  var stopCount = 0;
  if (data.years_practice === 'Менее 3 лет') stopCount++;
  if (data.has_product === 'Нет, хочу создать с нуля') stopCount++;
  var base = data.active_base || '';
  if (base === 'Нет / только начинаю' || base === 'Да, до 20 человек') stopCount++;

  if (stopCount >= 2) return 'red';

  var green = 0;
  if (data.years_practice === 'Более 10 лет' || data.years_practice === '5–10 лет') green++;
  if (data.has_product === 'Да, живой формат (группы, потоки)' || data.has_product === 'Да, записанный курс (видео, уроки)') green++;
  var totalAud = parseInt(data.aud_telegram || 0) + parseInt(data.aud_instagram || 0) + parseInt(data.aud_vk || 0);
  if (totalAud >= 300) green++;
  if (data.income === '150 000 – 300 000 руб.' || data.income === '300 000 – 500 000 руб.' || data.income === 'Более 500 000 руб.') green++;

  return green >= 3 ? 'green' : 'yellow';
}

// ---------- LEGACY API (direct POST to sheets) ----------

function doPost(e) {
  try {
    var data      = JSON.parse(e.postData.contents);
    var sheetName = data.sheet_name || 'Эксперты';
    var sheet     = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    var row       = data.row_data.map(function(v){ return v || ''; });
    var lastRow   = sheet.getLastRow() + 1;
    sheet.getRange(lastRow, 1, 1, row.length).setValues([row]);
    return ContentService
      .createTextOutput(JSON.stringify({ success: true, row: lastRow }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch(err) {
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
