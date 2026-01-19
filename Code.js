# 360-feedback-slack-bot
Serverless HR Feedback 360 bot for Slack based on Google Apps Script
// ==========================================
// ⚙️ НАЛАШТУВАННЯ (v24.0 ANTI-SPAM FIX)
// ==========================================
const SLACK_TOKEN = 'xoxb-ВАШ_ТОКЕН'; 
const SPREADSHEET_ID = 'ВАШ_ID_ТАБЛИЦІ'; 
const WEB_APP_URL = 'ВАША_URL_ВЕБ_ДОДАТКА'; 

// 👮‍♂️ АДМІНИ
const ADMIN_IDS = ['ВАШ ID']; 

// 🗓 ВІКНО ПОШУКУ (Днів +/- від дати вибраного рядка для звіту)
const PERIOD_WINDOW_DAYS = 30; 

// ⏳ СКІЛЬКИ ДНІВ НА ЗАПОВНЕННЯ (Робочих)
const DEADLINE_WORKING_DAYS = 3;

// 🛡 ЗАХИСТ ВІД ДУБЛІВ (Ігнорувати "Done" анкети, якщо вони створені менше N днів тому)
const IGNORE_DONE_DAYS = 30;

const QUESTIONS_LIST = [
  "1. Якість роботи", "2. Увага до деталей", "3. Самостійність", 
  "4. Надійність", "5. Комунікація", "6. Робота в команді", 
  "7. Проактивність", "8. Вирішення проблем", "9. Стресостійкість", 
  "10. Продаж ідей", "11. Розвиток", 
  "12. Сильні сторони", "13. Зони росту", "14. Що заважає"
];

// ==========================================
// 🟢 МЕНЮ
// ==========================================
function onOpen() {
  SpreadsheetApp.getUi().createMenu('💎 HR Admin')
    .addItem('📊 Сформувати звіт (Sidebar)', 'showSidebarReport')
    .addSeparator()
    .addItem('📝 Створити чернетку листа (Gmail)', 'createDraftFromActiveRow') 
    .addToUi();
}

function showSidebarReport() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const row = sheet.getActiveCell().getRow();
  
  if (sheet.getName() !== 'Database') { SpreadsheetApp.getUi().alert('Перейдіть на вкладку "Database".'); return; }
  if (row <= 1) { SpreadsheetApp.getUi().alert('Виберіть рядок.'); return; }
  
  const subjectName = sheet.getRange(row, 2).getValue(); 
  
  const html = generateReportPage(subjectName)
      .setTitle(`Звіт: ${subjectName}`)
      .setSandboxMode(HtmlService.SandboxMode.IFRAME);
      
  SpreadsheetApp.getUi().showSidebar(html);
}

// ==========================================
// 🌐 ВХІДНА ТОЧКА (З ПОКРАЩЕНИМ КЕШУВАННЯМ)
// ==========================================
function doPost(e) {
  try {
    if (e.postData && e.postData.contents) {
      // 1. Обробка кнопок (Interactive)
      if (e.parameter && e.parameter.payload) {
        handleInteractivity(JSON.parse(e.parameter.payload));
        return ContentService.createTextOutput(""); 
      }
      
      let params;
      try { params = JSON.parse(e.postData.contents); } catch(err) {}
      
      // 2. Verification URL
      if (params && params.type === "url_verification") return ContentService.createTextOutput(params.challenge);
      
      // 3. 🔥 ЗАХИСТ ВІД ПОВТОРІВ SLACK (Retry Logic)
      if (params && params.event_id) {
        const cache = CacheService.getScriptCache();
        if (cache.get(params.event_id)) {
          // Якщо ми вже бачили цей ID - просто кажемо ОК і виходимо
          return ContentService.createTextOutput("OK");
        }
        // Запам'ятовуємо ID на 5 хвилин
        cache.put(params.event_id, 'processed', 300);
      }
      
      // 4. Обробка події
      if (params && params.event && params.event.type === "message" && !params.event.bot_id) {
        handleSlackMessage(params.event);
      }
    }
    return ContentService.createTextOutput("OK");
  } catch (error) { 
    console.error("Global Error: " + error.toString()); 
    return ContentService.createTextOutput("OK"); 
  }
}

function doGet(e) {
  if (e.parameter.mode === 'report') return generateReportPage(e.parameter.subject || "");
  if (e.parameter.token) recordOpening(e.parameter.token);
  return generateSurveyPage(e.parameter.token);
}

// ==========================================
// 🧠 SLACK LOGIC
// ==========================================
function handleSlackMessage(event) {
  const text = event.text;
  const userId = event.user;
  if (!text) return; 

  try {
    const isAdminCmd = text.toLowerCase().includes('звіт') || text.toLowerCase().includes('report') || text.toLowerCase().includes('feedback') || text.toLowerCase().includes('оцінюємо');
    if (isAdminCmd) {
      let isAllowed = false;
      if (ADMIN_IDS && Array.isArray(ADMIN_IDS) && ADMIN_IDS.includes(userId)) isAllowed = true;
      if (!isAllowed) { postToSlack(userId, `⛔️ *Доступ заборонено.*`); return; }
    }
    if (text.toLowerCase().includes('звіт') || text.toLowerCase().includes('report')) {
       let subjectName = text.replace(/звіт|report/gi, '').replace(/\*/g, '').trim();
       if (!subjectName || subjectName.length < 2) subjectName = "Колега";
       sendReportCard(userId, subjectName);
       return;
    }
    if (text.includes("<@U")) { startSurveyProcess(text, userId); }
  } catch (err) { console.error(err); }
}

function handleInteractivity(payload) {
  try {
    const action = payload.actions[0];
    const userId = payload.user.id;
    const actionId = action.action_id;

    if (actionId.startsWith("urgent_remind_action_")) {
      if (!ADMIN_IDS.includes(userId)) { postToSlack(userId, "🚫 Тільки адмін."); return; }
      const subjectName = actionId.replace("urgent_remind_action_", "");
      const count = sendUrgentRemindersBatch(subjectName);
      postToSlack(userId, `✅ Термінове нагадування розіслано ${count} колегам.`);
    }
    if (actionId.startsWith("snooze_")) {
      const type = actionId.split("_")[1];
      const url = action.value;
      let minutes = 10;
      let label = "10 хв";
      if (type === "60m") { minutes = 60; label = "1 годину"; }
      if (type === "1d")  { minutes = 1440; label = "1 день"; }
      const result = setSnoozeTimeInDB(url, minutes);
      if (result.success) {
        const newBlocks = [
          { type: "section", text: { type: "mrkdwn", text: `✅ *Відкладено.* Нагадаю орієнтовно через ${label}.` } },
          { type: "divider" },
          { type: "actions", elements: [ { type: "button", text: { type: "plain_text", text: "✍️ Перервати і заповнити зараз" }, style: "primary", value: url, action_id: "interrupt_snooze" } ] }
        ];
        updateSlackMessage(payload.response_url, { blocks: newBlocks, replace_original: true });
      }
    }
    if (actionId === "interrupt_snooze") {
       const url = action.value;
       clearSnoozeTimeInDB(url); 
       const originalBlocks = [
        { type: "header", text: { type: "plain_text", text: "📝 Анкету відновлено", emoji: true } },
        { type: "divider" },
        { type: "section", text: { type: "mrkdwn", text: `Ти вирішив не чекати. Супер! Ось посилання:` } },
        { type: "actions", elements: [ { type: "button", text: { type: "plain_text", text: "👉 Відкрити анкету" }, style: "primary", url: url }, { type: "button", text: { type: "plain_text", text: "💤 10 хв" }, action_id: "snooze_10m", value: url }, { type: "button", text: { type: "plain_text", text: "💤 1 година" }, action_id: "snooze_60m", value: url }, { type: "button", text: { type: "plain_text", text: "💤 1 день" }, action_id: "snooze_1d", value: url } ] }
      ];
      updateSlackMessage(payload.response_url, { blocks: originalBlocks, replace_original: true });
    }
  } catch (e) { console.error(e); }
}

// ==========================================
// 🛠 CORE LOGIC (🔥 FIXED DUPLICATES)
// ==========================================

function startSurveyProcess(text, senderId) {
  const regex = /<@(U[A-Z0-9]+)(\|.*?)?>/g;
  const evaluators = [];
  let match;
  while ((match = regex.exec(text)) !== null) evaluators.push(match[1]);
  if (evaluators.length === 0) return; 
  
  let subjectName = text.replace(regex, '').replace(/оцінюємо/gi, '').replace(/feedback/gi, '').replace(/\*/g, '').trim();
  if (subjectName.length < 2) subjectName = "Колега";
  
  const sheet = getDatabaseSheet();
  const data = sheet.getDataRange().getValues(); 
  const requests = [];

  const deadlineDate = calculateDeadlineDate(DEADLINE_WORKING_DAYS);

  evaluators.forEach(uId => {
    let token = "";
    let isNewRow = true;
    let shouldSkip = false; // 🔥 Прапор для скасування відправки

    for (let i = 1; i < data.length; i++) {
      const rowUid = String(data[i][0]).trim();
      const rowSubj = String(data[i][1]).toLowerCase().trim();
      const rowStatus = String(data[i][3]).trim();
      const rowDate = new Date(data[i][4]);

      // Якщо знайшли збіг по юзеру і колезі
      if (rowUid === uId && rowSubj === subjectName.toLowerCase()) {
        
        // 1. Якщо анкета ще не заповнена -> Нагадуємо (Duplicate)
        if (rowStatus !== 'done') {
          token = data[i][2]; 
          isNewRow = false;
          shouldSkip = false; // Шлемо нагадування
          break;
        }

        // 2. Якщо анкета ВЖЕ заповнена (Done)
        if (rowStatus === 'done') {
           // Перевіряємо, як давно вона створена
           const diffTime = Math.abs(new Date() - rowDate);
           const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));
           
           // Якщо це свіжа анкета (менше 30 днів), то ми НЕ створюємо нову
           if (diffDays < IGNORE_DONE_DAYS) {
             isNewRow = false;
             shouldSkip = true; // 🔥 ІГНОРУЄМО! Не шлемо нічого.
             break;
           }
        }
      }
    }

    // Якщо ми вирішили пропустити (бо юзер вже заповнив свіжу анкету)
    if (shouldSkip) {
       // Можна написати в консоль або просто вийти
       return; 
    }

    if (isNewRow) {
      token = Utilities.getUuid();
      sheet.appendRow([uId, subjectName, token, 'pending', new Date(), '', '', '']);
    }

    const url = `${WEB_APP_URL}?token=${token}`;
    
    const blocks = [
      { type: "header", text: { type: "plain_text", text: "📬 Новий запит на фідбек", emoji: true } },
      { type: "divider" },
      { type: "section", text: { type: "mrkdwn", text: `Привіт! Нам потрібна твоя думка про колегу: *${subjectName}*.\n\n📅 *Дедлайн: ${deadlineDate}* (3 робочі дні).\nЗаповни, будь ласка, анкету.` } },
      {
        type: "actions",
        elements: [
          { type: "button", text: { type: "plain_text", text: "👉 Відкрити анкету" }, style: "primary", url: url },
          { type: "button", text: { type: "plain_text", text: "💤 10 хв" }, action_id: "snooze_10m", value: url },
          { type: "button", text: { type: "plain_text", text: "💤 1 година" }, action_id: "snooze_60m", value: url },
          { type: "button", text: { type: "plain_text", text: "💤 1 день" }, action_id: "snooze_1d", value: url }
        ]
      },
      { type: "context", elements: [{ type: "mrkdwn", text: "💾 _Відповіді зберігаються автоматично._" }] }
    ];

    requests.push({
      url: 'https://slack.com/api/chat.postMessage',
      method: 'post',
      headers: { Authorization: 'Bearer ' + SLACK_TOKEN },
      contentType: 'application/json',
      payload: JSON.stringify({ channel: uId, text: "Запит на фідбек", blocks: blocks })
    });
  });

  if (requests.length > 0) UrlFetchApp.fetchAll(requests);
  
  // Повідомляємо адміна тільки якщо реально щось відправили
  if (requests.length > 0) {
    sendSlackMessage(senderId, `✅ Запрошення оброблено для ${requests.length} колег.`);
  } else {
    // Якщо всі "скіпнуті", пишемо про це
    sendSlackMessage(senderId, `ℹ️ Всі вказані колеги вже мають свіжі анкети.`);
  }
}

// 📧 GMAIL DRAFTS
function createDraftFromActiveRow() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const ui = SpreadsheetApp.getUi();
  
  if (sheet.getName() !== 'Database') { ui.alert('⚠️ Перейдіть на вкладку "Database".'); return; }

  const row = sheet.getActiveCell().getRow();
  if (row <= 1) { ui.alert('⚠️ Виберіть рядок з даними.'); return; }

  const subjectName = sheet.getRange(row, 2).getValue(); 
  const anchorDateValue = sheet.getRange(row, 5).getValue(); 

  if (!subjectName) { ui.alert('⚠️ Немає імені.'); return; }
  
  let anchorDate = new Date();
  if (anchorDateValue && anchorDateValue instanceof Date) {
    anchorDate = anchorDateValue;
  }

  const htmlBody = generateEmailBody(subjectName, anchorDate);
  const emailSubject = `HR Feedback: ${subjectName}`;
  const recipient = ""; 

  try {
    GmailApp.createDraft(recipient, emailSubject, "", { htmlBody: htmlBody });
    ui.alert(`✅ Чернетку створено!\nПеревірте Gmail.`);
  } catch (e) { ui.alert(`❌ Помилка: ${e.toString()}`); }
}

function generateEmailBody(subjectName, anchorDate) {
  const sheet = getDatabaseSheet(); 
  const data = sheet.getDataRange().getValues();
  const questions = QUESTIONS_LIST;
  
  const startDate = new Date(anchorDate);
  startDate.setDate(startDate.getDate() - PERIOD_WINDOW_DAYS); 
  const endDate = new Date(anchorDate);
  endDate.setDate(endDate.getDate() + PERIOD_WINDOW_DAYS); 

  const aggregatedAnswers = new Array(questions.length).fill(0).map(() => []);
  let totalResponses = 0;
  
  for (let i = 1; i < data.length; i++) {
    const rowSubj = String(data[i][1]).toLowerCase().trim();
    const status = data[i][3];
    const createdDate = new Date(data[i][4]); 

    if (rowSubj === subjectName.toLowerCase().trim() && 
        status === 'done' && 
        createdDate >= startDate && 
        createdDate <= endDate) {
      
      totalResponses++;
      for (let q = 0; q < questions.length; q++) {
         const answer = data[i][8 + q]; 
         if (answer && String(answer).trim() !== "") {
           aggregatedAnswers[q].push(answer);
         }
      }
    }
  }

  let tableRows = "";
  for (let q = 0; q < questions.length; q++) {
    const answersList = aggregatedAnswers[q];
    let rightColumnContent = "";
    if (answersList.length > 0) {
      answersList.forEach(ans => {
        rightColumnContent += `<div style="border-bottom: 1px solid #eee; padding: 8px 0; font-size: 14px;">${ans}</div>`;
      });
    } else {
      rightColumnContent = "<span style='color:#bbb; font-size: 13px;'>—</span>";
    }

    tableRows += `
      <tr>
        <td style="border: 1px solid #e0e0e0; padding: 12px; vertical-align: top; width: 35%; background-color: #f9f9f9; color: #444; font-weight: bold; font-size: 14px;">
          ${questions[q]}
        </td>
        <td style="border: 1px solid #e0e0e0; padding: 12px; vertical-align: top; width: 65%; color: #333;">
          ${rightColumnContent}
        </td>
      </tr>
    `;
  }

  return `
    <div style="font-family: Helvetica, Arial, sans-serif; color: #333; max-width: 850px; line-height: 1.5;">
      <h2 style="color: #2c3e50; border-bottom: 2px solid #4285f4; padding-bottom: 10px;">
        HR Feedback: ${subjectName}
      </h2>
      
      <p style="font-size: 15px; margin-top: 20px;">
        Ми завершили збір зворотного зв'язку. Ось результати:
      </p>

      <div style="background: #e8f0fe; padding: 10px 15px; border-radius: 8px; margin-bottom: 25px; border: 1px solid #d2e3fc; display: inline-block;">
        ✅ <strong>Враховано анкет:</strong> ${totalResponses}
      </div>

      <table style="border-collapse: collapse; width: 100%; box-shadow: 0 1px 3px rgba(0,0,0,0.1);">
        ${tableRows}
      </table>
      
      <br/>
      <p style="color: #888; font-size: 12px; text-align: center;">
        <i>Згенеровано автоматично HR Bot Assistant | Конфіденційно</i>
      </p>
    </div>
  `;
}

// ==========================================
// 🛠 HELPERS
// ==========================================

function calculateDeadlineDate(workingDays) {
  let date = new Date();
  let added = 0;
  while (added < workingDays) {
    date.setDate(date.getDate() + 1);
    const day = date.getDay();
    if (day !== 0 && day !== 6) { 
      added++;
    }
  }
  return date.toLocaleDateString('uk-UA'); 
}

function checkSnoozes() { const sheet = getDatabaseSheet(); const data = sheet.getDataRange().getValues(); const now = new Date(); const requests = []; for (let i = 1; i < data.length; i++) { const status = data[i][3]; const snoozeTime = data[i][7] ? new Date(data[i][7]) : null; const uId = data[i][0]; const subjectName = data[i][1]; const token = data[i][2]; if (status !== 'done' && snoozeTime && snoozeTime <= now) { const url = `${WEB_APP_URL}?token=${token}`; const blocks = [ { type: "header", text: { type: "plain_text", text: "⏰ Час вийшов! Ти просив нагадати...", emoji: true } }, { type: "divider" }, { type: "section", text: { type: "mrkdwn", text: `Відтермінування закінчилося.\nДавай все ж таки заповнимо анкету про: *${subjectName}*.` } }, { type: "actions", elements: [ { type: "button", text: { type: "plain_text", text: "👉 Відкрити анкету" }, style: "primary", url: url }, { type: "button", text: { type: "plain_text", text: "💤 10 хв" }, action_id: "snooze_10m", value: url }, { type: "button", text: { type: "plain_text", text: "💤 1 година" }, action_id: "snooze_60m", value: url }, { type: "button", text: { type: "plain_text", text: "💤 1 день" }, action_id: "snooze_1d", value: url } ] } ]; requests.push({ url: 'https://slack.com/api/chat.postMessage', method: 'post', headers: { Authorization: 'Bearer ' + SLACK_TOKEN }, contentType: 'application/json', payload: JSON.stringify({ channel: uId, text: "Нагадування", blocks: blocks }) }); sheet.getRange(i + 1, 8).clearContent(); } } if (requests.length > 0) { try { UrlFetchApp.fetchAll(requests); } catch (e) { console.error(e); } } }
function setSnoozeTimeInDB(urlToken, minutes) { let token = urlToken; const uuidRegex = /[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}/; const match = urlToken.match(uuidRegex); if (match) token = match[0]; const sheet = getDatabaseSheet(); const data = sheet.getDataRange().getValues(); let found = false; for (let i = 1; i < data.length; i++) { if (String(data[i][2]).trim() === String(token).trim()) { const futureTime = new Date(); futureTime.setMinutes(futureTime.getMinutes() + minutes); sheet.getRange(i + 1, 8).setValue(futureTime); found = true; break; } } if (!found) return { success: false, error: "Token not found" }; return { success: true }; }
function clearSnoozeTimeInDB(urlToken) { let token = urlToken; const uuidRegex = /[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}/; const match = urlToken.match(uuidRegex); if (match) token = match[0]; const sheet = getDatabaseSheet(); const data = sheet.getDataRange().getValues(); for (let i = 1; i < data.length; i++) { if (String(data[i][2]).trim() === String(token).trim()) { sheet.getRange(i + 1, 8).clearContent(); break; } } }
function recordOpening(token) { const sheet = getDatabaseSheet(); const data = sheet.getDataRange().getValues(); for (let i = 1; i < data.length; i++) { if (data[i][2] == token) { sheet.getRange(i + 1, 7).setValue(new Date()); break; } } }
function sendReportCard(userId, subjectName) { const stats = getSurveyStats(subjectName); const sheetUrl = `https://docs.google.com/spreadsheets/d/${SPREADSHEET_ID}`; const reportUrl = `${WEB_APP_URL}?mode=report&subject=${encodeURIComponent(subjectName)}`; const blocks = [ { type: "header", text: { type: "plain_text", text: `📊 Статус: ${subjectName}`, emoji: true } }, { type: "divider" }, { type: "section", fields: [ { type: "mrkdwn", text: `*📅 Старт:*\n${stats.startDate}` }, { type: "mrkdwn", text: `*📩 Всього:*\n${stats.total}` }, { type: "mrkdwn", text: `*✅ Готово:*\n${stats.done}` }, { type: "mrkdwn", text: `*👀 В процесі:*\n${stats.inProgress}` }, { type: "mrkdwn", text: `*⏳ Очікуємо:*\n${stats.pending}` } ]}, { type: "divider" }, { type: "actions", elements: [ { type: "button", text: { type: "plain_text", text: "🚀 Відкрити Звіт" }, style: "primary", url: reportUrl }, { type: "button", text: { type: "plain_text", text: "📗 Таблиця" }, url: sheetUrl } ]} ]; if (stats.pending > 0 || stats.inProgress > 0) { blocks[4].elements.push({ type: "button", text: { type: "plain_text", text: "🔔 Нагадати терміново" }, style: "danger", value: "urgent_remind", action_id: "urgent_remind_action_" + subjectName }); } postToSlack(userId, "Статус звіту", blocks); }
function sendUrgentRemindersBatch(subjectName) { const sheet = getDatabaseSheet(); const data = sheet.getDataRange().getValues(); const searchKey = subjectName.toLowerCase().trim(); const requests = []; for (let i = 1; i < data.length; i++) { const rowSubj = String(data[i][1]).toLowerCase().trim(); const status = data[i][3]; const uId = data[i][0]; const token = data[i][2]; if (rowSubj === searchKey && status !== 'done') { const url = `${WEB_APP_URL}?token=${token}`; const blocks = [ { type: "header", text: { type: "plain_text", text: "🔥 Термінове нагадування!", emoji: true } }, { type: "divider" }, { type: "section", text: { type: "mrkdwn", text: `<@${uId}>, привіт!\nКритично не вистачає твого фідбеку по *${subjectName}*!` } }, { type: "actions", elements: [ { type: "button", text: { type: "plain_text", text: "✍️ Заповнити зараз" }, style: "primary", url: url }, { type: "button", text: { type: "plain_text", text: "💤 10 хв" }, action_id: "snooze_10m", value: url }, { type: "button", text: { type: "plain_text", text: "💤 1 година" }, action_id: "snooze_60m", value: url } ]} ]; requests.push({ url: 'https://slack.com/api/chat.postMessage', method: 'post', headers: { Authorization: 'Bearer ' + SLACK_TOKEN }, contentType: 'application/json', payload: JSON.stringify({ channel: uId, text: "Терміново", blocks: blocks }) }); } } if (requests.length > 0) UrlFetchApp.fetchAll(requests); return requests.length; }
function getSurveyStats(subjectName) { const sheet = getDatabaseSheet(); const data = sheet.getDataRange().getValues(); let total=0, done=0, pending=0, inProgress=0; let minDate = new Date(); const searchKey = subjectName.toLowerCase().trim(); for (let i=1; i<data.length; i++) { if (String(data[i][1]).toLowerCase().trim() === searchKey) { total++; if (data[i][3] === 'done') done++; else { if (data[i][6]) inProgress++; else pending++; } if (new Date(data[i][4]) < minDate) minDate = new Date(data[i][4]); } } if (total===0) minDate=new Date(); return { total, done, pending, inProgress, startDate: minDate.toLocaleDateString('uk-UA') }; }
function getDatabaseSheet() { return SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('Database'); }
function postToSlack(ch, txt, blk) { try { UrlFetchApp.fetch('https://slack.com/api/chat.postMessage', { method: 'post', contentType: 'application/json', headers: { Authorization: 'Bearer ' + SLACK_TOKEN }, payload: JSON.stringify({ channel: ch, text: txt, blocks: blk }) }); } catch (e) { console.error(e); } }
function sendSlackMessage(ch, txt) { postToSlack(ch, txt, null); }
function updateSlackMessage(responseUrl, payload) { try { UrlFetchApp.fetch(responseUrl, { method: 'post', contentType: 'application/json', payload: JSON.stringify(payload) }); } catch(e) { console.error(e); } }
function generateSurveyPage(token) { let subjectName = "..."; let isDone = false; let validToken = false; if (token) { const sheet = getDatabaseSheet(); const data = sheet.getDataRange().getValues(); for (let i = 1; i < data.length; i++) { if (data[i][2] == token) { subjectName = data[i][1]; if (data[i][3] === 'done') isDone = true; validToken = true; break; } } } if (token && !validToken) return HtmlService.createHtmlOutput("<h3>Link invalid.</h3>"); if (isDone) return HtmlService.createHtmlOutput("<h3>Done. Thank you!</h3>"); const template = HtmlService.createTemplateFromFile('index'); template.subjectName = subjectName; template.token = token || ""; return template.evaluate().setTitle('360 Survey').addMetaTag('viewport', 'width=device-width, initial-scale=1'); }
function generateReportPage(subjectToFind) { 
  if (!subjectToFind) return HtmlService.createHtmlOutput("Error: No name provided."); 
  const sheet = getDatabaseSheet(); 
  const data = sheet.getDataRange().getValues(); 
  const allAnswers = []; 
  const searchKey = subjectToFind.toLowerCase().trim(); 
  for (let i = 1; i < data.length; i++) { 
    const rowSubject = String(data[i][1]).toLowerCase().trim(); 
    if (rowSubject === searchKey) { 
      let dateVal = data[i][4]; 
      let dateStr = ""; 
      try { dateStr = new Date(dateVal).toISOString(); } catch(e) { dateStr = new Date().toISOString(); } 
      allAnswers.push({ date: dateStr, responses: data[i].slice(8, 22) }); 
    } 
  } 
  const template = HtmlService.createTemplateFromFile('report'); 
  template.subject = subjectToFind; 
  template.questions = QUESTIONS_LIST; 
  template.answersJson = JSON.stringify(allAnswers); 
  return template.evaluate().setTitle('Admin Dashboard'); 
}
function processForm(formObject) { const token = formObject.token; const sheet = getDatabaseSheet(); const data = sheet.getDataRange().getValues(); let rowIndex = -1; for (let i = 1; i < data.length; i++) { if (data[i][2] == token) { rowIndex = i; break; } } if (rowIndex === -1) throw new Error("Session not found"); const range = sheet.getRange(rowIndex + 1, 1, 1, 25); const rowValues = range.getValues()[0]; rowValues[3] = 'done'; rowValues[5] = new Date(); for (let q = 1; q <= 14; q++) { rowValues[7 + q] = formObject['q' + q] || ""; } range.setValues([rowValues]); return "Success"; }
