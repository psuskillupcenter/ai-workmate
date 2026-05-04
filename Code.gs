/************************************
 * AI WORKMATE: KNOWLEDGE LEGACY
 * Google Apps Script + Google Sheets + Gemini API
 ************************************/

const CONFIG = {
  SPREADSHEET_ID: '1cbNwLSnPB4LxbrT74eGW-aShKzSvrNZSgjNEV0GJTJ8',
  MODEL: 'gemini-2.5-flash-lite',

  RAW_SHEET: 'RAW_KNOWLEDGE',
  ANALYSIS_SHEET: 'AI_ANALYSIS',
  CARD_SHEET: 'KNOWLEDGE_CARD',
  CHAT_LOG_SHEET: 'CHAT_LOG'
};

/**
 * Open Google Sheet database
 */
function getDB() {
  return SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
}

/**
 * Web App entry point
 */
function doGet() {
  return HtmlService
    .createTemplateFromFile('Index')
    .evaluate()
    .setTitle('AI Workmate: Knowledge Legacy')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Include HTML partials
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * Custom menu in Google Sheets
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('AI Workmate')
    .addItem('1) วิเคราะห์ข้อมูลที่รอทั้งหมด', 'analyzeAllPendingRecords')
    .addItem('2) สร้าง Knowledge Card', 'generateKnowledgeCardsFromAnalysis')
    .addItem('3) ทดสอบ AI Workmate Chat', 'testKnowledgeChat')
    .addItem('4) ตรวจระบบ', 'debugCheckSystem')
    .addToUi();
}

/**
 * System check
 */
function debugCheckSystem() {
  const ss = getDB();

  const result = {
    spreadsheetName: ss.getName(),
    spreadsheetId: ss.getId(),
    sheets: ss.getSheets().map(s => s.getName()),
    apiKeyExists: !!PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY')
  };

  Logger.log(JSON.stringify(result, null, 2));
  return result;
}

/**
 * Dashboard data for Web App
 */
function getDashboardData() {
  const ss = getDB();

  const rawSheet = ss.getSheetByName(CONFIG.RAW_SHEET);
  const analysisSheet = ss.getSheetByName(CONFIG.ANALYSIS_SHEET);
  const cardSheet = ss.getSheetByName(CONFIG.CARD_SHEET);
  const chatSheet = ss.getSheetByName(CONFIG.CHAT_LOG_SHEET);

  return {
    rawCount: countDataRows_(rawSheet),
    analysisCount: countDataRows_(analysisSheet),
    cardCount: countDataRows_(cardSheet),
    chatCount: countDataRows_(chatSheet),
    pendingCount: countRawByStatus_('รอวิเคราะห์'),
    analyzedCount: countRawByStatus_('วิเคราะห์แล้ว')
  };
}

/**
 * Save raw knowledge from Web App
 */
function webSaveRawKnowledge(formData) {
  const ss = getDB();
  const sheet = ss.getSheetByName(CONFIG.RAW_SHEET);

  if (!sheet) {
    throw new Error('ไม่พบชีต RAW_KNOWLEDGE');
  }

  const recordId = 'TK' + new Date().getTime();

  sheet.appendRow([
    recordId,
    formData.staff_name || '',
    formData.department || '',
    formData.position || '',
    formData.expertise_area || '',
    formData.interview_question || '',
    formData.raw_answer || '',
    'รอวิเคราะห์',
    new Date(),
    formData.source_type || 'typed_document',
    formData.document_title || '',
    formData.consent_status || 'consented',
    formData.notes || ''
  ]);

  SpreadsheetApp.flush();

  return {
    success: true,
    message: 'บันทึกข้อมูลเรียบร้อยแล้ว',
    record_id: recordId
  };
}

/**
 * Run AI analysis from Web App
 */
function webAnalyzeAllPendingRecords() {
  const result = analyzeAllPendingRecords();
  return {
    success: true,
    message: result
  };
}

/**
 * Generate Knowledge Cards from Web App
 */
function webGenerateKnowledgeCards() {
  const result = generateKnowledgeCardsFromAnalysis();
  return {
    success: true,
    message: result
  };
}

/**
 * Get AI Analysis list for Web App
 */
function webGetAnalysisList() {
  const ss = getDB();
  const sheet = ss.getSheetByName(CONFIG.ANALYSIS_SHEET);

  if (!sheet) return [];

  return sheetToObjects_(sheet).reverse();
}

/**
 * Get Knowledge Card list for Web App
 */
function webGetKnowledgeCards() {
  const ss = getDB();
  const sheet = ss.getSheetByName(CONFIG.CARD_SHEET);

  if (!sheet) return [];

  return sheetToObjects_(sheet).reverse();
}

/**
 * Ask AI Workmate Chat from Web App
 */
function webAskKnowledgeChat(question) {
  if (!question || !question.trim()) {
    throw new Error('กรุณาพิมพ์คำถามก่อน');
  }

  return askKnowledgeChat(question);
}

/**
 * Analyze all pending records in RAW_KNOWLEDGE
 */
function analyzeAllPendingRecords() {
  const ss = getDB();
  const rawSheet = ss.getSheetByName(CONFIG.RAW_SHEET);
  const analysisSheet = ss.getSheetByName(CONFIG.ANALYSIS_SHEET);

  if (!rawSheet) throw new Error('ไม่พบชีต RAW_KNOWLEDGE');
  if (!analysisSheet) throw new Error('ไม่พบชีต AI_ANALYSIS');

  const rawData = rawSheet.getDataRange().getValues();

  if (rawData.length <= 1) {
    return 'ยังไม่มีข้อมูลใน RAW_KNOWLEDGE';
  }

  const headers = rawData[0];
  const recordIdCol = headers.indexOf('record_id');
  const rawAnswerCol = headers.indexOf('raw_answer');
  const statusCol = headers.indexOf('status');

  if (recordIdCol === -1 || rawAnswerCol === -1 || statusCol === -1) {
    throw new Error('RAW_KNOWLEDGE ยังมีหัวคอลัมน์ไม่ครบ');
  }

  let successCount = 0;
  let failCount = 0;
  let messages = [];

  for (let i = 1; i < rawData.length; i++) {
    const row = rawData[i];
    const recordId = row[recordIdCol];
    const rawAnswer = row[rawAnswerCol];
    const status = row[statusCol];

    if (!recordId || !rawAnswer) continue;
    if (status === 'วิเคราะห์แล้ว') continue;

    try {
      const input = rowToObject_(headers, row);
      const result = callGeminiForTacitAnalysis_(input);

      appendAnalysis_(analysisSheet, input, result);

      rawSheet.getRange(i + 1, statusCol + 1).setValue('วิเคราะห์แล้ว');

      successCount++;
      Utilities.sleep(800);

    } catch (err) {
      rawSheet.getRange(i + 1, statusCol + 1).setValue('วิเคราะห์ไม่สำเร็จ');
      failCount++;
      messages.push(recordId + ': ' + err.message);
    }
  }

  SpreadsheetApp.flush();

  const msg = `วิเคราะห์สำเร็จ ${successCount} รายการ / ไม่สำเร็จ ${failCount} รายการ`;
  Logger.log(msg);

  if (messages.length > 0) {
    Logger.log(messages.join('\n'));
  }

  return msg;
}

/**
 * Call Gemini API for Tacit Knowledge Analysis
 */
function callGeminiForTacitAnalysis_(input) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');

  if (!apiKey) {
    throw new Error('ยังไม่ได้ตั้งค่า GEMINI_API_KEY ใน Script Properties');
  }

  const url =
    'https://generativelanguage.googleapis.com/v1beta/models/' +
    CONFIG.MODEL +
    ':generateContent?key=' +
    apiKey;

  const prompt = `
คุณคือผู้เชี่ยวชาญด้าน Knowledge Management, Tacit Knowledge Transfer และการพัฒนาองค์กร

จงวิเคราะห์ข้อความประสบการณ์ของบุคลากร และสกัดเป็นความรู้ฝังลึกที่องค์กรสามารถนำไปถ่ายทอดต่อได้

ข้อมูลต้นฉบับ:
- record_id: ${input.record_id || ''}
- staff_name: ${input.staff_name || ''}
- department: ${input.department || ''}
- position: ${input.position || ''}
- expertise_area: ${input.expertise_area || ''}
- document_title: ${input.document_title || ''}
- interview_question: ${input.interview_question || ''}
- raw_answer: ${input.raw_answer || ''}

ให้ตอบกลับเป็น JSON เท่านั้น
ห้ามใส่ Markdown
ห้ามใส่ code block
ห้ามใส่คำอธิบายก่อนหรือหลัง JSON

โครงสร้าง JSON:
{
  "knowledge_category": "",
  "key_insight": "",
  "hidden_technique": "",
  "common_problem": "",
  "solution_pattern": "",
  "risk_warning": "",
  "step_by_step": "",
  "advice_for_new_staff": "",
  "keywords": "",
  "confidence_level": "",
  "recommended_action": ""
}
`;

  const payload = {
    contents: [
      {
        parts: [{ text: prompt }]
      }
    ],
    generationConfig: {
      temperature: 0.2,
      responseMimeType: 'application/json'
    }
  };

  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  let response;
  let responseText = '';

  for (let attempt = 1; attempt <= 3; attempt++) {
    response = UrlFetchApp.fetch(url, options);
    responseText = response.getContentText();

    if (response.getResponseCode() === 200) break;

    if (response.getResponseCode() === 429 || response.getResponseCode() === 503) {
      Utilities.sleep(2000 * attempt);
      continue;
    }

    throw new Error('Gemini API Error: ' + responseText);
  }

  if (response.getResponseCode() !== 200) {
    throw new Error('Gemini API ยังไม่พร้อม: ' + responseText);
  }

  const apiJson = JSON.parse(responseText);

  if (!apiJson.candidates || !apiJson.candidates[0]) {
    throw new Error('Gemini API ไม่ได้ส่งผลลัพธ์กลับมา: ' + responseText);
  }

  let outputText = apiJson.candidates[0].content.parts[0].text;

  outputText = outputText
    .replace(/```json/g, '')
    .replace(/```/g, '')
    .trim();

  return JSON.parse(outputText);
}

/**
 * Append AI Analysis result to AI_ANALYSIS sheet
 */
function appendAnalysis_(sheet, input, result) {
  const analysisId = 'AN' + new Date().getTime();

  sheet.appendRow([
    analysisId,
    input.record_id || '',
    result.knowledge_category || '',
    result.key_insight || '',
    result.hidden_technique || '',
    result.common_problem || '',
    result.solution_pattern || '',
    result.risk_warning || '',
    result.step_by_step || '',
    result.advice_for_new_staff || '',
    result.keywords || '',
    new Date(),
    result.confidence_level || 'Medium',
    result.recommended_action || '',
    'รอตรวจทาน'
  ]);
}

/**
 * Generate Knowledge Cards from AI_ANALYSIS
 */
function generateKnowledgeCardsFromAnalysis() {
  const ss = getDB();
  const analysisSheet = ss.getSheetByName(CONFIG.ANALYSIS_SHEET);
  const cardSheet = ss.getSheetByName(CONFIG.CARD_SHEET);
  const rawSheet = ss.getSheetByName(CONFIG.RAW_SHEET);

  if (!analysisSheet) throw new Error('ไม่พบชีต AI_ANALYSIS');
  if (!cardSheet) throw new Error('ไม่พบชีต KNOWLEDGE_CARD');

  const analysisData = analysisSheet.getDataRange().getValues();

  if (analysisData.length <= 1) {
    return 'ยังไม่มีข้อมูลใน AI_ANALYSIS';
  }

  const analysisHeaders = analysisData[0];

  const existingCards = cardSheet.getDataRange().getValues();
  const existingCardRecordIds = new Set();

  if (existingCards.length > 1) {
    const cardHeaders = existingCards[0];
    const recordIdIndex = cardHeaders.indexOf('record_id');

    for (let i = 1; i < existingCards.length; i++) {
      if (existingCards[i][recordIdIndex]) {
        existingCardRecordIds.add(String(existingCards[i][recordIdIndex]));
      }
    }
  }

  const rawMap = buildRawMap_(rawSheet);

  let count = 0;

  for (let i = 1; i < analysisData.length; i++) {
    const analysis = rowToObject_(analysisHeaders, analysisData[i]);

    if (!analysis.record_id) continue;
    if (existingCardRecordIds.has(String(analysis.record_id))) continue;

    const raw = rawMap[analysis.record_id] || {};
    const cardId = 'KC' + new Date().getTime() + count;
    const title = createCardTitle_(analysis, raw);

    cardSheet.appendRow([
      cardId,
      analysis.record_id || '',
      analysis.analysis_id || '',
      title,
      analysis.key_insight || '',
      analysis.key_insight || '',
      analysis.solution_pattern || '',
      analysis.hidden_technique || '',
      analysis.step_by_step || '',
      analysis.risk_warning || '',
      analysis.keywords || '',
      raw.staff_name || '',
      raw.department || '',
      'รอตรวจทาน',
      '',
      new Date()
    ]);

    count++;
  }

  SpreadsheetApp.flush();

  return 'สร้าง Knowledge Card สำเร็จ ' + count + ' รายการ';
}

/**
 * Ask AI Workmate Knowledge Chat
 */
function askKnowledgeChat(userQuestion) {
  const ss = getDB();
  const cardSheet = ss.getSheetByName(CONFIG.CARD_SHEET);
  const chatLogSheet = ss.getSheetByName(CONFIG.CHAT_LOG_SHEET);

  if (!cardSheet) throw new Error('ไม่พบชีต KNOWLEDGE_CARD');
  if (!chatLogSheet) throw new Error('ไม่พบชีต CHAT_LOG');

  const cards = getApprovedOrAllKnowledgeCards_(cardSheet);

  if (cards.length === 0) {
    return {
      answer: 'ยังไม่มี Knowledge Card ในระบบ จึงยังไม่สามารถตอบจากฐานความรู้ได้',
      usedCards: ''
    };
  }

  const context = cards.map(card => {
    return `
รหัสการ์ด: ${card.card_id}
ชื่อเรื่อง: ${card.title}
สรุป: ${card.summary}
Tacit Insight: ${card.tacit_insight}
Lesson Learned: ${card.lesson_learned}
Hidden Technique: ${card.hidden_technique}
Practical Guide: ${card.practical_guide}
Risk Warning: ${card.risk_warning}
Keywords: ${card.keywords}
Department: ${card.department}
`;
  }).join('\n---\n');

  const answer = callGeminiForKnowledgeChat_(userQuestion, context);
  const usedCards = cards.map(c => c.card_id).join(', ');

  chatLogSheet.appendRow([
    'CH' + new Date().getTime(),
    userQuestion,
    answer,
    '',
    usedCards,
    new Date()
  ]);

  SpreadsheetApp.flush();

  return {
    answer: answer,
    usedCards: usedCards
  };
}

/**
 * Call Gemini API for Chat
 */
function callGeminiForKnowledgeChat_(question, context) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');

  if (!apiKey) {
    throw new Error('ยังไม่ได้ตั้งค่า GEMINI_API_KEY');
  }

  const url =
    'https://generativelanguage.googleapis.com/v1beta/models/' +
    CONFIG.MODEL +
    ':generateContent?key=' +
    apiKey;

  const prompt = `
คุณคือ AI Workmate Knowledge Chat ผู้ช่วยถาม–ตอบจากฐานความรู้ฝังลึกขององค์กร

หน้าที่ของคุณคือช่วยตอบคำถามจาก Knowledge Cards ที่ให้ไว้เท่านั้น

กติกา:
1. ตอบเป็นภาษาไทยที่เข้าใจง่าย
2. ใช้เฉพาะข้อมูลจาก Knowledge Cards ด้านล่าง
3. ถ้าข้อมูลไม่พอ ให้ตอบว่า "ยังไม่พบข้อมูลเพียงพอในฐานความรู้"
4. ห้ามแต่งข้อมูลเกินจากฐานความรู้
5. ถ้าเหมาะสม ให้ตอบเป็นข้อ ๆ เพื่อให้เจ้าหน้าที่นำไปใช้ได้จริง
6. ให้ตอบด้วยน้ำเสียงเป็นผู้ช่วยเพื่อนร่วมงานที่สุภาพและเป็นมิตร

Knowledge Cards:
${context}

คำถามผู้ใช้:
${question}
`;

  const payload = {
    contents: [
      {
        parts: [{ text: prompt }]
      }
    ],
    generationConfig: {
      temperature: 0.2
    }
  };

  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  let response;
  let responseText = '';

  for (let attempt = 1; attempt <= 3; attempt++) {
    response = UrlFetchApp.fetch(url, options);
    responseText = response.getContentText();

    if (response.getResponseCode() === 200) break;

    if (response.getResponseCode() === 429 || response.getResponseCode() === 503) {
      Utilities.sleep(2000 * attempt);
      continue;
    }

    throw new Error('Gemini Chat Error: ' + responseText);
  }

  if (response.getResponseCode() !== 200) {
    throw new Error('Gemini Chat ยังไม่พร้อม: ' + responseText);
  }

  const json = JSON.parse(responseText);

  if (!json.candidates || !json.candidates[0]) {
    throw new Error('Gemini Chat ไม่ได้ส่งคำตอบกลับมา: ' + responseText);
  }

  return json.candidates[0].content.parts[0].text;
}

/**
 * Test Chat
 */
function testKnowledgeChat() {
  const question = 'เจ้าหน้าที่ใหม่ควรระวังอะไรจากความรู้ที่มีในระบบ?';
  const result = askKnowledgeChat(question);

  Logger.log(result.answer);
  return result.answer;
}

/**
 * Helper: Convert sheet to objects
 */
function sheetToObjects_(sheet) {
  const data = sheet.getDataRange().getDisplayValues();

  if (!data || data.length <= 1) return [];

  const headers = data[0];

  return data.slice(1)
    .filter(row => row.join('').trim() !== '')
    .map(row => {
      const obj = {};
      headers.forEach((h, i) => {
        obj[h] = row[i];
      });
      return obj;
    });
}

/**
 * Helper: Count data rows
 */
function countDataRows_(sheet) {
  if (!sheet) return 0;

  const lastRow = sheet.getLastRow();
  return Math.max(lastRow - 1, 0);
}

/**
 * Helper: Count RAW_KNOWLEDGE by status
 */
function countRawByStatus_(statusText) {
  const ss = getDB();
  const sheet = ss.getSheetByName(CONFIG.RAW_SHEET);

  if (!sheet) return 0;

  const data = sheet.getDataRange().getDisplayValues();

  if (data.length <= 1) return 0;

  const headers = data[0];
  const statusIndex = headers.indexOf('status');

  if (statusIndex === -1) return 0;

  let count = 0;

  for (let i = 1; i < data.length; i++) {
    if (data[i][statusIndex] === statusText) {
      count++;
    }
  }

  return count;
}

/**
 * Helper: Get Knowledge Cards
 */
function getApprovedOrAllKnowledgeCards_(sheet) {
  const data = sheet.getDataRange().getValues();

  if (data.length <= 1) return [];

  const headers = data[0];

  return data.slice(1)
    .filter(row => row.join('').trim() !== '')
    .map(row => rowToObject_(headers, row));
}

/**
 * Helper: Build RAW_KNOWLEDGE map
 */
function buildRawMap_(rawSheet) {
  const map = {};

  if (!rawSheet) return map;

  const data = rawSheet.getDataRange().getValues();

  if (data.length <= 1) return map;

  const headers = data[0];

  for (let i = 1; i < data.length; i++) {
    const obj = rowToObject_(headers, data[i]);

    if (obj.record_id) {
      map[obj.record_id] = obj;
    }
  }

  return map;
}

/**
 * Helper: Convert row to object
 */
function rowToObject_(headers, row) {
  const obj = {};

  headers.forEach((h, i) => {
    obj[h] = row[i];
  });

  return obj;
}

/**
 * Helper: Create Knowledge Card title
 */
function createCardTitle_(analysis, raw) {
  if (raw.document_title) return raw.document_title;

  if (analysis.knowledge_category && analysis.key_insight) {
    return analysis.knowledge_category + ': ' + String(analysis.key_insight).slice(0, 60);
  }

  if (analysis.knowledge_category) return analysis.knowledge_category;

  return 'Knowledge Card from ' + (analysis.record_id || '');
}
