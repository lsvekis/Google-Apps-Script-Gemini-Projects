/**
 * AI Text Rewriter for Google Sheets using Gemini REST API
 * Author: Laurence “Lars” Svekis
 */

const GEMINI_API_KEY = 'YOUR_API_KEY_HERE';
const GEMINI_MODEL = 'gemini-2.5-flash';
const BASE_URL = 'https://generativelanguage.googleapis.com/v1beta';

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🧠 AI Tools')
    .addItem('Open Rewriter Sidebar', 'showRewriterSidebar')
    .addItem('Insert Sample Data', 'insertSampleData')
    .addToUi();
}

function showRewriterSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('Sidebar')
    .setTitle('Gemini Rewriter');
  SpreadsheetApp.getUi().showSidebar(html);
}

function callGeminiRewrite(text, style = 'professional tone', temperature = 0.6) {
  const url = `${BASE_URL}/models/${GEMINI_MODEL}:generateContent?key=${GEMINI_API_KEY}`;
  const payload = {
    contents: [
      { role: "user", parts: [{ text: `Rewrite the following text in a ${style}:
${text}` }] }
    ],
    generationConfig: { temperature, maxOutputTokens: 512 }
  };

  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  try {
    const res = UrlFetchApp.fetch(url, options);
    const data = JSON.parse(res.getContentText());
    return data?.candidates?.[0]?.content?.parts?.[0]?.text || '⚠️ No rewrite returned.';
  } catch (e) {
    return '❌ Error: ' + e.message;
  }
}

function AI_REWRITE(text, style) {
  return callGeminiRewrite(text, style || 'neutral tone');
}

function insertSampleData() {
  const data = [
    ['Original Text', 'Rewritten Text'],
    ['Our company values innovation and collaboration across all departments.', ''],
    ['This policy ensures all employees have equal access to training opportunities.', ''],
    ['Please submit your report by Friday at noon.', '']
  ];
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  sheet.clearContents();
  sheet.getRange(1, 1, data.length, data[0].length).setValues(data);
  SpreadsheetApp.getActive().toast('✅ Sample data inserted');
}
