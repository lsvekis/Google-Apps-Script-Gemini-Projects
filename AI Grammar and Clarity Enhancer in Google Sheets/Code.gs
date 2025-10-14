/**
 * AI Grammar & Clarity Enhancer for Google Sheets using Gemini REST API
 * Author: Laurence “Lars” Svekis
 */

// === CONFIG ===
const GEMINI_API_KEY = 'YOUR_API_KEY_HERE'; // or store in Script Properties as GEMINI_API_KEY
const GEMINI_MODEL = 'gemini-2.5-flash';
const BASE_URL = 'https://generativelanguage.googleapis.com/v1beta';
const MAX_TOKENS = 512;
const TEMPERATURE = 0.4; // lower = more deterministic

// Read key from Script Properties if present
function getApiKey() {
  const p = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  return (p && p.trim()) ? p.trim() : GEMINI_API_KEY;
}

// ===== MENU =====
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🧠 AI Tools')
    .addItem('Open Grammar Sidebar', 'showGrammarSidebar')
    .addItem('Correct Selection → next column', 'correctSelectionToRight')
    .addItem('Insert Sample Data', 'insertSampleData')
    .addToUi();
}

// ===== SIDEBAR =====
function showGrammarSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('Sidebar')
    .setTitle('Gemini Grammar Checker');
  SpreadsheetApp.getUi().showSidebar(html);
}

// ===== GEMINI CALL (robust) =====
function extractGeminiText(json) {
  try {
    if (json?.error?.message) return `❌ API error: ${json.error.message}`;
    const block = json?.promptFeedback?.blockReason;
    if (block) {
      const details = (json?.promptFeedback?.safetyRatings || [])
        .map(r => `${r.category}:${r.probability}`).join(', ');
      return `❌ Blocked by safety (${block}${details ? ' — ' + details : ''}).`;
    }
    const cand = json?.candidates?.[0];
    if (!cand) return '';
    if (cand.finishReason === 'MAX_TOKENS') {
      const parts = cand?.content?.parts || [];
      const t = parts.map(p => p?.text || '').join('').trim();
      return t ? `${t}\n\n⚠️ Truncated (MAX_TOKENS). Increase max tokens or shorten input.`
               : '⚠️ Truncated (MAX_TOKENS) and no visible text.';
    }
    const parts = cand?.content?.parts || [];
    const text = parts.map(p => p?.text || '').join('').trim();
    if (text) return text;
    const cText = cand?.text;
    if (cText && String(cText).trim()) return String(cText).trim();
    return '';
  } catch (e) {
    return '';
  }
}

function callGeminiGrammarCheck(text) {
  const apiKey = getApiKey();
  if (!apiKey || apiKey === 'YOUR_API_KEY_HERE') {
    return '❌ Error: Set GEMINI_API_KEY in Code.gs or Script Properties.';
  }
  const url = `${BASE_URL}/models/${GEMINI_MODEL}:generateContent?key=${encodeURIComponent(apiKey)}`;
  const prompt = [
    'You are a professional English editor.',
    'Correct grammar, spelling, and clarity while preserving meaning and tone.',
    'Return ONLY the corrected text (no explanations or headings).',
    '',
    'Text:',
    String(text || '')
  ].join('\n');

  const payload = {
    contents: [{ role: "user", parts: [{ text: prompt }] }],
    generationConfig: { temperature: TEMPERATURE, maxOutputTokens: MAX_TOKENS, topP: 0.95, topK: 40 }
  };

  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  try {
    const res = UrlFetchApp.fetch(url, options);
    const code = res.getResponseCode();
    const body = res.getContentText();
    if (code < 200 || code >= 300) {
      return `❌ Error: Gemini HTTP ${code}. ${body}`;
    }
    const data = JSON.parse(body);
    const out = extractGeminiText(data);
    if (out && out.trim()) return out;
    Logger.log('Empty correction output. Raw response:\n' + body);
    return '⚠️ No corrections returned (see Logs for raw response).';
  } catch (e) {
    return '❌ Error: ' + e.message;
  }
}

// ===== CUSTOM FUNCTION =====
// =AI_GRAMMAR(A2)
function AI_GRAMMAR(text) {
  return callGeminiGrammarCheck(text);
}

// ===== SHEET ACTIONS =====
function correctSelectionToRight() {
  const range = SpreadsheetApp.getActiveRange();
  if (!range) {
    SpreadsheetApp.getUi().alert('Please select one or more cells first.');
    return;
  }
  const values = range.getValues();
  const out = [];
  for (let i = 0; i < values.length; i++) {
    const original = values[i][0];
    const corrected = callGeminiGrammarCheck(original);
    out.push([corrected]);
    Utilities.sleep(150);
  }
  const sheet = range.getSheet();
  const startRow = range.getRow();
  const nextCol = range.getColumn() + 1;
  sheet.getRange(startRow, nextCol, out.length, 1).setValues(out);
  SpreadsheetApp.getActive().toast(`✅ Corrected ${out.length} row(s).`);
}

// ===== SAMPLE DATA =====
function insertSampleData() {
  const data = [
    ['Original Text', 'Corrected Text'],
    ['The company are planning to expand there operations.', ''],
    ['She dont have enough experience for the role.', ''],
    ['This report were finished yesterday.', ''],
    ['We was excited about the upcoming project.', ''],
    ['Every employees must complete the training.', '']
  ];
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  sheet.clearContents();
  sheet.getRange(1, 1, data.length, data[0].length).setValues(data);
  SpreadsheetApp.getActive().toast('✅ Sample data inserted.');
}
