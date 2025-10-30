/**
 * Exercise 12 — AI Translator with Terminology Glossary (Sheets + Gemini)
 * Author: Laurence “Lars” Svekis
 *
 * Features
 * - Translate text via Gemini REST API (UrlFetchApp)
 * - Enforce custom terminology from a “Glossary” sheet (A:B)
 * - Batch: Selection → next column
 * - Custom function: =AI_TRANSLATE(A2, "French", Glossary!A:B)
 * - Sidebar: quick single-shot translator (+ ping test)
 */

// ===== CONFIG =====
// Prefer Script Properties (Project Settings → Script properties: GEMINI_API_KEY)
// You can set a fallback value here if you want:
const GEMINI_API_KEY = ''; // e.g., 'AIza...'
const GEMINI_MODEL   = 'gemini-2.5-flash';
const GEMINI_BASE    = 'https://generativelanguage.googleapis.com/v1beta';

const DEFAULT_TEMP   = 0.2;
const DEFAULT_TOKENS = 1024;  // raised from 512 for long strings
const COOLDOWN_MS    = 120;

// ===== MENU / UI =====
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🧠 AI Tools')
    .addItem('Translate Selection → next column', 'translateSelectionToRight')
    .addItem('Open Translator Sidebar', 'showTranslatorSidebar')
    .addSeparator()
    .addItem('Insert Sample Text', 'insertSampleText')
    .addItem('Insert Sample Glossary', 'insertSampleGlossary')
    .addToUi();
}

// Sidebar loader (HTML file must be named "Sidebar")
function showTranslatorSidebar() {
  try {
    const html = HtmlService.createHtmlOutputFromFile('Sidebar')
      .setTitle('AI Translator (Gemini)');
    SpreadsheetApp.getUi().showSidebar(html);
  } catch (e) {
    SpreadsheetApp.getUi().alert('Failed to open sidebar: ' + e.message);
  }
}

// ===== DIAGNOSTICS =====
// Ping from Sidebar to prove wiring/scopes are OK
function ping() { return 'pong'; }

// Simple model listing to verify API key
function testGemini() {
  const key = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY') || GEMINI_API_KEY;
  if (!key) return Logger.log('❌ No GEMINI_API_KEY set');
  const url = `https://generativelanguage.googleapis.com/v1beta/models?key=${encodeURIComponent(key)}`;
  const res = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  Logger.log(res.getContentText());
}

// ===== API KEY =====
function getApiKey() {
  const p = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  return (p && p.trim()) ? p.trim() : (GEMINI_API_KEY || '').trim();
}

// ===== ROBUST EXTRACTOR =====
// Handles candidates with empty parts/text, and MAX_TOKENS truncation.
function extractGeminiText_(json) {
  try {
    if (json?.error?.message) return `❌ API error: ${json.error.message}`;

    const cand = json?.candidates?.[0];
    if (!cand) return '';

    const parts = cand?.content?.parts || [];
    const joined = parts.map(p => (p?.text || '')).join('').trim();
    if (joined) return joined;

    if (cand?.text && String(cand.text).trim()) return String(cand.text).trim();

    if (cand?.finishReason === 'MAX_TOKENS') {
      return '⚠️ Truncated (MAX_TOKENS) and no visible text. Try lowering input length or increasing max tokens.';
    }
    return '';
  } catch (e) {
    return '';
  }
}

// ===== GEMINI CALL (with logs + extractor) =====
function callGemini_(prompt) {
  const key = getApiKey();
  if (!key) return '❌ Missing GEMINI_API_KEY.';

  const payload = {
    contents: [{ role: "user", parts: [{ text: prompt }] }],
    generationConfig: { temperature: DEFAULT_TEMP, maxOutputTokens: DEFAULT_TOKENS, topP: 0.9, topK: 40 }
    // If you want to try plain text with v1beta, add at top level (NOT inside generationConfig):
    // responseMimeType: "text/plain"
  };

  const url = `${GEMINI_BASE}/models/${encodeURIComponent(GEMINI_MODEL)}:generateContent?key=${encodeURIComponent(key)}`;
  const res = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  const code = res.getResponseCode();
  const body = res.getContentText();

  if (code < 200 || code >= 300) {
    Logger.log(`Gemini HTTP ${code}\\n${body}`);
    return `❌ HTTP ${code}: ${body}`;
  }

  let out = '';
  try {
    const j = JSON.parse(body);
    out = extractGeminiText_(j);
  } catch (e) {
    Logger.log('Parse error: ' + e);
    return '❌ Parse error';
  }

  if (!out) {
    Logger.log('Empty translation body: ' + body);
    return '⚠️ Empty translation';
  }
  return out.trim();
}

// ===== PROMPT BUILDER (strict) =====
function buildTranslatePrompt_(text, targetLang, glossaryPairs) {
  const rules = [
    `Translate the text to ${targetLang}.`,
    `Output ONLY the translated text. No quotes, no brackets, no preface, no code fences.`,
    `Preserve meaning and tone. Respect proper nouns and capitalization.`
  ];

  if (Array.isArray(glossaryPairs) && glossaryPairs.length) {
    const entries = glossaryPairs
      .filter(r => r && r.length >= 2 && String(r[0]).trim())
      .map(r => `- "${String(r[0]).trim()}" → "${String(r[1]||'').trim()}"`)
      .join('\\n');
    if (entries) {
      rules.push(
        `Use these exact terminology mappings wherever applicable:`,
        entries
      );
    }
  }

  return `${rules.join('\\n')}\\n\\nText:\\n${String(text || '').trim()}`;
}

// ===== PUBLIC TRANSLATE (with retry) =====
function translateWithGlossary_(text, lang, glossary) {
  // 1) Strict “just the translation”
  const prompt1 = buildTranslatePrompt_(text, lang, glossary);
  let out = callGemini_(prompt1);
  if (out && !/^⚠️|^❌/.test(out)) return out;

  // 2) Retry once with a simpler instruction
  const prompt2 = [
    `Translate to ${lang}.`,
    `Return ONLY the translated sentence. No quotes, no commentary, no extra lines.`,
    `Text:\\n${String(text || '').trim()}`
  ].join('\\n');
  out = callGemini_(prompt2);
  return out;
}

// ===== SHEETS ACTIONS =====
function translateSelectionToRight() {
  const ui = SpreadsheetApp.getUi();
  const resp = ui.prompt('Translate Selection', 'Enter target language (e.g., French, es, de):', ui.ButtonSet.OK_CANCEL);
  if (resp.getSelectedButton() !== ui.Button.OK) return;

  const lang = (resp.getResponseText() || '').trim();
  if (!lang) { ui.alert('Please provide a language.'); return; }

  const range = SpreadsheetApp.getActiveRange();
  if (!range) { ui.alert('Select a column/region of text to translate.'); return; }

  const sheet = range.getSheet();
  const glossarySheet = SpreadsheetApp.getActive().getSheetByName('Glossary');
  const glossary = glossarySheet ? glossarySheet.getRange(1, 1, glossarySheet.getLastRow(), 2).getValues() : [];

  const values = range.getValues();
  const out = [];
  for (let r = 0; r < values.length; r++) {
    const src = String(values[r][0] || '').trim();
    if (!src) { out.push(['']); continue; }
    const translated = translateWithGlossary_(src, lang, glossary);
    out.push([translated]);
    Utilities.sleep(COOLDOWN_MS);
  }

  const outCol = range.getColumn() + range.getNumColumns();
  sheet.getRange(range.getRow(), outCol, out.length, 1).setValues(out);
  SpreadsheetApp.getActive().toast(`Translated ${out.length} row(s) → column ${outCol}`);
}

// ===== CUSTOM FUNCTION =====
// Usage: =AI_TRANSLATE(A2, "French", Glossary!A:B)
function AI_TRANSLATE(text, targetLanguage, glossaryRange) {
  const glossVals = Array.isArray(glossaryRange) ? glossaryRange : [];
  return translateWithGlossary_(text, targetLanguage || 'French', glossVals);
}

// ===== SAMPLE DATA =====
function insertSampleText() {
  const data = [
    ['Source (EN)', 'Target (→ after run)'],
    ['Please reset your password using the link in your email.', ''],
    ['Our AI tutor integrates directly with the LMS template.', ''],
    ['Download the syllabus and follow the weekly schedule closely.', ''],
    ['The marketing campaign launches next Monday at 9 AM.', ''],
    ['Customer data is stored securely according to our policy.', '']
  ];
  const sheet = SpreadsheetApp.getActiveSheet();
  sheet.clearContents();
  sheet.getRange(1, 1, data.length, data[0].length).setValues(data);
  SpreadsheetApp.getActive().toast('✅ Sample text inserted (A–B).');
}

function insertSampleGlossary() {
  let sh = SpreadsheetApp.getActive().getSheetByName('Glossary');
  if (!sh) sh = SpreadsheetApp.getActive().insertSheet('Glossary');
  sh.clearContents();
  sh.getRange(1, 1, 1, 2).setValues([['Source Term', 'Preferred Translation']]);
  const rows = [
    ['learning management system', 'système de gestion de l’apprentissage'],
    ['LMS', 'SGA'],
    ['AI tutor', 'tuteur IA'],
    ['syllabus', 'plan de cours'],
    ['privacy policy', 'politique de confidentialité'],
    ['marketing campaign', 'campagne marketing']
  ];
  sh.getRange(2, 1, rows.length, 2).setValues(rows);
  SpreadsheetApp.getActive().toast('✅ Sample Glossary inserted (Glossary!A:B).');
}

// ===== SIDEBAR ENTRYPOINT =====
function translateFromSidebar(text, targetLang) {
  try {
    const glossarySheet = SpreadsheetApp.getActive().getSheetByName('Glossary');
    const glossaryValues = glossarySheet
      ? glossarySheet.getRange(1, 1, glossarySheet.getLastRow(), 2).getValues()
      : [];
    const translation = translateWithGlossary_(text || '', targetLang || 'French', glossaryValues);
    Logger.log('Sidebar -> translate ok');
    return translation || '⚠️ Empty translation';
  } catch (e) {
    const msg = `Sidebar error: ${e && e.message ? e.message : e}`;
    Logger.log(msg);
    return '❌ ' + msg;
  }
}

// ===== OPTIONAL: quick local test =====
function manualTest() {
  const out = translateWithGlossary_(
    'Our AI tutor integrates directly with the LMS template.',
    'French',
    [['AI tutor', 'tuteur IA'], ['LMS', 'SGA']]
  );
  Logger.log(out);
}
