/**
 * Exercise 4 — AI Translator & Tone Converter for Google Sheets (Gemini REST)
 * Author: Laurence “Lars” Svekis
 *
 * Features:
 *  - Translate selection → next column
 *  - Tone-convert selection → next column
 *  - Sidebar UI: translate & tone settings (see Sidebar.html)
 *  - Custom functions: =AI_TRANSLATE(), =AI_TONE()
 */

const GEMINI_API_KEY = ''; // Optional fallback, prefer Script Properties
const GEMINI_MODEL   = 'gemini-2.5-flash';
const BASE_URL       = 'https://generativelanguage.googleapis.com/v1beta';

const DEFAULT_TEMP   = 0.3;
const DEFAULT_TOKENS = 1024;
const MAX_TOKENS_HARD_CAP = 2048;
const COOLDOWN_MS    = 120;

const CHUNK_CHAR_LIMIT = 3000;
const CHUNK_JOINER     = '\n';

// === MENU ===
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🧠 AI Tools')
    .addItem('Open Translator Sidebar', 'showTranslatorSidebar')
    .addSeparator()
    .addItem('Translate Selection → next column', 'translateSelectionToRight')
    .addItem('Tone-convert Selection → next column', 'toneSelectionToRight')
    .addSeparator()
    .addItem('Insert Sample Data', 'insertSampleData')
    .addToUi();
}

function showTranslatorSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('Sidebar')
    .setTitle('Gemini Translator & Tone');
  SpreadsheetApp.getUi().showSidebar(html);
}

// === API KEY ===
function getApiKey() {
  const prop = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  return prop && prop.trim() ? prop.trim() : GEMINI_API_KEY;
}

// === ROBUST RESPONSE EXTRACTION ===
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
      const partial = parts.map(p => p?.text || '').join('').trim();
      return partial
        ? `${partial}\n\n⚠️ Truncated (MAX_TOKENS). Increase max tokens or shorten input.`
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

// === GENERATE TEXT (Gemini Request) ===
function geminiGenerateText(prompt, { temperature = DEFAULT_TEMP, maxTokens = DEFAULT_TOKENS } = {}) {
  const apiKey = getApiKey();
  if (!apiKey) return '❌ Set GEMINI_API_KEY first in Script Properties.';

  const capped = Math.min(Number(maxTokens) || DEFAULT_TOKENS, MAX_TOKENS_HARD_CAP);
  const payload = {
    contents: [{ role: "user", parts: [{ text: prompt }] }],
    generationConfig: { temperature, maxOutputTokens: capped, topP: 0.95, topK: 40 }
  };

  const url = `${BASE_URL}/models/${encodeURIComponent(GEMINI_MODEL)}:generateContent?key=${encodeURIComponent(apiKey)}`;
  const options = { method: 'post', contentType: 'application/json', payload: JSON.stringify(payload), muteHttpExceptions: true };

  try {
    const res  = UrlFetchApp.fetch(url, options);
    const code = res.getResponseCode();
    const body = res.getContentText();

    if (code < 200 || code >= 300) return `❌ Error: Gemini HTTP ${code}. ${body}`;
    const json = JSON.parse(body);
    const text = extractGeminiText(json);
    return text || '⚠️ No output (see Logs).';
  } catch (e) {
    return '❌ Error: ' + e.message;
  }
}

// === CHUNKING ===
function splitTextIntoChunks(input, limit = CHUNK_CHAR_LIMIT) {
  const text = String(input || '').replace(/\r\n/g, '\n');
  if (text.length <= limit) return [text];

  const paras = text.split(/\n{2,}/);
  if (paras.length > 1) return packChunks_(paras, limit);

  const sents = text.split(/(?<=[.!?])\s+/);
  if (sents.length > 1) return packChunks_(sents, limit);

  const chunks = [];
  for (let i = 0; i < text.length; i += limit) chunks.push(text.slice(i, i + limit));
  return chunks;
}

function packChunks_(parts, limit) {
  const out = [];
  let buf = '';
  for (const part of parts) {
    const candidate = buf ? buf + '\n\n' + part : part;
    if (candidate.length > limit) {
      if (buf) out.push(buf);
      if (part.length > limit) {
        for (let i = 0; i < part.length; i += limit) out.push(part.slice(i, i + limit));
        buf = '';
      } else buf = part;
    } else buf = candidate;
  }
  if (buf) out.push(buf);
  return out;
}

// === LANGUAGE NORMALIZATION ===
function normalizeLanguage(lang) {
  if (!lang) return 'English';
  const s = String(lang).trim().toLowerCase();
  const map = {
    en:'English', fr:'French', es:'Spanish', de:'German',
    it:'Italian', pt:'Portuguese', hi:'Hindi', ja:'Japanese',
    ko:'Korean', zh:'Chinese'
  };
  return map[s] || lang;
}

// === TRANSLATE & TONE ===
function callGeminiTranslate(text, targetLanguage, tone, temperature, maxTokens) {
  const tgt = normalizeLanguage(targetLanguage || 'English');
  const t   = tone ? ` and adjust tone to "${tone}"` : '';
  const mkPrompt = (chunk) => [
    `Translate the text below into ${tgt}${t}.`,
    'Keep meaning accurate; do not add or drop information.',
    'Return ONLY the translation as plain text.',
    '',
    chunk
  ].join('\n');

  const chunks = splitTextIntoChunks(text);
  const outs = [];
  for (const c of chunks) {
    const out = geminiGenerateText(mkPrompt(c), {
      temperature: (typeof temperature === 'number' ? temperature : DEFAULT_TEMP),
      maxTokens:   (typeof maxTokens   === 'number' ? maxTokens   : DEFAULT_TOKENS * 2)
    });
    outs.push(out);
    Utilities.sleep(COOLDOWN_MS);
  }
  return outs.join(CHUNK_JOINER);
}

function callGeminiTone(text, tone, temperature, maxTokens) {
  const tn = tone || 'professional tone';
  const mkPrompt = (chunk) => [
    `Rewrite the text below in a ${tn}.`,
    'Preserve meaning; return ONLY the rewritten text.',
    '',
    chunk
  ].join('\n');

  const chunks = splitTextIntoChunks(text);
  const outs = [];
  for (const c of chunks) {
    const out = geminiGenerateText(mkPrompt(c), {
      temperature: (typeof temperature === 'number' ? temperature : DEFAULT_TEMP),
      maxTokens:   (typeof maxTokens   === 'number' ? maxTokens   : DEFAULT_TOKENS * 2)
    });
    outs.push(out);
    Utilities.sleep(COOLDOWN_MS);
  }
  return outs.join(CHUNK_JOINER);
}

// === CUSTOM FUNCTIONS ===
function AI_TRANSLATE(text, lang, tone) {
  return callGeminiTranslate(text, lang, tone, DEFAULT_TEMP, DEFAULT_TOKENS);
}

function AI_TONE(text, tone) {
  return callGeminiTone(text, tone, DEFAULT_TEMP, DEFAULT_TOKENS);
}

// === SHEETS ACTIONS ===
function translateSelectionToRight() {
  const range = SpreadsheetApp.getActiveRange();
  if (!range) { SpreadsheetApp.getUi().alert('Please select one or more cells first.'); return; }
  const ui = SpreadsheetApp.getUi();
  const resp = ui.prompt('Translate Selection', 'Enter target language:', ui.ButtonSet.OK_CANCEL);
  if (resp.getSelectedButton() !== ui.Button.OK) return;
  const lang = resp.getResponseText().trim() || 'English';

  const values = range.getValues();
  const out = values.map(r => [callGeminiTranslate(r[0], lang)]);
  range.offset(0, 1).setValues(out);
  SpreadsheetApp.getActive().toast(`Translated ${values.length} row(s)`);
}

function toneSelectionToRight() {
  const range = SpreadsheetApp.getActiveRange();
  if (!range) { SpreadsheetApp.getUi().alert('Please select one or more cells first.'); return; }
  const ui = SpreadsheetApp.getUi();
  const resp = ui.prompt('Tone Conversion', 'Enter desired tone:', ui.ButtonSet.OK_CANCEL);
  if (resp.getSelectedButton() !== ui.Button.OK) return;
  const tone = resp.getResponseText().trim() || 'professional tone';

  const values = range.getValues();
  const out = values.map(r => [callGeminiTone(r[0], tone)]);
  range.offset(0, 1).setValues(out);
  SpreadsheetApp.getActive().toast(`Tone-converted ${values.length} row(s)`);
}

// === SAMPLE DATA ===
function insertSampleData() {
  const sheet = SpreadsheetApp.getActiveSheet();
  sheet.clearContents();
  sheet.getRange(1, 1, 6, 2).setValues([
    ['Original Text', 'Result'],
    ['Please submit your report by Friday.', ''],
    ['Welcome to our annual conference!', ''],
    ['The product helps students learn faster.', ''],
    ['Thank you for your feedback!', ''],
    ['Our platform supports over 10 languages.', '']
  ]);
}
