/**
 * Exercise 7 — Web Article Summarizer & Citation Extractor (Sheets + Gemini REST)
 * (Patched to avoid MAX_TOKENS truncation and add resilient merge w/ retry)
 *
 * What this file gives you:
 * - Menu items (onOpen): Open Sidebar, Summarize URL Selection → next columns, Insert Sample URLs
 * - Custom function: =AI_URL_SUMMARY(A2)
 * - Robust pipeline:
 *     fetchUrlText → chunk (short) → per-chunk 3-bullet summaries → merge w/ big token budget
 *     → retry with minimal JSON if needed → normalized result to sheet
 *
 * Tip: Store your API key in Script Properties as GEMINI_API_KEY. (Editor → Project Settings → Script properties)
 */

// ===== CONFIG =====
const GEMINI_API_KEY = '';               // Optional fallback—prefer Script Properties (see getApiKey)
const GEMINI_MODEL   = 'gemini-2.5-flash';
const BASE_URL       = 'https://generativelanguage.googleapis.com/v1beta';

const DEFAULT_TEMP     = 0.2;            // low creativity for factual accuracy
const DEFAULT_TOKENS   = 640;            // ↓ slightly smaller per-chunk budget
const MERGE_TOKENS     = 1536;           // ↑ much larger budget for the merge step
const COOLDOWN_MS      = 150;

const CHUNK_CHAR_LIMIT = 5500;           // ↓ smaller chunks → less to merge overall
const CHUNK_JOINER     = '\\n\\n';

const CACHE_SEC        = 6 * 3600;       // 6 hours for fetched pages cache

// ===== MENU & UI =====
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🧠 AI Tools')
    .addItem('Open Web Summarizer Sidebar', 'showWebSummarizerSidebar')
    .addItem('Summarize URL Selection → next columns', 'summarizeUrlSelectionToRight')
    .addSeparator()
    .addItem('Insert Sample URLs', 'insertSampleUrls')
    .addToUi();
}

function showWebSummarizerSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('Sidebar')
    .setTitle('Gemini Web Summarizer');
  SpreadsheetApp.getUi().showSidebar(html);
}

// ===== API KEY =====
function getApiKey() {
  const p = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  return (p && p.trim()) ? p.trim() : (GEMINI_API_KEY || '').trim();
}

// ===== GEMINI HELPERS =====
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
      return t ? `${t}\\n\\n⚠️ Truncated (MAX_TOKENS).` : '⚠️ Truncated (MAX_TOKENS) and no visible text.';
    }

    const parts = cand?.content?.parts || [];
    const text = parts.map(p => p?.text || '').join('').trim();
    if (text) return text;

    const cText = cand?.text;
    if (cText && String(cText).trim()) return String(cText).trim();

    return '';
  } catch (e) { return ''; }
}

function safeParseJson(maybeJson) {
  try {
    if (!maybeJson) return null;
    return JSON.parse(maybeJson);
  } catch (e) {
    // Salvage the first {...} block if extra prose leaks in
    const m = String(maybeJson).match(/\\{[\\s\\S]*\\}/);
    if (!m) return null;
    try { return JSON.parse(m[0]); } catch (e2) { return null; }
  }
}

function geminiGenerateText(prompt, maxTokens) {
  const apiKey = getApiKey();
  if (!apiKey) return '❌ Set GEMINI_API_KEY in Script Properties or Code.gs.';

  const payload = {
    contents: [{ role: "user", parts: [{ text: prompt }] }],
    generationConfig: {
      temperature: DEFAULT_TEMP,
      maxOutputTokens: maxTokens || DEFAULT_TOKENS,
      topP: 0.95,
      topK: 40
    }
  };

  const url = `${BASE_URL}/models/${encodeURIComponent(GEMINI_MODEL)}:generateContent?key=${encodeURIComponent(apiKey)}`;
  const options = { method: 'post', contentType: 'application/json', payload: JSON.stringify(payload), muteHttpExceptions: true };

  try {
    const res  = UrlFetchApp.fetch(url, options);
    const code = res.getResponseCode();
    const body = res.getContentText();

    if (code < 200 || code >= 300) return `❌ Error: Gemini HTTP ${code}. ${body}`;

    const text = extractGeminiText(JSON.parse(body));
    return text || '⚠️ No output (see Logs).';
  } catch (e) {
    return '❌ Error: ' + e.message;
  }
}

// ===== FETCH & CLEAN HTML =====
function fetchUrlText(url) {
  const cache = CacheService.getScriptCache();
  const key = 'fetch:' + url;
  const cached = cache.get(key);
  if (cached) return cached;

  let res, code, body;
  try {
    res = UrlFetchApp.fetch(url, {
      method: 'get',
      muteHttpExceptions: true,
      followRedirects: true,
      validateHttpsCertificates: true,
      headers: { 'User-Agent': 'AppsScript-GeminiSummarizer/1.0' }
    });
    code = res.getResponseCode();
    body = res.getContentText();
  } catch (e) {
    return '';
  }

  if (code < 200 || code >= 400) return '';

  // Remove scripts/styles/noscript/comments
  body = body.replace(/<script[\\s\\S]*?<\\/script>/gi, ' ')
             .replace(/<style[\\s\\S]*?<\\/style>/gi, ' ')
             .replace(/<noscript[\\s\\S]*?<\\/noscript>/gi, ' ')
             .replace(/<!--[\\s\\S]*?-->/g, ' ');

  // Attempt to capture title
  const titleMatch = body.match(/<title[^>]*>([\\s\\S]*?)<\\/title>/i);
  const t = titleMatch ? 'TITLE: ' + titleMatch[1].trim() + '\\n\\n' : '';

  // Strip remaining tags and normalize whitespace
  const text = t + body.replace(/<[^>]+>/g, ' ')
                       .replace(/\\s+/g, ' ')
                       .replace(/&nbsp;/g, ' ')
                       .trim();

  cache.put(key, text, CACHE_SEC);
  return text;
}

// ===== CHUNKING =====
function splitTextIntoChunks(input, limit = CHUNK_CHAR_LIMIT) {
  const text = String(input || '');
  if (text.length <= limit) return [text];

  const parts = [];
  for (let i = 0; i < text.length; i += limit) {
    parts.push(text.slice(i, i + limit));
  }
  return parts;
}

// ===== MERGE HELPERS (PATCHED) =====
function looksTruncated_(text) {
  return /Truncated \\(MAX_TOKENS\\)/i.test(String(text || ''));
}

/**
 * Try to merge chunk bullet summaries into strict JSON.
 * 1st attempt: full schema (title, summary, key_quotes, entities, topics, reading_time_minutes, source_url)
 * Retry (if truncated/invalid): minimal schema (title, summary, topics, reading_time_minutes, source_url)
 */
function mergeChunkSummaries_(chunkSummaries, url) {
  // Attempt 1 — full schema
  const mergePrompt = [
    'Consolidate these bullet summaries from one web page into STRICT JSON:',
    '{ "title": string, "summary": string[], "key_quotes": string[], "entities": { "people": string[], "organizations": string[], "locations": string[] }, "topics": string[], "reading_time_minutes": number, "source_url": string }',
    'Rules: 3–7 bullets; up to 3 short verbatim quotes; 3–6 lowercase topics; integer minutes.',
    `Set "source_url" to: ${url}`,
    '',
    'Chunk bullets:',
    chunkSummaries.join('\\n\\n')
  ].join('\\n');

  let jsonText = geminiGenerateText(mergePrompt, MERGE_TOKENS);
  let obj = safeParseJson(jsonText);
  if (obj && !looksTruncated_(jsonText)) return obj;

  // Attempt 2 — minimal schema (shorter to avoid truncation)
  const retryPrompt = [
    'Return STRICT JSON with keys exactly:',
    '{ "title": string, "summary": string[], "topics": string[], "reading_time_minutes": number, "source_url": string }',
    'Rules: summary 3–5 bullets, topics 3–5 lowercase tags, integer minutes.',
    `Set "source_url" to: ${url}`,
    '',
    'Chunk bullets:',
    chunkSummaries.join('\\n\\n')
  ].join('\\n');

  jsonText = geminiGenerateText(retryPrompt, MERGE_TOKENS);
  obj = safeParseJson(jsonText);
  if (!obj) return null;

  // Normalize to include entities & key_quotes fields even if minimal response
  obj.summary = Array.isArray(obj.summary) ? obj.summary : [];
  obj.topics  = Array.isArray(obj.topics)  ? obj.topics  : [];
  if (!obj.entities) obj.entities = { people: [], organizations: [], locations: [] };
  if (!Array.isArray(obj.entities.people))        obj.entities.people = [];
  if (!Array.isArray(obj.entities.organizations)) obj.entities.organizations = [];
  if (!Array.isArray(obj.entities.locations))     obj.entities.locations = [];
  if (!Array.isArray(obj.key_quotes))             obj.key_quotes = [];

  return obj;
}

// ===== SUMMARIZATION PIPELINE =====
function summarizeUrl(url) {
  const raw = fetchUrlText(url);
  if (!raw) return { error: '❌ Could not fetch URL or empty content.', source_url: url };

  // 1) Per-chunk summarization — ask for *short* bullets to reduce merge size
  const chunks = splitTextIntoChunks(raw);
  const chunkSummaries = [];
  for (let i = 0; i < chunks.length; i++) {
    const p = [
      'Summarize this chunk in 3 concise fact-only bullets. No preface; bullets only.',
      '',
      chunks[i]
    ].join('\\n');

    const out = geminiGenerateText(p, DEFAULT_TOKENS);
    chunkSummaries.push(out);
    Utilities.sleep(COOLDOWN_MS);

    // keep bounded—summarize first 4 chunks max
    if (i >= 3) break;
  }

  // 2) Merge all bullet summaries into structured JSON (with retry)
  const obj = mergeChunkSummaries_(chunkSummaries, url);
  if (!obj) {
    return {
      error: '⚠️ Could not parse consolidated JSON after retry.',
      raw: '(truncated or invalid)',
      source_url: url
    };
  }

  // Normalize fields
  const title = String(obj.title || '').trim();
  const summary = Array.isArray(obj.summary) ? obj.summary : [];
  const key_quotes = Array.isArray(obj.key_quotes) ? obj.key_quotes : [];
  const entities = obj.entities || {};
  const people = Array.isArray(entities.people) ? entities.people : [];
  const orgs   = Array.isArray(entities.organizations) ? entities.organizations : [];
  const locs   = Array.isArray(entities.locations) ? entities.locations : [];
  const topics = Array.isArray(obj.topics) ? obj.topics : [];

  // Better minutes fallback based on words in bullet summaries (~200 wpm)
  const minutes = Number(obj.reading_time_minutes) ||
                  Math.max(1, Math.round((chunkSummaries.join(' ').split(/\\s+/).length) / 200));

  return {
    title,
    summary,
    key_quotes,
    entities: { people, organizations: orgs, locations: locs },
    topics,
    reading_time_minutes: minutes,
    source_url: url
  };
}

// ===== CUSTOM FUNCTION =====
// Use in a cell: =AI_URL_SUMMARY(A2)
function AI_URL_SUMMARY(url) {
  const r = summarizeUrl(String(url || ''));
  return JSON.stringify(r);
}

// ===== SHEETS ACTION =====
function summarizeUrlSelectionToRight() {
  const range = SpreadsheetApp.getActiveRange();
  if (!range) { SpreadsheetApp.getUi().alert('Please select one or more URL cells first.'); return; }

  const values = range.getValues();
  const out = []; // [Title, SummaryJoined, Entities, Quotes, Topics, Minutes, Source]

  for (let i = 0; i < values.length; i++) {
    const url = String(values[i][0] || '').trim();
    if (!/^https?:\\/\\//i.test(url)) {
      out.push(['Invalid URL', '', '', '', '', '', url]);
      continue;
    }

    const r = summarizeUrl(url);
    if (r.error) {
      out.push([r.error, '', '', '', '', '', url]);
    } else {
      const entitiesJoined = [
        (r.entities.people || []).join(', '),
        (r.entities.organizations || []).join(', '),
        (r.entities.locations || []).join(', ')
      ].filter(Boolean).join(' | ');

      out.push([
        r.title || '(untitled)',
        (r.summary || []).join(' • '),
        entitiesJoined,
        (r.key_quotes || []).map(q => `“${q}”`).join(' '),
        (r.topics || []).join(', '),
        r.reading_time_minutes || '',
        r.source_url || url
      ]);
    }

    Utilities.sleep(COOLDOWN_MS);
  }

  const sheet = range.getSheet();
  const startRow = range.getRow();
  const nextCol  = range.getColumn() + range.getNumColumns();
  sheet.getRange(startRow, nextCol, out.length, 7).setValues(out);

  if (startRow === 1) {
    sheet.getRange(1, nextCol, 1, 7).setValues([[
      'Title', 'Summary (bullets)', 'Entities (people | orgs | locs)', 'Key Quotes', 'Topics', 'Reading Minutes', 'Source URL'
    ]]);
  }

  SpreadsheetApp.getActive().toast(`Summarized ${out.length} URL(s) → ${colToLetter(nextCol)}:${colToLetter(nextCol+6)}`);
}

// ===== SAMPLE DATA =====
function insertSampleUrls() {
  const sheet = SpreadsheetApp.getActiveSheet();
  sheet.clearContents();
  sheet.getRange(1, 1, 6, 2).setValues([
    ['URL', 'Output →'],
    ['https://blog.google/technology/ai/ai-announcements/', ''],
    ['https://workspace.google.com/blog/product-announcements', ''],
    ['https://developers.google.com/apps-script', ''],
    ['https://ai.google.dev/gemini-api', ''],
    ['https://aistudio.google.com', '']
  ]);
  SpreadsheetApp.getActive().toast('✅ Sample URLs inserted (A–B).');
}

// ===== UTIL =====
function colToLetter(col) {
  let temp = '', letter = '';
  while (col > 0) {
    temp = (col - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    col = (col - temp - 1) / 26;
  }
  return letter;
}
