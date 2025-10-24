/**
 * Exercise 9 — AI Meeting Minutes Generator
 * Author: Laurence “Lars” Svekis
 * Google Sheets + Gemini REST API
 */

const GEMINI_API_KEY = ''; // or set in Script Properties
const GEMINI_MODEL = 'gemini-2.5-flash';
const GEMINI_BASE = 'https://generativelanguage.googleapis.com/v1beta';
const DEFAULT_TEMP = 0.3;
const DEFAULT_TOKENS = 1024;
const COOLDOWN_MS = 150;

// === MENU ===
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🧠 AI Tools')
    .addItem('Summarize Meeting', 'summarizeMeeting')
    .addItem('Open Sidebar', 'showSidebar')
    .addItem('Insert Sample Data', 'insertSampleMeetings')
    .addToUi();
}

function showSidebar() {
  try {
    const html = HtmlService.createHtmlOutputFromFile('Sidebar')
      .setTitle('AI Meeting Minutes Generator')
      .setWidth(400);
    SpreadsheetApp.getUi().showSidebar(html);
  } catch (e) {
    SpreadsheetApp.getUi().alert('❌ Failed to open sidebar: ' + e.message);
  }
}

function getApiKey() {
  return PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY') || GEMINI_API_KEY;
}

function callGemini(prompt, tokens = DEFAULT_TOKENS) {
  const apiKey = getApiKey();
  if (!apiKey) return '❌ Missing API key';
  const payload = {
    contents: [{ role: "user", parts: [{ text: prompt }] }],
    generationConfig: { temperature: DEFAULT_TEMP, maxOutputTokens: tokens, topP: 0.9, topK: 40 }
  };
  const url = `${GEMINI_BASE}/models/${encodeURIComponent(GEMINI_MODEL)}:generateContent?key=${encodeURIComponent(apiKey)}`;
  const res = UrlFetchApp.fetch(url, { method: 'post', contentType: 'application/json', payload: JSON.stringify(payload), muteHttpExceptions: true });
  const data = JSON.parse(res.getContentText());
  return data?.candidates?.[0]?.content?.parts?.[0]?.text || '⚠️ No response';
}

// Strip ```json fences and parse JSON safely
function safeJson(raw) {
  try {
    let s = String(raw || '').trim();
    s = s.replace(/^```(?:json)?\s*/i, '').replace(/```$/i, '').trim();
    return JSON.parse(s);
  } catch (e) { return null; }
}

function summarizeMeeting() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const range = sheet.getActiveRange();
  const values = range.getValues();
  const out = [];

  for (const row of values) {
    const transcript = String(row[0] || '').trim();
    if (!transcript) { out.push(['', '', '', '']); continue; }

    const prompt = [
      "You are an AI meeting assistant. Summarize the following meeting transcript into JSON.",
      "Output only JSON with these keys:",
      "{ \"summary\": string, \"key_points\": string[], \"action_items\": [{\"owner\": string, \"task\": string, \"due\": string}], \"follow_ups\": string[] }",
      "",
      transcript
    ].join('\n');

    const raw = callGemini(prompt, 1536);
    const parsed = safeJson(raw);
    if (!parsed) {
      out.push(['⚠️ Could not parse JSON', '', '', '']);
    } else {
      const summary = parsed.summary || '';
      const points = (parsed.key_points || []).map(p => `• ${p}`).join('\n');
      const actions = (parsed.action_items || []).map(a => `${a.owner || 'N/A'} — ${a.task || ''} (${a.due || 'no date'})`).join('\n');
      const follow = (parsed.follow_ups || []).join('\n');
      out.push([summary, points, actions, follow]);
    }
    Utilities.sleep(COOLDOWN_MS);
  }

  const startCol = range.getColumn() + 1;
  sheet.getRange(range.getRow(), startCol, out.length, 4).setValues(out);
  SpreadsheetApp.getActive().toast('✅ Meeting summaries generated');
}

function insertSampleMeetings() {
  const data = [
    ['Meeting Transcript', 'Summary', 'Key Points', 'Action Items', 'Follow Ups'],
    ['Today we discussed the Q4 marketing campaign. Sarah will handle the new social media plan, John will redesign the landing page by next week, and we agreed to review budget allocations on Friday.', '', '', '', ''],
    ['Engineering sync: Maria confirmed API integration done, Alex testing new endpoints, and QA will start regression testing tomorrow. Decision: Launch postponed to next Monday.', '', '', '', ''],
    ['Finance update: Reviewed Q3 numbers, expenses within limits. Next: finalize vendor payments by Thursday and prepare Q4 forecast.', '', '', '', '']
  ];
  const sheet = SpreadsheetApp.getActiveSheet();
  sheet.clearContents();
  sheet.getRange(1, 1, data.length, data[0].length).setValues(data);
  SpreadsheetApp.getActive().toast('✅ Sample meeting data inserted (A–E)');
}
