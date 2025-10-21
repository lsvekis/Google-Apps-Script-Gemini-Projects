# 🌐 Exercise 7 — Web Article Summarizer & Citation Extractor (Patched)

Google Sheets + Apps Script + Gemini. Paste URLs → get **title, bullet summary, key quotes, entities, topics, reading time, and source URL**.

## Files
- `Code.gs` — patched pipeline with chunked 3-bullet summaries and resilient merge+retry
- `Sidebar.html` — one-click URL summarizer UI
- `appsscript.json` — required scopes
- `README.md` — this file

## Setup
1. Open a Google Sheet → **Extensions → Apps Script**.
2. Create the files above and paste their contents.
3. Add Script Property `GEMINI_API_KEY` with your Gemini key.
4. Run **onOpen()** once to authorize → reload the sheet.

## Use
- **🧠 AI Tools → Insert Sample URLs**
- Select URLs in column A → **🧠 AI Tools → Summarize URL Selection → next columns**
- Or open **Open Web Summarizer Sidebar** and paste one URL.
- Custom function: `=AI_URL_SUMMARY(A2)`

## Notes
- Some sites block automated fetches (those rows will show an error).
- If you still hit truncation: reduce `CHUNK_CHAR_LIMIT` or change chunk prompt to “2 bullets”.
