
# 🧠 Exercise 9 — AI Meeting Minutes Generator (Sheets + Gemini)

Generate meeting **summary, key points, action items, and follow-ups** directly from a transcript in Google Sheets.

## Files
- `Code.gs` — menu, API client, summarizer, sample data
- `Sidebar.html` — paste a transcript and summarize interactively
- `appsscript.json` — required scopes

## Setup
1. Open a Sheet → **Extensions → Apps Script**.
2. Create the three files above and paste their contents.
3. Add Script Property `GEMINI_API_KEY` with your Gemini key (Project Settings → Script properties).
4. Run `onOpen()` once to authorize → reload Sheet.

## Use
- **AI Tools → Insert Sample Data** to create sample rows in A:E
- Select column A (transcripts) → **AI Tools → Summarize Meeting**
- Or **AI Tools → Open Sidebar** to test manually

## Notes
- The parser strips ```json fences if the model returns Markdown.
- If parsing fails, the row will show a warning so you can retry.
