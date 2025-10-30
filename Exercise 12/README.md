# 🧠 Exercise 12 — AI Translator with Terminology Glossary (Sheets + Gemini)

Translate text in Google Sheets using Gemini and **enforce a terminology glossary** (Sheet "Glossary").

## Files
- `Code.gs` — menu, REST call, glossary-aware translation, custom function, sidebar entry
- `Sidebar.html` — quick single-shot translator UI (+ ping button)
- `appsscript.json` — required scopes

## Setup
1. Open a Google Sheet → **Extensions → Apps Script**.
2. Create the three files above and paste their contents.
3. Project Settings → **Script properties** → add `GEMINI_API_KEY` with your Gemini key.
4. Run `onOpen()` once to authorize and reload the Sheet.

## Use
- **AI Tools → Insert Sample Text** to populate A–B.
- **AI Tools → Insert Sample Glossary** to create a Glossary sheet with pairs in A:B.
- Select your source column → **AI Tools → Translate Selection → next column**.
- Or open the **Translator Sidebar** and run a single-shot translation.

## Custom Function
Use the glossary directly in a cell:
```
=AI_TRANSLATE(A2, "French", Glossary!A:B)
```

## Notes
- The glossary is enforced in the prompt. Tune it by editing "Glossary" A:B.
- Temperature kept low to limit rewording drift.
