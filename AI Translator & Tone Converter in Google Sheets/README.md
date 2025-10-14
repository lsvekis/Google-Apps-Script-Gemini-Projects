
# 🌐 Exercise 4 — Google Sheets + Gemini AI Translator & Tone Converter (Updated)

Translate text and adjust tone directly inside Google Sheets using **Google Apps Script** and **Gemini**.

## 🚀 Features
- Custom menu: **🧠 AI Tools**
- Translate selection → next column
- Tone-convert selection → next column
- Sidebar UI with language + tone + temperature + token controls
- Custom functions: `=AI_TRANSLATE(A2, "French", "friendly tone")`, `=AI_TONE(A2, "concise professional tone")`
- **Robust extraction**, **chunking for long inputs**, **larger token budgets**

## 📦 Files
- `Code.gs` — Apps Script backend (robust Gemini client, chunking, menu, actions, utilities)
- `Sidebar.html` — Translator & tone sidebar
- `appsscript.json` — Manifest with required scopes
- `README.md` — This documentation

## ⚙️ Setup
1. Open a new Google Sheet → **Extensions → Apps Script**.
2. Create the files above and paste the contents.
3. Get a Gemini API key: https://aistudio.google.com/app/apikey
4. Add it to Script Properties as `GEMINI_API_KEY` (recommended), or set `GEMINI_API_KEY` in `Code.gs`.
5. In the editor, **Run → onOpen()** → authorize scopes.
6. Reload the Sheet; the **🧠 AI Tools** menu appears.

## 🧠 Usage
- **Insert Sample Data** from the menu to test.
- Select cells → **Translate Selection → next column** (choose target language).
- Select cells → **Tone-convert Selection → next column** (enter tone).
- Or open **Open Translator Sidebar**, paste text, choose options, and **Run**.
- Custom functions:
  - `=AI_TRANSLATE(A2, "Spanish", "friendly tone")`
  - `=AI_TONE(A2, "formal tone")`

## 🔧 Troubleshooting
- `❌ Set GEMINI_API_KEY…` → Add your API key in Script Properties.
- `⚠️ Truncated (MAX_TOKENS)…` → Raise max tokens (Sidebar control) or shorten input; chunking helps automatically.
- `❌ Error: Gemini HTTP 404` → Ensure the model exists for your key (e.g., `gemini-2.5-flash`).
- Sidebar won’t open → ensure appsscript.json includes `script.container.ui` and the project is bound to the Sheet.

## 🪪 License
MIT
