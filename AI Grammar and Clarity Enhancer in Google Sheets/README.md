
# 🧠 Google Sheets + Gemini AI Grammar & Clarity Enhancer

Correct grammar, spelling, and clarity directly inside Google Sheets using **Google Apps Script** and **Gemini**.

## 🚀 Features
- Custom menu: **🧠 AI Tools**
- Sidebar UI for quick grammar checks
- One-click **Correct Selection → next column**
- Custom function: `=AI_GRAMMAR(A2)`
- Robust REST client with helpful error messages

## 📦 Files
- `Code.gs` — Apps Script backend (API + sheet logic)
- `Sidebar.html` — Sidebar UI
- `appsscript.json` — Manifest with OAuth scopes
- `README.md` — This documentation

## ⚙️ Setup
1. Open a new Google Sheet → **Extensions → Apps Script**.
2. Create files above and paste contents.
3. Get an API key from https://aistudio.google.com/app/apikey
4. Add it to `Code.gs` or store as Script Property `GEMINI_API_KEY`.
5. In Apps Script, **Run → onOpen()** → authorize.
6. Return to the Sheet and reload.

## 🧠 Usage
- **Insert Sample Data** from the menu to test.
- Select cells in column A → **Correct Selection → next column**.
- Or open **Open Grammar Sidebar** and paste text manually.
- Try formula: `=AI_GRAMMAR(A2)`

## 🔧 Troubleshooting
- `❌ Error: Gemini HTTP 404` → Use a supported model (e.g., `gemini-2.5-flash`).
- `⚠️ No corrections returned` → Increase `MAX_TOKENS` or shorten input.
- Sidebar fails to open → make sure manifest includes `script.container.ui` and project is **bound** to the Sheet.

## 🪪 License
MIT
