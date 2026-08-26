# Excel Reader

Upload any Excel or CSV file, ask a question in plain English, get back an exact answer.

```
"names in marketing department"  →  "Alice, Brian, Clara"
```

The core design rule the whole app is built around: **the AI is never allowed to invent a number.** Python always does the actual computation on your real data — the AI only translates your question into code, and later translates the result back into a sentence.

---

## 1. What it does

- Upload `.xlsx`, `.xls`, `.xlsm`, `.csv`, or `.tsv`
- Ask anything about the data in plain English — totals, filters, comparisons, summaries
- Get an answer that's either:
  - ✅ **Computed by Python (exact)** — a real number, calculated by code that actually ran on your data, or
  - ⚠️ **LLM analysis (verify numbers)** — a fallback guess, used only when the code path fails, and always labelled as such
- Export any uploaded file as clean, structured **JSON** — no AI required for this part
- Convert loose JSON ↔ readable text
- Automatically handles messy real-world spreadsheets: merged cells, multiple stacked tables on one sheet, missing headers

## 2. How it works — the pipeline

One question goes through three hops, not one:

```
Your question ──┐
                ├──▶  1. AI writes pandas code (never runs it)
Your data ──────┘              │
                                ▼
                     2. Python executes that code
                        on your FULL real dataset
                                │
                 ┌──────────────┴──────────────┐
                 ▼ succeeded                    ▼ failed twice
     3a. AI turns the raw           3b. AI answers directly from
         number into a sentence         the schema — no code, no
         (number is untouched)          guarantee — clearly labelled
                 │                                │
                 ▼                                ▼
        ✅ exact answer                  ⚠️ unverified answer
```

- If the AI's first attempt at code has a bug (wrong column name, typo, etc.), the error message is handed straight back to the AI to fix — it retries automatically up to 2 more times before giving up and falling back.
- The fallback path exists so the app never just crashes on a hard question — it always answers something, but it's honest about whether that answer was actually computed or just guessed.

## 3. Project structure

```
excel reader/
├── app.py              ← the Streamlit UI (upload box, chat, sidebar, tabs)
├── excel_parser.py      ← Excel/CSV → structured JSON. Pure Python, no AI.
├── llm_analyzer.py       ← the "brain": question → code → execute → explain
├── requirements.txt      ← Python dependencies
├── .env.example          ← template for your API key (copy to .env)
└── venv/                 ← local virtual environment (not shared)
```

| File | Responsibility | Needs an API key? |
|---|---|---|
| `app.py` | Renders the page, wires buttons/chat to the other two files | No |
| `excel_parser.py` | Reads every sheet, unmerges cells, detects headers, splits stacked tables, builds the JSON export | No |
| `llm_analyzer.py` | Loads data as pandas DataFrames, calls the AI to write/execute/format answers | Yes (Gemini or Claude) |

## 4. Setup

```bash
# 1. Create and activate a virtual environment
python -m venv venv
venv\Scripts\activate          # Windows
source venv/bin/activate       # macOS/Linux

# 2. Install dependencies
pip install -r requirements.txt

# 3. Set up your API key
copy .env.example .env         # Windows
cp .env.example .env           # macOS/Linux
# then edit .env and paste your real key
```

### Getting an API key (free option)

1. Go to [aistudio.google.com/apikey](https://aistudio.google.com/apikey)
2. Sign in with a Google account → **Create API Key** (no credit card required)
3. Paste it into `.env` as `GEMINI_API_KEY=...`, or directly into the sidebar of the running app

Anthropic Claude is also supported (`ANTHROPIC_API_KEY`) as a paid alternative — useful if Gemini's free daily quota runs out, since it's billed and rate-limited separately.

## 5. Running the app

```bash
streamlit run app.py
```

Opens at `http://localhost:8501`. Upload a file, and the sidebar API key field also works without touching `.env` at all — it's set live for that session.

## 6. Using it

**Upload a file** → the app immediately parses it (no AI call yet) and shows sheet names, row/column counts, and merged-cell counts.

**Ask a question** in the chat box at the bottom — this is the only part that spends API quota. Every answer is tagged:
- `✅ Computed by Python (exact)` — trust the number
- `⚠️ LLM analysis (verify numbers)` — double-check it

**Multi-table sheets are handled automatically.** If one worksheet contains several separate tables stacked with blank rows between them (a common pattern in dashboard-style exports), the app detects each blank-row boundary and treats them as separate tables — so filtering by "Marketing" in a *Department Sales* table never accidentally pulls in rows from an unrelated *Regional Revenue* table above it.

**JSON export** — go to the **"JSON ↔ Text"** tab → **"Export Parsed Excel"** → **⬇️ JSON**. This works with no API key and no quota cost; it's the same structured data used internally, just handed to you as a file.

**JSON ↔ Text converter boxes** (top of the same tab) are a *different*, AI-powered feature — they do need a key/quota, unlike the export button below them.

## 7. Why the numbers can be trusted

Every "exact" answer is the literal printed output of a real Python script that ran against your full, uploaded data — the AI never sees or reports a number it computed itself. It only:
1. Writes the pandas code (the *question*, in code form)
2. Rewrites Python's raw printed output into readable English (the *wording*, not the number)

If you ask for a specific format — e.g. *"...as json"* — the formatting step is instructed to return valid JSON instead of a sentence, while keeping the same underlying exact numbers.

## 8. Troubleshooting

**"It's giving obviously wrong answers after I fixed something"**
Streamlit keeps your uploaded file's parsed data in that browser session's memory. A code fix doesn't retroactively reprocess data that's already loaded — restart the Streamlit server and re-upload the file so it parses fresh.

**"API key quota exhausted"**
Gemini's free tier resets roughly every 24 hours. A second key on the *same* Google account shares the same quota — a different account (or enabling billing) gets a separate limit. Meanwhile, upload/parsing and the JSON export still work with zero API calls.

**"IndentationError" or the app won't start at all**
A Python syntax issue, unrelated to your data — check that recent edits to `llm_analyzer.py` / `excel_parser.py` didn't break indentation, then re-run `streamlit run app.py`.

## 9. Tech stack

- **UI:** Streamlit
- **Data:** pandas, openpyxl (Excel/CSV reading, merged-cell handling)
- **AI:** Google Gemini (free tier) or Anthropic Claude (paid) — either works, auto-detected from whichever API key is set
