# AI CV Screener — Google Apps Script Template

> **100% AI-generated, 100% human-directed.**
> Directed and iterated by **Yordhan Fitrians Akhmad B.** based on real HR pain points — refined repeatedly for business fit, user experience, and workflow comfort. Published as a generic, reusable template.

## What It Does

An AI-powered CV screening tool that runs inside Google Sheets. It reads candidate data from a Google Form response sheet, pulls CVs from Google Drive, and uses the Claude API (Anthropic) to evaluate each candidate against your job descriptions.

**Features:**
- **AI screening** — Claude reads each CV (PDF, DOCX, Google Docs, images) and evaluates role fit, experience, salary expectations, and JD match
- **Job Description knowledge base** — add your JDs to a sheet tab; the AI screens candidates against your *specific* posted requirements, not just generic knowledge
- **Duplicate detection** — flags double-submissions (same email + same role) without burning API calls
- **Batch or auto mode** — screen 10 at a time manually, or let it run all candidates in the background using time-based triggers
- **Color-coded verdicts** — 4 (Qualified), 3 (Need Consideration), 2 (Not a Strong Fit), 1 (Big No), 0 (Duplicate)
- **Resume-safe** — skips rows that already have a verdict, so you can stop and restart without re-screening

## Setup

### 1. Create the Script
1. Open your Google Sheet (the one with form responses)
2. Go to **Extensions → Apps Script**
3. Delete the default code → paste the contents of `ai-cv-screener.gs` → **Save**
4. Reload the spreadsheet — the **🤖 AI CV Screener** menu appears

### 2. Run Setup
Click **🤖 AI CV Screener → ⚙️ Setup** and follow the 4-step wizard:

| Step | What to enter |
|------|---------------|
| 1. API Key | Your Anthropic API key (`sk-ant-...`) |
| 2. Company | Your company name + 1-2 sentence description (gives the AI context) |
| 3. Sheet names | The tab name for candidate responses + the tab name for JDs |
| 4. Column mapping | Which columns contain Email, LinkedIn, CV link, Role, Experience, Current Salary, Expected Salary (and where to write Verdict/Notes) |

### 3. Add Job Descriptions (optional but recommended)
Create a sheet tab (default name: `Job Desc Knowledge`) with:
- **Column A**: Role title (must match the dropdown values in your form)
- **Column B**: The full job description / ad text you posted

Use **📋 Check JD Coverage** to verify which roles are matched.

### 4. Screen Candidates
- **▶ Screen Batch** — processes 10 candidates, asks for confirmation
- **🚀 Screen ALL** — runs through all unscreened candidates automatically in the background
- **⏹ Stop** — halts auto-screening (preserves all completed results)
- **📊 View Progress** — shows counts per verdict category

## Important Caveats

- **Salary benchmarks are generic.** The system prompt tells Claude to use its market knowledge. If you need specific salary ranges for your market, edit the system prompt in `callClaude_()`.
- **LinkedIn profiles cannot be read.** LinkedIn blocks programmatic access. The script only checks whether a URL was provided.
- **CV access requires sharing.** The script can only read files that are accessible to the Google account running the script. Files with restricted sharing will show "Access denied."
- **Some data-source logic is deliberately simplified.** The author works with confidential HR data and cannot publish production-specific schemas, column layouts, or screening criteria. This template uses generic placeholders — adapt them to your own spreadsheet structure via the Setup wizard.
- **API costs.** Each candidate = 1 Claude API call. At Haiku pricing (~$0.005/candidate), 1,000 candidates ≈ $5. Costs increase slightly when JD context is included.

## Column Mapping Defaults

These match a typical Google Form response layout. Change via ⚙️ Setup if yours differs.

| Column | Position | Content |
|--------|----------|---------|
| C (3)  | Email Address |
| D (4)  | LinkedIn URL |
| E (5)  | CV / Portfolio (Drive link) |
| F (6)  | Role applied for |
| G (7)  | Years of experience |
| H (8)  | Current salary |
| I (9)  | Expected salary |
| M (13) | 🤖 Verdict (output) |
| N (14) | 🤖 AI Notes (output) |

## License

MIT — use it, modify it, share it.

## Attribution

```
100% AI-generated, 100% human-directed.
Directed and iterated by Yordhan Fitrians Akhmad B.
```
