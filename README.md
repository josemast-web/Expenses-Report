# QB Expense Report Bot

Disclaimer: This repository is the public version of an original repository that contains private information.

A Python automation that fetches the latest **QuickBooks Excel report** from a Google Drive folder, processes expense data, and delivers a formatted **HTML email report** with plain-text attachments — triggered on demand or via an external scheduler through GitHub Actions.

---

## Architecture

```
Google Drive Folder
      |
      | Drive API v3 (service account)
      | List files -> find latest by date in filename
      | Download -> temp_qb_report.xlsx
      v
  DriveManager
      |
      v
  DataProcessor
  - Read Excel (header at row 4)
  - Exclude configured memo terms
  - Parse dates, amounts, quantities
  - Slice by time window (current week / last 30 days)
      |
      v
  ReportRenderer
  - HTML sections: KPI cards, Top-10 tables, high-value alert
  - TXT attachments: weekly list + monthly project breakdown
      |
      v
  EmailService
  - Gmail SMTP (TLS port 587)
  - HTML body + TXT file attachments
```

---

## Report Sections

### HTML Email
- **KPI Cards** — Total Spend, Transaction count, Vendor count, Active Projects
- **Top 10 Vendors** — sorted by spend
- **Top 10 Projects** — sorted by spend
- **Unassigned Alert** — highlights transactions with no Customer/Project
- **High-Value Alert** — lists individual purchases above the configured threshold

### TXT Attachments
| File | Content |
|---|---|
| `Weekly_Transactions_List.txt` | Line-by-line transaction log for the current week |
| `Monthly_By_Project.txt` | Transactions grouped by Project and ISO week number |

---

## Expected File Naming Convention

The bot searches the Drive folder for files matching:

```
QB*report*MMDDYYYY*   (case-insensitive)
```

Example: `QB_report_03152025.xlsx`

The file with the most recent date in the filename is downloaded.

---

## Setup

### 1. Clone and install

```bash
git clone https://github.com/your-username/your-repo.git
cd your-repo
python -m venv .venv
source .venv/bin/activate   # Windows: .venv\Scripts\activate
pip install -r requirements.txt
```

### 2. Create a GCP service account

1. Go to [console.cloud.google.com](https://console.cloud.google.com) > **IAM & Admin > Service Accounts**
2. Create a service account and download its JSON key as `credentials.json`
3. Enable the **Google Drive API** for your project
4. Share the target Drive folder with the service account's email address (Viewer access)

### 3. Configure environment variables

```bash
cp .env.example .env
# Edit .env with your values
```

### 4. Run locally

```bash
# Load .env variables (Linux/macOS)
export $(grep -v '^#' .env | xargs)

python main.py
```

---

## Environment Variables Reference

| Variable | Description |
|---|---|
| `DRIVE_FOLDER_ID` | Google Drive folder ID (from the folder URL) |
| `GCP_CREDENTIALS_FILE` | Path to the service account JSON (default: `credentials.json`) |
| `EMAIL_SENDER` | Gmail address used to send reports |
| `EMAIL_PASSWORD` | [Gmail App Password](https://myaccount.google.com/apppasswords) |
| `EMAIL_RECIPIENTS` | Comma-separated recipient addresses |
| `REPORT_LABEL` | Footer label in the HTML email |
| `HIGH_VALUE_THRESHOLD` | USD amount above which purchases trigger an alert (default: `3000`) |

---

## GitHub Actions Deployment

Add all variables above as **repository secrets** (`Settings > Secrets and variables > Actions`), plus:

| Secret | Description |
|---|---|
| `GCP_CREDENTIALS` | Full content of `credentials.json` as a single secret string |

The workflow supports two triggers:

- **`workflow_dispatch`** — manual run from the GitHub Actions UI
- **`repository_dispatch`** (event type `run-expense-bot`) — triggered via the GitHub API, useful for connecting an external cron service for precise scheduling

```bash
# Example: trigger via curl
curl -X POST \
  -H "Authorization: token YOUR_PAT" \
  -H "Accept: application/vnd.github.v3+json" \
  https://api.github.com/repos/<owner>/<repo>/dispatches \
  -d '{"event_type":"run-expense-bot"}'
```

---

## Project Structure

```
.
├── main.py                          # Full pipeline (Drive -> Process -> Report -> Email)
├── requirements.txt
├── .env.example                     # Environment variable reference
├── .gitignore                       # Excludes credentials.json and .env
└── .github/
    └── workflows/
        └── expense_bot.yml          # GitHub Actions workflow
```

---

## Key Design Decisions

- **Single-file architecture** — the entire pipeline lives in `main.py` with four cohesive classes (`DriveManager`, `DataProcessor`, `ReportRenderer`, `EmailService`). This keeps the project easy to read and deploy without a package structure.
- **Date-based file selection** — the bot parses dates from filenames rather than relying on Drive metadata (modification time), which is more resilient to re-uploads or file copies.
- **Configurable exclusions** — memo-based exclusion terms (shipping, taxes, surcharges) are stored in `Config.EXCLUDE_TERMS` and can be extended without touching report logic.
- **No hard-coded identifiers** — all Drive folder IDs, email addresses, and thresholds are injected at runtime via environment variables, making the codebase fully portable.
- **Gmail-safe HTML** — all report styles use inline CSS; no external stylesheets that email clients would strip.
