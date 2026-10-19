# Clockify × Jira Controlling

Turn Clockify time entries and Jira original estimates into an Excel workbook and a PDF report. Compare estimated and tracked hours by task, see project totals and time by person, and retain manual task notes and estimate overrides across refreshes.

## Requirements

- Python 3.10 or newer, with `venv` and pip
- Bash on Linux, macOS, or WSL
- A Clockify API key and workspace ID, with access to workspace reports
- A Jira Cloud account with permission to browse the relevant issues and read their original estimates
- A personal Atlassian API token for that account
- Internet access for installation and refreshes

## Practical setup

### 1. Download and install

```bash
git clone git@github.com:MayBeeGG/clockify-jira-controlling.git
cd clockify-jira-controlling
./setup.sh
cp .env.example .env
```

For HTTPS cloning, use `https://github.com/MayBeeGG/clockify-jira-controlling.git`. `setup.sh` creates a local `.venv` and installs the pinned packages in `requirements.txt`.

### 2. Configure your accounts

Edit the new `.env` file in a text editor:

```dotenv
CLOCKIFY_API_KEY=your_clockify_api_key
CLOCKIFY_WORKSPACE_ID=your_clockify_workspace_id
JIRA_URL=https://example.atlassian.net
JIRA_PROJECT_KEY=DEMO
JIRA_EMAIL=your_account@example.com
JIRA_API_TOKEN=your_personal_atlassian_api_token
```

Replace every example value with your own. Keep this file private.

- **Clockify:** obtain your API key from your Clockify profile settings. Use the workspace ID belonging to the workspace you want to report on, rather than its display name. The script queries its projects, tasks, users, and detailed time reports; your account must have the necessary access.
- **Jira URL:** use only your HTTPS site origin, such as `https://example.atlassian.net`, without an issue path or `/jira` suffix.
- **Jira email and token:** use the email of the Atlassian account that owns the personal API token. This implementation uses Basic authentication against your site URL. Use a personal Jira Cloud API token without scopes for this flow; an organization admin API key is not a substitute.
- **Jira project:** set the project key, such as `DEMO`, to narrow the search. Leaving it empty searches all issues accessible to the account and can cause ambiguous matches across projects.

Both Clockify and Jira configuration are required by this version. The script checks Jira authentication before writing reports. A Jira HTTP 401 generally means the account email/token pair needs correction or replacement.

### 3. Align task names and enter Jira estimates

The match is **Clockify task name ↔ Jira issue summary**. Clockify project names, entry descriptions, and Jira issue keys do not determine the match.

For each task you want to compare:

1. Give the Clockify task the same name as the Jira issue summary.
2. Log completed Clockify time entries against that task and project.
3. Set the Jira issue's **Original estimate**. The script reads `timeoriginalestimate` and converts seconds to hours.
4. Stop running timers, then refresh the report for the desired date range.

Names are normalized by trimming and collapsing whitespace and ignoring letter case. Everything else must match; there is no fuzzy matching or issue-key parsing.

| Clockify task | Jira summary | Result |
| --- | --- | --- |
| `Build login page` | `Build login page` | Match |
| ` build   LOGIN page ` | `Build login page` | Match |
| `DEMO-12 Build login page` | `Build login page` | No match |
| `Build login page` | Two issues named `Build login page` | Ambiguous; no estimate imported |

There must be exactly one matching issue within the Jira search results. A missing match, duplicate summary, or missing original estimate leaves the imported estimate empty. Check the Dashboard's **Jira match** column for the reason or matched issue key. Clockify estimates are not used. Time entries without a task have no task name available for Jira matching.

### 4. Generate and open reports

```bash
./refresh_clockify.sh 2026-01-01
./refresh_clockify.sh 2026-01-01 --until 2026-03-31
```

Dates use `YYYY-MM-DD`. Without arguments, the start is `2026-01-01`; the end defaults to today. Choose the start explicitly if you need a different year or reporting window.

The first run builds `Clockify_Controlling.xlsx` from scratch and writes `Clockify_Controlling.pdf` alongside it. No template or existing export is needed. Subsequent runs refresh those files. Open the workbook in Excel or LibreOffice to recalculate formulas.

The workbook contains:

- **Dashboard:** tasks, Jira original estimates, tracked and remaining hours, match status, and top contributor per task.
- **People:** tracked, billable, and non-billable time per person, plus task details.
- **Time entries:** individual time entries, including dates and descriptions.

The PDF summarizes projects, tasks, and people. Its task difference is **tracked hours minus Jira original estimate**: a positive number means an overrun. The PDF comparison always uses the Jira estimate, even when the workbook has a manual override.

## Which Excel edits survive a refresh?

Use the two designated Dashboard columns:

| Edit | Preserved? | Behavior |
| --- | --- | --- |
| **I — Priority / note** cell text | Yes | Restored for the same Clockify project ID and task ID |
| **L — Estimate override (h)** numeric value | Yes | Restored for the same IDs; takes precedence in workbook remaining-hours calculations |
| Excel cell comments or threaded comments on generated table rows | No | The generated rows are deleted and rebuilt |
| Other edits inside generated task, people, or time-entry tables | No | Replaced by fresh data and formulas |

**Type your note directly into the Priority / note cell; using Excel's “New Comment” or “New Note” feature is different and is not preserved on regenerated rows.** Save and close the workbook before refreshing. Only saved values in the workbook being refreshed can be read.

Before rebuilding the tables, the script reads the two manual columns and stores their values under `(project ID, task ID)`. It then restores them onto the matching task rows. Sorting or renaming a task does not break this mapping as long as the IDs remain the same. Do not edit the hidden ID columns or move/rename the generated sheets and columns.

Deleting and recreating a Clockify task gives it a new ID, so its old note will not follow. If a task is absent from a refreshed report, its manual values are not retained in a separate history store. Keep private workbook backups if those notes matter. Changing `--workbook` to a new file starts with no saved manual values. Notes and overrides are stored in the workbook, not uploaded to Clockify or Jira, and are not included as note fields in the PDF.

## Direct Python usage

```bash
.venv/bin/python sync_clockify_controlling.py --since 2026-01-01 --until 2026-03-31
.venv/bin/python sync_clockify_controlling.py --help
```

Options include `--workbook PATH`, `--pdf PATH`, and `--env-file PATH`. Output parent directories must already exist. When `--pdf` is omitted, the PDF uses the workbook's path with a `.pdf` extension.

Python can also read configuration from environment variables. Existing environment variables take precedence over values in `.env`. The shell launcher requires a local `.env` file; use Python directly for an environment-only setup.

## Files and privacy

- `sync_clockify_controlling.py`: API requests, workbook creation/refresh, and PDF generation
- `setup.sh`: virtual environment and dependency installation
- `refresh_clockify.sh`: refresh launcher
- `requirements.txt`: pinned Python dependencies
- `.env.example`: blank configuration template
- `.gitignore`: excludes credentials, local environments, caches, and generated reports

This repository includes no credentials or real exports. `.env` and generated Excel/PDF/CSV/TSV files are ignored by Git because they can contain credentials, names, descriptions, hours, notes, and private business data. Do not force-add them. API credentials are not written into the reports.
