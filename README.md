# Zendesk to JSM Migration Script

Production-ready Python tooling to migrate ~10,000 support tickets from
**Zendesk** to **Jira Service Management (JSM) Cloud** using the `requests`
library.

## Features

- **Cursor-based incremental export** from Zendesk with **side-loaded users**
  to avoid N+1 metadata calls.
- **Idempotent + resumable**: a local SQLite database records each migrated
  ticket, its new Jira issue key, and the last successful `after_cursor`. Crash
  or stop the script and re-run; it skips done tickets and resumes from the
  saved cursor.
- **Rate-limit aware**: intercepts `429 Too Many Requests`, honours the
  `Retry-After` header, and retries automatically (plus 5xx backoff).
- **HTML to ADF**: converts Zendesk HTML/plain text into valid Atlassian
  Document Format before creating issues and comments.
- **Full per-ticket flow**: ensure customer -> create issue -> migrate comments
  in chronological order -> download & re-upload attachments -> optional status
  transition.
- **Unmapped-value summary** printed at the end of every run so mapping gaps
  surface during your pilot.

## Project layout

```
.
├── README.md
├── requirements.txt
├── .env.example
├── .gitignore
├── migrate.py            # main migration engine
├── build_mappings.py     # CSV -> mappings.xlsx
└── mappings/
    ├── priority.csv
    ├── status.csv
    ├── tags.csv
    ├── request_type.csv
    └── config.csv
```

## Setup

```bash
python -m venv .venv && source .venv/bin/activate
pip install -r requirements.txt
cp .env.example .env   # then fill in real credentials
```

## Mappings

The **source of truth** is the CSV files in `mappings/` (Git-diff friendly).
Each CSV has a header row (ignored at load time); column A is the Zendesk value,
column B is the JSM value.

| Tab / CSV       | Maps                                         | Example |
|-----------------|----------------------------------------------|---------|
| `priority`      | Zendesk priority -> JSM priority name        | `urgent -> Highest` |
| `status`        | Zendesk status -> JSM status name            | `solved -> Done` |
| `tags`          | Zendesk tag -> JSM label (unmapped pass through) | `bug -> defect` |
| `request_type`  | Zendesk product/value -> JSM request type id (optional) | `software -> 10001` |
| `config`        | key/value defaults                           | `transition_status -> true` |

`config` keys: `default_priority`, `default_status`, `transition_status`
(whether to move the issue to its mapped status after creation).

Generate the single workbook (optional; the loader also reads the CSVs directly):

```bash
python build_mappings.py            # writes mappings.xlsx
```

Point `MAPPINGS_PATH` in `.env` at either `mappings.xlsx` or the `mappings/`
folder.

## Run

```bash
python migrate.py
```

Progress is logged to the console and to `migration.log`. State is stored in
`migration_state.db`. Re-running after an interruption resumes automatically.

At the end of each run an **unmapped value summary** lists any Zendesk
priorities, statuses, or request types that fell back to defaults so you can
top up the mapping files before the full migration.

## Notes & caveats

- The first Zendesk comment usually duplicates the ticket description and is
  skipped as a comment (its attachments are still migrated).
- Customer creation via `/rest/servicedeskapi/customer` is treated as
  idempotent; existing customers (400/409) are tolerated.
- One failing ticket or attachment is logged and skipped so it never aborts the
  remaining migration.
- Run a small pilot first (e.g. a Zendesk view of a few tickets) and review the
  unmapped summary before the full ~10k run.

## Wrike: Data Migration tasks

Suggested subtasks to create in the Wrike **Data Migration** task, nested under
the two existing items. These align the migration script's mapping files and
run flow with the project plan.

### Under "Verify Products/Subproducts/Tags mapping"

- [ ] Export the full list of Zendesk products, subproducts, and tags actually in use
- [ ] Decide which tags become JSM **labels** vs. which map to **request types** / fields
- [ ] Populate `mappings/tags.csv` (Zendesk tag -> JSM label; unmapped pass through unchanged)
- [ ] Populate `mappings/request_type.csv` (Zendesk product/value -> JSM request type id) for queue routing
- [ ] Confirm each JSM request type id with the Portal **Request Forms** / **Setup Queues** work
- [ ] Run a pilot migration and review the **unmapped-value summary** for missed tags / products
- [ ] Top up the CSVs for any misses and re-run the pilot until the summary is clean

### Under "Build a field mapping matrix (Zendesk > JSM)"

- [ ] List every Zendesk ticket field (standard + custom) and its JSM target field
- [ ] Populate `mappings/priority.csv` (e.g. `low->Low`, `normal->Medium`, `high->High`, `urgent->Highest`)
- [ ] Populate `mappings/status.csv` and confirm each JSM status name has a valid **transition** (status move on create)
- [ ] Set `config.csv` defaults: `default_priority`, `default_status`, `transition_status`
- [ ] Confirm `JSM_ISSUE_TYPE` and `JSM_PROJECT_KEY` match the target project scheme
- [ ] Decide requester/reporter handling and align with the **Import users from Zendesk** task
- [ ] Define the ticket-history cut-off date (links to **Import ticket history from Zendesk**)
- [ ] Run `python build_mappings.py` to generate `mappings.xlsx` and attach it to the Wrike task for sign-off

### New task: Environment & credentials setup

- [ ] Generate a Zendesk API token and confirm read access to tickets, comments, and attachments
- [ ] Generate an Atlassian API token with permission to create issues, comments, attachments, and customers
- [ ] Capture `ZENDESK_SUBDOMAIN_URL`, `ATLASSIAN_BASE_URL`, `JSM_PROJECT_KEY`, and `JSM_SERVICE_DESK_ID` in `.env`
- [ ] Store credentials in a secrets manager / vault (never commit `.env`)
- [ ] Provision the machine that runs the migration (Python 3, network access to both clouds)

### New task: Attachment migration validation

- [ ] Confirm attachment file-size and type limits on the JSM project
- [ ] Pilot a ticket with attachments and verify binaries re-upload correctly to JSM
- [ ] Decide handling for inline images embedded in descriptions/comments
- [ ] Spot-check that large attachments are not silently truncated or rejected

### New task: Pilot run & data validation / QA

- [ ] Run a controlled pilot (e.g. 25-50 representative tickets) end to end
- [ ] Verify summary, description (ADF), priority, status, labels, and reporter on migrated issues
- [ ] Confirm comments appear in chronological order with no description duplication
- [ ] Validate the `migration_state.db` resume behaviour (stop mid-run, re-run, confirm no duplicates)
- [ ] Reconcile counts: Zendesk tickets exported vs. JSM issues created
- [ ] Sign-off from Customer Experience team on pilot results

### New task: Production cutover plan

- [ ] Schedule the full ~10k run and a maintenance / freeze window
- [ ] Decide whether Zendesk is read-only during cutover to avoid drift
- [ ] Coordinate timing with the two parallel-use customers (Journal Inc, Santa Fe)
- [ ] Define success criteria and a go/no-go checkpoint
- [ ] Communicate cutover schedule to stakeholders and agents

### New task: Rollback & error recovery

- [ ] Document how to re-run safely (idempotent skip via `migration_state.db`)
- [ ] Define how to handle failed tickets logged in `migration.log` (triage & re-run)
- [ ] Decide rollback approach if a run must be abandoned (bulk-delete migrated issues by recorded Jira keys)
- [ ] Back up `migration_state.db` and `migration.log` after each major run

### New task: Post-migration cleanup & verification

- [ ] Final reconciliation of total ticket counts and a sample audit
- [ ] Verify requesters/customers were created and linked correctly in JSM
- [ ] Confirm queues route migrated tickets correctly (links to **Setup Queues**)
- [ ] Spot-check reports/dashboards reflect migrated data (links to **Configure reports/dashboards**)
- [ ] Archive migration artifacts (state DB, logs, mappings.xlsx) for audit
- [ ] Formally decommission / archive the Zendesk source per the agreed retention policy
