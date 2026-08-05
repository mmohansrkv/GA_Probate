# LNGA GA Probate Courts — Web App

A web version of the desktop `LNGA_ProbateCourts_Combined_01.py` tool
(Citation Search + Validation), deployable to [Render](https://render.com)
or any Docker host. Same scraping/parsing/Excel logic as the original —
only the Tkinter GUI was replaced with a Flask web UI + JSON API.

## Files

| File | Purpose |
|---|---|
| `core.py` | All non-GUI logic ported unchanged from the desktop tool: input parsing, the Selenium scrape of the live site, offline HTML/Excel validation, and result-workbook writing. |
| `app.py` | Flask app: a one-page web UI and a JSON API, both backed by `core.py`. Runs jobs in a background thread and exposes status/log polling. |
| `Dockerfile` | Python + headless Chromium/Chromedriver image (Chromium is required for the live-scrape feature). |
| `render.yaml` | Render.com Blueprint — one Docker web service, health check on `/healthz`. |
| `requirements.txt` | Python dependencies. |

## What each tool does

- **Citation Search** (live scrape) — for each `County / Starting Date / Ending
  Date` row in your uploaded input file, drives headless Chrome against
  `https://georgiaprobaterecords.com/Traffic/SearchCitations.aspx`, scrapes
  every result row, and writes a combined `result_*.xlsx`. If you also
  upload an existing case-number workbook, it's checked against the fresh
  scrape and a `CaseNumber_MatchStatus_*.xlsx` (Match / New / Missing) is
  produced too.
- **Validation** — fully offline. Cross-checks a folder of `debug_<county>.html`
  dumps against an Output Excel and an Input Excel, and writes a match report.

## Running locally

```bash
pip install -r requirements.txt
python app.py
# -> http://localhost:5000
```

Selenium needs a real Chrome/Chromium + matching driver on your machine for
Citation Search to work locally; Validation has no such requirement.

## Running with Docker

```bash
docker build -t lnga-probate .
docker run -p 5000:5000 lnga-probate
```

The image bakes in Chromium + chromedriver, so Citation Search works out of
the box in the container.

## Deploying to Render

1. Push this repo to GitHub.
2. In Render: **New → Blueprint**, point it at the repo — `render.yaml` is
   picked up automatically and provisions one Docker web service.
3. Once deployed, open the service URL for the web UI, or use the API below.

> **Note:** `georgiaprobaterecords.com` must be reachable from wherever this
> is deployed for Citation Search to work — check Render's outbound network
> isn't blocked for your plan if scrapes fail immediately.

## Web UI

Open `/` — two tabs:

- **Citation Search**: upload your input file (`.txt` tab-separated or
  `.xlsx`, same `County / Starting Date / Ending Date` columns as always),
  optionally an existing case-number file to validate against, set an
  optional county limit / inter-county delay, and run. Watch the live log,
  then download the result workbook(s).
- **Validation**: upload three `.zip` files (or single `.xlsx`/`.html` for
  single-file cases) — your Offline HTML dumps, Output Excel, and Input
  Excel — and run. Download the validation result workbook when done.

## API

All endpoints return JSON. Long-running work is async: POST returns a
`job_id` immediately; poll `/api/jobs/<job_id>` for progress.

### `POST /api/citation-search`

`multipart/form-data`:

| Field | Required | Notes |
|---|---|---|
| `input_file` | yes | `.txt` or `.xlsx`, columns `County / Starting Date / Ending Date` |
| `compare_file` | no | existing case-number `.xlsx` to validate the scrape against |
| `headless` | no | `"true"` (default) or `"false"` |
| `limit` | no | only process the first N counties (useful for testing) |
| `delay_min`, `delay_max` | no | random delay range (seconds) between counties |

Response: `202 { "job_id": "..." }`

### `POST /api/validate`

`multipart/form-data`, each field a `.zip` of a folder (or a single file):

| Field | Contents |
|---|---|
| `offline` | `debug_<county>.html` files |
| `output` | Output Excel workbook(s) |
| `input` | Input Excel workbook(s) |

Response: `202 { "job_id": "..." }`

### `GET /api/jobs/<job_id>`

```json
{
  "id": "954c7e42789d",
  "kind": "citation_search",
  "status": "running | done | error | pending",
  "log": ["...", "..."],
  "stats": { "...": "pipeline-specific summary numbers" },
  "error": null,
  "downloads": ["result_workbook", "match_status_workbook"]
}
```

### `GET /api/jobs/<job_id>/download/<label>`

Streams the named result file (`label` comes from the `downloads` list
above).

### `GET /healthz`

Liveness check; also reports whether Selenium/Chrome could be detected.

## Example: run your weekly county list against the API with `curl`

```bash
curl -F "input_file=@LNGA_ProbateCourts_Input.txt" \
     https://<your-render-url>/api/citation-search
# -> {"job_id": "..."}

curl https://<your-render-url>/api/jobs/<job_id>

curl -OJ https://<your-render-url>/api/jobs/<job_id>/download/result_workbook
```

## Known limitations vs. the desktop tool

- Jobs are tracked in an in-memory dict — fine for a single container
  instance; if you scale to multiple instances behind a load balancer,
  swap `JOBS` in `app.py` for Redis or a DB-backed table.
- Uploaded files and job results live under `/tmp/lnga_jobs` (ephemeral
  container storage) — download results promptly after a run, or mount a
  persistent disk and point `LNGA_DATA_ROOT` at it.
- The site's RadGrid/Telerik interactions in `core.py` are exactly as
  fragile as they were in the desktop tool — if the site changes its
  markup, the scrape will need the same kind of selector fixes as before.
