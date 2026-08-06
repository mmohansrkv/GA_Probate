# LNGA Probate Courts Tool — web edition

This is a Render.com-deployable web version of the original
`LNGA_ProbateCourts_Combined_Tool.py` Tkinter desktop app. It exposes the
same two modules as a small Flask site instead of a desktop GUI:

- **Citation Search** — live-scrapes
  `https://georgiaprobaterecords.com/Traffic/SearchCitations.aspx` for every
  County / Starting Date / Ending Date row in an uploaded file (Selenium +
  headless Chrome), and produces an **Output** workbook (every case number
  found) plus a **Result** workbook (Match / New / Missing vs. an optional
  existing case-number list).
- **Validation** — compares uploaded offline HTML case files against an
  uploaded Output Excel, field by field, and produces a **Result** workbook
  with `Validation_Result`, `Case#_Diff_County_Check`, and
  `Input_vs_Output` sheets.

## Files

| File | Purpose |
|---|---|
| `app.py` | Flask app: upload forms, background job runner, status polling, downloads. |
| `core.py` | All the original business logic (HTML parsing, Excel building, Selenium scraping, validation) with the Tkinter GUI removed. |
| `Dockerfile` | Installs Google Chrome (needed by Selenium) on top of Python 3.11. |
| `render.yaml` | Render Blueprint — one Docker web service + a persistent disk. |
| `requirements.txt` | Python dependencies. |
| `LNGA_ProbateCourts_Input.txt` | Sample County / Starting Date / Ending Date input file (tab-separated), for the Citation Search form. |

## Why a `Dockerfile`?

Selenium needs an actual Chrome browser on the machine it runs on. Render's
native Python runtime has no `apt-get`/root access to install one, so this
service builds from `render.yaml`'s `env: docker`, using the included
`Dockerfile` to install `google-chrome-stable`. Selenium 4.6+'s built-in
"Selenium Manager" downloads a matching `chromedriver` automatically at
runtime — you don't need to install/pin a driver version yourself.

## Deploying to Render

1. Push this folder to a GitHub/GitLab repo.
2. In the Render dashboard: **New → Blueprint**, point it at the repo.
   Render reads `render.yaml` and creates:
   - a Docker web service (`lnga-probate-courts-tool`)
   - a 5 GB persistent disk mounted at `/var/data` (holds uploads + generated
     workbooks so they survive restarts/redeploys)
3. Deploy. First build takes a few minutes (installing Chrome).
4. Open the service URL — you'll see the home page with the two module
   cards.

**Plan note:** Chrome is memory-hungry; `render.yaml` requests the
`starter` plan. The free plan's 512 MB will likely OOM mid-scrape.

## Environment variables

| Var | Set by | Purpose |
|---|---|---|
| `DATA_DIR` | `render.yaml` | Where job uploads/results are written (defaults to `./data` locally). |
| `CHROME_BIN` | `Dockerfile` | Path to the Chrome binary `core.py`'s `build_driver()` should use. |
| `CHROMEDRIVER_PATH` | (unset) | Only needed if you want to pin a specific chromedriver instead of letting Selenium Manager fetch one. |
| `PORT` | Render (automatic) | Port gunicorn binds to. |

## Running locally

```bash
pip install -r requirements.txt
# Chrome/chromedriver must be installed locally for Citation Search to work;
# Validation works without them.
python app.py
# -> http://localhost:5000
```

## Using the site

### Citation Search (`/citation-search`)
1. Upload a County / Starting Date / Ending Date file — either the
   tab-separated `.txt` format (see `LNGA_ProbateCourts_Input.txt`) or an
   `.xlsx` with the same three columns (dates as `DD-MM-YYYY`).
2. Optionally upload an existing case-number workbook (`.xlsx` with a
   "Case #" column) to diff this run against; if you skip it, the run's own
   results are saved as a fresh baseline for next time.
3. Optionally set a county **limit** (for a quick test) and a random delay
   range between counties (be polite to the source site).
4. Submit — you land on a live status page that polls progress every 1.5s
   and shows download buttons for the **Output** and **Result** workbooks
   once the run finishes.

### Validation (`/validation`)
1. Upload all the offline `.html`/`.htm` case files you want to validate.
2. Upload one or more Output Excel (`.xlsx`) files (each row = one citation).
3. Optionally upload Input Excel/txt file(s) to also run the
   Input-vs-Output check.
4. Submit — same live status page, with a download button for the Result
   workbook when done.

## Architecture notes / limitations

- **Single worker, in-memory job state.** `app.py` keeps job status/logs in
  a Python dict guarded by a lock, not a database. The Dockerfile's
  `gunicorn` command therefore runs `--workers 1 --threads 8` — multiple
  threads keep the UI responsive during a long scrape, but you must **not**
  scale this service to more than one instance/worker without first moving
  job state to something shared (Redis, Postgres, etc.), or different
  requests will land on processes that don't know about each other's jobs.
- **Long-running jobs run in a background thread** kicked off by the upload
  route; the HTTP request returns immediately with a status page, so Render
  request timeouts aren't an issue. `Dockerfile`'s gunicorn `--timeout 600`
  is just a safety margin for the initial POST/file-upload itself.
- **No authentication.** Add one (e.g. Render's built-in basic auth add-on,
  or a login check in `app.py`) before exposing this publicly if the data
  is sensitive.
- **Site fragility.** The Citation Search scraper is written against the
  current markup/control IDs of `georgiaprobaterecords.com`'s ASP.NET/
  Telerik grid. If that site changes its markup, `core.py`'s selectors
  (`select_county`, `scrape_result_rows`, etc.) will need updating — this
  was already true of the original desktop tool.
- **Disk cleanup.** Every job's uploads/results live under
  `$DATA_DIR/jobs/<job_id>/`. There's a `POST /jobs/<id>/cleanup` endpoint
  to delete a finished job's files; nothing calls it automatically, so
  point a cron/scheduled job at it (or prune `/var/data/jobs` periodically)
  if disk usage matters.
