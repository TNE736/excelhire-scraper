# Excel Hire candidate search

Pastes a job description into Excel Hire's **Job Seekers → AI Search → External** search, runs it,
and saves the top 20 candidates to an Excel file.

- **No credits used.** It never clicks **Reveal** (the only thing that spends credits) and checks the
  credit counter before and after every run.
- **No AI of its own.** It reads the results table exactly as Excel Hire shows it.
- Email and phone are not collected (they sit behind Reveal).

## What you need

- Windows with **Python 3.10+** ([python.org](https://www.python.org/downloads/)) and **Git**
- **Your own** Excel Hire account with access to AI Search → External

## Setup (once)

In PowerShell:

```powershell
git clone https://github.com/TN729/excelhire-scraper.git
cd excelhire-scraper
pip install -r requirements.txt
python -m playwright install chromium
copy .env.example .env
notepad .env
```

In `.env`, fill in **your** Excel Hire email and password, then save. `.env` never leaves your PC
(git ignores it). Never put real values in `.env.example`.

## Run a search

1. Put the job description in a text file:
   ```powershell
   notepad jd.txt
   ```
   Paste the JD, save, close. (Any name works: `python-dev.txt`, `maximo.txt`, …)
2. Run:
   ```powershell
   python search.py jd.txt
   ```

The browser runs **hidden**. On the first run it logs in by itself (email → Login → password →
Login) and remembers the login in `browser-profile/`, so later runs go straight to the search. If
Excel Hire ever asks for something else, such as a code, the run stops with `Login did not
complete`: run `python search.py jd.txt --show` once and finish the login in the window.

### How many candidates

It saves **20** by default. To change that:

- **For one run:** add `--count`, for example `python search.py jd.txt --count 10`.
- **For every run:** open `search.py` and change the `20` in this line (near the end of the file):
  ```python
  parser.add_argument("--count", type=int, default=20, help="profiles to save (max 25)")
  ```

The most it can save is **25**: that is how many Excel Hire shows on the first results page, and
the scraper reads only that page. Ask for more and you still get 25.

### Options

| Option | Meaning |
|---|---|
| `--mode "Best Fit"` | Default. Uses job titles and skills only: a wider pool |
| `--mode Exact` | Uses every filter Excel Hire parsed from the JD (experience, location, …) |
| `--mode "Skill Match"` | Ranks on skills |
| `--count 10` | How many profiles to save: 1 to 25, default 20 (see above) |
| `--show` | Show the browser window (it runs hidden by default). Use it if the login needs a code, or to watch a run |

## The result

An Excel file in `output/`, named after the JD and the time, with two sheets:

- **Candidates**: #, Name, Score (green = strongest), Job Title, Company, Location, Experience
  (one role per line), Skills (full list), Education, LinkedIn link, and Excel Hire's "why matched"
  note. A blank Education means Excel Hire has none for that person.
- **Search**: when it ran, the mode, Excel Hire's total matches and the JD used.

The file holds people's personal data: keep it on your PC, share it only with people who need it,
and delete it when the role is closed. `output/` is git-ignored.

## If something goes wrong

| Message | Fix |
|---|---|
| `playwright` is not recognized | Use `python -m playwright install chromium` |
| `Login did not complete` | Check the email and password in `.env`, or run once with `--show` and log in in the window |
| `The results table has changed` | Excel Hire changed its page; the scraper stops rather than save wrong columns. Ask the maintainer to update `search.py` |
| `WARNING: credits changed` | Something spent credits during the run. Stop using the tool and tell the maintainer |

## Files

| File | In git | What it is |
|---|---|---|
| `search.py` | yes | The scraper |
| `requirements.txt` | yes | Python packages |
| `.env.example` | yes | Template for `.env`, with empty values |
| `.env` | **no** | Your Excel Hire login |
| `browser-profile/` | **no** | Your saved Excel Hire session |
| `output/` | **no** | The Excel results |
| `*.txt` (JDs) | **no** | Your job descriptions |
