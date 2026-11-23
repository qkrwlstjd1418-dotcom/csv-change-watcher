# change_watch — what changed since yesterday's export?

One Python file. It compares two snapshots of the same export (yesterday and
today) and tells you the three things that get missed:

- a row that **appeared** (a new product, order, account, branch)
- a row that **vanished** (usually the interesting one)
- a value that **moved quietly** — with before, after and the difference

Then it appends every run to `change_log.csv`, so you can see *when* a value
moved and not just that it did.

```
$ python change_watch.py
  2 new · 1 removed · 4 changed field(s)
  report: output/change_report_20261124_0635.xlsx
     sku  field before    after  difference
SKU-1001  price   19.0     21.0         2.0
SKU-1002  stock  610.0      0.0      -610.0
SKU-1002 status active sold out         NaN
SKU-1003  stock   88.0     71.0       -17.0
```

That is the real output of the bundled sample data — two small snapshots of a
product export, where a price moved, a stock ran out, one product dropped out
and two appeared.

## Run it

```
git clone https://github.com/qkrwlstjd1418-dotcom/csv-change-watcher.git
cd csv-change-watcher
python -m pip install -r requirements.txt
python change_watch.py
```

Python 3.10+ and two packages (`pandas`, `openpyxl`). Windows, macOS, Linux.
Nothing is written outside `output/` unless you pass `--output`.

Point it at your own files:

```
python change_watch.py --previous "C:/work/export_mon.csv" --current "C:/work/export_tue.csv" \
    --key sku --compare price,stock
```

`--help` lists every flag.

## What it writes

One workbook, four sheets:

| Sheet | Contents |
|---|---|
| `summary` | row counts, new / removed / changed counts, duplicate-key count, run time |
| `changed` | one line per field that moved: key, field, before, after, difference |
| `new` | rows that exist only in today's file |
| `removed` | rows that exist only in yesterday's file |

Plus `output/change_log.csv`, appended on every run.

## Decisions worth knowing about

**The key is your choice.** Default is the first column; `--key sku` or
`--key branch,sku` for a composite key. Duplicates are counted and reported in
`summary` before anything is dropped — silently de-duplicating on every column
is how a real order disappears from a total.

**Numbers are compared as numbers.** `1,200`, `₩1,200`, `$1,200` and `(300)` all
parse; a difference under `--tolerance` (default `0.005`) counts as unchanged.
Text is compared trimmed.

**Exit code 2 means something moved.** Schedule it and let the exit code decide
whether a person needs to look:

```
# cron, every weekday at 07:00; mail only when something changed
0 7 * * 1-5 cd /work && python change_watch.py || echo "changed" | mail -s "export watch" you@example.com
```

Windows Task Scheduler: action `python C:\work\change_watch.py`, trigger daily
07:00, and notify on a non-zero exit code.

**Input files:** `.csv` (tries utf-8-sig, utf-8, cp1252, latin-1) and `.xlsx`.
For fixed-width exports or a Google Sheets link, export to CSV first.

## Where it came from

This is one of six scripts I wrote for the report work that repeats every month:
merging a folder of files whose headers almost match, pulling tables out of
PDFs, an API into Excel, the weekly pack, two-file reconciliation — and this
one.

The other five, with the sample data and field notes for each, are in the
**Python Automation Pack**:
[siljeondatalab.gumroad.com/l/python-automation-pack](https://siljeondatalab.gumroad.com/l/python-automation-pack)

SQL side, if that is the part you need:
[Cohort, Funnel & RFM Kit](https://siljeondatalab.gumroad.com/l/cohort-funnel-rfm-kit) ·
[Oracle & MySQL Performance Analytics Field Manual](https://siljeondatalab.gumroad.com/l/sql-excel-field-manual)

## Licence

MIT — use it at work, in client projects, fork it, change it. No warranty: it is
a reporting helper, and you are responsible for the numbers you publish.
