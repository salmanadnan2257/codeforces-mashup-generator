# Codeforces Mashup Generator

A single-script tool that builds fair mashup contests for a study group. It queries the public Codeforces API, filters the problemset by rating band, tags, and solver history, then writes the candidate problems to an Excel sheet.

## Why

When our study group ran mashup contests, someone always got a problem they'd already solved, which ruins the point. This script fixes that: it only picks problems that none of the listed group members (`--exclude-users`) have an accepted submission on. Optionally it also requires that at least one trusted strong user (`--include-users`, defaults to a list of well-known competitive programmers) has solved the problem, as a rough quality signal that the problem is worth setting.

This was solo work by Salman Adnan.

## Features

- Filter by any set of rating bands (default 800 to 2000 in steps of 100).
- Include filter: keep only problems solved by at least one handle in a trusted list.
- Exclude filter: drop any problem solved by any group member.
- Tag filtering with four modes: `some` (at least one listed tag), `all` (every listed tag present), `strict` (exact tag set), `off`.
- Configurable number of options per rating, output filename, request delay, and statement language.
- Output is a flat Excel sheet with rating, contest ID, index, name, tags, and a direct problem link.

## How it works

Three API endpoints do all the work:

1. `problemset.problems` fetches the full problemset once, bucketed by rating.
2. `user.status` is called per handle (paginated 10,000 submissions at a time) to build two sets of `(contestId, index)` pairs: solved-by-includes and solved-by-excludes.
3. Each bucketed problem passes if it is in the include set, not in the exclude set, and matches the tag rules. The first N survivors per rating go to the sheet via pandas + openpyxl.

API errors (bad handle, rate limit, network failure) are caught per request: the script prints the error, skips that handle or aborts that fetch, and continues with whatever data it has.

## Setup

Python 3.10+ recommended (tested on 3.12).

```
python -m venv .venv
.venv/bin/pip install -r requirements.txt
```

No API key or environment variables are needed. The Codeforces API is public but rate limited, so the script sleeps between calls (`--delay`, default 0.5s).

## Usage

Defaults reproduce the original study-group configuration (46 trusted handles, 34 group members, 10 choices per rating from 800 to 2000). Note the default include/exclude lists take a while: the script downloads the full submission history of every listed handle.

A small run, verified against the live API (took about 14 seconds):

```
python generate_mashup.py \
    --include-users ksun48 \
    --exclude-users salmanadnan2025 Musab1Blaser \
    --ratings 800 1200 1600 \
    --choices-per-rating 3 \
    --include-tags greedy --tag-mode some \
    -o mashup_test.xlsx
```

This produced 9 rows (3 per rating), all tagged `greedy`, none solved by the two excluded members, all solved by ksun48.

Other examples:

```
# Skip the trusted-solver requirement, just avoid the group's solved problems
python generate_mashup.py --no-include-filter --ratings 1000 1100 1200

# Pure bitmask practice set
python generate_mashup.py --include-tags bitmasks --tag-mode all

# See every flag
python generate_mashup.py --help
```

One default changed during the CLI refactor: the old hardcoded config had `include_tags = ["bitmasks"]` baked in from the last session it was used for. The CLI defaults to no tag filter, since that's the sensible general case; pass `--include-tags bitmasks` to get the old behavior.

## Challenges

- The Codeforces API splits the data I needed across two differently shaped endpoints. `problemset.problems` returns a flat list of problem objects (each with `contestId`, `index`, `rating`, `tags`), while `user.status` returns submission objects that wrap the problem under a `problem` key. To compare them I had to reduce both to the same `(contestId, index)` identity. Building the solved set from submissions and then testing each problemset entry against that set is what makes the include/exclude filters line up.
- Some problems in a submission record have no top-level `contestId`, or the field lives on the submission rather than the nested problem. I resolved it in `collect_solved` with `prob.get('contestId') or s.get('contestId')`, and I only add a key when both the contest id and index are present. Without that guard a handful of gym or malformed entries would have polluted the solved set with partial keys.
- The include and exclude logic both need the exact same "which problems has this group of handles solved" computation, just over different handle lists. Early on that was one copied loop pasted twice, which meant a fix to the verdict check had to be made in two places. I pulled it into a single `collect_solved(handles, delay)` helper and call it once for includes and once for excludes, so the dedup and the `verdict == 'OK'` rule live in exactly one spot.
- The script started life as a hardcoded config block (handle lists, ratings, and an `include_tags = ["bitmasks"]` value left over from the last session it was run for). Turning that into `argparse` meant deciding what stays a default and what becomes a flag. I kept the original handle lists as `DEFAULT_INCLUDE_USERS` / `DEFAULT_EXCLUDE_USERS` so a bare run still reproduces the study-group setup, but dropped the stale `bitmasks` tag default because that was specific to one session, not a sensible general default.
- User submission histories are large and the API paginates, so a single request can't fetch everything. `fetch_user_status` loops with `from` and `count=10000`, stops when a page comes back shorter than the batch size, and sleeps `--delay` between pages to stay under the rate limit. Bad handles return an HTTP 400 or an API-level non-OK status. Both are caught: the network path via `raise_for_status` inside a try/except, and the application error via checking `data.get('status')`. A failed handle prints a message and is skipped rather than crashing the whole run, which I confirmed live when two stale handles in the default exclude list returned 400 and the sheet still generated correctly.

## What I learned

- A REST API can hand you the same logical entity in two different envelope shapes. Picking one canonical identity, here the `(contestId, index)` tuple, and normalizing everything to it before doing set math is cleaner than trying to compare the raw objects.
- `raise_for_status()` only covers transport-level failures. Codeforces returns HTTP 200 with a JSON body of `{"status": "FAILED", "comment": ...}` for things like an unknown handle, so you need a second check on the payload's own status field. Handling only one of the two leaves a real class of errors unhandled.
- Cursor-style pagination with `from`/`count` needs an explicit stop condition. Treating "a page shorter than the requested batch size" as the end is simple and avoids an off-by-one fetch of an empty final page.
- Set membership is the right shape for this problem. Building `include_solved` and `exclude_solved` once as sets turns every per-problem check into an O(1) lookup instead of rescanning submission lists per problem.
- Passing timeouts to every `requests.get` matters for a script that fans out over dozens of handles. A single hung connection with no timeout can stall the entire run, so the problemset fetch uses 60 seconds and each status page uses 30.

## What I'd do differently

- Downloading every submission of every trusted handle is the wrong tool. Handles like tourist have tens of thousands of submissions, so a default run pulls hundreds of megabytes to answer a set-membership question. Caching solved-sets to disk between runs, or dropping the include filter in favor of `solvedCount` from `problemset.problems`, would cut runtime by orders of magnitude.
- Selection is `plist[:n]`, so it always takes the newest problems in API order rather than sampling randomly. Repeat runs with the same filters produce near-identical sheets; a `random.sample` with an optional seed would be better.
- A failed `user.status` call for an excluded member silently weakens the guarantee: their solved problems are treated as unsolved. That case should be a hard error, not a warning, because fairness is the whole point.
- Errors go to stdout via `print` and the exit code is always 0. Proper logging and a nonzero exit when nothing was saved would make it scriptable.
- The choices-per-rating value is now a single global number; the old dict allowed per-rating counts, which nobody used, but a `RATING=N` syntax would restore that flexibility.
