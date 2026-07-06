#!/usr/bin/env python3
"""
generate_mashup.py

Fetch multiple Codeforces problems per rating, ensuring:
    1. Optionally: Each problem is solved by at least one of --include-users.
    2. Optionally: No problem is solved by any of --exclude-users.
    3. Optionally: Problems include/exclude specific tags according to --tag-mode.

Configure filters and how many choices you want per rating via CLI arguments
(run with --help for the full list). Defaults match the original study-group setup.
Outputs a single Excel file (default `filtered_problems.xlsx`) with all selected options.
"""
import argparse
import time

import requests
import pandas as pd

# === DEFAULTS ===
DEFAULT_INCLUDE_USERS = [  # only used unless --no-include-filter
    "-is-this-fft-", "A_G", "Acanikolic73", "AlphaMale06", "Beng",
    "Dominator069", "Error_Yuan", "FEDIKUS", "HS90R", "LMeyling",
    "Little_Sheep_Yawn", "N_z_", "Nachia", "NemanjaSo2005", "Ormlis",
    "Pekiban", "Phantom_Performer", "Prady", "Qingyu", "RTE",
    "Radewoosh", "Shayan", "TimDee", "_.Ali._", "istil",
    "abc864197532", "adamant", "bssss_a", "culver0412", "ecnerwala",
    "errorgorn", "hashman", "jiangly", "jqdai0815", "kotatsugame",
    "ksun48", "larush", "maomao90", "maroonrk", "maspy",
    "milisav", "miumah", "oreg0na1", "orzdevinwang", "prvocislo",
    "tourist"
]
DEFAULT_EXCLUDE_USERS = [  # only used unless --no-exclude-filter
    "Alpha_Anas_70", "Ansari_Hamza", "Blingblong", "Genesis_X",
    "Moiz08229", "Musab1Blaser", "SadCivic", "Yasha_zaidi",
    "ZainNaqi", "hassan343", "lanaya-ramis", "lucario_knight",
    "meesum", "msaad01", "rk10_-3", "sa07885", "it.expert2210",
    "shayaanqazi", "syanide", "wsaleem", "Panta_Iilisha", "Kakusei",
    "Pramag_IIITD", "Satyanshu111", "Yapper", "imzero34802", "lsjo",
    "peter_griffin_codes", "salmanadnan2025", "Aritro_", "Drakozs",
    "chpp", "Vyaduct", "chiru200513"
]
DEFAULT_RATINGS = [800, 900, 1000, 1100, 1200, 1300, 1400, 1500, 1600, 1700, 1800, 1900, 2000]

# === CLI ===
def parse_args():
    parser = argparse.ArgumentParser(
        description="Generate a Codeforces mashup sheet: pick problems per rating "
                    "that trusted users solved but your group members have not.",
        formatter_class=argparse.ArgumentDefaultsHelpFormatter,
    )
    parser.add_argument("--include-users", nargs="*", default=DEFAULT_INCLUDE_USERS,
                        metavar="HANDLE",
                        help="problems must be solved by at least one of these handles")
    parser.add_argument("--exclude-users", nargs="*", default=DEFAULT_EXCLUDE_USERS,
                        metavar="HANDLE",
                        help="problems must NOT be solved by any of these handles")
    parser.add_argument("--no-include-filter", action="store_true",
                        help="skip the include-users filter entirely")
    parser.add_argument("--no-exclude-filter", action="store_true",
                        help="skip the exclude-users filter entirely")
    parser.add_argument("--ratings", nargs="+", type=int, default=DEFAULT_RATINGS,
                        metavar="R",
                        help="rating bands to fetch problems for")
    parser.add_argument("--choices-per-rating", type=int, default=10, metavar="N",
                        help="how many problem options to pick per rating")
    parser.add_argument("--include-tags", nargs="*", default=[], metavar="TAG",
                        help="tags to filter on (empty = no include filter)")
    parser.add_argument("--exclude-tags", nargs="*", default=[], metavar="TAG",
                        help="problems must NOT have any of these tags")
    parser.add_argument("--tag-mode", choices=["off", "some", "all", "strict"],
                        default="all",
                        help="how to match --include-tags: 'some' = at least one, "
                             "'all' = every listed tag present, 'strict' = exact tag set, "
                             "'off' = ignore include tags")
    parser.add_argument("--lang", default="en",
                        help="language of fetched problem statements (API lang param)")
    parser.add_argument("--delay", type=float, default=0.5, metavar="SECONDS",
                        help="delay between API calls")
    parser.add_argument("-o", "--output", default="filtered_problems.xlsx",
                        help="output Excel filename")
    return parser.parse_args()

# === CF API HELPERS ===
def fetch_user_status(handle, delay):
    """Fetch *all* submissions for a given user, via pagination."""
    all_subs = []
    batch_size = 10000
    start = 1
    while True:
        url = (
            f"https://codeforces.com/api/user.status"
            f"?handle={handle}&from={start}&count={batch_size}"
        )
        try:
            r = requests.get(url, timeout=30)
            r.raise_for_status()
            data = r.json()
            if data.get('status') != 'OK':
                print(f"API error for {handle}: {data.get('comment')}")
                break
            subs = data.get('result', [])
            if not subs:
                break
            all_subs.extend(subs)
            time.sleep(delay)
            if len(subs) < batch_size:
                break
            start += batch_size
        except Exception as e:
            print(f"Error fetching submissions for {handle}: {e}")
            break
    return all_subs

def fetch_problemset(lang, delay):
    """Fetch full problemset from CF."""
    url = f"https://codeforces.com/api/problemset.problems?lang={lang}"
    try:
        r = requests.get(url, timeout=60)
        r.raise_for_status()
        data = r.json()
        if data.get('status') == 'OK':
            time.sleep(delay)
            return data['result'].get('problems', [])
        print(f"API error fetching problemset: {data.get('comment')}")
    except Exception as e:
        print(f"Error fetching problemset: {e}")
    return []

def collect_solved(handles, delay):
    """Set of (contestId, index) pairs with an OK verdict across the given handles."""
    solved = set()
    for u in handles:
        for s in fetch_user_status(u, delay):
            if s.get('verdict') == 'OK':
                prob = s.get('problem', {})
                cid = prob.get('contestId') or s.get('contestId')
                idx = prob.get('index')
                if cid and idx:
                    solved.add((cid, idx))
    return solved

# === FILTER LOGIC ===
def get_filtered_problems(args):
    problems = fetch_problemset(args.lang, args.delay)
    by_rating = {r: [] for r in args.ratings}
    for p in problems:
        r = p.get('rating')
        if r in by_rating:
            by_rating[r].append(p)

    include_solved = set()
    if not args.no_include_filter:
        include_solved = collect_solved(args.include_users, args.delay)

    exclude_solved = set()
    if not args.no_exclude_filter:
        exclude_solved = collect_solved(args.exclude_users, args.delay)

    def passes_tag_filter(problem):
        # always apply exclude_tags
        tags = set(problem.get('tags', []))
        if args.exclude_tags and tags & set(args.exclude_tags):
            return False

        # include_tags enforcement based on tag mode
        if args.include_tags and args.tag_mode != "off":
            inc_set = set(args.include_tags)

            if args.tag_mode == "some":
                if not (tags & inc_set):
                    return False
            elif args.tag_mode == "all":
                if not inc_set.issubset(tags):
                    return False
            elif args.tag_mode == "strict":
                if tags != inc_set:
                    return False

        return True

    filtered = {}
    for r, plist in by_rating.items():
        valid = []
        for p in plist:
            key = (p['contestId'], p['index'])
            if (
                (args.no_include_filter or key in include_solved)
                and (args.no_exclude_filter or key not in exclude_solved)
                and passes_tag_filter(p)
                ):
                valid.append(p)
        filtered[r] = valid
    return filtered

# === CHOICES & OUTPUT ===
def pick_choices_per_rating(filtered, n):
    selection = {}
    for r, plist in filtered.items():
        selection[r] = plist[:n]
    return selection

def save_to_excel(choices, filename):
    rows = []
    for r, progs in choices.items():
        for p in progs:
            rows.append({
                'rating': r,
                'contestId': p['contestId'],
                'index': p['index'],
                'name': p['name'],
                'tags': ", ".join(p.get('tags', [])),
                'link': f"https://codeforces.com/problemset/problem/{p['contestId']}/{p['index']}"
            })
    if not rows:
        print("No problems to save.")
        return
    df = pd.DataFrame(rows)
    df.to_excel(filename, index=False)
    print(f"Results saved to {filename}")

# === MAIN ===
if __name__ == '__main__':
    args = parse_args()
    filtered = get_filtered_problems(args)
    choices = pick_choices_per_rating(filtered, args.choices_per_rating)
    save_to_excel(choices, args.output)
    print("Done.")
