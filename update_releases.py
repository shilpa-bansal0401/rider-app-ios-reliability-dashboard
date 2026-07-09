#!/usr/bin/env python3
"""
Adds the next weekly release to PINNED_RELEASES in both report scripts,
removing the oldest entry. Run every Monday before generating reports.

Version rule:  increment the week number (middle component), patch = 1
Dist rule:     last dist + 3
"""

import re
import sys

FILES = ["generate_hang_report.py", "generate_crash_report.py"]

BLOCK_RE      = re.compile(r'(PINNED_RELEASES\s*=\s*\[)(.*?)(\])', re.DOTALL)
ENTRY_RE      = re.compile(r'\{"version":\s*"([\d.]+)",\s*"dist":\s*"(\d+)"\}')
DISTS_ENTRY_RE = re.compile(r'\{"version":\s*"([\d.]+)",\s*"dists":\s*\[([^\]]+)\]\}')

def parse_releases(block):
    """Return list of (version, last_dist) tuples from a PINNED_RELEASES block."""
    entries = []
    for m in re.finditer(r'\{[^}]+\}', block):
        chunk = m.group(0)
        single = ENTRY_RE.match(chunk)
        if single:
            entries.append((single.group(1), single.group(2)))
            continue
        multi = DISTS_ENTRY_RE.match(chunk)
        if multi:
            ver = multi.group(1)
            dists = [d.strip().strip('"') for d in multi.group(2).split(",")]
            entries.append((ver, max(dists, key=int)))
    return entries

for path in FILES:
    with open(path) as f:
        text = f.read()

    m = BLOCK_RE.search(text)
    if not m:
        sys.exit(f"ERROR: PINNED_RELEASES not found in {path}")

    releases = parse_releases(m.group(2))
    if not releases:
        sys.exit(f"ERROR: no entries in PINNED_RELEASES in {path}")

    last_ver, last_dist = releases[-1]
    parts = last_ver.split(".")
    if len(parts) < 3:
        parts.append("0")
    parts[1] = str(int(parts[1]) + 1)
    parts[2] = "1"
    new_ver  = ".".join(parts)
    new_dist = str(int(last_dist) + 3)

    updated = releases[1:] + [(new_ver, new_dist)]
    new_body = "\n" + "".join(
        f'    {{"version": "{v}", "dist": "{d}"}},\n'
        for v, d in updated
    )

    new_text = text[:m.start()] + m.group(1) + new_body + m.group(3) + text[m.end():]

    with open(path, "w") as f:
        f.write(new_text)

    print(f"{path}: removed {releases[0][0]}, added {new_ver} (dist {new_dist})")
