import json, re, glob, os
from collections import Counter

files = sorted(glob.glob(r"D:\parmatmapanchvinshika\parmatmapanch\src\shlokas_*.json"))
banned = ["हे मुमुक्ष", "भो भव्य", "पश्यन्तु", "श्रूयताम्"]
all_ok = True

all_prashnas = []
all_prayogs = []
all_sutras = []
all_quotes = []

for fp in files:
    with open(fp, encoding="utf-8") as f:
        data = json.load(f)
    for item in data:
        n = item["num"]
        av_w = len([w for w in re.split(r"\s+", item["avataranika_sa"]) if w])
        vr_w = len([w for w in re.split(r"\s+", item["vritti_sa"]) if w and w != ">"])
        gk_lines = [l for l in item["gurjar_kavya"].splitlines() if l.strip()]
        bh_lines = [l for l in item["bhavanuvada"].splitlines() if l.strip()]
        issues = []
        if not (70 <= av_w <= 80):
            issues.append(f"Avataranika={av_w} (need 70-80)")
        if vr_w < 300:
            issues.append(f"Vritti={vr_w} (need >=300)")
        if len(gk_lines) != 4:
            issues.append(f"GurjarKavya lines={len(gk_lines)} (need 4)")
        if not (4 <= len(bh_lines) <= 5):
            issues.append(f"Bhavanuvada lines={len(bh_lines)} (need 4-5)")
        if len(item["sadhana_prayog"]) != 3:
            issues.append(f"Prayog={len(item['sadhana_prayog'])}")
        if len(item["atmaprashna"]) != 5:
            issues.append(f"Prashna={len(item['atmaprashna'])}")
        if len(item["sadhana_sutra"]) != 5:
            issues.append(f"Sutra={len(item['sadhana_sutra'])}")
        for b in banned:
            if b in item["avataranika_sa"] or b in item["vritti_sa"]:
                issues.append(f"Banned vocative: {b}")

        all_prashnas.extend([(n, p) for p in item["atmaprashna"]])
        all_prayogs.extend([(n, p) for p in item["sadhana_prayog"]])
        all_sutras.extend([(n, p) for p in item["sadhana_sutra"]])

        # Extract reference quotes in vritti_sa
        for line in item["vritti_sa"].splitlines():
            if line.strip().startswith(">"):
                all_quotes.append((n, line.strip()))

        if issues:
            all_ok = False
            print(f"[FAIL] Shloka {n}: " + ", ".join(issues))
        else:
            print(f"[OK]   Shloka {n}: Av={av_w}, Vr={vr_w}")

# Check exact duplicates in prashnas/prayogs/sutras
for label, lst in [("Atmaprashna", all_prashnas), ("SadhanaPrayog", all_prayogs), ("SadhanaSutra", all_sutras)]:
    c = Counter([x[1] for x in lst])
    dups = [k for k, v in c.items() if v > 1]
    if dups:
        all_ok = False
        print(f"[DUP] {label} has {len(dups)} exact duplicates!")
    else:
        print(f"[OK]  {label}: 0 exact duplicates ({len(lst)} total)")

# Check quote reuse frequency
qc = Counter([x[1] for x in all_quotes])
reused = [(k[:45], v) for k, v in qc.items() if v > 2]
print(f"Total quote lines: {len(all_quotes)}, Unique quote lines: {len(qc)}")
if reused:
    print("Quotes used >2 times:")
    for k, v in reused:
        print(f"  ({v}x) {k}...")

if all_ok:
    print("ALL CHECKS PASSED 100%!")
