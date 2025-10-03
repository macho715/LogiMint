
from hvdc.extractors.case import extract_cases
from hvdc.extractors.site import extract_sites
from hvdc.extractors.lpo import extract_lpos

samples = [
    "RE: HVDC-ADOPT-HE-0476 - JPTW-71 / GRM-123 (urgent)",
    "PRL-O-046-O4(HE-0486) - PO 5001005009 AGI",
    "HVDC-AGI-OPS-AB12: phase-2 update",
]

for s in samples:
    print("-----")
    print(s)
    print("cases:", [(h['kind'], h['value']) for h in extract_cases(s)])
    print("sites:", extract_sites(s))
    print("lpos :", extract_lpos(s))
