
from hvdc.extractors.case import extract_cases

SAMPLES = [
    "RE: HVDC-ADOPT-HE-0476 - JPTW-71 / GRM-123",
    "Fwd: HVDC-AGI-OPS-AB12 - schedule",
    "PRL-O-046-O4(HE-0486) | PO 5001005009 | AGI",
    "HVDC-ADOPT-SCT-0136 HVDC-ADOPT-SCT-0136 / GRM-123",
]

def _vals(lst):
    return [(h["kind"], h["value"]) for h in lst]

def test_extract_cases_basic():
    out = extract_cases(SAMPLES[0])
    assert ("HVDC_ADOPT", "HVDC-ADOPT-HE-0476") in _vals(out)
    assert ("JPTW", "JPTW-71") in _vals(out)
    assert ("GRM", "GRM-123") in _vals(out)

def test_extract_cases_generic():
    out = extract_cases(SAMPLES[1])
    assert ("HVDC_GENERIC", "HVDC-AGI-OPS-AB12") in _vals(out)

def test_extract_cases_prl():
    out = extract_cases(SAMPLES[2])
    assert ("PRL", "PRL-O-046-O4(HE-0486)") in _vals(out)

def test_extract_cases_dedup():
    out = extract_cases(SAMPLES[3])
    vals = _vals(out)
    assert vals.count(("HVDC_ADOPT", "HVDC-ADOPT-SCT-0136")) == 1
