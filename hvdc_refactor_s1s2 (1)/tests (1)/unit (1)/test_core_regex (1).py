
import re
from hvdc.core.regex import PATTERNS, COMPILED

def test_patterns_compile_once():
    assert set(PATTERNS.keys()) == set(COMPILED.keys())
    for k, rx in COMPILED.items():
        assert isinstance(rx, re.Pattern)

def test_hvdc_adopt_matches_sample():
    s = "Fwd: HVDC-ADOPT-HE-0476 - Docs"
    m = COMPILED["HVDC_ADOPT"].search(s)
    assert m and m.group(0).upper() == "HVDC-ADOPT-HE-0476"

def test_prl_with_parenthesis():
    s = "PRL-O-046-O4(HE-0486) - status"
    m = COMPILED["PRL"].search(s)
    assert m and m.group(0).startswith("PRL-O-046-O4")

def test_jptw_grm_pair():
    s = "… JPTW-71 / GRM-123 …"
    m = COMPILED["JPTW_GRM"].search(s)
    assert m and m.group(1) == "71" and m.group(2) == "123"
