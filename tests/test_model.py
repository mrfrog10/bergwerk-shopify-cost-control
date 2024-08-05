"""Tests: the Excel formulas must agree with an independent Python
calculation (src/oracle.py) for more than one scenario, and the integrity
checks must catch bad inputs.

Run:  python -m pytest -q        (or)  python tests/test_model.py
Needs LibreOffice for recalculation.
"""
import sys
from copy import deepcopy
from pathlib import Path

from openpyxl import load_workbook

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))

from build_workbook import build  # noqa: E402
from inputs import BASE, scenario_what_if  # noqa: E402
from oracle import compute  # noqa: E402
from verify_workbook import formula_errors, named, recalc  # noqa: E402

TMP = ROOT / "build" / "tests"
SCALARS = ["BAC", "PV", "EV", "AC", "CPI", "SPI", "EAC", "VAC",
           "RESERVE_LEFT", "CUR_TOTAL", "OVERRUN", "OVERRUN_PCT", "ES", "AT",
           "LAG_DAYS", "EAC_OPT", "EAC_CPI", "EAC_CPISPI", "TCPI_BAC",
           "TCPI_EAC", "COST_RAG", "SCHED_RAG", "COMMITTED"]


def run(scenario, tag):
    src = build(scenario, TMP / f"{tag}.xlsx")
    return load_workbook(recalc(src, TMP / f"{tag}_calc.xlsx"),
                         data_only=True)


def assert_matches_oracle(wb, scenario):
    exp = compute(scenario)
    for key in SCALARS:
        got = named(wb, key)
        want = exp[key]
        if isinstance(want, str):
            assert got == want, f"{key}: {got} != {want}"
        else:
            assert abs(got - want) < 1e-6, f"{key}: {got} != {want}"
    assert named(wb, "FORECAST_LAUNCH").date() == exp["FORECAST_LAUNCH"]
    assert named(wb, "FORECAST_END").date() == exp["FORECAST_END"]
    assert named(wb, "INTEGRITY") == "ALL PASS"
    assert not formula_errors(wb), formula_errors(wb)[:5]


def test_base_scenario_matches_oracle_and_brief():
    wb = run(BASE, "base")
    assert_matches_oracle(wb, BASE)
    assert named(wb, "BRIEF_CHECKS") == "ALL PASS"


def test_what_if_scenario_matches_oracle():
    s = scenario_what_if()
    wb = run(s, "what_if")
    assert_matches_oracle(wb, s)
    # CR-02 approved and client-funded: contract value rises, reserve does not
    assert named(wb, "CUR_TOTAL") == 62500
    assert named(wb, "RESERVE_LEFT") == 2500


def test_integrity_checks_catch_bad_inputs():
    s = deepcopy(BASE)
    s["pct_complete"][2] = 1.2            # >100% complete
    s["etc"][4] = 1000                    # below the 2,000 QA commitment
    s["changes"][1] = s["changes"][1][:6] + ("Approved", "no phasing")
    wb = run(s, "bad")                    # CR-02 approved without phasing
    assert named(wb, "INTEGRITY") == "REVIEW"
    ck = wb["Checks"]
    failed = {ck[f"A{r}"].value for r in range(1, ck.max_row + 1)
              if ck[f"E{r}"].value == "FAIL"}
    for expected in ("Revised PV per WP equals current budget",
                     "Approved CRs have a phasing month",
                     "% complete between 0% and 100%",
                     "ETC covers committed costs per WP"):
        assert expected in failed, f"not caught: {expected}"


if __name__ == "__main__":
    for t in (test_base_scenario_matches_oracle_and_brief,
              test_what_if_scenario_matches_oracle,
              test_integrity_checks_catch_bad_inputs):
        t()
        print(f"PASS {t.__name__}")
