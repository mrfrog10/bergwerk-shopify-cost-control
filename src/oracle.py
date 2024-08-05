"""Independent plain-Python earned value calculation.

Used by the tests as a second opinion: the Excel formulas must give the
same answers as this code for any scenario in inputs.py.
"""
import calendar
from datetime import timedelta


def compute(s):
    n = len(s["wps"])
    months = s["months"]
    approved = [c for c in s["changes"] if c[6] == "Approved"]
    wp_ids = [w[0] for w in s["wps"]]

    budget = [w[4] for w in s["wps"]]
    for c in approved:
        budget[wp_ids.index(c[3])] += c[4]
    bac = sum(budget)
    reserve_used = sum(c[4] for c in approved if c[5] == "Reserve")
    client_funded = sum(c[4] for c in approved if c[5] != "Reserve")
    reserve_left = s["reserve"] - reserve_used
    current_total = s["contract_budget"] + client_funded

    pv = [row[:] for row in s["pv"]]
    by_id = {c[0]: c for c in approved}
    for cr, (wi, mi) in s["cr_phasing"].items():
        if cr in by_id:
            pv[wi][mi] += by_id[cr][4]
    elapsed = [m <= s["status"] for m in months]
    pv_wp = [sum(v for v, e in zip(pv[i], elapsed) if e) for i in range(n)]
    ac_wp = [sum((v or 0) for v, e in zip(s["ac"][i], elapsed) if e)
             for i in range(n)]
    ev_wp = [budget[i] * s["pct_complete"][i] for i in range(n)]
    PV, AC, EV = sum(pv_wp), sum(ac_wp), sum(ev_wp)
    cpi, spi = EV / AC, EV / PV

    eac_wp = [ac_wp[i] + s["etc"][i] for i in range(n)]
    EAC = sum(eac_wp)
    vac_wp = [eac_wp[i] - budget[i] for i in range(n)]
    overrun = EAC - bac - reserve_left

    cum, run = [], 0
    for k in range(len(months)):
        run += sum(pv[i][k] for i in range(n))
        cum.append(run)
    at = sum(elapsed)
    full = sum(c <= EV for c in cum)
    if full == 0:
        es = EV / cum[0]
    elif full >= len(cum):
        es = float(len(cum))
    else:
        es = full + (EV - cum[full - 1]) / (cum[full] - cum[full - 1])
    days_in_month = calendar.monthrange(s["status"].year,
                                        s["status"].month)[1]
    lag_days = round((at - es) * days_in_month + 1e-9)
    # a forecast launch cannot lie in the past
    forecast_launch = max(s["launch"] + timedelta(days=lag_days),
                          s["status"] + timedelta(days=1))
    # aftercare runs a fixed number of weeks after the (forecast) launch
    forecast_end = forecast_launch + timedelta(weeks=s["aftercare_weeks"])

    if overrun > 0:
        cost_rag = "RED"
    elif EAC - bac > 0:
        cost_rag = "AMBER"
    else:
        cost_rag = "GREEN"
    sched_rag = ("GREEN" if spi >= s["spi_green"] else
                 "AMBER" if spi >= s["spi_amber"] else "RED")

    return {
        "BAC": bac, "PV": PV, "EV": EV, "AC": AC, "AC_MINUS_PV": AC - PV,
        "CPI": cpi, "SPI": spi, "EAC": EAC, "VAC": EAC - bac,
        "RESERVE_LEFT": reserve_left, "CUR_TOTAL": current_total,
        "OVERRUN": overrun, "OVERRUN_PCT": overrun / current_total,
        "ES": es, "AT": at, "LAG_DAYS": lag_days,
        "FORECAST_LAUNCH": forecast_launch, "FORECAST_END": forecast_end,
        "EAC_OPT": AC + (bac - EV), "EAC_CPI": bac / cpi,
        "EAC_CPISPI": AC + (bac - EV) / (cpi * spi),
        "TCPI_BAC": (bac - EV) / (bac - AC),
        "TCPI_EAC": (bac - EV) / (EAC - AC),
        "COST_RAG": cost_rag, "SCHED_RAG": sched_rag,
        "COMMITTED": sum(c[2] for c in s["commitments"]),
        "wp_budget": budget, "wp_ev": ev_wp, "wp_pv": pv_wp,
        "wp_ac": ac_wp, "wp_eac": eac_wp, "wp_vac": vac_wp,
    }
