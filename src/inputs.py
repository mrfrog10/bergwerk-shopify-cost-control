"""Project inputs for the BergWerk Outdoor Shopify migration (fictional).

The workbook builder, the independent oracle and the tests all read from
here, so a scenario is defined once.
"""
from copy import deepcopy
from datetime import date

BASE = {
    "client": "BergWerk Outdoor GmbH, Munich, Germany (fictional)",
    "project": "WooCommerce to Shopify migration with JTL-Wawi integration "
               "and German payment methods",
    "markets": "Germany, Austria, Switzerland",
    "start": date(2024, 4, 1),
    "end": date(2024, 9, 30),
    "launch": date(2024, 8, 15),
    "status": date(2024, 7, 31),
    "aftercare_weeks": 6,  # WP6 support after go-live
    "hourly_rate": 25,
    "contract_budget": 60000,
    "reserve": 4000,
    "spi_green": 0.95,
    "spi_amber": 0.80,
    "months": [date(2024, m, 1) for m in range(4, 10)],  # Apr..Sep
    "wps": [
        ("WP1", "Discovery and requirements",
         "Online workshops with BergWerk's Munich team; German legal "
         "requirements (Impressum, GDPR cookie consent, returns policy)",
         "Apr", 6000),
        ("WP2", "Design",
         "Shopify theme and user experience in German and English",
         "Apr-May", 9000),
        ("WP3", "Development and integrations",
         "Shopify theme, JTL-Wawi integration, payments (PayPal, SEPA, "
         "credit card)", "May-Jul", 21000),
        ("WP4", "Data migration",
         "Products, customers and orders from WooCommerce", "Jun-Jul", 8000),
        ("WP5", "Testing and quality checks",
         "Functional tests, payment tests, mobile checks", "Jul-Aug", 6000),
        ("WP6", "Launch, aftercare and training",
         "Go-live, 6 weeks of support, staff training", "Aug-Sep", 6000),
    ],
    # plan vs actual / forecast per work package, as of the status date:
    # (planned start, planned finish, status at status date, forecast finish)
    "wp_schedule": [
        (date(2024, 4, 1), date(2024, 4, 30),
         "Done in April; two extra workshops (GDPR, Impressum)", "Done"),
        (date(2024, 4, 1), date(2024, 5, 31),
         "Done in May; extra revision round (CR-01, approved 15 May)",
         "Done"),
        (date(2024, 5, 1), date(2024, 7, 31),
         "85% done; JTL-Wawi API problems since June", "Mid-Aug 2024"),
        (date(2024, 6, 1), date(2024, 7, 31),
         "Started late; 60% done", "Mid-Aug 2024"),
        (date(2024, 7, 1), date(2024, 8, 14),
         "Only 20% done", "Mid-Aug to 28 Aug 2024"),
        (date(2024, 8, 15), date(2024, 9, 30),
         "Not started; 6 weeks after launch",
         "29 Aug to ~10 Oct 2024"),
    ],
    "pv": [  # WP1..WP6 x Apr..Sep, original baseline
        [6000, 0, 0, 0, 0, 0],
        [3000, 6000, 0, 0, 0, 0],
        [0, 4000, 8000, 9000, 0, 0],
        [0, 0, 3000, 5000, 0, 0],
        [0, 0, 0, 2000, 4000, 0],
        [0, 0, 0, 0, 3000, 3000],
    ],
    # month phasing of APPROVED change requests: {cr_id: (wp_index, month_index)}
    "cr_phasing": {"CR-01": (1, 1)},
    "ac": [  # WP1..WP6 x Apr..Sep (None = not yet incurred)
        [6800, 0, 0, 0, None, None],
        [2500, 7200, 0, 0, None, None],
        [0, 3500, 9500, 9800, None, None],
        [0, 0, 2000, 4500, None, None],
        [0, 0, 0, 1000, None, None],
        [0, 0, 0, 0, None, None],
    ],
    "ac_notes": [
        "Two extra workshops to agree the GDPR and legal page requirements",
        "BergWerk asked for an extra design revision round",
        "Problems with the JTL-Wawi API caused rework; data migration "
        "started late",
        "Integration rework continued; product data had duplicate product "
        "codes and German umlauts (ä, ö, ü) that broke in the export file",
    ],
    "changes": [
        ("CR-01", date(2024, 5, 15), "Extra design revision round", "WP2",
         1500, "Reserve", "Approved",
         "WP2 budget becomes 10,500; reserve drops to 2,500"),
        ("CR-02", date(2024, 6, 20),
         "Add Klarna as a payment option (\"Pay later\" is very popular in "
         "Germany)", "WP3", 2500, "BergWerk", "Pending",
         "Waiting for client sign-off (decision for the project manager)"),
    ],
    "commitments": [
        ("Shopify app licence (yearly)", "WP3", 1200),
        ("Freelance QA tester contract", "WP5", 2000),
        ("English to German translation and legal-text review", "WP6", 800),
    ],
    "pct_complete": [1.00, 1.00, 0.85, 0.60, 0.20, 0.00],
    "etc": [0, 0, 4500, 3500, 5000, 6000],
    "explanations": [
        "Extra workshops for the German legal requirements (GDPR, Impressum)",
        "Under budget: the CR-01 design round cost less than approved",
        "Problems with the JTL-Wawi API caused rework on the inventory "
        "integration",
        "Migration started late; duplicate product codes and broken umlauts "
        "needed cleaning",
        "On budget so far; QA freelancer committed",
        "On budget; translation and legal-text review committed",
    ],
}


def scenario_what_if():
    """A different scenario used to prove the formulas are general:
    CR-02 approved (client-funded, phased into Aug), more progress, a
    status date one month later with August actuals."""
    s = deepcopy(BASE)
    s["changes"][1] = s["changes"][1][:6] + ("Approved",
                                             "Signed off by BergWerk")
    s["cr_phasing"]["CR-02"] = (2, 4)  # WP3, Aug
    s["status"] = date(2024, 8, 31)
    s["pct_complete"] = [1.00, 1.00, 0.97, 0.90, 0.70, 0.30]
    for i, v in enumerate([0, 0, 3600, 2100, 3900, 1700]):
        s["ac"][i][4] = v
    s["etc"] = [0, 0, 2600, 900, 1800, 4600]
    # half of the QA freelancer contract was invoiced in August
    s["commitments"][1] = ("Freelance QA tester contract", "WP5", 1000)
    return s
