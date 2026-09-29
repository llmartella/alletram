from datetime import date
from engine import Transaction, simulate


def test_bronze_two_category_example():
    """
    From Lori's spec:
      - chem spend $50,000 -> earn on $10,000 (over the $40k bronze threshold)
      - corn spend $25,000 -> earn on $5,000 (over the $20k bronze threshold)
      - at this point Bronze is reached (2 categories)
      - $10,000 on beans_wheat -> earns 0.012 on ALL $10,000, because the
        minimum was already met in 2 categories.
    Transactions occur in date order: chem, then corn, then beans_wheat.
    """
    txns = [
        Transaction(date(2026, 10, 1), "chem", 50000, "t1"),
        Transaction(date(2026, 10, 2), "corn", 25000, "t2"),
        Transaction(date(2026, 10, 3), "beans_wheat", 10000, "t3"),
    ]
    results = simulate(txns, enroll_date=date(2026, 9, 15))
    by_id = {r.txn_id: r for r in results}

    # chem: bracket earning, only the $10,000 over the $40k bronze threshold,
    # at bronze chem rate (0.010) = $100
    assert round(by_id["t1"].earning_quantity, 2) == 10000.00
    assert round(by_id["t1"].points_earned, 2) == 100.00
    assert round(by_id["t1"].effective_rate, 4) == 0.0100
    assert by_id["t1"].category_mode in ("locked/bracket", "bracket")

    # corn: bracket earning, only $5,000 over $20k bronze threshold,
    # at bronze corn rate (0.005) = $25
    assert round(by_id["t2"].earning_quantity, 2) == 5000.00
    assert round(by_id["t2"].points_earned, 2) == 25.00
    assert round(by_id["t2"].effective_rate, 4) == 0.0050
    assert by_id["t2"].active_tier_after == "bronze"  # gate now satisfied

    # beans_wheat: gate already satisfied by chem+corn -> full-dollar earning
    # on all $10,000 at bronze beans_wheat rate (0.012) = $120
    assert round(by_id["t3"].earning_quantity, 2) == 10000.00
    assert round(by_id["t3"].points_earned, 2) == 120.00
    assert round(by_id["t3"].effective_rate, 4) == 0.0120
    assert by_id["t3"].category_mode == "full"


def test_chemical_marginal_bracket_example():
    """
    From Lori's spec: chemical alone reaching $140,000 cumulative (assume it's
    one of the qualifying/bracket categories throughout) earns:
      - 0% on the first $40,000 (below bronze threshold)
      - 1.0% on the next $60,000 (from $40k to $100k -> bronze rate, since that
        band sits between the bronze and silver thresholds)
      - 1.5% on the remaining $40,000 (from $100k to $140k -> silver rate,
        since that band sits between the silver and gold thresholds)
    Total = 0 + 600 + 600 = $1,200
    Single lump-sum transaction to isolate the bracket math itself.
    """
    txns = [Transaction(date(2026, 10, 1), "chem", 140000, "t1")]
    results = simulate(txns, enroll_date=date(2026, 9, 15))
    r = results[0]
    # earning_quantity excludes the first $40,000 (0% bracket) -> $100,000 earned
    assert round(r.earning_quantity, 2) == 100000.00
    assert round(r.points_earned, 2) == 1200.00
    # blended rate across the 1.0% and 1.5% bands: 1200 / 100000 = 0.012
    assert round(r.effective_rate, 4) == 0.0120


def test_pre_enrollment_activity_ignored():
    txns = [
        Transaction(date(2026, 9, 20), "chem", 50000, "before"),
        Transaction(date(2026, 10, 5), "corn", 25000, "after"),
    ]
    results = simulate(txns, enroll_date=date(2026, 10, 1))
    by_id = {r.txn_id: r for r in results}
    assert by_id["before"].points_earned == 0.0
    assert by_id["before"].category_mode == "locked (pre-enrollment)"
    # corn, on its own, hasn't hit bronze gate (only 1 qualifying category so far)
    assert by_id["after"].active_tier_after is None


def test_energy_gate_requires_refined_plus_one_side():
    """
    Propane alone hitting its bronze threshold should NOT count as a completed
    energy slot -- refined_fuels must ALSO be met (with propane or lubes).
    """
    txns = [
        Transaction(date(2026, 10, 1), "propane", 5000, "t1"),   # meets propane bronze alone
        Transaction(date(2026, 10, 2), "chem", 40000, "t2"),     # meets chem bronze
    ]
    results = simulate(txns, enroll_date=date(2026, 9, 15))
    # only chem slot qualifies; propane alone doesn't complete "energy" ->
    # gate needs 2 slots, only 1 (chem) present -> no active tier yet
    assert results[-1].active_tier_after is None

    txns2 = txns + [
        Transaction(date(2026, 10, 3), "refined_fuels", 10000, "t3"),  # now energy slot completes
    ]
    results2 = simulate(txns2, enroll_date=date(2026, 9, 15))
    assert results2[-1].active_tier_after == "bronze"


if __name__ == "__main__":
    test_bronze_two_category_example()
    test_chemical_marginal_bracket_example()
    test_pre_enrollment_activity_ignored()
    test_energy_gate_requires_refined_plus_one_side()
    print("All tests passed.")