"""
Rewards program earnings engine.

RULES IMPLEMENTED (confirmed with Lori):

1. TIER = number of "gate slots" (fert, chem, seed, corn, beans_wheat, energy)
   that have independently reached a given tier's threshold. Bronze/Silver need
   2+ slots at that tier's threshold; Gold needs 3+.

2. ENERGY SLOT: counts as complete at tier T only if refined_fuels has reached
   its tier-T threshold AND (propane OR lubes has reached its tier-T threshold).
   Each of refined_fuels/propane/lubes still earns independently at ITS OWN rate —
   the combined check only determines whether "energy" counts as a gate slot.

3. "QUALIFYING" categories (the ones whose thresholds were crossed, in
   transaction-date order, to satisfy the CURRENT tier's gate) earn on a
   marginal/bracket basis using their OWN cumulative spend:
     - $0 earned on cumulative dollars below the Bronze threshold
     - Bronze rate on the band between Bronze and Silver thresholds
     - Silver rate on the band between Silver and Gold thresholds
     - Gold rate on everything above the Gold threshold
   (This is Lori's chemical example: $140k spent = 0% on $40k, 1% on next $60k,
   1.5% on next $40k.)

4. Once the gate is satisfied (2, then later 3, qualifying slots reached, IN
   TRANSACTION-DATE ORDER), every OTHER category — one not needed to complete
   the gate — earns on ALL of its dollars (not just the amount over its own
   threshold) at the CURRENT active tier's rate for that category, from that
   point forward. This is not retroactive to transactions before the gate was
   satisfied.

5. As the season progresses and the gate is satisfied again at a HIGHER tier
   (Silver, then Gold), the active tier — and therefore the rate applied to
   "already-unlocked" categories — increases accordingly. A category that was
   one of the ORIGINAL 2 qualifying slots keeps using bracket/marginal earning
   on its own cumulative spend for as long as it keeps being one of the slots
   used to clear each new tier's gate; a category that becomes "unlocked"
   (full-dollar earning) stays in that mode for the rest of the season, at
   whatever the current active tier rate is — it does not need to reclear its
   own threshold to get the upgraded rate.

6. Enrollment: transactions before a participant's enroll_date are ignored
   entirely (zero spend, zero points, don't count toward any threshold).

7. Mid-season tier drop (e.g. returns/negative transactions dropping a
   category's cumulative spend below a threshold): supported via negative
   amounts in the transaction feed. The account's ACTIVE TIER is recomputed
   from scratch, live, after every transaction, based on how many slots'
   CURRENT cumulative spend meets each tier's threshold right now -- so a
   return can immediately demote the account (e.g. Silver -> Bronze) the
   moment a qualifying slot's spend drops back under that tier's minimum.
   (This is distinct from EARNING MODE per category, which is intentionally
   sticky/monotonic -- see assumption C below. Points already earned are not
   clawed back in this model — only prospective earning rate/tier status
   changes. Flag if the flyer intends actual point clawbacks; not specified.)

ASSUMPTIONS FLAGGED FOR REVIEW (not yet explicitly confirmed):
  A. Which 2 (or 3) slots become "the qualifying slots" is determined strictly
     by which slots cross the CURRENT tier's threshold first, in transaction
     date order — ties on the same date broken by transaction input order.
  B. When advancing from Bronze to Silver (or Silver to Gold), the "qualifying
     slots" for the new tier's gate are whichever slots cross the NEW tier's
     threshold first — not necessarily the same slots that qualified the
     previous tier.
  C. A category that is currently "unlocked" (earning full-dollar) and is ALSO
     one of the slots still being tracked toward the next tier's gate: it
     keeps earning full-dollar at the current active tier's rate throughout
     (assumption 5 above) rather than reverting to bracket-only. This matters
     for e.g. a category that helped clear Bronze, was therefore "qualifying"
     (bracket-only) at Bronze, then becomes one of the 2 that clears Silver
     too — in which case rule 3 (bracket earning) continues to apply to it at
     Silver, not rule 4.
"""

from dataclasses import dataclass, field
from datetime import date
from typing import Optional

from config import (
    CATEGORIES, CATEGORY_TO_SLOT, SLOTS, TIERS,
    CATEGORY_THRESHOLDS, CATEGORY_RATES, GATE_REQUIREMENT, MAX_TOTAL_POINTS,
)

TIER_RANK = {None: 0, "bronze": 1, "silver": 2, "gold": 3}


@dataclass
class Transaction:
    date: date
    category: str
    amount: float
    txn_id: Optional[str] = None


@dataclass
class TxnResult:
    txn_id: Optional[str]
    date: date
    category: str
    amount: float
    cum_before: float
    cum_after: float
    earning_quantity: float  # portion of `amount` that actually earned (nonzero-rate dollars)
    points_earned: float     # = earnings = earning_quantity blended against the rate(s) it crossed
    effective_rate: Optional[float]  # points_earned / earning_quantity; None if nothing earned
    active_tier_after: Optional[str]
    category_mode: str  # "locked/bracket" | "bracket" | "full"
    running_totals: dict  # snapshot of cumulative spend in EVERY category as of this transaction


ENERGY_SUBITEMS = ("refined_fuels", "propane", "lubes")


def _energy_combo_ok(category: str, cum_context: dict, tier: str) -> bool:
    """
    For refined_fuels/propane/lubes ONLY: is the OTHER required side of the
    combo also satisfied at this tier, using the OTHER categories' current
    cumulative amounts (cum_context) -- refined_fuels needs propane OR
    lubes; propane and lubes each individually need refined_fuels.

    This gates a category's OWN bracket-based earning only (see rule 3 in
    the module docstring) -- it does NOT apply to full-dollar mode, where a
    category is riding on a gate already satisfied by two OTHER categories
    and doesn't need to independently qualify at all (confirmed with Lori).
    """
    if category == "refined_fuels":
        return (
            cum_context["propane"] >= CATEGORY_THRESHOLDS["propane"][tier]
            or cum_context["lubes"] >= CATEGORY_THRESHOLDS["lubes"][tier]
        )
    elif category in ("propane", "lubes"):
        return cum_context["refined_fuels"] >= CATEGORY_THRESHOLDS["refined_fuels"][tier]
    return True  # not an energy sub-item -- no gating


def _bracket_earning(
    category: str, prev_cum: float, new_cum: float, cum_context: Optional[dict] = None
) -> tuple[float, float]:
    """
    Marginal/bracket earning for a category using ITS OWN cumulative thresholds.
    $0 below bronze threshold; bronze rate between bronze & silver thresholds;
    silver rate between silver & gold thresholds; gold rate above gold threshold.
    Handles a single transaction spanning multiple brackets.

    For refined_fuels/propane/lubes: each nonzero band ALSO requires the
    energy combo condition to hold at that band's tier (using cum_context,
    the OTHER two energy sub-items' current amounts) -- otherwise that band
    earns $0, even though this category's own dollars are technically past
    that threshold. This only applies here, in bracket mode; full-dollar
    mode (see simulate()) has no such gating.

    Returns (points_earned, earning_quantity) where earning_quantity is the
    slice of the transaction that fell in a nonzero-rate bracket (i.e.
    excludes the portion below the Bronze threshold, which earns $0). If a
    transaction spans multiple nonzero brackets (e.g. crosses from Bronze
    into Silver rate), points_earned / effective rate is a blend across them.
    """
    th = CATEGORY_THRESHOLDS[category]
    rt = CATEGORY_RATES[category]
    bounds = [0, th["bronze"], th["silver"], th["gold"], float("inf")]
    rates = [0.0, rt["bronze"], rt["silver"], rt["gold"]]
    band_tiers = [None, "bronze", "silver", "gold"]

    earned = 0.0
    earning_quantity = 0.0
    for i in range(4):
        lo, hi = bounds[i], bounds[i + 1]
        overlap_lo = max(prev_cum, lo)
        overlap_hi = min(new_cum, hi)
        if overlap_hi > overlap_lo:
            segment = overlap_hi - overlap_lo
            rate = rates[i]
            if rate > 0 and cum_context is not None and category in ENERGY_SUBITEMS:
                if not _energy_combo_ok(category, cum_context, band_tiers[i]):
                    rate = 0.0
            earned += segment * rate
            if rate > 0:
                earning_quantity += segment
    return earned, earning_quantity


def _slot_tier_reached(slot: str, cum: dict) -> Optional[str]:
    """Highest tier this slot's cumulative spend currently satisfies."""
    if slot == "energy":
        best = None
        for tier in TIERS:
            rf_ok = cum["refined_fuels"] >= CATEGORY_THRESHOLDS["refined_fuels"][tier]
            side_ok = (
                cum["propane"] >= CATEGORY_THRESHOLDS["propane"][tier]
                or cum["lubes"] >= CATEGORY_THRESHOLDS["lubes"][tier]
            )
            if rf_ok and side_ok:
                best = tier
        return best
    else:
        best = None
        for tier in TIERS:
            if cum[slot] >= CATEGORY_THRESHOLDS[slot][tier]:
                best = tier
        return best


def simulate(
    transactions: list[Transaction],
    enroll_date: Optional[date] = None,
) -> list[TxnResult]:
    """
    Run the full earnings simulation for one participant, in date order.
    Returns one TxnResult per input transaction (ignored/pre-enrollment
    transactions still appear, with points_earned=0 and category_mode="locked").
    """
    ordered = sorted(
        enumerate(transactions), key=lambda pair: (pair[1].date, pair[0])
    )

    cum = {c: 0.0 for c in CATEGORIES}
    # qualifying_slots[tier] = set of slots credited as "qualifying" (bracket-mode)
    # for that tier's gate, in the order they completed.
    qualifying_slots: dict[str, list[str]] = {t: [] for t in TIERS}
    active_tier: Optional[str] = None

    results: list[Optional[TxnResult]] = [None] * len(transactions)

    for orig_idx, txn in ordered:
        if enroll_date is not None and txn.date < enroll_date:
            results[orig_idx] = TxnResult(
                txn.txn_id, txn.date, txn.category, txn.amount,
                cum_before=cum[txn.category], cum_after=cum[txn.category],
                earning_quantity=0.0, points_earned=0.0, effective_rate=None,
                active_tier_after=active_tier,
                category_mode="locked (pre-enrollment)",
                running_totals=dict(cum),
            )
            continue

        prev_cum = cum[txn.category]
        new_cum = prev_cum + txn.amount
        slot = CATEGORY_TO_SLOT[txn.category]

        # Is this category's slot one of the qualifying slots for the tier it's
        # currently helping to clear (or already cleared)? If so -> bracket mode.
        is_qualifying = any(slot in qualifying_slots[t] for t in TIERS)

        if is_qualifying:
            points, earning_qty = _bracket_earning(txn.category, prev_cum, new_cum, cum_context=cum)
            mode = "bracket"
        elif active_tier is not None:
            # Gate already satisfied by other slots -> full-dollar earning at
            # the current active tier's rate for this category. No combo
            # gating here even for energy sub-items -- riding on a gate
            # already met by two OTHER categories needs no independent
            # qualification (confirmed with Lori).
            points = txn.amount * CATEGORY_RATES[txn.category][active_tier]
            earning_qty = txn.amount
            mode = "full"
        else:
            # No tier active yet, and this category isn't (yet) a qualifying
            # slot -> only earns once it independently clears bronze, same as
            # bracket mode's $0 floor. Use bracket formula (gives 0 below
            # bronze threshold) until we know whether it becomes qualifying.
            points, earning_qty = _bracket_earning(txn.category, prev_cum, new_cum, cum_context=cum)
            mode = "locked/bracket"

        effective_rate = (points / earning_qty) if earning_qty > 0 else None

        cum[txn.category] = new_cum

        # qualifying_slots is intentionally MONOTONIC (never removes a slot) --
        # it tracks which categories have ever earned "bracket" mode for
        # EARNING-MODE purposes (rule 5's assumption C: once a category
        # starts bracket-earning, it keeps bracket-earning for the rest of
        # the season, even if it later helps unlock others or its own spend
        # dips). This does NOT determine the account's active tier -- see
        # below.
        for check_tier in TIERS:
            if len(qualifying_slots[check_tier]) >= GATE_REQUIREMENT[check_tier]:
                continue  # already fully qualified for this tier
            for s in SLOTS:
                if s in qualifying_slots[check_tier]:
                    continue
                if _slot_tier_reached(s, cum) is not None and (
                    TIER_RANK[_slot_tier_reached(s, cum)] >= TIER_RANK[check_tier]
                ):
                    qualifying_slots[check_tier].append(s)
                    if len(qualifying_slots[check_tier]) >= GATE_REQUIREMENT[check_tier]:
                        break

        # Active tier, by contrast, is recomputed LIVE from current
        # cumulative spend every transaction -- NOT from the monotonic list
        # above. This is what lets a return correctly demote the account:
        # if a slot's spend drops back under a tier's threshold, it stops
        # counting toward that tier's gate right away, per rule 6.
        active_tier = None
        for t in TIERS:
            slots_meeting_t = sum(
                1 for s in SLOTS
                if _slot_tier_reached(s, cum) is not None
                and TIER_RANK[_slot_tier_reached(s, cum)] >= TIER_RANK[t]
            )
            if slots_meeting_t >= GATE_REQUIREMENT[t]:
                active_tier = t

        results[orig_idx] = TxnResult(
            txn.txn_id, txn.date, txn.category, txn.amount,
            cum_before=prev_cum, cum_after=new_cum,
            earning_quantity=earning_qty, points_earned=points,
            effective_rate=effective_rate, active_tier_after=active_tier,
            category_mode=mode,
            running_totals=dict(cum),
        )

    return results  # type: ignore


def simulate_scenario(transactions: list[Transaction], enroll_date: Optional[date] = None) -> list[TxnResult]:
    """
    The scenario-level entry point (use this, not simulate(), for real
    scenarios) -- implements the SETTLED-TIER rule confirmed with Lori:

    - If the account's tier only ever climbs (or holds steady) for the whole
      scenario -- never drops -- earnings work exactly like simulate() above:
      each category blends across bands as its own spend climbs through them
      (the original Chemical $140,000 example: 0% / bronze rate / silver rate
      bands, each locked in as earned).

    - If a return causes the account's tier to drop AT ANY POINT in the
      scenario, the ENTIRE scenario is discarded and recalculated flat,
      using ONLY the FINAL settled tier's own threshold and rate for every
      category -- no blending across bands at all, even for a category
      whose own total individually reached a higher tier. Confirmed example:
      $150,000 Chemical, settled at Bronze -> ($150,000-$40,000) x 1% on the
      full excess, NOT split between Bronze/Silver bands.

      Within that flat recompute, categories are still split into
      "qualifying" (the first N, in date order, to independently cross the
      SETTLED tier's own threshold -- earn on excess only) vs "full-dollar"
      (gate already met by others -- earn on their ENTIRE total, no
      subtraction), mirroring simulate()'s qualifying/full-dollar split, but
      using the one settled tier as the only frame of reference.

      Every row's actual_level is also reported as this one settled tier,
      uniformly, since the whole scenario is being judged by where the
      account ends up -- not by what was true at each date historically.

    CONFIRMED: if the scenario's FINAL settled tier is None at the very end
    -- whether because a drop took it all the way down, or simply because
    the account never reached even Bronze (2 slots) in the first place --
    NOTHING earns, anywhere, for the whole scenario. This applies even to a
    category that individually crossed its own threshold along the way:
    that early crossing is only valid if the account eventually holds at
    least Bronze by the end. It was never "wrong" for the original Chemical/
    Corn Bronze example (both categories DID eventually clear together,
    validating Chemical's early excess) -- it's specifically the case where
    the gate is NEVER met, ever, in the whole scenario.
    """
    climbing_results = simulate(transactions, enroll_date=enroll_date)

    chrono_indices = sorted(range(len(transactions)), key=lambda i: (transactions[i].date, i))
    tier_sequence = [climbing_results[i].active_tier_after for i in chrono_indices]
    ranks = [TIER_RANK[t] for t in tier_sequence]
    final_tier = tier_sequence[-1] if tier_sequence else None

    if final_tier is None:
        # Gate never reached Bronze by the end of the scenario -- zero out
        # every transaction's earnings, regardless of any individual
        # category's own pre-gate bracket crossings along the way.
        zeroed = _zero_out_earnings(climbing_results)
        return _apply_points_cap(zeroed, chrono_indices)

    drop_occurred = any(ranks[i + 1] < ranks[i] for i in range(len(ranks) - 1))
    if not drop_occurred:
        return _apply_points_cap(climbing_results, chrono_indices)

    settled_results = _settle_mode_results(transactions, enroll_date, final_tier, chrono_indices)
    return _apply_points_cap(settled_results, chrono_indices)


def _zero_out_earnings(results: list[TxnResult]) -> list[TxnResult]:
    """Returns a copy of results with every transaction's earnings zeroed
    out (points_earned, earning_quantity, effective_rate) -- used when the
    scenario's final settled tier is None, meaning the 2-category gate was
    never satisfied by the end. cum_before/cum_after/running_totals are left
    as-is (still useful to see what spend happened), but nothing earned."""
    zeroed = []
    for r in results:
        zeroed.append(TxnResult(
            r.txn_id, r.date, r.category, r.amount,
            cum_before=r.cum_before, cum_after=r.cum_after,
            earning_quantity=0.0, points_earned=0.0, effective_rate=None,
            active_tier_after=r.active_tier_after,
            category_mode=r.category_mode + " (gate never met)"
            if r.category_mode != "locked (pre-enrollment)" else r.category_mode,
            running_totals=r.running_totals,
        ))
    return zeroed


def _apply_points_cap(results: list[TxnResult], chrono_indices: list[int]) -> list[TxnResult]:
    """
    Enforces MAX_TOTAL_POINTS as a running ceiling on the scenario's
    cumulative earnings, walked in chronological order:

    - A transaction that would push the running total above the cap is
      clipped down to exactly whatever headroom remains (so the running
      total lands exactly at the cap, never over) -- everything after that,
      in the same category or any other, earns $0 more until the running
      total drops back under the cap.
    - A NEGATIVE contribution (a return, in settle-mode) always applies in
      full -- the cap only throttles positive earnings, it never blocks a
      return from reducing the running total, which can free up headroom
      for later transactions to earn again.

    effective_rate is recomputed against the clipped points_earned so it
    stays internally consistent; earning_quantity is left as-is, since it
    describes how much of the transaction was eligible under the tier
    rules, which the cap doesn't change.
    """
    running_total = 0.0
    capped = list(results)  # shallow copy; we'll replace individual entries
    for i in chrono_indices:
        r = capped[i]
        contribution = r.points_earned

        if contribution > 0:
            headroom = max(0.0, MAX_TOTAL_POINTS - running_total)
            allowed = min(contribution, headroom)
        else:
            allowed = contribution  # returns/negatives always apply in full

        running_total += allowed

        if allowed != contribution:
            new_effective_rate = (allowed / r.earning_quantity) if r.earning_quantity > 0 else None
            capped[i] = TxnResult(
                r.txn_id, r.date, r.category, r.amount,
                cum_before=r.cum_before, cum_after=r.cum_after,
                earning_quantity=r.earning_quantity, points_earned=allowed,
                effective_rate=new_effective_rate, active_tier_after=r.active_tier_after,
                category_mode=r.category_mode + " (capped)",
                running_totals=r.running_totals,
            )
    return capped


def _settle_mode_results(
    transactions: list[Transaction],
    enroll_date: Optional[date],
    final_tier: Optional[str],
    chrono_indices: list[int],
) -> list[TxnResult]:
    cum = {c: 0.0 for c in CATEGORIES}
    qualifying_slots_settled: list[str] = []
    results: list[Optional[TxnResult]] = [None] * len(transactions)

    for i in chrono_indices:
        txn = transactions[i]

        if enroll_date is not None and txn.date < enroll_date:
            results[i] = TxnResult(
                txn.txn_id, txn.date, txn.category, txn.amount,
                cum_before=cum[txn.category], cum_after=cum[txn.category],
                earning_quantity=0.0, points_earned=0.0, effective_rate=None,
                active_tier_after=final_tier, category_mode="locked (pre-enrollment)",
                running_totals=dict(cum),
            )
            continue

        prev_cum = cum[txn.category]
        new_cum = prev_cum + txn.amount
        slot = CATEGORY_TO_SLOT[txn.category]

        if final_tier is None:
            # Flagged assumption: no tier held at settlement -> nobody earns.
            points, earning_qty = 0.0, 0.0
            mode = "settled/no-tier"
        else:
            threshold = CATEGORY_THRESHOLDS[txn.category][final_tier]
            rate = CATEGORY_RATES[txn.category][final_tier]
            is_qualifying_settled = slot in qualifying_slots_settled
            gate_already_met = len(qualifying_slots_settled) >= GATE_REQUIREMENT[final_tier]

            if is_qualifying_settled or not gate_already_met:
                # Bracket-excess above the ONE settled threshold (degenerates
                # to 0 if this transaction doesn't push cum past it).
                overlap = max(0.0, new_cum - threshold) - max(0.0, prev_cum - threshold)
                rate_to_use = rate
                if overlap > 0 and txn.category in ENERGY_SUBITEMS:
                    if not _energy_combo_ok(txn.category, cum, final_tier):
                        rate_to_use = 0.0
                points = overlap * rate_to_use
                earning_qty = overlap if rate_to_use > 0 else 0.0
                mode = "settled/bracket"
            else:
                # Gate already satisfied by other slots -> full-dollar. No
                # combo gating here even for energy sub-items -- same as
                # climbing mode's full-dollar branch (confirmed with Lori).
                points = txn.amount * rate
                earning_qty = txn.amount
                mode = "settled/full"

        effective_rate = (points / earning_qty) if earning_qty > 0 else None
        cum[txn.category] = new_cum

        if final_tier is not None and len(qualifying_slots_settled) < GATE_REQUIREMENT[final_tier]:
            for s in SLOTS:
                if s in qualifying_slots_settled:
                    continue
                reached = _slot_tier_reached(s, cum)
                if reached is not None and TIER_RANK[reached] >= TIER_RANK[final_tier]:
                    qualifying_slots_settled.append(s)
                    if len(qualifying_slots_settled) >= GATE_REQUIREMENT[final_tier]:
                        break

        results[i] = TxnResult(
            txn.txn_id, txn.date, txn.category, txn.amount,
            cum_before=prev_cum, cum_after=new_cum,
            earning_quantity=earning_qty, points_earned=points,
            effective_rate=effective_rate, active_tier_after=final_tier,
            category_mode=mode,
            running_totals=dict(cum),
        )

    return results  # type: ignore