"""
Rewards program configuration — pulled directly from the Tier Requirements flyer.

UNITS:
  - All thresholds are in dollars EXCEPT "fert", which is in tons.
  - Rates for fert are $-per-ton; all other rates are cents-on-the-dollar (fractions).

SEASON (confirmed by Lori):
  - Enrollment opens 9/15/2026. Points only accrue on activity AFTER a participant
    enrolls (no retroactive credit for pre-enrollment spend).
  - SEASON_ROLLOVER_DATE (currently 8/31/2027, UNCONFIRMED) is the date whose tier
    becomes the participant's starting tier for the next season. Kept as a single
    named constant below so it can be corrected without touching logic.
  - SEASON_END_DATE (10/31/2027) is the last day of the earning season.
  - POINTS_EXPIRATION_DATE (3/31/2028) — unused points are forfeited after this date.
"""

from datetime import date

TIERS = ["bronze", "silver", "gold"]

# How many gate "slots" must be satisfied (at that tier's thresholds) to be IN that tier.
GATE_REQUIREMENT = {"bronze": 2, "silver": 2, "gold": 3}

MAX_TOTAL_POINTS = 15000.0  # hard ceiling on total points earned per scenario, confirmed by Lori
SEASON_ROLLOVER_DATE = date(2027, 8, 31)  # UNCONFIRMED — update here if corrected
SEASON_END_DATE = date(2027, 8, 31)
POINTS_EXPIRATION_DATE = date(2028, 3, 31)

# Base categories as they appear in transaction data.
# "energy" is NOT a transaction category itself — refined_fuels / propane / lubes are.
CATEGORIES = [
    "fert", "chem", "seed", "corn", "beans_wheat",
    "refined_fuels", "propane", "lubes",
]

# Which "gate slot" each category counts toward for the 2+/3+ category requirement.
# fert/chem/seed/corn/beans_wheat are each their own slot.
# refined_fuels + propane + lubes all roll up into a single "energy" slot.
CATEGORY_TO_SLOT = {
    "fert": "fert",
    "chem": "chem",
    "seed": "seed",
    "corn": "corn",
    "beans_wheat": "beans_wheat",
    "refined_fuels": "energy",
    "propane": "energy",
    "lubes": "energy",
}
SLOTS = ["fert", "chem", "seed", "corn", "beans_wheat", "energy"]

CATEGORY_THRESHOLDS = {
    "fert":          {"bronze": 175,   "silver": 400,    "gold": 800},
    "chem":          {"bronze": 40000, "silver": 100000, "gold": 200000},
    "seed":          {"bronze": 45000, "silver": 120000, "gold": 225000},
    "corn":          {"bronze": 20000, "silver": 50000,  "gold": 110000},
    "beans_wheat":   {"bronze": 16000, "silver": 40000,  "gold": 80000},
    "refined_fuels": {"bronze": 10000, "silver": 25000,  "gold": 50000},
    "propane":       {"bronze": 5000,  "silver": 8000,   "gold": 15000},
    "lubes":         {"bronze": 900,   "silver": 1500,   "gold": 3000},
}

CATEGORY_RATES = {
    "fert":          {"bronze": 2.50,  "silver": 5.00,  "gold": 7.50},   # $ per ton
    "chem":          {"bronze": 0.010, "silver": 0.015, "gold": 0.020},
    "seed":          {"bronze": 0.005, "silver": 0.010, "gold": 0.015},
    "corn":          {"bronze": 0.005, "silver": 0.010, "gold": 0.015},
    "beans_wheat":   {"bronze": 0.012, "silver": 0.018, "gold": 0.024},
    "refined_fuels": {"bronze": 0.015, "silver": 0.030, "gold": 0.045},
    "propane":       {"bronze": 0.015, "silver": 0.025, "gold": 0.030},
    "lubes":         {"bronze": 0.025, "silver": 0.040, "gold": 0.055},
}

# ENERGY gate rule: refined_fuels must be met AND (propane met OR lubes met) —
# at the SAME tier — for "energy" to count as a completed slot for that tier.