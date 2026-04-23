"""
LOB (Line of Business) Registry for NL-39.

In NL-39, LOBs appear as ROW LABELS (col 1) — opposite to NL-4/NL-5/NL-7
where LOBs are column headers. The canonical LOB keys are consistent with
the other extractors for cross-form joins in the master output.

NL-39 uses 15 LOB rows. No sub-group totals (no total_marine, total_motor,
total_health, grand_total) — those don't appear in the ageing schedule.
"""

# Canonical LOB keys — ordered to match the PDF row order.
LOB_ORDER = [
    "fire",
    "marine_cargo",
    "marine_hull",              # "Marine Other than Cargo" in NL-39
    "motor_od",
    "motor_tp",
    "health",
    "personal_accident",
    "travel_insurance",
    "wc_el",                    # Workmen's Compensation / Employer's Liability
    "public_product_liability",
    "engineering",
    "aviation",
    "crop_insurance",
    "other_segments",
    "total_miscellaneous",      # "Miscellaneous" in NL-39 (catch-all row)
]

# Display-friendly names for each canonical LOB key.
LOB_DISPLAY_NAMES = {
    "fire":                     "Fire",
    "marine_cargo":             "Marine Cargo",
    "marine_hull":              "Marine Other than Cargo",
    "motor_od":                 "Motor OD",
    "motor_tp":                 "Motor TP",
    "health":                   "Health",
    "personal_accident":        "Personal Accident",
    "travel_insurance":         "Travel",
    "wc_el":                    "WC / Employer's Liability",
    "public_product_liability": "Public / Product Liability",
    "engineering":              "Engineering",
    "aviation":                 "Aviation",
    "crop_insurance":           "Crop Insurance",
    "other_segments":           "Other Segments",
    "total_miscellaneous":      "Miscellaneous",
}

# Maps every observed NL-39 row label (normalised, lowercase) → canonical LOB key.
LOB_ALIASES = {
    # Fire
    "fire":                                             "fire",
    # Marine
    "marine cargo":                                     "marine_cargo",
    "marine other than cargo":                          "marine_hull",
    "marine (other than cargo)":                        "marine_hull",
    # Motor
    "motor od":                                         "motor_od",
    "motor ownerdamage":                                "motor_od",
    "motor tp":                                         "motor_tp",
    "motor thirdparty":                                 "motor_tp",
    # Health
    "health":                                           "health",
    "personal accident":                                "personal_accident",
    "travel":                                           "travel_insurance",
    "travel insurance":                                 "travel_insurance",
    # Liability / Misc
    "workmen's compensation/ employer's liability":     "wc_el",
    "workmen\u2019s compensation/ employer\u2019s liability": "wc_el",
    "workmen's compensation/employer's liability":      "wc_el",
    "workmen’s compensation / employer’s liability":    "wc_el",
    "workmen's compensation / employer's liability":    "wc_el",
    "workmen’s compensation /":                         "wc_el",
    "workmen's compensation /":                         "wc_el",
    "workmen’s compensation/":                          "wc_el",
    "workmen's compensation/":                          "wc_el",
    "workmen's compensation":                           "wc_el",
    "wc/el":                                            "wc_el",
    "public/ product liability":                        "public_product_liability",
    "public/product liability":                         "public_product_liability",
    "public / product liability":                        "public_product_liability",
    "public /product liability":                        "public_product_liability",
    "epmubplilco/y eprr’osd liuacbti llitiaybility":    "public_product_liability",
    "engineering":                                      "engineering",
    "aviation":                                         "aviation",
    "crop insurance":                                   "crop_insurance",
    "crop":                                             "crop_insurance",
    "weather / crop insurance":                         "crop_insurance",
    "weather/crop insurance":                           "crop_insurance",
    # Segments
    "other segments (a)":                               "other_segments",
    "other segments":                                   "other_segments",
    "others (a)":                                       "other_segments",   # Magma
    "other liability":                                  "other_segments",   # Shriram
    "miscellaneous":                                    "total_miscellaneous",
    "total miscellaneous":                              "total_miscellaneous",
    "others":                                           "total_miscellaneous",
}

# Company-specific LOB alias overrides (populated as quirks are discovered).
COMPANY_SPECIFIC_ALIASES: dict = {}
