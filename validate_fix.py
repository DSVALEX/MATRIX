"""
Validate the combined-weight fix against the two invoiced shipments from
Barry's e-mail. Runs the ACTUAL patched builders in pipeline.py.

Expected freight (RATE_BASE) from the rate card:
    DHL  IT     9 boxes / 0.294832 m3  -> 73.7 kg  -> EUR 49.18
    UPSNL z4    4 boxes / 0.22593  m3  -> 37.75 kg -> EUR 72.53
"""
import openpyxl
import pipeline as pl

RATE_CARD = "MDK__FENDER__PARCEL_RATES__S2026.xlsx"
wb = openpyxl.load_workbook(RATE_CARD, data_only=True)


def tiers_from(sheet, col_label):
    ws = wb[sheet]
    rows = list(ws.iter_rows(values_only=True))
    hdr = rows[17]
    col = next(j for j, c in enumerate(hdr) if str(c) == str(col_label))
    out = []
    for row in rows[18:]:
        f, t = row[1], row[2]
        if isinstance(f, (int, float)) and isinstance(t, (int, float)) and row[col] is not None:
            out.append({"from": f, "to": t, "rate": float(row[col]), "per_kg": False})
    return out


def cheapest_covering(rows, boxes, volume, divisor):
    """Mimic CargoWrite: cheapest row whose total caps fit the shipment."""
    cand = []
    for r in rows:
        max_w = r["MAX_PARCEL"] * r["EACH_WEIGHT"]
        max_v = max_w / divisor
        if r["MAX_PARCEL"] >= boxes and max_v >= volume - 1e-9:
            cand.append((r["RATE_BASE"], round(max_w, 2)))
    return sorted(cand)[0] if cand else None


cfg = {
    "site_id": "SITE", "client_id": "FENDER",
    "max_parcel_count": 15, "max_each_weight_kg": 31.5,
    "each_weight_grid": sorted(set(list(range(1, 32)) + [31.5])),
    "postcode_prefix_range": (1, 99),
}


def check(name, got, expected):
    ok = got is not None and abs(got[0] - expected) < 0.005
    flag = "PASS" if ok else "FAIL"
    print(f"[{flag}] {name}: matrix={got[0] if got else None} "
          f"(cap {got[1] if got else '-'} kg)  verwacht={expected}")
    return ok

all_ok = True

# ---- DHL standard, Italy, 9 boxes ----
dhl_cfg = {**cfg, "iso2": "IT"}
dhl_rows = pl.build_rows_dhl({"STANDARD": tiers_from("PARCEL - DHL - Other countries", "IT")}, dhl_cfg)
got = cheapest_covering(dhl_rows, boxes=9, volume=0.294832, divisor=250)
all_ok &= check("DHL IT 9 dozen", got, 49.18)

# ---- UPS NL EXPRESS SAVER, zone 4, 4 boxes ----
nl_cfg = {**cfg, "iso2": "IT"}
nl_data = {
    "zones": [{"zone": 4, "pc_from": 0, "pc_to": 99999}],
    "rates_by_zone": {4: tiers_from("PARCEL - EXPRESS SAVER UPSNL", 4)},
}
nl_rows = pl.build_rows_upsnl(nl_data, nl_cfg)
got = cheapest_covering(nl_rows, boxes=4, volume=0.22593, divisor=167)
all_ok &= check("UPS NL z4 4 dozen", got, 72.53)

# ---- UPS DE EXPRESS SAVER, zone 3 (sanity, same shipment) -> 88.43 ----
de_tiers = tiers_from("PARCEL - EXPRESS SAVER UPSDE", 3)
de_rows = pl.build_combined_weight_rows({"CARRIER_ID": "UPDE", "COUNTRYISO2": "IT"},
                                        pl.collapse_same_rate_tiers(de_tiers),
                                        15, "EXPRESS SAVER")
got = cheapest_covering(de_rows, boxes=4, volume=0.22593, divisor=167)
all_ok &= check("UPS DE z3 4 dozen", got, 88.43)

print("\nRESULTAAT:", "ALLES KLOPT TOT OP DE CENT" if all_ok else "ER IS IETS MIS")
