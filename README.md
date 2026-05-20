# Rate Matrix Builder

A Streamlit app that turns a rate-card Excel into carrier pricing matrices.

## Files

| File | Purpose |
|------|---------|
| `app.py` | Streamlit UI — the only file managers interact with |
| `pipeline.py` | Core logic: build, compute, optimize, write Excel |
| `master_parser.py` | Parses the DSV master rate card (all countries in one file) |
| `requirements.txt` | Python dependencies |

## Two input formats (auto-detected)

**1. Master file** — the DSV "MDK - FENDER - PARCEL RATES" workbook with one sheet
per rate table and all countries in one file. Upload this and the app:
- reads every carrier and country in one pass
- shows only the countries actually present in the file
- reads **MAUT surcharges per country, per carrier** directly from the file
  (the MAUT inputs disappear from the sidebar)
- you still set **fuel %** in the sidebar (fuel is a separate monthly surcharge)

This replaces the old manual step of copying rates into per-country files.

**2. Per-country file** — the older one-country-per-workbook format (tabs named
UPSDE, DHL, DPD, UPSNL, POSTNORD). Still fully supported.

## Carriers supported

UPS DE, UPS NL, UPS GB (UK domestic), DHL, DPD, PostNord.

Special handling baked in:
- **UPS WorldEase (WEA)** — flat per-parcel rate for CH and NO
- **DHL BNL** — "1st parcel + each additional" pricing for BE, LU, NL
- **PostNord** — flat rate per service (B2B / Home / PUDO), multi-country table
- **UPS DE zones** — postcode/zone-based, including alphanumeric zones (ES4/ES5/ES6)

## Deploy to Streamlit Cloud (free)

1. Create a GitHub repo (can be private)
2. Upload all files to the root: `app.py`, `pipeline.py`, `master_parser.py`,
   `requirements.txt`, `README.md`
3. Go to [share.streamlit.io](https://share.streamlit.io), sign in with GitHub
4. New app → pick repo → branch `main` → main file `app.py` → Deploy
5. Share the URL

## Using the app

1. **Upload** the rate card (master or per-country — detected automatically)
2. **Adjust fuel %** in the sidebar if rates changed
3. **Select countries** (only the ones in the file are shown for a master file)
4. Optionally open **Advanced** to change carriers / postcode ranges per country
5. **▶ Generate matrices**
6. **Download** per country or everything as a ZIP

Each country yields three files: `*_extended` (all rows), `*_optimized`
(per-carrier dominance removed), `*_minimal` (cross-carrier dominance removed —
the one to use).

## Changing defaults

Carrier defaults and country configs live in `pipeline.py`
(`CARRIER_DEFAULTS`, `COUNTRY_CONFIG`). The master-file sheet names and parsing
logic live in `master_parser.py`. Edit and redeploy.
