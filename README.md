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


## Exceptions & buckets (add-on)

CargoWrite picks the first matrix row (cheapest-first) whose constraints all
fit an order. A constraint like a **size limit** means a row only matches
parcels at/under that size — so oversized orders need a **bucket** row lower
down that drops the limit, flags it, and adds a surcharge. Without buckets,
some orders match nothing.

In the app, open **📐 Exceptions & buckets** and edit the table — one line per
carrier (or country). Each line says: *for this carrier, the normal size limit
is X metres; anything bigger gets a €Y-per-parcel surcharge.* On Generate, every
output gets:
- the **normal limit** stamped on the cheap base rows
- a **catch-all bucket twin** (limit blank, `AWKWARD=y`, surcharge added),
  shown in **amber** for oversight

The mechanism is general: in `pipeline.py`, `apply_exceptions()` takes a list of
rule dicts (constraint column, normal value, surcharge, scope). Today the UI
exposes the size case; the same function already supports flat surcharges and
other constraint columns. Documented future buckets (parcel-count overflow,
weight overflow, postcode catch-all) are noted in the code and need only a new
rule, not new logic.

Leave the table empty for a plain matrix with no buckets.


### Future buckets (now implemented, off by default)

Two more bucket types live under the same expander, both **off by default**:

**Overflow buckets** — catch orders heavier or with more parcels than the grid.
For each parcel count *n* the matrix gets a row with `MIN_PARCEL=n`,
`MIN_WEIGHT=n×max-each-weight` and **no upper caps**, priced
`overflow rate × n + surcharge × n` and flagged. The overflow rate is the
contract heavy/per-kg rate **you enter** — it is never guessed from the grid.

**Postcode catch-all** — for zoned carriers (e.g. UPS DE), adds a blank-postcode
fallback at the worst zone's rate so a prefix not present in any zone still
matches something, flagged for review.

All three bucket types stack and are added to the *surviving* rows after
optimization, keeping the final list as short as possible. The engine functions
are `apply_exceptions`, `add_overflow_buckets`, and `add_postcode_catchall` in
`pipeline.py` — each takes generic rule dicts, so new constraint columns or
scopes need a new rule, not new logic.


## Pallets (major add-on)

The builder now handles **pallets** as a second shipping mode alongside parcels,
in the same output matrix. Pallets use carrier **DHL-FENDER** (service
EUROCONNECT) and are priced on two axes — destination **postcode zone** and a
**weight bucket** — instead of the parcel each-weight/parcel-count grid.

### Pallet rate card (separate upload)

Pallets come from their own file (the DHL "pricing with factor" workbook): one
sheet, `Country | Zip | Country+Zip | 0,1-100 kg | 100,1-200 kg | … | FTL`. It
covers 31 countries. Upload it in the sidebar under **Pallet rate card
(optional)** — it is parsed by `pallet_parser.py`. Pallet rows are added for
every selected country that has pallet data; countries without it stay
parcel-only.

### Bucket map

The source has 58 fine weight bands; the matrix collapses them into **21
operationally-meaningful pallet buckets** (a pallet/LDM arrangement may share a
source band with another). The collapse is encoded in `PALLET_BUCKET_MAP` in
`pipeline.py` as `(output MAX_WEIGHT, source band upper-bound)` pairs and is
applied to every country. `MIN_WEIGHT` is left blank so CargoWrite's
cheapest-first matching picks the smallest bucket that fits.

### Pricing chain (matches the combi reference exactly)

For each pallet row:
- `RATE_BASE = FACTORED RATE PALLET ÷ FACTOR` (Variables `B15`)
- `Mobility = mobility% × RATE_BASE` (Variables `B12`, default 4%)
- `BASE_TOTAL = RATE_BASE + Mobility`
- `FUEL = fuel% × BASE_TOTAL` (Variables `B11`, default 15% — DHL freight)
- `TOLL = toll% × BASE_TOTAL` (per-country, UK 0.43%)
- `ADMIN = flat € per shipment` (per-country, UK €46.51)
- `TOTAL = RATE_BASE + RATE_EXTRA + Mobility + FUEL + TOLL + ADMIN`

Fuel/mobility/factor are written as **formulas** referencing the Variables sheet
(tweak once, recalculates everywhere). Toll/admin are **per-country literal
values** so a single matrix can carry several countries with different pallet
surcharges. All four are editable in the sidebar (**Pallet surcharges**) and
configurable per country via `PALLET_COUNTRY_OVERRIDES`.

### New columns

The output gains: `FACTORED RATE PALLET`, `USER_DEF_TYPE_1` (PARCEL/PALLET),
`USER_DEF_TYPE_2` (Single/Multi — mirrors RATE_TYPE for parcels), `Mobility`,
`TOLL`, `ADMIN`. Parcel rows leave the pallet-only columns blank; pallet rows
leave the parcel-grid columns (MAX_PARCEL/EACH_WEIGHT/volumes/MAUT/Linehaul)
blank.

### Architecture notes

- `build_rows_pallet()` emits zone × bucket rows; `PALLET_DEFAULTS` /
  `pallet_carrier_cfg()` hold the economics.
- Pallets **bypass the parcel optimizer** — every (zone, bucket) is a distinct
  CargoWrite match target, and the parcel dominance logic assumes numeric
  postcodes/weights that pallets don't share. They are split out before
  optimization and re-attached before writing.
- `RATE_TYPE` (Single/Multi) was added for UPDE and UPSGB STANDARD rows; it is
  mirrored into `USER_DEF_TYPE_2`.
