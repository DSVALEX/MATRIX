# Rate Matrix Builder

A small website (a Streamlit app) that turns a carrier **rate card** — a normal
Excel file full of shipping prices — into clean **pricing matrices** that our
transport system, **CargoWrite**, can read.

You upload one Excel file, click a button, and download ready-to-use price
tables. That's the whole job. You do **not** need to write code or understand
the formulas to use it.

---

# Part 1 — For first-time users

If you have never opened this tool before, read this part and nothing else.
Everything below "Part 2" is for the person who maintains the code.

## What problem does this solve?

Carriers (UPS, DHL, DPD, PostNord…) send us their prices as messy spreadsheets,
each laid out differently. CargoWrite needs them in **one consistent format**,
with fuel and other surcharges already added in, and with the cheapest valid
option listed first. Doing that by hand is slow and error-prone. This tool does
it in a few seconds and produces the exact layout CargoWrite expects.

## Words you'll see (plain-language glossary)

| Term | What it means here |
|------|--------------------|
| **Rate card** | The Excel price file a carrier gives us. The thing you upload. |
| **Matrix** | The finished price table this tool produces. One row = one priced shipping option. |
| **Carrier** | The shipping company: UPS DE, UPS NL, UPS GB, DHL, DPD, PostNord, DHL-FENDER (pallets). |
| **Parcel** | A normal box shipment, priced by weight and number of parcels. |
| **Pallet** | A freight shipment on a wooden pallet, priced by destination zone and weight band. Different carrier (DHL-FENDER), different pricing. |
| **Surcharge** | An extra cost added on top of the base rate — most importantly **fuel %**, plus toll, MAUT, admin, mobility. |
| **MAUT** | A road-toll surcharge that some countries (mostly in central Europe) add. Read automatically from the master file; you don't type it in. |
| **CargoWrite** | Our transport management system. It reads the matrix you download and uses it to price real orders. |
| **Bucket / catch-all row** | A safety-net row that catches unusual orders (oversized, too heavy, or a postcode the carrier didn't list) so nothing falls through with no price. Shown highlighted for review. |
| **Variables sheet** | A second tab inside the downloaded file listing the surcharge percentages, so they can be checked or tweaked in Excel later. |

## Before you start

You need two things:

1. **The website link** — the person who set this up will give you a
   `…streamlit.app` URL. Open it in any browser. There's nothing to install.
2. **A rate card to upload.** The usual one is the DSV master workbook
   (file name like `MDK__FENDER__PARCEL_RATES__S2026.xlsx`). For pallets you
   also have a second file (the DHL "pricing with factor" workbook).

## Using the website — step by step

1. **Open the link.** You'll see a sidebar on the left and a main panel.
2. **Upload your rate card.** Use the upload box at the top of the sidebar. The
   tool figures out the format on its own — you don't pick a type.
3. *(Pallets only)* **Upload the pallet rate card** in the separate
   "Pallet rate card (optional)" box just below. Skip this if you only need
   parcels.
4. **Check the fuel %.** Fuel changes monthly, so confirm the box in the sidebar
   shows the current rate. Most other surcharges are read from the file for you.
5. **Pick the countries** you want matrices for. Only countries actually present
   in your file appear in the list.
6. *(Optional)* Open **Advanced** or **📐 Exceptions & buckets** only if someone
   has told you to change carriers, postcode ranges, or oversized-parcel rules.
   You can safely ignore these the first time.
7. **Click ▶ Generate matrices.** Wait a few seconds.
8. **Download** — either one country at a time, **all of them as a single
   combined file**, or everything zipped together.

That's it. If you only ever do steps 1, 4, 5, 7, 8, you're using it correctly.

## Understanding what you downloaded

**Each country produces three files.** They are the *same* prices filtered three
different ways:

| File | What it is | Use it? |
|------|-----------|---------|
| `*_extended` | Every possible row, nothing removed. | For auditing only. |
| `*_optimized` | Cheaper-or-equal duplicates removed *within* each carrier. | Intermediate. |
| `*_minimal` | Cheaper options removed *across all carriers* — only the genuinely best rows survive. | **This is the one to give CargoWrite.** |

There's also a **Combined** download that merges every selected country's
`minimal` table into one sheet, sorted by country then price — handy when you
want a single file instead of many.

### Row colours in the Excel

Some rows are highlighted so you can spot them at a glance:

- **Blue rows** = **pallet** shipments (DHL-FENDER freight). They're priced on
  a completely different basis from parcels, so the colour keeps the two modes
  visually separate in a mixed sheet.
- **Pale amber rows** = **catch-all / bucket** rows — the safety-net rows for
  oversized, overweight, or unlisted-postcode orders. They carry an extra
  surcharge, so ops should glance over them.

Everything else is a normal parcel row.

### The "Variables" tab

Most downloads include a second tab called **Variables** listing the surcharge
percentages (fuel, MAUT, mobility, toll, admin). In the per-country files these
feed live Excel formulas, so changing a number there recalculates the whole
sheet. In the **combined** and **pallet** files the prices are written as fixed
numbers (because surcharges differ by country and can't all share one cell), and
the Variables tab is there for reference only.

## Quick troubleshooting

- **A country I expected isn't in the list.** It isn't present in the uploaded
  file, or it has no data for the carriers you selected.
- **Pallet rows didn't appear.** You didn't upload the separate pallet rate
  card, or that country has no pallet data (those stay parcel-only).
- **Prices look off after I changed fuel.** Re-generate — the sidebar value is
  only applied when you click Generate.
- **It looks stale after an update.** Re-upload the file; the app keys its cache
  on the file's name *and* size.

---

# Part 2 — For whoever maintains the tool

Everything below is developer reference. Day-to-day users can stop here.

## Files

| File | Purpose |
|------|---------|
| `app.py` | Streamlit UI — the only file users interact with |
| `pipeline.py` | Core parcel logic: build, compute, optimize, write Excel |
| `pallet_pipeline.py` | Pallet (DHL freight / EUROCONNECT) matrix builder |
| `master_parser.py` | Parses the DSV master rate card (all countries in one file) |
| `pallet_parser.py` | Parses the DHL "pricing with factor" pallet workbook |
| `requirements.txt` | Python dependencies |

## Two input formats (auto-detected)

**1. Master file** — the DSV "MDK - FENDER - PARCEL RATES" workbook with one
sheet per rate table and all countries in one file. Upload this and the app:
- reads every carrier and country in one pass
- shows only the countries actually present in the file
- reads **MAUT surcharges per country, per carrier** directly from the file
  (the MAUT inputs disappear from the sidebar)
- you still set **fuel %** in the sidebar (fuel is a separate monthly surcharge)

**2. Per-country file** — the older one-country-per-workbook format (tabs named
UPSDE, DHL, DPD, UPSNL, POSTNORD). Still fully supported.

## Carriers supported

UPS DE, UPS NL, UPS GB (UK domestic), DHL, DPD, PostNord, and DHL-FENDER
(pallets).

Special handling baked in:
- **UPS WorldEase (WEA)** — flat per-parcel rate for CH and NO
- **DHL BNL** — "1st parcel + each additional" pricing for BE, LU, NL
- **PostNord** — flat rate per service (B2B / Home / PUDO), multi-country table
- **UPS DE zones** — postcode/zone-based, including alphanumeric zones (ES4/ES5/ES6)

## Deploy to Streamlit Cloud (free)

1. Create a GitHub repo (can be private).
2. Upload all files to the root: `app.py`, `pipeline.py`, `pallet_pipeline.py`,
   `master_parser.py`, `pallet_parser.py`, `requirements.txt`, `README.md`.
3. Go to [share.streamlit.io](https://share.streamlit.io), sign in with GitHub.
4. New app → pick repo → branch → main file `app.py` → Deploy.
5. Share the URL.

## Changing defaults

Carrier defaults and country configs live in `pipeline.py`
(`CARRIER_DEFAULTS`, `COUNTRY_CONFIG`). The master-file sheet names and parsing
logic live in `master_parser.py`. Pallet economics live in `PALLET_DEFAULTS`,
`PALLET_COUNTRY_OVERRIDES`, and `PALLET_MAUT` in `pipeline.py`. Edit and redeploy.

## Exceptions & buckets (add-on)

CargoWrite picks the first matrix row (cheapest-first) whose constraints all fit
an order. A constraint like a **size limit** means a row only matches parcels
at/under that size — so oversized orders need a **bucket** row lower down that
drops the limit, adds a surcharge, and is flagged for review. Without buckets,
some orders match nothing.

In the app, open **📐 Exceptions & buckets** and edit the table — one line per
carrier (or country). Each line says: *for this carrier, the normal size limit
is X metres; anything bigger gets a €Y-per-parcel surcharge.* On Generate, every
output gets:
- the **normal limit** stamped on the cheap base rows
- a **catch-all bucket twin** (limit blank, surcharge added), highlighted in
  **amber** for oversight

> **Note (changed):** the old standalone `AWKWARD` flag column has been removed
> from the output. Bucket / catch-all rows are now identified by their **amber
> highlight** rather than a `Y` in a column. The rule engine still supports a
> `flag_col` (it defaults to the now-unexported `AWKWARD` name); point it at a
> real, exported column such as `USER_DEF_TYPE_2` if you want a machine-readable
> marker back.

The mechanism is general: in `pipeline.py`, `apply_exceptions()` takes a list of
rule dicts (constraint column, normal value, surcharge, scope). Today the UI
exposes the size case; the same function already supports flat surcharges and
other constraint columns.

Leave the table empty for a plain matrix with no buckets.

### Future buckets (implemented, off by default)

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

The builder handles **pallets** as a second shipping mode alongside parcels, in
the same output matrix. Pallets use carrier **DHL-FENDER** (service EUROCONNECT)
and are priced on two axes — destination **postcode zone** and a **weight
bucket** — instead of the parcel each-weight/parcel-count grid. Pallet rows are
written in **orange** so they stand out from parcel rows in a mixed sheet.

### Pallet rate card (separate upload)

Pallets come from their own file (the DHL "pricing with factor" workbook): one
sheet, `Country | Zip | Country+Zip | 0,1-100 kg | 100,1-200 kg | … | FTL`. It
covers 31 countries. Upload it in the sidebar under **Pallet rate card
(optional)** — parsed by `pallet_parser.py`. Pallet rows are added for every
selected country that has pallet data; countries without it stay parcel-only.

### Bucket map

The source has 58 fine weight bands; the matrix collapses them into **21
operationally-meaningful pallet buckets**. The collapse is encoded in
`PALLET_BUCKET_MAP` in `pipeline.py` as `(output MAX_WEIGHT, source band
upper-bound)` pairs and is applied to every country. `MIN_WEIGHT` is left blank
so CargoWrite's cheapest-first matching picks the smallest bucket that fits.

### Pricing chain (matches the combi reference exactly)

For each pallet row:
- `RATE_BASE = FACTORED RATE PALLET ÷ FACTOR`
- `Mobility = mobility% × RATE_BASE` (default 4%)
- `BASE_TOTAL = RATE_BASE + Mobility`
- `FUEL = fuel% × BASE_TOTAL` (default 15% — DHL freight)
- `TOLL = toll% × BASE_TOTAL` (per-country, UK 0.43%)
- `ADMIN = flat € per shipment` (per-country, UK €46.51)
- `TOTAL = RATE_BASE + RATE_EXTRA + Mobility + FUEL + TOLL + ADMIN`

Fuel/mobility/factor are written as **formulas** referencing the Variables sheet
in per-country files (tweak once, recalculates everywhere). Toll/admin are
**per-country literal values** so a single matrix can carry several countries
with different pallet surcharges. All four are editable in the sidebar
(**Pallet surcharges**) and configurable per country via
`PALLET_COUNTRY_OVERRIDES`.

### Columns

Pallet output carries: `FACTORED RATE PALLET`, `USER_DEF_TYPE_1` (PARCEL/PALLET),
`USER_DEF_TYPE_2` (Single/Multi — mirrors RATE_TYPE for parcels), `Mobility`,
`TOLL`, `ADMIN`. Parcel rows leave the pallet-only columns blank; pallet rows
leave the parcel-grid columns (MAX_PARCEL/EACH_WEIGHT/volumes/MAUT/Linehaul)
blank.

### Architecture notes

- `build_pallet_df()` emits zone × bucket rows; `PALLET_DEFAULTS` hold the
  economics.
- Pallets **bypass the parcel optimizer** — every (zone, bucket) is a distinct
  CargoWrite match target, and the parcel dominance logic assumes numeric
  postcodes/weights that pallets don't share. They are split out before
  optimization and re-attached before writing.
- `RATE_TYPE` (Single/Multi) was added for UPDE and UPSGB STANDARD rows; it is
  mirrored into `USER_DEF_TYPE_2`.

## Combined export (all countries in one sheet)

Alongside the per-country files and the ZIP, the app offers a **combined**
download: every selected country's minimal matrix merged into a single sheet
(`Combined_Matrix_minimal.xlsx`), sorted by country then price.

Key design point: the combined sheet writes computed columns as **numeric
values, not formulas**. A single sheet can't reference one Variables cell for a
surcharge that varies by country (MAUT differs per country; pallet toll/admin
are UK-only), so formulas would silently apply one country's rate to all.
Numeric values keep every country correct in one file. A Variables sheet is
still included for reference. The engine function is
`pipeline.write_combined_matrix(frames, path, variables_layout)`; the app builds
it from each country's persisted minimal frame (`result['minimal_df']`).
