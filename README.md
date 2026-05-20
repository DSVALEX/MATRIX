# Rate Matrix Builder

A Streamlit app that turns a rate-card Excel into carrier pricing matrices.

## Files

| File | Purpose |
|------|---------|
| `app.py` | Streamlit UI — the only file managers interact with |
| `pipeline.py` | All logic: parsing, building, optimizing, writing Excel |
| `requirements.txt` | Python dependencies |

## Deploy to Streamlit Cloud (free, 5 minutes)

1. **Create a GitHub repository** (can be private)
2. **Upload these four files** to the root of the repo:
   - `app.py`
   - `pipeline.py`
   - `requirements.txt`
   - `README.md`
3. Go to **[share.streamlit.io](https://share.streamlit.io)** and sign in with GitHub
4. Click **New app** → select your repo → main branch → `app.py` → **Deploy**
5. Share the URL with your team

## Using the app

1. **Upload** your rate-card Excel (tabs: UPSDE, DHL, DPD, UPSNL, POSTNORD)
2. **Adjust fuel & MAUT %** in the sidebar if rates have changed
3. **Select countries** using the checkboxes
4. Optionally expand **Advanced settings** to change carriers or postcode ranges per country
5. Click **▶ Generate matrices**
6. **Download** individual countries or everything as a ZIP

Each country produces three files:
- `*_Matrix_extended.xlsx` — every possible combination
- `*_Matrix_optimized.xlsx` — dominated rows removed per carrier/service
- `*_Matrix_minimal.xlsx` — dominated rows removed across all carriers (smallest file, use this one)

## Rate card Excel format

The workbook needs one sheet per carrier. Sheet names are matched fuzzily
(e.g. "UPS DE", "UPS-DE", and "UPSDE" all work):

| Carrier | Expected sheet name |
|---------|-------------------|
| UPS DE contract | `UPSDE` |
| UPS NL contract | `UPSNL` |
| DHL / ROS | `DHL` |
| DPD | `DPD` |
| PostNord | `POSTNORD` |

## Changing defaults

All defaults (carriers per country, site/client IDs, weight grids) live in
`pipeline.py` under `CARRIER_DEFAULTS` and `COUNTRY_CONFIG`. Edit those
constants and redeploy.
