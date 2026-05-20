"""
app.py — Rate Matrix Builder  ·  Streamlit UI
Run locally:   streamlit run app.py
Deploy:        push to GitHub, connect on share.streamlit.io
"""

import io
import logging
import tempfile
import zipfile
from copy import deepcopy
from pathlib import Path

import pandas as pd
import streamlit as st

import pipeline as pl

# ── Page config ───────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Rate Matrix Builder",
    page_icon="📦",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── Minimal custom CSS (clean, professional, no purple gradients) ─────────────
st.markdown("""
<style>
  /* Tight top padding */
  .block-container { padding-top: 1.5rem; }

  /* Section headers */
  .section-title {
    font-size: 0.75rem;
    font-weight: 700;
    letter-spacing: 0.12em;
    text-transform: uppercase;
    color: #888;
    margin: 1.4rem 0 0.5rem 0;
  }

  /* Country pill buttons */
  div[data-testid="column"] .stCheckbox label {
    font-size: 0.85rem;
  }

  /* Result table rows */
  .result-row {
    display: flex;
    align-items: center;
    gap: 1rem;
    padding: 0.5rem 0;
    border-bottom: 1px solid #e5e7eb;
    font-size: 0.9rem;
  }
</style>
""", unsafe_allow_html=True)

logging.basicConfig(level=logging.INFO)


# ── Helpers ───────────────────────────────────────────────────────────────────

ALL_COUNTRIES = sorted(pl.COUNTRY_CONFIG.keys())

CARRIER_LABELS = {cid: cfg['label'] for cid, cfg in pl.CARRIER_DEFAULTS.items()}

FUEL_CARRIERS = ['UPDE', 'DHL-ROS', 'DPD', 'UPSNL', 'POSTNORD']
MAUT_CARRIERS = ['DPD', 'DHL-ROS']


def _pct_input(label, key, default):
    """Small percentage input that returns a fraction (e.g. 0.27)."""
    val = st.number_input(
        label, min_value=0.0, max_value=1.0,
        value=default, step=0.01, format="%.2f", key=key,
    )
    return val


def build_carrier_defaults(fuel_vals, maut_vals):
    """Build a modified copy of CARRIER_DEFAULTS from user inputs."""
    cd = deepcopy(pl.CARRIER_DEFAULTS)
    for cid, pct in fuel_vals.items():
        if cid in cd:
            cd[cid]['fuel_pct'] = pct
    for cid, pct in maut_vals.items():
        if cid in cd:
            cd[cid]['maut_pct'] = pct
    return cd


def build_variables_layout(fuel_vals, maut_vals):
    """Build an updated Variables sheet layout from user inputs."""
    return [
        ('FUEL UPSDE',    fuel_vals.get('UPDE',     0.27)),
        ('FUEL DHL',      fuel_vals.get('DHL-ROS',  0.27)),
        ('FUEL DPD',      fuel_vals.get('DPD',      0.27)),
        ('FUEL UPSNL',    fuel_vals.get('UPSNL',    0.27)),
        ('FUEL POSTNORD', fuel_vals.get('POSTNORD', 0.27)),
        (None, None),
        ('MAUT DPD',      maut_vals.get('DPD',      0.05)),
        ('MAUT DHL',      maut_vals.get('DHL-ROS',  0.06)),
    ]


def file_as_bytes(path):
    return Path(path).read_bytes()


def make_zip(results):
    """Pack all minimal matrices into one ZIP; return bytes."""
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, 'w', zipfile.ZIP_DEFLATED) as zf:
        for country, r in results.items():
            for key, label in [('extended', 'extended'),
                                ('optimized', 'optimized'),
                                ('minimal', 'minimal')]:
                zf.write(r[key], f'{country}/{country}_Matrix_{label}.xlsx')
    buf.seek(0)
    return buf.read()


# ── Sidebar ───────────────────────────────────────────────────────────────────

with st.sidebar:
    st.markdown("## 📦 Rate Matrix Builder")
    st.caption("Fender Musical Instruments · Logistics")
    st.divider()

    # ── Upload ────────────────────────────────────────────────────────────────
    st.markdown('<p class="section-title">Rate card</p>', unsafe_allow_html=True)
    uploaded = st.file_uploader(
        "Upload Excel file",
        type=["xlsx", "xls"],
        help="The rate-card workbook with tabs for each carrier (UPSDE, DHL, DPD, UPSNL, POSTNORD).",
        label_visibility="collapsed",
    )

    # ── Surcharge rates ───────────────────────────────────────────────────────
    st.markdown('<p class="section-title">Fuel surcharge (%)</p>', unsafe_allow_html=True)
    fuel_defaults = {cid: pl.CARRIER_DEFAULTS[cid]['fuel_pct'] for cid in FUEL_CARRIERS}
    fuel_vals = {}
    cols = st.columns(2)
    for i, cid in enumerate(FUEL_CARRIERS):
        with cols[i % 2]:
            fuel_vals[cid] = _pct_input(
                CARRIER_LABELS[cid], f'fuel_{cid}', fuel_defaults[cid]
            )

    st.markdown('<p class="section-title">MAUT surcharge (%)</p>', unsafe_allow_html=True)
    maut_defaults = {cid: pl.CARRIER_DEFAULTS[cid]['maut_pct'] for cid in MAUT_CARRIERS}
    maut_vals = {}
    cols = st.columns(2)
    for i, cid in enumerate(MAUT_CARRIERS):
        with cols[i % 2]:
            maut_vals[cid] = _pct_input(
                CARRIER_LABELS[cid], f'maut_{cid}', maut_defaults[cid]
            )

    st.divider()

    # ── Run button ────────────────────────────────────────────────────────────
    run_btn = st.button(
        "▶ Generate matrices",
        type="primary",
        use_container_width=True,
        disabled=(uploaded is None),
    )
    if uploaded is None:
        st.caption("Upload a rate card first.")


# ── Main area ─────────────────────────────────────────────────────────────────

st.markdown("## Select countries")
st.caption("Choose which countries to generate matrices for.")

# Country grid — 10 columns
COLS_PER_ROW = 10
selected_countries = []

# Initialise selection state
if 'country_selection' not in st.session_state:
    st.session_state.country_selection = {c: False for c in ALL_COUNTRIES}

# Select all / none helpers
col_a, col_b, *_ = st.columns([1, 1, 8])
if col_a.button("Select all"):
    for c in ALL_COUNTRIES:
        st.session_state.country_selection[c] = True
if col_b.button("Clear"):
    for c in ALL_COUNTRIES:
        st.session_state.country_selection[c] = False

grid_cols = st.columns(COLS_PER_ROW)
for i, country in enumerate(ALL_COUNTRIES):
    with grid_cols[i % COLS_PER_ROW]:
        st.session_state.country_selection[country] = st.checkbox(
            country,
            value=st.session_state.country_selection[country],
            key=f'chk_{country}',
        )

selected_countries = [c for c in ALL_COUNTRIES
                      if st.session_state.country_selection.get(c)]

# ── Advanced: per-country carrier selection ───────────────────────────────────
if selected_countries:
    with st.expander("⚙️ Advanced — carrier & postcode settings per country",
                     expanded=False):
        st.caption(
            "Defaults are pre-configured. Only change if the rate card for a "
            "specific country differs."
        )
        if 'country_overrides' not in st.session_state:
            st.session_state.country_overrides = {}

        for country in selected_countries:
            base = pl.COUNTRY_CONFIG[country]
            ov   = st.session_state.country_overrides.setdefault(country, {
                'carriers':              list(base['carriers']),
                'max_parcel_count':      base['max_parcel_count'],
                'max_each_weight_kg':    base['max_each_weight_kg'],
                'postcode_prefix_range': list(base['postcode_prefix_range']),
            })

            st.markdown(f"**{country}**")
            c1, c2, c3, c4, c5 = st.columns([3, 1, 1, 1, 1])

            all_carriers = list(pl.CARRIER_DEFAULTS.keys())
            with c1:
                ov['carriers'] = st.multiselect(
                    "Carriers", all_carriers,
                    default=ov['carriers'],
                    key=f'carriers_{country}',
                )
            with c2:
                ov['max_parcel_count'] = st.number_input(
                    "Max parcels", 1, 20,
                    value=ov['max_parcel_count'], key=f'mp_{country}',
                )
            with c3:
                ov['max_each_weight_kg'] = st.number_input(
                    "Max weight (kg)", 1.0, 70.0,
                    value=float(ov['max_each_weight_kg']),
                    step=0.5, key=f'mw_{country}',
                )
            with c4:
                ov['postcode_prefix_range'][0] = st.number_input(
                    "PC from", 0, 99,
                    value=ov['postcode_prefix_range'][0],
                    key=f'pc0_{country}',
                )
            with c5:
                ov['postcode_prefix_range'][1] = st.number_input(
                    "PC to", 0, 99,
                    value=ov['postcode_prefix_range'][1],
                    key=f'pc1_{country}',
                )
            st.session_state.country_overrides[country] = ov

st.divider()

# ── Results area ──────────────────────────────────────────────────────────────

if 'results' not in st.session_state:
    st.session_state.results = {}

if run_btn and uploaded and selected_countries:
    st.session_state.results = {}
    carrier_defaults  = build_carrier_defaults(fuel_vals, maut_vals)
    variables_layout  = build_variables_layout(fuel_vals, maut_vals)

    # Save uploaded file to a temp location once
    with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as tmp_in:
        tmp_in.write(uploaded.read())
        input_path = tmp_in.name

    progress = st.progress(0, text="Starting…")
    errors   = {}

    for idx, country in enumerate(selected_countries):
        progress.progress(
            idx / len(selected_countries),
            text=f"Processing {country}…  ({idx+1}/{len(selected_countries)})",
        )
        try:
            # Build country config with any advanced overrides
            base_cfg = deepcopy(pl.COUNTRY_CONFIG[country])
            ov = st.session_state.get('country_overrides', {}).get(country, {})
            if ov:
                base_cfg.update({
                    'carriers':              ov['carriers'],
                    'max_parcel_count':      ov['max_parcel_count'],
                    'max_each_weight_kg':    ov['max_each_weight_kg'],
                    'postcode_prefix_range': tuple(ov['postcode_prefix_range']),
                    'each_weight_grid':      sorted(
                        set(list(range(1, int(ov['max_each_weight_kg']) + 1))
                            + [ov['max_each_weight_kg']])
                    ),
                })

            with tempfile.TemporaryDirectory() as tmp_out:
                result = pl.run_pipeline(
                    input_path, country, output_dir=tmp_out,
                    country_cfg=base_cfg,
                    carrier_defaults=carrier_defaults,
                    variables_layout=variables_layout,
                )
                # Copy files out of the temp dir before it is deleted
                persistent = tempfile.mkdtemp()
                import shutil
                for key in ('extended', 'optimized', 'minimal'):
                    src = result[key]
                    dst = Path(persistent) / Path(src).name
                    shutil.copy(src, dst)
                    result[key] = str(dst)

            st.session_state.results[country] = result

        except Exception as e:
            errors[country] = str(e)

    progress.progress(1.0, text="Done.")

    if errors:
        for country, msg in errors.items():
            st.error(f"**{country}**: {msg}")

elif run_btn and not selected_countries:
    st.warning("Please select at least one country.")

# ── Show results ──────────────────────────────────────────────────────────────

if st.session_state.results:
    st.markdown("## Results")

    results = st.session_state.results

    # Summary table
    summary = []
    for country, r in results.items():
        summary.append({
            'Country':   country,
            'Extended':  f"{r['rows_extended']:,}",
            'Optimized': f"{r['rows_optimized']:,}",
            'Minimal':   f"{r['rows_minimal']:,}",
        })
    st.dataframe(
        pd.DataFrame(summary).set_index('Country'),
        use_container_width=True,
    )

    st.caption("**Extended** = all combinations · **Optimized** = per-carrier dominance removed · **Minimal** = cross-carrier dominance removed")

    # Per-country download buttons
    st.markdown("#### Download individual countries")
    for country, r in results.items():
        c1, c2, c3, c4 = st.columns([1, 2, 2, 2])
        c1.markdown(f"**{country}**")
        for col, key, label in [
            (c2, 'extended',  '📥 Extended'),
            (c3, 'optimized', '📥 Optimized'),
            (c4, 'minimal',   '📥 Minimal'),
        ]:
            col.download_button(
                label,
                data=file_as_bytes(r[key]),
                file_name=Path(r[key]).name,
                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                key=f'dl_{country}_{key}',
            )

    # Download all as ZIP
    st.markdown("#### Download everything")
    st.download_button(
        "📦 Download all countries as ZIP",
        data=make_zip(results),
        file_name="rate_matrices.zip",
        mime="application/zip",
        type="primary",
        use_container_width=False,
    )
