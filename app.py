"""
app.py — Rate Matrix Builder  ·  Streamlit UI

Supports two input formats, auto-detected on upload:
  • MASTER file  — the DSV "MDK - FENDER - PARCEL RATES" workbook (all countries
    in one file). MAUT is read per-country from the file automatically.
  • Per-country file — the older one-country-per-workbook format.

Run locally:   streamlit run app.py
Deploy:        push to GitHub, connect on share.streamlit.io
"""

import io
import logging
import shutil
import tempfile
import zipfile
from copy import deepcopy
from pathlib import Path

import pandas as pd
import streamlit as st

import pipeline as pl
import master_parser as mp

st.set_page_config(page_title="Rate Matrix Builder", page_icon="📦",
                   layout="wide", initial_sidebar_state="expanded")

st.markdown("""
<style>
  .block-container { padding-top: 1.5rem; }
  .section-title { font-size:.75rem; font-weight:700; letter-spacing:.12em;
                   text-transform:uppercase; color:#888; margin:1.4rem 0 .5rem 0; }
  div[data-testid="column"] .stCheckbox label { font-size:.85rem; }
</style>
""", unsafe_allow_html=True)

logging.basicConfig(level=logging.INFO)

ALL_COUNTRIES  = sorted(pl.COUNTRY_CONFIG.keys())
CARRIER_LABELS = {cid: cfg['label'] for cid, cfg in pl.CARRIER_DEFAULTS.items()}
FUEL_CARRIERS  = ['UPDE', 'DHL-ROS', 'DPD', 'UPSNL', 'POSTNORD', 'UPSGB']
MAUT_CARRIERS  = ['DPD', 'DHL-ROS']


# ── Helpers ───────────────────────────────────────────────────────────────────

def _pct_input(label, key, default):
    return st.number_input(label, min_value=0.0, max_value=1.0,
                           value=float(default), step=0.01, format="%.2f", key=key)


def variables_layout(fuel_vals, maut_dhl, maut_dpd):
    return [
        ('FUEL UPSDE',    fuel_vals.get('UPDE',     0.27)),
        ('FUEL DHL',      fuel_vals.get('DHL-ROS',  0.27)),
        ('FUEL DPD',      fuel_vals.get('DPD',      0.27)),
        ('FUEL UPSNL',    fuel_vals.get('UPSNL',    0.27)),
        ('FUEL POSTNORD', fuel_vals.get('POSTNORD', 0.27)),
        ('FUEL UPSGB',    fuel_vals.get('UPSGB',    0.27)),
        (None, None),
        ('MAUT DPD',      maut_dpd),
        ('MAUT DHL',      maut_dhl),
    ]


def carrier_defaults(fuel_vals, maut_dhl, maut_dpd):
    cd = deepcopy(pl.CARRIER_DEFAULTS)
    for cid, pct in fuel_vals.items():
        if cid in cd:
            cd[cid]['fuel_pct'] = pct
    cd['DHL-ROS']['maut_pct'] = maut_dhl
    cd['DPD']['maut_pct']     = maut_dpd
    return cd


def country_cfg_with_overrides(country):
    cfg = deepcopy(pl.COUNTRY_CONFIG[country])
    ov  = st.session_state.get('country_overrides', {}).get(country)
    if ov:
        cfg.update({
            'carriers':              ov['carriers'],
            'max_parcel_count':      ov['max_parcel_count'],
            'max_each_weight_kg':    ov['max_each_weight_kg'],
            'postcode_prefix_range': tuple(ov['postcode_prefix_range']),
            'each_weight_grid':      sorted(set(
                list(range(1, int(ov['max_each_weight_kg']) + 1))
                + [ov['max_each_weight_kg']])),
        })
    return cfg


def persist(result):
    """Copy result files out of a temp dir into a longer-lived temp dir."""
    out = tempfile.mkdtemp()
    for key in ('extended', 'optimized', 'minimal'):
        dst = Path(out) / Path(result[key]).name
        shutil.copy(result[key], dst)
        result[key] = str(dst)
    return result


def file_bytes(p):
    return Path(p).read_bytes()


def make_zip(results):
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

    st.markdown('<p class="section-title">Rate card</p>', unsafe_allow_html=True)
    uploaded = st.file_uploader("Upload Excel file", type=["xlsx", "xls"],
                                label_visibility="collapsed",
                                help="Upload the DSV master rate card, or an older "
                                     "per-country file. The format is detected automatically.")

    # Detect master vs per-country (cache parse in session)
    is_master = False
    master = None
    if uploaded is not None:
        # Save bytes once; reuse across reruns
        if st.session_state.get('uploaded_name') != uploaded.name:
            with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as tmp:
                tmp.write(uploaded.read())
                st.session_state['input_path']  = tmp.name
                st.session_state['uploaded_name'] = uploaded.name
                st.session_state.pop('master', None)
        input_path = st.session_state['input_path']

        is_master = mp.is_master_file(input_path)
        if is_master:
            if 'master' not in st.session_state:
                with st.spinner("Reading master rate card…"):
                    st.session_state['master'] = mp.parse_master_rate_card(input_path)
            master = st.session_state['master']
            avail = mp.available_countries(master)
            st.success(f"Master rate card detected — {len(avail)} countries available. "
                       "MAUT is read per-country from the file.")
        else:
            st.info("Per-country rate card detected.")

    st.markdown('<p class="section-title">Fuel surcharge (%)</p>', unsafe_allow_html=True)
    fuel_vals = {}
    cols = st.columns(2)
    for i, cid in enumerate(FUEL_CARRIERS):
        with cols[i % 2]:
            fuel_vals[cid] = _pct_input(CARRIER_LABELS[cid], f'fuel_{cid}',
                                        pl.CARRIER_DEFAULTS[cid]['fuel_pct'])

    # MAUT inputs only matter for per-country files (master reads MAUT itself)
    maut_dhl = pl.CARRIER_DEFAULTS['DHL-ROS']['maut_pct']
    maut_dpd = pl.CARRIER_DEFAULTS['DPD']['maut_pct']
    if not is_master:
        st.markdown('<p class="section-title">MAUT surcharge (%)</p>', unsafe_allow_html=True)
        cols = st.columns(2)
        with cols[0]:
            maut_dpd = _pct_input('DPD', 'maut_DPD', maut_dpd)
        with cols[1]:
            maut_dhl = _pct_input('DHL', 'maut_DHL', maut_dhl)

    st.divider()
    run_btn = st.button("▶ Generate matrices", type="primary",
                        use_container_width=True, disabled=(uploaded is None))
    if uploaded is None:
        st.caption("Upload a rate card first.")


# ── Country selection ─────────────────────────────────────────────────────────

st.markdown("## Select countries")
if is_master and master is not None:
    avail = mp.available_countries(master)
    selectable = [c for c in ALL_COUNTRIES if c in avail]
    if 'GB' in avail and 'GB' not in selectable:
        selectable.append('GB')
    selectable = sorted(set(selectable))
    st.caption(f"Showing the {len(selectable)} countries present in this rate card.")
else:
    selectable = ALL_COUNTRIES
    st.caption("Choose which countries to generate matrices for.")

if 'country_selection' not in st.session_state:
    st.session_state.country_selection = {}
for c in selectable:
    st.session_state.country_selection.setdefault(c, False)

ca, cb, *_ = st.columns([1, 1, 8])
if ca.button("Select all"):
    for c in selectable:
        st.session_state.country_selection[c] = True
if cb.button("Clear"):
    for c in selectable:
        st.session_state.country_selection[c] = False

COLS = 10
grid = st.columns(COLS)
for i, country in enumerate(selectable):
    with grid[i % COLS]:
        st.session_state.country_selection[country] = st.checkbox(
            country, value=st.session_state.country_selection.get(country, False),
            key=f'chk_{country}')

selected = [c for c in selectable if st.session_state.country_selection.get(c)]

# ── Advanced per-country settings ─────────────────────────────────────────────
if selected:
    with st.expander("⚙️ Advanced — carrier & postcode settings per country", expanded=False):
        st.caption("Defaults are pre-configured. Change only if a country differs.")
        if 'country_overrides' not in st.session_state:
            st.session_state.country_overrides = {}
        for country in selected:
            base = pl.COUNTRY_CONFIG.get(country, pl._default_country_cfg(country))
            ov = st.session_state.country_overrides.setdefault(country, {
                'carriers':              list(base['carriers']),
                'max_parcel_count':      base['max_parcel_count'],
                'max_each_weight_kg':    base['max_each_weight_kg'],
                'postcode_prefix_range': list(base['postcode_prefix_range']),
            })
            st.markdown(f"**{country}**")
            c1, c2, c3, c4, c5 = st.columns([3, 1, 1, 1, 1])
            with c1:
                ov['carriers'] = st.multiselect("Carriers", list(pl.CARRIER_DEFAULTS),
                                                default=ov['carriers'], key=f'car_{country}')
            with c2:
                ov['max_parcel_count'] = st.number_input("Max parcels", 1, 20,
                                                         value=ov['max_parcel_count'], key=f'mp_{country}')
            with c3:
                ov['max_each_weight_kg'] = st.number_input("Max kg", 1.0, 70.0,
                                                           value=float(ov['max_each_weight_kg']),
                                                           step=0.5, key=f'mw_{country}')
            with c4:
                ov['postcode_prefix_range'][0] = st.number_input("PC from", 0, 99,
                                                                 value=ov['postcode_prefix_range'][0], key=f'p0_{country}')
            with c5:
                ov['postcode_prefix_range'][1] = st.number_input("PC to", 0, 99,
                                                                 value=ov['postcode_prefix_range'][1], key=f'p1_{country}')
            st.session_state.country_overrides[country] = ov

st.divider()

# ── Run ────────────────────────────────────────────────────────────────────────
if 'results' not in st.session_state:
    st.session_state.results = {}

if run_btn and uploaded and selected:
    st.session_state.results = {}
    input_path = st.session_state['input_path']
    errors = {}
    progress = st.progress(0, text="Starting…")

    for idx, country in enumerate(selected):
        progress.progress(idx / len(selected),
                          text=f"Processing {country}…  ({idx+1}/{len(selected)})")
        try:
            cfg = country_cfg_with_overrides(country)

            if is_master:
                parsed = mp.country_rate_data(master, country)
                maut   = mp.country_maut(master, country)
                cd = carrier_defaults(fuel_vals, maut['DHL-ROS'], maut['DPD'])
                vl = variables_layout(fuel_vals, maut['DHL-ROS'], maut['DPD'])
                with tempfile.TemporaryDirectory() as tmp:
                    result = pl.run_pipeline_from_parsed(parsed, country, tmp, cfg, cd, vl)
                    result = persist(result)
            else:
                cd = carrier_defaults(fuel_vals, maut_dhl, maut_dpd)
                vl = variables_layout(fuel_vals, maut_dhl, maut_dpd)
                with tempfile.TemporaryDirectory() as tmp:
                    result = pl.run_pipeline(input_path, country, output_dir=tmp,
                                             country_cfg=cfg, carrier_defaults=cd,
                                             variables_layout=vl)
                    result = persist(result)

            st.session_state.results[country] = result
        except Exception as e:
            errors[country] = str(e)

    progress.progress(1.0, text="Done.")
    for country, msg in errors.items():
        st.error(f"**{country}**: {msg}")

elif run_btn and not selected:
    st.warning("Please select at least one country.")

# ── Results ──────────────────────────────────────────────────────────────────
if st.session_state.results:
    st.markdown("## Results")
    results = st.session_state.results

    summary = [{'Country': c, 'Extended': f"{r['rows_extended']:,}",
                'Optimized': f"{r['rows_optimized']:,}", 'Minimal': f"{r['rows_minimal']:,}"}
               for c, r in results.items()]
    st.dataframe(pd.DataFrame(summary).set_index('Country'), use_container_width=True)
    st.caption("**Extended** = all combinations · **Optimized** = per-carrier dominance "
               "removed · **Minimal** = cross-carrier dominance removed (use this one)")

    st.markdown("#### Download individual countries")
    for country, r in results.items():
        c1, c2, c3, c4 = st.columns([1, 2, 2, 2])
        c1.markdown(f"**{country}**")
        for col, key, label in [(c2, 'extended', '📥 Extended'),
                                (c3, 'optimized', '📥 Optimized'),
                                (c4, 'minimal', '📥 Minimal')]:
            col.download_button(label, data=file_bytes(r[key]),
                                file_name=Path(r[key]).name,
                                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                                key=f'dl_{country}_{key}')

    st.markdown("#### Download everything")
    st.download_button("📦 Download all countries as ZIP", data=make_zip(results),
                       file_name="rate_matrices.zip", mime="application/zip", type="primary")
