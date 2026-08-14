"""
engine.py — MMU Hub Report computation + workbook-building engine.
Implements Build Spec v8, generalized to accept uploaded files (any subset
may be missing) instead of fixed local paths, with graceful per-parameter
degradation instead of an all-or-nothing failure.

Only Raw Data carries live formulas (with cached values written alongside,
via xlsxwriter's write_formula(..., value=...), so the numbers show up
immediately without needing a LibreOffice/Excel recalculation pass — this
keeps the app self-contained for Streamlit Cloud deployment). Every other
visible sheet is hardcoded pre-computed values, exactly as in the desktop
build.
"""
import io
import os
import numpy as np
import pandas as pd
import xlsxwriter  # noqa: F401 (re-exported for convenience elsewhere)

SUB_DIVISIONS = ['ASP Karad West', 'SDIP Karad East', 'SDIP Vaduj']

DATA_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'data')
MASTER_PATH = os.path.join(DATA_DIR, 'master_file.xlsx')
LOGO_PATH = os.path.join(DATA_DIR, 'logo.png')


def ensure_data_dir():
    os.makedirs(DATA_DIR, exist_ok=True)


def save_master_file(uploaded_file):
    """Persist an uploaded Master file to local disk, overwriting any previous one."""
    ensure_data_dir()
    with open(MASTER_PATH, 'wb') as fh:
        fh.write(uploaded_file.getbuffer())


def load_persisted_master():
    """Return the persisted Master file path if one exists, else None."""
    return MASTER_PATH if os.path.exists(MASTER_PATH) else None


def save_logo_file(uploaded_file):
    ensure_data_dir()
    ext = os.path.splitext(uploaded_file.name)[1].lower() or '.png'
    path = os.path.join(DATA_DIR, f'logo{ext}')
    with open(path, 'wb') as fh:
        fh.write(uploaded_file.getbuffer())
    return path


def load_persisted_logo():
    if not os.path.isdir(DATA_DIR):
        return None
    for ext in ('.png', '.jpg', '.jpeg', '.svg'):
        p = os.path.join(DATA_DIR, f'logo{ext}')
        if os.path.exists(p):
            return p
    return None

# ----------------------------------------------------------------- I/O ---

def read_csv_any(fileobj):
    """Read a Streamlit UploadedFile (or path) as a DataFrame."""
    return pd.read_csv(fileobj)


def read_many_csv(fileobjs):
    """Concatenate zero or more uploaded CSV parts (e.g. DP Range 1 + Range 2)
    into a single DataFrame. Returns None if the list is empty."""
    if not fileobjs:
        return None
    dfs = [read_csv_any(f) for f in fileobjs]
    return pd.concat(dfs, ignore_index=True)


def read_master(fileobj):
    m = pd.read_excel(fileobj)
    m = m.rename(columns={
        'Office ID': 'office_id', 'Office Name': 'office_name',
        'Sub Office Name': 'sub_office', 'Sub Division Name': 'sub_division',
        'Office Type Code': 'office_type'
    })
    return m[['office_id', 'office_name', 'sub_office', 'sub_division', 'office_type']].copy()


# ------------------------------------------------------------- helpers ---

def safe_div(num, den):
    if den is None or pd.isna(den) or den == 0:
        return np.nan
    return num / den


# --------------------------------------------------------- computation ---

def compute_delivery(master, dp_r1, dp_r2, dp_sd):
    """dp_r1 / dp_r2 / dp_sd are DataFrames or None. Range-1 and Range-2 are the
    two explicit uploads named in the spec (Range-2 exists because the source
    system caps a single export at ~15 days) — summed together per office."""
    out = master[['office_id']].copy()
    out['invoice_r'] = np.nan
    out['deposit_r'] = np.nan
    out['invoice_sd'] = np.nan
    out['deposit_sd'] = np.nan

    parts = [d for d in (dp_r1, dp_r2) if d is not None]
    if parts:
        dp_r = pd.concat(parts, ignore_index=True)
        g = dp_r.groupby('office-id').agg(invoice_r=('invoice-count', 'sum'),
                                           deposit_r=('deposit-count', 'sum')).reset_index()
        g = g.rename(columns={'office-id': 'office_id'})
        out = out.drop(columns=['invoice_r', 'deposit_r']).merge(g, on='office_id', how='left')

    if dp_sd is not None:
        g = dp_sd.groupby('office-id').agg(invoice_sd=('invoice-count', 'sum'),
                                            deposit_sd=('deposit-count', 'sum')).reset_index()
        g = g.rename(columns={'office-id': 'office_id'})
        keep = [c for c in out.columns if c not in ('invoice_sd', 'deposit_sd')]
        out = out[keep].merge(g, on='office_id', how='left')

    if 'invoice_r' not in out.columns:
        out['invoice_r'] = np.nan
        out['deposit_r'] = np.nan
    if 'invoice_sd' not in out.columns:
        out['invoice_sd'] = np.nan
        out['deposit_sd'] = np.nan

    out['delivery_pct_r'] = out.apply(
        lambda r: safe_div(r['invoice_r'] - r['deposit_r'], r['invoice_r']) if pd.notna(r['invoice_r']) else np.nan, axis=1)
    out['delivery_pct_sd'] = out.apply(
        lambda r: safe_div(r['invoice_sd'] - r['deposit_sd'], r['invoice_sd']) if pd.notna(r['invoice_sd']) else np.nan, axis=1)
    return out[['office_id', 'invoice_r', 'deposit_r', 'invoice_sd', 'deposit_sd',
                'delivery_pct_r', 'delivery_pct_sd']]


def compute_dss(master, dss_r, dss_sd):
    master_ids = set(master['office_id'])
    excl_r = pd.DataFrame(columns=['office_id', 'office_name'])
    excl_sd = pd.DataFrame(columns=['office_id', 'office_name'])

    def calc(row):
        pdm = row['total_pdm_art_count']
        dss = row['total_dss_art_count']
        return np.nan if pdm == 0 else dss / pdm

    out = master[['office_id']].copy()
    out['pdm_r'] = np.nan; out['dssart_r'] = np.nan; out['dss_pct_r'] = np.nan
    out['pdm_sd'] = np.nan; out['dssart_sd'] = np.nan; out['dss_pct_sd'] = np.nan

    if dss_r is not None:
        excl_r = dss_r[(dss_r['office_id'] == 0) | (~dss_r['office_id'].isin(master_ids))].copy()
        f = dss_r[dss_r['office_id'].isin(master_ids)].copy()
        f['dss_pct_r'] = f.apply(calc, axis=1)
        f = f.rename(columns={'total_pdm_art_count': 'pdm_r', 'total_dss_art_count': 'dssart_r'})
        out = out.drop(columns=['pdm_r', 'dssart_r', 'dss_pct_r']).merge(
            f[['office_id', 'pdm_r', 'dssart_r', 'dss_pct_r']], on='office_id', how='left')

    if dss_sd is not None:
        excl_sd = dss_sd[(dss_sd['office_id'] == 0) | (~dss_sd['office_id'].isin(master_ids))].copy()
        f = dss_sd[dss_sd['office_id'].isin(master_ids)].copy()
        f['dss_pct_sd'] = f.apply(calc, axis=1)
        f = f.rename(columns={'total_pdm_art_count': 'pdm_sd', 'total_dss_art_count': 'dssart_sd'})
        out = out.drop(columns=['pdm_sd', 'dssart_sd', 'dss_pct_sd']).merge(
            f[['office_id', 'pdm_sd', 'dssart_sd', 'dss_pct_sd']], on='office_id', how='left')

    return out, excl_r, excl_sd


def compute_cod(master, cod_r, cod_sd):
    out = master[['office_id']].copy()
    out['digital_r'] = np.nan; out['total_r'] = np.nan; out['cod_pct_r'] = np.nan
    out['digital_sd'] = np.nan; out['total_sd'] = np.nan; out['cod_pct_sd'] = np.nan

    if cod_r is not None:
        g = cod_r.groupby('Office ID').agg(digital_r=('COD Digital Count', 'sum'),
                                            total_r=('Total COD Count', 'sum')).reset_index()
        g = g.rename(columns={'Office ID': 'office_id'})
        g['cod_pct_r'] = g.apply(lambda r: safe_div(r['digital_r'], r['total_r']), axis=1)
        out = out.drop(columns=['digital_r', 'total_r', 'cod_pct_r']).merge(g, on='office_id', how='left')

    if cod_sd is not None:
        g = cod_sd.groupby('Office ID').agg(digital_sd=('COD Digital Count', 'sum'),
                                             total_sd=('Total COD Count', 'sum')).reset_index()
        g = g.rename(columns={'Office ID': 'office_id'})
        g['cod_pct_sd'] = g.apply(lambda r: safe_div(r['digital_sd'], r['total_sd']), axis=1)
        out = out.drop(columns=['digital_sd', 'total_sd', 'cod_pct_sd']).merge(g, on='office_id', how='left')

    return out[['office_id', 'digital_r', 'total_r', 'digital_sd', 'total_sd', 'cod_pct_r', 'cod_pct_sd']]


def compute_lb(master, lb_r, lb_sd):
    out = master[['office_id']].copy()
    out['total_letterboxes'] = np.nan
    out['cleared_r'] = np.nan; out['boxes_r'] = np.nan; out['lb_pct_r'] = np.nan
    out['cleared_sd'] = np.nan; out['boxes_sd'] = np.nan; out['lb_pct_sd'] = np.nan

    boxes_const_r = None
    if lb_r is not None:
        g = lb_r.groupby('office-id').agg(cleared_r=('total-cleared', 'sum'),
                                           boxes_r=('total-letterboxes', 'sum')).reset_index()
        g = g.rename(columns={'office-id': 'office_id'})
        g['lb_pct_r'] = g.apply(lambda r: safe_div(r['cleared_r'], r['boxes_r']), axis=1)
        out = out.drop(columns=['cleared_r', 'boxes_r', 'lb_pct_r']).merge(g, on='office_id', how='left')
        boxes_const_r = lb_r.groupby('office-id')['total-letterboxes'].first().reset_index()
        boxes_const_r = boxes_const_r.rename(columns={'office-id': 'office_id', 'total-letterboxes': 'boxes_const_r'})

    if lb_sd is not None:
        lb_sd = lb_sd[lb_sd['office-id'] != 0]
        g = lb_sd.groupby('office-id').agg(cleared_sd=('total-cleared', 'sum'),
                                            boxes_sd=('total-letterboxes', 'sum')).reset_index()
        g = g.rename(columns={'office-id': 'office_id'})
        g['lb_pct_sd'] = g.apply(lambda r: safe_div(r['cleared_sd'], r['boxes_sd']), axis=1)
        out = out.drop(columns=['cleared_sd', 'boxes_sd', 'lb_pct_sd']).merge(g, on='office_id', how='left')

    if boxes_const_r is not None:
        out = out.merge(boxes_const_r, on='office_id', how='left')
        out['total_letterboxes'] = out['boxes_sd'].where(out['boxes_sd'].notna(), out['boxes_const_r'])
        out = out.drop(columns=['boxes_const_r'])
    else:
        out['total_letterboxes'] = out['boxes_sd']

    return out[['office_id', 'total_letterboxes', 'cleared_r', 'boxes_r', 'cleared_sd', 'boxes_sd',
                'lb_pct_r', 'lb_pct_sd']]


def build_full(master, dp_r1, dp_r2, dp_sd, dss_r, dss_sd, cod_r, cod_sd, lb_r, lb_sd):
    delivery = compute_delivery(master, dp_r1, dp_r2, dp_sd)
    dss, excl_r, excl_sd = compute_dss(master, dss_r, dss_sd)
    cod = compute_cod(master, cod_r, cod_sd)
    lb = compute_lb(master, lb_r, lb_sd)

    df = master.merge(delivery, on='office_id', how='left')
    df = df.merge(dss, on='office_id', how='left')
    df = df.merge(cod, on='office_id', how='left')
    df = df.merge(lb, on='office_id', how='left')

    df.loc[df['office_type'] == 'BPO', ['lb_pct_r', 'lb_pct_sd', 'total_letterboxes']] = np.nan

    avail = {
        'delivery_r': (dp_r1 is not None) or (dp_r2 is not None), 'delivery_sd': dp_sd is not None,
        'dss_r': dss_r is not None, 'dss_sd': dss_sd is not None,
        'cod_r': cod_r is not None, 'cod_sd': cod_sd is not None,
        'lb_r': lb_r is not None, 'lb_sd': lb_sd is not None,
    }
    return df, excl_r, excl_sd, avail


# --------------------------------------------------------- aggregation ---

def weighted_row(df):
    return {
        'delivery_pct_r': safe_div((df['invoice_r'] - df['deposit_r']).sum(skipna=True), df['invoice_r'].sum(skipna=True)),
        'delivery_pct_sd': safe_div((df['invoice_sd'] - df['deposit_sd']).sum(skipna=True), df['invoice_sd'].sum(skipna=True)),
        'dss_pct_r': safe_div(df['dssart_r'].sum(skipna=True), df['pdm_r'].sum(skipna=True)),
        'dss_pct_sd': safe_div(df['dssart_sd'].sum(skipna=True), df['pdm_sd'].sum(skipna=True)),
        'cod_pct_r': safe_div(df['digital_r'].sum(skipna=True), df['total_r'].sum(skipna=True)),
        'cod_pct_sd': safe_div(df['digital_sd'].sum(skipna=True), df['total_sd'].sum(skipna=True)),
        'lb_pct_r': safe_div(df['cleared_r'].sum(skipna=True), df['boxes_r'].sum(skipna=True)),
        'lb_pct_sd': safe_div(df['cleared_sd'].sum(skipna=True), df['boxes_sd'].sum(skipna=True)),
    }


def build_divisional_summary(df):
    """Sub-division row order (v9): alphabetical by name, Division Total always
    last. For Karad's 3 sub-divisions this happens to be byte-identical to the
    old fixed SUB_DIVISIONS order, but computing it means this keeps working
    correctly if SUB_DIVISIONS ever changes (e.g. for another Division)."""
    tables = {}
    for key, filt in [('excl_bo', df['office_type'] != 'BPO'),
                       ('only_bo', df['office_type'] == 'BPO'),
                       ('incl_bo', df['office_type'].notna())]:
        sub = df[filt]
        rows = []
        for sd in sorted(SUB_DIVISIONS):
            row = weighted_row(sub[sub['sub_division'] == sd])
            row['sub_division'] = sd
            rows.append(row)
        total_row = weighted_row(sub)
        total_row['sub_division'] = 'Karad Division Total'
        rows.append(total_row)
        tables[key] = rows
    return tables


def build_excl_bo_table(df):
    """37 SPO+HPO offices. Sort rule (v9, updated): offices with a genuine
    Delivery % (Range) value first, sorted by Delivery % (Range) ASCENDING
    (worst performer first); non-delivery offices (blank Delivery % R)
    unchanged — still grouped at the bottom, A-Z by name."""
    sub = df[df['office_type'] != 'BPO'].copy()
    has_data = sub[sub['delivery_pct_r'].notna()].sort_values(
        ['delivery_pct_r', 'office_name'], ascending=[True, True])
    no_data = sub[sub['delivery_pct_r'].isna()].sort_values('office_name')
    return pd.concat([has_data, no_data], ignore_index=False)


def build_defaulters(df):
    """Sort rule (v9, updated): each of the 3 tables (Delivery/DSS/COD) sorted
    by that metric's own Range % ascending — worst defaulter first, instead of
    alphabetical by office name."""
    out = {}
    for sd in SUB_DIVISIONS:
        s = df[df['sub_division'] == sd]
        out[sd] = {
            'delivery': s[s['delivery_pct_r'] < 0.90].sort_values(
                ['delivery_pct_r', 'office_name'], ascending=[True, True]),
            'dss': s[s['dss_pct_r'] < 0.90].sort_values(
                ['dss_pct_r', 'office_name'], ascending=[True, True]),
            'cod': s[s['cod_pct_r'] < 0.90].sort_values(
                ['cod_pct_r', 'office_name'], ascending=[True, True]),
        }
    return out


def build_exceptional_defaulters(df):
    params = [
        ('Delivery % (Range)', 'delivery_pct_r'), ('Delivery % (Single Date)', 'delivery_pct_sd'),
        ('DSS % (Range)', 'dss_pct_r'), ('DSS % (Single Date)', 'dss_pct_sd'),
        ('COD % (Range)', 'cod_pct_r'), ('COD % (Single Date)', 'cod_pct_sd'),
        ('LB % (Range)', 'lb_pct_r'), ('LB % (Single Date)', 'lb_pct_sd'),
    ]
    rows = []
    for _, r in df.iterrows():
        for label, col in params:
            v = r[col]
            if pd.notna(v) and v < 0.30:
                rows.append({'office_id': r['office_id'], 'office_name': r['office_name'],
                              'sub_division': r['sub_division'], 'office_type': r['office_type'],
                              'parameter': label, 'value': v})
    result = pd.DataFrame(rows)
    if len(result):
        result = result.sort_values(['sub_division', 'office_name', 'parameter']).reset_index(drop=True)
    return result
