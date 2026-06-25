"""
North Texas Sales Rep Sourcing — web app
Run-and-download tool. Users pick a location and size; verticals are fixed
(baked into the Apify task). Deploys free to Streamlit Community Cloud.
"""
import io, datetime, re
import pandas as pd
import streamlit as st
from apify_client import ApifyClient
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ---- config (token + task come from Streamlit secrets) ----
TASK_ID = st.secrets.get("TASK_ID", "natural_viburnum~north-texas-sales-reps-a")
# Dropdown options. Texas metros first (primary sourcing area), then every
# Albireo Energy office location from the global footprint (22 U.S. + 4 Europe).
LOCATION_PRESETS = [
    # Texas metros
    "Dallas-Fort Worth", "Dallas", "Fort Worth", "Austin", "Houston", "San Antonio", "Texas",
    # Albireo U.S. offices
    "Redmond, WA", "Vancouver, WA", "Anaheim, CA", "Poway, CA", "Tempe, AZ", "Tucson, AZ",
    "Salt Lake City, UT", "Denver, CO", "Omaha, NE", "Huntsville, AL", "Hoover, AL",
    "Montgomery, AL", "Tampa, FL", "Fort Lauderdale, FL", "Miami, FL", "Moosic, PA",
    "Sterling, VA", "New Castle, DE", "Gambrills, MD", "Edison, NJ", "New York City, NY",
    "Norwood, MA", "Chelmsford, MA",
    # Albireo Europe offices
    "London, United Kingdom", "Westerham, United Kingdom", "Scotland, United Kingdom",
    "Midlands, United Kingdom",
]

# Fixed search configuration (your verticals). Location + maxItems are set per run
# from the UI below; everything else stays locked so targeting never drifts.
TITLES = ["Sales Rep", "PLC", "Data Centers", "Power Quality", "Building Automation",
          "Sales", "Account Executive", "Business Development", "Sales Engineer", "Territory Manager"]
BASE_INPUT = {
    "autoQuerySegmentation": False,
    "currentJobTitles": TITLES,
    "pastJobTitles": TITLES,
    "industryIds": ["147", "923", "453", "983", "2468", "382"],
    "profileLanguages": ["English"],
    "profileScraperMode": "Full + email search",
    "recentlyChangedJobs": False,
    "recentlyPostedOnLinkedIn": False,
}

# ===================== Fit Score weights (tune here) =====================
# Priority order: job-title match (highest) > vertical > industry > the rest.
# A CURRENT-title match is weighted higher than a FORMER-title match.
# Location is scored RELATIVE to the run's target location (see fit_score);
# it's skipped entirely when the run has no location context.
FIT_W_TITLE_CURRENT = 50   # current job title vs the run's target titles (highest)
FIT_W_TITLE_FORMER  = 25   # former/past job titles (below current)
FIT_W_VERTICAL      = 20   # building automation / HVAC / controls / energy
FIT_W_INDUSTRY      = 12   # industry match
FIT_W_TENURE        =  6   # tenure (lower, equal-weight bucket)
FIT_W_SENIORITY     =  6   # seniority signal (lower, equal-weight bucket)
FIT_W_LOCATION      = 15   # proximity to the run's target location (relative)

# Keywords used by the industry / seniority factors and title normalization.
FIT_INDUSTRY_KEYWORDS = ['building automation', 'hvac', 'controls', 'control system',
    'energy', 'power', 'electrical', 'industrial automation', 'data center', 'mechanical',
    'facilities', 'building', 'automation', 'utilities', 'renewable']
FIT_SENIORITY_KEYWORDS = ['senior', 'lead', 'principal', 'manager', 'director', 'vice president',
    'head', 'chief', 'territory', 'regional', 'national', 'vp']
_TITLE_SYNONYMS = {'sr': 'senior', 'jr': 'junior', 'vp': 'vice president', 'svp': 'senior vice president',
    'ae': 'account executive', 'bd': 'business development', 'bdr': 'business development',
    'mgr': 'manager', 'exec': 'executive', 'rep': 'representative', 'eng': 'engineer',
    'biz': 'business', 'dev': 'development'}

# ============ data helpers (same logic as the local pipeline) ============
def flatten(obj, prefix=''):
    out = {}
    if isinstance(obj, dict):
        for k, v in obj.items():
            out.update(flatten(v, f"{prefix}/{k}" if prefix else k))
    elif isinstance(obj, list):
        for i, v in enumerate(obj):
            out.update(flatten(v, f"{prefix}/{i}"))
    else:
        out[prefix] = obj
    return out

def val(d, key):
    v = d.get(key)
    if v is None:
        return ''
    try:
        if isinstance(v, float) and pd.isna(v):
            return ''
    except Exception:
        pass
    s = str(v).strip()
    return '' if s.lower() == 'nan' else s

VMAP = {
    'Data Center': ['data center','datacenter','critical facilit','mission critical','colocation'],
    'Building Automation': ['building automation','hvac control','bas ','automated logic','distech','niagara','building control'],
    'PLC / Industrial Automation': ['plc','programmable logic','scada','industrial automation'],
    'Energy Management': ['energy management','energy efficiency','sustainab'],
    'Power Systems / Power Quality': ['power quality','power systems','switchgear','uninterruptible',' ups ','electrical distribution','power distribution'],
    'HVAC / Heating & Air': ['hvac','heating and air','mechanical contractor','refrigeration'],
}
COMPANY_VERTICAL = {
    'Building Automation': ['johnson controls','honeywell','siemens','schneider','automated logic','distech','alerton','delta controls','resource data management','reliable controls','kmc controls'],
    'HVAC / Heating & Air': ['lennox','trane','carrier','daikin','texas airsystems','york','goodman','rheem','ruud','aaon','mitsubishi electric','airsystems'],
    'PLC / Industrial Automation': ['ifm','sensopart','vaisala','rockwell','emerson','abb','beckhoff','omron','banner engineering','pepperl','balluff','sick ','turck','wago'],
    'Power Systems / Power Quality': ['eaton','vertiv','hitachi energy','apc','resa power','generac','cummins','kohler','mitsubishi electric power','schweitzer'],
    'Electrical Distribution': ['graybar','rexel','consolidated electrical','border states','wesco','sonepar'],
}
def vmatch(d):
    company = val(d,'currentPosition/0/companyName').lower()
    blob = ' '.join([val(d,'headline'), val(d,'currentPosition/0/position'), company, val(d,'about')]).lower()
    tags = []
    for k, names in COMPANY_VERTICAL.items():
        if any(n in company for n in names):
            tags.append(k)
    for k, kw in VMAP.items():
        if k not in tags and any(w in blob for w in kw):
            tags.append(k)
    return ', '.join(tags) if tags else '—'

def prev_roles(d):
    cur = (val(d,'currentPosition/0/position').lower(), val(d,'currentPosition/0/companyName').lower())
    seen, out = set(), []
    for i in range(25):
        pos, comp, dur = val(d,f'experience/{i}/position'), val(d,f'experience/{i}/companyName'), val(d,f'experience/{i}/duration')
        if not pos and not comp:
            continue
        k = (pos.lower(), comp.lower())
        if k == cur or k in seen:
            continue
        seen.add(k)
        s = f"{pos} @ {comp}" if comp else pos
        if dur:
            s += f"  ({dur})"
        out.append('• ' + s)
    return '\n'.join(out[:4])

def skills(d):
    s = []
    for i in range(8):
        t = val(d, f'topSkills/{i}')
        if t:
            s.append(t)
    for i in range(60):
        n = val(d, f'skills/{i}/name')
        if n and n not in s:
            s.append(n)
    return ', '.join(s[:8])

def email(d):
    for k in ['email','workEmail','emailAddress','professionalEmail','personalEmail',
              'emails/0','emails/0/email','emails/0/value','emailSearch/email',
              'emailSearch/0/email','contactInfo/email','contact/email','emailAddresses/0']:
        e = val(d, k)
        if e and '@' in e:
            return e
    return ''

# ============================ Fit Score ============================
def _tokens(s):
    """Lower-cased word tokens with a few title abbreviations expanded."""
    s = re.sub(r'[^a-z0-9 ]', ' ', (s or '').lower())
    out = []
    for t in s.split():
        out.extend(_TITLE_SYNONYMS.get(t, t).split())
    return set(out)

def _title_match_fraction(title, targets):
    """Fuzzy/partial match of one title against the target title list -> 0..1.
    Full credit when every word of a target title appears in the candidate
    title (so 'Sr. Account Executive' fully matches 'Account Executive');
    partial credit for partial word overlap."""
    toks = _tokens(title)
    if not toks:
        return 0.0
    best = 0.0
    for tgt in targets:
        tt = _tokens(tgt)
        if not tt:
            continue
        if tt <= toks:
            return 1.0
        best = max(best, len(tt & toks) / len(tt))
    return best

def _former_titles(d):
    out = []
    for i in range(25):
        p = val(d, f'experience/{i}/position')
        if p:
            out.append(p)
    for i in range(1, 6):  # any current positions beyond the primary one
        p = val(d, f'currentPosition/{i}/position')
        if p:
            out.append(p)
    return out

def _vertical_fraction(vertical):
    tags = [t for t in (vertical or '').split(',') if t.strip() and t.strip() != '—']
    if not tags:
        return 0.0
    return 0.7 if len(tags) == 1 else 1.0

def _industry_text(d):
    for k in ['currentPosition/0/companyIndustry', 'currentPosition/0/industry',
              'company/industry', 'companyIndustry', 'industry', 'industryName', 'occupation']:
        t = val(d, k)
        if t:
            return t.lower()
    return ''

def _industry_fraction(d):
    t = _industry_text(d)
    if not t:
        return 0.0
    return 1.0 if any(k in t for k in FIT_INDUSTRY_KEYWORDS) else 0.25

def _years_from_duration(s):
    s = (s or '').lower()
    yrs = 0.0
    m = re.search(r'(\d+)\s*(?:yr|year)', s)
    if m:
        yrs += int(m.group(1))
    m = re.search(r'(\d+)\s*(?:mo|month)', s)
    if m:
        yrs += int(m.group(1)) / 12.0
    return yrs

def _tenure_fraction(tenure):
    y = _years_from_duration(tenure)
    if y <= 0:
        return 0.0
    if y >= 3:
        return 1.0
    return 0.7 if y >= 1.5 else 0.4

def _seniority_fraction(title):
    t = (title or '').lower()
    return 1.0 if any(k in t for k in FIT_SENIORITY_KEYWORDS) else 0.0

# Location scoring is relative to the run's target location (no hardcoded city).
_STATE_ABBR = {'al','ak','az','ar','ca','co','ct','de','fl','ga','hi','id','il','in','ia','ks',
    'ky','la','me','md','ma','mi','mn','ms','mo','mt','ne','nv','nh','nj','nm','ny','nc','nd',
    'oh','ok','or','pa','ri','sc','sd','tn','tx','ut','vt','va','wa','wv','wi','wy'}
_STATE_NAME_ABBR = {'alabama':'al','alaska':'ak','arizona':'az','arkansas':'ar','california':'ca',
    'colorado':'co','connecticut':'ct','delaware':'de','florida':'fl','georgia':'ga','hawaii':'hi',
    'idaho':'id','illinois':'il','indiana':'in','iowa':'ia','kansas':'ks','kentucky':'ky',
    'louisiana':'la','maine':'me','maryland':'md','massachusetts':'ma','michigan':'mi',
    'minnesota':'mn','mississippi':'ms','missouri':'mo','montana':'mt','nebraska':'ne',
    'nevada':'nv','ohio':'oh','oklahoma':'ok','oregon':'or','pennsylvania':'pa','tennessee':'tn',
    'texas':'tx','utah':'ut','vermont':'vt','virginia':'va','washington':'wa','wisconsin':'wi',
    'wyoming':'wy'}
_LOC_STOP = {'area','greater','metro','metropolitan','city','of','the','region','county'}

def _loc_tokens(s):
    s = re.sub(r'[^a-z0-9 ]', ' ', (s or '').lower())
    out = []
    for t in s.split():
        if not t or t in _LOC_STOP:
            continue
        out.append(_STATE_NAME_ABBR.get(t, t))  # normalize full state names -> abbr
    return out

def _location_fraction(cand_loc, target_loc):
    """Higher when the candidate is in/near the run's target location.
    Returns None when there's no location context (so it's left out of the
    score rather than penalizing anyone)."""
    if not (target_loc or '').strip():
        return None
    cand = set(_loc_tokens(cand_loc))
    if not cand:
        return 0.0
    tgt = _loc_tokens(target_loc)
    cities = [t for t in tgt if t not in _STATE_ABBR]
    states = [t for t in tgt if t in _STATE_ABBR]
    if cities:
        if any(c in cand for c in cities):
            return 1.0                      # same city / metro / region
        if any(s in cand for s in states):
            return 0.5                      # same state -> "near"
        return 0.0
    # state-only target (e.g. "Texas") — being in that state is a full match
    return 1.0 if any(s in cand for s in states) else 0.0

def fit_score(d, rec, target_titles, target_location):
    """0–100 fit score; normalized over whichever factors apply this run."""
    targets = [t for t in (target_titles or []) if t]
    parts = [
        (FIT_W_TITLE_CURRENT, _title_match_fraction(rec.get('Current Title', ''), targets)),
        (FIT_W_TITLE_FORMER,  max((_title_match_fraction(t, targets)
                                   for t in _former_titles(d)), default=0.0)),
        (FIT_W_VERTICAL,      _vertical_fraction(rec.get('Vertical Match', ''))),
        (FIT_W_INDUSTRY,      _industry_fraction(d)),
        (FIT_W_TENURE,        _tenure_fraction(rec.get('Tenure', ''))),
        (FIT_W_SENIORITY,     _seniority_fraction(rec.get('Current Title', ''))),
    ]
    loc = _location_fraction(rec.get('Location', ''), target_location)
    if loc is not None:
        parts.append((FIT_W_LOCATION, loc))
    total = sum(w for w, _ in parts) or 1
    return round(sum(w * f for w, f in parts) / total * 100)

# Display columns: Fit Score first, then Location moved to right after Current Title.
COLUMNS = ['Fit Score','First Name','Last Name','Current Title','Location','Current Company',
           'Vertical Match','Headline','Tenure','Open To Work','Email','Previous Roles',
           'Top Skills','LinkedIn URL','Company LinkedIn','Connections','Summary']

def condense(d):
    otw = val(d,'openToWork').lower()
    otw = 'Yes' if otw == 'true' else ('No' if otw == 'false' else '')
    return {
        'First Name': val(d,'firstName'), 'Last Name': val(d,'lastName'),
        'Current Title': val(d,'currentPosition/0/position'),
        'Current Company': val(d,'currentPosition/0/companyName'),
        'Vertical Match': vmatch(d), 'Headline': val(d,'headline'),
        'Location': val(d,'location/linkedinText'),
        'Tenure': val(d,'currentPosition/0/duration'), 'Open To Work': otw,
        'Email': email(d), 'Previous Roles': prev_roles(d), 'Top Skills': skills(d),
        'LinkedIn URL': val(d,'linkedinUrl'),
        'Company LinkedIn': val(d,'currentPosition/0/companyLinkedinUrl'),
        'Connections': val(d,'connectionsCount'), 'Summary': val(d,'about'),
    }

def _excel_summary_sheet(ws, summary):
    """Write a small run-summary block (counts) onto its own sheet."""
    title_fill = PatternFill('solid', fgColor='1F4E5F')
    head = Font(name='Calibri', bold=True, color='FFFFFF', size=13)
    label = Font(name='Calibri', bold=True, size=11); base = Font(name='Calibri', size=11)
    ws['A1'] = 'Run Summary'; ws['A1'].font = head; ws['A1'].fill = title_fill
    ws['B1'].fill = title_fill
    rows = [
        ('Search location', summary.get('Location', '')),
        ('Date', summary.get('Timestamp', '')),
        ('Profiles requested', summary.get('Requested', '')),
        ('Profiles found', summary.get('Found', '')),
        ('New (added to master)', summary.get('New', '')),
        ('Duplicates (already on file)', summary.get('Dupes', '')),
        ('Master list total', summary.get('MasterTotal', '')),
    ]
    for i, (k, v) in enumerate(rows, start=3):
        ws.cell(i, 1, k).font = label
        ws.cell(i, 2, v).font = base
    ws.column_dimensions['A'].width = 30; ws.column_dimensions['B'].width = 28
    ws.sheet_view.showGridLines = False

def build_excel_bytes(rows, cols=COLUMNS, summary=None):
    wb = Workbook()
    if summary is not None:
        ws_sum = wb.active; ws_sum.title = 'Summary'; _excel_summary_sheet(ws_sum, summary)
        ws = wb.create_sheet('Sales Reps')
    else:
        ws = wb.active; ws.title = 'Sales Reps'
    hf = PatternFill('solid', fgColor='1F4E5F'); hfont = Font(name='Calibri', bold=True, color='FFFFFF', size=11)
    band = PatternFill('solid', fgColor='F2F6F7'); thin = Side(style='thin', color='E2E2E2')
    bd = Border(left=thin, right=thin, top=thin, bottom=thin)
    lf = Font(name='Calibri', color='0563C1', underline='single', size=11); base = Font(name='Calibri', size=11)
    wrap = {'Current Title','Vertical Match','Headline','Previous Roles','Top Skills','Location'}
    links = {'LinkedIn URL','Company LinkedIn'}
    ws.append(cols)
    for ci, h in enumerate(cols, 1):
        c = ws.cell(1, ci); c.fill = hf; c.font = hfont; c.border = bd
        c.alignment = Alignment(vertical='center', wrap_text=True)
    for ri, row in enumerate(rows, start=2):
        for ci, h in enumerate(cols, 1):
            cell = ws.cell(ri, ci, row.get(h, '')); cell.border = bd
            cell.alignment = Alignment(vertical='top', horizontal=('right' if h in ('Connections','Fit Score') else 'left'), wrap_text=(h in wrap))
            if h in links and cell.value:
                cell.hyperlink = cell.value; cell.font = lf
            else:
                cell.font = base
            if ri % 2 == 0:
                cell.fill = band
    W = {'Fit Score':10,'Date Added':13,'Search Location':18,'First Name':12,'Last Name':13,
         'Current Title':26,'Current Company':22,'Vertical Match':22,
         'Headline':38,'Location':24,'Tenure':12,'Open To Work':11,'Email':30,'Previous Roles':46,
         'Top Skills':28,'LinkedIn URL':34,'Company LinkedIn':34,'Connections':12,'Summary':55}
    for ci, h in enumerate(cols, 1):
        ws.column_dimensions[get_column_letter(ci)].width = W.get(h, 16)
    ws.freeze_panes = 'A2'; ws.auto_filter.ref = f"A1:{get_column_letter(len(cols))}1"
    ws.sheet_view.showGridLines = False
    buf = io.BytesIO(); wb.save(buf); buf.seek(0)
    return buf.getvalue()

def _dsid(r):
    if isinstance(r, dict):
        return r.get("defaultDatasetId") or r.get("default_dataset_id")
    return getattr(r, "default_dataset_id", None) or getattr(r, "defaultDatasetId", None)

# ============ persistence: Google Sheets master list + search history ============
# Keeps a running master list (deduped by LinkedIn URL) and a log of every search,
# so new-vs-duplicate counts work across runs and days. If the Google Sheet secrets
# aren't configured yet, everything falls back to in-session memory so the app still
# runs — see SETUP_GOOGLE_SHEETS.md to connect a sheet.
# Master sheet order: Date Added, then Fit Score (2nd), then Search Location, then the rest.
MASTER_HEADERS = ["Date Added", "Fit Score", "Search Location"] + [c for c in COLUMNS if c != "Fit Score"]
SEARCH_HEADERS = ["Timestamp", "Location", "Requested", "Found", "New", "Dupes"]

def _norm_url(u):
    u = (u or "").strip().lower()
    for p in ("https://", "http://", "www."):
        u = u.replace(p, "")
    return u.rstrip("/")

def gs_configured():
    try:
        has_creds = ("gcp_service_account_json" in st.secrets
                     or "gcp_service_account" in st.secrets)
        return has_creds and "MASTER_SHEET_ID" in st.secrets
    except Exception:
        return False

@st.cache_resource(show_spinner=False)
def _gs_book():
    import json, gspread
    from google.oauth2.service_account import Credentials
    # Accept either the whole service-account JSON pasted as one string
    # (easiest), or a classic [gcp_service_account] TOML table.
    if "gcp_service_account_json" in st.secrets:
        info = json.loads(st.secrets["gcp_service_account_json"])
    else:
        info = dict(st.secrets["gcp_service_account"])
    creds = Credentials.from_service_account_info(
        info, scopes=["https://www.googleapis.com/auth/spreadsheets"])
    return gspread.authorize(creds).open_by_key(st.secrets["MASTER_SHEET_ID"])

# ---- cosmetic formatting for the Google Sheet tabs (best-effort) ----
_NAVY_RGB = {"red": 0.055, "green": 0.176, "blue": 0.322}  # #0E2D52
_COLW = {
    'Fit Score': 75, 'Date Added': 95, 'Search Location': 135, 'First Name': 95, 'Last Name': 100,
    'Current Title': 190, 'Current Company': 165, 'Vertical Match': 165, 'Headline': 290,
    'Location': 175, 'Tenure': 95, 'Open To Work': 95, 'Email': 220, 'Previous Roles': 320,
    'Top Skills': 200, 'LinkedIn URL': 240, 'Company LinkedIn': 240, 'Connections': 100,
    'Summary': 380, 'Timestamp': 135, 'Requested': 95, 'Found': 80, 'New': 75, 'Dupes': 80,
}

def _format_worksheet(ws, headers):
    """Frozen bold header, sane column widths, zebra rows. Cosmetic only —
    each piece is best-effort so styling can never break the app."""
    header = ws.row_values(1) or list(headers)  # style the sheet's actual columns
    ncols = len(header)
    last = get_column_letter(ncols)
    try:
        ws.freeze(rows=1)
        ws.format(f"A1:{last}1", {
            "backgroundColor": _NAVY_RGB, "verticalAlignment": "MIDDLE",
            "textFormat": {"foregroundColor": {"red": 1, "green": 1, "blue": 1},
                           "bold": True, "fontSize": 10}})
    except Exception:
        pass
    try:
        ws.spreadsheet.batch_update({"requests": [
            {"updateDimensionProperties": {
                "range": {"sheetId": ws.id, "dimension": "COLUMNS",
                          "startIndex": i, "endIndex": i + 1},
                "properties": {"pixelSize": _COLW.get(h, 130)}, "fields": "pixelSize"}}
            for i, h in enumerate(header)]})
    except Exception:
        pass
    try:  # alternating row shading; ignored if a band already exists
        ws.spreadsheet.batch_update({"requests": [{"addBanding": {"bandedRange": {
            "range": {"sheetId": ws.id, "startRowIndex": 0,
                      "startColumnIndex": 0, "endColumnIndex": ncols},
            "rowProperties": {"headerColor": _NAVY_RGB,
                              "firstBandColor": {"red": 1, "green": 1, "blue": 1},
                              "secondBandColor": {"red": 0.949, "green": 0.965, "blue": 0.969}}}}}]})
    except Exception:
        pass

def _maybe_format(ws, title, headers, force=False):
    done = st.session_state.setdefault("_ws_formatted", set())
    if force or title not in done:
        _format_worksheet(ws, headers)
        done.add(title)

def _ws(title, headers):
    import gspread
    sh = _gs_book()
    created = False
    try:
        ws = sh.worksheet(title)
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet(title=title, rows=2000, cols=len(headers))
        ws.append_row(headers, value_input_option="USER_ENTERED")
        created = True
    if not ws.row_values(1):
        ws.append_row(headers, value_input_option="USER_ENTERED")
    _maybe_format(ws, title, headers, force=created)
    return ws

def load_master():
    """Return list of dicts (oldest first) for every unique rep saved so far."""
    if gs_configured():
        try:
            return _ws("Master", MASTER_HEADERS).get_all_records()
        except Exception as e:
            st.warning(f"Couldn't read the master sheet: {e}")
            return []
    return st.session_state.setdefault("master", [])

def load_searches():
    """Return list of dicts (oldest first) for every search run so far."""
    if gs_configured():
        try:
            return _ws("Searches", SEARCH_HEADERS).get_all_records()
        except Exception as e:
            st.warning(f"Couldn't read search history: {e}")
            return []
    return st.session_state.setdefault("searches", [])

def _ensure_columns(ws, needed):
    """Make sure every needed column exists in the sheet's header, appending any
    new ones (e.g. 'Fit Score') at the end so existing rows stay aligned.
    Returns the sheet's actual header order to write rows against."""
    header = ws.row_values(1)
    if not header:
        ws.update(range_name="A1", values=[list(needed)])
        return list(needed)
    missing = [h for h in needed if h not in header]
    if missing:
        header = header + missing
        ws.update(range_name="A1", values=[header])
    return header

def save_master(rows):
    if not rows:
        return
    if gs_configured():
        ws = _ws("Master", MASTER_HEADERS)
        header = _ensure_columns(ws, MASTER_HEADERS)   # adds Fit Score column if absent
        ws.append_rows([[r.get(h, "") for h in header] for r in rows],
                       value_input_option="USER_ENTERED")
    else:
        st.session_state.setdefault("master", []).extend(rows)

def save_search(rec):
    if gs_configured():
        ws = _ws("Searches", SEARCH_HEADERS)
        ws.append_row([rec.get(h, "") for h in SEARCH_HEADERS],
                      value_input_option="USER_ENTERED")
    else:
        st.session_state.setdefault("searches", []).append(rec)

# ============================ UI ============================
st.set_page_config(page_title="Albireo Energy · Sales Rep Sourcing", page_icon="🎯", layout="wide")

NAVY = "#0E2D52"; NAVY_DARK = "#081C36"; BLUE = "#1F6FB2"
GOLD = "#FDB813"; GOLD_DARK = "#E0A200"; INK = "#16263B"
GRAY = "#E5EAF0"; CARD = "#FFFFFF"; BORDER = "#D6DEE8"

st.markdown(f"""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');
html, body, .stApp, button, input, textarea, select,
[data-testid="stMarkdownContainer"] {{ font-family: 'Inter', -apple-system, sans-serif; }}

.stApp {{ background: {GRAY}; }}
.block-container {{ padding-top: 2.2rem; padding-bottom: 2rem;
    padding-left: 3rem; padding-right: 3rem; max-width: 1320px; }}

/* header banner */
.ae-header {{
    background: linear-gradient(135deg, {NAVY} 0%, {NAVY_DARK} 100%);
    border-bottom: 5px solid {GOLD};
    border-radius: 18px; padding: 40px 44px; margin-bottom: 30px;
    box-shadow: 0 10px 30px rgba(8,28,54,0.25);
}}
.ae-eyebrow {{ color: {GOLD} !important; font-weight: 800; letter-spacing: 3px;
    font-size: 14px; text-transform: uppercase; margin: 0; }}
.ae-title {{ color: #ffffff !important; font-size: 46px; font-weight: 800;
    margin: 8px 0 0 0; line-height: 1.05; }}
.ae-chips {{ display: flex; flex-wrap: wrap; gap: 8px; margin-top: 18px; }}
.ae-chip {{ background: rgba(255,255,255,0.10); color: #cfdbea;
    border: 1px solid rgba(255,255,255,0.18); font-size: 12px; font-weight: 600;
    padding: 5px 13px; border-radius: 999px; }}

/* input section labels */
.ae-label {{ font-size: 22px; font-weight: 800; color: {NAVY}; margin: 2px 0 2px 0; }}
.ae-help  {{ font-size: 13px; color: #66788c; margin: 0 0 12px 0; }}

/* gold run button (navy text) */
div.stButton, div[data-testid="stButton"] {{ width: 100% !important; }}
div.stButton > button, div[data-testid="stButton"] > button {{
    background: linear-gradient(135deg, #FFC62E 0%, {GOLD} 100%);
    color: {NAVY}; font-size: 24px; font-weight: 800; letter-spacing: .4px;
    padding: 20px 0; border: none; border-radius: 14px; width: 100% !important;
    white-space: nowrap; line-height: 1.2;
    box-shadow: 0 8px 22px rgba(253,184,19,0.45); transition: all .15s ease;
}}
div.stButton > button:hover, div[data-testid="stButton"] > button:hover {{
    background: linear-gradient(135deg, {GOLD} 0%, {GOLD_DARK} 100%);
    color: {NAVY}; transform: translateY(-2px); box-shadow: 0 12px 28px rgba(253,184,19,0.52);
}}
div.stButton > button:active, div[data-testid="stButton"] > button:active {{ transform: translateY(0); }}

/* navy download button */
div.stDownloadButton > button {{
    background: {NAVY}; color: #fff; font-size: 15px; font-weight: 700;
    padding: 12px 26px; border: none; border-radius: 10px;
    box-shadow: 0 4px 14px rgba(14,45,82,0.25); transition: all .15s ease;
}}
div.stDownloadButton > button:hover {{ background: {NAVY_DARK}; color: #fff; transform: translateY(-1px); }}

/* input cards */
div[data-testid="stVerticalBlockBorderWrapper"] {{
    background: {CARD}; border-radius: 16px;
    box-shadow: 0 4px 18px rgba(14,45,82,0.09); border: 1px solid {BORDER} !important;
}}
div[data-testid="stVerticalBlockBorderWrapper"] > div {{ padding: 18px 22px; min-height: 150px; }}

/* empty state */
.ae-empty {{ background: #fff; border: 2px dashed #c4d0de; border-radius: 16px;
    padding: 60px 24px; text-align: center; margin-top: 8px; }}
.ae-empty .big {{ font-size: 46px; line-height: 1; }}
.ae-empty .t {{ font-size: 18px; font-weight: 700; color: {NAVY}; margin-top: 12px; }}
.ae-empty .s {{ font-size: 14px; color: #7a8a9c; margin-top: 4px; }}

/* branded results banner */
.ae-result {{ background: {BLUE}; color: #fff; font-weight: 700; font-size: 16px;
    padding: 13px 20px; border-radius: 12px; margin: 8px 0 16px 0;
    box-shadow: 0 4px 14px rgba(31,111,178,0.25); }}

/* top tabs */
div[data-testid="stTabs"] div[role="tablist"] {{
    gap: 4px; border-bottom: 2px solid {BORDER}; margin-bottom: 18px; }}
div[data-testid="stTabs"] button[role="tab"] {{
    padding: 10px 22px; border-radius: 10px 10px 0 0; }}
div[data-testid="stTabs"] button[role="tab"] p {{
    font-size: 16px; font-weight: 700; color: #66788c; }}
div[data-testid="stTabs"] button[role="tab"][aria-selected="true"] {{
    background: #fff; }}
div[data-testid="stTabs"] button[role="tab"][aria-selected="true"] p {{
    color: {NAVY}; font-weight: 800; }}

/* footer */
.ae-footer {{ text-align: center; color: #8a98a8; font-size: 12px;
    margin-top: 36px; padding-top: 18px; border-top: 1px solid {BORDER}; }}
</style>
""", unsafe_allow_html=True)

st.markdown("""
<div class="ae-header">
    <p class="ae-eyebrow">Albireo Energy</p>
    <h1 class="ae-title">Sales Rep Sourcing</h1>
</div>
""", unsafe_allow_html=True)

tab_search, tab_history, tab_master = st.tabs(
    ["🔎  New Search", "🕘  Previous Searches", "📒  Master List"])

# ------------------------------- New Search -------------------------------
with tab_search:
    col1, col2 = st.columns(2, gap="large")
    with col1:
        with st.container(border=True):
            st.markdown('<p class="ae-label">📍 Location</p>', unsafe_allow_html=True)
            st.markdown('<p class="ae-help">Where to search for reps</p>', unsafe_allow_html=True)
            choice = st.selectbox("Location", LOCATION_PRESETS + ["Custom…"], label_visibility="collapsed")
            location = (st.text_input("Enter a location", label_visibility="collapsed",
                                      placeholder="Type a city or metro area")
                        if choice == "Custom…" else choice)
    with col2:
        with st.container(border=True):
            st.markdown('<p class="ae-label">🎯 Number of Profiles</p>', unsafe_allow_html=True)
            st.markdown('<p class="ae-help">How many reps to pull this run</p>', unsafe_allow_html=True)
            size = st.slider("Number of profiles to pull", 10, 200, 25, step=5, label_visibility="collapsed")

    st.write("")
    bcol = st.columns([1, 6, 1])
    with bcol[1]:
        clicked = st.button("🔎  Run Sourcing", type="primary", use_container_width=True)

    if not gs_configured():
        st.info("⚠️ Google Sheet not connected yet — searches still run, but the master list and "
                "history won't be saved between sessions. See **SETUP_GOOGLE_SHEETS.md** to connect one.")

    if clicked:
        if not location:
            st.warning("Please choose or enter a location.")
            st.stop()
        try:
            with st.spinner("Searching LinkedIn… this usually takes a minute or two."):
                client = ApifyClient(st.secrets["APIFY_TOKEN"])
                # full config sent every run, so verticals/titles/mode can never drift
                run_input = dict(BASE_INPUT)
                run_input["locations"] = [location]
                run_input["maxItems"] = size
                run = client.task(TASK_ID).call(task_input=run_input)
                ds = _dsid(run)
                items = [flatten(i) for i in client.dataset(ds).iterate_items()]
            recs = []
            for d in items:
                if not val(d, 'linkedinUrl'):
                    continue
                rec = condense(d)
                # Fit Score is scored against this run's target titles + location.
                rec['Fit Score'] = fit_score(d, rec, TITLES, location)
                recs.append(rec)
            if not recs:
                st.error("No results came back. This usually means the Apify free-tier run limit was hit "
                         "or the location returned no matches. Check the Apify console.")
                st.stop()
            # highest Fit Score first (ties broken by company)
            recs.sort(key=lambda r: (-r['Fit Score'], r['Current Company']))

            # ---- dedup against the running master list ----
            master = load_master()
            seen = {_norm_url(m.get('LinkedIn URL', '')) for m in master}
            new_recs, dupes = [], 0
            for r in recs:
                k = _norm_url(r['LinkedIn URL'])
                if k and k not in seen:
                    seen.add(k); new_recs.append(r)
                else:
                    dupes += 1

            # ---- persist: new reps -> master, this run -> history ----
            today = datetime.date.today().isoformat()
            now = datetime.datetime.now().isoformat(timespec='minutes').replace('T', ' ')
            save_master([{'Date Added': today, 'Search Location': location, **r} for r in new_recs])
            search_rec = {'Timestamp': now, 'Location': location, 'Requested': size,
                          'Found': len(recs), 'New': len(new_recs), 'Dupes': dupes}
            save_search(search_rec)
            summary = {**search_rec, 'MasterTotal': len(master) + len(new_recs)}

            df = pd.DataFrame(recs, columns=COLUMNS)
            st.markdown(
                f'<div class="ae-result">✓ Found {len(df)} reps near {location}'
                f' &nbsp;·&nbsp; <b>{len(new_recs)} new</b> added to master'
                f' &nbsp;·&nbsp; {dupes} already on file</div>', unsafe_allow_html=True)

            m1, m2, m3, m4 = st.columns(4)
            m1.metric("Profiles found", len(df))
            m2.metric("New", len(new_recs))
            m3.metric("Duplicates", dupes)
            m4.metric("Master total", summary['MasterTotal'])

            st.dataframe(df, use_container_width=True, hide_index=True)
            st.download_button("⬇ Download Excel", build_excel_bytes(recs, summary=summary),
                               file_name=f"sales_reps_{location.replace(' ','_')}_{today}.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        except Exception as e:
            st.error(f"Something went wrong talking to Apify: {e}")
    else:
        st.markdown(
            '<div class="ae-empty">'
            '<div class="big">🔍</div>'
            '<div class="t">Your sourced reps will appear here</div>'
            '<div class="s">Pick a location and size above, then run a search to build your list.</div>'
            '</div>', unsafe_allow_html=True)

# ---------------------------- Previous Searches ----------------------------
with tab_history:
    st.markdown('<p class="ae-label">🕘 Previous Searches</p>', unsafe_allow_html=True)
    st.markdown('<p class="ae-help">Every run, newest first — with new vs. duplicate counts.</p>',
                unsafe_allow_html=True)
    searches = load_searches()
    if not searches:
        st.markdown(
            '<div class="ae-empty"><div class="big">🕘</div>'
            '<div class="t">No searches yet</div>'
            '<div class="s">Run a search and it will be logged here automatically.</div></div>',
            unsafe_allow_html=True)
    else:
        hdf = pd.DataFrame(searches)
        for c in SEARCH_HEADERS:
            if c not in hdf.columns:
                hdf[c] = ''
        hdf = hdf[SEARCH_HEADERS].iloc[::-1].reset_index(drop=True)  # newest first
        total_new = pd.to_numeric(hdf['New'], errors='coerce').fillna(0).astype(int).sum()
        c1, c2, c3 = st.columns(3)
        c1.metric("Total searches", len(hdf))
        c2.metric("Profiles added (all time)", int(total_new))
        c3.metric("Master list size", len(load_master()))
        st.dataframe(hdf, use_container_width=True, hide_index=True)

# ------------------------------- Master List -------------------------------
with tab_master:
    st.markdown('<p class="ae-label">📒 Master List</p>', unsafe_allow_html=True)
    st.markdown('<p class="ae-help">Every unique rep sourced so far, deduped by LinkedIn URL.</p>',
                unsafe_allow_html=True)
    master = load_master()
    if not master:
        st.markdown(
            '<div class="ae-empty"><div class="big">📒</div>'
            '<div class="t">Master list is empty</div>'
            '<div class="s">New reps from each search land here automatically.</div></div>',
            unsafe_allow_html=True)
    else:
        mdf = pd.DataFrame(master)
        order = [c for c in MASTER_HEADERS if c in mdf.columns]
        mdf = mdf[order]
        if 'Fit Score' in mdf.columns:  # highest Fit Score first
            mdf = (mdf.assign(_fs=pd.to_numeric(mdf['Fit Score'], errors='coerce').fillna(-1))
                      .sort_values('_fs', ascending=False).drop(columns='_fs').reset_index(drop=True))
        st.markdown(f'<div class="ae-result">📒 {len(mdf)} unique reps in the master list</div>',
                    unsafe_allow_html=True)
        st.dataframe(mdf, use_container_width=True, hide_index=True)
        stamp = datetime.date.today().isoformat()
        st.download_button("⬇ Download full master (Excel)",
                           build_excel_bytes(mdf.to_dict('records'), cols=order),
                           file_name=f"sales_reps_master_{stamp}.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                           key="dl_master")

st.markdown('<div class="ae-footer">Albireo Energy · Internal sourcing tool · Powered by LinkedIn public data via Apify</div>',
            unsafe_allow_html=True)
