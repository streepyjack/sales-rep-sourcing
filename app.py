"""
North Texas Sales Rep Sourcing — web app
Run-and-download tool. Users pick a location and size; verticals are fixed
(baked into the Apify task). Deploys free to Streamlit Community Cloud.
"""
import io, datetime
import pandas as pd
import streamlit as st
from apify_client import ApifyClient
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ---- config (token + task come from Streamlit secrets) ----
TASK_ID = st.secrets.get("TASK_ID", "natural_viburnum~north-texas-sales-reps-a")
LOCATION_PRESETS = ["Dallas-Fort Worth", "Dallas", "Fort Worth", "Austin",
                    "Houston", "San Antonio", "Texas"]

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

COLUMNS = ['First Name','Last Name','Current Title','Current Company','Vertical Match','Headline',
           'Location','Tenure','Open To Work','Email','Previous Roles','Top Skills',
           'LinkedIn URL','Company LinkedIn','Connections','Summary']

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

def build_excel_bytes(rows):
    wb = Workbook(); ws = wb.active; ws.title = 'Sales Reps'
    hf = PatternFill('solid', fgColor='1F4E5F'); hfont = Font(name='Calibri', bold=True, color='FFFFFF', size=11)
    band = PatternFill('solid', fgColor='F2F6F7'); thin = Side(style='thin', color='E2E2E2')
    bd = Border(left=thin, right=thin, top=thin, bottom=thin)
    lf = Font(name='Calibri', color='0563C1', underline='single', size=11); base = Font(name='Calibri', size=11)
    wrap = {'Current Title','Vertical Match','Headline','Previous Roles','Top Skills','Location'}
    links = {'LinkedIn URL','Company LinkedIn'}
    ws.append(COLUMNS)
    for ci, h in enumerate(COLUMNS, 1):
        c = ws.cell(1, ci); c.fill = hf; c.font = hfont; c.border = bd
        c.alignment = Alignment(vertical='center', wrap_text=True)
    for ri, row in enumerate(rows, start=2):
        for ci, h in enumerate(COLUMNS, 1):
            cell = ws.cell(ri, ci, row.get(h, '')); cell.border = bd
            cell.alignment = Alignment(vertical='top', horizontal=('right' if h=='Connections' else 'left'), wrap_text=(h in wrap))
            if h in links and cell.value:
                cell.hyperlink = cell.value; cell.font = lf
            else:
                cell.font = base
            if ri % 2 == 0:
                cell.fill = band
    W = {'First Name':12,'Last Name':13,'Current Title':26,'Current Company':22,'Vertical Match':22,
         'Headline':38,'Location':24,'Tenure':12,'Open To Work':11,'Email':30,'Previous Roles':46,
         'Top Skills':28,'LinkedIn URL':34,'Company LinkedIn':34,'Connections':12,'Summary':55}
    for ci, h in enumerate(COLUMNS, 1):
        ws.column_dimensions[get_column_letter(ci)].width = W.get(h, 16)
    ws.freeze_panes = 'C2'; ws.auto_filter.ref = f"A1:{get_column_letter(len(COLUMNS))}1"
    ws.sheet_view.showGridLines = False
    buf = io.BytesIO(); wb.save(buf); buf.seek(0)
    return buf.getvalue()

def _dsid(r):
    if isinstance(r, dict):
        return r.get("defaultDatasetId") or r.get("default_dataset_id")
    return getattr(r, "default_dataset_id", None) or getattr(r, "defaultDatasetId", None)

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
    color: {NAVY}; font-size: 23px; font-weight: 800; letter-spacing: .4px;
    padding: 22px 0; border: none; border-radius: 14px; width: 100% !important;
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
bcol = st.columns([1, 3, 1])
with bcol[1]:
    clicked = st.button("🔎  Run Sourcing", type="primary")

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
        recs = [condense(d) for d in items if val(d, 'linkedinUrl')]
        if not recs:
            st.error("No results came back. This usually means the Apify free-tier run limit was hit "
                     "or the location returned no matches. Check the Apify console.")
            st.stop()
        # on-target first
        recs.sort(key=lambda r: (r['Vertical Match'] == '—', r['Current Company']))
        df = pd.DataFrame(recs, columns=COLUMNS)
        st.markdown(f'<div class="ae-result">✓ Found {len(df)} reps near {location}</div>',
                    unsafe_allow_html=True)
        st.dataframe(df, use_container_width=True, hide_index=True)
        stamp = datetime.date.today().isoformat()
        st.download_button("⬇ Download Excel", build_excel_bytes(recs),
                           file_name=f"sales_reps_{location.replace(' ','_')}_{stamp}.xlsx",
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

st.markdown('<div class="ae-footer">Albireo Energy · Internal sourcing tool · Powered by LinkedIn public data via Apify</div>',
            unsafe_allow_html=True)
