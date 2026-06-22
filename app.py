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
st.set_page_config(page_title="North Texas Sales Rep Sourcing", page_icon="🎯", layout="wide")
st.title("🎯 North Texas Sales Rep Sourcing")
st.caption("Find sales reps in controls, building automation, HVAC, power, and data-center companies. "
           "Pick a location and how many to pull, then download the list.")

col1, col2 = st.columns([2, 1])
with col1:
    choice = st.selectbox("Location", LOCATION_PRESETS + ["Custom…"])
    location = st.text_input("Enter a location") if choice == "Custom…" else choice
with col2:
    size = st.slider("Number of profiles to pull", 10, 200, 25, step=5)

if st.button("Run sourcing", type="primary"):
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
        st.success(f"Found {len(df)} reps near {location}.")
        st.dataframe(df, use_container_width=True, hide_index=True)
        stamp = datetime.date.today().isoformat()
        st.download_button("⬇ Download Excel", build_excel_bytes(recs),
                           file_name=f"sales_reps_{location.replace(' ','_')}_{stamp}.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    except Exception as e:
        st.error(f"Something went wrong talking to Apify: {e}")
