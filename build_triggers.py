"""
build_triggers.py — Run this script whenever GSheet trigger data is updated.
Fetches Triggers_raw from GSheet, aggregates, and embeds into dashboard index.html.
Then commit + push to deploy.

Usage:
    python build_triggers.py
"""
import sys, os, json, csv, io, urllib.request, re, openpyxl
sys.stdout.reconfigure(encoding='utf-8')

SPREADSHEET_ID = '1Pbckm2WZ3b0NGMiZ7QeWOuPRHPnUso1EVnETud6t1aw'  # old sheet (Demand/Ref tabs)
ENT_SHEET_ID  = '1BCOWyyUTVmWkVIxKzMV28oZPhBWdlWAr_S1iEYVimJI'  # new spends sheet
ENT_GID_NEW   = '0'
TRIG_SHEET_ID = '127wAiu04aIeV8V7RicOwlutE5ALv0KyKwPBfhnZi1ZU'  # Nov'26 triggers sheet
TRIG_GID_NEW  = '0'
EXCEL_PATH = r'C:\Users\Deepanshi Ahuja\Desktop\Enterprise-dashboard-sold_cpl.xlsx'
DASH_PATH  = os.path.join(os.path.dirname(__file__), 'margin', 'index.html')

# Nov'26 daily demand (run rates) — update each month
DEMAND_OCT = {
    'Toyota Camry': 24,
    'Toyota Vellfire': 24,
    'Toyota Land Cruiser 300': 10,
    'Toyota Hilux': 23,
    'Toyota Urban Cruiser EBELLA': 19,
    'Toyota Hyryder': 302,
    'Toyota Taisor': 156,
    'Toyota Rumion': 3,
    'Volkswagen Virtus': 82,
    'Volkswagen Taigun': 152,
    'Volkswagen Tayron': 10,
    'Honda Elevate': 500,
    'Honda City': 130,
    'Honda Amaze': 280,
    'Honda CB350RS': 70,
    'Honda Hness CB350': 82,
    'Honda CB350': 61,
    'Citroen Aircross': 80,
    'Citroen C3': 40,
    'Citroen Basalt': 38,
    'Land Rover Range Rover Evoque': 100,
    'Land Rover Range Rover Velar': 200,
    'Land Rover Defender': 130,
    'Land Rover Range Rover': 40,
    'Land Rover Range Rover Sport': 90,
    'Tata Nexon': 572,
    'Tata Tiago': 233,
    'Tata Altroz': 253,
    'Tata Harrier': 200,
    'Tata Safari': 40,
    'Tata Punch': 1005,
    'Tata Curvv': 89,
    'Tata Sierra': 336,
    'Tata Nexon EV': 420,
    'Tata Curvv EV': 70,
    'Tata Harrier EV': 195,
    'Tata Sierra EV': 226,
    'Lexus NX': 90,
    'Lexus ES 350h': 12,
    'Lexus RX': 15,
    'Skoda Slavia': 322,
    'Skoda Kushaq': 516,
    'Skoda Kylaq': 405,
    'Skoda Kodiaq': 26,
    'Skoda Octavia RS': 13,
    'Mahindra XEV 9e': 64,
    'Mahindra BE 6 SPORTEQ': 124,
    'Mahindra Bolero Max Pik-Up City': 498,
    'Mahindra Bolero Pik-Up': 316,
    'Mahindra Veero': 82,
    'Audi Q7': 2,
    'Royal Enfield': 1200,
    'HMSI Bigwing': 210,
    'HMSI Redwing': 200,
}

# BU Types to include from new spends sheet
VALID_BU_TYPES = {'media sales-ga', 'media sales-fb', 'media sales-whatsapp'}  # excludes media sales-ga_ims

def gsheet_csv(sheet_id, gid):
    url = f'https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=csv&gid={gid}'
    print(f'  Fetching {sheet_id[:20]}... gid={gid} ...', end=' ', flush=True)
    data = urllib.request.urlopen(url, timeout=120).read().decode('utf-8-sig')
    rows = list(csv.DictReader(io.StringIO(data)))
    print(f'{len(rows):,} rows')
    return rows

def src_to_channel(src):
    s = (src or '').strip().lower()
    if s in ('ms_fb','ms-fb','meta','facebook','instagram','media sales-fb'): return 'MS-FB'
    if s in ('ctwa','ctwa (meta)','whatsapp','whatsapp_mkt','rampwin_meta_whatsapp','rampwin meta whatsapp','media sales-whatsapp','media sales-ctwa'): return 'CTWA'
    if s == 'adwords_ims': return 'Adwords_IMS'
    if s in ('adwords','google','pmax','demand_gen','uac','display','youtube','video','media sales-ga'): return 'Adwords'
    return None

# ── 1. Fetch new spends sheet to get valid brand/model set ──
print('Fetching spends sheet (Enterprise_Raw)...')
try:
    ent_rows_raw = gsheet_csv(ENT_SHEET_ID, ENT_GID_NEW)
    ent_rows = [r for r in ent_rows_raw
                if (r.get('Brand') or '').strip() not in ('', '#N/A', 'N/A')
                and (r.get('BU Type') or '').strip().lower() in VALID_BU_TYPES
                and (r.get('Month') or '').strip() == "Nov'26"]
    spend_brands = set()
    for r in ent_rows:
        b = (r.get('Brand') or '').strip().lower()
        m = (r.get('Model') or '').strip().lower()
        if b and m:
            spend_brands.add((b, m))
    print(f'  After filter: {len(ent_rows):,} rows, {len(spend_brands)} brand/models')
except Exception as e:
    print(f'  WARNING: Could not fetch spends sheet ({e}), continuing without it...')
    ent_rows, spend_brands = [], set()

# ── 2. Fetch new triggers sheet ──
print('Fetching Triggers sheet...')
trig_rows = gsheet_csv(TRIG_SHEET_ID, TRIG_GID_NEW)

# ── 3. Aggregate triggers + Unique Gen ──
print('Aggregating triggers...')
trig_agg = {}
ug_agg   = {}   # Unique Gen aggregated by same key
ct_trig_agg = {}  # utm_campaign → trigger count
skipped_channel = set()
total = 0

TRIG_BRAND_NORM = {
    ('honda', 'honda shine'):     'hmsi redwing',
    ('honda', 'honda shine 100'): 'hmsi redwing',
}
TRIG_MODEL_NORM_STATIC = {
    'mahindra pv': {
        'xev 9s':            'mahindra xev 9s',
        'xev 9e':            'mahindra xev 9e',
        'xuv 7xo':           'mahindra xuv 7xo',
        'xuv 3xo':           'mahindra xuv 3xo',
        'bolero neo':        'mahindra bolero neo',
        'mahindra bolero':   'mahindra bolero',
        'mahindra scorpion': 'mahindra scorpio n',
    },
    'simple energy': {'simple energy': 'simple one'},
    'toyota':        {'toyota ebella': 'toyota urban cruiser ebella'},
    'vinfast':       {'vinfast vf6': 'vf6', 'vinfast vf7': 'vf7'},
}

for r in trig_rows:
    brand_mapped = (r.get('Brand_Mapped') or '').strip()
    model_mapped = (r.get('Model_Mapped') or '').strip()
    if not brand_mapped or not model_mapped:
        continue

    # Channel (validate early so we can build the key)
    src = (r.get('final_source') or r.get('Campaign Type') or '').strip()
    channel = src_to_channel(src)
    if channel is None:
        skipped_channel.add(src)
        continue

    # Date
    date_raw = (r.get('Date') or '').strip()
    if not date_raw:
        continue
    try:
        from datetime import datetime
        if '-' in date_raw and len(date_raw) == 10:
            date_str = date_raw
        else:
            date_str = datetime.strptime(date_raw, '%d-%b-%Y').strftime('%Y-%m-%d')
    except:
        date_str = date_raw
    if not date_str.startswith('2026-11'):
        continue

    # Brand/model normalization
    brand_l   = brand_mapped.lower()
    model_l_tmp = model_mapped.lower()
    brand_l = TRIG_BRAND_NORM.get((brand_l, model_l_tmp), brand_l)

    fb_low = brand_mapped.lower()
    is_jlr     = 'jlr' in fb_low or 'land rover' in fb_low or 'jaguar' in fb_low
    is_lexus   = 'lexus' in fb_low
    is_citroen = 'citroen' in fb_low
    is_river   = 'river' in fb_low
    use_list_id = is_jlr or is_lexus or is_citroen or is_river
    # Normalize model names
    model_l = model_mapped.lower()
    if brand_l in TRIG_MODEL_NORM_STATIC:
        model_l = TRIG_MODEL_NORM_STATIC[brand_l].get(model_l, model_l)
    key = f"{date_str}|{brand_l}|{model_l}|{channel.lower()}"

    # Always aggregate Unique Gen (regardless of triggered count)
    unique_gen = float(r.get('Unique Gen') or 0 or 0)
    if unique_gen > 0:
        ug_agg[key] = ug_agg.get(key, 0) + unique_gen

    # Trigger aggregation (only when triggered)
    triggered    = float(r.get('Triggered') or 0 or 0)
    list_id_trig = float(r.get('Triggered in List_ID') or 0 or 0)
    if not triggered and not (use_list_id and list_id_trig):
        continue
    effective = (triggered + list_id_trig) if use_list_id else triggered
    if not effective:
        continue
    trig_agg[key] = trig_agg.get(key, 0) + effective
    total += effective

    # Campaign-name level aggregation (for Google Campaigns tab exact matching)
    utm = (r.get('utm_campaign') or '').strip()
    if utm:
        ct_key = f'{utm}'
        ct_trig_agg[ct_key] = ct_trig_agg.get(ct_key, 0) + effective

print(f'  Aggregated: {len(trig_agg):,} trig keys, {total:,.0f} total triggers, {len(ug_agg):,} unique gen keys')
print(f'  Skipped channels: {skipped_channel}')

# ── 4. Build month-wise rates from Excel ──
print('Building rates from Excel...')
# Model aliases: Excel name → canonical GSheet name (lowercase, hyphens already removed)
MODEL_ALIASES = {
    'honda cb125 hornet': 'honda cb125',
    'ola': 'ola electric',                      # Excel uses 'Ola'; GSheet uses 'Ola Electric'
    'harley-davidson x440 t': 'harley davidson x440 t',  # hyphen vs space variant
    'mercedes-benz c-class': 'mercedes benz c class',
    'mercedes-benz gla': 'mercedes benz gla',
    'mercedes-benz glc': 'mercedes benz glc',
    'mercedes-benz gle': 'mercedes benz gle',
    'mercedes-benz v-class': 'mercedes benz v class',
    'mercedes-benz a-class': 'mercedes benz a class',
}
CHANNEL_MAP = {
    'media sales-fb':'MS-FB','ms-fb':'MS-FB','ms_fb':'MS-FB',
    'media sales-ga':'Adwords','adwords':'Adwords',
    'adwords_ims':'Adwords_IMS',
    'media sales-whatsapp':'CTWA','media sales-ctwa':'CTWA','ctwa':'CTWA',
}
rates_by_month = {}
wb = openpyxl.load_workbook(EXCEL_PATH, data_only=True)
for sheet_name in wb.sheetnames:
    ws = wb[sheet_name]
    hdr = {}
    for c in range(1, 20):
        v = ws.cell(1, c).value
        if v: hdr[str(v).strip()] = c
    sold_cpl, validation = {}, {}
    for row in range(2, ws.max_row + 1):
        brand   = ws.cell(row, hdr.get('Brand', 1)).value
        model   = ws.cell(row, hdr.get('Model', 2)).value
        channel = ws.cell(row, hdr.get('Channel', 4)).value
        val_pct = ws.cell(row, hdr.get('Validation%', 5)).value
        s_cpl   = ws.cell(row, hdr.get('Sold CPL', 6)).value
        if not brand or not model or not channel: continue
        ch  = CHANNEL_MAP.get(str(channel).strip().lower(), str(channel).strip())
        # Normalize hyphens→spaces and apply aliases so key matches GSheet model names
        norm_model = str(model).strip().lower().replace('-', ' ')
        norm_model = MODEL_ALIASES.get(norm_model, norm_model)
        key = f'{str(brand).strip().lower()}||{norm_model}||{ch}'
        if val_pct is not None:
            try:
                v2 = float(str(val_pct).replace('%','').strip())
                if 0 < v2 <= 1.0: v2 *= 100
                validation[key] = round(v2, 4)
            except: pass
        if s_cpl is not None:
            try: sold_cpl[key] = float(str(s_cpl).replace('₹','').replace(',','').strip())
            except: pass
    norm = sheet_name.replace('June','Jun').replace('July','Jul').replace('Sept','Sep')
    rates_by_month[norm] = {'soldCPL': sold_cpl, 'validation': validation}
    print(f'  {sheet_name}: {len(sold_cpl)} CPL, {len(validation)} validation rates')

# ── 5. Embed into dashboard ──
print('Embedding into dashboard...')
content = open(DASH_PATH, encoding='utf-8').read()

# Replace _localTrigData
TRIG_JSON = json.dumps(trig_agg, ensure_ascii=False)
embed_trig = f'const _localTrigData={TRIG_JSON};'
existing = content.find('const _localTrigData=')
if existing >= 0:
    end = content.find(';', existing) + 1
    if content[end:end+1] == '\n': end += 1
    content = content[:existing] + embed_trig + '\n' + content[end:]
    print(f'  _localTrigData replaced ({len(TRIG_JSON)//1024}KB, {len(trig_agg)} keys)')
else:
    insert = content.find('\nloadRates().then(loadData)')
    content = content[:insert] + '\n' + embed_trig + content[insert:]
    print(f'  _localTrigData inserted ({len(TRIG_JSON)//1024}KB, {len(trig_agg)} keys)')

# Replace _localUniqueGenData
UG_JSON = json.dumps(ug_agg, ensure_ascii=False)
embed_ug = f'const _localUniqueGenData={UG_JSON};'
existing_ug = content.find('const _localUniqueGenData=')
if existing_ug >= 0:
    end_ug = content.find(';', existing_ug) + 1
    if content[end_ug:end_ug+1] == '\n': end_ug += 1
    content = content[:existing_ug] + embed_ug + '\n' + content[end_ug:]
    print(f'  _localUniqueGenData replaced ({len(UG_JSON)//1024}KB, {len(ug_agg)} keys)')
else:
    insert = content.find('\nloadRates().then(loadData)')
    content = content[:insert] + '\n' + embed_ug + content[insert:]
    print(f'  _localUniqueGenData inserted ({len(UG_JSON)//1024}KB, {len(ug_agg)} keys)')

# Replace _localCampTrigData (utm_campaign → trigger count)
CT_JSON = json.dumps(ct_trig_agg, ensure_ascii=False)
embed_ct = f'const _localCampTrigData={CT_JSON};'
existing_ct = content.find('const _localCampTrigData=')
if existing_ct >= 0:
    end_ct = content.find(';', existing_ct) + 1
    if content[end_ct:end_ct+1] == '\n': end_ct += 1
    content = content[:existing_ct] + embed_ct + '\n' + content[end_ct:]
    print(f'  _localCampTrigData replaced ({len(CT_JSON)//1024}KB, {len(ct_trig_agg)} campaigns)')
else:
    insert = content.find('\nloadRates().then(loadData)')
    content = content[:insert] + '\n' + embed_ct + content[insert:]
    print(f'  _localCampTrigData inserted ({len(CT_JSON)//1024}KB, {len(ct_trig_agg)} campaigns)')

# Replace _localHistData (from Historic_data_enterprise.xlsx)
HIST_XLSX = r'C:\Users\Deepanshi Ahuja\Desktop\Historic_data_enterprise.xlsx'
if os.path.exists(HIST_XLSX):
    import pandas as _pd
    _hdf = _pd.read_excel(HIST_XLSX).fillna(0)
    _hdf['Brand'] = _hdf['Brand'].str.strip().replace({'citroen': 'Citroen', 'TATA EV': 'Tata EV'})
    _hdf['Model'] = _hdf['Model'].str.strip()
    hist_agg = {}
    for _, _r in _hdf.iterrows():
        _key = f"{_r['Month']}|{_r['Brand']}|{_r['Model']}|{_r['Channel']}"
        _pvl = float(_r['Triggered MS']) * float(_r['Validation%'])
        _rev = _pvl * float(_r['Sold CPL'])
        if _key not in hist_agg:
            hist_agg[_key] = {'seg': _r['Segment'], 's': 0, 'l': 0, 'ug': 0, 't': 0, 'pvl': 0, 'rev': 0, 'sc': float(_r['Sold CPL']), 'vp': float(_r['Validation%'])}
        _h = hist_agg[_key]
        _h['s'] += float(_r['Spends']); _h['l'] += float(_r['Leads'])
        _h['ug'] += float(_r['Unique Gen']); _h['t'] += float(_r['Triggered MS'])
        _h['pvl'] += _pvl; _h['rev'] += _rev
        _h['sc'] = float(_r['Sold CPL']); _h['vp'] = float(_r['Validation%'])
    for _v in hist_agg.values():
        _v['s'] = round(_v['s']); _v['l'] = round(_v['l'], 2); _v['ug'] = round(_v['ug'], 2)
        _v['t'] = round(_v['t'], 2); _v['pvl'] = round(_v['pvl'], 2); _v['rev'] = round(_v['rev'], 2)
        _v['sc'] = round(_v['sc'], 2); _v['vp'] = round(_v['vp'], 4)
    HIST_JSON = json.dumps(hist_agg, ensure_ascii=False)
    embed_hist = f'const _localHistData={HIST_JSON};'
    existing_hist = content.find('const _localHistData=')
    if existing_hist >= 0:
        end_hist = content.find(';', existing_hist) + 1
        if content[end_hist:end_hist+1] == '\n': end_hist += 1
        content = content[:existing_hist] + embed_hist + '\n' + content[end_hist:]
        print(f'  _localHistData replaced ({len(HIST_JSON)//1024}KB, {len(hist_agg)} keys)')
    else:
        insert = content.find('\nloadRates().then(loadData)')
        content = content[:insert] + '\n' + embed_hist + content[insert:]
        print(f'  _localHistData inserted ({len(HIST_JSON)//1024}KB, {len(hist_agg)} keys)')
else:
    print('  _localHistData skipped (Excel not found)')

# Replace _localDemandData
DEMAND_JSON = json.dumps(DEMAND_OCT, ensure_ascii=False)
embed_demand = f'const _localDemandData={DEMAND_JSON};'
existing_d = content.find('const _localDemandData=')
if existing_d >= 0:
    end_d = content.find(';', existing_d) + 1
    if content[end_d:end_d+1] == '\n': end_d += 1
    content = content[:existing_d] + embed_demand + '\n' + content[end_d:]
    print(f'  _localDemandData replaced ({len(DEMAND_OCT)} models)')
else:
    insert = content.find('\nloadRates().then(loadData)')
    content = content[:insert] + '\n' + embed_demand + content[insert:]
    print(f'  _localDemandData inserted ({len(DEMAND_OCT)} models)')

# Replace loadRates — always a single line, so replace line-by-line
new_fn = 'async function loadRates(){_rates={byMonth:' + json.dumps(rates_by_month, ensure_ascii=False) + '}}'
lines = content.split('\n')
replaced = False
for idx, line in enumerate(lines):
    if line.startswith('async function loadRates()'):
        lines[idx] = new_fn
        replaced = True
        break
if replaced:
    content = '\n'.join(lines)
    print('  loadRates updated')
else:
    # Insert before loadRates().then(loadData)
    insert = content.find('\nloadRates().then(loadData)')
    content = content[:insert] + '\n' + new_fn + content[insert:]
    print('  loadRates inserted')

open(DASH_PATH, 'w', encoding='utf-8').write(content)
print('\nDone! Now run:')
print('  git add margin/index.html')
print('  git commit -m "Refresh trigger data"')
print('  git push origin main')
