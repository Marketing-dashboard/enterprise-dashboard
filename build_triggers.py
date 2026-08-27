"""
build_triggers.py — Run this script whenever GSheet trigger data is updated.
Fetches Triggers_raw from GSheet, aggregates, and embeds into dashboard index.html.
Then commit + push to deploy.

Usage:
    python build_triggers.py
"""
import sys, os, json, csv, io, urllib.request, re, openpyxl
sys.stdout.reconfigure(encoding='utf-8')

SPREADSHEET_ID = '1Pbckm2WZ3b0NGMiZ7QeWOuPRHPnUso1EVnETud6t1aw'
TRIG_GID   = '942825587'   # Triggers_raw tab
ENT_GID    = '1372441356'  # Enterprise_Raw tab
EXCEL_PATH = r'C:\Users\Deepanshi Ahuja\Desktop\Enterprise-dashboard-sold_cpl.xlsx'
DASH_PATH  = os.path.join(os.path.dirname(__file__), 'margin', 'index.html')

def gsheet_csv(gid):
    url = f'https://docs.google.com/spreadsheets/d/{SPREADSHEET_ID}/export?format=csv&gid={gid}'
    print(f'  Fetching gid={gid} ...', end=' ', flush=True)
    data = urllib.request.urlopen(url, timeout=120).read().decode('utf-8-sig')
    rows = list(csv.DictReader(io.StringIO(data)))
    print(f'{len(rows):,} rows')
    return rows

def src_to_channel(src):
    s = (src or '').strip().lower()
    if s in ('ms_fb','ms-fb','meta','facebook','instagram'): return 'MS-FB'
    if s in ('ctwa','ctwa (meta)','whatsapp','whatsapp_mkt','rampwin_meta_whatsapp','rampwin meta whatsapp'): return 'CTWA'
    if s == 'adwords_ims': return 'Adwords_IMS'
    if s in ('adwords','google','pmax','demand_gen','uac','display','youtube','video'): return 'Adwords'
    return None

# ── 1. Fetch Enterprise_Raw to get valid brand/model set ──
print('Fetching Enterprise_Raw...')
ent_rows = gsheet_csv(ENT_GID)
spend_brands = set()
for r in ent_rows:
    b = (r.get('Brand') or '').strip().lower()
    m = (r.get('Model') or '').strip().lower()
    if b and m:
        spend_brands.add((b, m))
print(f'  Enterprise brand/models: {len(spend_brands)}')

# ── 2. Fetch Triggers_raw ──
print('Fetching Triggers_raw...')
trig_rows = gsheet_csv(TRIG_GID)

# ── 3. Aggregate triggers ──
print('Aggregating triggers...')
trig_agg = {}
skipped_channel = set()
total = 0

for r in trig_rows:
    brand_mapped = (r.get('Brand_Mapped') or '').strip()
    model_mapped = (r.get('Model_Mapped') or '').strip()
    if not brand_mapped or not model_mapped:
        continue

    fb_low = brand_mapped.lower()
    is_jlr     = 'jlr' in fb_low or 'land rover' in fb_low or 'jaguar' in fb_low
    is_lexus   = 'lexus' in fb_low
    is_citroen = 'citroen' in fb_low
    use_list_id = is_jlr or is_lexus or is_citroen

    triggered    = float(r.get('Triggered') or 0 or 0)
    list_id_trig = float(r.get('Triggered in List_ID') or 0 or 0)

    # Filter: others need Triggered>0; special brands also accept List_ID>0
    if not triggered and not (use_list_id and list_id_trig):
        continue

    # Effective trigger count
    if is_citroen:
        effective = triggered + list_id_trig * 0.20
    elif use_list_id:
        effective = triggered + list_id_trig
    else:
        effective = triggered

    if not effective:
        continue

    # Channel
    src = (r.get('final_source') or r.get('Campaign Type') or '').strip()
    channel = src_to_channel(src)
    if channel is None:
        skipped_channel.add(src)
        continue

    # Date
    date_raw = (r.get('Date') or '').strip()
    if not date_raw:
        continue
    # Normalize date to YYYY-MM-DD
    try:
        from datetime import datetime
        if '-' in date_raw and len(date_raw) == 10:
            date_str = date_raw
        else:
            date_str = datetime.strptime(date_raw, '%d-%b-%Y').strftime('%Y-%m-%d')
    except:
        date_str = date_raw

    key = f"{date_str}|{brand_mapped.lower()}|{model_mapped.lower()}|{channel.lower()}"
    trig_agg[key] = trig_agg.get(key, 0) + effective
    total += effective

print(f'  Aggregated: {len(trig_agg):,} keys, {total:,.0f} total triggers')
print(f'  Skipped channels: {skipped_channel}')

# ── 4. Build month-wise rates from Excel ──
print('Building rates from Excel...')
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
        key = f'{str(brand).strip().lower()}||{str(model).strip().lower()}||{ch}'
        if val_pct is not None:
            try:
                v2 = float(str(val_pct).replace('%','').strip())
                if 0 < v2 <= 1.0: v2 *= 100
                validation[key] = round(v2, 4)
            except: pass
        if s_cpl is not None:
            try: sold_cpl[key] = float(str(s_cpl).replace(',','').strip())
            except: pass
    norm = sheet_name.replace('June','Jun').replace('July','Jul')
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
