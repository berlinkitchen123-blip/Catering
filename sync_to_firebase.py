#!/Users/patelharsh513/Desktop/bb-kitchen/venv/bin/python3
"""
Bella & Bona — Firebase Sync
Reads Catering_BOMBEOGenerator_Macros_1.xlsb and pushes all data to Firestore.

Usage:
    python3 sync_to_firebase.py <path-to-xlsb>

Example:
    python3 sync_to_firebase.py ~/Desktop/Catering_BOMBEOGenerator_Macros_1.xlsb

Firestore structure written:
    /kitchen/meta       — week label, day names, updated timestamp
    /kitchen/events     — all events with menu + equipment
    /kitchen/recipes    — all recipe details (ingredients, steps, cooking, logistics)
    /kitchen/bom        — weekly ingredient totals
"""

import sys, json, os, re, requests
from pyxlsb import open_workbook
from datetime import datetime, timedelta
from collections import OrderedDict

# ─── Firebase config ──────────────────────────────────────────────────────────

PROJECT_ID = "stock-count-d0abf"
API_KEY    = "AIzaSyAAli1-_46ehgM7hzaLLNybkP4MIUw0XZQ"
FS_BASE    = f"https://firestore.googleapis.com/v1/projects/{PROJECT_ID}/databases/(default)/documents"

# ─── Firestore helpers ────────────────────────────────────────────────────────

def to_fs(val):
    """Convert Python value to Firestore REST API format."""
    if val is None:              return {"nullValue": None}
    if isinstance(val, bool):    return {"booleanValue": val}
    if isinstance(val, int):     return {"integerValue": str(val)}
    if isinstance(val, float):   return {"doubleValue": val}
    if isinstance(val, str):     return {"stringValue": val}
    if isinstance(val, list):    return {"arrayValue": {"values": [to_fs(v) for v in val]}}
    if isinstance(val, dict):    return {"mapValue": {"fields": {k: to_fs(v) for k, v in val.items()}}}
    return {"stringValue": str(val)}

def write_doc(path, data):
    """PATCH a Firestore document. Creates or overwrites."""
    fields = {k: to_fs(v) for k, v in data.items()}
    url = f"{FS_BASE}/{path}?key={API_KEY}"
    r = requests.patch(url, json={"fields": fields}, timeout=30)
    if r.status_code not in (200, 201):
        raise RuntimeError(f"Firestore write failed [{r.status_code}] {path}: {r.text[:300]}")
    return r.json()

# ─── Data helpers ─────────────────────────────────────────────────────────────

def cv(cell):
    if cell is None: return ''
    v = cell.v
    return '' if v is None else v

def serial_to_dt(serial):
    try:
        return datetime(1899, 12, 30) + timedelta(days=float(serial))
    except:
        return None

def fmt_num(v):
    if v == '': return ''
    try:
        f = float(v)
        if f == int(f): return str(int(f))
        return f'{f:.3g}'
    except:
        return str(v)

def slug(s):
    return re.sub(r'[^a-z0-9]+', '-', s.lower()).strip('-')

# ─── Parse xlsb ───────────────────────────────────────────────────────────────

def parse_beos(wb):
    with wb.get_sheet('BEOs Extract - Current Week') as ws:
        rows = list(ws.rows())

    r0 = [cv(c) for c in rows[0]]
    week_start = serial_to_dt(r0[1])
    week_end   = serial_to_dt(r0[3])

    events  = OrderedDict()
    recipes = {}

    for i, row in enumerate(rows):
        if i < 5: continue
        vals = [cv(c) for c in row]
        if not vals[0]: continue
        d = serial_to_dt(vals[0])
        if not d: continue

        date_str = d.strftime('%A %d/%m/%Y')
        client   = str(vals[1]).strip()
        recipe   = str(vals[2]).strip()
        if not recipe: continue

        if date_str not in events: events[date_str] = OrderedDict()
        if client not in events[date_str]: events[date_str][client] = OrderedDict()
        if recipe not in events[date_str][client]:
            events[date_str][client][recipe] = {
                'portions': vals[3], 'diet': str(vals[14]).strip(),
                'allergens': str(vals[15]).strip(), 'category': str(vals[16]).strip(),
                'chafing': vals[12], 'packaging': str(vals[13]).strip(),
                'equipment': str(vals[27]).strip(),
            }

        rec = events[date_str][client][recipe]
        if vals[4]:
            qty_c = fmt_num(vals[5]); unit = str(vals[7]).strip()
            tot_c = fmt_num(vals[8])
            rec.setdefault('ingredients', []).append({
                'name':  str(vals[4]).strip(),
                'qty':   f'{qty_c} {unit}'.strip(),
                'total': f'{tot_c} {unit}'.strip(),
                'note':  str(vals[10]).strip(),
            })
        if vals[11]:
            instr = str(vals[11]).strip()
            if instr and instr not in rec.setdefault('instructions', []):
                rec['instructions'].append(instr)
        if vals[17]:
            rec.setdefault('cooking_stages', []).append({
                'stage': str(vals[17]).strip(), 'mode': str(vals[18]).strip(),
                'temp': str(vals[19]).strip(), 'time': str(vals[20]).strip(),
                'humidity': fmt_num(vals[21]), 'notes': str(vals[22]).strip(),
            })
        if vals[23] and not rec.get('transport'): rec['transport'] = str(vals[23]).strip()
        if vals[24] and not rec.get('rational'):  rec['rational']  = str(vals[24]).strip()
        if vals[25] and not rec.get('buffet'):    rec['buffet']    = str(vals[25]).strip()
        if vals[26] and not rec.get('regen'):     rec['regen']     = str(vals[26]).strip()

        if recipe not in recipes:
            recipes[recipe] = {
                'diet': str(vals[14]).strip(), 'allergens': str(vals[15]).strip(),
                'category': str(vals[16]).strip(),
                'ingredients': [], 'instructions': [], 'cooking_stages': [],
                'transport': '', 'rational': '', 'buffet': '', 'regen': '',
            }

    # merge per-recipe data into recipes dict
    for date_str, clients in events.items():
        for client, recs in clients.items():
            for rname, rdata in recs.items():
                if rname in recipes:
                    for field in ['ingredients','instructions','cooking_stages','transport','rational','buffet','regen']:
                        if rdata.get(field) and not recipes[rname].get(field):
                            recipes[rname][field] = rdata[field]

    return week_start, week_end, events, recipes


def parse_bom(wb):
    bom = []
    with wb.get_sheet('BOM - Current Week') as ws:
        for i, row in enumerate(ws.rows()):
            if i < 4: continue
            vals = [cv(c) for c in row]
            if not vals[0]: continue
            name = str(vals[0]).strip()
            if name == 'Ingredient': continue
            bom.append({
                'name': name,
                'unit': str(vals[1]).strip() if len(vals) > 1 else '',
                'qty':  fmt_num(vals[2]) if len(vals) > 2 else '',
            })
    return bom


def build_events_list(events):
    day_names = list(events.keys())
    ev_list = []
    for di, (date_str, clients) in enumerate(events.items()):
        for client, recs in clients.items():
            equip_set = set()
            menu = []
            for rname, rdata in recs.items():
                portions = rdata.get('portions','')
                try: portions = int(float(portions))
                except: portions = 0
                menu.append({
                    'name':      rname,
                    'portions':  portions,
                    'diet':      rdata.get('diet',''),
                    'allergens': rdata.get('allergens',''),
                    'category':  rdata.get('category',''),
                    'packaging': rdata.get('packaging',''),
                    'chafing':   fmt_num(rdata.get('chafing','')),
                })
                if rdata.get('equipment'):
                    for e in rdata['equipment'].split(','):
                        equip_set.add(e.strip())
            ev_list.append({
                'id':     slug(f'{date_str[:3]}-{client}'),
                'day':    di,
                'date':   date_str,
                'client': client,
                'menu':   menu,
                'equip':  sorted(equip_set),
            })
    return day_names, ev_list

# ─── Main sync ────────────────────────────────────────────────────────────────

def sync(xlsb_path):
    print(f"📖 Reading: {xlsb_path}")
    with open_workbook(xlsb_path) as wb:
        week_start, week_end, events, recipes = parse_beos(wb, 'BEOs Extract - Next Week')
        bom = parse_bom(wb, 'BOM - Next Week')

    day_names, ev_list = build_events_list(events)
    ws_str = week_start.strftime('%-d %b') if week_start else ''
    we_str = week_end.strftime('%-d %b %Y') if week_end else ''
    week_label = f"{ws_str}–{we_str}"
    updated_at = datetime.utcnow().isoformat() + 'Z'

    print(f"   Week: {week_label} | {len(ev_list)} events | {len(recipes)} recipes | {len(bom)} BOM items")
    print(f"🔥 Syncing to Firebase ({PROJECT_ID})...")

    # 1. Meta
    write_doc('kitchen/meta', {
        'weekLabel': week_label,
        'weekStart': ws_str,
        'weekEnd':   we_str,
        'dayNames':  day_names,
        'updatedAt': updated_at,
    })
    print("   ✓ kitchen/meta")

    # 2. Events
    write_doc('kitchen/events', {'data': ev_list, 'updatedAt': updated_at})
    print("   ✓ kitchen/events")

    # 3. Recipes (store as list to avoid Firestore key restrictions)
    recipe_list = [{'name': k, **v} for k, v in recipes.items()]
    write_doc('kitchen/recipes', {'data': recipe_list, 'updatedAt': updated_at})
    print("   ✓ kitchen/recipes")

    # 4. BOM
    write_doc('kitchen/bom', {'items': bom, 'updatedAt': updated_at})
    print("   ✓ kitchen/bom")

    print(f"✅ Sync complete — {updated_at}")
    return week_label

if __name__ == '__main__':
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(1)
    sync(os.path.expanduser(sys.argv[1]))
