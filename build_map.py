#!/usr/bin/env python3
# ============================================================================
# build_map.py  —  MyBuddy Pet Finder (V92 - Desktop Call Interceptor)
# ============================================================================
import json, re, time, urllib.request, urllib.parse, random, csv
from pathlib import Path
from openpyxl import load_workbook

WORKBOOK_VETS = "MyBuddy_Vet_Finder_MASTER_RECONCILED_2026-10-15_v9.xlsx"
WORKBOOK_PLACES = "PET_PLACES.xlsx"
WORKBOOK_REST = "Pets_places_rest_updated.xlsx"
PRICING_FILE = "MyBuddy_Clinic_Price_List_Register_2026-11-04.xlsx"
REVIEWS_CACHE_FILE = "reviews_cache.json"

OUTPUT   = "index.html"
CACHE_FILE = Path(__file__).parent / "geocache.json"
MANIFEST_OUTPUT = "manifest.json"
ICON_OUTPUT     = "icon.svg"

# PASTE YOUR API KEY HERE IF YOU WANT TO UPDATE REVIEWS
GOOGLE_API_KEY = "PASTE_YOUR_NEW_API_KEY_HERE"

for old in ["mybuddy-vet-map.html", "mybuddy-vet-map-v2.html", "mybuddy-vet-map-v3.html"]:
    old_file = Path(__file__).parent / old
    if old_file.exists(): old_file.unlink()

DAY_NAMES = ["Monday","Tuesday","Wednesday","Thursday","Friday","Saturday","Sunday"]
DAY_ALIASES = {"mon":"Monday","monday":"Monday","tue":"Tuesday","tues":"Tuesday","tuesday":"Tuesday","wed":"Wednesday","weds":"Wednesday","wednesday":"Wednesday","thu":"Thursday","thur":"Thursday","thurs":"Thursday","thursday":"Thursday","fri":"Friday","friday":"Friday","sat":"Saturday","saturday":"Saturday","sun":"Sunday","sunday":"Sunday"}

POSTAL_FALLBACKS = { "563338": "560338" }
EMERGENCY_OVERRIDES = { "MBV-0005", "MBV-0032" }

MANUAL_COORDS = {
    "MBV-0005": (1.3244, 103.9418), "MBV-0032": (1.3286, 103.8521),
    "PLC-105": (1.3622, 103.8475), "PLC-106": (1.2901, 103.7622),
    "PLC-107": (1.3008, 103.9124), "PLC-108": (1.2860, 103.8236),
    "PLC-109": (1.4647, 103.8336), "PLC-110": (1.3115, 103.7955),
    "PLC-111": (1.3506, 103.8721), "PLC-112": (1.3065, 103.7885),
    "PLC-113": (1.2644, 103.8223), "PLC-114": (1.3054, 103.8294),
    "PLC-115": (1.3313, 103.8502), "PLC-116": (1.4433, 103.8295),
    "PLC-117": (1.3245, 103.8412), "PLC-118": (1.3317, 103.8516),
    "PLC-119": (1.3311, 103.9235), "PLC-120": (1.3255, 103.9338),
    "PLC-121": (1.3067, 103.7629), "PLC-122": (1.3539, 103.9425),
    "PLC-123": (1.3177, 103.7955),
}

SPECIES_OVERRIDES = {}
SPECIES_ALIASES = {
    "dog":    ["dog", "canine"], "cat":    ["cat", "feline"],
    "bird":   ["bird", "avian", "parrot"],
    "exotic": ["exotic", "rabbit", "small", "reptile", "pocket", "hamster", "guinea"],
}

def to_bool(v):
    if isinstance(v, bool): return v
    if isinstance(v, str):  return v.strip().lower() in ("true","yes","1","y")
    return False

def determine_category(raw_cat):
    s = str(raw_cat or "").strip().lower()
    if not s: return "Other"
    
    if any(re.search(rf"\b{kw}\b", s) for kw in ["cafe", "coffee", "bakery", "brunch", "gelato", "dessert"]):
        return "Cafe"
    if any(re.search(rf"\b{kw}\b", s) for kw in ["asian", "chinese", "malay", "indian", "japanese", "korean", "thai", "vietnamese", "fusion"]):
        return "Asian"
    if any(re.search(rf"\b{kw}\b", s) for kw in ["western", "italian", "french", "steak", "bistro", "grill", "pub", "bar"]):
        return "Western"
    if any(re.search(rf"\b{kw}\b", s) for kw in ["restaurant", "dining", "eatery", "food", "kitchen", "dine"]):
        return "Dine" 

    if any(re.search(rf"\b{kw}\b", s) for kw in ["groom", "groomer", "salon", "spa"]):
        return "Groomer"
    if any(re.search(rf"\b{kw}\b", s) for kw in ["mall", "shopping"]):
        return "Mall"
    if any(re.search(rf"\b{kw}\b", s) for kw in ["train", "school", "obedience", "agility", "behavior"]):
        return "Training"
    if ("park" in s and "parking" not in s) or "dog run" in s:
        return "Park"
    if any(re.search(rf"\b{kw}\b", s) for kw in ["food", "supplies", "shop", "store", "mart"]):
        return "Pet Food"
        
    return raw_cat.strip().capitalize()

def clean_postal(raw, name=""):
    if raw is None: return ""
    s = str(raw).strip()
    if s.endswith(".0"): s = s[:-2]
    digits = re.sub(r"\D", "", s)
    if len(digits) == 6:
        return digits
    if len(digits) == 5:
        padded = "0" + digits
        if name: print(f"INFO: {name} — 5-digit postal '{digits}' padded to '{padded}'")
        return padded
    if digits and name:
        print(f"WARNING: {name} — invalid postal '{raw}' ({len(digits)} digits) — dropping")
    return ""

def safe_http_url(u):
    if not u: return ""
    s = str(u).strip()
    return s if s.lower().startswith(("http://","https://","tel:","mailto:")) else ""

def parse_time(s):
    s = str(s).strip().lower().replace(".", "").replace(" ", "")
    if not s: return None
    m = re.match(r"^(\d{1,2})(?::(\d{2}))?\s*(am|pm)?$", s)
    if not m: return None
    h = int(m.group(1)); mi = int(m.group(2) or 0); ap = m.group(3)
    if mi > 59: return None
    if ap == "am" and h == 12: h = 0
    elif ap == "pm" and h != 12: h += 12
    elif not ap and h > 23: return None
    return h * 60 + mi

def parse_intervals(text):
    if text is None: return None
    t = str(text).strip().lower()
    if not t: return None
    if re.search(r"\bby\s+appointment\b", t): return "appointment"
    if re.match(r"^closed\b", t): return []
    if re.search(r"\b(open\s+)?24\s*hours?\b", t) or "24/7" in t: return [(0, 1440)]
    ranges = re.findall(r"(\d{1,2}(?::\d{2})?\s*(?:am|pm)?)\s*(?:[\u2013\u2014\-]|to)\s*(\d{1,2}(?::\d{2})?\s*(?:am|pm)?)", t, re.I)
    if not ranges: return [] if re.search(r"\bclosed\b", t) else None
    intervals = []
    for a, b in ranges:
        sa = parse_time(a); sb = parse_time(b)
        if sa is None or sb is None: continue
        if sb <= sa: sb += 1440
        intervals.append((sa, sb))
    return intervals if intervals else None

def day_in_range(start, end, target):
    s = DAY_ALIASES.get(start.lower()); e = DAY_ALIASES.get(end.lower())
    if not s or not e: return False
    si = DAY_NAMES.index(s); ei = DAY_NAMES.index(e); cur = si
    while True:
        if DAY_NAMES[cur] == target: return True
        if cur == ei: break
        cur = (cur + 1) % 7
    return False

def extract_day_schedule(full_text, target_day):
    if not full_text: return ""
    chunks = re.split(r"[;\n]+", str(full_text))
    result = []; current_days = None
    for chunk in chunks:
        chunk = chunk.strip()
        if not chunk: continue
        dr = re.match(r"^(mon|monday|tue|tues|tuesday|wed|weds|wednesday|thu|thur|thurs|thursday|fri|friday|sat|saturday|sun|sunday)\s*[\u2013\u2014\-]\s*(mon|monday|tue|tues|tuesday|wed|weds|wednesday|thu|thur|thurs|thursday|fri|friday|sat|saturday|sun|sunday)\b", chunk, re.I)
        if dr:
            current_days = [d for d in DAY_NAMES if day_in_range(dr.group(1), dr.group(2), d)]
            if target_day in current_days:
                rest = chunk[dr.end():].strip()
                if rest: result.append(rest)
            continue
        ds = re.match(r"^(mon|monday|tue|tues|tuesday|wed|weds|wednesday|thu|thur|thurs|thursday|fri|friday|sat|saturday|sun|sunday)\b", chunk, re.I)
        if ds:
            d = DAY_ALIASES.get(ds.group(1).lower())
            current_days = [d]
            if d == target_day:
                rest = chunk[ds.end():].strip()
                if rest: result.append(rest)
            continue
        if current_days is None or target_day in current_days: result.append(chunk)
    return "; ".join(result)

def read_sheet(wb, name_or_index):
    try:
        if isinstance(name_or_index, int): ws = wb.worksheets[name_or_index]
        else: ws = wb[name_or_index]
    except KeyError: return []
    rows = list(ws.iter_rows(values_only=True))
    if not rows: return []
    headers = [str(h).strip() if h is not None else f"col{i}" for i, h in enumerate(rows[0])]
    return [{headers[i]: (row[i] if i < len(row) else None) for i in range(len(headers))} for row in rows[1:] if not all(c is None for c in row)]

def _species_from_text(s):
    s = str(s or "").strip().lower()
    if not s: return set()
    return {tag for tag, kws in SPECIES_ALIASES.items() if any(kw in s for kw in kws)}

def extract_species_tags(services_data):
    tags = {}
    if not services_data: return tags
    headers = list(services_data[0].keys())
    has_species_col = any(str(h).strip().lower() == "species" for h in headers)
    wide_cols = [(h, _species_from_text(h)) for h in headers if _species_from_text(h)]
    for r in services_data:
        mbv = str(r.get("MyBuddy Vet ID") or "").strip()
        if not mbv: continue
        if has_species_col:
            sp_set = _species_from_text(r.get("Species"))
            if sp_set: tags.setdefault(mbv, set()).update(sp_set)
        else:
            for col, sp_set in wide_cols:
                if to_bool(r.get(col)): tags.setdefault(mbv, set()).update(sp_set)
    return {k: sorted(v) for k, v in tags.items()}

def _query_onemap(q):
    url = "https://www.onemap.gov.sg/api/common/elastic/search?" + urllib.parse.urlencode({
        "searchVal": q, "returnGeom": "Y", "getAddrDetails": "Y", "pageNum": 1})
    try:
        req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
        with urllib.request.urlopen(req, timeout=5) as r:
            data = json.loads(r.read().decode("utf-8"))
        for p in (data.get("results") or []):
            try:
                lat = float(p["LATITUDE"]); lng = float(p["LONGITUDE"])
                if 1.1 <= lat <= 1.5 and 103.5 <= lng <= 104.2: return (lat, lng)
            except (KeyError, ValueError): continue
    except Exception: pass
    return None

def geocode_location(postal, address, name, cache):
    key = f"{postal}|{address}|{name}"
    if key in cache: return cache[key]
    queries = []
    if postal and len(str(postal).strip()) == 6: queries.append(str(postal).strip())
    if name: queries.append(f"{name} Singapore")
    if address:
        clean_addr = str(address).replace('\n', ' ').strip()
        queries.append(f"{clean_addr} Singapore")
    for q in queries:
        time.sleep(0.2)
        coords = _query_onemap(q)
        if coords:
            cache[key] = coords
            return coords
    return None

def fetch_google_data(name, address, postal):
    if GOOGLE_API_KEY == "PASTE_YOUR_NEW_API_KEY_HERE": return None, None, None, None, None
    url = "https://places.googleapis.com/v1/places:searchText"
    headers = {
        "Content-Type": "application/json",
        "X-Goog-Api-Key": GOOGLE_API_KEY,
        "X-Goog-FieldMask": "places.displayName,places.location,places.rating,places.userRatingCount,places.id"
    }
    queries = [f"{name} {postal} Singapore", f"{name} Singapore", f"{postal} Singapore", f"{address} Singapore"]
    for q in queries:
        if not q.strip() or q.strip() == "Singapore": continue
        payload = {"textQuery": q}
        try:
            req = urllib.request.Request(url, data=json.dumps(payload).encode("utf-8"), headers=headers, method="POST")
            with urllib.request.urlopen(req, timeout=5) as r:
                data = json.loads(r.read().decode("utf-8"))
                if data.get("places"):
                    first = data["places"][0]
                    loc = first.get("location", {})
                    if loc.get("latitude") and loc.get("longitude"):
                        return (first.get("rating", 0), first.get("userRatingCount", 0),
                                loc.get("latitude"), loc.get("longitude"), first.get("id", ""))
        except Exception: pass
    return None, None, None, None, None

def write_pwa_files():
    manifest = {
        "name": "MyBuddy Pet Finder", "short_name": "Pet Finder",
        "start_url": "./index.html", "scope": "./",
        "display": "standalone", "orientation": "portrait",
        "background_color": "#ffffff", "theme_color": "#2C3E50",
        "icons": [{"src": "icon.svg", "sizes": "any", "type": "image/svg+xml", "purpose": "any"}]
    }
    Path(MANIFEST_OUTPUT).write_text(json.dumps(manifest, indent=2), encoding="utf-8")
    icon_svg = '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 512 512"><rect width="512" height="512" rx="96" fill="#2C3E50"/><circle cx="176" cy="180" r="44" fill="#ffffff"/><circle cx="256" cy="140" r="44" fill="#ffffff"/><circle cx="336" cy="180" r="44" fill="#ffffff"/><circle cx="128" cy="272" r="40" fill="#ffffff"/><circle cx="384" cy="272" r="40" fill="#ffffff"/><path d="M256 240c-58 0-104 46-104 104 0 46 38 68 104 68s104-22 104-68c0-58-46-104-104-104z" fill="#ffffff"/></svg>'
    Path(ICON_OUTPUT).write_text(icon_svg, encoding="utf-8")

def build():
    wb_vets = load_workbook(WORKBOOK_VETS, data_only=True)
    master = read_sheet(wb_vets, "MASTER_LIST")
    hours = read_sheet(wb_vets, "VET_CONTACT_HOURS")
    branches = read_sheet(wb_vets, "BRANCH_RELATIONSHIPS")
    prices_data = read_sheet(wb_vets, "VET_PRICES")
    services_data = read_sheet(wb_vets, "VET_SERVICES")

    places_data = []
    try:
        wb_places = load_workbook(WORKBOOK_PLACES, data_only=True)
        places_data.extend(read_sheet(wb_places, 0))
    except Exception as e: print(f"Warning: Could not load {WORKBOOK_PLACES} - {e}")
    try:
        wb_rest = load_workbook(WORKBOOK_REST, data_only=True)
        places_data.extend(read_sheet(wb_rest, 0))
    except Exception as e: print(f"Warning: Could not load {WORKBOOK_REST} - {e}")
        # --- NEW: Load Vet Pricing Guide ---
    PRICING_DATA = {}
    if Path(PRICING_FILE).exists():
        print("Loading Vet Pricing Guide...")
        try:
            pricing_wb = load_workbook(PRICING_FILE, data_only=True)
            if "Price Entry Detail" in pricing_wb.sheetnames:
                ws = pricing_wb["Price Entry Detail"]
                headers = [cell.value for cell in ws[1]]
                
                id_idx = headers.index("MyBuddy Vet ID")
                service_idx = headers.index("Service")
                species_idx = headers.index("Species")
                cond_idx = headers.index("Weight / Condition")
                price_idx = headers.index("Price")
                type_idx = headers.index("Price Type")
                notes_idx = headers.index("Notes")
                
                for row in ws.iter_rows(min_row=2, values_only=True):
                    vet_id = str(row[id_idx]).strip() if row[id_idx] else None
                    if vet_id and vet_id != "GLOBAL" and vet_id != "None":
                        if vet_id not in PRICING_DATA:
                            PRICING_DATA[vet_id] = []
                        
                        PRICING_DATA[vet_id].append({
                            "service": str(row[service_idx] or "").strip(),
                            "species": str(row[species_idx] or "").strip(),
                            "condition": str(row[cond_idx] or "").strip(),
                            "price": str(row[price_idx] or "").strip(),
                            "type": str(row[type_idx] or "").strip(),
                            "notes": str(row[notes_idx] or "").strip()
                        })
            pricing_wb.close()
        except Exception as e:
            print(f"Warning: Could not load pricing file - {e}")

    funeral_keywords = ["funeral", "cremation", "cremate", "memorial"]
    funeral_vets = set()
    for r in services_data:
        mbv = str(r.get("MyBuddy Vet ID") or "").strip()
        if not mbv: continue
        row_text = " ".join(str(v).lower() for v in r.values() if v is not None)
        if any(kw in row_text for kw in funeral_keywords): funeral_vets.add(mbv)
    for r in master:
        mbv = str(r.get("MyBuddy Vet ID") or "").strip()
        if not mbv: continue
        row_text = " ".join(str(v).lower() for v in r.values() if v is not None)
        if any(kw in row_text for kw in funeral_keywords): funeral_vets.add(mbv)

    hours_by_id = {r["MyBuddy Vet ID"]: r for r in hours if r.get("MyBuddy Vet ID")}
    prices_by_id = {}
    for p in prices_data:
        mbv = str(p.get("MyBuddy Vet ID") or "").strip()
        if not mbv: continue
        service = p.get("Service"); price = p.get("Price")
        if service and price not in (None, ""): prices_by_id.setdefault(mbv, []).append(f"{service}: ${price}")

    species_by_id = extract_species_tags(services_data)
    for _id, _sp in SPECIES_OVERRIDES.items():
        species_by_id[_id] = sorted(set(species_by_id.get(_id, [])) | set(_sp))

    groups = {}
    for r in branches:
        gid = r.get("Group ID")
        if not gid: continue
        groups.setdefault(gid, []).append({"id": r.get("MBV ID"), "name": r.get("Branch Name") or ""})
    siblings = {}
    for gid, members in groups.items():
        for m in members: siblings[m["id"]] = [x for x in members if x["id"] != m["id"]]

    clinics = []
    discrepancies = []
    unknown_categories = {}

    for r in master:
        if (r.get("Active / Closed / Relocated Status") or "").strip() != "Current": continue
        mbv = str(r.get("MyBuddy Vet ID") or "").strip()
        h = hours_by_id.get(mbv, {})
        name = str(r.get("Clinic / Provider Name") or "")
        postal = clean_postal(r.get("Postal Code"), name)
        postal = POSTAL_FALLBACKS.get(postal, postal)
        addr = ", ".join(str(p).strip() for p in [r.get("Street Address"), r.get("Building Name")] if p)
        sched = {day: parse_intervals(extract_day_schedule(h.get(day), day)) if h.get(day) else None for day in DAY_NAMES}
        ecat = (h.get("Emergency Category") or "NOT_24_7_UNCONFIRMED").strip()
        if mbv in EMERGENCY_OVERRIDES: ecat = "E1_CONFIRMED_24_7"
        appt_req = str(h.get("Appointment Required") or "").strip()
        is_appt = "appointment" in appt_req.lower() and "no" not in appt_req.lower()
        clinics.append({
            "id": mbv, "category": "Vet", "name": name,
            "branch": r.get("Branch Name") or "", "address": addr, "postal": postal,
            "phone": (r.get("Phone / Main Contact") or h.get("Main Phone") or "").strip(),
            "whatsapp": (r.get("WhatsApp") or h.get("WhatsApp") or "").strip(),
            "booking": safe_http_url(h.get("Booking URL")), "ecat": ecat, "schedule": sched,
            "house_call": to_bool(r.get("Offers_House_Calls")), "hospice": to_bool(r.get("Hospice_Service")),
            "lat": None, "lng": None, "rating": r.get("Google Rating") or 0,
            "review_count": int(float(r.get("Review Count") or 0)),
            "prices": prices_by_id.get(mbv, []), "species": species_by_id.get(mbv, []),
            "hours_text": "", "place_id": "", "reviews": [],
            "has_funeral": mbv in funeral_vets
        })

    for idx, p in enumerate(places_data):
        excel_id = str(p.get("ID") or "").strip()
        raw_place_id = str(p.get("Place ID") or "").strip()
        original_id = excel_id or raw_place_id
        pid = f"LOC_{idx}"
        name = str(p.get("Name") or "").strip()
        raw_cat = str(p.get("Category") or "").strip()
        raw_hours = str(p.get("Operating Hours") or "").strip()
        
        cat = determine_category(raw_cat)
        
        if cat not in ["Vet", "Groomer", "Mall", "Training", "Park", "Pet Food", "Asian", "Western", "Cafe", "Dine"]:
            raw_norm = re.sub(r"\s+", " ", raw_cat).strip().lower()
            if raw_norm: unknown_categories[raw_norm] = unknown_categories.get(raw_norm, 0) + 1

        sched = {}
        if raw_hours:
            for day in DAY_NAMES:
                day_text = extract_day_schedule(raw_hours, day)
                sched[day] = parse_intervals(day_text)

        rating_val = p.get("Google Rating") or p.get("Rating")
        try: rating_val = float(rating_val) if rating_val else 0.0
        except: rating_val = 0.0

        postal = clean_postal(p.get("Postal Code"), name)
        postal = POSTAL_FALLBACKS.get(postal, postal)

        lat_val = p.get("Latitude"); lng_val = p.get("Longitude")
        try:
            if lat_val and lng_val: lat_val = float(lat_val); lng_val = float(lng_val)
            else: lat_val = None; lng_val = None
        except: lat_val = None; lng_val = None

        clinics.append({
            "id": pid, "original_id": original_id, "category": cat, "name": name,
            "branch": "", "address": p.get("Address") or "", "postal": postal,
            "phone": str(p.get("Phone") or "").strip(),
            "whatsapp": str(p.get("WhatsApp") or "").strip(),
            "booking": safe_http_url(p.get("Booking URL")),
            "ecat": "", "schedule": sched, "house_call": False, "hospice": False,
            "lat": lat_val, "lng": lng_val, "rating": rating_val, "review_count": int(float(p.get("Review Count") or 0)), "is_appt": False,
            "prices": [], "species": [], "hours_text": raw_hours,
            "place_id": raw_place_id, "has_funeral": False, "reviews": []
        })

    if unknown_categories:
        print("WARNING: Unknown categories:")
        for cat, cnt in sorted(unknown_categories.items()):
            print(f"  - '{cat}' ({cnt} rows)")

    cache = {}
    if CACHE_FILE.exists():
        try: cache = {k: tuple(v) for k, v in json.loads(CACHE_FILE.read_text()).items()}
        except Exception: pass

    reviews_cache = {}
    if Path(REVIEWS_CACHE_FILE).exists():
        try: reviews_cache = json.loads(Path(REVIEWS_CACHE_FILE).read_text(encoding="utf-8"))
        except Exception as e: print(f"Warning: Could not load {REVIEWS_CACHE_FILE} - {e}")

    # --- UNINDENTED! This now runs even if the cache is missing ---
    for c in clinics:
        manual_coords = None
        if c["category"] == "Vet":
            manual_coords = MANUAL_COORDS.get(c["id"])
        else:
            manual_coords = MANUAL_COORDS.get(c.get("original_id")) or MANUAL_COORDS.get(c["id"])

        g_lat = None
        g_lng = None

        if manual_coords:
            c["lat"], c["lng"] = manual_coords
            if c["category"] in ["Vet", "Dine", "Cafe", "Asian", "Western"]:
                live_rating, live_reviews, _, _, place_id = fetch_google_data(
                    c['name'], c['address'], c['postal']
                )
                if live_rating is not None:
                    try:
                        excel_r = float(c["rating"]) if c["rating"] else 0.0
                        excel_rc = int(c["review_count"]) if c["review_count"] else 0
                    except:
                        excel_r = 0.0
                        excel_rc = 0
                        
                    if abs(live_rating - excel_r) > 0.05 or (live_reviews or 0) != excel_rc:
                        discrepancies.append({
                            "ID": c["id"], "Name": c["name"], "Category": c["category"],
                            "Excel Rating": excel_r, "API Rating": live_rating,
                            "Excel Reviews": excel_rc, "API Reviews": live_reviews or 0
                        })
                        
                    if live_reviews is not None and live_reviews > 0:
                        c["rating"] = live_rating
                        c["review_count"] = live_reviews
                    c["place_id"] = place_id or ""
        else:
            live_rating, live_reviews, g_lat, g_lng, place_id = fetch_google_data(
                c['name'], c['address'], c['postal']
            )
            if live_rating is not None:
                try:
                    excel_r = float(c["rating"]) if c["rating"] else 0.0
                    excel_rc = int(c["review_count"]) if c["review_count"] else 0
                except:
                    excel_r = 0.0
                    excel_rc = 0
                    
                if abs(live_rating - excel_r) > 0.05 or (live_reviews or 0) != excel_rc:
                    discrepancies.append({
                        "ID": c["id"], "Name": c["name"], "Category": c["category"],
                        "Excel Rating": excel_r, "API Rating": live_rating,
                        "Excel Reviews": excel_rc, "API Reviews": live_reviews or 0
                    })
                    
                if live_reviews is not None and live_reviews > 0:
                    c["rating"] = live_rating
                    c["review_count"] = live_reviews
                c["place_id"] = place_id or ""
                
            if g_lat and g_lng:
                c["lat"], c["lng"] = g_lat, g_lng
                
            if c["lat"] is None or c["lng"] is None:
                coords = geocode_location(c["postal"], c["address"], c["name"], cache)
                if coords:
                    c["lat"], c["lng"] = coords
                else:
                    print(f"FAILED GEOCODE: {c['name']} - {c['address']} {c['postal']}")
                    c["lat"] = None
                    c["lng"] = None


        if c.get("lat") and c.get("lng"):
            if not (1.1 <= c["lat"] <= 1.5 and 103.5 <= c["lng"] <= 104.2):
                print(f"Geofence triggered for {c['name']} — forcing SG search...")
                c["place_id"] = ""; c["lat"] = None; c["lng"] = None
                coords = geocode_location(c["postal"], c["address"], c["name"], cache)
                if coords: c["lat"], c["lng"] = coords

        if c.get("place_id") and c["place_id"] in reviews_cache:
            c["reviews"] = reviews_cache[c["place_id"]]
        time.sleep(0.05)

    if discrepancies:
        with open('data_discrepancies.csv', 'w', newline='', encoding='utf-8') as f:
            writer = csv.DictWriter(f, fieldnames=["ID", "Name", "Category", "Excel Rating", "API Rating", "Excel Reviews", "API Reviews"])
            writer.writeheader(); writer.writerows(discrepancies)
        print(f"Report: Found {len(discrepancies)} data discrepancies. Saved to data_discrepancies.csv")
    else:
        print("Report: No data discrepancies found.")

    try: CACHE_FILE.write_text(json.dumps({k: v for k, v in cache.items()}, indent=2))
    except Exception: pass

    valid_clinics = [c for c in clinics if c["lat"] is not None and c["lng"] is not None]
    seen_coords = set()
    for c in valid_clinics:
        coord_key = (round(c["lat"], 5), round(c["lng"], 5))
        if coord_key in seen_coords:
            rng = random.Random(f"{c.get('id','')}|{coord_key}")
            c["lat"] += rng.uniform(-0.0005, 0.0005)
            c["lng"] += rng.uniform(-0.0005, 0.0005)
        else:
            seen_coords.add(coord_key)

    directory = {}
    for c in valid_clinics:
        letter = c["name"][:1].upper()
        if not letter.isalpha(): letter = "#"
        directory.setdefault(letter, []).append({"id": c["id"], "name": c["name"], "category": c["category"]})
    for L in directory: directory[L].sort(key=lambda x: x["name"].lower())

    data_json = json.dumps({"clinics": valid_clinics, "siblings": siblings, "directory": directory, "pricing_data": PRICING_DATA }, ensure_ascii=False, separators=(",", ":"))
    data_json = data_json.replace("</", "<\\/")

    Path(OUTPUT).write_text(HTML_TEMPLATE.replace("/*DATA*/", data_json), encoding="utf-8")
    write_pwa_files()
    print(f"SUCCESS: Wrote {OUTPUT} with {len(valid_clinics)} mapped locations.")

HTML_TEMPLATE = r'''<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1,viewport-fit=cover,maximum-scale=1">
<title>MyBuddy Pet Finder</title>
<link rel="manifest" href="manifest.json">
<meta name="theme-color" content="#ffffff">
<link rel="icon" type="image/svg+xml" href="icon.svg">
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap" rel="stylesheet">
<link rel="stylesheet" href="https://unpkg.com/leaflet@1.9.4/dist/leaflet.css">
<style>
html,body,#map{height:100%;margin:0;font-family:'Inter', sans-serif; background:#eef2f5;}
#map{background:#eef2f5;}

#panel{
  position:absolute; z-index:1000; top:0; left:0; bottom:0;
  width:min(360px,100vw);
  overflow-y:auto; 
  background:#ffffff;
  box-shadow:2px 0 10px rgba(0,0,0,0.05); 
  font-size:14px; display:flex; flex-direction:column;
  border-right: 1px solid #EAEAEA;
}
.panel-content { padding: 20px; flex: 1; }
.panel-content::-webkit-scrollbar { width: 0px; }

.header { display: flex; justify-content: space-between; align-items: center; margin-bottom: 24px; }
.header a { display: flex; align-items: center; gap: 8px; text-decoration: none; }
.header h1 { font-size: 20px; font-weight: 800; color: #1A1A1A; margin: 0; }
.header h1 span.blue { color: #3B82F6; }
.header img { height: 32px; width: auto; object-fit: contain; }
.header .toggle { display: flex; background: #F3F4F6; border-radius: 8px; padding: 4px; }
.header .toggle button { border: none; background: transparent; padding: 4px 12px; font-size: 12px; font-weight: 600; color: #6B7280; border-radius: 6px; cursor: pointer; }
.header .toggle button.active { background: #fff; color: #1A1A1A; box-shadow: 0 1px 3px rgba(0,0,0,0.1); }

.search-box { position: relative; margin-bottom: 8px; }
.search-box input { width: 100%; box-sizing: border-box; padding: 12px 12px 12px 36px; border: 1px solid #E5E7EB; border-radius: 10px; font-size: 14px; color: #1A1A1A; background: #F9FAFB; outline: none; }
.search-box input:focus { border-color: #3B82F6; background: #fff; }
.search-box::before { content: "📍"; position: absolute; left: 12px; top: 50%; transform: translateY(-50%); font-size: 14px; color: #9CA3AF; }

.directory { display: none; position: absolute; top: 100%; left: 0; right: 0; background: #fff; border-radius: 10px; box-shadow: 0 4px 15px rgba(0,0,0,0.1); max-height: 300px; overflow-y: auto; z-index: 2000; margin-top: 4px; border: 1px solid #E5E7EB; }
.directory.open { display: block; }
.dir-letter { border-bottom: 1px solid #F3F4F6; }
.dir-header { padding: 10px 16px; font-weight: 700; color: #1A1A1A; cursor: pointer; display: flex; justify-content: space-between; background: #F9FAFB; }
.dir-header::after { content: "+"; font-weight: 600; color: #9CA3AF; }
.dir-letter.open .dir-header::after { content: "−"; }
.dir-items { display: none; padding: 4px 0; }
.dir-letter.open .dir-items { display: block; }
.dir-item { padding: 10px 16px 10px 24px; font-size: 13px; color: #4B5563; cursor: pointer; font-weight: 500;}
.dir-item:hover { background: #F3F4F6; color: #3B82F6; }

.near-loc { display: flex; justify-content: space-between; font-size: 12px; font-weight: 600; margin-bottom: 24px; }
.near-loc span { color: #3B82F6; display: flex; align-items: center; gap: 4px; cursor:pointer; }

.section-label { font-size: 11px; font-weight: 700; color: #9CA3AF; text-transform: uppercase; letter-spacing: 0.5px; margin-bottom: 12px; display: flex; justify-content: space-between; }
.section-label a { color: #3B82F6; text-decoration: none; cursor: pointer; }
.service-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 10px; margin-bottom: 24px; }

.service-card {
    display: flex; align-items: center; gap: 10px;
    padding: 12px; border: 1px solid #E5E7EB; border-radius: 12px;
    background: #fff; cursor: pointer; text-align: left; transition: all 0.2s;
}
.service-card:hover { border-color: #3B82F6; background: #F0F9FF; }
.service-card.active { border-color: #3B82F6; background: #F0F9FF; box-shadow: 0 0 0 1px #3B82F6; }

.service-icon { 
    width: 32px; height: 32px; border-radius: 8px; 
    display: flex; align-items: center; justify-content: center; 
    background: #F3F4F6; color: #1A1A1A; 
}
.service-icon svg { width: 20px; height: 20px; stroke: currentColor; stroke-width: 2; fill: none; stroke-linecap: round; stroke-linejoin: round; }
.service-card.active .service-icon { background: #1A1A1A; color: #fff; }

.service-text h4 { margin: 0; font-size: 13px; font-weight: 700; color: #1A1A1A; }
.service-text p { margin: 0; font-size: 10px; color: #6B7280; font-weight: 500; }

.pill-row { display: flex; flex-wrap: wrap; gap: 8px; margin-bottom: 24px; }
.pill { padding: 6px 14px; border-radius: 99px; border: 1px solid #E5E7EB; background: #fff; font-size: 12px; font-weight: 600; color: #4B5563; cursor: pointer; transition: all 0.2s; display:flex; align-items:center; gap:6px;}
.pill:hover { background: #F3F4F6; }
.pill.active { background: #1A1A1A; color: #fff; border-color: #1A1A1A; }
.pill svg { width: 14px; height: 14px; stroke: currentColor; stroke-width: 2; fill: none; stroke-linecap: round; stroke-linejoin: round; }

.radius-box { background: #fff; border: 1px solid #E5E7EB; border-radius: 12px; padding: 16px; margin-bottom: 16px; }
.radius-box h4 { margin: 0 0 12px 0; font-size: 12px; font-weight: 700; color: #1A1A1A; display: flex; justify-content: space-between; align-items: center; }
.radius-box h4 span { background: #F3F4F6; padding: 2px 8px; border-radius: 6px; font-size: 11px; color: #1A1A1A; }
input[type=range] { width: 100%; accent-color: #1A1A1A; }
.radius-labels { display: flex; justify-content: space-between; font-size: 10px; color: #9CA3AF; font-weight: 600; margin-top: 8px; }

.toggle-item { display: flex; justify-content: space-between; align-items: center; margin-bottom: 12px; font-size: 12px; font-weight: 600; color: #4B5563; }
.toggle-item input { width: 16px; height: 16px; accent-color: #3B82F6; }

.disclaimer { font-size: 10px; color: #9CA3AF; text-align: center; padding-top: 12px; border-top: 1px solid #F3F4F6;}
.disclaimer b { color: #E74C3C; }

.bottom-bar {
    position: absolute; bottom: 20px; left: 50%; transform: translateX(-50%);
    background: #fff; padding: 10px 20px; border-radius: 99px;
    display: flex; align-items: center; gap: 12px;
    box-shadow: 0 4px 15px rgba(0,0,0,0.1); border: 1px solid #EAEAEA; z-index: 1000;
    font-size: 13px; color: #4B5563; font-weight: 500;
}
.bottom-bar button { background: #1A1A1A; color: #fff; border: none; padding: 8px 16px; border-radius: 99px; font-weight: 600; cursor: pointer; font-size: 12px; }

.custom-pin { background: transparent !important; border: none !important; }
.pin-img { width: 36px; height: 36px; transition: transform 0.15s ease-out; cursor: default !important; filter: drop-shadow(0px 4px 6px rgba(0,0,0,0.3)); transform-origin: bottom center; }
.custom-pin:hover .pin-img, .custom-pin:active .pin-img { transform: scale(1.1); }

.pin-img.closed-pin {
    filter: grayscale(100%) opacity(0.6) drop-shadow(0px 4px 6px rgba(0,0,0,0.1)) !important;
}

@keyframes pulse-emergency {
    0% { transform: scale(1); filter: drop-shadow(0 0 0 rgba(231, 76, 60, 0.8)); }
    50% { transform: scale(1.15); filter: drop-shadow(0 0 12px rgba(231, 76, 60, 1)); }
    100% { transform: scale(1); filter: drop-shadow(0 0 0 rgba(231, 76, 60, 0)); }
}
.emergency-pin .pin-img { animation: pulse-emergency 1.5s infinite; }

.centre-handle-icon{background:transparent!important;border:none!important}
.centre-handle{
    width: 36px; height: 36px; border: 3px solid #fff; box-shadow: 0 2px 10px rgba(0,0,0,0.3);
    border-radius: 50%; background: #3B82F6; color: #fff; font-size: 24px; font-weight: bold;
    display: flex; align-items: center; justify-content: center; cursor: pointer;
}

/* Popup outer container */
.leaflet-popup-content-wrapper {
    border-radius: 16px !important;
    padding: 0 !important;
    overflow: hidden !important;
}

/* Actual scrollable content */
.leaflet-popup-content {
    margin: 0 !important;
    padding: 0 !important;
    max-height: 65vh !important;
    overflow-y: auto !important;
    overscroll-behavior: contain;
    font-family: 'Inter', sans-serif;
    scrollbar-width: thin;
    -webkit-overflow-scrolling: touch;
}

.popup { padding: 20px; }
.popup-header { display: flex; justify-content: space-between; align-items: flex-start; margin-bottom: 8px; }
.popup h3 { font-size: 16px; margin: 0; color: #1A1A1A; font-weight: 800; line-height: 1.3; }
.popup .badge { display: inline-block; font-size: 10px; padding: 3px 8px; border-radius: 4px; margin-bottom: 12px; color: #fff; font-weight: 700; text-transform: uppercase; letter-spacing: 0.5px;}
.popup p { margin: 4px 0; font-weight: 500; font-size: 12px; color: #4B5563; line-height: 1.4; }
.popup .rating { display: inline-flex; align-items: center; gap: 4px; font-size: 13px; font-weight: 700; color: #F59E0B; margin-bottom: 8px; }
.popup .rating span { color: #9CA3AF; font-weight: 500; font-size: 12px; }

.popup-actions { display: flex; flex-direction: column; gap: 8px; margin-top: 16px; border-top: 1px solid #E5E7EB; padding-top: 16px; }
.popup-btn { display: block; text-align: center; padding: 10px; border-radius: 8px; text-decoration: none; color: #fff !important; font-size: 13px; font-weight: 600; transition: opacity 0.2s; cursor: pointer;}
.popup-btn:hover { opacity: 0.9; }

.transport-btns { display: flex; gap: 6px; flex-wrap: wrap; }
.trans-btn { border: none; font-weight: 600; padding: 8px 12px; border-radius: 6px; font-family: inherit; cursor: pointer; font-size: 12px; flex: 1; text-align: center; text-decoration: none; display: block;}

.reviews-box { background: #F9FAFB; border: 1px solid #E5E7EB; border-radius: 8px; padding: 12px; margin-top: 12px; font-size: 12px; color: #4B5563; font-style: italic; }
.price-list { max-height: 120px; overflow-y: auto; background: #F9FAFB; padding: 10px; border-radius: 8px; font-size: 12px; border: 1px solid #E5E7EB; margin-top: 8px; }
.price-item { padding: 4px 0; border-bottom: 1px solid #E5E7EB; font-weight: 500; color: #4B5563; }
.price-item:last-child { border-bottom: none; }

#map-toast { position: fixed; bottom: 80px; left: 50%; transform: translateX(-50%); background: #1A1A1A; color: #fff; padding: 12px 24px; border-radius: 99px; font-size: 14px; z-index: 9999; box-shadow: 0 4px 15px rgba(0,0,0,0.2); display: none; text-align: center; }

#map { margin-left: 360px; }
/* --- MOBILE RESPONSIVE LAYOUT --- */
@media (max-width: 768px) {
    #panel {
        width: 100%;
        top: auto;
        bottom: 0;
        height: 25vh; /* 1/4 of the screen */
        border-right: none;
        border-top: 1px solid #E5E7EB;
        border-radius: 20px 20px 0 0;
        box-shadow: 0 -4px 12px rgba(0,0,0,0.1);
        z-index: 2000;
        transition: height 0.3s ease-in-out; /* Smooth slide effect */
    }
    
    /* When the user taps the handle, it expands to 80% */
    #panel.expanded {
        height: 80vh; 
    }
    
    #map {
        margin-left: 0;
        height: 75vh;
        position: absolute;
        top: 0;
        width: 100%;
    }

    /* Hide the search area button on mobile */
    .bottom-bar {
        display: none !important;
    }

    /* Show the drag handle */
    #mobile-handle {
        display: block !important;
    }

    /* Make the pricing modal look better on mobile */
    #pricing-modal > div {
        width: 95% !important;
        padding: 15px !important;
        max-height: 85vh !important;
    }
    
    /* Make buttons easier to tap */
    .service-card, .pill, #pricing-btn {
        padding: 14px !important;
    }
}

/* THE GREY FILTER FOR ONEMAP */
.leaflet-tile-pane { filter: grayscale(0.7) contrast(0.9) brightness(1.05); }
.leaflet-marker-pane { filter: none !important; }
</style>
</head>
<body>

<div id="map-toast"></div>


<section id="panel">
        <div id="mobile-handle" style="display: none; padding: 15px; cursor: pointer; text-align: center; width: 100%; box-sizing: border-box;">
        <div style="width: 50px; height: 5px; background: #D1D5DB; border-radius: 10px; margin: 0 auto; pointer-events: none;"></div>
    </div>

    <div class="panel-content">
      <div class="header">
          <a href="#">
              <img src="logo.png" alt="Logo" onerror="this.style.display='none'">
              <h1>MY<span class="blue">BUDDY.SG</span></h1>
          </a>
          
      </div>

      <div class="search-box">
          <input id="search" placeholder="Search or browse A-Z...">
          <div id="directory" class="directory"></div>
      </div>
      <div class="near-loc">
          <span style="color:#10B981;"><span style="font-size:10px;">●</span> Near: Current Location</span>
          <span onclick="initLocation()">Detect GPS</span>
      </div>
  <div style="padding: 0 20px 15px 20px;">
    <button id="pricing-btn" style="width: 100%; display: flex; align-items: center; gap: 12px; padding: 12px; background: white; border: 1px solid #E5E7EB; border-radius: 8px; cursor: pointer; text-align: left; transition: 0.2s; box-shadow: 0 1px 2px rgba(0,0,0,0.02);">
        <div style="background: #F3F4F6; border-radius: 6px; padding: 8px; display: flex; align-items: center; justify-content: center;">
            <svg viewBox="0 0 24 24" width="20" height="20" stroke="#1A1A1A" stroke-width="2" fill="none" stroke-linecap="round" stroke-linejoin="round">
                <line x1="12" y1="1" x2="12" y2="23"></line>
                <path d="M17 5H9.5a3.5 3.5 0 0 0 0 7h5a3.5 3.5 0 0 1 0 7H6"></path>
            </svg>
        </div>
        <div>
            <div style="font-size: 14px; font-weight: 600; color: #1A1A1A; font-family: 'Inter', sans-serif;">Vet Pricing Guide</div>
            <div style="font-size: 12px; color: #6B7280; font-family: 'Inter', sans-serif;">Compare clinic rates</div>
        </div>
    </button>
</div>

      <div class="section-label">PRIMARY SERVICE</div>
      <div class="service-grid">
          <button class="service-card active cat-btn" data-cat="Vet">
              <div class="service-icon">
                  <svg viewBox="0 0 24 24"><path d="M12 2l9 4.9V17L3 17V6.9L12 2z"></path><path d="M12 12v6"></path><path d="M9 15h6"></path></svg>
              </div>
              <div class="service-text"><h4>Veterinary</h4><p>Trusted Care</p></div>
          </button>
          <button class="service-card cat-btn" data-cat="Groomer">
              <div class="service-icon">
                  <svg viewBox="0 0 24 24"><circle cx="6" cy="6" r="3"></circle><circle cx="6" cy="18" r="3"></circle><line x1="20" y1="4" x2="8.12" y2="15.88"></line><line x1="14.47" y1="14.48" x2="20" y2="20"></line><line x1="8.12" y1="8.12" x2="12" y2="12"></line></svg>
              </div>
              <div class="service-text"><h4>Grooming</h4><p>Spas & Styling</p></div>
          </button>
          <button class="service-card cat-btn" data-cat="Dine">
              <div class="service-icon">
                  <svg viewBox="0 0 24 24"><path d="M3 2v7c0 1.1.9 2 2 2h4a2 2 0 0 0 2-2V2"></path><path d="M7 2v20"></path><path d="M21 15V2v0a5 5 0 0 0-5 5v6c0 1.1.9 2 2 2h3Zm0 0v7"></path></svg>
              </div>
              <div class="service-text"><h4>Pet Dining</h4><p>Cafes & Bistros</p></div>
          </button>
          <button class="service-card cat-btn" data-cat="Mall">
              <div class="service-icon">
                  <svg viewBox="0 0 24 24"><path d="M6 2L3 6v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2V6l-3-4z"></path><line x1="3" y1="6" x2="21" y2="6"></line><path d="M16 10a4 4 0 0 1-8 0"></path></svg>
              </div>
              <div class="service-text"><h4>Pet Malls</h4><p>Shopping & Hubs</p></div>
          </button>
      </div>

      <div class="section-label">SPECIALTY & SPECIES <span>Optional</span></div>
      <div class="pill-row">
          <button class="pill sp-btn active" data-sp="dog">
              <svg viewBox="0 0 24 24"><circle cx="11" cy="4" r="2"></circle><circle cx="18" cy="8" r="2"></circle><circle cx="20" cy="16" r="2"></circle><circle cx="4" cy="8" r="2"></circle><circle cx="6" cy="16" r="2"></circle><path d="M8 12a4 4 0 0 1 8 0c0 4-8 6-8 6s-8-2-8-6a4 4 0 0 1 8 0z"></path></svg>
              Dogs
          </button>
          <button class="pill sp-btn" data-sp="cat">
              <svg viewBox="0 0 24 24"><path d="M12 5c-4.4 0-8 3.6-8 8s3.6 8 8 8 8-3.6 8-8-3.6-8-8-8z"></path><path d="M8 3l2 4"></path><path d="M16 3l-2 4"></path><circle cx="9" cy="11" r="1"></circle><circle cx="15" cy="11" r="1"></circle></svg>
              Cats
          </button>
          <button class="pill sp-btn" data-sp="exotic">
              <svg viewBox="0 0 24 24"><path d="M21 16V8a2 2 0 0 0-1-1.73l-7-4a2 2 0 0 0-2 0l-7 4A2 2 0 0 0 3 8v8a2 2 0 0 0 1 1.73l7 4a2 2 0 0 0 2 0l7-4A2 2 0 0 0 21 16z"></path><polyline points="3.27 6.96 12 12.01 20.73 6.96"></polyline><line x1="12" y1="22.08" x2="12" y2="12"></line></svg>
              Exotics
          </button>
          <button class="pill cat-btn" data-cat="Training">
              <svg viewBox="0 0 24 24"><path d="M22 10v6M2 10l10-5 10 5-10 5z"></path><path d="M6 12v5c3 3 9 3 12 0v-5"></path></svg>
              Training
          </button>
          <button class="pill cat-btn" data-cat="Park">
              <svg viewBox="0 0 24 24"><path d="M12 22v-7l-2-2"></path><path d="M12 22v-7l2-2"></path><path d="M12 15a5 5 0 0 0 5-5c0-4-5-8-5-8S7 6 7 10a5 5 0 0 0 5 5z"></path></svg>
              Parks
          </button>
      </div>

      <div class="section-label" style="margin-top: 10px;">RATING FILTER</div>
<div class="radius-box" style="padding-bottom: 20px; margin-bottom: 0; border: none; box-shadow: none; padding-left: 0; padding-right: 0;">
    <h4 style="margin-bottom: 10px; font-size: 13px; color: #1A1A1A;">Minimum Rating: <span id="ratingVal" style="color: #10B981;">4★ & Above</span></h4>
    <input type="range" id="rating-slider" min="4" max="5" step="0.05" value="4" style="width: 100%; accent-color: #10B981; cursor: pointer;">
</div>

      <div class="section-label">FILTERS <a id="reset-btn">Reset All</a></div>
      <div class="radius-box">
          <h4>Search Radius <span id="radiusVal">5 km</span></h4>
          <input id="radius" type="range" min="1" max="20" value="5" step="1">
          <div class="radius-labels">
              <span>1 km</span><span>5 km</span><span>10 km</span><span>20 km</span>
          </div>
      </div>

  </div>
  <div class="disclaimer">
      <b>100% Unbiased</b> • Zero Commission
  </div>
</section>

<div id="map"></div>

<div class="bottom-bar">
  <span id="bottom-text">Showing <b>0</b> vetted clinics within your 5 km search radius.</span>
  <button onclick="map.setView(map.getCenter(), map.getZoom())">Search this area</button>
</div>

<script src="https://unpkg.com/leaflet@1.9.4/dist/leaflet.js"></script>
<script>
L.TileLayer.prototype.options.referrerPolicy = 'strict-origin-when-cross-origin';
const DATA = /*DATA*/;
const SINGAPORE = [1.3521, 103.8198];
const SGT_OFFSET_MIN = 8 * 60;
const PRICE_EMAIL = "info@mybuddy.sg";

const southWest = L.latLng(1.24, 103.65);
const northEast = L.latLng(1.47, 104.01);
const bounds = L.latLngBounds(southWest, northEast);

const map = L.map("map", { 
    center: SINGAPORE, 
    zoom: 12, 
    zoomControl: false,
    maxBounds: bounds,
    maxBoundsViscosity: 1.0, 
    minZoom: 11
});

L.tileLayer("https://www.onemap.gov.sg/maps/tiles/Default/{z}/{x}/{y}.png", { 
    minZoom: 11, 
    maxZoom: 19, 
    attribution: '&copy; OneMap' 
}).addTo(map);

const maskColor = "#eef2f5"; 
L.rectangle([[1.47, 103.5], [2.0, 104.2]], {color: maskColor, fillColor: maskColor, fillOpacity: 1, stroke: false}).addTo(map);
L.rectangle([[1.1, 103.5], [1.24, 104.2]], {color: maskColor, fillColor: maskColor, fillOpacity: 1, stroke: false}).addTo(map);
L.rectangle([[1.24, 104.01], [1.47, 104.2]], {color: maskColor, fillColor: maskColor, fillOpacity: 1, stroke: false}).addTo(map);
L.rectangle([[1.24, 103.5], [1.47, 103.65]], {color: maskColor, fillColor: maskColor, fillOpacity: 1, stroke: false}).addTo(map);

function singaporeNow() {
  const now = new Date();
  const sgt = new Date(now.getTime() + SGT_OFFSET_MIN * 60000);
  return { day: ["Sunday","Monday","Tuesday","Wednesday","Thursday","Friday","Saturday"][sgt.getUTCDay()], mins: sgt.getUTCHours() * 60 + sgt.getUTCMinutes() };
}

function statusFor(c) {
  if (c.ecat === "E1_CONFIRMED_24_7") return { key:"open", label:"OPEN NOW · 24/7 EMERGENCY" };
  const { day, mins } = singaporeNow();
  const iv = c.schedule[day];
  if (iv === "appointment") return { key:"appt", label:"BY APPOINTMENT ONLY" };
  if (iv === null || iv === undefined) return { key:"unknown", label:"HOURS UNCERTAIN" };
  if (!iv.length) return { key:"closed", label:"CLOSED NOW" };
  return iv.some(([a,b]) => a <= mins && mins < b) ? { key:"open",  label:"OPEN NOW" } : { key:"closed", label:"CLOSED NOW" };
}

const centreCircle = L.circle(SINGAPORE, { radius: 5000, color:'#3B82F6', fillColor:'#3B82F6', fillOpacity:0.05, weight:2, dashArray: '5, 5', interactive:false }).addTo(map);
const centreHandle = L.marker(SINGAPORE, { 
    icon: L.divIcon({className:'centre-handle-icon', html:'<div class="centre-handle">+</div>', iconSize:[36,36], iconAnchor:[18,18]}), 
    draggable: true, 
    zIndexOffset: 1000 
}).addTo(map);

function updateMapCentre(ll) { 
    centreCircle.setLatLng(ll); 
    centreHandle.setLatLng(ll); 
    updateVisibility(); 
}

centreHandle.on('drag', e => updateMapCentre(e.target.getLatLng()));
centreHandle.on('dragend', e => updateMapCentre(e.target.getLatLng()));

function pinIcon(c, st) {
  let imgName = "pin-vet.png";
  let catMatch = c.category;
  let isEmergency = c.ecat === "E1_CONFIRMED_24_7";
  
  if (isEmergency) {
      return L.divIcon({ className: "custom-pin emergency-pin", html: `<img src="Pin-emergency.png" class="pin-img" />`, iconSize: [36, 36], iconAnchor: [18, 36], popupAnchor: [0, -36] });
  }
  
  if (catMatch === "Groomer") imgName = "pin-groomer.png";
  else if (catMatch === "Park") imgName = "pin-park.png";
  else if (catMatch === "Training") imgName = "pin-training.png";
  else if (catMatch === "Pet Food") imgName = "pin-food.png";
  else if (catMatch === "Mall") imgName = "pin-mall.png";
  else if (["Asian", "Western", "Cafe", "Dine"].includes(catMatch)) imgName = "pin-dine.png";

  let closedClass = (st.key === "closed" || st.key === "unknown") ? "closed-pin" : "";

  return L.divIcon({ 
      className: "custom-pin", 
      html: `<img src="${imgName}" class="pin-img ${closedClass}" />`, 
      iconSize: [36, 36], 
      iconAnchor: [18, 36], 
      popupAnchor: [0, -36] 
  });
}

function getWaNumber(p) {
    if (!p) return "";
    let d = String(p).split('/')[0].replace(/\D/g, "");
    if (d.length === 8) d = "65" + d;
    return d;
}
function getTelLink(p) {
    let d = getWaNumber(p);
    return d ? "+" + d : "";
}
function esc(s){ return String(s??"").replace(/[&<>"']/g, c=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}[c])); }

// Desktop Interceptor
window.handleCall = function(e, phoneStr) {
    if (!phoneStr) return;
    const isMobile = /iPhone|iPad|iPod|Android/i.test(navigator.userAgent);
    if (!isMobile) {
        e.preventDefault();
        navigator.clipboard.writeText(phoneStr).catch(()=>{});
        const toast = document.getElementById('map-toast');
        toast.innerHTML = `📞 Number copied: <b>${phoneStr}</b><br><span style="font-size:11px;opacity:0.8;">Open your calling app to dial.</span>`;
        toast.style.display = 'block';
        setTimeout(() => { toast.style.display = 'none'; }, 4000);
    }
}

window.handleGrab = function(e) {
    const isMobile = /iPhone|iPad|iPod|Android/i.test(navigator.userAgent);
    if (!isMobile) {
        e.preventDefault();
        const toast = document.getElementById('map-toast');
        toast.innerHTML = `🚗 GrabPet is a mobile app.<br><span style="font-size:11px;opacity:0.8;">Please use the Grab app on your phone.</span>`;
        toast.style.display = 'block';
        setTimeout(() => { toast.style.display = 'none'; }, 4000);
    }
}

function popupHTML(c, st) {
  const color = st.key==="open" ? "#10B981" : "#9CA3AF";
  
  let actionLinks = "";
    // --- Directions Link ---
  const destQuery = encodeURIComponent(`${c.name} ${c.address} Singapore ${c.postal}`);
  const dirUrl = `https://www.google.com/maps/dir/?api=1&destination=${destQuery}`;
  let dirBtn = `<a class="popup-btn" href="${dirUrl}" target="_blank" style="background:#4285F4;">🚗 Directions</a>`;
  const searchQuery = encodeURIComponent(`${c.name} ${c.address} Singapore`);
  const gUrl = `https://www.google.com/search?q=${searchQuery}`;

  function getRootUrl(u) {
      if (!u) return "";
      try { return new URL(u).origin; } catch(e) { return u; }
  }
  
  function isWhatsApp(url) {
      return url && (url.includes('wa.me') || url.includes('whatsapp.com') || url.includes('api.whatsapp'));
  }

  let waLink = "";
  if (c.whatsapp) waLink = "https://wa.me/" + getWaNumber(c.whatsapp);
  else if (isWhatsApp(c.booking)) waLink = c.booking;
  
  let hasWebsite = c.booking && !isWhatsApp(c.booking);

  if (c.category === "Vet" || c.category === "Groomer" || c.category === "Training") {
      if (hasWebsite) {
          actionLinks += `<a class="popup-btn" href="${esc(getRootUrl(c.booking))}" target="_blank" style="background:#3B82F6;">🌐 Visit Website</a>`;
      } else {
          actionLinks += `<a class="popup-btn" href="${gUrl}" target="_blank" style="background:#F59E0B;">📖 More Details</a>`;
      }
  } else if (c.category === "Mall") {
      if (hasWebsite) {
          actionLinks += `<a class="popup-btn" href="${esc(getRootUrl(c.booking))}" target="_blank" style="background:#8B5CF6;">🌐 Visit Website</a>`;
      } else {
          actionLinks += `<a class="popup-btn" href="${gUrl}" target="_blank" style="background:#8B5CF6;">📖 More Details</a>`;
      }
  } else if (["Asian", "Western", "Cafe", "Dine"].includes(c.category)) {
      actionLinks += `<a class="popup-btn" href="${gUrl}" target="_blank" style="background:#EC4899;">📖 View Menu</a>`;
      if (hasWebsite) {
          actionLinks += `<a class="popup-btn" href="${esc(getRootUrl(c.booking))}" target="_blank" style="background:#1A1A1A;">🌐 Visit Website</a>`;
      }
  }

  let phoneNum = getTelLink(c.phone);
  let contactHTML = `
    ${waLink ? `<a class="popup-btn" href="${esc(waLink)}" target="_blank" style="background:#25D366;">💬 WhatsApp</a>` : ''}
    ${c.phone ? `<a class="popup-btn" href="tel:${phoneNum}" onclick="handleCall(event, '${phoneNum}')" style="background:#1A1A1A;">📞 Call ${esc(c.phone.split('/')[0].trim())}</a>` : ''}
  `;

  let transportHtml = "";
  if (c.category === "Vet") {
      transportHtml = `
        <div style="margin-top:16px; border-top:1px solid #E5E7EB; padding-top:16px;">
          <div style="font-weight:700; color:#1A1A1A; margin-bottom:8px; font-size:12px;">Need transport?</div>
          <div class="transport-btns">
            <a class="trans-btn" href="grab://open?screenType=TRANSPORT" onclick="handleGrab(event)" style="background:#FEF3C7; color:#D97706;">GrabPet</a>
            <a class="trans-btn" href="tel:+6594891781" onclick="handleCall(event, '+6594891781')" style="background:#FEE2E2; color:#DC2626;">Petbulance</a>
            <a class="trans-btn" href="tel:+6597552015" onclick="handleCall(event, '+6597552015')" style="background:#F3E5F5; color:#6A1B9A;">Mobile Vet</a>
          </div>
        </div>
      `;
  }

  let ratingHTML = c.rating > 0 ? `<div class="rating">⭐ ${parseFloat(c.rating).toFixed(1)} ${c.review_count > 0 ? `<span>(${c.review_count} reviews)</span>` : ''}</div>` : "";

  let badgeColor = "#8B5CF6"; 
  if (c.category === "Groomer") badgeColor = "#F59E0B";
  if (c.category === "Mall") badgeColor = "#8B5CF6";
  if (c.category === "Training") badgeColor = "#FFC107";
  if (["Asian", "Western", "Cafe", "Dine"].includes(c.category)) badgeColor = "#EC4899";
  if (c.category === "Vet") badgeColor = "#3B82F6";

  let txtColor = ["Training", "Cafe"].includes(c.category) ? "#000" : "#fff";

  return `<div class="popup">
    <div class="popup-header">
      <h3>${esc(c.name)}</h3>
    </div>
    <span class="badge" style="background:${badgeColor}; color:${txtColor}; border: 1px solid rgba(0,0,0,0.1);">${esc(c.category)}</span>
    ${ratingHTML}
    <p style="color:${color}; font-weight:700; font-size:12px; text-transform:uppercase;">⏰ ${st.label}</p>
    <p>${esc(c.address)}<br>Singapore ${esc(c.postal)}</p>
    ${c.hours_text ? `<p>🕒 ${esc(c.hours_text)}</p>` : ''}
    
    <div class="popup-actions">
      ${dirBtn}
      ${actionLinks}
      ${contactHTML}
    </div>
    ${transportHtml}
  </div>`;
}

const layer = L.layerGroup().addTo(map);
const markers = []; const markersById = {}; const clinicsById = {};
let currentCategory = "Vet";
let minRating = 4; // Sets the default filter to 4 stars on load
let currentSpecies = "all";
let searchText = ""; let selectedId = null;
let filterTopRated = false;

function updateVisibility() {
  const centre = centreCircle.getLatLng(); const radius = centreCircle.getRadius();
  const q = searchText.trim().toLowerCase(); const hasSearch = q.length > 0;
  let visible = 0;

  markers.forEach(m => {
    const c = m.clinic, st = m.status;
    const isEmergency = c.ecat === "E1_CONFIRMED_24_7";
    
    let show = true;
    
    if (isEmergency) {
        show = true;
    } else {
        if (hasSearch) {
            let match = c.name.toLowerCase().includes(q) || c.id.toLowerCase().includes(q) || c.category.toLowerCase().includes(q);
            if (!match && /^\d{6}$/.test(q)) match = (c.postal === q);
            show = match;
        } else {
            if (currentCategory === "All") {
                // Show everything
            } else if (currentCategory === "Park") {
                if (c.category !== "Park") show = false;
            } else {
                if (c.category === "Park") show = false; 
                else if (c.category !== currentCategory) {
                    if (currentCategory === "Dine" && ["Asian", "Western", "Cafe", "Dine"].includes(c.category)) {
                        // allow
                    } else {
                        show = false;
                    }
                }
            }
            
            if (show && currentSpecies !== "all" && currentCategory === "Vet" && c.category === "Vet") {
              const sp = c.species || [];
              if (sp.length > 0 && !sp.includes(currentSpecies)) show = false;
            }
        }
        
        if (show && filterTopRated && c.rating < 4.5) show = false;

        // Skip the radius check for Parks and Training so they show islandwide
        if (show && !hasSearch && c.category !== "Park" && c.category !== "Training") {
          show = map.distance(centre, m.marker.getLatLng()) <= radius;
        }
    } // <--- THIS IS THE MISSING BRACKET THAT WAS CRASHING YOUR MAP!

    // Final check: Hide if rating is below the slider, but ALWAYS show emergencies
    if (show && !isEmergency && typeof minRating !== 'undefined') {
        let placeRating = parseFloat(c.rating);
        
        // If a place has no rating (or is 0), default it to 4.0
        if (isNaN(placeRating) || placeRating === 0) {
            placeRating = 4.0;
        }
        
        // Only hide if the calculated rating is below the slider
        if (placeRating < minRating - 0.01) {
            show = false;
        }
    } 

    if (show) { 
        if (!layer.hasLayer(m.marker)) layer.addLayer(m.marker); 
        visible++; 
    }
    else if (layer.hasLayer(m.marker)) layer.removeLayer(m.marker);
  });
  
  let catName = "locations";
  if (currentCategory === "Vet") catName = "vetted clinics";
  if (currentCategory === "Groomer") catName = "groomers";
  if (currentCategory === "Mall") catName = "pet malls";
  if (currentCategory === "Training") catName = "training schools";
  if (currentCategory === "Dine") catName = "pet dining spots";

  const radiusKm = radius / 1000;
  document.getElementById("bottom-text").innerHTML = `Showing <b>${visible}</b> ${catName} within your ${radiusKm} km search radius.`;
}

function setCategory(cat, element) {
    currentCategory = cat;
    document.querySelectorAll('.service-card').forEach(x => x.classList.remove('active'));
    document.querySelectorAll('.pill[data-cat]').forEach(x => x.classList.remove('active'));
    
    if (element) element.classList.add('active');
    
    if (cat !== "Vet" && element.classList.contains('pill')) {
        document.querySelectorAll('.service-card').forEach(x => x.classList.remove('active'));
    } else if (element.classList.contains('service-card')) {
        document.querySelectorAll('.pill[data-cat]').forEach(x => x.classList.remove('active'));
    }

    // --- Reset Rating Slider to 4.0 ---
    minRating = 4;
    const ratingSlider = document.getElementById('rating-slider');
    if (ratingSlider) ratingSlider.value = 4;
    const ratingVal = document.getElementById('ratingVal');
    if (ratingVal) ratingVal.textContent = "4★ & Above";
    
    searchText = "";
    document.getElementById("search").value = "";
    updateVisibility();
    buildDirectory(); 
}

document.querySelectorAll('.cat-btn').forEach(b => {
  b.addEventListener('click', () => setCategory(b.dataset.cat, b));
});

document.querySelectorAll('.sp-btn').forEach(b => {
  b.addEventListener('click', () => {
    const sp = b.dataset.sp;
    currentSpecies = (currentSpecies === sp) ? "all" : sp;
    
    if (currentSpecies !== "all") {
        currentCategory = "Vet";
        document.querySelectorAll('.service-card').forEach(x => x.classList.remove('active'));
        document.querySelector('.service-card[data-cat="Vet"]').classList.add('active');
        document.querySelectorAll('.pill').forEach(x => x.classList.remove('active'));
    }
    
    document.querySelectorAll('.sp-btn').forEach(x => x.classList.toggle('active', x.dataset.sp === currentSpecies));
    updateVisibility();
  });
});

let searchTimeout;
const searchInput = document.getElementById("search");
const directory = document.getElementById("directory");

function buildDirectory() {
    directory.innerHTML = '';
    const letters = Object.keys(DATA.directory).sort();
    letters.forEach(L => {
        const items = DATA.directory[L].filter(c => {
            if (currentCategory === "All") return true;
            if (currentCategory === "Dine") return ["Asian", "Western", "Cafe", "Dine"].includes(c.category);
            return c.category === currentCategory;
        });
        if (!items.length) return;

        const section = document.createElement('div'); section.className = 'dir-letter';
        const header = document.createElement('div'); header.className = 'dir-header'; header.textContent = `${L} (${items.length})`;
        header.onclick = () => section.classList.toggle('open');
        const list = document.createElement('div'); list.className = 'dir-items';
        items.forEach(c => {
            const item = document.createElement('div'); item.className = 'dir-item'; item.textContent = c.name;
            item.onclick = () => {
                directory.classList.remove('open');
                const m = markersById[c.id];
                if (m) {
                    updateMapCentre(m.marker.getLatLng());
                    map.flyTo(m.marker.getLatLng(), 16);
                    m.marker.openPopup();
                }
            };
            list.appendChild(item);
        });
        section.appendChild(header); section.appendChild(list); directory.appendChild(section);
    });
}

searchInput.addEventListener("focus", () => { buildDirectory(); directory.classList.add("open"); });
searchInput.addEventListener("input", () => {
  const q = searchInput.value.trim();
  clearTimeout(searchTimeout);
  
  if (!q) {
      searchText = ""; 
      updateVisibility();
      return;
  }
  
  if (/^\d{6}$/.test(q)) {
      searchTimeout = setTimeout(async () => {
          const url = "https://www.onemap.gov.sg/api/common/elastic/search?searchVal=" + q + "&returnGeom=Y&getAddrDetails=Y&pageNum=1";
          const r = await fetch(url); const j = await r.json();
          const first = (j.results || [])[0];
          if (first) {
              const ll = L.latLng(parseFloat(first.LATITUDE), parseFloat(first.LONGITUDE));
              updateMapCentre(ll);
              map.flyTo(ll, 15);
              searchText = ""; 
              updateVisibility();
          }
      }, 600);
  } else {
      searchText = q.toLowerCase();
      updateVisibility();
      directory.classList.remove("open");
  }
});

document.addEventListener("click", e => { if (!e.target.closest(".search-box")) directory.classList.remove("open"); });

document.getElementById("reset-btn").addEventListener('click', () => {
    searchInput.value = "";
    searchText = "";
    currentCategory = "Vet";
    currentSpecies = "all";
    filterTopRated = false;
    document.getElementById("radius").value = 5;
    centreCircle.setRadius(5000);
    document.getElementById("radiusVal").textContent = "5 km";
    
    document.querySelectorAll('.service-card').forEach(x => x.classList.remove('active'));
    document.querySelector('.service-card[data-cat="Vet"]').classList.add('active');
    document.querySelectorAll('.pill').forEach(x => x.classList.remove('active'));
    document.querySelector('.sp-btn[data-sp="dog"]').classList.add('active');
    
    updateVisibility();
});

let unresolved = 0;
DATA.clinics.forEach(c => {
  clinicsById[c.id] = c;
  if (c.lat == null || c.lng == null) { unresolved++; return; }
  const st = statusFor(c);
  const marker = L.marker([c.lat, c.lng], { icon: pinIcon(c, st) });
  marker.bindPopup(() => popupHTML(c, statusFor(c)), {
    maxHeight: Math.floor(window.innerHeight * 0.65),
    maxWidth: 340,
    autoPan: true,
    autoPanPaddingTopLeft: [20, 20],
    autoPanPaddingBottomRight: [20, 20],
    keepInView: true
});
  marker.clinic = c; marker.status = st;
  markers.push({ marker, clinic: c, status: st }); markersById[c.id] = { marker, clinic: c, status: st };
  layer.addLayer(marker);
  });

updateVisibility();
if (unresolved) console.log("Unresolved coords: " + unresolved);


    // --- Mobile Panel Toggle Logic ---
    const mobileHandle = document.getElementById('mobile-handle');
    const mobilePanel = document.getElementById('panel');
    if (mobileHandle && mobilePanel) {
        mobileHandle.addEventListener('click', () => {
            console.log("Handle tapped!"); 
            mobilePanel.classList.toggle('expanded');
        });
    }

function initLocation() {
  if (!navigator.geolocation) return;
  navigator.geolocation.getCurrentPosition(
    (pos) => {
      if (pos.coords.latitude > 1.1 && pos.coords.latitude < 1.5 && pos.coords.longitude > 103.5 && pos.coords.longitude < 104.2) {
        const ll = L.latLng(pos.coords.latitude, pos.coords.longitude);
        updateMapCentre(ll);
        map.setView(ll, 14);
      }
    },
    (err) => { console.log("Location permission denied."); },
    { timeout: 5000, maximumAge: 600000 }
  );
}

setTimeout(() => { map.invalidateSize(); }, 200);
setTimeout(initLocation, 1000);

    // Force Leaflet to recalculate its size when switching to mobile view
    window.addEventListener('resize', function() {
        map.invalidateSize();
    });

    // --- Rating Slider Logic ---
    var ratingSlider = document.getElementById('rating-slider');
    var ratingVal = document.getElementById('ratingVal');
    if (ratingSlider && ratingVal) {
        ratingSlider.addEventListener('input', function(e) {
            // REMOVED 'var' here so it updates the global variable
            minRating = parseFloat(e.target.value); 
            
            ratingVal.textContent = (minRating > 4) ? parseFloat(minRating.toFixed(2)) + "★ & Above" : "4★ & Above";
            
            // Uncommented this so the map actually redraws the pins!
            if (typeof updateVisibility === 'function') {
                updateVisibility(); 
            }
        });
    }

// --- Pricing Guide Logic ---
document.addEventListener('DOMContentLoaded', function() {
    var pricingBtn = document.getElementById('pricing-btn');
    var pricingModal = document.getElementById('pricing-modal');
    var closePricingBtn = document.getElementById('close-pricing');
    var pricingContent = document.getElementById('pricing-content');

    if (pricingBtn && DATA.pricing_data) {
        pricingBtn.addEventListener('click', function() {
            var htmlStr = "";
            for (var vetId in DATA.pricing_data) {
                if (DATA.pricing_data.hasOwnProperty(vetId)) {
                    var prices = DATA.pricing_data[vetId];
                    var clinic = DATA.clinics.find(function(c) { return c.id === vetId || c.original_id === vetId; });
                    var clinicName = clinic ? clinic.name : vetId;

                    htmlStr += '<h3 style="color: #333; margin-top: 20px; border-bottom: 2px solid #388E3C; padding-bottom: 5px;">' + clinicName + '</h3>';
                    htmlStr += '<table style="width: 100%; border-collapse: collapse; font-size: 14px; margin-top: 10px;">';
                    htmlStr += '<tr style="background: #f1f8e9; text-align: left;">';
                    htmlStr += '<th style="padding: 8px; border: 1px solid #c8e6c9;">Service</th>';
                    htmlStr += '<th style="padding: 8px; border: 1px solid #c8e6c9;">Species / Condition</th>';
                    htmlStr += '<th style="padding: 8px; border: 1px solid #c8e6c9;">Price (SGD)</th>';
                    htmlStr += '</tr>';

                    for (var i = 0; i < prices.length; i++) {
                        var p = prices[i];
                        var condition = [p.species, p.condition].filter(Boolean).join(" - ");
                        if (!condition) condition = "All";
                        var priceStr = p.price;
                        
                        if (p.type === "From") priceStr = "From $" + priceStr;
                        else if (p.type === "Range") priceStr = "Range: $" + priceStr;
                        else if (priceStr && priceStr.indexOf("$") === -1) priceStr = "$" + priceStr;

                        htmlStr += '<tr>';
                        htmlStr += '<td style="padding: 8px; border: 1px solid #e0e0e0;">' + p.service + '</td>';
                        htmlStr += '<td style="padding: 8px; border: 1px solid #e0e0e0;">' + condition + '</td>';
                        htmlStr += '<td style="padding: 8px; border: 1px solid #e0e0e0; font-weight: bold; color: #388E3C;">' + priceStr + '</td>';
                        htmlStr += '</tr>';
                    }
                    htmlStr += '</table>';
                }
            }
            pricingContent.innerHTML = htmlStr;
            pricingModal.style.display = 'flex';
        });

        closePricingBtn.addEventListener('click', function() {
            pricingModal.style.display = 'none';
        });

        window.addEventListener('click', function(e) {
            if (e.target === pricingModal) pricingModal.style.display = 'none';
        });
    }
});
</script>

<div id="pricing-modal" style="display: none; position: fixed; top: 0; left: 0; width: 100%; height: 100%; background: rgba(0,0,0,0.6); z-index: 9999; justify-content: center; align-items: center;">
    <div style="background: white; width: 90%; max-width: 800px; max-height: 80vh; border-radius: 12px; padding: 20px; display: flex; flex-direction: column; position: relative;">
        <button id="close-pricing" style="position: absolute; top: 15px; right: 15px; background: none; border: none; font-size: 24px; cursor: pointer; color: #333;">&times;</button>
        <h2 style="color: #388E3C; font-family: 'Baloo 2', sans-serif; margin-top: 0;">Vet Pricing Guide</h2>
        <p style="font-size: 13px; color: #d32f2f; background: #ffebee; padding: 10px; border-radius: 6px; border-left: 4px solid #d32f2f;">
            <b>Disclaimer:</b> Prices are subject to change and were accurate at the time of recording. Please verify directly with the clinic before visiting.
        </p>
        <div id="pricing-content" style="overflow-y: auto; flex-grow: 1; margin-top: 15px;"></div>
    </div>
</div>

</body>
</html>
'''

if __name__ == "__main__":
    build()
