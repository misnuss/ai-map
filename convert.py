import openpyxl
import json
from collections import defaultdict
from datetime import datetime

# 1. Excel-Datei laden
wb = openpyxl.load_workbook("AI_Guidelines_Master_Git.xlsx", read_only=True)
ws = wb.active
headers = []
rows = []
for i, row in enumerate(ws.iter_rows(values_only=True)):
    if i == 0:
        headers = [str(h).strip() if h else '' for h in row]
    else:
        rows.append(dict(zip(headers, row)))

by_country = defaultdict(list)
for r in rows:
    if r.get('country'): 
        by_country[r['country']].append(r)

features = []
# 2. Daten verarbeiten
for country, country_rows in by_country.items():
    first = country_rows[0]
    by_org = defaultdict(list)
    for r in country_rows:
        if r.get('organization'): 
            by_org[r['organization']].append(r)
            
    orgs = []
    for org_name, org_rows in by_org.items():
        fo = org_rows[0]
        lu = fo.get('last_updated')
        last_updated = str(lu.year) if isinstance(lu, datetime) else str(lu) if lu else ''
        
        links = [str(lk).strip() for lk in [fo.get('link1'), fo.get('link2')]
                 if lk and str(lk).strip() and not str(lk).startswith('chrome-extension')]
                 
        categories = list(dict.fromkeys(str(r['guideline_category']) for r in org_rows if r.get('guideline_category')))
        codes = list(dict.fromkeys(str(r['code']) for r in org_rows if r.get('code')))
        
        segments = [{"category": str(r.get('guideline_category','')), "code": str(r.get('code','')),
                     "subcode": str(r.get('subcode','') or ''), "text": str(r.get('text',''))}
                    for r in org_rows if r.get('guideline_category') and r.get('code') and r.get('text')]
                    
        orgs.append({
            "organization": str(org_name), 
            "type": str(fo.get('organization_type','')),
            "subtype": str(fo.get('organization_subtype','')), 
            "country": str(country),
            "continent": str(first.get('continent','')), 
            "last_updated": last_updated,
            "links": links, 
            "categories": categories, 
            "codes": codes, 
            "segments": segments
        })
        
    features.append({
        "type": "Feature",
        "geometry": {"type": "Point", "coordinates": [float(first.get('longitude') or 0), float(first.get('latitude') or 0)]},
        "properties": {
            "country": str(country), 
            "continent": str(first.get('continent','')),
            "city": str(first.get('city','')), 
            "org_count": len(orgs), 
            "orgs": orgs  # Direkte Objektübergabe passend zur index.html
        }
    })

geojson = {"type": "FeatureCollection", "features": features}

# 3. data.js schreiben
with open("data.js", "w", encoding="utf-8") as f:
    f.write(f"const COUNTRY_DATA = {json.dumps(geojson, ensure_ascii=False)};")

print(f"✅ {len(features)} Länder konvertiert.")
