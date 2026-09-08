#!/usr/bin/env python3
import csv, io, json, os, re, sys, urllib.request, zipfile
from openpyxl import load_workbook
from pathlib import Path

HOSPITAL_URL = "https://www.mhlw.go.jp/content/11121000/01-1_hospital_facility_info_20260601.csv.zip"
CLINIC_URL = "https://www.mhlw.go.jp/content/11121000/02-1_clinic_facility_info_20260601.csv.zip"
BED_REPORT_URL = "https://www.mhlw.go.jp/content/10800000/001717798.xlsx"

CRITICAL = [
    "いわき市医療センター",
    "太田西ノ内病院",
    "会津中央病院",
    "福島県立医科大学附属病院",
]
DISASTER = [
    "福島県立医科大学附属病院",
    "福島赤十字病院",
    "枡記念病院",
    "太田西ノ内病院",
    "総合南東北病院",
    "公立岩瀬病院",
    "白河厚生総合病院",
    "会津中央病院",
    "福島県立南会津病院",
    "南相馬市立総合病院",
    "ふたば医療センター附属病院",
    "いわき市医療センター",
]

def norm(s):
    s = (s or "").replace("\u3000"," ").strip()
    s = re.sub(r"\s+", "", s)
    s = s.replace("―","ー").replace("－","ー").replace("ｰ","ー")
    for t in ["公立大学法人","一般財団法人","公益財団法人","医療法人","社団法人","独立行政法人","福島県厚生農業協同組合連合会","一般財団法人脳神経疾患研究所附属","一般財団法人太田綜合病院附属","一般財団法人温知会"]:
        s = s.replace(t, "")
    return s

def matches(name, targets):
    n = norm(name)
    return any(norm(t) in n for t in targets)

def download(url):
    req = urllib.request.Request(url, headers={"User-Agent":"fukushima-disaster-map3/1.0"})
    with urllib.request.urlopen(req, timeout=120) as r:
        return r.read()

def unzip_csv(blob):
    with zipfile.ZipFile(io.BytesIO(blob)) as z:
        names = [n for n in z.namelist() if n.lower().endswith(".csv")]
        if not names:
            raise RuntimeError("CSV not found in ZIP")
        with z.open(names[0]) as f:
            return io.TextIOWrapper(f, encoding="utf-8-sig", newline="").read()

def find_col(headers, exact=(), contains=()):
    nh = {h.strip():h for h in headers}
    for e in exact:
        if e in nh:
            return nh[e]
    for h in headers:
        hs = h.strip()
        if all(x in hs for x in contains):
            return h
    return None

def first_matching(headers, patterns):
    for pat in patterns:
        for h in headers:
            if all(x in h for x in pat):
                return h
    return None

def to_num(v):
    if v is None: return None
    s = str(v).replace(",","").replace("件","").replace("台","").strip()
    if not s or s in {"-","―","－","未確認","不明"}: return None
    m = re.search(r"-?\d+(?:\.\d+)?", s)
    return float(m.group()) if m else None

def parse_facilities(text, kind):
    reader = csv.DictReader(io.StringIO(text))
    headers = reader.fieldnames or []
    name_col = first_matching(headers, [
        ("正式名称",), ("施設名称",), ("医療機関名",), ("機関名",)
    ])
    lat_col = first_matching(headers, [("緯度",)])
    lon_col = first_matching(headers, [("経度",)])
    pref_cols = [h for h in headers if "都道府県" in h]
    city_col = first_matching(headers, [("市区町村",),("市町村",)])
    address_cols = [h for h in headers if any(x in h for x in ["所在地","町名","番地","住所"]) and "英語" not in h]
    phone_col = first_matching(headers, [("電話番号",),("代表電話",)])
    ambulance_cols = [h for h in headers if "救急車" in h and any(x in h for x in ["搬送","受入","受け入れ","件数","患者"])]
    if not name_col:
        raise RuntimeError("facility name column not found; first headers="+repr(headers[:30]))
    if not lat_col or not lon_col:
        raise RuntimeError("lat/lon columns not found")
    out = []
    pref_samples = []
    sample_pref07 = []
    for row in reader:
        if len(pref_samples) < 30:
            for pc in pref_cols:
                v=(row.get(pc) or "").strip()
                if v and v not in pref_samples:
                    pref_samples.append(v)
        pref_hit = any((row.get(c) or "").strip() in {"07","7","福島県"} for c in pref_cols)
        if not pref_hit:
            # fallback: search all address-like values
            pref_hit = any("福島県" in (row.get(c) or "") for c in address_cols)
        if pref_hit and len(sample_pref07) < 5:
            sample_pref07.append({
                "name": row.get(name_col),
                "lat": row.get(lat_col),
                "lon": row.get(lon_col),
                "pref": {pc: row.get(pc) for pc in pref_cols},
                "address": {ac: row.get(ac) for ac in address_cols[:8]}
            })
        if not pref_hit:
            continue
        name = (row.get(name_col) or "").strip()
        lat = to_num(row.get(lat_col)); lng = to_num(row.get(lon_col))
        if not name or lat is None or lng is None:
            continue
        address_parts=[]
        for c in pref_cols + ([city_col] if city_col else []) + address_cols:
            if not c: continue
            v=(row.get(c) or "").strip()
            if v and v not in address_parts:
                address_parts.append(v)
        address="".join(address_parts)
        ambulance=None
        if kind=="hospital":
            vals=[to_num(row.get(c)) for c in ambulance_cols]
            vals=[v for v in vals if v is not None and v >= 0]
            ambulance=max(vals) if vals else None
        item={
            "name":name,
            "lat":lat,
            "lng":lng,
            "type":kind,
            "address":address,
            "phone":(row.get(phone_col) or "").strip() if phone_col else "",
            "ambulance":int(round(ambulance)) if ambulance is not None else None,
            "critical": matches(name, CRITICAL) if kind=="hospital" else False,
            "disaster": matches(name, DISASTER) if kind=="hospital" else False,
        }
        out.append(item)
    # de-duplicate by normalized name + coordinates
    dedup={}
    for x in out:
        key=(norm(x["name"]), round(x["lat"],5), round(x["lng"],5))
        dedup[key]=x
    return list(dedup.values()), {
        "name_col":name_col,"lat_col":lat_col,"lon_col":lon_col,
        "ambulance_cols":ambulance_cols,"pref_cols":pref_cols,
        "emergency_headers":[h for h in headers if "救急" in h or "搬送" in h or "救急車" in h],
        "pref_samples":pref_samples,
        "sample_pref07":sample_pref07,
        "count":len(dedup)
    }


def inspect_bed_report():
    blob = download(BED_REPORT_URL)
    wb = load_workbook(io.BytesIO(blob), read_only=True, data_only=True)
    hits=[]
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            vals=[c.value for c in row]
            for idx,v in enumerate(vals):
                s="" if v is None else str(v)
                if ("救急車" in s or "救急搬送" in s or "救急医療" in s) and len(hits)<80:
                    lo=max(0,idx-4); hi=min(len(vals),idx+8)
                    hits.append({"sheet":ws.title,"row":row[0].row,"col":idx+1,"value":s,"context":[None if x is None else str(x) for x in vals[lo:hi]]})
    ws=wb[wb.sheetnames[0]]
    row5=[ws.cell(5,i).value for i in range(1,min(ws.max_column,220)+1)]
    return {"sheetnames":wb.sheetnames,"hits":hits,"row5":[None if v is None else str(v) for v in row5]}

def category(x):
    if x["type"]=="clinic": return "clinic"
    if x.get("critical"): return "critical"
    if x.get("disaster"): return "disaster"
    a=x.get("ambulance")
    if a is not None and a>=1000: return "ambulance1000"
    if a is not None and a>=500: return "ambulance500"
    return "hospital"

def main():
    hospitals, hm = parse_facilities(unzip_csv(download(HOSPITAL_URL)), "hospital")
    clinics, cm = parse_facilities(unzip_csv(download(CLINIC_URL)), "clinic")
    facilities=hospitals+clinics
    for x in facilities:
        x["category"]=category(x)
    counts={}
    for x in facilities:
        counts[x["category"]]=counts.get(x["category"],0)+1
    bed_meta=inspect_bed_report()
    payload={
        "source":"厚生労働省 医療情報ネット オープンデータ",
        "as_of":"2026-06-01",
        "generated_from":{"hospital":HOSPITAL_URL,"clinic":CLINIC_URL},
        "counts":counts,
        "facilities":facilities,
    }
    Path("data").mkdir(exist_ok=True)
    Path("data/fukushima_medical.json").write_text(json.dumps(payload, ensure_ascii=False, separators=(",",":")),encoding="utf-8")
    Path("data/build-meta.json").write_text(json.dumps({"hospital":hm,"clinic":cm,"bed_report":bed_meta,"counts":counts},ensure_ascii=False,indent=2),encoding="utf-8")
    print(json.dumps({"hospital":hm,"clinic":cm,"bed_report":bed_meta,"counts":counts}, ensure_ascii=False, indent=2))

if __name__=="__main__":
    main()
