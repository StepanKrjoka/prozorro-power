#!/usr/bin/env python3
"""
Хто продає електроенергію публічним замовникам: аналіз закупівель у Prozorro.

Конвеєр (кожен крок можна запускати окремо):
  probe    — перевірка: скільки результатів дає пошук, які типи процедур індексуються
  search   — збір кандидатів через пошук prozorro.gov.ua (за датою оголошення)
  details  — повні дані тендерів з public-api.prozorro.gov.ua + договори з модуля contracting
             (має ліміт часу --budget-min: якщо не встигає, наступний запуск продовжить з того ж місця)
  build    — CSV-таблиці, рейтинги, docs/index.html
  all      — search + details + build;  resume — details + build

Лише стандартна бібліотека Python 3.10+. Відповіді API зберігаються в data/raw/ в скороченому вигляді
(лише поля, потрібні для аналізу) — по одному стиснутому файлу на місяць, щоб репозиторій не роздувався.
"""
import argparse
import csv
import datetime as dt
import gzip
import json
import os
import re
import sys
import time
import urllib.error
import urllib.parse
import urllib.request
from collections import defaultdict

ROOT = os.path.dirname(os.path.abspath(__file__))
DATA = os.path.join(ROOT, "data")
RAW_T = os.path.join(DATA, "raw", "tenders")
RAW_C = os.path.join(DATA, "raw", "contracts")
DOCS = os.path.join(ROOT, "docs")

SEARCH_URL = "https://prozorro.gov.ua/api/search/tenders"
SUMMARY_URL = "https://prozorro.gov.ua/api/tenders/{}/summary"
TENDER_URL = "https://public-api.prozorro.gov.ua/api/2.5/tenders/{}"
CONTRACT_URL = "https://public-api.prozorro.gov.ua/api/2.5/contracts/{}"
PUBLIC_TENDER = "https://prozorro.gov.ua/tender/{}"
SEARCH_CAP = 1000           # пошук віддає не більше 1000 результатів на запит
HEX_RE = re.compile(r"^[0-9a-f]{32}$")
FINAL_STATUSES = {"complete", "cancelled", "unsuccessful"}

HEADERS = {"User-Agent": "prozorro-power-research/1.0 (journalism; github actions)",
           "Accept": "application/json"}
GAP = {"prozorro.gov.ua": 1.0, "public-api.prozorro.gov.ua": 0.15}   # пауза між запитами до хоста, с
_last = {}
T0 = time.time()


def log(*a):
    print(time.strftime("%H:%M:%S"), *a, flush=True)


# ------------------------------------------------------------------ HTTP
def http(url, body=None, tries=6, timeout=60):
    host = urllib.parse.urlparse(url).hostname
    headers = dict(HEADERS)
    data = None
    if body is not None:
        data = json.dumps(body).encode()
        headers["Content-Type"] = "application/json"
    for i in range(tries):
        wait = _last.get(host, 0) + GAP.get(host, 0.2) - time.time()
        if wait > 0:
            time.sleep(wait)
        _last[host] = time.time()
        try:
            with urllib.request.urlopen(urllib.request.Request(url, data, headers), timeout=timeout) as r:
                return json.load(r)
        except urllib.error.HTTPError as e:
            if e.code in (404, 410):
                return None
            if e.code in (400, 401, 403) or i == tries - 1:
                raise RuntimeError(f"HTTP {e.code} від {host}: {url} :: {e.read()[:300]!r}")
            time.sleep((15 if e.code == 429 else 3) * (i + 1))
        except (urllib.error.URLError, TimeoutError, OSError) as e:
            if i == tries - 1:
                raise RuntimeError(f"{host} недоступний: {e}")
            time.sleep(3 * (i + 1))


# ------------------------------------------------------------------ files
def gz_write(path, obj):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    tmp = path + ".tmp"
    with gzip.open(tmp, "wt", encoding="utf-8") as f:
        json.dump(obj, f, ensure_ascii=False, separators=(",", ":"))
    os.replace(tmp, path)


def gz_read(path):
    try:
        with gzip.open(path, "rt", encoding="utf-8") as f:
            return json.load(f)
    except (FileNotFoundError, OSError, ValueError):
        return None


def json_read(path, default):
    try:
        with open(path, encoding="utf-8") as f:
            return json.load(f)
    except (FileNotFoundError, ValueError):
        return default


def json_write(path, obj):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(obj, f, ensure_ascii=False, indent=1)


def csv_write(path, rows, fields):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "w", encoding="utf-8-sig", newline="") as f:   # BOM — щоб Excel відкрив кирилицю
        w = csv.DictWriter(f, fieldnames=fields, extrasaction="ignore")
        w.writeheader()
        for r in rows:
            w.writerow(r)


# ------------------------------------------------------------------ config & classification
def load_config():
    cfg = json_read(os.path.join(ROOT, "config.json"), None)
    if not cfg:
        sys.exit("Немає config.json")
    for g in cfg["groups"]:
        g["_inc"] = [re.compile(p, re.I) for p in g.get("patterns", [])]
        g["_exc"] = [re.compile(p, re.I) for p in g.get("exclude", [])]
    for k in ("communal", "state"):
        cfg["_own_" + k] = [re.compile(p, re.I) for p in cfg["ownership"][k]]
    cfg["_kw"] = [k.lower() for k in cfg.get("item_keywords", [])]
    return cfg


def read_override(name, value_col):
    out = {}
    path = os.path.join(ROOT, "overrides", name)
    if os.path.exists(path):
        with open(path, encoding="utf-8-sig") as f:
            for r in csv.DictReader(f):
                code = re.sub(r"\D", "", r.get("edrpou") or "")
                if code and (r.get(value_col) or "").strip():
                    out[code] = r[value_col].strip()
    return out


def norm(name):
    return " ".join((name or "").lower().replace("’", "'").replace("`", "'").split())


def classify(cfg, name, kind=None, edrpou=None, overrides=None):
    """Повертає (group_id, причина). Остання група в config — «інші», тож групу отримує кожен замовник;
    None — лише якщо замовника вручну виключено (overrides/buyers.csv, group=exclude)."""
    if overrides and edrpou and edrpou in overrides:
        g = overrides[edrpou]
        return (None, None) if g == "exclude" else (g, "ручне віднесення (overrides/buyers.csv)")
    n = norm(name)
    for g in cfg["groups"]:
        if kind and kind in g.get("kind", []):
            return g["id"], f"kind = {kind}"
        if g["_inc"] and any(p.search(n) for p in g["_inc"]) and not any(p.search(n) for p in g["_exc"]):
            return g["id"], "назва замовника"
    return cfg["groups"][-1]["id"], "не підпадає під жодне правило"


def ownership(cfg, name, kind=None):
    n = norm(name)
    if any(p.search(n) for p in cfg["_own_communal"]):
        return "комунальна"
    if kind in ("central", "defense") or any(p.search(n) for p in cfg["_own_state"]):
        return "державна"
    return "не визначено"


def is_target_item(cfg, item):
    pref = cfg["item_cpv_prefixes"]
    cpv = ((item.get("classification") or {}).get("id") or "")
    if any(cpv.startswith(p) for p in pref):
        return True
    d = (item.get("description") or "").lower()
    return any(k in d for k in cfg["_kw"])


def edr(org):
    return re.sub(r"\s", "", str(((org or {}).get("identifier") or {}).get("id") or ""))


def amount(v):
    try:
        return float((v or {}).get("amount"))
    except (TypeError, ValueError):
        return None


KWH_UNITS = {"KWH": 1, "MWH": 1000, "GWH": 1_000_000}


def kwh_of(items):
    """Обсяг електроенергії в кВт·год за позиціями; None, якщо одиниці не розпізнано."""
    total, seen = 0.0, False
    for it in items:
        u = it.get("unit") or {}
        code, name = (u.get("code") or "").upper(), (u.get("name") or "").lower()
        f = KWH_UNITS.get(code)
        if f is None:
            if "мегават" in name or "мвт" in name:
                f = 1000
            elif "кіловат" in name or "квт" in name:
                f = 1
        try:
            q = float(it.get("quantity") or 0)
        except (TypeError, ValueError):
            q = 0
        if f is None or not q:
            return None
        total += q * f
        seen = True
    return total if seen else None


# ------------------------------------------------------------------ скорочене зберігання
def _org(o):
    o = o or {}
    return {"name": o.get("name"), "identifier": {"id": edr(o), "legalName": (o.get("identifier") or {}).get("legalName")},
            "kind": o.get("kind"), "address": {k: (o.get("address") or {}).get(k) for k in ("region", "locality")}}


def _items(items):
    return [{"id": i.get("id"), "description": (i.get("description") or "")[:300],
             "classification": {"id": (i.get("classification") or {}).get("id")},
             "quantity": i.get("quantity"), "unit": i.get("unit"), "relatedLot": i.get("relatedLot")}
            for i in items or []]


def trim_tender(t):
    return {
        "id": t.get("id"), "tenderID": t.get("tenderID"), "title": (t.get("title") or "")[:400],
        "status": t.get("status"), "procurementMethodType": t.get("procurementMethodType"),
        "date": t.get("date"), "dateCreated": t.get("dateCreated"), "dateModified": t.get("dateModified"),
        "value": t.get("value"), "procuringEntity": _org(t.get("procuringEntity")),
        "items": _items(t.get("items")),
        "lots": [{k: l.get(k) for k in ("id", "title", "status", "value")} for l in t.get("lots") or []],
        "awards": [{"id": a.get("id"), "lotID": a.get("lotID"), "status": a.get("status"), "value": a.get("value"),
                    "suppliers": [_org(s) for s in a.get("suppliers") or []]} for a in t.get("awards") or []],
        "contracts": [{"id": k.get("id"), "awardID": k.get("awardID"), "status": k.get("status"),
                       "value": k.get("value"), "dateSigned": k.get("dateSigned"), "period": k.get("period"),
                       "contractNumber": k.get("contractNumber"), "suppliers": [_org(s) for s in k.get("suppliers") or []],
                       "items": _items(k.get("items"))} for k in t.get("contracts") or []],
        "n_bids": len(t.get("bids") or []),
    }


def trim_contract(c):
    return {
        "id": c.get("id"), "status": c.get("status"), "value": c.get("value"), "amountPaid": c.get("amountPaid"),
        "dateSigned": c.get("dateSigned"), "period": c.get("period"), "contractNumber": c.get("contractNumber"),
        "suppliers": [_org(s) for s in c.get("suppliers") or []], "items": _items(c.get("items")),
        "changes": [{"date": (ch.get("date") or "")[:10], "status": ch.get("status"),
                     "rationaleTypes": ch.get("rationaleTypes")} for ch in c.get("changes") or []],
    }


class Shards:
    """Словник {ключ: запис}, розкладений по файлах data/raw/<kind>/YYYY-MM.json.gz."""
    def __init__(self, folder):
        self.folder, self.cache, self.dirty = folder, {}, set()

    def _load(self, month):
        if month not in self.cache:
            self.cache[month] = gz_read(os.path.join(self.folder, month + ".json.gz")) or {}
        return self.cache[month]

    def get(self, month, key):
        return self._load(month).get(key)

    def put(self, month, key, obj):
        self._load(month)[key] = obj
        self.dirty.add(month)

    def flush(self):
        for m in sorted(self.dirty):
            gz_write(os.path.join(self.folder, m + ".json.gz"), self.cache[m])
        self.dirty.clear()

    def months(self):
        if os.path.isdir(self.folder):
            for f in sorted(os.listdir(self.folder)):
                if f.endswith(".json.gz"):
                    self._load(f[:-8])
        return self.cache


def month_of(tid):
    m = re.match(r"UA-(\d{4})-(\d{2})", tid or "")
    return f"{m.group(1)}-{m.group(2)}" if m else "other"


# ------------------------------------------------------------------ search
def month_windows(start, end):
    d = dt.date.fromisoformat(start)
    end = dt.date.fromisoformat(end)
    while d <= end:
        nxt = (d.replace(day=1) + dt.timedelta(days=32)).replace(day=1)
        yield d.isoformat(), min(nxt - dt.timedelta(days=1), end).isoformat()
        d = nxt


def search_page(base, start, end, page):
    body = dict(base)
    body["date"] = {"notice": {"start": start, "end": end}}
    return http(f"{SEARCH_URL}?page={page}", body) or {}


def search_window(base, start, end, sink, stats, label):
    """Збирає всі результати за вікно дат; якщо їх >= ліміту (1000) — ділить вікно навпіл."""
    first = search_page(base, start, end, 1)
    total = int(first.get("total") or 0)
    s, e = dt.date.fromisoformat(start), dt.date.fromisoformat(end)
    if total >= SEARCH_CAP and s < e:
        mid = s + (e - s) // 2
        log(f"  {label} {start}..{end}: {total} ≥ ліміту, ділю вікно")
        search_window(base, start, mid.isoformat(), sink, stats, label)
        search_window(base, (mid + dt.timedelta(days=1)).isoformat(), end, sink, stats, label)
        return
    if total >= SEARCH_CAP:
        stats["truncated"].append({"query": label, "date": start, "total": total})
        log(f"  УВАГА: {label} {start}: {total} результатів за один день, отримаю лише перші {SEARCH_CAP}")
    per = int(first.get("per_page") or 20) or 20
    pages = min((total + per - 1) // per, SEARCH_CAP // per)
    got = 0
    for p in range(1, pages + 1):
        r = first if p == 1 else search_page(base, start, end, p)
        rows = r.get("data") or []
        if not rows:
            break
        for x in rows:
            sink(x, label)
        got += len(rows)
    stats["windows"].append({"query": label, "start": start, "end": end, "total": total, "fetched": got})
    log(f"  {label} {start}..{end}: {total} знайдено, отримано {got}")


def queries(cfg):
    yield "cpv", {"cpv": cfg["cpv_search"]}
    for t in cfg.get("text_search", []):
        yield f"text:{t}", {"text": t}


def cmd_probe(cfg, args):
    """Діагностика: формат відповіді пошуку, кількість результатів за місяць, які типи процедур індексуються."""
    start = args.probe_from or cfg["date_from"]
    end = ((dt.date.fromisoformat(start).replace(day=1) + dt.timedelta(days=32)).replace(day=1)
           - dt.timedelta(days=1)).isoformat()
    base = {"cpv": cfg["cpv_search"]}
    r = search_page(base, start, end, 1)
    sample = (r.get("data") or [None])[0]
    print("== Ключі відповіді пошуку:", sorted(r.keys()))
    print("== Перший запис:", json.dumps(sample, ensure_ascii=False)[:1500])
    print(f"\n== Результатів за {start}..{end} (ліміт пошуку {SEARCH_CAP}):")
    print("  усього:", r.get("total"))
    for p in cfg.get("probe_proc_types", []):
        print(f"  {p}: {search_page({**base, 'proc_type': [p]}, start, end, 1).get('total')}")
    day = start
    print(f"  за один день {day}:", search_page(base, day, day, 1).get("total"))
    print("  будь-які закупівлі, тип reporting, за день:",
          search_page({"proc_type": ["reporting"]}, day, day, 1).get("total"))
    for t in cfg.get("text_search", []):
        print(f"  текст «{t}»: {search_page({'text': t}, start, end, 1).get('total')}")


def cmd_search(cfg, args):
    today = dt.date.today().isoformat()
    date_to = cfg.get("date_to") or today
    state = json_read(os.path.join(DATA, "state.json"), {})
    done = set(state.get("searched_months", []))
    refresh_from = (dt.date.today() - dt.timedelta(days=args.refresh_days)).isoformat()
    cand_path = os.path.join(DATA, "candidates.json")
    cands = json_read(cand_path, {})
    overrides = read_override("buyers.csv", "group")
    stats = {"windows": [], "truncated": [], "seen": 0, "kept": 0}

    def sink(x, label):
        stats["seen"] += 1
        tid = x.get("tenderID")
        if not tid:
            return
        pe = x.get("procuringEntity") or {}
        name = pe.get("name")
        group, _ = classify(cfg, name, None, None, overrides)
        if group is None:
            return
        stats["kept"] += 1
        c = cands.setdefault(tid, {"found_by": []})
        if HEX_RE.match(str(x.get("id") or "")):
            c["id"] = x["id"]
        c["status"] = x.get("status")
        if label not in c["found_by"]:
            c["found_by"].append(label)

    for start, end in month_windows(cfg["date_from"], date_to):
        month = start[:7]
        if month in done and end < refresh_from and not args.full:
            continue
        log(f"Місяць {month}")
        for label, base in queries(cfg):
            search_window(base, start, end, sink, stats, label)
        if end < refresh_from:
            done.add(month)
        json_write(cand_path, cands)
        state["searched_months"] = sorted(done)
        json_write(os.path.join(DATA, "state.json"), state)

    state["last_search"] = dt.datetime.now(dt.timezone.utc).isoformat(timespec="seconds")
    json_write(os.path.join(DATA, "state.json"), state)
    log_path = os.path.join(DATA, "search_log.json")
    hist = json_read(log_path, [])
    hist.append({"run": state["last_search"], **stats})
    json_write(log_path, hist[-10:])
    log(f"Пошук завершено: переглянуто {stats['seen']}, кандидатів {len(cands)}, обрізаних вікон {len(stats['truncated'])}")


# ------------------------------------------------------------------ details
def cmd_details(cfg, args):
    cand_path = os.path.join(DATA, "candidates.json")
    cands = json_read(cand_path, {})
    T, C = Shards(RAW_T), Shards(RAW_C)
    today = dt.date.today()
    stale_end = (today - dt.timedelta(days=90)).isoformat()
    n_t = n_c = n_skip = 0
    incomplete = False
    order = sorted(cands, reverse=True)                  # спершу найновіші
    for i, tid in enumerate(order):
        if time.time() - T0 > args.budget_min * 60:
            incomplete = True
            log(f"Ліміт часу {args.budget_min} хв: оброблено {i}/{len(order)}, решту продовжить наступний запуск")
            break
        c, month = cands[tid], month_of(tid)
        t = T.get(month, tid)
        stale = t is not None and t.get("status") not in FINAL_STATUSES and \
            time.time() - float(t.get("_fetched") or 0) > args.tender_ttl * 86400
        if t is None or stale or args.refetch:
            h = c.get("id")
            if not h:
                s = http(SUMMARY_URL.format(tid))
                h = (s or {}).get("id")
                if not h:
                    n_skip += 1
                    continue
                c["id"] = h
            r = http(TENDER_URL.format(h))
            if not r:
                n_skip += 1
                continue
            t = trim_tender(r["data"])
            t["_fetched"] = time.time()
            T.put(month, tid, t)
            n_t += 1
        if not any(is_target_item(cfg, it) for it in t.get("items") or []):
            continue                                     # договори нецільових тендерів не тягнемо
        for k in t.get("contracts") or []:
            if k.get("status") not in ("active", "terminated"):
                continue
            cc = C.get(month, k["id"])
            if cc and not args.refetch:
                if cc.get("status") == "terminated":
                    continue
                p_end = ((cc.get("period") or {}).get("endDate") or "")[:10]
                if p_end and p_end < stale_end:
                    continue                             # строк дії давно минув — договір уже не зміниться
                if time.time() - float(cc.get("_fetched") or 0) < args.contract_ttl * 86400:
                    continue
            r = http(CONTRACT_URL.format(k["id"]))
            cc = trim_contract(r["data"]) if r else {"_missing": True}   # 404: договору немає в модулі contracting
            cc["_fetched"] = time.time()
            C.put(month, k["id"], cc)
            n_c += 1
        if (i + 1) % 500 == 0:
            log(f"  оброблено {i + 1}/{len(order)}: нових тендерів {n_t}, договорів {n_c}")
            T.flush(); C.flush(); json_write(cand_path, cands)
    T.flush(); C.flush(); json_write(cand_path, cands)
    state = json_read(os.path.join(DATA, "state.json"), {})
    state["details_incomplete"] = incomplete
    json_write(os.path.join(DATA, "state.json"), state)
    log(f"Деталі: завантажено тендерів {n_t}, договорів {n_c}, пропущено {n_skip}; "
        f"{'НЕ ЗАВЕРШЕНО' if incomplete else 'завершено'}")
    return incomplete


# ------------------------------------------------------------------ build
def lot_items(t, lot_id):
    return [i for i in t.get("items") or [] if not lot_id or i.get("relatedLot") == lot_id]


def fmt_kwh(v):
    return f"{v:,.0f} кВт·год".replace(",", " ") if v else ""


def cmd_build(cfg, args):
    cands = json_read(os.path.join(DATA, "candidates.json"), {})
    overrides = read_override("buyers.csv", "group")
    holdings = read_override("suppliers.csv", "holding")
    active_from, date_from = cfg["active_from"], cfg["date_from"]
    groups = {g["id"]: g["title"] for g in cfg["groups"]}
    T, C = Shards(RAW_T), Shards(RAW_C)
    T.months(); C.months()
    tenders_rows, contract_rows = [], []
    rejected = defaultdict(int)
    for tid in sorted(cands):
        month = month_of(tid)
        t = T.get(month, tid)
        if not t:
            rejected["ще не завантажено"] += 1
            continue
        pe = t.get("procuringEntity") or {}
        b_code, b_name = edr(pe), pe.get("name") or ""
        group, why = classify(cfg, b_name, pe.get("kind"), b_code, overrides)
        if not group:
            rejected["замовника виключено вручну"] += 1
            continue
        own = ownership(cfg, b_name, pe.get("kind"))
        t_items = [i for i in t.get("items") or [] if is_target_item(cfg, i)]
        if not t_items:
            rejected["немає позицій електроенергії"] += 1
            continue
        tval = amount(t.get("value"))
        awards = {a.get("id"): a for a in t.get("awards") or []}
        lots = {l.get("id"): l for l in t.get("lots") or []}
        t_kwh = kwh_of(t_items)
        tenders_rows.append({
            "tender_id": tid, "url": PUBLIC_TENDER.format(tid), "title": t.get("title"),
            "status": t.get("status"), "procedure": t.get("procurementMethodType"),
            "notice_date": (t.get("date") or t.get("dateCreated") or "")[:10],
            "buyer": b_name, "buyer_edrpou": b_code, "buyer_kind": pe.get("kind"),
            "buyer_region": (pe.get("address") or {}).get("region"),
            "group": group, "group_title": groups.get(group, group), "group_reason": why, "ownership": own,
            "expected_value": tval, "currency": (t.get("value") or {}).get("currency"),
            "vat_included": (t.get("value") or {}).get("valueAddedTaxIncluded"),
            "kwh": t_kwh, "expected_price_kwh": round(tval / t_kwh, 4) if tval and t_kwh else None,
            "quantity": fmt_kwh(t_kwh), "bids": t.get("n_bids"),
            "n_contracts": sum(1 for k in t.get("contracts") or [] if k.get("status") in ("active", "terminated")),
            "large": bool(tval and tval >= cfg["large_tender_uah"]),
        })
        for k in t.get("contracts") or []:
            if k.get("status") not in ("active", "terminated"):
                continue
            full = C.get(month, k["id"]) or {}
            if full.get("_missing"):
                full = {}
            award = awards.get(k.get("awardID")) or {}
            lot_id = award.get("lotID")
            sup = ((full.get("suppliers") or k.get("suppliers") or award.get("suppliers") or [{}])[0]) or {}
            items0 = k.get("items") or lot_items(t, lot_id)
            items1 = full.get("items") or items0
            e0 = [i for i in items0 if is_target_item(cfg, i)]
            e1 = [i for i in items1 if is_target_item(cfg, i)]
            if not e0 and not e1:
                continue
            signed = (full.get("dateSigned") or k.get("dateSigned") or "")[:10]
            period = full.get("period") or k.get("period") or {}
            p_end = (period.get("endDate") or "")[:10]
            status = full.get("status") or k.get("status")
            v_init = amount(k.get("value"))
            v_cur = amount(full.get("value")) if full else None
            paid = amount(full.get("amountPaid"))
            v_rank = paid if (status == "terminated" and paid is not None) else (v_cur if v_cur is not None else v_init)
            kwh0, kwh1 = kwh_of(e0), kwh_of(e1)
            p0 = v_init / kwh0 if v_init and kwh0 else None
            vc = v_cur if v_cur is not None else v_init
            p1 = vc / kwh1 if vc and kwh1 else p0
            changes = [ch for ch in full.get("changes") or [] if ch.get("status") in (None, "active")]
            s_code = edr(sup)
            contract_rows.append({
                "tender_id": tid, "url": PUBLIC_TENDER.format(tid), "procedure": t.get("procurementMethodType"),
                "buyer": b_name, "buyer_edrpou": b_code, "group": group, "group_title": groups.get(group, group),
                "ownership": own, "buyer_region": (pe.get("address") or {}).get("region"),
                "contract_id": k["id"], "contract_number": full.get("contractNumber") or k.get("contractNumber"),
                "contract_status": status, "date_signed": signed,
                "period_start": (period.get("startDate") or "")[:10], "period_end": p_end,
                "supplier": sup.get("name") or "не розкрито", "supplier_edrpou": s_code,
                "holding": holdings.get(s_code, ""),
                "value_initial": v_init, "value_current": v_cur, "amount_paid": paid, "value_rank": v_rank,
                "currency": ((full.get("value") or k.get("value") or {}).get("currency")),
                "vat_included": ((full.get("value") or k.get("value") or {}).get("valueAddedTaxIncluded")),
                "n_changes": len(changes),
                "change_types": "; ".join(sorted({r for ch in changes for r in ch.get("rationaleTypes") or []})),
                "expected_value": amount((lots.get(lot_id) or {}).get("value")) if lot_id else tval,
                "kwh_initial": kwh0, "kwh_current": kwh1,
                "price_initial": round(p0, 4) if p0 else None, "price_current": round(p1, 4) if p1 else None,
                "price_change_pct": round(100 * (p1 / p0 - 1), 1) if p0 and p1 else None,
                "mixed": len(e1 or e0) < len(items1 or items0),
                "active_2026": bool((p_end and p_end >= active_from) or (not p_end and signed >= active_from)),
                "signed_in_period": signed >= date_from,
                "contract_data": "contracting" if full else "лише з тендера",
                "large": bool(v_rank and v_rank >= cfg["large_contract_uah"]),
            })

    def rank(rows, key_fn, label_fn):
        agg = {}
        for r in rows:
            if (r.get("currency") or "UAH") != "UAH" or r["value_rank"] is None:
                continue
            key = key_fn(r)
            a = agg.setdefault(key, {"name": label_fn(r), "edrpou": r["supplier_edrpou"] if key == r["supplier_edrpou"] else "",
                                     "holding": r["holding"], "total": 0.0, "contracts": 0, "buyers": set(),
                                     "groups": set(), "entities": set(), "names": set(), "max_contract": 0.0,
                                     "v_priced": 0.0, "kwh": 0.0, "up10": 0})
            a["total"] += r["value_rank"]
            a["contracts"] += 1
            a["buyers"].add(r["buyer_edrpou"])
            a["groups"].add(r["group"])
            a["entities"].add(r["supplier_edrpou"])
            a["names"].add(r["supplier"])
            a["max_contract"] = max(a["max_contract"], r["value_rank"])
            if r["kwh_current"] and r["price_current"]:
                a["kwh"] += r["kwh_current"]
                a["v_priced"] += r["price_current"] * r["kwh_current"]
            if (r["price_change_pct"] or 0) > 10:
                a["up10"] += 1
        grand = sum(a["total"] for a in agg.values()) or 1
        out = []
        for a in sorted(agg.values(), key=lambda x: -x["total"]):
            out.append({"name": a["name"], "edrpou": a["edrpou"], "holding": a["holding"],
                        "total": round(a["total"], 2), "share_pct": round(100 * a["total"] / grand, 2),
                        "contracts": a["contracts"], "buyers": len(a["buyers"]), "entities": len(a["entities"]),
                        "kwh": round(a["kwh"]), "avg_price_kwh": round(a["v_priced"] / a["kwh"], 3) if a["kwh"] else None,
                        "contracts_price_up_10pct": a["up10"],
                        "groups": ", ".join(sorted(groups.get(g, g) for g in a["groups"])),
                        "max_contract": round(a["max_contract"], 2),
                        "other_names": " | ".join(sorted(a["names"] - {a["name"]}))[:500]})
        for i, r in enumerate(out, 1):
            r["rank"] = i
        return out

    in_period = [r for r in contract_rows if r["signed_in_period"]]
    act = [r for r in in_period if r["active_2026"]]
    by_entity = (lambda r: r["supplier_edrpou"] or "? " + r["supplier"], lambda r: r["supplier"])
    by_holding = (lambda r: r["holding"] or r["supplier_edrpou"] or "? " + r["supplier"],
                  lambda r: r["holding"] or r["supplier"])
    r_fields = ["rank", "name", "edrpou", "holding", "total", "share_pct", "contracts", "buyers", "entities", "kwh",
                "avg_price_kwh", "contracts_price_up_10pct", "groups", "max_contract", "other_names"]
    csv_write(os.path.join(DATA, "suppliers_active_2026.csv"), rank(act, *by_entity), r_fields)
    csv_write(os.path.join(DATA, "suppliers_signed_since_start.csv"), rank(in_period, *by_entity), r_fields)
    csv_write(os.path.join(DATA, "holdings_active_2026.csv"), rank(act, *by_holding), r_fields)

    buyers = {}
    for r in tenders_rows:
        b = buyers.setdefault(r["buyer_edrpou"] or r["buyer"], {
            "buyer": r["buyer"], "buyer_edrpou": r["buyer_edrpou"], "group": r["group"],
            "group_title": r["group_title"], "group_reason": r["group_reason"], "ownership": r["ownership"],
            "kind": r["buyer_kind"], "region": r["buyer_region"], "tenders": 0, "contracts_value": 0.0})
        b["tenders"] += 1
    for r in in_period:
        b = buyers.get(r["buyer_edrpou"] or r["buyer"])
        if b and r["value_rank"] and (r.get("currency") or "UAH") == "UAH":
            b["contracts_value"] += r["value_rank"]
    buyer_rows = sorted(buyers.values(), key=lambda x: (x["group"], -x["contracts_value"]))

    c_fields = list(contract_rows[0].keys()) if contract_rows else ["tender_id"]
    csv_write(os.path.join(DATA, "contracts.csv"), sorted(contract_rows, key=lambda r: -(r["value_rank"] or 0)), c_fields)
    csv_write(os.path.join(DATA, "tenders.csv"), sorted(tenders_rows, key=lambda r: -(r["expected_value"] or 0)),
              list(tenders_rows[0].keys()) if tenders_rows else ["tender_id"])
    csv_write(os.path.join(DATA, "buyers.csv"), buyer_rows,
              ["group", "group_title", "ownership", "buyer", "buyer_edrpou", "kind", "region", "group_reason",
               "tenders", "contracts_value"])

    meta = {
        "built": dt.datetime.now(dt.timezone.utc).isoformat(timespec="seconds"),
        "date_from": date_from, "active_from": active_from,
        "large_contract_uah": cfg["large_contract_uah"], "large_tender_uah": cfg["large_tender_uah"],
        "groups": groups, "candidates": len(cands), "tenders": len(tenders_rows),
        "contracts": len(contract_rows), "rejected": dict(rejected),
        "details_incomplete": json_read(os.path.join(DATA, "state.json"), {}).get("details_incomplete", False),
        "search_log": {k: v for k, v in ((json_read(os.path.join(DATA, "search_log.json"), []) or [{}])[-1]).items()
                       if k != "windows"},
    }
    json_write(os.path.join(DATA, "summary.json"), meta)

    # компактні дані для сторінки: словники постачальників і замовників + масиви
    sup_idx, buy_idx, S, B = {}, {}, [], []
    rows = []
    for r in in_period:
        if (r.get("currency") or "UAH") != "UAH" or r["value_rank"] is None:
            continue
        sk = r["supplier_edrpou"] or "?" + r["supplier"]
        if sk not in sup_idx:
            sup_idx[sk] = len(S); S.append([r["supplier"], r["supplier_edrpou"], r["holding"]])
        bk = r["buyer_edrpou"] or "?" + r["buyer"]
        if bk not in buy_idx:
            buy_idx[bk] = len(B); B.append([r["buyer"], r["buyer_edrpou"], r["group"], r["ownership"]])
        rnd = lambda v, n=0: None if v is None else round(v, n)
        rows.append([r["tender_id"], buy_idx[bk], sup_idx[sk], rnd(r["value_rank"]), r["date_signed"], r["period_end"],
                     1 if r["contract_status"] == "terminated" else 0, r["n_changes"],
                     (1 if r["mixed"] else 0) | (2 if r["active_2026"] else 0),
                     rnd(r["expected_value"]), rnd(r["kwh_current"]), rnd(r["price_initial"], 3),
                     rnd(r["price_current"], 3), rnd(r["value_initial"])])
    big_t = [{k: r[k] for k in ("tender_id", "url", "title", "status", "notice_date", "buyer", "buyer_edrpou",
                                "group", "ownership", "expected_value", "quantity", "expected_price_kwh", "bids")}
             for r in tenders_rows if r["large"]]
    top_buyers = []
    per_group = defaultdict(int)
    for b in sorted(buyer_rows, key=lambda x: -x["contracts_value"]):
        if per_group[b["group"]] < 200:
            per_group[b["group"]] += 1
            top_buyers.append(b)
    meta["buyers_per_group"] = {g: sum(1 for b in buyer_rows if b["group"] == g) for g in groups}
    payload = {"meta": meta, "S": S, "B": B, "c": rows, "tenders": big_t, "buyers": top_buyers}
    with open(os.path.join(ROOT, "report_template.html"), encoding="utf-8") as f:
        html = f.read()
    blob = json.dumps(payload, ensure_ascii=False, separators=(",", ":")).replace("</", "<\\/")
    os.makedirs(DOCS, exist_ok=True)
    with open(os.path.join(DOCS, "index.html"), "w", encoding="utf-8") as f:
        f.write(html.replace("/*__DATA__*/null", blob))
    log(f"Звіт: тендерів {len(tenders_rows)}, договорів {len(contract_rows)} (у звіті {len(rows)}), "
        f"сторінка {len(blob) / 1e6:.1f} МБ; відсіяно {dict(rejected)}")


# ------------------------------------------------------------------ main
def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("command", choices=["probe", "search", "details", "build", "all", "resume"])
    ap.add_argument("--full", action="store_true", help="переглянути пошук за всі місяці заново")
    ap.add_argument("--refetch", action="store_true", help="перезавантажити всі тендери і договори")
    ap.add_argument("--tender-ttl", type=float, default=3, help="оновлювати незавершені тендери раз на N днів")
    ap.add_argument("--refresh-days", type=int, default=60, help="скільки останніх днів шукати повторно")
    ap.add_argument("--contract-ttl", type=int, default=30, help="оновлювати чинні договори раз на N днів")
    ap.add_argument("--budget-min", type=float, default=300, help="ліміт часу на details, хв")
    ap.add_argument("--probe-from", help="місяць для probe, напр. 2026-03-01")
    args = ap.parse_args()
    cfg = load_config()
    if args.command == "probe":
        cmd_probe(cfg, args)
    if args.command in ("search", "all"):
        cmd_search(cfg, args)
    if args.command in ("details", "all", "resume"):
        cmd_details(cfg, args)
    if args.command in ("build", "all", "resume"):
        cmd_build(cfg, args)


if __name__ == "__main__":
    main()
