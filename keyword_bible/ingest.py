import warnings
import openpyxl
from . import config

def normalize(kw):
    return " ".join(str(kw).strip().lower().split())

def split_multi(value):
    if value is None:
        return []
    return [p.strip() for p in str(value).split(",") if p.strip()]

def _to_int(x):
    try: return int(float(x))
    except (TypeError, ValueError): return 0

def _to_float(x):
    try: return float(x)
    except (TypeError, ValueError): return 0.0

def load_site(path, site):
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb.worksheets[0]
    out = {}
    for r in ws.iter_rows(min_row=2, values_only=True):
        if not r or r[0] is None:
            continue
        kw = str(r[0])
        norm = normalize(kw)
        pos = _to_int(r[1])
        if pos <= 0:
            continue
        rec = {
            "site": site, "keyword": kw, "keyword_norm": norm, "position": pos,
            "volume": _to_int(r[3]), "kd": _to_float(r[4]), "cpc": _to_float(r[5]),
            "url": r[6] or "", "intents": split_multi(r[15]), "serp": split_multi(r[14]),
        }
        prev = out.get(norm)
        if prev is None or pos < prev["position"]:
            out[norm] = rec
    wb.close()
    return out

def load_all():
    return {site: load_site(path, site) for site, path in config.SEMRUSH_FILES.items()}
