# production_docs.py • v0.4
# ─────────────────────────────────────────────────────────────────────────
import itertools, uuid, datetime
from collections import defaultdict
from copy import deepcopy
from typing import Dict, Any, List, Tuple
from uuid import uuid4

from catalogs import NOMENCLATURE                      # метод 3d / rubber

METHOD_LABEL = {"3d": "3D печать", "rubber": "Резина"}
WAX_JOBS_POOL: list[dict] = []     # все открытые наряды
ORDERS_POOL:  list[dict] = []     # все проведённые заказы (шапка+docs)

# статусы нарядов
JOB_STATES = [
    "created",   # создан
    "given",     # выдан исполнителю
    "done",      # выполнен и сдан
    "accepted",  # принят контролем
    "tree_ready" # ёлка собрана
]

# ─────────────  helpers  ────────────────────────────────────────────────
def _barcode(p):     return f"{p}-{uuid.uuid4().hex[:8].upper()}"
def new_order_code(): return _barcode("ORD")
def new_batch_code(): return _barcode("BTH")
def new_item_code():  return _barcode("ITM")

# ─────────────  1. разворачиваем qty в единицы  ─────────────────────────
def expand_items(order_json: Dict[str, Any]) -> List[Dict[str, Any]]:
    lst=[]
    for row in order_json["rows"]:
        unit_w = row["weight"]/row["qty"] if row["qty"] else 0
        for _ in range(row["qty"]):
            it = row.copy()  # avoid deepcopy of possible COM objects
            it["item_barcode"] = new_item_code()
            it["weight"] = unit_w
            lst.append(it)
    return lst

# ─────────────  2. партии (металл-проба-цвет)  ──────────────────────────
GROUP_KEYS_WAX_CAST = ("metal","hallmark","color")

def group_by_keys(items: list[dict], keys: tuple[str]):
    batches, mapping = [], defaultdict(list)
    items.sort(key=lambda r: tuple(r[k] for k in keys))
    for key, grp in itertools.groupby(items, key=lambda r: tuple(r[k] for k in keys)):
        grp=list(grp); code=new_batch_code()
        batches.append(dict(
            batch_barcode = code,
            **{k:v for k,v in zip(keys,key)},
            qty     = len(grp),
            total_w = round(sum(i["weight"] for i in grp),3)
        ))
        mapping[code] = [i["item_barcode"] for i in grp]
    return batches, mapping

# ─────────────  3. метод (3d / rubber) по артикулу  ─────────────────────
def _wax_method(article: str) -> str:
    """Определяем метод по артикулу.

    Если в артикула присутствует буква «д»/"d", считаем его 3D моделью.
    В противном случае возвращается метод "rubber".
    """
    art = str(article).lower()
    if "д" in art or "d" in art:
        return "3d"
    return "rubber"

# ─────────────  4. формируем 2 операции: cast & tree  ───────────────────
OPS = {"cast":"Отлив восковых заготовок",
       "tree":"Сборка восковых ёлок"}

def build_wax_jobs(order: dict, batches: list[dict]) -> list[dict]:
    result = []
    from uuid import uuid4
    grouped = defaultdict(list)

    # Привязка строк к batch_code по цвету/пробе/металлу
    batch_map = {}
    for b in batches:
        key = (b["metal"], b["hallmark"], b["color"])
        batch_map[key] = b["batch_barcode"]

    for row in order["rows"]:
        method = _wax_method(row["article"])
        key = (row["metal"], row["hallmark"], row["color"])
        batch_code = batch_map.get(key)
        row["batch_code"] = batch_code  # ← добавляем вручную
        grouped[method].append(row)

    for method, rows in grouped.items():
        wax_code = f"WX-{uuid4().hex[:8].upper()}"
        for row in rows:
            result.append({
                "wax_job": wax_code,
                "method": method,
                "articles": row["article"],
                "qty": row["qty"],
                "weight": row["weight"],
                "batch_code": row["batch_code"],
                "metal": row["metal"],
                "hallmark": row["hallmark"],
                "color": row["color"],
                "operation": "Отлив восковых заготовок" if method == "3d" else "Восковка из пресс-формы",
                "status": "created"
            })

    return result

# ─────────────  5. главный вход  ────────────────────────────────────────
def process_new_order(order_json: Dict[str,Any]) -> Dict[str,Any]:
    order_code = order_json.get("number", new_order_code())  # fallback на случай отладки без 1С

    items      = expand_items(order_json)
    batches,mapping = group_by_keys(items, GROUP_KEYS_WAX_CAST)
    wax_jobs   = build_wax_jobs(order_json, batches)
    from .state import ORDERS_POOL, WAX_JOBS_POOL
    WAX_JOBS_POOL.extend(wax_jobs)

    ORDERS_POOL.append(dict(order=order_json, docs=dict(
        order_code=order_code,
        items=items,
        batches=batches,
        mapping=mapping,
        wax_jobs=wax_jobs)))

# ─────────────  service helpers  ────────────────────────────────────────
def _find_job(code: str) -> dict | None:
    """Возвращает словарь наряда по его коду."""
    return next((j for j in WAX_JOBS_POOL if j.get("wax_job") == code), None)

def get_wax_job(job_code: str) -> dict | None:
    """Публичная обёртка для поиска наряда."""
    return _find_job(job_code)

def update_wax_job(job_code: str, updates: Dict[str, Any]) -> dict | None:
    """Обновляет поля наряда и возвращает его."""
    job = _find_job(job_code)
    if not job:
        return None
    job.update(updates)
    return job


def log_event(job_code: str, stage: str, user: str | None = None, extra: Dict[str, Any] | None = None) -> None:
    """Добавляет событие в журнал наряда."""
    job = _find_job(job_code)
    if not job:
        return
    if "signed_log" not in job or not isinstance(job["signed_log"], list):
        job["signed_log"] = []
    rec = dict(stage=stage, user=user,
               time=datetime.datetime.now().isoformat(timespec="seconds"))
    if extra:
        rec.update(extra)
    job["signed_log"].append(rec)
