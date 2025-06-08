# production_docs.py • v0.3
# ─────────────────────────────────────────────────────────────────────────
import itertools, uuid, datetime
from collections import defaultdict
from copy import deepcopy
from typing import Dict, Any, List, Tuple

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
    """Формирует наряды из партий.

    Наряды разделяются по методам внутри каждой партии. Базовые партии
    остаются сгруппированными только по металлу, пробе и цвету.
    """
    jobs = []
    for b in batches:
        rows = [r for r in order["rows"]
                if (r["metal"], r["hallmark"], r["color"]) ==
                   (b["metal"], b["hallmark"], b["color"])]
        rows_by_method: dict[str, list[dict]] = defaultdict(list)
        for r in rows:
            rows_by_method[_wax_method(r["article"])].append(r)

        for method, m_rows in rows_by_method.items():
            arts = {r["article"] for r in m_rows}
            qty_sum = sum(r["qty"] for r in m_rows)
            w_sum = round(sum(r["weight"] for r in m_rows), 3)
            for op in ("cast", "tree"):
                jobs.append(dict(
                    wax_job      = new_batch_code().replace("BTH", "WX"),
                    operation    = OPS[op],
                    method       = method,
                    method_title = METHOD_LABEL[method],
                    batch_code   = b["batch_barcode"],
                    articles     = ", ".join(sorted(arts)),
                    metal        = b["metal"],
                    hallmark     = b["hallmark"],
                    color        = b["color"],
                    qty          = qty_sum,
                    weight       = w_sum,
                    created      = datetime.datetime.now().isoformat(timespec="seconds"),
                    status       = "created",
                    assigned_to  = None,
                    received_by  = None,
                    completed_by = None,
                    accepted_by  = None,
                    weight_wax   = None,
        yt8arc-codex/реализация-логики-для-управления-нарядами-и-партиями
                    sync_doc_num = None,
=======
        ep5fca-codex/реализация-логики-для-управления-нарядами-и-партиями
                    sync_doc_num = None,
 
        main
        main
                    signed_log   = []
                ))
    return jobs

# ─────────────  5. главный вход  ────────────────────────────────────────
def process_new_order(order_json: Dict[str,Any]) -> Dict[str,Any]:
    order_code = new_order_code()
    items      = expand_items(order_json)
    batches,mapping = group_by_keys(items, GROUP_KEYS_WAX_CAST)
    wax_jobs   = build_wax_jobs(order_json, batches)

    ORDERS_POOL.append(dict(order=order_json, docs=dict(
        order_code=order_code, items=items, batches=batches,
        mapping=mapping, wax_jobs=wax_jobs)))
    WAX_JOBS_POOL.extend(wax_jobs)

    return dict(order_code=order_code,
                items=items, batches=batches,
                mapping=mapping, wax_jobs=wax_jobs)

# ─────────────  service helpers  ────────────────────────────────────────
def _find_job(code: str) -> dict | None:
    """Возвращает словарь наряда по его коду."""
    return next((j for j in WAX_JOBS_POOL if j.get("wax_job") == code), None)


        yt8arc-codex/реализация-логики-для-управления-нарядами-и-партиями
=======
        ep5fca-codex/реализация-логики-для-управления-нарядами-и-партиями
        main
def get_wax_job(job_code: str) -> dict | None:
    """Публичная обёртка для поиска наряда."""
    return _find_job(job_code)


        yt8arc-codex/реализация-логики-для-управления-нарядами-и-партиями
=======
=======
        main
        main
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
