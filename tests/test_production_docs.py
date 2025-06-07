import sys
from pathlib import Path

# Add project paths similar to main.py
BASE_DIR = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(BASE_DIR))
sys.path.insert(0, str(BASE_DIR / "core"))
sys.path.insert(0, str(BASE_DIR / "logic"))

from logic.production_docs import process_new_order

import pytest

@pytest.fixture
def sample_order():
    return {
        "rows": [
            {
                "article": "R-1001",
                "qty": 2,
                "weight": 6.4,
                "metal": "Золото",
                "hallmark": "585",
                "color": "красный",
                "size": 16,
            },
            {
                "article": "3D-1003",
                "qty": 1,
                "weight": 3.2,
                "metal": "Золото",
                "hallmark": "585",
                "color": "красный",
                "size": 18,
            },
        ]
    }


def test_keys(sample_order):
    res = process_new_order(sample_order)
    assert set(res.keys()) == {"order_code", "items", "batches", "mapping", "wax_jobs"}


def test_items_structure(sample_order):
    res = process_new_order(sample_order)
    items = res["items"]
    assert len(items) == 3  # qty expanded
    required = {"article", "qty", "weight", "metal", "hallmark", "color", "item_barcode", "size"}
    for item in items:
        assert required <= set(item.keys())
    # barcodes unique
    barcodes = [i["item_barcode"] for i in items]
    assert len(barcodes) == len(set(barcodes))


def test_batches_and_mapping(sample_order):
    res = process_new_order(sample_order)
    batches = res["batches"]
    assert len(batches) == 1  # all rows share metal/hallmark/color
    batch = batches[0]
    required_batch = {"batch_barcode", "metal", "hallmark", "color", "qty", "total_w"}
    assert required_batch <= set(batch.keys())
    mapping = res["mapping"]
    assert isinstance(mapping, dict)
    assert batch["batch_barcode"] in mapping
    mapped_items = mapping[batch["batch_barcode"]]
    assert isinstance(mapped_items, list)
    assert len(mapped_items) == len(res["items"])


def test_wax_jobs_structure(sample_order):
    res = process_new_order(sample_order)
    jobs = res["wax_jobs"]
    assert len(jobs) == 2 * len(res["batches"])
    required_job = {"wax_job", "operation", "method", "method_title", "batch_code", "articles", "metal", "hallmark", "color", "qty", "weight", "created"}
    for job in jobs:
        assert required_job <= set(job.keys())
