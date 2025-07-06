# logic/state.py

ORDERS_POOL = []
WAX_JOBS_POOL = []

# Очередь нарядов для сборки ёлок
ASSEMBLY_POOL: list[dict] = []

# Сформированные ёлки после сборки
TREES_POOL: list[dict] = []
