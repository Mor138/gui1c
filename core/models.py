import dataclasses

@dataclasses.dataclass
class ProductionOrder:
    __formable__ = True
    number: str
    # остальные поля…
