# -*- coding: utf-8 -*-
from __future__ import annotations
from typing import Any
from .com_bridge import log, safe_str


class WaxBridge:
    """Часть COM-моста для работы с нарядами и заданиями."""

    def __init__(self, bridge: 'COM1CBridge'):
        self.bridge = bridge

    def post_task(self, number: str) -> bool:
        return self.bridge.post_task(number)
