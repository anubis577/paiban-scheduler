#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""席位模型"""


class Seat:
    """席位类（app_name作为主键）"""
    def __init__(self, app_name, available=True, persons_count=3, required_score=5):
        self.app_name = app_name  # 主键
        self.name = app_name  # 兼容属性
        self.available = available
        self.persons_count = persons_count
        self.required_score = required_score