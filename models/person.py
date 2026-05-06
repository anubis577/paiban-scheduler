#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""人员模型"""


class Person:
    """人员类(用于排班)"""
    def __init__(self, id, name, level, score=0, active=True, score_modifier=0, seat_history=None, total_count=0, locked=False, position=None, last_seat_app_name=None):
        self.id = id
        self.name = name
        self.level = level
        self.score = score
        self.active = active
        self.score_modifier = score_modifier  # 规则分数修改
        self.seat_history = seat_history or {}  # {app_name: count}
        self.total_count = total_count  # 最近30天执勤次数
        self.locked = locked  # 是否锁定（手动指定的排班不动）
        self.position = position  # 锁定位置索引（0=A, 1=B, 2=C, 3=D）
        self.last_seat_app_name = last_seat_app_name  # 上次安排的席位

    @property
    def effective_score(self):
        """实际参与排班的分数（基础分 + 规则修改）"""
        return self.score + self.score_modifier

    def is_c_level(self):
        """是否为C1或C2"""
        return self.level in ("C1", "C2")