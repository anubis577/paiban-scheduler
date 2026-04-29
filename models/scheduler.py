#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""排班调度器模块"""

import random
import math
from itertools import combinations


class ShiftScheduler:
    """排班调度器"""
    
    def __init__(self, persons, seats):
        self.persons = persons
        self.seats = seats

    def can_place_together(self, p1, p2):
        """检查两人是否可以同席
        规则：
        - C1不能和C1/C2同席
        - C2可以和C2同席
        """
        # C1不能和任何C1/C2同席
        if p1.level == "C1" and p2.level in ("C1", "C2"):
            return False
        if p2.level == "C1" and p1.level in ("C1", "C2"):
            return False
        
        # C2可以和C2同席
        return True
    
    def can_add_to_seat(self, seat_people, new_person, seat, required_score=5):
        """检查是否可以添加到现有席位
        规则：
        - C1不能与C1/C2同席
        - 4人席位：分别检查ABC和BCD组合（C1不超过1，C1和C2不同时出现，C2各组合不超过2个）
        - 非4人席位：C2不能超过2个
        """
        # 先检查等级兼容性（两人之间）
        for p in seat_people:
            if not self.can_place_together(p, new_person):
                return False
        
        # 4人席位：检查如果凑齐4人时的组合是否合法
        if seat and seat.persons_count == 4 and len(seat_people) >= 2:
            # 模拟添加新人后的完整列表
            test_people = seat_people + [new_person]
            
            # 需要凑齐4人才检查
            if len(test_people) == 4:
                # ABC组合检查
                abc_c1 = sum(1 for p in test_people[:3] if p.level == 'C1')
                abc_c2 = sum(1 for p in test_people[:3] if p.level == 'C2')
                if abc_c1 > 1:  # C1超过1个
                    return False
                if abc_c1 > 0 and abc_c2 > 0:  # C1和C2同席
                    return False
                if abc_c2 > 2:  # ABC组合中C2超过2个
                    return False
                
                # BCD组合检查
                bcd_c1 = sum(1 for p in test_people[1:4] if p.level == 'C1')
                bcd_c2 = sum(1 for p in test_people[1:4] if p.level == 'C2')
                if bcd_c1 > 1:
                    return False
                if bcd_c1 > 0 and bcd_c2 > 0:
                    return False
                if bcd_c2 > 2:  # BCD组合中C2超过2个
                    return False
        else:
            # 非4人席位：C2不能超过2个
            current_c2_count = sum(1 for p in seat_people if p.level == 'C2')
            new_is_c2 = new_person.level == 'C2'
            if current_c2_count + new_is_c2 > 2:
                return False
        
        # 检查分数是否满足要求（>=3人时总分>=required_score，4人席位分别检查ABC/BCD）
        all_people = seat_people + [new_person]
        if len(all_people) >= 3:
            if seat and seat.persons_count == 4:
                # 4人席位：分别检查ABC和BCD组合
                # 假设按添加顺序排列位置
                if len(all_people) >= 3:
                    # ABC组合（位置0,1,2）
                    abc_score = sum(p.effective_score for p in all_people[:3])
                    if abc_score < required_score:
                        return False
                if len(all_people) == 4:
                    # BCD组合（位置1,2,3）
                    bcd_score = sum(p.effective_score for p in all_people[1:4])
                    if bcd_score < required_score:
                        return False
            else:
                # 非4人席位：检查总分
                total_score = sum(p.effective_score for p in all_people)
                if total_score < required_score:
                    return False
        
        return True

    def calc_balance_score(self, person, seat_id, seats_dict):
        """计算人员到席位的均衡分数（越低越优先选择）
        考虑：1. 该席位历史安排次数越少越好
              2. 人员总工时越多越不优先
              3. 同一人同一席位历史次数越少越好
              4. 避免连续安排到上次同一席位
        """
        score = 0

        # 同一席位历史安排次数越多，得分越高（不优先）
        history_seat_count = person.seat_history.get(seat_id, 0)
        score += history_seat_count * 10

        # 人员总工时（越多越不优先）
        score += person.total_hours / 60 * 2  # 每小时加2分

        # 避免连续安排到上次同一席位
        if person.last_seat_app_name and seat_id == person.last_seat_app_name:
            score += 50  # 大幅加分，尽量避免同席位连续安排

        return score

    def select_best_seat(self, person, seats_dict):
        """为人员选择最优席位"""
        best_seat = None
        best_score = float('inf')

        for seat in self.seats:
            if not seat.available:
                continue
            if len(seats_dict.get(seat.app_name, [])) >= seat.persons_count:
                continue

            # 检查是否能加入
            if not self.can_add_to_seat(seats_dict.get(seat.app_name, []), person, seat, seat.required_score):
                continue

            # 计算均衡分数
            balance_score = self.calc_balance_score(person, seat.app_name, seats_dict)
            if balance_score < best_score:
                best_score = balance_score
                best_seat = seat

        return best_seat
    
    def check_seat_score(self, seat_people, required_score, persons_count=3):
        """检查席位内所有人员是否满足分数要求：4人席位分别检查ABC和BCD组合"""
        if required_score <= 0:
            return True
        
        # 转换为列表以便按位置索引
        people_list = list(seat_people)
        
        # 4人席位：分别检查ABC和BCD组合
        if persons_count == 4 and len(people_list) >= 3:
            # ABC组合（位置0,1,2）
            abc_people = [p for p in people_list if hasattr(p, 'position') and p.position in ['A', 'B', 'C']]
            if len(abc_people) >= 3:
                abc_score = sum(p.effective_score for p in abc_people)
                if abc_score < required_score:
                    return False
            
            # BCD组合（位置1,2,3）
            bcd_people = [p for p in people_list if hasattr(p, 'position') and p.position in ['B', 'C', 'D']]
            if len(bcd_people) >= 3:
                bcd_score = sum(p.effective_score for p in bcd_people)
                if bcd_score < required_score:
                    return False
            
            return True
        else:
            # 非4人席位：直接检查总分
            if len(people_list) < 3:
                return True
            total_score = sum(p.effective_score for p in people_list)
            return total_score >= required_score

    def generate_schedule(self):
        """生成排班表 - 使用均衡算法安排人员"""
        available = [p for p in self.persons if p.active]
        total_capacity = sum(s.persons_count for s in self.seats if s.available)

        best_seats = None
        best_balance_score = float('inf')
        best_assigned = 0

        for attempt in range(5000):
            seats = {s.app_name: [] for s in self.seats if s.available}
            assigned = set()

            # 打乱顺序但保持均衡考虑
            persons_copy = list(available)
            random.shuffle(persons_copy)

            # 按均衡分数排序，优先安排分数高（需要优先安排）的人员
            # 分数高意味着：历史工时少、同一席位安排少
            persons_copy.sort(key=lambda p: (
                p.total_hours,  # 工时少的先安排
                -sum(p.seat_history.get(s.app_name, 0) for s in self.seats)  # 历史席位次数少的先安排
            ))

            # 尝试安排每个人
            for person in persons_copy:
                # 使用均衡选择
                best_seat = self.select_best_seat(person, seats)
                if best_seat:
                    seats[best_seat.app_name].append(person)
                    assigned.add(person.name)

            # 验证约束
            valid = True
            balance_score = 0
            for seat in self.seats:
                if not seat.available:
                    continue
                seat_people = seats.get(seat.app_name, [])
                # 检查等级兼容性
                for i, p1 in enumerate(seat_people):
                    for p2 in seat_people[i+1:]:
                        if not self.can_place_together(p1, p2):
                            valid = False
                            break
                    if not valid:
                        break
                if not valid:
                    break
                # 检查分数要求
                if not self.check_seat_score(seat_people, seat.required_score, seat.persons_count):
                    valid = False
                    break
            
            # 统一规则检查：调用 check_all_rules
            if valid:
                selections = {}
                for sid, plist in seats.items():
                    # sid就是app_name字符串
                    selections[sid] = [p.name if p else None for p in plist]
                
                persons_dict = {p.name: {'name': p.name, 'level': p.level} for p in self.persons}
                seats_dict = {s.app_name: {'name': s.app_name, 'count': s.persons_count, 'available': s.available} for s in self.seats}
                
                warnings = check_all_rules(selections, seats_dict, persons_dict)
                if warnings:
                    valid = False
            
            # 计算均衡分数
            for seat in self.seats:
                if not seat.available:
                    continue
                seat_people = seats.get(seat.app_name, [])
                for p in seat_people:
                    balance_score += p.total_hours

            if valid:
                current_assigned = len(assigned)
                # 优先选择安排人数多的，其次选择均衡分数低的
                if current_assigned > best_assigned or \
                   (current_assigned == best_assigned and balance_score < best_balance_score):
                    best_assigned = current_assigned
                    best_seats = seats
                    best_balance_score = balance_score

                if current_assigned >= total_capacity:
                    break

        if best_seats:
            # 提取所有已安排人员的名字
            assigned_names = {person.name for persons in best_seats.values() for person in persons}
            unplaced = [p for p in available if p.name not in assigned_names]

            placed_count = sum(len(v) for v in best_seats.values())
            if len(unplaced) == 0:
                return best_seats, "排班成功"
            else:
                # 显示未安排的人员
                unplaced_names = [p.name for p in unplaced[:3]]
                msg = f"已安排{placed_count}人/"
                if len(unplaced) > 3:
                    msg += f"还有{len(unplaced)}人待安排: {','.join(unplaced_names)}..."
                else:
                    msg += f"未安排: {','.join(unplaced_names)}"
                return best_seats, msg

        return None, "无法生成有效的排班方案"
    
    def generate_schedule_with_prefill(self, pre_filled):
        """生成排班表 - 考虑预先安排的人员，锁定人员优先
        pre_filled: {seat_id: [Person objects]}
        锁定的人员(locked=True)不会被移动，position指定位置索引
        """
        all_persons = list(self.persons)
        
        # 分离锁定和非锁定人员
        locked_persons = [p for p in all_persons if p.active and p.locked]
        unlocked_persons = [p for p in all_persons if p.active and not p.locked]
        
        # 构建锁定人员的席位和位置映射
        # locked_seat_map: {person_name: (seat_id, position)}
        locked_seat_map = {}  # {person_name: (seat_id, position)}
        for seat_id, persons in pre_filled.items():
            for p in persons:
                if p.locked and p.position is not None:
                    locked_seat_map[p.name] = (seat_id, p.position)
        
        # 获取总容量
        total_capacity = sum(s.persons_count for s in self.seats if s.available)
        
        # 锁定人员占用的总席位容量
        locked_count = len(locked_persons)
        
        best_seats = None
        best_assigned = 0
        best_unplaced = []

        for attempt in range(5000):
            random.shuffle(unlocked_persons)
            
            # 构建初始席位：{seat_id: {position: Person 或 None}}
            # 每个席位的位置从 0 到 persons_count-1
            seats = {}
            for s in self.seats:
                if s.available:
                    seats[s.app_name] = {i: None for i in range(s.persons_count)}
            
            assigned = set()
            
            # 放入锁定人员到指定位置
            for p in locked_persons:
                if p.name in locked_seat_map:
                    seat_id, position = locked_seat_map[p.name]
                    if seat_id in seats and position in seats[seat_id]:
                        seats[seat_id][position] = p
                        assigned.add(p.name)
            
            # 按均衡分数排序解锁人员（工时少的先安排）
            unlocked_persons.sort(key=lambda p: (
                p.total_hours,
                -sum(p.seat_history.get(s.app_name, 0) for s in self.seats)
            ))
            
            # 尝试安排剩余每个人
            for person in unlocked_persons:
                # 找到空余位置
                best_pos = self._find_best_position_for_person(person, seats)
                if best_pos:
                    seat_id, position = best_pos
                    seats[seat_id][position] = person
                    assigned.add(person.name)
            
            # 验证约束
            valid = True
            for seat in self.seats:
                if not seat.available:
                    continue
                seat_people = [p for p in seats.get(seat.app_name, {}).values() if p is not None]
                # 检查等级兼容性
                for i, p1 in enumerate(seat_people):
                    for p2 in seat_people[i+1:]:
                        if not self.can_place_together(p1, p2):
                            valid = False
                            break
                if not valid:
                    break
                # 检查分数要求
                if not self.check_seat_score(seat_people, seat.required_score, seat.persons_count):
                    valid = False
                    break
            
            if valid:
                current_assigned = len(assigned)
                if current_assigned > best_assigned:
                    best_assigned = current_assigned
                    best_seats = seats
                    # 记录未安排的人员
                    best_unplaced = [p for p in all_persons if p.active and p.name not in assigned]
                
                if current_assigned >= total_capacity:
                    break
        
        if best_seats:
            # 检查空位置
            empty_positions_info = []  # [(seat_id, seat_name, position_letter), ...]
            for seat in self.seats:
                if seat.app_name not in best_seats:
                    continue
                pos_map = best_seats[seat.app_name]
                for pos in range(seat.persons_count):
                    if pos_map.get(pos) is None:
                        # 位置字母: 0=A, 1=B, 2=C, 3=D
                        pos_letter = chr(ord('A') + pos) if pos < 4 else str(pos)
                        empty_positions_info.append((seat.app_name, seat.name, pos_letter))
            
            # 转换格式：{seat_id: [Person]} 用于返回，按位置索引排序
            result = {}
            for seat_id, pos_map in best_seats.items():
                # 按位置索引排序
                sorted_people = []
                for pos in sorted(pos_map.keys()):
                    p = pos_map.get(pos)
                    if p is not None:
                        sorted_people.append(p)
                result[seat_id] = sorted_people
            
            placed_count = sum(len(v) for v in result.values())
            
            # 构建消息
            if len(best_unplaced) == 0 and len(empty_positions_info) == 0:
                return result, "排班成功"
            else:
                msg_parts = []
                if len(best_unplaced) > 0:
                    unplaced_names = [p.name for p in best_unplaced[:3]]
                    if len(best_unplaced) > 3:
                        msg_parts.append(f"未安排: {','.join(unplaced_names)}...等{len(best_unplaced)}人")
                    else:
                        msg_parts.append(f"未安排: {','.join(unplaced_names)}")
                if len(empty_positions_info) > 0:
                    # 按席位分组
                    seat_empty_map = {}
                    for seat_id, seat_name, pos_letter in empty_positions_info:
                        if seat_name not in seat_empty_map:
                            seat_empty_map[seat_name] = []
                        seat_empty_map[seat_name].append(pos_letter)
                    empty_details = [f"{name}: {','.join(pos)}" for name, pos in seat_empty_map.items()]
                    msg_parts.append(f"空位置: {'; '.join(empty_details)}")
                
                msg = f"已安排{placed_count}人/" + "；".join(msg_parts)
                return result, msg

        return None, "无法生成有效的排班方案"
    
    def _find_best_position_for_person(self, person, seats):
        """为解锁人员找到最佳空余位置"""
        best_seat = None
        best_position = None
        best_score = float('inf')
        
        for seat in self.seats:
            if seat.app_name not in seats:
                continue
            
            # 找到所有空余位置
            empty_positions = [pos for pos, p in seats[seat.app_name].items() if p is None]
            if not empty_positions:
                continue
            
            for position in empty_positions:
                # 模拟放置人员
                original = seats[seat.app_name][position]
                seats[seat.app_name][position] = person
                
                # 检查是否能放入
                current_people = [p for p in seats[seat.app_name].values() if p is not None]
                can_add = True
                for p in current_people:
                    if not self.can_place_together(p, person):
                        can_add = False
                        break
                
                # 恢复
                seats[seat.app_name][position] = original
                
                if can_add:
                    # 计算均衡分数
                    balance_score = self.calc_balance_score(person, seat.app_name, {})
                    if balance_score < best_score:
                        best_score = balance_score
                        best_seat = seat.app_name
                        best_position = position
        
        if best_seat is not None:
            return (best_seat, best_position)
        return None

    def _select_best_seat_for_person(self, person, seats):
        """为解锁人员选择最优席位（考虑均衡）"""
        best_seat = None
        best_score = float('inf')
        
        for seat in self.seats:
            if not seat.available:
                continue
            if len(seats.get(seat.app_name, [])) >= seat.persons_count:
                continue
            
            current_people = seats.get(seat.app_name, [])
            if not self.can_add_to_seat(current_people, person, seat, seat.required_score):
                continue
            
            # 计算均衡分数
            balance_score = self.calc_balance_score(person, seat.app_name, seats)
            if balance_score < best_score:
                best_score = balance_score
                best_seat = seat
        
        return best_seat


# ================================================
# 规则验证层 - 纯业务逻辑，不含UI
# ================================================

import math
from typing import List, Dict
from enum import Enum


class RuleType(Enum):
    """规则类型"""
    DUPLICATE = "重复选择"
    C1_COUNT = "C1超员"
    C1_C2_SAME = "C1C2同席" 
    COLUMN_LIMIT = "列限制"
    SCORE = "分数不足"


class ScheduleData:
    """
    统一数据容器，只存数据，不含业务逻辑
    
    数据结构:
        selections: seat_id → [name_a, name_b, name_c, name_d]  # 按位置顺序
        seats: seat_id → {name, app, count, required}
        persons: name → {id, level, score}
    """
    
    def __init__(self):
        self.selections = {}      # seat_id → [names]
        self.seats = {}           # seat_id → {name, app, count, required}
        self.persons = {}         # name → {id, level, score}
        self._dirty = False
    
    def get_selection(self, seat_id: str) -> List[str]:
        """获取席位选择"""
        return self.selections.get(seat_id, [])
    
    def set_selection(self, seat_id: str, names: List[str]):
        """设置席位选择"""
        self.selections[seat_id] = names.copy()
        self._dirty = True
    
    def get_all_selected(self) -> List[str]:
        """获取所有已选择的人员"""
        all_names = []
        for names in self.selections.values():
            all_names.extend(names)
        return all_names
    
    def get_all_selected_names_set(self) -> set:
        """获取所有已选择的人员（去重）"""
        return set(self.get_all_selected())
    
    def validate(self) -> bool:
        """验证数据完整性"""
        for seat_id, names in self.selections.items():
            if seat_id not in self.seats:
                return False
            if len(names) > self.seats[seat_id].get('count', 0):
                return False
        return True


class RuleWarning:
    """规则警告"""
    def __init__(self, rule_type: RuleType, message: str, seat_id: str = None, position: str = None):
        self.rule_type = rule_type
        self.message = message
        self.seat_id = seat_id
        self.position = position
    
    def __str__(self):
        return self.message


def get_exclusions(selections: Dict, current_seat_id, current_position: int = None) -> List[str]:
    """
    计算当前席位需要排除的人员列表
    规则：
    - 跨席位互斥：排除其他所有席位已选择的人员
    - 同席位内不互斥：允许4人席位内A/B/C/D选同一个人
    - 只弹出C1/C2同席警告，不自动排除
    """
    # current_seat_id已经是app_name字符串
    
    excluded = []
    for seat_id, names in selections.items():
        if seat_id != current_seat_id:
            excluded.extend([n for n in names if n])
    return excluded


def check_duplicate(selections: Dict[str, List[str]]) -> List[RuleWarning]:
    """检查重复选择"""
    warnings = []
    all_names = []
    for names in selections.values():
        all_names.extend(names)
    all_names = [n for n in all_names if n]
    
    if len(all_names) != len(set(all_names)):
        from collections import Counter
        counts = Counter(all_names)
        duplicates = [name for name, count in counts.items() if count > 1]
        for dup_name in duplicates:
            warnings.append(RuleWarning(
                RuleType.DUPLICATE,
                f"人员重复选择: {dup_name}",
                None
            ))
    
    return warnings


def check_seat_score(selections: Dict[str, List[str]], seats: Dict, persons: Dict) -> List[RuleWarning]:
    """检查席位分数限制：4人席位分别检查ABC和BCD组合"""
    warnings = []
    
    for seat_id_str, names in selections.items():
        if seat_id_str not in seats:
            continue
        
        seat = seats[seat_id_str]
        required_score = seat.get('required', 0)
        persons_count = seat.get('count', 3)
        if required_score <= 0:
            continue
        
        # 获取有效选择的人员
        selected = [n for n in names if n]
        # 席位分数限制：选择3人及以上时才判断
        if len(selected) < 3:
            continue
        
        # 4人席位：分别检查ABC和BCD组合
        if persons_count == 4:
            # ABC组合（位置0,1,2）
            abc_names = selected[:3]
            if len(abc_names) >= 3:
                abc_score = sum(persons.get(n, {}).get('score', 0) for n in abc_names)
                if abc_score < required_score:
                    seat_name = seat.get('name', seat_id_str)
                    warnings.append(RuleWarning(
                        RuleType.SCORE,
                        f"{seat_name}: ABC组合分数不足({abc_score}/{required_score}分)",
                        seat_id_str
                    ))
                    continue  # 已有违规，不再检查BCD
            
            # BCD组合（位置1,2,3）
            if len(selected) >= 4:
                bcd_names = selected[1:4]
                bcd_score = sum(persons.get(n, {}).get('score', 0) for n in bcd_names)
                if bcd_score < required_score:
                    seat_name = seat.get('name', seat_id_str)
                    warnings.append(RuleWarning(
                        RuleType.SCORE,
                        f"{seat_name}: BCD组合分数不足({bcd_score}/{required_score}分)",
                        seat_id_str
                    ))
        else:
            # 非4人席位：直接检查总分
            total_score = sum(persons.get(n, {}).get('score', 0) for n in selected)
            if total_score < required_score:
                seat_name = seat.get('name', seat_id_str)
                warnings.append(RuleWarning(
                    RuleType.SCORE,
                    f"{seat_name}: 分数不足({total_score}/{required_score}分)",
                    seat_id_str
                ))
    
    return warnings


def check_c1_c2_rules(selections: Dict[str, List[str]], seats: Dict, persons: Dict, template_slots: Dict[str, List] = None) -> List[RuleWarning]:
    """检查C1/C2席位规则
    规则：根据席位执勤时间（前后移除10分钟缓冲）有重叠的人员之间，C1和C2不能同时存在
    template_slots: {seat_id: [(ctrl_time, ctrl_position), (asst_time, asst_position), ...]}
    """
    warnings = []
    
    # 解析时间段为分钟数
    def parse_time(time_str):
        """解析 HHMM-HHMM 格式为 (start_minutes, end_minutes)"""
        if not time_str:
            return None
        try:
            parts = time_str.split('-')
            if len(parts) != 2:
                return None
            start = int(parts[0])
            end = int(parts[1])
            # 转换为分钟
            start_min = start // 100 * 60 + start % 100
            end_min = end // 100 * 60 + end % 100
            # 前后移除10分钟缓冲
            start_min += 10
            end_min -= 10
            if start_min >= end_min:
                return None
            return (start_min, end_min)
        except:
            return None
    
    # 检查两个时间段是否重叠
    def times_overlap(t1, t2):
        # 如果任一时间段为None，无法判断重叠，认为可以共存
        if not t1 or not t2:
            return False
        return not (t1[1] <= t2[0] or t2[1] <= t1[0])
    
    # 位置名称到索引的映射
    position_to_index = {'A': 0, 'B': 1, 'C': 2, 'D': 3}
    
    # 为每个席位构建时段映射
    # seat_times[位置索引] - 管制时段
    # seat_asst_times[行号] - 助理时段（按slot行号索引）
    seat_times = {}
    seat_asst_times = {}
    if template_slots:
        for seat_id, slots in template_slots.items():
            seat_times[seat_id] = {0: [], 1: [], 2: [], 3: []}
            seat_asst_times[seat_id] = {}
            for slot_idx, slot in enumerate(slots):
                # 管制时段 - 按位置索引存储
                ctrl_time = slot.get('ctrl_time', '')
                ctrl_pos = slot.get('ctrl_position', '')

                if ctrl_pos and ctrl_time:
                    idx = position_to_index.get(ctrl_pos)
                    if idx is not None:
                        t = parse_time(ctrl_time)
                        if t:
                            seat_times[seat_id][idx].append(t)
                
                # 助理时段 - 按slot行号存储
                asst_time = slot.get('asst_time', '')
                if asst_time:
                    t = parse_time(asst_time)
                    if t:
                        if slot_idx not in seat_asst_times[seat_id]:
                            seat_asst_times[seat_id][slot_idx] = []
                        seat_asst_times[seat_id][slot_idx].append(t)
    
    for seat_id, names in selections.items():
        name_to_level = {p['name']: p['level'] for p in persons.values()}
        
        seat_info = seats.get(seat_id, {})
        seat_name = seat_info.get('name', f'席位{seat_id}')
        
        # 获取有效的人员列表（过滤空值）
        valid_names = [n for n in names if n]
        
        # 收集每个人的所有时段
        # 每个人负责: 位置的管制时段 + asst_position对应位置的助理时段
        person_times = {}  # {name: [time_range, ...]}
        
        # 首先解析出 slot位置索引 -> asst_position 的映射
        # 例如: 位置0的asst是'B'表示slot 1，位置1的asst是'C'表示slot 2
        if template_slots and seat_id in template_slots:
            slot_list = template_slots[seat_id]
            slot_idx_to_asst_pos = {}  # {slot_idx: asst_position}
            for slot_idx, slot in enumerate(slot_list):
                ctrl_pos = slot.get('ctrl_position', '')
                asst_pos = slot.get('asst_position', '')
                if ctrl_pos:
                    ctrl_idx = position_to_index.get(ctrl_pos)
                    if ctrl_idx is not None:
                        slot_idx_to_asst_pos[ctrl_idx] = asst_pos
        
        # 然后为每个人收集时段
        for pos_idx, name in enumerate(names):
            if not name:
                continue
            
            level = name_to_level.get(name)
            if not level:
                continue
            
            if name not in person_times:
                person_times[name] = []
            
            # 管制时段 - 位置索引pos_idx
            if seat_id in seat_times and pos_idx in seat_times[seat_id]:
                for t in seat_times[seat_id][pos_idx]:
                    person_times[name].append(t)
            
            # 助理时段 - asst_position对应slot的助理时段
            if seat_id in seat_asst_times and pos_idx in slot_idx_to_asst_pos:
                asst_pos = slot_idx_to_asst_pos[pos_idx]
                # asst_position 转为位置索引，然后找那个位置的slot行号
                asst_idx = position_to_index.get(asst_pos)
                if asst_idx is not None:
                    # 找asst_position对应的slot行号
                    for slot_idx, slot in enumerate(template_slots[seat_id]):
                        if slot.get('ctrl_position', '') == asst_pos:
                            if slot_idx in seat_asst_times[seat_id]:
                                for t in seat_asst_times[seat_id][slot_idx]:
                                    person_times[name].append(t)
                            break
        
        # C1和C2人员列表（使用set去重）
        c1_names = list(set([n for n in names if n and name_to_level.get(n) == 'C1']))
        c2_names = list(set([n for n in names if n and name_to_level.get(n) == 'C2']))
        
        # 检查任意两个人的所有时段是否重叠
        # 不管位置是否相邻，只要时段重叠就报警
        def times_overlap_any(times_dict, name1, name2):
            """检查两个人的所有时段是否有任意重叠"""
            if name1 not in times_dict or name2 not in times_dict:
                return False
            times1 = times_dict[name1]  # 只是[t1, t2, ...]列表
            times2 = times_dict[name2]
            for t1 in times1:
                for t2 in times2:
                    if not t1 or not t2:
                        continue
                    # 不管位置是否相同，只要时间重叠就返回True
                    if not (t1[1] <= t2[0] or t2[1] <= t1[0]):
                        return True
            return False
        
        # 检查C1和C2之间 - C1不可以和C2同时存在
        for name1 in c1_names:
            for name2 in c2_names:
                if times_overlap_any(person_times, name1, name2):
                    warnings.append(RuleWarning(
                        RuleType.C1_C2_SAME,
                        f"{seat_name}: {name1}C1与{name2}C2执勤时间重叠，不能同时存在",
                        seat_id
                    ))
    
    return warnings


def check_c1c2_column_limit(selections: Dict[str, List[str]], seats: Dict, persons: Dict, template_slots: Dict[str, List] = None) -> List[RuleWarning]:
    """
    检查C1/C2跨席位限制（按时间重叠分组）
    
    按实际执勤时间重叠分组：
    - 时间有重叠的席位分到同一组
    
    每组C1/C2数量 ≤ max(2, floor(组内3人及以上席位数 × 0.5))
    生效席位：可用的、人数>=3的席位
    """
    warnings = []
    
    # 统计所有可用的、人数>=3的席位
    active_seats = []
    for seat_id, seat_info in seats.items():
        if seat_info.get('available', True) and seat_info.get('count', 0) >= 3:
            active_seats.append(seat_id)
    
    if len(active_seats) < 2:
        return warnings
    
    # 解析时间段为分钟数
    def parse_time(time_str):
        if not time_str:
            return None
        try:
            parts = time_str.split('-')
            if len(parts) != 2:
                return None
            start = int(parts[0])
            end = int(parts[1])
            start_min = start // 100 * 60 + start % 100
            end_min = end // 100 * 60 + end % 100
            start_min += 10
            end_min -= 10
            if start_min >= end_min:
                return None
            return (start_min, end_min)
        except:
            return None
    
    # 检查两段时间是否重叠
    def time_overlap(t1, t2):
        if not t1 or not t2:
            return False
        return not (t1[1] <= t2[0] or t2[1] <= t1[0])
    
    # 收集每个席位的所有时段
    seat_times = {}  # {seat_id: [(start, end), ...]}
    position_to_index = {'A': 0, 'B': 1, 'C': 2, 'D': 3}
    
    if template_slots:
        for seat_id in active_seats:
            if seat_id not in template_slots:
                continue
            slots = template_slots[seat_id]
            times = []
            for slot in slots:
                ctrl_time = slot.get('ctrl_time', '')
                if ctrl_time:
                    t = parse_time(ctrl_time)
                    if t:
                        times.append(t)
            seat_times[seat_id] = times
    
    # 按时间重叠分组
    groups = []  # 每组是席位ID列表
    assigned = set()
    
    for seat_id in active_seats:
        if seat_id in assigned:
            continue
        
        # 新建一个重叠组
        group = [seat_id]
        assigned.add(seat_id)
        
        seat_time = seat_times.get(seat_id, [])
        
        # 找所有与该席位时间重叠的其他席位
        for other_id in active_seats:
            if other_id in assigned:
                continue
            
            other_time = seat_times.get(other_id, [])
            
            # 检查是否重叠（任意两段时间重叠即可）
            is_overlap = False
            for t1 in seat_time:
                for t2 in other_time:
                    if time_overlap(t1, t2):
                        is_overlap = True
                        break
                if is_overlap:
                    break
            
            if is_overlap:
                group.append(other_id)
                assigned.add(other_id)
        
        groups.append(group)
    
    name_to_level = {}
    for p in persons.values():
        name_to_level[p['name']] = p['level']
    
    # 检查每组
    for group in groups:
        # 只统计3人及以上席位
        valid_seats_in_group = [sid for sid in group if seats.get(sid, {}).get('count', 0) >= 3]
        group_count = len(valid_seats_in_group)
        
        if group_count < 2:
            continue
        
        max_c = max(2, math.floor(group_count * 0.5))
        
        # 统计该组的C1/C2数量（该席位有选择的人员）
        c_count = 0
        for seat_id in valid_seats_in_group:
            names = selections.get(seat_id, [])
            for name in names:
                if name and name_to_level.get(name) in ('C1', 'C2'):
                    c_count += 1
        
        if c_count > max_c:
            seat_names = ', '.join(valid_seats_in_group)
            warnings.append(RuleWarning(
                RuleType.COLUMN_LIMIT,
                f"{seat_names} 共{group_count}个席位，C1/C2同时有{c_count}人，"
                f"限制最多{max_c}人",
                None
            ))
    
    return warnings


def check_all_rules(selections: Dict[str, List[str]], seats: Dict, persons: Dict, template_slots: Dict[str, List] = None) -> List[RuleWarning]:
    """
    检查所有规则
    返回规则警告列表（空表示通过）
    """
    warnings = []
    warnings.extend(check_duplicate(selections))
    warnings.extend(check_seat_score(selections, seats, persons))
    warnings.extend(check_c1_c2_rules(selections, seats, persons, template_slots))
    warnings.extend(check_c1c2_column_limit(selections, seats, persons, template_slots))
    return warnings


def format_warnings(warnings: List[RuleWarning]) -> str:
    """格式化警告消息"""
    if not warnings:
        return ""
    
    lines = []
    for w in warnings:
        lines.append(f"• {w.message}")
    
    return "\n".join(lines)