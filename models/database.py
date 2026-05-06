#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""数据库管理模块"""

import sqlite3
import os
import re
import sys
from datetime import datetime


def get_db_path():
    """获取数据库路径，兼容开发环境和打包后的多平台"""
    app_name = "排班管理系统"

    # 打包状态检测（PyInstaller）
    if getattr(sys, 'frozen', False):
        base_dir = sys._MEIPASS
    else:
        base_dir = os.path.dirname(os.path.abspath(__file__))

    # 多平台用户数据目录
    if sys.platform == 'darwin':
        user_dir = os.path.expanduser(f"~/Library/Application Support/{app_name}")
    elif sys.platform == 'win32':
        user_dir = os.path.join(os.environ.get('APPDATA', os.path.expanduser('~')), app_name)
    else:
        # Linux 等其他平台
        user_dir = os.path.expanduser(f"~/.local/share/{app_name}")

    # 确保目录存在
    os.makedirs(user_dir, exist_ok=True)

    return os.path.join(user_dir, "scheduler.db")


DB_FILE = get_db_path()


class Database:
    """数据库管理类"""

    def __init__(self):
        self.conn = sqlite3.connect(DB_FILE)
        self.conn.row_factory = sqlite3.Row
        self.create_tables()
        self.init_default_seats()
        self.init_default_rules()
        self.init_default_time_slots()
        self.init_test_data_if_empty()

    def create_tables(self):
        """创建数据表"""
        cursor = self.conn.cursor()

        # 人员表
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS persons (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT NOT NULL UNIQUE,
                level TEXT NOT NULL,
                score REAL DEFAULT 0,
                active INTEGER DEFAULT 1,
                locked INTEGER DEFAULT 0,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        ''')

        # 为已存在的数据库添加locked列（如果不存在）
        cursor.execute("PRAGMA table_info(persons)")
        columns = [col[1] for col in cursor.fetchall()]
        if 'locked' not in columns:
            cursor.execute("ALTER TABLE persons ADD COLUMN locked INTEGER DEFAULT 0")
        
        # 为已存在的数据库修改score列支持小数（如果还是INTEGER类型）
        try:
            cursor.execute("PRAGMA table_info(persons)")
            for col in cursor.fetchall():
                if col[1] == 'score' and col[2] == 'INTEGER':
                    # SQLite不支持直接修改列类型，需要创建新表迁移数据
                    cursor.execute("ALTER TABLE persons RENAME TO persons_old")
                    cursor.execute('''
                        CREATE TABLE persons (
                            id INTEGER PRIMARY KEY AUTOINCREMENT,
                            name TEXT NOT NULL UNIQUE,
                            level TEXT NOT NULL,
                            score REAL DEFAULT 0,
                            active INTEGER DEFAULT 1,
                            locked INTEGER DEFAULT 0,
                            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                            updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                        )
                    ''')
                    cursor.execute("INSERT INTO persons (id, name, level, score, active, locked, created_at, updated_at) SELECT id, name, level, score, active, locked, created_at, updated_at FROM persons_old")
                    cursor.execute("DROP TABLE persons_old")
        except Exception:
            pass  # 如果出错（比如列已经是REAL类型），忽略继续

        # 席位表（app_name作为主键，关联席位模版）
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS seats (
                app_name TEXT PRIMARY KEY,
                template_id INTEGER,
                available INTEGER DEFAULT 1,
                persons_count INTEGER DEFAULT 3,
                required_score INTEGER DEFAULT 5
            )
        ''')
        
        # 席位模版表
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS seat_templates (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT NOT NULL,
                description TEXT DEFAULT '',
                time_start TEXT DEFAULT '08:40',
                time_end TEXT DEFAULT '10:30',
                available INTEGER DEFAULT 1,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        ''')
        
        # 席位模版时段表（新结构：班次+管制+助理）
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS seat_template_time_slots (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                template_id INTEGER NOT NULL,
                shift_name TEXT NOT NULL,
                ctrl_time TEXT DEFAULT '',
                ctrl_position TEXT DEFAULT '',
                asst_time TEXT DEFAULT '',
                asst_position TEXT DEFAULT '',
                row_order INTEGER DEFAULT 0,
                FOREIGN KEY (template_id) REFERENCES seat_templates(id) ON DELETE CASCADE
            )
        ''')

        # 席位模版位置表（存储位置顺序）
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS seat_template_positions (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                template_id INTEGER NOT NULL,
                position_name TEXT NOT NULL,
                position_index INTEGER DEFAULT 0,
                FOREIGN KEY (template_id) REFERENCES seat_templates(id) ON DELETE CASCADE
            )
        ''')

        # 初始化默认席位模版
        cursor.execute("SELECT COUNT(*) FROM seat_templates")
        if cursor.fetchone()[0] == 0:
            default_templates = [
                (1, "3人标准模版", "08:40", "10:30"),
                (2, "4人标准模版", "08:50", "10:40"),
                (3, "APP02专用模版", "08:40", "10:30"),
                (4, "APP04专用模版", "08:50", "10:40"),
            ]
            cursor.executemany(
                "INSERT OR IGNORE INTO seat_templates (id, name, time_start, time_end) VALUES (?, ?, ?, ?)",
                default_templates
            )
            # 默认位置：ABC
            for tpl_id in [1, 3]:
                for i, pos in enumerate(["A", "B", "C"]):
                    cursor.execute(
                        "INSERT INTO seat_template_positions (template_id, position_name, position_index) VALUES (?, ?, ?)",
                        (tpl_id, pos, i)
                    )
            for tpl_id in [2, 4]:
                for i, pos in enumerate(["A", "B", "C", "D"]):
                    cursor.execute(
                        "INSERT INTO seat_template_positions (template_id, position_name, position_index) VALUES (?, ?, ?)",
                        (tpl_id, pos, i)
                    )
            self.conn.commit()

        # 为已存在的数据库添加last_seat_app_name列（如果不存在）
        cursor.execute("PRAGMA table_info(persons)")
        columns = [col[1] for col in cursor.fetchall()]
        if 'last_seat_app_name' not in columns:
            cursor.execute("ALTER TABLE persons ADD COLUMN last_seat_app_name TEXT")
        
        # 清理旧的last_seat_id列（如果存在）
        if 'last_seat_id' in columns:
            cursor.execute("CREATE TEMP TABLE IF NOT EXISTS temp_persons AS SELECT * FROM persons")
            # 注意：这里我们不删除旧列，只是代码不再使用它

        # 人员执勤统计表（每人每席位执勤次数）
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS person_duty_stats (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                person_id INTEGER NOT NULL,
                app_name TEXT NOT NULL,
                duty_count INTEGER DEFAULT 0,
                UNIQUE(person_id, app_name),
                FOREIGN KEY (person_id) REFERENCES persons(id) ON DELETE CASCADE
            )
        ''')

        # 为已存在的数据库添加app_name列（如果不存在）
        cursor.execute("PRAGMA table_info(seats)")
        columns = [col[1] for col in cursor.fetchall()]
        if 'app_name' not in columns:
            cursor.execute("ALTER TABLE seats ADD COLUMN app_name TEXT")

        # 规则表
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS rules (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT NOT NULL UNIQUE,
                description TEXT,
                score_modifier REAL DEFAULT 0,
                active INTEGER DEFAULT 1,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        ''')

        # 为已存在的数据库添加规则名称列（用于显示标签）
        cursor.execute("PRAGMA table_info(rules)")
        columns = [col[1] for col in cursor.fetchall()]
        if 'display_name' not in columns:
            cursor.execute("ALTER TABLE rules ADD COLUMN display_name TEXT")

        # 人员规则关联表
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS person_rules (
                person_id INTEGER,
                rule_id INTEGER,
                enabled INTEGER DEFAULT 1,
                PRIMARY KEY (person_id, rule_id),
                FOREIGN KEY (person_id) REFERENCES persons(id) ON DELETE CASCADE,
                FOREIGN KEY (rule_id) REFERENCES rules(id) ON DELETE CASCADE
            )
        ''')

        # 排班历史记录表
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS schedule_history (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                schedule_date DATE NOT NULL,
                person_id INTEGER NOT NULL,
                seat_id TEXT,
                app_name TEXT NOT NULL,
                time_slot TEXT NOT NULL,
                duration_minutes INTEGER DEFAULT 0,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                FOREIGN KEY (person_id) REFERENCES persons(id) ON DELETE CASCADE
            )
        ''')

        # 时间段配置表
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS time_slots (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT NOT NULL,
                start_time TEXT NOT NULL,
                end_time TEXT NOT NULL,
                duration_minutes INTEGER NOT NULL,
                shift_type TEXT NOT NULL,
                seat_position TEXT NOT NULL
            )
        ''')

        # 应用设置表
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS settings (
                key TEXT PRIMARY KEY,
                value TEXT
            )
        ''')

        self.conn.commit()

    def get_setting(self, key, default=None):
        """获取设置值"""
        cursor = self.conn.cursor()
        cursor.execute("SELECT value FROM settings WHERE key = ?", (key,))
        row = cursor.fetchone()
        return row[0] if row else default

    def set_setting(self, key, value):
        """设置值"""
        cursor = self.conn.cursor()
        cursor.execute(
            "INSERT OR REPLACE INTO settings (key, value) VALUES (?, ?)",
            (key, str(value))
        )
        self.conn.commit()

    def init_default_seats(self):
        """初始化默认物理席位"""
        cursor = self.conn.cursor()
        cursor.execute("SELECT COUNT(*) FROM seats")
        if cursor.fetchone()[0] == 0:
            default_seats = [
                ("APP01", 1, 3, 5),
                ("APP02", 1, 3, 5),
                ("APP03", 1, 3, 5),
                ("APP04", 1, 4, 5),
                ("APP05", 1, 3, 5),
            ]
            cursor.executemany(
                "INSERT OR IGNORE INTO seats (app_name, available, persons_count, required_score) VALUES (?, ?, ?, ?)",
                default_seats
            )
            self.conn.commit()
    
    def init_test_data_if_empty(self):
        """初始化测试数据（如果数据库为空）"""
        cursor = self.conn.cursor()
        cursor.execute("SELECT COUNT(*) FROM persons")
        if cursor.fetchone()[0] == 0:
            # 等级对应的分数
            level_scores = {"C1": 1, "C2": 2, "C3": 3, "I": 3, "S": 3}
            
            # 添加测试人员 - C1/C2/C3/I/S 各等级
            test_persons = [
                ("张伟", "C1", 1), ("李强", "C1", 1),
                ("王芳", "C2", 1), ("刘洋", "C2", 1), ("陈明", "C2", 1),
                ("赵敏", "C3", 1), ("黄磊", "C3", 1), ("周涛", "C3", 1), ("吴静", "C3", 1),
                ("郑凯", "I", 1), ("孙悦", "I", 1), ("钱磊", "I", 1), ("赵强", "I", 1), ("王涛", "I", 1),
                ("冯琳", "S", 1), ("陈晨", "S", 1), ("林雪", "S", 1),
            ]
            # 添加分数
            test_data = [(name, level, level_scores.get(level, 0), active) for name, level, active in test_persons]
            cursor.executemany(
                "INSERT INTO persons (name, level, score, active) VALUES (?, ?, ?, ?)",
                test_data
            )
            self.conn.commit()

    def get_all_persons(self):
        """获取所有人员"""
        cursor = self.conn.cursor()
        cursor.execute("SELECT * FROM persons ORDER BY level, name")
        return cursor.fetchall()
    
    def get_active_persons(self):
        """获取参与排班的人员"""
        cursor = self.conn.cursor()
        cursor.execute("SELECT * FROM persons WHERE active = 1 ORDER BY level, name")
        return cursor.fetchall()

    def get_available_seats(self):
        """获取可用的物理席位"""
        cursor = self.conn.cursor()
        cursor.execute("SELECT * FROM seats WHERE available = 1 AND app_name IS NOT NULL AND app_name != '' AND app_name != '空' ORDER BY app_name")
        rows = cursor.fetchall()
        # 转换为字典
        return [dict(row) for row in rows]
    
    def get_all_seats(self):
        """获取所有物理席位"""
        cursor = self.conn.cursor()
        cursor.execute("SELECT * FROM seats ORDER BY app_name")
        rows = cursor.fetchall()
        # 转换为字典，因为sqlite3.Row不支持get方法
        return [dict(row) for row in rows]
    
    # ===== 席位模版相关方法 =====
    def get_all_templates(self):
        """获取所有席位模版"""
        cursor = self.conn.cursor()
        cursor.execute("SELECT * FROM seat_templates ORDER BY id")
        return [dict(row) for row in cursor.fetchall()]
    
    def get_next_template_id(self):
        """获取下一个模版ID（数字）"""
        cursor = self.conn.cursor()
        cursor.execute("SELECT MAX(id) FROM seat_templates")
        result = cursor.fetchone()[0]
        if result is None:
            return 1
        # 返回最大ID+1（保持为数字）
        return int(result) + 1
    
    def get_template(self, template_id):
        """获取单个席位模版"""
        cursor = self.conn.cursor()
        cursor.execute("SELECT * FROM seat_templates WHERE id = ?", (template_id,))
        row = cursor.fetchone()
        return dict(row) if row else None
    
    def add_template(self, template_id, name, description="", time_start="08:40", time_end="10:30", available=1):
        """添加席位模版"""
        cursor = self.conn.cursor()
        try:
            cursor.execute(
                "INSERT INTO seat_templates (id, name, description, time_start, time_end, available) VALUES (?, ?, ?, ?, ?, ?)",
                (template_id, name, description, time_start, time_end, available)
            )
            self.conn.commit()
            return True, "模版添加成功"
        except sqlite3.IntegrityError as e:
            return False, f"模版已存在: {e}"
    
    def update_template(self, template_id, name=None, description=None, time_start=None, time_end=None, new_id=None, available=None):
        """更新席位模版
        template_id: 旧ID（用于定位）
        new_id: 新ID（如果ID变更）"""
        cursor = self.conn.cursor()
        
        # 如果ID变更了（注意类型可能不同：int vs str，需要统一比较）
        if new_id is not None and int(new_id) != int(template_id):
            # 检查新ID是否已存在
            cursor.execute("SELECT id FROM seat_templates WHERE id = ?", (new_id,))
            if cursor.fetchone():
                return False, f"新ID {new_id} 已存在"
            
            # 更新关联的时段配置
            cursor.execute("UPDATE seat_template_time_slots SET template_id = ? WHERE template_id = ?", (new_id, template_id))
            
            # 更新模板记录
            cursor.execute("UPDATE seat_templates SET id = ? WHERE id = ?", (new_id, template_id))
            template_id = new_id
        
        if name is not None:
            cursor.execute("UPDATE seat_templates SET name = ? WHERE id = ?", (name, template_id))
        if description is not None:
            cursor.execute("UPDATE seat_templates SET description = ? WHERE id = ?", (description, template_id))
        if time_start:
            cursor.execute("UPDATE seat_templates SET time_start = ? WHERE id = ?", (time_start, template_id))
        if time_end:
            cursor.execute("UPDATE seat_templates SET time_end = ? WHERE id = ?", (time_end, template_id))
        if available is not None:
            cursor.execute("UPDATE seat_templates SET available = ? WHERE id = ?", (available, template_id))

        self.conn.commit()
        return True, "模版更新成功"

    def delete_template(self, template_id):
        """删除席位模版"""
        cursor = self.conn.cursor()
        cursor.execute("DELETE FROM seat_templates WHERE id = ?", (template_id,))
        self.conn.commit()
        return True, "模版删除成功"

    def update_template_available(self, template_id, available):
        """更新模版可用状态"""
        cursor = self.conn.cursor()
        cursor.execute("UPDATE seat_templates SET available = ? WHERE id = ?", (available, template_id))
        self.conn.commit()
        return True, "状态更新成功"

    def get_template_time_slots(self, template_id):
        """获取席位模版的时段配置"""
        cursor = self.conn.cursor()
        cursor.execute("""
            SELECT id, shift_name, ctrl_time, ctrl_position, asst_time, asst_position, row_order 
            FROM seat_template_time_slots WHERE template_id = ? ORDER BY row_order""",
            (template_id,)
        )
        return [dict(row) for row in cursor.fetchall()]

    def set_template_time_slots(self, template_id, slots):
        """设置席位模版的时段配置"""
        cursor = self.conn.cursor()
        cursor.execute("DELETE FROM seat_template_time_slots WHERE template_id = ?", (template_id,))
        for i, slot in enumerate(slots):
            cursor.execute("""
                INSERT INTO seat_template_time_slots 
                (template_id, shift_name, ctrl_time, ctrl_position, asst_time, asst_position, row_order)
                VALUES (?, ?, ?, ?, ?, ?, ?)""",
                (template_id, slot.get('shift_name', ''), slot.get('ctrl_time', ''), slot.get('ctrl_position', ''),
                 slot.get('asst_time', ''), slot.get('asst_position', ''), i)
            )
        self.conn.commit()
        return True, "时段设置成功"

    # 旧方法保留兼容
    def get_template_positions(self, template_id):
        """获取席位模版的位置顺序（兼容旧方法）"""
        slots = self.get_template_time_slots(template_id)
        # 从slots提取位置
        positions = []
        for slot in slots:
            if slot.get('ctrl_position') and slot['ctrl_position'] not in positions:
                positions.append({'position_name': slot['ctrl_position'], 'position_index': len(positions)})
            if slot.get('asst_position') and slot['asst_position'] not in [p['position_name'] for p in positions]:
                positions.append({'position_name': slot['asst_position'], 'position_index': len(positions)})
        return positions

    def set_template_positions(self, template_id, positions):
        """设置席位模版的位置顺序（兼容旧方法）"""
        # 兼容旧调用，转换为新结构
        slots = [{'shift_name': '早班', 'ctrl_time': '08:40', 'ctrl_position': positions[0] if len(positions) > 0 else '',
                'asst_time': '08:40', 'asst_position': positions[1] if len(positions) > 1 else ''}]
        return self.set_template_time_slots(template_id, slots)

    def update_seat_available(self, app_name, available):
        """更新物理席位可用状态"""
        cursor = self.conn.cursor()
        cursor.execute("UPDATE seats SET available = ? WHERE app_name = ?", (available, app_name))
        self.conn.commit()

    def update_seat_persons_count(self, app_name, count):
        """更新物理席位人数"""
        cursor = self.conn.cursor()
        cursor.execute("UPDATE seats SET persons_count = ? WHERE app_name = ?", (count, app_name))
        self.conn.commit()

    def update_seat_required_score(self, app_name, required_score):
        """更新物理席位要求分数"""
        cursor = self.conn.cursor()
        cursor.execute("UPDATE seats SET required_score = ? WHERE app_name = ?", (required_score, app_name))
        self.conn.commit()
    
    def update_seat_template(self, app_name, template_id):
        """更新物理席位关联的模版ID"""
        cursor = self.conn.cursor()
        cursor.execute("UPDATE seats SET template_id = ? WHERE app_name = ?", (template_id, app_name))
        self.conn.commit()

    def update_seat_app_name(self, app_name, available=None, persons_count=None, required_score=None, template_id=None):
        """更新物理席位配置（app_name是主键，不能修改）"""
        cursor = self.conn.cursor()
        if available is not None:
            cursor.execute("UPDATE seats SET available = ? WHERE app_name = ?", (available, app_name))
        if persons_count is not None:
            cursor.execute("UPDATE seats SET persons_count = ? WHERE app_name = ?", (persons_count, app_name))
        if required_score is not None:
            cursor.execute("UPDATE seats SET required_score = ? WHERE app_name = ?", (required_score, app_name))
        if template_id is not None:
            cursor.execute("UPDATE seats SET template_id = ? WHERE app_name = ?", (template_id, app_name))
        self.conn.commit()
    
    def add_seat(self, app_name, template_id=None, available=1, persons_count=3, required_score=5):
        """添加物理席位"""
        cursor = self.conn.cursor()
        try:
            cursor.execute(
                "INSERT INTO seats (app_name, template_id, available, persons_count, required_score) VALUES (?, ?, ?, ?, ?)",
                (app_name, template_id, available, persons_count, required_score)
            )
            self.conn.commit()
            return True, "席位添加成功"
        except sqlite3.IntegrityError as e:
            return False, f"席位已存在: {e}"
    
    def delete_seat(self, app_name):
        """删除物理席位"""
        cursor = self.conn.cursor()
        cursor.execute("DELETE FROM seats WHERE app_name = ?", (app_name,))
        self.conn.commit()
        return True, "席位删除成功"

    def add_person(self, name, level, active=1):
        """添加人员"""
        try:
            # 输入验证
            name = name.strip()
            if not name:
                return False, "姓名不能为空"
            if len(name) < 2:
                return False, "姓名至少需要2个字符"
            if len(name) > 20:
                return False, "姓名不能超过20个字符"
            if not re.match(r'^[\u4e00-\u9fa5a-zA-Z0-9_]+$', name):
                return False, "姓名只能包含中英文、数字和下划线"

            level_scores = {"C1": 1, "C2": 2, "C3": 3, "I": 3, "S": 3}
            score = level_scores.get(level, 0)
            cursor = self.conn.cursor()
            cursor.execute('''
                INSERT INTO persons (name, level, active, score)
                VALUES (?, ?, ?, ?)
            ''', (name, level, active, score))
            self.conn.commit()
            return True, "添加成功"
        except sqlite3.IntegrityError:
            return False, "人员已存在"
        except Exception as e:
            return False, str(e)

    def update_person(self, person_id, name, level, active, score=None, locked=None, last_seat_app_name=None):
        """更新人员信息"""
        try:
            cursor = self.conn.cursor()
            if score is not None and locked is not None and last_seat_app_name is not None:
                cursor.execute('''
                    UPDATE persons
                    SET name = ?, level = ?, active = ?, score = ?, locked = ?, last_seat_app_name = ?, updated_at = CURRENT_TIMESTAMP
                    WHERE id = ?
                ''', (name, level, active, score, locked, last_seat_app_name, person_id))
            elif score is not None:
                cursor.execute('''
                    UPDATE persons
                    SET name = ?, level = ?, active = ?, score = ?, updated_at = CURRENT_TIMESTAMP
                    WHERE id = ?
                ''', (name, level, active, score, person_id))
            elif locked is not None:
                cursor.execute('''
                    UPDATE persons
                    SET name = ?, level = ?, active = ?, locked = ?, updated_at = CURRENT_TIMESTAMP
                    WHERE id = ?
                ''', (name, level, active, locked, person_id))
            else:
                cursor.execute('''
                    UPDATE persons
                    SET name = ?, level = ?, active = ?, updated_at = CURRENT_TIMESTAMP
                    WHERE id = ?
                ''', (name, level, active, person_id))
            self.conn.commit()
            return True, "更新成功"
        except sqlite3.IntegrityError:
            return False, "人员已存在"
        except Exception as e:
            return False, str(e)

    def update_person_locked(self, person_id, locked):
        """更新人员锁定状态"""
        cursor = self.conn.cursor()
        cursor.execute("UPDATE persons SET locked = ?, updated_at = CURRENT_TIMESTAMP WHERE id = ?", (locked, person_id))
        self.conn.commit()

    def delete_person(self, person_id):
        """删除人员"""
        try:
            cursor = self.conn.cursor()
            cursor.execute("DELETE FROM persons WHERE id=?", (person_id,))
            self.conn.commit()
            return True, "删除成功"
        except Exception as e:
            return False, str(e)

    def get_person_by_id(self, person_id):
        """根据ID获取人员"""
        cursor = self.conn.cursor()
        cursor.execute("SELECT * FROM persons WHERE id=?", (person_id,))
        return cursor.fetchone()

    def get_person_by_name(self, name):
        """根据姓名获取人员"""
        cursor = self.conn.cursor()
        cursor.execute("SELECT * FROM persons WHERE name=?", (name,))
        return cursor.fetchone()

    def update_person_active(self, person_id, active):
        """更新人员参与排班状态"""
        cursor = self.conn.cursor()
        cursor.execute("UPDATE persons SET active = ? WHERE id = ?", (active, person_id))
        self.conn.commit()

    def search_persons(self, keyword):
        """搜索人员"""
        cursor = self.conn.cursor()
        cursor.execute('''
            SELECT * FROM persons 
            WHERE name LIKE ?
            ORDER BY level, name
        ''', (f'%{keyword}%',))
        return cursor.fetchall()

    def get_statistics(self):
        """获取统计信息"""
        cursor = self.conn.cursor()
        
        cursor.execute("SELECT COUNT(*) as total FROM persons")
        total = cursor.fetchone()['total']
        
        cursor.execute("SELECT COUNT(*) as active FROM persons WHERE active = 1")
        active = cursor.fetchone()['active']
        
        cursor.execute("SELECT level, COUNT(*) as count FROM persons WHERE active = 1 GROUP BY level")
        level_stats = {row['level']: row['count'] for row in cursor.fetchall()}
        
        cursor.execute("SELECT COUNT(*) as available FROM seats WHERE available = 1")
        available_seats = cursor.fetchone()['available']
        
        return {
            'total': total,
            'active': active,
            'c1': level_stats.get('C1', 0),
            'c2': level_stats.get('C2', 0),
            'c3': level_stats.get('C3', 0),
            'i': level_stats.get('I', 0),
            's': level_stats.get('S', 0),
            'available_seats': available_seats
        }

    # ==================== 执勤数据管理 ====================

    def get_person_duty_stats(self, person_id=None):
        """获取人员执勤统计，person_id为None时获取所有"""
        cursor = self.conn.cursor()
        if person_id:
            cursor.execute('''
                SELECT pds.*, p.name as person_name, s.app_name
                FROM person_duty_stats pds
                JOIN persons p ON p.id = pds.person_id
                JOIN seats s ON s.app_name = pds.app_name
                WHERE pds.person_id = ?
                ORDER BY p.name, s.app_name
            ''', (person_id,))
        else:
            cursor.execute('''
                SELECT pds.*, p.name as person_name, p.level, p.last_seat_app_name,
                       s.app_name
                FROM person_duty_stats pds
                JOIN persons p ON p.id = pds.person_id
                JOIN seats s ON s.app_name = pds.app_name
                ORDER BY p.name, s.app_name
            ''')
        return cursor.fetchall()

    def get_all_persons_with_duty_stats(self):
        """获取所有人员及其执勤统计和上次席位"""
        cursor = self.conn.cursor()
        cursor.execute('''
            SELECT p.*, p.last_seat_app_name as last_seat_app
            FROM persons p
            ORDER BY p.name
        ''')
        persons = cursor.fetchall()

        # 获取每个人员的执勤统计
        result = []
        for person in persons:
            p_dict = dict(person)
            cursor.execute('''
                SELECT app_name, duty_count
                FROM person_duty_stats
                WHERE person_id = ?
            ''', (person['id'],))
            p_dict['duty_stats'] = {row['app_name']: row['duty_count'] for row in cursor.fetchall()}
            result.append(p_dict)
        return result

    def update_person_last_seat(self, person_id, app_name):
        """更新人员上次安排的席位"""
        cursor = self.conn.cursor()
        cursor.execute("UPDATE persons SET last_seat_app_name = ? WHERE id = ?", (app_name, person_id))
        self.conn.commit()

    def increment_duty_count(self, person_id, app_name, count=1):
        """增加人员某席位执勤次数"""
        cursor = self.conn.cursor()
        cursor.execute('''
            INSERT INTO person_duty_stats (person_id, app_name, duty_count)
            VALUES (?, ?, ?)
            ON CONFLICT(person_id, app_name) DO UPDATE SET duty_count = duty_count + ?
        ''', (person_id, app_name, count, count))
        self.conn.commit()

    def set_duty_count(self, person_id, app_name, count):
        """设置人员某席位执勤次数"""
        cursor = self.conn.cursor()
        cursor.execute('''
            INSERT INTO person_duty_stats (person_id, app_name, duty_count)
            VALUES (?, ?, ?)
            ON CONFLICT(person_id, app_name) DO UPDATE SET duty_count = ?
        ''', (person_id, app_name, count, count))
        self.conn.commit()

    def set_person_last_seat(self, person_id, app_name):
        """设置人员上次安排的席位"""
        cursor = self.conn.cursor()
        cursor.execute("UPDATE persons SET last_seat_app_name = ? WHERE id = ?", (app_name, person_id))
        self.conn.commit()

    def clear_person_duty_stats(self, person_id):
        """清空人员所有执勤统计数据"""
        cursor = self.conn.cursor()
        cursor.execute("DELETE FROM person_duty_stats WHERE person_id = ?", (person_id,))
        cursor.execute("UPDATE persons SET last_seat_app_name = NULL WHERE id = ?", (person_id,))
        self.conn.commit()

    # ==================== 规则管理 ====================

    def get_all_rules(self):
        """获取所有规则"""
        cursor = self.conn.cursor()
        cursor.execute("SELECT * FROM rules ORDER BY id")
        return cursor.fetchall()

    def get_active_rules(self):
        """获取启用的规则"""
        cursor = self.conn.cursor()
        cursor.execute("SELECT * FROM rules WHERE active = 1 ORDER BY id")
        return cursor.fetchall()

    def add_rule(self, name, description, score_modifier):
        """添加规则"""
        try:
            cursor = self.conn.cursor()
            cursor.execute('''
                INSERT INTO rules (name, description, score_modifier, active)
                VALUES (?, ?, ?, 1)
            ''', (name, description, score_modifier))
            self.conn.commit()
            return True, "添加成功"
        except sqlite3.IntegrityError:
            return False, "规则已存在"
        except Exception as e:
            return False, str(e)

    def update_rule(self, rule_id, name, description, score_modifier, active):
        """更新规则"""
        try:
            cursor = self.conn.cursor()
            cursor.execute('''
                UPDATE rules
                SET name = ?, description = ?, score_modifier = ?, active = ?
                WHERE id = ?
            ''', (name, description, score_modifier, active, rule_id))
            self.conn.commit()
            return True, "更新成功"
        except sqlite3.IntegrityError:
            return False, "规则已存在"
        except Exception as e:
            return False, str(e)

    def delete_rule(self, rule_id):
        """删除规则"""
        try:
            cursor = self.conn.cursor()
            cursor.execute("DELETE FROM rules WHERE id = ?", (rule_id,))
            self.conn.commit()
            return True, "删除成功"
        except Exception as e:
            return False, str(e)

    def get_rule_by_id(self, rule_id):
        """获取规则"""
        cursor = self.conn.cursor()
        cursor.execute("SELECT * FROM rules WHERE id = ?", (rule_id,))
        return cursor.fetchone()

    def get_person_rules(self, person_id):
        """获取人员的规则"""
        cursor = self.conn.cursor()
        cursor.execute('''
            SELECT r.*, pr.enabled
            FROM rules r
            LEFT JOIN person_rules pr ON r.id = pr.rule_id AND pr.person_id = ?
            ORDER BY r.id
        ''', (person_id,))
        return cursor.fetchall()

    def set_person_rule(self, person_id, rule_id, enabled):
        """设置人员的规则"""
        try:
            cursor = self.conn.cursor()
            cursor.execute('''
                INSERT OR REPLACE INTO person_rules (person_id, rule_id, enabled)
                VALUES (?, ?, ?)
            ''', (person_id, rule_id, enabled))
            self.conn.commit()
            return True
        except Exception as e:
            return False

    def get_person_score_modifier(self, person_id):
        """获取人员的规则分数修改值"""
        cursor = self.conn.cursor()
        cursor.execute('''
            SELECT COALESCE(SUM(r.score_modifier), 0) as total
            FROM person_rules pr
            JOIN rules r ON pr.rule_id = r.id
            WHERE pr.person_id = ? AND pr.enabled = 1 AND r.active = 1
        ''', (person_id,))
        return cursor.fetchone()['total']

    def init_default_rules(self):
        """初始化默认规则（如果为空）"""
        cursor = self.conn.cursor()
        cursor.execute("SELECT COUNT(*) FROM rules")
        if cursor.fetchone()[0] == 0:
            # 6个默认规则：生病/培训/出差/值班超时/婚假/丧假
            default_rules = [
                ("生病", "因生病扣除分数", -0.1),
                ("培训", "参加培训扣除分数", -0.15),
                ("出差", "因出差扣除分数", -0.1),
                ("值班超时", "值班超时奖励", 0.1),
                ("婚假", "休婚假扣除分数", -0.2),
                ("丧假", "休丧假扣除分数", -0.3),
            ]
            cursor.executemany(
                "INSERT INTO rules (name, description, score_modifier) VALUES (?, ?, ?)",
                default_rules
            )
            self.conn.commit()

    def init_default_time_slots(self):
        """初始化默认时间段（如果为空）"""
        cursor = self.conn.cursor()
        cursor.execute("SELECT COUNT(*) FROM time_slots")
        if cursor.fetchone()[0] == 0:
            # 从Excel样表解析的时间段
            default_slots = [
                # 早班 - APP04
                ("早班-时段1", "0850", "1040", 110, "早班", "APP04-管制席"),
                ("早班-时段2", "1030", "1200", 90, "早班", "APP04-管制席"),
                ("早班-时段3", "1150", "1240", 50, "早班", "APP04-管制席"),
                # 早班 - APP04 助理席
                ("早班-时段1", "0840", "1000", 80, "早班", "APP04-助理席"),
                ("早班-时段2", "0950", "1120", 90, "早班", "APP04-助理席"),
                ("早班-时段3", "1110", "1230", 80, "早班", "APP04-助理席"),
                # 早班 - APP05
                ("早班-时段1", "0840", "1030", 110, "早班", "APP05-管制席"),
                ("早班-时段2", "1020", "1150", 90, "早班", "APP05-管制席"),
                ("早班-时段3", "1140", "1230", 50, "早班", "APP05-管制席"),
                # 早班 - APP05 助理席
                ("早班-时段1", "0850", "0950", 60, "早班", "APP05-助理席"),
                ("早班-时段2", "0940", "1110", 90, "早班", "APP05-助理席"),
                ("早班-时段3", "1100", "1240", 100, "早班", "APP05-助理席"),
                # 晚班 - APP04
                ("晚班-时段1", "1800", "1940", 100, "晚班", "APP04-管制席"),
                ("晚班-时段2", "1930", "2100", 90, "晚班", "APP04-管制席"),
                ("晚班-时段3", "2050", "2220", 90, "晚班", "APP04-管制席"),
                # 晚班 - APP04 助理席
                ("晚班-时段1", "1750", "1900", 70, "晚班", "APP04-助理席"),
                ("晚班-时段2", "1850", "2020", 90, "晚班", "APP04-助理席"),
                ("晚班-时段3", "2010", "2140", 90, "晚班", "APP04-助理席"),
                ("晚班-时段4", "2130", "2210", 40, "晚班", "APP04-助理席"),
                # 晚班 - APP05
                ("晚班-时段1", "1750", "1930", 100, "晚班", "APP05-管制席"),
                ("晚班-时段2", "1920", "2050", 90, "晚班", "APP05-管制席"),
                ("晚班-时段3", "2040", "2210", 90, "晚班", "APP05-管制席"),
                # 晚班 - APP05 助理席
                ("晚班-时段1", "1800", "1850", 50, "晚班", "APP05-助理席"),
                ("晚班-时段2", "1840", "2010", 90, "晚班", "APP05-助理席"),
                ("晚班-时段3", "2000", "2130", 90, "晚班", "APP05-助理席"),
                ("晚班-时段4", "2120", "2220", 60, "晚班", "APP05-助理席"),
                # 下午班 - APP04
                ("下午班-时段1", "1230", "1400", 90, "下午班", "APP04-管制席"),
                ("下午班-时段2", "1350", "1550", 120, "下午班", "APP04-管制席"),
                ("下午班-时段3", "1540", "1730", 110, "下午班", "APP04-管制席"),
                ("下午班-时段4", "1720", "1810", 50, "下午班", "APP04-管制席"),
                # 下午班 - APP04 助理席
                ("下午班-时段1", "1220", "1320", 60, "下午班", "APP04-助理席"),
                ("下午班-时段2", "1310", "1500", 110, "下午班", "APP04-助理席"),
                ("下午班-时段3", "1450", "1650", 120, "下午班", "APP04-助理席"),
                ("下午班-时段4", "1640", "1800", 80, "下午班", "APP04-助理席"),
                # 下午班 - APP05
                ("下午班-时段1", "1220", "1350", 90, "下午班", "APP05-管制席"),
                ("下午班-时段2", "1340", "1540", 120, "下午班", "APP05-管制席"),
                ("下午班-时段3", "1530", "1720", 110, "下午班", "APP05-管制席"),
                ("下午班-时段4", "1710", "1800", 50, "下午班", "APP05-管制席"),
                # 下午班 - APP05 助理席
                ("下午班-时段1", "1230", "1310", 40, "下午班", "APP05-助理席"),
                ("下午班-时段2", "1300", "1450", 110, "下午班", "APP05-助理席"),
                ("下午班-时段3", "1440", "1640", 120, "下午班", "APP05-助理席"),
                ("下午班-时段4", "1630", "1810", 100, "下午班", "APP05-助理席"),
            ]
            cursor.executemany(
                "INSERT INTO time_slots (name, start_time, end_time, duration_minutes, shift_type, seat_position) VALUES (?, ?, ?, ?, ?, ?)",
                default_slots
            )
            self.conn.commit()

    def get_person_seat_history(self, person_id, limit=10):
        """获取人员在各席位的最近排班历史（用于均衡）"""
        cursor = self.conn.cursor()
        cursor.execute('''
            SELECT s.app_name,
                   COUNT(*) as count, SUM(duration_minutes) as total_minutes
            FROM schedule_history h
            JOIN seats s ON h.app_name = s.app_name
            WHERE h.person_id = ?
            GROUP BY s.app_name
            ORDER BY count DESC
        ''', (person_id,))
        return cursor.fetchall()

    def get_person_total_count(self, person_id, days=30):
        """获取人员最近N天的执勤次数"""
        cursor = self.conn.cursor()
        cursor.execute('''
            SELECT COALESCE(SUM(duration_minutes), 0) as total
            FROM schedule_history 
            WHERE person_id = ? AND schedule_date >= date('now', '-' || ? || ' days')
        ''', (person_id, days))
        return cursor.fetchone()['total']

    def get_seat_assignment_count(self, app_name, days=30):
        """获取席位最近N天的安排次数"""
        cursor = self.conn.cursor()
        cursor.execute('''
            SELECT COUNT(*) as cnt FROM schedule_history 
            WHERE app_name = ? AND schedule_date >= date('now', '-' || ? || ' days')
        ''', (app_name, days))
        return cursor.fetchone()['cnt']

    def save_schedule(self, schedule_date, assignments):
        """保存排班结果到历史记录
        assignments: [(person_id, app_name, time_slot, duration_minutes), ...]
        同时更新 person_duty_stats 的 duty_count
        """
        cursor = self.conn.cursor()
        # 先删除当天的记录
        cursor.execute("DELETE FROM schedule_history WHERE schedule_date = ?", (schedule_date,))
        # 插入新记录，同时更新 person_duty_stats
        for record in assignments:
            if len(record) == 5:
                person_id, seat_id, app_name, time_slot, duration_minutes = record
            else:
                person_id, app_name, time_slot, duration_minutes = record
                seat_id = app_name  # 兼容：main_UI 场景下 seat_id 等于 app_name
            cursor.execute('''
                INSERT INTO schedule_history (schedule_date, person_id, seat_id, app_name, time_slot, duration_minutes)
                VALUES (?, ?, ?, ?, ?, ?)
            ''', (schedule_date, person_id, seat_id, app_name, time_slot, duration_minutes))
            # 更新执勤统计（每次导出算1次）
            cursor.execute('''
                INSERT INTO person_duty_stats (person_id, app_name, duty_count)
                VALUES (?, ?, 1)
                ON CONFLICT(person_id, app_name) DO UPDATE SET duty_count = duty_count + 1
            ''', (person_id, app_name))
        self.conn.commit()

    def save_schedule_with_seat_ids(self, schedule_date, assignments):
        """保存排班结果到历史记录（使用seat_id，scheduler.py导出专用）
        assignments: [(person_id, seat_id, app_name, time_slot, duration_minutes), ...]
        duration_minutes 在此上下文中实际是 duty_count（班次计数，1或2）
        同时更新 person_duty_stats 的 duty_count
        """
        cursor = self.conn.cursor()
        # 先删除当天的记录
        cursor.execute("DELETE FROM schedule_history WHERE schedule_date = ?", (schedule_date,))

        # 统计每个人每个席位的总执勤次数（因为同一人同一天同一席位可能有多条记录）
        duty_totals = {}  # {(person_id, app_name): total_duty_count}
        for person_id, seat_id, app_name, time_slot, duration_minutes in assignments:
            key = (person_id, app_name)
            duty_totals[key] = duty_totals.get(key, 0) + duration_minutes

        # 插入历史记录（注意：表结构只有 schedule_date, person_id, app_name, time_slot, duration_minutes）
        for person_id, seat_id, app_name, time_slot, duration_minutes in assignments:
            cursor.execute('''
                INSERT INTO schedule_history (schedule_date, person_id, app_name, time_slot, duration_minutes)
                VALUES (?, ?, ?, ?, ?)
            ''', (schedule_date, person_id, app_name, time_slot, duration_minutes))

        # 更新执勤统计：先删除当天的旧统计（避免重复累加），再用新统计覆盖
        for (person_id, app_name), total_duty in duty_totals.items():
            cursor.execute('''
                INSERT INTO person_duty_stats (person_id, app_name, duty_count)
                VALUES (?, ?, ?)
                ON CONFLICT(person_id, app_name) DO UPDATE SET duty_count = ?
            ''', (person_id, app_name, total_duty, total_duty))

        self.conn.commit()

    def get_time_slots(self, shift_type=None):
        """获取时间段配置"""
        cursor = self.conn.cursor()
        if shift_type:
            cursor.execute("SELECT * FROM time_slots WHERE shift_type = ? ORDER BY seat_position, start_time", (shift_type,))
        else:
            cursor.execute("SELECT * FROM time_slots ORDER BY shift_type, seat_position, start_time")
        return cursor.fetchall()

    def close(self):
        """关闭数据库连接"""
        self.conn.close()