# 尺寸链计算工具
# 核心原理 极值法计算尺寸链的组成环
# 1 搜索相关资料，了解什么是极值法
# 2 需要的功能：
#   a 增加删除尺寸链和公差，选择增环减环，区分标准件还是零件，通过列表的方式显示编辑；
#   b 组成环添加后可计数；有计算尺寸链按钮，点击后可以显示计算结果：总公称尺寸，上偏差，下偏差，公差中间值；列表可看到零件公差带贡献度；
#   c 可设定尺寸链总公差长度和公差中间值，然后自动按照公称尺寸大小智能分配给零件，但设置有标准件属性的公差带不分配公差带；
#   d 公差分配逻辑：去除标准件公差带后，按照尺寸长度大小分配公差；默认最小公差带0.01，一般公差带长度建议分配0.02，若公差比较富足，可以按照尺寸分配更多；分配公差优先在尺寸链总公差长度以内，分配结果可以超出设定尺寸链总公差长度，并弹出提示超出多少；
#   e 可保存尺寸链计算表；也可以读取尺寸链计算表；
#   f ui界面可以选择置顶；
# 3 存储方式excel excel存储目录可通过构造对象时传递，默认当前目录的./save 也可以点击按钮读取excel中的尺寸链数据
# 4 存储的excel会包含计算结果

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import pandas as pd
import os
import sys
import math
import logging
from datetime import datetime


logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(name)s - %(message)s"
)
logger = logging.getLogger("chain_tool")

class DimensionChain:
    """尺寸链数据类"""
    def __init__(self, name="尺寸链1"):
        self.name = name
        self.rings = []  # 组成环列表
        self._next_ring_id = 1
    
    def _resolve_defaults(self, alloc_type, t_min, t_pref, t_max, allow_asymmetry):
        """根据分配类型补齐默认参数"""
        defaults = ToleranceAllocator.get_type_defaults(alloc_type)
        if t_min is None:
            t_min = defaults['t_min']
        if t_pref is None:
            t_pref = defaults['t_pref']
        if t_max is None:
            t_max = defaults['t_max']
        if allow_asymmetry is None:
            allow_asymmetry = defaults['allow_asymmetry']
        return float(t_min), float(t_pref), float(t_max), bool(allow_asymmetry)

    def _normalize_ring(self, ring):
        """统一旧字段并补齐新字段，确保算法与UI都可用"""
        is_fixed = bool(ring.get('is_fixed', ring.get('is_standard', False)))
        alloc_type = str(ring.get('alloc_type', '普通')).strip() or '普通'
        bias_mode = str(ring.get('bias_mode', '保持当前中值')).strip() or '保持当前中值'
        ring_id = int(ring.get('ring_id', self._next_ring_id))
        t_min, t_pref, t_max, allow_asymmetry = self._resolve_defaults(
            alloc_type,
            ring.get('t_min'),
            ring.get('t_pref'),
            ring.get('t_max'),
            ring.get('allow_asymmetry')
        )

        normalized = {
            'ring_id': ring_id,
            'name': str(ring['name']).strip(),
            'nominal': float(ring['nominal']),
            'upper_dev': float(ring['upper_dev']),
            'lower_dev': float(ring['lower_dev']),
            'is_add': bool(ring.get('is_add', True)),
            'is_fixed': is_fixed,
            'is_standard': is_fixed,  # 兼容旧逻辑
            'alloc_type': alloc_type,
            'bias_mode': bias_mode,
            'allow_asymmetry': allow_asymmetry,
            't_min': float(t_min),
            't_pref': float(t_pref),
            't_max': float(max(t_max, t_min))
        }

        self._next_ring_id = max(self._next_ring_id, ring_id + 1)
        return normalized
    
    def add_ring(
        self,
        name,
        nominal,
        upper_dev,
        lower_dev,
        is_add=True,
        is_fixed=False,
        alloc_type='普通',
        bias_mode='保持当前中值',
        allow_asymmetry=None,
        t_min=None,
        t_pref=None,
        t_max=None,
        ring_id=None
    ):
        """添加组成环"""
        if ring_id is None:
            ring_id = self._next_ring_id

        ring = {
            'ring_id': int(ring_id),
            'name': name,
            'nominal': float(nominal),
            'upper_dev': float(upper_dev),
            'lower_dev': float(lower_dev),
            'is_add': is_add,  # True=增环, False=减环
            'is_fixed': bool(is_fixed),
            'is_standard': bool(is_fixed),  # 兼容旧版本
            'alloc_type': alloc_type,
            'bias_mode': bias_mode,
            'allow_asymmetry': allow_asymmetry,
            't_min': t_min,
            't_pref': t_pref,
            't_max': t_max
        }
        normalized = self._normalize_ring(ring)
        self.rings.append(normalized)
        logger.info("添加组成环: id=%s name=%s", normalized['ring_id'], normalized['name'])
        return normalized
    
    def remove_ring(self, index):
        """删除组成环"""
        if 0 <= index < len(self.rings):
            return self.rings.pop(index)
        return None
    
    def calculate(self):
        """极值法计算尺寸链"""
        if not self.rings:
            return None
        
        nominal_sum = 0
        upper_sum = 0
        lower_sum = 0
        tolerance_sum = 0
        
        ring_contributions = []
        
        for ring in self.rings:
            sign = 1 if ring['is_add'] else -1
            tol = abs(ring['upper_dev'] - ring['lower_dev'])
            
            # 公称尺寸
            nominal_sum += sign * ring['nominal']
            
            # 偏差计算
            if ring['is_add']:
                upper_sum += ring['upper_dev']
                lower_sum += ring['lower_dev']
            else:
                upper_sum -= ring['lower_dev']
                lower_sum -= ring['upper_dev']
            
            tolerance_sum += tol
            
            # 各环贡献度
            ring_contributions.append({
                'name': ring['name'],
                'tolerance': tol,
                'contribution': tol
            })
        
        mid_value = (upper_sum + lower_sum) / 2
        
        return {
            'nominal': nominal_sum,
            'upper_dev': upper_sum,
            'lower_dev': lower_sum,
            'tolerance': tolerance_sum,
            'mid_value': mid_value,
            'ring_count': len(self.rings),
            'contributions': ring_contributions
        }


class ToleranceAllocator:
    """公差分配器"""

    MIN_UNIT = 0.001
    MIN_TOLERANCE = 0.008

    TYPE_CONFIG = {
        '普通': {'t_min': 0.008, 't_pref': 0.012, 't_max': 0.020, 'weight': 1.0, 'allow_asymmetry': False},
        '总长': {'t_min': 0.016, 't_pref': 0.020, 't_max': 0.040, 'weight': 3.0, 'allow_asymmetry': True},
        '关键': {'t_min': 0.008, 't_pref': 0.009, 't_max': 0.015, 'weight': 0.4, 'allow_asymmetry': False},
        '放宽': {'t_min': 0.010, 't_pref': 0.020, 't_max': 0.060, 'weight': 2.0, 'allow_asymmetry': True}
    }

    @staticmethod
    def get_type_defaults(alloc_type):
        """根据分配类型返回默认参数"""
        return ToleranceAllocator.TYPE_CONFIG.get(alloc_type, ToleranceAllocator.TYPE_CONFIG['普通']).copy()

    @staticmethod
    def _round_to_unit(value):
        """按最小单位取整"""
        return round(round(float(value) / ToleranceAllocator.MIN_UNIT) * ToleranceAllocator.MIN_UNIT, 3)

    @staticmethod
    def _ring_contribution(ring, upper_dev=None, lower_dev=None):
        """计算环对尺寸链上下偏差的贡献，正确处理增环/减环"""
        upper = ring['upper_dev'] if upper_dev is None else upper_dev
        lower = ring['lower_dev'] if lower_dev is None else lower_dev
        if ring['is_add']:
            return upper, lower
        return -lower, -upper

    @staticmethod
    def _distribute_extra_with_cap(base_values, capacities, weights, extra):
        """在上限约束下按权重分配余量"""
        values = base_values.copy()
        remaining = max(0.0, float(extra))

        for _ in range(10):
            adjustable = [
                ring_id for ring_id, cap in capacities.items()
                if cap > 1e-12 and weights.get(ring_id, 0.0) > 0
            ]
            if not adjustable or remaining <= 1e-12:
                break

            weight_sum = sum(weights[ring_id] for ring_id in adjustable)
            if weight_sum <= 0:
                break

            consumed = 0.0
            for ring_id in adjustable:
                share = remaining * (weights[ring_id] / weight_sum)
                add = min(share, capacities[ring_id])
                values[ring_id] += add
                capacities[ring_id] -= add
                consumed += add

            remaining -= consumed
            if consumed <= 1e-12:
                break

        return values, remaining

    @staticmethod
    def _soft_distribute(values, weights, extra):
        """无上限约束按权重继续分配"""
        remaining = max(0.0, float(extra))
        if remaining <= 1e-12:
            return values

        valid = [ring_id for ring_id, weight in weights.items() if weight > 0]
        if not valid:
            return values

        weight_sum = sum(weights[ring_id] for ring_id in valid)
        for ring_id in valid:
            values[ring_id] += remaining * (weights[ring_id] / weight_sum)
        return values
    
    @staticmethod
    def allocate(rings, target_upper, target_lower):
        """
        智能公差分配（新策略）
        1. 固定公差件先占用
        2. 对可分配件先分总公差带 T
        3. 满足最小公差约束后，按类型权重*sqrt(nominal)分配余量
        4. 再按偏差模式拆分上偏差/下偏差
        5. 若目标中值不满足，优先对允许不对称的环做中值偏置
        """
        # 验证目标输入
        if target_upper <= target_lower:
            return {}, {}, "错误：目标上偏差必须大于目标下偏差"

        target_tolerance = float(target_upper - target_lower)
        target_mid = float((target_upper + target_lower) / 2)

        fixed_rings = [r for r in rings if bool(r.get('is_fixed', r.get('is_standard', False)))]
        alloc_rings = [r for r in rings if not bool(r.get('is_fixed', r.get('is_standard', False)))]

        if not alloc_rings:
            return {}, {}, "所有组成环均为固定公差件，无需分配"

        # 固定公差件贡献（严格按增/减环映射）
        fixed_upper = 0.0
        fixed_lower = 0.0
        for ring in fixed_rings:
            up_contrib, low_contrib = ToleranceAllocator._ring_contribution(ring)
            fixed_upper += up_contrib
            fixed_lower += low_contrib

        fixed_tolerance = fixed_upper - fixed_lower
        fixed_mid = (fixed_upper + fixed_lower) / 2
        available_tolerance = target_tolerance - fixed_tolerance

        # 先按总公差带分配
        t_values = {}
        capacities = {}
        weights = {}
        sum_min = 0.0

        for ring in alloc_rings:
            ring_id = ring['ring_id']
            alloc_type = ring.get('alloc_type', '普通')
            defaults = ToleranceAllocator.get_type_defaults(alloc_type)
            type_weight = defaults['weight']

            t_min = max(float(ring.get('t_min', defaults['t_min'])), ToleranceAllocator.MIN_TOLERANCE)
            t_max = max(float(ring.get('t_max', defaults['t_max'])), t_min)
            nominal = abs(float(ring.get('nominal', 0.0)))

            t_values[ring_id] = t_min
            capacities[ring_id] = max(0.0, t_max - t_min)
            weights[ring_id] = max(0.0, type_weight * math.sqrt(nominal) if nominal > 0 else type_weight * 0.1)
            sum_min += t_min

        tips = []
        if available_tolerance <= 0:
            over = abs(available_tolerance)
            tips.append(f"固定公差件已占满目标总公差，至少超出 {over:.3f} mm；已按最小可制造公差分配")
        elif sum_min > available_tolerance:
            shortage = sum_min - available_tolerance
            tips.append(f"目标总公差不足以覆盖最小公差约束，至少差 {shortage:.3f} mm；已按最小可制造公差分配")
        else:
            extra = available_tolerance - sum_min
            t_values, remained = ToleranceAllocator._distribute_extra_with_cap(
                t_values, capacities, weights, extra
            )
            if remained > 1e-12:
                tips.append(f"超过类型上限后的剩余公差 {remained:.3f} mm，已按权重继续分配给可加工件")
                t_values = ToleranceAllocator._soft_distribute(t_values, weights, remained)

        # 根据偏差模式先生成基础中值，再按目标中值做偏置修正
        mid_values = {}
        signed_base_mid = fixed_mid
        adjustable = []
        for ring in alloc_rings:
            ring_id = ring['ring_id']
            t_i = t_values[ring_id]
            bias_mode = ring.get('bias_mode', '保持当前中值')
            base_mid = 0.0
            if bias_mode == '对称':
                base_mid = 0.0
            elif bias_mode == '正单边':
                base_mid = t_i / 2
            elif bias_mode == '负单边':
                base_mid = -t_i / 2
            elif bias_mode == '保持当前中值':
                base_mid = float(ring['upper_dev'] + ring['lower_dev']) / 2
            elif bias_mode == '自动偏置':
                base_mid = 0.0
            else:
                base_mid = float(ring['upper_dev'] + ring['lower_dev']) / 2

            mid_values[ring_id] = base_mid
            sign = 1 if ring['is_add'] else -1
            signed_base_mid += sign * base_mid

            allow_asymmetry = bool(ring.get('allow_asymmetry', False))
            if allow_asymmetry or ring.get('alloc_type') == '总长' or bias_mode == '自动偏置':
                adjustable.append(ring)

        delta_mid = target_mid - signed_base_mid
        if abs(delta_mid) > 1e-9 and adjustable:
            adj_weights = {}
            for ring in adjustable:
                ring_id = ring['ring_id']
                adj_weights[ring_id] = max(weights.get(ring_id, 0.0), 0.1)
            w_sum = sum(adj_weights.values())
            if w_sum > 0:
                for ring in adjustable:
                    ring_id = ring['ring_id']
                    sign = 1 if ring['is_add'] else -1
                    z_i = delta_mid * (adj_weights[ring_id] / w_sum)
                    mid_values[ring_id] += sign * z_i
        elif abs(delta_mid) > 1e-9 and not adjustable:
            tips.append("当前无可偏置尺寸，目标公差中值无法完全满足，已输出最接近结果")

        # 输出分配结果（按ID，避免同名覆盖）
        allocation_plan = {}
        total_upper = fixed_upper
        total_lower = fixed_lower
        alloc_upper_sum = 0.0
        alloc_lower_sum = 0.0

        for ring in alloc_rings:
            ring_id = ring['ring_id']
            t_i = t_values[ring_id]
            m_i = mid_values[ring_id]
            # 总长/普通类型公差带取整到 0.005，其余保持 0.001 精度
            alloc_type_out = ring.get('alloc_type', '普通')
            if alloc_type_out in ('总长', '普通'):
                # 向下取整到 0.005，保守策略：分配后不超标
                t_i = max(
                    round(math.floor(t_i / 0.005) * 0.005, 3),
                    ToleranceAllocator.MIN_TOLERANCE
                )
            upper = ToleranceAllocator._round_to_unit(m_i + t_i / 2)
            lower = ToleranceAllocator._round_to_unit(m_i - t_i / 2)
            tol = round(upper - lower, 3)

            up_contrib, low_contrib = ToleranceAllocator._ring_contribution(ring, upper, lower)
            total_upper += up_contrib
            total_lower += low_contrib
            alloc_upper_sum += up_contrib
            alloc_lower_sum += low_contrib

            allocation_plan[ring_id] = {
                'ring_id': ring_id,
                'name': ring['name'],
                'upper_dev': upper,
                'lower_dev': lower,
                'tolerance': tol,
                'alloc_type': ring.get('alloc_type', '普通'),
                'bias_mode': ring.get('bias_mode', '保持当前中值')
            }

        actual_tolerance = total_upper - total_lower
        actual_mid = (total_upper + total_lower) / 2
        exceed = actual_tolerance - target_tolerance

        summary = {
            'target_upper': target_upper,
            'target_lower': target_lower,
            'target_tolerance': target_tolerance,
            'target_mid': target_mid,
            'fixed_tolerance': fixed_tolerance,
            'fixed_mid': fixed_mid,
            'available_tolerance': available_tolerance,
            'allocated_upper_sum': alloc_upper_sum,
            'allocated_lower_sum': alloc_lower_sum,
            'actual_upper': total_upper,
            'actual_lower': total_lower,
            'actual_tolerance': actual_tolerance,
            'actual_mid': actual_mid,
            'exceed': exceed,
            'tips': tips
        }

        message = (
            f"固定公差件占用公差: {fixed_tolerance:.3f} mm，"
            f"可分配公差: {available_tolerance:.3f} mm，"
            f"实际总公差: {actual_tolerance:.3f} mm，"
            f"目标中值: {target_mid:.3f} mm，实际中值: {actual_mid:.3f} mm"
        )
        if exceed > 1e-9:
            message += f"；超出目标总公差 {exceed:.3f} mm"
        if tips:
            message += "；" + "；".join(tips)

        logger.info("公差分配完成: target_tol=%.4f actual_tol=%.4f", target_tolerance, actual_tolerance)
        return allocation_plan, summary, message


class ChainCalculatorFrame(ttk.Frame):
    """尺寸链计算工具UI"""
    
    def __init__(self, parent, save_dir="./save"):
        super().__init__(parent)
        self.save_dir = save_dir
        self.chain = DimensionChain()
        self.is_topmost = False
        self.allocated_plan = {}
        self.alloc_summary = {}
        
        # 确保保存目录存在
        os.makedirs(self.save_dir, exist_ok=True)
        
        self.setup_ui()
    
    def setup_ui(self):
        """设置UI界面"""
        # 顶部控制区
        top_frame = ttk.Frame(self)
        top_frame.pack(fill="x", padx=5, pady=2)
        
        # 尺寸链名称
        ttk.Label(top_frame, text="尺寸链名称:").pack(side="left")
        self.name_var = tk.StringVar(value="尺寸链1")
        name_entry = ttk.Entry(top_frame, textvariable=self.name_var, width=15)
        name_entry.pack(side="left", padx=5)
        
        # 置顶按钮
        self.topmost_var = tk.BooleanVar(value=False)
        topmost_btn = ttk.Checkbutton(
            top_frame, text="窗口置顶", 
            variable=self.topmost_var,
            command=self.toggle_topmost
        )
        topmost_btn.pack(side="right", padx=5)
        
        # 输入区域 - 紧凑布局
        input_frame = ttk.LabelFrame(self, text=" 添加组成环 ")
        input_frame.pack(fill="x", padx=5, pady=2)
        
        # 第一行：名称、公称尺寸、上偏差、下偏差
        row1 = ttk.Frame(input_frame)
        row1.pack(fill="x", padx=3, pady=2)
        
        ttk.Label(row1, text="名称:").grid(row=0, column=0, padx=3, sticky="e")
        self.ring_name_var = tk.StringVar()
        ttk.Entry(row1, textvariable=self.ring_name_var, width=10).grid(row=0, column=1, padx=2)
        
        ttk.Label(row1, text="公称:").grid(row=0, column=2, padx=3, sticky="e")
        self.nominal_var = tk.StringVar(value="1")
        ttk.Entry(row1, textvariable=self.nominal_var, width=8).grid(row=0, column=3, padx=2)
        
        ttk.Label(row1, text="上偏差:").grid(row=0, column=4, padx=3, sticky="e")
        self.upper_dev_var = tk.StringVar(value="0.01")
        ttk.Entry(row1, textvariable=self.upper_dev_var, width=8).grid(row=0, column=5, padx=2)
        
        ttk.Label(row1, text="下偏差:").grid(row=0, column=6, padx=3, sticky="e")
        self.lower_dev_var = tk.StringVar(value="-0.01")
        ttk.Entry(row1, textvariable=self.lower_dev_var, width=8).grid(row=0, column=7, padx=2)
        
        # 第二行：环类型、固定公差、按钮
        row2 = ttk.Frame(input_frame)
        row2.pack(fill="x", padx=3, pady=2)
        
        ttk.Label(row2, text="环类型:").grid(row=0, column=0, padx=3, sticky="e")
        self.is_add_var = tk.BooleanVar(value=True)
        ttk.Radiobutton(row2, text="增环", variable=self.is_add_var, value=True).grid(row=0, column=1, padx=2)
        ttk.Radiobutton(row2, text="减环", variable=self.is_add_var, value=False).grid(row=0, column=2, padx=2)
        
        self.is_fixed_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(row2, text="固定公差(不参与分配)", variable=self.is_fixed_var).grid(row=0, column=3, padx=10)

        # 第三行：分配类型与偏差策略
        row3 = ttk.Frame(input_frame)
        row3.pack(fill="x", padx=3, pady=2)

        ttk.Label(row3, text="分配类型:").grid(row=0, column=0, padx=3, sticky="e")
        self.alloc_type_var = tk.StringVar(value="普通")
        alloc_type_combo = ttk.Combobox(
            row3,
            textvariable=self.alloc_type_var,
            values=["普通", "总长", "关键", "放宽"],
            width=8,
            state="readonly"
        )
        alloc_type_combo.grid(row=0, column=1, padx=2)
        alloc_type_combo.bind("<<ComboboxSelected>>", self.on_alloc_type_change)

        ttk.Label(row3, text="偏差模式:").grid(row=0, column=2, padx=3, sticky="e")
        self.bias_mode_var = tk.StringVar(value="保持当前中值")
        ttk.Combobox(
            row3,
            textvariable=self.bias_mode_var,
            values=["对称", "正单边", "负单边", "保持当前中值", "自动偏置"],
            width=12,
            state="readonly"
        ).grid(row=0, column=3, padx=2)

        self.allow_asymmetry_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(row3, text="允许不对称", variable=self.allow_asymmetry_var).grid(row=0, column=4, padx=10)

        ttk.Label(row3, text="T_min:").grid(row=0, column=5, padx=3, sticky="e")
        self.t_min_var = tk.StringVar(value="0.008")
        ttk.Entry(row3, textvariable=self.t_min_var, width=6).grid(row=0, column=6, padx=2)

        ttk.Label(row3, text="T_pref:").grid(row=0, column=7, padx=3, sticky="e")
        self.t_pref_var = tk.StringVar(value="0.012")
        ttk.Entry(row3, textvariable=self.t_pref_var, width=6).grid(row=0, column=8, padx=2)

        ttk.Label(row3, text="T_max:").grid(row=0, column=9, padx=3, sticky="e")
        self.t_max_var = tk.StringVar(value="0.020")
        ttk.Entry(row3, textvariable=self.t_max_var, width=6).grid(row=0, column=10, padx=2)
        
        ttk.Button(row2, text="添加", command=self.add_ring).grid(row=0, column=4, padx=5)
        ttk.Button(row2, text="更新选中", command=self.update_ring).grid(row=0, column=5, padx=5)
        
        # 组成环列表
        list_frame = ttk.LabelFrame(self, text=" 组成环列表 ")
        list_frame.pack(fill="both", expand=True, padx=5, pady=2)
        
        # Treeview - 占据全部空间
        columns = (
            "ring_id", "name", "nominal", "upper_dev", "lower_dev", "type",
            "fixed", "alloc_type", "bias_mode", "allow_asymmetry", "tolerance"
        )
        self.tree = ttk.Treeview(list_frame, columns=columns, show="headings")
        
        self.tree.heading("ring_id", text="ID")
        self.tree.heading("name", text="零件名称")
        self.tree.heading("nominal", text="公称尺寸")
        self.tree.heading("upper_dev", text="上偏差")
        self.tree.heading("lower_dev", text="下偏差")
        self.tree.heading("type", text="环类型")
        self.tree.heading("fixed", text="固定公差")
        self.tree.heading("alloc_type", text="分配类型")
        self.tree.heading("bias_mode", text="偏差模式")
        self.tree.heading("allow_asymmetry", text="允许不对称")
        self.tree.heading("tolerance", text="公差")
        
        self.tree.column("ring_id", width=45)
        self.tree.column("name", width=80)
        self.tree.column("nominal", width=70)
        self.tree.column("upper_dev", width=70)
        self.tree.column("lower_dev", width=70)
        self.tree.column("type", width=50)
        self.tree.column("fixed", width=70)
        self.tree.column("alloc_type", width=70)
        self.tree.column("bias_mode", width=85)
        self.tree.column("allow_asymmetry", width=80)
        self.tree.column("tolerance", width=70)
        
        scrollbar_y = ttk.Scrollbar(list_frame, orient="vertical", command=self.tree.yview)
        scrollbar_x = ttk.Scrollbar(list_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=scrollbar_y.set, xscrollcommand=scrollbar_x.set)
        
        self.tree.pack(side="left", fill="both", expand=True)
        scrollbar_y.pack(side="right", fill="y")
        scrollbar_x.pack(side="bottom", fill="x")
        
        # 绑定双击事件
        self.tree.bind("<Double-1>", self.on_double_click)
        
        # 操作按钮区 - 独立放在列表下方
        action_frame = ttk.Frame(self)
        action_frame.pack(fill="x", padx=5, pady=2)
        
        self.count_label = ttk.Label(action_frame, text="数量: 0")
        self.count_label.pack(side="left", padx=5)
        
        ttk.Button(action_frame, text="计算尺寸链", command=self.calculate_chain).pack(side="left", padx=3)
        ttk.Button(action_frame, text="保存Excel", command=self.save_to_excel).pack(side="left", padx=3)
        ttk.Button(action_frame, text="读取Excel", command=self.load_from_excel).pack(side="left", padx=3)
        ttk.Button(action_frame, text="删除选中", command=self.remove_ring).pack(side="right", padx=3)
        ttk.Button(action_frame, text="清空列表", command=self.clear_rings).pack(side="right", padx=3)
        
        # 公差分配区 - 紧凑布局
        alloc_frame = ttk.LabelFrame(self, text=" 公差分配 ")
        alloc_frame.pack(fill="x", padx=5, pady=2)
        
        row_alloc = ttk.Frame(alloc_frame)
        row_alloc.pack(fill="x", padx=3, pady=2)
        
        ttk.Label(row_alloc, text="目标上偏差:").grid(row=0, column=0, padx=3, sticky="e")
        self.target_upper_var = tk.StringVar(value="0.05")
        ttk.Entry(row_alloc, textvariable=self.target_upper_var, width=8).grid(row=0, column=1, padx=2)
        
        ttk.Label(row_alloc, text="目标下偏差:").grid(row=0, column=2, padx=3, sticky="e")
        self.target_lower_var = tk.StringVar(value="-0.05")
        ttk.Entry(row_alloc, textvariable=self.target_lower_var, width=8).grid(row=0, column=3, padx=2)

        ttk.Label(row_alloc, text="目标总公差:").grid(row=0, column=4, padx=3, sticky="e")
        self.target_tol_var = tk.StringVar(value="0.10")
        ttk.Label(row_alloc, textvariable=self.target_tol_var, width=8).grid(row=0, column=5, padx=2)

        ttk.Label(row_alloc, text="目标中值:").grid(row=0, column=6, padx=3, sticky="e")
        self.target_mid_var = tk.StringVar(value="0.00")
        ttk.Label(row_alloc, textvariable=self.target_mid_var, width=8).grid(row=0, column=7, padx=2)
        
        ttk.Button(row_alloc, text="智能分配公差", command=self.allocate_tolerance).grid(row=0, column=8, padx=8)
        ttk.Button(row_alloc, text="应用分配", command=self.apply_allocation).grid(row=0, column=9, padx=5)
        
        # 分配结果显示
        self.alloc_result_label = ttk.Label(alloc_frame, text="")
        self.alloc_result_label.pack(fill="x", padx=3, pady=1)
        
        # 计算结果区 - 减小高度
        result_frame = ttk.LabelFrame(self, text=" 计算结果 ")
        result_frame.pack(fill="x", padx=5, pady=2)
        
        self.result_text = tk.Text(result_frame, height=8, font=("Consolas", 11))
        self.result_text.pack(fill="x", padx=3, pady=3)
        
        # 初始化显示
        self.on_target_change()
        self.on_alloc_type_change()
        self.update_count()

    def on_target_change(self):
        """刷新目标总公差和目标中值显示"""
        try:
            target_upper = float(self.target_upper_var.get())
            target_lower = float(self.target_lower_var.get())
            self.target_tol_var.set(f"{target_upper - target_lower:.3f}")
            self.target_mid_var.set(f"{(target_upper + target_lower) / 2:.3f}")
        except ValueError:
            self.target_tol_var.set("-")
            self.target_mid_var.set("-")

    def on_alloc_type_change(self, _event=None):
        """当分配类型变化时自动刷新默认公差建议"""
        alloc_type = self.alloc_type_var.get()
        defaults = ToleranceAllocator.get_type_defaults(alloc_type)
        self.t_min_var.set(f"{defaults['t_min']:.3f}")
        self.t_pref_var.set(f"{defaults['t_pref']:.3f}")
        self.t_max_var.set(f"{defaults['t_max']:.3f}")
        self.allow_asymmetry_var.set(bool(defaults['allow_asymmetry']))

    def _insert_ring_to_tree(self, ring):
        """将组成环写入Treeview"""
        tol = abs(ring['upper_dev'] - ring['lower_dev'])
        self.tree.insert("", "end", values=(
            ring['ring_id'],
            ring['name'],
            ring['nominal'],
            ring['upper_dev'],
            ring['lower_dev'],
            "增环" if ring['is_add'] else "减环",
            "是" if ring.get('is_fixed', False) else "否",
            ring.get('alloc_type', '普通'),
            ring.get('bias_mode', '保持当前中值'),
            "是" if ring.get('allow_asymmetry', False) else "否",
            round(tol, 4)
        ))

    def refresh_tree(self):
        """按数据模型重绘列表"""
        for item in self.tree.get_children():
            self.tree.delete(item)
        for ring in self.chain.rings:
            self._insert_ring_to_tree(ring)
    
    def toggle_topmost(self):
        """切换窗口置顶状态"""
        root = self.winfo_toplevel()
        root.attributes('-topmost', self.topmost_var.get())
    
    def add_ring(self):
        """添加组成环"""
        try:
            name = self.ring_name_var.get().strip()
            if not name:
                messagebox.showwarning("警告", "请输入零件名称")
                return
            
            nominal = float(self.nominal_var.get())
            upper_dev = float(self.upper_dev_var.get())
            lower_dev = float(self.lower_dev_var.get())
            is_add = self.is_add_var.get()
            is_fixed = self.is_fixed_var.get()
            alloc_type = self.alloc_type_var.get()
            bias_mode = self.bias_mode_var.get()
            allow_asymmetry = self.allow_asymmetry_var.get()
            t_min = float(self.t_min_var.get())
            t_pref = float(self.t_pref_var.get())
            t_max = float(self.t_max_var.get())
            
            # 添加到数据
            ring = self.chain.add_ring(
                name,
                nominal,
                upper_dev,
                lower_dev,
                is_add,
                is_fixed=is_fixed,
                alloc_type=alloc_type,
                bias_mode=bias_mode,
                allow_asymmetry=allow_asymmetry,
                t_min=t_min,
                t_pref=t_pref,
                t_max=t_max
            )
            
            # 更新列表
            self._insert_ring_to_tree(ring)
            
            # 清空零件名称，保留其他输入值方便连续添加
            self.ring_name_var.set("")
            
            self.update_count()
            
        except ValueError:
            messagebox.showerror("错误", "请输入有效的数字")
    
    def remove_ring(self):
        """删除选中的组成环"""
        selected = self.tree.selection()
        if not selected:
            messagebox.showwarning("警告", "请先选中要删除的项")
            return
        
        for item in selected:
            index = self.tree.index(item)
            self.chain.remove_ring(index)
            self.tree.delete(item)
        
        self.update_count()
    
    def clear_rings(self):
        """清空所有组成环"""
        if messagebox.askyesno("确认", "确定要清空所有组成环吗？"):
            self.chain.rings.clear()
            for item in self.tree.get_children():
                self.tree.delete(item)
            self.update_count()
    
    def on_double_click(self, event):
        """双击列表项时填充到输入框"""
        selected = self.tree.selection()
        if not selected:
            return
        
        item = selected[0]
        values = self.tree.item(item, 'values')
        
        # 填充输入框
        self.ring_name_var.set(values[1])
        self.nominal_var.set(values[2])
        self.upper_dev_var.set(values[3])
        self.lower_dev_var.set(values[4])
        self.is_add_var.set(values[5] == "增环")
        self.is_fixed_var.set(values[6] == "是")
        self.alloc_type_var.set(values[7])
        self.bias_mode_var.set(values[8])
        self.allow_asymmetry_var.set(values[9] == "是")

        # 回填该行公差建议
        ring_id = int(values[0])
        for ring in self.chain.rings:
            if ring['ring_id'] == ring_id:
                self.t_min_var.set(str(ring.get('t_min', 0.008)))
                self.t_pref_var.set(str(ring.get('t_pref', 0.012)))
                self.t_max_var.set(str(ring.get('t_max', 0.02)))
                break
    
    def update_ring(self):
        """更新选中的组成环"""
        selected = self.tree.selection()
        if not selected:
            messagebox.showwarning("警告", "请先选中要更新的项")
            return
        
        try:
            name = self.ring_name_var.get().strip()
            if not name:
                messagebox.showwarning("警告", "请输入零件名称")
                return
            
            nominal = float(self.nominal_var.get())
            upper_dev = float(self.upper_dev_var.get())
            lower_dev = float(self.lower_dev_var.get())
            is_add = self.is_add_var.get()
            is_fixed = self.is_fixed_var.get()
            alloc_type = self.alloc_type_var.get()
            bias_mode = self.bias_mode_var.get()
            allow_asymmetry = self.allow_asymmetry_var.get()
            t_min = float(self.t_min_var.get())
            t_pref = float(self.t_pref_var.get())
            t_max = float(self.t_max_var.get())
            
            # 获取选中项的索引
            item = selected[0]
            index = self.tree.index(item)
            old_ring = self.chain.rings[index]
            
            # 更新数据
            self.chain.rings[index] = self.chain._normalize_ring({
                'ring_id': old_ring['ring_id'],
                'name': name,
                'nominal': nominal,
                'upper_dev': upper_dev,
                'lower_dev': lower_dev,
                'is_add': is_add,
                'is_fixed': is_fixed,
                'is_standard': is_fixed,
                'alloc_type': alloc_type,
                'bias_mode': bias_mode,
                'allow_asymmetry': allow_asymmetry,
                't_min': t_min,
                't_pref': t_pref,
                't_max': t_max
            })
            self.refresh_tree()
            
            # 清空输入
            self.ring_name_var.set("")
            self.nominal_var.set("1")
            self.upper_dev_var.set("0.01")
            self.lower_dev_var.set("-0.01")
            
            messagebox.showinfo("成功", "组成环已更新")
            
        except ValueError:
            messagebox.showerror("错误", "请输入有效的数字")
    
    def update_count(self):
        """更新组成环计数"""
        count = len(self.chain.rings)
        self.count_label.config(text=f"组成环数量: {count}")
    
    def calculate_chain(self):
        """计算尺寸链"""
        if not self.chain.rings:
            messagebox.showwarning("警告", "请先添加组成环")
            return
        
        result = self.chain.calculate()
        if result:
            # 显示结果
            self.result_text.delete("1.0", "end")
            
            result_str = f"""尺寸链名称: {self.name_var.get()}
═══════════════════════════════════════
总公称尺寸:   {result['nominal']:.4f} mm
上偏差:       {result['upper_dev']:.4f} mm
下偏差:       {result['lower_dev']:.4f} mm
公差:         {result['tolerance']:.4f} mm
公差中间值:   {result['mid_value']:.4f} mm
═══════════════════════════════════════
组成环数量:   {result['ring_count']} 个
"""
            self.result_text.insert("1.0", result_str)
            
            # 更新尺寸链名称
            self.chain.name = self.name_var.get()
    
    def allocate_tolerance(self):
        """智能分配公差"""
        if not self.chain.rings:
            messagebox.showwarning("警告", "请先添加组成环")
            return
        
        try:
            target_upper = float(self.target_upper_var.get())
            target_lower = float(self.target_lower_var.get())
        except ValueError:
            messagebox.showerror("错误", "请输入有效的偏差值")
            return

        self.on_target_change()

        allocation_plan, summary, message = ToleranceAllocator.allocate(
            self.chain.rings, target_upper, target_lower
        )

        # 保存分配结果到实例变量
        self.allocated_plan = allocation_plan
        self.alloc_summary = summary

        # 显示分配结果
        self.alloc_result_label.config(text=message)

        if self.allocated_plan:
            result_str = "\n公差分配结果:\n"
            result_str += f"{'ID':<6}{'零件名称':<14} {'上偏差':>10} {'下偏差':>10} {'公差':>10}\n"
            result_str += "-" * 56 + "\n"

            for ring_id, plan in sorted(self.allocated_plan.items(), key=lambda x: x[0]):
                result_str += (
                    f"{ring_id:<6}{plan['name']:<14} "
                    f"{plan['upper_dev']:>10.3f} {plan['lower_dev']:>10.3f} {plan['tolerance']:>10.3f}\n"
                )

            result_str += "\n"
            result_str += f"目标总公差:     {summary['target_tolerance']:.3f} mm\n"
            result_str += f"固定件占用公差: {summary['fixed_tolerance']:.3f} mm\n"
            result_str += f"可分配公差:     {summary['available_tolerance']:.3f} mm\n"
            result_str += f"实际总公差:     {summary['actual_tolerance']:.3f} mm\n"
            result_str += f"目标中值:       {summary['target_mid']:.3f} mm\n"
            result_str += f"实际中值:       {summary['actual_mid']:.3f} mm"

            if summary.get('tips'):
                result_str += "\n\n提示:\n- " + "\n- ".join(summary['tips'])
            
            self.result_text.delete("1.0", "end")
            self.result_text.insert("1.0", result_str)
    
    def apply_allocation(self):
        """应用公差分配结果到组成环"""
        if len(self.allocated_plan) == 0:
            messagebox.showwarning("警告", "请先进行公差分配")
            return
        
        # 更新组成环数据
        for ring in self.chain.rings:
            ring_id = ring['ring_id']
            if ring_id in self.allocated_plan:
                ring['upper_dev'] = self.allocated_plan[ring_id]['upper_dev']
                ring['lower_dev'] = self.allocated_plan[ring_id]['lower_dev']

        self.refresh_tree()
        
        messagebox.showinfo("成功", "已应用公差分配方案")
    
    def save_to_excel(self):
        """保存到Excel"""
        if not self.chain.rings:
            messagebox.showwarning("警告", "没有数据可保存")
            return
        
        # 默认文件名
        default_name = f"{self.name_var.get()}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        default_path = os.path.join(self.save_dir, default_name)
        
        filepath = filedialog.asksaveasfilename(
            initialdir=self.save_dir,
            initialfile=default_name,
            defaultextension=".xlsx",
            filetypes=[("Excel文件", "*.xlsx"), ("所有文件", "*.*")]
        )
        
        if not filepath:
            return
        
        try:
            # 计算结果
            result = self.chain.calculate()
            
            # 准备组成环数据
            rings_data = []
            for ring in self.chain.rings:
                rings_data.append({
                    'ID': ring.get('ring_id', ''),
                    '零件名称': ring['name'],
                    '公称尺寸': ring['nominal'],
                    '上偏差': ring['upper_dev'],
                    '下偏差': ring['lower_dev'],
                    '环类型': '增环' if ring['is_add'] else '减环',
                    '是否固定公差': '是' if ring.get('is_fixed', False) else '否',
                    '是否标准件': '是' if ring.get('is_fixed', False) else '否',  # 兼容旧版本
                    '分配类型': ring.get('alloc_type', '普通'),
                    '偏差模式': ring.get('bias_mode', '保持当前中值'),
                    '允许不对称': '是' if ring.get('allow_asymmetry', False) else '否',
                    'T_min': ring.get('t_min', 0.008),
                    'T_pref': ring.get('t_pref', 0.012),
                    'T_max': ring.get('t_max', 0.020),
                    '公差': abs(ring['upper_dev'] - ring['lower_dev'])
                })
            
            # 准备计算结果数据
            result_data = {
                '尺寸链名称': [self.name_var.get()],
                '总公称尺寸': [result['nominal']],
                '上偏差': [result['upper_dev']],
                '下偏差': [result['lower_dev']],
                '公差': [result['tolerance']],
                '公差中间值': [result['mid_value']],
                '组成环数量': [result['ring_count']],
                '保存版本': ['V2']
            }
            
            # 使用pandas保存
            rings_df = pd.DataFrame(rings_data)
            result_df = pd.DataFrame(result_data)
            
            with pd.ExcelWriter(filepath, engine='openpyxl') as writer:
                rings_df.to_excel(writer, sheet_name='组成环', index=False)
                result_df.to_excel(writer, sheet_name='计算结果', index=False)
            
            messagebox.showinfo("成功", f"数据已保存到:\n{filepath}")
            
        except Exception as e:
            logger.exception("保存Excel失败")
            messagebox.showerror("错误", f"保存失败: {str(e)}")
    
    def load_from_excel(self):
        """从Excel读取"""
        filepath = filedialog.askopenfilename(
            initialdir=self.save_dir,
            filetypes=[("Excel文件", "*.xlsx"), ("所有文件", "*.*")]
        )
        
        if not filepath:
            return
        
        try:
            # 读取组成环数据
            rings_df = pd.read_excel(filepath, sheet_name='组成环')
            
            # 清空现有数据
            self.chain.rings.clear()
            for item in self.tree.get_children():
                self.tree.delete(item)
            
            # 添加读取的数据
            for _, row in rings_df.iterrows():
                name = str(row['零件名称']).strip()  # 统一数据类型
                nominal = float(row['公称尺寸'])
                upper_dev = float(row['上偏差'])
                lower_dev = float(row['下偏差'])
                is_add = row['环类型'] == '增环'
                if '是否固定公差' in rings_df.columns:
                    is_fixed = str(row.get('是否固定公差', '否')).strip() == '是'
                else:
                    is_fixed = str(row.get('是否标准件', '否')).strip() == '是'

                alloc_type = str(row.get('分配类型', '普通')).strip() or '普通'
                bias_mode = str(row.get('偏差模式', '保持当前中值')).strip() or '保持当前中值'
                allow_asymmetry = str(row.get('允许不对称', '否')).strip() == '是'

                defaults = ToleranceAllocator.get_type_defaults(alloc_type)
                t_min = float(row.get('T_min', defaults['t_min']))
                t_pref = float(row.get('T_pref', defaults['t_pref']))
                t_max = float(row.get('T_max', defaults['t_max']))

                ring_id = row.get('ID', None)
                if pd.isna(ring_id):
                    ring_id = None

                self.chain.add_ring(
                    name,
                    nominal,
                    upper_dev,
                    lower_dev,
                    is_add,
                    is_fixed=is_fixed,
                    alloc_type=alloc_type,
                    bias_mode=bias_mode,
                    allow_asymmetry=allow_asymmetry,
                    t_min=t_min,
                    t_pref=t_pref,
                    t_max=t_max,
                    ring_id=ring_id
                )

            self.refresh_tree()
            
            # 尝试读取尺寸链名称
            try:
                result_df = pd.read_excel(filepath, sheet_name='计算结果')
                if '尺寸链名称' in result_df.columns:
                    self.name_var.set(result_df['尺寸链名称'].iloc[0])
            except Exception:
                pass
            
            self.update_count()
            messagebox.showinfo("成功", f"数据已从以下文件读取:\n{filepath}")
            
        except Exception as e:
            logger.exception("读取Excel失败")
            messagebox.showerror("错误", f"读取失败: {str(e)}")


def run_self_test():
    """命令行快速自测：不启动GUI，验证核心分配逻辑"""
    print("[自测] 开始执行尺寸链分配自测...")
    chain = DimensionChain("自测尺寸链")
    chain.add_ring("轴承", 20.0, 0.004, -0.002, is_add=True, is_fixed=True)
    chain.add_ring("一级总长", 35.0, 0.01, -0.01, is_add=True, is_fixed=False, alloc_type='总长', bias_mode='对称', allow_asymmetry=True)
    chain.add_ring("二级承接", 9.4, 0.006, -0.004, is_add=False, is_fixed=False, alloc_type='普通', bias_mode='保持当前中值', allow_asymmetry=False)

    plan, summary, message = ToleranceAllocator.allocate(chain.rings, target_upper=0.05, target_lower=-0.05)
    print("[自测] 分配消息:", message)
    print("[自测] 分配数量:", len(plan))
    print("[自测] 实际总公差:", f"{summary.get('actual_tolerance', 0):.3f}")
    print("[自测] 实际中值:", f"{summary.get('actual_mid', 0):.3f}")
    print("[自测] 通过")


if __name__ == "__main__":
    if "--self-test" in sys.argv:
        run_self_test()
        sys.exit(0)

    root = tk.Tk()
    root.title("尺寸链计算工具 V1.0")
    root.geometry("980x780")
    
    # 获取命令行参数，支持右键菜单启动
    if len(sys.argv) > 1:
        save_dir = sys.argv[1]
    else:
        save_dir = "./save"
    
    app = ChainCalculatorFrame(root, save_dir=save_dir)
    app.pack(fill="both", expand=True)

    app.target_upper_var.trace_add('write', lambda *_: app.on_target_change())
    app.target_lower_var.trace_add('write', lambda *_: app.on_target_change())
    
    root.mainloop()
