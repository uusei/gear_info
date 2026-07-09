import tkinter as tk
from tkinter import ttk, messagebox
import math

class NGWLogic:
    """核心计算逻辑"""
    @staticmethod
    def calculate_stage(m, zs, zp, zr, a, ds_s, ds_p, ds_r, pos_tol, gr, alpha_std_deg=20.0):
        alpha_std = math.radians(alpha_std_deg)
        # 1. 计算工作压力角 (acosd 逻辑)
        try:
            cos_aw_sp = (m * (zs + zp) / 2) * math.cos(alpha_std) / a
            aw_sp = math.acos(max(-1, min(1, cos_aw_sp)))
            
            cos_aw_pr = (m * (zr - zp) / 2) * math.cos(alpha_std) / a
            aw_pr = math.acos(max(-1, min(1, cos_aw_pr)))
        except:
            return None, "几何参数不合理，无法形成啮合（请检查中心距）"

        # 2. 侧隙计算 (法向)
        # 齿厚减薄 + 位置度贡献 + 轴承游隙贡献
        jn_thin = (ds_s + 2 * ds_p + ds_r) * math.cos(alpha_std)
        jn_pos = 2 * pos_tol * (math.sin(aw_sp) + math.sin(aw_pr))
        jn_bearing = gr * (math.sin(aw_sp) + math.sin(aw_pr))
        
        j_total = jn_thin + jn_pos + jn_bearing
        
        # 3. 转化为转角
        rb_s = (m * zs / 2) * math.cos(alpha_std)
        ratio = 1 + zr / zs
        phi_rad = j_total / (2 * rb_s * ratio)
        phi_arcmin = math.degrees(phi_rad) * 60
        
        return {"phi": phi_arcmin, "ratio": ratio}, None

class BacklashFrame(ttk.Frame):
    def __init__(self, parent):
        super().__init__(parent)

        # self.title("NGW 行星减速器回程间隙计算器 V1.0")
        # self.geometry("500x650")
        self.all_inputs = {"single": {}, "double": {}}
        # 创建选项卡
        self.tab_control = ttk.Notebook(self)
        self.tab1 = ttk.Frame(self.tab_control)
        self.tab2 = ttk.Frame(self.tab_control)
        
        self.tab_control.add(self.tab1, text='  单级计算  ')
        self.tab_control.add(self.tab2, text='  二级计算  ')
        self.tab_control.pack(expand=1, fill="both")
        
        self.setup_stage_ui(self.tab1, is_second_stage=False)
        self.setup_stage_ui(self.tab2, is_second_stage=True)

    def setup_stage_ui(self, frame, is_second_stage):
        # 参数输入区域
        input_frame = ttk.LabelFrame(frame, text=" 齿轮与公差参数 ")
        input_frame.pack(padx=10, pady=10, fill="x")

        labels = [
            ("模数 m", "0.4"), ("太阳轮齿数 zs", "13"), ("行星轮齿数 zp", "31"), 
            ("内齿圈齿数 zr", "77"), ("实际中心距 a", "9.2"),
            ("太阳轮减薄 ds_s (mm)", "0.01"), ("行星轮减薄 ds_p (mm)", "0.01"), 
            ("齿圈减薄 ds_r (mm)", "0.012"), ("行星架位置度 (mm)", "0.01"),
            ("轴承径向游隙 (mm)", "0.005")
        ]

        entries = {}
        # 如果是第二级，需要两组参数
        target_stages = ["Stage1"] if not is_second_stage else ["Stage1", "Stage2"]
        
        for stage_name in target_stages:
            stage_box = ttk.LabelFrame(input_frame, text=f" {stage_name} ")
            stage_box.pack(padx=5, pady=5, fill="x")
            
            row_entries = {}
            for i, (label_text, default_val) in enumerate(labels):
                row, col = divmod(i, 2)
                ttk.Label(stage_box, text=label_text).grid(row=row, column=col*2, padx=5, pady=2, sticky="e")
                ent = ttk.Entry(stage_box, width=10)
                ent.insert(0, default_val)
                ent.grid(row=row, column=col*2+1, padx=5, pady=2)
                row_entries[label_text.split()[0]] = ent
            entries[stage_name] = row_entries

        mode = "double" if is_second_stage else "single"
        self.all_inputs[mode] = entries
        # 计算按钮
        btn = ttk.Button(frame, text="立即计算", command=lambda: self.run_calc(entries, is_second_stage))
        btn.pack(pady=10)

        # 结果显示
        res_frame = ttk.LabelFrame(frame, text=" 计算结果 ")
        res_frame.pack(padx=10, pady=5, fill="both", expand=True)
        self.res_label = tk.Label(res_frame, text="等待输入...", font=("微软雅黑", 12, "bold"), fg="blue", justify="left")
        self.res_label.pack(padx=10, pady=20)
        
        if is_second_stage: self.res_label_2 = self.res_label
        else: self.res_label_1 = self.res_label

    def run_calc(self, entries_map, is_second_stage):
        try:
            results = []
            for name in ["Stage1", "Stage2"] if is_second_stage else ["Stage1"]:
                if name not in entries_map: continue
            
                e = entries_map[name]
                res, err = NGWLogic.calculate_stage(
                    float(e['模数'].get()), float(e['太阳轮齿数'].get()), float(e['行星轮齿数'].get()),
                    float(e['内齿圈齿数'].get()), float(e['实际中心距'].get()), float(e['太阳轮减薄'].get()),
                    float(e['行星轮减薄'].get()), float(e['齿圈减薄'].get()), float(e['行星架位置度'].get()),
                    float(e['轴承径向游隙'].get())
                )
                if err: 
                    messagebox.showerror("错误", f"{name}: {err}")
                    return
                results.append(res)

            if not is_second_stage:
                total_bl = results[0]['phi']
                ratio = results[0]['ratio']
                msg = f"一级总间隙: {total_bl:.2f} arcmin\n一级传动比: {ratio:.2f}"
                self.res_label_1.config(text=msg)
            else:
                # 二级叠加: S2 + S1/i2
                i1 = results[0]['ratio']
            i2 = results[1]['ratio']
            phi1 = results[0]['phi']
            phi2 = results[1]['phi']
            
            # 计算公式：负载端总间隙 = 最后一级间隙 + 前一级间隙/后一级传动比
            total_bl = phi2 + (phi1 / i2)
            
            msg = (f"二级整机总间隙: {total_bl:.2f} arcmin\n"
                   f"总传动比: {i1 * i2:.2f}\n"
                   f"--------------------------------\n"
                   f"高速级(Stage1)贡献: {phi1 / i2:.2f} arcmin\n"
                   f"低速级(Stage2)贡献: {phi2:.2f} arcmin")
            self.res_label_2.config(text=msg)

        except ValueError:
            messagebox.showerror("错误", "请输入有效的数字格式")
    
    def sync_gear_data(self, config_type, gear_data):
        def fill_stage(target_dict, za, zr, zp, mn, a):
            if not target_dict: return
            self._fill_val(target_dict['太阳轮齿数'], za)
            self._fill_val(target_dict['内齿圈齿数'], abs(float(zr)))
            self._fill_val(target_dict['行星轮齿数'], zp)
            self._fill_val(target_dict['模数'], mn)       # 对应 Label 中的 "模数 m"
            self._fill_val(target_dict['实际中心距'], a)  # 对应 Label 中的 "实际中心距 a"
    # 映射逻辑：主计算器的 Za 对应 Backlash 的 zs，Zb 对应 zr
        try:
            if config_type == "NGW1":
                # 更新单级/第一级标签页中的输入框
                self.tab_control.select(0)
                fill_stage(self.all_inputs["single"].get("Stage1"), 
                       gear_data['za1'], gear_data['zb1'], gear_data['zc1'], gear_data['mn1'], gear_data['a1'])
            
            elif config_type == "NGW2":
                # 更新二级标签页
                # 第一级
                self.tab_control.select(1)
                # 同步第一级
                fill_stage(self.all_inputs["double"].get("Stage1"), 
                           gear_data['za1'], gear_data['zb1'], gear_data['zc1'], gear_data['mn1'], gear_data['a1'])
                # 同步第二级
                fill_stage(self.all_inputs["double"].get("Stage2"), 
                           gear_data['za2'], gear_data['zb2'], gear_data['zc2'], gear_data['mn2'], gear_data['a2'])
        
        except Exception as e:
            print(f"同步数据失败: {e}")
    def _fill_val(self, entry_widget, value):
        """辅助方法：清空并填充 Entry"""
        if value is not None:
            entry_widget.delete(0, tk.END)
            entry_widget.insert(0, str(value))
# if __name__ == "__main__":
    # root = tk.Tk()
    # app = BacklashApp(root)
    # root.mainloop()