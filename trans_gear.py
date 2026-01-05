import os
import pdfplumber
import pandas as pd
import re
from openpyxl import Workbook

def extract_gear_parameters_from_pdf(pdf_path):
    """
    从PDF文件中提取齿轮参数
    """
    # 读取PDF文本
    text = ""
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text += page.extract_text() + "\n"
    
    # 初始化三个齿轮的数据字典
    sun_gear = {}
    planet_gear = {}
    ring_gear = {}
    
    # 提取基本参数
    patterns = {
        '齿数': r'齿数.*?\[z\].*?(\d+)\s+(\d+)\s+([-\d]+)',
        '法向模数': r'法向模数.*?\[mn\].*?([\d.]+)',
        '压力角': r'法向压力角.*?\[αn\].*?([\d.]+)',
        '螺旋角': r'分度圆上的螺旋角.*?\[β\].*?([\d]+)',
        '螺旋方向': r'螺旋线方向.*?[\u4e00-\u9fa5]+啮合',
        '齿顶高系数': r'基准齿廓齿顶高.*?\[haP\*\].*?([\d.]+)',
        '齿根高系数': r'基准齿廓齿根高.*?\[hfP\*\].*?([\d.]+)\s+([\d.]+)\s+([\d.]+)',
        '齿廓变位系数': r'齿廓变位系数.*?\[x\].*?([\d.-]+)\s+([\d.-]+)\s+([-\d.-]+)',
        '齿根圆直径': r'齿根圆直径.*?\[df\].*?([\d.]+)\s+([\d.]+)\s+([-\d.]+)',
        '齿顶圆直径': r'齿顶圆直径.*?\[da\].*?([\d.]+)\s+([\d.]+)\s+([-\d.]+)',
        '渐开线起始圆': r'齿根成形圆直径.*?\[dFf\].*?([\d.]+)\s+([\d.]+)\s+([-\d.]+)',
        '齿根圆角系数': r'基准齿廓齿根半径.*?\[ρfP\*\].*?([\d.]+)\s+([\d.]+)\s+([\d.]+)',
        '中心距': r'中心距.*?\[a\].*?([\d.]+)',
        '跨齿数': r'跨齿数.*?\[k\].*?([\d.]+)\s+([\d.]+)\s+([-\d.]+)',
        '量棒直径': r'有效量规直径.*?\[DMeff\].*?([\d.]+)\s+([\d.]+)\s+([\d.]+)',
        '单个齿距偏差': r'单个齿距偏差的公差.*?\[fpt\].*?([\d.]+)\s+([\d.]+)\s+([\d.]+)',
        '齿距累计总偏差': r'齿距累积总偏差的公差.*?\[FPT\].*?([\d.]+)\s+([\d.]+)\s+([\d.]+)',
        '齿廓总偏差': r'齿廓总偏差的公差.*?\[FαT\].*?([\d.]+)\s+([\d.]+)\s+([\d.]+)',
        '螺旋线总偏差': r'螺旋线总偏差的公差.*?\[FβT\].*?([\d.]+)\s+([\d.]+)\s+([\d.]+)',
        '径向跳动偏差': r'径跳偏差的公差.*?\[FrT\].*?([\d.]+)\s+([\d.]+)\s+([\d.]+)'
    }
    
    # 提取数据
    for param_name, pattern in patterns.items():
        match = re.search(pattern, text, re.IGNORECASE | re.DOTALL)
        if match:
            if param_name in ['齿数', '齿廓变位系数', '齿根圆直径', '齿顶圆直径', '渐开线起始圆', '齿根高系数','齿根圆角系数',
                            '跨齿数', '量棒直径', '单个齿距偏差', '齿距累计总偏差',
                            '齿廓总偏差', '螺旋线总偏差', '径向跳动偏差']:
                groups = match.groups()
                if len(groups) >= 3:
                    sun_gear[param_name] = groups[0]
                    planet_gear[param_name] = groups[1]
                    ring_gear[param_name] = groups[2]
            elif param_name in ['法向模数', '压力角', '螺旋角', '齿顶高系数',  
                               '中心距']:
                sun_gear[param_name] = match.group(1)
                planet_gear[param_name] = match.group(1)
                ring_gear[param_name] = match.group(1)
            elif param_name == '螺旋方向':
                sun_gear[param_name] = '直齿'
                planet_gear[param_name] = '直齿'
                ring_gear[param_name] = '直齿'
    
    # 处理行星轮数量
    planet_count_match = re.search(r'齿轮数量.*?\[p\].*?1\s+(\d+)\s+1', text)
    if planet_count_match:
        planet_gear['数量'] = planet_count_match.group(1)
    
    # 处理公法线长度（太阳轮和行星轮）
    w_pattern = r'\[Wk\.e/i\].*?([\d.]+)\s+/\s+([\d.]+)\s+'
    sun_w_match = re.search(w_pattern, text)
    if sun_w_match:
        sun_gear['公法线长度_Wmax'] = sun_w_match.group(1)
        sun_gear['公法线长度_Wmin'] = sun_w_match.group(2)
    else:
        w_pattern = r'\[Wk\.e/i\].*?([\d.]+)\s+/([\d.]+)\s+'
        sun_w_match = re.search(w_pattern, text)
        sun_gear['公法线长度_Wmax'] = sun_w_match.group(1)
        sun_gear['公法线长度_Wmin'] = sun_w_match.group(2)
        
    if sun_w_match:
        remaining_text = text[sun_w_match.end():]
        planet_w_pattern = r'([\d.]+)\s+/\s+([\d.]+)'
        planet_w_match = re.search(planet_w_pattern, remaining_text)
        if planet_w_match:
            planet_gear['公法线长度_Wmax'] = planet_w_match.group(1)
            planet_gear['公法线长度_Wmin'] = planet_w_match.group(2)

    # 处理齿顶圆公差带
    da_pattern = r'\[da\.e/i\].*?([\d.]+)\s+/\s+([\d.]+)'
    sun_da_match = re.search(da_pattern, text)
    if sun_da_match:
        try:
            sun_da_max = float(sun_da_match.group(1))
            sun_da_min = float(sun_da_match.group(2))
            sun_gear['齿顶公差范围'] = sun_da_max - sun_da_min
        except:
            sun_gear['齿顶公差范围'] = 0
    else:
        da_pattern = r'\[da\.e/i\].*?([\d.]+)\s+/([\d.]+)'
        sun_da_match = re.search(da_pattern, text)
        try:
            sun_da_max = float(sun_da_match.group(1))
            sun_da_min = float(sun_da_match.group(2))
            sun_gear['齿顶公差范围'] = sun_da_max - sun_da_min
        except:
            sun_gear['齿顶公差范围'] = 0
    
    if sun_da_match:
        remaining_text = text[sun_da_match.end():]
        planet_pattern = r'([\d.]+)\s+/([\d.]+)'
        planet_da_match = re.search(planet_pattern, remaining_text)
        if planet_da_match:
            try:
                planet_da_max = float(planet_da_match.group(1))
                planet_da_min = float(planet_da_match.group(2))
                planet_gear['齿顶公差范围'] = planet_da_max - planet_da_min
            except:
                planet_gear['齿顶公差范围'] = 0
        if planet_da_match:
            remaining_text = text[planet_da_match.end():]
            ring_pattern = r'\[Ada\.e/i\].*?([\d.]+)'
            ring_da_match = re.search(ring_pattern, remaining_text)
            remaining_text1 = remaining_text[ring_da_match.end():]
            ring_da_match1 = re.search(ring_pattern, remaining_text1)
            remaining_text2 = remaining_text1[ring_da_match1.end():]
            ring_da_match2 = re.search(ring_pattern, remaining_text2)
            remaining_text2 = remaining_text2[ring_da_match1.end():]
            
            if ring_da_match2:
                try:
                    ring_da_max = float(ring_da_match2.group(1))
                    ring_gear['齿顶公差范围'] = ring_da_max
                except:
                    ring_gear['齿顶公差范围'] = 0

    # 处理跨棒距（齿圈）- 径向二针跨球距
    md_pattern = r'径向二针跨球距.*?\[MdK\.e/i\].*?([\d.]+)\s+/\s+([\d.]+)'
    md_match = re.search(md_pattern, text)
    if not md_match:
        md_pattern = r'径向二针跨球距.*?\[MdK\.e/i\].*?([\d.]+)\s+/([\d.]+)'
        md_match = re.search(md_pattern, text)
    
    remaining_text = text[md_match.end():]
    md_pattern1 = r'([\d.]+)\s+/([\d.]+)'
    md_match1 = re.search(md_pattern1, remaining_text)
    remaining_text1 = remaining_text[md_match1.end():]
    md_pattern2 = r'([\d.]+)\s+/([\d.]+)'
    md_match2 = re.search(md_pattern2, remaining_text1)

    if md_match2:
        ring_gear['跨棒距_max'] = md_match2.group(1)
        ring_gear['跨棒距_min'] = md_match2.group(2)
    
    

    # 如果没有找到齿廓变位系数，尝试从产形齿廓变位系数中提取（作为备选）
    if '齿廓变位系数' not in sun_gear:
        backup_pattern = r'产形齿廓变位系数.*?\[xE e/i\].*?([\d.-]+).*?([\d.-]+).*?([\d.-]+)'
        backup_match = re.search(backup_pattern, text)
        if backup_match:
            sun_gear['齿廓变位系数'] = backup_match.group(1)
            planet_gear['齿廓变位系数'] = backup_match.group(2)
            ring_gear['齿廓变位系数'] = backup_match.group(3)
    
    # 计算顶隙系数：齿根高系数 - 齿顶高系数
    if '齿根高系数' in sun_gear and '齿顶高系数' in sun_gear:
        try:
            hf = float(sun_gear['齿根高系数'])
            ha = float(sun_gear['齿顶高系数'])
            c_value = hf - ha
            sun_gear['顶隙系数'] = f"{c_value:.2f}"
            hf = float(planet_gear['齿根高系数'])
            ha = float(planet_gear['齿顶高系数'])
            c_value = hf - ha
            planet_gear['顶隙系数'] = f"{c_value:.2f}"
            hf = float(ring_gear['齿根高系数'])
            ha = float(ring_gear['齿顶高系数'])
            c_value = hf - ha
            ring_gear['顶隙系数'] = f"{c_value:.2f}"
        except ValueError:
            sun_gear['顶隙系数'] = "0.25"
            planet_gear['顶隙系数'] = "0.25"
            ring_gear['顶隙系数'] = "0.25"
    else:
        # 如果无法计算，使用默认值
        sun_gear['顶隙系数'] = "0.25"
        planet_gear['顶隙系数'] = "0.25"
        ring_gear['顶隙系数'] = "0.25"
    
    # 设置精度等级
    sun_gear['精度等级'] = "ISO1328"
    planet_gear['精度等级'] = "ISO1328"
    ring_gear['精度等级'] = "ISO1328"
    
    # 格式化数值
    def format_value(value, param_name):
        if not value:
            return ""
        if param_name in ['中心距', '齿根圆直径', '渐开线起始圆']:
            try:
                return f"{float(value):.4f}"
            except:
                return value
        if param_name in ['压力角', '螺旋角']:
            try:
                return f"{float(value):.1f}°"
            except:
                return value
        if param_name in ['齿数','跨齿数']:
            try:
                return f"{int(value.split('.')[0])}"
            except:
                return value        
        if param_name in ['单个齿距偏差']:
            try:
                return "±"+f"{str(value)}"
            except:
                return value
        return value
    
    # 构建标准化的数据结构
    def create_gear_data(gear_dict, gear_type):
        # 齿轮参数部分
        gear_params = {
            '参数名称': [],
            '符号': [],
            '数值': []
        }
        
        # 基本参数映射
        param_mapping = {
            '齿数': ('Z', '齿数', 'gear'),
            '法向模数': ('mn', '法向模数', 'gear'),
            '压力角': ('ɑ', '压力角', 'gear'),
            '螺旋角': ('β', '螺旋角', 'gear'),
            '螺旋方向': ('', '螺旋方向', 'gear'),
            '齿顶高系数': ('ha*', '齿顶高系数', 'gear'),
            '顶隙系数': ('C*', '顶隙系数', 'gear'),
            '齿廓变位系数': ('x', '径向变位系数', 'gear'),
            '齿根圆直径': ('df', '齿根圆直径', 'gear'),
            '齿顶圆直径': ('da', '齿顶圆直径', 'gear'),
            '渐开线起始圆': ('dFf', '渐开线起始圆', 'gear'),
            '齿根圆角系数': ('rhofP*', '齿根圆角系数', 'gear'),
            '中心距': ('a', '中心距', 'gear'),
            '相配齿轮图号': ('', '相配齿轮图号', 'gear'),
            '相配齿轮齿数': ('', '相配齿轮齿数', 'gear'),
            '精度等级': ('6', '精度等级', 'gear'),
            '跨齿数': ('k', '跨齿数', 'gear'),
            '公法线长度_Wmax': ('Wmax', '公法线长度', 'gear'),
            '公法线长度_Wmin': ('Wmin', '', 'gear'),
            '量棒直径': ('DM', '量棒直径', 'gear'),
            '跨棒距_max': ('Mmax', '跨棒距', 'gear'),
            '跨棒距_min': ('Mmin', '', 'gear'),
            '数量': ('', '数量', 'gear')
            
        }
        
        # 精度参数映射
        accuracy_mapping = {
            '单个齿距偏差': ('±fpt', '单个齿距偏差', 'accuracy'),
            '齿距累计总偏差': ('Fp', '齿距累计总偏差', 'accuracy'),
            '齿廓总偏差': ('Fɑ', '齿廓总偏差', 'accuracy'),
            '螺旋线总偏差': ('Fβ', '螺旋线总偏差', 'accuracy'),
            '径向跳动偏差': ('Fr', '径向跳动偏差', 'accuracy'),
            '齿顶公差范围': ('', '齿顶公差范围', 'accuracy')
        }
        
        # 处理齿轮参数
        for param_key, (symbol, display_name, param_type) in param_mapping.items():
            if param_type == 'gear':
                # 太阳轮和行星轮没有量棒直径
                if gear_type in ['太阳轮', '行星轮'] and param_key == '量棒直径':
                    continue
                    
                if param_key in gear_dict:
                    formatted_value = format_value(gear_dict[param_key], display_name)
                    gear_params['参数名称'].append(display_name)
                    gear_params['符号'].append(symbol)
                    gear_params['数值'].append(formatted_value)
                elif display_name in ['相配齿轮图号', '相配齿轮齿数']:
                    # 保留空行
                    gear_params['参数名称'].append(display_name)
                    gear_params['符号'].append(symbol)
                    gear_params['数值'].append('')
        
        # 添加齿轮精度标题
        gear_params['参数名称'].append('齿轮精度')
        gear_params['符号'].append('')
        gear_params['数值'].append('')
        
        # 处理齿轮精度参数
        for param_key, (symbol, display_name, param_type) in accuracy_mapping.items():
            if param_type == 'accuracy' and param_key in gear_dict:
                formatted_value = format_value(gear_dict[param_key], display_name)
                gear_params['参数名称'].append(display_name)
                gear_params['符号'].append(symbol)
                gear_params['数值'].append(formatted_value)
        
        return gear_params
    
    # 特殊处理：齿圈没有跨齿数，太阳轮和行星轮没有跨棒距和量棒直径
    if '跨齿数' in ring_gear:
        del ring_gear['跨齿数']
    if '跨棒距_max' in sun_gear:
        del sun_gear['跨棒距_max']
        del sun_gear['跨棒距_min']
    if '跨棒距_max' in planet_gear:
        del planet_gear['跨棒距_max']
        del planet_gear['跨棒距_min']
    if '公法线长度_Wmax' in ring_gear:
        del ring_gear['公法线长度_Wmax']
        del ring_gear['公法线长度_Wmin']
    if '量棒直径' in sun_gear:
        del sun_gear['量棒直径']
    if '量棒直径' in planet_gear:
        del planet_gear['量棒直径']
    
    sun_data = create_gear_data(sun_gear, '太阳轮')
    planet_data = create_gear_data(planet_gear, '行星轮')
    ring_data = create_gear_data(ring_gear, '齿圈')
    
    # 修改列名为指定的名称
    def rename_columns(data, sheet_name):
        df = pd.DataFrame(data)
        if sheet_name == '齿圈':
            df = df.rename(columns={'参数名称': '齿圈参数'})
        else:
            df = df.rename(columns={'参数名称': '齿轮参数'})
        return df
    
    sun_df = rename_columns(sun_data, '太阳轮')
    planet_df = rename_columns(planet_data, '行星轮')
    ring_df = rename_columns(ring_data, '齿圈')
    
    return sun_df, planet_df, ring_df

def process_all_pdfs(input_dir="./input", output_dir="./excel"):
    """
    处理input文件夹中的所有PDF文件
    """
    if not os.path.exists(input_dir):
        print(f"输入文件夹 {input_dir} 不存在")
        return
    
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
    
    # 获取所有PDF文件
    pdf_files = [f for f in os.listdir(input_dir) if f.lower().endswith('.pdf') and f.lower().startswith('gear')]
    
    if not pdf_files:
        print(f"在 {input_dir} 文件夹中未找到PDF文件")
        return
    
    for pdf_file in pdf_files:
        pdf_path = os.path.join(input_dir, pdf_file)
        excel_name = os.path.splitext(pdf_file)[0] + '.xlsx'
        output_path = os.path.join(output_dir, excel_name)
        
        try:
            print(f"正在处理: {pdf_file}")
            sun_df, planet_df, ring_df = extract_gear_parameters_from_pdf(pdf_path)
            
            # 保存为Excel文件
            with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
                sun_df.to_excel(writer, sheet_name='太阳轮', index=False)
                planet_df.to_excel(writer, sheet_name='行星轮', index=False)
                ring_df.to_excel(writer, sheet_name='齿圈', index=False)
            
            print(f"成功输出: {excel_name}")
            
        except Exception as e:
            print(f"处理文件 {pdf_file} 时出错: {str(e)}")

# 执行处理
if __name__ == "__main__":
    process_all_pdfs()