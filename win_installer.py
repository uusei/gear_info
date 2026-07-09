import winreg
import sys
import os

# 1. 配置你的程序信息
# 请确保这里的路径指向你真正的 python.exe 和 脚本文件
python_exe = sys.executable  # 获取当前 Python 解释器路径
script_path = os.path.abspath("update_file.exe")  # 你的主程序文件名
# script_path = os.path.abspath("chain.exe")

# 2. 定义右键菜单显示的名称
menu_name = "使用图纸管理系统打开"
# menu_name = "使用尺寸链计算器打开"

def add_to_context_menu(target_key, command_val):
    try:
        # 创建主键
        key_path = rf"{target_key}\shell\DrawingManager"
        # key_path = rf"{target_key}\shell\ChainCalculator"
        with winreg.CreateKey(winreg.HKEY_CLASSES_ROOT, key_path) as key:
            winreg.SetValue(key, "", winreg.REG_SZ, menu_name)
            # winreg.SetValueEx(key, "Icon", 0, winreg.REG_SZ, python_exe) # 可选：添加图标
            
        # 创建执行命令子键
        with winreg.CreateKey(winreg.HKEY_CLASSES_ROOT, rf"{key_path}\command") as key:
            winreg.SetValue(key, "", winreg.REG_SZ, command_val)
            
        print(f"成功添加到: {target_key}")
    except Exception as e:
        print(f"添加失败: {e}")

if __name__ == "__main__":
    # 执行命令：python.exe "脚本路径" "目标文件夹路径"
    # %1 代表右键点击的目标路径
    cmd = f' "{script_path}" "%1"'
    
    # 分别添加到：文件夹右键、目录背景右键
    add_to_context_menu(r"Directory", cmd)           # 点击文件夹图标时
    add_to_context_menu(r"Directory\Background", cmd) # 在文件夹空白处右键时 (此时传的是 %V)
    
    # 背景右键需要特殊处理 %V
    # cmd_bg = f'"{python_exe}" "{script_path}" "%V"'
    cmd_bg = f'"{script_path}" "%V" source:ExplorerBackground'
    add_to_context_menu(r"Directory\Background", cmd_bg)

    print("\n--- 设置完成！现在你可以在文件夹上点击右键试试了 ---")
    os.system("pause")