# code: utf-8
import os
import subprocess
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import sys
import io
import importlib.util
import traceback

# 导入genshin_3dmigoto_collect.py中的功能
sys.path.append(os.path.join(os.path.dirname(os.path.abspath(__file__)), "3dmigoto-GIMI-for-development"))

# 导入genshin_3dmigoto_collect模块
spec = importlib.util.spec_from_file_location(
    "genshin_3dmigoto_collect", 
    os.path.join(os.path.dirname(os.path.abspath(__file__)), "3dmigoto-GIMI-for-development", "genshin_3dmigoto_collect.py")
)
genshin_3dmigoto_collect = importlib.util.module_from_spec(spec)
sys.modules["genshin_3dmigoto_collect"] = genshin_3dmigoto_collect
spec.loader.exec_module(genshin_3dmigoto_collect)

class VBExtractorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("3DMigoto VB数据提取工具")
        self.root.geometry("560x560")  # 增加窗口高度以确保日志界面完全显示
        self.root.resizable(True, True)
        
        # 设置字体以支持中文
        self.font = ('Microsoft YaHei', 10)
        
        # 获取当前脚本所在目录
        self.current_dir = os.path.dirname(os.path.abspath(__file__))
        
        # 创建主框架
        self.main_frame = ttk.Frame(root, padding="20")
        self.main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 创建界面元素
        self.create_widgets()
        
        # 在创建完界面元素后再读取d3dx.ini文件获取target值
        self.target_exe = self.read_target_from_d3dx_ini()
        # 更新注入进程路径输入框的值
        self.process_path_var.set(self.target_exe)
        
        # 检测当前键位配置模式并设置复选框状态
        is_no_numpad = self.detect_keyboard_layout()
        self.numpad_checkbox_var.set(is_no_numpad)
        
    def read_target_from_d3dx_ini(self):
        """从d3dx.ini文件中读取target值"""
        target_exe = ""
        try:
            # 构建d3dx.ini文件的完整路径
            d3dx_ini_path = os.path.join(self.current_dir, "3dmigoto-GIMI-for-development", "d3dx.ini")
            
            # 检查文件是否存在
            if os.path.exists(d3dx_ini_path):
                # 读取文件内容
                with open(d3dx_ini_path, 'r', encoding='utf-8') as f:
                    lines = f.readlines()
                    
                # 查找[Loader]部分下的target值
                in_loader_section = False
                for line in lines:
                    line = line.strip()
                    if line.lower() == '[loader]':
                        in_loader_section = True
                    elif in_loader_section and line.lower().startswith('target ='):
                        target_exe = line.split('=', 1)[1].strip()
                        break
            
            self.log_message(f"从d3dx.ini读取到的目标进程: {target_exe}")
        except Exception as e:
            self.log_message(f"读取d3dx.ini文件时发生错误: {str(e)}")
        
        return target_exe
    
    def save_target_to_d3dx_ini(self, target_exe):
        """保存target值到d3dx.ini文件"""
        try:
            # 构建d3dx.ini文件的完整路径
            d3dx_ini_path = os.path.join(self.current_dir, "3dmigoto-GIMI-for-development", "d3dx.ini")
            
            # 检查文件是否存在
            if os.path.exists(d3dx_ini_path):
                # 读取文件内容
                with open(d3dx_ini_path, 'r', encoding='utf-8') as f:
                    lines = f.readlines()
                     
                # 查找[Loader]部分，并更新target值
                new_lines = []
                in_loader_section = False
                target_updated = False
                
                for line in lines:
                    stripped_line = line.strip()
                    if stripped_line.lower() == '[loader]':
                        in_loader_section = True
                        new_lines.append(line)
                    elif in_loader_section and stripped_line.lower().startswith('target ='):
                        # 更新target值
                        new_lines.append(f"target = {target_exe}\n")
                        target_updated = True
                        in_loader_section = False  # 更新后退出loader部分
                    elif in_loader_section and stripped_line.startswith('['):
                        # 如果到达下一个section但还没找到target项，添加target项
                        if not target_updated:
                            new_lines.append(f"target = {target_exe}\n")
                            target_updated = True
                        new_lines.append(line)
                        in_loader_section = False
                    else:
                        new_lines.append(line)
                
                # 如果文件中没有[Loader]部分，添加它和target值
                if not any(line.strip().lower() == '[loader]' for line in lines):
                    new_lines.append("[Loader]\n")
                    new_lines.append(f"target = {target_exe}\n")
                
                # 写回文件
                with open(d3dx_ini_path, 'w', encoding='utf-8') as f:
                    f.writelines(new_lines)
                
                self.log_message(f"已将目标进程保存到d3dx.ini: {target_exe}")
                return True
            else:
                self.log_message(f"d3dx.ini文件不存在: {d3dx_ini_path}")
                return False
        except Exception as e:
            self.log_message(f"保存d3dx.ini文件时发生错误: {str(e)}")
            return False
    
    def browse_process_path(self):
        # 打开文件选择对话框，限制选择可执行文件
        file_path = filedialog.askopenfilename(
            title="选择注入进程可执行文件",
            filetypes=[("可执行文件", "*.exe"), ("所有文件", "*.*")]
        )
        
        if file_path:
            self.process_path_var.set(file_path)
            self.log_message(f"已选择注入进程路径: {file_path}")
            
            # 将选择的exe文件名保存到d3dx.ini文件
            if self.save_target_to_d3dx_ini(file_path):
                self.log_message("目标进程已成功保存到d3dx.ini")
            else:
                self.log_message("目标进程保存到d3dx.ini失败")
    
    def toggle_key_layout(self):
        # 切换键位配置
        try:
            # 获取3dmigoto目录
            three_dmigoto_dir = os.path.join(
                self.current_dir, 
                "3dmigoto-GIMI-for-development"
            )
            
            # 构建d3dx.ini文件的完整路径
            d3dx_ini_path = os.path.join(three_dmigoto_dir, "d3dx.ini")
            
            # 检查文件是否存在
            if os.path.exists(d3dx_ini_path):
                # 读取文件内容
                with open(d3dx_ini_path, 'r', encoding='utf-8') as f:
                    lines = f.readlines()
                    
                # 定义小键盘和无小键盘的键位配置
                numpad_layout = {
                    "next_marking_mode": "no_modifiers VK_DECIMAL VK_NUMPAD0",
                    "previous_pixelshader": "no_modifiers NO_VK_DECIMAL VK_NUMPAD1",
                    "next_pixelshader": "no_modifiers NO_VK_DECIMAL VK_NUMPAD2",
                    "mark_pixelshader": "no_modifiers NO_VK_DECIMAL VK_NUMPAD3",
                    "previous_vertexshader": "no_modifiers NO_VK_DECIMAL VK_NUMPAD4",
                    "next_vertexshader": "no_modifiers NO_VK_DECIMAL VK_NUMPAD5",
                    "mark_vertexshader": "no_modifiers NO_VK_DECIMAL VK_NUMPAD6",
                    "previous_indexbuffer": "no_modifiers NO_VK_DECIMAL VK_NUMPAD7",
                    "next_indexbuffer": "no_modifiers NO_VK_DECIMAL VK_NUMPAD8",
                    "mark_indexbuffer": "no_modifiers NO_VK_DECIMAL VK_NUMPAD9",
                    "previous_vertexbuffer": "no_modifiers NO_VK_DECIMAL VK_DIVIDE",
                    "next_vertexbuffer": "no_modifiers NO_VK_DECIMAL VK_MULTIPLY",
                    "mark_vertexbuffer": "no_modifiers NO_VK_DECIMAL VK_SUBTRACT",
                    "done_hunting": "NO_MODIFIERS NO_VK_DECIMAL VK_ADD",
                    "toggle_hunting": "no_modifiers NO_VK_DECIMAL VK_NUMPAD0"
                }
                
                no_numpad_layout = {
                    "next_marking_mode": "no_modifiers NO_VK_DECIMAL VK_F11",
                    "previous_pixelshader": "no_modifiers NO_VK_DECIMAL VK_OEM_COMMA",
                    "next_pixelshader": "no_modifiers NO_VK_DECIMAL VK_OEM_PERIOD",
                    "mark_pixelshader": "no_modifiers NO_VK_DECIMAL VK_OEM_2",
                    "previous_vertexshader": "no_modifiers NO_VK_DECIMAL P",
                    "next_vertexshader": "no_modifiers NO_VK_DECIMAL [",
                    "mark_vertexshader": "no_modifiers NO_VK_DECIMAL ]",
                    "previous_indexbuffer": "no_modifiers NO_VK_DECIMAL 7",
                    "next_indexbuffer": "no_modifiers NO_VK_DECIMAL 8",
                    "mark_indexbuffer": "no_modifiers NO_VK_DECIMAL 9",
                    "previous_vertexbuffer": "no_modifiers NO_VK_DECIMAL 0",
                    "next_vertexbuffer": "no_modifiers NO_VK_DECIMAL VK_OEM_MINUS",
                    "mark_vertexbuffer": "no_modifiers NO_VK_DECIMAL VK_OEM_PLUS",
                    "done_hunting": "NO_MODIFIERS NO_VK_DECIMAL VK_F5",
                    "toggle_hunting": "no_modifiers NO_VK_DECIMAL VK_F12"
                }
                
                # 根据复选框状态选择键位配置
                target_layout = no_numpad_layout if self.numpad_checkbox_var.get() else numpad_layout
                mode_name = "无小键盘" if self.numpad_checkbox_var.get() else "小键盘"
                
                # 查找[Hunting]部分，并更新键位配置
                new_lines = []
                in_hunting_section = False
                hunting_lines_processed = set()
                
                for line in lines:
                    stripped_line = line.strip()
                    if stripped_line.lower() == '[hunting]':
                        in_hunting_section = True
                        new_lines.append(line)
                    elif in_hunting_section and stripped_line.startswith('['):
                        # 如果到达下一个section，停止处理[Hunting]部分
                        in_hunting_section = False
                        new_lines.append(line)
                    elif in_hunting_section:
                        # 处理[Hunting]部分内的行
                        if '=' in stripped_line:
                            key, _ = stripped_line.split('=', 1)
                            key = key.strip()
                            if key in target_layout:
                                # 如果这是需要更新的键位配置行，更新它
                                new_lines.append(f"{key} = {target_layout[key]}\n")
                                hunting_lines_processed.add(key)
                                continue
                        # 其他行保持不变
                        new_lines.append(line)
                    else:
                        # 不在[Hunting]部分，直接添加
                        new_lines.append(line)
                
                # 检查是否找到了[Hunting]部分
                if not any(line.strip().lower() == '[hunting]' for line in lines):
                    # 如果没有找到[Hunting]部分，在文件末尾添加
                    new_lines.append("\n[Hunting]\n")
                    for key, value in target_layout.items():
                        new_lines.append(f"{key} = {value}\n")
                elif hunting_lines_processed != set(target_layout.keys()):
                    # 如果[Hunting]部分存在但缺少某些键位配置，添加它们
                    # 找到[Hunting]部分的结束位置
                    hunting_end_index = -1
                    for i, line in enumerate(new_lines):
                        if line.strip().lower() == '[hunting]':
                            hunting_end_index = i
                            break
                    
                    if hunting_end_index != -1:
                        # 在[Hunting]部分开始后添加缺少的键位配置
                        for key, value in target_layout.items():
                            if key not in hunting_lines_processed:
                                new_lines.insert(hunting_end_index + 1, f"{key} = {value}\n")
                                hunting_end_index += 1
                
                # 写回文件
                with open(d3dx_ini_path, 'w', encoding='utf-8') as f:
                    f.writelines(new_lines)
                
                self.log_message(f"键位配置已更新为{mode_name}模式")
                messagebox.showinfo("成功", f"键位配置已更新为{mode_name}模式，请使用F10重载3dmigoto以应用更改")
                return True
            else:
                self.log_message(f"d3dx.ini文件不存在: {d3dx_ini_path}")
                messagebox.showerror("错误", f"d3dx.ini文件不存在: {d3dx_ini_path}")
                return False
        except Exception as e:
            self.log_message(f"更新键位配置时发生错误: {str(e)}")
            messagebox.showerror("错误", f"更新键位配置时发生错误: {str(e)}")
            return False

    def show_key_mappings(self):
        # 显示从配置文件读取的实际键位映射关系（表格形式）
        try:
            # 构建d3dx.ini文件的完整路径
            d3dx_ini_path = os.path.join(self.current_dir, "3dmigoto-GIMI-for-development", "d3dx.ini")
            
            # 检查文件是否存在
            if not os.path.exists(d3dx_ini_path):
                self.log_message(f"d3dx.ini文件不存在: {d3dx_ini_path}")
                messagebox.showerror("错误", f"d3dx.ini文件不存在: {d3dx_ini_path}")
                return
            
            # 读取文件内容并提取[Hunting]部分的键位配置
            with open(d3dx_ini_path, 'r', encoding='utf-8') as f:
                lines = f.readlines()
            
            # 查找[Hunting]部分的键位配置
            in_hunting_section = False
            hunting_settings = {}
            
            for line in lines:
                stripped_line = line.strip()
                if stripped_line.lower() == '[hunting]':
                    in_hunting_section = True
                elif in_hunting_section and stripped_line.startswith('['):
                    # 到达下一个section，退出hunting部分
                    in_hunting_section = False
                elif in_hunting_section and '=' in stripped_line:
                    # 解析设置项
                    key, value = stripped_line.split('=', 1)
                    key = key.strip()
                    value = value.strip()
                    hunting_settings[key] = value
            
            # 定义键位功能的中文描述
            key_descriptions = {
                "toggle_hunting": "开启/关闭查找",
                "next_marking_mode": "切换标记模式",
                "done_hunting": "完成查找/刷新索引",
                "previous_pixelshader": "上一个像素着色器",
                "next_pixelshader": "下一个像素着色器",
                "mark_pixelshader": "标记像素着色器",
                "previous_vertexshader": "上一个顶点着色器",
                "next_vertexshader": "下一个顶点着色器",
                "mark_vertexshader": "标记顶点着色器",
                "previous_indexbuffer": "上一个索引缓冲区",
                "next_indexbuffer": "下一个索引缓冲区",
                "mark_indexbuffer": "标记索引缓冲区",
                "previous_vertexbuffer": "上一个顶点缓冲区",
                "next_vertexbuffer": "下一个顶点缓冲区",
                "mark_vertexbuffer": "标记顶点缓冲区",
            }
            
            # 创建新窗口显示键位映射
            mapping_window = tk.Toplevel(self.root)
            mapping_window.title("当前键位配置")
            mapping_window.geometry("600x450")  # 减小窗口宽度
            mapping_window.resizable(True, True)
            
            # 创建框架来容纳表格和滚动条
            frame = ttk.Frame(mapping_window)
            frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
            
            # 创建垂直滚动条
            y_scrollbar = ttk.Scrollbar(frame)
            y_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
            
            # 创建水平滚动条
            x_scrollbar = ttk.Scrollbar(frame, orient=tk.HORIZONTAL)
            x_scrollbar.pack(side=tk.BOTTOM, fill=tk.X)
            
            # 创建表格视图
            tree = ttk.Treeview(
                frame, 
                yscrollcommand=y_scrollbar.set, 
                xscrollcommand=x_scrollbar.set, 
                columns=('功能', '键位'), 
                show='headings'
            )
            
            # 配置列宽和标题
            tree.heading('功能', text='功能描述')
            tree.heading('键位', text='键位设置')
            tree.column('功能', width=250, anchor=tk.W)  # 减小列宽
            tree.column('键位', width=280, anchor=tk.W)  # 减小列宽
            
            # 配置滚动条
            y_scrollbar.config(command=tree.yview)
            x_scrollbar.config(command=tree.xview)
            
            # 填充表格数据 - 只显示key_descriptions字典中包含的按键键位
            if hunting_settings:
                # 先添加固定显示的快捷键映射（显示在最前面）
                tree.insert('', tk.END, values=("帮助页面", "F1"))
                tree.insert('', tk.END, values=("重载配置", "F10"))
                
                # 再显示从配置文件读取的键位映射
                has_displayed_keys = False
                for key, description in key_descriptions.items():
                    if key in hunting_settings:
                        value = hunting_settings[key]
                        # 过滤掉no_modifiers和NO_VK_DECIMAL文字
                        filtered_value = value
                        if "no_modifiers" in filtered_value:
                            filtered_value = filtered_value.replace("no_modifiers", "").strip()
                        if "NO_VK_DECIMAL" in filtered_value:
                            filtered_value = filtered_value.replace("NO_VK_DECIMAL", "").strip()
                        
                        # 处理F功能键，将VK_F1显示为F1，VK_F2显示为F2等
                        for i in range(1, 13):  # 处理F1-F12
                            if f"VK_F{i}" in filtered_value:
                                filtered_value = filtered_value.replace(f"VK_F{i}", f"F{i}")
                        
                        # 处理OEM键位，显示为对应的键盘键位
                        oem_mappings = {
                            "VK_OEM_COMMA": ",",
                            "VK_OEM_PERIOD": ".",
                            "VK_OEM_QUESTION": "?",
                            "VK_OEM_MINUS": "-",
                            "VK_OEM_PLUS": "+",
                            "VK_OEM_1": ";",
                            "VK_OEM_2": "/",
                            "VK_OEM_3": "~",
                            "VK_OEM_4": "[",
                            "VK_OEM_5": "\\",
                            "VK_OEM_6": "]",
                            "VK_OEM_7": "'"
                        }
                        for oem_key, display_char in oem_mappings.items():
                            if oem_key in filtered_value:
                                filtered_value = filtered_value.replace(oem_key, display_char)
                        
                        # 确保不会有多余的空格
                        filtered_value = ' '.join(filtered_value.split())
                        tree.insert('', tk.END, values=(description, filtered_value))
                        has_displayed_keys = True
                
                if not has_displayed_keys:
                    # 即使没有找到匹配的键位，也要显示固定的快捷键
                    tree.column('功能', width=250)
                    tree.column('键位', width=280)
                    tree['show'] = 'headings'
            else:
                # 如果没有找到配置，只显示固定的快捷键
                tree.insert('', tk.END, values=("帮助页面", "F1"))
                tree.insert('', tk.END, values=("重载配置", "F10"))
                tree.column('功能', width=250)
                tree.column('键位', width=280)
                tree['show'] = 'headings'
            
            # 设置表格样式
            style = ttk.Style()
            style.configure("Treeview", font=self.font)
            style.configure("Treeview.Heading", font=(self.font[0], self.font[1], 'bold'))
            
            # 显示表格
            tree.pack(fill=tk.BOTH, expand=True)
            
            # 添加关闭按钮
            close_button = ttk.Button(mapping_window, text="关闭", command=mapping_window.destroy)
            close_button.pack(pady=10)
        except Exception as e:
            self.log_message(f"显示键位配置时发生错误: {str(e)}")
            messagebox.showerror("错误", f"无法显示键位配置: {str(e)}")
    
    def create_widgets(self):
        # 创建顶部框架，包含标题和按钮
        top_frame = ttk.Frame(self.main_frame)
        top_frame.pack(side=tk.TOP, fill=tk.X, pady=5, padx=5)
        
        # 标题标签
        title_label = ttk.Label(top_frame, text="3DMigoto VB数据提取工具", 
                               font=('Microsoft YaHei', 14, 'bold'))
        title_label.pack(side=tk.LEFT, padx=5)
        
        # 创建右上角按钮框架
        top_right_frame = ttk.Frame(top_frame)
        top_right_frame.pack(side=tk.RIGHT, anchor=tk.NE)
        
        # 启动3dmigoto按钮
        self.start_3dmigoto_button = ttk.Button(
            top_right_frame, 
            text="启动3dmigoto", 
            command=self.start_3dmigoto,
            width=12
        )
        self.start_3dmigoto_button.pack(side=tk.RIGHT, padx=2)
        
        # 帮助按钮
        self.help_button = ttk.Button(
            top_right_frame, 
            text="使用帮助", 
            command=self.show_help,
            width=10
        )
        self.help_button.pack(side=tk.RIGHT, padx=2)
        
        # 创建输入区域
        input_frame = ttk.LabelFrame(self.main_frame, text="输入参数", padding="10")
        input_frame.pack(fill=tk.BOTH, expand=True, pady=10)
        
        # 创建表格视图 - show='headings' 仅显示列标题
        self.hash_table = ttk.Treeview(input_frame, columns=('id', 'hash', 'name'), show='headings', height=5)
        
        # 定义列
        
        self.hash_table.heading('id', text='ID')
        self.hash_table.column('id', width=40, anchor=tk.CENTER)
        
        self.hash_table.heading('hash', text='哈希值')
        self.hash_table.column('hash', width=160, anchor=tk.W)  # 减小宽度以适应窗口
        
        self.hash_table.heading('name', text='名称')
        self.hash_table.column('name', width=160, anchor=tk.W)  # 减小宽度以适应窗口
        
        # 绑定双击事件，实现单元格编辑功能
        self.hash_table.bind('<Double-1>', self.on_double_click)
        
        # 放置表格
        self.hash_table.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, pady=5)
        
        # 添加滚动条
        scrollbar = ttk.Scrollbar(input_frame, orient=tk.VERTICAL, command=self.hash_table.yview)
        self.hash_table.configure(yscroll=scrollbar.set)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # 创建表格操作按钮框架
        table_buttons_frame = ttk.Frame(input_frame)
        table_buttons_frame.pack(side=tk.RIGHT, pady=5, padx=5)
        
        # 删除行按钮
        self.delete_row_button = ttk.Button(
            table_buttons_frame, 
            text="删除行", 
            command=self.delete_row,
            width=10
        )
        self.delete_row_button.pack(fill=tk.X, pady=2)
        
        # 初始状态自动添加一行，包含ID
        row_id = 1
        self.hash_table.insert('', tk.END, values=(f'{row_id}', '', ''))
        
        # 创建注入进程路径输入区域
        process_path_frame = ttk.Frame(self.main_frame)
        process_path_frame.pack(fill=tk.X, pady=5, padx=5)
        
        # 注入进程路径标签
        process_path_label = ttk.Label(process_path_frame, text="注入进程路径:", width=12)
        process_path_label.pack(side=tk.LEFT, padx=5)
        
        # 注入进程路径输入框
        self.process_path_var = tk.StringVar(value="")
        self.process_path_entry = ttk.Entry(process_path_frame, textvariable=self.process_path_var)
        self.process_path_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        
        # 浏览按钮
        browse_button = ttk.Button(process_path_frame, text="浏览...", command=self.browse_process_path, width=8)
        browse_button.pack(side=tk.LEFT, padx=5)
        
        # 创建新的框架来放置键位设置，确保单独一行显示
        key_settings_frame = ttk.Frame(self.main_frame)
        key_settings_frame.pack(fill=tk.X, pady=5, padx=5)
        
        # 键位复选框
        self.numpad_checkbox_var = tk.BooleanVar(value=False)  # 默认不使用无小键盘模式
        numpad_checkbox = ttk.Checkbutton(
            key_settings_frame, 
            text="使用无小键盘模式键位", 
            variable=self.numpad_checkbox_var,
            command=self.toggle_key_layout
        )
        numpad_checkbox.pack(side=tk.LEFT, padx=5)
        
        # 查看键位按钮
        self.view_keys_button = ttk.Button(
            key_settings_frame, 
            text="查看键位", 
            command=self.show_key_mappings,
            width=10
        )
        self.view_keys_button.pack(side=tk.LEFT, padx=5)
        
        # 创建按钮区域
        button_frame = ttk.Frame(self.main_frame)
        button_frame.pack(pady=10)
        
        # 执行按钮
        self.run_button = ttk.Button(button_frame, text="执行提取", command=self.run_extraction, 
                                    width=60)  # 增加宽度使按钮更大
        # 设置按钮字体为加粗
        style = ttk.Style()
        style.configure("Bold.TButton", font=(self.font[0], self.font[1], 'bold'))
        self.run_button.configure(style="Bold.TButton")
        self.run_button.pack(side=tk.LEFT, padx=5)
        
        # 创建日志区域
        log_frame = ttk.LabelFrame(self.main_frame, text="操作日志", padding="10")
        log_frame.pack(fill=tk.BOTH, expand=True, pady=10)
        
        self.log_text = tk.Text(log_frame, height=8, width=50, font=self.font, state=tk.DISABLED)
        self.log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        scrollbar = ttk.Scrollbar(log_frame, command=self.log_text.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.log_text.config(yscrollcommand=scrollbar.set)
    
    def detect_keyboard_layout(self):
        """检测d3dx.ini中当前使用的键位配置模式"""
        try:
            # 构建d3dx.ini文件的完整路径
            d3dx_ini_path = os.path.join(self.current_dir, "3dmigoto-GIMI-for-development", "d3dx.ini")
            
            # 检查文件是否存在
            if os.path.exists(d3dx_ini_path):
                # 读取文件内容
                with open(d3dx_ini_path, 'r', encoding='utf-8') as f:
                    lines = f.readlines()
                
                # 查找[Hunting]部分的键位配置
                in_hunting_section = False
                current_settings = {}
                
                for line in lines:
                    stripped_line = line.strip()
                    if stripped_line.lower() == '[hunting]':
                        in_hunting_section = True
                    elif in_hunting_section and stripped_line.startswith('['):
                        # 到达下一个section，退出hunting部分
                        in_hunting_section = False
                    elif in_hunting_section and '=' in stripped_line:
                        # 解析设置项
                        key, value = stripped_line.split('=', 1)
                        key = key.strip().lower()
                        value = value.strip()
                        current_settings[key] = value
                
                # 定义用于检测的关键键位
                key_to_check = "toggle_hunting"  # 使用toggle_hunting作为检测键位
                
                # 如果找到了检测键位
                if key_to_check in current_settings:
                    # 检查是否匹配无小键盘模式的键位
                    # 无小键盘模式的toggle_hunting值为"no_modifiers NO_VK_DECIMAL VK_F12"
                    if current_settings[key_to_check] == "no_modifiers NO_VK_DECIMAL VK_F12":
                        return True
                    else:
                        return False
                else:
                    return False
            else:
                self.log_message(f"d3dx.ini文件不存在: {d3dx_ini_path}")
                return False
        except Exception as e:
            self.log_message(f"检测键位配置时发生错误: {str(e)}")
            return False
    
    def log_message(self, message):
        self.log_text.config(state=tk.NORMAL)
        self.log_text.insert(tk.END, message + "\n")
        self.log_text.see(tk.END)
        self.log_text.config(state=tk.DISABLED)
        
    def add_row(self):
        # 向表格中添加新行，包含ID
        row_id = len(self.hash_table.get_children()) + 1
        self.hash_table.insert('', tk.END, values=(f'{row_id}', '', ''))
        
        # 自动进入编辑模式
        item_id = self.hash_table.get_children()[-1]
        self.hash_table.focus(item_id)
        self.hash_table.selection_set(item_id)
        self.hash_table.see(item_id)
    
    def delete_row(self):
        # 从表格中删除选中的行
        selected_item = self.hash_table.selection()
        items = self.hash_table.get_children()
        # 确保表格中至少有一行且有选中的行
        if selected_item and items:
            # 检查选中的行是否是最后一行
            if selected_item[0] == items[-1]:
                # 不允许删除最后一行预添加的输入行
                return
            
            # 获取选中行的索引位置
            selected_index = items.index(selected_item[0])
            
            self.hash_table.delete(selected_item)
            # 重新编号所有行的ID
            for i, item in enumerate(self.hash_table.get_children(), start=1):
                values = list(self.hash_table.item(item, 'values'))
                values[0] = f'{i}'  # 更新ID值
                self.hash_table.item(item, values=values)
            
            # 获取删除后的所有行
            updated_items = self.hash_table.get_children()
            # 如果还有行，选中上一行
            if updated_items:
                # 计算要选中的行的索引
                # 如果删除的是第一行，则选中新的第一行；否则选中上一行
                new_select_index = min(selected_index - 1, len(updated_items) - 1)
                new_select_index = max(new_select_index, 0)  # 确保索引不小于0
                
                # 选中上一行
                self.hash_table.selection_set(updated_items[new_select_index])
                self.hash_table.focus(updated_items[new_select_index])
    
    def on_double_click(self, event):
        # 双击编辑单元格功能
        # 获取点击的位置信息
        region = self.hash_table.identify_region(event.x, event.y)
        if region != "cell":
            return
        
        # 获取点击的项和列
        item = self.hash_table.identify_row(event.y)
        column = self.hash_table.identify_column(event.x)
        
        if not item or not column:
            return
        
        # 获取列索引
        column_index = int(column.replace('#', ''))
        if column_index == 2:
            column_name = 'hash'
        elif column_index == 3:
            column_name = 'name'
        else:
            return  # 不编辑ID列
        
        # 获取单元格当前值
        current_value = self.hash_table.item(item, 'values')[column_index-1]
        
        # 获取单元格位置
        x, y, width, height = self.hash_table.bbox(item, column)
        
        # 创建输入框
        self.edit_entry = ttk.Entry(self.hash_table)
        self.edit_entry.place(x=x, y=y, width=width, height=height)
        self.edit_entry.insert(0, current_value)
        self.edit_entry.select_range(0, tk.END)
        self.edit_entry.focus()
        
        # 保存当前编辑的项和列信息
        self.edit_entry._item = item
        self.edit_entry._column = column_name
        
        # 绑定事件
        self.edit_entry.bind('<Return>', self.on_entry_confirm)
        self.edit_entry.bind('<FocusOut>', self.on_entry_confirm)
        self.edit_entry.bind('<Escape>', lambda e: self.edit_entry.destroy())
    
    def on_entry_confirm(self, event=None):
        # 确认单元格编辑
        if hasattr(self, 'edit_entry'):
            # 获取输入的值
            new_value = self.edit_entry.get()
            
            # 获取保存的项和列信息
            item = self.edit_entry._item
            column = self.edit_entry._column
            
            # 更新Treeview中的值
            current_values = list(self.hash_table.item(item, 'values'))
            if column == 'hash':
                current_values[1] = new_value  # hash对应索引1
            elif column == 'name':
                current_values[2] = new_value  # name对应索引2
            self.hash_table.item(item, values=current_values)
            
            # 检查当前行的两个字段是否都有值
            # 只有当两个字段都有值时，才在下方添加新行
            if current_values[0] and current_values[1]:
                # 获取当前项在表格中的位置
                items = self.hash_table.get_children()
                current_index = items.index(item)
                
                # 检查是否已经在最后一行，如果是，则直接在末尾添加新行
                if current_index == len(items) - 1:
                    # 检查最后一行是否已经是空行
                    last_item = items[-1]
                    last_values = self.hash_table.item(last_item, 'values')
                    if last_values[0] and last_values[1]:
                        # 最后一行有值，添加新的空行，包含ID
                        row_id = len(self.hash_table.get_children()) + 1
                        self.hash_table.insert('', tk.END, values=(f'{row_id}', '', ''))
            
            # 销毁输入框
            self.edit_entry.destroy()
            delattr(self, 'edit_entry')
            
    def start_3dmigoto(self):
        try:
            # 3dmigoto.exe的路径
            three_dmigoto_path = os.path.join(
                self.current_dir, 
                "3dmigoto-GIMI-for-development", 
                "3DMigoto Loader.exe"
            )
            
            # 检查文件是否存在
            if not os.path.exists(three_dmigoto_path):
                messagebox.showerror("错误", f"找不到3dmigoto Loader.exe文件\n路径: {three_dmigoto_path}")
                return
            
            # 获取3dmigoto.exe所在的目录
            three_dmigoto_dir = os.path.dirname(three_dmigoto_path)
            
            # 启动3dmigoto.exe，并指定工作目录为其自身目录，同时在新窗口中显示输出
            self.log_message("正在启动3dmigoto Loader.exe...")
            subprocess.Popen([three_dmigoto_path], cwd=three_dmigoto_dir, creationflags=subprocess.CREATE_NEW_CONSOLE)
            self.log_message("3dmigoto Loader.exe已启动，输出将显示在新窗口中")
            
        except Exception as e:
            self.log_message(f"启动3dmigoto时发生错误: {str(e)}")
            messagebox.showerror("错误", f"启动3dmigoto时发生错误: {str(e)}")
            
    def run_extraction(self):
        # 从表格中获取所有的哈希值和名称
        items = self.hash_table.get_children()
        hash_name_pairs = []
        
        for item in items:
            # 确保正确解包表格数据
            values = self.hash_table.item(item, "values")
            if len(values) >= 3:
                # 如果有3个值（包含ID），只取哈希值和名称
                hash_value = values[1].strip() if values[1] else ""
                name = values[2].strip() if values[2] else ""
            elif len(values) == 2:
                # 兼容只有两个值的情况
                hash_value = values[0].strip() if values[0] else ""
                name = values[1].strip() if values[1] else ""
            else:
                continue  # 跳过不完整的数据
            
            if hash_value:
                # 如果名称为空，使用哈希值的前8个字符作为名称
                if not name:
                    name = hash_value[:8]
                hash_name_pairs.append((hash_value, name))
        
        # 验证输入 - 当没有任何有效的哈希值时，不能进行提取
        if not hash_name_pairs:
            messagebox.showerror("错误", "请在表格中至少输入一个有效的哈希值")
            return
        
        try:
            # 创建主输出目录
            main_output_folder = "outputs"
            if not os.path.exists(main_output_folder):
                os.makedirs(main_output_folder)
            
            self.log_message(f"开始提取VB数据...")
            self.log_message(f"总共需要处理 {len(hash_name_pairs)} 个哈希值")
            
            # 按顺序处理每个哈希值
            for index, (vb_hash, name) in enumerate(hash_name_pairs):
                self.log_message(f"\n处理项目 {index+1}/{len(hash_name_pairs)}:")
                self.log_message(f"VB哈希值: {vb_hash}")
                self.log_message(f"名称: {name}")
                
                # 创建当前项目的输出目录
                folder_name = os.path.join(main_output_folder, name)
                self.log_message(f"输出文件夹: {folder_name}")
                
                # 重定向标准输出以便在日志中显示
                old_stdout = sys.stdout
                sys.stdout = mystdout = io.StringIO()
                
                try:
                    # 直接调用genshin_3dmigoto_collect模块中的功能
                    # 准备参数
                    frame_dump_folder = "dumps"
                    force_ids = False
                    texture_only = False
                    component_names = None
                    object_classifications = ["Body", "Head", "Hair"]
                    has_normalmap = False
                    
                    # 创建输出目录
                    if not os.path.exists(folder_name):
                        os.makedirs(folder_name)
                    
                    # 调用关键函数执行提取过程
                    # 收集点列表候选
                    pointlist_vb = genshin_3dmigoto_collect.collect_pointlist_candidates(frame_dump_folder)
                    
                    # 收集相关ID
                    relevant_ids = genshin_3dmigoto_collect.collect_relevant_ids(frame_dump_folder, vb_hash, pointlist_vb, force_ids)
                    
                    # 收集模型数据
                    model_data, position_vbs, texcoord_vbs = genshin_3dmigoto_collect.collect_model_data(
                        frame_dump_folder, relevant_ids, texture_only, force_ids
                    )
                    
                    # 处理每个部分的数据
                    for i in range(len(model_data)):
                        current_part = i
                        position_vb = position_vbs[i]
                        texcoord_vb = texcoord_vbs[i]
                        
                        # 收集缓冲区数据
                        position_data, position_format = genshin_3dmigoto_collect.collect_buffer_data(
                            frame_dump_folder, position_vb, ["POSITION0"]
                        )
                        texcoord_data, texcoord_format = genshin_3dmigoto_collect.collect_buffer_data(
                            frame_dump_folder, texcoord_vb, ["TEXCOORD0"]
                        )
                        
                        # 合并缓冲区数据
                        buffer_data = []
                        for j in range(len(position_data)):
                            buffer_data.append(position_data[j] + texcoord_data[j])
                        
                        # 构建合并后的缓冲区
                        element_format = position_format + texcoord_format
                        vb_merged = genshin_3dmigoto_collect.construct_combined_buffer(buffer_data, element_format)
                        
                        # 输出结果
                        genshin_3dmigoto_collect.output_results(
                            frame_dump_folder, folder_name, component_names, model_data, vb_merged, 
                            position_vb, current_part, texture_only, object_classifications, has_normalmap
                        )
                    
                    self.log_message(f"项目 {name} 提取完成!")
                except Exception as e:
                    # 显示详细错误信息
                    self.log_message(f"处理项目 {name} 时发生错误: {str(e)}")
                    self.log_message(f"错误详情:\n{traceback.format_exc()}")
                    messagebox.showerror("错误", f"处理项目 {name} 时发生错误: {str(e)}")
                finally:
                    # 恢复标准输出
                    sys.stdout = old_stdout
                    
                    # 读取并重定向捕获的输出到日志
                    output = mystdout.getvalue()
                    if output:
                        for line in output.strip().split('\n'):
                            self.log_message(line)
            
            self.log_message(f"\n所有项目提取完成!")
            messagebox.showinfo("成功", f"所有项目提取完成! 结果保存在 'outputs' 文件夹中")
        except Exception as e:
            # 显示详细错误信息
            self.log_message(f"执行过程中发生错误: {str(e)}")
            self.log_message(f"错误详情:\n{traceback.format_exc()}")
            messagebox.showerror("错误", f"执行过程中发生错误: {str(e)}")
        finally:
            # 恢复标准输出
            sys.stdout = old_stdout
            
            # 读取并重定向捕获的输出到日志
            output = mystdout.getvalue()
            if output:
                for line in output.strip().split('\n'):
                    self.log_message(line)

    
    def show_help(self):
        help_text = """3DMigoto VB数据提取工具使用说明:

1. 运行程序
   运行项目根目录下的 luncher.py 文件

2. 配置注入路径
   在程序界面中配置正确的3DMigoto注入路径

3. 启动3DMigoto Loader
   使用界面上的按钮启动3DMigoto Loader

4. 启动游戏
   启动需要提取数据的游戏

5. 查找VB哈希值
   使用小键盘上的 7 和 8 键找到所需的VB的哈希值

6. 填入哈希值
   在本工具中:
   - 点击"添加行"按钮添加新的输入行
   - 将查找到的哈希值和对应的文件夹名称填入程序界面的表格中
   - 哈希值：要提取的模型的VB哈希值
   - 名称：用于保存结果的文件夹名称，结果将保存在 'outputs/{名称}' 文件夹中
   - 支持批量添加多个哈希值
   - 哈希值为空时将无法进行提取操作
   - 如果需要删除某一行，可以选中后点击"删除行"按钮

7. 提取原始数据
   在游戏中使用3DMigoto的 F8 键提取数据，并等待提取完成（此过程耗时较久）

8. 处理数据
   在软件界面中点击 "执行提取" 按钮开始处理数据
   处理完成后，结果会保存在对应的文件夹中

9. 【可选】处理贴图
   使用DB工具将结果中的贴图（dds格式）提取并转换成jpg格式

10. 【可选】导入Blender
    在Blender中通过插件导入数据，选择 ib.txt 和 vb.txt 这两个文件导入
    - 3.6版本：使用3dmigoto-GIMI-for-development在github上提供的插件
    - 4.2版本：使用XXMITools

11. 【可选】绑定贴图
    在Blender中将模型与jpg图片贴图进行绑定

注意事项:
1. 确保所有参数都正确填写，哈希值不能为空
2. 程序会自动创建输出目录
3. 遇到问题时，可以查看日志输出获取更多信息"""
        
        messagebox.showinfo("使用帮助", help_text)

if __name__ == "__main__":
    root = tk.Tk()
    app = VBExtractorApp(root)
    root.mainloop()