import os
import shutil
import csv
import tkinter as tk
from tkinter import filedialog, messagebox
import re
import difflib

# 尝试导入 openpyxl 用于支持 xlsx 文件
try:
    import openpyxl
    HAS_OPENPYXL = True
except ImportError:
    HAS_OPENPYXL = False

def get_file_path(title, filetypes):
    """弹出文件选择框获取文件路径"""
    filepath = filedialog.askopenfilename(title=title, filetypes=filetypes)
    return filepath

def get_folder_path(title):
    """弹出文件夹选择框获取文件夹路径"""
    folderpath = filedialog.askdirectory(title=title)
    return folderpath

def clean_name(name):
    """清洗文件名：去除扩展名、特殊标点、空格等干扰字符，只保留核心中英文数字"""
    name = os.path.splitext(name)[0]  # 去除扩展名
    name = re.sub(r'[^\w\u4e00-\u9fa5]', '', name)  # 只保留字母、数字、下划线、汉字
    return name

def find_best_match(target_name, available_files):
    """在可用文件池中寻找最佳的模糊匹配"""
    # 1. 尝试完全一致 (忽略后缀名)
    for f in available_files:
        if f == target_name or os.path.splitext(f)[0] == os.path.splitext(target_name)[0]:
            return f
            
    # 2. 模糊匹配 (提取核心字符比对)
    clean_target = clean_name(target_name)
    if not clean_target:
        return None
        
    best_match = None
    max_score = 0
    set_target = set(clean_target)
    
    for f in available_files:
        clean_f = clean_name(f)
        if not clean_f: continue
        
        # 指标1：字符覆盖率 (解决乱序问题，比如 16北正道 vs 正道16北)
        set_f = set(clean_f)
        common_chars = set_target & set_f
        coverage = len(common_chars) / len(set_target) if set_target else 0
        
        # 指标2：连续序列相似度 (解决少量错别字)
        seq_ratio = difflib.SequenceMatcher(None, clean_target, clean_f).ratio()
        
        # 指标3：直接包含关系 (比如 16北正道 包含于 2月16北正道投资)
        is_subset = (clean_target in clean_f) or (clean_f in clean_target)
        len_valid = min(len(clean_target), len(clean_f)) >= 3
        
        # 综合打分
        score = (coverage * 0.5) + (seq_ratio * 0.5)
        if is_subset and len_valid:
            score += 0.3  # 互相包含则大幅加权
            
        if score > max_score:
            max_score = score
            best_match = f
            
    # 设定阈值：分数大于 0.7 认为可以信赖
    if max_score >= 0.7:
        return best_match
        
    return None

def read_excel_rules(file_path):
    """读取 xlsx 格式的规则表"""
    expected_files = {}
    wb = openpyxl.load_workbook(file_path, data_only=True)
    sheet = wb.active
    
    col_mapping = {}
    
    # 1. 遍历第一行获取表头（分类名）
    for col_idx, cell in enumerate(sheet[1]):
        if cell.value is not None:
            header = str(cell.value).strip()
            if header:  # 过滤空表头
                col_mapping[col_idx] = header
                expected_files[header] = []
                
    # 2. 遍历剩余行获取文件名
    for row in sheet.iter_rows(min_row=2):
        for col_idx, cell in enumerate(row):
            if col_idx in col_mapping and cell.value is not None:
                val = str(cell.value).strip()
                if val: # 过滤空白单元格
                    expected_files[col_mapping[col_idx]].append(val)
                    
    return expected_files

def read_csv_rules(file_path):
    """健壮地读取 CSV 格式的规则表 (支持多种编码 fallback)"""
    expected_files = {}
    encodings_to_try = ['utf-8-sig', 'gbk', 'gb18030', 'utf-8', 'gb2312']
    
    for enc in encodings_to_try:
        try:
            with open(file_path, 'r', encoding=enc) as f:
                reader = csv.reader(f)
                headers = next(reader)
                
                col_mapping = {}
                for i, header in enumerate(headers):
                    header = header.strip()
                    if header:
                        col_mapping[i] = header
                        expected_files[header] = []
                        
                for row in reader:
                    for i, cell in enumerate(row):
                        cell = cell.strip()
                        if i in col_mapping and cell:
                            expected_files[col_mapping[i]].append(cell)
            return expected_files # 成功读取则立即返回
        except UnicodeDecodeError:
            continue # 如果当前编码失败，尝试下一种
        except StopIteration:
            return {} # 空文件
            
    raise ValueError(f"文件编码不支持，无法读取CSV文件: {file_path}")

def main():
    root = tk.Tk()
    root.withdraw()

    # 1. 选择规则表（默认支持 xlsx 和 csv）
    messagebox.showinfo("步骤 1/2", "请选择你的【分类规则表】(支持 Excel 或 CSV)")
    rule_path = get_file_path("选择分类规则表", [("表格文件", "*.xlsx;*.csv"), ("Excel 文件", "*.xlsx"), ("CSV 文件", "*.csv"), ("所有文件", "*.*")])
    if not rule_path:
        print("未选择分类规则表，程序退出。")
        return

    # 2. 选择源文件夹
    messagebox.showinfo("步骤 2/2", "请选择存放电表数据的【源文件夹】")
    source_dir = get_folder_path("选择源文件夹")
    if not source_dir:
        print("未选择源文件夹，程序退出。")
        return

    print(f"正在读取规则表: {rule_path}")
    print(f"目标处理文件夹: {source_dir}\n")

    # 3. 根据文件扩展名，使用不同的方式读取规则
    ext = os.path.splitext(rule_path)[1].lower()
    expected_files = {}
    
    try:
        if ext == '.xlsx':
            if not HAS_OPENPYXL:
                messagebox.showerror("缺少依赖", "读取 Excel 文件需要 openpyxl 库。\n请在终端运行: pip install openpyxl")
                return
            expected_files = read_excel_rules(rule_path)
        elif ext == '.csv':
            expected_files = read_csv_rules(rule_path)
        else:
            messagebox.showerror("格式错误", f"不支持的文件格式: {ext}。\n请提供 .xlsx 或 .csv 格式的文件。")
            return
    except Exception as e:
        messagebox.showerror("读取规则表失败", f"发生了错误:\n{str(e)}")
        return

    if not expected_files:
        messagebox.showwarning("空数据", "在规则表中没有读取到任何有效的分类数据。")
        return

    # 4. 执行文件分类转移
    missing_files = [] 
    success_count = 0  
    
    # --- 新增：获取源文件夹中所有真实存在的待选文件池 ---
    available_files = []
    for f in os.listdir(source_dir):
        if os.path.isfile(os.path.join(source_dir, f)):
            available_files.append(f)
            
    # 记录那些通过模糊匹配上的文件，方便后期人工核对是否误判
    fuzzy_matched_records = []

    for category, file_list in expected_files.items():
        # 清除类别名称中可能导致文件夹创建失败的特殊字符
        safe_category = "".join(c for c in category if c not in r'\/:*?"<>|')
        target_dir = os.path.join(source_dir, safe_category)
        os.makedirs(target_dir, exist_ok=True)
        print(f"--> 已创建/确认文件夹: {safe_category} (共需处理 {len(file_list)} 个文件)")

        for filename in file_list:
            # --- 使用智能模糊搜索寻找匹配项 ---
            matched_real_file = find_best_match(filename, available_files)

            if matched_real_file:
                src_file_path = os.path.join(source_dir, matched_real_file)
                target_file_path = os.path.join(target_dir, matched_real_file)
                
                # 防重复复制保护
                if not os.path.exists(target_file_path):
                    shutil.copy2(src_file_path, target_file_path)
                success_count += 1
                
                # 如果是名字有出入但被算法抓中的，记录下来
                if clean_name(filename) != clean_name(matched_real_file):
                    fuzzy_matched_records.append(f"表格写: [{filename}] --> 实际匹配: [{matched_real_file}]")
                    
                # 【关键】文件被认领后，从待选池中剔除，防止被其他相似名字重复抓取
                if matched_real_file in available_files:
                    available_files.remove(matched_real_file)
            else:
                missing_files.append(f"类别【{safe_category}】缺失文件: {filename}")

    # 5. 输出结果与缺失/审查清单
    report_msg = f"分类完成！\n成功匹配并复制文件: {success_count} 个。\n"
    
    # 生成综合报告
    if missing_files or fuzzy_matched_records:
        report_path = os.path.join(source_dir, "处理报告清单.txt")
        with open(report_path, 'w', encoding='utf-8') as rf:
            if missing_files:
                rf.write("【一、 完全未找到的文件】 (请检查是否漏拷或名字差异过大):\n")
                rf.write("-" * 60 + "\n")
                for missing in missing_files:
                    rf.write(missing + "\n")
                rf.write("\n\n")
                
            if fuzzy_matched_records:
                rf.write("【二、 智能模糊匹配记录】 (系统觉得它们是同一个文件，请扫一眼确认没抓错):\n")
                rf.write("-" * 60 + "\n")
                for record in fuzzy_matched_records:
                    rf.write(record + "\n")
                    
        report_msg += f"\n生成了详细报告，请在源文件夹查看:\n【处理报告清单.txt】"
        print(f"\n[完成] 报告已保存至: {report_path}")
    else:
        report_msg += "完美！表格中的所有文件都精准无误地找到了。"
        print("\n[完美] 所有文件均已成功分类。")

    messagebox.showinfo("处理完成", report_msg)

if __name__ == "__main__":
    main()