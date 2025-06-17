import pandas as pd
import re
import os

def extract_info_from_filename(filename):
    """从文件名中提取学号和姓名"""
    match = re.search(r'(?<!\d)(\d{6,12})-([\u4e00-\u9fa5]{2,4})', filename)
    if match:
        return match.group(1), match.group(2)
    return None, None


import pandas as pd
import os


def reorder_score_file(score_file, order_file, output_file):
    # 读取评分结果
    df_score = pd.read_excel(score_file)
    df_score['学号'] = None
    df_score['姓名'] = None

    for idx, row in df_score.iterrows():
        filename = str(row.iloc[0])  # ⚠️修复 FutureWarning：用 iloc 更安全
        sid, name = extract_info_from_filename(filename)
        df_score.at[idx, '学号'] = sid
        df_score.at[idx, '姓名'] = name

    # 读取排序模板（处理掉前几行无用数据）
    df_order_raw = pd.read_excel(order_file, header=None)
    df_order_raw.columns = [f"Unnamed: {i}" for i in range(len(df_order_raw.columns))]

    name_col, id_col = None, None
    for i in range(min(10, len(df_order_raw))):
        for j in range(len(df_order_raw.columns)):
            val = str(df_order_raw.iat[i, j])
            if '姓名' in val:
                name_col = val
            if '学号' in val:
                id_col = val
        if name_col and id_col:
            df_order = df_order_raw.iloc[i+1:].copy()
            df_order.columns = df_order_raw.iloc[i]
            break

    if not (name_col and id_col):
        raise ValueError("❌ 排序模板中必须同时包含‘姓名’和‘学号’列")

    df_order = df_order[[id_col, name_col]].copy()
    df_order.columns = ['学号', '姓名']
    df_order['学号'] = df_order['学号'].astype(str).str.strip()
    df_order['姓名'] = df_order['姓名'].astype(str).str.strip()

    df_score['学号'] = df_score['学号'].astype(str).str.strip()
    df_score['姓名'] = df_score['姓名'].astype(str).str.strip()

    # 合并（以“学号+姓名”双键匹配）
    df_merge = pd.merge(df_order, df_score, on=['学号', '姓名'], how='left')

    # 输出
    os.makedirs(os.path.dirname(output_file), exist_ok=True)
    df_merge.to_excel(output_file, index=False)
    print(f"✅ 排序完成，输出文件：{output_file}")




import pandas as pd
import openpyxl

def remove_second_row_completely(file_path, output_path=None):
    # 读取文件，保留原始样式
    wb = openpyxl.load_workbook(file_path)
    sheet = wb.active

    # 删除第二行（openpyxl中行编号从1开始）
    sheet.delete_rows(2)

    # 保存
    if not output_path:
        output_path = file_path.replace('.xlsx', '_修正.xlsx')
    wb.save(output_path)
    print(f"✅ 第二行已彻底删除，保存至：{output_path}")




# 替换为你实际路径
reorder_score_file(
    score_file=r"C:\MyDocument\ToDoList\D20_DoingPlatform\D20_人工智能与大数据\23临床药学-2024-2025春季期末考试A卷(word)\学生评分结果.xlsx",
    order_file=r"C:\MyDocument\ToDoList\D20_DoingPlatform\D20_人工智能与大数据\23临床药学-2024-2025春季期末考试A卷(word)\23临床药学-2024-2025春季期末考试A卷.xlsx",
    output_file=r"C:\MyDocument\ToDoList\D20_DoingPlatform\D20_人工智能与大数据\23临床药学-2024-2025春季期末考试A卷(word)\学生评分结果_已排序.xlsx"
)

# 示例调用
remove_second_row_completely(
    file_path=r"C:\MyDocument\ToDoList\D20_DoingPlatform\D20_人工智能与大数据\23临床药学-2024-2025春季期末考试A卷(word)\学生评分结果_已排序.xlsx"
)

