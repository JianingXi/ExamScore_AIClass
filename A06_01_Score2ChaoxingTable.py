import pandas as pd
import pyexcel as pe
import os

def generate_sorted_score_table(
    score_path: str,
    final_xls_path: str,
    output_path: str
):
    """
    1️⃣ 将 .xls 转为 .xlsx
    2️⃣ 读取成绩表，提取学号、姓名、成绩、评语
    3️⃣ 按照基准顺序整理成绩，若无成绩则空值
    4️⃣ 输出整理后的成绩表

    参数:
    - score_path: 成绩表 Excel 文件路径（含列“学生文件夹”）
    - final_xls_path: 基准顺序 .xls 文件路径（含列“学号/工号”和“学生姓名”）
    - output_path: 输出整理好的文件路径（.xlsx）
    """
    # 1️⃣ 将 .xls 转为 .xlsx（如果已存在同名 .xlsx 文件可跳过转换）
    final_xlsx_path = os.path.splitext(final_xls_path)[0] + '.xlsx'
    if not os.path.exists(final_xlsx_path):
        sheet = pe.get_sheet(file_name=final_xls_path)
        sheet.save_as(final_xlsx_path)
        print("已将 .xls 转换为 .xlsx 文件：", final_xlsx_path)
    else:
        print("已存在 .xlsx 文件：", final_xlsx_path)

    # 2️⃣ 读取成绩表
    score_df = pd.read_excel(score_path, header=0)
    # 拆分学号和姓名
    score_df['学号'] = score_df['学生文件夹'].apply(lambda x: str(x).split('-')[0] if '-' in str(x) else '')
    score_df['姓名'] = score_df['学生文件夹'].apply(lambda x: str(x).split('-')[1] if '-' in str(x) else '')
    # 提取所需列
    score_df_cleaned = score_df[['学号', '姓名', '总分', '最终评语']]
    score_df_cleaned.columns = ['学号', '姓名', '成绩', '评语']

    # 3️⃣ 读取基准顺序表（正式数据从第2行开始）
    final_df = pd.read_excel(final_xlsx_path, header=0)
    final_df_data = final_df.iloc[1:].reset_index(drop=True)
    final_df_data.columns = final_df.iloc[0].str.strip()  # 第1行作为新列名

    final_student_ids = final_df_data['学号/工号'].astype(str).tolist()
    final_student_names = final_df_data['学生姓名'].astype(str).tolist()

    # 4️⃣ 按照基准顺序整理成绩
    new_records = []
    for stu_id, stu_name in zip(final_student_ids, final_student_names):
        matched = score_df_cleaned[score_df_cleaned['学号'] == stu_id]
        if not matched.empty:
            row = matched.iloc[0]
        else:
            row = pd.Series({'学号': stu_id, '姓名': stu_name, '成绩': '', '评语': ''})
        new_records.append(row)

    new_score_df = pd.DataFrame(new_records)

    # 5️⃣ 输出结果文件
    new_score_df.to_excel(output_path, index=False)
    print(f'整理后的成绩表已保存至：{output_path}')

def a06_01_score2chaoxing_table(score_file: str, final_xls_file: str, output_file: str):
    # score_file = r'C:\MyDocument\ToDoList\D20_DoingPlatform\D20_医学人工智能\结课论文\成绩表_明细版.xlsx'
    # final_xls_file = r'C:\MyDocument\ToDoList\D20_DoingPlatform\D20_医学人工智能\结课论文\【Final】结课论文和结课汇报.xls'
    # output_file = r'C:\MyDocument\ToDoList\D20_DoingPlatform\D20_医学人工智能\结课论文\整理后的成绩表.xlsx'
    generate_sorted_score_table(score_file, final_xls_file, output_file)
