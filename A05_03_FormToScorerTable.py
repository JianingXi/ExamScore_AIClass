import pandas as pd

def generate_final_score_excel(
    input_path: str,
    output_path: str,
    delete_col_index: int = 2,
    selected_col_indices: list[int] = [1, 2, 5, 6, 8, 10],
    merge_col_indices: list[int] = [3, 4, 7, 9, 11]
):
    """
    从输入 Excel 中：
    1️⃣ 删除指定列（默认索引 2）
    2️⃣ 选择指定列，拼接为新 DataFrame
    3️⃣ 添加空白列（论文分、视频分、总分）
    4️⃣ 合并若干列内容为【最终评语】
    5️⃣ 输出到新的 Excel 文件

    参数：
    - input_path: 输入文件路径
    - output_path: 输出文件路径
    - delete_col_index: 要删除的列索引
    - selected_col_indices: 新表格要添加的列索引（相对删除后的 DataFrame）
    - merge_col_indices: 要拼接合并的列索引
    """
    df = pd.read_excel(input_path, header=0)

    # 删除指定列
    df.drop(df.columns[delete_col_index], axis=1, inplace=True)

    # 创建新 DataFrame（第1列是原第1列）
    new_df = df.iloc[:, [0]].copy()

    # 新表格的第2-7列
    new_df = pd.concat([new_df, df.iloc[:, selected_col_indices]], axis=1)

    # 添加空白列（论文分、视频分、总分）
    new_df.insert(7, '论文分', '')
    new_df.insert(8, '视频分', '')
    new_df.insert(9, '总分', '')

    # 合并指定列内容（去掉 nan/空值，直接拼接）
    def concat_no_space(row):
        parts = [str(x) for x in row if pd.notnull(x)]
        return ''.join(parts)

    merged_column = df.iloc[:, merge_col_indices].apply(concat_no_space, axis=1)

    # 插入合并列
    new_df.insert(10, '最终评语', merged_column)

    # 保存新 Excel 文件
    new_df.to_excel(output_path, index=False)
    print(f'已生成定制好的文件：{output_path}')

def a05_03_form_to_score_table(input_file: str, output_file: str):
    # input_file = r'C:\MyDocument\ToDoList\D20_DoingPlatform\D20_医学人工智能\结课论文\MergedResult_QualitativeDescriptions.xlsx'
    # output_file = r'C:\MyDocument\ToDoList\D20_DoingPlatform\D20_医学人工智能\结课论文\成绩表_明细版.xlsx'
    generate_final_score_excel(input_file, output_file)
