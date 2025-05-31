import pandas as pd
import os

def merge_and_clean_excel_files(folder_path: str, file_names: list[str], output_file: str):
    """
    读取指定文件夹中的多个 Excel 文件：
    1️⃣ 将所有工作表拼接
    2️⃣ 删除第一列中斜杠及其后内容
    3️⃣ 按第一列进行左右拼接
    4️⃣ 拼接后，再对所有列的斜杠及其后内容进行彻底清洗
    5️⃣ 保存最终结果到输出 Excel 文件
    """
    def clean_excel(file_path):
        excel = pd.ExcelFile(file_path)
        df_all = []
        for sheet in excel.sheet_names:
            df = excel.parse(sheet)
            if not df.empty:
                df.iloc[:, 0] = df.iloc[:, 0].astype(str).str.replace(r'/.*', '', regex=True)
                df_all.append(df)
        df_concat = pd.concat(df_all, axis=0, ignore_index=True)
        return df_concat

    # 读取并初步清洗
    dfs = []
    for file_name in file_names:
        file_path = os.path.join(folder_path, file_name)
        df = clean_excel(file_path)
        dfs.append(df)

    # 按第一列进行左右拼接
    df_merged = dfs[0]
    for df in dfs[1:]:
        df_merged = pd.merge(df_merged, df, how='outer',
                              left_on=df_merged.columns[0], right_on=df.columns[0])

    # 全面再清洗一次：删除所有列中斜杠及其后内容
    df_merged_cleaned = df_merged.replace(r'/.*', '', regex=True)

    # 保存结果
    output_path = os.path.join(folder_path, output_file)
    df_merged_cleaned.to_excel(output_path, index=False)
    print("拼接和彻底清洗完成！所有斜杠及其后内容都已删除，已保存到：", output_path)

def a05_01_concatenate(folder: str, output_file: str):
    # folder = r"C:\MyDocument\ToDoList\D20_DoingPlatform\D20_医学人工智能\结课论文"
    files = ['ContentSemanticScore.xlsx', 'MissingSections_Report.xlsx', 'AllScores.xlsx']
    merge_and_clean_excel_files(folder, files, output_file)
