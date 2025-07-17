import pandas as pd
from pathlib import Path

def overwrite_first_col_with_id_name(folder_path: str):
    folder = Path(folder_path)
    excel_files = list(folder.glob("*.xlsx"))

    for file in excel_files:
        try:
            df = pd.read_excel(file)

            if df.shape[1] == 0:
                print(f"⚠️ 文件无内容：{file.name}")
                continue

            first_col = df.columns[0]

            # 替换第一列值：保留前两个用下划线分隔的部分
            df[first_col] = df[first_col].apply(
                lambda x: "_".join(str(x).split("_")[:2]) if isinstance(x, str) and "_" in x else x
            )

            # 覆盖原文件
            df.to_excel(file, index=False)
            print(f"✅ 已覆盖原文件：{file.name}")

        except Exception as e:
            print(f"❌ 处理失败：{file.name}，原因：{e}")




def overwrite_excels_by_union(file1_path, file2_path):
    file1 = Path(file1_path)
    file2 = Path(file2_path)

    # 读取两个文件
    df1 = pd.read_excel(file1)
    df2 = pd.read_excel(file2)

    key_col = df1.columns[0]  # 默认第一列为索引键

    # 处理重复值：保留第一次出现的记录
    df1 = df1.drop_duplicates(subset=key_col, keep="first")
    df2 = df2.drop_duplicates(subset=key_col, keep="first")

    # 取并集
    union_keys = sorted(set(df1[key_col]).union(set(df2[key_col])))

    # 设置索引
    df1_indexed = df1.set_index(key_col)
    df2_indexed = df2.set_index(key_col)

    # 扩展并填充0
    df1_aligned = df1_indexed.reindex(union_keys, fill_value=0).reset_index()
    df2_aligned = df2_indexed.reindex(union_keys, fill_value=0).reset_index()

    # 覆盖保存
    df1_aligned.to_excel(file1, index=False)
    df2_aligned.to_excel(file2, index=False)

    print(f"✅ 已覆盖并对齐：\n - {file1.name}\n - {file2.name}")



def compare_excel_rows(file1_path, file2_path, output_path, label_file1="file1", label_file2="file2"):
    # 读取两个Excel文件
    df1 = pd.read_excel(file1_path)
    df2 = pd.read_excel(file2_path)

    # 检查行数是否相等
    if len(df1) != len(df2):
        raise ValueError("两个文件的行数不一致，无法逐行比较。")

    merged_rows = []
    for i in range(len(df1)):
        row1 = df1.iloc[i]
        row2 = df2.iloc[i]
        score1 = row1.iloc[-1]
        score2 = row2.iloc[-1]

        if pd.isna(score1) and pd.isna(score2):
            selected_row = row1
            source = label_file1
        elif pd.isna(score1):
            selected_row = row2
            source = label_file2
        elif pd.isna(score2):
            selected_row = row1
            source = label_file1
        elif score1 >= score2:
            selected_row = row1
            source = label_file1
        else:
            selected_row = row2
            source = label_file2

        row_with_source = selected_row.to_dict()
        row_with_source["来源"] = source
        merged_rows.append(row_with_source)

    merged_df = pd.DataFrame(merged_rows)
    merged_df.to_excel(output_path, index=False)
    print(f"✅ 已保存合并结果到：{output_path}")




# 分数和评语分开
def rearrange_columns(file_path):
    # 读取 Excel 文件
    df = pd.read_excel(file_path)

    # 保留第0列（学号/姓名）
    first_column = df.columns[0]
    first_column_data = df[first_column]

    # 拆分其余的列
    odd_columns = df.columns[1::2]  # 奇数列
    even_columns = df.columns[2::2]  # 偶数列

    # 重新排列：奇数列放前，偶数列放后
    new_order = [first_column] + list(odd_columns) + list(even_columns)

    # 重排列
    df_reordered = df[new_order]

    # 将结果保存回原文件（覆盖原文件）
    df_reordered.to_excel(file_path, index=False)
    print(f"✅ 已处理并覆盖文件：{file_path}")



# page1-X 合并


def merge_excels_by_first_column(file_dir: str, file_names: list, output_name="merged_result.xlsx"):
    dfs = []
    all_keys = set()

    for file_name in file_names:
        file_path = Path(file_dir) / file_name
        df = pd.read_excel(file_path)
        df = df.drop_duplicates(subset=df.columns[0])  # 去除重复 key
        df = df.set_index(df.columns[0])               # 设置第一列为索引
        dfs.append(df)
        all_keys.update(df.index.tolist())

    all_keys = sorted(all_keys)  # 保证输出顺序一致

    # 对每个df进行 reindex，使所有 key 对齐，缺失填0
    aligned_dfs = [df.reindex(all_keys, fill_value=0) for df in dfs]

    # 合并所有内容（横向拼接）
    merged_df = pd.concat(aligned_dfs, axis=1)

    # 添加“学号_姓名”列回第一列
    merged_df.insert(0, "学号_姓名", merged_df.index)

    # 保存输出
    output_path = Path(file_dir) / output_name
    merged_df.to_excel(output_path, index=False)
    print(f"✅ 合并完成，输出文件：{output_path}")
