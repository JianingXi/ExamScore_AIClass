import os
import re
import pandas as pd
import numpy as np
from functools import reduce
import shutil
from pathlib import Path


def clean_text(text):
    return re.sub(r'[\s\u3000]', '', text)


def longest_common_substring(s1, s2):
    m, n = len(s1), len(s2)
    dp = [[0] * (n + 1) for _ in range(m + 1)]
    max_len = 0
    for i in range(m):
        for j in range(n):
            if s1[i] == s2[j]:
                dp[i + 1][j + 1] = dp[i][j] + 1
                max_len = max(max_len, dp[i + 1][j + 1])
    return max_len


def split_into_segments(text):
    raw_segments = re.split(r'[。！？\n]+', text)
    segments = [seg.strip() for seg in raw_segments if len(seg.strip()) >= 10]
    return segments


def compare_segments_recursive(txt_root_folder, html_path, match_suffix, threshold=0.5):
    with open(html_path, 'r', encoding='utf-8') as f:
        html_content = clean_text(f.read())

    result_summary = {}
    result_table = []

    for root, dirs, files in os.walk(txt_root_folder):
        for file in files:
            if file.endswith(match_suffix):
                full_path = os.path.join(root, file)
                with open(full_path, 'r', encoding='utf-8') as f:
                    txt_content = f.read()

                segments = split_into_segments(txt_content)
                similarities = []

                # print(f"\n文件：{file}（共 {len(segments)} 段） | 路径：{full_path}")

                for idx, segment in enumerate(segments, 1):
                    cleaned_seg = clean_text(segment)
                    lcs_len = longest_common_substring(cleaned_seg, html_content)
                    similarity = lcs_len / max(len(cleaned_seg), 1)
                    similarities.append(similarity)
                    preview = segment[:30].replace('\n', '')
                    # print(f"  段 {idx:02d} 相似度：{similarity:.2%} | 内容前30字：{preview}")

                if similarities:
                    max_sim = max(similarities)
                    avg_sim = sum(similarities) / len(similarities)
                    count_over_threshold = sum(1 for s in similarities if s >= threshold)

                    result_summary[full_path] = {
                        'max_similarity': round(max_sim, 4),
                        'avg_similarity': round(avg_sim, 4),
                        'match_count': count_over_threshold,
                        'total_segments': len(similarities)
                    }

                    result_table.append({
                        '文件名': file,
                        '平均相似度': round(avg_sim, 4)
                    })

    df = pd.DataFrame(result_table)
    return result_summary, df


def build_per_question_similarity_matrix(txt_folder_path, html_path, skip_question, num_question):
    result_list = []
    for i in range(skip_question + 1, skip_question + num_question + 1):
        suffix = f'_0{i:02}.txt'
        _, similarity_df = compare_segments_recursive(txt_folder_path, html_path, match_suffix=suffix)

        if similarity_df is None or similarity_df.empty:
            print(f"跳过空文件：{suffix}")
            continue

        df = similarity_df.copy().reset_index(drop=True)

        # 自动找出文件名列
        filename_col = None
        for col in df.columns:
            if df[col].astype(str).str.contains(r'\.txt$').any():
                filename_col = col
                break
        if filename_col is None:
            raise ValueError(f"{suffix} 没有发现文件名列")

        # 得分列（选项）
        score_cols = [col for col in df.columns if col != filename_col]

        # 获取 basename 并建立新 df
        basename = df[filename_col].apply(lambda x: re.sub(r'_0\d{2}\.txt$', '', str(x)))
        score_df = df[score_cols].copy()
        score_df.columns = [f"{i:02}_{col}" for col in score_cols]  # 加题号前缀

        score_df.insert(0, 'basename', basename)  # 加上 basename

        result_list.append(score_df)

    if not result_list:
        print("⚠️ 所有文件为空")
        return pd.DataFrame()

    # merge 所有题的得分列（按 basename 对齐）
    final_df = reduce(lambda left, right: pd.merge(left, right, on='basename', how='outer'), result_list)
    final_df.sort_values(by='basename', inplace=True)
    return final_df


def add_max_column_name(df):
    score_only = df.iloc[:, 1:]  # 除 basename 外的所有题目+选项列
    max_col = score_only.idxmax(axis=1)  # 找出每行最大值所在列名
    df['max_col'] = max_col
    return df


def build_similarity_tensor(exam_file_temp, html_path, html_file_base_name, skip_question, num_question):
    df_list = []

    for i_quiz in range(len(html_file_base_name)):
        html_path_temp = os.path.join(html_path, html_file_base_name[i_quiz] + ".html")

        # 构造矩阵（相似度值）
        matrix = build_per_question_similarity_matrix(exam_file_temp, html_path_temp, skip_question, num_question)

        # 去除 max_col，只保留 basename 和各题列
        matrix_clean = matrix.drop(columns=['max_col'], errors='ignore')
        df_list.append(matrix_clean)

    # 获取统一的 basename 和题号列表
    all_basenames = sorted(set().union(*[set(df['basename']) for df in df_list]))
    all_questions = df_list[0].columns[1:]  # 第一列是 basename，后面是 10,11,...

    # 初始化张量 (学生数 × 题数 × html版本数)
    tensor = np.full((len(all_basenames), len(all_questions), len(df_list)), np.nan)

    basename_to_idx = {b: i for i, b in enumerate(all_basenames)}
    question_to_idx = {q: i for i, q in enumerate(all_questions)}

    for k, df in enumerate(df_list):
        for _, row in df.iterrows():
            bname = row['basename']
            for q in all_questions:
                tensor[basename_to_idx[bname], question_to_idx[q], k] = row[q]

    return tensor, all_basenames, all_questions


def find_max_similarity_question_indices(tensor):
    """
    对于每位学生 i 和每个 html版本 v，返回相似度最高的题目索引 q。
    输出 shape 为 (N, V)
    """
    N, Q, V = tensor.shape
    result = []

    for i in range(N):
        max_indices = []
        for v in range(V):
            col = tensor[i, :, v]
            if np.isnan(col).all():
                max_indices.append(-1)
            else:
                max_indices.append(int(np.nanargmax(col)))
        result.append(max_indices)

    return pd.DataFrame(result)


def sim_match_best_ans(txt_folder_path, exam_file_base_name, html_path, html_file_base_name, skip_question, num_question):
    for i_class in range(len(exam_file_base_name)):
        exam_file_temp = os.path.join(txt_folder_path, exam_file_base_name[i_class])
        print(exam_file_temp)

        try:
            tensor, students, questions = build_similarity_tensor(
                exam_file_temp, html_path, html_file_base_name, skip_question, num_question)
        except Exception as e:
            print(f"❌ 构建相似度张量失败：{exam_file_temp}")
            print(html_path)
            print(html_file_base_name)
            print(skip_question)
            print(num_question)
            print(f"错误详情：{e}")
            continue  # ✅ 跳过当前循环，不再执行后续逻辑

        question_max_df = find_max_similarity_question_indices(tensor) + skip_question + 1
        question_max_df.columns = [f"page{i + 1}" for i in range(question_max_df.shape[1])]
        question_max_df.insert(0, "basename", students)

        # 保存为 Excel 文件
        excel_file_name = os.path.join(txt_folder_path, exam_file_base_name[i_class] + "_相似性.xlsx")
        question_max_df.to_excel(excel_file_name, index=False)



# 防止中英文括号不匹配
def normalize_parentheses_pattern(text: str) -> str:
    escaped = re.escape(text)
    return re.sub(r"\\[（(].+?\\[)）]", r"[(（][^()（）]+[)）]", escaped)



# move best match file according to excel dictionary
def copy_and_rename_by_similarity_excel(
    input_folder: str,
    output_folder: str,
    excel_path: str
):
    """
    递归扫描 input_folder 中所有 _0XX.txt 文件，根据 excel_path 中的相似性表格，
    找出对应的 pageX，复制文件并将其命名为 _pageX.txt 到 output_folder。
    """

    os.makedirs(output_folder, exist_ok=True)
    df = pd.read_excel(excel_path)

    # 自动识别 pageX 列
    page_cols = [col for col in df.columns if str(col).startswith("page")]

    # 递归查找所有 .txt 文件
    for root, _, files in os.walk(input_folder):
        for filename in files:
            if not filename.endswith(".txt") or "_0" not in filename:
                continue

            full_path = os.path.join(root, filename)
            basename_match = filename.rsplit("_0", 1)[0]
            suffix = filename[-7:-4]  # 取 '013'

            # 查找 DataFrame 匹配的 basename 行  row = df[df["basename"].str.contains(basename_match)]
            pattern = normalize_parentheses_pattern(basename_match)
            row = df[df["basename"].str.contains(pattern, regex=True, na=False)]

            if row.empty:
                print(f"❌ 未找到匹配 basename: {basename_match}")
                continue

            row = row.iloc[0]
            new_page = None

            for col in page_cols:
                if str(row[col]).zfill(3) == suffix:
                    new_page = col
                    break

            if new_page is None:
                print(f"❗ 未匹配题号 {suffix} 于 {basename_match}")
                continue

            new_filename = f"{basename_match}_{new_page}.txt"
            new_path = os.path.join(output_folder, new_filename)

            shutil.copy(full_path, new_path)
            # print(f"✅ 已复制并重命名：{filename} → {new_filename}")



# move best match file according to excel dictionary
def batch_similarity_move_by_excel(exam_folder_sim_excel_path: str, exam_folder_ordered_path: str, exam_file_base_name: list):
    """
    对多个试卷目录批量调用 copy_and_rename_by_similarity_excel，
    前提是每个试卷文件夹下存在同名的 "_相似性.xlsx" 文件。
    """
    for name in exam_file_base_name:
        file_path = os.path.join(exam_folder_sim_excel_path, name + "_相似性.xlsx")
        print(f"\n📂 正在处理：{name}")

        if os.path.exists(file_path):
            df_file_match = pd.read_excel(file_path)
            print(df_file_match)

            # 调用你已有的函数（无需重写）
            copy_and_rename_by_similarity_excel(
                input_folder=os.path.join(exam_folder_sim_excel_path, name),
                output_folder=os.path.join(exam_folder_ordered_path, name),
                excel_path=file_path
            )
        else:
            print(f"❌ 文件不存在：{file_path}")



def rename_txt_files_by_student_info(exam_folder_ordered_path: str, class_folder_name: str):
    """
    将指定班级文件夹下所有 .txt 文件重命名为：
    学号_姓名_student_pageN.txt 的格式
    """
    folder = Path(exam_folder_ordered_path) / class_folder_name

    if not folder.exists():
        print(f"❌ 文件夹不存在：{folder}")
        return

    pattern = re.compile(
        r"^.+?-(\d{10})-([\u4e00-\u9fa5·]{1,20})-.*?(student_page\d)\.txt$"
    )
    print(f"❌")
    for file in folder.glob("*.txt"):
        match = pattern.match(file.name)
        if match:
            student_id, name, page = match.groups()
            new_name = f"{student_id}_{name}_{page}.txt"
            file.rename(file.with_name(new_name))
            print(f"✅ 重命名：{file.name} → {new_name}")
        else:
            print(f"❌ 跳过：{file.name}（未匹配）")



# move best match file according to excel dictionary
def batch_txt_rename_by_student_info(exam_folder_ordered_path: str, exam_file_base_name: list):
    """
    对多个试卷目录批量调用 copy_and_rename_by_similarity_excel，
    前提是每个试卷文件夹下存在同名的 "_相似性.xlsx" 文件。
    """
    for name in exam_file_base_name:
        # 调用你已有的函数（无需重写）
        rename_txt_files_by_student_info(
            exam_folder_ordered_path=exam_folder_ordered_path,
            class_folder_name=name
        )