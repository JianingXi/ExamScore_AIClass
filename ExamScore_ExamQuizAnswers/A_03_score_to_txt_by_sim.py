import os
import re
from difflib import SequenceMatcher

def normalize_text(text):
    """去除空格、换行、制表符等，得到纯净文本"""
    return re.sub(r'\s+', '', text)

def compare_with_chinese_feedback(gold_path, student_path, total_score=100):
    # 读取文件内容
    with open(gold_path, 'r', encoding='utf-8') as f1:
        gold_text = f1.read()
    with open(student_path, 'r', encoding='utf-8') as f2:
        student_text = f2.read()

    # 内容归一化
    gold_clean = normalize_text(gold_text)
    student_clean = normalize_text(student_text)

    # 相似度计算（支持局部平移）
    matcher = SequenceMatcher(None, gold_clean, student_clean)
    similarity = matcher.ratio()
    penalty = (1 - similarity) * total_score
    final_score = round(max(total_score - penalty, 0))

    # 基础评语
    feedback = f""

    if similarity > 0.95:
        feedback += "所批改答案与标准答案高度一致，仅有轻微格式或字符差异，表现优异。"
    elif similarity > 0.85:
        feedback += "所批改答案与标准答案大体相符，存在部分术语或表达差异，建议注意专业表述与顺序。"
    elif similarity > 0.70:
        feedback += "所批改答案与标准答案存在明显差异，可能在内容顺序、术语或数据上出现偏差，建议加强理解。"
    else:
        feedback += "所批改答案与标准答案出入较大，建议重新审题并加强对相关知识点的掌握。"


    # 差异内容说明（最多列出3处）
    diff_blocks = matcher.get_opcodes()
    key_diffs = []
    for tag, i1, i2, j1, j2 in diff_blocks:
        if tag != 'equal':
            std_seg = gold_clean[i1:i2]
            stu_seg = student_clean[j1:j2]
            if std_seg and stu_seg:
                key_diffs.append(f"标准为“{std_seg}”，所批改答案为“{stu_seg}”")
            elif std_seg and not stu_seg:
                key_diffs.append(f"标准中出现“{std_seg}”，但所批改答案中缺失")
            elif stu_seg and not std_seg:
                key_diffs.append(f"所批改答案中多出了“{stu_seg}”")
        if len(key_diffs) >= 3:
            break

    if key_diffs:
        feedback += "存在如下具体差异：" + "；".join(key_diffs) + "。"

    return round(final_score, 1), feedback



def multi_blank_sweep(gold_base: str, student_base: str, num_blank: int, blank_score: float):
    # 预设固定长度的空数组
    scores = [None] * num_blank
    feedbacks = [""] * num_blank

    for i in range(1, num_blank + 1):
        gold_path = gold_base % i
        student_path = student_base % i
        idx = i - 1  # 0-based 索引

        if os.path.exists(gold_path) and os.path.exists(student_path):
            score, feedback = compare_with_chinese_feedback(gold_path, student_path, blank_score)

            # 将得分四舍五入为 0.5 的整数倍
            score = round(score * 2) / 2

            scores[idx] = score
            feedbacks[idx] = feedback

            print(f"\n📝 题目 B{i:03d}")
            print("得分：", score)
            print("评语：", feedback)
        else:
            print(f"\n⚠️ 缺失文件：B{i:03d}")
            scores[idx] = None
            feedbacks[idx] = "文件缺失，未评分。"

    return scores, feedbacks



def split_by_chinese_blank_labels(input_path):
    """
    将文件按“第X空：”进行拆分，每一空内容保存为一个新文件，命名以 _blank_XXX 结尾。

    参数:
        input_path (str): 原始txt文件路径
    """
    # 读取原始文本内容
    with open(input_path, 'r', encoding='utf-8') as f:
        text = f.read()

    # 正则匹配所有“第X空：”的位置
    pattern = re.compile(r"(第[一二三四五六七八九十百零〇两\d]+空：)")
    matches = list(pattern.finditer(text))

    if not matches:
        print("⚠️ 未找到任何“第X空：”结构。")
        return

    # 获取路径信息
    folder = os.path.dirname(input_path)
    base_name = os.path.splitext(os.path.basename(input_path))[0]

    # 遍历每段拆分
    for i in range(len(matches)):
        start = matches[i].end()
        end = matches[i + 1].start() if i + 1 < len(matches) else len(text)
        content = text[start:end].strip()

        file_index = f"{i + 1:03d}"  # 三位数编号
        output_filename = f"{base_name}_blank_{file_index}.txt"
        output_path = os.path.join(folder, output_filename)

        with open(output_path, 'w', encoding='utf-8') as f:
            f.write(content)


def batch_split_and_score_students(
        folder: str,
        str_student: str,
        num_q_st: int,
        num_q_en: int,
        num_blank: int,
        blank_score: float
):
    """
    对多个学生编号进行：拆分 + 自动评分

    参数：
        folder      根目录路径
        str_student 文件名前缀（不含 student_id）
        num_q_st    学生编号起始（int），如 13
        num_q_en    学生编号结束（int），如 15
        num_blank   每位学生的题目空数（int）

    返回：
        result_dict: {学号: {"scores": [...], "feedbacks": [...]}}
    """
    result_dict = {}

    for sid in range(num_q_st, num_q_en + 1):
        question_id = f"{sid:03d}"
        print(f"\n📘 开始处理 student_{question_id}")

        # 主文件路径
        filename = f"txt_files/{str_student}-student_{question_id}.txt"
        input_path = os.path.join(folder, filename)

        if os.path.exists(input_path):
            split_by_chinese_blank_labels(input_path)
        else:
            print(f"❌ 文件不存在：{input_path}")
            result_dict[question_id] = {
                "scores": [None] * num_blank,
                "feedbacks": ["未找到学生文件，未评分。"] * num_blank
            }
            continue

        # Gold + student_blank 路径模板
        gold_base = os.path.join(folder, f"GoldStandard/Q{question_id}_B%03d.txt")
        student_base = os.path.join(folder, f"txt_files/{str_student}-student_{question_id}_blank_%03d.txt")

        # 评分
        scores, feedbacks = multi_blank_sweep(gold_base, student_base, num_blank, blank_score)

        # 保存结果
        result_dict[question_id] = {
            "scores": scores,
            "feedbacks": feedbacks
        }

    return result_dict


import pandas as pd

def get_all_docx_filenames(root_folder):
    docx_filenames = []

    for dirpath, _, filenames in os.walk(root_folder):
        for filename in filenames:
            if filename.lower().endswith('.docx'):
                name_without_ext = os.path.splitext(filename)[0]
                docx_filenames.append(name_without_ext)  # 只保留无后缀文件名

    return docx_filenames





# 主批量处理函数
def batch_export_student_scores_to_excel(folder, str_student_list, num_q_st, num_q_en, num_blank, blank_score: float):
    all_rows = []

    for str_student in str_student_list:
        print(f"\n📂 处理学生：{str_student}")
        row = {"学生编号": str_student}

        for sid in range(num_q_st, num_q_en + 1):
            question_id = f"{sid:03d}"

            gold_base = os.path.join(folder, f"GoldStandard/Q{question_id}_B%03d.txt")
            student_base = os.path.join(folder, f"txt_files/{str_student}-student_{question_id}_blank_%03d.txt")

            scores, feedbacks = multi_blank_sweep(gold_base, student_base, num_blank, blank_score)

            for i in range(num_blank):
                col_score = f"题目{question_id}_B{i+1:03d}_得分"
                col_feedback = f"题目{question_id}_B{i+1:03d}_评语"
                row[col_score] = scores[i]
                row[col_feedback] = feedbacks[i]

        all_rows.append(row)

    df = pd.DataFrame(all_rows)
    output_path = os.path.join(folder, "学生评分结果.xlsx")
    df.to_excel(output_path, index=False)
    print(f"\n✅ 所有评分结果已写入 Excel：{output_path}")


# 示例调用
folder = r"C:\MyDocument\ToDoList\D20_DoingPlatform\D20_人工智能与大数据\新建文件夹\23临床药学-2024-2025春季期末考试A卷(word)"
num_q_st = 13
num_q_en = 15
num_blank = 5
blank_score = 3.0
str_student_list = get_all_docx_filenames(folder)

batch_export_student_scores_to_excel(
    folder=folder,
    str_student_list=str_student_list,
    num_q_st=num_q_st,
    num_q_en=num_q_en,
    num_blank=num_blank,
    blank_score=blank_score
)