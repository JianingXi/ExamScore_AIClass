import pandas as pd
from pathlib import Path
import re

# ================== 各题评分函数 ==================

def score_question_6_table_structure(code: str):
    if code.strip() == "":
        return 0, "完全未写表格结构。"
    row_count = code.count("<tr>")
    colspan_count = code.count('colspan="2"')
    if row_count >= 4 and colspan_count >= 2:
        return 4, "表格结构完整，跨列设置正确。"
    elif row_count >= 2 and colspan_count >= 1:
        return 3, "表格结构部分实现，行或合并有遗漏。"
    else:
        return 2, "表格结构明显不完整，仅给起评分。"

def score_question_7_table_content(code: str):
    if code.strip() == "":
        return 0, "完全未写表格内容。"
    matched = 0
    if "<input" in code and "type=\"text\"" in code:
        matched += 1
    if "<input" in code and "type=\"button\"" in code:
        matched += 1
    if any(k in code for k in ["小区停车", "时长", "收费"]):
        matched += 1
    if matched == 3:
        return 4, "表格内容完整，文字与控件正确。"
    elif matched == 2:
        return 3, "内容部分正确，有轻微遗漏。"
    else:
        return 2, "仅写了部分元素，起评分。"

def score_question_8_button_event(code: str):
    if code.strip() == "":
        return 0, "完全未写按钮事件绑定。"
    if 'onclick="parking()"' in code or "onclick='parking()'" in code:
        return 4, "按钮事件绑定正确。"
    elif "onclick" in code and "parking" in code:
        return 3, "事件绑定基本正确，语法略有瑕疵。"
    elif "<input" in code and "button" in code:
        return 2, "仅写了按钮，未正确绑定事件。"
    else:
        return 2, "按钮绑定逻辑不完整，仅起评分。"

def score_question_9_function_defined(code: str):
    if code.strip() == "":
        return 0, "完全未定义 parking() 函数。"
    if "function parking()" in code and ("=" not in code.split("function parking()")[1]):
        return 4, "函数定义完整，结构正确。"
    elif "function parking()" in code:
        return 3, "函数大致写出，缺部分逻辑。"
    elif "parking" in code:
        return 2, "有函数名字但结构不完整，仅起评分。"
    else:
        return 2, "仅写了函数名或空函数体。"

def score_question_10_parking_logic(code: str):
    if code.strip() == "":
        return 0, "未实现任何计费逻辑。"
    logic_ok = all([
        "Math" in code or "//" in code or "%" in code,
        "15" in code,
        "24" in code,
        "if" in code
    ])
    if logic_ok:
        return 4, "收费逻辑基本正确，含免费时长、分段收费和封顶判断。"
    elif "15" in code and "24" in code:
        return 3, "收费思路存在但逻辑不完整。"
    elif "15" in code or "24" in code:
        return 2, "部分常量出现但未成系统逻辑，仅起评分。"
    else:
        return 2, "仅写函数框架，逻辑内容缺失。"

# ================== 主函数 ==================

def grade_html_js_questions(folder: str, output_excel: str):
    base_path = Path(folder)
    records = []

    for txt_file in base_path.rglob("*.txt"):
        try:
            code = txt_file.read_text(encoding="utf-8")
        except:
            print(f"⚠️ 无法读取：{txt_file}")
            continue

        s6, c6 = score_question_6_table_structure(code)
        s7, c7 = score_question_7_table_content(code)
        s8, c8 = score_question_8_button_event(code)
        s9, c9 = score_question_9_function_defined(code)
        s10, c10 = score_question_10_parking_logic(code)
        total = s6 + s7 + s8 + s9 + s10

        records.append({
            "文件名": txt_file.name,
            "题6得分": s6,
            "题6评语": c6,
            "题7得分": s7,
            "题7评语": c7,
            "题8得分": s8,
            "题8评语": c8,
            "题9得分": s9,
            "题9评语": c9,
            "题10得分": s10,
            "题10评语": c10,
            "总分": total,
        })

    df = pd.DataFrame(records)
    df.to_excel(output_excel, index=False)
    print(f"✅ 评分完成，共处理 {len(df)} 份试卷。结果保存于：{output_excel}")


grade_html_js_questions(
    folder=r"C:\Users\xijia\Desktop\批改web\S03_test_files_ordered\儿生预web1班2-1-2024-2025-2_Web_A卷_-2_word_",
    output_excel=r"C:\Users\xijia\Desktop\批改web\S03_test_files_ordered\儿生预web1班2-1-2024-2025-2_Web_A卷_-2_word_评分结果page2.xlsx"
)

