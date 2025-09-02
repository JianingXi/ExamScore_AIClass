import re

def grade_page2_checkbox_js(code: str):
    result = []

    # Q6: 检查是否写了6个正确的复选框
    checkbox_matches = re.findall(r'<input\s+[^>]*type=["\']checkbox["\'][^>]*name=["\']book["\'][^>]*value=', code)
    score6 = 4 if len(checkbox_matches) >= 6 else 2
    comment6 = f"找到 {len(checkbox_matches)} 个复选框。" if len(checkbox_matches) >= 6 else f"复选框数量不足：{len(checkbox_matches)}"

    # Q7: 获取复选框列表代码
    found7 = re.search(r"document\.getElementsByName\(['\"]book['\"]\)", code)
    score7 = 4 if found7 else 2
    comment7 = "使用 getElementsByName 正确。" if found7 else "未正确使用 getElementsByName('book')。"

    # Q8: 获取结果区域代码
    found8 = re.search(r"document\.getElementById\(['\"]selection-result['\"]\)", code)
    score8 = 4 if found8 else 2
    comment8 = "使用 getElementById 正确。" if found8 else "未正确使用 getElementById('selection-result')。"

    # Q9: 判断是否选中 checked
    found9 = re.search(r"\[\s*i\s*\]\.checked", code)
    score9 = 4 if found9 else 2
    comment9 = "成功判断 checked 属性。" if found9 else "未检测 checked 属性。"

    # Q10: 显示 join('、') 的代码
    found10 = "join('、')" in code or 'join("、")' in code
    score10 = 4 if found10 else 2
    comment10 = "结果正确使用顿号分隔。" if found10 else "未正确使用 join('、') 显示结果。"

    return {
        "文件名": "",
        "题6得分": score6, "题6评语": comment6,
        "题7得分": score7, "题7评语": comment7,
        "题8得分": score8, "题8评语": comment8,
        "题9得分": score9, "题9评语": comment9,
        "题10得分": score10, "题10评语": comment10,
        "总分20": score6 + score7 + score8 + score9 + score10
    }

import pandas as pd
from pathlib import Path

def batch_grade_page2_checkbox_to_excel(folder: str, output_excel: str):
    folder_path = Path(folder)
    records = []

    for file in folder_path.rglob("*page2.txt"):
        try:
            code = file.read_text(encoding="utf-8")
        except Exception as e:
            print(f"⚠️ 跳过无法读取的文件：{file.name}，错误：{e}")
            continue

        result = grade_page2_checkbox_js(code)
        result["文件名"] = file.name
        records.append(result)

    df = pd.DataFrame(records)
    df = df[[
        "文件名",
        "题6得分", "题6评语",
        "题7得分", "题7评语",
        "题8得分", "题8评语",
        "题9得分", "题9评语",
        "题10得分", "题10评语",
        "总分20"
    ]]
    df.to_excel(output_excel, index=False)
    print(f"✅ 批改完成，结果已保存到：{output_excel}")



if __name__ == "__main__":
    exam_file_base_name = [
        "护理检验web1班4-3-2024-2025-2_Web期末考试（护理、检验）(word)",
        "口麻影Web1班4-4-2024-2025-2_Web期末考试（影像、麻醉、口腔）(word)",
    ]
    html_or_ordered = [
        "txt_files_1_ordered",
        "txt_files_2_html"
    ]
    txt_base_folder = r"C:\Users\xijia\Desktop\批改web\S02_TXT"
    for name in exam_file_base_name:
        for subdir in html_or_ordered:
            folder = fr"{txt_base_folder}\{subdir}\{name}"
            output_excel = fr"{folder}_评分结果page2.xlsx"
            batch_grade_page2_checkbox_to_excel(folder, output_excel)
