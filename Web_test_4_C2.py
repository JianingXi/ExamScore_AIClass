import re
import pandas as pd
from pathlib import Path

def grade_page2_js_single(txt_code: str):
    results = []

    # ---------- 题1：添加按钮表行 + 跨2列 + foo事件绑定 ----------
    if txt_code.strip() == "":
        results.append((0, "未提交任何代码。"))
    else:
        has_button_row = bool(re.search(r'<tr[^>]*>.*?(input|button).*?(input|button).*?</tr>', txt_code, re.DOTALL | re.IGNORECASE))
        has_colspan_2 = re.search(r'colspan\s*=\s*["\']2["\']', txt_code, re.IGNORECASE)
        has_foo_bind = re.search(r'onclick\s*=\s*["\']?foo\s*\(', txt_code)
        has_reset = re.search(r'type=["\']reset["\']', txt_code, re.IGNORECASE)
        matched = sum(bool(x) for x in [has_button_row, has_colspan_2, has_foo_bind, has_reset])

        score_map = [4, 6, 7, 8, 8]
        score = score_map[matched] if matched > 0 else 4

        fb = []
        fb.append("写出按钮表行" if has_button_row else "未写出按钮表行")
        fb.append("包含跨2列单元格" if has_colspan_2 else "未设置 colspan=2")
        fb.append("计算按钮绑定了 foo 函数" if has_foo_bind else "缺少 foo() 绑定事件")
        fb.append("含 reset 重置按钮" if has_reset else "未添加重置按钮")
        results.append((score, "，".join(fb)))

    # ---------- 题2：span 样式设计 ----------
    css_block = re.findall(r"<style[^>]*>(.*?)</style>", txt_code, re.DOTALL | re.IGNORECASE)
    css = css_block[0] if css_block else ""
    matched_css = sum(kw in css for kw in ["span", "border", "border-radius", "width", "height"])
    score_map = [3, 4, 5, 6]
    score = score_map[min(matched_css, 3)] if matched_css > 0 else 3
    results.append((score, f"样式匹配关键词 {matched_css}/5：包含{'、'.join([kw for kw in ['span','border','border-radius','width','height'] if kw in css])}"))

    # ---------- 题3：foo函数计算逻辑 ----------
    js_block = re.findall(r"<script[^>]*>(.*?)</script>", txt_code, re.DOTALL | re.IGNORECASE)
    js = js_block[0] if js_block else ""
    has_function = re.search(r'function\s+foo\s*\(', js)
    has_weight_logic = bool(re.search(r'(weight|parseFloat).*?[<>=]', js))
    has_area_logic = any(x in js for x in ["local", "regional", "overseas"])
    has_output = bool(re.search(r'(innerHTML|value).*total', js))
    matched = sum([bool(has_function), has_weight_logic, has_area_logic, has_output])
    score_map = [3, 4, 5, 6]
    score = score_map[min(matched, 3)] if matched > 0 else 3

    logic_fb = []
    logic_fb.append("包含foo函数定义" if has_function else "缺少foo函数定义")
    logic_fb.append("有重量条件判断" if has_weight_logic else "未实现重量判断逻辑")
    logic_fb.append("含地区附加逻辑" if has_area_logic else "未判断地区附加费用")
    logic_fb.append("有输出语句" if has_output else "缺少输出显示代码")
    results.append((score, "，".join(logic_fb)))

    return {
        "题6得分": results[0][0], "题6评语": results[0][1],
        "题7得分": results[1][0], "题7评语": results[1][1],
        "题8得分": results[2][0], "题8评语": results[2][1],
        "总分20": sum(r[0] for r in results)
    }




def batch_grade_page2_from_txt_to_excel(folder: str, output_excel: str):
    folder_path = Path(folder)
    records = []

    for file in folder_path.rglob("*page2.txt"):
        try:
            txt_code = file.read_text(encoding="utf-8")
        except Exception as e:
            print(f"⚠️ 无法读取文件：{file}，错误：{e}")
            continue

        score_result = grade_page2_js_single(txt_code)
        score_result["文件名"] = file.name
        records.append(score_result)

    df = pd.DataFrame(records)
    df = df[["文件名", "题6得分", "题6评语", "题7得分", "题7评语", "题8得分", "题8评语", "总分20"]]
    df.to_excel(output_excel, index=False)
    print(f"✅ 批量评分完成，导出：{output_excel}")
if __name__ == "__main__":
    base_names = [
        "精食药临药web1班4-2-2024-2025-2_web期末考试（C卷）(word)",
        "公管心法Web1班4-1-2024-2025-2_web期末考试（C卷）(word)"
    ]

    subfolders = [
        "txt_files_1_ordered",
        "txt_files_2_html"
    ]
    txt_base_folder = r"C:\Users\xijia\Desktop\批改web\S02_TXT"
    for name in base_names:
        for sub in subfolders:
            fold_1 = fr"{txt_base_folder}\{sub}\{name}"
            file_excel = fr"{fold_1}_评分结果page2.xlsx"
            batch_grade_page2_from_txt_to_excel(fold_1, file_excel)
