import re

def grade_page2_js_single(txt_code: str):
    results = []

    # ---------- 第1题：添加按钮行 + foo()事件绑定 ----------
    if txt_code.strip() == "":
        results.append((0, "未提交任何代码。"))
    else:
        has_button_row = bool(re.search(r'<tr>.*button.*</tr>', txt_code, re.DOTALL))
        has_colspan_2 = "colspan=\"2\"" in txt_code
        has_foo_bind = "onclick=\"foo()" in txt_code or "onclick='foo()" in txt_code
        matched = sum([has_button_row, has_colspan_2, has_foo_bind])
        if matched == 3:
            results.append((8, "添加了按钮行，设置了跨两列单元格，并正确绑定了 foo() 函数。"))
        elif matched == 2:
            results.append((6, "按钮行和结构基本正确，缺少事件绑定或跨列设置。"))
        elif matched == 1:
            results.append((4, "仅写出按钮或结构，缺少绑定事件或跨列设置。"))
        else:
            results.append((4, "代码非空但未正确添加按钮或事件。"))

    # ---------- 第2题：span 样式 ----------
    css_block = re.findall(r"<style[^>]*>(.*?)</style>", txt_code, re.DOTALL)
    css = css_block[0] if css_block else ""
    matched_css = sum([
        "span" in css,
        "border" in css,
        "border-radius" in css,
        "width" in css,
        "height" in css
    ])
    if matched_css >= 4:
        results.append((6, "span 样式设置完整，宽高、边框、圆角均合理。"))
    elif matched_css >= 2:
        results.append((4, f"部分样式设置成功（{matched_css} 项），尚不完整。"))
    elif matched_css >= 1:
        results.append((3, "只设置了少量 span 样式。"))
    else:
        results.append((3, "代码非空但未设置 span 样式。"))

    # ---------- 第3题：foo 函数逻辑 ----------
    script_block = re.findall(r"<script[^>]*>(.*?)</script>", txt_code, re.DOTALL)
    js_code = script_block[0] if script_block else ""
    has_function = "function foo()" in js_code
    has_weight_check = any(w in js_code for w in ["if", ">", "<=", "switch"])
    has_area_fee = any(w in js_code for w in ["省内", "省外", "海外", "dest"])
    has_output = "innerHTML" in js_code or "document.getElementById(" in js_code

    matched_js = sum([has_function, has_weight_check, has_area_fee, has_output])
    if matched_js >= 4:
        results.append((6, "函数逻辑完整，重量判断、地区判断、结果输出均实现。"))
    elif matched_js >= 2:
        results.append((4, "实现了部分逻辑，如地区或重量判断，但输出或结构不全。"))
    elif matched_js >= 1:
        results.append((3, "函数结构或变量定义存在，逻辑实现不完整。"))
    else:
        results.append((3, "代码非空但未实现有效逻辑。"))

    return {
        "题6得分": results[0][0], "题6评语": results[0][1],
        "题7得分": results[1][0], "题7评语": results[1][1],
        "题8得分": results[2][0], "题8评语": results[2][1],
        "总分20": sum(r[0] for r in results)
    }



import pandas as pd
from pathlib import Path

def batch_grade_page2_from_txt_to_excel(folder: str, output_excel: str):
    folder_path = Path(folder)
    records = []

    for file in folder_path.rglob("*.txt"):
        try:
            txt_code = file.read_text(encoding="utf-8")
        except:
            print(f"⚠️ 读取失败：{file}")
            continue

        score_result = grade_page2_js_single(txt_code)
        score_result["文件名"] = file.name
        records.append(score_result)

    df = pd.DataFrame(records)
    df = df[["文件名", "题6得分", "题6评语", "题7得分", "题7评语", "题8得分", "题8评语", "总分20"]]
    df.to_excel(output_excel, index=False)
    print(f"✅ 评分完成，导出至：{output_excel}")



batch_grade_page2_from_txt_to_excel(
    folder=r"C:\Users\xijia\Desktop\批改web\S03_test_files_html\精食药临药web1班4-2-2024-2025-2_web期末考试_C卷__word_",  # 替换为你的目录
    output_excel=r"C:\Users\xijia\Desktop\批改web\S03_test_files_html\精食药临药web1班4-2-2024-2025-2_web期末考试_C卷__word_评分结果page2.xlsx"
)


