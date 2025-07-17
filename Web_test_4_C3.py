import re
import pandas as pd
from pathlib import Path

def grade_page3_comprehensive(code: str):
    result = []

    # Q1: 修改 action、onsubmit、type
    action_ok = "action=" in code and "simple.html" in code
    onsubmit_ok = "onsubmit=" in code and "check()" in code
    type_text_ok = re.search(r'<input[^>]+type="text"', code) is not None
    type_pwd_ok = re.search(r'<input[^>]+type="password"', code) is not None
    q1_score = sum([action_ok, onsubmit_ok, type_text_ok, type_pwd_ok])
    q1_comment = []
    if not action_ok:
        q1_comment.append("未正确设置表单的 action=simple.html")
    if not onsubmit_ok:
        q1_comment.append("未添加 onsubmit='check()'")
    if not type_text_ok:
        q1_comment.append("用户名输入框类型未设置为 text")
    if not type_pwd_ok:
        q1_comment.append("密码输入框类型未设置为 password")
    if q1_score >= 3:
        score1 = 6
    else:
        score1 = 3
    result.append((score1, "；".join(q1_comment) or "action、onsubmit 和类型设置均正确"))

    # Q2: CSS 样式
    css = re.findall(r"<style[^>]*>(.*?)</style>", code, re.DOTALL)
    css_code = css[0] if css else ""
    font_ok = "font-size: 20px" in css_code
    margin_ok = "margin" in css_code and ("10px" in css_code or "20px" in css_code)
    if font_ok and margin_ok:
        score2 = 6
        comment2 = "已设置字体大小和按钮外边距，样式完整。"
    elif font_ok or margin_ok:
        score2 = 4
        comment2 = "只设置了字体大小或按钮外边距，样式不完整。"
    else:
        score2 = 3
        comment2 = "未正确设置样式。"
    result.append((score2, comment2))

    # Q3: check 函数逻辑判断
    js = re.findall(r"<script[^>]*>(.*?)</script>", code, re.DOTALL)
    js_code = js[0] if js else ""
    has_func = "function check()" in js_code
    has_account_pwd = "abc" in js_code and "123" in js_code
    has_alert = "alert" in js_code
    has_return = "return false" in js_code or "return true" in js_code
    js_score = sum([has_func, has_account_pwd, has_alert, has_return])
    if js_score >= 3:
        score3 = 6
        comment3 = "check() 函数逻辑正确，能判断密码、弹窗提醒并返回跳转状态。"
    elif js_score >= 1:
        score3 = 4
        comment3 = "check() 函数部分实现，例如有弹窗或变量判断，但逻辑不完整。"
    else:
        score3 = 3
        comment3 = "代码非空但未实现有效的判断逻辑。"
    result.append((score3, comment3))

    # Q4: 密码字段是否使用 password 类型
    pwd_field = re.search(r'<input[^>]+type="password"', code)
    score4 = 6 if pwd_field else 3
    comment4 = "密码框类型设置为 password，输入时会显示 *" if pwd_field else "密码框未设置为 password，用户输入不安全。"
    result.append((score4, comment4))

    # Q5: simple.html 的背景色切换函数
    fun1_ok = "function fun1()" in code and "background" in code
    fun2_ok = "function fun2()" in code and "background" in code
    fun_count = sum([fun1_ok, fun2_ok])
    if fun_count == 2:
        score5 = 6
        comment5 = "fun1() 与 fun2() 实现背景色切换逻辑，功能完整。"
    elif fun_count == 1:
        score5 = 4
        comment5 = "仅写出一个背景切换函数，功能部分实现。"
    else:
        score5 = 3
        comment5 = "代码非空但未定义背景切换函数。"
    result.append((score5, comment5))

    return {
        "文件名": "",
        "题1得分": result[0][0], "题1评语": result[0][1],
        "题2得分": result[1][0], "题2评语": result[1][1],
        "题3得分": result[2][0], "题3评语": result[2][1],
        "题4得分": result[3][0], "题4评语": result[3][1],
        "题5得分": result[4][0], "题5评语": result[4][1],
        "总分30": sum(r[0] for r in result)
    }

def batch_grade_page3_comprehensive_to_excel(folder: str, output_excel: str):
    folder_path = Path(folder)
    records = []

    for file in folder_path.rglob("*.txt"):
        try:
            code = file.read_text(encoding="utf-8")
        except Exception as e:
            print(f"⚠️ 跳过无法读取的文件：{file.name}，原因：{e}")
            continue

        result = grade_page3_comprehensive(code)
        result["文件名"] = file.name
        records.append(result)

    df = pd.DataFrame(records)
    df = df[["文件名", "题1得分", "题1评语",
             "题2得分", "题2评语",
             "题3得分", "题3评语",
             "题4得分", "题4评语",
             "题5得分", "题5评语",
             "总分30"]]
    df.to_excel(output_excel, index=False)
    print(f"✅ 批量评分完成，结果已导出至：{output_excel}")




# 设置输入文件夹路径与输出 Excel 路径
input_folder = r"C:\Users\xijia\Desktop\批改web\S03_test_files_html\精食药临药web1班4-2-2024-2025-2_web期末考试_C卷__word_"  # ← 修改为你的文件夹路径
output_excel = r"C:\Users\xijia\Desktop\批改web\S03_test_files_html\精食药临药web1班4-2-2024-2025-2_web期末考试_C卷__word_评分结果page3.xlsx"                # ← 修改为你的输出路径

# 执行评分与导出
batch_grade_page3_comprehensive_to_excel(input_folder, output_excel)



