import re
import pandas as pd
from pathlib import Path

def grade_page3_comprehensive(code: str):
    result = []

    # ---------- Q1: 修改 action、onsubmit、input type ----------
    action_ok = re.search(r'<form[^>]+action=["\']simple\.html["\']', code)
    onsubmit_ok = re.search(r'onsubmit=["\']check\(\)', code)
    type_text_ok = re.search(r'<input[^>]+type=["\']text["\']', code)
    type_pwd_ok = re.search(r'<input[^>]+type=["\']password["\']', code)

    matched_q1 = sum(bool(x) for x in [action_ok, onsubmit_ok, type_text_ok, type_pwd_ok])
    score_q1 = 6 if matched_q1 >= 3 else 3
    fb_q1 = []
    fb_q1.append("action 设置正确" if action_ok else "未正确设置 form 的 action 为 simple.html")
    fb_q1.append("onsubmit 使用 check()" if onsubmit_ok else "未设置 onsubmit=check()")
    fb_q1.append("账号输入框类型为 text" if type_text_ok else "账号输入框未设置为 type='text'")
    fb_q1.append("密码输入框类型为 password" if type_pwd_ok else "密码输入框未设置为 type='password'")
    result.append((score_q1, "；".join(fb_q1)))

    # ---------- Q2: CSS 样式 ----------
    css = re.findall(r"<style[^>]*>(.*?)</style>", code, re.DOTALL | re.IGNORECASE)
    css_code = css[0] if css else ""
    has_font_size = re.search(r'font-size\s*:\s*20px', css_code)
    has_margin = re.search(r'margin\s*:\s*10px', css_code) or re.search(r'margin-right\s*:\s*20px', css_code)
    matched_css = sum(bool(x) for x in [has_font_size, has_margin])
    score_q2 = 6 if matched_css == 2 else (4 if matched_css == 1 else 3)
    fb_q2 = []
    fb_q2.append("字体大小设置为20px" if has_font_size else "未设置字体大小")
    fb_q2.append("按钮设置了外边距" if has_margin else "未设置按钮外边距")
    result.append((score_q2, "；".join(fb_q2)))

    # ---------- Q3: check 函数逻辑判断 ----------
    js = re.findall(r"<script[^>]*>(.*?)</script>", code, re.DOTALL | re.IGNORECASE)
    js_code = js[0] if js else ""
    has_check = "function check()" in js_code
    has_correct_logic = "abc" in js_code and "123" in js_code
    has_alert = "alert" in js_code
    has_return = "return" in js_code

    matched_js = sum([has_check, has_correct_logic, has_alert, has_return])
    score_q3 = 6 if matched_js >= 3 else (4 if matched_js >= 1 else 3)
    fb_q3 = []
    fb_q3.append("定义了 check 函数" if has_check else "未定义 check 函数")
    fb_q3.append("含账号/密码判断逻辑" if has_correct_logic else "未判断账号或密码是否为 abc / 123")
    fb_q3.append("有弹窗提示" if has_alert else "未设置 alert 提示")
    fb_q3.append("有返回语句" if has_return else "未设置 return true / false")
    result.append((score_q3, "；".join(fb_q3)))

    # ---------- Q4: 密码字段是否为 password ----------
    score_q4 = 6 if type_pwd_ok else 3
    comment_q4 = "密码输入框已设置为 type='password'" if type_pwd_ok else "密码框未设置为 password，不能显示 *"
    result.append((score_q4, comment_q4))

    # ---------- Q5: simple.html 的背景切换函数 ----------
    fun1_ok = "function fun1()" in code and "background" in code
    fun2_ok = "function fun2()" in code and "background" in code
    matched_fun = sum([fun1_ok, fun2_ok])
    score_q5 = 6 if matched_fun == 2 else (4 if matched_fun == 1 else 3)
    fb_q5 = []
    fb_q5.append("fun1() 定义正确" if fun1_ok else "缺少 fun1() 或未设置背景色")
    fb_q5.append("fun2() 定义正确" if fun2_ok else "缺少 fun2() 或未设置背景色")
    result.append((score_q5, "；".join(fb_q5)))

    return {
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

    expected_keys = [
        "文件名",
        "题1得分", "题1评语",
        "题2得分", "题2评语",
        "题3得分", "题3评语",
        "题4得分", "题4评语",
        "题5得分", "题5评语",
        "总分30"
    ]

    for file in folder_path.rglob("*page3.txt"):  # ✅ 匹配所有第三页代码文件
        try:
            code = file.read_text(encoding="utf-8")
        except Exception as e:
            print(f"⚠️ 跳过无法读取的文件：{file.name}，原因：{e}")
            continue

        result = grade_page3_comprehensive(code)
        result["文件名"] = file.name

        # 显式补全所有字段，避免KeyError
        for key in expected_keys:
            result.setdefault(key, "")

        records.append(result)

    if not records:
        print(f"❌ 没有成功评分的记录，无法导出 Excel。目录：{folder}")
        return

    df = pd.DataFrame(records)

    for col in expected_keys:
        if col not in df.columns:
            df[col] = ""

    df = df[expected_keys]
    df.to_excel(output_excel, index=False)
    print(f"✅ 批量评分完成，已导出至：{output_excel}")



if __name__ == "__main__":
    exam_file_base_name = [
        "精食药临药web1班4-2-2024-2025-2_web期末考试（C卷）(word)",
        "公管心法Web1班4-1-2024-2025-2_web期末考试（C卷）(word)",
    ]
    html_or_ordered = [
        "txt_files_1_ordered",
        "txt_files_2_html"
    ]
    txt_base_folder = r"C:\Users\xijia\Desktop\批改web\S02_TXT"
    for name in exam_file_base_name:
        for subdir in html_or_ordered:
            folder = fr"{txt_base_folder}\{subdir}\{name}"
            output_excel = fr"{folder}_评分结果page3.xlsx"
            batch_grade_page3_comprehensive_to_excel(folder, output_excel)

