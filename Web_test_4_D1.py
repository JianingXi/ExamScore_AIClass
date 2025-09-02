import re
import pandas as pd
from pathlib import Path

def grade_page1_from_comments(code: str):
    result = []
    css_code = "\n".join(re.findall(r"<style.*?>(.*?)</style>", code, re.DOTALL))
    html_code = re.sub(r"<style.*?>.*?</style>", "", code, flags=re.DOTALL)

    def non_blank(txt): return bool(txt.strip())

    # 题1：结构嵌套
    tags = ["class=\"header\"", "class=\"content\"", "class=\"box\"", "class=\"footer\""]
    count = sum(tag in code for tag in tags)
    if count == 4:
        score1, comment1 = 6, "结构完整，已正确设置 header、content、box 和 footer 四个区域。"
    elif count >= 2:
        score1, comment1 = 4, "结构设置部分完成，部分区域缺失或层级不全。"
    elif count >= 1:
        score1, comment1 = 3, "只写了少部分结构区域，整体层次尚不清晰。"
    elif non_blank(code):
        score1, comment1 = 3, "有部分代码，但结构嵌套和命名不符合要求。"
    else:
        score1, comment1 = 0, "没有看到结构层设置的相关内容。"

    # 题2：插入指定图文内容
    count = sum([
        "lsqs.jpg" in code,
        "绿水青山就是金山银山" in code,
        "shengtai.jpg" in code and ("生态优先" in code or "绿色发展" in code),
        "ziran.jpg" in code and ("和谐共生" in code or "人与自然" in code),
        any(x in code for x in ["© 2025", "版权所有", "推动绿色发展"])
    ])
    if count == 5:
        score2, comment2 = 6, "图文内容齐全，已覆盖题目中所有指定元素。"
    elif count >= 3:
        score2, comment2 = 4, "部分图文内容已填写，建议补充遗漏图片或文字。"
    elif count >= 1:
        score2, comment2 = 3, "图文内容填写较少，整体不完整。"
    elif non_blank(code):
        score2, comment2 = 3, "页面存在内容，但未体现指定图文要求。"
    else:
        score2, comment2 = 0, "缺少所有图片和文字内容。"

    # 题3：header/footer 样式
    h_ok = all(x in css_code for x in ["background-color", "text-align", "padding: 24px"])
    f_ok = all(x in css_code for x in ["#2d4a3a", "#e6f2e6", "font-size: 14px"])
    total = sum([h_ok, f_ok])
    if total == 2:
        score3, comment3 = 6, "header 和 footer 的样式设置准确全面，符合题目要求。"
    elif total == 1:
        score3, comment3 = 4, "仅设置了 header 或 footer 的样式，建议补充完整。"
    elif any(x in css_code for x in ["header", "footer"]):
        score3, comment3 = 3, "样式存在但不符合具体要求，颜色、字号或对齐方式缺失。"
    elif non_blank(css_code):
        score3, comment3 = 3, "样式代码存在，但未针对 header/footer 设置具体样式。"
    else:
        score3, comment3 = 0, "未设置 header/footer 的相关样式。"

    # 题4：图片与标题样式
    hit_items = [
        "header img", "max-width", "border-radius: 10px", "box-shadow",
        "font-size: 20px", "color: #d6f5d6", "line-height"
    ]
    count = sum(x in css_code for x in hit_items)
    if count >= 6:
        score4, comment4 = 6, "图片与标题的样式设置全面，满足视觉效果要求。"
    elif count >= 4:
        score4, comment4 = 4, "样式设置较好，但仍有少量细节未完成，如圆角或颜色。"
    elif count >= 2:
        score4, comment4 = 3, "样式设置不够，缺少关键属性如阴影、字号等。"
    elif non_blank(css_code):
        score4, comment4 = 3, "样式部分存在，但未体现图片和标题要求的设置。"
    else:
        score4, comment4 = 0, "未设置与图片或标题相关的样式。"

    # 题5：box 样式
    hit_items = [
        ".box", "width: 45%", "background", "border-radius", "padding",
        "float", ".box img", "width: 100%", "margin-bottom", "border-radius: 8px"
    ]
    count = sum(x in css_code for x in hit_items)
    if count >= 8:
        score5, comment5 = 6, "box 的样式设置较完整，表现良好。"
    elif count >= 5:
        score5, comment5 = 4, "box 样式部分完成，仍需补充如图片圆角、内边距等细节。"
    elif count >= 2:
        score5, comment5 = 3, "box 样式设置较少，建议参考题干补足核心属性。"
    elif non_blank(css_code):
        score5, comment5 = 3, "样式代码存在，但缺少 box 相关定义。"
    else:
        score5, comment5 = 0, "未完成 box 样式设置。"

    return {
        "文件名": "",
        "题1得分": score1, "题1评语": comment1,
        "题2得分": score2, "题2评语": comment2,
        "题3得分": score3, "题3评语": comment3,
        "题4得分": score4, "题4评语": comment4,
        "题5得分": score5, "题5评语": comment5,
        "总分30": sum([score1, score2, score3, score4, score5])
    }

def batch_grade_page1_to_excel(folder: str, output_excel: str):
    folder_path = Path(folder)
    records = []
    for file in folder_path.rglob("*page1.txt"):
        try:
            code = file.read_text(encoding="utf-8")
        except Exception as e:
            print(f"❌ 无法读取 {file.name}，错误：{e}")
            continue
        result = grade_page1_from_comments(code)
        result["文件名"] = file.name
        records.append(result)

    df = pd.DataFrame(records)
    df = df[[  # 输出顺序固定
        "文件名",
        "题1得分", "题1评语",
        "题2得分", "题2评语",
        "题3得分", "题3评语",
        "题4得分", "题4评语",
        "题5得分", "题5评语",
        "总分30"
    ]]
    df.to_excel(output_excel, index=False)
    print(f"✅ 批改完成，结果已保存至：{output_excel}")

# 主程序使用方式
if __name__ == "__main__":
    base_names = [
        "护理检验web1班4-3-2024-2025-2_Web期末考试（护理、检验）(word)",
        "口麻影Web1班4-4-2024-2025-2_Web期末考试（影像、麻醉、口腔）(word)"
    ]
    subdirs = ["txt_files_1_ordered", "txt_files_2_html"]
    base_folder = r"C:\Users\xijia\Desktop\批改web\S02_TXT"

    for name in base_names:
        for sub in subdirs:
            folder = fr"{base_folder}\{sub}\{name}"
            out_file = fr"{folder}_评分结果page1.xlsx"
            batch_grade_page1_to_excel(folder, out_file)
