import os
import re
from docx import Document

CH_NUM = r'(?:[一二三四五六七八九十]+|\d+)'
SECTION_ORDER = ['题目', '摘要', '引言', '结论', '参考文献']

def evaluate_docx(path):
    """解析文档，返回 (sections: dict, two_column: bool, total_char_count: int)。"""
    try:
        doc = Document(path)
    except Exception:
        return {}, False, 0

    sections = {sec: [] for sec in SECTION_ORDER}
    current = '题目'
    title_assigned = False
    total_text = []

    def is_double_column(doc):
        ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
        for sec in doc.sections:
            sect_pr = sec._sectPr
            if sect_pr is None:
                continue
            cols = sect_pr.findall('.//w:cols', ns)
            for col in cols:
                num = col.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}num')
                if num and int(num) >= 2:
                    return True
        return False

    two_col = is_double_column(doc)

    for p in doc.paragraphs:
        text = p.text.strip()
        if not text:
            continue

        total_text.append(text)  # 统计全文字数

        if not title_assigned:
            sections['题目'].append(text)
            title_assigned = True
            continue

        if (re.match(rf'^{CH_NUM}\s*[、\.]\s*摘\s*要', text) or
            re.match(rf'^{CH_NUM}\.\s*摘\s*要', text) or
            re.match(r'^\s*摘\s*要', text)):
            current = '摘要'
            sections[current].append(text)
            continue

        if (re.match(rf'^{CH_NUM}\s*[、\.]?\s*引言', text) or
            re.match(r'^\s*引言', text)):
            current = '引言'
            sections[current].append(text)
            continue

        if (re.match(rf'^{CH_NUM}\s*[、\.]?\s*结论', text) or
            re.match(r'^\s*结论', text)):
            current = '结论'
            sections[current].append(text)
            continue

        if (re.match(rf'^{CH_NUM}\s*[、\.]?\s*参\s*考\s*文\s*献', text) or
            re.match(r'^\s*参\s*考\s*文\s*献', text)):
            current = '参考文献'
            sections[current].append(text)
            continue

        sections[current].append(text)

    char_count = sum(len(t) for t in total_text)
    return sections, two_col, char_count

def generate_report(base_dir):
    summary = ['文件夹\t得分\t评语']
    missing = []
    for folder in sorted(os.listdir(base_dir)):
        sub = os.path.join(base_dir, folder)
        if not os.path.isdir(sub):
            continue
        docs = [
            f for f in os.listdir(sub)
            if f.lower().endswith('.docx') and not f.lower().endswith('_summary.docx')
        ]
        # 额外剔除结尾为_summary.docx的文件
        if not docs:
            summary.append(f"{folder}\t0/10\t未检测到.docx 文件。")
            missing.append(folder)
            continue

        path = os.path.join(sub, docs[0])
        sections, is_two_column, char_count = evaluate_docx(path)

        for sec in SECTION_ORDER:
            with open(os.path.join(sub, f"Section_{sec}.txt"), 'w', encoding='utf-8') as f:
                f.write('\n'.join(sections.get(sec, [])))

        score = 10
        comments = []

        # 双栏评分
        if is_two_column:
            comments.append("采用双栏排版")
        else:
            score -= 3
            comments.append("未采用双栏排版（扣3分）")

        # 各核心章节
        for sec in ['摘要', '引言', '结论', '参考文献']:
            content = sections.get(sec, [])
            if not content or len("".join(content).strip()) < 10:
                score -= 1
                comments.append(f"{sec}缺失或内容过少（扣1分）")

        # 标题检测
        title = sections.get('题目', [])
        if not title or len(title[0]) > 100:
            score -= 1
            comments.append("标题格式异常或缺失（扣1分）")

        # 结构检测（段落数）
        total_paragraphs = sum(len(v) for v in sections.values())
        if total_paragraphs < 10:
            score -= 2
            comments.append("段落数过少，结构可能不完整（扣2分）")

        # 🔥 字数检测
        if char_count < 2000:
            score -= 2
            comments.append(f"总字数为 {char_count}，低于2000字（扣2分）")
        elif char_count < 2500:
            score -= 1
            comments.append(f"总字数为 {char_count}，低于2500字（扣1分）")
        elif char_count < 3000:
            comments.append(f"总字数为 {char_count}，略少于3000字（不扣分，提醒）")
        else:
            comments.append(f"总字数为 {char_count}，满足字数要求")

        score = max(score, 0)
        final_comment = "；".join(comments)
        summary.append(f"{folder}\t{score}/10\t{final_comment}")

        # 写入 FormatScore.txt
        with open(os.path.join(sub, 'Main_FormatScore.txt'), 'w', encoding='utf-8') as cf:
            cf.write(f"Score:\t{score}/10\n{final_comment}")

    # 输出缺失列表与汇总表
    with open(os.path.join(base_dir, 'MissingDocxFolders.txt'), 'w', encoding='utf-8') as f:
        f.write('\n'.join(missing))
    with open(os.path.join(base_dir, 'FormatScore_Summary.txt'), 'w', encoding='utf-8') as f:
        f.write('\n'.join(summary))

def a04_01_format_scorer(BASE_DIR: str):
    generate_report(BASE_DIR)
