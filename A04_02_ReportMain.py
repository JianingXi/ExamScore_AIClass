import os
import pandas as pd

def generate_missing_sections_report(base_dir, output_path):
    sections = {
        "题目": "题目",
        "摘要": "摘要",
        "引言": "引言",
        "结论": "结论",
        "参考文献": "参考文献",
    }

    report_lines = []
    records = []

    for folder_name in sorted(os.listdir(base_dir)):
        folder_path = os.path.join(base_dir, folder_name)
        if not os.path.isdir(folder_path):
            continue

        folder_report = []
        missing_items = []
        missing_count = 0

        # 检查章节文件是否缺失或为空
        for key, display in sections.items():
            filename = f"Section_{key}.txt"
            filepath = os.path.join(folder_path, filename)
            if not os.path.isfile(filepath):
                missing_count += 1
                msg = f"论文缺少{display}（文件不存在）"
                folder_report.append(msg)
                missing_items.append(msg)
            else:
                with open(filepath, 'r', encoding='utf-8') as f:
                    content = f.read().strip()
                if not content:
                    missing_count += 1
                    msg = f"论文缺少{display}"
                    folder_report.append(msg)
                    missing_items.append(msg)

        # 检查 Main_FormatScore.txt
        fs_path = os.path.join(folder_path, "Main_FormatScore.txt")
        format_score, format_comment = "", ""
        format_penalty = 0
        if os.path.isfile(fs_path):
            with open(fs_path, 'r', encoding='utf-8') as f:
                lines = [line.strip() for line in f if line.strip()]
            for line in lines:
                if line.startswith("Score:") or line.startswith("Score:\t"):
                    format_score = line.split(":")[-1].strip()
                else:
                    format_comment += line + " "

            # 简单格式化评语
            format_comment = format_comment.strip()

            # 根据关键词判断格式扣分
            fs_content = "\n".join(lines).strip()
            if "不符合" in fs_content or "格式存在问题" in fs_content or "非双栏" in fs_content:
                format_penalty = 2
            else:
                try:
                    score_val = int(format_score.split("/")[0])
                    if score_val < 8:
                        format_penalty = 1
                except:
                    pass

        # 计算建议评分
        section_penalty = min(4, missing_count * 2)
        suggested_score = max(5, 10 - section_penalty - format_penalty)

        folder_report.append(f"建议评分：{suggested_score}/10")

        # 写入TXT报告
        if folder_report:
            report_lines.append(f"【{folder_name}】")
            report_lines.extend(folder_report)
            report_lines.append("")

        # 记录Excel数据
        records.append({
            "学生文件夹": folder_name,
            "格式得分": format_score,
            "格式评语": format_comment,
            "缺失章节或空章节": "; ".join(missing_items),
            "章节得分": f"{suggested_score}/10"
        })

    # 写入TXT报告
    with open(output_path, 'w', encoding='utf-8') as out_f:
        out_f.write("\n".join(report_lines))
    print(f"✅ 已生成TXT报告：{output_path}")

    # 写入Excel表格
    if records:
        df = pd.DataFrame(records)
        excel_path = output_path.replace(".txt", ".xlsx")
        df.to_excel(excel_path, index=False)
        print(f"✅ 已生成Excel汇总：{excel_path}")

def a04_02_report_main(base_directory: str):
    report_file = os.path.join(base_directory, "MissingSections_Report.txt")
    generate_missing_sections_report(base_directory, report_file)
