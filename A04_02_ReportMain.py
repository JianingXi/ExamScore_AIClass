import os

def generate_missing_sections_report(base_dir, output_path):
    """
    遍历 base_dir 下的所有子文件夹，检查每个子文件夹中的以下文件：
      - Section_题目.txt
      - Section_摘要.txt
      - Section_引言.txt
      - Section_结论.txt
      - Section_参考文献.txt
      - Main_FormatScore.txt

    评分逻辑：
    - 每缺失或空的 section 扣 2 分，最多扣 4 分；
    - 格式不符合规范（如缺双栏）再酌情扣 1～2 分；
    - 最低建议评分为 5 分；
    - 若 Main_FormatScore 为满分且说明齐全，不输出任何格式评语。
    """

    sections = {
        "题目":     "题目",
        "摘要":     "摘要",
        "引言":     "引言",
        "结论":     "结论",
        "参考文献": "参考文献",
    }

    report_lines = []

    for folder_name in os.listdir(base_dir):
        folder_path = os.path.join(base_dir, folder_name)
        if not os.path.isdir(folder_path):
            continue

        folder_report = []
        missing_count = 0  # 缺失计数

        # 检查章节文件是否缺失或为空
        for key, display in sections.items():
            filename = f"Section_{key}.txt"
            filepath = os.path.join(folder_path, filename)

            if not os.path.isfile(filepath):
                missing_count += 1
                folder_report.append(f"论文缺少{display}（文件不存在）")
            else:
                with open(filepath, 'r', encoding='utf-8') as f:
                    content = f.read().strip()
                if not content:
                    missing_count += 1
                    folder_report.append(f"论文缺少{display}")

        # 检查格式评分文件
        fs_path = os.path.join(folder_path, "Main_FormatScore.txt")
        fs_content = ""
        format_penalty = 0
        if os.path.isfile(fs_path):
            with open(fs_path, 'r', encoding='utf-8') as f:
                lines = [line.strip() for line in f if line.strip() and "Score: 0/10" not in line]

            fs_content = "\n".join(lines).strip()
            perfect_msg = "Score: 10/10"
            ok_format_msg = "文档为双栏且所有核心环节齐全，格式符合要求。"

            # 判断是否需要输出格式评语
            if fs_content and not (perfect_msg in fs_content and ok_format_msg in fs_content):
                folder_report.extend(lines)

                # 粗略评分判断（根据关键词）
                if "不符合" in fs_content or "格式存在问题" in fs_content or "非双栏" in fs_content:
                    format_penalty = 2
                else:
                    for line in lines:
                        if "Score:" in line:
                            try:
                                score_val = int(line.split("Score:")[1].split("/")[0])
                                if score_val < 8:
                                    format_penalty = 1
                            except:
                                pass

        # 计算建议评分
        section_penalty = min(4, missing_count * 2)
        score = max(5, 10 - section_penalty - format_penalty)
        folder_report.append(f"建议评分：{score}/10")

        # 写入总报告
        if folder_report:
            report_lines.append(f"【{folder_name}】")
            report_lines.extend(folder_report)
            report_lines.append("")  # 空行分隔

    # 写入最终报告文件
    with open(output_path, 'w', encoding='utf-8') as out_f:
        out_f.write("\n".join(report_lines))

    print(f"✅ 已生成遗漏章节与评语报告：{output_path}")


def a04_02_report_main(base_directory: str):
    """
    执行入口函数，生成报告。
    """
    report_file = os.path.join(base_directory, "MissingSections_Report.txt")
    generate_missing_sections_report(base_directory, report_file)
