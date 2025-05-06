import os


def generate_errorref_report(base_dir, output_path):
    """
    遍历 base_dir 下每个子文件夹，读取其中的 ErrorRef.txt；
    如果条目数量 >= 3，则将：
        文件夹名
        不严谨引用如下：
        原文件所有条目
    拼成一段，收集到 report_entries 列表，最后写入 output_path。
    """
    report_entries = []

    for name in os.listdir(base_dir):
        folder = os.path.join(base_dir, name)
        if not os.path.isdir(folder):
            continue

        errorref_file = os.path.join(folder, "ErrorRef.txt")
        if not os.path.isfile(errorref_file):
            continue

        # 读取所有行（去掉末尾换行符）
        with open(errorref_file, "r", encoding="utf-8") as f:
            lines = [line.rstrip("\n") for line in f.readlines()]

        # 只有条目数 >= 3 时才生成报告片段
        if len(lines) >= 3:
            entry_lines = []
            entry_lines.append(f"【{name}】")
            entry_lines.append(
                "不严谨引用，包括论文不存在、作者和题目对不上、论文期刊/会议和题目对不上、引文中英混用等，具体如下：")
            entry_lines.extend(lines)
            entry_lines.append("")  # 空行，分隔不同文件夹段落
            report_entries.append("\n".join(entry_lines))

    # 把所有片段写入统一报告
    with open(output_path, "w", encoding="utf-8") as f:
        f.write("\n".join(report_entries))


def a04_05_report_ref_error(base_directory: str):
    # base_directory = r"C:\MyPython\ExamScore_AIClass\ExamFiles"
    new_dir = "\\ErrorRef_Report.txt"
    output_report = f"{base_directory}{new_dir}"
    generate_errorref_report(base_directory, output_report)
    print(f"已生成报告：{output_report}")
