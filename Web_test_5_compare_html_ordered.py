from ExamScore_ExamQuizAnswers.A_03_Web03_compare_and_merge import (overwrite_first_col_with_id_name,
                                                                    overwrite_excels_by_union,
                                                                    compare_excel_rows,
                                                                    rearrange_columns,
                                                                    merge_excels_by_first_column)


import shutil
from pathlib import Path

def reset_and_merge_folders(final_dir_path: str, source1_path: str, source2_path: str):
    final_dir = Path(final_dir_path)
    source_dirs = [Path(source1_path), Path(source2_path)]

    # Step 1：删除 Final 目录下所有内容
    if final_dir.exists():
        for item in final_dir.iterdir():
            if item.is_dir():
                shutil.rmtree(item)
            else:
                item.unlink()
        print(f"✅ 已清空 Final 目录：{final_dir}")
    else:
        final_dir.mkdir(parents=True)
        print(f"📁 已创建 Final 目录：{final_dir}")

    # Step 2：复制整个 source 文件夹到 Final（包含文件夹名本体）
    for src in source_dirs:
        if not src.exists():
            print(f"⚠️ 源目录不存在：{src}")
            continue

        dest = final_dir / src.name
        if dest.exists():
            shutil.rmtree(dest)
        shutil.copytree(src, dest)
        print(f"📋 已复制目录 {src.name} 到 {final_dir}")

    print("🎉 所有操作完成")


if __name__ == "__main__":
    Final_dir = r"C:\Users\xijia\Desktop\批改web\S02_TXT\Final"

    # ✅ 示例调用（你可以直接放到主程序中使用）
    reset_and_merge_folders(
        Final_dir,
        r"C:\Users\xijia\Desktop\批改web\S02_TXT\txt_files_2_html",
        r"C:\Users\xijia\Desktop\批改web\S02_TXT\txt_files_1_ordered",
    )

    exam_file_base_name = \
        ["儿生预web1班2-1-2024-2025-2_Web（A卷）-2(word)",
         "临I南山Web2班2-2-2024-2025-2_Web（A卷）-2(word)",
         "临2六中web2班2-3-2024-2025-2_Web期末考试（B卷）(word)",
         "公管心法Web1班4-1-2024-2025-2_web期末考试（C卷）(word)",
         "精食药临药web1班4-2-2024-2025-2_web期末考试（C卷）(word)",
         "护理检验web1班4-3-2024-2025-2_Web期末考试（护理、检验）(word)",
         "口麻影Web1班4-4-2024-2025-2_Web期末考试（影像、麻醉、口腔）(word)",
         ]
    
    html_or_ordered = [
        "txt_files_1_ordered",
        "txt_files_2_html"
    ]
    page_vec = [
        "page1",
        "page2",
        "page3"
    ]

    overwrite_first_col_with_id_name(Final_dir + "\\" + "txt_files_1_ordered")
    overwrite_first_col_with_id_name(Final_dir + "\\" + "txt_files_2_html")

    for name in exam_file_base_name:
        for page_i in page_vec:
            exam_xlsx_file_name = fr"{name}_评分结果{page_i}.xlsx"

            overwrite_excels_by_union(
                Final_dir + "\\" + "txt_files_1_ordered" + "\\" + exam_xlsx_file_name,
                Final_dir + "\\" + "txt_files_2_html" + "\\" + exam_xlsx_file_name,
            )

            compare_excel_rows(
                file1_path=Final_dir + "\\" + "txt_files_1_ordered" + "\\" + exam_xlsx_file_name,
                file2_path=Final_dir + "\\" + "txt_files_2_html" + "\\" + exam_xlsx_file_name,
                output_path=Final_dir + "\\" + exam_xlsx_file_name,
                label_file1="学习通提交版评分",
                label_file2="学习通提交版本比压缩包短，以html评分"
            )

            rearrange_columns(Final_dir + "\\" + exam_xlsx_file_name)

    for name in exam_file_base_name:
        merge_excels_by_first_column(
            file_dir=Final_dir,
            file_names=[
                Final_dir + "\\" + name + "_评分结果page1.xlsx",
                Final_dir + "\\" + name + "_评分结果page2.xlsx",
                Final_dir + "\\" + name + "_评分结果page3.xlsx"
            ],
            output_name=Final_dir + "\\" + name + "_评分总表.xlsx"
        )
