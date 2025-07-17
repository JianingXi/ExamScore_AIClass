from ExamScore_ExamQuizAnswers.A_03_Web03_compare_and_merge import (overwrite_first_col_with_id_name,
                                                                    overwrite_excels_by_union,
                                                                    compare_excel_rows,
                                                                    rearrange_columns,
                                                                    merge_excels_by_first_column)

# overwrite_first_col_with_id_name(r"C:\Users\xijia\Desktop\批改web\S04_test_files_merged\html")
# overwrite_first_col_with_id_name(r"C:\Users\xijia\Desktop\批改web\S04_test_files_merged\ordered")



exam_xlsx_file_name = [
    "儿生预web1班2-1-2024-2025-2_Web_A卷_-2_word_评分结果page1.xlsx",
    "儿生预web1班2-1-2024-2025-2_Web_A卷_-2_word_评分结果page2.xlsx",
    "儿生预web1班2-1-2024-2025-2_Web_A卷_-2_word_评分结果page3.xlsx",
    "公管心法Web1班4-1-2024-2025-2_web期末考试_C卷__word_评分结果page1.xlsx",
    "公管心法Web1班4-1-2024-2025-2_web期末考试_C卷__word_评分结果page2.xlsx",
    "公管心法Web1班4-1-2024-2025-2_web期末考试_C卷__word_评分结果page3.xlsx",
    "护理检验web1班4-3-2024-2025-2_Web期末考试_护理_检验__word_评分结果page1.xlsx",
    "护理检验web1班4-3-2024-2025-2_Web期末考试_护理_检验__word_评分结果page2.xlsx",
    "护理检验web1班4-3-2024-2025-2_Web期末考试_护理_检验__word_评分结果page3.xlsx",
    "精食药临药web1班4-2-2024-2025-2_web期末考试_C卷__word_评分结果page1.xlsx",
    "精食药临药web1班4-2-2024-2025-2_web期末考试_C卷__word_评分结果page2.xlsx",
    "精食药临药web1班4-2-2024-2025-2_web期末考试_C卷__word_评分结果page3.xlsx",
    "口麻影Web1班4-4-2024-2025-2_Web期末考试_影像_麻醉_口腔__word_评分结果page1.xlsx",
    "口麻影Web1班4-4-2024-2025-2_Web期末考试_影像_麻醉_口腔__word_评分结果page2.xlsx",
    "口麻影Web1班4-4-2024-2025-2_Web期末考试_影像_麻醉_口腔__word_评分结果page3.xlsx",
    "临2六中web2班2-3-2024-2025-2_Web期末考试_B卷__word_评分结果page1.xlsx",
    "临2六中web2班2-3-2024-2025-2_Web期末考试_B卷__word_评分结果page2.xlsx",
    "临2六中web2班2-3-2024-2025-2_Web期末考试_B卷__word_评分结果page3.xlsx",
    "临I南山Web2班2-2-2024-2025-2_Web_A卷_-2_word_评分结果page1.xlsx",
    "临I南山Web2班2-2-2024-2025-2_Web_A卷_-2_word_评分结果page2.xlsx",
    "临I南山Web2班2-2-2024-2025-2_Web_A卷_-2_word_评分结果page3.xlsx",
]

for name in exam_xlsx_file_name:
    overwrite_excels_by_union(
        r"C:\Users\xijia\Desktop\批改web\S04_test_files_merged\html" + "\\" + name,
        r"C:\Users\xijia\Desktop\批改web\S04_test_files_merged\ordered" + "\\" + name,
    )

    compare_excel_rows(
        file1_path=r"C:\Users\xijia\Desktop\批改web\S04_test_files_merged\ordered" + "\\" + name,
        file2_path=r"C:\Users\xijia\Desktop\批改web\S04_test_files_merged\html" + "\\" + name,
        output_path=r"C:\Users\xijia\Desktop\批改web\S04_test_files_merged" + "\\" + name,
        label_file1="学习通提交版评分",
        label_file2="学习通提交版本比压缩包短，以html评分"
    )

    rearrange_columns(r"C:\Users\xijia\Desktop\批改web\S04_test_files_merged" + "\\" + name)



exam_file_base_name = [
    "儿生预web1班2-1-2024-2025-2_Web_A卷_-2_word_",
    "临I南山Web2班2-2-2024-2025-2_Web_A卷_-2_word_",
    "临2六中web2班2-3-2024-2025-2_Web期末考试_B卷__word_",
    "精食药临药web1班4-2-2024-2025-2_web期末考试_C卷__word_",
    "公管心法Web1班4-1-2024-2025-2_web期末考试_C卷__word_",
    "护理检验web1班4-3-2024-2025-2_Web期末考试_护理_检验__word_",
    "口麻影Web1班4-4-2024-2025-2_Web期末考试_影像_麻醉_口腔__word_",
]

for name in exam_file_base_name:
    merge_excels_by_first_column(
        file_dir=r"C:\Users\xijia\Desktop\批改web\S04_test_files_merged",
        file_names=[
            name + "评分结果page1.xlsx",
            name + "评分结果page2.xlsx",
            name + "评分结果page3.xlsx"
        ],
        output_name=r"C:\Users\xijia\Desktop\批改web\S04_test_files_merged" + "\\" + name + "评分总表.xlsx"
    )
