from ExamScore_ExamQuizAnswers.A_03_Web01_sim_match_best_ans import batch_similarity_move_by_excel, rename_txt_files_by_student_info, batch_txt_files_by_student_info


# "7.2计算机课程重考A2-502_10_10-11_10_-计算机课程_web_2024-2025-2重修A卷_7.2日__word_",
# "7.3计算机课程重考A2-502_10_10-11_10_-2024-2025-2_Web重修_C卷-7.3_word_",


exam_folder_path = r"C:\Users\xijia\Desktop\批改web"
exam_file_base_name = [
    "儿生预web1班2-1-2024-2025-2_Web_A卷_-2_word_",
    "临I南山Web2班2-2-2024-2025-2_Web_A卷_-2_word_",
    "临2六中web2班2-3-2024-2025-2_Web期末考试_B卷__word_",
    "精食药临药web1班4-2-2024-2025-2_web期末考试_C卷__word_",
    "公管心法Web1班4-1-2024-2025-2_web期末考试_C卷__word_",
    "护理检验web1班4-3-2024-2025-2_Web期末考试_护理_检验__word_",
    "口麻影Web1班4-4-2024-2025-2_Web期末考试_影像_麻醉_口腔__word_",
]


batch_similarity_move_by_excel(exam_folder_path, exam_file_base_name)
batch_txt_files_by_student_info(exam_folder_path, exam_file_base_name)