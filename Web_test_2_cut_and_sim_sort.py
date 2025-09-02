from ExamScore_ExamQuizAnswers.A_02_convert_docx_to_txt import select_ans_txt_files
from ExamScore_ExamQuizAnswers.A_03_Web01_sim_match_best_ans import sim_match_best_ans
from ExamScore_ExamQuizAnswers.A_03_Web01_sim_match_best_ans import batch_similarity_move_by_excel, batch_txt_rename_by_student_info
import numpy as np
import os
import pickle

folder_path = r"C:\Users\xijia\Desktop\批改web\S01_Raw"
txt_folder_path = r"C:\Users\xijia\Desktop\批改web\S02_TXT"


delete_num_array = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10]
select_ans_txt_files(folder_path, txt_folder_path, delete_num_array)















exam_file_base_name = \
    ["儿生预web1班2-1-2024-2025-2_Web（A卷）-2(word)",
     "临I南山Web2班2-2-2024-2025-2_Web（A卷）-2(word)",
     "临2六中web2班2-3-2024-2025-2_Web期末考试（B卷）(word)",
     "公管心法Web1班4-1-2024-2025-2_web期末考试（C卷）(word)",
     "精食药临药web1班4-2-2024-2025-2_web期末考试（C卷）(word)",
     "护理检验web1班4-3-2024-2025-2_Web期末考试（护理、检验）(word)",
     "口麻影Web1班4-4-2024-2025-2_Web期末考试（影像、麻醉、口腔）(word)",
    ]


indices = [0, 1]  # 手动指定索引
exam_file_base_name_ABCD = np.array(exam_file_base_name)[indices].tolist()
html_path = r"C:\Users\xijia\Desktop\批改web\S01_Dict\A卷"
html_file_base_name = ["page1", "page2", "page3"]  # "page0",

skip_question = 10
num_question = 4

sim_match_best_ans(txt_folder_path, exam_file_base_name_ABCD, html_path, html_file_base_name, skip_question, num_question)



indices = [2]  # 手动指定索引
exam_file_base_name_ABCD = np.array(exam_file_base_name)[indices].tolist()
html_path = r"C:\Users\xijia\Desktop\批改web\S01_Dict\B卷"
html_file_base_name = ["page1", "page2", "page3"]  # "page0",

skip_question = 10
num_question = 3

sim_match_best_ans(txt_folder_path, exam_file_base_name_ABCD, html_path, html_file_base_name, skip_question, num_question)




indices = [3, 4]  # 手动指定索引
exam_file_base_name_ABCD = np.array(exam_file_base_name)[indices].tolist()
html_path = r"C:\Users\xijia\Desktop\批改web\S01_Dict\C卷"
html_file_base_name = ["page1", "page2", "page3"]  # "page0",

skip_question = 10
num_question = 3

sim_match_best_ans(txt_folder_path, exam_file_base_name_ABCD, html_path, html_file_base_name, skip_question, num_question)





indices = [5, 6]  # 手动指定索引
exam_file_base_name_ABCD = np.array(exam_file_base_name)[indices].tolist()
html_path = r"C:\Users\xijia\Desktop\批改web\S01_Dict\D卷"
html_file_base_name = ["page1", "page2", "page3"]

skip_question = 10
num_question = 3

sim_match_best_ans(txt_folder_path, exam_file_base_name_ABCD, html_path, html_file_base_name, skip_question, num_question)




exam_folder_sim_excel_path = txt_folder_path
exam_folder_ordered_path = txt_folder_path + "\\" + "txt_files_1_ordered"

batch_similarity_move_by_excel(exam_folder_sim_excel_path, exam_folder_ordered_path, exam_file_base_name)
batch_txt_rename_by_student_info(exam_folder_ordered_path, exam_file_base_name)


