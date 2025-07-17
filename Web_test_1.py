from ExamScore_ExamQuizAnswers.A_00_extract_files_in_directory import extract_all_in_directory
from ExamScore_ExamQuizAnswers.A_02_convert_docx_to_txt import select_ans_txt_files
from ExamScore_ExamQuizAnswers.A_03_Web01_sim_match_best_ans import sim_match_best_ans
import numpy as np
import os
import pickle


exam_file_base_name = \
    ["儿生预web1班2-1-2024-2025-2_Web_A卷_-2_word_",
     "临I南山Web2班2-2-2024-2025-2_Web_A卷_-2_word_",
     "临2六中web2班2-3-2024-2025-2_Web期末考试_B卷__word_",
     "7.2计算机课程重考A2-502_10_10-11_10_-计算机课程_web_2024-2025-2重修A卷_7.2日__word_",
     "7.3计算机课程重考A2-502_10_10-11_10_-2024-2025-2_Web重修_C卷-7.3_word_",
     "精食药临药web1班4-2-2024-2025-2_web期末考试_C卷__word_",
     "公管心法Web1班4-1-2024-2025-2_web期末考试_C卷__word_",
     "护理检验web1班4-3-2024-2025-2_Web期末考试_护理_检验__word_",
     "口麻影Web1班4-4-2024-2025-2_Web期末考试_影像_麻醉_口腔__word_",
    ]





folder_path = r"C:\Users\xijia\Desktop\批改web\S01_Raw"
"""
extract_all_in_directory(folder_path)

# 获取所有子文件夹名称
exam_file_base_name = [name for name in os.listdir(folder_path)
                if os.path.isdir(os.path.join(folder_path, name))]

for i_rar in range(len(exam_file_base_name)):
    folder_path_t = os.path.join(folder_path, exam_file_base_name[i_rar])
    extract_all_in_directory(folder_path_t)
"""
# delete_num_array = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10]
# select_ans_txt_files(folder_path, delete_num_array)


txt_folder_path = r"C:\Users\xijia\Desktop\批改web\S01_Raw\txt_files"



indices = [0, 1, 3]  # 手动指定索引
exam_file_base_name_ABCD = np.array(exam_file_base_name)[indices].tolist()
html_path = r"C:\Users\xijia\Desktop\批改web\S01_Raw\A卷"
html_file_base_name = ["page1", "page2", "page3"]  # "page0",

skip_question = 10
num_question = 4

sim_match_best_ans(txt_folder_path, exam_file_base_name_ABCD, html_path, html_file_base_name, skip_question, num_question)



indices = [2]  # 手动指定索引
exam_file_base_name_ABCD = np.array(exam_file_base_name)[indices].tolist()
html_path = r"C:\Users\xijia\Desktop\批改web\S01_Raw\B卷"
html_file_base_name = ["page1", "page2", "page3"]  # "page0",

skip_question = 10
num_question = 3

sim_match_best_ans(txt_folder_path, exam_file_base_name_ABCD, html_path, html_file_base_name, skip_question, num_question)




indices = [4, 5, 6]  # 手动指定索引
exam_file_base_name_ABCD = np.array(exam_file_base_name)[indices].tolist()
html_path = r"C:\Users\xijia\Desktop\批改web\S01_Raw\C卷"
html_file_base_name = ["page1", "page2", "page3"]  # "page0",

skip_question = 10
num_question = 3
print(exam_file_base_name_ABCD)
sim_match_best_ans(txt_folder_path, exam_file_base_name_ABCD, html_path, html_file_base_name, skip_question, num_question)





indices = [7, 8]  # 手动指定索引
exam_file_base_name_ABCD = np.array(exam_file_base_name)[indices].tolist()
html_path = r"C:\Users\xijia\Desktop\批改web\S01_Raw\D卷"
html_file_base_name = ["page1", "page2", "page3"]  # "page0",

skip_question = 10
num_question = 3
print(exam_file_base_name_ABCD)
sim_match_best_ans(txt_folder_path, exam_file_base_name_ABCD, html_path, html_file_base_name, skip_question, num_question)
