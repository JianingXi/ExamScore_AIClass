from ExamScore_ExamQuizAnswers.A_00_extract_files_in_directory import extract_all_in_directory
from ExamScore_ExamQuizAnswers.A_02_convert_docx_to_txt import select_ans_txt_files

import numpy as np
import os
import pickle

folder_path_t0 = r"C:\Users\xijia\Desktop\批改web\S01_Raw"

extract_all_in_directory(folder_path_t0)  # 解压当前文文件夹所有压缩包
exam_file_base_name_t1 = [name for name in os.listdir(folder_path_t0)  # 获取所有子文件夹名称
                          if os.path.isdir(os.path.join(folder_path_t0, name))]

for i_rar_t1 in range(len(exam_file_base_name_t1)):
    folder_path_t1 = os.path.join(folder_path_t0, exam_file_base_name_t1[i_rar_t1])
    extract_all_in_directory(folder_path_t1)  # 解压当前文文件夹所有压缩包
    exam_file_base_name_t2 = [name for name in os.listdir(folder_path_t1)  # 获取所有子文件夹名称
                              if os.path.isdir(os.path.join(folder_path_t1, name))]
    for i_rar_t2 in range(len(exam_file_base_name_t2)):
        folder_path_t2 = os.path.join(folder_path_t1, exam_file_base_name_t2[i_rar_t2])
        extract_all_in_directory(folder_path_t2)  # 解压当前文文件夹所有压缩包



delete_num_array = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10]
select_ans_txt_files(folder_path, txt_folder_path, delete_num_array)

