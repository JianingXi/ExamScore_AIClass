from ExamScore_ExamQuizAnswers.A_00_extract_files_in_directory import extract_all_in_directory
from ExamScore_ExamQuizAnswers.A_02_convert_docx_to_txt import select_ans_txt_files
from ExamScore_ExamQuizAnswers.A_03_Web01_sim_match_best_ans import sim_match_best_ans
import numpy as np
import os
import pickle


folder_path = r"C:\Users\xijia\Desktop\腰部肌电信号采集数据\B01处理_验证码"
extract_all_in_directory(folder_path)
# 获取所有子文件夹名称
exam_file_base_name = [name for name in os.listdir(folder_path)
                if os.path.isdir(os.path.join(folder_path, name))]

for i_rar in range(len(exam_file_base_name)):
    folder_path_t = os.path.join(folder_path, exam_file_base_name[i_rar])
    extract_all_in_directory(folder_path_t)

    inner_file_base_name = [name for name in os.listdir(folder_path_t)
                           if os.path.isdir(os.path.join(folder_path_t, name))]

    for i_inner_rar in range(len(inner_file_base_name)):
        # 获取所有子文件夹名称
        folder_path_t_inner = os.path.join(folder_path_t, inner_file_base_name[i_inner_rar])
        extract_all_in_directory(folder_path_t_inner)

    delete_num_array = [1, 2]
    txt_folder_path = folder_path_t + r"\txt_files"
    select_ans_txt_files(folder_path_t, txt_folder_path, delete_num_array)



    # 保存结果的列表
    combined_folder_names = []

    # 遍历一级文件夹
    for first_level_name in os.listdir(folder_path_t):
        first_level_path = os.path.join(folder_path_t, first_level_name)
        if os.path.isdir(first_level_path):
            # 初始化拼接字符串：从一级目录名称开始
            folder_name_str = first_level_name

            # 遍历该一级目录下的所有子目录（递归）
            for dirpath, dirnames, filenames in os.walk(first_level_path):
                # 获取相对路径中剩下的子目录
                for dirname in dirnames:
                    folder_name_str += "_" + dirname

            # 添加到结果列表中
            combined_folder_names.append(folder_name_str)

    # 输出结果
    print(combined_folder_names)

# 然后需要人来检查命名情况
# 然后需要人来检查命名情况

