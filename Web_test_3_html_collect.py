from ExamScore_ExamQuizAnswers.A_03_Web02_html_backup_ans import convert_html_to_backup_answers


exam_folder_path_raw = r"C:\Users\xijia\Desktop\批改web\S01_Raw"
exam_folder_path_html = r"C:\Users\xijia\Desktop\批改web\S02_TXT\txt_files_2_html"
exam_file_base_name = \
    ["儿生预web1班2-1-2024-2025-2_Web（A卷）-2(word)",
     "临I南山Web2班2-2-2024-2025-2_Web（A卷）-2(word)",
     "临2六中web2班2-3-2024-2025-2_Web期末考试（B卷）(word)",
     "公管心法Web1班4-1-2024-2025-2_web期末考试（C卷）(word)",
     "精食药临药web1班4-2-2024-2025-2_web期末考试（C卷）(word)",
     "护理检验web1班4-3-2024-2025-2_Web期末考试（护理、检验）(word)",
     "口麻影Web1班4-4-2024-2025-2_Web期末考试（影像、麻醉、口腔）(word)",
    ]


convert_html_to_backup_answers(exam_folder_path_raw, exam_folder_path_html, exam_file_base_name)
