import pandas as pd
import os

txt_folder_path = r"C:\Users\xijia\Desktop\批改web\S02_txt_files"





import os
import shutil
import pandas as pd

def copy_and_rename_by_similarity_excel(
    input_folder: str,
    output_folder: str,
    excel_path: str
):
    """
    递归扫描 input_folder 中所有 _0XX.txt 文件，根据 excel_path 中的相似性表格，
    找出对应的 page_X，复制文件并将其命名为 _page_X.txt 到 output_folder。
    """

    os.makedirs(output_folder, exist_ok=True)
    df = pd.read_excel(excel_path)

    # 自动识别 page_X 列
    page_cols = [col for col in df.columns if str(col).startswith("page_")]

    # 递归查找所有 .txt 文件
    for root, _, files in os.walk(input_folder):
        for filename in files:
            if not filename.endswith(".txt") or "_0" not in filename:
                continue

            full_path = os.path.join(root, filename)
            basename_match = filename.rsplit("_0", 1)[0]
            suffix = filename[-7:-4]  # 取 '013'

            # 查找 DataFrame 匹配的 basename 行
            row = df[df["basename"].str.contains(basename_match)]
            if row.empty:
                print(f"❌ 未找到匹配 basename: {basename_match}")
                continue

            row = row.iloc[0]
            new_page = None

            for col in page_cols:
                if str(row[col]).zfill(3) == suffix:
                    new_page = col
                    break

            if new_page is None:
                print(f"❗ 未匹配题号 {suffix} 于 {basename_match}")
                continue

            new_filename = f"{basename_match}_{new_page}.txt"
            new_path = os.path.join(output_folder, new_filename)

            shutil.copy(full_path, new_path)
            # print(f"✅ 已复制并重命名：{filename} → {new_filename}")








exam_file_base_name = [
    "儿生预web1班2-1-2024-2025-2_Web_A卷_-2_word_",
]
if 1 < 0:
    """
    "临I南山Web2班2-2-2024-2025-2_Web_A卷_-2_word_",
    "临2六中web2班2-3-2024-2025-2_Web期末考试_B卷__word_",
    "7.2计算机课程重考A2-502_10_10-11_10_-计算机课程_web_2024-2025-2重修A卷_7.2日__word_",
    "7.3计算机课程重考A2-502_10_10-11_10_-2024-2025-2_Web重修_C卷-7.3_word_",
    "精食药临药web1班4-2-2024-2025-2_web期末考试_C卷__word_",
    "公管心法Web1班4-1-2024-2025-2_web期末考试_C卷__word_",
    "护理检验web1班4-3-2024-2025-2_Web期末考试_护理_检验__word_",
    "口麻影Web1班4-4-2024-2025-2_Web期末考试_影像_麻醉_口腔__word_",
    """
    a = 1

# 存储所有 DataFrame 的字典
excel_data_dict = {}

for name in exam_file_base_name:
    file_path = os.path.join(txt_folder_path, name + "_相似性.xlsx")
    if os.path.exists(file_path):
        df_file_match = pd.read_excel(file_path)
        # excel_data_dict[name] = df
        print(df_file_match)
        copy_and_rename_by_similarity_excel(
            input_folder=r"C:\Users\xijia\Desktop\批改web\S02_txt_files\儿生预web1班2-1-2024-2025-2_Web_A卷_-2_word_",
            output_folder=r"C:\Users\xijia\Desktop\批改web\S03_new_files\儿生预web1班2-1-2024-2025-2_Web_A卷_-2_word_",
            excel_path=r"C:\Users\xijia\Desktop\批改web\S02_txt_files\儿生预web1班2-1-2024-2025-2_Web_A卷_-2_word__相似性.xlsx"
        )


    else:
        print(f"❌ 文件不存在：{file_path}")
