import pandas as pd
import re

def qualitative_replace_comments(input_path: str, output_path: str):
    """
    读取 Excel 文件，针对包含“评语”关键词的列：
    1️⃣ 删除单元格中所有斜杠及其后内容
    2️⃣ 使用正则替换数字/单位为定性描述
    3️⃣ 句子拆分和清理
    4️⃣ 将结果保存到新的 Excel 文件
    """
    df = pd.read_excel(input_path)
    comment_cols = [col for col in df.columns if '评语' in col]

    # 1️⃣ 删除所有斜杠及其后内容
    def remove_slash_content(cell):
        if pd.isna(cell):
            return cell
        cell_str = str(cell)
        if '/' in cell_str:
            return cell_str.split('/')[0]
        return cell_str

    df = df.applymap(remove_slash_content)

    # 2️⃣ 替换成定性描述
    def replace_with_qualitative(text):
        if pd.isna(text):
            return text
        text = str(text)

        text = re.sub(r'演讲时长[:：]?\s*[\d\.]+\s*(秒|分钟)?', '演讲时长适中', text)
        text = re.sub(r'时长不足.*?(分钟)?', '时长偏短', text)
        text = re.sub(r'语速[\d\.]*\s*(字/分|字)?', '语速适中', text)
        text = re.sub(r'平均每页.*?行.*?行以内', '文字行数适中', text)
        text = re.sub(r'图片覆盖面积[:：]?\s*[\d\.]+', '图片覆盖面积适中', text)
        text = re.sub(r'比例[:：]?\s*[\d\.]+%', '图片覆盖比例适中', text)
        text = re.sub(r'扣\d+(\.\d+)?分', '存在扣分项', text)

        # 删除括号及其中内容
        text = re.sub(r'（.*?）|\(.*?\)', '', text)

        # 处理多余空格
        text = re.sub(r'\s+', ' ', text).strip()

        # 句子拆分，过滤掉太短或残缺的句子
        sentences = re.split(r'[。；]', text)
        cleaned_sentences = []
        for sentence in sentences:
            sentence = sentence.strip('，。； ')
            if len(sentence) >= 4 and re.search(r'[\u4e00-\u9fa5]{2,}', sentence):
                cleaned_sentences.append(sentence)

        cleaned_text = '。'.join(cleaned_sentences) + '。' if cleaned_sentences else ''
        return cleaned_text

    # 3️⃣ 应用于所有“评语”列
    for col in comment_cols:
        df[col] = df[col].apply(replace_with_qualitative)

    # 4️⃣ 保存结果
    df.to_excel(output_path, index=False)
    print("已将所有数字及单位替换为定性描述（低中高等），文件已保存到：", output_path)

def a05_02_remove_data_from_comments(input_file: str, output_file: str):
    # input_file = r"C:\MyDocument\ToDoList\D20_DoingPlatform\D20_医学人工智能\结课论文\MergedResult.xlsx"
    # output_file = r"C:\MyDocument\ToDoList\D20_DoingPlatform\D20_医学人工智能\结课论文\MergedResult_QualitativeDescriptions.xlsx"
    qualitative_replace_comments(input_file, output_file)
