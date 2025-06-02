import os
import pickle
from sklearn.feature_extraction.text import TfidfVectorizer

SECTION_KEYS = ['题目', '摘要', '引言', '结论']
SECTION_FILES = [f'Section_{k}.txt' for k in SECTION_KEYS]


def collect_all_texts(base_dir):
    """遍历所有文件夹，收集所有段落文本，准备构建共享词典"""
    corpus = []
    folder_map = {}  # {folder: {section: content}}

    for folder in sorted(os.listdir(base_dir)):
        folder_path = os.path.join(base_dir, folder)
        if not os.path.isdir(folder_path):
            continue

        section_texts = {}
        for sec_file, sec_key in zip(SECTION_FILES, SECTION_KEYS):
            path = os.path.join(folder_path, sec_file)
            if os.path.exists(path):
                with open(path, 'r', encoding='utf-8') as f:
                    text = f.read().strip()
                    section_texts[sec_key] = text
                    corpus.append(text)
            else:
                section_texts[sec_key] = ""

        folder_map[folder_path] = section_texts

    return corpus, folder_map


def extract_and_save_tfidf(base_dir):
    corpus, folder_map = collect_all_texts(base_dir)

    # 共享词袋词典
    vectorizer = TfidfVectorizer(max_features=1000)  # 可调节维度
    vectorizer.fit(corpus)

    for folder_path, sec_dict in folder_map.items():
        tfidf_result = {}
        for sec_key in SECTION_KEYS:
            text = sec_dict[sec_key]
            vector = vectorizer.transform([text])
            tfidf_result[sec_key] = vector

        # 保存为 pkl
        out_path = os.path.join(folder_path, 'tf_idf.pkl')
        with open(out_path, 'wb') as f:
            pickle.dump(tfidf_result, f)

    # 也可以选择保存 vectorizer，以便未来使用
    with open(os.path.join(base_dir, 'shared_vectorizer.pkl'), 'wb') as f:
        pickle.dump(vectorizer, f)


def a04_06_content_tf_idf(base_directory: str):
    # base_directory = r"C:\MyPython\ExamScore_AIClass\ExamFiles"
    extract_and_save_tfidf(base_directory)
