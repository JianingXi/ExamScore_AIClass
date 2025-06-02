import os
import pickle
import numpy as np
import pandas as pd
from sklearn.ensemble import IsolationForest
from scipy.sparse import hstack

SECTION_KEYS = ['题目', '摘要', '引言', '结论']

def load_embeddings(base_dir):
    embeddings = []
    folder_names = []

    for folder in sorted(os.listdir(base_dir)):
        folder_path = os.path.join(base_dir, folder)
        tfidf_path = os.path.join(folder_path, 'tf_idf.pkl')
        if not os.path.exists(tfidf_path):
            continue

        try:
            with open(tfidf_path, 'rb') as f:
                tfidf = pickle.load(f)

            vectors = [tfidf.get(k) for k in SECTION_KEYS if k in tfidf]
            if len(vectors) != 4:
                continue

            merged_vector = hstack(vectors).toarray()[0]
            embeddings.append(merged_vector)
            folder_names.append(folder)

        except Exception as e:
            print(f"读取失败：{folder_path}，原因：{e}")

    return folder_names, np.array(embeddings)

def detect_outliers_and_score(folder_names, embeddings, base_dir):
    if len(embeddings) < 5:
        print("数据过少，无法有效进行离群点检测。")
        return

    clf = IsolationForest(n_estimators=100, contamination=0.1, random_state=42)
    preds = clf.fit_predict(embeddings)  # -1 表示异常点

    outliers = [folder for folder, pred in zip(folder_names, preds) if pred == -1]

    # 写入离群点列表
    out_path = os.path.join(base_dir, 'Outliers.txt')
    with open(out_path, 'w', encoding='utf-8') as f:
        f.write('\n'.join(outliers))
    print(f"检测完成，共识别出 {len(outliers)} 个离群点，已写入 {out_path}")

    # Excel 结果
    records = []
    for folder, pred in zip(folder_names, preds):
        if pred == -1:
            score = 0
            remark = "离群点"
        else:
            score = 8
            remark = "正常"
        records.append({
            "学生文件夹": folder,
            "内容得分": f"{score}/10",
            "备注": remark
        })

    df = pd.DataFrame(records)
    excel_path = os.path.join(base_dir, "ContentSemanticScore.xlsx")
    df.to_excel(excel_path, index=False)
    print(f"✅ 已生成内容得分 Excel：{excel_path}")

def a04_07_content_emb_outlier(BASE_DIR: str):
    folder_names, embeddings = load_embeddings(BASE_DIR)
    detect_outliers_and_score(folder_names, embeddings, BASE_DIR)
