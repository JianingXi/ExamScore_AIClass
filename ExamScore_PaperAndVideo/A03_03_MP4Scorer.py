import os
import hashlib
from docx import Document
import pandas as pd

def score_docx_by_variation(docx_path):
    doc = Document(docx_path)
    hashes = []
    for shape in doc.inline_shapes:
        rId = shape._inline.graphic.graphicData.pic.blipFill.blip.embed
        blob = doc.part.related_parts[rId].blob
        h = hashlib.md5(blob).hexdigest()
        hashes.append(h)

    unique_count = len(set(hashes))
    if unique_count >= 8:
        score = 5
        comment = "本次作业画面种类丰富，展示了多个不同环节，步骤衔接自然，信息覆盖全面，整体演示非常出色。"
    elif unique_count >= 5:
        score = 4
        comment = "本次作业展示了主要环节的多张不同画面，能够较好地呈现操作流程，建议再补充一两处关键步骤以臻完美。"
    elif unique_count >= 3:
        score = 3
        comment = "画面涵盖了核心步骤，信息传递基本到位，下次可以增加更多节点画面，使流程展示更完整。"
    elif unique_count >= 1:
        score = 2
        comment = "画面高度雷同（疑似是没开共享屏幕），无法充分了解全部过程，建议视频显示关键画面。"
    else:
        score = 0
        comment = "未检测到有效画面，无法评估演示内容，请检查视频有清晰画面。"
    return score, comment

def read_score_file(file_path):
    if not os.path.exists(file_path):
        return "", ""
    with open(file_path, 'r', encoding='utf-8') as f:
        lines = [line.strip() for line in f.readlines() if line.strip()]
        if not lines:
            return "", ""
        score = lines[0]
        comment = ' '.join(lines[1:]) if len(lines) > 1 else ""
        return score, comment

def process_all(base_dir):
    records = []
    for student in os.listdir(base_dir):
        folder = os.path.join(base_dir, student)
        if not os.path.isdir(folder):
            continue

        # MP4打分
        docs = [f for f in os.listdir(folder) if f.endswith("_mp4.docx")]
        if docs:
            path = os.path.join(folder, docs[0])
            mp4_score, mp4_comment = score_docx_by_variation(path)
        else:
            mp4_score, mp4_comment = 0, "未上传MP4文件，0分。"

        # 写入 MP4Score.txt
        with open(os.path.join(folder, "MP4Score.txt"), "w", encoding="utf-8") as f:
            f.write(f"Score: {mp4_score}/5\nComment: {mp4_comment}\n")

        # 读取 AudioScore.txt
        audio_score, audio_comment = read_score_file(os.path.join(folder, "AudioScore.txt"))
        # 读取 PPTScore.txt
        ppt_score, ppt_comment = read_score_file(os.path.join(folder, "PPTScore.txt"))
        # 记录
        records.append({
            "学生文件夹": student,
            "Audio得分": audio_score,
            "Audio评语": audio_comment,
            "PPT得分": ppt_score,
            "PPT评语": ppt_comment,
            "MP4得分": f"{mp4_score}/5",
            "MP4评语": mp4_comment
        })

    # 写入Excel
    if records:
        df = pd.DataFrame(records)
        excel_path = os.path.join(base_dir, "VideoScores.xlsx")
        df.to_excel(excel_path, index=False)
        print(f"✅ 已生成完整汇总表：{excel_path}")

def delete_media_docx_files(root_dir):
    media_extensions = ['_mp4.docx', '_mov.docx', '_avi.docx', '_mkv.docx']
    for root, _, files in os.walk(root_dir):
        for file in files:
            if any(file.lower().endswith(ext) for ext in media_extensions):
                try:
                    os.remove(os.path.join(root, file))
                    print(f"已删除: {file}")
                except Exception as e:
                    print(f"删除失败 {file}: {e}")

def a03_03_mp4_scorer(root_dir: str):
    process_all(root_dir)
    delete_media_docx_files(root_dir)
