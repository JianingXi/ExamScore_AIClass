import os
import hashlib
from docx import Document

def score_docx_by_variation(docx_path):
    """
    基于画面“多样性”打分，修正图片提取方式：
    - 提取图片 rId 后，使用 doc.part.related_parts 获取 blob
    """
    doc = Document(docx_path)
    hashes = []
    for shape in doc.inline_shapes:
        # 获取嵌入图片 rId
        rId = shape._inline.graphic.graphicData.pic.blipFill.blip.embed
        # 从文档总体 part 中提取图片 blob
        blob = doc.part.related_parts[rId].blob
        # 计算 MD5 哈希
        h = hashlib.md5(blob).hexdigest()
        hashes.append(h)

    unique_count = len(set(hashes))
    # 打分与评语
    if unique_count >= 8:
        score = 5
        comment = (
            "本次作业画面种类丰富，展示了多个不同环节，"
            "步骤衔接自然，信息覆盖全面，整体演示非常出色。"
        )
    elif unique_count >= 5:
        score = 4
        comment = (
            "本次作业展示了主要环节的多张不同画面，"
            "能够较好地呈现操作流程，建议再补充一两处关键步骤以臻完美。"
        )
    elif unique_count >= 3:
        score = 3
        comment = (
            "画面涵盖了核心步骤，信息传递基本到位，"
            "下次可以增加更多节点画面，使流程展示更完整。"
        )
    elif unique_count >= 1:
        score = 2
        comment = (
            "画面高度雷同（疑似是没开共享屏幕），无法充分了解全部过程，"
            "建议视频显示关键画面。"
        )
    else:
        score = 0
        comment = (
            "未检测到有效画面，无法评估演示内容，"
            "请检查视频有清晰画面。"
        )

    return score, comment

def process_all(base_dir):
    """
    遍历 base_dir 下各子文件夹，对 *_mp4.docx 文档应用新规则打分，
    并将结果写入子文件夹内的 MP4Score.txt。
    """
    for student in os.listdir(base_dir):
        folder = os.path.join(base_dir, student)
        if not os.path.isdir(folder):
            continue

        docs = [f for f in os.listdir(folder) if f.endswith("_mp4.docx")]
        if docs:
            path = os.path.join(folder, docs[0])
            score, comment = score_docx_by_variation(path)
        else:
            score, comment = 0, (
                "未检测到MP4文件，0分。"
            )

        output_path = os.path.join(folder, "MP4Score.txt")
        with open(output_path, "w", encoding="utf-8") as f:
            f.write(f"Score: {score}/5\nComment: {comment}\n")

    print("MP4Score.txt")



def delete_mp4docx_files(root_dir):
    for root, dirs, files in os.walk(root_dir):
        for file in files:
            if file.lower().endswith(('_mp4.docx')):
                file_path = os.path.join(root, file)
                try:
                    os.remove(file_path)
                except Exception as e:
                    print(f"删除失败 {file_path}: {e}")

def a03_03_mp4_scorer(root_dir: str):
    # root_dir = r"C:\MyPython\ExamScore_AIClass\ExamFiles"
    # 将此处路径替换为你的 ExamFiles 目录
    process_all(root_dir)

    delete_mp4docx_files(root_dir)
