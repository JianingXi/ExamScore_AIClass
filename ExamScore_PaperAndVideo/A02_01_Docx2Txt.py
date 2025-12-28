import os
import re
from docx import Document
import win32com.client as win32

# -------------------- 文件转换部分 --------------------
def convert_doc_to_docx(root_dir):
    word = win32.Dispatch("Word.Application")
    word.Visible = False

    for dirpath, _, filenames in os.walk(root_dir):
        for filename in filenames:
            if filename.lower().endswith('.doc') and not filename.lower().endswith('.docx'):
                doc_path = os.path.join(dirpath, filename)
                docx_path = os.path.join(dirpath, os.path.splitext(filename)[0] + '.docx')

                print(f"正在转换: {doc_path} 到 {docx_path}")
                try:
                    doc = word.Documents.Open(doc_path)
                    doc.SaveAs2(docx_path, FileFormat=16)  # 16 = docx
                    doc.Close()
                    os.remove(doc_path)
                    print(f"转换成功并删除: {doc_path}")
                except Exception as e:
                    print(f"处理文件时发生错误: {doc_path}\n错误信息: {e}")

    word.Quit()

# -------------------- 转录识别部分 --------------------
# Regex patterns
TIMESTAMP_PATTERN = re.compile(r"\d{1,2}:\d{2}(?::\d{2})?")
TIMESTAMP_BRACKET_PATTERN = re.compile(r"\[\d{1,2}:\d{2}(?::\d{2})?\]")
SPEAKER_PATTERN = re.compile(r"^[^\n]+\(\d{1,2}:\d{2}(?::\d{2})?\):", re.MULTILINE)
NAME_COLON_PATTERN = re.compile(r"^[A-Za-z][A-Za-z0-9_ ]{0,20}:")
PAPER_KEYWORDS_PATTERN = re.compile(r"^(摘要|Abstract|Keywords?[:：]|关键词[:：]|References|参考文献)", re.MULTILINE)
CHINESE_HEADING_PATTERN = re.compile(r"^[一二三四五六七八九十]、")

# 新增：演讲稿/讲稿等关键词
SPEECH_KEYWORDS = [
    "讲稿", "演讲稿", "演讲提纲", "讲演稿", "演说稿", "发言稿",
    "讲话稿", "发言提纲", "会议发言", "汇报稿", "文字稿"
]

def is_transcript(text: str) -> bool:
    """
    Heuristic to determine if a document text resembles a transcript or a speech script.
    """
    if not text:
        return False

    # ✅ 关键词命中：演讲稿/讲稿相关
    if any(keyword in text for keyword in SPEECH_KEYWORDS):
        return True

    if PAPER_KEYWORDS_PATTERN.search(text):
        return False
    if CHINESE_HEADING_PATTERN.search(text):
        return False

    lines = [l for l in text.splitlines() if l.strip()]
    if not lines:
        return False

    short_lines_count = sum(1 for l in lines if len(l) < 60)
    long_lines_count = sum(1 for l in lines if len(l) > 100)

    timestamps = TIMESTAMP_PATTERN.findall(text)
    bracketed = TIMESTAMP_BRACKET_PATTERN.findall(text)
    speakers = SPEAKER_PATTERN.findall(text)
    names = NAME_COLON_PATTERN.findall(text)

    if len(timestamps) >= 3 or len(bracketed) >= 3:
        return True
    if speakers:
        return True
    if len(names) >= 3:
        return True
    if len(lines) >= 10 and (short_lines_count / len(lines)) > 0.6:
        return True
    if len(lines) >= 5 and (long_lines_count / len(lines)) > 0.5:
        return False

    return False

# -------------------- 文件读取 --------------------
def extract_docx_text(path: str) -> str | None:
    try:
        doc = Document(path)
    except Exception as e:
        print(f"Error reading DOCX '{path}': {e}")
        return None
    full_text = []
    for para in doc.paragraphs:
        if para.text:
            full_text.append(para.text.strip())
    return "\n".join(full_text)

# -------------------- 文件夹处理 --------------------
def process_folder(folder_path: str):
    try:
        files = os.listdir(folder_path)
    except Exception as e:
        print(f"Cannot list directory '{folder_path}': {e}")
        return

    if any(f.lower().endswith('.txt') for f in files):
        return

    for f in files:
        if f.lower().endswith('.docx'):
            docx_path = os.path.join(folder_path, f)
            text = extract_docx_text(docx_path)
            if text is None:
                continue
            if is_transcript(text):
                txt_name = os.path.splitext(f)[0] + '_srt' + '.txt'
                txt_path = os.path.join(folder_path, txt_name)
                try:
                    with open(txt_path, 'w', encoding='utf-8') as out:
                        out.write(text)
                    os.remove(docx_path)
                    print(f"已生成TXT并删除原始文件: {docx_path}")
                except Exception as e:
                    print(f"Error writing TXT '{txt_path}': {e}")
            else:
                print(f"Skipped non-transcript: {docx_path}")

def convert_srt_to_txt(folder_path: str):
    for filename in os.listdir(folder_path):
        if filename.lower().endswith('.srt'):
            srt_path = os.path.join(folder_path, filename)
            txt_filename = os.path.splitext(filename)[0] + '_srt' + '.txt'
            txt_path = os.path.join(folder_path, txt_filename)
            try:
                with open(srt_path, 'r', encoding='utf-8', errors='ignore') as sf:
                    content = sf.read()
                with open(txt_path, 'w', encoding='utf-8') as tf:
                    tf.write(content)
                os.remove(srt_path)
                print(f"已将SRT转TXT并删除: {srt_path}")
            except Exception as e:
                print(f"Error processing {srt_path}: {e}")

def rename_txt_files(folder_path: str):
    for filename in os.listdir(folder_path):
        if filename.lower().endswith('.txt') and not filename.lower().endswith('_srt.txt'):
            old_path = os.path.join(folder_path, filename)
            base_name = filename[:-4]
            new_filename = f"{base_name}_srt.txt"
            new_path = os.path.join(folder_path, new_filename)
            try:
                os.rename(old_path, new_path)
                print(f"已重命名: {old_path} -> {new_path}")
            except Exception as e:
                print(f"Error renaming {old_path}: {e}")

# -------------------- 主执行函数 --------------------
def a02_01_docx2txt(BASE_DIR: str):
    convert_doc_to_docx(BASE_DIR)

    for root, dirs, _ in os.walk(BASE_DIR):
        for d in dirs:
            folder = os.path.join(root, d)
            process_folder(folder)
            convert_srt_to_txt(folder)
            rename_txt_files(folder)

# ✅ 使用例：
# BASE_DIR = r"C:\MyPython\ExamScore_AIClass\ExamFiles"
# a02_01_docx2txt(BASE_DIR)
