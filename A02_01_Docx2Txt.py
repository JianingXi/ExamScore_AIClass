import os
import re
from docx import Document

# Regex patterns to detect various transcript-like content
TIMESTAMP_PATTERN = re.compile(r"\d{1,2}:\d{2}(?::\d{2})?")
TIMESTAMP_BRACKET_PATTERN = re.compile(r"\[\d{1,2}:\d{2}(?::\d{2})?\]")
SPEAKER_PATTERN = re.compile(r"^[^\n]+\(\d{1,2}:\d{2}(?::\d{2})?\):", re.MULTILINE)
NAME_COLON_PATTERN = re.compile(r"^[A-Za-z][A-Za-z0-9_ ]{0,20}:")
PAPER_KEYWORDS_PATTERN = re.compile(r"^(摘要|Abstract|Keywords?[:：]|关键词[:：]|References|参考文献)", re.MULTILINE)
CHINESE_HEADING_PATTERN = re.compile(r"^[一二三四五六七八九十]、")


def is_transcript(text: str) -> bool:
    """
    Heuristic to determine if a document text resembles a transcript.
    Excludes common academic paper structures and looks for transcript cues.
    """
    if not text:
        return False

    # Exclude if academic paper keywords appear
    if PAPER_KEYWORDS_PATTERN.search(text):
        return False
    if CHINESE_HEADING_PATTERN.search(text):
        # Likely an academic section heading
        return False

    # Gather basic stats
    lines = [l for l in text.splitlines() if l.strip()]
    if not lines:
        return False
    short_lines_count = sum(1 for l in lines if len(l) < 60)
    long_lines_count = sum(1 for l in lines if len(l) > 100)

    # Transcript indicators
    timestamps = TIMESTAMP_PATTERN.findall(text)
    bracketed = TIMESTAMP_BRACKET_PATTERN.findall(text)
    speakers = SPEAKER_PATTERN.findall(text)
    names = NAME_COLON_PATTERN.findall(text)

    # 1. Timestamp occurrences
    if len(timestamps) >= 3 or len(bracketed) >= 3:
        return True
    # 2. Speaker labels
    if speakers:
        return True
    # 3. Generic 'Name:' prefix
    if len(names) >= 3:
        return True
    # 4. High proportion of short lines
    if len(lines) >= 10 and (short_lines_count / len(lines)) > 0.6:
        return True

    # Non-transcript signals: many long lines in academic prose
    if len(lines) >= 5 and (long_lines_count / len(lines)) > 0.5:
        return False

    return False


def extract_docx_text(path: str) -> str | None:
    """
    Extracts text from a .docx file; returns None on error.
    """
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


def process_folder(folder_path: str):
    """
    If no .txt present, convert eligible .docx transcripts to .txt.
    """
    try:
        files = os.listdir(folder_path)
    except Exception as e:
        print(f"Cannot list directory '{folder_path}': {e}")
        return

    # Skip if any .txt exists
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

                    # Remove the original docx after successful conversion
                    os.remove(docx_path)
                except Exception as e:
                    print(f"Error writing TXT '{txt_path}': {e}")
            else:
                print(f"Skipped non-transcript: {docx_path}")


def convert_srt_to_txt(folder_path: str):
    """
    Convert all .srt files in folder_path to .txt, then delete the original .srt files.
    """
    for filename in os.listdir(folder_path):
        if filename.lower().endswith('.srt'):
            srt_path = os.path.join(folder_path, filename)
            txt_filename = os.path.splitext(filename)[0] + '_srt' + '.txt'
            txt_path = os.path.join(folder_path, txt_filename)
            try:
                # Read SRT content
                with open(srt_path, 'r', encoding='utf-8', errors='ignore') as sf:
                    content = sf.read()
                # Write to TXT
                with open(txt_path, 'w', encoding='utf-8') as tf:
                    tf.write(content)
                # Delete original SRT
                os.remove(srt_path)
            except Exception as e:
                print(f"Error processing {srt_path}: {e}")


def rename_txt_files(folder_path: str):
    """
    Rename .txt files that do not already end with '_srt.txt' by appending '_srt' before the extension.
    """
    for filename in os.listdir(folder_path):
        # Only process .txt files
        if filename.lower().endswith('.txt') and not filename.lower().endswith('_srt.txt'):
            old_path = os.path.join(folder_path, filename)
            base_name = filename[:-4]  # remove '.txt'
            new_filename = f"{base_name}_srt.txt"
            new_path = os.path.join(folder_path, new_filename)
            try:
                os.rename(old_path, new_path)
            except Exception as e:
                print(f"Error renaming {old_path}: {e}")


def a02_01_docx2txt(BASE_DIR: str):
    # Directory to scan
    # BASE_DIR = r"C:\MyPython\ExamScore_AIClass\ExamFiles"

    for root, dirs, _ in os.walk(BASE_DIR):
        for d in dirs:
            process_folder(os.path.join(root, d))

            folder = os.path.join(root, d)
            convert_srt_to_txt(folder)

            rename_txt_files(folder)
