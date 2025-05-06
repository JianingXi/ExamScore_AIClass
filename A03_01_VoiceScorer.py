import os
import re
import math
from docx import Document

# 文件后缀
SRT_SUFFIX = '_srt.txt'
MP4_SUFFIX = '_mp4.docx'

# 正则模式，用于提取时长信息
duration_pattern = re.compile(r"时长[:：]\s*([\d\.]+)\s*秒")
time_pattern = re.compile(r"(\d{2}):(\d{2}):(\d{2}),\d{3}\s*-->\s*(\d{2}):(\d{2}):(\d{2}),\d{3}")

# 解析时长：支持.docx和.txt，优先标注，其次提取字幕最大时间戳
def parse_duration(path):
    try:
        ext = path.lower().split('.')[-1]
        text = ''
        if ext == 'docx':
            doc = Document(path)
            text = '\n'.join(p.text for p in doc.paragraphs)
        else:
            with open(path, 'r', encoding='utf-8', errors='ignore') as f:
                text = f.read()
        m = duration_pattern.search(text)
        if m:
            return float(m.group(1))
        if ext == 'txt':
            times = [int(h)*3600 + int(mi)*60 + int(s)
                     for (h,mi,s,_,_,_) in time_pattern.findall(text)]
            if times:
                return max(times)
    except:
        pass
    return 0.0

# 分析字幕：统计字数与>5秒停顿次数
def analyze_srt(path):
    words = pauses = 0
    timestamps = []
    with open(path, 'r', encoding='utf-8', errors='ignore') as f:
        for line in f:
            ln = line.strip()
            if not ln:
                continue
            m = time_pattern.match(ln)
            if m:
                start = int(m.group(1))*3600 + int(m.group(2))*60 + int(m.group(3))
                end   = int(m.group(4))*3600 + int(m.group(5))*60 + int(m.group(6))
                timestamps.append((start, end))
            elif not ln.isdigit():
                words += len(ln.split())
    for i in range(1, len(timestamps)):
        if timestamps[i][0] - timestamps[i-1][1] > 5:
            pauses += 1
    return words, pauses

# 评分函数
def evaluate(folder):
    files = os.listdir(folder)
    has_srt = any(f.endswith(SRT_SUFFIX) for f in files)
    has_mp4 = any(f.endswith(MP4_SUFFIX) for f in files)
    missing = (not has_srt) + (not has_mp4)

    # 缺两项材料
    if missing == 2:
        score = 0
        comments = ['未提交字幕和视频文件，成绩0分。']
    # 缺一项材料，固定2分，无其他扣分
    elif missing == 1:
        score = 2
        if has_srt:
            comments = ['仅有字幕，未提供视频文件，基础2分。']
        else:
            comments = ['仅有视频文件，未提供字幕，基础2分。']
    # 材料齐全，综合评估5-10分
    else:
        score = 10
        comments = []
        mp4_file = next(f for f in files if f.endswith(MP4_SUFFIX))
        srt_file = next(f for f in files if f.endswith(SRT_SUFFIX))
        duration = parse_duration(os.path.join(folder, mp4_file))
        comments.append(f'演讲时长：{duration:.1f}秒，符合要求。')
        # 时长不足9分钟扣分
        if duration < 540:
            mins_short = math.ceil((540 - duration) / 60)
            deduction = mins_short * 0.5
            score -= deduction
            comments.append(f'时长不足9分钟，短{mins_short}分钟，扣{deduction:.1f}分。')
        # 语速评估
        words, pauses = analyze_srt(os.path.join(folder, srt_file))
        wpm = words / (duration/60) if duration > 0 else 0
        if 100 <= wpm <= 140:
            comments.append(f'语速{wpm:.1f}字/分钟，流畅自然。')
        elif 80 <= wpm < 100 or 140 < wpm <= 160:
            score -= 0.5
            comments.append(f'语速{wpm:.1f}字/分钟，略有波动，扣0.5分。')
        else:
            score -= 1
            comments.append(f'语速{wpm:.1f}字/分钟，与理想范围差距较大，扣1分。')
        # 停顿评估
        if pauses > 0:
            d = min(pauses, 2)
            score -= d * 0.5
            comments.append(f'含{pauses}次明显停顿，共扣{d*0.5}分。')
        else:
            comments.append('演讲连贯，无明显停顿。')
        # 保持最低5分
        if score < 5:
            comments.append('考虑整体表现，最低保留5分。')
            score = 5

    # 写入结果
    with open(os.path.join(folder, 'AudioScore.txt'), 'w', encoding='utf-8') as f:
        f.write(f'{score}/10\n')
        for c in comments:
            f.write(c + '\n')
    print(f'{folder} 得分：{score}/10')



import os

def delete_srt_txt_files(root_dir):
    """
    遍历 root_dir 下的每个子文件夹，删除其中所有以 '_srt.txt' 结尾的文件。
    """
    if not os.path.isdir(root_dir):
        print(f"指定路径不存在或不是文件夹：{root_dir}")
        return

    # 遍历 root_dir 下的每个条目
    for entry in os.listdir(root_dir):
        sub_path = os.path.join(root_dir, entry)
        # 只处理子文件夹
        if os.path.isdir(sub_path):
            # 遍历该子文件夹下的所有文件
            for fname in os.listdir(sub_path):
                if fname.endswith('_srt.txt'):
                    file_path = os.path.join(sub_path, fname)
                    try:
                        os.remove(file_path)
                    except Exception as e:
                        print(f"  删除失败：{file_path}，错误：{e}")


def a03_01_voice_scorer(BASE_DIR: str):
    # BASE_DIR = r"C:\MyPython\ExamScore_AIClass\ExamFiles"

    for root, dirs, _ in os.walk(BASE_DIR):
        for d in dirs:
            evaluate(os.path.join(root, d))

    delete_srt_txt_files(BASE_DIR)