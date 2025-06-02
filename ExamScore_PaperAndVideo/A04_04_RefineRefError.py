import os

def filter_error_refs(base_dir):
    """
    遍历 base_dir 下的所有子文件夹，读取子文件夹中的 ErrorRef.txt，
    删除包含政府类关键词（中英文）的引用条目，并写回文件。
    """
    gov_keywords = [
        # —— 中文关键词 —— #
        "国务院", "中央", "国家", "政府",
        "财政部", "教育部", "人力资源和社会保障部", "自然资源部", "生态环境部",
        "住房和城乡建设部", "商务部", "工业和信息化部", "交通运输部",
        "文化和旅游部", "卫生健康委员会", "民政部", "司法部", "农业农村部",
        "科技部", "外交部", "国土资源部", "发改委", "国资委", "人民银行",
        "证监会", "银保监会", "保监会", "海关总署", "税务总局", "审计署",
        "国务院令", "国务院通知", "意见", "规划", "法规", "法令", "条例",
        "总理",

        # —— 英文关键词 —— #
        "government", "ministry", "department", "agency", "commission",
        "council", "administration", "authority", "federal", "state",
        "provincial", "municipal", "parliament", "senate", "congress",
        "house of representatives", "white house", "prime minister",
        "president", "cabinet", "directive", "regulation", "law", "act",
        "decree", "ordinance", "statute", "bill", "resolution",

        # —— 常见国家/地区及其机构简称 —— #
        "us government", "united states", "usa", "u.s.", "u.s.a.",
        "white house", "department of state", "department of education",
        "eu", "european union", "european commission",
        "british government", "uk government", "united kingdom",
        "germany", "france", "japan", "australia", "canada", "india",
        "south korea", "republic of korea", "russia", "french government",
        "german government", "japanese government", "australian government",
        "canadian government", "indian government",

        # —— 其他国际或区域组织 —— #
        "united nations", "un", "world bank", "imf", "who", "world health organization"
    ]

    # 大小写无关匹配
    gov_keywords = [kw.lower() for kw in gov_keywords]

    for name in os.listdir(base_dir):
        folder = os.path.join(base_dir, name)
        if not os.path.isdir(folder):
            continue

        errorref_path = os.path.join(folder, "ErrorRef.txt")
        if not os.path.isfile(errorref_path):
            continue

        with open(errorref_path, 'r', encoding='utf-8') as f:
            lines = f.readlines()

        filtered = []
        for line in lines:
            text = line.lower()
            if any(kw in text for kw in gov_keywords):
                continue
            filtered.append(line)

        # 覆盖写回
        with open(errorref_path, 'w', encoding='utf-8') as f:
            f.writelines(filtered)

        print(f"Processed {errorref_path}: kept {len(filtered)}/{len(lines)} lines.")


def a04_04_refine_ref_error(base_directory: str):
    # base_directory = r"C:\MyPython\ExamScore_AIClass\ExamFiles"
    filter_error_refs(base_directory)
