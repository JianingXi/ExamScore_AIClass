import os
import re
import requests
import time
import xml.etree.ElementTree as ET
from urllib.parse import quote_plus
# 如需使用 Google Scholar 检索，请先安装：pip install scholarly
from scholarly import scholarly

# 无需 API Key 的公开检索接口
CROSSREF_API      = "https://api.crossref.org/works"
SS_API            = "https://api.semanticscholar.org/graph/v1/paper/search"
OPENALEX_API      = "https://api.openalex.org/works"
ARXIV_API         = "http://export.arxiv.org/api/query"
CONF_API          = "https://dblp.org/search/publ/api"
PUBMED_SEARCH     = "https://eutils.ncbi.nlm.nih.gov/entrez/eutils/esearch.fcgi"
PUBMED_SUMMARY    = "https://eutils.ncbi.nlm.nih.gov/entrez/eutils/esummary.fcgi"
EUROPEPMC_API     = "https://www.ebi.ac.uk/europepmc/webservices/rest/search"
PAUSE             = 0.1  # 请求之间的等待时间（秒）


def format_citation(item, source):
    """按数据源统一格式化引用字符串"""
    if source == 'crossref':
        authors = item.get('author', [])
        author_str = ', '.join(
            f"{a.get('family')}{(' ' + a.get('given')) if a.get('given') else ''}"
            for a in authors[:3]
        ) + (' et al.' if len(authors) > 3 else '')
        title = item.get('title', [''])[0]
        venue = item.get('container-title', [''])[0]
        year = (item.get('published-print') or item.get('published-online') or {}) \
               .get('date-parts', [[None]])[0][0]
        return f"{author_str}. {title}. {venue}. {year}."

    if source == 'ss':
        authors = item.get('authors', [])
        author_str = ', '.join(a.get('name') for a in authors[:3]) \
                   + (' et al.' if len(authors) > 3 else '')
        title = item.get('title', '')
        venue = item.get('venue', '')
        year  = item.get('year', '')
        return f"{author_str}. {title}. {venue}. {year}."

    if source == 'openalex':
        auths = item.get('authorships', [])
        author_str = ', '.join(
            a.get('author', {}).get('display_name')
            for a in auths[:3]
        ) + (' et al.' if len(auths) > 3 else '')
        title = item.get('display_name', '')
        venue = item.get('host_venue', {}).get('display_name', '')
        year  = item.get('publication_year', '')
        return f"{author_str}. {title}. {venue}. {year}."

    if source == 'arxiv':
        authors = item.get('authors', [])
        author_str = ', '.join(authors[:3]) + (' et al.' if len(authors) > 3 else '')
        title      = item.get('title', '')
        year       = item.get('published', '').split('-')[0]
        return f"{author_str}. {title}. arXiv. {year}."

    if source == 'dblp':
        recs = item.get('result', {}).get('hits', {}).get('hit', [])
        if recs:
            r0 = recs[0]
            auths = r0.get('authors', [])
            author_str = ', '.join(
                a.get('author', {}).get('text') for a in auths[:3]
            ) + (' et al.' if len(auths) > 3 else '')
            title = r0.get('title', '')
            venue = r0.get('venue', '')
            year  = r0.get('year', '')
            return f"{author_str}. {title}. {venue}. {year}."

    if source == 'pubmed':
        authors = item.get('authors', [])
        author_str = ', '.join(authors[:3]) + (' et al.' if len(authors) > 3 else '')
        title = item.get('title', '')
        venue = item.get('fulljournalname', '')
        year  = item.get('pubdate', '')
        return f"{author_str}. {title}. {venue}. {year}."

    if source == 'europepmc':
        auths = item.get('authorList', {}).get('author', [])
        author_str = ', '.join(
            f"{a.get('firstName','')} {a.get('lastName','')}"
            for a in auths[:3]
        ) + (' et al.' if len(authors) > 3 else '')
        title = item.get('title', '')
        venue = item.get('journalTitle', '')
        year  = item.get('pubYear', '')
        return f"{author_str}. {title}. {venue}. {year}."

    if source == 'scholar':
        author_str = item.get('authors', '')
        title      = item.get('title', '')
        venue      = item.get('venue', '')
        year       = item.get('year', '')
        return f"{author_str}. {title}. {venue}. {year}."

    return ''


def check_crossref(ref):
    try:
        r = requests.get(CROSSREF_API,
                         params={'query.bibliographic': ref, 'rows': 1},
                         timeout=10)
        r.raise_for_status()
        items = r.json().get('message', {}).get('items', [])
        if items and items[0].get('author') and items[0].get('container-title'):
            return True, format_citation(items[0], 'crossref')
    except:
        pass
    return False, ''


def check_semantic_scholar(ref):
    try:
        r = requests.get(SS_API,
                         params={'query': ref, 'limit': 1},
                         timeout=10)
        r.raise_for_status()
        data = r.json().get('data', [])
        if data and data[0].get('authors') and data[0].get('venue'):
            return True, format_citation(data[0], 'ss')
    except:
        pass
    return False, ''


def check_openalex(ref):
    try:
        r = requests.get(
            f"{OPENALEX_API}?search={quote_plus(ref)}&per_page=1",
            timeout=10
        )
        r.raise_for_status()
        results = r.json().get('results', [])
        if results and results[0].get('authorships') and results[0].get('host_venue'):
            return True, format_citation(results[0], 'openalex')
    except:
        pass
    return False, ''


def check_arxiv(ref):
    # 优先提取 arXiv ID
    m = re.search(r'arXiv:(\d+\.\d+(v\d+)?)', ref)
    if m:
        arxiv_id = m.group(1)
        try:
            # 精准 id_list 检索
            r = requests.get(ARXIV_API,
                             params={'id_list': arxiv_id},
                             timeout=10)
            r.raise_for_status()
            ns = {'atom': 'http://www.w3.org/2005/Atom'}
            root = ET.fromstring(r.text)
            entry = root.find('atom:entry', ns)
            if entry is None:
                # fallback: search query by id
                r2 = requests.get(ARXIV_API,
                                  params={'search_query': f'id:{arxiv_id}',
                                          'max_results': 1},
                                  timeout=10)
                r2.raise_for_status()
                root2 = ET.fromstring(r2.text)
                entry = root2.find('atom:entry', ns)
                if entry is None:
                    return False, ''
            title = entry.find('atom:title', ns).text.strip().replace('\n', ' ')
            authors = [
                ae.find('atom:name', ns).text.strip()
                for ae in entry.findall('atom:author', ns)
            ]
            published = entry.find('atom:published', ns).text.strip()
            item = {'title': title, 'authors': authors, 'published': published}
            return True, format_citation(item, 'arxiv')
        except:
            return False, ''
    return False, ''


def check_dblp(ref):
    try:
        r = requests.get(CONF_API,
                         params={'q': ref, 'format': 'json'},
                         timeout=10)
        r.raise_for_status()
        hits = r.json().get('result', {}).get('hits', {}).get('hit', [])
        if hits:
            return True, format_citation({'result': {'hits': {'hit': hits}}}, 'dblp')
    except:
        pass
    return False, ''


def check_pubmed(ref):
    try:
        r = requests.get(PUBMED_SEARCH,
                         params={'db': 'pubmed', 'term': ref,
                                 'retmax': 1, 'retmode': 'json'},
                         timeout=10)
        r.raise_for_status()
        ids = r.json().get('esearchresult', {}).get('idlist', [])
        if ids:
            pmid = ids[0]
            r2 = requests.get(PUBMED_SUMMARY,
                              params={'db': 'pubmed', 'id': pmid,
                                      'retmode': 'json'},
                              timeout=10)
            info = r2.json().get('result', {}).get(pmid, {})
            authors = [a.get('name') for a in info.get('authors', [])]
            item = {
                'authors': authors,
                'title': info.get('title', ''),
                'fulljournalname': info.get('fulljournalname', ''),
                'pubdate': info.get('pubdate', '')
            }
            return True, format_citation(item, 'pubmed')
    except:
        pass
    return False, ''


def check_europepmc(ref):
    try:
        r = requests.get(EUROPEPMC_API,
                         params={'query': ref, 'format': 'json', 'pageSize': 1},
                         timeout=10)
        r.raise_for_status()
        res = r.json().get('resultList', {}).get('result', [])
        if res:
            return True, format_citation(res[0], 'europepmc')
    except:
        pass
    return False, ''


def check_scholar(ref):
    try:
        search = scholarly.search_pubs(ref)
        pub    = next(search)
        bib    = pub.bib
        item   = {
            'authors': bib.get('author', ''),
            'title':   bib.get('title', ''),
            'venue':   bib.get('journal', '') or bib.get('conference', ''),
            'year':    bib.get('year', '')
        }
        return True, format_citation(item, 'scholar')
    except:
        pass
    return False, ''


def process_folder(folder):
    section_file = os.path.join(folder, 'Section_参考文献.txt')
    if not os.path.isfile(section_file):
        return

    refs = []
    for line in open(section_file, encoding='utf-8'):
        text = line.strip()
        # 1. 移除所有空白和常见中英文标点
        norm = re.sub(r"[\s、，,。.]+", "", text)
        # 2. 去掉中文数字 + 标点前缀
        norm = re.sub(r"^[一二三四五六七八九十]+[、\.．]?", "", norm)
        # 3. 若规范化后包含“参考文献”，则跳过
        if not text or '参考文献' in norm:
            continue
        refs.append(text)

    found, errors = [], []
    for ref in refs:
        ok, cit = check_crossref(ref)
        if not ok: time.sleep(PAUSE); ok, cit = check_semantic_scholar(ref)
        if not ok: time.sleep(PAUSE); ok, cit = check_openalex(ref)
        if not ok: time.sleep(PAUSE); ok, cit = check_arxiv(ref)
        if not ok: time.sleep(PAUSE); ok, cit = check_dblp(ref)
        if not ok: time.sleep(PAUSE); ok, cit = check_pubmed(ref)
        if not ok: time.sleep(PAUSE); ok, cit = check_europepmc(ref)
        if not ok: time.sleep(PAUSE); ok, cit = check_scholar(ref)

        if ok and cit:
            found.append(cit)
        else:
            errors.append(ref)
        time.sleep(PAUSE)

    with open(os.path.join(folder, 'FoundRefs.txt'), 'w', encoding='utf-8') as ff:
        for c in found:
            ff.write(c + '\n')
    with open(os.path.join(folder, 'ErrorRef.txt'), 'w', encoding='utf-8') as ef:
        for e in errors:
            ef.write(e + '\n')

    print(f"Processed {os.path.basename(folder)}: {len(found)} successes, {len(errors)} failures.")


def a04_03_ref_check(BASE_DIR: str):
    # 根目录，包含各学生子文件夹
    # BASE_DIR = r"C:\MyPython\ExamScore_AIClass\ExamFiles"

    for entry in os.listdir(BASE_DIR):
        folder = os.path.join(BASE_DIR, entry)
        if os.path.isdir(folder):
            process_folder(folder)
    print("Done—请检查各子文件夹下的 FoundRefs.txt 与 ErrorRef.txt。")

