from A01_01_UnzipFiles import a01_01_unzip_files
from A02_01_Docx2Txt import a02_01_docx2txt
from A02_11_VedioSummary import a02_11_video_summary
from A03_01_VoiceScorer import a03_01_voice_scorer
from A03_02_PPTScorer import a03_02_ppt_scorer
from A03_03_MP4Scorer import a03_03_mp4_scorer
from A04_01_FormatScorer import a04_01_format_scorer
from A04_02_ReportMain import a04_02_report_main
from A04_03_RefCheck import a04_03_ref_check
from A04_04_RefineRefError import a04_04_refine_ref_error
from A04_05_ReportRefError import a04_05_report_ref_error
from A04_06_ContentTF_IDF import a04_06_content_tf_idf
from A04_07_ContentEmbOutlier import a04_07_content_emb_outlier

from A05_01_Concatenate import a05_01_concatenate
from A05_02_RemoveDataFromComments import a05_02_remove_data_from_comments
from A05_03_FormToScorerTable import a05_03_form_to_score_table

from A06_01_Score2ChaoxingTable import a06_01_score2chaoxing_table

root_dir = r"C:\MyDocument\ToDoList\D20_DoingPlatform\D20_人工智能与大数据\临床药学-小组报告-课程综述与答辩视频"
chaoxing_format_file = r'\小组报告-课程综述与答辩视频.xls'

include_video_subscript = 1
include_formal_paper = 1

a01_01_unzip_files(root_dir)

if include_video_subscript:
    # 如有视频汇报
    a02_01_docx2txt(root_dir)
    a02_11_video_summary(root_dir)
    a03_01_voice_scorer(root_dir)
    a03_02_ppt_scorer(root_dir)
    a03_03_mp4_scorer(root_dir)

if include_formal_paper:
    # 如有课程论文
    a04_01_format_scorer(root_dir)
    a04_02_report_main(root_dir)

    a04_03_ref_check(root_dir)  # 慢
    
    a04_04_refine_ref_error(root_dir)
    a04_05_report_ref_error(root_dir)

    a04_06_content_tf_idf(root_dir)
    a04_07_content_emb_outlier(root_dir)

    con_file = root_dir + r"\MergedResult.xlsx"
    detailed_file = root_dir + r"\MergedResult_Descriptions.xlsx"
    score_1_file = root_dir + r"\成绩表_明细版.xlsx"

    a05_01_concatenate(root_dir, con_file)
    a05_02_remove_data_from_comments(con_file, detailed_file)
    a05_03_form_to_score_table(detailed_file, score_1_file)

    final_xls_file = root_dir + chaoxing_format_file
    score_f_file = root_dir + r"\整理后的成绩表.xlsx"
    a06_01_score2chaoxing_table(score_1_file, final_xls_file, score_f_file)
