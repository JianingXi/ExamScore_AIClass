from ExamScore_PaperAndVideo.A01_01_UnzipFiles import a01_01_unzip_files
from ExamScore_PaperAndVideo.A02_01_Docx2Txt import a02_01_docx2txt
from ExamScore_PaperAndVideo.A02_11_VedioSummary import a02_11_video_summary
from ExamScore_PaperAndVideo.A03_01_VoiceScorer import a03_01_voice_scorer
from ExamScore_PaperAndVideo.A03_02_PPTScorer import a03_02_ppt_scorer
from ExamScore_PaperAndVideo.A03_02_01_PPT2LongImage import batch_pptx_to_long_images
from ExamScore_PaperAndVideo.A03_03_MP4Scorer import a03_03_mp4_scorer
from ExamScore_PaperAndVideo.A03_03_01_MP42LongImage import batch_mp4_to_long_images
from ExamScore_PaperAndVideo.A04_01_FormatScorer import a04_01_format_scorer
from ExamScore_PaperAndVideo.A04_02_ReportMain import a04_02_report_main
from ExamScore_PaperAndVideo.A04_03_RefCheck import a04_03_ref_check
from ExamScore_PaperAndVideo.A04_04_RefineRefError import a04_04_refine_ref_error
from ExamScore_PaperAndVideo.A04_05_ReportRefError import a04_05_report_ref_error
from ExamScore_PaperAndVideo.A04_06_ContentTF_IDF import a04_06_content_tf_idf
from ExamScore_PaperAndVideo.A04_07_ContentEmbOutlier import a04_07_content_emb_outlier

from ExamScore_PaperAndVideo.A05_01_Concatenate import a05_01_concatenate
from ExamScore_PaperAndVideo.A05_02_RemoveDataFromComments import a05_02_remove_data_from_comments
from ExamScore_PaperAndVideo.A05_03_FormToScorerTable import a05_03_form_to_score_table

from ExamScore_PaperAndVideo.A06_01_Score2ChaoxingTable import a06_01_score2chaoxing_table

from ExamScore_PaperAndVideo.B01_batch_deepseek_folder_scorer import batch_score_to_excel

root_dir = r"C:\MyDocument\ToDoList\D20_DoingPlatform\D20251228_作品汇总与评分表_广州医科大学第三届研究生创新论坛\赛道1：科研成果交流与合作_广州医科大学第三届研究生创新论坛"

include_video_subscript = 1
include_formal_paper = 1

start_up = 1


if start_up == 1:
    a01_01_unzip_files(root_dir)

    if include_video_subscript:
        # 如有视频汇报
        a02_01_docx2txt(root_dir)

        batch_mp4_to_long_images(root_dir)
        # a02_11_video_summary(root_dir)

        batch_pptx_to_long_images(root_dir)
        # a03_02_ppt_scorer(root_dir)

        batch_score_to_excel(root_dir)
