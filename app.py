import os
from openai import OpenAI
from docx import Document
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
import io
from PyPDF2 import PdfReader
import streamlit as st
# 唯一新增：本地PDF转Word库
from pdf2docx import Converter

# ==============================================================================
# 🌐 【多语言配置区】中英文语言包，完全复原你原版完整内容 + 修复key错误
# ==============================================================================
LANG_PACK = {
    "zh": {
        # 全局通用
        "page_title": "小倍AI助手",
        "main_title": "🙂 小倍AI助手",
        "warning": "警告",
        "error": "错误",
        "success": "完成",
        "start": "开始",
        "download": "下载",
        "upload": "上传",
        "preview": "预览",
        "generating": "生成中...",
        "processing": "处理中...",
        # 侧边栏
        "sidebar_title": "功能导航",
        "select_func": "选择功能",
        "sidebar_footer": "✅修复了一些已知问题ㅤㅤㅤ☺️版本：测试版 ㅤㅤㅤ©️beelicn.com",
        "lang_select": "选择语言",
        # 功能菜单（label+副标题）
        "menu_search_label": "💻 全网报告搜索",
        "menu_search_sub": "🤔 小倍正在搜索报告中",
        "menu_summary_label": "💡 文档总结/数据提取",
        "menu_summary_sub": "🤨 小倍正在分析你的文档",
        "menu_generate_label": "📝 行业报告生成",
        "menu_generate_sub": "🤯 小倍正在生成你的报告",
        "menu_compare_label": "📈 多文档竞品/赛道对比分析",
        "menu_compare_sub": "😁 小倍正在对比中",
        "menu_rewrite_label": "✏️ 仿照模板改写文档",
        "menu_rewrite_sub": "🧐 小倍正在改写你的文档",
        "menu_translate_label": "🌐 商务文档翻译",
        "menu_translate_sub": "😏 小倍翻译中",
        "menu_pdf2word_label": "💾 无损PDF转Word",
        "menu_pdf2word_sub": "😎 小倍PDF格式转换助手",
        # 赛道配置
        "track_general": "通用全行业",
        "track_ai": "AI市场研究",
        "track_consulting": "战略咨询",
        "track_risk": "企业风险管理",
        "track_manufacture": "制造业出海欧洲市场",
        # 功能1：全网报告搜索
        "search_input_tip": "请输入行业/赛道关键词，越详细检索结果越精准",
        "search_btn": "开始检索",
        "search_loading": "正在检索合规行业报告...",
        "search_kw_empty": "请输入检索关键词",
        "search_pub_org": "🏢 发布机构：",
        "search_pub_year": "📅 发布年份：",
        "search_abstract": "📄 核心摘要：",
        # 功能2：文档总结/数据提取
        "summary_mode": "选择分析模式",
        "summary_mode_general": "通用文档总结",
        "summary_mode_indicator": "行研核心指标提取",
        "summary_upload_tipip": "上传TXT/DOCX格式的文档、财报、行业白皮书",
        "summary_upload_tip": "上传TXT/DOCX格式的文档、财报、行业白皮书",
        "summary_analyze_btn": "开始分析",
        "summary_analyze_loading": "正在执行{mode}...",
        "summary_original_preview": "文档原文预览",
        "summary_result_title": "✅ {mode}结果",
        "summary_download_btn": "📍 下载分析结果Word文档",
        "summary_download_filename": "文档{mode}结果.docx",
        # 功能3：行业报告生成
        "generate_track_select": "选择垂直赛道模板",
        "generate_name_input": "输入目标行业/赛道/产品名称",
        "generate_ref_tip": "【可选】上传自有参考资料/报告模板（生成内容优先匹配参考资料的格式与规范）",
        "generate_ref_upload": "上传参考资料TXT/DOCX文档",
        "generate_ref_preview": "预览参考资料内容",
        "generate_btn": "📝 生成行业报告",
        "generate_loading": "正在生成{track}赛道专属行业报告...",
        "generate_name_empty": "请输入目标行业/赛道名称",
        "generate_ref_rule": "【参考资料要求】生成内容必须优先参考以下资料的格式规范、行业定义、数据口径：",
        "generate_report_title": "✉️ {name} | {track}赛道行业报告",
        "generate_download_word": "📍 下载Word版报告",
        "generate_download_ppt": "📊 下载咨询标准PPT版",
        "generate_word_filename": "{name}_{track}_行业报告.docx",
        "generate_ppt_filename": "{name}_{track}_行业报告.pptx",
        # 功能4：多文档对比分析
        "compare_tip": "支持上传2-5份同赛道行业报告、竞品财报、行业白皮书，自动生成战略咨询级对比分析报告",
        "compare_upload_tip": "上传需要对比分析的TXT/DOCX文档",
        "compare_btn": "📈 生成对比分析报告",
        "compare_loading": "正在解析文档并生成对比分析报告...",
        "compare_file_min": "请至少上传2份文档进行对比分析",
        "compare_result_title": "✅ 对比分析报告",
        "compare_download_word": "📍 下载Word版分析报告",
        "compare_download_ppt": "📊 下载咨询标准PPT版",
        "compare_word_filename": "赛道竞品对比分析报告.docx",
        "compare_ppt_filename": "赛道竞品对比分析报告.pptx",
        # 功能5：仿照模板改写文档
        "rewrite_flow": "流程：上传模板文档 → 上传待改写文档 → 一键改写 → 在线预览 → 下载双格式文件",
        "rewrite_template_upload": "1. 上传模板文档",
        "rewrite_content_upload": "2. 上传待改写文档",
        "rewrite_template_preview": "预览模板内容",
        "rewrite_content_preview": "预览待改写内容",
        "rewrite_btn": "✏️ 开始改写",
        "rewrite_loading": "正在按模板风格改写文档...",
        "rewrite_file_empty": "请先上传模板文档和待改写文档",
        "rewrite_result_title": "✅ 改写结果",
        "rewrite_download_word": "📍 下载Word文档",
        "rewrite_download_ppt": "📊 下载PPT版",
        "rewrite_word_filename": "文档改写结果.docx",
        "rewrite_ppt_filename": "文档改写结果.pptx",
        # 功能6：商务文档翻译
        "translate_tip": "支持直接输入文本翻译，或上传TXT/DOCX文档批量翻译，适配商务/咨询正式文档场景",
        "translate_target_lang": "选择目标翻译语言",
        "translate_mode": "翻译模式",
        "translate_mode_text": "直接输入文本",
        "translate_mode_file": "上传文档翻译",
        "translate_textarea_tip": "请输入需要翻译的商务文档内容",
        "translate_upload_tip": "上传需要翻译的TXT/DOCX文档",
        "translate_original_preview": "预览原文内容",
        "translate_btn": "🌐 开始翻译",
        "translate_loading": "正在翻译中，请稍等...",
        "translate_content_empty": "请输入需要翻译的内容，或上传有效文档",
        "translate_result_title": "✅ 翻译结果",
        "translate_download_btn": "📍 下载翻译结果Word文档",
        "translate_download_filename": "商务文档翻译结果.docx",
        # 功能7：PDF转Word
        "pdf2word_tip": "上传PDF → AI智能修复乱换行/乱分段 → 还原整洁排版 → 预览下载双格式文件",
        "pdf2word_upload_tip": "上传PDF文件",
        "pdf2word_loading": "正在提取PDF内容，并AI智能规整排版...",
        "pdf2word_preview_title": "📋 AI规整后内容预览",
        "pdf2word_download_word": "📍 下载无损Word文档",
        "pdf2word_download_ppt": "📊 下载PPT版",
        "pdf2word_word_filename": "PDF转换结果.docx",
        "pdf2word_ppt_filename": "PDF转换结果.pptx",
        # 其他
        "func_not_found": "该功能暂未实现，请检查配置",
        "ppt_title_default": "咨询报告",
        "ppt_footer": "小倍咨询级AI报告助手\n合规生成 | 数据可溯源",
        "ppt_end_page": "报告结束",
    },
    "en": {
        # Global General
        "page_title": "Xiaobei AI Assistant",
        "main_title": "🙂 Xiaobei AI Assistant",
        "warning": "Warning",
        "error": "Error",
        "success": "Success",
        "start": "Start",
        "download": "Download",
        "upload": "Upload",
        "preview": "Preview",
        "generating": "Generating...",
        "processing": "Processing...",
        # Sidebar
        "sidebar_title": "Function Navigation",
        "select_func": "Select Function",
        "sidebar_footer": "✅ Fixed known issuesㅤㅤㅤㅤ☺️Version: Beta ㅤㅤ©️beelicn.com",
        "lang_select": "Select Language",
        # Menu Config
        "menu_search_label": "💻 Full-web Report Search",
        "menu_search_sub": "🤔 Xiaobei is searching reports",
        "menu_summary_label": "💡 Doc Sum/Data Extraction",
        "menu_summary_sub": "🤨 Xiaobei is analyzing your document",
        "menu_generate_label": "📝 Industry Report Generation",
        "menu_generate_sub": "🤯 Xiaobei is generating your report",
        "menu_compare_label": "📈  Competitor Analysis",
        "menu_compare_sub": "😁 Xiaobei is comparing documents",
        "menu_rewrite_label": "✏️ Template-based Rewrite",
        "menu_rewrite_sub": "🧐 Xiaobei is rewriting your document",
        "menu_translate_label": "🌐 Business Translation",
        "menu_translate_sub": "😏 Xiaobei is translating",
        "menu_pdf2word_label": "💾 Lossless PDF to Word",
        "menu_pdf2word_sub": "😎 Xiaobei PDF Converter",
        # Industry Tracks
        "track_general": "General Industry",
        "track_ai": "AI Market Research",
        "track_consulting": "Strategy Consulting",
        "track_risk": "Enterprise Risk Management",
        "track_manufacture": "Manufacturing EU Go-to-Market",
        # Function 1: Report Search
        "search_input_tip": "Enter industry/track keywords, more details bring more accurate results",
        "search_btn": "Start Search",
        "search_loading": "Searching compliant industry reports...",
        "search_kw_empty": "Please enter search keywords",
        "search_pub_org": "🏢 Publisher: ",
        "search_pub_year": "📅 Publish Year: ",
        "search_abstract": "📄 Abstract: ",
        # Function 2: Doc Summary
        "summary_mode": "Select Analysis Mode",
        "summary_mode_general": "General Document Summary",
        "summary_mode_indicator": "Industry Research Indicator Extraction",
        "summary_upload_tip": "Upload TXT/DOCX document, financial report, white paper",
        "summary_analyze_btn": "Start Analysis",
        "summary_analyze_loading": "Executing {mode}...",
        "summary_original_preview": "Original Document Preview",
        "summary_result_title": "✅ {mode} Result",
        "summary_download_btn": "📍 Download Word Result",
        "summary_download_filename": "Document_{mode}_Result.docx",
        # Function 3: Report Generation
        "generate_track_select": "Select Vertical Track Template",
        "generate_name_input": "Enter target industry/track/product name",
        "generate_ref_tip": "【Optional】Upload reference materials (generated content matches format first)",
        "generate_ref_upload": "Upload reference TXT/DOCX document",
        "generate_ref_preview": "Preview Reference Content",
        "generate_btn": "📝 Generate Consulting Report",
        "generate_loading": "Generating report for {track} track...",
        "generate_name_empty": "Please enter target industry/track name",
        "generate_ref_rule": "【Reference Rule】Generated content must prioritize the format from reference below:",
        "generate_report_title": "✉️ {name} | {track} Track Report",
        "generate_download_word": "📍 Download Word Report",
        "generate_download_ppt": "📊 Download Consulting PPT",
        "generate_word_filename": "{name}_{track}_Industry_Report.docx",
        "generate_ppt_filename": "{name}_{track}_Industry_Report.pptx",
        # Function 4: Multi-doc Compare
        "compare_tip": "Support 2-5 documents of the same track to generate strategic consulting comparative analysis report",
        "compare_upload_tip": "Upload TXT/DOCX documents for comparison",
        "compare_btn": "📈 Generate Comparative Analysis",
        "compare_loading": "Parsing documents and generating report...",
        "compare_file_min": "Please upload at least 2 documents for comparison",
        "compare_result_title": "✅ Comparative Analysis Report",
        "compare_download_word": "📍 Download Word Report",
        "compare_download_ppt": "📊 Download Consulting PPT",
        "compare_word_filename": "Track_Competitor_Analysis_Report.docx",
        "compare_ppt_filename": "Track_Competitor_Analysis_Report.pptx",
        # Function 5: Template Rewrite
        "rewrite_flow": "Flow: Upload Template → Upload Target Document → One-click Rewrite → Preview → Download",
        "rewrite_template_upload": "1. Upload Template Document",
        "rewrite_content_upload": "2. Upload Target Document",
        "rewrite_template_preview": "Preview Template Content",
        "rewrite_content_preview": "Preview Target Content",
        "rewrite_btn": "✏️ Start Rewrite",
        "rewrite_loading": "Rewriting document with template style...",
        "rewrite_file_empty": "Please upload both template and target document first",
        "rewrite_result_title": "✅ Rewrite Result",
        "rewrite_download_word": "📍 Download Word Document",
        "rewrite_download_ppt": "📊 Download PPT Version",
        "rewrite_word_filename": "Document_Rewrite_Result.docx",
        "rewrite_ppt_filename": "Document_Rewrite_Result.pptx",
        # Function 6: Translation
        "translate_tip": "Support direct text translation or batch translation via TXT/DOCX upload",
        "translate_target_lang": "Select Target Language",
        "translate_mode": "Translation Mode",
        "translate_mode_text": "Direct Text Input",
        "translate_mode_file": "Upload Document",
        "translate_textarea_tip": "Enter business document content to translate",
        "translate_upload_tip": "Upload TXT/DOCX document to translate",
        "translate_original_preview": "Preview Original Content",
        "translate_btn": "🌐 Start Translation",
        "translate_loading": "Translating, please wait...",
        "translate_content_empty": "Please enter content or upload a valid document",
        "translate_result_title": "✅ Translation Result",
        "translate_download_btn": "📍 Download Word Result",
        "translate_download_filename": "Business_Document_Translation.docx",
        # Function 7: PDF to Word
        "pdf2word_tip": "Upload PDF → AI fix line breaks → Restore neat layout → Preview & Download",
        "pdf2word_upload_tip": "Upload PDF File",
        "pdf2word_loading": "Extracting PDF content and formatting with AI...",
        "pdf2word_preview_title": "📋 AI Formatted Content Preview",
        "pdf2word_download_word": "📍 Download Lossless Word",
        "pdf2word_download_ppt": "📊 Download PPT Version",
        "pdf2word_word_filename": "PDF_Conversion_Result.docx",
        "pdf2word_ppt_filename": "PDF_Conversion_Result.pptx",
        # Others
        "func_not_found": "This function is not available yet",
        "ppt_title_default": "Consulting Report",
        "ppt_footer": "Xiaobei Consulting AI Assistant\nCompliant Generation | Traceable Data",
        "ppt_end_page": "End of Report",
    }
}

# 翻译目标语言选项（中英文适配）
TARGET_LANG_OPTIONS = {
    "zh": ["简体中文", "English", "日本語", "한국어", "繁体中文"],
    "en": ["Simplified Chinese", "English", "Japanese", "Korean", "Traditional Chinese"]
}

# ==============================================================================
# 🎯 【用户核心配置区】完全复原你原版
# ==============================================================================
LOCAL_CONFIG = {
    "base_url": "https://ark.cn-beijing.volces.com/api/v3",
    "api_key": "ark-fc3c7e9f-d50d-48f5-8698-4955a37db662-5b27a",
    "model_name": "doubao-seed-2-0-pro-260215"
}

PROMPT_CONFIG = {
    "compliance_rule": """
    【强制合规要求，必须严格遵守】
    1. 所有数据、市场规模、增速、市场份额等量化内容，必须标注权威数据来源，包括但不限于：欧睿、IDC、乘联会、国家统计局、行业协会、上市公司财报、海关总署、贝恩/麦肯锡/波士顿咨询等权威机构发布的报告
    2. 绝对禁止虚构、编造任何数据、机构、事件、案例，所有内容必须符合行业真实情况
    3. 所有观点必须有对应的事实和数据支撑，禁止无依据的主观判断
    4. 严格遵循咨询行业报告的专业规范、结构逻辑和专业术语，语言正式、严谨、客观
    """,
    "industry_report_general": """
    为【{name}】生成专业、合规的咨询级行业报告，必须严格遵守以下要求：
    1. 报告结构必须包含7个核心部分：①行业定义与分类 ②市场规模与增长趋势 ③产业链上下游分析 ④竞争格局与核心玩家 ⑤用户画像与需求分析 ⑥行业痛点与发展趋势 ⑦投资机会与风险建议
    2. {compliance_rule}
    3. 结构清晰，段落分明，标题层级明确，符合正式咨询报告的排版规范
    """,
    "industry_report_ai": """
    为【{name}】生成AI领域专业市场研究报告，严格遵守AI行业研究规范，必须包含：
    1. 核心结构：①赛道定义与技术路径 ②市场规模与投融资情况 ③技术成熟度与落地场景 ④核心厂商与竞争格局 ⑤政策监管环境 ⑥技术趋势与商业化痛点 ⑦市场机会与风险提示
    2. {compliance_rule}
    3. 重点突出AI技术落地的商业价值、市场竞争壁垒、客户付费意愿，符合一级市场AI赛道研究的专业规范
    """,
    "industry_report_consulting": """
    为【{name}】生成战略咨询级行业研究报告，严格遵循顶级咨询公司报告规范，必须包含：
    1. 核心结构：①行业宏观环境（PEST分析）②市场规模与增长预测 ③产业链价值分布与利润池分析 ④竞争格局与五力模型分析 ⑤标杆企业战略与商业模式拆解 ⑥行业关键成功要素 ⑦企业进入战略与增长路径建议
    2. {compliance_rule}
    3. 重点突出战略洞察、可落地的商业建议，符合战略咨询项目的交付标准，逻辑严谨，洞察深刻
    """,
    "industry_report_risk": """
    为【{name}】生成企业风险管理视角的行业分析报告，严格遵循企业全面风险管理规范，必须包含：
    1. 核心结构：①行业基本情况与经营环境 ②行业核心风险点识别（市场风险、信用风险、运营风险、合规风险、政策风险）③风险传导路径分析 ④行业标杆企业风险管理实践 ⑤风险应对策略与缓释措施 ⑥行业风险预警指标体系
    2. {compliance_rule}
    3. 重点突出风险的量化分析、发生概率与影响程度，符合企业内控与风险管理的专业要求
    """,
    "industry_report_manufacture": """
    为【{name}】生成制造业出海欧洲市场的专项分析报告，严格遵守跨境贸易与出海咨询规范，必须包含：
    1. 核心结构：①欧洲目标市场准入政策与合规要求 ②市场规模与消费需求特征 ③欧洲本地竞争格局 ④供应链与物流方案分析 ⑤关税与税务筹划要点 ⑥本地化运营策略 ⑦出海风险与应对建议
    2. {compliance_rule}
    3. 重点突出欧洲市场合规要求、本地化运营难点、跨境供应链解决方案，符合制造业出海的真实业务需求
    """,
    "doc_summary_general": """
    对以下文档内容进行专业总结，核心输出4部分：1. 文档核心观点 2. 关键数据与信息 3. 行业竞争格局 4. 未来趋势与风险提示
    {compliance_rule}
    文档内容：{text}
    """,
    "doc_summary_indicator": """
    对以下财报/行业白皮书/行研报告内容，进行行研核心指标提取，严格遵守以下要求：
    1. 必须提取的核心指标：市场规模、年复合增长率、市场集中度CR5/CR10、行业平均毛利率、核心竞品市场份额、核心财务指标、政策关键节点
    2. 所有指标必须标注对应的来源、统计年份、统计口径
    3. 最终输出必须是**标准Markdown结构化表格**，表格列名：指标名称、指标数值、统计周期、数据来源、备注说明
    4. 禁止虚构任何指标，无明确数据的指标标注「文档未提及」即可
    5. 表格输出完成后，补充100字以内的核心指标洞察总结
    文档内容：{text}
    """,
    "report_search": """
    关键词：{keyword}，返回10条真实存在的行业报告，严格遵守格式要求：标题|机构|发布年份|核心摘要
    {compliance_rule}
    禁止输出链接、网址、虚构内容，每条报告必须真实可查
    """,
    "template_rewrite": """
    你是专业咨询文档改写助手，严格遵守要求：
    1. 完全模仿【模板文档】的文风、结构、段落格式、专业度、语气、标题层级
    2. 把【待改写内容】按照模板风格完整重写，不改变原文核心意思、核心数据、核心观点
    3. 优化语句的专业度、严谨性，符合咨询报告的写作规范
    4. {compliance_rule}
    5. 不要添加任何多余解释、装饰符号，只输出改写后的完整内容

    【模板文档】：
    {template_content}

    【待改写内容】：
    {original_content}
    """,
    "doc_translate": """
    你是专业商务文档翻译专家，严格遵守翻译要求：
    1. 目标语言：{target_lang}，严格按照目标语言进行专业翻译
    2. 精准翻译行业专业术语、商务表达、金融/咨询专业词汇，符合目标语言的正式商务文档规范
    3. 严格保留原文的段落结构、标题层级、表格格式，不改变原文核心意思、核心数据
    4. 翻译流畅自然，符合目标语言的商务写作习惯，无语法错误
    5. 不要添加任何额外解释、备注，只输出翻译后的完整内容

    需要翻译的原文：
    {text}
    """,
    "pdf_format": """
    你是专业文档排版整理助手，请对下面PDF提取的乱序文字做无损规整排版：
    要求：
    1. 严格保留原文所有内容、核心数据、观点，不删减、不修改原文意思
    2. 按照原文的逻辑结构重新分段、换行、区分标题和正文，还原标题层级
    3. 修复PDF自动拆行、断句、乱换行、乱码问题
    4. 排版整洁、段落清晰、格式规范，适合直接保存为Word/PPT
    5. 不要加多余解释、序号、装饰符号，只输出规整后的完整内容

    需要整理的PDF原文：
    {text}
    """,
    "multi_doc_compare": """
    你是专业战略咨询顾问，基于以下上传的多份同赛道行业报告/竞品财报/白皮书，生成专业的对比分析报告，严格遵守要求：
    1. 报告核心结构：①分析背景与对比范围 ②核心指标横向对比（市场规模、增速、盈利能力、市场份额等，输出结构化表格）③竞争格局与商业模式对比 ④核心优劣势差异分析 ⑤赛道机会与风险提示 ⑥战略建议
    2. {compliance_rule}
    3. 所有对比内容必须基于上传的文档内容，禁止添加文档外的虚构信息，重点突出核心差异与战略洞察
    4. 结构清晰，符合战略咨询项目对比分析报告的专业规范

    上传的文档内容合集：
    {all_doc_text}
    """
}

PAGE_CONFIG = {
    "page_icon": "😆"
}

# ==============================================================================
# 【禁止修改区】完全复原你原版初始化
# ==============================================================================
def init_session_state():
    if "language" not in st.session_state:
        st.session_state.language = "zh"
    if "selected_tab" not in st.session_state:
        st.session_state.selected_tab = ""
    if "rewrite_result" not in st.session_state:
        st.session_state.rewrite_result = ""
    if "rewrite_generating" not in st.session_state:
        st.session_state.rewrite_generating = False
    if "translate_result" not in st.session_state:
        st.session_state.translate_result = ""
    if "translate_generating" not in st.session_state:
        st.session_state.translate_generating = False
    if "compare_result" not in st.session_state:
        st.session_state.compare_result = ""
    if "compare_generating" not in st.session_state:
        st.session_state.compare_generating = False

init_session_state()
current_lang = st.session_state.language
lang = LANG_PACK[current_lang]

MENU_CONFIG = [
    {"id": "search", "label": lang["menu_search_label"], "sub_title": lang["menu_search_sub"]},
    {"id": "summary", "label": lang["menu_summary_label"], "sub_title": lang["menu_summary_sub"]},
    {"id": "generate", "label": lang["menu_generate_label"], "sub_title": lang["menu_generate_sub"]},
    {"id": "compare", "label": lang["menu_compare_label"], "sub_title": lang["menu_compare_sub"]},
    {"id": "rewrite", "label": lang["menu_rewrite_label"], "sub_title": lang["menu_rewrite_sub"]},
    {"id": "translate", "label": lang["menu_translate_label"], "sub_title": lang["menu_translate_sub"]},
    {"id": "pdf2word", "label": lang["menu_pdf2word_label"], "sub_title": lang["menu_pdf2word_sub"]}
]

INDUSTRY_TRACKS = [lang["track_general"], lang["track_ai"], lang["track_consulting"], lang["track_risk"], lang["track_manufacture"]]
TRACK_PROMPT_MAP = {
    lang["track_general"]: PROMPT_CONFIG["industry_report_general"],
    lang["track_ai"]: PROMPT_CONFIG["industry_report_ai"],
    lang["track_consulting"]: PROMPT_CONFIG["industry_report_consulting"],
    lang["track_risk"]: PROMPT_CONFIG["industry_report_risk"],
    lang["track_manufacture"]: PROMPT_CONFIG["industry_report_manufacture"]
}

MENU_LABELS = [item["label"] for item in MENU_CONFIG]
MENU_MAP = {item["label"]: item for item in MENU_CONFIG}

client = OpenAI(
    base_url=os.getenv("ARK_BASE_URL", LOCAL_CONFIG["base_url"]),
    api_key=os.getenv("ARK_API_KEY", LOCAL_CONFIG["api_key"]),
)

# ==============================================================================
# 【通用工具函数】100%原样不动
# ==============================================================================
def ai_request(prompt):
    try:
        response = client.responses.create(
            model=LOCAL_CONFIG["model_name"],
            input=[{"role": "user", "content": [{"type": "input_text", "text": prompt}]}]
        )
        full_text = ""
        if hasattr(response, "output") and response.output:
            for output in response.output:
                if hasattr(output, "content") and output.content:
                    for content_item in output.content:
                        if hasattr(content_item, "text") and content_item.text:
                            full_text += content_item.text
        return full_text.strip() if full_text else "AI生成内容为空，请重试~"
    except Exception as e:
        st.error(f"{lang['error']}: {str(e)}")
        return ""

def read_file(uploaded_file):
    try:
        raw = uploaded_file.read()
        if uploaded_file.name.lower().endswith(".docx"):
            doc = Document(io.BytesIO(raw))
            full_text = "\n".join([p.text for p in doc.paragraphs])
            return full_text.strip()
        elif uploaded_file.name.lower().endswith(".txt"):
            return raw.decode("utf-8", errors="ignore").strip()
        else:
            return "不支持的文件格式"
    except Exception as e:
        return f"文件读取失败：{str(e)}"

def generate_word_file(content):
    doc = Document()
    for para in content.split("\n"):
        if para.strip():
            doc.add_paragraph(para.strip())
    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

def generate_ppt_file(content, title, footer, end_text):
    prs = Presentation()
    title_slide_layout = prs.slide_layouts[0]
    slide = prs.slides.add_slide(title_slide_layout)
    slide.shapes.title.text = title
    slide.placeholders[1].text = footer

    paragraphs = content.split("\n")
    current_content = ""
    slide_layout = prs.slide_layouts[1]
    current_title = "正文内容"

    for para in paragraphs:
        para = para.strip()
        if not para:
            continue
        if para.startswith("#") or para.startswith("1.") or para.startswith("一、") or "核心" in para or "报告" in para:
            if current_content:
                slide = prs.slides.add_slide(slide_layout)
                slide.shapes.title.text = current_title
                tf = slide.placeholders[1].text_frame
                tf.text = current_content
                for p in tf.paragraphs:
                    p.font.size = Pt(12)
                    p.font.name = "微软雅黑"
                current_content = ""
            current_title = para.replace("#", "").strip()
        else:
            current_content += para + "\n"

    if current_content and current_title:
        slide = prs.slides.add_slide(slide_layout)
        slide.shapes.title.text = current_title
        tf = slide.placeholders[1].text_frame
        tf.text = current_content
        for p in tf.paragraphs:
            p.font.size = Pt(12)
            p.font.name = "微软雅黑"

    end_slide_layout = prs.slide_layouts[5]
    slide = prs.slides.add_slide(end_slide_layout)
    title_shape = slide.shapes.title
    title_shape.text = end_text
    title_shape.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
    title_shape.text_frame.paragraphs[0].font.size = Pt(32)

    buffer = io.BytesIO()
    prs.save(buffer)
    buffer.seek(0)
    return buffer

def extract_pdf_text(pdf_file):
    try:
        reader = PdfReader(pdf_file)
        text = ""
        for page in reader.pages:
            page_text = page.extract_text()
            if page_text:
                text += page_text + "\n\n"
        return text.strip() if text else "PDF文本提取为空"
    except Exception as e:
        return f"PDF读取失败：{str(e)}"

# ==============================================================================
# 【业务功能】全部原样不动
# ==============================================================================
def render_search():
    kw = st.text_input(lang["search_input_tip"])
    if st.button(lang["search_btn"], use_container_width=True):
        if kw:
            with st.spinner(lang["search_loading"]):
                prompt = PROMPT_CONFIG["report_search"].format(keyword=kw, compliance_rule=PROMPT_CONFIG["compliance_rule"])
                content = ai_request(prompt)
                if content:
                    lines = content.strip().split("\n")
                    for i, line in enumerate(lines):
                        if "|" in line:
                            p = line.split("|")
                            if len(p)>=4:
                                st.markdown(f"### {i+1}. {p[0]}")
                                st.write(f"{lang['search_pub_org']}{p[1]} | {lang['search_pub_year']}{p[2]}")
                                st.write(f"{lang['search_abstract']}{p[3]}")
                                st.divider()
        else:
            st.warning(f"{lang['warning']}: {lang['search_kw_empty']}")

def render_summary():
    summary_mode = st.radio(
        lang["summary_mode"],
        options=[lang["summary_mode_general"], lang["summary_mode_indicator"]],
        horizontal=True
    )
    st.markdown("---")
    f = st.file_uploader(lang["summary_upload_tip"], type=["txt","docx"])
    if f and st.button(lang["summary_analyze_btn"], use_container_width=True):
        with st.spinner(lang["summary_analyze_loading"].format(mode=summary_mode)):
            txt = read_file(f)
            st.text_area(lang["summary_original_preview"], txt, height=200)
            st.markdown("---")
            if summary_mode == lang["summary_mode_general"]:
                prompt = PROMPT_CONFIG["doc_summary_general"].format(text=txt[:3500], compliance_rule=PROMPT_CONFIG["compliance_rule"])
            else:
                prompt = PROMPT_CONFIG["doc_summary_indicator"].format(text=txt[:6000], compliance_rule=PROMPT_CONFIG["compliance_rule"])
            res = ai_request(prompt)
            st.markdown(f"### {lang['summary_result_title'].format(mode=summary_mode)}")
            st.write(res)
            word_buf = generate_word_file(res)
            st.download_button(
                label=lang["summary_download_btn"],
                data=word_buf,
                file_name=lang["summary_download_filename"].format(mode=summary_mode),
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True
            )

def render_generate():
    col1, col2 = st.columns(2)
    with col1:
        selected_track = st.selectbox(lang["generate_track_select"], options=INDUSTRY_TRACKS)
    with col2:
        industry_name = st.text_input(lang["generate_name_input"])
    st.markdown("---")
    st.markdown(lang["generate_ref_tip"])
    reference_file = st.file_uploader(lang["generate_ref_upload"], type=["txt","docx"], key="reference_file")
    reference_text = ""
    if reference_file:
        reference_text = read_file(reference_file)
        with st.expander(lang["generate_ref_preview"]):
            st.text_area("参考资料", reference_text, height=200)
    st.markdown("---")
    if st.button(lang["generate_btn"], use_container_width=True):
        if not industry_name:
            st.warning(f"{lang['warning']}: {lang['generate_name_empty']}")
        else:
            with st.spinner(lang["generate_loading"].format(track=selected_track)):
                base_prompt = TRACK_PROMPT_MAP[selected_track]
                full_prompt = base_prompt.format(name=industry_name, compliance_rule=PROMPT_CONFIG["compliance_rule"])
                if reference_text:
                    full_prompt += f"\n\n{lang['generate_ref_rule']}\n{reference_text[:3000]}"
                report_content = ai_request(full_prompt)
                st.markdown(f"### {lang['generate_report_title'].format(name=industry_name, track=selected_track)}")
                st.write(report_content)
                col_word, col_ppt = st.columns(2)
                with col_word:
                    word_buf = generate_word_file(report_content)
                    st.download_button(
                        label=lang["generate_download_word"],
                        data=word_buf,
                        file_name=lang["generate_word_filename"].format(name=industry_name, track=selected_track),
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        use_container_width=True
                    )
                with col_ppt:
                    ppt_buf = generate_ppt_file(
                        content=report_content,
                        title=f"{industry_name} {lang['ppt_title_default']}",
                        footer=lang["ppt_footer"],
                        end_text=lang["ppt_end_page"]
                    )
                    st.download_button(
                        label=lang["generate_download_ppt"],
                        data=ppt_buf,
                        file_name=lang["generate_ppt_filename"].format(name=industry_name, track=selected_track),
                        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                        use_container_width=True
                    )

def render_compare():
    st.markdown(lang["compare_tip"])
    upload_files = st.file_uploader(
        lang["compare_upload_tip"],
        type=["txt","docx"],
        accept_multiple_files=True
    )
    st.markdown("---")
    if st.button(lang["compare_btn"], use_container_width=True):
        if not upload_files or len(upload_files) < 2:
            st.warning(f"{lang['warning']}: {lang['compare_file_min']}")
        else:
            with st.spinner(lang["compare_loading"]):
                all_doc_text = ""
                for i, file in enumerate(upload_files):
                    file_text = read_file(file)
                    all_doc_text += f"===== 文档{i+1}：{file.name} =====\n{file_text[:3000]}\n\n"
                prompt = PROMPT_CONFIG["multi_doc_compare"].format(all_doc_text=all_doc_text, compliance_rule=PROMPT_CONFIG["compliance_rule"])
                compare_result = ai_request(prompt)
                st.session_state.compare_result = compare_result
    if st.session_state.compare_result:
        st.markdown(f"### {lang['compare_result_title']}")
        st.write(st.session_state.compare_result)
        col_word, col_ppt = st.columns(2)
        with col_word:
            word_buf = generate_word_file(st.session_state.compare_result)
            st.download_button(
                label=lang["compare_download_word"],
                data=word_buf,
                file_name=lang["compare_word_filename"],
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True
            )
        with col_ppt:
            ppt_buf = generate_ppt_file(
                content=st.session_state.compare_result,
                title=lang["compare_result_title"],
                footer=lang["ppt_footer"],
                end_text=lang["ppt_end_page"]
            )
            st.download_button(
                label=lang["compare_download_ppt"],
                data=ppt_buf,
                file_name=lang["compare_ppt_filename"],
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                use_container_width=True
            )

def render_rewrite():
    st.markdown(lang["rewrite_flow"])
    col1, col2 = st.columns(2)
    template_text = ""
    content_text = ""
    with col1:
        template_file = st.file_uploader(lang["rewrite_template_upload"], type=["txt","docx"], key="template_file")
        if template_file:
            template_text = read_file(template_file)
            with st.expander(lang["rewrite_template_preview"]):
                st.text_area("Template", template_text, height=280, key="template_preview")
    with col2:
        content_file = st.file_uploader(lang["rewrite_content_upload"], type=["txt","docx"], key="content_file")
        if content_file:
            content_text = read_file(content_file)
            with st.expander(lang["rewrite_content_preview"]):
                st.text_area("Original", content_text, height=280, key="content_preview")
    st.markdown("---")
    if st.button(lang["rewrite_btn"], use_container_width=True, disabled=st.session_state.rewrite_generating):
        if not template_file or not content_file:
            st.warning(f"{lang['warning']}: {lang['rewrite_file_empty']}")
        else:
            st.session_state.rewrite_result = ""
            st.session_state.rewrite_generating = True
    if st.session_state.rewrite_generating and not st.session_state.rewrite_result:
        with st.spinner(lang["rewrite_loading"]):
            prompt = PROMPT_CONFIG["template_rewrite"].format(template_content=template_text[:2500], original_content=content_text[:3500], compliance_rule=PROMPT_CONFIG["compliance_rule"])
            result_text = ai_request(prompt)
            st.session_state.rewrite_result = result_text
            st.session_state.rewrite_generating = False
    if st.session_state.rewrite_result and not st.session_state.rewrite_generating:
        st.markdown(f"### {lang['rewrite_result_title']}")
        st.text_area("Result", st.session_state.rewrite_result, height=450, key="rewrite_result_preview")
        col_word, col_ppt = st.columns(2)
        with col_word:
            word_buf = generate_word_file(st.session_state.rewrite_result)
            st.download_button(
                label=lang["rewrite_download_word"],
                data=word_buf,
                file_name=lang["rewrite_word_filename"],
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True
            )
        with col_ppt:
            ppt_buf = generate_ppt_file(
                content=st.session_state.rewrite_result,
                title=lang["rewrite_result_title"],
                footer=lang["ppt_footer"],
                end_text=lang["ppt_end_page"]
            )
            st.download_button(
                label=lang["rewrite_download_ppt"],
                data=ppt_buf,
                file_name=lang["rewrite_ppt_filename"],
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                use_container_width=True
            )

def render_translate():
    st.markdown(lang["translate_tip"])
    col1, col2 = st.columns(2)
    with col1:
        target_lang = st.selectbox(
            lang["translate_target_lang"],
            options=TARGET_LANG_OPTIONS[current_lang],
            index=1
        )
    with col2:
        translate_mode = st.radio(
            lang["translate_mode"],
            options=[lang["translate_mode_text"], lang["translate_mode_file"]],
            horizontal=True
        )
    st.markdown("---")
    source_text = ""
    if translate_mode == lang["translate_mode_text"]:
        source_text = st.text_area(lang["translate_textarea_tip"], height=200)
    else:
        translate_file = st.file_uploader(lang["translate_upload_tip"], type=["txt","docx"])
        if translate_file:
            source_text = read_file(translate_file)
            with st.expander(lang["translate_original_preview"]):
                st.text_area("Original", source_text, height=250)
    st.markdown("---")
    if st.button(lang["translate_btn"], use_container_width=True, disabled=st.session_state.translate_generating):
        if not source_text.strip():
            st.warning(f"{lang['warning']}: {lang['translate_content_empty']}")
        else:
            st.session_state.translate_result = ""
            st.session_state.translate_generating = True
    if st.session_state.translate_generating and not st.session_state.translate_result:
        with st.spinner(lang["translate_loading"]):
            prompt = PROMPT_CONFIG["doc_translate"].format(target_lang=target_lang, text=source_text[:6000])
            result_text = ai_request(prompt)
            st.session_state.translate_result = result_text
            st.session_state.translate_generating = False
    if st.session_state.translate_result and not st.session_state.translate_generating:
        st.markdown(f"### {lang['translate_result_title']}")
        st.text_area("Result", st.session_state.translate_result, height=400, key="translate_result_preview")
        word_buf = generate_word_file(st.session_state.translate_result)
        st.download_button(
            label=lang["translate_download_btn"],
            data=word_buf,
            file_name=lang["translate_download_filename"],
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True
        )

# ==============================================
# ✅【唯一修改】PDF转Word 纯本地pdf2docx，无AI、无大模型、无API
# ==============================================
def render_pdf2word():
    """功能7：无损PDF转Word 【纯本地离线转换，不调用任何AI】"""
    st.markdown(lang["pdf2word_tip"])
    pdf_file = st.file_uploader(lang["pdf2word_upload_tip"], type=["pdf"], key="pdf_file")

    if pdf_file:
        with st.spinner(lang["pdf2word_loading"]):
            # 读取PDF
            raw_bytes = pdf_file.read()
            temp_pdf = "temp_only_pdf.pdf"
            with open(temp_pdf, "wb") as f:
                f.write(raw_bytes)

            # 本地pdf2docx 直接转换
            temp_docx = "temp_only_word.docx"
            cv = Converter(temp_pdf)
            cv.convert(temp_docx)
            cv.close()

            # 读取转换后的Word
            with open(temp_docx, "rb") as f:
                word_data = f.read()

            # 清理临时文件
            try:
                os.remove(temp_pdf)
                os.remove(temp_docx)
            except:
                pass

            # 下载按钮
            st.success(lang["success"])
            st.download_button(
                label=lang["pdf2word_download_word"],
                data=word_data,
                file_name=lang["pdf2word_word_filename"],
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True
            )

# ==============================================================================
# 映射、页面CSS、布局、侧边栏 100%完全复原
# ==============================================================================
RENDER_FUNC_MAP = {
    "search": render_search,
    "summary": render_summary,
    "generate": render_generate,
    "compare": render_compare,
    "rewrite": render_rewrite,
    "translate": render_translate,
    "pdf2word": render_pdf2word
}

st.set_page_config(
    page_title=lang["page_title"],
    page_icon=PAGE_CONFIG["page_icon"],
    layout="wide",
    initial_sidebar_state="expanded"
)

st.markdown("""
<style>
#MainMenu {visibility: hidden !important;}
footer {visibility: hidden !important;}
button[title="View fullscreen"] {visibility: hidden !important;}
.stDeployButton {display: none !important;}
::-webkit-scrollbar {display: none !important;}
header {height: 0 !important;background: transparent !important;border: none !important;}
button[aria-label="Open sidebar"] {position: fixed !important;top: 1rem !important;left: 1rem !important;z-index: 9999 !important;background-color: rgba(255,255,255,0.9) !important;border-radius: 50% !important;width: 2.5rem !important;height: 2.5rem !important;box-shadow: 0 2px 8px rgba(0,0,0,0.1) !important;}
button[aria-label="Close sidebar"] {display: none !important;}
.block-container {padding-top: 0 !important;padding-bottom: 1rem !important;max-width: 95% !important;}
h1 {margin-top: 0.5rem !important;}
</style>
""", unsafe_allow_html=True)

st.title(lang["main_title"])
st.markdown("---")

with st.sidebar:
    st.radio(
        lang["lang_select"],
        options=["zh", "en"],
        format_func=lambda x: "中文" if x == "zh" else "English",
        key="language",
        horizontal=True
    )
    st.divider()
    st.header(lang["sidebar_title"])
    if st.session_state.selected_tab == "":
        st.session_state.selected_tab = MENU_LABELS[0]
    st.radio(
        lang["select_func"],
        MENU_LABELS,
        key="selected_tab",
        label_visibility="visible"
    )
    st.markdown("---")
    st.info(lang["sidebar_footer"])

current_tab = MENU_MAP[st.session_state.selected_tab]
st.subheader(current_tab["sub_title"])
st.markdown("---")

if current_tab["id"] in RENDER_FUNC_MAP:
    RENDER_FUNC_MAP[current_tab["id"]]()
else:
    st.warning(lang["func_not_found"])